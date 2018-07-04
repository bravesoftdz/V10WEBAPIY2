unit InitSoc;
     
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, DB, Hctrls, ComCtrls, Buttons, Ent1, HEnt1, hmsgbox,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  MajTable, HStatus, HSysMenu,LicUtil,hCRC, Hqry, ADODB ;

type
  TFInitSoc = class(TForm)
    PBouton: TPanel;
    TT: TLabel;
    HelpBtn: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    Pages: TPageControl;
    PInfos: TTabSheet;
    Precup: TTabSheet;
    DBSOURCE: TDatabase;
    DBDest: TDatabase;
    Bevel1: TBevel;
    GroupBox1: TGroupBox;
    EXERCICE: TCheckBox;
    SUIVCPTA: TCheckBox;
    CORRESP: TCheckBox;
    REFAUTO: TCheckBox;
    GUIDES: TCheckBox;
    RUBRIQUE: TCheckBox;
    RUPTURE: TCheckBox;
    GroupBox3: TGroupBox;
    PAYS: TCheckBox;
    REGION: TCheckBox;
    CODEPOST: TCheckBox;
    CODEAFB: TCheckBox;
    DEVISE: TCheckBox;
    MODEPAIE: TCheckBox;
    MODEREGL: TCheckBox;
    GroupBox4: TGroupBox;
    USERGRP: TCheckBox;
    UTILISAT: TCheckBox;
    MODELES: TCheckBox;
    ETABLISS: TCheckBox;
    Msg: THMsgBox;
    FRef: THValComboBox;
    GSOCIETE: TGroupBox;
    SO_SOCIETE: TEdit;
    HLabel2: THLabel;
    HLabel1: THLabel;
    Bevel2: TBevel;
    MODEDATA: TCheckBox;
    Entete: THTable;
    LLigne: THTable;
    RempliDefaut: THValComboBox;
    GUSERGRP: TGroupBox;
    HLabel7: THLabel;
    UG_GROUPE: TEdit;
    GUTILISAT: TGroupBox;
    HLabel4: THLabel;
    HLabel6: THLabel;
    US_UTILISATEUR: TEdit;
    US_PASSWORD: TEdit;
    TUS_ABREGE: THLabel;
    US_ABREGE: TEdit;
    US_GROUPE: TEdit;
    US_SUPERVISEUR: TEdit;
    UG_NIVEAUACCES: THNumEdit;
    UG_NUMERO: THNumEdit;
    HMTrad: THSystemMenu;
    HLabel3: THLabel;
    SO_LIBELLE: TEdit;
    UG_LIBELLE: TEdit;
    HLabel8: THLabel;
    HLabel5: THLabel;
    US_LIBELLE: TEdit;
    DBREF: TDatabase;
    BTag: THBitBtn;
    Bdetag: THBitBtn;
    FPassWord2: TEdit;
    TUS_PASSWORD2: THLabel;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MODELESClick(Sender: TObject);
    procedure UTILISATClick(Sender: TObject);
    procedure US_LIBELLEChange(Sender: TObject);
    procedure UG_GROUPEChange(Sender: TObject);
    procedure MODEPAIEClick(Sender: TObject);
    procedure MODEREGLClick(Sender: TObject);
    procedure USERGRPClick(Sender: TObject);
    procedure FRefChange(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FDestination : String ;
    RDriver,DDriver,SDriver,SocDriver : TDBDriver ;
    RODBC,DODBC,SODBC,SocODBC : Boolean ;
    procedure InitialiseSoc ;
    function  OkDonnees : boolean ;
    procedure ForceChecked(C1,C2 : TCheckBox) ;
    Procedure MajRupture ;
    procedure CreerTables ;
    Function  Nbrecords(T : TDataSet) : Integer ;
    Procedure InitialiseLesTables ;
    function  RecupChpDB (DB,Fichier,Champ,Where : String ; var Valeur : Variant ) : Boolean ;
    procedure SelDeSel(S : TCheckBoxState) ;
  public
    { Déclarations publiques }
  end;

procedure InitSociete(FDestination : String) ;

implementation

{$R *.DFM}

procedure InitSociete(FDestination : String) ;
var FInitSoc : TFInitSoc ;
BEGIN
FInitSoc:=TFInitSoc.Create(Application) ;
Try
 FInitSoc.FDestination:=FDestination ;
 FInitSoc.ShowModal ;
 Finally
 FInitSoc.Free ;
 end ;
END ;

procedure TFInitSoc.FormShow(Sender: TObject);
begin
SocDriver:=V_PGI.Driver ; SocODBC:=V_PGI.ODBC ; V_PGI.StopCourrier:=TRUE ;
PRecup.TabVisible:=False ;
ChargeDossier(FRef.Items,True) ; FRef.ItemIndex:=-1 ;
SO_SOCIETE.Text:=V_PGI.CodeSociete ;
SO_LIBELLE.Text:=V_PGI.NomSociete ;
UG_GROUPE.Text:=V_PGI.Groupe ;
UG_LIBELLE.Text:=RechDom('ttUserGroupe',V_PGI.Groupe,False) ;
US_UTILISATEUR.Text:=V_PGI.User ;
US_LIBELLE.Text:=RechDom('ttUtilisateur',V_PGI.User,False) ;
end;

procedure TFInitSoc.MODELESClick(Sender: TObject);
begin MODEDATA.Checked:=MODELES.Checked ; end;

procedure TFInitSoc.UTILISATClick(Sender: TObject);
begin ForceChecked(UTILISAT,USERGRP) ; end;

procedure TFInitSoc.USERGRPClick(Sender: TObject);
begin ForceChecked(UTILISAT,USERGRP) ; end;

procedure TFInitSoc.MODEPAIEClick(Sender: TObject);
begin ForceChecked(MODEREGL,MODEPAIE) ; end;

procedure TFInitSoc.MODEREGLClick(Sender: TObject);
begin ForceChecked(MODEREGL,MODEPAIE) ;end;

procedure TFInitSoc.ForceChecked(C1,C2 : TCheckBox) ;
BEGIN if C1.Checked then C2.Checked:=True ; END ;

procedure TFInitSoc.SelDeSel(S : TCheckBoxState) ;
var G,C : Tcontrol ;
    i,j : integer ;
begin
for i:=0 to PRecup.Controlcount-1 do
  BEGIN
  G:=PRecup.Controls[i] ;
  if G is TGroupBox then
    for j:=0 to TGroupBox(G).ControlCount-1 do
      BEGIN
      C:=TGroupBox(G).Controls[j] ;
      if C is TCheckBox then TCheckBox(C).State:=S ;
      END ;
  END ;
MODEPAIE.State:=S ; USERGRP.State:=S ;
end;

function TFInitSoc.OkDonnees : boolean ;
var G,C : TControl ;
    i,j : integer ;
BEGIN
Result:=False ;
for i:=0 to PInfos.ControlCount-1 do
  BEGIN
  G:=PInfos.Controls[i] ;
  if G is TGroupBox then
    for j:=0 to TGroupBox(G).ControlCount-1 do
      BEGIN
      C:=TGroupBox(G).Controls[j] ;
        if C is TEdit then
            if (TEdit(C).Text='') then BEGIN Pages.ActivePage:=PInfos ; TEdit(C).SetFocus ; Exit ; END ;
      END ;
  END ;
Result:=True ;
END ;

procedure TFInitSoc.BValiderClick(Sender: TObject);
begin
ModalResult:=mrNone ;
if (FRef.Items[FRef.ItemIndex]=FDestination) then
  BEGIN
  Pages.ActivePage:=PInfos ;
  FRef.SetFocus ;
  Msg.Execute(7,Caption,'') ;
  Exit ;
  END ;
if Not OkDonnees then BEGIN Msg.Execute(2,Caption,'') ; Exit ; END ;
if (FPassWord2.Text<>US_PASSWORD.Text) then BEGIN FPassword2.SetFocus ; Msg.Execute(9,Caption,'') ; Exit ; END ;
case Msg.Execute(0,Caption,'') of
  mrNo     : ;
  mrCancel : Exit ;
  END ;
InitialiseSoc ;
Msg.Execute(3,Caption,Chr(10)+Msg.Mess[8]) ;
ModalResult:=mrOk ;
end;

procedure TFInitSoc.InitialiseSoc ;
BEGIN
ConnecteDB(GetSocRef,DBREF,'DBREF') ; RDriver:=V_PGI.Driver ; RODBC:=V_PGI.ODBC ;
Entete.DatabaseName:='DBREF' ; LLigne.DatabaseName:='DBREF' ;
DBREF.Connected := True ; Entete.Open ; LLigne.Open ;
ConnecteDB(FRef.Items[FRef.ItemIndex],DBSOURCE,'DBSOURCE') ; SDriver:=V_PGI.Driver ; SODBC:=V_PGI.ODBC ;
DBSOURCE.Connected:=True ;
ConnecteDB(FDestination,DBDest,'DBDest') ; DDriver:=V_PGI.Driver ; DODBC:=V_PGI.ODBC ;
DBDest.Connected := True ;

V_PGI.Driver:=SocDriver ; V_PGI.ODBC:=SocODBC ;

CreerTables ;

END ;

function TFInitSoc.RecupChpDB (DB,Fichier,Champ,Where : String ; var Valeur : Variant ) : Boolean ;
Var Q : TQuery ;
    SQL : String ;
BEGIN
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DB ;
SQL:='Select '+Champ+ ' from '+Fichier + Where ;
Q.SQL.Clear ; Q.SQL.Add(SQL) ;
V_PGI.Driver:=RDriver ; V_PGI.ODBC:=RODBC ;
ChangeSQL(Q) ;
V_PGI.Driver:=SocDriver ; V_PGI.ODBC:=SocODBC ;
Q.RequestLive:=False ; Q.Open ;
Result:=(not Q.EOF) ;
if not Q.Eof then Valeur:=Q.Fields[0].Value ;
Q.Close ; Q.Free ;
END ;

Procedure TFInitSoc.MajRupture ;
Var QD,QS : THQuery ;
BEGIN
QS:=THQuery.Create(Application) ;
QS.DataBaseName:=DBSOURCE.DataBaseName ;
QS.SessionName:=DBSOURCE.SessionName ;
QD:=THQuery.Create(Application) ;
QD.DataBaseName:=DBDest.DataBaseName ;
QD.SessionName:=DBDest.SessionName ;
QD.RequestLive:=TRUE ;

QS.SQL.Clear ;
QS.SQL.Add('Select * from Choixcod where CC_TYPE="RUG"') ;
V_PGI.Driver:=SDriver ; V_PGI.ODBC:=SODBC ;
ChangeSQL(QS) ; QS.Open ;
V_PGI.Driver:=SocDriver ; V_PGI.ODBC:=SocODBC ;

QD.SQL.Clear ;
QD.SQL.Add('Select * From CHOIXCOD where CC_TYPE="'+W_W+'"') ;
V_PGI.Driver:=DDriver ; V_PGI.ODBC:=DODBC ;
ChangeSQL(QD) ; QD.Open ;
V_PGI.Driver:=SocDriver ; V_PGI.ODBC:=SocODBC ;

While Not QS.Eof Do
  BEGIN
  QD.Insert ;
  InitNew(QD) ;
  QD.FindField('CC_TYPE').AsString:=QS.FindField('CC_TYPE').AsString ;
  QD.FindField('CC_CODE').AsString:=QS.FindField('CC_CODE').AsString ;
  QD.FindField('CC_LIBELLE').AsString:=QS.FindField('CC_LIBELLE').AsString ;
  QD.FindField('CC_ABREGE').AsString:=QS.FindField('CC_ABREGE').AsString ;
  QD.FindField('CC_LIBRE').AsString:=QS.FindField('CC_LIBRE').AsString ;
  QD.Post ;
  QS.Next ;
  END ;
QS.Close ; QD.Close ;
QS.SQL.Clear ;
QS.SQL.Add('DELETE From RUPTURE where RU_NATURERUPT="RUG"') ;
V_PGI.Driver:=SDriver ; V_PGI.ODBC:=SODBC ;
ChangeSQL(QS) ; QS.ExecSQL ;
V_PGI.Driver:=SocDriver ; V_PGI.ODBC:=SocODBC ;
QS.Free ; QD.Free ;
END ;

procedure TFInitSoc.CreerTables ;
var StLibTable,StTable,Prefixe : String ;
    T      : THTable ;
    G,C    : TControl ;
    j,i,Nb(*,VersionTable*) : integer ;
    Prem   : Boolean ;
    OkMajRupture : Boolean ;
BEGIN
Nb:=NbRecords(Entete)+PInfos.ControlCount ;
Initmove(Nb,'') ;
Entete.First ; OkMajRupture:=FALSE ;
While not Entete.EOF do
  BEGIN
  StLibTable:=Entete.FindField('DT_LIBELLE').AsString ;
  StTable:=Entete.FindField('DT_NOMTABLE').AsString ;
  //VersionTable:=Entete.FindField('DT_NUMVERSION').AsInteger ;
  TT.Caption:=Msg.Mess[4]+' '+StLibTable ;
  Application.ProcessMessages ;
  Prefixe:=Entete.FindField('DT_PREFIXE').AsString ;
  LLigne.SetRange([Prefixe,000],[Prefixe,9999]) ;
  DBCreateTable(DBDest,Entete,LLigne,DDriver,False) ;
  if DDriver=dbMSACCESS then ConnecteTRUEFALSE(DBDest) ;
  if (DDriver<>dbMSSQL) and (DDriver<>dbMSSQL2005) then DBDest.StartTransaction ;  //10/08/2006 YMO Ajout test SQL2005
  if ((FRef.ItemIndex>-1 ) and ((FindComponent(StTable)<>nil) and (TCheckBox(FindComponent(StTable)).Checked)) or (RempliDefaut.Items.IndexOf(StTable)>-1)) then
    BEGIN
    TT.Caption:=Msg.Mess[5]+' '+StLibTable ;
    Application.ProcessMessages ;
    if (RempliDefaut.Items.IndexOf(StTable)>-1) then DBCopyTableBM(DBREF,DBDest,StTable,StTable)
                                                else DBCopyTableBM(DBSOURCE,DBDest,StTable,StTable) ;
    If StTable='RUPTURE' Then OkMajRupture:=TRUE ;
    if StTable='AXE' then
      BEGIN
      T:=THTable.Create(Application) ;
      T.DataBaseName:=DBDest.DataBaseName ;
      T.TableName:=StTable ; T.IndexName:='X_CLE1' ;
      T.Open ;
      While not T.Eof do
        BEGIN
        T.Edit ;
        T.FindField('X_SECTIONATTENTE').AsString:='' ;
        T.FindField('X_GENEATTENTE').AsString:='' ;
        T.FindField('X_BOURREANA').asString:='0' ;
        T.Post ;
        T.Next ;
        END ;
      T.Close ;
      T.Free ;
      END ;
    END ;
  if (DDriver<>dbMSSQL) and (DDriver<>dbMSSQL2005) then DBDest.Commit ; //10/08/2006 YMO Ajout test SQL2005
  Entete.Next ;
  MoveCur(False) ;
  END ;
T:=THTable.Create(Application) ;
T.DataBaseName:=DBDest.DataBaseName ;
for j:=0 to PInfos.ControlCount-1 do
   BEGIN
   G:=PInfos.Controls[j] ;
   if G is TGroupBox then
      BEGIN
      T.TableName:=Copy(G.Name,2,length(G.Name)-1) ; T.IndexName:=G.Hint ;
      T.Open ;
      TT.Caption:=Msg.Mess[6]+' '+T.TableName ;
      Application.ProcessMessages ;
      Prem:=True ;
      if Pos('USERGRP',T.TableName)>0 then
         BEGIN
         Prem:=not T.FindKey([UG_GROUPE.Text]) ;
         END else if Pos('UTILISAT',T.TableName)>0 then
         BEGIN
         Prem:=not T.FindKey([US_UTILISATEUR.Text]) ;
         END ;
      if Prem then BEGIN T.Insert ; InitNew(T) ; END else T.Edit ;
      for i:=0 to TGroupBox(G).ControlCount-1 do
         BEGIN
         C:=TGroupBox(G).Controls[i] ;
         if (T.FindField(C.Name)<>Nil) then
           if (C.Name='US_PASSWORD') then T.FindField(C.Name).Value:=CryptageSt(TEdit(C).Text)
                                     else T.FindField(C.Name).Value:=TEdit(C).Text ;
         END ;
      if Pos('UTILISAT',T.TableName)>0 then T.FindField('US_CRC').AsInteger:=GetCRC32ForData(T) ;
      T.Post ; T.Close ;
      END ;
   MoveCur(False) ;
   END ;
T.Free ;
If OkMajRupture Then MajRupture ;
InitialiseLesTables ;
FiniMove ;
END ;

procedure TFInitSoc.InitialiseLesTables ;
var Val : Variant ;
    T : THTable ;
    i : integer ;
    CodeD : String ;
BEGIN
Val:='' ;
RecupChpDB('DBREF','SOCIETE','SO_VERSIONBASE','',Val) ;
T:=THTable.Create(Application) ;
T.DataBaseName:=DBDest.DataBaseName ;
// Societe
T.TableName:='SOCIETE' ; T.IndexName:='SO_CLE1' ;
T.Open ;
if T.Eof then BEGIN T.Close ; Exit ; END ;
T.Edit ;
T.FindField('SO_VERSIONBASE').Value:=Val ;
T.FindField('SO_SOCIETE').AsString:=SO_SOCIETE.Text ;
T.FindField('SO_LIBELLE').AsString:=SO_LIBELLE.Text ;
T.Post ;
T.Close ;
// SuivDec
T.TableName:='CHOIXCOD' ; T.IndexName:='CC_CLE1' ;
T.Open ;
for i:=1 to 5 do
    BEGIN
    if T.FindKey(['CID','CD'+IntToStr(i)]) then T.Edit else T.Insert ;
    InitNew(T) ;
    T.FindField('CC_TYPE').AsString:='CID' ;
    T.FindField('CC_CODE').AsString:='CD'+IntToStr(i) ;
    T.FindField('CC_LIBELLE').AsString:='' ;
    T.FindField('CC_ABREGE').AsString:='' ;
    T.FindField('CC_LIBRE').AsString:='' ;
    T.Post ;
    END ;
T.Close ;
// Si pas recopie des devises, créer au moins la devise principale
if ((FRef.ItemIndex<=-1) or (Not Devise.Checked)) then
   BEGIN
   T.TableName:='PARAMSOC' ; T.IndexName:='SOC_CLE1' ; T.Open ;
   if T.FindKey(['SO_DEVISEPRINC']) then CodeD:=T.FindField('SOC_DATA').AsString else CodeD:='FRF' ;
   T.Close ;
   T.TableName:='DEVISE' ; T.IndexName:='D_CLE1' ; T.Open ;
   T.Insert ; InitNew(T) ;
   T.FindField('D_DEVISE').AsString:=CodeD  ; T.FindField('D_LIBELLE').AsString:='Francs' ;
   T.FindField('D_DECIMALE').AsInteger:=2   ; T.FindField('D_QUOTITE').AsFloat:=1 ;
   T.FindField('D_MONNAIEIN').AsString:='X' ; T.FindField('D_FONGIBLE').AsString:='X' ;
   T.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
   T.Post ;
   T.Close ;
   T.TableName:='PARAMSOC' ; T.IndexName:='SOC_CLE1' ; T.Open ;
   if T.FindKey(['SO_TAUXEURO']) then BEGIN T.Edit ; T.FindField('SOC_DATA').AsVariant:=6.55957 ; T.Post ; END ;
   if T.FindKey(['SO_DECVALEUR']) then BEGIN T.Edit ; T.FindField('SOC_DATA').AsVariant:=2 ; T.Post ; END ;
   if T.FindKey(['SO_DEVISEPRINC']) then BEGIN T.Edit ; T.FindField('SOC_DATA').AsVariant:=CodeD ; T.Post ; END ;
   if T.FindKey(['SO_TENUEEURO']) then BEGIN T.Edit ; T.FindField('SOC_DATA').AsVariant:='-' ; T.Post ; END ;
   if T.FindKey(['SO_DATEDEBUTEURO']) then BEGIN T.Edit ; T.FindField('SOC_DATA').AsVariant:=36161 ; T.Post ; END ;
   if T.FindKey(['SO_SOCIETE']) then BEGIN T.Edit ; T.FindField('SOC_DATA').AsVariant:=SO_SOCIETE.Text ; T.Post ; END ;
   if T.FindKey(['SO_LIBELLE']) then BEGIN T.Edit ; T.FindField('SOC_DATA').AsVariant:=SO_LIBELLE.Text ; T.Post ; END ;
   T.Close ;
   END ;
T.Free ;
END ;

Function TFInitSoc.Nbrecords(T : TDataSet) : Integer ;
var Nb : Integer ;
BEGIN
T.First ; Nb:=0 ;
While not T.Eof do BEGIN Inc(Nb) ; T.Next ; END ;
Result:=Nb ;
END ;

procedure TFInitSoc.US_LIBELLEChange(Sender: TObject);
begin US_ABREGE.Text:=Copy(US_LIBELLE.Text,1,17) ; end;

procedure TFInitSoc.UG_GROUPEChange(Sender: TObject);
begin US_GROUPE.Text:=UG_GROUPE.Text ; end;

procedure TFInitSoc.FRefChange(Sender: TObject);
begin PRecup.TabVisible:=(FRef.ItemIndex>-1) ; end;

procedure TFInitSoc.BTagClick(Sender: TObject);
begin
SelDeSel(cbChecked) ;
BTag.Visible:=False ;BDeTag.Visible:=True ;
end;

procedure TFInitSoc.BdetagClick(Sender: TObject);
begin
SelDeSel(cbUnChecked) ;
BDeTag.Visible:=False ;BTag.Visible:=True ;
end;

procedure TFInitSoc.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFInitSoc.FormClose(Sender: TObject; var Action: TCloseAction);
begin
V_PGI.Driver:=SocDriver ; V_PGI.ODBC:=SocODBC ; V_PGI.StopCourrier:=FALSE ;
end;

end.

