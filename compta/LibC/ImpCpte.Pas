unit ImpCpte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, Hctrls, StdCtrls, Buttons, ExtCtrls, DB, DBCtrls,  HEnt1,
  Ent1, HStatus, FmtChoix{, Filtre SG6 10/11/04 FQ 14826}, ComCtrls, Menus, HSysMenu,MajTable, HTB97,
  Hqry, HPanel, UiUtil, Impfic, RappType, TImpFic, ImpFicU,UObjFiltres,udbxDataset,
  ADODB {SG6 10/11/04 FQ 14826} ;

type PlanCpt = (cptGEN,cptAUX,cptANA,cptBUDBJ,cptBUDBG,cptBUDBS) ;

type
  TFImpCpte = class(TForm)
    MsgBox: THMsgBox;
    SSource: TDataSource;
    Sauve: TSaveDialog;
    HErr: THMsgBox;
    HMTrad: THSystemMenu;
    Msg: THMsgBox;
    HMTitre: THMsgBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    T: THTable;
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    PTools: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BFormat: TToolbarButton97;
    BStop: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Pages: TPageControl;
    TabSheet1: TTabSheet;
    HLabel7: THLabel;
    HLabel2: THLabel;
    FFormat: THValComboBox;
    GModif: TGroupBox;
    iLib: TImage;
    iCol: TImage;
    iTel: TImage;
    iCar: TImage;
    iAdr: TImage;
    iRIB: TImage;
    Lib: TCheckBox;
    Car: TCheckBox;
    Coll: TCheckBox;
    Adr: TCheckBox;
    Rib: TCheckBox;
    Tel: TCheckBox;
    LgComptes: TCheckBox;
    Bevel1: TBevel;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BFormatClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure FFormatChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RechFileClick(Sender: TObject);
  private
    ObjFiltre : TObjFiltre; //Gestion des Filtres SG6 10/11/04  FQ 14826
    Arreter,OkLecture,OkNew : boolean ;
    NbErreur, NbImport,NbModifies : integer ;
    TErreur : TList ;
    CurNat : String ;
    NumTypCpt : PlanCpt ;
    Fichier : TextFile ;
    FiltreImp:String ;
    OnREQ : Boolean ;
    procedure ImporteFormatParametre ;
    procedure ImporteCptes ;
    Function  ImporteCpteSAARI (St : String ; CptColl : PtTTabCptLu) : boolean ;
    procedure ImporteTiers (St : String) ;
    procedure ImporteHalley (St : String ; FFic : TFichierBase) ;
    function  ImporteEDI(var St : String) : boolean ;
    procedure SousDetail(var St : String) ;
    function  Format_Date(St : String) : String ;
    function  Controles(St : String ; Var FFic : TFichierBase) : byte ;
    function  MajCompteRendu(NumErr : Integer ; Bloquant : Boolean ;  Code,Valeur : String ; j : Integer) : boolean ;
    function  TestBreak : boolean ;
    Procedure ModifSurExistant ;
    procedure AjouteErrCombo(NomChp,Where,Val : String ; Res : byte ; Code : Variant) ;
    procedure AfficheCompteRendu ;
    function  ModifLeCompte(T : TDataSet) : boolean ;
    Function MajSQLDefaut : TQuery ;
    procedure QuelleTable ;
    Function  ImporteCpteCEGID(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : Boolean ;
  public
    Lequel : String ;
    {JP 18/07/05 : en attendant que l'agl gère le dbo dans TableToPrefix : cf FQ AGL 11668}
    function RecupPrefixSansDBO(NomTable, NomChamp : string) : string;
  end;

procedure ImportCpte(Lequel : String) ;
procedure LanceImportCptExt(Lequel : String) ;

implementation

uses RapSuppr,ImpUtil,VerCpte,ImporFmt ;

{$R *.DFM}


Procedure ImportCpte(Lequel : String) ;
var FImpCpte:TFImpCpte ;
    PP : THPanel ;
    Composants : TControlFiltre; //SG6   Gestion des Filtes 10/11/04   FQ 14826
BEGIN
if Blocage(['nrCloture','nrBatch'],False,'') then Exit ;
FImpCpte:=TFImpCpte.Create(Application) ;

FImpCpte.Lequel:=Lequel ;

//SG6 10/11/04 Gestion des Filtres FQ 14826
Composants.PopupF   := FImpCpte.POPF;
Composants.Filtres  := FImpCpte.FFILTRES;
Composants.Filtre   := FImpCpte.BFILTRE;
Composants.PageCtrl := FImpCpte.Pages;
FImpCpte.ObjFiltre := TObjFiltre.Create(Composants,'');


PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FImpCpte.ShowModal ;
     finally
     FImpCpte.Free ;
//     Bloqueur('nrBatch',False) ;
     END ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FImpCpte,PP) ;
   FImpCpte.Show ;
   END ;
END ;

Procedure LanceImportCptExt(Lequel : String) ;
var FImpCpte:TFImpCpte ;
BEGIN
if Blocage(['nrCloture','nrBatch'],False,'') then Exit ;
FImpCpte:=TFImpCpte.Create(Application) ;
FImpCpte.Lequel:=Lequel ;
try
  FImpCpte.ShowModal ;
  finally
  FImpCpte.Free ;
//  Bloqueur('nrBatch',False) ;
  END ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFImpCpte.FormShow(Sender: TObject);
begin
OnReq:=FALSE ;
if Lequel='G' then //Généraux
   BEGIN
   FiltreImp:='IMPFGE' ;
   NumTypCpt :=CptGEN ;
   CurNat:='FGE' ;
   Caption:=HMtitre.Mess[0] ;
   HelpContext:=6210000 ;
   END else
if Lequel='A' then //Analytiques
   BEGIN
   FiltreImp:='IMPFSE';
   NumTypCpt :=CptANA ;
   CurNat:='FSE' ;
   Caption:= HMtitre.Mess[1] ;
   HelpContext:=6230000 ;
   END else
if (Lequel='T') or (Lequel='X') then //Auxiliaires
   BEGIN
   FiltreImp:='IMPFTI' ;
   NumTypCpt :=CptAUX ;
   CurNat:='FTI' ;
   Caption:= HMtitre.Mess[2] ;
   HelpContext:=6220000 ;
   END else
if (Lequel='FBJ') then //Journaux Budgétaires
   BEGIN
   FiltreImp:='IMPFBJ' ;
   NumTypCpt :=CptBUDBJ ;
   CurNat:='FBJ' ;
   Caption:=HMtitre.Mess[3] ;
   HelpContext:=6240000 ;
   END else
if (Lequel='FBG') then //Comptes Budgétaires
   BEGIN
   FiltreImp:='IMPFBG' ;
   NumTypCpt :=CptBUDBG ;
   CurNat:='FBG' ;
   Caption:=HMtitre.Mess[4] ;
   HelpContext:=6250000 ;
   END ;
if (Lequel='FBS') then //Secitons Budgétaires
   BEGIN
   FiltreImp:='IMPFBS' ;
   NumTypCpt :=CptBUDBS ;
   CurNat:='FBS' ;
   Caption:=HMtitre.Mess[5] ;
   HelpContext:=6260000 ;
   END ;
Sauve.Title:=Caption ;
ChangeDataTypeFmt('X',CurNat,FFormat) ;
GereFMTChoix('X',CurNat,FFormat,FileName,RechFile) ;
TErreur:=TList.Create;


//SG6 10/11/04 Gestion des Filtres  FQ 14826
ObjFiltre.FFI_TABLE:=FiltreImp;
ObjFiltre.Charger;



if (FFormat.ItemIndex=-1) then FFormatChange(nil) ;
QuelleTable ;
UpdateCaption(Self) ;


end;


procedure TFImpCpte.BFormatClick(Sender: TObject);
begin
ChoixFormatImpExp('X' ,CurNat) ;
If not (ChoixFmt.OkSauve) then Exit ;
With ChoixFmt do
  BEGIN
  if Format<>'' then FFormat.Value:=Format ;
  FFormat.Enabled:=not (FixeFmt) ;
  if Fichier<>'' then FileName.Text:=Fichier ;
  FileName.Enabled:=not (FixeFichier) ;
  RechFile.Enabled:=not (FixeFichier) ;
  END ;
end;

procedure TFImpCpte.BStopClick(Sender: TObject);
begin Arreter:=True ; end;

procedure TFImpCpte.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FreeAndNil(ObjFiltre);  //SG6 10/11/04 Gestion des Filtres FQ 14826
TErreur.Free ;

if IsInside(Self) then
   BEGIN
//   Bloqueur('nrBatch',False) ;
   Action:=caFree ;
   END ;
end;

function TFImpCpte.TestBreak : Boolean ;
BEGIN
Application.ProcessMessages ;
if Arreter then if Msgbox.Execute(18,Caption,'')<>mryes then Arreter:=False ;
Result:=Arreter ;
END ;

function TFimpCpte.Format_Date(St : String) : String ;
var Dat : String ;
    Y,M,J : Word ;
BEGIN
if St='' then St:='19000101' ;
Y:=StrToInt(Copy(St,1,4)) ;
M:=StrToInt(Copy(St,5,2)) ;
J:=StrToInt(Copy(Dat,7,2)) ;
Result:=DateToStr(EncodeDate(Y,M,J)) ;
END ;

(*
Procedure TFImpCpte.ModifSurExistant ;
var H : integer ;
BEGIN
LgComptes.Enabled:=FALSE ;
Case Lequel[1] of
  'G','A':BEGIN
          if Coll.Visible then
            BEGIN
            Coll.Visible:=False ; iCol.Visible:=False ;
            Adr.Visible:=False  ; iAdr.Visible:=False ;
            Tel.Visible:=False  ; iTel.Visible:=False ;
            RIB.Visible:=False  ; iRIB.Visible:=False ;
            H:=(Tel.Top-Coll.Top)+(Coll.Height+10) ;
            GModif.Height:=GModif.Height-H ;
            ClientHeight:=ClientHeight-H ;
            END ;
          END ;
  'X'    :BEGIN
          RIB.Visible:=(FFormat.Value='SAB') ; iRIB.Visible:=RIB.Visible ;
          if not Coll.Visible then
            BEGIN
            Coll.Visible:=True ; iCol.Visible:=True ;
            Adr.Visible:=True  ; iAdr.Visible:=True  ;
            Tel.Visible:=True  ; iTel.Visible:=True  ;
            H:=(Tel.Top-Coll.Top)+(Coll.Height+10) ;
            GModif.Height:=GModif.Height+H ;
            ClientHeight:=ClientHeight+H ;
            END ;
          LgComptes.Enabled:=TRUE ;
          END ;
  'F'    :BEGIN
          if Coll.Visible then
            BEGIN
            Coll.Visible:=False ; iCol.Visible:=False ;
            Adr.Visible:=False  ; iAdr.Visible:=False  ;
            Tel.Visible:=False  ; iTel.Visible:=False  ;
            RIB.Visible:=False  ; iRIB.Visible:=False ;
            H:=(Tel.Top-Coll.Top)+(Coll.Height+10) ;
            GModif.Height:=GModif.Height-H ;
            ClientHeight:=ClientHeight-H ;
            END ;
          END ;
  END ;
If Not LgComptes.Enabled Then LgComptes.Checked:=FALSE ;
END ;
*)

Procedure TFImpCpte.ModifSurExistant ;
var H : integer ;
BEGIN
LgComptes.Enabled:=FALSE ;
Case Lequel[1] of
  'G','A':BEGIN
          if Coll.Enabled then
            BEGIN
            Coll.Enabled:=False ; iCol.Visible:=False ;
            Adr.Enabled:=False  ; iAdr.Visible:=False ;
            Tel.Enabled:=False  ; iTel.Visible:=False ;
            RIB.Enabled:=False  ; iRIB.Visible:=False ;
            (*
            H:=(Tel.Top-Coll.Top)+(Coll.Height+10) ;
            GModif.Height:=GModif.Height-H ;
            ClientHeight:=ClientHeight-H ;
            *)
            END ;
          END ;
  'X'    :BEGIN
          RIB.Enabled:=(FFormat.Value='SAB') ; iRIB.Visible:=RIB.Enabled ;
          if not Coll.Enabled then
            BEGIN
            Coll.Enabled:=True ; iCol.Visible:=True ;
            Adr.Enabled:=True  ; iAdr.Visible:=True  ;
            Tel.Enabled:=True  ; iTel.Visible:=True  ;
            (*
            H:=(Tel.Top-Coll.Top)+(Coll.Height+10) ;
            GModif.Height:=GModif.Height+H ;
            ClientHeight:=ClientHeight+H ;
            *)
            END ;
          LgComptes.Enabled:=TRUE ;
          END ;
  'F'    :BEGIN
          if Coll.Enabled then
            BEGIN
            Coll.Enabled:=False ; iCol.Visible:=False ;
            Adr.Enabled:=False  ; iAdr.Visible:=False  ;
            Tel.Enabled:=False  ; iTel.Visible:=False  ;
            RIB.Enabled:=False  ; iRIB.Visible:=False ;
            (*
            H:=(Tel.Top-Coll.Top)+(Coll.Height+10) ;
            GModif.Height:=GModif.Height-H ;
            ClientHeight:=ClientHeight-H ;
            *)
            END ;
          END ;
  END ;
If FFormat.Value='CGN' Then
  BEGIN
  Lib.Enabled:=False  ; iLib.Visible:=False ;
  Car.Enabled:=False  ; iCar.Visible:=False ;
  Coll.Enabled:=False ; iCol.Visible:=False ;
  Adr.Enabled:=False  ; iAdr.Visible:=False  ;
  Tel.Enabled:=False  ; iTel.Visible:=False  ;
  RIB.Enabled:=False  ; iRIB.Visible:=False ;
  LgComptes.Enabled:=FALSE ;
  END Else
  BEGIN
  Lib.Enabled:=TRUE  ; iLib.Visible:=TRUE ;
  Car.Enabled:=TRUE  ; iCar.Visible:=TRUE ;
  END ;
If Not LgComptes.Enabled Then LgComptes.Checked:=FALSE ;
END ;

procedure TFImpCpte.AfficheCompteRendu ;
var Mss,Mss1:integer ;
    MssIMp,MssMod : String ;
    Reste,Index,Titre : String ;
    QuelCpte : Byte ;
BEGIN
if (NbModifies=0) and (NbImport=0) then Exit ;
Mss:=9 ; Mss1:=1 ; QuelCpte:=1 ;
// initialisé
//if (Lequel='G') then                 BEGIN Mss:=9  ; Mss1:=1 ; QuelCpte:=1 ; END else
if (CurNat='FSE') then                 BEGIN Mss:=10 ; Mss1:=3 ; QuelCpte:=3 ; END else
if (CurNat='FTI') then BEGIN Mss:=11 ; Mss1:=5 ; QuelCpte:=2 ; END else
if (CurNat='FBJ') then BEGIN Mss:=12 ; Mss1:=7 ; QuelCpte:=2 ; END else
if (CurNat='FBG') then BEGIN Mss:=13 ; Mss1:=9 ; QuelCpte:=2 ; END else
if (CurNat='FBS') then BEGIN Mss:=14 ; Mss1:=11 ; QuelCpte:=2 ; END ;

if (NbImport>1) then Inc(Mss1) ;
MssImp:=IntToStr(NbImport)+' '+Msg.Mess[Mss1] ;
if (NbImport>1) then Inc(Mss1,19) else Inc(Mss1,20) ;
if (NbModifies>1) then Inc(Mss1) ;
MssMod:=Chr(13)+Chr(13) ;
MssMod:=MssMod+IntToStr(NbModifies)+' '+Msg.Mess[Mss1] ;
if ChoixFmt.CompteRendu then
  BEGIN
  Reste:=Msg.Mess[0] ; Index:=ReadTokenSt(Reste) ; Titre:=ReadTokenSt(Reste) ;
  Msg.Mess[0]:=Index+';'+caption+';'+Reste;
  Msg.Execute(0,MssImp,MssMod) ;
  if (NbImport>0) then RapportDeSuppression(TErreur,2) ;
  END else MsgBox.Execute(Mss,caption,'') ;
if (CurNat='FGE') or (CurNat='FTI') or (CurNat='FSE') then
 if MsgBox.Execute(20,Caption,'')=mrYes then
  if VerCompte(QuelCpte,False)
    then MsgBox.Execute(21,Caption,' '+MsgBox.Mess[QuelCpte+21]+' '+MsgBox.Mess[28])
    else MsgBox.Execute(29,Caption,'') ;
END ;

function TFImpCpte.ModifLeCompte(T : TDataSet) : boolean ;
var Axe : String ;
    Val,Compte,Lib : String ;
    FFic : TFichierBase ;
    OkExiste : Boolean ;
    Chp, ////
    Adr1,Adr2,Adr3,CodePost,Ville,Tele : String ;
BEGIN
Result:=False ; FFic:=fbGene ;
if CurNat='FSE' then Axe:=T.Findfield('S_AXE').AsString else
 if CurNat='FBS' then Axe:=T.Findfield('BS_AXE').AsString ;

if CurNat='FBG' then BEGIN Compte:='BG_BUDGENE' ; END
 else if CurNat='FBS' then BEGIN Compte:='BS_BUDSECT' ; END
  else if CurNat='FBJ' then BEGIN Compte:='BJ_BUDJAL' ; END
  else if CurNat='FGE' then BEGIN Compte:='G_GENERAL' ; END
   else if CurNat='FTI' then BEGIN Compte:='T_AUXILIAIRE' ; END
    else if CurNat='FSE' then
      BEGIN
      Compte:='S_SECTION' ;
      FFic:=fbGene ;
      if (Axe='A1') then FFic:=fbAxe1 else
      if (Axe='A2') then FFic:=fbAxe2 else
      if (Axe='A3') then FFic:=fbAxe3 else
      if (Axe='A4') then FFic:=fbAxe4 else
      if (Axe='A5') then FFic:=fbAxe5 ;
      END ;
Val:=T.FindField(Compte).AsString ;
If CurNat='FGE' Then Val:=BourreEtLess(Val,fbGene) Else
  If CurNat='FTI' Then Val:=BourreEtLess(Val,fbAux) Else
    If CurNat='FSE' Then Val:=BourreEtLess(Val,FFic) ;
Adr1:='' ; Adr2:='' ; Adr3:='' ; CodePost:='' ; Ville:='' ; Tele:='' ;

{JP 03/01/06 : FQ AGL 11668 : Corrigé avec l'agl 5.8.1.161 : la fonction RecupPrefixSansDBO
               est devenue inutile}
chp := TableToPrefixe(THTable(T).TableName)+'_LIBELLE';
//chp := RecupPrefixSansDBO(THTable(T).TableName, '_LIBELLE');
Lib:=T.FindField(Chp).AsString ;
////Lib:=T.FindField(TableToPrefixe(THTable(T).TableName)+'_LIBELLE').AsString ;

If CurNat='FTI' Then
  BEGIN
  Adr1:=T.FindField('T_ADRESSE1').AsString ;
  Adr2:=T.FindField('T_ADRESSE2').AsString ;
  Adr3:=T.FindField('T_ADRESSE3').AsString ;
  CodePost:=T.FindField('T_CODEPOSTAL').AsString ;
  Ville:=T.FindField('T_VILLE').AsString ;
  Tele:=T.FindField('T_TELEPHONE').AsString ;
  END ;

OkExiste:=TRUE ;
if CurNat='FSE' then
  BEGIN
  if PresenceComplexe('SECTION',['S_AXE','S_SECTION'],['=','='],[Axe,Val],['S','S']) then T.Cancel Else OkExiste:=FALSE ;
  END else
if CurNat='FBS' then
  BEGIN
  if PresenceComplexe('BUDSECT',['BS_AXE','BS_BUDSECT'],['=','='],[Axe,Val],['S','S']) then T.Cancel Else OkExiste:=FALSE ;
  END else if Presence(THTable(T).TableName,Compte,Val) then T.Cancel Else OkExiste:=FALSE ;

If OkExiste Then
  BEGIN
  if (CurNat='FSE') or (CurNat='FBS') then THTable(T).FindKey([Axe,Val])
                                      else THTable(T).FindKey([Val]) ;
  END ;

if OkExiste And (not T.Eof) then
  BEGIN
  T.Edit ;
//  T.FindField(TableToPrefixe(THTable(T).TableName)+'_LIBELLE').AsString:=Lib ;
  T.FindField(RecupPrefixSansDBO(THTable(T).TableName, '_LIBELLE')).AsString := Lib;
  If CurNat='FTI' Then
    BEGIN
    If Adr.Checked Then
      BEGIN
      T.FindField('T_ADRESSE1').AsString:=Adr1 ;
      T.FindField('T_ADRESSE2').AsString:=Adr2 ;
      T.FindField('T_ADRESSE3').AsString:=Adr3 ;
      T.FindField('T_CODEPOSTAL').AsString:=CodePost ;
      T.FindField('T_VILLE').AsString:=Ville ;
      END ;
    If Tel.Checked Then
      BEGIN
      T.FindField('T_TELEPHONE').AsString:=Tele ;
      END ;
    END ;
  Inc(NbModifies) ;
  T.Post ;
  Result:=True ;
  END else
  BEGIN
  T.FindField(Compte).AsString:=Val ;
  T.Post ;
  END ;
END ;

procedure TFImpCpte.QuelleTable ;
BEGIN
if CurNat='FBG' then T.TableName:='BUDGENE' else
 if CurNat='FBS' then T.TableName:='BUDSECT' else
  if CurNat='FBJ' then T.TableName:='BUDJAL' else
   if CurNat='FGE' then T.TableName:='GENERAUX' else
    if CurNat='FTI' then T.TableName:='TIERS' else
     if CurNat='FSE' then T.TableName:='SECTION' ;
{JP 18/07/05 : ici, il n'est pas nécessaire de faire un RecupPrefixSansDBO, car "DBO." n'est
               rajouté que dans le BeforeOpen}
T.IndexName:=TableToPrefixe(T.TableName)+'_CLE1' ;
END ;

procedure TFImpCpte.ImporteFormatParametre ;
var Entete   : TFmtEntete ;
    Detail   : TTabFmtDetail ;
    Debut    : Boolean ;
    Fichier  : TextFile ;
    f : String ;
BEGIN
if CurNat='FBG' then f:='BG_BUDGENE' else
 if CurNat='FBS' then f:='BS_BUDSECT' else
  if CurNat='FBJ' then f:='BJ_BUDJAL' else
   if CurNat='FGE' then f:='G_GENERAL' else
    if CurNat='FTI' then f:='T_AUXILIAIRE' else
     if CurNat='FSE' then f:='S_SECTION' ;
if not ChargeFormat(Fichier,Filename.Text,CurNat,'X',FFormat.Value,Entete,Detail,Debut) then Exit ;
MajSQLDefaut ;
InitMove(NbImport,'') ;
NbImport:=0 ;
While not Eof(Fichier) do
  BEGIN
  if TestBreak then Break ;
  T.Insert ; InitNew(T) ;
  LireFormat(Fichier,Entete,Detail,Debut,T) ;
  Try
    if ModifLeCompte(T) then
      BEGIN
      MajCompteRendu(1,False,T.FindField(f).AsString,HErr.Mess[3],0) ;
      END else
      BEGIN
      Inc(NbImport) ;
      MajCompteRendu(0,False,T.FindField(f).AsString,HErr.Mess[5],0) ;
      END ;
    finally
    MoveCur(false) ;
    End ;
  END ;
T.Close ;
FiniMove ;
CloseFile(Fichier) ;
END ;

procedure TFImpCpte.BValiderClick(Sender: TObject);
Var F : TextFile ;
    ImpParam : boolean ;
    St : String ;
BEGIN
if (FileName.text='') then BEGIN Msgbox.Execute(6,caption,'') ; Exit ; END ;
if (FFormat.ItemIndex=-1) then BEGIN MsgBox.Execute(17,caption,'') ; Exit ; END ;
if FFormat.Value='SAB' then Lequel:='T' ;
AssignFile(F,FileName.Text) ;
{$I-} Reset (F) ; {$I+}
if IoResult<>0 then BEGIN Msgbox.Execute(7,caption,'') ; Exit ; END ;
CloseFile(F) ;
ImpParam:=((FFormat.Value<>'SAA') and (FFormat.Value<>'SAB') and (FFormat.Value<>'EDI')
          and (FFormat.Value<>'HAL') and (FFormat.Value<>'CLB')and (FFormat.Value<>'CPR')
          and (FFormat.Value<>'CGN')
          and (FFormat.Value<>'CVI')) or (Lequel='FBG') or (Lequel='FBS') or (Lequel='FBJ') ;
if (not ImpParam) and Not FormatOk(FileName.Text,FFormat.Value,Lequel) then BEGIN Msgbox.Execute(32,caption,'') ; Exit ; END ;
AssignFile(F,FileName.Text) ;
{$I-} Reset (F) ; {$I+}
if IoResult<>0 then BEGIN Msgbox.Execute(7,caption,'') ; Exit ; END ;
ActivePanels(Self,False,False) ;
NbImport:=0 ; NbErreur:=0 ; NbModifies:=0 ;
ReadLn(F,St) ;
While not EOF(F) do BEGIN NbImport:=NbImport+1 ; Readln(F,St) ; END ;
CloseFile(F) ; Arreter:=False ;
if ImpParam then BEGIN ImporteFormatParametre ; END else ImporteCptes ;
//  if (Transactions(ImporteCptes,1)<>OeOk) then MessageAlerte(HErr.Mess[6]) ;
ActivePanels(Self,True,False) ;
if (NbModifies=0) and (NbImport=0) then Exit ;
AfficheCompteRendu ;
if ChoixFmt.Detruire then DeleteFile(FileName.text) ;
Screen.cursor:=SynCrDefault ;
END;

{----------------------------- Lancement de l'importation ----------------------------------}
(*
Procedure CreateListe(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.LGenLu:=TStringList.Create ;
InfoImp.LAuxLu:=TStringList.Create ;
InfoImp.LAnaLu:=TStringList.Create ;
InfoImp.LJalLu:=TStringList.Create ;
InfoImp.LMP:=TStringList.Create ;
InfoImp.LMR:=TStringList.Create ;
InfoImp.LRGT:=TStringList.Create ;
InfoImp.ListeCptFaux:=TList.Create ;
InfoImp.ListePieceFausse:=TList.Create ;
InfoImp.ListeEntetePieceFausse:=TStringList.Create ;
InfoImp.ListeEnteteDoublon:=TStringList.Create ;
InfoImp.ListePieceIntegre:=TStringList.Create ;
InfoImp.CRListeEnteteDoublon:=TList.Create ;
END ;
*)
Function AlimInfoImp(Var InfoImp :TInfoImport) : Boolean ;
BEGIN
FillChar(InfoImp,SizeOf(InfoImp),#0) ; CreateListeImp(InfoImp) ;
InfoImp.Lequel:='FEC' ; InfoImp.Format:='CGN' ;
ChargeScenarioImport(InfoImp,FALSE) ;
END ;


procedure TFImpCpte.ImporteCptes ;
var St : String ;
    CptColl : PtTTabCptLu ;
    QT : TQuery ;
    FFic : TFichierBase ;
    InfoImp : PtTInfoImport ;
    QFiche : TQFiche ;
    i : Integer ;
BEGIN
try
  If FFormat.Value='CGN' Then
    BEGIN
    New(InfoImp) ; AlimInfoImp(InfoImp^) ;
    For i:=0 To 3 Do InitRequete(QFiche[i],i) ;
    If InfoImp.Sc.UseCorresp Then InitRequete(QFiche[5],5) ;
    END ;
  InitMove(NbImport,'') ;
  MoveCur(False) ;
//  ChargeJalSoc(1,CptColl) ;
  TErreur.Clear ; NbImport:=0 ;
  AssignFile(Fichier,FileName.Text) ;
  {$I-} Reset (Fichier) ; {$I+}
  Readln(Fichier,St) ;
  OkLecture:=True ;
  SourisSablier ;
  If FFormat.Value<>'CGN' Then QT:=MajSQLDefaut ;
  While Not EOF(Fichier) do
    BEGIN
    if TestBreak then Break ;
    MoveCur(False) ;
    // Readln fait dans le RecupNatureCpte pour L'EDI
    if OkLecture then Readln(Fichier,St) ;
    if ChoixFmt.Ascii then St:=Ascii2Ansi(St) ;
    OkNew:=False ;
    If FFormat.Value<>'CGN' Then
      BEGIN
      if (FFormat.Value<>'EDI') then
        BEGIN
        case Controles(St,FFic) of
          1   : ; // Modification.
          255 : break ;
          else BEGIN OkNew:=True ; T.Insert ; InitNew(T) ; END ;
          end ;
        END else if not ImporteEDI(St) then Break ;
      END ;
    if (FFormat.Value='CGN') then ImporteCpteCEGID(St,InfoImp^,QFiche) else
      if (FFormat.Value='HAL') then ImporteHalley(St,FFic) else
        if (Lequel[1]='T') then ImporteTiers(St) else
          BEGIN
          New(CptColl) ; ChargeJalSoc(1,CptColl^) ;
          ImporteCpteSAARI(St,CptColl) ;
          Dispose(CptColl) ;
          END ;
    END ;
  finally
  CloseFile(Fichier) ; If FFormat.Value<>'CGN' Then T.Close ;
  If FFormat.Value='CGN' Then
    BEGIN
    VideListeInfoImp(InfoImp^,TRUE) ; Dispose(InfoImp) ;
    For i:=0 To 3 Do Ferme(QFiche[i]) ;
    If InfoImp.Sc.UseCorresp Then Ferme(QFiche[5]) ;
    END ;
  FiniMove ;
  SourisNormale ;
  end ;
END ;

{----------------------------- Format CEGID ----------------------------------}

Function TFImpCpte.ImporteCpteCEGID(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : Boolean ;
Var OkOk : Boolean ;
    ResImp : tResultImportCpte ;
    Val : String ;
BEGIN
Result:=FALSE ;
Case Lequel[1] Of
  'G' : OkOk:=EstUneLigneCptOk(St,'CGE') ;
  'X' : OkOk:=EstUneLigneCptOk(St,'CAU') Or EstUneLigneCptOk(St,'CAE');
  'A' : OkOk:=EstUneLigneCptOk(St,'SAN') Or EstUneLigneCptOk(St,'SSA');
  Else OkOk:=FALSE ;
  END ;
If OkOk Then
  BEGIN
  Result:=TRUE ; Val:=Trim(Copy(St,7,17)) ;
  ResImp:=TraiteImportCompte(St,InfoImp,QFiche) ;
  Case ResImp Of
    resCreer    : BEGIN
                  Inc(NbImport) ;
                  MajCompteRendu(0,False,Val,Val,17) ;
                  END ;
    resModifier : BEGIN
                  Inc(NbModifies) ;
                  MajCompteRendu(1,False,Val,Val,17) ;
                  END ;
    END ;
  END ;
END ;


{----------------------------- Format SAARI ----------------------------------}

Function TFImpCpte.ImporteCpteSAARI(St : String ; CptColl : PtTTabCptLu) : Boolean ;
Var TypeCpte:String ;
    LeCpt : String ;
    LeCptColl : TCptLu ;
BEGIN
Result:=False ;
TypeCpte:=Copy(St,1,1) ;
If ((Lequel[1]<>TypeCpte) and (Lequel[1]<>'T')) then Exit ;
Case TypeCpte[1] of
  'X' : TypeCpte:='T' ;
  'A' : TypeCpte:='S' ;
  END ;
if OkNew or Lib.checked then
  BEGIN
  T.FindField(TypeCpte+'_LIBELLE').AsString:=Trim(Copy(St,15,31)) ;
  T.FindField(TypeCpte+'_ABREGE').AsString:=Trim(Copy(St,15,17)) ;
  END ;
Case Lequel[1] of
  'G' :BEGIN
       if OkNew then
         BEGIN
         LeCpt:=Trim(Copy(St,2,13)) ;
         If LgComptes.Checked Then LeCpt:=BourreEtLess(LeCpt,fbGene) ;
         T.FindField('G_GENERAL').AsString:=LeCpt ;
         T.FindField('G_NATUREGENE').AsString:=Trim(Copy(St,64,1)) ;
         T.FindField('G_CREERPAR').AsString:='IMP' ;
         if (Copy(St,69,1)<>'0') then T.FindField('G_CONFIDENTIEL').AsString:='0' ;
         END ;
       if OkNew or Car.Checked then
         BEGIN
         if (Copy(St,46,1)='C') then T.FindField('G_COLLECTIF').AsString:='X' else T.FindField('G_COLLECTIF').AsString:='-' ;
         if (Copy(St,46,1)='P') or (Copy(St,47,1)='P') then T.FindField('G_POINTABLE').AsString:='X' else T.FindField('G_POINTABLE').AsString:='-' ;
         if (Copy(St,46,1)='E') or (Copy(St,47,1)='E') or (Copy(St,48,1)='E') then T.FindField('G_CENTRALISABLE').AsString:='X' else T.FindField('G_CENTRALISABLE').AsString:='-' ;
         if (Copy(St,46,1)='V') or (Copy(St,47,1)='V') or (Copy(St,48,1)='V') or (Copy(St,49,1)='V') then T.FindField('G_VENTILABLE').AsString:='X' else T.FindField('G_VENTILABLE').AsString:='-' ;
         if (Copy(St,46,1)='L') or (Copy(St,47,1)='L') or (Copy(St,48,1)='L') or (Copy(St,49,1)='L') then T.FindField('G_LETTRABLE').AsString:='X' else T.FindField('G_LETTRABLE').AsString:='-' ;
         T.FindField('G_SENS').AsString:=Copy(St,63,1) ;
         END ;
       (*T.FindField('G_ADRESSE1').AsString:=Copy(St,2,13) ;
       T.FindField('G_ADRESSE2').AsString:=Copy(St,2,13) ;
       T.FindField('G_CODEPOSTAL').AsString:=Copy(St,2,13) ;
       T.FindField('G_VILLE').AsString:=Copy(St,2,13) ;
       T.FindField('G_TELEPHONE').AsString:=Copy(St,2,13) ;*)
       END ;
  'X' :BEGIN
       if OkNew then
         BEGIN
         LeCpt:=Trim(Copy(St,2,13)) ;
         If LgComptes.Checked Then LeCpt:=BourreEtLess(LeCpt,fbAux) ;
         T.FindField('T_AUXILIAIRE').AsString:=LeCpt ;
         T.FindField('T_CREERPAR').AsString:='IMP' ;
         if (Copy(St,69,1)<>'0') then T.FindField('T_CONFIDENTIEL').AsString:='0' ;
         END ;
       //T.FindField('G_SENS').AsString:=Copy(St,46,1) ;
       //T.FindField('T_NATUREAUXI').AsString:=Copy(St,2,13) ;
       if OkNew or Car.Checked then
         BEGIN
         if (Copy(St,47,1)='L') or (Copy(St,48,1)='L') then T.FindField('T_LETTRABLE').AsString:='X' else T.FindField('T_LETTRABLE').AsString:='-' ;
         END ;
       //if (Copy(St,47,P)='P') or (Copy(St,48,1)='P') then T.FindField('G_POINTABLE').AsString:=Copy(St,2,13) ;
       if OkNew or Coll.Checked then T.FindField('T_COLLECTIF').AsString:=BourreEtLess(Trim(Copy(St,50,13)),fbGene) ;
       if OkNew or Adr.Checked then
         BEGIN
         T.FindField('T_ADRESSE1').AsString:=Trim(Copy(St,63,21)) ;
         T.FindField('T_ADRESSE2').AsString:=Trim(Copy(St,84,21)) ;
         T.FindField('T_CODEPOSTAL').AsString:=Trim(Copy(St,105,6)) ;
         T.FindField('T_VILLE').AsString:=Trim(Copy(St,111,21)) ;
         END ;
       if OkNew or Tel.Checked then T.FindField('T_TELEPHONE').AsString:=Trim(Copy(St,132,15)) ;
       If OkNew Then
          BEGIN
          T.FindField('T_REGIMETVA').AsString:='FRA' ;
          T.FindField('T_MODEREGLE').AsString:='CHQ' ;
          LeCptColl:=TrouveJalSoc(BourreEtLess(Trim(Copy(St,50,13)),fbGene),CptColl^) ;
          If LeCptColl.Cpt<>'' Then
             BEGIN
             If LeCptColl.Nature='COC' Then T.FindField('T_NATUREAUXI').AsString:='CLI' ;
             If LeCptColl.Nature='COF' Then T.FindField('T_NATUREAUXI').AsString:='FOU' ;
             If LeCptColl.Nature='COS' Then T.FindField('T_NATUREAUXI').AsString:='SAL' ;
             If LeCptColl.Nature='COD' Then T.FindField('T_NATUREAUXI').AsString:='DIV' ;
             END ;
          END ;
       END ;
  'A' :BEGIN
       if OkNew then
         BEGIN
         T.FindField('S_SECTION').AsString:=Trim(Copy(St,2,13)) ;
         T.FindField('S_AXE').AsString:='A1' ;
         T.FindField('S_CREERPAR').AsString:='IMP' ;
         if (Copy(St,69,1)<>'0') then T.FindField('S_CONFIDENTIEL').AsString:='0' ;
         END ;
       if OkNew or Car.Checked then T.FindField('S_SENS').AsString:=Trim(Copy(St,46,1)) ;
      END ;
  END ;
T.Post ;
Result:=True ;
END ;

{----------------------------- Format SAARI Banque ----------------------------------}

procedure TFImpCpte.ImporteTiers (St : String) ;
Var LeCpt : String ;
BEGIN
if OkNew then
  BEGIN
  T.FindField('T_EAN').AsString:=Copy(St,1,13) ;
  LeCpt:=Trim(Copy(St,38,13)) ; LeCpt:=BourreEtLess(LeCpt,fbAux) ;
  T.FindField('T_AUXILIAIRE').AsString:=LeCpt ;
  T.FindField('T_CREERPAR').AsString:='IMP' ;
  END ;
if OkNew or Lib.Checked then T.FindField('T_LIBELLE').AsString:=Copy(St,14,24) ;
if OkNew or Coll.Checked then T.FindField('T_COLLECTIF').AsString:=Copy(St,51,13) ;
if OkNew or Adr.Checked then
  BEGIN
  T.FindField('T_ADRESSE1').AsString:=Copy(St,64,24) ;
  T.FindField('T_ADRESSE2').AsString:=Copy(St,88,24) ;
  T.FindField('T_CODEPOSTAL').AsString:=Copy(St,112,8) ;
  T.FindField('T_VILLE').AsString:=Copy(St,120,24) ;
  END ;
T.Post ;
END ;

{------------------- Importation des comptes au Format HALLEY --------------------------}

function OkValExiste (OkCCO : boolean ; StCode,StWhere,StPrefixe,StTable,Val : String ) : boolean ;
var Q  : TQuery ;
    P  : Integer ;
    SQL,StOrder : String ;
BEGIN
Result:=True ;
if (Val='') then Exit ;
if OKCCO then SQL:='Select '+StCode+', '+StPrefixe+'_LIBELLE, '+StPrefixe+'_ABREGE FROM '+StTable
            else SQL:='Select '+StCode+', '+StPrefixe+'_LIBELLE FROM '+StTable ;
   if StWhere<>'' then SQL:=SQL+' Where '+StWhere+' AND '+StCode+'="'+Val+'"' else SQL:=SQL+' Where '+StCode+'="'+Val+'"' ;
   P:=Pos('DISTINCT',StCode) ; if P>0 then StOrder:=Copy(StCode,p+9,50) else StOrder:=StCode ;
   if OKCCO then SQL:=SQL+' ORDER BY '+StPrefixe+'_TYPE, '+StOrder
            else SQL:=SQL+' ORDER BY '+StOrder ;
Q:=OpenSQL(SQL,TRUE) ;
Result:=not (Q.EOF) ;
Ferme(Q) ;
END ;

function RecupJoin(NomChp : String ; MultiTT : string ) : string ;
BEGIN
Result:='' ;
if (NomChp='CREERPAR') or (NomChp='EXPORTE') or (NomChp='UTILISATEUR') or (NomChp='SOCIETE') then Exit ;
if (NomChp='CLEREPARTITION') then
 BEGIN
 if MultiTT='A1' then Result:='ttCleRepart1' else
  if MultiTT='A2' then Result:='ttCleRepart2' else
   if MultiTT='A3' then Result:='ttCleRepart3' else
    if MultiTT='A4' then Result:='ttCleRepart4' else
     if MultiTT='A5' then Result:='ttCleRepart5' ;
 END else Result:=Get_Join('X_'+NomChp) ;
END ;

function MajValCombo(ttCorr: string ; var Where : String ; Val,Lib,Abr : String) : byte ;
var St,StCode,StWhere,StTable,StPrefixe,StLib : String ;
    OkCCO : boolean ;
    i,First : integer ;
    StCle,StVal : Array[1..3] of String ;

BEGIN
Result:=0 ; Where:='' ;
if (ttCorr<>'') then
    BEGIN
    GetCorrespType(ttCorr,StTable,StCode,StWhere,StPrefixe,StLib) ;
    if ((StWhere<>'') AND (Pos('="'+W_W+'"',StWhere)>0)) then
      BEGIN
      StWhere:=FindEtReplace(StWhere,'="'+W_W+'"','<>"'+W_W+'"',TRUE) ;
      END ;
    OkCCO:=((StTable='COMMUN') or (StTable='CHOIXCOD')) ;
    if not OkValExiste(OkCCO,StCode,StWhere,StPrefixe,StTable,Val) then
      BEGIN
      St:=StWhere ;
      StCle[1]:=StCode ;
      StVal[1]:='"'+Val+'"' ;
      i:=1 ;
      While (St<>'') and (i<3) do
        BEGIN
        i:=i+1 ;
        First:=Pos('="',St) ;
        StCle[i]:=Copy(St,1,First-1) ;
        St:=Copy(St,First+2,Length(St)-First+3) ;
        First:=Pos('"',St) ;
        StVal[i]:='"'+Copy(St,1,First-1)+'"' ;
        St:=Copy(St,First+1,Length(St)-First+1) ;
        END ;
      StWhere:=Copy(StVal[2],2,Length(StVal[2])-2) ;
      StCle[3]:=StLib ;
      StVal[3]:='"'+Lib+'"' ;
//Simon      if OkCCO then AddValCombo((StTable='COMMUN'),StWhere,Val,Lib,Abr,'') else
      ValueInsert(StTable,StCle,StVal,TRUE) ;
      Result:=1 ;
      if OkCCO then Result:=Result+1 ;
      Where:=StWhere ;
      END ;
    END ;
END ;

procedure TFImpCpte.ImporteHalley (St : String ; FFic : TFichierBase) ;
var i, lg,deb,PPref : integer ;
    AxeEnCours,Cle,NomChp,Pref,Where,NomChamp : String ;
    PrendLe : boolean ;
    Res : Byte ;
BEGIN
Deb:=1 ; LG:=1 ; AxeEnCours:='' ; Cle:='' ; Res:=0 ;
for i:=0 to T.FieldCount-1 do
  BEGIN
  Case T.Fields[i].DataType of
    ftString  : BEGIN
                Lg:=T.Fields[i].DataSize-1 ;
                if (Pos(T.Fields[i].FieldName,'DERNLETTRAGE')<>0) or
                (Pos(T.Fields[i].FieldName,'DATEPAQUET')<>0) then BEGIN Deb:=Deb+Lg ; Continue ; END ;
                NomChamp:=T.Fields[i].FieldName ;
                END ;
    ftBlob : Continue ;  // Lg:=70 else
    ftDate,ftDateTime : Lg:=10 ;
    ftTime : Lg:=8 ;
    //ftDateTime : Lg:=19 ;
    ftAutoInc,ftInteger,ftSmallInt,ftWord : BEGIN
                                            Lg:=9 ;
                                            if (Pos(T.Fields[i].FieldName,'NUMDERNMVT')<>0) then BEGIN Deb:=Deb+Lg ; Continue  ; END ;
                                            END ;
    ftBCD,ftCurrency,ftFloat : BEGIN
                               Lg:=15  ;
                               if (T.Fields[i].FieldName<>'T_COEFCOMMR') and (T.Fields[i].FieldName<>'T_COEFCOMMA')
                                and (T.Fields[i].FieldName<>'T_REMISE') and (T.Fields[i].FieldName<>'T_FRANCO')
                                and (T.Fields[i].FieldName<>'T_ESCOMPTE') then
                                  BEGIN
                                  if Deb+Lg>1024 then
                                    BEGIN
                                    if Copy(St,1024,1)=SepLigneIE then
                                      BEGIN
                                      Readln(Fichier,St) ;
                                      Deb:=1 ;
                                      END ;
                                    END Else Deb:=Deb+Lg ;
                                  Continue ;
                                  END ;
                               END ;
    ftBoolean : Lg:=1 ;
    else continue ;
    END ;
  if Deb+Lg>1024 then
    BEGIN
    if Copy(St,1024,1)=SepLigneIE then
      BEGIN
      Readln(Fichier,St) ;
      Deb:=1 ;
      END else Continue ;
    END ;
  PPref:=Pos('_',T.Fields[i].FieldName) ;
  Pref:=Copy(T.Fields[i].FieldName,1,PPref) ;
  NomChp:=Copy(T.Fields[i].FieldName,PPref+1,Length(T.Fields[i].FieldName)-PPRef) ;
  if (NomChp='CREERPAR') then T.Fields[i].Value:='IMP' else
    BEGIN
    PrendLe:=False ;
    if not OkNew then
      BEGIN
      if (Lib.Checked and (NomChp='LIBELLE')) then PrendLe:=True ;
      if (Coll.Checked and (NomChp='COLLECTIF') and ((Lequel='X') or (Lequel='T'))) then PrendLe:=True ;
      if (Car.Checked and ((NomChp='SENS') or (NomChp='LETTRABLE') or (NomChp='POINTABLE')
        or (NomChp='VENTILABLE') or (NomChp='CENTRALISABLE') or (NomChp='MODEREGLE')
        or ((NomChp='COLLECTIF') and (Lequel='G')))) then PrendLe:=True ;
      if (Tel.Checked and (NomChp='TELEPHONE')) then PrendLe:=True ;
      if (Adr.Checked and ((Pos('ADRESSE',NomChp)<>0) or (NomChp='CODEPOSTAL')
      or (NomChp='VILLE') or (NomChp='PAYS'))) then PrendLe:=True ;
      END ;
    if OkNew or PrendLe then
      BEGIN
      if (Trim(Copy(St,Deb,Lg))<>'') then
        BEGIN
        case T.Fields[i].DataType of
          ftBCD,ftCurrency,ftFloat : T.Fields[i].Value:=Valeur(Trim(Copy(St,Deb,Lg))) ;
          else BEGIN
               If ((NomChamp='G_GENERAL') And (FFic=fbGene)) Or ((NomChamp='T_AUXILIAIRE') And (FFic=fbAux)) Or
                  ((NomChamp='S_SECTION') And (FFic In [fbAxe1..fbAxe5])) Then T.Fields[i].Value:=BourreEtLess(Trim(Copy(St,Deb,Lg)),FFic)
                                                                          Else T.Fields[i].Value:=Trim(Copy(St,Deb,Lg)) ;
               END ;
          END ;
        //T.Fields[i].Value:=Trim(Copy(St,Deb,Lg)) ;
        if (NomChp='AXE') then AxeEnCours:=Trim(Copy(St,Deb,Lg)) else
         if (NomChp='CLEREPARTITION') then Cle:=Trim(Copy(St,Deb,Lg)) else
           // combos uniquement
           if Lg=3 then
             BEGIN
             Res:=MajValCombo(RecupJoin(NomChp,''),Where,Trim(Copy(St,Deb,Lg)),HErr.Mess[11],HErr.Mess[12]) ;
             AjouteErrCombo(Pref+NomChp,Where,Trim(Copy(St,Deb,Lg)),Res,T.Fields[0].Value) ;
             END ;
         END ;
      END ;
    END ;
  Deb:=Deb+Lg ;
  END ;
// Spécif aux clés de répartition
if AxeEnCours<>'' then Res:=MajValCombo(RecupJoin('CLEREPARTITION',AxeEnCours),Where,Cle,HErr.Mess[11],HErr.Mess[12]) ;
if Cle<>'' then AjouteErrCombo(Pref+NomChp,Where,Cle,Res,T.Fields[0].Value) ;
T.Post ;
END ;


procedure TFImpCpte.AjouteErrCombo(NomChp,Where,Val : String ; Res : Byte ; Code : Variant) ;
var ErrMess : String ;
    Q : Tquery ;
    OkCCO : boolean ;
BEGIN
if Res=0 then Exit ;
OkCCO:=(Res=2) ;
if OkCCO then
  BEGIN
  ErrMess:=TTTypeToLibelle(Where)+' "'+Val+'" '+HErr.Mess[10] ;
  END else
  BEGIN
  Q:=OPENSQL('SELECT DH_LIBELLE FROM DECHAMPS WHERE DH_NOMCHAMP="'+NomChp+'"',True) ;
  if not Q.Eof then ErrMess:=Q.Fields[0].AsString+' "'+Val+'" '+HErr.Mess[10] ;
  Ferme(Q) ;
  END ;
//MajCompteRendu(4,False,VarAsType(Code,VarString),ErrMess,0) ;
MajCompteRendu(4,False, Code,ErrMess,0) ;
END ;
{----------------------------- Format EDIFICAS ----------------------------------}

function TFImpCpte.ImporteEDI(var St : String) : boolean ;
var Pref, Compte, LECpt : String ;
    FFic : TFichierBase ;
BEGIN
Result:=False ;
// Detail uniquement...( pour l'instant ?).
if (StrToInt(Copy(St,1,5))<2200) or (StrToInt(Copy(St,1,5))>2300) then BEGIN Result:=True ; Exit ; END ;
case Controles(St,FFic) of
  1 : ;
  255 : Exit ;
  else BEGIN OkNew:=True ; T.Insert ; InitNew(T) ; END ;
  end ;
Case NumTypCpt of
   CptGEN : BEGIN Pref:='G_' ; Compte:='G_GENERAL' ; LeCpt:=Trim(Copy(St,9,17)) ; LeCpt:=BourreEtLess(LeCpt,fbGene) ; END ;
   CptAUX : BEGIN Pref:='T_' ; Compte:='T_AUXILIAIRE' ; LeCpt:=Trim(Copy(St,9,17)) ; LeCpt:=BourreEtLess(LeCpt,fbAux) ; END ;
   CptANA : BEGIN Pref:='S_' ; Compte:='S_SECTION' ; LeCpt:=Trim(Copy(St,9,17)) ; LeCpt:=BourreEtLess(LeCpt,fbAxe1) ; END ;
   END ;
if OkNew then
  BEGIN
//  T.FindField(Compte).AsString:=Trim(Copy(St,9,17)) ;
  T.FindField(Compte).AsString:=LeCpt ;
  T.FindField(Pref+'ABREGE').AsString:=Trim(Copy(St,32,17)) ;
  T.FindField(Pref+'DATECREATION').AsString:=Format_Date(Copy(St,125,8)) ;
  T.FindField(Pref+'DATEMODIF').AsString:=Format_Date(Copy(St,125,8)) ;
  T.FindField(Pref+'DATEOUVERTURE').AsString:=Format_Date(Copy(St,125,8)) ;
  T.FindField(Pref+'DATEFERMETURE').AsString:=Format_Date(Copy(St,133,8)) ;
  if NumTypCpt=CptANA then T.FindField(Pref+'AXE').AsString:='A1' ;
  T.FindField(Pref+'CREERPAR').AsString:='IMP' ;
  if NumTypCpt=CptAUX then T.FindField(Pref+'DEVISE').AsString:=Copy(St,119,3) ;
  END ;
if OkNew or Lib.Checked then T.FindField(Pref+'LIBELLE').AsString:=Trim(Copy(St,49,35)) ;
if OkNew or Car.Checked then if Copy(St,122,3)='DOF' then T.FindField(Pref+'FERME').AsString:='X' else T.FindField(Pref+'FERME').AsString:='-' ;
SousDetail(St) ;
T.Post ;
Result:=True ;
END;

Procedure TFImpCpte.SousDetail(var St : String) ;
var RecupNatCpte,NatureCpte : String ;
    Pref : String ;
    OkNat : String1 ;

    function RecupNat(Var St : String) : String ;
    var OkNature : String1 ;
    BEGIN
    Result:='' ;
    OkNature :='-' ;
    Readln(Fichier,St) ;
    MoveCur(False) ;
    if (StrToInt(Copy(St,1,5))=2250) then BEGIN OkNature:='X' ; Readln(Fichier,St) ; MoveCur(False) ; END ;
    if EOF(Fichier) then Exit ;
    //Changement de compte  ;
    if (StrToInt(Copy(St,1,5))<>2300) then
      BEGIN
      if (StrToInt(Copy(St,1,5))=2200) then OkLecture:=False else if not OkLecture then OkLecture:=True  ;
      Exit ;
      END ;
    Result:=Copy(St,26,3)+OkNature ;
    END ;

BEGIN
Case NumTypCpt of
   cptGEN : BEGIN Pref:='G_' ; NatureCpte:='G_NATUREGENE' ; END ;
   CptAUX : BEGIN Pref:='T_' ; NatureCpte:='T_NATUREAUXI' ; END ;
   CptANA : BEGIN Pref:='S_' ; NatureCpte:='' ; END ;
   END ;
RecupNatCpte:=RecupNat(St) ;
Repeat
  OkNat:=Copy(RecupNatCpte,4,1) ;
  RecupNatCpte:=Copy(RecupNatCpte,1,3) ;
  if (NumTypCpt=CptAUX) and (OkNew or Coll.Checked) then T.FindField(Pref+'COLLECTIF').AsString:=Trim(Copy(St,38,17)) ;
  if OkNew or Car.Checked then
    BEGIN
    if (OkNat='X') then T.FindField(NatureCpte).AsString:=RecupNatCpte ;
    if NumTypCpt=CptGEN then
      BEGIN
      if (RecupNatCpte='CCE') then T.FindField(Pref+'CENTRALISABLE').AsString:='X' else T.FindField(Pref+'CENTRALISABLE').AsString:='-' ;
      if (RecupNatCpte='CPO') then T.FindField('G_POINTABLE').AsString:='X' else T.FindField('G_POINTABLE').AsString:='-' ;
      if (RecupNatCpte='CVE') then T.FindField(Pref+'VENTILABLE').AsString:='X' else T.FindField(Pref+'VENTILABLE').AsString:='-' ;
      if (RecupNatCpte='CCO') then T.FindField(Pref+'COLLECTIF').AsString:='X' else T.FindField(Pref+'COLLECTIF').AsString:='-' ;
      END ;
    if (NumTypCpt=CptGEN) or (NumTypCpt=CptAUX) then
      if (RecupNatCpte='CLE') or (RecupNatCpte='CEC') then T.FindField(Pref+'LETTRABLE').asString:='X' else T.FindField(Pref+'LETTRABLE').asString:='-' ;
    if NumTypCpt<>CptAUX then
      if (RecupNatCpte='CCS') then T.FindField(Pref+'SENS').AsString:='C' else
       if (RecupNatCpte='CDS') then T.FindField(Pref+'SENS').AsString:='D' else
        if (RecupNatCpte='CCD') then T.FindField(Pref+'SENS').AsString:='M' ;
    END ;
 //            if RecupNatCpte='CBU' then
//            if RecupNatCpte='COR' then
  RecupNatCpte:=RecupNat(St) ;
  until (RecupNatCpte='') ;
END ;

{----------------------------- Requêtes et Contrôles ----------------------------------}

Function TFImpCpte.MajSQLDefaut : TQuery ;
var Tabl,St : String ;
    QT : TQuery ;
BEGIN
Result:=NIL ;
if Lequel='G' then BEGIN Tabl:='GENERAUX' ; END else
 if Lequel='A' then BEGIN Tabl:='SECTION' ; END else
  if (Lequel='T') or (lequel='X') then BEGIN Tabl:='TIERS' ; END else
   if Lequel='FBJ' then BEGIN Tabl:='BUDJAL' ; END else
    if Lequel='FBG' then BEGIN Tabl:='BUDGENE' ; END else
     if Lequel='FBS' then BEGIN Tabl:='BUDSECT' ; END Else Exit ;
If OnREQ Then
  BEGIN
  (*
  St:='SELECT * FROM' ;
  QT:=OpenSQL(ST,FALSE) ;
  *)
  END Else
  BEGIN
  T.Close; {JP 18/07/05 : FQ 15587}
  T.TableName:=Tabl ;
  T.Open ;
  END ;
END ;

function TFImpCpte.Controles(St : String ; Var FFic : TFichierBase) : byte ;
var Axe,Compte,CpteColl,
    WhereAxe,ValColl,Tabl,Val : string ;
    Ok   : boolean ;
    Q    : TQuery ;
    Err  : byte ;
    Mss  : Integer ;
    LeRIB  : String ;
    NatTiers : String3 ;
    CodeAxe : String ;
BEGIN
Result:=0 ; Err:=0 ; Axe:='' ; WhereAxe:='' ; CodeAxe:='' ;
if ((FFormat.Value='EDI') and (Copy(St,6,3)<>'ECG') and (Copy(St,6,3)<>'ECD')
  and (Copy(St,6,3)<>'ECA')) then Exit ;
{
T.Close ;
T.SQL.Clear ;
}
Mss:=0 ; FFic:=fbGene ;
if Lequel='G' then
   BEGIN
   Compte:='G_GENERAL' ;
   Tabl:='GENERAUX' ; Mss:=22 ;
   END else
if Lequel='A' then
   BEGIN
   Compte:='S_SECTION' ;
   Axe:=',S_AXE' ;
   Tabl:='SECTION' ; Mss:=24 ;
   //if (FFormat.Value='HAL') then WhereAxe:=' AND S_AXE="'+Trim(Copy(St,278,3))+'" ' ;
   END else
if (Lequel='T') or (lequel='X') then
   BEGIN
   Compte:='T_AUXILIAIRE' ;
   CpteColl:='G_GENERAL' ;
   Tabl:='TIERS' ; Mss:=23 ;
   FFic:=fbAux ;
   END else
if Lequel='FBJ' then
   BEGIN
   Compte:='BJ_BUDJAL' ;
   Tabl:='BUDJAL' ; Mss:=25 ;
   END else
if Lequel='FBG' then
   BEGIN
   Compte:='BG_BUDGENE' ;
   Tabl:='BUDGENE' ; Mss:=26 ;
   END else
if Lequel='FBS' then
   BEGIN
   Compte:='BS_BUDSECT' ;
   Tabl:='BUDSECT' ; Mss:=27 ;
   END ;
{
T.SQL.Add('SELECT * FROM '+Tabl) ;
}
if (FFormat.Value='EDI') then
  BEGIN
  Val:=Copy(St,9,17) ;
  END else
  if (FFormat.Value='HAL') then
    BEGIN
    Val:=Copy(St,1,17) ; If Axe<>'' Then CodeAxe:=Trim(Copy(St,248,3)) ;
    END else if (FFormat.Value='SAB') then Val:=Copy(St,38,13) else Val:=Copy(St,2,13) ;
Val:=UpperCase(Trim(Val)) ;
if (Axe<>'') then
  BEGIN
  if (CodeAxe='A1') then FFic:=fbAxe1 else
  if (CodeAxe='A2') then FFic:=fbAxe2 else
  if (CodeAxe='A3') then FFic:=fbAxe3 else
  if (CodeAxe='A4') then FFic:=fbAxe4 else
  if (CodeAxe='A5') then FFic:=fbAxe5 ;
  If CodeAxe<>'' Then WhereAxe:=' AND S_AXE="'+CodeAxe+'" ' ;
  END ;
if (Axe<>'') then
  BEGIN
  Q:=OpenSQL('SELECT '+Compte+Axe+' FROM '+Tabl+' WHERE '+ Compte+'="'+Val+'"'+WhereAxe,True) ;
  if Not Q.Eof then CodeAxe:=Q.Fields[1].AsString ;
  If Trim(CodeAxe)='' Then CodeAxe:='A1' ;
  Ferme(Q) ;
  if (CodeAxe='A1') then FFic:=fbAxe1 else
  if (CodeAxe='A2') then FFic:=fbAxe2 else
  if (CodeAxe='A3') then FFic:=fbAxe3 else
  if (CodeAxe='A4') then FFic:=fbAxe4 else
  if (CodeAxe='A5') then FFic:=fbAxe5 ;
  END ;
// Longueur des comptes
if FFic in [fbAxe1..fbAux] then Val:=BourreEtLess(Val,FFic) ;
 if Length(Trim(Val))<>VH^.Cpta[FFic].Lg then
  BEGIN
  MsgBox.Execute(30,'',' '+MsgBox.Mess[Mss]+' '+MsgBox.Mess[31]) ;
  Result:=255 ;
  Exit ;
  END ;
// L'enregistrement a déja été créé.

if CodeAxe<>'' then
  BEGIN
  if PresenceComplexe('SECTION',['S_AXE','S_SECTION'],['=','='],[CodeAxe,Val],['S','S']) then Err:=1
  END else if Presence(Tabl,Compte,Val) then Err:=1 ;

// Modification
if (Err=1) then
  BEGIN
  T.Cancel ;
  if CodeAxe<>'' then
    BEGIN
    if T.FindKey([CodeAxe,Val]) then Err:=1 ;
    END else if T.Findkey([Val]) then Err:=1 ;
  Ok:=False ;
  if Lib.Checked or Car.Checked then Ok:=True else
    BEGIN
    case lequel[1] of
      'X' : if (Tel.Checked or Adr.Checked) and (FFormat.Value<>'EDI') then Ok:=True ;
      'T' : if Adr.Checked or RIB.Checked then Ok:=True ;
      END ;
    END ;
  if Ok then NbModifies:=NbModifies+1 ;
  //T.SQL.Add(' WHERE '+Compte+'="'+Trim(Val)+'"') ;
  //if Axe<>'' then T.SQL[1]:=T.SQL[1]+' AND S_AXE="'+Axe+'"' ;
  //ChangeSQL(T) ; T.Open ;
  T.Edit ;
  END else
  BEGIN
  MajCompteRendu(Err,False,Val,Val,0) ;
  NbImport:=NbImport+1 ;
  //T.SQL.Add(' WHERE '+Compte+'="'+W_W+'"') ;
  //ChangeSQL(T) ; T.Open ;
  END ;

// Tiers
if (Lequel[1]='X') or (Lequel[1]='T') then
  BEGIN
  if (FFormat.Value<>'EDI') then
    BEGIN
    if (Lequel='T') then
      BEGIN
      LeRIB:=EncodeLeRIB(Copy(St,166,5),Copy(St,171,5),Copy(St,176,11),Copy(St,187,2),Copy(St,146,20)) ;
      TraiteRIBImport(Val,LeRIB,Trim(Copy(St,120,24)),Trim(Copy(St,190,3)),'-') ;
      END ;
(*
      if not Presence('RIB','R_AUXILIAIRE',Val) then
        BEGIN
        Err:=3 ;
        MajCompteRendu(Err,False,Val,Val) ;
        ExecuteSQL('INSERT INTO RIB (R_AUXILIAIRE,R_NUMERORIB,R_DOMICILIATION,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_PRINCIPAL)'+
        ' VALUES("'+Val+'",'+IntToStr(1)+',"'+Copy(St,146,20)+'","'+Copy(St,166,5)+'","'+Copy(St,171,5)+'","'+Copy(St,176,11)+'","'+Copy(St,187,2)+'","X")') ;
        END else
        // RIB Déja créé et Nouvel enregistrement (Err<>1)
        if (Err<>1) or (RIB.Checked) then
           BEGIN
           ExecuteSQL('UPDATE RIB SET R_NUMERORIB='+IntToStr(1)+',R_DOMICILIATION="'+Copy(St,146,20)+
           '",R_ETABBQ="'+Copy(St,166,5)+'",R_GUICHET="'+Copy(St,171,5)+
           '",R_NUMEROCOMPTE="'+Copy(St,176,11)+'",R_CLERIB="'+Copy(St,187,2)+'",R_PRINCIPAL="X"'+
           'WHERE R_AUXILIAIRE="'+Copy(St,38,13)+'"') ;
           END ;
*)
    // Compte Collectif de l'auxiliaire
    if (FFormat.Value='HAL') then ValColl:=BourreEtLess(Copy(St,73,17),fbGene) else
      if (FFormat.Value='SAA') then ValColl:=BourreEtLess(Copy(St,50,13),fbGene) else
        if (FFormat.Value='SAB') then ValColl:=BourreEtLess(Copy(St,51,13),fbGene) ;
    if Coll.Checked or (Err=0) then
      if not Presence('GENERAUX',CpteColl,Trim(ValColl)) and (Trim(ValColl)<>'') then
        BEGIN
        MsgBox.Execute(33,'',' ('+ValColl+') ') ;
        Result:=255 ;
        Exit ;
        END ;
        (*
        MajCompteRendu(Err,False,Val,ValColl) ;
        NatTiers:='COD' ;
        if Copy(ValColl,1,2)='40' then NatTiers:='COF'else
         if Copy(ValColl,1,2)='41' then NatTiers:='COC' ;
        ExecuteSQL('INSERT INTO GENERAUX (G_GENERAL,G_LIBELLE,G_ABREGE,G_COLLECTIF,G_LETTRABLE,G_NATUREGENE)'+
        ' VALUES("'+Trim(ValColl)+'","'+HErr.Mess[11]+'","'+HErr.Mess[12]+'","X","X","'+NatTiers+'")') ;
        *)
    END ;
  END ;
Result:=Err ;
END ;

Function TFImpCpte.MajCompteRendu(NumErr : Integer ; Bloquant : Boolean ; Code,Valeur : String ; j : Integer) : Boolean ;
Var X : DelInfo ;
    NMess : String ;
    i : integer ;
BEGIN
Result:=Bloquant ;
if not ChoixFmt.CompteRendu then Exit ;
Valeur:=Trim(Valeur) ;
NbErreur:=NbErreur+1 ;
X:=DelInfo.Create ; X.LeCod:=IntToStr(NbErreur) ;
//if Bloquant then X.LeLib:='X' else X.LeLib:='-' ;
X.LeLib:=Code ;
case NumErr of
  0,1:BEGIN // Création / Existant
      if CurNat='FGE' then NMess:=HErr.Mess[0] else
       if CurNat='FTI' then NMess:=HErr.Mess[1] else
        if CurNat='FSE' then NMess:=HErr.Mess[2] else
         if CurNat='FBG' then NMess:=HErr.Mess[14] else
          if CurNat='FBS' then NMess:=HErr.Mess[15] else
           if CurNat='FBJ' then NMess:=HErr.Mess[16] ;
      if NumErr=0 then i:=5 else If j<>0 Then i:=j Else i:=3 ;
      NMess:=NMess+' '+Valeur+' '+HErr.Mess[i] ;
      END ;
  2:NMess:=HErr.Mess[7]+' '+Valeur+' '+HErr.Mess[10] ; // Collectif Créé
  3:NMess:=HErr.Mess[4]+' '+Valeur+' '+HErr.Mess[5] ;  // RIB Créé
  4:NMess:=Valeur ;
  END ;
X.LeMess:=NMess ;
TErreur.Add(X) ;
END ;

procedure TFImpCpte.FFormatChange(Sender: TObject);
begin
if FFormat.Value='EDI' then Sauve.FilterIndex:=4 else
  if FFormat.Value='HAL' then Sauve.FilterIndex:=1 else
    if (FFormat.Value='SAA') then Sauve.FilterIndex:=3 ;
ModifSurExistant ;
end;

procedure TFImpCpte.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;


//SG6 13/01/05 FQ 15242
procedure TFImpCpte.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFImpCpte.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //SG6 13/01/05 FQ 15242
  if key = VK_ESCAPE then BFermeClick(nil);
end;

procedure TFImpCpte.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

{JP 18/07/05 : en attendant que l'agl gère le dbo dans TableToPrefix : cf FQ AGL 11668
               en effet, l'agl rajoute "dbo." à la propriété TableName dans le OnBeforeOpen
{---------------------------------------------------------------------------------------}
function TFImpCpte.RecupPrefixSansDBO(NomTable, NomChamp : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := TableToPrefixe(FindEtReplace(LowerCase(NomTable), 'dbo.', '', True));
  Result := Result + NomChamp;
end;

end.


