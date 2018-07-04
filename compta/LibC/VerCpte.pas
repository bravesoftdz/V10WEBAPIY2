unit VerCpte;
(* Document Intégrateur n° 60017_91.doc (G:\client\0\7)
Attention si GesCom<>'' alors FicheTiers Run depuis la gescom
                           sinon FicheTiers Run depuis la Compta
Ceci pour le datatype de la nature tiers "ttNatTiers ou ttNatTiersCpta"
et les différents controle sur collectif
*)

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  hmsgbox, StdCtrls, Buttons, Hctrls, ExtCtrls, Ent1, HEnt1,
  HStatus, Hcompte, Rapsuppr, HSysMenu;

procedure VerifCompte ;

type  TINFOSCPTE = Record
      Code    : String17 ;
      Libelle : String35 ;
      Axe     : String ;
      Pref    : Char ;
      END ;

type  TINF = Class
      If1     : String17 ;
      If2     : String35 ;
      END ;

type
  TFVerifCpte = class(TForm)
    Panel1: TPanel;
    TFVerification: THLabel;
    Panel2: TPanel;
    TNBError1: TLabel;
    TNBError2: TLabel;
    TNBError3: TLabel;
    TNBError5: TLabel;
    TNBError4: TLabel;
    FVerification: TComboBox;
    HPB: TPanel;
    TTravail: TLabel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BStop: THBitBtn;
    MsgBar: THMsgBox;
    MsgRien: THMsgBox;
    QCG: TQuery;
    MsgLibel: THMsgBox;
    MsgLibel2: THMsgBox;
    MsgCor: THMsgBox;
    TCModR: THValComboBox;
    TCRegTVA: THValComboBox;
    TCTvaEnc: THValComboBox;
    TCCpte: TComboBox;
    TCTypeConPart: THValComboBox;
    TCDev: THValComboBox;
    QCT: TQuery;
    QCS: TQuery;
    QCJ: TQuery;
    TCnaGen: THValComboBox;
    TCnaAux: THValComboBox;
    TCnaJal: THValComboBox;
    TCRelTraite: THDBValComboBox;
    TCRelRegle: THDBValComboBox;
    HMTrad: THSystemMenu;
    CbEcr: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BValiderClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Déclarations privées }
    LaListe                  : TList ;
    OkVerif, StopVerif, MAJ,
    TousGenOk, TousAuxOk,
    TousSecOk, TousJalOK,
    ParLaForme, EnMvt        : Boolean ;
    NbEnreg, NbError, ErrGene,
    ErrTiers,ErrSect, ErrJal,
    Manuel                   : Integer ;
    QuelleVerif              : Byte ;
    Cp                       : TINFOSCPTE ;
    ListeGENCOL              : TStringList ;
    ListeGENFIN              : TStringList ;
    Function GoListe1(ID : TINFOSCPTE ; Rem : String ; I : Byte) : DelInfo ;
    Function GoListe2(ID : TINFOSCPTE ; Ch : TField ; Rem : String) : DelInfo ;
    procedure Corrige(Champ, Valeur : String) ;
    procedure TuFaisQuoi(Montre : TINFOSCPTE) ;
    function  TestBreak : Boolean ;
    Function CompteATraiter : Boolean ;
    procedure LanceVerif ;
    procedure InitLabel ;
    procedure SqlGene ;
    procedure SQLTiers ;
    procedure SQLSect ;
    procedure SQLJal ;
    Function VerifCode : Boolean ;
    Function VerifASCII : Boolean ;
    Function VerifLibel : Boolean ;
    Function VerifNatureGene : Boolean ;
    Function VerifBool(Q : TQuery) : Boolean ;
    Function Bool(Ch : String) : Boolean ;
    Function OKCorresp(Code,Compte : String) : Boolean ;
    Function VerifCorresp(Q : TQuery) : Boolean ;
    Function VerifQualifQte : Boolean ;
    Function VerifTaxe : Boolean ;
    Function VerifBanque : Boolean ;
    Function VerifDernLet(Q : TQuery) : Boolean ;
    Function VerifTicTid : Boolean ;
    Function VerifGeneCol : Boolean ;
    Function VerifNatureAuxi : Boolean ;
    Function VerifTVA : Boolean ;
    Function VerifReglenemt : Boolean ;
    Function VerifNatJal : Boolean ;
    Function VerifFacturier : Boolean ;
    Function VerifContrePartie : Boolean ;
    Function VerifRegul : Boolean ;
    Function ExisteCode(A : Byte ; Code : String ; Var St : String) : Boolean ;
    procedure CorrigeASCII ;
    procedure ResouCaracGene(A : Byte) ;
    procedure DonneCollectif ;
    procedure DonneContrePartie ;
    procedure FlagueCompte ;
    Function  CompteOk : Boolean ;
    Procedure ChargeInfos ;
    Procedure VideInfos ;
    Procedure MAJBQEDEVISE(St : String) ;
  public
    EnBatch : boolean ;
  end;

function VerCompte(Y : Byte ; EnBatch : boolean ) : boolean ; { 0, Tous ; 1, Généraux ; 2, Tiers ; 3, Section ; 4, Journaux }
Procedure VerCompteMAJ(Y : Byte ; Var NbPasPossible : Integer ; EnControl : Boolean) ;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  CPGeneraux_TOM ;

{$R *.DFM}

procedure VerifCompte ;
var VCpte : TFVerifCpte ;
BEGIN
VCpte:=TFVerifCpte.Create(Application) ;
try
 VCpte.EnBatch:=False ;
 VCpte.FVerification.ItemIndex:=0 ;
 VCpte.FVerification.Enabled:=True ;
 VCpte.ShowModal ;
 finally
 VCpte.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFVerifCpte.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
//FVerification.ItemIndex:=0 ;
InitLabel ;
QuelleVerif:=200 ; ParLaForme:=True ;
end;

procedure TFVerifCpte.FormCreate(Sender: TObject);
begin
LaListe:=TList.Create ;
end;

procedure TFVerifCpte.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//
end;

procedure TFVerifCpte.BValiderClick(Sender: TObject);
begin
OkVerif:=True ; NbError:=0 ; MAJ:=False ;
TousGenOk:=True ; TousAuxOk:=True ; TousSecOk:=True ; TousJalOk:=True ;
ErrGene:=0 ; ErrTiers:=0 ; ErrSect:=0 ;
LaListe.Clear ; StopVerif:=False ; Manuel:=0;
TNBError1.Caption:=MsgBar.Mess[11] ;
EnableControls(Self,False) ;
Application.ProcessMessages ;
if CompteATraiter then
   BEGIN
   QuelleVerif:=FVerification.ItemIndex ;
   ChargeInfos ;
   LanceVerif ;
   if Not OkVerif then RapportdErreurMvt(Laliste,4,MAJ,False) Else MsgRien.Execute(1,caption,'') ; // sinon message tout est ok
   END Else MsgRien.Execute(0,caption,'') ;
if MAJ then
   BEGIN
   InitLabel ; StopVerif:=False ;
   LanceVerif ;
   if Manuel<>0 then
      BEGIN
      MsgRien.Execute(4,IntToStr(Manuel)+' ','');
      END else MsgRien.Execute(3,caption,'') ;
   END ;
Caption:=MsgBar.Mess[15] ;
VideInfos ;
InitLabel ;
EnableControls(Self,True) ;
end;

Procedure TFVerifCpte.ChargeInfos ;
Var Q : TQuery ; inf : TINF ;
BEGIN
ListeGENCOL:=TStringList.Create ;
ListeGENFIN:=TStringList.Create ;
Q:=OpenSql('select G_GENERAL, G_NATUREGENE from GENERAUX '
          +'where G_NATUREGENE like "CO%" order by G_GENERAL',True) ;
While Not Q.Eof do
      BEGIN
      Inf:=TINF.Create ;
      Inf.If1:=Q.FindField('G_GENERAL').AsString ;
      Inf.If2:=Q.FindField('G_NATUREGENE').AsString ;
      ListeGENCOL.AddObject('',Inf) ; Q.Next ;
      END ;
Ferme(Q) ;
Q:=OpenSql('select G_GENERAL, G_NATUREGENE from GENERAUX '
          +'where G_NATUREGENE="BQE" or G_NATUREGENE="CAI" order by G_GENERAL',True) ;
While Not Q.Eof do
      BEGIN
      Inf:=TINF.Create ;
      Inf.If1:=Q.FindField('G_GENERAL').AsString ;
      Inf.If2:=Q.FindField('G_NATUREGENE').AsString ;
      ListeGENFIN.AddObject('',Inf) ; Q.Next ;
      END ;
Ferme(Q) ;
TCCpte.clear ;  // Init d'Enreg des cptes faux
END ;

Procedure TFVerifCpte.VideInfos ;
Var i : integer ; inf : TINF ;
BEGIN
If ListeGENCOL<>NIL Then
  BEGIN
  For i:=0 to ListeGENCOL.Count-1 do
     BEGIN
     inf:=TINF(ListeGENCOL.Objects[i]) ;
     inf.Free ;
     END ;
  END ;
If ListeGENCOL<>Nil then begin ListeGENCOL.Clear ; ListeGENCOL.Free ; ListeGENCOL:=NIL ; end ;
If ListeGENFIN<>NIL THen
  BEGIN
  For i:=0 to ListeGENFIN.Count-1 do
     BEGIN
     inf:=TINF(ListeGENFIN.Objects[i]) ;
     inf.Free ;
     END ;
  END ;
If ListeGENFIN<>Nil then begin ListeGENFIN.Clear ; ListeGENFIN.Free ; ListeGenFin:=NIL ; end ;
TCCpte.clear ;
END ;

function TFVerifCpte.TestBreak : Boolean ;
BEGIN
Application.ProcessMessages ;
if StopVerif then if MsgRien.Execute(2,caption,'')<>mryes then StopVerif:=False ;
Result:=StopVerif ;
END ;

Function TFVerifCpte.CompteATraiter : Boolean ;
Var Rien : Boolean ;
BEGIN
Rien:=FALSE ;
If (QuelleVerif=FVerification.ItemIndex) then begin Result:=true ; exit ; end ;
NbEnreg:=0 ;
QCG.Close ; QCT.Close ;QCS.Close ;QCJ.Close ;
QCG.SQL.Clear ; QCT.SQL.Clear ; QCS.SQL.Clear ; QCJ.SQL.Clear ;
QCG.RequestLive:=False ; QCT.RequestLive:=False ; QCS.RequestLive:=False ; QCJ.RequestLive:=False ;
Case FVerification.ItemIndex of   {tous}
  0 : BEGIN
      QCG.SQL.Add('Select * from GENERAUX order by G_GENERAL') ;
      ChangeSql(QCG) ; //QCG.Prepare ;
      PrepareSQLODBC(QCG) ;
      QCG.Open ; Rien:=QCG.Eof ;
      QCT.SQL.Add('Select * from TIERS order by T_AUXILIAIRE') ;
      ChangeSql(QCT) ; //QCT.Prepare ;
      PrepareSQLODBC(QCT) ;
      QCT.Open ; Rien:=Rien And QCT.Eof ;
      QCS.SQL.Add('Select * from SECTION order by S_AXE, S_SECTION') ;
      ChangeSql(QCS) ; //QCS.Prepare ;
      PrepareSQLODBC(QCS) ;
      QCS.Open ; Rien:=Rien And QCS.Eof ;
      QCJ.SQL.Add('Select * from JOURNAL order by J_JOURNAL') ;
      ChangeSql(QCJ) ; //QCJ.Prepare ;
      PrepareSQLODBC(QCJ) ;
      QCJ.Open ; Rien:=Rien And QCJ.Eof ;
      InitMove(RecordsCount(QCG),MsgBar.Mess[0]) ;
      END ;
  1 : BEGIN
      QCG.SQL.Add('Select * from GENERAUX order by G_GENERAL') ;
      ChangeSql(QCG) ; //QCG.Prepare ;
      PrepareSQLODBC(QCG) ;
      QCG.Open ; Rien:=QCG.Eof ;
      InitMove(RecordsCount(QCG),MsgBar.Mess[0]) ;
      END ;
  2 : BEGIN
      QCT.SQL.Add('Select * from TIERS order by T_AUXILIAIRE') ;
      ChangeSql(QCT) ; //QCT.Prepare ;
      PrepareSQLODBC(QCT) ;
      QCT.Open ; Rien:=QCT.Eof ;
      InitMove(RecordsCount(QCT),MsgBar.Mess[0]) ;
      END ;
  3 : BEGIN
      QCS.SQL.Add('Select * from SECTION order by S_AXE, S_SECTION') ;
      ChangeSql(QCS) ; //QCS.Prepare ;
      PrepareSQLODBC(QCS) ;
      QCS.Open ; Rien:=QCS.Eof ;
      InitMove(RecordsCount(QCS),MsgBar.Mess[0]) ;
      END ;
  4 : BEGIN
      QCJ.SQL.Add('Select * from JOURNAL order by J_JOURNAL') ;
      ChangeSql(QCJ) ; //QCJ.Prepare ;
      PrepareSQLODBC(QCJ) ;
      QCJ.Open ; Rien:=QCJ.Eof ;
      InitMove(RecordsCount(QCJ),MsgBar.Mess[0]) ;
      END ;
 End ;
CompteATraiter:=Not Rien ;
END ;

procedure TFVerifCpte.LanceVerif ;
BEGIN
TNBError1.Caption:='' ;
If MAJ then Caption:=MsgBar.Mess[14] ;
Application.ProcessMessages ;
Case FVerification.ItemIndex of {Tous}
  0 : BEGIN
      if not StopVerif then SqlGene ;
      if not StopVerif then SqlTiers ;
      if not StopVerif then SqLSect ;
      if not StopVerif then SqLJal ;
      END ;
  1 : if not StopVerif then SqlGene ;
  2 : if not StopVerif then SqlTiers ;
  3 : if not StopVerif then SqLSect ;
  4 : if not StopVerif then SqLJal ;
  End ;
FiniMove ;
END ;


procedure TFVerifCpte.InitLabel ;
BEGIN
TTravail.Caption:='' ; TNBError1.Caption:='' ; TNBError2.Caption:='' ;
TNBError3.Caption:='' ; TNBError4.Caption:='' ;TNBError5.Caption:='' ;
TNBError1.Font.color:=clNavy ; TNBError2.Font.color:=clNavy ;
TNBError3.Font.color:=clNavy ; TNBError4.Font.color:=clNavy ;
TNBError5.Font.color:=clNavy ;
END ;

Function TFVerifCpte.GoListe1(ID : TINFOSCPTE ; Rem : String ; I : Byte) : DelInfo ;
Var X : DelInfo ;
BEGIN
Inc(NbError) ;
X:=DelInfo.Create ;
X.LeLib:=ID.Libelle ; X.LeMess2:=Traduirememoire(MsgLibel.Mess[i])+' '+Rem ;
X.LeMess:=ID.Pref ; X.LeCod:=ID.Code ;
IF ID.Pref='S' then X.LeMess3:=ID.Axe ;
Result:=X ;
END ;

Function TFVerifCpte.GoListe2(ID : TINFOSCPTE ; Ch : TField ; Rem : String) : DelInfo ;
Var X : DelInfo ;
BEGIN
Inc(NbError) ;
X:=DelInfo.Create ;
X.LeLib:=ID.Libelle ; X.LeMess2:=TraduireMemoire(MsgLibel2.Mess[0])+' "'+Ch.FieldName+'" '+TraduireMemoire(MsgLibel2.Mess[1])+Rem ;
X.LeMess:=ID.Pref ; X.LeCod:=ID.Code ;
IF ID.Pref='S' then X.LeMess3:=ID.Axe ;
Result:=X ;
END ;

Function TFVerifCpte.Bool(Ch : String) : Boolean ;
BEGIN
Result:=((Ch='X')or(Ch='-')) ;
END ;

Function TFVerifCpte.OKCorresp(Code,Compte : String) : Boolean ;
BEGIN
Result:=PresenceComplexe('CORRESP',['CR_TYPE','CR_CORRESP'],['=','='],[Code,Compte],['S','S']) ;
END ;

procedure TFVerifCpte.BStopClick(Sender: TObject);
begin
StopVerif:=True ;
end;

procedure TFVerifCpte.TuFaisQuoi(Montre : TINFOSCPTE) ;
BEGIN
If Not ParLaForme then exit ;
Case FVerification.ItemIndex of   {tous}
  0 : BEGIN
      If Montre.Pref='G' then
         BEGIN
         TTravail.Caption:=MsgBar.Mess[1]+' '+Montre.Code+' '+
                           MsgBar.Mess[2]+' '+Montre.Libelle ;
         IF Not MAJ then
            BEGIN
            ErrGene:=NbError ;
            if ErrGene=0 then TNBError1.Caption:=MsgBar.Mess[7]+'... '+MsgBar.Mess[6] ;
            if ErrGene=1 then TNBError1.Caption:=MsgBar.Mess[7]+'... '+IntToStr(ErrGene)+' '+MsgBar.Mess[5] ;
            if ErrGene>1 then TNBError1.Caption:=MsgBar.Mess[7]+'... '+IntToStr(ErrGene)+' '+MsgBar.Mess[4] ;
            if ErrGene>=1 then TNBError1.Font.color:=clRed ;
            END Else TNBError1.Caption:=MsgBar.Mess[7] ;
         END Else
      If Montre.Pref='T' then
         BEGIN
         TTravail.Caption:=MsgBar.Mess[1]+' '+Montre.Code+' '+
                           MsgBar.Mess[2]+' '+Montre.Libelle ;
         IF Not MAJ then
            BEGIN
            ErrTiers:=NbError-ErrGene ;
            if ErrTiers=0 then TNBError2.Caption:=MsgBar.Mess[8]+'... '+MsgBar.Mess[6] ;
            if ErrTiers=1 then TNBError2.Caption:=MsgBar.Mess[8]+'... '+IntToStr(ErrTiers)+' '+MsgBar.Mess[5] ;
            if ErrTiers>1 then TNBError2.Caption:=MsgBar.Mess[8]+'... '+IntToStr(ErrTiers)+' '+MsgBar.Mess[4] ;
            if ErrTiers>=1 then TNBError2.Font.color:=clRed ;
            END Else TNBError2.Caption:=MsgBar.Mess[8] ;
         END Else
      If Montre.Pref='S' then
         BEGIN
         TTravail.Caption:=MsgBar.Mess[1]+' '+Montre.Code+' '+
                           MsgBar.Mess[2]+' '+Montre.Libelle ;
         If Not MAJ then
            BEGIN
            ErrSect:=NbError-ErrGene-ErrTiers ;
            if ErrSect=0 then TNBError3.Caption:=MsgBar.Mess[9]+'... '+MsgBar.Mess[6] ;
            if ErrSect=1 then TNBError3.Caption:=MsgBar.Mess[9]+'... '+IntToStr(ErrSect)+' '+MsgBar.Mess[5] ;
            if ErrSect>1 then TNBError3.Caption:=MsgBar.Mess[9]+'... '+IntToStr(ErrSect)+' '+MsgBar.Mess[4] ;
            if ErrSect>=1 then TNBError3.Font.color:=clRed ;
            END Else TNBError3.Caption:=MsgBar.Mess[9] ;
         END Else
      If Montre.Pref='J' then
         BEGIN
         TTravail.Caption:=MsgBar.Mess[12]+' '+Montre.Code+' '+
                           MsgBar.Mess[2]+' '+Montre.Libelle ;
         IF Not MAJ then
            BEGIN
            ErrJal:=NbError-ErrGene-ErrTiers-ErrSect ;
            if ErrJal=0 then TNBError4.Caption:=MsgBar.Mess[13]+'... '+MsgBar.Mess[6] ;
            if ErrJal=1 then TNBError4.Caption:=MsgBar.Mess[13]+'... '+IntToStr(ErrJal)+' '+MsgBar.Mess[5] ;
            if ErrJal>1 then TNBError4.Caption:=MsgBar.Mess[13]+'... '+IntToStr(ErrJal)+' '+MsgBar.Mess[4] ;
            if ErrJal>=1 then TNBError4.Font.color:=clRed ;
            END Else TNBError4.Caption:=MsgBar.Mess[13] ;
         END ;
      END ;
  1 : BEGIN
      TTravail.Caption:=MsgBar.Mess[1]+' '+Montre.Code+' '+
                        MsgBar.Mess[2]+' '+Montre.Libelle ;
      IF Not MAJ then
         BEGIN
         ErrGene:=NbError ;
         if ErrGene=0 then TNBError1.Caption:=MsgBar.Mess[7]+'... '+MsgBar.Mess[6] ;
         if ErrGene=1 then TNBError1.Caption:=MsgBar.Mess[7]+'... '+IntToStr(ErrGene)+' '+MsgBar.Mess[5] ;
         if ErrGene>1 then TNBError1.Caption:=MsgBar.Mess[7]+'... '+IntToStr(ErrGene)+' '+MsgBar.Mess[4] ;
         if ErrGene>=1 then TNBError1.Font.color:=clRed ;
         END Else TNBError1.Caption:=MsgBar.Mess[7] ;
      END ;
  2 : BEGIN
      TTravail.Caption:=MsgBar.Mess[1]+' '+Montre.Code+' '+
                        MsgBar.Mess[2]+' '+Montre.Libelle ;
      IF Not MAJ then
         BEGIN
         ErrTiers:=NbError-ErrGene ;
         if ErrTiers=0 then TNBError1.Caption:=MsgBar.Mess[8]+'... '+MsgBar.Mess[6] ;
         if ErrTiers=1 then TNBError1.Caption:=MsgBar.Mess[8]+'... '+IntToStr(ErrTiers)+' '+MsgBar.Mess[5] ;
         if ErrTiers>1 then TNBError1.Caption:=MsgBar.Mess[8]+'... '+IntToStr(ErrTiers)+' '+MsgBar.Mess[4] ;
         if ErrTiers>=1 then TNBError1.Font.color:=clRed ;
         END Else TNBError1.Caption:=MsgBar.Mess[8]
      END ;
  3 : BEGIN
      TTravail.Caption:=MsgBar.Mess[1]+' '+Montre.Code+' '+
                        MsgBar.Mess[2]+' '+Montre.Libelle ;
      If Not MAJ then
         BEGIN
         ErrSect:=NbError-ErrGene-ErrTiers ;
         if ErrSect=0 then TNBError1.Caption:=MsgBar.Mess[9]+'... '+MsgBar.Mess[6] ;
         if ErrSect=1 then TNBError1.Caption:=MsgBar.Mess[9]+'... '+IntToStr(ErrSect)+' '+MsgBar.Mess[5] ;
         if ErrSect>1 then TNBError1.Caption:=MsgBar.Mess[9]+'... '+IntToStr(ErrSect)+' '+MsgBar.Mess[4] ;
         if ErrSect>=1 then TNBError1.Font.color:=clRed ;
         END Else TNBError1.Caption:=MsgBar.Mess[9] ;
      END ;
  4 : BEGIN
      TTravail.Caption:=MsgBar.Mess[12]+' '+Montre.Code+' '+
                        MsgBar.Mess[2]+' '+Montre.Libelle ;
      If Not MAJ then
         BEGIN
         ErrJal:=NbError-ErrGene-ErrTiers-ErrSect ;
         if ErrJal=0 then TNBError1.Caption:=MsgBar.Mess[13]+'... '+MsgBar.Mess[6] ;
         if ErrJal=1 then TNBError1.Caption:=MsgBar.Mess[13]+'... '+IntToStr(ErrJal)+' '+MsgBar.Mess[5] ;
         if ErrJal>1 then TNBError1.Caption:=MsgBar.Mess[13]+'... '+IntToStr(ErrJal)+' '+MsgBar.Mess[4] ;
         if ErrJal>=1 then TNBError1.Font.color:=clRed ;
         END Else TNBError1.Caption:=MsgBar.Mess[13] ;
      END ;
 End ;
END ;

procedure TFVerifCpte.Corrige(Champ, Valeur : String) ;
BEGIN
if Cp.Pref='G' then begin QCG.Edit ; QCG.FindField(Champ).AsString:=Valeur ; QCG.Post ; end else
if Cp.Pref='T' then begin QCT.Edit ; QCT.FindField(Champ).AsString:=Valeur ; QCT.Post ; end else
if Cp.Pref='S' then begin QCS.Edit ; QCS.FindField(Champ).AsString:=Valeur ; QCS.Post ; end else
if Cp.Pref='J' then begin QCJ.Edit ; QCJ.FindField(Champ).AsString:=Valeur ; QCJ.Post ; end ;
END ;

procedure TFVerifCpte.CorrigeASCII ;
Var i : Byte ;
BEGIN
for i:=1 to Length(Cp.Code) do
    BEGIN
    if Not (Cp.Code[i] in ['0'..'9','A'..'z',VH^.Cpta[fbGene].Cb]) then Cp.Code[i]:=VH^.Cpta[fbGene].Cb ;
    END ;
END ;

Procedure TFVerifCpte.MAJBQEDEVISE(St : String) ;
BEGIN
ExecuteSql('Update BANQUECP set BQ_DEVISE="'+V_PGI.DevisePivot+'" where BQ_GENERAL="'+St
          +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ') ; // 19/10/2006 YMO Multisociétés
//ExecuteSql('Update BANQUECP set BQ_DEVISE="'+TCDev.Values[0]+'" where BQ_GENERAL="'+St+'" ') ;
END ;

Function TFVerifCpte.ExisteCode(A : Byte ; Code : String ; Var St : String ) : Boolean ;
Var I : Integer ; inf : TINF ; Trouve : Boolean ;
BEGIN  {  }
Trouve:=False ;
if A=1 then
   BEGIN
   For i:=0 to ListeGENCOL.Count-1 do
       BEGIN
       if ListeGENCOL.Objects[i]<>Nil then
          BEGIN
          Inf:=TINF(ListeGENCOL.Objects[i]) ;
          Trouve:=(Inf.If1=Code) ;
          IF Trouve then Begin St:=Inf.If2 ; Break ; End ;
          END Else begin Break ; end ;
       END ;
   END Else
If A=2 then Trouve:=(TCModR.values.IndexOf(Code)>-1) else
If A=3 then Trouve:=(TCRegTVA.values.IndexOf(Code)>-1) else
If A=4 then Trouve:=(TCTvaEnc.values.IndexOf(Code)>-1) else
if A=5 then Trouve:=(TCTypeConPart.values.IndexOf(Code)>-1) else
if A=6 then
   BEGIN
   For i:=0 to ListeGENFIN.Count-1 do
       BEGIN
       if ListeGENFIN.Objects[i]<>Nil then
          BEGIN
          Inf:=TINF(ListeGENFIN.Objects[i]) ;
          Trouve:=(Inf.If1=Code) ;
          IF Trouve then Begin St:=Inf.If2 ; Break ; End ;
          END Else begin Break ; end ;
       END ;
   END ;
Result:=Trouve ;
END ;

procedure TFVerifCpte.DonneCollectif ;
Var NatAuxi, NatGene : String3 ; Inf : TINF ; Cpt : String ; i : integer ;
BEGIN
NatAuxi:=QCT.FindField('T_NATUREAUXI').AsString ; cpt:='' ;
If NatAuxi='CLI' then NatGene:='COC' else
If NatAuxi='FOU' then NatGene:='COF' else
If NatAuxi='SAL' then NatGene:='COS' else
If NatAuxi='AUC' then NatGene:='COD' else
If NatAuxi='AUD' then NatGene:='COD' else
If NatAuxi='DIV' then NatGene:='COD' ;
For i:=0 to ListeGENCOL.Count-1 do
    BEGIN
    if ListeGENCOL.Objects[i]<>Nil then
       BEGIN
       Inf:=TINF(ListeGENCOL.Objects[i]) ;
       IF (Inf.If2=NatGene) then Begin Cpt:=Inf.If1 ; Break ; End ;
       END Else begin Break ; end ;
    END ;
Corrige('T_COLLECTIF',Cpt) ;
END ;

procedure TFVerifCpte.DonneContrePartie ;
Var NatJal : String3 ; i : integer ; Inf : TINF ; Cpt : String ;
BEGIN
NatJal:=QCJ.FindField('J_NATUREJAL').AsString ; Cpt:='' ;
For i:=0 to ListeGENFIN.Count-1 do
    BEGIN
    if ListeGENFIN.Objects[i]<>Nil then
       BEGIN
       Inf:=TINF(ListeGENFIN.Objects[i]) ;
       if (Inf.If2=NatJal) then Begin Cpt:=Inf.If1 ; Break ; End ;
       END Else begin Break ; end ;
    END ;
Corrige('J_CONTREPARTIE',Cpt) ;
END ;

procedure TFVerifCpte.ResouCaracGene(A : Byte) ;
Var Trouve : Boolean ; i : Byte ;
BEGIN
if A=1 then
   BEGIN
   Corrige('G_LETTRABLE','-') ; Corrige('G_POINTABLE','-') ;
   END Else
if A=2 then
   BEGIN
   Trouve:=False ;
   for i:=1 to 5 do if QCG.FindField('G_VENTILABLE'+IntToStr(i)+'').AsString='X' then begin Trouve:=True ; Break end ;
   If Trouve then corrige('G_LETTRABLE','-') Else begin corrige('G_VENTILABLE','-') ; Corrige('G_LETTRABLE','-') ; end ;
   Exit ;
   END ;
// si Proposition Ok
//if A=1 then St:='G_LETTRABLE="X";G_POINTABLE="X"' ;
//FicheGeneMZS('',Cp.Code ,TaModifenSerie,0,St);
// -- - - - - -
//if (Na='BQE')or(Na='CAI')or(Na='DIV')or(Na='EXT') then Corrige('G_LETTRABLE','-') else
//if (Na='TIC')or(Na='TID') then Corrige('G_POINTABLE','-') else begin Corrige('G_LETTRABLE','-') ; Corrige('G_POINTABLE','-') end ;
END ;

procedure TFVerifCpte.FlagueCompte ;
BEGIN
if (Cp.Pref='S')or(Cp.Pref='J') then TCCpte.Items.Add(Cp.Axe+Cp.Code) else TCCpte.Items.Add(Cp.Code) ;
END ;

Function TFVerifCpte.CompteOk : Boolean ;
Var Cpt : String ;
BEGIN
Cpt:=Cp.Axe+Cp.Code ;
Result:=(TCCpte.Items[TCCpte.Items.IndexOf(Cpt)]<>Cpt) ;
END ;

procedure TFVerifCpte.SqlGene ;
BEGIN
If MAJ then
   BEGIN
   If TousGenOk then Exit ;
   QCG.Close ; QCG.SQL.Clear ;
   QCG.SQL.Add('Select Distinct E_GENERAL From ECRITURE order by E_GENERAL') ;
   ChangeSQL(QCG) ; QCG.Open ; CbEcr.Items.Clear ;
   While Not QCG.Eof do
      BEGIN
      CbEcr.Items.Add(QCG.Fields[0].AsString) ;
      QCG.Next ;
      END ;
   QCG.Close ; QCG.SQL.Clear ;
   QCG.SQL.Add('Select Distinct Y_GENERAL From ANALYTIQ order by Y_GENERAL') ;
   ChangeSQL(QCG) ; QCG.Open ; CbEcr.Values.Clear ;
   While Not QCG.Eof do
      BEGIN
      CbEcr.Values.Add(QCG.Fields[0].AsString) ;
      QCG.Next ;
      END ;
   QCG.Close ; QCG.SQL.Clear ;
   QCG.SQL.Add('Select * from GENERAUX order by G_GENERAL') ; QCG.RequestLive:=MAJ ;
   ChangeSQL(QCG) ; QCG.Open ; InitMove(RecordsCount(QCG),MsgBar.Mess[0]) ;
   END ;
Fillchar(Cp,SizeOf(Cp),#0) ; QCG.First ;
While Not QCG.Eof do
   BEGIN
   Cp.Code:=QCG.FindField('G_GENERAL').AsString ; Cp.Libelle:=QCG.FindField('G_LIBELLE').AsString ;
   Cp.Pref:='G' ;
   MoveCur(False) ;
   if MAJ then if CompteOk then begin QCG.Next ; continue ; end ;
   if MAJ then EnMvt:=(CbEcr.Items.IndexOf(Cp.Code)>=0) Or (CbEcr.Values.IndexOf(Cp.Code)>=0) ;
   if VerifCode then
      BEGIN
      If Not VerifASCII then begin OkVerif:=False ; TousGenOk:=False ; end ;
      If Not VerifLibel then begin OkVerif:=False ; TousGenOk:=False ; end ;
      if VerifNatureGene then
         BEGIN
         IF Not VerifCorresp(QCG) then begin OkVerif:=False ; TousGenOk:=False ; end ;
         IF Not VerifQualifQte then begin OkVerif:=False ; TousGenOk:=False; end ;
         IF Not VerifTaxe then begin OkVerif:=False ;TousGenOk:=False ; end ;
         IF Not VerifBanque then begin OkVerif:=False ; TousGenOk:=False; end ;
         IF Not VerifDernLet(QCG) then begin OkVerif:=False ; TousGenOk:=False ; end ;
         IF Not VerifTicTid then begin OkVerif:=False ; TousGenOk:=False ; end ;
         If Not VerifBool(QCG) then begin OkVerif:=False ; TousGenOk:=False ; end ;
         END Else begin OkVerif:=False ; TousGenOk:=False ; end ;
      END Else begin OkVerif:=False ; TousGenOk:=False ; end ;
   If Not MAJ then IF Not TousGenOk then FlagueCompte ;  { Enreg des Cptes Error }
   TuFaisQuoi(Cp) ; if TestBreak then Break ;
   QCG.Next ;
   END ;
END ;

procedure TFVerifCpte.SqLSect ;
BEGIN
If MAJ then
   BEGIN
   If TousSecOk then exit ;
   QCS.Close ; QCS.SQL.Clear ;
   QCS.SQL.Add('Select * from SECTION order by S_AXE, S_SECTION') ; QCS.RequestLive:=MAJ ;
   ChangeSQL(QCS) ; QCS.Open ; InitMove(RecordsCount(QCS),MsgBar.Mess[0]) ;
   END ;
Fillchar(Cp,SizeOf(Cp),#0) ; QCS.First ;
While Not QCS.Eof do
   BEGIN
   Cp.Code:=QCS.FindField('S_SECTION').AsString ; Cp.Libelle:=QCS.FindField('S_LIBELLE').AsString ;
   Cp.Axe:=QCS.FindField('S_AXE').AsString ; Cp.Pref:='S' ;
   MoveCur(False) ;
   if MAJ then if CompteOk then begin QCS.Next ; continue ; end ;
   if VerifCode then
      BEGIN
      If Not VerifLibel then begin OkVerif:=False ; TousSecOk:=False ; end ;
      IF Not VerifCorresp(QCS) then begin OkVerif:=False ; TousSecOk:=false ; end;
      If Not VerifBool(QCS) then begin OkVerif:=False ; TousSecOk:=false ; end ;
      END Else begin OkVerif:=False ; TousSecOk:=false ;end ;
   If Not MAJ then IF Not TousSecOk then
                      begin
                      FlagueCompte ;
                      end ; { Enreg Cptes Error }
   TuFaisQuoi(Cp) ; if TestBreak then Break ;
   QCS.Next ;
   END ;
END ;

procedure TFVerifCpte.SqLJal ;
BEGIN
If MAJ then
   BEGIN
   If TousJalOk then Exit ;
   QCJ.Close ; QCJ.SQL.Clear ;
   QCJ.SQL.Add('Select Distinct E_JOURNAL From ECRITURE order by E_JOURNAL') ;
   ChangeSQL(QCJ) ; QCJ.Open ; CbEcr.Items.Clear ;
   While Not QCJ.Eof do
      BEGIN
      CbEcr.Items.Add(QCJ.Fields[0].AsString) ;
      QCJ.Next ;
      END ;
   QCJ.Close ; QCJ.SQL.Clear ;
   QCJ.SQL.Add('Select Distinct Y_JOURNAL From ANALYTIQ order by Y_JOURNAL') ;
   ChangeSQL(QCJ) ; QCJ.Open ; CbEcr.Values.Clear ;
   While Not QCJ.Eof do
      BEGIN
      CbEcr.Values.Add(QCJ.Fields[0].AsString) ;
      QCJ.Next ;
      END ;
   QCJ.Close ; QCJ.SQL.Clear ; 
   QCJ.SQL.Add('Select * from JOURNAL order by J_JOURNAL') ; QCJ.RequestLive:=MAJ ;
   ChangeSQL(QCJ) ; QCJ.Open ; InitMove(RecordsCount(QCJ),MsgBar.Mess[0]) ;
   END ;
Fillchar(Cp,SizeOf(Cp),#0) ; QCJ.First ;
While Not QCJ.Eof do
   BEGIN
   Cp.Code:=QCJ.FindField('J_JOURNAL').AsString ; Cp.Libelle:=QCJ.FindField('J_LIBELLE').AsString ;
   Cp.Axe:=QCJ.FindField('J_AXE').AsString ; Cp.Pref:='J' ;
   MoveCur(False) ;
   if MAJ then if CompteOk then begin QCJ.Next ; continue ; end ;
   if MAJ then EnMvt:=(CbEcr.Items.IndexOf(Cp.Code)>=0) Or (CbEcr.Values.IndexOf(Cp.Code)>=0) ;
   if VerifCode then
      BEGIN
      If Not VerifLibel then begin OkVerif:=False ; TousJalOk:=False ; end ;
      if VerifNatJAl then
         BEGIN
         IF Not VerifFacturier then begin OkVerif:=False ; TousJalOk:=False ; end ;
         IF Not VerifContrePartie then begin OkVerif:=False ; TousJalOk:=False ; end ;
         IF Not VerifRegul then begin OkVerif:=False ; TousJalOk:=False ; end ;
         If Not VerifBool(QCJ) then begin OkVerif:=False ; TousJalOk:=False ; end ;
         END Else begin OkVerif:=False ;TousJalOk:=False ; end ;
      END Else begin OkVerif:=False ; TousJalOk:=False ; end ;
   If Not MAJ then IF Not TousJalOk then FlagueCompte ;
   TuFaisQuoi(Cp) ; if TestBreak then Break ;
   QCJ.Next ;
   END ;
END ;

Function TFVerifCpte.VerifCode : Boolean ;
BEGIN
Result:=True ;
if Cp.Pref='G' then
   BEGIN
   If Cp.Code='' then
      BEGIN
      LaListe.Add(GoListe1(Cp,'',0));
      Result:=False ;
      END Else
      BEGIN
      if VH^.Cpta[fbgene].lg<>Length(Cp.Code) Then
         BEGIN
         LaListe.Add(GoListe1(Cp,IntToStr(Length(Cp.Code)),1));
         Result:=False ;
         END ;
      END ;
  END Else
if Cp.Pref='T' then
   BEGIN
   If Cp.Code='' then
      BEGIN
      LaListe.Add(GoListe1(Cp,'',0));
      Result:=False ;
      END Else
      BEGIN
      if VH^.Cpta[fbAux].lg<>Length(Cp.Code) Then
         BEGIN
         LaListe.Add(GoListe1(Cp,IntToStr(Length(Cp.Code)),1));
         Result:=False ;
         END ;
      END ;
   END Else
if Cp.Pref='S' then
   BEGIN
   IF Cp.Axe='' then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',2)); Result:=False ; end
                 Else begin Corrige('S_AXE','A1') ; end ;
      END Else
      BEGIN
      If ((Cp.Axe<>'A1')and(Cp.Axe<>'A2')and(Cp.Axe<>'A3')and(Cp.Axe<>'A4')and(Cp.Axe<>'A5'))  then
         BEGIN
         IF Not MAJ then begin LaListe.Add(GoListe1(Cp,Cp.Axe,30)); Result:=False ; end
                    Else begin Corrige('S_AXE','A1') ; end ;
         END Else
         If Cp.Code='' then
            BEGIN
            LaListe.Add(GoListe1(Cp,'',0)); Result:=False ;
            END Else
            BEGIN
            if VH^.Cpta[AxeToFb(Cp.Axe)].lg<>Length(Cp.Code) Then
               BEGIN
               LaListe.Add(GoListe1(Cp,IntToStr(Length(Cp.Code)),3));
               Result:=False ;
               END ;
            END ;
      END ;
   END Else
if Cp.Pref='J' then
   BEGIN
   If Cp.Code='' then
      BEGIN
      LaListe.Add(GoListe1(Cp,'',0));
      Result:=False ;
      END Else
      BEGIN
      if Length(Cp.Code)>3 Then
         BEGIN
         LaListe.Add(GoListe1(Cp,IntToStr(Length(Cp.Code)),1));
         Result:=False ;
         END ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifASCII : Boolean ;
Var i : Byte ; Trouve : Boolean ;
BEGIN
Trouve:=False ;
for i:=1 to Length(Cp.Code) do
    BEGIN
    if Not (Cp.Code[i] in ['0'..'9','A'..'z',VH^.Cpta[fbGene].Cb,' ','.','/',':','-']) then begin trouve:=True ; Break ; End ;
    END ;
if Trouve then
   BEGIN
   if Not MAJ then LaListe.Add(GoListe1(Cp,Cp.Code[i],9))
              Else begin if Not EnMvt then begin CorrigeASCII ; Trouve:=False ; end else Inc(Manuel) ; end ;
   END ;
Result:=not Trouve ;
END ;

Function TFVerifCpte.VerifLibel : Boolean ;
BEGIN
Result:=True ;
if Cp.Libelle='' then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',4)); Result:=False ; end
                 else begin Corrige(''+Cp.Pref+'_LIBELLE',Cp.Code) ; end ;
      END ;
END ;

Function TFVerifCpte.VerifNatureGene : Boolean ;
Var Na : String ;
    Let, Poin : Boolean ;
BEGIN
Result:=True ; Na:=QCG.FindField('G_NATUREGENE').AsString ;
if na='' then
   BEGIN
   if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',5)); result:=False ; exit ; end
              Else begin if Not EnMvt then Corrige('G_NATUREGENE','DIV') else Inc(Manuel); end ;
   END Else
if Not (TCnaGen.values.IndexOf(na)>-1) then
   BEGIN
   if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',6)); result:=False ; exit ; end
              Else begin if Not EnMvt then Corrige('G_NATUREGENE','DIV') else Inc(Manuel); end ;
   END ;
if (Na='BQE') or (Na='CAI') then
   BEGIN
   if QCG.FindField('G_COLLECTIF').AsString<>'-' then
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_COLLECTIF').FieldName+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('G_COLLECTIF','-') else Inc(Manuel); end ;
   if QCG.FindField('G_LETTRABLE').AsString<>'-' then
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_LETTRABLE').FieldName+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('G_LETTRABLE','-') else Inc(Manuel); end ;
   END Else
if (Na='CHA') or (Na='PRO') or (Na='IMO') then
   BEGIN
if vh^.CPIFDEFCEGID then
begin
   If na='IMO' Then
     BEGIN
     if QCG.FindField('G_COLLECTIF').AsString<>'-' then
        If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_COLLECTIF').FieldName+') ',7)) ; result:=False ; end
                   Else begin if Not EnMvt then Corrige('G_COLLECTIF','-') else Inc(Manuel); end ;
     END ;
end else
   if QCG.FindField('G_COLLECTIF').AsString<>'-' then
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_COLLECTIF').FieldName+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('G_COLLECTIF','-') else Inc(Manuel); end ;

   if QCG.FindField('G_LETTRABLE').AsString<>'-' then
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_LETTRABLE').FieldName+') ',7)) ; result:=False ; end
                 Else begin If Not EnMvt then Corrige('G_LETTRABLE','-') else Inc(Manuel); end ;
   if QCG.FindField('G_POINTABLE').AsString<>'-' then
      IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_POINTABLE').FieldName+') ',7)) ; result:=False ; end
                 Else begin If Not EnMvt then Corrige('G_POINTABLE','-') else Inc(Manuel); end ;
   END Else
if (Na='COC')or(Na='COD')or(Na='COF')or(Na='COS') then
   BEGIN
   if QCG.FindField('G_COLLECTIF').AsString<>'X' then
      IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_COLLECTIF').FieldName+') ',7)) ; result:=False ; end
                 Else begin If Not EnMvt then Corrige('G_COLLECTIF','X') else Inc(Manuel); end ;
   if QCG.FindField('G_LETTRABLE').AsString<>'-' then
      IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_LETTRABLE').FieldName+') ',7)) ; result:=False ; end
                 Else begin If Not EnMvt then Corrige('G_LETTRABLE','-') else Inc(Manuel); end ;
   END Else
if (Na='DIV') or (Na='EXT') then
   BEGIN
   if QCG.FindField('G_COLLECTIF').AsString<>'-' then
      IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_COLLECTIF').FieldName+') ',7)) ; result:=False ; end
                 Else begin If Not EnMvt then Corrige('G_COLLECTIF','-') else Inc(Manuel); end ;
   END Else
if (Na='TIC')or(Na='TID') then
   BEGIN
   if QCG.FindField('G_COLLECTIF').AsString<>'-' then
      IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_COLLECTIF').FieldName+') ',7)) ; result:=False ; end
                 Else begin If Not EnMvt then Corrige('G_COLLECTIF','-') else Inc(Manuel); end ;
   if QCG.FindField('G_POINTABLE').AsString<>'-' then
      IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatGene',Na,False)+') '+'('+QCG.FindField('G_POINTABLE').FieldName+') ',7)) ; result:=False ; end
                 Else begin if not EnMvt then Corrige('G_POINTABLE','-') else Inc(Manuel); end ;
{ GP : Tous les comptes peuvent être ventilables
   if QCG.FindField('G_VENTILABLE').AsString<>'-' then
      IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom(ttNatGene,Na,False)+') '+'('+QCG.FindField('G_VENTILABLE').FieldName+') ',7)) ; result:=False ; end
                 Else begin If Not EnMvt then Corrige('G_VENTILABLE','-') else Inc(Manuel); end ;
}
   END ;
Let:=(QCG.FindField('G_LETTRABLE').AsString='X') ;
//Ven:=(QCG.FindField('G_VENTILABLE').AsString='X') ;
Poin:=(QCG.FindField('G_POINTABLE').AsString='X') ;
{ GP : Tous les comptes peuvent être ventilables
If Let and (Poin or Ven) then
}
If Let and (Poin) then
   BEGIN
   If Not MAJ then
      BEGIN
      LaListe.Add(GoListe1(Cp,'',8)) ; Result:=False ;
      END Else
      BEGIN
      If Not EnMvt then begin if Poin then ResouCaracGene(1) else ResouCaracGene(2) ; end
                   Else begin Inc(Manuel) ; end ;
      END ;
   END ;

END ;

Function TFVerifCpte.VerifBool(Q : TQuery) : Boolean ;
Var na : String3 ; i : byte ; Trouve : Boolean ;
BEGIN
Result:=True;
If Not Bool(Q.FindField(''+Cp.Pref+'_FERME').AsString) then
   BEGIN
   if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_FERME'),'')); Result:=False ; end
              Else begin Corrige(''+Cp.Pref+'_FERME','-') ; end ;
   END ;
if (Cp.Pref='G')or(Cp.Pref='T') then
   BEGIN
   If Not Bool(Q.FindField(''+Cp.Pref+'_LETTRABLE').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_LETTRABLE'),''));Result:=False ; end
                 Else begin Corrige(''+Cp.Pref+'_LETTRABLE','-') ; end ;
      END ;
   If Not Bool(Q.FindField(''+Cp.Pref+'_SOUMISTPF').AsString) then
      BEGIN
      if not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_SOUMISTPF'),''));Result:=False ; end
                 Else begin Corrige(''+Cp.Pref+'_SOUMISTPF','-') ; end ;
      END ;
   END ;
if (Cp.Pref='G')or(Cp.Pref='J') then
   BEGIN
   If Not Bool(Q.FindField(''+Cp.Pref+'_CENTRALISABLE').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_CENTRALISABLE'),''));Result:=False ;end
                 Else begin Corrige(''+Cp.Pref+'_CENTRALISABLE','-') ; end ;
      END ;
   END ;
if (Cp.Pref='J')or(Cp.Pref='T') then
   BEGIN
   If Not Bool(Q.FindField(''+Cp.Pref+'_MULTIDEVISE').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_MULTIDEVISE'),''));Result:=False ;end
                 Else Corrige(''+Cp.Pref+'_MULTIDEVISE','-') ;
                 //Else begin if Cp.Pref='T' then Corrige(''+Cp.Pref+'_MULTIDEVISE','-') ; end ;
      END ;
   END ;
If Cp.Pref='G' then
   BEGIN
   Na:=Q.FindField('G_NATUREGENE').AsString ;
   //TaxePourNature:=(Na='IMO')Or(Na='CHA')Or(Na='PRO') :
   If Not Bool(Q.FindField('G_VENTILABLE').AsString) then
      BEGIN
      if not MAJ then
         BEGIN
         LaListe.Add(GoListe2(Cp,Q.FindField('G_VENTILABLE'),''));Result:=False ;
         END Else
         BEGIN
         Trouve:=False ;
         for i:=1 to 5 do if QCG.FindField('G_VENTILABLE'+IntToStr(i)+'').AsString='X' then begin Trouve:=True ; Break end ;
         if trouve then Corrige('G_VENTILABLE','X') else Corrige('G_VENTILABLE','-') ;
         END ;
      END ;
   If Not Bool(Q.FindField('G_POINTABLE').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField('G_POINTABLE'),'')); Result:=False ; end ;
      END ;
   If Not Bool(Q.FindField('G_COLLECTIF').AsString) then
      BEGIN
      if not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField('G_COLLECTIF'),'')); Result:=False ; end ;
      END ;
   If Not Bool(Q.FindField('G_PURGEABLE').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField('G_PURGEABLE'),''));Result:=False ; end
                 Else begin Corrige('G_PURGEABLE','X') ; end ;
      END ;
   If Not Bool(Q.FindField('G_RISQUETIERS').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField('G_RISQUETIERS'),''));Result:=False ;end
                 Else begin Corrige('G_RISQUETIERS','-') ; end ;
      END ;
   If Not Bool(Q.FindField('G_TVASURENCAISS').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField('G_TVASURENCAISS'),'')); Result:=False ; end
                 Else begin Corrige('G_TVASURENCAISS','-') ; end ;
      END ;
   END ;
If Cp.Pref='T' then
   BEGIN
   If Not Bool(Q.FindField('T_FACTUREHT').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField('T_FACTUREHT'),'')); Result:=False ; end
                 Else begin Corrige('T_FACTUREHT','-') ; end ;
      END ;
   If Not Bool(Q.FindField('T_RELEVEFACTURE').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField('T_RELEVEFACTURE'),'')); Result:=False ; end
                 Else begin if (Q.FindField('T_LETTRABLE').AsString='X') then Corrige('T_RELEVEFACTURE','-') else Corrige('T_RELEVEFACTURE','X') ; end ;
      END ;
   END ;
if Cp.Pref<>'J' then
   BEGIN
   If Q.FindField(''+Cp.Pref+'_CONFIDENTIEL').AsString<>'0' then
      if Q.FindField(''+Cp.Pref+'_CONFIDENTIEL').AsString<>'1' then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_CONFIDENTIEL'),'')); Result:=False ; end
                    Else begin Corrige(''+Cp.Pref+'_CONFIDENTIEL','0') ; end ;
         END ;
   If Not Bool(Q.FindField(''+Cp.Pref+'_SOLDEPROGRESSIF').AsString) then
      BEGIN
      IF Not MAJ then
         BEGIN
         LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_SOLDEPROGRESSIF'),''));
         Result:=False ;
         END Else begin Corrige(''+Cp.Pref+'_SOLDEPROGRESSIF','X') ; end ;
      END ;
   If Not Bool(Q.FindField(''+Cp.Pref+'_SAUTPAGE').AsString) then
      BEGIN
      If Not MAJ then
         BEGIN
         LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_SAUTPAGE'),''));
         Result:=False ;
         END Else begin Corrige(''+Cp.Pref+'_SAUTPAGE','-') ; end ;
      END ;
   If Not Bool(Q.FindField(''+Cp.Pref+'_TOTAUXMENSUELS').AsString) then
      BEGIN
      If Not MAJ then
         BEGIN
         LaListe.Add(GoListe2(Cp,Q.FindField(''+Cp.Pref+'_TOTAUXMENSUELS'),''));
         Result:=False ;
         END Else begin Corrige(''+Cp.Pref+'_TOTAUXMENSUELS','-') ; end ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifCorresp(Q : TQuery) : Boolean ;
Var Code1, Code2 : String3 ;
BEGIN
Result:=True ;
if Cp.Pref='G' then begin code1:='GE1' ; code2:='GE2' ; end else
if Cp.Pref='T' then begin code1:='AU1' ; code2:='AU2' ; end else
if Cp.Pref='S' then begin code1:=Cp.Axe+'1' ; code2:=Cp.Axe+'2' ; end ;
IF Q.FindField(''+Cp.Pref+'_CORRESP1').AsString<>'' then
   BEGIN
   If Not OKCorresp(Code1,Q.FindField(''+Cp.Pref+'_CORRESP1').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,Q.FindField(''+Cp.Pref+'_CORRESP1').AsString,11)); Result:=False ; end
                 Else begin Corrige(''+Cp.Pref+'_CORRESP1','') ; end ;
      END ;
   END ;
IF Q.FindField(''+Cp.Pref+'_CORRESP2').AsString<>'' then
   BEGIN
   If Not OKCorresp(Code2,Q.FindField(''+Cp.Pref+'_CORRESP2').AsString) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,Q.FindField(''+Cp.Pref+'_CORRESP2').AsString,12)); Result:=False ; end
                 Else begin Corrige(''+Cp.Pref+'_CORRESP2','') ; end ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifQualifQte : Boolean ;
BEGIN
Result:=True ;
if (QCG.FindField('G_QUALIFQTE1').AsString<>'') then
   BEGIN
{FQ 19936 BVE 10.04.07
   IF Not PresenceComplexe('COMMUN',['CO_TYPE','CO_CODE'],['=','='],['QME',QCG.FindField('G_QUALIFQTE1').AsString],['S','S']) then
}
   IF Not PresenceComplexe('CHOIXCOD',['CC_TYPE','CC_CODE'],['=','='],['QME',QCG.FindField('G_QUALIFQTE1').AsString],['S','S']) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCG.FindField('G_QUALIFQTE1').AsString,23)); Result:=False ; end
                 Else Corrige('G_QUALIFQTE1','') ;
      END ;
   END ;
if (QCG.FindField('G_QUALIFQTE2').AsString<>'') then
   BEGIN
{FQ 19936 BVE 10.04.07
   IF Not PresenceComplexe('COMMUN',['CO_TYPE','CO_CODE'],['=','='],['QME',QCG.FindField('G_QUALIFQTE2').AsString],['S','S']) then
}
   IF Not PresenceComplexe('CHOIXCOD',['CC_TYPE','CC_CODE'],['=','='],['QME',QCG.FindField('G_QUALIFQTE2').AsString],['S','S']) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCG.FindField('G_QUALIFQTE2').AsString,24)); Result:=False ; end
                 else Corrige('G_QUALIFQTE2','') ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifTaxe : Boolean ;
Var TaxePourNature : Boolean ;
    Na : String3 ;
BEGIN
Result:=True ; Na:=QCG.FindField('G_NATUREGENE').AsString ;
TaxePourNature:=(Na='IMO')Or(Na='CHA')Or(Na='PRO') ;

IF Not TaxePourNature then
   BEGIN
   If QCG.FindField('G_TVA').AsString<>'' then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',13)); Result:=False ; end
                 Else begin Corrige('G_TVA','' ) ; End ;
      END ;
   If QCG.FindField('G_TVASURENCAISS').AsString='X' then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',14)); end
                 Else begin Corrige('G_TVASURENCAISS','-' ) ; end ;
      END ;
   If QCG.FindField('G_TPF').AsString<>'' then
      BEGIN
      if Not Maj then begin LaListe.Add(GoListe1(Cp,'',15)); Result:=False ; end
                 Else begin Corrige('G_TPF','' ) ; end ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifTicTid : Boolean ;
Var TicTid, Let : Boolean ;
    Na : String3 ; toto : String ;
BEGIN
Result:=True ; Na:=QCG.FindField('G_NATUREGENE').AsString ;
TicTid:=(Na='TIC')Or(Na='TID') ; Let:=(QCG.FindField('G_LETTRABLE').AsString='X') ;
IF TicTid then
   BEGIN
   if QCG.FindField('G_REGIMETVA').AsString<>'' then
      BEGIN
      if Not ExisteCode(3,QCG.FindField('G_REGIMETVA').AsString,Toto) then
         BEGIN
         IF Not MAJ then begin LaListe.Add(GoListe1(Cp,QCG.FindField('G_REGIMETVA').AsString,17)); Result:=False ; end
                    Else begin if Not Let then Corrige('G_REGIMETVA','') else Corrige('G_REGIMETVA',TCRegTVA.Values[0]) ; end ;
         END ;
      END Else If Let then
                  BEGIN
                  If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',37)); Result:=False ; end
                             Else begin Corrige('G_REGIMETVA',TCRegTVA.Values[0]) ; end ;
                  END ;
  if QCG.FindField('G_TVAENCAISSEMENT').AsString='' then
      BEGIN
      If Na='TIC' then
         BEGIN
         IF Not MAJ then begin LaListe.Add(GoListe1(Cp,'',18)); Result:=False ; end
                    Else begin Corrige('G_TVAENCAISSEMENT',TCTvaEnc.Values[0]) ; end ;
         END ;
      END Else
      BEGIN
      if Not ExisteCode(4,QCG.FindField('G_TVAENCAISSEMENT').AsString,toto) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCG.FindField('G_TVAENCAISSEMENT').AsString,19)); Result:=False ; end
                    Else begin Corrige('G_TVAENCAISSEMENT',TCTvaEnc.Values[0]) ; end ;
         END ;
      END ;
   if QCG.FindField('G_MODEREGLE').AsString='' then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',20)); Result:=False ; end
                 Else begin Corrige('G_MODEREGLE',TCModR.Values[0]) ; end ;
      END Else
      BEGIN
      if Not ExisteCode(2,QCG.FindField('G_MODEREGLE').AsString,toto) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCG.FindField('G_MODEREGLE').AsString,21)); Result:=False ; end
                    Else begin Corrige('G_MODEREGLE',TCModR.Values[0]) ; end ;
         END ;
      END ;
   If QCG.FindField('G_RELANCEREGLEMENT').AsString<>'' then
      if Not (TCRelRegle.values.IndexOf(QCG.FindField('G_RELANCEREGLEMENT').AsString)>-1) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCG.FindField('G_RELANCEREGLEMENT').AsString,31)); Result:=False ; end
                    Else begin Corrige('G_RELANCEREGLEMENT','') ; end ;
         END ;
   If QCG.FindField('G_RELANCETRAITE').AsString<>'' then
      if Not (TCRelTraite.values.IndexOf(QCG.FindField('G_RELANCETRAITE').AsString)>-1) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCG.FindField('G_RELANCETRAITE').AsString,32)); Result:=False ; end
                    Else begin Corrige('G_RELANCETRAITE','') ; end ;
         END ;
   END ;
END ;

Function TFVerifCpte.VerifBanque : Boolean ;
Var Q : TQuery ; Devise : String3 ;
BEGIN
Result:=True ;
If QCG.FindField('G_NATUREGENE').AsString<>'BQE' then exit ;
Q:=OpenSql('Select BQ_DEVISE from BANQUECP where BQ_GENERAL="'+Cp.Code
          +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ',true) ; // 19/10/2006 YMO Multisociétés
if not Q.Eof then
   BEGIN
   if Q.FindField('BQ_DEVISE').AsString='' then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',34)); Result:=False ; end
                 Else begin if Not EnMvt then MAJBQEDEVISE(Cp.Code) else Inc(Manuel); end ;
      END Else
      BEGIN
      Devise:=Q.FindField('BQ_DEVISE').AsString ;
      if Not (TCDev.values.IndexOf(Devise)>-1) then
         BEGIN
         If Not MAJ then begin LaListe.Add(GoListe1(Cp,RechDom('ttDevise',Devise,False),35)); Result:=False ; end
                    Else begin if Not EnMvt then MAJBQEDEVISE(Cp.Code) else Inc(Manuel); end ;
         END ;
      END ;
   END ;
Ferme(Q) ;
END;

Function TFVerifCpte.VerifDernLet(Q : TQuery) : Boolean ;
BEGIN
Result:=True ;
IF Q.FindField(''+Cp.Pref+'_LETTRABLE').AsString='X' then exit ;
if Q.FindField(''+Cp.Pref+'_DERNLETTRAGE').AsString<>'' then
   BEGIN
   if Not MAJ then begin LaListe.Add(GoListe1(Cp,Q.FindField(''+Cp.Pref+'_DERNLETTRAGE').AsString,36)); Result:=False ; end
              Else begin Corrige(''+Cp.Pref+'_DERNLETTRAGE','') ; end ;
   END ;
END ;

procedure TFVerifCpte.SqlTiers ;
BEGIN
If MAJ then
   BEGIN
   If TousAuxOk then Exit ;
   QCT.Close ; QCT.SQL.Clear ;
   QCT.SQL.Add('Select Distinct E_AUXILIAIRE From ECRITURE order by E_AUXILIAIRE') ;
   ChangeSQL(QCT) ; QCT.Open ; CbEcr.Items.Clear ;
   While Not QCT.Eof do
      BEGIN
      CbEcr.Items.Add(QCT.Fields[0].AsString) ;
      QCT.Next ;
      END ;
   QCT.Close ; QCT.SQL.Clear ; 
   QCT.SQL.Add('Select * from TIERS order by T_AUXILIAIRE') ; QCT.RequestLive:=MAJ ;
   ChangeSQL(QCT) ; QCT.Open ; InitMove(RecordsCount(QCT),MsgBar.Mess[0]) ;
   END ;
Fillchar(Cp,SizeOf(Cp),#0) ; QCT.First ;
While Not QCT.Eof do
   BEGIN
   Cp.Code:=QCT.FindField('T_AUXILIAIRE').AsString ; Cp.Libelle:=QCT.FindField('T_LIBELLE').AsString ;
   Cp.Pref:='T' ;
   MoveCur(False) ;
   if MAJ then if CompteOk then begin QCT.Next ; continue ; end ;
   if MAJ then EnMvt:=(CbEcr.Items.IndexOf(Cp.Code)>=0) ;
   if VerifCode then
      BEGIN
      If Not VerifLibel then begin OkVerif:=False ; TousAuxOk:=False ; end ;
      if VerifGeneCol then
         BEGIN
         if VerifNatureAuxi then
            BEGIN
            IF Not VerifCorresp(QCT) then begin OkVerif:=False ; TousAuxOk:=False ; end ;
            IF Not VerifDernLet(QCT) then begin OkVerif:=False ; TousAuxOk:=False ; end ;
            if Not VerifTVA then begin OkVerif:=False ; TousAuxOk:=False ; end ;
            if Not VerifReglenemt then begin OkVerif:=False ; TousAuxOk:=False ; end ;
            If Not VerifBool(QCT) then begin OkVerif:=False ; TousAuxOk:=False ; end ;
            END Else begin OkVerif:=False ; TousAuxOk:=False ; end ;
         END Else begin OkVerif:=False ; TousAuxOk:=False ; end ;
      END Else begin OkVerif:=False ; TousAuxOk:=False ; end ;
   If Not MAJ then IF Not TousAuxOk then FlagueCompte ;
   TuFaisQuoi(Cp) ; if TestBreak then Break ;
   QCT.Next ;
   END ;
END ;

Function TFVerifCpte.VerifGeneCol : Boolean ;
BEGIN
Result:=True ;
END ;

Function TFVerifCpte.VerifNatureAuxi : Boolean ;
Var NatAuxi, NatGene : String ;
BEGIN
Result:=True ; NatAuxi:=QCT.FindField('T_NATUREAUXI').AsString ;
if NatAuxi='' then
   BEGIN
   If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',5)); result:=False ; exit ; end
              Else begin If Not EnMvt then Corrige('T_NATUREAUXI','DIV') else inc(Manuel); end ;
   END Else
if Not (TCnaAux.values.IndexOf(NatAuxi)>-1) then
   BEGIN
   If Not MAJ then begin LaListe.Add(GoListe1(Cp,NatAuxi,6)); result:=False ; exit ; end
              Else begin if Not EnMvt then Corrige('T_NATUREAUXI','DIV') else Inc(Manuel) ; end ;
   END ;
if QCT.FindField('T_COLLECTIF').AsString='' then
   BEGIN
   If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',25)); result:=False ; Exit ; end
              else begin if Not EnMvt then DonneCollectif else inc(Manuel); end ;
   END Else
   if Not ExisteCode(1,QCT.FindField('T_COLLECTIF').AsString,NatGene) then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,QCT.FindField('T_COLLECTIF').AsString,26)); result:=False ; exit ; end
                 Else begin if Not EnMvt then DonneCollectif else inc(manuel) ; end ;
      END ;

if (NatAuxi='CLI') then
   BEGIN
   if NatGene<>'COC' then Result:=False ;
   END Else
if NatAuxi='FOU' then
   BEGIN
   if NatGene<>'COF' then Result:=False ;
   END Else
if NatAuxi='SAL' then
   BEGIN
   if NatGene<>'COS' then Result:=False ;
   END Else
if NatAuxi='AUC' then
   BEGIN
   If NatGene<>'COD' then Result:=False ;
   END Else
if NatAuxi='AUD' then
   BEGIN
   if NatGene<>'COD' then Result:=False ;
   END Else
if NatAuxi='DIV' then
   BEGIN
   if NatGene<>'COD' then Result:=False ;
   END ;
If Not Result then If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatTiersCPTA',NatAuxi,False)+')'+' ('+RechDom('ttNatGene',NatGene,False)+') ',27)) ; end
                              else begin If Not EnMvt then begin DonneCollectif ; Result:=True ; end else inc(Manuel) ; end ;
END ;

Function TFVerifCpte.VerifTVA : Boolean ;
Var Na : String3 ; toto : string ;
BEGIN
Result:=True ; Na:=QCT.FindField('T_NATUREAUXI').AsString ;
if QCT.FindField('T_REGIMETVA').AsString='' then
   BEGIN
   If na<>'SAL' Then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',38)); Result:=False ; end
                 Else begin Corrige('T_REGIMETVA',TCRegTVA.Values[0]) ; end ;
      END ;
   END Else
   if Not ExisteCode(3,QCT.FindField('T_REGIMETVA').AsString,toto) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCT.FindField('T_REGIMETVA').AsString,39)); Result:=False ; end
                 Else begin Corrige('T_REGIMETVA',TCRegTVA.Values[0]) ; end ;
      END ;
If VH^.OuiTvaEnc Then
   BEGIN
   if QCT.FindField('T_TVAENCAISSEMENT').AsString='' then
      BEGIN
      if ((Na='FOU')or(Na='AUC')) then
         BEGIN
         If Not MAJ then begin (* fiche 5069 LaListe.Add(GoListe1(Cp,'',40)); Result:=False ;*) end
                    Else begin Corrige('T_TVAENCAISSEMENT',TCTvaEnc.Values[0]) ; end ;
         END ;
      END Else
      BEGIN
      if Not ExisteCode(4,QCT.FindField('T_TVAENCAISSEMENT').AsString,toto) then
         BEGIN
         If Not MAJ then begin LaListe.Add(GoListe1(Cp,QCT.FindField('T_TVAENCAISSEMENT').AsString,41)); Result:=False ; end
                    Else begin Corrige('T_TVAENCAISSEMENT',TCTvaEnc.Values[0]) ; end ;
         END ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifReglenemt : Boolean ;
Var Toto : String ;
BEGIN
Result:=True ;
If QCT.FindField('T_RELANCEREGLEMENT').AsString<>'' then
   if Not (TCRelRegle.values.IndexOf(QCT.FindField('T_RELANCEREGLEMENT').AsString)>-1) then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,QCT.FindField('T_RELANCEREGLEMENT').AsString,42)); Result:=False ; end
                 Else begin Corrige('T_RELANCEREGLEMENT','') ; end ;
      END ;
If QCT.FindField('T_RELANCETRAITE').AsString<>'' then
   if Not (TCRelTraite.values.IndexOf(QCT.FindField('T_RELANCETRAITE').AsString)>-1) then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,QCT.FindField('T_RELANCETRAITE').AsString,43)); Result:=False ; end
                 ELse begin Corrige('T_RELANCETRAITE','') end ;
      END ;
IF QCT.FindField('T_MODEREGLE').AsString='' then
   BEGIN
   if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',28)); Result:=False ; end
              Else begin Corrige('T_MODEREGLE',TCModR.Values[0]) ; end ;
   END Else
   if Not ExisteCode(2,QCT.FindField('T_MODEREGLE').AsString,toto) then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,QCT.FindField('T_MODEREGLE').AsString,29)); Result:=False ; end
                 Else begin Corrige('T_MODEREGLE',TCModR.Values[0]) ; end ;
      END ;
If QCT.FindField('T_LETTRABLE').AsString='X' then
   if QCT.FindField('T_RELEVEFACTURE').AsString='X' then
      if QCT.FindField('T_NATUREAUXI').AsString='CLI'then Exit else //Lhe le 30/04/97 n°1580
         if QCT.FindField('T_NATUREAUXI').AsString='AUD' then Exit else
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',45)); Result:=False ; end
                    Else begin Corrige('T_RELEVEFACTURE','-') ; end ;
         END ;
END ;

Function TFVerifCpte.VerifNatJal : Boolean ;
Var Na : String3 ; LibNa : String ;
BEGIN
Result:=True ;
Na:=QCJ.FindField('J_NATUREJAL').AsString ;
if Na='' then
   BEGIN
   if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',5));result:=False ; exit ; end
              Else begin if not EnMvt then Corrige('J_NATUREJAL','OD') else inc(Manuel); end ;
   END else
if Not (TCnaJal.values.IndexOf(na)>-1) then
   BEGIN
   if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+Na+') ',6)); result:=False ; exit ; end
              Else begin if not EnMvt then Corrige('J_NATUREJAL','OD') else inc(Manuel); end ;
   END ;
LibNa:=RechDom('ttNatJal',Na,False) ; // A mettre dans les conditions (vitesse)

if(Na<>'REG')then
  BEGIN
  if QCJ.FindField('J_CPTEREGULDEBIT1').AsString<>'' then
     if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CPTEREGULDEBIT1').AsString+') ',7)) ; result:=False ; end
                Else begin Corrige('J_CPTEREGULDEBIT1','') ; end ;
  if QCJ.FindField('J_CPTEREGULCREDIT1').AsString<>'' then
     if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CPTEREGULCREDIT1').AsString+') ',7)) ; result:=False ; end
                Else begin Corrige('J_CPTEREGULCREDIT1','') ; end ;
  if QCJ.FindField('J_CPTEREGULDEBIT2').AsString<>'' then
     If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CPTEREGULDEBIT2').AsString+') ',7)) ; result:=False ; end
                Else begin Corrige('J_CPTEREGULDEBIT2','') ; end ;
  if QCJ.FindField('J_CPTEREGULCREDIT2').AsString<>'' then
     If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CPTEREGULCREDIT2').AsString+') ',7)) ; result:=False ; end
                Else begin Corrige('J_CPTEREGULCREDIT2','') ; end ;
  if QCJ.FindField('J_CPTEREGULDEBIT3').AsString<>'' then
     If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CPTEREGULDEBIT3').AsString+') ',7)) ; result:=False ; end
                Else begin Corrige('J_CPTEREGULDEBIT3','') ; end ;
  if QCJ.FindField('J_CPTEREGULCREDIT3').AsString<>'' then
     If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CPTEREGULCREDIT3').AsString+') ',7)) ; result:=False ; end
                Else begin Corrige('J_CPTEREGULCREDIT3','') ; end ;
  END Else
if (Na='BQE') or (Na='CAI') then
   BEGIN
   if QCJ.FindField('J_AXE').AsString<>'' then
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_AXE').AsString+') ',7)) ; result:=False ; end
                 Else begin Corrige('J_AXE','') ; end ;
   if QCJ.FindField('J_COMPTEAUTOMAT').AsString<>'' then
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_COMPTEAUTOMAT').AsString+') ',7)) ; result:=False ; end
                 Else begin Corrige('J_COMPTEAUTOMAT','') ; end ;
   END Else
if (Na='ODA')or(Na='ANA') then
   BEGIN
   If QCJ.FindField('J_AXE').AsString='' then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_AXE').AsString+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('J_AXE','A1') else Inc(Manuel); end ;
      END Else
      BEGIN
      If (QCJ.FindField('J_AXE').AsString<>'A1')or(QCJ.FindField('J_AXE').AsString<>'A2')or
         (QCJ.FindField('J_AXE').AsString<>'A3')or(QCJ.FindField('J_AXE').AsString<>'A4')
         or (QCJ.FindField('J_AXE').AsString<>'A5') then
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_AXE').AsString+') ',7)) ; result:=False ; end
                    Else begin If Not EnMvt then Corrige('J_AXE','A1') else Inc(Manuel); end;
      END ;
   If QCJ.FindField('J_MULTIDEVISE').AsString='X' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_MULTIDEVISE').FieldName+') ',7)) ; result:=False ; end
                 Else begin Corrige('J_MULTIDEVISE','-') ; end ;
   If QCJ.FindField('J_COMPTEINTERDIT').AsString<>'' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_COMPTEINTERDIT').AsString+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('J_COMPTEINTERDIT','') else Inc(Manuel); end;
   If QCJ.FindField('J_COMPTEAUTOMAT').AsString<>'' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_COMPTEAUTOMAT').AsString+') ',7)) ; result:=False ; end
                 Else begin Corrige('J_COMPTEAUTOMAT','') ; end ;
   If QCJ.FindField('J_COMPTEURSIMUL').AsString<>'' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_COMPTEURSIMUL').AsString+') ',7)) ; result:=False ; end
                 Else begin Corrige('J_COMPTEURSIMUL','') ; end ;
   If QCJ.FindField('J_CONTREPARTIE').AsString<>'' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CONTREPARTIE').AsString+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('J_CONTREPARTIE','') else inc(Manuel); end;
   If QCJ.FindField('J_TYPECONTREPARTIE').AsString<>'' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_TYPECONTREPARTIE').AsString+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('J_TYPECONTREPARTIE','') else inc(Manuel); end ;
   END Else
if (Na='ACH')Or(Na='ECC')Or(Na='EXT')
   Or(Na='OD')Or(Na='REG')Or(Na='VTE') then
   BEGIN
   if QCJ.FindField('J_AXE').AsString<>'' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_AXE').AsString+') ',7)) ; result:=False ; end
                 Else begin Corrige('J_AXE','') ; end ;
   If QCJ.FindField('J_CONTREPARTIE').AsString<>'' then
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_CONTREPARTIE').AsString+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('J_CONTREPARTIE','') else inc(Manuel); end ;
   If QCJ.FindField('J_TYPECONTREPARTIE').AsString<>'' then
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+LibNa+') '+'('+QCJ.FindField('J_TYPECONTREPARTIE').AsString+') ',7)) ; result:=False ; end
                 Else begin if Not EnMvt then Corrige('J_TYPECONTREPARTIE','') else inc(Manuel); end ;
   END ;

END ;

Function TFVerifCpte.VerifFacturier : Boolean ;
BEGIN
Result:=True ;
if QCJ.FindField('J_COMPTEURNORMAL').AsString<>'' then
   BEGIN
   If (QCJ.FindField('J_NATUREJAL').AsString='ODA')or(QCJ.FindField('J_NATUREJAL').AsString='ANA') then
      BEGIN
      if Not PresenceComplexe('SOUCHE',['SH_SOUCHE','SH_TYPE','SH_SIMULATION','SH_ANALYTIQUE'],['=','=','=','='],[QCJ.FindField('J_COMPTEURNORMAL').AsString,'CPT','-','X'],['S','S','S','S']) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttSoucheComptaODA',QCJ.FindField('J_COMPTEURNORMAL').AsString,False)+') ',46)); result:=False ;end
                    Else begin Corrige('J_COMPTEURNORMAL','') ; end ;
         END ;
      END Else
      BEGIN
      if Not PresenceComplexe('SOUCHE',['SH_SOUCHE','SH_TYPE','SH_SIMULATION','SH_ANALYTIQUE'],['=','=','=','='],[QCJ.FindField('J_COMPTEURNORMAL').AsString,'CPT','-','-'],['S','S','S','S']) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttSoucheCompta',QCJ.FindField('J_COMPTEURNORMAL').AsString,False)+') ',46)); result:=False ; end
                    Else begin Corrige('J_COMPTEURNORMAL','') ; end ;
         END ;
      END ;
   END ;
if QCJ.FindField('J_COMPTEURSIMUL').AsString<>'' then
   BEGIN
   if Not PresenceComplexe('SOUCHE',['SH_SOUCHE','SH_TYPE','SH_SIMULATION'],['=','=','='],[QCJ.FindField('J_COMPTEURSIMUL').AsString,'CPT','X'],['S','S','S','S']) then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttSoucheComptaSimul',QCJ.FindField('J_COMPTEURSIMUL').AsString,False)+') ',47)); result:=False ;end
                 Else begin Corrige('J_COMPTEURSIMUL','') ; end ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifContrePartie : Boolean ;
Var Na : String3 ; ContPartie, TypeContPartie, NatGene, Toto : String ;
   JalEffet : Boolean;
BEGIN
Result:=True ;
Na:=QCJ.FindField('J_NATUREJAL').AsString ;
ContPartie:=QCJ.FindField('J_CONTREPARTIE').AsString ;
TypeContPartie:=QCJ.FindField('J_TYPECONTREPARTIE').AsString ;
JalEffet := (QCJ.FindField('J_EFFET').AsString = 'X') ;
if (Na<>'BQE')And(Na<>'CAI') then
   BEGIN
   if ContPartie<>'' then
      BEGIN
      if Not MAJ then begin if (not JalEffet) then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatJal',Na,False)+') '+' ('+ContPartie+') ',54)); result:=False ; end; end
                 Else begin if (not JalEffet) then Corrige('J_CONTREPARTIE','') ; end ;
      END ;
   if TypeContPartie<>'' then
      BEGIN
      if Not MAJ then begin if (not JalEffet) then begin LaListe.Add(GoListe1(Cp,' ('+RechDom('ttNatJal',Na,False)+') '+' ('+TypeContPartie+') ',55)); result:=False ; end; end
                 Else begin if (not JalEffet) then Corrige('J_TYPECONTREPARTIE','') ; end ;
      END ;
   Exit ;
   END ;
if Na='BQE' then
   BEGIN
   If ContPartie='' then
      BEGIN
      if NOT MAJ then begin LaListe.Add(GoListe1(Cp,'',48)); result:=False ; end
                 Else begin DonneContrePartie end ;
      END Else
      BEGIN
      if Not ExisteCode(6,ContPartie,NatGene) then
         BEGIN
         If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',49)); result:=False ; end
                    Else begin DonneContrePartie end ;
         END Else
         BEGIN
         If na<>NatGene then
            BEGIN
            If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+ContPartie+') '+' ('+RechDom('ttNatGene',NatGene, False)+') ',50)); result:=False ;end
                       Else begin DonneContrePartie end ;
            END ;
         END ;
      END ;
   If TypeContPartie='' then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',52)); result:=False ; end
                 Else begin Corrige('J_TYPECONTREPARTIE',TCTypeConPart.Values[0]) ; end ;
      END Else
      BEGIN
      if Not ExisteCode(5,TypeContPartie,Toto) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+TypeContPartie+') ',53)); result:=False ;end
                    Else begin Corrige('J_TYPECONTREPARTIE',TCTypeConPart.Values[0]) ; end ;
         END ;
      END ;
   END Else
if Na='CAI' then
   BEGIN
   If ContPartie='' then
      BEGIN
      If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',48)); result:=False ; end
                 Else begin DonneContrePartie ; end ;
      END Else
      BEGIN
      If Not ExisteCode(6,ContPartie,NatGene) then
         BEGIN
         If Not MAJ then begin LaListe.Add(GoListe1(Cp,'',49)); result:=False ; end
                    Else begin DonneContrePartie ; end ;
         END Else
         BEGIN
         if Na<>NatGene then
            BEGIN
            If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+ContPartie+') '+' ('+RechDom('ttNatGene',NatGene, False)+') ',51)); result:=False ; end
                       else begin DonneContrePartie ; end ;
            END ;
         END ;
      END ;
   If TypeContPartie='' then
      BEGIN
      if Not MAJ then begin LaListe.Add(GoListe1(Cp,'',52)); result:=False ; end
                 Else begin Corrige('J_TYPECONTREPARTIE',TCTypeConPart.Values[0]) ; end ;
      END Else
      BEGIN
      if Not ExisteCode(5,TypeContPartie,toto) then
         BEGIN
         if Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+TypeContPartie+') ',53)); result:=False ; end
                    Else begin Corrige('J_TYPECONTREPARTIE',TCTypeConPart.Values[0]) ; end ;
         END ;
      END ;
   END ;
END ;

Function TFVerifCpte.VerifRegul : Boolean ;
Var I : Byte ;
BEGIN
Result:=True ;
if QCJ.FindField('J_NATUREJAL').AsString<>'REG' then
   BEGIN
   For i:=1 to 3 do
       BEGIN
       if QCJ.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').AsString<>'' then
          If not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+QCJ.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').FieldName+') '+'('+QCJ.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').AsString+')',57)); result:=False ; end
                     Else begin Corrige('J_CPTEREGULDEBIT'+IntToStr(i)+'','') ; end ;
       if QCJ.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').AsString<>'' then
          IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+QCJ.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').FieldName+') '+'('+QCJ.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').AsString+')',57)); result:=False ; end
                     Else begin Corrige('J_CPTEREGULCREDIT'+IntToStr(i)+'','') ; end ;
       END ;
   END else
   BEGIN
   For i:=1 to 3 do
       BEGIN
       if QCJ.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').AsString<>'' then
          BEGIN
          IF Not Presence('GENERAUX','G_GENERAL',QCJ.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').AsString) then
             BEGIN
             If Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+QCJ.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').FieldName+') '+' ('+FloatToStr(QCJ.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').AsFloat)+') ',58)); result:=False ;end
                        Else begin Corrige('J_CPTEREGULDEBIT'+IntToStr(i)+'','') ; end ;
             END ;
          END ;
       if QCJ.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').AsString<>'' then
          BEGIN
          IF Not Presence('GENERAUX','G_GENERAL',QCJ.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').AsString) then
             BEGIN
             IF Not MAJ then begin LaListe.Add(GoListe1(Cp,' ('+QCJ.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').FieldName+') '+' ('+FloatToStr(QCJ.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').AsFloat)+') ',58)); result:=False ; end
                        Else begin Corrige('J_CPTEREGULCREDIT'+IntToStr(i)+'','') ; end ;
             END ;
          END ;
       END ;
   END ;
END ;

function VerCompte(Y : Byte ; EnBatch : boolean ) : boolean ;
var VCpte : TFVerifCpte ;
BEGIN
VCpte:=TFVerifCpte.Create(Application) ;
try
  VCpte.EnBatch:=EnBatch ;
 With VCpte do
      BEGIN
      FVerification.ItemIndex:=Y ;
      QuelleVerif:=200 ; ParLaForme:=False ;
      OkVerif:=True ; NbError:=0 ; MAJ:=False ;
      TousGenOk:=True ; TousAuxOk:=True ; TousSecOk:=True ; TousJalOk:=True ;
      LaListe.Clear ; StopVerif:=False ; Manuel:=0;
      If CompteATraiter then
         BEGIN
         ChargeInfos ;
         LanceVerif ;
         END ;
      If Not OkVerif then RapportdErreurMvt(Laliste,4,MAJ,EnBatch) ; //Else MAJ:=True ;
      Result:=OkVerif ;
{      if MAJ then
         BEGIN
         InitLabel ; StopVerif:=False ;
         LanceVerif ;
         NbPasPossible:=Manuel;
         END ;
}
      VideInfos ;
      END ;
 finally
 VCpte.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Procedure VerCompteMAJ(Y : Byte ; Var NbPasPossible : Integer ; EnControl : Boolean ) ;
var VCpte : TFVerifCpte ;
BEGIN
VCpte:=TFVerifCpte.Create(Application) ;
try
 With VCpte do
      BEGIN
      FVerification.ItemIndex:=Y ;
      QuelleVerif:=200 ; ParLaForme:=False ;
      OkVerif:=True ; NbError:=0 ; MAJ:=False ;
      TousGenOk:=True ; TousAuxOk:=True ; TousSecOk:=True ; TousJalOk:=True ;
      LaListe.Clear ; StopVerif:=False ; Manuel:=0;
      If CompteATraiter then
         BEGIN
         ChargeInfos ;
         LanceVerif ;
         END ;
      If Not OkVerif then
         if EnControl then RapportdErreurMvt(Laliste,4,MAJ,False) Else MAJ:=True ;
      if MAJ then
         BEGIN
         InitLabel ; StopVerif:=False ;
         LanceVerif ;
         NbPasPossible:=Manuel;
         END ;
      VideInfos ;
      END ;
 finally
 VCpte.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;


procedure TFVerifCpte.FormDestroy(Sender: TObject);
begin
 LaListe.Free ;
end;

end.
