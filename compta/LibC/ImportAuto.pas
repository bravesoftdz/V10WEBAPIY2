unit ImportAuto;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, ComCtrls, HStatus, DBTables, ImpUtil, FmtChoix, Ent1, HEnt1,
  HCtrls,DB,MajTable,TImpFic,LicUtil,ed_tools,VerCpta, StdCtrls, ImpFicU,RappType,
{$IFDEF AGL570D}
{$ELSE}
  PGIEnv,
{$ENDIF}
  PGIExec, HmsgBox
  ,EntPGI
 ;

type
  TFImpauto = class(TForm)
    Status: THStatusBar;
    Panel1: TPanel;
    Timer1: TTimer;
    MessImport: THLabel;
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    FMultiDossier : boolean;
    FFormatMulti : string;
    FScenarioMulti : string;
    FSociete : string;
    procedure LanceImportAuto ;
    procedure InitParametreMultiDossier;
  public
    { Déclarations publiques }
  end;

var
  FImpauto: TFImpauto;

implementation

uses SISCO, ImpAutoParamMulti;

{$R *.DFM}

procedure TFImpauto.FormShow(Sender: TObject);
begin
  VStatus:=Status ;
  Status.Caption := Copyright ;
end;

procedure init ;
BEGIN
V_PGI.Debug:=FALSE ;
V_PGI.Versiondev:=FALSE ;
V_PGI.Synap:=FALSE ;
VH^.GrpMontantMin:=0 ;
VH^.GrpMontantMax:=1000000 ;
V_PGI.DateEntree := Date ;
VH^.Mugler:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.NumVersion:='3' ;
V_PGI.SAV:=TRUE ;
V_PGI.PassWord:=CryptageSt(DayPass(Date)) ;
V_PGI.Versiondev:=FALSE ;
V_PGI.Synap:=FALSE ;
VH^.GereSousPlan:=True ;
V_PGI.Halley:=TRUE ;
END ;

Function RepositionneNomFic(Var InfoImp : TInfoImport) : Boolean ;
Var LePath,LeNom,Lextension : String ;
    i,j,k : Integer ;
    pbSISCO : Boolean ;
BEGIN
Result:=TRUE ;
If InfoImp.NomFic='AUTO' Then Exit ;
pbSISCO:=FALSE ;
If InfoImp.FormatOrigine='SIS' Then
   BEGIN
   j:=TransfertSISCO(InfoImp.NomFic,FALSE,k) ;
   If j<>0 Then BEGIN PbSISCO:=TRUE ; Result:=FALSE ; END ;
   InfoImp.NomFicOrigine:=InfoImp.NomFic ;
   InfoImp.NomFic:=NewNomFic(InfoImp.NomFic,'CGN') ;
   END ;
LePath:=ExtractFilePath(InfoImp.NomFic) ;
LExtension:=ExtractFileExt(InfoImp.NomFic) ;
LeNom:=ExtractFileName(InfoImp.NomFic) ;
i:=Pos('.',LeNom) ; If i>0 Then LeNom:=Copy(LeNom,1,i-1) ;
InfoImp.NomFicDoublon:=NewNomFicEtDir(InfoImp.NomFic,'Dbl','Doublons') ;
InfoImp.NomFicRejet:=NewNomFicEtDir(InfoImp.NomFic,'Err','Rejets') ;
InfoImp.NomFicRapport:=NewNomFicEtDir(InfoImp.NomFic,'Rap','Rapports') ;
CreateFicRapport(InfoImp) ;

END ;

Procedure RecupNomFic(Var InfoImp : TInfoImport ; ListeFichier : TStringList) ;
Var SearchRec : TSearchRec ;
    LeChemin,LePrefixe,LeSuffixe,LeNom : String ;
BEGIN
LeChemin:='' ; LePrefixe:='*' ; LeSuffixe:='*' ;
If InfoImp.Sc.Chemin<>'' Then LeChemin:=InfoImp.Sc.Chemin+'\' ;
If InfoImp.Sc.Prefixe<>'' Then LePrefixe:=InfoImp.Sc.Prefixe+'*' ;
If InfoImp.Sc.Suffixe<>'' Then LeSuffixe:=InfoImp.Sc.Suffixe ;
LeNom:=LeChemin+LePrefixe+'.'+LeSuffixe ;
If FindFirst(LeNom, faAnyFile, SearchRec)=0 Then
  BEGIN
  If (Pos('Dbl',SearchRec.Name)=0) And  (Pos('Err',SearchRec.Name)=0) Then ListeFichier.Add(LeChemin+SearchRec.Name) ;
  While (FindNext(SearchRec) = 0) Do
    If (Pos('Dbl',SearchRec.Name)=0) And  (Pos('Err',SearchRec.Name)=0) Then ListeFichier.Add(LeChemin+SearchRec.Name) ;
  END ;
FindClose(SearchRec);
END ;

Function FichierSuivant(ListeFichier : TStringList ; Var InfoImp : TInfoImport ; Var NumFic : Integer) : Boolean ;
BEGIN
Result:=FALSE ;
If NumFic<=ListeFichier.Count-1 Then
  BEGIN
  InfoImp.NomFic:=ListeFichier[NumFic] ;
  If Not RepositionneNomFic(InfoImp) Then BEGIN Result:=FALSE ; Exit ;END ;
  Inc(NumFic) ;
  END Else Exit ;
Result:=TRUE ;
END ;

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
InfoImp.LSoucheBOR:=TStringList.Create ;
END ;

Function AlimInfoImp(Var InfoImp :TInfoImport; bMulti : boolean = False) : Boolean ;
BEGIN
Result:=TRUE ;
FillChar(InfoImp,SizeOf(InfoImp),#0) ;
CreateListeImp(InfoImp) ;
if bMulti then
begin
  LanceFicheParametrageImportMulti ( InfoImp.Format, InfoImp.CodeFormat);
  InfoImp.Lequel:='FEC' ;
  InfoImp.NomFic:='AUTO' ;
end else
begin
  InfoImp.Lequel:=ParamStr(4) ;
  InfoImp.Format:=ParamStr(5) ;
  If InfoImp.Format='ORLI' Then BEGIN InfoImp.Format:='SN2' ; Sn2Orli:=TRUE ; END ;
  InfoImp.ImportAuto:=TRUE ;
  InfoImp.NomFic:=ParamStr(6) ;
  If ParamCount>6 Then InfoImp.CodeFormat:=ParamStr(7) ;
end;
PasDeBlanc:=FALSE ;
If ParamCount>0 Then If ParamStr(ParamCount)='ORLIBLANCS' Then PasDeBlanc:=TRUE ;
InfoImp.ForceQualif:='' ;
InfoImp.ForcePositif:=FALSE ;
InfoImp.ForceBourrage:=TRUE ;
InfoImp.CtrlDB:=FALSE ;
InfoImp.ForceNumPiece:=FALSE ;
ChargeScenarioImport(InfoImp,True) ;
If InfoImp.Sc.EstCharge Then
  BEGIN
  If InfoImp.SC.Doublon Then
    BEGIN
    InfoImp.CtrlDB:=TRUE ;
    If InfoImp.SC.ForcePiece Then InfoImp.ForceNumPiece:=TRUE ;
    END ;
  END ;
If InfoImp.Format='SIS' Then BEGIN InfoImp.FormatOrigine:='SIS' ; InfoImp.Format:='CGN' ;END ;
If Not RepositionneNomFic(InfoImp) Then Result:=FALSE ;
END ;


Procedure initImporte(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.NbGenFaux:=0 ; InfoImp.NbAuxFaux:=0 ;
InfoImp.NbAnaFaux:=0 ; InfoImp.NbJalFaux:=0 ;
VideListeInfoImp(InfoImp,FALSE) ;
InfoImp.ListeEnteteDoublon.Sorted:=TRUE ;
InfoImp.ListeEnteteDoublon.Duplicates:=DupIgnore ;
END ;


Procedure initVerif(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.ListeEntetePieceFausse.Clear ;
InfoImp.ListeEntetePieceFausse.Sorted:=TRUE ;
InfoImp.ListeEntetePieceFausse.Duplicates:=DupIgnore ;
VideListe(InfoImp.ListePieceFausse) ;
InfoImp.ListePbAna.Clear ;
InfoImp.ListePbAna.Sorted:=TRUE ;
InfoImp.ListePbAna.Duplicates:=DupIgnore ;
END ;

Procedure InitIntegre(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.ListePieceIntegre.Clear ;
InfoImp.ListePieceIntegre.Sorted:=TRUE ;
InfoImp.ListePieceIntegre.Duplicates:=DupIgnore ;
END ;

Procedure DetruitLeFichier(St : String) ;
Var Fichier : TextFile ;
    FichierOk : Boolean ;
    i : Integer ;
BEGIN
i:=0 ;
Repeat
  Inc(i) ;
  AssignFile(Fichier,St) ;
  {$I-} Reset (Fichier) ; {$I+}
  FichierOk:=ioresult=0 ;
  If FichierOk Then Close(Fichier) ;
Until FichierOk Or (i>200) ;
If FichierOk Then
  BEGIN
  {$i-}
  AssignFile(Fichier,St) ; Erase(Fichier) ;
  {$i+}
  END ;
END ;

Procedure DetruitFichier(Var InfoImp :TInfoImport) ;
Var Fichier : TextFile ;
    FichierOk : Boolean ;
    i : Integer ;
BEGIN
If InfoImp.Sc.DetruitFic Then
  BEGIN
  DetruitLeFichier(InfoImp.NomFic) ;
  If (InfoImp.FormatOrigine='SIS') And (InfoImp.NomFicOrigine<>'') Then DetruitLeFichier(InfoImp.NomFicOrigine) ;
  END ;
END ;

procedure TFImpauto.LanceImportAuto ;
var Soc : String ;
    InfoImp : PtTInfoImport ;
    Qui : String ;
    Auto,PlusDeFichierAImporter,OkFic : Boolean ;
    ListeFichier : TStringList ;
    NumFic : Integer ;
  Label 0 ;
begin
  if FMultiDossier then
  begin
    WindowState := wsNormal;
    Visible := True;
    BringToFront;
    PlusDeFichierAImporter:=TRUE ;
    Auto:=FALSE ; Init ;
    if DBSOC<>NIL Then DeconnecteHalley ;
    Soc := FSociete;
  end else
  begin
    If ParamCount<6 Then Exit ; PlusDeFichierAImporter:=TRUE ;
    Auto:=FALSE ; Init ;
    if DBSOC<>NIL Then DeconnecteHalley ;
    Qui:=ParamStr(1) ;
    Soc:=ParamStr(2) ;
    V_PGI.UserLogin:=ParamStr(3) ;
  end;
if ConnecteHalley(Soc,True,@ChargeMagHalley,NIL,NIL,NIL) then
   BEGIN
   END else
   BEGIN
   If (DBSOC<>NIL) AND (DBSOC.Connected) then DBSOC.Connected:=FALSE ;
   SourisNormale ; Exit ;
   END ;
New(InfoImp) ; If Not AlimInfoImp(InfoImp^, FMultiDossier) Then Goto 0 ;
if FMultiDossier then
begin
  if (InfoImp.Format='') or (InfoImp.CodeFormat='') then
  begin
    VideListeInfoImp(InfoImp^,TRUE) ;
    Dispose(InfoImp) ;
    Logout ; DeconnecteHalley ;
    SourisNormale ;
    exit;
  end;
end;
If InfoImp.NomFic='AUTO' Then Auto:=TRUE ;
If Auto Then
  BEGIN
  ListeFichier:=TStringList.Create ;
  RecupNomFic(InfoImp^,ListeFichier) ;
  END ;
NumFic:=0 ;
Repeat
  OkFic:=TRUE ; If Auto Then BEGIN OkFic:=FichierSuivant(ListeFichier,InfoImp^,NumFic) ; VideListeInfoImp(InfoImp^,FALSE) ; END ;
  If OkFic Then
    BEGIN
    PlusDeFichierAImporter:=FALSE ;
    MessImport.Visible:=TRUE ;
    MessImport.Caption:='Importation en cours du fichier '+InfoImp^.NomFic ;
    If ImporteLesEcritures(nil,InfoImp^) then
      BEGIN
      InitVerif(InfoImp^) ;
//      if not VerPourImp2(InfoImp^.ListeEntetePieceFausse,InfoImp^.ListePieceFausse,TRUE,tvEcr) then
      if not VerPourImp3(InfoImp^.ListeEntetePieceFausse,InfoImp^.ListePbAna,InfoImp^.ListePieceFausse,TRUE,tvEcr,InfoImp^.SC.ShuntPbAna,InfoImp^.PbAna) then
        BEGIN
        MajImpErr(InfoImp.ListeEntetePieceFausse) ;
        END ;
    //     if AucunMvtControlesOk then BEGIN HM.Execute(42,'','') ; Exit ; END ;
      InitIntegre(InfoImp^) ;
      IntegreEcr(Nil,InfoImp^) ;
      END else VH^.ImportRL:=FALSE ;
    FaitFichierDoublon(InfoImp^) ; FaitFichierRejet(InfoImp^) ; DetruitFichier(InfoImp^) ;
    END Else PlusDeFichierAImporter:=TRUE ;
  If Not Auto Then PlusDeFichierAImporter:=TRUE ;
Until PlusDeFichierAImporter ;
If Auto Then ListeFichier.Free ;
VideListeInfoImp(InfoImp^,TRUE) ;
0:Dispose(InfoImp) ;
Logout ; DeconnecteHalley ;
SourisNormale ;
END ;


procedure TFImpauto.Timer1Timer(Sender: TObject);
begin
Timer1.Enabled:=FALSE ;
LanceImportAuto ;
Close ;
end;

procedure RenseigneSerie ;
var Buffer: array[0..1023] of Char;
    sWinPath,sIni : String ;
begin
HalSocIni:='CCS5.ini' ;
GetWindowsDirectory(Buffer,1023);
SetString(sWinPath, Buffer, StrLen(Buffer));
sIni:=sWinPath+'\'+HALSOCINI ;
if FileExists(sIni) then
   BEGIN
   HalSocIni:='CCS5.ini' ;
   NomHalley:= 'Comptabilité S5' ;
   TitreHalley := 'Imports / Exports S5' ;
   V_PGI.LaSerie:=S5 ;
   Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS5.HLP' ;
   END else
   BEGIN
   HalSocIni:='CCS7.ini' ;
   NomHalley:= 'Comptabilité S7' ;
   TitreHalley := 'Imports / Exports S7' ;
   V_PGI.LaSerie:=S7 ;
   Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS7.HLP' ;
   END ;
END ;

Procedure InitLaVariablePGI;
Begin
  Apalatys:='CEGID' ;
  Copyright:='© Copyright ' + Apalatys ;
  V_PGI.OutLook:=FALSE ;
  V_PGI.OfficeMsg:=FALSE ;
  V_PGI.VersionDemo:=FALSE ;
  V_PGI.SAV:=True ;

  V_PGI.CegidAPalatys:=FALSE ;
  V_PGI.CegidBureau:=TRUE ;
  V_PGI.StandardSurDP:=True ;
  V_PGI.MajPredefini:=False ;
  V_PGI.MultiUserLogin:=False ;

  ChargeXuelib ;

  V_PGI.MenuCourant:=0 ;
{$IFDEF CCS3}
//V_PGI.NumVersion:='1.50' ; V_PGI.NumBuild:='1' ; V_PGI.NumVersionBase:=563;
V_PGI.NumVersion:='4.2.0' ; V_PGI.NumBuild:='10' ; V_PGI.NumVersionBase:=595;
{$ELSE}
V_PGI.NumVersion:='4.2.0' ; V_PGI.NumBuild:='10' ; V_PGI.NumVersionBase:=595;
{$ENDIF}
(*
If EstSerie(S3) Then V_PGI.NumVersion:='1.9.0' Else
 If EstSerie(S5) Then V_PGI.NumVersion:='3.9.4' Else
  If EstSerie(S7) Then V_PGI.NumVersion:='3.9.4' ;
 *)
V_PGI.DateVersion:=EncodeDate(2002,01,17) ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.ParamSocLast:=True ;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=False ;

RenseigneLaSerie(ExeCCAUTO) ;

end;

procedure TFImpauto.FormCreate(Sender: TObject);
begin
  FMultiDossier :=  ( Pos ('/DOSSIER', CmdLine ) > 0 );
  if FMultiDossier then
  begin
    Timer1.Enabled := False;
    if V_PGI.CegidBureau then
    begin
      PGIAppAlone:=True ;
      CreatePGIApp ;
      InitParametreMultiDossier;
      Timer1.Enabled := True;
    end;
  end;
end;

procedure TFImpauto.InitParametreMultiDossier;
var i : integer;
    St, Nom, Value : string;
begin
  VStatus:=Status ;
  Status.Caption := Copyright ;
  if FMultiDossier then
  begin
    VH^.ModeSilencieux:=TRUE ;
    FSociete := '';
{$IFDEF AGL570D}
  HalSocIni := 'CEGIDPGI.INI';
{$ELSE}
    V_PGI_Env := TPGIEnv.Create ;
    //# lit param de PGIApp et màj V_PGI.RunFromLanceur
    InitPGIEnv();
    //# retour au mode APA
    if V_PGI_Env.ModeFonc='APA' then
    begin
      V_PGI_Env.Free;
      V_PGI_Env := Nil;
    end
    //# routage halsocini
    else If Copy(V_PGI_Env.ModeFonc,1,1)='M' then HalSocIni := 'CEGIDPGI.INI';
{$ENDIF}

    //# évite message bloquant
    If V_PGI.RunFromLanceur then V_PGI.MultiUserLogin := True;

    V_PGI.UserLogin:='' ;
    V_PGI.PassWord:='' ;
    FSociete:='' ;
    for i:=1 to ParamCount do
    begin
      St:=ParamStr(i) ;
      Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
      Value:=UpperCase(Trim(St)) ;
      //Paramètres de connexion
      if Nom='/MAJSTRUCTURE' then BEGIN (*kVerifStructure:=(Value<>'FALSE') ;*) END ;
      if Nom='/USER'     then
      begin
        V_PGI.UserLogin:=Value ;
      end;
      if Nom='/PASSWORD' then BEGIN V_PGI.PassWord:=DecryptageSt(Value) ; END ;
      if Nom='/DATE'     then BEGIN  V_PGI.DateEntree:=StrToDate(Value) ; END ;
      if Nom='/DOSSIER'  then
      begin
        FSociete:=Value ;
      end;
    end;
  end;
end;

Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;

end.

