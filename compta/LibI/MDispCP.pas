unit MDispCP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, HStatus, StdCtrls, Hctrls, ExtCtrls, PGIExec, HEnt1, PGIEnv, LicUtil,
  MajTable,Ent1, dbtables, HMsgBox, ImpCegid, HFLabel;

type
  TFMain = class(TForm)
    Panel1: TPanel;
    Status: THStatusBar;
    Timer: TTimer;
    FlashingLabel1: TFlashingLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    SocPCL : String ;
    RepPCl,FichierPCL,MonnaiePCL : String ;
    procedure LanceImportAuto;
    Procedure ApresConnecte ;
  public
    { Déclarations publiques }
  end;

var
  FMain: TFMain;

implementation

{$R *.DFM}

procedure TFMain.ApresConnecte;
BEGIN
  if V_PGI_Env<>Nil then
  begin
    //# mono-entreprise : pointe l'environnement de la soc choisie
    if V_PGI_Env.ModeFonc='MONO' then V_PGI_Env.SocCommune := SOCPCL ;
    //# mode multi-dossier : appli lancée sans passer par le lanceur...
    if (V_PGI_Env.ModeFonc='MULTI') and (Not V_PGI.RunFromLanceur) and (Not V_PGI_Env.InTheLanceur) then
    //manque le double choix : soc commune, dossier
    PGIInfo('Vous lancez une application en mode multi-dossier sans passer par le lanceur.', TitreHalley);
    //# mode multi-dossier : on est dans le lanceur => connecter la soc...
  end;
  if ((V_PGI_Env<>Nil) and (V_PGI_Env.ModePCL='1')) then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
end;

procedure TFMain.FormCreate(Sender: TObject);
begin
  PGIAppAlone:=True ;
  CreatePGIApp ;
end;

procedure TFMain.FormShow(Sender: TObject);
Var i : Integer ;
    St,St1,Nom,Value : String ;
begin
  VStatus:=Status ;
  Status.Caption := Copyright ;
  SocPCL:='' ; RepPCl:='' ; FichierPCL:='' ; MonnaiePCL:='' ;
  // Création de l'instance unique
  V_PGI_Env := TPGIEnv.Create ;
  //# lit param de PGIApp et màj V_PGI.RunFromLanceur
  InitPGIEnv();
  //# retour au mode APA
  if V_PGI_Env.ModeFonc='APA' then
     begin V_PGI_Env.Free; V_PGI_Env := Nil; end
  //# routage halsocini
  else If Copy(V_PGI_Env.ModeFonc,1,1)='M' then HalSocIni := 'CEGIDPGI.INI';
  //# évite message bloquant
  If V_PGI.RunFromLanceur then V_PGI.MultiUserLogin := True;
  V_PGI.UserName:='' ;
  V_PGI.PassWord:='' ;
  SocPCL:='' ;
  for i:=1 to ParamCount do
  BEGIN
    St:=ParamStr(i) ;
    Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
    Value:=UpperCase(Trim(St)) ;
    //Paramètres de connexion
    if Nom='/MAJSTRUCTURE' then BEGIN (*kVerifStructure:=(Value<>'FALSE') ;*) END ;
    if Nom='/USER'     then
    BEGIN
      V_PGI.UserName:=Value ;
    END ;
    if Nom='/PASSWORD' then BEGIN V_PGI.PassWord:=DecryptageSt(Value) ; END ;
    V_PGI.PassWord:=CryptageSt(DayPass(Date)) ;
    if Nom='/DATE'     then BEGIN  V_PGI.DateEntree:=StrToDate(Value) ; END ;
    if Nom='/DOSSIER'  then
    BEGIN
      SocPCL:=Value ;
    END ;
    if Nom='/TRF'  then
    BEGIN
      If Value[Length(Value)]<>';' Then Value:=Value+';' ;
      St1:=ReadTokenSt(Value) ; If St1<>'' Then RepPCL:=St1 ;
      St1:=ReadTokenSt(Value) ; If St1<>'' Then FichierPCL:=St1 ;
      St1:=ReadTokenSt(Value) ; If St1<>'' Then MonnaiePCL:=St1 ;
    END ;
  END;
end;

procedure TFMain.LanceImportAuto;
begin
  VH^.ModeSilencieux := True;
  VH^.OkModImmo := True;
  VH^.OkModCompta := True;  
  if DBSOC<>NIL Then BEGIN Logout ; DeconnecteHalley ; END ;
  if ConnecteHalley(SocPCL,True,@ChargeMagHalley,NIL,NIL,NIL) then
  BEGIN
    ApresConnecte ;
//    SavImo2PGI ( RepPCL+FichierPCL, True);
    MessageAlerte('Hello');
  END else
  BEGIN
    If (DBSOC<>NIL) AND (DBSOC.Connected) then DBSOC.Connected:=FALSE ;
    SourisNormale ; Exit ;
  END ;
end;

procedure TFMain.TimerTimer(Sender: TObject);
begin
Timer.Enabled:=FALSE ;
LanceImportAuto;
Close ;
end;

procedure TFMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Logout;DeconnecteHalley;
end;

initialization
Apalatys:='CEGID' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.OutLook:=FALSE ;
V_PGI.OfficeMsg:=FALSE ;
//V_PGI.ToolsBarRight:=TRUE ;
V_PGI.VersionDemo:=FALSE ;
V_PGI.SAV:=True ;
ChargeXuelib ;
V_PGI.MenuCourant:=0 ;
//V_PGI.NumVersion:='3.50 ß' ;
V_PGI.NumVersion:='3.80' ; V_PGI.NumBuild:='103' ; V_PGI.NumVersionBase:=506 ;
V_PGI.DateVersion:=EncodeDate(2000,03,08) ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.NiveauAcces:=1 ;
V_PGI.LaSerie:=S5 ;
V_PGI.ParamSocLast:=True ;
V_PGI.PGIContexte:=[ctxCompta] ;
V_PGI.CegidAPalatys:=FALSE ;
V_PGI.CegidBureau:=TRUE ;
V_PGI.StandardSurDP:=True ;
V_PGI.MajPredefini:=False ;
end.
