unit MDispIM;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, HStatus, StdCtrls, Hctrls, ExtCtrls, PGIExec, HEnt1, LicUtil,
  MajTable,Ent1, dbtables, HMsgBox, ImpCegid, HFLabel, EntPGI, ImSaiCoef,
  ParamSoc;

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
  if V_PGI.ModePCL='1' then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
end;

procedure TFMain.FormCreate(Sender: TObject);
begin
  PGIAppAlone:=True ;
  CreatePGIApp ;
end;

procedure TFMain.FormShow(Sender: TObject);
Var i,j : Integer ;
    St,St1,Nom,Value : String ;
begin
  VStatus:=Status ;
  Status.Caption := Copyright ;
  SocPCL:='' ; RepPCl:='' ; FichierPCL:='' ; MonnaiePCL:='' ;
  HalSocIni := 'CEGIDPGI.INI';
  V_PGI.UserLogin:='' ;
  V_PGI.PassWord:='' ;
  SocPCL:='' ;
  for i:=1 to ParamCount do
  BEGIN
    St := ParamStr(i);
    Nom := UpperCase(Trim(ReadTokenPipe(St, '=')));
    Value := UpperCase(Trim(St));
    { Utilisateur }
    if Nom = '/USER' then V_PGI.UserLogin :=Value
    { Mot de passe }
    else if Nom = '/PASSWORD' then V_PGI.PassWord:= Value
    { Date d'entrée }
    else if Nom = '/DATE' then V_PGI.DateEntree := StrToDate(Value)
    { Dossier }
    else if Nom = '/DOSSIER' then
    begin
      SocPCL:=Value;
      V_PGI.CurrentAlias := Value;
      V_PGI.DefaultSectionName := '';
      j := pos('@', Value);
      if (j > 0) then
      begin
        V_PGI.DefaultSectionName := Copy(Value, j + 1, 255);
        V_PGI.CurrentAlias := Copy(Value, 1, j - 1);
        if (j > 3) and (Copy(Value, 1, 2) = 'DB') then
          V_PGI.NoDossier := Copy(Value, 3, j - 3); // sinon reste à '000000'
      end;
      V_PGI.RunFromLanceur := (V_PGI.DefaultSectionName <> '');
      if v_pgi.DefaultSectionName <> '' then
        if v_pgi.DBName = v_pgi.DefaultSectionDBName then
          v_pgi.InBaseCommune := true;
    end
    { Mode de fonctionnement }
    else if Nom = '/MODEPCL' then
    begin
      begin
        if assigned(v_pgi) then v_pgi.ModePCL := Value;
      end;
    end else
    { Paramètres de transfert }
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
  V_PGI.RunFromLanceur := True;
  if DBSOC<>NIL Then BEGIN Logout ; DeconnecteHalley ; END ;
  if ConnecteHalley(SocPCL,True,@ChargeMagHalley,NIL,NIL,NIL) then
  BEGIN
    ApresConnecte ;
    Visible := False;
    if not ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="ICD"') then
      InstalleLesCoefficientsDegressifs ;
    SavImo2PGI ( RepPCL+FichierPCL, True);
    SetParamSoc('SO_I_CPTAPGI',False);    
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
// ajout me
  if (V_PGI.NoDossier<>'') and  (V_PGI.Driver in [dbMssql, dbMssql2005]) then
  begin
    try
      ExecuteSql ('checkpoint');
      ExecuteSql ('backup log '+v_pgi.DBName+' with no_log');
      ExecuteSql ('dbcc shrinkdatabase('+v_pgi.DBName+')');
      ExecuteSql ('dbcc shrinkfile(2)');
    except;
    end;
  end;
  Logout;DeconnecteHalley;
end;

Procedure InitLaVariablePGI;
Begin
  RenseigneLaSerie(ExeCCAUTO) ;

  Apalatys:='CEGID' ;
  Copyright:='© Copyright ' + Apalatys ;
  V_PGI.OutLook:=FALSE ;
  V_PGI.OfficeMsg:=FALSE ;
  V_PGI.VersionDemo:=FALSE ;
  V_PGI.SAV:=True ;
  V_PGI.BlockMAJStruct:=True ;
  V_PGI.EuroCertifiee:=false ;

  ChargeXuelib ;

  V_PGI.MenuCourant:=0 ;
  V_PGI.ImpMatrix := True ;
  V_PGI.OKOuvert:=FALSE ;
  V_PGI.Halley:=TRUE ;
  V_PGI.LaSerie:=S5 ;
  V_PGI.ParamSocLast:=True ;
  V_PGI.PGIContexte:=[ctxCompta, ctxPCL] ;
  V_PGI.CegidAPalatys:=FALSE ;
  V_PGI.CegidBureau:=TRUE ;
  V_PGI.StandardSurDP:=True ;
  V_PGI.MajPredefini:=False ;
end;

Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;
end.
