{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 28/01/2004
Modifié le ... :   /  /
Description .. : Menu utilitaire archivage/restauration/transport dossier
Mots clefs ... : PAIE
*****************************************************************}

unit MenuDispCIEXPS5;

interface
uses HEnt1,
  Forms,
  sysutils,
  HMsgBox,
  Classes,
  Controls,
  HPanel,
  UIUtil,
  Windows,
  Hdebug,
  FE_Main,
  AGLInit,
  MenuOLG,
  PgiExec,
  PGIEnv,
  EntPgi,
  {$IFDEF MEMCHECK}
  memcheck,
  {$ENDIF}
  ParamSoc,
  Hctrls;

procedure InitApplication;
type
  TFMenuDisp = class(TDatamodule)
  end;

var
  FMenuDisp: TFMenuDisp;

implementation

{$R *.DFM}
// Procedure qui indique que le dossier en mode nomade . Uniquement PCL !!!
procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
begin
  case Num of
    10: ;
    11: ;
    12:
      begin // avant connexion
        if (V_PGI_Env <> nil) and (V_PGI_env.ModePcl = '1') then
          FMenuG.SetSeria('00280011', ['00250050'], ['Paie et GRH']) //PCL
        else FMenuG.SetSeria('05970012', ['05050050'], ['Paie et GRH']); // Paie entreprise
      end;

    13: ;
    16:;
    100: ; // action A Faire quand on passe par le lanceur
    46610 : AglLanceFiche ('PAY', 'GESTION_STD','','', 'P');
    46620 : AglLanceFiche ('PAY', 'GESTION_DOS','','', '');
  end;
end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT: string; Range: string);
begin
  case Num of
    1: ;
  end;
end;

procedure ZoomEdt(Qui: integer; Quoi: string);
begin
  Screen.Cursor := crDefault;
end;

procedure ZoomEdtEtat(Quoi: string);
var
  i, Qui: integer;
begin
  i := Pos(';', Quoi);
  Qui := StrToInt(Copy(Quoi, 1, i - 1));
  Quoi := Copy(Quoi, i + 1, Length(Quoi) - i);
  ZoomEdt(Qui, Quoi);
end;


procedure RechargeMenuPopPaie;
begin
  //
end;
//
procedure AfterChangeModule(NumModule: Integer);
begin
  FMenuG.RemoveGroup(49100, TRUE);
  FMenuG.RemoveGroup(49200, TRUE);
  FMenuG.RemoveGroup(49400, TRUE);
  FMenuG.RemoveGroup(-49500, TRUE);
  FMenuG.RemoveGroup(49300, TRUE);
  FMenuG.RemoveGroup(49700, TRUE);
  FMenuG.RemoveGroup(-49800, TRUE);
  FMenuG.RemoveGroup(49900, TRUE);
end;

procedure AfterProtec(sAcces: string); //Sérialisation des modules
begin
  // cas où on coche "version non sérialisée' dans
  // l'écran de séria. On force la non sérialisation
  if V_PGI.NoProtec then V_PGI.VersionDemo := TRUE
  else V_PGI.VersionDemo := False;
  if (sAcces[1] = '-') then V_PGI.VersionDemo := TRUE;

  FMenuG.SetModules([49], [49]);
end;

procedure AGLAfterChangeModule(Parms: array of variant; nb: integer);
begin
  AfterChangeModule(Integer(Parms[0]));
end;
procedure InitApplication;
begin
  {$IFDEF MEMCHECK}
  MemCheckLogFileName := ChangeFileExt(Application.exename, '.log');
  MemChk;
  {$ENDIF}

  //  ProcCalcEdt:=CalcOLEEtatPG;
  ProcZoomEdt := ZoomEdtEtat;

  FMenuG.OnDispatch := Dispatch;

  //  FMenuG.OnChargeMag:=ChargeMagHalley  ;
  {$IFNDEF EAGLCLIENT}
  //  FMenuG.OnMajAvant:=MAJHalleyAvant ;
  //  FMenuG.OnMajApres:=MAJHalleyApres ;
  {$ENDIF}
  // Mais initialisation faite quand meme pour avoir la variable sAcces definie pour la fonction AfterProtect
  if (V_PGI_Env <> nil) and (V_PGI_env.ModePcl = '1') then
    FMenuG.SetSeria('00280011', ['00250050'], ['Paie et GRH']) //PCL
  else FMenuG.SetSeria('05970012', ['05050050'], ['Paie et GRH']); // Paie entreprise
  FMenuG.OnAfterProtec := AfterProtec;
  FMenuG.SetModules([49], [49]);
  FMenuG.OnChangeModule := AfterChangeModule;
  V_PGI.DispatchTT := DispatchTT;
end;

{ TFMenuDisp }
procedure InitLaVariablePGI;
begin
  Apalatys := 'CEGID';
  HalSocIni := 'cegidpgi.ini';
  Copyright := '© Copyright ' + Apalatys;
  V_PGI.OutLook := TRUE;
  V_PGI.LaSerie := S5;
  TitreHalley := 'PAIE-OUTIL S5';
  NomHalley := 'CIEXPS5';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CIEXPS5.HLP';
  V_PGI.OfficeMsg := FALSE;
  V_PGI.ToolsBarRight := TRUE;
  ChargeXuelib;
  if (V_PGI_Env <> nil) and (V_PGI_env.ModePcl = '1') then V_PGI.PortailWeb := 'http://experts.cegid.fr/home.asp'
  else V_PGI.PortailWeb := 'http://utilisateurs.cegid.fr';
  V_PGI.OfficeMsg := TRUE;
  V_PGI.VersionDemo := True;
  V_PGI.MenuCourant := 0;
  V_PGI.VersionReseau := False;
  V_PGI.NumVersion := '5.05';
  V_PGI.NumBuild := '000.035';
  V_PGI.NumVersionBase := 620;
  V_PGI.DateVersion := EncodeDate(2004, 03, 10);
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  V_PGI.MenuCourant := 0;
  V_PGI.DebugSQL := False;
  V_PGI.Debug := False;
  V_PGI.QRPdf := True;
  V_PGI.NumMenuPop := 27;
  V_PGI.CegidBureau := True;
  V_PGI.CegidApalatys := FALSE;
  // pour autoriser l'acces au dp comme dossier de travail,au multi dossier
  V_PGI.StandardSurDP := true;
  // pour autoriser l'acces au dp comme base de reference pour tous les dossiers ==> Table DP+STD
  V_PGI.MajPredefini := True;
  // Confidentialité
  V_PGI.PGIContexte := [CtxPaie];
  V_PGI.RAZForme := TRUE;
  // Blocage des mises à jour de structure par l'appli. Elle doit etre faite par PgiMajVEr fourni dans le kit socref
  V_PGI.BlockMAJStruct := TRUE;
  v_pgi.sav := False;
end;

initialization
  ProcChargeV_PGI := InitLaVariablePGI;
end.

