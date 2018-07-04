{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 08/04/2004
Modifié le ... :   /  /
Description .. : Menu de la TOX paie
Mots clefs ... : PAIE
*****************************************************************}

unit MenuDispPTS5;

interface
Uses HEnt1,
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
//     PGIEnv,
     ParamSoc,
     Hctrls,
     UtoxConf,
     UtoxConst,
     UtoxFiches,
{$IFDEF MEMCHECK}
     memcheck ,
{$ENDIF}
     EntToxPaie ;

Procedure InitApplication ;
type
  TFMenuDisp = class(TDatamodule)
    end ;

Var FMenuDisp : TFMenuDisp ;

implementation

uses  PaieToxCtrl,
      PaieToxWork,
      PaieToxConsultBis,
      PaieToxMoteur,
      PaieToxDetop,
      PaieToxSim,
      PaieToxIntegre ;

{$R *.DFM}
// Procedure qui indique que le dossier en mode nomade . Uniquement PCL !!!
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
BEGIN
   case Num of
   10 : ;
   11 :  ;
   12 : begin  // avant connexion
          FMenuG.SetSeria('05970012',['05050050'],['Paie et GRH']) ;  // Paie entreprise
        end;
   13 : PaieToxLibereInfoPaie ;
   100 : ;  // action A Faire quand on passe par le lanceur
   49312 : AglToxConf( aceConfigure, '', PaieNotConfirmeTox, PaieTraitementTox, Nil, PaieAvantArchivageTOX ) ;
   49314 : AglToxMulChronos ; // Consultation des échanges
   49316 : PaieAfficheCorbeille; // Consultation des corbeilles
   49341 : AglToxSaisieParametres ; // Paramètres par défaut
   49342 : AglToxSaisieVariables ; // Variables
   49343 : AglToxMulSites ;      // Sites
   49344 : AglToxSaisieGroupes ; //  Groupes
   49345 : AglToxMulConditions ; // Requêtes
   49346 : AglToxMulEvenements ; // Evénements
   49362 : PaieToxSimulation ; // Intégration d'un fichier
   49364 : PaieConsultManuelTox ; // Visualisation d'un fichier
   49366 : PaieDetopeTox ; // Détopage d'un fichier
   end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT : String; Range : String ) ;
BEGIN
   case Num of
     1 : ;
     end ;
END ;

procedure ZoomEdt(Qui : integer;Quoi : string);
begin
  Screen.Cursor:=crDefault;
end;

procedure ZoomEdtEtat(Quoi : string);
var i,Qui : integer;
begin
  i:=Pos(';',Quoi);
  Qui:=StrToInt(Copy(Quoi,1,i-1));
  Quoi:=Copy(Quoi,i+1,Length(Quoi)-i);
  ZoomEdt(Qui,Quoi);
end;


procedure RechargeMenuPopPaie;
begin
//
end;
//
procedure AfterChangeModule (NumModule : Integer);
begin
  FMenuG.RemoveGroup(49100,TRUE);
  FMenuG.RemoveGroup(49200,TRUE);
  FMenuG.RemoveGroup(49400,TRUE);
  FMenuG.RemoveGroup(-49500,TRUE);
  FMenuG.RemoveGroup(49600,TRUE);
  FMenuG.RemoveGroup(49700,TRUE);
  FMenuG.RemoveGroup(-49800,TRUE);
  FMenuG.RemoveGroup(49900,TRUE);
end;

Procedure AfterProtec ( sAcces : String ) ;//Sérialisation des modules
BEGIN
// cas où on coche "version non sérialisée' dans
// l'écran de séria. On force la non sérialisation
  If V_PGI.NoProtec then V_PGI.VersionDemo:= TRUE
    else V_PGI.VersionDemo :=False;
  if (sAcces[1]='-') then   V_PGI.VersionDemo:= TRUE;

  FMenuG.SetModules([49],[49]) ;
END ;

Procedure AGLAfterChangeModule ( Parms : array of variant ; nb : integer ) ;
BEGIN
  AfterChangeModule(Integer(Parms[0])) ;
END ;
procedure InitApplication ;
BEGIN
{$IFDEF MEMCHECK}
   MemCheckLogFileName:=ChangeFileExt(Application.exename,'.log') ;
   MemChk ;
{$ENDIF}

//  ProcCalcEdt:=CalcOLEEtatPG;
  ProcZoomEdt:=ZoomEdtEtat ;

  FMenuG.OnDispatch:=Dispatch ;

  //  FMenuG.OnChargeMag:=ChargeMagHalley  ;
{$IFNDEF EAGLCLIENT}
//  FMenuG.OnMajAvant:=MAJHalleyAvant ;
//  FMenuG.OnMajApres:=MAJHalleyApres ;
{$ENDIF}
// Mais initialisation faite quand meme pour avoir la variable sAcces definie pour la fonction AfterProtect
  FMenuG.SetSeria('05970012',['05050065'],['Paie et GRH']) ;  // Paie entreprise
  FMenuG.OnAfterProtec:=AfterProtec ;
  FMenuG.SetModules([49],[49]) ;
  FMenuG.OnChangeModule:=AfterChangeModule;
  V_PGI.DispatchTT:=DispatchTT ;
END ;

{ TFMenuDisp }
Procedure InitLaVariablePGI;
Begin
  Apalatys:='CEGID' ;
  HalSocIni:='cegidpgi.ini' ;
  Copyright:='© Copyright ' + Apalatys ;
  V_PGI.OutLook:=TRUE ;
  V_PGI.LaSerie:=S5 ;
  TitreHalley:='PAIE-TOX S5' ;
  NomHalley:='CTPS5' ;
  Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CTPS5.HLP';
  V_PGI.OfficeMsg:=FALSE;
  V_PGI.ToolsBarRight:=TRUE ;
  ChargeXuelib ;
  if V_PGI.ModePcl='1' then V_PGI.PortailWeb:='http://experts.cegid.fr/home.asp'
  else V_PGI.PortailWeb:='http://utilisateurs.cegid.fr';
  V_PGI.OfficeMsg:=TRUE;
  V_PGI.VersionDemo:=True ;
  V_PGI.MenuCourant:=0 ;
  V_PGI.VersionReseau:=False ;
  V_PGI.NumVersion := '7.00';
  V_PGI.NumBuild := '000.0';
  V_PGI.NumVersionBase:=749;
  V_PGI.DateVersion:=EncodeDate(2006,05,31) ;
  V_PGI.ImpMatrix := True ;
  V_PGI.OKOuvert:=FALSE ;
  V_PGI.Halley:=TRUE ;
  V_PGI.MenuCourant:=0 ;
  V_PGI.DebugSQL:=False;
  V_PGI.Debug:=False;
  V_PGI.QRPdf:=True;
  V_PGI.NumMenuPop :=27;
  V_PGI.CegidBureau := True;
  V_PGI.CegidApalatys  := FALSE;
  // pour autoriser l'acces au dp comme dossier de travail,au multi dossier
  V_PGI.StandardSurDP  := true;
  // pour autoriser l'acces au dp comme base de reference pour tous les dossiers ==> Table DP+STD
  V_PGI.MajPredefini := True;
// Confidentialité
  V_PGI.PGIContexte := [CtxPaie];
  V_PGI.RAZForme := TRUE;
// Blocage des mises à jour de structure par l'appli. Elle doit etre faite par PgiMajVEr fourni dans le kit socref
  V_PGI.BlockMAJStruct := TRUE;
  v_pgi.sav := False ;
end;

Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;
end.
