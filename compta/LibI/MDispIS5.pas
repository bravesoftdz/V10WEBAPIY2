unit MDispIS5 ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  UIUtil, ExtCtrls, DBTables, StdCtrls, Mask, DBCtrls, Db, Grids, DBGrids,
  Menus, ComCtrls, HPanel, HStatus, HMenu97, Hctrls, HDB,UtilPGI,
  HTB97,HEnt1,Ed_Tools, HRotLab, hmsgbox, HFLabel, Courrier, MulCour,
  LicUtil,HCalc,UserChg, SQL,HMacro,AddDico, HCapCtrl,IniFiles, Seria,Prefs,
  Buttons,Xuelib,About,Traduc,TradForm,Tradmenu,Graph,
  FE_Main,UIUtil3, ImgList, Hout,
  EdtREtat,EdtEtat,UEdtComp,
  PGIExec, PgiEnv, PGIAppli,
  GrpFavoris, MenuOLG, ParamSoc, Tablette;

Procedure InitApplication ;
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;

type
  TFMDispIS5 = class(TDatamodule)
    procedure FMDispS5Create(Sender: TObject);
    private
      NumModule : integer;
      procedure ExercClick(Sender: TObject);
    public
      procedure AfterChangeGroup ( Sender : TObject) ;
    end ;

Var FMDispIS5 : TFMDispIS5 ;

implementation

uses
{ COMMUN }
  UtilSoc,
{ DP }
  GalEnv,
{ COMPTABILITE}
  Ent1, AssistPL,
{ IMMOBILISATIONS }
  MulImmo, Outils, ImMulHis, SupImmo, IntegEcr, MulInteg, ImOutGen, ImRapCpt,
  ImoClo, ImMigEur, ImPlan, CptesAss, ImSaiCoef, CtrlImmo, RepImmo, ImEdGene,
  ImOle, ImEdCalc, ImEnt
;

{$R *.DFM}

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
begin
  case Num of
    10 :  BEGIN
            UpdateComboslookUp ;
            ChargeHalleyUser ;
            TripoteStatus ;
          END ;
    11 : ; // Après déconnexion
    12 :  BEGIN
            if ((V_PGI_ENV<>Nil) and (V_PGI_ENV.ModePCL='1')) then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
            if ctxPCL in V_PGI.PGIContexte then VH^.ModeSilencieux:=Not IsAppliActivee('CCS5.EXE') ;
            if ctxPCL in V_PGI.PGIContexte then
            BEGIN
              FMenuG.bUser.Hint := TraduireMemoire('Exercice de référence');
              FMenuG.bUser.Onclick := FMDispIS5.ExercClick;
              VK_CHOIXMUL:=VK_F5 ; VK_RECHERCHE:=VK_F6 ;
            END ;
          end;
    13 : ;
    15 : ;
    100 : if ctxPCL in V_PGI.PGIContexte then
          begin
            if not LanceAssistantComptaPCL(True) then begin FMenuG.ForceClose := True; FMenuG.Close; Exit; end;
            RetourForce:=True ;
          end;
    2111 : ConsultationImmo(taModif, toNone,True,'') ;
    2115 : AfficheHistoriqueImmo ;
    2120 : SuppressionImmo ;
    2410 : AfficheMulIntegration;
    2412 : IntegrationEcritures (toDotation,Nil,True,False);
    2414 : IntegrationEcritures (toEcheance,Nil,True,False);
    2415 : EtatRapprochementCompta ;
    2416 : AfficheClotureImmo ;
    2417 : LanceMigrationImmoEuro ;
    2418 : UpdateBaseImmo ;    //ajout TD le 8/6/2001 menu Recalcul des plans
    2210 : ConsultationComptesAsso(taModif) ;
//      2340 : ParamSociete(False,'','SCO_IMMOBILISATIONS','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000) ;
//      2340 : ParamSociete(False,'','SCO_IMMOBILISATIONS','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000) ;
    2340 : begin ParamSociete(False,'','SCO_IMMOBILISATIONS','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000) ; end ; //YCP
    2342 : ParamTable('TILIEUGEO',taCreat,2342000,PRien) ;
    2345 : ParamTable('TIMOTIFCESSION',taCreat,2345000,PRien) ;
    2347 : SaisieDesCoefficientsDegressifs ;
    2370 : ControleFichiersImmo ;
    2375 : RepareFichiersImmo ;
    2505 : LanceEditionImmo ('FIM');
    2511 : LanceEditionImmo ('ACS');
    2512 : LanceEditionImmo ('INV');
    2513 : LanceEditionImmo ('SIM');
    2514 : LanceEditionImmo ('STS');
    2515 : LanceEditionImmo ('MUT');
    2516 : LanceEditionImmo ('PMV');
    2521 : LanceEditionImmo ('DOT');
    2522 : LanceEditionImmo ('DEC');
    2523 : LanceEditionImmo ('DFI');
    2524 : LanceEditionImmo ('DDE');
    2531 : LanceEditionImmo ('PRE');
    2532 : LanceEditionImmo ('PRF');
    2535 : LanceEditionImmo ('DCB');
    2536 : LanceEditionImmo ('ECB');
    2539 : LanceEditionImmo ('PTP');
    2540 : LanceEditionImmo ('ETI');
  else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
  end ;
end;


Procedure AfterProtec ( sAcces : String ) ;
BEGIN
  VH^.OkModImmo:=sAcces='X';
  FMenuG.SetModules([20],[]) ;
END ;

procedure AssignLesZoom ( LesImmos : boolean ) ;
BEGIN
  ProcZoomEdt:=imZoomEdtEtatImmo ;
  ProcCalcEdt:=CalcOLEEtatImmo ;
END ;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
  UpdateSeries ;
  AssignLesZoom(NumModule=20) ;
  Case NumModule of
    9,11,13,14,16,17,52,54,56,96 : ChargeMenuPop(hm2,FMenuG.Dispatch) ;
                12 : ChargeMenuPop(hm3,FMenuG.Dispatch) ;
                20 : ChargeMenuPop(hm4,FMenuG.Dispatch) ;
    else ChargeMenuPop(TTypeMenu(NumModule),FMenuG.Dispatch) ;
  END ;
  if FMenuG.PopupMenu=Nil then FMenuG.PopUpMenu:=ADDMenuPop(Nil,'','') ;
  FMDispIS5.NumModule:=NumModule ;
END ;

procedure TFMDispIS5.AfterChangeGroup(Sender : TObject);
begin
  if ctxPCL in V_PGI.PGIContexte then
  begin
    AssignLesZoom( (NumModule=53) and (fMenuG.Outlook.ActiveGroup = 2) );
    if ((NumModule=53) and (fMenuG.Outlook.ActiveGroup=2)) then
       ChargeMenuPop(hm4,FMenuG.Dispatch)
    else  if ((NumModule=55) and (fMenuG.Outlook.ActiveGroup=2)) then
       ChargeMenuPop(hm3,FMenuG.Dispatch)
    else ChargeMenuPop(hm2,FMenuG.Dispatch) ;
  end;
end;

procedure AssignZoom ;
BEGIN
AssignLesZoom(False) ;
ProcGetVH:=GetVH ;
//ProcGetDate:=GetDate ;
(*if Not Assigned(ProcZoomGene)    then ProcZoomGene    :=FicheGene    ;
if Not Assigned(ProcZoomTiers)   then ProcZoomTiers   :=FicheTiers   ;
if Not Assigned(ProcZoomSection) then ProcZoomSection :=FicheSection ;
if Not Assigned(ProcZoomJal)     then ProcZoomJal     :=FicheJournal ;
if Not Assigned(ProcZoomCorresp) then ProcZoomCorresp :=ZoomCorresp  ;
if Not Assigned(ProcZoomBudGen)  then ProcZoomBudGen  :=FicheBudgene ;
if Not Assigned(ProcZoomBudSec)  then ProcZoomBudSec  :=FicheBudsect ;
if Not Assigned(ProcZoomRub)     then ProcZoomRub     :=FicheRubrique ;
If Not Assigned(ProcZoomNatCpt)  Then ProcZoomNatCpt  :=FicheNatCpte ;*)
END ;

function ChargeFavoris : boolean ;
begin
//  AddGroupFavoris( FMenuG , '54;55;96') ;
  result := true;
end;

procedure InitApplication ;
Var sDom : String ;
BEGIN
  AssignZoom ;
  FMenuG.OnDispatch:=Dispatch ;
  FMenuG.OnChangeModule:=AfterChangeModule ;
  FMenuG.OnChargeMag:=ChargeMagHalley ;
  FMenuG.OnAfterProtec:=AfterProtec ;
//  FMenuG.OnMajAvant:=MajHalleyAvant ;
//  FMenuG.OnMajApres:=MajHalleyApres ;
//  FMenuG.OnMajPendant:=MajHalleyPendant ;
//  FMenuG.OnChargeFavoris := ChargeFavoris;
  sDom:='00280011' ;
  if GetModePCL('')='1' then
  begin
    sDom:='00280011' ;
    VH^.SerProdImmo  :='00170040';
  end
  else
  if GetModePCL('')='1' then
  begin
    sDom:='05990011' ;
    VH^.SerProdImmo  :='05040040' ;
    FMenuG.SetSeria(sDom,[VH^.SerProdImmo],['Immobilisations']) ;
  end;
  FMenuG.Outlook.OnChangeActiveGroup:=FMDispIS5.AfterChangeGroup;
END ;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
(*  case Num of
    1 : FicheGene(Nil,'',LeQuel,Action,0) ;
    2 : FicheTiers(Nil,'',LeQuel,Action,1) ;
    4 : FicheJournal(Nil,'',Lequel,Action,0) ;
    5 : FicheNatCpte(nil,'I0'+AnsiLastChar(TT),'',Action,1) ; //YCP
  end ;*)
end ;

procedure TFMDispIS5.ExercClick(Sender: TObject);
begin
//  ParamSociete(False,'','SCO_DATESDIVERS','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,0) ;
end;

procedure TFMDispIS5.FMDispS5Create(Sender: TObject);
begin
  PGIAppAlone:=True ;
  CreatePGIApp ;
end;

initialization

{Version}
Apalatys:='CEGID' ;
NomHalley:='Immobilisations S5' ;
TitreHalley:='CEGID Immobilisations S5' ;
HalSocIni:='CEGIDPGI.INI' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.NumVersion:='4.0.1' ;
V_PGI.NumBuild:='1' ; V_PGI.NumVersionBase:=590 ;
V_PGI.DateVersion:=EncodeDate(2002,05,22) ;

{Généralités}
V_PGI.VersionDemo:=True ;
V_PGI.SAV:=False ;
V_PGI.VersionReseau:=True ;
V_PGI.PGIContexte:=[ctxCompta] ;
V_PGI.CegidAPalatys:=FALSE ;
V_PGI.CegidBureau:=TRUE ;
V_PGI.StandardSurDP:=True ;
V_PGI.MajPredefini:=False ;
V_PGI.MultiUserLogin:=False ;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=False ;

ChargeXuelib ;

{Série}
V_PGI.LaSerie:=S5 ;
V_PGI.NumMenuPop:=27 ;
V_PGI.OfficeMsg:=True ;
V_PGI.OutLook:=TRUE ;
V_PGI.NoModuleButtons:=FALSE ;
V_PGI.NbColModuleButtons:=1 ;

{Divers}
V_PGI.MenuCourant:=0 ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.VersionDEV:=TRUE ;
V_PGI.ImpMatrix := True ;
V_PGI.DispatchTT:=DispatchTT ;
V_PGI.ParamSocLast:=False ;
V_PGI.Decla100:=TRUE ;
V_PGI.RAZForme:=TRUE ;

//RenseignelaSerie(ExeCCS5) ;

//RegisterHalleyWindow ;
RegisterAmortissement;

end.
