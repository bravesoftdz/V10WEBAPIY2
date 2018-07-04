unit MDispS7 ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  UIUtil, ExtCtrls, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, StdCtrls, Mask, DBCtrls, Db, Grids, DBGrids,
  Menus, ComCtrls, HPanel, HStatus, HMenu97, Hctrls, HDB,UtilPGI,
  HTB97,HEnt1,Ed_Tools, HRotLab, hmsgbox, HFLabel, Courrier, MulCour,
  LicUtil,HCalc,UserChg, SQL,HMacro,AddDico, HCapCtrl,IniFiles, Seria,Prefs,
  Buttons,Xuelib,About,Traduc,TradForm,Tradmenu,Graph,MenuOLG,Corresp_TOM,
  FE_Main,UIUtil3, ImgList, Hout, uTOFEtatsChaines,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,UEdtComp,
{$ENDIF}
  HalOLE,PGIExec, PgiEnv
  ,EntPGI
  ,Galoutil
  ;

Procedure InitApplication ;
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;

type
  TFMDispS7 = class(TDatamodule)
    procedure FMDispS7Create(Sender: TObject);
    private
      NumModule : integer;
    public
    end ;

Var FMDispS7 : TFMDispS7 ;

implementation

uses MulTiers, mulJal, MulBudg, EXERCICE_TOM, QR,
     REGION, QRBrouil, PlanRef, VENTIL, VENTTYPE, repartbu, Structur,
     CODEAFB, RUPTURE, QRPLANGE, QRSectio, Pointage, Saisie, HCompte, Ent1,
     Section, Axe, Budgene, Journal, Tablette, General, Tiers,
     CodePost, Pays, Devise, Tva, MulGene, OuvrFerm,
     MulSecti, MULMVT, QRPlanJo, MULANA, QRBudgen,
     Guide, Paramlib, FichComm, SOUCHE_TOM, //MajHalley,
     QRJDivis, QRCumulG, ParamSoc,CpteUtil,SaisUtil,
     SupprEcr, SupprAna, ValidSim, LetRegul, EncaDeca, MulAbo,
     QRJCpte, REPARCLE, ContaBon, MULLETTR, MULDLETT,
     Rubrique, QRGLANA, QRGLAux, MULLOG, QRTier,
     QRPointg, VideSect, QRRappro, QRBalGen, QRBLANA, QRGLAuGe, QRGLGeAu,
     QRGLGen, QRGLGESE, QRGLSEGE, RappAuto, QRBLSEGE, QRBLGESE,
     Suprgene, BanqueCP, SuprSect, SuprJal, UtilSoc,
     ReconAbo,GenerAbo, Edtlegal,
{$IFDEF V530}
     EdtDoc,
{$ELSE}
     EdtRDoc,
{$ENDIF}
     QRBaAuGe, QRBaGeAu,QRJuSold,
     QRBalAux, QREchean, RelevTie, ValPerio, QRBalAge, QRBalVen, QRGLAge,
     QRGLVen, Scenario, EtapeReg, HDebug,
     SaisieTr, EcheModf, Relance,RelanceNew, CalcScor, RelCompt,
     SplashG, CreerSoc, QRLegal, RAZ, CopySoc, TvaContr,
     CloPerio, Resulexo, QRJustBq, OuvreExo, Revision,
     Suprbudg, ReImput, Extraibq, Cloture, Choix, Extourne,
     AnulClo, CALCOLE, Totrub, CtrlCai, CoutTier, CtrRub, CtrRubBu, SuprAuxi,
     CtrlFic, CtrlDiv, MulRub, SuprRub, Reseau, AssistEF, QRRub,
     AssistSO, QrBroAna, CritEdt, ValJal, SuivTres, (*DefSat,*)
     NatCpte, Audit, Budsect,ModEcheMP,
     Mulbuds, Budjal, Mulbudj, MulMvtBu, QRBudSec, QRBudJal, MulValBu, Suprbudj,
     QRBLteG, QRBLteS, QRBLteGS, QRBLteSG, QrBroBuG, QrBroBuS, Suprbuds, Supecrbu,
     CopieBud, CreRubBu, QRBLecG, SaisBud, QRBLecS, QRBLecGS,
     QRBLecSG, CreCodBu, CreSecBu, PssToTli, Reindex, DocRegl,
     LibTabLi, ChoiTali, CopiTabL, ModTlEcr, ModiTali, CroisCpt, CpyBudMu,
     LibChpLi,QRBLEL,QRBLAL,CFONBAT, CodBuRup, SuivDec, RevBuRea,
     ImputAna, TvaModif,QRTvaHT,QRTvaAcE,QRTvaFac,TVAEdite,BonPayer,TvaAcc,
     QRBlAge2, StrucBud, Csebucat, TvaValid, QRCatSG, CtrBud, Constant,
     RubFam, EdtQR, InitSoc, calceuro, GenToBud, SerLic,
     AnulCAN, ClotANA, ConsEcr,

     SaisBor, SaisBal, ZCentral, MulBds, EtbMce,
     TomProfilUser, TofRubImport, CHGNATUREGENE_TOF,

     MulImmo, MulInteg, IntegEcr, CptesAss, Outils, SupImmo, Ctrlimmo, RepImmo,
     ImedGene, ImCrGuid, ImMulHis, ImoClo, ImEdCalc,ImRapCpt, ImoGen, 
     QRBAGTP, QRBAVTP, QRBRT, QRGLATP, QRGLVTP, QRGLAuxL, ImOutGen, ImPlan,

     CritSynth, IccGlobale, ImSaiCoef, UTOMUTILISAT, USERGRP_TOM,
     RELANCE_TOM, PointageNew, CPETATPOINTAGE_TOF, CPRAPPRO_TOF, CPJUSTBQ_TOF,CPAFAUXCPTA_TOF,
// Prorata de TVA
     Prorata_tof, CPECRPRORATISEES_TOF, CPPRORATA_TOM,
// Longueur des comptes
     LGCOMPTE_TOF,
// Modif entête de pièce
     UTOFMODIFENTPIE,
// Balance générale
     CPBALGEN_TOF, CPBALAUXI_TOF, 
// Balance âgée et ventilée
     TofBalagee
// Sources TOF Etat sur tables libres, ajout SB
     ,CPBALANCETL_TOF, CPGRANDLIVRETL_TOF
// Grand-livre général et auxiliaire
      , uTofCPGLGENE, uTOFCPGLAUXI
// Journal divisionnaire
      , UTOFCPJALECR
// Synchronisation tables libre
     , Synchro
// Export Etafi
     , TofExpEtafi
     ;

{$R *.DFM}

Procedure LanceEtatLibreS7 ( NatureEtat : String ) ;
Var CodeD : String ;
BEGIN
CodeD:=Choisir('Choix d''un état libre','MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="'+NatureEtat+'" AND MO_MENU="X"','', False, false, 999999112, '') ;
if CodeD<>'' then LanceEtat('E',NatureEtat,CodeD,True,False,False,Nil,'','',False) ;
END ;

Function OkRevis : boolean ;
BEGIN
Result:=V_PGI.Controleur ;
if Not Result then PGIBox('Fonction interdite : vous devez  être réviseur pour accéder aux fonctions de révision','Comptabilité S7') ;
END ;

procedure ChangeMenuDeca ;
// Modifie les boutons circuits de deca
var Q      : TQuery ;
    n    : integer ;
    s,StInter,sAuto : string ;
begin
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" ORDER BY CC_CODE',True) ;
n:=0 ;
StInter:='---------------------------------------------------------------------------------------------------' ;
SAuto  :='000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000' ;
While not Q.EOF do
   begin
   Inc(n) ;
   s:=Q.FindField('CC_LIBELLE').AsString ;
   if Trim(s)='' then
      BEGIN
      if Not V_PGI.OutLook then ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+StInter+'" WHERE MN_1=7 AND MN_2=4 AND MN_3=2 AND MN_TAG='+IntToStr(74990+n)) ;
      FMenuG.RemoveItem(74990+n) ;
      END else
      BEGIN
      if Not V_PGI.OutLook then ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+SAuto+'", MN_LIBELLE="'+s+'" WHERE MN_1=7 AND MN_2=4 AND MN_3=2 AND MN_TAG='+IntToStr(74990+n)) ;
      FMenuG.RenameItem(74990+n,s) ;
      END ;
   Q.Next ;
   end ;
Ferme(Q) ;
end ;

Function GetStRevis : String ;
BEGIN
Result:='' ;
if VH^.DateRevision>IDate1900 then Result:='Révision au ' +FormatDateTime('c',VH^.DateRevision) ;
END ;

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
var stMessage : string ;
BEGIN
MajBeforeDispatch(Num) ;
Case Num of
  10 : BEGIN
       UpdateComboslookUp ; ChargeHalleyUser ; TripoteStatus(GetStRevis) ; ChangeMenuDeca ;
       if PCL_IMPORT_BOB('CCS5') = 1 then
             begin
                  stMessage := 'LE SYSTEME A DETERMINE UNE MAJ DE MENU';
                  stMessage := stMessage +#13#10 +'POUR ETRE PRISE EN COMPTE';
                  stMessage := stMessage +#13#10 +'LE LOGICIEL VA SE FERMER AUTOMATIQUEMENT';
                  stMessage := stMessage +#13#10 +'VOUS DEVEZ RELANCER L''APPLICATION';
                  PGIInfo(stMessage+' : '+V_PGI_Env.DBSocName,'MAJ MENU');
                  FMenuG.ForceClose := TRUE;
                  PostMessage(FmenuG.Handle,WM_CLOSE,0,0) ;
                  RetourForce:=true;
                  FermeHalley:=true;
             end;
       END ;
  11 : ;
  12 : ;
  13 : ;
  15 : ;
  16 : ;
  100 : ;
// BOITE A OUTILS
   // Fichiers
   1104 : ParamSociete(False,'','SCO_PARAMETRESGENERAUX','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,1105000) ;
   1105 : FicheAxeAna ('') ;
   1110 : FicheEtablissement_AGL(taModif) ;
      // Tables soc
   1120 : ParamTable('ttFormeJuridique',taCreat,1120000,Nil) ;
   1122 : ParamTable('ttCivilite',taCreat,1122000,Nil) ;
   1125 : ParamTable('ttFonction',taCreat,1125000,Nil) ;
   1135 : OuvrePays ;
   1140 : FicheRegion('','',False) ;
   1145 : OuvrePostal(Nil) ;
   1150 : FicheDevise('',tamodif,False) ;
   1155 : ParamTable('ttLangue',taCreat,1155000,Nil) ;
   1165 : ParamTable('TTREGIMETVA',taCreat,1165000,Nil) ;
   1170 : ParamTvaTpf(true) ;
   1175 : ParamTvaTpf(false) ;
   1185 : FicheModePaie_AGL('');
   1190 : FicheRegle_AGL('',FALSE,taModif) ;
   1191 : ParamTable('ttQualUnitMesure',taCreat,1190030,Nil) ;
   1193 : ParamTable('ttNivCredit',taCreat,1190040,Nil) ;
   1194 : ModifLibelleTableLibre ;
   1195 : ModifLibelleChampLibre('','') ;
   1196 : ChoixRunNatCpte ;
   1197 : CopiTableLibre ;
   1198 : SynchroTaLi;
      // Perso/Traduc
   1205 : ParamTraduc(TRUE,Nil) ;
   1210 : ParamTraduc(FALSE,NIL) ;
   // Documents
      // Compta
   1235 : EditDocumentS5S7('L','REL','',TRUE) ;
   1240 : EditDocumentS5S7('L','RLV','RLV',FALSE) ;
   1245 : EditDocumentS5S7('L','RLC','',True) ;
   1247 : EditEtatS5S7('E','SAI','',True,nil,'','') ;
   1248 : EditEtatS5S7('E','BAP','',True,nil,'','') ;
   1257 : EditEtatS5S7('E','BOR','',True,nil,'','') ;
   // Banque
   1267 : EditDocumentS5S7('L','LCH','',True) ;
   1268 : EditDocumentS5S7('L','LTR','',True) ;
   // Editions paramétrable
   1280 : EditEtatS5S7('E','JDI','',TRUE,nil,'','') ;
   1281 : EditEtatS5S7('E','JAC','',TRUE,nil,'','') ;
   1282 : EditEtatS5S7('E','JPE','',TRUE,nil,'','') ;
   1283 : EditEtatS5S7('E','JAL','',TRUE,nil,'','') ;
   1284 : EditEtatS5S7('E','GLP','',TRUE,nil,'','') ;
   1285 : EditEtatS5S7('E','GLA','',TRUE,nil,'','') ;
   1286 : EditEtatS5S7('E','BGE','',TRUE,nil,'','') ;
   1287 : EditEtatS5S7('E','BAU','',TRUE,nil,'','') ;
   1288 : EditEtatS5S7('E','BVE','',TRUE,nil,'','') ;
   1289 : EditEtatS5S7('E','JTA','',TRUE,nil,'','') ;
   1291 : EditEtatS5S7('E','RES','',TRUE,nil,'','') ;
   1292 : EditEtatS5S7('E','BIL','',TRUE,nil,'','') ;
   1293 : EditEtatS5S7('E','SIG','',TRUE,nil,'','') ;
   // Générateur d'état
   1270 : EditEtatS5S7('E','UCO','',True,nil,'','') ;
   1271 : EditEtatS5S7('E','UBU','',True,nil,'','') ;
   // COMPTABILITE
//   1290 : FicheExercice('',taModif,0) ;
   1290 : YYLanceFiche_Exercice('0');
   1300 : YYLanceFiche_Souche('CPT') ;
//   1305 : YYLanceFiche_Souche('BUD') ;
   1315 : FichePlanRef(1,'',taModif) ;
      // Correspondance
(*
   1325 : ParamCorresp('GE') ;
   1330 : ParamCorresp('AU') ;
//   1335 : ParamCorresp('BU') ;
   1345 : ParamCorresp('A1') ;
   1350 : ParamCorresp('A2') ;
   1355 : ParamCorresp('A3') ;
   1360 : ParamCorresp('A4') ;
   1365 : ParamCorresp('A5') ;
*)
      // VLSB
   1325 : CCLanceFiche_Correspondance('GE');	// Plan de correspondance Généraux
   1330 : CCLanceFiche_Correspondance('AU');	// Plan de correspondance Auxiliaire
   1345 : CCLanceFiche_Correspondance('A1');	// Plan de correspondance Axe analytique 1
   1350 : CCLanceFiche_Correspondance('A2');	// Plan de correspondance Axe analytique 2
   1355 : CCLanceFiche_Correspondance('A3');	// Plan de correspondance Axe analytique 3
   1360 : CCLanceFiche_Correspondance('A4');	// Plan de correspondance Axe analytique 4
   1365 : CCLanceFiche_Correspondance('A5');	// Plan de correspondance Axe analytique 5

      // Plan rupture
   1375 : PlanRupture('RUG') ;
   1380 : PlanRupture('RUT') ;
//   1385 : PlanRupture('RUB') ;
   1395 : PlanRupture('RU1') ;
   1400 : PlanRupture('RU2') ;
   1405 : PlanRupture('RU3') ;
   1410 : PlanRupture('RU4') ;
   1415 : PlanRupture('RU5') ;
      // Divers
   1425 : ParamLibelle ;
   1426 : CCLanceFiche_Correspondance('ZG') ;
   1430 : ParamGuide('','NOR',taModif) ;
   1435 : ParamScenario('','') ;
//   1436 : SaisieTrs('', taModif) ; // Pram Strucutre - Ecriture - Imputation banque
      // Analytique
   1450 : ParamPlanAnal('') ;
   1455 : ParamRepartCle ;
   1460 : ParamVentilType ;
      // Etapes regl
   1485 : ParamEtapeReg(True,'',True) ;
   1490 : ParamEtapeReg(False,'',True) ;
   1491 : CircuitDec ;
      // Budget
// Outils Budgets
  1496 : YYLanceFiche_Souche('BUD') ;
  1498 : ParamStructureBudget ;
  1501 : CodeGeneVersCodeBudg ;
  1502 : CreationCodBudSurRupture(fbGene) ;
  1504 : CreationDesCodesBudget(fbBudGen) ;
  1522 : CreationCodBudSurRupture(fbAxe1) ;
  1524 : CreationDesCodesBudget(fbBudSec1) ;
  1526 : CopiePlanSousSectionToTableLibre ;
  1528 : If Not VH^.DupSectBud
            then PGIBox('Vous n''avez pas paramétré la copie des sections analytiques vers les sections budgétaires dans la fiche société.Vous ne pouvez pas effectuer cette opération!','Comptabilité S7')
            else GenerationSectionBudgetaire ;
  1529 : GenSecBudCategorie ;
  1530 : ParamRepartBUDGET('',False) ;
   // BANQUE
   1705 : FicheBanque_AGL('',taModif,0) ;
   1720 : ParamTable('ttCFONB',taCreat,1705030,Nil) ;
   1725 : ParamCodeAFB ;
   1730 : ParamTable('ttCodeResident',taCreat,1705050,Nil) ;
   1732 : ParamTeletrans ;
// ADMINISTRATEUR
   // Installation
//   3110 : DefSatellites ;
   // Société
   3130 : GestionSociete(Nil,@InitSociete,Nil) ;
   3140 : LanceAssistSO ;
   3145 : RecopieSoc(nil,True) ;
   3155 : RAZSociete ;
   // Utilisateurs
   3165 : FicheUserGrp ;
   3170 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;
   3172 : YYLanceFiche_ProfilUser;
   3180 : ReseauUtilisateurs(False) ;
   3185 : VisuLog ;
   3195 : ReseauUtilisateurs(True) ;
   // Utilitaires
   3205 : CCLanceFiche_LongueurCompte ;
   3210 : ControlFic ;
   3215 : ReparationFic ;
   3220 : ReindexSoc ;
   3221 : CPLanceFiche_RubImport;
   3245 : if AnnuleclotureComptable(TRUE) Then FermeHalley:=TRUE ;
   3246 : if BasculeCompta Then FermeHalley:=TRUE ;
// COMPTABILITE
   // Fichiers
      // Comptes généraux
   7106 : MulticritereCpteGene(taConsult) ;
   7109 : FicheGene(Nil,'','',taCreatEnSerie,0) ;
   7112 : MulticritereCpteGene(taModif) ;
   7115 : MulticritereCpteGene(taModifEnSerie) ;
   7116 : ModifSerieTableLibre(fbGene) ;
   // Modif nature compte généraux
   7117 : CPLanceFiche_ChgNatureGene;
   7118 : SuppressionCpteGene ;
   7124 : OuvreFermeCpte(fbGene,FALSE) ;
   7127 : OuvreFermeCpte(fbGene,TRUE) ;
   7133 : PlanGeneral('', false) ;
      // Comptes auxiliaires
   7139 : MultiCritereTiers(taConsult) ;
   7142 : FicheTiers(Nil,'','',taCreatEnSerie,1) ;
   7145 : MultiCritereTiers(taModif) ;
   7148 : MultiCritereTiers(taModifEnSerie) ;
   7149 : ModifSerieTableLibre(fbAux) ;
   7151 : SuppressionCpteAuxi ;
   7157 : OuvreFermeCpte(fbAux,FALSE) ;
   7160 : OuvreFermeCpte(fbAux,True) ;
   7166 : PlanAuxi('',False) ;
      // Sections
   7172 : MulticritereSection(taConsult)  ;
   7175 : FicheSection(Nil,'A1','',taCreatEnSerie,0) ;
   7178 : MulticritereSection(taModif)  ;
   7181 : MulticritereSection(taModifEnSerie)  ;
   7182 : ModifSerieTableLibre(fbSect) ;
   7184 : SuppressionSection ;
   7190 : OuvreFermeCpte(fbSect,FALSE) ;
   7193 : OuvreFermeCpte(fbSect,TRUE) ;
   7199 : PlanSection('','',FALSE) ;
      // Journaux
   7205 : MulticritereJournal(taConsult) ;
   7208 : FicheJournal(Nil,'','',taCreatEnSerie,0) ;
   7211 : MulticritereJournal(taModif) ;
   7214 : SuppressionJournaux ;
   7220 : OuvreFermeCpte(fbJal,FALSE) ;
   7223 : OuvreFermeCpte(fbJal,TRUE) ;
   7229 : PlanJournal('',FALSE) ;
   // Ecritures
      // Saisie
   7241 : SaisieGuidee('') ;
   7242 : SaisieFolio(taModif) ;
   7244 : MultiCritereMvt(taCreat,'N',False) ;
   7245 : PrepareSaisTresEffet ;
   7247 : PrepareSaisTres ;
      // Ecritures courantes
   7256 : MultiCritereMvt(taConsult,'N',False) ;
   7259 : MultiCritereMvt(taModif,'N',False) ;
   7261 : Centralisateur(taModif, 'N', FALSE) ;
   7260 : ModifSerieTableLibreEcr('E',[tpReel]) ;
   // Modif entete piece
   7264 : CCLanceFiche_ModifEntPie;
   7262 : DetruitEcritures('N',False) ;
   7263 : DetruitEcritures('N',True) ;
   7268 : Brouillard('N') ;
      // Simulation
   7274 : MultiCritereMvt(taConsult,'S',False) ;
   7277 : MultiCritereMvt(taCreat,'S',False) ;
   7280 : MultiCritereMvt(taModif,'S',False) ;
   7278 : Centralisateur(taModif, 'S', FALSE) ;
   7282 : ModifSerieTableLibreEcr('E',[tpSim]) ;
   7281 : ValideEnReel('S') ;
   7283 : DetruitEcritures('S',False) ;
   7289 : Brouillard('S') ;
      // Situation
   7295 : MultiCritereMvt(taConsult,'U',False) ;
   7298 : MultiCritereMvt(taCreat,'U',False) ;
   7301 : MultiCritereMvt(taModif,'U',False) ;
   7302 : Centralisateur(taModif, 'U', FALSE) ;
   7305 : ModifSerieTableLibreEcr('E',[tpSitu]) ;
   7304 : DetruitEcritures('U',False) ;
      // Prévision
   7310 : MultiCritereMvt(taConsult,'P',False) ;
   7313 : MultiCritereMvt(taCreat,'P',False) ;
   7316 : MultiCritereMvt(taModif,'P',False) ;
   7318 : Centralisateur(taModif, 'P', FALSE) ;
   7317 : ModifSerieTableLibreEcr('E',[tpPrev]) ;
   7319 : DetruitEcritures('P',False) ;
      // Abonnements
   7328 : GenereAbonnements(True) ;
   7331 : GenereAbonnements(False) ;
   7337 : ParamAbonnement(TRUE,'',taModif) ;
   7340 : ReconductionAbonnements ;
   7346 : ListeAbonnements ;
      // Guides & scénarios
   7352 : ParamGuide('','NOR',taModif) ;
   7355 : ParamScenario('','') ;
      // Ecr ana
   7364 : MultiCritereAna(taConsult) ;
   7367 : MultiCritereAna(taCreat) ;
   7370 : MultiCritereAna(taModif) ;
   7371 : ModifSerieTableLibreEcr('A',[]) ;
   7373 : DetruitAnalytiques ;
   7379 : VidageInterSections ;
   7381 : ReImputAnalytiques ;
   // Reventil Ana
   7382 :  AGLLanceFiche ('CP' , 'CPREVENTILANA', '', '' , '' ) ;
   7385 : BrouillardAna('N') ;
   // Editions
      // Journaux
   7394 : CPLanceFiche_CPJALECR ; // 26/11/2002 nv journal des ecritures
   7397 : CumulPeriodique(fbJal,neJalCentr) ;
   7403 : CumulPeriodique(fbJal,neJalPer) ;
   7406 : CumulPeriodique(fbJal,neJaG) ;
   7409 : JalCpteGe  ;
      //  GL
   7415 : CPLanceFiche_CPGLGENE('');
   7418 : CPLanceFiche_CPGLAUXI('');
   7421 : GLAnal ;
   7419 : CPLanceFiche_EtatGrandLivreSurTables; // Grand Livre sur table libre
   7427 : GLGenAuxi ;
   7430 : GLAuxiGen ;
   7436 : GLGESE ;
   7439 : GLSEGE ;
      // Balances
//   7445 : BalanceGene ;
// CA - 20/12/2001
   7444 : CPLanceFiche_EtatsChaines(';CP');
   7445 : CPLanceFiche_BalanceGeneral;
   //7448 : BalanceLibreGene ;
   //7449 : BalanceLibreAna ;
   7448 : CPLanceFiche_BalanceAuxiliaire;
   7451 : BLAnal ;
   7429 : CPLanceFiche_EtatBalanceSurTables; // Balance sur table libre
   7457 : BalGeAu ;
   7460 : BalAuGe ;
   7466 : BLGESE ;
   7469 : BLSEGE ;
      // Cumuls
   7478 : CumulPeriodique(fbGene,neJalRien) ;
   7481 : CumulPeriodique(fbAux,neJalRien) ;
   7484 : CumulPeriodique(fbSect,neJalRien) ;
   7490 : VisuEditionLegale(FALSE) ;
   // Edition de synthèse PCL
   7440 : EtatPCL (esCR)  ; // Document de synthèse - Résultat
   7441 : EtatPCL (esBIL) ; // Document de synthèse - Actif
   7442 : EtatPCL (esSIG) ; // Document de synthèse - SIG
   7446 : EtatPCL (esCRA) ; // Document de synthèse analytique - Résultat
   7447 : EtatPCL (esSIGA) ; // Document de synthèse analytique - SIG
   7449 : EtatPCL (esBILA) ; // Document de synthèse analytique - Bilan
   // EDITIONS TIERS PAYEUR
   7231 : AGLLanceFiche('CP', 'EPTIERSP', '', 'TFP', 'TPF') ; // Tiers facturé par payeur
   7232 : AGLLanceFiche('CP', 'EPTIERSP', '', 'TFA', 'TPF') ; // Tiers facturé avec payeur
   7233 : BrouillardTP ;
   7234 : BalanceAge2TP ;
   7236 : GLivreAgeTP ;
   7235 : BalVentileTP ;
   7237 : GLVentileTP ;
  // Editions "Sisco Like" Menu S7 non mis à jour (en attente JLD)
      7410 : AGLLanceFiche('CP', 'EPJALDIV', '', 'JDP', 'JDI') ; // Journal divisionnaire
      7411 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JCJ', 'JAC') ; // Journal centralisateur par journal
      7412 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JCD', 'JAC') ; // Journal centralisateur par période-journal
      7413 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JPP', 'JPE') ; // Journal périodique par Journal-Période
      7414 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JPJ', 'JPE') ; // Journal périodique par Période-Journal
      7422 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JPN', 'JPE') ; // Journal périodique par Période-Journal-Nature
      7423 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JGE', 'JAL') ; // Journal général
      7424 : AGLLanceFiche('CP', 'EPGLIVRE', '', 'GLP', 'GLP') ; // Grand livre général
      7425 : AGLLanceFiche('CP', 'EPGLIVRE', '', 'GLA', 'GLA') ; // Grand livre auxiliaire
      7426 : AGLLanceFiche('CP', 'EPBGLGEN', '', 'BGE', 'BGE') ; // Balance générale
      7450 : AGLLanceFiche('CP', 'EPBGLGEN', '', 'BAU', 'BAU') ; // Balance générale auxiliaire
      7428 : AGLLanceFiche('CP', 'EPBGLGEN', '', 'BVE', 'BVE') ; // Balance générale auxiliaire
//      7431 : AGLLanceFiche('CP', 'EPBALBDS', '', 'BDS', 'BDS') ; // Balance générale comparative
//      7432 : AGLLanceFiche('CP', 'EPETASYN', '', 'RC1', 'RLI') ; // Document de synthèse - Charge
//      7433 : AGLLanceFiche('CP', 'EPETASYN', '', 'RP1', 'RLI') ; // Document de synthèse - Produit
      7434 : if ctxPCL in V_PGI.PGIContexte then AGLLanceFiche('CP', 'EPETASYN', '', 'RCX', 'RES')  // Document de synthèse - Résultat
                                            else AGLLanceFiche('CP', 'EPETASYN', '', 'RCP', 'RES') ; // Document de synthèse - Résultat
      7435 : if ctxPCL in V_PGI.PGIContexte then AGLLanceFiche('CP', 'EPETASYN', '', 'BA1', 'BIL')  // Document de synthèse - Actif
                                            else AGLLanceFiche('CP', 'EPETASYN', '', 'BAF', 'BIL') ; // Document de synthèse - Actif
//      7452 : AGLLanceFiche('CP', 'EPETASYN', '', 'BPF', 'RLI') ; // Document de synthèse - Passif
      7453 : if ctxPCL in V_PGI.PGIContexte then AGLLanceFiche('CP', 'EPETASYN', '', 'SIX', 'SIG')  // Document de synthèse - SIG
                                            else AGLLanceFiche('CP', 'EPETASYN', '', 'SIG', 'SIG') ; // Document de synthèse - SIG
      7454 : AGLLanceFiche('CP', 'EPTABMVT', '', 'JTL', 'JTA') ; // Tableau d'avancement - Mouvements
      7455 : AGLLanceFiche('CP', 'EPTABMVT', '', 'JTM', 'JTA') ; // Tableau d'avancement - Montants
      7456 : LanceEtatLibreS7('UCO') ;
   // Tiers
      // Enca
   7496       : EncaisseDecaisse(True,'','',True,False,tslAucun) ;
   7497       : EncDecChequeTraiteBOR(tslTraite) ;
      // Deca
   7499,74999 : EncaisseDecaisse(False,'','',True,True,tslAucun) ;
   7501       : EncDecChequeTraiteBOR(tslCheque) ;
   7502       : EncDecChequeTraiteBOR(tslBOR) ;
   74997      : LanceBonAPayer ;
   74991      : DecaisseCircuit(1) ;
   74992      : DecaisseCircuit(2) ;
   74993      : DecaisseCircuit(3) ;
   74994      : DecaisseCircuit(4) ;
   74995      : DecaisseCircuit(5) ;
      // Lettrage
   7508 : LanceLettrage ;
   7511 : LanceDeLettrage ;
   7514 : RapprochementAuto('','') ;
   7520 : RegulLettrage(True,False) ;
   7523 : RegulLettrage(False,False) ;
   7524 : RegulLettrage(False,True) ;
   7529 : ModifEcheances ;
      // Editions regl
   7535 : JustSolde ;
   7541 : Echeancier ;
   7547 : CPLanceFiche_BalanceAgee;
   7550 : GLivreAge ;
   7556 : CPLanceFiche_BalanceVentilee;
   7559 : GLVentile ;
   7560 : GLAuxiliaireL ;
      // Relances, rlv, rlc
   (*
   7565 : RelanceClient(True,True) ;
   7566 : RelanceClient(True,False) ;
   7567 : RelanceClient(False,True) ;
   7568 : RelanceClient(False,False) ;
   *)
   7565 : If Not EstSpecif('51197') Then RelanceClientNew(True,True) Else RelanceClient(True,True) ;
   7566 : If Not EstSpecif('51197') Then RelanceClientNew(True,False) Else RelanceClient(True,False) ;
   7567 : If Not EstSpecif('51197') Then RelanceClientNew(False,True) Else RelanceClient(False,True) ;
   7568 : If Not EstSpecif('51197') Then RelanceClientNew(False,False) Else RelanceClient(False,False) ;
//   7574 : ParamRelance('RTR','',taModif) ;
//   7577 : ParamRelance('RRG','',taModif) ;
   7574 : CCLanceFiche_ParamRelance('RTR');
   7577 : CCLanceFiche_ParamRelance('RRG');
   7583 : CalculScoring ;
   7589 : ReleveCompte ;
   7592 : ReleveFacture ;
   7598 : LanceCoutTiers('',0) ;
   7599 : ExportCFONBBatch(False,True,tslAucun) ;
//   7585 : ModifLesRib ;
   7585 : ModifEcheMP ;
   7586 : ExportCFONBBatch(True,True,tslAucun) ;
   7594 : AGLLanceFiche('CP', 'RLVEMISSION', '', '', '') ;
   7595 : ExportCFONBBatch(True,False,tslAucun) ;
   7596 : ExportCFONBBatch(False,False,tslTraite) ;
   7597 : ExportCFONBBatch(False,False,tslBOR) ;
   // Contrôles
      // Pointage
   7602 : OperationsSurComptes('','','') ;
   7604 : If Not EstSpecif('51191') Then PointageCompteNew(TRUE) Else PointageCompte(true) ;
   7607 : If Not EstSpecif('51191') Then PointageCompteNew(FALSE) Else PointageCompte(false) ;
   7616 : If EstSpecif('51190') Then EtatPointage Else CC_LanceFicheEtatPointage ;
   7619 : If EstSpecif('51190') Then EtatRappro Else CC_LanceFicheEtatRappro ;
   7622 : If EstSpecif('51190') Then JustifPointage Else CC_LanceFicheJustifPointage ;
   7625 : SuiviTresoParPeriode ;
   7631 : ControlCaisse(fbGene,tzGCaisse,'') ;
      // TVA
   7640 : ControleTVAFactures('',0,0,'') ;
   7646 : TvaModifEnc ;
   7648 : TvaModifAcc ;
   7647 : EditionTvaHT ;
   7650 : EditionTvaFac(True) ;
   7651 : EditionTvaFac(False) ;
   7652 : EditionTvaAcE ;
   7653 : TVAEditeEtat ;
   7654 : TvaValidEnc ;
   // Prorata TVA
   7655 : CCLanceFiche_ParamTauxTVA;
   7656 : CCLanceFiche_EcrProrataTVA('First;');
   7657 : CCLanceFiche_EcrProrataTVA('Last;');
   7659 : CCLanceFiche_EcrTVAProratisees;

   // Saisie balance
   7710 : SaisieBalance('N1A', taModif) ;
   7711 : SaisieBalance('N1C', taModif) ;
   7716 : SaisieBalance('N0A', taModif) ;
   7713 : SaisieBalance('BDS', taCreat) ;
   7714 : MultiCritereBds(taModif) ;
   7715 : SaisieBalance('BSA', taCreat) ;
   6369 : CP_LanceFicheExpEtafi;
   // Situation PCL
   7717 : AGLLanceFiche ('CP','BALSIT_MUL','','','');
   7719 : TLAuxversTLEcr ;

   // Saisie relevé
//   7770 : SaisieRbp(taModif) ;  // Saisie relevé
//   7771 :  ;
   7772 : AGLLanceFiche('CP', 'RLVRECEPTION', '', '', '') ;
   7773 : AGLLanceFiche('CP', 'RLVINTEGRE', '', '', '') ;
   7774 : AGLLanceFiche('CP', 'RLVPOINTAGE', '', '', '') ;

   // Révision
      // Ecr rev
   7649 : if OkRevis then MultiCritereMvt(taCreat,'R',False) ;
   7658 : if OkRevis then MultiCritereMvt(taConsult,'R',False) ;
   7661 : if OkRevis then MultiCritereMvt(taModif,'R',False) ;
   7662 : if OkRevis then ValideEnReel('R') ;
   7664 : if OkRevis then DetruitEcritures('R',False) ;
   7670 : if OkRevis then Brouillard('R') ;
      // Reimput & mode rev
   7676 : if OkRevis then ReImputation ;
   7682 : if OkRevis then begin LanceRevision ; {JLD AfficheRevision ;} end ;
   7683 : if OkRevis then ResultatDeLexo ;
   7677 : if OkRevis then LanceExtourne ;
   // Clôture
      // Validation
   7691 : ValidationPeriode(TRUE,FALSE) ;
   7694 : ValidationPeriode(FALSE,FALSE) ;
   7700 : ValidationJournal(TRUE) ;
   7703 : ValidationJournal(FALSE) ;
      // Editions
   7709 : LanceEdtLegal ;
   7712 : VisuEditionLegale(TRUE) ;
   7718 : OuvertureExo ;
      // A nouveaux
   7724 : MultiCritereMvt(taCreat,'N',True) ;
   7727 : MultiCritereMvt(taConsult,'N',True) ;
   7728 : MultiCritereAnaANouveau(taCreat);
   7729 : MultiCritereAnaANouveau(taConsult);
   7730 : MultiCritereAnaANouveau(taModif);
      // Cloture
   7736 : CloPer ;
   7737 : AnnuleCloPer ;
   7742 : ClotureComptable(FALSE) ;
   7745 : AnnuleclotureComptable(FALSE) ;
   7751 : If LanceAuditPourCloture Then FermeHalley:=TRUE ;
   7752 : ClotureComptableAna ;
   7753 : AnnuleclotureAnalytique ;
      // Editions de clôture
   7757 : JalDivisioAno ;
   7760 : BalanceAno ;
   7766 : JalDivisioClo ;
   7769 : BalanceClo ;
   // Reporting
   7775 : FamillesRub ;// ParamTable('TTRUBFAMILLE',taCreat,7775000,NIL,17);
   7781 : MultiCritereRubriqueV2(taConsult,False) ;
   7784 : ParametrageRubrique('',taCreatEnSerie,FALSE) ;
   7787 : MultiCritereRubriqueV2(taModif,False) ;
   7790 : SuppressionRubriqueV2(False) ;
   7796 : ImpRubrique('',False,False) ;
   7802 : ControleRubrique ;
   7805 : TotalRubriques ;
   7811 : Constantes ;
   7812 : AGLLanceFiche('CP', 'MULBUDRUB', '', '', '') ;
   7813 : AGLLanceFiche('CP', 'MULRUPRUB', '', '', '') ;
   7814 : AGLLanceFiche('CP', 'MULCORRUB', '', '', '') ;
   7815 : AGLLanceFiche('CP', 'MULTABRUB', '', '', '') ;
// Immobilisations
   2110 : ConsultationImmo(taConsult, toNone,True,'') ;
   2111 : ConsultationImmo(taModif, toNone,True,'') ;
   2101 : FicheImmobilisation(nil,'',taCreat,'PRO') ;
   2102 : FicheImmobilisation(nil,'',taCreat,'FI') ;
   2103 : FicheImmobilisation(nil,'',taCreat,'CB') ;
   2104 : FicheImmobilisation(nil,'',taCreat,'LOC') ;
   2115 : AfficheHistoriqueImmo ;
   2120 : SuppressionImmo ;
   2410 : AfficheMulIntegration;
   2412 : IntegrationEcritures (toDotation,Nil,True,False) ;
   2414 : IntegrationEcritures (toEcheance,Nil,True,False) ;
   2415 : EtatRapprochementCompta ;
   2416 : AfficheClotureImmo ;
   2418 : UpdateBaseImmo ;    //ajout TD le 8/6/2001 menu Recalcul des plans
   2210 : ConsultationComptesAsso(taModif) ;
//   2340 : ParamSociete(False,'','SCO_IMMOBILISATIONS','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000) ;
   2340 : begin ParamSociete(False,'','SCO_IMMOBILISATIONS','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000) ; end ; //YCP
   2342 : ParamTable('TILIEUGEO',taCreat,2342000,Nil) ;
   2345 : ParamTable('TIMOTIFCESSION',taCreat,2345000,Nil) ;
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
   //Menu Pop Compta
   25710 : MulticritereCpteGene(taConsult) ;
   25720 : MultiCritereTiers(taConsult) ;
   25730 : MulticritereSection(taConsult)  ;
   25740 : MulticritereJournal(taConsult) ;
   25750 : MultiCritereMvt(taConsult,'N',False) ;
   25755 : OperationsSurComptes('','','') ;
   25760 : ResultatDeLexo ;
   25770 : ParamSociete(False,'','SCO_PARAMETRESGENERAUX','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,1105000) ;
   //Menu Pop Budget
   26160 : MulticritereBudgene(taConsult) ;
   26165 : MulticritereBudsect(taConsult) ;
   26170 : MulticritereBudjal(taConsult) ;
   26175 : MultiCritereMvtBud(taConsult,'') ;
   26180 : MulticritereCpteGene(taConsult) ;
   26185 : MulticritereSection(taConsult)  ;
   26190 : ParamSociete(False,'','SCO_PARAMETRESGENERAUX','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,1105000) ;
   //Menu Pop Immos
   26510 : ConsultationImmo(taConsult, toNone,True,'') ;
   26520 : MulticritereCpteGene(taConsult) ;
   26530 : MultiCritereTiers(taConsult) ;
   26540 : MulticritereSection(taConsult)  ;
// BUDGET
// Comptes Budgetaires généraux
  15111 : MulticritereBudgene(taConsult) ;
  15113 : FicheBudgene(Nil,'','',taCreatEnSerie,0) ;
  15115 : MulticritereBudgene(taModif) ;
  15117 : MulticritereBudgene(taModifEnSerie) ;
  15118 : ModifSerieTableLibre(fbBudGen) ;
  15119 : SuppressionCpteBudG ;
  15121 : OuvreFermeCpte(fbBudgen,False) ;
  15123 : OuvreFermeCpte(fbBudgen,True) ;
  15125 : PlanBudget('',FALSE) ;
  15126 : ControleBudget(fbBudGen) ;
//Sections Budgetaires
  15131 : MulticritereBudsect(taConsult) ;
  15133 : FicheBudsect(Nil,'','',taCreatEnSerie,0) ;
  15135 : MulticritereBudsect(taModif) ;
  15137 : MulticritereBudsect(taModifEnSerie) ;
  15138 : ModifSerieTableLibre(fbBudSec1) ;
  15139 : SuppressionCpteBudS ;
  15141 : OuvreFermeCpte(fbBudSec1,False) ;
  15143 : OuvreFermeCpte(fbBudSec1,True) ;
  15145 : PlanBudSec('','',FALSE) ;
  15146 : ControleBudget(fbAxe1) ;
//Journaux Budgétaire
  15151 : MulticritereBudjal(taConsult) ;
  15153 : FicheBudjal(Nil,'','',taCreatEnSerie,0) ;
  15155 : MulticritereBudjal(taModif) ;
  15157 : MulticritereBudjal(taModifEnSerie) ;
  15159 : SuppressionJournauxBud ;
  15161 : OuvreFermeCpte(fbBudJal,False) ;
  15163 : OuvreFermeCpte(fbBudJal,True) ;
  15165 : PlanBudJal('','',False) ;

  15211 : MultiCritereMvtBud(taCreat,'G') ;
  15213 : MultiCritereMvtBud(taCreat,'S') ;
  15217 : MultiCritereMvtBud(taCreat,'GS') ;
  15219 : MultiCritereMvtBud(taCreat,'SG') ;

  15221 : MultiCritereMvtBud(taConsult,'') ;
  15225 : VisuConsoBudget('G','') ;
  15227 : VisuConsoBudget('S','') ;

  15230 : MultiCritereMvtBud(taModif,'') ;
  15240 : DetruitBudgets ;
  15250 : MultiCritereValBud(True) ;
  15260 : MultiCritereValBud(False) ;
  15270 : RecopieBudgetSimple ;
  15271 : RecopieBudgetMultiple ;
  15272 : RevisionBudgetaireRealiser ;
//  15280 : TableauBudget('INI',taCreat) ;
  15282 : BrouillardBud('N') ;
  15284 : BrouillardBudAna('N') ;

// Editions budgétaire
  15321 : BalBudteGen ;
  15322 : BalBudteSec ;
  15323 : BalBudteGenSec ;
  15324 : BalBudteSecGen ;

  15331 : BalBudecGen ;
  15332 : BalBudecSec ;
  15333 : BalBudecGenSec ;
  15334 : BalBudecSecGen ;
  15340 : EcartParCategorie ;
  15350 : LanceEtatLibreS7('UBU') ;
// Reporting Budgétaire
  15411 : CreationRubrique ;
  15412 : ParamCroiseCompte ;
  15415 : MultiCritereRubriqueV2(taConsult,True) ;
  15417 : ParametrageRubrique('',taCreatEnSerie,True) ;
  15419 : MultiCritereRubriqueV2(taModif,True) ;
  15421 : SuppressionRubriqueV2(True) ;
  15425 : ImpRubrique('',False,True) ;
  15430 : ControleRubriqueBud ;
  15440 : TotalRubriquesV2(True) ;
// POSTE BANQUE
  16110 : FicheBanque_AGL('',taModif,0) ;
  16120 : FicheBanqueCP('',taModif,0) ;
// Expert
  31001 : SaisieFolio(taModif) ;
  31002 : SaisieBalance('N1A',taModif) ;
  31003 : SaisieBalance('N1C',taModif) ;
  31004 : SaisieBalance('N0A',taModif) ;
  31005 : SaisieBalance('BDS',taCreat) ;
  31010 : SaisieBalance('BSA',taCreat) ;
  31015 : MultiCritereBds(taModif) ;

     // Analyses Cubes, TobViewer, Croisements Axes
  8205  : AGLLanceFiche('CP','CPGENERECROISEAXE','','','') ;
  8210  : AGLLanceFiche('CP','CPCROISEAXE_CUBE','','','') ;
  8215  : AGLLanceFiche('CP','CPCROISEAXE_STAT','','','') ;
  8220  : AGLLanceFiche('CP','CPCROISEAXE_ETAT','','','') ;
  8305  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','') ;
  8307  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','CONTREPARTIE') ;
  8310  : AGLLanceFiche('CP','CPANALYTIQ_CUBE','','','') ;
  8405  : AGLLanceFiche('CP','CPECRITURE_TOBV','','','') ;
  8410  : AGLLanceFiche('CP','CPANALYTIQ_TOBV','','','') ;
  8312  : AGLLanceFiche('CP','CPBALAGEE_CUBE','','','') ;

   // Anciennes éditions
   52581 : BalanceGeneComp;
   52582 : BalanceAuxi;
   52583 : GLGeneral;
   52584 : GLAuxiliaire;
   52585 : JalDivisio;
   52701 : BalanceAge;
   52702 : BalVentile;
   // PGE
   52587 : EtatPointage ;
   52588 : EtatRappro ;
   52589 : JustifPointage ;
  // Icc
  96110 : AGLLanceFiche('CP', 'ICCGENERAUX', '', '', 'ACTION=CONSULTATION');
  96130 : AGLLanceFiche('CP', 'ICCFICHEGENERAUX', '', GetParamSoc('SO_ICCCOMPTECAPITAL'), 'ACTION=MODIFICATION');
  96140 : AGLLanceFiche('CP', 'ICCMULELTNAT', '', '','');
  // Fin Icc

   else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
  END ;
END ;

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
{$IFDEF CEGID}
VH^.OkModCompta:=True ; VH^.OkModBudget:=True ; VH^.OkModImmo:=True ; VH^.OkModEtebac:=True ;
V_PGI.VersionDemo:=False ;
FMenuG.SetModules([7,2,15,8,1,3],[108,1,22,11,28,49]) ;
{$ELSE}
{$IFDEF SANSSERIA}
VH^.OkModCompta:=True ; VH^.OkModBudget:=True ; VH^.OkModImmo:=True ; VH^.OkModEtebac:=True ;
V_PGI.VersionDemo:=False ;
FMenuG.SetModules([7,2,15,8,1,3],[108,1,22,11,28,49]) ;
{$ELSE}
VH^.OkModCompta:=(sAcces[1]='X') ; VH^.OkModBudget:=(sAcces[2]='X') ;
{$IFDEF SANSIMMO}
VH^.OkModImmo:=FALSE ;
//VH^.OkModEtebac:=(sAcces[3]='X') ;
VH^.OkModIcc:=FALSE;
{$ELSE}
VH^.OkModImmo:=(sAcces[3]='X') ;
//VH^.OkModEtebac:=(sAcces[4]='X') ;
VH^.OkModIcc:=FALSE;
{$ENDIF}
If Not VH^.OldTeleTrans Then VH^.OkModEtebac:=TRUE ;

// Si version démo --> Tout accessible mais limité
if V_PGI.VersionDemo then
  BEGIN
{$IFDEF SANSIMMO}
  FMenuG.SetModules([7,15,8,1,3],[108,22,11,28,49]) ;
{$ELSE}
  FMenuG.SetModules([7,2,15,8,1,3],[108,1,22,11,28,49]) ;
{$ENDIF}
  END else
  BEGIN
   // Si pas module compta --> pas les autres modules, seulement BAO
   if Not VH^.OkModCompta then FMenuG.SetModules([1,3],[28,49]) else
   // Nécessairement compta + éventuels autres modules
      BEGIN
      if VH^.OkModBudget then
         BEGIN
         if VH^.OkModImmo then FMenuG.SetModules([7,2,15,8,1,3],[108,1,22,11,28,49])
                          else FMenuG.SetModules([7,15,8,1,3],[108,22,11,28,49]) ;
         END else
         BEGIN
         if VH^.OkModImmo then FMenuG.SetModules([7,2,8,1,3],[108,1,11,28,49])
                          else FMenuG.SetModules([7,8,1,3],[108,11,28,49]) ;
         END ;
      END ;
  END ;
{$ENDIF}
{$ENDIF}
END ;

procedure AssignLesZoom ( LesImmos : boolean ) ;
BEGIN
if LesImmos then
   BEGIN
   ProcZoomEdt:=ZoomEdtEtatImmo ;
   ProcCalcEdt:=CalcOLEEtatImmo ;
   END else
   BEGIN
   ProcZoomEdt:=ZoomEdtEtat ;
   ProcCalcEdt:=CalcOLEEtat ;
   END ;
END ;

procedure ChangeMenuAnalyses ;
BEGIN
{$IFNDEF CEGID}
//FMenuG.RemoveGroup(2) ; {Virer le croisement des axes si pas CEGID}
{$ENDIF}
END ;

procedure RemoveGroupTN(i : Integer ; b : Boolean) ;
Var TN : TTreeNode ;
BEGIN
FMenuG.OutLook.RemoveGroup(i,b) ; TN:=GetTreeItem(FMenuG.TreeView,i) ; If TN<>NIL Then TN.Delete ;
END ;

Procedure RemoveItemTN(i,mn1,mn2,mn3,mn4 : Integer ; OkDel : Boolean) ;
Var HOI : THOutItem ;
    TN : TTreeNode ;
    StInter : String ;
BEGIN
If OkDel Then StInter:='---------------------------------------------------------------------------------------------------'
         Else StInter :='000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000' ;
If Not V_PGI.OutLook Then
  BEGIN
  ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+StInter+'" WHERE MN_1='+IntToStr(Mn1)+' AND MN_2='+IntToStr(Mn2)+' AND MN_3='+IntToStr(Mn3)+' AND MN_4='+IntToStr(Mn4)+'') ;
  If OkDel Then FMenuG.removeItem(i) ;
  END ;
If Not OkDel Then Exit ;
HOI:=FMenuG.OutLook.GetItemByTag(i) ; If HOI<>NIL Then FMenuG.OutLook.RemoveItem(HOI) ;
TN:=GetTreeItem(FMenuG.TreeView,i) ; If TN<>NIL Then TN.Delete ;
END ;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
UpdateSeries ;
AssignLesZoom(NumModule=2) ;
Case NumModule of
  7 : ChargeMenuPop(integer(hm7),FMenuG.DispatchX) ;
  15 : ChargeMenuPop(integer(hm15),FMenuG.DispatchX) ;
  2 : ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ;
  else ChargeMenuPop(integer(hm7),FMenuG.DispatchX) ;
END ;
if NumModule=7 then
  BEGIN
  ChangeMenuDeca ;
  //RR Le 02032003
  //If Not VH^.OldTeleTrans Then RemoveItemTN(7772,7,5,20,1,TRUE) Else RemoveItemTN(7772,7,5,20,1,FALSE) ; // Réception Etebac 3
  RemoveItemTN( 7772,7,5,21,1,TRUE) ; // Réception Etebac 3
  (*
  If Not EstSpecif('51198') Then RemoveItemTN(7264,7,2,7,4,TRUE) Else RemoveItemTN(7264,7,2,7,4,FALSE) ; // Modif entete pièces
  *)
(* passe en Standard !!!
  If Not EstSpecif('51196') Then RemoveItemTN(7719,7,5,24,0,TRUE) Else RemoveItemTN(7719,7,5,24,0,FALSE) ; // Affectation Interco
*)
  END ;
if NumModule=8 then ChangeMenuAnalyses ;
if FMenuG.PopupMenu=Nil then FMenuG.PopUpMenu:=ADDMenuPop(Nil,'','') ;
FMDispS7.NumModule:=NumModule ;
END ;

procedure AssignZoom ;
BEGIN
AssignLesZoom(False) ;
ProcGetVH:=GetVH ;
ProcGetDate:=GetDate ;
if Not Assigned(ProcZoomGene)    then ProcZoomGene    :=FicheGene    ;
if Not Assigned(ProcZoomTiers)   then ProcZoomTiers   :=FicheTiers   ;
if Not Assigned(ProcZoomSection) then ProcZoomSection :=FicheSection ;
if Not Assigned(ProcZoomJal)     then ProcZoomJal     :=FicheJournal ;
if Not Assigned(ProcZoomCorresp) then ProcZoomCorresp :=ZoomCorresp  ;
if Not Assigned(ProcZoomBudGen)  then ProcZoomBudGen  :=FicheBudgene ;
if Not Assigned(ProcZoomBudSec)  then ProcZoomBudSec  :=FicheBudsect ;
if Not Assigned(ProcZoomRub)     then ProcZoomRub     :=FicheRubrique ;
If Not Assigned(ProcZoomNatCpt)  Then ProcZoomNatCpt  :=FicheNatCpte ;
END ;

procedure InitApplication ;
Var sDom : String ;
BEGIN
AssignZoom ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChangeModule:=AfterChangeModule ;
FMenuG.OnChargeMag:=ChargeMagHalley ;
FMenuG.OnAfterProtec:=AfterProtec ;
FMenuG.OnMajAvant:=nil ;
FMenuG.OnMajApres:=nil ;
FMenuG.OnMajPendant:=nil ;
sDom:='07990011' ;
VH^.SerProdCompta:='07010042' ; VH^.SerProdBudget:='07020042' ;
VH^.SerProdImmo  :='07040042' ; VH^.SerProdEtebac:='07030011' ;
{$IFDEF SANSIMMO}
FMenuG.SetSeria(sDom,[VH^.SerProdCompta,VH^.SerProdBudget],
                     ['Comptabilité','Budget']) ;
{$ELSE}
FMenuG.SetSeria(sDom,[VH^.SerProdCompta,VH^.SerProdBudget,VH^.SerProdImmo],
                     ['Comptabilité','Budget','Immobilisations']) ;
{$ENDIF}
FMenuG.SetModules([1],[]) ;
FMenuG.SetPreferences(['Saisies'],False) ;
{$IFDEF CEGID}
FMenuG.bSeria.Visible:=False ;
{$ENDIF}
{$IFDEF SANSSERIA}
FMenuG.bSeria.Visible:=False ;
{$ENDIF}
END ;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel, TT,Range : String) ;
begin
  case Num of
    1 : FicheGene(Nil,'',LeQuel,Action,0) ;
    2 : FicheTiers(Nil,'',LeQuel,Action,1) ;
    4 : FicheJournal(Nil,'',Lequel,Action,0) ;
    5 : FicheNatCpte(nil,'I0'+AnsiLastChar(TT),'',Action,1) ; //YCP
  end ;
end ;

procedure TFMDispS7.FMDispS7Create(Sender: TObject);
begin
PGIAppAlone:=True ;
CreatePGIApp ;
end;

Procedure InitLaVariablePGI;
Begin
  {Version}
  Apalatys:='CEGID' ;
  NomHalley:='Comptabilité S7' ;
  TitreHalley:='CEGID Comptabilité S7' ;
  HalSocIni:='CEGIDPGI.INI' ;
  Copyright:='© Copyright ' + Apalatys ;
  V_PGI.NumVersion:='4.20' ;  V_PGI.NumVersionBase:=595 ;
  V_PGI.NumBuild:='8' ; V_PGI.DateVersion:=EncodeDate(2002,12,16);
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
  V_PGI.EuroCertifiee:=True ;

  ChargeXuelib ;

  {Série}
  V_PGI.LaSerie:=S7 ;
  V_PGI.NumMenuPop:=25 ;
  V_PGI.OfficeMsg:=False ;
  V_PGI.OutLook:=False ;
  V_PGI.NoModuleButtons:=False ;
  V_PGI.NbColModuleButtons:=1 ;
  {Divers}
  V_PGI.MenuCourant:=0 ;
  V_PGI.OKOuvert:=FALSE ;
  V_PGI.Halley:=TRUE ;
  V_PGI.VersionDEV:=TRUE ;
  V_PGI.ImpMatrix := True ;
  V_PGI.DispatchTT:=DispatchTT ;
  V_PGI.ParamSocLast:=False ;
  V_PGI.RAZForme:=TRUE ;

  RenseignelaSerie(ExeCCS7) ;

  RegisterHalleyWindow ;
end;

Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;

end.

