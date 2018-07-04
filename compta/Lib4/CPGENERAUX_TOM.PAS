{***********UNITE*************************************************
Auteur  ...... : COMPTA
Créé le ...... : 12/02/2003
Modifié le ... : 26/03/2007
Description .. : Source TOM de la TABLE : GENERAUX (GENERAUX)
Suite ........ :
Suite ........ : GCO - 03/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheGene en 2/3 et CWAS
Suite ........ : 
Suite ........ : CA - 12/08/2004
Suite ........ : -> FQ 13071 : lien avec GENERAUXREF en contexte PCL
Suite ........ : 
Suite ........ : CA - 17/08/2004
Suite ........ : -> FQ 12543 : Appel fiche RIB en cours de création
Suite ........ : 
Suite ........ : SG6 - 29/12/2004
Suite ........ : -> Gestion du mode analytique multi axe
Suite ........ : 
Suite ........ : GCO - 24/08/2005 - FQ 16423
Suite ........ : 
Suite ........ : Lek FQ16663, A la validation d'un compte Géné de type
Suite ........ : banque, je vérifie si l'enregistrement banque n'existe pas
Suite ........ : j'ouvre la fiche banque pour que l'utilisateur remplisse, sinon 
Suite ........ : je ne l'ouvre pas.
Suite ........ : 
Suite ........ : JP 18/01/06 : FQ 17201 : gestion de l'état en eAgl
Suite ........ : 
Suite ........ : JP 05/06/06 : FQ 17971 : Modification de la gestion du 
Suite ........ : bouton BCopier et du LookUp de duplication
Suite ........ : mbo 17/10/06 - en suppression cpte : ajout vérif sur compte
Suite ........ : de subvention immo
Suite ........ : CA - 13/10/2006
Suite ........ : -> FQ 18925 : correction plantage sous ORCL en mise à
Suite ........ : jour des ventilations analytiques
Suite ........ : CA 21/11/2006
Suite ........ : -> FA 12063 : Interdire modification de la nature si le compte 
Suite ........ :  est mouvementé. 
Suite ........ : SBO 26/03/2007 
Suite ........ : -> DEV 5942 : Affectation auto des paramètres de
Suite ........ : ventilation
Suite ........ : BTY 06/11/07 IsMouvemente:ajout test présence cpte ds IMMOREGFR
Suite ........ : BTY 07/11/07 FQ 21792 estCpteAmortissement : ajout test compte
Suite ........ : est un compte de charge i_comptelie dans IMMO
Mots clefs ... : TOM;GENERAUX
*****************************************************************}
Unit CPGENERAUX_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
{$IFDEF VER150}
     variants,
{$ENDIF}
     ComCtrls,
     Graphics, // clWindow, clBtnFace, clRed
     Windows,  // VK_ ,
     Buttons,  // TBitBtn
     Menus,    // TPopupMenu
     //ExtCtrls, // TPanel
     HCtrls,   // ReadTokenSt, ReadTokenI
{$IFNDEF CCADM}
     ImpPrefG, // FicheImportPlanRef
{$ENDIF CCADM}
     LookUp,   // LookUpValueExist
{$IFDEF AMORTISSEMENT}
     AMLISTE_TOF,  // ConsultationImmo
     //OUTILS,       // toNone
{$ENDIF}

{$IFDEF COMPTA}
     CUMMENS,   // CumulCpteMensuel
{$IFNDEF CMPGIS35}
     CPBALGENAUXI_TOF,
     CPBALGENANAL_TOF,
{$ENDIF}
{$ENDIF}

{$IFNDEF PGIIMMO}
     CRITEDT,  // TCritEdt
{$ENDIF}

{$IFDEF EAGLCLIENT}
     MenuOLX,
     eFiche,           // TTFiche
     MaineAGL,         // AGLLanceFiche
     UtileAGL,         // TNavigateBtn
     CPREGION_TOF,     // pour le Paysregion
     CPCODEPOSTAL_TOF, // Pour le VerifCodePostal
{$ELSE}
     MenuOLG,
     db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     HDB,
     FE_Main,  // AGLLanceFiche
     dbCtrls,  // Contrôles DB, TNavigateBtn
     REGION,   // PaysRegion
     CODEPOST, // VerifCodePostal
     Fiche,        // TTFiche
     edtREtat,     // AglLanceFiche

{$ENDIF EAGLCLIENT}

     UtilSais, // MajSoldeSectionTOB
{$IFNDEF CCADM}
     LetBatch, // MajSoldeSection, VentileGenerale
{$ENDIF CCADM}
     VENTIL,   // ParamVentil

     {$IFNDEF COMPTAPAIE}
       {$IFNDEF PGIIMMO}
         {$IFNDEF CCADM}
         {$IFNDEF IMP}
           ZECRIMVT_TOF, {ZoomEcritureMvt}
           CPJUSTISOL_TOF, // FQ 17548
         {$ENDIF IMP}
         uRecupSQLModele, // RecupSQLMultiSoc
         {$ENDIF CCADM}
       {$ENDIF PGIIMMO}
     {$ENDIF COMPTAPAIE}

     //SG6 02.02.05 Gestion ventil types croisaxe
     CPVENTILTYPECROIS_TOF, //ParamVentilCroisaxe
     SaisUtil, // EstTVATPF
     FichComm, // FicheRIB_AGL
     {$IFNDEF PGIIMMO}
     {$IFNDEF IMP}
     {$IFNDEF CCADM}
     BANQUECP_TOM, // FicheBanqueCP
     {$ENDIF CCADM}
     {$endif}
     {$ENDIF}
     HTB97,    // TToolBarButton97,
     CHOIX,    // Choisir
     PARAMSOC, // GetParamSocSecur
     UTILPGI,  // _Blocage,
     HMsgBox,  // ExJaiLeDroitConcept, THMsgBox
     UTOM,
     HEnt1,    // MakeZoomOLE,
     UTOB,     // IsNull
     AglInit,  // ActionToString
     Ent1,     // TFichierBase, fbGene, fbAux, OkSynChro, BoureLess, LibellesTableLibre, ChangeMask, AfficheLeSolde, CoherencePaysRegion
     HRichOle, // THRichEditOle
     HRichEdt,  // THRichEditEdt
     uLibWindows  ,UentCommun // IIF
{$IFDEF MODENT1}
     , CPTypeCons
{$ENDIF MODENT1}
{$IFNDEF EAGLCLIENT}
     , AglIsoFlex
{$ENDIF}
     ;
Procedure FicheGene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer ) ;
Procedure FicheGeneMZS(Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
function  FicheGeneCreateOne : string ;
//ajout me  01/09/2005 déplacer dans ent1 procedure MajLettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
procedure MajDelettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;

function  CPChargeBlocNoteMillesime( vStType, vStCompte, vStExercice : string) : string;
function  CPGetMaxRangBlob( vStType, vStCompte : string) : integer;

// Ventilation de compte
{ Déplacement dans ulibanalytique pour EaglServer
function  CountNbEcriture ( LeCpte : String ; Where : String; aDossier : string = '') : integer;
function  VentileCompte( LeCpte : String ; Where : String ; VentilAxe : array of string ; PreVentil : Boolean; aDossier : string = '') : boolean;
procedure MajVentilation(LeCpte: String ; VentilAxe : array of string ;PreVentil : Boolean; aDossier : string = '');
}
{JP 23/10/07 : FQ 21714 : Gestion de la mise des ventilations en multi dossiers}
procedure MajVentilationMulti(LeCpte : string; VentilAxe : array of string; PreVentil : Boolean);


////////////////////////////////////////////////////////////////////////////////

Type
  TOM_GENERAUX = Class (TOM)
    MsgBox                  : THMsgBox;
    Pages                   : TPageControl;
    BMenuZoom               : TBitBtn;
    PopZ                    : TPopupMenu;
{$IFDEF EAGLCLIENT}
    G_GENERAL               : THEdit;
    G_LETTRABLE             : TCheckBox;
    G_COLLECTIF             : TCheckBox;
{$ELSE}
    G_GENERAL               : THDBEdit;
    G_LETTRABLE             : TDBCheckBox;
    G_COLLECTIF             : TDBCheckBox;
{$ENDIF}

    // ======= Evènement TOM
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

    procedure OnAfterFormShow;

    // ======= Evènement Forme
    procedure FormKeyDown                (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure PagesChanging              (Sender: TObject; var AllowChange: Boolean);
    procedure FormCloseQuery             (Sender: TObject; var CanClose: Boolean);

    // ======= Evènements sur boutons
    procedure BInsertClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BRIBClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BCumulClick(Sender: TObject);
    procedure BZecrimvtClick(Sender: TObject);
    procedure BJustifClick(Sender: TObject);
{$IFDEF COMPTA}
    procedure BBalGenAnaClick(Sender: TObject);
    procedure BBALGENAUXClick(Sender: TObject);
{$ENDIF}
    procedure BMenuZoomClick(Sender: TObject);
    procedure BImmoClick(Sender: TObject);
    procedure BGCycleRevisionClick(Sender : TObject);

    // ======= Gestion automatisme coordonnées bancaire ( ville, CP, div. terr., pays...)
    procedure G_VILLEDblClick(Sender: TObject);
    procedure G_CODEPOSTALDblClick(Sender: TObject);
    procedure G_DIVTERRITDblClick(Sender: TObject);
    procedure G_PAYSChange(Sender: TObject);

    // Gestion coche confidentiel -> G_CONFIDENTIEL
    Procedure OnClickConfidentiel(Sender : TObject);
    procedure OnClickCutOff(Sender : TObject);
    procedure OnClickCutOffEchue(Sender : TObject);

    //Gestion Coche G_VENTILABLE //SG6
    procedure OnClickVentilable(Sender : TObject);
    procedure OnClickGVentilableX(Sender : TObject); // 21209
    procedure NATUREGENEExit(Sender: TObject);
    // MAJ ABREGE POUR L'EAGL
    {$IFDEF EAGLCLIENT}
    procedure G_LIBELLEExit(Sender: TObject);
    {$ENDIF EAGLCLIENT}

    procedure OnMouseDownG_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure OnKeyPressG_General(Sender: TObject; var Key: Char);

    procedure DuplicVentil(StGene : string);
  private
    ChangeConfidentiel      : Boolean ;
    modifEnSerie            : Boolean ;             // indicateur de mode "modif en série" (en attendant éventuelle prise en compte AGL dans le TypeAction)
    LesModif                : String ;              // Pour modif en série
    LeCompte                : String ;              // Stocke le code du compte actuel pour éviter rechargements intempestifs
    AvecMvt                 : Boolean ;             // Indicateur compte mouvementé
    LgCode                  : Byte ;                // Indicateur du nb de car. du compte
    EnSaisieCpta            : Boolean ;             // Indicateur d'ouverture de la fiche saisie
    FInQuellePage           : Integer ;             // Numéro de l'Onglet visible à l'ouverture de l'écran

    MA0,MA1,MA2,MA3,MA4,MA5 : String3 ;             // SAV des caractéristiques du compte
    MP,ML,MNat,MColl        : String3 ;             // SAV des caractéristiques du compte
    RTVA,TVAE,TVA,TVASE     : String3 ;             // SAV des caractéristiques TVA

    Import                  : Boolean ;             // Indicateur d'importation des plan de Ref.

    SectAtt                 : Array[1..5] of Boolean ; // stockage des section d'attente

    FBoCutOff               : Boolean;
    {$IFDEF COMPTA}
    FBoVisaAutoEnCreation   : Boolean;
    {$ENDIF}

    FTobCumuls              : TOB ;                    // Contient les infos de cumuls locales en mode multisoc

    FAGLCloseQuery          : TCloseQueryEvent ;
    gbLoad                  : boolean;

    // ======= Initialisations
    procedure InitMsgBox;
    procedure InitComposants ;
    procedure InitEvenements ;
    Procedure InitParPays ;
    procedure AnalyseParametre( vStArg : String ) ;
    procedure InitVariables ;
    procedure InitPresentation ;
    function  OkConfidentiel : Boolean ;
    procedure InitModifEnSerie(StModif : string) ;      // pour les modif en série

    // ====== Fonctions de calcul / Maj affichage diverse
    procedure MajEcranInfos ;                            // Affichage des totaux / dates d'informations
    procedure FormateLesMontants ;                       // Paramétrage du formatage les montants
    procedure Ergo35 ;                                   // Gestion série S3/S5/S7
    procedure EnabledTva ;                               // ?? utilité à teser
    function  GetChampsMS( vStChamp : String ) : double ;   // Retourne le cumul demandé soit issu de la table commune, soit de la table locale en multisoc
    procedure ChargeCumulsMultiSoc ;                     // Renseigne FTobCumuls avec les valeurs locales en mode multisoc

    // ======= Fonctions TESTs
    function  IsMouvemente     : Boolean ;                 // ok
    function  IsCptBqe         : Boolean ;                 // ok
    function  IsCptRib         : Boolean ;                 // ok
    function  IsCptHt          : Boolean ;                 // ok
    function  IsCptTICTID      : Boolean ;                 // ok
    function  IsCptNatColl     : Boolean ;                 // ok
    function  IsCptVentilable  : Boolean ;
    function  IsCptLettrable   : Boolean ;
    function  IsCptPointable   : Boolean ;
{$IFDEF COMPTA}
    function  IsCptCollectif   : Boolean ;
{$ENDIF}
    function  IsCptIAS         : Boolean ;
{$IFNDEF IMP}
    function  IsCptImmo        : Boolean ;
{$ENDIF}
    function  IsCptVentilAxe( Quelaxe : Byte ) : Boolean ;
    function  IsCptAvecEcritureMultiEch        : Boolean ;

    // ======= Gestion dépendance interface
    procedure majEcranNatureGene ;
    procedure majEcranVentilAxe ;
    procedure ActivationControl;                         // Utilité à tester ??

    // ======= Gestion dépendance de données
    procedure MajChampsVentilAxe ;
    procedure MajChampsNatureGene ;
    procedure MajChampsGeneral ;
    procedure MajChampsLettrable ;
    function  VerifiNatureValide : Boolean ;
    function  VerifiSectionAtt( vStChamp : String ) : boolean ;
    function  CompleteGeneral : Boolean ;

    // ====== Gestion des accès
    procedure majAccesNatureGene ;
    procedure majAccesLettrage ;
    procedure majAccesPointage ;
    procedure majAccesVentil ;
    procedure majAccesIAS ;

    // ==== TEST ET MAJ VALIDATION
    procedure AutoriseModifications ;
    // -> Tests autorisation des modifs du compte
    function  VentilChanger : Boolean ;
    function  LettrageChanger : Boolean ;
    function  TVAChanger: Boolean;
    function  CollectifChanger : Boolean ;
    function  CollectifChangerChaPro : Boolean ;
    function  BqeCaiChanger : Boolean ;
    // -> Test saisie
    function  CodeValide : Boolean ;
    function  TvaTicTidOk : Boolean ;
    function  CaracteristiqueIncompatible(Var j : Byte) : Boolean ;
    // -> traitement suivant modification du compte
    procedure RepositionneVentilation ;
    procedure SupprimePreventilRib ;
    procedure MajCollChaProEcriture(Cpte,Aux : String) ;
    procedure CreerRibDefaut;
    procedure ChercheEtImportCptePlanRef;
    function  GetGoodPlanRef(var Compte : string ; NumPlanRef : string) : TQuery ; 
    //SG6 14.03.05 FQ 15483 > on ne met plus a jour le plan de référence
    //procedure MajDuPlanReference ;

    function InitEcrInfoPointage( vStCompte : string ) : Boolean;
{$IFNDEF EAGLCLIENT}
    //ISOFLEX
    function GereIsoflex: Boolean;
    procedure AppelIsoflex(Sender: TObject);
{$ENDIF}
{$IFDEF COMPTA}
    function  TrouveLiasseDuGene( vStCpte : string ) : string;
    function  TrouveGeneralSurTableRubrique( vStSelectSurRubrique, vStCpte : string): string;
{$ENDIF}
    function  ExisteEnregBanque:boolean; {Lek 23/09/05 FQ16663}


  public
    CptesCrees : String ;
    procedure BCopierClick(Sender: TObject);
end;  

function  EstEcrGuide          (LeCpte : String) : Boolean ;
function  EstAnalytiqPure      (LeCpte : String) : Boolean ;
function  EstCpteAxe           (LeCpte : String) : Boolean ;
function  EstBanquaire         (LeCpte : String) : Boolean ;
function  EstCpteCorresp       (LeCpte : String) : Boolean ;
function  EstCpteDevise        (LeCpte : String) : Boolean ;
function  EstCpteJournal       (LeCpte : String) : Boolean ;
function  EstCpteModepaie      (LeCpte : String) : Boolean ;
function  EstCpteSociete       (LeCpte : String) : Boolean ;
function  EstCpteTva           (LeCpte : String) : Boolean ;
function  EstCpteCollRef       (LeCpte : String) : Boolean ; //retourne si le cpte est référencé dans les Tiers
function  EstCpteAmortissement (LeCpte : String) : Boolean ;
function  EstCpteDansParamSoc  (LeCpte : String) : Boolean ;

Implementation

uses
{$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  CPVersion,
{$ENDIF MODENT1}

{$IFDEF COMPTA}
  Constantes,         // JP 03/08/05 : FQ 15990
{$IFNDEF CMPGIS35}
  Commun,             // JP 01/08/06 : Pour la gestion des comptes courants
  uLibRevision,       // MiseAJourInfoRevision, TrouveCycleRevisionDuGene
{$ENDIF}
  {$IFNDEF CCMP}
{$IFNDEF CMPGIS35}
  CRevParamCycle_Tom, // CPLanceFiche_CRevParamCycle
{$ENDIF}
  {$ENDIF}
{$ENDIF}
  uLibAnalytique,
  UProcGen, {StrRight}
  ed_tools, //VideListe
  HCompte,
  UFonctionsCBP;            // JP 02/06/06 : FQ 17971 : THCpteEdit

const  // YMO 03/2006 Champs à copier à partir du compte source par le bouton dupliquer
    DuplicValues : array[1..39] of string = ('G_SENS','G_VENTILABLE',
    // Caractéristiques
    'G_LETTRABLE','G_PURGEABLE','G_POINTABLE','G_IAS14',
    // Options du Grand-Livre
    'G_TOTAUXMENSUELS','G_SOLDEPROGRESSIF','G_SAUTPAGE','G_CENTRALISABLE','G_SUIVITRESO',
    'G_EFFET','G_CONSO','G_QUALIFQTE1','G_QUALIFQTE2','G_BLOCNOTE',
  // onglet règlements
    'G_REGIMETVA','G_SOUMISTPF','G_TVAENCAISSEMENT',
    'G_RISQUETIERS','G_MODEREGLE','G_JOURPAIEMENT1',
    'G_JOURPAIEMENT2','G_RELANCEREGLEMENT',
    'G_RELANCETRAITE','G_LETTREPAIEMENT',
    'G_MOTIFVIREMENT','G_RESIDENTETRANGER',
    'G_NATUREECONOMIQUE','G_PLAFOND','G_RISQUE',
    // onglet révision
    'G_CUTOFF','G_CUTOFFPERIODE','G_CUTOFFCOMPTE','G_CUTOFFECHUE',
    'G_VISAREVISION','G_TVA','G_TVASURENCAISS','G_TPF');


Procedure FicheGene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer ) ;
var lStAction : String ;
begin
  Case Comment of
    taCreat,taCreatEnSerie, taCreatOne : begin
      if Not ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then Exit ;
      lStAction := ActionToString( Comment ) ;
      if lStAction='' then lStAction := 'ACTION=CREATION' ;
      AGLLanceFiche('CP','CPGENERAUX','',Compte, lStAction );
      end;
    taModif,taModifEnSerie : begin
      if Not ExJaiLeDroitConcept(TConcept(ccGenModif),True) then Exit ;
      // _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrEnca','nrDeca'],True,'nrBatch') then Exit ;
      AGLLanceFiche('CP','CPGENERAUX','',Compte, 'ACTION=MODIFICATION;' + IntToStr(QuellePage));
    end;

    // Consultation
    taConsult : AGLLanceFiche('CP','CPGENERAUX','',Compte,'ACTION=CONSULTATION;' + IntToStr(QuellePage));

  end;
  if Comment in [taModif,taModifEnSerie] then _Bloqueur('nrBatch',False);
end;

function FicheGeneCreateOne : string ;
begin
  if Not ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then Exit ;
  AGLLanceFiche('CP','CPGENERAUX','','','ACTION=CREATION');
END ;

Procedure FicheGeneMZS(Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
begin
  Case Comment of
     taCreat,taCreatEnSerie,taCreatOne : begin
       if Not ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then Exit ;
      AGLLanceFiche('CP','CPGENERAUX','','','ACTION=CREATION;'+IntToStr(QuellePage)+';SERIE;'+LesModif);
       end;

     taModif,taModifEnSerie : begin
       if Not ExJaiLeDroitConcept(TConcept(ccGenModif),True) then Exit ;
       AGLLanceFiche('CP','CPGENERAUX','',compte,'ACTION=MODIFICATION;'+IntToStr(QuellePage)+';SERIE;'+LesModif);
       end;
  end;
end;

function EstEcrGuide(LeCpte : String) : Boolean ;
begin
  Result:=Presence('ECRGUI','EG_GENERAL',LeCpte);
END ;

function EstAnalytiqPure(LeCpte : String) : Boolean ;
begin
  Result:=Presence('ANALYTIQ','Y_GENERAL',LeCpte);
end;

function EstCpteAxe(LeCpte : String) : Boolean ;
var
  i : Byte ;
begin
  Result:=False ;
  for i:=1 to 5 do begin
    if VH^.Cpta[AxeToFb('A'+IntToStr(i))].AxGenAttente=LeCpte then begin Result:=True; Exit; end;
  end;
end;

function EstBanquaire(LeCpte : String) : Boolean ;
begin
  Result:=Presence('BANQUECP','BQ_GENERAL',LeCpte);
end;

function EstCpteCorresp(LeCpte : String) : Boolean ;
begin
  Result:=Presence('CORRESP','CR_CORRESP',LeCpte);
end;

function EstCpteDevise(LeCpte : String) : Boolean ;
begin
  Result := ExisteSQL('SELECT D_DEVISE FROM DEVISE WHERE(D_CPTLETTRDEBIT="'+LeCpte+'") OR (D_CPTLETTRCREDIT="'+LeCpte+'") '+'OR (D_CPTPROVDEBIT="'+LeCpte+'") OR (D_CPTPROVCREDIT="'+LeCpte+'")');
end;

function EstCpteJournal(LeCpte : String) : Boolean ;
begin
  Result := ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE="'+LeCpte+'" '+'AND ((J_NATUREJAL="BQE") OR (J_NATUREJAL="CAI") OR (J_NATUREJAL="REG"))');
end;

function EstCpteModepaie(LeCpte : String) : Boolean ;
begin
  Result:=Presence('MODEPAIE','MP_GENERAL',LeCpte) ;
end;

function EstCpteSociete(LeCpte : String) : Boolean ;
var St : String ;
begin
  Result:=True ;
  St:=GetParamSocSecur('SO_GENATTEND', ''); if St=LeCpte then Exit ;
  St:=GetParamSocSecur('SO_FERMEBIL',  ''); if St=LeCpte then Exit ;
  St:=GetParamSocSecur('SO_OUVREBIL',  ''); if St=LeCpte then Exit ;
  St:=GetParamSocSecur('SO_RESULTAT',  ''); if St=LeCpte then Exit ;
  St:=GetParamSocSecur('SO_FERMEPERTE',''); if St=LeCpte then Exit ;
  St:=GetParamSocSecur('SO_OUVREPERTE',''); if St=LeCpte then Exit ;
  St:=GetParamSocSecur('SO_FERMEBEN',  ''); if St=LeCpte then Exit ;
  St:=GetParamSocSecur('SO_OUVREBEN',  ''); if St=LeCpte then Exit ;
  Result:=False ;
end;

function EstCpteTva(LeCpte : String) : Boolean ;
begin
  Result:=EstTvaTpf(LeCpte,True) ;
  if Not Result then Result:=EstTvaTpf(LeCpte,False) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/10/2006
Modifié le ... :   /  /    
Description .. : modif mbo le 17.10.2006
Suite ........ : Ajout sur test compte de subvention ou reprise de
Suite ........ : subvention ou compte quote part de subvention
Suite ........ : BTY 06/11/07 Ajout test présence compte dans IMMOREGFR
Suite ........ : (table des passages forfait à réel en Agricole)
Suite ........ : BTY 07/11/07 FQ 21792 ajout test compte de charge i_comptelie
Mots clefs ... :
*****************************************************************}
Function  EstCpteAmortissement(LeCpte : String) : boolean;
var
CptSub : string;
begin

  result := ExisteSQL ( 'SELECT I_IMMO FROM IMMO WHERE '+
      'I_COMPTEIMMO="'+LeCpte+'" OR I_COMPTEAMORT="'+LeCpte+'" OR I_COMPTEDOTATION="'+LeCpte+
      '" OR I_COMPTEDEROG="'+LeCpte+'" OR I_REPRISEDEROG="'+LeCpte+'" OR I_PROVISDEROG="'+LeCpte+
      '" OR I_DOTATIONEXC="'+LeCpte+'" OR I_VACEDEE="'+LeCpte+'" OR I_AMORTCEDE="'+LeCpte+
      '" OR I_COMPTELIE="'+LeCpte+'" OR I_CPTSBVB="'+LeCpte+'" OR I_CPTSBVR="'+LeCpte+
      '" OR I_REPEXPLOIT="'+LeCpte+'" OR I_REPEXCEP="'+LeCpte+'" OR I_VAOACEDEE="'+LeCpte+'"');

  if (result = false) and (copy(LeCpte, 1,3) = '777') then
  begin
     CptSub := '131' + COPY(LeCpte, 4, (length (LeCpte))-3);
     result := ExisteSQL ('SELECT * FROM IMMO WHERE I_CPTSBVB="'+ CptSub+'" ');

     if result = false then
     begin
        CptSub := '138' + COPY(LeCpte, 4, (length (LeCpte))-3);
        result := ExisteSQL ('SELECT * FROM IMMO WHERE I_CPTSBVB="'+ CptSub+'" ');
     end;
  end;
  // BTY 06/11/07
  if not result then
     result := ExisteSQL ( 'SELECT IR_COMPTEIMMO FROM IMMOREGFR WHERE '+
                           'IR_COMPTEREF="'+LeCpte+'" OR IR_COMPTEAMORT="'+LeCpte+'" OR IR_COMPTEDOTATION="'+LeCpte+
                           '" OR IR_COMPTEDEROG="'+LeCpte+'" OR IR_REPRISEDEROG="'+LeCpte+'" OR IR_PROVISREDOG="'+LeCpte+
                           '" OR IR_DOTATIONEXC="'+LeCpte+'" OR IR_AMORTCEDE="'+LeCpte+
                           '" OR IR_REPEXPLOIT="'+LeCpte+'" OR IR_REPEXCEP="'+LeCpte+'"');
end;

///////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/09/2004
Modifié le ... :   /  /
Description .. : Test si le compte général est renseigné dans les PARAMSOC
Mots clefs ... :
*****************************************************************}
function EstCpteDansParamSoc(LeCpte : String) : Boolean ;
begin
  Result := (LeCpte = GetParamSocSecur('SO_CPTEAMORTINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEAMORTSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTECBSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDEPOTINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDEPOTSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDEROGINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDEROGSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDOTEXCINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDOTEXCSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDOTINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEDOTSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEEXPLOITINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEEXPLOITSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEFININF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEFINSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEIMMOINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEIMMOSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTELOCINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTELOCSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEPROVDERINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEPROVDERSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEREPDERINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEREPDERSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEREPEXCINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEREPEXCSUP','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEVACEDEEINF','')) or
            (LeCpte = GetParamSocSecur('SO_CPTEVACEDEESUP','')) or
            (LeCpte = GetParamSocSecur('SO_DEFCOLCDIV','')) or
            (LeCpte = GetParamSocSecur('SO_DEFCOLCLI','')) or
            (LeCpte = GetParamSocSecur('SO_DEFCOLDDIV','')) or
            (LeCpte = GetParamSocSecur('SO_DEFCOLDIV','')) or
            (LeCpte = GetParamSocSecur('SO_DEFCOLFOU','')) or
            (LeCpte = GetParamSocSecur('SO_DEFCOLSAL','')) or
            (LeCpte = GetParamSocSecur('SO_GENATTEND','')) or
            (LeCpte = GetParamSocSecur('SO_FERMEBEN','')) or
            (LeCpte = GetParamSocSecur('SO_FERMEBIL','')) or
            (LeCpte = GetParamSocSecur('SO_FERMEPERTE','')) or
            (LeCpte = GetParamSocSecur('SO_OUVREBEN','')) or
            (LeCpte = GetParamSocSecur('SO_OUVREBIL','')) or
            (LeCpte = GetParamSocSecur('SO_OUVREPERTE','')) or
            (LeCpte = GetParamSocSecur('SO_PRODEB1','')) or
            (LeCpte = GetParamSocSecur('SO_PRODEB2','')) or
            (LeCpte = GetParamSocSecur('SO_PRODEB3','')) or
            (LeCpte = GetParamSocSecur('SO_PRODEB4','')) or
            (LeCpte = GetParamSocSecur('SO_PRODEB5','')) or
            (LeCpte = GetParamSocSecur('SO_PROFIN1','')) or
            (LeCpte = GetParamSocSecur('SO_PROFIN2','')) or
            (LeCpte = GetParamSocSecur('SO_PROFIN3','')) or
            (LeCpte = GetParamSocSecur('SO_PROFIN4','')) or
            (LeCpte = GetParamSocSecur('SO_PROFIN5','')) or
            (LeCpte = GetParamSocSecur('SO_RESULTAT','')) or
            (LeCpte = GetParamSocSecur('SO_LETCHOIXGEN','')) or
            (LeCpte = GetParamSocSecur('SO_LETCHOIXGENC','')) or
            (LeCpte = GetParamSocSecur('SO_ICCCOMPTECAPITAL','')) or
            (LeCpte = GetParamSocSecur('SO_TELETVABQE1','')) or
            (LeCpte = GetParamSocSecur('SO_TELETVABQE2','')) or
            (LeCpte = GetParamSocSecur('SO_TELETVABQE3','')) or
            // DOMAINE PAIE (P)
            (LeCpte = GetParamSocSecur('SO_PGCPTNETAPAYER','')) or
            // DOMAINE GESCOM (G)
            (LeCpte = GetParamSocSecur('SO_GCCPTEESCACH','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEESCVTE','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEREMACH','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEESCVTE','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEHTACH','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEHTVTE','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEHTVTE','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEHTVTE','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEPORTACH','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEPORTVTE','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTERGVTE','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTESTOCK','')) or
            (LeCpte = GetParamSocSecur('SO_GCCPTEVARSTK',''));
end;

////////////////////////////////////////////////////////////////////////////////
function  EstCpteCollRef(LeCpte : String) : Boolean ;
begin
    result := Presence('TIERS','T_COLLECTIF',LeCpte);
end;

(* ajout me  01/09/2005
procedure MajLettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
var
  LeWhere : String ;
begin
Case Lefb of
//     fbGene :LeWhere:='WHERE E_GENERAL="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
     fbGene :LeWhere:='WHERE E_GENERAL="'+Cpte+'"' ;  // CA - 22/04/2003 - On fait la manip sur toutes les écritures sinon pb dans rapport erreur ( compte lettrable avec état letrage incohérent ).
//     fbAux  :LeWhere:='WHERE E_AUXILIAIRE="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
     fbAux  :LeWhere:='WHERE E_AUXILIAIRE="'+Cpte+'"' ; // CA - 22/04/2003 - idem préc.
    End ;

{JP 03/08/05 : FQ 15990 : Gestion de E_TRESOSYNCHRO}
ExecuteSql('UPDATE ECRITURE SET E_MODEPAIE = "' + Mp + '", E_DATEECHEANCE = E_DATECOMPTABLE, ' +
           'E_ETATLETTRAGE = "AL", E_NUMECHE = 1, E_ECHE = "X", E_PAQUETREVISION = 1, E_TRESOSYNCHRO = "' +
           ets_Nouveau + '" ' + LeWhere)
end;
*)

procedure MajDelettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
var
  LeWhere : String ;
begin
Case Lefb of
     fbGene :LeWhere:='WHERE E_GENERAL="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
     fbAux  :LeWhere:='WHERE E_AUXILIAIRE="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
    End ;
BeginTrans ;
ExecuteSql('UPDATE ECRITURE SET E_MODEPAIE="'+Mp+'", E_DATEECHEANCE=E_DATECOMPTABLE, '+'E_ETATLETTRAGE="RI", E_NUMECHE=0, E_ECHE="-" '+LeWhere) ;
CommitTrans ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPChargeBlocNoteMillesime( vStType, vStCompte, vStExercice : string) : string;
var lQuery : TQuery;
begin
  Result := '';
  lQuery := nil;
  try
    lQuery := OpenSQL('SELECT LO_OBJET FROM LIENSOLE WHERE ' +
                     'LO_TABLEBLOB = "' + vStType + '" AND ' +
                     'LO_IDENTIFIANT = "' + vStCompte + '" AND ' +
                     'LO_EMPLOIBLOB = "' + vStExercice + '"', True);

    if not lQuery.Eof then
      Result := lQuery.FindField('LO_OBJET').AsString;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPGetMaxRangBlob( vStType, vStCompte : string) : integer;
var lQuery : TQuery;
begin
  lQuery := nil;
  try
    lQuery := OpenSQL('SELECT MAX(LO_RANGBLOB) LEMAX FROM LIENSOLE WHERE ' +
                      'LO_TABLEBLOB = "' + vStType + '" AND ' +
                      'LO_IDENTIFIANT = "' + vStCompte + '"', True);

    if not lQuery.Eof then
      Result := lQuery.FindField('LEMAX').AsInteger
    else
      Result := 0;

  finally
    Ferme(lQuery);
  end;

end;

////////////////////////////////////////////////////////////////////////////////

procedure TOM_GENERAUX.OnNewRecord ;
begin
  Inherited ;

  // Init variables
  gbLoad  := True;
  AvecMvt := False ;
  Import  := True ;
  MNat    := '' ;
  MA0     := '-' ;
  MA1     := '-' ;
  MA2     := '-' ;
  MA3     := '-' ;
  MA4     := '-' ;
  MA5     := '-' ;
  MP      := '-' ;
  ML      := '-' ;
  MColl   := '-' ;

  RTVA:= GetControlText('G_REGIMETVA');
  TVAE:= GetControlText('G_TVAENCAISSEMENT');
  TVA:=  GetControlText('G_TVA');
  TVASE:=GetControlText('G_TVASURENCAISS');

  // Init certains champs
  SetField('G_SUIVITRESO',         'RIE');
  SetField('G_PURGEABLE',          'X');
  SetField('G_SOLDEPROGRESSIF',    'X');
  SetField('G_CONFIDENTIEL',       '0');
  SetField('G_SENS',               'M');
  SetField('G_RISQUETIERS',        '-');

  // Accès aux champs code / nature
  SetControlEnabled('G_GENERAL',True);
  SetControlEnabled('G_NATUREGENE',True);

  // Resmies à zéro des soldes
  MajEcranInfos ;

  // Page par défaut, démarrage saisie sur code
  Pages.ActivePage := TTabSheet(GetControl('PGENERAL',True));
  SetFocusControl('G_GENERAL');

  SetControlEnabled('BMENUZOOM', False );
  SetControlEnabled('BIMPRIMER', False );

  SetControlEnabled('BCopier', True);{JP 05/06/06 : FQ17971}
  gbLoad := false;
end ;

procedure TOM_GENERAUX.OnDeleteRecord ;
var lStCompte : String ;
begin
  Inherited ;

  lStcompte := GetField('G_GENERAL') ;
  if lStcompte = '' then Exit ;

  ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="' + lStcompte + '" AND V_NATURE LIKE "GE%"') ;
  ExecuteSql('DELETE FROM BANQUECP WHERE BQ_GENERAL="' + lStcompte + '"');
  ExecuteSql('DELETE FROM CLIENGENEMODELA WHERE CLA_GENERAL="' + lStcompte + '"');  {FP 29/12/2005}

  // GCO - 08/02/2007 - NB : Jamais de ONDELETERECORD, car pas de bouton supprimer
  // Cependant j'ajoute tout de même les DELETE au cas où, on le rajoute
  ExecuteSql('DELETE FROM CREVGENERAUX WHERE CRG_GENERAL = "' + lStCompte + '"');
  ExecuteSql('DELETE FROM CREVBLOCNOTE WHERE CBN_NATURE = "GEN" AND CBN_CODE = "' + lStCompte + '"');
  // FIN GCO 
end ;

procedure TOM_GENERAUX.OnUpdateRecord ;
var j : Byte ;
{$IFDEF COMPTA}
    lStEtatCycle : string;
{$ENDIF}
begin
  Inherited ;

  j := 0;

  // ============================================
  // TEST DONNEES OBLIGATOIRES / COHERENCE SAISIE
  // ============================================

  // Code obligatoire et unique
  if (Not CodeValide) then
    begin
    LastError := 1;
    Exit;
    end;

  // Libellé obligatoire
  if (GetField('G_LIBELLE') = '') then
    begin
    MsgBox.Execute(3,'','');  // Vous devez renseigner un libellé.
    Pages.ActivePage := TTabSheet(GetControl('PGENERAL',True));
    SetFocuscontrol('G_LIBELLE');
    LastError := 1;
    Exit;
    end;

  // Nature du compte obligatoire
  if GetField('G_NATUREGENE') = '' then
    begin
    MsgBox.Execute(47,'','');  // Vous devez renseigner une nature
    Pages.ActivePage := TTabSheet(GetControl('PGENERAL',True));
    SetFocuscontrol('G_NATUREGENE');
    LastError := 1;
    Exit;
    end ;

  // Test de cohérence des caractériqtiques du compte
  if (CaracteristiqueIncompatible(j)) then
    Case j of
        1 : begin MsgBox.Execute(25,'',''); LastError := 1; Exit; end; // Vous ne pouvez pas rendre un compte ventilable, pointable et lettrable à la fois.
        2 : begin MsgBox.Execute(26,'',''); LastError := 1; Exit; end; // Vous ne pouvez pas rendre un compte ventilable et lettrable à la fois.
        3 : begin MsgBox.Execute(27,'',''); LastError := 1; Exit; end; // Vous ne pouvez pas rendre un compte pointable et lettrable à la fois.
        4 : begin MsgBox.Execute(28,'',''); LastError := 1; Exit; end; // Vous ne pouvez pas rendre un compte ventilable et pointable à la fois.
    end;

  // Test de cohérence des données de TVA
  if (Not TvaTicTidOk) then
    begin
    LastError := 1;
    Exit;
    end;

  // Cohérence Pays / Région ?
  if (Not CoherencePaysRegion('G_',DS)) then
    begin
    MsgBox.Execute(43,'','');  // Incohérence dans la saisie du couple "Pays - Division territoriale" !
    Pages.ActivePage := TTabSheet(GetControl('PADRESSE',True));
    SetFocuscontrol('G_PAYS');
    LastError := 1;
    Exit;
    end;

  // Gestion de la validité des compte de correspondance
  Case (VerifCorrespondance(1,GetField('G_CORRESP1'),GetField('G_CORRESP2'))) of
      0 : ;
      1 : begin Msgbox.Execute(41,'','') ; LastError := 1; Exit ; end; // Le compte de correspondance n°1 que vous avez saisi n''existe pas !
      2 : begin Msgbox.Execute(42,'','') ; LastError := 1; Exit ; end; // Le compte de correspondance n°2 que vous avez saisi n''existe pas !
  end;

  // MAJ liste des comptes créés
  if (DS.State in [dsInsert]) and ((GetField('G_GENERAL')<>'')
    and (Pos(GetField('G_GENERAL')+';',CptesCrees)<=0)) then
    CptesCrees := CptesCrees + GetField('G_GENERAL')+';';

  // =========================================================================================
  // TESTS SUIVANTS CONCERNENT LES AUTORISATIONS DE MODIFICATION DES CARACTERITIQUES DU COMPTE
  // =========================================================================================
  if not (DS.State in [dsBrowse,dsInsert]) then
  begin
    // ==> MAJ de l'indicateur "compte mouvementé" au cas ou
    if Not AvecMvt  then
      AvecMvt := IsMouvemente ;
    // EN PCL : Modification interdite de la nature si le compte est mouvementé
    // FQ 12063 - En entreprise aussi, on interdit la modification
    if AvecMvt (*and (ctxPcl in V_Pgi.PGIContexte) *) and (MNat <> GetField('G_NATUREGENE')) then
    begin
      MsgBox.Execute(54,'','');  // La modification de la nature est impossible : ce compte est mouvementé.
      // SetField('G_NATUREGENE',MNat); ??
      LastError := 1;
      Exit;
    end;
    // Autres tests
    AutoriseModifications ;
  end;

  {$IFDEF COMPTA}
  if DS.State = DsInsert then
  begin
    try
      ExecuteSQL('INSERT INTO GENERAUX (G_GENERAL) VALUES( "' + GetField('G_GENERAL') + '")');

      if VH^.Revision.Plan <> '' then
      begin
        // Ici, calcul à quel Cycle de Révision appartient le compte saisi, étant
        // que l'on gère le dossier avec un Plan de Révision

        // Problématique, il faut enregister momentanément un enregistrement fictif
        // pour savoir si le compte que l'on crée va ressortir dans un select sur
        // les cyles de révision
{$IFNDEF CMPGIS35}
        SetField('G_CYCLEREVISION', TrouveCycleRevisionDuGene(GetField('G_GENERAL')));

        if GetField('G_CYCLEREVISION') = '' then
        begin
          // Aucun cycle de révision ne contient le général saisi
          if GetParamSocSecur('SO_CPREVISBLOQUEGENE', False) then
          begin
            PgiInfo('Enregistement impossible. Le compte général n''appartient à aucun cycle de révision.', 'Comptes généraux');
            LastError := -1;
            Exit;
          end
          else
            PgiInfo('Le compte général n''appartient à aucun cycle de révision.', 'Comptes généraux');
        end
        else
        begin
          lStEtatCycle := GetColonneSQL('CREVCYCLE',  'CCY_ETATCYCLE',
                                        'CCY_CODECYCLE = "' + GetField('G_CYCLEREVISION') + '" AND ' +
                                        'CCY_EXERCICE = "' + VH^.EnCours.Code + '"');

          if (lStEtatcycle = cValide) or (lStEtatCycle = cSupervise) then
          begin
            FBoVisaAutoEnCreation := True;
            SetField('G_VISAREVISION', 'X'); (**)
            // La mise à jour de CREVGENERAUX est faite dans le OnAfterUpdateRecord
          end;
        end;
{$ENDIF}
      end;

      if (not GetParamSocSecur('SO_CPREVISLIASSEMODIF', False)) and
         (GetParamSocSecur('SO_CPCONTROLELIASSE', False) <> '') and
         (GetParamSocSecur('SO_CPCONTROLELIASSE', False) <> 'AUC') then
      begin
        // Recherche du général dans les liasses ( sur Table RUBRIQUE )
        if TrouveLiasseDuGene( GetField('G_GENERAL')) = '' then
        begin
          if GetParamSocSecur('SO_CPREVISBLOQUELIASSE', False) then
          begin
          PgiInfo('Enregistrement impossible. Le compte général n''appartient pas à la liasse spécifiée dans vos paramètres société.', 'Comptes généraux');
            LastError := -1;
            Exit;
          end
          else
            PgiInfo('Le compte général n''appartient pas à la liasse spécifiée dans vos paramètres société.', 'Comptes généraux');
        end;
      end;

    finally
      // Suppression systématique de l'enregistrement temporaire
      ExecuteSQL('DELETE FROM GENERAUX WHERE G_GENERAL = "' + GetField('G_GENERAL') + '"');
{$IFDEF COMPTA}
  // fiche 20549
    MAJHistoparam ('GENERAUX',  GetField('G_GENERAL'));
{$ENDIF}

    end;
  end;
  {$ENDIF}
end;
////////////////////////////////////////////////////////////////////////////////

procedure TOM_GENERAUX.OnAfterUpdateRecord ;
var lStCompte : String ;
begin

  Inherited ;

  lStcompte := GetField('G_GENERAL') ;
  if lStcompte = '' then Exit ;

  // GESTION DES PREVENTILATIONS
  if not IsCptVentilable then
    ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="' + lStcompte + '" AND V_NATURE LIKE "GE%"') ;

  // GESTION DES RIBS
  if IsCptBqe then
  begin {Lek 23/09/05 FQ16663}
    if not ExisteEnregBanque then BRibClick(Nil);
  end else ExecuteSql('DELETE FROM BANQUECP WHERE BQ_GENERAL = "' + lStcompte + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');

  //SG6 14.03.05 FQ 15483 > on ne met plus a jour le plan de référence
 { if (not (ctxPCL in V_PGI.PGIContexte)) then
    MajDuPlanReference ;}

  // GCO - 20/07/2004 - Traitement sur écritures sur compte devenu POINTABLE
  if IsFieldModified('G_POINTABLE') then
  begin
    if GetField('G_POINTABLE') = 'X' then
    begin
      if not InitEcrInfoPointage( lStCompte ) then
      begin
        PgiInfo('Traitement annulé. Erreur lors de la mise à jour des écritures.', 'Fonction : InitEcrInfoPointage');
        LastError := -1;
        Exit;
      end;
    end;
  end;
  // FIN GCO

  // GCO - 08/02/2007 - Recopie de G_VISAREVISON et G_CYCLEREVISION dans CREVGENERAUX
{$IFDEF COMPTA}
  if IsFieldModified('G_VISAREVISION') then
  begin
{$IFNDEF CMPGIS35}
    if not MiseAJourCREVGeneraux( lStCompte, FBoVisaAutoEnCreation ) then
    begin
      PgiError('Traitement annulé. Erreur lors de la mise à jour de la table CREVGENERAUX.', 'Fonction : MiseAJourCREVGeneraux');
      LastError := -1;
      Exit;
    end;
{$ENDIF}
  end;
{$ENDIF}

  if modifEnSerie then
    TFFiche(Ecran).BFermeClick(nil) ;

  //SG6 05/01/05 FQ 14597 Mis à jour de l'interface après enregistrement
  majAccesVentil;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : GC
Créé le ...... : 08/03/2005
Modifié le ... :   /  /
Description .. : - LG - 08/03/2005 - optimisation du chargement de liensole
Mots clefs ... :
*****************************************************************}
procedure TOM_GENERAUX.OnLoadRecord ;
{$IFDEF COMPTA}
var lQuery : TQuery;
{$ENDIF}
begin
  gbLoad := true;
  Inherited ;
  
  { FQ 21209 BVE 01.08.07 }
  TCheckBox(GetControl('G_VENTILABLE1',true)).OnClick     := OnClickGVentilableX;
  TCheckBox(GetControl('G_VENTILABLE2',true)).OnClick     := OnClickGVentilableX;
  TCheckBox(GetControl('G_VENTILABLE3',true)).OnClick     := OnClickGVentilableX;
  TCheckBox(GetControl('G_VENTILABLE4',true)).OnClick     := OnClickGVentilableX;
  TCheckBox(GetControl('G_VENTILABLE5',true)).OnClick     := OnClickGVentilableX;
  { END FQ 21209 }

  TCheckBox(GetControl('GCONFIDENTIEL',true)).OnClick := nil;
  TCheckBox(GetControl('GCONFIDENTIEL',true)).Checked := (GetField('G_CONFIDENTIEL') = '1');
  TCheckBox(GetControl('GCONFIDENTIEL',true)).OnClick := OnClickConfidentiel;
  FBoCutOff := ( GetField('G_NATUREGENE') = 'CHA' ) or ( GetField('G_NATUREGENE') = 'PRO' ) ;

  // Utilité ??? Pour en pas y passer 2 fois ? ca arrive ?
  if not (DS.State in [dsInsert]) then
    begin
    if (LeCompte = GetField('G_GENERAL')) then begin
      LastError := 1;
      Exit ;
    end;

  LeCompte := GetField('G_GENERAL');

  // Code du compte acessible uniquement à la création
  SetControlEnabled('G_GENERAL',False);

  // YMO Bouton dupliquer acessible uniquement à la création
  SetControlEnabled('BCopier', (DS.State = dsInsert)); {JP 05/06/06 : FQ17971}

  // chargement des infos de cumuls MULTISOC
  ChargeCumulsMultiSoc ;


    // compte mouvementé ?
    AvecMvt := IsMouvemente ;
    // FQ 12063 - En entreprise aussi, on interdit la modification
    if (*(ctxPcl in V_Pgi.PGIContexte) and *) avecMvt then
      begin
      SetControlEnabled('G_NATUREGENE', not AvecMvt );
      if GetControlEnabled('G_NATUREGENE')
        then SetControlProperty('G_NATUREGENE', 'COLOR', ClWindow )
        else SetControlProperty('G_NATUREGENE', 'COLOR', ClBtnFace ) ;
      end ;


    // SAV des caratérisques du compte
    MA0     := GetField('G_VENTILABLE') ;
    MA1     := GetField('G_VENTILABLE1') ;
    MA2     := GetField('G_VENTILABLE2') ;
    MA3     := GetField('G_VENTILABLE3') ;
    MA4     := GetField('G_VENTILABLE4') ;
    MA5     := GetField('G_VENTILABLE5') ;
    MP      := GetField('G_POINTABLE') ;
    ML      := GetField('G_LETTRABLE') ;
    MColl   := GetField('G_COLLECTIF') ;
    MNat    := GetField('G_NATUREGENE') ;

    RTVA:= GetControlText('G_REGIMETVA');
    TVAE:= GetControlText('G_TVAENCAISSEMENT');
    TVA:=  GetControlText('G_TVA');
    TVASE:=GetControlText('G_TVASURENCAISS');

    // Accès boutons
    if EnSaisieCpta or modifEnSerie then
      SetControlEnabled('BINSERT', False) ;
    SetControlEnabled('BMENUZOOM', True );
    SetControlEnabled('BIMPRIMER', True );

    // Gestion mode modif en série (affichage des zones modifiées en rouge)
    if modifEnSerie then
      InitModifEnSerie(LesModif);

    // Gestion des acces caractèristiques du compte
    majAccesNatureGene ;
  //  majEcranNatureGene ;

    // Affichage des soldes
    MajEcranInfos ;

  end ;

  SetControlProperty( 'G_CUTOFF'       , 'Enabled' ,  FBoCutOff and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  SetControlProperty( 'G_CUTOFFPERIODE', 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  SetControlProperty( 'G_CUTOFFECHUE'  , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  SetControlProperty( 'LCUTOFFPER'     , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  SetControlProperty( 'G_CUTOFFECHUE'  , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  SetControlProperty( 'G_CUTOFFCOMPTE' , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  SetControlProperty( 'LCUTOFFCOMPTE'  , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  SetControlProperty( 'LCUTOFFECHUE'   , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;

  // Gestion du Cycle de révision
  SetControlVisible('BGCYCLEREVISION', False);
  SetcontrolProperty('GCYCLEREVISION', 'Caption', 'Aucun');

{$IFDEF COMPTA}
  // Ajout GCO - 23/04/2007 - Dossier est supervisé, blocage de la révision
{$IFNDEF CMPGIS35}
  SetControlProperty('G_VISAREVISION', 'ENABLED',
                     AutoriseSuppresionVisaRevision(GetField('G_GENERAL')) and
                     (TFFiche(Ecran).TypeAction <> taConsult) and
                     (not VH^.Revision.DossierSupervise));
{$ENDIF}

  if GetField('G_CYCLEREVISION') <> '' then
  begin
    SetControlVisible('BGCYCLEREVISION', True);

    // GCO - 10/05/2007 - FQ 20277
    lQuery := nil;
    try
      lQuery := OpenSQL('SELECT CCY_lIBELLECYCLE, CCY_ETATCYCLE FROM CREVCYCLE WHERE ' +
                        'CCY_CODECYCLE = "' + GetField('G_CYCLEREVISION') + '" AND ' +
                        'CCY_EXERCICE = "' + VH^.EnCours.Code + '"', True);

      if not lQuery.Eof then
      begin
        SetControlProperty('GCYCLEREVISION', 'CAPTION', GetField('G_CYCLEREVISION') + ' - ' +
                           lQuery.FindField('CCY_LIBELLECYCLE').AsString);

{$IFNDEF CMPGIS35}
        if (lQuery.FindField('CCY_ETATCYCLE').AsString = cSupervise) or
           ((lQuery.FindField('CCY_ETATCYCLE').AsString = cValide) and (not JaiLeRoleCompta(rcSuperviseur))) then
        begin
          SetControlProperty('G_BLOCNOTE', 'READONLY', True);
          SetControlProperty('G_BLOCNOTE', 'COLOR', ClBtnFace);
        end;
{$ENDIF}
      end;

    finally
      Ferme( lQuery );
    end;
  end;  
{$ENDIF}

  // GCO - 05/01/2004 - FQ 12344 et FQ 16353
  if (CtxPCl in V_Pgi.PgiContexte) then
  begin
    if FInQuellePage = 2 then
      SetFocusControl('G_BLOCNOTE');
  end;

  ChangeConfidentiel := False ;
  gbLoad := false;
end ;

procedure TOM_GENERAUX.OnChangeField ( F: TField ) ;
var
  res      : boolean;
  i        : integer;
  lAxeVent : array [1..5] of Boolean; // FQ 21259
begin
  Inherited ;

  SetControlEnabled('BMENUZOOM', ( not (DS.State in [dsInsert]) ) );
  SetControlEnabled('BIMPRIMER', ( not (DS.State in [dsInsert]) ) );

  // GESTION NATURE GENE
  if F.FieldName='G_NATUREGENE' then
    begin
    if not VerifiNatureValide then
      begin
      majEcranNatureGene ;
      majChampsNatureGene ;
      ActivationControl;
      end;
    end
  // GESTION GENERAL
  else if F.FieldName='G_GENERAL' then
    begin
    if not CompleteGeneral then
      begin
      majAccesVentil;
      MajChampsGeneral ;
      ActivationControl ;
      if DS.State = dsInsert then
      begin                           
         if not(Presence('GENERAUX','G_GENERAL',GetField('G_GENERAL'))) then
         begin
            res := InitialiseVentilDefaut(GetField('G_GENERAL'),SectAtt,lAxeVent) ; // FQ 21259
            // MAJ Axe de ventilation
            if res then
            begin
               for i := 1 to 5 do
                  if lAxeVent[i] then  // FQ 21259
                     SetField('G_VENTILABLE' + IntToStr(i),'X')
                  else
                     SetField('G_VENTILABLE' + IntToStr(i),'-');
               // MAJ accès zone du code
               SetControlEnabled( 'G_GENERAL', Not res );
               SetControlEnabled( 'BVENTIL',   res );  // Accès ventilation par défaut si compte ventilable
            end;
         end;
      end;
      end ;
    end
  // GESTION LETTRABLE
  else if F.FieldName='G_LETTRABLE' then
    begin
    // Test le paramsoc du mode de paiement par défaut
    if ( GetField('G_LETTRABLE') = 'X' ) and ( GetField('G_NATUREGENE') = 'DIV' )
                                         and ( GetParamSocSecur('SO_GCMODEREGLEDEFAUT', False) = '' )
      then
        begin
        PGIBox( 'Vous devez renseigner un mode de règlement par défaut dans les paramètres société, avant de paramètrer des comptes divers lettrables.', Ecran.Caption ) ;
        SetField('G_LETTRABLE', '-' ) ;
        SetControlText('G_LETTRABLE', '-' ) ;
        end
      else majChampsLettrable ;
    end
  // GESTION VENTILABLE
  else if ( F.FieldName='G_VENTILABLE1' ) or ( F.FieldName='G_VENTILABLE2' ) or
          ( F.FieldName='G_VENTILABLE3' ) or ( F.FieldName='G_VENTILABLE4' ) or
          ( F.FieldName='G_VENTILABLE5' ) then
    begin
      if not VerifiSectionAtt( F.FieldName ) then
      begin
        majEcranVentilAxe ;
        MajChampsVentilAxe ;
      end ;
    end
    else
      if ( F.FieldName='G_CUTOFFPERIODE' ) or ( F.FieldName='G_CUTOFFECHUE' ) then
       OnClickCutOff(GetControl('G_CUTOFF')) ;

  if (VH^.PaysLocalisation=CodeISOES) and (F.FieldName='G_TVA') and (DS.State in [dsInsert,dsEdit]) then
     SetField('G_TPF',F.AsString) ;
end ;



procedure TOM_GENERAUX.OnArgument ( S: String ) ;
{$IFNDEF EAGLCLIENT}
var
  BIsoFlex : TToolBarButton97;
{$ENDIF}
begin
  Inherited ;

{$IFNDEF EAGLCLIENT}
    if V_PGI.ModePcl <> '1' then
    begin
      if GereIsoflex then
      begin
        BIsoFlex := TToolBarButton97(GetControl('BISOFLEx'));
        if BISOFLEX <> nil then Bisoflex.onClick := AppelIsoflex;
      end;
    end;
{$ENDIF}

  // GCO - 03/03/2005
  SetControlProperty('G_GENERAL', 'MAXLENGTH', VH^.CPta[fbGene].Lg);

  // Initialisation fiche
  InitMsgBox;
  InitComposants ;
  InitEvenements ;
  InitParPays ;

  // Récupère les arguments
  AnalyseParametre( S ) ;

  // Initialisation Variables
  InitVariables ;

  // Initialisation présentation
  InitPresentation ;
  
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_GENERAUX.OnAfterFormShow;
begin
  Pages.ActivePage := Pages.Pages[FInQuellePage];
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOM_GENERAUX.OnClose;
var
  Gen : string;
begin
  Inherited ;
  if Assigned( MsgBox ) then FreeAndNil( MsgBox ) ;
  if Assigned( FTobCumuls ) then FreeAndNil( FTobCumuls ) ; // SBO Ajout MultiSoc

  {JP 03/08/05 : FQ 16125 : Si on annule la création du général et que la ventilation et le rib ont
                 été créés, il faut les détruire : je me suis inspiré de ce qui se fait dans les Tiers}
  Gen := GetField('G_GENERAL');
  if Gen = #0 then Gen := '' ;
  if (BourreEtLess(Gen, fbGene) = '') then Exit;

  if (not (ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + Gen + '"'))) then
    begin
    ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE = "' + Gen + '" AND V_NATURE LIKE "GE%"') ;
    ExecuteSql('DELETE FROM BANQUECP WHERE BQ_GENERAL = "' + Gen + '"');
    end;
end ;

procedure TOM_GENERAUX.OnCancelRecord ;
begin
  Inherited ;
  SupprimePreventilRib ;
end ;

procedure TOM_GENERAUX.InitMsgBox;
begin
  MsgBox := THMsgBox.create(FMenuG);
  MsgBox.Mess.Add('0;Comptes généraux;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;C;');
  MsgBox.Mess.Add('1;Comptes généraux;Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;');
  MsgBox.Mess.Add('2;Comptes généraux;Vous devez renseigner un compte.;W;O;O;O;');
  MsgBox.Mess.Add('3;Comptes généraux;Vous devez renseigner un libellé.;W;O;O;O;');
  MsgBox.Mess.Add('4;Comptes généraux;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
  MsgBox.Mess.Add('5;Comptes généraux;Vous devez associer un compte bancaire à ce compte.;W;O;O;O;');
  MsgBox.Mess.Add('6;Comptes généraux;Ce compte ne possède pas de compte bancaire associé. Voulez-vous le créer ?;Q;YNC;Y;C;');
  MsgBox.Mess.Add('7;Comptes généraux;Suppression impossible : ce compte est défini en tant que modèle.;W;O;O;O;');
  MsgBox.Mess.Add('8;Comptes généraux;Suppression impossible : ce compte possède des écritures analytiques.;W;O;O;O;');
  MsgBox.Mess.Add('9;Comptes généraux;Suppression impossible : ce compte est défini en tant que compte d''attente d''un axe analytique.;W;O;O;O;');
  MsgBox.Mess.Add('10;Comptes généraux;Suppression impossible : ce compte est associé à un compte bancaire.;W;O;O;O;');
  MsgBox.Mess.Add('11;Comptes généraux;Suppression impossible : ce compte est un compte de correspondance.;W;O;O;O;');
  MsgBox.Mess.Add('12;Comptes généraux;Suppression impossible : ce compte est défini en tant que compte d''écart de change.;W;O;O;O;');
  MsgBox.Mess.Add('13;Comptes généraux;Suppression impossible : ce compte possède des mouvements d''immobilisation.;W;O;O;O;');
  MsgBox.Mess.Add('14;Comptes généraux;Suppression impossible : ce compte est défini en tant que compte de contrepartie.;W;O;O;O;');
  MsgBox.Mess.Add('15;Comptes généraux;Suppression impossible : ce compte est associé à un mode de paiement.;W;O;O;O;');
  MsgBox.Mess.Add('16;Comptes généraux;Suppression impossible : ce compte est un compte d''attente, de clôture ou d''ouverture.;W;O;O;O;');
  MsgBox.Mess.Add('17;Comptes généraux;Suppression impossible : ce compte est associé à un paramétrage de TVA ou de TPF.;W;O;O;O;');
  MsgBox.Mess.Add('18;Comptes généraux;Le lettrage a été modifié. Désirez-vous continuer ?;Q;YNC;N;C;');
  MsgBox.Mess.Add('19;Comptes généraux;La ventilation a été modifiée. Désirez-vous continuer ?;Q;YNC;N;C;');
  MsgBox.Mess.Add('20;Comptes généraux;La modification de la nature collective est impossible : ce compte est déjà référencé par des tiers.;W;O;O;O;');
  MsgBox.Mess.Add('21;Comptes généraux;La modification de la nature est impossible : ce compte est compte de contrepartie d''un journal.;W;O;O;O;');
  MsgBox.Mess.Add('22;Comptes généraux;La modification de la nature collective est impossible : ce compte est mouvementé.;W;O;O;O;');
  MsgBox.Mess.Add('23;Comptes généraux;La modification de la caractéristique lettrable va entraîner des modifications d''écritures. Vous devez sélectionner un type de mode de paiement.;E;O;O;O;');
  MsgBox.Mess.Add('24;Comptes généraux;Le traitement n''est pas effectué. Le mode de paiement n''est pas renseigné.;W;O;O;O;');
  MsgBox.Mess.Add('25;Comptes généraux;Vous ne pouvez pas rendre un compte ventilable, pointable et lettrable à la fois.;W;O;O;O;');
  MsgBox.Mess.Add('26;Comptes généraux;Vous ne pouvez pas rendre un compte ventilable et lettrable à la fois.;W;O;O;O;');
  MsgBox.Mess.Add('27;Comptes généraux;Vous ne pouvez pas rendre un compte pointable et lettrable à la fois.;W;O;O;O;');
  MsgBox.Mess.Add('28;Comptes généraux;Vous ne pouvez pas rendre un compte ventilable et pointable à la fois.;W;O;O;O;');
  MsgBox.Mess.Add('Choix d''un mode de paiement');
  MsgBox.Mess.Add('30;Comptes généraux;Suppression impossible : ce compte est mouvementé.;W;O;O;O;');
  MsgBox.Mess.Add('Cet enregistrement n''est plus accessible.');
  MsgBox.Mess.Add('32;Comptes généraux;Suppression impossible : ce compte est référencé dans un guide.;W;O;O;O;');
  MsgBox.Mess.Add('L''accès à ce compte  est interdit.');
  MsgBox.Mess.Add('34;Comptes généraux;Vous devez renseigner le régime fiscal du compte.;W;O;O;O;');
  MsgBox.Mess.Add('RIB d''attente');
  MsgBox.Mess.Add('36;Comptes généraux;Vous n''avez pas créé de RIB associé à ce compte général.Un RIB d''attente va être créé par défaut.;E;O;O;O;');
  MsgBox.Mess.Add('37;Comptes généraux;ne possède aucun code section et n''a pas de section d''attente.;W;O;O;O;');
  MsgBox.Mess.Add('L''axe');
  MsgBox.Mess.Add('39;Comptes généraux;Vous ne pouvez pas modifier la caractéristique lettrable. Un utilisateur vient de saisir une pièce avec ce compte général.;W;O;O;O;');
  MsgBox.Mess.Add('40;Comptes généraux;Vous ne pouvez pas modifier la caractéristique ventilable. Un utilisateur vient de saisir une pièce avec ce compte général.;W;O;O;O;');
  MsgBox.Mess.Add('41;Comptes généraux;Le compte de correspondance n°1 que vous avez saisi n''existe pas !;W;O;O;O;');
  MsgBox.Mess.Add('42;Comptes généraux;Le compte de correspondance n°2 que vous avez saisi n''existe pas !;W;O;O;O;');
  MsgBox.Mess.Add('43;Comptes généraux;Incohérence dans la saisie du couple "Pays - Division territoriale" !;W;O;O;O;');
  MsgBox.Mess.Add('44;Comptes généraux;ATTENTION. Il y a des codes de tables libres qui n''existent pas. Désirez-vous continuer ?;Q;YNC;N;C;');
  MsgBox.Mess.Add('45;Comptes généraux;Ce compte général n''a pas de RIB associé.;E;O;O;O;');
  MsgBox.Mess.Add('46;Comptes généraux;Ce compte général n''a pas de RIB associé. Voulez-vous le créer ?;Q;YNC;Y;C;');
  MsgBox.Mess.Add('47;Comptes généraux;Vous devez renseigner une nature.;W;O;O;O;');
  MsgBox.Mess.Add('Analytique');
  MsgBox.Mess.Add('Ventilable');
  MsgBox.Mess.Add('50;Comptes généraux;La nature "collectif" a été modifiée. Désirez-vous continuer ?;Q;YNC;N;C;');
  MsgBox.Mess.Add('51;Comptes généraux;La modification de la caractéristique Collectif va entraîner des modifications d''écritures. Vous devez sélectionner un auxiliaire.;E;O;O;O;');
  MsgBox.Mess.Add('Choix d''un mode auxiliaire');
  MsgBox.Mess.Add('53;Comptes généraux;Le traitement n''est pas effectué. Le compte auxiliaire n''est pas renseigné.;W;O;O;O;');
  MsgBox.Mess.Add('54;Comptes généraux;La modification de la nature est impossible : ce compte est mouvementé.;W;O;O;O;');
  MsgBox.Mess.Add('55;');

  MsgBox.Mess.Add('56;Comptes généraux;Opération impossible : Vous n''avez pas renseigné le code identifiant du compte général.;W;O;O;O;');
  MsgBox.Mess.Add('57;Comptes généraux;Confirmez-vous la duplication du compte sélectionné ?;Q;YN;N;N;');
  MsgBox.Mess.Add('58;Choix d''un compte général à dupliquer');
  MsgBox.Mess.Add('59;Comptes généraux;Vous ne pouvez pas modifier la caractéristique ventilable. Vous êtes en liaison avec '+RechDom('CPLIENCOMPTABILITE',GetParamsoc ('SO_CPLIENGAMME') , False)+'.;W;O;O;O;');
  MsgBox.Mess.Add('60;Comptes auxiliaires;Attention : La modification du paramètrage de TVA n''est pas répercutée sur les écritures existantes.;W;O;O;O;');
end;

procedure TOM_GENERAUX.BImprimerClick(Sender: TObject);
begin

{$IFNDEF PGIIMMO}
{$IFNDEF IMP}

  // Gestion du multisoc : soldes des comptes repris de la table cumuls via la vue TIERSMS
  if EstTablePartagee('GENERAUX') then
    V_PGI.EnableTableToView := True ;
    {JP 18/01/06 : FQ 17201 : En CWas, Les Booleans FicExport et DeuxPages à True (!) posaient un problème}
    LanceEtat('E','CST','GEN',true, False, False,nil,'G_GENERAL="'+GetField('G_GENERAL')+'"',Ecran.Caption,false,0,'');
    // Impression 2° pages si TID ou TIC
    if IsCptTicTid then
      {JP 18/01/06 : FQ 17201 : En CWas, Les Booleans FicExport et DeuxPages à True (!) posaient un problème}
      LanceEtat('E','CST','GE2',true, False, False,nil,'G_GENERAL="'+GetField('G_GENERAL')+'"',Ecran.Caption,false,0,'');

  // Gestion du multisoc : désactivation de la vue TIERSMS
  V_PGI.EnableTableToView := False ;

  {$ENDIF}
{$ENDIF}

end;

////////////////////////////////////////////////////////////////////////////////
procedure TOM_GENERAUX.PagesChanging(Sender: TObject; var AllowChange: Boolean);
begin
  // FQ 12063 - En entreprise aussi, on interdit la modification
//  if not (ctxPcl in V_PGI.PGIContexte) then Exit;

  AllowChange := ( Trim(GetField('G_GENERAL'))    <> '' ) and
                 ( Trim(GetField('G_NATUREGENE')) <> '' );

  if not AllowChange then
    begin
    if Trim(GetField('G_GENERAL')) = ''
      then SetFocusControl('G_GENERAL')
      else SetFocusControl('G_NATUREGENE') ;
    end;

end;

procedure TOM_GENERAUX.ActivationControl;
var lBoEnabled : Boolean;
begin
    // FQ 12063 - En entreprise aussi, on interdit la modification
  (* if not (ctxPcl in V_PGI.PGIContexte ) then Exit; *)

  lBoEnabled    := (TFFiche(Ecran).TypeAction <> taConsult) and
                   ( Trim( GetField('G_GENERAL') )    <> '' ) and
                   ( Trim( GetField('G_NATUREGENE') ) <> '' ) ;

  SetControlEnabled('G_SENS',        lBoEnabled);
  SetControlEnabled('G_LIBELLE',     lBoEnabled);
  SetControlEnabled('G_ABREGE',      lBoEnabled);
  SetControlEnabled('G_PURGEABLE',   lBoEnabled);
  SetControlEnabled('GCONFIDENTIEL', lBoEnabled);

  if lBoEnabled then
    begin
    SetControlProperty('G_SENS',    'COLOR', clWindow );
    SetControlProperty('G_LIBELLE', 'COLOR', clWindow );
    SetControlProperty('G_ABREGE',  'COLOR', clWindow );
    end
  else
    begin
    SetControlProperty('G_SENS',    'COLOR', clBtnFace );
    SetControlProperty('G_LIBELLE', 'COLOR', clBtnFace );
    SetControlProperty('G_ABREGE',  'COLOR', clBtnFace );
    end;

end;

procedure TOM_GENERAUX.BVentilClick(Sender: TObject);
var Axes : String ;
begin
  if DS.State=dsInsert then
    if Not CodeValide then Exit ;
  NextPrevControl( Ecran ) ;
  if not VH^.AnaCroisaxe then
  begin
    Axes:='' ;
    if (GetField('G_VENTILABLE1')= 'X') then Axes:=Axes+'1';
    if (GetField('G_VENTILABLE2')= 'X') then Axes:=Axes+'2' ;
    if (GetField('G_VENTILABLE3')= 'X') then Axes:=Axes+'3' ;
    if (GetField('G_VENTILABLE4')= 'X') then Axes:=Axes+'4' ;
    if (GetField('G_VENTILABLE5')= 'X') then Axes:=Axes+'5' ;
      if (Axes<>'') then //fb 26/06/2007
    ParamVentil('GE',GetField('G_GENERAL'),Axes, TFFiche(Ecran).TypeAction ,False) ;
  end
  else
  begin
    ParamVentilCroise('GE',GetField('G_GENERAL'),TFFiche(Ecran).TypeAction ,False);
  end;
  if DS.State=dsInsert then
    SetControlEnabled('G_GENERAL',Not ExisteSql('SELECT V_COMPTE FROM VENTIL WHERE V_COMPTE="'+GetField('G_GENERAL')+'" AND V_NATURE LIKE "GE%"'));
end;


{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_GENERAUX.OnClickConfidentiel(Sender : TObject);
//var lQuery : Tquery;
begin
    if (Sender = nil) then exit;
    if (not (DS.State in [dsEdit,dsInsert])) then DS.Edit;
    if (GetCheckBoxState('GCONFIDENTIEL') = cbChecked) then SetField('G_CONFIDENTIEL','1')
    else SetField('G_CONFIDENTIEL','0') ;

    ChangeConfidentiel:=true;
    {FQ18530 YMO 03/04/2007 Maj de l'historique des écritures
    lQuery := nil;
  try

    lQuery := OpenSQL('SELECT DISTINCT E_NUMEROPIECE,E_PERIODE,E_QUALIFPIECE,E_JOURNAL'
                     +' FROM ECRITURE WHERE E_GENERAL="'+GetField('G_GENERAL')+'"', True);

    While not lQuery.EOF do
    begin

      ExecuteSQL('UPDATE ECRITURE SET E_CONFIDENTIEL="8"'//'+GetField('G_CONFIDENTIEL')+'"'
                +' WHERE E_NUMEROPIECE='+lQuery.FindField('E_NUMEROPIECE').AsString
                +' AND E_PERIODE="'+lQuery.FindField('E_PERIODE').AsString+'"'
                +' AND E_QUALIFPIECE="'+lQuery.FindField('E_QUALIFPIECE').AsString+'"'
                +' AND E_JOURNAL="'+lQuery.FindField('E_JOURNAL').AsString+'"');

      lQuery.next;
    end;

  finally
    Ferme(lQuery);
  end;               }
                      
end;

procedure TOM_GENERAUX.OnClickCutOff(Sender : TObject) ;
var
 lBoCutOff : boolean ;
begin
 lBoCutOff := TCheckBox(Sender).Checked ;
 SetControlProperty( 'G_CUTOFFPERIODE', 'Enabled' , lBoCutOff  and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
 SetControlProperty( 'G_CUTOFFECHUE'  , 'Enabled' , lBoCutOff  and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
 SetControlProperty( 'LCUTOFFPER'     , 'Enabled' , lBoCutOff  and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
 SetControlProperty( 'G_CUTOFFCOMPTE' , 'Enabled' , lBoCutOff  and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
 SetControlProperty( 'LCUTOFFCOMPTE'  , 'Enabled' , lBoCutOff  and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
 SetControlProperty( 'LCUTOFFECHUE'   , 'Enabled' , lBoCutOff  and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
 if lBoCutOff and ( GetControlText('G_CUTOFFPERIODE') = '' ) then
  SetControlProperty( 'G_CUTOFFPERIODE' , 'Value' , 'CN'  ) ;
 if not lBoCutOff then
  begin
   SetControlText('G_CUTOFFPERIODE', '' ) ;
   SetControlText('G_CUTOFFECHUE', '-' ) ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 08/06/2005
Modifié le ... :   /  /    
Description .. : FB 15875 - 08/06/2205 - LG -
Suite ........ : Case non cochée = Payable d'avance
Suite ........ : Case cochée =  A terme échu
Mots clefs ... : 
*****************************************************************}
Procedure TOM_GENERAUX.OnClickCutOffEchue(Sender : TObject) ;
begin
 if TCheckBox(Sender).Checked then
  SetControlProperty( 'LCUTOFFECHUE' , 'Caption' , TraduireMemoire('( A terme échu )')  )
   else
    SetControlProperty( 'LCUTOFFECHUE' , 'Caption' , TraduireMemoire('( Payable d''avance )')  ) ;
end ;

function TOM_GENERAUX.CodeValide: Boolean;
begin

  Result := False;

  if DS.State in [dsInsert] then
    begin
    // Code obligatoire
    // FQ 17634 : ajout test GetControlText pour le cas VK_RETURN sur zone G_GENERAL
    if ((Trim(GetField('G_GENERAL')) = '') or (GetControlText('G_GENERAL') = '') )then
      begin
      Pages.ActivePage := TTabSheet(GetControl('PGENERAL',True));
      SetFocusControl('G_GENERAL');
      MsgBox.Execute(2,'','');  // Vous devez renseigner un code
      Exit;
      end;
    // Test unicité
    if Presence('GENERAUX','G_GENERAL',GetField('G_GENERAL')) then
      begin
      Pages.ActivePage := TTabSheet(GetControl('PGENERAL',True));
      SetFocusControl('G_GENERAL');
      MsgBox.Execute(4,'',''); // Le code que vous avez saisi existe déjà. Vous devez le modifier
      Exit;
      end;
    end;

  Result:=True ;

end;

function TOM_GENERAUX.LettrageChanger: Boolean;
begin

  Result := False ;

  // Paramétre de lettrage modifié ?
  if ML = GetField('G_LETTRABLE') then Exit ;

  if AvecMvt then
    begin
    // Si on décoche et mouvements existant alors erreur
    // CA - FQ 17536 - 27/02/2006 : Pas de modification de lettrage si présence de mouvements
    if ( ML='X' ) And ( GetField('G_LETTRABLE') <> 'X' ) (*And (GetField('G_NATUREGENE')<>'DIV')*) then
      begin
      // Vous ne pouvez pas modifier la caractéristique lettrable.
      // Un utilisateur vient de saisir une pièce avec ce compte général.
      MsgBox.Execute(39,'','');
      SetField('G_LETTRABLE',ML);
      Exit ;
    end;

    // Retourne vrai si mouvements exitants + on coche
    if not ( DS.State in [dsBrowse,dsInsert] ) then
      Result := ((ML='-') And (( GetField('G_LETTRABLE') = 'X' ))) ;

  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 20/12/2007
Modifié le ... :   /  /    
Description .. : Avertissement : non-répercussion des modifs de TVA sur les 
Suite ........ : écritures existantes
Mots clefs ... : FQ19946 YMO
*****************************************************************}
Function TOM_GENERAUX.TVAChanger : Boolean ;
Var AeteModifie : Boolean ;
BEGIN
    Result := False;
    AeteModifie := False;

    if (RTVA) <> (GetControlText('G_REGIMETVA')) then
      AeteModifie := True;

    if (TVAE) <> (GetControlText('G_TVAENCAISSEMENT')) then
      AeteModifie := True;

    if (TVA) <> (GetControlText('G_TVA')) then
      AeteModifie := True;

    if (TVASE) <> (GetControlText('G_TVASURENCAISS')) then
      AeteModifie := True;

    if (DS.State = dsInsert) then Exit ;
    Result := AeteModifie;
end;

function TOM_GENERAUX.VentilChanger: Boolean;
begin
  Result       := False ;

  if AvecMvt then
    begin
    // Annulation des modifs si mouvements présents + suppression d'un axe
    if ( ( ( MA1='X' ) And ( GetField('G_VENTILABLE1')<>'X' ) ) Or
         ( ( MA2='X' ) And ( GetField('G_VENTILABLE2')<>'X' ) ) Or
         ( ( MA3='X' ) And ( GetField('G_VENTILABLE3')<>'X' ) ) Or
         ( ( MA4='X' ) And ( GetField('G_VENTILABLE4')<>'X' ) ) Or
         ( ( MA5='X' ) And ( GetField('G_VENTILABLE5')<>'X' ) )
       ) then
      begin
      MsgBox.Execute(40,'','') ;
      RepositionneVentilation ;
      Exit ;
      end;

      if GetParamSocSecur('SO_CPLIENGAMME', '') = 'S1' then
      begin
          if    ( ( MA2='-' ) And ( GetField('G_VENTILABLE2') = 'X' ) ) or
                ( ( MA3='-' ) And ( GetField('G_VENTILABLE3') = 'X' ) ) or
                ( ( MA4='-' ) And ( GetField('G_VENTILABLE4') = 'X' ) ) or
                ( ( MA5='-' ) And ( GetField('G_VENTILABLE5') = 'X' ) ) then
          begin
          MsgBox.Execute(59,'','') ;
          RepositionneVentilation ;
          Exit ;
          end;
      end;

    // Retourne vrai si on a ajouter un axe de ventilation + mouvements présents
    if not ( DS.State in [dsBrowse,dsInsert] ) then
      begin
      result := ( ( MA1='-' ) And ( GetField('G_VENTILABLE1') = 'X' ) ) or
                ( ( MA2='-' ) And ( GetField('G_VENTILABLE2') = 'X' ) ) or
                ( ( MA3='-' ) And ( GetField('G_VENTILABLE3') = 'X' ) ) or
                ( ( MA4='-' ) And ( GetField('G_VENTILABLE4') = 'X' ) ) or
                ( ( MA5='-' ) And ( GetField('G_VENTILABLE5') = 'X' ) ) ;
      end ;

    end ;

end;

procedure TOM_GENERAUX.EnabledTva;
var i : Integer ;
    GBTVA : TGroupBox;
begin
  GBTVA := TGroupBox(GetControl('GBTVA',  True ));
  for i:=0 To GBTVA.ControlCount-1 do
    GBTVA.Controls[i].Enabled := GBTVA.Enabled and ( TFFiche(Ecran).TypeAction <> taConsult ) ;
end;

procedure TOM_GENERAUX.SupprimePreventilRib;
var lStcompte : String ;
begin
  if AvecMvt then Exit ;
  if DS.State <> dsBrowse then Exit ;

  lStcompte := GetField('G_GENERAL') ;
  if lStcompte = '' then Exit ;

  if not IsCptVentilable then
    ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="' + lStcompte + '" AND V_NATURE LIKE "GE%"') ;

  if not IsCptBqe then
    ExecuteSql('DELETE FROM BANQUECP WHERE BQ_GENERAL="' + lStcompte + '"');
end;

function TOM_GENERAUX.TvaTicTidOk: Boolean;
begin
  Result := True ;
  if ( not IsCptTicTid ) or ( IsCptTicTid and not IsCptLettrable ) then Exit;
  if GetField('G_REGIMETVA') = '' then
    begin
    Pages.ActivePage := TTabSheet(GetControl('PINFO',True));
    SetFocusControl('G_REGIMETVA');
    MsgBox.Execute(34,'',''); // Vous devez renseigner le régime fiscal du compte.
    Result:=False ;
    end;
end;

function TOM_GENERAUX.IsMouvemente: Boolean;
var lStCode : String ;
begin
  if ((TFFiche(Ecran).TypeAction=taConsult) or (DS.State=dsInsert)) then
    begin
    Result:=False ;
    Exit ;
    end ;

  lStCode := GetField('G_GENERAL') ;
  {
  Result := ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'             + lStCode + '" AND' +
                      ' ( (EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'  + lStCode + '") )'  +
                      ' OR (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL="' + lStCode + '") ) )'
                      ) ;}

  Result := ExisteSQL('SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL = "' + lStCode + '"');
  if not Result and (GetField('G_VENTILABLE') = 'X') then
    Result := ExisteSQL('SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL = "' + lStCode + '"');

 if not Result then Result := EstCpteAmortissement (lStCode);
end;

procedure TOM_GENERAUX.RepositionneVentilation;
begin
  SetField( 'G_VENTILABLE',  MA0 );
  SetField( 'G_VENTILABLE1', MA1 );
  SetField( 'G_VENTILABLE2', MA2 );
  SetField( 'G_VENTILABLE3', MA3 );
  SetField( 'G_VENTILABLE4', MA4 );
  SetField( 'G_VENTILABLE5', MA5 );
end;

function  TOM_GENERAUX.VerifiNatureValide : Boolean ;
begin
  Result := False ;
  if Not AvecMvt then Exit ;
  if MNat = GetField('G_NATUREGENE') then Exit ;

  if (MNat='COC') Or (MNat='COD') Or (MNat='COF') Or (MNat='COS') then
    begin
    if pos( GetField('G_NATUREGENE'), 'BQE;CAI;CHA;DIV;EXT;IMO;PRO;TIC;TID') > 0 then
      begin
      SetField('G_NATUREGENE', MNat );
      {$IFNDEF EAGLCLIENT}
        SetControlText('G_NATUREGENE', MNat );
      {$ENDIF}
      Result := True ;
      Exit;
      end;
    end
  else
  (*
    if (MNat<>'DIV') then
      begin
      SetField('G_NATUREGENE', MNat );
      {$IFNDEF EAGLCLIENT}
        SetControlText('G_NATUREGENE', MNat );
      {$ENDIF}
      Result := True ;
      Exit;
      end;
    *)
end;


{$IFDEF COMPTA}
procedure TOM_GENERAUX.BBalGenAnaClick(Sender: TObject);
var
  ACritEdt : ClassCritEdt;
  D1, D2 : TDateTime;
  Etab : string;
  Axe : string;
begin
  Axe:='A1' ;
  if (GetCheckBoxState('G_VENTILABLE1')= cbChecked) then Axe:='A1' else
  if (GetCheckBoxState('G_VENTILABLE2')= cbChecked) then Axe:='A2' else
  if (GetCheckBoxState('G_VENTILABLE3')= cbChecked) then Axe:='A3' else
  if (GetCheckBoxState('G_VENTILABLE4')= cbChecked) then Axe:='A4' else
  if (GetCheckBoxState('G_VENTILABLE5')= cbChecked) then Axe:='A5' else Exit;
  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
    If VH^.Entree.Code=VH^.Suivant.Code Then begin D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; end;
    ACritEdt.CritEdt.Cpt1  := GetField('G_GENERAL');
    ACritEdt.CritEdt.Cpt2  := GetField('G_GENERAL');
    ACritEdt.CritEdt.SCpt1 := '';
    ACritEdt.CritEdt.SCpt2 := '';
    ACritEdt.CritEdt.Date1:=D1 ; ACritEdt.CritEdt.Date2:=D2 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ; ACritEdt.CritEdt.DateFin:=ACritEdt.CritEdt.Date2 ;
    Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
    ACritEdt.CritEdt.BAL.Axe := Axe;
    TheData := ACritEdt;
{$IFNDEF CMPGIS35}
    CPLanceFiche_BalanceGenAnal;
{$ENDIF}
  finally
    ACritEdt.Free;
    TheData := nil;
  end;
end;

procedure TOM_GENERAUX.BBALGENAUXClick(Sender: TObject);
var
  ACritEdt : ClassCritEdt;
  D1, D2 : TDateTime;
  Etab : string;
begin
  ACritEdt := ClassCritEdt.Create;
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
  If VH^.Entree.Code=VH^.Suivant.Code Then begin D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; end;
  ACritEdt.CritEdt.Cpt1  := GetField('G_GENERAL');
  ACritEdt.CritEdt.Cpt2  := GetField('G_GENERAL');
  ACritEdt.CritEdt.SCpt1 := '';
  ACritEdt.CritEdt.SCpt2 := '';
  ACritEdt.CritEdt.Date1:=D1 ; ACritEdt.CritEdt.Date2:=D2 ;
  ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ; ACritEdt.CritEdt.DateFin:=ACritEdt.CritEdt.Date2 ;
  Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
  TheData := ACritEdt;
{$IFNDEF CMPGIS35}
  CPLanceFiche_BalanceGenAuxi;
{$ENDIF}
  ACritEdt.Free;
  TheData := nil;
end;
{$ENDIF}

procedure TOM_GENERAUX.BCumulClick(Sender: TObject);
begin
{$IFDEF COMPTA}
    CumulCpteMensuel(fbGene,GetField('G_GENERAL'),GetField('G_LIBELLE'),VH^.Entree);
{$ENDIF}
end;

procedure TOM_GENERAUX.BImmoClick(Sender: TObject);
begin
{$IFDEF AMORTISSEMENT}
  AMLanceFiche_ListeDesImmobilisations ( GetField('G_GENERAL'), False , taConsult) ;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/04/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_GENERAUX.BGCycleRevisionClick(Sender: TObject);
{$IFDEF COMPTA}
{$IFNDEF CCMP}
var lStPredefini : string;
{$ENDIF}
{$ENDIF}
begin
{$IFDEF COMPTA}
{$IFNDEF CCMP}
  lStPredefini := GetColonneSQL('CREVCYCLE', 'CCY_PREDEFINI', 'CCY_CODECYCLE = "' + GetField('G_CYCLEREVISION') + '"');
{$IFNDEF CMPGIS35}
  CPLanceFiche_CRevParamCycle( GetField('G_CYCLEREVISION')+ ';' + lStPredefini, TaConsult );
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOM_GENERAUX.BInsertClick(Sender: TObject);
begin
  if ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then
    TFFiche(Ecran).BInsertClick(Sender);
end;

procedure TOM_GENERAUX.BJustifClick(Sender: TObject);
begin
{ FQ 17458 BVE 12.04.07 }
(* rony 14/04/97 JustifSolde(GetField('G_GENERAL'),fbGene);*)
     {$IFNDEF COMPTAPAIE}
       {$IFNDEF PGIIMMO}
         {$IFNDEF IMP}
           {$IFNDEF CCADM}
           JustifSolde(GetField('G_GENERAL'),fbGene);
           {$ENDIF CCADM}
         {$ENDIF IMP}
       {$ENDIF PGIIMMO}
     {$ENDIF COMPTAPAIE}
{ END FQ 17458 }
end;

procedure TOM_GENERAUX.BMenuZoomClick(Sender: TObject);
begin
  PopZoom(BMenuZoom,POPZ) ;
end;


function TOM_GENERAUX.BqeCaiChanger: Boolean;
begin
  Result := False ;
  if DS.State in [dsBrowse,dsInsert] then Exit ;
  if (MNat='BQE') or (MNat='CAI') then
    if (GetField('G_NATUREGENE')='DIV') then
      Result:=True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lek
Créé le ...... : 23/09/2005
Modifié le ... : 23/09/2005
Description .. : Cette fonction vérifie l'existance de l'enregistrement de
Suite ........ : banque pour un compte géné de type banque
Suite ........ : Elle rend vrai si l'enregistrement existe
Suite ........ : sinon elle rend faux
Mots clefs ... :
*****************************************************************}
function  TOM_GENERAUX.ExisteEnregBanque:boolean; //Lek 23/09/05 FQ16663
begin
  Result:=ExisteSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_GENERAL = "' +
                GetField('G_GENERAL') + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"') ;
end;

procedure TOM_GENERAUX.BRIBClick(Sender: TObject);
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
var okok   : boolean ;
    LeMode : TActionFiche ;
    Existe : Boolean ;
{$ENDIF}
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
  if DS.State=dsInsert then
    if Not CodeValide then Exit ;
  LeMode:=TFFiche(Ecran).TypeAction ;
  if IsCptBqe then begin
    okok:=ExisteEnregBanque; {Lek 23/09/05 FQ16663}
    if ((LeMode=taConsult) and (Not okok)) then begin
      Msgbox.Execute(5,'','') ;  // Vous devez associer un compte bancaire à ce compte.
      Exit ;
    end;
    if ((LeMode<>taConsult) and (Not okok)) then LeMode:=taCreatOne ;
    if OkOk And (LeMode in [taCreat..taCreatOne]) and (DS.state<>dsInsert) then Exit ;
    if (OkOk And (LeMode in [taCreat..taCreatOne]) and (DS.State=dsInsert)) then LeMode:=taModif; // FQ 12543
{$IFNDEF CCADM}
    FicheBanqueCP(GetField('G_GENERAL'), LeMode, 0, V_PGI.NoDossier);
{$ENDIF CCADM}
    Existe :=ExisteEnregBanque; {Lek 23/09/05 FQ16663}
    SetControlEnabled('G_GENERAL',Not Existe);
    if Not Existe then CreerRibDefaut;
    end
  else begin
    OkOk := ExisteSQL('SELECT R_AUXILIAIRE FROM RIB WHERE R_AUXILIAIRE="'+GetField('G_GENERAL')+'"') ;
    if ((LeMode=taConsult) And  (Not OkOk)) then begin
      Msgbox.Execute(45,'','') ;  // Ce compte général n'a pas de RIB associé.
      Exit ;
    end;
    if ((LeMode<>taConsult) And (Not OkOk)) then
      // Ce compte général n'a pas de RIB associé. Voulez-vous le créer ?
      if Msgbox.Execute(46,'','')<>mrYes then Exit
                                         else LeMode:=taCreatOne ;
      if OkOk And (LeMode in [taCreat..taCreatOne]) then LeMode:=taModif ;
      FicheRIB_AGL(GetField('G_GENERAL'),LeMode,False,'',FALSE) ;
      SetControlEnabled('G_GENERAL',  Not(ExisteSQL('SELECT R_AUXILIAIRE FROM RIB WHERE R_AUXILIAIRE="'+GetField('G_GENERAL')+'"')) ) ;
  end;
{$ENDIF}
{$ENDIF}
end;

procedure TOM_GENERAUX.BZecrimvtClick(Sender: TObject);
begin
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
  {$IFNDEF COMPTAPAIE}
    ZoomEcritureMvt( GetField('G_GENERAL'), fbGene,'MULMMVTS' ) ;
  {$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF CCADM}
end;

function TOM_GENERAUX.CaracteristiqueIncompatible(var j: Byte): Boolean;
begin
  Result:=False ;

  // MAJ champ G_VENTILABLE
  MajChampsVentilAxe ;

  // Test compte tri-caractérisque
  if IsCptVentilable And IsCptLettrable And IsCptPointable then
     begin
     Result := true ;
     j:=1 ;
     Exit ;
     end;

  // Test compte double-caractérisque
  if IsCptLettrable And IsCptPointable then
     begin
     Result:=true ;
     j:=3 ;
     Exit ;
     end;

end;

procedure TOM_GENERAUX.ChercheEtImportCptePlanRef;
var St    : String ;
    QPRef : TQuery;
    stPrefixe : string;
begin

  if DS.State <> dsInsert then Exit ;
  if VH^.NumPlanRef=0 then Exit ;
  if Not Import then Exit ;

  St:=Bourreless(GetField('G_GENERAL'),fbGene);
  if St='' then Exit ;

  if Presence('GENERAUX','G_GENERAL',GetField('G_GENERAL')) then
    begin
    Pages.ActivePage := TTabSheet(GetControl('PGENERAL',True));
    SetFocusControl('G_GENERAL');
    MsgBox.Execute(4,'','');  // Le code que vous avez saisi existe déjà. Vous devez le modifier.
    Exit ;
    end;

  Import := False;
  if ctxPCL in V_PGI.PGIContexte then
   QPref := OpenSQL('SELECT * FROM GENERAUXREF WHERE GER_NUMPLAN='+IntToStr(VH^.NumPlanRef)+' AND GER_GENERAL="'+St+'"',True)
  else
  { FQ 20958 BVE 16.07.07 }
    QPref := GetGoodPlanRef(St, IntToStr(VH^.NumPlanRef));
  { END FQ 20958 }
  if (QPRef <> nil) and (not QPRef.EOF) then
    begin
    {$IFNDEF PGIIMMO}
    {$IFNDEF CCADM}
      if FicheImportPlanRef(VH^.NumPlanRef,St)=1 then
        begin
        if ctxPCL in V_PGI.PGIContexte then stPrefixe := 'GER' else stPrefixe := 'PR';
        SetField('G_LIBELLE',        QPRef.FindField(stPrefixe+'_LIBELLE').AsString);
        SetField('G_ABREGE',         QPRef.FindField(stPrefixe+'_ABREGE').AsString);
        SetField('G_CENTRALISABLE',  QPRef.FindField(stPrefixe+'_CENTRALISABLE').AsString);
        SetField('G_SOLDEPROGRESSIF',QPRef.FindField(stPrefixe+'_SOLDEPROGRESSIF').AsString);
        SetField('G_SAUTPAGE',       QPRef.FindField(stPrefixe+'_SAUTPAGE').AsString);
        SetField('G_TOTAUXMENSUELS', QPRef.FindField(stPrefixe+'_TOTAUXMENSUELS').AsString);
        SetField('G_COLLECTIF',      QPRef.FindField(stPrefixe+'_COLLECTIF').AsString);
        if not (ctxPCL in V_PGI.PGIContexte) then
          SetField('G_BLOCNOTE',       QPRef.FindField(stPrefixe+'_BLOCNOTE').AsString);
        SetField('G_SENS',           QPRef.FindField(stPrefixe+'_SENS').AsString);
        SetField('G_LETTRABLE',      QPRef.FindField(stPrefixe+'_LETTRABLE').AsString);
        SetField('G_POINTABLE',      QPRef.FindField(stPrefixe+'_POINTABLE').AsString);
        SetField('G_VENTILABLE1',    QPRef.FindField(stPrefixe+'_VENTILABLE1').AsString);
        SetField('G_VENTILABLE2',    QPRef.FindField(stPrefixe+'_VENTILABLE2').AsString);
        SetField('G_VENTILABLE3',    QPRef.FindField(stPrefixe+'_VENTILABLE3').AsString);
        SetField('G_VENTILABLE4',    QPRef.FindField(stPrefixe+'_VENTILABLE4').AsString);
        SetField('G_VENTILABLE5',    QPRef.FindField(stPrefixe+'_VENTILABLE5').AsString);
        SetField('G_NATUREGENE',     QPRef.FindField(stPrefixe+'_NATUREGENE').AsString);
        end;
        { FQ 22217 BVE 17.01.08 }
        SetFocusControl('G_GENERAL');
        { END 22217 BVE 17.01.08 }
    {$ENDIF CCADM}    
    {$ENDIF}
    end;

  Ferme(QPRef);

end;

function TOM_GENERAUX.GetGoodPlanRef(var Compte : string ; NumPlanRef : string) : TQuery;
var SQL : string;
begin
  result := nil;
  SQL := 'SELECT 1 FROM PLANREF ' +
         'WHERE PR_NUMPLAN = ' + NumPlanRef + ' ' +
         'AND PR_COMPTE = "' + Compte + '"';
  while not ExisteSQL(SQL) do
  begin
     if Length(Compte) = 1 then Exit;
     SetLength(Compte,Length(Compte) - 1);
     SQL := 'SELECT 1 FROM PLANREF ' +
            'WHERE PR_NUMPLAN = ' + NumPlanRef + ' ' +
            'AND PR_COMPTE = "' + Compte + '"';
  end;
  // Il existe bien un plan paramétré. Reste à savoir si c'est un prédéfini Cegid ou un Standard.
  SQL := 'SELECT ##TOP 1## * FROM PLANREF ' +
         'WHERE PR_NUMPLAN = ' + NumPlanRef + ' ' +
         'AND PR_COMPTE = "' + Compte + '"' +
         'ORDER BY PR_PREDEFINI DESC';
  result := OpenSQL(SQL,true);
end;

function TOM_GENERAUX.CollectifChanger: Boolean;
begin
  Result:=False ;
  if DS.State in [dsBrowse,dsInsert] then Exit ;
  if isCptNatColl then
     Result := GetField('G_NATUREGENE') <> MNat ;
end;

function TOM_GENERAUX.CollectifChangerChaPro: Boolean;
begin
  result := False ;

  if not VH^.CPIFDEFCEGID then exit ;
  if DS.State in [dsBrowse,dsInsert] then Exit ;

  if (GetField('G_NATUREGENE')='CHA') or (GetField('G_NATUREGENE')='PRO') Then
    begin
    // Si mouvements existant >> retourne Paramètre collectif modifié
    if AvecMvt then
      result := GetField('G_COLLECTIF') <> MColl ;
   end ;

end;

procedure TOM_GENERAUX.MajCollChaProEcriture( Cpte, Aux : String ) ;
var LeWhere,LeSet,Mp,Mr : String ;
    Q                   : TQuery ;
    OkLet               : Boolean ;
begin
  if not VH^.CPIFDEFCEGID then exit ;

  LeSet:='' ;
  if Aux<>'' then
    begin
    OkLet:=FALSE ;
    Mr:='' ;
    Mp:='' ;
    Q := OpenSQL('SELECT T_LETTRABLE, T_MODEREGLE FROM TIERS WHERE T_AUXILIAIRE="'+Aux+'" ',TRUE) ;
    If Not Q.Eof Then
      begin
      OkLet := Q.Fields[0].AsString = 'X' ;
      Mr    := Q.Fields[1].AsString ;
      end ;
    Ferme(Q) ;

    Q := OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+MR+'" ',TRUE) ;
    If Not Q.Eof Then
      Mp:=Q.Fields[0].AsString ;
    Ferme(Q) ;

    If OkLet Then
      LeSet:=', E_MODEPAIE="' + Mp + '", E_DATEECHEANCE=E_DATECOMPTABLE, E_ETATLETTRAGE="AL", E_NUMECHE=1, E_ECHE="X" ' ;
    end
  else
    LeSet := ', E_MODEPAIE="", E_ETATLETTRAGE="RI", E_NUMECHE=0, E_ECHE="-", E_COUVERTURE=0, E_COUVERTUREDEV=0, E_LETTRAGEDEV="-" ' ;

  LeWhere := ' WHERE E_GENERAL="' + Cpte + '" ' ;
  LeSet := LeSet + ' , E_PAQUETREVISION=1 ' ;

  ExecuteSql('UPDATE ECRITURE SET E_AUXILIAIRE="' + Aux + '" ' + LeSet + LeWhere ) ;

end;


procedure TOM_GENERAUX.CreerRibDefaut;
var QLoc : TQuery ;
    St   : String ;
begin
  // Vous n'avez pas créé de RIB associé à ce compte général.Un RIB d'attente va être créé par défaut.
  MsgBox.Execute(36,'','') ;

  // Chargement code banque
  QLoc:=OpenSql('Select PQ_BANQUE From BANQUES',True) ;
  St:='';
  if Not QLoc.Eof then
    St:=QLoc.Fields[0].AsString ;
  Ferme(QLoc) ;

  {05/04/05 : FQ 10242 : Le code société n'est pas renseigné, ce qui pose un problème en Trésorerie
   31/07/06 : Tréso multi sociétés : ajout du nodossier et de la nature de compte Pour les comptes bancaires ...}
  if GetField('G_NATUREGENE') = 'BQE' then begin
    {$IFDEF COMPTA}
    ExecuteSql('INSERT INTO BANQUECP ( BQ_CODE, BQ_GENERAL, BQ_BANQUE, BQ_LIBELLE, BQ_DEVISE, BQ_PAYS, ' +
               'BQ_SOCIETE, BQ_NODOSSIER, BQ_NATURECPTE ) ' +
               'VALUES ("' + Copy(VarToStr(GetField('G_GENERAL')), 1, 12) + StrRight(V_PGI.NoDossier, 5) + '", "' +
                             GetField('G_GENERAL') + '", "' + St + '", "' + MsgBox.Mess[35] + '", "' +
                             V_PGI.DevisePivot + '", "' + CodePaysDeIso(GetParamSocSecur('SO_PAYSLOCALISATION','')) +
                             '", "' + V_PGI.CodeSociete + '", "' + V_PGI.NoDossier + '", "' + tcb_Bancaire + '")');
  {... s'il s'agit d'un compte courant, création de l'enregistrement "complet"}
    {$ENDIF COMPTA}
  end
  else begin
    st := V_PGI.SchemaName + ' vers ' + '????';
    {$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
    ExecuteSql('INSERT INTO BANQUECP (BQ_CODE, BQ_GENERAL, BQ_BANQUE, BQ_AGENCE, BQ_LIBELLE, BQ_DEVISE, BQ_PAYS, ' +
               'BQ_SOCIETE, BQ_NODOSSIER, BQ_NATURECPTE, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, '+
               'BQ_CODECIB, BQ_DOMICILIATION ) ' +
               'VALUES ("' + Copy(VarToStr(GetField('G_GENERAL')), 1, 12) + StrRight(V_PGI.NoDossier, 5) + '", "' +
                             GetField('G_GENERAL') + '", "' + GetBanqueCourant + '", "' + GetAgenceCourant +
                             '", "' + st + '", "' + V_PGI.DevisePivot + '", "' +
                             CodePaysDeIso(GetParamSocSecur('SO_PAYSLOCALISATION','')) + '", "' +
                             V_PGI.CodeSociete + '", "' + V_PGI.NoDossier + '", "' + tcb_Courant + '", "' +
                             CODEATTENTE5 + '", "' + CODEATTENTE5 + '", "' + CODEATTENTE5 + '", "53", "' +
                             GetCibCourant + '", "' + TraduireMemoire('Agence comptes spéciaux') + '")');
{$ENDIF}
    {$ENDIF COMPTA}
  end;
  // ??
  SetControlEnabled('G_GENERAL',False);
end;

procedure TOM_GENERAUX.Ergo35;
var i  : integer  ;
    CC : TControl ;
begin
  {JP 30/09/05 : FQ 16569 : On cache les zones si les paramsoc sont décochés}
  SetControlVisible('G_CORRESP2',  GetParamSocSecur('SO_CORSGE2', True, True));
  SetControlVisible('TG_CORRESP2', GetParamSocSecur('SO_CORSGE2', True, True));
  SetControlVisible('G_CORRESP1',    GetParamSocSecur('SO_CORSGE1', True, True));
  SetControlVisible('TG_CORRESP1',   GetParamSocSecur('SO_CORSGE1', True, True));

  // gestion des tables libres     ???
  if TTabSheet(GetControl('ZL',True)).TabVisible then
    for i:=0 to 9 do
      begin
      CC := GetControl('G_TABLE'+IntTostr(i),True);
      if CC<>Nil then
        if GetControlText(CC.Name) <> '' then
          if not LookUpValueExist(CC)
            then SetControlText(CC.Name,'')
            else SetControlText('H_'+IntToStr(i),'');
      end;

end;

function TOM_GENERAUX.IsCptVentilAxe( Quelaxe : Byte ) : Boolean;
begin
  Result := False ;
  Case QuelAxe of
    1: if (MA1='-') And (GetField('G_VENTILABLE1') = 'X') then
         Result:=True ;
    2: if (MA2='-') And (GetField('G_VENTILABLE2') = 'X') then
         Result:=True ;
    3: if (MA3='-') And (GetField('G_VENTILABLE3') = 'X') then
         Result:=True ;
    4: if (MA4='-') And (GetField('G_VENTILABLE4') = 'X') then
         Result:=True ;
    5: if (MA5='-') And (GetField('G_VENTILABLE5') = 'X') then
         Result:=True ;
  end;
end;

procedure TOM_GENERAUX.FormateLesMontants;
begin
  ChangeMask( THNumEdit(GetControl('G_SOLCREP',True)),        V_PGI.OkdecV, '' );
  ChangeMask( THNumEdit(GetControl('G_SOLCREE',True)),        V_PGI.OkdecV, '' );
  ChangeMask( THNumEdit(GetControl('G_SOLCRES',True)),        V_PGI.OkdecV, '' );
  SetControlProperty('G_RISQUE', 'DISPLAYFORMAT', StrfMask(V_PGI.OkdecV,'',True) );
  SetControlProperty('G_PLAFOND','DISPLAYFORMAT', StrfMask(V_PGI.OkdecV,'',True) );
end;

procedure TOM_GENERAUX.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Vide : boolean ;
begin
  Vide:=(Shift=[]) ;
  if Vide then
  begin
    Case Key of
      VK_RETURN : begin
                  if Ecran.ActiveControl is TCustomMemo then Exit ;
                  NextControl(Ecran);
                  end;
    else
      TFFiche(Ecran).FormKeyDown(Sender,Key,Shift);
    end;
  end;
end;

procedure TOM_GENERAUX.G_CODEPOSTALDblClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  VerifCodePostal(DS,THEdit(GetControl('G_CODEPOSTAL',True)),THEdit(GetControl('G_VILLE',True)),FALSE) ;
{$ELSE}
  VerifCodePostal(DS,THDBEdit(GetControl('G_CODEPOSTAL',True)),ThDBEdit(GetControl('G_VILLE',True)),FALSE) ;
{$ENDIF}
end;

procedure TOM_GENERAUX.G_DIVTERRITDblClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  PaysRegion(THValComboBox(GetControl('G_PAYS',True)),THEdit(GetControl('G_DIVTERRIT',True)),True) ;
{$ELSE}
  PaysRegion(THDBValComboBox(GetControl('G_PAYS',True)),TDBEdit(GetControl('G_DIVTERRIT',True)),True) ;
{$ENDIF}
end;

procedure TOM_GENERAUX.G_PAYSChange(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  PaysRegion(THValComboBox(GetControl('G_PAYS',True)),THEdit(GetControl('G_DIVTERRIT',True)),False);
{$ELSE}
  PaysRegion(THDBValComboBox(GetControl('G_PAYS',True)),TDBEdit(GetControl('G_DIVTERRIT',True)),False);
{$ENDIF}
end;

procedure TOM_GENERAUX.G_VILLEDblClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  VerifCodePostal(DS,THEdit(GetControl('G_CODEPOSTAL',True)),THEdit(GetControl('G_VILLE',True)),FALSE) ;
{$ELSE}
  VerifCodePostal(DS,ThDBEdit(GetControl('G_CODEPOSTAL',True)),ThDBEdit(GetControl('G_VILLE',True)),FALSE) ;
{$ENDIF}
end;


procedure TOM_GENERAUX.InitModifEnSerie(StModif: string);
var szName, St    : string;
    Champ, Valeur : string;
    i             : integer;
    Bouton        : TToolBarButton97;
    C             : TControl;
begin

  if not modifEnSerie then Exit ;

  if DS.State = dsBrowse then DS.Edit ;

  // MAJ des champs modifiés
  while StModif <> '' do
    begin
    // Calcul Champ / Valeur
    St := ReadTokenSt( StModif ) ;
    // controle du cutoff
    if ( Pos('G_CUTOFF',St) > 0  ) and not FBoCutOff then continue ;

    i  := Pos('=',St);
    if i > 0 then Champ := Trim( Copy(St,1,i-1) );
    i := Pos('"',St);
    if i > 0 then St := Trim( Copy(St,i+1,Length(St)) );
    i := Pos('"',St);
    if i > 0 then Valeur := Trim( Copy(St,1,i-1) );

    // Affectation valeur / modif couleur
    if Champ = 'G_CONFIDENTIEL' then
      C := GetControl('GCONFIDENTIEL', True)
    else
      C := GetControl(Champ, True);

    {$IFDEF EAGLCLIENT}
      if (C is TCheckBox) or (C is THValComboBox) or (C is TEdit) then
    {$ELSE}
      if (C is TDBCheckBox) or (C is THDBValComboBox) or (C is TDBEdit) or
         (C is TCheckBox) then
    {$ENDIF}
    begin
      if Champ = 'G_CONFIDENTIEL' then
      begin
        TCheckBox(GetControl('GCONFIDENTIEL', True)).Checked := (Valeur = '1');
      end;
      SetField( Champ, Valeur ) ;
      TEdit(C).Font.Color := clRed ;
      //SetControlText( Champ , Valeur ) ; ??
    end;
  end;

  // restriction d'utilisation des fonctinnalités de la fiche
  for i:=0 to TToolWindow97(GetControl('PBOUTON',True)).ControlCount-1 do
    if TToolWindow97(GetControl('PBOUTON',True)).Controls[i] is TToolBarButton97 then
    begin
    Bouton := TToolBarButton97(TToolWindow97(GetControl('PBOUTON',True)).Controls[i]) ;
    szName := Uppercase(Bouton.Name) ;
    if ((szName<>'BVALIDER') and (szName<>'BFERME') and (szName<>'BHELPBTN')) then
      Bouton.Enabled := False ;
    end;

end;

//SG6 14.03.05 FQ 15483 > on ne met plus a jour le plan de référence
{procedure TOM_GENERAUX.MajDuPlanReference;
var lStCompte : String ;
    lQPRef    : TQuery ;
    lTobMere  : TOB ;
    lTobFille : TOB ;
    lInCpt    : Integer ;
begin

  if VH^.NumPlanRef=0 then Exit ;

  lStCompte := Bourreless( GetField('G_GENERAL'), fbGene ) ;

  // Récupération des plans de référence existants
  lTobMere := Tob.Create('V_PLANS', nil, -1) ;
  lQPRef := OpenSQL('SELECT * FROM PLANREF WHERE PR_NUMPLAN=' + IntToStr( VH^.NumPlanRef )
                                          +' AND PR_COMPTE="' + lStCompte + '"' , True);
  lTobMere.LoadDetailDB( 'PLANREF', '', '', lQPRef, False ) ;
  Ferme( lQPRef );

  // Si aucun, création d'un plan par défaut
  if lTobMere.Detail.Count = 0 then
    begin
    lTobFille := Tob.Create('PLANREF', lTobMere, -1 ) ;
    lTobFille.InitValeurs ;
    lTobFille.PutValue('PR_NUMPLAN',      VH^.NumPlanRef) ;
    lTobFille.PutValue('PR_COMPTE',       lStCompte) ;
    lTobFille.PutValue('PR_PREDEFINI',    'CEG') ;
    end ;

  // MAJ des données du plan avec celles du compte
  for lInCpt := 0 to lTobMere.Detail.Count - 1 do
    begin
    lTobFille := lTobMere.Detail[ lInCpt ] ;
    lTobFille.PutValue('PR_LIBELLE',         GetField('G_LIBELLE') ) ;
    lTobFille.PutValue('PR_ABREGE',          GetField('G_ABREGE') ) ;
    lTobFille.PutValue('PR_NATUREGENE',      GetField('G_NATUREGENE') ) ;
    lTobFille.PutValue('PR_CENTRALISABLE',   GetField('G_CENTRALISABLE') ) ;
    lTobFille.PutValue('PR_SOLDEPROGRESSIF', GetField('G_SOLDEPROGRESSIF') ) ;
    lTobFille.PutValue('PR_SAUTPAGE',        GetField('G_SAUTPAGE') ) ;
    lTobFille.PutValue('PR_TOTAUXMENSUELS',  GetField('G_TOTAUXMENSUELS') ) ;
    lTobFille.PutValue('PR_COLLECTIF',       GetField('G_COLLECTIF') ) ;
    lTobFille.PutValue('PR_SENS',            GetField('G_SENS') ) ;
    lTobFille.PutValue('PR_LETTRABLE',       GetField('G_LETTRABLE') ) ;
    lTobFille.PutValue('PR_POINTABLE',       GetField('G_POINTABLE') ) ;
    lTobFille.PutValue('PR_VENTILABLE1',     GetField('G_VENTILABLE1') ) ;
    lTobFille.PutValue('PR_VENTILABLE2',     GetField('G_VENTILABLE2') ) ;
    lTobFille.PutValue('PR_VENTILABLE3',     GetField('G_VENTILABLE3') ) ;
    lTobFille.PutValue('PR_VENTILABLE4',     GetField('G_VENTILABLE4') ) ;
    lTobFille.PutValue('PR_VENTILABLE5',     GetField('G_VENTILABLE5') ) ;
    lTobFille.PutValue('PR_BLOCNOTE',        GetField('G_BLOCNOTE') ) ;
    end ;

  // MAJ de la base
  lTobMere.InsertOrUpdateDB ;

end;
}

function TOM_GENERAUX.OkConfidentiel: Boolean;
begin
  Result := False ;
  if ((V_PGI.Confidentiel = '0') and (GetCheckBoxState('GCONFIDENTIEL') = cbChecked)) then
    begin
    MessageAlerte(MsgBox.Mess[33]);   // L'accès à ce compte est interdit.
    Ecran.Close; // Anciennement PostMessage(Ecran.Handle,WM_CLOSE,0,0) ;
    Exit ;
    end;
  SetControlVisible('GBCONFIDENTIEL', (V_PGI.Confidentiel>'0') ) ;
  SetControlEnabled('GBCONFIDENTIEL', V_PGI.Superviseur ) ;
  Result:=True ;
end;

procedure  TOM_GENERAUX.InitComposants ;
begin
  // FICHE
  PopZ         := TPopUpMenu(GetControl('POPZ',True));
  BMenuZoom    := TBitBtn(GetControl('BMENUZOOM',True));
  //SG6
  if VH^.AnaCroisaxe then
  begin
    SetControlVisible('G_VENTILABLE',True);
    SetcontrolEnabled('G_VENTILABLE1',False);
    SetcontrolEnabled('G_VENTILABLE2',False);
    SetcontrolEnabled('G_VENTILABLE3',False);
    SetcontrolEnabled('G_VENTILABLE4',False);
    SetcontrolEnabled('G_VENTILABLE5',False);
    //SetcontrolEnabled('BVentil',False);
  end;


  // Onglet Comptable
{$IFDEF EAGLCLIENT}
  G_GENERAL       := THEdit(GetControl('G_GENERAL', True));
  G_LETTRABLE     := TCheckBox(GetControl('G_LETTRABLE',True));
  G_COLLECTIF     := TCheckBox(GetControl('G_COLLECTIF',True));
  SetControlVisible('BIMPRIMER', True); {JP 18/01/06 : FQ 17201}
{$ELSE}
  G_GENERAL       := THDBEdit(GetControl('G_GENERAL', True));
  G_LETTRABLE     := TDBCheckBox(GetControl('G_LETTRABLE',True));
  G_COLLECTIF     := TDBCheckBox(GetControl('G_COLLECTIF',True));
{$ENDIF}
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 04/01/2005
Modifié le ... :   /  /
Description .. : - FB 15171 - LG - 04/01/2005 - il etait possible de faire
Suite ........ : nouveau alors que le concept applicatif ne me permet pas
Suite ........ : de créer de nouveaux comptes.
Mots clefs ... :
*****************************************************************}
procedure TOM_GENERAUX.InitEvenements ;
begin
  TFFiche(Ecran).OnAfterFormShow := Self.OnAfterFormShow;

  // Evènement fiches
  Ecran.OnKeyDown   := FormKeyDown;
  Pages             := TPageControl(GetControl('PAGES',True));
  Pages.OnChanging  := PagesChanging;

  BMenuZoom.OnClick := BMenuZoomClick;

  // Evénements des boutons
  TToolBarButton97(GetControl('BIMPRIMER',True)).OnClick       := BImprimerClick;
  TToolBarButton97(GetControl('BVENTIL',True)).OnClick         := BVentilClick;
  TToolBarButton97(GetControl('BRIB',True)).OnClick            := BRIBClick;
  TToolBarButton97(GetControl('BZECRIMVT',True)).OnClick       := BZecrimvtClick;
  TToolBarButton97(GetControl('BCUMUL',True)).OnClick          := BCumulClick;
  TToolBarButton97(GetControl('BJUSTIF',True)).OnClick         := BJustifClick;
  {$IFDEF COMPTA}
  TToolBarButton97(GetControl('BBALGENANA',True)).OnClick      := BBalGenAnaClick;
  TToolBarButton97(GetControl('BBALGENAUX',True)).OnClick      := BBALGENAUXClick;
  {$ENDIF}
  TToolBarButton97(GetControl('BIMMO',True)).OnClick           := BImmoClick;
  TToolBarButton97(GetControl('BINSERT',true)).OnClick         := BInsertClick;
  TToolBarButton97(GetControl('BGCYCLEREVISION', True)).OnClick := BGCycleRevisionClick;
  TToolBarButton97(GetControl('BCOPIER',True)).OnClick         := BCopierClick;

  TCheckBox(GetControl('G_CUTOFF',true)).OnClick          := OnClickCutOff ;
  TCheckBox(GetControl('G_CUTOFFECHUE',true)).OnClick     := OnClickCutOffEchue ;

  // Onglet Comptable
  TCheckBox(GetControl('GCONFIDENTIEL',true)).OnClick     := OnClickConfidentiel;
  //SG6
  TCheckBox(GetControl('G_VENTILABLE',true)).OnClick      := OnClickVentilable;
  // Onglet Coordonnées
  TToolBarButton97(GetControl('G_VILLE',True)).OnDblClick      := G_VILLEDblClick;
  TToolBarButton97(GetControl('G_CODEPOSTAL',True)).OnDblClick := G_CODEPOSTALDblClick;
  TToolBarButton97(GetControl('G_DIVTERRIT',True)).OnDblClick  := G_DIVTERRITDblClick;
  {$IFDEF EAGLCLIENT}
    THValComboBox(GetControl('G_PAYS',True)).OnExit := G_PAYSChange;
    THRichEditOle(GetControl('G_BLOCNOTE', True)).OnMouseDown := OnMouseDownG_BlocNote;
  {$ELSE}
    THDBValComboBox(GetControl('G_PAYS',True)).OnExit := G_PAYSChange;
    THDBRichEditOle(GetControl('G_BLOCNOTE', True)).OnMouseDown := OnMouseDownG_BlocNote;
  {$ENDIF}

  // En attendant correction par l'agl de l'automatisme Libelle->abrégé
  {$IFDEF EAGLCLIENT}
    THEdit( GetControl('G_LIBELLE', True ) ).OnExit := G_LibelleExit ;
  {$ENDIF EAGLCLIENT}

  { FQ 22217 BVE 17.01.08 : Pbm de cast }
  {$IFNDEF EAGLCLIENT}
    THDBValComboBox( GetControl('G_NATUREGENE', True ) ).OnExit := NATUREGENEExit ;
  {$ELSE}
    THValComboBox( GetControl('G_NATUREGENE', True ) ).OnExit := NATUREGENEExit ;
  {$ENDIF EAGLCLIENT}
  { END FQ 22217 BVE 17.01.08 }
  
  G_General.OnKeyPress := OnKeyPressG_General;

  // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
  FAGLCloseQuery := TFFiche(Ecran).OnCloseQuery ;
  TFFiche(Ecran).OnCloseQuery := FormCloseQuery ;

end ;

Procedure TOM_GENERAUX.InitParPays ;
begin
  //Adaptation Pays
  SetControlVisible('G_TVASURENCAISS',VH^.PaysLocalisation<>CODEISOES) ;
  SetControlVisible('G_TPF',VH^.PaysLocalisation<>CODEISOES) ;
  SetControlVisible('TG_TPF',VH^.PaysLocalisation<>CODEISOES) ;
  //XVI 24/02/2005
end ;

procedure TOM_GENERAUX.AnalyseParametre( vStArg : String ) ;
var lStChaine : String ;
begin

  // ACTION = ...
  lStChaine := ReadTokenSt ( vStArg ) ;
  if lStChaine='ACTION=CREATIONENSERIE' then
    TFFiche(Ecran).MonoFiche := False ;

  // Choix des onglets
  lStChaine := ReadTokenSt ( vStArg ) ;
  if (lStChaine <> '') then
    FInQuellePage := StrToInt(lStChaine)
  else
    FInQuellePage := 0;

  // MODIF EN SERIE ? (appelé depuis FicheGeneMSZ)
  lStChaine := ReadTokenSt ( vStArg ) ;
  if lStChaine = 'SERIE' then
    begin
    modifEnSerie := True ;
    LesModif := vStArg;
    end
  else
    modifEnSerie := False ;


end ;

procedure TOM_GENERAUX.InitVariables ;
var
  i : Byte ;
begin

  CptesCrees          := '' ;
  LeCompte            := 'azeaze';
  Import              := False ;

  // Longueur des comptes généraux
  LgCode              := VH^.Cpta[fbGene].Lg ;

  // Sections d'attente
  for i:=1 to 5 do
    if VH^.Cpta[AxeToFb('A'+IntToStr(i))].Attente='' then
      SectAtt[i]:=False
    else
      SectAtt[i]:=True ;

  // Saisie comptable ouverte ?
  if TFFiche(Ecran).TypeAction = taModif
    then EnSaisieCpta := SaisieLancer
    else EnSaisieCpta := False ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 24/01/2005
Modifié le ... :   /  /    
Description .. : - 24/01/2005 - LG - gestion du cut-off
Mots clefs ... : 
*****************************************************************}
procedure TOM_GENERAUX.InitPresentation ;
begin

  // affectation du bon libelle pour le cutoff
  OnClickCutOffEchue(GetControl('G_CUTOFFECHUE',true)) ;

  // contexte d'aide
  Ecran.HelpContext := 7109000;

  // Paramètrage de l'onglet tables libres
  LibellesTableLibre(TTabSheet(GetControl('ZL',True)),'TG_TABLE','G_TABLE','G') ;

  // ??
  MakeZoomOLE(Ecran.Handle);

  // Spécif CEGID
  if VH^.CPIFDEFCEGID then
  begin
    G_COLLECTIF.Top        := G_LETTRABLE.Top;
    G_COLLECTIF.Left       := G_LETTRABLE.Left;
    G_COLLECTIF.Height     := G_LETTRABLE.Height ;
    G_COLLECTIF.Width      := G_LETTRABLE.Width ;
    G_COLLECTIF.Color      := G_LETTRABLE.Color;
    G_COLLECTIF.Font.Color := G_LETTRABLE.Font.Color ;
  end;

  // Présentation des montants
  FormateLesMontants ;

  // Coche IAS14 visible uniquement en cas de séralisation
  SetControlVisible('G_IAS14',  EstComptaIFRS );

  // Test permission confidentialité
  if OkConfidentiel then
   begin

    // Limitation accès si saisie ouverte
    if EnSaisieCpta or modifEnSerie then
      SetControlEnabled('BINSERT',False);

    // Gestion présentation série S3/S5/S7
    Ergo35 ;

    {$IFDEF COMPTA}
      (*
      CA - 12/05/2004 - Pour éviter de tirer toute la comptabilité, le menu
      d'accès au cumuls des comptes mensuels est rendu invisible lorsque la
      fiche des généraux est appelée depuis un module non compilé avec la directive COMPTA
      *)
    {$ELSE}
      SetControlProperty( 'BCUMUL','Tag',0 );  // Le tag est différent de -(POPUP) donc le menu ne s'affiche pas.
    {$ENDIF}

   end ;

  // MODE MULTISOC : si la table GENERAUX est partagée, on cache les infos de dernier mouvement
  if EstTablePartagee( 'GENERAUX' ) then
    begin
    SetControlVisible( 'HGBDERNMOUV', False ) ;
    end ; 

end ;

procedure TOM_GENERAUX.MajChampsVentilAxe ;
begin

  if DS.State in [dsBrowse] then Exit ;
  if VH^.AnaCroisaxe then Exit; //SG6;
  if (GetField('G_VENTILABLE1') = 'X') or (GetField('G_VENTILABLE2') = 'X') or
     (GetField('G_VENTILABLE3') = 'X') or (GetField('G_VENTILABLE4') = 'X') or
     (GetField('G_VENTILABLE5') = 'X')
    then SetField('G_VENTILABLE','X')
    else SetField('G_VENTILABLE','-') ;

end ;

function TOM_GENERAUX.IsCptBqe: Boolean;
begin
  Result := (GetField('G_NATUREGENE')='BQE') ;
  {Il est possible d'affecter un RIB sur les comptes courants : est compte courant tout compte Divers
   étant paramétré dans la table CLIENSSOC}
  Result := Result or
            ((GetField('G_NATUREGENE') = 'DIV') and EstComptaTreso and
             ExisteSQL('SELECT CLS_GENERAL FROM CLIENSSOC WHERE CLS_GENERAL = "' + GetField('G_GENERAL') + '"'));
end;

function TOM_GENERAUX.IsCptRib: Boolean;
begin
  Result := (GetField('G_NATUREGENE')='BQE') or (GetField('G_NATUREGENE')='TID')
                                             or (GetField('G_NATUREGENE')='TIC') ;
  {Il est possible d'affecter un RIB sur les comptes courants : est compte courant tout compte Divers
   étant paramétré dans la table CLIENSSOC}
  Result := Result or
            ((GetField('G_NATUREGENE') = 'DIV') and EstComptaTreso and
             ExisteSQL('SELECT CLS_GENERAL FROM CLIENSSOC WHERE CLS_GENERAL = "' + GetField('G_GENERAL') + '"'));
end;

function  TOM_GENERAUX.IsCptHt: Boolean ;
begin
  Result := (GetField('G_NATUREGENE')='IMO') or (GetField('G_NATUREGENE')='CHA')
                                             or (GetField('G_NATUREGENE')='PRO') ;
end ;

function  TOM_GENERAUX.IsCptTicTid: Boolean ;
begin
  Result := (GetField('G_NATUREGENE')='TID') or (GetField('G_NATUREGENE')='TIC') ;
end ;

function  TOM_GENERAUX.IsCptNatColl: Boolean ;
begin
  Result := (GetField('G_NATUREGENE')='COC') or (GetField('G_NATUREGENE')='COD') or
            (GetField('G_NATUREGENE')='COF') or (GetField('G_NATUREGENE')='COS') ;
end ;

procedure TOM_GENERAUX.MajEcranInfos ;
var lStVal : String ;
begin

  // INFOS DERNIER MOUVEMENT ( CACHES EN MODE MULTISOC )
  if not EstTablePartagee( 'GENERAUX' ) then
    begin
    // Date de dernier mouvement
    lStVal := FormatDateTime(Traduitdateformat('dd mmm yyyy'), GetField('G_DATEDERNMVT') ) ;
    SetControlText('DATEDERNMVT', lStVal ) ;
    // Montants du dernier mouvement //StrfMask(V_PGI.OkdecV,'',True)
    lStVal := StrFMontant( GetChampsMS('G_DEBITDERNMVT'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('G_DEBITDERNMVT1',  lStVal ) ;
    lStVal := StrFMontant( GetChampsMS('G_CREDITDERNMVT'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('G_CREDITDERNMVT1', lStVal ) ;
    end ;

  // TOTAUX DU COMPTE
  // Maj des champs tampons pour affichage des montants d'infos
  lStVal := StrFMontant( GetChampsMS('G_TOTDEBP'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('G_TOTDEBP1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('G_TOTCREP'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('G_TOTCREP1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('G_TOTDEBE'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('G_TOTDEBE1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('G_TOTCREE'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('G_TOTCREE1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('G_TOTDEBS'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('G_TOTDEBS1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('G_TOTCRES'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('G_TOTCRES1', lStVal ) ;

  // maj de l'affichage des soldes
  AfficheLeSolde( THNumEdit(GetControl('G_SOLCREP',True)),
                  GetChampsMS('G_TOTDEBP'),
                  GetChampsMS('G_TOTCREP')  );
  AfficheLeSolde( THNumEdit(GetControl('G_SOLCREE',True)),
                  GetChampsMS('G_TOTDEBE'),
                  GetChampsMS('G_TOTCREE')  );
  AfficheLeSolde( THNumEdit(GetControl('G_SOLCRES',True)),
                  GetChampsMS('G_TOTDEBS'),
                  GetChampsMS('G_TOTCRES')  );

  // PAGE INFORMATIONS : DATES
  lStVal := FormatDateTime(Traduitdateformat('dd mmm yyyy'), GetField('G_DATEMODIF') ) ;
  SetControlText('DATEMODIF', lStVal ) ;
  lStVal := FormatDateTime(Traduitdateformat('dd mmm yyyy'), GetField('G_DATECREATION') ) ;
  SetControlText('DATECREATION', lStVal ) ;
  lStVal := FormatDateTime(Traduitdateformat('dd mmm yyyy'), GetField('G_DATEOUVERTURE') ) ;
  SetControlText('DATEOUVERTURE', lStVal ) ;
  lStVal := FormatDateTime(Traduitdateformat('dd mmm yyyy'), GetField('G_DATEFERMETURE') ) ;
  SetControlText('DATEFERMETURE', lStVal ) ;

end ;

function TOM_GENERAUX.IsCptVentilable  : Boolean ;
begin
  Result := GetField('G_VENTILABLE') = 'X' ;
end ;

function TOM_GENERAUX.IsCptLettrable   : Boolean ;
begin
  Result := GetField('G_LETTRABLE') = 'X' ;
end ;

{$IFDEF COMPTA}
function TOM_GENERAUX.IsCptCollectif   : Boolean ;
begin
  Result := GetField('G_COLLECTIF') = 'X' ;
end ;
{$ENDIF}

procedure  TOM_GENERAUX.majEcranNatureGene ;
var i : Integer ;
    lBoCutOff : boolean ;
begin

  // Gestion des pages
  TTabSheet(GetControl('PADRESSE',True)).TabVisible := IsCptTicTid ; // Affichage Page "Coordonnées" ?
  TTabSheet(GetControl('PINFO',True)).TabVisible    := IsCptTicTid ;    // Affichage Page "Informations" ?

  // Accès TVA
  SetControlEnabled('GBTVA', IsCptHT );
  EnabledTva ;

  // Limitations fonctionnalités annexes
  SetControlEnabled('BVENTIL',    IsCptVentilable);  // Accès ventilation par défaut si compte ventilable
{$IFDEF COMPTA}
  SetControlEnabled('BBALGENANA', IsCptVentilable ); // Accès balance général par analytique si compte ventilable
  SetControlEnabled('BBALGENAUX', IsCptCollectif );  // Accès balance général par auxi si compte collectif
{$ENDIF}
  SetControlEnabled('BJUSTIF',    IsCptLettrable );  // Accès justif de solde si compte lettrable
  SetControlEnabled('BRIB',       IsCptRib );        // Accès Rib
  {$IFNDEF IMP}
    {$IFDEF SANSIMMO}
      SetControlEnabled('BIMMO', False);  // Accès Immo
    {$ELSE}
      SetControlEnabled('BIMMO', ( (VH^.OkModImmo) and IsCptImmo )) ; // Accès Immo
    {$ENDIF}
  {$ENDIF}

  // EN MODE CREATION si la nature n'est pas encore renseignée...
  if (GetField('G_NATUREGENE') = '') And (DS.State=dsInsert) then
    begin
    // Caractéristiques accéssible sauf IAS
    SetControlEnabled('G_LETTRABLE',  True);
    SetControlEnabled('G_POINTABLE',  True);
    SetControlEnabled('G_IAS14',      False);
    // Ventilation des axes accessibles
    SetControlEnabled('G_VENTILABLE', True);
    for i:=1 to 5 do
    begin
      if ({EstComptaPackAvance and }not(GetParamSocSecur('SO_CROISAXE', False))) then SetControlEnabled('G_VENTILABLE'+IntToStr(i), True); //SG6
    end;
    end
  else
  begin // SINON Gestion de l'accès au carac. / ventil. classique
    // Accès caractéristiques du compte
    majAccesLettrage ;
    majAccesPointage ;
    majAccesVentil ;
    majAccesIAS ;

    lBoCutOff := ( GetField('G_NATUREGENE') = 'CHA' ) or ( GetField('G_NATUREGENE') = 'PRO' ) ;
    SetControlProperty( 'G_CUTOFF'       , 'Enabled' ,  lBoCutOff and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
    SetControlProperty( 'G_CUTOFFPERIODE', 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
    SetControlProperty( 'G_CUTOFFECHUE'  , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
    SetControlProperty( 'LCUTOFFPER'     , 'Enabled' ,  (GetField('G_CUTOFF') = 'X') and ( TFFiche(Ecran).TypeAction <> taConsult )) ;
  end ;
end ;


procedure  TOM_GENERAUX.majChampsNatureGene ;
var lStCode : String ;
begin

  if (DS.State in [dsBrowse]) then Exit ;

  // VALEURS TVA
  if Not IsCptHT then
    begin
    if GetField('G_TVA')<>'' then SetField('G_TVA','') ;
    if GetField('G_TPF')<>'' then SetField('G_TPF','') ;
    if GetField('G_TVASURENCAISS')<>'-' then SetField('G_TVASURENCAISS','-') ;
    end;

  // VALEURS LETTRABLE
  if IsCptTicTid then
    if GetField('G_LETTRABLE')<>'X' then SetField('G_LETTRABLE','X') ;

  majChampsLettrable ;

  // Maj G_COLLECTIF
  {JP 10/01/08 : FQ 22136 : Pour SIC, si on passe d'un collectif à un divers, G_COLLECTIF n'était pas remis à -}
  if IsCptNatColl then
    SetField('G_COLLECTIF','X')         // valable pour tous
  else if not VH^.CPIFDEFCEGID then
    SetField('G_COLLECTIF','-')
  else begin
    {JP 10/01/08 : FQ 22136 : SIC gérant les comptes de charge et de produit collectifs ...}
    if not ((GetField('G_NATUREGENE') = 'PRO') or (GetField('G_NATUREGENE') = 'CHA')) then
      SetField('G_COLLECTIF','-')
  end;

  
  // MAj RIBs ?
  If (Not IsCptBqe) And (Not IsCptRib) Then
    begin
    lStCode := GetField('G_GENERAL') ;
    if BourreLess( lStCode , fbgene ) = '' then Exit ;
    if TFFiche(Ecran).TypeAction in [taCreat..taCreatOne] then
      ExecuteSQL('DELETE FROM RIB WHERE R_AUXILIAIRE="' + lStCode + '"');
    end ;

end ;

procedure TOM_GENERAUX.majAccesLettrage ;
Var
  Q : TQuery ;
  CodeL : string ;
begin
  Q := OpenSQL('Select G_DERNLETTRAGE from GENERAUX Where G_GENERAL="' + GetField('G_GENERAL') +'"',true) ;
  if Not Q.EOF then
    CodeL := Q.FindField('G_DERNLETTRAGE').asString ;
  ferme(Q);

  // ===== Limitation totale =====
  // GCO - 16/10/2007 - FQ 21570
  // GCO - 29/01/2008 - FQ 22169
  if EstCpteSociete(GetField('G_GENERAL')) then
  begin
    SetControlEnabled( 'G_LETTRABLE', false);
    SetControlEnabled( 'G_VENTILABLE', false);
  end
  else
  // FIN GCO - 16/10/2007 - FQ 21570
  if (( EnSaisieCpta ) or                                     // Saisie ouverte
      ( not IsCptTicTid ) or                                  // pas un Tic / Tid
      ( AvecMvt and IsCptLettrable ) or                       // compte lettrable mouvementé
      ( GetField('G_GENERAL') = VH^.Cpta[fbGene].Attente ) or // compte d'attente
      ( TFFiche(Ecran).TypeAction = taConsult )) and          // consultation
      (GetField('G_NATUREGENE')<>'DIV') then
  begin
    SetControlEnabled( 'G_LETTRABLE', false) ;
    if ( DS.State in [dsEdit,dsInsert] ) and ( GetField('G_LETTRABLE') <> '-' ) then
      SetField('G_LETTRABLE','-');
  end
  else
  begin
    // GCO - 05/04/2006 - FQ 17537
    if TFFiche(Ecran).TypeAction = taConsult then
      SetControlEnabled('G_LETTRABLE', False)
    else
    begin // ======= Compte Lettrable Divers =======
      if (GetField('G_NATUREGENE') = 'DIV') then // FQ 17536 - CA pas de modification de lettrage si mouvement sur le compte
        SetControlEnabled('G_LETTRABLE', not(AvecMvt and IsCptLettrable ) (*or (CodeL='')*) ) // si lettrable et pas mouvementé
      else // ======= Accessible pour les TicTid =======
        SetControlEnabled('G_LETTRABLE', IsCptTicTid) ;
    end;  
  end;

  // ==== Spécif CEGID ====
  if VH^.CPIFDEFCEGID then
  begin
    SetControlVisible('G_LETTRABLE', True);
    SetControlVisible('G_COLLECTIF',False);
    if (GetField('G_NATUREGENE')='PRO') or (GetField('G_NATUREGENE')='CHA') then
      if not IsCptLettrable then
        SetControlVisible('G_COLLECTIF',True);
  end;
end ;

procedure TOM_GENERAUX.majAccesPointage ;
var lBoNatOk : Boolean ;
begin

  // Nature compatible ?
  lBoNatOk := ( Getfield('G_NATUREGENE') = 'DIV' ) or
              ( Getfield('G_NATUREGENE') = 'EXT' ) or
              ( Getfield('G_NATUREGENE') = 'BQE' ) or
              ( Getfield('G_NATUREGENE') = 'CAI' ) ;

  // ===== CAS de Limitation totale : =====
  if ( EnSaisieCpta ) or                                         // Saisie ouverte
     ( AvecMvt and IsCptPointable ) or                           // Compte mouvementé pointable
     ( not lBoNatOk ) or                                         // Nature Géné comptabile
     (GetField('G_GENERAL') = VH^.Cpta[fbGene].Attente) or       // compte d'attente
     ( TFFiche(Ecran).TypeAction = taConsult ) then             // consultation
    begin
    SetControlEnabled('G_POINTABLE', False) ;
    if ( DS.State in [dsEdit,dsInsert] ) and ( GetField('G_POINTABLE') <> '-' ) then
      SetField('G_POINTABLE','-');
    end
  else
  // ===== Sinon ok ====
    SetControlEnabled('G_POINTABLE', True ) ;

end ;

procedure TOM_GENERAUX.majAccesVentil ;
var i          : Integer ;
    lBoEtat    : Boolean ;
begin

  // ===== Limitation totale =====
  if ( EnSaisieCpta ) or                                       // Saisie ouverte
 //    ( AvecMvt and IsCptNatColl ) or                           // Compte collectif mouvementé //SG6 24/11/04 FQ 14597
     ( GetField('G_GENERAL') = VH^.Cpta[fbGene].Attente ) or   // compte d'attente
     ( TFFiche(Ecran).TypeAction = taConsult ) or             // consultation
     ( IsCptAvecEcritureMultiEch ) or //Multi Echéances   //SG6 05/01/05 FQ 14597 Pas de ventilation sur multi echéance
     // Compta sans analytique :
     ( ((ctxPCL in V_PGI.PGIContexte) or (EstComptaSansAna)) and (GetParamSocSecur('SO_CPPCLSANSANA', True)=TRUE) ) then
    begin
    SetcontrolEnabled('HGBOPTIONAXES',False);
    SetControlEnabled('G_VENTILABLE',False);
    SetcontrolEnabled('G_VENTILABLE1',False);
    SetcontrolEnabled('G_VENTILABLE2',False);
    SetcontrolEnabled('G_VENTILABLE3',False);
    SetcontrolEnabled('G_VENTILABLE4',False);
    SetcontrolEnabled('G_VENTILABLE5',False);
    end
  else
  // ======= Limitation partielle si mouvements existant =======
  if ( AvecMvt and IsCptVentilable ) then
      begin
      SetControlEnabled('G_VENTILABLE',False);
      // grise les checkbox des axes qui ne sont pas cochées
      for i := 1 to 5 do
        begin
        if ({EstComptaPackAvance and }not(GetParamSocSecur('SO_CROISAXE', False))) then
        begin
        lBoEtat := GetField('G_VENTILABLE'+IntToStr(i)) <> 'X' ;
        SetControlEnabled( 'G_VENTILABLE'+IntToStr(i) , lBoEtat );
        end ;
      end;
      end
  else
    begin
    SetControlEnabled('G_VENTILABLE', ( TFFiche(Ecran).TypeAction <> taConsult ) );
    for i:=1 to 5 do
      if ({EstComptaPackAvance and }not(GetParamSocSecur('SO_CROISAXE', False))) then SetControlEnabled('G_VENTILABLE'+IntToStr(i), ( TFFiche(Ecran).TypeAction <> taConsult ) ); //SG6
    end ;

end ;

procedure TOM_GENERAUX.majAccesNatureGene ;
begin

  // Blocage totale
  if EnSaisieCpta or                                       // saisie en cours
     (GetField('G_GENERAL') = VH^.Cpta[fbGene].Attente) or // compte d'attente
     ( TFFiche(Ecran).TypeAction = taConsult ) then        // consultation
    SetControlEnabled('G_NATUREGENE',False)
  // Blocage si mouvementé en mode PCL
  // FQ 12063 - En entreprise aussi, on interdit la modification
  else (*if (ctxPcl in V_Pgi.PGIContexte) then*)
    SetControlEnabled('G_NATUREGENE', not AvecMvt) ;

  // maj visuel
  if GetControlEnabled('G_NATUREGENE')
    then SetControlProperty('G_NATUREGENE', 'COLOR', ClWindow )
    else SetControlProperty('G_NATUREGENE', 'COLOR', ClBtnFace ) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/07/2006
Modifié le ... :   /  /    
Description .. : LG - 28/07/2006 - on rend ventilable automatiquement les 
Suite ........ : comptes ds tranches de tva si elle est activée
Mots clefs ... : 
*****************************************************************}
procedure TOM_GENERAUX.MajChampsGeneral ;
var QQ        : TQuery ;
    lStCode   : String ;
    lStRacine : String ;
    lStNature : String ;
    lStAxe    : string ;
begin

  if DS.State <> dsInsert then Exit ;

  lStCode := GetField('G_GENERAL') ;
  if ( lStCode <> '' ) then
    begin
    if GetParamSocSecur('SO_CPPCLSAISIETVA',false) and  ( ( Copy(lStCode,1,2) = '41' ) or ( Copy(lStCode,1,2) = '70' ) ) then
     begin
      SetField('G_VENTILABLE','X') ;
      lStAxe := GetParamSocSecur('SO_CPPCLAXETVA', '') ;
      SetField('G_VENTILABLE'+Copy(lStAxe,2,1) , 'X' ) ;
     end ;
    // MAJ de la nature suivant racine
    if GetField('G_NATUREGENE') = '' then
      begin
      lStRacine := Copy( lStCode , 1 , 3 ) ;
      lStNature := '' ;
      QQ := OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL LIKE "'+lStRacine+'%"',True) ;
      if Not QQ.EOF then
        lStNature := QQ.Fields[0].AsString ;
      Ferme(QQ) ;
      if lStNature <> '' then
        begin
        // Maj du sens
        if (lStNature = 'CHA') then
          SetField('G_SENS','D')
        else if (lStNature = 'PRO') then
          SetField('G_SENS','C');
        // MAJ de la nature
        SetField('G_NATUREGENE', lStNature ) ;
      end;
    end;
  end;
  ChercheEtImportCptePlanRef ;

end ;

procedure TOM_GENERAUX.MajChampsLettrable ;
begin
  if DS.State = dsBrowse then Exit ;

  if ( not IsCptTicTid ) or ( IsCptTicTid and not IsCptLettrable ) then
    begin

    if GetField('G_TVAENCAISSEMENT') = '' then
       SetField('G_TVAENCAISSEMENT', '') ;
  //    SetControlProperty('G_TVAENCAISSEMENT','ITEMINDEX', 0 );

    if GetField('G_MODEREGLE') = '' then
       SetField('G_MODEREGLE', '') ;
  //    SetControlProperty('G_MODEREGLE','ITEMINDEX', 0 );

    if GetField('G_REGIMETVA') = '' then
       SetField('G_REGIMETVA', '') ;

    end ;

end ;

function TOM_GENERAUX.IsCptIAS : Boolean ;
begin
  Result := ( GetField('G_NATUREGENE') = 'COF' ) or ( GetField('G_NATUREGENE') = 'COC' ) or
            ( GetField('G_NATUREGENE') = 'COS' ) or ( GetField('G_NATUREGENE') = 'COD' ) or
            ( GetField('G_NATUREGENE') = 'TIC' ) or ( GetField('G_NATUREGENE') = 'TID' ) ;
//          ( GetField('G_NATUREGENE') = 'DIV' ) or 
end ;

procedure TOM_GENERAUX.majAccesIAS ;
begin
  if IsCptIAS then
    SetControlEnabled('G_IAS14', ( TFFiche(Ecran).TypeAction <> taConsult ) )
  else
    begin
    SetControlEnabled('G_IAS14', False ) ;
    if (DS.State in [dsInsert,dsEdit]) and ( GetField('G_IAS14')<>'-' ) then
      SetField('G_IAS14', '-' ) ;
    end ;
end ;


Function TOM_GENERAUX.IsCptPointable   : Boolean ;
begin
  Result := GetField('G_POINTABLE') = 'X' ;
end ;

{$IFNDEF IMP}
function TOM_GENERAUX.IsCptImmo: Boolean;
begin
  Result := GetField('G_NATUREGENE') = 'IMO' ;
end;
{$ENDIF}

function TOM_GENERAUX.VerifiSectionAtt ( vStChamp : String ) : Boolean ;
var lInAxe   : Integer ;
begin

  Result := False ;
  if (DS.State in [dsBrowse]) then Exit ;
  if GetField( vStChamp ) <> 'X' then Exit ;

  if (vStChamp <> 'G_VENTILABLE') and (vStChamp <> '') then
    begin
    lInAxe := StrToInt( vStChamp[ Length(vStChamp) ] );
    if Not SectAtt[lInAxe] then
      begin
      MsgBox.Execute(37, MsgBox.Mess[38] + ' ' + IntToStr(lInAxe) + ' ' , '' ) ; // L'axe ? ne possède aucun code section, et n'a pas de section d'attente.
      SetField( vStChamp, '-' );
      {$IFNDEF EAGLCLIENT}
        SetControlText(vStChamp, '-' );
      {$ENDIF}
      Result := True ;
      Exit ;
      end;
    end;

end;
                                                             
procedure TOM_GENERAUX.majEcranVentilAxe;
var lBoEnabled : Boolean;
begin
  lBoEnabled := ( GetField('G_VENTILABLE1') = 'X' ) or
                ( GetField('G_VENTILABLE2') = 'X' ) or
                ( GetField('G_VENTILABLE3') = 'X' ) or
                ( GetField('G_VENTILABLE4') = 'X' ) or
                ( GetField('G_VENTILABLE5') = 'X' ) ;
  SetControlEnabled('BVENTIL',     lBoEnabled);
{$IFDEF COMPTA}
  SetControlEnabled('BBALGENANA',  lBoEnabled );
{$ENDIF}
end;

function TOM_GENERAUX.CompleteGeneral : Boolean ;
var lStCode : String ;
begin
  result := False ;

  if DS.State <> dsInsert then exit;

  lStCode := GetField('G_GENERAL') ;
  if lStCode <> '' then
  begin
    lStCode := BourreEtLess( lStCode, fbGene ) ;
    Result := lStCode <> GetField('G_GENERAL') ;
    if Result then
      SetField('G_GENERAL', lStCode) ;
  end ;
end;

procedure TOM_GENERAUX.AutoriseModifications;
var ModPaie  : String ;
    AuxDiv   : String ;
    ModRegle : String ;
    Tableau  : Array[0..MaxAxe] of string;
    i        : integer;
//    lQMR     : TQuery ;
begin

  // Gestion de la modification de la caractéristique collectif
  if (CollectifChanger) then
    begin
    if EstCpteCollRef( GetField('G_GENERAL') ) then
      begin
      MsgBox.Execute(20,'','');  // La modification de la nature collective est impossible : ce compte est déjà référencé par des tiers.
      SetField('G_NATUREGENE',MNat);
      end
    else
      begin
      if (AvecMvt) then
        begin
        MsgBox.Execute(22,'','');  // La modification de la nature collective est impossible : ce compte est mouvementé.
        SetField('G_NATUREGENE',MNat);
        end;
      end;
    end;

  // Si on a changer la nature du compte a partire d'un collectif vers un divers ... test de referencement !
  if (GetField('G_NATUREGENE') = 'DIV') then
    begin
    if ((MNat='COC') or (MNat='COD') or (MNat='COF') or (MNat='COS')) then
      begin
      if EstCpteCollRef( GetField('G_GENERAL') ) then
        begin
        MsgBox.Execute(20,'','');  // La modification de la nature collective est impossible : ce compte est déjà référencé par des tiers.
        SetField('G_NATUREGENE',MNat);
        end;
      end;
    end;

  // COMPTE DE CONTREPARTIE -> Si référencé dans un journal, nature non modifiable
  if (BqeCaiChanger) then
    begin
    if Presence( 'JOURNAL', 'J_CONTREPARTIE', GetField('G_GENERAL') ) then
      begin
      MsgBox.Execute(21,'','');  // La modification de la nature est impossible : ce compte est compte de contrepartie d''un journal.
      SetField('G_NATUREGENE',MNat);
      end;
    end;

  {FQ19946 20.12.2007 YMO Message d'avertissement sur modif paramètres TVA sur un compte}
  if (TVAChanger) then MsgBox.Execute(60,'','');

  // Modification de la caractéristique de lettrage du compte
  if (LettrageChanger) then
    begin
    if (Not _BlocCarFiche) then
      if MNat <> 'DIV' then
        if MsgBox.Execute(18,'','')=mrYes then // Le lettrage a été modifié. Désirez-vous continuer ?
        begin
          MsgBox.Execute(23,'',''); // La modification de la caractéristique lettrable va entraîner des modifications d''écritures. Vous devez sélectionner un type de mode de paiement.
          ModPaie := Choisir(MsgBox.Mess[29],'MODEPAIE','MP_LIBELLE','MP_MODEPAIE','','');  // Choix d''un mode de paiement
          if (ModPaie <> '') then
            MajLettrageEcriture(ModPaie,GetField('G_GENERAL'),fbGene)
          else
          begin
            MsgBox.Execute(24,'',''); // Le traitement n''est pas effectué. Le mode de paiement n''est pas renseigné.
            SetField('G_LETTRABLE',ML);
          end;
        end
        else
          SetField('G_LETTRABLE',ML)
      else //Nature 'DIV'
      begin
        // SBO 23/08/2005 : Totalement faux, on veux affecter un mode de reglement à une zone qui doit contenir un mode de paiement !!!
        //        ModPaie := GetParamSoc('SO_GCMODEREGLEDEFAUT'); //Choix d''un mode de paiement
        // Il faut aller chercher le 1er mode de paiement paramétré dans le mode de règlement
        (*
        ModRegle := GetParamSocSecur('SO_GCMODEREGLEDEFAUT', '');
        ModPaie  := '' ;
        if ModRegle <> '' then
          begin
          lQMR := OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+ModRegle+'" ',TRUE) ;
          if Not lQMR.Eof Then
            ModPaie := lQMR.FindField('MR_MP1').AsString ;
          Ferme(lQMR) ;
          if (ModPaie <> '') then
            MajLettrageEcriture( ModPaie, GetField('G_GENERAL'), fbGene)
          end ;
          *)

          // ajout me 09/02/2005 pour avoir une fontion unique
          ModPaie := ModePaiementParDefaut (ModRegle);
          if (ModPaie <> '') then
            MajLettrageEcriture( ModPaie, GetField('G_GENERAL'), fbGene);


        if (ModRegle = '' ) or (ModPaie = '') then
          begin
            MsgBox.Execute(24,'',''); // Le traitement n''est pas effectué. Le mode de paiement n''est pas renseigné.
            SetField('G_LETTRABLE',ML);
          end;

      end;
    end;

  // Modification des axes à ventiler
  if (VentilChanger) then
    begin
    // La ventilation a été modifiée. Désirez-vous continuer ?
    if ((Not _BlocCarFiche) and (MsgBox.Execute(19,'','')=mrYes))
      then
      begin
        Tableau[0] := '-';
        for i := 1 to MaxAxe do
        begin
           if IsCptVentilAxe(i) then
           begin
              Tableau[i] := 'X';
              Tableau[0] := 'X';
           end
           else
              Tableau[i] := '-';
        end;

        {JP 23/10/07 : FQ 21714 : Gestion de la mise des ventilations en multi dossiers
        MajVentilation(GetField('G_GENERAL'),Tableau,False);}
        MajVentilationMulti(GetField('G_GENERAL'), Tableau, False);

        // FQ 18925 - CA - 13/10/2006 : On recharge l'état des ventilations pour ne pas faire 2 fois le travail
        // et ainsi éviter le plantage sous ORCL
        MA0     := GetField('G_VENTILABLE') ;
        MA1     := GetField('G_VENTILABLE1') ;
        MA2     := GetField('G_VENTILABLE2') ;
        MA3     := GetField('G_VENTILABLE3') ;
        MA4     := GetField('G_VENTILABLE4') ;
        MA5     := GetField('G_VENTILABLE5') ;
      end
      else RepositionneVentilation ;
    end;

  // Spécifique CEGID
  if VH^.CPIFDEFCEGID then
    begin
    if (CollectifChangerChaPro) then
      begin
      if ((Not _BlocCarFiche) and (MsgBox.Execute(50,'','')=mrYes)) then // La nature "collectif" a été modifiée. Désirez-vous continuer ?
        begin
        AuxDiv:='' ;
        if (Not G_COLLECTIF.Checked) then
          begin
          MajCollChaProEcriture(GetField('G_GENERAL'),AuxDiv) ;
          MColl := GetField('G_COLLECTIF') ;
          end
        else
          begin
          MsgBox.Execute(51,'','') ; // La modification de la caractéristique Collectif va entraîner des modifications d''écritures. Vous devez sélectionner un auxiliaire.
          AuxDiv := Choisir(MsgBox.Mess[52],'TIERS','T_AUXILIAIRE','T_AUXILIAIRE','','T_AUXILIAIRE') ; // Choix d''un mode auxiliaire
          if (AuxDiv <> '') then
            begin
            MajCollChaProEcriture(GetField('G_GENERAL'),AuxDiv) ;
            MColl := GetField('G_COLLECTIF') ;
            end
          else
            begin
            MsgBox.Execute(53,'',''); // Le traitement n''est pas effectué. Le compte auxiliaire n''est pas renseigné.
            SetField('G_COLLECTIF',MColl);
            end;
          end;
        end
      else
        SetField('G_COLLECTIF',MColl); //    SetControlText('G_Collectif',MColl);
      end;
    end;

    {FQ18530 YMO 03/04/2007 Maj de l'historique des écritures}
    If ChangeConfidentiel then
    begin
        ExecuteSQL('UPDATE ECRITURE SET E_CONFIDENTIEL="'+GetField('G_CONFIDENTIEL')+'" '
                  +' WHERE CAST(E_NUMEROPIECE AS CHAR(6))||E_JOURNAL||CAST(E_PERIODE AS CHAR(6))||E_QUALIFPIECE IN'
                  +' (SELECT CAST(E_NUMEROPIECE AS CHAR(6))||E_JOURNAL||CAST(E_PERIODE AS CHAR(6))||E_QUALIFPIECE'
                  +' FROM ECRITURE WHERE E_GENERAL="'+GetField('G_GENERAL')+'")');
        ChangeConfidentiel :=False ;                  
    end;
end;

{$IFDEF EAGLCLIENT}
procedure TOM_GENERAUX.G_LIBELLEExit(Sender: TObject);
var lStAbrege  : String ;
    lStLibelle : String ;
begin
  if (DS.State<>dsInsert) then Exit ;

  lStLibelle := Trim( GetControlText('G_LIBELLE'));
  lStAbrege  := Trim( GetField('G_ABREGE') );

  if ( lStAbrege = '' ) And (lStLibelle <> '') then
    SetField( 'G_ABREGE' , Copy( lStLibelle, 1, 17 )  ) ;

end;
{$ENDIF EAGLCLIENT}

procedure TOM_GENERAUX.NATUREGENEExit(Sender: TObject);
begin
        // ajout me fiche 11635
  if (GetField('G_REGIMETVA') = '') and ((GetField('G_NATUREGENE')='TID')  or (GetField('G_NATUREGENE')='TIC')) then
      SetField('G_REGIMETVA', 'FRA');
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/07/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_GENERAUX.InitEcrInfoPointage( vStCompte : string ) : Boolean;
begin
  Result := True;
  try
    ExecuteSQL('UPDATE ECRITURE SET ' +
               'E_LETTRAGE = "", ' +
               'E_LETTRAGEDEV = "-", ' +
               'E_ECHE = "X", ' +
               'E_ETATLETTRAGE = "RI", ' +
               'E_COUVERTURE = 0, ' +
               'E_COUVERTUREDEV = 0, ' +
               'E_DATEPAQUETMAX = E_DATECOMPTABLE, ' +
               'E_DATEPAQUETMIN = E_DATECOMPTABLE, ' +
               'E_DATEECHEANCE = E_DATECOMPTABLE, ' +
               'E_NUMECHE = 1, ' +
               'E_MODEPAIE = "DIV" ' +
               'WHERE E_GENERAL = "' + vStCompte + '"');
  except
    Result := False;
  end;
end;
////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 03/09/2004
Modifié le ... : 03/09/2004
Description .. : Ajout des modifications afin d'accéder à SGEP
Mots clefs ... :
*****************************************************************}

{$IFNDEF EAGLCLIENT}
procedure TOM_GENERAUX.AppelIsoflex(Sender: TObject);
begin
  AglIsoflexViewDoc(NomHalley, Ecran.Name, 'GENERAUX', 'G_GENERAL', '', GetField('G_GENERAL'), '');
end;

function TOM_GENERAUX.GereIsoflex : boolean ;
begin
  if AglIsoflexPresent then
  begin
    SetControlVisible('BISOFLEX', TRUE);
    result := TRUE;
  end
  else
    result := FALSE;
end;
//FIN?
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 01/08/2007
Modifié le ... :   /  /    
Description .. : Permet de mettre à jour la ventilation type si on clique sur un 
Suite ........ : nouvel axe
Suite ........ : FQ 21209
Mots clefs ... : 
*****************************************************************}
procedure TOM_GENERAUX.OnClickGVentilableX(Sender : TObject);
var
  SectTemp : Array[1..5] of Boolean ; // stockage des section d'attente
  i        : integer;
  lAxeVent : array [1..5] of Boolean; // FQ 21259
begin
  if gbLoad then Exit;
  if ( TCheckBox(Sender).Checked ) then
  begin
     for i := 1 to 5 do
     begin
        if ( TCheckBox(Sender).Name = 'G_VENTILABLE' + IntToStr(i) ) then
           SectTemp[i] := true
        else
           SectTemp[i] := false;
     end;
     InitialiseVentilDefaut(GetField('G_GENERAL'),SectTemp, lAxeVent); // FQ 21259
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 09/12/2004
Modifié le ... :   /  /
Description .. : Gestion des ventilations en mode croiseaxe
Mots clefs ... :
*****************************************************************}
procedure TOM_GENERAUX.OnClickVentilable(Sender : TObject);
var
  i : integer;
  sTmp : string;
begin
  if not(GetParamSocSecur('SO_CROISAXE', False)) then exit;
  if TCheckBox(Sender).Checked then
  begin
    for i:=1 to 5 do
    begin
      if GetParamSocSecur('SO_VENTILA'+IntToStr(i), False) then
        sTmp := 'X'
      else
        sTmp := '-';
      SetField('G_VENTILABLE'+IntToStr(i),sTmp);
    end;
  end
  else
  begin
    for i:=1 to 5 do
    begin
      SetField('G_VENTILABLE'+IntToStr(i),'-');
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 05/01/2005
Modifié le ... : 08/03/2005
Description .. : Fonction qui permet de voir si le compte général a servi pour 
Suite ........ :  saisir des écritures MULTI échéances
Suite ........ : - LG - 08/03/2005 - optimisation du code
Mots clefs ... : MULTI ECHEANCE
*****************************************************************}
function TOM_GENERAUX.IsCptAvecEcritureMultiEch: Boolean ;
begin
  Result := ExisteSQL('SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="' + GetField('G_GENERAL') + '" ' +
                      ' AND E_NUMECHE>1') ;
end;

////////////////////////////////////////////////////////////////////////////////


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 25/05/2005
Modifié le ... :   /  /
Description .. : Retourne la valeur de champ de cumul identifié par
Suite ........ : vStChamp. Cette valeur est récupérée de la table CUMULS
Suite ........ : en mode Multisoc, ou directement de la table GENERAUX
Suite ........ : sinon.
Mots clefs ... : CUMULS;MULTISOC
*****************************************************************}
function TOM_GENERAUX.GetChampsMS(vStChamp: String): double;
begin
  result := 0 ;
  if DS.State = dsInsert then Exit ;

  if EstTablePartagee( 'GENERAUX' ) then
    begin
    if not Assigned( FTobCumuls ) then Exit ;
    result := FTobCumuls.GetDouble(  vStChamp ) ;
    end
  else result := Valeur( GetField( vStChamp ) ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 25/05/2005
Modifié le ... :   /  /
Description .. : UNIQUEMENT EN MULTISOC
Suite ........ : Renseigne FTobCumuls avec les valeurs locales de cumuls
Suite ........ : en mode multisoc.
Suite ........ : Infos chargées depuis la table CUMULS
Mots clefs ... : MULTISOC;CUMULS
*****************************************************************}
procedure TOM_GENERAUX.ChargeCumulsMultiSoc;
var lstReq  : String ;
    lQCumul : TQuery ;
begin

  if DS.State = dsInsert then Exit ;
  if not EstTablePartagee( 'GENERAUX' ) then Exit ;

  if Assigned( FTobCumuls ) then FreeAndNil( FTobCumuls ) ;

  lStReq := 'SELECT CU_DEBIT1 as G_TOTALDEBIT ,  CU_CREDIT1 as G_TOTALCREDIT ,'
                 + 'CU_DEBIT2 as G_TOTDEBP ,CU_CREDIT2 as G_TOTCREP ,'
                 + 'CU_DEBIT3 as G_TOTDEBE ,CU_CREDIT3 as G_TOTCREE ,'
                 + 'CU_DEBIT4 as G_TOTDEBS ,CU_CREDIT4 as G_TOTCRES ,'
                 + 'CU_DEBITAN as G_TOTDEBANO ,CU_CREDITAN as G_TOTCREANO ,'
                 + 'CU_DEBIT5 as G_TOTDEBANON1,CU_CREDIT5 as G_TOTCREANON1,'
                 + 'CU_DEBIT6 as G_DEBNONPOINTE ,CU_CREDIT6 as G_CREDNONPOINTE ,'
                 + 'CU_DEBIT7 G_TOTDEBPTP ,CU_CREDIT7 G_TOTCREPTP ,'
                 + 'CU_DEBIT8 G_TOTDEBPTD ,CU_CREDIT8 G_TOTCREPTD '
             + 'FROM CUMULS WHERE CU_TYPE="' + fbToCumulType(fbGene) + '" '
                           + 'AND CU_COMPTE1="' + GetField('G_GENERAL') + '"' ;

  lQCumul := OpenSQL( lStReq, True ) ;
  if not lQCumul.Eof then
    begin
    FTobCumuls := TOB.Create('_CUMULS', nil, -1 ) ;
    FTobCumuls.SelectDB('', lQCumul ) ;
    end ;

  Ferme(lQCumul);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/06/2005
Modifié le ... :   /  /
Description .. : Correction FQ 15897
Mots clefs ... :
*****************************************************************}
procedure TOM_GENERAUX.OnMouseDownG_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DS.State = dsBrowse then
    DS.Edit;
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. : Bouton copier ; duplication des champs du compte choisi
Mots clefs ... :
*****************************************************************}
procedure TOM_GENERAUX.BCopierClick(Sender: TObject);
// Copie la rubrique courante et place les données en mode nouveau
Var St : String ;
    Q : TQuery ;
    i : integer ;
    Cache : THCpteEdit; {JP 02/06/06 : FQ 17971}
    tz : tZoomTable;
begin

	// Test code saisi ?
   If GetControlText('G_GENERAL')='' Then
     BEGIN
     MsgBox.Execute(56,'','') ;
  //   G_GENERAL.SetFocus ;
     Exit ;
     END ;

    tz := tzGeneral;
    Cache := THCpteEdit.Create(Ecran);
    Cache.Visible := false;
    Cache.Parent := Ecran;
    Cache.ZoomTable := tz;
    Cache.okLocate := true;
    Cache.Text := GetControlText('G_GENERAL') ;

  {JP 02/06/06 : FQ 17971 : changement du LookUp}
  if (GChercheCompte(Cache, FicheGene)) then begin
    st := Cache.Text;
	// Choix de la rubrique a dupliquer
    //  St := Choisir(MsgBox.Mess[58],'GENERAUX','G_GENERAL || " / " || G_LIBELLE','G_GENERAL','G_GENERAL<>"'+'ymo'+'"','G_GENERAL');
    if St='' then exit;
    // Confirmation
    if MsgBox.Execute(57,'','')<>mrYes then Exit ;
    SetControlEnabled('BCopier',False);

    // Changement d'état si nécessaire
    if (not (DS.State in [dsEdit,dsInsert])) then DS.Edit;

          // Copie
    Q := OpenSQL('Select * From GENERAUX Where G_GENERAL="' + St + '"', TRUE) ;
    if not Q.Eof then
      begin
        // YMO 16/05/20006 FQ17889 Remplacement des SetControlText par SetField
        for i:=Low(DuplicValues) to High(DuplicValues) do
          SetField(DuplicValues[i], Q.FindField(DuplicValues[i]).AsString) ;

        // onglet tables libres
        for i:=0 to 9 do
          SetField('G_TABLE'+inttostr(i), Q.FindField('G_TABLE'+inttostr(i)).AsString) ;

        for i:=1 to 5 do
          SetField('G_VENTILABLE'+inttostr(i), Q.FindField('G_VENTILABLE'+inttostr(i)).AsString) ;

        if Q.FindField('G_CYCLEREVISION').AsString <> '' then
        begin
        SetControlVisible('BGCYCLEREVISION', True);
        SetControlProperty('GCYCLEREVISION', 'Caption',
                           GetColonneSql( 'RUBRIQUE', 'RB_LIBELLE',
                           'RB_RUBRIQUE= "' + GetField('G_CYCLEREVISION') + '" AND ' +
                           'RB_CLASSERUB = "CDR"')
                          );
      end;

      TCheckBox(GetControl('GCONFIDENTIEL',true)).Checked := (Q.FindField('G_CONFIDENTIEL').AsString = '1');

      {YMO 26/12/2006 La duplication de la ventilation est effectuée dès le choix du compte à copier ;
           FQ18924    Si on annule, les enregs dans VENTIL seront supprimés dans le OnClose}
      DuplicVentil(St);

      end ;
    Ferme(Q) ;
  end;
  SetControlEnabled('G_GENERAL',DS.State in [dsInsert]);

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/12/2006
Modifié le ... :   /  /
Description .. : Duplication des sections
Mots clefs ... : FQ18924
*****************************************************************}
procedure TOM_GENERAUX.DuplicVentil(StGene : string);
var
  TobVentil : Tob ;
  StVentil  : String ;
  QVentil   : TQuery ;
  i         : integer ;
begin

    StVentil := 'SELECT VENTIL.* FROM VENTIL WHERE VENTIL.V_NATURE LIKE "GE%'
               +'" AND VENTIL.V_COMPTE="'+StGene+'"';

    TobVentil := Tob.Create('', nil, -1);

    QVentil:=OpenSQL(StVentil,True) ;

    if not QVentil.Eof then
       TobVentil.LoadDetailDB('VENTIL','','', QVentil, False, False);

    Ferme(QVentil);

    for i:=0 to TobVentil.Detail.Count-1 do
        TobVentil.Detail[i].PutValue('V_COMPTE', GetControlText('G_GENERAL'));

    Tobventil.SetAllModifie(true);
    TobVentil.InsertDBByNivel(false) ;

    FreeAndNil(TobVentil);

end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/07/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_GENERAUX.OnKeyPressG_General(Sender: TObject; var Key: Char);
begin
  if EstSpecif('51205') then Exit;

  if not (key in ['a'..'z', 'A'..'Z', '0'..'9', #8, GetInfoCpta(fbgene).cb]) then
    Key := #0;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOM_GENERAUX.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

  // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
  if ( DS.State = dsInsert ) and
     ( GetControlText('G_GENERAL') = '' ) then Exit ;

  if Assigned( FAGLCloseQuery ) then
    FAGLCloseQuery(Sender, Canclose ) ;

end;

{JP 23/10/07 : FQ 21714 : Gestion de la mise des ventilations en multi dossiers
{---------------------------------------------------------------------------------------}
procedure MajVentilationMulti(LeCpte : string; VentilAxe : array of string; PreVentil : Boolean);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if not(CtxPCL in V_PGI.PGIContexte) and EstTablePartagee('GENERAUX') then begin
    Q := OpenSQL('SELECT DOS_NODOSSIER, DOS_SOCIETE, DOS_LIBELLE, DOS_VERSIONBASE, DOS_NOMBASE FROM DOSSIER', True);
    try
      if Q.EOF then
        {Théoriquement on ne devrait jamais passé par ici}
        MajVentilation(LeCpte, VentilAxe, PreVentil)
      else
        while not Q.EOF do begin
          MajVentilation(LeCpte, VentilAxe, PreVentil, Q.FindField('DOS_NOMBASE').AsString);
          Q.Next;
        end;
    finally
      Ferme(Q);
    end;
  end
  else
    MajVentilation(LeCpte, VentilAxe, PreVentil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
function TOM_GENERAUX.TrouveLiasseDuGene(vStCpte: string): string;
var lStSql : string;
begin
  lStSql := 'SELECT RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1 FROM RUBRIQUE WHERE ' +
            'RB_FAMILLES LIKE "%' + GetParamSocSecur('SO_CPCONTROLELIASSE', '', True) + '%" ' +
            'ORDER BY RB_RUBRIQUE';
  Result := TrouveGeneralSurTableRubrique( lStSql, vStCpte );

end;
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_GENERAUX.TrouveGeneralSurTableRubrique( vStSelectSurRubrique, vStCpte : string): string;
var lQuery : TQuery;
    lStWhereCompte1    : string;
    lStWhereExclusion1 : string;
    lStSqlGene         : string;
begin
  Result := '';
  if vStSelectSurRubrique = '' then Exit;
  try
    try
      lQuery := OpenSql( vStSelectSurRubrique, True);
      while not lQuery.Eof do
      begin
        lStWhereCompte1 := AnalyseCompte( lQuery.FindField('RB_COMPTE1').AsString, fbGene, False, False) ;

        lStWhereExclusion1 := AnalyseCompte( lQuery.FindField('RB_EXCLUSION1').AsString, fbGene, True, False);

        lStSqlGene := 'SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + vStCpte + '"';

        if lStWhereCompte1 <> '' then
          lStSqlGene := lStSqlGene + ' AND ' + lStWhereCompte1;

        if lStWhereExclusion1 <> '' then
          lStSqlGene := lStSqlGene + ' AND ' + lStWhereExclusion1;

        if ExisteSQL( lStSqlGene ) then
        begin
          Result := lQuery.FindField('RB_RUBRIQUE').AsString;
          Break;
        end;
        lQuery.Next;
      end;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : TrouveCycleDeRevision');
    end;

  finally
    Ferme( lQuery );
  end;
end;
{$ENDIF}


////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOM_GENERAUX ] ) ;
end.
