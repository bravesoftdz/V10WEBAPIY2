{***********UNITE*************************************************
Auteur  ...... : SANTUCCI Lionel
Créé le ...... : 11/02/2004
Modifié le ... : 05/04/2006 par Franck Vautrain
Description .. : Source TOF de la FICHE : BTSAISIECONSO ()
Mots clefs ... : TOF;BTSAISIECONSO
*****************************************************************}
Unit SaisieConsommations ;

Interface

Uses StdCtrls,
     Messages,
     Controls,
     Classes,
     ComCtrls,
     sysUtils,
     HTB97,
     HPanel,
     variants,
     NewCalendar,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fiche,
{$else}
     Maineagl,
     eMul,
     efiche,
{$ENDIF}
{$IFDEF BTP}
     CalcOLEGenericBTP,
{$ENDIF}
     EntGC,
     windows,
     graphics,
     menus,
     LookUp,
     uTob,
     utofAfBaseCodeAffaire,
     UtilSaisieConso,
     forms,
     Grids,
     HCtrls,
     HEnt1,
     HMsgBox,
     SelectPhase,
     UtilPhases,
     UTofRessource_Mul,
     BTSelectDate,
     Splash,
     SaisUtil,
     Vierge,
     UTOF,
     AffaireUtil,
     FactUtil,
     HRichOle,
     AppelsUtil,
     UtilBlocage;

Type
  TOF_BTSAISIECONSO = Class (TOF_AFBASECODEAFFAIRE)
  private
    CurrTOB 			: TOB;
    ActiveMode 		: TGrilleModeSaisie;
    QTE						: THEdit;
    DPA						: THEdit;
    DPR						: THEdit;
    DPV						: THEdit;
    MONTANTPA			: THEdit;
    MONTANTPR			: THEdit;
    MONTANTPV 		: THEdit;
    THRSSAISIE 		: THlabel;
    NBHRSSAISIE 	: THNumEdit;
    PlusWindow 		: TToolWindow97;
    RetourWindow	: TToolWindow97;
    WithPhases 		: boolean;
    //
    TOBMo				  : TOB;
    TOBFRS			  : TOB;
    TOBMAT			  : TOB;
    TOBMATERIAU	  : TOB;
    TOBRES			  : TOB;
    TOBMOEXT		  : TOB;
    TOBRECETTES		: TOB;
    TOBTYPEHEURE	: TOB;
    TOBConso      : TOB;
    TOBConsoOld   : TOB;
    TOBLIENEQUIPE : TOB;
    TOBLIENEQUIPE_O : TOB;

    // Dans le cas de saisie par ressource
    TOBRessource	: TOB;
    //
    PGTemps 			: TPageControl;
    Calendrier 		: TmxCalendar;

    CHOIX_CHANTIER	: TRadioButton;
    CHOIX_MO	  		: TRadioButton;
    CHOIX_RESSOURCE	: TRadioButton;
    CHOIX_MATERIAUX	: TRadioButton;
    CHOIX_MOEXT 		: TRadioButton;
    CHOIX_CONTRAT 	: TRadioButton;
    CHOIX_APPEL  		: TRadioButton;

    PWinPlus 	  	: THPanel;
    PDESCRIPTIF 	: THPanel;
    PHAUT					: THPanel;
    PSAISIE				: THPanel;
    PCALENDAR 		: THPanel;
    P_CHANTIER		: THPanel;
    P_RESSOURCE  	: THPanel;
    P_MO					: THPanel;
    P_MATERIAUX  	: THPanel;
    P_MOEXT 	  	: THPanel;

    PANELCONTRAT	: THPanel;
    PANELAPPEL		: THPanel;

    THEURES				:	TTabSheet;
    TFRAIS				: TTabSheet;
    TMATERIELS		: TTabSheet;
    TFOURNITURES	: TTabSheet;
    TMOEXT				: TTabSheet;
    TRECETTES 		: TTabSheet;
    POPMAINDOEUVRE: TPopupMenu;
    MOToutes 			: TMenuItem;

    BBlocNote		  : TToolbarButton97;
    BBlocNote1		: TToolbarButton97;
    BBlocNote2		: TToolbarButton97;

    BDuplic 		    : TToolbarButton97;
    BDuplicSemaine  : TToolbarButton97;
    BCherche		    : TToolbarButton97;
    Bdelete 		    : TToolbarButton97;
    BValider		    : TToolbarButton97;
    Bannul 		      : TToolbarButton97;
    BSelectAff1  	  : TToolBarButton97;
    BEffaceAff1  	  : TToolBarButton97;
    BRechPhase 	 	  : TToolbarButton97;
    BChangeEtat     : TToolbarButton97;
    BEffaceCha      : TToolbarButton97;

    //
    LastDayMo				: TdateTime;
    LastDayFrs			: TdateTime;
    LastDayMat			: TdateTime;
    LastDayFourn		: TdateTime;
    LastDayMoExt		: TdateTime;
    LastDayRecettes : TdateTime;

    // Les grilles de saisie
    GHeures				: THGrid;
    GFrais				: THGrid;
    GMAteriel			: THGrid;
    GFourniture		: THGrid;
    GMoext				: THGrid;
    GRecettes 		: THGrid;

    // Saisie par ressource
    TYPERESSOURCE	  : THValCOmboBox;
    TYPERESSOURCEEXT: THValCOmboBox;
    RESSOURCES		  : THEdit;
    RESSOURCESEXT 	: THEdit;
    LRESSOURCE		  : THLabel;
    LRESSOURCEEXT 	: THLabel;

    PHEURES				: THPanel;

    // Saisie par matériaux
    ARTICLE 			: THEdit;
    MATERIAUX 		: THEdit;
    LMATERIAUX 		: THLabel;

    // saisie sur MO
    MAINDOEUVRE 	: THEdit;
    LMAINDOEUVRE 	: THLabel;

    // saisie par chantier
    CH_CHANTIER  	: THEdit;
    CH_CHANTIER0 	: THEdit;
    CH_CHANTIER1 	: THEdit;
    CH_CHANTIER2 	: THEdit;
    CH_CHANTIER3 	: THEdit;
    CH_AVENANT 	 	: THEdit;
    LCHANTIER 	 	: THlabel;
    FONCTION 	  	: THValCOmboBox;
    LIBPHASE 	  	: THLabel;
    PHASE 	    	: THEdit;

    // Saisie Par Mo externe
    // -- gestion des grids
    stCols						: String;
    stColListe 				: String;
    FPerso 						: String;
    NomList						: String;
    FNomTable					: String;
    FLien							: String;
    FSortBy						: String;
    FTitre						: HString;
    FLargeur					: String;
    FAlignement				: String;
    FParams						: String;
    title							: HString;
    NC 								: HString;
    stprevHeure				: String;
    stPrevFrs					: String;
    StPrevMAT					: String;
    StPrevFOur				: String;
    stPrevMoext				: String;
    stPrevRecettes 		: String;
    PrestationDefaut	: String;
    CodeDomaine       : String;
    LibAffaire        : String;

    OkTri					: Boolean;
    OkNumCol 			: Boolean;
    OkSaisEquipe  : Boolean;
    nbcolsInListe : Integer;

    // -- Colonnes dans les grids
    G_SELECT			: integer;
    G_JOUR				: integer;
    G_RESSOURCE		: integer;
    G_CODEARTICLE	: integer;
    G_AFFAIRE0		: integer;
    G_AFFAIRE			: integer;
    G_PHASE				: integer;
    G_QTE					: integer;
    G_QUALIF			: integer;
    G_MONTANT			: integer;
    G_FACT				: integer;
    G_PU					: integer;
    G_NATURE			: integer;
    G_LIBELLE			: integer;
    G_TYPEHEURE 	: integer;

    // Zones ajoutés pour la saisie du retour des appels
    CODERESSOURCE : THEdit;
    CODETIERS     : THEdit;

    LCONTRAT		  : THLabel;
    LLABELHEURE		: THLabel;
    LLABELETAT		: THLabel;
    LLabelTiers   : THLabel;
    LQTEFACTURE		: THLabel;
    LLIBCONTRAT   : THLabel;
    QTEFACTURE		: THEdit;
    CONTRAT   		: THEdit;
    FACTURABLE		: TCheckBox;
    LPHASE				: THLabel;
    FAMILLETAXE1  : ThValComboBox;
    TypeSaisie    : String;
    //
    ValideSansCloture : Boolean;
    //

    // -- méthodes privées
    procedure BannulClick(Sender: TObject);
    procedure BChangeEtatClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BrechPhaseClick (Sender : TObject);
    procedure CreeTobs;
    procedure ChangeChantier (Sender : TObject);
    procedure ChargeResource (CodeResource: string);
    procedure CHOIX_CHANTIERClick(Sender: TObject);
    procedure CHOIX_CONTRATClick(Sender: TObject);
    procedure CHOIX_APPELClick(Sender: TObject);
    procedure CHOIX_MATERIAUXClick(Sender: TObject);
    procedure CHOIX_MOClick(Sender: TOBject);
    procedure CHOIX_RESSOURCEClick(Sender: TObject);
    procedure CHOIX_MOEXTClick(Sender: TOBject);
    function  ConstitueWhere: string;
    Procedure ContratClick(Sender: TObject);
    function  ControleSaisieOK: boolean;
    procedure DefiniGrilles;
    procedure EventGrilles(Active: boolean);
    procedure EffaceAffaireClick(Sender: TObject);
    procedure FONCTIONChange(Sender: Tobject);
    procedure GetComponents;
    procedure GetConso;
    procedure LibereTObs;
    procedure MAINDOEUVREExit(Sender: TOBject);
    procedure MAINDOEUVREElipsisClick(Sender: TOBject);
    procedure MATERIAUXRECH(Sender: TObject);
    procedure OnExitPartieAffaire (Sender : Tobject);
    procedure PhaseChange(Sender: TObject);
    procedure RESSOURCEExit(Sendre: TObject);
    procedure SaisieChantier;
    procedure SaisieMateriaux;
    procedure SaisieMO;
    procedure SaisieRessource;
    procedure SetChoix;
    procedure SetEvent;
    procedure SetFonction;
    procedure SetParamGrilles;
    procedure TYPEPRESTATIONCHANGE(Sender: TOBject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GestionColonneAffaire(TOBL : TOB);
    // Methode de la grille des heures de MO interne
    procedure GheuresCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GheuresCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
    procedure GheuresElipsisClick(Sender: TOBject);
    procedure GheuresRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GheuresRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GheuresPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

    // Methode de la grille des frais de MO interne
    procedure GFraisCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GFraisCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
    procedure GFraisElipsisClick(Sender: TOBject);
    procedure GFraisRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GFraisRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
    procedure GFraisPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

    // Methode de la grille des heures de materiels
    procedure GMaterielCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GMaterielCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GMaterielElipsisClick(Sender: TOBject);
    procedure GMaterielRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GMaterielRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
    procedure GMaterielPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

    // Methode de la grille des fournitures
    procedure GfournitureCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GfournitureCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
    procedure GfournitureElipsisClick(Sender: TOBject);
    procedure GFourniturePostDrawCell (ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GfournitureRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GfournitureRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);

    // Methode de la grille des MO externes
    procedure GMOextCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GMOextCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GMOextElipsisClick(Sender: TOBject);
    procedure GMOextRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GMOextRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GMOextPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

    // Methode de la grille des REcettes et frais annexes
    procedure GrecettesCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrecettesCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrecettesElipsisClick(Sender: TOBject);
    procedure GrecettesRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrecettesRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
    procedure GrecettesGetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GRecettesPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

    procedure Remplitgrilles;
    procedure AfficheLigne(Grid: THgrid; TOBL: TOB; indice: integer);
    function AddNewLigne(LaTOB: TOB; Themode : TGrilleModeSaisie) : TOB;
    procedure InitNewConso(TOBL: TOB; Themode : TGrilleModeSaisie);
    function ZoneAccessible(Grille: THgrid; var ACol, ARow: integer): boolean;
    procedure ZoneSuivanteOuOk(Grille: THGrid; var ACol, ARow: integer; var Cancel: boolean);
    function VerifAffaire(TheMode: TGrilleModeSaisie;TOBL: TOB; valeur: string): boolean;
    function VerifArticle(TOBL: TOB; valeur: string): boolean;
    function VerifDate(TOBL: TOB; valeur: string): boolean;
    function VerifFacturable(TOBL: TOB; valeur: string): boolean;
    function VerifPhase(TOBL: TOB; valeur: string): boolean;
    function VerifQTE(TOBL: TOB; Thevaleur: string): boolean;
    function VerifQualif(TOBL: TOB; valeur: string): boolean;
    function VerifRessource(TOBL: TOB; valeur: string): boolean;
    function ControleLigne(ligne : integer; Mode : TGrilleModeSaisie;var Acol : integer) : boolean;
    function LigneVide(Ligne : integer; Mode : TGrilleModeSaisie ) : boolean;
    procedure EnteteActive(status: boolean);
    function VerifMontant(TOBL: TOB; Thevaleur: string): boolean;
    procedure MATERIAUXExit(Sender: TObject);
    procedure ValideLesEcritures;
    procedure SetUniqueNumberForCreat;
    procedure SupprimeAncien;
    procedure EcritLesLignes;
    procedure PurgeLesLignes;
    procedure AfficheLibelle(Mode: TGrilleModeSaisie; Acol, Arow: integer);
    procedure BBlocNoteClick(Sender: TOBject);
    procedure ShowToolWindow(Name : String; Actif: boolean);
    procedure SetToolValue(Mode: TGrilleModeSaisie; Arow: integer); overload;
    procedure SetToolValue(TOBL: TOB); overload;
    procedure ChangeOnglet(Sender: TObject);
    procedure AvantChangeOnglet(Sender: TObject; var AllowChange : Boolean);
    procedure DPAExit(Sender: TObject);
    procedure DPRExit(Sender: TObject);
    procedure DPVExit(Sender: TObject);
    procedure MONTANTPAExit(Sender: TObject);
    procedure MONTANTPRExit(Sender: TObject);
    procedure MONTANTPVExit(Sender: TObject);
    procedure TYPEPRESTATIONEXTCHANGE(Sender: TOBject);
    procedure RESSOURCEEXTExit(Sendre: TObject);
    procedure SaisieMOExt;
    procedure PlusWindowClose (Sender : TObject);
    procedure RetourWindowClose (Sender : TObject);
    procedure BdeleteClick (sender : Tobject);
    procedure DeleteLigne(Mode: TGrilleModeSaisie);
    procedure initgrilles;
    procedure GridKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    function VerifPu(TOBL: TOB; Thevaleur: string): boolean;
    procedure GridElipsisClick(Sender: Tobject);
    procedure dupliqueDonnee;
    procedure BDuplicClick (Sender : Tobject);
    procedure BDuplicSemaineClick (Sender : Tobject);
    function AnnulOk: boolean;
    function SortieOk: boolean;
    function ExisteValorisation(TOBL: TOB): boolean;
    function VerifNature(TOBL: TOB; valeur: string): boolean;
    function VerifTypeHeure(TOBL: TOB; Valeur: string): boolean;
    procedure RecupTypeHeure;
    procedure  PositionneValeurInit (TOBConso : TOB);
    procedure PositionneDansGrid(TheGrid: THGrid; Arow, Acol: integer;Cancel : Boolean = False);
    procedure GridsPostDrawCell(grid: Thgrid; TOBATrait: TOB; ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    function Controlgrids : boolean;
    function GetcumulTemps : double;

    //Procedure et fonction créées par FV.
    procedure GestionAffichageContrat(VraiFaux: Boolean);
    function GetLibelleContrat(Contrat: string): string;
    function VerifNatAffaire(TOBL: TOB; valeur: string): boolean;
    procedure QTEFACTUREExit(Sender: TObject);
    procedure CocheFact(Sender: TObject);

    procedure RechArticleGrid(Grille: THGrid; TOBL: TOB);
    procedure AfficheNature(Grille: THGrid);
    Procedure RechAffaireGrid(Grille: THGrid; TOBL: TOB);
    procedure RechPhaseGrid(Grille: THGrid; TOBL: TOB);
    procedure RechPrestationGrid(Grille: THGrid; TOBL: TOB);
    procedure RechRessourceGrid(Grille: THGrid; StArg : String);
    procedure RechDateGrid(Grille: THGrid);
    procedure ControleGrille(TheMode: TGrilleModeSaisie;Grille: THGrid; TOBL: TOB; Acol,Arow: Integer; var Cancel: Boolean);
    procedure RechFraisGrid(Grille: THGrid; TOBL: TOB; StArg: String);

    //Modif FV
    procedure CellExitGrid(TOBL: TOB; Grille: THGrid; Acol, Arow: Integer; TheMode: TGrilleModeSaisie; var StPrev: String; var Cancel: Boolean);
    procedure RowEnterGrid(Ou: Integer; Grille: THGrid; TOBGrille: Tob; TheMode: TGrilleModeSaisie; var Cancel: Boolean);
    procedure RowExitGrid(Ou: Integer; Grille: THGrid; TheMode: TGrilleModeSaisie; var Cancel: Boolean);
    procedure ControleChangeOnglet(Grille: THGrid; var Cancel: Boolean);
    procedure ChargementFichierOLE(Statut, Etat, Appel: string);
    procedure ValideDesAppel(DateInt: TDateTime; Affaire: String); Overload;
    procedure ValideDesAppel(TOBLC : TOB); Overload;
    procedure ValideLesAppel(DateInt: TDateTime; TOBDet : TOB);
    procedure GestionCheckedClick(ok_aff: boolean);
    procedure BDUPLICCONSOClick (Sender : TObject);
    procedure BEFFACECHAClick (Sender : TObject);
    function GetConsoADuplic(Ressource: string): boolean;
    procedure ChangeRessourceEtVAlorisation(TOBL: TOB);
    procedure BVALIDEPLUSClick (Sender : TObject);
    procedure TraiteStatusMoisOD(TOBATRAIT: TOB);
    function ChargementRequeteConso(DateDebut,DateDeFin: Tdatetime): String;

    // traitement par equipe
		procedure TraiteLigneEquipe (TOBL : TOB);
    procedure UpdConsoFromLig (grid : Thgrid;TOBLC,TOBL : TOB);
    procedure AddDataEquipe(TOBL ,TOBEquipe : TOB; Equipe : string);
		procedure AddDataEquipeExt(TOBL ,TOBEquipe : TOB; Equipe : string);
    procedure AddConsoFromLig(Nature: string; TOBRef, TOBL, TOBRess: TOB;ModeS: TGrilleModeSaisie; grid : Thgrid; Equipe : string);
    procedure DeleteLigneEquipe(LinkEquipe: string);
    procedure DelLigne(Mode: TGrilleModeSaisie; TheGrid: Thgrid; LaTOB,OneTOB: TOB);
    function ExisteLigneEquipe(TOBL: TOB; LinkEquipe: string): boolean;
    procedure UpdatelesLignesEquipe(TOBL: TOB; LinkEquipe : string);
    procedure PositionnePrevValeur(TheGrid: THGrid; Acol, Arow: integer);
    procedure RecupInfoEquipe(TOBC: TOB);
    procedure UpdateLignesEquipeExt(TOBL, TOBLE: TOB);
    procedure DeleteLigneEquipeExt(LinkEquipe: string);
    procedure SetGridDefault;
    procedure FAMILLETAXEChange (Sender : TObject);
    procedure ChangeMaterielEtVAlorisation(TOBL: TOB);
    //FV1 - 28/12/20015 : FS#1832 - CODRIS : Pb en saisie du code chantier
    procedure ChargeAffaire;
    procedure OnClickSelectAff(Sender: Tobject);
    procedure ControleSaisieEquipe;
    procedure MAJDescriptifInterv(Affaire: String);
  public
    Action : TActionFiche;

    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

procedure SaisieConsommation;
procedure SaisieIntervention;

Implementation
uses  BtpUtil,
      BTGENODANAL_TOF,
      Paramsoc,
      uTOFComm,
      DateUtils,
      UdateUtils;

procedure SaisieConsommation;
begin
  AGLLanceFiche('BTP','BTSAISIECONSO','','','MODIFICATION') ;
end;

procedure SaisieIntervention;
begin
  AGLLanceFiche('BTP','BTSAISIECONSO','','','MODIFICATION;INTERVENTION') ;
end;

procedure TOF_BTSAISIECONSO.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_, Aff1_, Aff2_,Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff  := THEdit(GetControl('CH_CHANTIER'));
Aff0 := THEdit(GetControl('CH_CHANTIER0'));
Aff1 := THEdit(GetControl('CH_CHANTIER1'));
Aff2 := THEdit(GetControl('CH_CHANTIER2'));
Aff3 := THEdit(GetControl('CH_CHANTIER3'));
Aff4 := THEdit(GetControl('CH_AVENANT'));
end;

procedure TOF_BTSAISIECONSO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIECONSO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIECONSO.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    Case Key of
        VK_F9 : BEGIN
                    Key := 0 ;
                    if TToolBarButton97(GetControl('BCHERCHE')).visible = true then
											BchercheClick(Sender);
                END;
    END;
end;

procedure TOF_BTSAISIECONSO.SetUniqueNumberForCreat;
var Indice,II : integer;
    TOBL      : TOB;
    TOBLE     : TOB;
    TheRetour : TGncERROR;
    UnNumero  : double;
begin

  for Indice := 0 to TOBMoEXT.detail.count -1 do
  begin
    TOBL := TOBMOEXT.detail[Indice];
    if TOBL.GetValue('NEW')='X' Then
    begin
      TheRetour := GetNumUniqueConso (UnNumero);
      if TheRetour = gncAbort then
      BEGIN
        V_PGI.IOError := oeUnknown;
        Exit;
      END;
      TOBL.putValue ('BCO_NUMMOUV',UnNumero);
    end;
  end;

  for Indice := 0 to TOBMo.detail.count -1 do
  begin
    TOBL := TOBMO.detail[Indice];
    if TOBL.GetValue('NEW')='X' Then
    begin
      TheRetour := GetNumUniqueConso (UnNumero);
      if TheRetour = gncAbort then
         BEGIN
         V_PGI.IOError := oeUnknown;
         Exit;
         END;
      TOBL.putValue ('BCO_NUMMOUV',UnNumero);
    end;
  end;

  for Indice := 0 to TOBFrs.detail.count -1 do
  begin
    TOBL := TOBFrs.detail[Indice];
    if TOBL.GetValue('NEW')='X' Then
    begin
      TheRetour := GetNumUniqueConso (UnNumero);
      if TheRetour = gncAbort then BEGIN V_PGI.IOError := oeUnknown; Exit; END;
      TOBL.putValue ('BCO_NUMMOUV',UnNumero);
    end;
  end;

  for Indice := 0 to TOBMat.detail.count -1 do
  begin
    TOBL := TOBMat.detail[Indice];
    if TOBL.GetValue('NEW')='X' Then
    begin
      TheRetour := GetNumUniqueConso (UnNumero);
      if TheRetour = gncAbort then BEGIN V_PGI.IOError := oeUnknown; Exit; END;
      TOBL.putValue ('BCO_NUMMOUV',UnNumero);
    end;
  end;

  for Indice := 0 to TOBres.detail.count -1 do
  begin
    TOBL := TOBres.detail[Indice];
    if TOBL.GetValue('NEW')='X' Then
    begin
      TheRetour := GetNumUniqueConso (UnNumero);
      if TheRetour = gncAbort then BEGIN V_PGI.IOError := oeUnknown; Exit; END;
      TOBL.putValue ('BCO_NUMMOUV',UnNumero);
    end;
  end;

  for Indice := 0 to TOBRecettes.detail.count -1 do
  begin
    TOBL := TOBRecettes.detail[Indice];
    if TOBL.GetValue('NEW')='X' Then
    begin
      TheRetour := GetNumUniqueConso (UnNumero);
      if TheRetour = gncAbort then BEGIN V_PGI.IOError := oeUnknown; Exit; END;
      TOBL.putValue ('BCO_NUMMOUV',UnNumero);
    end;
  end;

  //FV1 : 01/03/2016 - FS#1912 - TEAM RESEAUX : Pb en saisie consommations avec option "par équipe"
  if (not CHOIX_CHANTIER.Checked) And (not CHOIX_CONTRAT.Checked) And (not CHOIX_APPEL.Checked) then ControleSaisieEquipe;

  for Indice := 0 to TOBLIENEQUIPE.detail.count -1 do
  begin
    TOBLE := TOBlienEquipe.detail[Indice];
    For II := 0 to TOBLE.detail.count -1 do
    begin
      TOBL := TOBLE.detail[II];
      if TOBL.GetValue('NEW')='X' Then
      begin
        TheRetour := GetNumUniqueConso (UnNumero);
        if TheRetour = gncAbort then
        Begin
          V_PGI.IOError := oeUnknown;
          Exit;
        end;
        TOBL.putValue ('BCO_NUMMOUV',UnNumero);
      end;
    end;
  end;

end;

Procedure TOF_BTSAISIECONSO.ControleSaisieEquipe;
Var Indice, II  : Integer;
    Resref      : String;
    TOBL        : TOB;
    Ok_Equipe   : Boolean;
    NbLien      : Integer;
begin

  if CHOIX_MO.Checked then
    ResRef := MAINDOEUVRE.text
  else if CHOIX_MOEXT.Checked then
    ResRef := MAINDOEUVRE.text
  else if CHOIX_RESSOURCE.Checked then
    ResRef := RESSOURCES.text
  else if CHOIX_MATERIAUX.Checked then
    ResRef := MATERIAUX.text
  else
    ResRef := '';

  if Resref = '' then Exit;

  Ok_Equipe := True;

  //Parcour de la table des Equipe pour s'assurer qu'il y a soit une création soit une modification...
  TOBL := TOBlienEquipe.FindFirst(['NEW'], ['X'], True);
  if TOBL = nil then
  begin
    TOBL := TOBlienEquipe.FindFirst(['MODIF'], ['X'], True);
    if TOBL <> nil then
    begin
      if pgiAsk('Voulez-vous appliquer ces modifications à l''ensemble de l''équipe', 'Gestion Equipe') = Mrno then Ok_Equipe := False;
    end;
  end
  else
  begin
    if pgiAsk('Voulez-vous appliquer ces saisies à l''ensemble de l''équipe', 'Gestion Equipe') = Mrno then Ok_Equipe := False;
  end;

  If Ok_Equipe then Exit;

  TOBLIENEQUIPE.ClearDetail;

end;

procedure TOF_BTSAISIECONSO.SupprimeAncien;
var Indice,II : Integer;
		TOBL      : TOB;
begin

  if not TOBConsoOld.DeleteDB (false) then
  Begin
    V_PGI.ioError := oeUnknown;
    Exit;
  End;

  if TOBLIENEQUIPE = nil then exit;

  if TOBLIENEQUIPE.detail.count > 0 then
  begin
    for Indice := 0 to TOBLIENEQUIPE.detail.count -1 do
    begin
      For II := 0 to TOBLIENEQUIPE.detail[Indice].detail.count -1 do
      begin
        TOBL := TOBLIENEQUIPE.detail[Indice].detail[II];
        TOBL.DeleteDB(false);
      end;
    end;
  end;

end;

procedure TOF_BTSAISIECONSO.PurgeLesLignes;
var indice : integer;
    TOBL : TOB;
begin

  Indice := 0;
  if TOBMOEXT.detail.count > 0 then
  begin
    repeat
      TOBL := TOBMOEXT.detail[Indice];
      if LigneVide (indice+1,TgsMoext) then TOBL.free else inc(Indice);
    until indice > TOBMOEXT.detail.count -1;
  end;

  Indice := 0;
  if TOBMO.detail.count > 0 then
  begin
    repeat
      TOBL := TOBMO.detail[Indice];
      if LigneVide (indice+1,TgsMO) then TOBL.free else inc(Indice);
    until indice > TOBMO.detail.count -1;
  end;

  Indice := 0;
  if TOBMat.detail.count > 0 then
  begin
    repeat
      TOBL := TOBMAT.detail[Indice];
      if LigneVide (indice+1,TgsFOURN) then TOBL.free else inc(Indice);
    until indice > TOBMat.detail.count -1;
  end;

  Indice := 0;
  if TOBFrs.detail.count > 0 then
  begin
    repeat
      TOBL := TOBFrs.detail[Indice];
      if LigneVide (indice+1,TgsFRS ) then TOBL.free else inc(Indice);
    until indice > TOBFrs.detail.count -1;
  end;

  Indice := 0;
  if TOBRes.detail.count > 0 then
  begin
    repeat
      TOBL := TOBRES.detail[Indice];
      if LigneVide (indice+1,TgsRES) then TOBL.free else inc(Indice);
    until indice > TOBRes.detail.count -1;
  end;

  Indice := 0;
  if TOBRecettes.detail.count > 0 then
  begin
    repeat
      TOBL := TOBRecettes.detail[Indice];
      if LigneVide (indice+1,TgsRecettes) then TOBL.free else inc(Indice);
    until indice > TOBRecettes.detail.count -1;
  end;

end;

procedure TOF_BTSAISIECONSO.TraiteStatusMoisOD (TOBATRAIT :TOB);
var Indice : Integer;
begin
  for indice := 0 to TOBATRAIT.Detail.Count -1 do
  begin
  	UpdateStatusMoisOD (TOBATRAIT.detail[Indice]);
  end;
end;

procedure TOF_BTSAISIECONSO.EcritLesLignes;
var TOBL    : Tob;
    indice  : integer;
    Ind_2   : Integer;
    DateNow : TDateTime;
    ResRef  : String;
begin

  if CHOIX_MO.Checked then
    ResRef := MAINDOEUVRE.text
  else if CHOIX_MOEXT.Checked then
    ResRef := MAINDOEUVRE.text
  else if CHOIX_RESSOURCE.Checked then
    ResRef := RESSOURCES.text
  else
    ResRef := '';

  TOBMOEXT.SetAllModifie(true);
  if not TOBMOEXT.InsertDB (nil,false) Then
  BEGIN
    V_PGI.IOError := oeUnknown;
    Exit;
  END;

  if GetparamSocSecur ('SO_OPTANALCONSO',false) then
  begin
    TraiteStatusMoisOD (TOBMOEXT);
    if V_PGI.IOError <> oeOk then Exit;
  end;

  TOBMO.SetAllModifie(true);
  if not TOBMO.InsertDB (nil,false) Then
  BEGIN
    V_PGI.IOError := oeUnknown;
    Exit;
  END;

  if (GetParamSocSecur('SO_AFLIENPAIEVAR',False)) and (GetparamSocSecur ('SO_OPTANALCONSO',false)) then
  begin
    TraiteStatusMoisOD (TOBMO);
    if V_PGI.IOError <> oeOk then Exit;
  end;

  TOBMAt.SetAllModifie(true);
  if not TOBMat.InsertDB (nil,false) Then
  BEGIN
    V_PGI.IOError := oeUnknown;
    Exit;
  END;

  if (GetparamSocSecur ('SO_OPTANALCONSO',false)) then
  begin
    TraiteStatusMoisOD (TOBMAT);
    if V_PGI.IOError <> oeOk then Exit;
  end;

  TOBFrs.SetAllModifie(true);
  if not TOBFrs.InsertDB (nil,false) Then
  BEGIN
    V_PGI.IOError := oeUnknown;
    Exit;
  END;

  if (GetParamSocSecur('SO_AFLIENPAIEVAR',False)) and (GetparamSocSecur ('SO_OPTANALCONSO',false)) then
  begin
    TraiteStatusMoisOD (TOBFrs);
    if V_PGI.IOError <> oeOk then Exit;
  end;


  TOBRES.SetAllModifie(true);
  if not TOBRES.InsertDB (nil,false) Then
  BEGIN
    V_PGI.IOError := oeUnknown;
    Exit;
  END;

  TOBRecettes.SetAllModifie(true);
  for Indice := 0 to TOBRecettes.detail.count -1 do
  begin
    TOBL := TOBRecettes.detail[Indice];
    if TOBL.GetValue('BCO_NATUREMOUV')='RAN' Then
    begin
      TOBL.PutValue('BCO_QUANTITE',TOBL.GetValue('BCO_QUANTITE') * -1);
      calculeLaLigne (TOBL);
    end;
  end;

  //FV1 : 16/12/2015 -FS#1784 - TEAM RESEAUX : Si saisie avec option "par équipe", message d'erreur et pas d'enregistrement
  for Indice := 0 to TOBLIENEQUIPE.detail.count -1 do
  begin
    TOBL := TOBLIENEQUIPE.detail[Indice];
    For Ind_2 := 0 to TOBL.Detail.count -1 do
    begin
      if TOBL.detail[ind_2].GetString('BCO_RESSOURCE') = ResRef then
        Continue
      else
      begin
        If TOBL.detail[Ind_2].GetBoolean('SUPPRESSION') then
        begin
          TOBL.Detail[Ind_2].DeleteDB(False);
        end
        else
        begin
          if Not TOBL.detail[Ind_2].InsertDB(nil,false) then
          begin
            V_PGI.IOError := oeUnknown;
            Exit;
          end;
        end;
      end;
    end;
  end;

  if not TOBRecettes.InsertDB (nil,false) Then
  Begin
    V_PGI.IOError := oeUnknown;
    Exit;
  End;

  //Mise à jour de l'affaire dans le cas d'un appel.
  if Choix_Appel.Checked then
  Begin
    ValideDesAppel(Now, CH_CHANTIER.text);
  End
  else
  Begin
    ValideLesAppel (Now,TOBMO);
    ValideLesAppel (Now,TOBMOEXT);
    ValideLesAppel (Now,TOBFrs);
  End;

  MAJDescriptifInterv(CH_CHANTIER.text);

  ValideSansCloture := false;

end;

Procedure TOF_BTSAISIECONSO.ValideDesAppel(TOBLC : TOB);
Var TOBAff : TOB;
    Req    : String;
    QQ     : TQuery;
Begin

  TobAff := Tob.Create('AFFAIRE',Nil, -1);

  Req := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE ="' + TOBLC.GetValue('BCO_AFFAIRE') + '"';

  QQ := OpenSql (Req,true,-1,'',true);

  if Not QQ.Eof then
  Begin
    TobAff.SelectDB ('',QQ);
    TobAff.PutValue('AFF_ETATAFFAIRE', TOBLC.GetVALUE('ETATAFFAIRE'));
    TobAff.PutValue('AFF_DATEFIN', Now);
    TobAff.UpdateDB(false);
  end;

  Ferme (QQ);
  TOBAff.free;

end;

Procedure TOF_BTSAISIECONSO.ValideDesAppel(DateInt : TDateTime; Affaire : String);
Var TobOLE : Tob;
    Motif  : THRichEditOLE;
    TOBApp : Tob;
    QQ		 : TQuery;
    Req		 : String;
begin

  TobApp := Tob.Create('AFFAIRE',Nil, -1);

  Req := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE ="' + Affaire + '"';
  QQ := OpenSql (Req,true,-1,'',true);
  if Not QQ.Eof then
  Begin
    TobApp.SelectDB ('',QQ);
    //FV1 : 22/12/2015 - FS#1818 - VEODIS : saisie retour d'interv., ne pas autoriser la validation temporaire d'un appel facturé.
    if ValideSansCloture then
    begin
      if (TobApp.GetString('AFF_ETATAFFAIRE') = 'FAC') then    // En cours de réalisation
      else if (TobApp.GetString('AFF_ETATAFFAIRE') = 'REA') then    //FV1 : 09/06/2017 - FS#2585 - VEODIS GROUP : Bon de Commande N° GB-170515-1
        TobApp.PutValue('AFF_ETATAFFAIRE','ECR')
      else if (TobApp.GetString('AFF_ETATAFFAIRE') = 'ECR') then    //FV1 : 09/06/2017 - FS#2585 - VEODIS GROUP : Bon de Commande N° GB-170515-1
        TobApp.PutValue('AFF_ETATAFFAIRE','ECR')
      else
        TobApp.PutValue('AFF_ETATAFFAIRE','ECR');
    end
    else
    //FV1 : 10/02/2015 - FS#1898 - DELABOUDINIERE : après un retour d'intervention, l'appel ne passe plus en réalisé
    begin
      if (TobApp.GetString('AFF_ETATAFFAIRE') <> 'FAC') AND (TobApp.GetString('AFF_ETATAFFAIRE') <> 'TER') then
        TobApp.PutValue('AFF_ETATAFFAIRE','REA');       // Réalisé
    end;
    TobApp.PutValue('AFF_affaireinit', Contrat.text);
    TobApp.PutValue('AFF_DATEFIN', DateInt);
    TobApp.UpdateDB(false);
  end;

  Ferme (QQ);
  TOBApp.free;

end;

Procedure TOF_BTSAISIECONSO.MAJDescriptifInterv(Affaire : String);
Var TobOLE : Tob;
    Motif  : THRichEditOLE;
Begin

  //Mise à jour du descriptif de l'intervention
  Motif := THRichEditOLE(GetControl('RETOURINT'));

  if (Length(Motif.Text) = 0) or (Motif.Text = #$D#$A)  then exit;

  //Chargement des Tob liens OLE
  TobOle := Tob.Create('LIENSOLE' ,Nil, -1);

  TobOle.PutValue('LO_TABLEBLOB', 'APP');
  TobOle.PutValue('LO_QUALIFIANTBLOB', 'MOT');
  TobOle.PutValue('LO_EMPLOIBLOB', 'REA');
  TobOle.PutValue('LO_IDENTIFIANT', Affaire);
  TobOle.PutValue('LO_RANGBLOB', 1);

  TobOle.PutValue('LO_LIBELLE', 'Intervention Appel ' + Affaire + 'le ' + DateToStr(Now));

  TobOle.PutValue('LO_PRIVE', '-');
  TobOle.PutValue('LO_DATEBLOB', Now);
  TobOle.PutValue('LO_OBJET', GetControlText('RETOURINT'));

  TobOle.SetAllModifie(true);
  TobOle.InsertOrUpdateDB(true);
  TobOle.free;

  Motif.Text := '';

end;

procedure TOF_BTSAISIECONSO.ValideLesEcritures;
begin
  SupprimeAncien;
  if V_PGI.IOError = OeOk then PurgeLesLignes;
  if V_PGI.IOError = OeOk then EcritLesLignes;
end;

procedure TOF_BTSAISIECONSO.OnUpdate ;
var  splashScreen: TFsplashScreen;
begin
  Inherited ;

  RetourWindow.Visible := false;

  IF TFvierge(ecran).ActiveControl.Parent.Name <> 'PHEURES' then
    SendMessage(TFvierge(ecran).ActiveControl.Handle , WM_KeyDown, VK_TAB, 0);

  if not Controlgrids then
  BEGIN
  	ecran.ModalResult := mrNone;
    exit;
  END;

  V_PGI.IOerror := OeOk;
  splashScreen := TFsplashScreen.Create (application);
  splashScreen.Caption := Ecran.Caption;
  splashScreen.Label1.Caption := TraduireMemoire('Validation des consommations');
  splashScreen.Show;
  splashScreen.Update;


  TRY
	  // CORRECTIONS : FQ 11904
	  if not VerrouilleValidation ('XXX', 'XXX;00') then exit;

    if Transactions (SetUniqueNumberForCreat,1) <> oeOk then
    begin
      PGIBox (TraduireMemoire('Erreur durant la numérotation'),ecran.Caption);
      ecran.ModalResult := mrNone;
      exit;
    end;

    if Transactions (ValideLesEcritures,1) = oeOK then
    begin
      TOBMo.ClearDetail;
      TOBMoEXT.ClearDetail;
      TOBFRS.ClearDetail;
      TOBMAT.ClearDetail;
      TOBRES.ClearDetail;
      TOBRecettes.ClearDetail;
      TOBConso.ClearDetail; TOBConsoOld.clearDetail;
      TOBLIENEQUIPE.ClearDetail;
      TOBLIENEQUIPE_O.ClearDetail;
      //
      EnteteActive (True);
      EventGrilles(false);
      //
      GHeures.VidePile (false);
      GFrais.videpile (false);
      GMAteriel.videpile (false);
      GFourniture.videpile (false);
      GMOEXT.videpile (false);
      Grecettes.VidePile (false);
    end else
    begin
      PGIBox (TraduireMemoire('Erreur durant la validation'),ecran.Caption);
      ecran.ModalResult := mrNone;
    end;
  FINALLY
    splashScreen.Free;
    PSaisie.visible := false;
    // CORRECTIONS : FQ 11904
    DeverouilleValidation('XXX');
  END;

  THRSSAISIE.Visible := false;
  NBHRSSAISIE.Visible := false;
  ecran.ModalResult := mrNone;

  BChangeEtat.visible := False;

  exit;

end ;

procedure TOF_BTSAISIECONSO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIECONSO.OnArgument (S : String ) ;
begin

  Inherited ;

	AppliqueFontDefaut (THRichEditOle(GetControl('RETOURINT')));

  TypeSaisie := Copy(S,14,12);
  OkSaisEquipe := GetParamSocSecur('SO_BTSAISEQUIPE',false);

  CreeTobs;
  RecupTypeHeure;
  GetComponents;
  SetEvent;
  EffaceAffaireClick (self);
  SetFonction;

  //FV1 : 26/08/2013 - FS#564 - VERRE & METAL - Saisie de consommations / ressource de type matériel fermée
  TYPEPRESTATIONCHANGE(self);
  TYPEPRESTATIONEXTCHANGE(self);

  //FV1 : 10/07/2017 - FS#2614 - VEODIS : En saisie Intervenant passage de REA=>ECR et ECR=>REA
  If BChangeEtat <> nil then BChangeEtat.Visible := False;

  //
  SetparamGrilles;
	EnteteActive (True);

  If TypeSaisie = 'INTERVENTION' then
     Choix_Appel.Checked := True
  else
  begin
     Choix_Chantier.Checked :=True;
     CHOIX_CHANTIERClick(Self);
  end;
  //
  TToolBarButton97(GetControl('BDUPLICCONSO')).visible := false;

  //  TToolBarButton97(GetControl('BDUPLIC')).visible      := false;
  If Assigned(Getcontrol('BDuplicSemaine')) then BDuplicSemaine.visible := false;
  //
  (*
  TToolBarButton97(GetControl('BVALIDER')).visible := false;
  TToolBarButton97(GetControl('BCHERCHE')).visible := true;
  TToolBarButton97(GetControl('BANNUL')).visible := false;
  TToolBarButton97(GetControl('BDELETE')).visible := false;
  *)

end ;

procedure TOF_BTSAISIECONSO.OnClose ;
begin
  Inherited ;

  if (not AnnulOk) and (not SortieOk) then
  begin
    ecran.ModalResult := 0; //
    exit;
  end;

  Calendrier.free;
  PlusWindow.Free;
  RetourWindow.Free;
  LibereTObs;

end ;

procedure TOF_BTSAISIECONSO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIECONSO.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTSAISIECONSO.AnnulOk : boolean;
var indice : integer;
    TOBL : TOB;
begin
  result := true;
  for Indice := 0 to TOBMo.detail.count -1 do
  begin
    TOBL := TOBMo.detail[indice];
    if (TOBL.GetValue('NEW')='X') and (not LigneVide (indice+1,TgsMO)) then
    begin
      result := false;
      Exit;
    end;
    if TOBL.IsOneModifie (false) then
    BEGIN
      result := false;
      Exit;
    END;
  end;

  for Indice := 0 to TOBFrs.detail.count -1 do
  begin
    TOBL := TOBFrs.detail[indice];
    if (TOBL.GetValue('NEW')='X') and (not LigneVide (indice+1,TgsFRS)) then
    begin
      result := false;
      Exit;
    end;
    if TOBL.IsOneModifie (false) then
    BEGIN
      result := false;
      Exit;
    END;
  end;

  for Indice := 0 to TOBMOEXT.detail.count -1 do
  begin
    TOBL := TOBMoext.detail[indice];
    if (TOBL.GetValue('NEW')='X') and (not LigneVide (indice+1,TgsMoext)) then
    begin
      result := false;
      Exit;
    end;
    if TOBL.IsOneModifie (false) then
    BEGIN
      result := false;
      Exit;
    END;
  end;

  for Indice := 0 to TOBres.detail.count -1 do
  begin
    TOBL := TOBRES.detail[indice];
    if (TOBL.GetValue('NEW')='X') and (not LigneVide (indice+1,TgsRES)) then
    begin
      result := false;
      Exit;
    end;
    if TOBL.IsOneModifie (false) then
    BEGIN
      result := false;
      Exit;
    END;
  end;

  for Indice := 0 to TOBMAT.detail.count -1 do
  begin
    TOBL := TOBMat.detail[indice];
    if (TOBL.GetValue('NEW')='X') and (not LigneVide (indice+1,TgsFOURN)) then
    begin
      result := false;
      Exit;
    end;
    if TOBL.IsOneModifie (false) then
    BEGIN
      result := false;
      Exit;
    END;
  end;

  for Indice := 0 to TOBrecettes.detail.count -1 do
  begin
    TOBL := TOBrecettes.detail[indice];
    if (TOBL.GetValue('NEW')='X') and (not LigneVide (indice+1,TgsRecettes)) then
    begin
      result := false;
      Exit;
    end;
    if TOBL.IsOneModifie (false) then
    BEGIN
      result := false;
      Exit;
    END;
  end;
end;

function TOF_BTSAISIECONSO.SortieOk : boolean;
begin
  result := (PGIAsk(TraduireMemoire('Confirmez-vous l''abandon de la saisie ?'), ecran.caption) = mryes);
end;

procedure TOF_BTSAISIECONSO.EnteteActive(status : boolean);
begin

	ValideSansCloture := false;

  THPanel(GetControl('PSEL')).Enabled  := status;
  THPanel(GetControl('PHAUT')).Enabled := status;
  //
  if ((TypeSaisie = 'INTERVENTION')  and (TToolbarButton97 (GetControl('BVALIDEPLUS')) <> nil)) OR
  	 ((TypeSaisie <> 'INTERVENTION') and (CHOIX_APPEL.Checked) and (TToolbarButton97 (GetControl('BVALIDEPLUS')) <> nil)) then
  begin
		TToolbarButton97 (GetControl('BVALIDER')).Hint        := 'Enregistrer et clôturer';
		TToolbarButton97 (GetControl('BVALIDEPLUS')).OnClick  := nil;
    TToolbarButton97 (GetControl('BVALIDEPLUS')).Visible  := (Not status);
    if TToolbarButton97 (GetControl('BVALIDEPLUS')).Visible then
    begin
      TToolbarButton97 (GetControl('BVALIDEPLUS')).OnClick := BVALIDEPLUSClick;
    end;
  end else
  begin
		TToolbarButton97 (GetControl('BVALIDER')).Hint  := 'Enregistrer';
  end;
  //
  TToolBarButton97(GetControl('BVALIDER')).visible  := (not Status);
  TToolBarButton97(GetControl('BFERME')).visible    := (Status);
  TToolBarButton97(GetControl('BCHERCHE')).visible  := (Status);
  TToolBarButton97(GetControl('BANNUL')).visible    := (not Status);
  TToolBarButton97(GetControl('BDELETE')).visible   := (not Status);
  TToolBarButton97(GetControl('BDELETE')).enabled   := (not Status);
  TToolBarButton97(GetControl('BDUPLIC')).visible   := (not Status);

  if (CHOIX_CHANTIER.Checked)   OR
     (CHOIX_CONTRAT.Checked)    OR
     (CHOIX_APPEL.Checked)      OR
     (CHOIX_MATERIAUX.Checked)  OR
     (CHOIX_MOEXT.Checked)      then
  begin
    TToolBarButton97(GetControl('BDUPLICCONSO')).visible := False;
  end
  else
  begin
    TToolBarButton97(GetControl('BDUPLICCONSO')).visible := True;
  end;

  If TypeSaisie = 'INTERVENTION' then
  Begin
     Choix_MO.Caption         := 'Intervenant';
     Choix_MOExt.Visible      := false;
     Choix_Ressource.Visible  := false;
     Choix_Materiaux.Visible  := false;
  end
  else
  Begin
     if Not VH_GC.BTSeriaContrat then
     Begin
        Choix_Contrat.Visible := False;
        PANELCONTRAT.Visible  := false;
     end;
     if Not VH_GC.BTSeriaIntervention then
     Begin
        Choix_Appel.visible   := False;
        PANELAPPEL.Visible    := False;
		 End;
  end;

  if Choix_MO.Checked then
  begin
    SetControlEnabled ('BDUPLICCONSO',not Status)
  end
  else if CHOIX_RESSOURCE.Checked then SetControlEnabled ('BDUPLICCONSO',not Status);

end;

procedure TOF_BTSAISIECONSO.InitNewConso (TOBL : TOB; Themode : TGrilleModeSaisie);
var year,Month,day : Word;
    WithMajPrix : boolean;
    Aff0,Aff1,Aff2,Aff3,Aff4 : string;
    DateDebutT : TdateTime;
begin

  TOBL.InitValeurs;
  AddLesSupLignesConso (TOBL);
  WithMajPrix := true;

  LLabelHeure.Caption := '';
  LLabelEtat.Caption  := '';
  LLabelTiers.Caption := '';

  TOBL.PutValue ('NEW','X');
  if TypeSaisie = 'INTERVENTION' then
  begin
    TOBL.SetString('BCO_FAMILLETAXE1','TN');
  end else
  begin
    TOBL.SetString('BCO_FAMILLETAXE1','');
  end;

  if TheMode = TgsMO then
  begin
    TOBL.PutValue('BCO_NATUREMOUV','MO');
    TOBL.PutValue ('BCO_DATEMOUV',LastDayMo);
  end else if TheMode = TgsFRS then
  begin
    TOBL.PutValue('BCO_NATUREMOUV','FRS');
    TOBL.PutValue ('BCO_DATEMOUV',LastDayFrs);
  end else if TheMode = TgsMOEXT then
  begin
    TOBL.PutValue('BCO_NATUREMOUV','EXT');
    TOBL.PutValue ('BCO_DATEMOUV',LastDayMoExt);
  end else if TheMode = TgsRES then
  begin
    TOBL.PutValue('BCO_NATUREMOUV','RES');
    TOBL.PutValue ('BCO_DATEMOUV',LastDayMat);
  end else if TheMode = TgsFOURN then
  begin
    TOBL.PutValue('BCO_NATUREMOUV','FOU');
    TOBL.PutValue ('BCO_DATEMOUV',LastDayFourn);
    //FV1 : 10/12/2014 - FS#1342 - SCETEC - saisie de consos matériaux
    TOBL.PutValue('BCO_LIBELLE',  LMATERIAUX.Caption);
  end else if TheMode = TgsRecettes then
  begin
    TOBL.PutValue('BCO_NATUREMOUV','RAN');
    TOBL.PutValue ('BCO_DATEMOUV',LastDayRecettes);
    TOBL.PutValue('LIBELLENATURE',RechDom ('BTNATUREMOUV',TOBL.GetValue('BCO_NATUREMOUV'),false));
  end;

  if (CHOIX_CHANTIER.Checked) OR (CHOIX_CONTRAT.Checked) OR (CHOIX_APPEL.Checked) then
    begin
    // Suite a erreur d'alimentation du BCo_AFFAIRE0 on revoit la copie
    (*
    TOBL.PutValue ('BCO_AFFAIRE',CH_CHANTIER.Text);
    TOBL.PutValue ('BCO_AFFAIRE0',CH_CHANTIER0.Text);
    TOBL.PutValue ('BCO_AFFAIRE1',CH_CHANTIER1.Text);
    TOBL.PutValue ('BCO_AFFAIRE2',CH_CHANTIER2.Text);
    TOBL.PutValue ('BCO_AFFAIRE3',CH_CHANTIER3.Text);
    TOBL.PutValue ('BCO_PHASETRA',PHASE.Text);
    *)
    TOBL.PutValue ('BCO_AFFAIRE',CH_CHANTIER.Text);
    BTPCodeAffaireDecoupe (TOBL.GetValue('BCO_AFFAIRE'),Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);
    CH_CHANTIER0.Text := Aff0; // grmfff
    TOBL.PutValue ('BCO_AFFAIRE0',Aff0);
    TOBL.PutValue ('BCO_AFFAIRE1',Aff1);
    TOBL.PutValue ('BCO_AFFAIRE2',Aff2);
    TOBL.PutValue ('BCO_AFFAIRE3',Aff3);
    //FV1 : 28/08/2013 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
    //TOBL.SetString('DOMAINE', CodeDomaine); (???)
    TOBL.PutValue('DOMAINE', CodeDomaine); 
    //
    TOBL.PutValue ('BCO_PHASETRA',PHASE.Text);
    if Choix_APPEL.Checked then
       Begin
         ChargeResource(CODERESSOURCE.text);
         if (TheMode = TgsMo) and (TOBRessource.Getvalue('ARS_TYPERESSOURCE') = 'SAL') then
         begin
           TOBL.putValue ('BCO_RESSOURCE', trim(CODERESSOURCE.text));
           TOBL.putValue ('LIBELLEMO', GetMainDoeuvre(CODERESSOURCE.text));
           SetInfoRessource (TOBL,TOBRESSOURCE,WithMajPrix,True,PrestationDefaut);
         end;
         if (TheMode = TgsMOExt) and (TOBRessource.Getvalue('ARS_TYPERESSOURCE') = 'ST') then
         begin
           TOBL.putValue ('BCO_RESSOURCE', trim(CODERESSOURCE.text));
           TOBL.putValue ('LIBELLEMO', GetMainDoeuvre(CODERESSOURCE.text));
           SetInfoRessource (TOBL,TOBRESSOURCE,WithMajPrix,True,PrestationDefaut);
         end;
         If (TOBL.GetString('BCO_AFFAIRE0') = 'W') then
         begin
           AppliqueCoeff(TOBL);
         end;
	     end;
    end
  else if CHOIX_MO.Checked then
  begin
    if (TheMode = TgsMO) Or (TheMode = TgsFrs) Or (theMode = Tgsfourn)  then
    begin
      ChargeResource(MAINDOEUVRE.Text);
      TOBL.PutValue ('BCO_RESSOURCE',Trim(MAINDOEUVRE.Text));
      if (TheMode <> TgsFRS) and (TheMode <> TgsFourn) then SetInfoRessource (TOBL,TOBRESSOURCE,WithMajPrix,True,PrestationDefaut);
    end;
  end
  else if CHOIX_RESSOURCE.Checked then
  begin
    if TheMode = TgsRes then
    begin
      ChargeResource(RESSOURCES.Text);
      TOBL.PutValue ('BCO_RESSOURCE',Trim(RESSOURCES.Text));
      SetInfoRessource (TOBL,TOBRESSOURCE,WithMajPrix,True,PrestationDefaut);
    end;
  end
  else if CHOIX_MOEXT.Checked then
  begin
    if TheMode = TgsMOEXT then
    begin
      ChargeResource(RESSOURCESEXT.Text);
      TOBL.PutValue ('BCO_RESSOURCE',Trim(RESSOURCESEXT.Text));
      SetInfoRessource (TOBL,TOBRESSOURCE,WithMajPrix,True,PrestationDefaut);
    end;
  end
  else if CHOIX_MATERIAUX.Checked then
  begin
    ARTICLE.Text := Trim(copy(MATERIAUX.Text,1,18));
    TOBL.PutValue ('BCO_CODEARTICLE',ARTICLE.Text);
    TOBL.PutValue ('BCO_ARTICLE',MATERIAUX.Text );
    //DefiniInfoArticle (TOBL,TOBMateriau,True);
    calculeLaLigne (TOBL);
  end;

  // Par défaut positionné facturable à non
  TOBL.PutValue ('BCO_FACTURABLE','N');
  TOBL.putValue ('BCO_AFFAIRESAISIE','');
  TOBL.putValue ('BCO_QTEFACTUREE',1);
  QTEFACTURE.Text := TOBL.GetValue('BCO_QTEFACTUREE');

  //chargement par défaut de la ressource si affaire = appel
  if Choix_APPEL.Checked then
    Begin
    GetLibelleChantier(TOBL);
    //TOBL.putValue ('BCO_AFFAIRESAISIE', CONTRAT.Text);
    TOBL.putValue ('BCO_QTEFACTUREE',1);
    if Contrat.Text = '' then
			 TOBL.PutValue ('BCO_FACTURABLE','A')
    else
	     TOBL.PutValue ('BCO_FACTURABLE','N');
    IF ((TheMode = TgsMO) or (TheMode = TgsFrs)) and (TOBRessource.Getvalue('ARS_TYPERESSOURCE') = 'SAL')  Then
       Begin
       TOBL.putValue ('BCO_RESSOURCE', Trim(CODERESSOURCE.text));
       TOBL.putValue ('LIBELLEMO', GetMainDoeuvre(CODERESSOURCE.text));
       end;
    end;

  TOBL.PutValue('BCO_DATETRAVAUX',ConstitueDateDebutTravaux(TOBL));
  DecodeDate (TOBL.GetValue('BCO_DATEMOUV'),year,month,day);
  
  TOBL.putValue ('JOUR',IntToStr (day));
  TOBL.putValue('BCO_SEMAINE',NumSemaine(TOBL.GetValue('BCO_DATEMOUV')));
  TOBL.PutValue ('BCO_MOIS',month);

  TOBL.putValue ('BCO_QUANTITE',1);
  if TypeSaisie = 'INTERVENTION' then
    TOBL.SetString('BCO_FAMILLETAXE1',GetFamilleTaxe(TOBL.GetString('BCO_ARTICLE')));

  // Modif Pour éviter qu'une ligne n'arrive complètement renseignée lors de sa création
  // Correction pouchain
	If (TypeSaisie = 'INTERVENTION') and (TheMode=TgsMo) then
  begin
  	TOBL.putValue ('BCO_QUANTITE',0);
    TOBL.putValue ('BCO_QTEFACTUREE',0);
    calculeLaLigne (TOBL);
  end;
  //
end;

function TOF_BTSAISIECONSO.AddNewLigne (LaTOB: TOB; Themode : TGrilleModeSaisie) : TOB;
var TOBL : TOB;
begin
  TOBL := TOB.Create ('CONSOMMATIONS',laTob,-1);
  initNewconso (TOBL,TheMode);
  result := TOBL;
end;

procedure TOF_BTSAISIECONSO.CreeTobs;
begin
  TOBMo := TOB.Create ('LES MO',nil,-1); // Main d'oeuvre interne
  TOBMoEXT := TOB.Create ('LES MO EXTERNES',nil,-1); // Main d'oeuvre Externes
  TOBFRS := TOB.Create ('LES PRIMES',nil,-1); // FRAIS - Primes
  TOBMAT := TOB.Create ('LES MATERIAUX',nil,-1); // Fournitures
  TOBRES := TOB.Create ('LES RESSOURCES',nil,-1); // ressources materielle ou prestation externe
  TOBCONSO := TOB.Create ('LES CONSO', nil,-1);
  TOBCONSOOLD := TOB.Create ('LES CONSO', nil,-1);
  TOBRecettes := TOB.Create ('LES RECETTES',nil,-1); // recettes et frais annexes
  TOBRessource := TOB.Create ('RESSOURCE',nil,-1);
  TOBTYPEHEURE := TOB.Create ('LES TYPES ET VALO',nil,-1);
  TOBMateriau := TOB.create ('ARTICLE',nil,-1);
  // Géré lors pour la saisie sur autres chose que chantiers ou appels
  TOBLIENEQUIPE := TOB.Create ('LIENS EQUIPES EXT',nil,-1);
  TOBLIENEQUIPE_O := TOB.Create ('LIENS EQUIPES OLD',nil,-1);
end;

procedure TOF_BTSAISIECONSO.LibereTObs;
begin
  TOBMo.free;
  TOBMoEXT.free;
  TOBFRS.free;
  TOBMAT.free;
  TOBRES.free;
  TOBREcettes.free;
  TOBCONSO.free;
  TOBCONSOOLD.free;
  TOBREssource.free;
  TOBTYPEHEURE.free;
  TOBMateriau.free;
  TOBLIENEQUIPE.free;
  TOBLIENEQUIPE_O.free;
end;

procedure TOF_BTSAISIECONSO.EffaceAffaireClick (Sender : TObject);
begin
  CH_CHANTIER.Text  := '';
  CH_CHANTIER1.Text := '';
  CH_CHANTIER2.Text := '';
  CH_CHANTIER3.Text := '';
  CH_AVENANT.Text   := '';
  PHASE.Text        := '';
end;

procedure TOF_BTSAISIECONSO.GetComponents;
begin

  PGTemps   := TPageControl (GetControl('PGTEMPS'));
  PHAUT     := THpanel (GetControl('PHAUT'));
  PSAISIE   := THpanel (GetControl('PSAISIE'));
  PCALENDAR := THpanel (GetControl('PCALENDAR'));
  //
  PANELCONTRAT	:= THPanel (GetControl('PANELCONTRAT'));
  PANELAPPEL    := THPanel (GetControl('PANELAPPEL'));
  //
  LCHANTIER         := THLabel(GetControl('LCHANTIER'));
  LCHANTIER.Caption := '';
  //
  P_CHANTIER  := THpanel (GetControl('P_CHANTIER'));
  P_RESSOURCE := THpanel (GetControl('P_RESSOURCE'));
  P_MO        := THpanel (GetControl('P_MO'));
  P_MATERIAUX := THpanel (GetControl('P_MATERIAUX'));
  P_MOEXT     := THpanel (GetControl('P_MOEXT'));
  //
  THEURES     := TTabSheet (GetControl('THEURES'));
  TFRAIS      := TTabSheet (GetControl('TFRAIS'));
  TMATERIELS  := TTabSheet (GetControl('TMATERIELS'));
  TFOURNITURES:= TTabSheet (GetControl('TFOURNITURES'));
  TMOEXT      := TTabSheet (GetControl('TMOEXT'));
  TRECETTES   := TTabSheet (GetControl('TRECETTES'));
  //
  Calendrier            := TmxCalendar.Create (PCALENDAR);
  Calendrier.Align      := alClient;
  calendrier.Parent     := PCALENDAR;
  Calendrier.DateFormat := 'dd/mm/yyyy';
  calendrier.Options    := calendrier.Options + [csSelectionEnabled,coClearButtonVisible];

  //
  CHOIX_CHANTIER  := TRadioButton (GetControl('CHOIX_CHANTIER'));
  CHOIX_CONTRAT   := TRadioButton (GetControl('CHOIX_CONTRAT'));
  CHOIX_APPEL     := TRadioButton (GetControl('CHOIX_APPEL'));
  CHOIX_MO        := TRadioButton (GetControl('CHOIX_MO'));
  CHOIX_RESSOURCE := TRadioButton (GetControl('CHOIX_RESSOURCE'));
  CHOIX_MATERIAUX := TRadioButton (GetControl('CHOIX_MATERIAUX'));
  CHOIX_MOEXT     := TRadioButton (GetControl('CHOIX_MOEXT'));
  //
  BEffaceAff1     := TToolbarButton97 (GetControl('BEffaceAff1'));
  BCherche        := TToolbarButton97 (GetControl('BCherche'));
  BChangeEtat     := TToolbarButton97 (GetControl('BChangeEtat'));
  BEffaceCha      := TToolbarButton97 (GetControl('BEFFACECHA'));
  BValider        := TToolbarButton97 (GetControl('BValider'));
  Bannul          := TToolbarButton97 (GetControl('BANNUL'));
  Bannul.onclick  := BannulClick;
  //LLIBELLE := THLabel(GetControl('LLIBELLE'));
  //
  FONCTION := THValComboBox(GetControl('FONCTION'));

  // Les grilles de saisie
  GHeures     := THGrid(GetCOntrol('GHEURES'));
  GFrais      := THGrid(GetCOntrol('GFRAIS'));
  GMAteriel   := THGrid(GetCOntrol('GMATERIEL'));
  GFourniture := THGrid(GetCOntrol('GFOURNITURE'));
  GMOExt      := THGrid(GetCOntrol('GMOEXT'));
  GRecettes   := THGrid(GetCOntrol('GRECETTES'));

  // Saisie par ressource
  TYPERESSOURCE := THValComboBox (GetCOntrol('TYPERESSOURCE'));
  RESSOURCES    := THEdit (GetControl('RESSOURCES'));
  LRESSOURCE    := THLabel (GetControl('LRESSOURCE'));
  LRESSOURCE.caption := '';

  // Saisie par matériaux
  MATERIAUX   := THEdit( getCOntrol('MATERIAUX'));
  ARTICLE     := THEdit( getCOntrol('ARTICLE'));

  MATERIAUX.OnElipsisClick  := MATERIAUXRECH;
  LMATERIAUX  := THLabel( getCOntrol('LMATERIAUX'));
  LMATERIAUX.Caption := '';

  // saisie sur MO
  POPMAINDOEUVRE := TPopupMenu (GetControl ('POPMAINDOEUVRE'));
  MoToutes := TMenuItem(GetControl('MnMOToutes'));
  // MOToutes.OnClick := MAINDOEUVRETOUTE ;
  MAINDOEUVRE := THEdit (getControl('MAINDOEUVRE'));
  LMAINDOEUVRE := THLabel (getControl('LMAINDOEUVRE'));
  LMAINDOEUVRE.Caption := '';

  // Main d'oeuvre externe
  TYPERESSOURCEEXT := THValComboBox (GetCOntrol('TYPERESSOURCEEXT'));
  RESSOURCESEXT := THEdit (GetControl('RESSOURCESEXT'));
  LRESSOURCEEXT := THLabel (GetControl('LRESSOURCEEXT'));
  LRESSOURCEEXT.caption := '';

  // saisie par chantier
  BSelectAff1   := TToolBarButton97 (GetControl('BSELECTAFF1'));
  CH_CHANTIER   := THEdit(GetControl('CH_CHANTIER'));
  CH_CHANTIER0  := THEdit(GetControl('CH_CHANTIER0'));
  CH_CHANTIER1  := THEdit(GetControl('CH_CHANTIER1'));
  CH_CHANTIER2  := THEdit(GetControl('CH_CHANTIER2'));
  CH_CHANTIER3  := THEdit(GetControl('CH_CHANTIER3'));
  CH_AVENANT    := THEdit(GetControl('CH_AVENANT'));

  BRechPhase := TToolbarButton97 (GetControl('BRECHPHASE'));
  LIBPHASE := THLabel (GetCOntrol('LIBPHASE'));
  PHASE := THEdit (GetControl('PHASE'));
  //
  PHEURES := THPanel(GetCOntrol('PHEURES'));
  //
  PWinPlus := THPanel (GetCOntrol('TOOLWINPLUS'));
  PlusWindow := TToolWindow97.Create (ecran);
  PlusWindow.Parent := ecran;
  PlusWindow.Top := PwinPlus.top ;
  PlusWindow.left := PWinPlus.left;
  PlusWindow.ClientHeight := 162 ;
  PlusWindow.ClientWidth := 379;
  PlusWindow.Caption := TraduireMemoire('Valorisation');
  PlusWindow.Resizable := false;
  PlusWindow.Visible := false;
  PlusWindow.OnClose := PlusWindowClose;
  //
  PwinPlus.Parent := PlusWindow;
  PwinPlus.Align := alClient;
  PwinPlus.visible := true;
  //
  PDESCRIPTIF:= THPanel (GetCOntrol('TWDESCRIPTIF'));
  RetourWindow := TToolWindow97.Create (ecran);
  RetourWindow.Parent := ecran;
  RetourWindow.Top := PDESCRIPTIF.top;
  RetourWindow.left := PDESCRIPTIF.left;
  RetourWindow.ClientHeight := 162 ;
  RetourWindow.ClientWidth := 379;
  RetourWindow.Caption := TraduireMemoire('Saisie détail Intervention');
  RetourWindow.Resizable := false;
  RetourWindow.Visible := false;
  RetourWindow.OnClose := RetourWindowClose;
  //
  PDESCRIPTIF.Parent := RetourWindow;
  PDESCRIPTIF.Align := alClient;
  PDESCRIPTIF.visible := true;
  //
  BBLocNote  := TToolbarButton97 (getControl('BBLOCNOTE'));
  BBLocNote1 := TToolbarButton97 (getControl('BBLOCNOTE1'));
  BBLocNote2 := TToolbarButton97 (getControl('BBLOCNOTE2'));

  THRSSAISIE := THLabel (GetCOntrol('THRSSAISIE'));
  NBHRSSAISIE := THNumEdit(GetControl('NBHRSSAISIE'));

  //
  QTE := THedit(GetControl('QTE'));
  DPA := THedit(GetControl('DPA'));
  DPR := THedit(GetControl('DPR'));
  DPV := THedit(GetControl('DPV'));
  MONTANTPA := THedit(GetControl('MONTANTPA'));
  MONTANTPR := THedit(GetControl('MONTANTPR'));
  MONTANTPV := THedit(GetControl('MONTANTPV'));
  //
  Bdelete := TToolbarButton97 (GetControl('BDELETE'));
  BDuplic := TToolbarButton97 (GetControl('BDUPLIC'));

  If Assigned(Getcontrol('BDuplicSemaine')) then BDuplicSemaine := TToolbarButton97 (GetControl('BDUPLICSEMAINE'));

  //PHeures
  LPHASE        := THLabel (GetCOntrol('LPHASE_'));

  CODERESSOURCE := THEdit(GetControl('CODERESSOURCE'));
  CODETIERS     := THEdit(GetControl('CODETIERS'));

  LCONTRAT      := THLabel (GetCOntrol('LCONTRAT'));
  LLABELHEURE   := THLabel (GetCOntrol('LLABELHEURE'));
  LLABELETAT    := THLabel (GetCOntrol('LLABELETAT'));
  LLABELTIERS   := THLabel (GetCOntrol('LLABELTIERS'));

  LQTEFACTURE   := THLabel (GetCOntrol('LQTEFACTURE'));
  LLIBCONTRAT   := THLabel (GetCOntrol('LLIBCONTRAT'));
  QTEFACTURE    := THedit(GetControl('QTEFACTURE'));
  CONTRAT       := THedit(GetControl('CONTRAT'));
  FACTURABLE    := TCheckBox (GetCOntrol('FACTURABLE'));
  FAMILLETAXE1  := THValComboBox (GetControl('FAMILLETAXE1'));
end;

procedure TOF_BTSAISIECONSO.SetChoix;
begin
  CHOIX_CHANTIER.OnClick  := CHOIX_CHANTIERClick;
  CHOIX_CONTRAT.OnClick   := CHOIX_CONTRATClick;
  CHOIX_APPEL.OnClick     := CHOIX_APPELClick;
  CHOIX_MO.onClick        := CHOIX_MOClick;
  CHOIX_RESSOURCE.OnClick := CHOIX_RESSOURCEClick;
  CHOIX_MATERIAUX.onClick := CHOIX_MATERIAUXClick;
  CHOIX_MOEXT.onClick     := CHOIX_MOEXTClick;
end;

procedure TOF_BTSAISIECONSO.SetToolValue (TOBL : TOB);
begin
 	if (LigneDejaValide (CurrTOB)) or (LigneFromPieces (CurrTOB)) then
 		PlusWindow.enabled := False
  else
 		PlusWindow.enabled := True;
  //
  QTE.Text := strf00(TOBL.GetValue('BCO_QUANTITE'),V_PGI.okDecQ);
  //
  DPA.Text := strf00(TOBL.GetValue('BCO_DPA'),V_PGI.okDecP);
  DPR.Text := strf00(TOBL.GetValue('BCO_DPR'),V_PGI.okDecP);
  DPV.Text := strf00(TOBL.GetValue('BCO_PUHT'),V_PGI.okDecP);
  //
  MONTANTPA.Text := strf00(TOBL.GetValue('BCO_MONTANTACH'),V_PGI.okDecV);
  MONTANTPR.Text := strf00(TOBL.GetValue('BCO_MONTANTPR'),V_PGI.okDecV);
  MONTANTPV.Text := strf00(TOBL.GetValue('BCO_MONTANTHT'),V_PGI.okDecV);
  //
end;

procedure TOF_BTSAISIECONSO.SetToolValue(Mode : TGrilleModeSaisie ; Arow : integer);
var TOBL : TOB;
begin

  TOBL := nil;

  if Mode = TgsMo then
  begin
    if Arow <= TOBMO.detail.count then TOBL := TOBMO.detail[Arow-1];
  end else if Mode = TgsMoEXT then
  begin
    if Arow <= TOBMOEXT.detail.count then TOBL := TOBMOEXT.detail[Arow-1];
  end else if Mode = TgsFrs then
  begin
    if Arow <= TOBFRS.detail.count then  TOBL := TOBFRS.detail[Arow-1]
  end else if Mode = Tgsfourn then
  begin
    if Arow <= TOBMAT.detail.count then TOBL := TOBMAT.detail[Arow-1]
  end else if Mode = TgsRecettes then
  begin
    if Arow <= TOBRecettes.detail.count then TOBL := TOBRecettes.detail[Arow-1]
  end else if Mode = TgsRES then
  begin
    if Arow <= TOBRES.detail.count then TOBL := TOBRES.detail[Arow-1];
  end;
  if TOBL = nil then exit;
  //
  CurrTOB := TOBL;
  SetToolvalue (TOBL);
  //
end;

procedure TOF_BTSAISIECONSO.ShowToolWindow(Name : String; Actif : boolean);
begin
	if (GetActiveTabSheet('PGTEMPS').Name = 'TRECETTES') then actif := false;

  if Actif Then
     Begin
     if Name = 'BBLOCNOTE' then
        PlusWindow.Visible := true
     else if (Name = 'BBLOCNOTE1') OR (Name = 'BBLOCNOTE2') then
        RetourWindow.Visible := true;
     end
  else
     Begin
     if Name = 'BBLOCNOTE' then
        PlusWindow.Visible := false
     else if (Name = 'BBLOCNOTE1') OR (Name = 'BBLOCNOTE2') then
        RetourWindow.Visible := false;
     end;
     
end;

procedure TOF_BTSAISIECONSO.BBlocNoteClick (Sender : TOBject);
begin
  if not (Sender Is TToolbarButton97) then exit;

  ShowToolWindow(TToolBarButton97(Sender).Name, TToolBarButton97(Sender).Down);

end;

procedure TOF_BTSAISIECONSO.SetEvent;
begin

  SetChoix;

  BRechPhase.OnClick          := BrechPhaseClick;
  //
  CH_CHANTIER.OnChange        := ChangeChantier;
  CH_CHANTIER1.OnExit         := OnExitPartieAffaire;
  CH_CHANTIER2.OnExit         := OnExitPartieAffaire;
  CH_CHANTIER3.OnExit         := OnExitPartieAffaire;
  //
  BSelectAff1.OnClick         := OnClickSelectAff;
  //
  PHASE.OnChange              := PhaseChange;
  BEffaceAff1.onClick         := EffaceAffaireClick;
  BCherche.onclick            := BChercheClick;
  BChangeEtat.OnClick         := BChangeEtatClick;

	Ecran.OnKeyDown             := FormKeyDown;

  // Sur Main d'oeuvre
  FONCTION.OnChange           := FONCTIONChange;
  MAINDOEUVRE.OnExit          := MAINDOEUVREExit;
  MAINDOEUVRE.OnElipsisClick  := MAINDOEUVREElipsisClick;

  // prestation
  TYPERESSOURCE.OnChange      := TYPEPRESTATIONCHANGE;
  RESSOURCES.OnExit           := RESSOURCEExit;

  // Materiaux
  MATERIAUX.OnExit            := MATERIAUXExit;

  // MO externe
  TYPERESSOURCEEXT.OnChange   := TYPEPRESTATIONEXTCHANGE;
  RESSOURCESEXT.OnExit        := RESSOURCEEXTExit;
  //
  BBlocNote.OnClick           := BBlocNoteClick;
  BBlocNote1.OnClick          := BBlocNoteClick;
  BBlocNote2.OnClick          := BBlocNoteClick;
  //
  //BBlocNote2.OnClick := BBlocNoteClick;
  //BBlocNote3.OnClick := BBlocNoteClick;
  //BBlocNote4.OnClick := BBlocNoteClick;
  //BBlocNote5.OnClick := BBlocNoteClick;
  //
  PGTemps.OnChange            := ChangeOnglet;
  PGTemps.OnChanging          := AvantChangeOnglet;
  //
  DPA.OnExit                  := DPAExit;
  DPR.OnExit                  := DPRExit;
  DPV.OnExit                  := DPVExit;

  MONTANTPA.OnExit            := MONTANTPAExit;
  MONTANTPR.OnExit            := MONTANTPRExit;
  MONTANTPV.OnExit            := MONTANTPVExit;

  CONTRAT.OnElipsisClick      := CONTRATClick;

  QTEFACTURE.OnExit           := QTEFACTUREExit;

  //
  Bdelete.onclick             := BdeleteClick;
  BDuplic.onclick             := BDuplicClick;
  //
  If Assigned(GetControl('BDuplicSemaine')) then BDuplicSemaine.onclick := BDuplicSemaineClick;
  //
  TToolBarButton97(GetCOntrol('BDUPLICCONSO')).OnClick := BDUPLICCONSOClick;
  BEffaceCha.OnClick := BEFFACECHAClick;

end;

Procedure TOF_BTSAISIECONSO.OnClickSelectAff(Sender : Tobject);
Var Stchamps        : String;
    StArgument      : String;
    TheTypeaffaire  : String;
    Tmp             : String;
    Aff0            : String;
    Aff1            : string;
    Aff2            : string;
    Aff3            : string;
    Avenant         : string;
    CodeAff         : string;
begin

  //StArgument := 'ACTION=CONSULTATION';
  StArgument := 'ACTION=RECH';
  //Stargument := Stargument + ';NOCHANGETIERS';
  //Stargument := Stargument + ';ACTION=RECH';
  //Stargument := Stargument + ';NOFILTRE';
  //StArgument := Stargument + ';NOAFFETAT';

  if ((CH_CHANTIER0 <> nil) and (CH_CHANTIER0.text<>'')) or
     ((CH_CHANTIER1 <> nil) and (CH_CHANTIER1.text<>'')) or
     ((CH_CHANTIER2 <> nil) and (CH_CHANTIER2.text<>'')) or
     ((CH_CHANTIER3 <> nil) and (CH_CHANTIER3.text<>'')) then
  begin
    if CH_CHANTIER0.text <> '' then StChamps := StChamps + ';AFF_AFFAIRE0=' + CH_CHANTIER0.text;
    if CH_CHANTIER1.text <> '' then StChamps := StChamps + ';AFF_AFFAIRE1=' + CH_CHANTIER1.Text;
    if CH_CHANTIER2.text <> '' then StChamps := StChamps + ';AFF_AFFAIRE2=' + CH_CHANTIER2.Text;
    if CH_CHANTIER3.text <> '' then StChamps := StChamps + ';AFF_AFFAIRE3=' + CH_CHANTIER3.Text;
    if (CH_AVENANT <> nil) then
    begin
      if (CH_AVENANT.Text <> '00') or (CH_AVENANT.Text <> '') then StChamps := StChamps + ';AFF_AVENANT=' + CH_AVENANT.Text;
    end;
  end;

  if CHOIX_CHANTIER.Checked	then
  begin
     TheTypeAffaire := 'A';
     StChamps := 'AFF_AFFAIRE0=' + TheTypeAffaire;
     StChamps := stChamps + ' ;AFF_STATUTAFFAIRE=AFF';
     StChamps := stChamps + ' ;AFF_AVENANT=';
     StArgument := Stargument + ';NOCHANGESTATUT'; //';STATUT:AFF;NOCHANGESTATUT';
   	 tmp := AGLLanceFiche('BTP', 'BTAFFAIRE_MUL', StChamps, '', StArgument);
  end
  else if CHOIX_CONTRAT.Checked 	then
  begin
    TheTypeAffaire := 'I';
    StChamps := 'AFF_STATUTAFFAIRE=INT';
    StArgument := Stargument + ';STATUT:INT;NOCHANGESTATUT';
    tmp := AGLLanceFiche('BTP', 'BTCONTRAT_MUL', StChamps, '', StArgument);
  end
  else if CHOIX_APPEL.Checked  	then
  begin
    TheTypeAffaire := 'W';
    StArgument := 'ACTION=CONSULTATION';
    Stargument := Stargument + ';NOCHANGETIERS';
    Stargument := Stargument + ';ACTION=RECH';
    StChamps := 'AFF_STATUTAFFAIRE=APP';
    StArgument := Stargument + ';STATUT:APP;NOCHANGESTATUT';
    tmp := AGLLanceFiche('BTP', 'BTMULAPPELS', StChamps, '', StArgument);
  end;

  if tmp <> '' then
  begin
    CodeAff := ReadTokenSt(tmp);
    BTPCodeAffaireDecoupe(Codeaff,Aff0,Aff1,Aff2,Aff3,Avenant,taCreat, False);
    if CH_CHANTIER  <> Nil then CH_CHANTIER.Text  := CodeAff;
    if CH_CHANTIER0 <> Nil then CH_CHANTIER0.Text := Aff0;
    if CH_CHANTIER1 <> Nil then CH_CHANTIER1.Text := Aff1;
    if CH_CHANTIER2 <> Nil then CH_CHANTIER2.Text := Aff2;
    if CH_CHANTIER3 <> Nil then CH_CHANTIER3.Text := Aff3;
  end;

end;

procedure TOF_BTSAISIECONSO.DPAExit (Sender : TObject);
var LastprPv,lastPaPr : double;
begin

  if CurrTOB.GetValue ('BCO_DPA') > 0  then LastPaPr := CurrTOB.GetValue ('BCO_DPR')/CurrTOB.GetValue ('BCO_DPA')
  else LastPaPr := 0;

  if CurrTOB.GetValue ('BCO_DPR') > 0  then LastPrPV := CurrTOB.GetValue ('BCO_PUHT')/CurrTOB.GetValue ('BCO_DPR')
  else LastPrPv := 0;

  CurrTOB.PutValue ('BCO_DPA',valeur(DPA.Text));
  calculeLaLigne (CurrTob,LastPaPr,LastPrPv);

  if ActiveMode = TgsMo then CurrTOB.PutLigneGrid (GHeures,Gheures.row,false,false,stColListe) else
  if ActiveMode = TgsFrs then CurrTOB.PutLigneGrid (GFrais,GFrais.row,false,false,stColListe) else
  if ActiveMode = Tgsfourn then CurrTOB.PutLigneGrid (GFourniture ,GFourniture.row,false,false,stColListe) else
  if ActiveMode = TgsMOEXT then CurrTOB.PutLigneGrid (GMOEXT ,GMOEXT.row,false,false,stColListe) else
  if ActiveMode = TgsRecettes then CurrTOB.PutLigneGrid (GRecettes ,GRecettes.row,false,false,stColListe) else
  if ActiveMode = TgsRES then CurrTOB.PutLigneGrid (GMAteriel ,GMAteriel.row,false,false,stColListe);

  SetToolValue (CurrTob);

end;

procedure TOF_BTSAISIECONSO.DPRExit (Sender : TObject);
var LastprPv : double;
begin

  if CurrTOB.GetValue ('BCO_DPR') > 0  then LastPrPV := CurrTOB.GetValue ('BCO_PUHT')/CurrTOB.GetValue ('BCO_DPR')
  else LastPrPv := 0;

  CurrTOB.PutValue ('BCO_DPR',valeur(DPR.Text));

  calculeLaLigne (CurrTob,0,LastprPv);

  if ActiveMode = TgsMo then CurrTOB.PutLigneGrid (GHeures,Gheures.row,false,false,stColListe) else
  if ActiveMode = TgsFrs then CurrTOB.PutLigneGrid (GFrais,GFrais.row,false,false,stColListe) else
  if ActiveMode = Tgsfourn then CurrTOB.PutLigneGrid (GFourniture ,GFourniture.row,false,false,stColListe) else
  if ActiveMode = TgsMOEXT then CurrTOB.PutLigneGrid (GMOEXT ,GMOEXT.row,false,false,stColListe) else
  if ActiveMode = TgsRecettes then CurrTOB.PutLigneGrid (GRecettes ,GRecettes.row,false,false,stColListe) else
  if ActiveMode = TgsRES then CurrTOB.PutLigneGrid (GMAteriel ,GMAteriel.row,false,false,stColListe);

  SetToolValue (CurrTob);

end;

procedure TOF_BTSAISIECONSO.DPVExit (Sender : TObject);
begin
  CurrTOB.PutValue ('BCO_PUHT',valeur(DPV.Text));
  calculeLaLigne (CurrTob);

  if ActiveMode = TgsMo then CurrTOB.PutLigneGrid (GHeures,Gheures.row,false,false,stColListe) else
  if ActiveMode = TgsFrs then CurrTOB.PutLigneGrid (GFrais,GFrais.row,false,false,stColListe) else
  if ActiveMode = Tgsfourn then CurrTOB.PutLigneGrid (GFourniture ,GFourniture.row,false,false,stColListe) else
  if ActiveMode = TgsMOEXT then CurrTOB.PutLigneGrid (GMOEXT ,GMOEXT.row,false,false,stColListe) else
  if ActiveMode = TgsRecettes then CurrTOB.PutLigneGrid (GRecettes ,GRecettes.row,false,false,stColListe) else
  if ActiveMode = TgsRES then CurrTOB.PutLigneGrid (GMAteriel ,GMAteriel.row,false,false,stColListe);

  SetToolValue (CurrTob);
end;

procedure TOF_BTSAISIECONSO.MONTANTPAExit (Sender : TObject);
begin
  CurrTOB.PutValue ('BCO_MONTANTACH',valeur(MONTANTPA.Text));
  CurrTOB.putValue('BCO_DPA',Arrondi(Valeur(MONTANTPA.text)/CurrTOB.GetValue('BCO_QUANTITE'),V_PGI.OkdecP));
  calculeLaLigne (CurrTob);
  SetToolValue (CurrTob);
end;

procedure TOF_BTSAISIECONSO.MONTANTPRExit (Sender : TObject);
begin
  CurrTOB.PutValue ('BCO_MONTANTPR',valeur(MONTANTPR.Text));
  CurrTOB.putValue('BCO_DPR',Arrondi(Valeur(MONTANTPR.text)/CurrTOB.GetValue('BCO_QUANTITE'),V_PGI.OkdecP));
  calculeLaLigne (CurrTob);
  SetToolValue (CurrTob);
end;

procedure TOF_BTSAISIECONSO.MONTANTPVExit (Sender : TObject);
begin
  CurrTOB.PutValue ('BCO_MONTANTHT',valeur(MONTANTPV.Text));
  CurrTOB.putValue('BCO_PUHT',Arrondi(Valeur(MONTANTPV.text)/CurrTOB.GetValue('BCO_QUANTITE'),V_PGI.OkdecP));
  if ActiveMode = TgsMo then CurrTOB.PutLigneGrid (GHeures,Gheures.row,false,false,stColListe) else
  if ActiveMode = TgsMoExt then CurrTOB.PutLigneGrid (GMOext,GMOext.row,false,false,stColListe) else
  if ActiveMode = TgsFrs then CurrTOB.PutLigneGrid (GFrais,GFrais.row,false,false,stColListe) else
  if ActiveMode = Tgsfourn then CurrTOB.PutLigneGrid (GFourniture ,GFourniture.row,false,false,stColListe) else
  if ActiveMode = TgsRecettes then CurrTOB.PutLigneGrid (GRecettes ,GRecettes.row,false,false,stColListe) else
  if ActiveMode = TgsRES then CurrTOB.PutLigneGrid (GMAteriel ,GMAteriel.row,false,false,stColListe);
  SetToolValue (CurrTob);
end;

procedure TOF_BTSAISIECONSO.ChangeOnglet (Sender : TObject);
var Cancel : boolean;
begin
  //
  BChangeEtat.visible := False;

  if (GetActiveTabSheet('PGTEMPS').Name = 'THEURES') then
  	 begin
		 ControleChangeOnglet(GHeures, cancel);
     //if (CHOIX_MO.checked) Or (CHOIX_RESSOURCE.checked) Or CHOIX_MOEXT.checked then BChangeEtat.visible := True;
  	 end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFRAIS') then
     begin
		 ControleChangeOnglet(GFrais, cancel);
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMATERIELS') then
     begin
		 ControleChangeOnglet(GMateriel, cancel);
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFOURNITURES') then
     begin
     ControleChangeOnglet(GFourniture, cancel);
		 end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMOEXT') then
     begin
     ControleChangeOnglet(GMOEXT, cancel);
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TRECETTES') then
     begin
     ControleChangeOnglet(Grecettes, cancel);
     end;

  if Sender is TToolbarButton97 then ShowToolWindow(TToolBarButton97(Sender).Name, TToolBarButton97(Sender).Down);

  //
end;

procedure TOF_BTSAISIECONSO.AvantChangeOnglet (Sender : TObject; var AllowChange : Boolean);
begin
  // utile pour valider le champ en cours de saisie :cf FQ 15665
  if TFvierge(ecran).ActiveControl <> nil then
    SendMessage(TFvierge(ecran).ActiveControl.Handle , WM_KeyDown, VK_TAB, 0);
end;

function TOF_BTSAISIECONSO.constitueWhere : string;
var CodeAffaire : String;
begin

  CodeAffaire := CH_CHANTIER.text;

  if (CHOIX_CHANTIER.Checked) or (CHOIX_CONTRAT.Checked) Or (CHOIX_APPEL.Checked) then
     begin
     Result := 'BCO_AFFAIRE="' + CodeAffaire + '" ' + 'AND BCO_NATUREPIECEG NOT IN ("CF","BLF","BFA","FF","LFR","LCR")';
     if PHASE.text <> '' then
        begin
        result := result + ' AND BCO_PHASETRA="' + PHASE.text+'"';
        end;
     end
  else if CHOIX_MATERIAUX.checked Then
     begin
     Result := 'BCO_NATUREMOUV="FOU" ';
     result := result + 'AND BCO_ARTICLE="' + MATERIAUX.text + '" ';
     result := result + 'AND BCO_NATUREPIECEG NOT IN ("CF","BLF","BFA","FF","AF","LFR","LCR","AFS")';
     end
  else if CHOIX_MO.checked then
     begin
     Result := '(BCO_NATUREMOUV="MO" OR BCO_NATUREMOUV="FRS" OR BCO_NATUREMOUV="FOU")';
     result := result + ' AND BCO_RESSOURCE="'+Trim(MAINDOEUVRE.text)+'"';
     end
  else if CHOIX_RESSOURCE.checked then
     begin
     Result := 'BCO_NATUREMOUV="RES"';
     result := result + ' AND BCO_RESSOURCE="'+trim(RESSOURCES.text)+'"';
     end
  else if CHOIX_MOEXT.checked then
     begin
     Result := 'BCO_NATUREMOUV="EXT"';
     result := result + ' AND BCO_RESSOURCE="'+trim(RESSOURCESEXT.text)+'"';
     result := result + ' AND BCO_NATUREPIECEG NOT IN ("CF","BLF","BFA","FF","AF","LFR","LCR","AFS")';
     end;

  if Calendrier.SelectionEnd <> Calendrier.SelectionStart then
     begin
     result := result + ' AND BCO_DATEMOUV >= "'+USDATETIME (Calendrier.SelectionStart)+'"';
     result := result + ' AND BCO_DATEMOUV <= "'+USDATETIME (Calendrier.SelectionEnd)+'"';
     end
  else
     result := result + ' AND BCO_DATEMOUV = "'+USDATETIME (Calendrier.SelectionStart)+'"';

  result := result + ' AND BCO_TRANSFORME="-"'; // ne reprends pas les transformé
  result := result + ' AND BCO_TRAITEVENTE<>"X"'; // ne reprends pas les passés en livraisons
  result := result + ' ORDER BY BCO_DATEMOUV';

end;

procedure TOF_BTSAISIECONSO.GetConso;
var Req : String;
    QQ : TQuery;
    TOBL,TOBC : TOB;
    Indice,IndDate : integer;
    Nbhrs : double;
begin

    NBhrs := 0;
    req := 'SELECT * FROM CONSOMMATIONS WHERE '+ constitueWhere;
    QQ := OpenSql (Req,true,-1,'',true);

  TRY
    if not QQ.eof then
    begin
       TOBCONSO.LoadDetailDB ('CONSOMMATIONS','','',QQ,FALSE);
       TOBConsoOLd.dupliquer(TOBCOnso,true,true);
       //
       IndDate := TOBConso.detail[0].GetNumChamp ('BCO_DATEMOUV');
      for Indice := 0 to TOBConso.detail.count -1 do
      begin
        TOBC := TOBCONSO.detail[indice];
        //
        if not ((CHOIX_CHANTIER.Checked) or (CHOIX_CONTRAT.Checked) OR (CHOIX_APPEL.Checked)) then
        begin
          if TOBC.GetString('BCO_LINKEQUIPE')<>'' then
          begin
           RecupInfoEquipe(TOBC);
          end;
        end;
        // Frais - primes
        if TOBC.GetValue('BCO_NATUREMOUV')='FRS' then
        begin
          TOBL := TOB.Create ('CONSOMMATIONS',TOBFRS,-1);
          if TOBC.GetValeur (inddate) > LastDayFrs Then LastDayFrs := TOBC.GetValeur (inddate);
          TOBL.Dupliquer (TOBC,true,true);
          AddLesSupLignesConso (TOBL);
        end
        // Matériaux (fournitures)
        else if TOBC.GetValue('BCO_NATUREMOUV')='FOU' then
        begin
          TOBL := TOB.Create ('CONSOMMATIONS',TOBMAT,-1);
          if TOBC.GetValeur (inddate) > LastDayFourn Then LastDayFourn := TOBC.GetValeur (inddate);
          TOBL.Dupliquer (TOBC,true,true);
          AddLesSupLignesConso (TOBL);
        end
        // Main d'oeuvre interne
        else if TOBC.GetValue('BCO_NATUREMOUV')='MO' then
        begin
          TOBL := TOB.Create ('CONSOMMATIONS',TOBMO,-1);
          if TOBC.GetValeur (inddate) > LastDayMo Then LastDayMo := TOBC.GetValeur (inddate);
          TOBL.Dupliquer (TOBC,true,true);
          AddLesSupLignesConso (TOBL);
          NBHRS:= NBHRS + TOBL.GetValue('BCO_QUANTITE');
        end
        // Main d'oeuvre Externe
        else if TOBC.GetValue('BCO_NATUREMOUV')='EXT' then
        begin
          TOBL := TOB.Create ('CONSOMMATIONS',TOBMOEXT,-1);
          if TOBC.GetValeur (inddate) > LastDayMoExt  Then LastDayMoExt := TOBC.GetValeur (inddate);
          TOBL.Dupliquer (TOBC,true,true);
          AddLesSupLignesConso (TOBL);
        end
        // RESSOURCE (engins)
        else if TOBC.GetValue('BCO_NATUREMOUV')='RES' then
        begin
          TOBL := TOB.Create ('CONSOMMATIONS',TOBRES,-1);
          if TOBC.GetValeur (inddate) > LastDayMat  Then LastDayMat := TOBC.GetValeur (inddate);
          TOBL.Dupliquer (TOBC,true,true);
          AddLesSupLignesConso (TOBL);
        end
        // recettes annexes
        else if TOBC.GetValue('BCO_NATUREMOUV')='RAN' then
        begin
          TOBL := TOB.Create ('CONSOMMATIONS',TOBRecettes,-1);
          if TOBC.GetValeur (inddate) > LastDayrecettes  Then LastDayrecettes := TOBC.GetValeur (inddate);
          TOBL.Dupliquer (TOBC,true,true);
          TOBL.PutValue ('BCO_QUANTITE',TOBL.GetValue ('BCO_QUANTITE') * -1); // remet la quantite en positif
          calculeLaLigne (TOBL);
          AddLesSupLignesConso (TOBL);
        end
        // Frais Annexes
        else if TOBC.GetValue('BCO_NATUREMOUV')='FAN' then
        begin
          TOBL := TOB.Create ('CONSOMMATIONS',TOBRecettes,-1);
          if TOBC.GetValeur (inddate) > LastDayRecettes  Then LastDayrecettes := TOBC.GetValeur (inddate);
          TOBL.Dupliquer (TOBC,true,true);
          AddLesSupLignesConso (TOBL);
        end;        //
        PositionneValeurInit (TOBL);
        //
        GetLibelleChantier(TOBL);
        //
      end;
    end;
  FINALLY
    ferme (QQ);
    //
    TOBLIENEQUIPE_O.Dupliquer(TOBLIENEQUIPE,True,true);
    ///////
    if LastDayMo = 0        Then LastDayMo        := Calendrier.SelectionStart;
    if LastDayFrs = 0       Then LastDayFrs       := Calendrier.SelectionStart;
    if LastDayMat = 0       Then LastDayMat       := Calendrier.SelectionStart;
    if LastDayFourn = 0     Then LastDayFourn     := Calendrier.SelectionStart;
    if LastDayMoExt = 0     Then LastDayMoExt     := Calendrier.SelectionStart;
    if LastDayRecettes = 0  Then LastDayrecettes  := Calendrier.SelectionStart;
    //
    if (CHOIX_CHANTIER.Checked) OR (CHOIX_CONTRAT.Checked) Or (CHOIX_Appel.Checked) then
    begin
       WithPhases := IsExistPhases (CH_CHANTIER.text);
       if TOBMo.detail.count        = 0 Then TOBL := AddNewLigne (TOBMO,TgsMO);
       if TOBMoEXT.detail.count     = 0 Then TOBL := AddNewLigne (TOBMOEXT,TgsMoext);
       if TOBFRS.detail.count       = 0 Then TOBL := AddNewLigne (TOBFrs,TgsFRS);
       if TOBMAT.detail.count       = 0 Then TOBL := AddNewLigne (TOBMAT,TgsFOURN);
       if TOBRES.detail.count       = 0 Then TOBL := AddNewLigne (TOBRES,TgsRES);
       if TOBRecettes.detail.count  = 0 Then TOBL := AddNewLigne (TOBREcettes,TgsREcettes);
    end
    else if CHOIX_MATERIAUX.checked Then
    begin
      if TOBMAT.detail.count = 0       Then
      Begin
        TOBL := AddNewLigne (TOBMAT,TgsFOURN);
      end;
    end
    else if CHOIX_MO.checked then
    begin
      if TOBMo.detail.count = 0        Then TOBL := AddNewLigne (TOBMO,TgsMO);
      if TOBFRS.detail.count = 0       Then TOBL := AddNewLigne (TOBFrs,TgsFRS);
    end
    else if CHOIX_MOEXT.checked then
    begin
      if TOBMoEXT.detail.count = 0     Then TOBL := AddNewLigne (TOBMOEXT,TgsMOEXT);
    end
    else if CHOIX_RESSOURCE.checked then
    begin
      if TOBRES.detail.count = 0       Then TOBL := AddNewLigne (TOBRES,TgsRES);
    end;

    NBHRSSAISIE.Value := NBHRS;
//
  END;
end;

procedure TOF_BTSAISIECONSO.SetGridDefault;
begin

GHeures.Col := 1;
GHeures.Row := 1;
//
GFourniture.Col := 1;
GFourniture.Row := 1;
//
GFrais.Col := 1;
GFrais.Row := 1;
//
GMateriel.Col := 1;
GMateriel.Row := 1;
//
GMoExt.Col := 1;
GMoExt.Row := 1;
//
GRecettes.Col := 1;
GRecettes.Row := 1;

end;

procedure TOF_BTSAISIECONSO.BChercheClick (Sender : TObject);
var Arow,Acol : integer;
begin
  FAMILLETAXE1.OnChange := nil;

	EventGrilles(false);
  SetGridDefault;
  BBLocnote1.visible := False;

  BChangeEtat.Visible := False;

  if ControleSaisieOK then
  begin
     //if CHOIX_MO.checked then
     //   ChargeResource (MAINDOEUVRE.text)
     //else if CHOIX_RESSOURCE.checked then
     //   ChargeResource (RESSOURCES.text)
     //else if CHOIX_MOEXT.checked then
     //   ChargeResource (RESSOURCESEXT.text);

    LastDayMo := 0;
    LastDayFrs := 0;
    LastDayMat := 0;
    LastDayFourn := 0;
    LastDayMoExt := 0;
    LastDayRecettes := 0;

    PSAISIE.visible := true;
    
    TOBMo.ClearDetail;
    TOBFRS.ClearDetail;
    TOBMAT.ClearDetail;
    TOBRES.ClearDetail;
    TOBMOEXT.ClearDetail;
    TOBConso.ClearDetail;
    TOBConsoOld.ClearDetail;
    TOBRecettes.ClearDetail;
    TOBLIENEQUIPE.ClearDetail;
    TOBLIENEQUIPE_O.ClearDetail;

    GetConso;

    DefiniGrilles;
    Remplitgrilles;
    EventGrilles(true);

    if (CHOIX_CHANTIER.Checked) OR (CHOIX_CONTRAT.Checked) OR (CHOIX_APPEL.Checked) then
    begin
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GHeures,Arow,Acol);
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GFourniture,Arow,Acol);
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GFrais,Arow,Acol);
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GMateriel,Arow,Acol);
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GMoExt,Arow,Acol);
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GRecettes,Arow,Acol);
      //
      ActiveMode := TgsMO;
      PGTemps.ActivePage := THeures;
      ChangeOnglet (Self);
    end
    else if CHOIX_MATERIAUX.checked Then
    begin
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GFourniture,Arow,Acol);
      ActiveMode := TgsFOURN;
      PGTemps.ActivePage := TFOURNITURES;
      ChangeOnglet (Self);
    end
    else if CHOIX_MO.checked then
    begin
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GHeures,Arow,Acol);
    //
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GFrais,Arow,Acol);
    //
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GFourniture,Arow,Acol);
      //
      ActiveMode := TgsMo;
      PGTemps.ActivePage := THEURES;
      // -- MODIF LS Pour visu total heures saisie
      THRSSAISIE.Visible := true;
      NBHRSSAISIE.Visible := true;
      //
      ChargeResource (MAINDOEUVRE.text);
      ChangeOnglet (Self);
      //FV1 - 10/07/2017 :  FS#2614 - VEODIS : En saisie Intervenant passage de REA=>ECR et ECR=>REA
      BChangeEtat.Visible := True;
    end
    else if CHOIX_RESSOURCE.checked then
    begin
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GMateriel,Arow,Acol);
      ActiveMode := TgsRES;
      PGTemps.ActivePage := TMATERIELS;
      ChargeResource (RESSOURCES.text);
      ChangeOnglet (Self);
      //FV1 - 10/07/2017 :  FS#2614 - VEODIS : En saisie Intervenant passage de REA=>ECR et ECR=>REA
      BChangeEtat.Visible := True;
    end
    else if CHOIX_MOEXT.checked then
    begin
      Arow := 1;
      Acol := 1;
      PositionneDansGrid (GMoExt,Arow,Acol);
      ActiveMode := TgsMoext;
      PGTemps.ActivePage := TMOEXT;
      ChargeResource (RESSOURCESEXT.text);
      ChangeOnglet (Self);
      //FV1 - 10/07/2017 :  FS#2614 - VEODIS : En saisie Intervenant passage de REA=>ECR et ECR=>REA
      BChangeEtat.Visible := True;
    end;
    EnteteActive(False);
  end;

  FAMILLETAXE1.OnChange := FAMILLETAXEChange;

end;

procedure TOF_BTSAISIECONSO.CHOIX_CHANTIERClick(Sender : TObject);
begin

  //Déselection des autre choix
  CHOIX_CHANTIER.Checked := true;
  CHOIX_MO.Checked := False;
  CHOIX_RESSOURCE.Checked := False;
  CHOIX_MATERIAUX.Checked := False;
  CHOIX_MOEXT.Checked := False;
  CHOIX_CONTRAT.Checked := False;
  CHOIX_APPEL.Checked := False;

  //
  CH_CHANTIER.text := 'A';
  CH_CHANTIER0.text := 'A';
  //
  setcontroltext('TCHANTIER', 'Chantier : ');
  //
  TheStatus := 'AFF';
  TheEtatAffaire := '';
  //
  GestionCheckedClick(false);
  //
  If Assigned(Getcontrol('BDuplicSemaine')) then BDUPLICSEMAINE.visible := false;
  //
end;

procedure TOF_BTSAISIECONSO.GestionCheckedClick(ok_aff : boolean);
Begin
  //
  LLabelHeure.Caption := '';
  LLabelEtat.Caption  := '';
  LLabelTiers.Caption := '';
  //
  BchangeEtat.visible := False;
  //
  if TheStatus = 'AFF' then
  Begin
    BRECHPHASE.Visible := True;
    LIBPHASE.Visible := True;
    LPHASE.Visible := True;
  End
  Else
  Begin
    BRECHPHASE.Visible := False;
    LIBPHASE.Visible := False;
    LPHASE.Visible := False;
  End;
  //
  P_RESSOURCE.visible := false;
  P_MATERIAUX.visible := false;
  P_MO.visible := false;
  P_MOEXT.visible := false;
  P_CHANTIER.visible := true;
  //
  TMATERIELS.tabvisible := true;
  TFOURNITURES.tabvisible := true;
  THEURES.tabvisible := true;
  TFRAIS.tabvisible := true;
  TMOEXT.tabvisible := true;
  TRECETTES.tabvisible := true;

  //Formatage du code affaire pour appel - ChargeClefAppel
  ChargeCleAffaire(CH_CHANTIER0, CH_CHANTIER1, CH_CHANTIER2, CH_CHANTIER3, CH_AVENANT, BSelectAFF1, TaModif, CH_CHANTIER.Text, false);

  //Formatage de l'écran dans le cas de la saisie d'un appel
  GestionAffichageContrat(ok_aff);

end;

procedure TOF_BTSAISIECONSO.CHOIX_CONTRATClick(Sender : TObject);
begin

  //Déselection des autre choix
  CHOIX_CHANTIER.Checked := False;
  CHOIX_MO.Checked := False;
  CHOIX_RESSOURCE.Checked := False;
  CHOIX_MATERIAUX.Checked := False;
  CHOIX_MOEXT.Checked := False;
  CHOIX_CONTRAT.Checked := True;
  CHOIX_APPEL.Checked := False;
  //
  CH_CHANTIER.text := 'I';
  CH_CHANTIER0.text := 'A';

  setcontroltext('TCHANTIER', 'Contrat : ');

  //ChargeCleAffaire(CH_CHANTIER0, CH_CHANTIER1, CH_CHANTIER2, CH_CHANTIER3, CH_AVENANT, BSelectAFF1, TaModif, CH_CHANTIER.Text, false);

  TheStatus := 'INT';
  TheEtatAffaire := '';

  GestionCheckedClick(False);

end;

procedure TOF_BTSAISIECONSO.CHOIX_APPELClick(Sender : TObject);
begin

  //Déselection des autre choix
  CHOIX_CHANTIER.Checked := False;
  CHOIX_MO.Checked := False;
  CHOIX_RESSOURCE.Checked := False;
  CHOIX_MATERIAUX.Checked := False;
  CHOIX_MOEXT.Checked := False;
  CHOIX_CONTRAT.Checked := False;
  CHOIX_APPEL.Checked := True;
  //
  CH_CHANTIER.text := 'W';
  CH_CHANTIER0.text := 'W';

  If Assigned(Getcontrol('BDuplicSemaine')) then BDuplicSemaine.visible := False;

  setcontroltext('TCHANTIER', 'Appel : ');

  //Formatage du code affaire pour appel - ChargeClefAppel
  //ChargeCleAffaire(CH_CHANTIER0, CH_CHANTIER1, CH_CHANTIER2, CH_CHANTIER3, CH_AVENANT, BSelectAFF1, TaModif, CH_CHANTIER.Text, false);

  TheStatus := 'APP';
  TheEtatAffaire := '';

  GestionCheckedClick(True);

end;


procedure TOF_BTSAISIECONSO.CHOIX_RESSOURCEClick (Sender : TObject);
begin
  //
  //Déselection des autre choix
  CHOIX_CHANTIER.Checked := False;
  CHOIX_MO.Checked := False;
  CHOIX_RESSOURCE.Checked := True;
  CHOIX_MATERIAUX.Checked := False;
  CHOIX_MOEXT.Checked := False;
  CHOIX_CONTRAT.Checked := False;
  CHOIX_APPEL.Checked := False;
  //
  BRECHPHASE.Visible := False;
  LIBPHASE.Visible := False;
  LPHASE.Visible := False;
  //
  P_CHANTIER.visible := false;
  P_MATERIAUX.visible := false;
  P_MO.visible := false;
  P_RESSOURCE.visible := true;
  P_MOEXT.visible := false;
  //
  TMATERIELS.tabvisible := True;
  TMOEXT.tabvisible := false;
  THEURES.tabvisible := false;
  TFRAIS.tabvisible := false;
  TFOURNITURES.tabvisible := false;
  TRECETTES.tabvisible := false;

  If Assigned(Getcontrol('BDuplicSemaine')) then BDuplicSemaine.visible := True;

 	//Formatage de l'écran dans le cas de la saisie d'un appel
  GestionAffichageContrat(false);

end;

procedure TOF_BTSAISIECONSO.CHOIX_MOClick(Sender : TOBject);
begin
  //Déselection des autre choix
  CHOIX_CHANTIER.Checked  := False;
  CHOIX_MO.Checked        := True;
  CHOIX_RESSOURCE.Checked := False;
  CHOIX_MATERIAUX.Checked := False;
  CHOIX_MOEXT.Checked     := False;
  CHOIX_CONTRAT.Checked   := False;
  CHOIX_APPEL.Checked     := False;
  //
  P_CHANTIER.visible      := false;
  P_MATERIAUX.visible     := false;
  P_RESSOURCE.visible     := false;
  P_MO.visible            := true;
  P_MOEXT.visible         := false;
  //
  TMATERIELS.tabvisible   := false;
  TMOEXT.tabvisible       := false;
  THEURES.tabvisible      := true;
  TFRAIS.tabvisible       := true;
  TFOURNITURES.tabvisible := True;
  TRECETTES.tabvisible    := false;
  //
  BDuplic.Visible         := False;
  //
  If Assigned(Getcontrol('BDuplicSemaine')) then BDuplicSemaine.visible := True;

  //A controler surement un bug !!!!!
  CODETIERS.Text := '';
  CODERESSOURCE.text := '';

  CONTRAT.text := '';

  LLIBCONTRAT.caption := '';

  MAINDOEUVRE.Text := '';
  LMAINDOEUVRE.caption := '';

  GestionAffichageContrat(false);

end;


procedure TOF_BTSAISIECONSO.CHOIX_MOEXTClick(Sender : TOBject);
begin
	//
  //Déselection des autre choix
  CHOIX_CHANTIER.Checked := False;
  CHOIX_MO.Checked := False;
  CHOIX_RESSOURCE.Checked := False;
  CHOIX_MATERIAUX.Checked := False;
  CHOIX_MOEXT.Checked := True;
  CHOIX_CONTRAT.Checked := False;
  CHOIX_APPEL.Checked := False;
  //
  P_CHANTIER.visible := false;
  P_MATERIAUX.visible := false;
  P_RESSOURCE.visible := false;
  P_MO.visible := false;
  P_MOEXT.visible := True;
  //
  TMATERIELS.tabvisible := false;
  THEURES.tabvisible := false;
  TFRAIS.tabvisible := false;
  TFOURNITURES.tabvisible := false;
  TMOEXT.tabvisible := true;
  TRECETTES.tabvisible := false;
  //
  If Assigned(Getcontrol('BDuplicSemaine')) then BDuplicSemaine.visible := False;
  //
  GestionAffichageContrat(false);
  //
end;

procedure TOF_BTSAISIECONSO.CHOIX_MATERIAUXClick(Sender : TObject);
begin
  //Déselection des autre choix
  CHOIX_CHANTIER.Checked := False;
  CHOIX_MO.Checked := False;
  CHOIX_RESSOURCE.Checked := False;
  CHOIX_MATERIAUX.Checked := True;
  CHOIX_MOEXT.Checked := False;
  CHOIX_CONTRAT.Checked := False;
  CHOIX_APPEL.Checked := False;
  //
  P_CHANTIER.visible := false;
  P_RESSOURCE.visible := false;
  P_MO.visible := false;
  P_MOEXT.visible := false;
  P_MATERIAUX.visible := true;
  //
  TMATERIELS.tabvisible :=false;
  THEURES.tabvisible := false;
  TFRAIS.tabvisible := false;
  TMOEXT.tabvisible := false;
  TFOURNITURES.tabvisible := true;
  TRECETTES.tabvisible := false;
  //
  If Assigned(Getcontrol('BDuplicSemaine')) then BDuplicSemaine.visible := False;
  //
  GestionAffichageContrat(false);
  //
end;

procedure TOF_BTSAISIECONSO.BrechPhaseClick(Sender: TObject);
var Code : string;

begin
  if CH_CHANTIER.Text = '' then exit;

  if SelectionPhase (Ch_CHANTIER.text,Code) then
  	 PHASE.Text := Code;

end;

procedure TOF_BTSAISIECONSO.ChangeChantier(Sender: TObject);
var QQ     : TQuery;
    Sql    : String;
    Statut : string;
    Etat   : String;
    Appel  : String;
    //
    LibEtat  : string;
    LibTiers : String;
begin

  CODETIERS.Text := '';
  CODERESSOURCE.text := '';
  BBlocNote1.Visible  := False;
  //
  LCHANTIER.Caption   := '';
  LLabelHeure.Caption := '';
  LLabelEtat.Caption  := '';
  LLabelTiers.Caption := '';

  if Ch_CHANTIER.text = '' then exit;

  Ch_CHANTIER.text := Trim(UpperCase (Ch_CHANTIER.text));


  if CHOIX_APPEL.checked then
     Sql := 'SELECT AFF_AFFAIRE, AFF_LIBELLE, AFF_TIERS, AFF_RESPONSABLE, AFF_AFFAIREINIT, AFF_STATUTAFFAIRE, AFF_ETATAFFAIRE, AFF_DOMAINE, T_LIBELLE '
  else
     Sql := 'SELECT AFF_LIBELLE, AFF_TIERS, AFF_DOMAINE, AFF_ETATAFFAIRE, T_LIBELLE ';

  Sql := Sql + ' FROM AFFAIRE LEFT JOIN TIERS ON AFF_TIERS=T_TIERS WHERE AFF_AFFAIRE="'+ Ch_CHANTIER.text +'"';

  QQ := OpenSql (SQL,true,-1,'',true);
  if not QQ.eof then
  begin
    LCHANTIER.Caption := QQ.findfield('AFF_LIBELLE').asString ;
    //
    Etat := QQ.findfield('AFF_ETATAFFAIRE').asstring;
    CodeDomaine := QQ.findfield('AFF_DOMAINE').AsString;
    //
    LLabelTiers.Caption := QQ.findfield('AFF_TIERS').asString + ' ' + QQ.findfield('T_LIBELLE').asString;
    //
    if CHOIX_APPEL.checked then
    Begin
      CODERESSOURCE.Text := QQ.findfield('AFF_RESPONSABLE').asString ;
      CODETIERS.text := QQ.findfield('AFF_TIERS').asString ;
      CONTRAT.Text := QQ.findfield('AFF_AFFAIREINIT').asString ;
      BBlocnote1.Visible  := True;
      //Chargement du fichier OLE
      Statut := QQ.findfield('AFF_STATUTAFFAIRE').asstring;
      Appel := QQ.findfield('AFF_AFFAIRE').asstring;
      LLabelHeure.Caption := 'Appel : ' + QQ.findfield('AFF_LIBELLE').asString;
      LLABELETAT.Caption := RechDom('BTETATAPPEL',   QQ.findfield('AFF_ETATAFFAIRE').asString, False);
      ChargementFichierOLE(statut, Etat, Appel);
    End
    else
    begin
      LLABELETAT.Caption := RechDom('AFETATAFFAIRE', QQ.findfield('AFF_ETATAFFAIRE').asString, False);
      //
      if CHOIX_CONTRAT.Checked then
        LLabelHeure.Caption := 'Contrat : ' +QQ.findfield('AFF_LIBELLE').asString
      else if CHOIX_CHANTIER.Checked then
        LLabelHeure.Caption := 'Chantier : ' + QQ.findfield('AFF_LIBELLE').asString
    end;
    //
  end;

  ferme (QQ);

  LLIBCONTRAT.Caption := GetLibelleContrat(contrat.text);

end;
Procedure TOF_BTSAISIECONSO.ChargementFichierOLE(Statut, Etat, Appel : string);
var QQ  : TQuery;
    Req : String;
Begin

  //Modification d'un enregistrement
  req := 'SELECT * FROM LIENSOLE WHERE ';
  Req := Req + ' LO_TABLEBLOB = "' + Statut + '"';
  Req := Req + ' AND LO_QUALIFIANTBLOB = "MOT"';
  Req := Req + ' AND LO_EMPLOIBLOB = "' + Etat + '"';
  Req := Req + ' AND LO_IDENTIFIANT ="' + Appel + '"';
  Req := Req + ' AND LO_RANGBLOB = 1';

  QQ := OpenSql (Req,true,-1,'',true);

  if Not QQ.eof then
     begin
     SetControlText('RETOURINT', QQ.findfield('LO_OBJET').asString);
     end;

  ferme (QQ);

end;

procedure TOF_BTSAISIECONSO.PhaseChange(Sender: TObject);
begin
  LIBPhase.Caption := GetLibellePhase (Ch_CHANTIER.text,Phase.text);
end;

procedure TOF_BTSAISIECONSO.FONCTIONChange (Sender : Tobject);
begin
  SetFonction;
end;

procedure TOF_BTSAISIECONSO.SetFonction;
begin
  if FONCTION.Value <> '' then
     MAINDOEUVRE.plus := '(ARS_TYPERESSOURCE="SAL" AND ARS_FONCTION1="'+ FONCTION.Value  +'") AND (ARS_FERME<>"X")'
  else
     MAINDOEUVRE.plus := 'ARS_TYPERESSOURCE="SAL" AND (ARS_FERME<>"X")';
end;

procedure TOF_BTSAISIECONSO.MAINDOEUVREExit (Sender : TOBject);
begin

  LMAINDOEUVRE.Caption := GetMainDoeuvre (MAINDOEUVRE.Text);
  if LMAINDOEUVRE.Caption= '' then MAINDOEUVRE.Text := '';

end;

procedure TOF_BTSAISIECONSO.TYPEPRESTATIONCHANGE (Sender : TOBject);
begin

  if TYPERESSOURCE.Value = '' then
  begin
    RESSOURCES.Plus := '(ARS_TYPERESSOURCE = "MAT"  OR ARS_TYPERESSOURCE="OUT" ) AND (ARS_FERME<>"X")';
  end
  else
  begin
    RESSOURCES.Plus := 'ARS_TYPERESSOURCE="'+TYPERESSOURCE.Value+'" AND (ARS_FERME<>"X")';
  end;

end;

procedure TOF_BTSAISIECONSO.TYPEPRESTATIONEXTCHANGE (Sender : TOBject);
begin

  if TYPERESSOURCEEXT.Value = '' then
  begin
    RESSOURCESEXT.Plus := '(ARS_TYPERESSOURCE = "INT" OR ARS_TYPERESSOURCE="ST" OR ARS_TYPERESSOURCE="AUT" OR ARS_TYPERESSOURCE="LOC") AND (ARS_FERME<>"X")';
  end
  else
  begin
    RESSOURCESEXT.Plus := 'ARS_TYPERESSOURCE="'+TYPERESSOURCEEXT.Value+'" AND (ARS_FERME<>"X")';
  end;

end;

procedure TOF_BTSAISIECONSO.RESSOURCEExit(Sendre : TObject);
begin
  LRESSOURCE.Caption := GetMainDoeuvre (RESSOURCES.Text);
  if LRESSOURCE.Caption= '' then RESSOURCES.Text := '';
end;

procedure TOF_BTSAISIECONSO.RESSOURCEEXTExit (Sendre : TObject);
begin
  LRESSOURCEEXT.Caption := GetMainDoeuvre (RESSOURCESEXT.Text);
  if LRESSOURCEEXT.Caption= '' then RESSOURCESEXT.Text := '';
end;

procedure TOF_BTSAISIECONSO.MATERIAUXExit (Sender : TObject);
var TheArticle 	: string;
		QQ 					: TQuery;
    Requete			: string;
begin

	TOBMateriau.InitValeurs;

  //TheArticle := trim(ARTICLE.text);
  TheArticle := trim(MATERIAUX.text);
  if Length(TheArticle)>18 then
  begin
    Requete := 'SELECT GA_STATUTART,GA_ARTICLE FROM ARTICLE WHERE ' +
               'GA_ARTICLE="'+TheArticle+'" AND GA_TYPEARTICLE IN ("ARP","MAR")';
    //Lecture de l'article
  end else
  begin
    Requete := 'SELECT GA_STATUTART,GA_ARTICLE FROM ARTICLE WHERE ' +
               'GA_CODEARTICLE="'+TheArticle+'" AND GA_TYPEARTICLE IN ("ARP","MAR")';
  end;

  QQ := OpenSql(Requete ,true,-1,'',true);
  if not QQ.eof then
  begin
    if (QQ.findField ('GA_STATUTART').asString = 'DIM') then
      TheArticle  := SelectUneDimension (QQ.findField ('GA_ARTICLE').asString)
    else
      TheArticle := QQ.findField ('GA_ARTICLE').asString;
    ferme (QQ);
    ARTICLE.Text := TheArticle;
    Requete := 'SELECT GA_STATUTART,GA_ARTICLE FROM ARTICLE WHERE ' +
               ' GA_ARTICLE="'+ TheArticle +'" AND GA_TYPEARTICLE IN ("ARP","MAR")';
    QQ := OpenSql (Requete ,true,-1,'',true);
  end;

  if QQ.eof then TheArticle := '';
  ferme (QQ);
  
  LMATERIAUX.Caption := GetLibelleMateriaux(TheArticle);
  
  MATERIAUX.text := TheArticle;

  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TheArticle+'"',true,-1,'',true);
  TOBMateriau.selectDb ('',QQ);
  ferme (QQ);

end;

procedure TOF_BTSAISIECONSO.MATERIAUXRECH (Sender : TObject);
var stchamps  : string;
    NoErreur  : Single;
begin
  stChamps := 'XX_WHERE=GA_CODEARTICLE LIKE "'+MATERIAUX.text+'%" AND ((GA_TYPEARTICLE="MAR") OR (GA_TYPEARTICLE="ARP"))';
	stChamps := stChamps + ';FIXEDTYPEART';
  MATERIAUX.Text    := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps);
  ARTICLE.text  := Trim(Copy(MATERIAUX.Text ,1,18));
end;


function TOF_BTSAISIECONSO.ControleSaisieOK : boolean;
var year,Month,Day : word;
		Prest     : string;
    NoErreur  : single;
begin

  result := false;
  decodedate (Calendrier.SelectionStart,year,month,day);

  //FV1 : 17/12/2014 - FS#1354 - NUANCE3, VERRE&METAL : Ajout contrôle date limites de saisies paramétrées
  if not CtrlDateLimiteSaisie(Trunc(Calendrier.SelectionStart)) then
  begin
    PgiBox (TraduireMemoire('La date que vous avez renseignée est hors des dates limites de saisies'));
    Exit;
  end;

  if not CtrlDateLimiteSaisie(Trunc(Calendrier.SelectionEnd)) then
  begin
    PgiBox (TraduireMemoire('La date que vous avez renseignée est hors des dates limites de saisies'));
    Exit;
  end;

  // calcul du nombre de jour dans la sélection
  if trunc(Calendrier.SelectionEnd) - trunc(Calendrier.SelectionStart) + 1 > DaysPerMonth (year,Month) then
  	 begin
     PgiBox (TraduireMemoire('Le nombre de jours de la sélection excède le nombre de jour dans le mois'),ecran.caption);
     exit;
	   end;

  //controle existance du chantier si coché
  if (CHOIX_CHANTIER.Checked) then
  begin
	  OnExitPartieAffaire (self);
    if CH_CHANTIER.text = '' Then
    begin
      PGIBox (TraduireMemoire('Vous devez renseigner un chantier'),ecran.caption);
      CH_CHANTIER1.SetFocus;
      exit;
    end;
    if not OkChantier (CH_CHANTIER.text) then
    begin
      PGIBox (TraduireMemoire('Chantier inexistant'),ecran.caption);
      CH_CHANTIER.text := '';
      CH_CHANTIER1.SetFocus;
      exit;
    end;
  end;

    //controle existance du contrat si coché
  if (CHOIX_CONTRAT.Checked) then
  begin
    OnExitPartieAffaire (self);
    if CH_CHANTIER.text = '' Then
    begin
      PGIBox (TraduireMemoire('Vous devez renseigner un contrat'),ecran.caption);
      CH_CHANTIER1.SetFocus;
      exit;
    end;
    if not OkChantier (CH_CHANTIER.text) then
    begin
      PGIBox (TraduireMemoire('Contrat inexistant'),ecran.caption);
      CH_CHANTIER.text := '';
      CH_CHANTIER1.SetFocus;
      exit;
    end;
      end;

  //controle existance de l'appel si coché
  if (CHOIX_APPEL.Checked)then
  begin
	  OnExitPartieAffaire (self);
    if CH_CHANTIER.text = '' Then
    begin
      PGIBox (TraduireMemoire('Vous devez renseigner un Appel'),ecran.caption);
      BBlocnote1.Visible  := False;
      CH_CHANTIER1.SetFocus;
      exit;
    end;
    if not OkChantier (CH_CHANTIER.text) then
    begin
        PGIBox (TraduireMemoire('Appel inexistant'),ecran.caption);
        CH_CHANTIER.text := '';
        BBlocnote1.Visible  := False;
        CH_CHANTIER1.SetFocus;
        exit;
    end;
  end;

  //FV1 : 10/06/2014 - FS#921 - DELABOUDINIERE : Revoir les contrôles sur appels et contrats en fonction du code état
  if (CHOIX_CHANTIER.Checked) Or
     (CHOIX_CONTRAT.Checked)  Or
     (CHOIX_APPEL.Checked)    then
  begin
    if not ControleAffaire(Ch_chantier.text,Ecran.caption,'SCO') then
    begin
      CH_CHANTIER.text := '';
      CH_CHANTIER1.SetFocus;
      exit;
    end;
  end;
  //
  if CHOIX_MATERIAUX.checked Then
  begin
    if MATERIAUX.Text  = '' Then
    begin
      PGIBox (TraduireMemoire('Vous devez renseigner une fourniture'),ecran.caption);
      MATERIAUX.SetFocus;
      exit;
    end;
    if not OkMATERIAUX (MATERIAUX.text) then
    begin
      //PGIBox (TraduireMemoire('Fourniture inexistante'),ecran.caption);
      MATERIAUX.SetFocus;
      exit;
    end;
     //FV1 : 10/12/2014 - FS#1342 - SCETEC - saisie de consos matériaux
      MATERIAUXExit(Self);
  end
  else if CHOIX_MO.checked then
  begin
    if MAINDOEUVRE.Text = '' then
    begin
      PGIBox (TraduireMemoire('Vous devez renseigner une main d''oeuvre'),ecran.caption);
      MAINDOEUVRE.SetFocus;
      exit;
    end;
    NoErreur := OKRessource (nil,MAINDOEUVRE.text,'"SAL","INT"',Prest,false,FONCTION.value,TRue );
    if NoErreur = 0 then
    begin
      PGIBox (TraduireMemoire('Main d''oeuvre inexistante'),ecran.caption);
      MAINDOEUVRE.SetFocus;
      exit;
    end
    Else if NoErreur = 2 then
    begin
      PGIBox (TraduireMemoire('Main d''oeuvre Fermée'),ecran.caption);
      MAINDOEUVRE.SetFocus;
      exit;
    end;
  end
  else if CHOIX_RESSOURCE.checked then
  begin
    if RESSOURCES.Text = '' then
    begin
      PGIBox (TraduireMemoire('Vous devez renseigner une ressource'),ecran.caption);
      RESSOURCES.SetFocus;
      exit;
    end;
    NoErreur := OKRessource (Nil,RESSOURCES.text,'"'+TYPERESSOURCE.value+'"',Prest,false,'',TRue);
    if NoErreur = 0 then
    begin
      PGIBox (TraduireMemoire('Ressource inexistante'),ecran.caption);
      RESSOURCES.SetFocus;
      exit;
    end
    Else if NoErreur = 2 then
    begin
      PGIBox (TraduireMemoire('Ressource Fermée'),ecran.caption);
      RESSOURCES.SetFocus;
      exit;
    end
  end else if CHOIX_MOEXT.checked then
  begin
    if RESSOURCESEXT.Text = '' then
    begin
      PGIBox (TraduireMemoire('Vous devez renseigner une main d''oeuvre externe'),ecran.caption);
      RESSOURCESEXT.SetFocus;
      exit;
    end;
    NoErreur := OKRessource (nil,RESSOURCESEXT.text,'"'+TYPERESSOURCEEXT.value+'"',Prest,false,'',TRue);
    if NoErreur = 0  then
    begin
      PGIBox (TraduireMemoire('Ressource inexistante'),ecran.caption);
      RESSOURCESEXT.SetFocus;
      exit;
    end
    Else if NoErreur = 2  then
    begin
      PGIBox (TraduireMemoire('Ressource Fermée'),ecran.caption);
      RESSOURCESEXT.SetFocus;
      exit;
    end;

  end;
  result := true;
end;

procedure TOF_BTSAISIECONSO.SetparamGrilles;
var LaListe,lelement : string;
begin
  // récupération du paramétrage général des grilles
  NomList := 'BTSAISIECONSO';
  ChargeHListe (NomList, FNomTable, FLien, FSortBy, stCols, FTitre,
                FLargeur, FAlignement, FParams, title, NC, FPerso, OkTri, OkNumCol);

  // Ajout des 2 colonnes de bases : selection + jour du mois
  stColListe := 'SELECT;JOUR;'+stCols;
  flargeur := '2;3;'+flargeur;
  Falignement := 'C.0  ---;C.0  ---;'+Falignement;
  FTitre := ' ;Jour;'+Ftitre;
  NC := '1;1;'+NC;
  nbcolsInListe := 0;
  LaListe := stColListe;
  G_SELECT := 0;
  G_JOUR := 1;
  repeat
    lelement := READTOKENST (laListe);
    if lelement <> '' then
    begin
      if lelement = 'SELECT'            then G_SELECT       := nbcolsinListe else
      if lelement = 'JOUR'              then G_JOUR         := nbcolsinListe else
      if lelement = 'BCO_NATUREMOUV'    then G_NATURE       := nbcolsinListe else
      if lelement = 'BCO_RESSOURCE'     then G_RESSOURCE    := nbcolsinListe else
      if lelement = 'BCO_CODEARTICLE'   then G_CODEARTICLE  := nbcolsinListe else
      if lelement = 'BCO_AFFAIRE0'      then G_AFFAIRE0     := nbcolsinListe else
      if lelement = 'BCO_AFFAIRE'       then G_AFFAIRE      := nbcolsinListe else
      if lelement = 'BCO_PHASETRA'      then G_PHASE        := nbcolsinListe else
      if lelement = 'BCO_LIBELLE'       then G_LIBELLE      := nbcolsinListe else
      if lelement = 'BCO_QUANTITE'      then G_QTE          := nbcolsinListe else
      if lelement = 'BCO_QUALIFQTEMOUV' then G_QUALIF       := nbcolsinListe else
      if lelement = 'BCO_MONTANTPR'     then G_MONTANT      := nbcolsinListe else
      if lelement = 'BCO_DPR'           then G_PU           := nbcolsinListe else
      //if lelement = 'BCO_FACTURABLE'  then G_FACT         := nbcolsinListe else
      if lelement = 'BCO_TYPEHEURE'     then G_TYPEHEURE    := nbcolsinListe;
      inc(nbcolsInListe);
    end;
  until lelement = '';
end;

procedure TOF_BTSAISIECONSO.SaisieChantier;
begin

  GHeures.colwidths[G_AFFAIRE0] := 0;
  GHeures.colLengths[G_AFFAIRE0] := -1;
  GHeures.colwidths[G_AFFAIRE] := 0;
  GHeures.ColLengths [G_AFFAIRE] := -1;
  GHeures.ColWidths [G_NATURE] := 0;

  if (PHASE.Text <> '') or (not WithPhases) then
     begin
     GHeures.colwidths[G_PHASE] := 0;
     GHeures.ColLengths [G_PHASE] := -1;
     end;

  //
  GFrais.colwidths[G_AFFAIRE0] := 0;
  GFrais.colLengths[G_AFFAIRE0] := -1;
  GFrais.colwidths[G_AFFAIRE] := 0;
  GFrais.ColLengths [G_AFFAIRE] := -1;
  GFrais.ColLengths [G_NATURE] := -1;
  GFrais.ColWidths [G_NATURE] := 0;

  if (PHASE.Text <> '') or (not WithPhases) then
     begin
     GFrais.colwidths[G_PHASE] := 0;
     GFrais.ColLengths [G_PHASE] := -1;
     end;

  GFrais.cells[G_CODEARTICLE,0] := TraduireMemoire('Frais');
  GFrais.ColLengths [G_TYPEHEURE] := -1;
  GFrais.ColWidths [G_TYPEHEURE] := 0;

  //
  GMateriel.colwidths[G_AFFAIRE0] := 0;
  GMateriel.colLengths[G_AFFAIRE0] := -1;
  GMateriel.colwidths[G_AFFAIRE] := 0;
  GMateriel.ColLengths [G_AFFAIRE] := -1;
  GMateriel.ColLengths [G_NATURE] := -1;
  GMateriel.ColWidths [G_NATURE] := 0;

  if (PHASE.Text <> '') or (not WithPhases) then
  begin
    GMateriel.colwidths[G_PHASE] := 0;
    GMateriel.ColLengths [G_PHASE] := -1;
  end;

  GMateriel.cells[G_CODEARTICLE,0] := TraduireMemoire('Prestation');
  GMateriel.ColLengths [G_TYPEHEURE] := -1;
  GMateriel.ColWidths [G_TYPEHEURE] := 0;
  //
  GMOEXT.colwidths[G_AFFAIRE0] := 0;
  GMOEXT.colLengths[G_AFFAIRE0] := -1;
  GMOEXT.colwidths[G_AFFAIRE] := 0;
  GMOEXT.ColLengths [G_AFFAIRE] := -1;
  GMOEXT.ColLengths [G_NATURE] := -1;
  GMOEXT.ColWidths [G_NATURE] := 0;

  if (PHASE.Text <> '') or (not WithPhases) then
  begin
    GMOEXT.colwidths[G_PHASE] := 0;
    GMOEXT.ColLengths [G_PHASE] := -1;
  end;
  GMOEXT.cells[G_CODEARTICLE,0] := TraduireMemoire('Prestation');
  GMOEXT.ColLengths [G_TYPEHEURE] := -1;
  GMOEXT.ColWidths [G_TYPEHEURE] := 0;
  //
  GFourniture.colwidths[G_AFFAIRE0] := 0;
  GFourniture.colLengths[G_AFFAIRE0] := -1;
  GFourniture.colwidths[G_AFFAIRE] := 0;
  GFourniture.ColLengths [G_AFFAIRE] := -1;
  GFourniture.ColLengths [G_NATURE] := -1;
  GFourniture.ColWidths [G_NATURE] := 0;

  if (PHASE.Text <> '') or (not WithPhases) then
  begin
    GFourniture.colwidths[G_PHASE] := 0;
    GFourniture.ColLengths [G_PHASE] := -1;
  end;
  GFourniture.colwidths[G_RESSOURCE] := 0;
  GFourniture.ColLengths [G_RESSOURCE] := -1;
  GFourniture.cells[G_CODEARTICLE,0] := TraduireMemoire('Article');
  GFourniture.ColLengths [G_TYPEHEURE] := -1;
  GFourniture.ColWidths [G_TYPEHEURE] := 0;
  //
  GRecettes.colwidths[G_AFFAIRE0] := 0;
  GRecettes.colLengths[G_AFFAIRE0] := -1;
  Grecettes.colwidths[G_AFFAIRE] := 0;
  Grecettes.ColLengths [G_AFFAIRE] := -1;

  if (PHASE.Text <> '') or (not WithPhases) then
  begin
    Grecettes.colwidths[G_PHASE] := 0;
    Grecettes.ColLengths [G_PHASE] := -1;
  end;

  Grecettes.colwidths[G_RESSOURCE] := 0;
  Grecettes.ColLengths [G_RESSOURCE] := -1;
  Grecettes.colwidths[G_CODEARTICLE] := 0;
  Grecettes.ColLengths [G_CODEARTICLE] := -1;
  Grecettes.colwidths[G_QTE] := 0;
  Grecettes.ColLengths [G_QTE] := -1;
  Grecettes.colwidths[G_QUALIF] := 0;
  Grecettes.ColLengths [G_QUALIF] := -1;
  //Grecettes.colwidths[G_FACT] := 0;
  //Grecettes.ColLengths [G_FACT] := -1;
  Grecettes.colwidths[G_PU] := 0;
  Grecettes.ColLengths [G_PU] := -1;
  Grecettes.ColLengths [G_TYPEHEURE] := -1;
  Grecettes.ColWidths [G_TYPEHEURE] := 0;

  TFFiche(ecran).HMTrad.ResizeGridColumns (Gheures);
  TFFiche(ecran).HMTrad.ResizeGridColumns (GFourniture);
  TFFiche(ecran).HMTrad.ResizeGridColumns (GFrais);
  TFFiche(ecran).HMTrad.ResizeGridColumns (GMateriel);
  TFFiche(ecran).HMTrad.ResizeGridColumns (GMOEXT);
  TFFiche(ecran).HMTrad.ResizeGridColumns (GRecettes);

end;

procedure TOF_BTSAISIECONSO.SaisieMateriaux;
begin
  //
  GFourniture.ColLengths [G_NATURE] := -1;
  GFourniture.ColWidths [G_NATURE] := 0;
  GFourniture.colwidths[G_CODEARTICLE] := 0;
  GFourniture.ColLengths [G_CODEARTICLE] := -1;
  GFourniture.colwidths[G_RESSOURCE] := 0;
  GFourniture.ColLengths [G_RESSOURCE] := -1;
  GFourniture.cells[G_CODEARTICLE,0] := TraduireMemoire('Article');
  GFourniture.ColLengths [G_TYPEHEURE] := -1;
  GFourniture.ColWidths [G_TYPEHEURE] := 0;
  TFFiche(ecran).HMTrad.ResizeGridColumns (GFourniture);
end;

procedure TOF_BTSAISIECONSO.SaisieMO;
begin
  GHeures.colwidths[G_RESSOURCE] := 0;
  GHeures.ColLengths [G_RESSOURCE] := -1;
  GHeures.ColLengths [G_NATURE] := -1;
  GHeures.ColWidths [G_NATURE] := 0;
  //
  GFrais.colwidths[G_RESSOURCE] := 0;
  GFrais.ColLengths [G_RESSOURCE] := -1;
  GFrais.cells[G_CODEARTICLE,0] := TraduireMemoire('Frais');
  GFrais.ColLengths [G_NATURE] := -1;
  GFrais.ColWidths [G_NATURE] := 0;
  GFrais.ColLengths [G_TYPEHEURE] := -1;
  GFrais.ColWidths [G_TYPEHEURE] := 0;
  //
  GFourniture.ColLengths [G_NATURE] := -1;
  GFourniture.ColWidths [G_NATURE] := 0;
  GFourniture.colwidths[G_RESSOURCE] := 0;
  GFourniture.ColLengths [G_RESSOURCE] := -1;
  GFourniture.cells[G_CODEARTICLE,0] := TraduireMemoire('Article');
  GFourniture.ColLengths [G_TYPEHEURE] := -1;
  GFourniture.ColWidths [G_TYPEHEURE] := 0;
  //
  TFFiche(ecran).HMTrad.ResizeGridColumns (Gheures);
  TFFiche(ecran).HMTrad.ResizeGridColumns (GFrais);
  TFFiche(ecran).HMTrad.ResizeGridColumns (GFourniture);
end;


procedure TOF_BTSAISIECONSO.SaisieMOExt;
begin
  GMOext.colwidths[G_RESSOURCE] := 0;
  GMOext.ColLengths [G_RESSOURCE] := -1;
  GMOext.cells[G_CODEARTICLE,0] := TraduireMemoire('Prestation');
  GMOext.ColLengths [G_NATURE] := -1;
  GMOext.ColWidths [G_NATURE] := 0;
  GMOext.ColLengths [G_TYPEHEURE] := -1;
  GMOext.ColWidths [G_TYPEHEURE] := 0;
  //
  TFFiche(ecran).HMTrad.ResizeGridColumns (GMOext);
end;

procedure TOF_BTSAISIECONSO.Saisieressource;
begin
  GMateriel.colwidths[G_RESSOURCE] := 0;
  GMateriel.ColLengths [G_RESSOURCE] := -1;
  GMateriel.cells[G_CODEARTICLE,0] := TraduireMemoire('Prestation');
  GMateriel.ColLengths [G_NATURE] := -1;
  GMateriel.ColWidths [G_NATURE] := 0;
  GMateriel.ColLengths [G_TYPEHEURE] := -1;
  GMateriel.ColWidths [G_TYPEHEURE] := 0;
  //
  TFFiche(ecran).HMTrad.ResizeGridColumns (GMateriel);
end;

procedure TOF_BTSAISIECONSO.DefiniGrilles;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec : integer;
    indice : integer;
    FFPU : string;
begin
  FFPU := '#';
  if V_PGI.OkDecP > 0 then
  begin
    FFPU := '0.';
    for indice := 1 to V_PGI.OkDecP - 1 do
    begin
      FFPU := FFPU + '#';
    end;
    FFPU := FFPU + '0';
  end;

  GHeures.ColCount := NbColsInListe;
  GFrais.ColCount := NbColsInListe;
  GMAteriel.ColCount := NbColsInListe;
  GFourniture.ColCount := NbColsInListe;
  GMOEXT.ColCount := NbColsInListe;
  GRecettes.ColCount := NbColsInListe;

  st := stColListe;
  lesalignements := Falignement;
  lestitres := FTitre;
  leslargeurs := flargeur;

  for indice := 0 to nbcolsInListe -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;

    GHeures.cells[Indice,0] := leTitre;
    GFrais.cells[Indice,0] := leTitre;
    GMateriel.cells[Indice,0] := leTitre;
    GFourniture.cells[Indice,0] := leTitre;
    GMOEXT.cells[Indice,0] := leTitre;
    GRecettes.cells[Indice,0] := leTitre;

    if copy(Alignement,1,1)='G' then
    begin
      GHeures.ColAligns[indice] := taLeftJustify;
      GFrais.ColAligns[indice] := taLeftJustify;
      GMateriel.ColAligns[indice] := taLeftJustify;
      GFourniture.ColAligns[indice] := taLeftJustify;
      GMOEXT.ColAligns[indice] := taLeftJustify;
      GRecettes.ColAligns[indice] := taLeftJustify;
    end else
    if copy(Alignement,1,1)='D' then
    begin
      GHeures.ColAligns[indice] := taRightJustify;
      GFrais.ColAligns[indice] := taRightJustify;
      GMateriel.ColAligns[indice] := taRightJustify;
      GFourniture.ColAligns[indice] := taRightJustify;
      GMOExt.ColAligns[indice] := taRightJustify;
      GRecettes.ColAligns[indice] := taRightJustify;
    end else
    if copy(Alignement,1,1)='C' then
    begin
      GHeures.ColAligns[indice] := taCenter;
      GFrais.ColAligns[indice] := taCenter;
      GMateriel.ColAligns[indice] := taCenter;
      GFourniture.ColAligns[indice] := taCenter;
      GMOext.ColAligns[indice] := taCenter;
      GRecettes.ColAligns[indice] := taCenter;
    end;

    GHeures.ColWidths[indice] := strtoint(lalargeur);
    GFrais.ColWidths[indice] := strtoint(lalargeur);
    GMateriel.ColWidths[indice] := strtoint(lalargeur);
    GFourniture.ColWidths[indice] := strtoint(lalargeur);
    GMOExt.ColWidths[indice] := strtoint(lalargeur);
    GRecettes.ColWidths[indice] := strtoint(lalargeur);

    if OkLib then
       begin
       GHeures.ColFormats[indice] := 'CB=' + Get_Join(Nam);
       GFrais.ColFormats[indice] := 'CB=' + Get_Join(Nam);
       GMateriel.ColFormats[indice] := 'CB=' + Get_Join(Nam);
       GFourniture.ColFormats[indice] := 'CB=' + Get_Join(Nam);
       GMOEXt.ColFormats[indice] := 'CB=' + Get_Join(Nam);
       GRecettes.ColFormats[indice] := 'CB=' + Get_Join(Nam);
       end
    else if (Dec<>0) or (Sep) then
       begin
       GHeures.ColFormats[indice] := FF ;
       GFrais.ColFormats[indice] := FF ;
       GMateriel.ColFormats[indice] := FF ;
       GFourniture.ColFormats[indice] := FF ;
       GMOEXt.ColFormats[indice] := FF ;
       GRecettes.ColFormats[indice] := FF ;
       end;

    if nam = 'BCO_DPR' then
    begin
      GHeures.ColFormats[indice] := FFPU ;
      GFrais.ColFormats[indice] := FFPU ;
      GMateriel.ColFormats[indice] := FFPU ;
      GFourniture.ColFormats[indice] := FFPU ;
      GMOEXt.ColFormats[indice] := FFPU ;
      GRecettes.ColFormats[indice] := FFPU ;
    end;

    if nam = 'BCO_LIBELLE' then
    begin
      GHeures.ColLengths[indice] := 35 ;
      GFrais.ColLengths[indice] := 35 ;
      GMateriel.ColLengths[indice] := 35 ;
      GFourniture.ColLengths[indice] := 35 ;
      GMOEXt.ColLengths[indice] := 35 ;
      GRecettes.ColLengths[indice] := 35 ;
    end;

  end ;

  initgrilles;

  // Particularité
  GHeures.cells[G_CODEARTICLE,0] := 'Prestation';

  if CHOIX_CHANTIER.Checked then SaisieChantier else
  if CHOIX_CONTRAT.Checked then SaisieChantier else
  if CHOIX_APPEL.Checked then SaisieChantier else
  if CHOIX_MATERIAUX.Checked then SaisieMateriaux else
  if CHOIX_MO.Checked then SaisieMO else
  if CHOIX_MOEXT.Checked then SaisieMOExt else
  if CHOIX_RESSOURCE.Checked then Saisieressource;

end;

procedure TOF_BTSAISIECONSO.Remplitgrilles;
var indice : integer;
begin

  if (CHOIX_CHANTIER.Checked) OR (CHOIX_CONTRAT.Checked) OR (CHOIX_APPEL.Checked) then
     begin
     if TOBMo.detail.count > 0    then GHeures.RowCount := TOBMO.detail.count +2;
     if GHeures.RowCount < 2      then GHeures.rowcount := 3;
     if TOBFrs.detail.count > 0   then GFrais.RowCount := TOBFRS.detail.count +2;
     if GFrais.RowCount < 2       then GFrais.rowcount := 3;
     if TOBRES.detail.count > 0   then GMateriel.RowCount := TOBRES.detail.count +2;
     if GMateriel.RowCount < 2    then GMateriel.rowcount := 3;
     if TOBMAT.detail.count > 0   then GFourniture.RowCount := TOBMAT.detail.count +2;
     if GFourniture.RowCount < 2  then GFourniture.rowcount := 3;
     if TOBMoEXT.detail.count > 0 then GMOEXT.RowCount := TOBMOEXT.detail.count +2;
     if GMOEXT.RowCount < 2       then GMOEXT.rowcount := 3;
     if TOBRecettes.detail.count > 0 then GRecettes.RowCount := TOBRecettes.detail.count +2;
     if GRecettes.RowCount < 2 then GRecettes.rowcount := 3;
     for Indice := 0 to TOBMO.detail.count -1 do
         begin
         AfficheLigne(GHeures,TOBMO.detail[Indice],Indice);
         end;
     for Indice := 0 to TOBMOEXT.detail.count -1 do
        begin
        AfficheLigne(GMOEXT,TOBMOEXT.detail[Indice],Indice);
        end;
     for Indice := 0 to TOBFRS.detail.count -1 do
        begin
        AfficheLigne(GFrais,TOBFrs.detail[Indice],Indice);
        end;
     for Indice := 0 to TOBMAT.detail.count -1 do
        begin
        AfficheLigne(GFourniture,TOBMAT.detail[Indice],Indice);
        end;
     for Indice := 0 to TOBRES.detail.count -1 do
        begin
        AfficheLigne(GMAteriel,TOBRES.detail[Indice],Indice);
        end;
     for Indice := 0 to TOBRecettes.detail.count -1 do
        begin
        AfficheLigne(GRecettes,TOBRecettes.detail[Indice],Indice);
        end;
     TFFiche(ecran).HMTrad.ResizeGridColumns (Gheures);
     TFFiche(ecran).HMTrad.ResizeGridColumns (GMOEXT);
     TFFiche(ecran).HMTrad.ResizeGridColumns (GFrais);
     TFFiche(ecran).HMTrad.ResizeGridColumns (GMateriel);
     TFFiche(ecran).HMTrad.ResizeGridColumns (GFourniture);
     TFFiche(ecran).HMTrad.ResizeGridColumns (Grecettes);
     end
  else if CHOIX_MATERIAUX.Checked then
     begin
     if TOBMAT.detail.count > 0 then GFourniture.RowCount := TOBMAT.detail.count +2;
     if GFourniture.RowCount < 2 then GFourniture.rowcount := 3;
     for Indice := 0 to TOBMAT.detail.count -1 do
         begin
         AfficheLigne(GFourniture,TOBMAT.detail[Indice],Indice);
         end;
     TFFiche(ecran).HMTrad.ResizeGridColumns (GFourniture);
     end
  else if CHOIX_MO.Checked then
     begin
     if TOBMo.detail.count > 0 then GHeures.RowCount := TOBMO.detail.count +2;
     if GHeures.RowCount < 2 then GHeures.rowcount := 3;
     if TOBFrs.detail.count > 0 then GFrais.RowCount := TOBFRS.detail.count +2;
     if GFrais.RowCount < 2 then GFrais.rowcount := 3;
     if TOBMAT.detail.count > 0 then GFourniture.RowCount := TOBMAT.detail.count +2;
     if GFourniture.RowCount < 2 then GFourniture.rowcount := 3;
     for Indice := 0 to TOBMO.detail.count -1 do
         begin
         AfficheLigne(GHeures,TOBMO.detail[Indice],Indice);
         end;
     for Indice := 0 to TOBFRS.detail.count -1 do
         begin
         AfficheLigne(GFrais,TOBFrs.detail[Indice],Indice);
         end;
     for Indice := 0 to TOBMAT.detail.count -1 do
         begin
         AfficheLigne(GFourniture,TOBMAT.detail[Indice],Indice);
         end;
     TFFiche(ecran).HMTrad.ResizeGridColumns (Gheures);
     TFFiche(ecran).HMTrad.ResizeGridColumns (GFrais);
     TFFiche(ecran).HMTrad.ResizeGridColumns (GFourniture);
     end
  else if CHOIX_RESSOURCE.Checked then
     begin
     if TOBRES.detail.count > 0 then GMateriel.RowCount := TOBRES.detail.count +2;
     if GMateriel.RowCount < 2 then GMateriel.rowcount := 3;
     for Indice := 0 to TOBRES.detail.count -1 do
         begin
         AfficheLigne(GMAteriel,TOBRES.detail[Indice],Indice);
         end;
     TFFiche(ecran).HMTrad.ResizeGridColumns (GMateriel);
     end
  else if CHOIX_MOEXT.Checked then
     begin
     if TOBMOEXT.detail.count > 0 then GMOEXT.RowCount := TOBMOEXT.detail.count +2;
     if GMOEXT.RowCount < 2 then GMOEXT.rowcount := 3;
     for Indice := 0 to TOBMOEXT.detail.count -1 do
         begin
         AfficheLigne(GMOEXT,TOBMOEXT.detail[Indice],Indice);
         end;
     TFFiche(ecran).HMTrad.ResizeGridColumns (GMOEXT);
     end;

end;

procedure TOF_BTSAISIECONSO.AfficheLigne(Grid : THgrid;TOBL : TOB;indice : integer);
begin
  TOBL.PutLigneGrid (Grid,Indice+1,false,false,stColListe);
end;

procedure TOF_BTSAISIECONSO.EventGrilles(Active : boolean);
begin

  if Active then
	  begin
    // grilles de saisie des heure de MO
    GHeures.OnRowEnter := GheuresRowEnter;
    GHeures.OnRowExit := GheuresRowExit;
    GHeures.OnCellEnter  := GheuresCellEnter;
    GHeures.OnCellExit  := GheuresCellExit;
    GHeures.OnElipsisClick := GheuresElipsisClick;
    GHeures.OnKeyDown := GridKeyDown;
    Gheures.PostDrawCell := GheuresPostDrawCell;

    // grilles de saisie des frais
    GFrais.OnRowEnter := GFraisRowEnter;
    GFrais.OnRowExit := GFraisRowExit;
    GFrais.OnCellEnter  := GFraisCellEnter;
    GFrais.OnCellExit  := GFraisCellExit;
    GFrais.OnElipsisClick := GFraisElipsisClick;
    GFrais.OnKeyDown := GridKeyDown;
    GFrais.PostDrawCell := GFraisPostDrawCell;

    // grilles de saisie des ressources externes ou matériels
    GMAteriel.OnRowEnter := GMAterielRowEnter;
    GMAteriel.OnRowExit := GMAterielRowExit;
    GMAteriel.OnCellEnter  := GMAterielCellEnter;
    GMAteriel.OnCellExit  := GMAterielCellExit;
    GMateriel.OnElipsisClick := GMaterielElipsisClick;
    GMateriel.OnKeyDown := GridKeyDown;
    GMateriel.PostDrawCell := GMAterielPostDrawCell;

    // grilles de saisie des fournitures
    GFourniture.OnRowEnter := GFournitureRowEnter;
    GFourniture.OnRowExit := GFournitureRowExit;
    GFourniture.OnCellEnter  := GFournitureCellEnter;
    GFourniture.OnCellExit  := GFournitureCellExit;
    GFourniture.OnElipsisClick := GFournitureElipsisClick;
    GFourniture.PostDrawCell  := GFourniturePostDrawCell;
    GFourniture.OnKeyDown := GridKeyDown;

    // grilles de saisie des heure de MO
    GMOExt.OnRowEnter := GMOEXTRowEnter;
    GMOExt.OnRowExit := GMOExtRowExit;
    GMOExt.OnCellEnter  := GMOExtCellEnter;
    GMOExt.OnCellExit  := GMOExtCellExit;
    GMOExt.OnElipsisClick := GMOExtElipsisClick;
    GMOExt.OnKeyDown := GridKeyDown;
    GMOEXT.PostDrawCell := GMOEXTPostDrawCell;

    // grilles de saisie des recettes et frais annexes
    GRecettes.OnRowEnter := GRecettesRowEnter;
    GRecettes.OnRowExit := GRecettesRowExit;
    GRecettes.OnCellEnter  := GRecettesCellEnter;
    GRecettes.OnCellExit  := GRecettesCellExit;
    GRecettes.OnKeyDown := GridKeyDown;
    Grecettes.OnElipsisClick := GRecettesElipsisClick;
    Grecettes.GetCellCanvas  := GrecettesGetCellCanvas;
    Grecettes.PostDrawCell  := GrecettesPostDrawCell;
  	end
  else
  	begin
    // grilles de saisie des heure de MO
    GHeures.OnRowEnter := nil;
    GHeures.OnRowExit := nil;
    GHeures.OnCellEnter  := nil;
    GHeures.OnCellExit  := nil;
    GHeures.OnElipsisClick := nil;
    Gheures.PostDrawCell := nil;
    // grilles de saisie des frais
    GFrais.OnRowEnter := nil;
    GFrais.OnRowExit := nil;
    GFrais.OnCellEnter  := nil;
    GFrais.OnCellExit  := nil;
    GFrais.OnElipsisClick := nil;
    GFrais.PostDrawCell := nil;
    // grilles de saisie des ressources externes ou matériels
    GMAteriel.OnRowEnter := nil;
    GMAteriel.OnRowExit := nil;
    GMAteriel.OnCellEnter  := nil;
    GMAteriel.OnCellExit  := nil;
    GMateriel.OnElipsisClick := nil;
    GMateriel.PostDrawCell := GMAterielPostDrawCell;
    // grilles de saisie des fournitures
    GFourniture.OnRowEnter := nil;
    GFourniture.OnRowExit := nil;
    GFourniture.OnCellEnter  := nil;
    GFourniture.OnCellExit  := nil;
    GFourniture.OnElipsisClick := nil;
    // grilles de saisie des heure de MO
    GMOExt.OnRowEnter := nil;
    GMOExt.OnRowExit := nil;
    GMOExt.OnCellEnter  := nil;
    GMOExt.OnCellExit  := nil;
    GMOExt.OnElipsisClick := nil;
    GMOEXT.PostDrawCell := GMOEXTPostDrawCell;
    // grilles de saisie des recettes et frais annexes
    GRecettes.OnRowEnter := nil;
    GRecettes.OnRowExit := nil;
    GRecettes.OnCellEnter  := nil;
    GRecettes.OnCellExit  := nil;
    GRecettes.OnKeyDown := nil;
    GRecettes.OnElipsisClick := nil;
  	end;
// --
end;

procedure TOF_BTSAISIECONSO.GheuresRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowEnterGrid(Ou, Gheures, TOBMO, TgsMo, Cancel);

end;

procedure TOF_BTSAISIECONSO.GheuresRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowExitGrid(Ou, Gheures, TgsMO, Cancel);

end;

procedure TOF_BTSAISIECONSO.GheuresElipsisClick (Sender : TOBject);
var TOBL : TOB;
begin

  TOBL := TOBMO.detail[GHeures.row-1] ;

  if Gheures.col = G_JOUR then
  	  RechDateGrid(GHeures)
  else if Gheures.col = G_RESSOURCE then
     RechRessourceGrid(GHeures, 'TYPERESSOURCE=SAL,INT')
  else if Gheures.col = G_CODEARTICLE then
     RechPrestationGrid(GHeures, TOBL)
  else if Gheures.col = G_AFFAIRE0 then
	  AfficheNature(GHeures)
  else if Gheures.col = G_AFFAIRE then
	  RechAffaireGrid(GHeures, TOBL)
  else if Gheures.col = G_PHASE then
     RechPhaseGrid(GHeures, TOBL)
  else if Gheures.col = G_TYPEHEURE then
     LookupList (THEdit(Sender),'Type d''heure','CHOIXCOD','CC_CODE','CC_LIBELLE','CC_TYPE="ATH"','',true,0);
  //else if Gheures.col = G_FACT then
  //  LookupList (THEdit(Sender),'Sélection du mode','COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="ARE" AND CO_LIBRE="ART"','',true,0);

end;

procedure TOF_BTSAISIECONSO.GheuresCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  if Action = taConsult then Exit;

  if not Cancel then
  begin
  ZoneSuivanteOuOk(GHeures,ACol, ARow, Cancel);

    Gheures.ElipsisButton := ((GHeures.Col = G_JOUR)       and (Calendrier.SelectionStart<>Calendrier.SelectionEnd)) or
                           (GHeures.col = G_RESSOURCE) or
                           (GHeures.col = G_CODEARTICLE) or
                           (GHeures.col = G_TYPEHEURE) or
                           (GHeures.col = G_AFFAIRE0) or
                           (GHeures.col = G_AFFAIRE) or
                           //(GHeures.col = G_FACT) or
                           (GHeures.col = G_PHASE);
    //                          
     stprevHeure := GHeures.Cells [Acol,Arow];
     AfficheLibelle(TgsMO,Acol,Arow);
     SetToolValue (TgsMo,Arow);
     end;


end;

procedure TOF_BTSAISIECONSO.GheuresCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  CellExitGrid(TOBMO, GHeures, Acol, Arow,TgSMO, StPrevHeure,Cancel);

end;

// gestion Grid des frais

procedure TOF_BTSAISIECONSO.GFraisRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowEnterGrid(Ou, Gfrais, TOBFrs, TgsFrs, Cancel);

end;

procedure TOF_BTSAISIECONSO.GFraisRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowExitGrid(Ou, GFrais, TgsFrs, Cancel);

end;

procedure TOF_BTSAISIECONSO.GFraisElipsisClick (Sender : TOBject);
var TOBL : TOB;
begin

  TOBL := TOBFRS.detail[GFrais.row-1] ;

  if GFrais.col = G_JOUR then
  	RechDateGrid(GFrais)
	else if GFrais.col = G_RESSOURCE then
  	RechRessourceGrid(GFrais, 'TYPERESSOURCE=SAL,INT')
  else if GFrais.col = G_CODEARTICLE then
		RechFraisGrid(GFrais, TOBL, '(GA_TYPEARTICLE="FRA")')
  else if GFrais.col = G_AFFAIRE0 then
		AfficheNature(GFrais)
  else if GFrais.col = G_AFFAIRE then
		RechAffaireGrid(GFrais, TOBL)
  else if GFrais.col = G_PHASE then
    RechPhaseGrid(GFrais, TOBL);
  //else if GFrais.col = G_FACT then
  //  LookupList (THEdit(Sender),'Sélection du mode','COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="ARE" AND CO_LIBRE="ART"','',true,0);

end;

procedure TOF_BTSAISIECONSO.GFraisCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  if Action = taConsult then Exit;

  ZoneSuivanteOuOk(GFrais,ACol, ARow, Cancel);

  GFrais.ElipsisButton := ((GFrais.Col = G_JOUR) and (Calendrier.SelectionStart<>Calendrier.SelectionEnd) ) or
                           (GFrais.col = G_RESSOURCE) or
                           (GFrais.col = G_CODEARTICLE) or
                           (GFrais.col = G_AFFAIRE0) or
													 (GFrais.col = G_AFFAIRE) or
                           //(GFrais.col = G_FACT) or
                           (GFrais.col = G_PHASE);

  if not Cancel then
  begin
    stPrevFrs := GFrais.Cells [Acol,Arow];
    SetToolValue (TgsFrs,Arow);
    AfficheLibelle(TgsFrs,Acol,Arow);
  end;

end;

procedure TOF_BTSAISIECONSO.GFraisCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  CellExitGrid(TOBFrs, GFrais, Acol, Arow,TgsFrs, StPrevFrs,Cancel);
end;
// Materiels
procedure TOF_BTSAISIECONSO.GMaterielRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowEnterGrid(Ou, GMateriel, TOBRES, TgsRES, Cancel);

end;

procedure TOF_BTSAISIECONSO.GMaterielRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowExitGrid(Ou, GMateriel, TgsRES, Cancel);

end;

procedure TOF_BTSAISIECONSO.GMaterielElipsisClick (Sender : TOBject);
var TOBL : TOB;
    stchamps,Article : string;
    result : string;
    Action : String;
begin

  TOBL := TOBRES.detail[GMateriel.row-1] ;

  Action := ';ACTION=RECH';


  if GMateriel.col = G_JOUR then
  	 RechDateGrid(GMateriel)
  else if GMateriel.col = G_RESSOURCE then
  begin
    if CHOIX_RESSOURCE.checked then
    begin
      result := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+GMateriel.cells[GMateriel.col,GMateriel.row],'TYPERESSOURCE='+TYPERESSOURCE.Value + Action);
      if result <> '' then GMateriel.cells[GMateriel.col,GMateriel.row] := result;
    end else
    begin
      result := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+GMateriel.cells[GMateriel.col,GMateriel.row],'TYPERESSOURCE=MAT,OUT' + Action);
      if result <> '' then GMateriel.cells[GMateriel.col,GMateriel.row] := result;
    end;
  end
  else if GMateriel.col = G_CODEARTICLE then
  begin
    if GMateriel.cells[G_RESSOURCE,GMateriel.row] <> '' then
    begin
      stChamps := 'TYPERESSOURCE='+ GetTypeRessource(GMateriel.cells[G_RESSOURCE,GMateriel.row]);
    end else
    begin
      stChamps := 'TYPERESSOURCE=MAT,OUT';
    end;
    stChamps := stChamps+';GA_TYPEARTICLE=PRE';
    Article := AGLLanceFiche('BTP', 'BTPREST_RECH', 'GA_CODEARTICLE='+GMateriel.cells[GMateriel.col,GMateriel.row], '', StChamps);
    if Article <> '' then
    begin
      TOBL.PutValue('BCO_ARTICLE',Article);
      GMateriel.cells[GMateriel.col,GMateriel.row] := Copy(Article,1,18);
    end;
	end
  else if GMateriel.col = G_AFFAIRE0 then
	  AfficheNature(GMateriel)
  else if GMateriel.col = G_AFFAIRE then
  	RechAffaireGrid(GMateriel, TOBL)
  else if GMateriel.col = G_PHASE then
  	RechPhaseGrid(GMateriel, TOBL)
  //else if GMateriel.col = G_FACT then
  //  LookupList (THEdit(Sender),'Sélection du mode','COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="ARE" AND CO_LIBRE="ART"','',true,0);

end;

procedure TOF_BTSAISIECONSO.GMaterielCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(GMateriel,ACol, ARow, Cancel);
  GMateriel.ElipsisButton := ((GMateriel.Col = G_JOUR) and (Calendrier.SelectionStart<>Calendrier.SelectionEnd) ) or
                           (GMateriel.col = G_RESSOURCE) or
                           (GMateriel.col = G_CODEARTICLE) or
                           (GMateriel.col = G_AFFAIRE0) or
                           (GMateriel.col = G_AFFAIRE) or
                           //(GMateriel.col = G_FACT) or
                           (GMateriel.col = G_PHASE);
  if not Cancel then
  	 begin
     AfficheLibelle(TgsRES,Acol,Arow);
     stprevMat := GMateriel.Cells [Acol,Arow];
     SetToolValue (TgsRes,Arow);
  	 end;

end;

procedure TOF_BTSAISIECONSO.GMaterielCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  CellExitGrid(TOBRes, GMateriel, Acol, Arow,TgsRes, StPrevMat,Cancel);

end;

// gestion grille fourniture
procedure TOF_BTSAISIECONSO.GfournitureRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  RowEnterGrid(Ou, Gfourniture, TOBMAT, TgsFourn, Cancel);
end;

procedure TOF_BTSAISIECONSO.GfournitureRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowExitGrid(Ou, Gfourniture, TgsFourn, Cancel);

end;

procedure TOF_BTSAISIECONSO.GFourniturePostDrawCell (ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
	GridsPostDrawCell (Gfourniture,TOBMAT,Acol,Arow,Canvas,Astate);
end;

procedure TOF_BTSAISIECONSO.GfournitureElipsisClick (Sender : TOBject);
var TOBL : TOB;
begin

  TOBL := TOBMAT.detail[Gfourniture.row-1] ;

  if Gfourniture.col = G_JOUR then
		 RechDateGrid(GFourniture)
	else if Gfourniture.col = G_RESSOURCE then
  	 RechRessourceGrid(GFourniture, 'TYPERESSOURCE=SAL')
	else if Gfourniture.col = G_CODEARTICLE then
     RechFraisGrid(Gfourniture, TOBL,'(GA_TYPEARTICLE="MAR") OR (GA_TYPEARTICLE="ARP")')
  else if Gfourniture.col = G_AFFAIRE0 then
	  AfficheNature(Gfourniture)
	else if Gfourniture.col = G_AFFAIRE then
  	RechAffaireGrid(GFourniture, TOBL)
	else if Gfourniture.col = G_PHASE then
  	RechphaseGrid(GFourniture, TOBL);
	//else if Gfourniture.col = G_FACT then
  //  LookupList (THEdit(Sender),'Sélection du mode','COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="ARE" AND CO_LIBRE="ART"','',true,0);

end;

procedure TOF_BTSAISIECONSO.GfournitureCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(Gfourniture,ACol, ARow, Cancel);
  Gfourniture.ElipsisButton := ((Gfourniture.Col = G_JOUR) and (Calendrier.SelectionStart<>Calendrier.SelectionEnd) ) or
                           (Gfourniture.col = G_RESSOURCE) or
                           (Gfourniture.col = G_CODEARTICLE) or
                           (Gfourniture.col = G_AFFAIRE0) or
													 (Gfourniture.col = G_AFFAIRE) or
                           //(Gfourniture.col = G_FACT) or
                           (Gfourniture.col = G_PHASE);
  if not Cancel then
  begin
    AfficheLibelle(TgsFOURN,Acol,Arow);
    stprevFour := Gfourniture.Cells [Acol,Arow];
    SetToolValue (TgsFourn,Arow);
  end;
end;

procedure TOF_BTSAISIECONSO.GfournitureCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  CellExitGrid(TOBMat, GFourniture, Acol, Arow, TgsFourn, StPrevFour,Cancel);

end;

// MO Externes
procedure TOF_BTSAISIECONSO.GMOextRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
 	RowEnterGrid(Ou, GMoExt, TOBMOEXT, TgsMoExt, Cancel);
end;

procedure TOF_BTSAISIECONSO.GMOextRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowExitGrid(Ou, GMoExt, TgsMoExt, Cancel);

end;

procedure TOF_BTSAISIECONSO.GMOextElipsisClick (Sender : TOBject);
var TOBL : TOB;
    stchamps,Article : string;
    result : string;
    Action : String;
begin

	Action := ';ACTION=RECH';

  TOBL := TOBMOEXT.detail[GMOext.row-1] ;
  if GMOext.col = G_JOUR then
  	 RechDateGrid(GMOExt)
  else if GMOext.col = G_RESSOURCE then
  begin
  //
    if CHOIX_RESSOURCE.checked then
    begin
      result := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+GMOext.cells[GMOext.col,GMOext.row],'TYPERESSOURCE='+TYPERESSOURCEEXT.Value + Action);
      if result <> '' then GMOext.cells[GMOext.col,GMOext.row] := result;
    end else
    begin
      result := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+GMOext.cells[GMOext.col,GMOext.row],'TYPERESSOURCE=INT,ST,AUT,LOC' + Action);
      if result <> '' then GMOext.cells[GMOext.col,GMOext.row] := result;
    end;
  //
  end else if GMOext.col = G_CODEARTICLE then
  begin
    if GMOext.cells[G_RESSOURCE,GMOext.row] <> '' then
    begin
      stChamps := 'TYPERESSOURCE='+ GetTypeRessource(GMOext.cells[G_RESSOURCE,GMOext.row]);
    end else
    begin
      stChamps := 'TYPERESSOURCE=INT,ST,AUT,LOC';
    end;
    stChamps := stChamps+';GA_TYPEARTICLE=PRE';
    Article := AGLLanceFiche('BTP', 'BTPREST_RECH', '', '', StChamps);
    if Article <> '' then
    begin
      TOBL.PutValue('BCO_ARTICLE',Article);
      GMOext.cells[GMOext.col,GMOext.row] := Copy(Article,1,18);
    end;
  end else if GMOext.col = G_AFFAIRE0 then
  	AfficheNature(GMOext)
	else if GMOext.col = G_AFFAIRE then
  	RechAffaireGrid(GMOExt, TOBL)
  else if GMOext.col = G_PHASE then
  	rechPhaseGrid(GMOExt, TOBL)
  //else if GMOext.col = G_FACT then
  //  LookupList (THEdit(Sender),'Sélection du mode','COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="ARE" AND CO_LIBRE="ART"','',true,0);

end;

procedure TOF_BTSAISIECONSO.GMOextCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(GMOext,ACol, ARow, Cancel);
  GMOext.ElipsisButton := ((GMOext.Col = G_JOUR) and (Calendrier.SelectionStart<>Calendrier.SelectionEnd) ) or
                           (GMOext.col = G_RESSOURCE) or
                           (GMOext.col = G_CODEARTICLE) or
                           (GMOext.col = G_AFFAIRE0) or
                           (GMOext.col = G_AFFAIRE) or
                           //(GMOext.col = G_FACT) or
                           (GMOext.col = G_PHASE);
  if not Cancel then
  begin
    AfficheLibelle(TgsMoext,Acol,Arow);
    stPrevMoext := GMOext.Cells [Acol,Arow];
    SetToolValue (TgsMoext,Arow);
  end;
end;

procedure TOF_BTSAISIECONSO.GMOextCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  CellExitGrid(TOBMoExt, GMoExt, Acol, Arow,TgsMoext, StPrevMoext,Cancel);

end;

// Recettes Annexes
procedure TOF_BTSAISIECONSO.GrecettesRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	RowEnterGrid(Ou, GRecettes, TOBRecettes, TgsRecettes, Cancel);
end;

procedure TOF_BTSAISIECONSO.GrecettesRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

	RowExitGrid(Ou, GRecettes, TgsRecettes, Cancel);

end;

procedure TOF_BTSAISIECONSO.GrecettesElipsisClick (Sender : TOBject);
var TOBL      : TOB;
    stchamps  : string;
    TheDate   : TDateTime;
    year,Month,day : word;
begin
  TOBL := TOBRecettes.detail[Grecettes.row-1] ;
  if Grecettes.col = G_JOUR then
  begin
    TheDate := SelectDateFromCalendar (calendrier.SelectionStart,Calendrier.SelectionEnd);
    if TheDate <> 0 then
    begin
      DecodeDate (TheDate,year,month,day);
      Grecettes.cells[Grecettes.col,Grecettes.row] := inttostr(day);
      if TheDate > LastDayMoExt Then LastDayMoExt := TheDate;
    end;
  end
  else if Grecettes.col = G_AFFAIRE0 then
  	AfficheNature(GRecettes)
  else if Grecettes.col = G_AFFAIRE then
  begin
    Grecettes.cells[Grecettes.col,Grecettes.row] := GetChantier (Grecettes.cells[Grecettes.col,Grecettes.row]);
  end else if Grecettes.col = G_PHASE then
  begin
    if TOBL.GetValue('BCO_AFFAIRE') = '' then exit;
    stChamps := Grecettes.cells[Grecettes.col,Grecettes.row];
    if SelectionPhase (TOBL.GetValue('BCO_AFFAIRE'),stChamps) then
    begin
       Grecettes.cells[Grecettes.col,Grecettes.row] := stChamps;
    end;
  end else if GRecettes.col = G_NATURE then
  begin
    LookupList (THEdit(Sender),'Sélection de la nature','COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="BNM" AND (CO_CODE="RAN" OR CO_CODE="FAN")','',true,0);
  end;
end;

procedure TOF_BTSAISIECONSO.GrecettesCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(Grecettes,ACol, ARow, Cancel);
  Grecettes.ElipsisButton := ((Grecettes.Col = G_JOUR) and (Calendrier.SelectionStart<>Calendrier.SelectionEnd)) or
                           	 (Grecettes.col = G_NATURE) or
													 	 (Grecettes.col = G_AFFAIRE0) or
                           	 (Grecettes.col = G_AFFAIRE) or
                           	 (Grecettes.col = G_PHASE);
  if not Cancel then
  begin
    AfficheLibelle(TgsRecettes,Acol,Arow);
    stPrevRecettes := Grecettes.Cells [Acol,Arow];
    SetToolValue (TgsRecettes,Arow);
  end;
end;

procedure TOF_BTSAISIECONSO.GrecettesCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  CellExitGrid(TOBRecettes, GRecettes, Acol, Arow,TgsRecettes, StPrevRecettes,Cancel);
end;

//
function TOF_BTSAISIECONSO.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
begin
  result := true;
  if Grille.ColWidths[acol] = 0 then BEGIN result := false; exit; END;

  // MODIF BRL 160609
  // Contrôle accès aux zones :
  // impossible si ligne validée ou provenant d'une pièce sauf pour le libellé ou la phase
	if (ACol <> G_LIBELLE) and (Acol <> G_PHASE) then
  begin
  	if LigneDejaValide (CurrTOB) then Result := False
  	else if LigneFromPieces (CurrTOB) then Result := False;
  end;

end;

procedure TOF_BTSAISIECONSO.ZoneSuivanteOuOk(Grille : THGrid;var ACol, ARow : integer; var Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  ChgLig := (Grille.Row <> ARow);
  ChgSens := False;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol <= Grille.Col)) then Sens := 1;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  while not ZoneAccessible(Grille,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := Grille.RowCount ;
      // ---
      if ((ACol = Grille.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          // Ajout d'une ligne
          break;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.ColCount;
        ChgLig := False;
      end;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

function TOF_BTSAISIECONSO.VerifDate(TOBL : TOB;valeur : string) : boolean;
var TheDate : TDateTime;
    MoisDeb,MoisFin,AnneeDeb,AnneeFin,JourDeb,JourFin,TheJour : word;
begin

  result := false;
  TheJour := StrToInt (valeur);

  DecodeDate (Calendrier.SelectionStart,AnneeDeb,MoisDeb,JourDeb);
  DecodeDate (Calendrier.SelectionEnd,AnneeFin,Moisfin,Jourfin);

  if not IsValidDate (AnneeDeb,Moisdeb,JourDeb) then
  begin
    PgiBox (TraduireMemoire('Date invalide'),ecran.caption);
    exit;
  end;

  TheDate := EncodeDate (AnneeDeb,MoisDeb,TheJour);
  if (TheDate < Calendrier.SelectionStart) then
  begin
    inc(MoisDeb);
    if MoisDeb > 12 then
    begin
      MoisDeb := 1;
      AnneeDeb := AnneeDeb + 1;
    end;
    TheDate := EncodeDate (AnneeDeb,MoisDeb,TheJour);
  end;

  if (TheDate >= Calendrier.SelectionStart) and (TheDate <=Calendrier.SelectionEnd) then
  begin
    if TOBL.GetValue('BCO_DATEMOUV') <> TheDate then
    Begin
      TOBL.PutValue('MODIF', 'X');
      TOBL.PutValue('TYPEMODIF', 'D');
    end;
    TOBL.putValue('JOUR',TheJour);
    TOBL.putValue('BCO_DATEMOUV',TheDate);
    ConstitueDateDebutTravaux(TOBL);
    TOBL.putValue('BCO_SEMAINE',NumSemaine(TheDate));
    TOBL.putValue('BCO_MOIS', moisDeb);
    result := true;
  end else
  begin
    PgiBox (TraduireMemoire('la date doit être comprise entre le ') +
            DateToStr (Calendrier.SelectionStart) + TraduireMemoire(' et le ')+
            DateToStr (Calendrier.SelectionEnd),ecran.caption);
  end;
end;

function TOF_BTSAISIECONSO.VerifRessource(TOBL : TOB;valeur : string) : boolean;
var Bidon    : string;
    NoErreur : Single;
begin
  NoErreur := -1;
  result := false;
  valeur := Trim(UpperCase (valeur));
  if valeur = '' then Exit;
  if TOBL.GetValue('BCO_NATUREMOUV')='MO'  then
  begin
    //FV1 : 25/02/2014 - FS#881 - En saisie d'heures et duplication de ligne, la prestation associée à la ressource est reprise.
    if TOBL.GetString('BCO_ARTICLE') <> '' then
      NoErreur := OKRessource (TOBL,valeur,'"SAL","INT"',Bidon,false,'',false)
    else
      NoErreur := OKRessource (TOBL,valeur,'"SAL","INT"',Bidon,true,'',True);
    if NoErreur = 0 then
    begin
      PGIBox (TraduireMemoire('Main d''oeuvre inexistante'),ecran.caption);
    end
    Else if NoErreur = 2 then
    begin
      PGIBox (TraduireMemoire('Main d''oeuvre Fermée'),ecran.caption);
    end
    Else if NoErreur = 1 then Result := true;
  end
  else if TOBL.GetValue('BCO_NATUREMOUV')='RES' then
  	 result := OKRessourceMAT (TOBL,valeur,true,Bidon)
  else if TOBL.GetValue('BCO_NATUREMOUV')='EXT' then
  begin
     if Trim(Valeur) <> '' then
       	result := OKRessourceExterne (TOBL,valeur,False,Bidon)
     else
        result := true; // on autorise la non saisie
  end
  else if TOBL.GetValue('BCO_NATUREMOUV')='FRS' then
  begin
     NoErreur := OKRessource (TOBL,valeur,'"SAL","INT"',Bidon,False,'',false);
    if NoErreur = 0 then
    begin
      PGIBox (TraduireMemoire('Main d''oeuvre inexistante'),ecran.caption);
    end
    Else if NoErreur = 2 then
    begin
      PGIBox (TraduireMemoire('Main d''oeuvre Fermée'),ecran.caption);
    end
    Else if NoErreur = 1 then Result := True
  end
  else if TOBL.GetValue('BCO_NATUREMOUV')='FOU' then result := true;

  if result then
  begin
	  ChargeResource (Valeur);
    // -----------
    if (OkSaisEquipe) and
       (TOBRessource.GetString('ARS_EQUIPERESS')<>'') and
       (TOBL.GetString('BCO_LINKEQUIPE')='') and
       (Pos(TOBL.GetValue('BCO_NATUREMOUV'),'MO;RES;EXT')>0) and
       (TOBL.GetString('NEW')='X') then
    begin
      TOBL.putValue('BCO_LINKEQUIPE',TOBRessource.GetString('ARS_EQUIPERESS')+';'+DateTimeToStr(Now));
    end;
    // -----------
    //
    if TOBL.GetValue('BCO_RESSOURCE') <> Trim(valeur) then
    Begin
      TOBL.PutValue('MODIF', 'X');
      TOBL.PutValue('TYPEMODIF', 'Q');
    end
    else
    Begin
      if Action = TaCreat then
      begin
        TOBL.PutValue('MODIF', 'X');
        TOBL.PutValue('TYPEMODIF', 'Q');
      end
      else
        TOBL.PutValue('MODIF', '-');
    end;

    TOBL.putValue('BCO_RESSOURCE', trim(valeur));
    result := true;
  end;

end;

function TOF_BTSAISIECONSO.VerifTypeHeure (TOBL: TOB; Valeur : string) : boolean;
begin

	result := false;
  valeur := trim (Uppercase(Valeur));

  result := OKTypeHeure(TOBL,valeur,TOBTypeHeure);

  if not result then
  begin
    PgiBox (TraduireMemoire('Type d''heure inconnu'),ecran.caption);
    exit;
  end;

end;

function TOF_BTSAISIECONSO.VerifArticle(TOBL : TOB; valeur : string) : boolean;
var WithMajPrixRessource : boolean;
begin

  result := false;
  valeur := Trim(UpperCase (Valeur));
  if valeur = '' then exit;
  // Article non encore référencé en TOB
  if TOBL.GetValue('BCO_NATUREMOUV')='MO'  then
  	 Begin
  	 WithMajPrixRessource := true;
  	 result :=  IsPrestationKnown (Valeur,TOBL,false,TgsMO);
     End
  else if TOBL.GetValue('BCO_NATUREMOUV')='RES' then
     Begin
  	 WithMajPrixRessource := false;
  	 result := IsPrestationKnown (Valeur,TOBL,false,Tgsres);
     End
  else if TOBL.GetValue('BCO_NATUREMOUV')='EXT' then
     Begin
  	 WithMajPrixRessource := false;
  	 result := IsPrestationKnown (Valeur,TOBL,false,TgsMOext);
     End
  else if TOBL.GetValue('BCO_NATUREMOUV')='FRS' then
     Begin
  	 WithMajPrixRessource := false;
  	 result := IsFraisKnown(valeur,TOBL,false);
     End
  else if TOBL.GetValue('BCO_NATUREMOUV')='FOU' then
     Begin
  	 WithMajPrixRessource := false;
  	 result := isArticleKnown (valeur, TOBL,FALSE);
     End;

  if not result then
     begin
     PgiBox (TraduireMemoire('Elément inconnu'),ecran.caption);
     exit;
     end;
     //
     if (TOBressource.getString('ARS_RESSOURCE')<> TOBL.getString('BCO_RESSOURCE')) then
     begin
       ChargeResource (TOBL.getString('BCO_RESSOURCE'));
     end;
     //
  SetInfoArticle (TOBL,TOBRessource,TOBL.GetValue('BCO_ARTICLE'),PrestationDefaut,WithMajPrixRessource);

end;

function TOF_BTSAISIECONSO.VerifAffaire(TheMode: TGrilleModeSaisie;TOBL : TOB;valeur : string) : boolean;
var Part0,Part1,Part2,Part3,Avenant, st : string;
		Nat0        : string;
    Etat        : String;
    WithMajPrix : boolean;
    SAVValeur   : String;
begin

  result := false;

  valeur := Trim(Uppercase(Valeur));

  if Valeur = '' then
  begin
    PgiBox (TraduireMemoire('Code Chantier non renseigné'),ecran.caption);
    exit;
  end;

  // Formatage du code affaire dans le cas où l'utilisateur
  // a saisi directement le code sans le type d'affaire et le numero d'avenant
  st := valeur;

	Nat0 := copy(UpperCase (st),1,1);
	if (Nat0 <> 'A') and (Nat0 <> 'W') and (Nat0 <> 'I') then
  begin
    valeur := 'A'+copy(st,1,length(st));
  end;

  valeur := Format_String(valeur,15);
  valeur := valeur + '00';

  Etat := GetChampsAffaire (valeur,'AFF_ETATAFFAIRE');
  //

  //if Etat = '' then
  //begin
  //  PgiBox (TraduireMemoire('Chantier dans un drôle d''état???'),ecran.caption);
  //  exit;
  //end;

  //FV1 : 27/01/2014 - FS#827 - DELABOUDINIERE : Contrôle sur chantier non accepté en saisie de consommations
  //                   FS#921 - DELABOUDINIERE : Revoir les contrôles sur appels et contrats en fonction du code état
  if not ControleAffaire(Valeur, Ecran.caption,'SCO') then exit;

  if IsExisteAffaire (valeur, CodeDomaine, LibAffaire) then
    begin
      //
    	WithMajPrix := true;
      {$IFDEF BTP}
      BTPCodeAffaireDecoupe (valeur,Part0,Part1,Part2,Part3,Avenant,taModif,false);
      {$ELSE}
      CodeAffaireDecoupe (valeur,Part0,Part1,Part2,Part3,Avenant,taModif,false);
      {$ENDIF}
      IF (TOBL.GetString ('BCO_AFFAIRE0') <> Part0) Or
         (TOBL.GetString ('BCO_AFFAIRE1') <> Part1) OR
         (TOBL.GetString ('BCO_AFFAIRE2') <> Part2) Or
         (TOBL.GetString ('BCO_AFFAIRE3') <> Part3) Then
      Begin
        TOBL.PutValue('MODIF', 'X');
        TOBL.PutValue('TYPEMODIF', 'A');
      end;
      TOBL.PutValue ('BCO_AFFAIRE',valeur);
      TOBL.PutValue ('BCO_AFFAIRE0',Part0);
      TOBL.PutValue ('BCO_AFFAIRE1',Part1);
      TOBL.PutValue ('BCO_AFFAIRE2',Part2);
      TOBL.PutValue ('BCO_AFFAIRE3',Part3);
      //
      //FV1 : 28/08/2013- FS#632 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
      TOBL.SetString('DOMAINE', CodeDomaine);

      //FV1 : 09/12/2014 - FS#1342 - SCETEC - saisie de consos matériaux
      //Il n'existe aucune gestion du libellé (sic)
      If (TheMode = tgsFOURN) and (CHOIX_MATERIAUX.Checked) then
      begin
      end
			Else IF (TheMode = TgsMO) and (CHOIX_MO.Checked) then
      begin
        SetInfoRessource (TOBL,TOBRESSOURCE,WithMajPrix,false,PrestationDefaut);
        OKTypeHeure(TOBL,TOBL.GetValue('BCO_TYPEHEURE'),TOBTypeHeure);
      end;

      //FV1 : 26/09/2013 - FS#669 - BAGE : En saisie de consommations sur matériel, ligne non valorisée.
      if (TheMode = TgsRes) and (CHOIX_RESSOURCE.Checked) then
      begin
        if (copy(TOBL.getValue('BCO_AFFAIRE'),1,1)='A') and (getParamSocSecur('SO_BTVALOSCONSO',false)) then
        begin
        	if TOBL.GetValue('BCO_ARTICLE')<> '' then
          begin
          	GetValoPrestation (TOBL);
          end;
        end
        else if (copy(TOBL.getValue('BCO_AFFAIRE'),1,1)='W') and (getParamSocSecur('SO_BTVALOAPPELS',false)) then
        begin
        	if TOBL.GetValue('BCO_ARTICLE')<> '' then
          begin
          	GetPVPrestation (TOBL);
          end;
        end;
        //AlimenteValeurInit (TOBL);
      end;
      //
      result := true;
    end
  else
    begin
      PgiBox (TraduireMemoire('Chantier inconnu'),ecran.caption);
    end;

  if result then
  begin
    if (OkSaisEquipe) and
       (TOBRessource.GetString('ARS_EQUIPERESS')<>'') and
       (TOBL.GetString('BCO_LINKEQUIPE')='') and
       (Pos(TOBL.GetValue('BCO_NATUREMOUV'),'MO;RES;EXT')>0) and
       (TOBL.GetString('NEW')='X') then
    begin
      TOBL.putValue('BCO_LINKEQUIPE',TOBRessource.GetString('ARS_EQUIPERESS')+';'+DateTimeToStr(Now));
    end;
  end;

end;

function TOF_BTSAISIECONSO.VerifNatAffaire(TOBL : TOB;valeur : string) : boolean;
begin

  result := false;
  valeur := Trim(Uppercase(Valeur));

  if valeur <> '' then
  Begin
    if TOBL.GetValue('BCO_AFFAIRE0') <> Valeur then
    Begin
      TOBL.PutValue('MODIF', 'X');
      TOBL.PutValue('TYPEMODIF', 'A');
    end;
    TOBL.PutValue ('BCO_AFFAIRE0',valeur);
    result := true;
  end
  else
  	 begin
     PgiBox (TraduireMemoire('Nature d''Affaire inconnue'),ecran.caption);
  	 end;

end;

function TOF_BTSAISIECONSO.VerifNature(TOBL : TOB;valeur : string) : boolean;
begin
  result := false;
  valeur := Trim(Uppercase(Valeur));
  if (valeur = 'FAN') or (Valeur = 'RAN') then
  begin
  	TOBL.PutValue('BCO_NATUREMOUV',Valeur);
    TOBL.PutValue('LIBELLENATURE',RechDom ('BTNATUREMOUV',TOBL.GetValue('BCO_NATUREMOUV'),false));
    result := true;
  end else
  begin
    PgiBox (TraduireMemoire('Type inconnu'),ecran.caption);
  end;
end;

function TOF_BTSAISIECONSO.VerifPhase(TOBL : TOB;valeur : string) : boolean;
Var CodePhase : String;
    Affaire   : string;
begin

  result := True;

  //FV1 : 29/10/20014 - FS#1287 - DURET : Pbs en saisie de consommations par MO.
  Affaire:= TOBL.GetValue('BCO_AFFAIRE');

  if Affaire = '' then exit;
  if Valeur  = '' then exit;

  if TOBL.GetValue('BCO_PHASETRA') <> Valeur then
  Begin
    TOBL.PutValue('MODIF', 'X');
    TOBL.PutValue('TYPEMODIF', 'P');
  end
  else
    TOBL.PutValue('MODIF', '-');


  if IsExistePhaseAffaire (TOBL.GetValue('BCO_AFFAIRE'),valeur) then
     TOBL.PutValue('BCO_PHASETRA',valeur)
  else
  begin
    result := true;
    PgiBox (TraduireMemoire('Phase inconnue'),ecran.caption);
  end;

end;

function TOF_BTSAISIECONSO.VerifMontant (TOBL : TOB;Thevaleur : string) : boolean;
begin
  result := true;

  if TOBL.GetValue('BCO_MONTANTPR') <> Valeur(Thevaleur) then
  Begin
    TOBL.PutValue('MODIF', 'X');
    TOBL.PutValue('TYPEMODIF', 'X');
  end
  Else
    TOBL.PutValue('MODIF', '-');


  TOBL.PutValue('BCO_MONTANTPR', valeur(TheValeur));

  if TOBL.GetValue('BCO_QUANTITE') > 0 then
  begin
    TOBL.PutValue('BCO_DPR', arrondi(TOBL.GEtValue('BCO_MONTANTPR')/TOBL.GetValue('BCO_QUANTITE'),V_PGI.okdecp));
  end;

  if ActiveMode = TgsRecettes then
  begin
  	TOBL.PutValue('BCO_MONTANTHT',  TOBL.GEtValue('BCO_MONTANTPR'));
  	TOBL.PutValue('BCO_PUHT',       TOBL.GEtValue('BCO_MONTANTPR'));
  	TOBL.PutValue('BCO_MONTANTACH', TOBL.GEtValue('BCO_MONTANTPR'));
  	TOBL.PutValue('BCO_DPA',        TOBL.GEtValue('BCO_MONTANTPR'));
  end;

end;

function TOF_BTSAISIECONSO.VerifPu (TOBL : TOB;Thevaleur : string) : boolean;
var LastprPv : double;
begin
  result := true;
	if TOBL.GetValue ('BCO_DPR') > 0 then
    LastPrPV := TOBL.GetValue ('BCO_PUHT')/TOBL.GetValue ('BCO_DPR')
  else
    LastPrPv := 0;

  if TOBL.GetValue('BCO_DPR') <> Valeur(Thevaleur) then
  Begin
    TOBL.PutValue('MODIF', 'X');
    TOBL.PutValue('TYPEMODIF', 'X');
  end
  Else
    TOBL.PutValue('MODIF', '-');


  TOBL.PutValue('BCO_DPR', valeur(TheValeur));
  calculeLaLigne (TOBL,0,LastprPv);
end;


function TOF_BTSAISIECONSO.VerifQTE(TOBL : TOB;Thevaleur : string) : boolean;
begin
(*
  if (valeur(TheValeur) > 24 ) and (TOBL.GetValue('BCO_NATUREMOUV')='MO') then
  BEGIN
    PgiBox (TraduireMemoire('Nb heures supérieure à 24'),ecran.caption);
    result := false;
    exit;
  END;
*)

  if TOBL.GetValue('BCO_QUANTITE') <> Valeur(Thevaleur) then
  Begin
    TOBL.PutValue('MODIF', 'X');
    TOBL.PutValue('TYPEMODIF', 'Q');
  end
  else
  Begin
    if Action = TaCreat then
    begin
      TOBL.PutValue('MODIF', 'X');
      TOBL.PutValue('TYPEMODIF', 'Q');
    end
    else
      TOBL.PutValue('MODIF', '-');
  end;

  TOBL.PutValue('BCO_QUANTITE',Valeur (TheValeur));
  result := true;

end;

function TOF_BTSAISIECONSO.VerifQualif(TOBL : TOB;valeur : string) : boolean;
begin
  result := true;
end;

function TOF_BTSAISIECONSO.VerifFacturable(TOBL : TOB;valeur : string) : boolean;
begin
  valeur := trim(uppercase(Valeur));
  result := (rechdom('AFTACTIVITEREPRISE',valeur,false,'AND CO_LIBRE="ART"')<>'');
  if result then
  	 begin
     TOBL.PutValue('BCO_FACTURABLE',Valeur);
  	 end;
end;

function TOF_BTSAISIECONSO.ControleLigne (ligne : integer; Mode : TGrilleModeSaisie;var Acol : integer) : boolean;
var TOBL      : TOB;
    TheDate   : TDateTime;
    Phase,Article,Ressource ,Libmessage: string;
    Bidon     : String;
    NoErreur  : Single;
    TypeMvt   : String;
begin

  result := true;
  if Ligne = 0 then exit;

  TOBL := nil;
  //
  if Mode = TgsMo then TOBL := TOBMO.detail[ligne-1] else
  if Mode = TgsMoEXT then TOBL := TOBMOEXT.detail[ligne-1] else
  if Mode = TgsFrs then TOBL := TOBFRS.detail[ligne-1] else
  if Mode = Tgsfourn then TOBL := TOBMAT.detail[ligne-1] else
  if Mode = TgsRES then TOBL := TOBRES.detail[ligne-1] else
  if Mode = TgsRecettes then TOBL := TOBRecettes.detail[ligne-1];
  //
  if TOBL = nil then exit;
  theDate := TOBL.getValue('BCO_DATEMOUV');

  if (TheDate < Calendrier.SelectionStart) or (TheDate > Calendrier.SelectionEnd) then
  begin
    result := false;
    Acol := G_JOUR;
    exit;
  end;
  //
  if (Mode = TgsRecettes) and
  	 (TOBL.GetValue('BCO_NATUREMOUV') <> 'FAN') and
     (TOBL.GetValue('BCO_NATUREMOUV') <> 'RAN') then
  begin
  	Acol := G_NATURE;
  	result := false;
    exit;
  end;
  //
  if (not CHOIX_CHANTIER.Checked) then
  begin
    if not IsExisteAffaire (TOBL.getValue('BCO_AFFAIRE'), CodeDomaine, LibAffaire) then
    begin
      PgiBox (TraduireMemoire('Chantier inconnu'),ecran.caption);
      Acol := G_AFFAIRE;
      result := false;
      exit;
    end;
    Phase := TOBL.GetValue('BCO_PHASETRA');
    if (Phase <> '') and (not IsExistePhaseAffaire (TOBL.GetValue('BCO_AFFAIRE'),Phase)) then
    begin
      PgiBox (TraduireMemoire('Phase inconnu'),ecran.caption);
      Acol := G_PHASE;
      result := false;
      exit;
    end;
  end;

  if (not CHOIX_CONTRAT.Checked) then
  	 begin
     if not IsExisteAffaire (TOBL.getValue('BCO_AFFAIRE'), CodeDomaine,LibAffaire) then
        begin
       	PgiBox (TraduireMemoire('Contrat inconnu'),ecran.caption);
    	  Acol := G_AFFAIRE;
        result := false;
        exit;
        end;
       //FV1 : 28/08/2013- FS#632 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
       TOBL.AddChampSupValeur('DOMAINE', CodeDomaine);
     end;

  if (not CHOIX_APPEL.Checked) then
  	 Begin
     if not IsExisteAffaire (TOBL.getValue('BCO_AFFAIRE'), Codedomaine, LibAffaire) then
 		    begin
    		PgiBox (TraduireMemoire('Appel inconnu'),ecran.caption);
    		Acol := G_AFFAIRE;
      	result := false;
      	exit;
    		end;
       //FV1 : 28/08/2013- FS#632 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
       TOBL.AddChampSupValeur('DOMAINE', CodeDomaine);    
     end;

  //FV1 : 28/08/2013- FS#632 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
  TOBL.SetString('DOMAINE', CodeDomaine);
  // controle de la prestation
  Article := TOBL.getValue('BCO_ARTICLE');

  if TOBL.GetValue('BCO_NATUREMOUV')='MO'  then result := IsPrestationKnown (Article,TOBL,true,TgsMo,False) else
  if TOBL.GetValue('BCO_NATUREMOUV')='EXT' then result := IsPrestationKnown (Article,TOBL,true,TgsMoext,False) else
  if TOBL.GetValue('BCO_NATUREMOUV')='RES' then result := IsPrestationKnown (Article,TOBL,true,Tgsres,False) else
  if TOBL.GetValue('BCO_NATUREMOUV')='FRS' then result := IsFraisKnown(Article,TOBL,true,False) else
  if TOBL.GetValue('BCO_NATUREMOUV')='FOU' then result := isArticleKnown (Article,TOBL,true,False); //FS#1342 - SCETEC - saisie de consos matériaux isArticleKnown (Article,TOBL,True,False)

  if not result then
  begin
  	Libmessage := '';
    if Mode = TgsMo     then LibMessage := 'Prestation'          else
    if Mode = TgsMoEXT  then LibMessage := 'Prestation'          else
    if Mode = TgsFrs    then LibMessage := 'Prestation'          else
    if Mode = Tgsfourn  then LibMessage := 'Article/marchandise' else
    if Mode = TgsRES    then LibMessage := 'Prestation';
    if LibMessage <> '' then PgiBox (TraduireMemoire(LibMessage)+ TraduireMemoire(' inconnu(e)'),ecran.caption);
    Acol := G_CODEARTICLE;
    Exit;
  end;

  // Controle Ressource
  Ressource := TOBL.getValue('BCO_RESSOURCE');
  TypeMvt   := TOBL.GetValue('BCO_NATUREMOUV');

  result := true;

  if TypeMvt ='MO' then
  begin
    NoErreur := OKRessource (NIL,Ressource,'"SAL","INT"',Bidon,false,'',True);
    If Noerreur = 0 then Result := False
  end
  else if TypeMvt='EXT' then
  begin
    if Ressource <> '' then result := OKRessourceExterne (nil,Ressource,false,Bidon)
  end
  else if TypeMvt='RES' then
  begin
    if Ressource <> '' then result := OKRessourceMAT (nil,Ressource,false,Bidon);
  end
  else if TypeMvt='FRS' then
  begin
    NoErreur := OKRessource (nil,Ressource,'"SAL","INT"',Bidon,False,'',True);
    If NoErreur = 0 then Result := False
  end
  else if TypeMvt='FOU' then
    result := true;

  if not result then
  begin
  	Acol := G_RESSOURCE;
    if TypeMvt ='MO' then
    begin
      if NoErreur = 0 then
        PgiBox (TraduireMemoire('Main d''oeuvre inexistante'),ecran.caption)
      else
        PgiBox (TraduireMemoire('Main d''oeuvre Fermée'),ecran.caption);
    end
    else if TypeMvt ='FRS' then
    begin
      if NoErreur = 0 then
        PgiBox (TraduireMemoire('Ressource inexistante'),ecran.caption)
      else
        PgiBox (TraduireMemoire('Ressource Fermée'),ecran.caption);
    end;
    exit;
  end;

  // controle Qte
(*
  if (TOBL.getValue('BCO_QUANTITE') > 24 ) and (TOBL.GetValue('BCO_NATUREMOUV')='MO') then
  BEGIN
    PgiBox (TraduireMemoire('Nb heures supérieure à 24'),ecran.caption);
    result := false;
    exit;
  END;
*)

  //if (OkSaisEquipe) and (TOBL.GetString('BCO_LINKEQUIPE')<>'')then
  //begin
  //  TraiteLigneEquipe (TOBL);
  //end;

end;

function TOF_BTSAISIECONSO.ExisteValorisation(TOBL : TOB) : boolean;
begin
  result := (TOBL.GetValue('BCO_DPA') <> 0) OR
            (TOBL.GetValue('BCO_DPR') <> 0) OR
            (TOBL.GetValue('BCO_PUHT') <> 0);
end;

function TOF_BTSAISIECONSO.LigneVide (Ligne : integer; Mode : TGrilleModeSaisie ) : boolean;
var TOBL : TOB;
begin
  result := true;
  TOBL := nil;

  if      (Mode = TgsMo)        and (TOBMO.detail.count > (ligne-1))        then TOBL := TOBMO.detail[ligne-1]
  else if (Mode = TgsMoEXT)     and (TOBMOEXT.detail.count > (ligne-1))     then TOBL := TOBMOEXT.detail[ligne-1]
  else if (Mode = TgsFrs)       and (TOBFRS.detail.count > (ligne-1))       then TOBL := TOBFRS.detail[ligne-1]
  else if (Mode = Tgsfourn)     and (TOBMAT.detail.count > (ligne-1))       then TOBL := TOBMAT.detail[ligne-1]
  else if (Mode = TgsRES)       and (TOBRES.detail.count > (ligne-1))       then TOBL := TOBRES.detail[ligne-1]
  else if (Mode = TgsRecettes)  and (TOBRecettes.detail.count > (ligne-1))  then TOBL := TOBRecettes.detail[ligne-1];

  if TOBL = nil then exit;

  if TOBL.GetValue('BCO_NATUREMOUV')='MO' then
     begin
     // controle ligne de main d'oeuvre
     if (not CHOIX_CHANTIER.Checked) AND (not CHOIX_CONTRAT.Checked) AND (not CHOIX_APPEL.Checked)then
        begin
        if TOBL.GetValue('BCO_AFFAIRE')<> '' then
           BEGIN
           result := false;
           Exit;
           END;
        if TOBL.GetValue('BCO_PHASETRA')<> '' then
           BEGIN
           result := false;
           Exit;
           END;
        end;
     if (not CHOIX_MO.checked) and (not CHOIX_MATERIAUX.Checked ) then
        begin
					If (TypeSaisie = 'INTERVENTION') and (TOBL.GetValue('BCO_QUANTITE')=0) then
          begin
          	exit; // la on se dit que la ligne est vide (correction pour ligne vide)
          end;

        if (TOBL.getValue('BCO_RESSOURCE') <> '')  then
           BEGIN
           result := false;
           Exit;
           END;
        end;
     if (not CHOIX_MATERIAUX.Checked ) then
        begin
        if (TOBL.getValue('BCO_ARTICLE') <> '') and (TOBL.getValue('BCO_ARTICLE') <> PrestationDefaut) then
           BEGIN
           result := false;
           Exit;
           END;
        end;
     end
  else if TOBL.GetValue('BCO_NATUREMOUV')= 'RES' then
  begin
		if (not CHOIX_CHANTIER.Checked) AND (not CHOIX_CONTRAT.Checked) AND (not CHOIX_APPEL.Checked) then
    begin
      if TOBL.GetValue('BCO_AFFAIRE') <> '' then BEGIN result := false; Exit; END;
      if TOBL.GetValue('BCO_PHASETRA')<> '' then BEGIN result := false; Exit; END;
    end;
    if not CHOIX_RESSOURCE.Checked then
    begin
      if TOBL.getValue('BCO_RESSOURCE') <> '' then BEGIN result := false; Exit; END;
    end;
//    if TOBL.getValue('BCO_ARTICLE') <> '' then BEGIN result := false; Exit; END;
    if (not CHOIX_MATERIAUX.Checked ) then
    begin
      if (TOBL.getValue('BCO_ARTICLE') <> '') and (TOBL.getValue('BCO_ARTICLE') <> PrestationDefaut) then
      BEGIN
        result := false;
        Exit;
      END;
    end;
  end else if TOBL.GetValue('BCO_NATUREMOUV')='EXT' then
  begin
    if (not CHOIX_CHANTIER.Checked) AND (not CHOIX_CONTRAT.Checked) AND (not CHOIX_APPEL.Checked) then    begin
      if TOBL.GetValue('BCO_AFFAIRE')<> '' then BEGIN result := false; Exit; END;
      if TOBL.GetValue('BCO_PHASETRA')<> '' then BEGIN result := false; Exit; END;
    end;
    if not CHOIX_MOEXT.Checked then
    begin
      if TOBL.getValue('BCO_RESSOURCE') <> '' then BEGIN result := false; Exit; END;
    end;
    if TOBL.getValue('BCO_ARTICLE') <> '' then BEGIN result := false; Exit; END;
  end else if TOBL.GetValue('BCO_NATUREMOUV')='FRS' then
  begin
    if (not CHOIX_CHANTIER.Checked) AND (not CHOIX_CONTRAT.Checked) AND (not CHOIX_APPEL.Checked) then    begin
      if TOBL.GetValue('BCO_AFFAIRE')<> '' then BEGIN result := false; Exit; END;
      if TOBL.GetValue('BCO_PHASETRA')<> '' then BEGIN result := false; Exit; END;
    end;
    if TOBL.getValue('BCO_ARTICLE') <> '' then BEGIN result := false; Exit; END;
  end else if TOBL.GetValue('BCO_NATUREMOUV')='FOU' then
  begin
    if (not CHOIX_CHANTIER.Checked) AND (not CHOIX_CONTRAT.Checked) AND (not CHOIX_APPEL.Checked) then    begin
      if TOBL.GetValue('BCO_AFFAIRE')<> '' then BEGIN result := false; Exit; END;
      if TOBL.GetValue('BCO_PHASETRA')<> '' then BEGIN result := false; Exit; END;
    end;
    if not CHOIX_MATERIAUX.Checked then
    begin
      if TOBL.getValue('BCO_ARTICLE') <> '' then BEGIN result := false; Exit; END;
    end;
  end else if (TOBL.GetValue('BCO_NATUREMOUV')='FAN') or (TOBL.GetValue('BCO_NATUREMOUV')='RAN') then
  begin
      if TOBL.getValue('BCO_MONTANTPR') <> 0 then BEGIN result := false; Exit; END;
  end;
  //if ExisteValorisation(TOBL) THEN BEGIN result := false; Exit; END;
  if TOBL.getValue('BCO_QUANTITE') <> 1 THEN BEGIN result := false; Exit; END;
end;

procedure TOF_BTSAISIECONSO.AfficheLibelle(Mode : TGrilleModeSaisie;Acol,Arow : integer);
var TOBL        : TOB;
    TheLibelle  : String;
    LeLibelle   : String;
    restclick   : boolean;
    LibEtat     : String;
    Libtiers    : String;
begin

  RestClick := false;

  TOBL := nil;

  if Mode = TgsMo then
  begin
    if Arow <= TOBMO.detail.count then TOBL := TOBMO.detail[Arow-1];
  end else if Mode = TgsMoEXT then
  begin
    if Arow <= TOBMOEXT.detail.count then TOBL := TOBMOEXT.detail[Arow-1];
  end else if Mode = TgsFrs then
  begin
    if Arow <= TOBFRS.detail.count then  TOBL := TOBFRS.detail[Arow-1]
  end else if Mode = Tgsfourn then
  begin
    if Arow <= TOBMAT.detail.count then TOBL := TOBMAT.detail[Arow-1]
  end else if Mode = TgsRecettes then
  begin
    if Arow <= TOBRecettes.detail.count then TOBL := TOBRecettes.detail[Arow-1]
  end else if Mode = TgsRES then
  begin
    if Arow <= TOBRES.detail.count then TOBL := TOBRES.detail[Arow-1];
  end;

  if TOBL = nil then exit;

  If TOBL.GetVALUE('BCO_AFFAIRE') <> '' then
  begin
    //Libellé de l'affaire
    LeLibelle   := TOBL.GetVALUE('LIBAFFAIRE');
    //Libellé de l'état
    If TOBL.GetValue('BCO_AFFAIRE0') = 'W' then
    Begin
      BChangeEtat.visible := True;
      LeLibelle := 'Appel : ' + LeLibelle;
      LibEtat   := RechDom('BTETATAPPEL',   TOBL.GetValue('ETATAFFAIRE') , False);
    end
    else
    Begin
      BChangeEtat.visible := False;
      if Tobl.GetValue('BCO_AFFAIRE0') = 'I' then
        LeLibelle := 'Contrat : ' + LeLibelle
      else if Tobl.GetValue('BCO_AFFAIRE0') = 'A' then
        LeLibelle := 'Chantier : ' + LeLibelle
      else
        LeLibelle := 'Affaire : ' + LeLibelle;
      //
      LibEtat   := RechDom('AFETATAFFAIRE', TOBL.GetValue('ETATAFFAIRE') , False);
    end;
    //Libellé du tiers
    LibTiers    := TOBL.GetVALUE('TIERS') + ' ' + TOBL.GetVALUE('LIBTIERS');
  end
  else
  begin
    BChangeEtat.visible := False;
    LeLibelle := '';
    LibEtat   := '';
    Libtiers  := '';
  end;

  LLabelHeure.Caption := leLibelle;
  LLabelEtat.Caption  := LibEtat;
  LLabelTiers.Caption := Libtiers;

  if Acol = G_JOUR Then
  	// heuuuu...
    LeLibelle := ''
  Else if ACol = G_NATURE Then
  	leLibelle := TOBL.GetValue('LIBELLENATURE')
  ELSE if Acol = G_RESSOURCE Then
    leLibelle := TOBL.GetValue('LIBELLEMO')
  ELSE if Acol = G_CODEARTICLE Then
  BEGIN
    TheLibelle := GetLibelleMateriaux (TOBL.GetValue('BCO_ARTICLE'));
    if (Mode=TgsMo) or (Mode=TgsRes) or (Mode = TgsMoext) then LeLibelle := 'Prestation : ' + TheLibelle else
    if (Mode=TgsFourn) then LeLibelle := 'Fourniture : ' + TheLibelle else
    if Mode = TgsFrs then LeLibelle := 'Frais : ' + TheLibelle;
  END
  ELSE if Acol = G_AFFAIRE Then
  BEGIN
    GestionColonneAffaire(TOBL);
  END
  ELSE if Acol = G_PHASE Then
  BEGIN
    TheLibelle := GetLibellePhase(TOBL.GetValue('BCO_AFFAIRE'),TOBL.GetValue('BCO_PHASETRA'));
    if TheLibelle <> '' then
      LeLibelle := 'Phase : ' + TheLibelle
    else
      LeLibelle := '';
    LLabelHeure.Caption := leLibelle;
  END
  ELSE if Acol = G_TYPEHEURE Then
  BEGIN
     TheLibelle := TOBL.GetValue('LIBELLETYPE');
     if TheLibelle <> '' then
        LeLibelle := 'Type : ' + TheLibelle
     else
        LeLibelle := '';
     LLabelHeure.Caption := leLibelle;
  END
  ELSE if (Acol = G_QTE) OR (Acol = G_MONTANT) OR (Acol = G_QUALIF) Then
    LLabelHeure.Caption := '';

  if TOBL = nil then exit;

end;

Procedure Tof_BTSAISIECONSO.GestionColonneAffaire(TOBL : TOB);
Var LeLibelle : string;
    LibEtat   : string;
    LibTiers  : String;
    RestClick : Boolean;
begin

  //On efface les informations liées au contrat
  GestionAffichageContrat(false);

  LeLibelle := '';
  LibEtat   := '';
  LibTiers  := '';

  if TOBL.GetValue('BCO_AFFAIRE') = '' then Exit;

  GetLibelleChantier(TOBL);

  if Tobl.GetValue('BCO_AFFAIRE0') = 'W' then
  begin
    GestionAffichageContrat(True);
    contrat.text := TOBL.GetValue('BCO_AFFAIRESAISIE');
    LLIBCONTRAT.Caption := GetLibelleContrat(Contrat.text);

    if TOBL.GetValue('BCO_QTEFACTUREE') <> 0 then
    Begin
      TOBL.putValue('BCO_FACTURABLE', 'A');
    end
    else
      QteFacture.Text := TOBL.GetValue('BCO_QTEFACTUREE');

    if assigned(FACTURABLE.OnClick) then
    begin
      Facturable.onclick := nil;
      RestClick := true;
    end;

    if TOBL.GetValue('BCO_FACTURABLE') = 'A' then
      Facturable.Checked := true
    else if TOBL.GetValue('BCO_FACTURABLE') = 'F' then
      Facturable.State := cbGrayed
    else
      Facturable.Checked := False;

    if restclick then FACTURABLE.OnClick := CocheFact;
  end;

end;

procedure TOF_BTSAISIECONSO.BannulClick(Sender: TObject);
begin
  //
  TOBMo.ClearDetail;
  TOBMoEXT.ClearDetail;
  TOBFRS.ClearDetail;
  TOBMAT.ClearDetail;
  TOBRES.ClearDetail;
  TOBConso.ClearDetail; TOBConsoOld.clearDetail;
  TOBrecettes.ClearDetail;
  TOBLIENEQUIPE.ClearDetail;
  TOBLIENEQUIPE_O.ClearDetail;
  //
  EnteteActive (True);
  EventGrilles(false);
  //
  GHeures.VidePile (false);
  GFrais.videpile (false);
  GMAteriel.videpile (false);
  GFourniture.videpile (false);
  GMOEXT.videpile (false);
  GRecettes.videpile (false);
  PSaisie.visible     := false;
  THRSSAISIE.Visible  := false;
  NBHRSSAISIE.Visible := false;
  //
  BChangeEtat.Visible := False;

end;

procedure TOF_BTSAISIECONSO.PlusWindowClose(Sender: TObject);
begin
  BBlocNote.Down := false;
end;

procedure TOF_BTSAISIECONSO.RetourWindowClose(Sender: TObject);
begin
  BBlocNote1.Down := false;
  BBlocNote2.Down := false;
end;

function TOF_BTSAISIECONSO.ExisteLigneEquipe(TOBL : TOB; LinkEquipe : string) : boolean;
var OneTOB : TOB;
begin

	result := false;

  OneTOB :=  TOBMO.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  while OneTOB <> nil do
  begin
    if (OneTOB <> nil) and (TOBL <> OneTOB ) then BEGIN result := True; Break; end;
  	OneTOB :=  TOBMO.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;
  if result then Exit;

  OneTOB :=  TOBMoExt.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
	while OneTOB <> nil do
  begin
    if (OneTOB <> nil) and (TOBL <> OneTOB ) then BEGIN result := True; Break; end;
  	OneTOB :=  TOBMoExt.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;
  if result then Exit;

  OneTOB :=  TOBRES.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
	while OneTOB <> nil do
  begin
    if (OneTOB <> nil) and (TOBL <> OneTOB ) then BEGIN result := True; Break; end;
  	OneTOB :=  TOBRES.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;

end;

Procedure TOF_BTSAISIECONSO.DeleteLigneEquipeExt(LinkEquipe : string);
var TOBLE : TOB;
begin

  TOBLE :=  TOBLIENEQUIPE.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],True);
	while TOBLE <> nil do
  begin
    //FreeAndNil(TOBLE);
    TOBLE.AddChampSupValeur('SUPPRESSION', 'X');
	  TOBLE :=  TOBLIENEQUIPE.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],True);
  end;

end;

Procedure TOF_BTSAISIECONSO.DeleteLigneEquipe(LinkEquipe : string);
var OneTOB,LocTOB: TOB;
		Mode : TGrilleModeSaisie;
    TheGrid : THGrid;
begin
  OneTOB :=  TOBMO.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
	while OneTOB <> nil do
  begin
    DelLigne (TgsMo,GHeures,TOBMO,OneTOB);
	  OneTOB :=  TOBMO.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;

  OneTOB :=  TOBMoExt.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
	while OneTOB <> nil do
  begin
    DelLigne (TgsMOEXT,GMoext,TOBMoExt,OneTOB);
	  OneTOB :=  TOBMoExt.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;

  OneTOB :=  TOBRES.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
	while OneTOB <> nil do
  begin
    DelLigne (TgsRES,GMAteriel,TOBRES,OneTOB);
	  OneTOB :=  TOBRES.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;

end;

Procedure TOF_BTSAISIECONSO.DelLigne (Mode : TGrilleModeSaisie; TheGrid : Thgrid; LaTOB,OneTOB: TOB);
var Ou : Integer;
begin
  Ou := OneTOB.getIndex+1;
  //
  OneTOB.free;
  TheGrid.DeleteRow (ou);

  if LaTOB.detail.count = 0 then
  begin
    OneTOB := AddNewLigne (LaTOB,Mode);
  	Ou := OneTOB.getIndex+1;
    AfficheLaLigne (OneTOB,TheGrid,stColListe,Ou);
  end
  else if LaTOB.detail.count = 1 then
  begin
    Ou := 1;
  end;

  if LaTOB.detail.count > 0 then TheGrid.RowCount := LaTOB.detail.count +2;
  if TheGrid.RowCount < 2 then TheGrid.rowcount := 3;

  PositionneDansGrid (TheGrid,ou,1);

end;

Procedure TOF_BTSAISIECONSO.DeleteLigne (Mode : TGrilleModeSaisie);
var TOBL,TheTOB : TOB;
    theGrid : THGrid;
    TOBEquipe   : TOB;
    DelEquipe : boolean;
begin
  DelEquipe :=false;
  TheGrid := nil;
  TheTob := nil;

  if Mode = TgsMo then
  begin
    theGrid := GHeures;
    TheTOB := TOBMO;
  end else if Mode = TgsFrs then
  begin
    TheGrid := GFrais;
    TheTOB := TOBFRS;
  end else if Mode = Tgsfourn then
  begin
    TheGrid := GFourniture;
    TheTOB := TOBMAT;
  end else if Mode = TgsMOEXT then
  begin
    TheGrid := GMoext;
    TheTOb := TOBMoExt;
  end else if Mode = TgsRES then
  begin
    TheGrid := GMAteriel;
    TheTOB := TOBRES;
  end else if Mode = TgsREcettes then
  begin
    TheGrid := GRecettes;
    TheTOB := TOBRecettes;
  end;
  if TheGrid = nil then exit;

  if LigneVide (Thegrid.row,Mode) then exit;
  TOBL := TheTOB.Detail[TheGrid.row -1];
  if LigneFromPieces (TOBL)then
  begin
    PGIBox (TraduireMemoire('Impossible : provenance d''un document'),ecran.caption);
    exit;
  end;
  if LigneDejaValide(TOBL) then
  begin
    PGIBox (TraduireMemoire('Impossible : Ligne validée'),ecran.caption);
    exit;
  end;

  if TOBL.GetString('BCO_LINKEQUIPE') <> '' then
  begin
    if (CHOIX_CHANTIER.Checked) or (CHOIX_CONTRAT.Checked) Or (CHOIX_APPEL.Checked) then
      DelEquipe := True
    else
    begin
      TOBEquipe := TOBLIENEQUIPE.FindFIRST(['BCO_LINKEQUIPE'],[TOBL.GetString('BCO_LINKEQUIPE')], True);
      if TOBEquipe = nil then
        DelEquipe := False
      else
        DelEquipe := True;
    end;
  end;

  if delEquipe then
  begin
    If PGIAsk('Désirez-vous supprimer toute l''équipe ?')= Mryes then
    begin
      if (CHOIX_CHANTIER.Checked) OR (CHOIX_CONTRAT.Checked) OR (CHOIX_APPEL.Checked) then
        DeleteLigneEquipe(TOBL.GetString('BCO_LINKEQUIPE'))
      else
      begin
        DeleteLigneEquipeExt(TOBL.GetString('BCO_LINKEQUIPE'));
        DelLigne (Mode,TheGrid,TheTOB,TOBL);
      end;
    end
    else
     DelLigne(Mode,TheGrid,TheTOB,TOBL);
  end
  else
  	DelLigne(Mode,TheGrid,TheTOB,TOBL);
(*
  TOBL.free;
  TheGrid.DeleteRow (TheGrid.row);

  if TheTOB.detail.count = 0 then
  begin
    TOBL := AddNewLigne (TheTOB,Mode);
    AfficheLaLigne (TOBL,TheGrid,stColListe,TheGrid.row);
  end
  else if TheTOB.detail.count = 1 then
  begin
    TheGrid.row := 1;
  end;

  if TheTob.detail.count > 0 then TheGrid.RowCount := TheTob.detail.count +2;
  if TheGrid.RowCount < 2 then TheGrid.rowcount := 3;

  PositionneDansGrid (TheGrid,TheGrid.row,1);
*)
  TheTOB := nil;

end;

procedure TOF_BTSAISIECONSO.BdeleteClick(sender: Tobject);
begin
  DeleteLigne(ActiveMode);
end;

procedure TOF_BTSAISIECONSO.initgrilles;
begin

  // Grilles MO
  GHeures.ColLengths [G_AFFAIRE0] := 1;
  GHeures.ColLengths [G_AFFAIRE] := 17;
  GHeures.ColLengths [G_PHASE] := 35;
  GHeures.ColLengths [G_RESSOURCE] := 17;
  GHeures.ColLengths [G_CODEARTICLE] := 17;

  // Grilles frais
  GFrais.ColLengths [G_AFFAIRE0] := 1;
  GFrais.ColLengths [G_AFFAIRE] := 17;
  GFrais.ColLengths [G_PHASE] := 35;
  GFrais.ColLengths [G_RESSOURCE] := 17;
  GFrais.ColLengths [G_CODEARTICLE] := 17;

  // Grilles materiel
  GMateriel.ColLengths [G_AFFAIRE0] := 1;
  GMAteriel.ColLengths [G_AFFAIRE] := 17;
  GMAteriel.ColLengths [G_PHASE] := 35;
  GMAteriel.ColLengths [G_RESSOURCE] := 17;
  GMAteriel.ColLengths [G_CODEARTICLE] := 17;

  // Grilles fourniture
  GFourniture.ColLengths [G_AFFAIRE0] := 1;
  GFourniture.ColLengths [G_AFFAIRE] := 17;
  GFourniture.ColLengths [G_PHASE] := 35;
  GFourniture.ColLengths [G_RESSOURCE] := 17;
  GFourniture.ColLengths [G_CODEARTICLE] := 17;

  // Grilles Main d'oeuvre externe
  GMOext.ColLengths [G_AFFAIRE0] := 1;
  GMoext.ColLengths [G_AFFAIRE] := 17;
  GMoext.ColLengths [G_PHASE] := 35;
  GMoext.ColLengths [G_RESSOURCE] := 17;
  GMoext.ColLengths [G_CODEARTICLE] := 17;

  // Grilles Recettes et frais annexex
  GRecettes.ColLengths [G_AFFAIRE0] := 1;
  GRecettes.ColLengths [G_AFFAIRE] := 17;
  GRecettes.ColLengths [G_PHASE] := 35;

  // QUALIFIANT DE MESURE
  GFrais.ColLengths [G_QUALIF] := -1;
  GHeures.ColLengths [G_QUALIF] := -1;
  GMateriel.ColLengths [G_QUALIF] := -1;
  GFourniture.ColLengths [G_QUALIF] := -1;
  GMOExt.ColLengths [G_QUALIF] := -1;

end;

procedure TOF_BTSAISIECONSO.GridElipsisClick(Sender : Tobject);
begin
  if ActiveMode = TgsMo then GheuresElipsisClick (sender) else
  if ActiveMode = TgsFrs then GFraisElipsisClick (Sender)  else
  if ActiveMode = Tgsfourn then GfournitureElipsisClick (Sender) else
  if ActiveMode = TgsMOEXT then GMOextElipsisClick (Sender) else
  if ActiveMode = TgsRecettes then GRecettesElipsisClick (Sender) else
  if ActiveMode = TgsRES then GMaterielElipsisClick (Sender) ;
end;

procedure TOF_BTSAISIECONSO.GridKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_DELETE: if (Shift = [ssCtrl]) then
    begin
      DeleteLigne(ActiveMode);
      Key := 0;
    end;
    VK_F5: if (Shift = []) then
    begin
      GridElipsisClick(Sender);
      Key := 0;
    end;
    VK_F2: if (Shift = []) then
    begin
      dupliqueDonnee;
      Key := 0;
    end;
    VK_RETURN : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
    VK_ESCAPE:
    BEGIN
    	Key := 0;
      BannulClick (self);
    END;
  end;
end;

procedure TOF_BTSAISIECONSO.dupliqueDonnee;

  procedure DupliqueChamps (TOBLP,TOBL : TOB ; LesChamps : string);
  var TheChamps,UnChamp: string;
  begin
    TheChamps := LesChamps;
    repeat
      UnChamp := READTOKENST (TheChamps);
      if unChamp <> '' then
      begin
        TOBL.PutValue (UnChamp,TOBLP.GetValue(UnChamp));
      end;
    until unchamp = '';
  end;

var TOBL,TOBLP,MaTOB : TOB;
    Arow,Acol : integer;
    lesChamps : string;
    LaGrid : THGrid;
begin
  if ActiveMode = TgsMo then
  BEGIN
    laGrid := Gheures;
    MaTOB := TOBMO;
  END else if ActiveMode = TgsFrs then
  BEGIN
    laGrid := GFrais;
    MaTOB := TOBFRS;
  END else if ActiveMode = Tgsfourn then
  BEGIN
    laGrid := GFourniture;
    MaTOB := TOBMAT;
  END else if ActiveMode = TgsMOEXT then
  BEGIN
    laGrid := GMoext;
    MaTOB := TOBMOEXT;
  END else if ActiveMode = TgsRES then
  BEGIN
    laGrid := GMAteriel;
    MaTOB := TOBRES;
  END else if ActiveMode = TgsRecettes then
  BEGIN
    laGrid := GRecettes;
    MaTOB := TOBRecettes;
  END;

  Arow := LaGrid.Row; if Arow < 2 then exit;
  ACol := LaGrid.Col;
  TOBL := MaTob.detail[Arow-1];
  TOBLP := MaTob.detail[Arow-2];

  lesChamps := '';
  if Acol = G_JOUR Then lesChamps := 'BCO_DATEMOUV' else
  if Acol = G_NATURE Then lesChamps := 'BCO_NATUREMOUV' else
  if Acol = G_RESSOURCE Then lesChamps := 'BCO_RESSOURCE' else
  if Acol = G_CODEARTICLE Then lesChamps := 'BCO_CODEARTICLE' else
  if Acol = G_AFFAIRE0 Then lesChamps := 'BCO_AFFAIRE0' else
  if Acol = G_AFFAIRE Then lesChamps := 'BCO_AFFAIRE' else
  if Acol = G_PHASE Then lesChamps := 'BCO_PHASETRA' else
  if Acol = G_QTE Then lesChamps := 'BCO_QUANTITE' else
  if Acol = G_PU Then lesChamps := 'BCO_DPR' else
  if Acol = G_MONTANT Then lesChamps := 'BCO_MONTANTPR' else
  if Acol = G_TYPEHEURE Then lesChamps := 'BCO_TYPEHEURE' else
  //if Acol = G_FACT Then lesChamps := 'BCO_FACTURABLE';
  if LesChamps = '' then exit;

  DupliqueChamps (TOBLP,TOBL,LesChamps);
  AfficheLaLigne (TOBL,lagrid,stColListe,Arow);
  TOBL.SetString('BCO_LINKEQUIPE','');
  if (TOBressource.getString('ARS_RESSOURCE')<> TOBL.getString('BCO_RESSOURCE')) then
  begin
    ChargeResource (TOBL.getString('BCO_RESSOURCE'));
  end;
  if (OkSaisEquipe) and
     (TOBRessource.GetString('ARS_EQUIPERESS')<>'') and
     (Pos(TOBL.GetValue('BCO_NATUREMOUV'),'MO;RES;EXT')>0) and
     (TOBL.GetString('NEW')='X') then
  begin
    TOBL.putValue('BCO_LINKEQUIPE',TOBRessource.GetString('ARS_EQUIPERESS')+';'+DateTimeToStr(Now));
    TraiteLigneEquipe (TOBL);
  end;


end;

procedure TOF_BTSAISIECONSO.BDuplicClick(Sender: Tobject);
var TOBL,TOBLP,MaTOB : TOB;
    Arow,Acol : integer;
    lesChamps : string;
    LaGrid : THGrid;
    TheJour : TDateTime;
    TheMois,TheSemaine : integer;
    year,Month,Day : word;
begin

  if ActiveMode = TgsMo then
  BEGIN
    laGrid := Gheures;
    MaTOB  := TOBMO;
  END else if ActiveMode = TgsFrs then
  BEGIN
    laGrid := GFrais;
    MaTOB  := TOBFRS;
  END else if ActiveMode = Tgsfourn then
  BEGIN
    laGrid := GFourniture;
    MaTOB  := TOBMAT;
  END else if ActiveMode = TgsMOEXT then
  BEGIN
    laGrid := GMoext;
    MaTOB  := TOBMOEXT;
  END else if ActiveMode = TgsRES then
  BEGIN
    laGrid := GMAteriel;
    MaTOB  := TOBRES;
  END else if ActiveMode = TgsRecettes then
  BEGIN
    laGrid := Grecettes;
    MaTOB  := TOBRecettes;
  END;

  Arow := LaGrid.Row;

  if Arow < 2 then exit;

  ACol := LaGrid.Col;

  TOBL        := MaTob.detail[Arow-1];
  TOBLP       := MaTob.detail[Arow-2];
  
  TheJour     := TOBL.GetValue('BCO_DATEMOUV');
  TheMois     := TOBL.GetValue('BCO_MOIS');
  TheSemaine  := TOBL.GetValue('BCO_SEMAINE');

  TOBL.Dupliquer (TOBLP,true,true);
  TOBL.putValue('BCO_DATEMOUV',TheJour);
  ConstitueDateDebutTravaux(TOBL);
  TOBL.putValue('BCO_MOIS',TheMois);
  TOBL.putValue('BCO_SEMAINE',TheSemaine);
  DecodeDate (TOBL.GetValue('BCO_DATEMOUV'),year,Month,Day);
  TOBL.PutValue ('JOUR', InttoStr(Day));
  TOBL.PutValue('NEW','X');

  AfficheLaLigne (TOBL,lagrid,stColListe,Arow);

  QTEFACTURE.Text := TOBL.GetValue('BCO_QTEFACTUREE');
  contrat.text := TOBL.GetValue('BCO_AFFAIRESAISIE');
  FAMILLETAXE1.OnChange := nil;
  if TypeSaisie = 'INTERVENTION' then SetControlText('FAMILLETAXE1',TOBL.GetString('BCO_FAMILLETAXE1'));
  FAMILLETAXE1.OnChange := FAMILLETAXEChange;
  LLIBCONTRAT.Caption :=GetLibelleContrat(Contrat.text);
  if TOBL.GetValue('BCO_FACTURABLE') = 'A' then
     Facturable.Checked := true
  else
	   Facturable.Checked := False;

  TOBL.SetString('BCO_LINKEQUIPE','');
  TOBL.SetInteger('BCO_PAIENUMFIC',0);
  if (TOBressource.getString('ARS_RESSOURCE')<> TOBL.getString('BCO_RESSOURCE')) then
  begin
    ChargeResource (TOBL.getString('BCO_RESSOURCE'));
  end;
  if (OkSaisEquipe) and
     (TOBRessource.GetString('ARS_EQUIPERESS')<>'') and
     (Pos(TOBL.GetValue('BCO_NATUREMOUV'),'MO;RES;EXT')>0) and
     (TOBL.GetString('NEW')='X') then
  begin
    TOBL.putValue('BCO_LINKEQUIPE',TOBRessource.GetString('ARS_EQUIPERESS')+';'+DateTimeToStr(Now));
    TraiteLigneEquipe (TOBL);
  end;
end;

//FV1 : 21/02/2014 - FS#896 - BAGE : en saisie de consommations, dupliquer la saisie d'une semaine d'un salarié.
procedure TOF_BTSAISIECONSO.BDuplicSemaineClick(Sender: Tobject);
var StSQL         : String;
    StSQLEnCrs    : String;
    StWhere       : String;
    Titre         : String;
    QQ            : TQuery;
    TOBEnCours    : TOB;
    TOBDuplic     : TOB;
    TOBL          : TOB;
    DateDebut     : TDateTime;
    DateDeFin     : TDateTime;
    DatePrecDebut : TDateTime;
    DatePrecFin   : TDateTime;
    DateTrait     : TdateTime;
    Indice        : Integer;
    PreviousWeek  : Integer;
    IndDate       : integer;
    Diff          : Extended;
    TheMois       : Integer;
    TheSemaine    : integer;
    year          : Word;
    Month         : Word;
    Day           : word;
begin
  //
  TheSemaine:= Calendrier.weekselect;
  DateDebut := Calendrier.SelectionStart;
  DateDeFin := Calendrier.SelectionEnd;
  //
  Diff := (DateDeFin - DateDebut)+1;

  Titre := 'Récupération Semaine précédente';

  if Datedebut = DateDeFin then
  begin
    PGIBOX('Veuillez sélectionner une semaine complète !', Titre);
    Exit;
  end;

  if (Diff > 7) then
  begin
    PGIBOX('Vous avez sélectionné plus d''une semaine. Veuillez sélectionner une semaine complète !', Titre);
    Exit;
  end;

  if (Diff < 7) then
  begin
    PGIBOX('Vous avez sélectionné moins d''une semaine, Veuillez sélectionner une semaine complète !', Titre);
    Exit;
  end;
  //
  if  TheSemaine = 0 then
  begin
    PGIBOX('Récupération impossible aucune semaine selectionnée !', Titre);
    exit;
  end;

  if (choix_MO.Checked) then
  begin
    if Trim(MAINDOEUVRE.text) = '' then
    begin
      PGIBOX('La main d''oeuvre doit être renseignée !', Titre);
      MAINDOEUVRE.SetFocus;
      Exit;
    end;
  end;

  if (CHOIX_MATERIAUX.Checked) then
  Begin
    if Trim(MATERIAUX.text) = '' then
    begin
      PGIBOX('Le matériaux doit être renseignée !', Titre);
      MATERIAUX.SetFocus;
      Exit;
    end;
  end;         

  //Chargement des dates Sélectionnées moins 7 jours
  DatePrecDebut := DateDebut - 7;
  DatePrecFin   := DateDeFin - 7;
  //
  StSQL := ChargementRequeteConso(DatePrecDebut, DatePrecFin);
  //
  QQ := OpenSql (StSQL,true,-1,'',true);

  if QQ.eof then
  begin
    PGIBOX('Aucune saisie sur la semaine précédente !', Titre);
    Ferme(QQ);
    Exit;
  end
  else
  begin
    TOBDuplic := TOB.Create ('LES CONSO', nil,-1);
    TOBCONSO.LoadDetailDB ('CONSOMMATIONS','','',QQ,FALSE);
  end;

  //Vérification si aucunes saisies sur la semaine sélectionnée
  StSQLEnCrs := ChargementRequeteConso(DateDebut, DateDeFin);
  If existeSQL(StSQLEnCrs) then
  begin
    if PGIAsk('Des consommations existent pour la semaine sélectionnée.' + CHR(10) +
              'Attention vos consommations précédentes sur cette semaine seront supprimées !' + CHR(10) +
              'Voulez-vous les Remplacer ?', Titre)=MrNo then
      Exit
    else
    begin
      //Suppression des enregistrement déjà existant
      StSQL := 'DELETE FROM CONSOMMATIONS WHERE ';
      //FV1 : 30/11/2015 - FS#1810 - TEAM RESEAUX : ajout duplication idem MO dans onglet matériels
      if CHOIX_RESSOURCE.Checked then
      begin
        StSQL := StSQL + '(BCO_NATUREMOUV="RES")';
        StSQL := StSQL + ' AND BCO_RESSOURCE="' + Trim(RESSOURCES.text) + '"';
      end
      else
      begin
        StSQL := StSQL + '(BCO_NATUREMOUV="MO" OR BCO_NATUREMOUV="FRS" OR BCO_NATUREMOUV="FOU")';
        StSQL := StSQL + ' AND BCO_RESSOURCE="' + Trim(MAINDOEUVRE.text) + '"';
      end;
      //
      if DateDeFin <> DateDebut then
      begin
       StSQL := StSQL + ' AND BCO_DATEMOUV >= "'+USDATETIME (DateDebut)+'"';
       StSQL := StSQL + ' AND BCO_DATEMOUV <= "'+USDATETIME (DateDeFin)+'"';
      end
      else
       StSQL := StSQL + ' AND BCO_DATEMOUV = "' +USDATETIME (DateDebut)+'"';

      StSQL := StSQL + ' AND BCO_TRANSFORME="-"'; // ne reprends pas les transformé
      StSQL := StSQL + ' AND BCO_TRAITEVENTE<>"X"'; // ne reprends pas les passés en livraisons
      //
      ExecuteSQL(StSQL);
    end;
  end;

  //Chargement des consos de la semaine précédente pour duplication sur semaine sélectionnée
  if TOBCONSO.Detail.count > 0 then
  begin
    TOBDuplic.Dupliquer(TOBCONSO,true,true);
    For indice := 0 to TOBDuplic.detail.count - 1 do
    begin
      TOBL := TOBDuplic.detail[Indice];
      DateTrait := TOBL.GetDateTime('BCO_DATEMOUV');
      DateTrait := DateTrait + 7;
      TheSemaine := WeekOfTheYear(DateTrait);
      DecodeDate(DateTrait, year, Month, Day);
      TOBL.PutValue('BCO_DATEMOUV', DateTrait);
      ConstitueDateDebutTravaux(TOBL);
      TOBL.PutValue('BCO_MOIS', Month);
      TOBL.PutValue('BCO_SEMAINE', TheSemaine);
      TOBL.SetInteger('BCO_PAIENUMFIC',0);
    end;
    //Mise à jour des consos de la semaines sélectionnées
    TOBDuplic.SetAllModifie(True);
    TobDuplic.InsertOrUpdateDB;
    //
  end;

  Ferme(QQ);

  TOBMo.ClearDetail;
  TOBMoEXT.ClearDetail;
  TOBFRS.ClearDetail;
  TOBMAT.ClearDetail;
  TOBRES.ClearDetail;
  TOBRecettes.ClearDetail;

  GHeures.VidePile (false);
  GFrais.videpile (false);
  GMAteriel.videpile (false);
  GFourniture.videpile (false);
  GMOEXT.videpile (false);
  Grecettes.VidePile (false);

  //Affichages des consos de la semaine sélectionnées
  BChercheClick(Self);

  FreeAndNil(TobDuplic);

end;

function TOF_BTSAISIECONSO.ChargementRequeteConso(DateDebut, DateDeFin : Tdatetime) : String;
Begin

  Result := '';

  //génération de la requete de recherche des conso en fontion de la date de début et la date de fin tout onglets confondus
  Result :='SELECT * FROM CONSOMMATIONS WHERE ';

  if CHOIX_RESSOURCE.Checked then
  begin
    //FV1 : 30/11/2015 - FS#1810 - TEAM RESEAUX : ajout duplication idem MO dans onglet matériels
    Result := Result + '(BCO_NATUREMOUV="RES") ';
    Result := Result + ' AND BCO_ARTICLE="' + Trim(Ressources.text) + '"';
  end
  else
  begin
    Result := Result + '(BCO_NATUREMOUV="MO" OR BCO_NATUREMOUV="FRS" OR BCO_NATUREMOUV="FOU")';
    Result := Result + ' AND BCO_RESSOURCE="' + Trim(MAINDOEUVRE.text) + '"';
  end;

  if DateDeFin <> DateDebut then
     begin
     Result := Result + ' AND BCO_DATEMOUV >= "'+USDATETIME (DateDebut)+'"';
     Result := Result + ' AND BCO_DATEMOUV <= "'+USDATETIME (DateDeFin)+'"';
     end
  else
     Result := Result + ' AND BCO_DATEMOUV = "' +USDATETIME (DateDebut)+'"';

  Result := Result + ' AND BCO_TRANSFORME="-"';   // ne reprends pas les transformé
  Result := Result + ' AND BCO_TRAITEVENTE<>"X"'; // ne reprends pas les passés en livraisons
  Result := Result + ' ORDER BY BCO_DATEMOUV';

end;


procedure TOF_BTSAISIECONSO.ChargeResource(CodeResource: string);
var Sql : STring;
		QQ : Tquery;
begin
  Sql := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+CodeResource+'"';
  QQ := OpenSQL (Sql,true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBRessource.SelectDB ('',QQ,true);
  end;
  ferme (QQ);
end;

procedure TOF_BTSAISIECONSO.GrecettesGetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
begin

  if ACol < Grecettes.FixedCols then Exit;
  if Arow < Grecettes.FixedRows then Exit;
  if Arow > TObrecettes.detail.count then exit;

  TOBL := TOBRecettes.detail[ARow-1];

  if TOBL = nil then Exit;

  if TOBL.GetValue('BCO_NATUREMOUV') = 'FAN' then Canvas.Font.Color := clRed;

end;

procedure TOF_BTSAISIECONSO.RecupTypeHeure;
var req : String;
	 QQ :TQuery;
    UneTOB : TOB;
begin

  Req := 'SELECT CC_CODE,CC_LIBELLE,CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="ATH"';

  QQ := OpenSql (Req,True,-1,'',true);

  if not QQ.eof Then TOBTYPEHEURE.LoadDetailDB ('CHOIXCOD','','',QQ,false,true);

  ferme(QQ);

end;

procedure TOF_BTSAISIECONSO.PositionneValeurInit(TOBConso: TOB);
var coef : double;
		leType,LaValeur ,Sql: string;
    QQ : TQuery;
    result : Boolean;
begin

  coef := 1;
  result := false;
  leType := TOBConso.GetValue('BCO_TYPEHEURE');

  if LeType <> '' then
     begin
     LATOB := TOBTypeHeure.findFirst(['CC_CODE'],[LeType],true);
     if LaTOB = nil then
        begin
        Sql := 'SELECT CC_ABREGE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="ATH" AND CC_CODE="'+LeType+'"';
        QQ := OpenSql (SQl,true,-1,'',true);
        if not QQ.eof then
           begin
           LaTOB := TOB.Create ('CHOIXCOD',TOBTypeheure,-1);
           LaTOB.selectdb ('',QQ);
           result := true;
           end;
        ferme (QQ);
        end
    else
        result := true;
    end
  else
    result := true;

  if result and (LeType <> '') then
     begin
     LaValeur :=  LaTOB.GetValue ('CC_ABREGE');
     if (LaValeur <> '') and (IsNumeric (LaValeur)) then
        begin
        Coef := 1 + VALEUR(laValeur)/100;
        end;
     end;

  if (result) and (Coef > 0) then
     begin
  	  TOBConso.PutValue('_PA_INIT',Arrondi(TOBConso.GetValue('BCO_DPA')/Coef,V_PGI.okdecP));
  	  TOBConso.PutValue('_PR_INIT',Arrondi(TOBConso.GetValue('BCO_DPR')/Coef,V_PGI.okdecP));
  	  TOBConso.PutValue('_PV_INIT',Arrondi(TOBConso.GetValue('BCO_PUHT')/Coef,V_PGI.okdecP));
     if LeType <> '' then
        TOBConso.PutValue('LIBELLETYPE',TOBConso.GetValue('CC_LIBELLE'))
     else
        TOBConso.PutValue('LIBELLETYPE','');
     end;

  if TOBConso.getValue ('BCO_RESSOURCE') <> '' then
     begin
  	  Sql := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+TOBConso.getValue ('BCO_RESSOURCE')+'"';
     QQ := OpenSql (SQL,true,-1,'',true);
     if not QQ.eof then TOBConso.putValue('LIBELLEMO',QQ.findField('ARS_LIBELLE').asString);
     ferme (QQ);
     end;

  TOBConso.PutValue('LIBELLENATURE',RechDom ('BTNATUREMOUV',TOBConso.GetValue('BCO_NATUREMOUV'),false));

end;

procedure TOF_BTSAISIECONSO.PositionnePrevValeur (TheGrid: THGrid;Acol,Arow : integer);
begin
  if      TheGrid = GHeures     then stprevHeure    := TheGrid.Cells [Acol,Arow]
  else if TheGrid = GFrais      then stPrevFrs      := TheGrid.Cells [Acol,Arow]
  else if TheGrid = GMateriel   then stprevMat      := TheGrid.Cells [Acol,Arow]
  else if TheGrid = Gfourniture then stprevFour     := TheGrid.Cells [Acol,Arow]
  else if TheGrid = GMOext      then stPrevMoext    := TheGrid.Cells [Acol,Arow]
  else if TheGrid = Grecettes   then stPrevRecettes := TheGrid.Cells [Acol,Arow];
end;

//FV1 - 01/07/2015 : FS#1568 - VIROT : Message en saisie conso par main d'oeuvre
procedure TOF_BTSAISIECONSO.PositionneDansGrid (TheGrid : THGrid;Arow,Acol : integer;Cancel : Boolean = False);
begin

  TheGrid.CacheEdit;
  TheGrid.SynEnabled := false;
  TheGrid.OnRowEnter (self,Arow,Cancel,false);
  TheGrid.OnCellEnter (self,Acol,Arow,Cancel);
  if not cancel then
  begin
    TheGrid.row := Arow;
    TheGrid.col := Acol;
    PositionnePrevValeur (TheGrid,Acol,Arow);
  end else
  begin
    PositionnePrevValeur (TheGrid,Acol,Arow);
  end;
  TheGrid.SynEnabled := True;
  TheGrid.ShowEditor;
end;

procedure TOF_BTSAISIECONSO.GridsPostDrawCell(grid : Thgrid; TOBATrait : TOB ; ACol, ARow: Integer;
  																						Canvas: TCanvas; AState: TGridDrawState);
var Arect : Trect ;
  	Fix: boolean;
    TOBL : TOB;
    PosG : integer;
    chaine : string;
begin

  if grid.RowHeights[ARow] <= 0 then Exit;
//  if ARow > grid.TopRow + grid.VisibleRowCount - 1 then Exit;
  if (Arow < grid.fixedRows) or (Acol < grid.fixedcols) then exit;
  if (Arow < 1) or (Arow > TOBATrait.detail.count) then exit;

	TOBL := TOBATrait.detail[Arow-1]; if TOBL = nil then Exit;

  ARect := grid.CellRect(ACol, ARow);
  Fix := ((ACol < grid.FixedCols) or (ARow < grid.FixedRows));
  grid.Canvas.Pen.Style := psSolid;
  grid.Canvas.Pen.Color := clgray;

  if Acol = G_RESSOURCE then
     begin
     grid.Canvas.Brush.Style := BsSolid;
     Canvas.FillRect(ARect);
     Chaine := TOBL.GetValue('LIBELLEMO');
     PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
     canvas.TextOut (PosG,ARect.top + 1,Chaine);
     end
  else if Acol = G_NATURE then
  	 begin
     grid.Canvas.Brush.Style := BsSolid;
     Canvas.FillRect(ARect);
     Chaine := TOBL.GetValue('LIBELLENATURE');
     PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
     canvas.TextOut (PosG,ARect.top + 1,Chaine);
  	 end
  else if Acol = G_AFFAIRE0 then
     begin
     grid.Canvas.Brush.Style := BsSolid;
     Canvas.FillRect(ARect);
     Chaine := TOBL.GetValue('BCO_AFFAIRE0');
     PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
     canvas.TextOut (PosG,ARect.top + 1,Chaine);
     end
  else if Acol = G_AFFAIRE then
     begin
     grid.Canvas.Brush.Style := BsSolid;
     Canvas.FillRect(ARect);
     Chaine := BTPCodeAffaireAffiche (TOBL.GetValue('BCO_AFFAIRE'));
     PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
     canvas.TextOut (PosG,ARect.top + 1,Chaine);
     end;

end;

procedure TOF_BTSAISIECONSO.GheuresPostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
begin
	GridsPostDrawCell (Gheures,TOBMO,Acol,Arow,Canvas,Astate);
end;

procedure TOF_BTSAISIECONSO.GFraisPostDrawCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
	GridsPostDrawCell (Gfrais,TOBFRS,Acol,Arow,Canvas,Astate);
end;

procedure TOF_BTSAISIECONSO.GMOextPostDrawCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
	GridsPostDrawCell (GMOExt,TOBMOEXT,Acol,Arow,Canvas,Astate);
end;

procedure TOF_BTSAISIECONSO.GMaterielPostDrawCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
	GridsPostDrawCell (GMateriel,TOBRES,Acol,Arow,Canvas,Astate);
end;

procedure TOF_BTSAISIECONSO.GRecettesPostDrawCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
	GridsPostDrawCell (GRecettes,TOBRecettes,Acol,Arow,Canvas,Astate);
end;

function TOF_BTSAISIECONSO.Controlgrids: boolean;
var indice,Acol : integer;
begin

	result := true;

	// controle de la grille des Heures;
  for indice := 0 to TOBMO.detail.count -1 do
  begin
  	if LigneVide (Indice+1,TgsMo) then continue;
  	if not ControleLigne (Indice+1,TgsMo,Acol) then
    begin
      result := false;
      // Positionnement sur l'onglet
      PgTemps.ActivePage := Theures;
      PositionneDansGrid (Gheures,Indice+1,Acol, not Result);
      exit;
    end;
  end;

	// controle de la grille des Frais;
  for indice := 0 to TOBFRS.detail.count -1 do
  begin
  	if LigneVide (Indice+1,TgsFRS) then continue;
  	if not ControleLigne (Indice+1,TgsFRS,Acol) then
    begin
    	result := false;
      // Positionnement sur l'onglet
      PgTemps.ActivePage := TFrais;
			PositionneDansGrid (GFrais,Indice+1,Acol, not Result);
      exit;
    end;
  end;

	// controle de la grille des Main d'oeuvres externes;
  for indice := 0 to TOBMOExt.detail.count -1 do
  begin
  	if LigneVide (Indice+1,TgsMoext) then continue;
  	if not ControleLigne (Indice+1,TgsMoext,Acol) then
    begin
    	result := false;
      // Positionnement sur l'onglet
      PgTemps.ActivePage := TMOExt;
			PositionneDansGrid (GMoext,Indice+1,Acol, not Result);
      exit;
    end;
  end;

	// controle de la grille des Materiels;
  for indice := 0 to TOBRES.detail.count -1 do
  begin
  	if LigneVide (Indice+1,TgsRES) then continue;
  	if not ControleLigne (Indice+1,TgsRES,Acol) then
    begin
    	result := false;
      // Positionnement sur l'onglet
      PgTemps.ActivePage := TMATERIELS;
			PositionneDansGrid (GMAteriel,Indice+1,Acol, not Result);
      exit;
    end;
  end;

	// controle de la grille des Fourniture;
  for indice := 0 to TOBMAT.detail.count -1 do
  begin
  	if LigneVide (Indice+1,TgsFourn) then continue;
  	if not ControleLigne (Indice+1,TgsFourn,Acol) then
    begin
    	result := false;
      // Positionnement sur l'onglet
      PgTemps.ActivePage := TFOURNITURES;
			PositionneDansGrid (GFourniture,Indice+1,Acol, not Result);
      exit;
    end;
  end;

	// controle de la grille des Frais recttes annexes;
  for indice := 0 to TOBRECETTES.detail.count -1 do
  begin
  	if LigneVide (Indice+1,TgsRecettes) then continue;
  	if not ControleLigne (Indice+1,TgsRecettes,Acol) then
    begin
    	result := false;
      // Positionnement sur l'onglet
      PgTemps.ActivePage := TRECETTES;
			PositionneDansGrid (GRecettes,Indice+1,Acol, not Result);
    end;
  end;

end;

procedure TOF_BTSAISIECONSO.MAINDOEUVREElipsisClick(Sender: TOBject);
var result : string;
		Action : String;
begin

	Action := ';ACTION=RECH';

  result := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+MAINDOEUVRE.text+';ARS_FONCTION1='+FONCTION.value,'TYPERESSOURCE=SAL,INT' + Action);
  if Result <> '' then MAINDOEUVRE.Text := result;
end;

function TOF_BTSAISIECONSO.GetcumulTemps: double;
var Indice : integer;
		TOBL : TOB;
begin
	result := 0;
	For Indice := 0 to TOBMO.Detail.count -1 do
  begin
  	TOBL := TOBMO.detail[Indice];
    result := result + TOBL.GetValue('BCO_QUANTITE');
  end;
end;

procedure TOF_BTSAISIECONSO.OnExitPartieAffaire(Sender: Tobject);
var iErr : integer;
begin

    if (CHOIX_CHANTIER.checked) or (CHOIX_CONTRAT.checked ) then
    begin
      if (CH_CHANTIER1.text = '') and (CH_CHANTIER2.text = '') and (CH_CHANTIER3.text = '') then exit;
      CH_CHANTIER.Text := DechargeCleAffaire(CH_CHANTIER0, CH_CHANTIER1, CH_CHANTIER2, CH_CHANTIER3, CH_AVENANT, '', TaCreat, False, True, false, Ierr);
  	end;

    //Rechargement clé appel si modification manuel dans entete saisie (modif FV 16/01/2008)
    if (CHOIX_APPEL.checked) then
      CH_CHANTIER.Text := DechargeCleAppel(CH_CHANTIER0.text, CH_CHANTIER1.text, CH_CHANTIER2.text, CH_CHANTIER3.text, CH_AVENANT.text, '', Tamodif, True, True, false, Ierr);
    //
    //FV1 - 28/12/20015 : FS#1832 - CODRIS : Pb en saisie du code chantier
    ChargeAffaire;

end;

//FV1 - 28/12/20015 : FS#1832 - CODRIS : Pb en saisie du code chantier
Procedure TOF_BTSAISIECONSO.ChargeAffaire;
Var CodeAffaire : string;
    Affaire0    : string;
    Affaire1    : string;
    Affaire2    : string;
    Affaire3    : string;
    Avenant     : string;
    StSQL       : string;
    QQ          : TQuery;
    NbAff       : Integer;
begin

  CodeAffaire := CH_CHANTIER.Text;
  //
  Affaire0    := CH_CHANTIER0.Text;
  Affaire1    := CH_CHANTIER1.Text;
  Affaire2    := CH_CHANTIER2.Text;
  Affaire3    := CH_CHANTIER3.Text;
  Avenant     := CH_AVENANT.Text;


  CodeAffaire:=trim(Copy (CodeAffaire,1,15));   // supression de l'avenant

  If (CodeAffaire = '') Then Exit;

  StSQL := 'SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE0="' + Affaire0 + '"';

  if Affaire0  = ''   then
  begin
    if (CHOIX_CHANTIER.checked) then
      Affaire0 := 'A'
    else if (CHOIX_CONTRAT.checked ) then
      Affaire0 := 'I'
    else if (CHOIX_APPEL.checked ) then
      Affaire0 := 'W'
    else
      Affaire0 := 'A';
    //
    StSQL := StSQL + ' AND AFF_AFFAIRE0="'+  Affaire0 + '"';
  end;

  if Affaire1 <> '' then StSQL := StSQL + ' AND AFF_AFFAIRE1="'+  Affaire1 + '"';
  if Affaire2 <> '' then StSQL := StSQL + ' AND AFF_AFFAIRE2="'+  Affaire2 + '"';
  if Affaire3 <> '' then StSQL := StSQL + ' AND AFF_AFFAIRE3="'+  Affaire3 + '"';
  if Avenant  <> '' then StSQL := StSQL + ' AND AFF_AVENANT="' +  Avenant  + '"';

  QQ := OpenSQL(StSQL,True,2) ;

  // si plus d'une affaire stop le fetch sur les enreg du query ...
  While Not(QQ.EOF) And (NbAff<2) do
  Begin
    Inc(NbAff);
    QQ.Next;
  End;

  If (NbAff = 1) then
  Begin
    QQ.First;
    CodeAffaire := QQ.FindField('AFF_AFFAIRE').AsString;
    BTPCodeAffaireDecoupe(CodeAffaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant, Action, false);
  End;

  // Si aucune affaire correspondante
  If (NbAff = 0) then
  Begin
    PGIBox('Aucune affaire correspondante', TitreHalley);
    Affaire1  := '';
    Affaire2  := '';
    Affaire3  := '';
    Avenant   := '00';
    if CHOIX_CHANTIER.checked     then
      CodeAffaire := 'A' + Affaire1
    Else if CHOIX_CONTRAT.checked then
      CodeAffaire := 'I' + Affaire1
    Else if CHOIX_APPEL.checked   then
      CodeAffaire := 'W' + Affaire1
    else
      CodeAffaire := 'A' + Affaire1;
  End;

  CH_CHANTIER.Text  := CodeAffaire;
  CH_CHANTIER0.Text := Affaire0;
  CH_CHANTIER1.Text := Affaire1;
  CH_CHANTIER2.Text := Affaire2;
  CH_CHANTIER3.Text := Affaire3;
  CH_AVENANT.Text   := Avenant;

  Ferme(QQ);

end;

Procedure TOF_BTSAISIECONSO.GestionAffichageContrat(VraiFaux : Boolean);
Begin

  if VraiFaux then
     begin
     PHEURES.top := 252;
     PHEURES.Refresh;
     end
  else
  	  begin
     PHEURES.top := 233;
     PHEURES.Refresh;
     end;
  //
  LCONTRAT.visible      := VraiFaux;
  //
  Contrat.Visible       := VraiFaux;
  //
  LLibContrat.visible   := VraiFaux;
  //
  LQteFacture.Visible   := VraiFaux;
  //
  QteFacture.visible    := VraiFaux;
  //
  Facturable.visible    := VraiFaux;
  //
  BBlocnote2.Visible    := VraiFaux;
  //
  BEffaceCha.Visible    := VraiFaux;

  FAMILLETAXE1.Visible  := VraiFaux;

end;

procedure TOF_BTSAISIECONSO.ContratClick(Sender: TObject);
var req 			 : string;
	  NumContrat : THEdit;
    LibContrat : THLabel;
    TOBL       : Tob;
begin

  if (GetActiveTabSheet('PGTEMPS').Name = 'THEURES') then
     TOBL := TOBMO.detail[GHeures.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFRAIS') then
     TOBL := TOBFRS.detail[GFrais.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMATERIELS') then
     TOBL := TOBRES.detail[GMateriel.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFOURNITURES') then
     TOBL := TOBMAT.detail[GFourniture.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMOEXT') then
     TOBL := TOBMOEXT.detail[GMOEXT.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TRECETTES') then
     TOBL := TOBRecettes.detail[GRecettes.row-1] ;

  Req := ' AFF_STATUTAFFAIRE="INT"';

  if Codetiers.text = '' then
  begin
    IF TOBL.GetVALUE('TIERS') <> '' then Req := Req + ' AND AFF_TIERS="' + TOBL.GetVALUE('TIERS') + '"';
  end
  else
    Req := Req + ' AND AFF_TIERS="' + CodeTiers.text + '"';

  CONTRAT.Plus := req;

  LookupCombo(Contrat);

  //Lecture du contract sélectionné et affichage des informations
  TOBL.putvalue('BCO_AFFAIRESAISIE', Contrat.Text );
  LLibContrat.caption := GetLibelleContrat(Contrat.text);

end;

procedure TOF_BTSAISIECONSO.QTEFACTUREExit(Sender: TObject);
Var TOBL     : TOB;
Begin
  //
  if (GetActiveTabSheet('PGTEMPS').Name = 'THEURES') and (TOBMO.detail.count > 0) then
     TOBL := TOBMO.detail[GHeures.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFRAIS') and (TOBFRS.detail.count > 0) then
     TOBL := TOBFRS.detail[GFrais.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMATERIELS') and (TOBRES.detail.count > 0)then
     TOBL := TOBRES.detail[GMateriel.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFOURNITURES') and (TOBMAT.detail.count > 0) then
  	 TOBL := TOBMAT.detail[GFourniture.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMOEXT') and (TOBMOEXT.detail.count > 0) then
  	 TOBL := TOBMOEXT.detail[GMOEXT.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TRECETTES') and (TOBRecettes.detail.count > 0) then
  	 TOBL := TOBRecettes.detail[GRecettes.row-1];
  //
  If TOBL <> Nil then TOBL.putvalue('BCO_QTEFACTUREE', Valeur(QTEFACTURE.text));
  //
end;

procedure TOF_BTSAISIECONSO.CocheFact(Sender: TObject);
var TOBL 		: Tob;
    Grille  : THGrid;
    Cancel	: Boolean;
begin
  //
  if (GetActiveTabSheet('PGTEMPS').Name = 'THEURES') then
     Begin
     TOBL := TOBMO.detail[GHeures.row-1] ;
     Grille := Gheures;
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFRAIS') then
     Begin
     TOBL := TOBFRS.detail[GFrais.row-1] ;
     Grille := GFrais;
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMATERIELS') then
     Begin
     TOBL := TOBRES.detail[GMateriel.row-1] ;
     Grille := GMateriel;
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFOURNITURES') then
     Begin
  	 TOBL := TOBMAT.detail[GFourniture.row-1] ;
     Grille := GFourniture;
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMOEXT') then
     Begin
  	 TOBL := TOBMOEXT.detail[GMOEXT.row-1] ;
     Grille := GMOEXT;
     end
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TRECETTES') then
     Begin
  	 TOBL := TOBRecettes.detail[GRecettes.row-1] ;
     Grille := GRecettes;
     end;

  if Facturable.Checked then
     Begin
	   TOBL.putvalue('BCO_FACTURABLE', 'A');
//  	 Grille.cells[G_FACT,Grille.row] := 'A';
     End
  Else
     Begin
	   TOBL.putvalue('BCO_FACTURABLE', 'N');
//  	 Grille.cells[G_FACT,Grille.row] := 'N';
     end;

end;

Procedure TOF_BTSAISIECONSO.RechPrestationGrid(Grille : THGrid; TOBL: TOB);
Var  stchamps : String;
     Article	: string;
begin

	if GMateriel.cells[G_RESSOURCE,GMateriel.row] <> '' then
      stchamps := 'TYPERESSOURCE='+ GetTypeRessource(Grille.cells[G_RESSOURCE,Grille.row])
  else
      stchamps := 'TYPERESSOURCE=SAL,INT';

  stchamps := stchamps + ';GA_TYPEARTICLE=PRE';
  Article := AGLLanceFiche('BTP', 'BTPREST_RECH', 'GA_CODEARTICLE='+Grille.cells[Grille.col,Grille.row] , '', stchamps);

  if Article <> '' then
     begin
     TOBL.PutValue('BCO_ARTICLE',Article);
     Grille.cells[Grille.col,Grille.row] := Copy(Article,1,18);
     end;

end;

Procedure TOF_BTSAISIECONSO.AfficheNature(Grille : THGrid);
var ThetypeC : THEdit;
Begin

  TheTYpeC := ThEdit.create (Grille);
  TheTypeC.Parent := Grille;
  TheTypeC.DataType := 'BTAFFAIREPART0';
  TheTYpeC.Text := Grille.cells[Grille.col,Grille.row];

  LookupCombo (TheTypec);

	Grille.cells[Grille.col,Grille.row] := TheTypec.text;

  if TheTypec.Text= 'W' then
  	GestionAffichageContrat(True)
  else
  	GestionAffichageContrat(False);

  TheTypeC.free;

end;

Procedure TOF_BTSAISIECONSO.RechAffaireGrid(Grille: THGrid; TOBL: Tob);
var   Result  : string;
      TheMode : TGrilleModeSaisie;
      Nat0    : String;
      ChangeStatut : Boolean;
Begin
  ChangeStatut := (CHOIX_MO.Checked) or (CHOIX_RESSOURCE.Checked) or (CHOIX_MATERIAUX.Checked) or (CHOIX_MOEXT.Checked);
  //FV1 : 29/10/2014 - FS#1287 - DURET : Pbs en saisie de consommations par MO.
  //On charge le type d'heure (sic)
  //Result := Grille.cells[Grille.col-1,Grille.row] + ';';
  Result := Copy(grille.cells[Grille.col,Grille.row],1,1) + ';';
  //On charge la main d'oeuvre (Spécial APPEL)
  Result := Result + MAINDOEUVRE.Text + ';';
  //On Charge l'affaire (oui c'est ça)
  Result := Result + grille.cells[Grille.col,Grille.row];
  //On appel le mul de recherche affaire... Mais ça marche pas :(
  Result := GetChantier (Result,ChangeStatut);

  if result = '' then exit;

  //FV1 : 27/01/2014 - FS#827 - DELABOUDINIERE : Contrôle sur chantier non accepté en saisie de consommations
  //                   FS#921 - DELABOUDINIERE : Revoir les contrôles sur appels et contrats en fonction du code état
  //if not ControleAffaire(Result, Ecran.caption,'SCO') then exit;

  Grille.cells[Grille.col,Grille.row] := result;
  TOBL.PutValue('BCO_AFFAIRE',result);

	//Lecture du libellé de l'affaire
  if grille.Name = 'GHEURES' then
     TheMode := TgsMO
  Else if grille.Name = 'GFRAIS' then
	  TheMode := Tgsfrs
  Else if grille.Name = 'GMATERIEL' then
	  TheMode := TgsRes
  Else if grille.Name = 'GFOURNITURE' then
	  TheMode := TgsFourn
  Else if grille.Name = 'GMOEXT' then
	  TheMode := TgsMoExt
  Else if grille.Name = 'GRECETTES' then
	  TheMode := TgsRecettes;

  AfficheLibelle(TheMode,Grille.col,Grille.row);

end;

Procedure TOF_BTSAISIECONSO.RechPhaseGrid(Grille : THGrid; TOBL: TOB);
var   Result : string;
Begin

  if TOBL.GetValue('BCO_AFFAIRE') = '' then exit;

  Result := Grille.cells[Grille.col,Grille.row];

  if SelectionPhase(TOBL.GetValue('BCO_AFFAIRE'),Result) then
     Grille.cells[Grille.col,Grille.row] := Result;
end;

procedure TOF_BTSAISIECONSO.RechRessourceGrid(Grille: THGrid; StArg : String);
Var result : String;
		Action : String;
begin

	Action := ';ACTION=RECH';

  Result := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+Grille.cells[Grille.col,Grille.row],StArg + Action);

  if Result <> '' then Grille.cells[Grille.col,Grille.row] := Result;

end;

procedure TOF_BTSAISIECONSO.RechDateGrid(Grille: THGrid);
Var TheDate : TDateTime;
    year		: Word;
    Month		: Word;
    Day 		: Word;
begin

  TheDate := SelectDateFromCalendar(calendrier.SelectionStart,Calendrier.SelectionEnd);

  if TheDate <> 0 then
  	 begin
     DecodeDate (TheDate,year,month,day);
     Grille.cells[Grille.col,Grille.row] := inttostr(day);
     if TheDate > LastDayMo then LastDayMo := TheDate;
     end;

end;

procedure TOF_BTSAISIECONSO.RechFraisGrid(Grille: THGrid; TOBL: TOB; StArg: String);
Var  stchamps : String;
     Article	: string;
begin

  stChamps := 'XX_WHERE=GA_CODEARTICLE LIKE "'+Grille.cells[Grille.col,Grille.row]+'%" AND ' +StArg;
	stChamps := stChamps + ';FIXEDTYPEART';
  Article := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps);

  if Article <> '' then
     begin
     TOBL.PutValue('BCO_ARTICLE',Article);
     Grille.cells[Grille.col,Grille.row] := Copy(Article,1,18);
     end;

end;

procedure TOF_BTSAISIECONSO.ControleGrille(TheMode: TGrilleModeSaisie;Grille: THGrid; TOBL: TOB; Acol, Arow: Integer; var Cancel: Boolean);
Var ZoneGrille : Variant;
begin

    ZoneGrille := Grille.cells[Acol,Arow];

    if Acol = G_JOUR Then
       Begin
       if not VerifDate(TOBL,ZoneGrille) then cancel := true
       end
    Else if Acol = G_RESSOURCE Then
       Begin
         if ((TheMode = TgsRES) or (TheMode = TgsMoExt))  and (ZoneGrille='') then exit; // on autorise de ne pas saisir 
       if not VerifRessource(TOBL,ZoneGrille) then cancel := true
       end
    Else if Acol = G_TYPEHEURE Then
       Begin
       if not VerifTypeHeure(TOBL,ZoneGrille) then cancel := true
       end
    Else if Acol = G_CODEARTICLE Then
       Begin
    	 if not VerifArticle(TOBL,ZoneGrille) then cancel := true
       End
    Else if Acol = G_LIBELLE Then
       Begin
       TOBL.Putvalue('BCO_LIBELLE',ZoneGrille)
       end
    Else if Acol = G_AFFAIRE0 Then
       Begin
    	 if not VerifNatAffaire(TOBL,ZoneGrille) then cancel := true
       End
    Else if Acol = G_AFFAIRE Then
    	 Begin
         if not VerifAffaire(TheMode,TOBL,ZoneGrille) then
         begin
           cancel := true;
           ZoneGrille := '';
           Grille.cells[Acol,Arow] := ZoneGrille;
           TOBL.PutValue ('BCO_AFFAIRE',ZoneGrille);
           TOBL.PutValue ('BCO_AFFAIRE0','');
           TOBL.PutValue ('BCO_AFFAIRE1','');
           TOBL.PutValue ('BCO_AFFAIRE2','');
           TOBL.PutValue ('BCO_AFFAIRE3','');
         end
       end
    Else if Acol = G_PHASE Then
       Begin
		   if not VerifPhase(TOBL,ZoneGrille) then cancel := true
       End
    Else if Acol = G_QTE Then
       Begin
	     if not VerifQTE(TOBL,ZoneGrille) then cancel := true
       end
    Else if Acol = G_PU Then
       Begin
    	 if not VerifPu(TOBL,ZoneGrille) then cancel := true
       End
    Else if Acol = G_MONTANT Then
       Begin
    	 if not VerifMontant(TOBL,ZoneGrille) then cancel := true
       end
    Else if Acol = G_QUALIF Then
       Begin
       if not VerifQualif(TOBL,ZoneGrille) then cancel := true
    //   end
    //Else if Acol = G_FACT Then
    //   Begin
    //   if not VerifFacturable(TOBL,ZoneGrille) then cancel := true
       end;

  If TOBL.GetString('MODIF') = 'X' Then
  begin
    if (OkSaisEquipe) and (TOBL.GetString('BCO_LINKEQUIPE')<>'')then
    begin
      TraiteLigneEquipe (TOBL);
    end;
  End;

end;

procedure TOF_BTSAISIECONSO.RechArticleGrid(Grille: THGrid; TOBL: TOB);
begin

end;

Procedure TOF_BTSAISIECONSO.ControleChangeOnglet(Grille:THGrid; Var Cancel : Boolean);
Var Arow	: integer;
    Acol 	: integer;
Begin

    Grille.CacheEdit;

    Arow := Grille.row;
    ACol := 1;

    if Grille.Name = 'GHEURES' then
       Begin
    	 GheuresRowEnter (self,Grille.row,cancel,false);
       GheuresCellEnter (self,Acol,Arow,cancel);
       end
    Else if Grille.Name = 'GFRAIS' then
       Begin
		   GFraisRowEnter (self,Grille.row,cancel,false);
	     GFraisCellEnter (self,Acol,Arow,cancel);
       end
  	Else if Grille.Name = 'GMATERIEL' then
       Begin
       GMAterielRowEnter (self,Grille.row,cancel,false);
  	   GMAterielCellEnter (self,Acol,Arow,cancel);
       end
    Else if Grille.Name = 'GFOURNITURE' then
       Begin
       GFournitureRowEnter (self,Grille.row,cancel,false);
  	   GFournitureCellEnter (self,Acol,Arow,cancel);
       end
    Else if Grille.Name = 'GMOEXT' then
    	 Begin
	     GMOEXTRowEnter (self,Grille.row,cancel,false);
  	   GMOEXTCellEnter (self,Acol,Arow,cancel);
       end
    Else if Grille.Name = 'GRECETTES' then
       Begin
	     GRecettesRowEnter (self,Grille.row,cancel,false);
  	   GRecettesCellEnter (self,Acol,Arow,cancel);
       end;

    Grille.row := ARow;
    Grille.col := Acol;

    ShowToolWindow(BBLocNote.name, BBLocNote.Down);

    Grille.ShowEditor;

end;

procedure TOF_BTSAISIECONSO.RowEnterGrid(Ou : Integer; Grille:THGrid; TOBGrille : Tob; TheMode : TGrilleModeSaisie; Var Cancel : Boolean);
Var TobLigne : Tob;
    Affaire  : String;
    TypeAff  : String;
begin

  if (Ou > 1) and (LigneVide (Ou-1,TheMode)) then
     BEGIN
     cancel := true;
     Exit;
     END;

  FACTURABLE.OnClick := nil;
  if (Ou >= Grille.rowCount -1) then
     Grille.rowCount := Grille.rowCount +1;

  if Ou > TOBGrille.Detail.count then
     begin
     TOBLigne := AddNewLigne (TOBGrille,TheMode);
     AfficheLigne (Grille,TOBLigne,Ou-1);
     end;

  TOBLigne := TOBGrille.detail[Ou-1];

  Affaire := TOBLigne.GetValue('BCO_AFFAIRE');
  TypeAff := TOBLigne.GetValue('BCO_AFFAIRE0');

  //chargement des zone propre aux appels
  if TypeAff = 'W' then
  Begin
     GestionAffichageContrat(True);
     ChargementFichierOLE('APP', 'REA', Affaire);
	   if TOBLigne.GetValue('BCO_AFFAIRESAISIE') <> '' then
     Begin
        Contrat.text := TOBLigne.GetValue('BCO_AFFAIRESAISIE');
	      LLibcontrat.Caption := GetLibelleContrat(Contrat.text );
     end;

	   QteFacture.Text := TOBLigne.GetValue('BCO_QTEFACTUREE');
	   if TOBLigne.GetValue('BCO_FACTURABLE') = 'F' then
	   begin
        Facturable.State := cbGrayed;
        CONTRAT.enabled := False;
        QTEFACTURE.enabled := False;
        FACTURABLE.enabled := False;
	   end
     else
     begin
        CONTRAT.enabled := True;
        QTEFACTURE.enabled := True;
        FACTURABLE.enabled := True;
        if TOBLigne.GetValue('BCO_FACTURABLE') = 'A' then
	         Facturable.Checked := true
	      else
	         Facturable.Checked := False;
     end;

     FAMILLETAXE1.OnChange := nil;

     if TypeSaisie = 'INTERVENTION' then
     begin
      if TypeSaisie = 'INTERVENTION' then
      SetControlText('FAMILLETAXE1',TobLigne.GetString('BCO_FAMILLETAXE1'));
      FAMILLETAXE1.OnChange := FAMILLETAXEChange;
     end;
     
  end else if TypeAff = 'I' then
  Begin
	   GestionAffichageContrat(False);
	   Facturable.visible := True;
     Facturable.Checked := False;
  end else if TypeAff = 'A' then
     GestionAffichageContrat(False);

  SetToolValue (TheMode,Ou);
  ActiveMode := TheMode;
  FACTURABLE.OnClick := CocheFact;

end;

Procedure TOF_BTSAISIECONSO.RowExitGrid(Ou : Integer; Grille:THGrid; TheMode : TGrilleModeSaisie;  Var Cancel : Boolean);
var Acol : integer;
Begin

  // Controle de la ligne saisie
  if Action = taConsult then Exit;

  // si vide pas besoin de controler le contenu
  if LigneVide (Ou,TheMode) then exit;

  if not ControleLigne (Ou,TheMode,Acol) then
  BEGIN
  	cancel := true;
    PositionneDansGrid (Grille,Ou,Acol, Cancel);
  END ELSE
  BEGIN
    if TheMode = TgsMo then NBHRSSAISIE.value := GetcumulTemps;
  END;

end;

Procedure TOF_BTSAISIECONSO.CellExitGrid(TOBL : TOB; Grille:THGrid; Acol, Arow : Integer; TheMode : TGrilleModeSaisie; var StPrev : String; Var Cancel : Boolean);
Var TOBLigne : TOB;
    ZoneGrille : Variant;
begin

  if Action = taConsult then Exit;

  // pas de controle si remonté et ligne vide
  if (Grille.row < Arow) and (LigneVide (Arow,TheMode))  then exit;

  ZoneGrille := Grille.cells[Acol,Arow];

  if (stPrev <> ZoneGrille) or (ZoneGrille = '') then
  begin
    TOBLigne := TOBL.detail[Arow-1];
    if LigneDejaValide(TOBLigne) and (Acol <> G_LIBELLE) and (Acol <> G_PHASE) then // modif BRL 150609 : on autorise tjs la modif du libelle et de la phase
    	AfficheLaLigne (TOBLigne,Grille,stColListe,Arow);
    if (TheMode = TgsRes) or (TheMode = TgsFourn ) or (TheMode = TgsMoExt) or (TheMode = TgsRecettes) then
    	if LigneFromPieces(TOBLigne) and (Acol <> G_LIBELLE) and (Acol <> G_PHASE) then AfficheLaLigne (TOBLigne,Grille,stColListe,Arow); // modif BRL 150609 : on autorise tjs la modif du libelle et de la phase
    //Modif FV pour passage qté grille en qté à facturer qualque soit l'onglet de saisie (19062007)
    if ACol = G_QTE then
    Begin
      QTEFACTURE.Text := ZoneGrille;
      TOBLigne.putvalue('BCO_QTEFACTUREE', QTEFACTURE.text);
    end;

    if TheMode = TgsMO then
    begin
    end
    else if TheMode = TgsFourn then
    Begin
      //traitement de récup du prix de l'article
      if (Acol = G_AFFAIRE0) then AfficheNature(Grille);
      if (Acol = G_AFFAIRE) and (TOBLigne.getValue('PRIXRECUP') = '-') Then
      Begin
        SetInfoArticle (TOBLigne,nil,TOBLigne.GetValue('BCO_ARTICLE'),PrestationDefaut,false);
        TOBLigne.PutValue('PRIXRECUP','X');
      end;
    end
    else if TheMode = TgsRecettes then
    Begin
      if Acol = G_JOUR Then
      begin
      	if not VerifDate(TOBLigne,ZoneGrille) then cancel := true;
      end else if Acol = G_NATURE Then
      begin
      	if not VerifNature(TOBLigne,ZoneGrille) then cancel := true;
      end else if Acol = G_AFFAIRE0 Then
      begin
      	if not VerifNatAffaire(TOBLigne,ZoneGrille) then cancel := true;
      end Else if Acol = G_AFFAIRE Then
      begin
        if not VerifAffaire(TheMode,TOBL,ZoneGrille) then
         begin
           cancel := true;
           ZoneGrille := '';
           Grille.cells[Acol,Arow] := ZoneGrille;
         end
      end Else if Acol = G_LIBELLE Then
      begin
      	TOBL.Putvalue('BCO_LIBELLE',ZoneGrille);
      end Else if Acol = G_PHASE Then
      begin
      	if not VerifPhase(TOBLigne,ZoneGrille) then cancel := true;
      end else if Acol = G_MONTANT Then
      begin
      	if not VerifMontant(TOBLigne,ZoneGrille) then cancel := true;
      end;
    end;
    //
		ControleGrille(TheMode,Grille, TOBLigne, Acol, Arow, Cancel);
    //
    if not cancel then
    begin
      //FV1 : 28/08/2013- FS#632 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions

      if (Acol = G_RESSOURCE) Or (Acol = G_CODEARTICLE) Or (Acol = G_AFFAIRE) or (Acol=G_TYPEHEURE) then
      begin
        If ((TOBLigne.GetString('BCO_AFFAIRE0') = 'W') Or (TypeSaisie = 'INTERVENTION')) then
        begin
        	AppliqueCoeff(TOBLigne);
        end;
      end;
      AfficheLibelle(TheMode,Acol,Arow);
      CalculeAndAfficheLaLigne(Grille,stColListe,Arow,TOBL);
      if (TypeSaisie = 'INTERVENTION') then
      begin
        FAMILLETAXE1.OnChange := nil;
        SetControltext('FAMILLETAXE1',TOBLigne.GetValue('BCO_FAMILLETAXE1'));
        FAMILLETAXE1.OnChange := FAMILLETAXEChange;
      end;
    end;
    stPrev := ZoneGrille;
  end;

end;

function TOF_BTSAISIECONSO.GetLibelleContrat (Contrat : string) : string;
var QQ : TQuery;
    Sql : String;
Begin

  result := '';

  if Contrat = '' then
  Begin
    Result := '';
    exit;
  end;

	Sql := 'SELECT AFF_LIBELLE, AFF_TOTALHT FROM AFFAIRE WHERE AFF_AFFAIRE="'+ Contrat +'"';
  QQ := OpenSql (SQL,true,-1,'',true);

  if not QQ.eof then
  	 Result := QQ.findfield('AFF_LIBELLE').asString + ' (' + QQ.findfield('AFF_TOTALHT').AsString + ')';

  ferme (QQ);

end;

procedure TOF_BTSAISIECONSO.ValideLesAppel(DateInt: TDateTime; TOBDet: TOB);
var indice : integer;
    Affaire : string;
    TOBLC   : TOB;
begin

  for Indice := 0 to TOBDet.detail.count - 1 do
  begin
    TOBLC := TOBdet.detail[Indice];
    if TOBLC.GetValue('BCO_AFFAIRE0') <> 'W' then continue;
    Affaire := TOBLC.GetValue('BCO_AFFAIRE');
    If TOBLC.GetValue('CHANGEETAT') = 'X' then
      ValideDesAppel(TOBLC)
    else
      ValideDesAppel(DateInt,Affaire);
  end;

  ValideSansCloture := false;

end;

procedure TOF_BTSAISIECONSO.BDUPLICCONSOClick(Sender: TObject);
var ressource : string;
		Action : string;
begin
	//
	Action := ';ACTION=RECH';

  if CHOIX_MO.checked then
    ressource := AFLanceFiche_Rech_Ressource ('ARS_FONCTION1='+FONCTION.value,'TYPERESSOURCE=SAL,INT' + Action)
  else if CHOIX_RESSOURCE.Checked then
    ressource := AFLanceFiche_Rech_Ressource ('ARS_FONCTION1='+TypeRessource.value,'TYPERESSOURCE=MAT,OUT' + Action)
  else Ressource := '';

  if ressource = '' then exit;

  if not GetConsoAduplic (ressource) then exit;

  EventGrilles(false);
  //
  GHeures.VidePile (false);
  GFrais.videpile (false);
  GMAteriel.videpile (false);
  GFourniture.videpile (false);
  GMOEXT.videpile (false);
  GRecettes.videpile (false);

  Remplitgrilles;

  EventGrilles(true);

end;

procedure TOF_BTSAISIECONSO.BEFFACECHAClick(Sender: TObject);
var TOBL : TOB;
begin
  CONTRAT.text:='';
  //
  if (GetActiveTabSheet('PGTEMPS').Name = 'THEURES') and (TOBMO.detail.count > 0) then
     TOBL := TOBMO.detail[GHeures.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFRAIS') and (TOBFRS.detail.count > 0) then
     TOBL := TOBFRS.detail[GFrais.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMATERIELS') and (TOBRES.detail.count > 0)then
     TOBL := TOBRES.detail[GMateriel.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFOURNITURES') and (TOBMAT.detail.count > 0) then
  	 TOBL := TOBMAT.detail[GFourniture.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMOEXT') and (TOBMOEXT.detail.count > 0) then
  	 TOBL := TOBMOEXT.detail[GMOEXT.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TRECETTES') and (TOBRecettes.detail.count > 0) then
  	 TOBL := TOBRecettes.detail[GRecettes.row-1];
  //
  If TOBL <> Nil then TOBL.putvalue('BCO_AFFAIRESAISIE', '');
  //
end;

procedure TOF_BTSAISIECONSO.ChangeRessourceEtVAlorisation (TOBL: TOB);
var WithMajPrix : boolean;
begin
	WithMajPrix := false;
  if TOBL.GetValue ('BCO_NATUREMOUV')='MO' then WithMajPrix := true;
  TOBL.putValue ('BCO_RESSOURCE', Trim(MAINDOEUVRE.Text));
  TOBL.putValue ('LIBELLEMO', GetMainDoeuvre(MAINDOEUVRE.Text));
  SetInfoRessource (TOBL,TOBRESSource,WithMajPrix,false,PrestationDefaut);
  OKTypeHeure(TOBL,TOBL.GetValue('BCO_TYPEHEURE'),TOBTypeHeure);
end;

procedure TOF_BTSAISIECONSO.ChangeMaterielEtVAlorisation (TOBL: TOB);
var WithMajPrix : boolean;
begin
	WithMajPrix := false;
  TOBL.putValue ('BCO_RESSOURCE', Trim(Ressources.Text));
  TOBL.putValue ('LIBELLEMO', GetMainDoeuvre(Ressources.Text));
  SetInfoRessource (TOBL,TOBRESSource,WithMajPrix,false,PrestationDefaut);
  OKTypeHeure(TOBL,TOBL.GetValue('BCO_TYPEHEURE'),TOBTypeHeure);
end;


function TOF_BTSAISIECONSO.GetConsoADuplic (Ressource : string) : boolean;
var Req : String;
    QQ : TQuery;
    TOBL,TOBC : TOB;
    Indice,IndDate : integer;
    Nbhrs : double;
begin

		result := true;

    NBhrs := 0;

    req := 'SELECT * FROM CONSOMMATIONS WHERE ';

    if choix_MO.checked then
      req := req + '(BCO_NATUREMOUV="MO" OR BCO_NATUREMOUV="FRS" OR BCO_NATUREMOUV="FOU")'
    else
      req := req + '(BCO_NATUREMOUV="RES")';

    req := req + ' AND BCO_RESSOURCE="'+Trim(ressource) + '"';

    if Calendrier.SelectionEnd <> Calendrier.SelectionStart then
    begin
      Req := Req + ' AND BCO_DATEMOUV >= "'+USDATETIME (Calendrier.SelectionStart)+'"';
      Req := Req + ' AND BCO_DATEMOUV <= "'+USDATETIME (Calendrier.SelectionEnd)+'"';
    end
    else
    begin
    	Req := Req + ' AND BCO_DATEMOUV = "'+USDATETIME (Calendrier.SelectionStart)+'"';
    end;

    Req := Req + ' AND BCO_TRANSFORME="-"'; // ne reprends pas les transformé
    Req := Req + ' AND BCO_TRAITEVENTE<>"X"'; // ne reprends pas les passés en livraisons
    Req := Req + ' ORDER BY BCO_DATEMOUV';

    QQ := OpenSql (Req,true,-1,'',true);
    if QQ.eof then
    begin
    	PgiInfo ('Aucune donnée à récupérer sur cette ressource');
    	ferme (QQ);
      result := false;
      exit;
    end;
    //
    TOBMo.ClearDetail;
    TOBFRS.ClearDetail;
    TOBMAT.ClearDetail;
    TOBRES.ClearDetail;
    TOBMOEXT.ClearDetail;
    TOBConso.ClearDetail;
    TOBRecettes.ClearDetail;
    //
    TRY
    TOBCONSO.LoadDetailDB ('CONSOMMATIONS','','',QQ,FALSE);
    IndDate := TOBConso.detail[0].GetNumChamp ('BCO_DATEMOUV');
    for Indice := 0 to TOBConso.detail.count -1 do
    begin
      TOBC := TOBCONSO.detail[indice];
      // Frais - primes
      if TOBC.GetValue('BCO_NATUREMOUV')='FRS' then
      begin
        AddLesSupLignesConso (TOBC);
        // --
        ChangeRessourceEtVAlorisation (TOBC);
        calculeLaLigne (TOBC);
        // --
        TOBL := TOB.Create ('CONSOMMATIONS',TOBFRS,-1);
        AddLesSupLignesConso (TOBL);
        if TOBC.GetValeur (inddate) > LastDayFrs Then LastDayFrs := TOBC.GetValeur (inddate);
        TOBL.Dupliquer (TOBC,true,true);
        TOBL.PutValue('NEW','X');
      end
      // Main d'oeuvre interne
      else if TOBC.GetValue('BCO_NATUREMOUV')='MO' then
      begin
        AddLesSupLignesConso (TOBC);
        // --
        ChangeRessourceEtVAlorisation (TOBC);
        calculeLaLigne (TOBC);
        // --
        TOBL := TOB.Create ('CONSOMMATIONS',TOBMO,-1);
        if TOBC.GetValeur (inddate) > LastDayMo Then LastDayMo := TOBC.GetValeur (inddate);
        AddLesSupLignesConso (TOBL);
        TOBL.Dupliquer (TOBC,true,true);
        TOBL.PutValue('NEW','X');
        IF TOBL.GEtValue('BCO_LINKEQUIPE') <> '' then TOBL.PutValue('BCO_LINKEQUIPE', '');
        NBHRS:= NBHRS + TOBL.GetValue('BCO_QUANTITE');
      end
      else if TOBC.GetValue('BCO_NATUREMOUV')='FOU' then
      begin
        AddLesSupLignesConso (TOBC);
        // --
        ChangeRessourceEtVAlorisation (TOBC);
        calculeLaLigne (TOBC);
        // --
        TOBL := TOB.Create ('CONSOMMATIONS',TOBMAT,-1);
        if TOBC.GetValeur (inddate) > LastDayFourn Then LastDayFourn := TOBC.GetValeur (inddate);
        AddLesSupLignesConso (TOBL);
        TOBL.Dupliquer (TOBC,true,true);
        TOBL.PutValue('NEW','X');
        NBHRS:= NBHRS + TOBL.GetValue('BCO_QUANTITE');
      end
      else if TOBC.GetValue('BCO_NATUREMOUV')='RES' then //Ressource Matériel
      begin
        AddLesSupLignesConso (TOBC);
        // --
        ChangeMaterielEtVAlorisation (TOBC);
        CalculeLaLigne (TOBC);
        // --
        TOBL := TOB.Create ('CONSOMMATIONS',TOBRes,-1);
        if TOBC.GetValeur (inddate) > LastDayMat Then LastDayMat := TOBC.GetValeur (inddate);
        AddLesSupLignesConso (TOBL);
        TOBL.Dupliquer (TOBC,true,true);
        TOBL.PutValue('NEW','X');
        NBHRS:= NBHRS + TOBL.GetValue('BCO_QUANTITE');
      end
      else
        continue;
        PositionneValeurInit (TOBL);
      end;
    FINALLY
      ferme (QQ);
      //
      if LastDayMo       = 0 Then LastDayMo       := Calendrier.SelectionStart;
      if LastDayFrs      = 0 Then LastDayFrs      := Calendrier.SelectionStart;
      if LastDayMat      = 0 Then LastDayMat      := Calendrier.SelectionStart;
      if LastDayFourn    = 0 Then LastDayFourn    := Calendrier.SelectionStart;
      if LastDayMoExt    = 0 Then LastDayMoExt    := Calendrier.SelectionStart;
      if LastDayRecettes = 0 Then LastDayrecettes := Calendrier.SelectionStart;
      //
      if TOBMo.detail.count  = 0 then TOBL := AddNewLigne (TOBMO ,TgsMO);
      if TOBRES.detail.count = 0 then TOBL := AddNewLigne (TOBRES,TgsRES);
      if TOBFRS.detail.count = 0 then TOBL := AddNewLigne (TOBFrs,TgsFRS);
      if TOBMAT.detail.count = 0 then TOBL := AddNewLigne (TOBMAT,TgsFOURN);
      //
      NBHRSSAISIE.Value := NBHRS;
      //
    END;
end;

procedure TOF_BTSAISIECONSO.BVALIDEPLUSClick(Sender: TObject);
begin
	ValideSansCloture := True;
  BValider.Click;
end;

procedure TOF_BTSAISIECONSO.UpdConsoFromLig ( grid : Thgrid;TOBLC,TOBL : TOB);
var year,Month,Day : word;
begin

  if TOBL.GetString('NEW') = 'X' then
  begin
    if CHOIX_CONTRAT.Checked or CHOIX_APPEL.Checked OR CHOIX_CHANTIER.Checked then
    begin
      if TOBL.GetString('MODIF') = 'X' then
      begin
        if (TOBL.GetString('TYPEMODIF') = 'A') Then
        begin
          TOBLC.SetString('BCO_AFFAIRE',  TOBL.GetString('BCO_AFFAIRE'));
          TOBLC.SetString('BCO_AFFAIRE0', TOBL.GetString('BCO_AFFAIRE0'));
          TOBLC.SetString('BCO_AFFAIRE1', TOBL.GetString('BCO_AFFAIRE1'));
          TOBLC.SetString('BCO_AFFAIRE2', TOBL.GetString('BCO_AFFAIRE2'));
          TOBLC.SetString('BCO_AFFAIRE3', TOBL.GetString('BCO_AFFAIRE3'));
        end
        else if (TOBL.GetString('TYPEMODIF') = 'P') Then
          TOBLC.SetString('BCO_PHASETRA', TOBL.GetString('BCO_PHASETRA'))
        else if (TOBL.GetString('TYPEMODIF') = 'D') Then
        Begin
          TOBLC.SetString('BCO_MOIS',     TOBL.GetString('BCO_MOIS'));
          TOBLC.SetString('BCO_SEMAINE',  TOBL.GetString('BCO_SEMAINE'));
          TOBLC.SetString('BCO_DATEMOUV', TOBL.GetString('BCO_DATEMOUV'));
          //
          ConstitueDateDebutTravaux(TOBLC);
          //
          DecodeDate (TOBLC.GetValue('BCO_DATEMOUV'),year,month,day);
          //
          TOBLC.putValue ('JOUR',IntToStr (day));
        end
        else if (TOBL.GetString('TYPEMODIF') = 'Q') Then
        begin
          TOBLC.SetString('BCO_QUANTITE',   TOBL.GetString('BCO_QUANTITE'));
          TOBLC.SetString('BCO_QTEFACTUREE',TOBL.GetString('BCO_QTEFACTUREE'));
        end;
      end;
    end
    else
    begin
      TOBLC.SetString('BCO_AFFAIRE',    TOBL.GetString('BCO_AFFAIRE'));
      TOBLC.SetString('BCO_AFFAIRE0',   TOBL.GetString('BCO_AFFAIRE0'));
      TOBLC.SetString('BCO_AFFAIRE1',   TOBL.GetString('BCO_AFFAIRE1'));
      TOBLC.SetString('BCO_AFFAIRE2',   TOBL.GetString('BCO_AFFAIRE2'));
      TOBLC.SetString('BCO_AFFAIRE3',   TOBL.GetString('BCO_AFFAIRE3'));
      TOBLC.SetString('BCO_PHASETRA',   TOBL.GetString('BCO_PHASETRA'));
      TOBLC.SetString('BCO_MOIS',       TOBL.GetString('BCO_MOIS'));
      TOBLC.SetString('BCO_SEMAINE',    TOBL.GetString('BCO_SEMAINE'));
      TOBLC.SetString('BCO_DATEMOUV',   TOBL.GetString('BCO_DATEMOUV'));
      //
      ConstitueDateDebutTravaux(TOBLC);
      //
      DecodeDate (TOBLC.GetValue('BCO_DATEMOUV'),year,month,day);
      //
      TOBLC.putValue ('JOUR',IntToStr (day));
      //
      TOBLC.SetString('BCO_QUANTITE',   TOBL.GetString('BCO_QUANTITE'));
      TOBLC.SetString('BCO_QTEFACTUREE',TOBL.GetString('BCO_QTEFACTUREE'));
    end;
  end
  else if TOBL.GetString('MODIF') = 'X' then
  Begin
    if (TOBL.GetString('TYPEMODIF') = 'A') Then
    begin
      TOBLC.SetString('BCO_AFFAIRE',  TOBL.GetString('BCO_AFFAIRE'));
      TOBLC.SetString('BCO_AFFAIRE0', TOBL.GetString('BCO_AFFAIRE0'));
      TOBLC.SetString('BCO_AFFAIRE1', TOBL.GetString('BCO_AFFAIRE1'));
      TOBLC.SetString('BCO_AFFAIRE2', TOBL.GetString('BCO_AFFAIRE2'));
      TOBLC.SetString('BCO_AFFAIRE3', TOBL.GetString('BCO_AFFAIRE3'));
    end
    Else if (TOBL.GetString('TYPEMODIF') = 'P') Then
    Begin
      TOBLC.SetString('BCO_PHASETRA', TOBL.GetString('BCO_PHASETRA'));
    end
    else if (TOBL.GetString('TYPEMODIF') = 'D') Then
    Begin
      TOBLC.SetString('BCO_MOIS',     TOBL.GetString('BCO_MOIS'));
      TOBLC.SetString('BCO_SEMAINE',  TOBL.GetString('BCO_SEMAINE'));
      TOBLC.SetString('BCO_DATEMOUV', TOBL.GetString('BCO_DATEMOUV'));
      //
      ConstitueDateDebutTravaux(TOBLC);
      //
      DecodeDate (TOBLC.GetValue('BCO_DATEMOUV'),year,month,day);
      //
      TOBLC.putValue ('JOUR',IntToStr (day));
    end
    else if (TOBL.GetString('TYPEMODIF') = 'Q') Then
    begin
      TOBLC.SetString('BCO_QUANTITE',   TOBL.GetString('BCO_QUANTITE'));
      TOBLC.SetString('BCO_QTEFACTUREE',TOBL.GetString('BCO_QTEFACTUREE'));
    end;
  end;
  //
  calculeLaLigne(TOBLC);
  //
	TOBLC.SetString('BCO_AFFAIRESAISIE',TOBL.GetString('BCO_AFFAIRESAISIE'));
  TOBLC.SetInteger('BCO_PAIENUMFIC',0);

  if TypeSaisie = 'INTERVENTION' then TOBLC.SetString('BCO_FAMILLETAXE1',TOBL.GetString('BCO_FAMILLETAXE1'));

  TOBLC.SetString('MODIF',      TOBL.GetString('MODIF'));
  TOBLC.SetString('TYPEMODIF',  TOBL.GetString('TYPEMODIF'));

  if grid <> nil then AfficheLigne (GriD,TOBLC,TOBLC.getIndex);

end;

procedure TOF_BTSAISIECONSO.AddConsoFromLig (Nature : string; TOBRef,TOBL,TOBRess : TOB;ModeS : TGrilleModeSaisie; Grid : Thgrid; Equipe : string);
var TOBLC : TOB;
		WithMajPrix : boolean;
    Ou : Integer;
begin
  WithMajPrix := (Nature='MO');
  if LigneVide (TOBRef.Detail.count,Modes) then TOBLC := TOBRef.detail[TOBRef.Detail.count-1]
  																				 else TOBLC := AddNewLigne (TOBRef,ModeS);

  TOBLC.SetString('BCO_NATUREMOUV',Nature);
  TOBLC.SetString('BCO_RESSOURCE',TOBRess.GetString('ARS_RESSOURCE'));
  TOBLC.putValue('BCO_LINKEQUIPE',Equipe);
  TOBLC.SetInteger('BCO_PAIENUMFIC',0);
  //
  SetInfoRessource (TOBLC,TOBRESS,WithMajPrix,true,PrestationDefaut);
  //
  Ou := TOBLC.GetIndex+1;
  if Ou >= grid.RowCount then grid.RowCount := Ou+1;
  AfficheLigne (GriD,TOBLC,TOBLC.GetIndex);
end;


procedure TOF_BTSAISIECONSO.UpdatelesLignesEquipe (TOBL : TOB; LinkEquipe : string);
var OneTOB  : TOB;
begin

  OneTOB :=  TOBMO.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  while OneTOB <> nil do
  begin
    if OneTOB <> TOBL THEN UpdConsoFromLig (GHeures,OneTOB,TOBL);
  	OneTOB :=  TOBMO.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;

  OneTOB :=  TOBMoExt.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
	while OneTOB <> nil do
  begin
    if OneTOB <> TOBL THEN UpdConsoFromLig (GMoext,OneTOB,TOBL);
  	OneTOB :=  TOBMoExt.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;

  OneTOB :=  TOBRES.FindFirst(['BCO_LINKEQUIPE'],[LinkEquipe],true);
	while OneTOB <> nil do
  begin
    if OneTOB <> TOBL THEN UpdConsoFromLig (GMAteriel,OneTOB,TOBL);
  	OneTOB :=  TOBRES.FindNext(['BCO_LINKEQUIPE'],[LinkEquipe],true);
  end;

  TOBL.PutVALUE('MODIF', '-');

end;

procedure TOF_BTSAISIECONSO.AddDataEquipe(TOBL ,TOBEquipe : TOB; Equipe :string);
var II          : Integer;
		TOBE        : TOB;
    TOBLC       : TOB;
    TOBREf      : TOB;
    LinkEquipe  : string;
    ModeS       : TGrilleModeSaisie;
    Grid        : THGrid;
    Nature      : string;
    T           : TOB;
    OK_Equipe   : Boolean;
begin

  TOBREf     := nil;

  if TOBL.Getstring('MODIF') = 'X' then
  begin
    if pgiAsk('Voulez-vous appliquer ces modifications à l''ensemble de l''équipe', 'Gestion Equipe') = Mrno then
    Begin
      Ok_Equipe := False;
      TOBL.PutVALUE('MODIF', '-');
      Exit;
    end
    else
    Begin
      Ok_Equipe := True;
      if not ExisteLigneEquipe (TOBL,Equipe) then
      begin
        for II :=0 to TOBEquipe.detail.count -1 do
        begin
          TOBE := TOBEquipe.detail[II];
          if TOBE.GetString('ARS_RESSOURCE')<>TOBL.GetString('BCO_RESSOURCE') then
          begin
            TOBREf := nil;
            if TOBE.GetString('ARS_TYPERESSOURCE')='SAL' then
            begin
              TOBREf := TOBMo;
              ModeS := TgsMO;
              Nature := 'MO';
              Grid := GHeures;
            end else if TOBE.GetString('ARS_TYPERESSOURCE')='INT' then
            begin
              TOBREf := TOBMOEXT;
              ModeS := TgsMoext;
              Nature := 'EXT';
              Grid := GMoext;
            end else if TOBE.GetString('ARS_TYPERESSOURCE')='LOC' then
            begin
              TOBREf := TOBMOEXT;
              ModeS := TgsMoext;
              Nature := 'EXT';
              Grid := GMoext;
            end else if TOBE.GetString('ARS_TYPERESSOURCE')='MAT' then
            begin
              TOBREf := TOBRES;
              ModeS := TgsRES;
              Nature := 'RES';
              Grid := GMAteriel;
            end else if TOBE.GetString('ARS_TYPERESSOURCE')='ST' then
            begin
              TOBREf := TOBMOEXT;
              ModeS := TgsMoext;
              Nature := 'EXT';
              Grid := GMoext;
            end;
            if TOBRef = nil then Exit;
            //
            TOBLC := TOBRef.FindFirst(['BCO_RESSOURCE','BCO_LINKEQUIPE'],[TOBE.GetString('ARS_RESSOURCE'),Equipe],false);
            if TOBLC <> nil then
            begin
               UpdConsoFromLig (grid,TOBLC,TOBL);
            end else
            begin
               AddConsoFromLig (Nature,TOBREf,TOBL,TOBE,ModeS,Grid,Equipe);
            end;
          end;
        end;
      end else
      begin
        if not Ok_Equipe then TOBL.PutVALUE('MODIF', '-');
        UpdatelesLignesEquipe (TOBL,Equipe);
       end;
    end;
  end;
  
end;


Procedure TOF_BTSAISIECONSO.UpdateLignesEquipeExt (TOBL,TOBLE : TOB);
var II : integer;
		TOBLC : TOB;
begin
	for II := 0 to TOBLE.Detail.Count -1 do
  begin
    TOBLC := TOBLE.detail[II];
    UpdConsoFromLig (nil,TOBLC,TOBL);
  end;
end;

procedure TOF_BTSAISIECONSO.AddDataEquipeExt(TOBL ,TOBEquipe : TOB; Equipe : string);
var TOBLE,TOBE,TOBLC : TOB;
		PrestationDefaut : string;
    II : Integer;
    Nature : string;
    WithMajprix : Boolean;
begin
	TOBLE := TOBLIENEQUIPE.FindFirst(['LINKEQUIPE'],[Equipe],false);
  if TOBLE = nil then
  begin
    TOBLE := TOB.Create('XXXXXXXXX XX',TOBLIENEQUIPE,-1);
    TOBLE.AddChampSupValeur('LINKEQUIPE',Equipe);
    for II := 0 to TOBEquipe.detail.count -1 do
    begin
    	TOBE := TOBEquipe.detail[II];
      if TOBE.GetString('ARS_TYPERESSOURCE')='SAL' then
      begin
        Nature := 'MO';
      end else if TOBE.GetString('ARS_TYPERESSOURCE')='INT' then
      begin
        Nature := 'EXT';
      end else if TOBE.GetString('ARS_TYPERESSOURCE')='LOC' then
      begin
        Nature := 'EXT';
      end else if TOBE.GetString('ARS_TYPERESSOURCE')='MAT' then
      begin
        Nature := 'RES';
      end else if TOBE.GetString('ARS_TYPERESSOURCE')='ST' then
      begin
        Nature := 'EXT';
      end;
      TOBLC := TOB.Create('CONSOMMATIONS',TOBLE,-1);

      TOBLC.AddChampSupValeur('NEW','X');
      TOBLC.AddChampSupValeur('MODIF','-');
      TOBLC.AddChampSupValeur('TYPEMODIF','-');

      TOBLC.SetString('BCO_NATUREMOUV',Nature);
      TOBLC.SetString('BCO_RESSOURCE',TOBE.GetString('ARS_RESSOURCE'));
      TOBLC.putValue('BCO_LINKEQUIPE',Equipe);
      TOBLC.SetInteger('BCO_PAIENUMFIC',0);
      //
      WithMajprix := (Nature='MO');
      SetInfoRessource (TOBLC,TOBE,WithMajPrix,true,PrestationDefaut);
  //
    end;
  end else
  begin
    UpdateLignesEquipeExt(TOBL,TOBLE);
  end;
end;

procedure TOF_BTSAISIECONSO.TraiteLigneEquipe(TOBL: TOB);

	procedure DecodeLinkEquipe (LinkEquipe: string; var Equipe,Link : string);
  begin
  	Equipe := '';
    Link := '';
    if LinkEquipe = '' then exit;
    Equipe := READTOKENST(LinkEquipe);
    if LinkEquipe <> '' then Link := READTOKENST(LinkEquipe);
  end;

  procedure GetEquipe(Equipe: string;TOBEquipe : TOB);
  var QQ : TQuery;
  begin
		QQ := OpenSQL('SELECT * FROM RESSOURCE WHERE ARS_EQUIPERESS="'+Equipe+'"',True,-1,'',true);
    if not QQ.eof then
    begin
			TOBEquipe.LoadDetailDB ('RESSOURCE','','',QQ,false);
    end;
    Ferme (QQ);
  end;

var Equipe : string;
		Link : string;
    TOBEquipe : TOB;
begin
  TOBEquipe := TOB.Create('UN EQUIPE',nil,-1);
  TRY
    DecodeLinkEquipe (TOBL.GetString('BCO_LINKEQUIPE'), Equipe,Link);
    if Equipe = '' then Exit;
    GetEquipe(Equipe,TOBEquipe);
    if TOBEquipe.detail.count = 0 then Exit;

    if (Pos(TOBL.GetValue('BCO_NATUREMOUV'),'MO;RES;EXT')>0)  then
    begin
      TOBL.PutValue('MODIF', 'X');
      if (CHOIX_CHANTIER.Checked) OR (CHOIX_CONTRAT.Checked) OR (CHOIX_APPEL.Checked) then
      begin
        // Dans ce cadre on enregistre dans les bonnes TOB (en creation ou en mise a jour)
        AddDataEquipe(TOBL,TOBEquipe,TOBL.GetString('BCO_LINKEQUIPE'));
      end else if (CHOIX_MO.checked) or (CHOIX_RESSOURCE.Checked) then
      begin
        // Dans ce cadre les infos ne sont pas en tob de saisie standard
        AddDataEquipeExt(TOBL,TOBEquipe,TOBL.GetString('BCO_LINKEQUIPE'));
      end;
    end;
  FINALLY
    TOBEquipe.free;
  END;
end;

procedure TOF_BTSAISIECONSO.RecupInfoEquipe(TOBC : TOB);
var TOBCC,TOBLE,TOBCD : TOB;
		QQ : TQuery;
    ii : Integer;
    LinkEquipe : string;
    Req : string;
    first : boolean;
begin
  first := True;

	LinkEquipe := TOBC.GetString('BCO_LINKEQUIPE');

  if LinkEquipe = '' then Exit;

	TOBCC := TOB.Create('LES CONSO',nil,-1);

  try
    Req := 'SELECT * FROM CONSOMMATIONS WHERE BCO_LINKEQUIPE="' + LinkEquipe + '" ';
    //Req := Req + '    AND BCO_RESSOURCE <> "' + TOBC.GetString('BCO_RESSOURCE') + '"';

		QQ := OpenSQL(Req,True,-1,'',true);
    if Not QQ.eof then
    begin
			TOBCC.LoadDetailDB('CONSOMMATIONS','','',QQ,false);
      II := 0;
      for II := TOBCC.detail.count -1  downto 0 do
      begin
        TOBCD := TOBCC.detail[II];
      	if first then
        begin
					TOBLE := TOB.Create('XXXXXX XX',TOBLIENEQUIPE,-1);
          TOBLE.AddChampSupValeur('LINKEQUIPE',LinkEquipe);
          first := false;
        end;
        AddLesSupLignesConso (TOBCD);
        TOBCD.ChangeParent(TOBLE,-1);
      end;
    end;
    Ferme(QQ);
  finally
    TOBCC.free;
  end;

end;

procedure TOF_BTSAISIECONSO.FAMILLETAXEChange(Sender: TObject);
Var TOBL     : TOB;
Begin
  //
  if (GetActiveTabSheet('PGTEMPS').Name = 'THEURES') and (TOBMO.detail.count > 0) then
     TOBL := TOBMO.detail[GHeures.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFRAIS') and (TOBFRS.detail.count > 0) then
     TOBL := TOBFRS.detail[GFrais.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMATERIELS') and (TOBRES.detail.count > 0)then
     TOBL := TOBRES.detail[GMateriel.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TFOURNITURES') and (TOBMAT.detail.count > 0) then
  	 TOBL := TOBMAT.detail[GFourniture.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TMOEXT') and (TOBMOEXT.detail.count > 0) then
  	 TOBL := TOBMOEXT.detail[GMOEXT.row-1]
  else if (GetActiveTabSheet('PGTEMPS').Name = 'TRECETTES') and (TOBRecettes.detail.count > 0) then
  	 TOBL := TOBRecettes.detail[GRecettes.row-1];
  //
  if TypeSaisie = 'INTERVENTION' then
  begin
    If TOBL <> Nil then TOBL.putvalue('BCO_FAMILLETAXE1', GetControlText('FAMILLETAXE1'));
  end;
  //
end;

procedure TOF_BTSAISIECONSO.BChangeEtatClick(Sender: TObject);
Var TOBL    : TOB;
    Affaire : String;
begin

  //Cette fonction ne sera valable que sur la saisie des heures...
  TOBL := TobMO.Detail[GHeures.Row-1];

  if TOBL = nil then exit;

  if not TOBL.FieldExists('ETATAFFAIRE') then exit;

  Affaire := TOBL.GetValue('BCO_AFFAIRE');

  //On recherche dans la TOBMO toutes les lignes qui correspondent à cette affaire/Appel
  TOBL := TobMO.FindFirst(['BCO_AFFAIRE'], [Affaire], False);
  While TOBL <> nil do
  begin
    if TOBL.getValue('ETATAFFAIRE') = 'REA' then
      TOBL.PutValue('ETATAFFAIRE', 'ECR')
    else if TOBL.getValue('ETATAFFAIRE') = 'ECR' then
      TOBL.PutValue('ETATAFFAIRE', 'REA');
    //
    TOBL.PutValue('CHANGEETAT', 'X');
    //
    If TOBL.GetValue('BCO_AFFAIRE0') = 'W' then
      LLABELETAT.Caption := RechDom('BTETATAPPEL',   TOBL.GetValue('ETATAFFAIRE'), False)
    else
      LLABELETAT.Caption := RechDom('AFETATAFFAIRE', TOBL.GetValue('ETATAFFAIRE'), False);
    //
    TOBL := TobMO.FindNext(['BCO_AFFAIRE'], [Affaire], False);
  end;

end;

Initialization
  registerclasses ( [ TOF_BTSAISIECONSO ] ) ;
end.

