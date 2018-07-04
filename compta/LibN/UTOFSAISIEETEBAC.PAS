{***********UNITE*************************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001 (2/20)
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CRELBQE ()
Mots clefs ... : TOF;CRELBQE
*****************************************************************}
Unit UTOFSAISIEETEBAC ;

Interface

Uses StdCtrls, Controls, Classes,
     {$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,
     // composant AGL
     HTB97,
     HSysMenu,
     HFLabel,
     // VCL
     Windows,
     Messages,
     Graphics,
     Dialogs,
     Menus,
     buttons,
     Grids,  // pour TGridDrawState
     extctrls, // pour le TBevel
     HPanel,
     // lib
     Ent1,
     ParamDat,
     LookUp,
     CPGeneraux_TOM,
     ZCompte,
     eSaisAnal,
     {$IFNDEF EAGLCLIENT}
     FE_Main,   // pour AGLLanceFiche()
     PrintDBG, // pour le prindDB
     {$ELSE}
     MaineAGL,
     {$ENDIF}
     SaisUtil,  // pour RDevise
     ParamSoc,  // pour le GetParam
     AGLInit,
     Vierge, // pour le form closeQuery
     ed_tools, // pour le TQRProgressForm
     // Objet de gestion
     ZEtebac,
     ZEtebacGuide,
     UCstEtebac,
     ULibEcriture,
     ULibExercice,
     ULibWindows,
     ZEtebacVersEcr,
     uTOFCPMULGUIDE,
     // AGL
     UTOB,
     HXLSPAS,
     UTOF,
     uObjFiltres;  // GCO - 07/06/2006 - FQ 18307


Type


  TOF_SAISIEETEBAC = Class (TOF)
   private
    // variable interne
    FStatutGrilleReleve     : TActionFiche ;          // Etat de la grille ( modif,creation...)
    FStatutGrilleImput      : TActionFiche ;          // Etat de la grille ( modif,creation...)
    FBoGuideRun             : boolean;                // True la saisie par guide est active
    FCurrentCol             : integer;                // Colonne courante de la grille d'imputation
    FCurrentRow             : integer;                // Ligne couranet de la grille d'imputation
    FRdNumeroPieceMax       : double;                 // Numero de piece max de la grille : sert dans pour le Shift+F7
    FRdNumeroPieceMaxOld    : double;                 // on utilise cette variable pour retrouver le precedent numero en cas d'annulation
    FRdSolDeT               : double;                 // solde théorique
    FBoChoixEffectue        : boolean;                // La saisie est bloque tant que l'on n'a pas choisit un compte de banque et de contrepartie
    FBoGrilleVide           : boolean;                // true si la grille est en creation mais que l'on n'a rien enregistré
    FInNumReleve            : integer;
    FInFirstNumReleve       : integer;
    FBoModeSelection        : boolean;
    FBoSorted               : boolean;
    FStListeNumASupp        : string;
    FStArgument             : string;                 // Paramètres passés par AGLLanceFiche
    FListeIntegrer          : TStringList;            // StringList contenant le code des lignes à intégrer dans la tabel ecriture
    FExceptionEvent         : TExceptionEvent;        // evenement gerant toutes les exceptions non gérées de la saisies
    // liste des TOB
    FZEtebac                : TZEtebac;               // TOB de gestion de la grille de saisie
    FZEtebacGuide           : TZEtebacGuide;          // TOB de gestion de la grille d'imputation
    FTOBSauveLigne          : TOB;
    FInfoEcr                : TInfoEcriture;          // objet contenant la liste des journaux,... partage entre par l ZEtebac et les guides
    FMessCompta             : TMessageCompta ;        // affichage des messages
    // liste des contrôles
    PgcCritere              : TPageControl;           // Page à onglet de recherche
    TbsStandard             : TTabSheet;              // Onglet standard
    HGSaisie                : THGrid;                 // Grille de saisie
    HGImputation            : THGrid;                 // Grille d'imputation
    PFiltres                : TToolWindow97;          // Fenêtre contenant l'enregistrement des filtres
    PBouton                 : TToolWindow97;          // Fenêtre contenant les boutons
    BFermer2                : TToolbarButton97;       // Bouton de fermeture du mode d'intégration
 //   BFerme                  : TToolbarButton97;       // Bouton de fermeture du mode d'intégration
    BMode                   : TToolbarButton97;       // Bouton de passage en mode selection
    DckDock                 : TDock97;                // Dock de PFiltres
    DckDock971              : TDock97;                // Dock de PBouton
    BInsert                 : TToolbarButton97;       // Bouton BInsert
    BCherche                : TToolbarButton97;       // Bouton 'Mouette'
    BDelete                 : TToolbarButton97;       // Bouton BCherche
    BValider                : TToolbarButton97;       // Bouton Valider : coche verte
    BImprimer               : TToolbarButton97;       // Bouton imprimer
    BDefaire                : TToolbarButton97;       // Bouton Annulation
    BAgrandir               : TToolbarButton97;       // Bouton Agrandir la liste
    BMenuZoom               : TToolbarButton97;       // Bouton Zoom
    EdtMulDateDu            : THCritMaskEdit;         // Edit Date 'du'
    EdtMulDateAu            : THCritMaskEdit;         // Edit Date 'au'
    EdtMulNumeroPiece       : THCritMaskEdit;         // Edit des pièces bancaire
    EdtMulNumeroPieceCpml   : THCritMaskEdit;         // Edit des pièces complaimentaire
    CmbMulEtat              : THMultiValComboBox;     // Edit de l'etat des imputations
    EdtMulLibelle           : THCritMaskEdit;         // Edit du libellé
    EdtMulGeneral           : THCritMaskEdit;         // Edit du compte d'imputation
    CmbMulValeur1           : THValComboBox;          // CombBox Debit ou Crédit
    CmbMulValeur2           : THValComboBox;          // CombBox Debit ou Crédit
    CmbMulChoix1            : THValComboBox;
    CmbMulChoix2            : THValComboBox;
    EdtMulValeur1           : THCritMaskEdit;
    EdtMulValeur2           : THCritMaskEdit;
    CmbMulOp                : THValComboBox;
    CmbEtablissement        : THValComboBox;         // Edit Etablissement
    CmbContrepartie         : THValComboBox;         // Edit Compte bancaire
    CmbJournal              : THValComboBox;         // Edit Compte bancaire
    HlbDebitSaisie          : THLabel;
    HlbDebitImput           : THLabel;
    HlbCreditSaisie         : THLabel;
    HlbCreditImput          : THLabel;
    BvlDebitSaisie          : TBevel;
    BvlCreditSaisie         : TBevel;
    BvlDebitImput           : TBevel;
    BvlCreditImput          : TBevel;
    BvlSoldeImput           : TBevel;
    HlbSoldeImput           : THLabel;
    HSystemMenu1            : THSystemMenu;
    EdtDate                 : THCritMaskEdit;
    FlashGuide              : TFlashingLabel;
    FlashCurrentReleve      : TFlashingLabel;
    HlbDevise               : THLabel;                // affiche les info sur la devise du dossier
    HlbSoldeCompte          : THLabel;
    BvlSoldeCompte          : TBevel;
    HlbLabelSoldeCompte     : THLabel;
    ChcBouclage             : TCheckBox;              // CheckBox de bouclage en creation de guide
    EdtFFiltres             : THValComboBox;          // Liste des filtres disponibles
    BFiltre                 : TToolBarButton97;
    POPF                    : TPopupMenu;
    POPS                    : TPopupMenu;
    POPZOOM                 : TPopupMenu;
    BCompte                 : TToolbarButton97;
    BGuide                  : TToolbarButton97;
    FlashCompte             : TFlashingLabel;
    FlashJournal            : TFlashingLabel;
    HlbNomGuide             : THLabel;               // label affichant le nom du guide
    ImgEuro                 : TImage;                // Image du sigle EURO
    HlbCaptionSoldeImput    : THLabel;
    BIntegrer               : TToolbarButton97;       // Bouton BCherche
    BSelectionner           : TToolbarButton97;       // Bouton BCherche
    BDeSelectionner         : TToolbarButton97;       // Bouton BCherche
    BExport                 : TToolbarButton97;       // Bouton BCherche
    BImport                 : TToolbarButton97 ;
    BCreerGuide             : TToolbarButton97 ;
    SD                      : TSaveDialog;
    FListeMessage           : HTStringList ;
    FBoParle                : Boolean ;
    FBoChainage             : Boolean ;
    FStLastCrit             : string ;
    FInNumFolio             : integer ;
    FBoExoVide              : boolean ;
    FObjFiltre : TObjFiltre; // GCO - 07/06/2006 - FQ 18307

    // liste des evenement
    procedure BInsertClick(Sender: TObject);          // Insert une ligne de releve
    procedure BChercheClick(Sender: TObject);         // Bouton 'Mouette'
    procedure BDeleteClick(Sender: TObject);          // Bouton Supprimer
    procedure BValiderClick(Sender: TObject);         // Bouton Supprimer
    procedure BModeClick(Sender: TObject);            // Bouton Supprimer
    procedure BFermerClick(Sender: TObject);            // Bouton Supprimer
    procedure BImprimerClick(Sender: TObject);        // Bouton Supprimer
    procedure BDefaireClick(Sender: TObject);         // Bouton Annulation
    procedure BAgrandirClick(Sender: TObject);        // Bouton Agrandir Liste
    procedure BIntegrerClick(Sender: TObject);
    procedure BSelectionnerClick(Sender: TObject);
    procedure BDeSelectionnerClick(Sender: TObject);
    procedure BCompteClick(Sender: TObject);
    procedure BGuideClick(Sender: TObject);
    procedure BExportClick(Sender: TObject);
    procedure BImportClick(Sender: TObject);
    procedure BCreerGuideClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);

    // contrôle que la date de debut est supérieur a la date de debut de l'excercice N
    procedure EdtMulDateDuChange( Sender : TObject );
    // contrôle que la date de fin est inférieur a la date de fin de l'excercice N+1
    procedure EdtMulDateAuChange( Sender : TObject );

    // Gestion de la grille de saisie des relevées

    // Gestion du déplacement dans la grille pour les cases débit et crédit
    procedure HGSaisieCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    // Gestion de l'affichage des généraux
    procedure HGSaisieElipsisClick(Sender: TObject);
    // vérif de la validité des imputations et formatage des cellules
    procedure HGSaisieCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
     // enregistrement en cas de modif
    procedure HGSaisieRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    // suppression d'une ligne
    procedure HGSaisieDeleteRow( ARow : integer) ;
    // grise les cellules débit /crédit en fct de leur état
    procedure HGSaisiePostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    // affectation du statut de la grille
    procedure HGSaisieKeyPress(Sender: TObject; var Key: Char);
    // idem ( gere le VK_DELETE )
    procedure HGSaisieKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    // rafraichie la grille avec les données de la TOB
    procedure HGSaisieRefreshLigne( ARow : integer );
    // initialise une nouvelle ligne dans la grille ( valeur par defaut )
    procedure HGSaisieInitLigne( value : integer );
    // rempli la grille avec les valeurs de la TOB
    procedure HGSaisieRemplirGrille;
    // test la validité d'une ligne ( utilise le ZReleveBanque.IsRowValidate
    function  HGSaisieIsRowValid(Row : LongInt) : Boolean ;
    // remplie la TOB avec les cellules de la grilles
    function  HGSaisieAssignTOB( ARow : integer) : boolean;
    // recherche des imputations
    function  HGSaisieRechercheImputation : boolean;
    // affiche le solde des debit credit
    procedure HGSaisieAfficheSolde;
    // rafraichit le numero de ligne, calcul les soldes des comptes ...
    procedure HGSaisieRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure HGSaisieClick(Sender: TObject);

    // Gestion de la grille de saisie des imputations
    procedure HGImputationPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    // Gestion du déplacement dans la grille pour les cases débit et crédit
    procedure HGImputationCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    // Gestion des deplacements dans la grille des imputations : mode guide  + mode normal
    // la fin du guide est gere dasn cette methode
    procedure HGImputationCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    // Affecte la ligne d'imputation courante avec la ligne de releve courante
    function  HGImputationAssignLigneTOB( ARow : integer) : boolean;
    // affectation du statut de la grille
    procedure HGImputationKeyPress(Sender: TObject; var Key: Char);
    // idem ( gere le VK_DELETE )
    procedure HGImputationKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    // Gestion de l'affichage des collectifs
    procedure HGImputationElipsisClick( Sender : TObject );
    // affecte la tob des imputations avec l'ensembles des enregistrement de la grille
    function  HGImputationAssignTOB : boolean;
    // rempli la grille avec les données presente en base
    function  HGImputationRemplirGrille : boolean;
    // on repasse HGImputation.ElipsisButton à false car le cellenter n'est pas declencher quand on sort et revient dans la
    // grille par code. C'est le cas quand on enchaine les modes guides.
    procedure HGImputationExit(Sender: TObject);
    // Test la validite de la ligne d'imputation avant de sortir
    procedure HGImputationRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    // enregistrement de l'imputation, passage à ligne suivante de releve et bouclage su cmbBouclage est selectionne
    procedure VK_F10Imputation;
    procedure VK_F6Imputation;


    // Gestion de la combobox du choix du compte de
    // contrepartie
    procedure CmbContrepartieChange(Sender: TObject);
    procedure CmbJournalChange(Sender: TObject);
    procedure CmbEtablissementChange(Sender: TObject);

    // On recherche l'ensemble des comptes de contrepartie des
    // journaux de type banques et on calcule le solde de la
    // periode precedente du compte
    procedure RemplirComboBoxContrepartie;

    // Value = True -> on affecte les evenement à la grille   Value = false on debranche tous les evenements ( sert lors de l'ajout de nouvelle ligne
    procedure SetEvent( value : boolean);
    // gestion des touches VK_DOWN,VK_F5,VK_RETURN dans les grille de saisie
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    // lance la recherche d'un guide et remplit la grille des imputations
    function  AfficheGuide( vBoLastGuide : boolean = false ) : boolean;
    // Ajout d'un nouvelle enregistrement dans la grille des releves
    procedure CreateRow;
    // Ajout d'un nouvelle enregistrement dans la grille des imputations
    procedure NouvelleImputation;
    // enregistrement d'un ligne en base - utiliser sur le RowExit
    function  EnregistreLigne ( ARow : integer ) : boolean;

    // affecte la statut de la grille de saisie
    procedure SetStatutGrilleReleve ( Value : TActionFiche );
    // affecte la statut de la grille de saisie
    procedure SetStatutGrilleImput ( Value : TActionFiche );

    procedure SynchroniseHGSaisie;
    procedure AfficheSoldeImputation;
    procedure SetModeGuide( Value : boolean);
    procedure EnabledControl;
    procedure RafraichirLigne;
    procedure NumeroteLigne( ARow : integer; G : THGrid);
    procedure VK_F7Grille( G : THGrid );
    procedure ShiftVK_F7Grille( G : THGrid ; vIncrement :integer = 1 ) ;
    procedure VideGrille ( G : THGrid );
    procedure HGSaisieGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure HGImputationDeleteRow(ARow: integer);
    procedure RemettreEnEditionHGSaisie;
    procedure HGImputationGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure EdtMulDateAuExit(Sender: TObject);
    procedure EdtMulDateDuExit(Sender: TObject);
    procedure EdtDateValide( Sender : THCritMaskEdit );
    procedure CalculSoldeTheorique;
    function  AltDImputation: boolean;
    procedure RechercheRibAIntegrer;
    function  NextColStop(ARow, ACol: integer): TPoint;
    procedure FormCloseQuery( Sender : TObject ; var CanClose : Boolean);
    procedure ZoomGeneral;
    procedure ZoomEcrs;
    procedure POPSPopup(Sender: TObject);
    function  Integration(vNumEtebac : integer ; vTOBLettrage : TOB = nil): boolean;
    procedure ReaffecterLigneReleve;
    procedure SetModeSelection;
    procedure TransactionSaveReleve;
    procedure GestionException (Sender: TObject; E: Exception);
    procedure HGImputationRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    function  SelectionnerLigne : boolean;
    procedure ChargeParam;
    procedure AfficheLookUpImput;
    function  IsDateValide(Sender: string) : boolean;
    procedure HGSaisieOnSorted(Sender: TObject);
    // affiche les erreurs retroune par la recherche des comptes
    procedure OnError (sender : TObject; Error : TRecError ) ;
    function  GetCol : integer;
    procedure OnArgumentSaisieTresorerie  (S : String ) ;
    procedure OnArgumentIntegrationTresorerie (S : String ) ;
    procedure OnClickF6(Sender: TObject);
    procedure OnClickAltD(Sender: TObject);
    procedure Integrer (vTOBLettrage : TOB = nil );
   protected
    function  GetControlTOF           : boolean; // recupere les contrôles de la grille
    function  AssignEvent             : boolean; // Affecte les evenements au contrôle
    function  InitControl             : boolean; // initialise les grilles
    function  CreateControl           : boolean; // creation de TOB
  public
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnClose                  ; override ;

    procedure ImportFromDP ;

    property  StatutGrilleReleve : TActionFiche read FStatutGrilleReleve write SetStatutGrilleReleve;
    property  StatutGrilleImput  : TActionFiche read FStatutGrilleImput  write SetStatutGrilleImput;
    procedure ApresChargementFiltre;

  end ;

  procedure CPLanceFiche_SaisieTresorerie( vStArgument : string = '');

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  UTofConsEcr, dpTOFCETEBAC, CPMULCETEBAC_TOF,
  CPGUIDE_TOM,
  GuidUtil;

  const cNomFiltre = 'CSAISIEETEBAC';

  type
  TTemp = Class
   StValue : string;
   StEta   : string ;
  end;

var
 TheLastError : string ;

{$IFDEF TT}
var
 TheLog : TStringList;

procedure DelEvenement ;
begin
 TheLog.Clear ;
 TheLog.SaveToFile('c:\SaisieEtabac.txt') ;
end;

procedure AddEvenement( value : string );
begin
 TheLog.Add(Value) ;
 TheLastError:=value;
 TheLog.SaveToFile('c:\SaisieEtabac.txt') ;
end;
{$ENDIF}

function _EstAutoriseEtebac : boolean ;
begin

 result := false ;

 if not ExJaiLeDroitConcept(TConcept(ccSaisEcritures), TRUE) then
  exit;

 if not CExoRefOuvert(true) then
  exit;

 if not VH^.ZACTIVEPFU then
  begin // l'option activé les nouvelles saisie n'est pas active
   PGIInfo('Vous ne pouvez pas accéder à cette fonction',cStTexteTitreFenetre);
   Exit;
  end ;

 if VH^.EtablisDefaut = '' then
  begin
   PGIINfo(cStTextePasEtablissement,cStTexteTitreFenetre);
   Exit ;
  end ;

 // pas de journal de banque
 if not ExisteSQL('select J_JOURNAL from JOURNAL where J_FERME="-" and J_NATUREJAL="BQE" ') then
  begin
   PGIInfo(cStTexteJournalPasPresent,cStTexteTitreFenetre);
   exit;
  end;

 if ( VH^.DateCloturePer > 0 ) and ( FinDeMois(VH^.Entree.Fin) <= VH^.DateCloturePer ) then
  begin
   PGIInfo(Format('Toutes les périodes sont clôturées entre le %s et le %s', [DateToStr(VH^.Entree.Deb), DateToStr(VH^.Entree.Fin)]), cStTexteTitreFenetre);
   exit;
  end ;

 result := true;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 12/01/2006
Modifié le ... :   /  /    
Description .. : - LG - 12/01/2005 - FB 17187 - Supression des 
Suite ........ : enregistrement ds cEtabac en import automatique
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.ImportFromDP ;
var
 lQ           : TQuery;
 lTOB         : TOB ;
 lTOBLettrage : TOB ;
 i            : integer ;
begin


 lQ := OpenSQL(' select cet_numreleve,j_contrepartie,j_journal,cet_etabbq,cet_guichet,cet_numerocompte, bq_general,j_modesaisie ' +
              ' from cetebac, banquecp, journal ' +
              ' where bq_general=j_contrepartie ' +
              ' AND BQ_NODOSSIER = "' + V_PGI.NoDossier+'" ' + // 19/10/2006 YMO Multisociétés
              ' and bq_etabbq=cet_etabbq and bq_guichet=cet_guichet and bq_numerocompte=cet_numerocompte ' +
              ' and cet_typeligne="01" ' , true ) ;

 if lQ.EOF then exit ;


 lTOB := TOB.Create('',nil,-1) ;
 lTOB.LoadDEtailDB('','','',lQ,true) ;
 Ferme(lQ) ;

 lTOBLettrage := TOB.Create('',nil,-1) ;

 try

 if Blocage([cStVerrouTreso], true , cStVerrouTreso ) then exit;

 for i := 0  to lTOB.Detail.Count - 1 do
  begin

   FInNumReleve                            := FZEtebac.ProchainCompteur;
   FInFirstNumReleve                       := FInNumReleve;
   FZEtebacGuide.StCompteTVA               := FZEtebac.StCompteTVA;
   FZEtebac.TypeContexte                   := TModeAuto;
   FZEtebac.StCompteContrepartie           := lTOB.Detail[i].GetValue('J_CONTREPARTIE') ;
   FZEtebac.StJournalContrepartie          := lTOB.Detail[i].GetValue('J_JOURNAL') ;
   FZEtebac.Initialize ;
   FZEtebacGuide.StJournalContrepartie     := FZEtebac.StJournalContrepartie;
   FZEtebacGuide.ZCompte                   := FInfoEcr.Compte;
   FZEtebacGuide.StLeNomTable              := 'CRELBQE';
   FZEtebacGuide.StEtablissement           := VH^.EtablisDefaut ;
   FZEtebacGuide.StDevise                  := FZEtebac.Devise.Code;
   FZEtebacGuide.StCompteContrepartie      := FZEtebac.StCompteContrepartie;
   FZEtebacGuide.SilentMode                := true ;
    // chargement de l'ensemble des guides
   FZEtebacGuide.Load;
   Integration( lTOB.Detail[i].GetValue('CET_NUMRELEVE'),lTOBLettrage ) ;
   FStListeNumASupp := lTOB.Detail[i].GetValue('CET_NUMRELEVE') ;

   if Transactions(TransactionSaveReleve,0) <> oeOK then
    begin
     MessageAlerte('Erreur lors de l''enregistrement des fichiers' + #13#10 + V_PGI.LastSQLError );
     exit ;
    end ;

    if FInfoEcr.Journal.GetValue('J_TRESOCHAINAGE') = 'X' then
     begin
      FBoParle    := false ;
      FBoChainage := true ;
      Integrer(lTOBLettrage) ;
      FBoParle    := true ;
      FBoChainage := false ;
     end ;

  end ; // for


  finally
   lTOB.Free ;
   lTOBLettrage.Free ;
   Bloqueur( cStVerrouTreso , false );
  end ;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/05/2006
Modifié le ... : 16/05/2006
Description .. : - LG - 16/05/2006 - suppression d'une fuite memoire. appel 
Suite ........ : du OnClose
Mots clefs ... : 
*****************************************************************}
procedure CPLanceFiche_IntegTresorerie ;
var
 lMyTOF : TOF_SAISIEETEBAC ;
begin

  if (ctxPCL in V_PGI.PGIContexte) and not VH^.OkModExpertETEBAC then exit ;

  lMyTOF := TOF_SAISIEETEBAC.Create(nil,false) ;
  lMyTOF.OnArgumentIntegrationTresorerie('') ;
  lMyTOF.ChargeParam ;
  lMyTOF.ImportFromDP ;
  lMyTOF.OnClose ;
  lMyTOF.Free ;

end ;




{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... : 20/07/2001
Description .. : Poinr d'entre de la saisie de tresorie
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_SaisieTresorerie( vStArgument : string = '');
begin
 if not _EstAutoriseEtebac then exit ;
 if vStArgument = 'DP' then
  CPLanceFiche_IntegTresorerie
   else
    AGLLanceFiche('CP','CSAISIEETEBAC','','', vStArgument ) ;
end ;




{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/11/2006
Modifié le ... :   /  /    
Description .. : - 29/11/2006 - LG - FB 19239 - affectation de l'exo ref s'il 
Suite ........ : n'est pas definie
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.OnArgument (S : String ) ;
begin
 inherited ;

 {$IFDEF TT}
  VH^.OkModExpertETEBAC := true ;
 {$ENDIF}

 if VH^.CPExoRef.Code = '' then
  begin
   FBoExoVide   := true ;
   VH^.CPExoRef := VH^.EnCours ;
  end
   else
    FBoExoVide := false ;

 FStArgument := S;
 FBoChainage := false ;

 {$IFDEF TT}
 TheLog:=TStringList.Create;
 TheLog.SaveToFile('c:\SaisieEtabac.txt');
 V_PGI.SAV := true ;
{$ENDIF}

  // on sauvegarde la valeur courante
 FExceptionEvent         := Application.OnException;
 // on branche le nouveau gestionnaire d'exception
 Application.OnException := GestionException;

 OnArgumentSaisieTresorerie (S ) ;
 
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.OnLoad;
begin
  inherited;
  // GCO - 07/06/2006 - FQ 18307
//  FObjFiltre.FFI_TABLE := cNomFiltre;
//  FObjFiltre.Charger;

//  ApresChargementFiltre ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... :   /  /
Description .. : Rem : on rentre directement en création
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.OnArgumentSaisieTresorerie (S : String ) ;
var lStArgument : string;
begin

 lStArgument := S;

 {$IFDEF TT}
 TheLog:=TStringList.Create;
 TheLog.SaveToFile('c:\SaisieEtabac.txt');
{$ENDIF}

 if not GetControlTOF then exit;
 if not CreateControl then exit;
 if not InitControl   then exit;
 if not AssignEvent   then exit;

 // creation de la tob de sauvegarde
 FTOBSauveLigne := TOB.Create('',nil,-1);

 // on remplit le combox contrepartie qui va remplir le comboxbox journal puis le OnChange du Combobox
 // journal va appeler la procedure Nouveau;
 RemplirComboBoxContrepartie;
 if not(FBoChoixEffectue) and ( CmbContrepartie.CanFocus) then
  CmbContrepartie.SetFocus; // l'utilisateur doit faire un choix

 FListeMessage:=HTStringList.Create ; CInitMessageBor(FListeMessage) ;

 // GCO - 13/09/2004 - FQ 10605
 if lStArgument <> '' then
 begin
   CmbContrepartie.ItemIndex := CmbContrepartie.Items.IndexOf(ReadTokenSt(lStArgument));
   CmbContrepartieChange(nil);
 end;

 FBoParle := true ;

end;

procedure TOF_SAISIEETEBAC.OnArgumentIntegrationTresorerie (S : String ) ;
var lStArgument : string;
begin

 {$IFDEF TT}
  TheLog:=TStringList.Create;
  TheLog.SaveToFile('c:\SaisieEtabac.txt');
{$ENDIF}


 lStArgument := S;

 if not CreateControl then exit;

 FListeMessage:=HTStringList.Create ; CInitMessageBor(FListeMessage) ;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 29/11/2006
Description .. : Destruction des objets
Suite ........ : - 29/11/2006 - LG - FB 19239 - affectation de l'exo ref s'il 
Suite ........ : n'est pas definie
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.OnClose ;
begin
 {$IFDEF TT}
  TheLog.Free;
{$ENDIF}

 if FBoExoVide then
  begin
   VH^.CPExoRef.Code              := '' ;
   VH^.CPExoRef.Deb               := 0 ;
   VH^.CPExoRef.Fin               := 0 ;
   VH^.CPExoRef.DateButoir        := 0 ;
   VH^.CPExoRef.DateButoirBud     := 0 ;
   VH^.CPExoRef.DateButoirBudgete := 0 ;
   VH^.CPExoRef.NombrePeriode     := 0 ;
   VH^.CPExoRef.EtatCpta          := '' ;
  end ;

 FreeAndNil(FObjFiltre);

 // on sauvegarde le numero de dossier max
 SetParamSoc('SO_CPTRESONUMPIECEMAX',FRdNumeroPieceMax) ;

 PurgePopup(POPF);
 PurgePopup(POPZOOM);
 PurgePopup(POPS);

 try

  if assigned(FInfoEcr)          then FInfoEcr.Free;
  if assigned(FZEtebacGuide)     then FZEtebacGuide.Free;
  if assigned(FListeIntegrer)    then FListeIntegrer.Free;
  if assigned(FTOBSauveLigne)    then FTOBSauveLigne.Free;
  if assigned(FZEtebac)          then FZEtebac.Free;
  if assigned(SD)                then SD.Free;
  if assigned(FMessCompta)       then FMessCompta.Free ;
  if assigned(FListeMessage)     then FListeMessage.Free ;

 finally
  FZEtebac                       := nil;
  FZEtebacGuide                  := nil;
  FListeIntegrer                 := nil;
  FTOBSauveLigne                 := nil;
  FInfoEcr                       := nil;
  FMessCompta                    := nil ;
  SD                             := nil ;
  FListeMessage                  := nil ;
 end;

 // on rebranche l'evenement par defaut
 Application.OnException := FExceptionEvent;

 Inherited ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 08/11/2001
Modifié le ... :   /  /
Description .. : Gestion des exceptions globales pour la fiche CRELBQE
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.GestionException ( Sender : TObject ; E : Exception);
begin
 MessageAlerte( 'Fermer la fenêtre ! ' + #13#10 + #13#10 + TheLastError + #13#10 + #13#10 +
               E.Message );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/09/2006
Modifié le ... :   /  /    
Description .. : - FB 18281 - LG  - on deactive le bouclag , plante sur la 
Suite ........ : generation de la TVA
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.ChargeParam;
begin

{$IFDEF TT}
AddEvenement('ChargeParam' );
{$ENDIF}

 // recuperation du numero de piece max en base
 if VarIsEmpty(GetParamSocSecur('SO_CPTRESONUMPIECEMAX',0)) then
  FRdNumeroPieceMax := 0
  else
   FRdNumeroPieceMax := GetParamSocSecur('SO_CPTRESONUMPIECEMAX',0);

 FRdNumeroPieceMaxOld := FRdNumeroPieceMax;

// if Ecran <> nil then
//  ChcBouclage.Checked := GetParamSocSecur('SO_BOUCLERSAISIECREAT',False);
 if Ecran <> nil then
  ChcBouclage.Checked := false ;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 20/07/2001
Description .. : Recupération des contrôles de la fiche
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.GetControlTOF: boolean;
var Composants : TControlFiltre;
begin
 PgcCritere              := TPageControl(GetControl('FE__PAGECONTROL_CRITERE'));
 TbsStandard             := TTabSheet(GetControl('FE__TABSHEET_STANDARD'));
 HGSaisie                := THGrid(GetControl('FE__HGRID_SAISIE'));
 HGImputation            := THGrid(GetControl('FE__HGRID_IMPUTATION'));
 PFiltres                := TToolWindow97(GetControl('PFILTRES'));
 PBouton                 := TToolWindow97(GetControl('PBouton'));
 DckDock                 := TDock97(GetControl('Dock'));
 DckDock971              := TDock97(GetControl('Dock971'));
 BInsert                 := TToolbarButton97(GetControl('BInsert')) ;
 BDelete                 := TToolbarButton97(GetControl('BDelete')) ;
 BCherche                := TToolbarButton97(GetControl('BCherche')) ;
 BValider                := TToolbarButton97(GetControl('BValider')) ;
 BFermer2                := TToolbarButton97(GetControl('BMODEOFF')) ;
// BFerme                  := TToolbarButton97(GetControl('BFERME')) ;
 BMode                   := TToolbarButton97(GetControl('BMode')) ;
 BImprimer               := TToolbarButton97(GetControl('BImprimer')) ;
 BDefaire                := TToolbarButton97(GetControl('BDefaire')) ;
 BAgrandir               := TToolbarButton97(GetControl('BAgrandir')) ;
 BMenuZoom               := TToolbarButton97(GetControl('BMenuZoom')) ;
 EdtMulDateDu            := THCritMaskEdit(GetControl('FE__HEDITDATE'));
 EdtMulDateAu            := THCritMaskEdit(GetControl('FE__HEDITDATE_'));
 EdtMulNumeroPiece       := THCritMaskEdit(GetControl('FE__HEDIT_NUMEROPIECE'));
 EdtMulNumeroPieceCpml   := THCritMaskEdit(GetControl('FE__HEDIT_NUMEROPIECECPML'));
 CmbMulEtat              := THMultiValComboBox(GetControl('CRL_ETAT'));
 EdtMulLibelle           := THCritMaskEdit(GetControl('FE__HEDIT_LIBELLE'));
 EdtMulGeneral           := THCritMaskEdit(GetControl('FE__HEDIT_GENERAL'));
 CmbMulValeur1           := THValComboBox(GetControl('FE__HVALCOMBOBOX_VALEUR1'));
 CmbMulValeur2           := THValComboBox(GetControl('FE__HVALCOMBOBOX_VALEUR2'));
 CmbMulChoix1            := THValComboBox(GetControl('FE__HVALCOMBOBOX_CHOIX1'));
 CmbMulChoix2            := THValComboBox(GetControl('FE__HVALCOMBOBOX_CHOIX2'));
 EdtMulValeur1           := THCritMaskEdit(GetControl('FE__HEDIT_VALEUR1'));
 EdtMulValeur2           := THCritMaskEdit(GetControl('FE__HEDIT_VALEUR2'));
 CmbMulOp                := THValComboBox(GetControl('FE__HVALCOMBOBOX_OP'));
 CmbEtablissement        := THValComboBox(GetControl('FE__HVALCOMBOBOX_ETABLISSEMENT'));
 CmbContrepartie         := THValComboBox(GetControl('FE__HVALCOMBOBOX_CONTREPARTIE'));
 CmbJournal              := THValComboBox(GetControl('FE__HVALCOMBOBOX_JOURNAL'));
 HlbDebitSaisie          := THLabel(GetControl('FE__HLABEL_DEBIT_SAISIE'));
 HlbDebitImput           := THLabel(GetControl('FE__HLABEL_DEBIT_IMPUT'));
 HlbCreditSaisie         := THLabel(GetControl('FE__HLABEL_CREDIT_SAISIE'));
 HlbCreditImput          := THLabel(GetControl('FE__HLABEL_CREDIT_IMPUT'));
 BvlDebitSaisie          := TBevel(GetControl('FE__BEVEL_DEBIT_SAISIE'));
 BvlCreditSaisie         := TBevel(GetControl('FE__BEVEL_CREDIT_SAISIE'));
 BvlDebitImput           := TBevel(GetControl('FE__BEVEL_DEBIT_IMPUT'));
 BvlCreditImput          := TBevel(GetControl('FE__BEVEL_CREDIT_IMPUT'));
 BvlSoldeImput           := TBevel(GetControl('FE__BEVEL_SOLDE_IMPUT'));
 HlbSoldeImput           := THLabel(GetControl('FE__HLABEL_SOLDE_IMPUT'));
 EdtDate                 := THCritMaskEdit(GetControl('FE__HEDIT_DATE'));
 HlbDevise               := THLabel(GetControl('FE__HLABEL_DEVISE'));
 HlbSoldeCompte          := THLabel(GetControl('FE__HLABEL_SOLDE_CALC'));
 HlbLabelSoldeCompte     := THLabel(GetControl('FE__HLABEL_SOLDE'));
 BvlSoldeCompte          := TBevel(GetControl('FE__BEVEL_SOLDE'));
 ChcBouclage             := TCheckBox(GetControl('FE__CHECKBOX_BOUCLAGE'));

 EdtFFiltres             := THValComboBox(GetControl('FFiltres'));
 BFiltre                 := TToolBarButton97(GetControl('BFILTRE', True));

 POPF                    := TPopupMenu(GetControl('POPF'));
 POPS                    := TPopupMenu(GetControl('POPS'));
 POPZOOM                 := TPopupMenu(GetControl('POPZOOM'));
 HlbNomGuide             := THLabel(GetControl('FE__HLABEL_NOMGUIDE'));
 ImgEuro                 := TImage(GetControl('FE__IMAGE_EURO'));
 HlbCaptionSoldeImput    := THLabel(GetControl('FE__HLABEL_SOLDEIMPUT'));
 BIntegrer               := TToolbarButton97(GetControl('BIntegrer')) ;
 BDeSelectionner         := TToolbarButton97(GetControl('BDETAG')) ;
 BSelectionner           := TToolbarButton97(GetControl('BTAG')) ;
 BCompte                 := TToolbarButton97(GetControl('BCompte'));
 BGuide                  := TToolbarButton97(GetControl('BGuide'));
 BExport                 := TToolbarButton97(GetControl('BEXPORT'));
 BImport                 := TToolbarButton97(GetControl('BIMPORT'));
 BCreerGuide             := TToolbarButton97(GetControl('BCREERGUIDE')) ;
 HSystemMenu1            := TFVierge(Ecran).HMTrad;

 NotifyErrorComponent ( PgcCritere                 , 'FE__PAGECONTROL_CRITERE');
 NotifyErrorComponent ( TbsStandard                , 'FE__TABSHEET_STANDARD');
 NotifyErrorComponent ( HGSaisie                   , 'FE__HGRID_SAISIE');
 NotifyErrorComponent ( HGImputation               , 'FE__HGRID_IMPUTATION');
 NotifyErrorComponent ( PFiltres                   , 'PFILTRES');
 NotifyErrorComponent ( DckDock                    , 'Dock');
 NotifyErrorComponent ( DckDock971                 , 'Dock971');
 NotifyErrorComponent ( PBouton                    , 'PBouton');
 NotifyErrorComponent ( BInsert                    , 'BInsert');
 NotifyErrorComponent ( BDelete                    , 'BDelete');
 NotifyErrorComponent ( BCherche                   , 'BCherche');
 NotifyErrorComponent ( BValider                   , 'BValider');
 NotifyErrorComponent ( BImprimer                  , 'BImprimer');
 NotifyErrorComponent ( BDefaire                   , 'BDefaire');
 NotifyErrorComponent ( BAgrandir                  , 'BAgrandir');
 NotifyErrorComponent ( BFermer2                   , 'BAnnuler');
 NotifyErrorComponent ( BMenuZoom                  , 'BMenuZoom');
 NotifyErrorComponent ( BMode                      , 'BMode');
 NotifyErrorComponent ( EdtMulDateDu               , 'FE__HEDITDATE');
 NotifyErrorComponent ( EdtMulDateAu               , 'FE__HEDITDATE_');
 NotifyErrorComponent ( EdtMulNumeroPiece          , 'FE__HEDIT_NUMEROPIECE');
 NotifyErrorComponent ( EdtMulNumeroPieceCpml      , 'FE__HEDIT_NUMEROPIECECPML');
 NotifyErrorComponent ( CmbMulEtat                 , 'CRL_ETAT');
 NotifyErrorComponent ( EdtMulLibelle              , 'FE__HEDIT_LIBELLE');
 NotifyErrorComponent ( EdtMulValeur1              , 'FE__HEDIT_VALEUR1');
 NotifyErrorComponent ( EdtMulValeur2              , 'FE__HEDIT_VALEUR2');
 NotifyErrorComponent ( CmbEtablissement           , 'FE__HEDIT_ETABLISSEMENT');
 NotifyErrorComponent ( CmbContrepartie            , 'FE__HEDIT_CONTREPARTIE');
 NotifyErrorComponent ( CmbJournal                 , 'FE__HEDIT_JOURNAL');
 NotifyErrorComponent ( HlbDebitSaisie             , 'FE__HLABEL_DEBIT_SAISIE');
 NotifyErrorComponent ( HlbDebitImput              , 'FE__HLABEL_DEBIT_IMPUT');
 NotifyErrorComponent ( HlbCreditSaisie            , 'FE__HLABEL_CREDIT_SAISIE');
 NotifyErrorComponent ( HlbCreditImput             , 'FE__HLABEL_CREDIT_IMPUT');
 NotifyErrorComponent ( BvlDebitSaisie             , 'FE__BEVEL_DEBIT_SAISIE');
 NotifyErrorComponent ( BvlCreditSaisie            , 'FE__BEVEL_CREDIT_SAISIE');
 NotifyErrorComponent ( BvlDebitImput              , 'FE__BEVEL_DEBIT_IMPUT');
 NotifyErrorComponent ( BvlCreditImput             , 'FE__BEVEL_CREDIT_IMPUT');
 NotifyErrorComponent ( BvlSoldeImput              , 'FE__BEVEL_SOLDE_IMPUT');
 NotifyErrorComponent ( HlbSoldeImput              , 'FE__HLABEL_SOLDE_IMPUT');
 NotifyErrorComponent ( EdtDate                    , 'FE__HEDIT_DATE');
 NotifyErrorComponent ( HlbDevise                  , 'FE__HLABEL_DEVISE');
 NotifyErrorComponent ( HlbSoldeCompte             , 'FE__HLABEL_SOLDE_CALC');
 NotifyErrorComponent ( HlbLabelSoldeCompte        , 'FE__HLABEL_SOLDE');
 NotifyErrorComponent ( BvlDebitSaisie             , 'FE__BEVEL_SOLDE');
 NotifyErrorComponent ( ChcBouclage                , 'FE__CHECKBOX_BOUCLAGE');
 NotifyErrorComponent ( EdtFFiltres                , 'FFiltres');
 NotifyErrorComponent ( POPF                       , 'POPF');
 NotifyErrorComponent ( POPS                       , 'POPS');
 NotifyErrorComponent ( POPZOOM                    , 'POPZOOM');
 NotifyErrorComponent ( HlbNomGuide                , 'FE__HLABEL_NOMGUIDE');
 NotifyErrorComponent ( ImgEuro                    , 'FE__IMAGE_EURO');
 NotifyErrorComponent ( HlbCaptionSoldeImput       , 'FE__HLABEL_SOLDEIMPUT');
 NotifyErrorComponent ( BIntegrer                  , 'BINTEGRER');
 NotifyErrorComponent ( BDeSelectionner            , 'BTAG');
 NotifyErrorComponent ( BSelectionner              , 'BDETAG');
 NotifyErrorComponent ( BGuide                     , 'BGuide');
 NotifyErrorComponent ( BCompte                    , 'BCompte');

 // GCO - 07/06/2006
 Composants.Filtres  := EdtFFILTRES;
 Composants.Filtre   := BFILTRE;
 Composants.PageCtrl := PgcCritere;
 FObjFiltre := TObjFiltre.Create(Composants, '');
 //FObjFiltre.ApresChangementFiltre := ApresChargementFiltre;
 // FIN GCO

 Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06i/2001
Modifié le ... : 20/07/2001
Description .. : Création des objets de la TOF
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.CreateControl: boolean;
begin
  FInfoEcr                          := TInfoEcriture.Create ;
  FInfoEcr.Devise.load([V_PGI.DevisePivot]);
  FZEtebac                          := TZEtebac.Create(FInfoEcr);
  FZEtebac.OnError                  := OnError ;
  FZEtebac.InitializeObject;
  FZEtebacGuide                     := TZEtebacGuide.Create(FInfoEcr);
  FListeIntegrer                    := TStringList.Create;
  FCurrentCol                       := 0;
  FCurrentRow                       := 0;
  if Ecran <> nil then
   FMessCompta                       := TMessageCompta.Create(Ecran.Caption)
    else
     FMessCompta                       := TMessageCompta.Create('') ;
  if Ecran <> nil then
   begin
    SD                                := TSaveDialog.Create(Ecran);
    SD.Filter                         := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
    SD.DefaultExt                     := 'XLS';
    SD.FilterIndex                    := 1;
    SD.Options                        := SD.Options + [ofOverwritePrompt, ofPathMustExist, ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];
  end ;
  result                            := true;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 16/01/2006
Description .. : Initialisation des contrôles
Suite ........ : FB 10174 - LG  - 13/06/2005 - ajout du code interbancaire
Suite ........ : - FB 15735 - LG - 14/06/2005 - plantait qd on 
Suite ........ : redimensionnait la fenetre
Suite ........ : - FB 16822 - LG  - 16/01/2005 - on ne peut plus saisir ds le 
Suite ........ : numero
Mots clefs ... : 
*****************************************************************}
function TOF_SAISIEETEBAC.InitControl : boolean;
var
 lDtDateCompable1,lDtDateCompable2 : TDateTime ;
begin

  ChargeParam;

  // parametrage de la grille de saisie des relevées
  HGSaisie.ColWidths[0]                          := 30;
  HGSaisie.ColWidths[cColDate]                   := 75;
  HGSaisie.ColWidths[cColRefInterne]             := 80;
  HGSaisie.ColWidths[cColDebit]                  := 70;
  HGSaisie.ColWidths[cColCredit]                 := 70;
  HGSaisie.ColWidths[cColLibelle]                := 150;
  HGSaisie.ColWidths[cColPiece]                  := 100;
  HGSaisie.ColWidths[cColImputation]             := 100;
  HGSaisie.ColWidths[cColTva]                    := 35;
  HGSaisie.ColWidths[cColEtat]                   := 30;
  HGSaisie.ColWidths[cColAFB]                    := 25;

  HGSaisie.Cells[cColDate,0]                     := TraduireMemoire('Date');
  HGSaisie.Cells[cColRefInterne,0]               := TraduireMemoire('Pièce banc.');
  HGSaisie.Cells[cColDebit,0]                    := TraduireMemoire('Débit');
  HGSaisie.Cells[cColCredit,0]                   := TraduireMemoire('Crédit');
  HGSaisie.Cells[cColLibelle,0]                  := TraduireMemoire('Libellé');
  HGSaisie.Cells[cColPiece,0]                    := TraduireMemoire('Pièce');
  HGSaisie.Cells[cColImputation,0]               := TraduireMemoire('Imputation');
  HGSaisie.Cells[cColTva,0]                      := TraduireMemoire('T');
  HGSaisie.Cells[cColEtat,0]                     := TraduireMemoire('Etat');
  HGSaisie.Cells[cColAFB,0]                      := TraduireMemoire('AFB');
  // limitation de la longeur des champs
  HGSaisie.ColLengths[cColLibelle]               := 35;
  HGSaisie.ColLengths[cColRefInterne]            := 17;
  HGSaisie.ColLengths[cColPiece]                 := 17;
  HGSaisie.ColLengths[cColImputation]            := VH^.Cpta[fbGene].Lg ; // longeur d'un compte general

  // formatage de la cellule Date
  HGSaisie.ColTypes[cColDate]                    := 'D';
  HGSaisie.ColFormats[cColDate]                  := ShortDateFormat;

  HGSaisie.ColTypes[cColDebit]                   := 'R';
  HGSaisie.ColAligns[cColDebit]                  := taRightJustify ;
  HGSaisie.ColFormats[cColDebit]                 := '#,##0.00';
  HGSaisie.ColTypes[cColCredit]                  := 'R';
  HGSaisie.ColFormats[cColCredit]                := '#,##0.00';
  HGSaisie.ColAligns[cColCredit]                 := taRightJustify ;
  HGSaisie.ColTypes[cColTva]                     := 'R';
  HGSaisie.ColFormats[cColTva]                   := '#0.00';
  HGSaisie.ColAligns[cColTva]                    := taRightJustify ;
  HGSaisie.ColAligns[0]                          := taCenter ;
  {$IFNDEF TT}
  HGSaisie.ColEditables[cColAFB]                 := false ;
  {$ENDIF}
  HGSaisie.ColEditables[0]                       := false ;

  // redimensionnement de la grille
  HSystemMenu1.ResizeGridColumns(HGSaisie);

  // placement des soldes des débits et crédits
  BvlDebitSaisie.Left                            := HGSaisie.LeftCol + HGSaisie.ColWidths[0] + HGSaisie.ColWidths[cColDate] + HGSaisie.ColWidths[cColRefInterne] +
                                                    + HGSaisie.ColWidths[cColLibelle] + 5;
  BvlDebitSaisie.Width                           := HGSaisie.ColWidths[cColDebit];
  HlbDebitSaisie.Left                            := BvlDebitSaisie.Left + 3;
  HlbDebitSaisie.Width                           := HGSaisie.ColWidths[cColDebit] - 6;

  BvlCreditSaisie.Left                           := BvlDebitSaisie.Left + HGSaisie.ColWidths[cColDebit] + 5;
  BvlCreditSaisie.Width                          := HGSaisie.ColWidths[cColCredit];
  HlbCreditSaisie.Left                           := BvlCreditSaisie.Left + 3;
  HlbCreditSaisie.Width                          := HGSaisie.ColWidths[cColCredit] - 6;

  HlbLabelSoldeCompte.Left                       := BvlCreditSaisie.Left + BvlDebitSaisie.Width + 10;
  BvlSoldeCompte.Left                            := HlbLabelSoldeCompte.Left + HlbLabelSoldeCompte.Width + 2;
  HlbSoldeCompte.left                            := BvlSoldeCompte.Left ;
  HlbSoldeCompte.Width                           := ( HlbLabelSoldeCompte.Width * 2 ) - 3;
  BvlSoldeCompte.Width                           := HlbLabelSoldeCompte.Width * 2;

   // parametrage de la grille de saisie des imputations
  HGImputation.ColWidths[0]                      := 10;
  HGImputation.ColWidths[cColGeneralImput]       := 100;
  HGImputation.ColWidths[cColAuxiliaireImput]    := 100;
  HGImputation.ColWidths[cColLibelleImput]       := 200;
  HGImputation.ColWidths[cColReferenceImput]     := 70;
  HGImputation.ColWidths[cColCreditImput]        := 70;
  HGImputation.ColWidths[cColDebitImput]         := 70;


  HGImputation.Cells[cColGeneralImput,0]         := TraduireMemoire('Général');
  HGImputation.Cells[cColAuxiliaireImput,0]      := TraduireMemoire('Auxiliaire');
  HGImputation.Cells[cColReferenceImput,0]       := TraduireMemoire('Référence');
  HGImputation.Cells[cColLibelleImput,0]         := TraduireMemoire('Libellé');
  HGImputation.Cells[cColDebitImput,0]           := TraduireMemoire('Débit');
  HGImputation.Cells[cColCreditImput,0]          := TraduireMemoire('Crédit');

  HGImputation.ColLengths[cColLibelleImput]      := 35;
  HGImputation.ColLengths[cColGeneralImput]      := VH^.Cpta[fbGene].Lg ; // longeur d'un compte general
  HGImputation.ColLengths[cColAuxiliaireImput]   := VH^.Cpta[fbGene].Lg ; // longeur d'un compte general

  // formatage des cellules
  HGImputation.ColTypes[cColDebitImput]          := 'F';
  HGImputation.ColAligns[cColDebitImput]         := taRightJustify ;
  HGImputation.ColFormats[cColDebitImput]        := '#,##0.00';
  HGImputation.ColTypes[cColCreditImput]         := 'F';
  HGImputation.ColFormats[cColCreditImput]       := '#,##0.00';
  HGImputation.ColAligns[cColCreditImput]        := taRightJustify ;
  HGImputation.ColAligns[0]                      := taCenter ;

  // redimensionnement de la grille
  HSystemMenu1.ResizeGridColumns(HGImputation);

  // placement des totaux des debti et credit
  BvlDebitImput.Left                             := HGImputation.LeftCol + HGImputation.ColWidths[0] + HGImputation.ColWidths[cColGeneralImput]
                                                 + HGImputation.ColWidths[cColAuxiliaireImput] + HGImputation.ColWidths[cColLibelleImput]
                                                 + HGImputation.ColWidths[cColReferenceImput] + 5;
  BvlDebitImput.Width                            := HGImputation.ColWidths[cColDebitImput];
  HlbDebitImput.Left                             := BvlDebitImput.Left + 3;
  HlbDebitImput.Width                            := HGImputation.ColWidths[cColDebitImput] - 6;

  BvlCreditImput.Left                            := BvlDebitImput.Left + HGImputation.ColWidths[cColDebitImput] + 5;
  BvlCreditImput.Width                           := HGImputation.ColWidths[cColCreditImput];
  HlbCreditImput.Left                            := BvlCreditImput.Left + 3;
  HlbCreditImput.Width                           := HGImputation.ColWidths[cColCreditImput] - 6;

  BvlSoldeImput.Left                             := BvlCreditImput.Left;
  BvlSoldeImput.Width                            := BvlCreditImput.Width;
  HlbSoldeImput.Left                             := BvlCreditImput.Left + 3;
  HlbSoldeImput.Width                            := HGImputation.ColWidths[cColCreditImput] - 6;

  HlbCaptionSoldeImput.Left                      := BvlDebitImput.Left + 40;
  ImgEuro.Left                                   := BvlDebitImput.Left;

  // mise en place de la ToolBar
  PFiltres.Left                                  := PBouton.Left + 70 ;
  BFiltre.Visible                                := false ;
  EdtFFiltres.Visible                            := false ;
  BCherche.Left                                  := PBouton.Left + 10 ;
//  PFiltres.Visible                               := false ;

  StatutGrilleReleve                             := taConsult;
  StatutGrilleImput                              := taConsult;

  EdtMulNumeroPiece.Text                         := '';
  EdtMulNumeroPieceCpml.Text                     := '';
  EdtMulLibelle.Text                             := '';
  EdtMulGeneral.Text                             := '';

  // affectation du libellé etablissement par défaut
  CmbEtablissement.Value                         := VH^.EtablisDefaut;

  HlbNomGuide.Left                               := 20;
  HlbNomGuide.Top                                := 7;

    // affiche la devise du dossier a cote de la case solde des credits
  HlbDevise.Caption                              := TraduireMemoire('Dossier tenu en ') + VH^.LibDevisePivot;
  HlbDevise.Left                                 := HlbNomGuide.Left;
  HlbDevise.Top                                  := HlbNomGuide.Top + 30;
  HlbDevise.Visible                              := false;

  FlashGuide                                     := TFlashingLabel.Create(Ecran);
  FlashGuide.parent                              := HlbDebitImput.Parent;
  FlashGuide.Top                                 := HlbNomGuide.Top - 3;
  FlashGuide.Left                                := HlbNomGuide.Left + HlbNomGuide.Width + 50;
  FlashGuide.Font.Height                         := 20;
  FlashGuide.Font.Style                          := [fsBold];
  FlashGuide.Transparent                         := true;
  FlashGuide.Caption                             := TraduireMemoire('GUIDE');
  FlashGuide.Color                               := clRed;
  FlashGuide.Visible                             := false;


  FlashCurrentReleve                             := TFlashingLabel.Create(Ecran);
  FlashCurrentReleve.parent                      := HlbDebitSaisie.Parent;
  FlashCurrentReleve.Top                         := 5;
  FlashCurrentReleve.Left                        := 10;
  FlashCurrentReleve.Font.Height                 := 15;
  FlashCurrentReleve.Font.Style                  := [fsBold];
  FlashCurrentReleve.Transparent                 := true;
  FlashCurrentReleve.Caption                     := TraduireMemoire(cStTexteLigneReleve);
  FlashCurrentReleve.Color                       := clActiveCaption; 
  FlashCurrentReleve.Visible                     := true;


  FlashCompte                                     := TFlashingLabel.Create(Ecran);
  FlashCompte.parent                              := CmbJournal.Parent;
  FlashCompte.Top                                 := 11;
  FlashCompte.Left                                := 7;
  FlashCompte.Font.Height                         := 13;
  FlashCompte.Transparent                         := true;
  FlashCompte.Caption                             := TraduireMemoire('Compte bancaire');
  FlashCompte.Color                               := clWindowText;
  FlashCompte.Visible                             := true;

  FlashJournal                                    := TFlashingLabel.Create(Ecran);
  FlashJournal.parent                             := CmbJournal.Parent;
  FlashJournal.Top                                := 11;
  FlashJournal.Left                               := CmbContrepartie.Left + CmbContrepartie.Width + 10;
  FlashJournal.Font.Height                        := 13;
  FlashJournal.Transparent                        := true;
  FlashJournal.Caption                            := TraduireMemoire('Journal');
  FlashJournal.Color                              := clWindowText;
  FlashJournal.Visible                            := true;

   // charge les filtres par defaut
 //ChargeFiltre('RLB_SAISIE',EdtFFiltres,PgcCritere);


 CInitialiseDateEtebac(lDtDateCompable1,lDtDateCompable2);
 EdtMulDateDu.Text                                := DateToStr(lDtDateCompable1);
 EdtMulDateAu.Text                                := DateToStr(lDtDateCompable2);

 Ecran.PopupMenu                                  := POPS;

 PgcCritere.ActivePageIndex                       := 0;

 FBoGrilleVide                                    := true;

// FBoModeGuideAuto                                 := false;
 FBoSorted                                        := false;
 BExport.Visible                                  := ExJaiLeDroitConcept(ccExportListe, False);


 TToolbarButton97(GetControl('BF6')).OnClick       := OnClickF6 ;
 TToolbarButton97(GetControl('BAltD')).OnClick     := OnClickAltD ;

 CmbEtablissement.Vide                             := true ;
 CmbEtablissement.ReLoad ; // LG - 27/08/2007 - le .vide n'etait pas pris en compte

 {$IFNDEF TT}
 BImport.Visible                                   := not ( ctxPCL in V_PGI.PGIContexte ) ;
 {$ENDIF}
 BImport.Hint                                      := 'Importation des fichiers Etebac' ;

 result                                            := true;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Affectation des evenements au contrôle
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.AssignEvent: boolean;
begin 
 HGSaisie.OnCellEnter           := HGSaisieCellEnter;    // affichage du Bouton sur les col. Imputation et Date
 HGSaisie.OnElipsisClick        := HGSaisieElipsisClick; // Click sur Bouton dans les col. Imputation et Date
 HGSaisie.OnCellExit            := HGSaisieCellExit;     // Contrôle des valeurs saisies
 HGSaisie.PostDrawCell          := HGSaisiePostDrawCell; // Gestion des grisé sur les colonnes Debit et Credit
 HGSaisie.OnRowExit             := HGSaisieRowExit;      // enregistrement de la ligne
 HGSaisie.OnKeyPress            := HGSaisieKeyPress;
 HGSaisie.OnKeyDown             := HGSaisieKeyDown;
 HGSaisie.OnRowEnter            := HGSaisieRowEnter;
 HGSaisie.GetCellCanvas         := HGSaisieGetCellCanvas;
 HGSaisie.OnClick               := HGSaisieClick;
 HGSaisie.OnSorted              := HGSaisieOnSorted ;

 HGImputation.PostDrawCell      := HGImputationPostDrawCell;
 HGImputation.OnCellEnter       := HGImputationCellEnter;    // affichage du Bouton sur les col. Imputation et Date
 HGImputation.OnCellExit        := HGImputationCellExit;
 HGImputation.OnElipsisClick    := HGImputationElipsisClick;
 HGImputation.OnKeyPress        := HGImputationKeyPress;
 HGImputation.OnKeyDown         := HGImputationKeyDown;
 HGImputation.OnExit            := HGImputationExit;
 HGImputation.OnRowExit         := HGImputationRowExit;      // enregistrement de la ligne
 HGImputation.GetCellCanvas     := HGImputationGetCellCanvas;
 HGImputation.OnRowEnter        := HGImputationRowEnter;

 BCompte.OnClick                 := BCompteClick;
 BGuide.OnClick                  := BGuideClick;
 BExport.OnClick                 := BExportClick ;
 BImport.OnClick                 := BImportClick ;
 BCreerGuide.OnClick             := BCreerGuideClick ; ;

 //EdtFFiltres.OnChange            := FFiltresChange;

 POPS.OnPopup                    := POPSPopup;

 CmbContrepartie.OnChange        := CmbContrepartieChange;
 CmbJournal.OnChange             := CmbJournalChange;
 CmbEtablissement.OnChange       := CmbEtablissementChange ;

 Ecran.OnKeyDown                 := FormKeyDown;
 Ecran.OnCloseQuery              := FormCloseQuery;

 // affectation des evenements sur les bp
 BDelete.OnClick                 := BDeleteClick;
 BCherche.OnClick                := BChercheClick;
 BValider.OnClick                := BValiderClick;
 BImprimer.OnClick               := BImprimerClick;
 BDefaire.OnClick                := BDefaireClick;
 BAgrandir.OnClick               := BAgrandirClick;
 BInsert.OnClick                 := BInsertClick;
 BFermer2.OnClick                := BFermerClick;
// BFermer.OnClick                 := BFermeClick;
 BMode.OnClick                   := BModeClick;
 BINTEGRER.OnClick               := BIntegrerClick;
 BSelectionner.OnClick           := BSelectionnerClick;
 BDeSelectionner.OnClick         := BDeSelectionnerClick;
 BMenuZoom.OnMouseEnter          := BMenuZoomMouseEnter;

 EdtMulDateDu.OnChange           := EdtMulDateDuChange;
 EdtMulDateDu.OnExit             := EdtMulDateDuExit;
 EdtMulDateAu.OnChange           := EdtMulDateAuChange;
 EdtMulDateAu.OnExit             := EdtMulDateAuExit;

 Result := true;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Initialise une nouvelle ligne dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieInitLigne( value : integer );
begin
 HGSaisie.Cells[cColDate,value]             := DateToStr(FZEtebac.DtDateDefaut);
 HGSaisie.Cells[cColRefInterne,value]       := '';
 HGSaisie.Cells[cColDebit,value]            := '';
 HGSaisie.Cells[cColCredit,value]           := '';
 HGSaisie.Cells[cColLibelle,value]          := '';
 HGSaisie.Cells[cColPiece,value]            := '';
 HGSaisie.Cells[cColImputation,value]       := '' ;
 HGSaisie.Cells[cColTva,value]              := '';
 HGSaisie.CellValues[cColEtat,value]        := FZEtebac.StETAT_R;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Ajout d'une nouvelle ligne dans la grille et la TOB
Suite ........ : TOB_CRELBQE
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.CreateRow;
var
 lBoInutilise : boolean;
begin

 if ( StatutGrilleReleve <> taConsult )  then exit; // on ajout une ligne uniquement si on etait en consultation

{$IFDEF TT}
 AddEvenement('CreateRow' );
{$ENDIF}

 SetEvent(false); // déactive tous les evenements pour crée une nouvelle enregistrement

 // on vide la grille des imputations
 VideGrille(HGImputation);

 // on rajoute une ligne si : on est sur la première ligne et elle est valide ou le nombre de ligne est supérieur à 1
 if ( ( HGSaisie.Row = 1 ) and HGSaisieIsRowValid(HGSaisie.Row ) ) or ( HGSaisie.Row > 1 ) then
  HGSaisie.RowCount := HGSaisie.RowCount + 1 ;

 if HGSaisie.CanFocus then // on donne le focus à la grille
  HGSaisie.setfocus;
 HGSaisie.Row                            := HGSaisie.RowCount - 1;          // on se place sur cette nouvelle ligne
 HGSaisie.Objects[0,HGSaisie.Row]        := FZEtebac.AddLigneReleve;   // creation de l'enregistrement et ajout à la grille
 //FZEtebac.InCOMPTEUR_R              := FZEtebac.ProchainCompteur; // initialisation du compteur de ligne
 HGSaisieInitLigne(HGSaisie.RowCount - 1);                                  // Initialiation de la grille
 StatutGrilleReleve                      := taCreat;                        // on passe en creation dans la grille
                                                                            // on recalcul de solde
 RafraichirLigne;
 HGImputation.Enabled := false;

 FBoGrilleVide := HGSaisie.RowCount < 3;

 BValider.Enabled := false;

 SetEvent(true); // on réactive les évènements

 // on declenche manuellement l'eve. pour assigner la TOB
 HGSaisieRowEnter( Self, HGSaisie.Row, lBoInutilise,lBoInutilise );
 HGSaisie.col := GetCol ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 08/11/2001
Modifié le ... :   /  /
Description .. : Insertion d'un nouvelle enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.BInsertClick(Sender: TObject);
begin
 CreateRow;
end;

procedure TOF_SAISIEETEBAC.BDeleteClick(Sender: TObject);
begin
 HGSaisieDeleteRow(HGSaisie.Row);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 16/01/2006
Modifié le ... :   /  /    
Description .. : - LG - 17052 - FB 17052 - on lance un vk_f10 pour valider 
Suite ........ : l'ecriture
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.BValiderClick(Sender: TObject);
begin
 SendMessage(Ecran.Handle,WM_KEYDOWN,VK_F10,0) ;
end;

procedure TOF_SAISIEETEBAC.BImprimerClick(Sender: TObject);
begin
 {$IFNDEF EAGLCLIENT}
 PrintDBGrid( HGSaisie , Nil , 'Liste des imputations','') ;
 {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 08/11/2001
Modifié le ... :   /  /
Description .. : Bouton d'intégration des écritures en compta
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.BIntegrerClick(Sender: TObject);
begin
 Integrer() ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/09/2007
Modifié le ... :   /  /    
Description .. : - LG  - 13/09/2007 - FB 21415 - correction de l'integration 
Suite ........ : auto depuis le bureau
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.Integrer( vTOBLettrage : TOB = nil );
var
 lZTresoVersEcr : TZTresoVersEcr;
 lStJal         : string ;
 lBoResult      : boolean ;
begin

 if not FBoChainage and ( FListeIntegrer.Count = 0 ) then exit;

// if FBoParle and ( PGIAskCancel( cStTexteConfirmIntegration , cStTexteTitreFenetre ) = mrNo) then

if FBoParle and ( PGIAskCancel( 'Confirmer vous l''intégration des écritures valides ?' , cStTexteTitreFenetre ) = mrNo) then
  begin
   // messaqe de confirmation de l'intégration
   FBoModeSelection := false;
   SetModeSelection;
   BDeSelectionnerClick(nil);
   exit;
  end; // if

if Ecran <> nil then
 lStJal := CmbJournal.Value
  else
   lStJal := FZEtebac.StJournalContrepartie ;
if FInfoEcr.Journal.Load([lStJal]) = - 1 then exit ;
 if FInfoEcr.Journal.GetValue('J_FERME') = 'X' then
  begin
   PGIInfo('Le journal est fermé !',cStTexteTitreFenetre) ;
   exit ;
  end ;

 try

 // creation de l'objet de transfert des ecritures ( definis dans //ZTresoVersEcr.pas )
 lZTresoVersEcr                       := TZTresoVersEcr.Create(FInfoEcr);
 lZTresoVersEcr.OnError               := OnError ;
 // affichage de la fenetre d'attente
 if ( Ecran <> nil ) and FBoParle then
 InitMoveProgressForm ( Ecran,  cStTexteTraitementEcr,
                        cStTexteTitreFenetre,
                        FListeIntegrer.Count - 1,
                        true,
                        true ) ;
 lZTresoVersEcr.StJournalContrepartie := FZEtebac.StJournalContrepartie;
 lZTresoVersEcr.StCompteContrepartie  := FZEtebac.StCompteContrepartie ;
 lZTresoVersEcr.StModeSaisie          := FZEtebac.StModeSaisie;
 lZTresoVersEcr.StCompteurNormal      := FZEtebac.StCompteurNormal;
 if Ecran <> nil then
  lZTresoVersEcr.StEtablissement      := cmbEtablissement.Value
   else
    lZTresoVersEcr.StEtablissement    := VH^.EtablisDefaut ;
 lZTresoVersEcr.PTOBLettrage          := vTOBLettrage ;

 //lZTresoVersEcr.OnError               :=

 try
 // chargement des releve ds l'objet de transfert
 if FBoChainage then
  begin
   lBoResult := lZTresoVersEcr.Transfert(FZEtebac.TOBReleve) ;
   FInNumFolio := - 1 ;
  end
   else
    begin
     FInNumFolio := - 1 ;
     lBoResult := lZTresoVersEcr.LoadReleve( FListeIntegrer ) ;
    end ;

  if lBoResult then
   begin
    // on enregistre les ecritures
    case lZTresoVersEcr.Save(FInNumFolio) of
     oeOk     : if FBoParle and ( Ecran <> nil ) then PGIInfo( cStTexteTransfertReussie , cStTexteTitreFenetre ) ;
     oeSaisie,
     oeStock,
     oeReseau : begin  // erreur sur la mise à jour des comptes ->
                 FBoModeSelection := false;
                 SetModeSelection;
                 if ( Ecran <> nil ) and ( BCherche.Enabled ) then
                  BChercheClick(nil);
                end; // if
    end; // case
   end; // if

 if Ecran <> nil then BDeSelectionnerClick(nil);

 finally
  if assigned(lZTresoVersEcr) then
   lZTresoVersEcr.Free;
  if FBoParle then FiniMoveProgressForm ;
 end; // try

 except
  on E : Exception do
   PGIError('Erreur la génération des écritures !' + #10#13 + E.Message ) ;
 end ;

end;



procedure TOF_SAISIEETEBAC.BMenuZoomMouseEnter(Sender: TObject);
begin
 PopZoom97(BMenuZoom,POPZOOM);
end;

procedure TOF_SAISIEETEBAC.BSelectionnerClick(Sender: TObject);
var
 i   : integer;
begin

 for i := 0 to ( FZEtebac.TOBReleve.Detail.Count - 1 ) do
  begin
   FZEtebac.TOBLigneReleve := FZEtebac.TOBReleve.Detail[i];
   SelectionnerLigne;
  end; // for

 if Ecran <> nil then
  begin
   HGSaisie.AllSelected := true;
   // on remet à jour l'affichage des boutons des fenetres
   EnabledControl;
  end ;

end;

procedure TOF_SAISIEETEBAC.BDeSelectionnerClick(Sender: TObject);
begin
 FListeIntegrer.Clear;
 HGSaisie.ClearSelected;
 EnabledControl;
end;

function TOF_SAISIEETEBAC.SelectionnerLigne : boolean;
var
 lInIndex : integer;
begin

 result   := true;
 lInIndex := FListeIntegrer.IndexOf( varToStr(FZEtebac.InCOMPTEUR_R) );

 if ( lInIndex = -1 ) then
  begin // l'enregistrement n' a jamais été selectionné
   if ( ( FZEtebac.StETAT_R = cImputCorrectNonValide ) or ( FZEtebac.StETAT_R = cImputValide ) ) then
    FListeIntegrer.Add( varToStr(FZEtebac.InCOMPTEUR_R) ) // on l'ajoute à la liste
     else
      result := false; // on ne doit pas pouvoir le selectionner
  end // l'enregistrement n' a jamais été selectionné
   else // l'enresitrement existe dans la liste
    FListeIntegrer.Delete( lInIndex ); // on l'enleve

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/07/2001
Modifié le ... : 20/07/2001
Description .. : Annule les modifs faite dans la grille des releves
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.BDefaireClick(Sender: TObject);
begin
 ReaffecterLigneReleve;
 FRdNumeroPieceMax  := FRdNumeroPieceMaxOld;
 StatutGrilleReleve := taConsult;
end;


procedure TOF_SAISIEETEBAC.BAgrandirClick(Sender: TObject);
begin
// if PgcCritere.Visible then
 PgcCritere.Visible := not PgcCritere.Visible;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/04/2007
Description .. : Gestion du déplacement dans la grille
Suite ........ : - FB 10713 - 10/06/2004 - LG - ajout d'un parametre 
Suite ........ : autorisant la modif ou non des zones montants, date et 
Suite ........ : libelle
Suite ........ : - FB 10714 - 21/06/2005 - LG - on saute la collen code afb
Suite ........ : - FB 19857 - 20/04/2007 - LG - le deplacement ds la cesse 
Suite ........ : credit vers le bas fct mal
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieCellEnter(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean);
var
 lInCurrentCol          : integer;
 lInCurrentRow          : Integer;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

{$IFDEF TT}
 AddEvenement('HGSaisieCellEnter ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow));
{$ENDIF}

 // on stocke la position courante car cette eve est redencher plusieur fois
 lInCurrentCol := HGSaisie.Col;
 lInCurrentRow := HGSaisie.Row;

 // gestion de l'affichage des boutons
 HGSaisie.ElipsisButton := ( lInCurrentCol = cColImputation ) or ( lInCurrentCol = cColDate);
 // gestion des deplacments dans la grille
 case lInCurrentCol of
  cColDebit    : if (( FZEtebac.StIMPORT_R  = 'X' ) and ( FInfoEcr.Journal.GetValue('J_TRESOMONTANT') = '-' )) or ( HGSaisie.Cells[cColCredit,lInCurrentRow] <> '' ) then
                 begin
                  PasseColSuivante(ACol,ARow,HGSaisie);
                  Cancel := true;
                 end;
  cColCredit   : if (( FZEtebac.StIMPORT_R  = 'X' ) and ( FInfoEcr.Journal.GetValue('J_TRESOMONTANT') = '-' )) or ( HGSaisie.Cells[cColDebit,lInCurrentRow] <> '' ) then
                 begin
                  {$IFDEF TT}
                  AddEvenement('PasseColSuivante(ACol,ARow,HGSaisie) credit '+intToStr(ACol)+' ARow:='+intToStr(ARow));
                  AddEvenement('------------------------------------ credit ACol:='+intToStr(HGSaisie.Col)+' ARow:='+intToStr(HGSaisie.Row));
                  {$ENDIF}
                  ACol   := cColPiece ;
                  ARow   := HGSaisie.Row ;
                  Cancel := true;
                 end;
  cColEtat    : PostMessage(HGSaisie.Handle, WM_KEYDOWN, VK_TAB, 0) ;
  cColPiece   : begin
                 if ( HGSaisie.Cells[cColDebit,lInCurrentRow] = '' ) and ( HGSaisie.Cells[cColCredit,lInCurrentRow] <> '' ) then
                  SetGridGrise(cColDebit,lInCurrentRow,HGSaisie)
                   else
                    if ( HGSaisie.Cells[cColDebit,lInCurrentRow] = '' ) and ( HGSaisie.Cells[cColCredit,lInCurrentRow] = '' ) then
                     HGSaisie.InvalidateCell(cColDebit,lInCurrentRow); // on redessine la cellule pour supprimer le grisé
                 end;
  cColLibelle : begin
                 if ( HGSaisie.Cells[cColCredit,lInCurrentRow] = '' ) and ( HGSaisie.Cells[cColDebit,lInCurrentRow] <> '') then
                  SetGridGrise(cColCredit,lInCurrentRow,HGSaisie)
                   else
                    if ( HGSaisie.Cells[cColCredit,lInCurrentRow] = '' ) and ( HGSaisie.Cells[cColDebit,lInCurrentRow] = '' ) then
                     HGSaisie.InvalidateCell(cColCredit,lInCurrentRow);  // on redessine la cellule pour supprimer le grisé
                   if (( FZEtebac.StIMPORT_R  = 'X' ) and ( FInfoEcr.Journal.GetValue('J_TRESOLIBELLE') = '-' )) then
                    begin
                     PasseColSuivante(ACol,ARow,HGSaisie);
                     Cancel := true;
                    end ;
                 end;
  cColTva     : if not FZEtebac.IsTvaAutorise(HGSaisie.Cells[cColImputation,lInCurrentRow]) then
                 begin
                   PasseColSuivante(ACol,ARow,HGSaisie);
                   SetGridGrise(cColTva,lInCurrentCol,HGSaisie);
                   Cancel := true;
                 end;
  cColDate    : if (( FZEtebac.StIMPORT_R  = 'X' ) and ( FInfoEcr.Journal.GetValue('J_TRESODATE') = '-' )) then
                 begin
                  PasseColSuivante(ACol,ARow,HGSaisie);
                  Cancel := true;
                 end ;
  {$IFNDEF TTa}
  cColAFB    : begin
                PasseColSuivante(ACol,ARow,HGSaisie) ;
                {$IFDEF TT}
                AddEvenement('PasseColSuivante(ACol,ARow,HGSaisie) code afb '+intToStr(ACol)+' ARow:='+intToStr(ARow));
                AddEvenement('------------------------------------ code afb ACol:='+intToStr(HGSaisie.Col)+' ARow:='+intToStr(HGSaisie.Row));
                {$ENDIF}
                Cancel := true ;
              end ;
  {$ENDIF}
 end; // case


end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Gestion du clik dasn la cellule imputation
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieElipsisClick(Sender: TObject);
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 case HGSaisie.Col of
  cColImputation : LookupList(HGSaisie,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL',true, 1);
  cColDate       : begin // affichage du calendrier AGL
                    V_PGI.ParamDateProc(HGSaisie);
                    StatutGrilleReleve := taModif;
                   end;
 end; // case
end;

procedure TOF_SAISIEETEBAC.AfficheLookUpImput;
var
 lStGeneral : string;
 lBoResult  : boolean;
begin

 lBoResult    := false;

 lStGeneral   := HGSaisie.Cells[cColImputation,HGSaisie.Row];

 if FInfoEcr.Compte.GetCompte(lStGeneral) > -1 then
  begin
   FZEtebac.StGENERAL_R := lStGeneral;
   FZEtebac.AssignInfoCompte;
   lBoResult                 := true;
  end
   else
    begin
     if HGSaisie.Col <> cColImputation then
      begin
       HGSaisie.SynEnabled := false;
       HGSaisie.Col        := cColImputation;
       HGSaisie.SynEnabled := true;
      end; // if
     if LookupList(HGSaisie,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL',true, 1) then
      begin
       FZEtebac.StGENERAL_R := HGSaisie.Cells[cColImputation,HGSaisie.Row];
       FZEtebac.AssignInfoCompte;
       lBoResult                 := false;
      end; // if
    end; // if

  if lBoResult then
   begin

    // remise à jour des infos
    HGSaisie.Cells[cColImputation,HGSaisie.Row]  := FZEtebac.StGENERAL_R;
    if   ( FZEtebac.StTVASAISIE_R = 'X' ) then
    HGSaisie.Cells[cColTva,HGSaisie.Row]         := STRFMONTANT ( FZEtebac.RdTAUXTVA_R , 15 , V_PGI.OkDecV, '' , true)
      else
       HGSaisie.Cells[cColTva,HGSaisie.Row]      := '';

  end; // if

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... :   /  /
Description .. : Recherche de l'imputation saisie dans le releve et mise à
Suite ........ : jour des info ( libelle + code tva )
Mots clefs ... :
*****************************************************************}
function  TOF_SAISIEETEBAC.HGSaisieRechercheImputation : boolean;
var
 lStGeneral : string;
begin

 lStGeneral  := HGSaisie.Cells[cColImputation,HGSaisie.Row];
 result      := FZEtebac.IsValideGeneralImput(lStGeneral);
 HGSaisie.Cells[cColImputation,HGSaisie.Row] := lStGeneral ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Vérification de la validité des comptes saisies
Suite ........ : Formatage des cellules
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
 RdMontant : double;
begin

 if ( csDestroying in Ecran.ComponentState ) or ( StatutGrilleReleve = taConsult )then Exit ;

{$IFDEF TT}
  AddEvenement('HGSaisieCellExit ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow));
{$ENDIF}

 case ACol of
  cColDate       : begin
                    Cancel := not IsDateValide(HGSaisie.Cells[cColDate,ARow]) ;
                   end;
  cColImputation : if (HGSaisie.Cells[cColImputation,ARow] <> '') then
                    begin
                     cancel :=  not HGSaisieRechercheImputation;
                     AfficheLookUpImput;
                    end;
  cColDebit,
  cColCredit     :  // formatage des nombres
                    if IsNumeric(HGSaisie.Cells[ACol,ARow]) then
                     begin
                      RdMontant := Valeur(HGSaisie.Cells[ACol,ARow]);
                      Cancel    := RdMontant < 0; // on ne peut pas saisir de valeur negative
                      if Cancel then
                       PGIInfo( cStTexteMontantNegatif, cStTexteTitreFenetre );
                      HGSaisie.Cells[ACol,ARow] := STRFMONTANT ( RdMontant , 15 , V_PGI.OkDecV, '' , true);
                     end; // if

  cColTva       : begin
                   if IsNumeric(HGSaisie.Cells[ACol,ARow]) then
                     begin
                      RdMontant := Valeur(HGSaisie.Cells[ACol,ARow]);
                      Cancel    := RdMontant < 0; // on ne peut pas saisir de valeur negative
                      if Cancel then
                       PGIInfo( cStTexteMontantNegatif, cStTexteTitreFenetre );
                      HGSaisie.Cells[ACol,ARow]   := STRFMONTANT ( RdMontant , 15 , V_PGI.OkDecV, '' , true);
                     end; // if
                  end;
 end; // case

 BValider.Enabled := ( HGSaisie.Cells[cColDebit,ARow] <> '' ) or ( HGSaisie.Cells[cColCredit,ARow] <> '' );

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Grise les cellules débit/crédit en fonction de leur état
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisiePostDrawCell( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState );
var
 lBoGrise  : boolean;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 try

 // gestion des cases debit et crédit - une seul des deux doit être renseignée
 lBoGrise := ( ( ACol = cColDebit )  and ( HGSaisie.Cells[cColCredit, ARow] <>'' ) and  ( ARow > 0 ) ) or
             ( ( ACol = cColCredit ) and ( HGSaisie.Cells[cColDebit, ARow]  <>'' ) and  ( ARow > 0 ) ) ; //or
 //            ( ( ACol = cColTva )    and  ( ARow > 0 ) and not FZEtebac.IsTvaAutorise(HGSaisie.Cells[cColImputation,ARow]));

 if lBoGrise then
  begin
   HGSaisie.PostDrawCell  := nil; // on debranche l'évènement lors du dessin de la grille
   SetGridGrise(ACol, ARow, HGSaisie);
   HGSaisie.PostDrawCell  := HGSaisiePostDrawCell;
  end;

 except
  on E : exception do
   begin
     HGSaisie.PostDrawCell := nil;
     PGIBox( 'Erreur lors du dessin de la grille' + #13#10 +
             E.Message ,
             cStTexteTitreFenetre);
     V_PGI.IoError := oeSaisie;
   end; // on

 end;

end ;

procedure TOF_SAISIEETEBAC.HGSaisieClick(Sender: TObject);
begin
 if not FBoChoixEffectue then
  PGIInfo( Ecran.Caption,cStTexteTitreFenetre ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 28/06/2001
Modifié le ... : 17/01/2006
Description .. : Gestion des touches dans les deux grilles
Suite ........ : - FB 16103 - 24/06/2005 - la touche fin place le curseur ds
Suite ........ : la zone tva
Suite ........ : - FB 16552 - 07/09/2005 - ctrl+fin = new line
Suite ........ : - 17/01/2005 - LG - FB 16207 - alt+Z agrandie la fenetre
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
 lBoHGSaisieActive        : boolean;
 lBoHGImputationActive    : boolean;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 lBoHGSaisieActive       := Screen.ActiveControl = HGSaisie;
 lBoHGImputationActive   := Screen.ActiveControl = HGImputation;

 // evenement qui doit être declencher sur tous les contrôles
 if ( Key = VK_F9 ) and BCherche.Enabled then
  BChercheClick(nil);

 if not (  lBoHGSaisieActive or lBoHGImputationActive ) then
  begin

   exit;
  end; // if

 case Key of
  VK_RETURN : begin
               Key:=VK_TAB ;
              end; // VK_RETURN
  VK_SPACE  : if lBoHGSaisieActive and FBoModeSelection then
               begin
                if not SelectionnerLigne then
                 Key := 0;
                EnabledControl;
               end; // if
  VK_INSERT : if lBoHGSaisieActive then
               CreateRow
                else
                 if not FBoGuideRun then
                  NouvelleImputation;
  VK_DOWN   : if lBoHGSaisieActive then
               begin // si la grille des releve est active
                if not EnregistreLigne(HGSaisie.Row) then Key := 0;

                if ( HGSaisie.Row = ( HGSaisie.RowCount - 1 ) )  then
                 CreateRow; // on insere un ligne que sur la dernière cellule
               end;
  VK_TAB    : if lBoHGSaisieActive then
               begin
                if ( HGSaisie.Row = HGSaisie.RowCount - 1 ) and ( HGSaisie.Col = HGSaisie.Colcount - 1 ) then
                 begin
                  if not EnregistreLigne(HGSaisie.Row) then
                   begin
                    Key := 0;
                    exit;
                   end; // if

                  CreateRow;
                 end;  // if

                if ( FZEtebac.StETAT_R <> cPasImput ) and not EnregistreLigne(HGSaisie.Row) then
                 Key := 0;
               end;
  VK_UP     : begin
               // on passe en mode auto pour ne afficher le message
               FZEtebac.TypeContexte := TModeAuto;
               if lBoHGSaisieActive and( StatutGrilleReleve = taCreat ) and not ( HGSaisieIsRowValid(HGSaisie.Row) ) then
                begin
                 HGSaisieDeleteRow(HGSaisie.Row);
                 Key := 0;
                end
                 else
                  if lBoHGSaisieActive and not EnregistreLigne(HGSaisie.Row) then
                   Key := 0;
                FZEtebac.TypeContexte := TModeSaisie;
              end; // if
  VK_ESCAPE : if lBoHGImputationActive then
               begin
                key := 0;
                HGImputationExit(nil);
               end
                else
                 begin
                  if not IsInside(Ecran) then
                   begin
                    Key :=  0 ;
                    PostMessage(Ecran.handle,WM_CLOSE,0,0) ;
                   end ;
                 end ;
          {      else
                 if lBoHGSaisieActive then
                  begin
                   if ( FStatutGrilleReleve <> taConsult ) and (PGIAskCancel(cStTexteEchap,cStTexteTitreFenetre) = mrNo) then
                    begin
                     ReaffecterLigneReleve;
                     Key := 0;
                    end // if
                     else

                  end; // if   }
  VK_DELETE : if Shift = [ssCtrl] then
               begin
                Key := 0 ;
                if  lBoHGSaisieActive then
                 HGSaisieDeleteRow(HGSaisie.Row)
                  else
                   HGImputationDeleteRow(HGImputation.Row);
               end; // VK_DELETE
  VK_F5     : if lBoHGSaisieActive then
               begin
                if Shift = [ssShift] then
                 begin
                  Key := 0;
                  ZoomEcrs;
                 end
                  else
                   HGSaisieElipsisClick(Self);
               end
                else
                 HGImputationElipsisClick(Self);

  Ord('D')  : if ( Shift = [ssAlt] ) and lBoHGSaisieActive then
               begin
                Key := 0;
                OnClickAltD(nil) ;
       //         if not EnregistreLigne(HGSaisie.Row) then exit;
        //        if not AltDImputation                then HGSaisie.SetFocus;
               end;

  Ord('A')  : if Shift = [ssCtrl] then
               begin
                Key := 0;
                if lBoHGSaisieActive and FBoModeSelection then
                 BSelectionnerClick(nil);
               end;  // Ord(A)

  ord('G')  : if Shift = [ssAlt] then BGuideClick(nil);

  VK_F6     : begin
               Key := 0 ;
               OnClickF6(nil) ;
              end ;
               
  VK_F7     : begin
               Key := 0;
               if ( Shift = [ssShift] ) then
                begin // on incremente le numero de référence
                 if lBoHGSaisieActive and ( HGSaisie.Col = cColRefInterne ) then
                  ShiftVK_F7Grille(HGSaisie)
                   else
                    if lBoHGImputationActive and ( StatutGrilleImput <> taConsult ) and ( HGImputation.Col = cColReferenceImput ) then
                     ShiftVK_F7Grille(HGImputation)
                end // if
                 else
                  begin
                   if lBoHGSaisieActive then
                    VK_F7Grille(HGSaisie)
                     else
                      if HGImputation.Row <> ( HGImputation.RowCount - 1 ) then
                       VK_F7Grille(HGImputation);
                  end;
                end;
{CTRL -}109  : if Shift=[ssCtrl]     then
                begin
             if lBoHGSaisieActive and ( HGSaisie.Col = cColRefInterne ) then
                  ShiftVK_F7Grille(HGSaisie,-1)
                   else
                    if lBoHGImputationActive and ( StatutGrilleImput <> taConsult ) and ( HGImputation.Col = cColReferenceImput ) then
                     ShiftVK_F7Grille(HGImputation,-1)
                end; // VK_F7
{CTRL +}107  : if Shift=[ssCtrl] then
                 begin
                 if lBoHGSaisieActive and ( HGSaisie.Col = cColRefInterne ) then
                  ShiftVK_F7Grille(HGSaisie)
                   else
                    if lBoHGImputationActive and ( StatutGrilleImput <> taConsult ) and ( HGImputation.Col = cColReferenceImput ) then
                     ShiftVK_F7Grille(HGImputation)
                end; // VK_F7
  VK_F8     : if lBoHGSaisieActive then
               begin
                 Key := 0;
                 if ( StatutGrilleReleve = taConsult ) then // sauvegarde de la ligne courante
                  FTOBSauveLigne.Dupliquer(FZEtebac.TOBLigneReleve,true,true);
                 if not EnregistreLigne(HGSaisie.Row) then exit;
                 AfficheGuide(true);
               end;
  VK_F10    : if lBoHGSaisieActive and not FBoModeSelection then
               begin
                Key := 0;
                if not EnregistreLigne(HGSaisie.Row) then
                 exit;
               end // if
                else
                 if lBoHGSaisieActive and FBoModeSelection and BIntegrer.Enabled  then
                  begin
                   BIntegrerClick(nil);
                   Key := 0;
                  end
                   else
                    if lBoHGImputationActive then
                     begin
                      Key := 0;
                      VK_F10Imputation;
                     end; // if
  VK_F11    : begin
               Key:=0;
               TPopUpMenu(GetControl('POPS')).Popup(Mouse.CursorPos.x,Mouse.CursorPos.y);
              end; // VK_F11
  VK_F12    : if lBoHGSaisieActive and PgcCritere.CanFocus then
               begin
                Key:=0;
                PgcCritere.SetFocus;
               end;
  VK_END : if (Shift=[ssCtrl]) and (StatutGrilleReleve=taConsult) then
           begin // on ne place a la fin quand consultation
            Key  := 0 ;
            CreateRow ;
           end
            else
             if (Shift=[]) then
              begin
               Key          := 0 ;
               HGSaisie.col := cColTva  ;
              end ;
{^Z}    90 :  if Shift=[ssAlt] then
              BEGIN Key:=0 ; if Ecran.WindowState=wsMaximized then Ecran.WindowState:=wsNormal else Ecran.WindowState:=wsMaximized END ; //LG
  end; // case
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... :   /  /
Description .. : Affecte la cellule courante avec la valeur de la cellule situé
Suite ........ : au dessus
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.VK_F7Grille( G : THGrid );
begin
 if ( G.Row > 1 ) then
  begin
   G.Cells[G.Col,G.Row] := G.Cells[G.Col,G.Row-1] ;
   StatutGrilleReleve := taModif;
  end; // if

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... : 20/07/2001
Description .. : Incrémente le numéro de pièce automatiquement
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.ShiftVK_F7Grille( G : THGrid ; vIncrement : integer = 1 ) ;
var
 lStRef         : string;
 lStNum         : string;
 lInIndex       : integer;
 l              : integer;
begin

 lStRef := G.Cells[G.Col,G.Row];
 if lStRef = '' then  // si la cellule est vide
  G.Cells[G.Col,G.Row] := FloatToStr(FRdNumeroPieceMax); // on affecte avec la plus grand numero de piece

 StatutGrilleReleve := taModif;

 lStNum         := '';
 lInIndex       := Length(lStRef);
 l              := 0;

 while ( ( lInIndex >= 1 ) and ( lStRef[lInIndex] in ['0'..'9'] ) ) do
  begin
   lStNum := lStRef[lInIndex] + lStNum;
   Dec(lInIndex);
   inc(l);
  end; // while

 if lStNum = '' then Exit ;

 lStNum               := IntToStr( ValeurI(lStNum) + vIncrement) ;
 lStRef               := Copy( lStRef, 1, Length(lStRef) - l) + lStNum ;
 G.Cells[G.Col,G.Row] := lStRef ;

 if isNumeric(lStNum) and ( StrToInt(lStNum) > FRdNumeroPieceMax ) then
  FRdNumeroPieceMax  := StrToInt(lStNum);

end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... : 12/09/2007
Description .. : On recherche l'ensemble des comptes de contrepartie des
Suite ........ : journaux de type banques et on calcule le solde de la
Suite ........ : periode precedente du compte
Suite ........ : - LG - 12/09/2007 - FB 19373 - on cahce ls compte ferme
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.RemplirComboBoxContrepartie;
var
 Q      : TQuery;
begin

 Q := nil;

 // gestion du compte bancaire
 try
   // on recherche l'ensemble des comptes de contrepartie des journal de type banques
   Q := OpenSQL(' select distinct J_CONTREPARTIE '  +
                'from JOURNAL, BANQUECP , GENERAUX '+
                'where J_NATUREJAL ="BQE" '         +
                'and G_FERME="-" '                  +
                'and J_FERME="-" '                  +
                'and J_CONTREPARTIE=G_GENERAL '     +
                'and BQ_GENERAL=J_CONTREPARTIE '    +
                'AND BQ_NODOSSIER = "'              + V_PGI.NoDossier+'" ' + // 19/10/2006 YMO Multisociétés
                'and BQ_DEVISE="'                   + V_PGI.DevisePivot + '" ' +
                ' and J_CONTREPARTIE is not null' ,
                true );

  if Q.EOF then
   begin
    PGIInfo('Aucun des comptes de banque ne possède de RIB !' + #10#13 + 'Vous devez les renseigner avant de saisir des écritures',Ecran.Caption);
   end; // if

  while not Q.Eof do
   begin
    CmbContrepartie.Items.Add(Q.FindField('J_CONTREPARTIE').Value);
    Q.Next;
   end; // while


 if  Q.RecordCount =  1 then
  begin
   CmbContrepartie.ItemIndex := 0; // on se place sur le premier element
  end;   

 finally
  if assigned(Q) then Ferme(Q);
 end; // try

 CmbContrepartieChange(Self);    // on remplit la liste des journaux de banque

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... : 02/09/2004
Description .. : Gestion de la combobox du choix du journal de banque
Suite ........ : associé au compte de contrepartie
Suite ........ : - LG - 02/09/2004 - FB 14491 - correction de l'opensql,
Suite ........ : ajout de guillement pour le champs j_contrepartie
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.CmbContrepartieChange(Sender: TObject);
var
 Q    : TQuery;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 CmbJournal.Clear;
 CmbJournal.Values.Clear;

 // les grilles repassent en modif
 FStatutGrilleImput     := taConsult;
 FStatutGrilleReleve    := taConsult;
 FBoGrilleVide          := true;

 FBoChoixEffectue       := CmbContrepartie.Text <> '';

 VideGrille(HGSaisie);
 VideGrille(HGImputation);
 FZEtebac.Initialize;

 if not FBoChoixEffectue then
  begin
   EnabledControl;
   // affichage d'un message pour prevenir l'utilisateur
   Ecran.Caption        := TraduireMemoire(cStTexteTitreFenetre + ' : ' +  cStTexteSaisirCompte );
   UpdateCaption(Ecran);
   FlashCompte.Flashing := true and not V_PGI.ModeTSE;
   FlashCompte.Color    := clRed;
   exit;
  end; // if

 FlashCompte.Flashing   := false;
 FlashCompte.Color      := clWindowText;

 FZEtebac.StCompteContrepartie   := CmbContrepartie.Text;

 HlbDevise.Caption := TraduireMemoire('Dossier tenu en ') + VH^.LibDevisePivot + '   ' +
                      TraduireMemoire('Compte tenu en ')  + FZEtebac.Devise.Libelle ;

 ImgEuro.Visible   := VH^.TenueEuro;
 // selectionne de l'ensemble des comptes de banques du comptes de contrepartie choisi

 Q := nil;

 try

  Q := OpenSQL('SELECT j_journal,j_libelle,j_contrepartie from journal ' +
                ' where J_NATUREJAL ="BQE" ' +
                ' and J_FERME="-" ' +
                ' and J_CONTREPARTIE = "' + CmbContrepartie.Text +'" ',
                true );

  while not Q.Eof do
   begin
    CmbJournal.Items.Add(Q.FindField('J_LIBELLE').Value);
    CmbJournal.Values.Add(Q.FindField('J_JOURNAL').Value);
    Q.Next;
   end; // while

 if Q.RecordCount = 1 then
  CmbJournal.ItemIndex := 0;

 CmbJournalChange(nil); // on force l'affectation du journal de banque

 finally
  if assigned(Q) then Ferme(Q);
 end; // try

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/07/2001
Modifié le ... : 02/08/2001
Description .. : Quand on change le journal de banque -> on vide la grille et
Suite ........ : on calcul la solde du compte de crontrepartie
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.CmbJournalChange(Sender: TObject);
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 FBoChoixEffectue := CmbJournal.Value <> '' ;

 EnabledControl;

 if not FBoChoixEffectue then
  begin
   // affichage d'un message pour prevenir l'utilisateur
   Ecran.Caption                       := TraduireMemoire(cStTexteTitreFenetre + ' : ' +  cStTexteSaisirJournal );
   UpdateCaption(Ecran);
   FlashJournal.Flashing               := true and not V_PGI.ModeTSE;;
   FlashJournal.Color                  := clRed;
   exit;
  end; // if

  FlashJournal.Flashing                := false;
  FlashJournal.Color                   := clWindowText;

  // les grilles repassent en modif
 FStatutGrilleImput                    := taConsult;
 FStatutGrilleReleve                   := taConsult;
 FBoGrilleVide                         := true;

 FZEtebac.Initialize;
 FZEtebac.StJournalContrepartie        := CmbJournal.Value;
 // chargement des info sur le journal
 if FInfoEcr.Journal.Load([CmbJournal.Value]) = - 1 then exit ;
 // Initialisation de l'objet guide
 FZEtebacGuide.StJournalContrepartie     := FZEtebac.StJournalContrepartie;
 FZEtebacGuide.ZCompte                   := FInfoEcr.Compte;
 FZEtebacGuide.StLeNomTable              := 'CRELBQE';
 FZEtebacGuide.StEtablissement           := CmbEtablissement.Value;
 FZEtebacGuide.StDevise                  := FZEtebac.Devise.Code;
 FZEtebacGuide.StCompteContrepartie      := FZEtebac.StCompteContrepartie;
  // chargement de l'ensemble des guides
 FZEtebacGuide.Load;

  // On recherche si l'on a des RIB a integrer
  RechercheRibAIntegrer;

 // On lance tous de suite la recherche
 BChercheClick(nil);
 CalculSoldeTheorique;
 HGSaisieAfficheSolde;

end;

procedure TOF_SAISIEETEBAC.CmbEtablissementChange(Sender: TObject);
begin
 FZEtebacGuide.StEtablissement := CmbEtablissement.Value ;
  // chargement de l'ensemble des guides
 FZEtebacGuide.Load ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Affichage des soldes des ecritures + lignes de releves
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieAfficheSolde;
begin
 HlbDebitSaisie.caption  := STRFMONTANT ( FZEtebac.RdSoldeDebit , 15 , V_PGI.OkDecV, '' , true);
 HlbCreditSaisie.caption := STRFMONTANT ( FZEtebac.RdSoldeCredit , 15 , V_PGI.OkDecV, '' , true);

  if FRdSolDeT > 0 then
   HlbSoldecompte.Caption := STRFMONTANT ( FRdSolDeT , 15 , V_PGI.OkDecV, '' , true) + ' D'
    else
     if FRdSolDeT < 0 then
      HlbSoldecompte.Caption :=  STRFMONTANT ( FRdSolDeT*(-1) , 15 , V_PGI.OkDecV, '' , true) + ' C'
       else
        HlbSoldecompte.Caption :=  STRFMONTANT ( FRdSolDeT*(-1) , 15 , V_PGI.OkDecV, '' , true);

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... :   /  /
Description .. : On change de ligne -> on regarde si l'on doit enregistrer en
Suite ........ : base
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieRowExit(Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin

  if csDestroying in Ecran.ComponentState then Exit ;

{$IFDEF TT}
 AddEvenement('HGSaisieRowExit Ou:='+intToStr(Ou));
{$ENDIF}

  Cancel := not EnregistreLigne(Ou);

  if Cancel and ( Ou = ( HGSaisie.RowCount - 1 ) ) then
   begin
    HGSaisieDeleteRow(Ou);
    Cancel := false;
   end; // if

  VideGrille(HGImputation);
  StatutGrilleImput := taConsult;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... : 20/07/2001
Description .. : Validation de la ligne
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.HGSaisieIsRowValid( Row: Integer ) : Boolean;
begin

 result := HGSaisieAssignTOB(Row);
 if result then
  result := FZEtebac.IsValideLigneReleve;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... : 18/05/2004
Description .. : Récupération des données de la grille des releves dans la
Suite ........ : TOBReleve
Suite ........ : - FB 13044 - tous les mvt sont genere en OD
Mots clefs ... : 
*****************************************************************}
function TOF_SAISIEETEBAC.HGSaisieAssignTOB( ARow : integer) : boolean;
begin

 result := false;

 if ( HGSaisie.Objects[0,ARow] = nil ) or ( not ( HGSaisie.Objects[0,ARow] is TOB ) ) then exit;

{$IFDEF TT}
 AddEvenement('HGSaisieAssignTOB ARow:='+intToStr(ARow));
{$ENDIF}


 try

  // récupération de l'objet dans la TOB
  FZEtebac.TOBLigneReleve       := TOB(HGSaisie.Objects[0,ARow]);
  FZEtebac.DtDATECOMPTABLE_R    := HGSaisie.Cells[cColDate,ARow];
  FZEtebac.StREFINTERNE_R       := HGSaisie.Cells[cColRefInterne,ARow];
  FZEtebac.STREFEXTERNE_R       := HGSaisie.Cells[cColPiece,ARow];
  FZEtebac.StGENERAL_R          := HGSaisie.Cells[cColImputation,ARow];
  FZEtebac.StETABLISSEMENT_R    := CmbEtablissement.Value;


  if isNumeric(HGSaisie.Cells[cColTva,ARow]) then
   begin
    FZEtebac.RdTAUXTVA_R        := HGSaisie.Cells[cColTva,ARow];
    FZEtebac.StTVASAISIE_R      := 'X';
   end
    else
     begin
      FZEtebac.RdTAUXTVA_R      := 0;
      FZEtebac.TOBLigneReleve.PutValue('CRL_TAUXTVA',0) ;
      FZEtebac.StTVASAISIE_R    := '-';
     end;

  CSetMontants( FZEtebac.TOBLigneReleve,
                Valeur(HGSaisie.Cells[cColDebit,ARow]) ,
                Valeur(HGSaisie.Cells[cColCredit,ARow]) ,
                FZEtebac.Devise,
             //   false,
                true );

 if FZEtebac.RdDEBIT_R = 0  then
{  FZEtebac.StNATUREPIECE_R := 'RC'
   else
    FZEtebac.StNATUREPIECE_R := 'RF'; }

  FZEtebac.StNATUREPIECE_R   := 'OD' ;
  FZEtebac.StLIBELLE_R       := HGSaisie.Cells[cColLibelle,ARow];

  result                          := true;

  except
  on E : exception do
   begin
    V_PGI.IoError := oeSaisie;
    PGIBox( cStTexteErreurRecupImput + #13#10 +
             E.Message ,
             cStTexteTitreFenetre);
   end; // on
 end; // try

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... : 28/10/2002
Description .. : Enregistrer une nouvelle ligne
Suite ........ : -28/10/2002 - correction deu calcul du solde
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.EnregistreLigne( ARow : integer ) : boolean;
var
 lRdSolde : double;
 lBoCancel : boolean ;
 ACol : integer ;
begin

 result := true;

 if ( StatutGrilleReleve = taConsult ) then exit; // on consultation on sort

{$IFDEF TT}
 AddEvenement('EnregistreLigne ARow:='+intToStr(ARow));
{$ENDIF}


 result := false;
 if FZEtebac.TOBLigneReleve =  nil then exit ;
 // sauvegarde de la ligne courante
 FTOBSauveLigne.Dupliquer(FZEtebac.TOBLigneReleve,true,true);


 if ( FZEtebac.StETAT_R <> cPasImput ) then
  begin
   if PGIAsk(cStTexteModifLigneImputer,cStTexteTitreFenetre) = mrNo then
    begin
     BDefaireClick(nil);
     exit;
    end
     else
      begin
       StatutGrilleReleve := taModif;
       result             :=  HGSaisieAssignTOB(ARow);
       if result then AfficheGuide;
       result             := false; // on repasse le flag a zero pour indiquer que l'on n'a pas enreistrer la ligne
       exit;
      end; // if
  end; // if

 ACol := HGSaisie.Col ;
 HGSaisieCellExit(self,ACol,ARow, lBoCancel) ;
 result := HGSaisieIsRowValid(ARow); // validation de la ligne

 if result then
  begin
   FTOBSauveLigne.Dupliquer(FZEtebac.TOBLigneReleve,true,true);
   result := FZEtebac.SaveLigneReleve; // enregistrement du releve
   if not result then exit;
   HGSaisieRefreshLigne(ARow);    //  rafraisir la grille
   // si on a effacé le numero de piece max on repart de l'ancien
   if FZEtebac.StREFINTERNE_R = '' then
    FRdNumeroPieceMax := FRdNumeroPieceMaxOld
     else
      FRDNumeroPieceMaxOld := FRdNumeroPieceMax;
   StatutGrilleReleve := taConsult;

   if ( StrToDate(FZEtebac.DtDATECOMPTABLE_R) >= StrToDate(EdtMulDateDu.Text) ) and
      ( StrToDate(FZEtebac.DtDATECOMPTABLE_R) <= StrToDate(EdtMulDateAu.Text) ) then
    begin
     lRdSolde  := FRdSolDeT - FZEtebac.RdDEBIT_R + FZEtebac.RdCREDIT_R ;
     FRdSolDeT := lRdSolde;
    end; // if
   HGSaisieAfficheSolde;
   FBoGrilleVide := false;
  end; // if

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... : 20/07/2001
Description .. : Refresh de la grille à partir des enregistrement de la TOB
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieRefreshLigne( ARow : integer );
begin

 try

  HGSaisie.Cells[0,ARow]               := intToStr(HGSaisie.Row);
  HGSaisie.Cells[cColDate,ARow]        := FZEtebac.DtDATECOMPTABLE_R;
  HGSaisie.Cells[cColRefInterne,ARow]  := FZEtebac.StREFINTERNE_R;
  HGSaisie.Cells[cColPiece,ARow]       := FZEtebac.StREFEXTERNE_R;
//  HGSaisie.Cells[cColPiece,ARow]       := FZEtebac.StETABLISSEMENT_R;
  HGSaisie.Cells[cColImputation,ARow]  := FZEtebac.StGENERAL_R;
  HGSaisie.Cells[cColAFB,ARow]         := FZEtebac.StCODEAFB_R;


  if   ( FZEtebac.StTVASAISIE_R = 'X' ) then
  HGSaisie.Cells[cColTva,ARow]         := STRFMONTANT ( FZEtebac.RdTAUXTVA_R , 15 , V_PGI.OkDecV, '' , true)
    else
     HGSaisie.Cells[cColTva,ARow]      := '';


  if FZEtebac.RdDEBIT_R <> 0 then // si debit = 0 la cellule doit afficher ''
   HGSaisie.Cells[cColDebit,ARow]      := STRFMONTANT ( FZEtebac.RdDEBIT_R , 15 , V_PGI.OkDecV, '' , true);
  if FZEtebac.RdCREDIT_R <> 0 then
   HGSaisie.Cells[cColCredit,ARow]     := STRFMONTANT ( FZEtebac.RdCREDIT_R , 15 , V_PGI.OkDecV, '' , true);
  HGSaisie.Cells[cColLibelle,ARow]     := FZEtebac.StLIBELLE_R;
  HGSaisie.CellValues[cColEtat,ARow]   := FZEtebac.StETAT_R ;

  // gestion du plus grand numero de pièce
  if ( trim(FZEtebac.StREFINTERNE_R) <> '' )               and
       isNumeric(FZEtebac.StREFINTERNE_R )           and
     ( FZEtebac.StREFINTERNE_R > FRdNumeroPieceMax ) then
   FRdNumeroPieceMax := (FZEtebac.StREFINTERNE_R);

 except
  on E : Exception do
   begin
    V_PGI.IoError := oeSaisie;
    PGIInfo('Erreur lors de la récupération des données en base.'               + #13#10 +
            'Ligne de relevée numéro ' + varToStr(FZEtebac.InCOMPTEUR_R)   + #13#10 +
             E.Message ,
             cStTexteTitreFenetre);
   end;
 end; // try

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Débranche les évènements en fonction de value
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.SetEvent(value: boolean);
begin
 if value then
  AssignEvent
   else
    begin
     Ecran.OnKeyDown            := nil;
     HGSaisie.OnElipsisClick    := nil;
     HGSaisie.OnCellExit        := nil;
     HGSaisie.OnRowExit         := nil;
     HGSaisie.OnKeyPress        := nil;
     HGSaisie.OnKeyDown         := nil;
     HGSaisie.OnRowEnter        := nil;
     HGImputation.OnCellEnter   := nil;
     HGImputation.OnCellExit    := nil;
     HGImputation.OnKeyPress    := nil;
     HGImputation.OnKeyDown     := nil;
     HGImputation.OnExit        := nil;
     HGImputation.OnRowExit     := nil;
    end; // if

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 28/10/2002
Description .. : Suppression d'une ligne
Suite ........ :
Suite ........ : -28/10/2002 - correction deu calcul du solde
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieDeleteRow(ARow: integer);
var
 lRdSolde : double;
begin

 SetEvent(false);

 lRdSolde           := FRdSolDeT + FZEtebac.RdDEBIT_R - FZEtebac.RdCREDIT_R;
 FRdSolDeT          := lRdSolde;
 StatutGrilleReleve := taModif;
 if HGSaisieAssignTOB(ARow) and FZEtebac.DeleteLigneReleve then
  begin
   HGSaisie.Objects[0,ARow] := nil;
   if HGSaisie.RowCount = 2 then
    begin
     StatutGrilleReleve  := taConsult;
     CReateRow;
    end
     else
      begin
       if HGSaisie.Row <> 1 then // si on n'est pas sur la derniere cellule on remonte d'une ligne
        HGSaisie.Row := HGSaisie.Row - 1;
       HGSaisie.DeleteRow(ARow); // on supprime la ligne
      end; // if
   HGSaisie.Refresh;
   NumeroteLigne(1,HGSaisie); // renumeroter les lignes
   HGSaisieAfficheSolde;         // recalcul du solde
   StatutGrilleReleve := taConsult;
   VideGrille(HGImputation);
   RafraichirLigne;
  end; // if

 SetEvent(true);

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 13/06/2005
Description .. : Gestion du filtre de recherche
Suite ........ : - FB 10174 - LG  - 13/06/2005 - ajout du code interbancaire
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.BChercheClick(Sender: TObject);
var
 lStSQL  : string;
 Q       : TQuery;
begin

 {$IFDEF TT}
  AddEvenement('BChercheClick' );
 {$ENDIF}

 if (not IsDateValide(EdtMulDateDu.Text)) or (not IsDateValide(EdtMulDateAu.Text)) then exit ;

 // on se replace dans le libelle
 if EdtMulLibelle.CanFocus then
  EdtMulLibelle.SetFocus;

 lStSQL := ' select * from CRELBQE ';

 lStSQL := lStSQL + 'where ( CRL_DATECOMPTABLE >= "' + UsDate(EdtMulDateDu) + '" and CRL_DATECOMPTABLE <= "' + UsDate(EdtMulDateAu) + '") ' ;

 if EdtMulLibelle.Text <> '' then
 lStSQL := lStSQL + ' and CRL_LIBELLE like "' + EdtMulLibelle.Text + '%"';

 if EdtMulNumeroPiece.Text <> '' then
  lStSQL := lStSQL + ' and CRL_REFINTERNE like "' + EdtMulNumeroPiece.Text + '%"';

 if EdtMulNumeroPieceCpml.Text <> '' then
  lStSQL := lStSQL + ' and CRL_REFEXTERNE like "' + EdtMulNumeroPieceCpml.Text + '%"';

 if EdtMulGeneral.Text <> '' then
  lStSQL := lStSQL + ' and CRL_GENERAL like "' + EdtMulGeneral.Text + '%"';

 if FZEtebac.StJournalContrepartie <> '' then
  lStSQL := lStSQL + ' and CRL_JOURNAL = "' + FZEtebac.StJournalContrepartie + '"';

 if FZEtebac.StCompteContrepartie <> '' then
  lStSQL := lStSQL + ' and CRL_GENERALBQE = "' + FZEtebac.StCompteContrepartie + '"';

 if Pos('<<', CmbEtablissement.Text) = 0 then
  lStSQL := lStSQL + ' and CRL_ETABLISSEMENT = "' + CmbEtablissement.Value + '" ' ;


 // on recupere les critères pour l'etat de l'imputation
 lStSQL := lStSQL + CmbMulEtat.GetSQLValue;

 if ( CmbMulValeur1.Text <> '' ) and ( CmbMulChoix1.Value <> '' ) and ( EdtMulValeur1.Text <> '' )  then
  lStSQL := lStSQL + 'and  ' + CmbMulValeur1.Value + CmbMulChoix1.Value + '"' + EdtMulValeur1.Text + '"' + CmbMulOp.Value + ' ';

 if ( CmbMulValeur2.Text <> '' ) and ( CmbMulOp.Text <> '' ) and ( CmbMulChoix2.Value <> '' )  then
  lStSQL := lStSQL + CmbMulValeur2.Value + CmbMulChoix2.Value + '"' + EdtMulValeur2.Text + '"';

 lStSQL := lStSQL + ' and CRL_TYPE=0'; // on ne recherche que les lignes de relevée

 lStSQL := lStSQL + ' order by CRL_DATECOMPTABLE ';

 Q := nil;

 try

   Q := openSQL( lStSQL, true );
   FZEtebac.LoadReleve(Q);      // chargement des relevées

 finally
   if assigned(Q) then Ferme(Q);
 end; // try

 StatutGrilleReleve := taConsult; // on replace la grille en consultation
 HGSaisieRemplirGrille;     // remplissage de la grille
 CalculSoldeTheorique;
 HGSaisieAfficheSolde;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/06/2001
Modifié le ... : 20/07/2001
Description .. : Remplissage de la grille de saisie après une nouvelle
Suite ........ : recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieRemplirGrille;
var
 i           : integer;
begin

 // on debranche les evenements temporairements
 HGSaisie.OnKeyPress            := nil;
 HGSaisie.OnKeyDown             := nil;
 Ecran.OnKeyDown                := nil;

 // remise à zéro de la grille
 VideGrille(HGSaisie);

 // déactiver les évènements
 HGSaisie.SynEnabled        := false;
 HGSaisie.OnSetEditText     := nil;
 // message envoyer à  la grille bloquant le rafraissisment
 HGSaisie.BeginUpdate;

 if FZEtebac.TOBReleve.Detail.Count = 0 then
   CreateRow  //HGSaisie.RowCount := 2
   else
    begin

     HGSaisie.RowCount := FZEtebac.TOBReleve.Detail.Count + 1;
     HGSaisie.Row      := 1;
     HGSaisie.Col      := GetCol ;

     for i := 0 to FZEtebac.TOBReleve.Detail.Count - 1 do
      begin
       HGSaisie.Row                       := i + 1;
       HGSaisie.Cells[0,i+1]              := intToStr(i+1);
       // faire une fct NextTOB(index);
       FZEtebac.TOBLigneReleve       := FZEtebac.TOBReleve.Detail[i];

       HGSaisieRefreshLigne(HGSaisie.Row);
       HGSaisie.Objects[0,HGSaisie.Row]   := FZEtebac.TOBLigneReleve; // creation de l'enregistrement et ajout à la grille
      end; // for

    FBoGrilleVide := false;

    end; // if Grille est vide


 // se replacer au début
 if HGSaisie.CanFocus then // on donne le focus à la grille
  HGSaisie.setfocus;
 HGSaisie.Row          := 1;
 HGSaisie.Col          := GetCol ;

 // la grille repasse en consultation
 StatutGrilleReleve    := taConsult;
 // rafraichir la grille
 RafraichirLigne;
 HGImputation.Enabled  := false;

 // on rebranche les evenements
 HGSaisie.OnKeyPress   := HGSaisieKeyPress;
 HGSaisie.OnKeyDown    := HGSaisieKeyDown;
 Ecran.OnKeyDown       := FormKeyDown;

 // message demandant à la grille de se mettre a jour
 HGSaisie.EndUpdate;
 HGSaisie.SynEnabled   := true;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Affectactino du statut dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.SetStatutGrilleReleve ( Value : TActionFiche );
begin

 if not ( ( Value = taModif ) and ( StatutGrilleReleve = taCreat ) ) then // le F5 place la grille en modification
  FStatutGrilleReleve := Value;
 EnabledControl;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/06/2001
Modifié le ... : 20/07/2001
Description .. : Dès on presse une touche alphanumerique dans la grille on
Suite ........ : passe en modification
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieKeyPress(Sender: TObject; var Key: Char);
begin
 if Key <> #9 then
  StatutGrilleReleve := taModif;

 if HGSaisie.Col = cColDate then
  ParamDate(ecran,sender,key);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/06/2001
Modifié le ... : 20/07/2001
Description .. : Met la grille en modification quand on appuie sur Suppr
Suite ........ : Rem : la touche Suppr ne declanche pas KeyPress
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if ( Shift = [] ) and ( key = VK_DELETE ) then
  StatutGrilleReleve := taModif;
end;


//****************************************************************************//
//    GESTION DE LA GRILLE DES IMPUTATIONS
//
//****************************************************************************//


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 27/09/2002
Description .. : Grise les cellules débit/crédit en fonction de leur état
Suite ........ :
Suite ........ : - 28/06/2002 - correction du grisage de la case auxiliaire
Suite ........ :
Suite ........ : - 27/09/2002 - correction de la fct, ne pas recherche
Suite ........ : l'auxiliaire sur la premiere colonne
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
 lBoGrise  : boolean;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 // gestion des cases debit et crédit - une seul des deux doit être renseignée
 lBoGrise := ( ( ACol = cColDebitImput )  and ( HGImputation.Cells[cColCreditImput, ARow] <>'' ) and  ( ARow > 0 ) ) or
             ( ( ACol = cColCreditImput ) and ( HGImputation.Cells[cColDebitImput, ARow]  <>'' ) and  ( ARow > 0 ) ) or
             (
              ( ACol = cColAuxiliaireImput )                                                            and
              ( ARow > 0 ) and
              ( not FZEtebac.IsCollectif( HGImputation.Cells[cColGeneralImput,ARow]) ) ) ;

 if lBoGrise then
  begin
   HGImputation.PostDrawCell  := nil; // on debranche l'évènement lors du dessin de la grille
   SetGridGrise(ACol, ARow, HGImputation);
   HGImputation.PostDrawCell  := HGImputationPostDrawCell;
  end;

end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 28/06/2005
Description .. : 2 facon de se déplacer dans la grille :
Suite ........ : - Guide de saisie
Suite ........ : - Standart
Suite ........ : - LG - FB 10173 - interdire la modif des monant suivant un 
Suite ........ : parametre
Suite ........ : -FB 15337 - LG - Il existe un paramétrage au niveau du 
Suite ........ : journal permettant d'interdire la modification de certaines 
Suite ........ : informations issues de la récupération d'un fichier ETEBAC.
Suite ........ : Ces paramétres ne doivent avoir une incidence qu'au 
Suite ........ : niveau de la partie "Mouvement bancaire" c'est-à-dire la 
Suite ........ : partie supérieure de l'écran.
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationCellEnter(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean);
var
 lInCurrentCol          : integer;
 lInCurrentRow          : Integer;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

{$IFDEF TT}
 AddEvenement('HGImputationCellEnter ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow));
{$ENDIF}


 lInCurrentCol       := HGImputation.Col;
 lInCurrentRow       := HGImputation.Row;

 if lInCurrentRow = HGImputation.RowCount - 1 then
  begin // on ne peut pas modifier la derniere ligne de l'imputation ( compte de contrepartie du journal )
   HGImputation.CacheEdit;
  // HGImputation.Options := HGImputation.Options - [GoEditing,GoAlwaysShowEditor];
   HGImputation.Options := HGImputation.Options + [goRowSelect] - [GoEditing,GoAlwaysShowEditor];
   HGImputation.MontreEdit;
  end // if
   else
    begin
     HGImputation.Options := HGImputation.Options - [goRowSelect] + [GoEditing,GoAlwaysShowEditor];
     // on bouge deans la grille pour enlever le ligne selectionné
     if HGImputation.Col = HGImputation.ColCount - 1 then
     begin
      HGImputation.Col := HGImputation.Col - 1 ;
      HGImputation.Col := HGImputation.Col + 1 ;
     end
      else
       begin
        HGImputation.Col := HGImputation.Col + 1 ;
        HGImputation.Col := HGImputation.Col - 1 ;
       end;

    end; // if

    //HGImputation.Options := HGImputation.Options + [GoEditing,GoAlwaysShowEditor];

 // gestion de l'affichage des boutons
 HGImputation.ElipsisButton := ( lInCurrentCol = cColGeneralImput ) or ( lInCurrentCol = cColAuxiliaireImput );

 case lInCurrentCol of
  cColAuxiliaireImput : if not FZEtebac.IsCollectif( HGImputation.Cells[cColGeneralImput,lInCurrentRow]) then
                        begin
                         HGImputation.Cells[cColAuxiliaireImput,lInCurrentRow] := '';
                         PasseColSuivante(ACol,ARow,HGImputation);
                         Cancel := true;
                        end; // if
  cColDebitImput     : if  ( HGImputation.Cells[cColCreditImput, lInCurrentRow] <> '' ) then
                       begin
                        PasseColSuivante(ACol,ARow,HGImputation);
                        Cancel := true;
                       end; // if
  cColCreditImput    : if ( HGImputation.Cells[cColDebitImput, lInCurrentRow] <> '' ) then
                       begin
                        if HGImputation.Row <> HGImputation.Rowcount - 2 then
                         PasseColSuivante(ACol,ARow,HGImputation);
                        Cancel := true;
                       end; // if
  end; // case

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/07/2001
Modifié le ... :   /  /
Description .. : Retourne la prochaine cellule d'arrêt du guide.
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.NextColStop( ARow, ACol : integer ) : TPoint;
var
 i,j  : integer;
 lCol : integer;
begin

 lCol := ACol + 1;

 for i := ARow to HGImputation.RowCount - 1 do
  begin
   for j := lCol to HGImputation.ColCount - 1 do
     if FZEtebacGuide.IsColStop(i,j) then
      begin
       Result.x := j;
       Result.y := i;
       exit;
      end; // if
    lCol := 1;
  end; // for

 // il n'y a plus de point d'arret, on se replace au debut
 Result.x  := 1;
 Result.y  := 1;

end;


procedure TOF_SAISIEETEBAC.HGImputationRowEnter( Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin
{$IFDEF TT}
 AddEvenement('HGImputationRowEnter Ou:='+intToStr(Ou));
{$ENDIF}

 cancel := not HGImputationAssignLigneTOB(HGImputation.Row);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... : 27/07/2001
Description .. : Gestion des deplacement dans la grille des imputations :
Suite ........ : mode guide + mode normal
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationCellExit(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean);
var
 lCell          : TPoint;
 lBoFirstCell   : boolean;
 RdMontant      : double;
begin
 // on stocke la cellule que l'on est en train de tester, et l'on sort de ce traitement si l'on n'a pas bougé
 // rem : cet evenement est appelé plusieur fois de suite quand on force le changement de cellule
 if ( csDestroying in Ecran.ComponentState ) or ( ( FCurrentCol = ACol ) and ( FCurrentRow = ARow ) ) then Exit ;
 if ARow > FZEtebac.TOBImput.Detail.Count then exit ;

 // recuperation des données de la grille vers la TOB
 if not HGImputationAssignLigneTOB(ARow) then exit;

{$IFDEF TT}
 AddEvenement('HGImputationCellExit ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow));
{$ENDIF}

 FCurrentCol            := ACol;
 FCurrentRow            := ARow;

 case ACol of
  cColGeneralImput :
                       begin
                        FZEtebac.CellExitGenImput(HGImputation,ACol,ARow,Cancel) ;
                        if Cancel then
                         begin
                          FCurrentCol             := 0;
                          FCurrentRow             := 0;
                          exit ;
                         end; // if
                       end;
  cColAuxiliaireImput :
                        begin
                        FZEtebac.CellExitAuxImput(HGImputation,ACol,ARow,Cancel) ;
                        if Cancel then
                         begin
                          FCurrentCol             := 0;
                          FCurrentRow             := 0;
                          exit ;
                         end; // if
                       end;
  cColDebitImput,
  cColCreditImput    : if IsNumeric(HGImputation.Cells[ACol,ARow]) then
                        begin
                         // on reinitialise ces valeurs à zero pour redeclencher l'evenement
                         FCurrentCol             := 0;
                         FCurrentRow             := 0;
                         RdMontant                     := Valeur(HGImputation.Cells[ACol,ARow]);
                         Cancel                        := not( VH^.MontantNegatif ) and ( RdMontant < 0 ); // on ne peut pas saisir de valeur negative
                         if Cancel then
                          begin
                           PGIInfo(cStTexteMontantNegatif,cStTexteTitreFenetre);
                           exit;
                          end; // if
                         HGImputation.Cells[ACol,ARow] := STRFMONTANT ( RdMontant , 15 , V_PGI.OkDecV, '' , true);
                         // on reaffecte la grille pour prendre en compte la nouvelleur valeur des montants
                         // ! a enlever pour la saisie en guide
                         if not HGImputationAssignLigneTOB(ARow) then
                          exit;
                        end; // if
 end; // case

 if FBoGuideRun then
  begin  // on en saisie guide

   HGBeginUpdate(HGImputation) ;
   HGImputationAssignTOB ;
   lCell            := NextColStop( ARow, ACol);
   FCurrentCol      := ACol;
   FCurrentRow      := ARow;
   ACol             := lCell.x;
   ARow             := lCell.y;
   Cancel           := true;

   lBoFirstCell     := ( ARow = 1 ) and ( ACol = 1 );

   if lBoFirstCell  then
    begin
     {$IFDEF TT}
     AddEvenement('Fin du mode guide');
     {$ENDIF}

     FZEtebacGuide.RecalculGuide ( FZEtebac.TOBImput , FCurrentRow  );
     FCurrentCol             := 0;
     FCurrentRow             := 0;
     SetModeGuide(false);
     // on supprime les lignes vides
     FZEtebac.SupprimerLigneImputVide;
     // on remet a jour la grille si jamais des lignes en été supprimer
     HGImputationRemplirGrille;
     HGEndUpdate(HGImputation) ;

    end
     else
      begin
        // HGImputationAssignTOB ;
        //HGImputationAssignLigneTOB(FCurrentRow);
        FZEtebacGuide.RecalculGuide ( FZEtebac.TOBImput , FCurrentRow  );
        HGImputationRemplirGrille;
      end; // if

  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Gestion des Flash labels
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.SetModeGuide( Value : boolean);
begin
 // on désactive se mode quand on se retrouve dans la première cellule de la grille ^
 FBoGuideRun                     :=  Value;
 // on ne fait plus clignoter le label
 FlashGuide.Flashing             := FBoGuideRun and not V_PGI.ModeTSE;
 FlashGuide.Visible              := FBoGuideRun;
 FlashCurrentReleve.Flashing     := FBoGuideRun and not V_PGI.ModeTSE;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Ajout d'une nouvelle ligne dans la grille et la TOB
Suite ........ : TOB_RELBQE
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.NouvelleImputation;
var
 lCol,lRow : integer;
 b   : boolean;
begin

 // if HGImputation.Row <> ( HGImputation.RowCount - 1 ) then exit; // on ajout une ligne uniquement si on etait en consultation

 HGImputation.SynEnabled := false;
 SetEvent(false); // déactive tous les evenements pour crée une nouvelle enregistrement

 HGImputation.InsertRow(HGImputation.Row);

 //if HGImputation.CanFocus then // on donne le focus à la grille
//  HGImputation.setfocus;
 HGImputation.Row    := HGImputation.Row - 1;  // on se place sur cette nouvelle ligne
 HGImputation.col    := cColGeneralImput;
 HGImputationAssignLigneTOB(HGImputation.Row);
 StatutGrilleImput   := taCreat;

 AfficheSoldeImputation;

 NumeroteLigne(HGImputation.Row,HGImputation);

 SetEvent(true); // on réactive les évènements
 lCol := HGImputation.Col;
 lRow := HGImputation.Row;

 // on lance manuellement cet evenement pour supprimer la ligne pleine si on etait sur la derniere ligne
 HGImputationCellEnter(nil,lCol, lRow , b);

 HGImputation.SynEnabled := true;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/07/2001
Modifié le ... : 14/10/2002
Description .. : Recherche d'un guide de saisie en fonction des info
Suite ........ : présente dans la ligne
Suite ........ : 
Suite ........ : - 14/10/2002 - correction d'une fuite memoire
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.AfficheGuide( vBoLastGuide : boolean = false ) : boolean;
var
 lCellStop  : TPoint;
 lTOB       : TOB;
begin
 // suppression des lignes d'imputation courantes
 FZEtebac.DeleteCurrentImputation;
 if FZEtebac.StGENERAL_R<> '' then
  begin
   FZEtebac.AssignInfoCompte;
   FZEtebacGuide.StCompteTVA := FZEtebac.StCompteTVA;
  end ;
// HGSaisieAssignTOB(HGSaisie.Row) ;

 lTOB:=TOB.Create('',nil,-1);

 try

 if vBoLastGuide then
   result:=FZEtebacGuide.RechercheDernierGuide( FZEtebac.TOBLigneReleve ,
                                                lTOB ,
                                                HGSaisie )
   else
    result:=FZEtebacGuide.RechercheGuideEnBase(  FZEtebac.TOBLigneReleve ,
                                                 lTOB ,
                                                 HGSaisie ) ; // on le repasse la grille pour l'affichage de lookup de choix des guides
                                                          // en affectation automatique on passe la valeur nil, par d'interaction avec l'utilisateur
 if result then
  begin // un guide existe
   // on affecte les enregistrements contenu dans lTOB à la ligne de releve courante
   FZEtebac.CreerImputation(lTOB);
   HGImputationRemplirGrille;
   lCellStop              := NextColStop(1,0);
   HGImputation.Col       := lCellStop.x;
   HGImputation.Row       := lCellStop.y;

   // on affecte le flag indiquant ce mode de saisie ( utilisée dans le CellExit et le CellEnter )
   // si le guide n'a pas de point d'arret en ne passe pas en mode guide
   SetModeGuide(FZEtebacGuide.PossedeArret);
   // on passe en modification dans la grille
   StatutGrilleImput := taCreat;
   // on complète le ligne de relevé avec les info du guide
   SynchroniseHGSaisie;
   result := true;

  end; //if

 AfficheSoldeImputation;
 CmbContrepartie.Enabled        := false;
 BIntegrer.Enabled              := false;

 finally
  if assigned(lTOB) then lTOB.Free;
 end ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Remplit la grille des imputations avec les valeurs de la base
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.AltDImputation : boolean;
begin

 result := false;
 
 if FZEtebac.StETAT_R = cPasImput  then exit;

 // sauvegarde de la ligne courante
 FTOBSauveLigne.Dupliquer(FZEtebac.TOBLigneReleve,true,true);

 if FZEtebac.RechercheImputation then
  result := HGImputationRemplirGrille
   else
    begin
     StatutGrilleReleve := taModif;
     EnregistreLigne(HGSaisie.Row);
    end; // if

 AfficheSoldeImputation;

 // rafraichir les grilles pour relance GetCellCanvas
 HGSaisie.Refresh;
 HGImputation.Refresh;

end;

function TOF_SAISIEETEBAC.HGImputationRemplirGrille : boolean;
var
 lInRowCount : integer;
begin

  // remise à zéro de la grille
   VideGrille(HGImputation);
   HGImputation.Enabled           := true;
   lInRowCount                    := 2;

   // déactiver les évènements
   HGImputation.SynEnabled        := false;
   HGImputation.OnSetEditText     := nil;
   // ne pas rafraichir la grille pendant la mise à jour
   HGImputation.BeginUpdate;

   HGImputation.Row               := 1;
   HGImputation.Col               := cColGeneralImput;

   CmbContrepartie.Enabled        := false;
   BIntegrer.Enabled              := false;

   // on se place sur la premiere imputation
   FZEtebac.FirstImputation;

   // affichage du nom du guide
   HlbNomGuide.Visible            := FZEtebac.StGUIDE_I <> '';
   HlbNomGuide.Caption            := TraduireMemoire('Guide n°') + varToStr(FZEtebac.StGUIDE_I) + ' ' + RechDom('TTGUIDEECR', FZEtebac.StGUIDE_I ,false);

   while assigned( FZEtebac.TOBLigneImputation ) do
    begin

       HGImputation.RowCount := lInRowCount;
       HGImputation.Row      := lInRowCount - 1; // on se place sur la ligne courante

       HGImputation.Cells[0,HGImputation.Row]                   := intToStr(HGImputation.Row);
       HGImputation.Cells[cColGeneralImput,HGImputation.Row]    := FZEtebac.StGENERAL_I;
       HGImputation.Cells[cColReferenceImput,HGImputation.Row]  := FZEtebac.StREFINTERNE_I;
       HGImputation.Cells[cColLibelleImput,HGImputation.Row]    := FZEtebac.StLIBELLE_I;
       HGImputation.Cells[cColAuxiliaireImput,HGImputation.Row] := FZEtebac.StAUXILIAIRE_I;
       if FZEtebac.RdDEBIT_I <> 0 then // si debit = 0 la cellule doit afficher ''
        HGImputation.Cells[cColDebitImput,HGImputation.Row]     := STRFMONTANT ( FZEtebac.RdDEBIT_I , 15 , V_PGI.OkDecV, '' , true);
       if FZEtebac.RdCREDIT_I <> 0 then
        HGImputation.Cells[cColCreditImput,HGImputation.Row]    := STRFMONTANT ( FZEtebac.RdCREDIT_I , 15 , V_PGI.OkDecV, '' , true);

       HGImputation.Objects[cColObjectReleve,HGImputation.Row]  := FZEtebac.TOBLigneImputation; // creation de l'enregistrement et ajout à la grille

     Inc(lInRowCount);
     FZEtebac.NextImputation;

    end; // while

     // se replacer au début
    if HGImputation.CanFocus then // on donne le focus à la grille
      HGImputation.setfocus;
    HGImputation.Row      := 1;
    HGImputation.Col      := cColGeneralImput;

    // rafraichir la grille
    HGImputation.EndUpdate;
    HGImputation.SynEnabled := true;

   result := true;

   HGSaisie.Options      := HGSaisie.Options + [goRowSelect];

   HGSaisie.TitleBold := false;

   NumeroteLigne(1,HGImputation);

end;

procedure TOF_SAISIEETEBAC.SynchroniseHGSaisie;
begin
 FZEtebac.StGENERAL_R := FZEtebac.StGENERAL_I;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Affiche le solde des imputations
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.AfficheSoldeImputation;
begin
 FZEtebac.CalculSoldeImput;
 HlbDebitImput.Caption          := STRFMONTANT ( FZEtebac.RdSoldeDebitImput , 15 , V_PGI.OkDecV, '' , true)  ;
 HlbCreditImput.Caption         := STRFMONTANT ( FZEtebac.RdSoldeCreditImput , 15 , V_PGI.OkDecV, '' , true) ;
 if FZEtebac.RdSoldeImput > 0 then
  HlbSoldeImput.Caption         := STRFMONTANT ( FZEtebac.RdSoldeImput , 15 , V_PGI.OkDecV, '' , true)  + ' D'
   else
    if FZEtebac.RdSoldeImput < 0 then
     HlbSoldeImput.Caption      := STRFMONTANT ( FZEtebac.RdSoldeImput * (-1) , 15 , V_PGI.OkDecV, '' , true) + ' C'
      else
       HlbSoldeImput.Caption      := STRFMONTANT ( FZEtebac.RdSoldeImput , 15 , V_PGI.OkDecV, '' , true);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 21/02/2007
Description .. : Gestion des click sur des cellules imputation
Suite ........ : et tva
Suite ........ : - LG - 21/02/2007 - on affectet la tob courante d'imput
Suite ........ : avant d'appeller la recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationElipsisClick(Sender: TObject);
begin
 if ( csDestroying in Ecran.ComponentState )  then Exit ;
 HGImputationAssignLigneTOB(HGImputation.Row) ;
 FZEtebac.ElipsisClick(HGImputation);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... : 30/07/2001
Description .. : Assigne la ligne d'imputation courante avec las valeurs de la
Suite ........ : ligne de la grille courantes
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.HGImputationAssignLigneTOB( ARow : integer) : boolean;
begin

 result := false;

{$IFDEF TT}
 AddEvenement('HGImputationAssignLigneTOB ARow:='+intToStr(ARow));
{$ENDIF}

 try

  if HGImputation.Objects[cColObjectReleve,ARow] = nil then
   begin
    FZEtebac.AddLigneImput(ARow-1);
    HGImputation.Objects[cColObjectReleve,ARow]  := FZEtebac.TOBLigneImputation;
   end
    else
     FZEtebac.TOBLigneImputation := TOB(HGImputation.Objects[cColObjectReleve,ARow]);  // récupération de l'objet dans la TOB

  if not assigned(FZEtebac.TOBLigneImputation) then
   begin
     PGIBox( cStTexteErreurRecupImput, cStTexteTitreFenetre);
    exit;
   end;

  FZEtebac.StGENERAL_I      := HGImputation.Cells[cColGeneralImput,ARow];
  FZEtebac.StAUXILIAIRE_I   := HGImputation.Cells[cColAuxiliaireImput,ARow];

  CSetMontants( FZEtebac.TOBLigneImputation ,
                Valeur(HGImputation.Cells[cColDebitImput,ARow]) ,
                Valeur(HGImputation.Cells[cColCreditImput,ARow]) ,
                FZEtebac.Devise,
               // false,
                true );

  FZEtebac.StLIBELLE_I         := HGImputation.Cells[cColLibelleImput,ARow];
  FZEtebac.StREFINTERNE_I      := HGImputation.Cells[cColReferenceImput,ARow];
  FZEtebac.InNUMLIGNE_I        := ARow;

  result                            := true;

 except
  on E : exception do
   begin
    V_PGI.IoError := oeSaisie;
    PGIBox( cStTexteErreurRecupImput + #13#10 +
            E.Message ,
            cStTexteTitreFenetre);
   end; // on
 end; // try

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/07/2001
Modifié le ... : 20/07/2001
Description .. : Affecte FTOBImput avec l'ensemble des lignes saisies dans
Suite ........ : la grille
Mots clefs ... :
*****************************************************************}
function TOF_SAISIEETEBAC.HGImputationAssignTOB : boolean;
var
 i : integer;
begin

 result := false;

 for i := 1 to ( HGImputation.RowCount - 1 ) do
  begin
   result := HGImputationAssignLigneTOB(i);
   if not result then exit;
  end; // for

 // On affiche le soldes des imputations
 AfficheSoldeImputation;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Vide la grille et reinitialise les objets associe
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.VideGrille ( G : THGrid );
var
 i : integer;
begin

 // Avant le VidePile pour bien placer le curseur
 G.Row := G.FixedRows;
 G.Col := G.FixedCols;
 G.VidePile(FALSE);

 // reinitialise les pointeur sur les objets
 for i:= 1 to G.RowCount - 1 do
  begin
   G.Objects[cColObjectGuide,i]  := nil;
   G.Objects[cColObjectReleve,i] := nil;
  end; // for

// G.RowCount := 2;
 G.Refresh;

end;

procedure TOF_SAISIEETEBAC.HGImputationKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

begin
 if ( Shift = [] ) and ( key = VK_DELETE ) then
  StatutGrilleImput := taModif;
end;

procedure TOF_SAISIEETEBAC.HGImputationKeyPress(Sender: TObject; var Key: Char);
begin
 if Key <> #27 then
  StatutGrilleImput := taModif;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : On repasse HGImputation.ElipsisButton à false car le
Suite ........ : cellenter n'est pas declencher quand on sort et revient dans
Suite ........ : la grille par code. C'est le cas quand on enchaine les modes
Suite ........ : guides.
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationExit(Sender: TObject);
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit;

{$IFDEF TT}
 AddEvenement('HGImputationExit');
{$ENDIF}


 if FBoGuideRun then
  begin

   if PGIAskCancel(cStTexteArretGuide,cStTexteTitreFenetre) = mrYes then
    begin
     SetModeGuide(false);
     FZEtebac.StatutImput := taCreat;
    // ChcBouclage.Checked := false;
    end; // if PGIAsk
   end // FBoGuideRun
    else
     begin

      if ( FStatutGrilleImput <> taConsult ) then
       begin
        if (PGIAskCancel(cStTexteEchap,cStTexteTitreFenetre) = mrYes) then
         begin
          FZEtebac.StatutImput := taCreat;
          ReaffecterLigneReleve;
          ChcBouclage.Checked := false;
        end
         else
          begin
           if HGImputation.CanFocus then
            HGImputation.SetFocus;
           exit;
          end;
        end;

        HGImputation.ElipsisButton := false;
        StatutGrilleImput := taConsult;
        VideGrille( HGImputation );
        AfficheSoldeImputation;
        RemettreEnEditionHGSaisie;
        // on vide la grille des imputations
        StatutGrilleReleve := taConsult;
        VideGrille( HGImputation );
        HGImputation.TitleBold := false;

        if HGSaisie.CanFocus then
         begin
          HGSaisie.SetFocus;
          HGSaisie.Refresh;
          HGSaisie.MontreEdit;
        end;

    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... "":   /  /
Description .. : Sauvegarde de la TOB quand on rentre dans la ligne
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieRowEnter( Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin
{$IFDEF TT}
 AddEvenement('HGSaisieRowEnter Ou:='+intToStr(Ou));
{$ENDIF}
 RafraichirLigne;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Rafraichissements des infos sur les lignes quand on en
Suite ........ : change
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.RafraichirLigne;
begin
 FlashCurrentReleve.Caption      := TraduireMemoire(cStTexteLigneReleve) + intToStr(HGSaisie.Row);
 if ( HGSaisie.Cells[ 0,HGSaisie.Row] = '' ) then
  HGSaisie.Cells[ 0,HGSaisie.Row] := intToStr(HGSaisie.Row);
 // on affiche le solde de la ligne de releve si on avait chargé les lignes d'imputations en memoire
 if not HGSaisieAssignTOB(HGSaisie.Row) then exit;
  // sauvegarde de la ligne courante
 FTOBSauveLigne.Dupliquer(FZEtebac.TOBLigneReleve,true,true);

 // mise a jour des info de la affiché dans la barre de menu
 Ecran.Caption                   := TraduireMemoire(cStTexteTitreFenetre + ' : compte ' + FZEtebac.StCompteContrepartie ) +
                                 '  ' + FZEtebac.StLibelleEtat;
 UpdateCaption(Ecran);

 {$IFDEF TT}
  Ecran.Caption                  := Ecran.Caption + ' : numero = ' + varToStr(FZEtebac.InCOMPTEUR_R);
 {$ENDIF}
 UpdateCaption(Ecran);
 HGSaisieAfficheSolde;
 VideGrille(HGImputation);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/07/2001
Modifié le ... : 29/06/2004
Description .. : Gestion de l'etat des bouttons
Suite ........ : - FB 13117 - LG - le bouton etablissement est valide
Suite ........ : uniquement que la grille est vide
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.EnabledControl;
begin

 BFermer2.Enabled         := FBoModeSelection and ( StatutGrilleReleve = taConsult );
 BSelectionner.Enabled    := FBoModeSelection and ( StatutGrilleReleve = taConsult );
 BDeSelectionner.Enabled  := FBoModeSelection and ( StatutGrilleReleve = taConsult );
 BIntegrer.Enabled        := FBoModeSelection and ( FListeIntegrer.Count > 0 );
 BValider.Enabled         := not FBoModeSelection;
 BMode.Enabled            := not FBoModeSelection and FBoChoixEffectue and ( StatutGrilleReleve = taConsult );

 PgcCritere.Enabled       :=  FBoChoixEffectue;
 HGSaisie.Enabled         :=  FBoChoixEffectue;
 HGImputation.Enabled     :=  HGImputation.Enabled and FBoChoixEffectue;

 BInsert.Enabled          := ( StatutGrilleReleve = taConsult ) and ( StatutGrilleReleve = taConsult ) and FBoChoixEffectue and not FBoModeSelection;
 BDelete.Enabled          := ( StatutGrilleReleve = taConsult ) and ( StatutGrilleReleve = taConsult ) and FBoChoixEffectue and not FBoModeSelection;
// BValider.Enabled         := ( StatutGrilleReleve <> taConsult ) and ( StatutGrilleReleve <> taConsult ) and FBoChoixEffectue;

 BDefaire.Enabled         := ( StatutGrilleReleve <> taConsult ) and ( StatutGrilleReleve <> taConsult ) and FBoChoixEffectue;
 BCherche.Enabled         := FBoChoixEffectue;

 CmbContrepartie.Enabled  := ( ( StatutGrilleReleve = taConsult ) or
                             (  FBoGrilleVide and ( StatutGrilleReleve <> taConsult ) ) ) and
                             ( CmbContrepartie.Items.Count > 1 ) and ( FStArgument = '' );

 CmbJournal.Enabled       := ( ( StatutGrilleReleve = taConsult ) or
                             (  FBoGrilleVide and ( StatutGrilleReleve <> taConsult ) ) ) and
                             ( CmbJournal.Items.Count > 1 );

 // GCO - 05/01/2005 - FQ 15186
 PositionneEtabUser(CmbEtablissement);
// lBoEtab                  := not CmbEtablissement.Enabled;
// if not lBoEtab then
//  CmbEtablissement.Enabled := FBoGrilleVide ;

 BExport.Enabled          := ExJaiLeDroitConcept(ccExportListe, False) and not FBoGrilleVide ;
 BCreerGuide.Enabled      := HGImputation.Enabled ;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Affectaction du statut dans la grille et mise à jour de l'etat
Suite ........ : des boutons
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.SetStatutGrilleImput ( Value : TActionFiche );
begin

 if not ( ( Value = taModif ) and ( StatutGrilleImput = taCreat ) ) then // le F5 place la grille en modification
  FStatutGrilleImput := Value;
 EnabledControl;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/07/2001
Modifié le ... : 20/07/2001
Description .. : Numérote les lignes de la grille à partir de la ligne ARow
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.NumeroteLigne ( ARow : integer ; G : THGrid );
var
 i : integer;
begin
 if FBoSorted then exit ;
 for i := ARow to G.RowCount - 1 do
  G.Cells[0,i] := intToStr(i);
end;

procedure TOF_SAISIEETEBAC.HGSaisieOnSorted(Sender: TObject);
begin
 FBoSorted := true ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Test la validite de la ligne d'imputation avant de sortir
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationRowExit( Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin

 if csDestroying in Ecran.ComponentState then Exit ;
 if Ou > FZEtebac.TOBImput.Detail.Count then exit ;

 if Ou = ( HGImputation.RowCount - 1 ) then exit ; 

{$IFDEF TT}
 AddEvenement('HGImputationRowExit Ou:='+intToStr(Ou));
{$ENDIF}

 HGImputationAssignLigneTOB(Ou);
 FZEtebac.TypeContexte := TModeSaisie;
 Cancel                := not FZEtebac.IsValideLigneImput;

 if not cancel then
  AfficheSoldeImputation;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Test de validiter  des dates
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.EdtDateValide( Sender : THCritMaskEdit );
begin
 if csDestroying in Ecran.ComponentState then Exit;
 IsDateValide(Sender.Text) ;
end;

function TOF_SAISIEETEBAC.IsDateValide( Sender : string ) : boolean ;
begin
 result := false ;
 if (length(trim(Sender))) < 10 then exit;
 try
 result := CControleDateBor(StrToDate(Sender),FInfoEcr.Exercice,true) ;
 except
  On E : Exception do
   PGIError('Date non valide') ;
 end ;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Contrôle que la date de debut ne peut inférieur à date de
Suite ........ : debut de l'exercice N
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.EdtMulDateDuChange( Sender : TObject );
begin
 EdtDateValide(EdtMulDateDu);
end;

procedure TOF_SAISIEETEBAC.EdtMulDateDuExit( Sender : TObject );
begin
 EdtDateValide(EdtMulDateDu);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... : 27/07/2001
Description .. : Contrôle que la daet saisie est inférieur à ka date de fin de
Suite ........ : l'excercice N+1
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.EdtMulDateAuChange( Sender : TObject );
begin
 EdtDateValide(EdtMulDateAu);
end;

procedure TOF_SAISIEETEBAC.EdtMulDateAuExit( Sender : TObject );
begin
 EdtDateValide(EdtMulDateAu);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : enregistrement de l'imputation, passage à ligne suivante de
Suite ........ : releve et bouclage su cmbBouclage est selectionne
Description .. : - FB 13184 - 06/01/2004 - on remets la grille en edition lors
Suite ........ : de l'enregsitrement
*****************************************************************}
procedure TOF_SAISIEETEBAC.VK_F10Imputation;


 procedure TresoValid ;
 begin
  if FInfoEcr.Journal.GetValue('J_TRESOVALID') =  'X' then
   begin
    FBoParle := false ;
    if FStLastCrit <> ( FZEtebac.StJournalContrepartie + IntToStr( GetPeriode(FZEtebac.DtDATECOMPTABLE_R) ) ) then
     FInNumFolio := - 1 ;
    SelectionnerLigne ;
    BIntegrerClick(nil) ;
    BDeleteClick(nil);
    FBoParle           := true ;
    FStatutGrilleImput := taConsult;
    FStLastCrit        := FZEtebac.StJournalContrepartie + IntToStr( GetPeriode(FZEtebac.DtDATECOMPTABLE_R) ) ;
   end ;
 end ;

 procedure BouclageGuide;
  begin
   // on boucle en mode guide sur la prochaine imputation non imputée
   if ( HGSaisie.Row = ( HGSaisie.RowCount - 1 ) ) then
    begin
     CreateRow;    // on est sur la derniere ligne -> on sort du mode guide et on ajoute une ligne
     RemettreEnEditionHGSaisie;
    end
     else
      begin
       HGSaisie.Row := HGSaisie.Row + 1;
       HGSaisieAssignTOB(HGSaisie.Row);
       if ( FZEtebac.StETAT_R <> cPasImput ) then
        BouclageGuide
         else
          begin
           RafraichirLigne;
           if not ( AfficheGuide ) and ChcBouclage.Checked then
            BouclageGuide;
          end; // else
      end; // if

  end; // procedure

begin
 // en mode guide on ne peut pas enregistrer
 if FBoGuideRun then
  begin
   PGIInfo(cStTextePasF10,cStTexteTitreFenetre);
   if HGImputation.CanFocus then
    HGImputation.SetFocus;
   exit;
  end;
 // FB 13184 on repasse la grille en edition
 HGImputation.Options := HGImputation.Options - [goRowSelect] + [GoEditing,GoAlwaysShowEditor];

 // on recupere les enregistrements de la grille
 if not HGImputationAssignTOB then
   exit;

  // on supprime les lignes vides
  FZEtebac.SupprimerLigneImputVide;

 // on remet a jour la grille si jamais des lignes en été supprimées
 HGImputationRemplirGrille;

 FZEtebac.TypeContexte := TModeSaisie;
 if not FZEtebac.IsValideImputation then
  exit;

  // on complete le ligne de relevé avec les info du guide
 // SynchroniseHGSaisie;

  FZEtebac.AssignInfoTVA;

  // l'imputation est coherente et validé
  FZEtebac.StETAT_R := cImputValide;

  if not FZEtebac.SaveLigneReleve then
   begin
    PGIInfo('Erreur lors de l''enregistrement',cStTexteTitreFenetre);
    exit;
   end;

  // on rafraichis la grille des releves
  HGSaisieRefreshLigne(HGSaisie.Row);
  StatutGrilleReleve := taConsult;

  if ( HGSaisie.Row = ( HGSaisie.RowCount - 1 ) )  then
   begin
    TresoValid ;
    CreateRow;    // on est sur la derniere ligne -> on sort du mode guide et on ajoute une ligne
    RemettreEnEditionHGSaisie;
   end
    else
     if ChcBouclage.Checked then
      begin
       TresoValid ;
       BouclageGuide;
      end
       else
        begin
         FStatutGrilleImput := taConsult;
         HGSaisie.Row       := HGSaisie.Row + 1;
         HGSaisie.Col       := GetCol ;
         HGSaisie.SetFocus;
         RafraichirLigne;
         RemettreEnEditionHGSaisie;
        end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Calcul du solde de la piece est affectation a la ligne
Suite ........ : courante
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.VK_F6Imputation;
var
 lRdSolde                 : double;
 lRdCredit                : double;
 lRdDebit                 : double;
begin

 // on calcule et on affiche le solde des imputations
 HGImputationAssignLigneTOB(HGImputation.Row);
 AfficheSoldeImputation;

 lRdSolde  := FZEtebac.RdSoldeDebitImput - FZEtebac.RdSoldeCreditImput;
 if HGImputation.Cells[cColCreditImput, HGImputation.Row] = '' then
  lRdCredit := 0
   else
   lRdCredit := Valeur(HGImputation.Cells[cColCreditImput, HGImputation.Row]);
 if HGImputation.Cells[cColDebitImput, HGImputation.Row] = '' then
  lRdDebit := 0
   else
    lRdDebit  := Valeur(HGImputation.Cells[cColDebitImput, HGImputation.Row]);

 if (lRdSolde <> 0 ) then
  begin
   if lRdDebit > 0 then
    begin
     lRdDebit := lRdDebit - lRdSolde;
     if ( lRdDebit < 0 ) and not( VH^.MontantNegatif ) then
      begin
        lRdCredit := lRdDebit * (-1);
        lRdDebit := 0;
       end;
    end // lRdDebit > 0
     else
      if lRdCredit > 0 then
       begin
        lRdCredit := lRdSolde + lRdCredit;
        if ( lRdCredit < 0 ) then
         begin
          lRdDebit := lRdCredit * (-1);
          lRdCredit := 0;
         end;
       end // lRdCredit > 0
        else // on est dans une ligne ou les cases debit et credit sont vides
         if lRdSolde > 0 then
          lRdCredit := lRdSolde
           else
            if lRdSolde < 0 then
             lRdDebit := lRdSolde * (-1);


    if lRdCredit = 0 then
     HGImputation.Cells[cColCreditImput, HGImputation.Row] := ''
      else
       HGImputation.Cells[cColCreditImput, HGImputation.Row] := STRFMONTANT ( lRdCredit , 15 , V_PGI.OkDecV, '' , true);

    if lRdDebit = 0 then
     HGImputation.Cells[cColDebitImput, HGImputation.Row] := ''
      else
       HGImputation.Cells[cColDebitImput, HGImputation.Row] := STRFMONTANT ( lRdDebit , 15 , V_PGI.OkDecV, '' , true);

    // on recalcule le nouveau solde
    HGImputationAssignLigneTOB(HGImputation.Row);
    AfficheSoldeImputation;

  end; // if Solde <> 0

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Colorisation des lignes de releve en fonction de leur etat
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGSaisieGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit;

 if FBoModeSelection then exit ;

 if ( Screen.ActiveControl = HGImputation ) and ( ARow <> HGSaisie.Row ) then
  Canvas.Font.Color := clInactiveCaption
   else
    if ( Screen.ActiveControl = HGSaisie ) then
     begin

      if HGSaisie.Cells[cColEtat,ARow] = '3' then
       Canvas.Font.Color := clInactiveCaption
        else
         Canvas.Font.Color := clWindowText;

     end; // if

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Colorisation des lignes d'imputation en fonction de leur etat
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit;

  if ( Screen.ActiveControl = HGSaisie ) and ( ARow <> HGSaisie.Row ) then
    begin
 //   Canvas.Font.Style := Canvas.Font.Style + [fsItalic] ;
    Canvas.Font.Color := clInactiveCaption;
    end
  //end
   else
    if ( Screen.ActiveControl = HGImputation ) then
     begin
   //   Canvas.Font.Style := Canvas.Font.Style - [fsItalic] ;
      Canvas.Font.Color := clWindowText;

 if ARow = ( HGImputation.RowCount - 1 ) then
  begin
  // Canvas.Font.Style := Canvas.Font.Style + [fsItalic] ;
   Canvas.Font.Color := clInactiveCaption;
  end
   else
    begin
  //   Canvas.Font.Style := Canvas.Font.Style - [fsItalic] ;
     Canvas.Font.Color := clWindowText;
    end
   end;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Suppression des lignes d'imputations
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.HGImputationDeleteRow( ARow : integer);
begin

 if ARow = ( HGImputation.RowCount - 1 )  then
  exit;

 if HGImputationAssignLigneTOB(ARow) and FZEtebac.DeleteCurrentLigneImputation then
  begin
   HGImputation.Objects[0,ARow] := nil;
   HGImputation.DeleteRow(ARow); // on supprime la ligne
   if HGImputation.RowCount = 2 then // si on n'est pas sur la derniere cellule on remonte d'une ligne
    HGImputation.Row := 1;

   HGImputation.Refresh;
   NumeroteLigne(ARow,HGImputation); // renumeroter les lignes
   HGImputationAssignTOB;
   AfficheSoldeImputation;           // on recalcul le nouveau solde
   StatutGrilleImput := taConsult;
  end; // if

end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Remet la grille de releve en edition ( suppression de la
Suite ........ : fonction de selection globale )
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.RemettreEnEditionHGSaisie;
begin

 // on desactive les evenements sur les touches
 HGSaisie.OnKeyPress            := nil;
 HGSaisie.OnKeyDown             := nil;
 Ecran.OnKeyDown                := nil;
 HGImputation.OnExit            := nil; // le MontreEdit redeclenche cet evenement

 HGSaisie.CacheEdit ;
 HGSaisie.Options := HGSaisie.Options - [goRowSelect] + [GoEditing,GoAlwaysShowEditor];
 HGSaisie.Invalidate;
 HGSaisie.MontreEdit;
 HlbNomGuide.Visible := false;

 HGImputation.Enabled := false;

 if HGSaisie.Col = HGSaisie.ColCount - 1 then
  begin
   HGSaisie.Col := HGSaisie.Col - 1 ;
   HGSaisie.Col := HGSaisie.Col + 1 ;
  end
   else
    begin
     HGSaisie.Col := GetCol ;
     HGSaisie.Col := HGSaisie.Col + 1 ;
     HGSaisie.Col := HGSaisie.Col - 1 ;
    end;

  // on rebranche les evenements
  HGSaisie.OnKeyPress            := HGSaisieKeyPress;
  HGSaisie.OnKeyDown             := HGSaisieKeyDown;
  Ecran.OnKeyDown                := FormKeyDown;
  HGImputation.OnExit            := HGImputationExit;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/10/2002
Modifié le ... :   /  /    
Description .. : -28/10/2002 - correction deu calcul du solde
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.CalculSoldeTheorique;
var
 lStSQL     : string;
 Q          : TQuery;
 lRdVar1    : double;
 lRdVar2    : double;
begin
  // calcul du solde du compte de contrepartie
 CGetBalanceParcompte('',FZEtebac.StCompteContrepartie,StrToDate(EdtMulDateDu.Text) , StrToDate(EdtMulDateAu.Text),lRdVar1,lRdVar2,FRdSolDeT);

  // calcul du solde du releve en ne tenant compte que des dates
 lStSQL := ' select '                                                           +
           'sum(CRL_DEBIT - CRL_CREDIT) as D '                                  +
           'from CRELBQE '                                                      +
           'where CRL_GENERALBQE = "' + FZEtebac.StCompteContrepartie + '"'     +
           ' and CRL_DATECOMPTABLE >= "' + UsDate(EdtMulDateDu) + '"'           +
           ' and CRL_DATECOMPTABLE <= "' + UsDate(EdtMulDateAu) + '"'           +
           ' and CRL_NUMLIGNE = 0 ';
 Q := nil;

 try

   Q := openSQL( lStSQL, true );
   FRdSolDeT := FRdSolDeT - Q.FindField('D').asFloat;

 finally
   if assigned(Q) then Ferme(Q);
 end; // try

end;

procedure TOF_SAISIEETEBAC.BCompteClick(Sender: TObject);
begin
 ZoomGeneral;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 23/07/2007
Modifié le ... :   /  /
Description .. : - LG - 23/07/2007 - FB 21137 - on appel les guides avec le 
Suite ........ : journal et l'etablissement courant
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.BGuideClick(Sender: TObject);
begin
 CPLanceFiche_MulGuide('','','CLASSIC;'+CmbJournal.Value+';;;'+CmbEtablissement.Value+';X') ;
 FZEtebacGuide.ClearGuide;
 FZEtebacGuide.Load;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 20/06/2007
Description .. : Si on etait en modif ou en creation -> msg pour prevenir
Suite ........ : l'utilisateur
Suite ........ : - LG - FB 18870 - on ne reaffecte plus les ligne qd on fait 
Suite ........ : echap
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin


 CanClose := ( StatutGrilleReleve = taConsult ) and ( StatutGrilleImput = taConsult );

 if not CanClose then
  CanClose := PGIAskCancel(cStTexteEchap,cStTexteTitreFenetre) = mrYes;

 if CanClose then
  TFVierge(Ecran).FormCloseQuery(Sender,CanClose) ;


end;

procedure TOF_SAISIEETEBAC.ZoomGeneral;
begin
 FicheGene(nil, '', FZEtebac.StGENERAL_R, taConsult, 0)
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Zoom sur le compte de contrepartie
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.ZoomEcrs;
var
 lStExo : string;
begin
 // Exercice
 // -1 = précédent
 // 0 = courant
 // 1 = suivant
 if ( VH^.CPExoRef.Code=VH^.Encours.Code ) then
  lStExo := '0'
   else
    if ( VH^.CPExoRef.Code=VH^.Suivant.Code ) then
     lStExo := '1'
      else
       lStExo:= '-1' ;
 OperationsSurComptes( FZEtebac.StCompteContrepartie, lStExo , '', '') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 13/06/2005
Description .. : Ajout des menus supplementaire
Suite ........ : 
Suite ........ : - 28/05/2002 - suppression du raccourci sur le menu
Suite ........ : passage en selection
Suite ........ : - FB 14586 - 13/06/2005 - 
Suite ........ : Calculer l'imputation			F6
Suite ........ : Accéder au détail de l'imputation	ALT+D
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.POPSPopup ( Sender : TObject) ;
var
 i   : integer;
 lT  : TMenuItem ;
 lRS : TShortCut ;
begin

 InitPopUp(Ecran);

 // Ajout des shortcuts au menu cree par InitPopUp
 for i := 0 to ( POPS.Items.Count - 1 ) do
  begin
   lT  := POPS.Items[i];
   lRS := 0;

   if UpperCase( lT.Name ) = 'PTAG' then
    lRS := ShortCut( ord('A') , [ssCtrl] )
     else
      if UpperCase( lT.Name ) = 'PINTEGRER' then
       lRS := ShortCut( VK_F10 , [] )
        else
         if UpperCase( lT.Name ) = 'PDELETE' then
          lRS := ShortCut( VK_DELETE , [ssCtrl] );

   if lRS <> 0 then
    lT.Caption:= lT.Caption + #9 + ShortCutToText(lRS) ;

  end; // for

 lT           := CreerLigPop(TBitBtn(GetControl('BF6')),PopS,True,'P') ;
 lT.Caption   := lT.Caption + #9 + ShortCutToText(ShortCut( VK_F6 , [] )) ;
 PopS.Items.Insert(1,lT) ;
 if FZEtebac.StETAT_R <> cPasImput  then
  begin
   lT           := CreerLigPop(TBitBtn(GetControl('BALTD')),PopS,True,'P') ;
   lT.Caption   := lT.Caption + #9 + ShortCutToText(ShortCut( ord('D') , [ssAlt] )) ;
   PopS.Items.Insert(1,lT) ;
  end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Réaffectation de la table après annulation
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.ReaffecterLigneReleve;
begin

 try

  FZEtebac.DeleteCurrentImputation;
  FZEtebac.TOBLigneReleve.Dupliquer(FTOBSauveLigne,true,true);
  HGSaisieRefreshLigne(HGSaisie.Row);
  StatutGrilleReleve := taModif;
 except
  on E : exception do
   begin
     PGIBox( 'Erreur sur la réaffectation des lignes' + #13#10 +
             E.Message ,
             cStTexteTitreFenetre);
     V_PGI.IoError := oeSaisie;
     {$I-}
    // FTOBSauveLigne.SaveToFile('c:\FTOBSauveLigne.log',false,true,true);
     {$I+}
   end; // on

 end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Passage en mode selecion
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.BModeClick(Sender: TObject);
begin
 if ( cmbEtablissement.Value = '' ) then
  begin
   PGIInfo('Vous devez choisir un établissement avant d''intégrer les écritures !') ;
   if CmbEtablissement.CanFocus then
    CmbEtablissement.setFocus ;
   exit ;
  end ;
 FBoModeSelection := true;
 SetModeSelection;
end;

procedure TOF_SAISIEETEBAC.BFermerClick(Sender: TObject);
begin
 BDeSelectionnerClick(nil);
 FBoModeSelection := false;
 SetModeSelection;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Passage en mode selection pour integrer les ecritures
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.SetModeSelection;
begin

 if Ecran = nil then exit ;
 
 HGSaisie.MultiSelect := FBoModeSelection;
 EnabledControl;

 if FBoModeSelection then
  begin
   HGSaisie.Options := HGSaisie.Options + [goRowSelect] - [GoEditing,GoAlwaysShowEditor];
   // on desactive les evenements sur les touches
   HGSaisie.OnKeyPress            := nil;
   HGSaisie.OnKeyDown             := nil;
   HGImputation.OnExit            := nil; // le MontreEdit redeclenche cet evenement
  end
   else
    begin
     HGSaisie.Options := HGSaisie.Options - [goRowSelect] + [GoEditing,GoAlwaysShowEditor];
      // on rebranche les evenements
     HGSaisie.OnKeyPress            := HGSaisieKeyPress;
     HGSaisie.OnKeyDown             := HGSaisieKeyDown;
     Ecran.OnKeyDown                := FormKeyDown;
     HGImputation.OnExit            := HGImputationExit;
      if HGSaisie.Col = HGSaisie.ColCount - 1 then
       begin
        HGSaisie.Col := HGSaisie.Col - 1 ;
        HGSaisie.Col := HGSaisie.Col + 1 ;
       end
        else
         begin
          HGSaisie.Col := HGSaisie.Col + 1 ;
          HGSaisie.Col := HGSaisie.Col - 1 ;
         end;

    end; // if

end;



{
 Integration des releves etebac
}


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 14/09/2004
Description .. : Transfert des lignes de CETEBAC vers CRELBQE
Suite ........ : - LG - 09/06/2004 - FB 10719 - si les ecritrues n'ont pas de
Suite ........ : guide, ellles sont sois detruites, génération avec compte
Suite ........ : d'attente ou rien
Suite ........ : - LG - 14/09/2004 - FB 14584 - affectation du code 
Suite ........ : etablissement
Suite ........ : - LG - 14/09/2004 - FB 14584 - c'etait le libelle de l'etablissemnt
Suite ........ : qui etait repris
Mots clefs ... : 
*****************************************************************}
function TOF_SAISIEETEBAC.Integration( vNumEtebac : integer ; vTOBLettrage : TOB ) : boolean;
var
 lTOBLigneEtebac            : TOB;
 lTOBEtebac                 : TOB;
 lTOB                       : TOB;
 Q                          : TQuery;
 lInNumLigne                : integer;
 lStChampsDateComptable     : string;
 lStChampsDateValeur        : string;
 lStTraitementImport        : string ;
begin

 result           := true;

 if (ctxPCL in V_PGI.PGIContexte) and not VH^.OkModExpertETEBAC then exit ;

 Q                := nil;
 lTOBEtebac       := TOB.Create('',nil,-1);

 try
 try
  try

   Q := OpenSQL('select CET_LIBELLE, '        +
                'CET_REFPIECE, '              +
                'CET_NUMEROPIECE,'            +
                'CET_DATEVALEUR, '            +
                'CET_DATEOPERATION, '         +
                'CET_DEVISE, '                +
                'CET_CODEAFB, '               +
                'CET_REFPIECE, '              +
                'CET_DEBIT, '                 +
                'CET_CREDIT, '                +
                {JP 19/01/07 : Gestion des libellés enrichis}
                'CET_LIBELLE1, '              +
                'CET_LIBELLE2, '              +
                'CET_LIBELLE3 '               +
                'from CETEBAC '               +
                'where CET_NUMRELEVE = '      + intToStr(vNumEtebac)      +
                ' and CET_TYPELIGNE = "04" '  +
                'and CET_DEVISE = "'          + V_PGI.DevisePivot         + '" ' ,
                true );


  lTOBEtebac.LoadDetailDB('CETEBAC','','',Q,false);

 finally
  if assigned(Q) then Ferme(Q);
 end; // try

 if lTOBEtebac.Detail.Count = 0 then exit;

 // chargement des info sur le journal
 if FInfoEcr.Journal.Load([FZEtebac.StJournalContrepartie ]) = - 1 then exit ;
 lStChampsDateComptable := FInfoEcr.Journal.GetValue('J_CHOIXDATE') ;
 lStTraitementImport    := FInfoEcr.Journal.GetValue('J_TRESOIMPORT') ;

  if lStChampsDateComptable = '' then
   lStChampsDateComptable := 'DATEOPERATION';

  if lStChampsDateComptable = 'DATEOPERATION' then
   lStChampsDateValeur := 'DATEVALEUR'
    else
     lStChampsDateValeur := 'DATEOPERATION';

 lInNumLigne     := 0;

 while ( lTOBEtebac.Detail.Count ) > lInNumLigne do
  begin
   FZEtebac.AddLigneReleve;
   lTOBLigneEtebac                      := lTOBEtebac.Detail[lInNumLigne];
   FZEtebac.StIMPORT_R                  := 'X';
   FZEtebac.StGENERALBQE_R              := FZEtebac.StCompteContrepartie; // compte de banque
   FZEtebac.StJOURNAL_R                 := FZEtebac.StJournalContrepartie;
   FZEtebac.DtDATECOMPTABLE_R           := lTOBLigneEtebac.GetValue('CET_' + lStChampsDateComptable); //
   FZEtebac.DtDATEVALEUR_R              := lTOBLigneEtebac.GetValue('CET_' + lStChampsDateValeur); //
   FZEtebac.StREFINTERNE_R              := lTOBLigneEtebac.GetValue('CET_NUMEROPIECE');   // numero de cheque ou de piece
   CSetMontants( FZEtebac.TOBLigneReleve,
                 lTOBLigneEtebac.GetDouble('CET_DEBIT') ,
                 lTOBLigneEtebac.GetDouble('CET_CREDIT') ,
                 FZEtebac.Devise,
               //  false,
                 true );
   FZEtebac.StNATUREPIECE_R             := 'OD' ;
   FZEtebac.StLIBELLE_R                 := lTOBLigneEtebac.GetValue('CET_LIBELLE');
   FZEtebac.StCODEAFB_R                 := lTOBLigneEtebac.GetValue('CET_CODEAFB');
   FZEtebac.StREFPIECE_R                := lTOBLigneEtebac.GetValue('CET_REFPIECE');
   FZEtebac.StDEVISE_R                  := lTOBLigneEtebac.GetValue('CET_DEVISE');

   {JP 19/01/07 : Gestion des libellés enrichis}
   FZEtebac.StLibEnrichi1_R := lTOBLigneEtebac.GetString('CET_LIBELLE1');
   FZEtebac.StLibEnrichi2_R := lTOBLigneEtebac.GetString('CET_LIBELLE2');
   FZEtebac.StLibEnrichi3_R := lTOBLigneEtebac.GetString('CET_LIBELLE3');

   if Ecran <> nil then
    FZEtebac.StETABLISSEMENT_R          := CmbEtablissement.Value
     else
      FZEtebac.StETABLISSEMENT_R        := VH^.EtablisDefaut ;
   FZEtebac.InCOMPTEUR_R                := FInNumReleve;
   FInNumReleve                         := FInNumReleve + 1;

   lTOB                                 := TOB.Create( '' , nil , -1 );

   result := FZEtebacGuide.RechercheGuideEnBase( FZEtebac.TOBLigneReleve , lTOB , nil ) ;

  if not result then  // il n'existe pas de guide ou plusieurs guide pour cette saisie
   begin // FB 10719
    if FInfoEcr.Journal.GetValue('J_TRESOECRITURE') = 'X' then
     begin
      if not FZEtebac.CreationDepuisEcriture(vTOBLettrage) then
       begin
        if lStTraitementImport = '2TS' then
         FZEtebac.DeleteLigneReleve
          else
           if lStTraitementImport = '3TS' then
            FZEtebac.CreationAvecCompteAttente ;
       end ;
     end
      else
       if lStTraitementImport = '2TS' then
        FZEtebac.DeleteLigneReleve
         else
          if lStTraitementImport = '3TS' then
           FZEtebac.CreationAvecCompteAttente ;
   end
    else
     begin
      // on affecte les enregistrement contenu dans lTOB à la ligne de releve courante
      FZEtebac.CreerImputation(lTOB);
      // on supprime les lignes vides
     // FZEtebac.SupprimerLigneImputVide;
      FZEtebac.CalculSoldeImput;

      if FZEtebac.IsValideImputation then
       FZEtebac.StETAT_R := cImputCorrectNonValide
        else
         FZEtebac.StETAT_R := cImputIncorrecte;

      end; // if

  lTOB.Free ;
  lInNumLigne  := lInNumLigne + 1;

   // gestion de la fenetre d'attente
  MoveCurProgressForm('Integration du releve ' + IntToStr( FInNumReleve )) ;

 end; // while

 finally
  if assigned(lTOBEtebac) then lTOBEtebac.Free;
 end; // try

 except
  on E : Exception do
  PGIError('Erreur lors de la génération des écritures par rapport aux guides/ecritures/comptes d''attente !' + #10#13 + E.Message ) ;
 end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Enregistrement en base des lignes de releves et
Suite ........ : suppression des lignes CETEBAC
Mots clefs ... :
*****************************************************************}
procedure TOF_SAISIEETEBAC.TransactionSaveReleve;
var
 lStValue : string;
begin

  if FZEtebac.SaveReleve then
   begin

    if not CMAJNumeroSouche ('TRE', 'NOR' ,FInNumReleve - 1,FInFirstNumReleve) then
     raise EAbort.Create('Erreur a la maj des compteurs');

    // suppression des lignes etebac integrees
    while ( FStListeNumASupp <> '' ) do
     begin

      lStValue     := ReadTokenST(FStListeNumASupp);
      {$IFNDEF TT}
      {$IFNDEF JOHN}
       if lStValue <> '' then
        ExecuteSQL('delete CETEBAC where CET_NUMRELEVE = ' + lStValue);
      {$ENDIF}
      {$ENDIF}

     end; // while

   end; // if

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... : 22/08/2007
Description .. : Recherche des lignes CETEBAC correspondant au RIB du
Suite ........ : journal de abnque
Suite ........ : - FB 13015 - 10/06/2004 - on peu chainer l'enregistrment
Suite ........ : apres l'importation
Suite ........ : - FB 12520 - passage d'un parametre supp a la fenetre
Suite ........ : d'affichage des rib : le journal. Pour le controlde de relevel
Suite ........ : deja presnet en base
Suite ........ : - FB 20561 - LG - 17/07/2007 - gestion des etablissemtn ds 
Suite ........ : les guides eteabc, on passe l'etablissemnt selectionne a la 
Suite ........ : fenetre d'integration
Suite ........ : - LG - 22/08/2007 - FB 21026 - suppression d emsg 
Suite ........ : intempestif
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.RechercheRibAIntegrer;
var
 Q               : TQuery;
 lStETABQ        : string;
 lStGUICHET      : string;
 lstNUMEROCOMPTE : string;
 lStListeNum     : string;
 lStValue        : string;
 lTempData       : TObject;
 T               : TTemp;
 lInNumEtabac    : integer;
 lTOBLettrage    : TOB ;
 lStEta          : string ;
begin

  Q            := nil;
  lTOBLettrage := nil ;
{$IFDEF TT}
 //AddEvenement('RechercheRibAIntegrer' );
 //CPLanceFiche_ImportCEtebac(CmbContrepartie.Text);
{$ENDIF}

  try
  {$IFDEF TT}
     Q := OpenSQL(  'select CET_NUMLIGNE, '                         +
                  'CET_ETABBQ, '                                  +
                  'CET_NUMEROCOMPTE, '                            +
                  'CET_GUICHET '                                  +
                  'from CETEBAC  '                      +
                  'where '       +
                  ' CET_TYPELIGNE = "01" '                     +
                  'and CET_DEVISE = "'                            + V_PGI.DevisePivot                   + '" '
                  ,
                  true );

  {$ELSE}
   Q := OpenSQL(  'select CET_NUMLIGNE, '                         +
                  'CET_ETABBQ, '                                  +
                  'CET_NUMEROCOMPTE, '                            +
                  'CET_GUICHET '                                  +
                  'from CETEBAC, BANQUECP  '                      +
                  'where CET_ETABBQ = BQ_ETABBQ '                 +
                  'and CET_NUMEROCOMPTE = BQ_NUMEROCOMPTE '       +
                  'and CET_GUICHET = BQ_GUICHET '                 +
                  'and BQ_GENERAL = "'                            + FZEtebac.StCompteContrepartie + '"' +
                  'AND BQ_NODOSSIER = "'                          + V_PGI.NoDossier+'" '   + // 19/10/2006 YMO Multisociétés
                  'and CET_TYPELIGNE = "01" '                     +
                  'and CET_DEVISE = "'                            + V_PGI.DevisePivot                   + '" '
                  ,
                  true );
   {$ENDIF}

   if Q.EOF then exit;

   if (ctxPCL in V_PGI.PGIContexte) and not VH^.OkModExpertETEBAC then
    begin
     PGIInfo('Des relévés sont en attente d''intégration.'+#10#13 + 'Le module Cegid Expert Etebac (code 00609) n’a pas été sérialisé') ;
     exit ;
    end ;

   lStETABQ        := Q.FindField('CET_ETABBQ').asString;
   lStGUICHET      := Q.FindField('CET_GUICHET').asString;
   lstNUMEROCOMPTE := Q.FindField('CET_NUMEROCOMPTE').asString;


   if assigned(Q) then Ferme(Q);
   Q := nil;

   lStListeNum     := '';

   lTempData       := TheData;

   if Ecran <> nil then
    begin
    lStEta := CmbEtablissement.Value ;
    if lStEta = '' then lStEta := 'T' ;
    // lance la fiche d'integration des releves fenetre RLVETEBAC et TOF  TOF_CETEBAC ( fichier DP\lib\dpTOFCETEBAC )
    CPLanceFiche_MulCEtebac(lStETABQ + ';' + lStGUICHET       + ';' + lstNUMEROCOMPTE  + ';' + 'N' + ';' +
                            FZEtebac.StJournalContrepartie + ';' + lStEta );

    T               := TTemp(TheData);

    if not assigned(T) then exit;

    if ( Ecran <> nil ) and ( FZEtebacGuide.EstVide ) and ( PGIAsk(cStTexteAucunGuideEtebac , cStTexteTitreFenetre ) = mrNo ) then exit;

    lStListeNum              := T.StValue;
    FStListeNumASupp         := lStListeNum;
    FZEtebacGuide.SilentMode := true ;
    CmbEtablissement.Value   := T.StEta ;
    FZEtebacGuide.SilentMode := false ;
    T.free;

    theData            := lTempData; // restauration de la valeur de theData
   end ;

   // l'utilisateur n'a pas integre de releve
   if lStListeNum = '' then
    exit;

   if Blocage([cStVerrouTreso], true , cStVerrouTreso ) then
    exit;

   FInNumReleve                 := FZEtebac.ProchainCompteur;
   FInFirstNumReleve            := FInNumReleve;
   FZEtebacGuide.StCompteTVA    := FZEtebac.StCompteTVA;
 //  FZEtebac.TypeContexte        := TModeAuto;

   if Ecran <> nil then
   InitMoveProgressForm (	Ecran,
                                cStTexteTraitementEcr,
                                cStTexteTitreFenetre,
                                100,
                                true,
                                true ) ;


  lTOBLettrage := TOB.Create('',nil,-1) ;

   while ( lStListeNum <> '' ) do
    begin
     lStValue     := ReadTokenST(lStListeNum);
     if lStValue <> '' then
      begin
       lInNumEtabac := StrToInt(lStValue);
       Integration(lInNumEtabac,lTOBLettrage);
      end; // if
    end; // while


   // gestion de la fenetre d'attente
  if ( Ecran <> nil )  and not MoveCurProgressForm('Enregistrement en base') then Exit ;

  if Transactions(TransactionSaveReleve,1) <> oeOK then
   MessageAlerte('Erreur lors de l''enregistrement des fichiers' + #13#10 + V_PGI.LastSQLError );

  if ( FInfoEcr.Journal.GetValue('J_TRESOCHAINAGE') = 'X' ) then
    begin
    FBoParle := false ;
    FBoChainage := true ;
    Integrer(lTOBLettrage) ;
    FBoChainage := false ;
    FBoParle := true ;
   end ;

 finally
  if assigned(Q)    then Ferme(Q);
  if Ecran <> nil then FiniMoveProgressForm ;
  if assigned(lTOBLettrage) then lTOBLettrage.Free ;
  Bloqueur( cStVerrouTreso , false );
 end; // try

end;

procedure TOF_SAISIEETEBAC.BExportClick(Sender: Tobject);
var
  lStHint: string;
begin
  if not ExJaiLeDroitConcept(ccExportListe, True) then
    exit;
  if SD.Execute then
  begin
    if SD.FilterIndex = 5 then //html
    begin
      lStHint := HGSaisie.Hint;
      HGSaisie.Hint := Ecran.Caption;
      ExportGrid(HGSaisie, nil, SD.FileName, SD.FilterIndex, True);
      HGSaisie.Hint := lStHint;
    end
    else
      ExportGrid(HGSaisie, nil, SD.FileName, SD.FilterIndex, True);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 28/09/2005
Modifié le ... :   /  /    
Description .. : - LG - FB 16737 - 28/09/2005- plantait ne import auto 
Suite ........ : depuis le bureau
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.OnError (sender : TObject; Error : TRecError ) ;
begin

 if FZEtebac.TypeContexte = TModeAuto then exit ;

 if (trim(Error.RC_Message) <> '' ) and ( Ecran <> nil ) then
  begin
   if HGSaisie.Row <> HGSaisie.RowCount - 1 then
    PGIInfo(Error.RC_Message, cStTexteTitreFenetre )
  end
  else
   if (Error.RC_Error<>RC_PASERREUR) then
    FMessCompta.Execute(Error.RC_Error) ;
//    CPrintMessageRC(cStTexteTitreFenetre,Error.RC_Error,FListeMessage) ;
end;


function TOF_SAISIEETEBAC.GetCol : integer ;
begin
 if (( FZEtebac.StIMPORT_R  = 'X' ) and ( FInfoEcr.Journal.GetValue('J_TRESODATE') = '-' )) then
  result := cColDate + 1
  else
   result := cColDate ;
end ;



procedure TOF_SAISIEETEBAC.OnClickAltD(Sender: TObject);
begin
if Screen.ActiveControl = HGSaisie then
 begin
  if not EnregistreLigne(HGSaisie.Row) then exit;
  if not AltDImputation                then HGSaisie.SetFocus;
 end;
end;

procedure TOF_SAISIEETEBAC.OnClickF6(Sender: TObject);
begin
if Screen.ActiveControl = HGSaisie then
 begin
  if ( StatutGrilleReleve = taConsult ) then // sauvegarde de la ligne courante
   FTOBSauveLigne.Dupliquer(FZEtebac.TOBLigneReleve,true,true);
  if not EnregistreLigne(HGSaisie.Row) then exit;
  AfficheGuide;
 end // if
  else
   if ( Screen.ActiveControl = HGImputation ) and ( HGImputation.Row <> ( HGImputation.RowCount - 1 )) then
    VK_F6Imputation; // calcul du solde de la ligne mais pas sur la derniere ligne ( ligne de contrepartie non modifiable
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.ApresChargementFiltre;
begin
  if EdtFFiltres.Text <> '' then BChercheClick(BCherche);
end;

procedure TOF_SAISIEETEBAC.BImportClick(Sender: Tobject);
begin
 if (ctxPCL in V_PGI.PGIContexte) and not VH^.OkModExpertETEBAC then
  begin
   PGIInfo('Des relévés sont en attente d''intégration.'+#10#13 + 'Le module Cegid Expert Etebac (code 00609) n’a pas été sérialisé') ;
   exit ;
  end ;
 CPLanceFiche_ImportCEtebac(CmbContrepartie.Text);
 RechercheRibAIntegrer ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 18/10/2007
Modifié le ... :   /  /    
Description .. : - LG - 18/10/2007 - FB 21136 - creation d'un guide a partir 
Suite ........ : d'une ecriture
Mots clefs ... : 
*****************************************************************}
procedure TOF_SAISIEETEBAC.BCreerGuideClick(Sender: Tobject) ;
var
 lStCode : string ;
begin

 if FBoGuideRun then exit ;

 FZEtebac.TypeContexte := TModeSaisie;
 if not FZEtebac.IsValideImputation then
  exit;

 lStCode := CCreerCommeGuide( FZEtebac.TOBImput,'CRL_' ) ;
 if lStCode <> '' then
  begin
   ParamGuide( lStCode , 'NOR' , taModif ) ;
   FZEtebacGuide.ClearGuide;
   FZEtebacGuide.Load;
  end ;
end ;

////////////////////////////////////////////////////////////////////////////////

Initialization
 registerclasses ( [ TOF_SAISIEETEBAC ] ) ;
end.
