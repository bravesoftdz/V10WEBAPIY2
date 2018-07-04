{***********UNITE*************************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CRELBQE ()
Mots clefs ... : TOF;CRELBQE
*****************************************************************}
Unit UTOFCRELBQE ;

Interface

Uses StdCtrls, Controls, Classes,
     {$IFNDEF EAGLCLIENT}
     db,dbTables,
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
     Grids,  // pour TGridDrawState
     extctrls, // pour le TBevel
     HPanel,
     // lib
     Ent1,
     ParamDat,
     LookUp,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF SANSCOMPTA}
     ZCompte,
     FE_Main,   // pour AGLLanceFiche()
     SaisAnal,
     HCompte,
     General,  // FicheGene ( fonction de Zoom des comptes )
     PrintDBG, // pour le prindDB
     ConsEcr, // pour les zooms sur les ecritures
     Zecrimvt, // pour ZoomEcr
     Devise, // pour FicheDevise
     {$ELSE}
     MaineAGL,
     {$ENDIF}
     {$ENDIF}
     SaisUtil,  // pour RDevise
     ParamSoc,  // pour le GetParam
     Filtre, // gestion du filtre
     AGLInit,
     Vierge, // pour le form closeQuery
     TZ,
     ed_tools, // pour le TQRProgressForm
     // Objet de gestion
     ZReleveBanque,
     ZGuideTreso,
     UCstEtebac,
     ULibEcriture,
     ULibExercice,
     ULibWindows,
     ZTresoVersEcr,
     Guide,
     // AGL
     UTOB,
     UTOF ;

Type


  TOF_CRELBQE = Class (TOF)
   private
    // variable interne
    FStatutGrilleReleve     : TActionFiche ;          // Etat de la grille ( modif,creation...)
    FStatutGrilleImput      : TActionFiche ;          // Etat de la grille ( modif,creation...)
    FBoGuideRun             : boolean;                // True la saisie par guide est active
    FCurrentCol             : integer;                // Colonne courante de la grille d'imputation
    FCurrentRow             : integer;                // Ligne couranet de la grille d'imputation
    FRdNumeroPieceMax       : double;                 // Numero de piece max de la grille : sert dans pour le Shift+F7
    FRdNumeroPieceMaxOld    : double;                 // on utilise cette variable pour retrouver le precedent numero en cas d'annulation
    FDtDateDebN             : TDateTime;              // Date de debut de l'exercice N
    FDtDateFinN1            : TDateTime;              // Date de fin de l'exercice N+1
    FRdSolDeT               : double;                 // solde théorique
    FBoChoixEffectue        : boolean;                // La saisie est bloque tant que l'on n'a pas choisit un compte de banque et de contrepartie
    FBoGrilleVide           : boolean;                // true si la grille est en creation mais que l'on n'a rien enregistré
    FBoModeGuideAuto        : boolean;
    FInNumReleve            : integer;
    FInFirstNumReleve       : integer;
    FBoModeSelection        : boolean;
    FStListeNumASupp        : string;
    FListeIntegrer          : TStringList;            // StringList contenant le code des lignes à intégrer dans la tabel ecriture
    FExceptionEvent         : TExceptionEvent;        // evenement gerant toutes les exceptions non gérées de la saisies
    FTT                     : TQRProgressForm;
    // liste des TOB
    ZReleveBanque           : TZReleveBanque;         // TOB de gestion de la grille de saisie
    ZGuideTreso             : TZGuideTreso;          // TOB de gestion de la grille d'imputation
    FTOBSauveLigne          : TOB;
    // liste des contrôles
    PgcCritere              : TPageControl;           // Page à onglet de recherche
    TbsStandard             : TTabSheet;              // Onglet standard
    HGSaisie                : THGrid;                 // Grille de saisie
    HGImputation            : THGrid;                 // Grille d'imputation
    PFiltres                : TToolWindow97;          // Fenêtre contenant l'enregistrement des filtres
    PBouton                 : TToolWindow97;          // Fenêtre contenant les boutons
    BFermer2                : TToolbarButton97;       // Bouton de fermeture du mode d'intégration
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
    POPF                    : TPopupMenu;
    POPS                    : TPopupMenu;
    POPZOOM                 : TPopupMenu;
    BCreerFiltre            : TMenuItem;
    BSaveFiltre             : TMenuItem;
    BDelFiltre              : TMenuItem;
    BRenFiltre              : TMenuItem;
    BNouvRech               : TMenuItem;
    BPopCompte              : TMenuItem;
    BPopJournal             : TMenuItem;
    FlashCompte             : TFlashingLabel;
    FlashJournal            : TFlashingLabel;
    HlbNomGuide             : THLabel;               // label affichant le nom du guide
    ImgEuro                 : TImage;                // Image du sigle EURO
    HlbCaptionSoldeImput    : THLabel;
    BIntegrer               : TToolbarButton97;       // Bouton BCherche
    BSelectionner           : TToolbarButton97;       // Bouton BCherche
    BDeSelectionner         : TToolbarButton97;       // Bouton BCherche
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
    procedure POPFPopup(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BIntegrerClick(Sender: TObject);
    procedure BSelectionnerClick(Sender: TObject);
    procedure BDeSelectionnerClick(Sender: TObject);
    procedure BPopCompteClick(Sender: TObject);
    procedure BPopJournalClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
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
    // recherche du tiers
    function  HGImputationRechercheTiers : boolean;
    // recherche imputation
    function  HGImputationRechercheImputation( ARow : integer ) : boolean;
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
    // On recherche l'ensemble des comptes de contrepartie des
    // journaux de type banques et on calcule le solde de la
    // periode precedente du compte
    procedure RemplirComboBoxContrepartie;

    // Value = True -> on affecte les evenement à la grille   Value = false on debranche tous les evenements ( sert lors de l'ajout de nouvelle ligne
    procedure SetEvent( value : boolean);
    // gestion des touches VK_DOWN,VK_F5,VK_RETURN dans les grille de saisie
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    // lance la recherche d'un guide et remplit la grille des imputations
    function  AfficheGuide : boolean;
    // Ajout d'un nouvelle enregistrement dans la grille des releves
    procedure Nouveau;
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
    procedure ShiftVK_F7Grille( G : THGrid );
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
    function  Integration(vNumEtebac: integer): boolean;
    procedure ReaffecterLigneReleve;
    procedure SetModeSelection;
    procedure TransactionSaveReleve;
    procedure GestionException (Sender: TObject; E: Exception);
    procedure HGImputationRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    function  SelectionnerLigne : boolean;
    procedure ChargeParam;
    procedure AfficheLookUpImput;
   protected
    function  GetControlTOF           : boolean; // recupere les contrôles de la grille
    function  AssignEvent             : boolean; // Affecte les evenements au contrôle
    function  InitControl             : boolean; // initialise les grilles
    function  CreateControl           : boolean; // creation de TOB
  public
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    property  StatutGrilleReleve : TActionFiche read FStatutGrilleReleve write SetStatutGrilleReleve;
    property  StatutGrilleImput  : TActionFiche read FStatutGrilleImput  write SetStatutGrilleImput;

  end ;

  procedure CPLanceFiche_SaisieTresorerie;

Implementation

  type
  TTemp = Class
   StValue : string;
  end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... : 20/07/2001
Description .. : Poinr d'entre de la saisie de tresorie
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_SaisieTresorerie;
begin

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

 AGLLanceFiche('CP','RLBSAISIE','','','') ;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... :   /  /
Description .. : Rem : on rentre directement en création
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.OnArgument (S : String ) ;
begin
 Inherited ;

  // on sauvegarde la valeur courante
 FExceptionEvent         := Application.OnException;
 // on branche le nouveau gestionnaire d'exception
 Application.OnException := GestionException;

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

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Destruction des objets
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.OnClose ;
begin

 // on sauvegarde le numero de dossier max
 SetParamSoc('SO_CPTRESONUMPIECEMAX',FRdNumeroPieceMax) ;

 try

  if assigned(ZGuideTreso)       then ZGuideTreso.Free;
  if assigned(FListeIntegrer)    then FListeIntegrer.Free;
  if assigned(FTOBSauveLigne)    then FTOBSauveLigne.Free;
  if assigned(ZReleveBanque)     then ZReleveBanque.Free;

 finally
  ZReleveBanque                  := nil;
  ZGuideTreso                    := nil;
  FListeIntegrer                 := nil;
  FTOBSauveLigne                 := nil;
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
procedure TOF_CRELBQE.GestionException ( Sender : TObject ; E : Exception);
begin
 MessageAlerte( 'Fermer la fenêtre ! ' + #13#10 + #13#10 +
               E.Message );
end;

procedure TOF_CRELBQE.ChargeParam;
begin

 // recuperation du numero de piece max en base
 if VarIsEmpty(GetParamSocSecur('SO_CPTRESONUMPIECEMAX',0)) then
  FRdNumeroPieceMax := 0
  else
   FRdNumeroPieceMax                            := GetParamSocSecur('SO_CPTRESONUMPIECEMAX',0);

 FRdNumeroPieceMaxOld                           := FRdNumeroPieceMax;

 // recherche de la derniere date de cloture periodique
 FDtDateDebN  := GetParamSocSecur('SO_DATECLOTUREPER',iDate1900) + 1;
 if FDtDateDebN = 1 then
  FDtDateDebN := VH^.CPExoRef.Deb;

 ChcBouclage.Checked                             := GetParamSocSecur('SO_BOUCLERSAISIECREAT',False);

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 20/07/2001
Description .. : Recupération des contrôles de la fiche
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.GetControlTOF: boolean;

 procedure NotifyErrorComponent(Sender : TComponent; Name: string);
 var
  lStMessage : string;
  begin
   if not assigned(Sender) then
    begin

     lStMessage := TraduireMemoire('Le contrôle ') + Name + TraduireMemoire(' est manquant !' ) + #13#10 +
                   TraduireMemoire(' Vérifiez la version de votre base');
      PgiBox( lStMessage,TraduireMemoire('Erreur création de la fiche') ) ;

      // pour stopper les traitements suivants
      raise EAbort.Create('Erreur a la creation');
    end; //if
  end; // proc

begin

 result                  := false;
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
 BMode                   := TToolbarButton97(GetControl('BMode')) ;
 BImprimer               := TToolbarButton97(GetControl('BImprimer')) ;
 BDefaire                := TToolbarButton97(GetControl('BDefaire')) ;
 BAgrandir               := TToolbarButton97(GetControl('BAgrandir')) ;
 BMenuZoom               := TToolbarButton97(GetControl('BZoom')) ;
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
 POPF                    := TPopupMenu(GetControl('POPF'));
 POPS                    := TPopupMenu(GetControl('POPS'));
 POPZOOM                 := TPopupMenu(GetControl('POPZOOM'));
 BCreerFiltre            := TMenuItem(GetControl('BCreerFiltre'));
 BSaveFiltre             := TMenuItem(GetControl('BSaveFiltre'));
 BDelFiltre              := TMenuItem(GetControl('BDelFiltre'));
 BRenFiltre              := TMenuItem(GetControl('BRenFiltre'));
 BNouvRech               := TMenuItem(GetControl('BNouvRech'));
 HlbNomGuide             := THLabel(GetControl('FE__HLABEL_NOMGUIDE'));
 ImgEuro                 := TImage(GetControl('FE__IMAGE_EURO'));
 HlbCaptionSoldeImput    := THLabel(GetControl('FE__HLABEL_SOLDEIMPUT'));
 BIntegrer               := TToolbarButton97(GetControl('BIntegrer')) ;
 BDeSelectionner         := TToolbarButton97(GetControl('BDETAG')) ;
 BSelectionner           := TToolbarButton97(GetControl('BTAG')) ;
 BPopCompte              := TMenuItem(GetControl('BPopCompte'));
 BPopJournal             := TMenuItem(GetControl('BPopJournal'));
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
 NotifyErrorComponent ( BPopJournal                , 'BPopJournal');
 NotifyErrorComponent ( BPopCompte                 , 'BPopCompte');


 result := true;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 20/07/2001
Description .. : Création des objets de la TOF
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.CreateControl: boolean;
begin
  result                            := false;
  ZReleveBanque                     := TZReleveBanque.Create(nil);
  ZReleveBanque.InitializeObject;
  ZGuideTreso                       := TZGuideTreso.Create;
  FListeIntegrer                    := TStringList.Create;
  FCurrentCol                       := 0;
  FCurrentRow                       := 0;
  result                            := true;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 20/07/2001
Description .. : Initialisation des contrôles
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.InitControl : boolean;
begin

  result := false;

  ChargeParam;

  // tri possible dans la grille
  // correction a mettre dans la socref
  HGSaisie.SortEnabled                               := true;
  CmbMulValeur1.Values[0]                            := 'CRL_DEBIT';
  CmbMulValeur1.Values[1]                            := 'CRL_CREDIT';
  CmbMulValeur2.Values[0]                            := 'CRL_DEBIT';
  CmbMulValeur2.Values[1]                            := 'CRL_CREDIT';
  POPZOOM.Items[0].Caption                           := TraduireMemoire('voir les guides');
  //THPanel(GetControl('FE__HPANEL2')).BackGroundEffect := bdFlat;
 // THPanel(GetControl('FE__HPANEL5')).BackGroundEffect := bdFlat;
 // THPanel(GetControl('FE__HPANEL4')).BackGroundEffect := bdFlat;

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

  HGSaisie.Cells[cColDate,0]                     := TraduireMemoire('Date');
  HGSaisie.Cells[cColRefInterne,0]               := TraduireMemoire('Pièce banc.');
  HGSaisie.Cells[cColDebit,0]                    := TraduireMemoire('Débit');
  HGSaisie.Cells[cColCredit,0]                   := TraduireMemoire('Crédit');
  HGSaisie.Cells[cColLibelle,0]                  := TraduireMemoire('Libellé');
  HGSaisie.Cells[cColPiece,0]                    := TraduireMemoire('Pièce');
  HGSaisie.Cells[cColImputation,0]               := TraduireMemoire('Imputation');
  HGSaisie.Cells[cColTva,0]                      := TraduireMemoire('T');
  HGSaisie.Cells[cColEtat,0]                     := TraduireMemoire('Etat');
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
  PFiltres.Left                                  := PBouton.Left + 70;

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

  FlashGuide                                     := TFlashingLabel.Create(nil);
  FlashGuide.parent                              := HlbDebitImput.Parent;
  FlashGuide.Top                                 := HlbNomGuide.Top - 3;
  FlashGuide.Left                                := HlbNomGuide.Left + HlbNomGuide.Width + 50;
  FlashGuide.Font.Height                         := 20;
  FlashGuide.Font.Style                          := [fsBold];
  FlashGuide.Transparent                         := true;
  FlashGuide.Caption                             := TraduireMemoire('GUIDE');
  FlashGuide.Color                               := clRed;
  FlashGuide.Visible                             := false;


  FlashCurrentReleve                             := TFlashingLabel.Create(nil);
  FlashCurrentReleve.parent                      := HlbDebitSaisie.Parent;
  FlashCurrentReleve.Top                         := 5;
  FlashCurrentReleve.Left                        := 10;
  FlashCurrentReleve.Font.Height                 := 15;
  FlashCurrentReleve.Font.Style                  := [fsBold];
  FlashCurrentReleve.Transparent                 := true;
  FlashCurrentReleve.Caption                     := TraduireMemoire(cStTexteLigneReleve);
  FlashCurrentReleve.Color                       := clActiveCaption;
  FlashCurrentReleve.Visible                     := true;


  FlashCompte                                     := TFlashingLabel.Create(nil);
  FlashCompte.parent                              := CmbJournal.Parent;
  FlashCompte.Top                                 := 11;
  FlashCompte.Left                                := 7;
  FlashCompte.Font.Height                         := 13;
//  FlashCompte.Font.Style                          := [fsBold];
  FlashCompte.Transparent                         := true;
  FlashCompte.Caption                             := TraduireMemoire('Compte bancaire');
  FlashCompte.Color                               := clWindowText;
  FlashCompte.Visible                             := true;


  FlashJournal                                    := TFlashingLabel.Create(nil);
  FlashJournal.parent                             := CmbJournal.Parent;
  FlashJournal.Top                                := 11;
  FlashJournal.Left                               := CmbContrepartie.Left + CmbContrepartie.Width + 10;
  FlashJournal.Font.Height                        := 13;
  //FlashJournal.Font.Style                         := [fsBold];
  FlashJournal.Transparent                        := true;
  FlashJournal.Caption                            := TraduireMemoire('Journal');
  FlashJournal.Color                              := clWindowText;
  FlashJournal.Visible                            := true;

   // charge les filtres par defaut
 ChargeFiltre('RLB_SAISIE',EdtFFiltres,PgcCritere);

  if FDtDateDebN > VH^.CPExoRef.Deb then // si la date de cloture est definie
   EdtMulDateDu.Text                             := DateToStr(FDtDateDebN)
    else
     EdtMulDateDu.Text                           := DateToStr(VH^.CPExoRef.Deb);  // sinon on prend la date de debut exercice

  EdtMulDateAu.Text                              := DateToStr(VH^.CPExoRef.Fin);

  if VH^.Suivant.Fin <> 0 then
   FDtDateFinN1 := VH^.Suivant.Fin
    else
     FDtDateFinN1 := VH^.EnCours.Fin;

 Ecran.PopupMenu                                 := POPS;

 PgcCritere.ActivePageIndex                      := 0;

 FBoGrilleVide                                   := true;

 FBoModeGuideAuto                                := false;

 result                                          := true;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Affectation des evenements au contrôle
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.AssignEvent: boolean;
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

 // gestion du filtre
 BCreerFiltre.OnClick            := BCreerFiltreClick;
 BSaveFiltre.OnClick             := BSaveFiltreClick;
 BDelFiltre.OnClick              := BDelFiltreClick;
 BRenFiltre.OnClick              := BRenFiltreClick;
 BNouvRech.OnClick               := BNouvRechClick;
 BPopCompte.OnClick              := BPopCompteClick;
 BPopJournal.OnClick             := BPopJournalClick;

 EdtFFiltres.OnChange            := FFiltresChange;
 POPF.OnPopup                    := POPFPopup;
 POPS.OnPopup                    := POPSPopup;

 CmbContrepartie.OnChange        := CmbContrepartieChange;
 CmbJournal.OnChange             := CmbJournalChange;


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
 BMode.OnClick                   := BModeClick;
 BINTEGRER.OnClick               := BIntegrerClick;
 BSelectionner.OnClick           := BSelectionnerClick;
 BDeSelectionner.OnClick         := BDeSelectionnerClick;


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
procedure TOF_CRELBQE.HGSaisieInitLigne( value : integer );
begin
 HGSaisie.Cells[cColDate,value]             := ZReleveBanque.DtDATECOMPTABLE_R;
 HGSaisie.Cells[cColRefInterne,value]       := '';
 HGSaisie.Cells[cColDebit,value]            := '';
 HGSaisie.Cells[cColCredit,value]           := '';
 HGSaisie.Cells[cColLibelle,value]          := '';
 HGSaisie.Cells[cColPiece,value]            := '';
 HGSaisie.Cells[cColImputation,value]       := '' ;
 HGSaisie.Cells[cColTva,value]              := '';
 HGSaisie.CellValues[cColEtat,value]        := ZReleveBanque.StETAT_R;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Ajout d'une nouvelle ligne dans la grille et la TOB
Suite ........ : TOB_CRELBQE
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.Nouveau;
var
 lBoInutilise : boolean;
begin

 if ( StatutGrilleReleve <> taConsult )  then exit; // on ajout une ligne uniquement si on etait en consultation

 SetEvent(false); // déactive tous les evenements pour crée une nouvelle enregistrement

 // on vide la grille des imputations
 VideGrille(HGImputation);

 // on rajoute une ligne si : on est sur la première ligne et elle est valide ou le nombre de ligne est supérieur à 1
 if ( ( HGSaisie.Row = 1 ) and HGSaisieIsRowValid(HGSaisie.Row ) ) or ( HGSaisie.Row > 1 ) then
  HGSaisie.RowCount := HGSaisie.RowCount + 1 ;

 if HGSaisie.CanFocus then // on donne le focus à la grille
  HGSaisie.setfocus;
 HGSaisie.Row                            := HGSaisie.RowCount - 1;          // on se place sur cette nouvelle ligne
 HGSaisie.col                            := cColDate;
 HGSaisie.Objects[0,HGSaisie.Row]        := ZReleveBanque.AddLigneReleve;   // creation de l'enregistrement et ajout à la grille
 //ZReleveBanque.InCOMPTEUR_R              := ZReleveBanque.ProchainCompteur; // initialisation du compteur de ligne
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

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 08/11/2001
Modifié le ... :   /  /
Description .. : Insertion d'un nouvelle enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.BInsertClick(Sender: TObject);
begin
 Nouveau;
end;

procedure TOF_CRELBQE.BDeleteClick(Sender: TObject);
begin
 HGSaisieDeleteRow(HGSaisie.Row);
end;

procedure TOF_CRELBQE.BValiderClick(Sender: TObject);
begin

 if EnregistreLigne(HGSaisie.Row) then
  VideGrille(HGImputation);

end;

procedure TOF_CRELBQE.BImprimerClick(Sender: TObject);
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
procedure TOF_CRELBQE.BIntegrerClick(Sender: TObject);
var
 lZTresoVersEcr : TZTresoVersEcr;
 lTT            : TQRProgressForm ;
begin

 if ( FListeIntegrer.Count = 0 ) then
  exit;

 if ( PGIAskCancel( cStTexteConfirmIntegration , cStTexteTitreFenetre ) = mrNo) then
  begin
   // messaqe de confirmation de l'intégration
   FBoModeSelection := false;
   SetModeSelection;
   BDeSelectionnerClick(nil);
   exit;
  end; // if
 // creation de l'objet de transfert des ecritures ( definis dans //ZTresoVersEcr.pas )
 lZTresoVersEcr                       := TZTresoVersEcr.Create;
 // affichage de la fenetre d'attente
 lTT                                  := DebutProgressForm ( Ecran,
                                                             cStTexteTraitementEcr,
                                                             cStTexteTitreFenetre,
                                                             FListeIntegrer.Count - 1,
                                                             true,
                                                             true ) ;
 lZTresoVersEcr.TT                    := lTT;   // on passe le pointeur d'attente a l'objet de transferts
 lZTresoVersEcr.StJournalContrepartie := ZReleveBanque.StJournalContrepartie;
 lZTresoVersEcr.StModeSaisie          := ZReleveBanque.StModeSaisie;
 lZTresoVersEcr.StCompteurNormal      := ZReleveBanque.StCompteurNormal;

 try
  // on charge l'ensemble des lignes à transferer
  if lZTresoVersEcr.LoadReleve( FListeIntegrer ) then
   begin
    // on enregistre les ecritures
    case lZTresoVersEcr.Save of
     oeOk,
     oeSaisie,
     oeStock,
     oeReseau : begin  // erreur sur la mise à jour des comptes ->
                 FBoModeSelection := false;
                 SetModeSelection;
                 if BCherche.Enabled then
                  BChercheClick(nil);
                end; // if
    end; // case
   end; // if

  BDeSelectionnerClick(nil);

 finally
  if assigned(lZTresoVersEcr) then lZTresoVersEcr.Free;
  if assigned(lTT)            then lTT.Free;
 end; // try

end;

procedure TOF_CRELBQE.BSelectionnerClick(Sender: TObject);
var
 i   : integer;
begin

 for i := 0 to ( ZReleveBanque.TOBReleve.Detail.Count - 1 ) do
  begin
   ZReleveBanque.TOBLigneReleve := ZReleveBanque.TOBReleve.Detail[i];
   SelectionnerLigne;
  end; // for

 HGSaisie.AllSelected := true;
 // on remet à jour l'affichage des boutons des fenetres
 EnabledControl;

end;

procedure TOF_CRELBQE.BDeSelectionnerClick(Sender: TObject);
begin
 FListeIntegrer.Clear;
 HGSaisie.ClearSelected;
 EnabledControl;
end;

function TOF_CRELBQE.SelectionnerLigne : boolean;
var
 lInIndex : integer;
begin

 result   := true;
 lInIndex := FListeIntegrer.IndexOf( varToStr(ZReleveBanque.InCOMPTEUR_R) );

 if ( lInIndex = -1 ) then
  begin // l'enregistrement n' a jamais été selectionné
   if ( ( ZReleveBanque.StETAT_R = cImputCorrectNonValide ) or ( ZReleveBanque.StETAT_R = cImputValide ) ) then
    FListeIntegrer.Add( varToStr(ZReleveBanque.InCOMPTEUR_R) ) // on l'ajoute à la liste
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
procedure TOF_CRELBQE.BDefaireClick(Sender: TObject);
begin
 ReaffecterLigneReleve;
 FRdNumeroPieceMax  := FRdNumeroPieceMaxOld;
 StatutGrilleReleve := taConsult;
end;


procedure TOF_CRELBQE.BAgrandirClick(Sender: TObject);
begin
// if PgcCritere.Visible then
 PgcCritere.Visible := not PgcCritere.Visible;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Gestion du déplacement dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGSaisieCellEnter(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean);
var
 lInCurrentCol          : integer;
 lInCurrentRow          : Integer;
begin

 if csDestroying in Ecran.ComponentState then Exit ;
 // on stocke la position courante car cette eve est redencher plusieur fois
 lInCurrentCol := HGSaisie.Col;
 lInCurrentRow := HGSaisie.Row;

 // gestion de l'affichage des boutons
 HGSaisie.ElipsisButton := ( lInCurrentCol = cColImputation ) or ( lInCurrentCol = cColDate);
 // gestion des deplacments dans la grille
 case lInCurrentCol of
  cColDebit    : if ( HGSaisie.Cells[cColCredit,lInCurrentRow] <> '' ) then
                 begin
                  PasseColSuivante(ACol,ARow,HGSaisie);
                  Cancel := true;
                 end;
  cColCredit   : if ( HGSaisie.Cells[cColDebit,lInCurrentRow] <> '' ) then
                 begin
                  PasseColSuivante(ACol,ARow,HGSaisie);
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
                 end;
  cColTva     : if not ZReleveBanque.IsTvaAutorise(HGSaisie.Cells[cColImputation,lInCurrentRow]) then
                 begin
                   PasseColSuivante(ACol,ARow,HGSaisie);
                   SetGridGrise(cColTva,lInCurrentCol,HGSaisie);
                   Cancel := true;

                 end;
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
procedure TOF_CRELBQE.HGSaisieElipsisClick(Sender: TObject);
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

procedure TOF_CRELBQE.AfficheLookUpImput;
var
 lStGeneral : string;
 lBoResult  : boolean;
begin

 lBoResult    := false;

 lStGeneral   := HGSaisie.Cells[cColImputation,HGSaisie.Row];

 if ZReleveBanque.ExisteImputation(lStGeneral) then
  begin
   ZReleveBanque.StGENERAL_R := lStGeneral;
   ZReleveBanque.AssignInfoCompte;
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
       ZReleveBanque.StGENERAL_R := HGSaisie.Cells[cColImputation,HGSaisie.Row];
       ZReleveBanque.AssignInfoCompte;
       lBoResult                 := false;
      end; // if
    end; // if

  if lBoResult then
   begin

    // remise à jour des infos
    HGSaisie.Cells[cColImputation,HGSaisie.Row]  := ZReleveBanque.StGENERAL_R;
    if   ( ZReleveBanque.StTVASAISIE_R = 'X' ) then
    HGSaisie.Cells[cColTva,HGSaisie.Row]         := STRFMONTANT ( ZReleveBanque.RdTAUXTVA_R , 15 , V_PGI.OkDecV, '' , true)
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
function  TOF_CRELBQE.HGSaisieRechercheImputation : boolean;
var
 lStGeneral : string;
begin

 lStGeneral  := HGSaisie.Cells[cColImputation,HGSaisie.Row];

 result      := ZReleveBanque.IsValideGeneral(lStGeneral); 

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Vérification de la validité des comptes saisies
Suite ........ : Formatage des cellules
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGSaisieCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
 RdMontant : double;
begin

 if ( csDestroying in Ecran.ComponentState ) or ( StatutGrilleReleve = taConsult )then Exit ;

 case ACol of
  cColDate       : begin
                    Cancel := not ( IsValidDate(HGSaisie.Cells[cColDate,ARow]) and
                                  ( strToDate(HGSaisie.Cells[cColDate,ARow]) >= FDtDateDebN ) and
                                  ( StrToDate(HGSaisie.Cells[cColDate,ARow]) <= FDtDateFinN1 ));
                    // nul a changer
                    if Cancel and ( strToDate(HGSaisie.Cells[cColDate,ARow]) < FDtDateDebN ) then
                     PGIInfo( cStDateInfDateDebExo,cStTexteTitreFenetre )
                      else
                       if Cancel and ( strToDate(HGSaisie.Cells[cColDate,ARow]) > FDtDateFinN1 ) then
                        PGIInfo( cStDateInfDateFinExo,cStTexteTitreFenetre );
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
procedure TOF_CRELBQE.HGSaisiePostDrawCell( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState );
var
 lBoGrise  : boolean;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 try

 // gestion des cases debit et crédit - une seul des deux doit être renseignée
 lBoGrise := ( ( ACol = cColDebit )  and ( HGSaisie.Cells[cColCredit, ARow] <>'' ) and  ( ARow > 0 ) ) or
             ( ( ACol = cColCredit ) and ( HGSaisie.Cells[cColDebit, ARow]  <>'' ) and  ( ARow > 0 ) ) OR
             ( ( ACol = cColTva )    and  ( ARow > 0 ) and not ZReleveBanque.IsTvaAutorise(HGSaisie.Cells[cColImputation,ARow]));

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
     PGIBox( 'Erreur sur lors du dessin de la grille' + #13#10 +
             E.Message ,
             cStTexteTitreFenetre);
     V_PGI.IoError := oeSaisie;
   end; // on

 end;

end ;

procedure TOF_CRELBQE.HGSaisieClick(Sender: TObject);
begin
   if not FBoChoixEffectue then
    begin
     PGIInfo( Ecran.Caption,cStTexteTitreFenetre );

    end; // if
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Gestion des touches dans les deux grilles
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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
               Nouveau
                else
                 if not FBoGuideRun then
                  NouvelleImputation;
  VK_DOWN   : if lBoHGSaisieActive then
               begin // si la grille des releve est active
                if not EnregistreLigne(HGSaisie.Row) then Key := 0;

                if ( HGSaisie.Row = ( HGSaisie.RowCount - 1 ) )  then
                 Nouveau; // on insere un ligne que sur la dernière cellule
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

                  Nouveau;
                 end;  // if

                if ( ZReleveBanque.StETAT_R <> cPasImput ) and not EnregistreLigne(HGSaisie.Row) then
                 Key := 0;
               end;
  VK_UP     : begin
               // on passe en mode auto pour ne afficher le message
               ZReleveBanque.TypeContexte := TModeAuto;
               if lBoHGSaisieActive and( StatutGrilleReleve = taCreat ) and not ( HGSaisieIsRowValid(HGSaisie.Row) ) then
                begin
                 HGSaisieDeleteRow(HGSaisie.Row);
                 Key := 0;
                end
                 else
                  if lBoHGSaisieActive and not EnregistreLigne(HGSaisie.Row) then
                   Key := 0;
                ZReleveBanque.TypeContexte := TModeSaisie;
              end; // if
  VK_ESCAPE : if lBoHGImputationActive then
               begin
                key := 0;
                HGImputationExit(nil);
               end
                else
                 if lBoHGSaisieActive then
                  begin
                   if ( FStatutGrilleReleve <> taConsult ) and (PGIAskCancel(cStTexteEchap,cStTexteTitreFenetre) = mrNo) then
                    begin
                     ReaffecterLigneReleve;
                     Key := 0;
                    end; // if
                  end; // if
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
                if not EnregistreLigne(HGSaisie.Row) then exit;
                if not AltDImputation                then HGSaisie.SetFocus;
               end;
  Ord('I')  : if Shift = [ssAlt]  then
               begin
                Key := 0;
                {$IFNDEF EAGLCLIENT}
                {$IFNDEF SANSCOMPTA}
                FicheDevise( ZReleveBanque.Devise.Code, taConsult, FALSE) ;
                {$ENDIF}
                {$ENDIF}
               end;
  Ord('A')  : if Shift = [ssCtrl] then
               begin
                Key := 0;
                if lBoHGSaisieActive and FBoModeSelection then
                 BSelectionnerClick(nil)
                { else
                  begin
                   HGImputationAssignLigneTOB(HGImputation.Row);
                   if ( ZReleveBanque.IsVentilable ) and ( VH^.ZGereAnal )  then
                    if FBoGuideRun then
                     ZreleveBanque.AppelAnalytique (TModeGuideAna)
                     else
                      ZreleveBanque.AppelAnalytique (TModeVisuAna);
                  end; // if  }
               end;  // Ord(A)
  VK_F6     : if lBoHGSaisieActive then
               begin
                 Key := 0;
                 if ( StatutGrilleReleve = taConsult ) then // sauvegarde de la ligne courante
                  FTOBSauveLigne.Dupliquer(ZReleveBanque.TOBLigneReleve,true,true);
                 if not EnregistreLigne(HGSaisie.Row) then exit;
                 AfficheGuide;
               end // if
                else
                 if lBoHGImputationActive and ( HGImputation.Row <> ( HGImputation.RowCount - 1 )) then
                  VK_F6Imputation; // calcul du solde de la ligne mais pas sur la derniere ligne ( ligne de contrepartie non modifiable
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

              end; // VK_F7
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
procedure TOF_CRELBQE.VK_F7Grille( G : THGrid );
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
procedure TOF_CRELBQE.ShiftVK_F7Grille( G : THGrid ) ;
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

 lStNum               := IntToStr( ValeurI(lStNum) + 1) ;
 lStRef               := Copy( lStRef, 1, Length(lStRef) - l) + lStNum ;
 G.Cells[G.Col,G.Row] := lStRef ;

 if isNumeric(lStNum) and ( StrToInt(lStNum) > FRdNumeroPieceMax ) then
  FRdNumeroPieceMax  := StrToInt(lStNum);

end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... : 27/07/2001
Description .. : On recherche l'ensemble des comptes de contrepartie des
Suite ........ : journaux de type banques et on calcule le solde de la
Suite ........ : periode precedente du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.RemplirComboBoxContrepartie;
var
 Q      : TQuery;
begin

 Q := nil;

 // gestion du compte bancaire
 try
   // on recherche l'ensemble des comptes de contrepartie des journal de type banques
   Q := OpenSQL(' select distinct J_CONTREPARTIE '  +
                'from JOURNAL, BANQUECP '           +
                'where J_NATUREJAL ="BQE" '         +
                'and BQ_GENERAL=J_CONTREPARTIE '    +
                'AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ' + // 24/10/2006 YMO Multisociétés
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
Modifié le ... : 21/08/2001
Description .. : Gestion de la combobox du choix du journal de banque
Suite ........ : associé au compte de contrepartie
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.CmbContrepartieChange(Sender: TObject);
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
 ZReleveBanque.Initialize;

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

 ZReleveBanque.StCompteContrepartie   := CmbContrepartie.Text;

 HlbDevise.Caption := TraduireMemoire('Dossier tenu en ') + VH^.LibDevisePivot + '   ' +
                      TraduireMemoire('Compte tenu en ')  + ZReleveBanque.Devise.Libelle;

 ImgEuro.Visible   := VH^.TenueEuro;
 // selectionne de l'ensemble des comptes de banques du comptes de contrepartie choisi

 Q := nil;

 try

  Q := OpenSQL('SELECT j_journal,j_libelle,j_contrepartie from journal ' +
                ' where J_NATUREJAL ="BQE" ' +
                ' and J_CONTREPARTIE = ' + CmbContrepartie.Text ,
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
procedure TOF_CRELBQE.CmbJournalChange(Sender: TObject);
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

 ZReleveBanque.Initialize;
 ZReleveBanque.StJournalContrepartie   := CmbJournal.Value;
 // Initialisation de l'objet guide
 ZGuideTreso.StJournalContrepartie     := ZReleveBanque.StJournalContrepartie;
 ZGuideTreso.ZCompte                   := ZReleveBanque.ZComptes;
 ZGuideTreso.StLeNomTable              := 'CRELBQE';
 ZGuideTreso.StEtablissement           := CmbEtablissement.Value;
 ZGuideTreso.StDevise                  := ZReleveBanque.Devise.Code;
 ZGuideTreso.StCompteContrepartie      := ZReleveBanque.StCompteContrepartie;
  // chargement de l'ensemble des guides
 ZGuideTreso.Load;
 // On recherche si l'on a des RIB a integrer
 RechercheRibAIntegrer;

 // On lance tous de suite la recherche
 BChercheClick(nil);
 CalculSoldeTheorique;
 HGSaisieAfficheSolde;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Affichage des soldes des ecritures + lignes de releves
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGSaisieAfficheSolde;
begin
 HlbDebitSaisie.caption  := STRFMONTANT ( ZReleveBanque.RdSoldeDebit , 15 , V_PGI.OkDecV, '' , true);
 HlbCreditSaisie.caption := STRFMONTANT ( ZReleveBanque.RdSoldeCredit , 15 , V_PGI.OkDecV, '' , true);

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
procedure TOF_CRELBQE.HGSaisieRowExit(Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin

  if csDestroying in Ecran.ComponentState then Exit ;

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
function TOF_CRELBQE.HGSaisieIsRowValid( Row: Integer ) : Boolean;
begin

 result := HGSaisieAssignTOB(Row);
 if result then
  result := ZReleveBanque.IsValideLigneReleve;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/06/2001
Modifié le ... : 20/07/2001
Description .. : Récupération des données de la grille des releves dans la
Suite ........ : TOBReleve
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.HGSaisieAssignTOB( ARow : integer) : boolean;
begin

 result := false;

 if HGSaisie.Objects[0,ARow] = nil then exit;

 try

  // récupération de l'objet dans la TOB
  ZReleveBanque.TOBLigneReleve       := TOB(HGSaisie.Objects[0,ARow]);
  ZReleveBanque.DtDATECOMPTABLE_R    := HGSaisie.Cells[cColDate,ARow];
  ZReleveBanque.StREFINTERNE_R       := HGSaisie.Cells[cColRefInterne,ARow];
  ZReleveBanque.STREFEXTERNE_R       := HGSaisie.Cells[cColPiece,ARow];
  ZReleveBanque.StGENERAL_R          := HGSaisie.Cells[cColImputation,ARow];
  ZReleveBanque.StETABLISSEMENT_R    := CmbEtablissement.Value;

  if isNumeric(HGSaisie.Cells[cColTva,ARow]) then
   begin
    ZReleveBanque.RdTAUXTVA_R        := HGSaisie.Cells[cColTva,ARow];
    ZReleveBanque.StTVASAISIE_R      := 'X';
   end
    else
     begin
      ZReleveBanque.RdTAUXTVA_R      := 0;
      ZReleveBanque.StTVASAISIE_R    := '-';
     end;

  CSetMontants( ZReleveBanque.TOBLigneReleve,
                Valeur(HGSaisie.Cells[cColDebit,ARow]) ,
                Valeur(HGSaisie.Cells[cColCredit,ARow]) ,
                ZReleveBanque.Devise,
                false,
                true );

 if ZReleveBanque.RdDEBIT_R = 0  then
  ZReleveBanque.StNATUREPIECE_R := 'RC'
   else
    ZReleveBanque.StNATUREPIECE_R := 'RF';

  ZReleveBanque.StLIBELLE_R       := HGSaisie.Cells[cColLibelle,ARow];

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
Modifié le ... : 20/07/2001
Description .. : Enregistrer une nouvelle ligne
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.EnregistreLigne( ARow : integer ) : boolean;
var
 lRdSolde : double;
begin

 result := true;

 if ( StatutGrilleReleve = taConsult ) then exit; // on consultation on sort

 result := false;
 // sauvegarde de la ligne courante
 FTOBSauveLigne.Dupliquer(ZReleveBanque.TOBLigneReleve,true,true);


 if HGSaisie.Col = cColImputation then AfficheLookUpImput;

 if ( ZReleveBanque.StETAT_R <> cPasImput ) then
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

 result := HGSaisieIsRowValid(ARow); // validation de la ligne

 if result then
  begin
   FTOBSauveLigne.Dupliquer(ZReleveBanque.TOBLigneReleve,true,true);
   result := ZReleveBanque.SaveLigneReleve; // enregistrement du releve
   if not result then exit;
   HGSaisieRefreshLigne(ARow);    //  rafraisir la grille
   // si on a effacé le numero de piece max on repart de l'ancien
   if ZReleveBanque.StREFINTERNE_R = '' then
    FRdNumeroPieceMax := FRdNumeroPieceMaxOld
     else
      FRDNumeroPieceMaxOld := FRdNumeroPieceMax;
   StatutGrilleReleve := taConsult;

   if ( StrToDate(ZReleveBanque.DtDATECOMPTABLE_R) >= StrToDate(EdtMulDateDu.Text) ) and
      ( StrToDate(ZReleveBanque.DtDATECOMPTABLE_R) <= StrToDate(EdtMulDateAu.Text) ) then
    begin
     lRdSolde  := FRdSolDeT + ZReleveBanque.RdDEBIT_R - ZReleveBanque.RdCREDIT_R;
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
procedure TOF_CRELBQE.HGSaisieRefreshLigne( ARow : integer );
begin

 try

  HGSaisie.Cells[0,ARow]               := intToStr(HGSaisie.Row);
  HGSaisie.Cells[cColDate,ARow]        := ZReleveBanque.DtDATECOMPTABLE_R;
  HGSaisie.Cells[cColRefInterne,ARow]  := ZReleveBanque.StREFINTERNE_R;
  HGSaisie.Cells[cColPiece,ARow]       := ZReleveBanque.StREFEXTERNE_R;
  HGSaisie.Cells[cColImputation,ARow]  := ZReleveBanque.StGENERAL_R;

  if   ( ZReleveBanque.StTVASAISIE_R = 'X' ) then
  HGSaisie.Cells[cColTva,ARow]         := STRFMONTANT ( ZReleveBanque.RdTAUXTVA_R , 15 , V_PGI.OkDecV, '' , true)
    else
     HGSaisie.Cells[cColTva,ARow]      := '';


  if ZReleveBanque.RdDEBIT_R <> 0 then // si debit = 0 la cellule doit afficher ''
   HGSaisie.Cells[cColDebit,ARow]      := STRFMONTANT ( ZReleveBanque.RdDEBIT_R , 15 , V_PGI.OkDecV, '' , true);
  if ZReleveBanque.RdCREDIT_R <> 0 then
   HGSaisie.Cells[cColCredit,ARow]     := STRFMONTANT ( ZReleveBanque.RdCREDIT_R , 15 , V_PGI.OkDecV, '' , true);
  HGSaisie.Cells[cColLibelle,ARow]     := ZReleveBanque.StLIBELLE_R;
  HGSaisie.CellValues[cColEtat,ARow]   := ZReleveBanque.StETAT_R ;

  // gestion du plus grand numero de pièce
  if ( trim(ZReleveBanque.StREFINTERNE_R) <> '' )               and
       isNumeric(ZReleveBanque.StREFINTERNE_R )           and
     ( ZReleveBanque.StREFINTERNE_R > FRdNumeroPieceMax ) then
   FRdNumeroPieceMax := (ZReleveBanque.StREFINTERNE_R);

 except
  on E : Exception do
   begin
    V_PGI.IoError := oeSaisie;
    PGIInfo('Erreur lors de la récupération des données en base.'               + #13#10 +
            'Ligne de relevée numéro ' + varToStr(ZReleveBanque.InCOMPTEUR_R)   + #13#10 +
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
procedure TOF_CRELBQE.SetEvent(value: boolean);
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
Modifié le ... : 20/07/2001
Description .. : Suppression d'une ligne
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGSaisieDeleteRow(ARow: integer);
var
 lRdSolde : double;
begin

 SetEvent(false);

 lRdSolde           := FRdSolDeT - ZReleveBanque.RdDEBIT_R + ZReleveBanque.RdCREDIT_R;
 FRdSolDeT          := lRdSolde;
 StatutGrilleReleve := taModif;
 if HGSaisieAssignTOB(ARow) and ZReleveBanque.DeleteLigneReleve then
  begin
   HGSaisie.Objects[0,ARow] := nil;
   if HGSaisie.RowCount = 2 then
    begin
     StatutGrilleReleve  := taConsult;
     Nouveau;
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
Modifié le ... : 20/07/2001
Description .. : Gestion du filtre de recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.BChercheClick(Sender: TObject);
var
 lStSQL  : string;
 Q       : TQuery;
begin

 // on se replace dans le libelle
 if EdtMulLibelle.CanFocus then
  EdtMulLibelle.SetFocus;

 lStSQL := ' select '                      +
           'CRL_COMPTEUR, '                +
           'CRL_JOURNAL, '                 +
           'CRL_DATECOMPTABLE, '           +
           'CRL_REFINTERNE, '              +
           'CRL_REFEXTERNE, '              +
           'CRL_ETAT, '                    +
           'CRL_GENERAL, '                 +
           'CRL_TAUXTVA, '                 +
           'CRL_TVASAISIE, '               +
           'CRL_DEBIT, '                   +
           'CRL_CREDIT, '                  +
           'CRL_LIBELLE, '                 +
           'CRL_COMPTEURPERE, '            +
           'CRL_TYPE, '                    +
           'CRL_IMPORT, '                  +
           'CRL_DATECREATION, '            +
           'CRL_DEVISE, '                  +
           'CRL_NATUREPIECE, '             +
           'CRL_GENERALBQE '               +
           ' from CRELBQE ';

 lStSQL := lStSQL + 'where ( CRL_DATECOMPTABLE >= "' + UsDate(EdtMulDateDu) + '" and CRL_DATECOMPTABLE <= "' + UsDate(EdtMulDateAu) + '") ' ;

 if EdtMulLibelle.Text <> '' then
 lStSQL := lStSQL + ' and CRL_LIBELLE like "' + EdtMulLibelle.Text + '%"';

 if EdtMulNumeroPiece.Text <> '' then
  lStSQL := lStSQL + ' and CRL_REFINTERNE like "' + EdtMulNumeroPiece.Text + '%"';

 if EdtMulNumeroPieceCpml.Text <> '' then
  lStSQL := lStSQL + ' and CRL_REFEXTERNE like "' + EdtMulNumeroPieceCpml.Text + '%"';

 if EdtMulGeneral.Text <> '' then
  lStSQL := lStSQL + ' and CRL_GENERAL like "' + EdtMulGeneral.Text + '%"';

 if ZReleveBanque.StJournalContrepartie <> '' then
  lStSQL := lStSQL + ' and CRL_JOURNAL = "' + ZReleveBanque.StJournalContrepartie + '"';

 if ZReleveBanque.StCompteContrepartie <> '' then
  lStSQL := lStSQL + ' and CRL_GENERALBQE = "' + ZReleveBanque.StCompteContrepartie + '"';


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
   ZReleveBanque.LoadReleve(Q);      // chargement des relevées

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
procedure TOF_CRELBQE.HGSaisieRemplirGrille;
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

 if ZReleveBanque.TOBReleve.Detail.Count = 0 then
   Nouveau  //HGSaisie.RowCount := 2
   else
    begin

     HGSaisie.RowCount := ZReleveBanque.TOBReleve.Detail.Count + 1;
     HGSaisie.Row      := 1;
     HGSaisie.Col      := cColDate;

     for i := 0 to ZReleveBanque.TOBReleve.Detail.Count - 1 do
      begin
       HGSaisie.Row                       := i + 1;
       HGSaisie.Cells[0,i+1]              := intToStr(i+1);
       // faire une fct NextTOB(index);
       ZReleveBanque.TOBLigneReleve       := ZReleveBanque.TOBReleve.Detail[i];

       HGSaisieRefreshLigne(HGSaisie.Row);
       HGSaisie.Objects[0,HGSaisie.Row]   := ZReleveBanque.TOBLigneReleve; // creation de l'enregistrement et ajout à la grille
      end; // for

    FBoGrilleVide := false;

    end; // if Grille est vide


 // se replacer au début
 if HGSaisie.CanFocus then // on donne le focus à la grille
  HGSaisie.setfocus;
 HGSaisie.Row          := 1;
 HGSaisie.Col          := cColDate;

 // la grille repasse en consultation
 StatutGrilleReleve    := taConsult;
 // rafraichir la grille
 RafraichirLigne;
 HGImputation.Enabled  := false;

 FBoModeGuideAuto      := false;

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
procedure TOF_CRELBQE.SetStatutGrilleReleve ( Value : TActionFiche );
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
procedure TOF_CRELBQE.HGSaisieKeyPress(Sender: TObject; var Key: Char);
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
procedure TOF_CRELBQE.HGSaisieKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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
Modifié le ... : 20/07/2001
Description .. : Grise les cellules débit/crédit en fonction de leur état
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGImputationPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
 lBoGrise  : boolean;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 // gestion des cases debit et crédit - une seul des deux doit être renseignée
 lBoGrise := ( ( ACol = cColDebitImput )  and ( HGImputation.Cells[cColCreditImput, ARow] <>'' ) and  ( ARow > 0 ) ) or
             ( ( ACol = cColCreditImput ) and ( HGImputation.Cells[cColDebitImput, ARow]  <>'' ) and  ( ARow > 0 ) ) or
             (
              ( ACol = cColAuxiliaireImput )                                                            and
              ( not ZReleveBanque.IsCollectif( HGImputation.Cells[cColGeneralImput,HGImputation.Row]) ) and
              ( ARow > 0 )
              );

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
Modifié le ... : 20/07/2001
Description .. : 2 facon de se déplacer dans la grille :
Suite ........ : - Guide de saisie
Suite ........ : - Standart
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGImputationCellEnter(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean);
var
 lInCurrentCol          : integer;
 lInCurrentRow          : Integer;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

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
     // on bouge deans la grille pour enlever le le ligne selectionné
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
  cColAuxiliaireImput : if not ZReleveBanque.IsCollectif( HGImputation.Cells[cColGeneralImput,lInCurrentRow]) then
                        begin
                         HGImputation.Cells[cColAuxiliaireImput,lInCurrentRow] := '';
                         PasseColSuivante(ACol,ARow,HGImputation);
                         Cancel := true;
                        end; // if
  cColDebitImput     : if ( HGImputation.Cells[cColCreditImput, lInCurrentRow] <> '' ) then
                       begin
                        PasseColSuivante(ACol,ARow,HGImputation);
                        Cancel := true;
                       end; // if
  cColCreditImput    : if ( HGImputation.Cells[cColDebitImput, lInCurrentRow] <> '' ) then
                       begin
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
function TOF_CRELBQE.NextColStop( ARow, ACol : integer ) : TPoint;
var
 i,j  : integer;
 lCol : integer;
begin

 lCol := ACol + 1;

 for i := ARow to HGImputation.RowCount - 1 do
  begin
   for j := lCol to HGImputation.ColCount - 1 do
     if ZGuideTreso.IsColStop(i,j) then
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


procedure TOF_CRELBQE.HGImputationRowEnter( Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin
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
procedure TOF_CRELBQE.HGImputationCellExit(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean);
var
 lCell          : TPoint;
 lBoFirstCell   : boolean;
 RdMontant      : double;
begin

 // on stocke la cellule que l'on est en train de tester, et l'on sort de ce traitement si l'on n'a pas bougé
 // rem : cet evenement est appelé plusieur fois de suite quand on force le changement de cellule
 if ( csDestroying in Ecran.ComponentState ) or
    ( ( FCurrentCol = ACol ) and ( FCurrentRow = ARow ) ) then Exit ;

 // recuperation des données de la grille vers la TOB

 if not HGImputationAssignLigneTOB(HGImputation.Row) then exit;

 FCurrentCol      := ACol;
 FCurrentRow      := ARow;


 case ACol of
  cColGeneralImput : if not HGImputationRechercheImputation(ARow) then
                     begin
                       // on reinitialise ces valeurs à zero pour redeclencher l'evenement
                       FCurrentCol             := 0;
                       FCurrentRow             := 0;
                       Cancel                  := true ;
                     //  HGImputationElipsisClick(nil);
                       PostMessage(HGImputation.Handle, WM_KEYDOWN, VK_F5, 0); // on simule un VK_F5 pour afficher la fenetre des comptes
                       exit;
                     end // if
                      else
                       if ( VH^.ZGereAnal ) and ( ZReleveBanque.IsVentilable ) and ( ZReleveBanque.stANA_I <> 'X' ) then
                         ZreleveBanque.AppelAnalytique (TModeManuAna ); // on affiche le fentre d'analytique
  cColAuxiliaireImput : if not HGImputationRechercheTiers then
                        begin
                         Cancel:= true ;
                         PostMessage(HGImputation.Handle, WM_KEYDOWN, VK_F5, 0); // on simule un VK_F5 pour afficher la fenetre des comptes
                         exit;
                       end; // if
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

   lCell            := NextColStop( ARow, ACol);
   FCurrentCol      := ACol;
   FCurrentRow      := ARow;
   ACol             := lCell.x;
   ARow             := lCell.y;
   Cancel           := true;

   lBoFirstCell     := ( ARow = 1 ) and ( ACol = 1 );

   if lBoFirstCell  then
    begin
     FCurrentCol             := 0;
     FCurrentRow             := 0;
     SetModeGuide(false);
     // on supprime les lignes vides
     ZReleveBanque.SupprimerLigneImputVide;

     // on remet a jour la grille si jamais des lignes en été supprimer
     HGImputationRemplirGrille;
     //HGImputation.Refresh;
     
    end
     else
      begin
              HGImputationAssignLigneTOB(FCurrentRow);
       // ZReleveBanque.RecupereGuide(ZGuideTreso.RecalculGuide ( ZReleveBanque.AffecteGuide(FCurrentRow) , FCurrentRow ));
        ZGuideTreso.RecalculGuide ( ZReleveBanque.TOBImput , FCurrentRow  );
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
procedure TOF_CRELBQE.SetModeGuide( Value : boolean);
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
procedure TOF_CRELBQE.NouvelleImputation;
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
Modifié le ... : 20/07/2001
Description .. : Recherche d'un guide de saisie en fonction des info
Suite ........ : présente dans la ligne
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.AfficheGuide : boolean;
var
 lCellStop  : TPoint;
 lTOB       : TOB;
begin

 result := false;

 // suppression des lignes d'imputation courantes
 ZReleveBanque.DeleteCurrentImputation;

 //if ZReleveBanque.StCompteTVA = '' then
 ZReleveBanque.AssignInfoCompte;

 ZGuideTreso.StCompteTVA := ZReleveBanque.StCompteTVA;

 lTOB:=TOB.Create('',nil,-1);

 result:=ZGuideTreso.RechercheGuideEnBase(  HGSaisie.Cells[cColCredit      , HGSaisie.Row],
                                            HGSaisie.Cells[cColDebit       , HGSaisie.Row],
                                            HGSaisie.Cells[cColLibelle     , HGSaisie.Row],
                                            HGSaisie.Cells[cColImputation  , HGSaisie.Row],
                                            HGSaisie.Cells[cColTVA         , HGSaisie.Row],
                                            HGSaisie.Cells[cColDate        , HGSaisie.Row],
                                            HGSaisie.Cells[cColRefInterne  , HGSaisie.Row],
                                            HGSaisie.Cells[cColPiece       , HGSaisie.Row],
                                            '', // il n'y a pas de CodeAFB en saisie manuel
                                            lTOB ,
                                            HGSaisie );  // on le repasse la grille pour l'affichage de lookup de choix des guides
                                                          // en affectation automatique on passe la valeur nil, par d'interaction avec l'utilisateur

 if result then
  begin // un guide existe
   // on affecte les enregistrements contenu dans lTOB à la ligne de releve courante
   ZReleveBanque.CreerImputation(lTOB);

   if assigned(lTOB) then lTOB.Free;

   HGImputationRemplirGrille;
   lCellStop              := NextColStop(1,0);
   HGImputation.Col       := lCellStop.x;
   HGImputation.Row       := lCellStop.y;

   // on affecte le flag indiquant ce mode de saisie ( utilisée dans le CellExit et le CellEnter )
   // si le guide n'a pas de point d'arret en ne passe pas en mode guide
   SetModeGuide(ZGuideTreso.PossedeArret);

   // on passe en modification dans la grille
   StatutGrilleImput := taCreat;
   // on complète le ligne de relevé avec les info du guide
   SynchroniseHGSaisie;

   result := true;

  end; //if

 AfficheSoldeImputation;

 CmbContrepartie.Enabled        := false;
 BIntegrer.Enabled              := false;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Remplit la grille des imputations avec les valeurs de la base
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.AltDImputation : boolean;
begin

 result := false;
 
 if ZReleveBanque.StETAT_R = cPasImput  then exit;

 // sauvegarde de la ligne courante
 FTOBSauveLigne.Dupliquer(ZReleveBanque.TOBLigneReleve,true,true);

 if ZReleveBanque.RechercheImputation then
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

function TOF_CRELBQE.HGImputationRemplirGrille : boolean;
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
   ZReleveBanque.FirstImputation;

   // affichage du nom du guide
   HlbNomGuide.Visible            := ZReleveBanque.StGUIDE_I <> '';
   HlbNomGuide.Caption            :=  TraduireMemoire('Guide n°') + varToStr(ZReleveBanque.StGUIDE_I)
                                                       + ' ' + RechDom('TTGUIDEECR', ZReleveBanque.StGUIDE_I ,false);

   while assigned( ZReleveBanque.TOBLigneImputation ) do
    begin

     if ( ZReleveBanque.StGENERAL_I <> VH^.EccEuroDebit )
     and ( ZReleveBanque.StGENERAL_I <> VH^.EccEuroCredit ) then
      begin

       HGImputation.RowCount := lInRowCount;
       HGImputation.Row      := lInRowCount - 1; // on se place sur la ligne courante

       HGImputation.Cells[0,HGImputation.Row]                   := intToStr(HGImputation.Row);
       HGImputation.Cells[cColGeneralImput,HGImputation.Row]    := ZReleveBanque.StGENERAL_I;
       HGImputation.Cells[cColReferenceImput,HGImputation.Row]  := ZReleveBanque.StREFINTERNE_I;
       HGImputation.Cells[cColLibelleImput,HGImputation.Row]    := ZReleveBanque.StLIBELLE_I;
       if ZReleveBanque.RdDEBIT_I <> 0 then // si debit = 0 la cellule doit afficher ''
        HGImputation.Cells[cColDebitImput,HGImputation.Row]     := STRFMONTANT ( ZReleveBanque.RdDEBIT_I , 15 , V_PGI.OkDecV, '' , true);
       if ZReleveBanque.RdCREDIT_I <> 0 then
        HGImputation.Cells[cColCreditImput,HGImputation.Row]    := STRFMONTANT ( ZReleveBanque.RdCREDIT_I , 15 , V_PGI.OkDecV, '' , true);

       HGImputation.Objects[cColObjectReleve,HGImputation.Row]  := ZReleveBanque.TOBLigneImputation; // creation de l'enregistrement et ajout à la grille

      end;

     Inc(lInRowCount);
     ZReleveBanque.NextImputation;

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

procedure TOF_CRELBQE.SynchroniseHGSaisie;
begin
 ZReleveBanque.StGENERAL_R := HGImputation.Cells[1,1];
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Affiche le solde des imputations
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.AfficheSoldeImputation;
begin
 ZReleveBanque.CalculSoldeImput;
 HlbDebitImput.Caption          := STRFMONTANT ( ZReleveBanque.RdSoldeDebitImput , 15 , V_PGI.OkDecV, '' , true)  ;
 HlbCreditImput.Caption         := STRFMONTANT ( ZReleveBanque.RdSoldeCreditImput , 15 , V_PGI.OkDecV, '' , true) ;
 if ZReleveBanque.RdSoldeImput > 0 then
  HlbSoldeImput.Caption         := STRFMONTANT ( ZReleveBanque.RdSoldeImput , 15 , V_PGI.OkDecV, '' , true)  + ' D'
   else
    if ZReleveBanque.RdSoldeImput < 0 then
     HlbSoldeImput.Caption      := STRFMONTANT ( ZReleveBanque.RdSoldeImput * (-1) , 15 , V_PGI.OkDecV, '' , true) + ' C'
      else
       HlbSoldeImput.Caption      := STRFMONTANT ( ZReleveBanque.RdSoldeImput , 15 , V_PGI.OkDecV, '' , true);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Gestion des click sur des cellules imputation
Suite ........ : et tva
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGImputationElipsisClick(Sender: TObject);
var
 lStCompte : string;
begin

 if ( csDestroying in Ecran.ComponentState ) or ( StatutGrilleImput = taConsult ) then Exit ;

 lStCompte := ZReleveBanque.StGENERAL_I;

 case HGImputation.Col of
  cColGeneralImput : begin
                     if FBoGuideRun and
                     ( length( ZReleveBanque.StGENERAL_I ) < VH^.Cpta[fbGene].Lg ) then
                      LookupList(HGImputation,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL',true, 1)
                       else

                        //if not HGImputationRechercheImputation(HGImputation.Row) then
                      //  if not ZReleveBanque.IsValideGeneral( lStCompte ) then
                      //   HGImputationRechercheImputation( HGImputation.Row )
                         LookupList(HGImputation,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL',true, 1) ;
                     end; // imputation
  cColAuxiliaireImput : begin
                          ZReleveBanque.RechercheAuxiliaire(HGImputation);
                        end;
 end; // case
end;

function  TOF_CRELBQE.HGImputationRechercheTiers : boolean;
var
 lStAuxi     : string;
 lStGeneral  : string;
begin

 lStAuxi    := HGImputation.Cells[cColAuxiliaireImput,HGImputation.Row];
 lStGeneral := HGImputation.Cells[cColGeneralImput,HGImputation.Row];
 result     := ZReleveBanque.ExisteTiers( lStAuxi,
                                           lStGeneral );
 HGImputation.Cells[cColAuxiliaireImput,HGImputation.Row] := lStAuxi;
 HGImputation.Cells[cColGeneralImput,HGImputation.Row]    := lStGeneral;

end;

function  TOF_CRELBQE.HGImputationRechercheImputation( ARow : integer ) : boolean;
var
 lStCompte : string;
begin

 lStCompte := HGImputation.Cells[cColGeneralImput,ARow];

 result    := ZReleveBanque.IsValideGeneral( lStCompte  );

 if not result then exit;

 result    := ZReleveBanque.ExisteImputation( lStCompte  );

 if result then
  begin
   ZReleveBanque.StGENERAL_I                  := lStCompte ;
   HGImputation.Cells[cColGeneralImput,ARow]  := lStCompte ;
  end; // if

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... : 30/07/2001
Description .. : Assigne la ligne d'imputation courante avec las valeurs de la
Suite ........ : ligne de la grille courantes
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.HGImputationAssignLigneTOB( ARow : integer) : boolean;
begin

 result := false;

 try

  if HGImputation.Objects[cColObjectReleve,ARow] = nil then
   begin
    ZReleveBanque.AddLigneImput(ARow-1);
    HGImputation.Objects[cColObjectReleve,ARow]  := ZReleveBanque.TOBLigneImputation;
   end
    else
     ZReleveBanque.TOBLigneImputation := TOB(HGImputation.Objects[cColObjectReleve,ARow]);  // récupération de l'objet dans la TOB

  if not assigned(ZReleveBanque.TOBLigneImputation) then
   begin
     PGIBox( cStTexteErreurRecupImput,
             cStTexteTitreFenetre);
    exit;
   end;

  ZReleveBanque.StGENERAL_I := HGImputation.Cells[cColGeneralImput,ARow];

  CSetMontants( ZReleveBanque.TOBLigneImputation ,
                Valeur(HGImputation.Cells[cColDebitImput,ARow]) ,
                Valeur(HGImputation.Cells[cColCreditImput,ARow]) ,
                ZReleveBanque.Devise,
                false,
                true );

  ZReleveBanque.StLIBELLE_I         := HGImputation.Cells[cColLibelleImput,ARow];
  ZReleveBanque.StREFINTERNE_I      := HGImputation.Cells[cColReferenceImput,ARow];
  ZReleveBanque.InNUMLIGNE_I        := ARow;

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
function TOF_CRELBQE.HGImputationAssignTOB : boolean;
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
procedure TOF_CRELBQE.VideGrille ( G : THGrid );
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

procedure TOF_CRELBQE.HGImputationKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

begin
 if ( Shift = [] ) and ( key = VK_DELETE ) then
  StatutGrilleImput := taModif;
end;

procedure TOF_CRELBQE.HGImputationKeyPress(Sender: TObject; var Key: Char);
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
procedure TOF_CRELBQE.HGImputationExit(Sender: TObject);
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit;

 if FBoGuideRun then
  begin

   if PGIAskCancel(cStTexteArretGuide,cStTexteTitreFenetre) = mrYes then
    begin
     SetModeGuide(false);
     ZReleveBanque.StatutImput := taCreat;
    // ChcBouclage.Checked := false;
    end; // if PGIAsk
   end // FBoGuideRun
    else
     begin

      if ( FStatutGrilleImput <> taConsult ) then
       begin
        if (PGIAskCancel(cStTexteEchap,cStTexteTitreFenetre) = mrYes) then
         begin
          ZReleveBanque.StatutImput := taCreat;
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
procedure TOF_CRELBQE.HGSaisieRowEnter( Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin
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
procedure TOF_CRELBQE.RafraichirLigne;
begin
 FlashCurrentReleve.Caption      := TraduireMemoire(cStTexteLigneReleve) + intToStr(HGSaisie.Row);
 HGSaisie.Cells[ 0,HGSaisie.Row] := intToStr(HGSaisie.Row);
 // on affiche le solde de la ligne de releve si on avait chargé les lignes d'imputations en memoire
 if not HGSaisieAssignTOB(HGSaisie.Row) then exit;

 // mise a jour des info de la affiché dans la barre de menu
 Ecran.Caption                   := TraduireMemoire(cStTexteTitreFenetre + ' : compte ' + ZReleveBanque.StCompteContrepartie ) +
                                 '  ' + ZReleveBanque.StLibelleEtat;
 UpdateCaption(Ecran);

 {$IFDEF TEST}
 if V_PGI.SAV then
  Ecran.Caption                  := Ecran.Caption + ' : numero = ' + varToStr(ZReleveBanque.InCOMPTEUR_R);
 {$ENDIF}
 UpdateCaption(Ecran);
 HGSaisieAfficheSolde;
 VideGrille(HGImputation);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/07/2001
Modifié le ... : 20/07/2001
Description .. : Gestion de l'etat des bouttons
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.EnabledControl;
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
                             ( CmbContrepartie.Items.Count > 1 );
 CmbJournal.Enabled       := ( ( StatutGrilleReleve = taConsult ) or
                             (  FBoGrilleVide and ( StatutGrilleReleve <> taConsult ) ) ) and
                             ( CmbJournal.Items.Count > 1 );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Affectaction du statut dans la grille et mise à jour de l'etat
Suite ........ : des boutons
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.SetStatutGrilleImput ( Value : TActionFiche );
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
procedure TOF_CRELBQE.NumeroteLigne ( ARow : integer ; G : THGrid );
var
 i : integer;
begin
 for i := ARow to G.RowCount - 1 do
  G.Cells[0,i] := intToStr(i);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Test la validite de la ligne d'imputation avant de sortir
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.HGImputationRowExit( Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 HGImputationAssignLigneTOB(Ou);
 ZReleveBanque.TypeContexte := TModeSaisie;
 Cancel                     := not ZReleveBanque.IsValideLigneImput;

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
procedure TOF_CRELBQE.EdtDateValide( Sender : THCritMaskEdit );
begin

 if csDestroying in Ecran.ComponentState then Exit;

 if (length(trim(Sender.Text))) < 10 then exit;

 if IsValidDate(Sender.Text) and ( StrToDate(Sender.Text) < FDtDateDebN ) then
    begin
     PGIInfo(cStDateInfDateDebExo,cStTexteTitreFenetre);
     if Sender.CanFocus then
      Sender.SetFocus;
      exit;
    end;

  if IsValidDate(Sender.Text) and ( StrToDate(Sender.Text) > FDtDateFinN1 ) then
    begin
     PGIInfo(cStDateInfDateFinExo,cStTexteTitreFenetre);
     if Sender.CanFocus then
      Sender.SetFocus;
      exit;
    end;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Contrôle que la date de debut ne peut inférieur à date de
Suite ........ : debut de l'exercice N
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.EdtMulDateDuChange( Sender : TObject );
begin
 EdtDateValide(EdtMulDateDu);
end;

procedure TOF_CRELBQE.EdtMulDateDuExit( Sender : TObject );
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
procedure TOF_CRELBQE.EdtMulDateAuChange( Sender : TObject );
begin
 EdtDateValide(EdtMulDateAu);
end;

procedure TOF_CRELBQE.EdtMulDateAuExit( Sender : TObject );
begin
 EdtDateValide(EdtMulDateAu);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : enregistrement de l'imputation, passage à ligne suivante de
Suite ........ : releve et bouclage su cmbBouclage est selectionne
Mots clefs ... :
*****************************************************************}
procedure TOF_CRELBQE.VK_F10Imputation;

 procedure BouclageGuide;
  begin
   // on boucle en mode guide sur la prochaine imputation non imputée
   if ( HGSaisie.Row = ( HGSaisie.RowCount - 1 ) ) then
    begin
     Nouveau;    // on est sur la derniere ligne -> on sort du mode guide et on ajoute une ligne
     RemettreEnEditionHGSaisie;
    end
     else
      begin
       HGSaisie.Row := HGSaisie.Row + 1;
       HGSaisieAssignTOB(HGSaisie.Row);
       if ( ZReleveBanque.StETAT_R <> cPasImput ) then
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

 // on recupere les enregistrements de la grille
 if not HGImputationAssignTOB then
   exit;

  // on supprime les lignes vides
  ZReleveBanque.SupprimerLigneImputVide;

 // on remet a jour la grille si jamais des lignes en été supprimer
 HGImputationRemplirGrille;

 ZReleveBanque.TypeContexte := TModeSaisie;
 if not ZReleveBanque.IsValideImputation then
  exit;

  // on complete le ligne de relevé avec les info du guide
  SynchroniseHGSaisie;

  ZReleveBanque.AssignInfoTVA;

  // l'imputation est coherente et validé
  ZReleveBanque.StETAT_R := cImputValide;

  if not ZReleveBanque.SaveLigneReleve then
   begin
    PGIInfo('Erreur lors de l''enregistrement',cStTexteTitreFenetre);
    exit;
   end;

  // on rafraichis la grille des releves
  HGSaisieRefreshLigne(HGSaisie.Row);
  StatutGrilleReleve := taConsult;

  if ( HGSaisie.Row = ( HGSaisie.RowCount - 1 ) )  then
   begin
    Nouveau;    // on est sur la derniere ligne -> on sort du mode guide et on ajoute une ligne
    RemettreEnEditionHGSaisie;
   end
    else
     if ChcBouclage.Checked then
      begin
        BouclageGuide;
      end
       else
        begin
         FStatutGrilleImput := taConsult;
         HGSaisie.Row       := HGSaisie.Row + 1;
         HGSaisie.Col       := cColDate;
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
procedure TOF_CRELBQE.VK_F6Imputation;
var
 lRdSolde                 : double;
 lRdCredit                : double;
 lRdDebit                 : double;
begin

 // on calcule et on affiche le solde des imputations
 HGImputationAssignLigneTOB(HGImputation.Row);
 AfficheSoldeImputation;

 lRdSolde  := ZReleveBanque.RdSoldeDebitImput - ZReleveBanque.RdSoldeCreditImput;
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
procedure TOF_CRELBQE.HGSaisieGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit;

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
procedure TOF_CRELBQE.HGImputationGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
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
procedure TOF_CRELBQE.HGImputationDeleteRow( ARow : integer);
begin

 if ARow = ( HGImputation.RowCount - 1 )  then
  exit;

 if HGImputationAssignLigneTOB(ARow) and ZReleveBanque.DeleteCurrentLigneImputation then
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
procedure TOF_CRELBQE.RemettreEnEditionHGSaisie;
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
     HGSaisie.Col := HGSaisie.Col + 1 ;
     HGSaisie.Col := HGSaisie.Col - 1 ;
    end;

  // on rebranche les evenements
  HGSaisie.OnKeyPress            := HGSaisieKeyPress;
  HGSaisie.OnKeyDown             := HGSaisieKeyDown;
  Ecran.OnKeyDown                := FormKeyDown;
  HGImputation.OnExit            := HGImputationExit;

end;


procedure TOF_CRELBQE.CalculSoldeTheorique;
var
 lStSQL     : string;
 lDtDate    : TDateTime;
 Q          : TQuery;
 lRdVar1    : double;
 lRdVar2    : double;
begin
  // calcul du solde du compte de contrepartie
 CGetBalanceParcompte('',ZReleveBanque.StCompteContrepartie,StrToDate(EdtMulDateDu.Text) , StrToDate(EdtMulDateAu.Text),lRdVar1,lRdVar2,FRdSolDeT);

  // calcul du solde du releve en ne tenant compte que des dates
 lStSQL := ' select '                                                           +
           'sum(CRL_DEBIT - CRL_CREDIT) as D '                                   +
           'from CRELBQE '                                                        +
           'where CRL_GENERALBQE = "' + ZReleveBanque.StCompteContrepartie + '"' +
           ' and CRL_DATECOMPTABLE >= "' + UsDate(EdtMulDateDu) + '"'             +
           ' and CRL_DATECOMPTABLE <= "' + UsDate(EdtMulDateAu) + '"'            +
           ' and CRL_NUMLIGNE = 0 ';
 Q := nil;

 try

   Q := openSQL( lStSQL, true );
   FRdSolDeT := FRdSolDeT + Q.FindField('D').asFloat;

 finally
   if assigned(Q) then Ferme(Q);
 end; // try

end;

//******************************************************************************
//
//           Gestion du filtre
//
//******************************************************************************
procedure TOF_CRELBQE.BCreerFiltreClick(Sender: TObject);
begin
 NewFiltre('RLB_SAISIE',EdtFFiltres,PgcCritere);
end;

procedure TOF_CRELBQE.BSaveFiltreClick(Sender: TObject);
begin
 SaveFiltre('RLB_SAISIE',EdtFFiltres,PgcCritere);
end;

procedure TOF_CRELBQE.BDelFiltreClick(Sender: TObject);
begin
 DeleteFiltre('RLB_SAISIE',EdtFFiltres);
end;

procedure TOF_CRELBQE.BRenFiltreClick(Sender: TObject);
begin
 RenameFiltre('RLB_SAISIE',EdtFFiltres);
end;

procedure TOF_CRELBQE.BNouvRechClick(Sender: TObject);
begin
 VideFiltre(EdtFFiltres,PgcCritere) ;
 //ValeursParDefaut;
end;

procedure TOF_CRELBQE.POPFPopup(Sender: TObject);
begin
 UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,EdtFFiltres) ;
end;

procedure TOF_CRELBQE.FFiltresChange(Sender: TObject);
begin
 LoadFiltre('RLB_SAISIE',EdtFFiltres,PgcCritere);
end;

procedure TOF_CRELBQE.BPopCompteClick(Sender: TObject);
begin
 ZoomGeneral;
end;

procedure TOF_CRELBQE.BPopJournalClick(Sender: TObject);
begin
 // ouverture du guide  ( a faire mieux )
 ParamGuide('','NOR',taModif) ;
 ZGuideTreso.ClearGuide;
 ZGuideTreso.Load;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /    
Description .. : Si on etait en modif ou en creation -> msg pour prevenir 
Suite ........ : l'utilisateur
Mots clefs ... : 
*****************************************************************}
procedure TOF_CRELBQE.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

 CanClose := ( StatutGrilleReleve = taConsult ) and ( StatutGrilleImput = taConsult );

 if not CanClose then
  CanClose := PGIAskCancel(cStTexteEchap,cStTexteTitreFenetre) = mrYes;

 if CanClose then
  TFVierge(Ecran).FormCloseQuery(Sender,CanClose);

end;

procedure TOF_CRELBQE.ZoomGeneral;
begin
 {$IFNDEF EAGLCLIENT}
 {$IFNDEF SANSCOMPTA}
 FicheGene(nil, '', ZReleveBanque.StGENERAL_R, taConsult, 0)
 {$ENDIF}
 {$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /    
Description .. : Zoom sur le compte de contrepartie
Mots clefs ... : 
*****************************************************************}
procedure TOF_CRELBQE.ZoomEcrs;
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
 {$IFNDEF EAGLCLIENT}
 {$IFNDEF SANSCOMPTA}
 OperationsSurComptes( ZReleveBanque.StCompteContrepartie, lStExo , '', '') ;
 {$ENDIF}
 {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /    
Description .. : Ajout des menus supplementaire
Mots clefs ... : 
*****************************************************************}
procedure TOF_CRELBQE.POPSPopup ( Sender : TObject) ;
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

   if UpperCase( lT.Name ) = 'PMODE' then
    lRS := ShortCut( VK_F6 , [] )
     else
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

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /    
Description .. : Réaffectation de la table après annulation
Mots clefs ... : 
*****************************************************************}
procedure TOF_CRELBQE.ReaffecterLigneReleve;
begin

 try

  ZReleveBanque.DeleteCurrentImputation;
  ZReleveBanque.TOBLigneReleve.Dupliquer(FTOBSauveLigne,true,true);
  HGSaisieRefreshLigne(HGSaisie.Row);
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
procedure TOF_CRELBQE.BModeClick(Sender: TObject);
begin
 FBoModeSelection := true;
 SetModeSelection;
end;

procedure TOF_CRELBQE.BFermerClick(Sender: TObject);
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
procedure TOF_CRELBQE.SetModeSelection;
begin

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
Modifié le ... :   /  /
Description .. : Transfert des lignes de CETEBAC vers CRELBQE
Mots clefs ... :
*****************************************************************}
function TOF_CRELBQE.Integration( vNumEtebac : integer ) : boolean;
var
 lTOBLigneEtebac            : TOB;
 lTOBEtebac                 : TOB;
 lTOB                       : TOB;
 Q                          : TQuery;
 lInNumLigne                : integer;
 lStChampsDateComptable     : string;
 lStChampsDateValeur        : string;
begin

 result           := true;

 Q                := nil;

 lTOBEtebac       := TOB.Create('',nil,-1);

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
                'CET_CREDIT '                 +
                'from CETEBAC '               +
                'where CET_NUMRELEVE = '      + intToStr(vNumEtebac)      +
                'and CET_TYPELIGNE = "04" '   +
                'and CET_DEVISE = "'          + V_PGI.DevisePivot         + '" ' ,
                true );


  lTOBEtebac.LoadDetailDB('CETABAC','','',Q,false);

 finally
  if assigned(Q) then Ferme(Q);
 end; // try

 Q := nil;

 try

  Q := OpenSQL ('select J_CHOIXDATE from JOURNAL '    +
                'where J_JOURNAL="' + ZReleveBanque.StJournalContrepartie + '"' , true);


  lStChampsDateComptable := Q.findField('J_CHOIXDATE').asString;

  if lStChampsDateComptable = '' then
   lStChampsDateComptable := 'DATEOPERATION';

 if lStChampsDateComptable = 'DATEOPERATION' then
  lStChampsDateValeur := 'DATEVALEUR'
   else
    lStChampsDateValeur := 'DATEOPERATION';

 finally
  if assigned(Q) then Ferme(Q);
 end; // try

 if lTOBEtebac.Detail.Count = 0 then
  exit;

 lInNumLigne     := 0;


 while ( lTOBEtebac.Detail.Count ) > lInNumLigne do
  begin

   ZReleveBanque.AddLigneReleve;
   lTOBLigneEtebac                           := lTOBEtebac.Detail[lInNumLigne];
   ZReleveBanque.StIMPORT_R                  := 'X';
   ZReleveBanque.StGENERALBQE_R              := ZReleveBanque.StCompteContrepartie; // compte de banque
   ZReleveBanque.StJOURNAL_R                 := ZReleveBanque.StJournalContrepartie;
   ZReleveBanque.DtDATECOMPTABLE_R           := lTOBLigneEtebac.GetValue('CET_' + lStChampsDateComptable); //
   ZReleveBanque.DtDATEVALEUR_R              := lTOBLigneEtebac.GetValue('CET_' + lStChampsDateValeur); //
   ZReleveBanque.StREFINTERNE_R              := lTOBLigneEtebac.GetValue('CET_NUMEROPIECE');   // numero de cheque ou de piece
   CSetMontants( ZReleveBanque.TOBLigneReleve,
                 lTOBLigneEtebac.GetValue('CET_DEBIT') ,
                 lTOBLigneEtebac.GetValue('CET_CREDIT') ,
                 ZReleveBanque.Devise,
                 false,
                 true );
   if ZReleveBanque.RdDEBIT_R > 0  then
    ZReleveBanque.StNATUREPIECE_R := 'RC'
     else
     ZReleveBanque.StNATUREPIECE_R := 'RF';

   ZReleveBanque.StLIBELLE_R                 := lTOBLigneEtebac.GetValue('CET_LIBELLE');
   ZReleveBanque.StCODEAFB_R                 := lTOBLigneEtebac.GetValue('CET_CODEAFB');
   ZReleveBanque.StREFPIECE_R                := lTOBLigneEtebac.GetValue('CET_REFPIECE');
   ZReleveBanque.StDEVISE_R                  := lTOBLigneEtebac.GetValue('CET_DEVISE');
   ZReleveBanque.InCOMPTEUR_R                := FInNumReleve;

   FInNumReleve                              := FInNumReleve + 1;

   lTOB                                      := TOB.Create( '' , nil , -1 );

   result := ZGuideTreso.RechercheGuideEnBase( FloatToStr( ZReleveBanque.RdCREDIT_R),
                                               FloatToStr( ZReleveBanque.RdDEBIT_R),
                                               VarToStr( ZReleveBanque.StLIBELLE_R ),
                                               VarToStr( ZReleveBanque.StGENERAL_R ),
                                               FloatToStr( ZReleveBanque.RdTAUXTVA_R),
                                               DateToStr( ZReleveBanque.DtDATECOMPTABLE_R ),
                                               VarToStr( ZReleveBanque.StREFINTERNE_R ),
                                               VarToStr( ZReleveBanque.StREFEXTERNE_R ),
                                               VarToStr( ZReleveBanque.StCODEAFB_R ),
                                               lTOB,
                                               nil );  // on le repasse la grille pour l'affichage de lookup de choix des guides

   if result then
    begin
     // on affecte les enregistrement contenu dans lTOB à la ligne de releve courante
     ZReleveBanque.CreerImputation(lTOB);
     // on supprime les lignes vides
     ZReleveBanque.SupprimerLigneImputVide;
     lTOB.Free;
     ZReleveBanque.CalculSoldeImput;

     if ZReleveBanque.IsValideImputation then
      ZReleveBanque.StETAT_R := cImputCorrectNonValide
       else
        ZReleveBanque.StETAT_R := cImputIncorrecte;

      end; // if

  lInNumLigne  := lInNumLigne + 1;

   // gestion de la fenetre d'attente
  if FTT.Canceled then  exit;
  FTT.Value     := FTT.Value + 1 ;
  FTT.SubText   := 'Integration du releve ' + IntToStr( FInNumReleve ) ;
  Application.ProcessMessages;

 end; // while

 finally
  if assigned(lTOBEtebac) then lTOBEtebac.Free;
 end; // try

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /    
Description .. : Enregistrement en base des lignes de releves et 
Suite ........ : suppression des lignes CETEBAC
Mots clefs ... : 
*****************************************************************}
procedure TOF_CRELBQE.TransactionSaveReleve;
var
 lStValue : string;
begin

  if ZReleveBanque.SaveReleve then
   begin

    if not CMAJNumeroSouche ('TRE', 'NOR' ,FInNumReleve - 1,FInFirstNumReleve) then
     raise EAbort.Create('Erreur a la maj des compteurs');

    // suppression des lignes etebac integrees
    while ( FStListeNumASupp <> '' ) do
     begin

      lStValue     := ReadTokenST(FStListeNumASupp);
      if lStValue <> '' then
       ExecuteSQL('delete CETEBAC where CET_NUMRELEVE = ' + lStValue);

     end; // while

   end; // if

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /    
Description .. : Recherche des lignes CETEBAC correspondant au RIB du 
Suite ........ : journal de abnque
Mots clefs ... : 
*****************************************************************}
procedure TOF_CRELBQE.RechercheRibAIntegrer;
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
begin

  Q     := nil;

  try

   Q := OpenSQL( 'select CET_NUMLIGNE, '                          +
                  'CET_ETABBQ, '                                  +
                  'CET_NUMEROCOMPTE, '                            +
                  'CET_GUICHET '                                  +
                  'from CETEBAC, BANQUECP  '                      +
                  'where CET_ETABBQ = BQ_ETABBQ '                 +
                  'and CET_NUMEROCOMPTE = BQ_NUMEROCOMPTE '       +
                  'and CET_GUICHET = BQ_GUICHET '                 +
                  'and BQ_GENERAL = "'                            + ZReleveBanque.StCompteContrepartie + '" ' +
                  'AND BQ_NODOSSIER="'                            + V_PGI.NoDossier                    + '" ' + // 24/10/2006 YMO Multisociétés
                  'and CET_TYPELIGNE = "01" '                     +
                  'and CET_DEVISE = "'                            + V_PGI.DevisePivot                  + '" '
                  ,
                  true );

   if Q.EOF then
    exit;

   lStETABQ        := Q.FindField('CET_ETABBQ').asString;
   lStGUICHET      := Q.FindField('CET_GUICHET').asString;
   lstNUMEROCOMPTE := Q.FindField('CET_NUMEROCOMPTE').asString;


   if assigned(Q) then Ferme(Q);
   Q := nil;

   lStListeNum     := '';

   lTempData       := TheData;

   // lance la fiche d'integration des releves fenetre RLVETEBAC et TOF  TOF_CETEBAC ( fichier DP\lib\dpTOFCETEBAC )
   AGLLanceFiche('CP','RLVETEBAC','','', lStETABQ         + ';' +
                                         lStGUICHET       + ';' +
                                         lstNUMEROCOMPTE  + ';' +
                                         'N'              + ';'  );

   T                  := TTemp(TheData);

   if not assigned(T) then
    exit;

   if ( ZGuideTreso.EstVide ) and ( PGIAsk(cStTexteAucunGuideEtebac , cStTexteTitreFenetre ) = mrNo ) then exit;

   lStListeNum        := T.StValue;
   FStListeNumASupp   := lStListeNum;
   T.free;

   theData            := lTempData; // restauration de la valeur de theData

   // l'utilisateur n'a pas integre de releve
   if lStListeNum = '' then
    exit;

   if Blocage([cStVerrouTreso], true , cStVerrouTreso ) then
    exit;

   FInNumReleve                 := ZReleveBanque.ProchainCompteur;
   FInFirstNumReleve            := FInNumReleve;
   ZGuideTreso.StCompteTVA      := ZReleveBanque.StCompteTVA;
   ZReleveBanque.TypeContexte   := TModeAuto;

   FTT                          := DebutProgressForm ( Ecran,
                                                       cStTexteTraitementEcr,
                                                       cStTexteTitreFenetre,
                                                       100,
                                                       true,
                                                       true ) ;


   while ( lStListeNum <> '' ) do
    begin
     lStValue     := ReadTokenST(lStListeNum);
     if lStValue <> '' then
      begin
       lInNumEtabac := StrToInt(lStValue);
       Integration(lInNumEtabac);
      end; // if
    end; // while


   // gestion de la fenetre d'attente
  if FTT.Canceled then  exit;
  FTT.Value     := FTT.Value + 1 ;
  FTT.SubText   := 'Enregistrement en base' ;
  Application.ProcessMessages;

  if Transactions(TransactionSaveReleve,1) <> oeOK then
   MessageAlerte('Erreur sur lors de l''enregistrement des fichiers' + #13#10 + V_PGI.LastSQLError );

 finally
  if assigned(Q)    then Ferme(Q);
  if assigned(FTT)  then FTT.Free;
  Bloqueur( cStVerrouTreso , false );
 end; // try

end;


Initialization
  registerclasses ( [ TOF_CRELBQE ] ) ;
end.
