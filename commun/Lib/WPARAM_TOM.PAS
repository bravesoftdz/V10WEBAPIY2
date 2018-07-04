{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 08/02/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : WPARAM (WPARAM)

ATTENTION cette unité peut être dans un contexte HORS GPAO

Pensez donc à bien encadrer vos fonctions GPAO ou utiliser les structures existantes

($IFDEF GPAO)
  // Fonction GPAO
($ENDIF)

// Fonction GC


Mots clefs ... : TOM;WPARAM
*****************************************************************}
Unit wParam_Tom;

Interface

Uses
	StdCtrls,
  DbCtrls,
  Controls,  // MrYes
  Classes,
  Dialogs,
  Forms,
  SysUtils,
	ComCtrls,
	ComObj,
  Menus,
  Windows,
  Messages,
 	Graphics,
  extctrls,
  strUtils,
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    {$IFNDEF EAGLSERVER}
      Fe_Main,
      Fiche,
      Mul,
    {$ENDIF !EAGLSERVER}
  {$ELSE}
    MainEagl,
    eFiche,
    eMul,
  {$ENDIF}
  HCtrls,
  Grids,
  HEnt1,
  HDB,
  HMsgBox,
  Htb97,
  Formule,
  {$IFDEF GPAOLIGHT}
    EntGP,
  {$ENDIF GPAOLIGHT}  
  {$IFDEF GPAO}
    wVersion,
  {$ENDIF GPAO}
	wTom,
  wCommuns,
  {$IFNDEF EAGLSERVER}
    hcapCtrl,
  {$ENDIF !EAGLSERVER}
  uTob
  ;

Type
  TOM_WPARAM = Class (twTOM)
    procedure OnNewRecord               ; override ;
    procedure OnDeleteRecord            ; override ;
    procedure OnUpdateRecord            ; override ;
    procedure OnAfterUpdateRecord       ; override ;
    procedure OnLoadRecord              ; override ;
    procedure OnChangeField (F: TField) ; override ;
    procedure OnArgument (S: String )   ; override ;
    procedure OnClose                   ; override ;
    procedure OnCancelRecord            ; override ;
    private
      Title         : boolean;
      MemoLastError : integer;
      iIdentifiantWPE: integer;
      WithCancelInfo: Boolean;
      {$IFDEF GPAO}
      lInitdata : Boolean;
      {$ENDIF GPAO}
      {$IFDEF EAGLCLIENT}
        EnabledOnLoadRecord: boolean;
      {$ENDIF}

      {$IFDEF GPAO}
      LBITINERAIRE, LBCIRCUIT  : tHLabel;
      LBNOMENCLATURELIBELLE, LBNOMENCLATUREREVISION, LBNOMENCLATUREMAJEUR, LBNOMENCLATUREMINEUR : tHLabel;
      LBGAMMELIBELLE       , LBGAMMEREVISION       , LBGAMMEMAJEUR       , LBGAMMEMINEUR        : tHLabel;
      LBUNITEPROD1, LBUNITEPROD2, LBUNITEPROD3 : tHLabel;
      {$ENDIF GPAO}

      LBARTICLE, LBAFFAIRE, LBATELIER : thLabel;

      {$IFNDEF EAGLSERVER}
        LBTIERS, LBLIVRE  : tHLabel;
        BTRechercheTarifaire  : tToolBarButton97;

        {$IFDEF GPAO}
          BTCalcul    : tToolBarButton97;
        {$ENDIF GPAO}
      {$ENDIF !EAGLSERVER}

      BMP: TBitmap;

      {$IFDEF GPAO}
        CleWNT : tCleWNT;
        CleWGT : tCleWGT;
        WPA_COMBO01, WPA_COMBO02 : thValComboBox;
      {$ENDIF GPAO}
      {$IFNDEF EAGLSERVER}
        {$IFDEF GPAO}
          WPA_BOOLEAN01 : tDbCheckBox;
          RGPRIX, RGUNPDRPAR : tHRadioGroup;
        {$ENDIF GPAO}
        {$IFNDEF ERADIO}
          {$IFDEF GPAOLIGHT}
            PlanLivrTobFields, PlanLivrTobItems, PlanLivrTobIndicateurs: Tob;
            PlanLivrSB: TScrollBox;
          {$ENDIF GPAOLIGHT}
        {$ENDIF ERADIO}
        GAFFICHAGE: thGrid;      //Grille d'affichage des origines des différentes composantes du système tarifaire
        MBLOCNOTE : tMemo;       //Champs Commentaire contenant les autres détails de l'origine tarifaire
      {$ENDIF !EAGLSERVER}

      GBSousTraitance         : tGroupBox;
      NatureTravail           : String;
      Article, CodeArticle    : String;
      sFonctionnalite, sRang  : String;

      sCodePort, sTiersFrais    : string;
      sTypeFourni, sRepartition : string;
      sTypePDR: string;

      TobContexte, TobTarif, TobPrixMarge, TobAffichage, TobUneLigneTarif : Tob;
      sAppel : string;

      sParametre : string;

      {$IFDEF GPAO}
        l1erClickBoolean01 : boolean;
      {$ENDIF GPAO}

      {$IFDEF STK}
        NbDeci:integer;
      {$ENDIF STK}

      OpeCirc, CodeComposant : string;
      LigneOrdre	: Integer;

      { VAR : Remplacement de composant --------------------------------------------------------------------------- }
      GRILLE: thGrid;
      //CODECOMPOSANT: thEdit;
      TobWPA, TobWPA_Fils: Tob;
      FromPutEcran, FromApplyChangement, FromBtMoulinette: Boolean;
      LastErrorDejaTraite: Boolean;
      { ----------------------------------------------------------------------------------------------------------- }

      {$IFNDEF EAGLSERVER}
        {$IFNDEF ERADIO}
          { Bouton fermer }
          procedure bFerme_Click(Sender: tObject);

          {$IFDEF GPAO}
            procedure WPA_VARCHAR05_OnElipsisClick(Sender: Tobject);
            procedure WPA_VARCHAR06_OnElipsisClick(Sender: Tobject);
          {$ENDIF GPAO}

          {$IFDEF SAV}
            procedure WPA_VARCHAR01_OnElipsisClick(Sender: Tobject);
            procedure WPA_VARCHAR02_OnElipsisClick(Sender: Tobject);
          {$ENDIF SAV}
        {$ENDIF !ERADIO}
      {$ENDIF !EAGLSERVER}

      //Calcul Prix De Revient
      {$IFDEF GPAO}
	      {$IFNDEF EAGLSERVER}
          procedure BTCalcul_OnClick(Sender: tObject);
	      {$ENDIF !EAGLSERVER}
        procedure AffichageParDefaut(lDefaut : boolean);
      {$ENDIF GPAO}

      {$IFNDEF EAGLSERVER}
        {$IFDEF GPAO}
          //Configuration par défaut ?
          procedure WPA_BOOLEAN01_OnClick(Sender: tObject);
        {$ENDIF GPAO}

        //Recherche Tarifaire
        procedure BTRechercheTarifaire_OnClick(Sender: tObject);

        // Loupe
        procedure MnLpLegende_OnClick(Sender: TObject);
        {$IFDEF GPAOLIGHT}
          procedure ShowCodeNomenAss;
        {$ENDIF GPAOLIGHT}
      {$ENDIF !EAGLSERVER}

      {$IFDEF GPAOLIGHT}
      procedure SetConversion;
      {$ENDIF GPAOLIGHT}

      procedure ClientOuFournisseur;
      procedure CtrlAffichageMarqueChoixQualiteIndice;
      procedure ControleAffichageDimensionsArticle;
      procedure SetSayOuGet(sControl: string; lConditionSay: boolean);

      Procedure SetAdresseTiers(InfosTiers:MyArrayValue);
      {$IFNDEF EAGLSERVER}
        procedure GAffichageCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
        procedure GAffichageOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure wAffichageDesTarifsGrille(iCol, iRow: Integer; Rect: Trect; Canvas: TCanvas; lFocus : boolean);
      {$ENDIF !EAGLSERVER}

      {$IFDEF GPAOLIGHT}
        procedure SetCodeArticleWNT;
        function  GetCleWNT: tCleWNT;
        procedure SetQWNT;
        {$IF not(Defined(EAGLSERVER) or Defined(ERADIO))}
          procedure BTCalculFaisabilite_OnClick(Sender: tObject);
        {$IFEND !EAGLSERVER OR !ERADIO}
      {$ENDIF GPAOLIGHT}
      {$IFDEF GPAO}
        function GetCleWVS: tCleWVS;
        {$IFNDEF EAGLSERVER}
          {$IFNDEF ERADIO}
            procedure btMoulinette_OnClick(Sender: tObject);
            procedure Cell_OnElipsisClick(Sender: TObject);
            procedure GRILLE_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
            procedure GRILLE_OnDblClick(Sender: TObject);
          {$ENDIF !ERADIO}
        {$ENDIF !EAGLSERVER}
      {$ENDIF GPAO}

      { METHODES : Remplacement de composants --------------------------------------------------------------------- }
      { événements : boutons + grille + RadioButton + Champs numériques... }
      procedure BTVALIDCOMPOSANTS_OnClick(Sender: tObject);
      procedure btSupprimerFils_OnClick(Sender: tObject);
      procedure btAnnulerFils_OnClick(Sender: tObject);
      procedure btNouveauFils_OnClick(Sender: tObject);
      procedure btEraseFils_OnClick(Sender: tObject);
      {$IFNDEF EAGLSERVER}
        {$IFNDEF ERADIO}
          {$IFDEF GPAO}
            procedure PmWND_OnPopup(Sender: tObject);
            procedure MnWNA_OnClick(Sender: tObject);
            procedure MnLpVisuNome_OnClick(Sender: TObject);
          {$ENDIF GPAO}
        {$ENDIF !ERADIO}
      {$ENDIF !EAGLSERVER}
      procedure RafraichirGrille;
      procedure GRILLE_OnExit(Sender: TObject);
      procedure GRILLE_OnEnter(Sender: TObject);
      procedure GRILLE_OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
      procedure GRILLE_OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
      procedure GRILLE_OnKeyPress(Sender: TObject; var Key: Char);
      procedure RBRIENLIEN_OnClick(Sender: tObject);
      procedure RBNOUVQLIEN_OnClick(Sender: tObject);
      procedure RBCOEFQLIEN_OnClick(Sender: tObject);
      procedure RBREVISION_OnClick(Sender: tObject);
      procedure Fields_OnChange(Sender: tObject);
      procedure CBAPPLIQUER_OnClick(Sender: tObject);
      { validité des données }
      function wEstUnFilsValide(var iCol, iRow: Integer): tWinControl;
      { Affichage }
      procedure wInitLienFields(var T: Tob);
      procedure wInitUnitesLien(var TobWPA_Fils: Tob);
      procedure wSetAffichageQLien(ctx_QTE_ou_COEF: String);
      function ParticularitesSousTraitance : boolean;

      procedure CB_ANALYSE_OnClick(Sender: tObject);
      {$IFNDEF EAGLSERVER}
        { Touches }
        procedure Ecran_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

        {$IFNDEF ERADIO}
          { Champs, Cellules }
          procedure Tiers_OnElipsisClick(Sender: TObject);
          procedure TiersLivre_OnElipsisClick(Sender: TObject);
        {$ENDIF !ERADIO}
      {$ENDIF !EAGLSERVER}

{$IFDEF QUALITE}
				procedure Ressource_OnElipsisClick(Sender: TOBject);
{$ENDIF QUALITE}

      {CheckBox pour ControlField}
      procedure WPA_BOOLEAN12_OnClick(Sender: tObject);
      procedure WPA_BOOLEAN13_OnClick(Sender: tObject);
      procedure MyOnGetArticle(const sArticle: String);

      {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
      {$IFDEF GPAOLIGHT}
        { Paramétrage Planning de livraison }
        procedure SetPlanLivrOnLoad;
        procedure PlanLivrSetColor;
        procedure PlanLivrSetIndicateurs;
        procedure PlanLivrRecolleLeMemo;
        procedure PlanLivrSetEvents;
        procedure PlanLivrSelectColor(Sender: TObject);
        procedure PlanLivrAddToList(L: THListBox);
        procedure PlanLivrUpdateList(Const L: THListBox; Const R: THRadioGroup);
        function  PlanLivrGetItemCaption(Const R: THRadioGroup; Const T: Tob): HString;
        procedure PlanLivrChangeMode(Sender: TObject);
        procedure PlanLivrBlocNoteChange(Sender: TObject);
        procedure PlanLivrOnDblClick(Sender: TObject);
        procedure PlanLivrComboEtatOnChange(Sender: TObject);
        procedure PlanLivrComboEtatOnEnter(Sender: TObject);
        procedure PlanLivrLabelOnClick(Sender: TObject);
        procedure PlanLivrTriAddOnClick(Sender: TObject);
        procedure PlanLivrTriDelOnClick(Sender: TObject);
        procedure PlanLivrSetWPA_LONGVARCHAR02;
        procedure PlanLivrReadWPA_LONGVARCHAR02;
        {$ENDIF GPAOLIGHT}

        {$IFDEF GPAO}
          { Paramétrage Planning Plan de charge }
          procedure SetPlanDeChargeOnLoad;
          procedure PlanDeChargeSetColor;
          procedure PlanDeChargeSetEvents;
          procedure PlanDeChargeSelectColor(Sender: TObject);
          {$IFDEF AFFAIRE}
            Procedure WPA_VARCHAR01_OnChange(Sender: TObject);
          {$ENDIF AFFAIRE}
        {$ENDIF GPAO}
      {$ENDIF !ERADIO}
      {$ENDIF !EAGLSERVER}

  protected
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        procedure BValiderOnClick(Sender: TObject);
        {$IFDEF AFFAIRE}
          procedure BRECHAFFAIRE_OnClick(Sender: TObject); override;
        {$ENDIF AFFAIRE}
      {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}

    {$IFDEF GPAO}
      {$IFNDEF EAGLSERVER}
        {$IFNDEF ERADIO}
          {$IFDEF AFFAIRE}
            procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Tiers: THEdit); override;
          {$ENDIF AFFAIRE}
        {$ENDIF ERADIO}
      {$ENDIF EAGLSERVER}
    {$ENDIF GPAO}

    procedure ControlField(Const FieldName: string); override;
    procedure CalculField(Const FieldName: string); override;
    Function  RecordIsValid: boolean; override;
  public
  end ;

Var
  IsoWizard          : Variant;
	CellNewComposant, CellOldComposant: String;

const

//GP_DKZ_GP_PDC_20070927
  TexteMessage: array[1..45] of string=
       (
        {1} 'La nature de travail de cet article n''est pas active.'
        {2} ,'Code article inconnu'
        {3} ,'Nature travail inconnu'
        {4} ,'Itineraire inconnu'
        {5} ,'Circuit inconnu pour l''itinéraire'
        {6} ,'Nature travail + Nomenclature + Indice Majeur Inconnu'
        {7} ,'Indice Majeur Nomenclature est obligatoire'
        {8} ,'Nature travail + Gamme + Indice Majeur Inconnu'
        {9} ,'Indice Majeur Gamme est obligatoire'
       {10} ,'Vous devez valider votre écran des paramètres avant de lancer le calcul'
       {11} ,'Vous ne pouvez pas sur-lancer une phase'
       {12} ,'Vous devez valider votre écran des paramètres avant de lancer le calcul des besoins nets'
       {13} ,'Code tiers inconnu ou non renseigné'
       {14} ,'Code Article inconnu ou non renseigné'
       {15} ,'Code Tarif Spécial inconnu'
       {16} ,'Code Tiers Inconnu ou non renseigné'
       {$IFDEF GPAO}
       {17} ,'Le couple article / nature travail n''existe pas'
       {$ELSE GPAO}
       {17} ,'L''article n''est pas de type assemblage'
       {$ENDIF GPAO}
       {18} ,'Pas de nomenclature pour ce couple article / nature.'
       {19} ,'Le code nomenclature n'' est pas valide.' + #13 + ' Vérifiez dans la liste.'
       {20} ,'Vous devez sélectionner un composant existant.'
       {21} ,'Vous ne pouvez pas remplacer un composant par lui-même si ce n''est pour modifier les quantité ou unité de lien.' + #13 + #10 + ' Vous devez sélectionner un autre composant ou modifier les quantité ou unité de lien.'
       {22} ,'Vous avez déjà décidé de remplacer ce composant.' + #13 + #10 + ' Vous devez sélectionner un autre composant.'
       {23} ,'Vous devez saisir une quantité de lien positive.'
       {24} ,'Vous devez saisir un coefficient multiplicateur pour la quantité de lien, positif.'
       {25} ,'La version n''est pas renseignée ou n''existe pas pour cet article'
       {26} ,'Transporteur inconnu ou non renseigné'
       {27} ,'Le code circuit est inconnu'
       {28} ,'Le code circuit est inconnu pour l''itinéraire de l''article'
       {29} ,'Le code phase est inconnu pour le circuit'
       {30} ,'Le code traitement est inconnu'
       {31} ,'Code affaire inconnu'
       {32} ,''
       {33} ,'Le code atelier est inconnu'
       {34} ,''
       {35} ,'Le n° de regroupement est inconnu.' + #13 + ' Vérifiez dans la liste.'
       {36} ,'Le nombre maximum de colonnes est de 13.'
       {37} ,'Saisir une valeur positive.'
       {38} ,'Vérifier la fourchette de dates.'
       {39} ,'La quantité réceptionnée doit être positive'
       {40} ,'La quantité lancée doit être positive'
       {41} ,'Vous devez saisir un composant dont le mode de consommation est renseigné.'
       {42}  ,'La ressource est obligatoire.'
       {43}  ,'La ressource saisie est inconnue.'
       {44}  ,'La méthode de valorisation est obligatoire.'
//GP_DKZ_GP_PDC_20070927
       {45}  ,'Vérifier la fourchette des heures.'

       );
 	ErrNatArtInactive                   =  1;
	ErrCodeArticle                      =  2;
  ErrCodeNature                       =  3;
  ErrCodeItineraire                   =  4;
  ErrCodeCircuit                      =  5;
  ErrCodeNomenclature                 =  6;
  ErrCodeIndiceMajeurNomenclature     =  7;
  ErrCodeGamme                        =  8;
  ErrCodeIndiceMajeurGamme            =  9;
  ErrCalculEnModeEdit                 = 10;
  ErrQLancement                       = 11;
  ErrCalculCBNEnModeEdit              = 12;
  ErrCodeTiers                        = 13;
  ErrArticle                          = 14;
  ErrCodeTarifSpecial                 = 15;
  ErrCodeTiersLivre                   = 16;
  ErrArtNat                           = 17;
  ErrArticleWNT                       = 18;
  ErrArticleWNT2                      = 19;
  ErrEmptyComposant							      = 20;
  ErrSameComposant							      = 21;
  ErrEvenChangedComposant					    = 22;
  ErrQLien										        = 23;
  ErrCoefLien									        = 24;
  ErrCodeVersion                      = 25;
  ErrTransporteur                     = 26;
	ErrCircuit                          = 27;
	ErrCircuitDeItineraire              = 28;
	ErrPhaseduCircuit                   = 29;
	ErrTraitement                       = 30;
	ErrCodeAffaire                      = 31;
  ErrAtelier                          = 33;
  ErrCodeOrdre                        = 35;

  ErrParamStkProjete                  = 36;
  ErrValeurPositif                    = 37;
  ErrFourchetteDate                   = 38;
  ErrQRecWOL                          = 39;
  ErrQLanWOP                          = 40;
  ErrModeConso									      = 41;
  ErrRessource												= 42;
  ErrExistRessource				 						= 43;
  ErrMethodeValorisation              = 44;
//GP_DKZ_GP_PDC_20070927
  ErrFourchetteHeure                  = 45;
  { -------------------------- }

  sPrixDeRevient     = 'PRIXDEREVIENT';
  sTarifsRecherche   = 'TARIFSRECH';
  sFournisseur       = 'FOU';

  { Remplacement de composants ------------------- }
  { Noms de champs }
  chgNATURETRAVAIL = 'WPA_VARCHAR01';
  chgARTICLE 		   = 'WPA_VARCHAR02';
  chgFAMILLE1		   = 'WPA_VARCHAR03';
  chgFAMILLE2   	 = 'WPA_VARCHAR04';
  chgFAMILLE3 	 	 = 'WPA_VARCHAR05';
  REVISION			   = 'WPA_VARCHAR06';
  DATEAPPLICATION  = 'WPA_DATE01';
  { Codeparam }
  sChangementComposant = 'CHGCOMPOSANT';
  { ---------------------------------------------- }

Implementation

uses
  {$IFDEF GPAOLIGHT}
    wNatureTravail,
    wFaisabilite,
    wArtNat,
    wNomeTet,
    wNomelig,
    wRevision,
    FactNomen,
//GP_20071206_TP_GC15634
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        wPlanLivr,
        AfDialog,
      {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
  {$ENDIF GPAOLIGHT}
  {$IFDEF GPAO}
    wGammeTet,
    wItineraire,
    wPrixDeRevient,
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        wChgComposant,
      {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
    {$IFNDEF QUALITE}
      wRessource,
    {$ENDIF !QUALITE}
    wChangement,
    wNomeDec,
  {$ENDIF GPAO}
  {$IFDEF STK}
    DispoDetail,
    StockUtil,
  {$ENDIF STK}
  UtilArticle,
  {$IFDEF AFFAIRE}
    AffaireUtil,
  {$ENDIF AFFAIRE}
  yTarifs,
  yTarifsCommun,
  TiersUtil,
  {$IFNDEF EAGLSERVER}
    LookUp,
  {$ENDIF !EAGLSERVER}
  AGLInitGC,
  EntGC,
  ParamSoc,
  hPanel,
  {$IFDEF QUALITE}
    wRessource,
  {$ENDIF QUALITE}

  {$IFDEF GPAO}
    wPilotageDeFlux,
  {$ENDIF GPAO}
  wParam,
//GP_20071011_PCO_GP14299
    UtilGc,
  Math,
  {$IFDEF GPAO}
    gcMarque,
    gcChoixQualite,
    {$IFDEF AFFAIRE}
      wAffaires,
    {$ENDIF AFFAIRE}
  {$ENDIF GPAO}
  {$IFDEF STK}
    IndiceArticle,
  {$ENDIF STK}
  Ent1,
  UtilPGI,
  UTOM,
  TntStdCtrls
  ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnNewRecord ;
begin
  {$IFDEF EAGLCLIENT}
    EnabledOnLoadRecord := true;
  {$ENDIF}
  Inherited ;
  if (copy(GetString('WPA_CODEPARAM'), 1, 10) = sTarifsRecherche) then
  begin
    {Restriction par utilisateur}
    if FieldExists('WPA_COMBO02') then
    begin
      SetString('WPA_COMBO02', VH^.ProfilUserC[PrEtablissement].Depot);
      if Assigned(Ecran) and Assigned(GetControl('WPA_COMBO02')) and tWinControl(GetControl('WPA_COMBO02')).Enabled then
        SetControlEnabled('WPA_COMBO02', not VH^.ProfilUserC[PrEtablissement].ForceDepot);
    end;
  end
  ;  
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnDeleteRecord ;
begin
  Inherited ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnUpdateRecord ;
var
  tsBlocNote : tStringList;
begin
  DisableControl;
  try
    if Ikc = 'C' then
    begin
      if (copy(GetString('WPA_CODEPARAM'),1,10) = sTarifsRecherche)   then
      begin
        SetString('WPA_COMBO01'   , Copy(GetString('WPA_CODEPARAM'),14,3)); //Fonctionnalité
        SetString('WPA_VARCHAR15' , Copy(GetString('WPA_CODEPARAM'),18,18)); //CodePort
        SetString('WPA_COMBO03' , V_PGI.DevisePivot);                     //Devise du tarif : Monnaie de Tenue
        SetDateTime('WPA_DATE01', V_PGI.DateEntree);                      //Pour le : pour tenir compte des dates de validité
        SetDouble('WPA_DOUBLE09', 1);                                     //Pour une quantité de : ^pour tenir compte des fourchettes
      end
      {$IFDEF GPAO}
        else if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
        begin
          SetString('WPA_COMBO03', 'FAB');          //NatureTravail de Nomenclature - Gamme
          SetDouble('WPA_DOUBLE01', 1   );          //Pour une qté de prix de revient de 1
          SetDouble('WPA_DOUBLE02', 0   );          //Pour une quantité économique de fabrication de 1
          SetDouble('WPA_DOUBLE03', 1   );          //Prix de Revient Exprimé pour une qté de
          SetString('WPA_COMBO08', 'E'  );          //Fractionnement des frais fixe par la quantité économique
          SetString('WPA_COMBO04', '01' );          //Basé sur les tarif Achat pour valoriser la matière
          SetString('WPA_COMBO05', '010');          //Enregistrement d'un prix de revient par An
          SetString('WPA_COMBO06', 'S'  );          //Taux de la ressources issus de la comptabilité
          SetString('WPA_VARCHAR07', 'ACT');        //Type de Prix de Revient
          SetString('WPA_VARCHAR08', 'BTH');        //Nature de prix de revient
          SetBoolean('WPA_BOOLEAN01', False);       //Pas de Configuration par défaut
          SetBoolean('WPA_BOOLEAN02', True );       //Mémorisation par défaut
          SetBoolean('WPA_BOOLEAN03', True );       //Avec les tarifs achat
          SetBoolean('WPA_BOOLEAN04', False);       //en tenant compte des quantité : Acheté par multiple de
          SetBoolean('WPA_BOOLEAN05', True );       //Calcul tous les niveaux de nomenclature
          SetBoolean('WPA_BOOLEAN06', True );       //en tenant compte des pertes
          SetBoolean('WPA_BOOLEAN07', True);        //en ne tenant pas compte des coùts indirects
          SetBoolean('WPA_BOOLEAN08', False);       //en ne tenant pas compte des flux annexes
          SetBoolean('WPA_BOOLEAN09', False);       //en ne tenant pas compte des marges
        end
      {$ENDIF GPAO}
      {$IFDEF STK}
        else if GetString('WPA_CODEPARAM') = 'WPA_GETSTKPROJETE' then
        begin
          SetBoolean('WPA_BOOLEAN01', True);                        //Analyse consolidée
          SetInteger('WPA_INTEGER01', 5);                           //5 colonnes
          SetString('WPA_VARCHAR01' , '1');                         //exprimées en Journée
          SetInteger('WPA_INTEGER02', 5);                           //5 colonnes
          SetString('WPA_VARCHAR02' , '2');                         //exprimées en Semaine
          SetInteger('WPA_INTEGER03', 3);                           //3 colonnes
          SetString('WPA_VARCHAR03' , '4');                         //exprimées en Mois
        end
      {$ENDIF STK}
    end
    else if ikc = 'M' then
    begin

      if GetString('WPA_CODEPARAM') = 'WPA_BLOCNOTE' then
        SetString('WPA_BLOCNOTE', GetControlText('WPA_BLOCNOTE'))
      else if GetString('WPA_CODEPARAM') = 'WPA_GETSTKPROJETE' then
        SetBoolean('WPA_BOOLEAN01', tCheckBox(GetControl('CB_ANALYSE')).Checked)
    end;

    if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
    begin
      tsBlocNote := TStringList.Create;
      try
        if (GetString('WPA_VARCHAR01')='') then
        begin
          if (GetControlText('GA_COLLECTION' )<>'') then tsBlocNote.Add('GA_COLLECTION=' +GetControlText('GA_COLLECTION' ));
          if (GetControlText('GA_DOMAINE'    )<>'') then tsBlocNote.Add('GA_DOMAINE='    +GetControlText('GA_DOMAINE'    ));
          if (GetControlText('GA_REFCONSTRUC')<>'') then tsBlocNote.Add('GA_REFCONSTRUC='+GetControlText('GA_REFCONSTRUC'));

          if (GetControlText('GA_FAMILLENIV1')<>'') then tsBlocNote.Add('GA_FAMILLENIV1='+GetControlText('GA_FAMILLENIV1'));
          if (GetControlText('GA_FAMILLENIV2')<>'') then tsBlocNote.Add('GA_FAMILLENIV2='+GetControlText('GA_FAMILLENIV2'));
          if (GetControlText('GA_FAMILLENIV3')<>'') then tsBlocNote.Add('GA_FAMILLENIV3='+GetControlText('GA_FAMILLENIV3'));

          if (GetControlText('GA_LIBREART1'  )<>'') then tsBlocNote.Add('GA_LIBREART1='  +GetControlText('GA_LIBREART1'  ));
          if (GetControlText('GA_LIBREART2'  )<>'') then tsBlocNote.Add('GA_LIBREART2='  +GetControlText('GA_LIBREART2'  ));
          if (GetControlText('GA_LIBREART3'  )<>'') then tsBlocNote.Add('GA_LIBREART3='  +GetControlText('GA_LIBREART3'  ));
          if (GetControlText('GA_LIBREART4'  )<>'') then tsBlocNote.Add('GA_LIBREART4='  +GetControlText('GA_LIBREART4'  ));
          if (GetControlText('GA_LIBREART5'  )<>'') then tsBlocNote.Add('GA_LIBREART5='  +GetControlText('GA_LIBREART5'  ));
          if (GetControlText('GA_LIBREART6'  )<>'') then tsBlocNote.Add('GA_LIBREART6='  +GetControlText('GA_LIBREART6'  ));
          if (GetControlText('GA_LIBREART7'  )<>'') then tsBlocNote.Add('GA_LIBREART7='  +GetControlText('GA_LIBREART7'  ));
          if (GetControlText('GA_LIBREART8'  )<>'') then tsBlocNote.Add('GA_LIBREART8='  +GetControlText('GA_LIBREART8'  ));
          if (GetControlText('GA_LIBREART9'  )<>'') then tsBlocNote.Add('GA_LIBREART9='  +GetControlText('GA_LIBREART9'  ));
          if (GetControlText('GA_LIBREARTA'  )<>'') then tsBlocNote.Add('GA_LIBREARTA='  +GetControlText('GA_LIBREARTA'  ));

          if (GetControlText('GA_DATELIBRE1' )<>'') and (StrToDate(GetControlText('GA_DATELIBRE1'  ))<>iDate1900) then tsBlocNote.Add('GA_DATELIBRE1=' +GetControlText('GA_DATELIBRE1' ));
          if (GetControlText('GA_DATELIBRE1_')<>'') and (StrToDate(GetControlText('GA_DATELIBRE1_' ))<>iDate2099) then tsBlocNote.Add('GA_DATELIBRE1_='+GetControlText('GA_DATELIBRE1_'));
          if (GetControlText('GA_DATELIBRE2' )<>'') and (StrToDate(GetControlText('GA_DATELIBRE2'  ))<>iDate1900) then tsBlocNote.Add('GA_DATELIBRE2=' +GetControlText('GA_DATELIBRE2' ));
          if (GetControlText('GA_DATELIBRE2_')<>'') and (StrToDate(GetControlText('GA_DATELIBRE2_' ))<>iDate2099) then tsBlocNote.Add('GA_DATELIBRE2_='+GetControlText('GA_DATELIBRE2_'));
          if (GetControlText('GA_DATELIBRE3' )<>'') and (StrToDate(GetControlText('GA_DATELIBRE3'  ))<>iDate1900) then tsBlocNote.Add('GA_DATELIBRE3=' +GetControlText('GA_DATELIBRE3' ));
          if (GetControlText('GA_DATELIBRE3_')<>'') and (StrToDate(GetControlText('GA_DATELIBRE3_' ))<>iDate2099) then tsBlocNote.Add('GA_DATELIBRE3_='+GetControlText('GA_DATELIBRE3_'));

          if (Valeur(GetControlText('GA_VALLIBRE1' ))<>0.0) then tsBlocNote.Add('GA_VALLIBRE1=' +Trim(StrFPoint(Valeur(GetControlText('GA_VALLIBRE1' )))));
          if (Valeur(GetControlText('GA_VALLIBRE1_'))<>0.0) then tsBlocNote.Add('GA_VALLIBRE1_='+Trim(StrFPoint(Valeur(GetControlText('GA_VALLIBRE1_')))));
          if (Valeur(GetControlText('GA_VALLIBRE2' ))<>0.0) then tsBlocNote.Add('GA_VALLIBRE2=' +Trim(StrFPoint(Valeur(GetControlText('GA_VALLIBRE2' )))));
          if (Valeur(GetControlText('GA_VALLIBRE2_'))<>0.0) then tsBlocNote.Add('GA_VALLIBRE2_='+Trim(StrFPoint(Valeur(GetControlText('GA_VALLIBRE2_')))));
          if (Valeur(GetControlText('GA_VALLIBRE3' ))<>0.0) then tsBlocNote.Add('GA_VALLIBRE3=' +Trim(StrFPoint(Valeur(GetControlText('GA_VALLIBRE3' )))));
          if (Valeur(GetControlText('GA_VALLIBRE3_'))<>0.0) then tsBlocNote.Add('GA_VALLIBRE3_='+Trim(StrFPoint(Valeur(GetControlText('GA_VALLIBRE3_')))));

          if assigned(GetControl('GA_FERME'      )) and (tCheckBox(GetControl('GA_FERME'      )).State <>cbGrayed) then tsBlocNote.Add('GA_FERME='     +GetControlText('GA_FERME')     );
          if assigned(GetControl('GA_BOOLLIBRE1' )) and (tCheckBox(GetControl('GA_BOOLLIBRE1' )).State <>cbGrayed) then tsBlocNote.Add('GA_BOOLLIBRE1='+GetControlText('GA_BOOLLIBRE1'));
          if assigned(GetControl('GA_BOOLLIBRE2' )) and (tCheckBox(GetControl('GA_BOOLLIBRE2' )).State <>cbGrayed) then tsBlocNote.Add('GA_BOOLLIBRE2='+GetControlText('GA_BOOLLIBRE2'));
          if assigned(GetControl('GA_BOOLLIBRE3' )) and (tCheckBox(GetControl('GA_BOOLLIBRE3' )).State <>cbGrayed) then tsBlocNote.Add('GA_BOOLLIBRE3='+GetControlText('GA_BOOLLIBRE3'));
        end
        else
        begin
          if (not assigned(ecran)) then  // cas d'appel de l'acran de paramétrage depuis les affaires
          begin
            SetBoolean('WPA_BOOLEAN01', True);
            ControlField('WPA_BOOLEAN01');
          end;
        end;
      finally
        SetString('WPA_BLOCNOTE', tsBlocNote.Text);
        tsBlocNote.Free;
      end;
    end;
  finally
    EnableControl;
  end;

  if RecordIsValid then
  begin
    {$IFDEF GPAOLIGHT}
    {$IFNDEF ERADIO}
    {$IFNDEF EAGLSERVER}
      if EcranIs('WPA_WPLANLIVR') then
      begin
        if Assigned(GetControl('CADRAGEITEM')) then
          SetInteger('WPA_BOOLEAN05', THValComboBox(GetControl('CADRAGEITEM')).ItemIndex);
        PlanLivrSetIndicateurs;
        PlanLivrRecolleLeMemo;
      end;
    {$ENDIF !EAGLSERVER}
    {$ENDIF ERADIO}
    {$ENDIF GPAOLIGHT}
  end;

  Inherited ;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnAfterUpdateRecord ;
{$IFNDEF EAGLSERVER}
var
  Prefixe: String;
  Identifiant: Integer;
{$ENDIF !EAGLSERVER}
begin
  Inherited ;
  if Assigned(Ecran) then
  begin
    {$IFNDEF EAGLSERVER}
      tfFiche(Ecran).Retour := wTrue;
      if (Pos(sChangementComposant, GetString('WPA_CODEPARAM')) = 1) and EcranIs('WPA_CHGCOMPOSANT') then
      begin
        TobWPA.LoadDB(True);
      end
      else if EcranIs('WPA_CODEARTICLE') then
      begin
        TFFiche(Ecran).Retour := GetString('WPA_VARCHAR01');
        Prefixe := GetArgumentValue(StArgument, 'PREFIXE');
        Identifiant := GetArgumentInteger(StArgument, 'IDENTIFIANT');
        if (Prefixe <> '') and (Identifiant <> 0) then
      end;
      {$IFDEF EAGLCLIENT}
        EnabledOnLoadRecord := false;
      {$ENDIF}
    {$ENDIF !EAGLSERVER}
  end;
end;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnLoadRecord ;

{$IFNDEF EAGLSERVER}
{$IFDEF GPAO}
  var
    { Remplacement de composants }
    iCpt, iCol, iRow : Integer;
    bCancel    : Boolean;
    DateApp    : tDateTime;
    sRevision  : String;
    tsBlocNote : tStringList;
{$ENDIF GPAO}
{$ENDIF !EAGLSERVER}
begin
  {$IFNDEF EAGLSERVER}
    {$IFDEF EAGLCLIENT}
      if not EnabledOnLoadRecord then exit;
    {$ENDIF EAGLCLIENT}
  {$ENDIF !EAGLSERVER}

  Inherited ;

{$IFNDEF EAGLSERVER}
  { Titre }
  if not Title then
  begin
    if (GetString('WPA_LONGVARCHAR01') <> '') then
      Ecran.Caption := GetString('WPA_LONGVARCHAR01')
  end
  else
  begin
    if copy(GetArgumentValue(StArgument, 'TITLE'),1,4)='WPA_' then
      Ecran.Caption := GetString(GetArgumentValue(StArgument, 'TITLE'))
    else
      Ecran.Caption := GetArgumentString(StArgument, 'TITLE', False);
  end;
  UpdateCaption(ecran);

  {$IFDEF GPAO}
  if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
  begin
    if (sAppel<>'AFFAIRE') then // Pour l'initialisation du budget des affaires le titre est géré par la variable TITLE
    begin
      if (GetString('WPA_CODEPARAM')='PRIXDEREVIENT') then
        Ecran.Caption := TraduireMemoire('Calcul des prix de revient')
      else
        Ecran.Caption := TraduireMemoire('Paramètres de lancement du calcul des prix de revient par batch');
      UpdateCaption(ecran);
    end;

    LBARTICLE               := thLabel(GetControl('LBARTICLE'));
    LBITINERAIRE            := thLabel(GetControl('LBITINERAIRE'));
    LBCIRCUIT               := thLabel(GetControl('LBCIRCUIT'));
    LBNOMENCLATURELIBELLE   := thLabel(GetControl('LBNOMENCLATURELIBELLE'));
    LBNOMENCLATUREREVISION  := thLabel(GetControl('LBNOMENCLATUREREVISION'));
    LBNOMENCLATUREMAJEUR    := thLabel(GetControl('LBNOMENCLATUREMAJEUR'));
    LBNOMENCLATUREMINEUR    := thLabel(GetControl('LBNOMENCLATUREMINEUR'));
    LBGAMMELIBELLE          := thLabel(GetControl('LBGAMMELIBELLE'));
    LBGAMMEREVISION         := thLabel(GetControl('LBGAMMEREVISION'));
    LBGAMMEMAJEUR           := thLabel(GetControl('LBGAMMEMAJEUR'));
    LBGAMMEMINEUR           := thLabel(GetControl('LBGAMMEMINEUR'));
    LBUNITEPROD1            := thLabel(GetControl('LBUNITEPROD1'));
    LBUNITEPROD2            := thLabel(GetControl('LBUNITEPROD2'));
    LBUNITEPROD3            := thLabel(GetControl('LBUNITEPROD3'));
    WPA_COMBO01             := thValComboBox(GetControl('WPA_COMBO01'));
    WPA_COMBO02             := thValComboBox(GetControl('WPA_COMBO02'));
    WPA_BOOLEAN01           := tDbCheckBox(GetControl('WPA_BOOLEAN01'));
    RGPRIX                  := thRadioGroup(GetControl('WPA_COMBO04'));
    RGUNPDRPAR              := thRadioGroup(GetControl('WPA_COMBO05'));
    BTCalcul                := tToolBarButton97(GetControl('BTCALCUL'));

    SetControlProperty('BTCALCUL', 'Visible', (GetString('WPA_CODEPARAM') =  sPrixDeRevient));
    SetControlProperty('BVALIDER', 'Visible', (GetString('WPA_CODEPARAM') <> sPrixDeRevient));

    if Assigned(BTCalcul) then
      BTCalcul.OnClick := BTCalcul_OnClick;
    if Assigned(WPA_BOOLEAN01) then
      WPA_BOOLEAN01.OnClick := WPA_BOOLEAN01_OnClick;

    {$IFNDEF ERADIO}
      if Assigned(GetControl('WPA_VARCHAR05')) then
        thdbedit(GetControl('WPA_VARCHAR05')).OnElipsisClick := WPA_VARCHAR05_OnElipsisClick;
      if Assigned(GetControl('WPA_VARCHAR06')) then
        thdbedit(GetControl('WPA_VARCHAR06')).OnElipsisClick := WPA_VARCHAR06_OnElipsisClick;
    {$ENDIF !ERADIO}

    //A supprimer dès que les développements seront terminés et validés
    if Assigned(GetControl('WPA_BOOLEAN09')) then SetControlEnabled('WPA_BOOLEAN09', V_PGI.SAV);

    l1erClickBoolean01 := True;

    tsBlocNote := TStringList.Create;
    try
      if (GetString('WPA_VARCHAR01')='') then
      begin
        tsBlocNote.Text := GetString('WPA_BLOCNOTE');
        for iCpt := 0 to tsBlocNote.Count-1 do
        begin
          SetControlText(tsBlocNote.Names[iCpt] , tsBlocNote.Values[tsBlocNote.Names[iCpt]]);
        end;
      end;
    finally
      tsBlocNote.Free;
    end;
    if (Assigned(GetControl('WPA_VARCHAR07'))) and (sAppel <> 'AFFAIRE') then
    begin
      SetControlProperty('WPA_VARCHAR07', 'PLUS', 'WRT_TYPEPDR NOT LIKE "OR%" AND WRT_TYPEPDR<>"AFF"');
      if Getstring('WPA_VARCHAR07')='AFF' then
      begin
        if not(DS.State in [dsInsert,dsEdit]) then
          DS.Edit;
        SetString('WPA_VARCHAR07', 'ACT');
      end;
    end;
    // Code affaire
    if (Assigned(GetControl('WPA_VARCHAR09'))) and (sAppel <> 'AFFAIRE') and (GetString('WPA_VARCHAR09')<>'') then
    begin
      if not(DS.State in [dsInsert,dsEdit]) then
        DS.Edit;
      SetString('WPA_VARCHAR09', '');
    end;
	{"Avec marges"}
		if Assigned(GetControl('WPA_BOOLEAN09')) then
  		SetControlProperty('WPA_BOOLEAN09', 'ENABLED', False);
    if not GereFraisAvances then
    begin
      {"Flux annexes"}
      if (Assigned(GetControl('WPA_BOOLEAN08'))) and GetBoolean('WPA_BOOLEAN08') then
      begin
        if not(DS.State in [dsInsert,dsEdit]) then
          DS.Edit;
        SetBoolean('WPA_BOOLEAN08', False);
      end;
      SetControlProperty('WPA_BOOLEAN08'	, 'ENABLED', False);
    end;
    SetControlEnabled('WPA_COMBO04', GetBoolean('WPA_BOOLEAN12') or GetBoolean('WPA_BOOLEAN13'));

    { Calcul du budget d'affaire : l'article n'est pas modifiable }
    SetControlEnabled('WPA_VARCHAR01' , sTypePDR <> 'AFF');
    SetControlEnabled('WPA_VARCHAR07' , sTypePDR <> 'AFF');
    SetControlVisible('WPA_VARCHAR09' , sTypePDR = 'AFF');
    SetControlVisible('TWPA_VARCHAR09', sTypePDR = 'AFF');
    SetControlEnabled('WPA_VARCHAR09' , False);
  end;
  {$ENDIF GPAO}

  {$IFDEF GPAOLIGHT}
  if EcranIs('WPA_ORDRERECEP') then
  begin
    { Titre}
    Ecran.Caption := 'Réception partielle  -  Ordre ' + IntToStr(LigneOrdre);

  {  Unités:
        > WPA_COMBO01 : unité de saisie
        > WPA_COMBO02 : unité de stock

      Quantité à réceptionner:
        > WPA_DOUBLE01
        > WPA_DOUBLE05

       Quantité déjà réceptionnée:
        > WPA_DOUBLE06
          > WPA_DOUBLE02

       Quantité réceptionnée:
        > WPA_DOUBLE03
          > WPA_DOUBLE07

       Total:
        > WPA_DOUBLE04
          > WPA_DOUBLE08

        Coef:
        > WPA_DOUBLE09
    }

    { Calcul du reste à réceptionner }
    DS.Edit;
    SetDouble('WPA_DOUBLE03', max(0,GetDouble('WPA_DOUBLE01') - GetDouble('WPA_DOUBLE02')));

    { Calcul du total }
    SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') + GetDouble('WPA_DOUBLE03'));

    { Conversion }
    SetConversion;
  end
  else if EcranIs('WPA_ORDRELANC') then
  begin
    { Titre}
    Ecran.Caption := 'Lancement partiel  -  Phase : ' + OpeCirc + ' de l''ordre ' + IntToStr(LigneOrdre);

  {  Unités:
        > WPA_COMBO01 : unité de saisie
        > WPA_COMBO02 : unité de stock

      Quantité à lancer:
        > WPA_DOUBLE01
        > WPA_DOUBLE05

       Quantité déjà lancée:
        > WPA_DOUBLE06
          > WPA_DOUBLE02

       Quantité lancement:
        > WPA_DOUBLE03
          > WPA_DOUBLE07

       Total:
        > WPA_DOUBLE04
          > WPA_DOUBLE08

        Coef:
        > WPA_DOUBLE09
    }

    { Calcul du reste à lancer }
    DS.Edit;
    SetDouble('WPA_DOUBLE03', max(0,GetDouble('WPA_DOUBLE01') - GetDouble('WPA_DOUBLE02')));

    { Calcul du total }
    SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') + GetDouble('WPA_DOUBLE03'));

    { Conversion }
    SetConversion;
  end
  else if EcranIs('WPA_ORDREINTEGRAT') or EcranIs('WPA_ORDRECPDCOR') then
  begin
    { Titre}
    if EcranIs('WPA_ORDREINTEGRAT') then
      Ecran.Caption := TraduireMemoire('Réintégration') + ' : '+CodeComposant
    else
      Ecran.Caption := TraduireMemoire('Correction de réception')+' : ' + CodeComposant;
  {  Unités:
        > WPA_COMBO01 : unité de saisie
        > WPA_COMBO02 : unité de stock

      Quantité de besoin:
        > WPA_DOUBLE01
        > WPA_DOUBEL05

       Quantité déjà servi:
        > WPA_DOUBLE02
        > WPA_DOUBLE06

       Quantité réintégré:
        > WPA_DOUBLE03
          > WPA_DOUBLE07

       Total:
        > WPA_DOUBLE04
          > WPA_DOUBLE08
    }

    { Calcul du reste à réceptionner }
    DS.Edit;
    SetDouble('WPA_DOUBLE03', GetDouble('WPA_DOUBLE02'));

    { Calcul du total }
    SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') - GetDouble('WPA_DOUBLE03'));

    { Conversion }
    SetConversion;
  end
  else if EcranIs('WPA_ORDRESERVICE') or EcranIs('WPA_ORDRECPDREC') then
  begin
    { Titre}
    if EcranIs('WPA_ORDRESERVICE') then
      Ecran.Caption := TraduireMemoire('Service manuel')+' : ' + CodeComposant
    else
      Ecran.Caption := TraduireMemoire('Réception manuelle')+' : ' + CodeComposant;
  {  Unités:
        > WPA_COMBO01 : unité de saisie
        > WPA_COMBO02 : unité de stock

      Quantité de besoin:
        > WPA_DOUBLE01
        > WPA_DOUBEL05

       Quantité déjà servi:
        > WPA_DOUBLE02
        > WPA_DOUBLE06

       Quantité servi:
        > WPA_DOUBLE03
          > WPA_DOUBLE07

       Total:
        > WPA_DOUBLE04
          > WPA_DOUBLE08
    }

    { Calcul du reste à réceptionner }
    DS.Edit;

    { Calcul du total }
    SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') + GetDouble('WPA_DOUBLE03'));

    { Conversion }
    SetConversion;
  end
  else if EcranIs('WPA_ORDREREBUT') then
  begin
    { Titre}
    Ecran.Caption := 'Saisie de rebuts  -  Phase ' + OpeCirc + ' de l''ordre ' + IntToStr(LigneOrdre);

    { Init. de la qté à rebuter }
    DS.Edit;
    SetDouble('WPA_DOUBLE03', 0);
  end
  else if EcranIs('WPA_CODEARTICLE') then
  begin
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.edit;
    { Filtre sur les natures de travail }
    if Assigned(GetControl('WPA_VARCHAR01')) then
      // {GETPLUSWARTNAT} SetControlProperty('WPA_VARCHAR01', 'PLUS', wGetPlusArtNat(NatureTravail));
      SetString('WPA_VARCHAR01', CodeArticle);
  end;
  {$ENDIF GPAOLIGHT}

  {$IFDEF GPAO}
  if EcranIs('WPA_CHGVERSION') then
  begin
    if Assigned(GetControl('WPA_VARCHAR02')) then
    begin
      DS.Edit;

      DisableControl;
      SetString('WPA_VARCHAR02', '');
      linitData := True;
      EnableControl;
{      SetString('WPA_VARCHAR02', GetString('WPA_VARCHAR01'));
      if wGetVersion(Article) <> '' then
        SetString('WPA_VARCHAR02', wGetVersion(Article))
      else
        if not wExistUneVersion(GetCleWVS.article) then SetString('WPA_VARCHAR02', '');
}
      { filtre sur les versions de l'article }
      SetControlProperty('WPA_VARCHAR02', 'PLUS', wGetPlusVersion(Article) );
    end;
  end
  else if EcranIs('WPA_ARTNAT') then
  begin
    { Filtre sur les natures de travail }
    DS.Edit;
    if Assigned(GetControl('WPA_VARCHAR01')) then
    begin
      // {GETPLUSWARTNAT} SetControlProperty('WPA_VARCHAR01', 'PLUS', wGetPlusArtNat(NatureTravail));
      SetString('WPA_VARCHAR01', CodeArticle);
      SetString('WPA_VARCHAR02', NatureTravail);
    end;
  end;
  if EcranIs('WPA_FAISABILITE') then
  begin
    { Filtre sur les nomenclatures }
    SetControlProperty(wFaisabilite.Wpa_CodeArticleWNT, 'PLUS', wGetPlusCodeArticleWNT(GetString(wFaisabilite.Wpa_NatureTravail),GetString(wFaisabilite.Wpa_Coditi)));
    { Filtre sur les indices majeur }
    SetControlProperty(wFaisabilite.Wpa_MajeurWNT, 'PLUS', wGetPlusMajeurWNT(GetString(wFaisabilite.Wpa_NatureTravail), wgetArticleFromCodeArticle(GetString(wFaisabilite.Wpa_CodeArticleWNT))));
    { Filtre sur les circuits }
    SetControlProperty(wFaisabilite.Wpa_Circuit, 'PLUS', 'QCI_CODITI="' + GetString(wFaisabilite.Wpa_CodIti) + '"');

    DS.Edit;
    SetDateTime('WPA_DATE01', V_Pgi.DateEntree);
  end;
  if EcranIs('WPA_CHGCOMPOSANT') then
  begin
    { TobWPA }
    TobWPA := Tob.Create('WPARAM', nil, -1);
    TobWPA.SetString('WPA_CODEPARAM', sChangementComposant);
    TobWPA.SetString('WPA_UTILISATEUR', V_PGI.User);
    DS.Edit;
    TobWPA.LoadDB(True);
    SetString(chgNATURETRAVAIL, wGetFieldFromWPA(sChangementComposant, chgNATURETRAVAIL));
    SetString(chgARTICLE,       wGetFieldFromWPA(sChangementComposant, chgARTICLE));
    SetString(chgFAMILLE1,      wGetFieldFromWPA(sChangementComposant, chgFAMILLE1));
    SetString(chgFAMILLE2,      wGetFieldFromWPA(sChangementComposant, chgFAMILLE2));
    SetString(chgFAMILLE3,      wGetFieldFromWPA(sChangementComposant, chgFAMILLE3));
    sRevision := wGetFieldFromWPA(sChangementComposant, REVISION);
    SetString(REVISION, iifV(sRevision = '', 'MINVAL', sRevision));
    DateApp := Date;
    SetDateTime(DATEAPPLICATION, DateApp);
    if Assigned(GetControl('RB' + GetString(REVISION))) then
      tRadioButton(GetControl('RB' + GetString(REVISION))).Checked := True;

    { Init d'au moins 1 fils }
    TobWPA_Fils := Tob.Create('CHGCOMPOSANT', TobWPA, -1);
    TobWPA_Fils.AddChampSupValeur('OLDCOMPOSANT', '');
    TobWPA_Fils.AddChampSupValeur('NEWCOMPOSANT', '');
    TobWPA_Fils.AddChampSupValeur('TYPECHGLIEN', 'RIEN');
    TobWPA_Fils.AddChampSupValeur('CBQLIEN', wTrue);
    TobWPA_Fils.AddChampSupValeur('CBPERTEPROP', wFalse);
    TobWPA_Fils.AddChampSupValeur('CBQPPER', wFalse);
    TobWPA_Fils.AddChampSupValeur('CBQPERIODE', wFalse);
    TobWPA_Fils.AddChampSupValeur('CBQPFIXE', wFalse);
    wInitLienFields(TobWPA_Fils);

    FromPutEcran := True;
    FromApplyChangement := False;
    FromBtMoulinette := False;
    TobWPA_Fils.PutEcran(Ecran, tWinControl(GetControl('GBQLIEN')));
    FromPutEcran := False;
    { divers }
    iCol := 0;
    iRow := 1;
    bCancel := False;
    GRILLE_OnCellEnter(Self, iCol, iRow, bCancel);
    SetFocusControl(chgNATURETRAVAIL);
    SetControlEnabled('BTANNULERFILS', False);
  end;
  {$ENDIF GPAO}

  {$IFDEF GPAOLIGHT}
  if EcranIs('WPA_GETCODEORDRE') then
  begin
    { Titre}
    Ecran.Caption := 'Choix d''un regroupement';
    if Assigned(GetControl('WPA_INTEGER01')) then
    begin
      { filtre sur les regroupements autorisés }
      SetControlProperty('WPA_INTEGER01', 'PLUS', '(WOT_TYPEORDRE= "SIM" or WOT_TYPEORDRE = "MUL") AND WOT_NATURETRAVAIL= "' + NatureTravail + '"');
    end;
  end
  else if EcranIs('WPA_FAISABILITE') then
  begin
    ShowCodeNomenAss; { Affichage code nomenclature assemblage }
  end
  ;
  {$ENDIF GPAOLIGHT}

  if (copy(GetString('WPA_CODEPARAM'), 1, 10) = sTarifsRecherche) then
  begin

    if (sFonctionnalite<>'') then
    begin
      if       (sAppel='RECHERCHE') then
      begin
        Ecran.Caption     := RechDom('YTARIFSFONCTIONNALITE',sFonctionnalite,False)+' : recherche et justification';
        if sCodePort <> '' then
          Ecran.Caption     := Ecran.Caption + ' pour le frais '+RechDom('GCPORT', sCodePort, False);
        if sTiersFrais <> '' then
          Ecran.Caption     := Ecran.Caption + ' pour le tiers '+sTiersFrais;
      end
      else if  (Pos(sAppel,'PIECE/LIGNE')>0)     then
        Ecran.Caption     := RechDom('YTARIFSFONCTIONNALITE',sFonctionnalite,False)+' : justification '
      else
        Ecran.Caption     := RechDom('YTARIFSFONCTIONNALITE',sFonctionnalite,False)+' Problème de paramétrage, veuillez contacter le S.A.T.';
	    UpdateCaption(ecran);
    end;

    SetSayOuGet('WPA_COMBO01'  , True);  //Fonctionnalite  : Non modifiable
    SetSayOuGet('WPA_VARCHAR15', True);  //Frais  : Non modifiable

    GAFFICHAGE := thGrid(GetControl('GAFFICHAGE'));
    if (GAFFICHAGE<>nil)    then thGrid(GetControl('GAFFICHAGE')).PostDrawCell := GAffichageCell;
    if (GAFFICHAGE<>nil)    then thGrid(GetControl('GAFFICHAGE')).OnRowEnter := GAffichageOnRowEnter;

    LBTIERS     := thLabel(GetControl('LBTIERS'));
    LBLIVRE     := thLabel(GetControl('LBLIVRE'));
    LBARTICLE   := thLabel(GetControl('LBARTICLE'));
    LBAFFAIRE   := thLabel(GetControl('LBAFFAIRE'));
    LBATELIER   := thLabel(GetControl('LBATELIER'));
    ClientOuFournisseur;
    CtrlAffichageMarqueChoixQualiteIndice;
    ControleAffichageDimensionsArticle;
    
    BTRechercheTarifaire  := tToolBarButton97(GetControl('BTRECHERCHETARIFAIRE'));
    if (BTRechercheTarifaire  <> nil)  then BTRechercheTarifaire.OnClick  := BTRechercheTarifaire_OnClick;
    BTRechercheTarifaire.Visible := (sAppel='RECHERCHE');

    GBSousTraitance := tGroupBox(GetControl('GBSOUSTRAITANCE'));
    ParticularitesSousTraitance;

    MBLOCNOTE         := tMemo(GetControl('MBLOCNOTE'));

    TobTarif          := Tob.Create('_TobTarif_'        ,nil,-1);
    TobPrixMarge      := Tob.Create('_TobPrixMarge_'      ,nil,-1);
    TobAffichage      := Tob.Create('_TobAffichage_'      ,nil,-1);
    TobContexte       := Tob.Create('_TobContexte_'       ,nil,-1);
    TobUneLigneTarif  := Tob.Create('_TobUneLigneTarif_'  ,nil,-1);
    if (Pos(sAppel,'PIECE/LIGNE')>0) then
    begin
      { Récupère les tobs }
      TobContexte.Dupliquer( Tob(TransfertTobs[0]), True, True);
      TobUneLigneTarif.Dupliquer( Tob(TransfertTobs[1]), True, True);;
      { Masque l'onglet de saisie des paramètres }
      SetControlVisible('PGENERAL'      , False);
      SetControlVisible('TSCOMPLEMENT1' , False);
      SetControlVisible('TSCOMPLEMENT2' , False);
      { Lance la purée }
      BTRechercheTarifaire_OnClick(GetControl('BVALIDER'));
    end;
    {$IFNDEF ERADIO}
      if Assigned(GetControl('WPA_VARCHAR02') ) then
        thDBEdit(GetControl('WPA_VARCHAR02')).OnElipsisClick := Tiers_OnElipsisClick;
      if Assigned(GetControl('WPA_VARCHAR06') ) then
        thDBEdit(GetControl('WPA_VARCHAR06')).OnElipsisClick := TiersLivre_OnElipsisClick;
    {$ENDIF !ERADIO}
    { Champ n° d'adresse}
    SetControlEnabled ('WPA_VARCHAR13', (GetString('WPA_VARCHAR06')<>'') and (LookupValueExist(THEdit(Getcontrol('WPA_VARCHAR06')))) );
    SetControlProperty('WPA_VARCHAR13', 'Plus', 'ADR_TYPEADRESSE="TIE"'
                                             + ' AND ADR_REFCODE="' + GetControlText('WPA_VARCHAR06') + '"'
                                             + ' AND ADR_LIVR="X"'
                                             );
    {Restriction par utilisateur}
    if FieldExists('WPA_COMBO02') and (sAppel='RECHERCHE') then
    begin
      if not(DS.State in [dsInsert,dsEdit]) then
        DS.Edit;
      SetString('WPA_COMBO02', VH^.ProfilUserC[PrEtablissement].Depot);
      if Assigned(Ecran) and Assigned(GetControl('WPA_COMBO02')) and tWinControl(GetControl('WPA_COMBO02')).Enabled then
        SetControlEnabled('WPA_COMBO02', not VH^.ProfilUserC[PrEtablissement].ForceDepot);
    end;
  end
  else if EcranIs('WPA_NOMPARC') or EcranIs('WPA_ARTNSER') then
  begin
    thDBEdit(GetControl('WPA_VARCHAR01')).Plus := 'GA_GEREPARC=' + wTrue;
  end
  else if EcranIs('WPA_CBNTRANSUSER') then
  begin
    { Gestion du changement d'utilisateur sur les propositions validées (CBN) }
    tfFiche(Ecran).Retour:=wFalse;
    SetControlProperty('WPA_COMBO01', 'Enabled', V_PGI.Superviseur);
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.Edit;
  end
  else if (GetString('WPA_CODEPARAM') = 'WPA_GETDATE'  )
       or (GetString('WPA_CODEPARAM') = 'WPA_DEPOT'    )
       or (GetString('WPA_CODEPARAM') = 'WPA_GETSTRING') then
  begin
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.Edit;
  end
  else if EcranIs('WPA_GETSTKPROJETE') then
  begin
    DS.Edit;
    tCheckBox(GetControl('CB_ANALYSE')).Checked := StrToBool_ (wGetFieldFromWPA('WPA_GETSTKPROJETE', 'WPA_BOOLEAN01'));
    SetControlProperty('GBANALYSECONSOLIDE', 'Visible', tCheckBox(GetControl('CB_ANALYSE')).Checked);
    SetControlProperty('TWPA_VARCHAR01'    , 'Visible', tCheckBox(GetControl('CB_ANALYSE')).Checked);
    SetControlProperty('TWPA_TOTAL'        , 'Visible', tCheckBox(GetControl('CB_ANALYSE')).Checked);
  end
{$IFDEF QUALITE}
  else if EcranIs('RQGETCLOTURERQN') then
  begin
    DS.Edit;
    SetString('WPA_BLOCNOTE' , '');
  end
{$ENDIF QUALITE}
  {$IFDEF GPAO}
    {$IFDEF AFFAIRE}
      else if EcranIs('WPA_GETAFCONTEXTE') then
      begin
        SetControlEnabled('WPA_DATE01', GetString('WPA_VARCHAR01')<>sContexteORIGINEWLB );
        if GetString('WPA_VARCHAR01')= sContexteORIGINEWLB then
        begin
          DS.Edit;
          SetDateTime('WPA_DATE01' , V_PGI.DateEntree);
        end;
      end
    {$ENDIF AFFAIRE}
  {$ENDIF GPAO}
  ;

  if EcranIs('WPA_ORDREREBUT') or EcranIs('WPA_ORDRERECEP') or EcranIs('WPA_ORDRELANC')
  or EcranIs('WPA_ORDRESERVICE') or EcranIs('WPA_ORDRECPDREC') or EcranIs('WPA_ORDREINTEGRAT')
  or EcranIs('WPA_ORDRECPDCOR') then
    SetFocusControl('WPA_DOUBLE03');

  {$IFDEF GPAOLIGHT}
  {$IFNDEF ERADIO}
    if EcranIs('WPA_WPLANLIVR') then
      SetPlanLivrOnLoad
    {$IFDEF GPAO}
    else if EcranIs('WPA_PDCPARAM') then
      SetPlanDeChargeOnLoad
    {$ENDIF GPAO}
    ;
  {$ENDIF ERADIO}
  {$ENDIF GPAOLIGHT}
{$ENDIF !EAGLSERVER}
end;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  if EcranIs('WPA_NOMPARC') then
  begin
    if (f.FieldName='WPA_VARCHAR01') and (trim(getControltext(f.FieldName)) <> '') then
      SetControlProperty('WPA_VARCHAR02', 'PLUS', 'WPC_CODEARTICLE="'+getControltext(f.FieldName)+'"')
  end;
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnArgument ( S: String ) ;
var
  iRangPanel, iHauteurDisponible : integer;
  {$IFDEF GPAO}
  {$IFNDEF ERADIO}
  {$IFNDEF EAGLSERVER}
    Maillage, TypePlanning : string;
  {$ENDIF !EAGLSERVER}
  {$ENDIF !ERADIO}
  {$ENDIF GPAO}

  {-----------------------------------------------------------------------------
  ------------------------------------------------------------------------------}
  procedure ActivePanel(sNomPanel: string; var iRangPanel: integer; var iHauteurDisponible : integer; iHauteurPanel : integer);
  begin {ActivePanel}
    //Tabsheet TSCaractérisitiques complet on active au Tabsheet TSComplements
    if (iHauteurDisponible < iHauteurPanel) then
    begin
      if (not tTabSheet(GetControl('TSCOMPLEMENT1')).TabVisible) then
      begin
        iHauteurDisponible := tTabSheet(GetControl('TSCOMPLEMENT1')).Height; //Hauteur du TabSheet
        iRangPanel         :=  10; //Premier Panel du Tabsheet complément 1
        tTabSheet(GetControl('TSCOMPLEMENT1')).TabVisible := True;
      end
      else
      begin
        iHauteurDisponible := tTabSheet(GetControl('TSCOMPLEMENT1')).Height; //Hauteur du TabSheet
        iRangPanel         :=  19; //Premier Panel du Tabsheet complément 2
        tTabSheet(GetControl('TSCOMPLEMENT2')).TabVisible := True;
      end;
    end;

    thPanel(GetControl(sNomPanel)).Parent                 := tHPanel(GetControl('PN'+IntToStr(iRangPanel)));
    thPanel(GetControl('PN'+IntToStr(iRangPanel))).Height := iHauteurPanel;

    iRangPanel := iRangPanel+1;
    iHauteurDisponible := iHauteurDisponible - iHauteurPanel;
  end; {ActivePanel}
//GP_20071011_PCO_GP14299
  procedure SetLibelleLibre;
  var
    iLibre  : integer;
    Libelle : string;
  begin
    For  iLibre:=1 to 10 do
    begin
      GCTitreZoneLibre( IIF(iLibre = 10,'GA_LIBREARTA','GA_LIBREART' + IntToStr(iLibre)), Libelle, False);
      if Libelle <> ''
        then SetControlCaption('TWPA_LIBREWPA' + IntToStr(iLibre), Libelle)
      else
      begin
        SetControlVisible('TWPA_LIBREWPA' + IntToStr(iLibre), False);
        if iLibre < 9 then
          SetControlVisible('WPA_LONGVARCHAR0' + IntToStr(iLibre), False)
        else
          SetControlVisible(IIF(iLibre = 10,'WPA_VARCHAR','WPA_VARCHAR0') + IntToStr(iLibre), False) ;
      end;
    end ;
  end ;
begin
  {$IFDEF EAGLCLIENT}
    EnabledOnLoadRecord := true;
  {$ENDIF}
  { Identifiant WPE}
  iIdentifiantWPE:=0;

  { Gestion d'un titre en particulier }
  Title := Pos('TITLE', s) > 0;
  Inherited ;

  { Récup. des valeurs de la phase }
  OpeCirc       := GetArgumentValue(S, 'WOP_OPECIRC');
  CodeComposant := GetArgumentValue(S, 'WOB_CODECOMPOSANT');
  LigneOrdre    := ValeurI(GetArgumentValue(S, 'WOP_LIGNEORDRE'));
  WithCancelInfo := GetArgumentBoolean(S, 'WITHCANCELINFO');

  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      { événements }
      if Assigned(GetControl('BFERME')) then
        tToolBarButton97(GetControl('BFERME')).OnClick := bFerme_Click;
	    if Assigned(GetControl('BVALIDER')) then
	      TToolBarButton97(GetControl('BVALIDER')).OnClick := BValiderOnClick;

      {$IFDEF GPAOLIGHT}
        if Assigned(GetControl('BTCALCULFAISABILITE')) then
          tToolBarButton97(GetControl('BTCALCULFAISABILITE')).OnClick := BTCalculFaisabilite_OnClick;
      {$ENDIF GPAOLIGHT }
      {$IFDEF GPAO}
        {$IFDEF AFFAIRE}
          if Assigned(GetControl('WPA_VARCHAR01')) then
            thDbValComboBox(GetControl('WPA_VARCHAR01')).OnChange := WPA_VARCHAR01_OnChange;
    	    if Assigned(GetControl('BRECHAFFAIRE')) then
	          TToolBarButton97(GetControl('BRECHAFFAIRE')).OnClick := BRECHAFFAIRE_OnCLick ;
        {$ENDIF AFFAIRE}
      {$ENDIF GPAO}
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}

  if EcranIs('WPA_CODEARTICLE') then
  begin
    CodeArticle  := GetArgumentValue(S, 'CODEARTICLE');
  end
  else if EcranIs('WPA_CHGVERSION') then
  begin
    Article  := GetArgumentValue(S, 'ARTICLE');
  end
  else if EcranIs('WPA_ARTNAT') then
  begin
    CodeArticle  := GetArgumentValue(S, 'CODEARTICLE');
    NatureTravail:= GetArgumentValue(S, 'NATURETRAVAIL');
  end
  else if EcranIs('WPA_CHGCOMPOSANT') then {remplacement de composants}
  begin
    { composants }
    if Assigned(GetControl('GRILLE')) then GRILLE := thGrid(GetControl('GRILLE'));
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        {$IFDEF GPAO}
          if Assigned(GetControl('GRILLE'))       then GRILLE.OnElipsisClick := Cell_OnElipsisClick;
          if Assigned(GetControl('GRILLE'))       then GRILLE.OnKeyDown := GRILLE_OnKeyDown;
          if Assigned(GetControl('GRILLE'))       then GRILLE.OnDblClick := GRILLE_OnDblClick;
          if Assigned(GetControl('BTMOULINETTE')) then tToolBarButton97(GetControl('BTMOULINETTE')).OnClick       := btMoulinette_OnClick;
        {$ENDIF GPAO}
      {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
    if Assigned(GetControl('GRILLE')) then GRILLE.OnKeyPress := GRILLE_OnKeyPress;
    if Assigned(GetControl('GRILLE')) then GRILLE.OnExit := GRILLE_OnExit;
    if Assigned(GetControl('GRILLE')) then GRILLE.OnEnter := GRILLE_OnEnter;
//GP_20071011_PCO_GP14299
    SetLibelleLibre;
    { événements }
    if Assigned(GetControl('BTVALIDCOMPOSANTS'))  then tToolBarButton97(GetControl('BTVALIDCOMPOSANTS')).OnClick  := BTVALIDCOMPOSANTS_OnClick;
    if Assigned(GetControl('BTSUPPRIMERFILS'))    then tToolBarButton97(GetControl('BTSUPPRIMERFILS')).OnClick    := btSupprimerFils_OnClick;
    if Assigned(GetControl('BTANNULERFILS'))      then tToolBarButton97(GetControl('BTANNULERFILS')).OnClick      := btAnnulerFils_OnClick;
    if Assigned(GetControl('BTNOUVEAUFILS'))      then tToolBarButton97(GetControl('BTNOUVEAUFILS')).OnClick      := btNouveauFils_OnClick;
    if Assigned(GetControl('BTERASEFILS'))        then tToolBarButton97(GetControl('BTERASEFILS')).OnClick        := btEraseFils_OnClick;
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        {$IFDEF GPAO}
          if Assigned(GetControl('PmWND')) then
            TPopupMenu(GetControl('PmWND')).OnPopup := PmWND_OnPopup;
          if Assigned(GetControl('MNLPVISUNOME')) then
            tMenuItem(GetControl('MNLPVISUNOME')).OnClick := MnLpVisuNome_OnClick;
        {$ENDIF GPAO}
      {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
    if Assigned(GetControl('RBCOEFQLIEN'))        then tRadioButton    (GetControl('RBCOEFQLIEN')).OnClick        := rbCoefQLien_OnClick;
    if Assigned(GetControl('RBNOUVQLIEN'))        then tRadioButton    (GetControl('RBNOUVQLIEN')).OnClick        := rbNouvQLien_OnClick;
    if Assigned(GetControl('RBRIENLIEN'))         then tRadioButton    (GetControl('RBRIENLIEN')).OnClick         := rbRienLien_OnClick;
    if Assigned(GetControl('RBRIEN'))             then tRadioButton    (GetControl('RBRIEN')).OnClick             := RBREVISION_OnClick;
    if Assigned(GetControl('RBMINVAL'))           then tRadioButton    (GetControl('RBMINVAL')).OnClick           := RBREVISION_OnClick;
    if Assigned(GetControl('RBMAJVALPEROLD'))     then tRadioButton    (GetControl('RBMAJVALPEROLD')).OnClick     := RBREVISION_OnClick;
    if Assigned(GetControl('QLIEN')) 				      then thNumEdit(GetControl('QLIEN')).OnChange                    := Fields_OnChange;
    if Assigned(GetControl('UNITELIEN')) 		      then thValComboBox(GetControl('UNITELIEN')).OnChange            := Fields_OnChange;
    if Assigned(GetControl('COEFQLIENMULT'))      then thNumEdit(GetControl('COEFQLIENMULT')).OnChange            := Fields_OnChange;
    if Assigned(GetControl('COEFQLIENDIV'))       then thNumEdit(GetControl('COEFQLIENDIV')).OnChange             := Fields_OnChange;
    if Assigned(GetControl('PERTEPROP')) 				  then thNumEdit(GetControl('PERTEPROP')        ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFPERTEPROPMULT')) 	then thNumEdit(GetControl('COEFPERTEPROPMULT')).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFPERTEPROPDIV')) 	then thNumEdit(GetControl('COEFPERTEPROPDIV') ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('QPPER')) 				      then thNumEdit(GetControl('QPPER')            ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('UNITEPPER')) 		      then thValComboBox(GetControl('UNITEPPER')    ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFQPPERMULT'))      then thNumEdit(GetControl('COEFQPPERMULT')    ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFQPPERDIV'))       then thNumEdit(GetControl('COEFQPPERDIV')     ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('QPERIODE')) 		      then thNumEdit(GetControl('QPERIODE')         ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('UNITEPERIODE'))       then thValComboBox(GetControl('UNITEPERIODE') ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFQPERIODEMULT'))   then thNumEdit(GetControl('COEFQPERIODEMULT') ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFQPERIODEDIV'))    then thNumEdit(GetControl('COEFQPERIODEDIV')  ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('QPFIXE')) 				    then thNumEdit(GetControl('QPFIXE')           ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('UNITEPFIXE')) 		    then thValComboBox(GetControl('UNITEPFIXE')   ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFQPFIXEMULT'))     then thNumEdit(GetControl('COEFQPFIXEMULT')   ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('COEFQPFIXEDIV'))      then thNumEdit(GetControl('COEFQPFIXEDIV')    ).OnChange        := Fields_OnChange;
    if Assigned(GetControl('CBQLIEN'))            then tCheckBox(GetControl('CBQLIEN')          ).OnClick        := CBAPPLIQUER_OnClick;
    if Assigned(GetControl('CBPERTEPROP'))        then tCheckBox(GetControl('CBPERTEPROP')      ).OnClick        := CBAPPLIQUER_OnClick;
    if Assigned(GetControl('CBQPPER'))            then tCheckBox(GetControl('CBQPPER')          ).OnClick        := CBAPPLIQUER_OnClick;
    if Assigned(GetControl('CBQPERIODE'))         then tCheckBox(GetControl('CBQPERIODE')       ).OnClick        := CBAPPLIQUER_OnClick;
    if Assigned(GetControl('CBQPFIXE'))           then tCheckBox(GetControl('CBQPFIXE')         ).OnClick        := CBAPPLIQUER_OnClick;
    if Assigned(GRILLE) then
    begin
      { événements }
      GRILLE.OnCellEnter  := GRILLE_OnCellEnter;
      GRILLE.OnCellExit   := GRILLE_OnCellExit;
    end;
    {$IFDEF STK}
    // Récupère le nombre de décimales des zones quantités de la fiche
      NbDeci:= THNumEdit(GetControl('QLIEN')).Decimals;
    {$ENDIF STK}

    { Enabled des champ quantité }
    CBAPPLIQUER_OnClick(TCheckBox(GetControl('CBQLIEN')));
    CBAPPLIQUER_OnClick(TCheckBox(GetControl('CBPERTEPROP')));
    CBAPPLIQUER_OnClick(TCheckBox(GetControl('CBQPPER')));
    CBAPPLIQUER_OnClick(TCheckBox(GetControl('CBQPERIODE')));
    CBAPPLIQUER_OnClick(TCheckBox(GetControl('CBQPFIXE')));
  end
  else if EcranIs('YTARIFSRECH_FIC') then
  begin
    BMP := TImage(GetControl('BMP')).Picture.Bitmap;

    sFonctionnalite := GetArgumentValue(S, 'FONCTIONNALITE');
    sCodePort       := GetArgumentValue(S, 'CODEPORT');
    sTiersFrais     := GetArgumentValue(S, 'TIERSFRAIS');
    sRang           := GetArgumentValue(S, 'RANG');
    sAppel          := GetArgumentValue(S, 'APPEL');
    sTypeFourni     := GetArgumentValue(S, 'TYPEFOURNI');
    sRepartition    := GetArgumentValue(S, 'REPARTITION');

    sParametre := GetParametreFromYFO(sFonctionnalite, sCodePort, 'TIE', 'EntitéActive');

    //GP_DS_GP14214_20071011
    {Type de tiers -> Pour les recherches affaires}
    fTypeTiers := GetTypeTiers(sFonctionnalite);

    //Fixe l'écran de saisi en fonction des paramètres de gestion des informations de la fonctionnalite
    //Ces trois TabSheets regroupent toutes les données disponibles mais sont invisibles
    tTabSheet(GetControl('TSJAMAISAFFICHE1')).TabVisible := False;
    tTabSheet(GetControl('TSJAMAISAFFICHE2')).TabVisible := False;
    tTabSheet(GetControl('TSJAMAISAFFICHE3')).TabVisible := False;
    tTabSheet(GetControl('TSJAMAISAFFICHE4')).TabVisible := False;
    tTabSheet(GetControl('TSCOMPLEMENT1')).TabVisible    := False;
    tTabSheet(GetControl('TSCOMPLEMENT2')).TabVisible    := False;
    //Activation du premier TabSheet de saisie
    tPageControl(GetControl('Pages')).ActivePage := tTabSheet(GetControl('PGENERAL'));
    //Sur ces TabSheets de saisis transfert des données saisissables : paramétrées comme telles dans la fonctionnalité demandée
    iRangPanel:=1; iHauteurDisponible := tTabSheet(GetControl('PGENERAL')).Height;
    if (sParametre[iPosDepot        ]=wTrue) then ActivePanel('GDEPOT'             , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosDevise       ]=wTrue) then ActivePanel('GDEVISE'            , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosSpecial      ]=wTrue) then ActivePanel('GTARIFSPECIAL'      , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosTiers        ]=wTrue) then ActivePanel('GTIERS'             , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosAffaire      ]=wTrue) then ActivePanel('GAFFAIRE'           , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosAtelier      ]=wTrue) then ActivePanel('GATELIER'           , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosSecteurGeo   ]=wTrue) then ActivePanel('GSECTEURGEO'        , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosCodePostal   ]=wTrue) then ActivePanel('GTIERSLIVRE'        , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosCodePostal   ]=wTrue) then ActivePanel('GADRESSE'           , iRangPanel, iHauteurDisponible,  92);
    if (sParametre[iPosModeExp      ]=wTrue) then ActivePanel('GMODEEXP'           , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosNatureTra    ]=wTrue) then ActivePanel('GNATURETRAVAIL'     , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosArticle      ]=wTrue) then ActivePanel('GARTICLE'           , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosArticle      ]=wTrue) then ActivePanel('GBDIMENSIONS'       , iRangPanel, iHauteurDisponible,  58);
    if (sParametre[iPosMarque       ]=wTrue) then ActivePanel('GMARQUE'            , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosChoixQualite ]=wTrue) then ActivePanel('GCHOIXQUALITE'      , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosIndiceArticle]=wTrue) then ActivePanel('GINDICEARTICLE'     , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosRessource    ]=wTrue) then ActivePanel('GRESSOURCE'         , iRangPanel, iHauteurDisponible,  44);
    if (sParametre[iPosCircuit      ]=wTrue) then ActivePanel('GSOUSTRAITANCE'     , iRangPanel, iHauteurDisponible,  92);
    ActivePanel('GCONTEXTE', iRangPanel, iHauteurDisponible, 70);
    //Activation du premier TabSheet de saisie
    tPageControl(GetControl('Pages')).ActivePage  := tTabSheet(GetControl('PGENERAL'));
    tPageControl(GetControl('Pages')).TabWidth    := 100;
    {Permet de ramener le CodeArticle et l'Article}
    OnGetArticle := MyOnGetArticle;
  end
  else if EcranIs('WPA_GETCODEORDRE') then
  begin
    // Récup. de la natureTravail
    NatureTravail := GetArgumentValue(S, 'NATURETRAVAIL');
  end
  else if EcranIs('WPA_GETSTKPROJETE') then
  begin
    if Assigned(GetControl('CB_ANALYSE')) then tCheckBox(GetControl('CB_ANALYSE')).OnClick := CB_ANALYSE_OnClick;
  end
  {$IFDEF SAV}
    else if EcranIs('WPA_NOMPARC') Then
    begin
      {$IFNDEF EAGLSERVER}
        {$IFNDEF ERADIO}
          if Assigned(GetControl('WPA_VARCHAR01')) then
            thdbedit(GetControl('WPA_VARCHAR01')).OnElipsisClick := WPA_VARCHAR01_OnElipsisClick;
          if Assigned(GetControl('WPA_VARCHAR02')) then
            thdbedit(GetControl('WPA_VARCHAR02')).OnElipsisClick := WPA_VARCHAR02_OnElipsisClick;
        {$ENDIF !ERADIO}
      {$ENDIF !EAGLSERVER}
    end
  {$ENDIF SAV}
  {$IFDEF EAGLCLIENT}
    else if EcranIs('WPA_GETPREVCOM') then
    begin
      TFFiche(Ecran).CodeName := '';
    end
  {$ENDIF EAGLCLIENT}
  {$IFDEF QUALITE}
    else if EcranIs('RQGETCLOTURERQN') then
    begin
      if Assigned(GetControl('WPA_VARCHAR01')) then
        thDbEdit(GetControl('WPA_VARCHAR01')).OnElipsisClick := Ressource_OnElipsisClick;
    end
  {$ENDIF QUALITE}
  else if EcranIs('WPDRCALCUL_FIC') then
  begin
    { Prise en compte de ... }
    SetControlVisible('WPA_BOOLEAN07', not GereFraisAvances);
    SetControlVisible('WPA_BOOLEAN08',     GereFraisAvances);
    SetControlVisible('WPA_BOOLEAN09', false               );

    if Assigned(GetControl('WPA_BOOLEAN12')) then TDBCheckBox(GetControl('WPA_BOOLEAN12')).OnClick := WPA_BOOLEAN12_OnClick;
    if Assigned(GetControl('WPA_BOOLEAN13')) then TDBCheckBox(GetControl('WPA_BOOLEAN13')).OnClick := WPA_BOOLEAN13_OnClick;

    sTypePDR:= GetArgumentString(S, 'TYPEPDR');
    sAppel  := GetArgumentValue(S, 'APPEL');
    { Prix de revient d'affaires }
    if sTypePDR = 'AFF' then
      {Seuls enregistrements possibles : "Jour" ou "A Chaque fois"}
      SetControlProperty('WPA_COMBO05', 'DATATYPE', 'WAFPDRUNPAR');
  end
  else if EcranIs('WPA_GETSTRING') then
  begin
    if Assigned(GetControl(GetArgumentString(S, 'FIELDSTRING'))) then
    begin
      with THDBEdit(GetControl(GetArgumentString(S, 'FIELDSTRING'))) do
      begin
        DataType := GetArgumentString(S, 'DATATYPE');
        Plus     := GetArgumentString(S, 'PLUS');
        if DataType <> '' then
        begin
          CharCase := ecUpperCase;
          ElipsisButton := True;
        end
      end
    end
  end
  else if EcranIs('WPA_WPLANLIVR') then {Planning livraison}
  begin
    {$IFDEF GPAOLIGHT}
    {$IFNDEF ERADIO}
    {$IFNDEF EAGLSERVER}
      PlanLivrTobFields := Tob.Create('', nil, - 1);
      PlanLivrTobIndicateurs := Tob.Create('', nil, -1);
      PlanLivrTobItems := GetArgumentTob(s, 'TOBITEMS');
      PlanLivrSetEvents;
      { }
      PlanLivrSB := TScrollBox.Create(GetControl('GBETATS'));
      PlanLivrSB.Parent := TWinControl(GetControl('GBETATS'));
      PlanLivrSB.Name := 'PLSB';
      PlanLivrSB.Align := alClient;
    {$ENDIF !EAGLSERVER}
    {$ENDIF !ERADIO}
    {$ENDIF GPAOLIGHT}
  end
  else if EcranIs('WPA_PDCPARAM') then { Planning Plan de charge}
  begin
    {$IFDEF GPAO}
    {$IFNDEF ERADIO}
    {$IFNDEF EAGLSERVER}
      PlanDeChargeSetEvents;
      Maillage := GetArgumentString(StArgument, 'MAILLAGE');
      TypePlanning := GetArgumentString(StArgument, 'TYPEPLANNING');
      SetControlVisible('WPA_BOOLEAN01', (Maillage = '01J'));
      SetControlVisible('LSAMEDI',       (Maillage = '01J'));
      SetControlVisible('BSAMEDI',       (Maillage = '01J'));
      SetControlVisible('VSAMEDI',       (Maillage = '01J'));
      SetControlVisible('LDIMANCHE',     (Maillage = '01J'));
      SetControlVisible('BDIMANCHE',     (Maillage = '01J'));
      SetControlVisible('VDIMANCHE',     (Maillage = '01J'));
      SetControlVisible('LJOURSFERIES',  (Maillage = '01J'));
      SetControlVisible('BJOURSFERIES',  (Maillage = '01J'));
      SetControlVisible('VJOURSFERIES',  (Maillage = '01J'));
      if      TypePlanning = 'FAM' then
        SetControlCaption('WPA_BOOLEAN04', TraduireMemoire('Montrer 1ère colonne fixe (code famille de ressource)'))
      else if TypePlanning = 'GAN' then
        SetControlCaption('WPA_BOOLEAN04', TraduireMemoire('Montrer 1ère colonne fixe (Ordre - Phase - Opération)'));
    {$ENDIF !EAGLSERVER}
    {$ENDIF !ERADIO}
    {$ENDIF GPAO}
  end
  ;

  {$IFNDEF EAGLSERVER}
    { Loupe }
    if Assigned(GetControl('MNLPLEGENDE')) then TMenuItem(GetControl('MNLPLEGENDE')).OnClick := MnLpLegende_OnClick;

    { Touches }
    TFFiche(Ecran).OnKeyDown := Ecran_OnKeyDown;
  {$ENDIF !EAGLSERVER}
end;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnClose ;
begin
  Inherited ;
  if (copy(GetString('WPA_CODEPARAM'),1,10)=sTarifsRecherche) then
  begin
    TobUneLigneTarif.Free;
    TobContexte.Free;
    TobTarif.Free;
    TobPrixMarge.Free;
    TobAffichage.Free;
  end
  {$IFNDEF EAGLSERVER}
    else if (GetString('WPA_CODEPARAM') = sPrixDeRevient) then
      tfFiche(Ecran).Retour := intToStr(iIdentifiantWPE)
  {$ENDIF !EAGLSERVER}
  ;
  { Remplacement de composants }
  if (Pos(sChangementComposant, GetString('WPA_CODEPARAM')) = 1) and Assigned(TobWPA) then TobWPA.Free;
  {$IFDEF GPAOLIGHT}
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        { Planning de livraison }
        if EcranIs('WPA_WPLANLIVR') then
        begin
          if Assigned(PlanLivrTobFields) then
            FreeAndNil(PlanLivrTobFields);
          if Assigned(PlanLivrTobIndicateurs) then
            FreeAndNil(PlanLivrTobIndicateurs);
          if Assigned(PlanLivrSB) then
            FreeAndNil(PlanLivrSB);
        end;
      {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
  {$ENDIF GPAOLIGHT}
end ;

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.OnCancelRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
   Calcul de champs
---------------------------------------------------------------------------------}
procedure TOM_WPARAM.CalculField(Const FieldName: string);
var
  aArticle, aNomenclature, aGamme, aItineraire, aTiers, aAffaire, aRessource : MyArrayValue;
  {$IFDEF GPAO}
    lMemeArticle   : boolean;
  {$ENDIF GPAO}
  sArticle : string;
begin
  aGamme        := nil;
  aNomenclature := nil;
  aTiers        := nil;
  aItineraire   := nil;
  aArticle      := nil;
  aAffaire      := nil;
  aRessource    := nil;

  inherited;
  {$IFDEF GPAO }
  if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
  begin
    lMemeArticle:=(GetString('WPA_VARCHAR01')<>'');
    { Pour l'article }
    if (FieldName='WPA_VARCHAR01') and (GetString('WPA_VARCHAR01')<>'') then
    begin
      aArticle := WGetSqlFieldsValues(['GA_LIBELLE', 'GA_UNITEPROD'], 'ARTICLE', 'GA_ARTICLE="'+wGetArticleFromCodeArticle(GetString(FieldName))+'"');
      if Assigned(aArticle) then
      begin
        LBARTICLE.Caption    := aArticle[0];
        LBUNITEPROD1.Caption := iifV(lMemeArticle,aArticle[1],'');
        LBUNITEPROD2.Caption := iifV(lMemeArticle,aArticle[1],'');
        LBUNITEPROD3.Caption := iifV(lMemeArticle,aArticle[1],'');
      end;
      SetControlVisible('WPA_BOOLEAN01',lMemeArticle);
      SetControlProperty('TSARTICLE', 'TabVisible', not lMemeArticle);
    end
    { Itinéraire }
    else if (FieldName='WPA_VARCHAR03') then
    begin
      aItineraire := WGetSqlFieldsValues(['QIT_CTX', 'QIT_CODITILIB'], 'QITI', 'QIT_CTX = "0   " AND QIT_CODITI = "'+GetString(FieldName)+'"');
      if (not lMemeArticle) then
      begin
        LBITINERAIRE.Caption    := 'Par défaut';
      end
      else if Assigned(aItineraire) then
      begin
        LBITINERAIRE.Caption    := aItineraire[1];
        SetControlProperty('WPA_VARCHAR04', 'PLUS', 'QCI_CODITI="' + GetString('WPA_VARCHAR03') + '"');
      end
      else
      begin
        LBITINERAIRE.Caption    := 'Pas d''itinéraire';
      end;
    end
    { Circuit }
    else if (FieldName='WPA_VARCHAR04') then
    begin
      if      (not lMemeArticle) then
        LBCIRCUIT.Caption := 'Par défaut'
      else if (GetString(FieldName)='') then
        LBCIRCUIT.Caption := 'Pas de circuit'
      else
        LBCIRCUIT.Caption := WGetSqlFieldValue('QCI_CIRCUITLIB', 'QCIRCUIT', 'QCI_CTX = "0   " AND QCI_CIRCUIT = "'+GetString(FieldName)+'"');
    end
    // Si changement de : Code Nomenclature ou NatureTravail ou Indice Majeur Nomenclature
    else if (FieldName='WPA_VARCHAR05') or (FieldName='WPA_COMBO03') or (FieldName='WPA_COMBO01') then
    begin
      //Affichage : Libellé Révision - Indice Majeur - Indice Mineur
      CleWNT.NatureTravail := GetString('WPA_COMBO03');
      CleWNT.Article       := wGetArticleFromCodeArticle(GetString('WPA_VARCHAR05'));
      CleWNT.Majeur        := GetString('WPA_COMBO01');
      aNomenclature := WGetSqlFieldsValues(['WNT_LIBELLE', 'WNT_ETATREV', 'WNT_MAJEUR', 'WNT_MINEUR'], 'WNOMETET', WhereWNT(CleWNT));
      if (not lMemeArticle) then
      begin
        LBNOMENCLATURELIBELLE.Caption    := 'Par défaut';
        LBNOMENCLATUREREVISION.Caption   := '';
        LBNOMENCLATUREMAJEUR.Caption     := '';
        LBNOMENCLATUREMINEUR.Caption     := '';
      end
      else if Assigned(aNomenclature) then
      begin
        LBNOMENCLATURELIBELLE.Caption    := wPadRight(aNomenclature[0],35);
        LBNOMENCLATUREREVISION.Caption   := aNomenclature[1];
        LBNOMENCLATUREMAJEUR.Caption     := aNomenclature[2];
        LBNOMENCLATUREMINEUR.Caption     := aNomenclature[3];
      end
      else
      begin
        LBNOMENCLATURELIBELLE.Caption    := 'Pas de Nomenclature';
        LBNOMENCLATUREREVISION.Caption   := '';
        LBNOMENCLATUREMAJEUR.Caption     := '';
        LBNOMENCLATUREMINEUR.Caption     := '';
      end;
      // Filtre sur les indices majeurs
      WPA_COMBO01.Plus := 'WNT_ARTICLE = "'+wGetArticleFromCodeArticle(GetString('WPA_VARCHAR05'))+'" AND WNT_NATURETRAVAIL = "'+GetString('WPA_COMBO03')+'" AND WNT_ETATREV <> "PER"';
    end
    // Si changement de : Code Gamme ou NatureTravail ou Indice Majeur Gamme
    else if (FieldName='WPA_VARCHAR06') or (FieldName='WPA_COMBO03') or (FieldName='WPA_COMBO02') then
    begin
      //Affichage : Libellé Révision - Indice Majeur - Indice Mineur
      CleWGT.NatureTravail := GetString('WPA_COMBO03');
      CleWGT.Article       := wGetArticleFromCodeArticle(GetString('WPA_VARCHAR06'));
      CleWGT.Majeur        := GetString('WPA_COMBO02');
      aGamme := WGetSqlFieldsValues(['WGT_LIBELLE', 'WGT_ETATREV', 'WGT_MAJEUR', 'WGT_MINEUR'], 'WGAMMETET', WhereWGT(CleWGT));
      if (not lMemeArticle) then
      begin
        LBGAMMELIBELLE.Caption    := 'Par défaut';
        LBGAMMEREVISION.Caption   := '';
        LBGAMMEMAJEUR.Caption     := '';
        LBGAMMEMINEUR.Caption     := '';
      end
      else if Assigned(aGamme) then
      begin
        LBGAMMELIBELLE.Caption    := wPadRight(aGamme[0],35);
        LBGAMMEREVISION.Caption   := aGamme[1];
        LBGAMMEMAJEUR.Caption     := aGamme[2];
        LBGAMMEMINEUR.Caption     := aGamme[3];
      end
      else
      begin
        LBGAMMELIBELLE.Caption    := 'Pas de Gamme';
        LBGAMMEREVISION.Caption   := '';
        LBGAMMEMAJEUR.Caption     := '';
        LBGAMMEMINEUR.Caption     := '';
      end;

      // Filtre sur les indices majeurs
      WPA_COMBO02.Plus := 'WGT_ARTICLE = "'+wGetArticleFromCodeArticle(GetString('WPA_VARCHAR06'))+'" AND WGT_NATURETRAVAIL = "'+GetString('WPA_COMBO03')+'" AND WGT_ETATREV <> "PER"';
    end
    else if (FieldName='WPA_BOOLEAN01') then
    begin
      AffichageParDefaut(GetBoolean('WPA_BOOLEAN01'));
    end
    ;
  end
  ;
  {$ENDIF GPAO}

  if (copy(GetString('WPA_CODEPARAM'),1,10)=sTarifsRecherche) then
  begin
    //Pour la fonctionnalite
    if           (FieldName='WPA_COMBO01') then
    begin
      if Assigned(GBSousTraitance)    then  GBSousTraitance.Visible    := (GetString('WPA_COMBO01')=sTarifSousTraitantAchat) or (GetString('WPA_COMBO01')=sTarifSousTraitantPhase);
    end
    //Pour le tiers
    {$IFNDEF EAGLSERVER}
      else if      (FieldName='WPA_VARCHAR02') and (LookupValueExist(THEdit(Getcontrol('WPA_VARCHAR02')))) then
      begin
        aTiers := wGetSqlFieldsValues(['T_LIBELLE', 'T_TARIFTIERS'], 'TIERS', 'T_TIERS="'+GetString(FieldName)+'"');
        if Assigned(aTiers) then
          LBTIERS.Caption  := aTiers[0]+' '+aTiers[1]
        else
          LBTIERS.Caption  := 'Pas de Tiers'
      end
    {$ENDIF !EAGLSERVER}
    //Pour l'article
    else if (FieldName='WPA_VARCHAR03') then
    begin

      if (GetString('WPA_VARCHAR16')<>'') and (GetString('WPA_VARCHAR03')=wGetCodeArticleFromArticle(GetString('WPA_VARCHAR16'))) then
        sArticle := GetString('WPA_VARCHAR16')
      else
        sArticle := wGetArticleFromCodeArticle(GetString('WPA_VARCHAR03'));

      aArticle := WGetSqlFieldsValues(['GA_LIBELLE', 'GA_TARIFARTICLE', 'GA_QUALIFUNITESTO', 'GA_QUALIFUNITEVTE', 'GA_STATUTART'], 'ARTICLE', 'GA_ARTICLE="'+sArticle+'"');
      if Assigned(aArticle) then
      begin
        LBARTICLE.Caption  := aArticle[0]+' '+aArticle[1];
        {Frais annexes / Coûts indirects}
        if (pos(sFonctionnalite, sFraisAnnexes+';'+sCoutsIndirects)>0) then
          SetControlProperty('LBUNITE', 'CAPTION', iif(aArticle[2]<>'', RechDom('WUNITE',aArticle[2],False),'')) // Unité de stock
        else
          SetControlProperty('LBUNITE', 'CAPTION', iif(aArticle[3]<>'', RechDom('WUNITE',aArticle[3], False),'')); // Unité de facturation
      end
      else
      begin
        LBARTICLE.Caption  := TraduireMemoire('Pas d''article');
        SetControlProperty('LBUNITE', 'CAPTION', '');
      end;
      ControleAffichageDimensionsArticle;
    end
    else if (FieldName='WPA_VARCHAR05') then
    begin
      if (GetString(FieldName)<>'') then
      begin
        aAffaire := WGetSqlFieldsValues(['AFF_LIBELLE'], 'AFFAIRE', 'AFF_AFFAIRE="'+GetString(FieldName)+'"');
        if Assigned(aAffaire) then
          LBAFFAIRE.Caption  := aAffaire[0]
        else
          LBAFFAIRE.Caption  := TraduireMemoire('Affaire inconnue');
      end
      else
        LBAFFAIRE.Caption  := '';
    end
    else if (FieldName='WPA_VARCHAR12') then
    begin
      if (GetString(FieldName)<>'') then
      begin
        LBATELIER.Caption := wGetSqlFieldValue('QSI_SITELIB'  , 'QSITE'  , 'QSI_CTX="0" and QSI_SITE="'+GetString('WPA_VARCHAR12')+'"');
        if (LBATELIER.Caption='') then
          LBATELIER.Caption    := TraduireMemoire('Atelier inconnu');
      end
      else
        LBATELIER.Caption := '';
    end
    else if (FieldName='WPA_VARCHAR06') then    // Tiers livré
    begin
      {$IFNDEF EAGLSERVER}
      if (GetString(FieldName)<>'') and (LookupValueExist(THEdit(Getcontrol('WPA_VARCHAR06')))) then
      begin
        SetControlEnabled('WPA_VARCHAR13', True);
        aTiers := wGetSqlFieldsValues(['YTC_SECTEURGEO','T_CODEPOSTAL','T_PAYS','T_REGION','YTC_MODEEXP','T_LIBELLE','YTC_NADRESSELIV'], 'TIERS LEFT JOIN TIERSCOMPL ON (YTC_TIERS=T_TIERS)', '(T_NATUREAUXI="'+sNatureAuxiClient+'" OR T_NATUREAUXI="'+sNatureAuxiProspect+'") AND T_TIERS="'+GetString(FieldName)+'"');
        if Assigned(aTiers) then
        begin
          LBLIVRE.Caption  := aTiers[5];
          if aTiers[6]<>0 then  // N° adresse livraison
          begin
            if ValeurI(GetString('WPA_VARCHAR13')) <> aTiers[6] then
              SetString('WPA_VARCHAR13' , aTiers[6])
            else
              CalculField('WPA_VARCHAR13');
          end
          else
          begin
            if ValeurI(GetString('WPA_VARCHAR13')) <> aTiers[6] then
              { Remise à blanc n° adresse de livraison : Cela réinitialisera l'adresse du client }
              SetString('WPA_VARCHAR13','')
            else
              CalculField('WPA_VARCHAR13');
          end;

        end;
      end
      else  // Tiers n'existe pas ou vide
      {$ENDIF !EAGLSERVER}
      begin
        SetControlText   ('WPA_VARCHAR13' , '');
        SetControlEnabled('WPA_VARCHAR13' , False);
        SetControlEnabled('WPA_COMBO04'   , True);
        SetControlEnabled('WPA_VARCHAR10' , True);
        SetControlEnabled('WPA_COMBO05'   , True);
        SetControlEnabled('WPA_VARCHAR11' , True);
        SetControlEnabled('WPA_COMBO07'   , True);
      end;
    end
    else if (FieldName='WPA_VARCHAR13') then    // N° Adresse livré
    begin
      if (GetString(FieldName)<>'') and (GetInteger(FieldName)<>0) then
      begin
        { Adresse livraison ou facturation }
        aTiers := wGetSqlFieldsValues(['ADR_SECTEURGEO','ADR_CODEPOSTAL','ADR_PAYS','ADR_REGION','ADR_MODEEXP'], 'ADRESSES', 'ADR_REFCODE="'+GetString('WPA_VARCHAR06')+'" AND ADR_NADRESSE='+GetString(FieldName));
        if Assigned(aTiers) then
        begin
            SetAdresseTiers(aTiers);
        end
        else
        begin
        { N° Adresse inexistant -> On remet l'adresse tiers};
          aTiers := wGetSqlFieldsValues(['YTC_SECTEURGEO','T_CODEPOSTAL','T_PAYS','T_REGION','YTC_MODEEXP','T_LIBELLE','YTC_NADRESSELIV'], 'TIERS LEFT JOIN TIERSCOMPL ON (YTC_TIERS=T_TIERS)', '(T_NATUREAUXI="'+sNatureAuxiClient+'" OR T_NATUREAUXI="'+sNatureAuxiProspect+'") AND T_TIERS="'+GetString('WPA_VARCHAR06')+'"');
          if Assigned(aTiers) then
            SetAdresseTiers(aTiers);
        end;
      end
      else  // N° adresse à zéro ou blanc
      begin
          aTiers := wGetSqlFieldsValues(['YTC_SECTEURGEO','T_CODEPOSTAL','T_PAYS','T_REGION','YTC_MODEEXP','T_LIBELLE','YTC_NADRESSELIV'], 'TIERS LEFT JOIN TIERSCOMPL ON (YTC_TIERS=T_TIERS)', '(T_NATUREAUXI="'+sNatureAuxiClient+'" OR T_NATUREAUXI="'+sNatureAuxiProspect+'") AND T_TIERS="'+GetString('WPA_VARCHAR06')+'"');
        if Assigned(aTiers) then
          { Adresse Tiers }
          SetAdresseTiers(aTiers);
      end;
    end;
  end
  ;
end;

{--------------------------------------------------------------------------------
   Controle de champs
---------------------------------------------------------------------------------}
procedure TOM_WPARAM.ControlField(Const FieldName: string);
  {$IFDEF GPAO}
  var
    lMemeArticle   : boolean;
    aArticleNature : MyArrayValue;
    aTypePDR       : MyArrayValue;
		QEcoProd			 : Double;
  {$ENDIF GPAO}

  {-----------------------------------------------------------------------------
    Recherche de l'itinéraire    par défaut de l'article
    Recherche du circuit         par défaut de l'article
    Recherche de la nomenclature par défaut de l'article
    Recherche de la gamme        par défaut de l'article
  ------------------------------------------------------------------------------}

 {$IFDEF GPAO}
  procedure InitParDefautArticle;
  var
    sCodeArticleWNT, sCodeArticleWGT : string;
  begin
    //Code Itinéraire - Code Circuit
    aArticleNature := WGetSqlFieldsValues(['WAN_CODITI', 'WAN_CIRCUIT', 'WAN_CODEARTICLEWNT', 'WAN_CODEARTICLEWGT'], 'WARTNAT', 'WAN_ARTICLE = "'+wGetArticleFromCodeArticle(GetString('WPA_VARCHAR01'))+'" AND WAN_NATURETRAVAIL = "'+GetString('WPA_COMBO03')+'"');
    if Assigned(aArticleNature) then
    begin
      SetString('WPA_VARCHAR03',aArticleNature[0]); ControlField('WPA_VARCHAR03');
      SetString('WPA_VARCHAR04',aArticleNature[1]); ControlField('WPA_VARCHAR04');

      { Code Nomenclature + Indice majeur }
      if (aArticleNature[2]<>'') then
        sCodeArticleWNT := aArticleNature[2]
      else
        sCodeArticleWNT := GetString('WPA_VARCHAR01');

      SetString('WPA_VARCHAR05',sCodeArticleWNT); ControlField('WPA_VARCHAR05');
      SetString('WPA_COMBO01', wGetActifMajeur('WNT', GetString('WPA_COMBO03'), wGetArticleFromCodeArticle(GetString('WPA_VARCHAR05')), wndNone, '2'));
      ControlField('WPA_COMBO01');
      if (GetString('WPA_COMBO01')='') then SetString('WPA_VARCHAR05','');

      { Code Gamme + Indice majeur }
      if (aArticleNature[3]<>'') then
        sCodeArticleWGT := aArticleNature[3]
      else
        sCodeArticleWGT := GetString('WPA_VARCHAR01');

      SetString('WPA_VARCHAR06',sCodeArticleWGT); ControlField('WPA_VARCHAR06');
      SetString('WPA_COMBO02', wGetActifMajeur('WGT', GetString('WPA_COMBO03'), wGetArticleFromCodeArticle(GetString('WPA_VARCHAR06')), wndNone, '2'));
      ControlField('WPA_COMBO02');
      if (GetString('WPA_COMBO02')='') then SetString('WPA_VARCHAR06','');

      { Qté Economique de Production }
			QEcoProd:= wGetFieldFromGA('GA_QECOPROD',wGetArticleFromCodeArticle(GetString('WPA_VARCHAR01')));
      SetDouble('WPA_DOUBLE02', iif(QEcoProd<>0, QEcoProd, 1));
      { Par défaut : Qté demandée = Qté économique de fabrication }
      SetDouble('WPA_DOUBLE01',GetDouble('WPA_DOUBLE02'));

    end
    else
    begin
      SetString('WPA_VARCHAR03',''); ControlField('WPA_VARCHAR03'); //Itinéraire
      SetString('WPA_VARCHAR04',''); ControlField('WPA_VARCHAR04'); //Circuit
      SetString('WPA_VARCHAR05','');  //Nomenclature
      SetString('WPA_VARCHAR06','');  //Gamme
      SetDouble('WPA_DOUBLE01' ,1 );  //Prix de Revient pour une qté de
      SetDouble('WPA_DOUBLE02' ,0 );  //Qté Economique de Production
      SetDouble('WPA_DOUBLE03' ,1 );  //Prix de Revient exprimé pour une quantité de
    end;
  end;
  {$ENDIF GPAO}

begin
  {$IFDEF GPAO}
    aTypePdr := nil;
    aArticleNature := nil;
  {$ENDIF GPAO}  

  DisableControl;
  try
    {$IFDEF GPAO}
    if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
    begin
      //Ne pas faire ce test après le New Record : Attente de Afterinserting
      if (FieldName='WPA_VARCHAR01')  then
      begin
        SetBoolean('WPA_BOOLEAN01', True); ControlField('WPA_BOOLEAN01'); l1erClickBoolean01 := True;

        lMemeArticle:=(GetString('WPA_VARCHAR01')<>'');
        SetControlVisible('WPA_BOOLEAN01',lMemeArticle);
        SetControlProperty('TSARTICLE', 'TabVisible', not lMemeArticle);
        if (not lMemeArticle) then
          tPageControl(GetControl('Pages')).ActivePage := tTabSheet(GetControl('TSARTICLE'));

        if (not lMemeArticle) then
        begin
          SetString('WPA_VARCHAR03', '');  //Itinéraire
          SetString('WPA_VARCHAR04', '');  //Circuit
          SetString('WPA_VARCHAR05', '');  //Nomenclature
          SetString('WPA_VARCHAR06', '');  //Gamme
          SetDouble('WPA_DOUBLE01' , 1 );  //Prix de Revient pour une qté de
          SetDouble('WPA_DOUBLE02' , 0 );  //Qté Economique de Production
          SetDouble('WPA_DOUBLE03' , 1 );  //Prix de Revient exprimé pour une qté de
        end;
      end
      //Champ Par Defaut et Même Article alors Init de la configuration par défaut
      else if (FieldName='WPA_BOOLEAN01') and GetBoolean('WPA_BOOLEAN01') and (GetString('WPA_VARCHAR01')<>'') then
      begin
        InitParDefautArticle;
      end
      else if (FieldName='WPA_VARCHAR07') then
      begin
        aTypePDR := WGetSqlFieldsValues(['WRT_FRACTFIXE', 'WRT_AVECPERTE', 'WRT_AVECINDIRECT', 'WRT_AVECFLUX', 'WRT_AVECMARGE', 'WRT_AVECTARIF', 'WRT_TVACHAT', 'WRT_TVRESSOURCE', 'WRT_AVECQPCB', 'WRT_NATUREPDR', 'WRT_MAJDPADPR', 'WRT_AVECCATALOGUE', 'WRT_AVECDISPO', 'WRT_AVECARTICLE', 'WRT_AVECARTICLEDEF', 'WRT_AVECARTICLESIM', 'WRT_PDRENREGWPL'], 'WPDRTYPE', 'WRT_TYPEPDR= "'+GetString('WPA_VARCHAR07')+'"');
        if Assigned(aTypePDR) then
        begin
          SetString('WPA_COMBO08'  , aTypePDR[ 0]);  //Fractionnement des fraix fixes
          SetString('WPA_BOOLEAN06', aTypePDR[ 1]);  //Avec Perte
          SetString('WPA_BOOLEAN07', aTypePDR[ 2]);  //Avec Indirect
          SetString('WPA_BOOLEAN08', aTypePDR[ 3]);  //Avec Flux Annexe
          SetString('WPA_BOOLEAN09', aTypePDR[ 4]);  //Avec Marge
          SetString('WPA_BOOLEAN03', aTypePDR[ 5]);  //Avec Tarif Achat
          SetString('WPA_COMBO04'  , aTypePDR[ 6]);  //Type de Valorisation Achat
          SetString('WPA_COMBO06'  , aTypePDR[ 7]);  //Type de Valorisation Ressource
          SetString('WPA_BOOLEAN04', aTypePDR[ 8]);  //En tenant des quantités : Achetés par multiple de
          SetString('WPA_VARCHAR08', aTypePDR[ 9]);  //Nature de Prix de revient
          SetString('WPA_BOOLEAN10', aTypePDR[10]);  //Mise à jour des DPA et DPR des fiches article et stock
          SetString('WPA_BOOLEAN11', aTypePDR[11]);  //Avec prix catalogue
          SetString('WPA_BOOLEAN12', aTypePDR[12]);  //Avec prix article/dépôt
          SetString('WPA_BOOLEAN13', aTypePDR[13]);  //Avec prix article
          SetString('WPA_BOOLEAN14', aTypePDR[14]);  //Avec prix article : prix d'achat par défaut
          SetString('WPA_BOOLEAN15', aTypePDR[15]);  //Avec prix article : prix de simulation
          SetString('WPA_BOOLEAN16', aTypePDR[16]);  //Avec enregistrement du détail
          if (GetString('WPA_COMBO05')='') then
            SetString('WPA_COMBO05', '010');
          ;
        end
      end
      ;
    end
    else if EcranIs('WPA_ORDREINTEGRAT') or EcranIs('WPA_ORDRECPDCOR') then
    begin
      if FieldName = 'WPA_DOUBLE03' then
      begin
        { Total }
        SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') - GetDouble('WPA_DOUBLE03'));

        { Conversion }
        SetConversion;
      end;
    end
    else if EcranIs('WPA_ORDRESERVICE') or EcranIs('WPA_ORDRECPDREC') then
    begin
      if FieldName = 'WPA_DOUBLE03' then
      begin
        { Total }
        SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') + GetDouble('WPA_DOUBLE03'));
        
        { Conversion }
        SetConversion;
      end;
    end;
    {$ENDIF GPAO}
    {$IFDEF GPAOLIGHT}
    if EcranIs('WPA_ORDRELANC') or EcranIs('WPA_ORDRERECEP') then
    begin
      if FieldName = 'WPA_DOUBLE03' then
      begin
        if EcranIs('WPA_ORDRELANC') and ((GetDouble(FieldName) +  GetDouble('WPA_DOUBLE02')) > GetDouble('WPA_DOUBLE01')) then LastError := ErrQLancement
        else
        begin
          { Total }
          SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') + GetDouble('WPA_DOUBLE03'));

          { Conversion }
          SetConversion;
        end;
      end;
    end;
    {$ENDIF GPAOLIGHT}
    {$IFDEF GPAO}
    if EcranIs('WPA_ORDREREBUT') then
    begin
      if FieldName = 'WPA_DOUBLE03' then
      begin
        { Total }
        SetDouble('WPA_DOUBLE04', GetDouble('WPA_DOUBLE02') + GetDouble('WPA_DOUBLE03'));
        
        { Conversion }
        SetConversion;
      end;
    end
    else if EcranIs('WPA_CHGVERSION') then
    begin
      if FieldName = 'WPA_VARCHAR02' then
      begin
        if not lInitdata and (wExistUneVersion(GetCleWVS.article) and ((GetString('WPA_VARCHAR02') = '') or (not wExistVersion(GetCleWVS)))) then
          LastError := ErrCodeVersion;
        linitData := False;
      end;
    end;
    {$ENDIF GPAO}
    {$IFDEF GPAOLIGHT}
    if EcranIs('WPA_FAISABILITE') then
    begin
      if FieldName = wFaisabilite.Wpa_CodeArticle then
      begin
        if GetString(wFaisabilite.Wpa_CodeArticle) = '' then
        begin
          SetString(wFaisabilite.Wpa_CodeArticleWNT, '');
          SetString(wFaisabilite.Wpa_MajeurWNT, '');
          SetString(wFaisabilite.Wpa_Coditi, '');
          SetString(wFaisabilite.Wpa_Circuit, '');
          SetDouble(wFaisabilite.Wpa_QLotSais, 0);
          SetDouble(wFaisabilite.Wpa_QLotStoc, 0);
          SetDouble(wFaisabilite.Wpa_CoefLot, 0);
          SetString(wFaisabilite.Wpa_QualifUniteSto, '');
          SetString(wFaisabilite.Wpa_UniteLot, '');
        end;
        {$IFDEF GPAO}
          if not wExistWAN(wGetCleWAN(GetString(wFaisabilite.Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(wFaisabilite.Wpa_CodeArticle)))) then
            LastError := ErrArtNat
          { Gestion de la nature de travail de l'article inactive }
          else if not wExistWANActif(wGetCleWAN(GetString(Wpa_NatureTravail),wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle)))) then
            LastError := ErrNatArtInactive;
        {$ELSE GPAO}
          if wGetFieldFromGA('GA_TYPENOMENC', wGetArticleFromCodeArticle(GetString(wFaisabilite.Wpa_CodeArticle))) <> 'ASS' then
            LastError := ErrArtNat;
        {$ENDIF !GPAO}
        if (GetString(wFaisabilite.Wpa_CodeArticle) <> '') and (LastError = 0) then 
        begin
          {$IFNDEF GPAO}
            SetString(wFaisabilite.Wpa_NatureTravail, 'FAB');
          {$ENDIF !GPAO}
          { Nomenclature + indice majeur + Quantité + Unité}
          SetCodeArticleWNT;
          { Filtre sur l' indice majeur WNT }
          SetControlProperty(wFaisabilite.Wpa_MajeurWNT, 'PLUS', wGetPlusMajeurWNT(GetString('WPA_COMBO01'), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle))));
          { itinéraire }
          SetString(wFaisabilite.Wpa_Coditi, wGetFieldFromWAN('WAN_CODITI', wGetCleWAN(GetString(Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle)))));
          { Filtre sur les nomenclatures de même nature + itinéraire }
          SetControlProperty(wFaisabilite.Wpa_CodeArticleWNT, 'PLUS', wGetPlusCodeArticleWNT(GetString(wFaisabilite.Wpa_NatureTravail),GetString(wFaisabilite.Wpa_Coditi)));
          { Filtre sur les circuits }
          SetControlProperty(wFaisabilite.Wpa_Circuit, 'PLUS', 'QCI_CODITI="' + GetString(wFaisabilite.Wpa_CodIti) + '"');
          SetString(wFaisabilite.Wpa_Circuit, wGetFieldFromWAN('WAN_CIRCUIT', wGetCleWAN(GetString(wFaisabilite.Wpa_NatureTravail),wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle)))));
        end;
      end
      else if FieldName = wFaisabilite.Wpa_NatureTravail then
      begin
        ControlField(Wpa_CodeArticle);
      end
      {$IFDEF GPAO}
      else if FieldName = wFaisabilite.Wpa_CodeArticleWNT then
      begin
        if GetString(FieldName) = ''  then SetString(wFaisabilite.Wpa_MajeurWNT, '')
        else
        begin
          if (not wExistWAN(wGetCleWAN(GetString(wFaisabilite.Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticleWNT))))) then LastError := ErrArticleWNT
          else
          begin
            { Indice majeur }
            SetString(wFaisabilite.Wpa_MajeurWNT, wGetActiveCleWNT(GetString(Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(wFaisabilite.Wpa_CodeArticleWNT))).majeur);
            { Filtre sur les indices majeurs }
            SetControlProperty(wFaisabilite.Wpa_MajeurWNT, 'PLUS', wGetPlusMajeurWNT(GetString(wFaisabilite.Wpa_NatureTravail), wgetArticleFromCodeArticle(GetString(wFaisabilite.Wpa_CodeArticleWNT))));
            if not wWNTisOk(GetCleWNT) then LastError := ErrArticleWNT2
            else
              { Quantité }
              SetQWNT;
          end;
        end;
      end
      else if FieldName = Wpa_Circuit then
      begin
        if (GetString(wFaisabilite.Wpa_Circuit) <> '') and  (not wExistQCI(GetString(wFaisabilite.Wpa_Circuit)) ) then LastError := ErrCodeCircuit
      end
      {$ENDIF GPAO}
      else if FieldName = 'WPA_DOUBLE01' then
      begin
         SetDouble('WPA_DOUBLE02', GetDouble('WPA_DOUBLE01') * GetDouble('WPA_DOUBLE03'));
      end;
    end;
    {$ENDIF GPAOLIGHT }
    if (copy(GetString('WPA_CODEPARAM'),1,10)=sTarifsRecherche) then
    begin
      if (FieldName='WPA_COMBO01') then     //Fonctionnalité
      begin
        ClientOuFournisseur;
      end
      else if (FieldName='WPA_VARCHAR03') then // Article
      begin
        ParticularitesSousTraitance;
        CtrlAffichageMarqueChoixQualiteIndice;
      end
      else if (FieldName='WPA_VARCHAR02') then // Tiers
      begin
        CtrlAffichageMarqueChoixQualiteIndice;
      end
      ;
    end
//GP_DKZ_GP_PDC_20070927
    {$IFDEF GPAO}
    else if EcranIs('WPA_PDCPARAM') then
    begin
      if FieldName = 'WPA_BOOLEAN02' then
        SetBoolean('WPA_BOOLEAN07', not tCheckBox(GetControl('WPA_BOOLEAN02')).Checked)
      else if FieldName = 'WPA_BOOLEAN07' then
        SetBoolean('WPA_BOOLEAN02', not tCheckBox(GetControl('WPA_BOOLEAN07')).Checked)
//GP_20080708_DKZ_GP15150 Déb
      else if FieldName = 'WPA_BOOLEAN09' then
      begin
        SetControlEnabled('WPA_DOUBLE08', not tCheckBox(GetControl('WPA_BOOLEAN09')).Checked);
        SetControlEnabled('WPA_DOUBLE09', not tCheckBox(GetControl('WPA_BOOLEAN09')).Checked);
        SetInteger('WPA_DOUBLE08', 0);
        SetInteger('WPA_DOUBLE09', 0);
      end
//GP_20080708_DKZ_GP15150 Fin
      ;
    end
    {$ENDIF GPAO}
{$IFDEF QUALITE}
    else if GetString('WPA_CODEPARAM') = 'RQGETCLOTURERQN' then
    begin
      if ((FieldName='WPA_VARCHAR01')) and (GetString(FieldName)<>'') then     //Tiers
	    	if not wExistRessource(GetString('WPA_VARCHAR01'),'','',True) then
          LastError := ErrExistRessource;
    end
{$ENDIF QUALITE}
  finally
    EnableControl;
  end;

  inherited;

  if (LastError>0) then
  begin
    if LastError <> 0 then
      LastErrorMsg := TexteMessage[LastError];
  end;
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF SAV}
procedure TOM_WPARAM.WPA_VARCHAR01_OnElipsisClick(Sender: Tobject);
var Range, Cle_Parc : string;
begin
  Range := 'WPC_CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('WPA_VARCHAR01'))
           + ';WPC_SERIEINTERNE='  + GetString('WPA_VARCHAR02')
           ;
  Cle_Parc := AglLanceFiche('W', 'WPARC_MUL', Range, '', 'ACTION=MODIFICATION;MULRECHERCHE;NOMENCLATURE');

  if (Cle_Parc <> '') then
  begin
    //Pour passer DS.state en mode dsEdit sinon SetString ne fonctionne pas
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.edit;
    //Affectation des champs par rapport aux champs retournés
    SetString('WPA_VARCHAR01',wGetArticleFromCodeArticle(ReadTokenSt(Cle_Parc)));
    SetString('WPA_VARCHAR02'  ,ReadTokenSt(Cle_Parc));
  end;
end;
{$ENDIF SAV}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF SAV}
procedure TOM_WPARAM.WPA_VARCHAR02_OnElipsisClick(Sender: Tobject);
var Range, Cle_Parc : string;
begin
  Range := 'WPC_CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('WPA_VARCHAR01'))
           + ';WPC_SERIEINTERNE='  + GetString('WPA_VARCHAR02')
           ;
  Cle_Parc := AglLanceFiche('W', 'WPARC_MUL', Range, '', 'ACTION=MODIFICATION;MULRECHERCHE;NOMENCLATURE');

  if (Cle_Parc <> '') then
  begin
    //Pour passer DS.state en mode dsEdit sinon SetString ne fonctionne pas
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.edit;
    //Affectation des champs par rapport aux champs retournés
    SetString('WPA_VARCHAR01',wgetArticleFromCodeArticle(ReadTokenSt(Cle_Parc)));
    SetString('WPA_VARCHAR02'  ,ReadTokenSt(Cle_Parc));
  end;
end;
{$ENDIF SAV}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
{--------------------------------------------------------------------------------
   Appel de la liste de choix des nomenclatures
---------------------------------------------------------------------------------}
procedure TOM_WPARAM.WPA_VARCHAR05_OnElipsisClick(Sender: Tobject);
var
  F4Return : string;
  Range    : string;
begin
  if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
  begin
    Range := 'NATURETRAVAIL='+ GetString('WPA_COMBO03')
           + ';ARTICLE='     + wGetArticleFromCodeArticle(GetString('WPA_VARCHAR05'))
           + ';RECHERCHE='   + wTrue
           ;
    F4Return := CallMulWNT(Range);

    if (F4return <> '') then
    begin
      //Pour passer DS.state en mode dsEdit sinon SetString ne fonctionne pas
      if not(DS.State in [dsInsert,dsEdit]) then
        DS.edit;
      //Affectation des champs par rapport aux champs retournés
      SetString('WPA_COMBO03'  ,ReadTokenSt(F4Return));
      SetString('WPA_VARCHAR05',wGetCodeArticleFromArticle(ReadTokenSt(F4Return)));
      SetString('WPA_COMBO01'  ,ReadTokenSt(F4Return));
    end;
  end;
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
{--------------------------------------------------------------------------------
   Appel de la liste de choix des gammes
---------------------------------------------------------------------------------}
procedure TOM_WPARAM.WPA_VARCHAR06_OnElipsisClick(Sender: Tobject);
var
  F4Return : string;
  Range    : string;
begin
  if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
  begin
    Range := 'NATURETRAVAIL=' + GetString('WPA_COMBO03')
           + ';ARTICLE='      + GetString('WPA_VARCHAR06')
           + ';RECHERCHE='    + wTrue
           ;
    F4Return := CallMulWGT(Range);

    if (F4return <> '') then
    begin
      //Pour passer DS.state en mode dsEdit sinon SetString ne fonctionne pas
      if not(DS.State in [dsInsert,dsEdit]) then
        DS.edit;
      //Affectation des champs par rapport aux champs retournés
      SetString('WPA_COMBO03'  ,ReadTokenSt(F4Return));
      SetString('WPA_VARCHAR06',wGetCodeArticleFromArticle(ReadTokenSt(F4Return)));
      SetString('WPA_COMBO02'  ,ReadTokenSt(F4Return));
    end;
  end;
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{--------------------------------------------------------------------------------*
   Controle si l'enregistrement est valide
---------------------------------------------------------------------------------}
function TOM_WPARAM.RecordIsValid: boolean;
var
  {$IFDEF GPAO}
    lMemeArticle: boolean;
    iCpt        : Integer;
  {$ENDIF GPAO}

  { Remplacement de composant }
  iCol, iRow  : Integer;
  bCancel     : Boolean;
  C           : tWinControl;
  sItineraire : string;
  { ------------------------ }
//GP_20080507_DKZ_GP14995 Déb
  Article     : string;
  { ------------------------ }
//GP_20080507_DKZ_GP14995 Fin
begin
  { Remplacement de composant }
  C := nil;
  { ------------------------ }

  {$IFDEF GPAO}
  if (copy(GetString('WPA_CODEPARAM'),1,13) = sPrixDeRevient) then
  begin
    lMemeArticle := (GetString('WPA_VARCHAR01')<>'');
    CleWNT.NatureTravail := GetString('WPA_COMBO03');
    CleWNT.Article       := wGetArticleFromCodeArticle(GetString('WPA_VARCHAR05'));
    CleWNT.Majeur        := GetString('WPA_COMBO01');

    CleWGT.NatureTravail := GetString('WPA_COMBO03');
    CleWGT.Article       := wGetArticleFromCodeArticle(GetString('WPA_VARCHAR06'));
    CleWGT.Majeur        := GetString('WPA_COMBO02');

    if ((GetString('WPA_VARCHAR01')<>'') and (not wExistGA(wGetArticleFromCodeArticle(GetString('WPA_VARCHAR01'))))) then LastError := ErrCodeArticle
    else if ( (GetString('WPA_COMBO03')  ='')                                                                ) then LastError := ErrCodeNature
    else if ( (lMemeArticle) and (not wExistQIT(GetString('WPA_VARCHAR03')))                                 ) then LastError := ErrCodeItineraire
    else if ( (lMemeArticle) and (not wExistQCI(GetString('WPA_VARCHAR04'), GetString('WPA_VARCHAR03')))     ) then LastError := ErrCodeCircuit
    else if ( (lMemeArticle) and (GetString('WPA_VARCHAR05')<>'') and (not ExistWNT(CleWNT))                 ) then LastError := ErrCodeNomenclature
    else if ( (lMemeArticle) and (GetString('WPA_VARCHAR05')<>'') and (GetString('WPA_COMBO01')  ='')        ) then LastError := ErrCodeIndiceMajeurNomenclature
    else if ( (lMemeArticle) and (GetString('WPA_VARCHAR06')<>'') and (not ExistWGT(CleWGT))                 ) then LastError := ErrCodeGamme
    else if ( (lMemeArticle) and (GetString('WPA_VARCHAR06')<>'') and (GetString('WPA_COMBO02')  ='')        ) then LastError := ErrCodeIndiceMajeurGamme
    else if ( (GetString('WPA_COMBO04')='') and (GetBoolean('WPA_BOOLEAN12') or GetBoolean('WPA_BOOLEAN13')) ) then LastError := ErrMethodeValorisation
    ;
  end
  else if EcranIs('WPA_ORDRELANC') then
  begin
    if (getDouble('WPA_DOUBLE03') +  GetDouble('WPA_DOUBLE02')) > GetDouble('WPA_DOUBLE01') then LastError := ErrQLancement   ;
  end;
  {$ENDIF GPAO}
  {$IFDEF GPAOLIGHT}
  if EcranIs('WPA_FAISABILITE') then
  begin
    {$IFNDEF GPAO}
    if (GetString(Wpa_CodeArticle) = '') or (wGetFieldFromGA('GA_TYPENOMENC', wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle))) <> 'ASS') then
      LastError := ErrArtNat
    {$ELSE GPAO}
    if (GetString(Wpa_CodeArticle) = '') or (not wExistWAN(wGetCleWAN(GetString(Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle))))) then
      LastError := ErrArtNat
    { Gestion de la nature de travail de l'article inactive }
    else if not wExistWANActif(wGetCleWAN(GetString(Wpa_NatureTravail),wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle)))) then
      LastError := ErrNatArtInactive
    else if (GetString(Wpa_CodeArticleWNT) = '') then
      LastError := ErrArticleWNT
    else if (GetString(Wpa_CodeArticleWNT) = '') or not (wExistWAN(wGetCleWAN(GetString(Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticleWNT))))) then
      LastError := ErrArticleWNT
    else if not wWNTisOk(GetCleWNT) then
      LastError := ErrArticleWNT2
    else if (GetString(Wpa_Circuit) = '') or (not wExistQCI(GetString(Wpa_Circuit)) ) then LastError := ErrCodeCircuit
    {$ENDIF GPAO}
  end;
  {$ENDIF GPAOLIGHT}
  {$IFDEF GPAO}
  { Remplacement de composants }
  if EcranIs('WPA_CHGCOMPOSANT') and (Pos(sChangementComposant, GetString('WPA_CODEPARAM')) = 1) then
  begin

    if	    tRadioButton(GetControl('RBRIEN')).Checked         then SetString(REVISION, 'RIEN')
    else if tRadioButton(GetControl('RBMINVAL')).Checked       then SetString(REVISION, 'MINVAL')
    else if tRadioButton(GetControl('RBMAJVALPEROLD')).Checked then SetString(REVISION, 'MAJVALPEROLD');

    if FromApplyChangement then
    begin
      iCol := 1;
      bCancel := False;

      iCpt := 1;
      while (LastError = 0) and (iCpt < GRILLE.RowCount) do
      begin
        iRow := iCpt;
        if FromBtMoulinette then
          GRILLE_OnCellExit(GetControl('BTMOULINETTE'), iCol, iRow, bCancel)
        else
          GRILLE_OnCellExit(Self, iCol, iRow, bCancel);
        Inc(iCpt);
      end;

      if LastError = 0 then
        C := wEstUnFilsValide(iCol, iRow)
      else
        LastError := 0;
    end;
  end
  else if EcranIs('WPA_CHGVERSION') then
  begin
    if wExistUneVersion(GetCleWVS.article) and ((GetString('WPA_VARCHAR02') = '') or (not wExistVersion(GetCleWVS))) then
      LastError := ErrCodeVersion;
  end;
  {$ENDIF GPAO}
  {$IFDEF GPAOLIGHT}
  if GetString('WPA_CODEPARAM') = 'WPA_ORDRERECEP' then
  begin
    if GetDouble('WPA_DOUBLE07') < 0 then LastError := ErrQRecWOL
  end;
  if GetString('WPA_CODEPARAM') = 'WPA_ORDRELANC' then
  begin
    if GetDouble('WPA_DOUBLE03') < 0 then LastError := ErrQLanWOP
  end;
  {$ENDIF GPAOLIGHT}

  if copy(GetString('WPA_CODEPARAM'),1,10) = sTarifsRecherche then
  begin
    sItineraire := '';
    if      (GetString('WPA_VARCHAR03')<>'') and (not wExistGA(wGetArticleFromCodeArticle(GetString('WPA_VARCHAR03'))))                 then LastError := ErrArticle
    else if (GetString('WPA_VARCHAR02')<>'') and (sFonctionnalite<>sFraisAnnexes)  and (sFonctionnalite<>sCoutsIndirects) and (not ExistTiers('', GetString('WPA_VARCHAR02'))  )                               then LastError := ErrCodeTiers
    else if (GetString('WPA_VARCHAR02')<>'') and (sFonctionnalite=sFraisAnnexes) and (pos('TRA', sTypeFourni) > 0)        and (not ExistTiers('', GetString('WPA_VARCHAR02'), ' YTC_TYPEFOURNI LIKE "%TRA%"')) then LastError := ErrTransporteur
    else if (GetString('WPA_VARCHAR02')<>'') and (sFonctionnalite=sFraisAnnexes)                                          and (not ExistTiers('', GetString('WPA_VARCHAR02')))                                 then LastError := ErrCodeTiers
    else if (GetString('WPA_VARCHAR02')<>'') and (sFonctionnalite=sCoutsIndirects)                                        and (not ExistTiers('', GetString('WPA_VARCHAR02')) )                                                                        then LastError := ErrCodeTiers
    else if (GetString('WPA_VARCHAR01')<>'') and (not ExistTarifSpecial(GetString('WPA_COMBO01'),GetString('WPA_VARCHAR01')))           then LastError := ErrCodeTarifSpecial
    else if (GetString('WPA_VARCHAR06')<>'') and (not ExistTiers('CLI',GetString('WPA_VARCHAR06')))  then LastError := ErrCodeTiersLivre
    {$IFDEF AFFAIRE}
    else if (GetString('WPA_VARCHAR05')<>'') and (not Existeaffaire(GetString('WPA_VARCHAR05'), ''))                                    then LastError := ErrCodeAffaire
    {$ENDIF AFFAIRE}
    {$IFDEF GPAO}
      else if (GetString('WPA_VARCHAR07')<>'') and (                      not wExistQCI(GetString('WPA_VARCHAR07')))                    then LastError := ErrCircuit
      else if (GetString('WPA_VARCHAR07')<>'') and (sItineraire<>'') and (not wExistQCI(GetString('WPA_VARCHAR07'), sItineraire))       then LastError := ErrCircuitDeItineraire
      else if (GetString('WPA_VARCHAR08')<>'') and (not wExistQDE(GetString('WPA_VARCHAR08'), GetString('WPA_VARCHAR07')) )             then LastError := ErrPhaseduCircuit
      else if (GetString('WPA_VARCHAR12')<>'') and (not ExistQSI(GetString('WPA_VARCHAR12')))                                           then LastError := ErrAtelier
    {$ENDIF GPAO}
    ;
  end
  else if GetString('WPA_CODEPARAM') = 'WPA_GETPREVCOM' then
  begin
    if assigned(Ecran) then
    begin
//GP_20080507_DKZ_GP14995 Déb
      if (GetString('WPA_VARCHAR01')<>'') then
      begin
        Article := GetString('WPA_VARCHAR01');
        if wGetArticleFromCodeArticle(Article) <> Article then
          Article := wGetArticleFromCodeArticle(Article);
        if not wExistGA(Article) then
          LastError := ErrArticle;
      end;
//GP_20080507_DKZ_GP14995 Fin
      if (GetInteger('WPA_INTEGER01') <= 0) or (GetDouble('WPA_DOUBLE02') <= 0) then LastError := ErrValeurPositif;
      if GetDateTime('WPA_DATE02') < GetDateTime('WPA_DATE01') then LastError := ErrFourchetteDate;
    end
  end
//GP_DKZ_GP_PDC_20070927
  {$IFDEF GPAO}
  else if pos('WPDC_VUE', GetString('WPA_CODEPARAM')) > 0 then
  begin
    if assigned(Ecran) then
    begin
//GP_20080708_DKZ_GP15150 Déb
      if not GetBoolean('WPA_BOOLEAN09') then
      begin
        if GetInteger('WPA_DOUBLE09') < GetInteger('WPA_DOUBLE08') then LastError := ErrFourchetteHeure;
      end;
//GP_20080708_DKZ_GP15150 Fin
    end
  end
  {$ENDIF GPAO}
{$IFDEF QUALITE}
  else if (Assigned(ecran)) and (GetString('WPA_CODEPARAM')='RQGETCLOTURERQN') then
  begin
    if (GetString('WPA_VARCHAR01') <>'') and (not wExistRessource(GetString('WPA_VARCHAR01'),'','',True)) then
      LastError := ErrExistRessource
    else if (GetString('WPA_VARCHAR01') ='') then
      LastError := ErrRessource
    ;
  end
{$ENDIF QUALITE}
  {$IFDEF STK}
    else if GetString('WPA_CODEPARAM') = 'WPA_GETSTKPROJETE' then
    begin
      if GetInteger('WPA_INTEGER01') + GetInteger('WPA_INTEGER02') + GetInteger('WPA_INTEGER03') > 13 then LastError := ErrParamStkProjete
    end;
  {$ENDIF STK}
  ;

  Result := LastError = 0;
  MemoLastError := LastError;

  if LastError > 0 then
  begin
    LastErrorMsg := TexteMessage[LastError];
    if Assigned(Ecran) then
    begin
      { Remplacement de composant }
      if (Pos(sChangementComposant, GetString('WPA_CODEPARAM')) = 1) and EcranIs('WPA_CHGCOMPOSANT') then
      begin
        GRILLE.Col := iCol;
        GRILLE.Row := iRow;
        GRILLE_OnCellEnter(Self, iCol, iRow, bCancel);
        if Assigned(C) and (not (C is thGrid)) then
        begin
        if tPageControl(GetControl('PAGES')).ActivePage <> tTabSheet(GetControl('PGENERAL')) then
          tPageControl(GetControl('PAGES')).ActivePage := tTabSheet(GetControl('PGENERAL'));
        C.SetFocus;
        end;
      end
      { ------------------------- }
      {$IFDEF STK}
        else if GetString('WPA_CODEPARAM') = 'WPA_GETSTKPROJETE' then
        begin
          if LastError = ErrParamStkProjete then SetFocusControl('WPA_INTEGER01');
        end
      {$ENDIF STK}
        else if GetString('WPA_CODEPARAM') = 'WPA_GETPREVCOM' then
        begin
          if      LastError = ErrValeurPositif  then SetFocusControl('WPA_INTEGER01')
          else if LastError = ErrFourchetteDate then SetFocusControl('WPA_DATE01')
//GP_20080507_DKZ_GP14995
          else if LastError = ErrArticle        then SetFocusControl('WPA_VARCHAR01')
          ;
        end
      {$IFDEF GPAO}
        else if pos('WPDC_VUE', GetString('WPA_CODEPARAM')) > 0 then
        begin
//GP_20080708_DKZ_GP15150
          if LastError = ErrFourchetteHeure then SetFocusControl('WPA_DOUBLE08');
        end
      {$ENDIF GPAO}
      else
      begin
        Case LastError of
          ErrCodeArticle                   : SetFocusControl('WPA_VARCHAR01');
          ErrCodeNature                    : SetFocusControl('WPA_COMBO03')  ;
          ErrCodeItineraire                : SetFocusControl('WPA_VARCHAR03');
          ErrCodeCircuit                   : SetFocusControl('WPA_VARCHAR04');
          ErrCodeNomenclature              : SetFocusControl('WPA_VARCHAR05');
          ErrCodeIndiceMajeurNomenclature  : SetFocusControl('WPA_COMBO01')  ;
          ErrCodeGamme                     : SetFocusControl('WPA_VARCHAR06');
          ErrCodeIndiceMajeurGamme         : SetFocusControl('WPA_COMBO02')  ;
          ErrQLancement                    : SetFocusControl('WPA_DOUBLE03');
          ErrArticle                       : SetFocusControl('WPA_VARCHAR03');
          ErrCodeTiers, ErrTransporteur    : SetFocusControl('WPA_VARCHAR02');
          ErrCodeTarifSpecial              : SetFocusControl('WPA_VARCHAR01');
          ErrCodeTiersLivre                : SetFocusControl('WPA_VARCHAR06');
          {$IFDEF GPAO}
            ErrArticleWNT2                 : SetFocusControl(Wpa_CodeArticleWNT);
            ErrArtNat                      : SetFocusControl(Wpa_CodeArticle);
            { Gestion de la nature de travail de l'article inactive }
            ErrNatArtInactive              :  SetFocusControl(Wpa_CodeArticle);
          {$ENDIF GPAO}
          {$IFDEF GPAOLIGHT}
            ErrQRecWOL                     : SetFocusControl('WPA_DOUBLE07');
            ErrQLanWOP                     : SetFocusControl('WPA_DOUBLE03');
          {$ENDIF GPAOLIGHT}
          ErrCodeVersion                   : SetFocusControl('WPA_VARCHAR01');
          ErrRessource										 : SetFocusControl('WPA_VARCHAR01');
          ErrExistRessource								 : SetFocusControl('WPA_VARCHAR01');
          ErrMethodeValorisation           : SetFocusControl('WPA_COMBO04');
        end;
      end
      ;
    end;
  end;
end;

{$IFNDEF EAGLSERVER}
{$IFDEF GPAO}
{-------------------------------------------------------------------------------
   Lancement du Calcul du Prix de Revient
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.BTCalcul_OnClick(Sender: tObject);
var
  TobWPA, TobWRT : Tob; //Tob des paramètres
begin

//  vIdentifiantWPE:=0;
  if PgiAsk(TraduireMemoire('Confirmez-vous le lancement du calcul ?'),'') = mrYes then
  begin
    if not (DS.State in [dsInsert,dsEdit]) then
      DS.Edit;

    { Sauvegarde de la saisie}
    tToolBarButton97(GetControl('BVALIDER')).Click;

    if (MemoLastError=0) then
    begin
      { Chargement des paramètres de lancement }
      TobWPA := Tob.Create('WPARAM'    , nil, -1); // Paramètre de Lancement du calcul des prix de revient
      TobWRT := Tob.Create('_WPDRTYPE_', nil, -1);
      try
        wGetTobWPA(sPrixDeRevient, TobWPA);

        { Chargement des paramètres nécessaire depuis la table des types de prix de revient }
        if TobWRT.LoadDetailDBFromSql('WPDRTYPE', 'SELECT WRT_TYPEPDR, WRT_MAJDPADPR, WRT_PDRENREGWPL FROM WPDRTYPE WHERE WRT_TYPEPDR="'+TobWPA.GetString('WPA_VARCHAR07')+'"') then
        begin
          TobWPA.SetBoolean('WPA_BOOLEAN10', TobWRT.Detail[0].GetBoolean('WRT_MAJDPADPR'));
          TobWPA.SetBoolean('WPA_BOOLEAN16', TobWRT.Detail[0].GetBoolean('WRT_PDRENREGWPL'));
        end;

        {$IFDEF GPAO}
          wCalculPrixDeRevient('EcranParamètre', TobWPA, False, iIdentifiantWPE, 'ORIGINEAPPEL='+sAppel);
        {$ENDIF GPAO}
      finally
        TobWPA.Free;
        TobWRT.Free;
      end;
    end;
  end;
end;
{$ENDIF GPAO}
{$ENDIF !EAGLSERVER}

{-------------------------------------------------------------------------------
--------------------------------------------------------------------------------}
{$IFDEF GPAOLIGHT}
procedure TOM_WPARAM.SetConversion;
begin
  SetDouble('WPA_DOUBLE05', GetDouble('WPA_DOUBLE01') * GetDouble('WPA_DOUBLE09'));
  SetDouble('WPA_DOUBLE06', GetDouble('WPA_DOUBLE02') * GetDouble('WPA_DOUBLE09'));
  SetDouble('WPA_DOUBLE07', GetDouble('WPA_DOUBLE03') * GetDouble('WPA_DOUBLE09'));
  SetDouble('WPA_DOUBLE08', GetDouble('WPA_DOUBLE04') * GetDouble('WPA_DOUBLE09'));
end;
{$ENDIF GPAOLIGHT}

{--------------------------------------------------------------------------------
   En fonction  de la nature du tiers personnalisation des listes de choix et des libellés
---------------------------------------------------------------------------------}
procedure TOM_WPARAM.ClientOuFournisseur;
begin
  SetControlCaption('TWPA_VARCHAR02', LibelleParFonctionnalite(sFonctionnalite, 'TIERS', sTypeFourni));

  if sFonctionnalite = sFraisAnnexes then
  begin
    SetControlProperty('WPA_VARCHAR02', 'DataType', 'GCTIERS');
    SetControlProperty('WPA_VARCHAR01', 'PLUS', 'YTP_FONCTIONNALITE="'+GetString('WPA_COMBO01')+'"');
    SetControlProperty('WPA_VARCHAR02', 'PLUS', ' AND T_AUXILIAIRE IN (SELECT YTC_AUXILIAIRE FROM TIERSCOMPL WHERE YTC_TYPEFOURNI LIKE "%'+sTypeFourni+'%")');
  end
  else if sFonctionnalite = sCoutsIndirects then
  begin
    SetControlProperty('WPA_VARCHAR02', 'DataType', 'GCTIERS');
    SetControlProperty('WPA_VARCHAR01', 'PLUS', 'YTP_FONCTIONNALITE="'+GetString('WPA_COMBO01')+'"');
  end
  else if (sFonctionnalite=sTarifFournisseur) or (sFonctionnalite=sTarifSousTraitantAchat) or (sFonctionnalite=sTarifSousTraitantPhase) then
  begin
    SetControlProperty('WPA_VARCHAR02', 'DataType', 'GCTIERS');
    SetControlProperty('WPA_VARCHAR02', 'PLUS', 'T_NATUREAUXI="'+sNatureAuxiFournisseur+'"');
    SetControlProperty('WPA_VARCHAR01', 'PLUS', 'YTP_FONCTIONNALITE="'+GetString('WPA_COMBO01')+'"');
  end
  else
  begin
    SetControlProperty('WPA_VARCHAR02', 'DataType', 'GCTIERS');
    SetControlProperty('WPA_VARCHAR02', 'PLUS', 'T_NATUREAUXI="'+sNatureAuxiClient+'" OR T_NATUREAUXI="'+sNatureAuxiProspect+'"');
    SetControlCaption('TWPA_VARCHAR02', TraduireMemoire('Client'));
    SetControlProperty('WPA_VARCHAR01', 'PLUS', 'YTP_FONCTIONNALITE="'+GetString('WPA_COMBO01')+'"');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Affecte les propriété Enabled et Color a un Control en
Suite ........ : fonction d'une condition de SAY
Mots clefs ... : 
*****************************************************************}
procedure TOM_WPARAM.SetSayOuGet(sControl: string; lConditionSay: boolean);
begin {SetSayOuGet}
  if (lConditionSay) then
  begin
    SetControlProperty(sControl, 'Enabled', False);
    SetControlProperty(sControl, 'Color'  , ClBtnFace);
  end
  else
  begin
    if (sControl='WPA_COMBO02') then
      SetControlProperty(sControl, 'Enabled', not VH^.ProfilUserC[PrEtablissement].ForceDepot)
    else
      SetControlProperty(sControl, 'Enabled', True);
    SetControlProperty(sControl, 'Color'  , ClWindow);
  end
end; {SetSayOuGet}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Controle l'affichage et positionne le filtre sur la tablette des
Suite ........ : zones Marque Choixqualité et IndiceArticle dépendant de
Suite ........ : l'article saisi
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.CtrlAffichageMarqueChoixQualiteIndice;
var
  {$IFDEF STK}
    lActif   : boolean;
    lVisible : boolean;
  {$ENDIF STK}
  sArticle : string;

begin {CtrlAffichageMarqueChoixQualiteIndice}
  sArticle := wGetArticleFromCodeArticle(GetString('WPA_VARCHAR03'));
  { Marque }
  {$IFDEF GPAO}
  if StkGereMarque then
  begin
    lVisible := ((sArticle='') or GereMarque(sArticle)) and (sParametre[iPosMarque]=wTrue);
    SetControlVisible('GMARQUE', lVisible);
    if (not lVisible) then
      SetString('WPA_COMBO06', '')
    else
    begin
      SetControlVisible('WPA_COMBO06' , lVisible);
      SetControlVisible('TWPA_COMBO06', lVisible);
      SetControlProperty('WPA_COMBO06', 'PLUS', MyGetPlusGM(GetFluxFromFonctionnalite(sFonctionnalite), GetString('WPA_VARCHAR02'), sArticle));
      lActif   := (sParametre[iPosMarque]=wTrue);
      SetSayOuGet('WPA_COMBO06', (not lActif));
    end;
  end;
  {$ENDIF GPAO}

  { ChoixQualite }
  {$IFDEF GPAO}
  if StkGereChoixQualite then
  begin
    lVisible := ((sArticle='') or GereChoixQualite(sArticle))and (sParametre[iPosChoixQualite]=wTrue);
    SetControlVisible('GCHOIXQUALITE', lVisible);
    if (not lVisible) then
      SetString('WPA_COMBO08', '')
    else
    begin
      SetControlVisible('WPA_COMBO08' , lVisible);
      SetControlVisible('TWPA_COMBO08', lVisible);
      SetControlProperty('WPA_COMBO08', 'PLUS', GetPlusGCQ(GetFluxFromFonctionnalite(sFonctionnalite), GetString('WPA_VARCHAR02'), sArticle));
      lActif   := (sParametre[iPoschoixQualite]=wTrue);
      SetSayOuGet('WPA_COMBO08', (not lActif));
    end;
  end;
  {$ENDIF GPAO}

  { Indice article }
  {$IFDEF STK}
  if StkGereIndiceArticle then
  begin
    lVisible := ((sArticle='') or GereIndiceArticle(sArticle)) and (sParametre[iPosIndiceArticle]=wTrue);
    SetControlVisible('GINDICEARTICLE', lVisible);
    if (not lVisible) then
      SetString('WPA_COMBO09', '')
    else
    begin
      SetControlVisible('WPA_COMBO09' , lVisible);
      SetControlVisible('TWPA_COMBO09', lVisible);
      SetControlProperty('WPA_COMBO09', 'PLUS', 'GIA_ARTICLE="' + sArticle + '" AND GIA_ETATREV="VAL"');
      lActif   := (sParametre[iPosIndiceArticle]=wTrue);
      SetSayOuGet('WPA_COMBO09', (not lActif));
    end;
  end;
  {$ENDIF STK}
end; {CtrlAffichageMarqueChoixQualiteIndice}

{-------------------------------------------------------------------------------
   Lancement de la Recherche Tarifaire
--------------------------------------------------------------------------------}
{$IFNDEF EAGLSERVER}
procedure TOM_WPARAM.BTRechercheTarifaire_OnClick(Sender: tObject);
var
  iCpt : integer;
  lConditionsApplications : boolean;
  TobRemises, TobTiers, TobArticle  : tob;
  aPort: MyArrayValue;

  {---------------------------------------------------------------------------
  ----------------------------------------------------------------------------}
  procedure GetTobContexte;
  begin {GetTobContexte}
    {Initialisation de la TobContexte par les valeur par défaut de chaque champ}
    InitTobContexte(TobContexte);

    TobContexte.SetString('FONCTIONNALITE', GetString('WPA_COMBO01')  );
    TobContexte.SetString('CODEPORT'      , GetString('WPA_VARCHAR15'));
    TobContexte.SetString('NATUREAUXI'    , NatureAuxiliaire(GetString('WPA_COMBO01')));
    TobContexte.SetString('DEPOT'         , GetString('WPA_COMBO02')  );
    TobContexte.SetString('DEVISETARIF'   , GetString('WPA_COMBO03')  );
    TobContexte.SetString('DEVISEPIECE'   , GetString('WPA_COMBO03')  );

    TobContexte.SetString('TARIFSPECIAL'  , GetString('WPA_VARCHAR01'));
    TobContexte.SetString('TIERS'         , GetString('WPA_VARCHAR02'));
    TobContexte.SetString('SECTEURGEO'    , GetString('WPA_COMBO04')  );
    TobContexte.SetString('CODEPOSTAL'    , GetString('WPA_VARCHAR10'));
    TobContexte.SetString('PAYS'          , GetString('WPA_COMBO05')  );
    TobContexte.SetString('REGION'        , GetString('WPA_VARCHAR11'));
    TobContexte.SetString('MODEEXP'       , GetString('WPA_COMBO07')  );
    TobContexte.SetString('AFFAIRE'       , GetString('WPA_VARCHAR05'));
    TobContexte.SetString('ATELIER'       , GetString('WPA_VARCHAR12'));

    TobContexte.SetString('NATURETRAVAIL' , GetString('WPA_VARCHAR14'));

    if (GetString('WPA_VARCHAR16')<>'') and (GetString('WPA_VARCHAR03')=wGetCodeArticleFromArticle(GetString('WPA_VARCHAR16'))) then
      TobContexte.SetString('ARTICLE', GetString('WPA_VARCHAR16'))
    else
      TobContexte.SetString('ARTICLE', wGetArticleFromCodeArticle(GetString('WPA_VARCHAR03')));

    TobContexte.SetString('MARQUE'        , GetString('WPA_COMBO06')  );
    TobContexte.SetString('CHOIXQUALITE'  , GetString('WPA_COMBO08')  );
    TobContexte.SetString('INDICEARTICLE' , GetString('WPA_COMBO09')  );
    TobContexte.SetString('RESSOURCE'     , GetString('WPA_VARCHAR04'));
    TobContexte.SetString('CIRCUIT'       , GetString('WPA_VARCHAR07'));
    TobContexte.SetString('PHASE'         , GetString('WPA_VARCHAR08'));
    TobContexte.SetString('TRAITEMENT'    , GetString('WPA_VARCHAR09'));
    TobContexte.SetDateTime('DATEPIECE'     , GetDateTime('WPA_DATE01') );
    TobContexte.SetDateTime('DATELIVRAISON' , GetDateTime('WPA_DATE01') );
    TobContexte.SetDateTime('DATEEXTERNE'   , GetDateTime('WPA_DATE01') );
    TobContexte.SetDouble('QUANTITE'      , GetDouble('WPA_DOUBLE09') );
  end; {GetTobcontexte}

begin
  aPort:=nil;
  NextPrevControl(Ecran);
  if (DS.State in [dsInsert,dsEdit]) then
     { Sauvegarde de la saisie}
     tToolBarButton97(GetControl('BVALIDER')).Click;

  if (MemoLastError=0) then
  begin
    {Si appel = Recherche alors chargement de la TobContexte, Recherche des tarifs et chargement de la TobTarif                   }
    {                     sinon la TobContexte est passée par TransfertTobs, la TobTarif est crée à partir de la TobUneLigneTarif }
    TobTarif.ClearDetail; TobPrixMarge.ClearDetail; TobAffichage.ClearDetail;
    if (sAppel='RECHERCHE') then
    begin
      GetTobContexte;
      lConditionsApplications := GetParamSoc('SO_TARIFCONDAPPL');
      TobTiers   := Tob.Create('_TIERS_', nil, -1);
      TobArticle := Tob.Create('_ARTICLE_', nil, -1);
      try
        //Si gestion des conditions particuliaires d'application alors chargement de la  TobTiers et de la TobArticle
        if (lConditionsApplications) then
        begin
          wGetTobGA('*', whereGA(TobContexte.GetString('ARTICLE')), TobArticle, True);
          GetTobTiers('*', whereTiers('',TobContexte.GetString('TIERS')), TobTiers, True);
        end;
        { Recherche du système tarifaire }
        RechercheTarifs('RECHERCHE', TobContexte, TobTarif, TobPrixMarge, TobTiers, TobArticle, nil, nil, True);
        tPageControl(GetControl('PAGES')).ActivePage := tTabSheet(GetControl('PRESULTATS'));
      finally
        TobTiers.Free;
        TobArticle.Free;
      end;
    end
    else if (Pos(sAppel,'PIECE/LIGNE')>0) then
    begin
      //A Partir de TobUneLigneTarif génération de la TobTarif (Mère + Filles)
      TobTarifAddChampsSupValeur('', TobTarif );
      TobTarifAddChampsSupValeur('', TobPrixMArge );

      TobTarif.SetDouble('REMISETIERS_VAL', TobContexte.GetDouble('REMISEPIED'));

      for iCpt:=0 to (TobUneLigneTarif.Detail.Count-1) do
      begin
        if     (TobUneLigneTarif.Detail[iCpt].GetString('GLT_FONCTIONNALITE')= sFonctionnalite)
           and (TobUneLigneTarif.Detail[iCpt].GetString('GLT_CODEPORT')      = sCodePort)
           and (TobUneLigneTarif.Detail[iCpt].GetString('GLT_TIERSFRAIS')    = sTiersFrais)
        then
        begin
          if      (TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG')=sRang+'100') then
          begin
            TobTarif.SetDouble('FORFAIT_VAL',TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobTarif.SetString('FORFAIT_OU' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
          end
          else if (TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG')=sRang+'200') then
          begin
            TobTarif.SetDouble('COUTLANCEMENT_VAL',TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobTarif.SetString('COUTLANCEMENT_OU' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
          end
          else if (TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG')=sRang+'300') then
          begin
            TobTarif.SetDouble('REMISEMONTANT_VAL',TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobTarif.SetString('REMISEMONTANT_OU' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
          end
          else if (TobUneLigneTarif.Detail[iCpt].GetValue('GLT_RANG')=sRang+'800') then
          begin
            TobTarif.PutValue('REMISETIERS_VAL',TobUneLigneTarif.Detail[iCpt].GetValue('GLT_VALEUR'));
          end
          else if (TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG')=sRang+'400') then
          begin
            TobTarif.SetDouble('PRIXBRUT_VAL',TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobTarif.SetString('PRIXBRUT_OU' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
          end
          else if (TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG')=sRang+'900') then
          begin
            TobTarif.SetDouble('PRIXMARGE_VAL',TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
          end
          else if (TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG')=sRang+'700') then
          begin
            TobTarif.SetDouble('PRIXNET_VAL',TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobTarif.SetString('PRIXNET_OU' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
          end
          else if (copy(TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG'),1,2)=sRang+'5') then //Remise Libre
          begin
//GP20080228 GP14772 MM
            TobRemises := CreationTobRemises(TobTarif, nil);
            TobRemises.SetDouble('REMISE1_VAL'      ,TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobRemises.SetDouble('REMISE2_VAL'      ,0);
            TobRemises.SetDouble('REMISE3_VAL'      ,0);
            TobRemises.SetString('REMISES_OU'       ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
            TobRemises.SetString('REMISES_INTITULE' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_QUOI'));
          end
          else if (copy(TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG'),1,2)=sRang+'6') then //Remise Système
          begin
//GP20080228 GP14772 MM
            TobRemises := CreationTobRemises(TobTarif, nil);
            TobRemises.SetDouble('REMISE1_VAL'      ,TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobRemises.SetDouble('REMISE2_VAL'      ,0);
            TobRemises.SetDouble('REMISE3_VAL'      ,0);
            TobRemises.SetString('REMISES_OU'       ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
            TobRemises.SetString('REMISES_INTITULE' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_QUOI'));
          end;
        end
        else if (TobUneLigneTarif.Detail[iCpt].GetString('GLT_FONCTIONNALITE')=sPrixVenteMarge) then
        begin
          if      (TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG')=sRang+'100') then
          begin
            TobPrixMarge.SetString('PRIXDEBASE_PARAM', TobUneLigneTarif.Detail[iCpt].GetString('GLT_QUOI'));
            TobPrixMarge.SetDouble('PRIXDEBASE_VAL'  , TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobPrixMarge.SetString('PRIXDEBASE_OU'   , TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
          end
          else if (copy(TobUneLigneTarif.Detail[iCpt].GetString('GLT_RANG'),1,2)=sRang+'6') then
          begin
//GP20080228 GP14772 MM
            TobRemises := CreationTobRemises(TobPrixMarge, nil);
            TobRemises.SetDouble('REMISE1_VAL'      ,TobUneLigneTarif.Detail[iCpt].GetDouble('GLT_VALEUR'));
            TobRemises.SetDouble('REMISE2_VAL'      ,0);
            TobRemises.SetDouble('REMISE3_VAL'      ,0);
            TobRemises.SetString('REMISES_OU'       ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_OU'));
            TobRemises.SetString('REMISES_INTITULE' ,TobUneLigneTarif.Detail[iCpt].GetString('GLT_QUOI'));
          end;
        end;
      end;
    end;

    { Affichage de l'origine du système tarifaire trouvé }
    { Depuis le menu de recherche, on ne connait pas encore les info relatives au port, il faut aller les rechercher }
    if (sAppel='RECHERCHE') and (sTypeFourni = '') then
    begin
      aPort:= WGetSqlFieldsValues(['GPO_TYPEFOURNI','GPO_REPARTITION'], 'PORT', 'GPO_CODEPORT="'+TobContexte.GetString('CODEPORT')+'"');
      if aPort <> nil then
      begin
        sTypeFourni  := aPort[0];
        sRepartition := aPort[1];
      end
      else
      	sRepartition := 'Q';
    end;
    AffichageOrigineSystemeTarifaire(sAppel, sTypeFourni, sRepartition, TobContexte, TobTarif, TobPrixMarge, TobAffichage, GAFFICHAGE, MBLOCNOTE);

    if Assigned(TobTarif) and (sAppel='RECHERCHE') then
    begin
      //Pour passer DS.state en mode dsEdit sinon SetString ne fonctionne pas
      if not(DS.State in [dsInsert,dsEdit]) then
        DS.edit;
      SetDouble('WPA_DOUBLE01',TobTarif.GetDouble('FORFAIT_VAL'))          ; SetString('WPA_LONGVARCHAR01',TobTarif.GetString('FORFAIT_OU'));
      SetDouble('WPA_DOUBLE02',TobTarif.GetDouble('COUTLANCEMENT_VAL'))    ; SetString('WPA_LONGVARCHAR02',TobTarif.GetString('COUTLANCEMENT_OU'));
      SetDouble('WPA_DOUBLE03',TobTarif.GetDouble('PRIXBRUT_VAL'))         ; SetString('WPA_LONGVARCHAR03',TobTarif.GetString('PRIXBRUT_OU'));
      if (TobTarif.Detail.Count <>0) then
      begin
        SetDouble('WPA_DOUBLE04',TobTarif.Detail[0].GetDouble('REMISE1_VAL')); SetString('WPA_LONGVARCHAR04',TobTarif.Detail[0].GetString('REMISES_OU'));
        SetDouble('WPA_DOUBLE05',TobTarif.Detail[0].GetDouble('REMISE2_VAL'));
        SetDouble('WPA_DOUBLE06',TobTarif.Detail[0].GetDouble('REMISE3_VAL'));
      end;
      SetDouble('WPA_DOUBLE07',TobTarif.GetDouble('PRIXNET_VAL'))          ; SetString('WPA_LONGVARCHAR07',TobTarif.GetString('PRIXNET_OU'));

      if (DS.State in [dsInsert,dsEdit]) then
        tToolBarButton97(GetControl('BVALIDER')).Click;
    end;
  end;
end;
{$ENDIF !EAGLSERVER}

{-------------------------------------------------------------------------------
   Affichage des couleurs des cellules
--------------------------------------------------------------------------------}
{$IFNDEF EAGLSERVER}
procedure TOM_WPARAM.GAffichageCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
  Rect: TRect;
  GAffichage: THgrid;
begin
  GAffichage := THGrid(GetControl('GAFFICHAGE'));
  Rect := GAffichage.CellRect(Acol,Arow);
  wAffichageDesTarifsGrille(ACol, ARow, Rect, Canvas, gdSelected in aState);
end;
{$ENDIF !EAGLSERVER}

{-------------------------------------------------------------------------------
   Affichage des compléments d'une lignes de la grille d'affichage
--------------------------------------------------------------------------------}
{$IFNDEF EAGLSERVER}
procedure TOM_WPARAM.GAffichageOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  sParametre : string;
begin
  sParametre := GetParametreFromYFO(sFonctionnalite, '', 'TIE', 'EntitéActive');
  AffichageCommentaireTarif(sAppel, sFonctionnalite, sTypeFourni, sParametre, TobContexte, TobAffichage, GAffichage, MBLOCNOTE);
end;
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GPAO}
{-------------------------------------------------------------------------------
   checkbox : par défaut ?
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.WPA_BOOLEAN01_OnClick(Sender: tObject);
begin
  NextPrevControl(Ecran);
  if (l1erClickBoolean01) then
    AffichageParDefaut( not GetBoolean('WPA_BOOLEAN01'))
  else
    AffichageParDefaut( GetBoolean('WPA_BOOLEAN01'));
  l1erClickBoolean01 := False;
end;
{$ENDIF GPAO}
{$ENDIF !EAGLSERVER}

{$IFDEF GPAO}
{-------------------------------------------------------------------------------
   Affichage en fonction si le contexte est le contexte par defaut de l'article ou non
--------------------------------------------------------------------------------}
procedure TOM_WPARAM.AffichageParDefaut(lDefaut: boolean);
begin
  SetSayOuGet('WPA_VARCHAR03',lDefaut);  //Itinéraire  : Non modifiable car pour le moment un seul itinéraire par article
  SetSayOuGet('WPA_VARCHAR04',lDefaut);  //Circuit
  SetSayOuGet('WPA_VARCHAR05',lDefaut);  //Nomenclature
  SetSayOuGet('WPA_COMBO01'  ,lDefaut);  //Indice Majeur Nomenclature
  SetSayOuGet('WPA_VARCHAR06',lDefaut);  //Gamme
  SetSayOuGet('WPA_COMBO02'  ,lDefaut);  //Indice Majeur Gamme
  SetSayOuGet('WPA_DOUBLE01' ,lDefaut);  //Prix de Revient pour une qté de
  SetSayOuGet('WPA_DOUBLE02' ,lDefaut);  //Qté économique de production
  SetSayOuGet('WPA_DOUBLE03' ,lDefaut);  //Prix de Revient exprimé pour une qté de
end;
{$ENDIF GPAO}

{-------------------------------------------------------------------------------------
   Permet de colorer les cellules des informations ayant participées à la recherche
--------------------------------------------------------------------------------------}
{$IFNDEF EAGLSERVER}
procedure TOM_WPARAM.wAffichageDesTarifsGrille(iCol, iRow: Integer; Rect: Trect; Canvas: TCanvas; lFocus : boolean);
var
  iLeftImage, iTopImage: integer;
  G   : THGrid;
  Val : String;
begin
  G := THGrid(GetControl('GAFFICHAGE'));
  Val := G.CellValues[iCol, iRow];

  if Pos('IMAGE_', Val) <> 0 then
  begin
    if      (Val=sImageMain) then
      BMP := tImage(GetControl('IMAGE_MAIN')).Picture.Bitmap
    else if (Val=sImageSommePleine) then
      BMP := tImage(GetControl('IMAGE_SOMMEPLEINE')).Picture.Bitmap
    else if (Val=sImageSommeVide) then
      BMP := tImage(GetControl('IMAGE_SOMMEVIDE')).Picture.Bitmap
    else if (Val=sImageRecherchePleine) then
      BMP := tImage(GetControl('IMAGE_RECHERCHEPLEINE')).Picture.Bitmap
    else if (Val=sImageRechercheVide) then
      BMP := tImage(GetControl('IMAGE_RECHERCHEVIDE')).Picture.Bitmap
    else if (Val=sImageSommeRecherchePleine) then
      BMP := tImage(GetControl('IMAGE_SOMMERECHERCHEPLEINE')).Picture.Bitmap
    else if (Val=sImageSommeRechercheVide) then
      BMP := tImage(GetControl('IMAGE_SOMMERECHERCHEVIDE')).Picture.Bitmap
    else
      BMP := tImage(GetControl('IMAGE_VIDE')).Picture.Bitmap
    ;

      if (lFocus) then
      begin
        if (val<>sImageVide) then
          Canvas.Brush.Color := wEclaircirCouleur(ClHighLight, 200)
        else
          Canvas.Brush.Color := ClHighLight;
      end;

    iLeftImage := Rect.Left + (G.ColWidths[iCol]  - BMP.Width ) div 2;
    iTopImage  := Rect.Top  + (G.RowHeights[iRow] - BMP.Height) div 2;

    Canvas.Fillrect(Rect);
    Canvas.Draw(iLeftImage, iTopImage, BMP);
  end;
end;
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.bFerme_Click(Sender: tObject);
begin
  { Fermeture sans sauvegarde }
  if (ecran.Name='WPA_GETAFCONTEXTE') then
  begin
    DS.Cancel;
    tfFiche(Ecran).Retour:=wFalse;
  end
  {$IFDEF QUALITE}
    else if (EcranIs('RQGETCLOTURERQN')) and (tfFiche(Ecran).Retour<>'X') then
      tfFiche(Ecran).Retour:='FERME'
  {$ENDIF QUALITE}
  else if EcranIs('WPA_BLOCNOTE') and WithCancelInfo and (tfFiche(Ecran).Retour<>'X') then
    tfFiche(Ecran).Retour:='-'
  ;

  tfFiche(Ecran).BFermeClick(Sender);
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IF not(Defined(EAGLSERVER) or Defined(ERADIO))}
{$IFDEF GPAOLIGHT}
procedure TOM_WPARAM.BTCalculFaisabilite_OnClick(Sender: tObject);
Var
  Params: string;

  procedure CalculFaisabilite;
  var
    TobWPA: Tob;
  begin
    TobWPA := Tob.Create(GetTableName, nil, -1);
    try
      TobWPA.SetString('WPA_CODEPARAM', GetString('WPA_CODEPARAM'));
      TobWPA.SetString('WPA_UTILISATEUR', V_Pgi.User);
      TobWPA.LoadDB;
      wCalculFaisabilite(TobWPA);
    finally
      TobWPA.free;
    end;
  end;
begin
  tToolBarButton97(Sender).Enabled := false;
  try
    { Sauvegarde }
    tToolBarButton97(GetControl('BVALIDER')).Click;
    {$IFDEF GPAOLIGHT}  { Assemblage si plusieurs GNE pour l'article }
      {$IFNDEF GPAO}
        ShowCodeNomenAss;
      {$ENDIF GPAO}
    {$ENDIF GPAOLIGHT}

    if MemoLastError = 0 then
    begin
      { Lancement du calcul }
      CalculFaisabilite;

      { Résultat du calcul }
      Params:='NATURETRAVAIL=' + GetString(wfaisabilite.Wpa_NatureTravail) + ';CODEARTICLE=' + GetString(wfaisabilite.Wpa_CodeArticle)
           + ';CODEARTICLEWNT='+ GetString(wfaisabilite.Wpa_CodeArticleWNT)
           + ';MAJEURWNT=' + GetString(wfaisabilite.Wpa_MajeurWNT)
           + ';DATEDEC=' + GetString(wfaisabilite.Wpa_DateDec)
           + ';QLOTSAIS=' + FloatToStr(GetDouble(wfaisabilite.Wpa_QlotSais))
           + ';UNITELOT=' + GetString(wfaisabilite.Wpa_UniteLot)
           + ';CODITI=' + GetString(wfaisabilite.Wpa_CodIti)
           + ';CIRCUIT=' + GetString(wfaisabilite.Wpa_Circuit)
           + ';PFSTPHY=' + GetString(wfaisabilite.Wpa_Param1)
           + ';PFSTRES=' + GetString(wfaisabilite.Wpa_Param2)
           + ';SESTPHY=' + GetString(wfaisabilite.Wpa_Param3)
           + ';SESTRES=' + GetString(wfaisabilite.Wpa_Param4)
           + ';COSTRES=' + GetString(wfaisabilite.Wpa_Param5)
           + ';REGROU=' + GetString(wfaisabilite.Wpa_Param6)
           ;
      wCallResultatFaisabilite(Params);
    end;
  finally
    tToolBarButton97(Sender).Enabled := true;
  end;
end;
{$ENDIF GPAOLIGHT }
{$IFEND !EAGLSERVER OR !ERADIO}

{$IFDEF GPAOLIGHT}
procedure TOM_WPARAM.SetCodeArticleWNT;
var
{$IFDEF GPAO}
	CleWNT: tCleWNT;
{$ELSE GPAO}
  CodeNomenASS: String;

  procedure SetMajeurWNTFromNomenAss;
  var
    T: Tob;
    Sql: String;
  begin
    if (CodeNomenAss <> '') then
    begin
      Sql := 'SELECT WNT_MAJEUR '
           + 'FROM WNOMETET '
           + 'WHERE WNT_IDENTIFIANT=ISNULL((SELECT GNE_IDENTIFIANTWNT '
                                           + 'FROM NOMENENT '
                                           + 'WHERE GNE_ARTICLE="' + wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle)) + '" '
                                           +  ' AND GNE_NOMENCLATURE="' + CodeNomenASS + '"), 0)'
           ;
      T := Tob.Create('', nil, -1);
      try
        T.LoadDetailDBFromSQL('', Sql);
        if T.Detail.Count > 0 then
          SetString(Wpa_MajeurWNT, T.Detail[0].GetString('WNT_MAJEUR'));
      finally
        T.Free;
      end;
    end;
  end;

{$ENDIF GPAO}
begin
  {$IFDEF GPAO}
    CleWNT := wGetActiveCleWNT(GetString(Wpa_natureTravail), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle)));
    SetString(Wpa_CodeArticleWNT, wGetCodeArticleFromArticle(CleWNT.Article));
    SetString(Wpa_MajeurWNT, CleWNT.Majeur);
  {$ELSE GPAO} { Assemblage }
    SetString(Wpa_CodeArticleWNT, GetString(Wpa_CodeArticle));
    SetString(Wpa_MajeurWNT, wGetActifMajeur('WNT', GetString(Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle))));
    SetString(Wpa_GNENomenclature, '');
    if GetString(Wpa_MajeurWNT) <> 'A' then { Il y a plusieurs nomenclatures d'assemblage différentes }
    begin
      { On choisi une nomenclature GNE }
      CodeNomenASS := ChoixNomenclature(wGetArticleFromCodeArticle(GetString(Wpa_CodeArticle)));
      if CodeNomenAss <> '' then
      begin
        SetString(Wpa_GNENomenclature, CodeNomenAss);
        SetMajeurWNTFromNomenAss;
      end;
    end;
    if Assigned(Ecran) then
      ShowCodeNomenAss;

  {$ENDIF GPAO}
  SetQWNT;
end;
{$ENDIF GPAOLIGHT }

{$IFNDEF EAGLSERVER}
{$IFDEF GPAOLIGHT}
procedure TOM_WPARAM.ShowCodeNomenAss;
var
  TobGNE: Tob;
  Sql: String;
begin
  SetControlText(Wpa_CodeArticleWNT, ''); SetControlCaption('LBCODEARTICLEWNT', '');

  TobGNE := Tob.Create('', nil, -1);
  try
    Sql := 'SELECT GNE_NOMENCLATURE,GNE_LIBELLE '
         + 'FROM NOMENENT '
         + 'WHERE GNE_IDENTIFIANTWNT=(SELECT WNT_IDENTIFIANT '
                                  +  'FROM WNOMETET '
                                  +  'WHERE (WNT_NATURETRAVAIL="' + GetString(Wpa_NatureTravail) + '") '
                                  +    'AND (WNT_ARTICLE="' + wGetArticleFromCodeArticle(GetString(Wpa_CodeArticleWNT)) + '") '
                                  +    'AND (WNT_MAJEUR="' + GetString(Wpa_MajeurWNT) + '")'
                                  +  ')'
        ;
    if TobGNE.LoadDetailFromSQL(Sql) then
    begin
      SetControlText(Wpa_CodeArticleWNT, TobGNE.Detail[0].GetString('GNE_NOMENCLATURE'));
      SetControlCaption('LBCODEARTICLEWNT', TobGNE.Detail[0].GetString('GNE_LIBELLE'));
    end;
  finally
    TobGNE.Free;
  end;
end;
{$ENDIF GPAOLIGHT}
{$ENDIF !EAGLSERVER}


{$IFDEF GPAOLIGHT}
function TOM_WPARAM.GetCleWNT: tCleWNT;
begin
  Result := wNomeTet.GetCleWNT(GetString(Wpa_NatureTravail), wGetArticleFromCodeArticle(GetString(Wpa_CodeArticleWNT)), GetString(Wpa_MajeurWNT))
end;
{$ENDIF GPAOLIGHT }

{$IFDEF GPAOLIGHT}
procedure TOM_WPARAM.SetQWNT;
begin
  if GetCleWNT.Article <> '' then
  begin
    SetString('WPA_COMBO04', wGetFieldFromWNT('WNT_UNITELOT', GetCleWNT));
    SetString('WPA_COMBO05', wGetFieldFromWNT('WNT_QUALIFUNITESTO', GetCleWNT));
    SetString('WPA_DOUBLE01', wGetFieldFromWNT('WNT_QLOTSAIS', GetCleWNT));
    SetString('WPA_DOUBLE02', wGetFieldFromWNT('WNT_QLOTSTOC', GetCleWNT));
    SetString('WPA_DOUBLE03', wGetFieldFromWNT('WNT_COEFLOT', GetCleWNT));
  end;
end;
{$ENDIF GPAOLIGHT }

{---------------------------------METHODES POUR LE REMPLACEMENT DE COMPOSANTS------------------------------------------}

{-------------------------------------      GESTION DE LA GRILLE      -------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 03/01/2003
Modifié le ... :   /  /
Description .. : Gère le click de l'Elipsis
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
procedure TOM_WPARAM.Cell_OnElipsisClick(Sender: TObject);
begin
  if GetParamSocSecur('SO_GCRECHARTAV', False) then
  begin
    SetControlText('CODECOMPOSANT', GRILLE.Cells[GRILLE.Col, GRILLE.Row]);
    DispatchRecherche(ThCritMaskEdit(GetControl('CODECOMPOSANT')), 1, '', 'RECHERCHEARTICLE;RETOUR_CODEARTICLE=X;GA_CODEARTICLE=' + GRILLE.Cells[GRILLE.Col, GRILLE.Row], NomMulRechArticle);
    GRILLE.Cells[GRILLE.Col, GRILLE.Row] := GetControlText('CODECOMPOSANT');
  end
  else
    LookupList(GRILLE, TraduireMemoire('Codes composants'), 'ARTICLE', 'GA_CODEARTICLE', 'GA_LIBELLE', '', 'GA_CODEARTICLE', True, 0);
end;
{$ENDIF GPAO }
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Place le composant de saisie sur la bonne cellule
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.GRILLE_OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
	CellValue: String;
begin
	if ARow > 0 then
   begin
    { MàJ de l'affichage des données de lien }
    FromPutEcran := True;
    TobWPA.Detail[GRILLE.Row - 1].PutEcran(Ecran, tWinControl(GetControl('GBQLIEN')));
    FromPutEcran := False;

    GRILLE.ElipsisButton := True;
    GRILLE.EditorMode := True;
    CellValue := GRILLE.Cells[GRILLE.Col, GRILLE.Row];
    if GRILLE.Col = 1 then   // Nouveau composant
      CellNewComposant := GRILLE.Cells[1, GRILLE.Row]
    else
      CellOldComposant := GRILLE.Cells[0, GRILLE.Row];
    if TobWPA.Detail[GRILLE.Row - 1].GetString('TYPECHGLIEN') = 'COEF' then
    begin
      tRadioButton(GetControl('RBCOEFQLIEN')).Checked := True;
      RBCOEFQLIEN_onClick(Self);
    end
    else if TobWPA.Detail[GRILLE.Row - 1].GetValue('TYPECHGLIEN') = 'QTE' then
    begin
      tRadioButton(GetControl('RBNOUVQLIEN')).Checked := True;
      RBNOUVQLIEN_OnClick(Self);
    end
    else
    begin
      tRadioButton(GetControl('RBRIENLIEN')).Checked := True;
      RBRIENLIEN_OnClick(Self);
    end
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 17/10/2002
Modifié le ... :   /  /
Description .. : Rempli la grille avec les données de la tob
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.RafraichirGrille;
begin
	GRILLE.VidePile(False);
	TobWPA.PutGridDetail(GRILLE, False, False, 'OLDCOMPOSANT;NEWCOMPOSANT', True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 03/01/2003
Modifié le ... :   /  /
Description .. : Gestion du F5
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
procedure TOM_WPARAM.GRILLE_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
	case Key of
   	VK_F5 : if GRILLE.ElipsisButton then Cell_OnElipsisClick(Sender);
  end;
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : Force la saise en majuscules
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.GRILLE_OnKeyPress(Sender: TObject; var Key: Char);
begin
  Key := UpCase(Key);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : provoque l'ellipsis clique au dbl clique de la cellule
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
procedure TOM_WPARAM.GRILLE_OnDblClick(Sender: TObject);
begin
	if GRILLE.ElipsisButton then Cell_OnElipsisClick(Sender);
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : contrôle lorsque la grille perd le focus
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.GRILLE_OnExit(Sender: TObject);
var
  ACol, ARow: Integer;
  bCancel: Boolean;
begin
  bCancel := False;
  ACol := Grille.Col;
  ARow := Grille.Row;
  FromApplyChangement := True;
  if ARow > 0 then
  	Grille_OnCellExit(Sender, ACol, ARow, bCancel);
  FromApplyChangement := False;
end;

procedure TOM_WPARAM.GRILLE_OnEnter(Sender: TObject);
var
  ACol, ARow: Integer;
  bCancel: Boolean;
begin
  bCancel := False;
  ACol := Grille.Col;
  ARow := Grille.Row;
  FromApplyChangement := True;
  if ARow > 0 then
  	Grille_OnCellEnter(Sender, ACol, ARow, bCancel);
  FromApplyChangement := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 03/12/2002
Modifié le ... : 03/12/2002
Description .. : Vérifie la validité des données modifiées :
Suite ........ : Si OK, on enregistre dans la TOB
Suite ........ : Sinon, Erreur et on reste à la même position dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.GRILLE_OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  iSaveCol, iSaveRow: Integer;
  T: Tob;
  {***********A.G.L.***********************************************
  Auteur  ...... : Thibaut SUBLET
  Créé le ...... : 03/12/2002
  Modifié le ... :   /  /
  Description .. : Test la validité du composant
  Mots clefs ... :
  *****************************************************************}
  function ComposantEstValide(ValeurComposant: String; NumLig, NumCol: Integer): Boolean;
  var
    OffSet, i: Integer;
  begin
    OffSet := iifV(NumCol=0, 1, -1);
    Result := (ValeurComposant <> '') and (wExistGA(wGetArticleFromCodeArticle(ValeurComposant)));
    if not Result then
      LastError := ErrEmptyComposant
    else
    begin
      if FromApplyChangement and (ValeurComposant = GRILLE.Cells[NumCol + OffSet, NumLig]) and tRadioButton(GetControl('RBRIENLIEN')).Checked then
      begin
        LastError := ErrSameComposant;
        Result := False;
      end
    	else if (wExistGA(wGetArticleFromCodeArticle(ValeurComposant),False, 'GA_MODECONSO=""')) then
      begin
      	LastError := ErrModeConso;
        result    := false;
      end
      else if NumCol = 0 then
      begin
        Result := True;
        i := 0;
        while Result and (i <= TobWPA.Detail.Count - 1) do
        begin
          if NumLig - 1 <> i then
            Result := ValeurComposant <> TobWPA.Detail[i].GetString('OLDCOMPOSANT');
          inc(i);
        end;
        if not Result then
          LastError := ErrEvenChangedComposant;
      end;
      if Result then
        LastError := 0;
    end;
  end;

begin
  if ARow <= 0 then Exit;

  TobWPA.Detail[ARow-1].AddChampSupValeur(iifV(ACol = 0, 'OLDCOMPOSANT', 'NEWCOMPOSANT'), GRILLE.Cells[ACol, ARow]);
  if ComposantEstValide(GRILLE.Cells[ACol, ARow], ARow, ACol) then
  begin
    iSaveCol := GRILLE.Col;
    iSaveRow := GRILLE.Row;
    RafraichirGrille;
    GRILLE.Col := iSaveCol;
    GRILLE.Row := iSaveRow;
    if (TobWPA.Detail[ARow-1].GetString('OLDCOMPOSANT') <> '') and (TobWPA.Detail[ARow-1].GetString('NEWCOMPOSANT') <> '') then
      SetControlEnabled('BTANNULERFILS', False);
    GRILLE.EditorMode := False;
    GRILLE.ElipsisButton := False;
    T:= TobWPA.Detail[ARow-1];
    wInitUnitesLien(T);
  end
  else
  begin
    TobWPA.Detail[ARow-1].AddChampSupValeur(iifV(ACol = 0, 'OLDCOMPOSANT', 'NEWCOMPOSANT'), '');
    if (ARow <> 0) then
    begin
      GRILLE.Col := ACol;
      GRILLE.Row := ARow;
      GRILLE.ElipsisButton := True;
      GRILLE.EditorMode := True;
      if tPageControl(GetControl('PAGES')).ActivePage <> tTabSheet(GetControl('PGENERAL')) then
        tPageControl(GetControl('PAGES')).ActivePage := tTabSheet(GetControl('PGENERAL'));
      if ((LastError <> 0) and (LastError <> ErrSameComposant)) or ((LastError = ErrSameComposant) and ((Sender is TToolBarButton97) or FromBtMoulinette)) then
        PGIERROR(TexteMessage[LastError]);
    end;
  end;
  LastErrorDejaTraite := LastError <> 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : Test la validité de saisie d'un couple OLD/NEW
Suite ........ : COMPOSANT et des inits de qte lien.
Mots clefs ... :
*****************************************************************}
function TOM_WPARAM.wEstUnFilsValide(var iCol, iRow: Integer): tWinControl;
var
	i: Integer;
begin
  Result := nil;
  i := 0;
  while (Result = nil) and (i < TobWPA.Detail.Count) do
  begin
    if TobWPA.Detail[i].GetString('OLDCOMPOSANT') = '' then
    begin
      iCol := 0;
      if LastError = 0 then LastError := ErrEmptyComposant;
      Result := GRILLE;
    end
    else if TobWPA.Detail[i].GetString('NEWCOMPOSANT') = '' then
    begin
      iCol := 1;
      if LastError = 0 then LastError := ErrEmptyComposant;
      Result := GRILLE;
    end
    else if (wExistGA(wGetArticleFromCodeArticle(TobWPA.Detail[i].GetString('NEWCOMPOSANT')),False, 'GA_MODECONSO=""')) then
    begin
      iCol := 1;
      if LastError = 0 then LastError := ErrModeConso;
      Result := GRILLE;
    end
    else if (TobWPA.Detail[i].GetString('OLDCOMPOSANT') = TobWPA.Detail[i].GetString('NEWCOMPOSANT')) and (TobWPA.Detail[i].GetString('TYPECHGLIEN') = 'RIEN') then
    begin
      if LastError = 0 then LastError := ErrSameComposant;
      Result := tWinControl(GetControl('RBRIENLIEN'));
    end
    else if TobWPA.Detail[i].GetString('TYPECHGLIEN') = 'QTE' then
    begin
      if TobWPA.Detail[i].GetDouble('QLIEN') = 0 then
      begin
      iCol := 0;
      if LastError = 0 then LastError := ErrQLien;
      Result := tWinControl(GetControl('QLIEN'));
      end;
    end
    else if TobWPA.Detail[i].GetString('TYPECHGLIEN') = 'COEF' then
    begin
      if TobWPA.Detail[i].GetDouble('COEFQLIENMULT') = 0 then
      begin
        iCol := 0;
        if LastError = 0 then LastError := ErrCoefLien;
        Result := tWinControl(GetControl('COEFQLIENMULT'));
      end
      else if TobWPA.Detail[i].GetDouble('COEFQLIENDIV') = 0 then
      begin
        iCol := 0;
        if LastError = 0 then LastError := ErrCoefLien;
        Result := tWinControl(GetControl('COEFQLIENDIV'));
      end;
    end
    ;
    if Assigned(result) then
      iRow := i + 1;
    inc(i);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : initialise les données sur les liens denomenclature
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.wInitLienFields(var T: Tob);
begin
  TobWPA_Fils.AddChampSupValeur('QLIEN', 1);
  TobWPA_Fils.AddChampSupValeur('UNITELIEN', 'UNI');
  TobWPA_Fils.AddChampSupValeur('COEFQLIENMULT', 1);
  TobWPA_Fils.AddChampSupValeur('COEFQLIENDIV', 1);
  TobWPA_Fils.AddChampSupValeur('PERTEPROP', 1);
  TobWPA_Fils.AddChampSupValeur('UNITEPERTEPROP', 'UNI');
  TobWPA_Fils.AddChampSupValeur('COEFPERTEPROPMULT', 1);
  TobWPA_Fils.AddChampSupValeur('COEFPERTEPROPDIV', 1);
  TobWPA_Fils.AddChampSupValeur('QPPER', 1);
  TobWPA_Fils.AddChampSupValeur('UNITEPPER', 'UNI');
  TobWPA_Fils.AddChampSupValeur('COEFQPPERMULT', 1);
  TobWPA_Fils.AddChampSupValeur('COEFQPPERDIV', 1);
  TobWPA_Fils.AddChampSupValeur('QPERIODE', 1);
  TobWPA_Fils.AddChampSupValeur('UNITEPERIODE', 'UNI');
  TobWPA_Fils.AddChampSupValeur('COEFQPERIODEMULT', 1);
  TobWPA_Fils.AddChampSupValeur('COEFQPERIODEDIV', 1);
  TobWPA_Fils.AddChampSupValeur('QPFIXE', 1);
  TobWPA_Fils.AddChampSupValeur('UNITEPFIXE', 'UNI');
  TobWPA_Fils.AddChampSupValeur('COEFQPFIXEMULT', 1);
  TobWPA_Fils.AddChampSupValeur('COEFQPFIXEDIV', 1);
  TobWPA_Fils.AddChampSupValeur('UNITESTOCK', 'UNI');
  TobWPA_Fils.AddChampSupValeur('CBQLIEN'     , wTrue);
  TobWPA_Fils.AddChampSupValeur('CBPERTEPROP' , wFalse);
  TobWPA_Fils.AddChampSupValeur('CBQPPER'     , wFalse);
  TobWPA_Fils.AddChampSupValeur('CBQPERIODE'  , wFalse);
  TobWPA_Fils.AddChampSupValeur('CBQPFIXE'    , wFalse);
end;

procedure TOM_WPARAM.wInitUnitesLien(var TobWPA_Fils: Tob);
Var
  Unites: MyArrayValue;
  ACol, ARow : integer;
  bCancel: boolean;
begin
  Unites := nil;
  if (TobWPA_Fils.GetValue('NEWCOMPOSANT')<> CellNewComposant) or (TobWPA_Fils.GetValue('OLDCOMPOSANT')<> CellOldComposant)then
  begin
    if TobWPA_Fils.GetValue('NEWCOMPOSANT')<> TobWPA_Fils.GetValue('OLDCOMPOSANT') then
    begin
      Unites := wGetFieldsFromGA(['GA_UNITECONSO', 'GA_QUALIFUNITESTO'], wGetArticleFromCodeArticle(TobWPA_Fils.GetValue('NEWCOMPOSANT')) );
      if Assigned(Unites) then
      begin
        Acol := 1;
        ARow := GRILLE.Row;
        bCancel:=false;
        GRILLE_OnCellEnter(Self, ACol, ARow, bCancel);
        SetControlProperty('UNITELIEN'      , 'VALUE', Unites[0]);
        SetControlProperty('UNITEPERTEPROP' , 'VALUE', Unites[0]);
        SetControlProperty('UNITEPERIODE'   , 'VALUE', Unites[0]);
        SetControlProperty('UNITEPPER'      , 'VALUE', Unites[0]);
        SetControlProperty('UNITEPFIXE'     , 'VALUE', Unites[0]);
        SetControlProperty('UNITESTOCK'     , 'VALUE', Unites[1]);
        SetControlProperty('UNITELIEN'      , 'VISIBLE', True);
        SetControlProperty('UNITEPERTEPROP' , 'VISIBLE', True);
        SetControlProperty('UNITEPERIODE'   , 'VISIBLE', True);
        SetControlProperty('UNITEPPER'      , 'VISIBLE', True);
        SetControlProperty('UNITEPFIXE'     , 'VISIBLE', True);
        SetControlProperty('UNITELIENNOME'  , 'VISIBLE', false);
        SetControlProperty('UNITESTOCK'     , 'VALUE', Unites[1]);
      end;
    end
    else
    begin
      SetControlProperty('UNITELIEN'      , 'VISIBLE', false);
      SetControlProperty('UNITEPERTEPROP' , 'VISIBLE', false);
      SetControlProperty('UNITEPERIODE'   , 'VISIBLE', false);
      SetControlProperty('UNITEPPER'      , 'VISIBLE', false);
      SetControlProperty('UNITEPFIXE'     , 'VISIBLE', false);
      SetControlProperty('UNITELIENNOME'  , 'VISIBLE', true);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : Ajoute une ligne dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.btNouveauFils_OnClick(Sender: tObject);
var
  iCol, iRow: Integer;
  bCancel: Boolean;
  C: tWinControl;
begin
  iCol := GRILLE.Col;
  iRow := GRILLE.Row;
  bCancel := False;
  GRILLE_OnCellExit(Sender, iCol, iRow, bCancel);
  if LastError = 0 then
  begin
    C := wEstUnFilsValide(iCol, iRow);
    if (LastError = 0) and (C = nil) then
    begin
      { Tob }
      TobWPA_Fils := Tob.Create('CHGCOMPOSANT', TobWPA, -1);
      TobWPA_Fils.AddChampSupValeur('OLDCOMPOSANT', '');
      TobWPA_Fils.AddChampSupValeur('NEWCOMPOSANT', '');
      TobWPA_Fils.AddChampSupValeur('TYPECHGLIEN', 'RIEN');
      GRILLE.RowCount := GRILLE.RowCount + 1;
      wInitLienFields(TobWPA_Fils);

      iCol := 0;
      iRow := GRILLE.RowCount - 1;
      GRILLE.Col := iCol;
      GRILLE.Row := iRow;
      FromPutEcran := True;
      TobWPA_Fils.PutEcran(Ecran, tWinControl(GetControl('GBQLIEN')));
      FromPutEcran := False;
      SetControlEnabled('BTANNULERFILS', True);
    end
    else
    begin
      GRILLE.Col := iCol;
      GRILLE.Row := iRow;
    end;

    if LastError <> 0 then PGIERROR(TexteMessage[LastError]);
    GRILLE_OnCellEnter(Self, iCol, iRow, bCancel);
    if Assigned(C) and (not (C is thGrid)) then C.SetFocus;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : Valide la saisie d'un composant (simulation ergonomique)
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.BTVALIDCOMPOSANTS_OnClick(Sender: tObject);
var
  bCancel: Boolean;
  ACol, ARow: Integer;
begin
  FromApplyChangement := True;
  bCancel := False;
  ACol := Grille.Col;
  ARow := Grille.Row;
  if ARow > 0 then
    Grille_OnCellExit(Sender, ACol, ARow, bCancel);
  FromApplyChangement := False;
  if LastError = 0 then
  begin
    Grille.SetFocus;
    Grille.Col := 0;
    Grille.Row := 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : Supprime une ligne dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.btSupprimerFils_OnClick(Sender: tObject);
var
iCol, iRow: Integer;
bCancel: Boolean;
begin
  bCancel := False;
  iCol := GRILLE.Col;
  iRow := GRILLE.Row;
  GRILLE_OnCellExit(Self, iCol, iRow, bCancel);
  if (LastError = 0) and (GRILLE.RowCount > 2) then
  begin
    if PGIASK('Voulez-vous vraiment supprimer cette ligne ?', 'Confirmation de suppression') = mrYes then
    begin
       GRILLE.DeleteRow(GRILLE.Row);
       iCol := 0;
       iRow := GRILLE.Row;
       TobWPA.GetGridDetail(GRILLE, GRILLE.RowCount - 1, '', '');
       RafraichirGrille;
       GRILLE.Col := iCol;
       GRILLE.Row := iRow;
       GRILLE_OnCellEnter(Self, iCol, iRow, bCancel);
    end;
  end
  else
    GRILLE_OnCellEnter(Self, iCol, iRow, bCancel);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : Annuler la saisie d'un nouveau changement de composant
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.btAnnulerFils_OnClick(Sender: tObject);
var
  iCol, iRow: Integer;
  bCancel: Boolean;
begin
  if (GRILLE.Rowcount > 2) and (GRILLE.Row = GRILLE.RowCount - 1) {and (CODECOMPOSANT.Visible)} and ((TobWPA.Detail[GRILLE.Row - 1].GetString('OLDCOMPOSANT') = '') or (TobWPA.Detail[GRILLE.Row - 1].GetString('NEWCOMPOSANT') = '')) then
  begin
    TobWPA.Detail[GRILLE.Row - 1].Free;
    RafraichirGrille;
    bCancel := False;
    iCol := 0;
    iRow := GRILLE.RowCount - 1;
    GRILLE.Col := iCol;
    GRILLE.Row := iRow;
    GRILLE_OnCellEnter(Self, iCol, iRow, bCancel);
  end;
  SetControlEnabled('BTANNULERFILS', False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : Efface le contenu d'une ligne
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.btEraseFils_OnClick(Sender: tObject);
var
  iCol, iRow: Integer;
  bCancel: Boolean;
begin
  TobWPA.Detail[GRILLE.Row - 1].SetString('OLDCOMPOSANT', '');
  TobWPA.Detail[GRILLE.Row - 1].SetString('NEWCOMPOSANT', '');
  TobWPA.Detail[GRILLE.Row - 1].PutLigneGrid(GRILLE, GRILLE.Row, False, False, 'OLDCOMPOSANT;NEWCOMPOSANT');
  bCancel := False;
  iCol := 0;
  iRow := GRILLE.Row;
  GRILLE.Col := iCol;
  GRILLE_OnCellEnter(self, iCol, iRow, bCancel);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : MàJ de l'affichage des qte et coef.
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.wSetAffichageQLien(ctx_QTE_ou_COEF: String);
var
	bQteCtx: Boolean;
begin
  if (ctx_QTE_ou_COEF = 'QTE') or (ctx_QTE_ou_COEF = 'COEF') or (ctx_QTE_ou_COEF = 'RIEN') then
  begin
    bQteCtx := iifV(ctx_QTE_ou_COEF = 'QTE', True, False);

    if ctx_QTE_ou_COEF = 'COEF' then
    begin
      SetControlProperty('GBQLIEN'    , 'CAPTION', 'Coefficients nouveau composant : Quantité de lien * ');
      SetControlProperty('LBAPPLIQUER', 'VISIBLE', False);
    end
    else
    begin
      SetControlProperty('GBQLIEN'    , 'CAPTION', 'Quantités nouveau composant');
      SetControlProperty('LBAPPLIQUER', 'VISIBLE', True);
    end;

    SetControlVisible('GBQLIEN'     , iifV(ctx_QTE_ou_COEF = 'RIEN', False, True));
    SetControlVisible('PANTITLE'    , iifV(ctx_QTE_ou_COEF = 'RIEN', False, True));
    SetControlVisible('PNCOEFF'     , iifV(ctx_QTE_ou_COEF = 'RIEN', False, not bQteCtx));
    SetControlVisible('PNQUANTITES' , iifV(ctx_QTE_ou_COEF = 'RIEN', False, bQteCtx));
    TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur('TYPECHGLIEN', ctx_QTE_ou_COEF);
    TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur('CBQLIEN'     , iif(TCheckBox(GetControl('CBQLIEN')).Checked, wTrue, wFalse));
    TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur('CBPERTEPROP' , iif(TCheckBox(GetControl('CBPERTEPROP')).Checked, wTrue, wFalse));
    TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur('CBQPPER'     , iif(TCheckBox(GetControl('CBQPPER')).Checked, wTrue, wFalse));
    TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur('CBQPERIODE'  , iif(TCheckBox(GetControl('CBQPERIODE')).Checked, wTrue, wFalse));
    TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur('CBQPFIXE'    , iif(TCheckBox(GetControl('CBQPFIXE')).Checked, wTrue, wFalse));
    {$IFDEF STK}
      // Si Action = nouvelle quantité et nouveau comoposant géré en série -> interdire saisie de décimales
      if (bQteCtx) then
      begin
        SetControlProperty('QLIEN'   , 'Decimals', iif(GereSerie(wGetArticleFromCodeArticle(GRILLE.Cells[1, GRILLE.Row])), 0, NbDeci));
        SetControlProperty('QPPER'   , 'Decimals', iif(GereSerie(wGetArticleFromCodeArticle(GRILLE.Cells[1, GRILLE.Row])), 0, NbDeci));
        SetControlProperty('QPERIODE', 'Decimals', iif(GereSerie(wGetArticleFromCodeArticle(GRILLE.Cells[1, GRILLE.Row])), 0, NbDeci));
        SetControlProperty('QPFIXE'  , 'Decimals', iif(GereSerie(wGetArticleFromCodeArticle(GRILLE.Cells[1, GRILLE.Row])), 0, NbDeci));
      end;
    {$ENDIF STK}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : Style de modification de lien : Application d'1 coef
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.RBCOEFQLIEN_OnClick(Sender: tObject);
begin
	wSetAffichageQLien('COEF');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : Style de modification de lien : Changement de quantité
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.RBRIENLIEN_OnClick(Sender: tObject);
begin
	wSetAffichageQLien('RIEN');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : Style de modification de lien : Changement de quantité
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.RBNOUVQLIEN_OnClick(Sender: tObject);
begin
	wSetAffichageQLien('QTE');
end;

{***********A.G.0.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : Evénement unique à tous les composant concernant la
Suite ........ : quantité de lien : Recharge la Tob à chaque changement
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.Fields_OnChange(Sender: tObject);
begin
	if not FromPutEcran then
   	TobWPA.Detail[GRILLE.Row - 1].GetEcran(Ecran, tWinControl(GetControl('GBQLIEN')));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : Gestion du champs permettant d'identifier le comportement
Suite ........ : sur les révisions.
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.RBREVISION_OnClick(Sender: tObject);
begin
  with (Sender as tRadioButton) do
  begin
    if Name = 'RBRIEN' then
    begin
      SetString(REVISION, 'RIEN');
      TobWPA.SetString(REVISION, 'RIEN');
    end
    else if Name = 'RBMINVAL' then
    begin
      SetString(REVISION, 'MINVAL');
      TobWPA.SetString(REVISION, 'MINVAL');
    end
    else if Name = 'RBMAJVALPEROLD' then
    begin
      SetString(REVISION, 'MAJVALPEROLD');
      TobWPA.SetString(REVISION, 'MAJVALPEROLD');
    end
    ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 06/12/2002
Modifié le ... :   /  /
Description .. : Lance le changement de composants
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
procedure TOM_WPARAM.btMoulinette_OnClick(Sender: tObject);
begin
  tToolBarButton97(Sender).Enabled := false;
  try
    if PGIASK('Confirmez-vous l''application de ces changements ?', Ecran.Caption) = mrYes then
    begin
      FromApplyChangement := True;
      FromBtMoulinette := True;

      if DS.State = dsBrowse then
        { Validation de la tob qui n'est pas dans WPA }
        RecordIsValid
      else
        { Sauvegarde des données + Validation de la tob qui n'est pas dans WPA }
        tToolBarButton97(GetControl('BVALIDER')).Click;

      FromBtMoulinette := False;
      FromApplyChangement := False;

      if (LastError = 0) and not LastErrorDejaTraite then
        wRemplaceComposants(TobWPA);
    end;
  finally
    tToolBarButton97(Sender).Enabled := true;
  end;
end;
{$ENDIF GPAO }
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{-----------------------------------      FIN DE LA GESTION DE LA GRILLE      -----------------------------------------}

{---------------------------------FIN DES METHODES POUR LE REMPLACEMENT DE COMPOSANTS----------------------------------}

{----------------------------------------------------------------------------------
Définition des DATATYPE et PLUS sur les Circuits et Phases en fonction de l'article
----------------------------------------------------------------------------------}
function TOM_WPARAM.ParticularitesSousTraitance : boolean;
var
  sArticle, sNatureTravail, sItineraire : string;
  aArticleNature : MyArrayValue;
  lSousTraitance  : boolean;
begin
  aArticleNature := nil;

  lSousTraitance := (sFonctionnalite=sTarifSousTraitantAchat) or (sFonctionnalite=sTarifSousTraitantPhase);
  if Assigned(GBSousTraitance) then GBSousTraitance.Visible     := lSousTraitance;

  // Si tarif Sous traitant
  if lSousTraitance and Assigned(GetControl('WPA_VARCHAR07')) and Assigned(GetControl('WPA_VARCHAR08')) then
  begin
    SetControlProperty('WPA_VARCHAR07', 'Plus'     , ''       ); // Pas de filtre sur les circuit
    SetControlProperty('WPA_VARCHAR08', 'DataType' ,'QUTPHASE'); // Liste sur la table des phases

    sArticle        := wGetArticleFromCodeArticle(GetString('WPA_VARCHAR03'));
    sNatureTravail  := 'FAB';

    // Si Tarif sur un article
    if (sArticle<>'') then
    begin
      aArticleNature := WGetSqlFieldsValues(['WAN_CODITI', 'WAN_CIRCUIT'], 'WARTNAT', 'WAN_ARTICLE = "'+sArticle+'" AND WAN_NATURETRAVAIL = "'+sNatureTravail+'"');
      // Article avec une nature de travail
      if Assigned(aArticleNature) then
      begin
        sItineraire := aArticleNature[0];
        //Article avec un itinéraire
        if (sItineraire<>'') then
        begin
          SetControlProperty('WPA_VARCHAR07', 'Plus'    , 'QCI_CODITI="'+sItineraire+'"'); // Filtre sur les circuit de l'itinéraire
          SetControlProperty('WPA_VARCHAR08', 'DataType', 'WPHASEITINERAIRE'            ); // Liste sur les phases de l'itinéraire
          SetControlProperty('WPA_VARCHAR08', 'Plus'    , 'QP_CODITI="' +sItineraire+'"'); // filtre sur les phases de l'itinéraire
        end;
      end;
    end;
  end;
  Result := True;
end;

{$IFDEF GPAO}
function TOM_WPARAM.GetCleWVS: tCleWVS;
begin
  Result.Article      := Article;
  Result.CodeVersion  := GetString('WPA_VARCHAR02');
end;
{$ENDIF GPAO }

{$IFNDEF EAGLSERVER}
procedure TOM_WPARAM.MnLpLegende_OnClick(Sender: TObject);
begin
  if EcranIs('YTARIFSRECH_FIC') then
    AglLanceFiche('Y', 'YTARIFSRECHERCHEL', '', '', '');
end;
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure TOM_WPARAM.Ecran_OnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  t : tTime; { Voir JL§ }
begin
  Case Key of
    VK_F10: begin
              t := time;
              repeat until (time-t) > 0.000001;
              TFFiche(Ecran).BValider.Click;
            end;
  end;
end;
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/09/2004
Modifié le ... :   /  /    
Description .. : Visu à pla des nomnclature en fonction de la nature de 
Suite ........ : travail
Mots clefs ... : 
*****************************************************************}
procedure TOM_WPARAM.MnWNA_OnClick(Sender: tObject);
begin
  CallFicWNDActiveForVisu(wGetNatureFromMenuItem(TMenuItem(Sender).Name), wGetArticleFromCodeArticle(GetControlText(chgARTICLE)), '');
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/09/2004
Modifié le ... :   /  /
Description .. : Visualisation de la nomenclature à plat
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.PmWND_OnPopup(Sender: tObject);
var
  Pm            : TPopupMenu;
  TabMN         : Array of TMenuItem;
  sNatureTravail: String;
  i             : Integer;
begin
  Pm := TPopupMenu(Sender);
  while Pm.Items.Count > 0 do
    Pm.Items[Pred(Pm.Items.Count)].Free;

  if Assigned(GetControl(chgNATURETRAVAIL)) then
  begin
    sNatureTravail := GetControlText(chgNATURETRAVAIL);
    if sNatureTravail = '' then
      for i := 0 to Pred(THMultiValComboBox(GetControl(chgNATURETRAVAIL)).Items.Count) do
        sNatureTravail := sNatureTravail + iif(sNatureTravail <> '', ';', '') + THMultiValComboBox(GetControl(chgNATURETRAVAIL)).Values[i];

    while sNatureTravail <> '' do
    begin
      SetLength(TabMN, Length(TabMN) + 1);
      TabMN[Pred(Length(TabMN))] := TMenuItem.Create(Pm);
      with TabMN[Pred(Length(TabMN))] do
      begin
        Name := 'MnWND_' + ReadTokenSt(sNatureTravail);
        Caption := wGetLibelleWNA(wGetNatureFromMenuItem(Name));
        OnClick := MnWNA_OnClick
      end
    end;
    Pm.Items.Add(TabMn);
    SetLength(TabMN, 0);
    ActivateXPPopUp(Pm)    
  end
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
procedure TOM_WPARAM.MnLpVisuNome_OnClick(Sender: TObject);
begin
  CallFicWNDActiveForVisu(wGetFirstWAN(wGetArticleFromCodeArticle(GRILLE.Cells[GRILLE.Col, GRILLE.Row])), wGetArticleFromCodeArticle(GRILLE.Cells[GRILLE.Col, GRILLE.Row]), '');
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

procedure TOM_WPARAM.CB_ANALYSE_OnClick(Sender: tObject);
var
  lConsolide : boolean;
begin
  lConsolide := tCheckBox(GetControl('CB_ANALYSE')).Checked;
  SetControlProperty('GBANALYSECONSOLIDE', 'Visible', lConsolide);
  SetControlProperty('TWPA_VARCHAR01'    , 'Visible', lConsolide);
  SetControlProperty('TWPA_TOTAL'        , 'Visible', lConsolide);
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.Tiers_OnElipsisClick(Sender: TObject);
Var
  sRange: string;
  sWhereNatureAuxi: string;
begin
  if not(DS.State in [dsInsert,dsEdit]) then
     DS.edit;
  if (sFonctionnalite=sTarifClient) then
    sWhereNatureAuxi := 'T_NATUREAUXI="'+sNatureAuxiClient+'" OR T_NATUREAUXI="'+sNatureAuxiProspect+'"'
  else if (NatureAuxiliaire(sFonctionnalite)='TIE')  then
    sWhereNatureAuxi := 'T_NATUREAUXI="'+sNatureAuxiClient+'" OR T_NATUREAUXI="'+sNatureAuxiProspect+'"'+'" OR T_NATUREAUXI="'+sNatureAuxiFournisseur+'"'
  else
    sWhereNatureAuxi := 'T_NATUREAUXI="'+NatureAuxiliaire(sFonctionnalite)+'"'
  ;

  { Tiers livré }
  if tControl(Sender).Name='WPA_VARCHAR06' then
    GetTiersRecherche(thEdit(Sender), sWhereNatureAuxi, '', '')
  else
  begin
  { Tiers transporteur }
    sRange:= iif(sFonctionnalite=sFraisAnnexes, 'YTC_TYPEFOURNI='+sTypeFourni, '');
    GetTiersRecherche(thEdit(Sender), sWhereNatureAuxi, sRange, '');
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.TiersLivre_OnElipsisClick(Sender: TObject);
Var
  sWhere: string;
begin
  if not(DS.State in [dsInsert,dsEdit]) then
    DS.edit;
  sWhere:= 'T_NATUREAUXI="'+sNatureAuxiClient+'" OR T_NATUREAUXI="'+sNatureAuxiProspect+'"';
  GetTiersRecherche(thEdit(Sender), sWhere, '', '')
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

procedure TOM_WPARAM.CBAPPLIQUER_OnClick(Sender: tObject);
var
  NomControl: string;

  function GetFieldName: string;
  begin
    Result := StringReplace (NomControl, 'CB', '', [rfReplaceAll]);
  end;

  function GetUniteName: string;
  begin
    if NomControl = 'CBPERTEPROP' then
      Result := ''
    else
      Result := StringReplace (NomControl, 'CBQ' , 'UNITE', [rfReplaceAll])
  end;

begin
  NomControl:= tCheckBox(Sender).Name;
  SetControlProperty(GetFieldName, 'ENABLED' , TCheckBox(GetControl(NomControl)).Checked);
  SetControlProperty(GetUniteName, 'ENABLED' , TCheckBox(GetControl(NomControl)).Checked);
  if TCheckBox(GetControl(NomControl)).Checked then
  begin
    SetControlProperty(GetFieldName, 'COLOR' , ClWindow);
    SetControlProperty(GetUniteName, 'COLOR' , ClWindow);
    if Assigned(TobWPA) then
      TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur(NomControl, wTrue);
  end
  else // CbUnChecked
  begin
    SetControlProperty(GetFieldName, 'COLOR' , ClBtnFace);
    SetControlProperty(GetUniteName, 'COLOR' , ClBtnFace);
    if Assigned(TobWPA) then
      TobWPA.Detail[GRILLE.Row - 1].AddChampSupValeur(NomControl, wFalse);
  end;
end;

procedure TOM_WPARAM.SetAdresseTiers(InfosTiers: MyArrayValue);
begin
  SetString         ('WPA_COMBO04'   , InfosTiers[0]); // Secteur géo
  SetString         ('WPA_VARCHAR10' , InfosTiers[1]); // Code postal
  SetControlText    ('LBCODEPOSTAL'  , RechDom('ICODEPOSTAL',GetControlText('WPA_VARCHAR10'),False));
  SetString         ('WPA_COMBO05'   , InfosTiers[2]); // Pays
  SetString         ('WPA_VARCHAR11' , InfosTiers[3]); // Région
  SetString         ('WPA_COMBO07'   , InfosTiers[4]); // Mode d'expédition
  SetControlEnabled ('WPA_COMBO04'   , false);
  SetControlEnabled ('WPA_VARCHAR10' , false);
  SetControlEnabled ('WPA_COMBO05'   , false);
  SetControlEnabled ('WPA_VARCHAR11' , false);
  SetControlEnabled ('WPA_COMBO07'   , false);
  SetControlProperty('WPA_VARCHAR13' , 'Plus', 'ADR_TYPEADRESSE="TIE" AND ADR_REFCODE="' + GetControlText('WPA_VARCHAR06') + '"'
                                     + ' AND ADR_LIVR="X"' );
end;

{$IFDEF QUALITE}
procedure TOM_WPARAM.Ressource_OnElipsisClick(Sender: TOBject);
Var
	Ressource: string;
begin
  if not(DS.State in [dsInsert,dsEdit]) then
    DS.edit;
  Ressource:= AGLLanceFiche('AFF','RESSOURCERECH_MUL','ARS_ESTHUMAIN=X;ARS_GENERIQUE=-','','ARS_RESSOURCE=' + GetControlText(TControl(Sender).Name)+';PRODUITPGI=QNC');
	if Ressource <> '' then
 	begin
 		SetControlText(TControl(Sender).Name, Ressource);
  	SetString(TControl(Sender).Name,Ressource);
  end;
end;
{$ENDIF QUALITE}

procedure TOM_WPARAM.WPA_BOOLEAN13_OnClick(Sender: tObject);
begin
  SetControlEnabled('WPA_COMBO04', GetBoolean('WPA_BOOLEAN12') or TDbCheckBox(GetControl('WPA_BOOLEAN13')).Checked);
  if (not GetControlEnabled('WPA_COMBO04')) then
    SetString('WPA_COMBO04', '');
end;

procedure TOM_WPARAM.WPA_BOOLEAN12_OnClick(Sender: tObject);
begin
  SetControlEnabled('WPA_COMBO04', TDbCheckBox(GetControl('WPA_BOOLEAN12')).Checked or GetBoolean('WPA_BOOLEAN13'));
  if (not GetControlEnabled('WPA_COMBO04')) then
    SetString('WPA_COMBO04', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Permet de ramener le GA_CODEARTICLE et
Suite ........ : le GA_ARTICLE même si différent (article dimensionné)
Mots clefs ... :
*****************************************************************}
procedure TOM_WPARAM.MyOnGetArticle(const sArticle: String);
begin
  SetString('WPA_VARCHAR16', sArticle);
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zones Dimensions de l'article
---------------------------------------------------------------------------------}
procedure TOM_WPARAM.ControleAffichageDimensionsArticle;
var
  lArticleDimensionne: boolean;
begin
  {Le tarif est associé a un article dimensionné}
  lArticleDimensionne :=     (sParametre[iPosArticle]=wTrue)
                         and (GetString('WPA_VARCHAR03')<>'')
                         and (GetString('WPA_VARCHAR03')=wGetCodeArticleFromArticle(GetString('WPA_VARCHAR16')))
                         and (wGetArticleFromCodeArticle(GetString('WPA_VARCHAR03'))<>GetString('WPA_VARCHAR16'));
  SetcontrolProperty('GBDIMENSIONS', 'Visible', lArticleDimensionne);

{$IFDEF STK}
  { Gestion de l'affichage du cadre Dimensions pour les articles dimensionnés }
  SetCadreDimensions(GetString('WPA_VARCHAR16'), False, GetControl('PNDIMENSIONS'),'');
{$ENDIF STK}

end;

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrSetColor;
begin
  SetControlProperty('VFOND', 'Color', GetInteger('WPA_INTEGER01'));
  if GetString('WPA_VARCHAR04') <> '' then
    SetControlProperty('VITEM', 'Color', ValeurI(GetString('WPA_VARCHAR04')));
  SetControlProperty('VSELECTION', 'Color', GetInteger('WPA_INTEGER02'));
  SetControlProperty('VSAMEDI', 'Color', GetInteger('WPA_INTEGER03'));
  SetControlProperty('VDIMANCHE', 'Color', GetInteger('WPA_INTEGER04'));
  SetControlProperty('VJOURSFERIES', 'Color', GetInteger('WPA_INTEGER05'));
  if GetString('WPA_VARCHAR01') <> '' then
    SetControlProperty('VTRAIT', 'Color', ValeurI(GetString('WPA_VARCHAR01')));
  if GetString('WPA_VARCHAR02') <> '' then
    SetControlProperty('VALERTES', 'Color', ValeurI(GetString('WPA_VARCHAR02')));
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.SetPlanLivrOnLoad;
var
  TSCellule, TSHint, TSDetail: TStrings;
  i: Integer;
  PlanLivrJour: Boolean;
begin
  { }
  PlanLivrJour := RightStr(GetString('WPA_CODEPARAM'), 1) = 'J';
  SetControlVisible('WPA_BOOLEAN01', PlanLivrJour);
  SetControlVisible('LSAMEDI', PlanLivrJour); SetControlVisible('BSAMEDI', PlanLivrJour); SetControlVisible('VSAMEDI', PlanLivrJour);
  SetControlVisible('LDIMANCHE', PlanLivrJour); SetControlVisible('BDIMANCHE', PlanLivrJour); SetControlVisible('VDIMANCHE', PlanLivrJour);
  SetControlVisible('LJOURSFERIES', PlanLivrJour); SetControlVisible('BJOURSFERIES', PlanLivrJour); SetControlVisible('VJOURSFERIES', PlanLivrJour);
  SetControlVisible('WPA_BOOLEAN04', False);
  { }
  PlanLivrSetColor;
  { }
  if V_PGI.Sav then
  begin
    SetControlText('WPA_INFOS', GetString('WPA_CODEPARAM') + '/' + GetString('WPA_UTILISATEUR'));
    //SetControlVisible('WPA_INFOS', True);
    //SetControlVisible('PANTRISAV', V_PGI.Sav);
  end;
//GP_DKZ_20071004
  SetControlVisible('WPA_INFOS', V_PGI.Sav);
  SetControlProperty('WPA_LONGVARCHAR02', 'ReadOnly', not V_Pgi.Sav);
  { Découpe le mémo }
  TSCellule := TStringList.Create;
  TSHint := TStringList.Create;
  TSDetail := TStringList.Create;
  try
    wplEclateMemos(GetString('WPA_BLOCNOTE'), TSCellule, TSHint, TSDetail);
    { }
    if Assigned(GetControl('BLOCNOTEITEM')) then
      THMemo(GetControl('BLOCNOTEITEM')).Lines.Assign(TSCellule);
    if Assigned(GetControl('BLOCNOTEHINT')) then
      THMemo(GetControl('BLOCNOTEHINT')).Lines.Assign(TSHint);
    if Assigned(GetControl('BLOCNOTEDETAIL')) then
      THMemo(GetControl('BLOCNOTEDETAIL')).Lines.Assign(TSDetail);
  finally
    TSCellule.free;
    TSHint.free;
    TSDetail.free;
  end;

  { Récupère la liste des champs utilisables }
  if Copy(GetString('WPA_CODEPARAM'), 1, 11) = 'WPLANLIVRGP' then
    wplGetTobFields(ctxwpGP, PlanLivrTobItems, PlanLivrTobFields)
  else
    wplGetTobFields(ctxwpGL, PlanLivrTobItems, PlanLivrTobFields);

  if Assigned(GetControl('ITEMLISTFIELD')) and Assigned(GetControl('HINTLISTFIELD'))
     and Assigned(GetControl('DETAILLISTFIELD')) then
  begin
    PlanLivrAddToList(THListBox(GetControl('ITEMLISTFIELD')));
    PlanLivrUpdateList(THListBox(GetControl('ITEMLISTFIELD')), THRadioGroup(GetControl('ITEMLISTMODE')));
    PlanLivrAddToList(THListBox(GetControl('HINTLISTFIELD')));
    PlanLivrUpdateList(THListBox(GetControl('HINTLISTFIELD')), THRadioGroup(GetControl('HINTLISTMODE')));
    PlanLivrAddToList(THListBox(GetControl('DETAILLISTFIELD')));
    PlanLivrUpdateList(THListBox(GetControl('DETAILLISTFIELD')), THRadioGroup(GetControl('DETAILLISTMODE')));
    PlanLivrAddToList(THListBox(GetControl('TRILISTFIELD')));
    PlanLivrUpdateList(THListBox(GetControl('TRILISTFIELD')), THRadioGroup(GetControl('TRILISTMODE')));
  end;

  if PlanLivrTobItems = nil then
  begin
    SetControlEnabled('PANITEMLISTFIELD', False);
    SetControlEnabled('PANDETAILISTFIELD', False);
    SetControlEnabled('PANHINTLISTFIELD', False);
    SetControlEnabled('TRILISTFIELD', False);
    SetControlEnabled('BTNITEM', False);
    SetControlEnabled('BTNDETAIL', False);
    SetControlEnabled('BTNHINT', False);
    SetControlEnabled('BTNTRIADD', False);
    SetControlEnabled('BTNTRIDEL', False);
  end;

  { Récupère la tob des indicateurs }
  wplLoadTobWPAIndicateurs(PlanLivrTobIndicateurs, GetTob);

  for i := 0 to PlanLivrTobIndicateurs.Detail.Count - 1 do
  begin
    SetControlText(PlanLivrTobIndicateurs.Detail[i].GetString('NAME'), PlanLivrTobIndicateurs.Detail[i].GetString('VALUE'));
  end;

  PlanLivrComboEtatOnEnter(THValComboBox(GetControl('ETATLIVRAISON')));

  { Cadrage contenu des items }
  if Assigned(GetControl('CADRAGEITEM')) then
  begin
    if GetString('WPA_BOOLEAN05') = '1' then
      SetControlProperty('CADRAGEITEM', 'ItemIndex', 1)
    else if GetString('WPA_BOOLEAN05') = '2' then
      SetControlProperty('CADRAGEITEM', 'ItemIndex', 2)
    else
      SetControlProperty('CADRAGEITEM', 'ItemIndex', 0)
  end;

  { Charge la liste des champs formant le tri }
  PlanLivrReadWPA_LONGVARCHAR02;

end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrSetIndicateurs;
var
  i, j: Integer;
  s: HString;
begin
  for i := 0 to PlanLivrTobIndicateurs.Detail.Count - 1 do
  begin
    s := 'VALUE=' + PlanLivrTobIndicateurs.Detail[i].GetString('VALUE');
    for j := 0 to PlanLivrTobIndicateurs.Detail[i].Detail.Count - 1 do
      s := s + ';' + PlanLivrTobIndicateurs.Detail[i].Detail[j].GetString('ETATCODE') + '=' + PlanLivrTobIndicateurs.Detail[i].Detail[j].GetString('ETATCOLOR');
    SetString('WPA_LONGVARCHAR0' + IntToStr(3 + i), s);
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrAddToList(L: THListBox);
var
  i: Integer;
begin
  L.Items.Clear;
  { Génère les listes de choix des champs }
  for i := 0 to PlanLivrTobFields.Detail.Count - 1 do
  begin
    L.Items.AddObject(PlanLivrTobFields.Detail[i].GetString('FIELDNAME'), PlanLivrTobFields.Detail[i]);
  end;
  if L.Items.Count = 0 then
    L.Items.Add(TraduireMemoire('<Aucun>'));
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
function TOM_WPARAM.PlanLivrGetItemCaption(Const R: THRadioGroup; Const T: Tob): HString;
begin
  if R.ItemIndex = 0 then
    Result := T.GetString('FIELDNAME')
  else
    Result := T.GetString('FIELDLIBELLE');
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrUpdateList(Const L: THListBox; Const R: THRadioGroup);
var
  i: Integer;
begin
  for i := 0 to L.Items.Count - 1 do
  begin
    if Assigned(L.Items.Objects[i]) then
      L.Items[i] := PlanLivrGetItemCaption(R, Tob(L.Items.Objects[i]));
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrRecolleLeMemo;
var
  i: Integer;
  TS: TStrings;
begin
  if Assigned(GetControl('BLOCNOTEITEM')) and Assigned(GetControl('BLOCNOTEHINT')) and Assigned(GetControl('BLOCNOTEDETAIL')) then
  begin
    TS := TStringList.Create;
    try
      TS.Add('[0]'); { Début Memo Item }
      for i := 0 to THMemo(GetControl('BLOCNOTEITEM')).Lines.Count - 1 do
        TS.Add(THMemo(GetControl('BLOCNOTEITEM')).Lines[i]);
      TS.Add('[1]'); { Début Memo Bulle d'aide }
      for i := 0 to THMemo(GetControl('BLOCNOTEHINT')).Lines.Count - 1 do
        TS.Add(THMemo(GetControl('BLOCNOTEHINT')).Lines[i]);
      TS.Add('[2]'); { Début Mémo Détail item }
      for i := 0 to THMemo(GetControl('BLOCNOTEDETAIL')).Lines.Count - 1 do
        TS.Add(THMemo(GetControl('BLOCNOTEDETAIL')).Lines[i]);
      { }
      SetString('WPA_BLOCNOTE', TS.Text);
    finally
      TS.Free;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrSetEvents;
begin
  if Assigned(GetControl('BLOCNOTEITEM')) then
    THMemo(GetControl('BLOCNOTEITEM')).OnChange := PlanLivrBlocNoteChange;
  if Assigned(GetControl('BLOCNOTEHINT')) then
    THMemo(GetControl('BLOCNOTEHINT')).OnChange := PlanLivrBlocNoteChange;
  if Assigned(GetControl('BLOCNOTEDETAIL')) then
    THMemo(GetControl('BLOCNOTEDETAIL')).OnChange := PlanLivrBlocNoteChange;

  if Assigned(GetControl('BFOND')) then
    TToolBarButton97(GetControl('BFOND')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BITEM')) then
    TToolBarButton97(GetControl('BITEM')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BSELECTION')) then
    TToolBarButton97(GetControl('BSELECTION')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BTRAIT')) then
    TToolBarButton97(GetControl('BTRAIT')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BCOLROWFIXED')) then
    TToolBarButton97(GetControl('BCOLROWFIXED')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BSAMEDI')) then
    TToolBarButton97(GetControl('BSAMEDI')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BDIMANCHE')) then
    TToolBarButton97(GetControl('BDIMANCHE')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BJOURSFERIES')) then
    TToolBarButton97(GetControl('BJOURSFERIES')).OnClick := PlanLivrSelectColor;
  if Assigned(GetControl('BALERTES')) then
    TToolBarButton97(GetControl('BALERTES')).OnClick := PlanLivrSelectColor;

  if assigned(GetControl('ITEMLISTFIELD')) then
    THListBox(GetControl('ITEMLISTFIELD')).OnDblClick := PlanLivrOnDblClick;
  if assigned(GetControl('HINTLISTFIELD')) then
    THListBox(GetControl('HINTLISTFIELD')).OnDblClick := PlanLivrOnDblClick;
  if assigned(GetControl('DETAILLISTFIELD')) then
    THListBox(GetControl('DETAILLISTFIELD')).OnDblClick := PlanLivrOnDblClick;
  if Assigned(GetControl('BTNITEM')) then
    TToolBarButton97(GetControl('BTNITEM')).OnClick := PlanLivrOnDblClick;
  if Assigned(GetControl('BTNHINT')) then
    TToolBarButton97(GetControl('BTNHINT')).OnClick := PlanLivrOnDblClick;
  if Assigned(GetControl('BTNDETAIL')) then
    TToolBarButton97(GetControl('BTNDETAIL')).OnClick := PlanLivrOnDblClick;

  if assigned(GetControl('BTNTRIADD')) then
    TToolBarButton97(GetControl('BTNTRIADD')).OnClick := PlanLivrTriAddOnClick;
  if assigned(GetControl('TRILISTFIELD')) then
    THListBox(GetControl('TRILISTFIELD')).OnDblClick := PlanLivrTriAddOnClick;
  if assigned(GetControl('BTNTRIDEL')) then
    TToolBarButton97(GetControl('BTNTRIDEL')).OnClick := PlanLivrTriDelOnClick;
  if assigned(GetControl('TRILISTFIELDSELECT')) then
    THListBox(GetControl('TRILISTFIELDSELECT')).OnDblClick := PlanLivrTriDelOnClick;

  if Assigned(GetControl('ITEMLISTMODE')) then
    THRadioGroup(GetControl('ITEMLISTMODE')).OnClick := PlanLivrChangeMode;
  if Assigned(GetControl('HINTLISTMODE')) then
    THRadioGroup(GetControl('HINTLISTMODE')).OnClick := PlanLivrChangeMode;
  if Assigned(GetControl('DETAILLISTMODE')) then
    THRadioGroup(GetControl('DETAILLISTMODE')).OnClick := PlanLivrChangeMode;
  if Assigned(GetControl('TRILISTMODE')) then
    THRadioGroup(GetControl('TRILISTMODE')).OnClick := PlanLivrChangeMode;

  if Assigned(GetControl('LETATLIVRAISON')) then
    THLabel(GetControl('LETATLIVRAISON')).OnClick := PlanLivrLabelOnClick;

  if Assigned(GetControl('ETATLIVRAISON')) then
  begin
    THValComboBox(GetControl('ETATLIVRAISON')).OnEnter := PlanLivrComboEtatOnEnter;
    THValComboBox(GetControl('ETATLIVRAISON')).OnClick := PlanLivrComboEtatOnChange;
  end;

  if Assigned(GetControl('LAFFECTSTOCK')) then
    THLabel(GetControl('LAFFECTSTOCK')).OnClick := PlanLivrLabelOnClick;

  if Assigned(GetControl('AFFECTSTOCK')) then
  begin
    THValComboBox(GetControl('AFFECTSTOCK')).OnEnter := PlanLivrComboEtatOnEnter;
    THValComboBox(GetControl('AFFECTSTOCK')).OnClick := PlanLivrComboEtatOnChange;
  end;

  if Assigned(GetControl('LRETARDS')) then
    THLabel(GetControl('LRETARDS')).OnClick := PlanLivrLabelOnClick;

  if Assigned(GetControl('RETARDS')) then
  begin
    THValComboBox(GetControl('RETARDS')).OnEnter := PlanLivrComboEtatOnEnter;
    THValComboBox(GetControl('RETARDS')).OnClick := PlanLivrComboEtatOnChange;
  end;

  if Assigned(GetControl('LTYPECADENCE')) then
    THLabel(GetControl('LTYPECADENCE')).OnClick := PlanLivrLabelOnClick;

  if Assigned(GetControl('TYPECADENCE')) then
  begin
    THValComboBox(GetControl('TYPECADENCE')).OnEnter := PlanLivrComboEtatOnEnter;
    THValComboBox(GetControl('TYPECADENCE')).OnChange := PlanLivrComboEtatOnChange;
  end;

  if Assigned(GetControl('LAFFECTSTOCK')) then
    THLabel(GetControl('LAFFECTSTOCK')).OnClick := PlanLivrLabelOnClick;

end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrBlocNoteChange(Sender: TObject);
begin
  if not(DS.State in [dsInsert,dsEdit]) then  { BlocNote = THMemo }
    DS.Edit;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrSelectColor(Sender: TObject);
var
  vColor: TColor;
begin
  if Sender is TToolBarButton97 then
  begin
    vColor := execAFColorDialog;
    if vColor <> 0 then
    begin
      case TToolBarButton97(Sender).Tag of
        1: SetInteger('WPA_INTEGER01', vColor);
        2: SetInteger('WPA_INTEGER02', vColor);
        3: SetInteger('WPA_INTEGER03', vColor);
        4: SetInteger('WPA_INTEGER04', vColor);
        5: SetInteger('WPA_INTEGER05', vColor);
        6: SetString('WPA_VARCHAR01', IntToStr(vColor));
        7: SetString('WPA_VARCHAR02', IntToStr(vColor));
        8: SetString('WPA_VARCHAR04', IntToStr(vColor));
      end;
      PlanLivrSetColor;
    end;
  end
  else if Sender is THPanel then
  begin
    if (THPanel(Sender).Tag <> 0) and (Copy(THPanel(Sender).Name, 1, 3) = 'EP_') then
    begin
      vColor := execAFColorDialog;
      if vColor <> 0 then
      begin
        Tob(THPanel(Sender).Tag).SetInteger('ETATCOLOR', vColor);
        THPanel(Sender).Font.Color := vColor;
        { Force le DS en mode Edit }
        if not(DS.State in [dsInsert, dsEdit]) then { BlocNote = THMemo }
          DS.Edit;
      end;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrChangeMode(Sender: TObject);
begin
  if Sender is THRadioGroup then
  begin
    if THRadioGroup(Sender).Name = 'ITEMLISTMODE' then
      PlanLivrUpdateList(THListBox(GetControl('ITEMLISTFIELD')), THRadioGroup(Sender))
    else if THRadioGroup(Sender).Name = 'HINTLISTMODE' then
      PlanLivrUpdateList(THListBox(GetControl('HINTLISTFIELD')), THRadioGroup(Sender))
    else if THRadioGroup(Sender).Name = 'DETAILLISTMODE' then
      PlanLivrUpdateList(THListBox(GetControl('DETAILLISTFIELD')), THRadioGroup(Sender))
    else if THRadioGroup(Sender).Name = 'TRILISTMODE' then
    begin
      PlanLivrUpdateList(THListBox(GetControl('TRILISTFIELD')), THRadioGroup(Sender));
      PlanLivrUpdateList(THListBox(GetControl('TRILISTFIELDSELECT')), THRadioGroup(Sender));
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrOnDblClick(Sender: TObject);
var
  i: Integer;
  T: Tob;
  LH: THListBox;

  function GetMemoName: String;
  begin
    if LH.Name = 'ITEMLISTFIELD' then
      Result := 'BLOCNOTEITEM'
    else if LH.Name = 'HINTLISTFIELD' then
      Result := 'BLOCNOTEHINT'
    else if LH.Name = 'DETAILLISTFIELD' then
      Result := 'BLOCNOTEDETAIL'
    else
      Result := '';
  end;

begin
  if (Sender is THListBox) then
    LH := THListBox(Sender)
  else if Sender is TToolbarButton97 then
  begin
    if TToolbarButton97(Sender).Name = 'BTNITEM' then
      LH := THListBox(GetControl('ITEMLISTFIELD'))
    else if TToolbarButton97(Sender).Name = 'BTNHINT' then
      LH := THListBox(GetControl('HINTLISTFIELD'))
    else if TToolbarButton97(Sender).Name = 'BTNDETAIL' then
      LH := THListBox(GetControl('DETAILLISTFIELD'))
    else
      LH := nil
  end;

  if Assigned(LH) then
  begin
    for i := 0 to LH.Items.Count - 1 do
    begin
      if LH.Selected[i] then
      begin
        T := Tob(LH.Items.Objects[i]);
        TMemo(GetControl(GetMemoName)).Lines.Add('[' + T.GetString('FIELDNAME') + ']');
      end;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrSetWPA_LONGVARCHAR02;
var
  i: Integer;
  NewStr, OldStr: hString;
  LHSel: THListBox;
  T: Tob;
begin
  LHSel := THListBox(GetControl('TRILISTFIELDSELECT'));
  if Assigned(LHSel) then
  begin
    OldStr := GetString('WPA_LONGVARCHAR02'); NewStr := '';
    for i := 0 to LHSel.Items.Count - 1 do
    begin
      T := Tob(LHsel.Items.Objects[i]);
      if Assigned(T) then
      begin
        if NewStr <> '' then
          NewStr := NewStr + ';';
        NewStr := NewStr + T.GetString('FIELDNAME');
      end;
    end;
    if OldStr <> NewStr then
    begin
      if not(DS.State in [dsEdit]) then
        DS.Edit;
      SetString('WPA_LONGVARCHAR02', NewStr);
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrReadWPA_LONGVARCHAR02;
var
  s, FieldName: hString;
  LHSel: THListBox;
  T: Tob;
begin
  LHSel := THListBox(GetControl('TRILISTFIELDSELECT'));
  if Assigned(LHSel) then
  begin
    s := GetString('WPA_LONGVARCHAR02');
    while (s <> '') do
    begin
      FieldName := ReadTokenSt(s);
      if FieldName <> '' then
      begin
        T := PlanLivrTobFields.FindFirst(['FIELDNAME'], [FieldName], False);
        if Assigned(T) then
          LHSel.Items.AddObject('[' + T.GetString('FIELDNAME') + ']', T);
      end;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrTriAddOnClick(Sender: TObject);
var
  i: Integer;
  T: Tob;
  LH, LHSel: THListBox;
begin
  if Assigned(GetControl('TRILISTFIELD')) and Assigned(GetControl('TRILISTFIELDSELECT')) then
  begin
    LH := THListBox(GetControl('TRILISTFIELD'));
    LHSel := THListBox(GetControl('TRILISTFIELDSELECT'));
    for i := 0 to LH.Items.Count - 1 do
    begin
      if LH.Selected[i] then
      begin
        T := Tob(LH.Items.Objects[i]);
        if LHSel.Items.IndexOfObject(T) = -1 then
          LHSel.Items.AddObject('[' + T.GetString('FIELDNAME') + ']', T);
      end;
    end;
    PlanLivrUpdateList(LHSel, THRadioGroup(GetControl('TRILISTMODE')));
    PlanLivrSetWPA_LONGVARCHAR02;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrTriDelOnClick(Sender: TObject);
var
  i: Integer;
  LHSel: THListBox;
begin
  if assigned(GetControl('TRILISTFIELDSELECT')) then
  begin
    LHSel := THListBox(GetControl('TRILISTFIELDSELECT'));
    for i := LHSel.Items.Count -1 downto 0 do
    begin
      if LHSel.Selected[i] then
        LHSel.Items.Delete(i);
    end;
    PlanLivrSetWPA_LONGVARCHAR02;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrComboEtatOnChange(Sender: TObject);
var
  T: Tob;
begin
  if Sender is THValComboBox then
  begin
    T := PlanLivrTobIndicateurs.FindFirst(['NAME'], [THValComboBox(Sender).Name], False);
    if Assigned(T) then
    begin
      T.SetString('VALUE', THValComboBox(Sender).Value);
      DS.Edit;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrComboEtatOnEnter(Sender: TObject);
var
  T: Tob;
  P: THPanel;
  LastTop: Integer;
  i: Integer;

  procedure SetLabelBold(const ControlName: String; const SetBold: boolean);
  begin
    if Assigned(GetControl(ControlName)) then
    begin
      if SetBold then
        ThLabel(GetControl(ControlName)).Font.Style := ThLabel(GetControl(ControlName)).Font.Style + [fsBold]
      else
        ThLabel(GetControl(ControlName)).Font.Style := ThLabel(GetControl(ControlName)).Font.Style - [fsBold]
    end
  end;

begin
  { Nettoie la ScrollBox }
  while (PlanLivrSB.ComponentCount > 0) do
  begin
    PlanLivrSB.Components[0].Free;
  end;
  { }
  if Sender is THValComboBox then
  begin
    T := PlanLivrTobIndicateurs.FindFirst(['NAME'], [THValComboBox(Sender).Name], False);
    if Assigned(T) then
    begin
      { }
      SetLabelBold('LETATLIVRAISON', THValComboBox(Sender).Name = 'ETATLIVRAISON');
      SetLabelBold('LAFFECTSTOCK', THValComboBox(Sender).Name = 'AFFECTSTOCK');
      SetLabelBold('LRETARDS', THValComboBox(Sender).Name = 'RETARDS');
      SetLabelBold('LTYPECADENCE', THValComboBox(Sender).Name = 'TYPECADENCE');

      { Titre de la GroupBox }
      if Assigned(GetControl('L' + THValComboBox(Sender).Name)) then
        SetControlCaption('GBETATS', Format(TraduireMemoire('Valeurs possibles pour l''indicateur : %s'),[THLabel(GetControl('L' + THValComboBox(Sender).Name)).Caption]));

      { Affiche les valeurs possibles pour l'indicateurs dans la GroupBox }
      LastTop := 0;
      for i := 0 to T.Detail.Count - 1 do
      begin
        P := THPanel.Create(PlanLivrSB);
        P.Parent := PlanLivrSB;
        P.Name := 'EP_' + T.Detail[i].GetString('ETATCODE');
        P.BevelOuter := bvNone;
        P.Height := 36;
        LastTop := LastTop + P.Height;
        P.Align := alTop;
        P.Top := LastTop;
        P.Caption := T.Detail[i].GetString('ETATLIB');
        P.Alignment := taCenter;
        P.Color := clWindowText;
        P.Font.Color := T.Detail[i].GetInteger('ETATCOLOR');
        P.Tag := LongInt(T.Detail[i]);
        P.OnClick := PlanLivrSelectColor;
      end;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.PlanLivrLabelOnClick(Sender: TObject);
begin
  if (Sender is THLabel) and Assigned(ThLabel(Sender).FocusControl) then
    PlanLivrComboEtatOnEnter(THValComboBox(ThLabel(Sender).FocusControl))
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF GPAOLIGHT}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF GPAO}
procedure TOM_WPARAM.PlanDechargeSetEvents;
begin
  if Assigned(GetControl('BFOND')) then
    TToolBarButton97(GetControl('BFOND')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BITEM')) then
    TToolBarButton97(GetControl('BITEM')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BSELECTION')) then
    TToolBarButton97(GetControl('BSELECTION')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BTRAIT')) then
    TToolBarButton97(GetControl('BTRAIT')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BCOLROWFIXED')) then
    TToolBarButton97(GetControl('BCOLROWFIXED')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BSAMEDI')) then
    TToolBarButton97(GetControl('BSAMEDI')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BDIMANCHE')) then
    TToolBarButton97(GetControl('BDIMANCHE')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BJOURSFERIES')) then
    TToolBarButton97(GetControl('BJOURSFERIES')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BRATIO0')) then
    TToolBarButton97(GetControl('BRATIO0')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BRATIO1')) then
    TToolBarButton97(GetControl('BRATIO1')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BRATIO2')) then
    TToolBarButton97(GetControl('BRATIO2')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BRATIO3')) then
    TToolBarButton97(GetControl('BRATIO3')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BRATIO4')) then
    TToolBarButton97(GetControl('BRATIO4')).OnClick := PlanDeChargeSelectColor;
  if Assigned(GetControl('BJALON')) then
    TToolBarButton97(GetControl('BJALON')).OnClick := PlanDeChargeSelectColor;
end;

//GP_DKZ_GP_PDC_20070927
procedure TOM_WPARAM.SetPlanDeChargeOnLoad;
var
  PDCJour, PDCGantt: Boolean;
  sTitre : string;
//GP_20080317_DKZ_GP14898
  i                : integer;
begin
  { }
  PDCJour  := RightStr(GetString('WPA_CODEPARAM'), 1) = 'J';
  PDCGantt := pos('GAN', GetString('WPA_CODEPARAM')) > 0;
  if PDCGantt then
    sTitre := TraduireMemoire('Gantt par ordre')
  else
    sTitre := TraduireMemoire('Graphique par famille de ressource');
  { Titre}
  Ecran.Caption := TraduireMemoire('Paramètres d''affichage')+' - '+ sTitre;

  SetControlVisible('WPA_BOOLEAN01', PDCJour);
  SetControlVisible('LSAMEDI', PDCJour); SetControlVisible('BSAMEDI', PDCJour); SetControlVisible('VSAMEDI', PDCJour);
  SetControlVisible('LDIMANCHE', PDCJour); SetControlVisible('BDIMANCHE', PDCJour); SetControlVisible('VDIMANCHE', PDCJour);
  SetControlVisible('LJOURSFERIES', PDCJour); SetControlVisible('BJOURSFERIES', PDCJour); SetControlVisible('VJOURSFERIES', PDCJour);
  SetControlVisible('LJALON', PDCGantt); SetControlVisible('BJALON', PDCGantt); SetControlVisible('VJALON', PDCGantt);
//GP_20080708_DKZ_GP15150 Déb
  SetControlVisible('TWPA_DOUBLE08', PDCGantt); SetControlVisible('WPA_DOUBLE08', PDCGantt);
  SetControlVisible('TWPA_DOUBLE09', PDCGantt); SetControlVisible('WPA_DOUBLE09', PDCGantt);
//GP_20080708_DKZ_GP15150 Fin
  SetControlVisible('WPA_BOOLEAN08', PDCGantt);
//GP_20080708_DKZ_GP15150 Déb
  SetControlVisible('WPA_BOOLEAN09', PDCGantt);
  SetControlEnabled('WPA_DOUBLE08', (PDCGantt) and not (GetBoolean('WPA_BOOLEAN09')));
  SetControlEnabled('WPA_DOUBLE09', (PDCGantt) and not (GetBoolean('WPA_BOOLEAN09')));
//GP_20080708_DKZ_GP15150 Fin
  SetControlVisible('WPA_BOOLEAN02', PDCGantt);
//GP_BUG810_DKZ GP_PDC 20071008
  SetControlVisible('WPA_BOOLEAN07', PDCGantt);
//GP_20080317_DKZ_GP14898 Déb
  SetControlVisible('LRATIO', not PDCGantt);
  for i:=0 to 4 do
  begin
    SetControlVisible('LRATIO'+IntToStr(i), not PDCGantt);
    SetControlVisible('BRATIO'+IntToStr(i), not PDCGantt);
    SetControlVisible('VRATIO'+IntToStr(i), not PDCGantt);
  end;
//GP_20080317_DKZ_GP14898 Fin
  { }
  PlanDeChargeSetColor;
  { }
//GP_DKZ_20071004 Déb
  if V_PGI.Sav then
    SetControlText('WPA_INFOS', GetString('WPA_CODEPARAM') + '/' + GetString('WPA_UTILISATEUR'));
  SetControlVisible('WPA_INFOS', V_PGI.Sav);
//GP_DKZ_20071004 Fin
end;

procedure TOM_WPARAM.PlanDeChargeSelectColor(Sender: TObject);
var
  vColor: TColor;
begin
  if Sender is TToolBarButton97 then
  begin
    vColor := execAFColorDialog;
    if vColor <> 0 then
    begin
      case TToolBarButton97(Sender).Tag of
        1: SetInteger('WPA_INTEGER01', vColor);
        2: SetInteger('WPA_INTEGER02', vColor);
        3: SetInteger('WPA_INTEGER03', vColor);
        4: SetInteger('WPA_INTEGER04', vColor);
        5: SetInteger('WPA_INTEGER05', vColor);
        6: SetString('WPA_VARCHAR01', IntToStr(vColor));
        7: SetInteger('WPA_DOUBLE02', vColor);
        8: SetInteger('WPA_DOUBLE03', vColor);
        9: SetInteger('WPA_DOUBLE04', vColor);
        10: SetInteger('WPA_DOUBLE05', vColor);
        11: SetInteger('WPA_DOUBLE06', vColor);
        12: SetInteger('WPA_DOUBLE07', vColor);
      end;
      PlanDeChargeSetColor;
    end;
  end;
end;

procedure TOM_WPARAM.PlanDeChargeSetColor;
begin
  SetControlProperty('VFOND', 'Color', GetInteger('WPA_INTEGER01'));
  SetControlProperty('VSELECTION', 'Color', GetInteger('WPA_INTEGER02'));
  SetControlProperty('VSAMEDI', 'Color', GetInteger('WPA_INTEGER03'));
  SetControlProperty('VDIMANCHE', 'Color', GetInteger('WPA_INTEGER04'));
  SetControlProperty('VJOURSFERIES', 'Color', GetInteger('WPA_INTEGER05'));
  if GetString('WPA_VARCHAR01') <> '' then
    SetControlProperty('VTRAIT', 'Color', ValeurI(GetString('WPA_VARCHAR01')));
  SetControlProperty('VRATIO0', 'Color', GetInteger('WPA_DOUBLE02'));
  SetControlProperty('VRATIO1', 'Color', GetInteger('WPA_DOUBLE03'));
  SetControlProperty('VRATIO2', 'Color', GetInteger('WPA_DOUBLE04'));
  SetControlProperty('VRATIO3', 'Color', GetInteger('WPA_DOUBLE05'));
  SetControlProperty('VRATIO4', 'Color', GetInteger('WPA_DOUBLE06'));
  SetControlProperty('VJALON', 'Color' , GetInteger('WPA_DOUBLE07'));
end;
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFDEF GPAO}
{$IFDEF AFFAIRE}
{$IF not(Defined(ERADIO) or Defined(EAGLSERVER))}
procedure TOM_WPARAM.WPA_VARCHAR01_OnChange(Sender: TObject);
begin
  if EcranIs('WPA_GETAFCONTEXTE') then
  begin
    if GetControlText('WPA_VARCHAR01')=sContexteORIGINEWLB then
      SetDateTime('WPA_DATE01', Date);

    SetControlProperty('WPA_DATE01', 'ENABLED', GetControlText('WPA_VARCHAR01')<>sContexteORIGINEWLB );
  end;
end;
{$IFEND !ERADIO && !EAGLSERVER}
{$ENDIF AFFAIRE}
{$ENDIF GPAO}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOM_WPARAM.BValiderOnClick(Sender: TObject);
begin
  if LastError = 0 then
    tfFiche(Ecran).Retour := wTrue;

  inherited BValiderOnClick(Sender);
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFDEF GPAO}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF AFFAIRE}
procedure TOM_WPARAM.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Tiers: THEdit);
begin
  InHerited;
  Aff   := THEdit(GetControl('WPA_VARCHAR05'));
  Aff1  := THEdit(GetControl('AFFAIRE1'));
  Aff2  := THEdit(GetControl('AFFAIRE2'));
  Aff3  := THEdit(GetControl('AFFAIRE3'));
  Aff4  := THEdit(GetControl('AVENANT'));
  Tiers := THEdit(GetControl('WPA_VARCHAR02'));
end;
{$ENDIF AFFAIRE}
{$ENDIF ERADIO}
{$ENDIF EAGLSERVER}
{$ENDIF GPAO}

{$IFNDEF PGIMAJVER}
{$IF not Defined(EAGLSERVER) and not Defined(ERADIO)}
{$IFDEF AFFAIRE}
procedure TOM_WPARAM.BRECHAFFAIRE_OnClick(Sender: TObject);
Var
  aAffaire: MyArrayValue;
begin
  { Pour avoir accès à toutes les affaires }
  fTypeAffaire        := 'TOUS';
  InHerited;

  if (GetControlText('AFFAIRE1')<>'') then
  begin
    aAffaire := WGetSqlFieldsValues(['AFF_LIBELLE'], 'AFFAIRE', 'AFF_AFFAIRE="'+GetString('WPA_VARCHAR05')+'"');
    if Assigned(aAffaire) then
      LBAFFAIRE.Caption  := aAffaire[0]
  end
  else
    LBAFFAIRE.Caption  := '';
end;
{$ENDIF AFFAIRE}
{$IFEND !EAGLSERVER && !ERADIO}
{$ENDIF !PGIMAJVER}

Initialization
  RegisterClasses([TOM_WPARAM]);
end.
