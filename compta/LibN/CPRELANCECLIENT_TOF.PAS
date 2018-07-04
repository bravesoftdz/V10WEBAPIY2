{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/06/2004
Modifié le ... : 24/11/2005
Description .. : Source TOF de la FICHE : CPRELANCECLIENT ()
Suite ........ : - FQ 15767 - CA - 05/07/2005 : Ne plus utiliser 
Suite ........ : E_REFPOINTAGE sous peine de supprimer les référence 
Suite ........ : de pointage utilisée en pointage sur journal.
Suite ........ : Desormais les mouvements à éditer sont répertoriés 
Suite ........ : uniquement par le champ E_NUMPIECEINTERNE quit 
Suite ........ : contient "EDITREL;'+Code_modèle_lettre_relance. Ce
Suite ........ : champ est remis à blanc en fin de traitement.
Suite ........ : 
Suite ........ : JP 24/11/05 : FQ 16499 : problème d'affectation des listes
Mots clefs ... : TOF;CPRELANCECLIENT
*****************************************************************}

unit CPRELANCECLIENT_TOF;

interface

uses
    StdCtrls,
    Controls,
    Classes,
    forms,
    sysutils,
    ComCtrls,
    Graphics, // TColor
    FileCtrl,
    Menus, // TPopUpMenu
    Dialogs, // TFindDialog
{$IFDEF EAGLCLIENT}
    eMul,
    MainEAGL,
    UtileAGL,
{$ELSE}
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    mul,
    hdb,
    dbGrids,
    Fe_Main,
    EdtRDoc,
    EdtREtat,
    Justisol, // JustifSolde
    UtilEdt, // InitCritEdt
//    QRJuSold, // JustSoldeZoom
{$IFNDEF CCS3}
    CoutTier,
{$ENDIF}
{$ENDIF}
    CritEdt, // TCritEdt
    uTob,
    HCtrls,
    HEnt1,
    Htb97,
    HMsgBox,
    UtilPgi,
    UTOF,
    URecupSQLModele,
    Ent1, // pour positionneEtabUser
    LettUtil, // InitTablesLibresTiers
    EncUtil, // WhereGeneCritEcr
    HQry, // THQuery
    HStatus, // Barre défilement
    ParamSoc, // SetParamSoc
    SaisUtil, // TMSEncaDeca
    SaisComm, // OBMToIdent
    ULibEcriture, // WhereEcritureTob
    SaisBor, // ChargeSaisieFolio
    Saisie, // LanceSaisie
    AGLInit,
    Relance_TOM // CCLanceFiche_ParamRelance
    ;

const
    // Index des colonnes de la grilles des clients
    RT_Aux = 1;
    RT_Int = 2;
    // Index des colonnes de la grille des échéances
    RE_Num = 1;
    RE_Pai = 2;
    RE_Ech = 3;
    RE_Tot = 4;
    RE_Niv = 5;
    RE_Dif = 6;
    RE_SEL = 7;

{$IFDEF CCS3}
    RELNIVOMAX = 2; // Nivo maximum de relance pour S3
{$ELSE}
    RELNIVOMAX = 6; // Nivo maximum de relance pour S5, S7
{$ENDIF CCS3}

procedure RelanceClient(Manuelle, OkTraite : boolean);
procedure CPLanceFiche_RelanceClient(vStRange, vStLequel, vStArgs : string);

type
    TOF_CPRELANCECLIENT = class(TOF)
        procedure OnNew; override;
        procedure OnDelete; override;
        procedure OnUpdate; override;
        procedure OnLoad; override;
        procedure OnArgument(S : string); override;
        procedure OnDisplay; override;
        procedure OnClose; override;
        procedure OnCancel; override;
    private

        // -------------------------------------------------------
        // Indicateur d'appel de la fiche (Mode de fonctionnement)
        // -------------------------------------------------------
        ModeManuel : boolean; // manuel ou auto ?
        ModeTraite : boolean; // gestion des traites ou autres moyens de paiement

        // -------
        // Données
        // -------
        TobClient : TOB; // Tob 2 nivo contenant les lignes de la grille client >> lignes de la grilles des échéances
        TobRelance : TOB; // Tob contenant les paramètres de relance chargés depuis la base
        GClient : THGrid; // Grille des cumuls clients
        GEche : THGrid; // Grille des échéances

        // ----------
        // Traitement
        // ----------
        FNbRelance : Integer; // Nb de relance généré
        FBoAppliqScore : Boolean; // Indicateur de maj du scoring
        FBoYATP : boolean; // Presence de tiers payeur ??
        FMSED : TMSEncaDeca; // Paramètres d'éditions par spooler
        FBoNivoMax : Boolean; // Indicateur de nivo de relance max atteind
        FDtNowFutur : TDateTime; // Date modif de pièce pour uniformiser toutes les lignes
        FBoGenePremier : Boolean; // Indicateur de 1ère relance
        FDtDateRel : TDateTime; // Date de relance saisie
        FInNivoRef : Integer; // Niveau de référence saisi
        FBoSAVCodeOmr : boolean; // pour mettre de coté le mode par défaut d'édition des code barres
        FBoChargEnCours : boolean; // Indicateur de chargement de l'interface en cours (pour éviter 1ère recherche...)

        // ---------------------------
        // Variables Gestion interface
        // ---------------------------
        RT_Aff : Integer; // Index colonne Total
        RT_Tot : Integer; // Index colonne Total
        RT_Niv : Integer; // Index colonne Niveau relance
        RT_Der : Integer; // Index colonne Date dernière relance
        RT_Sel : Integer; // Index colonne invisible pour paramétrage colonne
        FBoToutSel : Boolean; // Indicateur tout sélectionner
        FBoFlipAuto : Boolean; // Indicateur traitement sur FlipSelection des grilles
        FBoFindFirst : Boolean; // Indicateur de recherche dans la liste

        // ----------
        // Composants
        // ----------
        QEcr : THQuery;
{$IFDEF EAGLCLIENT}
        FListe : THGrid;
{$ELSE}
        FListe : THDBGrid;
{$ENDIF}
        Pages : TPageControl;
        BSelectAll : TToolbarButton97;
        BMenuzoom : TToolbarButton97;
        POPS : TPopUpMenu;
        PopZoom : TPopUpMenu;
        PopNiveau : TPopUpMenu;
        FindMvt : TFindDialog;

        // EVT du parent
        OnParentEcranResize : TNotifyEvent;
        OnParentBChercheClick : TNotifyEvent;

        // ---------------
        // Initialisations
        // ---------------
        procedure AnalyseParametre(vStArg : string);
        procedure InitComposants;
        procedure InitEvenements;
        procedure InitCriteres;
        procedure InitVariables;
        procedure InitPresentation;

        // -----------------------------
        // Mise en place de la recherche
        // -----------------------------
        function WhereCritMP : string;
        procedure RenseigneXXWhere;
        procedure RenseigneXXOrderBy;

        // -----------------------------
        // Chargement des données
        // -----------------------------
        function EstLigneOk : boolean;
        procedure RenseigneLaTob;
        procedure VideLesListes;

        // -----------------------------
        // Gestion de la grille des clients
        // -----------------------------
        procedure InitGClient;
        procedure GClientTailleCol;
        procedure RemplitGClient;
        procedure GClientToutSelect(vBoTout : Boolean);
        procedure GClientMajLigne(vInRow : Integer);
        procedure GClientRecalculLigne(vInRow : Integer;vBoMajGrille : Boolean);
        procedure GClientRowEnter(Sender : TObject;Ou : Longint;var Cancel : Boolean;Chg : Boolean);
        procedure GClientFlipSelection(Sender : TObject);
        procedure GClientDoubleClick(Sender : TObject);
        procedure GClientBeforeFlip(Sender : TObject;ARow : Integer;var Cancel : Boolean);
        procedure FlapLigne;

        // -----------------------------
        // Gestion de la grille des échéances
        // -----------------------------
        procedure InitGEche;
        procedure GEcheTailleCol;
        procedure RemplitGEche;
        procedure GEcheToutSelect(vInRow : Integer;vBoTout : Boolean;vBoMajGrille : Boolean);
        procedure GEcheMajLigne(vInRow : Integer);
        procedure GEcheFlipSelection(Sender : TObject);
        procedure GEcheDoubleClick(Sender : TObject);
        procedure GEcheBeforeFlip(Sender : TObject;ARow : Integer;var Cancel : Boolean);

        // -----------
        // EVT Boutons
        // -----------
        procedure BSelectAllClick(Sender : TObject);
        procedure BValideClick(Sender : TObject);
        procedure BImprimerClick(Sender : TObject);
        procedure BRechercherClick(Sender : TObject);
        procedure FindMvtFind(Sender : TObject);

        // --------
        // EVT Zoom
        // --------
        procedure ZoomEcriture    (Sender : TObject);
        procedure ZoomTiers       (Sender : TObject);
        procedure ZoomParam       (Sender : TObject);
        procedure ZoomCout        (Sender : TObject);
//        procedure ZoomEdtJustif   (Sender : TObject);
        procedure ZoomEdtJustifNew(Sender : TObject); {JP 28/10/05 : FQ 16881}
        procedure NiveauModif     (Sender : TObject);
        procedure NiveauUp        (Sender : TObject);
        procedure NiveauDown      (Sender : TObject);

        // ----------
        // EVT autres
        // ----------
        procedure SpoolerClick(Sender : TObject);
        procedure RuptureAffaireClick(Sender : TObject);
        procedure POPSPopup(Sender : TObject);
        procedure UnModeleClick(Sender : TObject);
        procedure OnEcranResize(Sender : TObject);
        procedure OnCritModif(Sender : TObject);
        procedure BChercheClick(Sender : TObject);
        procedure OnAfterShow;
        procedure AuxiElipsisClick(Sender : TObject);
        function  GetFormatCol : string; // FQ 20048

        // ----------
        // Traitement
        // ----------
        procedure ValideRelance;
        function EclateEches(vTobCli : TOB;vStFamille, vStAuxi, vStAffaire : string;vInNivScore, vInMaxNiv : Integer;ListeModele : HTStringList) : Boolean;
        procedure EclateEchesParModele(ListeModele : HTStringList);
        function MarqueEches(vInNivo : Integer;vTobCli : TOB;vStModele : string) : boolean;
        procedure MajListeTiersparModele(vStFamille, vStAuxi, vStAffaire, vStModele : string;vInNivScore, vInMaxNiv : Integer;ListeModele : HTStringList);
        procedure MajEcritures(vStAuxi, vStModele : string;vInMaxNiv : integer);
        procedure MajEcrScore(vStAuxi, vStModele : string;vInNivScore, vInMaxNiv : integer);
        procedure MajNiveau7(vStAuxi, vStModele : string);
        procedure EditeRelancesTP;
        procedure ModificationNiveau(Montant : integer;FromZero : Boolean;BlocNote : string;MajBlocNote : boolean);

        // ------------------
        // Fonctions diverses
        // ------------------
        function GetEcartQ : Integer;
        function GetRelance(vStFamille : string) : TOB;
        function ExisteModele(vStModele : string;var vCodeM : string) : boolean;
        procedure RecupMSED(lBoAvecTP : Boolean);
        function GetConditionSQLEcr(vInNivo : Integer;vBoNonEchu : Boolean;vStModele : string) : string;
        function GetOrderBy : string;
        procedure VideLaStringList(L : HTStrings);
        function SQLRetrouve(vStAuxi, vStModele : string) : string;
    end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  CPTIERS_TOM, {FicheTiers}
  {$IFNDEF IMP}
  UTOFCPGLAUXI {JP 28/10/05 : FQ 16881 : Justificatif de solde},
  {$ENDIF IMP}
  CPRELANCENIVEAU_TOF {CPLAnceFIche_RelanceNiveau}
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

// =====================================
// ======  LANCEMENT DE LA FICHE =======
// =====================================

procedure RelanceClient(Manuelle, OkTraite : boolean);
var
    lStArg : string;
begin
    if (Manuelle) then lStArg := 'MODE=MANUEL;'
    else lStArg := 'MODE=AUTO;';

    if (OkTraite) then lStArg := lStArg + 'TYPE=TRAITE'
    else lStArg := lStArg + 'TYPE=AUTRE';

    CPLanceFiche_RelanceClient('', '', lStArg);
end;

// ===========================================
// ======  APPEL LANCEMENT DE LA FICHE =======
// ===========================================

procedure CPLanceFiche_RelanceClient(vStRange, vStLequel, vStArgs : string);
begin
    AGLLanceFiche('CP', 'CPRELANCECLIENT', vStRange, vStLequel, vStArgs);
end;

// ============================================
// ============ EVENEMENT DE LA TOF ===========
// ============================================

procedure TOF_CPRELANCECLIENT.OnNew;
begin
    inherited;

    if (GetControlText('FFiltres') = '') then InitCriteres;
end;

procedure TOF_CPRELANCECLIENT.OnDelete;
begin
    inherited;
end;

procedure TOF_CPRELANCECLIENT.OnUpdate;
begin
    inherited;
    // CA - 05/07/2005 : ce code auparavant dans OnDisplay a été déplcé ici sinon problème avec AGL580
    // Remplissage des grilles
    RenseigneLaTob;
    GClient.Row := 1;
    GClientToutSelect(true);
    
end;

procedure TOF_CPRELANCECLIENT.OnLoad;
begin
    inherited;


    VideLesListes;

    FDtDateRel := StrToDate(GetControlText('DATERELANCE'));
    if (GetControlText('NIVEAURELANCE') <> '') then FInNivoRef := StrToInt(GetControlText('NIVEAURELANCE'))
    else FInNivoRef := -1;

    RenseigneXXWhere;
    RenseigneXXOrderBy;

    // MAJ indicateurs
    FBoToutSel := false;

end;

procedure TOF_CPRELANCECLIENT.OnArgument(S : string);
begin
    FBoChargEnCours := true;

    inherited;

    // Analyse des paramètres
    AnalyseParametre(S);

    // Initialisation de l'interface
    InitVariables;
    InitComposants;
    InitEvenements;
    InitPresentation;
    InitCriteres;

    if GetParamSocSecur('SO_CPMULTIERS', false) then
    begin
      THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
      THEdit(GetControl('E_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
    end;
end;

procedure TOF_CPRELANCECLIENT.OnClose;
begin
    inherited;

    TobClient.ClearDetail;
    FreeAndNil(TobClient);
    TobRelance.ClearDetail;
    FreeAndNil(TobRelance);

    V_PGI.MiseSousPli := FBoSAVCodeOmr;
end;

procedure TOF_CPRELANCECLIENT.OnDisplay;
begin
    inherited;

end;

procedure TOF_CPRELANCECLIENT.OnCancel;
begin
    inherited;
end;

procedure TOF_CPRELANCECLIENT.AnalyseParametre(vStArg : string);
var
    lStValue : string;
    lStChaine : string;
begin
    ModeManuel := True;
    ModeTraite := False;

    while vStArg <> '' do
    begin
        lStChaine := ReadTokenSt(vStArg);

        // Stocke le XXX de "MODE=XXX" dans la variable
        if (Pos('MODE=', Uppercase(lStChaine)) > 0) then
        begin
            lStValue := Copy(lStChaine, 6, length(lStChaine));
            ModeManuel := (lStValue <> 'AUTO');
        end

        // Stocke le XXX de "TYPE=XXX" dans la variable
        else
            if (Pos('TYPE=', Uppercase(lStChaine)) > 0) then
        begin
            lStValue := Copy(lStChaine, 6, length(lStChaine));
            ModeTraite := (lStValue = 'TRAITE');
        end;
    end;

end;

procedure TOF_CPRELANCECLIENT.InitComposants;
begin
    // Controles - Formes
    QEcr := THQuery(GetControl('Q', true));

{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FListe', true));
{$ELSE}
    FListe := THDBGrid(GetControl('FListe', true));
{$ENDIF}

    Pages := TPageControl(GetControl('Pages', true));
    POPS := TPopUpMenu(GetControl('POPS', true));
    PopZoom := TPopUpMenu(GetControl('POPZOOM', true));
    PopNiveau := TPopUpMenu(GetControl('POPNIVEAU', true));

    // grilles
    GClient := THGrid(GetControl('GClient', True)); // Grille des cumuls clients
    GEche := THGrid(GetControl('GEche', True)); // Grille des échéances

    // boutons
    BSelectAll := TToolbarButton97(GetControl('bSelectAll', true));
    BMenuZoom := TToolbarButton97(GetControl('BMenuZoom', true));
    FindMvt := TFindDialog(GetControl('FindDialog', true));
end;

procedure TOF_CPRELANCECLIENT.InitEvenements;
var
    lBouton : TToolbarButton97;
begin
    // EVT Fiche
    if Assigned(TFMul(Ecran).OnResize) then OnParentEcranResize := TFMul(Ecran).OnResize;
    TFMul(Ecran).OnResize := OnEcranResize;
    TFMul(Ecran).OnAfterFormShow := OnAfterShow;

    // EVT Recherche
    lBouton := TToolbarButton97(GetControl('BCHERCHE'));
    if Assigned(lBouton.OnClick) then OnParentBChercheClick := lBouton.OnClick;
    lBouton.OnClick := BChercheClick;

    // Grille des clients
    GClient.onFlipSelection := GClientFlipSelection;
    GClient.OnRowEnter := GClientRowEnter;
    GClient.OnDblClick := GClientDoubleClick;
    GClient.OnBeforeFlip := GClientBeforeFlip;

    // Grilles des échéances
    GEche.onFlipSelection := GEcheFlipSelection;
    GEche.OnDblClick := GEcheDoubleClick;
    GEche.OnBeforeFlip := GEcheBeforeFlip;

    // Boutons
    BSelectAll.OnClick := BSelectAllClick;
    TToolbarButton97(GetControl('BOUVRIR', true)).OnClick := BValideClick;
    TToolbarButton97(GetControl('BIMPRIMER', true)).OnClick := BImprimerClick;
    TToolbarButton97(GetControl('BRechercher', true)).OnClick := BRechercherClick;
    FindMvt.OnFind := FindMvtFind;

    // Menu popup principal
    POPS.OnPopup := POPSPopup;

    // Menu Zoom
    PopZoom.Items[0].OnClick := ZoomEcriture;
    PopZoom.Items[1].OnClick := ZoomTiers;

  {JP 28/10/05 : FQ 16881 : On branche le justificatif de solde ...}
  PopZoom.Items[2].OnClick := ZoomEdtJustifNew;
  {$IFDEF EAGLCLIENT}
  {... Et on cache la gestion des coûts}
  PopZoom.Items[4].Visible := False;
  {$ELSE}
  //  PopZoom.Items[2].OnClick := ZoomEdtJustif;
  {$ENDIF EAGLCLIENT}

    PopZoom.Items[3].OnClick := ZoomParam;
    PopZoom.Items[4].OnClick := ZoomCout;
    TToolbarButton97(GetControl('BZOOMPIECE', true)).OnClick := ZoomEcriture;
    TToolbarButton97(GetControl('BZOOMTIERS', true)).OnClick := ZoomTiers;
//    TToolbarButton97(GetControl('BZOOMSOLDE', true)).OnClick := ZoomEdtJustif;
    TToolbarButton97(GetControl('BZOOMPARAM', true)).OnClick := ZoomParam;
    TToolbarButton97(GetControl('BZOOMCOUT', true)).OnClick := ZoomCout;

    // Menu échéance
    PopNiveau.Items[0].OnClick := NiveauModif;
    PopNiveau.Items[1].OnClick := NiveauUp;
    PopNiveau.Items[2].OnClick := NiveauDown;
    TToolbarButton97(GetControl('BNIVOMODIF', true)).OnClick := NiveauModif;
    TToolbarButton97(GetControl('BNIVOUP', true)).OnClick := NiveauUp;
    TToolbarButton97(GetControl('BNIVODOWN', true)).OnClick := NiveauDown;

    // Critères / paramètres
    TCheckBox(GetControl('RUPTUREAFFAIRE', true)).OnClick := RuptureAffaireClick;
    TCheckBox(GetControl('SPOOLER', true)).OnClick := SpoolerClick;
    TCheckBox(GetControl('UNMODELE', true)).OnClick := UnModeleClick;

    TCheckBox(GetControl('AVECTRAITE', true)).OnClick := OnCritModif;
    { FQ 21824 BVE 12.11.07 }
    THEdit(GetControl('MONTANTMIN', true)).OnChange := OnCritModif;
    THEdit(GetControl('MONTANTMAX', true)).OnChange := OnCritModif;
    { END FQ 21824 }
    THEdit(GetControl('DATERELANCE', true)).OnChange := OnCritModif;
end;

procedure TOF_CPRELANCECLIENT.InitCriteres;
begin
    // Onglet comptes
    SetControlText('DateRelance', DateToStr(V_PGI.DateEntree));

    // onglet Ecritures
    SetControlText('E_JOURNAL', '');
    SetControlText('E_NATUREPIECE', '');
    SetControlText('E_DEVISE', V_PGI.DevisePivot);
    PositionneEtabUser(THValComboBox(GetControl('E_ETABLISSEMENT', True)));

    SetControlText('E_DATECOMPTABLE', StDate1900);
    SetControlText('E_DATECOMPTABLE_', StDate2099);
    SetControlText('E_DATEECHEANCE', StDate1900);
    SetControlText('E_DATEECHEANCE_', StDate2099);

    // onglet tiers
    SetControlText('T_RELANCEREGLEMENT', '');
    SetControlText('T_RELANCETRAITE', '');
    SetControlText('T_SECTEUR', '');

    // onglet paramètres
    SetControlText('NIVEAURELANCE', '');
    SetControlText('ORDRERELANCE', 'CPT');
    SetControlText('SENS', 'D');
end;

procedure TOF_CPRELANCECLIENT.InitPresentation;
var
    lInCpt : Integer;
begin

{$IFDEF EAGLCLIENT}
    // Spécif CWAS
    // Pas de gestin du spooler en eAGL
    TTabSheet(GetControl('PSPOOLER', True)).TabVisible := false;
    SetControlVisible('BNEXT', false);
    SetControlVisible('BPREV', false);
{$ENDIF EAGLCLIENT}

    // Onglet avec zone invisible
    TTabSheet(GetControl('PINVISIBLE', True)).TabVisible := false;

    // Interface
    if (ModeManuel) then
    begin
        if (ModeTraite) then Ecran.Caption := 'Relance manuelle des effets en retour d''acceptation'
        else Ecran.Caption := 'Relance manuelle des modes de paiement en retard de règlement';
        Ecran.HelpContext := 7568000;
    end
    else
    begin
        SetControlVisible('BMenuNivo', false);
        SetControlProperty('BMenuNivo', 'Tag', 0);
        SetControlProperty('BNivoModif', 'Tag', 0);
        SetControlProperty('BNivoUp', 'Tag', 0);
        SetControlProperty('BNivoDown', 'Tag', 0);
        SetcontrolProperty('BMenuZoom', 'Left', TButton(GetControl('BMenuNivo', true)).Left);
        if (ModeTraite) then Ecran.Caption := 'Relances automatiques des effets en retour d''acceptation'
        else Ecran.Caption := 'Relance automatique des modes de paiement en retard de règlement';
        Ecran.HelpContext := 7565000;
    end;
    Ecran.Caption:=TraduireMemoire(Ecran.Caption) ;

    UpdateCaption(Ecran);

    // Grille DB et panel cumul invisible !
    SetControlVisible('FLISTE', false);
    SetControlVisible('PCUMUL', false);

    // onglet comptes
    SetControlVisible('ChoixDate', ModeTraite);
    SetControlVisible('TSpinDelta', ModeTraite);
    SetControlVisible('SpinDelta', ModeTraite);
    SetControlVisible('AvecTraite', not ModeTraite);
    SetControlVisible('TAvecTraite', not ModeTraite);
    if (ModeTraite) then SetControlChecked('AvecTRaite', false);

    //onglet tiers
    SetControlVisible('T_RELANCEREGLEMENT', not ModeTraite);
    SetControlVisible('TT_RELANCEREGLEMENT', not ModeTraite);
    SetControlVisible('T_RELANCETRAITE', ModeTraite);
    SetControlVisible('TT_RELANCETRAITE', ModeTraite);

    // onglet tables libres
    InitTablesLibresTiers(TTabSheet(GetControl('PTablesLibres', true)));
    for lInCpt := 0 to 9 do SetControlProperty('T_TABLE' + IntToStr(LInCpt), 'DATATYPE', 'TZNATTIERS' + IntToStr(LInCpt));

    // onglet paramètres
    if (not VH^.OuiTP) then
    begin
        SetControlVisible('MODELETP', false);
        SetControlVisible('TMODELETP', false);
    end;

    // grille des cumuls clients
    InitGClient;

    // grille des échéances
    InitGEche;

  // Onglet Spooler ( Gestion du code omr pour ccmp uniquement )
{$IFDEF CCMP}
  SetControlChecked('CODEOMR', V_PGI.MiseSousPli);
  SetControlVisible('CODEOMR', true);
{$ELSE}
  SetControlVisible('CODEOMR', false);
{$ENDIF}

  {JP 04/06/07 : FQ 20060 : En relances automatiques, on cache les options}
  SetControlVisible('BSELECTALL', ModeManuel);
  SetControlVisible('BMENUNIVO' , ModeManuel);
end;

procedure TOF_CPRELANCECLIENT.InitVariables;
begin
    TobClient := TOB.Create('VGRILLES', nil, -1);
    TobRelance := TOB.Create('VRELANCES', nil, -1);
    FBoSAVCodeOmr := V_PGI.MiseSousPli;
    FBoFlipAuto := false;
    FBoToutSel := false;
    FBoYATP := false;
    FBoFindFirst := true;
    {JP 24/11/05 : FQ 16499 : Modification de la gestion des listes et filtres avec les
                   derniers Agl (~6.8.1.150)}
    if (ModeTraite) then TFMul(Ecran).SetDBListe('CPRLANCETRA')  // Attention, il manque le E dans le nom de la liste (SOCREF723) CPRLANCETRA à la place de CPRELANCETRA  
                    else TFMul(Ecran).SetDBListe('CPRELANCEDIV');
end;

procedure TOF_CPRELANCECLIENT.InitGClient;
begin
    GClient.RowCount := 2;

    if (GetCheckBoxState('RUPTUREAFFAIRE') = cbChecked) then
    begin
        RT_Aff := 3;
        RT_Tot := 4;
        RT_Niv := 5;
        RT_Der := 6;
        RT_Sel := 7;
        GClient.ColCount := 8;
        GClient.Width := 520;
    end
    else
    begin
        RT_Aff := 0;
        RT_Tot := 3;
        RT_Niv := 4;
        RT_Der := 5;
        RT_Sel := 6;
        GClient.ColCount := 7;
        GClient.Width := 420;
    end;

    // Colonne sélection invisible
    GClient.ColLengths[RT_Sel] := -1;

    // Colonne Auxiliaire
    GClient.Cells[RT_Aux, 0] := traduireMemoire('Auxiliaire');
    GClient.ColAligns[RT_Aux] := taLeftJustify;

    // Colonne Intitulé
    GClient.Cells[RT_Int, 0] := TraduireMemoire('Intitulé') ;
    GClient.ColAligns[RT_Int] := taLeftJustify;

    // Colonne Affaire
    if (RT_Aff > 0) then
    begin
        GClient.Cells[RT_Aff, 0] := TraduireMemoire('Affaire') ;
        GClient.ColAligns[RT_Aff] := taLeftJustify;
    end;

    // Colonne Total
    GClient.Cells[RT_Tot, 0] := TraduireMemoire('Total') ;
    GClient.ColTypes[RT_Tot] := 'R';
    { FQ 20048 BVE 26.04.07
    GClient.ColFormats[RT_Tot] := '#,##0.00'; }
    GClient.ColFormats[RT_Tot] := GetFormatCol;
    { END FQ 20048 }
    GClient.ColAligns[RT_Tot] := taRightJustify;

    // Colonne Niveau max
    GClient.Cells[RT_Niv, 0] := TraduireMemoire('Niv.') ;
    GClient.ColAligns[RT_Niv] := taCenter;

    // Colonne Date dernière relance
    GClient.Cells[RT_Der, 0] := TraduireMemoire('Dern Rel.') ;
    GClient.ColTypes[RT_Der] := 'D';
    GClient.ColFormats[RT_Der] := 'dd/mm/yy';
    GClient.ColAligns[RT_Der] := taLeftJustify;
    //YMO 27/04/2006 FQ17540 Sans cette ligne, le format en colonne Rt_Der(6) se répercute sur Rt_Niv(5)
    GClient.ColFormats[RT_Niv] := '#0';

    // Calcul des largeur de colonne ;
    GClientTailleCol;
end;

procedure TOF_CPRELANCECLIENT.InitGEche;
begin
  // taille
    GEche.RowCount := 2;
    GEche.ColCount := 8;

  // Colonne sélection invisible
    GEche.ColLengths[RE_Sel] := -1;

  // Colonne Numéro
    GEche.Cells[RE_Num, 0] := TraduireMemoire('Num.') ;
    GEche.ColAligns[RE_Num] := taRightJustify;

  // Colonne Mode paiement
    GEche.Cells[RE_Pai, 0] := TraduireMemoire('Paie.') ;
    GEche.ColAligns[RE_Pai] := taLeftJustify;

  // Colonne Date Echéance
    GEche.Cells[RE_Ech, 0] := TraduireMemoire('Echéance') ;
    GEche.ColTypes[RE_Ech] := 'D';
    GEche.ColFormats[RE_Ech] := 'dd/mm/yy';
    GEche.ColAligns[RE_Ech] := taLeftJustify;

  // Colonne Montant
    GEche.Cells[RE_Tot, 0] := TraduireMemoire('Montant') ;
    GEche.ColTypes[RE_Tot] := 'R';

    { FQ 20048 BVE 26.04.07
    GEche.ColFormats[RE_Tot] := '#,#00.00';  }
    GEche.ColFormats[RE_Tot] := GetFormatCol;
    { END FQ 20048 }
    GEche.ColAligns[RE_Tot] := taRightJustify;

  // Colonne Niveau Relance
    GEche.Cells[RE_Niv, 0] := TraduireMemoire('Niv.') ;
    GEche.ColAligns[RE_Niv] := taCenter;

  // Colonne Ecart
    GEche.Cells[RE_Dif, 0] := TraduireMemoire('Dif.') ;
    GEche.ColAligns[RE_Dif] := taCenter;

  // Calcul de la largeur des colonnes
    GEcheTailleCol;
end;

function TOF_CPRELANCECLIENT.WhereCritMP : string;
var
    lStCond : string;
    lStChaine : string;
    lStModePaie : string;
    lStCatPaie : string;
    lInCpt : integer;
begin
    Result := '';

    // ======================
    // === EN MODE TRAITE ===
    // ======================
    if (ModeTraite) then
    begin
        // Condition sur le mode de paiement pour les traites
        lStCond := '';
        for lInCpt := 0 to VH^.MPACC.Count - 1 do
        begin
            lStChaine := VH^.MPACC[lInCpt];
            lStModePaie := ReadtokenSt(lStChaine); // Mode de paiement
            ReadtokenSt(lStChaine); // Acc
            lStCatPaie := ReadTokenSt(lStChaine); // Catégorie de paiment
            if (lStCatPaie = 'LCR') then lStCond := lStCond + ' OR E_MODEPAIE="' + lStModePaie + '"';
        end;

        if (Copy(lStCond, 2, 3) = 'AND') then System.Delete(lStCond, 1, 5);
        if (Copy(lStCond, 2, 2) = 'OR') then System.Delete(lStCond, 1, 4);

        if (lStCond <> '') then result := ' AND (' + lStCond + ')'
        else Result := '';

        // Condition sur le code acceptation
        Result := Result + ' AND (E_CODEACCEPT="TRA" OR E_CODEACCEPT="NAC")';
    end
    // ==================
    // === AUTRE MODE ===
    // ==================
    else
        if (GetCheckBoxState('AvecTraite') <> cbChecked) then Result := ' AND ((E_CODEACCEPT<>"TRA" AND E_CODEACCEPT<>"NAC") OR (E_CODEACCEPT=""))';
end;

procedure TOF_CPRELANCECLIENT.RenseigneXXWhere;
var
    lStCond : string;
    lStV8 : string;
begin
    lStCond := ' E_AUXILIAIRE<>"" AND ' // Auxiliaire obligatoire
        + '( E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H" )' // A-nouveaux
        + WhereCritMP; // E_MODEPAIE + E_CODEACCEPT

    lStV8 := LWhereV8;
    if (lStV8 <> '') then lStCond := lStCond + ' AND ' + lStV8; // Exo V8

    if (ModeTraite) then lStCond := lStCond + ' AND T_RELANCETRAITE<>""' // Type de relance
    else lStCond := lStCond + ' AND T_RELANCEREGLEMENT<>""';

    SetControlText('XX_WHERE', lStCond);
end;

procedure TOF_CPRELANCECLIENT.RenseigneXXOrderBy;
begin
    SetControlText('XX_ORDERBY', GetOrderBy);
end;

procedure TOF_CPRELANCECLIENT.RemplitGClient;
var
    lStChamps : string;
    lInRow : Integer;
    lTobLigne : TOB;
begin
    GClient.rowCount := 2;

    if (GetCheckBoxState('RUPTUREAFFAIRE') = cbChecked) then lStChamps := 'AUXILIAIRE;LIBELLE;AFFAIRE;TOTAL;NIVEAUMAX;DATERELANCE;SELECTION'
    else lStChamps := 'AUXILIAIRE;LIBELLE;TOTAL;NIVEAUMAX;DATERELANCE;SELECTION';

    TOBClient.PutGridDetail(GClient, false, false, lStChamps, true);

    // Flip sélection des lignes de la grille
    FBoFlipAuto := true;
    for lInRow := 1 to TobClient.Detail.Count do
    begin
        lTobLigne := TobClient.Detail[lInRow - 1];
        if (GClient.IsSelected(lInRow) <> (lTobLigne.GetValue('SELECTION') = '*')) then GClient.FlipSelection(lInRow);
    end;

    FBoFlipAuto := false;
    GClient.Refresh;
end;

procedure TOF_CPRELANCECLIENT.RemplitGEche;
var
    lTobCli : TOB;
    lTobEche : TOB;
    lInRow : Integer;
begin
    GEche.rowCount := 2;

    if GClient.row > TobClient.Detail.Count then Exit;

    lTobCli := TobClient.Detail[GClient.row - 1];

    lTobCli.PutGridDetail(GEche, false, false, 'E_NUMEROPIECE;MODEPAIE;DATEECHEANCE;MONTANT;NIVEAURELANCE;ECART;SELECTION', true);

    // Flip-Sélection des lignes de la grille
    FBoFlipAuto := true;
    for lInRow := 1 to lTobCli.Detail.Count do
    begin
        lTobEche := lTobCli.Detail[lInRow - 1];
        if (GEche.IsSelected(lInRow) <> (lTobEche.GetValue('SELECTION') = '*')) then GEche.FlipSelection(lInRow);
    end;
    FBoFlipAuto := false;

    GEche.Refresh;

    // MAJ accès tableau des échéances
    GEche.Enabled := (lTobCli.GetValue('SELECTION') = '*');
    FlapLigne; // MAJ couleur
end;

procedure TOF_CPRELANCECLIENT.RenseigneLaTob;
var
    lTobEche, lTobClient : TOB;
    lStOldAux, lStOldAff, lStNewAux, lStNewAff : string;
    Debit, Credit, Couv, Solde : Double;
begin
    TobClient.ClearDetail;
    lTobClient := nil;
    lStOldAux := '';
    lStOldAff := '';

{$IFDEF EAGLCLIENT}
    // On a besoin de tous les enregistrements !
    InitMove(QEcr.TQ.Detail.Count, '');
{$ELSE}
    InitMove(QEcr.RecordCount, '');
{$ENDIF EAGLCLIENT}

    QEcr.First;

    while (not QEcr.EOF) do
    begin
        MoveCur(false);

        if (EstLigneOk) then
        begin
            lStNewAux := QEcr.FindField('E_AUXILIAIRE').AsString;
            lStNewAff := QEcr.FindField('E_AFFAIRE').AsString;

            if ((lStNewAux <> lStOldAux) or ((lStNewAff <> lStOldAff) and (GetCheckBoxState('RUPTUREAFFAIRE') = cbChecked))) then
            begin
                lTobClient := TOB.Create('VLIGNESCLIENT', TobClient, -1);

                lTobClient.AddChampSupValeur('SELECTION', '*');
                lTobClient.AddChampSupValeur('AUXILIAIRE', lStNewAux);
                lTobClient.AddChampSupValeur('AFFAIRE', lStNewAff);
                lTobClient.AddChampSupValeur('LIBELLE', QEcr.FindField('T_LIBELLE').AsString);
                lTobClient.AddChampSupValeur('TOTAL', 0);
                lTobClient.AddChampSupValeur('NIVEAUMAX', 0);
                lTobClient.AddChampSupValeur('DATERELANCE', iDate1900);
                lTobClient.AddChampSupValeur('SCORERELANCE', QEcr.FindField('T_SCORERELANCE').AsInteger);
                if (ModeTraite) then lTobClient.AddChampSupValeur('RELANCE', QEcr.FindField('T_RELANCETRAITE').AsString)
                else lTobClient.AddChampSupValeur('RELANCE', QEcr.FindField('T_RELANCEREGLEMENT').AsString);

                lStOldAux := lStNewAux;
                lStOldAff := lStNewAff;
            end;

            lTobEche := TOB.Create('VLIGNESECHE', lTobClient, -1);

            lTobEche.AddChampSupValeur('SELECTION', '*');
            lTobEche.AddChampSupValeur('E_NUMEROPIECE', QEcr.FindField('E_NUMEROPIECE').AsInteger);
            lTobEche.AddChampSupValeur('MODEPAIE', QEcr.FindField('E_MODEPAIE').AsString);
            lTobEche.AddChampSupValeur('DATEECHEANCE', QEcr.FindField('E_DATEECHEANCE').AsDateTime);
            lTobEche.AddChampSupValeur('NIVEAURELANCE', QEcr.FindField('E_NIVEAURELANCE').AsInteger + 1);
            lTobEche.AddChampSupValeur('DATERELANCE', QEcr.FindField('E_DATERELANCE').AsDateTime);
            {JP 25/10/05 : Dans le cadre de la FQ 16499, on suppime le champ Formule C24 qui faisait
                           simplement appel à E_QUALIFORIGINE !!!}
            lTobEche.AddChampSupValeur('QUALIFORIGINE', QEcr.FindField('E_QUALIFORIGINE').AsString);

            // Calcul du solde
            Debit := QEcr.FindField('E_DEBITDEV').AsFloat;
            Credit := QEcr.FindField('E_CREDITDEV').AsFloat;
            Couv := QEcr.FindField('E_COUVERTUREDEV').AsFloat;

            if (Debit <> 0) then Solde := Debit - Couv
            else Solde := Couv - Credit;

            lTobEche.AddChampSupValeur('DEBITDEV', Debit);
            lTobEche.AddChampSupValeur('CREDITDEV', Credit);
            lTobEche.AddChampSupValeur('COUVERTUREDEV', Couv);
            lTobEche.AddChampSupValeur('MONTANT', Solde);

            // Ecart
            lTobEche.AddChampSupValeur('ECART', GetEcartQ);

            // CLE ECRITURE
            lTobEche.AddChampSupValeur('E_JOURNAL', QEcr.FindField('E_JOURNAL').AsString);
            lTobEche.AddChampSupValeur('E_EXERCICE', QEcr.FindField('E_EXERCICE').AsString);
            lTobEche.AddChampSupValeur('E_DATECOMPTABLE', QEcr.FindField('E_DATECOMPTABLE').AsDateTime);
            lTobEche.AddChampSupValeur('E_QUALIFPIECE', QEcr.FindField('E_QUALIFPIECE').AsString);
            lTobEche.AddChampSupValeur('E_NATUREPIECE', QEcr.FindField('E_NATUREPIECE').AsString);
            lTobEche.AddChampSupValeur('E_NUMLIGNE', QEcr.FindField('E_NUMLIGNE').AsInteger);
            lTobEche.AddChampSupValeur('E_NUMECHE', QEcr.FindField('E_NUMECHE').AsInteger);
            lTobEche.AddChampSupValeur('E_DATEMODIF', QEcr.FindField('E_DATEMODIF').AsDateTime);
            // autres champs utilisés
            lTobEche.AddChampSupValeur('E_ETABLISSEMENT', QEcr.FindField('E_ETABLISSEMENT').AsString);
            lTobEche.AddChampSupValeur('E_DEVISE', QEcr.FindField('E_DEVISE').AsString);
            lTobEche.AddChampSupValeur('E_DATETAUXDEV', QEcr.FindField('E_DATETAUXDEV').AsDateTime);
            lTobEche.AddChampSupValeur('E_TAUXDEV', QEcr.FindField('E_TAUXDEV').AsFloat);
            lTobEche.AddChampSupValeur('E_VALIDE', QEcr.FindField('E_VALIDE').AsString);
            lTobEche.AddChampSupValeur('E_MODESAISIE', QEcr.FindField('E_MODESAISIE').AsString);
            lTobEche.AddChampSupValeur('MEMOCHARGE', False);
            {JP 21/07/06 : Modification de la fonction WhereEcritureTob : Le champ préfixe est maintenant inutile
            lTobEche.AddChampSupValeur('PREFIXE', 'E');}
            lTobEche.AddChampSupValeur('E_BLOCNOTE', QEcr.FindField('E_BLOCNOTE').AsString);
        end;

        QEcr.Next;
    end;

    FiniMove;
end;

function TOF_CPRELANCECLIENT.EstLigneOk : boolean;
var
    Famille, NatP : string;
    Ecart, Niveau : Integer;
    lTobRel : TOB;
    Solde, SoldeC, CouvP : double;
begin
    Result := false;

    // Test du niveau
    Niveau := QEcr.FindField('E_NIVEAURELANCE').AsInteger;
    if ((FInNivoRef >= 0) and (Niveau <> FInNivoRef - 1)) then Exit;

    // Test du retard
    Ecart := GetEcartQ;
    if (Ecart < 0) then Exit;

    // Test du retard / délai famille relance
    if (modeTraite) then Famille := QEcr.FindField('T_RELANCETRAITE').AsString
    else Famille := QEcr.FindField('T_RELANCEREGLEMENT').AsString;

    lTobRel := GetRelance(Famille);

    if (lTobRel = nil) then exit;
    if (Ecart < lTobRel.GetValue('RR_DELAI1')) then exit;

    // Option Test Jour de paiement // FQ 15315 SBO 25/01/2004 On prend en compte les délais de relance si demandé dans les paramètres de relances
    if lTobRel.GetValue('RR_ENJOURS')='X' then
      if Ecart <= lTobRel.GetValue('RR_DELAI'+ IntToStr(Niveau + 1) ) then
        Exit ; // Fin FQ 15315

    // Test de la couverture
    NatP := QEcr.FindField('E_NATUREPIECE').AsString;
    SoldeC := QEcr.FindField('E_DEBITDEV').AsFloat + QEcr.FindField('E_CREDITDEV').AsFloat - QEcr.FindField('E_COUVERTUREDEV').AsFloat;
    if ((SoldeC <= 0) and (GetControlText('SENS') <> 'M')) then exit;

    // Test des montants
    Solde := QEcr.FindField('E_DEBITDEV').AsFloat - QEcr.FindField('E_CREDITDEV').AsFloat;
    if ((GetControlText('SENS') = 'D') and (Solde <= 0)) then exit;
    if ((GetControlText('SENS') = 'C') and (Solde >= 0)) then exit;
    Solde := QEcr.FindField('E_DEBIT').AsFloat - QEcr.FindField('E_CREDIT').AsFloat;
    if (Abs(Solde) < Valeur(GetControlText('MONTANTMIN'))) or (Abs(Solde) > Valeur(GetControlText('MONTANTMAX'))) then exit;


    CouvP := QEcr.FindField('E_DEBIT').AsFloat + QEcr.FindField('E_CREDIT').AsFloat - QEcr.FindField('E_COUVERTURE').AsFloat;
    {JP 25/01/06 : FQ 17027 : on exclut tous les soldes à 0, quelle que soit la nature de pièce
    if (((NatP = 'FC') or (NatP = 'AC')) and (Arrondi(CouvP, V_PGI.OkDecV) = 0)) then
       if (GetControlText('SENS') <> 'M') then Exit;}
    if Arrondi(CouvP, V_PGI.OkDecV) = 0 then Exit;
    // TOUT EST OK
    Result := true;
end;

function TOF_CPRELANCECLIENT.GetEcartQ : Integer;
var
    Eche, DateC : TDateTime;
    lDelta : Double;
begin
    Eche := QEcr.FindField('E_DATEECHEANCE').AsDateTime;
    DateC := QEcr.FindField('E_DATECOMPTABLE').AsDateTime;

    if (modeTraite) then
        if (GetControlText('ChoixDate') = 'AVANT') then lDelta := FDtDateRel - (Eche - (ValeurI(GetControlText('SpinDelta'))))
        else lDelta := FDtDateRel - (DateC + (ValeurI(GetControlText('SpinDelta'))))
    else lDelta := FDtDateRel - Eche;

    Result := Trunc(lDelta);
end;

function TOF_CPRELANCECLIENT.GetRelance(vStFamille : string) : TOB;
var
    lStType : string;
begin
    Result := nil;

    if (vStFamille = '') then exit;

    if (ModeTraite) then lStType := 'RTR'
    else lStType := 'RRG';

    Result := TobRelance.FindFirst(['RR_TYPERELANCE', 'RR_FAMILLERELANCE'], [lStType, vStFamille], false);

    if (Result = nil) then
    begin
        Result := TOB.Create('RELANCE', TobRelance, -1);
        Result.putValue('RR_TYPERELANCE', lStType);
        Result.putValue('RR_FAMILLERELANCE', vStFamille);
        if (not Result.LoadDB) then FreeAndNil(Result);
    end;
end;

procedure TOF_CPRELANCECLIENT.GClientToutSelect(vBoTout : Boolean);
var
    lInRow : Integer;
    lTobCli : TOB;
begin
    // MAJ champ selection dans les TOB
    for lInRow := 1 to TobCLient.Detail.Count do
    begin
        lTobCli := TobClient.Detail[lInRow - 1];

        // MAJ de la ligne de la grille "client"
        if vBoTout then lTobCli.PutValue('SELECTION', '*')
        else lTobCli.PutValue('SELECTION', ' ');

        GEcheToutSelect(lInRow, vBoTout, false);
    end;

    // MAJ affichage
    RemplitGClient;
    RemplitGEche;

    // MAJ indicateur "tout sélectionner"
    BSelectAll.Down := vBoTout;
    FBoToutSel := vBoTout;
end;

procedure TOF_CPRELANCECLIENT.GClientRowEnter(Sender : TObject;Ou : Integer;var Cancel : Boolean;Chg : Boolean);
begin
    RemplitGEche;
end;

procedure TOF_CPRELANCECLIENT.BSelectAllClick(Sender : TObject);
begin
    if ((not FBoFlipAuto) and (not ModeManuel)) then exit;
    GClientToutSelect(not FBoToutSel);
end;

procedure TOF_CPRELANCECLIENT.GClientFlipSelection(Sender : TObject);
var
    lTobLigne : TOB;
    lBoSelect : Boolean;
begin
    if (FBoFlipAuto) then Exit;

    if (GClient.row > TobClient.Detail.Count) then Exit;

    lBoSelect := GClient.IsSelected(GClient.row);

    lTobLigne := TobClient.Detail[GClient.row - 1];

    // MAJ indicateur de sélection
    if (lBoSelect) then lTobLigne.PutValue('SELECTION', '*')
    else lTobLigne.PutValue('SELECTION', ' ');

    // MAJ sélecion échéances
    GEcheToutSelect(GClient.row, lBoSelect, true);

    // MAJ affichage ligne client
    GClientMAJLigne(GClient.row);

    // MAJ indicateur
    BSelectAll.Down := false;
    FBoToutSel := false;
end;

procedure TOF_CPRELANCECLIENT.GEcheFlipSelection(Sender : TObject);
var
    lTobLigne : TOB;
    lTobEche : TOB;
begin
    if (FBoFlipAuto) then Exit;

    if (GClient.row > TobClient.Detail.Count) then Exit;
    lTobLigne := TobClient.Detail[GClient.row - 1];

    if (GEche.row > lTobLigne.Detail.Count) then Exit;
    lTobEche := lTobLigne.Detail[GEche.row - 1];

    // MAJ indicateur de sélection
    if (GEche.IsSelected(GEche.row)) then lTobEche.PutValue('SELECTION', '*')
    else lTobEche.PutValue('SELECTION', ' ');

    // Maj total ligne client
    GClientRecalculLigne(GClient.row, true);

    // MAJ affichage lignes des grilles
    GClientMAJLigne(GClient.row);
    GEcheMAJLigne(GEche.row);

    // MAJ indicateur
    BSelectAll.Down := false;
    FBoToutSel := false;
end;

procedure TOF_CPRELANCECLIENT.GEcheToutSelect(vInRow : Integer;vBoTout : Boolean;vBoMajGrille : Boolean);
var
    lTobEche, lTobCli : TOB;
    lInEche : Integer;
begin
    if (vInRow > TobClient.Detail.Count) then Exit;
    lTobCli := TobClient.Detail[vInRow - 1];

    // Sélection / désélection des lignes d'échances
    for lInEche := 0 to lTobCli.Detail.count - 1 do
    begin
        lTobEche := lTobCli.Detail[lInEche];

        if vBoTout then lTobEche.PutValue('SELECTION', '*')
        else lTobEche.PutValue('SELECTION', ' ');
    end;

    // MAJ des cumuls lignes clients
    GClientRecalculLigne(vInRow, false);

    // Maj grille si demandé
    if (vBoMajGrille) then RemplitGEche;
end;

procedure TOF_CPRELANCECLIENT.GClientMajLigne(vInRow : Integer);
var
    lTobLigne : Tob;
    lStChamps : string;
begin
    if (vInRow > TobClient.Detail.Count) then Exit;
    lTobLigne := TobClient.Detail[vInRow - 1];

    if (GetCheckBoxState('RUPTUREAFFAIRE') = cbChecked) then lStChamps := 'AUXILIAIRE;LIBELLE;AFFAIRE;TOTAL;NIVEAUMAX;DATERELANCE;SELECTION'
    else lStChamps := 'AUXILIAIRE;LIBELLE;TOTAL;NIVEAUMAX;DATERELANCE;SELECTION';

    lTobLigne.PutLigneGrid(GClient, vInRow, false, false, lStChamps);

    GClient.Refresh;
end;

procedure TOF_CPRELANCECLIENT.GEcheMajLigne(vInRow : Integer);
var
    lTobLigne : Tob;
    lTobEche : Tob;
begin
    if (GClient.Row > TobClient.Detail.Count) then Exit;
    lTobLigne := TobClient.Detail[GClient.Row - 1];

    if (vInRow > lTobLigne.Detail.Count) then Exit;
    lTobEche := lTobLigne.Detail[vInRow - 1];

    lTobEche.PutLigneGrid(GEche, vInRow, false, false, 'E_NUMEROPIECE;MODEPAIE;DATEECHEANCE;MONTANT;NIVEAURELANCE;ECART;SELECTION');
end;

procedure TOF_CPRELANCECLIENT.GClientRecalculLigne(vInRow : Integer;vBoMajGrille : Boolean);
var
    lNiveau, lInEche : Integer;
    lTobEche, lTobCli : TOB;
    lDtRelance : TDateTime;
    lDebit, lTotal, lCredit, lCouv : Double;
begin
    if (vInRow > TobClient.Detail.Count) then Exit;
    lTobCli := TobClient.Detail[vInRow - 1];

    lTotal := 0;
    lNiveau := 0;
    lDtRelance := iDate1900;

    // Parcours des lignes d'échances
    for lInEche := 0 to lTobCli.Detail.count - 1 do
    begin
        lTobEche := lTobCli.Detail[lInEche];

        if (lTobEche.GetValue('SELECTION') = '*') then
        begin
            // maj total
            lDebit := lTobEche.GetValue('DEBITDEV');
            lCredit := lTobEche.GetValue('CREDITDEV');
            lCouv := lTobEche.GetValue('COUVERTUREDEV');

            if (lDebit <> 0) then lTotal := lTotal + lDebit - lCouv
            else lTotal := lTotal - (lCredit - lCouv);

            // maj niveau
            if (lTobEche.GetValue('NIVEAURELANCE') > lNiveau) then lNiveau := lTobEche.GetValue('NIVEAURELANCE');
            // maj date relance
            if (lTobEche.GetValue('DATERELANCE') > lDtRelance) then lDtRelance := lTobEche.GetValue('DATERELANCE');
        end;
    end;

    // MAJ de la ligne de la grille "client"
    lTobCli.PutValue('TOTAL', lTotal);
    lTobCli.PutValue('NIVEAUMAX', lNiveau);
    lTobCli.PutValue('DATERELANCE', lDtRelance);

    if (vBoMajGrille) then GClientMajLigne(vInRow);
end;

procedure TOF_CPRELANCECLIENT.GClientDoubleClick(Sender : TObject);
begin
    if (not ModeManuel) then Exit;
    GClient.FlipSelection(GClient.row);
end;

procedure TOF_CPRELANCECLIENT.GEcheDoubleClick(Sender : TObject);
begin
    if (not ModeManuel) then Exit;
    GEche.FlipSelection(GEche.row);
end;

procedure TOF_CPRELANCECLIENT.ValideRelance;
var
    i, lInMaxNiv, lInNivScore : Integer;
    lStFamille, lStAuxi, lStAffaire : string;
    lBoTraitementOk, lBoCarePDF : Boolean;
    lTobCli : TOB;
    ListeModele : HTStringList;
begin
    lBoCarePDF := false;
    FNbRelance := 0;
    lBoTraitementOk := true;

    // On vide le XX_ORDERBY afin de ne pas récupérer de clause ORDER BY dans le RecupWhereCritere
    SetControlText('XX_ORDERBY', '');

    InitMove(TobClient.Detail.Count, '');
    RecupMSED(false);

    if (FMSED.Spooler and (not V_PGI.QRPDF)) then
    begin
        V_PGI.QRPDF := true;
        lBoCarePDF := true;
    end;

    ListeModele := HTStringList.Create;
    ListeModele.Sorted:=true; {Lek 27/09/05 FQ16464}
    ListeModele.Duplicates := DupIgnore;

    for i := 0 to TobClient.Detail.count - 1 do
    begin
        MoveCur(false);
        FBoAppliqScore := false;

        lTobCli := TobClient.Detail[i];
        if ((lTobCli.GetValue('SELECTION') <> '*') or (lTobCli.GetValue('TOTAL') <= 0) or (lTobCli.Detail.Count <= 0)) then Continue;

        lStFamille := lTobCli.GetValue('RELANCE');
        if (lStFamille = '') then Continue;

        lInNivScore := lTobCli.GetValue('SCORERELANCE');
        lStAffaire := lTobCli.GetValue('AFFAIRE');
        lInMaxNiv := lTobCli.GetValue('NIVEAUMAX');
        lStAuxi := lTobCli.GetValue('AUXILIAIRE');

        lBoTraitementOk := EclateEches(lTobCli, lStFamille, lStAuxi, lStAffaire, lInNivScore, lInMaxNiv, ListeModele);

        if (V_PGI.IoError <> oeOk) then Break;
    end;

    FiniMove;

    if (lBoTraitementOk) then
    begin
        EclateEchesParModele(ListeModele);
        VideLaStringList(ListeModele);
        ListeModele.Free;
    end;

    if ((GetCheckBoxState('Spooler') = cbChecked) or (not (GetCheckBoxState('APERCU') = cbChecked))) then PGIInfo('Traitement terminé.' + #13 + #10 + ' ' + IntToStr(FNbRelance) + ' Relance(s) ont été envoyées à l''impression.', Ecran.Caption);

    if FMSED.Spooler then SetParamSoc('SO_CPCHEMINRELANCE', FMSED.RepSpooler);

    if lBoCarePDF then V_PGI.QRPDF := false;
end;

function TOF_CPRELANCECLIENT.EclateEches(vTobCli : TOB;vStFamille, vStAuxi, vStAffaire : string;vInNivScore, vInMaxNiv : Integer;ListeModele : HTStringList) : Boolean;
var
    lTobRelance : TOB;
    lStModele : string;
    lInNivo : Integer;
begin
    Result := true;

    lTobRelance := GetRelance(vStFamille);
    if (lTobRelance = nil) then exit;

    lStModele := '';

    // ====================================
    // Paramètre de la famille groupe = OUI
    // ====================================
    if (lTobRelance.GetValue('RR_GROUPELETTRE') = 'X') then
    begin
        if (vInMaxNiv <= 0) then vInMaxNiv := 1;
        if (lTobRelance.GetValue('RR_SCOORING') = 'X') then // FQ 11732
        begin
            if (vInNivScore < vInMaxNiv) then vInNivScore := vInMaxNiv;
        end
        else vInNivScore := vInMaxNiv;

        if (not ExisteModele(lTobRelance.GetValue('RR_MODELE' + IntToStr(vInNivScore)), lStModele)) then Exit;

        if (not MarqueEches(-1, vTobCli, lStModele)) then Exit;

        MajListeTiersparModele(vStFamille, vStAuxi, vStAffaire, lStModele, vInNivScore, vInMaxNiv, ListeModele);
    end
    // =====================
    // Une lettre par niveau
    // =====================
    else
    begin
        for lInNivo := 1 to vInMaxNiv do
        begin
            // Application du scoring si Score = OUI et T_SCORERELANCE>Niveau
            if ((lTobRelance.GetValue('RR_SCOORING') = 'X') and (vInNivScore > lInNivo) and (vInNivScore <= RELNIVOMAX)) then
                if (ExisteModele(lTobRelance.GetValue('RR_MODELE' + IntToStr(vInNivScore)), lStModele)) then FBoAppliqScore := true
                else Continue
            else
                if (not ExisteModele(lTobRelance.GetValue('RR_MODELE' + IntToStr(lInNivo)), lStModele)) then Continue;

            if (MarqueEches(lInNivo - 1, vTobCli, lStModele)) then
            begin
                MajListeTiersParModele(vStFamille, vStAuxi, vStAffaire, lStModele, vInNivScore, vInMaxNiv, ListeModele);
                V_PGI.IoError := oeOk;
            end;

            if (V_PGI.IoError <> oeOk) then
            begin
                Result := false;
                Break;
            end;
        end;
    end;
end;

function TOF_CPRELANCECLIENT.ExisteModele(vStModele : string;var vCodeM : string) : boolean;
begin
    Result := false;

    if (vStModele <> '') then
    begin
        vCodeM := vStModele;
        Result := true;
        if (GetCheckBoxState('UnModele') = cbChecked) then vCodeM := GetControlText('ModeleUnique');
    end;
end;

procedure TOF_CPRELANCECLIENT.EclateEchesParModele(ListeModele : HTStringList);
var
    lTobRel : TOB;
    lStModele, lStSQL, lStSousReq, lStAuxi, lStFamille, lStChaine, lStAffaire : string;
    lInNivScore, lInMaxNiv, i, j, k : Integer;
    TT : Tlist;
    LL : TStrings;
    LTiersModele : HTStringList;
    lBoNivoMax : Boolean;
begin
    FBoGenePremier := true;
    if (ListeModele = nil) then Exit;

    for i := 0 to ListeModele.Count - 1 do
    begin
        lStModele := ListeModele[i];
        LTiersModele := HTStringList(ListeModele.Objects[i]);

        if (LTiersModele <> nil) then
        begin
            TT := TList.Create;

            for j := 0 to LTiersModele.Count - 1 do
            begin
                lStChaine := LTiersModele[j];
                lStAuxi := '';
                lStFamille := '';
                lStAffaire := '';

                if (lStChaine <> '') then lStAuxi := ReadTokenSt(lStChaine);
                if (lStChaine <> '') then lStFamille := ReadTokenSt(lStChaine);
                if (lStChaine <> '') then lStAffaire := ReadTokenSt(lStChaine);

                lTobRel := GetRelance(lStFamille);

                if ((lTobRel <> nil) and (lStAuxi <> '')) then
                begin
                    lStSousReq := GetConditionSQLEcr(-1, (lTobRel.GetValue('RR_NONECHU') = 'X'), lStModele);
                    LL := TStringList.Create;

                    if (GetCheckBoxState('RuptureAffaire') = cbChecked) then lStSQL := RecupSQLModele('L', 'REL', lStModele, '', lStAuxi, '', ' WHERE E_AUXILIAIRE="' + lStAuxi + '" AND E_AFFAIRE="' + lStAffaire + '" ')
                    else lStSQL := RecupSQLModele('L', 'REL', lStModele, '', lStAuxi, '', ' WHERE E_AUXILIAIRE="' + lStAuxi + '" ');

                    LL.Add(lStSQL);
                    LL.Add(lStSousReq);
                    TT.Add(LL);
                    Inc(FNbRelance);
                end;
            end;

        {$IFNDEF EAGLCLIENT}
            if (FMSED.Spooler) then LanceDocumentSpool('L', 'REL', lStModele, TT, nil, (GetCheckBoxState('APERCU') = cbChecked), FBoGenePremier, false, FMSED.RepSpooler, FMSED.RacineSpooler, FMSED.XFichierSpooler, FMSED.SoucheSpooler)
            else begin
        {$ENDIF EAGLCLIENT}
                LanceDocument('L', 'REL', lStModele, TT, nil, (GetCheckBoxState('APERCU') = cbChecked), FBoGenePremier);
        {JP 02/08/05 : FQ 15627 : Comme on lance plusieurs fois le LanceEtat, mais qu'au-delà de première fois on
                       n'affiche plus la boîte de dialoque(cf ci-dessous), je récupère le nombre de copies demandées
                       sur la première impression et je l'affecte au nombre de copies par défaut}
                if FBoGenePremier then
                  V_PGI.DefaultDocCopies := V_PGI.NbDocCopies;
        {$IFNDEF EAGLCLIENT}
            end;
        {$ENDIF EAGLCLIENT}

            for k := 0 to TT.Count - 1 do
            begin
                LL := TStringList(TT[k]);
                LL.Clear;
                LL.Free;
            end;

            TT.Free;
            FBoGenePremier := false;
        end;

        for j := 0 to LTiersModele.Count - 1 do
        begin
            lStChaine := LTiersModele[j];
            lStAuxi := '';
            lStFamille := '';
            lStAffaire := '';
            lInNivScore := -10;
            lInMaxNiv := -10;

            if (lStChaine <> '') then lStAuxi := ReadTokenSt(lStChaine);
            if (lStChaine <> '') then lStFamille := ReadTokenSt(lStChaine);
            if (lStChaine <> '') then lStAffaire := ReadTokenSt(lStChaine);
            if (lStChaine <> '') then lInNivScore := StrToInt(ReadtokenSt(lStChaine));
            if (lStChaine <> '') then lInMaxNiv := StrToInt(ReadtokenSt(lStChaine));

            if ((lInNivScore <> -10) and (lInMaxNiv <> -10)) then
            begin
                if ((FBoAppliqScore) and (lInNivScore > 1)) then MajEcrScore(lStAuxi, lStModele, lInNivScore, lInMaxNiv)
                else
                begin
                    lBoNivoMax := false;
                    lTobRel := GetRelance(lStFamille);

                    if (lTobRel <> nil) then
                        if ((lTobRel.GetValue('RR_GROUPELETTRE') = 'X') and (lTobRel.GetValue('RR_MODELE' + IntToStr(lInMaxNiv)) <> '')) then lBoNivoMax := true;

                    if (lBoNivoMax) then MajEcrScore(lStAuxi, lStModele, lInMaxNiv, lInMaxNiv)
                    else MajEcritures(lStAuxi, lStModele, lInMaxNiv);
                end;
            end;
        end;
    end;
end;

function TOF_CPRELANCECLIENT.MarqueEches(vInNivo : Integer;vTobCli : TOB;vStModele : string) : boolean;
var
    lTobEche : TOB;
    lInEche : Integer;
    lStSQL, lStEMP : string;
begin
    Result := False;
    lStEMP := '';

    for lInEche := 0 to vTobCli.Detail.Count - 1 do
    begin
        lTobEche := vTobCli.Detail[lInEche];

        if (lTobEche.GetValue('SELECTION') <> '*') then Continue;
        if ((vInNivo >= 0) and (lTobEche.GetValue('NIVEAURELANCE') <> (vInNivo + 1))) then Continue;

        lStEMP := '';

        if ((GetControlText('MODELETP') <> '') and (lTobEche.GetValue('QUALIFORIGINE') = 'TP') and (VH^.OuiTP)) then
        begin
            lStEMP := 'TP';
            FBoYATP := true;
        end;

          // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
          //lStSQL := 'UPDATE ECRITURE SET E_NUMPIECEINTERNE="EDITREL", E_MULTIPAIEMENT="' + lStEMP +
          lStSQL := 'UPDATE ECRITURE SET E_NUMPIECEINTERNE="EDITREL;'+vStModele+'", E_MULTIPAIEMENT="' + lStEMP +
          //'", E_DATEMODIF="' + UsTime(FDtNowFutur) + '", E_REFPOINTAGE="' + vStModele +
          '", E_DATEMODIF="' + UsTime(FDtNowFutur) +
          {JP 21/07/06 : Modification de la fonction WhereEcritureTob}
          '" WHERE ' + WhereEcritureTob(tsGene, lTobEche, true, False, 'E') + ' AND E_DATEMODIF="' +
          UsTime(lTobEche.GetValue('E_DATEMODIF')) + '"';

        if (ExecuteSQL(lStSQL) <> 1) then
        begin
            V_PGI.IoError := oeSaisie;
            Exit;
        end;

        Result := true;
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lek
Créé le ...... : 27/09/2005
Modifié le ... :   /  /    
Description .. : La liste de tiers par modèle a été trier par le code du tiers.
Suite ........ : A la demande FQ16464, il faut que la liste du tiers édite par 
Suite ........ : l'ordre de sélection dans MUL.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRELANCECLIENT.MajListeTiersparModele(vStFamille, vStAuxi, vStAffaire, vStModele : string;vInNivScore, vInMaxNiv : Integer;ListeModele : HTStringList);
var
    lInIndice : Integer;
    LTiersModele : HTStringList;
    lStChaine : string;
    i         : integer; {Lek 27/09/05 FQ16464}
    trouve    : boolean;
begin
    // Recupérztion modèle
    lInIndice := ListeModele.IndexOf(vStModele);
    if lInIndice < 0 then
    begin
        LTiersModele := HTStringList.Create;
        LTiersModele.Sorted := False;{True; Lek 27/09/05 FQ16464}
        LTiersModele.Duplicates := DupIgnore;
    end
    else LTiersModele := HTStringList(ListeModele.Objects[lInIndice]);

    // Consitution de la chaîne
    lStChaine := vStAuxi + ';' + vStFamille + ';';

    if (GetCheckBoxState('RuptureAffaire') = cbChecked) then lStChaine := lStChaine + vStAffaire + ';'
    else lStChaine := lStChaine + ';';

    lStChaine := lStChaine + IntToStr(vInNivScore) + ';' + IntToStr(vInMaxNiv) + ';';

      // Maj de la liste des modèles
    trouve:=false; {Lek 27/09/05 FQ16464}
    for i:=0 to LTiersModele.count-1 do
      if lStChaine=LTiersModele[i] then trouve:=true;
    if not trouve then
      LTiersModele.Add(lStChaine);

    if lInIndice < 0 then ListeModele.AddObject(vStModele, LTiersModele);
end;

procedure TOF_CPRELANCECLIENT.RecupMSED(lBoAvecTP : Boolean);
var
    lStRep : string;
begin
    // Création du MSED
    FillChar(FMSED, SizeOf(FMSED), #0);

    FMSED.Spooler := GetCheckBoxState('Spooler') = cbChecked;
    FMSED.SoucheSpooler := 0;
    fMSED.XFichierSpooler := GetCheckBoxState('XFichierSpooler') = cbChecked;

    if (lBoAvecTP) then
    begin
        FMSED.RepSpooler := GetControlText('RepSpooler') + '\TP';
        FMSED.RacineSpooler := 'RELTP';
    end
    else
    begin
        FMSED.RepSpooler := GetControlText('RepSpooler');
        FMSED.RacineSpooler := GetParamSocSecur('SO_CPRACINERELANCE','');
    end;

    // Création du répertoire
    lStRep := FMSED.RepSpooler;
    if (lStRep <> '') then
    begin
        if (lStRep[Length(lStRep)] = '\') then lStRep := Copy(lStRep, 1, Length(lStRep) - 1);
        if (not DirectoryExists(lStRep)) then CreateDir(lStRep);
    end;
end;

function TOF_CPRELANCECLIENT.GetConditionSQLEcr(vInNivo : Integer;vBoNonEchu : Boolean;vStModele : string) : string;
var
    lStCond1 : string;
    lStCond2 : string;
begin
    // Critère MUL + mode de paiement
    Result := ' AND ' + WhereGeneCritEcr(Pages, QEcr, true) + ' ' + WhereCritMP;

    // Echéance
    if vBoNonEchu then
    begin
        // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
        //lStCond1 := 'E_NUMPIECEINTERNE="EDITREL"';
        lStCond1 := 'E_NUMPIECEINTERNE LIKE "EDITREL%"';
        if vInNivo >= 0 then lStCond1 := lStCond1 + ' AND E_NIVEAURELANCE=' + IntToStr(vInNivo);
        { FQ 21662 BVE 19.10.07
        lStCond2 := '( (E_DEBITDEV+E_CREDITDEV-E_COUVERTUREDEV>0) AND E_NIVEAURELANCE<=0 AND E_DATEECHEANCE>"' + UsDateTime(FDtDateRel) + '")';}
        lStCond2 := '( (E_DEBITDEV+E_CREDITDEV-E_COUVERTUREDEV>0) AND E_NIVEAURELANCE<= ' + GClient.Cells[RT_Niv,GClient.Row] + ')';
        { END 21662 }
        Result := Result + ' AND ( (' + lStCond1 + ') OR ' + lStCond2 + ')';
    end
    else
    begin
        // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
        // Result := Result + ' AND E_NUMPIECEINTERNE="EDITREL"';
        Result := Result + ' AND E_NUMPIECEINTERNE LIKE "EDITREL%"';        
        if vInNivo >= 0 then Result := Result + ' AND E_NIVEAURELANCE=' + IntToStr(vInNivo);
    end;

    // Critère modèle
    if vStModele <> '' then
        if vBoNonEchu then
        begin
            // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
            // lStCond1 := ' E_REFPOINTAGE="' + vStModele + '" ';
            lStCond1 := ' E_NUMPIECEINTERNE="EDITREL;' + vStModele + '" ';
            lStCond1 := '(' + lStCond1 + ')';
            { FQ 21662 BVE 19.10.07
            lStCond2 := '(E_DEBITDEV+E_CREDITDEV-E_COUVERTUREDEV>0 AND E_NIVEAURELANCE<=0 AND E_DATEECHEANCE>"' + UsDateTime(FDtDateRel) + '")';}
            lStCond2 := '(E_DEBITDEV+E_CREDITDEV-E_COUVERTUREDEV>0 AND E_NIVEAURELANCE<=' + GClient.Cells[RT_Niv, GClient.Row] + ')';
            { END FQ 21662 }
            Result := Result + ' AND (' + lStCond1 + ' OR ' + lStCond2 + ')';
        end
        // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
        //else Result := Result + ' AND E_REFPOINTAGE="' + vStModele + '" ';
        else Result := Result + ' AND E_NUMPIECEINTERNE="EDITREL;' + vStModele + '" ';

    // TRI
    // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
    //Result := Result + ' ORDER BY E_REFPOINTAGE, ' + GetOrderBy;
    Result := Result + ' ORDER BY E_NUMPIECEINTERNE, ' + GetOrderBy;

end;

function TOF_CPRELANCECLIENT.GetOrderBy : string;
begin
    if (GetControlText('ORDRERELANCE') = 'CPT') then Result := 'E_AUXILIAIRE, E_AFFAIRE, E_NIVEAURELANCE DESC, E_DATEECHEANCE'
    else
        if (GetControlText('ORDRERELANCE') = 'INT') then Result := 'T_LIBELLE, E_AUXILIAIRE, E_AFFAIRE, E_NIVEAURELANCE DESC, E_DATEECHEANCE'
    else
        if (GetControlText('ORDRERELANCE') = 'NIV') then Result := 'E_NIVEAURELANCE DESC, E_DATEECHEANCE'
    else
        if (GetControlText('ORDRERELANCE') = 'PCP') then Result := 'T_PAYS, T_CODEPOSTAL, E_AUXILIAIRE, E_AFFAIRE, E_NIVEAURELANCE DESC'
    else
        if (GetControlText('ORDRERELANCE') = 'ECH') then Result := 'E_DATEECHEANCE, E_AUXILIAIRE, E_AFFAIRE, E_NIVEAURELANCE DESC'
    else Result := 'E_AUXILIAIRE, E_AFFAIRE, E_DATEECHEANCE, E_NIVEAURELANCE DESC';
end;

procedure TOF_CPRELANCECLIENT.VideLaStringList(L : HTStrings);
var
    i : integer;
begin
    if (L = nil) then Exit;
    if (L.Count <= 0) then Exit;

    for i := 0 to L.Count - 1 do
    begin
        if (L.Objects[i] <> nil) then
        begin
            HTStringList(L.Objects[i]).Clear;
            HTStringList(L.Objects[i]).Free;
        end;

        L.Objects[i] := nil;
    end;
    L.Clear;
end;

procedure TOF_CPRELANCECLIENT.MajEcritures(vStAuxi, vStModele : string;vInMaxNiv : integer);
var
    lStReq : string;
begin
    lStReq := 'UPDATE ECRITURE SET E_DATERELANCE="' + UsDateTime(FDtDateRel) +
    // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
    //           '", E_NIVEAURELANCE=E_NIVEAURELANCE+1, E_NUMPIECEINTERNE="", E_REFPOINTAGE="" ' +
               '", E_NIVEAURELANCE=E_NIVEAURELANCE+1, E_NUMPIECEINTERNE="" ' +
               SQLRetrouve(vStAuxi, vStModele) + ' AND E_NIVEAURELANCE<' + IntToStr(RELNIVOMAX);

    ExecuteSQL(lStReq);

    if (vInMaxNiv = (RELNIVOMAX + 1)) then MajNiveau7(vStAuxi, vStModele);
end;

procedure TOF_CPRELANCECLIENT.MajEcrScore(vStAuxi, vStModele : string;vInNivScore, vInMaxNiv : integer);
var
    lStReq : string;
begin
    // Gestion borne nivo max
    if (vInNivScore > RELNIVOMAX) then vInNivScore := RELNIVOMAX;

    // 1er update
    lStReq := 'UPDATE ECRITURE SET E_DATERELANCE="' + UsDateTime(FDtDateRel) + '", E_NIVEAURELANCE=' +
              // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
              //IntToStr(vInNivScore) + ', E_NUMPIECEINTERNE="", E_REFPOINTAGE="" ' +
              IntToStr(vInNivScore) + ', E_NUMPIECEINTERNE="" ' +
              SQLRetrouve(vStAuxi, vStModele) + ' AND E_NIVEAURELANCE<' + IntToStr(vInNivScore);
    ExecuteSQL(lStReq);

    // 2ème update
    lStReq := 'UPDATE ECRITURE SET E_DATERELANCE="' + UsDateTime(FDtDateRel) +
               // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
               //'", E_NIVEAURELANCE=E_NIVEAURELANCE+1, E_NUMPIECEINTERNE="" ' + ', E_REFPOINTAGE="" ' +
               '", E_NIVEAURELANCE=E_NIVEAURELANCE+1, E_NUMPIECEINTERNE="" ' +
               SQLRetrouve(vStAuxi, vStModele) + ' AND E_NIVEAURELANCE<' + IntToStr(RELNIVOMAX) +
               ' AND E_NIVEAURELANCE>=' + IntToStr(vInNivScore);
    ExecuteSQL(lStReq);

    // Nivo max atteind
    if (vInMaxNiv = (RELNIVOMAX + 1)) then MajNiveau7(vStAuxi, vStModele);
end;

procedure TOF_CPRELANCECLIENT.MajNiveau7(vStAuxi, vStModele : string);
var
    lStReq : string;
begin
    lStReq := 'UPDATE ECRITURE SET E_DATERELANCE="' + UsDateTime(FDtDateRel) +
              // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
              //'", E_NUMPIECEINTERNE="", E_REFPOINTAGE="" ' + SQLRetrouve(vStAuxi, vStModele) +
              '", E_NUMPIECEINTERNE="" ' + SQLRetrouve(vStAuxi, vStModele) +
              ' AND E_NIVEAURELANCE=' + IntToStr(RELNIVOMAX);
    ExecuteSQL(lStReq);
    FBoNivoMax := True;
end;

function TOF_CPRELANCECLIENT.SQLRetrouve(vStAuxi, vStModele : string) : string;
begin
    // FQ 15767 - Ne plus utiliser E_REFPOINTAGE
//    Result := ' WHERE ' + WhereGeneCritEcr(Pages, QEcr, True) + ' ' + WhereCritMP + ' AND E_NUMPIECEINTERNE="EDITREL"' + ' AND E_DATEMODIF="' + UsTime(FDtNowFutur) + '"';
    Result := ' WHERE ' + WhereGeneCritEcr(Pages, QEcr, True) + ' ' + WhereCritMP + ' AND E_NUMPIECEINTERNE LIKE "EDITREL%"' + ' AND E_DATEMODIF="' + UsTime(FDtNowFutur) + '"';

    if (vStAuxi <> '') then Result := Result + ' AND E_AUXILIAIRE="' + vStAuxi + '" ';
//    if (vStModele <> '') then Result := Result + ' AND E_REFPOINTAGE="' + vStModele + '" ';
    if (vStModele <> '') then Result := Result + ' AND E_NUMPIECEINTERNE="EDITREL;' + vStModele + '" ';
end;

procedure TOF_CPRELANCECLIENT.BValideClick(Sender : TObject);
var
    lIoRel : TIOErr;
    lInSel : Integer;
    lInCli : Integer;
    lTobCli : TOB;
begin

    if TFMul(Ecran).CritModified then
    begin
        PGIBox('Relance impossible : vous avez modifié les critères.', Ecran.Caption);
        Exit;
    end;

    if (PGIAsk('Confirmez-vous la relance des tiers sélectionnés ?', Ecran.caption) <> mrYes) then Exit;

    // On compte les lignes clients sélectionnées sans lignes d'échéance
    lInSel := 0;
    for lInCli := 1 to TobClient.Detail.Count - 1 do
    begin
        lTobCli := TobClient.Detail[lInCli];
        if (lTobCli.GetValue('TOTAL') <= 0) and (lTobCli.GetValue('SELECTION') = '*') then Inc(lInSel);
    end;

    V_PGI.MiseSousPli := GetCheckBoxState('CODEOMR') = cbChecked;
    FDtNowFutur := NowH;
    FBoNivoMax := false;
    FBoGenePremier := true;
    FBoYATP := false;

    SourisSablier;

    lIoRel := Transactions(ValideRelance, 3);

    case lIoRel of
        oeUnknown : PGIBox('Validation non effectuée !', Ecran.Caption);
        oeSaisie : PGIBox('Traitement impossible! Certaines échéances sont en cours de traitement par un autre utilisateur', Ecran.Caption);
        oeOk :
            begin
                if (FBoYATP and (GetControlText('MODELETP') <> '')) then
                    if (PGIAsk('Vous avez relancé des tiers payeurs. Voulez-vous éditer des documents à destination des tiers facturés ?', Ecran.Caption) = mrYes) then EditeRelancesTP;
                if (FBoNivoMax) then PGIInfo('Génération terminée. ATTENTION ! Certains tiers ont atteint le niveau maximum de relance.', Ecran.Caption);
                BChercheClick(nil);
            end;
    end;

    if (lInSel > 0) then PGIInfo('Attention : certaines échéances n''ont pas été générées.', Ecran.Caption);

    SourisNormale;
end;

procedure TOF_CPRELANCECLIENT.EditeRelancesTP;
var
    Q : TQuery;
    TT : Tlist;
    LL : TStrings;
    ListAux : HTStrings;
    lStSQL : string;
    lStSousReq : string;
    lStCond : string;
    lStMinExo : string;
    lStMaxExo : string;
    lStExo : string;
    lStAuxi : string;
    lDtMinDate : TDateTime;
    lDtMaxDate : TDateTime;
    lDtDate : TDateTime;
    lInMinNum : Integer;
    lInMaxNum : Integer;
    lInNum : Integer;
    lInCpt : Integer;
    lBoCarePDF : Boolean;
begin

    lBoCarePdf := False;

    RecupMSED(True);

    if FMSED.Spooler and (not V_PGI.QRPDF) then
    begin
        V_PGI.QRPDF := TRUE;
        lBoCarePdf := TRUE;
    end;

    lStSQL := 'SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_CONTREPARTIEAUX FROM ECRITURE WHERE ' + WhereGeneCritEcr(Pages, QEcr, True) + ' ' + WhereCritMP + ' AND E_MULTIPAIEMENT="TP" AND E_DATEMODIF="' + UsTime(FDtNowFutur) + '" AND E_QUALIFORIGINE="TP"';

    lInMinNum := 999999999;
    lInMaxNum := 0;
    lDtMinDate := iDate2099;
    lDtMaxDate := iDate1900;
    lStMinExo := 'zzz';
    lStMaxExo := '   ';
    ListAux := HTStringList.Create;

    Q := OpenSQL(lStSQL, true);
    while not Q.EOF do
    begin
        lStExo := Q.FindField('E_EXERCICE').AsString;
        if (lStExo > lStMaxExo) then lStMaxExo := lStExo;
        if (lStExo < lStMinExo) then lStMinExo := lStExo;

        lDtDate := Q.FindField('E_DATECOMPTABLE').AsDateTime;
        if (lDtDate > lDtMaxDate) then lDtMaxDate := lDtDate;
        if (lDtDate < lDtMinDate) then lDtMinDate := lDtDate;

        lInNum := Q.FindField('E_NUMEROPIECE').AsInteger;
        if (lInNum > lInMaxNum) then lInMaxNum := lInNum;
        if (lInNum < lInMinNum) then lInMinNum := lInNum;

        lStAuxi := Q.FindField('E_CONTREPARTIEAUX').AsString;
        if (ListAux.indexOf(lStAuxi) < 0) then ListAux.Add(lStAuxi);

        lStSQL := 'UPDATE ECRITURE SET E_MULTIPAIEMENT="TF" WHERE E_JOURNAL="' + Q.FindField('E_JOURNAL').AsString + '" AND E_EXERCICE="' + lStExo + '" AND E_DATECOMPTABLE="' + UsDateTime(lDtDate) + '" AND E_NUMEROPIECE=' + IntToStr(lInNum) + ' AND E_NUMLIGNE=1 AND E_MULTIPAIEMENT="" ';
        ExecuteSQL(lStSQL);
        Q.Next;
    end;

    Ferme(Q);
    FBoGenePremier := True;

    lStCond := ' E_JOURNAL="' + VH^.JalVTP + '" AND E_EXERCICE>="' + lStMinExo + '" AND E_EXERCICE<="' + lStMaxExo + '" AND E_DATECOMPTABLE>="' + UsDateTime(lDtMinDate) + '" AND E_DATECOMPTABLE<="' + UsDateTime(lDtMaxDate) + '" AND E_NUMEROPIECE>=' + IntToStr(lInMinNum) + ' AND E_NUMEROPIECE<=' + IntToStr(lInMaxNum) + ' AND E_NUMLIGNE=1 AND E_MULTIPAIEMENT="TF" ';

    for lInCpt := 0 to ListAux.Count - 1 do
    begin
        lStSousReq := ' AND ' + lStCond + ' AND E_MULTIPAIEMENT="TF" ';
        lStAuxi := ListAux[lInCpt];
        TT := TList.Create;
        LL := TStringList.Create;
        lStSQL := RecupSQLModele('L', 'REL', GetControlText('MODELETP'), '', lStAuxi, '', ' WHERE E_AUXILIAIRE="' + lStAuxi + '" ');
        LL.Add(lStSQL);
        LL.Add(lStSousReq);
        TT.Add(LL);

        Inc(FNbRelance);

    {$IFNDEF EAGLCLIENT}
        if FMSED.Spooler then LanceDocumentSpool('L', 'REL', GetControlText('ModeleTP'), TT, nil, GetCheckBoxState('APERCU') = cbChecked, FBoGenePremier, FALSE, FMSED.RepSpooler, FMSED.RacineSpooler, FMSED.XFichierSpooler, FMSED.SoucheSpooler)
        else begin
    {$ENDIF EAGLCLIENT}
            LanceDocument('L', 'REL', GetControlText('ModeleTP'), TT, nil, GetCheckBoxState('APERCU') = cbChecked, FBoGenePremier);
        {JP 02/08/05 : FQ 15627 : Comme on lance plusieurs fois le LanceEtat, mais qu'au-delà de première fois on
                       n'affiche plus la boîte de dialoque(cf ci-dessous), je récupère le nombre de copies demandées
                       sur la première impression et je l'affecte au nombre de copies par défaut}
            if FBoGenePremier then
              V_PGI.DefaultDocCopies := V_PGI.NbDocCopies;
    {$IFNDEF EAGLCLIENT}
        end;
    {$ENDIF EAGLCLIENT}

        LL.Free;
        TT.Free;
        FBoGenePremier := False;
    end;

    if (ListAux.Count > 0) then
    begin
        lStCond := ' E_JOURNAL="' + VH^.JalVTP + '" AND E_EXERCICE>="' + lStMinExo + '" AND E_EXERCICE<="' + lStMaxExo + '" AND E_DATECOMPTABLE>="' + UsDateTime(lDtMinDate) + '" AND E_DATECOMPTABLE<="' + UsDateTime(lDtMaxDate) + '" AND E_NUMEROPIECE>=' + IntToStr(lInMinNum) + ' AND E_NUMEROPIECE<=' + IntToStr(lInMaxNum) + ' AND E_MULTIPAIEMENT<>"" ';
        ExecuteSQL('UPDATE ECRITURE SET E_MULTIPAIEMENT="" WHERE ' + lStCond);
    end;

    ListAux.Clear;
    ListAux.Free;

    if lBoCarePDF then V_PGI.QRPDF := false;
end;

procedure TOF_CPRELANCECLIENT.FlapLigne;
var
    lInRow : Integer;
    lBoEtat : Boolean;
    lInCol : Integer;
begin
    lInRow := GClient.row;
    lBoEtat := GClient.IsSelected(lInRow);

    if lBoEtat then
    begin
        GEche.Font.Color := clWindowText;
        for lInCol := 1 to GEche.ColCount - 1 do GEche.ColColors[lInCol] := clWindowText;
    end
    else
    begin
        GEche.Font.Color := clGrayText;
        for lInCol := 1 to GEche.ColCount - 1 do GEche.ColColors[lInCol] := clGrayText;
    end;
end;

procedure TOF_CPRELANCECLIENT.BImprimerClick(Sender : TObject);
var
    lTobEdt : TOB;
    lTobLigne : TOB;
    lInEche : Integer;
    lInRow : Integer;
    lPagesCrit : TPageControl;
    lStCodeEtat : string;
begin
    lInRow := GClient.Row;
    if ((lInRow <= 0) or (lInRow > TobClient.Detail.Count)) then Exit;

{$IFDEF EAGLCLIENT}
    lPagesCrit := nil;
{$ELSE}
    lPagesCrit := Pages;
{$ENDIF EAGLCLIENT}

    if (PGIAsk('Souhaitez-vous imprimer le détail ?', Ecran.Caption) = mrYes) then
    // Edition des échéances
    begin
        lInEche := GEche.Row;
        if ((lInEche <= 0) or (lInEche > TobClient.Detail[lInRow - 1].Detail.Count)) then Exit;
        lTobEdt := TobClient.Detail[lInRow - 1].Detail[lInEche - 1];

        lTobEdt := TOB.Create('VETAT', nil, -1);
        for lInEche := 0 to TOBClient.Detail[lInRow - 1].Detail.Count - 1 do
        begin
            lTobLigne := TOB.Create('VLIGNES', lTobEdt, -1);
            lTobLigne.Dupliquer(TobClient.Detail[lInRow - 1].Detail[lInEche], False, True);
            // ajout infos auxiliaire
            lTobLigne.AddChampSupValeur('AUXILIAIRE', TobClient.Detail[lInRow - 1].GetValue('AUXILIAIRE'));
            lTobLigne.AddChampSupValeur('AFFAIRE', TobClient.Detail[lInRow - 1].GetValue('AFFAIRE'));
            lTobLigne.AddChampSupValeur('LIBELLE', TobClient.Detail[lInRow - 1].GetValue('LIBELLE'));
            lTobLigne.AddChampSupValeur('TOTAL', TobClient.Detail[lInRow - 1].GetValue('TOTAL'));
            lTobLigne.AddChampSupValeur('NIVEAUMAX', TobClient.Detail[lInRow - 1].GetValue('NIVEAUMAX'));
            lTobLigne.AddChampSupValeur('DATERELANCE', TobClient.Detail[lInRow - 1].GetValue('DATERELANCE'));
        end;

{$IFDEF EAGLCLIENT}
        lStCodeEtat := 'CRH';
{$ELSE}
        lStCodeEtat := 'CRE';
{$ENDIF EAGLCLIENT}

    end
    else
    // Edition des cumuls clients
    begin
        lTobEdt := TOB.Create('VETAT', nil, -1);
        for lInRow := 0 to TOBClient.Detail.Count - 1 do (TOB.Create('VLIGNES', lTobEdt, -1)).Dupliquer(TOBClient.Detail[lInRow], False, True);

        if GetCheckBoxState('RUPTUREAFFAIRE') = cbChecked then
        begin
{$IFDEF EAGLCLIENT}
            lStCodeEtat := 'CRF';
{$ELSE}
            lStCodeEtat := 'CRA';
{$ENDIF EAGLCLIENT}
        end
        else
        begin
{$IFDEF EAGLCLIENT}
            lStCodeEtat := 'CRL';
{$ELSE}
            lStCodeEtat := 'CRC';
{$ENDIF EAGLCLIENT}
        end;

    end;

    // Lancement de l'état
    LanceEtatTOB('E', 'CRC', lStCodeEtat, lTobEdt, True, False, False, lPagesCrit, '', Ecran.Caption, False);

    // Libération tob pour état
    lTobEdt.ClearDetail;
    FreeAndnil(lTobEdt);
end;

procedure TOF_CPRELANCECLIENT.SpoolerClick(Sender : TObject);
var
    lBoAvec : Boolean;
begin
    lBoAvec := GetCheckBoxState('SPOOLER') = cbChecked;

    SetControlEnabled('RepSpooler', lBoAvec);
    SetControlEnabled('TRepSpooler', lBoAvec);
    SetControlEnabled('XFichierSpooler', lBoAvec);

    if (lBoAvec) and (GetControlText('RepSpooler') = '') then SetControlText('RepSpooler', GetParamSocSecur('SO_CPCHEMINRELANCE',''));

end;

procedure TOF_CPRELANCECLIENT.RuptureAffaireClick(Sender : TObject);
begin
    InitGClient;
    InitGEche;
    VideLesListes;
    OnCritModif(nil);
end;

procedure TOF_CPRELANCECLIENT.POPSPopup(Sender : TObject);
begin
    InitPopup(Ecran);
end;

procedure TOF_CPRELANCECLIENT.ZoomCout(Sender : TObject);
{$IFNDEF EAGLCLIENT}
var
    lInRow : integer;
    lStAuxi : string;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
    PGIInfo('Fonction non implémentée en CWAS.', Ecran.Caption);
{$ELSE}
{$IFNDEF CCS3}
    lInRow := GClient.Row;
    if (lInRow <= 0) or (lInRow > TobCLient.Detail.Count) then Exit;

    lStAuxi := TobClient.Detail[lInRow - 1].GetValue('AUXILIAIRE');
    if (lStAuxi <> '') then LanceCoutTiers(lStAuxi, StrToDate(GetControlText('DATERELANCE')));
{$ENDIF}
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 18/08/2004
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices
Suite ........ : decalees
Mots clefs ... :
*****************************************************************}

procedure TOF_CPRELANCECLIENT.ZoomEcriture(Sender : TObject);
var
    lInRow : Integer;
    lInEche : Integer;
    lTobEche : TOB;
    P : RParFolio;
    M : RMVT;
begin
    lInRow := GClient.Row;
    if ((lInRow <= 0) or (lInRow > TobClient.Detail.Count)) then Exit;

    lInEche := GEche.Row;
    if ((lInEche <= 0) or (lInEche > TobClient.Detail[lInRow - 1].Detail.Count)) then Exit;

    lTobEche := TobClient.Detail[lInRow - 1].Detail[lInEche - 1];
    if (lTobEche = nil) then Exit;

    FillChar(M, Sizeof(M), #0);
    M.Etabl := lTobEche.GetValue('E_ETABLISSEMENT');
    M.Jal := lTobEche.GetValue('E_JOURNAL');
    M.Exo := lTobEche.GetValue('E_EXERCICE');
    M.DateC := lTobEche.GetValue('E_DATECOMPTABLE');
    M.Nature := lTobEche.GetValue('E_NATUREPIECE');
    M.Num := lTobEche.GetValue('E_NUMEROPIECE');
    M.Simul := lTobEche.GetValue('E_QUALIFPIECE');
    M.CodeD := lTobEche.GetValue('E_DEVISE');
    M.DateTaux := lTobEche.GetValue('E_DATETAUXDEV');
    M.TauxD := lTobEche.GetValue('E_TAUXDEV');
    M.Valide := lTobEche.GetValue('E_VALIDE') = 'X';
    M.ModeSaisieJal := lTobEche.GetValue('E_MODESAISIE');

    // SAISIE BORDEREAU
    if ((M.ModeSaisieJal <> '-') and (M.ModeSaisieJal <> '')) then
    begin
        FillChar(P, Sizeof(P), #0);
        P.ParPeriode := DateToStr(M.DateC);
        P.ParCodeJal := M.Jal;
        P.ParNumFolio := IntToStr(M.Num);
        P.ParNumLigne := lTobEche.GetValue('E_NUMLIGNE');
        ChargeSaisieFolio(P, taConsult);
    end
    // SAISIE PIECE
    else LanceSaisie(nil, taConsult, M);
end;

procedure TOF_CPRELANCECLIENT.ZoomParam(Sender : TObject);
var
    lStTypeRel : string;
begin
    if ModeTraite then lStTypeRel := 'RTR'
    else lStTypeRel := 'RRG';

    CCLanceFiche_ParamRelance(lStTypeRel);
end;

procedure TOF_CPRELANCECLIENT.ZoomTiers(Sender : TObject);
var
    lInRow : integer;
    lStAuxi : string;
begin
    lInRow := GClient.Row;
    if ((lInRow <= 0) or (lInRow > TobCLient.Detail.Count)) then Exit;

    lStAuxi := TobClient.Detail[lInRow - 1].GetValue('AUXILIAIRE');
    if (lStAuxi <> '') then FicheTiers(nil, '', lStAuxi, taConsult, 1);
end;

procedure TOF_CPRELANCECLIENT.UnModeleClick(Sender : TObject);
begin
    SetControlVisible('ModeleUnique', GetCheckBoxState('UnModele') = cbChecked);
end;
(*
procedure TOF_CPRELANCECLIENT.ZoomEdtJustif(Sender : TObject);
{$IFNDEF EAGLCLIENT}
var
    lCrit : TCritEdt;
    lDtDeb : TdateTime;
    lInRow : integer;
    lStAuxi : string;
{$ENDIF EAGLCLIENT}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    FDtDateRel := StrToDate(GetControlText('DATERELANCE'));

    lInRow := GClient.Row;
    if (lInRow <= 0) or (lInRow > TobCLient.Detail.Count) then Exit;
    lStAuxi := TobClient.Detail[lInRow - 1].GetValue('AUXILIAIRE');
    if lStAuxi = '' then Exit;

    Fillchar(lCrit, SizeOf(lCrit), #0);

    lDtDeb := FDtDateRel;
    lCrit.Date1 := lDtDeb;
    lCrit.DateDeb := lCrit.Date1;
    lCrit.NatureEtat := neJU;

    InitCritEdt(lCrit);

    lCrit.JU.OnTiers := TRUE;
    lCrit.Cpt1 := lStAuxi;
    lCrit.Cpt2 := lCrit.Cpt1;

    JustSoldeZoom(lCrit);
{$ENDIF EAGLCLIENT}
end;
*)
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 07/10/2004
Modifié le ... : 07/10/2004
Description .. : - BPY le 07/10/2004 - Mise a jour de plusieur lignes a la fois
Suite ........ : !
Mots clefs ... :
*****************************************************************}

procedure TOF_CPRELANCECLIENT.NiveauModif(Sender : TObject);
var
    i, j : integer;
    lTobCli, lTobEche : TOB;
    NiveauRelance, NbEche : integer;
    BlocNote : string;
begin
    if not ModeManuel then Exit;

    // init des champs
    NiveauRelance := 0;
    NbEche := 0;
    BlocNote := '';

    // creation de la TOB
    TheData := TOB.Create('', nil, -1);

    // parcour de la grille des client
    for i := 0 to TobClient.Detail.count - 1 do
    begin
        // recup de la tob
        lTobCli := TobClient.Detail[i];
        if (lTobCli = nil) then continue;

        // si pas selectionné on quitte
        if (not (lTobCli.GetValue('SELECTION') = '*')) then continue;

        // parcour de la grille des echeance
        for j := 0 to lTobCli.Detail.count - 1 do
        begin
            // recup de la tob
            lTobEche := lTobCli.Detail[j];
            if (lTobEche = nil) then continue;

            // si pas selectionné on quitte
            if (not (lTobEche.GetValue('SELECTION') = '*')) then continue;

            // set des valeur
            if (NiveauRelance < lTobEche.GetValue('NIVEAURELANCE')) then NiveauRelance := lTobEche.GetValue('NIVEAURELANCE');
            NbEche := NbEche + 1;
            BlocNote := lTobEche.GetValue('E_BLOCNOTE');
        end;
    end;

    // init des valeurs
    TOB(TheData).AddChampSupValeur('NIVEAURELANCE', NiveauRelance);
    if (NbEche > 1) then TOB(TheData).AddChampSupValeur('BLOCNOTE', '')
    else TOB(TheData).AddChampSupValeur('BLOCNOTE', BlocNote);
    TOB(TheData).AddChampSupValeur('MAJBLOCNOTE', false);
    TOB(TheData).AddChampSupValeur('NBECHE', NbEche);

    // lancement de la fiche
    if not (CPLanceFiche_RelanceNiveau('', '', '')) then Exit;

    // modif
    ModificationNiveau(TOB(TheData).GetValue('NIVEAURELANCE'), true, TOB(TheData).GetValue('BLOCNOTE'), TOB(TheData).GetValue('MAJBLOCNOTE'));

    // fin
    FreeAndNil(TheData);
    TheData := nil;

    // Maj total ligne client
    GClientRecalculLigne(GClient.row, True);

    // MAJ affichage lignes des grilles
    RemplitGClient;
    RemplitGEche;
end;

procedure TOF_CPRELANCECLIENT.NiveauUp(Sender : TObject);
begin
    ModificationNiveau(1, false, '', false);
end;

procedure TOF_CPRELANCECLIENT.NiveauDown(Sender : TObject);
begin
    ModificationNiveau(-1, false, '', false);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 07/10/2004
Modifié le ... : 07/10/2004
Description .. : - BPY le 07/10/2004 - Mise a jour de plusieur lignes a la fois
Suite ........ : !
Mots clefs ... :
*****************************************************************}

procedure TOF_CPRELANCECLIENT.ModificationNiveau(Montant : integer;FromZero : Boolean;BlocNote : string;MajBlocNote : boolean);
var
    i, j : integer;
    lInNivo,lInMax : Integer;
    lTobCli,lTobEche : TOB;
    OldBlocNote : string;
    lTobEcr : TOB ; // Pour l'update DB // FQ 17060 SBO 05/12/2005
begin
    if (not ModeManuel) then exit;

    // parcour de la grille des client
    for i := 0 to TobClient.Detail.count - 1 do
    begin
        // init du max client
        lInMax := 0;

        // recup de la tob
        lTobCli := TobClient.Detail[i];
        if (lTobCli = nil) then continue;

        // si pas selectionné on quitte
        if (not (lTobCli.GetValue('SELECTION') = '*')) then continue;

        // parcour de la grille des echeance
        for j := 0 to lTobCli.Detail.count - 1 do
        begin
            // recup de la tob
            lTobEche := lTobCli.Detail[j];
            if (lTobEche = nil) then continue;

            // recup du niveau actuel d'echeance
            lInNivo := lTobEche.GetValue('NIVEAURELANCE');

            // si pas selectionné on quitte
            if (not (lTobEche.GetValue('SELECTION') = '*')) then
            begin
                // calcul nivo max pour toute les echeance non modifier
                if (lInNivo > lInMax) then lInMax := lInNivo;
                continue;
            end;

            // recup du niveau
            if (FromZero) then lInNivo := Montant
            else lInNivo := lInNivo + Montant;

            // test du niveau
            if (lInNivo < 1) then lInNivo := 1;
            if (lInNivo > (RELNIVOMAX + 1)) then lInNivo := RELNIVOMAX + 1;

            // calcul nivo max si cette echeance est modifier !
            if (lInNivo > lInMax) then lInMax := lInNivo;

            // Si valeur a été modifié alors traitement
            OldBlocNote := lTobEche.GetValue('E_BLOCNOTE');
            if ((not (lInNivo = lTobEche.GetValue('NIVEAURELANCE'))) or ((MajBlocNote) and (not (BlocNote = OldBlocNote)))) then
            begin
                if (MajBlocNote) then
                begin
                    // MAJ objet
                    lTobEche.PutValue('NIVEAURELANCE', lInNivo);
                    lTobEche.PutValue('E_BLOCNOTE',    BlocNote);
                    // FQ 17060 SBO 05/12/2005 : Remplacement de l'update via execute SQL par un Tob.UpdateDB
                    // Chargement de l'ecriture
                    lTobEcr := TOB.Create('ECRITURE',nil, -1);
                    lTObEcr.PutValue('E_JOURNAL',       lTobEche.GetValue('E_JOURNAL') );
                    lTObEcr.PutValue('E_EXERCICE',      lTobEche.GetValue('E_EXERCICE') );
                    lTObEcr.PutValue('E_DATECOMPTABLE', lTobEche.GetValue('E_DATECOMPTABLE') );
                    lTObEcr.PutValue('E_NUMEROPIECE',   lTobEche.GetValue('E_NUMEROPIECE') );
                    lTObEcr.PutValue('E_NUMLIGNE',      lTobEche.GetValue('E_NUMLIGNE') );
                    lTObEcr.PutValue('E_NUMECHE',       lTobEche.GetValue('E_NUMECHE') );
                    lTObEcr.PutValue('E_QUALIFPIECE',   lTobEche.GetValue('E_QUALIFPIECE') );
                    lTobEcr.LoadDB ;
                    // Maj Base
                    lTobEcr.PutValue('E_NIVEAURELANCE', lInNivo - 1 ) ;
                    lTobEcr.PutValue('E_BLOCNOTE',      BlocNote ) ;
                    lTobEcr.UpdateDB ;
                    FreeAndNil( lTobEcr ) ;
                    // MAJ base ( Attention : Nivo affiché - 1 ) !!!
//                    ExecuteSQL('UPDATE ECRITURE SET E_NIVEAURELANCE=' + IntToStr(lInNivo - 1) + ',E_BLOCNOTE="' + ( CheckdblQuote(BlocNote) ) + '" WHERE ' + WhereEcritureTob(tsGene, lTobEche, True))
                    // fin FQ 17060 SBO 05/12/2005 
                end
                else
                begin
                    // MAJ objet
                    lTobEche.PutValue('NIVEAURELANCE', lInNivo);
                    // MAJ base ( Attention : Nivo affiché - 1 ) !!!
                    {JP 21/07/06 : Modification de la fonction WhereEcritureTob}
                    ExecuteSQL('UPDATE ECRITURE SET E_NIVEAURELANCE=' + IntToStr(lInNivo - 1) + ' WHERE ' + WhereEcritureTob(tsGene, lTobEche, True, False, 'E'));
                end;
            end;
        end;

        // Maj nivo max du client si modifié
        if (lInMax <> lTobCli.GetValue('NIVEAUMAX')) then lTobCli.PutValue('NIVEAUMAX', lInMax);

        // MAJ grilles
        GClientMajLigne(i + 1);
        RemplitGEche;
    end;
end;

procedure TOF_CPRELANCECLIENT.BRechercherClick(Sender : TObject);
begin
    FBoFindFirst := true;
    FindMvt.Execute;
end;

procedure TOF_CPRELANCECLIENT.VideLesListes;
begin
    // Libération objets
    TobClient.ClearDetail;

    // On vide la grille des cumuls client
    GClient.RowCount := 2;
    GClient.Cells[RT_Aux, 1] := '';
    GClient.Cells[RT_Int, 1] := '';
    if (RT_Aff > 0) then GClient.Cells[RT_Aff, 1] := '';
    GClient.Cells[RT_Tot, 1] := '';
    GClient.Cells[RT_Niv, 1] := '';
    GClient.Cells[RT_Der, 1] := '';
    GClient.Cells[RT_Sel, 1] := ' ';

    // On vide la grille des échéances
    GEche.RowCount := 2;
    GEche.Cells[RE_Num, 1] := '';
    GEche.Cells[RE_Pai, 1] := '';
    GEche.Cells[RE_Ech, 1] := '';
    GEche.Cells[RE_Tot, 1] := '';
    GEche.Cells[RE_Niv, 1] := '';
    GEche.Cells[RE_Dif, 1] := '';
    GEche.Cells[RE_Sel, 1] := ' ';

end;

procedure TOF_CPRELANCECLIENT.FindMvtFind(Sender : TObject);
begin
    Rechercher(GClient, FindMvt, FBoFindFirst);
end;

procedure TOF_CPRELANCECLIENT.GClientBeforeFlip(Sender : TObject;ARow : Integer;var Cancel : Boolean);
begin
    Cancel := ((not FBoFlipAuto) and (not ModeManuel));
end;

procedure TOF_CPRELANCECLIENT.GEcheBeforeFlip(Sender : TObject;ARow : Integer;var Cancel : Boolean);
begin
    Cancel := ((not FBoFlipAuto) and (not ModeManuel));
end;

procedure TOF_CPRELANCECLIENT.OnEcranResize(Sender : TObject);
begin
    // Simulation de l'héritage
    if Assigned(OnParentEcranResize) then OnParentEcranResize(Sender);

    GClientTailleCol;
    GEcheTailleCol;

end;

procedure TOF_CPRELANCECLIENT.GClientTailleCol;
var
    lInBase : Integer;
    lInNew : Integer;
begin

    lInNew := GClient.Width;
    if (GetCheckBoxState('RUPTUREAFFAIRE') = cbChecked) then lInBase := 520
    else lInBase := 420;

    // Colonne sélection invisible
    GClient.ColWidths[RT_Sel] := -1;
    // Colonne Auxiliaire
    GClient.ColWidths[RT_Aux] := Trunc(80 * lInNew / lInBase);
    // Colonne Intitulé
    GClient.ColWidths[RT_Int] := Trunc(140 * lInNew / lInBase);
    // Colonne Affaire
    if (RT_Aff > 0) then GClient.ColWidths[RT_Aff] := Trunc(100 * lInNew / lInBase);
    // Colonne Total
    GClient.ColWidths[RT_Tot] := Trunc(88 * lInNew / lInBase);
    // Colonne Niveau max
    GClient.ColWidths[RT_Niv] := Trunc(30 * lInNew / lInBase);
    // Colonne Date dernière relance
    GClient.ColWidths[RT_Der] := Trunc(60 * lInNew / lInBase);

end;

procedure TOF_CPRELANCECLIENT.GEcheTailleCol;
var
    lInBase : Integer;
    lInNew : Integer;
begin

    lInNew := GEche.Width;
    lInBase := 370;
    // Colonne sélection invisible
    GEche.ColWidths[RE_Sel] := -1;
    // Colonne Numéro
    GEche.ColWidths[RE_Num] := Trunc(50 * lInNew / lInBase);
    // Colonne Mode paiement
    GEche.ColWidths[RE_Pai] := Trunc(50 * lInNew / lInBase);
    // Colonne Date Echéance
    GEche.ColWidths[RE_Ech] := Trunc(80 * lInNew / lInBase);
    // Colonne Montant
    GEche.ColWidths[RE_Tot] := Trunc(90 * lInNew / lInBase);
    // Colonne Niveau Relance
    GEche.ColWidths[RE_Niv] := Trunc(30 * lInNew / lInBase);
    // Colonne Ecart
    GEche.ColWidths[RE_Dif] := Trunc(50 * lInNew / lInBase);

end;

procedure TOF_CPRELANCECLIENT.OnCritModif(Sender : TObject);
begin
    TFMul(Ecran).CritModified := True;
end;

procedure TOF_CPRELANCECLIENT.BChercheClick(Sender : TObject);
begin

    // Au chargement, on ne fait rien !
    if FBoChargEnCours then Exit;

    // Simulation de l'héritage
    if Assigned(OnParentBChercheClick) then OnParentBChercheClick(TFMul(Ecran).bExport);
end;

procedure TOF_CPRELANCECLIENT.ZoomEdtJustifNew(Sender: TObject);
Var
    lInRow : integer;
    lStAuxi : String;
    lCrit : ClassCritEdt;
begin
  FDtDateRel := StrToDate(GetControlText('DATERELANCE'));

  lInRow := GClient.Row;
  if ((lInRow <= 0) or (lInRow > TobCLient.Detail.Count)) then Exit;

  lStAuxi := TobClient.Detail[lInRow - 1].GetValue('AUXILIAIRE');
  if (lStAuxi = '') then Exit;

  lCrit := ClassCritEdt.Create;
  Fillchar(lCrit.CritEdt, SizeOf(lCrit.CritEdt), #0);

  lCrit.CritEdt.Exo.Code := ''; // Exercice

  if GetControlText('E_NUMEROPIECE' ) = '' then lCrit.CritEdt.GL.NumPiece1 := 0
                                           else lCrit.CritEdt.GL.NumPiece1 := StrToInt(GetControlText('E_NUMEROPIECE'));
  if GetControlText('E_NUMEROPIECE_') = '' then lCrit.CritEdt.GL.NumPiece2 := 999999999
                                           else lCrit.CritEdt.GL.NumPiece2 := StrToInt(GetControlText('E_NUMEROPIECE_'));
  lCrit.CritEdt.GL.EnDateSituation := True; // Justificatif de Solde
  lCrit.CritEdt.Qualifpiece := 'N'; // Type d'Ecritures
  lCrit.CritEdt.Etab := GetControlText('E_ETABLISSEMENT'); // Etablissement
  lCrit.CritEdt.DeviseSelect := GetControlText('E_DEVISE'); // Devise
  lCrit.CritEdt.GL.Lettrable := 1; // Ecritures Lettrées

  lCrit.CritEdt.Date1 := StrToDate(GetControlText('E_DateComptable')); // Date Début
  lCrit.CritEdt.Date2 := FDtDateRel ;                                  // Date de situation

  lCrit.CritEdt.Cpt1 := lStAuxi; // auxiliaire min
  lCrit.CritEdt.Cpt2 := lStAuxi; // auxiliaire max

  TheData := lCrit;
  {$IFNDEF IMP}
  CPLanceFiche_CPGLAUXI('');
  {$ENDIF IMP}
  TheData := nil;
end;


procedure TOF_CPRELANCECLIENT.OnAfterShow;
begin
    FBoChargEnCours := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRELANCECLIENT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 26/04/2007
Modifié le ... :   /  /    
Description .. : Permet de retourner le bon format pour les colonnes 
Suite ........ : montant et total en fonction des paramètres sociétés
Suite ........ : (cf FQ 20048)
Mots clefs ... : 
*****************************************************************}
function TOF_CPRELANCECLIENT.GetFormatCol : string; // FQ 20048
var
  nb : integer;
  i  : integer;
begin
  result := '#,##0';
  nb := GetParamSocSecur('SO_DECVALEUR','2');
  for i := 1 to nb do
  begin
     if i = 1 then
        result := result + '.';
     result := result + '0';
  end;
end;

initialization
    registerclasses([TOF_CPRELANCECLIENT]);
end.
