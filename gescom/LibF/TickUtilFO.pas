{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/06/2001
Modifié le ... : 23/07/2001
Description .. : Fonctions utilitaires pour la gestion des tickets du Front
Suite ........ : Office
Mots clefs ... : FO
*****************************************************************}
unit TickUtilFO;

interface
uses
  classes, Controls, extctrls, Sysutils, graphics, Windows, UTob, LookUp,
  {$IFDEF EAGLCLIENT}
  Maineagl, UtileAgl,
  {$ELSE}
  dbtables, FE_Main, EdtREtat, EdtRDoc,
  {$ENDIF}
  HCtrls, Hent1, HDimension, ParamSoc, HMsgBox, AGLInit, ED_Tools,
  KB_Ecran, HPanel, LCD_Lab, SaisUtil, EntGC, FactPieceContainer,
  FactFormule; // JTR - eQualité 11203

///////////////////////////////////////////////////////////////////////////////////////
//  Caractéristiques des colonnes des grilles.
///////////////////////////////////////////////////////////////////////////////////////
type RColCarac = record
    NoPage: Integer; // N° page de pavé associée à la colonne
    Largeur: Integer; // Largeur de la colonne
    Align: TAlignment; // Alignement de la colonne
  end;

///////////////////////////////////////////////////////////////////////////////////////
//  Caractéristiques du pave secondaire.
///////////////////////////////////////////////////////////////////////////////////////
type RPropPave2 = record
    Visible: boolean; // Pavé secondaire visible
    Largeur: integer; // Largeur du pavé
    Align: TAlignment; // Alignement du pavé
    NbrBtnHeight: integer; // Nombre de boutons en hauteur
    NbrBtnWidth: integer; // Nombre de boutons en largeur
    ClcPosition: TPosClc; // Alignement de la calculette
    ClcVisible: boolean; // Affichage de la calculette
  end;

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Code de retour de la vérification d'un article FOArticleAutorise
  ///////////////////////////////////////////////////////////////////////////////////////
type T_FOArtAutorise = (afoOk, afoFerme, afoInCompatible, afoInterdit, afoConfidentiel);

  ///////////////////////////////////////////////////////////////////////////////////////
  //  déclarations des fonctions et procédures.
  ///////////////////////////////////////////////////////////////////////////////////////
function FORepresentantExiste(CodeComm, TypeCom: string; TOBComms: TOB): Boolean;
function FOGetVendeur(CC: TControl; TitreSel, ZoneCom, Etablis: string): boolean;
function FOPreAffecteVendeur(TOBTiers: TOB): string;
procedure FOPreAffecteVendeurLigne(TOBPiece, TOBLigne, TOBTiers: TOB; ARow: Integer);
function FOGetDemarqueSaisi(CC: TControl; TitreSel: string; Letl: TTypeLocate; BFidelite: Boolean = False): Boolean;
function FODemarqueClientOblig(CodeDem: string): Boolean;
function FODemarquePeriode(CodeDem: string; DatePiece: TDateTime): boolean;
function FOSeuilCliOblig(CA: Double): Boolean;
function FOOpCaisseCliOblig(TOBLig, TOBArt: TOB): string;
function FOMakeNomClient(TOBTiers: TOB; CodeTiers: string = ''): string;
function FOGetTiersMulLibelle (G_CodeTiers: THCritMaskEdit): boolean;
function FOMontreStockDispo(TOBArt, TOBLig: TOB): string;
function FOMontreInfoArticle(TOBArt: TOB): string;
function FOGetArticleRecherche(CC: TControl; TitreSel, NaturePieceG: string): boolean;
procedure FOAjoutArticleCompl(TOBArt: TOB);
function FOArticleAutorise(TOBPiece, TOBArticles: TOB; NaturePiece: string; ARow: Integer): T_FOArtAutorise;
function FOVerifMixageArticle(TOBPiece, TOBArticles: TOB; var ARow: Integer): Boolean;
function FOSigneQteArticle(TOBPiece, TOBArticles: TOB; ARow: integer; Qte: double): double;
function FOLibelleDimension(TOBArt: TOB; Abrege: Boolean = True; Compact: Boolean = False; AvecTitre: Boolean = True; Separateur: string = '-:'): string;
function FOLibelleFamille(TOBArt: TOB; Abrege: Boolean = True): string;
function FOLibelleUneFamille(TOBArt: TOB; Champ: string; Abrege: Boolean = False): string;
function FOAffichePhotoArticle(CodeArt: string; Image: TImage): Boolean;
function FOSetNumeroDefinitif(PieceContainer: TPieceContainer): Boolean;
function FOEstUneLigneVente(TOBLigne: TOB): Boolean;
function FOPieceUneLigne(TOBPiece: TOB): Boolean;
function FOLigneRemisable(TOBLigne: TOB; EnPied: Boolean): Boolean;
procedure FOAffecteFactureHT(TOBPiece: TOB);
function FOFabriqueCommentairePiece(TOBPiece: TOB; Texte: string): string;
procedure FOInsereLigneCommentaire(TOBPiece: TOB; Texte: string; EnFin: Boolean);
function FOQuantiteAutorise(Qte: double): boolean;
function FOPrixUnitaireAutorise(PU: double): boolean;
procedure FOMajQte(TOBLigne: TOB; EnHT: Boolean);
procedure FOImprimeTexte(TOBPiece: TOB; CleDoc: R_CleDoc; Nature, Modele, Langue: string; Init_Imprimante, PreView: Boolean);
procedure FOImprimerLaPiece(TOBPiece, TOBTiers, TOBEches, TOBComms: TOB; CleDoc: R_CleDoc; MiseAJour: Boolean = True; PreViewForce: Boolean = False;
  OuvrirTiroir: Boolean = False; ChoixModele: Boolean = False);
function FOChoixModeleImpression(var Format, Modele: string; var EnPlus: Boolean; var NbCopies: Integer): Boolean;
function FOImprimeTicket(TOBPiece: TOB; CleDoc: R_CleDoc): Boolean;
procedure FOReImpTicket(ChoixModele: boolean);
procedure FOReImprimeBons;
procedure FOImpressionBons(TOBPiece, TOBArticles: TOB);
procedure FOImpressionCheque(TOBEche, TOBMode: TOB);
procedure FOImpressionAvoirs(TOBEche: TOB);
function FOCalculeMargeLigne(TOBLigne: TOB): double;
procedure FOConvtMntEcheance(MntDevSaisi: Double; CodeDev: string; TOBPiec, TOBDev: TOB; DEV: RDEVISE; var MontantP, MontantD: Double);
function FOMarqueChqDifUtilise(TOBEches: TOB; MajSql: Boolean = False): Boolean;
procedure FOMarqueTicketAnnule(TOBPiece_O, TOBPiece: TOB);
function FOVerifTicketAnnulable(CleDoc: R_CleDoc): boolean;
function FOCalculBaseRemiseGlobale(TOBPiece: TOB; Tous: boolean = False): double;
procedure FOAppliqueRemiseGlobale(TOBPiece: TOB; RemiseGlobale: double; TypeRemise: string; DEV: RDEVISE);
function FOAppliqueRemiseSousTotal(TOBPiece: TOB; ARow: integer; TauxRemise, RemiseSousTotal, MontantNet: double; TypeRemise: string; DEV: RDEVISE; var MaxRem:
  integer): integer;
function FOMisePieceEnAttente(PieceContainer: TPieceContainer) : integer;

procedure FOSupprimePieceEnAttente(Caisse: string);
function FOReprisePieceEnAttente(TOBPiece: TOB; GS: THGrid; DejaSaisie: Boolean; var TobAtt : TOB; TobLigFormule : TOB = nil): Boolean;
procedure FOAlerteAnnoncesNonValidees(Etablis, Depot: string);
function FOComptaEstActive(NaturePiece: string): Boolean;
function FOGenereTicket(CodeArt, CodeVen, CodeMode, CodeDev, CodeTiers, Caption: string; MntEche: Double; Affichage, Impression, RattachNumZ: Boolean):
  Integer;
procedure AffichePaveContextuel(PnlBoutons: TClavierEcran; NoPage: Integer);
procedure ChargePaveContextuel(PP: array of THPanel; GS: array of THGrid; LCD: array of TLCDLabel; var ColCarac: array of RColCarac; WC: array of TWinControl; var PropPv2: RPropPave2);
procedure FOAppliqueColCarac(GS: THGrid; ColCarac: array of RColCarac; IndDebut: Integer);
procedure FOInitColCarac(ColCarac: array of RColCarac);
function FODonneParamClavierEcran(ValeurParDefaut: boolean): string;
procedure FOCreateClavierEcran(var PnlBoutons: TClavierEcran; AOwner: TComponent; Pmain: TPanel; ChargeAuto: boolean = True; NbBtnDefaut: boolean = False);
function FOIsOpCaisseNoBon(Prefixe, TypeOpe: string; Montant: double): boolean;
function FOIsOpCaisseCumul(Prefixe, TypeOpe: string; Montant: double): boolean;
function FOIsOpCaisseLiee(Prefixe, TypeOpe: string; Montant: double): boolean;
procedure FOValideLesOperCaisse(TOBPiece, TOBArticles, TOBEches, TOBPiece_O: TOB; AnnulPiece: boolean);
procedure FOLoadOperCaisse(TOBPiece, TOBEches, TOBOpCais: TOB);
function FORechercheOpCaisseDispo(TOBLigne, TOBPiece, TOBEches: TOB; TypeMode: string; MntReste: double; var Montant: double): boolean;
function FORechercheReglementDispo(TOBLigne, TOBPiece, TOBEches: TOB; TypeOpe: string; var Montant: double; var Trouve: boolean): boolean;
function FOVerifReglementLie(TOBPiece, TOBEches, TOBArticles: TOB; AnnulPiece: boolean; var NoLig: integer): integer;
function FORepriseOpCaisseLies(CleDoc: R_CleDoc): boolean;
function FORepriseReglementsLies(CleDoc: R_CleDoc): boolean;
procedure FOAnnulReglementLie(TOBPiece, TOBArticles, TOBEches, TOBPiece_O: TOB);
function FOAffecteNoBonReglementLie(TOBPiece, TOBArticles, TOBEches: TOB): boolean;
type T_MajPieceEnAttente = class
    procedure MajPieceEnAttente;
  end;

implementation
uses
  UtilPGI, uRecupSQLModele, Ent1, TarifUtil, UtilGC,
  UtilArticle, NomenUtil, FactComm, FactUtil, FactCalc, FactCpta,
  FactTOB, FactPiece, FactArticle, FactTiers,
  {$IFDEF FOS5}
  MC_Erreur, LP_impri,
  {$ENDIF}
  FODefi, FOUtil;

var PieceContainerLoc : TPieceContainer; // JTR - eQualité 11203

///////////////////////////////////////////////////////////////////////////////////////
//  déclarations du composant pour la génération de ticket
///////////////////////////////////////////////////////////////////////////////////////
type TFOGenTicket = class(TObject)
  private
    Titre: string; // Titre des messages
    Tiers: string; // Code du tiers
    Vendeur: string; // Code du vendeur
    Article: string; // Code de l'article
    ModePaie: string; // Modes de paiement
    Devise: string; // Devise du mode de paiement correspondant
    MontDev: Double; // Montant en devise du mode de paiement correspondant
    Montant: Double; // Montant en devise de la pièce
    DEV: RDEVISE; // Devise de la pièce
    CleDoc: R_CleDoc; // Clé de la pièce
    AffMsg: Boolean; // Affichage des messages d'erreur
    ImpTic: Boolean; // Impression du ticket
    RatNumZ: Boolean; // rattachement au n° de clôture
    ActiveCompta: Boolean; // Activation de la comptabilisation
    TOBPiece, TOBBases, TOBTiers, TOBArticles, TOBPorcs, TOBEches, TOBAffaire, TOBConds, TOBAcomptes, TOBCpta, TOBAnaP, TOBAnaS, TOBComms: TOB;
    PieceContainer: TPieceContainer;
    function MAJPiece: Boolean;
    function MAJLigne: Boolean;
    procedure TraiteLaCompta(ARow: Integer);
    function MAJPiedEche: Boolean;
    procedure InitToutModif;
    procedure ValideLaPiece;
    procedure ValideImpression;
  public
    constructor Create; overload;
    destructor Destroy; override;
    function Init(CodeArt, CodeVen, CodeMode, CodeDev, CodeTiers, Caption: string; MntEche: Double; Affichage, Impression, RattachNumZ: Boolean): Boolean;
    function GenerePiece: Boolean;
  end;

  ///////////////////////////////////////////////////////////////////////////////////////
  //  déclarations du composant pour la mise à jour des numéros des bons
  ///////////////////////////////////////////////////////////////////////////////////////
type
  TFOMajNoBon = class(TObject)
  public
    Num: integer;
    NbBon: integer;
    procedure MAJNumeroBon;
  end;

  {==============================================================================================}
  {============================== TRAITEMENT DES VENDEURS =======================================}
  {==============================================================================================}
  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 23/07/2001
  Modifié le ... : 23/07/2001
  Description .. : Vérifie si un code vendeur existe
  Mots clefs ... : FO
  *****************************************************************}

function FORepresentantExiste(CodeComm, TypeCom: string; TOBComms: TOB): Boolean;
var QQ: TQuery;
  sSql: string;
begin
  Result := False;
  if CodeComm = '' then Exit;
  // Recherche dans la TOB
  if TOBComms <> nil then
  begin
    if TOBComms.FindFirst(['GCL_COMMERCIAL'], [CodeComm], False) <> nil then
    begin
      Result := True;
      Exit;
    end;
  end;
  // Recherche dans la table
  if TypeCom = '' then TypeCom := 'VEN';
  sSql := 'SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_TYPECOMMERCIAL="' + TypeCom + '" '
    + ' AND GCL_COMMERCIAL="' + CodeComm + '" AND GCL_DATESUPP>"' + USDateTime(Date) + '"';
  QQ := OpenSQL(sSql, True);
  Result := (not QQ.EOF);
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 19/09/2003
Description .. : Recherche d'un vendeur sous forme d'un Lookup
Mots clefs ... : FO
*****************************************************************}

function FOGetVendeurLookUp(CC: TControl; TitreSel, Etablis: string; Letl: TTypeLocate): boolean;
var
  NomChamp, CodeC, sWhere: string;
begin
  sWhere := 'GCL_TYPECOMMERCIAL="VEN" AND GCL_DATESUPP>"' + USDateTime(Date) + '"';
  
  CodeC := GetInfoParPiece(FOGetNatureTicket(False, False), 'GPP_FILTRECOMM');
  NomChamp := RechDom('GCINFOCOMMERCIAL', CodeC, True);
  if NomChamp = 'GCL_ETABLISSEMENT' then
    sWhere := sWhere + ' AND GCL_ETABLISSEMENT="' + Etablis + '"';

  Result := LookupList(CC, TitreSel, 'COMMERCIAL', 'GCL_COMMERCIAL', 'GCL_LIBELLE,GCL_PRENOM',
    sWhere, 'GCL_COMMERCIAL', False, 0, '', Letl);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 19/09/2003
Description .. : Recherche d'un vendeur sous forme d'un MUL
Mots clefs ... : FO
*****************************************************************}

function FOGetVendeurMul(CC: TControl; ZoneCom, Etablis: string): boolean;
var
  StC, Repres, NomChamp, CodeC: string;
  GG: THGrid;
begin
  Result := False;
  if CC is THGrid then
  begin
    GG := THGrid(CC);
    Repres := GG.Cells[GG.Col, GG.Row];
  end else
  begin
    Repres := THCritMaskEdit(CC).Text;
  end;
  StC := 'GCL_ZONECOM=' + ZoneCom + ';GCL_TYPECOMMERCIAL=VEN;GCL_COMMERCIAL=' + Repres
    + ';GCL_DATESUPP=' + DateToStr(Date);
  CodeC := GetInfoParPiece(FOGetNatureTicket(False, False), 'GPP_FILTRECOMM');
  NomChamp := RechDom('GCINFOCOMMERCIAL', CodeC, True);
  if NomChamp = 'GCL_ETABLISSEMENT' then
    StC := StC + ';GCL_ETABLISSEMENT=' + Etablis;

  Repres := AGLLanceFiche('GC', 'GCCOMMERCIAL_RECH', StC, '', '');
  if Repres <> '' then
  begin
    Result := True;
    if CC is THGrid then
    begin
      GG := THGrid(CC);
      GG.Cells[GG.Col, GG.Row] := Repres;
    end else
    begin
      THCritMaskEdit(CC).Text := Repres;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Recherche d'un vendeur depuis un MaskEdit
Mots clefs ... : FO
*****************************************************************}

function FOGetVendeur(CC: TControl; TitreSel, ZoneCom, Etablis: string): boolean;
begin
  if VH_GC.GCRechCommAv then Result := FOGetVendeurMul(CC, ZoneCom, Etablis)
  else Result := FOGetVendeurLookUp(CC, TitreSel, Etablis, tlLocate);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/03/2003
Modifié le ... : 10/03/2003
Description .. : Retourne le code du vendeur associé à un code
Suite ........ : utilisateur
Mots clefs ... : FO
*****************************************************************}

function FOGetVendeurFromUtilisat(Utilisateur: string): string;
var sSql: string;
  QQ: TQuery;
begin
  Result := '';
  if Utilisateur = '' then Exit;

  sSql := 'SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_TYPECOMMERCIAL="VEN"'
    + ' AND GCL_UTILASSOCIE="' + Utilisateur + '"';
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then Result := QQ.FindField('GCL_COMMERCIAL').AsString;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : Retourne le code du vendeur en fonction du paramétrage 
Suite ........ : de la caisse
Mots clefs ... : FO
*****************************************************************}

function FOGetChoixVendeur(Choix: string; TOBTiers: TOB): string;
var
  sNature, sSouche, sSql: string;
  QQ: TQuery;
begin
  Result := '';
  if FOGetParamCaisse('GPK_VENDSAISIE') <> 'X' then Exit;
  Choix := FOGetParamCaisse('GPK_VENDREPRISE');
  if Choix = 'CAI' then // reprise du vendeur par défaut de la caisse
  begin
    Result := FOGetParamCaisse('GPK_VENDDEFAUT');
  end else
    if Choix = 'UTI' then // reprise du code utilisateur
  begin
    Result := FOGetVendeurFromUtilisat(V_PGI.FUser);
  end else
    if Choix = 'CLI' then // reprise du vendeur du client
  begin
    if (TOBTiers <> nil) and (TOBTiers.FieldExists('T_REPRESENTANT')) then
    begin
      Result := TOBTiers.GetValue('T_REPRESENTANT');
    end;
  end else
    if Choix = 'TIC' then // reprise du vendeur du ticket précédent
  begin
    Result := FOGetParamCaisse('LASTVENDEUR');
    if Result = '' then
    begin
      sNature := FOGetNatureTicket(False, False);
      sSouche := GetSoucheG(sNature, VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'), '');
      sSql := 'SELECT GP_REPRESENTANT FROM PIECE WHERE GP_NATUREPIECEG="' + sNature + '"'
        + ' AND GP_NUMERO=' + IntToStr(GetNumSoucheG(sSouche) -1)
        + ' AND GP_CAISSE="' + FOGetParamCaisse('GPK_CAISSE') +'"';
      QQ := OpenSQL(sSql, True);
      if not QQ.EOF then
      begin
        Result := QQ.FindField('GP_REPRESENTANT').AsString;
        FOMajChampSupValeur(VH_GC.TOBPCaisse, 'LASTVENDEUR', Result);
      end;
      Ferme(QQ);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Initialise le vendeur de l'en-tête de la pièce
Mots clefs ... : FO
*****************************************************************}

function FOPreAffecteVendeur(TOBTiers: TOB): string;
var
  Choix: string;
begin
  Result := '';
  if FOGetParamCaisse('GPK_VENDSAISIE') <> 'X' then Exit;
  Choix := FOGetParamCaisse('GPK_VENDREPRISE');
  Result := FOGetChoixVendeur(Choix, TOBTiers);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Initialise le vendeur de la ligne
Mots clefs ... : FO
*****************************************************************}

procedure FOPreAffecteVendeurLigne(TOBPiece, TOBLigne, TOBTiers: TOB; ARow: Integer);
var Choix, Vendeur: string;
  TOBL: TOB;
  Ind: Integer;
begin
  if FOGetParamCaisse('GPK_VENDSAISLIG') <> 'X' then Exit;
  Vendeur := '';
  if (ARow <= 1) and (FOGetParamCaisse('GPK_VENDSAISIE') = '-') then
  begin // règle pour le 1er vendeur (en-tête ou 1ère ligne)
    Choix := FOGetParamCaisse('GPK_VENDREPRISE');
  end else
  begin // règle pour les lignes suivantes
    Choix := FOGetParamCaisse('GPK_VENDLIGREP');
    if Choix = 'REG' then Choix := FOGetParamCaisse('GPK_VENDREPRISE');
    if (ARow <= 1) and ((Choix = 'FST') or (Choix = 'PRE')) then Choix := 'ENT';
  end;
  if Choix = 'ENT' then // reprise du vendeur de l'en-tête
  begin
    Vendeur := TOBPiece.GetValue('GP_REPRESENTANT');
  end else
    if Choix = 'FST' then // reprise du vendeur de la 1ère ligne
  begin
    TOBL := GetTOBLigne(TOBPiece, 1);
    if TOBL <> nil then Vendeur := TOBL.GetValue('GL_REPRESENTANT');
  end else
    if Choix = 'PRE' then // reprise du vendeur de la ligne précédente
  begin
    for Ind := ARow - 1 downto 0 do
    begin
      TOBL := GetTOBLigne(TOBPiece, Ind);
      if (TOBL <> nil) and (FOEstUneLigneVente(TOBL)) then
      begin
        Vendeur := TOBL.GetValue('GL_REPRESENTANT');
        Break;
      end;
    end;
  end else
  begin
    Vendeur := FOGetChoixVendeur(Choix, TOBTiers);
  end;
  TOBLigne.PutValue('GL_REPRESENTANT', Vendeur);
end;

{==============================================================================================}
{============================== TRAITEMENT DES DEMARQUES ======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Recherche d'un code démarque sous forme d'un Lookup
Mots clefs ... : FO
*****************************************************************}

function FOGetDemarqueSaisi(CC: TControl; TitreSel: string; Letl: TTypeLocate; BFidelite: Boolean = False): Boolean;
var StRestriction: string;
begin
  if BFidelite then StRestriction := 'GTR_FERME<>"X"'
  else StRestriction := 'GTR_FERME<>"X" AND GTR_FIDELITE="-"';
  Result := LookupList(CC, TitreSel, 'TYPEREMISE', 'GTR_TYPEREMISE', 'GTR_LIBELLE', StRestriction, 'GTR_LIBELLE', False, 0, '', Letl);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/11/2002
Modifié le ... : 13/11/2002
Description .. : vérifie si la saisie du client est obligatoire pour un code
Suite ........ : démarque.
Mots clefs ... : FO
*****************************************************************}

function FODemarqueClientOblig(CodeDem: string): Boolean;
var TOBDem, TOBL: TOB;
  Ind: Integer;
begin
  Result := False;
  if FOGetParamCaisse('GPK_CLISAISIE') = '-' then Exit;
  if FOGetParamCaisse('GPK_DEMSAISIE') = '-' then Exit;
  TOBDem := FOGetPTobDemarque;
  if TOBDem = nil then Exit;
  if CodeDem = '' then
  begin
    // Recherche si le client est obligatoire pour au moins un code démarque
    for Ind := 0 to TOBDem.Detail.Count - 1 do
    begin
      TOBL := TOBDem.Detail[Ind];
      if FODecodeLibreDemarque(TOBL, '1') = 'X' then
      begin
        Result := True;
        Break;
      end;
    end;
  end else
  begin
    // Recherche si le client est obligatoire pour le code démarque
    TOBL := TOBDem.FindFirst(['GTR_TYPEREMISE'], [CodeDem], False);
    Result := (FODecodeLibreDemarque(TOBL, '1') = 'X');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/05/2003
Modifié le ... : 13/05/2003
Description .. : vérifie si la date de la pièce est dans la période d'utilisation
Suite ........ : du motif de démarque
Mots clefs ... :
*****************************************************************}

function FODemarquePeriode(CodeDem: string; DatePiece: TDateTime): boolean;
var
  TOBDem, TOBL: TOB;
  DateDeb, DateFin: TDateTime;
begin
  Result := True;
  if CodeDem = '' then Exit;
  if FOGetParamCaisse('GPK_DEMSAISIE') = '-' then Exit;
  TOBDem := FOGetPTobDemarque;
  if TOBDem = nil then Exit;
  TOBL := TOBDem.FindFirst(['GTR_TYPEREMISE'], [CodeDem], False);
  if TOBL = nil then Exit;
  DateDeb := StrToDate(FODecodeLibreDemarque(TOBL, '4'));
  DateFin := StrToDate(FODecodeLibreDemarque(TOBL, '5'));
  Result := ((DatePiece >= DateDeb) and (DatePiece <= DateFin));
end;

{==============================================================================================}
{=============================== TRAITEMENT DES CLIENTS =======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/11/2002
Modifié le ... : 13/11/2002
Description .. : vérifie si le dépassement du seuil par le CA du ticket (hors
Suite ........ : opération de caisse) rend la saisie du client est obligatoire.
Mots clefs ... : FO
*****************************************************************}

function FOSeuilCliOblig(CA: Double): Boolean;
var Maxi: Double;
begin
  Result := False;
  if FOGetParamCaisse('GPK_CLISAISIE') = '-' then Exit;
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists('GPK_SEUILCLIOBLIG')) then
  begin
    Maxi := VH_GC.TOBPCaisse.GetValue('GPK_SEUILCLIOBLIG');
    if (Maxi <> 0) and (CA >= Maxi) then Result := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/11/2002
Modifié le ... : 15/11/2002
Description .. : Retourne si la saisie du client est obligatoire ou interdite
Suite ........ : pour une opération de caisse :
Suite ........ :   'FAC' => client facultatif
Suite ........ :   'INT' => client interdit
Suite ........ :   'OBL' => client obligatoire
Mots clefs ... : FO
*****************************************************************}

function FOOpCaisseCliOblig(TOBLig, TOBArt: TOB): string;
begin
  Result := 'FAC';
  if TOBLig = nil then Exit;
  if TOBArt = nil then Exit;
  if FOGetParamCaisse('GPK_CLISAISIE') = '-' then Exit;
  if TOBLig.GetValue('GL_TYPEARTICLE') <> 'FI' then Exit;
  if not TOBArt.FieldExists('GA2_RATTACHECLI') then Exit;
  Result := TOBArt.GetValue('GA2_RATTACHECLI');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/11/2003
Modifié le ... : 20/11/2003
Description .. : Fabrique le nom du client à partir de la civilité, du nom et du
Suite ........ : prénom.
Mots clefs ... : 
*****************************************************************}

function FOMakeNomClient(TOBTiers: TOB; CodeTiers: string = ''): string;
var
  Stg: string;
  TobLocal: boolean;
  QQ: TQuery;
begin
  Result := '';
  if TOBTiers = nil then
  begin
    if (CodeTiers = '') or (CodeTiers = FODonneClientDefaut) then Exit;

    TOBTiers := TOB.Create('TIERS', nil, -1);
    Stg := 'SELECT T_TIERS,T_PARTICULIER,T_JURIDIQUE,T_PRENOM,T_LIBELLE '
      + 'FROM TIERS WHERE T_TIERS="'+ CodeTiers + '"';
    QQ := OpenSQL(Stg, True);
    if not QQ.EOF then TOBTiers.SelectDB('', QQ);
    Ferme(QQ);
    TobLocal := True;
  end else
    TobLocal := False;

  if TOBTiers.GetValue('T_TIERS') <> FODonneClientDefaut then
  begin
    if TOBTiers.GetValue('T_PARTICULIER') = 'X' then
    begin
      Stg := TOBTiers.GetValue('T_JURIDIQUE');
      if Stg <> '' then Result := Stg + ' ';
      Stg := TOBTiers.GetValue('T_PRENOM');
      if Stg <> '' then Result := Result + Stg + ' ';
      Stg := TOBTiers.GetValue('T_LIBELLE');
      if Stg <> '' then Result := Result + Stg;
    end else
      Result := TOBTiers.GetValue('T_LIBELLE');
  end;

  if (TobLocal) and (TOBTiers <> nil) then TOBTiers.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 02/12/2003
Modifié le ... : 02/12/2003
Description .. : Recherche du tiers sur le libellé
Mots clefs ... :
*****************************************************************}

function FOGetTiersMulLibelle (G_CodeTiers: THCritMaskEdit): boolean;
var
  Stg, sWhere, sCodeTiers, sNat, sCod : string;
  Ind: integer;
begin
  // Constitution de la clause where pour la recherche des tiers
  Result := False ;
  sCodeTiers := '';
  sWhere := '' ;
  Stg := 'GCTIERSSAISIE';
  Ind := TTToNum(Stg) ;
  if Ind > 0 then sWhere := V_PGI.DECombos[Ind].Where ;
  if ctxFO in V_PGI.PGIContexte then
  begin
    sNat := 'MFO';
    sCod := 'FOCLIMUL_MODE';
    sWhere := 'SELECTION' + sWhere;
  end else
  if ctxMode in V_PGI.PGIContexte then
  begin
    {$IFDEF CHR}
    sNat := 'H';
    sCod := 'HRCLIENTS_MUL';
    sWhere := 'SELECTION' + sWhere;
    {$ELSE}
    sNat := 'GC';
    sWhere := 'SELECTION';
    if Pos('FOU', sWhere)  > 0 then sCod := 'GCFOURNISSEUR_MUL' else sCod := 'GCCLIMUL_MODE';
    {$ENDIF}
  end else
  begin
    sNat := 'GC';
    sCod := 'GCTIERS_RECH';
  end;

  sCodeTiers := AGLLanceFiche(sNat, sCod, 'T_LIBELLE='+G_CodeTiers.text, '', sWhere);
  if sCodeTiers <> '' then
  begin
    G_CodeTiers.Text := sCodeTiers;
    Result := True;
  end;
end;

{==============================================================================================}
{======================== AFFICHAGE INFORMATIONS SUR LA LIGNE =================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Affichage de la quantité en stock disponible de l'article
Mots clefs ... : FO
*****************************************************************}

function FOMontreStockDispo(TOBArt, TOBLig: TOB): string;
var Depot: string;
  Qte: Double;
  TOBD: TOB;
begin
  Result := '';
  Qte := 0;
  if TOBArt = nil then Exit;
  if TOBLig = nil then Exit;
  if TOBArt.GetValue('GA_TENUESTOCK') = '-' then Exit;
  Depot := TOBLig.GetValue('GL_DEPOT');
  TOBD := TOBArt.FindFirst(['GQ_DEPOT'], [Depot], False);
  if TOBD <> nil then Qte := TOBD.GetValue('GQ_PHYSIQUE') - TOBD.GetValue('GQ_RESERVECLI');
  Result := StrS(Qte, V_PGI.OkDecQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Affichage des informations sur l'article de la ligne
Mots clefs ... : FO
*****************************************************************}

function FOMontreInfoArticle(TOBArt: TOB): string;
var Champ: string;
  ArtDim: Boolean;
begin
  Result := '';
  if TOBArt = nil then Exit;
  // Si l'article est dimensionnée
  ArtDim := (TOBArt.GetValue('GA_STATUTART') = 'DIM');
  if ArtDim then Champ := VH_GC.TOBPCaisse.GetValue('GPK_INFOARTDIM')
  else Champ := VH_GC.TOBPCaisse.GetValue('GPK_INFOARTICLE');
  if Champ = '' then Exit;
  Champ := RechDom('GCINFOLIGNE', Champ, True);
  if Champ = 'DIMENSION' then Result := FOLibelleDimension(TOBArt) else
    if Champ = 'FAMILLE' then Result := FOLibelleFamille(TOBArt) else
    if pos('GA_FAMILLENIV', Champ) > 0 then Result := FOLibelleUneFamille(TOBArt, Champ) else
    if TOBArt.FieldExists(Champ) then Result := TOBArt.GetValue(Champ);
end;

{==============================================================================================}
{================================== RECHERCHE D'UN ARTICLE ====================================}
{==============================================================================================}
{========================== Recherches et chargements =============================}
///////////////////////////////////////////////////////////////////////////////////////
//  FOGetArticleMul : recherche d'un article sous forme d'un Mul
///////////////////////////////////////////////////////////////////////////////////////

function FOGetArticleMul(CC: TControl; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string): boolean;
var sw, sWhere, CodeArt, SNatArt: string;
begin
  Result := False;
  CodeArt := '';
  if CC is THGrid then CodeArt := THGrid(CC).Cells[THGrid(CC).Col, THGrid(CC).Row] else
    if CC is THCritMaskEdit then CodeArt := THCritMaskEdit(CC).Text;
  sWhere := '';
  sW := FabricWhereNatArt(NaturePieceG, DomainePiece, SelectFourniss);
  // Restriction possible de l'affichage des articles d'un seul fournisseur
  SNatArt := GetInfoParPiece(NaturePieceG, 'GPP_TYPEARTICLE');
  while SNatArt <> '' do
  begin
    if sWhere <> '' then sWhere := sWhere + ' OR ';
    sWhere := sWhere + '(CO_CODE="' + ReadTokenSt(SNatArt) + '")';
  end;
  sWhere := 'TYPEARTICLE=' + sWhere;
  sWhere := sWhere + ';GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18));
  sWhere := sWhere + ';XX_WHERE_=' + sw;
  if SelectFourniss <> '' then sWhere := sWhere + ';EXCLUSIF;GA_FOURNPRINC=' + SelectFourniss;
  {$IFDEF CHR}
  CodeArt:=AGLLanceFiche('H', 'HRPRESTATIONS_MUL', '', '', 'SELECTION');
  {$ELSE}
  CodeArt := DispatchArtMode(1, '', '', 'SELECTION;' + sWhere);
  {$ENDIF}
  if CodeArt <> '' then
  begin
    Result := True;
    if CC is THGrid then THGrid(CC).Cells[THGrid(CC).Col, THGrid(CC).Row] := CodeArt else
      if CC is THCritMaskEdit then THCritMaskEdit(CC).Text := CodeArt;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOGetArticleLookUp : recherche d'un article sous forme d'un Lookup
///////////////////////////////////////////////////////////////////////////////////////

function FOGetArticleLookUp(CC: TControl; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string): boolean;
var sWhere, sArt: string;
begin
  sArt := 'GA_STATUTART<>"DIM"';
  sWhere := FabricWhereNatArt(NaturePieceG, DomainePiece, SelectFourniss);
  if sWhere <> '' then sArt := sArt + ' AND ' + sWhere;
  Result := LookupList(CC, TitreSel, 'ARTICLE', 'GA_ARTICLE', 'GA_LIBELLE', sArt, 'GA_ARTICLE', True, 7);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Recherche d'un article
Mots clefs ... : FO
*****************************************************************}

function FOGetArticleRecherche(CC: TControl; TitreSel, NaturePieceG: string): boolean;
begin
  if CC is THGrid then Result := GetArticleRecherche(THGrid(CC), TitreSel, NaturePieceG, '', '') else
    if VH_GC.GCRechArtAv then Result := FOGetArticleMul(CC, TitreSel, NaturePieceG, '', '')
  else Result := FOGetArticleLookUp(CC, TitreSel, NaturePieceG, '', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 15/11/2002
Modifié le ... : 15/11/2002
Description .. : Ajoute à la TOB de l'article les informations complémentaires
Suite ........ : de la table ARTICLECOMPL
Suite ........ : Pour l'instant uniquement le champ  GA2_RATTACHECLI
Suite ........ : qui concerne les articles financiers.
Mots clefs ... : FO
*****************************************************************}

procedure FOAjoutArticleCompl(TOBArt: TOB);
var QQ: TQuery;
  Stg: string;
begin
  if TOBArt = nil then Exit;
  if TOBArt.GetValue('GA_TYPEARTICLE') <> 'FI' then Exit;
  if TOBArt.FieldExists('GA2_RATTACHECLI') then Exit;
  Stg := TOBArt.GetValue('GA_ARTICLE');
  if Stg = '' then Exit;
  // Lecture de la table complément article
  QQ := OpenSQL('SELECT GA2_RATTACHECLI FROM ARTICLECOMPL WHERE GA2_ARTICLE="' + Stg + '"', True);
  if QQ.EOF then Stg := 'FAC' else Stg := QQ.FindField('GA2_RATTACHECLI').AsString;
  Ferme(QQ);
  TOBArt.AddChampSupValeur('GA2_RATTACHECLI', Stg);
end;

{==============================================================================================}
{=============================== VERIFICATION D'UN ARTICLE ====================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 14/11/2002
Description .. : Vérifie si la vente de l'article est autorisée
Mots clefs ... : FO
*****************************************************************}

function FOArticleAutorise(TOBPiece, TOBArticles: TOB; NaturePiece: string; ARow: Integer): T_FOArtAutorise;
var TOBA: TOB;
  RefUnique, sFiArt: string;
  NoLig, CodeEvt: Integer;
begin
  Result := afoOk;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  RefUnique := TOBA.GetValue('GA_ARTICLE');
  if RefUnique = '' then Exit;
  if not ArticleAutorise(TOBPiece, TOBArticles, NaturePiece, ARow) then
  begin
    Result := afoFerme; // Article fermé
    Exit;
  end;
  if CtxMode in V_PGI.PGIContexte then
  begin
    if (TOBPiece.GetValue('GP_VENTEACHAT') = 'VEN') and
      (TOBA.GetValue('GA_INTERDITVENTE') = 'X') then
    begin
      Result := afoInterdit; // Article interdit à la vente
      Exit;
    end;
  end;
  // confidentialite sur les opérations de caisse
  if TOBA.GetValue('GA_TYPEARTICLE') = 'FI' then
  begin
    sFiArt := TOBA.GetValue('GA_TYPEARTFINAN');
    if sFiArt = 'ABA' then CodeEvt := 61 // Acquisition bon achat
    else if sFiArt = 'ECD' then CodeEvt := 62 // Encaissement chèque différé
    else if sFiArt = 'ECR' then CodeEvt := 63 // Encaissement crédit
    else if sFiArt = 'END' then CodeEvt := 64 // Entrée diverse
    else if sFiArt = 'FCA' then CodeEvt := 65 // Fond de caisse
    else if sFiArt = 'PRE' then CodeEvt := 66 // Prélèvement
    else if sFiArt = 'RAR' then CodeEvt := 67 // Remboursement arrhes
    else if sFiArt = 'RAV' then CodeEvt := 68 // Remboursement avoir
    else if sFiArt = 'RBA' then CodeEvt := 69 // Remboursement bon achat
    else if sFiArt = 'SOD' then CodeEvt := 70 // Sortie Diverse
    else if sFiArt = 'VAR' then CodeEvt := 71 // Versement arrhes
    else CodeEvt := 0;
    if (CodeEvt > 0) and (not FOJaiLeDroit(CodeEvt)) then
    begin
      Result := afoConfidentiel; // Article confidentiel
      Exit;
    end;
  end;
  // pour ne pas mixer certains opérations financières avec des articles
  if not FOVerifMixageArticle(TOBPiece, TOBArticles, NoLig) then
  begin
    Result := afoInCompatible; // Article incompatible avec les autres lignes
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/11/2002
Modifié le ... : 14/11/2002
Description .. : Interdit de mixer certaines opérations financières avec des
Suite ........ : marchandises ou des prestations.
Mots clefs ... : FO
*****************************************************************}

function FOVerifMixageArticle(TOBPiece, TOBArticles: TOB; var ARow: Integer): Boolean;
var TOBL, TOBA: TOB;
  RefUnique, sFiArt, sFiPiece, sTypeOp: string;
  Ind, NbLignes: Integer;
begin
  Result := True;
  sFiPiece := '';
  NbLignes := 0;
  sTypeOp := TYPEOPCFDCAISSE + ';' + TYPEOPCPRELEV + ';' + TYPEOPCENTREE
    + ';' + TYPEOPCSORTIE + ';' +  TYPEOPCECART + ';';
  for Ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[Ind];
    if (TOBL.GetValue('GL_TYPELIGNE') = 'ART') or (Ind + 1 = ARow) then
    begin
      RefUnique := TOBL.GetValue('GL_ARTICLE');
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
      sFiArt := '';
      if (TOBA <> nil) and (TOBA.GetValue('GA_TYPEARTICLE') = 'FI') then
      begin
        sFiArt := TOBA.GetValue('GA_TYPEARTFINAN');
        if not FOStrCmp(sFiArt, sTypeOp) then sFiArt := '';
      end;
      if sFiArt = '' then Inc(NbLignes);
      if sFiPiece = '' then sFiPiece := sFiArt;
      if (sFiPiece <> '') and ((sFiPiece <> sFiArt) or (NbLignes > 0)) then
      begin
        Result := False; // Article incompatible avec les autres lignes
        ARow := Ind;
        Exit;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/01/2003
Modifié le ... : 14/01/2003
Description .. : Recherche du signe de la quantité en fonction de l'article
Mots clefs ... : FO
*****************************************************************}

function FOSigneQteArticle(TOBPiece, TOBArticles: TOB; ARow: integer; Qte: double): double;
var TOBA: TOB;
  RefUnique, sFiArt: string;
begin
  Result := Qte;
  if TobPiece = nil then Exit;
  if TOBArticles = nil then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  RefUnique := TOBA.GetValue('GA_ARTICLE');
  if RefUnique = '' then Exit;
  if TOBA.GetValue('GA_TYPEARTICLE') = 'FI' then
  begin
    sFiArt := TOBA.GetValue('GA_TYPEARTFINAN');
    if FOStrCmp(sFiArt, 'PRE;RAR;RAV;RBA;SOD;') then
      Result := Arrondi((Qte * (-1)), V_PGI.OkDecQ); // on propose une quantité négative
  end;
end;

{==============================================================================================}
{============================== INFORMATIONS SUR UN ARTICLE ===================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : retourne le nom et le code de chacune des dimensions d'un
Suite ........ : article, par exemple :  Taille : 40 - Couleur : Blanc
Mots clefs ... : FO
*****************************************************************}

function FOLibelleDimension(TOBArt: TOB; Abrege: Boolean = True; Compact: Boolean = False; AvecTitre: Boolean = True; Separateur: string = '-:'): string;
var
  LibelleDim, GrilleDim, CodeDim, CodeGrille, SepT, SepP: string;
  Ind: Integer;
begin
  Result := '';
  if TOBArt = nil then Exit;
  SepT := Copy(Separateur, 1, 1);
  if SepT = '' then SepT := '-';
  SepP := Copy(Separateur, 2, 1);
  if SepP = '' then SepP := ':';
  if not Compact then
  begin
    SepT := ' '+ SepT +' ';
    SepP := ' '+ SepP +' ';
  end;
  LibelleDimensions('', TOBArt, LibelleDim, GrilleDim, Abrege);
  for Ind := 1 to MaxDimension do
  begin
    CodeDim := TraduireMemoire(Trim(ReadTokenSt(LibelleDim)));
    CodeGrille := TraduireMemoire(Trim(ReadTokenSt(GrilleDim)));
    if CodeDim <> '' then
    begin
      if Result <> '' then Result := Result + SepT;
      if AvecTitre then
        Result := Result + CodeDim + SepP + CodeGrille
      else
        Result := Result + CodeGrille;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : retourne le nom et le libellé de chacune des familles d'un
Suite ........ : article
Suite ........ : par exemple :  Rayon : Homme - Famille : Accessoire -
Suite ........ : Sous-Famille : Cravate
Mots clefs ... : FO
*****************************************************************}

function FOLibelleFamille(TOBArt: TOB; Abrege: Boolean = True): string;
var Table, Champ, CodeFam: string;
  Ind: Integer;
begin
  Result := '';
  if TOBArt = nil then Exit;
  for Ind := 1 to 3 do
  begin
    Champ := 'GA_FAMILLENIV' + IntToStr(Ind);
    CodeFam := TOBArt.GetValue(Champ);
    if CodeFam <> '' then
    begin
      if Ind > 1 then Result := Result + ' - ';
      Table := 'GCFAMILLENIV' + IntToStr(Ind);
      Champ := 'LF' + IntToStr(Ind);
      Result := Result + TraduireMemoire(RechDom('GCLIBFAMILLE', Champ, Abrege))
        + ' : ' + TraduireMemoire(RechDom(Table, CodeFam, Abrege));
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne le libellé d'une famille d'un article
Mots clefs ... : FO
*****************************************************************}

function FOLibelleUneFamille(TOBArt: TOB; Champ: string; Abrege: Boolean = False): string;
var Table, CodeFam: string;
begin
  Result := '';
  if TOBArt = nil then Exit;
  CodeFam := TOBArt.GetValue(Champ);
  if CodeFam <> '' then
  begin
    Table := 'GCFAMILLENIV' + Champ[Length(Champ)];
    Result := RechDom(Table, CodeFam, Abrege);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Affiche la photo d'un article
Mots clefs ... : FO
*****************************************************************}

function FOAffichePhotoArticle(CodeArt: string; Image: TImage): Boolean;
var QQ: TQuery;
  Stg: string;
  IsJpeg: Boolean;
  //Rec    : TRect ;
  XX, YY: Integer;
  Icon: TIcon;
begin
  Result := False;
  if Image = nil then Exit;
  Image.Picture := nil;
  ///if VH_GC.GCPHOTOFICHE = '' then  Exit ;
  // recherche de l'image de l'article générique
  if (CodeArt <> '') and (VH_GC.GCPHOTOFICHE <> '') then
  begin
    Stg := 'SELECT LO_QUALIFIANTBLOB, LO_OBJET from LIENSOLE '
      + 'where LO_TABLEBLOB="GA" AND LO_IDENTIFIANT="' + FOIdentArticleGen(CodeArt) + '" '
      + 'AND (LO_QUALIFIANTBLOB="PHO" OR  LO_QUALIFIANTBLOB="PHJ" OR  LO_QUALIFIANTBLOB="VIJ") '
      + 'AND LO_EMPLOIBLOB="' + VH_GC.GCPHOTOFICHE + '"';
    QQ := OpenSQL(Stg, True);
    if not QQ.EOF then
    begin
      // chargement de l'image dans le contrôle
      IsJpeg := ((QQ.Findfield('LO_QUALIFIANTBLOB').asString = 'PHJ') or (QQ.Findfield('LO_QUALIFIANTBLOB').asString = 'VIJ'));
      LoadBitMapFromChamp(QQ, 'LO_OBJET', Image, IsJpeg);
      if Image.Picture <> nil then Result := True;
    end;
    Ferme(QQ);
  end;
  // si pas de photo
  if not Result then
  begin
    // chargement du logo société
    Stg := FOLogoCaisse;
    if (Stg <> '') and (FileExists(Stg)) then
    try
      Image.Picture.LoadFromFile(Stg);
      SetJPEGOptions(Image);
      if Image.Picture <> nil then Result := True;
    finally
    end;
  end;
  // si pas de Logo société
  if not Result then
  begin
    // Icône PGI
    Icon := TIcon.Create;
    try
      Icon.Handle := Loadicon(AGLHinstance, PChar('ZPGI')) ;
      XX := (Image.Width - Icon.Width) div 2;
      if XX < 0 then XX := 0;
      YY := (Image.Height - Icon.Height) div 2;
      if YY < 0 then YY := 0;
      if Icon.Handle <> 0 then Image.Canvas.Draw(XX, YY, Icon);
    finally
      Icon.Free;
    end;
    {**
       // Cadre avec le texte Photo
       Rec := Image.ClientRect ;
       Image.Canvas.Font.Style := [fsBold] ;
       Stg := TraduireMemoire('PHOTO') ;
       XX := (Image.Width - Image.Canvas.TextWidth(Stg)) div 2 ;
       if XX < 0 then XX := 0 ;
       Image.Canvas.TextRect(Rec, XX, 10, Stg) ;
       Image.Canvas.Brush.Color := clBlack ;
       Image.Canvas.FrameRect(Rec) ;
    {**
       // cadre et diagonales sous forme de tiret
       Rec := Image.ClientRect ;
       Image.Canvas.Pen.Style := psDot;
       Image.Canvas.Rectangle(Rec) ;
       Image.Canvas.MoveTo(2, 2) ;
       Image.Canvas.LineTo(Image.Width -2, Image.Height -2) ;
       Image.Canvas.MoveTo(Image.Width -2, 2) ;
       Image.Canvas.LineTo(2, Image.Height -2) ;
    **}
  end;
end;

{==============================================================================================}
{==================================== GESTION DES PIECES ======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Attribue le numéro à une pièce à partir de la souche et
Suite ........ : vérifie qu'aucune pièce avec ce numéro n'existe
Mots clefs ... : FO
*****************************************************************}

function FOSetNumeroDefinitif(PieceContainer: TPieceContainer): Boolean;
var NCleDoc: R_CleDoc;
  Stg: string;
  SetNum, CreeTBas, CreeTEch, CreeTNom, CreeTAcpt: Boolean;
begin
  Result := False;
  SetNum := True;
  if PieceContainer.TCPiece = nil then Exit;
  // Création si besoin des TOB complémentaires
  CreeTBas := False;
  if PieceContainer.TCBases = nil then
  begin
    PieceContainer.TCBases := TOB.Create('BASES', nil, -1);
    CreeTBas := True;
  end;
  CreeTEch := False;
  if PieceContainer.TCEches = nil then
  begin
    PieceContainer.TCEches := TOB.Create('Les ECHEANCES', nil, -1);
    CreeTEch := True;
  end;
  CreeTNom := False;
  if PieceContainer.TCNomenclature = nil then
  begin
    PieceContainer.TCNomenclature := TOB.Create('', nil, -1);
    CreeTNom := True;
  end;
  CreeTAcpt := False;
  if PieceContainer.TCAcomptes = nil then
  begin
    PieceContainer.TCAcomptes := TOB.Create('NOMENCLATURES', nil, -1);
    CreeTAcpt := True;
  end;
  // Recherche du numéro dans la souche
  while SetNum do
  begin
    if SetNumeroDefinitif(PieceContainer) then
    begin
      // Vérifie si le numéro n'est pas déjà attribué
      NCleDoc := TOB2CleDoc(PieceContainer.TCPiece);
      Stg := 'SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="' + NCleDoc.NaturePiece + '"'
        + ' AND GP_SOUCHE="' + NCleDoc.Souche + '" AND GP_NUMERO=' + IntToStr(NCleDoc.NumeroPiece)
        + ' AND GP_INDICEG=' + IntToStr(NCleDoc.Indice) + ' ';
      if not ExisteSQL(Stg) then
      begin
        SetNum := False;
        Result := True;
      end;
    end else
    begin
      SetNum := False;
    end;
  end;
  // Libére les TOB créées localement
  if CreeTBas then PieceContainer.TCBases.Free;
  if CreeTEch then PieceContainer.TCEches.Free;
  if CreeTNom then PieceContainer.TCNomenclature.Free;
  if CreeTAcpt then PieceContainer.TCAcomptes.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si la ligne de la pièce est une ligne de vente d'un
Suite ........ : article
Mots clefs ... : FO
*****************************************************************}

function FOEstUneLigneVente(TOBLigne: TOB): Boolean;
const TypeLigne: string = 'TOT;COM';
begin
  Result := False;
  if TOBLigne = nil then Exit;
  Result := ((TOBLigne.GetValue('GL_ARTICLE') <> '') and (not FOStrCmp(TOBLigne.GetValue('GL_TYPELIGNE'), TypeLigne)));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si la pièce a au moins une ligne article
Mots clefs ... : FO
*****************************************************************}

function FOPieceUneLigne(TOBPiece: TOB): Boolean;
var Ind: Integer;
begin
  Result := False;
  if TOBPiece = nil then Exit;
  for Ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    if FOEstUneLigneVente(TOBPiece.Detail[Ind]) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Vérifie si une ligne est remisable
Mots clefs ... : FO
*****************************************************************}

function FOLigneRemisable(TOBLigne: TOB; EnPied: Boolean): Boolean;
begin
  Result := False;
  if TOBLigne = nil then Exit;
  // Remisable en ligne
  Result := (TOBLigne.GetValue('GL_REMISABLELIGNE') = 'X');
  // Remisable en pied
  if (EnPied) and (Result) then Result := (TOBLigne.GetValue('GL_REMISABLEPIED') = 'X');
  // Si une remise ligne a été consentie, non remisable en pied
  if (EnPied) and (Result) and (VH_GC.TOBPCaisse.FieldExists('GPK_REMPIEDLIG')) and
    (VH_GC.TOBPCaisse.GetValue('GPK_REMPIEDLIG') = '-') then
  begin
    Result := (TOBLigne.GetValue('GL_REMISELIGNE') = 0);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Positionne l'indicateur facture HT de la pièce selon le
Suite ........ : paramétrage de la caisse
Mots clefs ... : FO
*****************************************************************}

procedure FOAffecteFactureHT(TOBPiece: TOB);
var Stg: string;
begin
  if VH_GC.TOBPCaisse <> nil then
  begin
    Stg := VH_GC.TOBPCaisse.GetValue('GPK_APPELPRIXTIC');
    if Stg = 'HT' then TOBPiece.PutValue('GP_FACTUREHT', 'X') else
      if Stg = 'TTC' then TOBPiece.PutValue('GP_FACTUREHT', '-');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2002
Modifié le ... : 28/10/2002
Description .. : Fabrique un commentaire pour une pièce
Mots clefs ... : FO
*****************************************************************}

function FOFabriqueCommentairePiece(TOBPiece: TOB; Texte: string): string;
var sComment: string;
begin
  Result := '';
  if TOBPiece = nil then Exit;
  sComment := Trim(TraduireMemoire(Texte))
    + ' ' + TraduireMemoire(RechDom('GCNATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'), False))
    + ' ' + TraduireMemoire('N°') + ' ' + IntToStr(TOBPiece.GetValue('GP_NUMERO'))
    + ' ' + TraduireMemoire('du') + ' ' + DateToStr(TOBPiece.GetValue('GP_DATEPIECE'))
    + ' ' + TraduireMemoire('caisse') + ' ' + TOBPiece.GetValue('GP_CAISSE')
    + '  ' + TOBPiece.GetValue('GP_REFINTERNE');
  Result := Trim(sComment);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2002
Modifié le ... : 01/08/2002
Description .. : Ajout d'un commentaire dans une pièce
Mots clefs ... : FO
*****************************************************************}

procedure FOInsereLigneCommentaire(TOBPiece: TOB; Texte: string; EnFin: Boolean);
var TOBL: TOB;
  sComment: string;
  Ind: Integer;
  LigneExiste: Boolean;
begin
  if TOBPiece = nil then Exit;
  sComment := Trim(TraduireMemoire(Texte))
    + ' '
    + TraduireMemoire(RechDom('GCNATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'), False))
    + ' ' + TraduireMemoire('N°') + ' ';
  // Pour éviter l'empilement des commentaires
  LigneExiste := False;
  TOBL := nil;
  if EnFin then Ind := TOBPiece.Detail.Count - 1
  else Ind := 0;
  if Ind < TOBPiece.Detail.Count then
  begin
    TOBL := TOBPiece.Detail[Ind];
    if (TOBL.GetValue('GL_TYPELIGNE') = 'COM') and
      (TOBL.GetValue('GL_TYPEDIM') = 'NOR') and
      (TOBL.GetValue('GL_CREERPAR') = 'GEN') then
    begin
      if Copy(TOBL.GetValue('GL_LIBELLE'), 0, Length(sComment)) = sComment then
        LigneExiste := True;
    end;
  end;
  // Création de la ligne
  if (not LigneExiste) or (TOBL = nil) then
  begin
    if EnFin then Ind := TOBPiece.Detail.Count + 1
    else Ind := 1;
    TOBL := NewTOBLigne(TOBPiece, Ind);
    InitialiseTOBLigne(TOBPiece, nil, nil, Ind);
    TOBL.PutValue('GL_TYPELIGNE', 'COM');
    TOBL.PutValue('GL_TYPEDIM', 'NOR');
    TOBL.PutValue('GL_CREERPAR', 'GEN');
  end;
  // Constitution du commentaire
  sComment := sComment + IntToStr(TOBPiece.GetValue('GP_NUMERO'))
    + ' ' + TraduireMemoire('du') + ' '
    + DateToStr(TOBPiece.GetValue('GP_DATEPIECE'))
    + '  ' + TOBPiece.GetValue('GP_REFINTERNE');
  sComment := Trim(sComment);
  Ind := ChampToLength('GL_LIBELLE');
  if Ind > 0 then sComment := Copy(sComment, 1, Ind);
  TOBL.PutValue('GL_LIBELLE', sComment);
  // Renumérotation des lignes
  if (not EnFin) and (not LigneExiste) then NumeroteLignesGC(nil, TOBPiece);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 06/01/2003
Modifié le ... : 06/01/2003
Description .. : Vérifie si la quantité est valide pour la saisie d'un ticket
Mots clefs ... : FO
*****************************************************************}

function FOQuantiteAutorise(Qte: double): boolean;
{$IFDEF MODE}
var QteMax: double;
  {$ENDIF}
begin
  Result := True;
  {$IFDEF MODE}
  QteMax := Arrondi(FOGetParamCaisse('QTEMAX'), V_PGI.OkDecQ);
  Qte := Arrondi(Abs(Qte), V_PGI.OkDecQ);
  if QteMax > 0 then Result := (Qte <= QteMax);
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 06/01/2003
Modifié le ... : 06/01/2003
Description .. : Vérifie si le prix unitaire est valide pour la saisie d'un ticket
Mots clefs ... : FO
*****************************************************************}

function FOPrixUnitaireAutorise(PU: double): boolean;
{$IFDEF MODE}
var PrixMax: double;
  {$ENDIF}
begin
  Result := True;
  {$IFDEF MODE}
  PrixMax := Arrondi(FOGetParamCaisse('PRIXMAX'), V_PGI.OkDecV);
  PU := Arrondi(Abs(PU), V_PGI.OkDecV);
  if PrixMax > 0 then Result := (PU <= PrixMax);
  {$ENDIF}
end;

{==============================================================================================}
{================================= CALCULS D'UNE PIECE ========================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Calcul du montant d'une ligne après modification de la
Suite ........ : quantité
Mots clefs ... : FO
*****************************************************************}

procedure FOMajQte(TOBLigne: TOB; EnHT: Boolean);
var QteF, PQ, PUDEV, PUNETDEV, MontantNetDev, MontantBrutDev, RemLigneDev: Double;
begin
  if TOBLigne.GetValue('GL_ARTICLE') = '' then Exit;
  QteF := TOBLigne.GetValue('GL_QTEFACT');
  if QteF = 0 then Exit;
  if EnHT then
  begin
    PUDEV := TOBLigne.GetValue('GL_PUHTDEV');
    PUNETDEV := TOBLigne.GetValue('GL_PUHTNETDEV');
  end else
  begin
    PUDEV := TOBLigne.GetValue('GL_PUTTCDEV');
    PUNETDEV := TOBLigne.GetValue('GL_PUTTCNETDEV');
  end;
  if PUDEV = 0 then Exit;
  if PUNETDEV = 0 then Exit;
  PQ := TOBLigne.GetValue('GL_PRIXPOURQTE');
  if PQ <= 0 then
  begin
    PQ := 1.0;
    TOBLigne.PutValue('GL_PRIXPOURQTE', PQ);
  end;
  MontantBrutDev := Arrondi((PUDEV * QteF / PQ), 6);
  MontantNetDev := Arrondi((PUNETDEV * QteF / PQ), 6);
  RemLigneDev := MontantBrutDev - MontantNetDev;
  if RemLigneDev <> 0 then TOBLigne.PutValue('GL_VALEURREMDEV', RemLigneDev);
end;

{==============================================================================================}
{================================= IMPRESSION DES PIECES ======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 02/04/2002
Description .. : Impression d'une pièce selon au format Ticket.
Mots clefs ... : FO
*****************************************************************}

procedure FOImprimeTexte(TOBPiece: TOB; CleDoc: R_CleDoc; Nature, Modele, Langue: string; Init_Imprimante, PreView: Boolean);
{$IFDEF FOS5}
var Where, Imprimante, Port, Param: string;
  SQLBased, VersionDemo: Boolean;
  Err: TMC_Err;
begin
  Fillchar(Err, Sizeof(Err), #0);
  if (trim(Modele) = '') then Modele := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODTIC');
  Where := WherePiece(CleDoc, ttdPiece, False);
  FODonneParamLP(Imprimante, Port, Param);
  // Pour imprimer "Version de démonstration" dans la base Formation
  VersionDemo := V_PGI.VersionDemo;
  if GetParamSoc('SO_GCFOBASEFORMATION') then V_PGI.VersionDemo := True;
  SQLBased := (TOBPiece = nil);
  if (not FOImprimeLP('T', Nature, Modele, Where, Imprimante, Port, Param, Init_Imprimante, PreView, SQLBased, TOBPiece, Err))
    and (Err.COde <> 0) then FOMCErr(Err, 'Impression du ticket.');
  V_PGI.VersionDemo := VersionDemo;
end;
{$ELSE}
begin
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Impression d'une pièce selon le format (Ticket, Etat,
Suite ........ : Document) défini dans le paramétrage de la caisse
Mots clefs ... : FO
*****************************************************************}

procedure FOImprimerLaPiece(TOBPiece, TOBTiers, TOBEches, TOBComms: TOB; CleDoc: R_CleDoc;
  MiseAJour: boolean = True; PreViewForce: boolean = False;
  OuvrirTiroir: boolean = False; ChoixModele: boolean = False);
{$IFDEF FOS5}
var TL: TList;
  TT: TStrings;
  PreView, EnPlus: boolean;
  TOBB, TOBE, TOBP, TOBC: TOB;
  Numero, Indice, NbCopies, DefaultDocCopies, Ind: integer;
  Format, Modele, Nature, Langue, SQL, NaturePiece, Souche, Stg: string;
  ///////////////////////////////////////////////////////////////////////////////////////
  procedure FormatEnNature;
  begin
    Nature := '';
    if Format = 'T' then Nature := NATUREMODTICKET else
      if Format = 'E' then Nature := NATUREMODTICETAT else
      if Format = 'L' then Nature := NATUREMODTICDOC;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
begin
  TOBP := nil;
  NaturePiece := '';
  // Mise à jour éventuelle de l'indicateur pièce vivante
  if (TOBPiece <> nil) and (MiseAJour) then
  begin
    NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
    if GetInfoParPiece(NaturePiece, 'GPP_ACTIONFINI') = 'IMP' then
      PutValueDetail(TOBPiece, 'GP_VIVANTE', '-');
  end;
  // Recherche du format de l'impression, de la nature et du modèle à utiliser
  Format := '';
  Modele := '';
  NbCopies := 0;
  EnPlus := False;
  if not ChoixModele then
  begin
    if (TOBPiece <> nil) and (TOBPiece.FieldExists('IMPFORMAT')) and
      (TOBPiece.FieldExists('IMPMODELE')) then
    begin
      Format := TOBPiece.GetValue('IMPFORMAT');
      Modele := TOBPiece.GetValue('IMPMODELE');
      if TOBPiece.FieldExists('NBEXEMPLAIRES') then
        NbCopies := TOBPiece.GetValue('NBEXEMPLAIRES')
      else
        NbCopies := VH_GC.TOBPCaisse.GetValue('GPK_NBEXEMPTIC');
      if TOBPiece.FieldExists('PLUSMODIMP') then
        EnPlus := (TOBPiece.GetValue('PLUSMODIMP') = 'X');
    end else
    begin
      Format := VH_GC.TOBPCaisse.GetValue('GPK_IMPFMTTIC');
      Modele := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODTIC');
      NbCopies := VH_GC.TOBPCaisse.GetValue('GPK_NBEXEMPTIC');
    end;
  end;
  if (ChoixModele) or ((NaturePiece <> '') and
    (GetInfoParPiece(NaturePiece, 'GPP_VALMODELE') = 'X')) then
  begin
    if not FOChoixModeleImpression(Format, Modele, EnPlus, NbCopies) then
    begin
      // en réimpression si aucun modèle n'est choisi
      if ChoixModele then Exit;
    end;
  end;
  if (NbCopies < 1) or (NbCopies > 9) then NbCopies := 1;
  FormatEnNature;
  Langue := V_PGI.LanguePerso;
  if Trim(Modele) = '' then Exit;
  if not ExisteModele(Format, Nature, Modele, Langue) then     // TRUE
  begin
    Stg := TraduireMemoire('Le modèle de ') + RechDom('GCFORMATTICKETFO', Format, False)
      + ' "' + Modele + '" ' + TraduireMemoire('n''est pas défini') + '#10'
      + TraduireMemoire('pour la nature ') + Nature + '-'
      + RechDom('YYNATUREETAT', Nature, False);
    PGIBox(Stg, 'Impression du ticket.');
    Exit;
  end;
  // Aperçu avant impression
  if PreViewForce then
    PreView := True
  else
    PreView := (not FOExisteLP);
  if not (PreView) and not (V_PGI.NoPrintDialog) then
  begin
    if NaturePiece <> '' then
      PreView := (GetInfoParPiece(NaturePiece, 'GPP_APERCUAVIMP') = 'X');
  end;
  // Traitement
  if (OuvrirTiroir) and (not PreView) and
    (FOExisteTiroir) and (FOVerifOuvreTiroir(TobEches)) then
  begin
    FOOuvretiroir;
  end;
  DefaultDocCopies := V_PGI.DefaultDocCopies;
  if EnPlus then Ind := 2 else Ind := 1;
  while Ind > 0 do
  begin
    V_PGI.DefaultDocCopies := NbCopies;
    if Format = 'T' then
    begin
      if TOBPiece <> nil then
      begin
        TOBP := TOB.Create('La pièce', nil, -1);
        TOBP.Dupliquer(TOBPiece, True, True, True);
        // ajout du client
        FOMergeTob(TOBP, TOBTiers, False);
        if TOBComms <> nil then
        begin
          // ajout du vendeur
          TOBC := TOBComms.FindFirst(['GCL_COMMERCIAL'], [TOBPiece.GetValue('GP_REPRESENTANT')], False);
          if TOBC <> nil then FOMergeTob(TOBP, TOBC, False);
        end;
      end;
      FOImprimeTexte(TOBP, CleDoc, Nature, Modele, Langue, True, PreView);
      if TOBP <> nil then TOBP.Free;
    end else
    begin
      TL := TList.Create;
      TT := TStringList.Create;
      SQL := RecupSQLModele(Format, Nature, Modele, '', '', '', ' WHERE ' + WherePiece(CleDoc, ttdPiece, False));
      TT.Add(SQL);
      TL.Add(TT);
      if Format = 'E' then
        LanceEtat('E', Nature, Modele, True, False, False, nil, Trim(SQL), '', False)
      else
        LanceDocument('L', Nature, Modele, TL, nil, True, False);
      TT.Free;
      TL.Free;
    end;
    // impression du modèle par défaut de la caisse
    if EnPlus then
    begin
      EnPlus := False;
      Format := VH_GC.TOBPCaisse.GetValue('GPK_IMPFMTTIC');
      FormatEnNature;
      Modele := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODTIC');
      NbCopies := VH_GC.TOBPCaisse.GetValue('GPK_NBEXEMPTIC');
      if not PreView then Delay(3000);
    end;
    Dec(Ind);
  end;
  V_PGI.DefaultDocCopies := DefaultDocCopies;
  // Libération de la TOB pied de document
  if TOBPiece <> nil then
  begin
    NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
    Souche := TOBPiece.GetValue('GP_SOUCHE');
    Numero := TOBPiece.GetValue('GP_NUMERO');
    Indice := TOBPiece.GetValue('GP_INDICEG');
    TOBB := VH_GC.TOBGCB.FindFirst(['GPB_NATUREPIECEG', 'GPB_SOUCHE', 'GPB_NUMERO', 'GPB_INDICEG'],
      [NaturePiece, Souche, Numero, Indice], True);
    if TOBB <> nil then TOBB.Parent.Free;
    TOBE := VH_GC.TOBGCE.FindFirst(['GPE_NATUREPIECEG', 'GPE_SOUCHE', 'GPE_NUMERO', 'GPE_INDICEG'],
      [NaturePiece, Souche, Numero, Indice], True);
    if TOBE <> nil then TOBE.Parent.Free;
  end;
end;
{$ELSE}
begin
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/09/2002
Modifié le ... : 26/09/2002
Description .. : Choix du modèle d'impression d'un ticket
Mots clefs ... : FO
*****************************************************************}

function FOChoixModeleImpression(var Format, Modele: string; var EnPlus: Boolean; var NbCopies: Integer): Boolean;
var Retour, Stg: string;
  TOBImp: TOB;
begin
  Result := False;
  if EnPlus then Stg := 'X' else Stg := '-';
  if NbCopies = 0 then NbCopies := V_PGI.DefaultDocCopies;
  if (Format = '') and (Modele = '') then
  begin
    Format := GetSynRegKey('FFOIMPFORMAT', Format, True);
    Modele := GetSynRegKey('FFOIMPMODELE', Format, True);
    NbCopies := GetSynRegKey('FFONBEXEMPLAIRES', NbCopies, True);
    Stg := GetSynRegKey('FFOPLUSMODIMP', Stg, True);
  end;
  if (Format = '') and (Modele = '') then
  begin
    Format := VH_GC.TOBPCaisse.GetValue('GPK_IMPFMTTIC');
    Modele := VH_GC.TOBPCaisse.GetValue('GPK_IMPMODTIC');
    NbCopies := VH_GC.TOBPCaisse.GetValue('GPK_NBEXEMPTIC');
    Stg := 'X';
  end;
  TOBImp := TOB.Create('Modele du ticket', nil, -1);
  TOBImp.AddChampSupValeur('IMPFORMAT', Format);
  TOBImp.AddChampSupValeur('IMPMODELE', Modele);
  TOBImp.AddChampSupValeur('NBEXEMPLAIRES', NbCopies);
  TOBImp.AddChampSupValeur('PLUSMODIMP', Stg);
  TheTOB := TOBImp;
  Retour := AglLanceFiche('MFO', 'SELMODELEIMP', '', '', '');
  if (Retour <> '') and (TOBImp <> nil) then
  begin
    Format := TOBImp.GetValue('IMPFORMAT');
    SaveSynRegKey('FFOIMPFORMAT', Format, True);
    Modele := TOBImp.GetValue('IMPMODELE');
    SaveSynRegKey('FFOIMPMODELE', Modele, True);
    NbCopies := TOBImp.GetValue('NBEXEMPLAIRES');
    SaveSynRegKey('FFONBEXEMPLAIRES', NbCopies, True);
    Stg := TOBImp.GetValue('PLUSMODIMP');
    SaveSynRegKey('FFOPLUSMODIMP', Stg, True);
    EnPlus := (Stg = 'X');
    Result := True;
  end;
  if TheTob <> nil then TheTob := nil;
  if TOBImp <> nil then TOBImp.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Impression d'un ticket à partir d'une pièce
Mots clefs ... : FO
*****************************************************************}

function FOImprimeTicket(TOBPiece: TOB; CleDoc: R_CleDoc): Boolean;
var TypeEtat: TTypeEtatFO;
  DefaultDocCopies: Integer;
begin
  Result := False;
  if CleDoc.NaturePiece = FOGetNatureTicket(False, False) then TypeEtat := efoTicket else
    if CleDoc.NaturePiece = 'TEM' then TypeEtat := efoTransfert else
    if CleDoc.NaturePiece = 'TRE' then TypeEtat := efoTransfert else
    if CleDoc.NaturePiece = 'FCF' then TypeEtat := efoCommande else
    if CleDoc.NaturePiece = 'BLF' then TypeEtat := efoReception else Exit;
  Result := FOExisteCodeEtat(TypeEtat);
  if Result then
  begin
    DefaultDocCopies := V_PGI.DefaultDocCopies;
    if TypeEtat = efoTicket then V_PGI.DefaultDocCopies := FONbExemplaire('TIC')
    else V_PGI.DefaultDocCopies := FONbExemplaire('BON');
    FOLanceImprimeLP(TypeEtat, '', False, TOBPiece);
    V_PGI.DefaultDocCopies := DefaultDocCopies;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Réimpression du dernier ticket de vente enregistré sur la
Suite ........ : caisse
Mots clefs ... : FO
*****************************************************************}

procedure FOReImpTicket(ChoixModele: boolean);
{$IFDEF FOS5}
var CleDoc: R_CleDoc;
  {$ENDIF}
begin
  {$IFDEF FOS5}
  // Recherche du dernier n° de ticket de la caisse
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := FOGetNatureTicket(False, False) ;
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'), '');
  CleDoc.NumeroPiece := VH_GC.TOBPCaisse.GetValue('LASTNUMTIC');
  if CleDoc.NumeroPiece = 0 then CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche) - 1;
  CleDoc.Indice := 0;
  // Impression de la pièce
  if CleDoc.NumeroPiece > 0 then
  begin
    if not FOJaiLeDroit(21) then Exit;
    if ChoixModele then ChoixModele := FOJaiLeDroit(22, False, False);
    FOImprimerLaPiece(nil, nil, nil, nil, CleDoc, False, False, False, ChoixModele);
  end else
  begin
    PGIBox('Aucun ticket n''est défini.', 'Réimpression du ticket.');
  end;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/07/2003
Modifié le ... : 31/07/2003
Description .. : Charge en TOB les articles utilisés dans une pièce
Mots clefs ... : FO
*****************************************************************}

procedure FOLoadTobArt(TOBPiece, TOBArticles: TOB);
const
  NbArtParRequete: integer = 200;
var
  Ind, CountStArt, NbRequete: integer;
  StSelect, StArticle, StCodeArticle, RefGen, Statut: string;
  TabWhere: array of string;
  QArticle: TQuery;
  OkTab: Boolean;
begin
  StSelect := 'SELECT * FROM ARTICLE';
  CountStArt := 0;
  NbRequete := 0;
  SetLength(TabWhere, 1);
  for Ind := 0 to TOBPiece.Detail.Count -1 do
  begin
    StArticle := TOBPiece.Detail[Ind].GetValue('GL_ARTICLE');
    StCodeArticle := TOBPiece.Detail[Ind].GetValue('GL_CODEARTICLE');
    RefGen := TOBPiece.Detail[Ind].GetValue('GL_CODESDIM');
    Statut := TOBPiece.Detail[Ind].GetValue('GL_TYPEDIM');
    OkTab := False;
    if CountStArt >= NbArtParRequete then
    begin
      NbRequete := NbRequete + 1;
      SetLength(TabWhere, NbRequete + 1);
      CountStArt := 0;
    end;
    if (Statut = 'GEN') or (Statut = 'DIM') or (Statut = 'NOR') then
    begin
      if (StArticle = '') and (RefGen <> '') then RefGen := CodeArticleUnique2(RefGen, '');
      if StArticle <> '' then RefGen := StArticle;
      if TabWhere[NbRequete] = '' then TabWhere[NbRequete] := '"' + RefGen + '"'
      else TabWhere[NbRequete] := TabWhere[NbRequete] + ',"' + RefGen + '"';
      OkTab := True;
    end;
    if OkTab then CountStArt := CountStArt + 1;
  end;
  if TabWhere[0] <> '' then
  begin
    for Ind := Low(TabWhere) to High(TabWhere) do
    begin
      if TabWhere[Ind] <> '' then
      begin
        QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (' + TabWhere[Ind] + ')', True);
        if not QArticle.EOF then TOBArticles.LoadDetailDB('ARTICLE', '', '', QArticle, True);
        Ferme(QArticle);
      end;
    end;
    for Ind := 0 to TOBArticles.Detail.Count -1 do
    begin
      TOBArticles.detail[Ind].AddChampSup('UTILISE', False);
      TOBArticles.detail[Ind].PutValue('UTILISE', '-');
      TOBArticles.detail[Ind].AddChampSup('REFARTSAISIE', False);
      TOBArticles.detail[Ind].AddChampSup('REFARTBARRE', False);
      TOBArticles.detail[Ind].AddChampSup('REFARTTIERS', False);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/07/2003
Modifié le ... : 31/07/2003
Description .. : Réimpression des bons (avoirs, arrhes, ...) liés aux lignes et
Suite ........ : ceux (avoirs, ...) liés aux échéances d'une pièce
Mots clefs ... : FO
*****************************************************************}

procedure FOReImprimeBons;
var
  CleDoc: R_CleDoc;
  TOBPiece, TOBArticles, TOBEches, TOBOpCais: TOB;
  QQ: TQuery;
  sSql: string;
begin
  if not FOJaiLeDroit(39) then Exit;
  // Recherche du dernier n° de ticket de la caisse
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := FOGetNatureTicket(False, False) ;
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'), '');
  CleDoc.NumeroPiece := VH_GC.TOBPCaisse.GetValue('LASTNUMTIC');
  if CleDoc.NumeroPiece = 0 then CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche) - 1;
  CleDoc.Indice := 0;
  // Chargement de la pièce
  TOBPiece := TOB.Create('PIECE', nil, -1);
  sSql := 'SELECT * FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False);
  QQ := OpenSQL(sSql, True);
  TOBPiece.SelectDB('', QQ);
  Ferme(QQ);
  // Lecture Lignes
  sSql := 'SELECT * FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False) + ' ORDER BY GL_NUMLIGNE';
  QQ := OpenSQL(sSql, True);
  TOBPiece.LoadDetailDB('LIGNE', '', '', QQ, False, True);
  Ferme(QQ);
  PieceAjouteSousDetail(TOBPiece);
  // Lecture Articles
  TOBArticles := TOB.Create('ARTICLES', nil, -1);
  FOLoadTobArt(TOBPiece, TOBArticles);
  // Lecture Echéances
  TOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  sSQL := 'SELECT PIEDECHE.*,MP_TYPEMODEPAIE AS TYPEMODEPAIE FROM PIEDECHE LEFT JOIN MODEPAIE ON GPE_MODEPAIE=MP_MODEPAIE WHERE '
    + WherePiece(CleDoc, ttdEche, False);
  QQ := OpenSQL(sSql, True);
  TOBEches.LoadDetailDB('PIEDECHE', '', '', QQ, False);
  Ferme(QQ);
  // Lecture des opérations de caisse
  TOBOpCais := TOB.Create('', nil, -1);
  FOLoadOperCaisse(TOBPiece, TOBEches, TOBOpCais);
  // Impression des bons liés aux lignes de la pièce
  FOImpressionBons(TOBPiece, TOBArticles);
  // Impression des bons liés aux échéances de la pièce
  FOImpressionAvoirs(TOBEches);
  // Libération des TOBs
  TOBPiece.Free;
  TOBArticles.Free;
  TOBEches.Free;
  TOBOpCais.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Impression des bons (avoirs, arrhes, ...) liés aux lignes du
Suite ........ : ticket
Mots clefs ... : FO
*****************************************************************}

procedure FOImpressionBons(TOBPiece, TOBArticles: TOB);
{$IFDEF FOS5}
var TOBL, TOBA: TOB;
  Ind, DefaultDocCopies: Integer;
  Montant: Double;
  ImpDoc: Boolean;
  TypeEtat: TTypeEtatFO;
  TypeOpc, Stg, Titre: string;
begin
  Montant := 0;
  for Ind := 1 to TOBPiece.Detail.Count do
  begin
    ImpDoc := False;
    TypeEtat := efoBonAchat;
    TOBL := TOBPiece.Detail[(Ind - 1)];
    if (TOBL <> nil) and (TOBL.GetValue('GL_TYPELIGNE') = 'ART') and (TOBL.GetValue('GL_TYPEARTICLE') = 'FI') then
    begin
      TOBA := FindTOBArtRow(TOBPiece, TOBArticles, Ind);
      if TOBA <> nil then TypeOpc := TOBA.GetValue('GA_TYPEARTFINAN') else TypeOpc := '';
      Montant := TOBL.GetValue('GL_TOTALTTCDEV');
      if (TypeOpc = TYPEOPCBONACHAT) and (Montant > 0) and (FOExisteCodeEtat(efoBonAchat)) then
      begin
        // Impression d'un bon d'achat
        ImpDoc := True;
        TypeEtat := efoBonAchat;
        Titre := 'du bon d''achat';
      end;
      if (TypeOpc = TYPEOPCARRHES) and (Montant > 0) and (FOExisteCodeEtat(efoBonArrhes)) then
      begin
        // Impression d'un bon de versement d'arrhes
        ImpDoc := True;
        TypeEtat := efoBonArrhes;
        Titre := 'du versement d''arrhes';
      end;
    end;
    if ImpDoc then
    begin
      Stg := TraduireMemoire('Voulez-vous imprimer un bon de') + ' '
        + FOStrFMontant(Abs(Montant), 12, TOBL.GetValue('GL_DEVISE'))
        + ' ' + TraduireMemoire('qui correspond au montant')
        + ' ' + TraduireMemoire(Titre) + ' ?';
      Titre := TraduireMemoire('Impression ') + Titre;
      if PGIAsk(Stg, Titre) = mrYes then
      begin
        DefaultDocCopies := V_PGI.DefaultDocCopies;
        V_PGI.DefaultDocCopies := FONbExemplaire;
        FOLanceImprimeLP(TypeEtat, '', False, TOBL);
        V_PGI.DefaultDocCopies := DefaultDocCopies;
      end;
    end;
  end;
end;
{$ELSE}
begin
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Impression des chèques d'une pièce
Mots clefs ... : FO
*****************************************************************}

procedure FOImpressionCheque(TOBEche, TOBMode: TOB);
var TOBL, TOBM: TOB;
  Ind: Integer;
  Montant: Double;
  CodeM, TypeM, Rep: string;
begin
  for Ind := 1 to TOBEche.Detail.Count do
  begin
    TOBL := TOBEche.Detail[(Ind - 1)];
    Montant := TOBL.GetValue('GPE_MONTANTENCAIS');
    CodeM := TOBL.GetValue('GPE_MODEPAIE');
    TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [CodeM], False);
    if TOBM <> nil then
    begin
      TypeM := TOBM.GetValue('MP_TYPEMODEPAIE');
      if (TypeM = TYPEPAIECHEQUE) or (TypeM = TYPEPAIECHQDIFF) then
      begin
        if (Montant > 0) and (TOBM.GetValue('MP_EDITCHEQUEFO') = 'X') and
          (FOExisteLPCheque) and (FOExisteCodeEtat(efoCheque)) then
        begin
          // Impression d'un chèque
          TOBL.AddChampSupValeur('GPE_MONTANTLETTRE', ChiffreLettre(Montant, FODonneMaskDevise(TOBL.GetValue('GPE_DEVISEESP'))));
          TOBL.AddChampSupValeur('GPE_ORDRE', GetParamSoc('SO_LIBELLE'));
          TOBL.AddChampSupValeur('GPE_LOCALITE', GetParamSoc('SO_VILLE'));
          TOBL.AddChampSupValeur('GPE_DATE', DateToStr(Date));
          TheTOB := TOBL;
          Rep := AGLLanceFiche('MFO', 'IMPCHEQUE', '', '', '');
          if Rep = 'OK' then FOLanceImprimeLP(efoCheque, '', False, TOBL);
        end;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Impression des bons (avoirs, ...) liés aux
Suite ........ : échéances d'une pièce
Mots clefs ... : FO
*****************************************************************}

procedure FOImpressionAvoirs(TOBEche: TOB);
var TOBL: TOB;
  Ind, DefaultDocCopies: Integer;
  Montant: Double;
  TypeM, Rep: string;
begin
  for Ind := 1 to TOBEche.Detail.Count do
  begin
    TOBL := TOBEche.Detail[(Ind - 1)];
    Montant := TOBL.GetValue('GPE_MONTANTENCAIS');
    if TOBL.FieldExists('TYPEMODEPAIE') then
    begin
      TypeM := TOBL.GetValue('TYPEMODEPAIE');
      if (TypeM = TYPEPAIEAVOIR) and (Montant < 0) and (FOExisteCodeEtat(efoBonAvoir)) then
      begin
        // Impression d'un bon d'avoir
        Rep := TraduireMemoire('Voulez-vous imprimer un bon de') + ' '
          + FOStrFMontant(Abs(Montant), 12, TOBL.GetValue('GPE_DEVISEESP'))
          + ' ' + TraduireMemoire('qui correspond au montant de l''avoir ?');
        if PGIAsk(Rep, 'Impression d''un bon d''avoir') = mrYes then
        begin
          DefaultDocCopies := V_PGI.DefaultDocCopies;
          V_PGI.DefaultDocCopies := FONbExemplaire;
          FOLanceImprimeLP(efoBonAvoir, '', False, TOBL);
          V_PGI.DefaultDocCopies := DefaultDocCopies;
        end;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... :   /  /
Description .. : Calcule la marge (en montant) sur une ligne d'une pièce à
Suite ........ : partir du hors taxe.
Mots clefs ... : FO
*****************************************************************}

function FOCalculeMargeLigne(TOBLigne: TOB): double;
var PxValo, PuNet: double;
begin
  Result := 0;
  if TOBLigne = nil then Exit;
  if TOBLigne.GetValue('GL_ARTICLE') = '' then Exit;
  if not FOEstUneLigneVente(TOBLigne) then Exit;
  if VH_GC.GCMargeArticle = 'PMA' then PxValo := TOBLigne.GetValue('GL_PMAP') else
    if VH_GC.GCMargeArticle = 'PMR' then PxValo := TOBLigne.GetValue('GL_PMRP') else
    if VH_GC.GCMargeArticle = 'DPA' then PxValo := TOBLigne.GetValue('GL_DPA') else
    if VH_GC.GCMargeArticle = 'DPR' then PxValo := TOBLigne.GetValue('GL_DPR') else Exit;
  PuNet := TOBLigne.GetValue('GL_PUHTNET');
  Result := PuNet - PxValo;
end;

{==============================================================================================}
{=============================== CONVERSION DES MONTANTS ======================================}
{==============================================================================================}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Convertit le montant d'une échéance dans les différentes
Suite ........ : devises
Mots clefs ... : FO
*****************************************************************}

procedure FOConvtMntEcheance(MntDevSaisi: Double; CodeDev: string; TOBPiec, TOBDev: TOB; DEV: RDEVISE; var MontantP, MontantD: Double);
var MntSaisi: Double;
begin
  // Conversion en devise de la pièce
  MntSaisi := ToxConvToDev(MntDevSaisi, CodeDev, DEV.Code, TOBPiec.GetValue('GP_DATEPIECE'), TOBDev);
  // conversion en devise de la pièce et pivot
  if TOBPiec.GetValue('GP_DEVISE') = V_PGI.DevisePivot then
  begin
    // Saisie en devise pivot
    MontantP := MntSaisi;
    MontantD := MntSaisi;
  end else
  begin
    // Saisie en devise
    MontantP := DeviseToEuro(MntSaisi, TOBPiec.GetValue('GP_TAUXDEV'), DEV.Quotite);
    MontantD := MntSaisi;
  end;
end;

{==============================================================================================}
{=============================== ANNULATION D'UN TICKET =======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Marque les échéances comme remises en banque en cas
Suite ........ : d'annulation d'un ticket
Mots clefs ... : FO
*****************************************************************}

function FOMarqueChqDifUtilise(TOBEches: TOB; MajSql: Boolean = False): Boolean;
var TOBL: TOB;
  Ind: Integer;
  sSql: string;
  CleDoc: R_CLEDOC;
begin
  Result := False;
  if TOBEches = nil then Exit;
  sSql := '';
  for Ind := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBL := TOBEches.Detail[Ind];
    if TOBL.GetValue('GPE_CHQDIFUTIL') = '-' then
    begin
      if MajSql then
      begin
        if sSql = '' then
        begin
          CleDoc := TOB2CleDoc(TOBL);
          sSql := 'UPDATE PIEDECHE SET GPE_CHQDIFUTIL="X" WHERE' + WherePiece(CleDoc, ttdEche, False)
            + 'AND GPE_NUMECHE IN ("' + IntToStr(TOBL.GetValue('GPE_NUMECHE')) + '"';
        end else
        begin
          sSql := sSql + ',"' + IntToStr(TOBL.GetValue('GPE_NUMECHE')) + '"';
        end;
      end else
      begin
        TOBL.PutValue('GPE_CHQDIFUTIL', 'X');
      end;
      Result := True;
    end;
  end;
  if (Result) and (MajSql) and (sSql <> '') then
  begin
    sSql := sSql + ')';
    ExecuteSQL(sSql);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 30/10/2002
Modifié le ... : 30/10/2002
Description .. : Marque la pièce d'origine comme étant annulée sans mettre
Suite ........ : à jour la date de modification pour ne pas "toxer" la pièce.
Mots clefs ... : FO
*****************************************************************}

procedure FOMarqueTicketAnnule(TOBPiece_O, TOBPiece: TOB);
var sSql, sRefSuiv: string;
  CleDoc: R_CLEDOC;
begin
  if not TOBPiece_O.FieldExists('GP_TICKETANNULE') then Exit;
  sRefSuiv := EncodeRefPiece(TOBPiece);
  CleDoc := TOB2CleDoc(TOBPiece_O);
  sSql := 'UPDATE PIECE SET GP_DEVENIRPIECE="' + sRefSuiv + '",GP_TICKETANNULE="X"'
    + ' WHERE ' + WherePiece(CleDoc, ttdPiece, False);
  ExecuteSQL(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/02/2003
Modifié le ... : 20/02/2003
Description .. : Vérifie si un ticket peut être annulé.
Mots clefs ... : FO
*****************************************************************}

function FOVerifTicketAnnulable(CleDoc: R_CleDoc): boolean;
var NbJours: integer;
  DateMin: TDateTime;
  Stg: string;
begin
  Result := False;
  // Plage d'annulation
  NbJours := FOGetParamCaisse('GPK_NBJANNULTIC', 'INTEGER');
  if NbJours > 0 then
  begin
    // Nbjours = 1 signifie que l'on ne peut annuler que les tickets du jour.
    DateMin := PlusDate(V_PGI.DateEntree, ((NbJours - 1) * -1), 'J');
    if CleDoc.DatePiece < DateMin then
    begin
      Stg := TraduireMemoire('Vous ne pouvez pas annuler un ticket antérieur au ')
        + DateToStr(DateMin);
      PGIBox(Stg);
      Exit;
    end;
  end;
  // Pièce non annulée
  Stg := 'SELECT GP_TICKETANNULE FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False)
    + ' AND GP_TICKETANNULE="X"';
  if ExisteSQL(Stg) then
  begin
    PGIBox('Annulation impossible. Ce ticket a déjà été annulé.');
    Exit;
  end;
  Result := True;
end;

{==============================================================================================}
{================================= CALCUL DES REMISES =========================================}
{==============================================================================================}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 30/12/2002
Modifié le ... : 30/12/2002
Description .. : Calcul du montant brut d'une ligne
Mots clefs ... : FO
*****************************************************************}

function FOCalculMontantBrutLigne(TOBLigne: TOB): double;
var
  QteF, PUDEV, PQ: double;
begin
  Result := 0;
  if TOBLigne = nil then Exit;
  // Calcul du montant brut
  QteF := TOBLigne.GetValue('GL_QTEFACT');
  if QteF = 0 then Exit;
  if TOBLigne.GetValue('GL_FACTUREHT') = 'X' then
    PUDEV := TOBLigne.GetValue('GL_PUHTDEV')
  else
    PUDEV := TOBLigne.GetValue('GL_PUTTCDEV');
  PQ := TOBLigne.GetValue('GL_PRIXPOURQTE');
  if PQ <= 0 then
  begin
    PQ := 1.0;
    TOBLigne.PutValue('GL_PRIXPOURQTE', PQ);
  end;
  Result := Arrondi(PUDEV * QteF / PQ, 6);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Calcul le pourcentage de remise accordée à partir du
Suite ........ : montant de la remise
Mots clefs ... : FO
*****************************************************************}

procedure FOCalculRemiseLigne(TOBLigne: TOB; NomChamp: string);
var
  MntRemise, MontantL, RemL: double;
begin
  RemL := 0;
  if TOBLigne = nil then Exit;
  // Montant de la remise
  MntRemise := TOBLigne.GetValue(NomChamp);
  // Calcul du montant brut
  MontantL := FOCalculMontantBrutLigne(TOBLigne);
  // Calcul du pourcentage de remise
  if MontantL <> 0 then RemL := Arrondi(100.0 * MntRemise / MontantL, 6);
  TOBLigne.PutValue('GL_REMISELIGNE', RemL);
  TOBLigne.PutValue('GL_RECALCULER', 'X');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 30/12/2002
Modifié le ... : 30/12/2002
Description .. : Calcul de la remise d'une ligne à partir du montant de remise
Suite ........ : globale
Mots clefs ... : FO
*****************************************************************}

function FOCalculRemise(TOBLigne: TOB; TotalTicket, RemiseGlobale: double; DEV: RDEVISE; SurLeBrut: boolean): double;
var MontantLigne, MontantRemise, TauxRemise, RemiseLigne: double;
begin
  if SurLeBrut then
  begin
    // sur le brut (on remplace la remise existante)
    MontantLigne := FOCalculMontantBrutLigne(TOBLigne);
    MontantRemise := 0;
  end else
  begin
    // sur le net (on ajoute à la remise existante)
    MontantLigne := TOBLigne.GetValue('GL_MONTANTTTCDEV');
    MontantRemise := TOBLigne.GetValue('GL_TOTREMLIGNEDEV');
  end;
  TauxRemise := MontantLigne / TotalTicket;
  RemiseLigne := Arrondi((RemiseGlobale * TauxRemise), DEV.Decimale);
  TOBLigne.PutValue('GL_VALEURREMDEV', (RemiseLigne + MontantRemise));
  FOCalculRemiseLigne(TOBLigne, 'GL_VALEURREMDEV');
  Result := RemiseLigne;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 16/01/2004
Modifié le ... : 16/01/2004
Description .. : Calcul du montant de la base de calcul de la remise globale
Mots clefs ... : FO
*****************************************************************}

function FOCalculBaseRemiseGlobale(TOBPiece: TOB; Tous: boolean = False): double;
var
  Ind: integer;
  TOBL: TOB;
begin
  Result := 0;
  if TOBPiece = nil then Exit;

  if Tous then
  begin
    Result := TOBPiece.GetValue('GP_TOTALTTCDEV');
  end else
  begin
    // cumul du montant des lignes remisables
    for Ind := 0 to TOBPiece.Detail.Count - 1 do
    begin
      TOBL := TOBPiece.Detail[Ind];
      if (FOEstUneLigneVente(TOBL)) and (FOLigneRemisable(TOBL, True)) then
        Result := Result + TOBL.GetValue('GL_MONTANTTTCDEV');
    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 30/12/2002
Modifié le ... : 30/12/2002
Description .. : Répartition d'un montant de remise globale sur les lignes
Suite ........ : d'une pièce
Mots clefs ... : FO
*****************************************************************}

procedure FOAppliqueRemiseGlobale(TOBPiece: TOB; RemiseGlobale: double; TypeRemise: string; DEV: RDEVISE);
var
  Ind, iLast, iNumType: integer;
  TOBL: TOB;
  TotalTicket, ResteRemise: double;
begin
  if RemiseGlobale = 0 then Exit;
  if TOBPiece = nil then Exit;
  TotalTicket := FOCalculBaseRemiseGlobale(TOBPiece);
  if TotalTicket = 0 then Exit;
  // Mise à jour du motif de remise
  TOBPiece.PutValue('GP_TYPEREMISE', TypeRemise);
  TOBPiece.PutValue('GP_RECALCULER', 'X');
  // Répartition du montant de la remise sur les lignes de la pièce
  ResteRemise := RemiseGlobale;
  iLast := 0;
  iNumType := -1;
  for Ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[Ind];
    if iNumType < 0 then iNumType := TOBL.GetNumChamp('GL_TYPEREMISE');
    
    if (FOEstUneLigneVente(TOBL)) and (FOLigneRemisable(TOBL, True)) then
    begin
      iLast := Ind;
      ResteRemise := ResteRemise - FOCalculRemise(TOBL, TotalTicket, RemiseGlobale, DEV, False);
      if TOBL.GetValeur(iNumType) = '' then TOBL.PutValeur(iNumType, TypeRemise);
    end;
  end;
  // Affectation du résidu de la remise sur la dernière ligne
  if ResteRemise <> 0 then
  begin
    TOBL := TOBPiece.Detail[iLast];
    TOBL.PutValue('GL_VALEURREMDEV', (ResteRemise + TOBL.GetValue('GL_VALEURREMDEV')));
    FOCalculRemiseLigne(TOBL, 'GL_VALEURREMDEV');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 30/12/2002
Modifié le ... : 30/12/2002
Description .. : Répartition d'un montant de remise sur les lignes d'un
Suite ........ : sous-total d'une pièce
Suite ........ :
Suite ........ : Valeur retournée :
Suite ........ :  0 = ok
Suite ........ :  1 = maximum autorisé pour la caisse dépassé
Suite ........ :  2 = maximum autorisé pour le type de remise dépassé
Suite ........ :  3 = majoration interdite
Mots clefs ... : FO
*****************************************************************}

function FOAppliqueRemiseSousTotal(TOBPiece: TOB; ARow: integer; TauxRemise, RemiseSousTotal, MontantNet: double; TypeRemise: string; DEV: RDEVISE; var MaxRem:
  integer): integer;
var Ind, iLast: integer;
  TOBL, TOBBis: TOB;
  Ok, TypeUniquement, MajorOk: boolean;
  MontantTotal, ResteRemise, TauxR: double;
  CodeDem, CodeDemValide: string;
begin
  Result := 0;
  MaxRem := 0;
  CodeDemValide := '';
  MajorOk := False;
  if (TauxRemise = 0) and (RemiseSousTotal = 0) and (MontantNet = 0) and (TypeRemise = '') then Exit;
  if TOBPiece = nil then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_TYPELIGNE') <> 'TOT' then Exit;
  // Montant brut du sous-total
  MontantTotal := 0;
  for Ind := ARow - 1 downto 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, Ind);
    if TOBL = nil then Break;
    if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Break;
    if (FOEstUneLigneVente(TOBL)) and (FOLigneRemisable(TOBL, False)) then
      MontantTotal := MontantTotal + FOCalculMontantBrutLigne(TOBL);
  end;
  if (TauxRemise = 0) and (RemiseSousTotal = 0) and (MontantNet <> 0) then
    RemiseSousTotal := MontantTotal - MontantNet;
  if (RemiseSousTotal <> 0) and (MontantTotal = 0) then Exit;
  // Vérification du maximum de remise autorisé
  for Ind := ARow - 1 downto 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, Ind);
    if TOBL = nil then Break;
    if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Break;
    if (FOEstUneLigneVente(TOBL)) and (FOLigneRemisable(TOBL, False)) then
    begin
      if RemiseSousTotal <> 0 then
      begin
        // remise en montant
        TOBBis := TOB.Create('', nil, -1);
        TOBBis.Dupliquer(TOBL, True, True);
        FOCalculRemise(TOBBis, MontantTotal, RemiseSousTotal, DEV, True);
        TauxR := TOBBis.GetValue('GL_REMISELIGNE');
        TOBBis.Free;
      end else
        if TauxRemise <> 0 then
      begin
        // remise en pourcentage
        TauxR := TauxRemise;
      end else
      begin
        TauxR := TOBL.GetValue('GL_REMISELIGNE');
      end;
      // Vérification du % maximun autorisé
      if TypeRemise = '' then
        CodeDem := TOBL.GetValue('GL_TYPEREMISE')
      else
        CodeDem := TypeRemise;
      TypeUniquement := ((TauxRemise = 0) and (RemiseSousTotal = 0));
      //if TauxR > MaxRem then
      if ((MajorOk) and (TauxR < 0)) or ((TauxR <= MaxRem) and (TauxR > 0)) then
      begin
        Result := 0;
      end else
      begin
        Result := FOVerifMaxRemise(TauxR, CodeDem, TypeUniquement, MaxRem, CodeDemValide);
        if (Result = 0) and (TauxR < 0) then MajorOk := True;
      end;
      if Result <> 0 then Exit;
      // pour ne saisir qu'une fois le mot de passe pour le motif de démarque saisi sur le sous-total
      if (CodeDemValide <> '') and (CodeDemValide <> TypeRemise) then
        CodeDemValide := '';
    end;
  end;
  // Répartition du montant de la remise sur les lignes du sous-total
  iLast := 0;
  ResteRemise := RemiseSousTotal;
  Ok := False;
  for Ind := ARow - 1 downto 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, Ind);
    if TOBL = nil then Break;
    if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Break;
    if (FOEstUneLigneVente(TOBL)) and (FOLigneRemisable(TOBL, False)) then
    begin
      if RemiseSousTotal <> 0 then
      begin
        // remise en montant
        iLast := Ind;
        ResteRemise := ResteRemise - FOCalculRemise(TOBL, MontantTotal, RemiseSousTotal, DEV, True);
      end else
        if TauxRemise <> 0 then
      begin
        // remise en pourcentage
        TOBL.PutValue('GL_REMISELIGNE', TauxRemise);
        TOBL.PutValue('GL_VALEURREMDEV', 0);
        TOBL.PutValue('GL_RECALCULER', 'X');
      end;
      Ok := True;
      // motif de démarque
      if (TypeRemise <> '') and (TOBL.GetValue('GL_TOTREMLIGNEDEV') <> 0) and
        (TOBL.GetValue('GL_TYPEREMISE') = '') then
        TOBL.PutValue('GL_TYPEREMISE', TypeRemise);
    end;
  end;
  // Affectation du résidu de la remise sur la dernière ligne
  if ResteRemise <> 0 then
  begin
    TOBL := GetTOBLigne(TOBPiece, iLast);
    TOBL.PutValue('GL_VALEURREMDEV', (ResteRemise + TOBL.GetValue('GL_VALEURREMDEV')));
    FOCalculRemiseLigne(TOBL, 'GL_VALEURREMDEV');
    Ok := True;
  end;
  if Ok then TOBPiece.PutValue('GP_RECALCULER', 'X');
end;

{==============================================================================================}
{================================= PIECES EN ATTENTE ==========================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/03/2002
Modifié le ... : 12/03/2003
Description .. : Met une pièce en attente.
Suite ........ :
Suite ........ : Code retour :
Suite ........ :   0 => pièce mise en attente
Suite ........ :   1 => pièce abandonnée
Suite ........ :   2 => mise en attente interdite
Mots clefs ... : FO
*****************************************************************}

function FOMisePieceEnAttente(PieceContainer: TPieceContainer) : integer;
var Stg: string;
    MajPceAttente: T_MajPieceEnAttente;
begin
  Result := 1;
  if PieceContainer.TCPiece.GetValue('GP_NATUREPIECEG') <> FOGetNatureTicket(False, False) then Exit;
  Stg := AGLLanceFiche('MFO', 'MISEATTENTE', '', '', '');
  if ReadTokenSt(Stg) <> 'OK' then Exit;
  if not FOJaiLeDroit(37) then
  begin
    Result := 2;
    Exit;
  end;
  PieceContainerLoc := PieceContainer;
  SourisSablier;
  MajPceAttente := T_MajPieceEnAttente.Create;
  if Transactions(MajPceAttente.MajPieceEnAttente, 1) <> oeOk then
    Result := 1
    else
    Result := 0;
  MajPceAttente.Free;
  SourisNormale;
end;

procedure T_MajPieceEnAttente.MajPieceEnAttente;
var CleDoc: R_CleDoc;
    Ind: integer;
    NowFutur: TDateTime;
    Stg: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := FOGetNatureTicket(True, False);
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, PieceContainerLoc.TCPiece.GetValue('GP_ETABLISSEMENT'), '');
  CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche);
  CleDoc.Indice := 0;
  CleDoc.DatePiece := V_PGI.DateEntree;
  MajFromCleDoc(PieceContainerLoc.TCPiece, CleDoc);
  if Stg <> '' then PieceContainerLoc.TCPiece.PutValue('GP_REFEXTERNE', Stg);
  for Ind := 0 to PieceContainerLoc.TCPiece.Detail.Count - 1 do
    MajFromCleDoc(PieceContainerLoc.TCPiece.Detail[Ind], CleDoc);
  // JTR - eQualité 11203 (gestion des formules)
  for Ind := 0 to PieceContainerLoc.TCLigFormule.Detail.Count - 1 do
    MajFromCleDoc(PieceContainerLoc.TCLigFormule.Detail[Ind], CleDoc, True);
  // Fin JTR
  FOSetNumeroDefinitif(PieceContainerLoc);
  NowFutur := NowH;
  PieceContainerLoc.TCPiece.SetAllModifie(True);
  PieceContainerLoc.TCPiece.SetDateModif(NowFutur);
  PieceContainerLoc.TCPiece.InsertDBByNivel(False);
  // JTR - eQualité 11203
  if PieceContainerLoc.TCLigFormule.Detail.count > 0 then
  begin
    PieceContainerLoc.TCLigFormule.SetAllModifie(True);
    ValideLesFormules(PieceContainerLoc.TCPiece,PieceContainerLoc.TCLigFormule);
  end;
  FOIncrJoursCaisse(jfoNbTicMisatt);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/03/2002
Modifié le ... : 26/03/2002
Description .. : Reprise d'une pièce en attente
Suite ........ : JTR - eQualité 11203. Charge aussi les formules des quantités
Mots clefs ... : FO
*****************************************************************}
function FOReprisePieceEnAttente(TOBPiece: TOB; GS: THGrid; DejaSaisie: Boolean; var TobAtt : TOB; TobLigFormule : TOB = nil): Boolean;
var TOBLgAtt, TOBLigne, TobAttFormule: TOB;
  CleDoc: R_CleDoc;
  Stg, sSql, sEtab, sDep, sCais: string;
  Ind, NumLigne: Integer;
  QQ: TQuery;
  sCli, sCliA: string;
begin
  Result := False;
  if (TOBPiece = nil) or (GS = nil) then Exit;
  if TOBPiece.GetValue('GP_NATUREPIECEG') <> FOGetNatureTicket(False, False) then Exit;
  // Choix de la pièce en attente
  Stg := AGLLanceFiche('MFO', 'PIECEATTENTE_MUL', 'GP_NATUREPIECEG='+ FOGetNatureTicket(True, False)
       + ';GP_CAISSE=' + FOCaisseCourante, '', '');
  if Stg = '' then Exit;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StringToCleDoc(Stg, CleDoc);
  TOBAtt := TOB.Create('PIECE', nil, -1);
  // Chargement de la pièce en attente
  sSql := 'SELECT * FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False);
  QQ := OpenSQL(sSql, True);
  if QQ.Eof then
  begin
    TOBAtt.Free;
    Ferme(QQ);
    Exit;
  end;
  TOBAtt.SelectDB('', QQ);
  Ferme(QQ);
  // Chargement des lignes
  sSql := 'SELECT * FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False) + ' ORDER BY GL_NUMLIGNE';
  QQ := OpenSQL(sSql, True);
  if QQ.Eof then
  begin
    TOBAtt.Free;
    Ferme(QQ);
    Exit;
  end;
  TOBAtt.LoadDetailDB('LIGNE', '', '', QQ, False, True);
  Ferme(QQ);
  // JTR - eQualité 11203 - Charge les formules
  LoadLesFormules(TobAtt,TobLigFormule,CleDoc) ;
  TobAttFormule := TOB.Create('',nil,-1);
  TobAttFormule.Dupliquer(TobLigFormule,true,true);
  // Fin JTR
  Result := True;
  if not DejaSaisie then
  begin
    // JTR - eQualité 11203  - Chargement des formules
    for Ind:= 0 to TobLigFormule.detail.count-1 do
      TobLigFormule.detail[Ind].SetString('GLF_NATUREPIECEG',FOGetNatureTicket(False, False));
    // Fin JTR
    // Remplacement de la pièce courante par la pièce en attente
    CleDoc := TOB2CleDoc(TOBPiece);
    sEtab := TOBPiece.GetValue('GP_ETABLISSEMENT');
    sDep := TOBPiece.GetValue('GP_DEPOT');
    sCais := TOBPiece.GetValue('GP_CAISSE');
    TOBPiece.Dupliquer(TOBAtt, True, True);
    // Ajout d'une ligne vide pour se positionner dans l'écran de saisie
    TOB.Create('LIGNE', TOBPiece, -1);
    AddLesSupEntete(TOBPiece);
    InitTOBPiece(TOBPiece);
    TOBAtt.PutValue('GP_ETABLISSEMENT', sEtab);
    TOBAtt.PutValue('GP_DEPOT', sDep);
    TOBAtt.PutValue('GP_CAISSE', sCais);
    TOBAtt.PutValue('GP_REFEXTERNE', '');
    MajFromCleDoc(TOBPiece, CleDoc);
    for Ind := 0 to TOBPiece.Detail.Count - 1 do
      MajFromCleDoc(TOBPiece.Detail[Ind], CleDoc);
    PieceAjouteSousDetail(TOBPiece);
    GS.Row := TOBPiece.Detail.Count;
    NumeroteLignesGC(GS, TOBPiece);
    PieceAjouteSousDetail(TOBPiece);
    PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
  end else
  begin
    // Si le client du ticket en attente n'est pas celui du ticket courant
    sCli := TOBPiece.GetValue('GP_TIERS');
    sCliA := TOBAtt.GetValue('GP_TIERS');
    if (sCliA <> FODonneClientDefaut) and (sCliA <> sCli) then
      PGIInfo('Le client du ticket en attente est différent de celui du ticket courant');
    // Insertion de la pièce en attente dans la pièce courante
    GS.CacheEdit;
    GS.SynEnabled := False;
    NumLigne := GS.Row;
    for Ind := 0 to TOBAtt.Detail.Count - 1 do
    begin
      TOBLgAtt := TOBAtt.Detail[Ind];
      TOBLigne := TOB.Create('LIGNE', TOBPiece, NumLigne - 1);
      GS.InsertRow(NumLigne);
      TOBLigne.Dupliquer(TOBLgAtt, False, True);
      TOB.Create('', TOBLigne, -1); {Analytique vente/achat}
      TOB.Create('', TOBLigne, -1); {Analytique stock}
      AddLesSupLigne(TOBLigne, False);
      InitLesSupLigne(TOBLigne);
      PieceVersLigne(TOBPiece, TOBLigne);
      Stg := TOBLgAtt.GetValue('GL_REPRESENTANT');
      if Stg <> '' then TOBLigne.PutValue('GL_REPRESENTANT', Stg);
      TOBLigne.PutValue('GL_CAISSE', TOBPiece.GetValue('GP_CAISSE'));
      TOBLigne.PutValue('GL_RECALCULER', 'X');
      Inc(NumLigne);
    end;
    NumeroteLignesGC(nil, TOBPiece);
    // Pour forcer le réaffichage de la ligne
    if NumLigne > 0 then GS.Row := NumLigne - 1;
    GS.MontreEdit;
    GS.SynEnabled := True;
    GS.Row := NumLigne;
    NumeroteLignesGC(GS, TOBPiece);
    TOBPiece.PutValue('GP_RECALCULER', 'X');
  end;
  // Suppression de la pièce en attente
  if Result then
  begin
//    TOBAtt.DeleteDB;
    for Ind:= 0 to TobAttFormule.detail.count-1 do
      TobAttFormule.Detail[Ind].DeleteDB;
  end;
//  if TOBAtt <> nil then FreeAndNil(TOBAtt);
  if TobAttFormule <> nil then FreeAndNil(TobAttFormule); 
  FOIncrJoursCaisse(jfoNbTicattRepri);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/03/2002
Modifié le ... : 26/03/2002
Description .. : Supprime les pièces en attente d'une caisse.
Mots clefs ... : FO
*****************************************************************}

procedure FOSupprimePieceEnAttente(Caisse: string);
var sSql: string;
    Qry : TQuery;
    TobDesPieces : TOB;
    Cpt : integer;
    CleDoc : R_CleDoc;
begin
  BeginTrans;
  // JTR - eQualité 11203
  // Monte une tob des pièces pour supprimer les éventuelles lignesformules liées
  TobDesPieces := TOB.Create('Les pieces',nil, -1);
  sSQl := 'SELECT GP_NATUREPIECEG, GP_SOUCHE, GP_NUMERO, GP_INDICEG FROM PIECE'
        + ' WHERE GP_NATUREPIECEG='+ FOGetNatureTicket(True, True)+' AND GP_CAISSE="' + Caisse + '"';
  Qry := OpenSQL(sSQL,True);
  TObDesPieces.LoadDetailDB('PIECE','','',Qry,False);
  Ferme(Qry);
  // Fin JTR
  sSql := 'DELETE FROM PIECE WHERE GP_NATUREPIECEG='+ FOGetNatureTicket(True, True)
        + ' AND GP_CAISSE="' + Caisse + '"';
  ExecuteSQL(sSql);
  sSql := 'DELETE FROM LIGNE WHERE GL_NATUREPIECEG='+ FOGetNatureTicket(True, True)
        + ' AND GL_CAISSE="' + Caisse + '"';
  ExecuteSQL(sSql);
  // JTR - eQualité 11203
  for Cpt := 0 to TobDesPieces.detail.count -1 do
  begin
    CleDoc := TOB2CleDoc(TobDesPieces.detail[Cpt]);
    sSQL := 'DELETE FROM LIGNEFORMULE WHERE '+WherePiece(CleDoc, ttdLigneFormule, false);
    ExecuteSQL(sSQL);
  end;
  if TobDesPieces <> nil then FreeAndNil(TobDesPieces);
  // Fin JTR
  CommitTrans;
end;

{==============================================================================================}
{============================== ANNONCES DE LIVRAISON =========================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/10/2003
Modifié le ... : 13/10/2003
Description .. : Affiche une alerte si des annonces de livraison ou de
Suite ........ : transfert ne sont pas validées.
Mots clefs ... : FO
*****************************************************************}

procedure FOAlerteAnnoncesNonValidees(Etablis, Depot: string);
var
  sWhere, sNb, Stg: string;
  Ok: boolean;
begin
  if not (CtxMode in V_PGI.PGIContexte) then Exit;
  Ok := FOGetFromRegistry(REGJOURNEE, REGALERTEALF, True);
  if not Ok then Exit;

  Stg := '';
  if Etablis <> '' then
  begin
    sWhere := 'GP_NATUREPIECEG="ALF" AND GP_VIVANTE="X" AND GP_SUPPRIME<>"X" AND GP_ETABLISSEMENT="'+ Etablis +'"';
    sNb := GetColonneSQL('PIECE', 'COUNT(*)', sWhere);
    if StrToInt(sNb) > 0 then
       Stg := TraduireMemoire('Vous avez') + ' :#10  - '+ sNb +' '+ RechDom('GCNATUREPIECEG', 'ALF', False);
  end;

  if Depot <> '' then
  begin
    sWhere := 'GP_NATUREPIECEG="TRV" AND GP_VIVANTE="X" AND GP_SUPPRIME<>"X" AND GP_DEPOTDEST="'+ Depot +'"';
    sNb := GetColonneSQL('PIECE', 'COUNT(*)', sWhere);
    if StrToInt(sNb) > 0 then
    begin
      if Stg = '' then Stg := TraduireMemoire('Vous avez') + ' :';
      Stg := Stg +'#10  - '+ sNb +' '+ RechDom('GCNATUREPIECEG', 'TRV', False);
    end;
  end;

  if Stg <> '' then
  begin
    Stg := Stg +#10#10+ TraduireMemoire('Lors de la réception de la marchandise, pensez à valider ces annonces')
      +#10+ TraduireMemoire('pour mettre à jour votre stock.');
    SysUtils.Beep;
    PGIInfo(Stg);
  end;
end;

{==============================================================================================}
{================================= COMPTABILISATION ===========================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 16/05/2002
Description .. : Indique si la comptabilisation de la pièce est active
Mots clefs ... : FO
*****************************************************************}

function FOComptaEstActive(NaturePiece: string): Boolean;
var PassP: string;
begin
  Result := (not GetParamsoc('SO_GCDESACTIVECOMPTA'));
  if (Result) and (NaturePiece <> '') then
  begin
    PassP := GetInfoParPiece(NaturePiece, 'GPP_TYPEECRCPTA');
    Result := ((PassP <> '') and (PassP <> 'RIE'));
  end;
end;

{==============================================================================================}
{=============================== GENERATION D'UN TICKET =======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Génération d'un ticket composé d'une ligne article et d'une
Suite ........ : échéance
Suite ........ : par exemple ticket d'écart ou ticket de fond de caisse.
Mots clefs ... : FO
*****************************************************************}

function FOGenereTicket(CodeArt, CodeVen, CodeMode, CodeDev, CodeTiers, Caption: string; MntEche: Double; Affichage, Impression, RattachNumZ: Boolean):
  Integer;
var XX: TFOGenTicket;
  {$IFDEF GESCOM}
  SansCompta: boolean;
  {$ENDIF}
begin
  Result := 0;
  {$IFDEF GESCOM}
  SansCompta := False;
  if CodeTiers = '&#@' then
  begin
    CodeTiers := '';
    SansCompta := True;
  end;
  {$ENDIF}
  // Création du composant de génération de ticket
  XX := TFOGenTicket.Create;
  if XX <> nil then
  begin
    // Initialisation des variables
    if not XX.Init(CodeArt, CodeVen, CodeMode, CodeDev, CodeTiers, Caption, MntEche, Affichage, Impression, RattachNumZ) then
      Exit;
    {$IFDEF GESCOM}
    if SansCompta then
      XX.ActiveCompta := False;
    {$ENDIF}
    // Constitution de la pièce
    if XX.GenerePiece then
      Result := XX.CleDoc.NumeroPiece;
    XX.Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Create : création du composant
///////////////////////////////////////////////////////////////////////////////////////

constructor TFOGenTicket.Create;
begin
  inherited Create;
  // Création des TOB
  TOBPiece := TOB.Create('PIECE', nil, -1);
  AddLesSupEntete(TOBPiece);
  TOBBases := TOB.Create('BASES', nil, -1);
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBTiers.AddChampSup('RIB', False);
  TOBArticles := TOB.Create('ARTICLES', nil, -1);
  TOBPorcs := TOB.Create('PORCS', nil, -1);
  TOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  TOBAffaire := TOB.Create('AFFAIRE', nil, -1);
  TOBConds := TOB.Create('CONDS', nil, -1);
  TOBAcomptes := TOB.Create('', nil, -1);
  TOBCpta := TOB.Create('', nil, -1);
  TOBAnaP := TOB.Create('', nil, -1);
  TOBANAS := TOB.Create('', nil, -1);
  TOBComms := TOB.Create('COMMERCIAUX', nil, -1);
  PieceContainer := TPieceContainer.Create;
  PieceContainer.TCPiece        := TobPiece;
  PieceContainer.TCBases        := TOBBases;
  PieceContainer.TCEches        := TobEches;
  PieceContainer.TCTiers        := TobTiers;
  PieceContainer.TCArticles     := TobArticles;
  PieceContainer.TCPorcs        := TobPorcs;
  PieceContainer.TCConds        := TobConds;
  PieceContainer.TCComms        := TobComms;
  PieceContainer.TCAcomptes     := TobAcomptes;
  PieceContainer.TCAffaire      := TobAffaire;
  PieceContainer.TCCpta         := TobCpta;
  PieceContainer.TCAnaP         := TobAnaP;
  PieceContainer.TCAnaS         := TobAnaS;
{$IFDEF BTP}
  PieceContainer.TCPieceRG      := TobPieceRG;
  PieceContainer.TCBasesRG      := TobBasesRG;
{$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Destroy : destruction du composant
///////////////////////////////////////////////////////////////////////////////////////

destructor TFOGenTicket.Destroy;
///////////////////////////////////////////////////////////////////////////////////////
  procedure Libere(var TOBL: TOB);
  begin
    if TOBL <> nil then
    begin
      TOBL.Free;
      TOBL := nil;
    end;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
begin
  inherited Destroy;
  // Libération des TOB
  Libere(TOBPiece);
  Libere(TOBBases);
  Libere(TOBTiers);
  Libere(TOBArticles);
  Libere(TOBPorcs);
  Libere(TOBEches);
  Libere(TOBAffaire);
  Libere(TOBConds);
  Libere(TOBAcomptes);
  Libere(TOBCpta);
  Libere(TOBAnaP);
  Libere(TOBAnaS);
  Libere(TOBComms);
  PieceContainer.Free;  
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Init : Initialisation des variables
///////////////////////////////////////////////////////////////////////////////////////

function TFOGenTicket.Init(CodeArt, CodeVen, CodeMode, CodeDev, CodeTiers, Caption: string; MntEche: Double; Affichage, Impression, RattachNumZ: Boolean):  Boolean;
begin
  Result := True;
  RatNumZ := RattachNumZ;
  ImpTic := Impression;
  AffMsg := Affichage;
  Titre := Caption;
  Article := CodeArt;
  Vendeur := CodeVen;
  if CodeTiers = '' then
    Tiers := FODonneClientDefaut
    else
    Tiers := CodeTiers;
  ModePaie := CodeMode;
  Devise := CodeDev;
  MontDev := MntEche;
  // Initialisation de CleDoc
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := FOGetNatureTicket(False, False);
  CleDoc.DatePiece := V_PGI.DateEntree;
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'), '');
  CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche);
  if CleDoc.NumeroPiece <= 0 then
  begin
    Result := False;
    if AffMsg then PGIBox('La souche associée au ticket n''est pas valide !', Titre);
    Exit;
  end;
  CleDoc.Indice := 0;
  CleDoc.NumLigne := 1;
  // Devise
  DEV.Code := FOGetParamCaisse('DEVISECAISSE');
  if DEV.Code = '' then
    DEV.Code := V_PGI.DevisePivot;
  GetInfosDevise(DEV);
  DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, CleDoc.DatePiece);
  Montant := ToxConvToDev(MntEche, CodeDev, DEV.Code, CleDoc.DatePiece, nil);
  // Comptabilisation
  ActiveCompta := FOComptaEstActive(CleDoc.NaturePiece);
  PieceContainer.Dev := DEV;
  PieceContainer.CleDoc := CleDoc;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GenerePiece : Génération de la pièce
///////////////////////////////////////////////////////////////////////////////////////

function TFOGenTicket.GenerePiece: Boolean;
begin
  Result := False;
  // Génération de la pièce, des lignes et des échéances
  if (not MAJPiece) or (not MAJLigne) or (not MAJPiedEche) then Exit;
  // Calcul des différents montants de la piéce et des lignes
  PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X');
  CalculFacture(PieceContainer);
  // Enregistrement de la pièce
  if Transactions(ValideLaPiece, 10) <> oeOK then
  begin
    if AffMsg then PGIBox('La pièce n''est pas enregistrée !', Titre);
    Exit;
  end;
  // Impression de la pièce
  if (ImpTic) and (GetInfoParPiece(CleDoc.NaturePiece, 'GPP_IMPIMMEDIATE') = 'X') then
  begin
    if (Transactions(ValideImpression, 1) <> oeOk) and (AffMsg) then
    begin
      PGIBox('L''impression ne s''est pas correctement effectuée !', Titre);
    end;
  end;
  Result := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJPiece : maj de la table PIECE
///////////////////////////////////////////////////////////////////////////////////////

function TFOGenTicket.MAJPiece: Boolean;
var Stg: string;
begin
  Result := True;
  // Chargement du tiers par défaut de la caisse
  TOBTiers.InitValeurs;
  if RemplirTOBTiers(TOBTiers, Tiers, CleDoc.NaturePiece, True) <> trtOk then
  begin
    if AffMsg then PGIBox('Le tiers par défaut est incorrect !', Titre);
    V_PGI.IoError := oeUnknown;
    Result := False;
    Exit
  end;
  // Constitution de la pièce
  InitTOBPiece(TOBPiece);
  MajFromCleDoc(TOBPiece, CleDoc);
  TOBPiece.PutValue('GP_TIERS', TOBTiers.GetValue('T_TIERS'));
  TOBPiece.PutValue('GP_TVAENCAISSEMENT', PositionneExige(TOBTiers));
  TiersVersPiece(TOBTiers, TOBPiece);
  Stg := FOGetParamCaisse('REGIMETAXECAISSE');
  if Stg = '' then Stg := VH^.RegimeDefaut;
  TOBPiece.PutValue('GP_REGIMETAXE', Stg);
  FOAffecteFactureHT(TOBPiece);
  TOBPiece.PutValue('GP_ETABLISSEMENT', VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'));
  TOBPiece.PutValue('GP_DEPOT', VH_GC.TOBPCaisse.GetValue('GPK_DEPOT'));
  TOBPiece.PutValue('GP_CAISSE', FOCaisseCourante);
  if RatNumZ then TOBPiece.PutValue('GP_NUMZCAISSE', FOGetNumZCaisse(FOCaisseCourante, 'MAX'));
  TOBPiece.PutValue('GP_DEVISE', DEV.Code);
  TOBPiece.PutValue('GP_TAUXDEV', DEV.Taux);
  TOBPiece.PutValue('GP_DATETAUXDEV', CleDoc.DatePiece);
  if Vendeur <> '' then
  begin
    // Chargement du commercial
    AjouteRepres(Vendeur, GetInfoParPiece(CleDoc.NaturePiece, 'GPP_TYPECOMMERCIAL'), TOBComms);
    TOBPiece.PutValue('GP_REPRESENTANT', Vendeur);
  end;
  TOBPiece.PutValue('GP_CREEPAR', 'GEN');
  TOBPiece.PutValue('GP_VENTEACHAT', GetInfoParPiece(CleDoc.NaturePiece, 'GPP_VENTEACHAT'));
  AttribCotation(TOBPiece);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJLigne : maj de la table LIGNE
///////////////////////////////////////////////////////////////////////////////////////

function TFOGenTicket.MAJLigne: Boolean;
var TOBL, TOBA: TOB;
  ARow: Integer;
  QQ: TQuery;
  Stg: string;
begin
  Result := True;
  ARow := 1;
  // Chargement de l'article
  TOBA := CreerTOBArt(TOBArticles);
  QQ := OpenSQL('select * from ARTICLE where GA_CODEARTICLE="' + Article + '"', True);
  if QQ.Eof then
  begin
    Ferme(QQ);
    if AffMsg then
    begin
      Stg := TraduireMemoire('L''opération de caisse ') + Article
        + TraduireMemoire(' est incorrecte !');
      PGIBox(Stg, Titre);
    end;
    V_PGI.IoError := oeUnknown;
    Result := False;
    Exit;
  end;
  TOBA.SelectDB('', QQ, False);
  Ferme(QQ);
  // Constitution de la ligne
  TOBL := NewTOBLigne(TOBPiece, 0);
  InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
  TOBL.PutValue('GL_ARTICLE', TOBA.GetValue('GA_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE', TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_CODEARTICLE', TOBA.GetValue('GA_CODEARTICLE'));
  if PieceContainer.TCPiece = nil then
    PieceContainer.TCPiece := TobPiece;
  ArticleVersLigne(PieceContainer, TOBA, TOBL);
  TOBL.PutValue('GL_QTESTOCK', 1);
  TOBL.PutValue('GL_QTEFACT', 1);
  if TOBL.GetValue('GL_FACTUREHT') = 'X' then TOBL.PutValue('GL_PUHTDEV', Montant)
  else TOBL.PutValue('GL_PUTTCDEV', Montant);
  TOBL.PutValue('GL_CREERPAR', TOBPiece.GetValue('GP_CREEPAR'));
  TOBL.PutValue('GL_QUALIFMVT', GetInfoParPiece(CleDoc.NaturePiece, 'GPP_QUALIFMVT'));
  TOBL.PutValue('GL_REPRESENTANT', Vendeur);
  // Comptabilisation
  if ActiveCompta then TraiteLaCompta(ARow);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteLaCompta : comptabilisation d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOGenTicket.TraiteLaCompta(ARow: Integer);
var RefUnique: string;
  TOBL, TOBA, TOBC, TOBCata: TOB;
begin
  RefUnique := GetCodeArtUnique(TOBPiece, ARow);
  if RefUnique = '' then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if ActiveCompta then
  begin
    TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
    if TOBA = nil then Exit;
    TOBCata := nil;
    TOBC := ChargeAjouteCompta(TOBCpta, TOBPiece, TOBL, TOBA, TOBTiers, TOBCata, TobAffaire, True);
  end else
  begin
    TOBC := nil;
  end;
  PreVentileLigne(TOBC, TOBAnaP, TOBAnaS, TOBL);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJEcheance : maj de la table PIEDECHE
///////////////////////////////////////////////////////////////////////////////////////

function TFOGenTicket.MAJPiedEche: Boolean;
var MontantP, MontantD: Double;
  TOBL: TOB;
begin
  Result := True;
  // Contrôle du mode de paiement
  if not ExisteSQL('select MP_MODEPAIE from MODEPAIE where MP_MODEPAIE="' + ModePaie + '"') then
  begin
    if AffMsg then PGIBox('Le mode de paiement est incorrect !', Titre);
    V_PGI.IoError := oeUnknown;
    Result := False;
    Exit;
  end;
  // Constitution de l'échéance
  TOBL := TOB.Create('PIEDECHE', TOBEches, -1);
  TOBL.InitValeurs;
  MajFromCleDoc(TOBL, CleDoc);
  TOBL.PutValue('GPE_DEVISE', TOBPiece.GetValue('GP_DEVISE'));
  TOBL.PutValue('GPE_TAUXDEV', TOBPiece.GetValue('GP_TAUXDEV'));
  TOBL.PutValue('GPE_COTATION', TOBPiece.GetValue('GP_COTATION'));
  TOBL.PutValue('GPE_DATETAUXDEV', TOBPiece.GetValue('GP_DATETAUXDEV'));
  TOBL.PutValue('GPE_SAISIECONTRE', TOBPiece.GetValue('GP_SAISIECONTRE'));
  TOBL.PutValue('GPE_TIERS', TOBPiece.GetValue('GP_TIERS'));
  TOBL.PutValue('GPE_CAISSE', TOBPiece.GetValue('GP_CAISSE'));
  if RatNumZ then TOBL.PutValue('GPE_NUMZCAISSE', TOBPiece.GetValue('GP_NUMZCAISSE'));
  TOBL.PutValue('GPE_NUMECHE', 1);
  TOBL.PutValue('GPE_CHQDIFUTIL', 'X');
  TOBL.PutValue('GPE_ACOMPTE', '-');
  TOBL.PutValue('GPE_MODEPAIE', ModePaie);
  TOBL.PutValue('GPE_DATEECHE', Date);
  TOBL.PutValue('GPE_MONTANTENCAIS', MontDev);
  TOBL.PutValue('GPE_DEVISEESP', Devise);
  FOConvtMntEcheance(MontDev, Devise, TOBPiece, nil, DEV, MontantP, MontantD);
  // Montant en devise de tenu
  TOBL.PutValue('GPE_MONTANTECHE', MontantP);
  // Montant en devise de la pièce
  TOBL.PutValue('GPE_MONTANTDEV', MontantD);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InitToutModif : force l'indicateur de modification des TOBS
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOGenTicket.InitToutModif;
var NowFutur: TDateTime;
begin
  NowFutur := NowH;
  TOBPiece.SetAllModifie(True);
  TOBPiece.SetDateModif(NowFutur);
  TOBBases.SetAllModifie(True);
  TOBEches.SetAllModifie(True);
  TOBAcomptes.SetAllModifie(True);
  TOBPorcs.SetAllModifie(True);
  InvalideModifTiersPiece(TOBTiers);
  TOBAnaP.SetAllModifie(True);
  TOBAnaS.SetAllModifie(True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ValideLaPiece : Enregistrement de la piece
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOGenTicket.ValideLaPiece;
var OldEcr, OldStk: RMVT;
  {$IFDEF GESCOM}
  TOBPIECERG, TOBBASESRG: TOB;
  TobOpCaisse : TOB;
  {$ENDIF}
begin
  InitToutModif;
{$IFDEF GESCOM}
  TOBPieceRG := TOB.create('', nil, -1);
  TOBBasesRG := TOB.create('', nil, -1);
{$ENDIF}
  ValideLaCotation(PieceContainer);
  ValideLaPeriode(TOBPiece);
  if GetInfoParPiece(CleDoc.NaturePiece, 'GPP_IMPIMMEDIATE') = 'X' then TOBPiece.PutValue('GP_EDITEE', 'X');
  // Attribution du numéro de pièce
  if not FOSetNumeroDefinitif(PieceContainer) then V_PGI.IoError := oePointage;
  CleDoc.NumeroPiece := TOBPiece.GetValue('GP_NUMERO');
  if GetInfoParPiece(CleDoc.NaturePiece, 'GPP_ACTIONFINI') = 'ENR' then PutValueDetail(TOBPiece, 'GP_VIVANTE', '-');
  // Alimentation de la comptabilité
  if ActiveCompta then
  begin
    if V_PGI.IoError = oeOk then ValideAnalytiques(TOBPiece, TOBAnaP, TOBAnaS);
    if V_PGI.IoError = oeOk then
    begin
{$IFDEF GESCOM}
      if IsArtFiPce(TobPiece) then
      begin
        TobOpCaisse := TOB.Create('PIECE', nil, -1);
        TobOpCaisse.Dupliquer(TobPiece,true,true,true);
      end else
        TobOpCaisse := nil;
      if not PassationComptable(PieceContainer, DEV.Decimale, OldEcr, OldStk, True, TobOpCaisse) then
        V_PGI.IoError := oeLettrage;
      if TobOpCaisse <> nil then FreeAndNil(TobOpCaisse);
{$ELSE}
      if not PassationComptable(PieceContainer, DEV.Decimale, OldEcr, OldStk, True) then
        V_PGI.IoError := oeLettrage;
{$ENDIF GESCOM}
    end;
  end;
  // Enregistrement de la pièce
  if V_PGI.IoError = oeOk then TOBPiece.InsertDBByNivel(False);
  if V_PGI.IoError = oeOk then TOBBases.InsertDB(nil);
  if V_PGI.IoError = oeOk then TOBEches.InsertDB(nil);
  if V_PGI.IoError = oeOk then TOBAnaP.InsertDB(nil);
  if V_PGI.IoError = oeOk then TOBAnaS.InsertDB(nil);
  {$IFDEF GESCOM}
  TOBPIECERG.free;
  TOBBASESRG.free;
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ValideImpression : Impression de la pièce
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOGenTicket.ValideImpression;
var Editee: string;
begin
  // Pour ne pas imprimer la mention Duplicata
  Editee := TOBPiece.GetValue('GP_EDITEE');
  TOBPiece.PutValue('GP_EDITEE', '-');
  FOImprimerLaPiece(TOBPiece, TOBTiers, TOBEches, TOBComms, CleDoc, True, False, False);
  TOBPiece.PutValue('GP_EDITEE', Editee);
end;

{==============================================================================================}
{============================ GESTION DES PAVES CONTEXTUELS ===================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/06/2002
Modifié le ... : 28/06/2002
Description .. : Affichage de la page du pavé affectée au champ
Mots clefs ... : FO
*****************************************************************}

procedure AffichePaveContextuel(PnlBoutons: TClavierEcran; NoPage: Integer);
var sPlus: string;
begin
  if PnlBoutons = nil then Exit;
  if NoPage > 0 then
  begin
    PnlBoutons.PageCourante := NoPage - 1;
  end else
  begin
    case NoPage of
      -1: PnlBoutons.LanceChargePageMenu('OCT', '', '', clBtnFace, 13);
      -2: PnlBoutons.LanceChargePageMenu('REM', '', '', clBtnFace, 11);
      -3:
        begin
          sPlus := 'GCL_TYPECOMMERCIAL="VEN" and GCL_ETABLISSEMENT="' + FOGetParamCaisse('GPK_ETABLISSEMENT') + '"'
            + ' and GCL_DATESUPP>"' + USDateTime(Date) + '"';
          PnlBoutons.LanceChargePageMenu('VEN', '', sPlus, clBtnFace, 40);
        end;
      -4: PnlBoutons.LanceChargePageMenu('REG', '', '', clBtnFace, 26);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/03/2003
Modifié le ... : 24/03/2003
Description .. : Dimension un panel pour afficher un certain nombre de
Suite ........ : lignes d'une grille
Mots clefs ... : FO
*****************************************************************}

procedure AppliqueHauteurPanel(GS: THGrid; NbRow: integer);
var PCorps: THPanel;
begin
  if (NbRow >= 2) and (NbRow <= 22) and
    (GS.Parent <> nil) and (GS.Parent is THPanel) then
  begin
    // Un nouveau panel permet de gérer le pavé secondaire
    // PMain -> PGrille -> PnlCorps -> GS
    //                              -> PPied
    //                  -> PClavier2
    PCorps := THPanel(TWinControl(GS.Parent).Parent);
    if NbRow <> GS.VisibleRowCount then
    begin
      PCorps.Height := 84 + GS.RowHeights[0] + 2
        + ((GS.RowHeights[1] + 1) * NbRow);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/06/2002
Modifié le ... : 28/06/2002
Description .. : Charge le paramètrage du pavé et de l'écran de saisie de ticket
Mots clefs ... : FO
*****************************************************************}

procedure ChargePaveContextuel(PP: array of THPanel; GS: array of THGrid; LCD: array of TLCDLabel; var ColCarac: array of RColCarac; WC: array of TWinControl; var PropPv2: RPropPave2);
var
  sVal, Stg: string;
  Ind, Deb, NbRow: Integer;
  StgLst: TStrings;
begin
  FillChar(PropPv2, Sizeof(PropPv2), #0);
  PropPv2.Visible := False;
  if not VH_GC.TOBPCaisse.FieldExists('GPK_INFOS') then Exit;
  StgLst := TStringList.Create;
  try
    StgLst.Text := VH_GC.TOBPCaisse.GetValue('GPK_INFOS');
  except
    StgLst.Free;
    Exit;
  end;
  with StgLst do
  begin
    // propriétés des panels
    sVal := Values['PANEL'];
    if sVal <> '' then
    begin
      Deb := Low(PP);
      PP[Deb].ColorStart := TColor(StrToInt(ReadToKenST(sVal)));
      PP[Deb].ColorEnd := TColor(StrToInt(ReadToKenST(sVal)));
      PP[Deb].BackGroundEffect := TDirection(StrToInt(ReadToKenST(sVal)));
      if PP[Deb].BackGroundEffect = bdFond then PP[Deb].ColorNB := 6
      else PP[Deb].ColorNB := 100;
      for Ind := (Deb + 1) to High(PP) do
      begin
        PP[Ind].ColorStart := PP[Deb].ColorStart;
        PP[Ind].ColorEnd := PP[Deb].ColorEnd;
        PP[Ind].BackGroundEffect := PP[Deb].BackGroundEffect;
        PP[Ind].ColorNB := PP[Deb].ColorNB;
      end;
    end;
    if GetParamSoc('SO_GCFOBASEFORMATION') then
    begin
      for Ind := Low(PP) to High(PP) do
      begin
        PP[Ind].BackGroundEffect := bdFlat;
        PP[Ind].Color := clRed;
      end;
    end;
    // propriétés des LCD
    sVal := Values['LCD'];
    if sVal <> '' then
    begin
      Deb := Low(LCD);
      LCD[Deb].BackGround := TColor(StrToInt(ReadToKenST(sVal)));
      LCD[Deb].PixelOn := TColor(StrToInt(ReadToKenST(sVal)));
      LCD[Deb].PixelOff := TColor(StrToInt(ReadToKenST(sVal)));
      for Ind := (Deb + 1) to High(LCD) do
      begin
        LCD[Ind].BackGround := LCD[Deb].BackGround;
        LCD[Ind].PixelOn := LCD[Deb].PixelOn;
        LCD[Ind].PixelOff := LCD[Deb].PixelOff;
      end;
    end;
    if GetParamSoc('SO_GCFOBASEFORMATION') then
    begin
      for Ind := Low(LCD) to High(LCD) do
      begin
        LCD[Ind].PixelOn := clRed;
        LCD[Ind].PixelOff := clBlack;
      end;
    end;
    // propriétés des grilles
    sVal := Values['GRID'];
    if sVal <> '' then
    begin
      Deb := Low(GS);
      GS[Deb].FixedColor := TColor(StrToInt(ReadToKenST(sVal)));
      GS[Deb].TwoColors := (ReadToKenST(sVal) = 'X');
      for Ind := (Deb + 1) to High(GS) do
      begin
        GS[Ind].FixedColor := GS[Deb].FixedColor;
        GS[Ind].TwoColors := GS[Deb].TwoColors;
      end;
      Stg := ReadToKenST(sVal);
      if IsNumeric(Stg) then NbRow := StrToInt(Stg) else NbRow := 0;
      AppliqueHauteurPanel(GS[Deb], NbRow);
    end else
      if not FOExisteClavierEcran then
    begin
      Deb := Low(GS);
      AppliqueHauteurPanel(GS[Deb], 12);
    end;
    // ordre de tabulation en-tête
    sVal := Values['ETABORD'];
    if sVal <> '' then
    begin
      for Ind := Low(WC) to High(WC) do
      begin
        WC[Ind].TabOrder := StrToInt(ReadToKenST(sVal));
      end;
    end;
    // choix des pavés
    for Ind := Low(ColCarac) to High(ColCarac) do
    begin
      sVal := Values['P' + IntToStr(Ind)];
      if sVal <> '' then ColCarac[Ind].NoPage := StrToInt(ReadToKenST(sVal));
    end;
    // choix des propriétés des colonnes
    for Ind := Low(ColCarac) to High(ColCarac) do
    begin
      sVal := Values['C' + IntToStr(Ind)];
      if sVal <> '' then
      begin
        ColCarac[Ind].Largeur := Round(StrToInt(ReadToKenST(sVal)) * 6.3);
        ColCarac[Ind].Align := TAlignment(StrToInt(ReadToKenST(sVal)));
      end;
    end;
    for Ind := Low(GS) to High(GS) do FOAppliqueColCarac(GS[Ind], ColCarac, (Ind * 20));
    // propriétés du pavé secondaire
    sVal := Values['PAVE2'];
    if sVal <> '' then
    begin
      PropPv2.Visible := (ReadToKenST(sVal) = 'X');
      PropPv2.Largeur := StrToInt(ReadToKenST(sVal));
      PropPv2.Align := TAlignment(StrToInt(ReadToKenST(sVal)));
      PropPv2.NbrBtnWidth := StrToInt(ReadToKenST(sVal));
      PropPv2.NbrBtnHeight := StrToInt(ReadToKenST(sVal));
      PropPv2.ClcPosition:= TPosClc(StrToInt(ReadToKenST(sVal)));
      PropPv2.ClcVisible:= (ReadToKenST(sVal) <> '-');
    end;
end;
  StgLst.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/06/2002
Modifié le ... : 23/07/2002
Description .. : Applique sur une grille les caractéristiques des colonnes (n°
Suite ........ : page du pavé, taille et alignement)
Mots clefs ... : FO
*****************************************************************}

procedure FOAppliqueColCarac(GS: THGrid; ColCarac: array of RColCarac; IndDebut: Integer);
var Ind: Integer;
begin
  for Ind := GS.FixedRows to GS.ColCount - 1 do
  begin
    if ((IndDebut + Ind) >= Low(ColCarac)) and ((IndDebut + Ind) <= High(ColCarac)) and
      (ColCarac[IndDebut + Ind].Largeur > 0) then
    begin
      GS.ColWidths[Ind] := ColCarac[IndDebut + Ind].Largeur;
      GS.ColAligns[Ind] := ColCarac[IndDebut + Ind].Align;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2002
Modifié le ... : 23/07/2002
Description .. : Initialise un tableau des caractéristiques des colonnes (n°
Suite ........ : page du pavé, taille et alignement)
Mots clefs ... : FO
*****************************************************************}

procedure FOInitColCarac(ColCarac: array of RColCarac);
var Ind: Integer;
begin
  for Ind := Low(ColCarac) to High(ColCarac) do
  begin
    ColCarac[Ind].NoPage := 0;
    ColCarac[Ind].Largeur := -1;
    ColCarac[Ind].Align := taLeftJustify;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Donne le paramétrage du pavé tactile défini pour la caisse
Suite ........ : courante
Mots clefs ... : FO
*****************************************************************}

function FODonneParamClavierEcran(ValeurParDefaut: boolean): string;
begin
  if ValeurParDefaut then
    Result := ''
  else
    Result := FOGetParamCaisse('GPK_PARAMSCE');

  if (Trim(Result) = '') or (CountTokenPipe(Result, ';') < 3) then
    Result := IntToStr(CENbrBtnWidthDef) + ';' + IntToStr(CENbrBtnHeightDef)
      + ';' + IntToStr(Ord(pcLeft)) + ';X;';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/12/2003
Modifié le ... : 03/12/2003
Description .. : Création d'un pavé tactile
Mots clefs ... : FO
*****************************************************************}

procedure FOCreateClavierEcran(var PnlBoutons: TClavierEcran; AOwner: TComponent; Pmain: TPanel; ChargeAuto: boolean = True; NbBtnDefaut: boolean = False);
var
  Stg: string;
begin
  if Assigned(PnlBoutons) then FreeAndNil(PnlBoutons);
  PnlBoutons := TClavierEcran.Create(AOwner);
  PnlBoutons.Parent := Pmain;
  PnlBoutons.Align := alBottom;
  PnlBoutons.Height := Pmain.Tag;
  PnlBoutons.colornb := 6;
  PnlBoutons.ChargeAut := ChargeAuto;
  PnlBoutons.LanceBouton := nil;
  PnlBoutons.LanceCalculette := nil;
  PnlBoutons.BoutonCalculette := nil;
  PnlBoutons.Caisse := FOCaisseCourante;

  Stg := FODonneParamClavierEcran(False);
  if NbBtnDefaut then
  begin
    PnlBoutons.NbrBtnWidth := CENbrBtnWidthDef;
    PnlBoutons.NbrBtnHeight := CENbrBtnHeightDef;
    PnlBoutons.ClcPosition := pcLeft;
    PnlBoutons.ClcVisible := True;
  end else
  begin
    PnlBoutons.NbrBtnWidth := ValeurI(ReadTokenSt(Stg));
    PnlBoutons.NbrBtnHeight := ValeurI(ReadTokenSt(Stg));
    PnlBoutons.ClcPosition := tPosClc(ValeurI(ReadTokenSt(Stg)));
    PnlBoutons.ClcVisible := (ReadTokenSt(Stg) <> '-');
  end;
end;

{==============================================================================================}
{========================= GESTION DES LIAISONS DES REGLEMENTS ================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/04/2003
Modifié le ... : 18/04/2003
Description .. : Indique si un mode de paiement ou un article financier 
Suite ........ : possède un numéro de bon
Mots clefs ... :
*****************************************************************}

function FOIsOpCaisseNoBon(Prefixe, TypeOpe: string; Montant: double): boolean;
begin
  Result := False;
  if Prefixe = 'GPE' then
  begin
    // cas des échéances => modes de paiement
    if (TypeOpe = TYPEPAIEAVOIR) and (Montant < 0) then // Emission d'avoir
      Result := True;
  end else
    if Prefixe = 'GL' then
  begin
    // cas des lignes => articles financiers
    if TypeOpe = TYPEOPCBONACHAT then // Acquisition d'un bon d'achat
      Result := True
    else if TypeOpe = TYPEOPCARRHES then // Versement d'arrhes
      Result := True;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/04/2003
Modifié le ... : 18/04/2003
Description .. : Indique si un mode de paiement ou un article financier est
Suite ........ : un cumul pour la liaison des règlements
Mots clefs ... :
*****************************************************************}

function FOIsOpCaisseCumul(Prefixe, TypeOpe: string; Montant: double): boolean;
begin
  Result := False;
  if Prefixe = 'GPE' then
  begin
    // cas des échéances => modes de paiement
    if (TypeOpe = TYPEPAIEAVOIR) and (Montant < 0) then // Emission d'avoir
      Result := True
    else if TypeOpe = TYPEPAIERESTEDU then // Reste dû
      Result := True;
  end else
    if Prefixe = 'GL' then
  begin
    // cas des lignes => articles financiers
    if TypeOpe = TYPEOPCBONACHAT then // Acquisition d'un bon d'achat
      Result := True
    else if TypeOpe = TYPEOPCARRHES then // Versement d'arrhes
      Result := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/04/2003
Modifié le ... : 18/04/2003
Description .. : Indique si un mode de paiement ou un article financier est
Suite ........ : concerné par la liaison des règlements
Mots clefs ... :
*****************************************************************}

function FOIsOpCaisseLiee(Prefixe, TypeOpe: string; Montant: double): boolean;
begin
  Result := False;
  if Prefixe = 'GPE' then
  begin
    // cas des modes de paiement
    if (TypeOpe = TYPEPAIEAVOIR) and (Montant > 0) then // Avoir
      Result := True
    else if TypeOpe = TYPEPAIEBONACHAT then // Bon d'achat
      Result := True
    else if TypeOpe = TYPEPAIEARRHES then // Arrhes déjà versés
      Result := True;
  end else
    if Prefixe = 'GL' then
  begin
    // cas des opérations de caisse
    if TypeOpe = TYPEOPCENCCREDIT then // Encaissement de crédit
      Result := True
    else if TypeOpe = TYPEOPCREMBAVOIR then // Remboursement d'avoir
      Result := True
    else if TypeOpe = TYPEOPCREMBBONACH then // Remboursement de bon d'achat
      Result := True
    else if TypeOpe = TYPEOPCREMBARRHES then // Remboursement d'arrhes
      Result := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/04/2003
Modifié le ... : 14/04/2003
Description .. : Copie une ligne vers une opération de caisse
Mots clefs ... :
*****************************************************************}

function TOBLigneVersOpCaisse(TOBLigne, TOBOpCais: TOB; sType: string; NumZCaisse : integer; Caisse : string): TOB;
var
  TOBL: TOB;
  RefPiece: string;
  Montant: double;
begin
  Result := nil;
  if (TOBLigne = nil) or (TOBLigne = nil) then Exit;
  Montant := TOBLigne.GetValue('GL_TOTALTTC');
  TOBL := TOB.Create('OPERCAISSE', TOBOpCais, -1);
  Result := TOBL;
  TOBL.InitValeurs;
  TOBL.PutValue('GOC_NATUREPIECEG', TOBLigne.GetValue('GL_NATUREPIECEG'));
  TOBL.PutValue('GOC_DATEPIECE', TOBLigne.GetValue('GL_DATEPIECE'));
  TOBL.PutValue('GOC_SOUCHE', TOBLigne.GetValue('GL_SOUCHE'));
  TOBL.PutValue('GOC_NUMERO', TOBLigne.GetValue('GL_NUMERO'));
  TOBL.PutValue('GOC_INDICEG', TOBLigne.GetValue('GL_INDICEG'));
  TOBL.PutValue('GOC_PREFIXE', 'GL');
  TOBL.PutValue('GOC_NUMLIGNE', TOBLigne.GetValue('GL_NUMLIGNE'));
  TOBL.PutValue('GOC_NUMORDRE', TOBLigne.GetValue('GL_NUMORDRE'));
  TOBL.PutValue('GOC_ARTICLE', TOBLigne.GetValue('GL_ARTICLE'));
  TOBL.PutValue('GOC_TIERS', TOBLigne.GetValue('GL_TIERS'));
  TOBL.PutValue('GOC_ETABLISSEMENT', TOBLigne.GetValue('GL_ETABLISSEMENT'));
  TOBL.PutValue('GOC_TOTALTTC', Montant);
  TOBL.PutValue('GOC_TOTALTTCDEV', TOBLigne.GetValue('GL_TOTALTTCDEV'));
  TOBL.PutValue('GOC_DEVISE', TOBLigne.GetValue('GL_DEVISE'));
  if TOBLigne.FieldExists('GOC_NUMPIECELIEN') then
  begin
    RefPiece := TOBLigne.GetValue('GOC_NUMPIECELIEN');
    TOBL.PutValue('GOC_NUMPIECELIEN', RefPiece);
  end else
    RefPiece := '';
  if (FOIsOpCaisseCumul('GL', sType, Montant)) and (RefPiece = '') then
    TOBL.PutValue('GOC_TOTALDISPOTTC', Montant);
  if TOBLigne.FieldExists('GOC_NUMBON') then
    TOBL.PutValue('GOC_NUMBON', TOBLigne.GetValue('GOC_NUMBON'));
  // JTR
  TOBL.PutValue('GOC_NUMZCAISSE', NumZCaisse);
  TOBL.PutValue('GOC_CAISSE', Caisse);
  if TOBLigne.FieldExists('GOC_ANTERIEUR') then
    TOBL.PutValue('GOC_ANTERIEUR', TOBLigne.GetValue('GOC_ANTERIEUR'))
    else
    TOBL.PutValue('GOC_ANTERIEUR', '-');
  // Fin JTR
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/04/2003
Modifié le ... : 14/04/2003
Description .. : Copie une échéance vers une opération de caisse
Mots clefs ... :
*****************************************************************}

function TOBEcheanceVersOpCaisse(TOBEche, TOBOpCais: TOB; sType, Etablis: string; NumZCaisse : integer; Caisse : string): TOB;
var TOBL: TOB;
  RefPiece: string;
  Montant: double;
begin
  Result := nil;
  if (TOBOpCais = nil) or (TOBEche = nil) then Exit;
  Montant := TOBEche.GetValue('GPE_MONTANTECHE');
  TOBL := TOB.Create('OPERCAISSE', TOBOpCais, -1);
  Result := TOBL;
  TOBL.InitValeurs;
  TOBL.PutValue('GOC_NATUREPIECEG', TOBEche.GetValue('GPE_NATUREPIECEG'));
  TOBL.PutValue('GOC_DATEPIECE', TOBEche.GetValue('GPE_DATEPIECE'));
  TOBL.PutValue('GOC_SOUCHE', TOBEche.GetValue('GPE_SOUCHE'));
  TOBL.PutValue('GOC_NUMERO', TOBEche.GetValue('GPE_NUMERO'));
  TOBL.PutValue('GOC_INDICEG', TOBEche.GetValue('GPE_INDICEG'));
  TOBL.PutValue('GOC_PREFIXE', 'GPE');
  TOBL.PutValue('GOC_NUMLIGNE', TOBEche.GetValue('GPE_NUMECHE'));
  TOBL.PutValue('GOC_NUMORDRE', TOBEche.GetValue('GPE_NUMECHE'));
  TOBL.PutValue('GOC_MODEPAIE', TOBEche.GetValue('GPE_MODEPAIE'));
  TOBL.PutValue('GOC_TIERS', TOBEche.GetValue('GPE_TIERS'));
  TOBL.PutValue('GOC_ETABLISSEMENT', Etablis);
  TOBL.PutValue('GOC_TOTALTTC', Montant);
  TOBL.PutValue('GOC_TOTALTTCDEV', TOBEche.GetValue('GPE_MONTANTDEV'));
  TOBL.PutValue('GOC_DEVISE', TOBEche.GetValue('GPE_DEVISE'));
  if TOBEche.FieldExists('GOC_NUMPIECELIEN') then
  begin
    RefPiece := TOBEche.GetValue('GOC_NUMPIECELIEN');
    TOBL.PutValue('GOC_NUMPIECELIEN', RefPiece);
  end else
    RefPiece := '';
  if (FOIsOpCaisseCumul('GPE', sType, Montant)) and (RefPiece = '') then
    TOBL.PutValue('GOC_TOTALDISPOTTC', Montant);
  if TOBEche.FieldExists('GOC_NUMBON') then
    TOBL.PutValue('GOC_NUMBON', TOBEche.GetValue('GOC_NUMBON'));
  TOBL.PutValue('GOC_NUMZCAISSE', NumZCaisse);
  TOBL.PutValue('GOC_CAISSE',Caisse); 
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/04/2003
Modifié le ... : 18/04/2003
Description .. : Fabrique l'ordre SQL de mise à jour du total dispo du
Suite ........ : règlement lié
Mots clefs ... :
*****************************************************************}

function MAJTotalDispo(Prefixe, TypeOpe: string; Montant: double; ClauseWhere, AnnulPiece: boolean): string;
var
  Signe, Operateur: string;
begin
  Result := '';
  if AnnulPiece then
  begin
    // cas de l'annulation de ticket
    if not ClauseWhere then Result := ',GOC_TOTALDISPOTTC=0';
    Exit;
  end;
  if Prefixe = 'GPE' then
  begin
    // cas des échéances => modes de paiement
    if (TypeOpe = TYPEPAIEAVOIR) and (Montant >= 0) then // Avoir
    begin
      Signe := '+';
      Operateur := '<=';
    end else
      if TypeOpe = TYPEPAIEBONACHAT then // Bon d'achat
    begin
      Signe := '-';
      Operateur := '>=';
    end else
      if TypeOpe = TYPEPAIEARRHES then // Arrhes déjà versés
    begin
      Signe := '-';
      Operateur := '>=';
    end;
  end else
    if Prefixe = 'GL' then
  begin
    // cas des lignes => articles financiers
    if TypeOpe = TYPEOPCENCCREDIT then // Encaissement de crédit
    begin
      Signe := '-';
      Operateur := '>=';
    end else
      if TypeOpe = TYPEOPCREMBAVOIR then // Remboursement d'avoir
    begin
      Signe := '-';
      Operateur := '<=';
    end else
      if TypeOpe = TYPEOPCREMBBONACH then // Remboursement de bon d'achat
    begin
      Signe := '+';
      Operateur := '>=';
    end else
      if TypeOpe = TYPEOPCREMBARRHES then // Remboursement d'arrhes
    begin
      Signe := '+';
      Operateur := '>=';
    end;
  end;
  if ClauseWhere then
    Result := ' AND (GOC_TOTALDISPOTTC' + Signe + '(' + StrFPoint(Montant) + '))' + Operateur + '0'
  else
    Result := ',GOC_TOTALDISPOTTC=GOC_TOTALDISPOTTC' + Signe + '(' + StrFPoint(Montant) + ')';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/04/2003
Modifié le ... : 04/07/2003
Description .. : Enregistrement des opérations de caisse à partir des lignes 
Suite ........ : de la pièce
Mots clefs ... : 
*****************************************************************}

procedure EnregOperCaisseLigne(TOBOpCais, TOBPiece, TOBArticles, TOBPiece_O: TOB; AnnulPiece: boolean);
var
  TOBL, TOBA, TOBC: TOB;
  RefArt, RefLigne, Prfx, sSql, TypeOpe, Stg: string;
  Montant, Montant_O: double;
  CD, CD_O: R_CleDoc;
  NowFutur: TDateTime;
  NumZCaisse : integer; // JTR
  Caisse : string; //JTR
begin
  if (TOBOpCais = nil) or (TOBPiece = nil) or (TOBArticles = nil) then Exit;
  NowFutur := NowH;
  TOBC := nil;

  // Recherche des lignes avec un article financier
  NumZCaisse := TOBPiece.GetValue('GP_NUMZCAISSE');
  Caisse := TOBPiece.GetValue('GP_CAISSE');
  TOBL := TOBPiece.FindFirst(['GL_TYPEARTICLE'], ['FI'], False);
  while TOBL <> nil do
  begin
    RefArt := TOBL.GetValue('GL_ARTICLE');
    TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefArt], False);
    TypeOpe := TOBA.GetValue('GA_TYPEARTFINAN');
    Montant := TOBL.GetValue('GL_TOTALTTC');
    if AnnulPiece then Montant_O := Montant  * -1 else Montant_O := Montant;

    // mise à jour de la table des opérations de caisse
    if (FOIsOpCaisseCumul('GL', TypeOpe, Montant_O)) or
       (FOIsOpCaisseLiee('GL', TypeOpe, Montant_O)) then
      TOBC := TOBLigneVersOpCaisse(TOBL, TOBOpCais, TypeOpe, NumZCaisse, Caisse);

    // mise à jour du réglement lié
    RefLigne := FOGetChampSupValeur(TOBL, 'GOC_NUMPIECELIEN', '');
    if (RefLigne <> '') and (Montant <> 0) then
    begin
      Prfx := ReadTokenSt(RefLigne);
      DecodeRefPiece(RefLigne, CD);
      if (AnnulPiece) and (TOBPiece_O <> nil) and (FOIsOpCaisseLiee('GL', TypeOpe, Montant_O)) then
      begin
        // réactivation du règlement d'origine
        sSql := 'UPDATE OPERCAISSE'
              + ' SET GOC_NUMPIECELIEN="",'
              + ' GOC_DATEMODIF="' + USTime(NowFutur) + '",'
              + ' GOC_TOTALDISPOTTC=GOC_TOTALTTC'
              + ' WHERE GOC_NATUREPIECEG="' + CD.NaturePiece + '"'
              + ' AND GOC_SOUCHE="' + CD.Souche + '"'
              + ' AND GOC_NUMERO=' + IntToStr(CD.NumeroPiece)
              + ' AND GOC_INDICEG=' + IntToStr(CD.Indice)
              + ' AND GOC_PREFIXE="'+ Prfx +'"'
              + ' AND GOC_NUMORDRE=' + IntToStr(CD.NumOrdre);

        if V_PGI.IoError = oeOk then
          if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;

        // mise à jour du règlement de la pièce annulée
        CD_O := TOB2CleDoc(TOBPiece_O);
        CD_O.NumLigne := TOBL.GetValue('GL_NUMLIGNE');
        CD_O.NumOrdre := TOBL.GetValue('GL_NUMORDRE');
        sSql := 'UPDATE OPERCAISSE'
              + ' SET GOC_NUMPIECELIEN="GL;' + EncodeRefPiece(TOBL) + '",'
              + ' GOC_DATEMODIF="' + USTime(NowFutur) + '"'
              + MAJTotalDispo('GL', TypeOpe, Montant, False, AnnulPiece)
              + ' WHERE GOC_NATUREPIECEG="' + CD_O.NaturePiece + '"'
              + ' AND GOC_SOUCHE="' + CD_O.Souche + '"'
              + ' AND GOC_NUMERO=' + IntToStr(CD_O.NumeroPiece)
              + ' AND GOC_INDICEG=' + IntToStr(CD_O.Indice)
              + ' AND GOC_PREFIXE="GL"'
              + ' AND GOC_NUMORDRE=' + IntToStr(CD_O.NumOrdre)
              + MAJTotalDispo('GL', TypeOpe, Montant, True, AnnulPiece);

        if V_PGI.IoError = oeOk then
          if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;

        // mise à jour du règlement de la pièce d'annulation
        if TOBC <> nil then
        begin
          Stg := 'GL;' + EncodeRefPiece(TOBPiece_O) + IntToStr(CD_O.NumOrdre) + ';';
          FOMajChampSupValeur(TOBC, 'GOC_NUMPIECELIEN', Stg);
        end;
      end else
      begin
        sSql := 'UPDATE OPERCAISSE'
              + ' SET GOC_NUMPIECELIEN="GL;' + EncodeRefPiece(TOBL) + '",'
              + ' GOC_DATEMODIF="' + USTime(NowFutur) + '"'
              + MAJTotalDispo('GL', TypeOpe, Montant, False, AnnulPiece)
              + ' WHERE GOC_NATUREPIECEG="' + CD.NaturePiece + '"'
              + ' AND GOC_SOUCHE="' + CD.Souche + '"'
              + ' AND GOC_NUMERO=' + IntToStr(CD.NumeroPiece)
              + ' AND GOC_INDICEG=' + IntToStr(CD.Indice)
              + ' AND GOC_PREFIXE="'+ Prfx +'"'
              + ' AND GOC_NUMORDRE=' + IntToStr(CD.NumOrdre)
              + MAJTotalDispo('GL', TypeOpe, Montant, True, AnnulPiece);

        if V_PGI.IoError = oeOk then
          if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;
      end;
    end;

    TOBL := TOBPiece.FindNext(['GL_TYPEARTICLE'], ['FI'], False);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/04/2003
Modifié le ... : 04/07/2003
Description .. : Enregistrement des opérations de caisse à partir des
Suite ........ : échéances de la pièce
Mots clefs ... :
*****************************************************************}

procedure EnregOperCaisseEcheance(TOBOpCais, TOBPiece, TOBEches, TOBPiece_O: TOB; AnnulPiece: boolean);
var
  TOBL, TOBC : TOB;
  Etablis, RefLigne, RefPiece, Prfx, sSql, TypeOpe, Stg: string;
  Montant, Montant_O: double;
  CD, CD_O: R_CleDoc;
  NowFutur, DD: TDateTime;
  Ind, iNumEche, NumZCaisse: integer;
  Caisse : string; // JTR
begin
  if (TOBOpCais = nil) or (TOBPiece = nil) or (TOBEches = nil) then Exit;
  NowFutur := NowH;
  TOBC := nil;

  for Ind := 0 to TOBEches.Detail.Count - 1 do
  begin
    sSql := '';
    TOBL := TOBEches.Detail[Ind];

    Etablis := TOBPiece.GetValue('GP_ETABLISSEMENT');
    NumZCaisse := TOBPiece.GetValue('GP_NUMZCAISSE'); // JTR
    Caisse := TOBPiece.GetValue('GP_CAISSE'); // JTR
    TypeOpe := TOBL.GetValue('TYPEMODEPAIE');
    Montant := TOBL.GetValue('GPE_MONTANTECHE');
    if AnnulPiece then Montant_O := Montant  * -1 else Montant_O := Montant;

    // mise à jour de la table des opérations de caisse
    if (FOIsOpCaisseCumul('GPE', TypeOpe, Montant_O)) or
       (FOIsOpCaisseLiee('GPE', TypeOpe, Montant_O)) then
      TOBC := TOBEcheanceVersOpCaisse(TOBL, TOBOpCais, TypeOpe, Etablis, NumZCaisse, Caisse);

    // mise à jour du réglement lié
    RefLigne := FOGetChampSupValeur(TOBL, 'GOC_NUMPIECELIEN', '');
    if (RefLigne <> '') and (Montant <> 0) then
    begin
      // les n° d'échéance sont identiques entre la pièce annulée et la pièce d'annulation
      iNumEche := TOBL.GetValue('GPE_NUMECHE');
      // Référence de la ligne d'échéance du ticket courant
      DD := TOBL.GetValue('GPE_DATEPIECE');
      RefPiece := FormatDateTime('ddmmyyyy', DD) + ';'
        + TOBL.GetValue('GPE_NATUREPIECEG') + ';'
        + TOBL.GetValue('GPE_SOUCHE') + ';'
        + IntToStr(TOBL.GetValue('GPE_NUMERO')) + ';'
        + IntToStr(TOBL.GetValue('GPE_INDICEG')) + ';'
        + IntToStr(iNumEche) + ';';
      Prfx := ReadTokenSt(RefLigne);
      DecodeRefPiece(RefLigne, CD);
      if (AnnulPiece) and (TOBPiece_O <> nil) and (FOIsOpCaisseLiee('GPE', TypeOpe, Montant_O)) then
      begin
        // réactivation du règlement d'origine
        sSql := 'UPDATE OPERCAISSE'
              + ' SET GOC_NUMPIECELIEN="",'
              + ' GOC_DATEMODIF="' + USTime(NowFutur) + '",'
              + ' GOC_TOTALDISPOTTC=GOC_TOTALTTC'
              + ' WHERE GOC_NATUREPIECEG="' + CD.NaturePiece + '"'
              + ' AND GOC_SOUCHE="' + CD.Souche + '"'
              + ' AND GOC_NUMERO=' + IntToStr(CD.NumeroPiece)
              + ' AND GOC_INDICEG=' + IntToStr(CD.Indice)
              + ' AND GOC_PREFIXE="'+ Prfx +'"'
              + ' AND GOC_NUMORDRE=' + IntToStr(CD.NumOrdre);

        if V_PGI.IoError = oeOk then
          if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;

        // mise à jour du règlement de la pièce annulée
        CD_O := TOB2CleDoc(TOBPiece_O);
        sSql := 'UPDATE OPERCAISSE'
              + ' SET GOC_NUMPIECELIEN="GPE;' + RefPiece + '",'
              + ' GOC_DATEMODIF="' + USTime(NowFutur) + '"'
              + MAJTotalDispo('GPE', TypeOpe, Montant, False, AnnulPiece)
              + ' WHERE GOC_NATUREPIECEG="' + CD_O.NaturePiece + '"'
              + ' AND GOC_SOUCHE="' + CD_O.Souche + '"'
              + ' AND GOC_NUMERO=' + IntToStr(CD_O.NumeroPiece)
              + ' AND GOC_INDICEG=' + IntToStr(CD_O.Indice)
              + ' AND GOC_PREFIXE="GPE"'
              + ' AND GOC_NUMORDRE=' + IntToStr(iNumEche)
              + MAJTotalDispo('GPE', TypeOpe, Montant, True, AnnulPiece);

        if V_PGI.IoError = oeOk then
          if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;

        // mise à jour du règlement de la pièce d'annulation
        if TOBC <> nil then
        begin
          Stg := 'GPE;' + EncodeRefPiece(TOBPiece_O) + IntToStr(iNumEche) + ';';
          FOMajChampSupValeur(TOBC, 'GOC_NUMPIECELIEN', Stg);
        end;
      end else
      begin
        sSql := 'UPDATE OPERCAISSE'
              + ' SET GOC_NUMPIECELIEN="GPE;' + RefPiece + '",'
              + ' GOC_DATEMODIF="' + USTime(NowFutur) + '"'
              + MAJTotalDispo('GPE', TypeOpe, Montant, False, AnnulPiece)
              + ' WHERE GOC_NATUREPIECEG="' + CD.NaturePiece + '"'
              + ' AND GOC_SOUCHE="' + CD.Souche + '"'
              + ' AND GOC_NUMERO=' + IntToStr(CD.NumeroPiece)
              + ' AND GOC_INDICEG=' + IntToStr(CD.Indice)
              + ' AND GOC_PREFIXE="'+ Prfx +'"'
              + ' AND GOC_NUMORDRE=' + IntToStr(CD.NumOrdre)
              + MAJTotalDispo('GPE', TypeOpe, Montant, True, AnnulPiece);

        if V_PGI.IoError = oeOk then
          if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/04/2003
Modifié le ... : 14/04/2003
Description .. : Enregistrement des opérations de caisse
Mots clefs ... :
*****************************************************************}

procedure FOValideLesOperCaisse(TOBPiece, TOBArticles, TOBEches, TOBPiece_O: TOB; AnnulPiece: boolean);
var
  TOBOpCais: TOB ;
begin
  TOBOpCais := TOB.Create('', nil, -1);

  if V_PGI.IoError = oeOk then
    EnregOperCaisseLigne(TOBOpCais, TOBPiece, TOBArticles, TOBPiece_O, AnnulPiece);
  if V_PGI.IoError = oeOk then
    EnregOperCaisseEcheance(TOBOpCais, TOBPiece, TOBEches, TOBPiece_O, AnnulPiece);

  // création des opérations de caisse
  if (TOBOpCais.Detail.Count > 0) and (V_PGI.IoError = oeOk) then
  begin
    TOBOpCais.SetAllModifie(True);
    if not TOBOpCais.InsertDB(nil) then V_PGI.IoError := oeUnknown;
  end;
  TOBOpCais.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/04/2003
Modifié le ... : 22/04/2003
Description .. : Chargement des opérations de caisse d'une pièce
Mots clefs ... :
*****************************************************************}

procedure FOLoadOperCaisse(TOBPiece, TOBEches, TOBOpCais: TOB);
var
  TOBL, TOBO: TOB;
  sSql, sPrfx: string;
  Ind, NumLig: integer;
  QQ: TQuery;
begin
  if TOBOpCais = nil then Exit;
  if (TOBPiece = nil) and (TOBEches = nil) then Exit;

  // Lecture des opérations de caisse
  sSql := 'SELECT * FROM OPERCAISSE'
    + ' WHERE GOC_NATUREPIECEG="' + TOBPiece.GetValue('GP_NATUREPIECEG') + '"'
    + ' AND GOC_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '"'
    + ' AND GOC_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO'))
    + ' AND GOC_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG'));
  if TOBPiece = nil then
    // lecture uniquement des opérations liées à une échéance
    sSql := sSql + ' AND GOC_PREFIXE="GPE"'
  else if TOBEches = nil then
    // lecture uniquement des opérations liées à une ligne
    sSql := sSql + ' AND GOC_PREFIXE="GL"';

  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then
    TOBOpCais.LoadDetailDB('OPERCAISSE', '', '', QQ, False);
  Ferme(QQ);

  // Fusion des opérations de caisse avec les lignes ou les échéances
  for Ind := 0 to TOBOpCais.Detail.Count - 1 do
  begin
    TOBO := TOBOpCais.Detail[Ind];
    TOBL := nil;
    NumLig := TOBO.GetValue('GOC_NUMORDRE');
    sPrfx := TOBO.GetValue('GOC_PREFIXE');
    if (sPrfx = 'GL') and (TOBPiece <> nil) then
      TOBL := TOBPiece.FindFirst(['GL_NUMORDRE'], [NumLig], False)
    else if (sPrfx = 'GPE') and (TOBEches <> nil) then
      TOBL := TOBEches.FindFirst(['GPE_NUMECHE'], [NumLig], False);

    if TOBL <> nil then
    begin
      FOMajChampSupValeur(TOBL, 'GOC_NUMPIECELIEN', TOBO.GetValue('GOC_NUMPIECELIEN'));
      FOMajChampSupValeur(TOBL, 'GOC_NUMBON', TOBO.GetValue('GOC_NUMBON'));
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 07/07/2003
Modifié le ... : 07/07/2003
Description .. : Lancement de la fiche de recherche des opérations de 
Suite ........ : caisse disponibles.
Suite ........ : 
Suite ........ : Le code retour de la fiche est
Suite ........ : Montant;Prefixe;Date;Nature;Souche;Numéro;Indice;N°
Suite ........ : ordre
Mots clefs ... :
*****************************************************************}

function LanceRechercheOpCaisseDispo (CodeFiche, Range, Prefixe, TypeOpe : string; TOBLigne, TOBPiece, TOBEches: TOB; var Montant: double; var EnErreur: boolean): boolean;
var
  Stg, sRetour, sPrfx, sDate: string;
  DatePiece: TDateTime;
  dVal: double;
  // JTR - RD encaissés
  NumZCaisse, NumZGP : integer;
  DateGP : TDateTime;
  Qry : TQuery;
  stNature, stSouche : string;
  iNum, iIndice, iOrdre : integer;
begin
  EnErreur := False;
  Result := False;
  sRetour := AglLanceFiche('MFO', CodeFiche, Range, '', '');
//  NumZCaisse := Valeuri(ReadTokenST(sRetour)); // JTR - RD encaissés
  Stg := ReadTokenST(sRetour);
  if IsNumeric(Stg) then
  begin
    if Prefixe = 'GPE' then
    begin
     if TypeOpe = TYPEPAIEAVOIR then
        dVal := Valeur(Stg) * (-1)
      else
        dVal := Valeur(Stg);
    end else
    begin
      if TypeOpe = TYPEOPCREMBAVOIR then
        dVal := Valeur(Stg) * (-1)
      else
        dVal := Valeur(Stg);
    end ;

    if dVal > 0 then
    begin
      sPrfx := ReadTokenST(sRetour);
      sDate := ReadTokenST(sRetour);
      DatePiece := StrToDate(sDate);
      Stg := sPrfx + ';' + FormatDateTime('ddmmyyyy', DatePiece) + ';' + sRetour;
      // on vérifie qu'une même opération de caisse n'a été utilisé qu'une seule fois dans le ticket
      if (TOBPiece.FindFirst(['GOC_NUMPIECELIEN'], [Stg], False) = nil) and
         (TOBEches.FindFirst(['GOC_NUMPIECELIEN'], [Stg], False) = nil) then
      begin
        stNature := ReadTokenST(sRetour);
        stSouche := ReadTokenST(sRetour);
        iNum := StrToInt(ReadTokenSt(sRetour));
        iIndice := StrToInt(ReadTokenSt(sRetour));
        iOrdre := StrToInt(ReadTokenSt(sRetour));
        Qry := OpenSQL('SELECT GOC_NUMZCAISSE FROM OPERCAISSE'
                     + ' WHERE GOC_NATUREPIECEG = "' + stNature + '"'
                     + ' AND GOC_SOUCHE = "' + stSouche + '"'
                     + ' AND GOC_NUMERO = ' + IntToStr(iNum)
                     + ' AND GOC_INDICEG = ' + IntToStr(iIndice)
                     + ' AND GOC_PREFIXE = "' + sPrfx +'"' , True);
        if not Qry.EOF then
          NumZCaisse := Qry.FindField('GOC_NUMZCAISSE').AsInteger;
        Ferme(Qry);


        FOMajChampSupValeur(TOBLigne, 'GOC_NUMPIECELIEN', Stg);
        // JTR - RD encaissés
        DateGP := TOBPiece.GetValue('GP_DATEPIECE');
        NumZGP := TOBPiece.GetValue('GP_NUMZCAISSE');
        if (DatePiece < DateGP) or ((DatePiece = DateGP) and (NumZCaisse < NumZGP)) then
          FOMajChampSupValeur(TOBLigne, 'GOC_ANTERIEUR', 'X');
        // Fin JTR - RD encaissés
        Result := True;
        Montant := dVal;
      end else
      begin
        EnErreur := True;
        if sPrfx = 'GL' then
          PGIError('Cette opération de caisse est déjà utilisée.')
        else
          PGIError('Ce règlement est déjà utilisé.');
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 15/04/2003
Modifié le ... : 18/04/2003
Description .. : Recherche des opérations de caisse disponibles pour un
Suite ........ : règlement
Mots clefs ... :
*****************************************************************}

function FORechercheOpCaisseDispo(TOBLigne, TOBPiece, TOBEches: TOB; TypeMode: string; MntReste: double; var Montant: double): boolean;
var
  Stg: string;
  EnErreur: boolean;
begin
  Result := True;
  Montant := 0;
  if TOBLigne = nil then Exit;
  if FOGetParamCaisse('GPK_LIAISONREG') <> 'X' then Exit;

  EnErreur := False;
  if (TypeMode = TYPEPAIEAVOIR) and (MntReste > 0) then
  begin
    // Avoir => Emission d'avoir
    Stg := 'MP_TYPEMODEPAIE=' + TYPEPAIEAVOIR + ';'
      + 'XX_WHERE=GOC_TOTALDISPOTTC<0;'
      + 'GOC_NATUREPIECEG=' + TOBLigne.GetValue('GPE_NATUREPIECEG') + ';'
      + 'GOC_PREFIXE=GPE;'
      + 'GOC_TIERS=' + TOBLigne.GetValue('GPE_TIERS') + ';';
    Result := LanceRechercheOpCaisseDispo('MFOLIENECHE_MUL', Stg, 'GPE', TypeMode, TOBLigne, TOBPiece, TOBEches, Montant, EnErreur);
  end else
    if ((TypeMode = TYPEPAIEBONACHAT) or (TypeMode = TYPEPAIEARRHES)) and
    (MntReste > 0) then
  begin
    if TypeMode = TYPEPAIEBONACHAT then
      // Bon d'achat => Acquisition d'un bon d'achat
      Stg := 'GA_TYPEARTFINAN=' + TYPEOPCBONACHAT + ';'
    else
      // Arrhes déjà versés => Versement d'arrhes
      Stg := 'GA_TYPEARTFINAN=' + TYPEOPCARRHES + ';';
    Stg := Stg + 'XX_WHERE=GOC_TOTALDISPOTTC>0;'
      + 'GOC_NATUREPIECEG=' + TOBLigne.GetValue('GPE_NATUREPIECEG') + ';'
      + 'GOC_PREFIXE=GL;'
      + 'GOC_TIERS=' + TOBLigne.GetValue('GPE_TIERS') + ';';
    Result := LanceRechercheOpCaisseDispo('MFOLIENOPCAIS_MUL', Stg, 'GPE', TypeMode, TOBLigne, TOBPiece, TOBEches, Montant, EnErreur);
  end else
  begin
    // pas soumis à la liaison des règlements
    Result := True;
    FOMajChampSupValeur(TOBLigne, 'GOC_NUMPIECELIEN', '');
  end;

  if not (EnErreur) and not (Result) then
  begin
    Stg := 'Vous devez lier ce règlement à un autre règlement.';
    Result := FOJaiLeDroit(57, True, True, Stg);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/04/2003
Modifié le ... : 18/04/2003
Description .. : Recherche des règlements disponibles pour un article
Suite ........ : financier
Mots clefs ... :
*****************************************************************}

function FORechercheReglementDispo(TOBLigne, TOBPiece, TOBEches: TOB; TypeOpe: string; var Montant: double; var Trouve: boolean): boolean;
var
  Stg, CodeFiche: string;
  EnErreur: boolean;
begin
  Result := True;
  Trouve := False;
  Montant := 0;
  if (TOBPiece = nil) or (TOBLigne = nil) then Exit;
  if FOGetParamCaisse('GPK_LIAISONREG') <> 'X' then Exit;

  EnErreur := False;
  if TypeOpe = TYPEOPCENCCREDIT then
  begin
    // Encaissement de crédit => Reste dû
    Stg := 'MP_TYPEMODEPAIE=' + TYPEPAIERESTEDU + ';'
      + 'XX_WHERE=GOC_TOTALDISPOTTC>0;'
      + 'GOC_PREFIXE=GPE;'
      + 'GOC_NATUREPIECEG=' + TOBPiece.GetValue('GP_NATUREPIECEG') + ';'
      + 'GOC_TIERS=' + TOBPiece.GetValue('GP_TIERS') + ';';
    CodeFiche := 'MFOLIENECHE_MUL';
  end else
    if TypeOpe = TYPEOPCREMBAVOIR then
  begin
    // Remboursement d'avoir => Emission d'avoir
    Stg := 'MP_TYPEMODEPAIE=' + TYPEPAIEAVOIR + ';'
      + 'XX_WHERE=GOC_TOTALDISPOTTC<0;'
      + 'GOC_PREFIXE=GPE;'
      + 'GOC_NATUREPIECEG=' + TOBPiece.GetValue('GP_NATUREPIECEG') + ';'
      + 'GOC_TIERS=' + TOBPiece.GetValue('GP_TIERS') + ';';
    CodeFiche := 'MFOLIENECHE_MUL';
  end else
    if TypeOpe = TYPEOPCREMBBONACH then
  begin
    // Remboursement de bon d'achat => Acquisition d'un bon d'achat
    Stg := 'GA_TYPEARTFINAN=' + TYPEOPCBONACHAT + ';'
      + 'XX_WHERE=GOC_TOTALDISPOTTC>0;'
      + 'GOC_PREFIXE=GL;'
      + 'GOC_NATUREPIECEG=' + TOBPiece.GetValue('GP_NATUREPIECEG') + ';'
      + 'GOC_TIERS=' + TOBPiece.GetValue('GP_TIERS') + ';';
    CodeFiche := 'MFOLIENOPCAIS_MUL';
  end else
    if TypeOpe = TYPEOPCREMBARRHES then
  begin
    // Remboursement d'arrhes => Versement d'arrhes
    Stg := 'GA_TYPEARTFINAN=' + TYPEOPCARRHES + ';'
      + 'XX_WHERE=GOC_TOTALDISPOTTC>0;'
      + 'GOC_PREFIXE=GL;'
      + 'GOC_NATUREPIECEG=' + TOBPiece.GetValue('GP_NATUREPIECEG') + ';'
      + 'GOC_TIERS=' + TOBPiece.GetValue('GP_TIERS') + ';';
    CodeFiche := 'MFOLIENOPCAIS_MUL';
  end else
  begin
    // pas soumis à la liaison des règlements
    Result := True;
    if TOBLigne.FieldExists('GOC_NUMPIECELIEN') then
      TOBLigne.DelChampSup('GOC_NUMPIECELIEN', False);
    Exit;
  end;

  Result := LanceRechercheOpCaisseDispo(CodeFiche, Stg, 'GL', TypeOpe, TOBLigne, TOBPiece, TOBEches, Montant, EnErreur);
  if Result then
  begin
    Trouve := True;
    // pour rendre la ligne non modifiable dans la saisie de ticket
    if TOBLigne.FieldExists('MODIFPU') then
      TOBLigne.PutValue('MODIFPU', '-')
    else
      TOBLigne.AddChampSupValeur('MODIFPU', '-');
    TOBLigne.PutValue('GL_REMISABLELIGNE', '-');
    TOBLigne.PutValue('GL_REMISABLEPIED', '-');
  end else
  if not EnErreur then
  begin
    Stg := 'Vous devez lier cette opération de caisse à un règlement.';
    Result := FOJaiLeDroit(57, True, True, Stg);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/04/2003
Modifié le ... : 22/04/2003
Description .. : Vérification de la cohérence des règlements liés
Suite ........ :
Suite ........ : Valeur retournée :
Suite ........ :  0 = ok
Suite ........ :  1 = Le règlement lié est aussi utilisé pour une autre ligne
Suite ........ :  2 = Annulation impossible si un règlement est lié
Mots clefs ... :
*****************************************************************}

function FOVerifReglementLie(TOBPiece, TOBEches, TOBArticles: TOB; AnnulPiece: boolean; var NoLig: integer): integer;
var
  ii, jj: integer;
  RefPiece, RefArt, TypeOpe: string;
  Montant: double;
  TOBL, TOBA: TOB;
begin
  Result := 0;
  NoLig := 0;
  if (TOBPiece = nil) or (TOBEches = nil) then Exit;
  // on vérifie qu'une même opération de caisse n'a été utilisé qu'une seule fois dans le ticket
  for ii := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[ii];
    RefPiece := FOGetChampSupValeur(TOBL, 'GOC_NUMPIECELIEN', '');
    if RefPiece <> '' then
    begin
      if AnnulPiece then
      begin
        RefArt := TOBL.GetValue('GL_ARTICLE');
        TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefArt], False);
        TypeOpe := TOBA.GetValue('GA_TYPEARTFINAN');
        Montant := TOBL.GetValue('GL_TOTALTTC');
        if not FOIsOpCaisseLiee('GL', TypeOpe, Montant) then
        begin
          // annulation impossible si une opération de caisse est liée
          Result := 2;
          Exit;
        end;
      end;
      for jj := ii + 1 to TOBPiece.Detail.Count - 1 do
      begin
        if FOTesteChampSupValeur(TOBPiece.Detail[jj], 'GOC_NUMPIECELIEN', RefPiece) then
        begin
          NoLig := ii;
          Result := 1;
          Exit;
        end;
      end;
      if TOBEches.FindFirst(['GOC_NUMPIECELIEN'], [RefPiece], False) <> nil then
      begin
        NoLig := ii;
        Result := 1;
        Exit;
      end;
    end;
  end;
  // on vérifie qu'un même règlement n'a été utilisé qu'une seule fois dans le ticket
  for ii := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBL := TOBEches.Detail[ii];
    RefPiece := FOGetChampSupValeur(TOBL, 'GOC_NUMPIECELIEN', '');
    if RefPiece <> '' then
    begin
      if AnnulPiece then
      begin
        Montant := TOBL.GetValue('GPE_MONTANTECHE');
        TypeOpe := TOBL.GetValue('TYPEMODEPAIE');
        if not FOIsOpCaisseLiee('GPE', TypeOpe, Montant) then
        begin
          // Annulation impossible si un règlement est lié
          Result := 2;
          Exit;
        end;
      end;
      for jj := ii + 1 to TOBEches.Detail.Count - 1 do
      begin
        if RefPiece = FOGetChampSupValeur(TOBEches.Detail[jj], 'GOC_NUMPIECELIEN', '') then
        begin
          Result := 1;
          Exit;
        end;
      end;
    end;
  end;
  Result := 0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/04/2003
Modifié le ... : 14/04/2003
Description .. : Reprise de l'historique des opérations de caisse liées
Mots clefs ... :
*****************************************************************}

function FORepriseOpCaisseLies(CleDoc: R_CleDoc): boolean;
var
  TOBOpCais, TOBLig: TOB;
  RefArt, sSql, TypeOpe: string;
  Montant: double;
  QQ: TQuery;
  NumZCaisse : integer; // JTR
  Caisse : string; // JTR
begin
  Result := False;
  TypeOpe := '';
  // Lecture de la ligne
  TOBLig := TOB.Create('LIGNE', nil, -1);
  sSql := 'SELECT * FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, True);
  QQ := OpenSQL(sSql, True);
  TOBLig.SelectDB('', QQ);
  Ferme(QQ);
  RefArt := TOBLig.GetValue('GL_ARTICLE');
  Montant := TOBLig.GetValue('GL_TOTALTTC');
  // Recherche du type de l'article financier
  if (RefArt <> '') and (TOBLig.GetValue('GL_TYPEARTICLE') = 'FI') then
  begin
    sSql := 'SELECT GA_TYPEARTFINAN FROM ARTICLE WHERE GA_ARTICLE="' + RefArt + '"';
    QQ := OpenSQL(sSql, True);
    if not QQ.Eof then TypeOpe := QQ.FindField('GA_TYPEARTFINAN').AsString;
    Ferme(QQ);
  end;
  // Création de l'opération de caisse
  if FOIsOpCaisseCumul('GL', TypeOpe, Montant) then
  begin
    sSql := 'SELECT GP_NUMZCAISSE, GP_CAISSE FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, True);
    QQ := OpenSQL(sSql, True);
    if not QQ.EOF then
    begin
      Caisse := QQ.FindField('GP_CAISSE').AsString;
      NumZCaisse := QQ.FindField('GP_NUMZCAISSE').AsInteger;
    end else
    begin
      Caisse := '';
      NumZCaisse := 0;
    end;
    Ferme(QQ);
    TOBOpCais := TOB.Create('', nil, -1);
    TOBLigneVersOpCaisse(TOBLig, TOBOpCais, TypeOpe, NumZCaisse, Caisse);
    TOBOpCais.SetAllModifie(True);
    Result := TOBOpCais.InsertOrUpdateDB;
    TOBOpCais.Free;
  end;
  TOBLig.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/04/2003
Modifié le ... : 14/04/2003
Description .. : Reprise de l'historique des règlements liés
Mots clefs ... :
*****************************************************************}

function FORepriseReglementsLies(CleDoc: R_CleDoc): boolean;
var
  TOBOpCais, TOBEche: TOB;
  sSql, CodeMdR, TypeMdR, Etablis: string;
  Montant: double;
  QQ: TQuery;
  NumZCaisse : integer ; // JTR
  Caisse : string; // JTR
begin
  Result := False;
  TypeMdR := '';
  // Lecture de la ligne
  TOBEche := TOB.Create('PIEDECHE', nil, -1);
  sSql := 'SELECT PIEDECHE.*,MODEPAIE.MP_TYPEMODEPAIE,PIECE.GP_ETABLISSEMENT, PIECE.GP_NUMZCAISSE, PIECE.GP_CAISSE'
    + ' FROM PIEDECHE LEFT JOIN MODEPAIE ON MP_MODEPAIE=GPE_MODEPAIE'
    + ' LEFT JOIN PIECE ON GP_NATUREPIECEG=GPE_NATUREPIECEG'
    + ' AND GP_SOUCHE=GPE_SOUCHE AND GP_NUMERO=GPE_NUMERO AND GP_INDICEG=GPE_INDICEG'
    + ' WHERE ' + WherePiece(CleDoc, ttdEche, False)
    + ' AND GPE_NUMECHE=' + IntToStr(CleDoc.NumLigne);
  QQ := OpenSQL(sSql, True);
  TOBEche.SelectDB('', QQ);
  Ferme(QQ);
  CodeMdR := TOBEche.GetValue('GPE_MODEPAIE');
  Montant := TOBEche.GetValue('GPE_MONTANTECHE');
  TypeMdR := TOBEche.GetValue('MP_TYPEMODEPAIE');
  Etablis := TOBEche.GetValue('GP_ETABLISSEMENT');
  NumZCaisse := TOBEche.GetValue('GP_NUMZCAISSE');
  Caisse := TOBEche.GetValue('GP_CAISSE');
  // Mise à jour de l'échéance
  if FOIsOpCaisseCumul('GPE', TypeMdR, Montant) then
  begin
    TOBOpCais := TOB.Create('', nil, -1);
    TOBEcheanceVersOpCaisse(TOBEche, TOBOpCais, TypeMdR, Etablis, NumZCaisse, Caisse);
    TOBOpCais.SetAllModifie(True);
    Result := TOBOpCais.InsertOrUpdateDB;
    TOBOpCais.Free;
  end;
  TOBEche.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 25/04/2003
Modifié le ... : 25/04/2003
Description .. : Mise à jour des règlements liés en cas d'annulation de ticket
Mots clefs ... :
*****************************************************************}

procedure FOAnnulReglementLie(TOBPiece, TOBArticles, TOBEches, TOBPiece_O: TOB);
var
  TOBL, TOBA: TOB;
  Stg, TypeOpe: string;
  Ind: integer;
  DD: TDateTime;
  QQ: TQuery;
begin
  for Ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[Ind];
    if TOBL.GetValue('GL_TYPEARTICLE') = 'FI' then
    begin
      Stg := TOBL.GetValue('GL_ARTICLE');
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [Stg], False);
      TypeOpe := TOBA.GetValue('GA_TYPEARTFINAN');
      if FOIsOpCaisseCumul('GL', TypeOpe, TOBL.GetValue('GL_TOTALTTC')) then
      begin
        // Référence de la ligne du ticket courant
        DD := TOBPiece_O.GetValue('GP_DATEPIECE');
        Stg := 'GL;' + FormatDateTime('ddmmyyyy', DD) + ';'
          + TOBPiece_O.GetValue('GP_NATUREPIECEG') + ';'
          + TOBPiece_O.GetValue('GP_SOUCHE') + ';'
          + IntToStr(TOBPiece_O.GetValue('GP_NUMERO')) + ';'
          + IntToStr(TOBPiece_O.GetValue('GP_INDICEG')) + ';'
          + IntToStr(TOBL.GetValue('GL_NUMORDRE')) + ';';
        FOMajChampSupValeur(TOBL,'GOC_NUMPIECELIEN', Stg);
        FOMajChampSupValeur(TOBL,'GOC_TOTALDISPOTTC', 0);
      end;
    end;
  end;
  for Ind := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBL := TOBEches.Detail[Ind];
    if not TOBL.FieldExists('TYPEMODEPAIE') then
    begin
      TypeOpe := '';
      Stg := 'SELECT MP_TYPEMODEPAIE FROM MODEPAIE WHERE MP_MODEPAIE="'
        + TOBL.GetValue('GPE_MODEPAIE') + '"';
      QQ := OpenSQL(Stg, True);
      if not QQ.EOF then TypeOpe := QQ.FindField('MP_TYPEMODEPAIE').AsString;
      Ferme(QQ);
      TOBL.AddChampSupValeur('TYPEMODEPAIE', TypeOpe);
    end else
      TypeOpe := TOBL.GetValue('TYPEMODEPAIE');
    if FOIsOpCaisseCumul('GPE', TypeOpe, TOBL.GetValue('GPE_MONTANTECHE')) then
    begin
      // Référence de la ligne d'échéance du ticket courant
      DD := TOBPiece_O.GetValue('GP_DATEPIECE');
      Stg := 'GPE;' + FormatDateTime('ddmmyyyy', DD) + ';'
        + TOBPiece_O.GetValue('GP_NATUREPIECEG') + ';'
        + TOBPiece_O.GetValue('GP_SOUCHE') + ';'
        + IntToStr(TOBPiece_O.GetValue('GP_NUMERO')) + ';'
        + IntToStr(TOBPiece_O.GetValue('GP_INDICEG')) + ';'
        + IntToStr(TOBL.GetValue('GPE_NUMECHE')) + ';';
      FOMajChampSupValeur(TOBL,'GOC_NUMPIECELIEN', Stg);
      FOMajChampSupValeur(TOBL,'GOC_TOTALDISPOTTC', 0);
    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/04/2003
Modifié le ... : 22/04/2003
Description .. : Mise à jour du dernier numéro de bon utilisé
Mots clefs ... :
*****************************************************************}

procedure TFOMajNoBon.MAJNumeroBon;
var
  sSql: string;
begin
  if V_PGI.IoError <> oeOk then Exit;
  if NbBon <= 0 then Exit;
  if not GetParamSoc('SO_GCFONOBONAUTO') then Exit;
  // mise à jour du dernier no attribué
  Num := GetParamSoc('SO_GCFOCHRONOBON', True);
  sSql := 'UPDATE PARAMSOC SET SOC_DATA="' + IntToStr(Num + NbBon) + '"'
    + ' WHERE SOC_NOM="SO_GCFOCHRONOBON"';
  if Num = 0 then
    sSql := sSql + ' AND (SOC_DATA="' + IntToStr(Num) + '"' + ' OR SOC_DATA="")'
  else
    sSql := sSql + ' AND SOC_DATA="' + IntToStr(Num) + '"';
  if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/04/2003
Modifié le ... : 22/04/2003
Description .. : Mise en forme du numéro de bon
Mots clefs ... :
*****************************************************************}

function AttributionNoBon(var Numero: integer): string;
var
  Prefixe, TypeCAB, DernierChrono: string;
  Taille, LongMax: integer;
begin
  Result := '';
  if Numero < 0 then Exit;
  LongMax := ChampToLength('GOC_NUMBON');
  Taille := GetParamSoc('SO_GCFOLGNOBON');
  if (Taille < 1) or (Taille > LongMax) then Taille := LongMax;
  Prefixe := Trim(GetParamSoc('SO_GCFOPREFIXENOBON'));
  TypeCAB := Trim(GetParamSoc('SO_GCFOCABNOBON'));
  DernierChrono := IntToStr(Numero);
  Result := AttribNewCode('OPERCAISSE', 'GOC_NUMBON', Taille, Prefixe, DernierChrono, TypeCAB, False);
  Inc(Numero);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/04/2003
Modifié le ... : 22/04/2003
Description .. : Mise à jour des numéros des bons
Mots clefs ... :
*****************************************************************}

function FOAffecteNoBonReglementLie(TOBPiece, TOBArticles, TOBEches: TOB): boolean;
var
  TOBL, TOBA: TOB;
  RefArt, TypeOpe: string;
  Ind, NbBon, Num: integer;
  io: TIOErr;
  XX: TFOMajNoBon;
begin
  Result := True;
  if not GetParamSoc('SO_GCFONOBONAUTO') then Exit;

  // calcul du nombre de bons nécessaires
  NbBon := 0;
  for Ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[Ind];
    if TOBL.GetValue('GL_TYPEARTICLE') = 'FI' then
    begin
      RefArt := TOBL.GetValue('GL_ARTICLE');
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefArt], False);
      TypeOpe := TOBA.GetValue('GA_TYPEARTFINAN');
      if FOIsOpCaisseNoBon('GL', TypeOpe, TOBL.GetValue('GL_TOTALTTC')) then
        Inc(NbBon);
    end;
  end;
  for Ind := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBL := TOBEches.Detail[Ind];
    TypeOpe := TOBL.GetValue('TYPEMODEPAIE');
    if FOIsOpCaisseNoBon('GPE', TypeOpe, TOBL.GetValue('GPE_MONTANTECHE')) then
      Inc(NbBon);
  end;
  if NbBon > 0 then
  begin
    // mise à jour du dernier no attribué
    XX := TFOMajNoBon.Create;
    try
      XX.NbBon := NbBon;
      io := Transactions(XX.MAJNumeroBon, 5);
      Num := XX.Num;
    finally
      XX.Free;
    end;
    // mise à jour du numéro de bon sur les lignes et les échéances
    if io = oeOk then
    begin
      for Ind := 0 to TOBPiece.Detail.Count - 1 do
      begin
        TOBL := TOBPiece.Detail[Ind];
        if TOBL.GetValue('GL_TYPEARTICLE') = 'FI' then
        begin
          RefArt := TOBL.GetValue('GL_ARTICLE');
          TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefArt], False);
          TypeOpe := TOBA.GetValue('GA_TYPEARTFINAN');
          if FOIsOpCaisseNoBon('GL', TypeOpe, TOBL.GetValue('GL_TOTALTTC')) then
            FOMajChampSupValeur(TOBL,'GOC_NUMBON', AttributionNoBon(Num));
        end;
      end;
      for Ind := 0 to TOBEches.Detail.Count - 1 do
      begin
        TOBL := TOBEches.Detail[Ind];
        TypeOpe := TOBL.GetValue('TYPEMODEPAIE');
        if FOIsOpCaisseNoBon('GPE', TypeOpe, TOBL.GetValue('GPE_MONTANTECHE')) then
          FOMajChampSupValeur(TOBL,'GOC_NUMBON', AttributionNoBon(Num));
      end;
    end else
      Result := False;
  end;
end;

end.
