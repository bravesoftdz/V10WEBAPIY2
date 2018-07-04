unit ToxMessage;

interface

uses {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hctrls, SysUtils, uTob ;

/////////////////////////////////////////////////////////
// Liste des fonctions permettant d'afficher un message
////////////////////////////////////////////////////////
function MessageTOBAcomptes          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBAdresses          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBArrondi           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBArticle           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBaffaire           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBArticleCompl      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBChancell          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBChoixcod          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBChoixExt          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBClavierEcran      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBCommercial        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBCompteurEtab      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBContact           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBCtrlCaisMt        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBDepots            (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBDevise            (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBCommentaire       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBDimension         (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBDimasque          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBDispo             (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBEtabliss          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBExercice          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBFideliteEnt       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBFideliteLig       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBFonction          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBJoursCaisse       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBJoursEtab         (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBJoursEtabEvt      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBLiensOLE          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBLigne             (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBLigneLot          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBLigneNomen        (TPST : TOB; MessErreur : integer) : string ;
{$IFDEF BTP}
function MessageTOBLigneOuvrages     (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBMilliemes         (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBMemorisation      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParDoc           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPieceRG           (TPST : TOB; MessErreur : integer) : string ;
{$ENDIF}
function MessageTOBListeinvent       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBListeinvlig       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBMea               (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBModeles           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBModeData          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBModeRegl          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBModePaie          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBNaturePrest       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBNomenEnt          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBNomenLig          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBOperCaisse        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParCaisse         (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParamSoc          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParamTaxe         (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParFidelite       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParRegleFid       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParSeuilFid       (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBParPiecBil        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPays              (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPiedBase          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPiedEche          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPiedPort          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPiece             (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPort              (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPieceAdresse      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBPprofilArt      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBRegion            (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBSociete           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBStoxQuerys        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTarif             (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTarifMode         (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTarifPer          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTarifTypMode      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOByTarifs           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOByTarifsFourchette (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTiers             (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTiersCompl        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTradDico          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTradTablette      (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTraduc            (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTxCptTva          (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTypeMasque        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBTypeRemise        (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBUserGrp           (TPST : TOB; MessErreur : integer) : string ;
function MessageTOBUtilisat          (TPST : TOB; MessErreur : integer) : string ;
function ToxEditeMessage             (NomTable : string; ToxInfo : TOB; MessErreur : integer) : string;

//////////////////////////////////////////////////
// Codes Erreurs
//////////////////////////////////////////////////
const

      EnregNotTrt      = -100 ;     // Enregistrement non traité
      IntegOK          = 0  ;     	// Enregistrement OK
      TableNonReconnu  = 1  ;       // Table non reconnue et non gérée par la TOX
      ErrNaturePiece   = 3  ;				// Nature de pièce inconnue
      ErrFamilleNiv1   = 4  ;				// Famille Niveau 1 inexistante
      ErrFamilleNiv2   = 5  ;				// Famille Niveau 2 inexistante
      ErrFamilleNiv3   = 6  ;				// Famille Niveau 3 inexistante
      ErrFamilleTaxe1  = 7  ;				// Famille Taxe 1 inexistante
      ErrFamilleTaxe2  = 8  ;				// Famille Taxe 2 inexistante
      ErrCollection    = 9  ;				// Collection inexistante
      ErrTarifArticle  = 10 ;				// Tarif Article inexistant
      ErrStatutArticle = 11 ;       // Statut Article incorrect
      ErrTypeArticle   = 12 ;       // Type Article incorrect
      ErrDimMasque     = 13 ;       // Masque de saisie inexistante
      ErrGrilleDim1    = 14 ;				// Grille sur dimension 1 inexistante
      ErrGrilleDim2    = 15 ;				// Grille sur dimension 2 inexistante
      ErrGrilleDim3    = 16 ;				// Grille sur dimension 3 inexistante
      ErrGrilleDim4    = 17 ;				// Grille sur dimension 4 inexistante
      ErrGrilleDim5    = 18 ;				// Grille sur dimension 5 inexistante
      ErrStatArt1      = 19 ;				// Valeur table libre1 inexistante
      ErrStatArt2      = 20 ;				// Valeur table libre2 inexistante
      ErrStatArt3      = 21 ;				// Valeur table libre3 inexistante
      ErrStatArt4      = 22 ;				// Valeur table libre4 inexistante
      ErrStatArt5      = 23 ;				// Valeur table libre5 inexistante
      ErrStatArt6      = 24 ;				// Valeur table libre6 inexistante
      ErrStatArt7      = 25 ;				// Valeur table libre7 inexistante
      ErrStatArt8      = 26 ;				// Valeur table libre8 inexistante
      ErrStatArt9      = 27 ;				// Valeur table libre9 inexistante
      ErrStatArt10     = 28 ;				// Valeur table libre10 inexistante
      ErrDevise        = 29 ;				// Devise inexistante
      ErrTypeCommercial  = 30 ;				// Type de commercial incorrect
      ErrEtablissement   = 31 ;				// Etablissement inexistant
      ErrZoneCommercial  = 32 ;				// Zone commerciale inexistante
      ErrTypeCommission  = 33 ;				// Type de commissionnement incorrect
      ErrCivilite        = 35 ;				// Civilité inexistante
      ErrParente         = 36 ;				// Lien de parenté inexistant
      ErrSexe            = 37 ;				// Sexe incorrect !!!!!!
      ErrPays            = 38 ;				// Pays inexistant
      ErrArticle         = 39 ;       // Article inconnu
      ErrDepot           = 40 ;       // Dépôt
      ErrLangue          = 41 ;       // Langue inexistante
      ErrModePaie        = 42 ;       // Mode de paiement   inexistant
      ErrModePaie1       = 43 ;       // Mode de paiement 1 inexistant
      ErrModePaie2       = 44 ;       // Mode de paiement 2 inexistant
      ErrModePaie3       = 45 ;       // Mode de paiement 3 inexistant
      ErrModePaie4       = 46 ;       // Mode de paiement 4 inexistant
      ErrModePaie5       = 47 ;       // Mode de paiement 5 inexistant
      ErrModePaie6       = 48 ;       // Mode de paiement 6 inexistant
      ErrModePaie7       = 49 ;       // Mode de paiement 7 inexistant
      ErrModePaie8       = 50 ;       // Mode de paiement 8 inexistant
      ErrModePaie9       = 51 ;       // Mode de paiement 9 inexistant
      ErrModePaie10      = 52 ;       // Mode de paiement 10 inexistant
      ErrModePaie11      = 53 ;       // Mode de paiement 11 inexistant
      ErrModePaie12      = 54 ;       // Mode de paiement 12 inexistant
      ErrClient          = 55 ;       // Client inexistant
      ErrFournisseur     = 56 ;       // Fournisseur inexistant
      ErrSecteurAct      = 57 ;       // Secteur d'activité inexistant
      ErrZoneCom         = 58 ;       // Zone Commerciale inexistante
      ErrTarifTiers      = 59 ;       // Tarif Tiers inexistant
      ErrModeRegle       = 60 ;       // Mode de règlement inexistant
      ErrRegimeTVA       = 61 ;       // Régime de TVA inexistant
      ErrNationalite     = 62 ;       // Nationalité inexistante
      ErrFamComptable    = 63 ;       // Famille comptable inexistante
      ErrTiers           = 64 ;       // Tiers inexistant
      ErrTiersLivre      = 65 ;       // Tiers Livré inexistant
      ErrTiersFacture    = 66 ;       // Tiers Facturé inexistant
      ErrTiersPayeur     = 67 ;       // Tiers Payeur inexistant
      ErrRepresentant    = 68 ;       // Représentant inexistant
      ErrDepotDest       = 69 ;       // Dépot destinataire inexistant
      ErrArticleGen      = 70 ;       // Article générique inexistant
      ErrDeviseEsp       = 71 ;       // Devise montant encaissé inexsistante
      ErrCategorieTaxe   = 72 ;       // Catégorie de taxe interdite
      ErrFamilleTaxe     = 73 ;       // Famille de taxe
      ErrAnnulDoc        = 74 ;       // Erreur lors l'annulation du document
      ErrArticleRat      = 75 ;       // Article de référence inconnu
      ErrArticleRef      = 76 ;       // Article de rattachement inconnu
      ErrMeteo           = 77 ;       // Code météo inconnu
      ErrTypeArrondi     = 78 ;       // Type d'arrondi inexistant
      ErrTypeDemarque    = 79 ;       // Démarque inexistante
      ErrTypePort        = 80 ;       // Type de port inexistant
      ErrDico            = 81 ;       // Dictionnaire inexistant
      ErrCodeTablette    = 82 ;       // Code de tablette inexistant
      ErrForme           = 83 ;       // Forme inexistante
      ErrFamilleNiv4     = 84 ;		// Famille Niveau 4 inexistante
      ErrFamilleNiv5     = 85 ;		// Famille Niveau 5 inexistante
      ErrFamilleNiv6     = 86 ;		// Famille Niveau 6 inexistante
      ErrFamilleNiv7     = 87 ;		// Famille Niveau 7 inexistante
      ErrFamilleNiv8     = 88 ;		// Famille Niveau 8 inexistante
      ErrStat9Art1       = 89 ;		// Statistique article 1
      ErrStat9Art2       = 90 ;		// Statistique article 2
      ErrRegime          = 91 ;		// Régime fiscal
      ErrTiersVide       = 92 ;   // Tiers inexistant
      ErrTarif           = 93 ;   // Tarif inexistant
      ErrDate            = 94 ;   // Pb de date (fourchette)
      ErrAffaire         = 95 ;   // Affaire inexistante
{$IFDEF BTP}
			ErrTypeRg          = 96 ;   // type de retenue de garantie inexistante
{$ENDIF}

      ErrAnnulePiece         = 498 ;      // Erreur lors la mise en jour DB (UpdateDB)
      ErrInsertDB            = 499 ;      // Erreur lors la mise en jour DB (InsertDB)
      ErrUpdateDB            = 500 ;      // Erreur lors la mise en jour DB (UpdateDB)
      ErrLoadLesNomens       = 501 ;      // Erreur au chargement des Nomenclatures
      ErrLoadLesLots         = 502 ;      // Erreur au chargement des Nomenclatures
      ErrCalculDispo         = 503 ;      // Erreur lors du calcul de la MAJ du stock
      ErrMAJDispo            = 504 ;      // Erreur lors de la MAJ des stocks
      ErrMAJTiersPiece       = 505 ;      // Erreur lors de la MAJ du Tiers de la pièce
      ErrMAJAcomptes         = 506 ;      // Erreur lors de la MAJ des acomptes
      ErrMAJPorts            = 507 ;      // Erreur lors de la MAJ des lignes de ports
      ErrMAJNomencl          = 508 ;      // Erreur lors de la MAJ des nomenclatures
      ErrCalculInverseDispo  = 509 ;      // Erreur lors du calcul de MAJ du stock pour suppression de la pièce
      ErrMAJInverseDispo     = 510 ;      // Erreur lors de la MAJ du stock pour suppression de la pièce
{$IFDEF BTP}
      ErrLoadLesOuvrages     = 511 ;      // Erreur au chargement des ouvrages
      ErrMAJOuvrage          = 512 ;      // Erreur lors de la MAJ des ouvrages
      ErrMillieme            = 513 ;      // Erreur lors de la MAJ des milliemes de tva
      ErrMAJPieceRG          = 514 ;      // Erreur lors de la MAJ des Retenues de garantie sur piece
      ErrMAJPiedPieceRG      = 515 ;      // Erreur lors de la MAJ des base de Tva / retenue de garantie
{$ENDIF}
    	EnregNotTrtErr         = 1000 ;     // Enregistrement non traité car erreur dans traitement TOB tarnsactionnelle



implementation

////////////////////////////////////////////////////////
//  Message TOB Acomptes
////////////////////////////////////////////////////////
function MessageTOBAcomptes (TPST : TOB; MessErreur : integer) : string ;
var NumEcr       : integer  ;
    Probleme     : string   ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrModePaie then
    begin
      Probleme := TPST.GetValue ('GAC_MODEPAIE');
      result   := 'Mode de paiement inexistant : "' + Probleme + '"';
      exit ;
    end;
    result   := 'Erreur ACOMPTES n°' + IntToStr (MessErreur);
  end else
  begin
    NumEcr  := TPST.GetValue ('GAC_NUMECR');
    result       := 'Ligne accompte n°' + IntToStr(NumEcr);
  end;
end;


////////////////////////////////////////////////////////
// Message TOB ADRESSES
////////////////////////////////////////////////////////
function MessageTOBAdresses (TPST : TOB; MessErreur : integer) : string ;
var NumAdresse    : integer ;
    OrigineOK     : boolean ;
    RefAdresse    : string  ;
    NaturePiece   : string  ;
    LibellePiece  : string  ;
    RefTiers      : string  ;
    TypeTiers     : string  ;
    SQL           : string  ;
    Q             : TQUERY  ;
begin

  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur Adresses n°' + IntToStr (MessErreur);
  end else
  begin
   OrigineOK  := False ;
   NumAdresse := TPST.GetValue ('ADR_NUMEROADRESSE');

   ////////////////////////////////////////////////////////
   // Adresse d'une pièce
   ////////////////////////////////////////////////////////
   if (TPST.GetValue ('ADR_TYPEADRESSE') = 'PIE') then
   begin
                                      RefAdresse  := TPST.GetValue ('ADR_REFCODE');
   NaturePiece := ReadTokenSt(RefAdresse) ;

   SQL:='Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="'+NaturePiece+'"';
   Q:=OpenSQL(SQL,True) ;
   if Not Q.EOF then
   begin
   LibellePiece := Q.FindField('GPP_LIBELLE').AsString ;
   result       := 'L''adresse de la pièce ' +  '"' + LibellePiece+ '"';
   end else
   begin
   result := 'L''adresse de pièce (n° ' + IntToStr(NumAdresse) + ')';
   end;
   Ferme(Q);
   OrigineOK:=True ;
   end
   ////////////////////////////////////////////////////////
   // Adresse d'un tiers
   ////////////////////////////////////////////////////////
   else if (TPST.GetValue ('ADR_TYPEADRESSE') = 'TIE') then
   begin
                                      RefTiers := TPST.GetValue ('ADR_REFCODE');

   SQL := 'Select T_NATUREAUXI From TIERS WHERE T_AUXILIAIRE="'+RefTiers+'"';
   Q   := OpenSQL(SQL,True) ;
   if Not Q.EOF then
   begin
   TypeTiers := Q.FindField('T_NATUREAUXI').AsString ;
   if      (TypeTiers = 'CLI') then result := 'L''adresse du client ' + RefTiers
   else if (TypeTiers = 'FOU') then result := 'L''adresse du fournisseur ' + RefTiers
   else result := 'L''adresse du tiers ' + RefTiers;
   end
   else result := 'L''adresse du tiers ' + RefTiers;
   Ferme (Q);
   OrigineOK:=True ;
   end;

   if (OrigineOK=False) then result := 'L''adresse n°' + IntToStr(NumAdresse);
  end;
end;

////////////////////////////////////////////////////////
// Message TOB ARRONDI
////////////////////////////////////////////////////////
function MessageTOBArrondi (TPST : TOB; MessErreur : integer) : string ;
var CodeArrondi : string ;
		Libelle			: string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur Arrondi n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeArrondi := TPST.GetValue ('GAR_CODEARRONDI');
  	Libelle     := TPST.GetValue ('GAR_LIBELLE');
  	result      := 'Le Code arrondi "' + CodeArrondi + '" ' +  Libelle  ;
  end;
end;

////////////////////////////////////////////////////////
// Message TOB ARTICLE
////////////////////////////////////////////////////////
function  MessageTOBArticle (TPST : TOB; MessErreur : integer) : string ;
var CodeArticle : string ;
    Probleme    : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv1 then
    begin
	    Probleme := TPST.GetValue ('GA_FAMILLENIV1');
      result   := 'La famille article de niveau 1 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv2 then
    begin
	    Probleme := TPST.GetValue ('GA_FAMILLENIV2');
      result   := 'La famille article de niveau 2 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv3 then
    begin
	    Probleme := TPST.GetValue ('GA_FAMILLENIV3');
      result   := 'La famille article de niveau 3 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleTaxe1 then
    begin
	    Probleme := TPST.GetValue ('GA_FAMILLETAXE1');
      result   := 'La famille de taxe 1 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleTaxe2 then
    begin
	    Probleme := TPST.GetValue ('GA_FAMILLETAXE2');
      result   := 'La famille de taxe 2 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrCollection then
    begin
	    Probleme := TPST.GetValue ('GA_COLLECTION');
      result   := 'La collection est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrTarifArticle then
    begin
	    Probleme := TPST.GetValue ('GA_TARIFARTICLE');
      result   := 'Le code tarif est inexistant : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatutArticle then
    begin
	    Probleme := TPST.GetValue ('GA_STATUTART');
      result   := 'Le statut de l''article est incorrect : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrTypeArticle then
    begin
	    Probleme := TPST.GetValue ('GA_TYPEARTICLE');
      result   := 'Le type de l''article est incorrect : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrDimMasque then
    begin
	    Probleme := TPST.GetValue ('GA_DIMMASQUE');
      result   := 'Le masque de dimensions est inexistant : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrGrilleDim1 then
    begin
	    Probleme := TPST.GetValue ('GA_GRILLEDIM1');
      result   := 'La grille de dimension 1 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrGrilleDim2 then
    begin
	    Probleme := TPST.GetValue ('GA_GRILLEDIM2');
      result   := 'La grille de dimension 2 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrGrilleDim3 then
    begin
	    Probleme := TPST.GetValue ('GA_GRILLEDIM3');
      result   := 'La grille de dimension 3 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrGrilleDim4 then
    begin
	    Probleme := TPST.GetValue ('GA_GRILLEDIM4');
      result   := 'La grille de dimension 4 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrGrilleDim5 then
    begin
	    Probleme := TPST.GetValue ('GA_GRILLEDIM5');
      result   := 'La grille de dimension 5 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt1 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART1');
      result   := 'La statistique libre 1 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt2 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART2');
      result   := 'La statistique libre 2 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt3 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART3');
      result   := 'La statistique libre 3 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt4 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART4');
      result   := 'La statistique libre 4 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt5 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART5');
      result   := 'La statistique libre 5 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt6 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART6');
      result   := 'La statistique libre 6 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt7 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART7');
      result   := 'La statistique libre 7 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt8 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART8');
      result   := 'La statistique libre 8 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt9 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREART9');
      result   := 'La statistique libre 9 n''existe pas en table : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStatArt10 then
    begin
	    Probleme := TPST.GetValue ('GA_LIBREARTA');
      result   := 'La statistique libre 10 n''existe pas en table : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur article n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeArticle := TPST.GetValue ('GA_ARTICLE');
    result      := 'La fiche article référence ' + CodeArticle ;
  end;
end;

function MessageTOBaffaire           (TPST : TOB; MessErreur : integer) : string ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end else result := 'Erreur affaire n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'L''affaire ' + TPST.GetValue ('AFF_LIBELLE');
  end;
end;

////////////////////////////////////////////////////////
// Message TOB ARTICLECOMPL
////////////////////////////////////////////////////////
function  MessageTOBArticleCompl (TPST : TOB; MessErreur : integer) : string ;
var CodeArticle : string ;
    Probleme    : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv4 then
    begin
	    Probleme := TPST.GetValue ('GA2_FAMILLENIV4');
      result   := 'La famille article de niveau 4 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv5 then
    begin
	    Probleme := TPST.GetValue ('GA2_FAMILLENIV5');
      result   := 'La famille article de niveau 5 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv6 then
    begin
	    Probleme := TPST.GetValue ('GA2_FAMILLENIV6');
      result   := 'La famille article de niveau 6 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv7 then
    begin
	    Probleme := TPST.GetValue ('GA2_FAMILLENIV7');
      result   := 'La famille article de niveau 7 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleNiv8 then
    begin
	    Probleme := TPST.GetValue ('GA2_FAMILLENIV8');
      result   := 'La famille article de niveau 8 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStat9Art1 then
    begin
	    Probleme := TPST.GetValue ('GA2_STATART1');
      result   := 'La statistique 1 est inexistante : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrStat9Art2 then
    begin
	    Probleme := TPST.GetValue ('GA2_STATART2');
      result   := 'La statistique 2 est inexistante : "' + Probleme + '"';
      exit ;
    end;

    result   := 'Erreur articlecompl n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeArticle := TPST.GetValue ('GA2_ARTICLE');
    result      := 'La statistique de l''article référence ' + CodeArticle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB ARTICLELIE
////////////////////////////////////////////////////////
function  MessageTOBArticleLie (TPST : TOB; MessErreur : integer) : string ;
var ArticleRef : string ;
    ArticleRat : string ;
    Probleme   : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticleRef then
    begin
	    Probleme := TPST.GetValue ('GAL_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrArticleRat then
    begin
	    Probleme := TPST.GetValue ('GAL_ARTICLELIE');
      result   := 'L''article de rattachement n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur ARTICLELIE n°' + IntToStr (MessErreur);
  end else
  begin
  	ArticleRef  := TPST.GetValue ('GAL_ARTICLE');
    ArticleRat  := TPST.GetValue ('GAL_ARTICLELIE');

  	result := 'L''article ' + ArticleRat + ' lié à ' + ArticleRef;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB ARTICLEPIECE
////////////////////////////////////////////////////////
function  MessageTOBArticlePiece (TPST : TOB; MessErreur : integer) : string ;
var Article  : string ;
    Nature   : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GAP_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrNaturePiece then
    begin
	    Probleme := TPST.GetValue ('GAP_NATUREPIECEG');
      result   := 'La nature de document n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur ARTICLEPIECE n°' + IntToStr (MessErreur);
  end else
  begin
  	Article  := TPST.GetValue ('GAP_ARTICLE');
    Nature   := TPST.GetValue ('GAP_NATUREPIECEG');

  	result := 'L''exception pour l''article "' + Article + '" et la nature de document ' + Nature ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB ARTICLETIERS
////////////////////////////////////////////////////////
function  MessageTOBArticleTiers (TPST : TOB; MessErreur : integer) : string ;
var Article  : string ;
    Tiers    : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GAT_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTiers then
    begin
	    Probleme := TPST.GetValue ('GAT_REFTIERS');
      result   := 'Le tiers n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur ARTICLETIERS n°' + IntToStr (MessErreur);
  end else
  begin
  	Article := TPST.GetValue ('GAT_ARTICLE');
    Tiers   := TPST.GetValue ('GAT_REFTIERS');

  	result := 'La référence de l''article "' + Article + '" pour le tiers ' + Tiers ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB CATALOGU
////////////////////////////////////////////////////////
function  MessageTOBCatalogu (TPST : TOB; MessErreur : integer) : string ;
var Article  : string ;
    Tiers    : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GCA_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTiers then
    begin
	    Probleme := TPST.GetValue ('GCA_TIERS');
      result   := 'Le fournisseur n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur CATALOGU n°' + IntToStr (MessErreur);
  end else
  begin
  	Article := TPST.GetValue ('GCA_ARTICLE');
    Tiers   := TPST.GetValue ('GCA_TIERS');

  	result := 'La référence de l''article "' + Article + '" pour le fournisseur ' + Tiers;
  end;
end;


////////////////////////////////////////////////////////
//  Message TOB CHANCELL
////////////////////////////////////////////////////////
function  MessageTOBChancell (TPST : TOB; MessErreur : integer) : string ;
var CodeDevise   : string ;
		DateCours	   : string ;
    Probleme     : string ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDevise then
    begin
      Probleme := TPST.GetValue ('H_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit ;
    end
    else result := 'Erreur Chancellerie n°' + IntToStr (MessErreur);
  end else
  begin
    CodeDevise := TPST.GetValue ('H_DEVISE');
    DateCours  := TPST.GetValue ('H_DATECOURS');
    result     := 'Le taux de la devise ' +  CodeDevise + ' au ' + DateCours;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB CHOIXCOD
////////////////////////////////////////////////////////
function  MessageTOBChoixcod (TPST : TOB; MessErreur : integer) : string ;
var CCType       : string ;
    CCCode       : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur ChoixCod n°' + IntToStr (MessErreur);
  end else
  begin
  	CCType := TPST.GetValue ('CC_TYPE');
  	CCCode := TPST.GetValue ('CC_CODE');
		result := 'L''enregistrement de la table CHOIXCOD de type ' + CCType + ' et de code '+ CCCode;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB CHOIXEXT
////////////////////////////////////////////////////////
function  MessageTOBChoixExt (TPST : TOB; MessErreur : integer) : string ;
var YXType       : string ;
    YXCode       : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur ChoixExt n°' + IntToStr (MessErreur);
  end else
  begin
  	YXType := TPST.GetValue ('YX_TYPE');
  	YXCode := TPST.GetValue ('YX_CODE');
		result := 'L''enregistrement de la table CHOIXEXT de type ' + YXType + ' et de code '+ YXCode;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB CLAVIERECRAN
////////////////////////////////////////////////////////
function  MessageTOBClavierEcran (TPST : TOB; MessErreur : integer) : string ;
var Caisse : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur ClavierEcran n°' + IntToStr (MessErreur);
  end else
  begin
  	Caisse := TPST.GetValue ('CE_CAISSE');
	result := 'Le paramétrage du clavier de la caisse "' + Caisse + '"';
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB COMMERCIAL
////////////////////////////////////////////////////////
function  MessageTOBCommercial (TPST : TOB; MessErreur : integer) : string ;
var CodeCommercial : string ;
    Libelle        : string ;
    Probleme       : string ;
begin
  if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTypeCommercial then
    begin
	    Probleme := TPST.GetValue ('GCL_COMMERCIAL');
      result   := 'Le type de commercial est incorrect : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrEtablissement then
    begin
	    Probleme := TPST.GetValue ('GCL_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrZoneCommercial then
    begin
	    Probleme := TPST.GetValue ('GCL_ZONECOM');
      result   := 'La zone commerciale n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTypeCommission then
    begin
	  Probleme := TPST.GetValue ('GCL_TYPECOM');
      result   := 'Le type de commissionnement est incorrect : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrNaturePiece then
    begin
	    Probleme := TPST.GetValue ('GCL_NATUREPIECEG');
      if Probleme = 'ZZ1' then // JTR - eQualité 12036
        result   := 'Une des natures de pièce n''existe pas : "FAC", "AVS", "AVC", "FFO"'
        else
        result   := 'Le nature de pièce n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Commercial n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeCommercial := TPST.GetValue ('GCL_COMMERCIAL');
  	Libelle        := TPST.GetValue ('GCL_LIBELLE');
  	result         := 'Le commercial "' + CodeCommercial + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB COMMERCIAL
////////////////////////////////////////////////////////
function  MessageTOBCompteurEtab (TPST : TOB; MessErreur : integer) : string ;
var Etab     : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrEtablissement then
    begin
      Probleme := TPST.GetValue ('GCL_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur COMPTEURETAB n°' + IntToStr (MessErreur);
  end else
  begin
  	Etab   := TPST.GetValue ('GCE_ETABLISSEMENT');
  	result         := 'Le compteur d''entrée de l''établissement "' + Etab + '" ';
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB CONTACT
////////////////////////////////////////////////////////
function  MessageTOBContact (TPST : TOB; MessErreur : integer) : string ;
var Libelle  : string  ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrCivilite then
    begin
	    Probleme := TPST.GetValue ('C_CIVILITE');
      result   := 'La civilité n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrParente then
    begin
	    Probleme := TPST.GetValue ('C_LIPARENT');
      result   := 'Le lien de parenté n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrSexe then
    begin
	    Probleme := TPST.GetValue ('C_SEXE');
      result   := 'Le Sexe n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Contact n°' + IntToStr (MessErreur);
  end else
  begin
  	Libelle := TPST.GetValue ('C_NOM');
  	result  := 'Le contact ' + Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB  CTRLCAISMT
////////////////////////////////////////////////////////
function  MessageTOBCtrlCaisMt (TPST : TOB; MessErreur : integer) : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur CtrlCaisMt n°' + IntToStr (MessErreur);
  end else
  begin
  	result  := 'Le contrôle de caisse' ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB DEPOTS
////////////////////////////////////////////////////////
function  MessageTOBDepots (TPST : TOB; MessErreur : integer) : string ;
var CodeDepot : string ;
		Libelle   : string ;
    Probleme  : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
  	if MessErreur = ErrPays then
    begin
	    Probleme := TPST.GetValue ('GDE_PAYS');
      result   := 'Le pays n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Dépôt n°' + IntToStr (MessErreur);
  end else
  begin
    CodeDepot := TPST.GetValue ('GDE_DEPOT');
    Libelle   := TPST.GetValue ('GDE_LIBELLE');
    result    := 'Le dépôt "' + CodeDepot + '" ' + Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB DEVISE
////////////////////////////////////////////////////////
function  MessageTOBDevise (TPST : TOB; MessErreur : integer) : string ;
var CodeDevise : string ;
    Libelle    : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur Devise n°' + IntToStr (MessErreur);
  end else
  begin
    CodeDevise := TPST.GetValue ('D_DEVISE');
    Libelle    := TPST.GetValue ('D_LIBELLE');
    result     := 'La devise "' + CodeDevise + '" ' +  Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB COMMENTAIRE
////////////////////////////////////////////////////////
function  MessageTOBCommentaire (TPST : TOB; MessErreur : integer) : string ;
var CodeComm : string ;
    Libelle    : string ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur Commentaire n°' + IntToStr (MessErreur);
  end else
  begin
    CodeComm   := TPST.GetValue ('GCT_CODE');
    Libelle    := TPST.GetValue ('GCT_LIBELLE');
    result     := 'Le commentaire "' + CodeComm + '" ' +  Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB DIMENSION
////////////////////////////////////////////////////////
function MessageTOBDimension (TPST : TOB; MessErreur : integer) : string ;
var TypeDim		    : string ;
    GrilleDim     : string ;
    Probleme      : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
  	if MessErreur = ErrGrilleDim1 then
    begin
	    Probleme := TPST.GetValue ('GDI_GRILLEDIM');
      result   := 'La grille de tailles de la première dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim2 then
    begin
	    Probleme := TPST.GetValue ('GDI_GRILLEDIM');
      result   := 'La grille de tailles de la deuxième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim3 then
    begin
	    Probleme := TPST.GetValue ('GDI_GRILLEDIM');
      result   := 'La grille de tailles de la troisième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim4 then
    begin
	    Probleme := TPST.GetValue ('GDI_GRILLEDIM');
      result   := 'La grille de tailles de la quatrième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim5 then
    begin
	    Probleme := TPST.GetValue ('GDI_GRILLEDIM');
      result   := 'La grille de tailles  de la cinquième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Libellé Grille n°' + IntToStr (MessErreur);
  end else
  begin
  	TypeDim   := TPST.GetValue ('GDI_TYPEDIM');
  	GrilleDim := TPST.GetValue ('GDI_GRILLEDIM');
		result    := 'Le libellé de la grille ' + GrilleDim ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB DIMASQUE
////////////////////////////////////////////////////////
function  MessageTOBDimasque (TPST : TOB; MessErreur : integer) : string ;
var	CodeMasque : string ;
		Libelle    : string ;
    Probleme   : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
  	if MessErreur = ErrGrilleDim1 then
    begin
	    Probleme := TPST.GetValue ('GDM_TYPE1');
      result   := 'La grille de tailles de la première dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim2 then
    begin
	    Probleme := TPST.GetValue ('GDM_TYPE2');
      result   := 'La grille de tailles de la deuxième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim3 then
    begin
	    Probleme := TPST.GetValue ('GDM_TYPE3');
      result   := 'La grille de tailles de la troisième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim4 then
    begin
	    Probleme := TPST.GetValue ('GDM_TYPE4');
      result   := 'La grille de tailles de la quatrième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrGrilleDim5 then
    begin
	    Probleme := TPST.GetValue ('GDM_TYPE5');
      result   := 'La grille de tailles de la cinquième dimension n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Masque de dimension n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeMasque := TPST.GetValue ('GDM_MASQUE');
  	Libelle    := TPST.GetValue ('GDM_LIBELLE');
  	result     := 'Le masque de dimension "' + CodeMasque + '" ' + Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB DISPO
////////////////////////////////////////////////////////
function  MessageTOBDispo (TPST : TOB; MessErreur : integer) : string ;
var CodeArticle : string ;
		CodeDepot		: string ;
    Probleme    : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GQ_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDepot then
    begin
	    Probleme := TPST.GetValue ('GQ_DEPOT');
      result   := 'Le dépôt n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Dispo n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeArticle := TPST.GetValue ('GQ_ARTICLE');
  	CodeDepot   := TPST.GetValue ('GQ_DEPOT');
  	result      := 'Le stock de l''article ' + CodeArticle + ' pour le dépôt ' + CodeDepot;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB ETABLISS
////////////////////////////////////////////////////////
function  MessageTOBEtabliss (TPST : TOB; MessErreur : integer) : string ;
var CodeEta  : string ;
		Libelle  : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
  	if MessErreur = ErrPays then
    begin
	    Probleme := TPST.GetValue ('ET_PAYS');
      result   := 'Le pays n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('ET_LANGUE');
      result   := 'La langue n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Etablissement n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeEta := TPST.GetValue ('ET_ETABLISSEMENT');
  	Libelle := TPST.GetValue ('ET_LIBELLE');
  	result  := 'La fiche de l''établissement "' + CodeEta + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB EXERCICE
////////////////////////////////////////////////////////
function  MessageTOBExercice (TPST : TOB; MessErreur : integer) : string ;
var CodeExe  : string ;
	Libelle  : string ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur EXERCICE n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeExe := TPST.GetValue ('EX_EXERCICE');
  	Libelle := TPST.GetValue ('EX_LIBELLE');
  	result  := 'L''exercice "' + CodeExe + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB FIDELITEENT
////////////////////////////////////////////////////////
function  MessageTOBFideliteEnt (TPST : TOB; MessErreur : integer) : string ;
var Tiers    : string ;
    Etab     : string ;
    Probleme : string ;
    NumCarte : string ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTiers then
    begin
      Probleme := TPST.GetValue ('GFE_TIERS');
      result   := 'Le tiers n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrEtablissement then
    begin
	  Probleme := TPST.GetValue ('GFE_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    result   := 'Erreur FIDELITEENT n°' + IntToStr (MessErreur);
  end else
  begin
    Tiers    := TPST.GetValue ('GFE_TIERS');
    Etab     := TPST.GetValue ('GFE_ETABLISSEMENT');
    NumCarte := TPST.GetValue ('GFE_NUMCARTEINT');
    result := 'La fidélité du client "' + Tiers + '" pour l''établissement ' + Etab + ' numéro ' + NumCarte;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB FIDELITELIG
////////////////////////////////////////////////////////
function  MessageTOBFideliteLig (TPST : TOB; MessErreur : integer) : string ;
var Carte    : string ;
	Etab     : string ;
    Probleme : string ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrEtablissement then
    begin
	  Probleme := TPST.GetValue ('GFI_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    result   := 'Erreur FIDELITELIG n°' + IntToStr (MessErreur);
  end else
  begin
    Carte  := TPST.GetValue ('GFI_NUMCARTEINT');
    Etab   := TPST.GetValue ('GFI_ETABLISSEMENT');
    result := 'La ligne de fidélité de la carte "' + Carte + '" pour l''établissement ' + Etab;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB FONCTION
////////////////////////////////////////////////////////
function  MessageTOBFonction (TPST : TOB; MessErreur : integer) : string ;
var Fonction  : string ;
	Libelle   : string ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur FONCTION n°' + IntToStr (MessErreur);
  end else
  begin
  	Fonction := TPST.GetValue ('AFO_FONCTION');
  	Libelle  := TPST.GetValue ('AFO_LIBELLE');
  	result   := 'La fonction "' + Fonction + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB JOURSCAISSE
////////////////////////////////////////////////////////
function  MessageTOBJoursCaisse (TPST : TOB; MessErreur : integer) : string ;
var CodeCaisse : string  ;
    NumOuvert  : integer ;
begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
  	result   := 'Erreur table JOURSCAISSE n°' + IntToStr (MessErreur);
  end else
  begin
    CodeCaisse := TPST.GetValue ('GJC_CAISSE');
    NumOuvert  := TPST.GetValue ('GJC_NUMZCAISSE');
  	result  := 'La journée n°' + IntToStr (NumOuvert) + ' de la caisse ' + CodeCaisse;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB JOURSETAB
////////////////////////////////////////////////////////
function  MessageTOBJoursEtab (TPST : TOB; MessErreur : integer) : string ;
var CodeEta  : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
  	if MessErreur = ErrPays then
    begin
	    Probleme := TPST.GetValue ('GJE_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrMeteo then
    begin
	    Probleme := TPST.GetValue ('GJE_METEO');
      result   := 'Le code météo n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur table JOURSETAB n°' + IntToStr (MessErreur);
  end else
  begin
    CodeEta := TPST.GetValue ('GJE_ETABLISSEMENT');
  	result  := 'La journée de l''établissement ' + CodeEta;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB JOURSETABEVT
////////////////////////////////////////////////////////
function MessageTOBJoursEtabEvt (TPST : TOB; MessErreur : integer) : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur table JOURSETABEVT n°' + IntToStr (MessErreur);
  end else
  begin
  	result  := 'L''évènement de la journée';
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB LIENSOLE
////////////////////////////////////////////////////////
function  MessageTOBLiensOLE (TPST : TOB; MessErreur : integer) : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur LIENSOLE n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'Le lien OLE ' + IntToStr(TPST.GetValue ('LO_RANGBLOB'));
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB LIGNE
////////////////////////////////////////////////////////
function  MessageTOBLigne (TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GL_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrArticleGen then
    begin
	    Probleme := TPST.GetValue ('GL_CODEARTICLE');
      result   := 'L''article générique n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTarifArticle then
    begin
	    Probleme := TPST.GetValue ('GL_TARIFARTICLE');
      result   := 'Le code tarif article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur LIGNE n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'La ligne ' + IntToStr(TPST.GetValue ('GL_NUMLIGNE'));
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB LIGNELOT
////////////////////////////////////////////////////////
function  MessageTOBLigneLot (TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GLL_ARTICLE');
      result   := 'L''article  n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur LIGNELOT n°' + IntToStr (MessErreur);
  end else
  begin
  		result := 'La ligne lot n°' + IntToStr(TPST.GetValue ('GLL_ARTICLE'));
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB LIGNENOMEN
////////////////////////////////////////////////////////
function  MessageTOBLigneNomen (TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GLN_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrArticleGen then
    begin
	    Probleme := TPST.GetValue ('GLN_CODEARTICLE');
      result   := 'L''article générique n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur LIGNENOMEN n°' + IntToStr (MessErreur);
  end else
  begin
 		result := 'La ligne nomenclature n°' + IntToStr(TPST.GetValue ('GLN_NUMLIGNE')) ;
  end;
end;

{$IFDEF BTP}
////////////////////////////////////////////////////////
//  Message TOB LIGNEOUV
////////////////////////////////////////////////////////
function  MessageTOBLigneOuvrages (TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('BLO_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrArticleGen then
    begin
	    Probleme := TPST.GetValue ('BLO_CODEARTICLE');
      result   := 'L''article générique n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur LIGNEOUV n°' + IntToStr (MessErreur);
  end else
  begin
 		result := 'La ligne nomenclature n°' + IntToStr(TPST.GetValue ('BLO_NUMLIGNE')) ;
  end;
end;

function MessageTOBMilliemes (TPST : TOB; MessErreur : integer) : string ;
//var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur BTPIECEMILIEME n°' + IntToStr (MessErreur);
  end else
  begin
 		result := 'La TVA Categorie '+ TPST.GetValue('BPM_CATEGORIETAXE')+' Famille '+ TPST.GetValue('BPM_FAMILLETAXE');
  end;
end;

function MessageTOBMemorisation      (TPST : TOB; MessErreur : integer) : string ;
//var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur BMEMORISATION n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'La Mémorisation ' + TPST.GetValue ('BMO_CODEMEMO');
  end;
end;

function MessageTOBParDoc           (TPST : TOB; MessErreur : integer) : string ;
//var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur BTPARDOC n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'Le paramdoc ' + TPST.GetValue ('BPD_NATUREPIECE')+ ' '+IntToStr(TPST.GetValue ('BPD_NUMPIECE'));
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB PIECERG
////////////////////////////////////////////////////////
function MessageTOBPieceRG (TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string   ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTypeRg then
    begin
	    Probleme := TPST.GetValue ('PBR_TYPERG');
      result   := 'Le type de retenue de garantie n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    result   := 'Erreur PIECERG n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'Les retenues de garantie du document';
  end;
end;

{$ENDIF}

////////////////////////////////////////////////////////
//  Message TOB LISTEINVENT
////////////////////////////////////////////////////////
function  MessageTOBListeInvEnt(TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDepot then
    begin
	  Probleme := TPST.GetValue ('GIE_DEPOT');
      result   := 'Le dépôt n''existe pas :"' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur LISTEINVENT n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'La liste d''inventaire ' + TPST.GetValue ('GIE_CODELISTE') + ' du dépôt ' + TPST.GetValue ('GIE_DEPOT');
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB LISTEINVLIG
////////////////////////////////////////////////////////
function  MessageTOBListeInvLig(TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string  ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
      Probleme := TPST.GetValue ('GIL_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDepot then
    begin
	  Probleme := TPST.GetValue ('GIL_DEPOT');
      result   := 'Le dépôt n''existe pas :"' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur LISTEINVLIG n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'La ligne de la liste d''inventaire ' + TPST.GetValue ('GIL_CODELISTE') + ' du dépôt ' + TPST.GetValue ('GIL_DEPOT');
  end;
end;


////////////////////////////////////////////////////////
//  Message TOB MEA
////////////////////////////////////////////////////////
function  MessageTOBMea (TPST : TOB; MessErreur : integer) : string ;
var Libelle      : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur MEA n°' + IntToStr (MessErreur);
  end else
  begin
  	Libelle      := TPST.GetValue ('GME_LIBELLE');
  	result       := 'Le qualifiant de mesure ' + Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB MODELES
////////////////////////////////////////////////////////
function  MessageTOBModeles (TPST : TOB; MessErreur : integer) : string ;
var Probleme : string ;
    Libelle  : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('MO_LANGUE');
      result   := 'La langue n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur MODELES n°' + IntToStr (MessErreur);
  end else
  begin
  	Libelle := TPST.GetValue ('MO_LIBELLE');
  	result  := 'Le modèle de document "' + Libelle + '"' ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB MODEDATA
////////////////////////////////////////////////////////
function  MessageTOBModeData (TPST : TOB; MessErreur : integer) : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur MODEDATA n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'Le contenu du modèle de document' ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB MODEREGL
////////////////////////////////////////////////////////
function  MessageTOBModeRegl (TPST : TOB; MessErreur : integer) : string ;
var CodeRegl : string ;
    Libelle  : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP1');
      result   := 'Le mode de paiement n°1 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP2');
      result   := 'Le mode de paiement n°2 n''existe pas"' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP3');
      result   := 'Le mode de paiement n°3 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP4');
      result   := 'Le mode de paiement n°4 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP5');
      result   := 'Le mode de paiement n°5 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP6');
      result   := 'Le mode de paiement n°6 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP7');
      result   := 'Le mode de paiement n°7 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP8');
      result   := 'Le mode de paiement n°8 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP9');
      result   := 'Le mode de paiement n°9 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP10');
      result   := 'Le mode de paiement n°10 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP11');
      result   := 'Le mode de paiement n°11 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModePaie1 then
    begin
	    Probleme := TPST.GetValue ('MR_MP12');
      result   := 'Le mode de paiement n°12 n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur MODEREGL n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeRegl := TPST.GetValue ('MR_MODEREGLE');
  	Libelle  := TPST.GetValue ('MR_LIBELLE');
  	result   := 'Le mode de règlement "' + CodeRegl + '" ' + Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB MODEPAIE
////////////////////////////////////////////////////////
function  MessageTOBModePaie (TPST : TOB; MessErreur : integer) : string ;
var CodePaie : string ;
    Libelle  : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur MODEPAIE n°' + IntToStr (MessErreur);
  end else
  begin
  	CodePaie := TPST.GetValue ('MP_MODEPAIE');
  	Libelle  := TPST.GetValue ('MP_LIBELLE');
  	result   := 'Le mode de paiement "' + CodePaie + '" ' + Libelle;
  end;
end;

function MessageTOBNaturePrest       (TPST : TOB; MessErreur : integer) : string ;
var Natureprest : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur NATUREPREST n°' + IntToStr (MessErreur);
  end else
  begin
  	Natureprest := TPST.GetValue ('BNP_NATUREPRES');
  	result   := 'La nature de prestation "' + Natureprest + '"';
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB NOMENENT
////////////////////////////////////////////////////////
function  MessageTOBNomenEnt (TPST : TOB; MessErreur : integer) : string ;
var Probleme : string ;
    CodeNom  : string ;
    Libelle  : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GNE_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur NOMENENT n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeNom := TPST.GetValue ('GNE_NOMENCLATURE');
  	Libelle := TPST.GetValue ('GNE_LIBELLE');
  	result   := 'La nomenclatue "' + CodeNom + '" ' + Libelle ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB NOMENLIG
////////////////////////////////////////////////////////
function  MessageTOBNomenLig (TPST : TOB; MessErreur : integer) : string ;
var Probleme : string ;
    CodeNom  : string ;
    Libelle  : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GLN_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur NOMENLIG n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeNom := TPST.GetValue ('GNL_NOMENCLATURE');
  	Libelle := TPST.GetValue ('GNL_LIBELLE');
  	result   := 'Le composant ' + Libelle + ' de la nomenclature ' + CodeNom;
  end;
end;


///////////////////////////////////////////////////////
//  Message TOB OPERCAISSE
////////////////////////////////////////////////////////
function  MessageTOBOperCaisse (TPST : TOB; MessErreur : integer) : string ;
var Nature   : string ;
    Numero   : integer;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrEtablissement then
    begin
	  Probleme := TPST.GetValue ('GOC_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	  Probleme := TPST.GetValue ('GOC_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrTiers then
    begin
	  Probleme := TPST.GetValue ('GOC_TIERS');
      result   := 'Le tiers n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    result := 'Erreur OPERCAISSE n°' + IntToStr (MessErreur);
  end else
  begin
  	Nature := TPST.GetValue ('GOC_NATUREPIECEG');
  	Numero := TPST.GetValue ('GOC_NUMERO');
  	result  := 'L''opération de caisse de nature ' + Nature  + ' numéro ' +  IntToStr (Numero) ;
  end;
end;

///////////////////////////////////////////////////////
//  Message TOB PARCAISSE
////////////////////////////////////////////////////////
function  MessageTOBParCaisse (TPST : TOB; MessErreur : integer) : string ;
var Caisse   : string ;
    Libelle  : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrEtablissement then
    begin
	    Probleme := TPST.GetValue ('GPK_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrDepot then
    begin
	    Probleme := TPST.GetValue ('GPK_DEPOT');
      result   := 'Le dépôt n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur PARCAISSE n°' + IntToStr (MessErreur);
  end else
  begin
  	Caisse  := TPST.GetValue ('GPK_CAISSE');
  	Libelle := TPST.GetValue ('GPK_LIBELLE');
  	result  := 'Le paramétrage de la caisse "' + Caisse + '" ' + Libelle;
  end;
end;

function MessageTOBParamSoc          (TPST : TOB; MessErreur : integer) : string ;
var ParamSoc   : string ;
//    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur PARAMSOC n°' + IntToStr (MessErreur);
  end else
  begin
  	ParamSoc  := TPST.GetValue ('SOC_NOM');
  	result  := 'Le paramétrage du ParamSoc "' + ParamSoc + '"';
  end;
end;

function MessageTOBParamTaxe         (TPST : TOB; MessErreur : integer) : string ;
var ParamTaxe   : string ;
//    Libelle  : string ;
//    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur PARAMTAXE n°' + IntToStr (MessErreur);
  end else
  begin
  	ParamTaxe  := TPST.GetValue ('BPT_CATEGORIETAXE');
  	result  := 'Le paramétrage du ParamTaxe "' + ParamTaxe + '"';
  end;
end;

///////////////////////////////////////////////////////
//  Message TOB PARFIDELITE
////////////////////////////////////////////////////////
function  MessageTOBParFidelite (TPST : TOB; MessErreur : integer) : string ;
var CodeFidelite : string ;
    Libelle      : string ;
    Probleme     : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrModePaie then
    begin
	  Probleme := TPST.GetValue ('GFO_MODEPAIE');
      result   := 'Le mode de paiement n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrTypeDemarque then
    begin
	  Probleme := TPST.GetValue ('GFO_TYPEREMISE');
      result   := 'Le type de remise n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur PARFIDELITE n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeFidelite := TPST.GetValue ('GFO_CODEFIDELITE');
  	Libelle      := TPST.GetValue ('GFO_LIBELLE');
  	result  := 'Le paramétrage de la fidelité "' + CodeFidelite + '" ' + Libelle;
  end;
end;

///////////////////////////////////////////////////////
//  Message TOB PARREGLEFID
////////////////////////////////////////////////////////
function  MessageTOBParRegleFid (TPST : TOB; MessErreur : integer) : string ;
var CodeFidelite : string ;
    Libelle      : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur PARREGLEFID n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeFidelite := TPST.GetValue ('GFR_CODEFIDELITE');
  	Libelle      := TPST.GetValue ('GFR_LIBELLE');
  	result  := 'Le paramétrage des règles de la fidélité "' + CodeFidelite + '" ' + Libelle;
  end;
end;

///////////////////////////////////////////////////////
//  Message TOB PARSEUILFID
////////////////////////////////////////////////////////
function  MessageTOBParSeuilFid (TPST : TOB; MessErreur : integer) : string ;
var CodeFidelite : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur PARSEUILFID n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeFidelite := TPST.GetValue ('GFS_CODEFIDELITE');
  	result  := 'Le paramétrage des seuils de la fidelité "' + CodeFidelite + '" ';
  end;
end;

///////////////////////////////////////////////////////
//  Message TOB PARPIECBIL
////////////////////////////////////////////////////////
function  MessageTOBParPiecBil (TPST : TOB; MessErreur : integer) : string ;
var Devise   : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('GPI_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    result := 'Erreur PARPIECBIL n°' + IntToStr (MessErreur);
  end else
  begin
  	Devise  := TPST.GetValue ('GPI_DEVISE');
  	result  := 'Les pièces et billets de la devise "' + Devise + '"';
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB PAYS
////////////////////////////////////////////////////////
function  MessageTOBPays (TPST : TOB; MessErreur : integer) : string ;
var CodePays : string ;
    Libelle  : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('PY_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('PY_LANGUE');
      result   := 'La langue n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur PAYS n°' + IntToStr (MessErreur);
  end else
  begin
  	CodePays := TPST.GetValue ('PY_PAYS');
  	Libelle  := TPST.GetValue ('PY_LIBELLE');
  	result   := 'Le pays "' + CodePays + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB PIEDBASE
////////////////////////////////////////////////////////
function MessageTOBPiedBase (TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string   ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('GPB_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrCategorieTaxe then
    begin
	    Probleme := TPST.GetValue ('GPB_CATEGORIETAXE');
      result   := 'La catégorie de taxe n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrFamilleTaxe then
    begin
	    Probleme := TPST.GetValue ('GPB_FAMILLETAXE');
      result   := 'La famille de taxe n''existe pas "' + Probleme + '"';
      exit ;
    end;
    result   := 'Erreur PIEDBASE n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'Le pied de taxe du document';
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB PiedEche
////////////////////////////////////////////////////////
function MessageTOBPiedEche (TPST : TOB; MessErreur : integer) : string ;
var Probleme     : string   ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrModePaie then
    begin
	    Probleme := TPST.GetValue ('GPE_MODEPAIE');
      result   := 'Le mode de paiement n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('GPE_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    if MessErreur = ErrDeviseEsp then
    begin
	    Probleme := TPST.GetValue ('GPE_DEVISEESP');
      result   := 'La devise du montant encaissé n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    result   := 'Erreur PIEDECHE n°' + IntToStr (MessErreur);
  end else
  begin
		result  := 'La ligne de règlement n°' + IntToStr(TPST.GetValue ('GPE_NUMECHE')) ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB PiedPort
////////////////////////////////////////////////////////
function MessageTOBPiedPort (TPST : TOB; MessErreur : integer) : string ;
var  NumPort : integer  ;
    Probleme : string   ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('GPE_DEVISE');
      result   := 'La devise  n''existe pas : "' + Probleme + '"';
      exit ;
    end;
    result   := 'Erreur PIEDPORT n°' + IntToStr (MessErreur);
  end else
  begin
  	NumPort := TPST.GetValue ('GPT_NUMPORT');
  	result       := 'Ligne de port n° ' + IntToStr(NumPort);
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB PIECE
////////////////////////////////////////////////////////
function  MessageTOBPIECE (TPST : TOB; MessErreur : integer) : string ;
var Nature       : string ;
    Souche       : string ;
    Numero       : integer;
    Etabliss     : string ;
    LibellePiece : string ;
    Probleme     : string ;
    SQL          : string ;
    Q            : TQUERY ;
begin
 	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrNaturePiece then
    begin
	    Probleme := TPST.GetValue ('GP_NATUREPIECEG');
      result   := 'La nature de piece n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTiersVide then
    begin
	  result   := 'Le tiers de la pièce n''est pas renseigné ! Il est obligatoire ';
      exit;
    end;
    if MessErreur = ErrTiers then
    begin
	    Probleme := TPST.GetValue ('GP_TIERS');
      result   := 'Le tiers n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTiersLivre then
    begin
	    Probleme := TPST.GetValue ('GP_TIERSLIVRE');
      result   := 'Le tiers livré n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTiersFacture then
    begin
	    Probleme := TPST.GetValue ('GP_TIERSFACTURE');
      result   := 'Le tiers facturé  n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTiersPayeur then
    begin
	    Probleme := TPST.GetValue ('GP_TIERSPAYEUR');
      result   := 'Le tiers payeur n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrRepresentant then
    begin
	    Probleme := TPST.GetValue ('GP_REPRESENTANT');
      result   := 'Le représentant n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrEtablissement then
    begin
	    Probleme := TPST.GetValue ('GP_ETABLISSEMENT');
      result   := 'L''établissement n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDepot then
    begin
	    Probleme := TPST.GetValue ('GP_DEPOT');
      result   := 'Le dépôt n''existe pas :"' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDepotDest then
    begin
	    Probleme := TPST.GetValue ('GP_DEPOTDEST');
      result   := 'Le dépôt destinataire n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModeRegle then
    begin
	    Probleme := TPST.GetValue ('GP_MODEREGLE');
      result   := 'Le mode de règlement n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrRegimeTVA then
    begin
	    Probleme := TPST.GetValue ('GP_REGIMETVA');
      result   := 'Le régime de TVA n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('GP_DEVISE');
      result   := 'La devise  n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrAnnulDoc then
    begin
      result   := 'Attention : problème lors de l''annulation de la pièce !';
      exit;
    end;
    if MessErreur = ErrLoadLesNomens then
    begin
      result   := 'Attention : problème lors du chargement du descriptif des nomenclatures !';
      exit;
    end;
{$IFDEF BTP}
    if MessErreur = ErrLoadLesOuvrages then
    begin
      result   := 'Attention : problème lors du chargement du descriptif des ouvrages !';
      exit;
    end;
{$ENDIF}
    if MessErreur = ErrLoadLesLots then
    begin
      result   := 'Attention : problème lors du chargement du descriptif des lots !';
      exit;
    end;
    if MessErreur = ErrCalculDispo then
    begin
      result   := 'Attention : problème lors du calcul de la MAJ des stocks !';
      exit;
    end;
    if MessErreur = ErrMAJDispo then
    begin
      result   := 'Attention : problème lors de la MAJ des stocks !';
      exit;
    end;
    if MessErreur = ErrMAJTiersPiece then
    begin
      result   := 'Attention : problème lors de la MAJ du tiers de la pièce !';
      exit;
    end;
    if MessErreur = ErrMAJAcomptes then
    begin
      result   := 'Attention : problème lors de la MAJ des acomptes !';
      exit;
    end;
    if MessErreur = ErrMAJPorts then
    begin
      result   := 'Attention : problème lors des lignes de ports !';
      exit;
    end;
    if MessErreur = ErrCalculInverseDispo then
    begin
      result   := 'Attention : problème lors du calcul de la MAJ des stocks pour la suppression de la pièce!';
      exit;
    end;
    if MessErreur = ErrMAJInverseDispo then
    begin
      result   := 'Attention : problème lors de la MAJ des stocks pour la suppression de la pièce!';
      exit;
    end;
    result   := 'Erreur PIECE n°' + IntToStr (MessErreur);
  end else
  begin
    Nature   := TPST.GetValue ('GP_NATUREPIECEG');
    Souche   := TPST.GetValue ('GP_SOUCHE');
    Numero   := TPST.GetValue ('GP_NUMERO');
    Etabliss := TPST.GetValue ('GP_ETABLISSEMENT');

    SQL:='Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="'+Nature+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
      LibellePiece := Q.FindField('GPP_LIBELLE').AsString ;
      result       := LibellePiece + ' n° ' + IntToStr(Numero) + ' pour le site '+Souche+' et de l''établissement ' + Etabliss;
    end else
    begin
    	result := 'L''entête de document de nature' + Nature + ' et de numéro ' + IntToStr(Numero);
    end;
    Ferme (Q);
  end;
end;


////////////////////////////////////////////////////////
//  Message TOB PORT
////////////////////////////////////////////////////////
function MessageTOBPieceAdresse (TPST : TOB; MessErreur : integer) : string ;
var Numero       : integer;
    Nature       : string ;
    LibellePiece : string ;
    SQL          : string ;
    Q            : TQUERY ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur PIECEADRESSE n°' + IntToStr (MessErreur);
  end else
  begin
  	Nature   := TPST.GetValue ('GPA_NATUREPIECEG');
  	Numero   := TPST.GetValue ('GPA_NUMERO');

    SQL:='Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="'+Nature+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
    begin
      LibellePiece := Q.FindField('GPP_LIBELLE').AsString ;
      result       := 'L''adresse de la pièce ' + LibellePiece + ' n° ' + IntToStr(Numero);
    end else
    begin
      result := 'L''adresse de la pièce ' + Nature + ' et de numéro ' + IntToStr(Numero);
    end;
    Ferme (Q);
  end;
end ;

function MessageTOBPprofilArt      (TPST : TOB; MessErreur : integer) : string ;
var Numero       : integer;
    ProfilArt       : string ;
    LibellePiece : string ;
    SQL          : string ;
    Q            : TQUERY ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur PROFILART n°' + IntToStr (MessErreur);
  end else
  begin
  	ProfilArt   := TPST.GetValue ('GPF_PROFILARTICLE');
    result := 'Le profil article ' + ProfilArt ;
  end;
end ;

////////////////////////////////////////////////////////
//  Message TOB PORT
////////////////////////////////////////////////////////
function MessageTOBPort (TPST : TOB; MessErreur : integer) : string ;
var CodePort,Libelle,Probleme : string ;
begin
  if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTypePort then
    begin
	  Probleme := TPST.GetValue ('GPO_TYPEPORT');
      result   := 'Le type de port n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur PORT n°' + IntToStr (MessErreur);
  end else
  begin
  	CodePort := TPST.GetValue ('GPO_CODEPORT');
  	Libelle  := TPST.GetValue ('GPO_LIBELLE');
  	result   := 'Le port "' + CodePort + '" ' + Libelle;
  end;
end ;

///////////////////////////////////////////////////////
//  Message TOB STOXQUERYS
////////////////////////////////////////////////////////
function MessageTOBStoxQuerys (TPST : TOB; MessErreur : integer) : string ;
var CodeRequete : string ;
    Libelle     : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur STOXQUERYS n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeRequete := TPST.GetValue ('SQE_CODEREQUETE');
  	Libelle     := TPST.GetValue ('SQE_LIBELLE');
  	result      := 'La requête TOX "' + CodeRequete + '" ' + Libelle;
  end;
end ;

///////////////////////////////////////////////////////
//  Message TOB REGION
////////////////////////////////////////////////////////
function MessageTOBRegion (TPST : TOB; MessErreur : integer) : string ;
var CodePays   : string ;
    CodeRegion : string ;
    Libelle    : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur REGION n°' + IntToStr (MessErreur);
  end else
  begin
    CodePays   := TPST.GetValue ('RG_PAYS');
  	CodeRegion := TPST.GetValue ('RG_REGION');
  	Libelle    := TPST.GetValue ('RG_LIBELLE');
  	result     := 'La région "' + CodeRegion + '" ' + Libelle + ' du pays ' + CodePays;
  end;
end ;

///////////////////////////////////////////////////////
//  Message TOB SOCIETE
////////////////////////////////////////////////////////
function MessageTOBSociete (TPST : TOB; MessErreur : integer) : string ;
var CodeSociete : string ;
    Libelle     : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur SOCIETE n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeSociete := TPST.GetValue ('SO_SOCIETE');
  	Libelle     := TPST.GetValue ('SO_LIBELLE');
  	result      := 'Le numéro de version PGI "' + CodeSociete + '" ' + Libelle;
  end;
end ;

////////////////////////////////////////////////////////
//  Message TOB TARIFS
////////////////////////////////////////////////////////
function  MessageTOBTarif (TPST : TOB; MessErreur : integer) : string ;
var CodeTarif : string ;
    Libelle   : string ;
    Probleme  : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrArticle then
    begin
	    Probleme := TPST.GetValue ('GF_ARTICLE');
      result   := 'L''article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDepot then
    begin
	    Probleme := TPST.GetValue ('GF_DEPOT');
      result   := 'Le dépôt n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('GF_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrClient then
    begin
	    Probleme := TPST.GetValue ('GF_TIERS');
      result   := 'Le client n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrFournisseur then
    begin
	    Probleme := TPST.GetValue ('GF_TIERS');
      result   := 'Le fournisseur n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTarifArticle then
    begin
	    Probleme := TPST.GetValue ('GF_TARIFARTICLE');
      result   := 'Le tarif article n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur TARIF n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeTarif := TPST.GetValue ('GF_TARIF');
  	Libelle   := TPST.GetValue ('GF_LIBELLE');
  	result    := 'Le tarif "' + CodeTarif + '" ' + Libelle;
  end;
end;

//
//  Message TOB TARIFMODE
//
function  MessageTOBTarifMode (TPST : TOB; MessErreur : integer) : string ;
var TarifMode,Libelle,Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTypeArrondi then
    begin
	    Probleme := TPST.GetValue ('GFM_ARRONDI');
      result   := 'Le type d''arrondi n''existe pas : "' + Probleme + '"';
      exit;
    end ;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('GFM_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit;
    end ;
    if MessErreur = ErrTypeDemarque then
    begin
	    Probleme := TPST.GetValue ('GFM_DEMARQUE');
      result   := 'Le type de démarque n''existe pas : "' + Probleme + '"';
      exit;
    end
  end else
  begin
  	TarifMode := TPST.GetValue ('GFM_TARFMODE');
  	Libelle   := TPST.GetValue ('GFM_LIBELLE');
  	result    := 'Le tarif mode "' + TarifMode + '" ' + Libelle;
  end;
end ;

//
//  Message TOB TARIFPER
//
function  MessageTOBTarifPer (TPST : TOB; MessErreur : integer) : string ;
var CodePeriode,Libelle,Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTypeArrondi then
    begin
	    Probleme := TPST.GetValue ('GFP_ARRONDI');
      result   := 'Le type d''arrondi n''existe pas : "' + Probleme + '"';
      exit;
    end ;
    if MessErreur = ErrTypeDemarque then
    begin
	    Probleme := TPST.GetValue ('GFP_DEMARQUE');
      result   := 'Le type de démarque n''existe pas : "' + Probleme + '"';
      exit;
    end
  end else
  begin
  	CodePeriode := TPST.GetValue ('GFP_CODEPERIODE');
  	Libelle     := TPST.GetValue ('GFP_LIBELLE');
  	result      := 'Le tarif période "' + CodePeriode + '" ' + Libelle;
  end;
end ;

function  MessageTOBTarifTypMode (TPST : TOB; MessErreur : integer) : string ;
var CodeType,Libelle,Probleme : string ;
begin
  if MessErreur >= 0 then
  begin
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDevise then
    begin
      Probleme := TPST.GetValue ('GFT_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit;
    end
  end else
  begin
    CodeType := TPST.GetValue ('GFT_CODETYPE');
    Libelle  := TPST.GetValue ('GFT_LIBELLE');
    result   := 'Le tarif de type "' + CodeType + '" ' + Libelle;
  end;
end ;

function MessageTOByTarifs(TPST : TOB; MessErreur : integer) : string ;
var Libelle,Probleme : string ;
begin
  if MessErreur >= 0 then
  begin
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTiers then
    begin
      Probleme := TPST.GetValue ('YTS_TIERS');
      result   := 'Le tiers "'+ Probleme +'" n''existe pas';
    end;
    if MessErreur = ErrAffaire then
    begin
      Probleme := TPST.GetValue ('YTS_AFFAIRE');
      result   := 'L''affaire "'+ Probleme +'" n''existe pas';
      exit;
    end;
    if MessErreur = ErrArticle then
    begin
      Probleme := TPST.GetValue ('YTS_ARTICLE');
      result   := 'L''article "'+ Probleme +'" n''existe pas';
      exit;
    end;
    if MessErreur = ErrDepot then
    begin
      Probleme := TPST.GetValue ('YTS_DEPOT');
      result   := 'Le dépôt "'+ Probleme +'" n''existe pas';
      exit;
    end;
  end else
  begin
    Libelle  := TPST.GetValue ('YTS_LIBELLETARIF');
    result   := 'Le tarif "' + Libelle + '"';
  end;
end;

function MessageTOByTarifsFourchette (TPST : TOB; MessErreur : integer) : string ;
var CodeType,Libelle,Probleme : string ;
begin
  if MessErreur >= 0 then
  begin
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrTarif then
    begin
      Probleme := TPST.GetValue ('YTF_IDENTIFIANTYTS');
      result   := 'Le tarif n''existe pas : "' + Probleme + '"';
      exit;
    end
  end else
  begin
    CodeType := TPST.GetValue ('YTF_IDENTIFIANT');
    Libelle  := TPST.GetValue ('YTF_LIBELLETARIF');
    result   := 'Le tarif de type "' + CodeType + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB TIERS
////////////////////////////////////////////////////////
function  MessageTOBTiers (TPST : TOB; MessErreur : integer) : string ;
var Auxiliaire : string ;
		TypeTiers  : string ;
    Libelle    : string ;
    Probleme   : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrPays then
    begin
	    Probleme := TPST.GetValue ('T_PAYS');
      result   := 'Le pays n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('T_LANGUE');
      result   := 'La langue n''existe pas :"' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDevise then
    begin
	    Probleme := TPST.GetValue ('T_DEVISE');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrDepot then
    begin
	    Probleme := TPST.GetValue ('T_DEPOT');
      result   := 'La devise n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrSecteurAct then
    begin
	    Probleme := TPST.GetValue ('T_SECTEUR');
      result   := 'Le secteur d''activité n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrZoneCom then
    begin
	    Probleme := TPST.GetValue ('T_ZONECOM');
      result   := 'La zone commerciale n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrTarifTiers then
    begin
	    Probleme := TPST.GetValue ('T_TARIFTIERS');
      result   := 'Le code tarifs tiers n''existe pas : ' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrModeRegle then
    begin
	    Probleme := TPST.GetValue ('T_MODEREGLE');
      result   := 'Le mode de règlement n''existe pas :"' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrRegimeTVA then
    begin
	    Probleme := TPST.GetValue ('T_REGIMETVA');
      result   := 'Le régime de TVA n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrCivilite then
    begin
	    Probleme := TPST.GetValue ('T_JURIDIQUE');
      result   := 'La civilité du client n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrNationalite then
    begin
	    Probleme := TPST.GetValue ('T_NATIONALITE');
      result   := 'La nationalité n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrFamComptable then
    begin
	    Probleme := TPST.GetValue ('T_COMPTATIERS');
      result   := 'La famille comptable n''existe pas :"' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrSexe then
    begin
	    Probleme := TPST.GetValue ('T_SEXE');
      result   := 'Le sexe n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur TIERS n°' + IntToStr (MessErreur);
  end else
  begin
  	Auxiliaire := TPST.GetValue ('T_AUXILIAIRE');
  	TypeTiers  := TPST.GetValue ('T_NATUREAUXI');
  	Libelle    := TPST.GetValue ('T_LIBELLE');
  	if      (TypeTiers = 'CLI') then result:='La fiche du client ' + Auxiliaire + ' ' + Libelle
  	else if (TypeTiers = 'FOU') then result:='La fiche du fournisseur ' + Auxiliaire + ' ' + Libelle
    else if (TypeTiers = 'PRO') then result:='La fiche du prospect ' + Auxiliaire + ' ' + Libelle
  	else                             result:='La fiche du tiers ' + Auxiliaire + ' ' + Libelle ;
  end;
end;

function  MessageTOBTiersCompl (TPST : TOB; MessErreur : integer) : string ;
var Auxiliaire : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur TIERS n°' + IntToStr (MessErreur);
  end else
  begin
  	Auxiliaire := TPST.GetValue ('YTC_AUXILIAIRE');
    result:='La fiche du tiers ' + Auxiliaire ;
  end;
end;

//
// Message TOB TRADDICO
//
function MessageTOBTradDico     (TPST : TOB; MessErreur : integer) : string ;
var CodeDico,Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDico then
    begin
	    Probleme := TPST.GetValue ('DX_FRA');
      result   := 'Le libellé n''existe pas : "' + Probleme + '"';
      exit;
    end
  end else
  begin
  	CodeDico := TPST.GetValue ('DX_FRA');
  	result   := 'Le dictionnaire "' + CodeDico + '" ' ;
  end;
end ;

//
// Message TRADTABLETTE
//
function MessageTOBTradTablette (TPST : TOB; MessErreur : integer) : string ;
var sLangue, sTypeTablette, sCodeTablette ,Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end ;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('YTT_LANGUE');
      result   := 'La langue "'+ Probleme +'" n''existe pas.' ;
      exit;
    end ;
    if MessErreur = ErrCodeTablette then
    begin
	    Probleme := TPST.GetValue ('YTT_CODE');
      result   := 'La tablette de code "'+ Probleme + '" n''existe pas.' ;
      exit;
    end
  end else
  begin
  	sLangue       := TPST.GetValue ('YTT_LANGUE');
    sTypeTablette := TPST.GetValue ('YTT_TYPE') ;
    sCodeTablette := TPST.GetValue ('YTT_CODE') ;
  	result        := 'La tablette de type "'+sTypeTablette+'" de code "'+sCodeTablette+'" en langue "'+sLangue+'"' ;
  end;
end ;

//
// Message TRADUC
//
function MessageTOBTraduc (TPST : TOB; MessErreur : integer) : string ;
var sLangue, sForme, Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end ;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('TR_LANGUE');
      result   := 'La langue "'+ Probleme + '" n''existe pas.' ;
      exit;
    end ;
    if MessErreur = ErrForme then
    begin
	    Probleme := TPST.GetValue ('TR_FORME');
      result   := 'La fiche "'+ Probleme +'" n''existe pas.' ;
      exit;
    end
  end else
  begin
  	sLangue := TPST.GetValue ('TR_LANGUE');
    sForme  := TPST.GetValue ('TR_FORME') ;
  	result  := 'La fiche "'+sForme+'" en "'+sLangue+'"' ;
  end;
end ;

//
// Message TXCPTTVA
//
function MessageTOBTxCptTva (TPST : TOB; MessErreur : integer) : string ;
var Regime   : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end ;
    if MessErreur = ErrRegime then
    begin
	    Probleme := TPST.GetValue ('TV_REGIME');
      result   := 'Le régime fiscal "'+ Probleme + '" n''existe pas.' ;
      exit;
    end ;
  end else
  begin
  	Regime := TPST.GetValue ('TV_REGIME');
  	result  := 'La taxe du régime fiscal "'+Regime+'"' ;
  end;
end ;

////////////////////////////////////////////////////////
//  Message TOB TYPEMASQUE
////////////////////////////////////////////////////////
function  MessageTOBTypeMasque (TPST : TOB; MessErreur : integer) : string ;
var	TypeMasque : string ;
	Libelle    : string ;
begin
  if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur TYPEMASQUE n°' + IntToStr (MessErreur);
  end else
  begin
  	TypeMasque := TPST.GetValue ('GMQ_TYPEMASQUE');
  	Libelle    := TPST.GetValue ('GMQ_LIBELLE');
  	result     := 'Le type de masque "' + TypeMasque + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB TYPEMASQUE
////////////////////////////////////////////////////////
function  MessageTOBTypeRemise (TPST : TOB; MessErreur : integer) : string ;
var	TypeRemise : string ;
	Libelle    : string ;
begin
  if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur TYPEREMISE n°' + IntToStr (MessErreur);
  end else
  begin
  	TypeRemise := TPST.GetValue ('GTR_TYPEREMISE');
  	Libelle    := TPST.GetValue ('GTR_LIBELLE');
  	result     := 'Le type de remise "' + TypeRemise + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB USERGRP
////////////////////////////////////////////////////////
function  MessageTOBUserGrp (TPST : TOB; MessErreur : integer) : string ;
var	Groupe 	: string ;
		Libelle : string ;

begin
  if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur USERGRP n°' + IntToStr (MessErreur);
  end else
  begin
    Groupe  := TPST.GetValue ('UG_GROUPE');
    Libelle := TPST.GetValue ('UG_LIBELLE');
    result  := 'Le groupe d''utilisateur "' + Groupe + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB UTILISAT
////////////////////////////////////////////////////////
function  MessageTOBUtilisat (TPST : TOB; MessErreur : integer) : string ;
var	Utilisateur	: string ;
	Libelle     : string ;

begin
  if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur UTILISAT n°' + IntToStr (MessErreur);
  end else
  begin
  	Utilisateur := TPST.GetValue ('US_UTILISATEUR');
  	Libelle     := TPST.GetValue ('US_LIBELLE');
  	result      := 'L''utilisateur "' + Utilisateur + '" ' + Libelle;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Descriptif des informations intégrées
///////////////////////////////////////////////////////////////////////////////////////
function ToxEditeMessage (NomTable : string; ToxInfo : TOB; MessErreur : integer) : string;
var MessageInfo : string;
begin
  if MessErreur = EnregNotTrt then MessageInfo := 'L''enregistrement n''a pas encore été traité'
  else if MessErreur = EnregNotTrtErr      then MessageInfo := 'Une erreur a été détectée dans l''unité transactionnelle'
  else if (NomTable = 'AFFAIRE')          then MessageInfo := MessageTOBAffaire     (ToxInfo, MessErreur)
  else if (NomTable = 'ACOMPTES')          then MessageInfo := MessageTOBAcomptes     (ToxInfo, MessErreur)
  else if (NomTable = 'ADRESSES')          then MessageInfo := MessageTOBAdresses     (ToxInfo, MessErreur)
  else if (NomTable = 'ARRONDI')	        then MessageInfo := MessageTOBArrondi      (ToxInfo, MessErreur)
  else if (NomTable = 'ARTICLE')	        then MessageInfo := MessageTOBArticle      (ToxInfo, MessErreur)
  else if (NomTable = 'ARTICLECOMPL')      then MessageInfo := MessageTOBArticleCompl (ToxInfo, MessErreur)
  else if (NomTable = 'ARTICLELIE')        then MessageInfo := MessageTOBArticleLie   (ToxInfo, MessErreur)
  else if (NomTable = 'ARTICLEPIECE')      then MessageInfo := MessageTOBArticlePiece (ToxInfo, MessErreur)
{$IFDEF BTP}
  else if (NomTable = 'BMEMORISATION')      then MessageInfo := MessageTOBMemorisation (ToxInfo, MessErreur)
  else if (NomTable = 'BTPARDOC')          then MessageInfo := MessageTOBParDoc (ToxInfo, MessErreur)
{$ENDIF}
  else if (NomTable = 'ARTICLETIERS')      then MessageInfo := MessageTOBArticleTiers (ToxInfo, MessErreur)
  else if (NomTable = 'CATALOGU')	        then MessageInfo := MessageTOBCatalogu     (ToxInfo, MessErreur)
  else if (NomTable = 'CHANCELL')	        then MessageInfo := MessageTOBChancell     (ToxInfo, MessErreur)
  else if (NomTable = 'CHOIXCOD')	        then MessageInfo := MessageTOBChoixCod     (ToxInfo, MessErreur)
  else if (NomTable = 'CHOIXEXT')	        then MessageInfo := MessageTOBChoixExt     (ToxInfo, MessErreur)
  else if (NomTable = 'CLAVIERECRAN')      then MessageInfo := MessageTOBClavierEcran (ToxInfo, MessErreur)
  else if (NomTable = 'COMMENTAIRE')       then MessageInfo := MessageTOBCommentaire  (ToxInfo, MessErreur)
  else if (NomTable = 'COMMERCIAL')        then MessageInfo := MessageTOBCommercial   (ToxInfo, MessErreur)
  else if (NomTable = 'COMPTEURETAB')      then MessageInfo := MessageTOBCompteurEtab (ToxInfo, MessErreur)
  else if (NomTable = 'CONTACT')	        then MessageInfo := MessageTOBContact      (ToxInfo, MessErreur)
  else if (NomTable = 'CTRLCAISMT')        then MessageInfo := MessageTOBCtrlCaisMt   (ToxInfo, MessErreur)
  else if (NomTable = 'DEPOTS')	        then MessageInfo := MessageTOBDepots       (ToxInfo, MessErreur)
  else if (NomTable = 'DEVISE')	        then MessageInfo := MessageTOBDevise       (ToxInfo, MessErreur)
  else if (NomTable = 'DIMMASQUE')         then MessageInfo := MessageTOBDimasque     (ToxInfo, MessErreur)
  else if (NomTable = 'DIMENSION')         then MessageInfo := MessageTOBDimension    (ToxInfo, MessErreur)
  else if (NomTable = 'DISPO')	        then MessageInfo := MessageTOBDispo        (ToxInfo, MessErreur)
  else if (NomTable = 'ETABLISS')	        then MessageInfo := MessageTOBEtabliss     (ToxInfo, MessErreur)
  else if (NomTable = 'EXERCICE')	        then MessageInfo := MessageTOBExercice     (ToxInfo, MessErreur)
  else if (NomTable = 'FIDELITEENT')       then MessageInfo := MessageTOBFideliteEnt  (ToxInfo, MessErreur)
  else if (NomTable = 'FIDELITELIG')       then MessageInfo := MessageTOBFideliteLig  (ToxInfo, MessErreur)
  else if (NomTable = 'FONCTION')	        then MessageInfo := MessageTOBFonction     (ToxInfo, MessErreur)
  else if (NomTable = 'JOURSCAISSE')       then MessageInfo := MessageTOBJoursCaisse  (ToxInfo, MessErreur)
  else if (NomTable = 'JOURSETAB')         then MessageInfo := MessageTOBJoursEtab    (ToxInfo, MessErreur)
  else if (NomTable = 'JOURSETABEVT')      then MessageInfo := MessageTOBJoursEtabEvt (ToxInfo, MessErreur)
  else if (NomTable = 'LIENSOLE')	        then MessageInfo := MessageTOBLiensOLE     (ToxInfo, MessErreur)
  else if (NomTable = 'LIGNE')	        then MessageInfo := MessageTOBLigne        (ToxInfo, MessErreur)
  else if (NomTable = 'LIGNELOT')	        then MessageInfo := MessageTOBLigneLot     (ToxInfo, MessErreur)
  else if (NomTable = 'LIGNENOMEN')        then MessageInfo := MessageTOBLigneNomen   (ToxInfo, MessErreur)
{$IFDEF BTP}
  else if (NomTable = 'LIGNEOUV')        then MessageInfo := MessageTOBLigneOuvrages   (ToxInfo, MessErreur)
  else if (NomTable = 'BTPIECEMILIEME')  then MessageInfo := MessageTOBMilliemes   (ToxInfo, MessErreur)
  else if (NomTable = 'PIECERG')         then MessageInfo := MessageTOBPieceRG   (ToxInfo, MessErreur)
{$ENDIF}
  else if (NomTable = 'LISTEINVENT')       then MessageInfo := MessageTOBListeInvEnt  (ToxInfo, MessErreur)
  else if (NomTable = 'LISTEINVLIG')       then MessageInfo := MessageTOBListeInvLig  (ToxInfo, MessErreur)
  else if (NomTable = 'MEA')	        then MessageInfo := MessageTOBMea          (ToxInfo, MessErreur)
  else if (NomTable = 'MODELES')	        then MessageInfo := MessageTOBModeles      (ToxInfo, MessErreur)
  else if (NomTable = 'MODEDATA')	        then MessageInfo := MessageTOBModeData     (ToxInfo, MessErreur)
  else if (NomTable = 'MODEREGL')	        then MessageInfo := MessageTOBModeRegl     (ToxInfo, MessErreur)
  else if (NomTable = 'MODEPAIE')	        then MessageInfo := MessageTOBModePaie     (ToxInfo, MessErreur)
  else if (NomTable = 'NATUREPREST')      then MessageInfo := MessageTOBNaturePrest  (ToxInfo, MessErreur)
  else if (NomTable = 'NOMENENT')          then MessageInfo := MessageTOBNomenEnt     (ToxInfo, MessErreur)
  else if (NomTable = 'NOMENLIG')	        then MessageInfo := MessageTOBNomenLig     (ToxInfo, MessErreur)
  else if (NomTable = 'OPERCAISSE')        then MessageInfo := MessageTOBOperCaisse   (ToxInfo, MessErreur)
  else if (NomTable = 'PARCAISSE')         then MessageInfo := MessageTOBParCaisse    (ToxInfo, MessErreur)
  else if (NomTable = 'PARAMSOC')         then MessageInfo := MessageTOBParamSoc    (ToxInfo, MessErreur)
  else if (NomTable = 'PARAMTAXE')         then MessageInfo := MessageTOBParamTaxe    (ToxInfo, MessErreur)
  else if (NomTable = 'PARFIDELITE')       then MessageInfo := MessageTOBParFidelite  (ToxInfo, MessErreur)
  else if (NomTable = 'PARREGLEFID')       then MessageInfo := MessageTOBParRegleFid  (ToxInfo, MessErreur)
  else if (NomTable = 'PARSEUILFID')       then MessageInfo := MessageTOBParSeuilFid  (ToxInfo, MessErreur)
  else if (NomTable = 'PARPIECBIL')        then MessageInfo := MessageTOBParPiecBil   (ToxInfo, MessErreur)
  else if (NomTable = 'PAYS')	        then MessageInfo := MessageTOBPays         (ToxInfo, MessErreur)
  else if (NomTable = 'PIEDBASE')	        then MessageInfo := MessageTOBPiedBase     (ToxInfo, MessErreur)
  else if (NomTable = 'PIEDECHE')	        then MessageInfo := MessageTOBPiedEche     (ToxInfo, MessErreur)
  else if (NomTable = 'PIEDPORT')	        then MessageInfo := MessageTOBPiedPort     (ToxInfo, MessErreur)
  else if (NomTable = 'PIECE')	        then MessageInfo := MessageTOBPiece        (ToxInfo, MessErreur)
{$IFDEF BTP}
  else if (NomTable = 'PIECERG')           then MessageInfo := MessageTOBPieceRG      (ToxInfo, MessErreur)
{$ENDIF}
  else if (NomTable = 'PORT')	        then MessageInfo := MessageTOBPort         (ToxInfo, MessErreur)
  else if (NomTable = 'PIECEADRESSE')      then MessageInfo := MessageTOBPieceAdresse (ToxInfo, MessErreur)
  else if (NomTable = 'PROFILART')         then MessageInfo := MessageTOBPprofilArt (ToxInfo, MessErreur)
  else if (NomTable = 'REGION')            then MessageInfo := MessageTOBRegion       (ToxInfo, MessErreur)
  else if (NomTable = 'SOCIETE')           then MessageInfo := MessageTOBSociete      (ToxInfo, MessErreur)
  else if (NomTable = 'STOXQUERYS')        then MessageInfo := MessageTOBStoxQuerys   (ToxInfo, MessErreur)
  else if (NomTable = 'TARIF')	        then MessageInfo := MessageTOBTarif        (ToxInfo, MessErreur)
  else if (NomTable = 'TARIFMODE')         then MessageInfo := MessageTOBTarifMode    (ToxInfo, MessErreur)
  else if (NomTable = 'TARIFPER')          then MessageInfo := MessageTOBTarifPer     (ToxInfo, MessErreur)
  else if (NomTable = 'TARIFTYPMODE')      then MessageInfo := MessageTOBTarifTypMode (ToxInfo, MessErreur)
  else if (NomTable = 'YTARIFS')           then MessageInfo := MessageTOByTarifs      (ToxInfo, MessErreur)
  else if (NomTable = 'YTARIFSFOURCHETTE') then MessageInfo := MessageTOByTarifsFourchette(ToxInfo, MessErreur)
  else if (NomTable = 'TARIFTYPMODE')      then MessageInfo := MessageTOBTarifTypMode (ToxInfo, MessErreur)
  else if (NomTable = 'TIERS')	        then MessageInfo := MessageTOBTiers        (ToxInfo, MessErreur)
  else if (NomTable = 'TIERSCOMPL')        then MessageInfo := MessageTOBTiersCompl   (ToxInfo, MessErreur)
  else if (NomTable = 'TRADDICO')	        then MessageInfo := MessageTOBTradDico     (ToxInfo, MessErreur)
  else if (NomTable = 'TRADTABLETTE')      then MessageInfo := MessageTOBTradTablette (ToxInfo, MessErreur)
  else if (NomTable = 'TRADUC')	        then MessageInfo := MessageTOBTraduc       (ToxInfo, MessErreur)
  else if (NomTable = 'TXCPTTVA')          then MessageInfo := MessageTOBTxCptTva     (ToxInfo, MessErreur)
  else if (NomTable = 'TYPEMASQUE')        then MessageInfo := MessageTOBTypeMasque   (ToxInfo, MessErreur)
  else if (NomTable = 'TYPEREMISE')        then MessageInfo := MessageTOBTypeRemise   (ToxInfo, MessErreur)
  else if (NomTable = 'USERGRP')           then MessageInfo := MessageTOBUserGrp      (ToxInfo, MessErreur)
  else if (NomTable = 'UTILISAT')          then MessageInfo := MessageTOBUtilisat     (ToxInfo, MessErreur)
  else MessageInfo:='L''intégration de la table '+ Nomtable + ' n''est pas gérée' ;
  result:=MessageInfo;
end;

end.

