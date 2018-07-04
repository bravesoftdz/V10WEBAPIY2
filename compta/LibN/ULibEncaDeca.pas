unit ULibEncaDeca;

interface

uses
  Hctrls
  ,Classes
  {$IFDEF EAGLCLIENT}
    ,UtileAGL    // Pour LanceEtatTob
    ,HPdfPrev    // pour gestion du spooler
   {$ELSE}
    ,db
    {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
    ,EdtREtat    // Pour LanceEtatTob
    ,EdtRDoc     // Pour LanceDocumentSpool
    ,HPdfPrev    // pour gestion du spooler
    ,uPDFBatch   // pour gestion du spooler
  {$ENDIF}
  ,ParamSoc      // pour le GetParamSoc
  ,Controls      // pour le mrYes
  ,UTob
  ,HEnt1
  ,Ent1
  ,SysUtils
  ,ULibEcriture   // Pour fns de base sur tob écriture
  ,ULibAnalytique // Pour fns de base sur tob analytique
  ,SaisComm       // Pour TSorteMontant, OBMToIdent
  ,Saisutil       // pour GetNewNumJal
  ,LetBatch       // pour le lettrage
  ,LettUtil       // Pour le TL_Rappro, et formatCheque
  ,Ed_Tools       // pour VideListe
  ,UtilPGI        // pour EncodeRIB
  ,hmsgbox        // pour MessageAlerte
  ,utilSais       // pour MajSoldesEcritureTOB
  ,Math           // Pour IntPower
  ,LP_VIEW        // pour ExporteLP
  ,uLibPieceCompta// TPieceCompta
  ,CPSaisiePiece_Tof // ModifiePieceCompta
  ,Constantes        // ets_Rien
  ,MC_Erreur
  ;


//================================================
//===== GESTION DES OPERATIONS D'ENCA / DECA =====
//================================================

Type TTypeAcces = ( taJamais, taPossible, taObligatoire ) ;

Type TTypeEncaDeca = Record
               Code            : String ;       // Code de l'opération d'enca/déca
               AccesCompta     : TTypeAcces ;   // Génération de pièces comptables
               AccesEscompte   : TTypeAcces ;   // Gestion de l'escompte
               AccesEdition    : TTypeAcces ;   // Edition
               AccesCFONB      : TTypeAcces ;   // Export CFONB
               AccesBordereau  : TTypeAcces ;   // Edition bordereau
               AccesMethode    : String ;       // Liste des méthodes de génération possible
               AccesJournaux   : String ;       // Tablette de journaux autorisés
               AccesGeneraux   : String ;       // Tablette de comptes généraux autorisés
               NatureEtat      : String ;       // Nature des états associés
               NatureDoc       : String ;       // Nature des documents associés
               AccesCondRemise : Boolean ;
               end ;

Type TErrorEncaDeca = Class
       private
         FMsgCompta : TMessageCompta ;
         procedure Initialisation ;
       public
         // Construteurs / Destructeurs
         Class function CreerMsgErreur : TErrorEncaDeca ;
         // Construteurs / Destructeurs
         Destructor Destroy ; override ;

         procedure OnError (sender : TObject; Error : TRecError ) ;
       end ;


Function ChargeTypeEncaDeca ( vStCode : String ) : TTypeEncaDeca ;
Function CheminSpoolerPourEncaDeca ( vType : TTypeEncaDeca ) : String ;
Function TriPourGroupeEncaDeca( vStCode : String ; avecLot : Boolean ) : String ;

//===================================================
//===== CODE ERREUR DES SCENARIOS D'ENCA / DECA =====
//===================================================
Const
  CPG_PASERREUR           = 0 ;
  CPG_ERRUNICITE          = 1 ;
  CPG_ERRCODE             = 2 ;
  CPG_ERRLIBELLE          = 3 ;
  CPG_ERRMODELE           = 4 ;
  CPG_ERRTYPE             = 5 ;
  CPG_ERRFLUX             = 6 ;
  CPG_ERRJOURNAL          = 7 ;
  CPG_ERRGENERAL          = 8 ;
  CPG_ERRGROUPE           = 9 ;
  CPG_ERRMETHANA          = 10 ;
  CPG_ERRESCMETH          = 11 ;
  CPG_ERRESCCPT           = 12 ;
  CPG_ERRESCTAUX          = 13 ;
  CPG_ERRESCTVACPT        = 14 ;
  CPG_ERRESCTVATAUX       = 15 ;
  CPG_ERRMODEPAIE         = 16 ;
  CPG_ERRCFONBFORMAT      = 17 ;
  CPG_ERRBORDEREAUMOD     = 18 ;
  CPG_ERRGENERALSEL       = 19 ;
  CPG_WNGMODIFCASCADE     = 20 ;
  CPG_ERRMSJAL            = 21 ;
  CPG_ERRMSMETH           = 22 ;
  CPG_ERRMETHETAB         = 23 ;
  CPG_ERRSELECTETAB       = 24 ;
  CPG_ESP_ERR             = 100 ;
  CPG_ESP_ERRCNDREMINCOHERENTE = CPG_ESP_ERR+1 ;

//========================================================
//===== CODE ERREUR DE LA GENERATION DES ENCA / DECA =====
//========================================================
Const
  CGE_PASERREUR           = -1 ;
  CGE_ERRTRAITEMENT       =  0 ;
  CGE_ERRTRANSAC          =  1 ;
  CGE_ANNULATION          =  2 ;
  CGE_NOEXODOSSIER        =  3 ;
  CGE_EXOCLOS             =  4 ;
  CGE_ERRVALIDEPIECE      =  5 ;
  CGE_ERRVALIDEESC        =  6 ;
  CGE_ERRNUMPIECE         =  7 ;
  CGE_ERRSAVEPIECE        =  8 ;
  CGE_ERRLETTPIECE        =  9 ;
  CGE_ERRUPDATEECHE       = 10 ;
  CGE_ERREDITION          = 11 ;
  CGE_ERREDITBOR          = 12 ;
  CGE_ERRVALIDEPIECEMS    = 13 ;
  
//================================
//===== GLOBALES ENCA / DECA =====
//================================
Var GErrEncaDeca : TErrorEncaDeca ;
    GStTri       : string ;

// Fonctions diverses pour Tob sur CPARAMGENER
Function  EstEncaissement ( TobScenario : TOB ) : Boolean ;
Function  EstDecaissement ( TobScenario : TOB ) : Boolean ;
Function  AvecEtat        ( TobScenario : TOB ) : Boolean ;
Function  AvecEscompte    ( TobScenario : TOB ) : Boolean ;
Procedure CompleteScenario( TobScenario : TOB ) ;
Procedure AjouteChampsSuppScenario ( TobScenario : TOB ) ;

// Fonctions de traitement des EncaDecas:
Function  VerifieCoherenceSelection ( TobOrigine, TobScenario : TOB ) : Boolean ;
Function  ExecuteEncaDeca   ( vTobOrigine, vTobScenario : TOB ; vTobRefPieces : Tob = nil ) : integer ;
Procedure GenereEncaDeca    ( TobOrigine, TobScenario, vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais, vTobRefPieces : TOB ) ;
Function  SauveEncaDeca     ( vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais, vTobScenario : TOB ) : Boolean ;

// Gestion des ruptures
Function  RuptureEtab       ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ; var vEtab : string ) : Boolean ;
Function  RuptureEncaDeca   ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ;
                              var vGeneral : String ; var vAuxiliaire : String ; var vDateEche : TDateTime ; var vEtab : string ) : Boolean ;
Function  RuptureEche       ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ;
                              var vGeneral : String ; var vAuxiliaire : String ; var vEtab : string ) : Boolean ;
Function  RuptureModePaie   ( vTobLigne, vTobScenario : TOB ; var vModePaie : String ; var vBoPremier : Boolean ) : Boolean ;

Function  UpdateNumeroPiece ( vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais, vTobScenario : TOB ) : Boolean ;
Function  CorrigeLigneAZero ( vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais : TOB ) : Boolean ;
Function  RuptureDossier    ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ; var vDossier : string ) : Boolean ;
Procedure MajCumulsEncaDeca( vTobLigne : TOB ; var vDebit, vCredit : Double ) ;

// Fonctions de génération de lignes
procedure GenereEcriture    ( vTobScenario, vTobOrigine : TOB ; vInDebut, vInFin : Integer ; vTobGene : TOB ; vBoContrepartie : Boolean = false ) ;
Procedure GenereAnalytique  ( vTobOrigine , vTobScenario : TOB ; vInDebut, vInFin : Integer ; vPieceCpt : TPieceCompta ; vNumLigne : Integer ) ;

// Fonctions de gestion de l'escompte
Procedure GenereEscompte    ( vTobLigne, vTobEscompte, vTobScenario : TOB ; vInfoEcr : TInfoEcriture ) ;
Function  DetermineTauxEsc  ( vTobLigne, vTobScenario : TOB ; vInfoEcr: TInfoEcriture ; vDev : RDevise ) : Double ;

// Préparation au traitement
Function  CoherenceAnalytique ( vTobOrigine, vTobScenario : TOB ; vInDebut, vInFin : Integer ; vInfoEcr : TInfoEcriture ) : Boolean ;
Function  GetLeMontant        ( Montant, Couv, Sais : Double ) : Double ;

// Gestion du E_NUMTRAITECHQ
Procedure MajNumTraiteChq          ( vTobOrigine, vTobScenario : TOB ) ;
Function  UpdateEcheancesOrigines  ( vTobOrigine, vTobScenario : TOB ) : Boolean ;
Function  AvecNumTraiteChq         ( vTobScenario : TOB ) : Boolean ;

// Lettrage
Function  GetRapproTob    ( vTobLigne : TOB ; vDev : RDEVISE ) : TL_Rappro ;
Procedure LettrerEncaDeca ( vTobOrig : TOB ; vPieceCpt : TPieceCompta ; vDev : RDevise ; vEtab : string = '' ) ;

// Traitements de génération des états
Function  RuptureEtats      ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ;
                                var vAuxiliaire : String ; var vDateEche : TDateTime ; var vDossier, vEtab : string ) : Boolean ;
Procedure AjouteInfosEtat     ( vTLEche : TList ; vTobScenario : TOB ; vInfoEcr : TInfoEcriture ) ;

function  GenereEtats         ( vTLEche : TList ; vTobScenario : TOB ) : boolean ;

// Procédures spécifiques aux documents
Procedure GenereDOCFACT       ( vTobLigne, vTobScenario : TOB ; vIndexDR, vIndexDF : Integer ) ;
Procedure GenereDOCREGL       ( vTobLigne, vTobScenario : TOB ; vIndexDoc : Integer ; vMontant, vMontantDev : Double ) ;
Procedure GenereRequeteDoc    ( vTobLigne, vTobScenario : TOB ; vTLReq : TList ; vIndexDoc : Integer ) ;
Function  GenereRibPourGeneral( vStGeneral : String ) : String ;

// Traitement de l'émission de bordereaux
Function  ExecuteEmissionBOR  ( vTLEche : TList ; vStModeleBOR : string ; vBoPrintDialog, vBoApercu : Boolean; Criteres : string = '' ) : Boolean ;

// Traitement à génération des exports
Function  GenereExports ( vTLEche : TList ; vTobScenario : TOB ; vDev : RDevise ) : Boolean ;
//Function  ExecuteExportCFONB ( vTobScenario : TOB ) : Boolean ;
//function  DetermineFichierExport ( var vFichier : String ) : Boolean ;

// Fonctions diverses sur tob ecriture
Function  EncodeTOBEcriture       ( vTobEcr : TOB ) : String ;

// Fonctions diverses pour multi-critère
Procedure CategorieVersModePaiement( vStCat : String ; vComboMP : THMultiValComboBox ) ;
Function  ConditionTypeEncadeca ( LeFlux : String ) : String ; //XMG 05/04/04

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ULibEncaDecaESP,
  CFONB,            // pour GetNumCFONB, ExportCFONB
  uTobDebug ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 16/12/2003
Modifié le ... :   /  /
Description .. : Retourne un TTypeEncaDeca renseigné en fonction du
Suite ........ : code de l'opération passé en paramètre.
Suite ........ : Le TypeEncaDeca permet de spécifier les particularités de
Suite ........ : charque type d'opération effectuer en génération
Suite ........ : d'enca/deca.
Suite ........ : Les différents code d'opération sont listé dans la tablette
Suite ........ : CPTYPEENCADECA
Suite ........ : Pour l'instant, toutes les données sont renseignées en dur
Suite ........ : dans la fonction...
Mots clefs ... : TYPEENCADECA;OPERATION
*****************************************************************}
Function ChargeTypeEncaDeca ( vStCode : String ) : TTypeEncaDeca ;
const NatureEtatESP = 'CME' ; //Tous les EncaDeca Espagnoles utilissent le même code d'état
const NatureDocESP = 'CML' ; //Tous les EncaDeca Espagnoles utilissent le même code De Document //XVI 24/02/2005
begin
  With Result do
    begin
    // Code
    Code              := vStCode ;

    // Valeur par défaut
    AccesCompta       := taPossible ;
    AccesEscompte     := taJamais ;
    AccesEdition      := taPossible ;
    AccesCFONB        := taJamais ;
    AccesBordereau    := taJamais ;
    AccesJournaux     := 'TTJALBANQUEOD' ;
    AccesMethode      := 'AUX;ECH;GLO;ECT;DET' ;
    AccesGeneraux     := 'TZGBQCAISS' ;
    AccesCondRemise   := False ;
    NatureEtat        := '' ;
    NatureDoc         := '' ;

    // Code "TOUS" utilisé uniquement dans le mul !!!
    if Code = 'TOUS' then
              begin
              AccesCompta       := taPossible ;
              AccesEscompte     := taPossible ;
              AccesEdition      := taPossible ;
              AccesCFONB        := taPossible ;
              AccesBordereau    := taPossible ;
              AccesCondRemise   := False ;
              end
    // Relevés de comptes
    else if Code = 'RCP' then
              begin
              AccesEdition      := taObligatoire ;
              NatureEtat        := 'CRC' ;
              AccesMethode      := 'AUX' ;
              end
    // Relevés de facture
    else if Code = 'RCP' then
              begin
              AccesEdition      := taObligatoire ;
              NatureEtat        := 'CRF' ;
              AccesMethode      := 'AUX' ;
              end
    // Cartes bancaires
    else if Code = 'CB' then
              begin
              NatureEtat        := 'CCB' ;
              NatureDoc         := 'CCB' ;
              AccesBordereau    := taPossible ;
              end
    // Chèques
    else if Code = 'CHQ' then
              begin
              NatureEtat        := 'CCH' ;
              NatureDoc         := 'LCH' ;
              AccesEscompte     := taPossible ;
              AccesBordereau    := taPossible ;
              end
    // Espèces
    else if Code = 'ESP' then
              begin
              NatureEtat        := 'CES' ;
              AccesBordereau    := taPossible ;
              end
    // Traite LCR BOR
    else if Code = 'LCR' then
              begin
              NatureEtat        := 'CTR' ;
              NatureDoc         := 'LTR' ;
              AccesEscompte     := taJamais ;
              AccesCFONB        := taPossible ;
              AccesBordereau    := taPossible ;
              end
    // Prélèvements
    else if Code = 'PRE' then
              begin
              NatureEtat        := 'CPR' ;
              NatureDoc         := 'LPR' ;
              AccesCFONB        := taPossible ;
              AccesBordereau    := taPossible ;
              end
    // Titre de paiement
    else if Code = 'TEP' then
              begin
              NatureEtat        := 'CTP' ;
              AccesBordereau    := taPossible ;
              end
    // Titres interbancaires de paiement
    else if Code = 'TIP' then
              begin
              NatureEtat        := 'CTB' ;
              AccesBordereau    := taPossible ;
              end
    // Transferts internationnaux
    else if Code = 'TRI' then
              begin
              NatureEtat        := 'CTI' ;
              NatureDoc         := 'LVI' ;
              AccesEscompte     := taPossible ;
              AccesCFONB        := taPossible ;
              AccesBordereau    := taPossible ;
              end
    // Virements commerciaux
    else if Code = 'VIC' then
              begin
              NatureEtat        := 'CVC' ;
              AccesBordereau    := taPossible ;
              end
    // Virements
    else if Code = 'VIR' then
              begin
              NatureEtat        := 'CVI' ;
              NatureDoc         := 'LVI' ;
              AccesEscompte     := taPossible ;
              AccesCFONB        := taPossible ;
              AccesBordereau    := taPossible ;
              end
    // Virement de trésorerie
    else if Code = 'VIT' then
              begin
              NatureEtat        := 'CVT' ;
              AccesBordereau    := taPossible ;
              end ;

    if VH^.PaysLocalisation=CodeISOES then //Spécifique Espagnol....
       Begin
       //Commun à Tous les opèrations
       NatureEtat        := NatureEtatESP ;
       NatureDoc         := NatureDocESP ;
       // Remise en banque
       if (Code = 'REC') or (Code='REE') then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taPossible ;
            AccesCFONB        := taPossible ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGCOLLECTIF' ;
            AccesCondRemise   := TRUE ;
            end
       // Rejet en banque
       else if (Code = 'RJC') or (Code='RJE') then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGCOLLECTIF' ;
            AccesCondRemise   := TRUE ;
            end
       // Confirmation Remise Encaissement
       else if Code = 'COC' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALBANQUE' ;
            AccesGeneraux     := 'TZGBQCAISS' ;
            end
       // Confirmation Remise Escompte
       else if Code='COE' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGENERAL' ;
            end
       // Déclaration des Impayés
       else if Code = 'IMC' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGCOLLECTIF' ;
            AccesCondRemise   := TRUE ;
            end
       // Déclaration des Impayés escompte
       else if Code='IME' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGCOLLECTIF' ;
            AccesCondRemise   := TRUE ;
            end
       // Paiement des impayés (?)
       else if Code='PIM' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALBANQUE' ;
            AccesGeneraux     := 'TZGBQCAISS' ;
            end
       // Re-negotiation (re-entrée en portefeuille
       else if Code='RNG' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGCOLLECTIF' ;
            end
       // Mise en portefeuille
       else if Code = 'MPF' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGCOLLECTIF' ;
            end
       // Fusion d'écheances
       else if Code = 'FUS' then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taJamais ;
            AccesCFONB        := taJamais ;
            AccesBordereau    := taPossible ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALOD' ;
            AccesGeneraux     := 'TZGCOLLECTIF' ;
            end
       // Décaissement (Fournisseurs) ou Encaissement (clients)
       else if (Code = 'DEC') or (Code='ENC') then
            begin
            AccesCompta       := taObligatoire ;
            AccesEscompte     := taJamais ;
            AccesEdition      := taPossible ;
            //XMG 20/04/04 début
            if Code = 'DEC' then
               AccesCFONB     := taPossible
              else
                AccesCFONB     := taJamais ;
            //XMG 20/04/04 fin
            AccesBordereau    := taJamais ;
            AccesMethode      := 'AUX;ECT;DET' ;
            AccesJournaux     := 'TTJALBANQUE' ;
            AccesGeneraux     := 'TZGBQCAISS' ;
            end
       End ; //XVI 24/02/2005

    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 16/12/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Retourne le chemin stocké dans les paramètres sociétés,
Suite ........ : dans lequel seront spooler les éditions en .pdf.
Suite ........ :
Suite ........ : Le paramSoc utilisé depend du type d'opération souhaitée
Mots clefs ... :
*****************************************************************}
Function CheminSpoolerPourEncaDeca ( vType : TTypeEncaDeca ) : String ;
begin

  // Lettre-chèque
  if vType.Code = 'CHQ'
    then Result := GetParamSocSecur('SO_CPCHEMINCHEQUE', '')
  // Lettre-BOR
  else if vType.Code = 'LCR'
    then Result := GetParamSocSecur('SO_CPCHEMINBOR', '')
  // Lettre-Traite
  else if vType.Code = 'LCR'
    then Result := GetParamSocSecur('SO_CPCHEMINTRAITE', '')
  // Lettre-Virement
  else if vType.Code = 'VIR'
    then Result := GetParamSocSecur('SO_CPCHEMINVIREMENT', '')
  // Lettre-Virement internationnaux
  else if vType.Code = 'TRI'
    then Result := GetParamSocSecur('SO_CHEMINVIRIN', '')
  // Lettre-Prélèvement
  else if vType.Code = 'PRE'
    then Result := GetParamSocSecur('SO_CPCHEMINPRELEVEMENT', '')

  // Sinon...
  else Result := '' ;

end ;

Function TriPourGroupeEncaDeca( vStCode : String ; avecLot : Boolean ) : String ;
Begin

  if vStCode='ECT'
    // Par Tiers/ échéances
    then Result := 'E_DATEECHEANCE;E_AUXILIAIRE;E_GENERAL;E_NUMEROPIECE'
    // les autres cas
    else Result := 'E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' ;

  // Si Traitement pas lot, ajout du tri par lot
  if avecLot then
    Result := 'E_N0MLOT;' + Result ;

End ;

Function  EstEncaissement ( TobScenario : TOB ) : Boolean ;
begin
  Result := TobScenario.GetValue('CPG_FLUXENCADECA') = 'ENC' ;
end ;

Function  EstDecaissement ( TobScenario : TOB ) : Boolean ;
begin
  Result := TobScenario.GetValue('CPG_FLUXENCADECA') = 'DEC' ;
end ;

Function  AvecEtat ( TobScenario : TOB ) : Boolean ;
begin
  Result := TobScenario.GetValue('CPG_MODELEENCADECA') <> '' ;
end ;

Function  AvecEscompte ( TobScenario : TOB ) : Boolean ;
begin
  Result := TobScenario.GetValue('CPG_ESCMETHODE') <> 'RIE' ;
end ;


// =============================================================================
// =============================================================================
// ======                 TRAITEMENTS DES ENCA / DECA                   ========
// =============================================================================
// =============================================================================


function AutoriseTestContrepartie( vTobScenario : TOB ) : Boolean ;
begin

  result := False ;
  if not Assigned( vTobScenario ) then Exit ;

  result := ( vTobScenario.GetString('CPG_MULTISOC')<>'X' )
            and ( vTobScenario.GetString('CPG_MULTIETAB')<>'X' )
            and ( vTobScenario.GetString('CPG_EDITE')='X' )
            and ( vTobScenario.GetString('CPG_GROUPEENCADECA')<>'GLO' ) ;

end ;

function VireReglementNegatifs( vTobScenario, vTobGene : Tob ) : Boolean ;
var lTobEcr : Tob ;
    i, j    : integer ;
    lPiece  : TPieceCompta ;
    lFlux   : string ;
    lBoSupprEnCours : Boolean ;
begin

  result := False ;

  try
    lFlux           := vTobScenario.GetString('CPG_FLUXENCADECA') ;

    for i := 0 to vTobGene.Detail.count - 1 do
      begin

      lPiece := TPieceCompta( vTobGene.Detail[ i ] ) ;
      lBoSupprEnCours := False ;

      for j := (lPiece.Detail.count - 1) downto 0 do
        begin
        lTobEcr := lPiece.Detail[ j ] ;

        if lTobEcr.GetString('E_GENERAL') = vTobScenario.GetString('CPG_GENERAL') then
          lBoSupprEnCours := (lFlux <> lTobEcr.GetString('E_ENCAISSEMENT') ) ;

        if lBoSupprEnCours then
          FreeAndNil( lTobEcr ) ;

        end ;

      end ;

      result := True ;

  Except
    end ;

end ;


function SensReglementOk( vTobScenario, vTobGene : Tob ; var vInCodeErr : integer ) : Boolean ;
var lTobEcr : Tob ;
    i, j    : integer ;
    lPiece  : TPieceCompta ;
    lFlux   : string ;
begin

  result     := True ;
  vInCodeErr := CGE_PASERREUR ;
  if not AutoriseTestContrepartie( vTobscenario ) then Exit ;

  lFlux := vTobScenario.GetString('CPG_FLUXENCADECA') ;

  // Parcours des pièces générées
  for i := 0 to vTobGene.Detail.count - 1 do
    begin
    lPiece := TPieceCompta( vTobGene.Detail[ i ] ) ;
    // Parcours des lignes d'écritures
    for j := 0 to lPiece.Detail.count - 1 do
      begin
      lTobEcr := lPiece.Detail[ j ] ;
      // Est-on sur le compte de contrepartie
      if lTobEcr.GetString('E_GENERAL') = vTobScenario.GetString('CPG_GENERAL') then
        begin
          if (lFlux <> lTobEcr.GetString('E_ENCAISSEMENT') ) then
          begin
          Case PgiAskCancel('Certains documents générés présentent des totaux négatifs. Voulez-vous les inclure dans le traitement ?', 'Vérification des documents' ) of
            mrYes :     begin
                        result := true ;
                        end ;
            mrNo :      begin
                        result := VireReglementNegatifs( vTobScenario, vTobGene ) ;
                        if not result then
                          vInCodeErr := CGE_ERRTRAITEMENT ;
                        end ;
            mrCancel :  begin
                        result := False ;
                        vInCodeErr := CGE_ANNULATION ;
                        end ;
            end ;
          Exit ;
          end ;
        end ;
      end ;

    end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. : retourne un TPieceCompta paramétré pour la génération
Suite ........ : d'un réglement sur le dossier passé en paramètre
Mots clefs ... :
*****************************************************************}
function CreerPieceEncaDeca ( vTobScenario, vTobTemoin, vTobGene : Tob ; vDossier : string = '' ) : TPiececompta ;
var lInfoEcr        : TInfoEcriture ;
begin
  // instanciation
  lInfoEcr := TInfoEcriture.Create( vDossier ) ;
  result   := TPieceCompta.CreerPiece( lInfoEcr ) ;
  if Assigned( vTobGene ) then
    result.ChangeParent( vTobGene, -1 ) ;

  // Paramétrage
  result.InitSaisie ;
  if not Assigned( GErrEncaDeca ) then
    GErrEncaDeca  := TErrorEncaDeca.CreerMsgErreur ;
  result.OnError := GErrEncaDeca.OnError ;
  result.SetMultiEcheOff ; // pas de multi-échéance
  if (vTobScenario.GetValue('CPG_METHODEANA')='ATT') or ( not result.Contexte.local ) then
    result.SetVentilSurAttente ; // analytique sur section d'attente uniquement
  // pas d'attribution auto du rib :
  result.Contexte.AttribRIBOff := True ;

  // Entête
  result.PutEntete('E_QUALIFPIECE',   'S' ) ;
  result.PutEntete('E_ECRANOUVEAU',   'N' ) ;
  if not result.Contexte.local then
    begin
    result.PutEntete('E_NATUREPIECE',   'OD' ) ;
    result.PutEntete('E_JOURNAL',       vTobScenario.GetValue('CPG_MULTISOCJAL') ) ;
    end
  else
    begin
    result.PutEntete('E_NATUREPIECE',   vTobScenario.GetValue('NATUREPIECE') ) ;
    result.PutEntete('E_JOURNAL',       vTobScenario.GetValue('CPG_JOURNAL') ) ;
    end ;

  result.PutEntete('E_DATECOMPTABLE', vTobScenario.GetValue('DATECOMPTABLE') ) ;
  result.PutEntete('E_DEVISE',        vTobTemoin.GetValue('E_DEVISE') ) ;

  if ( vTobScenario.GetString('CPG_MULTIETAB') = 'X' ) then
    begin
    if ( vTobScenario.GetValue('CPG_SELECTETAB') = 'FAC' ) and result.Contexte.local
      then result.PutEntete( 'E_ETABLISSEMENT', vTobScenario.GetValue('ETABLISSEMENTPAYEUR') )
      else result.PutEntete( 'E_ETABLISSEMENT', result.Contexte.EtablisDefaut )
    end
    else result.PutEntete('E_ETABLISSEMENT', vTobTemoin.GetValue('E_ETABLISSEMENT') ) ;

  result.AttribNumeroTemp ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. : Retourne un nouveau TPieceCompta paramétré poru la
Suite ........ : génération d'une pièce d'escompte
Mots clefs ... :
*****************************************************************}
function CreerPieceEscompte ( vTobScenario, vTobTemoin, vTobEscompte : Tob ; vInfoEcr : TInfoEcriture ) : TPiececompta ;
begin
  // instanciation
  result := TPieceCompta.CreerPiece( vInfoEcr ) ;
  result.ChangeParent( vTobEscompte, -1 ) ;

  // Paramétrage
  result.InitSaisie ;
  if not Assigned( GErrEncaDeca ) then
    GErrEncaDeca  := TErrorEncaDeca.CreerMsgErreur ;
  result.OnError := GErrEncaDeca.OnError ;
  result.SetMultiEcheOff ;
  // pas d'attribution auto du rib :
  result.Contexte.AttribRIBOff := True ;

  // Entête
  result.PutEntete('E_QUALIFPIECE',   'S' ) ;
  result.PutEntete('E_ECRANOUVEAU',   'N' ) ;
  result.PutEntete('E_NATUREPIECE',   'OD' ) ;
  result.PutEntete('E_DATECOMPTABLE', vTobScenario.GetValue('DATECOMPTABLE') ) ;
  result.PutEntete('E_JOURNAL',       vTobScenario.GetValue('ESCJOURNAL') ) ;
  result.PutEntete('E_DEVISE',        vTobTemoin.GetValue('E_DEVISE') ) ;

  if ( vTobScenario.GetString('CPG_MULTIETAB') = 'X' ) then
    begin
    if ( vTobScenario.GetValue('CPG_SELECTETAB') = 'FAC' ) and result.Contexte.local
      then result.PutEntete( 'E_ETABLISSEMENT', vTobScenario.GetValue('ETABLISSEMENTPAYEUR') )
      else result.PutEntete( 'E_ETABLISSEMENT', result.Contexte.EtablisDefaut )
    end
  else result.PutEntete('E_ETABLISSEMENT', vTobTemoin.GetValue('E_ETABLISSEMENT') ) ;

  result.AttribNumeroTemp ; // FQ 21090 SBO 07/08/2007 : Pb génération de l'escompte

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. : retourne le TPieceCompta pointant sur le dossier passé en
Suite ........ : paramètre
Mots clefs ... :
*****************************************************************}
function GetPieceDossier( vTobGene : TOB ; vDossier : string ) : TPieceCompta ;
var lPC : TPieceCompta ;
    i   : integer ;
begin

  result := nil ;
  if vTobGene.Detail.Count = 0 then Exit ;

  // Pièce sur dossier local
  result := TPieceCompta( vTobGene.Detail[0] ) ;
  if (not EstMultiSoc) or (result.Dossier = vDossier) or (vDossier='') then Exit ;

  // recherche sur les autres dossiers
  result := nil ;
  for i := 1 to vTobGene.Detail.Count - 1 do
    begin
    lPC := TPieceCompta( vTobGene.Detail[i] ) ;
    if lPC.Dossier = vDossier then
      begin
      result := lPC ;
      exit ;
      end ;
    end ;

end ;

procedure ChargeInfosSoc( vTobScenario : Tob ) ;
var lTobNomSoc   : Tob ;
    lTobSoc      : Tob ;
    i            : integer ;
    lStLibelle   : string ;
begin
  // chargement des liaisons inter-sociétés
  vTobScenario.LoadDetailDB('CLIENSSOC', '', '', nil, False ) ;

  // Ajout libelles des sociétés
  lTobNomSoc := RecupInfosSocietes ( 'SO_SOCIETE;SO_LIBELLE' ) ;
  if Assigned(lTobNomSoc) then
    begin
    for i:=0 to vTobScenario.Detail.Count - 1 do
      begin
      lStLibelle := vTobScenario.Detail[i].GetString('CLS_SOCIETE') ;
      lTobSoc := lTobNomSoc.FindFirst( ['SO_SOCIETE'], [lStLibelle], False )  ;
      if Assigned(lTobSoc) then
        lStLibelle := lTobSoc.GetString('SO_LIBELLE') ;
      vTobScenario.Detail[i].AddChampSupValeur('LIBELLE', lStLibelle ) ;
      end ;
    lTobNomSoc.Free ;
    end ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. : Retourne le compte de liaison paramétré pour générer les
Suite ........ : règlement vers le dossier vDossier
Mots clefs ... :
*****************************************************************}
function GetCompteLiaison( vTobScenario : TOB ; vDossier : string ) : string ;
var   lTobLiensSoc : Tob ;
begin

  result := '' ;

  // Chargement des liaisons inter-sociétés
  if vTobScenario.Detail.Count = 0 then
    ChargeInfosSoc( vTobScenario ) ;

  lTobLiensSoc := vTobScenario.FindFirst( [ 'CLS_DOSSIER' ], [ vDossier ], False ) ;
  if Assigned( lTobLiensSoc ) then
    result := lTobLiensSoc.GetString('CLS_GENERAL') ;

end ;



function GetLiaisonSoc( vTobScenario : TOB ; vDossier : string ) : string ;
var   lTobLiensSoc : Tob ;
begin

  result := vDossier ;

  // Chargement des liaisons inter-sociétés
  if vTobScenario.Detail.Count = 0 then
    ChargeInfosSoc( vTobScenario ) ;

  lTobLiensSoc := vTobScenario.FindFirst( ['CLS_DOSSIER'], [vDossier], False ) ;
  if Assigned( lTobLiensSoc ) then
    begin
    result := lTobLiensSoc.GetString('LIBELLE') ;
    if Trim( result ) = '' then
      result := lTobLiensSoc.GetString('CLS_SOCIETE') ;
    end ;


end ;


procedure SetLibelleMultiSoc( vTobScenario, vTobOrig : Tob ; vPieceCpt : TPieceCompta ; vNumLigne : Integer ; vBoContrepartie : Boolean ; vDossierDest : string = '' ) ;
var lStSoc : string ;
    lStLib : string ;
    lStAdj : string ;
    lStTyp : string ;
begin

  if vBoContrepartie
    then lStTyp := 'Reglt'
    else lStTyp := 'Eché.' ;

  // détermination de la société
  if vPieceCpt.Contexte.local then
    begin
    if vDossierDest = '' then Exit ;
    lStSoc := GetLiaisonSoc(vTobScenario, vDossierDest ) ;
    lStAdj := ' pour ' ;
    end
  else
    begin
    lStSoc := GetParamSocSecur('SO_LIBELLE', '') ;
    lStAdj := ' par ' ;
    end ;

  // Affectation libellé suivant méthode de génération
  // 1. par Société
  if vTobScenario.GetValue('CPG_GROUPEENCADECA') = 'GLO' then
    lStLib := lStTyp + ' société ' + lStSoc
  // 2. par Echéance :
  else if vTobScenario.GetValue('CPG_GROUPEENCADECA') = 'ECT' then
    lStLib := lStTyp + ' du ' + DateToStr( vTobOrig.GetDateTime('E_DATEECHEANCE')  ) + lStAdj + lStSoc
  // 3. par tiers
  else
    begin
    if vPieceCpt.Info.LoadAux( vTobOrig.GetString('E_AUXILIAIRE') )
      then lStLib := ' ' + copy( vPieceCpt.Info.Aux_GetValue('T_LIBELLE'), 1, 15 )
    else if vPieceCpt.Info.LoadCompte( vTobOrig.GetString('E_GENERAL') )
      then lStLib := ' ' + copy( vPieceCpt.Info.Compte_GetValue('G_LIBELLE'), 1, 15 )
    else lStLib := '' ;

    lStLib := lStTyp + lStLib + lStAdj + lStSoc ;
    end ;

  vPieceCpt.PutValue( vNumLigne, 'E_LIBELLE', Copy( lStLib, 1, 35) ) ;

end ;


procedure InitInfosEtat ( vTobEcr : Tob  ) ;
begin
  // infos Tiers
  vTobEcr.AddChampSup( 'T_JURIDIQUE',   False ) ;
  vTobEcr.AddChampSup( 'T_LIBELLE',     False ) ;
  vTobEcr.AddChampSup( 'T_CODEPOSTAL',  False ) ;
  vTobEcr.AddChampSup( 'T_VILLE',       False ) ;
  vTobEcr.AddChampSup( 'T_ADRESSE1',    False ) ;
  vTobEcr.AddChampSup( 'T_ADRESSE2',    False ) ;
  vTobEcr.AddChampSup( 'T_ADRESSE3',    False ) ;
  vTobEcr.AddChampSup( 'T_LIBELLE',     False ) ;
  vTobEcr.AddChampSup( 'T_PRENOM',      False ) ;
  // infos General
  vTobEcr.AddChampSup( 'G_LIBELLE',     False ) ;
  vTobEcr.AddChampSup( 'G_CODEPOSTAL',  False ) ;
  vTobEcr.AddChampSup( 'G_VILLE',       False ) ;
  vTobEcr.AddChampSup( 'G_ADRESSE1',    False ) ;
  vTobEcr.AddChampSup( 'G_ADRESSE2',    False ) ;
  vTobEcr.AddChampSup( 'G_ADRESSE3',    False ) ;
  vTobEcr.AddChampSup( 'G_PAYS',        False ) ;
  vTobEcr.AddChampSup( 'G_TELEPHONE',   False ) ;
  // Info contact tiers
  vTobEcr.AddChampSup( 'C_CIVILITE',    False ) ;
  vTobEcr.AddChampSup( 'C_NOM',         False ) ;
  vTobEcr.AddChampSup( 'C_PRENOM',      False ) ;
  vTobEcr.AddChampSup( 'C_SERVICE',     False ) ;
  vTobEcr.AddChampSup( 'C_FONCTION',    False ) ;
  vTobEcr.AddChampSup( 'C_TELEPHONE',   False ) ;
  // infos banques
  vTobEcr.AddChampSup( 'BQ_CODE',          False ) ;
  vTobEcr.AddChampSup( 'BQ_LIBELLE',       False ) ;
  vTobEcr.AddChampSup( 'BQ_ADRESSE1',      False ) ;
  vTobEcr.AddChampSup( 'BQ_ADRESSE2',      False ) ;
  vTobEcr.AddChampSup( 'BQ_ADRESSE3',      False ) ;
  vTobEcr.AddChampSup( 'BQ_VILLE',         False ) ;
  vTobEcr.AddChampSup( 'BQ_GENERAL',       False ) ;
  vTobEcr.AddChampSup( 'BQ_CODEPOSTAL',    False ) ;
  vTobEcr.AddChampSup( 'BQ_PAYS',          False ) ;
  vTobEcr.AddChampSup( 'BQ_ETABBQ',        False ) ;
  vTobEcr.AddChampSup( 'BQ_GUICHET',       False ) ;
  vTobEcr.AddChampSup( 'BQ_NUMEROCOMPTE',  False ) ;
  vTobEcr.AddChampSup( 'BQ_CLERIB',        False ) ;
  vTobEcr.AddChampSup( 'BQ_CODEIBAN',      False ) ;
  vTobEcr.AddChampSup( 'BQ_TELEPHONE',     False ) ;
  vTobEcr.AddChampSup( 'BQ_FAX',           False ) ;
  vTobEcr.AddChampSup( 'BQ_TELEX',         False ) ;
  vTobEcr.AddChampSup( 'BQ_DOMICILIATION', False ) ;
  vTobEcr.AddChampSup( 'BQ_CODEBIC',       False ) ;
end ;

Procedure AjouteInfosSupp( vTobEcr, vTobScenario, vTobOrigine : TOB ; vBoContrepartie : Boolean ) ;
begin

  vTobEcr.AddChampSup( 'ESTECHEORIG',   False ) ;
  vTobEcr.AddChampSup( 'CLEECHEORIG',   False ) ;
  vTobEcr.AddChampSup( 'NUMENCADECA',   False ) ;
  vTobEcr.AddChampSup( 'MONTANTESC',    False ) ;
  vTobEcr.AddChampSup( 'MONTANTESCDEV', False ) ;
  vTobEcr.AddChampSup( 'TAUXESC',       False ) ;
{
  vTobEcr.AddChampSup( 'T_JURIDIQUE',   False ) ;
  vTobEcr.AddChampSup( 'T_LIBELLE',     False ) ;
  vTobEcr.AddChampSup( 'T_CODEPOSTAL',  False ) ;
  vTobEcr.AddChampSup( 'T_VILLE',       False ) ;
  vTobEcr.AddChampSup( 'T_ADRESSE1',    False ) ;
  vTobEcr.AddChampSup( 'T_ADRESSE2',    False ) ;
  vTobEcr.AddChampSup( 'T_ADRESSE3',    False ) ;
  vTobEcr.AddChampSup( 'T_LIBELLE',     False ) ;
  vTobEcr.AddChampSup( 'T_PRENOM',      False ) ;
  vTobEcr.AddChampSup( 'G_LIBELLE',     False ) ;
  vTobEcr.AddChampSup( 'G_CODEPOSTAL',  False ) ;
  vTobEcr.AddChampSup( 'G_VILLE',       False ) ;
  vTobEcr.AddChampSup( 'G_ADRESSE1',    False ) ;
  vTobEcr.AddChampSup( 'G_ADRESSE2',    False ) ;
  vTobEcr.AddChampSup( 'G_ADRESSE3',    False ) ;
  vTobEcr.AddChampSup( 'G_PAYS',        False ) ;
  vTobEcr.AddChampSup( 'G_TELEPHONE',   False ) ;
  vTobEcr.AddChampSup( 'C_CIVILITE',    False ) ;
  vTobEcr.AddChampSup( 'C_NOM',         False ) ;
  vTobEcr.AddChampSup( 'C_PRENOM',      False ) ;
  vTobEcr.AddChampSup( 'C_SERVICE',     False ) ;
  vTobEcr.AddChampSup( 'C_FONCTION',    False ) ;
  vTobEcr.AddChampSup( 'C_TELEPHONE',   False ) ;
}
  if vBoContrepartie then
    begin
    vTobEcr.PutValue('ESTECHEORIG', '-' ) ;
    vTobEcr.PutValue('CLEECHEORIG', '' ) ;
    end
  else
    begin
    vTobEcr.PutValue('ESTECHEORIG', 'X' ) ;
    vTobEcr.PutValue('CLEECHEORIG', EncodeTOBEcriture( vTobOrigine ) ) ;

    if ( vTobScenario.GetValue('CPG_CFONBEXPORT')='X' )
       or (vTobScenario.GetValue('CPG_BORDEREAUEXP')='X') then
      vTobEcr.PutValue('E_CFONBOK', '#');

    if vTobOrigine.GetNumChamp('MONTANTESC') > 0
       then vTobEcr.PutValue('MONTANTESC', vTobOrigine.GetValue('MONTANTESC') ) ;
    if vTobOrigine.GetNumChamp('MONTANTESCDEV') > 0
       then vTobEcr.PutValue('MONTANTESC', vTobOrigine.GetValue('MONTANTESCDEV') ) ;
    if vTobOrigine.GetNumChamp('TAUXESC') > 0
       then vTobEcr.PutValue('TAUXESC',    vTobOrigine.GetValue('TAUXESC') ) ;
    end ;

  // Gestion multisoc
  vTobEcr.AddChampSupValeur( 'DOSSIERDEST',   '' ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 05/01/2004
Modifié le ... : 08/02/2006
Description .. : Retourne le montant restant à facturer
Suite ........ : Respris de mulSmpUtil
Mots clefs ... :
*****************************************************************}
Function  GetLeMontant ( Montant, Couv, Sais : Double ) : Double ;
begin
  Result := 0 ;
  if Arrondi(Montant,4) = 0 then Exit ;
  if Arrondi(Sais,4) = 0
    then Result := Arrondi(Montant-Couv, 6)
    else Result := Sais ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 08/02/2006
Modifié le ... :   /  /
Description .. : Ajoute le montant de l'écriture passée en paramètre aux
Suite ........ : montants vDebit, vCredit.
Suite ........ : Le résultat est exprimé en solde (toujours 1 des 2 montants 
Suite ........ : est à 0)
Mots clefs ... : 
*****************************************************************}
Procedure MajCumulsEncaDeca( vTobLigne : TOB ; var vDebit, vCredit : Double ) ;
begin

  if vTobLigne = nil then Exit ;

  vDebit   := vDebit  + GetLeMontant( vTobLigne.GetValue('E_DEBITDEV'),  vTobLigne.GetValue('E_COUVERTUREDEV'),  vTobLigne.GetValue('E_SAISIMP') ) ;
  vCredit  := vCredit + GetLeMontant( vTobLigne.GetValue('E_CREDITDEV'), vTobLigne.GetValue('E_COUVERTUREDEV'),  vTobLigne.GetValue('E_SAISIMP') ) ;

  // calcul du solde si besoin
  if ((vCredit<>0) and (vDebit<>0)) then
    if Abs(vDebit)>Abs(vCredit) then
      begin
      vDebit      :=  vDebit - vCredit ;
      vCredit     :=  0 ;
      end
    else
      begin
      vCredit     := vCredit - vDebit ;
      vDebit      := 0 ;
      end ;

end ;

Procedure RecopieInfosAnalytique ( vTobOrigine , vTobGene : TOB ) ;
begin

  // Données principales
  vTobGene.PutValue('Y_AXE',              vTobOrigine.GetValue('Y_AXE') ) ;
  vTobGene.PutValue('Y_SECTION',          vTobOrigine.GetValue('Y_SECTION') ) ;

  // Infos Générale
  vTobGene.PutValue('Y_REFEXTERNE',       vTobOrigine.GetValue('Y_REFEXTERNE') ) ;
  vTobGene.PutValue('Y_DATEREFEXTERNE',   vTobOrigine.GetValue('Y_DATEREFEXTERNE') ) ;

  // QTES
  vTobGene.PutValue('Y_QUALIFQTE1',     vTobOrigine.GetValue('Y_QUALIFQTE1') ) ;
  vTobGene.PutValue('Y_QUALIFQTE2',     vTobOrigine.GetValue('Y_QUALIFQTE2') ) ;

  // Infos Complémentaires
  vTobGene.PutValue('Y_REFLIBRE',       vTobOrigine.GetValue('Y_REFLIBRE') ) ;
  vTobGene.PutValue('Y_AFFAIRE',        vTobOrigine.GetValue('Y_AFFAIRE') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE0',    vTobOrigine.GetValue('Y_LIBRETEXTE0') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE1',    vTobOrigine.GetValue('Y_LIBRETEXTE1') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE2',    vTobOrigine.GetValue('Y_LIBRETEXTE2') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE3',    vTobOrigine.GetValue('Y_LIBRETEXTE3') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE4',    vTobOrigine.GetValue('Y_LIBRETEXTE4') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE5',    vTobOrigine.GetValue('Y_LIBRETEXTE5') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE6',    vTobOrigine.GetValue('Y_LIBRETEXTE6') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE7',    vTobOrigine.GetValue('Y_LIBRETEXTE7') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE8',    vTobOrigine.GetValue('Y_LIBRETEXTE8') ) ;
  vTobGene.PutValue('Y_LIBRETEXTE9',    vTobOrigine.GetValue('Y_LIBRETEXTE9') ) ;
  vTobGene.PutValue('Y_TABLE0',         vTobOrigine.GetValue('Y_TABLE0') ) ;
  vTobGene.PutValue('Y_TABLE1',         vTobOrigine.GetValue('Y_TABLE1') ) ;
  vTobGene.PutValue('Y_TABLE2',         vTobOrigine.GetValue('Y_TABLE2') ) ;
  vTobGene.PutValue('Y_TABLE3',         vTobOrigine.GetValue('Y_TABLE3') ) ;
  vTobGene.PutValue('Y_LIBREDATE',      vTobOrigine.GetValue('Y_LIBREDATE') ) ;
  vTobGene.PutValue('Y_LIBREBOOL0',     vTobOrigine.GetValue('Y_LIBREBOOL0') ) ;
  vTobGene.PutValue('Y_LIBREBOOL1',     vTobOrigine.GetValue('Y_LIBREBOOL1') ) ;
  vTobGene.PutValue('Y_LIBREMONTANT0',  vTobOrigine.GetValue('Y_LIBREMONTANT0') ) ;
  vTobGene.PutValue('Y_LIBREMONTANT1',  vTobOrigine.GetValue('Y_LIBREMONTANT1') ) ;
  vTobGene.PutValue('Y_LIBREMONTANT2',  vTobOrigine.GetValue('Y_LIBREMONTANT2') ) ;
  vTobGene.PutValue('Y_LIBREMONTANT3',  vTobOrigine.GetValue('Y_LIBREMONTANT3') ) ;
  vTobGene.PutValue('Y_CONSO',          vTobOrigine.GetValue('Y_CONSO') ) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 08/01/2004
Modifié le ... :   /  /
Description .. : Affecte à la ligne d'écriture générée, les informations à
Suite ........ : reprendre de la ligne d'échéance d'origine.
Mots clefs ... :
*****************************************************************}
Procedure RecopieInfosEcheance      ( vTobOrigine : TOB ; vPiece : TPieceCompta ; vNumLigne : Integer ) ;
begin

  // Infos mode de paiement
//  if vPiece.GetValue( vNumLigne, 'E_ECHE')='X' then begin
  vPiece.PutValue( vNumLigne, 'E_MODEPAIE'        , vTobOrigine.GetString('E_MODEPAIE'));
  if NbJoursOk( vPiece.GetDateTime( vNumLigne, 'E_DATECOMPTABLE' ), vTobOrigine.GetDateTime('E_DATEECHEANCE') )
    then vPiece.PutValue( vNumLigne, 'E_DATEECHEANCE'    , vTobOrigine.GetDateTime('E_DATEECHEANCE'))
    else vPiece.PutValue( vNumLigne, 'E_DATEECHEANCE'    , vPiece.GetDateTime( vNumLigne, 'E_DATECOMPTABLE' )) ;
  vPiece.PutValue( vNumLigne, 'E_ORIGINEPAIEMENT' , vTobOrigine.GetString('E_ORIGINEPAIEMENT'));
// end;

  //Date de valeur
  vPiece.AffecteDateValeur( vNumLigne ) ;

  // Infos Générale
  vPiece.PutValue( vNumLigne, 'E_REFINTERNE',       vTobOrigine.GetValue('E_REFINTERNE') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBELLE',          vTobOrigine.GetValue('E_LIBELLE') ) ;
  vPiece.PutValue( vNumLigne, 'E_RIB',              vTobOrigine.GetValue('E_RIB') ) ;

  // TVA ???
  vPiece.PutValue( vNumLigne, 'E_TVAENCAISSEMENT',  vTobOrigine.GetValue('E_TVAENCAISSEMENT') ) ;
  vPiece.PutValue( vNumLigne, 'E_REGIMETVA',        vTobOrigine.GetValue('E_REGIMETVA') ) ;
  vPiece.PutValue( vNumLigne, 'E_TVA',              vTobOrigine.GetValue('E_TVA') ) ;
  vPiece.PutValue( vNumLigne, 'E_TPF',              vTobOrigine.GetValue('E_TPF') ) ;

  // Infos Complémentaires
  vPiece.PutValue( vNumLigne, 'E_REFEXTERNE',     vTobOrigine.GetValue('E_REFEXTERNE') ) ;
  vPiece.PutValue( vNumLigne, 'E_DATEREFEXTERNE', vTobOrigine.GetValue('E_DATEREFEXTERNE') ) ;
  vPiece.PutValue( vNumLigne, 'E_REFLIBRE',       vTobOrigine.GetValue('E_REFLIBRE') ) ;
  vPiece.PutValue( vNumLigne, 'E_AFFAIRE',        vTobOrigine.GetValue('E_AFFAIRE') ) ;
  vPiece.PutValue( vNumLigne, 'E_QTE1',           vTobOrigine.GetValue('E_QTE1') ) ;
  vPiece.PutValue( vNumLigne, 'E_QTE2',           vTobOrigine.GetValue('E_QTE2') ) ;
  vPiece.PutValue( vNumLigne, 'E_QUALIFQTE1',     vTobOrigine.GetValue('E_QUALIFQTE1') ) ;
  vPiece.PutValue( vNumLigne, 'E_QUALIFQTE2',     vTobOrigine.GetValue('E_QUALIFQTE2') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE0',    vTobOrigine.GetValue('E_LIBRETEXTE0') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE1',    vTobOrigine.GetValue('E_LIBRETEXTE1') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE2',    vTobOrigine.GetValue('E_LIBRETEXTE2') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE3',    vTobOrigine.GetValue('E_LIBRETEXTE3') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE4',    vTobOrigine.GetValue('E_LIBRETEXTE4') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE5',    vTobOrigine.GetValue('E_LIBRETEXTE5') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE6',    vTobOrigine.GetValue('E_LIBRETEXTE6') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE7',    vTobOrigine.GetValue('E_LIBRETEXTE7') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE8',    vTobOrigine.GetValue('E_LIBRETEXTE8') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBRETEXTE9',    vTobOrigine.GetValue('E_LIBRETEXTE9') ) ;
  vPiece.PutValue( vNumLigne, 'E_TABLE0',         vTobOrigine.GetValue('E_TABLE0') ) ;
  vPiece.PutValue( vNumLigne, 'E_TABLE1',         vTobOrigine.GetValue('E_TABLE1') ) ;
  vPiece.PutValue( vNumLigne, 'E_TABLE2',         vTobOrigine.GetValue('E_TABLE2') ) ;
  vPiece.PutValue( vNumLigne, 'E_TABLE3',         vTobOrigine.GetValue('E_TABLE3') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBREDATE',      vTobOrigine.GetValue('E_LIBREDATE') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBREBOOL0',     vTobOrigine.GetValue('E_LIBREBOOL0') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBREBOOL1',     vTobOrigine.GetValue('E_LIBREBOOL1') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBREMONTANT0',  vTobOrigine.GetValue('E_LIBREMONTANT0') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBREMONTANT1',  vTobOrigine.GetValue('E_LIBREMONTANT1') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBREMONTANT2',  vTobOrigine.GetValue('E_LIBREMONTANT2') ) ;
  vPiece.PutValue( vNumLigne, 'E_LIBREMONTANT3',  vTobOrigine.GetValue('E_LIBREMONTANT3') ) ;
  vPiece.PutValue( vNumLigne, 'E_CONSO',          vTobOrigine.GetValue('E_CONSO') ) ;

  {JP 04/06/07 : FQ 19502 : Sur un journal d'effets, on reprend E_BANQUEPREVI}
  vPiece.Info.LoadCompte(vTobOrigine.GetValue('E_GENERAL'));
  if vPiece.Info.Compte.GetString('G_EFFET') = 'X' then
    vPiece.PutValue(vNumLigne, 'E_BANQUEPREVI',  vTobOrigine.GetValue('E_BANQUEPREVI'));
end ;



procedure RenseigneEcriture ( vTobScenario, vTobOrigine : TOB ; vPieceCpt : TPieceCompta ; vNumLigne : Integer ; vBoContrepartie : Boolean ) ;
var lBoMemeDossier : Boolean ;
    lIdxChamp      : string ;
    lStCpt         : string ;
    lStAux         : string ;
begin

  // ---------------------------------------
  // --- Infos de l'échéance à récupérer ---
  // ---------------------------------------
  RecopieInfosEcheance ( vTobOrigine , vPieceCpt, vNumLigne ) ;
  if vBoContrepartie and vPieceCpt.Contexte.local and ( vTobScenario.GetString('CPG_GROUPEENCADECA') <> 'DET' )   then
    begin
    vPieceCPT.PutValue( vNumLigne, 'E_REFINTERNE', '' ) ;
    vPieceCPT.PutValue( vNumLigne, 'E_LIBELLE',    '' ) ;
    end ;

  // Recopie du E_NumTraiteChq si besoin
  if AvecNumTraiteChq( vTobScenario ) then
    vPieceCpt.PutValue( vNumLigne, 'E_NUMTRAITECHQ',   vTobOrigine.GetValue('E_NUMTRAITECHQ') ) ;

  // --------------------------------
  // -- GESTION DE L'ETABLISSEMENT --
  // --------------------------------
  if ( vTobOrigine.GetNumChamp('SYSDOSSIER') > 0 )
    then lBoMemeDossier := vTobOrigine.GetString('SYSDOSSIER') = vPieceCpt.Dossier
    else lBoMemeDossier := True ;
{  if ( vTobOrigine.GetNumChamp('SYSDOSSIER') > 0 ) and ( vTobOrigine.GetString('SYSDOSSIER') <> vPieceCpt.Dossier ) then
    begin
    // CAS MULTI-SOCIETE : On renseigne le compte de liaison sur l'établissement d'entête de la pièce
    // Rien à faire, affecté automatiquement par le TPieceCompta
    end
  else }
    if not vBoContrepartie and lBoMemeDossier and // (pour la contrepartie, toujours sur l'établissement d'entête)
       ( vTobOrigine.GetString('E_ETABLISSEMENT') <> vPieceCpt.GetEntete('E_ETABLISSEMENT') )
       and ( vTobScenario.GetString('CPG_MULTIETAB') = 'X' ) then
      begin
      // CAS MULTI-ETABLISSEMENT A GERER :
      if ( vTobScenario.GetValue('CPG_SELECTETAB') = 'FAC' ) and vPieceCpt.IsValidEtab( vTobOrigine.GetString('E_ETABLISSEMENT') )
        // 1. Affectation au détail, on récupère celui de la facture
        then vPieceCpt.PutValue( vNumLigne, 'E_ETABLISSEMENT', vTobOrigine.GetString('E_ETABLISSEMENT') )
        // 2. Affectation sur un établissement par défaut
        else vPieceCpt.PutEntete( 'E_ETABLISSEMENT', vPieceCpt.Contexte.EtablisDefaut ) ;
      end ;

  // ----------------------------------------
  // --- Infos Sup pour gestion enca/déca ---
  // ----------------------------------------
  AjouteInfosSupp( vPieceCpt.GetTob( vNumLigne ), vTobScenario, vTobOrigine, vBoContrepartie ) ;

  // ------------------------------------------------
  // -- Formules paramétrés dans le scénario  --
  // ------------------------------------------------
  if vBoContrepartie
    then lIdxChamp := '2'
    else lIdxChamp := '1' ;
  // ( Ref interne... (Gestion spéciale pour les LCR / CHQ ) )
  if AvecNumTraiteChq( vTobScenario ) then
     vPieceCPT.PutValue( vNumLigne, 'E_REFINTERNE',  FormatCheque( vTobScenario.GetValue('REFCHEQUE'), vPieceCpt.GetValue( vNumLigne, 'E_NUMTRAITECHQ') ) )
  else  if vTobScenario.GetValue('CPG_REFINTERNE' + lIdxChamp ) <> '' then
      vPieceCPT.ExecuteFormule( vNumLigne,
                                'E_REFINTERNE',
                                vTobScenario.GetValue('CPG_REFINTERNE' + lIdxChamp ),
                                vTobOrigine ) ;

  // Libelle
  if vTobScenario.GetValue('CPG_LIBELLE' + lIdxChamp ) <> '' then
      vPieceCPT.ExecuteFormule( vNumLigne,
                                'E_LIBELLE',
                                vTobScenario.GetValue('CPG_LIBELLE' + lIdxChamp ),
                                vTobOrigine ) ;

  // Ref externe
  if vTobScenario.GetValue('CPG_REFEXTERNE' + lIdxChamp ) <> '' then
      vPieceCPT.ExecuteFormule( vNumLigne,
                                'E_REFEXTERNE',
                                vTobScenario.GetValue('CPG_REFEXTERNE' + lIdxChamp ),
                                vTobOrigine ) ;

  // Ref libre
  if vTobScenario.GetValue('CPG_REFLIBRE' + lIdxChamp ) <> '' then
      vPieceCPT.ExecuteFormule( vNumLigne,
                                'E_REFLIBRE',
                                vTobScenario.GetValue('CPG_REFLIBRE' + lIdxChamp ),
                                vTobOrigine ) ;

  // ------------------------------------
  // -- MAJ pour émission de bordereau --
  // ------------------------------------
  vPieceCpt.PutValue( vNumLigne, 'E_NUMENCADECA', vTobScenario.GetValue('NUMENCADECA') ) ;

  // ------------------------------------------
  // -- Date d'échéance et mode de paiement  --
  // ------------------------------------------
  if vPieceCPt.GetValue(vNumLigne, 'E_ECHE')='X' then
    begin
    if ( vTobScenario.GetValue('CPG_FORCERMP')='X' ) or
       ( vTobScenario.GetValue('RUPTUREMODEPAIE') = 'X' ) and
       ( vTobScenario.GetValue('CPG_MODEPAIEMENT') <> '' ) then
      vPieceCpt.PutValue( vNumLigne, 'E_MODEPAIE', vTobScenario.GetValue('CPG_MODEPAIEMENT') ) ;
    if ( vTobScenario.GetValue('FORCERDATEECHE')='X' ) then
      vPieceCpt.PutValue( vNumLigne, 'E_DATEECHEANCE', vTobScenario.GetValue('DATEECHE') ) ;
    end ;

  // --------------------------------
  // -- Gestion de la contrepartie --
  // --------------------------------
  lStCpt := '' ;
  lStAux := '' ;
  if lBoMemeDossier then
    begin
    if vBoContrepartie then
      begin
      lStCpt := vPieceCpt.GetValue( vNumLigne - 1, 'E_GENERAL' ) ;
      lStAux := vPieceCpt.GetValue( vNumLigne - 1, 'E_AUXILIAIRE' ) ;
      end
    else lStCpt := vTobScenario.GetString('CPG_GENERAL') ;
    end
  else
    begin
    if vBoContrepartie then
      begin
      lStCpt := vPieceCpt.GetValue( vNumLigne - 1, 'E_GENERAL' ) ;
      lStAux := vPieceCpt.GetValue( vNumLigne - 1, 'E_AUXILIAIRE' ) ;
      end
    else
      lStCpt := GetCompteLiaison( vTobScenario, vPieceCpt.Dossier ) ;
    end ;
  vPieceCpt.PutValue( vNumLigne, 'E_CONTREPARTIEGEN', lStCpt ) ;
  vPieceCpt.PutValue( vNumLigne, 'E_CONTREPARTIEAUX', lStAux ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : X.Maluenda
Créé le ...... : 18/03/2004
Modifié le ... :   /  /
Description .. : Par rapport au pays de la société on fait les verifications de
Suite ........ : coherence sur la selection realisée par l'utilisateur
Mots clefs ... :
*****************************************************************}
Function  VerifieCoherenceSelection( TobOrigine, TobScenario : TOB ) : Boolean ;
var lInCpt    : Integer ;
    lStOldMP  : String ;
    lTobLigne : TOB ;
    lBoTest   : Boolean ;
Begin

  Result := True ;

  // Vérification modes de paiement homogènes si on ne force pas le mode de paiement
  if VH^.PaysLocalisation<>CodeISOES then
     Begin
     if TobScenario.GetValue('CPG_FORCERMP')<>'X' then
       begin
       // As-t-on des mode de paiement qui diffèrent
       lBoTest := False ;
       lStOldMP := '' ;
       for lInCpt := 0 to TobOrigine.Detail.Count - 1 do
         begin
         lTobLigne := TobOrigine.Detail[ lInCpt ] ;
         if lStOldMP = ''
           then lStOldMP := lTobLigne.GetValue('E_MODEPAIE')
           else if lStOldMP <> lTobLigne.GetValue('E_MODEPAIE') then
                   begin
                   lBoTest := True ;
                   break ;
                   end ;

         end ;
       // Si oui, message de confirmation :
       if lBoTest then
          if PgiAsk('Les lignes d''écritures sélectionnées ont des modes de paiement qui diffèrent. Voulez-vous continuer le traitement ?')
             <> mrYes then Result := False ;
       end ;
  End ; //XVI 24/02/2005

  if VH^.PaysLocalisation=CodeISOES then
    if Result then
       Result:=VerifieCoherenceSelection_ESP(TobOrigine,TobScenario) ; //XVI 24/02/2005

end ;


function GetStTri( vTobScenario : TOB ; vBoMultiSoc : Boolean ) : string ;
var lStChpAuxi : string ;
    lStChpGene : string ;
    lBoMultiEtab : boolean ;
begin

  // TIC / TID ou Auxiliaire ?
  if vTobScenario.GetValue('TICTID') = 'X' then
    begin
    lStChpAuxi := 'E_GENERAL' ;
    lStChpGene := '' ;
    end
  else
    begin
    lStChpAuxi := 'E_AUXILIAIRE' ;
    lStChpGene := ';E_GENERAL' ;
    end ;

  // gère-t-on le multi-étab ??
  lBoMultiEtab := ( vTobScenario.GetString('CPG_MULTIETAB')='X') and
                  ( vTobScenario.GetString('CPG_METHODEETAB')='DET') ;

  // Elément principal du tri
  if vTobScenario.GetValue('CPG_GROUPEENCADECA') = 'ECT'
    then result := 'E_DATEECHEANCE;' + lStChpAuxi  // Cas particulier pour la génération par échéance
    else result := lStChpAuxi ;

  // cas multiSoc
  if vBoMultiSoc
    then result := result + ';SYSDOSSIER' ;

  // cas multi-étab (Etablissement prioritaire ou general ?)
  if lBoMultiEtab
    then result := result + ';E_ETABLISSEMENT' + lStChpGene
    else result := result + lStChpGene + ';E_ETABLISSEMENT' ;

  if vTobScenario.GetValue('CPG_GROUPEENCADECA') <> 'ECT'
    then result := result + ';E_DATEECHEANCE' ;

end ;

procedure TriOrigine( vTobOrigine, vTobScenario : TOB ) ;
var lStTri : string ;
begin
  lStTri := GetStTri( vTobScenario, EstMultiSoc and ( vTobOrigine.Detail[0].GetNumChamp('SYSDOSSIER') > 0 ) ) ;
  vTobOrigine.Detail.Sort( lStTri ) ;
end ;


function _TriListeEche( Item1, Item2 : Pointer ) : Integer ;
var lStChp       : string ;
    lStListeChps : string ;
begin
result := 0 ;
lStListeChps := GStTri ;
while lStListeChps <> '' do
  begin
  lStChp := ReadTokenSt( lStListeChps ) ;
  result := StrComp( PChar( Tob(Item1).GetString( lStChp ) ),
                     PChar( Tob(Item2).GetString( lStChp ) ) ) ;
  if result <> 0 then Exit ;
  end ;

end ;


function GetListeEche ( vTobGene : Tob ; vTobScenario : Tob ) : TList ;
var i,j,k      : integer ;
    lPC        : TPieceCompta ;
    lStDossier : string ;
    lTobLigne  : Tob ;
    lStLib     : string ;
begin

  // ============================
  // === Création de la liste ===
  // ============================
  result := TList.Create ;
  for i := 0 to vTobGene.Detail.Count - 1 do
    begin
    lPC        := TPieceCompta( vTobGene.Detail[i] ) ;
    lStDossier := lPC.Dossier ;
    // Echéances de la pièce principale
    for j := 0 to lPC.Detail.Count - 1 do
      begin
      lTobLigne := lPC.Detail[j] ;
      if lTobLigne.GetValue('ESTECHEORIG') = 'X' then
        begin
        // Gestion multi-dossier
        if lTobLigne.GetNumChamp( 'SYSDOSSIER' ) > 0
          then lTobLigne.PutValue('SYSDOSSIER', lStDossier )
          else lTobLigne.AddChampSupValeur('SYSDOSSIER', lStDossier ) ;
        // Gestion cfonb batch
        if ( vTobScenario.GetNumChamp('CFONBBATCH') > 0 ) and ( vTobScenario.GetValue('CFONBBATCH') = 'X' ) then
          begin
          if lPC.Info.LoadAux( lTobLigne.GetString('E_AUXILIAIRE') )
            then lStLib := lPC.Info.Aux_GetValue('T_LIBELLE')
            else lStLib := '' ;
          if lTobLigne.GetNumChamp('T_LIBELLE') > 0
            then lTobLigne.putValue('T_LIBELLE', lStLib )
            else lTobLigne.AddChampSupValeur('T_LIBELLE', lStLib ) ;
          end ;
        result.Add( lTobLigne ) ;
        end ;
      end ;

    // Echéances des pièces sur multi-établissements
    for j := 0 to lPC.PiecesEtab.Count - 1 do
      for k := 0 to TPieceCompta(lPC.PiecesEtab.Objects[j]).Detail.Count - 1 do
        if TPieceCompta(lPC.PiecesEtab.Objects[j]).Detail[k].GetValue('ESTECHEORIG') = 'X' then
          begin
          lTobLigne := TPieceCompta(lPC.PiecesEtab.Objects[j]).Detail[k] ;
          // Gestion multi-dossier
          if lTobLigne.GetNumChamp( 'SYSDOSSIER' ) > 0
            then lTobLigne.PutValue('SYSDOSSIER', lStDossier )
            else lTobLigne.AddChampSupValeur('SYSDOSSIER', lStDossier ) ;
          // Gestion cfonb batch
          if ( vTobScenario.GetNumChamp('CFONBBATCH') > 0 ) and ( vTobScenario.GetValue('CFONBBATCH') = 'X' ) then
            begin
            if lPC.Info.LoadAux( lTobLigne.GetString('E_AUXILIAIRE') )
              then lStLib := lPC.Info.Aux_GetValue('T_LIBELLE')
              else lStLib := '' ;
            if lTobLigne.GetNumChamp('T_LIBELLE') > 0
              then lTobLigne.putValue('T_LIBELLE', lStLib )
              else lTobLigne.AddChampSupValeur('T_LIBELLE', lStLib ) ;
            end ;
          result.Add( lTobLigne ) ;
          end ;
    end ;

  // =======================
  // === Tri de la liste ===
  // =======================
  GStTri   := GetStTri( vTobScenario, ( vTobGene.Detail.Count > 1 ) ) ;
  result.Sort( _TriListeEche ) ;

  // === On désynchronise les lignes d'éché de leur parent pour éviter pb de libération mémoire
  for j := 0 to result.Count - 1 do
    TOB(result[j]).ChangeParent(nil, -1) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 02/03/2004
Modifié le ... : 02/03/2004
Description .. : Fonction globale d'éxecution du scenario contenu dans
Suite ........ : vTobScenario sur un ensemble de ligne d'échéance
Suite ........ : pouvant comprendre :
Suite ........ :  - chargment de l'analytique
Suite ........ :  - génération du reglement (avec éventuellement analytique)
Suite ........ :  - enregistrement du document + lettrage
Suite ........ :  - edition etat / document
Suite ........ :  - export
Suite ........ :
Suite ........ : RESTE A FAIRE :
Suite ........ :   - Gestion des erreurs / transaction
Suite ........ :   - Branchement d'un mécanisme d'export
Mots clefs ... :
*****************************************************************}
Function  ExecuteEncaDeca( vTobOrigine, vTobScenario : TOB ; vTobRefPieces : Tob = nil ) : integer ;
var lTobGene        : TOB ;  // Tob contenant la liste des pièces générées
    lTobEscompte    : TOB ;  // Tob contenant la pièce d'escompte générée
    lTobRemEscompte : TOB ;  // Tob contenant la deuxième pièce pour les Remises a l'escompte générée (ESP)
    lTobFrais       : TOB ;  // Tob contenant la pièce de frais générée (ESP)
    lInError        : TRecError ;
    lBoTransacOk    : Boolean ;
    lInfoEcr        : TInfoEcriture ;
    lInCpt          : Integer ;
    lNumLigne       : Integer ;
    lTobPiece       : TOB ;
    lPieceCpt       : TPieceCompta ;
    lPieceLocal     : TPieceCompta ;
    lPieceEtab      : TPieceCompta ;
    lTLEche         : TList ;
    lMsgCompta      : TMessageCompta ;
    lStEtab         : string ;
    i               : integer ;
    lStModeleBOR    : string ;
    lStFormatCFONB  : string ;
    lInfoTemp       : TInfoEcriture ;
    lTobEcr         : Tob ;
    lStExo          : string ;
    lInCodeErr      : integer ;

    procedure _LibereMemoire ;
      var k : integer ;
      begin
      // Libération des TPieceCompta
      if Assigned( lTobGene ) then
        begin
        for k := (lTobGene.Detail.Count - 1) downto 0 do
          begin
          lPieceCpt := TPieceCompta(lTobGene.Detail[ k ]) ;
          lPieceCpt.Free ;
          end ;
        FreeAndNil( lTobGene ) ;
        end ;
      if Assigned( lTobEscompte ) then
        begin
        for k := (lTobEscompte.Detail.Count - 1) downto 0 do
          begin
          lPieceCpt := TPieceCompta(lTobEscompte.Detail[ k ]) ;
          lPieceCpt.Free ;
          end ;
        FreeAndNil( lTobEscompte ) ;
        end ;
      // Autres Tob
      if Assigned( lTobRemEscompte ) then
        begin
        lTobRemEscompte.ClearDetail ;
        FreeAndNil( lTobRemEscompte ) ;
        end ;
      if Assigned( lTobFrais ) then
        begin
        lTobFrais.ClearDetail ;
        FreeAndNil( lTobFrais ) ;
        end ;
      // TInfoEcriture
      if Assigned( lInfoEcr ) then
        FreeAndNil( lInfoEcr ) ;

      // autres objets
      if Assigned( GErrEncaDeca ) then
        FreeAndNil( GErrEncaDeca ) ;

      if Assigned( lTLEche ) then
        begin
        VideListe( lTLEche ) ;
        FreeAndNil( lTLEche ) ;
        end ;

      if Assigned( lMsgCompta ) then // FQ 21090 SBO 07/08/2007 : Pb génération de l'escompte
        FreeAndNil( lMsgCompta ) ;

      lPieceCpt := nil ;
      end ;

begin
  V_PGI.IoError := oeOk ;
  lboTransacOk := (VH^.PaysLocalisation<>CodeISOES) ; //XVI 24/02/2005
  Result       := CGE_ERRTRANSAC ;
  lInfoEcr     := nil ;
  lPieceLocal  := nil ;

  // Complémente le scénario pour éviter multiples chargement de données
  CompleteScenario ( vTobScenario ) ;

  // Tri de la tob de départ ( // FQ16986 SBO 03/11/2005 )
  TriOrigine( vTobOrigine, vTobScenario ) ;

  // Chargement de l'analytique si génération aux détails
  if vTobScenario.GetValue('CPG_METHODEANA') = 'DET' then
    begin
    for lInCpt := 0 to ( vTobOrigine.detail.count - 1 ) do
      if vTobOrigine.detail[ lInCpt ].GetString('E_ANA') = 'X' then
        CChargeAna ( vTobOrigine.detail[ lInCpt ] ) ;
    end ;

  // Tob contenant les pièces générées
  lTobGene        :=  TOB.Create( 'TOB_ENCADECA',     nil, -1) ;
  lTobEscompte    :=  TOB.Create( 'TOB_ESCOMPTE',     nil, -1) ;
  lTobRemEscompte :=  TOB.Create( 'TOB_REMESCOMPTE',  nil, -1) ;
  lTobFrais       :=  TOB.Create( 'TOB_FRAIS',        nil, -1) ;

  // Génération des pièces comptables
  GenereEncaDeca ( vTobOrigine, vTobScenario, lTobGene, lTobEscompte, lTobRemEscompte, lTobFrais, vTobRefPieces ) ;

  // Gestion des tiers avec cumuls à 0 // FQ 16986
  if CorrigeLigneAZero( lTobGene, lTobEscompte, lTobRemEscompte, lTobFrais ) then
    if ( PgiAsk( TraduireMemoire('Votre sélection contient des cumuls à zéro pour certains tiers. Les écritures des tiers concernés seront lettrées. Souhaitez-vous poursuivre le traitement ?'), TraduireMemoire('Génération d''enca/déca') ) <> mrYes ) then
      begin
      _LibereMemoire ;
      Exit ;
      end ;

  // Gestion des traites négatives
  if not SensReglementOk( vTobScenario, lTObGene, lInCodeErr ) then
      begin
      result := lInCodeErr ;
      _LibereMemoire ;
      Exit ;
      end ;

  // Vérif exercice sur dossier cible
  for lInCpt := 0 to lTobGene.Detail.Count - 1 do
    begin
    lPieceCpt := TPieceCompta( lTobGene.Detail[ lInCpt ] ) ;
    if not lPieceCpt.Contexte.local then
      begin
      lStExo    := lPieceCpt.Contexte.Exercices.QuelExoDT( lPieceCpt.GetEnteteDt('E_DATECOMPTABLE') ) ;
      if (lStExo='') then
        begin
        PgiInfo( TraduireMemoire('Aucun exercice n''existe pour la date de génération sur le dossier ') + lPieceCpt.Dossier + ' !' ) ;
        result := CGE_NOEXODOSSIER ;
        end
      else if lPieceCpt.Contexte.Exercices.EstExoClos( lStExo ) then
        begin
        PgiInfo( TraduireMemoire('L''exercice ' + lStExo + ' est clos sur le dossier ') + lPieceCpt.Dossier) ;
        result := CGE_EXOCLOS ;
        end
      else continue ;
      // si on arrive la c'est qu'il y a eu une erreur
      _LibereMemoire ;
      Exit ;
      end ;
    end ;

  // Visualisation ?
  if ( vTobScenario.GetValue('CPG_AVECVERIFPIECE') = 'X' ) then
    for lInCpt := 0 to lTobGene.Detail.Count - 1 do
      begin
      lPieceCpt := TPieceCompta( lTobGene.Detail[ lInCpt ] ) ;
      if not ModifiePieceCompta( lPieceCpt, taModif ) then
        begin
        result := CGE_ANNULATION ;
        _LibereMemoire ;
        Exit ;
        end ;
      end ;

  if lTobGene.Detail.Count > 0  then
    begin
    // ================================
    // Vérification des pièces générées
    // ================================
    // *** Vérification des pièces générées ***
    for lInCpt := 0 to lTobGene.Detail.Count - 1 do
      begin
      lPieceCpt := TPieceCompta( lTobGene.Detail[ lInCpt ] ) ;
      if lPieceCpt.Contexte.local then
        begin
        lPieceLocal := lPieceCpt ;
        lInfoEcr    := lPieceCpt.Info ;
        end ;
      if not lPieceCpt.IsValidPiece then
        begin
        if lPieceCpt.Contexte.local
          then result := CGE_ERRVALIDEPIECE
          else result := CGE_ERRVALIDEPIECEMS ;
        _LibereMemoire ;
        Exit ;
        end ;
      end ;

    // *** Vérification des pièces d'escompte ***
    for lInCpt := 0 to lTobEscompte.Detail.Count - 1 do
      begin
      // FQ 21090 SBO 07/08/2007 : Pb génération de l'escompte (Message clair)
      lPieceCpt := TPieceCompta( lTobEscompte.Detail[ lInCpt ] ) ;
      if not lPieceCpt.IsValidPiece then
        begin
        result := CGE_ERRVALIDEESC ;
        _LibereMemoire ;
        Exit ;
        end ;
      end ;
    // *** Vérification de la pièce de Risque***
    if (VH^.PaysLocalisation=CodeISOES) and (ltobRemEscompte.Detail.count>0) then
       Begin
       lInError := CIsValidSaisiePiece( lTobRemEscompte, lInfoEcr ) ;
       if lInError.RC_Error <> RC_PASERREUR then
         begin
         lMsgCompta := TMessageCompta.Create( 'Génération de la pièce de Risque', msgSaisiePiece ) ;
         lMsgCompta.Execute( lInError.RC_Error ) ;
         result := CGE_ERRTRAITEMENT ;
         _LibereMemoire ;
         Exit ;
         end ;
       End ;

    // *** Vérification des pièces de frais ***
    if (VH^.PaysLocalisation=CodeISOES) and (lTobFrais.Detail.count>0) then
       for lInCpt := 0 to lTobFrais.Detail.Count - 1 do
         begin
         lTobPiece := lTobFrais.Detail[ lInCpt ] ;
         lInError := CIsValidSaisiePiece( lTobPiece, lInfoEcr ) ;
         if lInError.RC_Error <> RC_PASERREUR then
           begin
           lMsgCompta := TMessageCompta.Create( 'Génération de la pièce des frais', msgSaisiePiece ) ;
           lMsgCompta.Execute( lInError.RC_Error ) ;
           result := CGE_ERRTRAITEMENT ;
           _LibereMemoire ;
           Exit ;
           end ;
         end ; //XVI 24/02/2005

    // ==========================
    // Enregistrements + Lettrage
    // ==========================
    // Par défaut on considère que la transaction n'aboutie pas si on doit comptabiliser
    // Transaction ok si pas de comptabilisation et pas besoin de maj E_NUMTRAITECHQ
    lBoTransacOk := not ( AvecNumTraiteChq( vTobScenario ) or
                          ( vTobScenario.getValue('CPG_COMPTABILISE') = 'X' )
                        ) ;
    if not lBoTransacOk then
      begin
      try
          // *********************************
          // **** Debut de la transaction ****
          // *********************************
          BeginTrans ;

          // *************************************
          // **** Comptabilisation + Lettrage ****
          // *************************************
          if vTobScenario.getValue('CPG_COMPTABILISE') = 'X' then
            begin
            // =>  Affectation des numéros de pièces type NORMAL !!!
            if not UpdateNumeroPiece ( lTobGene, lTobEscompte, lTobRemEscompte, lTobFrais, vTobScenario ) then
              begin
              result := CGE_ERRNUMPIECE ;
              MessageAlerte('Erreur lors de l''affectation des numéros de pièces finaux !' ) ;
              Exit ;
              end ;
            // => Enregistrement des pièces
            if not SauveEncaDeca ( lTobGene, lTobEscompte, lTobRemEscompte, lTobFrais, vTobScenario ) then
              begin
              result := CGE_ERRSAVEPIECE ;
              MessageAlerte('Erreur lors de l''enregistrement de la pièce !' );
              Exit ;
              end ;

            // ================
            // === Lettrage ===
            // ================
            //YMO 15/12/2005 Pour les pièces en devise, choix de ne pas lettrer
            if ( vTobScenario.getValue('SANSLETTRAGE') = '-' ) // gestion global lettrage
              and ((lPieceLocal.Devise.Code = V_PGI.DevisePivot) or (vTobScenario.getValue('CPG_LETTRAGEAUTO') = 'X')) then
            begin
            // => Lettrage
              for lInCpt := 0 to lTobGene.Detail.Count - 1 do
                begin
                V_PGI.IoError := oeOk ;
                lPieceCpt := TPieceCompta(lTobGene.Detail[lInCpt]) ;
                lStEtab   := '' ;
                // Lettrage des pièces sur les établissements
                if lPieceCpt.PiecesEtab.count > 0 then
                  begin
                  For i := 0 to ( lPieceCpt.PiecesEtab.Count - 1 ) do
                    begin
                    lPieceEtab := TPieceCompta( lPieceCpt.PiecesEtab.Objects[ i ] ) ;
                    lStEtab    := lPieceEtab.GetEntete('E_ETABLISSEMENT') ;
                    LettrerEncaDeca ( vTobOrigine, lPieceEtab, lPieceEtab.Devise, lStEtab ) ;
                    if V_PGI.IoError<>oeOk then
                      begin
                      result := CGE_ERRLETTPIECE ;
                      MessageAlerte('Erreur lors du lettrage de la pièce !' );
                      Exit ;
                      end ;
                    end ;
                  lStEtab    := lPieceCpt.GetEntete('E_ETABLISSEMENT') ;
                  end ;
                // Lettrage de la pièce principale sur la société
                LettrerEncaDeca ( vTobOrigine, lPieceCpt, lPieceCpt.Devise, lStEtab ) ;
                if V_PGI.IoError<>oeOk then
                  begin
                  result := CGE_ERRLETTPIECE ;
                  MessageAlerte('Erreur lors du lettrage de la pièce !' );
                  Exit ;
                  end ;
                end ;
            end  // si on ne lettre pas la pièce en devise, on met à jour E_NUMTRAITECHQ
            // ====================
            // === FIN Lettrage ===
            // ====================
            else if AvecNumTraiteChq( vTobScenario ) then
            begin
            // => maj des pièces d'origine
            if not UpdateEcheancesOrigines ( vTobOrigine, vTobScenario) then
              begin
              result := CGE_ERRUPDATEECHE ;
              MessageAlerte('Erreur lors de la mise à jour des pièces d''origine !' );
              Exit ;
              end ;
            end ;
          end

          // *******************************************************
          // **** OU Maj du E_NUMTRAITECHQ des pièces d'origine ****
          // *******************************************************
          else if AvecNumTraiteChq( vTobScenario ) then
            begin
            // => maj des pièces d'origine
            if not UpdateEcheancesOrigines ( vTobOrigine, vTobScenario) then
              begin
              result := CGE_ERRUPDATEECHE ;
              MessageAlerte('Erreur lors de la mise à jour des pièces d''origine !' );
              Exit ;
              end ;
            end ;

          // *******************************
          // **** Fin de la transaction ****
          // *******************************
          CommitTrans ;
          lBoTransacOk := True ; // Si on arrive, le traitement s'est correctement passé

        finally

          // retourne la 1ère ligne de chaque pièce
          if Assigned( vTobRefPieces ) then
            begin
            for lInCpt := 0 to lTobGene.Detail.Count - 1 do
              begin
              lPieceCpt := TPieceCompta( lTobGene.Detail[ lInCpt ] ) ;
              // garde ref de la pièce
              lTobEcr := Tob.Create( 'ECRITURE', vTobRefPieces, -1 ) ;
              lTobEcr.Dupliquer( lPieceCpt.Detail[ lPieceCpt.Detail.count - 1 ], False, True ) ;
              //garde ref des pièce multi-etab
              For lNumLigne := 1 to lPieceCpt.PiecesEtab.count do
                begin
                // garde ref de la pièce
                lTobEcr := Tob.Create( 'ECRITURE', vTobRefPieces, -1 ) ;
                lTobEcr.Dupliquer( TPieceCompta( lPieceCpt.PiecesEtab.Objects[ lNumLigne ] ).Detail[ TPieceCompta( lPieceCpt.PiecesEtab.Objects[ lNumLigne ] ).Detail.count - 1 ], False, True ) ;
                end ;
              end ;

            end ;

          // ***********************************
          // **** En cas d'erreur, rollback ****
          // ***********************************
          if not lBoTransacOk then
            RollBack ;
        end ;

      end ;

    // Le reste du traitement n'est à effectuer que si la transaction ci-dessus est ok
    if lBoTransacOk and
       ( ( vTobScenario.GetValue('CPG_EDITE') = 'X' )        or
         ( vTobScenario.GetValue('CPG_CFONBEXPORT') = 'X' )  or
         ( vTobScenario.GetValue('CPG_BORDEREAUEXP') = 'X' ) ) then
      begin

      // Attention ! Si multi étab sans comptabilisation, il faut simuler l'enregistrement
      // de la pièce pour générer les pièces sur les établissements
      if vTobScenario.getValue('CPG_COMPTABILISE') <> 'X' then
        for lInCpt := 0 to lTobGene.Detail.Count - 1 do
          begin
          lPieceCpt := TPieceCompta( lTobGene.Detail[ lInCpt ] ) ;
          if lPieceCpt.EstMultiEtab then
            For lNumLigne := 1 to lPieceCpt.Detail.Count do
              lPieceCpt.TraiteMultiEtab( lNumLigne ) ;
          end ;

      // pour la suite des traitements, on se base sur une liste des échéances tiers
      lTLEche := GetListeEche( lTobGene, vTobScenario ) ;
      if vTobScenario.GetValue('CFONBBATCH') <> 'X' then
        AjouteInfosEtat ( lTLEche, vTobScenario, lInfoEcr ) ;

      // libération pièces princpales maintenant pour éviter pb mémoire
      if Assigned( lTobGene ) then
        begin
        for lInCpt := (lTobGene.Detail.Count - 1) downto 0 do
          begin
          lPieceCpt := TPieceCompta(lTobGene.Detail[ lInCpt ]) ;
          lInfoTemp := lPieceCpt.Info ;
          FreeAndNil( lPieceCpt ) ;
          if Assigned( lInfoTemp ) and ( lInfoTemp.Dossier <> V_PGI.SchemaName ) then
            FreeAndNil( lInfoTemp ) ;
          end ;
        FreeAndNil( lTobGene ) ;
        end ;


      // Editions
      if vTobScenario.GetValue('CPG_EDITE') = 'X' then
        begin
        if not GenereEtats( lTLEche, vTobScenario ) then
          result := CGE_ERREDITION ;
        end ;

      // MAJ des paramSoc pour les numéro de Traites / BOR
      if AvecNumTraiteChq( vTobScenario ) and ( vTobScenario.GetValue('NOMPARAMSOCTRAITE') <> '' ) then
        SetParamSoc( vTobScenario.GetValue('NOMPARAMSOCTRAITE'), vTobScenario.GetValue('NUMCHEQUE') + 1 ) ;

      // Cas non espagnol
      if (VH^.PaysLocalisation<>CodeISOES) then
        begin
        // Export CFONB + Bordereau
        if ( vTobScenario.GetValue('CPG_CFONBEXPORT') = 'X') then
          begin

          if (vTobScenario.GetValue('CPG_BORDEREAUEXP') = 'X')
            then lStModeleBOR := vTobScenario.GetValue('CPG_BORDEREAUMOD')
            else lStModeleBOR := '' ;

          if vTobScenario.GetValue('CPG_TYPEENCADECA') = 'LCR'
            then lStFormatCFONB := 'LEN'
            else lStFormatCFONB  := vTobScenario.GetValue('CPG_TYPEENCADECA') ;

          if vTobScenario.GetValue('CFONBBATCH') = 'X'
            then ExportCFONBBatch( vTobScenario, lTLEche )
            else ExportCFONBMS( vTobScenario, lTLEche ) ;

          end
          // Bordereau seul
        else if (vTobScenario.GetValue('CPG_BORDEREAUEXP') = 'X') then
          begin
          if not ExecuteEmissionBOR  ( lTLEche, vTobScenario.GetValue('CPG_BORDEREAUMOD'), vTobScenario.GetValue('CPG_BORDEREAUCHX')='X', vTobScenario.GetValue('APERCU')='X' ) then
            result := CGE_ERREDITBOR ;
          end
        end
      end
      // Cas espagnol
      else if (vTobScenario.GetValue('CPG_CFONBEXPORT') = 'X') then
        lBoTransacOk := GenereExports( lTLEche, vTobScenario, lInfoEcr.Devise.Dev ) ;

      end ;


  // Libération
  _LibereMemoire ;

  if lBoTransacOk
    then result := CGE_PASERREUR
    else if result = CGE_PASERREUR then
         result := CGE_ERRTRANSAC ;

end ;


Function  RuptureEtab       ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ; var vEtab : string ) : Boolean ;
begin
  // Si aucune ligne a traiter, la rupture est vrai, on sort...
  result := not assigned( vTobLigne ) ;
  if result then Exit ;

//  lStGroupe := vTobScenario.GetValue('CPG_GROUPEENCADECA') ;

  // Test Sauf première ligne
  if vInIndex>0 then
    result :=   ( ( vTobScenario.GetString('CPG_MULTIETAB')='X') and
                       ( vTobScenario.GetString('CPG_METHODEETAB')='DET') and
                       ( vEtab <> vTobLigne.GetString('E_ETABLISSEMENT') ) ) ;

  // MAJ des champs de rupture
  if Result or ( vInIndex = 0 ) then
    vEtab        := vTobLigne.GetValue('E_ETABLISSEMENT') ;

end ;


Function  RuptureEncaDeca   ( vInIndex : Integer ; vTobLigne, vTobScenario  : TOB ;
                              var vGeneral : String ; var vAuxiliaire : String ; var vDateEche : TDateTime ; var vEtab : string ) : Boolean ;
var lStGroupe   : String ;
    lBoRuptEtab : Boolean ;
begin

  // Si aucune ligne a traiter, la rupture est vrai, on sort...
  result := not assigned( vTobLigne ) ;
  if result then Exit ;

  lStGroupe := vTobScenario.GetValue('CPG_GROUPEENCADECA') ;

  // Première ligne
  if vInIndex>0 then
    begin

    lBoRuptEtab :=   ( ( vTobScenario.GetString('CPG_MULTIETAB')='X') and
                       ( vTobScenario.GetString('CPG_METHODEETAB')='DET') and
                       ( vEtab <> vTobLigne.GetString('E_ETABLISSEMENT') ) ) ;

    // Génération au détail, Rupture à chaque ligne
    if lStGroupe='DET'
      then Result := True

    // Génération globalisé, pas de rupture
    else if lStGroupe='GLO'
      then Result := lBoRuptEtab

    // Rupture par auxiliaire (et par généraux)
    else if lStGroupe='AUX' then
      begin
      if vTobScenario.GetValue('TICTID') = 'X'
        then Result := ( (vGeneral<>vTobLigne.GetValue('E_GENERAL'))       or
                         lBoRuptEtab
                         )
        else Result := ( (vAuxiliaire<>vTobLigne.GetValue('E_AUXILIAIRE')) or
                         lBoRuptEtab
                         ) ;
      end

    // Rupture par échéance
    else if lStGroupe='ECT'
      then Result:=(vDateEche<>vTobLigne.GetValue('E_DATEECHEANCE')) or lBoRuptEtab

    // Rupture par auxiliaire/général/échéance
    else if lStGroupe='ECH' then
      begin
      if vTobScenario.GetValue('TICTID') = 'X'
        then Result:=( (vGeneral<>vTobLigne.GetValue('E_GENERAL'))       or
                       (vDateEche<>vTobLigne.GetValue('E_DATEECHEANCE')) or
                       lBoRuptEtab
                       )
        else Result:=( (vAuxiliaire<>vTobLigne.GetValue('E_AUXILIAIRE')) or
                       (vDateEche<>vTobLigne.GetValue('E_DATEECHEANCE')) or
                       lBoRuptEtab
                       ) ;
      end ;

    end ;

  // MAJ des champs de rupture
  if Result or ( vInIndex = 0 ) then
    begin
    vAuxiliaire  := vTobLigne.GetValue('E_AUXILIAIRE') ;
    vGeneral     := vTobLigne.GetValue('E_GENERAL') ;
    vDateEche    := vTobLigne.GetValue('E_DATEECHEANCE') ;
    vEtab        := vTobLigne.GetValue('E_ETABLISSEMENT') ;
    end ;

end ;



procedure GereNomLot( vTobScenario, vTobGene, vTobRefPieces : Tob ) ;
var i,j         : integer ;
    lTobEcr     : Tob ;
begin

  for i := 0 to vTobGene.Detail.count - 1 do
    begin
    // MAJ du nomlot
    for j := 0 to vTobGene.Detail[i].Detail.count - 1 do
      begin
      lTobEcr := vTobGene.Detail[i].Detail[j] ;
      if ( lTobEcr.GetValue('ESTECHEORIG')<>'X' ) and
         ( lTobEcr.GetValue('E_GENERAL') = vTobScenario.GetValue('CPG_GENERAL') ) then
        begin
        lTobEcr.PutValue('E_NOMLOT', vTobScenario.GetValue('NOMLOT') ) ;
        end ;
      end ;  // for j

    end ; // for i

end ;

procedure GlobaliseMultiSoc( vTobScenario : Tob ; vPieceCpt : TPieceCompta ) ;
var lInIndex    : integer ;
    lTobEcr     : Tob ;
    lNumL       : integer ;
    lStCpt      : string ;
    lTobTemp    : Tob ;
    lInLast     : integer ;
    lInContr    : integer ;
    lTobContr   : Tob  ;
    lStContr    : string ;
    lStDossier  : string ;
begin

  if (vTobScenario.GetValue('CPG_GROUPEENCADECA')<>'GLO') then Exit ;

  for lInIndex := 0 to vTobScenario.detail.Count - 1 do
    begin

    lStCpt      := vTobScenario.detail[ lInIndex ].GetString('CLS_GENERAL') ;
    lStDossier  := vTobScenario.detail[ lInIndex ].GetString('CLS_DOSSIER') ;
    lTobEcr     := vPieceCpt.FindFirst( ['E_GENERAL','DOSSIERDEST'], [lStCpt,lStDossier], True ) ;

    if lTobEcr = nil then continue ;

    lTobTemp := Tob.Create('ECRITURE', nil, -1 ) ;
    lTobTemp.Dupliquer( lTobEcr, True, True ) ;

    // Effacement des lignes sur le comptes
    while lTobEcr <> nil do
      begin
      lInLast := lTobEcr.GetInteger('E_NUMLIGNE') ;
      vPieceCpt.DeleteRecord( lInLast ) ;
      lTobEcr := vPieceCpt.FindFirst( ['E_GENERAL','DOSSIERDEST'], [lStCpt,lStDossier], True ) ;
      end ;

    // recherche index d'insertion (sur la contrepartie concernée)
    if vTobScenario.GetString('CPG_MULTISOCMETH')='GLO' then
      begin
      lStContr    := vTobScenario.GetString('CPG_GENERAL') ;
      lTobContr   := vPieceCpt.FindFirst( ['E_GENERAL'], [lStContr], False ) ;
      if lTobContr = nil
        then lInContr := -1
        else lInContr := lTobContr.GetInteger('E_NUMLIGNE') ;
      end
    else
      begin
      lStContr    := vTobScenario.GetString('CPG_GENERAL') ;
      lTobContr   := vPieceCpt.FindFirst( ['E_GENERAL','DOSSIERDEST'], [lStContr,lStDossier], True ) ;
      if lTobContr = nil
        then lInContr := -1
        else lInContr := lTobContr.GetInteger('E_NUMLIGNE') ;
      end ;

    // Ajout de la ligne global
    vPieceCpt.NewRecord( lInContr ) ;
    lNumL   := vPieceCpt.CurIdx ;

    // -- MAJ du Compte
    vPieceCpt.PutValue( lNumL,  'E_GENERAL',    lStCpt ) ;
    // -- MAJ des infos
    RenseigneEcriture( vTobScenario, lTobTemp, vPieceCpt, lNumL, True ) ;

    vPieceCpt.PutValue( lNumL, 'E_REFINTERNE', '' ) ;
    vPieceCpt.PutValue( lNumL, 'E_LIBELLE',    lTobTemp.GetString('E_LIBELLE') ) ;

    // -- MAJ des montants
    vPieceCpt.AttribSolde( lNumL ) ;

    // Libération Tob Temporaire
    lTobTemp.Free ;

    end ;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/01/2004
Modifié le ... : 30/01/2006
Description .. : Algorithme global de génération de la pièce d'enca/déca
Suite ........ : suivant indication du scénario dans TobScenario, en
Suite ........ : prenant pour base les lignes d'échéance d'origine de
Suite ........ : TobOrigine.
Suite ........ : Retourne un pointeur sur la tob générée.
Suite ........ :
Suite ........ : 30/01/2006 : Adaptation au TPieceCompta
Mots clefs ... :
*****************************************************************}
Procedure GenereEncaDeca  ( TobOrigine, TobScenario, vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais, vTobRefPieces : TOB ) ;
 var lTobEche      : TOB ;       // Pointeur sur la ligne génerée
    lInIndex      : Integer ;   // Compteur
    lRuptGene     : String ;    // Valeur de E_GENERAL pour gestion des ruptures
    lRuptAuxi     : String ;    // Valeur de E_AUXILIAIRE pour gestion des ruptures
    lRuptDateEche : TDateTime ; // Valeur de E_DAETECHEANCE pour gestion des ruptures
    lInLastRupt   : Integer ;   // Index de la dernière rupture
    lBoRupt       : Boolean ;
    lEcheAuxi     : String ;
    lEcheGene     : String ;
    lInRuptEche   : Integer ;
    lBoRuptEche   : Boolean ;
    lStCodeMP     : String ;
    lBoPremierMP  : Boolean ;
    lNumL         : Integer ;
    ldbFraisGen   : Double  ;   // Cumule les frais gobales de la remise  //XMG 05/04/04
    lPieceCpt     : TPieceCompta ;
    lBoRuptMS     : Boolean ; // indicateur de rupture multisociété
    lEcheDossier  : string ;
    lEcheEtab     : string ;
    lRuptEtab     : string ;
begin

  lRuptGene     := '' ;
  lRuptAuxi     := '' ;
  lEcheAuxi     := '' ;
  lEcheGene     := '' ;
  lEcheDossier  := V_PGI.SchemaName ;
  lRuptDateEche := iDate1900 ;
  lInLastRupt   := 0 ;
  lInRuptEche   := 0 ;
  lStCodeMP     := '' ;
  lBoPremierMP  := True ;
  ldbFraisGen   := 0 ;
  TobScenario.PutValue('RuptureModePaie', '-' ) ;

  // Initialisation de la Piece principale (sur la base local)
  lPieceCpt := CreerPieceEncaDeca( TobScenario, TobOrigine.Detail[0], vTobGene )  ;

  // Paramètrage opti
  lPieceCpt.ChargeInfoOpti( TobOrigine ) ;

  // Gestion du E_NumTraiteCHQ dans les échéances d'origine
  if AvecNumTraiteChq( TobScenario ) then
    MajNumTraiteChq( TobOrigine, TobScenario ) ;

  // Parcours des lignes d'échéances pour génération des lignes de contrepartie
  // + 1 fois supplémentaire pour traitement de la dernière rupture
  for lInIndex:=0 to TobOrigine.Detail.Count do
    begin

    lTobEche := nil ;
    if lInIndex <= ( TobOrigine.Detail.Count - 1 ) then
      lTobEche := TobOrigine.Detail[lInIndex] ;

    // ==> Tests des ruptures sur méthodes de génération <== ATTENTION ORDRE IMPORTANT CAR RuptureDossier peut modifier lEcheDossier
{1} lBoRupt     := RuptureEncaDeca ( lInIndex, lTobEche, TobScenario, lRuptGene, lRuptAuxi, lRuptDateEche, lRuptEtab ) ;
{2} lBoRuptEche := RuptureEche ( lInIndex, lTobEche, TobScenario, lEcheAuxi, lEcheGene, lEcheEtab ) ; // and (lEcheDossier=V_PGI.SchemaName) ;
{3} lBoRuptMS   := RuptureDossier  ( lInIndex, lTobEche, TobScenario, lEcheDossier ) ;

    // ==================================================================================================
    // TRAITEMENT DE LA GLOBLISATION DES LIGNES DE TIERS ==> Test rupture pour la globalisation des tiers
    // ==================================================================================================
    if ( lInIndex > 0 ) and ( lBoRupt or lBoRuptMS or lBoRuptEche ) then
      begin
      // Génération des lignes d'échéance tiers
      GenereEcriture( TobScenario, TobOrigine, lInRuptEche, lInIndex - 1, vTobGene ) ;
      // Maj indicateur de rupture
      lInRuptEche   := lInIndex ;
      end ;

    // =============================
    // TRAITEMENT DE LA CONTREPARTIE
    // =============================
    if ( lInIndex > 0 ) and lBoRupt   // Dernière ligne ou rupture sur contrepartie ?
//       or ( (TobScenario.GetValue('CPG_MULTISOCMETH') = 'SOC') and (lBoRuptMS) ) ) // Rupture dossier
     then
      begin

      // Si Différents modes de paiement détectés
      // --> MAJ du mode de paiement pour les dernières d'echéances générées jusqu'à la dernière ligne de contrepartie
      if ( TobScenario.GetValue('CPG_FORCERMP') <> 'X' ) and
         ( TobScenario.GetValue('RuptureModePaie') = 'X' ) and
         ( TobScenario.GetValue('CPG_MODEPAIEMENT') <> '' ) then
         begin
         for lNumL := lPieceCpt.Detail.Count downto 1 do
           begin
           if lPieceCpt.GetValue( lNumL, 'ESTECHEORIG' ) <> 'X' then Break ;
           if lPieceCpt.GetValue( lNumL, 'E_ECHE' ) = 'X' then
             lPieceCpt.PutValue( lNumL, 'E_MODEPAIE', TobScenario.GetValue('CPG_MODEPAIEMENT') ) ;
           end ;
         end ;

      // Génération de la ligne de contrepartie
      GenereEcriture( TobScenario, TobOrigine, lInLastRupt, lInIndex - 1, vTobGene, True ) ;

      // maj des variables de ruptures
      lInLastRupt   := lInIndex ;
      if ( TobScenario.GetValue('CPG_FORCERMP') <> 'X' ) and
         ( TobScenario.GetValue('CPG_MODEPAIEMENT') <> '' ) then
        begin
        lBoPremierMP  := True ;
        TobScenario.PutValue('RuptureModePaie','-') ;
        end ;

       //Calculs des frais Pour la version espagnole....
      if ( VH^.PaysLocalisation=CodeISOES) and ( lInIndex = TobOrigine.Detail.Count ) then
         Begin
         //Calculs des frais globales .... (A faire)
         //géneration de la pièce des frais (Si Globals...)
         CalculeFraisRemiseParType_Esp(TobOrigine.Detail[lInIndex-1],Tobscenario,vtobFrais,ldbFraisGen,lPieceCpt.Devise,'REM',lPieceCpt.Info ) ;
         //géneration de la pièce des frais (contrepartie...)
         CalculeContrepartieFraisRemise_ESP(TobOrigine.Detail[lInIndex-1],TobScenario,VTobFrais,lPieceCpt.Devise,lPieceCpt.Info ) ;
         //Génération de la pièce supplementaire des remises en Escompte. (contrepartie....)
         GenereContrepartieEcritureRisque_ESP(TobOrigine.Detail[lInIndex-1],Tobscenario,vTobRemEscompte,lPieceCpt.Devise,lPieceCpt.Info ) ;
         End ; //XVI 24/02/2005

      end ;

    // ===============================================
    // TRAITEMENT A FAIRE POUR CHAQUE LIGNE D'ECHEANCE
    // ===============================================
    if lInIndex <= ( TobOrigine.Detail.Count - 1 ) then
      begin

      // Gestion de l'escompte
      //  Attention : peut modifier le montant de l'échéance d'origine laisser rester transparent par au traitement
      GenereEscompte( lTobEche, vTobEscompte, TobScenario, lPieceCpt.Info ) ;

      // Test de l'uniformisation du mode de paiement
      if ( TobScenario.GetValue('CPG_FORCERMP') <> 'X' ) and
         ( TobScenario.GetValue('CPG_MODEPAIEMENT') <> '' ) then
        if RuptureModePaie( lTobEche, TobScenario, lStCodeMP, lBoPremierMP ) then
          TobScenario.PutValue('RuptureModePaie','X') ;

      //Calculs des frais Pour la version espagnole....
      if (VH^.PaysLocalisation=CodeISOES) then
          Begin
          //Génération de la pièce supplementaire des remises en Escompte.
          CalculeEcritureRisque_ESP(lTobEche,TobScenario,vTobRemEscompte,lPieceCpt.Devise,lPieceCpt.Info ) ;
          //Génération Pièce Frais....
          CalculeFraisRemise_ESP(lTobEche,TobScenario,vTobFrais,lPieceCpt.Devise,ldbFraisGen,lPieceCpt.Info ) ;
          End ; //XVI 24/02/2005

      end ;

    end ;

  // ========
  // FINITION
  // ========
  // Globalisation des comptes de liaison si demandé
  if (TobScenario.GetValue('CPG_GROUPEENCADECA')='GLO') and (vTobGene.Detail.count > 1) then
    GlobaliseMultiSoc ( TobScenario, lPieceCpt ) ;

  // Spécif ESP
  if (VH^.PaysLocalisation=CodeISOES) then
     Begin

     if vTobRemEscompte.Detail.count > 0 then
        TerminePieceEncaDecaESP ( TOBOrigine, vTobRemEscompte, TobScenario, lPieceCpt.Devise ) ;

     if vTobFrais.Detail.count > 0 then
        for lInIndex:=0 to VtobFrais.Detail.Count-1 do
            TerminePieceEncaDecaESP ( TOBOrigine, VtobFrais.Detail[lInIndex], TobScenario, lPieceCpt.Devise ) ;
     End ; //XVI 24/02/2005

  // Gestion de la récap pour génération en batch (multi-lot)
  if ( TobScenario.GetValue('ENCADECABATCH')='X' ) then
    GereNomLot( TobScenario, vTobGene, vTobRefPieces ) ;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/01/2004
Modifié le ... : 30/01/2006
Description .. : Génère la ligne d'écriture de règlement correspondant à la
Suite ........ : ligne de contrepartie aux échéances tiers.
Suite ........ : Les montants affectés sont passés en paramètres pour
Suite ........ : correspondre à la méthode de génération utilisée.
Suite ........ : Rem : l'analytique est générée plus tard dans
Suite ........ : GenereAnalytique.
Suite ........ : OU
Suite ........ : Génère la ligne d'écriture de règlement correspondant à la
Suite ........ : ligne d'échéance d'origine du tiers.
Suite ........ : Rem : l'analytique est générée plus tard dans
Suite ........ : GenereAnalytique.
Suite ........ :
Suite ........ : 31/01/2006 : Adaptation au TPieceCompta
Mots clefs ... :
*****************************************************************}
Procedure GenereEcriture ( vTobScenario, vTobOrigine : TOB ; vInDebut, vInFin : Integer ; vTobGene : Tob ; vBoContrepartie : Boolean = false ) ;
var lNumLigne     : integer ;
    lTobLigne     : Tob ;
    lCredit       : double ;
    lDebit        : double ;
    lInCpt        : integer ;
    lPieceCpt     : TPieceCompta ;
    lPieceDossier : TPieceCompta ;
    lBoMultiSoc   : boolean ;
    lStDossier    : string ;
    lNumLD        : integer ;
    lSolde        : double ;
    lBoRupt       : boolean ;
    lStGen        : string ;
    lStAux        : string ;
    // ==> Crée une nouvelle écriture sur la pièce
    function _AjouteLigne( vPieceCpt : TPieceCompta ; vGene, vAuxi : String ; vIndex : Integer = -1 ) : integer ;
    begin
      // Ajout d'une ligne
      vPieceCpt.NewRecord( vIndex ) ;
      result    := vPieceCpt.CurIdx ;
      // MAJ comptes
      vPieceCpt.Info.LoadCompte( vGene ) ;
      vPieceCpt.PutValue( result, 'E_GENERAL', vGene ) ;
      if vPieceCpt.Info.Compte_GetValue('G_COLLECTIF') = 'X'
        then vPieceCpt.PutValue( result, 'E_AUXILIAIRE', vAuxi ) ;
      // MAJ Infos
      RenseigneEcriture( vTobScenario, lTobLigne, vPieceCpt, result, vBoContrepartie ) ;
    end ;
    // ==> Renseigne l'analytique sur la pièce si option report au détail
    procedure _GereAna ( vPieceCpt : TpieceCompta ; vNumL : integer );
    begin
      if ( vPieceCpt.GetValue( vNumL, 'E_ANA' ) = 'X' ) then
        begin
        if ( vTobScenario.GetValue('CPG_METHODEANA') = 'DET' ) then
          GenereAnalytique ( vTobOrigine, vTobScenario, vInDebut, vInFin, vPieceCpt, vNumL ) ;
        CSynchroVentil( vPieceCpt.GetTob( vNumL ) ) ;
        end ;
    end ;

begin

  if ( vInDebut > vInFin ) or (vTobOrigine.Detail.Count < vInFin ) then
    Exit ;

  // Première ligne de la fourchette des échéances concernées,
  //  utilisée pour récupérer les informations communes
  lTobLigne    := vTobOrigine.Detail[ vInDebut ] ;

  // ---------------------------------------------
  // --  Initialisation des ecritures générées  --
  // ---------------------------------------------

  // Pièce émettrice (base courante)
  lBoMultiSoc  := False ;
  lStDossier   := V_PGI.SchemaName ;
  lPieceCpt    := GetPieceDossier( vTobGene, lStDossier ) ;

  // Recherche du dossier de l'échéance
  if EstMultiSoc and (lTobLigne.GetNumChamp('SYSDOSSIER') > 0) then
    begin
    lStDossier  := lTobLigne.Getstring('SYSDOSSIER') ;
    lBoMultiSoc := lStDossier <> lPieceCpt.Dossier ;
    end ;

  // ------------------------------------------------------
  // -- Génération de l'écriture dans la pièce émettrice --
  // ------------------------------------------------------
  // ==> Report de la contrepartie
  if vBoContrepartie then
    begin

    // === CONTREPARTIE LOCALE ===
    // Création de la ligne
    lNumLigne := _AjouteLigne( lPieceCpt , vTobScenario.GetString('CPG_GENERAL'), lTobLigne.GetString('E_AUXILIAIRE') ) ;

    // === CONTREPARTIE SOCIETE ===
    // si Règlement par société, on ajoute les lignes de règlement soldant les pièces dossiers
    if (vTobScenario.GetValue('CPG_MULTISOCMETH') = 'SOC') and
       not (vTobScenario.GetValue('CPG_GROUPEENCADECA') = 'DET') then
      begin

      // === SI MULTISOC : Insertion des règlements des autres dossiers
      for lInCpt := vTobGene.Detail.Count - 1 downTo 0 do
        begin
        // Recherche de la pièce
        lPieceDossier   := TPieceCompta( vTobGene.Detail[ lInCpt ] ) ;
        if lPieceDossier.Contexte.local then Continue ;
        // Traitement uniquement si quelque chose à solder
        lSolde  := lPieceDossier.GetSolde ;
        if lPieceDossier.GetSolde = 0 then Continue ;
        // Création de la ligne
        lNumLD := _AjouteLigne( lPieceCpt , vTobScenario.GetString('CPG_GENERAL'), lTobLigne.GetString('E_AUXILIAIRE') ) ;

        // libellé spécifique si vide
        SetLibelleMultiSoc( vTobScenario, lTobLigne, lPieceCpt, lNumLD, True, lPieceDossier.Dossier ) ;
        lPieceCpt.PutValue( lNumLD, 'DOSSIERDEST', lPieceDossier.Dossier ) ;

        // Recupération des montants sur les dossier
        if lSolde > 0
          then lPieceCpt.PutValue( lNumLD, 'E_CREDITDEV', lSolde )
          else lPieceCpt.PutValue( lNumLD, 'E_DEBITDEV',  -1 * lSolde ) ;

        // Analytique
        _GereAna( lPieceCpt, lNumLD ) ;

        end ; // for
      end ; //  if vTobScenario.GetValue('CPG_MULTISOCMETH') = 'SOC' then

    // === SUITE CONTREPARTIE LOCALE ===
    // -- MAJ des montants  --
    if lPieceCpt.GetSolde <> 0 then
      begin
      // contrepartie au détail mais multisoc...
      if (vTobScenario.GetValue('CPG_GROUPEENCADECA') = 'DET') and lBoMultiSoc then
        begin
        SetLibelleMultiSoc( vTobScenario, lTobLigne, lPieceCpt, lNumLigne, True, lStDossier ) ;
        lPieceCpt.PutValue( lNumLigne, 'DOSSIERDEST', lTobLigne.GetString('SYSDOSSIER') ) ;
        end
      else lPieceCpt.PutValue( lNumLigne, 'DOSSIERDEST', lPieceCpt.Dossier ) ;
      // Montant
      lPieceCpt.AttribSolde( lNumLigne ) ;
      // Analytique
      _GereAna( lPieceCpt, lNumLigne ) ;
      end
    else lPieceCpt.DeleteRecord( lNumLigne ) ;

    end
  else
  // ==> Report de(s) échéance(s)
    begin

    if lBoMultiSoc then // En multiSoc : compte de liaison
      begin
      lNumLigne := _AjouteLigne( lPieceCpt , GetCompteLiaison( vTobScenario, lStDossier), '' ) ;
      lPieceCpt.PutValue( lNumLigne, 'E_REFINTERNE', '' ) ;
      // libellé spécifique si vide
      SetLibelleMultiSoc( vTobScenario, lTobLigne, lPieceCpt, lNumLigne, False, lStDossier ) ;
      lPieceCpt.PutValue( lNumLigne, 'DOSSIERDEST',  lStDossier ) ;
      lPieceCpt.PutValue( lNumLigne, 'ESTECHEORIG',  '-' ) ;
      end
    else
      begin
      lNumLigne := _AjouteLigne( lPieceCpt , lTobLigne.GetString('E_GENERAL'), lTobLigne.GetString('E_AUXILIAIRE') ) ;
      lPieceCpt.PutValue( lNumLigne, 'DOSSIERDEST',  lPieceCpt.Dossier ) ;
      end ;

    // Re-Calculs des cumuls pour globalisation
    lCredit     := 0 ;
    lDebit      := 0 ;
    for lInCpt := vInDebut to vInFin do
      MajCumulsEncaDeca( vTobOrigine.Detail[ lInCpt ], lDebit, lCredit ) ;
    // -- Attention les montants sont affectés à l'inverse de l'échéance d'origine --
    if lDebit=0
      then lPieceCpt.PutValue( lNumLigne, 'E_DEBITDEV',  lCredit )
      else lPieceCpt.PutValue( lNumLigne, 'E_CREDITDEV', lDebit ) ;
    // Analytique
    _GereAna( lPieceCpt, lNumLigne ) ;

    end ;

  // --------------------------------------------------------------
  // -- Génération des reports sur la société X en multi-société --
  // --------------------------------------------------------------
  // ==> Report des contreparties sur les pièces dossier concernées
  if vBoContrepartie then
    begin
    for lInCpt := 0 to vTobGene.Detail.Count - 1 do
      begin

      // Recherche de la pièce
      lPieceDossier   := TPieceCompta( vTobGene.Detail[ lInCpt ] ) ;

      // Sur la pièce local, pas de traitement à faire
      if lPieceDossier.Contexte.local then Continue ;
      // Sur les autres, uniquement si quelque chose à solder
      if lPieceDossier.GetSolde = 0 then Continue ;

      // Création de la ligne
      lNumLD := _AjouteLigne( lPieceDossier , GetCompteLiaison( vTobScenario, lPieceDossier.Dossier), '' ) ;

      // Référence spécifique
      lPieceDossier.PutValue( lNumLD, 'E_REFINTERNE', '' ) ;
      // libellé spécifique si vide
      SetLibelleMultiSoc( vTobScenario, lTobLigne, lPieceDossier, lNumLD, True ) ;
      // -- MAJ des montants  --
      lPieceDossier.AttribSolde( lNumLD ) ;

      end
    end
    // ==> Report de(s) échéance(s) : 2 cas possible suivant scénario...
  else if lBoMultiSoc then
    begin

    // Recherche de la pièce générée sur le dossier de l'échéance d'origine
    lPieceDossier   := GetPieceDossier( vTobGene, lStDossier ) ;
    if lPieceDossier = nil then
      lPieceDossier := CreerPieceEncaDeca( vTobScenario, lTobLigne, vTobGene, lStDossier ) ;

    // DEBUT FQ
    // 1ère ligne...
    lTobLigne   := vTobOrigine.Detail[ vInDebut ] ;
    lStGen      := lTobLigne.GetString('E_GENERAL') ;
    lStAux      := lTobLigne.GetString('E_AUXILIAIRE') ;
    // init montants
    lCredit     := 0 ;
    lDebit      := 0 ;
    MajCumulsEncaDeca( lTobLigne , lDebit, lCredit ) ;

    for lInCpt := (vInDebut + 1) to vInFin do
      begin
      // Tiers globalisé ou compte Général différent
      lBoRupt   := ( vTobScenario.GetValue('CPG_TIERSGLOBALISE') <> 'X' ) or
                   ( lStGen <> vTobOrigine.Detail[ lInCpt ].GetString('E_GENERAL') ) or
                   ( lStAux <> vTobOrigine.Detail[ lInCpt ].GetString('E_AUXILIAIRE') ) ;

      // Création de la ligne
      if lBoRupt then
        begin
        lNumLD := _AjouteLigne( lPieceDossier , lTobLigne.GetString('E_GENERAL'), lTobLigne.GetString('E_AUXILIAIRE') ) ;
        // -- Attention les montants sont affectés à l'inverse de l'échéance d'origine --
        if lDebit=0
          then lPieceDossier.PutValue( lNumLD, 'E_DEBITDEV',  lCredit )
          else lPieceDossier.PutValue( lNumLD, 'E_CREDITDEV', lDebit ) ;

        // remise à zéro des montant
        lCredit     := 0 ;
        lDebit      := 0 ;

        // MAJ général de référence
        lStGen      := vTobOrigine.Detail[ lInCpt ].GetString('E_GENERAL') ;
        lStAux      := vTobOrigine.Detail[ lInCpt ].GetString('E_AUXILIAIRE') ;

        end ;

      // Maj ligne + cumul
      lTobLigne := vTobOrigine.Detail[ lInCpt ] ;
      MajCumulsEncaDeca( lTobLigne , lDebit, lCredit ) ;

      end ;

    // Mise en place de la dernière rupture (écvetuellement la 1ère aussi si on est pas passé dans la boucle)
    lNumLD := _AjouteLigne( lPieceDossier , lTobLigne.GetString('E_GENERAL'), lTobLigne.GetString('E_AUXILIAIRE') ) ;

    // -- Attention les montants sont affectés à l'inverse de l'échéance d'origine --
    if lDebit=0
      then lPieceDossier.PutValue( lNumLD, 'E_DEBITDEV',  lCredit )
      else lPieceDossier.PutValue( lNumLD, 'E_CREDITDEV', lDebit ) ;

    end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/01/2004
Modifié le ... : 25/06/2007
Description .. : Génère les lignes d'analytique suivant méthode définit dans
Suite ........ : le scénario :
Suite ........ :  - ATT : sur section d'attente,
Suite ........ :  - VEN : selon préventilation du compte
Suite ........ :  - DET : avec report au détail de l'analytique d'origine
Suite ........ : Attention : la méthode DET demande d'avoir charger au
Suite ........ : préalable les lignes d'analytiques d'origine dans la tob des
Suite ........ : échéances d'origine.
Suite ........ : Si vInDebut et vInFin sont différents, alors un regroupement
Suite ........ : à lieu en cumulant des ventilations des lignes d'échéances
Suite ........ : compris entre ces 2 index.
Suite ........ : ---------------------------
Suite ........ : SBO 25/06/2007 : FQ 20657 / 20655
Suite ........ : gestion de report analytique sur les lignes cumulés
Suite ........ : ---------------------------
Mots clefs ... :
*****************************************************************}
Procedure GenereAnalytique  ( vTobOrigine , vTobScenario : TOB ; vInDebut, vInFin : Integer ; vPieceCpt : TPieceCompta ; vNumLigne : Integer ) ;
var lInAxe          : Integer ;
    lInVentil       : Integer ;
    lInCpt          : Integer ;
    lTobAnaGene     : TOB ;
    lTobAnaOrig     : TOB ;
    lTobAxeGene     : TOB ;
    lTobAxeOrig     : TOB ;
    lTobOrig        : TOB ;
    lStSection      : String ;
    lQte            : Double ;
    lPourcentOrig   : Double ;
    lPourcentGene   : Double ;
    lTotalOrig      : Double ;
    lTotalGene      : Double ;
    lCoef           : Double ;
    lTobLigne       : TOB ;
    lMontantEcr     : Double ;
    lMontantDev     : Double ;
    lPourcent       : Double ;
    lDebitAna       : Double ;
    lCreditAna      : Double ;
    lTotalPourcent  : Double ;
    lTotalDebit     : Double ;
    lTotalCredit    : Double ;

    function _CoherAna : boolean ;
    var lStAxeVentilOrig  : string ;
        lStAxeVentilCible : string ;
      begin
      lStAxeVentilOrig  := '' ;
      lStAxeVentilCible := '' ;
      // Cas de résolution automatique
      if (vInDebut <> vInFin) and (vTobScenario.GetValue('CPG_METHODEANA') = 'DET') then
        begin
        if vPieceCpt.Info.LoadCompte( vTobOrigine.Detail[vInDebut].GetString( 'E_GENERAL' ) ) then
          lStAxeVentilOrig := vPieceCpt.Info.GetString('G_VENTILABLE1')
                              + vPieceCpt.Info.GetString('G_VENTILABLE2')
                              + vPieceCpt.Info.GetString('G_VENTILABLE3')
                              + vPieceCpt.Info.GetString('G_VENTILABLE4')
                              + vPieceCpt.Info.GetString('G_VENTILABLE5') ;
        if vPieceCpt.Info.LoadCompte( vPieceCpt.GetString( vNumLigne, 'E_GENERAL' ) ) then
          lStAxeVentilCible := vPieceCpt.Info.GetString('G_VENTILABLE1')
                              + vPieceCpt.Info.GetString('G_VENTILABLE2')
                              + vPieceCpt.Info.GetString('G_VENTILABLE3')
                              + vPieceCpt.Info.GetString('G_VENTILABLE4')
                              + vPieceCpt.Info.GetString('G_VENTILABLE5') ;
        end ;
      // test des 2 valeurs ( vrai si rien a tester puisque les 2 seront à '' )
      result := lStAxeVentilOrig = lStAxeVentilCible ;

      end ;

begin

  if ( vTobScenario.GetValue('CPG_METHODEANA') <> 'DET' ) then Exit ;

  lTobLigne := vPieceCpt.GetTob( vNumLigne ) ;
  lTobLigne.ClearDetail ;
  AlloueAxe( lTobLigne ) ;

  // Gestion analytique sur la ligne générée ?
  if not vPieceCpt.EstVentilable( vNumLigne ) then Exit ;

  // ---------------------------------------------------------------------------------
  // Cas du report de l'analytique au détail de l'échéance d'origine vInDebut à vInFin
  // ---------------------------------------------------------------------------------
  // Vérification de la cohérence de structure analytique des différentes lignes à regrouper
//  if CoherenceAnalytique( vTobOrigine, vTobScenario, vInDebut, vInFin , vPieceCpt.Info ) then
  if _CoherAna then
    begin

    // ===========================================
    // ==== Parcours des échéances concernées ====
    // ===========================================
    for lInCpt := vInDebut to vInFin do
      begin
      lTobOrig := vTobOrigine.Detail[ lInCpt ] ;
      // Parcours des axes
      for lInAxe := 0 to lTobOrig.Detail.Count - 1 do
        begin
        // Pointeurs
        lTobAxeGene := lTobLigne.Detail[ lInAxe ] ;
        lTobAxeOrig := lTobOrig.Detail[ lInAxe ] ;
        // Parcours des ventilations de l'échéance d'origine courante
        for lInVentil := 0 to lTobAxeOrig.Detail.Count-1 do
          begin
          lTobAnaOrig := lTobAxeOrig.Detail[lInVentil] ;
          lStSection  := lTobAnaOrig.getValue('Y_SECTION') ;
          lTobAnaGene := lTobAxeGene.FindFirst(['Y_SECTION'], [lStSection], False) ;
          if lTobAnaGene = nil then
            begin
            lTobAnaGene := TOB.Create('ANALYTIQ', lTobLigne.Detail[lInAxe], -1) ;
            // Initialisations
            CPutDefautAna( lTobAnaGene ) ;
            InitCommunObjAnalNEW( lTobLigne, lTobAnaGene ) ;
            RecopieInfosAnalytique ( lTobAnaOrig , lTobAnaGene ) ;
            lTobAnaGene.PutValue('Y_POURCENTAGE',   0) ;
            lTobAnaGene.PutValue('Y_TOTALECRITURE',  0) ;
            lTobAnaGene.PutValue('Y_TOTALQTE1',  0) ;
            lTobAnaGene.PutValue('Y_TOTALQTE2',  0) ;
            lTobAnaGene.PutValue('Y_QTE1',  0) ;
            lTobAnaGene.PutValue('Y_QTE2',  0) ;
            end ;

          // ----------------------------------------------
          // --    MAJ du pourcentage / montant total    --
          // ----------------------------------------------
          lPourcentOrig := lTobAnaOrig.GetDouble('Y_POURCENTAGE') ;
          lTotalOrig    := lTobAnaOrig.GetDouble('Y_TOTALECRITURE') ;
          lPourcentGene := lTobAnaGene.GetDouble('Y_POURCENTAGE') ;
//          lTotalGene    := lTobAnaGene.GetValue('Y_TOTALECRITURE') ;
          lTotalGene    := lTobLigne.GetDouble('E_DEBIT') + lTobLigne.GetDouble('E_CREDIT') ;

          // Coefficient pour gestion du sens des écritures
          if ( ( lTobLigne.GetDouble('E_DEBIT') <> 0 ) and ( lTobAnaOrig.GetDouble('Y_CREDIT') <> 0 ) )
             or ( ( lTobLigne.GetDouble('E_CREDIT') <> 0 ) and ( lTobAnaOrig.GetDouble('Y_DEBIT') <> 0 ) )
             then lCoef := -1.0
             else lCoef := 1.0 ;

          // Formule pour récupérer le nouveau poucentage :
          //  ( Old_Taux * Old_Total ) + ( coef * new_taux * new_total )
          //  ----------------------------------------------------------
          //               ( Old_total + New_Total )

          if (lTotalOrig + lTotalGene) <> 0 // Test pour éviter division par 0
//            then lPourcentGene := ( ( lPourcentOrig * lTotalOrig ) + ( lCoef * lPourcentGene * lTotalGene ) )
//                                  / ( lTotalOrig + lTotalGene )
            then lPourcentGene := lPourcentGene + lCoef * ( ( lPourcentOrig * lTotalOrig ) / lTotalGene )
            else lPourcentGene := 0 ;
          lPourcentGene := Arrondi( lPourcentGene , ADecimP ) ;
          lTobAnaGene.PutValue('Y_POURCENTAGE', lPourcentGene ) ;

          // MAJ du total des écritures
          lTobAnaGene.PutValue('Y_TOTALECRITURE', lTotalGene ) ;

          // Affectation des Qtes
          lQte := lTobAnaGene.GetValue('Y_QTE1') ;
          lQte := lQte + lTobAnaOrig.GetValue('Y_QTE1') ;
          lTobAnaGene.PutValue('Y_QTE1',      lQte) ;
          lQte := lTobAnaGene.GetValue('Y_QTE2') ;
          lQte := lQte + lTobAnaOrig.GetValue('Y_QTE2') ;
          lTobAnaGene.PutValue('Y_QTE2',      lQte) ;
          lQte := lTobAnaGene.GetValue('Y_QTE1') ;
          lQte := lQte + lTobAnaOrig.GetValue('Y_QTE1') ;
          lTobAnaGene.PutValue('Y_TOTALQTE1', lQte) ;
          lQte := lTobAnaGene.GetValue('Y_TOTALQTE2') ;
          lQte := lQte + lTobAnaOrig.GetValue('Y_TOTALQTE2') ;
          lTobAnaGene.PutValue('Y_TOTALQTE2', lQte) ;

          end ;
        end ;

      end ;

    // ==================
    // ==== FINITION ====
    // ==================
    // remise en place du signe correcte pour le pourcentage si besoin
    for lInAxe:=0 to lTobLigne.Detail.Count - 1 do
      begin
      lTotalPourcent    := 0 ;
      lTobAxeGene := lTobLigne.Detail[lInAxe] ;
      // recherche total pourcentage
      for lInVentil:=0 to lTobAxeGene.Detail.Count-1 do
        lTotalPourcent := lTotalPourcent + lTobAxeGene.Detail[lInVentil].GetDouble('Y_POURCENTAGE') ;
      // inversion du signe
      if lTotalPourcent < 0 then
        for lInVentil:=0 to lTobAxeGene.Detail.Count-1 do
          lTobAxeGene.Detail[lInVentil].putValue('Y_POURCENTAGE', -1 * lTobAxeGene.Detail[lInVentil].GetDouble('Y_POURCENTAGE') ) ;
      end ;

    // Parcours des axes de l'écriture en cours
    for lInAxe:=0 to lTobLigne.Detail.Count - 1 do
      begin
      // Initialisation des cumuls
      lTotalPourcent    := 0 ;
      lTotalDebit       := 0 ;
      lTotalCredit      := 0 ;

      // Parcours des ventilations de chaque axe
      lTobAxeGene := lTobLigne.Detail[lInAxe] ;
      for lInVentil:=0 to lTobAxeGene.Detail.Count-1 do
        begin
        lTobAnaGene := lTobAxeGene.Detail[lInVentil] ;

        // Numéro de ligne
        lTobAnaGene.PutValue('Y_NUMVENTIL',      IntToStr(lInVentil+1)) ;

        // Montant total écriture générale
        lMontantEcr   := lTobLigne.GetValue('E_DEBIT') + lTobLigne.GetValue('E_CREDIT') ;
        lTobAnaGene.PutValue('Y_TOTALECRITURE', lMontantEcr ) ;
        lMontantDev   := lTobLigne.GetValue('E_DEBITDEV') + lTobLigne.GetValue('E_CREDITDEV') ;
        lTobAnaGene.PutValue('Y_TOTALDEVISE',   lMontantDev ) ;

        // Finition des calculs si report de l'analytique au détail de l'échéance
        // Sauf pour la dernière ligne qui est calculé par différence total - cumul
        if (lInVentil < lTobAxeGene.Detail.Count - 1 ) then
          begin
          lDebitAna      := Arrondi( lTobLigne.GetValue('E_DEBITDEV')  * lTobAnaGene.GetValue('Y_POURCENTAGE') / 100.0, vPieceCPT.Devise.Decimale ) ;
          lCreditAna     := Arrondi( lTobLigne.GetValue('E_CREDITDEV') * lTobAnaGene.GetValue('Y_POURCENTAGE') / 100.0, vPieceCPT.Devise.Decimale ) ;
          // MAJ montant tob
          CSetMontants( lTobAnaGene, lDebitAna, lCreditAna, vPieceCpt.Devise, True );
          // stockage du cumul pour calcul exact de la dernière ligne
          lTotalPourcent := lTotalPourcent + lTobAnaGene.GetValue('Y_POURCENTAGE') ;
          lTotalDebit    := lTotalDebit + lTobAnaGene.GetValue('Y_DEBITDEV') ;
          lTotalCredit   := lTotalCredit + lTobAnaGene.GetValue('Y_CREDITDEV') ;
          end
        else
          begin
          lDebitAna      := lTobLigne.GetValue('E_DEBITDEV') - lTotalDebit ;
          lCreditAna     := lTobLigne.GetValue('E_CREDITDEV') - lTotalCredit ;
          lPourcent      := 100 - lTotalPourcent ;
          // MAJ montant tob
          CSetMontants( lTobAnaGene, lDebitAna, lCreditAna, vPieceCpt.Devise, True );
          // MAJ pourcentage
          lTobAnaGene.PutValue('Y_POURCENTAGE', lPourcent ) ;
          end ;

        // Calcul pourcentage des quantités
        lPourcent := 0 ;
        if lTobAnaGene.GetValue('Y_TOTALQTE1')<>0 then
          lPourcent := Arrondi( 100.0*( lTobAnaGene.GetValue('Y_QTE1') ) / lTobAnaGene.GetValue('Y_TOTALQTE1'), V_PGI.OkDecQ ) ;
        lTobAnaGene.PutValue('Y_POURCENTQTE1',    lPourcent) ;
        lPourcent := 0 ;
        if lTobAnaGene.GetValue('Y_TOTALQTE2')<>0 then
          lPourcent := Arrondi( 100.0*( lTobAnaGene.GetValue('Y_QTE2') ) / lTobAnaGene.GetValue('Y_TOTALQTE2'), V_PGI.OkDecQ ) ;
        lTobAnaGene.PutValue('Y_POURCENTQTE2',    lPourcent) ;

        end ;

      end ;

    end
  else
   // Si recopie impossible de l'analytique, ventilation par défaut
   VentilerTOB( lTobLigne, '', 0, vPieceCpt.Devise.Decimale, False ) ;

end ;


Function  GenereExports ( vTLEche : TList ; vTobScenario : TOB ; vDev : RDevise ) : Boolean ;
var
    lErr    : TMC_ERR ;
    lTobExp : TOB ;
    lTobEcr : TOB ;
    i       : Integer ;
    lInCpt  : Integer ;
    lStMod  : String ;
    lStCle     : String ;
    lStRIB     : String ;
    lStCleIBAN : String ;
    lStPays    : String ;
    lStEtab    : String ;
    lStGuichet : String ;
    lStBidon   : String ;
begin
  Result:=FALSE ;
  if ( vTobScenario.GetValue('CPG_CFONBEXPORT') <> 'X' ) or
     ( Trim(vTobScenario.GetValue('EXPORTFICHIER')) = '' ) then Exit ;

  lInCpt  := 0 ;
  lTobExp := TOB.Create('EXPORTCFONB', nil, -1) ;
  for i := 0 to vTLEche.Count - 1 do
    if TOB(vTLEche[i]).GetValue('ESTECHEORIG') = 'X' then
      begin
      Inc( lInCpt ) ;
      lTobEcr := Tob.Create('ECRITURE', lTobExp, -1) ;
      lTobEcr.Dupliquer( TOB(vTLEche[i]), False, True) ;
      lTobEcr.AddChampSupValeur('REMISEREF',       vTobScenario.GetValue('REMISEREF') ) ;
      lTobEcr.AddChampSupValeur('REMISEDATE',      vTobScenario.GetValue('REMISEDATE') ) ;
      lTobEcr.AddChampSupValeur('FRAISIMPUTATION', vTobScenario.GetValue('FRAISIMPUTATION') ) ;
      lTobEcr.AddChampSupValeur('CFONBMONTANT',    (lTobEcr.GetValue('E_DEBITDEV')+lTobEcr.GetValue('E_CREDITDEV'))* (IntPower(10,vDev.Decimale) ) ) ;
      lTobEcr.AddChampSupValeur('CFONBDECIMAL',    IntToStr( vDev.Decimale ) ) ;
      lTobEcr.AddChampSupValeur('NUMLIGNE',        IntToStr( lInCpt ) ) ;
      if VH^.PaysLocalisation=CodeISOES then
      begin
         lTobEcr.AddChampSupValeur('IDCEDENTE',       vTobScenario.GetValue('IDCEDENTE')) ;
         lTobEcr.AddChampSupValeur('EFFETSPHYS',      vTobScenario.GetValue('EFFETSPHYS')) ;
         lTobEcr.AddChampSup('TYPDOC', FALSE) ;
      End ;//XVI 24/02/2005
      end ;

  if lTobExp.Detail.Count > 0 then
    Begin
    lStMod:=vTobScenario.GetValue('CPG_CFONBFORMAT') ;
    if VH^.PaysLocalisation=CodeISOEs then
    Begin
       lTobExp.Detail[0].Addchampsup('TRIPAR',TRUE) ;
       if vTobScenario.Detail.count>0 then
          lTobExp.Detail[0].AddchampsupValeur('LABANQUE',vTobScenario.Detail[0].GetValue('CCB_GENERAL'),TRUE) ;
       for i:=0 to lTobExp.Detail.Count-1 do
           Begin
           lstCle:='' ;
           if pos(';'+lstMod+';',';N19;N58;')>0 then
              Begin //Ordre de Tri RIB.EtabBQ + RIB.Guichet + Référence + code enregistrement (géré en interne)
              if IBANtoRIB(lTobExp.Detail[i].GetValue('E_RIB'),lStRib,lStPays,lStCleIBAN) then
                 Begin
                 DecodeRib(lstEtab,lstGuichet,lstBidon,lstBidon,lstBidon,lStRIB,lstPays) ;
                 lStCle:=lStEtab+lStGuichet+lTobExp.Detail[i].GetValue('E_REFGESCOM')+formatfloat('000',lTobExp.Detail[i].GetValue('E_NUMECHE')) ;
                 End ;
              End
           else if lStMod='N32' then
                   lStCle:=lTobExp.Detail[i].GetValue('NUMLIGNE') ;
           lTobExp.Detail[i].PutValue('TRIPAR',lStCle) ;
           //Charger les infos manquantes......
           //L'auxiliaire
           ChargeComplementExport(lTobExp.Detail[i],'TIERS','E_AUXILIAIRE','T_AUXILIAIRE') ;
           //Modepaiement
           ChargeComplementExport(lTobExp.Detail[i],'MODEPAIE','E_MODEPAIE','MP_MODEPAIE') ;
           //Rib Banque Remise....
           if vTobScenario.Detail.count>0 then
              ChargeComplementExport(lTobExp.Detail[i],'BANQUECP','LABANQUE','BQ_GENERAL') ;
                 //CPG_IDCEDENTE et CPG_EFFETSPHY.....
                 //Si  E_MODEPAIE.Categorie=TEP, alors TypeDoc=3, sinon TypeDoc=2 (il y a un souci, on ne  peut pas reconnaitre les traites des LCR)
           End ;

       lTobEXP.Detail.Sort('TRIPAR') ;
    End ; //XVI 24/02/2005
    Result := ExporteLP( 'F', 'FIC', lstMod,
                         '',
//                         'ESTECHEORIG="X"',                                   // Clause Where
                         vTobScenario.GetValue('EXPORTFICHIER'),                // Fichiers
                         vTobScenario.GetValue('EXPORTAPERCU')='X',             // Aperçu
                         False,                                                 // Sur base ?
                         lTobExp,                                               // TOB
                         lErr ) ;                                               // TMC_ERR ???
    End ;
//XMG 20/04/04 fin

end ;

Function  CoherenceAnalytique ( vTobOrigine, vTobScenario : TOB ; vInDebut, vInFin : Integer ; vInfoEcr : TInfoEcriture ) : Boolean ;
var lTobEche      : TOB ;
    lTobAxe       : TOB ;
    lInCpt        : Integer ;
    lInAxe        : Integer ;
    lTabVentil    : Array[1..MaxAxe] of Boolean ;
begin
  Result := False ;

  // Cas de résolution automatique
  if (vInDebut = vInFin) or (vTobScenario.GetValue('CPG_METHODEANA') <> 'DET') then
    begin
    Result := True ;
    Exit ;
    end ;

  // Parcours de la liste des lignes concernés par le
  for lInCpt := vInDebut to vInFin do
    begin
    lTobEche := vTobOrigine.Detail[ lInCpt ] ;
    FillChar( lTabVentil, Sizeof(lTabVentil), #0 ) ;
    // Parcours des ventilations concernées pour connaître le liste des axes ventilés
    for lInAxe := 0 to (lTobEche.Detail.Count - 1) do
        begin
        lTobAxe := lTobEche.Detail[ lInAxe ] ;
        if lTobAxe.Detail.Count > 0 then
          lTabVentil[ lInAxe + 1 ] := True ;
        end ;
    // Teste si les axes ventilés sont aussi les axes ventilables du compte de génération
    for lInAxe := 1 to MaxAxe do
      if lTabVentil[lInAxe] <> ( vInfoEcr.GetString('G_VENTILABLE' + IntToStr(lInAxe)) = 'X' )
        then Exit ;
    end ;

  // Si on a passé tout les tests, alors cohérence il y a.
  Result := True ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/01/2004
Modifié le ... :   /  /    
Description .. : Complète la tob contenant les infos du scenario par des
Suite ........ : infos supplémentaires utilisé durant le traitement des
Suite ........ : enca/deca
Mots clefs ... :
*****************************************************************}
Procedure CompleteScenario( TobScenario : TOB ) ;
var lQJal     : TQuery ;
    lStNatJal : String ;
begin

  // Code journal d'escompte
  if TobScenario.GetValue('CPG_ESCMETHODE')<>'RIE'
     then TobScenario.PutValue('ESCJOURNAL', GetParamSocSecur('SO_CPJALESCOMPTE', '') )
     else TobScenario.PutValue('ESCJOURNAL', '' ) ;

  // Recherche nature pièce à générer
  lStNatJal := 'OD' ;
  lQJal := OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="' + TobScenario.GetValue('CPG_JOURNAL')  + '" ' , TRUE ) ;
  If Not lQJal.Eof Then
    lStNatJal := lQJal.FindField('J_NATUREJAL').AsString ;
  Ferme( lQJal ) ;
  if ( lStNatJal = 'BQE' ) or ( lStNatJal = 'CAI' ) then
    begin
    if TobScenario.GetValue('CPG_FLUXENCADECA') = 'ENC'
      then TobScenario.PutValue( 'NATUREPIECE' ,  'RC' )
      else TobScenario.PutValue( 'NATUREPIECE' ,  'RF' ) ;
    end
  else
    TobScenario.PutValue( 'NATUREPIECE' ,  'OD' ) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/01/2004
Modifié le ... :   /  /
Description .. : Enregistre en base de données la pièce d'enca/déca
Suite ........ : préalablement générée dans la fonction GenereEncaDeca.
Suite ........ : L'analytique est enregistrée en même temps, la tob ayant
Suite ........ : l'arborescence suivante :
Suite ........ : ECRITURE --> AXE (virtuelle) --> ANALYTIQ
Suite ........ :
Mots clefs ... :
*****************************************************************}
Function  SauveEncaDeca( vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais, vTobScenario : TOB ) : Boolean ;
var lInCpt      : integer ;
    lPieceCpt   : TPieceCompta ;
    lPieceLocal : TPieceCompta ;
    lTobLiens   : Tob ;
begin

  // Par default, on considère qu'il peut y avoir erreur
  Result := False ;

  try

    // Enregistrement des pièces d'enca-déca
    lPieceLocal := TPieceCompta( vTobGene.Detail[ 0 ] ) ;
    for lInCpt := 0 to vTobGene.Detail.Count - 1 do
      begin
      lPieceCpt := TPieceCompta( vTobGene.Detail[ lInCpt ] ) ;
      if not lPieceCpt.Save then
        begin
        V_PGI.IoError := oeSaisie ;
        Exit ;
        end ;

      // Enregistrement liaison multi-soc
      if not lPieceCpt.Contexte.local then
        begin
        lTobLiens := CCreerLiensSoc( lPieceLocal,
                                     lPieceCpt,
                                     GetCompteLiaison( vTobScenario, lPieceCpt.Dossier ),
                                     Copy('Reglt société ' + GetLiaisonSoc( vTobScenario, lPieceCpt.Dossier), 1, 35)
                                     ) ; // function CCreerLiensSoc( vPieceLocal, vPieceDossier : TPieceCompta ; vStGene, vStLib : string ) : Tob ;
        lTobLiens.InsertDB(nil) ;
        lTobLiens.Free ;
        end ;

      end ;

    // Enregistrement des pièces d'escompte
    if vTobEscompte.Detail.Count > 0 then
      for lInCpt := 0 to vTobEscompte.Detail.Count - 1 do
        begin
        lPieceCpt := TPieceCompta( vTobEscompte.Detail[ lInCpt ] ) ;
        if not lPieceCpt.Save then
          begin
          V_PGI.IoError := oeSaisie ;
          Exit ;
          end ;
        end ;

    // Enregistrement de la deuxième pièce des remises en Escompte
    if (VH^.PaysLocalisation=CodeISOES) and (vTobRemEscompte.Detail.Count>0) then
       begin
       // Enregistrement base
       vTobRemEscompte.InsertDB( nil, True ) ;
        // MAJ des comptes
        MajSoldesEcritureTOB( vTobRemEscompte, False ) ;
        if V_PGI.IoError<>oeOk then Exit ;
       end ;

    // Enregistrement de la pièce de frais
    if (VH^.PaysLocalisation=CodeISOES) and (vTobFrais.Detail.Count>0) then
       begin
       for lInCpt := 0 to vTobFrais.Detail.Count - 1 do
           begin
           // Enregistrement base
           vTobFrais.Detail[lInCPT].InsertDB( nil, True ) ;
          // MAJ des comptes
           MajSoldesEcritureTOB( vTobFrais.Detail[lInCPT], False ) ;
           if V_PGI.IoError<>oeOk then Exit ;
           End ;
       end ; //XVI 24/02/2005

    // Tout s'est bien pasé si on arrive là :
    Result := True ;

    finally

  end;

end ;

Procedure AjouteChampsSuppScenario ( TobScenario : TOB ) ;
begin

  // ==============================================
  // ==== Infos remplis depuis le multicritere ====
  // ==============================================

  // MUL : Paramètres contextuelles
  TobScenario.AddChampSup('DATECOMPTABLE',       False ) ;
  TobScenario.AddChampSup('FORCERDATEECHE',      False ) ;
  TobScenario.AddChampSup('DATEECHE',            False ) ;
  TobScenario.AddChampSup('ETABLISSEMENTPAYEUR', False ) ;

  // Gestion du E_NUMTRAITECHQ
  TobScenario.AddChampSup('NUMCHEQUE',           False ) ;
  TobScenario.AddChampSup('REFCHEQUE',           False ) ;
  TobScenario.AddChampSup('NOMPARAMSOCTRAITE',   False ) ;

  // Uniquement Tic/Tid
  TobScenario.addChampSup('TICTID', False) ;

  // MUL : Paramètres d'éditions
  TobScenario.AddChampSup('APERCU',          False ) ;
  TobScenario.AddChampSup('SPOOLER',         False ) ;
  TobScenario.AddChampSup('XFICHIERSPOOLER', False ) ;
  TobScenario.AddChampSup('REPSPOOLER',      False ) ;
  TobScenario.AddChampSup('NUMENCADECA',     False ) ;

  // MUL : Paramètres d'export
  TobScenario.AddChampSup('EXPORTFICHIER',   False ) ;
  TobScenario.AddChampSup('EXPORTAPERCU',    False ) ;
  TobScenario.AddChampSup('REMISEREF',       False ) ;
  TobScenario.AddChampSup('REMISEDATE',      False ) ;
  TobScenario.AddChampSup('FRAISIMPUTATION', False ) ;
  if VH^.PaysLocalisation=CodeISOEs then
  begin
     TobScenario.AddChampSup('IDCEDENTE'       ,FALSE) ;
     TobScenario.AddChampSup('EFFETSPHYS'      ,FALSE) ;
  end ; //XVI 24/02/2005
  // ==================================================
  // ==== Infos remplis chargées pour l'algorithme ====
  // ==================================================

  TobScenario.AddChampSup( 'NATUREPIECE' ,  False ) ;

  // Journal d'escompte
  TobScenario.AddChampSup('ESCJOURNAL', False ) ;

  // Gestion de l'uniformité du mode de paiement
  TobScenario.AddChampSup('RuptureModePaie',False) ;

  // Nb d'échéance max à traiter par pièces
  TobScenario.AddChampSup('NBLIGNES', False ) ;

  // GEstion global du lettrage pour découpage par lot avec compte non lettrable
  TobScenario.AddChampSup('ENCADECABATCH',    False ) ;
  TobScenario.AddChampSup('CFONBBATCH',       False ) ;
  TobScenario.AddChampSup('NOMLOT',           False ) ;
  TobScenario.AddChampSup('SANSLETTRAGE',     False) ;
  TobScenario.AddChampSup('CFONBPREMIER',     False) ;
  TobScenario.AddChampSup('TRESOSYNCHRO',     False) ;

  // Valeurs par défaut
  TobScenario.PutValue   ('CFONBBATCH',       '-' ) ;
  TobScenario.PutValue   ('NBLIGNES',         0   ) ;
  TobScenario.PutValue   ('SANSLETTRAGE',     '-' ) ;
  TobScenario.PutValue   ('ENCADECABATCH',    '-' ) ;
  TobScenario.PutValue   ('NOMLOT',           ''  ) ;
  TobScenario.PutValue   ('CFONBPREMIER',     'X' ) ;
  TobScenario.PutValue   ('TRESOSYNCHRO',     ets_Rien ) ;

end ;

Procedure LettrerEncaDeca ( vTobOrig : TOB ; vPieceCpt : TPieceCompta ; vDev : RDevise ; vEtab : string = '' ) ;
Var i,j,k               : Integer ;
    LOrig,LSais,LEcr    : TList ;
    XOrig,XSais,XOrig2  : TL_Rappro ;
    FindC,OkMontants    : boolean ;
    lTobLigne           : TOB ;
//    lPieceEtab          : TPieceCompta ;
Begin

  if (vTobOrig.Detail.Count = 0) or (vPieceCpt.Detail.Count = 0) then
    Exit ;

  if ( vEtab <> '' ) and ( vEtab <> vPieceCpt.GetEntete('E_ETABLISSEMENT') ) then
    Exit ;


  LOrig := TList.Create ;
  LSais := TList.Create ;
  LEcr  := TList.Create ;

  // Ecriture générée
  For i := 0 to vPieceCpt.Detail.Count - 1 do
    begin
    lTobLigne := vPieceCpt.Detail[ i ] ;
    if lTobLigne.GetValue('ESTECHEORIG')='X' then
       begin
       XSais         := GetRapproTob( lTobLigne, vDev ) ;
       XSais.Dossier := vPieceCpt.Dossier ;
       LSais.Add( XSais ) ;
       end ;
    end ;

  // Multi-établissement : Récupération des règlements des autres établissements
{
  For i := 0 to ( vPieceCpt.PiecesEtab.Count - 1 ) do
    begin
    lPieceEtab := TPieceCompta( vPieceCpt.PiecesEtab.Objects[ i ] ) ;
    For j := 0 to lPieceEtab.Detail.Count - 1 do
      begin
      lTobLigne := lPieceEtab.Detail[ j ] ;
      if lTobLigne.GetValue('ESTECHEORIG')='X' then
         begin
         XSais         := GetRapproTob( lTobLigne, vDev ) ;
         XSais.Dossier := lPieceEtab.Dossier ;
         LSais.Add( XSais ) ;
         end ;
      end ;
    end ;
}
  // Echéances origine
  for i:=0 to vTobOrig.Detail.Count - 1 do
    begin
    lTobLigne := vTobOrig.Detail[ i ] ;
    if ( lTobLigne.GetNumChamp('SYSDOSSIER') > 1 ) and
       ( lTobLigne.GetString('SYSDOSSIER') <> vPieceCpt.Dossier ) then Continue ;
    if ( vEtab <> '' ) and ( vEtab <> lTobLigne.GetString('E_ETABLISSEMENT') ) then Continue ;
    XOrig         := GetRapproTob( lTobLigne, vDev ) ;
    XOrig.Dossier := vPieceCpt.Dossier ;
    LOrig.Add(XOrig) ;
    end ;

  // Lecture des paquets pour partiellement lettrés
  for i:=LOrig.Count-1 downto 0 do
    BEGIN
    XOrig:=TL_Rappro(LOrig[i]) ;
    if ((XOrig.CodeL<>'') and (DansSaisie(XOrig,LSais))) then
      ChargeLettrage(XOrig,LOrig,vDev) ;
    END ;

  // Lecture des origines passage 1 pour 1
  for i:=0 to LOrig.Count-1 do
    BEGIN
    XOrig:=TL_Rappro(LOrig[i]) ;
//    FindC:=False ;
    for j:=0 to LSais.Count-1 do
        BEGIN
        XSais:=TL_Rappro(LSais[j]) ;
        if XSais.Solution=1 then Continue ;
        if    ( (XSais.General=XOrig.General) and ( Length( XSais.General ) <> GetInfoCpta( fbAux ).Lg ) and ( Length( XOrig.General ) <> GetInfoCpta( fbAux ).Lg ) )
           or ( (XSais.General=XOrig.General) and (XSais.Auxiliaire=XOrig.Auxiliaire)) then
           BEGIN
//           FindC:=True ; {si même comptes --> tester les montants}
           if vDev.Code<>V_PGI.DevisePivot then
              BEGIN
              OkMontants:=((XSais.Debit=XOrig.Credit) and (XSais.Credit=XOrig.Debit)) ;
              END else
              BEGIN
              OkMontants:=((XSais.DebDev=XOrig.CredDev) and (XSais.CredDev=XOrig.DebDev)) ;
              END ;
           if OkMontants then
              BEGIN
              LEcr.Clear ; LEcr.Add(XOrig) ; LEcr.Add(XSais) ;
              LettrerUnPaquet(LEcr,False,True) ; {si ok montant --> lettrage 1 pour 1}
              XOrig.Solution:=1 ; {ne pas les reprendre}
              XSais.Solution:=1 ;
              Break ;
              END ;
           END ;
        END ;
// Fiche 12017 : Optimisation de trop pour les lignes crédit / débit de même montant
//                car a pour effet de zapper la lecture des origines sur elles-même
//                (dernière étape plus bas)
//    if Not FindC then XOrig.Solution:=1 ; {inutile de rechercher si cptes pas présents saisie}
    END ;

  // Lecture des origines passage n pour p
  for i:=0 to LOrig.Count-1 do
    BEGIN
    XOrig:=TL_Rappro(LOrig[i]) ; FindC:=False ; LEcr.Clear ;
    if XOrig.Solution=1 then Continue ;
    {Lecture des echés de saisie (normalement doivent exister) }
    for j:=0 to LSais.Count-1 do
        BEGIN
        XSais:=TL_Rappro(LSais[j]) ; if XSais.Solution=1 then Continue ;
        if    ( (XSais.General=XOrig.General) and ( Length( XSais.General ) <> GetInfoCpta( fbAux ).Lg ) and ( Length( XOrig.General ) <> GetInfoCpta( fbAux ).Lg ) )
           or ( (XSais.General=XOrig.General) and (XSais.Auxiliaire=XOrig.Auxiliaire)) then
           BEGIN
           FindC:=True ; LEcr.Add(XSais) ; XSais.Solution:=1 ;
           END ;
        END ;
    {Si trouvé des echés saisies, rechercher les origines sur les mêmes comptes}
    if FindC then
       BEGIN
       LEcr.Add(XOrig) ; XOrig.Solution:=1 ;
       for k:=i+1 to LOrig.Count-1 do
           BEGIN
           XOrig2:=TL_Rappro(LOrig[k]) ; if XOrig2.Solution=1 then Continue ;
           if    ( (XOrig2.General=XOrig.General) and ( Length( XOrig2.General ) <> GetInfoCpta( fbAux ).Lg ) and ( Length( XOrig2.General ) <> GetInfoCpta( fbAux ).Lg ) )
              or ( (XOrig2.General=XOrig.General) and (XOrig2.Auxiliaire=XOrig.Auxiliaire)) then
              BEGIN
              LEcr.Add(XOrig2) ; XOrig2.Solution:=1 ;
              END ;
           END ;
       END ;
    if LEcr.Count>0 then LettrerUnPaquet(LEcr,False,True) ;
    END ;

  // Lecture des origines sur elles-même
  for i:=0 to LOrig.Count-1 do
    BEGIN
    XOrig:=TL_Rappro(LOrig[i]) ; FindC:=False ; LEcr.Clear ;
    if XOrig.Solution=1 then Continue ;
    {Lecture des autres origines}
    for j:=i+1 to LOrig.Count-1 do
        BEGIN
        XOrig2:=TL_Rappro(LOrig[j]) ; if XOrig2.Solution=1 then Continue ;
        if    ( (XOrig2.General=XOrig.General) and ( Length( XOrig2.General ) <> GetInfoCpta( fbAux ).Lg ) and ( Length( XOrig2.General ) <> GetInfoCpta( fbAux ).Lg ) )
           or ( (XOrig2.General=XOrig.General) and (XOrig2.Auxiliaire=XOrig.Auxiliaire)) then
           BEGIN
           LEcr.Add(XOrig2) ; XOrig2.Solution:=1 ; FindC:=True ;
           END ;
        END ;
    if FindC then
       BEGIN
       LEcr.Add(XOrig) ; XOrig.Solution:=1 ;
       LettrerUnPaquet(LEcr,False,True) ;
       END ;
    END ;

  // Frees
  VideListe(LOrig) ;
  LOrig.Free ;
  VideListe(LSais) ;
  LSais.Free ;
  LEcr.Clear ;
  LEcr.Free ;
end ;

Function GetRapproTob ( vTobLigne : TOB ; vDev : RDEVISE ) : TL_Rappro ;
begin
  Result := TL_Rappro.Create ;

  Result.General      := vTobLigne.GetString('E_GENERAL') ;
  Result.Auxiliaire   := vTobLigne.GetString('E_AUXILIAIRE') ;
  Result.Exo          := vTobLigne.GetString('E_EXERCICE') ;
  Result.DateC        := vTobLigne.GetValue('E_DATECOMPTABLE') ;
  Result.DateE        := vTobLigne.GetValue('E_DATEECHEANCE') ;
  Result.DateR        := vTobLigne.GetValue('E_DATEECHEANCE') ; {JLD}
  Result.RefI         := vTobLigne.GetString('E_REFINTERNE') ;
  Result.RefL         := vTobLigne.GetString('E_REFLIBRE') ;
  Result.RefE         := vTobLigne.GetString('E_REFEXTERNE') ;
  Result.Lib          := vTobLigne.GetString('E_LIBELLE') ;
  Result.Jal          := vTobLigne.GetString('E_JOURNAL') ;
  Result.Numero       := vTobLigne.GetValue('E_NUMEROPIECE') ;
  Result.NumLigne     := vTobLigne.GetValue('E_NUMLIGNE') ;
  Result.NumEche      := vTobLigne.GetValue('E_NUMECHE') ;
  Result.CodeD        := vTobLigne.GetString('E_DEVISE') ;
  Result.TauxDev      := vTobLigne.GetValue('E_TAUXDEV') ;
  Result.CodeL        := vTobLigne.GetString('E_LETTRAGE') ;
  Result.Decim        := vDev.Decimale ;
  // Attention, les 4 lignes après ne sont pas une erreur
  Result.Debit        := vTobLigne.GetValue('E_DEBITDEV') ;
  Result.Credit       := vTobLigne.GetValue('E_CREDITDEV') ;
  Result.DebDev       := vTobLigne.GetValue('E_DEBIT') ;
  Result.CredDev      := vTobLigne.GetValue('E_CREDIT') ;

  Result.DebitCur     := vTobLigne.GetValue('E_DEBIT') ;
  Result.CreditCur    := vTobLigne.GetValue('E_CREDIT') ;
  Result.Nature       := vTobLigne.GetString('E_NATUREPIECE') ;
  Result.EditeEtatTva := vTobLigne.GetValue('E_EDITEETATTVA') = 'X' ;
  Result.Facture      := (Result.Nature='FC') or (Result.Nature='FF') or (Result.Nature='AC') or (Result.Nature='AF') ;
  Result.Client       := (Result.Nature='FC') or (Result.Nature='AC') or (Result.Nature='RC') or (Result.Nature='OC') ;
  Result.Solution     := 0 ;

  Result.NumTraCHQ    := vTobLigne.GetString('E_NUMTRAITECHQ') ;

{
  if vTobLigne.GetNumChamp( 'SYSDOSSIER' ) > 1
    then Result.Dossier := vTobLigne.GetString('SYSDOSSIER')
    else Result.Dossier := V_PGI.SchemaName ;
}
end ;

Function  EncodeTOBEcriture ( vTobEcr : TOB ) : String ;
begin
  Result := vTobEcr.GetValue('E_JOURNAL')  + ';'
          + vTobEcr.GetValue('E_EXERCICE') + ';'
          + UsDateTime( vTobEcr.GetValue('E_DATECOMPTABLE') ) + ';'
          + IntToStr( vTobEcr.GetValue('E_NUMEROPIECE') ) + ';'
          + IntToStr( vTobEcr.GetValue('E_NUMLIGNE') ) + ';'
          + IntToStr( vTobEcr.GetValue('E_NUMECHE') ) ;
end;



procedure _permuteTob ( vTLEche, vTLParent : TList ; vTobParent : TOB ) ;
var j : integer ;
begin
  for j := 0 to vTLEche.count - 1 do
    begin
    if Assigned( vTLParent ) then
      vTLParent.add ( TOB( vTLEche[j] ).Parent ) ;
    TOB( vTLEche[j] ).ChangeParent( vTobParent, -1 ) ;
    end ;
end ;

procedure _restaureTob ( vTLEche, vTLParent : TList ) ;
var j : integer ;
begin
  for j := 0 to vTLEche.count - 1 do
    if Assigned( vTLParent )
      then TOB( vTLEche[j] ).ChangeParent( Tob( vTLParent[j]) , -1 )
      else TOB( vTLEche[j] ).ChangeParent( nil , -1 ) ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Genère les états / documents en fonction du paramètrage
Suite ........ : de vTobScenario, pour les règlements de vTobGene
Mots clefs ... :
*****************************************************************}
function GenereEtats ( vTLEche : TList ; vTobScenario : TOB ) : boolean ;
var lStNature   : String ;
    lStCode     : String ;
    lStSpooler  : String ;
    lApercu     : Boolean ;
    lTobEtat    : Tob ;
    lTobLigne   : Tob ;
    lTobRupt    : Tob ;
    i           : Integer ;
    lStRuptAuxi : String ;
    lDtRuptEche : TdateTime ;
    lTLReq      : TList ;
    lIndexDR    : Integer ;
    lIndexDF    : Integer ;
{$IFNDEF EAGLCLIENT}
    lStRep      : String ;
    lBoXFichier : Boolean ;
    lStRacine   : String ;
    lInSouche   : Integer ;
{$ENDIF EAGLCLIENT}
    lBoDialog   : Boolean ;
    lMontant    : Double ;
    lMontantDev : Double ;
    lSens       : Double ;
    saveCodeOmr : Boolean ;
    lStRuptEtab : string ;
    lStRuptDoss : string ;
    lTLParent   : TList ;
begin

  result := False ;

  // Pas de pièces à éditer, pas d'édition...
  if vTLEche.Count = 0 then Exit ;

  // Init paramètres
  lStNature   := vTobScenario.GetValue('CPG_ETATNATURE') ;
  lStCode     := vTobScenario.GetValue('CPG_MODELEENCADECA') ;
  lBoDialog   := vTobScenario.getValue('CPG_MODELECHX') = 'X' ;

  lApercu     := vTobScenario.GetValue('APERCU')  = 'X' ;
  lIndexDR    := 1 ;
  lIndexDF    := 1 ;
  lMontant    := 0 ;
  lMontantDev := 0 ;
  lTobRupt    := nil ;
  saveCodeOmr := V_PGI.MiseSousPli ;
  V_PGI.MiseSousPli := ( vTobScenario.GetValue('CPG_CODEBARRE') = 'X' ) ;


  // Coéficient multiplicateur des montants <> pour ENC / DEC
  if vTobScenario.GetValue('CPG_FLUXENCADECA')='ENC'
     then lSens := -1.0
     else lSens := 1.0 ;

  // Gestion du spooler
  if vTobScenario.GetValue('SPOOLER') = 'X' then
    begin
    lStSpooler := vTobScenario.GetValue('REPSPOOLER') + '\' + lStCode + '.PDF' ;
    V_PGI.NoPrintDialog := True;
    V_PGI.QRPDF         := True;
    V_PGI.QRPDFQueue    := lStSpooler;
    V_PGI.QRPDFMerge    := lStSpooler;
//    lApercu             := True ; // FQ 17883 : Accès au spooler sans aperçu
    end ;

  try
    // Début du spooler
    if vTobScenario.GetValue('SPOOLER') = 'X' then
      StartPDFBatch( lStSpooler );

    // Pour une impression de documents -> gestion des tables temporaire DOCFACT et DOCREGL
    if (vTobScenario.GetValue('CPG_AVECDOC') = 'X') then
         // =================================================
         // ==== EDITION DOCUMENT VIA DOCFACT ET DOCREGL ====
         // =================================================
      begin
      // ********************************************************
      // **** Suppression préalable des chèques non imprimés ****
      // ********************************************************
      ExecuteSQL('DELETE FROM DOCREGLE WHERE DR_USER="'+V_PGI.User+'"') ;
      ExecuteSQL('DELETE FROM DOCFACT WHERE DF_USER="'+V_PGI.User+'"') ;
      lTLReq := TList.Create ;

      // ********************************
      // **** Parcours des échéances ****
      // ********************************
      for i := 0 to (vTLEche.Count - 1)  do
        begin

        lTobLigne := TOB( vTLEche[ i ] ) ;

        // Si rupture (ou dernière ligne), création d'un nouveau document DOCREGL
        if RuptureEtats( i, lTobLigne, vTobScenario, lStRuptAuxi, lDtRuptEche, lStRuptDoss, lStRuptEtab ) then
          begin
          GenereDOCREGL    ( lTobRupt, vTobScenario, lIndexDR, lMontant, lMontantDev ) ;
          GenereRequeteDoc ( lTobRupt, vTobScenario, lTLReq, lIndexDR ) ;
          lIndexDR    := lIndexDR + 1 ;
          lMontant    := 0 ;
          lMontantDev := 0 ;
          lIndexDF    := 0 ;
          end ;

        // MAJ des montants
        lMontant    := lMontant    + ( lTobLigne.GetValue('E_DEBIT')    - lTobLigne.GetValue('E_CREDIT') ) * lSens ;
        lMontantDev := lMontantDev + ( lTobLigne.GetValue('E_DEBITDEV') - lTobLigne.GetValue('E_CREDITDEV') ) * lSens ;

        // Création d'une ligne de DocFact
        GenereDOCFACT ( lTobLigne, vTobScenario, lIndexDR, lIndexDF ) ;
        lTobRupt    := lTobLigne ;
        lIndexDF    := lIndexDF + 1 ;

        end ;

      // Prise en compte de la dernière rupture
      GenereDOCREGL    ( lTobRupt, vTobScenario, lIndexDR, lMontant, lMontantDev ) ;
      GenereRequeteDoc ( lTobRupt, vTobScenario, lTLReq, lIndexDR ) ;

      // **************************************
      // **** Exécution du lanceDocument : ****
      // **************************************
      {$IFNDEF EAGLCLIENT} // Pas de spooler avec les documents en EAGL
      if vTobScenario.GetValue('SPOOLER') = 'X' then
        begin
        lStRep      := vTobScenario.GetValue('REPSPOOLER') ;           // Rép spooler
        lBoXFichier := vTobScenario.GetValue('XFICHIERSPOOLER')='X' ;  // Racine Nom fichier
        lStRacine   := vTobScenario.GetValue('CPG_MODELEENCADECA') ;   // X Fichiers
        lInSouche   := 0 ;                                             // Souche
        LanceDocumentSpool( 'L', lStNature, lStCode, lTLReq, nil, lApercu, lBoDialog,
                            FALSE, lStRep, lStRacine, lBoXFichier, lInSouche );
        end
      else
      {$ENDIF}
        LanceDocument     ( 'L', lStNature, lStCode, lTLReq, nil, lApercu, lBoDialog ) ;

      end
    else // Sinon utilisation d'une tob virtuelle avec lanceEtatTob
         // ==========================
         // ==== EDITION ETAT AGL ====
         // ==========================
        begin

        // Création TobEtat
        lTobEtat    := Tob.Create('TOB_ETAT', nil, -1 ) ;
        lTLParent := TList.create ;
//        _permuteTob( vTLEche, nil, nil ) ;
//        _permuteTob( vTLEche, lTLParent, nil ) ;

        // Parcours du règlement
        for i := 0 to vTLEche.Count - 1 do
          begin

          // Uniquement les lignes concernées (Echéances Tiers)
          lTobLigne := TOB( vTLEche[ i ] ) ;
          // Si rupture lancement de l'état avec pour auxi en cours
          if RuptureEtats ( i, lTobLigne, vTobScenario, lStRuptAuxi, lDtRuptEche, lStRuptDoss, lStRuptEtab ) then
            begin
            LanceEtatTOB( 'E' , lStNature , lStCode ,
                            lTobEtat,
                            lApercu, False, False, Nil, '', '', False ) ;
            lTobEtat.Detail.Clear ;
            end ;

          // duplication pour alimentation de l'état
          lTobLigne.ChangeParent( lTobEtat, -1 ) ;

          end ;
        // Si rupture lancement de l'état avec pour auxi en cours
        if lTobEtat.Detail.count > 0 then
          begin
          LanceEtatTOB( 'E' , lStNature , lStCode ,
                        lTobEtat,
                        lApercu, False, False, Nil, '', '', False ) ;  //manque le paramètre de PrintDialog.....
          end ;

        // Restauration pièces
//        _restaureTob( vTLEche, lTLParent ) ;
        _restaureTob( vTLEche, nil ) ;

        end ;

    // Fin du spooler
    if vTobScenario.GetValue('SPOOLER') = 'X' then
      begin
      CancelPDFBatch;
      {$IFNDEF EAGLCLIENT}
      if lApercu then // FQ 17883 : SBO 27/09/2006 : Permettre le spooler sans aperçu
        PreviewPDFFile('', lStSpooler );
      {$ENDIF}
      end ;

    result := True ;

  finally

    // Remise à zéro paramètres PDF
    if vTobScenario.GetValue('SPOOLER') = 'X' then
      begin
      V_PGI.QRPDF       := False;
      V_PGI.QRPDFQueue  := '';
      V_PGI.QRPDFMerge  := '';
      end ;
    // Gestion des codes barres
    V_PGI.MiseSousPli := saveCodeOmr ;

    // libération mémoire
    if Assigned( lTobEtat ) then
      begin
      lTobEtat.Detail.Clear ;
      FreeAndNil( lTobEtat ) ;
      end ;
    if Assigned( lTLParent ) then
      begin
      lTLParent.Clear ;
      FreeAndNil( lTLParent ) ;
      end ;
    if Assigned( lTLReq ) then
      FreeAndNil( lTLReq ) ;
      
  end;

end ;

Procedure AjouteInfosEtat     ( vTLEche : TList ; vTobScenario : TOB ; vInfoEcr : TInfoEcriture ) ;
var lTobInfos   : TOB ;
    lTobLigne   : TOB ;
    lTobContact : TOB ;
    lTobTiers   : TOB ;
    lTobBQ      : Tob ;
    lQBanque    : TQuery ;
    i         : Integer ;
    lStSql    : String ;
    lQInfos   : TQuery ;
    lStCode   : String ;
begin

  lTobInfos   := TOB.Create('$CONTACTS', nil, -1) ;
  lTobBQ      := nil ;

  // récup info banquecp
  lQBanque := OpenSelect( 'SELECT * FROM BANQUECP WHERE BQ_GENERAL="' + vTobScenario.GetString('CPG_GENERAL')
                         +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"', vInfoEcr.Dossier ) ; // 19/10/2006 YMO Multisociétés
  if not lQBanque.Eof then
    begin
    lTobBQ := TOB.Create('BANQUECP', nil, -1) ;
    lTobBQ.SelectDB( '', lQBanque ) ;
    end ;
  Ferme( lQBanque ) ;

  // Récupération des infos des tiers
  for i := 0 to vTLEche.Count - 1 do
    begin

    lTobLigne := Tob( vTLEche[ i ] ) ;

    // Ajout champs Supp
    InitInfosEtat( lTobLigne ) ;

    // Renseigne infos TIC TID
    if vTobScenario.GetString('TICTID')='X' then
      begin
      lStCode     := lTobLigne.GetValue('E_AUXILIAIRE') ;
      if vInfoEcr.LoadCompte( lStCode ) then
        begin
        lTobTiers := vInfoEcr.Compte.Item ;
        lTobLigne.PutValue('G_LIBELLE',     lTobTiers.GetValue('G_LIBELLE') ) ;
        lTobLigne.PutValue('G_CODEPOSTAL',  lTobTiers.GetValue('G_CODEPOSTAL') ) ;
        lTobLigne.PutValue('G_VILLE',       lTobTiers.GetValue('G_VILLE') ) ;
        lTobLigne.PutValue('G_ADRESSE1',    lTobTiers.GetValue('G_ADRESSE1') ) ;
        lTobLigne.PutValue('G_ADRESSE2',    lTobTiers.GetValue('G_ADRESSE2') ) ;
        lTobLigne.PutValue('G_ADRESSE3',    lTobTiers.GetValue('G_ADRESSE3') ) ;
        lTobLigne.PutValue('G_PAYS',        lTobTiers.GetValue('G_PAYS') ) ;
        lTobLigne.PutValue('G_TELEPHONE',   lTobTiers.GetValue('G_TELEPHONE') ) ;
        end ;
      end
    // Renseigne infos Tiers + contact
    else
      begin
      lStCode     := lTobLigne.GetValue('E_AUXILIAIRE') ;
      if vInfoEcr.LoadAux( lStCode ) then
        begin

        // infos tiers
        lTobTiers := vInfoEcr.Aux.Item ;
        lTobLigne.PutValue('T_JURIDIQUE',   lTobTiers.GetValue('T_JURIDIQUE') ) ;
        lTobLigne.PutValue('T_LIBELLE',     lTobTiers.GetValue('T_LIBELLE') ) ;
        lTobLigne.PutValue('T_CODEPOSTAL',  lTobTiers.GetValue('T_CODEPOSTAL') ) ;
        lTobLigne.PutValue('T_VILLE',       lTobTiers.GetValue('T_VILLE') ) ;
        lTobLigne.PutValue('T_ADRESSE1',    lTobTiers.GetValue('T_ADRESSE1') ) ;
        lTobLigne.PutValue('T_ADRESSE2',    lTobTiers.GetValue('T_ADRESSE2') ) ;
        lTobLigne.PutValue('T_ADRESSE3',    lTobTiers.GetValue('T_ADRESSE3') ) ;
        lTobLigne.PutValue('T_LIBELLE',     lTobTiers.GetValue('T_LIBELLE') ) ;
        lTobLigne.PutValue('T_PRENOM',      lTobTiers.GetValue('T_PRENOM') ) ;

        // infos contacts
        lTobContact := lTobInfos.FindFirst(['E_AUXILIAIRE'],[lStCode], False) ;
        if lTobContact = nil then
          begin
          // Chargement des infos pour le tiers
          lStSQL := 'SELECT C_CIVILITE, C_NOM, C_PRENOM, C_SERVICE, C_FONCTION, C_TELEPHONE'
                 + ' FROM CONTACT WHERE C_AUXILIAIRE="' + lStCode + '" AND C_PRINCIPAL="X"' ;
          lQInfos := OpenSql( lStSql, True ) ;
          if not lQInfos.Eof then
            lTobInfos.LoadDetailDB( 'CONTACT', '', '', lQinfos, True) ;
          Ferme(lQInfos) ;
          lTobContact := lTobInfos.FindFirst(['C_AUXILIAIRE'],[lStCode], False) ;
          end ;
        if lTobContact<>nil then
          begin
          lTobLigne.PutValue('C_CIVILITE',    lTobContact.GetValue('C_CIVILITE') ) ;
          lTobLigne.PutValue('C_NOM',         lTobContact.GetValue('C_NOM') ) ;
          lTobLigne.PutValue('C_PRENOM',      lTobContact.GetValue('C_PRENOM') ) ;
          lTobLigne.PutValue('C_SERVICE',     lTobContact.GetValue('C_SERVICE') ) ;
          lTobLigne.PutValue('C_FONCTION',    lTobContact.GetValue('C_FONCTION') ) ;
          lTobLigne.PutValue('C_TELEPHONE',   lTobContact.GetValue('C_TELEPHONE') ) ;
          end ;

        end ;

      end ;

    // Renseigne infos banquecp
    if Assigned( lTobBQ ) then
      begin
      lTobLigne.PutValue( 'BQ_CODE',          lTobBQ.getValue( 'BQ_CODE' )           ) ;
      lTobLigne.PutValue( 'BQ_LIBELLE',       lTobBQ.getValue( 'BQ_LIBELLE' )        ) ;
      lTobLigne.PutValue( 'BQ_ADRESSE1',      lTobBQ.getValue( 'BQ_ADRESSE1' )       ) ;
      lTobLigne.PutValue( 'BQ_ADRESSE2',      lTobBQ.getValue( 'BQ_ADRESSE2' )       ) ;
      lTobLigne.PutValue( 'BQ_ADRESSE3',      lTobBQ.getValue( 'BQ_ADRESSE3' )       ) ;
      lTobLigne.PutValue( 'BQ_VILLE',         lTobBQ.getValue( 'BQ_VILLE' )          ) ;
      lTobLigne.PutValue( 'BQ_GENERAL',       lTobBQ.getValue( 'BQ_GENERAL' )        ) ;
      lTobLigne.PutValue( 'BQ_CODEPOSTAL',    lTobBQ.getValue( 'BQ_CODEPOSTAL' )     ) ;
      lTobLigne.PutValue( 'BQ_PAYS',          lTobBQ.getValue( 'BQ_PAYS' )           ) ;
      lTobLigne.PutValue( 'BQ_ETABBQ',        lTobBQ.getValue( 'BQ_ETABBQ' )         ) ;
      lTobLigne.PutValue( 'BQ_GUICHET',       lTobBQ.getValue( 'BQ_GUICHET' )        ) ;
      lTobLigne.PutValue( 'BQ_NUMEROCOMPTE',  lTobBQ.getValue( 'BQ_NUMEROCOMPTE' )   ) ;
      lTobLigne.PutValue( 'BQ_CLERIB',        lTobBQ.getValue( 'BQ_CLERIB' )         ) ;
      lTobLigne.PutValue( 'BQ_CODEIBAN',      lTobBQ.getValue( 'BQ_CODEIBAN' )       ) ;
      lTobLigne.PutValue( 'BQ_TELEPHONE',     lTobBQ.getValue( 'BQ_TELEPHONE' )      ) ;
      lTobLigne.PutValue( 'BQ_FAX',           lTobBQ.getValue( 'BQ_FAX' )            ) ;
      lTobLigne.PutValue( 'BQ_TELEX',         lTobBQ.getValue( 'BQ_TELEX' )          ) ;
      lTobLigne.PutValue( 'BQ_DOMICILIATION', lTobBQ.getValue( 'BQ_DOMICILIATION' )  ) ;
      lTobLigne.PutValue( 'BQ_CODEBIC',       lTobBQ.getValue( 'BQ_CODEBIC' )        ) ;
      end ;

    // MONTANT ALS_MONTANT pour édition des bordereaux
    if EstDecaissement ( vTobScenario )
      then lTobLigne.AddChampSupValeur( 'ALS_MONTANT', lTobLigne.GetDouble('E_DEBIT') - lTobLigne.GetDouble('E_CREDIT') )
      else lTobLigne.AddChampSupValeur( 'ALS_MONTANT', lTobLigne.GetDouble('E_CREDIT') - lTobLigne.GetDouble('E_DEBIT') ) ;

    end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 26/02/2004
Modifié le ... :   /  /
Description .. : Remplit le multivalcombobox vComboMP avec la liste des
Suite ........ : modes de paiement référencant la catégorie de paiement
Suite ........ : vStCat.
Mots clefs ... : 
*****************************************************************}
Procedure CategorieVersModePaiement( vStCat : String ; vComboMP : THMultiValComboBox ) ;
Var QModePaie : TQuery ;
begin
  // Recherche
  if vStCat=''
     then QModePaie := OpenSQL('SELECT MP_MODEPAIE, MP_LIBELLE FROM MODEPAIE',True)
     else QModePaie := OpenSQL('SELECT MP_MODEPAIE, MP_LIBELLE FROM MODEPAIE WHERE MP_CATEGORIE="'+vStCat+'"',True) ;
  // Remplissage de la combo
  vComboMP.Values.Clear ;
  vComboMP.Items.Clear ;
  While not QModePaie.Eof do
    BEGIN
    vComboMP.Values.Add( QModePaie.FindField('MP_MODEPAIE').AsString ) ;
    vComboMP.Items.Add(  QModePaie.FindField('MP_LIBELLE').AsString ) ;
    QModePaie.Next ;
    END ;
  Ferme(QModePaie) ;
  // Sélectionner toutes la liste
  vComboMP.selectAll ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 02/03/2004
Modifié le ... :   /  /    
Description .. : Genère un ligne dans la table DocFact représentant
Suite ........ : l'échéance contenu dans vTobLigne
Mots clefs ... :
*****************************************************************}
Procedure GenereDOCFACT ( vTobLigne, vTobScenario : TOB ; vIndexDR, vIndexDF : Integer ) ;
var lCoef      : Double ;
    Montant    : Double ;
    Montantdev : Double;
    MontantEsc : Double ;
    TauxEsc    : Double ;
    vTobDocF   : TOB;
begin
  //SG6 06/01/05 Gestion sous forme de tob
  //             Ajoute de nouveaux champs dans DocFact

  vTobDocF:=TOB.Create('DOCFACT',nil,-1);

  // Coéficient multiplicateur des montants <> pour ENC / DEC
  if vTobScenario.GetValue('CPG_FLUXENCADECA')='ENC'
     then lCoef := -1.0
     else lCoef := 1.0 ;

  //Montant Esc
  Montant:= vTobLigne.GetValue('E_DEBIT') - vTobLigne.GetValue('E_CREDIT') ;
  if ( vTobScenario.GetValue('CPG_ESCMETHODE')<>'RIE' ) and
     ( vTobLigne.FieldExists('MONTANTESC') ) and
     ( vTobLigne.GetValue('MONTANTESC') <> 0) then
     montant := montant + vTobLigne.GetValue('MONTANTESC') ;

  //Montant Dev
  Montantdev:= vTobLigne.GetValue('E_DEBITDEV') - vTobLigne.GetValue('E_CREDITDEV') ;
  if ( vTobScenario.GetValue('CPG_ESCMETHODE')<>'RIE' ) and
     ( vTobLigne.FieldExists('MONTANTESCDEV') ) and
     ( vTobLigne.GetValue('MONTANTESCDEV') <> 0) then
     montantdev := montantdev + vTobLigne.GetValue('MONTANTESCDEV') ;
  // Gestion de l'escompte
  MontantEsc  := 0 ;
  TauxEsc     := 0 ;
  if vTobScenario.GetValue('CPG_ESCMETHODE')<>'RIE' then
    begin
    MontantEsc := vTobLigne.GetValue('MONTANTESCDEV') ;
    TauxEsc    := vTobLigne.GetValue('TAUXESC') ;
    end ;


  vTobDocF.PutValue('DF_USER',             V_PGI.User);
  vTobDocF.PutValue('DF_DATEECHEANCE',     vTobLigne.GetDateTime('E_DATEECHEANCE'));
  vTobDocF.PutValue('DF_MONTANT',          Montant*lCoef);
  vTobDocF.PutValue('DF_MONTANTDEV',       Montantdev*lCoef);
  vTobDocF.PutValue('DF_LIBELLE',          vTobLigne.GetString('E_LIBELLE'));
  vTobDocF.PutValue('DF_REFERENCE',        vTobLigne.GetString('E_REFINTERNE'));
  vTobDocF.PutValue('DF_REFEXTERNE',       vTobLigne.GetString('E_REFEXTERNE'));
  vTobDocF.PutValue('DF_ORIGINE',          vIndexDR);
  vTobDocF.PutValue('DF_DATECOMPTABLE',    vTobLigne.GetDateTime('E_DATECOMPTABLE'));
  vTobDocF.PutValue('DF_COUVERTURE',       vTobLigne.GetDouble('E_COUVERTURE'));
  vTobDocF.PutValue('DF_COUVERTUREDEV',    vTobLigne.GetDouble('E_COUVERTUREDEV'));
  vTobDocF.PutValue('DF_INDICE',           vIndexDF);
  vTobDocF.PutValue('DF_DATEREFEXTERNE',   vTobLigne.GetDateTime('E_DATEREFEXTERNE'));
  vTobDocF.PutValue('DF_NUMEROPIECE',      vTobLigne.GetInteger('E_NUMEROPIECE'));
  vTobDocF.PutValue('DF_NUMLIGNE',         vTobLigne.GetInteger('E_NUMLIGNE'));
  vTobDocF.PutValue('DF_CONTREPARTIEAUX',  vTobLigne.GetString('E_CONTREPARTIEAUX'));
  vTobDocF.PutValue('DF_TAUXESC',          TauxEsc);
  vTobDocF.PutValue('DF_MONTANTESC',       MontantEsc);
  //SG6 06/01/05 Nouveau Champ
  vTobDocF.PutValue('DF_JOURNAL',          vTobLigne.GetString('E_JOURNAL'));
  vTobDocF.PutValue('DF_NATUREPIECE',      vTobLigne.GetString('E_NATUREPIECE'));
  vTobDocF.PutValue('DF_SOCIETE',          vTobLigne.GetString('E_SOCIETE'));
  vTobDocF.PutValue('DF_ETABLISSEMENT',    vTobLigne.GetString('E_ETABLISSEMENT'));
  vTobDocF.PutValue('DF_REFLIBRE',         vTobLigne.GetString('E_REFLIBRE'));
  vTobDocF.PutValue('DF_AFFAIRE',          vTobLigne.GetString('E_AFFAIRE'));
  vTobDocF.PutValue('DF_MODEPAIE',         vTobLigne.GetString('E_MODEPAIE'));
  vTobDocF.PutValue('DF_DATERELANCE',      vTobLigne.GetDateTime('E_DATERELANCE'));
  vTobDocF.PutValue('DF_NIVEAURELANCE',    vTobLigne.GetInteger('E_NIVEAURELANCE'));
  vTobDocF.PutValue('DF_TIERSPAYEUR',      vTobLigne.GetString('E_TIERSPAYEUR'));
  vTobDocF.PutValue('DF_QTE1',             vTobLigne.GetDouble('E_QTE1'));
  vTobDocF.PutValue('DF_QTE2',             vTobLigne.GetDouble('E_QTE2'));
  vTobDocF.PutValue('DF_NOMLOT',           vTobLigne.GetString('E_NOMLOT'));
  vTobDocF.PutValue('DF_LIBRETEXTE0',      vTobLigne.GetString('E_LIBRETEXTE0'));
  vTobDocF.PutValue('DF_LIBRETEXTE1',      vTobLigne.GetString('E_LIBRETEXTE1'));
  vTobDocF.PutValue('DF_LIBRETEXTE2',      vTobLigne.GetString('E_LIBRETEXTE2'));
  vTobDocF.PutValue('DF_LIBRETEXTE3',      vTobLigne.GetString('E_LIBRETEXTE3'));
  vTobDocF.PutValue('DF_LIBRETEXTE4',      vTobLigne.GetString('E_LIBRETEXTE4'));
  vTobDocF.PutValue('DF_LIBRETEXTE5',      vTobLigne.GetString('E_LIBRETEXTE5'));
  vTobDocF.PutValue('DF_LIBRETEXTE6',      vTobLigne.GetString('E_LIBRETEXTE6'));
  vTobDocF.PutValue('DF_LIBRETEXTE7',      vTobLigne.GetString('E_LIBRETEXTE7'));
  vTobDocF.PutValue('DF_LIBRETEXTE8',      vTobLigne.GetString('E_LIBRETEXTE8'));
  vTobDocF.PutValue('DF_LIBRETEXTE9',      vTobLigne.GetString('E_LIBRETEXTE9'));
  vTobDocF.PutValue('DF_TABLE0',           vTobLigne.GetString('E_TABLE0'));
  vTobDocF.PutValue('DF_TABLE1',           vTobLigne.GetString('E_TABLE1'));
  vTobDocF.PutValue('DF_TABLE2',           vTobLigne.GetString('E_TABLE2'));
  vTobDocF.PutValue('DF_TABLE3',           vTobLigne.GetString('E_TABLE3'));
  vTobDocF.PutValue('DF_LIBREMONTANT0',    vTobLigne.GetDouble('E_LIBREMONTANT0'));
  vTobDocF.PutValue('DF_LIBREMONTANT1',    vTobLigne.GetDouble('E_LIBREMONTANT1'));
  vTobDocF.PutValue('DF_LIBREMONTANT2',    vTobLigne.GetDouble('E_LIBREMONTANT2'));
  vTobDocF.PutValue('DF_LIBREMONTANT3',    vTobLigne.GetDouble('E_LIBREMONTANT3'));
  vTobDocF.PutValue('DF_NUMTRAITECHQ',     vTobLigne.GetString('E_NUMTRAITECHQ'));
  vTobDocF.PutValue('DF_NUMENCADECA',      vTobLigne.GetString('E_NUMENCADECA'));
  vTobDocF.PutValue('DF_REFGESCOM',        vTobLigne.GetString('E_REFGESCOM'));
  vTobDocF.PutValue('DF_LIBREDATE',        vTobLigne.GetDateTime('E_LIBREDATE'));
  vTobDocF.PutValue('DF_SAISIMP',          vTobLigne.GetDouble('E_SAISIMP'));

  //Insertion
  vTobDocF.InsertDB(nil);


end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Genère un ligne dans la table DocRegl représentant le
Suite ........ : document obtenu à partir de vTobLigne
Mots clefs ... :
*****************************************************************}
Procedure GenereDOCREGL ( vTobLigne, vTobScenario : TOB ; vIndexDoc : Integer ; vMontant, vMontantDev : Double ) ;
var lTobDR : Tob ;
begin
  // Pb DB2 : gestion sous forme de TOB
  lTobDR := Tob.Create('DOCREGLE', nil, -1) ;

  lTobDR.InitValeurs ;

  lTobDR.PutValue( 'DR_USER',              V_PGI.User ) ;
  lTobDR.PutValue( 'DR_AUXILIAIRE',        vTobLigne.GetValue('E_AUXILIAIRE') ) ;
  lTobDR.PutValue( 'DR_DATEECHEANCE',      vTobLigne.GetValue('E_DATEECHEANCE') ) ;
  lTobDR.PutValue( 'DR_DATECOMPTABLE',     vTobLigne.GetValue('E_DATECOMPTABLE') ) ;
  lTobDR.PutValue( 'DR_MODEPAIE',          vTobLigne.GetValue('E_MODEPAIE') ) ;
  lTobDR.PutValue( 'DR_MONTANT',           vMontant ) ;
  lTobDR.PutValue( 'DR_MONTANTDEV',        vMontantDev ) ;
  lTobDR.PutValue( 'DR_ORIGINE',           vIndexDoc ) ;
  lTobDR.PutValue( 'DR_LIBELLE',           vTobLigne.GetValue('E_LIBELLE') ) ;
  lTobDR.PutValue( 'DR_NUMCHEQUE',         Copy( vTobLigne.GetValue('E_REFINTERNE') ,1 ,17 ) ) ;
  lTobDR.PutValue( 'DR_EDITE',             '-' ) ;
  lTobDR.PutValue( 'DR_RIB',               vTobLigne.GetValue('E_RIB') ) ;
  lTobDR.PutValue( 'DR_RIBEMETTEUR',       GenereRibPourGeneral(vTobScenario.GetValue('CPG_GENERAL')) ) ;
  lTobDR.PutValue( 'DR_CPTBANQUE',         vTobScenario.GetValue('CPG_GENERAL') ) ;
  lTobDR.PutValue( 'DR_GENERAL',           vTobLigne.GetValue('E_GENERAL') ) ;
  lTobDR.PutValue( 'DR_REFTIRE',           vTobLigne.GetValue('E_REFEXTERNE') ) ;
  lTobDR.PutValue( 'DR_NUMEROPIECE',       vTobLigne.GetValue('E_NUMEROPIECE') ) ;
  lTobDR.PutValue( 'DR_BANQUEPREVI',       vTobLigne.GetValue('E_BANQUEPREVI') ) ;

  lTobDR.InsertDB(nil) ;

  FreeAndNil( lTobDR ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Ajoute dans vTLReq la requête qui sera utilisé pour imprimer
Suite ........ : vTobLigne via le lanceDocument
Mots clefs ... :
*****************************************************************}
Procedure GenereRequeteDoc    ( vTobLigne, vTobScenario : TOB ; vTLReq : TList ; vIndexDoc : Integer ) ;
var LL     : TStringList ;
    lStSql : String ;
begin
  // Création de la requête
  lStSql := 'SELECT * FROM DOCREGLE' ;
  lStSql := lStSql + ' LEFT JOIN DOCFACT ON DF_USER=DR_USER AND DF_ORIGINE=DR_ORIGINE' ;

  if vTobScenario.GetValue('TICTID') = 'X'
    then lStSql := lStSql + ' LEFT JOIN GENERAUX ON G_GENERAL=DR_AUXILIAIRE'
    else lStSql := lStSql + ' LEFT JOIN TIERS ON T_AUXILIAIRE=DR_AUXILIAIRE'
                    + ' LEFT OUTER JOIN CONTACT ON C_AUXILIAIRE=T_AUXILIAIRE AND C_PRINCIPAL="X"' ;

  lStSql := lStSql + ' LEFT OUTER JOIN BANQUECP ON BQ_GENERAL=DR_CPTBANQUE AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"' ; // 19/10/2006 YMO Multisociétés
  lStSql := lStSql + ' WHERE DR_USER="' + V_PGI.User + '"' ;
  lStSql := lStSql + ' AND DR_ORIGINE=' + IntToStr( vIndexDoc ) ;

  // Ajout de la requête à la liste
  LL:=TStringList.Create ;
  LL.Add(lStSql) ;
  vTLReq.Add(LL) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Retourne un booléen indiquant si l'échéance de vTobLigne
Suite ........ : doit aparaître dans un nouveau document, en se basant sur
Suite ........ : l'auxiliaire et la date d'échéance.
Mots clefs ... :
*****************************************************************}
Function  RuptureEtats ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ;
                           var vAuxiliaire : String ; var vDateEche : TDateTime ; var vDossier, vEtab : string ) : Boolean ;
var lStGroupe       : String ;
    lBoRuptDossier  : Boolean ;
    lBoRuptEtab     : Boolean ;
begin

  Result    := False ;
  lStGroupe := vTobScenario.GetValue('CPG_GROUPEENCADECA') ;

  // Première ligne
  if vInIndex>0 then
  begin
    lBoRuptDossier := ( vTobScenario.GetString('CPG_MULTISOC') = 'X' ) and ( vTobLigne.GetNumChamp('SYSDOSSIER') > 0 ) and
                      ( vTobScenario.GetString('CPG_MULTISOCMETH') <> 'GLO' ) and ( vDossier <> vTobLigne.GetString('SYSDOSSIER') ) ;

    lBoRuptEtab    := ( vTobScenario.GetString('CPG_MULTIETAB') = 'X' ) and ( vTobLigne.GetNumChamp('E_ETABLISSEMENT') > 0 ) and
                      ( vTobScenario.GetString('CPG_METHODEETAB') = 'DET' ) and ( vEtab <> vTobLigne.GetString('E_ETABLISSEMENT') ) ;

    // Génération au détail, Rupture à chaque ligne
    if lStGroupe='DET'
      then Result := True
    // Rupture par auxiliaire ou globalisé
    else if (lStGroupe='AUX') or (lStGroupe='GLO')
      then Result := ( vAuxiliaire <> vTobLigne.GetValue('E_AUXILIAIRE') )
                     or lBoRuptEtab or lBoRuptDossier
    // Rupture par auxiliaire/échéance
    else if (lStGroupe='ECT') or (lStGroupe='ECH')
      then Result := ( vAuxiliaire <> vTobLigne.GetValue('E_AUXILIAIRE') )
                     or ( vDateEche <> vTobLigne.GetValue('E_DATEECHEANCE') )
                     or lBoRuptEtab or lBoRuptDossier ;
  end ;

  // MAJ des champs de rupture
  if Result or ( vInIndex = 0 ) then
    begin
    vAuxiliaire  := vTobLigne.GetString('E_AUXILIAIRE') ;
    vDateEche    := vTobLigne.GetDateTime('E_DATEECHEANCE') ;
    if vTobLigne.GetNumChamp('SYSDOSSIER') > 0 then
       vDossier     := vTobLigne.GetString('SYSDOSSIER') ;
    vEtab        := vTobLigne.GetString('E_ETABLISSEMENT') ;
    end ;

end ;

Function  GenereRibPourGeneral( vStGeneral : String ) : String ;
Var lQJal : TQuery ;
begin
  Result := '' ;
  if vStGeneral='' then Exit ;

  // Recherche RIB
  lQJal := OpenSQL('SELECT BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_DOMICILIATION, BQ_PAYS FROM BANQUECP WHERE BQ_GENERAL="'+vStGeneral
                  +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
  if Not lQJal.EOF then //XVI 24/02/2005
    begin
    Result := EncodeRIB( lQJal.FindField('BQ_ETABBQ').AsString,
                         lQJal.FindField('BQ_GUICHET').AsString,
                         lQJal.FindField('BQ_NUMEROCOMPTE').AsString,
                         lQJal.FindField('BQ_CLERIB').AsString,
                         lQJal.FindField('BQ_DOMICILIATION').AsString,
                         CodeISOduPays(lQJal.FindField('BQ_PAYS').AsString) ) ;
  END ;

  Ferme(lQJal) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 30/03/2004
Modifié le ... :   /  /
Description .. : Renseigne les numéro de chèques / traites tel qu'ils seront
Suite ........ : édités dans les pièces d'origine.
Mots clefs ... :
*****************************************************************}
Procedure MajNumTraiteChq  ( vTobOrigine, vTobScenario : TOB ) ;
var lInCpt       : Integer ;
    lStRuptAuxi  : String ;
    lDtRuptEche  : TDateTime ;
    lInNumChq    : Integer ;
    lStNoDepart  : String ;
    lStValeur    : String ;
    lTobLigne    : TOB ;
    lStRuptDoss  : string ;
    lStRuptEtab  : string ;
begin

  // Init param souche
  vTobScenario.PutValue('NOMPARAMSOCTRAITE', '' ) ;

  // Pour les LCR, on se base sur le paramsoc (en fonction type d'auxiliaire)
  if vTobScenario.GetValue('CPG_TYPEENCADECA') = 'LCR' then
       begin
         lTobLigne := vTobOrigine.Detail[0] ;
         if (lTobLigne.GetValue('T_NATUREAUXI') = 'CLI') or (lTobLigne.GetValue('T_NATUREAUXI') = 'AUC')
           or (lTobLigne.GetValue('G_NATUREGENE') = 'TIC') then
           begin
             vTobScenario.PutValue('NOMPARAMSOCTRAITE', 'SO_CPNUMTRACLI' ) ;
             vTobScenario.PutValue('REFCHEQUE', GetParamSocSecur('SO_CPSOUCHETRACLI', '') ) ;
             vTobScenario.PutValue('NUMCHEQUE', IntToStr( GetParamSocSecur('SO_CPNUMTRACLI', 0) ) ) ;
           end
         else
           begin
             vTobScenario.PutValue('NOMPARAMSOCTRAITE', 'SO_CPNUMTRAFOU' ) ;
             vTobScenario.PutValue('REFCHEQUE', GetParamSocSecur('SO_CPSOUCHETRAFOU', '') ) ;
             vTobScenario.PutValue('NUMCHEQUE', IntToStr( GetParamSocSecur('SO_CPNUMTRAFOU', 0) ) ) ;
           end ;
       end ;

  // MAJ des champs E_NUMTRAITECHQ des échéances
  lStNoDepart  := vTobScenario.GetValue('NUMCHEQUE') ;
  lStValeur    := lStNoDepart ;
  for lInCpt := 0 to vTobOrigine.Detail.Count - 1 do
    begin
    lTobLigne := vTobOrigine.Detail[ lInCpt ] ;
    if RuptureEtats(lInCpt, lTobLigne, vTobScenario, lStRuptAuxi, lDtRuptEche, lStRuptDoss, lStRuptEtab ) then
      begin
      lInNumChq := StrToInt(lStValeur) + 1 ;
      lStValeur  := IntToStr(lInNumChq) ;
      While Length(lStValeur) < Length(lStNoDepart)
          do lStValeur := '0' + lStValeur ;
      end ;
    lTobLigne.PutValue('E_NUMTRAITECHQ', lStValeur ) ;
    end ;

  // MAJ du scenario
  if lStValeur<>lStNoDepart then
     vTobScenario.PutValue('NUMCHEQUE',     lStValeur ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 30/03/2004
Modifié le ... :   /  /
Description .. : Retourne Vrai si l'opération inclus une gestion de
Suite ........ : E_NUMTRAITECHQ
Mots clefs ... :
*****************************************************************}
Function  AvecNumTraiteChq( vTobScenario : TOB ) : Boolean ;
begin
  Result := ( vTobScenario.GetValue('CPG_AVECNUMCHEQUE')='X' )
            or ( ( vTobScenario.GetValue('CPG_TYPEENCADECA')='LCR' ) and ( vTobScenario.GetValue('CPG_EDITE')='X' ) ); // FQ 17031
end ;

Function UpdateEcheancesOrigines    ( vTobOrigine, vTobScenario : TOB ) : Boolean ;
var lInCpt       : Integer ;
    lTobLigne    : TOB ;
    lStNumChq    : string ;
    lStDossier   : string ;
begin

  try

    Result := False ;

    // Parcours des échéances pour maj base
    for lInCpt := 0 to vTobOrigine.Detail.Count - 1 do
      begin
      lTobLigne  := vTobOrigine.Detail[ lInCpt ] ;
      lStNumChq  := lTobLigne.GetValue('E_NUMTRAITECHQ') ;
      if vTobOrigine.GetNumChamp('SYSDOSSIER')>0
        then lStDossier := vTobOrigine.GetString('SYSDOSSIER')
        else lStDossier := '' ;

      ExecuteSQL('UPDATE ' + GetTableDossier( lStDossier, 'ECRITURE' ) + ' SET E_NUMTRAITECHQ="' + lStNumChq +
                    '" WHERE E_JOURNAL="'        + lTobLigne.GetValue('E_JOURNAL')                    + '" AND '
                          + 'E_EXERCICE="'       + lTobLigne.GetValue('E_EXERCICE')                   + '" AND '
                          + 'E_DATECOMPTABLE="'  + UsDateTime( lTobLigne.GetValue('E_DATECOMPTABLE') )+ '" AND '
                          + 'E_NUMEROPIECE='     + IntToStr( lTobLigne.GetValue('E_NUMEROPIECE') )    + ' AND '
                          + 'E_NUMLIGNE='        + IntToStr( lTobLigne.GetValue('E_NUMLIGNE') )       + ' AND '
                          + 'E_NUMECHE='         + IntToStr( lTobLigne.GetValue('E_NUMECHE') )
                );
      end ;

    // Arrivé là, tout s'est bien passé
    Result := True ;

  finally

  end ;

end ;


Procedure GenereEscompte      ( vTobLigne, vTobEscompte, vTobScenario : TOB ; vInfoEcr : TInfoEcriture ) ;
var lTauxEsc        : Double ;        // Taux d'escompte
    lTauxTva        : Double ;        // Taux de tva
    lStCompteTva    : String ;        // compte d'escompte
    lStCompteHT     : String ;        // Compte de Tva
    lStLibelle      : String ;        // Libellé de l'escompte
    lDebitEsc       : Double ;        // Montant débiteur de l'escompte
    lCreditEsc      : Double ;        // Montant créditeur de l'escompte
    lDebitHT        : Double ;        // Montant débiteur de la ligne HT
    lCreditHT       : Double ;        // Montant créditeur de la ligne HT
    lDebitTva       : Double ;        // Montant débiteur de la ligne de Tva
    lCreditTva      : Double ;        // Montant débiteur de la ligne de Tva
    lBoAvecTva      : Boolean ;       // indicateur de génération de la ligne de Tva
    lPieceEsc       : TPieceCompta ;
    lDev            : RDevise ;
    lNumLigne       : Integer ;
begin

  if vTobLigne = nil then Exit ;
  if vTobScenario.GetValue('CPG_ESCMETHODE')='RIE' then Exit ;
  if GetParamSocSecur('SO_CPJALESCOMPTE', '')='' then Exit ;

  // gestion de la TVA ?
  lBoAvecTva := ( vTobScenario.GetValue('CPG_ESCTVAAVEC') = 'X' ) and
                ( vTobScenario.GetValue('CPG_ESCTVATAUX') <> 0  ) and
                ( vTobScenario.GetValue('CPG_ESCTVACPT') <> ''  ) ;

  // Init. variables
  lStCompteTva := vTobScenario.GetValue('CPG_ESCTVACPT') ;
  lStCompteHT  := vTobScenario.GetValue('CPG_ESCCPTGENE') ;
  lDebitEsc    := 0 ;
  lCreditEsc   := 0 ;
  lDev         := vInfoEcr.Devise.Dev ;

  // Init. Taux escompte
  lTauxEsc     := DetermineTauxEsc ( vTobLigne, vTobScenario, vInfoEcr, lDev ) ;
  if lTauxEsc = 0 then Exit ;

  // Init. Taux TVA
  if lBoAvecTva
     then lTauxTva := vTobScenario.GetValue('CPG_ESCTVATAUX')
     else lTauxTva := 0 ;

  // Init. de la pièce d'escompte
  lPieceEsc := CreerPieceEscompte( vTobScenario, vTobLigne, vTobEscompte, vInfoEcr ) ;

  // Construction du libelle de l'escompte (maxi 35 caractères) Fiche 10441 :
  // - Escpt S/pièce N° <Numéro pièce d'origine> <Libelle tiers tronqué>
  lStLibelle := 'Escpte S/pièce N° ' + IntToStr( vTobLigne.GetValue('E_NUMEROPIECE') ) ;
  if vInfoEcr.LoadAux( vTobLigne.GetValue('E_AUXILIAIRE') )
     then lStLibelle := lStLibelle + ' ' + vInfoEcr.Aux_GetValue('T_LIBELLE')
     else lStLibelle := lStLibelle + ' ' + vTobLigne.GetValue('E_AUXILIAIRE') + ' ' ;
  lStLibelle := Copy( lStLibelle, 1, 35) ;

  // =============================
  // ==== Calcul des montants ====
  // =============================
  // le montant de l'escompte
  if vTobLigne.GetValue('E_DEBITDEV') <> 0 then
    lDebitEsc    := vTobLigne.GetValue('E_DEBITDEV')  - vTobLigne.GetValue('E_COUVERTUREDEV') ;
  if vTobLigne.GetValue('E_CREDITDEV') <> 0 then
    lCreditEsc   := vTobLigne.GetValue('E_CREDITDEV') - vTobLigne.GetValue('E_COUVERTUREDEV') ;
  lDebitEsc  := Arrondi( ( lDebitEsc * ( lTauxEsc / 100 ) ) , lDev.Decimale ) ;
  lCreditEsc := Arrondi( ( lCreditEsc * ( lTauxEsc / 100 ) ) , lDev.Decimale ) ;
  // Le montant du HT
  lDebitHT   := Arrondi( lDebitEsc  / ( 1.0 + lTauxTva / 100.0 ), lDev.Decimale ) ;
  lCreditHT  := Arrondi( lCreditEsc / ( 1.0 + lTauxTva / 100.0 ), lDev.Decimale ) ;
  // Le montant de la TVA
  lDebitTva  := lDebitEsc  - lDebitHT ;
  lCreditTva := lCreditEsc - lCreditHT ;

  // =============================
  // ==== ligne TTC ( TIERS ) ====
  // =============================
  lPieceEsc.NewRecord ;
  lNumLigne := lPieceEsc.CurIdx ;
  lPieceEsc.PutValue( lNumLigne, 'E_GENERAL',        vTobLigne.GetValue('E_GENERAL')    ) ;
  lPieceEsc.PutValue( lNumLigne, 'E_AUXILIAIRE',     vTobLigne.GetValue('E_AUXILIAIRE') ) ;
  // Recopie données échéances
  RecopieInfosEcheance( vTobLigne , lPieceEsc, lNumLigne ) ;
  // FQ 22141 : initialiser la date d'échéance à la date comptable du règlement
  lPieceEsc.PutValue( lNumLigne, 'E_DATEECHEANCE',   vTobScenario.GetDateTime('DATECOMPTABLE') ) ;
  // Fin FQ 22141
  // Affecation infos
  lPieceEsc.PutValue( lNumLigne, 'E_QUALIFORIGINE',  'ESC') ;
  lPieceEsc.PutValue( lNumLigne, 'E_LIBELLE',        lStLibelle) ;
  // Affectation des montants
  //  Attention : debit/Crédit inversé sur le TTC, pour besoin de la pièce d'escompte.
  // FQ 21090 SBO 07/08/2007 : Pb génération de l'escompte
  if lCreditEsc = 0
    then lPieceEsc.PutValue( lNumLigne, 'E_CREDITDEV', lDebitEsc )
    else lPieceEsc.PutValue( lNumLigne, 'E_DEBITDEV',  lCreditEsc ) ;

  // ==================
  // ==== Ligne HT ====
  // ==================
  lPieceEsc.NewRecord ;
  lNumLigne := lPieceEsc.CurIdx ;
  lPieceEsc.PutValue( lNumLigne, 'E_GENERAL',        lStCompteHT    ) ;
  lPieceEsc.PutValue( lNumLigne, 'E_QUALIFORIGINE',  'ESC') ;
  lPieceEsc.PutValue( lNumLigne, 'E_LIBELLE',        lStLibelle) ;
  // Affectation des montants
  // FQ 21090 SBO 07/08/2007 : Pb génération de l'escompte
  if lDebitHT = 0
    then lPieceEsc.PutValue( lNumLigne, 'E_CREDITDEV', lCreditHT )
    else lPieceEsc.PutValue( lNumLigne, 'E_DEBITDEV',  lDebitHT  ) ;

  // ===================
  // ==== Ligne TVA ====
  // ===================
  if vTobScenario.GetValue('CPG_ESCTVAAVEC') = 'X' then
    begin
   lPieceEsc.NewRecord ;
   lNumLigne := lPieceEsc.CurIdx ;
   lPieceEsc.PutValue( lNumLigne, 'E_GENERAL',        lStCompteTVA    ) ;
   lPieceEsc.PutValue( lNumLigne, 'E_QUALIFORIGINE',  'ESC') ;
   lPieceEsc.PutValue( lNumLigne, 'E_LIBELLE',        lStLibelle) ;
   // Affectation des montants
   // FQ 21090 SBO 07/08/2007 : Pb génération de l'escompte
    if lDebitTva = 0
      then lPieceEsc.PutValue( lNumLigne, 'E_CREDITDEV', lCreditTva )
      else lPieceEsc.PutValue( lNumLigne, 'E_DEBITDEV',  lDebitTva  ) ;
    end ;


  // =============================================
  // ==== Report de l'escompte sur l'échéance ====
  // =============================================
  CSetMontants( vTobLigne,                                             // Ligne d'échéance
                ( vTobLigne.GetValue('E_DEBITDEV') - lDebitEsc ),      // Débit escompté
                ( vTobLigne.GetValue('E_CREDITDEV') - lCreditEsc ),    // Crédit escompté
                vInfoEcr.Devise.Dev,                                   // Devise
                False );
  vTobLigne.AddChampSupValeur('MONTANTESC',      lPieceEsc.GetValue(1, 'E_DEBIT') - lPieceEsc.GetValue(1, 'E_CREDIT') ) ;
  vTobLigne.AddChampSupValeur('MONTANTESCDEV',   lPieceEsc.GetValue(1, 'E_DEBITDEV') - lPieceEsc.GetValue(1, 'E_CREDITDEV') ) ;
  vTobLigne.AddChampSupValeur('TAUXESC',         lTauxEsc ) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 29/04/2004
Modifié le ... :   /  /    
Description .. : Retourne le taux d'escompte en fonction de la méthode 
Suite ........ : sélectionnée.
Mots clefs ... : 
*****************************************************************}
Function  DetermineTauxEsc    ( vTobLigne, vTobScenario : TOB ; vInfoEcr: TInfoEcriture ; vDev : RDevise ) : Double ;
var lDeltaGene  : Double ;
    lDeltaEche  : Double ;
begin

  // détermination du taux
  Result := vTobScenario.GetValue('CPG_ESCTAUX') ;
  if vTobScenario.GetValue('CPG_ESCMETHODE')='TXT' then
    begin
    if vInfoEcr.LoadAux( vTobLigne.GetValue('E_AUXILIAIRE') ) then
       if vInfoEcr.Aux_GetValue('T_ESCOMPTE')<>0 then
          Result := vInfoEcr.Aux_GetValue('T_ESCOMPTE') ;		// Taux d'escompte du tiers
    end ;

  // Gestion de la proratisation
  if vTobScenario.GetValue('CPG_ESCPRORATA')='X' then
    begin
    // Calcul des écarts de date pour l'éventuelle proratisation
    lDeltaGene  := Arrondi( vTobLigne.GetValue('E_DATEECHEANCE') - vTobScenario.GetValue('DATECOMPTABLE'), 0 ) ;
    lDeltaEche  := Arrondi( vTobLigne.GetValue('E_DATEECHEANCE') - vTobLigne.GetValue('E_DATECOMPTABLE'),  0 ) ;
    // Calcul du montant escompté max pour recalcul du taux
    if ( lDeltaGene <= 0 ) or ( lDeltaEche <= 0)
       then Result := 0
       // Calcul du taux proratisé
       else Result := Arrondi( result * ( lDeltaGene / lDeltaEche ) , 4 ) ;
    end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 29/04/2004
Modifié le ... :   /  /    
Description .. : Execute l'export de la pièce récement comptabilisé par 
Suite ........ : l'utilisateur courant.
Mots clefs ... : 
*****************************************************************}
{
Function  ExecuteExportCFONB ( vTobScenario : TOB ) : Boolean ;
var lErr        : TMC_ERR ;
    lTobExp     : TOB ;
    lStNumCFONB : String ;
    lDtNowFutur : TDateTime ;
    lStFichier  : String ;
begin

  Result := False ;

  if ( vTobScenario.GetValue('CPG_CFONBEXPORT') <> 'X' ) or
     ( Trim(vTobScenario.GetValue('EXPORTFICHIER')) = '' ) then Exit ;

  try
    // Détermination du nom du fichier
    lStFichier := vTobScenario.GetValue('EXPORTFICHIER') ;
    if not DetermineFichierExport ( lStFichier ) then
      EAbort.Create('Erreur lors de la détermination du nom du fichier d''export !' ) ;
    vTobScenario.PutValue('EXPORTFICHIER', lStFichier) ;

    // Effacement du paramètrage contextuel de l'export
    ExecuteSQL('DELETE FROM CPARAMEXPORT WHERE CPX_USER="' + V_PGI.User + '"' ) ;

    // Création du paramètrage contextuel pour l'export a venir
    lTobExp := TOB.Create('CPARAMEXPORT',nil, -1) ;
    lTobExp.putValue('CPX_USER',            V_PGI.User ) ;
    lTobExp.putValue('CPX_REMISEREF',       vTobScenario.GetValue('REMISEREF') ) ;
    lTobExp.putValue('CPX_REMISEDATE',      vTobScenario.GetValue('REMISEDATE') ) ;
    lTobExp.putValue('CPX_FRAISCODE',       vTobScenario.GetValue('FRAISIMPUTATION') ) ;
    lTobExp.InsertDb(nil);

    ExporteLP( 'F', 'FIC', vTobScenario.GetValue('CPG_CFONBFORMAT'),
               'E_CFONBOK="#" AND E_UTILISATEUR="' + V_PGI.User + '"',
               // AND E_DATECOMPTABLE="'                  // Clause Where
               //     + UsDateTime( vTobScenario.GetValue('DATECOMPTABLE')) + '"',
               lStFichier,                                            // Fichiers
               vTobScenario.GetValue('EXPORTAPERCU')='X',             // Aperçu
               True,                                                  // Sur base ?
               nil,                                                   // TOB
               lErr ) ;                                               // TMC_ERR ???

    // MAJ des écritures si fichier d'export créés
    Result := FileExists(lStFichier) ;
    if Result then
      begin
      lStNumCFONB := GetNumCFONB( vTobScenario.GetValue('CPG_TYPEENCADECA') ) ;
      lDtNowFutur := NowH ;
      ExecuteSQL('UPDATE ECRITURE SET  E_CFONBOK="X", ' +
                                      'E_NUMCFONB="'   + lStNumCFONB + '", ' +
                                      'E_DATEMODIF="'  + UsTime( lDtNowFutur ) +
                      '" WHERE E_CFONBOK="#" AND E_UTILISATEUR="' + V_PGI.User + '"' ) ;

      end ;

    finally

      // Effacement du paramètrage contextuel de l'export
      ExecuteSQL('DELETE FROM CPARAMEXPORT WHERE CPX_USER="' + V_PGI.User + '"' ) ;

      if result
        then PGIInfo('Le fichier d''export ' + lStFichier + ' a été correctement créé.')
        else PGIBox('Le fichier d''export n''a pas été créé. Les écritures n''ont pas été mises à jour.') ;

    end ;

end ;
}

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 29/04/2004
Modifié le ... :   /  /    
Description .. : Cherche un nom de fichier nom utilisé en se basant sur le 
Suite ........ : nom de fichier saisi par l'utilisateur.
Mots clefs ... :
*****************************************************************}
{
function DetermineFichierExport ( var vFichier : String ) : Boolean ;
var lStNewFichier  : String ;
    lStBaseFichier : String ;
    lStExtFichier  : String ;
    lInCompteur    : Integer ;
begin

  // Si le nom du fichier saisi par l'utilisteur n'existe pas alors ok
  Result := True ;
  if not FileExists(vFichier) then Exit ;

  // Sinon, ajout d'un compteur au nom du fichier puis test existence
  lStExtFichier  := ExtractFileExt( vFichier ) ;
  lStBaseFichier := vFichier ;
  if lStExtFichier <> '' then
    lStBaseFichier := Copy( lStBaseFichier, 1, length( lStBaseFichier ) - length(lStExtFichier) ) ;

  For lInCompteur := 1 to 99 do
    begin
    if Length( IntToStr( lInCompteur ) ) = 1
       then lStNewFichier := lStBaseFichier + '0' + IntToStr( lInCompteur ) + lStExtFichier
       else lStNewFichier := lStBaseFichier + IntToStr( lInCompteur ) + lStExtFichier ;
    // On d'arrête au 1er nom de fichier non existant
    result := not FileExists( lStNewFichier ) ;
    if result then break ;
    end ;

  // Maj du nom du fichier si besoin
  if Result then
     vFichier := lStNewFichier ;

end ;
}

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 29/04/2004
Modifié le ... :   /  /
Description .. : MAJ les pièces passées en paramètres avec un numéro
Suite ........ : basé sur les compteurs "normaux".
Suite ........ : Rem :
Suite ........ : Par défaut, pour éviter les erreurs de compteurs sur les
Suite ........ : pièces "normales" le numéro des pièces générées lors du
Suite ........ : traitement est renseigné avec un numéro de compteur
Suite ........ : "Simulation"
Mots clefs ... :
*****************************************************************}
Function  UpdateNumeroPiece ( vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais, vTobScenario : TOB ) : Boolean ;
var lPieceCpt : TPieceCompta ;
    lInCpt    : Integer ;
begin
  // Par default, on considère qu'il peut y avoir erreur
  Result := False ;

  try

    // MAJ de la pièce de règlement
    for lInCpt := 0 to vTobGene.Detail.Count - 1 do
      begin
      lPieceCpt := TPieceCompta( vTobGene.Detail[ lInCpt ] ) ;
      lPieceCpt.PutEntete( 'E_QUALIFPIECE', 'N' ) ;
//        lPieceCpt.AttribNumeroTemp ;
      if V_PGI.IoError<>oeOk then Exit ;
      end ;

    // Enregistrement des pièces d'escompte
    if vTobEscompte.Detail.Count > 0 then
      for lInCpt := 0 to vTobEscompte.Detail.Count - 1 do
        begin
        lPieceCpt := TPieceCompta( vTobEscompte.Detail[ lInCpt ] ) ;
        lPieceCpt.PutEntete( 'E_QUALIFPIECE', 'N' ) ;
//        lPieceCpt.AttribNumeroTemp ;
        if V_PGI.IoError<>oeOk then Exit ;
        end ;

    if (VH^.PaysLocalisation=CodeISOES) and (vTobFrais.Detail.Count>0) then
       for lInCpt:=0 to vtobFrais.Detail.Count-1 do
         if vtobFrais.Detail[lInCpt].Detail.Count>0 then
            begin
            AffecteNumeroPieceNormal( vtobFrais.Detail[lInCpt], vtobFrais.Detail[lInCpt].Detail[0].GetValue('E_JOURNAL'), vTobScenario.GetValue('DATECOMPTABLE') ) ;
            if V_PGI.IoError<>oeOk then Exit ;
            End ;

    // Enregistrement de la deuxième pièce des remises en Escompte
    if (VH^.PaysLocalisation=CodeISOES) and (vTobRemEscompte.Detail.Count>0) then
       for lInCpt:=0 to vtobEscompte.Detail.Count-1 do
         if vtobEscompte.Detail[lInCpt].Detail.Count>0 then
            begin
            AffecteNumeroPieceNormal( vtobEscompte.Detail[lInCpt], vtobEscompte.Detail[lInCpt].Detail[0].GetValue('E_JOURNAL'), vTobScenario.GetValue('DATECOMPTABLE') ) ;
            if V_PGI.IoError<>oeOk then Exit ;
            End ; //XVI 24/02/2005

    // Tout s'est bien pasé si on arrive là :
    Result := True ;

    finally

  end;


end ;

{JP 07/06/07 : FQ 18831 : gestion des critères pour l'état}
Function  ExecuteEmissionBOR  ( vTLEche : TList ; vStModeleBOR : string ; vBoPrintDialog, vBoApercu : Boolean; Criteres : string = '') : Boolean ;
Var lBoDialog : Boolean ;
    lTobEtat  : Tob ;
    lTlParent : TList ;
begin

  Result := False ;

  if vTLEche.Count = 0 then Exit ;

  lBoDialog := False ;

  try

    lTobEtat  := Tob.Create('TOB_ETAT', nil, -1 ) ;
    lTlParent := TList.create ;
//    _permuteTob( vTLEche, lTLParent, lTobEtat ) ;
    _permuteTob( vTLEche, nil, lTobEtat ) ;

    // Gestion du Print Dialog
    lBoDialog := V_PGI.NoPrintDialog ;
    if not lBoDialog then
      V_PGI.NoPrintDialog := not vBoPrintDialog ;

    LanceEtatTOB( 'E' , 'BOR' , vStModeleBOR,  // Type, Nature , Modèle d'état
                    lTobEtat,
                    vBoApercu,
                    False, False, Nil, '', '', False, 0, Criteres ) ; {JP 07/06/07 : FQ 18831 : ajout de criteres}

    Result := True ;

    finally

      // Restauration pièces
//      _restaureTob( vTLEche, lTLParent ) ;
      _restaureTob( vTLEche, nil ) ;

      if Assigned( lTobEtat ) then
        FreeAndNil( lTobEtat ) ;
      if Assigned( lTLParent ) then
        FreeAndNil( lTLParent ) ;

      V_PGI.NoPrintDialog := lBoDialog ; // Gestion du Print Dialog

    end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 29/04/2004
Modifié le ... : 08/02/2006
Description .. : Retourne vrai si l'auxiliaire ou le général de la ligne courante
Suite ........ : ont changés par rapport à ceux de référence passé en
Suite ........ : paramètre dans vGeneral et vAuxiliaire.
Suite ........ : OU
Suite ........ : Si les lignes de tiers ne sont pas globalisés
Mots clefs ... :
*****************************************************************}
Function  RuptureEche       ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ;
                              var vGeneral : String ; var vAuxiliaire : String ; var vEtab : string ) : Boolean ;
var lBoRuptEtab : boolean ;
begin

  // Si aucune ligne a traiter, la rupture est vrai, on sort...
  result := not assigned( vTobLigne ) ;
  if result then Exit ;

  if vTobScenario.GetValue('CPG_TIERSGLOBALISE') = 'X' then
    begin

    lBoRuptEtab :=   ( ( vTobScenario.GetString('CPG_MULTIETAB')='X') and
                       ( vTobScenario.GetString('CPG_SELECTETAB')='FAC') and
                       ( vEtab <> vTobLigne.GetString('E_ETABLISSEMENT') ) ) ;

    if vInIndex>0 then
       Result := (    ( vAuxiliaire <> vTobLigne.GetValue('E_AUXILIAIRE')    )
                   or ( vGeneral    <> vTobLigne.GetValue('E_GENERAL')       )
                   or   lBoRuptEtab
                  ) ;
    end
  else
    result := True ;

  // MAJ des champs de rupture
  if Result or ( vInIndex = 0 ) then
    begin
    vAuxiliaire  := vTobLigne.GetValue('E_AUXILIAIRE') ;
    vGeneral     := vTobLigne.GetValue('E_GENERAL') ;
    vEtab        := vTobLigne.GetValue('E_ETABLISSEMENT') ;
    end ;


end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/02/2006
Modifié le ... : 07/02/2006
Description .. : Retourne vrai si le dossier de la ligne courante
Suite ........ : a changé par rapport à celui de référence passé en
Suite ........ : paramètre dans vDossier
Mots clefs ... :
*****************************************************************}
Function  RuptureDossier       ( vInIndex : Integer ; vTobLigne, vTobScenario : TOB ; var vDossier : string ) : Boolean ;
begin

  result := False ;

  if not EstMultiSoc then Exit ;
  if not assigned( vTobLigne ) then Exit ;
  if vTobLigne.GetNumChamp('SYSDOSSIER') < 1 then Exit ;

  if vInIndex>0 then
     // Changement de société
     Result := ( vDossier <> vTobLigne.GetValue('SYSDOSSIER') ) ;

  // MAJ des champs de rupture
  if Result or ( vInIndex = 0 ) then
    vDossier  := vTobLigne.GetValue('SYSDOSSIER') ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 29/04/2004
Modifié le ... :   /  /
Description .. : Teste l'uniformité des modes de paiement sur les lignes de
Suite ........ : tiers donnant lieu à une seule ligne de contrepartie.
Suite ........ : Retourne vrai si détection d'une ligne de tiers avec un mode
Suite ........ : de paiement différent que celui de référence passé en
Suite ........ : paramètre dans vModePaie
Mots clefs ... :
*****************************************************************}
Function  RuptureModePaie   ( vTobLigne, vTobScenario : TOB ; var vModePaie : String ; var vBoPremier : Boolean ) : Boolean ;
begin

  result := False ;

  // Si aucune ligne a traiter, on sort...
  if not assigned( vTobLigne ) then Exit ;

  // Si méthode de génération au détail, alors jamais besoin de maj de l'échéance
  if vTobScenario.GetValue('CPG_GROUPEENCADECA') = 'DET' then Exit ;

  // Sur la 1ère ligne, pas de test, uniquement maj du mode de paiement
  if vBoPremier then
    begin
    vModePaie  := vTobLigne.GetValue('E_MODEPAIE') ;
    vBopremier := False ;
    end
  // Sinon test uniformité du mode de paiement
  else Result := vModePaie <> vTobLigne.GetValue('E_MODEPAIE') ;

end ;

Function ConditionTypeEncadeca ( LeFlux : String ) : String ; //XMG 05/04/04
begin
  //on ne prend que les type d'enca/Déca que corresponent....
  if VH^.PaysLocalisation=CodeISOES then
     Begin
     if Trim(LeFlux)<>'' then Result:=' and CO_LIBRE in ("'+VH^.PaysLocalisation+'","'+VH^.PaysLocalisation+LeFlux+'")'
                         else Result:=' and CO_LIBRE like "'+VH^.PaysLocalisation+'%"';

     End
  else
    Result:=' and (CO_LIBRE in ("CMP","FR"))' ; //XVI 20/04/2005
end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 01/12/2005
Modifié le ... :   /  /
Description .. : Permet de virer les lignes avec montatns à 0 dans le cadre
Suite ........ : des cumuls de tiers soldés
Suite ........ : ( FQ 16986 )
Mots clefs ... :
*****************************************************************}
Function  CorrigeLigneAZero ( vTobGene, vTobEscompte, vTobRemEscompte, vTobFrais : TOB ) : Boolean ;
var lNumP : Integer ;

    function _CorrigePiece( vTobPiece : TOB ) : Boolean ;
    var lNumL   : Integer ;
    begin
      result := False ;
      for lNumL := vTobPiece.Detail.count - 1 downto 0 do
        begin
        if ( vTobPiece.Detail[ lNumL ].GetDouble( 'E_CREDITDEV' ) = vTobPiece.Detail[ lNumL ].GetDouble( 'E_DEBITDEV' ) ) then
          begin
          result := True ;
          vTobPiece.Detail[ lNumL ].Free ;
          end ;
        end ;
    end ;

begin

  result := False ;
{
  // Test Piece principale
  for lNumP := vTobGene.Detail.count - 1 downTo 0 do
    begin
    if _CorrigePiece( vTobGene.Detail[ lNumP ] ) then
      begin
      result := True ;
      if vTobGene.Detail[ lNumP ].Detail.count < 2 then
         vTobGene.Detail[ lNumP ].Free ;
      end ;
    end ;

  // Test des pièces d'escompte
  for lNumP := vTobEscompte.Detail.count - 1 downTo 0 do
    begin
    if _CorrigePiece( vTobEscompte.Detail[ lNumP ] ) then
      begin
      result := True ;
      if vTobEscompte.Detail[ lNumP ].Detail.count < 2 then
         vTobEscompte.Detail[ lNumP ].Free ;
      end ;
    end ;
}
  // Test de la pièce de remise
  if (VH^.PaysLocalisation=CodeISOES) and ( vtobRemEscompte.Detail.count>0) then
    result := result or _CorrigePiece( vTobRemEscompte ) ;

  // Test des pièces de frais
  if (VH^.PaysLocalisation=CodeISOES) then
    for lNumP := vTobFrais.Detail.count - 1 downTo 0 do
      begin
      if _CorrigePiece( vTobFrais.Detail[ lNumP ] ) then
        begin
        result := True ;
        if vTobFrais.Detail[ lNumP ].Detail.count < 2 then
           vTobFrais.Detail[ lNumP ].Free ;
        end ;
      end ;

end ;


{ TErrorEncaDeca }

class function TErrorEncaDeca.CreerMsgErreur: TErrorEncaDeca;
begin
  result := TErrorEncaDeca.Create;
  result.Initialisation ;
end;

destructor TErrorEncaDeca.Destroy;
begin
  inherited;
  if Assigned(FMsgCompta) then
    FreeAndnil( FMsgCompta ) ;
end;

procedure TErrorEncaDeca.Initialisation;
begin
  FMsgCompta := TMessageCompta.Create( 'Génération d''enca/déca', msgSaisiePiece ) ;
end;

procedure TErrorEncaDeca.OnError(sender: TObject; Error: TRecError);
begin
  if trim( Error.RC_Message )<>''
    then PGIInfo( Error.RC_Message, 'Génération d''enca/déca' )
    else if ( Error.RC_Error <> RC_PASERREUR ) then
         FMsgCompta.Execute( Error.RC_Error ) ;
end ;


end.




