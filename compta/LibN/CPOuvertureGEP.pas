{***********UNITE*************************************************
Auteur  ...... : SBO
Créé le ...... : 26/10/2005
Modifié le ... :   /  /
Description .. :
Suite ........ : Fonctions d'ouverture de gestion des pièces compta pour le
Suite ........ : GEP.
Mots clefs ... :
*****************************************************************}

{
Liste des fonctions utilisables en script depuis la compta uniquement pour le moment :

> Fonction CP_CreerPiece ( CodeJournal, CodeNature, DateComptable, CodeDevise, Etablissement )
    Retourne une pièce comptable  (en fait un pointeur sur un objet (TPieceCompta) permettant la gestion d'une pièce comptable.)
    Retourne 0 si un problème est rencontré.

> Fonction CP_AjouterEcriture( LaPiece, CompteGeneral, CompteAuxiliaire, DebitEnDevise, CreditEnDevise )
    Ajoute une ligne d'écriture à la pièce passée en paramètre.
    Retourne un code erreur >=0 si un problème est rencontré, -1 si tout est ok

> Fonction CP_ModifierChamps ( LaPiece, IndexLigne, LeNomDuChamps, LaValeur ) ;
    Modifie un champ d'une ligne de la pièce.
    Retourne un code erreur >=0 si un problème est rencontré, -1 si tout est ok

> Fonction CP_TesterPiece ( LaPiece )
    Test la validité de la pièce.
    Retourne un code erreur >=0 si un problème est rencontré, -1 si tout est ok

> Fonction CP_SauverPiece ( LaPiece )
    Enregistre la pièce en base.
    Retourne un code erreur >=0 si un problème est rencontré, -1 si tout est ok

> Fonction CP_LibererPiece ( LaPiece )
    Libère la mémoire allouer pour la gestion de la pièce lors du CP_CreerPiece.


// EXEMPLE DE FONCTION SCRIPT UTILISANT LES FONCTIONS DE GESTION DES PIECES COMPTA

procedure BVALIDERPIECE_OnClick()
var Piece, codeErr, i,
    lStJal, lStNat, lStDate, lStDev, lStEtab, // entete
    lStGen, lStAux, lStText, lStDeb, lStCred ; // ligne
begin

  lStJal  := E_Journal.Value ;  // Journal
  lStNat  := E_Nature.Value;   // Nature de la pièce  
  lStDate := E_Date.Text;      // Date comptable
  lStDev  := E_Devise.Value; // Devise
  lStEtab := E_Etab.Value ;   // Etablissement

   // Création pièce
  Piece := CP_CreerPiece ( lStJal, lStNat, lStDate, lStDev, lStEtab ) ;

  // Ajout des lignes
  i := 0 ;
  while ( i < 3 ) do  // Parcours des lignes de la grille
    begin
    i := i + 1 ; // index de la ligne courante
    lStGen := GridGetCell('FLISTE',1,i) ;  // récupération du compte général de la ligne courante
    if ( lStGen <> '' ) then
      begin
      lStAux  := GridGetCell('FLISTE',2,i) ;  // auxiliaire
      lStText := GridGetCell('FLISTE',3,i) ;  // texte libre
      lStDeb  := GridGetCell('FLISTE',4,i) ; // débit
      lStCred := GridGetCell('FLISTE',5,i) ; // crédit

      // Ajouter une ligne d'écriture
      codeErr := CP_AjouterEcriture( Piece, lStGen, lStAux, lStDeb, lStCred ) ;

      // Modififier un champ d'une ligne
      codeErr := CP_ModifierChamps( Piece, i, 'E_LIBRETEXTE0', lStText ) ;

      end ;
    end ;

  // Test de Validation de la pièce
  codeErr := CP_TesterPiece( Piece ) ;

  // Enregistrement de la pièce
  if ( codeErr < 0 ) then
    codeErr := CP_SauverPiece( Piece ) ;

  // Libération de la mémoire utilisée par la pièce
  CP_LibererPiece( Piece ) ;

end

}

unit CPOuvertureGEP;

interface

uses
  UTob,
  Forms,
  Controls,
  uLibPieceCompta,
  uLibEcriture,
  M3FP,
  MenuSpec,
  HEnt1,          // IsValidDate
  SysUtils,       // StrToDate
  hmsgbox ;



// Validation de pièce
function OnBeforeValidePieceCompta ( LaForme : TForm ; vPiece : TPieceCompta ) : boolean;
function OnAfterValidePieceCompta ( LaForme : TForm ; vPiece : TPieceCompta ) : boolean;


implementation


function OnBeforeValidePieceCompta (LaForme : TForm; vPiece : TPieceCompta ) : boolean;
var TobData     : TOB;
    TobResultat : TOB;
begin
  Result := true;

  // données
//  TobData := vPiece.DupliquerEnTob ;
  TobData := Tob.Create ('', nil, -1);
  TobData.Dupliquer( vPiece, True, True ) ;

  // Résultat
  TobResultat := Tob.Create ('', nil, -1);
  TobResultat.AddChampSupValeur ('VALIDE', 0);
  TobResultat.AddChampSupValeur ('MESSAGE', '');

  // Appel script
  vmDisp_CallProc ([LongInt(LaForme), 'CP_BeforeValidePiece', LongInt(TobData), LongInt(TobResultat)], 4);

  // Question
  if TobResultat.GetInteger ('VALIDE') = 1 then
    begin
    if TobResultat.GetString ('MESSAGE') <> '' then
      Result := PgiAsk (TobResultat.GetString ('MESSAGE')) <> mrNo ;
    end
  // Message bloquant
  else if TobResultat.GetInteger ('VALIDE') = 2 then
    begin
    PgiBox ( TobResultat.GetString ('MESSAGE') ) ;
    Result := false;
    end;

  TobData.Free;
  TobResultat.Free;

end ;

function OnAfterValidePieceCompta (LaForme : TForm; vPiece : TPieceCompta) : boolean;
var TobData     : TOB;
    TobResultat : TOB;
begin
  Result := true;

  // données
//  TobData := vPiece.DupliquerEnTob ;
  TobData := Tob.Create ('', nil, -1);
  TobData.Dupliquer( vPiece, True, True ) ;

  // résultat
  TobResultat := Tob.Create ('', nil, -1);
  TobResultat.AddChampSupValeur ('VALIDE', 0);
  TobResultat.AddChampSupValeur ('MESSAGE', '');

  // appel script
  vmDisp_CallProc ([LongInt(LaForme), 'CP_AfterValidePiece', LongInt(TobData), LongInt(TobResultat)], 4);

  // Question
  if TobResultat.GetInteger ('VALIDE') = 1 then
    begin
    if TobResultat.GetString ('MESSAGE') <> '' then
      Result := PgiAsk (TobResultat.GetString ('MESSAGE')) <> mrNo ;
    end
  // Message bloquant
  else if TobResultat.GetInteger ('VALIDE') = 2 then
    begin
    PgiBox ( TobResultat.GetString ('MESSAGE') ) ;
    Result := false;
    end;

  TobData.Free;
  TobResultat.Free;
end;



function CCreerPiece (Parms : array of Variant; nb: integer ) : variant ;
var Piece       : TPieceCompta;
    lInfoEcr    : TInfoEcriture ;
    lStJournal  : String ;
    lStNature   : String ;
    lStDateC    : String ;
    lStDevise   : String ;
    lStEtab     : String ;
begin

  result := 0 ;

  try

    // Récupération des paramètres
    lStJournal   := Parms[0] ;
    lStNature    := Parms[1] ;
    lStDateC     := Parms[2] ;
    lStDevise    := Parms[3] ;
    lStEtab      := Parms[4] ;

    lInfoEcr := TInfoEcriture.Create ;
    Piece    := TPieceCompta.CreerPiece( lInfoEcr ) ;

    // Le journal
    if lStJournal <> '' then
      Piece.PutEntete( 'E_JOURNAL', lStJournal )
      else begin
           PgiBox( 'Le journal n''est pas renseigné !', 'CP_CreerPiece' ) ;
           Exit ;
           end ;

    // La Nature
    if lStNature <> '' then
      Piece.PutEntete( 'E_NATUREPIECE', lStNature )
      else begin
           PgiBox( 'La nature n''est pas renseignée !', 'CP_CreerPiece' ) ;
           Exit ;
           end ;

    // La date
    if ( lStDateC <> '' ) and IsValidDate( lStDateC ) then
      Piece.PutEntete( 'E_DATECOMPTABLE', StrToDate( lStDateC ) )
      else begin
           PgiBox( 'Le date comptable n''est pas renseignée ou incorrecte !', 'CP_CreerPiece' ) ;
           Exit ;
           end ;

    // La devise
    if lStDevise <> '' then
      Piece.PutEntete( 'E_DEVISE', lStDevise )
      else begin
           PgiBox( 'La devise n''est pas renseignée !', 'CP_CreerPiece' ) ;
           Exit ;
           end ;

    // L'établissement
    if lStEtab <> '' then
      Piece.PutEntete( 'E_ETABLISSEMENT', lStEtab )
      else begin
           PgiBox( 'L''établissement n''est pas renseigné !', 'CP_CreerPiece' ) ;
           Exit ;
           end ;

    // Attribution numéro temporaire
    Piece.AttribNumeroTemp ;

    result := LongInt( Piece ) ;

    except
      PgiBox( 'Une erreur est survenue pendant l''utilisation de la fonction. Veuillez vérifier votre SCRIPT.', 'CP_CreerPiece' ) ;

  end ;

end ;

function CAjoutEcrPiece (Parms : array of Variant; nb: integer ) : variant ;
var lTobLigne  : Tob ;
    Piece      : TPieceCompta;
    lNumLigne  : Integer ;
    lStGeneral : String ;
    lStAuxi    : String ;
    lDebit     : Double ;
    lCredit    : Double ;
begin

  result := RC_PASERREUR ;

  try

    if LongInt( Parms[0] ) = 0 then
      begin
      PgiBox( 'La référence à la pièce courante n''est pas renseignée !', 'CP_AjouterEcriture' ) ;
      Exit ;
      end ;

    // Récup paramètres
    Piece := TPieceCompta( LongInt( Parms[0] ) );
    lStGeneral := Parms[1] ;
    lStAuxi    := Parms[2] ;
    lDebit     := Valeur(Parms[3]) ;
    lCredit    := Valeur(Parms[4]) ;

    lTobLigne := Piece.NewRecord ;
    lNumLigne := lTobLigne.GetInteger('E_NUMLIGNE') ;

    // général
    if lStGeneral <> '' then
      Piece.PutValue( lNumLigne, 'E_GENERAL', lStGeneral )
      else begin
           PgiBox( 'Le compte général n''est pas renseigné !', 'CP_AjouterEcriture' ) ;
           result := RC_COMPTEINEXISTANT ;
           Exit ;
           end ;

    // auxiliaire
    if Piece.Info.Compte.IsCollectif then
      if lStAuxi <> '' then
        Piece.PutValue( lNumLigne, 'E_AUXILIAIRE', lStAuxi )
        else begin
             PgiBox( 'Le compte auxiliaire n''est pas renseigné !', 'CP_AjouterEcriture' ) ;
             result := RC_AUXOBLIG ;
             Exit ;
             end ;

      // Montants
      if lDebit <> 0
        then Piece.PutValue( lNumLigne, 'E_DEBITDEV', lDebit )
        else if lCredit <> 0
          then Piece.PutValue( lNumLigne, 'E_CREDITDEV', lCredit )
          else begin
             PgiBox( 'Le montant n''est pas renseigné !', 'CP_AjouterEcriture' ) ;
             result := RC_MONTANTINEXISTANT ;
             Exit ;
             end ;

    except
      result := RC_BADWRITE ;
      PgiBox( 'Une erreur est survenue pendant l''utilisation de la fonction. Veuillez vérifier votre SCRIPT.', 'CP_AjouterEcriture' ) ;

  end ;

end ;

function CPutValuePiece (Parms : array of Variant; nb: integer ) : variant ;
var Piece      : TPieceCompta;
    lInNumL    : Integer ;
    lStChamps  : String ;
    lStValeur  : Variant ;
begin

  result := RC_PASERREUR ;

  try

    if LongInt( Parms[0] ) = 0 then
      begin
      PgiBox( 'La référence à la pièce courante n''est pas renseignée !', 'CP_ModifierChamps' ) ;
      Exit ;
      end ;

    // Récup paramètres
    Piece := TPieceCompta( LongInt( Parms[0] ) );
    lInNumL      := Parms[1] ;
    lStChamps    := Parms[2] ;
    lStValeur    := Parms[3] ;

    // Modif champs
    if lStChamps <> '' then
      Piece.PutValue( lInNumL, lStChamps, lStValeur )
      else begin
           PgiBox( 'Le nom du champs n''est pas renseigné !', 'CP_ModifierChamps' ) ;
           result := RC_BADWRITE ;
           Exit ;
           end ;

    result := LongInt( Piece ) ;

    except
      result := RC_BADWRITE ;
      PgiBox( 'Une erreur est survenue pendant l''utilisation de la fonction. Veuillez vérifier votre SCRIPT.', 'CP_ModifierChamps' ) ;

  end ;

end ;


function CTesterPiece (Parms : array of Variant; nb: integer ) : variant ;
var  Piece   : TPieceCompta;
begin

  result := RC_PASERREUR ;

  try

    if LongInt( Parms[0] ) = 0 then
      begin
      PgiBox( 'La référence à la pièce courante n''est pas renseignée !', 'CP_TesterPiece' ) ;
      Exit ;
      end ;

    // Récup paramètres
    Piece := TPieceCompta( LongInt( Parms[0] ) );

    // La pièce est-elle valide
    if not Piece.isValidPiece
      then result := Piece.LastError.RC_Error ;

    except
      result := RC_BADWRITE ;
      PgiBox( 'Une erreur est survenue pendant l''utilisation de la fonction. Veuillez vérifier votre SCRIPT.', 'CP_TesterPiece' ) ;

  end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 13/02/2008
Modifié le ... :   /  /    
Description .. : Prise en compte demande VB sur FQ 22315 :
Suite ........ : En terme de méthodologie, il faudra prévoir une Tob de
Suite ........ : retour avec trois champs :
Suite ........ : - VALIDE : l'indicateur d'échec/réussite
Suite ........ : - MESSAGE : le message d'erreur en cas d'échec
Suite ........ : - NUMERO : le numéro de la pièce créée en cas de réussite
Suite ........ : Cette Tob :
Suite ........ : - sera un second paramètre de CP_SauverPiece( Piece, 
Suite ........ : TobR)
Suite ........ : - est créé et détruite par le script
Mots clefs ... : 
*****************************************************************}
function CSauverPiece (Parms : array of Variant; nb: integer ) : variant ;
var  Piece   : TPieceCompta;
     lTobRes : Tob ;        // Tob résultat
     lMess   : TMessageCompta ;
begin

  result  := RC_PASERREUR ;
  lTobRes := nil ;

  try

    if LongInt( Parms[0] ) = 0 then
      begin
      PgiBox( 'La référence à la pièce courante n''est pas renseignée !', 'CP_SauverPiece' ) ;
      Exit ;
      end ;

    if LongInt( Parms[1] ) = 0 then
      begin
      PgiBox( 'La Tob résultat n''est n''est pas renseignée !', 'CP_SauverPiece' ) ;
      Exit ;
      end ;

    // Récup paramètres
    Piece   := TPieceCompta( LongInt( Parms[0] ) ) ;
    lTobRes := TOB( LongInt( Parms[1] ) );
    // ajout indicateur reussite
    if lTobRes.GetNumChamp( 'VALIDE' ) < 0 then
       lTobRes.AddChampSupValeur( 'VALIDE', '-' ) ;
    // ajout message si echec
    if lTobRes.GetNumChamp( 'MESSAGE' ) < 0 then
       lTobRes.AddChampSupValeur( 'MESSAGE', '' ) ;
    // ajout numero de piece
    if lTobRes.GetNumChamp( 'NUMERO' ) < 0 then
       lTobRes.AddChampSupValeur( 'NUMERO', 0) ;

    // La pièce est-elle valide
    if Piece.isValidPiece then
      begin
      if Piece.Save then
        begin
        lTobRes.PutValue('VALIDE',  'X' ) ;
        lTobRes.PutValue('MESSAGE', '' ) ;
        lTobRes.PutValue('NUMERO',  Piece.GetEntete('E_NUMEROPIECE') ) ;
        end ;
      end
    else
      begin
      result := Piece.LastError.RC_Error ;
      lMess := TMessageCompta.Create( 'CP_SauverPiece', msgSaisiePiece ) ;
      lTobRes.PutValue('VALIDE',  '-' ) ;
      lTobRes.PutValue('MESSAGE', lMess.GetMessage( result ) ) ;
      lTobRes.PutValue('NUMERO',  0 ) ;
      FreeAndNil( lMess ) ;
      Exit ;
      end ;

    except
      result := RC_BADWRITE ;
      if Assigned( lTobRes ) then
        begin
        lTobRes.PutValue('VALIDE',  '-' ) ;
        lTobRes.PutValue('MESSAGE', 'Une erreur est survenue pendant l''utilisation de la fonction. Veuillez vérifier votre SCRIPT.' ) ;
        lTobRes.PutValue('NUMERO',  0 ) ;
        end ;
      PgiBox( 'Une erreur est survenue pendant l''utilisation de la fonction. Veuillez vérifier votre SCRIPT.', 'CP_SauverPiece' ) ;

  end ;

end;

procedure CLibererPiece (Parms : array of Variant; nb: integer ) ;
var  Piece   : TPieceCompta;
begin

  try

    if LongInt( Parms[0] ) = 0 then
      begin
      PgiBox( 'La référence à la pièce courante n''est pas renseignée !', 'CP_LibererPiece' ) ;
      Exit ;
      end ;

    // Récup paramètres
    Piece := TPieceCompta( LongInt( Parms[0] ) );

    // Libération mémoire
    Piece.Free ;

    except
      PgiBox( 'Une erreur est survenue pendant l''utilisation de la fonction. Veuillez vérifier votre SCRIPT.', 'CP_LibererPiece' ) ;

  end ;

end ;


procedure InitAglFunc;
begin
  RegisterAglFunc ('CP_TesterPiece',      False , 1, CTesterPiece);
  RegisterAglFunc ('CP_SauverPiece',      False , 1, CSauverPiece);
  RegisterAglProc ('CP_LibererPiece',     False , 1, CLibererPiece);
  RegisterAglFunc ('CP_CreerPiece',       False , 5, CCreerPiece);
  RegisterAglFunc ('CP_AjouterEcriture',  False , 5, CAjoutEcrPiece);
  RegisterAglFunc ('CP_ModifierChamps',   False , 4, CPutValuePiece);
end;

Initialization
  InitAglFunc
  ;

end.
