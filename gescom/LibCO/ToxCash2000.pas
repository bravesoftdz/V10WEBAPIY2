{***********UNITE*************************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Génération des fichiers CASH 2000 à partir des TOX
Mots clefs ... : TOX;CASH 2000
*****************************************************************}

{****************************************************************
 * PROBLEMES
 * ---------
 *  1/ Codes TVA différents par boutique
 ****************************************************************}
 
unit ToxCash2000;

interface

uses {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hctrls, Windows, SysUtils, uTob, StdCtrls, Dialogs, FactComm, FactUtil, FactCpta, 
     SaisUtil, HEnt1, FactNomen, ToxMessage, UtilPGI, Ent1, FactCalc, ParamSoc, EntGC,
{$IFDEF TOXCLIENT}
     uToxFtp,
{$ENDIF}
     TarifUtil, SaisieCorrespGPAO ;

//////////////////////////////////////////////////////////////////////////////
// Liste des fonctions permettant de contrôler ou d'intégrer un enregistrement
///////////////////////////////////////////////////////////////////////////////

procedure TOBArrondiToCash2000     (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBArticleToCash2000     (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;
procedure TOBChancellToCash2000    (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBChoixcodToCash2000    (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBChoixExtToCash2000    (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBCommercialToCash2000  (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;
procedure TOBDeviseToCash2000      (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBDimensionToCash2000   (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBDispoToCash2000       (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBEtablissToCash2000    (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;
procedure TOBModePaieToCash2000    (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBPaysToCash2000        (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBCodePostalToCash2000  (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBPieceToCash2000       (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBTarifToCash2000       (ToxInfo : TOB; var FichierBE : textfile) ;
procedure TOBTiersToCash2000       (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;

procedure TOBMarchandiseToCash2000 (TobDonnees : TOB; var FichierBE : textfile; CodeSite : string) ;
procedure TOBPrestationToCash2000  (TobDonnees : TOB; var FichierBE : textfile) ;
procedure TOBArtFinanToCash2000    (TobDonnees : TOB; var FichierBE : textfile) ;

procedure TOBFinitionToCash2000    (CodeParam, LibParam : String; var FichierBE : textfile) ;
procedure TOBModeleToCash2000      (CodeParam, LibParam : String; var FichierBE : textfile) ;
procedure TOBColorisToCash2000     (CodeParam, LibParam : String; var FichierBE : textfile) ;
procedure TOBMatiereToCash2000     (CodeParam, LibParam : String; var FichierBE : textfile) ;
procedure TOBFamilleArtToCash2000  (CodeParam, LibParam : String; var FichierBE : textfile) ;
procedure TOBGrilleTailleToCash2000(CodeGrille : String; CodeTaille : Array of String; var FichierBE : textfile) ;
procedure TOBTVAToCash2000         (CodeTVA : String; TauxTVA : Double; var FichierBE : textfile) ;

procedure CreeTobParamCorresp ;
procedure LibereTobParamCorresp ;
function  ToxNomFichierBE ( CodeSite, NumeroCaisse : String ; NumJour : Integer ) : String ;

procedure ToxToCash2000 (NomTable, CodeSite : string; ToxInfo : TOB; var FichierBE : textfile) ;
procedure AvantArchivageToxPourCASH (Tox : TOB ; CodeSite, LesCaisses : String) ;
procedure DefineChampTobParam (TobParametrage : TOB) ;
procedure ChargeTobParam (TobParametrage : TOB) ;

const
 ////////////////////////////////////////////////////////////////////
 // Noms des champs de la TOB Tob_param, avec leur valeur par défaut
 /////////////////////////////////////////////////////////////////////
 ChampTobParam : array[1..105,1..2] of String 	= (
   ('OPTIMISATIONARTICLE', '0'),
   ('RECODIFIEFAMILLE'  , '0'),
   ('FAMNIV1'           , '3'),
   ('FAMNIV2'           , '3'),
   ('FAMNIV3'           , '0'),
   ('REPFAMILLE'        , ''),
   ('REPMOD'            , ''),
   ('REPMAT'            , ''),
   ('REPCOL'            , ''),
   ('REPFIN'            , ''),
   ('REPSTR'            , ''),
   ('REPLIBCOL'         , ''),
   ('REPCAT'            , ''),
   ('REPFOR'            , ''),
   ('REPCP1'            , ''),
   ('REPCP2'            , ''),
   ('REPTYP'            , ''),
   ('CATEGORIEDIM'      , ''),
   ('GRILLEUNI'         , 'UNI'),
   ('REPCODEBARRE'      , ''),
   ('REPTVA'            , ''),
   ('REPREG'            , ''),
   ('CODECLI'           , '0'),
   ('REPCLIPROMAIL'     , ''),
   ('CLIEXPORT'         , ''),
   ('REPETACPTA'        , ''),
   ('FRNFAMCOMPTA'      , ''),
   ('FRNREGIMETVA'      , ''),
   ('NATCDE'            , ''),
   ('NATREC'            , ''),
   ('NATENT'            , ''),
   ('NATSOR'            , ''),
   ('NATTRFEM'          , ''),
   ('NATTRFRE'          , ''),
   ('NATNEGCDE'         , ''),
   ('NATNEGLIV'         , ''),
   ('NATNEGFAC'         , ''),
   ('NATNEGAVO'         , ''),
   ('NATVTE'            , ''),
   ('MODEREGLVTE'       , ''),
   ('CODECLIVTE'        , ''),
   ('VALSTOCK'          , '0'),
   ('TARIF'             , '0'),
   ('TARIFHT'           , '0'),
   ('TYPETARIFTTC'      , '' ),
   ('PERIODETARIFTTC'   , '' ),
   ('TYPETARIFSOLDE'    , '' ),
   ('PERIODETARIFSOLDE' , '' ),
   ('TARIFACH'          , '0'),
   ('TYPETARIFACH'      , '' ),
   ('PERIODETARIFACH'   , '' ),
   ('DELBOUTIQUES'      , ''),
   ('DELDEVISES'        , '0'),
   ('DELARRONDIS'       , '0'),
   ('DELPAYS'           , '0'),
   ('DELCODESPOSTAUX'   , '0'),
   ('DELFAMILLES'       , '0'),
   ('DELCOLLECTIONS'    , '0'),
   ('DELPOIDS'          , '0'),
   ('DELDEMARQUES'      , '0'),
   ('DELCIVILITES'      , '0'),
   ('DELREPRESENTANTS'  , '0'),
   ('DELCATEGORIES'     , '0'),
   ('DELGRILLEDIM1'     , '0'),
   ('DELDIM1'           , '0'),
   ('DELGRILLEDIM2'     , '0'),
   ('DELDIM2'           , '0'),
   ('DELGRILLEDIM3'     , '0'),
   ('DELDIM3'           , '0'),
   ('DELGRILLEDIM4'     , '0'),
   ('DELDIM4'           , '0'),
   ('DELGRILLEDIM5'     , '0'),
   ('DELDIM5'           , '0'),
   ('DELMASQUEDIM'      , '0'),
   ('DELTABLEA1'        , '0'),
   ('DELTABLEA2'        , '0'),
   ('DELTABLEA3'        , '0'),
   ('DELTABLEA4'        , '0'),
   ('DELTABLEA5'        , '0'),
   ('DELTABLEA6'        , '0'),
   ('DELTABLEA7'        , '0'),
   ('DELTABLEA8'        , '0'),
   ('DELTABLEA9'        , '0'),
   ('DELTABLEA10'       , '0'),
   ('DELTABLEC1'        , '0'),
   ('DELTABLEC2'        , '0'),
   ('DELTABLEC3'        , '0'),
   ('DELTABLEC4'        , '0'),
   ('DELTABLEC5'        , '0'),
   ('DELTABLEC6'        , '0'),
   ('DELTABLEC7'        , '0'),
   ('DELTABLEC8'        , '0'),
   ('DELTABLEC9'        , '0'),
   ('DELTABLEC10'       , '0'),
   ('DELARTICLES'       , '0'),
   ('DELCLIENTS'        , '0'),
   ('DELFOURNISSEURS'   , '0'),
   ('STOCKCLOTURE'      , 'N'),
   ('DTECLOTURE'        , '01/01/1900'),
   ('REPPAARTICLE'      , 'N'),
   ('REPPVARTICLE'      , 'N'),
   ('REPPVHTARTICLE'    , 'N'),
   ('REPPRARTICLE'      , 'N'),
   ('REPDEVISE'         , '' ),
   ('CODEARTVTE'        , '' ));

var
  Tob_Param       	: TOB    ;  // TOB utilisée pour stocker le paramétrage de la récup
  Tob_Paiement          : TOB    ;  // TOB de correspondance des modes de paiement
  Tob_Pays              : TOB    ;  // TOB de correspondance des pays
  Tob_Devise            : TOB    ;  // TOB des correspondance des devises
  Tob_Famille           : TOB    ;  // TOB des correspondances des familles
  Tob_TVA               : TOB    ;  // TOB des correspondances des TVA

  TobArticleGB          : TOB    ;  // TOB listant les articles à envoyer
  TobFamilleGB          : TOB    ;  // TOB listant les familles à pré-envoyer
  TobModeleGB           : TOB    ;  // TOB listant les modèles à pré-envoyer
  TobMatiereGB          : TOB    ;  // TOB listant les matières à pré-envoyer
  TobColorisGB          : TOB    ;  // TOB listant les coloris à pré-envoyer
  TobFinitionGB         : TOB    ;  // TOB listant les finitions à pré-envoyer
  TobGrilleGB           : TOB    ;  // TOB listant les grilles à pré-envoyer

implementation


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lionel Meunier
Créé le ...... : 21/03/2002
Modifié le ... :   /  /    
Description .. : Création des TOB des paramétrages GB/PGI et des tables
Suite ........ : de correspondances
Suite ........ : Chargement des TOB
Mots clefs ... : 
*****************************************************************}
procedure CreeTobParamCorresp ;
begin
  // Création des TOB contenant le paramétrage des récupération
  Tob_Param := TOB.CREATE ('Les Paramètres', NIL, -1);
  Tob_Param.initValeurs;
  // Création de la TOB de correspondance des règlements
  Tob_Paiement := TOB.CREATE ('Les Paiements', NIL, -1);
  Tob_Paiement.initValeurs;
  // Création de la TOB de correspondance des pays
  Tob_Pays := TOB.CREATE ('Les Pays', NIL, -1);
  Tob_Pays.initValeurs;
  // Création de la TOB de correspondance des pays
  Tob_Devise := TOB.CREATE ('Les Devises', NIL, -1);
  Tob_Devise.initValeurs;
  // Création de la TOB de correspondance des familles
  Tob_Famille := TOB.CREATE ('Les Familles', NIL, -1);
  Tob_Famille.initValeurs;
  // Création de la TOB de correspondance des TVA
  Tob_TVA := TOB.CREATE ('Les TVA', NIL, -1);
  Tob_TVA.initValeurs;
  //
  // Chargement en TOB
  //
  DefineChampTobParam (Tob_Param) ;
  ChargeTobParam (Tob_Param) ;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lionel Meunier
Créé le ...... : 21/03/2002
Modifié le ... :   /  /    
Description .. : Libération des TOB des paramétrages GB/PGI et des tables
Suite ........ : de correspondances
Mots clefs ... : 
*****************************************************************}
procedure LibereTobParamCorresp ;
begin
  Tob_Param.free;
  Tob_Param := nil ;
  Tob_Paiement.free;
  Tob_Paiement := nil ;
  Tob_Pays.free ;
  Tob_Pays := nil ;
  Tob_Devise.free ;
  Tob_Devise := nil ;
  Tob_Famille.free ;
  Tob_Famille := nil ;
  Tob_TVA.Free ;
  Tob_TVA := nil ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lionel Meunier
Créé le ...... : 21/03/2002
Modifié le ... :   /  /    
Description .. : Création des TOB listant les paramètres à
Suite ........ : pré-envoyer et les articles 
Mots clefs ... : 
*****************************************************************}
procedure CreeTobEnvoiGB ;
begin
  // Création de la TOB permettant de lister les articles envoyer
  TobArticleGB := TOB.CREATE ('Les Articles', NIL, -1);
  // Création de la TOB permettant de lister les familles à pré-envoyer
  TobFamilleGB := TOB.CREATE ('Les Familles', NIL, -1);
  // Création de la TOB permettant de lister les Modeles à pré-envoyer
  TobModeleGB := TOB.CREATE ('Les Modeles', NIL, -1);
  // Création de la TOB permettant de lister les Matieres à pré-envoyer
  TobMatiereGB := TOB.CREATE ('Les Matieres', NIL, -1);
  // Création de la TOB permettant de lister les Coloris à pré-envoyer
  TobColorisGB := TOB.CREATE ('Les Coloris', NIL, -1);
  // Création de la TOB permettant de lister les fintions à pré-envoyer
  TobFinitionGB := TOB.CREATE ('Les Finitions', NIL, -1);
  // Création de la TOB permettant de lister les grilles à pré-envoyer
  TobGrilleGB := TOB.CREATE ('Les Grilles', NIL, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lionel Meunier
Créé le ...... : 21/03/2002
Modifié le ... :   /  /
Description .. : Libération des TOB des paramétrages GB/PGI et des tables
Suite ........ : de correspondances
Mots clefs ... : 
*****************************************************************}
procedure LibereTobEnvoiGB ;
begin
  TobArticleGB.free  ;
  TobFamilleGB.free  ;
  TobModeleGB.free   ;
  TobMatiereGB.free  ;
  TobColorisGB.free  ;
  TobFinitionGB.free ;
  TobGrilleGB.free   ;

  TobArticleGB   := nil ;
  TobFamilleGB   := nil ;
  TobModeleGB    := nil ;
  TobMatiereGB   := nil ;
  TobColorisGB   := nil ;
  TobFinitionGB  := nil ;
  TobGrilleGB    := nil ;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 26/04/2002
Description .. : Conversion d'un booléen PGI (X/-) en booléen GB (O/N)
Mots clefs ... : TOX
*****************************************************************}
function BoolPGIToGB (Valeur : String) : String ;
begin
  if Valeur = 'X' then Result := 'O' else Result := 'N' ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 26/04/2002
Description .. : Conversion d'un code dimension PGI (001-ZZZ)
Suite ........ : en rang GB (01-30)
Mots clefs ... : TOX
*****************************************************************}
function CodeDimPGIToGB (Valeur : String) : Integer ;
var Stg : String ;
begin
  Result := 0 ;
  Stg := Copy(Valeur, Length(Valeur), 1) ;
  if Length(Stg) > 0 then
  begin
   if Stg[1] in ['A'..'Z'] then Result := (Ord(Stg[1]) - Ord('A')) + 10 else
   if Stg[1] in ['1'..'9'] then Result := (Ord(Stg[1]) - Ord('1')) + 1 ;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 26/04/2002
Description .. : Conversion d'un code grille de taxe et d'un code dimension
Suite ........ : PGI en un code grille et un code taille GB
Mots clefs ... : TOX
*****************************************************************}
procedure CodeDimensionPGIToTailleGB (TobDonnees : TOB ; var GrilleTailleGB : string ; var RangTailleGB : integer ; var FichierBE : textfile ) ;
var Rang       : String ;
    TailleGB   : Array of String ;
    StatutArt  : String ;
    sSQL       : String ;
    Ind        : Integer ;
    MaxInd     : Integer ;
    Tob_Taille : TOB ;
    TobFille   : TOB ;
begin
  StatutArt := TobDonnees.GetValue ('GA_STATUTART') ;
  if StatutArt = 'DIM' then
  begin
    // article dimensionné
    Rang := Copy (Tob_Param.GetValue ('CATEGORIEDIM'), 3, 1) ;
    Ind := StrToInt (Rang) ;
    if (Ind < 1) or (Ind > 5) then Rang := '1' ;
    RangTailleGB := CodeDimPGIToGB(TobDonnees.GetValue('GA_CODEDIM'+ Rang));
    GrilleTailleGB := TobDonnees.GetValue('GA_GRILLEDIM'+ Rang);
  end else
  begin
    // article unique
    RangTailleGB := 0;
    GrilleTailleGB := Tob_Param.GetValue ('GRILLEUNI') ;
  end;
  // Génération dans le fichier BE de la ligne correspondante
  if TobGrilleGB.FindFirst(['CODE'], [GrilleTailleGB], False) = Nil then
  begin
    TobFille := TOB.Create ('', TobGrilleGB, -1) ;
    TobFille.AddChampSupValeur ('CODE', GrilleTailleGB) ;
    if StatutArt = 'DIM' then
    begin
      // article dimensionné
      sSQL := 'Select GDI_LIBELLE From DIMENSION Where GDI_TYPEDIM="DI'+ Rang +'" ' +
              'And GDI_GRILLEDIM="'+ GrilleTailleGB +'" Order By GDI_CODEDIM' ;
      Tob_Taille := TOB.Create('', Nil, -1) ;
      Tob_Taille.LoadDetailFromSQL(sSQL) ;
      MaxInd := Tob_Taille.Detail.Count ;
      if MaxInd > 30 then MaxInd := 30 ;
      SetLength(TailleGB, MaxInd) ;
      for Ind := 0 to MaxInd -1 do
      begin
        TailleGB[Ind] := Tob_Taille.Detail[Ind].GetValue('GDI_LIBELLE') ;
      end;
      Tob_Taille.Free ;
    end else
    begin
      // article unique
      SetLength(TailleGB, 1) ;
      TailleGB[0] := 'UNI' ;
    end;
    TOBGrilleTailleToCash2000 (GrilleTailleGB, TailleGB, FichierBE) ;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 26/04/2002
Description .. : Conversion des codes familles article PGI en un famille
Suite ........ : article GB
Mots clefs ... : TOX
*****************************************************************}
function CodeFamillArtPGIToGB (TobDonnees : TOB ; var FichierBE : textfile ) : String ;
var TobFille         : TOB    ;
    CodeFam1PGI      : string ;
    CodeFam2PGI      : string ;
    CodeFam3PGI      : string ;
    CodeFamilleGB    : string ;
    LibeFamilleGB    : string ;
    TypeFamille      : string ;
    Fniv1            : string ;
    Fniv2            : string ;
    Fniv3            : string ;
    RecodifieFamille : string ;
    Stg              : string ;
    sSQL             : string ;
    QQ               : TQuery ;
begin
  CodeFamilleGB := '' ;
  LibeFamilleGB := '' ;
  RecodifieFamille := Tob_Param.GetValue ('RECODIFIEFAMILLE');
  if RecodifieFamille = 'X' then
  begin     // recodification des familles
    CodeFam1PGI := TobDonnees.GetValue ('GA_FAMILLENIV1');
    CodeFam2PGI := TobDonnees.GetValue ('GA_FAMILLENIV2');
    CodeFam3PGI := TobDonnees.GetValue ('GA_FAMILLENIV3');
    // Recherche des codes et libellés GB dans la table de correspondance
    Stg := TraduitPGIenGPAO(TypeRepriseGB2000, '$$_CODEFAMSF', [CodeFam1PGI,CodeFam2PGI,CodeFam3PGI], True, Tob_Famille);
    if Stg <> '' then
    begin
      CodeFamilleGB := ReadTokenST(Stg);
      LibeFamilleGB := ReadTokenST(Stg);
    end;
  end else
  begin     // Familles non recodifiées
    Fniv1 := Tob_Param.GetValue ('FAMNIV1') ;
    Fniv2 := Tob_Param.GetValue ('FAMNIV2') ;
    Fniv3 := Tob_Param.GetValue ('FAMNIV3') ;
    if StrToInt (Fniv1) > 0 then CodeFamilleGB := copy (TobDonnees.GetValue ('GA_FAMILLENIV1'), 1,  StrToInt (Fniv1));
    if StrToInt (Fniv2) > 0 then CodeFamilleGB := CodeFamilleGB + copy (TobDonnees.GetValue ('GA_FAMILLENIV2'), 1,  StrToInt (Fniv2));
    if StrToInt (Fniv3) > 0 then CodeFamilleGB := CodeFamilleGB + copy (TobDonnees.GetValue ('GA_FAMILLENIV3'), 1,  StrToInt (Fniv3));
  end;
  //
  // Génération dans le fichier BE de la ligne correspondante
  //
  if (CodeFamilleGB <> '') and (TobFamilleGB.FindFirst(['CODE'], [CodeFamilleGB], False) = Nil) then
  begin
    TobFille := TOB.Create ('', TobFamilleGB, -1) ;
    TobFille.AddChampSupValeur ('CODE', CodeFamilleGB) ;
    // recherche du libellé de la famille
    if RecodifieFamille <> 'X' then
    begin
      if StrToInt (Fniv3) > 0      then TypeFamille := 'FN3'
      else if StrToInt (Fniv2) > 0 then TypeFamille := 'FN2'
      else TypeFamille := 'FN1';
      sSQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="'+TypeFamille+'" AND CC_CODE="'+CodeFamilleGB+'"';
      QQ := OpenSQL(sSQL, True) ;
      if not QQ.EOF then LibeFamilleGB := QQ.FindField('CC_LIBELLE').AsString ;
      Ferme(QQ) ;
    end;
    if LibeFamilleGB = '' then LibeFamilleGB := CodeFamilleGB ;
    TOBFamilleArtToCash2000 (CodeFamilleGB, LibeFamilleGB, FichierBE) ;
  end;
  Result := CodeFamilleGB ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 26/04/2002
Description .. : Conversion d'un code régime de taxe et d'un code famille
Suite ........ : de taxe PGI en un code TVA GB pour un article
Mots clefs ... : TOX
*****************************************************************}
function CodeTaxePGIToTVAGB (TobDonnees : TOB ; var FichierBE : textfile ) : String ;
var Stg           : string ;
    NomChamp      : string ;
    CodeRegimePGI : string ;
    CodeTaxePGI   : string ;
    TauxTVA       : double ;
    TOBL          : TOB    ;
begin
  Result := '' ;
  TobL := Nil ;
  // recherche du régime de taxe
  CodeRegimePGI := GetParamSoc('SO_REGIMEDEFAUT') ;   // ??? problème des TVA différentes par boutique ???
  // recherche de la famille de taxe
  CodeTaxePGI := '' ;
  Stg := GetParamSoc('SO_GCDEFCATTVA') ;
  NomChamp := 'GA_FAMILLETAXE' + Copy(Stg, 3, 1) ;
  if TobDonnees.FieldExists (NomChamp) then CodeTaxePGI := TobDonnees.GetValue (NomChamp);
  if (CodeTaxePGI <> '') and (CodeRegimePGI <> '') then
  begin
    TOBL := Tob_TVA.FindFirst(['GRO_TYPEREPRISE','GRO_NOMCHAMP','GRO_VALEURPGI', 'GRO_VALEURPGI1'],
                              [TypeRepriseGB2000, ExtractSuffixe('$$_CODETVA'), CodeRegimePGI, CodeTaxePGI], False) ;
    Result := TraduitPGIenGPAO(TypeRepriseGB2000, '$$_CODETVA', [CodeRegimePGI,CodeTaxePGI], False, Tob_TVA);
  end;
  // Génération dans le fichier BE de la ligne correspondante
  if TobL = Nil then
  begin
    TauxTVA := TVA2TAUX(CodeRegimePGI, CodeTaxePGI, False);
    TOBTVAToCash2000 (Result, TauxTVA, FichierBE) ;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 26/04/2002
Description .. : Génération dans le fichier BE de la ligne de la tablette
Suite ........ : associée à un champ de la fiche article
Mots clefs ... : TOX
*****************************************************************}
procedure AjoutChoixCod ( ChampParam, CodeParam : String ; var FichierBE : textfile ) ;
var NomChamp     : String ;
    TypeChamp    : String ;
    NomTablette  : String ;
    LibParam     : String ;
    TOBParamGB   : TOB ;
    TobFille     : TOB ;
begin
  if (ChampParam <> '') and (Pos(';'+ChampParam+';', ';REPMOD;REPMAT;REPCOL;REPFIN;') > 0) and (CodeParam <> '') then
  begin
    LibParam := '' ;
    NomChamp := Tob_Param.GetValue (ChampParam) ;
    if NomChamp <> '' then TypeChamp := ChampToType (NomChamp) ;
    if (Copy(TypeChamp, 1, 7) = 'VARCHAR') or (TypeChamp = 'COMBO') then
    begin
      if ChampParam = 'REPFIN' then TOBParamGB := TobFinitionGB else
      if ChampParam = 'REPMOD' then TOBParamGB := TobModeleGB   else
      if ChampParam = 'REPCOL' then TOBParamGB := TobColorisGB  else
      if ChampParam = 'REPMAT' then TOBParamGB := TobMatiereGB  else TOBParamGB := Nil ;
      if TOBParamGB <> Nil then
      begin
        // recherche si le code a déjà été traité
        if TOBParamGB.FindFirst(['CODE'], [CodeParam], False) <> Nil then Exit ;
        TobFille := TOB.Create('', TOBParamGB, -1) ;
        TobFille.AddChampSupValeur('CODE', CodeParam) ;
      end;
      // recherche du libellé
      NomTablette := Get_Join(NomChamp) ;
      if NomTablette <> '' then LibParam := RechDom(NomTablette, CodeParam, False) ;
      if (LibParam <> '') and (LibParam <> 'ERROR') then
      begin
        if ChampParam = 'REPFIN' then TOBFinitionToCash2000 (CodeParam, LibParam, FichierBE) else
        if ChampParam = 'REPMOD' then TOBModeleToCash2000   (CodeParam, LibParam, FichierBE) else
        if ChampParam = 'REPCOL' then TOBColorisToCash2000  (CodeParam, LibParam, FichierBE) else
        if ChampParam = 'REPMAT' then TOBMatiereToCash2000  (CodeParam, LibParam, FichierBE) ;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB ARRONDI
Mots clefs ... : TOX
*****************************************************************}
procedure TOBArrondiToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
var Enreg1    : string  ;
    PoidsCash : string  ;
    Poids     : int64   ;
    Cpt       : integer ;
    TobDonnees: TOB     ;
begin

  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    //
    // Conversion du poids !! problème d'arrondi e-18 !!
    //
    poids := Trunc (TobDonnees.GetValue ('GAR_POIDSARRONDI') * 100) ;
    if      poids = 20      then PoidsCash := '0'
    else if poids = 10      then PoidsCash := '1'
    else if poids = 50      then PoidsCash := '2'
    else if poids = 100     then PoidsCash := '3'
    else if poids = 500     then PoidsCash := '4'
    else if poids = 900     then PoidsCash := '5'
    else if poids = 1000    then PoidsCash := '6'
    else if poids = 5000    then PoidsCash := '7'
    else if poids = 9900    then PoidsCash := '8'
    else if poids = 10000   then PoidsCash := '9'
    else if poids = 50000   then PoidsCash := 'A'
    else if poids = 100000  then PoidsCash := 'B'
    else if poids = 500000  then PoidsCash := 'C'
    else if poids = 1000000 then PoidsCash := 'D' ;

    Enreg1 := 'ParC1 ar' + Format ('%-1.1s'   , [TobDonnees.GetValue('GAR_CODEARRONDI')]) + StringOfChar (' ', 5) +
                           Format ('%-15.15s' , [TobDonnees.GetValue('GAR_LIBELLE')]) +
                           Format ('%-1.1s'   , [TobDonnees.GetValue('GAR_METHODE')]) + PoidsCash ;

    Enreg1 := Trim (Enreg1) ;
    writeln (FichierBE, Enreg1);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 25/04/2002
Description .. : Conversion au format CASH 2000 de la TOB ARTICLE de
Suite ........ : type "marchandise"
Mots clefs ... : TOX
*****************************************************************}
procedure TOBMarchandiseToCash2000 (TobDonnees : TOB; var FichierBE : textfile; CodeSite : string) ;
var Enreg1        : string  ;
    Tob_ArtCompl  : TOB     ;
    TobFille      : TOB     ;
    CodeFamilleGB : string  ;
    RefArtGB      : string  ;
    RangTailleGB  : integer ;
    GrilleTailleGB: string  ;
    CodeTVA       : string  ;
    StatutArt     : string  ;
    DebutSolde    : string  ;
    FinSolde      : string  ;
    DateSolde     : TDateTime ;
    PxVente       : double  ;
    PxSolde       : double  ;
    PxAchat       : double  ;
    //////////////////////////////////////////////////////////////////////////////
    function AjoutChamp ( ChampParam : String ) : String ;
    var NomChamp, Prefixe : String ;
    begin
      Result := '' ;
      NomChamp := Tob_Param.GetValue (ChampParam) ;
      if NomChamp <> '' then
      begin
        // recherche de la valeur
        Prefixe := ExtractPrefixe (NomChamp) ;
        if Prefixe = 'GA' then
        begin
          if TobDonnees.FieldExists (NomChamp) then Result := TobDonnees.GetValue (NomChamp) ;
        end else
        if Prefixe = 'GA2' then
        begin
          if Tob_ArtCompl = Nil then
          begin
            Tob_ArtCompl := TOB.Create('ARTICLECOMPL', Nil, -1) ;
            Tob_ArtCompl.InitValeurs ;
            Tob_ArtCompl.SelectDB('"'+TobDonnees.GetValue('GA_ARTICLE')+'"', Nil) ;
          end;
          if Tob_ArtCompl.FieldExists (NomChamp) then Result := Tob_ArtCompl.GetValue (NomChamp) ;
        end;
        // génération de la tablette associée
        if Result <> '' then AjoutChoixCod (ChampParam, Result, FichierBE) ;
      end;
    end;
    //////////////////////////////////////////////////////////////////////////////
begin
  //
  // on ne traite pas les articles génériques
  //
  StatutArt := TobDonnees.GetValue ('GA_STATUTART') ;
  if StatutArt = 'GEN' then Exit ;
  Tob_ArtCompl := Nil ;
  //
  // Rang et grille de taille
  //
  CodeDimensionPGIToTailleGB (TobDonnees, GrilleTailleGB, RangTailleGB, FichierBE) ;
  //
  // Reference article GB (= code article générique + code interne dimension 2 + code interne dimension 3)
  //
  RefArtGB := TobDonnees.GetValue('GA_CODEARTICLE');
  if StatutArt = 'DIM' then RefArtGB:= RefArtGB + TobDonnees.GetValue('GA_CODEDIM2') +
                                                  TobDonnees.GetValue('GA_CODEDIM3');
  //
  // l'article est déjà traité ?
  //
  Enreg1 := Format ('%-3.3s'  , [CodeSite]) +
            Format ('%-12.12s', [RefArtGB]) +
            Format ('%2.2d'   , [RangTailleGB]) ;
  if TobArticleGB.FindFirst(['CODE'], [Enreg1], False) <> Nil then Exit ;
  TobFille := TOB.Create ('', TobArticleGB, -1) ;
  TobFille.AddChampSupValeur ('CODE', Enreg1) ;
  //
  // Gestion des familles
  //
  CodeFamilleGB := CodeFamillArtPGIToGB (TobDonnees, FichierBE) ;
  //
  // Recherche du Code TVA GB
  //
  CodeTVA := CodeTaxePGIToTVAGB (TobDonnees, FichierBE) ;
  //
  // Prix de vente théorique, de solde et d'achat
  //
  PxAchat := TobDonnees.GetValue('GA_PAHT') * 100 ;
  //FindTarifArt (TobDonnees, CodeSite, True) ;
  if TobDonnees.FieldExists('PXBASE') then PxVente := TobDonnees.GetValue('PXBASE') * 100
                                      else PxVente := TobDonnees.GetValue('GA_PVTTC') * 100 ;
  if TobDonnees.FieldExists('PXSOLDE') then PxSolde := TobDonnees.GetValue('PXSOLDE') * 100
                                       else PxSolde := TobDonnees.GetValue('GA_PVTTC') * 100 ;
  DebutSolde := '';
  if TobDonnees.FieldExists('DATEDEB') then
  begin
    DateSolde := TobDonnees.GetValue('DATEDEB') ;
    if (DateSolde > iDate1900) and (DateSolde < iDate2099) then
      DebutSolde := FormatDateTime('yyyymmdd', DateSolde) ;
  end;
  FinSolde := '';
  if TobDonnees.FieldExists('DATEFIN') then
  begin
    DateSolde := TobDonnees.GetValue('DATEFIN') ;
    if (DateSolde > iDate1900) and (DateSolde < iDate2099) then
      FinSolde := FormatDateTime('yyyymmdd', DateSolde) ;
  end;

  Enreg1 := 'CATC1 ' + Format ('%-3.3s'   , [CodeSite]) +
                       Format ('%-12.12s' , [RefArtGB]) +
                       Format ('%2.2d'    , [RangTailleGB]) +
                       Format ('%-6.6s'   , [CodeFamilleGB]) +
                       Format ('%-1.1s'   , [CodeTVA]) +
                       Format ('%-6.6s'   , [AjoutChamp('REPMOD')]) +
                       Format ('%-6.6s'   , [AjoutChamp('REPMAT')]) +
                       Format ('%-5.5s'   , [AjoutChamp('REPCOL')]) +
                       Format ('%-3.3s'   , [AjoutChamp('REPFIN')]) +
                       Format ('%-3.3s'   , [GrilleTailleGB]) +
                       Format ('%-15.15s' , [TobDonnees.GetValue('GA_LIBELLE')]) +
                       BoolPGIToGB (TobDonnees.GetValue('GA_FERME')) +
                       Format ('%-5.5s'   , [TobDonnees.GetValue('GA_FOURNPRINC')]) +
                       Format ('%12.0f'   , [PxVente]) +
                       Format ('%12.0f'   , [PxSolde]) +
                       Format ('%12.0f'   , [PxAchat]) +
                       Format ('%-8.8s'   , [DebutSolde]) +
                       Format ('%-8.8s'   , [FinSolde]) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

  Enreg1 := 'CATC2 ' + Format ('%-3.3s'   , [CodeSite]) +
                       Format ('%-12.12s' , [RefArtGB]) +
                       Format ('%2.2d'    , [RangTailleGB]) +
                       Format ('%-10.10s' , [Copy (TobDonnees.GetValue('GA_LIBELLE'), 16, 10)]) +
                       Format ('%-3.3s'   , [TobDonnees.GetValue('GA_COLLECTION')]) +
                       Format ('%-13.13s' , [TobDonnees.GetValue('GA_CODEBARRE')]) ;
  if Length (Trim (Enreg1)) > 23 then
  begin
    Enreg1 := Enreg1 + StringOfChar ('N', 3) ;
    Enreg1 := Trim (Enreg1) ;
    writeln (FichierBE, Enreg1);
  end;

  if Tob_ArtCompl <> Nil then Tob_ArtCompl.Free ;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 25/04/2002
Description .. : Conversion au format CASH 2000 de la TOB ARTICLE de
Suite ........ : type "prestation"
Mots clefs ... : TOX
*****************************************************************}
procedure TOBPrestationToCash2000 (TobDonnees : TOB; var FichierBE : textfile) ;
var Enreg1        : string  ;
    CodeTVA       : string  ;
    PxAchat       : double ;
begin
  //
  // Recherche du Code TVA GB
  //
  CodeTVA := CodeTaxePGIToTVAGB (TobDonnees, FichierBE) ;
  //
  // Prix d'achat
  //
  PxAchat := TobDonnees.GetValue('GA_PAHT') * 100 ;

  Enreg1 := 'PvdC1 vd' + Format ('%-2.2s'   , [TobDonnees.GetValue('GA_CODEARTICLE')]) + StringOfChar (' ', 4) +
                         Format ('%-15.15s' , [TobDonnees.GetValue('GA_LIBELLE')]) +
                         Format ('%-1.1s'   , [CodeTVA]) +
                         Format ('%12.0f'   , [PxAchat]) + StringOfChar (' ', 11) + '0' + StringOfChar (' ', 22) + 'NN ' ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 25/04/2002
Description .. : Conversion au format CASH 2000 de la TOB ARTICLE de
Suite ........ : type "article financier"
Mots clefs ... : TOX
*****************************************************************}
procedure TOBArtFinanToCash2000 (TobDonnees: TOB; var FichierBE : textfile) ;
var Enreg1        : string  ;
    SensOpc       : string  ;
    TypeOpc       : string  ;
begin
  //
  // Sens de l'opération
  //
  SensOpc := '+' ;
  TypeOpc := TobDonnees.GetValue('GA_TYPEARTFINAN') ;
  if Pos(TypeOpc, ';PRE;RAR;RAV;RBA;SOD;') > 0 then SensOpc := '-' ;

  Enreg1 := 'PocC1 oc' + Format ('%-2.2s'   , [TobDonnees.GetValue('GA_CODEARTICLE')]) + StringOfChar (' ', 4) +
                         Format ('%-15.15s' , [TobDonnees.GetValue('GA_LIBELLE')]) + StringOfChar (' ', 2) +
                         SensOpc + 'N' + StringOfChar (' ', 20) + 'T' + StringOfChar (' ', 8) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB ARTICLE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBArticleToCash2000 (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;
var Cpt           : integer ;
    TobDonnees    : TOB     ;
    TypeArt       : string  ;
begin
  //
  // Création des TOB listant les paramètres à pré-envoyer et les articles
  //
  CreeTobEnvoiGB ;

  FindTarifArt (ToxInfo, CodeSite) ;

  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];

    //
    // Traitement en fonction du type d'article
    //
    TypeArt := TobDonnees.GetValue('GA_TYPEARTICLE') ;
    if TypeArt = 'MAR' then TOBMarchandiseToCash2000 (TobDonnees, FichierBE, CodeSite) else
    if TypeArt = 'PRE' then TOBPrestationToCash2000 (TobDonnees, FichierBE) else
    if TypeArt = 'FI'  then TOBArtFinanToCash2000 (TobDonnees, FichierBE) ;

  end;
  LibereTobEnvoiGB ;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB des taux des devises
Mots clefs ... : TOX
*****************************************************************}
procedure TOBChancellToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
begin

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB FINITION ARTICLE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBFinitionToCash2000 (CodeParam, LibParam : String; var FichierBE : textfile) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PfiC1 fi' + Format ('%-3.3s'   , [CodeParam]) + StringOfChar (' ', 3) +
                         Format ('%-25.25s' , [LibParam]) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB MODELE ARTICLE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBModeleToCash2000 (CodeParam, LibParam : String; var FichierBE : textfile) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PmoC1 mo' + Format ('%-6.6s'   , [CodeParam]) +
                         Format ('%-25.25s' , [LibParam]) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB COLORIS ARTICLE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBColorisToCash2000 (CodeParam, LibParam : String; var FichierBE : textfile) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PcoC1 co' + Format ('%-5.5s'   , [CodeParam]) + ' ' +
                         Format ('%-25.25s' , [LibParam]) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB MATIERE ARTICLE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBMatiereToCash2000 (CodeParam, LibParam : String; var FichierBE : textfile) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PmaC1 ma' + Format ('%-6.6s'   , [CodeParam]) +
                         Format ('%-25.25s' , [LibParam]) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB FAMILLE ARTICLE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBFamilleArtToCash2000 (CodeParam, LibParam : String; var FichierBE : textfile) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PfsC1 fs' + Format ('%-6.6s'   , [CodeParam]) + StringOfChar (' ', 12) +
                         Format ('%-25.25s' , [LibParam]) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB GRILLE DE TAILLES
Mots clefs ... : TOX
*****************************************************************}
procedure TOBGrilleTailleToCash2000 (CodeGrille : String; CodeTaille : Array of String; var FichierBE : textfile) ;
var Enreg1    : string  ;
    Ind       : integer ;
begin

  // Nombre de tailles dans la grille
  Ind := High (CodeTaille) ;
  Inc(Ind) ;

  Enreg1 := 'PgrC1 gr' + Format ('%-3.3s'   , [CodeGrille]) + StringOfChar (' ', 3) +
                         Format ('%2.2d'    , [Ind]) ;
  Ind := Low (CodeTaille) ;
  while (Ind <= High (CodeTaille)) and (Ind < 27) do
  begin
    Enreg1 := Enreg1 + Format ('%4.4s', [CodeTaille[Ind]]) ;
    Inc(Ind) ;
  end;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

  if High (CodeTaille) >= 27 then
  begin
    Enreg1 := 'PgrC2 gr' + Format ('%-3.3s'   , [CodeGrille]) + StringOfChar (' ', 3) ;
    while (Ind <= High (CodeTaille)) and (Ind <= 30) do
    begin
      Enreg1 := Enreg1 + Format ('%4.4s', [CodeTaille[Ind]]) ;
      Inc(Ind) ;
    end;
    Enreg1 := Trim (Enreg1) ;
    writeln (FichierBE, Enreg1);
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB TVA
Mots clefs ... : TOX
*****************************************************************}
procedure TOBTVAToCash2000 (CodeTVA : String; TauxTVA : Double; var FichierBE : textfile) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PtvC1 tv' + Format ('%-1.1s'   , [CodeTVA]) + StringOfChar (' ', 5) +
                         Format ('%12.0f'   , [(TauxTVA * 10000)]) + StringOfChar (' ', 31) + '0' ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB DEMARQUES
Mots clefs ... : TOX
*****************************************************************}
procedure TOBDemarqueToCash2000 (TobDonnees : TOB; var FichierBE : textfile) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PdmC1 dm' + Format ('%-1.1s'   , [TobDonnees.GetValue('CC_CODE')]) + StringOfChar (' ', 5) +
                         Format ('%-25.25s' , [TobDonnees.GetValue('CC_LIBELLE')]) + 'NNOO' ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB FORMEJURIDIQUE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBFormeJuridiqueToCash2000 (TobDonnees : TOB; var FichierBE : textfile; Individu : String) ;
var Enreg1    : string  ;
begin

  Enreg1 := 'PfjC1 fj' + Format ('%-3.3s'   , [TobDonnees.GetValue('CC_CODE')]) + StringOfChar (' ', 3) +
                         Format ('%-25.25s' , [TobDonnees.GetValue('CC_LIBELLE')]) +
                         Format ('%-1.1s'   , [Individu]) ;
  Enreg1 := Trim (Enreg1) ;
  writeln (FichierBE, Enreg1);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB CHOIXCOD
Mots clefs ... : TOX
*****************************************************************}
procedure TOBChoixCodToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
var Cpt       : integer ;
    TobDonnees: TOB     ;
begin

  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    //
    // Gestion des démarques
    //
    if TobDonnees.GetValue ('CC_TYPE') = 'GTM' then
    begin
      TOBDemarqueToCash2000 (TobDonnees, FichierBE) ;
    end else
    //
    // Gestion des formes juridiques
    if TobDonnees.GetValue ('CC_TYPE') = 'CIV' then
    begin
      TOBFormeJuridiqueToCash2000 (TobDonnees, FichierBE, 'I') ;
    end else
    if TobDonnees.GetValue ('CC_TYPE') = 'JUR' then
    begin
      TOBFormeJuridiqueToCash2000 (TobDonnees, FichierBE, 'S') ;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB des statistiques clients
Mots clefs ... : TOX
*****************************************************************}
procedure TOBLibreTiersToCash2000 (TobDonnees : TOB; var FichierBE : textfile) ;
var Enreg1      : string  ;
    IdStat      : string  ;
    TypeStat    : string  ;
    NomChamp    : string  ;
    NomTablette : string  ;
    Ind         : integer ;
begin

  IdStat := ' ' ;
  TypeStat := TobDonnees.GetValue('YX_TYPE') ;
  for Ind := 1 to 20 do
  begin
    NomChamp := Tob_Param.GetValue('REPSTATC'+ Format('%2.2d', [Ind])) ;
    NomTablette := Get_Join(NomChamp) ;
    if (NomTablette <> '') and (TypeStat = TTToTipe(NomTablette)) then
    begin
      IdStat[1] := Chr(Ord('A') + Ind - 1) ;
    end;
  end;

  if IdStat <> ' ' then
  begin
    Enreg1 := 'P1' + Format ('%-1.1s'   , [IdStat]) + 'C1 1' + 
                     Format ('%-1.1s'   , [IdStat]) +
                     Format ('%-6.6s'   , [TobDonnees.GetValue('YX_CODE')]) +
                     Format ('%-25.25s' , [TobDonnees.GetValue('YX_LIBELLE')]) ;
    Enreg1 := Trim (Enreg1) ;
    writeln (FichierBE, Enreg1);
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CHOIXEXT
Mots clefs ... : TOX
*****************************************************************}
procedure TOBChoixExtToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
var Cpt       : integer ;
    TobDonnees: TOB     ;
begin

  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    //
    // Gestion des statistiques clients
    //
    if (TobDonnees.GetValue ('YX_TYPE') = 'LT1') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT2') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT3') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT4') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT5') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT6') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT7') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT8') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LT9') or
       (TobDonnees.GetValue ('YX_TYPE') = 'LTA') then
    begin
      TOBLibreTiersToCash2000 (TobDonnees, FichierBE) ;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB COMMERCIAL
Mots clefs ... : TOX
*****************************************************************}
procedure TOBCommercialToCash2000 (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;
var Enreg1        : string  ;
    Cpt           : integer ;
    TobDonnees    : TOB     ;
    CodeVendeur   : string  ;
    CodeEtab      : string  ;
    DateFinValid  : string  ;
    MontantCom    : double  ;
begin
  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    CodeEtab := TobDonnees.GetValue('GCL_ETABLISSEMENT');

    if (TobDonnees.GetValue('GCL_TYPECOMMERCIAL') = 'VEN') and (CodeEtab = CodeSite) then
    begin
      CodeVendeur  := copy ( TobDonnees.GetValue('GCL_COMMERCIAL'), 4, 2);
      DateFinValid := DateToStr ( TobDonnees.GetValue('GCL_DATESUPP')) ;
      DateFinValid := copy (DateFinValid, 7, 4) + copy (DateFinValid, 4, 2) + copy (DateFinValid, 1, 2);
      MontantCom   := TobDonnees.GetValue('GCL_COMMISSION')*1000 ;

      DecimalSeparator := '.' ;

      Enreg1 := 'PveC1 ve' + Format ('%-3.3s'   , [CodeEtab]) +
                             Format ('%-2.2s'   , [CodeVendeur]) + ' ' +
                             Format ('%12.3f'   , [MontantCom]) +
                             Format ('%-20.20s' , [TobDonnees.GetValue('GCL_LIBELLE')]) +
                             Format ('%-15.15s' , [TobDonnees.GetValue('GCL_PRENOM')]) + StringOfChar (' ', 8) +
                             DateFinValid ;

      Enreg1 := Trim (Enreg1) ;
      writeln (FichierBE, Enreg1);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB DEVISE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBDeviseToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
var Enreg1        : string  ;
    Cpt           : integer ;
    TobDonnees    : TOB     ;
    CodeDevGB     : string  ;
    CodeDevPGI    : string  ;
    IdentDev      : string  ;
    Nbdec         : integer ;
begin
  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    //
    // Recherche du Devise GB dans la table de correcpondance GB/PGI
    //
    CodeDevPGI := TobDonnees.GetValue ('D_DEVISE');
    CodeDevGB := TraduitPGIenGPAO(TypeRepriseGB2000, 'D_DEVISE', [CodeDevPGI], False, Tob_Devise);
    if CodeDevGB <> '' then
    begin
      IdentDev := '' ;
      if ((V_PGI.TenueEuro) and (CodeDevPGI = V_PGI.DevisePivot)) or (CodeDevPGI = 'EUR') then IdentDev := 'EUR' else
      if TobDonnees.GetValue('D_MONNAIEIN') = 'X' then IdentDev := 'IN' ;
      NbDec := TobDonnees.GetValue('D_DECIMALE') ;

      Enreg1 := 'PdeC1 de' + Format ('%-3.3s'   , [CodeDevGB]) + StringOfChar (' ', 3) +
                             Format ('%-25.25s' , [TobDonnees.GetValue('D_LIBELLE')]) +
                             Format ('%1.1d'    , [NbDec]) +
                             Format ('%1.1d'    , [NbDec]) +
                             Format ('%-3.3s'   , [IdentDev]) +
                             Format ('%1.1d'    , [NbDec]) ;
      Enreg1 := Trim (Enreg1) ;
      writeln (FichierBE, Enreg1);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB DIMENSION
Mots clefs ... : TOX
*****************************************************************}
procedure TOBDimensionToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
begin

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB DISPO
Mots clefs ... : TOX
*****************************************************************}
procedure TOBDispoToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
begin

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ETABLISSEMENT
Mots clefs ... : TOX
*****************************************************************}
procedure TOBEtablissToCash2000 (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;
var Enreg1        : string  ;
    Cpt           : integer ;
    TobDonnees    : TOB     ;
    CodeDevGB     : string  ;
    CodeDevPGI    : string  ;
    LgNumArt      : integer ;
begin
  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    if TobDonnees.GetValue ('ET_ETABLISSEMENT') <> CodeSite then
    begin
      //
      // Recherche du Devise GB dans la table de correcpondance GB/PGI
      //
      CodeDevPGI := TobDonnees.GetValue ('ET_DEVISE');
      CodeDevGB := TraduitPGIenGPAO(TypeRepriseGB2000, 'ET_DEVISE', [CodeDevPGI], False, Tob_Devise);
      //
      // Taille de la référence article
      //
      LgNumArt := GetParamSoc('SO_GCLGNUMART') ;
      if (LgNumArt < 1) or (LgNumArt > 12) then LgNumArt := 12 ;

      Enreg1 := 'PboC1 bo' + Format ('%-3.3s'   , [TobDonnees.GetValue('ET_ETABLISSEMENT')]) + StringOfChar (' ', 3) +
                             Format ('%-15.15s' , [TobDonnees.GetValue('ET_LIBELLE')]) + StringOfChar (' ', 3) +
                             StringOfChar ('N', 14) + StringOfChar (' ', 3) + 'N' +
                             Format ('%2.2d'    , [LgNumArt]) + '2' +
                             Format ('%-3.3s'   , [CodeDevGB]) +
                             Format ('%-3.3s'   , [TobDonnees.GetValue('ET_LANGUE')]) + StringOfChar (' ', 6) +
                             StringOfChar ('N', 3) +
                             Format ('%-3.3s'   , [TobDonnees.GetValue('ET_PAYS')]) + ' NN' ;
      Enreg1 := Trim (Enreg1) ;
      writeln (FichierBE, Enreg1);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB MODEPAIE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBModePaieToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
var Enreg1        : string  ;
    Cpt           : integer ;
    TobDonnees    : TOB     ;
    Stg           : string  ;
    CodePaieGB    : string  ;
    LibellePaieGB : string  ;
    CodePaiePGI   : string  ;
    CodeDevGB     : string  ;
    CodeDevPGI    : string  ;
    TypePaie      : string  ;
    SensPaie      : string  ;
    CliOblig      : string  ;
begin
  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];

    if TobDonnees.GetValue ('MP_UTILFO') = 'X' then
    begin
      //
      // Recherche du Mode de paiement GB dans la table de correcpondance GB/PGI
      //
      CodePaiePGI := TobDonnees.GetValue ('MP_MODEPAIE');
      Stg := TraduitPGIenGPAO(TypeRepriseGB2000, 'MP_MODEPAIE', [CodePaiePGI], True, Tob_Paiement);
      CodePaieGB := ReadTokenST(Stg);
      if CodePaieGB <> '' then
      begin
        LibellePaieGB := ReadTokenST(Stg);
        if LibellePaieGB = '' then LibellePaieGB := TobDonnees.GetValue ('MP_LIBELLE');
        CodeDevPGI := TobDonnees.GetValue ('MP_DEVISEFO');
        CodeDevGB := TraduitPGIenGPAO(TypeRepriseGB2000, 'D_DEVISE', [CodeDevPGI], False, Tob_Devise);
        TypePaie := TobDonnees.GetValue ('MP_TYPEMODEPAIE');
        if TypePaie = '001' then
        begin
          TypePaie := 'E' ;     // Espèce
          Senspaie := 'O' ;
        end else
        if TypePaie = '002' then
        begin
          TypePaie := 'A' ;     // Avoir
          Senspaie := 'N' ;
        end else
        if TypePaie = '003' then
        begin
          TypePaie := 'C' ;     // Chèque
          Senspaie := 'O' ;
        end else
        if TypePaie = '004' then
        begin
          TypePaie := 'R' ;     // Chèque différé
          Senspaie := 'N' ;
        end else
        if TypePaie = '005' then
        begin
          TypePaie := 'T' ;     // Carte bancaire
          Senspaie := 'O' ;
        end else
        if TypePaie = '006' then
        begin
          TypePaie := 'H' ;     // Arrhes déjà versés
          Senspaie := 'N' ;
        end else
        if TypePaie = '007' then
        begin
          TypePaie := 'R' ;     // Reste dû
          Senspaie := 'N' ;
        end else
        if TypePaie = '008' then
        begin
          TypePaie := 'B' ;     // Bon d'achat
          Senspaie := 'N' ;
        end else
        begin
          TypePaie := 'Z' ;     // Autre mode de règlement
          Senspaie := 'O' ;
        end ;
        if TobDonnees.GetValue('MP_CLIOBLIGFO') = 'X' then CliOblig := 'O' else CliOblig := 'N' ;

        Enreg1 := 'PmrC1 mr' + Format ('%-3.3s'   , [CodePaieGB]) + StringOfChar (' ', 3) +
                               Format ('%-25.25s' , [LibellePaieGB]) + StringOfChar (' ', 2) +
                               Format ('%-3.3s'   , [CodeDevGB]) + TypePaie + SensPaie + StringOfChar (' ', 3) + CliOblig +
                               Format ('%-1.1s'   , [TobDonnees.GetValue('MP_ARRONDIFO')]) ;
        Enreg1 := Trim (Enreg1) ;
        writeln (FichierBE, Enreg1);
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB PAYS
Mots clefs ... : TOX
*****************************************************************}
procedure TOBPaysToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
var Enreg1        : string  ;
    Cpt           : integer ;
    TobDonnees    : TOB     ;
    TrouvePays    : boolean ;
    Stg           : string  ;
    CodePaysGB    : string  ;
    LibellePaysGB : string  ;
    CodePaysPGI   : string  ;
    PaysCEE       : string  ;
begin
  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    //
    // Recherche du Pays GB dans la table de correcpondance GB/PGI
    //
    TrouvePays  := False ;
    CodePaysPGI := TobDonnees.GetValue ('PY_PAYS');
    if CodePaysPGI <> '' then
    begin
      Stg := TraduitPGIenGPAO(TypeRepriseGB2000, 'PY_PAYS', [CodePaysPGI], True, Tob_Pays);
      CodePaysGB    := ReadTokenST(Stg);
      LibellePaysGB := ReadTokenST(Stg);
      if LibellePaysGB = '' then LibellePaysGB := TobDonnees.GetValue ('PY_LIBELLE');
      if CodePaysGB <> '' then TrouvePays := True ;
    end;
    if TrouvePays then
    begin
      //
      // Pays CEE ??
      //
      if TobDonnees.GetValue ('PY_MEMBRECEE') = 'X' then PaysCEE := 'O'
      else PaysCEE := 'N';

      Enreg1 := 'PpaC1 pa' + Format ('%-3.3s'   , [CodePaysGB]) + StringOfChar (' ', 3) +
                             Format ('%-25.25s' , [LibellePaysGB]) + StringOfChar (' ', 3) + PaysCEE ;
      Enreg1 := Trim (Enreg1) ;
      writeln (FichierBE, Enreg1);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Conversion au format CASH 2000 de la TOB CODEPOSTAL
Mots clefs ... : TOX
*****************************************************************}
procedure TOBCodePostalToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
var Enreg1        : string  ;
    Cpt           : integer ;
    TobDonnees    : TOB     ;
    CodePaysGB    : string  ;
    CodePaysPGI   : string  ;
begin
  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];
    //
    // Recherche du Pays GB dans la table de correcpondance GB/PGI
    //
    CodePaysPGI := TobDonnees.GetValue ('O_PAYS');
    CodePaysGB := TraduitPGIenGPAO(TypeRepriseGB2000, 'O_PAYS', [CodePaysPGI], False, Tob_Pays);
    if CodePaysGB <> '' then
    begin
      Enreg1 := 'PL1C1 1' + Format ('%-3.3s'   , [CodePaysGB]) +
                            Format ('%-7.7s'   , [TobDonnees.GetValue('O_CODEPOSTAL')]) + StringOfChar (' ', 5) +
                            Format ('%-32.32s' , [TobDonnees.GetValue('O_VILLE')]) ;
      Enreg1 := Trim (Enreg1) ;
      writeln (FichierBE, Enreg1);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PIECE
Mots clefs ... : TOX
*****************************************************************}
procedure TOBPieceToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
begin

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TARIF
Mots clefs ... : TOX
*****************************************************************}
procedure TOBTarifToCash2000 (ToxInfo : TOB; var FichierBE : textfile) ;
begin

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TIERS
Mots clefs ... : TOX
*****************************************************************}
procedure TOBTiersToCash2000 (ToxInfo : TOB; var FichierBE : textfile; CodeSite : string) ;
var Enreg1        : string  ;
    Cpt           : integer ;
    TobDonnees    : TOB     ;
    Tob_TiersCompl: TOB     ;
    CodeTiersGB   : string  ;
    CodePaysGB    : string  ;
    CodeNatGB     : string  ;
    CodeDevGB     : string  ;
    CodePGI       : string  ;
    RepParam      : string  ;
    ProMail       : string  ;
    EtatCpta      : string  ;
    DatNais       : string  ;
    CodeVendeur   : string  ;
    Stg           : string  ;
    Ind           : Integer ;
    Val           : Integer ;
    //////////////////////////////////////////////////////////////////////////////
    procedure LireTiersCompl ;
    begin
      if Tob_TiersCompl = Nil then
      begin
        Tob_TiersCompl := TOB.Create('TIERSCOMPL', Nil, -1) ;
        Tob_TiersCompl.InitValeurs ;
        Tob_TiersCompl.SelectDB('"'+TobDonnees.GetValue('T_AUXILIAIRE')+'"', Nil) ;
      end;
    end;
    //////////////////////////////////////////////////////////////////////////////
    procedure InsereChampStatistique ;
    var TypChp, IdStat : string  ;
        Nb             : integer ;
    begin
      TypChp := ChampToType (RepParam);
      if Tob_TiersCompl.FieldExists (RepParam) then
      begin
        if TypChp = 'BOOLEAN' then
        begin
          if Tob_TiersCompl.GetValue (RepParam) = 'X' then Stg := 'O' else Stg := 'N' ;
          Stg := Stg + StringOfChar (' ', 5) ;
        end
        else if TypChp = 'DATE' then Stg := FormatDateTime ('ddmmyy', Tob_TiersCompl.GetValue (RepParam))
        else if TypChp = 'DOUBLE' then Stg := StrfMontant (Tob_TiersCompl.GetValue (RepParam), 6, 0, '', False)
        else Stg := Format ('%-6.6s', [Tob_TiersCompl.GetValue (RepParam)]) ;
        if Stg <> '' then
        begin
          if Enreg1 = '' then
          begin
            if Ind < 19 then IdStat := 'A' else IdStat := 'B' ;
            Enreg1 := 'CLIC'+ IdStat +' ' + Format ('%-3.3s', [CodeSite]) + Format ('%-6.6s', [CodeTiersGB]) + StringOfChar (' ', 108) ;
          end;
          if Ind < 19 then Nb := 1 else Nb := 19 ;
          Insert (Stg, Enreg1, (16 + ((Ind - Nb) * 6))) ;
        end;
      end;
    end;
    //////////////////////////////////////////////////////////////////////////////
begin
  for Cpt:=0 to ToxInfo.Detail.Count-1 do
  begin
    TobDonnees   := ToxInfo.Detail[Cpt];

    //
    // Constitution du code client (code client ou code boutque + code client)
    //
    CodeTiersGB := TobDonnees.GetValue ('T_TIERS') ;
    RepParam := Tob_Param.GetValue ('CODECLI') ;
    if RepParam = '0' then
    begin
      if Copy (CodeTiersGB, 1, 3) = CodeSite then Delete(CodeTiersGB, 1, 3) else CodeTiersGB := '' ;
    end;

    if CodeTiersGB <> '' then
    begin
      Tob_TiersCompl := Nil ;
      //
      // Recherche du Pays GB dans la table de correcpondance GB/PGI
      //
      CodePGI     := TobDonnees.GetValue ('T_PAYS');
      CodePaysGB  := TraduitPGIenGPAO(TypeRepriseGB2000, 'T_PAYS', [CodePGI], False, Tob_Pays);
      CodePGI     := TobDonnees.GetValue ('T_NATIONALITE');
      CodeNatGB   := TraduitPGIenGPAO(TypeRepriseGB2000, 'T_PAYS', [CodePGI], False, Tob_Pays);
      CodePGI     := TobDonnees.GetValue ('T_DEVISE');
      CodeDevGB   := TraduitPGIenGPAO(TypeRepriseGB2000, 'T_DEVISE', [CodePGI], False, Tob_Devise);
      //
      // Envoi prochain mailing
      //
      ProMail := 'O' ;
      RepParam := Tob_Param.GetValue ('REPCLIPROMAIL');
      if RepParam <> '' then
      begin
        LireTiersCompl ;
        if Tob_TiersCompl <> Nil then
        begin
          if RepParam = 'Booléen libre n° 1' then RepParam := 'YTC_BOOLLIBRE1' else
          if RepParam = 'Booléen libre n° 2' then RepParam := 'YTC_BOOLLIBRE2' else
          if RepParam = 'Booléen libre n° 3' then RepParam := 'YTC_BOOLLIBRE3' else RepParam := '' ;
          if (RepParam <> '') and (Tob_TiersCompl.FieldExists (RepParam)) then
          begin
            if Tob_TiersCompl.GetValue (RepParam) <> 'X' then ProMail := 'N' ;
          end;
        end;
      end;
      //
      // Etat comptable
      //
      EtatCpta := 'V' ;
      RepParam := Tob_Param.GetValue ('REPETACPTA');
      if RepParam <> '' then
      begin
        LireTiersCompl ;
        if (Tob_TiersCompl <> Nil) and (Tob_TiersCompl.FieldExists (RepParam)) then
        begin
          EtatCpta := Tob_TiersCompl.GetValue (RepParam) ;
        end;
      end;
      //
      // Date de naissance
      //
      Val := TobDonnees.GetValue('T_ANNEENAISSANCE') ;
      if Val > 0 then DatNais := Format ('%-4.4d', [Val]) else DatNais := StringOfChar (' ', 4) ;
      Val := TobDonnees.GetValue('T_MOISNAISSANCE') ;
      if Val > 0 then DatNais := DatNais + Format ('%-2.2d', [Val]) else DatNais := DatNais + StringOfChar (' ', 2) ;
      Val := TobDonnees.GetValue('T_JOURNAISSANCE') ;
      if Val > 0 then DatNais := DatNais + Format ('%-2.2d', [Val]) else DatNais := DatNais + StringOfChar (' ', 2) ;
      //
      // vendeur
      //
      CodeVendeur  := copy ( TobDonnees.GetValue('T_REPRESENTANT'), 4, 2);

      Enreg1 := 'CLIC1 ' + Format ('%-3.3s'   , [CodeSite]) +
                           Format ('%-6.6s'   , [CodeTiersGB]) + StringOfChar (' ', 14) +
                           Format ('%-4.4s'   , [TobDonnees.GetValue('T_JURIDIQUE')]) +
                           Format ('%-32.32s' , [TobDonnees.GetValue('T_LIBELLE')]) +
                           Format ('%-32.32s' , [TobDonnees.GetValue('T_PRENOM')]) +
                           Format ('%-7.7s'   , [TobDonnees.GetValue('T_CODEPOSTAL')]) +
                           Format ('%-20.20s' , [TobDonnees.GetValue('T_TELEPHONE')]) + 'O' +
                           Format ('%-1.1s'   , [EtatCpta]) ;
      Enreg1 := Trim (Enreg1) ;
      if Length (Enreg1) > 15 then writeln (FichierBE, Enreg1);

      Enreg1 := 'CLIC2 ' + Format ('%-3.3s'   , [CodeSite]) +
                           Format ('%-6.6s'   , [CodeTiersGB]) +
                           Format ('%-32.32s' , [TobDonnees.GetValue('T_ADRESSE1')]) +
                           Format ('%-32.32s' , [TobDonnees.GetValue('T_ADRESSE2')]) +
                           Format ('%-32.32s' , [TobDonnees.GetValue('T_ADRESSE3')]) +
                           Format ('%-3.3s'   , [CodePaysGB]) +
                           Format ('%-3.3s'   , [CodeNatGB]) + DatNais ;
      Enreg1 := Trim (Enreg1) ;
      if Length (Enreg1) > 15 then writeln (FichierBE, Enreg1);

      Enreg1 := 'CLIC3 ' + Format ('%-3.3s'   , [CodeSite]) +
                           Format ('%-6.6s'   , [CodeTiersGB]) +
                           Format ('%-32.32s' , [TobDonnees.GetValue('T_VILLE')]) +
                           Format ('%-2.2s'   , [CodeVendeur]) +  
                           FormatDateTime ('ddmmyyyy', TobDonnees.GetValue('T_DATECREATION')) +
                           FormatDateTime ('ddmmyyyy', TobDonnees.GetValue('T_DATEMODIF')) +
                           Format ('%-8.8s'   , [TobDonnees.GetValue('T_UTILISATEUR')]) +
                           StringOfChar (' ', 8) + ProMail + StringOfChar (' ', 3) + 'C' +
                           Format ('%-3.3s'   , [TobDonnees.GetValue('T_TARIFTIERS')]) +
                           Format ('%-3.3s'   , [CodeDevGB]) +
                           Format ('%-20.20s' , [TobDonnees.GetValue('T_TELEX')]) +
                           Format ('%-3.3s'   , [TobDonnees.GetValue('T_LANGUE')]) ;
      Enreg1 := Trim (Enreg1) ;
      if Length (Enreg1) > 15 then writeln (FichierBE, Enreg1);
      //
      // Statistiques client
      //
      Enreg1 := '' ;
      for Ind := 1 to 20 do
      begin
        RepParam := Tob_Param.GetValue('REPSTATC'+ Format('%2.2d', [Ind])) ;
        if RepParam <> '' then
        begin
          // recherche et mise en place de la valeur
          LireTiersCompl ;
          if Tob_TiersCompl <> Nil then InsereChampStatistique ;
        end;
        if (Ind = 18) or (Ind = 20) then
        begin
          // ecriture dans le fichier 
          Enreg1 := Trim (Enreg1) ;
          if Enreg1 <> '' then
          begin
            writeln (FichierBE, Enreg1);
            Enreg1 := '' ;
          end;
        end;
      end;
      if Tob_TiersCompl <> Nil then Tob_TiersCompl.Free ;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Fonction générique qui transforme les info TOX en chaîne
Suite ........ : de caractère pour CASH 2000
Mots clefs ... : TOX
*****************************************************************}
procedure ToxToCash2000 (NomTable, CodeSite : string; ToxInfo : TOB; var FichierBE : textfile) ;
begin
  if      (NomTable = 'ARRONDI')      then TOBArrondiToCash2000   (ToxInfo, FichierBE)
  else if (NomTable = 'ARTICLE')      then TOBArticleToCash2000   (ToxInfo, FichierBE, CodeSite)
  else if (NomTable = 'CHANCELL')     then TOBChancellToCash2000  (ToxInfo, FichierBE)
  else if (NomTable = 'CHOIXCOD')     then TOBChoixCodToCash2000  (ToxInfo, FichierBE)
  else if (NomTable = 'CHOIXEXT')     then TOBChoixExtToCash2000  (ToxInfo, FichierBE)
  else if (NomTable = 'COMMERCIAL')   then TOBCommercialToCash2000(ToxInfo, FichierBE, CodeSite)
  else if (NomTable = 'DEVISE')	      then TOBDeviseToCash2000    (ToxInfo, FichierBE)
  else if (NomTable = 'DIMENSION')    then TOBDimensionToCash2000 (ToxInfo, FichierBE)
  else if (NomTable = 'DISPO')	      then TOBDispoToCash2000     (ToxInfo, FichierBE)
  else if (NomTable = 'ETABLISS')     then TOBEtablissToCash2000  (ToxInfo, FichierBE, CodeSite)
  else if (NomTable = 'MODEPAIE')     then TOBModePaieToCash2000  (ToxInfo, FichierBE)
  else if (NomTable = 'PAYS')	      then TOBPaysToCash2000      (ToxInfo, FichierBE)
  else if (NomTable = 'CODEPOST')     then TOBCodePostalToCash2000(ToxInfo, FichierBE)
  else if (NomTable = 'PIECE')	      then TOBPieceToCash2000     (ToxInfo, FichierBE)
  else if (NomTable = 'TARIF')	      then TOBTarifToCash2000     (ToxInfo, FichierBE)
  else if (NomTable = 'TIERS')	      then TOBTiersToCash2000     (ToxInfo, FichierBE, CodeSite);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 22/05/2002
Description .. : Envoi sur le concentrateur de caisses CASH
Mots clefs ... : TOX
*****************************************************************}
procedure FTPToCash2000 ;
{$IFDEF TOXCLIENT}
var xpFtp   : TOXFTP;
    PortFtp : Integer;
    Host, Username, Password, RepFtp, RepDepart, RepArchive, LocalFileName, HostFileName, ArchFileName : String;
    sr : TSearchRec;
{$ENDIF}
begin
{$IFDEF TOXCLIENT}
  // Recherche du paramétrage du transfert FTP
  PortFtp    := 21;                                                      // port IP affecté au service FTP
  Host       := GetParamSoc ('SO_GCORLIHOST');
  if Host = '' then Exit;
  Username   := GetParamSoc ('SO_GCORLIUSERNAME');
  if Username = '' then Exit;
  Password   := GetParamSoc ('SO_GCORLIPASSWORD');
  if Password = '' then Exit;
  RepFtp     := GetParamSoc ('SO_GCORLIREPFTP');
  if RepFtp = '' then Exit;
  if RepFtp[Length(RepFtp)] <> '/' then RepFtp := RepFtp + '/';          // serveur GO ou ORLI = SCO
  RepDepart  := IncludeTrailingBackslash(GetParamSoc ('SO_GCREPORLI'));
  RepArchive := IncludeTrailingBackslash(GetParamSoc ('SO_GCREPORLIARCH'));
  if (Host = '') or (Username = '') or (Password = '') or (RepFtp = '')  or (RepDepart = '') or (RepArchive = '')then Exit;
  // Recherche de fichiers BE
  if FindFirst(RepDepart+'BE*.*', faAnyFile, sr) = 0 then
  begin
    Try
      xpFtp := TOXFTP.Create (PortFtp);
      if xpFTP.Connect (Host, Username, Password) then
      begin
        repeat
          LocalFileName := RepDepart + sr.Name;
          HostFileName  := RepFtp + Uppercase(sr.Name);
          ArchFileName  := RepArchive + Uppercase(sr.Name);
          if xpFTP.PutFileAscii(LocalFileName, HostFileName) then
          begin
            if CopyFile (pchar(LocalFileName), pchar(ArchFileName), False) then
              DeleteFile(pchar(LocalFileName)) ;
          end;
        until (FindNext(sr) <> 0);
        xpFTP.DisConnect;
      end;
      xpFTP.Free;
    Finally
      FindClose(sr);
    End;
  end;
{$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Constitue le nom du fichier d'envoi au CASH
Suite ........ : BEXXXYY.ZZZ
Mots clefs ... : TOX
*****************************************************************}
function ToxNomFichierBE ( CodeSite, NumeroCaisse : String ; NumJour : Integer ) : String ;
begin
  if Numjour <= 0 then Numjour := NumSemaine (NowH) * 7 + DayOfWeek (NowH) ;
  Result := GetParamSoc ('SO_GCREPORLI') + '\BE' + CodeSite + NumeroCaisse + '.' + Format('%.3u', [Numjour]) ;
end;

//////////////////////////////////////////////////////////////////////////////////////
// Traitement de la TOX avant l'archivage
// Utilité :
//
//         1 - Si site en CASH 2000, les informations seront transformées et un fichier ASCII sera généré
//             Le code retour est False (on ne crée pas de fichier TOX
//         2 - On peut toper les pièces extraites pour interdire les modifications dès l'envoi TOX
//
///////////////////////////////////////////////////////////////////////////////////////
procedure AvantArchivageToxPourCASH (Tox : TOB ; CodeSite, LesCaisses : String) ;
var ToxCondition           : TOB 	      ;
    ToxTable0		   : TOB	      ;
    NomTable0		   : string	      ;
    ii, jj 		   : integer          ;
    NumJour		   : integer          ;
    NumeroCaisse           : string           ;
    FichierBE              : textfile         ;
    NomFichier             : string           ;
    NomFichierCopy         : string           ;
begin
  //
  // Chargement des TOB paramètres
  //
  CreeTobParamCorresp ;
  //
  // Récupération de la premiere caisse
  //
  NumeroCaisse  := ReadTokenSt (LesCaisses);
  //
  // Ouverture du fichier BE....
  //
  Numjour := NumSemaine (NowH) * 7 + DayOfWeek (NowH) ;
  NomFichier := ToxNomFichierBE (CodeSite, NumeroCaisse, NumJour) ;

  AssignFile(FichierBE, NomFichier);
  if FileExists(NomFichier) then Append(FichierBE) else Rewrite(FichierBE) ;

  Try
    if (Tox.FieldExists('EVENEMENT')) then
    begin
      for ii:=0 to Tox.Detail.Count-1 do
      begin
        //
        // TOX CONDITION CONTENANT LA CONDITION
        //
        ToxCondition:=Tox.Detail[ii];
        if (ToxCondition.FieldExists('CONDITION')) then
        begin
          for jj:=0 to ToxCondition.Detail.Count-1 do
          begin
            //
            // TOX TABLE Niveau 0
            //
            ToxTable0:=ToxCondition.Detail[jj];
            if (ToxTable0.FieldExists('TOX_TABLE')) then
            begin
              NomTable0 := ToxTable0.GetValue ('TOX_TABLE');			// Nom de la table principale
              ToxToCash2000 (NomTable0, CodeSite, ToxTable0, FichierBE) ;
            end;
          end;
        end;
      end;
    end;
  Finally
    //
    // Fermeture du fichier BE
    //
    CloseFile (FichierBE) ;
  End;
  //
  // Liberation des TOB
  //
  LibereTobParamCorresp ;
  //
  // Y a t il d'autres caisses ?
  //
  NumeroCaisse := ReadTokenSt(LesCaisses);
  while NumeroCaisse <> '' do
  begin
    NomFichierCopy := ToxNomFichierBE (CodeSite, NumeroCaisse, NumJour) ;
    CopyFile (pchar(NomFichier), pchar(NomFichierCopy), False) ;
    NumeroCaisse := ReadTokenSt(LesCaisses);
  end;
  //
  // Envoi sur le concentrateur de caisses CASH
  //
  FTPToCash2000 ;
end;

//
// Initialisation des champs de la TOB de paramétrage
//
//procedure TFRecupGB2000.DefineChampTobParam ;
procedure DefineChampTobParam (TobParametrage : TOB) ;
var cpt : integer;
    Chtmp : string;
begin
  // Création des champs dans la TOB
  for cpt:=Low(ChampTobParam) to High(ChampTobParam) do
  begin
    if (ChampTobParam[cpt, 1]<>'') and not (TobParametrage.FieldExists(ChampTobParam[cpt, 1])) then
    begin
      TobParametrage.AddChampSupValeur(ChampTobParam[cpt, 1], ChampTobParam[cpt, 2]);
    end ;
  end ;

  // Pour les articles
  for cpt:=1 to 20 do
  begin
    Chtmp := 'REPSTATA'+ Format('%2.2d', [cpt]) ;
    if not TobParametrage.FieldExists(Chtmp) then TobParametrage.AddChampSupValeur(Chtmp, '') ;
  end;

  // Pour les clients
  for cpt:=1 to 20 do
  begin
    Chtmp := 'REPSTATC'+ Format('%2.2d', [cpt]) ;
    if not TobParametrage.FieldExists(Chtmp) then TobParametrage.AddChampSupValeur(Chtmp, '') ;
  end;

  // Pour les Fournisseurs
  for cpt:=1 to 20 do
  begin
    Chtmp := 'REPSTATF'+ Format('%2.2d', [cpt]) ;
    if not TobParametrage.FieldExists(Chtmp) then TobParametrage.AddChampSupValeur(Chtmp, '') ;
  end;

end;

//
// Chargement de la TOB de paramétrage
//
procedure ChargeTobParam (TobParametrage : TOB) ;
var Stg     : string;
    TOBCorr : TOB ;
    TOBL    : TOB ;
    Ind     : Integer ;
begin
  // l'ancienne structure de la TOB de paramétrage est conservée
  TOBCorr := TOB.Create('', Nil, -1) ;
  Stg := 'SELECT GRO_NOMCHAMP,GRO_OPTIONS FROM REPRISECOGPAO'
       + ' WHERE GRO_TYPEREPRISE="'+ TypeRepriseParam +'" AND GRO_VALEURGPAO="..."' ;
  TOBCorr.LoadDetailFromSQL(Stg) ;
  for Ind := 0 to TOBCorr.Detail.Count -1 do
  begin
    TOBL := TOBCorr.Detail[Ind] ;
    Stg := TOBL.GetValue('GRO_NOMCHAMP') ;
    if TobParametrage.FieldExists(Stg) then
    begin
      TobParametrage.PutValue(Stg, TOBL.GetValue('GRO_OPTIONS')) ;
    end;
  end;
  TOBCorr.Free ;
end;

end.

