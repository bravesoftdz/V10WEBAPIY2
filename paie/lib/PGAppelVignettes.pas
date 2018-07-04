{***********UNITE*************************************************
Auteur  ...... : P.dumet
Créé le ...... : 02/2006
Modifié le ... : 02/2006
Description .. : Appel des vignettes Paie
Mots clefs ... :
*****************************************************************
PT1  | 19/01/2008 | FLO | Simplification du chargement des vignettes sans devoir spécifier chacune d'entre elles
PT2  | 18/03/2008 | FLO | Modification du numéro de dossier si mode PCL et connexion sur DB00000 avec partage formation
PT3  | 18/03/2008 | FLO | Récupération du type d'utilisateur connecté : Responsable, Assistant ou Secrétaire
PT4  | 03/06/2008 | FLO | Modification des codes séria des différents bouquets
}
unit PGAppelVignettes;

interface

uses classes,
  UTOB,
  windows,
  sysutils,
  eSession,
  uToolsPlugin,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  db,
  HEnt1,
  HCtrls,
  eDispatch,
  forms,
  uHTTP,
  PgVIGUTIL,
  hDTLinks;

procedure InitApplication;
function VignettesPGDispatch (Action, Param: string; RequestTOB: TOB): TOB;
function VignettesPGProcCalcEdt (sf, sp: string): variant;
function Dispatch (Action, Param: string; RequestTOB: TOB): TOB; stdcall;

implementation
uses uHTTPCS,
  PGVignettePaie, Paramsoc, PGOutilsFormation;

Const BOUQUET_GRATUIT   = '00485080';
	  BOUQUET_PAIE      = '00650080';
      BOUQUET_FORMATION = '00652080';
      BOUQUET_CONGES    = '00654080';

//TEMPORAIRE - DEBUT
procedure PgAffecteNoDossier ();
Var LeDos : String;
begin
	// Récupération du numéro de dossier actuel
	leDos := GetParamSocSecur ('SO_NODOSSIER', '');
	
	// Si le dossier est indiqué, on le garde, sinon on met 000000
	if (LeDos <> '000000') And (LeDos <> '') Then
   		V_PGI.NoDossier:= LeDos
	else if ((LeDos = '') And (Length (V_PGI.NoDossier)=8)) Then
   		V_PGI.NoDossier:= '000000';

    // Modification possible du numéro de dossier si partage formation et connexion sur la base partagée
    MultiDosFormation; //PT2
end;
//TEMPORAIRE - FIN

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/12/2007
Modifié le ... :   /  /
Description .. : Création du dictionnaire de vignettes disponibles en
Suite ........ : rapport avec les bouquets (gratuit, paie, formation, etc.)
Mots clefs ... :
***************************************²**************************}
Function GetVignetteDictionary : TOB;
var
	ResponseTOB : TOB;
	
    procedure add(AName: string; AModule: string = BOUQUET_GRATUIT);
    var t: tob;
    begin
      T := TOB.Create('dictionary_item', ResponseTOB, -1);
      T.AddChampSupValeur('CODE',AName);
      T.AddChampSupValeur('SERIA_MODULE',AModule);
    end;
begin
    ResponseTOB := TOB.Create ('DICTIONARY', nil, -1);
      
	// Bouquet gratuit
	Add('PGVIGNETCHARTABS');
	Add('PG_VIG_CONTEFF');
	Add('PG_VIG_FINCDD');
	Add('PGVIGNLSTSORTIE');
	Add('PG_VIG_COTOREP');
	Add('PG_VIG_FINESSAI');
	Add('PG_VIG_ABSPERIODE');
  
  	// Bouquet Paie
	Add('PG_VIG_TBEFF', 	BOUQUET_PAIE);
	Add('PG_VIG_HEURES', 	BOUQUET_PAIE);
	Add('PGVIGNETTECHARTBR',BOUQUET_PAIE);
	Add('PG_VIG_CARTSEJOUR',BOUQUET_PAIE);
	Add('PG_VIG_LSTENTREE', BOUQUET_PAIE);
	Add('PG_VIG_HANDICAP',  BOUQUET_PAIE);
	Add('PG_VIG_INTERIM',	BOUQUET_PAIE);
	Add('PG_VIG_VISITMED',	BOUQUET_PAIE);
	
	// Bouquet Congés
	Add('PG_VIG_GESABSSAL',	BOUQUET_CONGES);
	Add('PG_VIG_GESABS',	BOUQUET_CONGES);

	// Bouquet Formation
	Add('PG_VIG_COMPTDIF',	BOUQUET_FORMATION);
	Add('PG_VIG_SAISPREV',	BOUQUET_FORMATION);
	Add('PG_VIG_DDIFSAL',	BOUQUET_FORMATION);
	Add('PG_VIG_STAGESAL',	BOUQUET_FORMATION);
	Add('PG_VIG_LISTEDDIF',	BOUQUET_FORMATION);
	Add('PG_VIG_CATSTAGE',	BOUQUET_FORMATION);
	Add('PG_VIG_STAGIAIRES',BOUQUET_FORMATION);
	Add('PG_VIG_SESSIONS',	BOUQUET_FORMATION);
	Add('PG_VIG_SAISDEM',	BOUQUET_FORMATION);
	Add('PG_VIG_LISTEPREV',	BOUQUET_FORMATION);
	Add('PG_VIG_COMPAPR',	BOUQUET_FORMATION);
	Add('PG_VIG_PREVSAL',	BOUQUET_FORMATION);
	Add('PG_VIG_SCORING',	BOUQUET_FORMATION);

	Result := ResponseTOB;
end;


function VignettesPGDispatch (Action, Param: string; RequestTOB: TOB): TOB;
begin
//  RequestTOB.savetofile('c:\RequestTOB.txt',false,true,true);

  Result := nil;
  Action := UpperCase (Action);
  ddWriteLN ('PGIPAIEVIGNETTES : VignettesDispatch (' + Action + ')');

  if not Assigned (RequestTob) then
  begin
    ddWriteLN ('PGIPAIEVIGNETTES : RequestTOB non assignée');
    Exit;
  end;

  //PT1 - Début
  if (Pos('PGVIG', Action) = 1) Or (Pos('PG_VIG_', Action) = 1) then
    Result := TGAVignettePaie.NouvelleInstance (Param, RequestTOB, Action)
  else
  //PT1 - Fin
  begin
      //EPZ 27.08.2007 nom de vignette incorrect      Result := TOB.Create ('$PAIEVIGNETTES', nil, -1);
      Result := TOB.Create ('_PAIEVIGNETTES', nil, -1);
      Result.AddChampSupValeur ('ERROR', '');
      ddWriteLN ('PGIPAIEVIGNETTES : Action non trouvée');
  end;
end;

function Dispatch (Action, Param: string; RequestTOB: TOB): TOB; stdcall;
Var LaSession : TISession;
begin
    // Récupération de la session afin de savoir si on a déjà chargé les paramètres liés
    LaSession := LookupCurrentSession;
    // On teste l'un des types, en l'occurrence ici FORmation
    // S'il n'existe pas, c'est qu'il faut faire le traitement
    If LaSession.UserObjects.IndexOf(TYPE_UTILISATEUR+'FOR') = -1 Then
    Begin
      PgAffecteNoDossier;
      SauveTypeUtilisateur('FOR');
      SauveTypeUtilisateur('ABS');
    End;

  ddWriteLN ('PGIVIGNETTES : Dispatch');
  if Action = 'PgRecupAbsUser' then Result := PgRecupAbsUser(Param)
  else
    if Action = 'PgRecupRecapUser' then Result := PgRecupRecapUser()
  else
    if Action = 'PgRecupJourHeureAbs' then Result := PgRecupJourHeureAbs(RequestTOB)
  else
    if Action = 'PgValidAbsUser' then Result := PgValidAbsUser(RequestTOB)
  else
    if Action = 'PgRecupInfoUserAndResp' then Result := PgRecupInfoUserAndResp(Param)
  else
    if Action = 'PgValidMailIndicUser' then Result := PgValidMailIndicUser(Param)
  else
    if Action = 'PgRecupBodyMailUser' then Result := PgRecupBodyMailUser(Param,RequestTOB)
  else
    if Action = 'GetVignetteDictionary' then Result := GetVignetteDictionary // retourne la liste des vignettes dispos dans ce plugin

  else result := VignettesPGDispatch (Action, Param, RequestTOB);

//    result := VignettesPGDispatch (Action, Param, RequestTOB);
end;

procedure InitApplication;
begin
  ddWriteLN ('PGIPAIEVIGNETTES Initialisation...');
end;

function VignettesPGProcCalcEdt (sf, sp: string): variant;
begin
  result := '';
end;

end.
