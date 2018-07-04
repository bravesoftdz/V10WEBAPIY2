{***********UNITE*************************************************
Auteur  ...... : P.Dumet
Créé le ...... : 02/2006
Modifié le ... : 02/2006
Description .. : Lancement des vignettes Paie
Mots clefs ... :
*****************************************************************}
unit PGVignettePaie;

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HEnt1,
  uAncetreVignettePlugIn;

type
  TGAVignettePaie = class (TAncetreVignettePlugIn)

  private
    {Fait office de RegisterClasses}
    class function GetObjet (TobIn, TobOut: TOB; ClassName: string): TGAVignettePaie;
  protected
    {Récupération des controls/valeurs depuis la TobRequest}
    function GetInterface (NomGrille: string = ''): Boolean; override;
    {Affectation des controls/valeurs dans la TobRespons}
    function SetInterface: Boolean; override;
  public
    class function NouvelleInstance (Param: string; RequestTob: TOB; ClassName: string): TOB; override;

    constructor Create (TobIn, TobOut: Tob); override;
    destructor Destroy; override;
  end;
var
  ParamFich  : string;
  DateJour   : TDateTime;

implementation

uses
  SysUtils,
  AFVignetteChartBrutPaie,
  AFVignetteChartAbsPaie,
  PGVignLstSotie,
//  PGVIGVALABS,
//  PGVIGSAISINSC,
//  PGVIGSAISABS,
//  PGVIGSAISABSSAL,
  PGVIGABSPERIODE,
  PGVIGEFFECTIFS,
  PGVIGHEURES,
  PGVIGLSTENTREE,
  PGVIGFINCDD,
  PGVIGFINESSAI,
  PGVIGCARTSEJOUR,
  PGVIGCONTEFF,
  PGVIGTBEFF,
  PGVIGINTERIM,
  PGVIGSTAGESAL,
  PGVIGHANDICAP,
  PGVIGCOTOREP,
  PGVIGVISITMED,
  PGVIGTOOLSIDE,
  PGVIGSTAGIAIRES,
  PGVIGCOMPTDIF,
  PGVIGLISTEDDIF,
  PGVIGDDIFSAL,
  PGVIGSAISPREV,
  PGVIGCATSTAGE,
  PGVIGSESSIONS,
  PGVIGSAISDEM,
  PGVIGLISTEPREV,
  PGVIGCOMPAPR,
  PGVIGGESABSSAL,
  PGVIGGESABS,
//  PGVIGCALABSSAL,
  PGVIGPREVSAL,
  PGVIGSCORING;

{Fonction de création de l'objet et lancement du traitement métier
{---------------------------------------------------------------------------------------}

class function TGAVignettePaie.NouvelleInstance (Param: string; RequestTob: TOB; ClassName: string): TOB;
var
  ResponseTOB: TOB;
  MonObjet: TGAVignettePaie;
begin
  ddWriteLN ('AFVignettePaie : NouvelleInstance');

  ResponseTOB := TOB.Create ('Out', nil, -1);
  ResponseTOB.AddChampSupValeur ('ERROR', '');

  MonObjet := TGAVignettePaie.GetObjet (RequestTob, ResponseTOB, ClassName);
  try
    ddWriteLN ('AFVignettePaie : Début de traitement');
    ParamFich := Param;
    MonObjet.LanceTraitement (Param);
  finally
    ddWriteLN ('AFVignettePaie : Fin de traitement');
    if Assigned (MonObjet) then
      FreeAndNil (monObjet);
  end;

  Result := ResponseTOB;
  if Result = nil then
    ddWriteLN ('AFVignettePaie : RESPONSETOB VIDE');
  //Result.SaveToXMLFile ('C:\temp\ResponseTOB.txt', True, True);
end;

{---------------------------------------------------------------------------------------}

function TGAVignettePaie.GetInterface (NomGrille: string = ''): Boolean;
begin
  Result := inherited GetInterface (NomGrille);
end;

{---------------------------------------------------------------------------------------}

function TGAVignettePaie.SetInterface: Boolean;
var
  T: TOB;
  G: string;
begin
  Result := True;
  if Assigned (TobDonnees) then
  begin
    {Mets les données dans la grille}
    case Ancetre of
//      vaf_Vierge: G := 'GRID';
      vaf_Vierge: G := 'GRSALARIE';
      vaf_Mul: G := 'FListe';
    else
      G := '';
    end;
    ddWriteLN ('AFVignettePaie : Resultat (' + IntToStr (TobDonnees.Detail.Count) + ')');
    T := TOB.Create ('~TEMP', nil, -1);
    try
      {On est obligé de sauvegarder la Tob car le PutGridDetail change le Parent des filles}
      T.Dupliquer (TobDonnees, True, True);
//      TObDonnees.SaveToFile ('C:\temp\TOBDONNEE0.TXT',FALSE,TRUE, TRUE);
      {Mets les données dans la grille}
      PutGridDetail (G, TobDonnees);
      SetControlValue (G, 1);
      TobDonnees.Dupliquer (T, True, True);
      ddWriteLN ('AFVignettePaie : putGrid (' + IntToStr (TobDonnees.Detail.Count) + ')');
    finally
      FreeAndNil (T);
    end;
    {"Dessine la grille"}
    DrawGrid (G);
  end
  else
    Result := False;
end;

{---------------------------------------------------------------------------------------}

constructor TGAVignettePaie.Create (TobIn, TobOut: Tob);
begin
  inherited Create (TobIn, TobOut);
end;

{---------------------------------------------------------------------------------------}

destructor TGAVignettePaie.Destroy;
begin
  inherited Destroy;
end;

{Cette fonction fait office de RegisterClass
{---------------------------------------------------------------------------------------}

class function TGAVignettePaie.GetObjet (TobIn, TobOut: TOB; ClassName: string): TGAVignettePaie;
begin

  if ClassName = 'PGVIGNETTECHARTABS' then Result := TGChartAbsPAIE.Create (TobIn, TobOut)
  else if ClassName = 'PGVIGNETTECHARTBR' then Result := TGChartBrutPAIE.Create (TobIn, TobOut)
  else if ClassName = 'PGVIGNLSTSORTIE' then Result := PGVignLstSortie.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_ABSPERIODE' then Result := PG_VIG_ABSPERIODE.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_EFFECTIFS' then Result := PG_VIG_EFFECTIFS.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_TBEFF' then Result := PG_VIG_TBEFF.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_HEURES' then Result := PG_VIG_HEURES.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_LSTENTREE' then Result := PG_VIG_LSTENTREE.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_FINCDD' then Result := PG_VIG_FINCDD.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_FINESSAI' then Result := PG_VIG_FINESSAI.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_CARTSEJOUR' then Result := PG_VIG_CARTSEJOUR.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_CONTEFF' then Result := PG_VIG_CONTEFF.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_INTERIM' then Result := PG_VIG_INTERIM.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_HANDICAP' then Result := PG_VIG_HANDICAP.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_COTOREP' then Result := PG_VIG_COTOREP.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_VISITMED' then Result := PG_VIG_VISITMED.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_TOOLSIDE' then Result := PG_VIG_TOOLSIDE.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_STAGIAIRES' then Result := PG_VIG_STAGIAIRES.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_COMPTDIF' then Result := PG_VIG_COMPTDIF.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_LISTEDDIF' then Result := PG_VIG_LISTEDDIF.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_SAISPREV' then Result := PG_VIG_SAISPREV.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_STAGESAL' then Result := PG_VIG_STAGESAL.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_DDIFSAL' then Result := PG_VIG_DDIFSAL.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_CATSTAGE' then Result := PG_VIG_CATSTAGE.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_SESSIONS' then Result := PG_VIG_SESSIONS.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_SAISDEM'  then Result := PG_VIG_SAISDEM.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_LISTEPREV'  then Result := PG_VIG_LISTEPREV.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_COMPAPR'  then Result := PG_VIG_COMPAPR.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_GESABSSAL'  then Result := PG_VIG_GESABSSAL.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_GESABS'  then Result := PG_VIG_GESABS.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_PREVSAL'  then Result := PG_VIG_PREVSAL.Create (TobIn, TobOut)
  else if ClassName = 'PG_VIG_SCORING'  then Result := PG_VIG_SCORING.Create (TobIn, TobOut)
//  else if ClassName = 'PG_VIG_SAISABS' then Result := PG_VIG_SAISABS.Create (TobIn, TobOut)
//  else if ClassName = 'PG_VIG_SAISINSC' then Result := PG_VIG_SAISINSC.Create (TobIn, TobOut)
//  else if ClassName = 'PG_VIG_VALABS' then Result := PG_VIG_VALABS.Create (TobIn, TobOut)
//  else if ClassName = 'PG_VIG_SAISABSSAL' then Result := PG_VIG_SAISABSSAL.Create (TobIn, TobOut)
//  else if ClassName = 'PG_VIG_CALABSSAL' then Result := PG_VIG_CALABSSAL.Create (TobIn, TobOut)
  else Result := TGAVignettePaie.Create (TobIn, TobOut);
end;

end.
