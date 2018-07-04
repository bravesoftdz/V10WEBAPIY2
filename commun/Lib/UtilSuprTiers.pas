unit UtilSuprTiers;

interface

uses
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}
{$ENDIF}
  Hctrls,
  ParamSoc,
  Hent1,
  Ent1;

type
  T_SupTiersControle = class
  private
    sSocietes : string;
    bEncours  : boolean;

  public

    property Societes: string  read sSocietes write sSocietes;
    property Encours : boolean read bEncours  write bEncours;

    function SupAEstDansPiece(St: string): Boolean;
    function SupAEstDansActivite(St: string): Boolean;
    function SupAEstDansRessource(St: string): Boolean;
    function SupAEstDansActions(St: string): Boolean;
    function SupAEstDansPersp(St: string): Boolean;
    function SupAEstDansAffaire(St: string): Boolean;
    function SupAEstDansIntervenant(St: String): Boolean;
    function SupAEstDansCata(St: string): Boolean;
    function SupAEstDansPaie(St: string): Boolean;
    function SupAEstDansMvtPaie(St: string): Boolean;
    function SupAEstEcrGuide(St: string): Boolean;
    function SupAEstDansSociete(St: string): Boolean;
    function SupAEstDansSection(St: string): Boolean;
    function SupAEstDansUtilisat(St: string): Boolean;
    function SupAEstCpteCorresp(St: string): Boolean;
    function SupAEstUnPayeur(St: string): Boolean;
    function SupAEstMouvemente(St: string): Boolean;
    function SupAEstDansProjets(St: string): Boolean;
    function SupAEstDansGED(St: string): Boolean;
    function SupAEstDansParc(St: string): Boolean;

    function ExisteSqlMS(Sql: string): boolean;
  end;

implementation

{ SQL multi-sociétés }
 function T_SupTiersControle.ExisteSqlMS(Sql: string): boolean;
var
  Q: TQuery;
begin
  Q := OpenSql(Sql, true, -1, '', false, Societes);
  try
    result := not Q.Eof;
  finally                                            
    Ferme(Q);
  end;
end;

function T_SupTiersControle.SupAEstDansPiece(St: string): Boolean;
var sWhere : string;
begin
  if bEncours then sWhere := ' AND GP_VIVANTE="X" '
  else sWhere := '';
  result := ExisteSqlMS('SELECT GP_TIERS FROM PIECE WHERE (GP_TIERS="' + St + '" OR GP_TIERSLIVRE="' + St + '" OR GP_TIERSFACTURE="' + St + '") '+ sWhere);
end;

function T_SupTiersControle.SupAEstDansActivite(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT ACT_TIERS FROM ACTIVITE WHERE ACT_TIERS="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstDansRessource(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT ARS_AUXILIAIRE FROM RESSOURCE WHERE ARS_AUXILIAIRE="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstDansActions(St: string): Boolean;
var sWhere : string;
begin
  if bEncours then sWhere := ' AND RAC_ETATACTION="PRE" '
  else sWhere := '';
  result := ExisteSqlMS('SELECT RAC_AUXILIAIRE FROM ACTIONS WHERE RAC_AUXILIAIRE="' + St + '"' + sWhere);
end;

function T_SupTiersControle.SupAEstDansPersp(St: string): Boolean;
var sWhere : string;
begin
  if bEncours then sWhere := ' AND RPE_ETATPER="ENC" '
  else sWhere := '';
  result := ExisteSqlMS('SELECT RPE_AUXILIAIRE FROM PERSPECTIVES WHERE RPE_AUXILIAIRE="' + St + '"'  + sWhere);
end;

function T_SupTiersControle.SupAEstDansAffaire(St: string): Boolean;
var sWhere : string;
begin
  if bEncours then sWhere := ' AND AFF_ETATAFFAIRE <> "CLO" '
  else sWhere := '';
  result := ExisteSqlMS('SELECT AFF_TIERS FROM AFFAIRE WHERE AFF_TIERS="' + St + '"' + sWhere);
end;

// BDU - 11/04/07 - FQ : 13920
function T_SupTiersControle.SupAEstDansIntervenant(St: String): Boolean;
var
  sJoin,
  sWhere: String;
begin
  if bEnCours then
  begin
    sJoin := ' INNER JOIN AFFAIRE ON (AFT_AFFAIRE = AFF_AFFAIRE)';
    sWhere := ' AND AFF_ETATAFFAIRE <> "CLO"';
  end
  else
  begin
    sJoin := '';
    sWhere := '';
  end;
  Result := ExisteSqlMS('SELECT AFT_TIERS FROM AFFTIERS' + sJoin +
    ' WHERE AFT_TIERS = "' + St + '"' + sWhere);
end;

function T_SupTiersControle.SupAEstDansCata(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT GCA_TIERS FROM CATALOGU WHERE GCA_TIERS="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstDansPaie(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT PSA_AUXILIAIRE FROM SALARIES WHERE PSA_AUXILIAIRE="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstDansMvtPaie(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT PPU_SALARIE FROM PAIEENCOURS LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE WHERE PSA_AUXILIAIRE="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstEcrGuide(St: string): Boolean;
begin
  if not bEncours then Result := Presence('ECRGUI', 'EG_AUXILIAIRE', St)
  else result := false;
end;

function T_SupTiersControle.SupAEstDansSociete(St: string): Boolean;
var
  StLoc: string;
begin
  if not bEncours then
  begin
  {$IFDEF SPEC302}
    result := ExisteSqlMS('Select SO_CLIATTEND,SO_FOUATTEND From SOCIETE Where SO_CLIATTEND="' + St + '" OR SO_FOUATTEND="' + St + '"');
  {$ELSE}
    Result := True;
    StLoc := GetParamSoc('SO_CLIATTEND');
    if StLoc = St then
      Exit;
    StLoc := GetParamSoc('SO_FOUATTEND');
    if StLoc = St then
      Exit;
    if St = VH^.TiersDefSal then
      Exit;
    if St = VH^.TiersDefDiv then
      Exit;
    if St = VH^.TiersDefCli then
      Exit;
    if St = VH^.TiersDefFou then
      Exit;
    Result := False;
  {$ENDIF}
  end else result := false;
end;

function T_SupTiersControle.SupAEstDansSection(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('Select S_MAITREOEUVRE,S_CHANTIER From SECTION Where S_MAITREOEUVRE="' + St + '" OR S_CHANTIER="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstDansUtilisat(St: string): Boolean;
begin
  if not bEncours then Result := Presence('UTILISAT', 'US_AUXILIAIRE', St)
  else result := false;
end;

function T_SupTiersControle.SupAEstCpteCorresp(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('Select CR_CORRESP From CORRESP Where (CR_TYPE="AU1" OR CR_TYPE="AU2" )AND CR_CORRESP="' + St + '" ')
  else result := false;
end;

function T_SupTiersControle.SupAEstUnPayeur(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT T_PAYEUR FROM TIERS WHERE T_PAYEUR="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstMouvemente(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="' + St + '" ' +
    'AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE="' + St + '" )')
  else result := false;
end;

function T_SupTiersControle.SupAEstDansProjets(St: string): Boolean;
var sWhere : string;
begin
  if bEncours then sWhere := ' AND RPJ_ETAT="X" '
  else sWhere := '';
  result := ExisteSqlMS('SELECT RPJ_TIERS FROM PROJETS WHERE RPJ_TIERS="' + St + '"'  + sWhere);
end;

function T_SupTiersControle.SupAEstDansGED(St: string): Boolean;
begin
  if not bEncours then result := ExisteSqlMS('SELECT RTD_TIERS FROM RTDOCUMENT WHERE RTD_TIERS="' + St + '"')
  else result := false;
end;

function T_SupTiersControle.SupAEstDansParc(St: string): Boolean;
begin
  result := ExisteSqlMS('SELECT * FROM WPARC WHERE WPC_ETATPARC="ES" AND WPC_TIERS="' + St + '"');
end;

end.
 