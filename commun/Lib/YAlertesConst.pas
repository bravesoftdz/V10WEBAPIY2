unit YAlertesConst;

interface

const
  { Modes blocage }
  SansBlocage = '001';
  Blocage = '002';
  Interrogative = '003';
  CodeOuverture = '1OP';
  CodeModification = '2MF';
  CodeCreation = '3CR';
  CodeModifChamps = '4MC';
  CodeDateAnniv = '5DA';
  CodePieceTiers = '6PI';
  CodeTiersArticle = '7PI';
  CodeSuppression = '8SU';
  { évènements }
  CodeEvent: array[1..8] of string = (
        {1}'1OP',
        {2}'2MF',
        {3}'3CR',
        {4}'4MC',
        {5}'5DA',
        {6}'6PI',
        {7}'7PI',
        {8}'8SU'
    );
  { Flag pour TobAlertes ne mode sans écran }
  ALERT_SansEcran = 'ALERT_SANSECRAN';
  ALERT_TobMsg = 'Alerte';

  nomIndic: string = 'RESINDIC';
  nbEvents: integer = 50;

  { Pour filtrer les tables alerte non GP }
  WhereAlertesNotGP : string = ' YTA_PREFIXE<>"GPO" AND YTA_PREFIXE<>"GEM" AND YTA_PREFIXE<>"GDE" AND YTA_PREFIXE<>"YTS" '
                        + 'AND YTA_PREFIXE<>"WOB" AND YTA_PREFIXE<>"WOL" AND YTA_PREFIXE<>"WAN" AND YTA_PREFIXE<>"WGL" AND YTA_PREFIXE<>"WGO" '
                        + 'AND YTA_PREFIXE<>"WGR" AND YTA_PREFIXE<>"WGS" AND YTA_PREFIXE<>"WGT" AND YTA_PREFIXE<>"WLS" AND YTA_PREFIXE<>"WNL" '
                        + 'AND YTA_PREFIXE<>"WNT" AND YTA_PREFIXE<>"WOG" AND YTA_PREFIXE<>"WOP" AND YTA_PREFIXE<>"WOR" AND YTA_PREFIXE<>"WOT" '
                        + 'AND YTA_PREFIXE<>"WPF" ';
implementation

end.
