{***********A.G.L.***********************************************
Auteur  ...... : Paie PGI
Créé le ...... : 04/07/2001
Modifié le ... :   /  /
Description .. : Unit de définition des liens OLE de la paie
Suite ........ : Contient les fonctions de déclarations du dictionnaire
Suite ........ : pour les contrats de travail
Mots clefs ... : PAIE;LIENOLE
*****************************************************************}
{
PT-1-1 : 07/09/2001 SB V547 Modification des requêtes de selection des champs.
                            Selection sur le nom du champ et non sur le numéro
                            de champ..
PT-1-2 : 07/09/2001 SB V547 Nouvelle fonction de recherche d'un libélle associé
                            à un code
PT2    : 29/03/2002 JL V571 Ajout médecine du travail
PT3    : 19/08/2002 PH V585 Ajout Fonction liens OLE pour  EXCEL
PT4    : 19/08/2002 PH V585 Ajout Fonctions publiques pour le bilan social et
                            les liens OLE pour EXCEL
PT5    : 08/11/2002 JL V595 Ajout Formation
PT6    : 02/04/2003 SB V_42 L'écriture de la classe est inopérande depuis une
                            certaine version??
PT7    : 09/03/2004 PH V_50 Mise en place des concepts de la paie ie
                            confidentialité
PT8    : 23/04/2004 SB V_50 FQ 11084 Nouvelle fonction DICO Word :
                            DUREECONTRATJOUR
PT9    : 28/04/2004 SB V_50 FQ 10812 Nouvelle fonction DICO Word : Gestion des
                            déclarants
PT10   : 28/04/2004 SB V_50 FQ 11085 Nouvelle fonction DICO Word : Accord
                            grammaire
PT11   : 04/02/2005 SB V_60 FQ 11748 Ajout champ tob etablissement
PT12   : 07/02/2005 SB V_60 FQ 11785 Ajout fn renvoie durée préavis
                            @DUREEPREAVIS
PT13   : 21/03/2005 SB V_60 FQ 11423 Ajout fn renvoie donnée organisme
PT14-1 : 16/09/2005 SB V_65 FQ 12456 Ajout fn élement de la paie
PT14-2 : 16/09/2005 SB V_65 FQ 12456 Ajout Tri tob
PT15   : 20/09/2005 SB V_65 FQ 12421 Ajout info registre, table annuaire
PT16   : 26/09/2005 SB V_65 Refonte dico pgi
PT17   : 26/09/2005 SB V_65 Affichage des formules AGL
PT18   : 07/10/2005 SB V_65 Affichage de l'assistant en modal
PT19   : 16/11/2005 PH V_65 FQ 12685 rajout somme unite pris dans l'effectif
PT20   : 03/05/2006 PH V_65 prise en compte driver dbMSSQL2005
PT21   : 19/05/2006 SB V_65 Appel des fonctions communes pour identifier
                            V_PGI.driver
PT22   : 19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT23   : 20/06/2006 PH V_65 RAZ de la TOBAncienneté
PT24   : 12/02/2007 SB V_70 FQ 13782 Correction fn ole
PT25   : 07/03/2007 GGS V_80 lien OLE Multidossier
PT26   : 21/05/2007 FC V_72 Rajout dans le dico de la fontion @GETELTNAT dans
                            les utilitaires
PT27   : 12/06/2007 GGU V_72 FQ 13541 : Un effectif global équivalent temps plein
PT28   : 26/06/2007 VG V_72 Implémentation de la procédure OnTheLast
PT29   : 27/06/2007 FC V_72 FQ 13953 Rajout des fonctions DUREEESSAISEMAINE, DUREEPREAVISSEMAINE, DUREECONTRATSEMAINE
PT30   : 08/07/2008 NA V_850 FQ 11785 Ajout Fn renvoie durée préavis en jours @DUREEPREAVISJOUR
}
unit OlePaiePgi;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  SysUtils,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_MAIN,
{$ELSE}
  MainEagl,
{$ENDIF}
  OleDicoPGI, OLEAuto, UTob, HEnt1,Hmsgbox, HCtrls, Entpaie,utilpgi;


type
  // Automation Server
  TOlePaiePGI = class(TAutoObject)
  automated
    function New: Variant;
    function Init: variant;
    function Free: Variant;
    function ShowDico: Variant;
    function PaieCumulMois(Etab, Sal, LaTable, Quoi, Rub, TRV1, TRV2, TRV3, TRV4,
      LCMB1, LCMB2, LCMB3, LCMB4, CODESTAT, ZMoisD, ZAnneeD, ZMoisF, ZAnneeF,
      CatDucs, DadsProf: string; Present: string = ''): Variant;
    function PaieStatEffectif(Etab, TRV1, TRV2, TRV3, TRV4, LCMB1, LCMB2, LCMB3,
      LCMB4, CodeStat, ZMoisF, ZAnneeF, CatDucs, DadsProf, DadsCat, Sexe: string;
      LibEmpl: string = ''; LibCoeff: string = ''; LibInd: string = ''; LibNiv: string = ''; LibQual: string = ''; CodePCS: string = ''; Contrat: string = ''; EntSort: string =
      ''; MotifCont: string = ''; SommeEff : String =''): Variant;
    function PGOleGetInfoTable(Table, Sal, ChampSelect, Dollar: string): Variant;
    function PGBSSINFO(Quoi, val, val1, deb, fin, categorie: string): Variant;
    function PGAssistOle(st: string): string;
    function PAIEBSCalIndOLE(DateDeb, Datefin, Pres, Ind, Etab, Categ, Emploi, Coeff, Qualif,
      CodeStat, TravailN1, TravailN2, TravailN3, TravailN4,
      TabLibre1, TabLibre2, TabLibre3, TabLibre4: string): Variant;
    function RecupClauseSql(StVal, StChamp: string): string;
    function PGMultiSoc():Variant;   //PT25
  private
    Regroupsoc, NomBase : String;  //PT25
    function Getregroupsoc: String;
    procedure setRegroupsoc(const Value: String);
    function GetNomBase: String;
    procedure setNomBase(const Value: String);
  public
    IndexReg : Integer;
    constructor Create; override; //PT6
    destructor Destroy; override; //PT6
    property RegroupementSociete : String read Getregroupsoc write setRegroupsoc; //PT25
    property ListeBases : String read GetNomBase write setNomBase;                //PT25
  protected
    procedure OnTheLast (var shutdown: Boolean);
  end;

procedure RegisterPaiePGI;
procedure PaieVideAnc (); // PT23

procedure GetTobFunc(theTobFunc: TOB {a ne pas vider : a utiliser comme parent});
procedure GetTobDico(theTobDico: TOB);

// PT4 19/08/2002 PH V585 Ajout Fonctions publiques pour le bilan social et les liens OLE
procedure InitBSAnciennete(StWhere: string);
procedure FreeBSAnciennete;
function RendBSAnciennete(V1, V2: WORD; DD, DF: TDateTime; Categorie: string): integer;
function RendBSLeWhere(DD, DF: TDateTime; Champ, LaValeur, Categorie: string): string;
function RendBSAge(V1, V2: WORD; DF: TDateTime): string;
function ExecBSSQL(LeSQL: string): Integer;
function ExecBSSumTP(LeWhere: string): Extended; //PT27
function RendBSCumul(DD, DF: TDateTime; Cumul, LeWhere: string): Double;
function RendBSAbsences(DD, DF: TDateTime; Motif, LeWhere: string): Double;
// FIN PT4
function RendSQLMois(LaDate: string): string;
function PGControlConcept(): Boolean;

implementation

uses AssistPgExcelOle, pgcommun, P5def, DB;

var
  OleT_Sal, OleT_Contrat, TOB_AncienneteBS: TOB;
  PgOleSal: string;
  {================================ OLE =====================================}
const //DEB PT6 Déplacement code
  AutoClassInfo: TAutoClassInfo = (
    AutoClass: TOlePaiePGI;
    ProgID: 'CegidPGI.PaiePGI';
    ClassID: '{72D0CE25-4AAB-4d96-AC4C-FEF1C99E704D}';
    Description: 'Cegid Paie PGI';
    Instancing: acMultiInstance);

procedure RegisterPaiePGI;
begin
  ;
end;
//FIN PT6
// PT3 19/08/2002 PH V585 Ajout Fonction liens OLE pour  EXCEL
{ Fonction de gestion des messages d'erreur pour les liens OLE
}

function PGOLEErreur(i: Integer): string;
begin
  Result := '#ERREUR:';
  case i of
    1: Result := Result + 'Cumul non renseigné';
    2: Result := Result + 'Rubrique non renseignée';
    3: Result := Result + 'Date non renseignée';
    4: Result := Result + 'Année début non renseignée';
    5: Result := Result + 'Champ à cumuler non renseigné';
    6: Result := Result + 'Nature rubrique non renseignée';
    7: Result := Result + 'tablette non renseignée';
    10: Result := Result + 'générale SQL !';
  else Result := Result + 'Non définie ' + IntToStr(i);
  end;
end;

{Fonction qui remplie une tob pour eviter de faire de  multiples requêtes SQL pour
obtenir plusieurs champs sur le même enregistrement de la table
}

procedure PGOleAllocTob(LaTable, Sal: string);
var
  st: string;
  Q: TQuery;
  TheTob: TOB;
begin
  if PgOleSal = Sal then exit; // pas de rechargement car on travaille tjs sur le même salarié
  if LaTable = 'SALARIES' then
  begin
    if OleT_Sal <> nil then
    begin
      OleT_Sal.free;
      OleT_Sal := nil;
    end;
    st := 'SELECT * FROM SALARIES WHERE PSA_SALARIE="' + Sal + '"';
    TheTob := TOB.Create('Mon Salarie', nil, -1);
    Q := OpenSql(st, TRUE);
    TheTob.LoadDetailDB('SALARIES', '', '', Q, FALSE);
    Ferme(Q);
    OleT_Sal := TheTob.findFirst([''], [''], FALSE);
  end
  else
  begin
    if LaTable = 'CONTRATS' then
    begin
      if OleT_contrat <> nil then
      begin
        OleT_contrat.free;
        OleT_contrat := nil;
      end;
      st := 'SELECT * FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + Sal + '" ORDER BY PCI_SALARIE,PCI_ORDRE DESC';
      TheTob := TOB.Create('Mon contrat', nil, -1);
      Q := OpenSql(st, TRUE);
      TheTob.LoadDetailDB('CONTRATTRAVAIL', '', '', Q, FALSE);
      Ferme(Q);
      OleT_contrat := TheTob.findFirst([''], [''], FALSE);
    end;
  end;
  PgOleSal := Sal;
end;
{ Procedure de desallocation des objets TOB pour les liens OLE
}

procedure PGOleDesAllocTob(LaTable, Sal: string);
begin
  if OleT_Sal <> nil then
  begin
    OleT_Sal.free;
    OleT_Sal := nil;
  end;
  if OleT_contrat <> nil then
  begin
    OleT_contrat.free;
    OleT_contrat := nil;
  end;
end;
{
Fonction qui remplie la chaine de la requete SQL pour le traitement d'un cumul,d'une rubrique
}

procedure ConfectSqlOle(LePref, LeChamp, LaValeur: string; var St: string; var LeAnd: Boolean);
var
  Chaine: string;
  i, j: Integer;
begin
  if (LeChamp <> '') and (LaValeur <> '') then
  begin
    j := Pos(';', LaValeur);
    if LeAnd then
    begin
      if j = 0 then St := St + ' AND '
      else St := St + ' AND (';
    end
    else if j > 0 then st := st + ' (';
    i := 0;
    Chaine := readtokenst(LaValeur);
    while Chaine <> '' do
    begin
      if i > 0 then St := St + ' OR ';
      st := st + ' ' + LePref + '_' + LeChamp + '="' + Chaine + '"';
      LeAnd := TRUE;
      Chaine := readtokenst(LaValeur);
      i := 1;
    end;
    if j <> 0 then st := st + ')';
  end;
end;
{ TPaiePGI }
// FIN  PT3

function TOlePaiePGI.ShowDico: variant;
begin
  AfficherDictionnairePGI;
  result := unassigned;
end;

function TOlePaiePGI.Free: variant;
begin
  FreeDictionnairePGI;
  result := unassigned;
end;

function TOlePaiePGI.New: variant;
begin
  NewDictionnairePGI;
  result := unassigned;
end;

function TOlePaiePGI.Init: variant;
begin
  InitDictionnairePGI(GetTobDico, GetTobFunc, nil, '');
  result := unassigned;
end;

{ Exemples d'initialisation des tob champs et fonctions paie }

procedure GetTobFunc(theTobFunc: TOB {a ne pas vider : a utiliser comme parent});
var
  T, TT1, TT2, TT3, TT4, TT5: TOB;
begin
  //  theTobFunc.ClearDetail; PT17
  // pour cacher l'onglet des fonctions, sortir APRES AVOIR VIDER LE DETAIL
  // exit;

  TT1 := TOB.Create('Calendrier', theTobFunc, -1);

  T := TOB.Create('Durée du contrat', TT1, -1);
  //T.AddchampSupValeur('TYPE','');
  T.AddchampSupValeur('NOM', 'Durée du contrat');
  T.AddchampSupValeur('PROTO', '@DUREECONTRAT ( PSA_SALARIE )');
  T.AddchampSupValeur('HELP', 'Calcul la durée du contrat');

  { DEB PT8 }
  T := TOB.Create('Durée du contrat en jour', TT1, -1);
  T.AddchampSupValeur('NOM', 'Durée du contrat en jour');
  T.AddchampSupValeur('PROTO', '@DUREECONTRATJOUR ( PSA_SALARIE )');
  T.AddchampSupValeur('HELP', 'Calcul la durée du contrat en jour');
  { FIN PT8 }

  { DEB PT29 }
  T := TOB.Create('Durée du contrat en semaine', TT1, -1);
  T.AddchampSupValeur('NOM', 'Durée du contrat en semaine');
  T.AddchampSupValeur('PROTO', '@DUREECONTRATSEMAINE ( PSA_SALARIE )');
  T.AddchampSupValeur('HELP', 'Calcul la durée du contrat en semaine');
  { FIN PT29 }

  T := TOB.Create('Durée d''essai', TT1, -1);
  //??  T.AddchampSupValeur('TYPE','');
  T.AddchampSupValeur('NOM', 'Durée d''essai');
  T.AddchampSupValeur('PROTO', '@DUREEESSAI ( PSA_SALARIE )');
  T.AddchampSupValeur('HELP', 'Calcul la durée de la période d''essai');

  //DEB PT29
  T := TOB.Create('Durée d''essai en semaine', TT1, -1);
  T.AddchampSupValeur('NOM', 'Durée d''essai en semaine');
  T.AddchampSupValeur('PROTO', '@DUREEESSAISEMAINE ( PSA_SALARIE )');
  T.AddchampSupValeur('HELP', 'Calcul la durée de la période d''essai en semaine');
  //FIN PT29

  T := TOB.Create('Durée du préavis', TT1, -1); { PT12 }
  //??  T.AddchampSupValeur('TYPE','');
 // PT30 T.AddchampSupValeur('NOM', 'Durée du préavis');
  T.AddchampSupValeur('NOM', 'Durée du préavis en mois complet');          // pt30
  T.AddchampSupValeur('PROTO', '@DUREEPREAVIS ( PSA_SALARIE )');
 // pt30 T.AddchampSupValeur('HELP', 'Calcul la durée du préavis');
  T.AddchampSupValeur('HELP', 'Calcul la durée du préavis en mois complet');      // pt30

  // deb pt30
   T := TOB.Create('Durée du préavis en jours', TT1, -1); { PT12 }
   T.AddchampSupValeur('NOM', 'Durée du préavis en jours');
   T.AddchampSupValeur('PROTO', '@DUREEPREAVISJOUR ( PSA_SALARIE )');
   T.AddchampSupValeur('HELP', 'Calcul la durée du préavis en jour');
  // fin pt30

  //DEB PT29
  T := TOB.Create('Durée du préavis en semaine', TT1, -1);
  T.AddchampSupValeur('NOM', 'Durée du préavis en semaine');
  T.AddchampSupValeur('PROTO', '@DUREEPREAVISSEMAINE ( PSA_SALARIE )');
  T.AddchampSupValeur('HELP', 'Calcul la durée du préavis en semaine');
  //FIN PT29

  T := TOB.Create('Horaire hebdomadaire', TT1, -1);
  //??  T.AddchampSupValeur('TYPE','');
  T.AddchampSupValeur('NOM', 'Horaire hebdomadaire');
  T.AddchampSupValeur('PROTO', '@GETCALENDRIER( PSA_SALARIE )');
  T.AddchampSupValeur('HELP', 'Horaire hebdomadaire');


  TT2 := TOB.Create('Utilitaire', theTobFunc, -1);
  //DEB PT-1-2
  T := TOB.Create('Libellé associé', TT2, -1);
  T.AddchampSupValeur('NOM', 'Libellé associé');
  T.AddchampSupValeur('PROTO', '@GETLIBELLE(''NOMCHAMP'',VALEUR )');
  T.AddchampSupValeur('HELP', 'Libellé associé');
  //FIN PT-1-2
  //DEB PT26
  T := TOB.Create('Elément national', TT2, -1);
  T.AddchampSupValeur('NOM', 'Elément national');
  T.AddchampSupValeur('PROTO', '@GETELTNAT(''CODE_ELEMENT'',PSA_SALARIE,PSA_ETABLISSEMENT,DATEDEB)');
  T.AddchampSupValeur('HELP', 'Elément national');
  //FIN PT26

  { DEB PT9 }
  TT2 := TOB.Create('Déclarant', theTobFunc, -1);
  T := TOB.Create('Nom du déclarant', TT2, -1);
  T.AddchampSupValeur('NOM', 'Nom du déclarant');
  T.AddchampSupValeur('PROTO', '$INTERLOCUTEUR');
  T.AddchampSupValeur('HELP', 'Nom du déclarant');
  T := TOB.Create('Qualité du déclarant', TT2, -1);
  T.AddchampSupValeur('NOM', 'Qualité du déclarant');
  T.AddchampSupValeur('PROTO', '$STATUT');
  T.AddchampSupValeur('HELP', 'Qualité du déclarant');
  T := TOB.Create('Civilité du déclarant', TT2, -1);
  T.AddchampSupValeur('NOM', 'Civilité du déclarant');
  T.AddchampSupValeur('PROTO', '$CIVILITE');
  T.AddchampSupValeur('HELP', 'Civilité du déclarant');
  T := TOB.Create('Sexe du déclarant', TT2, -1);
  T.AddchampSupValeur('NOM', 'Sexe du déclarant');
  T.AddchampSupValeur('PROTO', '$SEXE');
  T.AddchampSupValeur('HELP', 'Sexe du déclarant');
  T := TOB.Create('Ville du déclarant', TT2, -1);
  T.AddchampSupValeur('NOM', 'Ville du déclarant');
  T.AddchampSupValeur('PROTO', '$VILLE');
  T.AddchampSupValeur('HELP', 'Ville du déclarant');
  { FIN PT9 }

  { DEB PT10 }
  TT3 := TOB.Create('Grammaire Accord', theTobFunc, -1);
  T := TOB.Create('Accord sur sexe salarié', TT3, -1);
  T.AddchampSupValeur('NOM', 'Accord sur sexe salarié');
  T.AddchampSupValeur('PROTO', 'PSA_E');
  T.AddchampSupValeur('HELP', 'Accord sur sexe salarié');
  { FIN PT10 }

  { DEB PT13 }
  TT4 := TOB.Create('Organisme libre', theTobFunc, -1);
  T := TOB.Create('Organisme libre', TT4, -1);
  T.AddchampSupValeur('NOM', 'Organisme libre');
  T.AddchampSupValeur('PROTO', '@GETORGANISME( PSA_ETABLISSEMENT,''CODE_ORGANISME'',''CHAMP'')');
  T.AddchampSupValeur('HELP', 'Organisme libre');
  { FIN PT13 }
  { DEB PT14-1 }
  TT5 := TOB.Create('Historique de la paie', theTobFunc, -1);
  T := TOB.Create('Cumul de la paie', TT5, -1);
  T.AddchampSupValeur('NOM', 'Cumul');
  T.AddchampSupValeur('PROTO', '@GETCUMUL(DATEDEB,DATEFIN,PSA_SALARIE,''Cumul1;Cumul2'')');
  T.AddchampSupValeur('HELP', 'Cumul de la paie');

  T := TOB.Create('Base rémunération de la paie', TT5, -1);
  T.AddchampSupValeur('NOM', 'Base rémunération');
  T.AddchampSupValeur('PROTO', '@GETBASEREM(DATEDEB,DATEFIN,PSA_SALARIE,''REM1;REM2'')');
  T.AddchampSupValeur('HELP', 'Base rémunération de la paie');

  T := TOB.Create('Montant rémunération de la paie', TT5, -1);
  T.AddchampSupValeur('NOM', 'Montant rémunération');
  T.AddchampSupValeur('PROTO', '@GETMTREM(DATEDEB,DATEFIN,PSA_SALARIE,''REM1;REM2'')');
  T.AddchampSupValeur('HELP', 'Montant rémunération de la paie');

  T := TOB.Create('Base cotisation de la paie', TT5, -1);
  T.AddchampSupValeur('NOM', 'Base cotisation');
  T.AddchampSupValeur('PROTO', '@GETBASECOT(DATEDEB,DATEFIN,PSA_SALARIE,''COT1;COT2'')');
  T.AddchampSupValeur('HELP', 'Base cotisation de la paie');

  T := TOB.Create('Montant salarial de la paie', TT5, -1);
  T.AddchampSupValeur('NOM', 'Montant salarial');
  T.AddchampSupValeur('PROTO', '@GETMTSALCOT(DATEDEB,DATEFIN,PSA_SALARIE,''COT1;COT2'')');
  T.AddchampSupValeur('HELP', 'Montant salarial de la paie');

  T := TOB.Create('Montant patronal de la paie', TT5, -1);
  T.AddchampSupValeur('NOM', 'Montant patronal');
  T.AddchampSupValeur('PROTO', '@GETMTPATCOT(DATEDEB,DATEFIN,PSA_SALARIE,''COT1;COT2'')');
  T.AddchampSupValeur('HELP', 'Montant patronal de la paie');
  { FIN PT14-1 }



   { T := TOB.Create('ExisteLien',TT,-1);
  //??  T.AddchampSupValeur('TYPE','');
    T.AddchampSupValeur('NOM','EXISTELIEN');
    T.AddchampSupValeur('PROTO','@EXISTELIEN ( Quoi , Lequel )');
    T.AddchampSupValeur('HELP','Permet de savoir si un lien existe');
    ////
    T := TOB.Create('Fonction1',TT,-1);
  //??  T.AddchampSupValeur('TYPE','');
    T.AddchampSupValeur('NOM','FONCTION1');
    T.AddchampSupValeur('PROTO','@FONCTION1 ( Lequel )');
    T.AddchampSupValeur('HELP','N''importe quoi ce test...');   }
  //
end;

procedure GetTobDico(theTobDico: TOB);
var
  T1, T2, T3, T4, T5, T6, T7, T8, T9, T10: TOB;
  Q: TQuery;
  St: string;
  { DEB PT16 }
  function AddTobTable(libelle, laquelle: string): tob;
  begin
    result := TOB.Create(libelle, theTobDico, -1);
    Q := OpenSql('select dt_nomtable, dt_libelle from detables where dt_prefixe="' + laquelle + '"', True);
    result.SelectDB('', Q);
    result.PutValue('DT_LIBELLE', libelle);
    Ferme(Q);
  end;
  { FIN PT16 }
begin
  theTobDico.ClearDetail;
  //DEB PT-1-1
  //Table salarie
  T1 := AddTobTable('Les salariés', 'PSA'); { PT16 }
  T1.LoadDetailDB('DECHAMPS', '"PSA";', '', nil, false);
  //Champ etablissement
  T2 := AddTobTable('Les établissements', 'ET'); { PT16 }
  St := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE (DH_PREFIXE="ET"  ' +
    'AND (DH_NOMCHAMP="ET_ETABLISSEMENT" OR DH_NOMCHAMP="ET_ABREGE" ' +
    'OR DH_NOMCHAMP="ET_ACTIVITE" OR DH_NOMCHAMP="ET_ADRESSE1" ' +
    'OR DH_NOMCHAMP="ET_ADRESSE2" OR DH_NOMCHAMP="ET_ADRESSE3" ' +
    'OR DH_NOMCHAMP="ET_APE" OR DH_NOMCHAMP="ET_CODEPOSTAL" ' +
    'OR DH_NOMCHAMP="ET_DIVTERRIT" OR DH_NOMCHAMP="ET_EMAIL" ' +
    'OR DH_NOMCHAMP="ET_ETABLIE" OR DH_NOMCHAMP="ET_ETABLISSEMENT" ' +
    'OR DH_NOMCHAMP="ET_FAX" OR DH_NOMCHAMP="ET_LANGUE" ' +
    'OR DH_NOMCHAMP="ET_LIBELLE" OR DH_NOMCHAMP="ET_PAYS" ' +
    'OR DH_NOMCHAMP="ET_SIRET" OR DH_NOMCHAMP="ET_TELEPHONE" ' +
    'OR DH_NOMCHAMP="ET_TELEX" OR DH_NOMCHAMP="ET_VILLE" ' +
    'OR DH_NOMCHAMP="ET_JURIDIQUE") ) ' + { PT11 }
    'OR (DH_PREFIXE="ETB" AND DH_NOMCHAMP="ETB_HORAIREETABL") ';
  Q := OpenSql(St, True);
  T2.LoadDetailDB('les-etabl', '', '', Q, false);
  Ferme(Q);
  //Champ contrat de travail
  T3 := AddTobTable('Le contrat de travail', 'PCI'); { PT16 }
  T3.LoadDetailDB('DECHAMPS', '"PCI";', '', nil, false);
  //  T3.Detail.Sort('DH_NOMCHAMP'); { PT14-2}
  //Champ convention collective
  T4 := AddTobTable('La convention collective', 'PCV'); { PT16 }
  Q := OpenSql('SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS ' +
    'WHERE (DH_NOMCHAMP LIKE "PCV_CONVENTION" OR DH_NOMCHAMP LIKE "PCV_LIBELLE%")', True);
  T4.LoadDetailDB('la-conv-collec', '', '', Q, false);
  Ferme(Q);
  //Champ Organismes et caisses
  T5 := AddTobTable('L''organisme URSSAF', 'POG'); { PT16 }
  St := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="POG" ' +
    'AND (DH_NOMCHAMP="POG_ADRESSE1" OR DH_NOMCHAMP="POG_ADRESSE2" ' +
    'OR DH_NOMCHAMP="POG_ADRESSE3" OR DH_NOMCHAMP="POG_AUTREPERIODUCS" ' +
    'OR DH_NOMCHAMP="POG_CAISSEDESTIN" OR DH_NOMCHAMP="POG_CODEPOSTAL" ' +
    'OR DH_NOMCHAMP="POG_CONTACT" OR DH_NOMCHAMP="POG_EMAIL" ' +
    'OR DH_NOMCHAMP="POG_ETABLISSEMENT" OR DH_NOMCHAMP="POG_FAX" ' +
    'OR DH_NOMCHAMP="POG_GENERAL" OR DH_NOMCHAMP="POG_INSTITUTION" ' +
    'OR DH_NOMCHAMP="POG_LIBELLE" OR DH_NOMCHAMP="POG_NATUREORG" ' +
    'OR DH_NOMCHAMP="POG_NUMAFFILIATION" OR DH_NOMCHAMP="POG_NUMINTERNE" ' +
    'OR DH_NOMCHAMP="POG_ORGANISME" OR DH_NOMCHAMP="POG_PERIODICITDUCS" ' +
    'OR DH_NOMCHAMP="POG_SIRET" OR DH_NOMCHAMP="POG_TELEPHONE" ' +
    'OR DH_NOMCHAMP="POG_TELEX" OR DH_NOMCHAMP="POG_VILLE")';
  Q := OpenSql(St, True);
  T5.LoadDetailDB('l-urssaf', '', '', Q, false);
  Ferme(Q);
  //Champ société
  T6 := AddTobTable('La société', 'SO'); { PT16 }
  St := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="SO" ' +
    'AND (DH_NOMCHAMP="SO_ADRESSE1" OR DH_NOMCHAMP="SO_ADRESSE2" ' +
    'OR DH_NOMCHAMP="SO_ADRESSE3" OR DH_NOMCHAMP="SO_DIVTERRIT" ' +
    'OR DH_NOMCHAMP="SO_CODEPOSTAL" OR DH_NOMCHAMP="SO_CONTACT" ' +
    'OR DH_NOMCHAMP="SO_APE" OR DH_NOMCHAMP="SO_LIBELLE" OR DH_NOMCHAMP="SO_NIF" ' +
    'OR DH_NOMCHAMP="SO_FAX" OR DH_NOMCHAMP="SO_PAYS" OR DH_NOMCHAMP="SO_TELEX" ' +
    'OR DH_NOMCHAMP="SO_LIBELLE" OR DH_NOMCHAMP="SO_SIRET" OR DH_NOMCHAMP="SO_VILLE" ' +
    'OR DH_NOMCHAMP="SO_SOCIETE" OR DH_NOMCHAMP="SO_TELEPHONE" ' +
    'OR DH_NOMCHAMP="SO_CAPITAL" OR DH_NOMCHAMP="SO_RC")';
  Q := OpenSql(St, True);
  T6.LoadDetailDB('la-soc', '', '', Q, false);
  Ferme(Q);
  //FIN PT-1-1
  //MEDECINE          //DEBUT PT2
  T7 := AddTobTable('La visite médicale', 'PVM'); { PT16 }
  St := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="PVM" ' +
    'AND (DH_NOMCHAMP="PVM_SALARIE" OR DH_NOMCHAMP="PVM_DATEVISITE"' +
    ' OR DH_NOMCHAMP="PVM_TYPEVISITMED" OR DH_NOMCHAMP="PVM_EMDTRAV" OR DH_NOMCHAMP="PVM_HEUREVISITE")';
  Q := OpenSQL(St, True);
  T7.LoadDetailDB('visit-med', '', '', Q, false);
  Ferme(Q);
  //ANNUAIRE
  T8 := AddTobTable('L''adresse médicale', 'ANN'); { PT16 }
  St := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="ANN" ' +
    'AND (DH_NOMCHAMP="ANN_APNOM" OR DH_NOMCHAMP="ANN_APRUE1"' +
    ' OR DH_NOMCHAMP="ANN_APRUE2" OR DH_NOMCHAMP="ANN_APRUE3" OR DH_NOMCHAMP="ANN_APCPVILLE")';
  Q := OpenSQL(St, True);
  T8.LoadDetailDB('adress-med', '', '', Q, false);
  Ferme(Q); //FIN PT2
  //  PT5
  T9 := AddTobTable('Inscriptions aux formations', 'PFO'); { PT16 }
  St := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="PFO" ' +
    'AND (DH_NOMCHAMP="PFO_SALARIE" OR DH_NOMCHAMP="PFO_CODESTAGE"' +
    ' OR DH_NOMCHAMP="PFO_DATEDEBUT" OR DH_NOMCHAMP="PFO_DATEFIN"' +
    ' OR DH_NOMCHAMP="PFO_DEBUTDJ" OR DH_NOMCHAMP="PFO_FINDJ"' +
    ' OR DH_NOMCHAMP="PFO_HEUREDEBUT" OR DH_NOMCHAMP="PFO_HEUREFIN"' +
    ' OR DH_NOMCHAMP="PFO_LIEUFORM" OR DH_NOMCHAMP="PFO_NATUREFORM"' +
    ' OR DH_NOMCHAMP="PFO_FORMATION1" OR DH_NOMCHAMP="PFO_FORMATION2" OR DH_NOMCHAMP="PFO_FORMATION3"' +
    ' OR DH_NOMCHAMP="PFO_FORMATION4" OR DH_NOMCHAMP="PFO_FORMATION4" OR DH_NOMCHAMP="PFO_FORMATION6"' +
    ' OR DH_NOMCHAMP="PFO_FORMATION7" OR DH_NOMCHAMP="PFO_FORMATION8")';
  Q := OpenSQL(St, True);
  T9.LoadDetailDB('insc-form', '', '', Q, false);
  Ferme(Q);

  { DEB PT15 }
  T10 := AddTobTable('Le registre', 'ANN'); { PT16 }
  St := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="ANN" ' +
    'AND (DH_NOMCHAMP="ANN_SIREN" OR DH_NOMCHAMP="ANN_RM" OR DH_NOMCHAMP="ANN_RMDEP")';
  Q := OpenSQL(St, True);
  T10.LoadDetailDB('le-registre', '', '', Q, false);
  Ferme(Q);
  { FIN PT15 }


end;
// PT3 19/08/2002 PH V585 Ajout Fonction liens OLE pour  EXCEL
{
Fonction qui rend la valeur d'un ou plusieurs cumuls ou rubriques de rémunérations
ou de cotisations pour 1 Salarie,1 mois
ou 1 ensemble de salarié
Une fourchette de dates
Une selection d'etablissements,
une selection sur combo+etab+code stat
traitement des effectifs et des effectifs moyens en fonction des paies faites
}

function TOlePaiePGI.PaieCumulMois(
  Etab, Sal, LaTable, Quoi, Rub, TRV1, TRV2, TRV3, TRV4, LCMB1, LCMB2, LCMB3,
  LCMB4, CODESTAT, ZMoisD, ZAnneeD, ZMoisF, ZAnneeF, CatDucs, DadsProf: string; Present: string = ''): Variant;
var
  Q: TQuery;
  St, LePlus: string;
  Mois, Jour, Annee: WORD;
  ZDate, LaDate: TDateTime;
  i: Integer;
  LeAnd: Boolean;
  Trv: array[1..4] of string;
  Lcmb: array[1..4] of string;
  Pref: string; // Prefixe de la table et nature de la rubrique
  LeQuoi: string; // Zone à cumuler dans histobulletin
  Nature: string; // Nature des rubriques COT ou REM
  LeJoin: string; // Jointure sur la table salarié pour test de la date de sortie
  PMois, PAn, NbMois: Word;
//PT25
  mdossiers, lstdossier : string;
  resdos: TQuery;
  resmul: Double;
begin
  // PT7 09/03/2004 PH V_50  Mise en place des concepts de la paie ie confidentialité
//  if not PGJaiLeDroitTag(282300) then exit;
  if (PGControlConcept () = FALSE) then Exit;
  if (LaTable <> 'CUM') and (LaTable <> 'COT') and (LaTable <> 'REM') and (LaTable <> 'PAI') then
  begin
    V_PGI.StOLEErr := PgOLEErreur(5);
    Exit;
  end;
  if (ZMoisD = '') or (ZAnneeD = '') or (ZMoisF = '') or (ZAnneeF = '') then
  begin
    V_PGI.StOLEErr := PgOLEErreur(3);
    Exit;
  end; // Date obligatoire

  if LaTable = 'COT' then Nature := LaTable
  else if LaTable = 'REM' then Nature := 'AAA'
  else if LaTable = 'CUM' then Nature := 'CUM'
  else if LaTable = 'PAI' then Nature := 'PAI';
  if Quoi = '' then
  begin
    if LaTable = 'CUM' then LeQuoi := 'MT' // on ne prend que le montant du cumul
    else V_PGI.StOLEErr := PgOLEErreur(2);
    Exit;
  end
  else
  begin
    if Quoi = 'BR' then LeQuoi := 'PHB_BASEREM' // Base rémunération
    else if Quoi = 'MR' then LeQuoi := 'PHB_MTREM' // montant rémunération
    else if Quoi = 'BC' then LeQuoi := 'PHB_BASECOT' // base de cotisation
    else if Quoi = 'MP' then LeQuoi := 'PHB_MTPATRONAL' // Montant patronal de cotisation
    else if Quoi = 'MS' then LeQuoi := 'PHB_MTSALARIAL' // Montant salarial de cotisation
    else if Quoi = 'MC' then LeQuoi := 'PHC_MONTANT' // Montant cumul
    else if Quoi = 'EFF' then LeQuoi := 'Effectif'
    else if Quoi = 'EFM' then LeQuoi := 'Effectif moyen'
    else
    begin
      V_PGI.StOLEErr := PgOLEErreur(5);
      exit;
    end;
  end;
  if ZMoisD = '' then
  begin
    V_PGI.StOLEErr := PgOLEErreur(3);
    Exit;
  end;
  if ZAnneeD = '' then
  begin
    V_PGI.StOLEErr := PgOLEErreur(4);
    Exit;
  end;
  for i := 1 to 4 do
  begin
    if (i = 1) then Trv[i] := TRV1 else
      if (i = 2) then Trv[i] := TRV2 else
        if (i = 3) then Trv[i] := TRV3 else
          if (i = 4) then Trv[i] := TRV4;
    if (i = 1) then Lcmb[i] := LCMB1 else
      if (i = 2) then Lcmb[i] := LCMB2 else
        if (i = 3) then Lcmb[i] := LCMB3 else
          if (i = 4) then Lcmb[i] := LCMB4;
  end;
  LeJoin := '';
  if LaTable = 'CUM' then Pref := 'PHC'
  else if LaTable = 'PAI' then Pref := 'PPU'
  else Pref := 'PHB';

  if (Present = 'O') and (LaTable <> 'PAI') then
    LeJoin := ' LEFT JOIN PAIEENCOURS ON PPU_SALARIE=' +
      Pref + '_SALARIE AND PPU_DATEDEBUT=' + Pref + '_DATEDEBUT AND PPU_DATEFIN=' +
      Pref + '_DATEFIN ';
  if (CatDucs <> '') or (DadsProf <> '') then LeJoin := LeJoin + ' LEFT JOIN SALARIES ON PSA_SALARIE=' + Pref + '_SALARIE ';
  if LaTable = 'CUM' then
  begin
    Pref := 'PHC';
    LeAnd := FALSE;
    St := 'SELECT SUM (PHC_MONTANT) MT FROM HISTOCUMSAL ' + LeJoin + ' WHERE ';
    ConfectSqlOle(Pref, 'CUMULPAIE', Rub, St, LeAnd);
  end
  else
  begin
    if LaTable = 'PAI' then
    begin
      Pref := 'PPU';
      if Quoi = 'EFF' then
        St := 'SELECT Count (DISTINCT(PPU_SALARIE)) MT FROM PAIEENCOURS ' + LeJoin
      else
      begin
        LePLus := RendSQLMois('PPU_DATEFIN');
        St := 'SELECT Count (DISTINCT(PPU_SALARIE)) MT,' + LePlus + ' DF FROM PAIEENCOURS ' + LeJoin;
      end;
      LeAnd := FALSE;
      St := St + ' WHERE ';
      if CatDucs <> '' then ConfectSqlOle(Pref, 'CATDADS', Rub, St, LeAnd);
      if DadsProf <> '' then ConfectSqlOle(Pref, 'DADSPROF', Rub, St, LeAnd);
    end
    else
    begin
      Pref := 'PHB';
      St := 'SELECT SUM (' + LeQuoi + ') MT FROM HISTOBULLETIN ' + LeJoin;
      st := St + ' WHERE PHB_NATURERUB="' + Nature + '" ';
      LeAnd := TRUE;
      ConfectSqlOle(Pref, 'RUBRIQUE', Rub, St, LeAnd);
    end;
  end;
  if Present = 'O' then
  begin
    if LeAnd then st := st + ' AND ';
    st := st + ' ((PPU_DATESORTIE > PPU_DATEFIN) OR (PPU_DATESORTIE ="' + USDateTime(IDate1900) + '" OR PPU_DATESORTIE IS NULL))';
    LeAnd := TRUE;
  end;
  if (Sal <> '') and (LaTable <> 'PAI') then
  begin
    if LeAnd then St := St + ' AND ' + Pref + '_SALARIE="' + Sal + '" '
    else St := Pref + '_SALARIE="' + Sal + '" ';
    LeAnd := TRUE;
  end;
  if LaTable <> 'PAI' then LeAnd := TRUE;
  for i := 1 to 4 do
  begin
    ConfectSqlOle(Pref, 'TRAVAILN' + IntToStr(i), Trv[i], St, LeAnd);
    ConfectSqlOle(Pref, 'LIBREPCMB' + IntToStr(i), Lcmb[i], St, LeAnd);
  end;

  ConfectSqlOle(Pref, 'CODESTAT', CODESTAT, St, LeAnd);
  ConfectSqlOle(Pref, 'ETABLISSEMENT', Etab, St, LeAnd);

  Jour := 1;
  Mois := StrToInt(Copy(ZMoisD, 1, 2));
  if (Mois < 1) or (Mois > 12) then exit;
  Annee := StrToInt(Copy(ZAnneeD, 1, 4));
  ZDate := EncodeDate(Annee, Mois, Jour);
  if LeAnd then st := St + ' AND ' + Pref + '_DATEFIN >="' + USDateTime(ZDate) + '"'
  else st := St + ' ' + Pref + '_DATEFIN >="' + USDateTime(ZDate) + '"';
  if LaTable = 'PAI' then LaDate := Zdate
  else LaDate := Idate1900; //memorisation de la date de debut
  if (ZMoisF <> '') and (ZAnneeF <> '') then
  begin // Traitement borne de fin de la periode
    Jour := 1;
    Mois := StrToInt(Copy(ZMoisF, 1, 2));
    if (Mois < 1) or (Mois > 12) then exit;
    Annee := StrToInt(Copy(ZAnneeF, 1, 4));
    ZDate := EncodeDate(Annee, Mois, Jour);
  end;
  ZDate := FINDEMOIS(ZDate);

  NOMBREMOIS(LaDate, ZDate, PMois, PAn, NbMois);
  if LaTable <> 'CON' then st := St + ' AND ' + Pref + '_DATEFIN <="' + USDateTime(ZDate) + '"';
  if Quoi = 'EFM' then st := St + ' GROUP BY ' + LePlus;

  try
//Début PT25
    mdossiers := GetBasesMS(RegroupementSociete);
    if mdossiers = '' then begin
      //Si non regroupement : comme avant
      Q := OpenSQL(st, TRUE);
      if (not Q.EOF) and (Q <> nil) then
      begin
        if (Quoi <> 'EFM') then
        begin
          if not Q.EOF then result := Q.FindField('MT').AsFloat
          else result := 0;
        end
        else
        begin
          result := 0;
          while not Q.Eof do
          begin // Boucle de 12 maxi
            result := result + Q.FindField('MT').AsFloat;
            Q.next;
          end;
          result := Arrondi(result / NbMois, 2); // Effectif moyen
        end;
      end
      else result := 0;
      Ferme(Q);
      end
    else begin
      //Regroupement : récupération de sélection éventuelle de bases
        resmul := 0;
        result := 0;
        if listebases <> '' then mdossiers := listebases;
        while (mdossiers <> '') do
        begin
          lstdossier :=  ReadTokenSt(mdossiers);
          resdos := OpenSelect(st,lstdossier);
          If (not resdos.EOF) and (resdos <> nil) then resmul := resdos.FindField('MT').AsFloat else resmul := 0;
          result := resmul + result;
          if assigned(resdos) then ferme(resdos);
        end;

    end
  except
    V_PGI.StOLEErr := PgOLEErreur(10);
    result := 0;
  end;
end;
// PT3 19/08/2002 PH V585 Ajout Fonction liens OLE pour  EXCEL
{
Fonction qui rend la valeur d'un ou plusieurs cumuls ou rubriques de rémunérations
ou de cotisations pour 1 Salarie,1 mois
ou 1 ensemble de salarié
Une fourchette de dates
Une selection d'etablissements,
une selection sur combo+etab+code stat
traitement des effectifs et des effectifs moyens en fonction des paies faites
}
// PT19 Rajout SommeEff pour le calcul unité somme de l'effectif
function TOlePaiePGI.PaieStatEffectif(Etab, TRV1, TRV2, TRV3, TRV4, LCMB1, LCMB2, LCMB3,
  LCMB4, CodeStat, ZMoisF, ZAnneeF, CatDucs, DadsProf, DadsCat, Sexe: string;
  LibEmpl: string = ''; LibCoeff: string = ''; LibInd: string = ''; LibNiv: string = ''; LibQual: string = ''; CodePCS: string = ''; Contrat: string = ''; EntSort: string = '';
  MotifCont: string = ''; SommeEff : String =''): Variant;
var
  Q: TQuery;
  St: string;
  Mois, Jour, Annee: WORD;
  ZDate: TDateTime;
  i: Integer;
  LeAnd: Boolean;
  Trv: array[1..4] of string;
  Lcmb: array[1..4] of string;
  Pref: string; // Prefixe de la table et nature de la rubrique
//PT25
  mdossiers, lstdossier : string;
  resdos: TQuery;
  resmul: Double;
begin
  if (PGControlConcept () = FALSE) then Exit;
  if (ZMoisF = '') or (ZAnneeF = '') then exit
  else
  begin // Traitement borne de fin de la periode
    Jour := 1;
    Mois := StrToInt(Copy(ZMoisF, 1, 2));
    if (Mois < 1) or (Mois > 12) then exit;
    Annee := StrToInt(Copy(ZAnneeF, 1, 4));
    ZDate := EncodeDate(Annee, Mois, Jour);
  end;

  for i := 1 to 4 do
  begin
    if (i = 1) then Trv[i] := TRV1 else
      if (i = 2) then Trv[i] := TRV2 else
        if (i = 3) then Trv[i] := TRV3 else
          if (i = 4) then Trv[i] := TRV4;
    if (i = 1) then Lcmb[i] := LCMB1 else
      if (i = 2) then Lcmb[i] := LCMB2 else
        if (i = 3) then Lcmb[i] := LCMB3 else
          if (i = 4) then Lcmb[i] := LCMB4;
  end;
  Pref := 'PSA';

  if Contrat = '' then
  begin // DEB PT19
    if SommeEff = '' then St := 'SELECT COUNT (*) MT FROM SALARIES WHERE '
    else St := 'SELECT SUM(PSA_UNITEPRISEFF) MT FROM SALARIES WHERE ';
  end // FIN PT19
  else st := 'SELECT COUNT(*) FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PSA_SALARIE=PCI_SALARIE WHERE';

  LeAnd := FALSE;
  if CatDucs <> '' then ConfectSqlOle(Pref, 'CATDADS', CatDucs, St, LeAnd);
  if DadsProf <> '' then ConfectSqlOle(Pref, 'DADSPROF', DadsProf, St, LeAnd);
  if DadsCat <> '' then ConfectSqlOle(Pref, 'DADSCAT', DadsCat, St, LeAnd);
  if CodeStat <> '' then ConfectSqlOle(Pref, 'CODESTAT', CodeStat, St, LeAnd);
  if Sexe <> '' then ConfectSqlOle(Pref, 'SEXE', Sexe, St, LeAnd);
  if LibEmpl <> '' then ConfectSqlOle(Pref, 'LIBELLEEMPLOI', Sexe, St, LeAnd);
  if LibNiv <> '' then ConfectSqlOle(Pref, 'NIVEAU', LibNiv, St, LeAnd);
  if LibInd <> '' then ConfectSqlOle(Pref, 'INDICE', LibInd, St, LeAnd);
  if LibCoeff <> '' then ConfectSqlOle(Pref, 'COEFFICIENT', LibCoeff, St, LeAnd);
  if LibQual <> '' then ConfectSqlOle(Pref, 'QUALIFICATION', LibQual, St, LeAnd);
  if CodePCS <> '' then ConfectSqlOle(Pref, 'CODEPCS', CodePCS, St, LeAnd);

  for i := 1 to 4 do
  begin
    ConfectSqlOle(Pref, 'TRAVAILN' + IntToStr(i), Trv[i], St, LeAnd);
    ConfectSqlOle(Pref, 'LIBREPCMB' + IntToStr(i), Lcmb[i], St, LeAnd);
  end;

  ConfectSqlOle(Pref, 'ETABLISSEMENT', Etab, St, LeAnd);
  if Contrat <> '' then
  begin
    if LeAnd then St := st + ' AND ';
    St := st + ' PCI_TYPECONTRAT="' + Contrat + '" ';
  end;

  ZDate := FINDEMOIS(ZDate);

  if (MotifCont <> '') and (Contrat <> '') and (EntSort <> '') then
  begin
    if LeAnd then St := st + ' AND ';
    if EntSort = 'S' then St := st + ' PSA_MOTIFSORTIE="' + MotifCont + '" '
    else St := st + ' PSA_MOTIFENTREE="' + MotifCont + '" ';
  end;

  if LeAnd then St := st + ' AND ';
  if Contrat = '' then
    St := St + '(PSA_DATEENTREE <="' + UsDateTime(ZDate) + '") AND ((PSA_DATESORTIE IS NULL) OR ' +
      '(PSA_DATESORTIE <= "' + UsDateTime(iDate1900) + '") OR (PSA_DATESORTIE >="' + UsDateTime(FinDeMois(ZDate)) + '"))'
  else
  begin
    if EntSort = 'S' then // Cas de sortie des salariés donc obligation de renseigner les contrats de travail
      St := St + ' (PCI_FINCONTRAT <="' + UsDateTime(ZDate) + '") AND (PCI_FINCONTRAT >="' + UsDateTime(DebutDeMois(ZDate)) + '")' +
        ' OR ((PCI_FINCONTRAT IS NULL) OR ' + '(PCI_FINCONTRAT = "' + UsDateTime(iDate1900) + '"'
    else
      St := St + ' (PCI_DEBUTCONTRAT <="' + UsDateTime(ZDate) + '") AND (PCI_DEBUTCONTRAT >="' + UsDateTime(DebutDeMois(ZDate)) + '")';
  end;


  try
//Début PT25
      if RegroupementSociete<>'' then mdossiers := GetBasesMS(RegroupementSociete);
      if mdossiers = '' then begin
        //Non regroupement : comme avant
        Q := OpenSQL(st, TRUE);
        if (not Q.EOF) and (Q <> nil) then
        begin
          if not Q.EOF then result := Q.FindField('MT').AsFloat
          else result := 0;
        end
        else result := 0;
        Ferme(Q);
      end
      else begin
        //Regroupement récupération selection éventuelle de Bases
        resmul := 0;
        result := 0;
        if listebases <> '' then mdossiers := listebases;
        while (mdossiers <> '') do
        begin
          lstdossier :=  ReadTokenSt(mdossiers);
          resdos := OpenSelect(st,lstdossier);
          If (not resdos.EOF) and (resdos <> nil) then resmul := resdos.FindField('MT').AsFloat else resmul := 0;
          result := resmul + result;
          if assigned(resdos) then ferme(resdos);
        end;
     end;
//Fin PT25
  except
    V_PGI.StOLEErr := PgOLEErreur(10);
    result := 0;
  end;
end;
{ Fonction qui rend le contenu d'un champ.
Attention, tous les champs de type BLOB ou DATA sont ignorés
}

function TOlePaiePGI.PGOleGetInfoTable(Table, Sal, ChampSelect, Dollar: string): Variant;
var
  T1: TOB;
  Pref, LeTyp, RR: string;
  iTable, iChamp: Integer;
begin
  if (PGControlConcept () = FALSE) then Exit;
  if Table = 'SALARIES' then Pref := 'PSA'
  else if Table = 'CONTRAT' then Pref := 'PCI'
  else
  begin
    result := '';
    exit;
  end;

  if PgOleSal <> Sal then PGOleAllocTob(Table, Sal);
  if Table = 'SALARIES' then T1 := OleT_Sal
  else if Table = 'CONTRAT' then T1 := OleT_Contrat
  else T1 := nil; // V_42 Qualité
  if T1 = nil then
  begin
    result := '';
    exit;
  end; // V_42 Qualité


  iTable := PrefixeToNum(Pref);
  ChargeDeChamps(iTable, Pref);
  //iChamp:=T1.GetNumChamp(Pref+'_'+ChampSelect) ;
  for iChamp := 1 to high(V_PGI.DeChamps[iTable]) do
    if V_PGI.DEChamps[iTable, iChamp].nom = Pref + '_' + ChampSelect then break;

  LeTyp := V_PGI.DEChamps[iTable, iChamp].tipe;

  try
    if (LeTyp = 'BLOB') or (LeTyp = 'DATA') then Result := ' '
    else result := T1.GetValue(Pref + '_' + ChampSelect);
    if ((Dollar = '$') or (Dollar = '$$')) and ((LeTyp = 'COMBO') or (Pos('VARCHAR',LeTyp)>0)) then { PT24 }
    begin
      RR := RechDom(Get_Join(ChampSelect), Result, FALSE);
      if (RR <> '') and (RR <> 'Error') then
      begin
        if Dollar = '$' then result := RR
        else result := result + ' ' + RR;
      end;
      if Result = 'Error' then Result := PgOLEErreur(7);
    end;
  except
    if (LeTyp = 'INTEGER') or (LeTyp = 'DOUBLE') or (LeTyp = 'RATE') then Result := 0
    else if (LeTyp = 'DATE') then Result := Idate1900
    else if (LeTyp = 'BLOB') or (LeTyp = 'DATA') then Result := ' '
    else result := '';
  end;
end;
// FIN PT3

{PT4 19/08/2002 PH V585 Ajout Fonctions publiques pour le bilan social et les liens
                    OLE pour  EXCEL

}

procedure InitBSAnciennete(StWhere: string);
var
  QRech: TQuery;
  RegulMois, i: Integer;
  DateAnciennete: TDateTime;
  TA: TOB;
begin
  if TOB_AncienneteBS <> nil then FreeBSAnciennete;
  TOB_AncienneteBS := TOB.Create('Le calcul ancienneté', nil, -1);
  QRech := OpenSql('SELECT PSA_REGULANCIEN,PSA_DATEANCIENNETE,PSA_DATEENTREE,PSA_DADSPROF FROM SALARIES WHERE ' + StWhere, True);
  TOB_AncienneteBS.LoadDetailDB('SALARIES', '', '', QRech, False);
  Ferme(QRech);
  for i := 0 to TOB_AncienneteBS.detail.count - 1 do
  begin
    TA := TOB_AncienneteBS.detail[i];
    if TA.GetValue('PSA_DATEANCIENNETE') <= 1 then DateAnciennete := TA.GetValue('PSA_DATEENTREE')
    else DateAnciennete := TA.GetValue('PSA_DATEANCIENNETE');
    RegulMois := TA.GetValue('PSA_REGULANCIEN');
    DateAnciennete := PLUSMOIS(DateAnciennete, RegulMois);
    TA.AddChampSup('ANCIENNETE', FALSE);
    TA.PutValue('ANCIENNETE', DateAnciennete);
  end;
end;

procedure FreeBSAnciennete;
begin
  if TOB_AncienneteBS <> nil then
  begin
    TOB_AncienneteBS.Free;
    TOB_AncienneteBS := nil;
  end;
end;

function RendBSAge(V1, V2: WORD; DF: TDateTime): string;
var
  D1, D2: TDateTime;
begin
  D1 := PLUSMOIS(DF, V1 * (-12));
  D2 := PLUSMOIS(DF, V2 * (-12));
  result := ' AND PSA_DATENAISSANCE <= "' + UsDateTime(D1) + '" AND PSA_DATENAISSANCE > "' +
    UsDateTime(D2) + '"';
end;

function RendBSAnciennete(V1, V2: WORD; DD, DF: TDateTime; Categorie: string): integer;
var
  i, Nbre: Integer;
  DateAnciennete: TDateTime;
  PremMois, PremAnnee, NbMois: Word;
  T1: TOB;
  cat, StWhere: string;
  Ok: Boolean;
begin
  if TOB_AncienneteBS = nil then
  begin
    StWhere := RendBSLeWhere(DD, DF, '', '', '');
    InitBSAnciennete(StWhere);
  end;
  PremMois := 0;
  PremAnnee := 0;
  // result:=0; // V_42 Qualité
  Nbre := 0;
  for i := 0 to TOB_AncienneteBS.detail.count - 1 do
  begin
    T1 := TOB_AncienneteBS.detail[i];
    cat := T1.getValue('PSA_DADSPROF');
    Ok := FALSE;
    if (Categorie = 'O') and (cat = '01') then Ok := TRUE
    else if (Categorie = 'A') and (cat = '04') then Ok := TRUE
    else if (Categorie = 'T') and (cat = '03') then Ok := TRUE
    else if (Categorie = 'E') and (cat = '02') then Ok := TRUE;
    if (Categorie = 'C') and ((cat = '12') or (cat = '13')) then Ok := TRUE;
    if (Categorie = 'D') and ((cat > '04') and (cat <> '12') and (cat <> '13')) then Ok := TRUE;
    if Ok then
    begin
      DateAnciennete := T1.GetValue('ANCIENNETE');
      NOMBREMOIS(DateAnciennete, DF, PremMois, PremAnnee, NbMois);
      if NbMois <> 0 then NbMois := NbMois div 12; // Donne nbre annees
      if (NbMois >= V1) and (NbMois < V2) then
        Nbre := Nbre + 1;
    end;
  end;
  result := Nbre;
end;

function RendBSLeWhere(DD, DF: TDateTime; Champ, LaValeur, Categorie: string): string;
var
  StWhere, cat: string;
begin
  if (not (isValidDate(DateToStr(DD)))) or (not (isValidDate(DateToStr(DD)))) then exit;

  StWhere := '(PSA_DATEENTREE <="' + USDateTime(DF) + '") AND ((PSA_DATESORTIE >="' + USDateTime(DD) + '")' +
    ' OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "' + USDateTime(IDate1900) + '"))';
  cat := '';
  if Categorie = 'O' then cat := '01'
  else if Categorie = 'A' then cat := '04'
  else if Categorie = 'T' then cat := '03'
  else if Categorie = 'E' then cat := '02';
  if cat <> '' then StWhere := StWhere + ' AND (PSA_DADSPROF = "' + Cat + '")';
  if Categorie = 'C' then StWhere := StWhere + ' AND NOT((PSA_DADSPROF <> "12") AND (PSA_DADSPROF <> "13"))'
  else
    if Categorie = 'D' then StWhere := StWhere + ' AND ((PSA_DADSPROF > "04") AND (PSA_DADSPROF <> "12") AND (PSA_DADSPROF <> "13"))';
  if Champ <> '' then
  begin
    if (Copy(LaValeur, 1, 1) = '!') then
    begin
      StWhere := StWhere + ' AND NOT ';
      LaValeur := Copy(LaValeur, 2, Length(LaValeur));
    end
    else StWhere := StWhere + ' AND ';
    StWhere := StWhere + '(PSA_' + Champ + '="' + LaValeur + '")';
  end;
  result := StWhere;
end;

function ExecBSSQL(LeSQL: string): Integer;
var
  Q: TQuery;
begin
  Q := OpenSql('SELECT COUNT(*) NBRE FROM SALARIES WHERE ' + LeSQL, True);
  if not Q.eof then result := Q.FindField('NBRE').asinteger
  else result := 0;
  Ferme(Q);
end;
// fonction qui rend le montant cumulé

function ExecBSSumTP(LeWhere: string): Extended;  //PT27 Calcul de l'effectif équivalent temps plein
var
  Q: TQuery;
  tempFloat : Extended;
begin
  result := 0;
  Q := OpenSql('SELECT PSA_TAUXPARTIEL/100 FROM SALARIES WHERE ' + LeWhere, True);
  While not Q.eof do
  begin
    tempFloat := Q.Fields[0].AsFloat;
    if tempFloat <> 0 then
      result := result + tempFloat
    else
      result := result+1;
    Q.Next;
  end;
  Ferme(Q);
end;


function RendBSCumul(DD, DF: TDateTime; Cumul, LeWhere: string): Double;
var
  st: string;
  Q: TQuery;
begin
  result := 0;
  if (Cumul = '') or (LeWhere = '') then exit;
  St := 'SELECT sum(PHC_MONTANT) MT from histocumsal left join salaries on PSA_SALARIE=PHC_SALARIE WHERE PHC_DATEDEBUT >= "';
  St := St + USDateTime(DD) + '" AND PHC_DATEFIN <="' + USDateTime(DF) + '" AND PHC_CUMULPAIE="' + CUMUL + '"';
  St := St + ' AND ' + LeWhere;
  Q := OpenSql(St, True);
  if not Q.eof then result := Q.FindField('MT').asFloat;
  Ferme(Q);
end;

function RendBSAbsences(DD, DF: TDateTime; Motif, LeWhere: string): Double;
var
  st: string;
  Q: TQuery;
begin
  result := 0;
  if (Motif = '') or (LeWhere = '') then exit;
  st := 'SELECT sum(PCN_JOURS) JOURS FROM ABSENCESALARIE LEFT JOIN SALARIES ON PSA_SALARIE=PCN_SALARIE WHERE ';
  st := st + ' PCN_ETATPOSTPAIE <> "NAN" '+ { PT22 }
  'AND ((PCN_DATEDEBUTABS >="' + USDateTime(DD) + '" ) AND (PCN_DATEFINABS <="' + USDateTime(DF) + '")) AND ';
  if Motif = 'PRI' then st := st + ' (PCN_MVTDUPLIQUE<>"X" AND PCN_TYPEMVT = "CPA" AND PCN_TYPECONGE="PRI")'
  else
  begin
    st := st + ' (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-") AND ';
    if Motif <> 'Autres' then st := st + '(PCN_TYPECONGE="' + Motif + '")'
    else
    begin
      st := st + '((PCN_TYPECONGE <> "PRI") AND (PCN_TYPECONGE <> "' + VH_Paie.PgBSAbs1 + '")';
      st := st + ' AND (PCN_TYPECONGE <> "' + VH_Paie.PgBSAbs2 + '"))';
    end;
  end;
  st := st + ' AND ' + LeWhere;
  Q := OpenSql(St, True);
  if not Q.eof then result := Q.FindField('JOURS').asFloat;
  Ferme(Q);
end;
// FIN PT4
{ verble OLE publié pour bilan social simplifié
Quoi = type de donnée (sexe, age, ancienneté, nationalite ...)
val et val1 valeurs de bornage --> val1 pour être nulle
deb et fin date de début et de fin de la sélection
categorie = PSA_DADSPROF pour identifier la colonne
}

function TOlePaiePGI.PGBSSINFO(Quoi, val, val1, deb, fin, categorie: string): Variant;
var
  StWhere: string;
begin
  if (PGControlConcept () = FALSE) then Exit;
  result := 0;
  StWhere := '';
  if (Quoi = '') or (deb = '') or (Fin = '') or (categorie = '') then exit;
  if Quoi = 'SEX' then StWhere := RendBSLeWhere(StrToDate(Deb), StrToDate(Fin), 'SEXE', Val, Categorie)
  else if Quoi = 'NAT' then StWhere := RendBSLeWhere(StrToDate(Deb), StrToDate(Fin), 'NATIONALITE', Val, Categorie)
  else if Quoi = 'EMP' then StWhere := RendBSLeWhere(StrToDate(Deb), StrToDate(Fin), 'CONDEMPLOI', Val, Categorie);
  if Quoi = 'AGE' then
  begin
    StWhere := RendBSLeWhere(StrToDate(deb), StrToDate(Fin), '', '', Categorie);
    StWhere := StWhere + RendBSAge(StrToInt(Val), StrToInt(Val1), StrToDate(Fin));
  end;
  if StWhere <> '' then result := ExecBSSQL(StWhere);
  if Quoi = 'ETP' then //PT27 effectifs équivalent temps plein
  begin
    StWhere := RendBSLeWhere(StrToDate(Deb), StrToDate(Fin), '', '', Categorie);
    result := ExecBSSumTP(StWhere);
  end;
  if Quoi = 'ANC' then
  begin
    if TOB_AncienneteBS <> NIL then FreeAndNil (TOB_AncienneteBS); // PT23
    result := RendBSAnciennete(StrToInt(Val), StrToInt(Val1), StrToDate(Deb), StrToDate(Fin), Categorie);
  end
  else
  begin
    if Quoi = 'CUM' then
    begin
      StWhere := RendBSLeWhere(StrToDate(Deb), StrToDate(Fin), '', '', Categorie);
      result := RendBSCumul(StrToDate(Deb), StrToDate(Fin), val, StWhere);
    end
    else
    begin
      if Quoi = 'ABS' then
      begin
        if val = '0' then val := 'PRI'
        else if val = '1' then val := VH_Paie.PgBSAbs1
        else if val = '2' then val := VH_Paie.PgBSAbs2
        else if val = 'A' then val := 'Autres';
        StWhere := RendBSLeWhere(StrToDate(Deb), StrToDate(Fin), '', '', Categorie);
        result := RendBSAbsences(StrToDate(Deb), StrToDate(Fin), val, StWhere);
      end;
    end;
  end;
end;

function RendSQLMois(LaDate: string): string;
begin
  { DEB PT21 }
  if PGisOracle then result := ' To_Char (' + LaDate + ', "mm")'
  else if PGisMssql or PGisSYBASE then result := ' datepart(Month, ' + LaDate + ')'
  else result := ' month (' + LaDate + ')';
  { FIN PT21 }
  {case V_PGI.Driver of PT21 Mise en commentaire
    dbORACLE7, dbORACLE8: result := ' To_Char (' + LaDate + ', "mm")';
    dbMSSQL, dbMSSQL2005,dbSYBASE: result := ' datepart(Month, ' + LaDate + ')';  //PT20
  else result := ' month (' + LaDate + ')';
  end;  }
end;


//DEB PT6

constructor TOlePaiePGI.Create;
begin
  inherited create;
  Automation.OnLastRelease:= OnTheLast;	//PT28
  RegroupementSociete := '';
  ListeBases := '';
  IndexReg := -1;
end;

destructor TOlePaiePGI.Destroy;
begin
  inherited Destroy;

end;
//FIN PT6

//PT28
procedure TOlePaiePGI.OnTheLast (var shutdown: Boolean);
begin
  ShutDown:=(FindCmdLineSwitch('REGSERVER') OR FindCmdLineSwitch('UNREGSERVER'));
end;
//FIN PT28

function TOlePaiePGI.PGAssistOle(st: string): string;
begin
  InitZoomOLE; { PT18 }
  MakeZoomOle(V_PGI.ZoomOleHwnd); { PT18 }
  result := LancePgAssistExcelOle(self);
  FinZoomOLE; { PT18 }
end;

function TOlePaiePGI.PAIEBSCalIndOLE(DateDeb, Datefin, Pres, Ind, Etab,
  Categ, Emploi, Coeff, Qualif, CodeStat, TravailN1, TravailN2, TravailN3,
  TravailN4, TabLibre1, TabLibre2, TabLibre3, TabLibre4: string): Variant;
var
  Q: TQuery;
  st, StWHere, StInd: string;
//PT25
  mdossiers, lstdossier : string;
  resdos: TQuery;
  resmul: Double;
begin
  result := 'Erreur';
  if (PGControlConcept () = FALSE) then Exit;
  if not IsValidDate(DateDeb) then exit;
  if not IsValidDate(DateFin) then exit;
  if Pres = '' then exit;
  if Ind = '' then exit;
  StWHere := RecupClauseSql(Pres, 'PBC_BSPRESENTATION');
  StWHere := StWhere + RecupClauseSql(Ind, 'PBC_INDICATEURBS');
  if Etab <> '' then StWHere := StWhere + RecupClauseSql(Etab, 'PBC_ETABLISSEMENT');
  if Categ <> '' then StWHere := StWhere + RecupClauseSql(Categ, 'PBC_CATBILAN');

  if Emploi <> '' then StWHere := StWhere + RecupClauseSql(Emploi, 'PBC_LIBELLEEMPLOI');
  if Coeff <> '' then StWHere := StWhere + RecupClauseSql(Coeff, 'PBC_COEFFICIENT');
  if Qualif <> '' then StWHere := StWhere + RecupClauseSql(Qualif, 'PBC_QUALIFICATION');
  if CodeStat <> '' then StWHere := StWhere + RecupClauseSql(CodeStat, 'PBC_CODESTAT');
  if TravailN1 <> '' then StWHere := StWhere + RecupClauseSql(TravailN1, 'PBC_TRAVAILN1');
  if TravailN2 <> '' then StWHere := StWhere + RecupClauseSql(TravailN2, 'PBC_TRAVAILN2');
  if TravailN3 <> '' then StWHere := StWhere + RecupClauseSql(TravailN3, 'PBC_TRAVAILN3');
  if TravailN4 <> '' then StWHere := StWhere + RecupClauseSql(TravailN4, 'PBC_TRAVAILN4');
  if TabLibre1 <> '' then StWHere := StWhere + RecupClauseSql(TabLibre1, 'PBC_LIBREPCMB1');
  if TabLibre2 <> '' then StWHere := StWhere + RecupClauseSql(TabLibre2, 'PBC_LIBREPCMB2');
  if TabLibre3 <> '' then StWHere := StWhere + RecupClauseSql(TabLibre3, 'PBC_LIBREPCMB3');
  if TabLibre4 <> '' then StWHere := StWhere + RecupClauseSql(TabLibre4, 'PBC_LIBREPCMB4');

  StInd := RecupClauseSql(Ind, 'PBI_INDICATEURBS');
  if ExisteSql('SELECT * FROM PBSINDICATEURS ' +
    'WHERE (PBI_TYPINDICATBS="032" OR PBI_TYPINDICATBS="033") ' + StInd) then
    St := 'SELECT SUM(PBC_VALCAT*PBC_NBELEMENT)/SUM(PBC_NBELEMENT) VALEUR FROM BILANSOCIAL ' +
      'WHERE PBC_DATEDEBUT="' + USDateTime(StrToDate(DateDeb)) + '" ' +
      'AND PBC_DATEFIN="' + USDateTime(StrToDate(DateFin)) + '" ' + StWhere
  else
    St := 'SELECT SUM(PBC_VALCAT) VALEUR FROM BILANSOCIAL ' +
      'WHERE PBC_DATEDEBUT="' + USDateTime(StrToDate(DateDeb)) + '" ' +
      'AND PBC_DATEFIN="' + USDateTime(StrToDate(DateFin)) + '" ' + StWhere;
      //Début PT25
  mdossiers := GetBasesMS(RegroupementSociete);
  if mdossiers = '' then begin
    //Non regroupement : comme avant
    Q := OpenSql(st, true);
    if not Q.eof then result := Q.FindField('VALEUR').AsFloat
    else result := 0;
    Ferme(Q);
    end
  else begin
    //Regroupement: récupération selection éventuelle de bases
    resmul := 0;
    result := 0;
    if listebases <> '' then mdossiers := listebases;
    while (mdossiers <> '') do
    begin
      lstdossier :=  ReadTokenSt(mdossiers);
      resdos := OpenSelect(st,lstdossier);
      If (not resdos.EOF) and (resdos <> nil) then resmul := resdos.FindField('VALEUR').AsFloat else resmul := 0;
      result := resmul + result;
      if assigned(resdos) then ferme(resdos);
      end;
    end;
  end;

function TOlePaiePGI.RecupClauseSql(StVal, StChamp: string): string;
begin
  Result := '';
  if StVal = '' then exit;
  Result := ' AND (';
  while StVal <> '' do
  begin
    Result := Result + ' ' + StChamp + '="' + ReadTokenSt(StVal) + '" OR ';
  end;
  Result := Copy(Result, 1, Length(Result) - 4) + ') ';
end;

function PGControlConcept: Boolean;
begin
  Result := ExJaiLeDroitConcept(ccOLE, FALSE);
  if (Result = FALSE)  then
  begin
    PgiInfo('Vous n''avez pas les droits d''utilisation du lien OLE', 'Lien OLE');
  end
  else
  begin
    if MonHabilitation.active then
    begin
      Result := FALSE ;
      PgiInfo('Votre habilitation vous interdit l''utilisation des liens OLE', 'Lien OLE');
    end;
  end;
end;

procedure PaieVideAnc ();
begin
    if TOB_AncienneteBS <> NIL then FreeAndNil (TOB_AncienneteBS); // PT23
end;

function TolePaiePgi.PGMultiSoc(): Variant;
begin
  Regroupsoc := AglLanceFiche('PAY','PGREGROUPSOC','','','');
end;

//PT25
function TOlePaiePGI.Getregroupsoc: String;
begin
  result := Regroupsoc;
end;

procedure TOlePaiePGI.setRegroupsoc(const Value: String);
begin
  Regroupsoc := Value;
end;

function TOlePaiePGI.GetNomBase: String;
begin
  result := NomBase;
end;

procedure TOlePaiePGI.setNomBase(const Value: String);
begin
  NomBase := Value;
end;
//Fin PT25

initialization //PT6 RegisterClass en phase d'initialization
  Automation.RegisterClass(AutoClassInfo);
end.
