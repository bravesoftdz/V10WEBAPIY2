{ Unité : Fonction de calcul des soldes de la Trésorerie
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |  Commentaires
--------------------------------------------------------------------------------------
 6.00.001.001  28/06/04   JP   Création de l'unité : regroupement de toutes les fonctions de
                               calcul des solde qui étaient avant dans d'autres unités en
                               particulier dans Commun.Pas
 6.51.001.003  21/11/05   JP   FQ 10304 : Inversion des champs TE_CLEVALEUR, TE_CLEOPERATION dans
                               la fonction de calcul simplifié des soldes RecalculSoldeSQL
 6.53.001.003  31/01/06   JP   FQ 10335 : refonte de GetSoldePointe
 7.00.001.007  17/05/06   JP   FQ 10339 : correction de RecalculSoldeSQL pour tenir compte des écritures
                               de réinitialisation et ne pas écraser leurs soldes ; problème posé lorsque
                               l'on synchronise une écriture du 01/01/XX (on allait chercher le solde de réinit.
                               de XX + 1 et non de XX !!)
7.00.001.010   05/06/06   JP   FQ 10339 : Reste les cas dont la date d'opé >= 02/01 et date de valeur <= 01/01
7.09.001.001   17/08/06   JP   Gestion du multi sociétés avec la création de GetSoldeNatureCompte
7.09.001.001   17/11/06   JP   Fonction de recalcul des écritures de réinitialisation lors du changement des
                               dates de valeur autour du début d'un millésime : RecalculInit
7.09.001.001   20/11/06   JP   RecalculSoldeValeur : Lorsque l'on déplace les dates de valeurs, il est inutile de faire
                               un recalcul depuis la date d'opération la plus vieille et de recalculer les solde en opération
7.06.001.001   26/12/06   JP   Adaptation des requêtes de RecalculSolde à la nouvelle clef primaire de TRECRITURE
8.01.001.010   28/03/07   JP   Nouvelle fonction de calcul de solde anticipant sur la suppression des soldes initiaux
                               au premier janvier : la date sera libre
8.10.001.001   09/07/07   JP   Nouvelle gestion des soldes : on ne reprend plus les soldes depuis TRECITURE, mais
                               on les calcule par requête. Chercher IsNewSoldes.
--------------------------------------------------------------------------------------}
unit UProcSolde;

interface

uses
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, Ent1, Hent1, UTOB, Classes, Constantes, HCtrls;

{28/06/04 : Recalcul tous les soldes à partir d'une date voire d'un exercice}
procedure CalculerTousLesSoles(Depart : TDateTime; Exercice : string = '');
{03/12/03 : Récupère le solde du 01/01 saisi auquel on retranche les opération du jours (MoinsOpe01) ou non }
function  GetSoldeMillesime   (CompteGeneral, Endate, Nature : string;
                               EnValeur, AuMatin : Boolean; EnPivot : Boolean = False) : Double;
{05/12/03 : Supprime la ligne d'initialisation de solde pour le compte et l'année}
procedure SupprimerInit       (CompteGeneral : string; Endate : TDateTime; SupprimerOk : Boolean);
{08/12/2003 : Préparation de la TList de mise à jour des soldes de la compta et appel de ExecMajSoldeTOB}
procedure PrepareMajSoldes    (TOBEcr : TOB; Plus : Boolean);
{Retourne la liste des années qui correspondent aux exercices ouverts}
function  GetListeMillesime   : HTStringList;
{Récupère le solde précédent en opération}
function  GetSoldePrec        (CompteGeneral, Clef : string) : Double ;
{Récupère le solde précédent en valeur}
function  GetSoldePrecValeur  (CompteGeneral, Clef : string) : Double ;
{Création d'une écriture de réinitialisation des soldes}
function  CreerEcritureInit   (DateDepart : TDateTime; Devise, General : string; MtVal, MtOpe : Double) : Boolean;
{13/10/04 : Récupére le solde des écriures pointées/non pointées pour le suivi bancaire}
function  GetSoldePointe      (CompteGeneral, Endate, Nature : string; EnValeur : Boolean) : Double;

function  GetSolde	      (CompteGeneral : string ; Endate : string; var DateCalc : string) : Double;
function  GetSoldeValeur      (CompteGeneral : string ; Endate : string; var DateCalc : string) : Double;
{27/10/04 : Appels mutliples au recalcul des soldes : les comptes et les dates de départ sont stockés dans la StringList}
procedure MultiRecalculSolde  (var ATraiter : TStringList);
{JP 18/09/03 : Calcul du solde à la date de "EnDate", en Valeur ou non, Ecritures réalisées ou non}
function  GetSoldeInit        (CompteGeneral, Endate, Nature : string; EnValeur : Boolean; EnPivot : Boolean = False) : Double;
{09/11/06 : Pour les fiches de suivi par flux (Détail du suivi + buget * 2), il est possible de ne séletionner
 aucun compte bancaire mais de filtrer sur un dossier}
function  GetSoldeInitDossier (NoDossier, Endate, Nature : string; EnValeur : Boolean;
                               EnPivot : Boolean = False; Banque : string = ''; Agence : string = '') : Double;
{JP 22/01/04 : récupère le solde cumulé des comptes d'une banque, ou d'une agence}
function  GetSoldeBanque      (Banque, Endate, Nature : string; EnValeur : Boolean; AgenceOk : Boolean = False; NoDossier : string = '') : Double;
{20/11/06 : Recalcul des soldes en valeurs uniquement}
function  RecalculSoldeValeur (Compte, DateDepart : string) : Double;
(*
{23/11/06 : Pour limiter le chargement de tob trop important, le recalcul va se faire par mois :
RecalculSolde boucle de DateDepart jusqu'à la fin, mois par mois et appelle RecalculSoldeDet pour chaque mois}
procedure RecalculSolde       (Compte, DateDepart : string; SoldeInit : Double; SoldeInitOk : Boolean = False; SoldeVal : Double = 0.001);
function  RecalculSoldeDet    (Compte : string; DateDe, DateA : TDateTime; SoldeInit : Boolean) : Boolean;
*)
function  RecalculSolde       (Compte, DateDepart : string; SoldeInit : Double; RecupSoldePrec : Boolean; SoldeVal : Double = 0.001) : Double;
{17/08/07 : Recalcul des soldes d'initialisation}
function  RecalculSoldeInit   (Compte : string; Millesime : Integer) : TOB;
{22/08/07 : Récupère la liste des soldes d'initialisation}
function  GetListeSoldesInit  (Compte : HString; DateDepart : TDateTime) : TOB;

{22/07/05 : Mise à jour des soldes par requête}
procedure RecalculSoldeSQL    (Cpte, CleVal, CleOpe : string; Dt : TDateTime; Mnt : Double; Etat : TEtatEcriture);
{05/01/05 : Mémorisation des comptes dont on doit recalculer les soldes}
procedure AddGestionSoldes    (var lSolde : TStringList; Cpte : string; Dt : TDateTime);
{28/12/05 : Calcul du solde sur un liste de comptes contenue dans Q}
function  GetSoldeSurListe    (Q : TQuery; Endate, Nature : string; EnValeur : Boolean) : Double;
{28/12/05 : Calcul d'un solde multi-comptes : CptTokenSt liste des comptes séparés par un ";"}
function  GetSoldeMultiComptes(CptTokenSt, Endate, Nature : string; EnValeur : Boolean) : Double;
{17/08/06 : Retourne le solde en DEVISE PIVOT pour une nature de compte et un dossier}
function  GetSoldeNatureCompte(Dossier, NatCpt, Endate, Nature, Banque : string; EnValeur, Millesime : Boolean) : Double;
{17/11/06 : recalcul du solde de réinitialisation lorsque l'on tourne autour du 01/01/XX : Le Montant doit être en devise}
function  RecalculInit(dtVal : TDateTime; Compte : string; MntDev : Double; Recalcul : Boolean = False) : Boolean;
{17/11/06 : vérifie si OldDt, NewDt sont sur le même millésime : sinon, appel de RecalculInit}
function  GereSoldeInit(OldDt, NewDt : TDateTime; Compte : string; MntDev : Double; Recalcul : Boolean = False) : Boolean;
{20/11/06 : OldClef, NewClef sont sur le même millésime : sinon, appel de RecalculInit}
function  GereSoldeInitClef(OldClef, NewClef : string; Compte : string; MntDev : Double; Recalcul : Boolean = False) : Boolean;
{28/03/07 : Nouvelle fonction de calcul de solde}
function  GetSoldeBancaire(Compte, Nature : string; DateRef : TDateTime; EnValeur : Boolean = True; EnPivot : Boolean = False) : Double;
{09/10/07 : Récupération du solde bancaire à partir des relevés bancaire}
function  GetSoldeRelevesBQE(Compte : string; DateVal : TDateTime; Dossier : string = '') : Double;

implementation

uses
  {$IFDEF TRCONF}
  UlibConfidentialite,
  {$ENDIF TRCONF}
  commun, UtilSais, UProcGen, UProcEcriture, SaisUtil, HStatus, ParamSoc, HMsgBox, Math,
  UtilPgi;

{---------------------------------------------------------------------------------------}
procedure CalculerTousLesSoles(Depart : TDateTime; Exercice : string = '');
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  D : string;
begin
  if IsNewSoldes then Exit;

  if Exercice = '' then D := DateToStr(Depart)
                   else D := DateToStr(EncodeDate(StrToInt(Exercice), 1, 1));
  Q := OpenSQL('SELECT BQ_GENERAUX FROM BANQUECP', True);
  try
    while not Q.EOF do begin
      RecalculSolde(Q.FindField('BQ_GENERAUX').AsString, D, 0, True);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;

{03/12/2003 : Récupère le solde du 01/01 saisi auquel on retranche les opération du jours
              ou non (MoinsOpe01), en devise pivot ou non
 07/11/06 : FQ 10256 : les soldes d'initialisation ayant été déplacés au matin, je remplace le paramètre
            MoinsOpe01 (avant solde au soir - opération su jour) par AuMatin (solde initial).
            Ce choix sémantique est fait de manière à ne pas changer la valeur du Boolean dans les
            différents algorithmes : if MoinsOpe01 équivaut au niveau fonctionnel à if AuMatin.
            Par contre dans la fonction, on fait l'inverse :
            - avant : if MoinsOpe01, on récupérait l'écriture d'init et on ôtait les opérations du 01/01
            - maintenant : if not AuMatin : on récupère l'écriture d'init et on ajoute les opérations du 01/01
{---------------------------------------------------------------------------------------}
function GetSoldeMillesime(CompteGeneral, Endate, Nature : string;
                           EnValeur, AuMatin : Boolean; EnPivot : Boolean = False) : Double;
{---------------------------------------------------------------------------------------}
var
  SQL 	: string ;
  Q	: TQuery ;
  Solde	: Double ;
  Ok    : Boolean;
begin
  Solde := 0;
  {A true si on trouvé un solde forcé}
  Ok := False;

  {On commence par récupérer le solde au premier janvier}
  if EnValeur then
    SQL := 'SELECT TE_SOLDEDEVVALEUR, TE_DEVISE FROM TRECRITURE WHERE TE_DATEVALEUR = "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" '
  else
    SQL := 'SELECT TE_SOLDEDEV, TE_DEVISE FROM TRECRITURE WHERE TE_DATECOMPTABLE = "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" ';

  {Si on travail sur un compte}
  if CompteGeneral <> '' then begin
    SQL := SQL + 'AND TE_GENERAL = "' + CompteGeneral + '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"';

    Q := OpenSQL(SQL, True);
    if not Q.EOF then begin
      {Conversion du montant dans la devise Pivot}
      if (Q.Fields[1].AsString = V_PGI.DevisePivot) or not EnPivot then
        Solde := Solde + Q.Fields[0].AsFloat
      else
        Solde := Solde + (Q.Fields[0].AsFloat * RetPariteEuro(Q.Fields[1].AsString, DebutAnnee(StrToDate(EnDate)){V_PGI.DateEntree}));
      Ok := True;
    end;
    Ferme(Q);
  end
  {Si demande d'un solde général à tous les comptes}
  else begin
    SQL := SQL + ' AND TE_QUALIFORIGINE = "' + CODEREINIT + '"';
    Q := OpenSQL(SQL, True);
    {On cumule les soldes initiaux de tous les comptes}
    while not Q.EOF do begin
      Ok := True;
      {Conversion du montant dans la devise Pivot}
      if Q.Fields[1].AsString = V_PGI.DevisePivot then
        Solde := Solde + Q.Fields[0].AsFloat
      else
        Solde := Solde + (Q.Fields[0].AsFloat * RetPariteEuro(Q.Fields[1].AsString, DebutAnnee(StrToDate(EnDate)){V_PGI.DateEntree}));
      Q.Next;
    end;
    Ferme(Q);
  end;

  {On retranche éventuellement les opérations du 01/01 pour avoir le solde au matin et non au soir, si
   bien sûr on a trouvé un solde forcé (Ok)
   07/11/06 : FQ 10256 : on ajoute les opérations du jour}
  if not AuMatin and Ok then begin
    if (CompteGeneral <> '') or not EnPivot then begin
      SQL := 'SELECT SUM(TE_MONTANTDEV) FROM TRECRITURE WHERE ';
      if (CompteGeneral <> '') then SQL := SQL + 'TE_GENERAL = "' + CompteGeneral + '" AND '
    end
    {... Sinon on récupère le solde dans la devise pivot}
    else
      SQL := 'SELECT SUM(TE_MONTANT) FROM TRECRITURE WHERE ';

    if EnValeur then
      SQL := SQL + 'TE_DATEVALEUR = "' + UsDateTime(StrToDate(Endate)) + '" '
    else
      SQL := SQL + 'TE_DATECOMPTABLE = "' +UsDateTime(StrToDate(Endate)) + '" ';
    {On ne retire que les écriture qui figurent à l'affichage}
    SQL := SQL + Nature;

    Q := OpenSQL(SQL, True);
    if not Q.EOF then Solde := Solde + Q.Fields[0].AsFloat;
    Ferme(Q);
  end;
  if Ok then Result := Solde
        else Result := 0.001;
end;

{05/12/03 : Supprime la ligne d'initialisation de solde pour le compte et l'année
{---------------------------------------------------------------------------------------}
procedure SupprimerInit(CompteGeneral : string; Endate : TDateTime; SupprimerOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  if not SupprimerOk then Exit;
  SQL := 'DELETE FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" AND TE_DATEVALEUR = "' +
         UsDateTime(EnDate) + '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"';
  ExecuteSQL(SQL);
end;

{---------------------------------------------------------------------------------------}
procedure PrepareMajSoldes(TOBEcr : TOB; Plus : Boolean);
{---------------------------------------------------------------------------------------}
var
 i             : Integer;
 lTS           : TList;
// lTA           : TList;
 lTOBLigneEcr  : TOB;
begin
 {L'idéal serait d'utiliser UtilSais.MajSoldesEcritureTOB en constituant une tob qui contiendrait
  toutes les lignes à OK}
 lTS := TList.Create;
// lTA := TList.Create; Gestion des ANO dynamiques
 try
   for i := 0 to TOBEcr.Detail.Count - 1 do begin
     lTOBLigneEcr := TOBEcr.Detail[i];
     {On ne traite que les écritures qui ont été intégrées}
     if (lTOBLigneEcr.GetValue('RESULTAT') = 'OK') and (lTOBLigneEcr.GetValue('E_GENERAL')<> '')then
       AjouteTOB(lTS,  lTOBLigneEcr , Plus);
       //if GetParamSocSecur('SO_CPANODYNA' then AjouteTobAno(lTA, lTOBLigneEcr, NatureGene
   end;

   {Ecriture dans la base}
   ExecMajSoldeTOB(lTS);
   //for n := 0 to lTA.Count - 1 do
     //ExecReqMajANo
  finally
    DisposeListe(lTS, True);
//    DisposeListe(lTA, True);
  end;
end;

{Retourne la liste des années qui correspondent aux exercices ouverts
{---------------------------------------------------------------------------------------}
function GetListeMillesime : HTStringList;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
  n : Integer;
begin
  Result := HTStringList.Create;
  Result.Duplicates := dupIgnore;
  Result.Sorted := True;
  DecodeDate(GetEnCours.Deb, a, m, j);
  Result.Add(IntToStr(a));
  for n := 1 to 4 do begin
    Result.Add(IntToStr(a + n));
    Result.Add(IntToStr(a - n));
  end;
end;

{---------------------------------------------------------------------------------------}
function GetSoldePrec(CompteGeneral, Clef : string) : Double ;
{---------------------------------------------------------------------------------------}
var
  SQL 	: string ;
  Q	: TQuery ;
  Solde	: Double ;
begin
  Solde  := 0;
  Result := 0;

  if IsNewSoldes then Exit; {N'est utilisé que dans la fonction RecalculSolde}

  {On récupère le dernier solde en devise et sa date comptable pour un compte donné à une date donnée}
  SQL := 'SELECT TE_SOLDEDEV FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" ';
  SQL := SQL + ' AND TE_CLEOPERATION < "' + Clef + '" ORDER BY TE_CLEOPERATION DESC';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Solde    := Q.FindField('TE_SOLDEDEV').asFloat ;
  Ferme(Q);

  Result := Solde;
end ;

{JP 29/08/03 : Récupère le dernier solde en valeur avant une date donnée
{---------------------------------------------------------------------------------------}
function GetSoldePrecValeur(CompteGeneral, Clef : string) : Double ;
{---------------------------------------------------------------------------------------}
var
  SQL 	: string ;
  Q	: TQuery ;
  Solde	: Double ;
begin
  Solde  := 0;
  Result := 0;

  if IsNewSoldes then Exit; {N'est utilisé que dans la fonction RecalculSolde}

  {On récupère le dernier solde en devise et sa date comptable pour un compte donné à une date donnée}
  SQL := 'SELECT TE_SOLDEDEVVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" ';
  SQL := SQL + ' AND TE_CLEVALEUR < "' + Clef + '" ORDER BY TE_CLEVALEUR DESC';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Solde    := Q.Fields[0].AsFloat;
  Ferme(Q);

  Result := Solde;
end ;

{Création d'une écriture de réinitialisation des soldes
{---------------------------------------------------------------------------------------}
function CreerEcritureInit(DateDepart : TDateTime; Devise, General : string; MtVal, MtOpe : Double) : Boolean;
{---------------------------------------------------------------------------------------}
var
  TEc : TOB;
  QQ  : TQuery;
  Par : Double;
  Unq : Integer;
  a, m, j : Word;
begin
  Result := True;
  {On regarde s'il n'y pas déjà une écriture de réinitialisation pour le compte sur l'année}
  if ExisteSQL('SELECT TE_CLEVALEUR FROM TRECRITURE WHERE TE_DATEVALEUR = "' + UsDateTime(DateDepart) + '" AND ' +
               'TE_GENERAL = "' + General + '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"') then begin
    {Si c'est le cas, on se contente de mettre à jour les soldes}
    ExecuteSQL('UPDATE TRECRITURE SET TE_SOLDEDEV = ' + StrFPoint(MtOpe) +
               ', TE_SOLDEDEVVALEUR = ' + StrFPoint(MtVal) +
               ' WHERE TE_DATEVALEUR = "' + UsDateTime(DateDepart) + '" AND ' +
               'TE_GENERAL = "' + General + '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"');
    Exit;
  end;

  {Sinon création d'une nouvelle écriture}
  TEc := TOB.Create('TRECRITURE', nil, -1);
  try
    if IsTresoMultiSoc then begin
      QQ := OpenSQL('SELECT BQ_SOCIETE, BQ_NODOSSIER FROM BANQUECP WHERE BQ_CODE = "' + General + '"', True);
      if not QQ.EOF then begin
        TEc.SetString('TE_NODOSSIER', QQ.FindField('BQ_NODOSSIER').AsString);
        TEc.SetString('TE_SOCIETE', QQ.FindField('BQ_SOCIETE').AsString);
      end;
      Ferme(QQ);
      InitNlleEcritureTob(TEc, General, Tec.GetString('TE_NODOSSIER'));
    end
    else
      InitNlleEcritureTob(TEc, General);


    {Génération des clés : à l'inverse des clefs classiques, on part du maximum pour que l'écriture soit toujours
                           la dernière du jour. CODEUNIQUEDEC n'étant utilisé que dans cette routine, on devrait
                           rester proche de 9999999
    FQ 10256 : 07/11/06 Déplacement des soldes de réinitialisation au matin clef := AA0000Unq
                           }
    DecodeDate(DateDepart, a, m, j);
    Unq := StrToInt(Commun.GetNum(CODEUNIQUEDEC, CODEUNIQUEDEC, CODEUNIQUEDEC));
// //   TEc.SetString('TE_CLEVALEUR'   , RetourneCleEcriture(DateDepart, 0000000));
//    TEc.SetString('TE_CLEOPERATION', RetourneCleEcriture(DateDepart, 0000000));
    TEc.SetString('TE_CLEVALEUR'   , Copy(IntToStr(a), 3, 2) + PadL(IntToStr(Unq), '0', 11));
    TEc.SetString('TE_CLEOPERATION', Copy(IntToStr(a), 3, 2) + PadL(IntToStr(Unq), '0', 11));
    Commun.SetNum(CODEUNIQUEDEC, CODEUNIQUEDEC, CODEUNIQUEDEC, IntToStr(Unq));

    TEc.SetDateTime('TE_DATECOMPTABLE', DateDepart);
    TEc.SetDateTime('TE_DATEVALEUR',    DateDepart);
    TEc.SetDouble('TE_SOLDEDEVVALEUR',  MtVal);
    TEc.SetDouble('TE_SOLDEDEV',        MtOpe);

    TEc.SetString('TE_DEVISE', Devise);
    Par := RetPariteEuro(Devise, DateDepart);
    if Par = 0 then TEc.SetDouble('TE_COTATION', 1)
               else TEc.SetDouble('TE_COTATION', 1 / Par);

    TEc.SetDouble('TE_MONTANTDEV', 0);
    TEc.SetDouble('TE_MONTANT',    0);

    TEc.SetString('TE_CODECIB',       '00');
    TEc.SetString('TE_CODEFLUX',      CODEREGULARIS);
    TEc.SetString('TE_LIBELLE',       'Initialisation');
    TEc.SetString('TE_REFINTERNE',    'Init. ' + V_PGI.User + ' - ' + DateToStr(V_PGI.DateEntree));
    TEc.SetString('TE_QUALIFORIGINE', CODEREINIT);
    TEc.SetString('TE_USERCREATION',   V_PGI.User);
    TEc.SetString('TE_USERVALID',      V_PGI.User);
    TEc.SetString('TE_USERCOMPTABLE',  'INI');
    TEc.SetString('TE_USERMODIF',      V_PGI.User);
    TEc.SetString('TE_CONTREPARTIETR', '------');
    TEc.SetDateTime('TE_DATECREATION', V_PGI.DateEntree);
    TEc.SetDateTime('TE_DATEMODIF',    V_PGI.DateEntree);
    TEc.SetDateTime('TE_DATEVALID',    V_PGI.DateEntree);
    TEc.SetDateTime('TE_DATERAPPRO',   DateDepart); {31/10/06 : c'est mieux que V_PGI.DateEntree pour le suivi bancaire}
    {Récupération des infos bancaires}
    QQ := OpenSQL('SELECT PQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN FROM BANQUECP ' +
                  'LEFT JOIN BANQUES ON PQ_BANQUE = BQ_BANQUE WHERE BQ_CODE = "' + General + '"', True);
    try
      if not QQ.EOF then begin
        TEc.SetString('TE_CODEBANQUE',  QQ.Fields[0].AsString);
        TEc.SetString('TE_CODEGUICHET', QQ.Fields[1].AsString);
        TEc.SetString('TE_NUMCOMPTE',   QQ.Fields[2].AsString);
        TEc.SetString('TE_CLERIB',      QQ.Fields[3].AsString);
        TEc.SetString('TE_IBAN',        QQ.Fields[4].AsString);
      end;
    finally
      ferme(QQ);
    end;

    TEc.SetString('TE_EXERCICE', TObjetExercice.GetCurrent.GetExoNodos(TEc.GetDateTime('TE_DATECOMPTABLE'),
                                                                       TEc.GetString('TE_NODOSSIER')));

    {Crée une valeur unique (varchar(17))}
    TEc.SetString('TE_NUMTRANSAC', 'INI' + Padr(Copy(General, 1, 10), '0', 10)  + Copy(TEc.GetString('TE_CLEVALEUR'), 1, 4));

    TEc.SetString('TE_NATURE', na_Realise);
    TEc.InsertDb(nil);
  finally
    if Assigned(TEc) then FreeAndNil(TEc);
  end;
end;

{13/10/04 : Récupére le solde des écritures pointées pour le suivi bancaire
 31/01/06 : FQ 10335 : Refonte de la fonction pour gérer un certain nombre de cas particulier
{---------------------------------------------------------------------------------------}
function GetSoldePointe(CompteGeneral, Endate, Nature : string; EnValeur : Boolean) : Double;
{---------------------------------------------------------------------------------------}
var
  SQL 	: string ;
  SQLR	: string ;
  Q	: TQuery ;
  Solde	: Double ;
  Ok    : Boolean;
  Regul : Double ;
  dtMil : string;
begin
  {Ok est à true si on veut le solde au "matin" du 01/01}
  Ok := (Endate = DateToStr(DebutAnnee(StrToDate(Endate))));

  {08/11/06 : FQ 10256 : Récupération du solde du millésime au matin <=> à l'écriture de réinitialisation}
  Solde := GetSoldeMillesime(CompteGeneral, EnDate, Nature, EnValeur, True, True);

  {On récupère le solde du compte dans la devise pivot ...}
  if (CompteGeneral <> '')  then begin
    SQL  := 'SELECT SUM(TE_MONTANT) FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" AND ';
    if Ok then
      SQLR := 'SELECT SUM(TE_MONTANT), TE_DATEVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" AND '
    else
      SQLR := 'SELECT SUM(TE_MONTANT), TE_DATEVALEUR, TE_DATERAPPRO FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" AND ';
  end
  {... Sinon on récupère le solde dans la devise pivot}
  else begin
    SQL  := 'SELECT SUM(TE_MONTANT) FROM TRECRITURE WHERE ';
    if Ok then
      SQLR := 'SELECT SUM(TE_MONTANT), TE_DATEVALEUR FROM TRECRITURE WHERE '
    else
      SQLR := 'SELECT SUM(TE_MONTANT), TE_DATEVALEUR, TE_DATERAPPRO FROM TRECRITURE WHERE ';
  end;

  if Trim(Nature) <> '' then begin
    SQL  := SQL  + ' ' + Nature + ' AND ';
    SQLR := SQLR + ' ' + Nature + ' AND ';
  end;

  dtMil := UsDateTime(DebutAnnee(StrToDate(Endate)));
  {31/01/06 : FQ 10335 : La fiche de suivi bancaire fonctionne sur deux requêtes, l'une sur les dates de
              valeur, l'autre sur les date de rapprochement mais dans le calcul du solde initial en valeur
              et du solde pointé (ici-même), on part du solde en valeur qui est calculé à partir du 01/01
              au matin : il faut donc créer une régul sur le solde pointé pour tenir compte des
              écritures pointées avant la date de traitement ou dont la date de rappro est postérieure à
              celle du millésime.}
  SQLR := SQLR + '((TE_DATERAPPRO < "' + dtMil + '" AND TE_DATERAPPRO > "' + UsDateTime(iDate1900) + '" AND TE_DATEVALEUR >= "' + dtMil + '") OR';
  SQLR := SQLR + ' (TE_DATEVALEUR < "' + dtMil + '" AND TE_DATERAPPRO = "' + UsDateTime(iDate1900) + '") OR';
  SQLR := SQLR + ' (TE_DATEVALEUR < "' + dtMil + '" AND TE_DATERAPPRO >= "' + dtMil + '")) ';
  if Ok then
    SQLR := SQLR + ' GROUP BY TE_DATEVALEUR ORDER BY TE_DATEVALEUR'
  else
    SQLR := SQLR + ' GROUP BY TE_DATEVALEUR, TE_DATERAPPRO ORDER BY TE_DATEVALEUR';

  Q := OpenSQL(SQLR, True);
  try
    Regul := 0;
    while not Q.EOF do begin
      {On déduit de la régulation :
       1/ les flux dont la date de valeur est antérieures à la date de début
       2/ les flux, si la date de début n'est pas un millésime, dont la date de valeur est >= au début et
          dont la date de rappro est >= au début de millésime mais antérieure au début}
      if (Q.FindField('TE_DATEVALEUR').AsDateTime < DebutAnnee(StrToDate(Endate))) or
         (not Ok and (Q.FindField('TE_DATEVALEUR').AsDateTime >= StrToDate(Endate)) and
          (Q.FindField('TE_DATERAPPRO').AsDateTime >= DebutAnnee(StrToDate(Endate))) and
          (Q.FindField('TE_DATERAPPRO').AsDateTime < StrToDate(Endate))) then
        Regul := Regul - Q.Fields[0].AsFloat
      {Et on ajoute à la régulation les flux dont la date de rappro est antérieure au millésime, mais
       la date de valeur est supérieure au millésime}
      else
        Regul := Regul + Q.Fields[0].AsFloat;

      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

  {Si on demande le solde au 01/01/ ou 02/01, il est inutile de lancer la requête qui renverra vide.
   FQ 10335 : on se contente de gérer la régul
   07/11/06 : FQ 10256 : on limite le Exit au 01/01. Pour le 02/01, il faut ajouter les opérations du 01/01}
  if Ok {or (Endate = DateToStr(DebutAnnee(StrToDate(Endate)) + 1))} then begin
    Result := Solde + Regul;
    Exit;
  end;

  {05/04/05 : FQ 10240 : Pour qu'une écriture soit rapprochée au "EnDate", il ne suffit pas que la date de rappro
              soit différente de 01/01/1900, il faut qu'elle soit inférieure à "EnDate".
              Par ailleurs, on ne travaille plus que sur les dates de rapprochement
   31/01/05 : FQ 10335 : Il ne faut traiter que les écritures pointées de puis le début du millésime
   07/11/06 : FQ 10256 : supériorité non stricte par rapport à dtMil, car il faut ajouter les opérations du 01/01.}
  SQL := SQL + ' (TE_DATERAPPRO >= "' + dtMil + '" AND TE_DATERAPPRO < "' + UsDateTime(StrToDate(Endate)) + '")  ';

  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Solde := Solde + Q.Fields[0].AsFloat + Regul;

  Ferme(Q);
  Result := Solde;
end;

{JP 22/01/04 : récupère le solde cumulé des comptes d'une banque, ou d'une agence
{---------------------------------------------------------------------------------------}
function GetSoldeBanque(Banque, Endate, Nature : string; EnValeur : Boolean; AgenceOk : Boolean = False; NoDossier : string = '') : Double;
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
  SQL : string;
  ch  : string;
  whC : string;
begin
  if AgenceOk then ch := 'BQ_AGENCE'
              else ch := 'BQ_BANQUE';
  SQL := 'SELECT BQ_CODE FROM BANQUECP WHERE ' + ch + ' = "' + Banque + '"';
  {$IFDEF TRCONF}
  whC := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if whC <> '' then whC := ' AND (' + whC + ') ';
  {$ENDIF TRCONF}
  SQL := SQL + whC;

  if (NoDossier <> '') and (Pos('<<', NoDossier) = 0) then begin
    if NoDossier[Length(NoDossier)] <> ';' then NoDossier := NoDossier + ';';
    SQL := SQL + ' AND BQ_NODOSSIER IN (' + GetClauseIn(NoDossier) + ')';
  end;

  Q := OpenSQL(SQL, True);
  try
    Result := GetSoldeSurListe(Q, Endate, Nature, EnValeur);
  finally
    Ferme(Q);
  end;
end;

{28/12/05 : Calcul d'un solde multi-comptes : CptTokenSt liste des comptes séparés par un ";"
{---------------------------------------------------------------------------------------}
function GetSoldeMultiComptes(CptTokenSt, Endate, Nature : string; EnValeur : Boolean) : Double;
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
  SQL : string;
  ch  : string;
  whC : string;
begin
  Result := 0;
  if CptTokenSt = '' then
    Exit
  else if CptTokenSt[Length(CptTokenSt)] <> ';' then
    CptTokenSt := CptTokenSt + ';';

  ch := 'BQ_CODE IN (' + GetClauseIn(CptTokenSt) + ')';
  SQL := 'SELECT BQ_CODE FROM BANQUECP WHERE ' + ch;
  {$IFDEF TRCONF}
  whC := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if whC <> '' then whC := ' AND (' + whC + ') ';
  {$ENDIF TRCONF}
  SQL := SQL + whC;

  Q := OpenSQL(SQL, True);
  try
    Result := GetSoldeSurListe(Q, Endate, Nature, EnValeur);
  finally
    Ferme(Q);
  end;
end;

{28/12/05 : Calcul du solde sur un liste de comptes}
{---------------------------------------------------------------------------------------}
function GetSoldeSurListe(Q : TQuery; Endate, Nature : string; EnValeur : Boolean) : Double;
{---------------------------------------------------------------------------------------}
var
  Mnt : Double;
begin
  Result := 0;
  {Cumul sur la liste des comptes}
  while not Q.EOF do begin
    {On récupère le SoldeInit en devise pivot}
    Mnt := GetSoldeInit(Q.FindField('BQ_CODE').AsString, Endate, Nature, EnValeur, True);
    {Si le résultat = 0.001, sauf cas excéptionnel, c'est que l'on n'a réussi à récupèrer un solde de millésime}
    if StrFPoint(Mnt) <> '0.001' then
      Result := Result + Mnt;
    Q.Next;
  end;
end;

{09/11/06 : Pour les fiches de suivi par flux (buget * 2), il est possible de ne séletionner
 aucun compte bancaire mais de filtrer sur un dossier
{---------------------------------------------------------------------------------------}
function GetSoldeInitDossier(NoDossier, Endate, Nature : string; EnValeur : Boolean;
                             EnPivot : Boolean = False; Banque : string = ''; Agence : string = '') : Double;
{---------------------------------------------------------------------------------------}
var
  SQL   : string;
  Q	: TQuery ;
  Solde	: Double ;
  whC : string;
begin
  Result := 0;
  SQL := 'SELECT BQ_CODE FROM BANQUECP WHERE BQ_NODOSSIER = "' + NoDossier + '"';

  if (Banque <> '') and (Pos('<<', Banque) = 0) then begin
    if Banque[Length(Banque)] <> ';' then Banque := Banque + ';';
    SQL := SQL + ' AND BQ_BANQUE IN (' + GetClauseIn(Banque) + ')';
  end;

  if (Agence <> '') and (Pos('<<', Agence) = 0) then begin
    if Agence[Length(Agence)] <> ';' then Agence := Agence + ';';
    SQL := SQL + ' AND BQ_AGENCE IN (' + GetClauseIn(Agence) + ')';
  end;

  {$IFDEF TRCONF}
  whC := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if whC <> '' then whC := ' AND (' + whC + ') ';
  {$ENDIF TRCONF}
  SQL := SQL + whC;

  Q := OpenSQL(SQL, True);
  try
    while not Q.EOF do begin
      Solde := GetSoldeInit(Q.FindField('BQ_CODE').AsString, Endate, Nature, EnValeur, EnPivot);
      if StrFPoint(Solde) <> '0.001' then Result := Result + Solde;
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;

{JP 18/09/03 : Calcul du solde au "matin" de la date de "EnDate", en Valeur ou non, Ecritures réalisées ou non
{---------------------------------------------------------------------------------------}
function GetSoldeInit(CompteGeneral, Endate, Nature : string; EnValeur : Boolean; EnPivot : Boolean = False) : Double;
{---------------------------------------------------------------------------------------}
var
  SQL 	: string ;
  Q	: TQuery ;
  Solde	: Double ;
  Ok    : Boolean;
begin
  {Ok est à true si on veut le solde au "matin" du 01/01}
  Ok := (Endate = DateToStr(DebutAnnee(StrToDate(Endate))));

  Solde := GetSoldeMillesime(CompteGeneral, EnDate, Nature, EnValeur, True{Ok}, EnPivot);
  {Si on demande le solde au 01/01 ou 02/01, il est inutile de lancer la requête qui renverra vide
   07/11/06 : FQ 10256 : on limite le Exit au 01/01. Pour le 02/01, il faut ajouter les opérations du 01/01}
  if Ok {or (Endate = DateToStr(DebutAnnee(StrToDate(Endate)) + 1))} then begin
    Result := Solde;
    Exit;
  end;

  {Si on est sur un compte on récupère le solde dans la devise du compte ...}
  if (CompteGeneral <> '') and not EnPivot then
    SQL := 'SELECT SUM(TE_MONTANTDEV) FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" AND '
  {JP 27/02/04 : ... Sinon on récupère le solde du compte dans la devise pivot ...}
  else if (CompteGeneral <> '')  then
    SQL := 'SELECT SUM(TE_MONTANT) FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" AND '
  {... Sinon on récupère le solde dans la devise pivot}
  else
    SQL := 'SELECT SUM(TE_MONTANT) FROM TRECRITURE WHERE ';

  {On récupère le dernier solde en devise et sa date comptable pour un compte donné à une date donnée.
   Par ailleur, on ne travaille que sur l'année en cours, l'outil de calcul des soldes permettant à
   l'utilisateur de réinitialiser le compte au 01/01
   07/11/06 : FQ 10256 : Maintenant que l'écriture de Réinitialisation est au Matin, il faut que la date soit
              supérieure ou égale au 01/01 et non plus strictement supérieure}
  if EnValeur then
    SQL := SQL + 'TE_DATEVALEUR < "' + UsDateTime(StrToDate(Endate)) + '" AND TE_DATEVALEUR >= "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" '
  else
    SQL := SQL + 'TE_DATECOMPTABLE < "' +UsDateTime(StrToDate(Endate)) + '" AND TE_DATECOMPTABLE >= "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" ';

  if Trim(Nature) <> '' then
    SQL := SQL + ' ' + Nature;

  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Solde := Solde + Q.Fields[0].AsFloat;

  Ferme(Q);
  Result := Solde;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /
Description .. : retourne le dernier solde en opé d'un compte général
Mots clefs ... :
*****************************************************************}
function GetSolde(CompteGeneral : string ; Endate : string; var DateCalc : string) : Double ;
var
  SQL 	: string ;
  Q	: TQuery ;
  Solde	: Double ;
begin
  Solde := 0.001;

  {Si la date est le 31/12, c'est que l'on veut récupérer le solde au matin du 01/01}
  if (StrToDate(Endate) + 1) = DebutAnnee(StrToDate(Endate) + 1) then begin
    Solde := GetSoldeMillesime(CompteGeneral, DateToStr(StrToDate(Endate) + 1), '', False, True);
    DateCalc := EnDate;
  end
  {07/07/06 : FQ 10367 : gestion du solde au 2 janvier au matin}
  else if StrToDate(Endate) = DebutAnnee(StrToDate(Endate)) then begin
    Solde := GetSoldeMillesime(CompteGeneral, DateToStr(StrToDate(Endate)), '', False, False);
    DateCalc := EnDate;
  end;

  if IsNewSoldes then begin
    if StrFPoint(Solde) = '0.001' then
      {13/08/07 : FQ 10508 : il est mieux de mémoriser le solde !!
       19/08/08 : FQ 10557 : Il faut faire +1 pour bien partir de la date de début}
      Solde := GetSoldeInit(CompteGeneral, DateToStr(StrToDate(Endate) + 1), '', False);
    DateCalc := EnDate;
  end

  else begin
    if StrFPoint(Solde) = '0.001' then begin
      SQL := 'SELECT ##TOP 1## TE_DATECOMPTABLE, TE_SOLDEDEV FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" ' +
             ' AND TE_DATECOMPTABLE <= "' + UsDateTime(StrToDate(Endate)) + '" ';
      {07/07/06 : FQ 10367 : ajout du supérieur au début d'année, sinon, on peut se retrouver avec un très vieux soldes}
      SQL := SQL + 'AND TE_DATECOMPTABLE >= "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" ORDER BY TE_CLEOPERATION DESC';
      Q := OpenSQL(SQL, True);
      try
        if not Q.EOF then begin
          solde    := Q.FindField('TE_SOLDEDEV').asFloat ;
          dateCalc := Q.FindField('TE_DATECOMPTABLE').AsString;
        end
        {11/07/07 : Problème vu en 7.06, qui ne devrait plus se poser avec la nouvelle gestion des soldes,
                    mais si on passe par un paramsoc ... Si on n'a pas d'écriture entre le premier janvier
                    et la date de début, on reprend le solde initiale}
        else begin
          Solde := GetSoldeMillesime(CompteGeneral, DateToStr(DebutAnnee(StrToDate(Endate))), '', False, False);
          DateCalc := DateToStr(DebutAnnee(StrToDate(Endate)));
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;

  if StrFPoint(Solde) = '0.001' then Result := 0
                                else Result := Solde;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /
Description .. : retourne le solde en valeur d'un compte général
Mots clefs ... :
*****************************************************************}
function GetSoldeValeur(CompteGeneral : string ; Endate : string; var DateCalc : string) : Double ;
var
  SQL 	: string ;
  Q	: TQuery ;
  Solde	: Double ;
begin
  Solde := 0.001;
  {Si la date est le 31/12, c'est que l'on veut récupérer le solde au matin du 01/01}
  if (StrToDate(Endate) + 1) = DebutAnnee(StrToDate(Endate) + 1) then begin
    Solde := GetSoldeMillesime(CompteGeneral, DateToStr(StrToDate(Endate) + 1), '', True, True);
    DateCalc := EnDate;
  end
  {07/07/06 : FQ 10367 : gestion du solde au 2 janvier au matin}
  else if StrToDate(Endate) = DebutAnnee(StrToDate(Endate)) then begin
    Solde := GetSoldeMillesime(CompteGeneral, Endate, '', True, False);
    DateCalc := EnDate;
  end;

  if IsNewSoldes then begin
    if StrFPoint(Solde) = '0.001' then
      {13/08/07 : FQ 10508 : il est mieux de mémoriser le solde !!
       19/08/08 : Il est mieux de mettre True car le solde est en valeur
       19/08/08 : FQ 10557 : Il faut faire +1 pour bien partir de la date de début}
      Solde := GetSoldeInit(CompteGeneral, DateToStr(StrToDate(Endate) + 1), '', True);
    DateCalc := EnDate;
  end

  else begin
    if StrFPoint(Solde) = '0.001' then begin
      SQL := 'SELECT ##TOP 1## TE_DATEVALEUR, TE_SOLDEDEVVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + CompteGeneral + '" ' +
             ' AND TE_DATEVALEUR <= "' + UsDateTime(StrToDate(Endate)) + '" ';
      {07/07/06 : FQ 10367 : ajout du supérieur au début d'année, sinon, on peut se retrouver avec un très vieux soldes}
      SQL := SQL + 'AND TE_DATEVALEUR >= "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" ORDER BY TE_CLEVALEUR DESC';
      Q := OpenSQL(SQL, True);
      try
        if not Q.EOF then begin
          Solde    := Q.FindField('TE_SOLDEDEVVALEUR').asFloat ;
          DateCalc := Q.FindField('TE_DATEVALEUR').Asstring;
        end
        {11/07/07 : Problème vu en 7.06, qui ne devrait plus se poser avec la nouvelle gestion des soldes,
                    mais si on passe par un paramsoc ... Si on n'a pas d'écriture entre le premier janvier
                    et la date de début, on reprend le solde initiale}
        else begin
          Solde := GetSoldeMillesime(CompteGeneral, DateToStr(DebutAnnee(StrToDate(Endate))), '', True, False);
          DateCalc := DateToStr(DebutAnnee(StrToDate(Endate)));
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;

  if StrFPoint(Solde) = '0.001' then Result := 0
                                else Result := Solde;
end;

{17/08/06 : Retourne le solde en DEVISE PIVOT pour une nature de compte et un dossier
 07/11/06 : FQ 10256 : Déplacement de l'écriture de réinitialisation au 01/01 au matin =>
            il n'est plus nécessaire de supprimer les opération du 01/01
{---------------------------------------------------------------------------------------}
function GetSoldeNatureCompte(Dossier, NatCpt, Endate, Nature, Banque : string; EnValeur, Millesime : Boolean) : Double;
{---------------------------------------------------------------------------------------}
var
  SQL 	: string ;
  Q	: TQuery ;
  Solde	: Double ;
begin
  {Ok est à true si on veut le solde au "matin" du 01/01}
  if Banque <> '' then
    Banque := GetClauseIn(Banque);
  Solde := 0;

  {1/ CALCUL DU SOLDE AU PREMIER JANVIER}
  if EnValeur then SQL := 'SELECT SUM(TE_SOLDEDEVVALEUR / TE_COTATION) FROM TRECRITURE'
              else SQL := 'SELECT SUM(TE_SOLDEDEV / TE_COTATION) FROM TRECRITURE';

  SQL := SQL + ' LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL WHERE BQ_NATURECPTE = "' + NatCpt + '" ';
  if Banque <> '' then
    SQL := SQL + 'AND BQ_BANQUE IN (' + Banque + ') ';

  if (Dossier <> '')  then
    SQL := SQL + 'AND TE_NODOSSIER = "' + Dossier + '" ';
  {Sur les écritures de réinitialisation, les dates d'opération et de valeur sont identitiques}
  SQL := SQL + 'AND TE_DATECOMPTABLE = "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" ';
  SQL := SQL + 'AND TE_CODEFLUX = "' + CODEREGULARIS + '"';

  Q := OpenSQL(SQL, True);
  if not Q.EOF then Solde := Q.Fields[0].AsFloat;
  Ferme(Q);

  {Si on demande le solde au 01/01 ou 02/01, il est inutile de lancer la requête qui renverra vide
   07/11/06 : FQ 10256 : on limite le Exit au 01/01. Pour le 02/01, il faut ajouter les opérations du 01/01
   09/11/06 : FQ 10256 : Pour le 31/12 au soir, on se contente du solde au 01/01 au matin !!!
  if Ok or (Endate = DateToStr(DebutAnnee(StrToDate(Endate)) + 1)) then begin}
  if Millesime or ((StrToDate(Endate) + 1) = DebutAnnee(StrToDate(Endate) + 1)) then begin
    Result := Solde;
    Exit;
  end;

  {3/ CALCUL DU SOLDE DEPUIS LE PREMIER JANVIER JUSQU'À LA VEILLE DE LA DATE PASSÉE EN PARAMÈTRE}
  SQL := 'SELECT SUM(TE_MONTANT) FROM TRECRITURE LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL WHERE BQ_NATURECPTE = "' + NatCpt + '" ';
  if Banque <> '' then
    SQL := SQL + 'AND BQ_BANQUE IN (' + Banque + ') ';

  if (Dossier <> '')  then
    SQL := SQL + 'AND TE_NODOSSIER = "' + Dossier + '"  ';

  {07/11/06 : FQ 10256 : Maintenant que l'écriture de Réinitialisation est au Matin, il faut que la date soit
              supérieure ou égale au 01/01 et non plus strictement supérieure}
  if EnValeur then
    SQL := SQL + 'AND TE_DATEVALEUR <= "' + UsDateTime(StrToDate(Endate)) + '" AND TE_DATEVALEUR >= "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" '
  else
    SQL := SQL + 'AND TE_DATECOMPTABLE <= "' +UsDateTime(StrToDate(Endate)) + '" AND TE_DATECOMPTABLE >= "' + UsDateTime(DebutAnnee(StrToDate(Endate))) + '" ';

  if Trim(Nature) <> '' then
    SQL := SQL + ' ' + Nature;


  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Solde := Solde + Q.Fields[0].AsFloat;
  Ferme(Q);

  Result := Solde;
end;

{27/10/04 : Appels mutliples au recalcul des soldes : les comptes et les dates de départ sont stockés dans la StringList
{---------------------------------------------------------------------------------------}
procedure MultiRecalculSolde (var ATraiter : TStringList);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  O : TObjDtValeur;
begin
  if IsNewSoldes then Exit;
  
  {Recalcul des soldes par compte bancaire}
  for n := 0 to ATraiter.Count - 1 do begin
    O := TObjDtValeur(ATraiter.Objects[n]);
    RecalculSolde(ATraiter[n], DateToStr(O.DateVal), 0, True);
  end;
end;

{20/11/06 : Recalcul des soldes en valeurs uniquement}
{---------------------------------------------------------------------------------------}
function  RecalculSoldeValeur (Compte, DateDepart : string) : Double;
{---------------------------------------------------------------------------------------}
var
  SoldePrec : Double ;
  I         : Integer;
  SoldeMAJ  : Double ;
  Montant   : Double ;
  TobTRE    : TOB ;
  PremOk    : Boolean;
begin
  Result := 0;

  if IsNewSoldes then Exit;

  SoldeMAJ := 0;
  if Compte <> '' then begin
    TobTRE := Tob.Create('£TRECRITURE', nil, -1);
    try
      TobTRE.LoadDetailDBFromSQL('£TRECRITURE', 'SELECT TE_CLEVALEUR, TE_QUALIFORIGINE, TE_SOLDEDEVVALEUR, ' +
                                 'TE_MONTANTDEV, TE_NUMTRANSAC, TE_NUMLIGNE, TE_NODOSSIER, TE_NUMEROPIECE, ' +
                                 'TE_DATEVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + Compte +
                                '" AND TE_DATEVALEUR >= "' + UsDateTime(StrToDateTime(DateDepart)) + '"', True);
      if TobTRE.Detail.Count > 0 then begin
        PremOk := True;
        {solde en date d'opération}
        TobTRE.detail.Sort('TE_CLEVALEUR');
        for i := 0 to TobTRE.Detail.Count - 1 do begin
          {Si on tombe sur une écriture de réinitialisation}
          if tobTRE.Detail[i].GetString('TE_QUALIFORIGINE') = CODEREINIT then begin
            {On récupère son solde, mais on ne le recalcule pas}
            SoldeMAJ := tobTRE.Detail[i].GetDouble('TE_SOLDEDEVVALEUR');
            PremOk := False;
          end

          else begin
            {Récupération du montant à cumuler}
            Montant := TobTRE.Detail[i].GetDouble('TE_MONTANTDEV');

            {C'est la première écriture que l'on traite}
            if PremOk then begin
              SoldePrec := GetSoldePrecValeur(Compte, TobTRE.Detail[0].GetString('TE_CLEVALEUR'));
              {Calcul du nouveau solde}
              SoldeMAJ := Montant + SoldePrec;
              PremOk := False;
            end
            else
              SoldeMAJ := Montant + SoldeMAJ; {On utilise la valeur précédente}

            tobTRE.Detail[i].SetDouble('TE_SOLDEDEVVALEUR', SoldeMAJ);
          end;
        end;
        Result := SoldeMAJ ;
      end;

      InitMove(TobTre.Detail.Count, 'Calcul des soldes');
      for i := 0 to TobTre.Detail.Count - 1 do begin
        MoveCur(False);
        ExecuteSQL('UPDATE TRECRITURE SET TE_SOLDEDEVVALEUR = ' + StrFPoint(TobTRE.Detail[i].GetDouble('TE_SOLDEDEVVALEUR')) +
                   ' WHERE TE_NODOSSIER = "' + TobTRE.Detail[i].GetString('TE_NODOSSIER') +
                   '" AND TE_NUMTRANSAC = "' + TobTRE.Detail[i].GetString('TE_NUMTRANSAC') +
                   '" AND TE_NUMEROPIECE = ' + TobTRE.Detail[i].GetString('TE_NUMEROPIECE') +
                   '  AND TE_NUMLIGNE = ' + TobTRE.Detail[i].GetString('TE_NUMLIGNE') );
      end;
      FiniMove;
    finally
      TobTRE.Free;
    end ;
  end;
end;

{Recalcule des soldes d'un compte à partir d'une date de départ
{---------------------------------------------------------------------------------------}
function RecalculSolde(Compte, DateDepart : string; SoldeInit : Double; RecupSoldePrec : Boolean; SoldeVal : Double = 0.001) : Double ;
{---------------------------------------------------------------------------------------}
var
  SoldePrec : Double ;
  I         : Integer;
  SoldeMAJ  : Double ;
  Montant   : Double ;
  DateRefC  : TDateTime; {27/04/05 : FQ 10243}
  DateRefV  : TDateTime; {27/04/05 : FQ 10243}
  TobTRE    : TOB ;
  PremOk    : Boolean;
  ACheval   : Boolean; {19/01/05}
  JourDeb   : string; {19/01/05}
  JourFin   : string; {19/01/05}
begin
  Result := 0;

  if IsNewSoldes then Exit;

  SoldeMAJ := 0;
  if Compte <> '' then begin
    TobTRE := Tob.Create('£TRECRITURE', nil, -1);
    try
      {JP 29/08/03 : On part de X jours avant la date demandée, pour être sûr que lors du calcul des soldes en valeur
                     il ne manque pas une ou plusieurs lignes dans le cas où des dates de valeur sont supérieures à
                     celles retournées par la requête ne figureraient dans la TOB.
      Exemple :  ID  | date comptable  |  date de valeur
                  1      01/01/2003         05/02/2003
                  2      01/02/2003         02/02/2003
                  3      01/03/2003         02/03/2003
                Ainsi si DateDepart = 01/02/2003, la TOB ne contiendra que les enregistrements 2 et 3 => le calcul
                en valeur sera faux car l'enregistrement 1 se situe en valeur entre le 2 et le 3

      Ceci ne vaut que pour un recalcul après création d'écritures. Dans le cas d'une initialisation ou d'un calcul
      annuel, une écriture a été ajoutée en fin de journée du 01/01 qui contient les nouveaux soldes en Valeur et
      en opération : on partira de ces soldes}

      {27/04/05 : FQ 10243 : pour la requête sur les dates de valeur}
      DateRefC := StrToDate(DateDepart);

      {On commence par regarder s'il s'agit d'un recalcul annuel}
//      DecodeDate(StrToDate(DateDepart), a, m, j);
  //    PremOk := (m = 1) and (j = 2);

      {Si SoldeVal est <> 0, c'est que l'on fait une initialisation.
       22/12/04 : le compilateur n'interprète pas 0.001 comme un Double ce qui fait que quand
                  SoldeVal = 0.001, l'évaluateur dit que SoldeVal <> 0.001 !! => FloatToStr}
      if (FloatToStr(SoldeVal) <> FloatToStr(0.001)) {or PremOk 03/05/05 : cela pose un problème pour les écritures en opé du 02/01} then
        DateRefV := DateRefC
      else
        {03/05/05 : Nouvelle gestion de l'écart de date}
        DateRefV := DateRefC - Abs(ValeurI(GetParamSocSecur('SO_TRNBJOURSSOLDE', 10)));

      {Par défaut, on part du principe que l'écriture de réinitialisation ne figure pas dans la Tob}
      ACheval := True;

      {26/12/06 : Remplacement de TE_EXERCICE par TE_NODOSSIER}
      TobTRE.LoadDetailDBFromSQL('£TRECRITURE', 'SELECT TE_CLEOPERATION, TE_CLEVALEUR, TE_DATECOMPTABLE, TE_QUALIFORIGINE, ' +
                               'TE_SOLDEDEV, TE_SOLDEDEVVALEUR, TE_MONTANTDEV, TE_NUMTRANSAC, TE_NUMLIGNE, ' +
                               'TE_NODOSSIER, TE_NUMEROPIECE, TE_DATEVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + Compte +
                              '" AND TE_DATECOMPTABLE >= "' + UsDateTime(DateRefC) + '"', True);
      if TobTRE.Detail.Count > 0 then begin
        PremOk := True;
        {solde en date d'opération}
        TobTRE.detail.Sort('TE_CLEOPERATION');
        for i := 0 to TobTRE.Detail.Count - 1 do begin
          {Si on tombe sur une écriture de réinitialisation}
          if tobTRE.Detail[i].GetString('TE_QUALIFORIGINE') = CODEREINIT then begin
            {On récupère son solde, mais on ne le recalcule pas}
            SoldeMAJ := tobTRE.Detail[i].GetDouble('TE_SOLDEDEV');
            ACheval := False;{19/01/05 : Pour dire que l'écriture de réinitialisation figure dans la TOB}
          end

          else begin
            {Récupération du montant à cumuler}
            Montant := TobTRE.Detail[i].GetDouble('TE_MONTANTDEV');

            {C'est la première écriture que l'on traite}
            if PremOk then begin
              {Récupération du solde de départ}
              if RecupSoldePrec then begin
            //    if i <> 0 then
              //    SoldePrec := TobTRE.Detail[i - 1].GetDouble('TE_SOLDEDEV')
                //else
                SoldePrec := GetSoldePrec(Compte, TobTRE.Detail[0].GetString('TE_CLEOPERATION'));
              end
              else
                SoldePrec := SoldeInit;
              {Calcul du nouveau solde}
              SoldeMAJ := Montant + SoldePrec;
              PremOk := False;
            end
            else
              SoldeMAJ := Montant + SoldeMAJ; {On utilise la valeur précédente}

            tobTRE.Detail[i].SetDouble('TE_SOLDEDEV', SoldeMAJ);
          end;
        end;
        Result := SoldeMAJ ;

        {27/04/05 : FQ 10243 : j'avais supprimé la deuxième requête, ce qui était une grande erreur, car il pouvait
                    manqué des écritures : on rajoute les écriture qui pourrait manquer
         26/12/06 : Remplacement de TE_EXERCICE par TE_NODOSSIER}
        TobTRE.LoadDetailDBFromSQL('£TRECRITURE', 'SELECT TE_CLEOPERATION, TE_CLEVALEUR, TE_DATECOMPTABLE, TE_QUALIFORIGINE, ' +
                                   'TE_SOLDEDEV, TE_SOLDEDEVVALEUR, TE_MONTANTDEV, TE_NUMTRANSAC, TE_NUMLIGNE, ' +
                                   'TE_NODOSSIER, TE_NUMEROPIECE, TE_DATEVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + Compte +
                                   '" AND TE_DATECOMPTABLE < "' + UsDateTime(DateRefC) +
                                   '" AND TE_DATEVALEUR >= "' + UsDateTime(DateRefV) + '"', True);
        {Solde en date de valeur}
        TobTRE.Detail.Sort('TE_CLEVALEUR');

        {19/01/05 : Il y a un cas qui n'est pas géré, c'est si l'écriture de réinitialisation n'est pas
                    dans la TOB mais que des écritures sont antérieures au 02/01 et ce malgré la requête
                    sur les conditions de valeurs (les conditions de valeurs s'appuient aussi sur la table
                    des frais dans le cas d'une écriture de commission), ou bien dans le recalcul des soldes
                    où l'on prend toutes écritures dont la date comptable >= 02/01:
                    date opération    date valeur     montant
                    19/01/05          01/01/05        150
                    20/01/05          15/01/05        200
                    21/01/05          23/01/05        -300
                    Dans le cas ci-dessus, on ne tient pas compte de l'écriture de réinitialisation et
                    les soldes en valeur sont cumulés avec le dernier solde de l'année 2004 !!}

        {Si au moins deux enregistrements et que l'écriture de réinitialisation ne figure pas dans la Tob}
        if (TobTRE.Detail.Count > 1) and ACheval then begin
          JourDeb := Copy(TobTre.Detail[0].GetString('TE_CLEVALEUR'), 3, 4);
          JourFin := Copy(TobTre.Detail[TobTRE.Detail.Count - 1].GetString('TE_CLEVALEUR'), 3, 4);
          {19/01/05 : on est à cheval si on a des écritures antérieures et postérieures au 01/01}
         // ACheval := (JourDeb <= '0101') and (JourFin > '0101');
        end
        else if ACheval then
          //ACheval := False;
        {23/11/06 : Avec les réinitialisation au matin, cela me semble bien inutile}
        ACheval := False;
        
        for i := 0 to TobTRE.Detail.Count - 1 do begin
          {25/05/05 : On peut se retrouvé avec une écriture dont la date d'opération est supérieure à la
                      date passée en paramêtre de la fonction, mais dont la date de valeur serait inférieure
                      à DateRefV, ce qui fait fait que, sans le continue, on fausserait tout les soldes}
          if TobTre.Detail[i].GetDateTime('TE_DATEVALEUR') < DateRefV then Continue;
          {19/01/05 : on récupère le jour sur la cle en valeur}
          JourDeb := Copy(TobTre.Detail[i].GetString('TE_CLEVALEUR'), 3, 4);
          {Si on tombe sur une écriture de réinitialisation}
          if tobTRE.Detail[i].GetString('TE_QUALIFORIGINE') = CODEREINIT then begin
            {On récupère son solde, mais on ne le recalcule pas}
            SoldeMAJ := tobTRE.Detail[i].GetDouble('TE_SOLDEDEVVALEUR');
            ACheval := False; {19/01/05}
          end

          {19/01/05 : Si ACheval et le jour est le 1er janvier, c'est que le solde de
           réinitialisation ne figure pas dans la Tob}
          else if ACheval and (JourDeb > '0101') then begin
            SoldeMAJ := GetSoldeMillesime(Compte, TobTRE.Detail[i].GetString('TE_DATECOMPTABLE'), '', True, False);
            SoldeMaj := SoldeMAJ + TobTRE.Detail[i].GetDouble('TE_MONTANTDEV');
            TobTRE.Detail[i].SetDouble('TE_SOLDEDEVVALEUR', SoldeMAJ);
            ACheval := False;
          end

          else begin
            Montant := TobTRE.Detail[i].GetDouble('TE_MONTANTDEV');
            if i = 0 then begin
              {Si SoldeVal est <> 0, c'est que l'on fait une initialisation}
              if (StrFPoint(SoldeVal) <> '0.001') then
                SoldePrec := SoldeVal
              else
                SoldePrec := GetSoldePrecValeur(Compte, TobTRE.Detail[i].GetString('TE_CLEVALEUR'));
              SoldeMAJ := Montant + SoldePrec; {On calcul avec la valeur de la base ou demandé}
            end
            else
              SoldeMAJ := Montant + SoldeMAJ; {On utilise la valeur précédente}
            TobTRE.Detail[i].SetDouble('TE_SOLDEDEVVALEUR', soldeMAJ);
          end;
        end;
      end;

      InitMove(TobTre.Detail.Count, 'Calcul des soldes');
      for i := 0 to TobTre.Detail.Count - 1 do begin
        MoveCur(False);
        {26/12/06 : Remplacement de TE_EXERCICE par TE_NODOSSIER}
        ExecuteSQL('UPDATE TRECRITURE SET TE_SOLDEDEVVALEUR = ' + StrFPoint(TobTRE.Detail[i].GetDouble('TE_SOLDEDEVVALEUR')) +
                   ', TE_SOLDEDEV = ' + StrFPoint(TobTRE.Detail[i].GetDouble('TE_SOLDEDEV')) +
                   ' WHERE TE_NODOSSIER = "' + TobTRE.Detail[i].GetString('TE_NODOSSIER') +
                   '" AND TE_NUMTRANSAC = "' + TobTRE.Detail[i].GetString('TE_NUMTRANSAC') +
                   '" AND TE_NUMEROPIECE = ' + TobTRE.Detail[i].GetString('TE_NUMEROPIECE') +
                   '  AND TE_NUMLIGNE = ' + TobTRE.Detail[i].GetString('TE_NUMLIGNE') );
      end;
      FiniMove;
    finally
      TobTRE.Free;
    end ;
  end;
end;

{17/08/07 : Recalcul des soldes d'initialisation
{---------------------------------------------------------------------------------------}
function RecalculSoldeInit(Compte : string; Millesime : Integer) : TOB;
{---------------------------------------------------------------------------------------}
var
  DateRef : TDateTime;
  TobInit : TOB;
  TobSolO : TOB;
  TobSolV : TOB;
  FV      : TOB;
  FO      : TOB;
  SoldeO  : Double;
  SoldeV  : Double;
  TempO   : Double;
  TempV   : Double;
  S       : string;
  n       : Integer;
  F       : TOB;
begin
  Result := nil;
  {On se positionne sur le millésime précédent de celui dont on recalcule le solde initial}
  DateRef := EncodeDate(Millesime - 1, 1, 1);

  TobInit := TOB.Create('mpmp', nil, -1);
  TobSolV := TOB.Create('pmpm', nil, -1);
  TobSolO := TOB.Create('ppmm', nil, -1);
  {Intégration dans TRECRITURE des écritures concernées par la synchronisation}
  Initmove(100, TraduireMemoire('Solde d''initialisation'));
  try
    MoveCur(False);
    {Tob contenant tous les soldes de réinitialisation}
    S := 'SELECT TE_GENERAL, TE_DATECOMPTABLE, TE_SOLDEDEV, TE_SOLDEDEVVALEUR, ' +
         'TE_DEVISE FROM TRECRITURE WHERE TE_CODEFLUX = "' + CODEREGULARIS + '"';
    S := S + ' AND TE_DATECOMPTABLE = "' + UsDateTime(DateRef) + '"';
    if Compte <> '' then S := S + ' AND TE_GENERAL = "' + Compte + '"';
    MoveCur(False);
    TobInit.LoadDetailFromSQL(S);
    MoveCur(False);

    {Tob contenant le cumul des mouvements en date d'opération par généraux et par millésime}
    S := 'SELECT SUM(TE_MONTANTDEV) ASOLDE, TE_GENERAL FROM TRECRITURE ' +
         'WHERE TE_NATURE = "' + na_Realise + '"';
    S := S + ' AND TE_DATECOMPTABLE BETWEEN "' + UsDateTime(DateRef) + '" AND "' + UsDateTime(FinAnnee(DateRef)) + '"';
    if Compte <> '' then S := S + ' AND TE_GENERAL = "' + Compte + '"';
    S := S + ' GROUP BY TE_GENERAL';
    MoveCur(False);
    TobSolO.LoadDetailFromSQL(S);
    MoveCur(False);

    {Tob contenant le cumul des mouvements en date de valeur par généraux et par millésime}
    S := 'SELECT SUM(TE_MONTANTDEV) ASOLDE, TE_GENERAL FROM TRECRITURE ' +
         'WHERE TE_CODEFLUX <> "' + CODEREGULARIS + '" AND TE_NATURE = "' + na_Realise + '"';
    S := S + ' AND TE_DATEVALEUR BETWEEN "' + UsDateTime(DateRef) + '" AND "' + UsDateTime(FinAnnee(DateRef)) + '"';
    if Compte <> '' then S := S + ' AND TE_GENERAL = "' + Compte + '"';
    S := S + ' GROUP BY TE_GENERAL';
    MoveCur(False);
    TobSolV.LoadDetailFromSQL(S);
    MoveCur(False);

    if TobInit.Detail.Count > 0 then Result := TOB.Create('MMMM', nil, -1);

    for n := 0 to TobInit.Detail.Count - 1 do begin
      MoveCur(False);
      FV := TobSolV.FindFirst(['TE_GENERAL'], [TobInit.Detail[n].GetString('TE_GENERAL')], True);
      FO := TobSolO.FindFirst(['TE_GENERAL'], [TobInit.Detail[n].GetString('TE_GENERAL')], True);
      if Assigned(FV) then TempV := FV.GetDouble('ASOLDE')
                      else TempV := 0;
      if Assigned(FO) then TempO := FO.GetDouble('ASOLDE')
                      else TempO := 0;
      SoldeV := {SoldeV +} TempV + TobInit.Detail[n].GetDouble('TE_SOLDEDEVVALEUR');
      SoldeO := {SoldeO +} TempO + TobInit.Detail[n].GetDouble('TE_SOLDEDEV');
      CreerEcritureInit(FinAnnee(TobInit.Detail[n].GetDateTime('TE_DATECOMPTABLE')) + 1,
                        TobInit.Detail[n].GetString('TE_DEVISE'),
                        TobInit.Detail[n].GetString('TE_GENERAL'), SoldeV, SoldeO);
      MoveCur(False);
      F := TOB.Create('MMMM', Result, -1);
      F.AddChampSupValeur('COMPTE'   , RechDom('TRBANQUECP', TobInit.Detail[n].GetString('TE_GENERAL'), False));
      F.AddChampSupValeur('DATE'     , DateToStr(FinAnnee(TobInit.Detail[n].GetDateTime('TE_DATECOMPTABLE')) + 1));
      F.AddChampSupValeur('VALEUR'   , SoldeV);
      F.AddChampSupValeur('OPERATION', SoldeO);
      F.AddChampSupValeur('DEVISE'   , TobInit.Detail[n].GetString('TE_DEVISE'));
    end;

  finally
    FreeAndNil(TobInit);
    FreeAndNil(TobSolV);
    FreeAndNil(TobSolO);
    FiniMove;
  end;
end;

{22/08/07 : Récupère la liste des soldes d'initialisation
{---------------------------------------------------------------------------------------}
function GetListeSoldesInit(Compte : HString; DateDepart : TDateTime) : TOB;
{---------------------------------------------------------------------------------------}
var
  TobInit : TOB;
  n       : Integer;
  S       : Hstring;
  F       : TOB;
begin
  Result := nil;
  TobInit := TOB.Create('mpmp', nil, -1);
  try
    {Tob contenant tous les soldes de réinitialisation}
    S := 'SELECT TE_GENERAL, TE_DATECOMPTABLE, TE_SOLDEDEV, TE_SOLDEDEVVALEUR, ' +
         'TE_DEVISE FROM TRECRITURE WHERE TE_CODEFLUX = "' + CODEREGULARIS + '"';
    S := S + ' AND TE_DATECOMPTABLE >= "' + UsDateTime(DateDepart) + '"';
    if Compte <> '' then S := S + ' AND TE_GENERAL = "' + Compte + '"';
    S := S + ' ORDER BY TE_GENERAL, TE_DATECOMPTABLE';
    TobInit.LoadDetailFromSQL(S);

    if TobInit.Detail.Count > 0 then begin
      Result := TOB.Create('MMMM', nil, -1);
      for n := 0 to TobInit.Detail.Count - 1 do begin
        F := TOB.Create('MMMM', Result, -1);
        F.AddChampSupValeur('COMPTE'   , RechDom('TRBANQUECP', TobInit.Detail[n].GetString('TE_GENERAL'), False));
        F.AddChampSupValeur('DATE'     , TobInit.Detail[n].GetString('TE_DATECOMPTABLE'));
        F.AddChampSupValeur('VALEUR'   , TobInit.Detail[n].GetDouble('TE_SOLDEDEVVALEUR'));
        F.AddChampSupValeur('OPERATION', TobInit.Detail[n].GetDouble('TE_SOLDEDEV'));
        F.AddChampSupValeur('DEVISE'   , TobInit.Detail[n].GetString('TE_DEVISE'));
      end;
    end;
  finally
    FreeAndNil(TobInit);
  end;
end;

{05/01/05 : Mémorisation des comptes dont on doit recalculer les soldes
{---------------------------------------------------------------------------------------}
procedure AddGestionSoldes(var lSolde : TStringList; Cpte : string; Dt : TDateTime);
{---------------------------------------------------------------------------------------}
var
  Obj : TObjDtValeur;
  n   : Integer;
begin
  n := lSolde.IndexOf(Cpte);
  {Le compte est déjà mémorisé dans la liste}
  if n > -1 then begin
    Obj := TObjDtValeur(lSolde.Objects[n]);
    {On mémorise la plus vieille date}
    if Obj.DateVal > Dt then Obj.DateVal :=  Dt;
  end
  {Sinon on crée une nouvelle entrée dans la liste pour ce compte}
  else begin
    Obj := TObjDtValeur.Create;
    Obj.DateVal := Dt;
    lSolde.AddObject(Cpte, Obj);
  end;
end;

{22/07/05 : Mise à jour des soldes par requête
{---------------------------------------------------------------------------------------}
procedure RecalculSoldeSQL(Cpte, CleVal, CleOpe : string; Dt : TDateTime; Mnt : Double; Etat : TEtatEcriture);
{---------------------------------------------------------------------------------------}

    {Dans le cas d'une insertion d'écriture, on récupère le solde précédent
    {------------------------------------------------------------------}
    procedure SetSoldesCurr;
    {------------------------------------------------------------------}
    var
      Q : TQuery;
      V : Double;
      O : Double;
    begin
      {Recherche du solde en opération}
      Q := OpenSQL('SELECT ##TOP 1## TE_SOLDEDEV FROM TRECRITURE WHERE TE_CLEOPERATION < "' + CleOpe +
                   '" AND TE_GENERAL = "' + Cpte + '" ORDER BY TE_CLEOPERATION DESC', True);
      if not Q.EOF then O := Q.FindField('TE_SOLDEDEV').AsFloat
                   else O := 0;
      Ferme(Q);

      {Recherche du solde en valeur}  {JP 21/11/05 : FQ 10304 : TE_CLEVALEUR et non TE_CLEOPERATION !!!!}
      Q := OpenSQL('SELECT ##TOP 1## TE_SOLDEDEVVALEUR FROM TRECRITURE WHERE TE_CLEVALEUR < "' + CleVal +
                   '" AND TE_GENERAL = "' + Cpte + '" ORDER BY TE_CLEVALEUR DESC', True);
      if not Q.EOF then V := Q.FindField('TE_SOLDEDEVVALEUR').AsFloat
                   else V := 0;
      Ferme(Q);

      {Mise à jour des soldes de l'écriture courantes}
      ExecuteSQL('UPDATE TRECRITURE SET TE_SOLDEDEV = ' + StrFPoint(O) + ', TE_SOLDEDEVVALEUR = ' + StrFPoint(V) +
                 ' WHERE TE_GENERAL = "' + Cpte + '" AND TE_CLEOPERATION = "' + CleOpe + '"');
    end;

    {Recalcule les soldes par requêtes en ajoutant le montant de l'écriture
     courante à toutes les écritures postérieures
    {------------------------------------------------------------------}
    procedure MajSoldeSQL;
    {------------------------------------------------------------------}
    var
      CC : string;
      CV : string;
      DC : TDateTime;
      DV : TDateTime;
      Q : TQuery;
      O : string;
      V : string;
      J : Word;
      M : Word;
      A : Word;
    begin
      {19/05/06 : FQ 10339 : Si l'écriture est le premier jour de l'année, il faut chercher l'écriture
                  de réinitialisation de ce jour et non de l'année suivante, comme cela était fait
                  systématiquement avant de l'ajout du test ci dessous}
      if Dt <> DebutAnnee(Dt) then DC := FinAnnee(Dt) + 1
                              else DC := Dt;

      {Recherche d'une éventuelle écriture de réinitialisation au début de l'exercice suivant de
       l'écriture courante : sur une écriture de réinitialisation, les deux clefs sont identiques}
      Q := OpenSQL('SELECT TE_CLEVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + Cpte +
                   '" AND TE_DATEVALEUR = "' + UsDateTime(DC) +
                   '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"', True);
      if not Q.EOF then CC := Q.FindField('TE_CLEVALEUR').AsString;
      Ferme(Q);

      {05/06/06 : FQ 10339 : Certaines dates n'ont pu être calculée faute d'un bon paramétrage}
      if CleVal = '' then Exit;

      A := StrToInt(Copy(CleVal, 1, 2)) + 2000; {Je fais le pari que l'on aura pas de 199X}
      M := StrToInt(Copy(CleVal, 3, 2));
      J := StrToInt(Copy(CleVal, 5, 2));
      DV := EncodeDate(A, M, J);
      {05/06/06 : FQ 10339 : si la date de valeur et la date d'opération entourent le 01/01,
                  dans un certain nombre de cas, l'écriture de réinitialisation en valeur ne
                  sera pas celle en opération. Exemples :
                  - DV = 02/01/2006 -> ini du 01/01/2007 et DC = 01/01/2006 -> ini du 01/01/2006
                  - DV = 30/12/2005 -> ini du 01/01/2006 et DC = 02/01/2006 -> ini du 01/01/2007
                  Pour simplifier et éviter une erreur dans un if and or, je lance la requête si 
                  les clefs en valeur et en opération sont différentes}
      if CleVal <> CleOpe  then begin
        if DV <> DebutAnnee(DV) then DV := FinAnnee(DV) + 1;

        {Recherche d'une éventuelle écriture de réinitialisation au début de l'exercice suivant de
         l'écriture courante : sur une écriture de réinitialisation, les deux clefs sont identiques}
        Q := OpenSQL('SELECT TE_CLEVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' + Cpte +
                     '" AND TE_DATEVALEUR = "' + UsDateTime(DV) +
                     '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"', True);
        if not Q.EOF then CV := Q.FindField('TE_CLEVALEUR').AsString;
        Ferme(Q);

      end
      else
        CV := CC;

      {Constitution de la clause where sur les clefs : si on a trouvé une écriture de réinitialisation
       postérieure à l'écriture courante, on arrête la mise à jour à celle-là}
      O := '" AND TE_CLEOPERATION >= "' + CleOpe + '"';
      V := '" AND TE_CLEVALEUR >= "' + CleVal + '"'; {JP 21/11/05 : FQ 10304 : TE_CLEVALEUR et non TE_CLEOPERATION !!!!}
      if (CC <> '') then
        {19/05/06 : FQ 10339 : modification de l'affectation de O et V}
        O := O + ' AND TE_CLEOPERATION < "' + CC + '"';
      if (CV <> '') then
        V := V + ' AND TE_CLEVALEUR < "' + CV + '"';  {JP 21/11/05 : FQ 10304 : TE_CLEVALEUR et non TE_CLEOPERATION !!!!}

      {Mise à jour des soldes en opération}
      ExecuteSQL('UPDATE TRECRITURE SET TE_SOLDEDEV = TE_SOLDEDEV + ' + StrFPoint(Mnt) + ' ' +
                 'WHERE TE_GENERAL = "' + Cpte + O);

      {Mise à jour des soldes en valeur}
      ExecuteSQL('UPDATE TRECRITURE SET TE_SOLDEDEVVALEUR = TE_SOLDEDEVVALEUR + ' + StrFPoint(Mnt) + ' ' +
                 'WHERE TE_GENERAL = "' + Cpte + V);
    end;

begin
  if IsNewSoldes then Exit;

  case Etat of
    cso_Delete,
    cso_Update : MajSoldeSQL;
    cso_Insert : begin
                   SetSoldesCurr;
                   MajSoldeSQL;
                 end;
    cso_Init   : ; {Évolution future ?!}
  end;
end;

{17/11/06 : vérifie si OldDt, NewDt sont sur le même millésime : sinon, appel de RecalculInit
{---------------------------------------------------------------------------------------}
function GereSoldeInit(OldDt, NewDt : TDateTime; Compte : string; MntDev : Double; Recalcul : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  if TestDateEtMillesime(OldDt, NewDt) then
    Result := RecalculInit(Min(OldDt, NewDt), Compte, MntDev, Recalcul);
end;

{20/11/06 : OldClef, NewClef sont sur le même millésime : sinon, appel de RecalculInit}
{---------------------------------------------------------------------------------------}
function GereSoldeInitClef(OldClef, NewClef : string; Compte : string; MntDev : Double; Recalcul : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
var
  OldDt : TDateTime;
  NewDt : TDateTime;
  a, m, j : Word;
begin
  Result := True;

  a := 2000 + StrToInt(Copy(OldClef, 1, 2));
  m := StrToInt(Copy(OldClef, 3, 2));
  j := StrToInt(Copy(OldClef, 5, 2));
  OldDt := EncodeDate(a, m, j);

  a := 2000 + StrToInt(Copy(NewClef, 1, 2));
  m := StrToInt(Copy(NewClef, 3, 2));
  j := StrToInt(Copy(NewClef, 5, 2));
  NewDt := EncodeDate(a, m, j);
  if TestDateEtMillesime(OldDt, NewDt) then
    Result := RecalculInit(Min(OldDt, NewDt), Compte, MntDev, Recalcul);
end;

{17/11/06 : recalcul du solde de réinitialisation lorsque l'on tourne autour du 01/01/XX
{---------------------------------------------------------------------------------------}
function RecalculInit(dtVal : TDateTime; Compte : string; MntDev : Double; Recalcul : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
var
  dtCal : TDateTime;
begin
  Result := True;
  BeginTrans;
  try
    MntDev := - MntDev;
    dtCal := dtVal - 1;
    dtVal := FinAnnee(dtVal) + 1;

    ExecuteSQL('UPDATE TRECRITURE SET TE_SOLDEDEVVALEUR = TE_SOLDEDEVVALEUR + ' + StrFPoint(MntDev) +
               ' WHERE TE_DATEVALEUR = "' + UsDateTime(DtVal) + '" AND TE_GENERAL = "' +
               Compte + '" AND TE_CODEFLUX = "' + CODEREGULARIS + '"');
    if Recalcul then
      RecalculSoldeValeur(Compte, DateToStr(dtCal));
    CommitTrans;
  except
    on E : Exception do begin
      Result := False;
      RollBackDiscret;
      if V_PGI.SAV then
        PGIError('Erreur lors de la mise à jour de la réinitialisation avec le message :'#13#13 + E.Message);
    end;
  end;

end;

{28/03/07 : Nouvelle fonction de calcul de solde anticipant sur la suppression des soldes initiaux au premier janvier :
            en V9, la date d'initialisation sera libre}
{---------------------------------------------------------------------------------------}
function  GetSoldeBancaire(Compte, Nature : string; DateRef : TDateTime; EnValeur : Boolean = True; EnPivot : Boolean = False) : Double;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  DateDeb : TDateTime;
  Solde   : Double;
  NoDoss  : string;
  ChpMnt  : string;
  ChpDate : string;
begin
//  Result  := 0;
  Solde   := 0;
  DateDeb := iDate1900;
  NoDoss  := V_PGI.NoDossier;

  if EnPivot then ChpMnt := 'TE_MONTANT'
             else ChpMnt := 'TE_MONTANTDEV';

  if EnValeur then ChpDate := 'TE_DATEVALEUR'
              else ChpDate := 'TE_DATECOMPTABLE';

  {Récupération de l'écriture de réinitialisation}
  Q := OpenSQL('SELECT TE_NODOSSIER, TE_DATEVALEUR, TE_SOLDEDEV, TE_SOLDEDEVVALEUR FROM TRECRITURE WHERE TE_GENERAL = "' +
               Compte + '" AND TE_CODEFLUX = "' + CODEREGULARIS + '" AND TE_DATECOMPTABLE <= "' +
               UsDateTime(DateRef) + '" ORDER BY TE_DATEVALEUR DESC', True);
  try
    if not Q.EOF then begin
      DateDeb := Q.FindField('TE_DATEVALEUR').AsDateTime;
      if EnValeur then Solde := Q.FindField('TE_SOLDEDEVVALEUR').AsFloat
                  else Solde := Q.FindField('TE_SOLDEDEV').AsFloat;
      NoDoss := Q.FindField('TE_NODOSSIER').AsString;
    end
    else
      NoDoss := GetDossierFromBQCP(Compte);
  finally
    Ferme(Q);
  end;

  {Somme des opérations}
  Q := OpenSQL('SELECT SUM(' + ChpMnt + ') TE_SOLDE FROM TRECRITURE WHERE TE_GENERAL = "' + Compte + '" AND ' +
               ChpDate + ' BETWEEN "' + UsDateTime(DateDeb) + '" AND "' + UsDateTime(DateRef) +
               '" AND TE_NODOSSIER = "' + NoDoss + '" ' + Nature,  True);
  try
    if not Q.EOF then Result := Solde + Q.FindField('TE_SOLDE').AsFloat
                 else Result := Solde;

  finally
    Ferme(Q);
  end;
end;

{09/10/07 : Récupération du solde bancaire à partir des relevés bancaire
{---------------------------------------------------------------------------------------}
function  GetSoldeRelevesBQE(Compte : string; DateVal : TDateTime; Dossier : string = '') : Double;
{---------------------------------------------------------------------------------------}
var
  QQ : TQuery;
begin
  Result := GetSoldeMillesime(Compte, DateToStr(DateVal), na_Realise, True, True{Ok}, False);
  QQ := OpenSQL('SELECT SUM(CEL_CREDITEURO + CEL_CREDITDEV - CEL_DEBITEURO - CEL_DEBITDEV) FROM ' +
                GetTableDossier(Dossier, 'EEXBQLIG') + ' WHERE CEL_GENERAL = "' + Compte + '" AND ' +
                'CEL_DATEVALEUR BETWEEN "' + UsDateTime(DebutAnnee(DateVal)) + '" AND "' +
                UsDateTime(DateVal - 1) + '"', True);
  if not QQ.EOF then Result := Result + QQ.Fields[0].AsFloat;
  Ferme(QQ);
end;

end.

