unit AccImportTreso;
{ Unité : Fonction de traitement des synchronisations
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |  Commentaires
--------------------------------------------------------------------------------------
 6.50.001.009  22/07/05   JP   Optimisation du recalcul des soldes avec ajout d'une option
                               FQ 10268 : nouvelle gestion des modes de paiement vides
 6.50.001.018  14/09/05   JP   FQ 10269 : On ne met pas à jour les champs de recalcul des dates
                               de valeur si on n'arrive pas à récupérer le CIB afin de pouvoir
                               relancer le traitement.
                               Exclusion des comptes Divers lettrables de la synchronisation
 6.60.001.001  21/11/05   JP   FQ 10304 : Mauvaise gestion des booleans PaieOk et SoldeOk
 7.09.001.001  08/08/06   JP   Gestion de la synchronisation Multi Sociétés
 7.09.001.010  05/04/07   JP   FQ 10428 : Erreur dans la requête de réparation des rubriques pour les C/C
 8.00.001.017  29/05/07   JP   Reprise de E_NATURETRESO dans TE_CODERAPPRO
 8.00.001.018  05/06/07   JP   FQ 10431 : ajout du montant dans la tob d'alerte sur millésime
--------------------------------------------------------------------------------------}
interface

uses
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HEnt1, HCtrls, Classes, UTob, SysUtils, UObjGen, Constantes, Controls;


{Génère la ligne treso en fonction de la ligne de compta
 22/07/05 : FQ 10268 : Ajout de deux options :
            1/ si PaieOk, on affiche les problèmes de modes de paiement,
               sinon on met à RIE les écritures sans mode de paiement
            2/ si SoldeOk, on fait un recalcul complet des comptes traités,
               sinon, on augmente les soldes du montant de l'écriture intégrée
 10/08/06 : retourne le nombre d'enregitrements traités}
function Synchronisation(T, R : TOB; var Grp : TGrpObjet; PaieOk, SoldeOk : Boolean) : Integer;
{Mise à jour de E_TRESOSYNCHRO d'une écriture comptables à partir d'une tob}
procedure  MajTRESOSYNCHRO(T : TOB; NomBase : string; Value : string = ''; dtValeur : TDateTime = 0);
{Mise à jour des dates de valeur et des cib}
procedure TravaillerDates(ClauseWh : string);
{Mise à jour des codes rubriques}
procedure TravaillerFlux;

var
  LaTob   : TOB;
  LaListe : TStringList;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Commun, UProcGen, UProcSolde, HMsgBox, HStatus, Ent1, UtilPgi, ParamSoc;

procedure UpdateRapprochement(T : TOB; NomBase, NoDossier : string); forward;
procedure TravaillerDevise   (var T : TOB; var Grp : TGrpObjet); forward;
procedure TravaillerDonnees  (var T : TOB; var Grp : TGrpObjet); forward;
{08/08/06 : Le traitement des comptes est plus simple et se fait donc à part pour
            éviter une génération d'erreur intempestive}
procedure TravaillerDonneesMS(var T : TOB; var Grp : TGrpObjet); forward;
{10/08/06 : Pour la gestion des rejets et le recalcul des soldes}
procedure TermineUnDossier(L : TStringList; var T : TOB; G : TGrpObjet); forward;


{---------------------------------------------------------------------------------------}
procedure AjouterRefus(TobExclus : TOB; Msg : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  if Msg = '' then Exit;
  T := TOB.Create('***', TobExclus, -1);
  if EstMultiSoc then T.AddChampSupValeur('DOSSIER', ReadTokenSt(Msg))
                 else ReadTokenSt(Msg);
  T.AddChampSupValeur('COMPTE' , ReadTokenSt(Msg));
  T.AddChampSupValeur('PIÈCE'  , ReadTokenSt(Msg));
  T.AddChampSupValeur('LIGNE'  , ReadTokenSt(Msg));
  {05/06/07 : FQ 10431 : ajout du montant}
  T.AddChampSupValeur('MONTANT', ReadTokenSt(Msg));
  T.AddChampSupValeur('MOTIF'  , ReadTokenSt(Msg));
end;

{---------------------------------------------------------------------------------------}
procedure UpdateRapprochement(T : TOB; NomBase, NoDossier : string);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  SQL := 'UPDATE TRECRITURE SET TE_DATERAPPRO = "' + USDateTime(T.GetDateTime('E_DATEPOINTAGE')) +
         '", TE_REFPOINTAGE = "' + T.GetString('E_REFPOINTAGE') +
         '", TE_CODERAPPRO = "' + T.GetString('E_NATURETRESO') + {29/05/07}
         '", TE_USERMODIF = "' + V_PGI.User +
         '" WHERE TE_EXERCICE = "' + T.GetString('E_EXERCICE') +
         '" AND TE_JOURNAL = "' + T.GetString('E_JOURNAL') +
         '" AND TE_DATECOMPTABLE = "' + USDateTime(T.GetDateTime('E_DATECOMPTABLE')) +
         '" AND TE_NUMEROPIECE = ' + T.GetString('E_NUMEROPIECE') +
         ' AND TE_CPNUMLIGNE = ' + T.GetString('E_NUMLIGNE') +
         ' AND TE_NUMECHE = ' + T.GetString('E_NUMECHE') +
         {08/08/06 : la clef comptable n'est plus suffisante : il faut ajouter la notion de dossier}
         ' AND TE_NODOSSIER = "' + NoDossier + '"';
  ExecuteSQL(SQL);
  MajTRESOSYNCHRO(T, NomBase);
end;

{---------------------------------------------------------------------------------------}
procedure GererListe(var J, G, C : TStringList; NomBase : string);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  {Constitution de la liste contenant les journaux de banque.}
  Q := OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL = "BQE"', True);
  try
    while not Q.EOF do begin
      J.Add(Q.FindField('J_JOURNAL').AsString);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

  {24/09/04 : on traite les journaux d'effets de la même manière suite à FQ 10104}
  Q := OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL = "OD" AND J_EFFET = "X"', True);
  try
    while not Q.EOF do begin
      G.Add(Q.FindField('J_JOURNAL').AsString);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

  {08/08/06 : Récupération des comptes courants}
  Q := OpenSQL('SELECT CLS_GENERAL FROM ' + GetTableDossier(NomBase, 'CLIENSSOC'), True);
  try
    while not Q.EOF do begin
      C.Add(Q.FindField('CLS_GENERAL').AsString);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TravaillerDonnees(var T : TOB; var Grp : TGrpObjet);
{---------------------------------------------------------------------------------------}
var
  Lib  : string;
  Rub  : string;
  Gene : string;
  Paie : string;
  Sens : string;
  dt   : TDateTime;
  CVa  : string;
  CCo  : string;
begin
  {GESTION DU CODE RUBRIQUE POUR TE_CODEFLUX}
  Gene := T.GetString('TE_CONTREPARTIETR');
  Paie := T.GetString('TE_CODECIB');
  if T.GetDouble('TE_MONTANT') > 0 then Sens := 'C'
                                   else Sens := 'D';
  Grp.oRub.GetCorrespondance(Gene, Rub, Lib, Sens, Paie);
  {On ajoute le compte général à la liste de correspondance}
  if Rub = '' then begin
    Grp.oRub.AjouterALaListe(Gene);
    Grp.oRub.GetCorrespondance(Gene, Rub, Lib, Sens, Paie);
  end;
  {On ne met à jour les deux champs que si on a récupéré la rubrique, de manière à ce
   que le traitement puisse être relancé depuis le mul après complétion du paramétrage
   des rubriques}
  if Rub <> '' then begin
    T.SetString('TE_CODEFLUX', Rub);
    T.SetString('TE_USERMODIF', V_PGI.User);
  end;

  {GESTION DES CODECIB ET DONC DES DATES DE VALEURS}
  Lib  := '';

  {Pour les écritures prévisionnelles, il faut commencer par récupérer la véritable date d'opération qui
   est stockée dans TE_DATEVALID, notamment pour que la clause de la requête ci-dessus n'exclut par
   certaines écritures si le filtre se fait sur la date comptable cf ImportTreso.TravaillerCompta}
   if T.GetString('TE_NATURE') = na_Prevision then
     T.SetDateTime('TE_DATECOMPTABLE', T.GetDateTime('TE_DATEVALID'));

  dt   := T.GetDateTime('TE_DATECOMPTABLE');
  Gene := T.GetString('TE_GENERAL');
  Sens := T.GetString('TE_USERVALID');
  Grp.oValeur.GetDateValeur(Gene, Paie, Sens, dt, Lib);
  if Lib = '' then begin
    Rub := Grp.oBqPrev.GetInfosBq(Gene).Banque;
    Grp.oValeur.AjouterALaListe(Rub, Gene, Paie, Sens);
    Grp.oValeur.GetDateValeur(Gene, Paie, Sens, dt, Lib);
  end;

  {On ne met à jour les champs que si on a récupéré le Cib et la date de valeur, de manière
   à ce que le traitement puisse être relancé depuis le mul après complétion du paramétrage
   des Cib et des conditions de valeur}
  if Lib <> '' then begin
    T.SetString('TE_USERCREATION', V_PGI.User);
    T.SetString('TE_CODECIB', Lib);
    {Jusqu'ici TE_USERVALID a servi de champ "SENS", on le met maintenant à vide}
    T.SetString('TE_USERVALID', '');
  end;
  
  {27/02/07 : FQ 10413 : pour éviter que les clefs ne soient vides
  {13/09/06 : Si l'écriture est pointée, théoriquement la date de valeur est renseignée grace à l'écriture comptable}
  if not((T.GetDateTime('TE_DATERAPPRO') > iDate1900) and (T.GetDateTime('TE_DATEVALEUR') > iDate1900)) then
    T.SetDateTime('TE_DATEVALEUR', Dt);

  {Calcul des champs clefs}
  CVa := RetourneCleEcriture(Dt, Grp.iCode);
  CCo := RetourneCleEcriture(T.GetDateTime('TE_DATECOMPTABLE'), Grp.iCode);
  T.SetString('TE_CLEVALEUR'   , CVa);
  T.SetString('TE_CLEOPERATION', CCo);
  Inc(Grp.iCode);

  {20/11/06 : gestion  des dates antérieures au millésime courant}
  Lib := '';
  {Inutile d'ajouter un refus si le montant est vide}
  if Arrondi(T.GetDouble('TE_MONTANT'), V_PGI.OkDecV) <> 0 then begin
    {05/06/07 : FQ 10431 : Ajout du montant}
    if T.GetDateTime('TE_DATEVALEUR') < DebutAnnee(Date) then
      Lib := Grp.NomBase + ';' + T.GetString('TE_GENERAL') + ';' + T.GetString('TE_NUMEROPIECE') + ';' +
             T.GetString('TE_NUMLIGNE') + ';' + T.GetString('TE_MONTANT') + ';' + TraduireMemoire('Date de valeur antérieure au millésime') + ';';
    if T.GetDateTime('TE_DATECOMPTABLE') < DebutAnnee(Date) then
      Lib := Grp.NomBase + ';' + T.GetString('TE_GENERAL') + ';' + T.GetString('TE_NUMEROPIECE') + ';' +
             T.GetString('TE_NUMLIGNE') + ';' + T.GetString('TE_MONTANT') + ';' + TraduireMemoire('Date d''opération antérieure au millésime') + ';';
    AjouterRefus(Grp.TobDate, Lib);
  end;

  {GESTION DES DEVISES}
  TravaillerDevise(T, Grp);
end;

{Concerne les comptes courants dans le traitement est simplifié
{---------------------------------------------------------------------------------------}
procedure TravaillerDonneesMS(var T : TOB; var Grp : TGrpObjet);
{---------------------------------------------------------------------------------------}
var
  Lib  : string;
  Rub  : string;
  Gene : string;
  Paie : string;
  Sens : string;
  CCo  : string;
begin
  {GESTION DU CODE RUBRIQUE POUR TE_CODEFLUX}
  Gene := GetGeneFromBqCode(T.GetString('TE_GENERAL'));
  {Il y a toujours un code Cib, car si le paramétrage de BanqueCp n'est pas complet
   GetCibCourant renvoie CODECIBCOURANT}
  Paie := GetCibCourant(T.GetString('TE_GENERAL'));
  T.SetString('TE_CODECIB', Paie);


  if T.GetDouble('TE_MONTANT') > 0 then Sens := 'C'
                                   else Sens := 'D';
  {On va boucler sur tous les modes de paiement du CIB pour trouver s'il existe une rubriqe
   correspondant au compte et au sens de l'écriture}
  Grp.oRub.GetCorrespondanceMS(Gene, Rub, Lib, Sens, Paie);
  {On ne met à jour les deux champs que si on a récupéré la rubrique, de manière à ce
   que le traitement puisse être relancé depuis le mul après complétion du paramétrage
   des rubriques}
  if Rub <> '' then begin
    T.SetString('TE_CODEFLUX', Rub);
    T.SetString('TE_USERMODIF', V_PGI.User);
  end;

  {GESTION DES CODECIB ET DONC DES DATES DE VALEURS}
  T.SetString('TE_USERCREATION', V_PGI.User);
  T.SetDateTime('TE_DATEVALEUR', T.GetDateTime('TE_DATECOMPTABLE'));
  {Pour que l'on ne puisse pas modifier les dates de valeur}
  T.SetDateTime('TE_DATERAPPRO', T.GetDateTime('TE_DATECOMPTABLE'));
  {Jusqu'ici TE_USERVALID a servi de champ "SENS", on le met maintenant à vide}
  T.SetString('TE_USERVALID', '');
  CCo := RetourneCleEcriture(T.GetDateTime('TE_DATECOMPTABLE'), Grp.iCode);
  T.SetString('TE_CLEVALEUR'   , CCo);
  T.SetString('TE_CLEOPERATION', CCo);
  Inc(Grp.iCode);

  {Pour les écritures prévisionnelles, il faut commencer par récupérer la véritable date d'opération qui
   est stockée dans TE_DATEVALID, notamment pour que la clause de la requête ci-dessus n'exclut par
   certaines écritures si le filtre se fait sur la date comptable cf ImportTreso.TravaillerCompta}
   if T.GetString('TE_NATURE') = na_Prevision then
     T.SetDateTime('TE_DATECOMPTABLE', T.GetDateTime('TE_DATEVALID'));

  {GESTION DES DEVISES}
  TravaillerDevise(T, Grp);
end;

{---------------------------------------------------------------------------------------}
procedure TravaillerDevise(var T : TOB; var Grp : TGrpObjet);
{---------------------------------------------------------------------------------------}
var
  MtD : Double;
  MtE : Double;
  DtV : TDateTime;
  Gen : string;
  Dev : string;
begin
  {Dans la compta, il est possible de saisir dans une autre devise que celle du compte.
   Pour le moment en tréso on ne met dans la table TRECRITURE que les montants en € et dans la devise du compte.}
  Dev := T.GetString('TE_DEVISE');
  Gen := T.GetString('TE_GENERAL');
  DtV := T.GetDateTime('TE_DATEVALEUR');
  {Si la devise a été saisie dans la devise du compte, il n'y a aucun traitement à faire}
  if Dev <> Grp.oDevise.GetDeviseCpt(Gen) then begin
    {JP 28/04/04 : pour limiter les risques d'erreur et les écarts de change, on traite à part
                   le cas du compte de destination en euros}
    if Grp.oDevise.GetDeviseCpt(Gen) = V_PGI.DevisePivot then begin
      MtD := T.GetDouble('TE_MONTANT');
      MtE := MtD;
    end else begin
      {On part du montant en devise car le montant est déjà un montant calculé}
      MtE := T.GetDouble('TE_MONTANTDEV');
      Grp.oDevise.ConvertitMnt(MtE, MtD, Dev, Gen, DtV);
    end;
    {Mais on ne change que le montant en devise, on garde le montant en Euro calculé en comptant}
    {24/05/05 : Uniformisation des formats des montants : Théoriquement, c'est fait dans ConvertitMnt, mais ...}
    T.PutValue('TE_MONTANTDEV', Arrondi(MtD, Grp.oDevise.GetNbDecimalesFromCpt(Gen)));
    T.PutValue('TE_DEVISE', Grp.oDevise.GetDeviseCpt(Gen));
    if MtE <> 0 then T.PutValue('TE_COTATION', Arrondi(MtD / MtE, NBDECIMALTAUX))
                else T.PutValue('TE_COTATION', 1);
  end;
end;

{25/07/05 : FQ 10268 : Nouvelle gestion du recalcul des soldes
{---------------------------------------------------------------------------------------}
function TresoFromCompta(T : TOB; var Grp : TGrpObjet; SoldeOk : Boolean = True) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;{25/07/05 : FQ 10268}
  M : Double;
  s : string;
  Nouveau : Boolean;
  aTob    : TOB;
  SQL     : string;
begin
  Result := '';
  {On ne traite pas les écritures prévisionnelles que l'on ne peut rattacher à aucune banque}
  if (T.GetString('E_BANQUEPREVI') = '') and not Grp.oBqPrev.IsValide(Grp.oBqPrev.NatToIndice(T.GetString('G_NATUREGENE'))) then begin
    Result := Grp.NomBase + ';' + T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Banque prévisionnelle non renseignée') + ';';
    Exit;
  end;

  {11/10/06 : FQ 10365 : synchronisation des écritures bordereau dont le numéro de pièce est inférieur à 10000}
  if (T.GetString('J_MODESAISIE') <> '-') and (T.GetInteger('E_NUMEROPIECE') > 9999) then begin
    Result := Grp.NomBase + ';' + T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Numéro de folio supérieur à 9999') + ';';
    Exit;
  end;

  Grp.sGene := '';
  Grp.dDate := iDate1900;

  aTob := TOB.Create('TRECRITURE', nil, -1);
  try
    if T.GetString('J_MODESAISIE') = '-' then
      {11/10/06 : FQ 10365 : saisie pièce, on concerve la même structure de transaction
                 (sur 17 caractères) : "ICP" + SOC + JAL + PIECE----}
      aTob.SetString('TE_NUMTRANSAC', TRANSACIMPORT + T.GetString('E_SOCIETE') + T.GetString('E_JOURNAL') + T.GetString('E_NUMEROPIECE'))
    else
      {11/10/06 : FQ 10365 : saisie boredreau / libre : "ICP" + SOC + JAL + PERI + PIEC}
      aTob.SetString('TE_NUMTRANSAC', TRANSACIMPORT + T.GetString('E_SOCIETE') + T.GetString('E_JOURNAL') +
                     Copy(T.GetString('E_PERIODE'), 3, 4) + T.GetString('E_NUMEROPIECE'));

    {11/08/06 : gros bug !! ex Ligne   Eche   Ligne Treso
                                 1       1      2
                                 1       2      3 !
                                 2       1      3 !!
                                 3       0      3 !!!
    aTob.SetInteger('TE_NUMLIGNE'    , T.GetInteger('E_NUMLIGNE') + T.GetInteger('E_NUMECHE'));}
    if T.GetInteger('E_NUMECHE') > 1  then
      aTob.SetInteger('TE_NUMLIGNE'  , 100 * T.GetInteger('E_NUMLIGNE') + T.GetInteger('E_NUMECHE'))
    else
      aTob.SetInteger('TE_NUMLIGNE'  , T.GetInteger('E_NUMLIGNE'));
    aTob.SetInteger('TE_NUMECHE'     , T.GetInteger('E_NUMECHE'));
    aTob.SetInteger('TE_CPNUMLIGNE'  , T.GetInteger('E_NUMLIGNE'));
    aTob.SetString('TE_CODECIB'      , T.GetString('E_MODEPAIE'));
    aTob.SetString('TE_JOURNAL'      , T.GetString('E_JOURNAL'));
    aTob.SetInteger('TE_NUMEROPIECE' , T.GetInteger('E_NUMEROPIECE'));
    aTob.SetString('TE_EXERCICE'     , T.GetString('E_EXERCICE'));
    aTob.SetString('TE_REFINTERNE'   , T.GetString('E_REFINTERNE'));
    aTob.SetString('TE_SOCIETE'      , T.GetString('E_SOCIETE'));
    aTob.SetString('TE_ETABLISSEMENT', T.GetString('E_ETABLISSEMENT'));
    aTob.SetString('TE_QUALIFORIGINE', QUALIFCOMPTA);
    aTob.SetString('TE_LIBELLE'      , T.GetString('E_LIBELLE'));
    aTob.SetString('TE_NODOSSIER'    , Grp.NoDossier); {29/05/06 : FQ 10360}

    Nouveau := False;
    {L'écriture comptable a pu être crée puis modifiée avant d'avoir été synchronisée : on s'assure
     qu'elle existe bien
     10/11/06 : Nouvelle gestion du numéro de transaction et de ligne}
    s := 'SELECT TE_MONTANTDEV FROM TRECRITURE WHERE TE_NUMECHE = ' + aTob.GetString('TE_NUMECHE') + ' ' +
         'AND TE_CPNUMLIGNE = ' + aTob.GetString('TE_CPNUMLIGNE') + ' AND TE_JOURNAL = "' +
         aTob.GetString('TE_JOURNAL') + '" AND TE_EXERCICE = "' +
         aTob.GetString('TE_EXERCICE') + '" AND TE_NUMEROPIECE = ' +
         aTob.GetString('TE_NUMEROPIECE') + ' AND TE_NODOSSIER = "' + aTob.GetString('TE_NODOSSIER') + '"';
    {17/11/06 : gestion des bordereaux}
    if (T.GetString('E_MODESAISIE') = 'BOR') or (T.GetString('E_MODESAISIE') = 'LIB') then
      S := S + ' AND TE_NUMTRANSAC LIKE "' + TRANSACIMPORT + T.GetString('E_SOCIETE') + T.GetString('E_JOURNAL') +
                 Copy(T.GetString('E_PERIODE'), 3, 4) + '%"';

    {Si une écriture a été pointée, il y a deux possibilités :
     a/ soit il s'agit d'une écriture présente dans la Trésorerie, et donc il n'y a que la date de rapprochement
        à mettre à jour : il est inutile de poursuivre le traitement
     b/ soit l'écriture a été pointée avant d'avoir été synchronisée une première fois}
    if T.GetString('E_TRESOSYNCHRO') = ets_Pointe then begin
      if ExisteSQL(s) then begin
        UpdateRapprochement(T, Grp.NomBase, Grp.NoDossier);
        Exit;
      end
      else
        Nouveau := True;
    end;

    {27/09/04 : FQ 10119 : Gestion des modes de paiement dans les rubriques}
    aTob.SetString('TE_USERMODIF'     , T.GetString('E_MODEPAIE'));
    aTob.SetString('TE_USERCOMPTABLE' , V_PGI.User);
    aTob.SetString('TE_DEVISE'        , T.GetString('E_DEVISE'));
    {E_Couverture étant un montant en valeur absolu ...}
    if (T.GetDouble('E_DEBIT') - T.GetDouble('E_CREDIT')) >= 0 then begin
      aTob.SetDouble('TE_MONTANT'   , T.GetDouble('E_DEBIT') - T.GetDouble('E_CREDIT') - T.GetDouble('E_COUVERTURE'));
      {24/05/05 : Le format du montant en devise sera effectué dans dans TravaillerDonnees}
      aTob.SetDouble('TE_MONTANTDEV', T.GetDouble('E_DEBITDEV') - T.GetDouble('E_CREDITDEV') - T.GetDouble('E_COUVERTUREDEV'));
    end else begin
      aTob.SetDouble('TE_MONTANT'   , T.GetDouble('E_DEBIT') - T.GetDouble('E_CREDIT') + T.GetDouble('E_COUVERTURE'));
      {24/05/05 : Le format du montant en devise sera effectué dans dans TravaillerDonnees}
      aTob.SetDouble('TE_MONTANTDEV', T.GetDouble('E_DEBITDEV') - T.GetDouble('E_CREDITDEV') + T.GetDouble('E_COUVERTUREDEV'));
    end;
    {La cotation en compta semble être l'inverse de la tréso et de la chancellerie.
     Pour plus de sécurité, on la recalcule, même si le champ est rarement utilisé !
     Dans les tables Chancell et Trecriture MntDev := Cotation * Mnt€}
    aTob.SetDouble('TE_COTATION', 1); {Par défaut}
    if aTob.GetDouble('TE_MONTANT') <> 0 then
     {24/05/05 : La cotation en devise sera effectué dans dans TravaillerDonnees}
      aTob.SetDouble('TE_COTATION'  , Arrondi(aTob.GetDouble('TE_MONTANTDEV') / aTob.GetDouble('TE_MONTANT'), NBDECIMALTAUX));
    aTob.SetDateTime('TE_DATERAPPRO', T.GetDateTime('E_DATEPOINTAGE'));
    aTob.SetString('TE_REFPOINTAGE', T.GetString('E_REFPOINTAGE'));
    aTob.SetString('TE_CODERAPPRO', T.GetString('E_NATURETRESO')); {29/05/07}

    {Dans le cas des écritures prévisionnelles, la date d'opération n'est pas la date comptable mais la date
     d'échéance. Dans un premier temps, on stocke la date d'échéance dans TE_DATEVALID et on mettra à jour
     TE_DATECOMPTABLE (date d'opération) dans le travail sur les dates de valeur : ceci pour une raison, à
     savoir que les critères de sélection reposent, entre autre, sur E_DATECOMPTABLE / TE_DATECOMPTABLE, donc
     si TE_DATECOMPTABLE était différent de E_DATECOMPTABLE, certains traitements ne seraient pas correctement
     effectués, notamment le calcul des dates de valeur car ceratines écritures sont exclues par le filtrage sur
     la date d'opération : MajDeviseEtValeur }
    aTob.SetDateTime('TE_DATECOMPTABLE', T.GetDateTime('E_DATECOMPTABLE'));

    {Ecritures bancaires ...}
    if Grp.oBqPrev.NatToIndice(T.GetString('G_NATUREGENE')) in [nc_BqCais, nc_CouTit] then begin
      aTob.SetString('TE_GENERAL'        , GetBqCodeFromGene(T.GetString('E_GENERAL'), Grp.NoDossier));
      aTob.SetString('TE_CONTREPARTIETR' , T.GetString('E_CONTREPARTIEGEN'));
      aTob.SetString('TE_NATURE'         , na_Realise);
      aTob.SetDateTime('TE_DATEVALID'    , iDate1900);
    end

    {... ou échéancier}
    else begin
      if T.GetString('E_BANQUEPREVI') <> '' then
        aTob.SetString('TE_GENERAL', GetBqCodeFromGene(T.GetString('E_BANQUEPREVI'), Grp.NoDossier))
      else
        aTob.SetString('TE_GENERAL', GetBqCodeFromGene(Grp.oBqPrev.GetCompte(Grp.oBqPrev.NatToIndice(T.GetString('G_NATUREGENE'))), Grp.NoDossier));

      aTob.SetString('TE_CONTREPARTIETR' , T.GetString('E_GENERAL'));
      aTob.SetString('TE_NATURE'         , na_Prevision);
      {Dans le cas des écritures prévisionnelles, E_DATECOMPTABLE n'a pas d'intérêt : on reprend la date d'échéance}
      aTob.SetDateTime('TE_DATEVALID'   , T.GetDateTime('E_DATEECHEANCE'));
    end;

    aTob.SetString('TE_CODEBANQUE' , Grp.oBqPrev.GetEtabBq (aTob.GetString('TE_GENERAL')));
    aTob.SetString('TE_CODEGUICHET', Grp.oBqPrev.GetGuichet(aTob.GetString('TE_GENERAL')));
    aTob.SetString('TE_NUMCOMPTE'  , Grp.oBqPrev.GetNumCpt (aTob.GetString('TE_GENERAL')));
    aTob.SetString('TE_CLERIB'     , Grp.oBqPrev.GetCleRib (aTob.GetString('TE_GENERAL')));
    aTob.SetString('TE_IBAN'       , Grp.oBqPrev.GetIban   (aTob.GetString('TE_GENERAL')));

    {Le code banque est nécessaire en tréso => on n'intègre pas la ligne}
    if (aTob.GetString('TE_CODEBANQUE') = '') then begin
      Result := Grp.NomBase + ';' + T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Code banque non renseigné') + ';';
      Exit;
    end
    {Le CIB est nécessaire en tréso => on n'intègre pas la ligne,
     09/08/06 : sauf s'il s'agit d'une écriture courantes où le mode de paiement n'est pas
                forcément renseigné et que l'on n'utilise pas pour calculer les cib et les rubriques}
    else if (aTob.GetString('TE_CODECIB') = '') and (T.GetString('G_NATUREGENE') <> 'DIV') then begin
      Result := Grp.NomBase + ';' + T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Mode de paiement non renseigné') + ';';
      Exit;
    end
    {Sans Contrepartie, on n'a pas de code rubrique => on n'intègre pas la ligne}
    else if (aTob.GetString('TE_CONTREPARTIETR') = '') then begin
      Result := Grp.NomBase + ';' + T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Contrepartie non renseignée') + ';';
      Exit;
    end;

    aTob.SetDateTime('TE_DATECREATION', Date);
    aTob.SetDateTime('TE_DATEMODIF'   , iDate1900);
    aTob.SetDateTime('TE_DATEVALEUR'  , T.GetDateTime('E_DATEVALEUR'));

    aTob.SetString('TE_COMMISSION', suc_SansCom);

    {Ces valeurs temporaires, permettent de connaître les écritures qui n'auront pas été correctement traités}
    aTob.SetString('TE_USERCREATION'  , CODETEMPO);
    aTob.SetString('TE_CODEFLUX'     , CODEIMPORT);
    {Le champ TE_USERVALID va servir temporairement de Champ SENS}
    if aTob.GetDouble('TE_MONTANTDEV') >= 0 then aTob.SetString('TE_USERVALID', 'ENC')
                                            else aTob.SetString('TE_USERVALID', 'DEC');
    {Renseignement des champs qui demandent des traitements : Date de valeur, clefs, Code flux, CIB ...
     09/08/06 : pour les comptes courants (et titres ...), on ne part du mode de paiement pour calculer les
                Cib, les rubriques mais du Cib paramétré dans BANQUECP}
    if T.GetString('G_NATUREGENE') = 'DIV' then TravaillerDonneesMS(aTob, Grp)
                                           else TravaillerDonnees  (aTob, Grp);

    {Si le montant est égal à zéro (on peut supposer que l'écriture vient d'être lettrée) et qu'il s'agit d'une
     écriture de l'échéancier ou bien si l'écriture est complètement lettrée, on la supprime de la trésorerie.
     On fait aussi le test sur le montant en devise pour contourner déventuels problèmes d'arrondi dans le cas
     d'un lettrage en devise}
    if ((aTob.GetDouble('TE_MONTANT') = 0) and (Grp.oBqPrev.NatToIndice(T.GetString('G_NATUREGENE')) <> nc_BqCais)) or
       ((aTob.GetDouble('TE_MONTANTDEV') = 0) and (Grp.oBqPrev.NatToIndice(T.GetString('G_NATUREGENE')) <> nc_BqCais)) or
       (T.GetString('E_ETATLETTRAGE') = 'TL') then begin

      {25/07/05 : FQ 10268 : Nouvelle option de recalcul des soldes}
      M := 0;
      if not SoldeOk then begin
        {Il faut commencer par récupérer l'ancien montant pour le déduire des soldes}
        Q := OpenSQL(s, True);
        if not Q.EOF then M := - Q.FindField('TE_MONTANTDEV').AsFloat;
        Ferme(Q);
      end;

      {Suppression de l'écriture
       10/11/06 : Nouvelle gestion du numéro de transaction et de ligne}
      SQL := 'DELETE FROM TRECRITURE WHERE TE_NUMECHE = ' + aTob.GetString('TE_NUMECHE') + ' ' +
             'AND TE_CPNUMLIGNE = ' + aTob.GetString('TE_CPNUMLIGNE') + ' AND TE_JOURNAL = "' +
             aTob.GetString('TE_JOURNAL') + '" AND TE_EXERCICE = "' +
             aTob.GetString('TE_EXERCICE') + '" AND TE_NUMEROPIECE = ' +
             aTob.GetString('TE_NUMEROPIECE') + ' AND TE_NODOSSIER = "' + aTob.GetString('TE_NODOSSIER') + '"';
      {17/11/06 : gestion des bordereaux}
      if (T.GetString('E_MODESAISIE') = 'BOR') or (T.GetString('E_MODESAISIE') = 'LIB') then
        SQL := SQL + ' AND TE_NUMTRANSAC LIKE "' + TRANSACIMPORT + T.GetString('E_SOCIETE') + T.GetString('E_JOURNAL') +
                     Copy(T.GetString('E_PERIODE'), 3, 4) + '%"';
      ExecuteSQL(SQL);

      {25/07/05 : FQ 10268 : Recalcul des soldes par requêtes
       05/06/07 : il est inutile de lancer la requête si le solde est à zéro (cela signifie que l'écriture
                  prévisionnelle n'avait pas été synchronisé avant d'être totalement lettrée}
      if not SoldeOk and (Arrondi(M, V_PGI.OkDecV) <> 0) then begin
        RecalculSoldeSQL(aTob.GetString('TE_GENERAL'), aTob.GetString('TE_CLEVALEUR'),
                     aTob.GetString('TE_CLEOPERATION'), aTob.GetDateTime('TE_DATECOMPTABLE'), M, cso_Delete);
      end;
      {Mise à Jour De E_TRESOSYNCHRO}
      MajTresoSynchro(T, Grp.NomBase);
    end

    {Gestion des autres cas : écritures bancaires, écritures d'échéanciers non lettrés}
    else begin
      {JP 15/05/07 : FQ 10467 : E_TRESOSYNCHRO est mieux que E_SYNCHROTRESO}
      Nouveau := (Nouveau or (T.GetValue('E_TRESOSYNCHRO') = ets_Nouveau)) and
                 not (T.GetValue('E_TRESOSYNCHRO') = ets_BqPrevi);

      if not Nouveau then
        Nouveau := not ExisteSQL(s);
      {On ne modifie la date de valeur que s'il s'agit d'une écriture de règlement non pointée}
      if (aTob.GetString('TE_NATURE') = na_Realise) and (aTob.GetDateTime('TE_DATERAPPRO') = iDate1900) then
        {Mise à Jour De E_TRESOSYNCHRO}
        MajTresoSynchro(T, Grp.NomBase, ets_Synchro, aTob.GetDateTime('TE_DATEVALEUR'))
      else
        {Mise à Jour De E_TRESOSYNCHRO}
        MajTresoSynchro(T, Grp.NomBase);

      {S'il s'agit d'une nouvelle écriture ou d'une écriture modifiée en compta avant d'avoir été synchronisée}
      if Nouveau then begin
        {Insertion de la nouvelle écriture}
        aTob.InsertDb(nil);

        {25/07/05 : Nouvelle option de recalcul des soldes}
        if not SoldeOk then
          RecalculSoldeSQL(aTob.GetString('TE_GENERAL'), aTob.GetString('TE_CLEVALEUR'),
                    aTob.GetString('TE_CLEOPERATION'), aTob.GetDateTime('TE_DATECOMPTABLE'), aTob.GetDouble('TE_MONTANTDEV'), cso_Insert);
      end
      else  begin
        {25/07/05 : FQ 10268 : Nouvelle option de recalcul des soldes}
        M := 0;
        if not SoldeOk then begin
          {Il faut commencer par récupérer l'ancien montant pour le déduire du nouveau}
          Q := OpenSQL(s, True);
          if not Q.EOF then M := Q.FindField('TE_MONTANTDEV').AsFloat;
          Ferme(Q);

          M := aTob.GetDouble('TE_MONTANTDEV') - M;
        end;

        {Mise à jour de l'écriture}
        aTob.UpdateDB;

        {25/07/05 : FQ 10268 : Recalcul des soldes par requêtes}
        if not SoldeOk then
          RecalculSoldeSQL(aTob.GetString('TE_GENERAL'), aTob.GetString('TE_CLEVALEUR'),
               aTob.GetString('TE_CLEOPERATION'), aTob.GetDateTime('TE_DATECOMPTABLE'), M, cso_Update);
      end;
    end;
    {Pour mémoriser le recalcul des soldes}
    Grp.sGene := aTob.GetString('TE_GENERAL');
    Grp.dDate := aTob.GetDateTime('TE_DATECOMPTABLE');

  finally
    if Assigned(aTob) then FreeAndNil(aTob);
  end;
end;

{22/07/05 : FQ 10268 : Ajout de deux options :
  1/ si PaieOk, on affiche les problèmes de modes de paiement, sinon on met à RIE les écritures sans mode de paiement
  2/ si SoldeOk, on fait un recalcul complet des comptes traités, sinon, on augmente les soldes du montant de l'écriture intégrée
 10/08/06 : retourne le nombre d'écritures traitées
{---------------------------------------------------------------------------------------}
function Synchronisation(T, R : TOB; var Grp : TGrpObjet; PaieOk, SoldeOk : Boolean) : Integer;
{---------------------------------------------------------------------------------------}
var
  TobExclus : TOB;
  lCalcSold : TStringList;
  Obj       : TObjDtValeur;
  ObjRub    : TObjRubrique;
  ObjDtCib  : TDateValeur;
  ObjDev    : TObjDevise;
  S : string;
  n : Integer;
  J : TStringList;
  G : TStringList;
  C : TStringList;
  k : Integer;

    {---------------------------------------------------------------------}
    procedure GererRecalculSolde(Gene : string; dt : TDateTime);
    {---------------------------------------------------------------------}
    begin
      {JP 21/11/05 : FQ 10304 : SoldeOk et non PaieOk !!!!}
      if SoldeOk then begin
        k := lCalcSold.IndexOf(Gene);
        if k = -1 then begin
          Obj := TObjDtValeur.Create;
          Obj.DateVal := dt;
          lCalcSold.AddObject(Gene, Obj);
        end
        else begin
          Obj := TObjDtValeur(lCalcSold.Objects[k]);
          if Obj.DateVal > dt then
            Obj.DateVal := dt;
        end;
      end;
    end;

begin
  Result := 0;
  J := TStringList.Create;
  {24/09/04 : FQ 10104 : Pour stocker les comptes généraux collectifs/TIC/TID mais pas d'effet. Le problème
              vient de la saisie d'effets et de règlement en compta  (Unité SaisieTr.RetouchePiece) où l'on
              peut avoir E_TRESOSYNCHRO à CRE au Lieu de LET}
  G := TStringList.Create;
  {08/08/06 : Liste des comptes courants fournis depuis la Table CLIENSSOC}
  C := TStringList.Create;
  lCalcSold := TStringList.Create;
  GererListe(J, G, C, Grp.NomBase);
  ObjRub   := TObjRubrique.Create;
  Grp.oRub := ObjRub;
  ObjDtCib := TDateValeur.Create;
  Grp.oValeur := ObjDtCib;
  Grp.oValeur.RecupConditionValeur;
  ObjDev := TObjDevise.Create(V_PGI.DateEntree);
  Grp.oDevise := ObjDev;
  Grp.oBqPrev.NoDossier := Grp.NoDossier;

  {Génération du code unique pour les clefs de la table TRECRITURE}
  Grp.sChaine := GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
  Grp.iCode   := StrToInt(Grp.sChaine);

  {Création de la tob contenant les écritures qui n'auront pas été traitées}
  TobExclus := TOB.Create('$$$', nil, -1);
  try
    {Intégration dans TRECRITURE des écritures concernées par la synchronisation}
    Initmove(T.Detail.Count, TraduireMemoire('Synchronisation des écritures'));
    for n := 0 to T.Detail.Count - 1 do begin
      {Les écritures de lettrage : pour ne prendre aucun risque dans la compta, et surtout pour
       ne pas alourdir le lettrage par plusieurs requêtes, toutes les écritures modifiées lors du
       lettrage voient E_TRESOSYNCHRO passer à "LET". Il s'agit donc ici de distinguer les écritures
       de règlement dont on passera E_TRESOSYNCHRO à "RIE" de celles de factures (E_TRESOSYNCHRO à "LET").
       Remarque : le problème ne se pose pas lors de la première synchronisation car toutes les
                  écritures sont réinitialisées.
       24/09/04 : FQ 10104 : dans SaisieTr, La valeur "LET" peut avoir été remplacé par "CRE"
                  on exclut les écritures : d'un journal de banque et
                                             soit avec E_TRESOSYNCHRO à BQP
                                             soit avec E_TRESOSYNCHRO à LET
                                             soit avec E_TRESOSYNCHRO à CRE et Collectif/TIC/TID
                                                                            et pas MultiSoc pour les DIV}
      if ((J.IndexOf(T.Detail[n].GetString('E_JOURNAL')) > - 1) and
          ((T.Detail[n].GetString('E_TRESOSYNCHRO') = ets_Lettre) or
           {20/09/05 : FQ 10294 : Dans CCMP, on peut affecter la banque prévisionnelle sur un règlement,
                       alors que c'est fait pour pour les factures ou les effets => avant de rajouter ce
                       test, les comptes collectifs des règlements pouvaient donc être synchronisés  !!!!}
           (T.Detail[n].GetString('E_TRESOSYNCHRO') = ets_BqPrevi) or
           ((T.Detail[n].GetString('E_TRESOSYNCHRO') = ets_Nouveau) and
            ((T.Detail[n].GetString('G_NATUREGENE') <> 'BQE') and
             ((T.Detail[n].GetString('G_NATUREGENE') <> 'DIV') or
              ((T.Detail[n].GetString('G_NATUREGENE') = 'DIV') and not EstMultiSoc)))))) then begin
        {Mise à jour de E_TRESOSYNCHRO}
        MajTRESOSYNCHRO(T.Detail[n], Grp.NomBase, ets_Rien);
        MoveCur(False);
        Continue;
      end;

     {On exclut aussi
        1/ soit un journal d'effets et
           soit avec E_TRESOSYNCHRO à LET et compte d'effets
           soit avec E_TRESOSYNCHRO à CRE et compte d'effets
        2/ not PaieOk and E_MODEPAIE = ""}
      if ((G.IndexOf(T.Detail[n].GetString('E_JOURNAL')) > - 1) and
          ((T.Detail[n].GetString('E_TRESOSYNCHRO') = ets_Lettre) or
           (T.Detail[n].GetString('E_TRESOSYNCHRO') = ets_Nouveau)) and
          (T.Detail[n].GetString('G_EFFET') <> 'X') and
          (T.Detail[n].GetString('G_NATUREGENE') <> 'DIV')) or
         {20/07/02 : Utilisation de Length pour le cas ou le champ est à null et que GetString renvoie #0}
         ((not PaieOk and (Length(T.Detail[n].GetString('E_MODEPAIE')) = 0)) and
         {14/09/05 : Maintenant que les comptes divers sont lettrables, il n'est pas impossible qu'à la suite
                     de certains traitements, il y ait certaines écritures à CRE.
          08/08/06 : Les comptes divers peuvent être des comptes courants : on ne les synchronise pas
                     si l'on n'est pas en Multi sociétés, sinon, ils sont traités ci-dessous}
         (T.Detail[n].GetString('G_NATUREGENE') = 'DIV') and not EstMultiSoc) then begin
        {Mise à jour de E_TRESOSYNCHRO}
        MajTRESOSYNCHRO(T.Detail[n], Grp.NomBase, ets_Rien);
        MoveCur(False);
        Continue;
      end;

      {08/08/06 : Gestion des comptes courants : sont synchronisables
                  1/ les comptes courants appartenant à la table CLIENSSOC
                  2/ sur des journaux bancaires ou l'OD
                  //3/ Pour le moment, je considère qu'ils ne doivent pas être lettrables}
      if (T.Detail[n].GetString('G_NATUREGENE') = 'DIV') and
         ((C.IndexOf(T.Detail[n].GetString('E_GENERAL')) = - 1) or
          ((T.Detail[n].GetString('J_NATUREJAL') <> 'BQE') and
           (T.Detail[n].GetString('J_NATUREJAL') <> 'OD')) {or
          (T.Detail[n].GetString('E_TRESOSYNCHRO') = ets_Lettre)}) then begin
        MajTRESOSYNCHRO(T.Detail[n], Grp.NomBase, ets_Rien);
        MoveCur(False);
        Continue;
      end;


      {15/09/04 : S'il y a un filtre sur les comptes bancaires et que l'écriture previsionnelle
                  a comme banque prévi ce compte, on ne la traite pas}
      if Grp.oBqPrev.IsFiltre(T.Detail[n].GetString('E_BANQUEPREVI'), T.Detail[n].GetString('G_NATUREGENE')) then Continue;

      {Sinon, on importe l'écriture}
      s := TresoFromCompta(T.Detail[n], Grp, SoldeOk); {JP 25/07/05 : Nouvelle gestion du recalcul des soldes}

      {Il y a un retour si le code ETABBQ n'est pas récupérable}
      if s <> '' then
        AjouterRefus(TobExclus, s) {FQ 10144}
      else begin
        GererRecalculSolde(Grp.sGene, Grp.dDate);
        Inc(Result);
      end;
      MoveCur(False);
    end;
    FiniMove;

    {Gestion des écritures de trésorerie modfiée en comptabilité par le pointage ou le rapprochement}
    Initmove(T.Detail.Count, TraduireMemoire('Mise à jour des dates de rapprochement'));
    for n := 0 to R.Detail.Count - 1 do begin
      UpdateRapprochement(R.Detail[n], Grp.NomBase, Grp.NoDossier);
      {En faisant le recalcul des soldes sur E_DATECOMPTABLE, il s'agit d'une approximation !!}
      GererRecalculSolde(GetBqCodeFromGene(R.Detail[n].GetString('E_GENERAL'), Grp.NoDossier),
                         R.Detail[n].GetDateTime('E_DATECOMPTABLE'));
      MoveCur(True);
    end;
    FiniMove;

    Result := Result + R.Detail.Count;

    TermineUnDossier(lCalcSold, TobExclus, Grp);
  finally
    LibereListe(lCalcSold, True);
    if Assigned(TobExclus) then FreeAndNil(TobExclus);
    if Assigned(G        ) then FreeAndNil(G);
    if Assigned(J        ) then FreeAndNil(J);
    if Assigned(C        ) then FreeAndNil(C);
    if Assigned(ObjRub   ) then FreeAndNil(ObjRub);
    if Assigned(ObjDev   ) then FreeAndNil(ObjDev);
    if Assigned(ObjDtCib ) then FreeAndNil(ObjDtCib);
    {Mise à jour des compteurs pour les clefs de la table TRECRITURE}
    if StrToInt(Grp.sChaine) < Grp.iCode then begin
      {On décremente iCode car il a été incrémenté après la dernière écriture}
      Dec(Grp.iCode);
      {Maj du Code unique dans la table COMPTEURTRESO}
      SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, PadL(IntToStr(Grp.iCode), '0', 7));
    end;
  end
end;

{---------------------------------------------------------------------------------------}
procedure MajTRESOSYNCHRO(T : TOB; NomBase : string; Value : string = ''; dtValeur : TDateTime = 0);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  if Value = '' then Value := ets_Synchro;
  {Mise à jour du champ E_TRESOSYNCHRO.}
  SQL := 'UPDATE ' + GetTableDossier(NomBase, 'ECRITURE') + ' SET E_TRESOSYNCHRO = "' + Value + '", ' +
         'E_DATEMODIF = "' + USDateTime(Now) + '" ';
  {On ne modifie la date de valeur que s'il s'agit d'une écriture de règlement non pointée}
  if dtValeur > 0 then
    SQL := SQL + ', E_DATEVALEUR = "' + USDateTime(dtValeur) + '" ';

  SQL := SQL + 'WHERE E_EXERCICE = "' + T.GetString('E_EXERCICE') +
               '" AND E_JOURNAL = "' + T.GetString('E_JOURNAL') +
               '" AND E_DATECOMPTABLE = "' + USDateTime(T.GetDateTime('E_DATECOMPTABLE')) +
               '" AND E_NUMEROPIECE = ' + T.GetString('E_NUMEROPIECE') +
               ' AND E_NUMLIGNE = ' + T.GetString('E_NUMLIGNE') +
               ' AND E_NUMECHE = ' + T.GetString('E_NUMECHE') +
               ' AND E_QUALIFPIECE = "N"';
  ExecuteSQL(SQL);
end;

{JP 04/05/04 : FQ 10062 : gestion du sens : maintenant un même compte pourra être rattaché
               à deux rubriques en fonction de son sens (positif ou négatif)
{---------------------------------------------------------------------------------------}
procedure TravaillerFlux;
{---------------------------------------------------------------------------------------}
var
  Lib : string;
  Rub : string;
  Gen : string;
  Pai : string;
  SQL : string;
  Obj : TObjRubrique;
  Q   : TQuery;
  n   : Integer;
begin
  Obj := TObjRubrique.Create;
  try
  {27/09/04 : FQ 10119 : Gestion des modes de paiement dans les rubriques : on rajoute
              TE_USERMODIF (dans lequel est stocké le mode de paiement) dans la requête
   09/08/06 : Gestion MS : on traite les comptes courants à part}
    Q := OpenSQL('SELECT DISTINCT TE_CONTREPARTIETR GENERAL, TE_USERMODIF AMODE FROM TRECRITURE ' +
                 'LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL ' +
                 'WHERE TE_CODEFLUX = "' + CODEIMPORT + '" AND BQ_BANQUE NOT IN ("' + CODECOURANTS + '", "' +
                 CODETITRES + '") ORDER BY TE_USERMODIF DESC', True);
    Obj.Generaux.LoadDetailDB('***', '', '', Q, False);
    Ferme(Q);
    {Construction de la liste de correspondance Comptes généraux / Rubriques}
    Obj.SetListeCorrespondance;
    {27/09/04 : la TOB est Triée par mode de paiement. On commence par les modes de paiement
                non-vides => on traite les cas particuliers avant le cas général}
    for n := 0 to Obj.Generaux.Detail.Count - 1 do begin
      Gen := Obj.Generaux.Detail[n].GetString('GENERAL');
      Pai := Obj.Generaux.Detail[n].GetString('AMODE'); {27/09/04}
      {Récupération de la rubrique rattachée au compte de contrepartie et au sens de l'écriture}
      Obj.GetCorrespondance(Gen, Rub, Lib, 'C', Pai);

      {JP 04/03/04 : Si la rubrique n'a pas été paramétrée, on laisse IMP pour pouvoir relancer
                     le traitement (FQ 10017)}
      if (Trim(Rub) <> '') then begin
        SQL := 'UPDATE TRECRITURE SET TE_CODEFLUX = "' + Rub + '" WHERE TE_CONTREPARTIETR = "' + Gen +
               '" AND TE_CODEFLUX = "' + CODEIMPORT + '" AND TE_MONTANT >= 0';
        if Trim(Pai) <> '' then {27/09/04 : si le mode de paiment n'est pas vide}
          SQL :=  SQL + ' AND TE_USERMODIF = "' + Pai + '"';
        ExecuteSQL(SQL);
      end;

      {Récupération de la rubrique rattachée au compte de contrepartie d'une écriture négative}
      Obj.GetCorrespondance(Gen, Rub, Lib, 'D', Pai);
      {JP 04/03/04 : Si la rubrique n'a pas été paramétrée, on laisse IMP pour pouvoir relancer
                     le traitement (FQ 10017)}
      if (Trim(Rub) = '') then Continue;
      SQL := 'UPDATE TRECRITURE SET TE_CODEFLUX = "' + Rub + '" WHERE TE_CONTREPARTIETR = "' + Gen +
             '" AND TE_CODEFLUX = "' + CODEIMPORT + '" AND TE_MONTANT < 0';
        if Trim(Pai) <> '' then {27/09/04 : si le mode de paiment n'est pas vide}
          SQL :=  SQL + ' AND TE_USERMODIF = "' + Pai + '"';
      ExecuteSQL(SQL);
    end;

    {09/08/06 : Gestion des comptes courants et titres}
    Q := OpenSQL('SELECT DISTINCT TE_GENERAL CODEGEN, BQ_GENERAL GENERAL, BQ_CODECIB AMODE FROM TRECRITURE ' +
                 'LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL ' +
                 {05/04/07 : FQ 10428 : BQ_BANQUE est mieux que TE_CODEBANQUE}
                 'WHERE TE_CODEFLUX = "' + CODEIMPORT + '" AND BQ_BANQUE IN ("' + CODECOURANTS + '", "' +
                 CODETITRES + '") ORDER BY BQ_GENERAL', True);
    Obj.Generaux.LoadDetailDB('***', '', '', Q, False);
    Ferme(Q);
    {Construction de la liste de correspondance Comptes généraux / Rubriques}
    Obj.SetListeCorrespondance;
    for n := 0 to Obj.Generaux.Detail.Count - 1 do begin
      Gen := Obj.Generaux.Detail[n].GetString('GENERAL');
      Pai := Obj.Generaux.Detail[n].GetString('AMODE'); {27/09/04}

      {Traitement des écritures négatives}
      Obj.GetCorrespondanceMS(Gen, Rub, Lib, 'D', Pai);
      if (Trim(Rub) <> '') then begin
        SQL := 'UPDATE TRECRITURE SET TE_CODEFLUX = "' + Rub + '" WHERE TE_GENERAL = "' +
               Obj.Generaux.Detail[n].GetString('CODEGEN') +
               '" AND TE_CODEFLUX = "' + CODEIMPORT + '" AND TE_MONTANT < 0';
        ExecuteSQL(SQL);
      end;

      {Traitement des écritures négatives}
      Obj.GetCorrespondanceMS(Gen, Rub, Lib, 'C', Pai);
      if (Trim(Rub) <> '') then begin
        SQL := 'UPDATE TRECRITURE SET TE_CODEFLUX = "' + Rub + '" WHERE TE_GENERAL = "' +
               Obj.Generaux.Detail[n].GetString('CODEGEN') +
               '" AND TE_CODEFLUX = "' + CODEIMPORT + '" AND TE_MONTANT >= 0';
        ExecuteSQL(SQL);
      end;

    end;
  finally
    FreeAndNil(Obj);
  end;
end;

{---------------------------------------------------------------------------------------}
function MajDeviseEtValeur(Obj : TDateValeur; Clause, ChpUser : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  Q : TQuery;
  n : Integer;
  D : TDateTime;
  DC : TDateTime;
  C : string;
  ODt : TObjDevise;
  MtD : Double;
  MtE : Double;
  Dev : string;
  Cpt : string;
  CVa : string;
  CCo : string;
  Sen : string;
  Per : Integer; {17/11/06 : Gestion des bordereaux}
  CodeUnq  : string;
  Roquette : string;
  Societe  : string;
begin
  Result := True;
  BeginTrans;
  try
    {Si Obj est passé en paramètre, on utilise les fonctions de l'objet}
    if Assigned(Obj) then begin
      {JP 22/03/04 : Les chaines commençant par "@@" ne sont pas compatible Oracle}
      if Trim(Clause) <> '' then Clause := Clause  + ' AND ' + ChpUser + ' = "' + CODETEMPO + '"'
                            else Clause := 'WHERE ' + ChpUser + ' = "' + CODETEMPO + '"';
      {18/12/06 : Impossible de me souvenir pourquoi j'ai posé cette condition, mais ce n'est pas bon !!! 
      if EstMultiSoc then CodeUnq := ''
                     else} CodeUnq := GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
      T := TOB.Create('***', nil, - 1);
      Q := OpenSQL('SELECT * FROM TRECRITURE ' + Clause, True);
      try
        if not Q.EOF then begin
          ODt := TObjDevise.Create(V_PGI.DateEntree);
          InitMove(Q.RecordCount, 'Travail des devises et des dates de valeur');
          try
            T.LoadDetailDB('TRECRITURE', '', '', Q, True);
            Ferme(Q);
            T.Detail.Sort('TE_SOCIETE;TE_GENERAL;TE_CODECIB;TE_DATECOMPTABLE');
            Societe := '';
            for n := 0 to T.Detail.Count - 1 do begin

              MoveCur(False);
              {Recherche du sens}
              MtE := T.Detail[n].GetDouble('TE_MONTANTDEV');
              if MtE >= 0 then Sen := 'ENC'
                          else Sen := 'DEC';
              {14/09/04 : FQ 10139 : on mémorise la date pour la mise à jour de la table ÉCRITURE}
              DC := T.Detail[n].GetValue('TE_DATECOMPTABLE');

              {Pour les écritures prévisionnelles, il faut commencer par récupérer la véritable date d'opération qui
               est stockée dans TE_DATEVALID, notamment pour que la clause de la requête ci-dessus n'exclut par
               certaines écritures si le filtre se fait sur la date comptable cf ImportTreso.TravaillerCompta}
               if T.Detail[n].GetValue('TE_NATURE') = na_Prevision then
                 T.Detail[n].SetDateTime('TE_DATECOMPTABLE', T.Detail[n].GetDateTime('TE_DATEVALID'));

              D := T.Detail[n].GetValue('TE_DATECOMPTABLE');

              {14/09/05 : FQ 10269 : Réinitialisation de C pour le test ci-dessous}
              c := '';

              {On va estimer la date de valeur de l'écriture}
              Obj.GetDateValeur(T.Detail[n].GetString('TE_GENERAL'), T.Detail[n].GetString('TE_CODECIB'), Sen, D, C);

              {14/09/05 : FQ 10269 : On ne réinitialise pas les champs si on ne récupère pas de cib afin de
                          pouvoir relancer le traitement}
              if C <> '' then begin
                T.Detail[n].PutValue(ChpUser, V_PGI.User);
                T.Detail[n].PutValue('TE_CODECIB', C);
                {Jusqu'ici TE_USERVALID a servi de champ "SENS", on le met maintenant à vide}
                T.Detail[n].PutValue('TE_USERVALID', '');
              end;

              {13/09/06 : Si l'écriture est pointée, théoriquement la date de valeur est renseignée
                          grace à l'écriture comptable}
              if not((T.GetDateTime('TE_DATERAPPRO') > iDate1900) and
                     (T.Detail[n].GetString('TE_NATURE') = na_Realise) and
                     (T.GetDateTime('TE_DATEVALEUR') > iDate1900)) then
                T.Detail[n].PutValue('TE_DATEVALEUR', D);

              {Génération des clefs de valeur et d'opération}
              CVa := RetourneCleEcriture(D, StrToInt(CodeUnq));
              CCo := RetourneCleEcriture(T.Detail[n].GetDateTime('TE_DATECOMPTABLE'), StrToInt(CodeUnq));
              T.Detail[n].PutValue('TE_CLEVALEUR'   , CVa);
              T.Detail[n].PutValue('TE_CLEOPERATION', CCo);

              {Gestion des devises : dans la compta, il est possible de saisir dans une autre devise que celle du compte.
               Pour le moment en tréso on ne met dans la table TRECITURE que les montants en € et dans la devise du compte.}
              Dev := T.Detail[n].GetString('TE_DEVISE');
              Cpt := T.Detail[n].GetString('TE_GENERAL');
              {Si la devise a été saisie dans la devise du compte, il n'y a aucun traitement à faire}
              if Dev <> ODt.GetDeviseCpt(Cpt) then begin
                {JP 28/04/04 : pour limiter les risques d'erreur et les écarts de change, on traite à part
                               le cas du compte de destination en euros}
                if ODt.GetDeviseCpt(Cpt) = V_PGI.DevisePivot then begin
                  MtD := T.Detail[n].GetDouble('TE_MONTANT');
                  MtE := T.Detail[n].GetDouble('TE_MONTANT');
                end else begin
                  {On part du montant en devise car le montant est déjà un montant calculé}
                  MtE := T.Detail[n].GetDouble('TE_MONTANTDEV');
                  ODt.ConvertitMnt(MtE, MtD, Dev, Cpt, D);
                end;
                {Mais on ne change que le montant en devise, on garde le montant en Euro calculé en comptant}
                T.Detail[n].PutValue('TE_MONTANTDEV', MtD);
                T.Detail[n].PutValue('TE_DEVISE', ODt.GetDeviseCpt(Cpt));
                if MtE <> 0 then T.Detail[n].PutValue('TE_COTATION', Arrondi(MtD / MtE, NBDECIMALTAUX))
                            else T.Detail[n].PutValue('TE_COTATION', 1);
              end;
              {Si tout s'est bien passé, on incrémente le CodeUnique}
              CodeUnq := IntToStr(StrToInt(CodeUnq) + 1);
              {14/09/04 : FQ 10139 : Mise à jour de E_TRESOSYNCHRO et de E_DATEVALEUR pour donner un peu plus de sens
                          à la règle d'accrochage sur la date de valeur lors du rapprochement bancaire}
              {08/08/06 : Gestion du multi sociétés}
              if (T.Detail[n].GetString('TE_NODOSSIER') = V_PGI.NoDossier) or not EstMultiSoc then
                Roquette := 'UPDATE ECRITURE SET E_TRESOSYNCHRO = "' + ets_Synchro + '" '
              else begin
                Roquette := GetInfosFromDossier('DOS_NODOSSIER', T.Detail[n].GetString('TE_NODOSSIER'), 'DOS_NOMBASE');
                Roquette := 'UPDATE ' + GetTableDossier(Roquette, 'ECRITURE') + ' SET E_TRESOSYNCHRO = "' + ets_Synchro + '" ';
              end;

              {17/11/06 : Recherche de la valeur de E_PERIODE}
              Per := GetNumPeriodeTransac(T.Detail[n].GetString('TE_NUMTRANSAC'), T.Detail[n].GetString('TE_JOURNAL'));

              Roquette := Roquette + ', E_DATEMODIF = "' + USDateTime(Now) + '" ';
              {On ne modifie la date de valeur que s'il s'agit d'une écriture de règlement non pointée}
              if (T.Detail[n].GetString('TE_NATURE') = na_Realise) and (T.Detail[n].GetDateTime('TE_DATERAPPRO') = iDate1900) then
                Roquette := Roquette + ', E_DATEVALEUR = "' + USDateTime(D) + '" ';
              Roquette := Roquette + 'WHERE' +
                         ' E_JOURNAL = "'       + T.Detail[n].GetString('TE_JOURNAL')     + '" AND' +
                         ' E_EXERCICE = "'      + T.Detail[n].GetString('TE_EXERCICE')    + '" AND' +
                         ' E_DATECOMPTABLE = "' + UsDateTime(DC) + '" AND' +
                         ' E_NUMEROPIECE = '    + T.Detail[n].GetString('TE_NUMEROPIECE') + ' AND'  +
                         ' E_NUMLIGNE = '       + T.Detail[n].GetString('TE_CPNUMLIGNE')  + ' AND'  +
                         ' E_NUMECHE = '        + T.Detail[n].GetString('TE_NUMECHE')     + ' AND';

              if Per > -1 then
                Roquette := Roquette + ' E_PERIODE = ' + IntToStr(Per) + ' AND';
              Roquette := Roquette + ' E_QUALIFPIECE = "N"';
              ExecuteSql(Roquette);
            end;
            T.UpdateDB;
          finally
            FreeAndNil(ODt);
            FiniMove;
          end;
        end
        else
          Ferme(Q);
      finally
        SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, CodeUnq);
        if Assigned(T) then FreeAndNil(T);
      end;
    end
    {Sinon, on met des valeurs par défaut}
    else begin

    end;
    CommitTrans;
  except
    on E : Exception do begin
      PGIError('Une erreur est intervenue lors de la mise à jours des dates de valeur et des devise avec le message :'#13 +
               E.Message + '.'#13#13 +
               'La synchronisation ne s''est pas correctement effectuée et la table ECRITURE n''a pas été mise à jour.');
      RollBack;
    end;
  end;
end;

{Calculs des dates de valeur, des codes CIB, des devises et des clefs
{---------------------------------------------------------------------------------------}
procedure TravaillerDates(ClauseWh : string);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TQuery;
  Obj : TDateValeur;
begin
  {22/03/04 : Les chaines commençant par "@@" ne sont pas compatible Oracle
   31/05/06 : Ajout d'un préfixe aux alias pour éviter une erreur en multi-sociétés}
  SQL := 'SELECT TE_CODECIB AAA_MODEPAIE, TE_GENERAL AAA_GENERAL, TE_USERVALID AAA_SENS, BQ_BANQUE AAA_BANQUE FROM TRECRITURE ' +
         'LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL ' +
         'WHERE TE_USERCREATION = "' + CODETEMPO + '" ' +
         'GROUP BY BQ_BANQUE, TE_CODECIB, TE_GENERAL, TE_USERVALID';
  Q := OpenSQL(SQL, True);
  try
    if not Q.EOF then begin
      Obj := TDateValeur.Create;
      try
        Obj.ModePaie.LoadDetailDB('', '', '', Q, False);
        Obj.RecupConditionValeur;
        MajDeviseEtValeur(Obj, ClauseWh, 'TE_USERCREATION')
      finally
        FreeAndNil(Obj);
      end;
    end;
  finally
    Ferme(Q);
  end;
end;

{10/08/06 : Pour la gestion des rejets et le recalcul des soldes
{---------------------------------------------------------------------------------------}
procedure TermineUnDossier(L : TStringList; var T : TOB; G : TGrpObjet);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  o : TObjDtValeur;
  v : TObjDtValeur;
begin
  for n := T.Detail.Count - 1 downto 0 do
    T.Detail[n].ChangeParent(G.TobRejet, -1);

  for n := 0 to L.Count - 1 do begin
    {On regarde si le compte en cours figure dans la liste ...}
    p := G.ListSolde.IndexOf(L[n]);
    {... Si oui ...}
    if p > -1  then begin
      o := TObjDtValeur(L.Objects[n]);
      v := TObjDtValeur(G.ListSolde.Objects[p]);
      {on compare les dates : on mémorise la plus ancienne pour le recalcul des soldes}
      if v.DateVal > o.DateVal then v.DateVal := o.DateVal;
    end
    {... si non, on ajoute le compte et l'objet dans la liste}
    else begin
      o := TObjDtValeur.Create;
      o.DateVal := TObjDtValeur(L.Objects[n]).DateVal;
      G.ListSolde.AddObject(L[n], o);
    end;
  end;
end;

end.
