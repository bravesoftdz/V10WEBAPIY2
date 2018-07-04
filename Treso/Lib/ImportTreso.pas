{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 06/10/2003
Modifié le ... : 15/01/2004
Description .. : contenant les fonctions d'import et de formatage d'écriture
Suite ........ : dans la tréso
Mots clefs ... : TRAVAILLERDATES;UPDATERAPPROCHEMENT;TRAVAILLERDONNEES;TRAVAILLERCOMPTA
*****************************************************************}
{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            06/10/03    JP     Création de l'unité
1.2X.001.001     25/03/04    JP     Mise en place de la synchronisation de l'échéancier.
--------------------------------------------------------------------------------------}
unit ImportTreso;

interface

uses
  {$IFNDEF EAGLCLIENT}
  DBTables,
  {$ENDIF}
  Commun, UProcGen, Constantes, UObjGen, HEnt1, HCtrls, Classes, UTob, SysUtils, Controls;


{Génère la ligne treso en fonction de la ligne de compta}
function  TravaillerCompta(T : TOB; var UneEcriture : TREcritures; Obj : TObjBanquePrevi) : string;
{Traitements sur les données intégrées dans la tables TRECRITURE : dates de valeur}
procedure TravaillerDonnees(ClauseWh : string);
{Maj des écritures de Treso intégrées en compta et modifiées lors du rapprochement bancaire}
procedure UpdateRapprochement(T : TOB);

procedure TravaillerDates(ClauseWh : string);
procedure TravaillerFlux;

implementation

var
  ListeCompte   : TStringList;
  ListeDateDep  : TStringList;

{---------------------------------------------------------------------------------------}
procedure UpdateRapprochement(T : TOB);
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  NLI : string;
  TSC : string;
  SQL : string;
  s   : string;
  d   : TDate;
begin
  for n := 0 to T.Detail.Count - 1 do begin
      s   := T.Detail[n].GetValue('E_REFINTERNE');
      NLI := Copy(s, Pos('-', s) + 1, 1);
      TSC := Copy(s, 1, Pos('-', s) - 1);
      d   := StrToDate(VarToStr(T.Detail[n].GetValue('E_DATEPOINTAGE')));
    {Mise à jour de la date de pointage}
    SQL := 'UPDATE TRECRITURE SET TE_DATERAPPRO = "' + USDateTime(d) +
           '" WHERE TE_NUMTRANSAC = "' + TSC + '" AND TE_NUMLIGNE = ' + NLI;
    ExecuteSQL(SQL);
    {Mise à jour du champ E_TRESOSYNCHRO.}
    SQL := 'UPDATE ECRITURE SET E_TRESOSYNCHRO = "SYN" WHERE E_EXERCICE = "' + T.Detail[n].GetValue('E_EXERCICE') +
           '" AND E_JOURNAL = "' + T.Detail[n].GetValue('E_JOURNAL') +
           '" AND E_DATECOMPTABLE = "' + USDateTime(StrToDate(VarToStr(T.Detail[n].GetValue('E_DATECOMPTABLE')))) +
           '" AND E_QUALIFPIECE = "' + T.Detail[n].GetValue('E_QUALIFPIECE') +
           '" AND E_NUMEROPIECE = ' + VarToStr(T.Detail[n].GetValue('E_NUMEROPIECE')) +
           ' AND E_NUMLIGNE = ' + VarToStr(T.Detail[n].GetValue('E_NUMLIGNE')) +
           ' AND E_NUMECHE = ' + VarToStr(T.Detail[n].GetValue('E_NUMECHE'));
    ExecuteSQL(SQL);
  end;
end;

{---------------------------------------------------------------------------------------}
function TravaillerCompta(T : TOB; var UneEcriture : TREcritures; Obj : TObjBanquePrevi) : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
  Nouveau : Boolean;

    {-----------------------------------------------------------------}
    procedure MajTresoSynchro;
    {-----------------------------------------------------------------}
    begin
      S := 'UPDATE ECRITURE SET E_TRESOSYNCHRO = "SYN", E_DATEMODIF = "' + USDateTime(Now) + '" ';
      S := S + 'WHERE' +
               ' E_JOURNAL = "'       + T.GetString('E_JOURNAL')     + '" AND' +
               ' E_EXERCICE = "'      + T.GetString('E_EXERCICE')    + '" AND' +
               ' E_DATECOMPTABLE = "' + UsDateTime(T.GetDateTime('E_DATECOMPTABLE')) + '" AND' +
               ' E_NUMEROPIECE = '    + T.GetString('E_NUMEROPIECE') + ' AND'  +
               ' E_NUMLIGNE = '       + T.GetString('E_NUMLIGNE')    + ' AND'  +
               ' E_NUMECHE = '        + T.GetString('E_NUMECHE')     + ' AND'  +
               ' E_QUALIFPIECE = "N"';
      ExecuteSql(S);
    end;

begin
  Result := '';
  {On ne traite pas les écritures prévisionnelles que l'on ne peut rattacher à aucune banque}
  if (T.GetString('E_BANQUEPREVI') = '') and not Obj.IsValide(Obj.NatToIndice(T.GetString('G_NATUREGENE'))) then begin
    Result := T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Banque prévisionnelle non renseignée') + ';';
    Exit;
  end;

  UneEcriture.TE_NUMTRANSAC     := TRANSACIMPORT + V_PGI.CodeSociete + T.GetString('E_JOURNAL') + IntToStr(T.GetInteger('E_NUMEROPIECE'));
  UneEcriture.TE_CODEFLUX       := CODEIMPORT; {JP 04/03/04}
  UneEcriture.TE_NUMLIGNE       := T.GetInteger('E_NUMLIGNE') + T.GetInteger('E_NUMECHE');
  UneEcriture.TE_NUMECHE        := T.GetInteger('E_NUMECHE');
  UneEcriture.TE_CPNUMLIGNE     := T.GetInteger('E_NUMLIGNE');
  UneEcriture.TE_CODECIB        := T.GetString('E_MODEPAIE');
  UneEcriture.TE_JOURNAL        := T.GetString('E_JOURNAL');
  UneEcriture.TE_NUMEROPIECE    := T.GetInteger('E_NUMEROPIECE');
  UneEcriture.TE_EXERCICE       := T.GetString('E_EXERCICE');
  UneEcriture.TE_REFINTERNE     := T.GetString('E_REFINTERNE');
  UneEcriture.TE_SOCIETE        := T.GetString('E_SOCIETE');
  UneEcriture.TE_ETABLISSEMENT  := T.GetString('E_ETABLISSEMENT');
  UneEcriture.TE_QUALIFORIGINE  := QUALIFCOMPTA;
  UneEcriture.TE_LIBELLE        := T.GetString('E_LIBELLE');
  {JP 22/03/04 : Les chaines commençant par "@@" ne sont pas compatible Oracle}
  UneEcriture.TE_USERCREATION   := CODETEMPO;
  {27/09/04 : FQ 10119 : Gestion des modes de paiement dans les rubriques}
  UneEcriture.TE_USERMODIF      := T.GetString('E_MODEPAIE');
  UneEcriture.TE_USERCOMPTABLE  := V_PGI.User;
  UneEcriture.TE_DEVISE         := T.GetString('E_DEVISE');
  {E_Couverture étant un montant en valeur absolu ...}
  if (T.GetDouble('E_DEBIT') - T.GetDouble('E_CREDIT')) >= 0 then begin
    UneEcriture.TE_MONTANT        := T.GetDouble('E_DEBIT') - T.GetDouble('E_CREDIT') - T.GetDouble('E_COUVERTURE');
    UneEcriture.TE_MONTANTDEV     := T.GetDouble('E_DEBITDEV') - T.GetDouble('E_CREDITDEV') - T.GetDouble('E_COUVERTUREDEV');
  end else begin
    UneEcriture.TE_MONTANT        := T.GetDouble('E_DEBIT') - T.GetDouble('E_CREDIT') + T.GetDouble('E_COUVERTURE');
    UneEcriture.TE_MONTANTDEV     := T.GetDouble('E_DEBITDEV') - T.GetDouble('E_CREDITDEV') + T.GetDouble('E_COUVERTUREDEV');
  end;
  {La cotation en compta semble être l'inverse de la tréso et de la chancellerie.
   Pour plus de sécurité, on la recalcule, même si le champ est rarement utilisé !
   Dans les tables Chancell et Trecriture MntDev := Cotation * Mnt€}
  UneEcriture.TE_COTATION       := 1; {Par défaut}
  if UneEcriture.TE_MONTANT <> 0 then
    UneEcriture.TE_COTATION     := UneEcriture.TE_MONTANTDEV / UneEcriture.TE_MONTANT;
  { JP 09/04/04 : Voilà quelque chose qui doit être un oubli, car la gestion en compta est différente
                  de celle de la tréso : dans cette dernière, l'écriture est forcément dans la devise
                  du compte, alors qu'en compta la devise est au choix en fonction du journal de la pièce
  UneEcriture.TE_COTATION       := T.GetDouble('E_COTATION');}
  UneEcriture.TE_DATERAPPRO     := T.GetString('E_DATEPOINTAGE');

  {Dans le cas des écritures prévisionnelles, la date d'opération n'est pas la date comptable mais la date
   d'échéance. Dans un premier temps, on stocke la date d'échéance dans TE_DATEVALID et on mettra à jour
   TE_DATECOMPTABLE (date d'opération) dans le travail sur les dates de valeur : ceci pour une raison, à
   savoir que les critères de sélection reposent, entre autre, sur E_DATECOMPTABLE / TE_DATECOMPTABLE, donc
   si TE_DATECOMPTABLE était différent de E_DATECOMPTABLE, certains traitements ne seraient pas correctement
   effectués, notamment le calcul des dates de valeur car ceratines écritures sont exclues par le filtrage sur
   la date d'opération : cf UObjGen.MajDeviseEtValeur }
  UneEcriture.TE_DATECOMPTABLE  := T.GetString('E_DATECOMPTABLE');
  {Ecritures bancaires ...}
  if Obj.NatToIndice(T.GetString('G_NATUREGENE')) = nc_BqCais then begin
    UneEcriture.TE_GENERAL        := T.GetString('E_GENERAL');
    UneEcriture.TE_CONTREPARTIETR := T.GetString('E_CONTREPARTIEGEN');
    UneEcriture.TE_NATURE         := na_Realise;
    UneEcriture.TE_DATEVALID      := DateToStr(iDate1900);
  end

  {... ou échéancier}
  else begin
    if T.GetString('E_BANQUEPREVI') <> '' then
      UneEcriture.TE_GENERAL      := T.GetString('E_BANQUEPREVI')
    else
      UneEcriture.TE_GENERAL      := Obj.GetCompte(Obj.NatToIndice(T.GetString('G_NATUREGENE')));

    UneEcriture.TE_CONTREPARTIETR := T.GetString('E_GENERAL');
    UneEcriture.TE_NATURE         := na_Prevision;
    UneEcriture.TE_DATEVALID      := T.GetString('E_DATEECHEANCE');
  end;

  UneEcriture.TE_CODEBANQUE       := Obj.GetEtabBq (UneEcriture.TE_GENERAL);
  UneEcriture.TE_CODEGUICHET      := Obj.GetGuichet(UneEcriture.TE_GENERAL);
  UneEcriture.TE_NUMCOMPTE        := Obj.GetNumCpt (UneEcriture.TE_GENERAL);
  UneEcriture.TE_CLERIB           := Obj.GetCleRib (UneEcriture.TE_GENERAL);
  UneEcriture.TE_IBAN             := Obj.GetIban   (UneEcriture.TE_GENERAL);

  {Le code banque est nécessaire en tréso => on n'intègre pas la ligne}
  if (UneEcriture.TE_CODEBANQUE = '') then begin
    Result := T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Code banque non renseigné') + ';';
    Exit;
  end
  else if (UneEcriture.TE_CODECIB = '') then begin
    Result := T.GetString('E_GENERAL') + ';' + T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' + TraduireMemoire('Mode de paiement non renseigné') + ';';
    Exit;
  end;

  UneEcriture.TE_DATECREATION := DateToStr(V_PGI.DateEntree);
  UneEcriture.TE_DATEMODIF    := DateToStr(iDate1900);
  UneEcriture.TE_DATEVALEUR   := DateToStr(iDate1900);

  {Le champ TE_USERVALID va servir temporairement de Champ SENS}
  if UneEcriture.TE_MONTANTDEV >= 0 then UneEcriture.TE_USERVALID := 'ENC'
                                    else UneEcriture.TE_USERVALID := 'DEC';

  {Si le montant est égal à zéro (on peut supposer que l'écriture vient d'être lettrée) et qu'il s'agit d'une
   écriture de l'échéancier ou bien si l'écriture est complètement lettrée, on la supprime de la trésorerie.
   On fait aussi le test sur le montant en devise pour contourner déventuels problèmes d'arrondi dans le cas
   d'un lettrage en devise}
  if ((UneEcriture.TE_MONTANT = 0) and (Obj.NatToIndice(T.GetString('G_NATUREGENE')) <> nc_BqCais)) or
     ((UneEcriture.TE_MONTANTDEV = 0) and (Obj.NatToIndice(T.GetString('G_NATUREGENE')) <> nc_BqCais)) or
     (T.GetString('E_ETATLETTRAGE') = 'TL') then begin
    ExecuteSQL('DELETE FROM TRECRITURE WHERE TE_NUMTRANSAC = "' + UneEcriture.TE_NUMTRANSAC + '" ' +
                'AND TE_NUMLIGNE = ' + IntToStr(UneEcriture.TE_NUMLIGNE));

    {05/10/04 : Il faut mettre à jour la table ECRITURE ici, car si la TRECRITURE est supprimée, le traitement
                dans UObjGen.MajDeviseEtValeur ne sera pas lancé et E_TRESOSYNCHRO restera à "LET"}
    MajTresoSynchro;
  end
  else begin
    {L'écriture comptable a pu être crée puis modifiée avant d'avoir été synchronisée : on s'assure
     qu'elle existe bien}
    s := 'SELECT TE_NUMTRANSAC FROM TRECRITURE WHERE TE_NUMTRANSAC = "' + UneEcriture.TE_NUMTRANSAC + '" ' +
         'AND TE_NUMLIGNE     = '  + IntToStr(UneEcriture.TE_NUMLIGNE);

    Nouveau := (T.GetValue('E_SYNCHROTRESO') = 'CRE');

    if not Nouveau then begin
      Nouveau := not ExisteSQL(s);
      {05/10/04 : Si l'écriture existe et que l'on passe par ici, c'est qu'on est soit en lettrage partiel soit
                  en modification d'écriture bancaire : a priori TE_USERCREATION n'est plus à CODETEMPO, donc
                  comme ci-dessus, UObjGen.MajDeviseEtValeur ne sera pas éxécuté sur l'écriture courante}
      if not Nouveau then
        MajTresoSynchro;
    end;
    {S'il s'agit d'une nouvelle écriture ou d'une écriture modifiée en compta avant d'avoir été synchronisée}
    EcritTREcriture(UneEcriture, Nouveau);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure RecalculerSoldes;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  {Recalcul des soldes par compte bancaire}
  for n := 0 to ListeCompte.Count - 1 do
    RecalculSolde(ListeCompte[n], ListeDateDep[n], 0, True);
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
  {JP 22/03/04 : Les chaines commençant par "@@" ne sont pas compatible Oracle}
  SQL := 'SELECT TE_CODECIB MODEPAIE, TE_GENERAL GENERAL, TE_USERVALID SENS, BQ_BANQUE BANQUE FROM TRECRITURE ' +
         'LEFT JOIN BANQUECP ON BQ_GENERAL = TE_GENERAL ' +
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
              TE_USERMODIF (dans lequel est stocké le mode de paiement) dans la requête}
    Q := OpenSQL('SELECT DISTINCT TE_CONTREPARTIETR GENERAL, TE_USERMODIF AMODE FROM TRECRITURE WHERE ' +
                 'TE_CODEFLUX = "' + CODEIMPORT + '" GROUP BY TE_CONTREPARTIETR, TE_USERMODIF ' +
                 'ORDER BY TE_USERMODIF DESC', True);
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
  finally
    FreeAndNil(Obj);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TravaillerDonnees(ClauseWh : string);
{---------------------------------------------------------------------------------------}
var
  Cpt : TQuery;
begin
  {Constitution des listes des comptes généraux et des dates de départ pour le recalcul des soldes}
  ListeCompte  := TStringList.Create;
  ListeDateDep := TStringList.Create;
  try
    Cpt := OpenSql('SELECT DISTINCT TE_GENERAL , MIN (TE_DATECOMPTABLE) FROM TRECRITURE WHERE TE_CODEFLUX = "' +
                   CODEIMPORT + '" GROUP BY TE_GENERAL', True);
    while not Cpt.EOF do begin
      ListeCompte .Add(Cpt.Fields[0].AsString);
      ListeDateDep.Add(Cpt.Fields[1].AsString);
      Cpt.Next;
    end;
    Ferme(Cpt);
    {Gestion de la date de valeur/ codecib}
    TravaillerDates(ClauseWh);
    {Mise à jour des codes flux}
    TravaillerFlux;
    {Recalcul des soldes}
    RecalculerSoldes;
  finally
    FreeAndNil(ListeCompte);
    FreeAndNil(ListeDateDep);
  end;
end ;

end.
