{-------------------------------------------------------------------------------------
  Version   |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
7.06.001.001  08/08/06  JP   Création de l'unité : suivi global au niveau du groupe (ensemble des
                             sociétés) de tous les types de comptes (bancaires, courants, Titres) et
                             ce de manière globalisée : 1 ligne par société et par type de compte
7.06.001.001  09/11/06  JP   FQ 10256 : Déplacement des soldes initiaux au 01/01 au MATIN
7.06.001.001  16/11/06  JP   Gestion des comptes absents 
--------------------------------------------------------------------------------------}

unit TRSUIVIGROUPE_TOF;

interface

uses
  StdCtrls, Controls, Classes, UFicTableau,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, UTOB, Grids, Windows;

type
  TOF_TRSUIVIGROUPE = class(FICTAB)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
  private
    procedure SupprimeLignePresente(Soc, Nat : string);
  protected
    {Evènements}
    procedure bGraphClick    (Sender : TObject); override;
    function  IsCelulleReinit(Col, Row : Integer) : Boolean; override;
    procedure DessinerCells  (Sender : TObject; Col, Row : Integer; Rect : TRect; State : TGridDrawState); override;
    procedure FormOnKeyDown  (Sender : TObject; var Key : Word; Shift : TShiftState); override;
    procedure DeviseOnChange (Sender : TObject);
    {Clause where sur les banques et les comptes}
    function  GetWhereBqCpt : string; override;
    procedure ParGroupes;
    function  GetLibelle(Lib, Nat : string) : string;
  public
    function  CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean; override;
  end ;

function TRLanceFiche_TRSUIVIGROUPE(Arg : string): string;

implementation

uses
  ExtCtrls, Commun, UProcGen, UProcSolde, Constantes, TRSUIVISOLDE_TOF, HMsgBox,
  uTobDebug, Graphics, UtilPgi, AglInit;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRSUIVIGROUPE(Arg : string): string;
{---------------------------------------------------------------------------------------}
begin
  Result := FICTAB.FicLance('TR', 'TRSUIVIGROUPE', '', '', Arg);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  TypeDesc := dsGroupe;
  inherited;
  Ecran.HelpContext := 5000150;
  LargeurCol := 100;
  InArgument := True;

  ListeRouge.Sorted := True;
  ListeRouge.Duplicates := dupIgnore;

  TRadioButton(GetControl('DATEVAL')).OnClick := CritereClick;

  {Initialisation de la devise}
  Devise := V_PGI.DevisePivot;
  cbDevise.OnChange := DeviseOnChange;
  MajComboDevise(Devise);

  InArgument := False;
  {09/11/06 : Pas le temps de fair un graph pour le moment}
  SetControlVisible('BGRAPH', False);
end;

{Appel de la fiche de suivi de solde correspondant à la ligne en cours
{---------------------------------------------------------------------------------------}
function  TOF_TRSUIVIGROUPE.CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  s : string;
  Q : TQuery;
begin
  {False, pour ne pas lancer le détail des flux}
  Result := inherited CommandeDetail(CurRow, TT);

  {LÉGENDE DES TYPES : "NODOSSIER" = Ligne contenant le solde d'une nature de compte
                       ""          = Ligne contenant le libellé d'une société
                       "*"         = Ligne de séparation entre deux banques
                       "+"         = Total général
                       "-"         = Total par nature de compte
                CODES  "+"         = solde total d'une société avec TYPE = "NODOSSIER"
                       "NATCPTE"   = Ligne contenant le solde d'une nature de compte
                       }

  {Si on est sur une ligne de de libellé, on sort}
  if CurRow[COL_CODE] = '' then Exit;

  if (CurRow[COL_TYPE] = '-') or (CurRow[COL_TYPE] = '+') then
    TT.SetString('REGROUP', '')
  else
    TT.SetString('REGROUP', CurRow[COL_TYPE]);
  TT.SetString('CODE'   , CurRow[COL_CODE]);
  TT.SetString('TYPE'   , CurRow[COL_TYPE]);
  TT.SetString('NATURE' , GetControlText('MCNATURE'));
  TT.SetString('DATEOPE', sDateOpe);
  TT.SetString('DEVISE' , Devise);
  TT.SetString('INTERVAL', GetControlText('INTERVAL'));
  TT.SetInteger('PERIODICITE', Ord(Periodicite));
  TT.SetInteger('APPEL'      , Ord(TypeDesc));
  TT.SetDateTime('DATEDEB', ColToDate(PtClick.X - COL_DATEDE));
  TT.SetDateTime('DATEFIN', ColToDate(PtClick.X - COL_DATEDE, True));
  {On n'affichera que les comptes correspondant à la nature de la ligne courante :
   le filtre va se faire avec la banque}
  if CurRow[COL_CODE] = tcb_Titre then
    {On affiche que la banque "comptes titres"}
    TT.SetString('BANQUE', CODETITRES)
  else if CurRow[COL_CODE] = tcb_Courant then
    {On affiche que la banque "comptes courants"}
    TT.SetString('BANQUE', CODECOURANTS)
  else if CurRow[COL_CODE] = tcb_Bancaire then begin
    {On affiche que les banques "normales"}
    s := cbBanque.Value;
    {Si pas de banque, on prend toutes les banques autre que titre et courant}
    if s = '' then begin
      Q := OpenSQL('SELECT PQ_BANQUE FROM BANQUES WHERE NOT (PQ_BANQUE IN ("' + CODETITRES + '","' + CODECOURANTS + '"))', True);
      while not Q.EOF do begin
        s := s + Q.FindField('PQ_BANQUE').AsString + ';';
        Q.Next;
      end;
      Ferme(Q);
    end
    {Sinon suppression éventuelle des banques Titre et Courant}
    else begin
      s := FindEtReplace(s, CODETITRES, '', True);
      s := FindEtReplace(s, CODECOURANTS, '', True);
      s := FindEtReplace(s, ';;', ';', True);
    end;

    TT.SetString('BANQUE', s);
  end
  else
    {On est sur une ligne de total société, on affichera tous les comptes de la société}
    TT.SetString('BANQUE' , cbBanque.Value);

  TheTob := TT;

  {Rafraichissement en cas de modifications}
  if TRLanceFiche_TRSUIVISOLDE('TR', 'TRSUIVISOLDE', '', '', 'SOLDE;')  <> '' then
    MouetteClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery;
  SQL    : string;
  I      : Integer;
  sCpt,
  Tmp1   : string;
  wBanque: string;
  wDate  : string;
  sNat   : string;
begin
  inherited;
  {31/01/06 : FQ 10335 : Dans l'immédiat, on ne gère plus le suivi bancaire en opération}
  if TypeDesc = dsBancaire then begin
    if rbDateOpe.Checked then rbDateOpe.Checked := False;
    sDateOpe := 'TE_DATEVALEUR';
  end;

  sNat := GetControlText('MCNATURE');
  if (Pos('<<' , sNat) = 0) and (Trim(sNat) <> '') then
    Nature := 'AND TE_NATURE IN (' + GetClauseIn(sNat) + ')'
  else
    Nature := '';

  wDate  := '(' + sDateOpe + ' >= "' + UsDateTime(DateDepart) + '" AND ' + sDateOpe + ' <= "' + UsDateTime(DateFin) + '") ';

  {12/05/04 : Création de la liste des comptes
   24/01/06 : Fait dans le OnLoad suite à la FQ 10328
  CreerListeCompte;}
  if (cbBanque.Value <> '') then begin
    wBanque := 'AND BQ_BANQUE IN (';
    Tmp1 := GetClauseIn(cbBanque.Value);
    wBanque := wBanque + Tmp1 + ') ';
  end
  else
    wBanque := '';

  Tmp1 := wBanque + sCpt;
  SQL := GetRequete(' WHERE ' + wDate + Tmp1);

  Q := OpenSQL(SQL, True); 
  TobListe.LoadDetailDB('1', '', '', Q, False, True);
  Ferme(Q);

  ParGroupes;

  Grid.ColCount := COL_DATEDE + NbColonne;

  {Formatage des montant : '#,##0.00'; Euros}
  SQL := StrFMask(CalcDecimaleDevise(Devise), '', True);
  {Formatage des dates}
  for I := COL_DATEDE to COL_DATEDE + NbColonne - 1 do
    Grid.ColFormats[I] := SQL;

  TobGrid.PutGridDetail(Grid, True, True, '', True);

  Grid.Cells[COL_LIBELLE, 0] := '';

  {Gestion du séparateur de ligne par solde entre chaque banque}
  for i := 0 to Grid.RowCount - 1 do
    if Grid.Cells[COL_TYPE, i] = '*' then Grid.RowHeights[i] := 3
                                     else Grid.RowHeights[i] := 18; {16/07/05 : dans certains cas, on a des éffets inattendus}

  SQL := 'Suivi du groupe ';
  if not MultiNature then begin
         if Pos('"' + na_Realise    + '"', Nature) > 0 then LibNature := ' réalisés'
    else if Pos('"' + na_Simulation + '"', Nature) > 0 then LibNature := ' de simulation'
                                                       else LibNature := ' prévisionnels';
    SQL := SQL + ' (flux' + LibNature + ')';
  end
  else
    LibNature := '';

  Ecran.Caption := TraduireMemoire(SQL);
  UpdateCaption(Ecran);

  Grid.Col := COL_DATEDE;
  Grid.Row := Grid.FixedRows;
  Grid.SetFocus;
end ;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVIGROUPE.IsCelulleReinit(Col, Row : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited IsCelulleReinit(Col, Row);
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVIGROUPE.GetWhereBqCpt : string;
{---------------------------------------------------------------------------------------}
begin
  {13/06/05 : FQ 10250 : On ne gère plus les comptes vides que sur les comptes du filtre}
  Result := inherited GetWhereBqCpt;
  {13/06/05 : FQ 10250 : sinon sur les banques, s'il n'y a pas de filtre sur les comptes}
  if cbBanque.Value <> '' then begin
    Result := 'BQ_BANQUE IN (';
    Result := Result + GetClauseIn(cbBanque.Value);
    Result := Result + ')';
  end
end;

{09/11/06 : FQ 10256 : Refonte de la quête des soldes 
{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.ParGroupes;
{---------------------------------------------------------------------------------------}
var
  TobL,
  TobG : TOB;
  lSociete : string;
  MntTmp   : Double;
  CodeSoc  : string;
  txConv   : Double;
  NbDec    : Integer;
  {LÉGENDE DES TYPES : "NODOSSIER" = Ligne contenant le solde d'une nature de compte
                       ""          = Ligne contenant le libellé d'une société
                       "*"         = Ligne de séparation entre deux banques
                       "+"         = Total général
                       "-"         = Total par nature de compte
                CODES  "+"         = solde total d'une société avec TYPE = "NODOSSIER"
                       "NATCPTE"   = Ligne contenant le solde d'une nature de compte
                       }

    {------------------------------------------------------------------}
    procedure CreeTobDetail(CodSoc, NatC, LibLig, Nat : string; Ind : Integer = -1);
    {------------------------------------------------------------------}
    var
      n : Integer;
    begin
      TobG := TOB.Create('', TobGrid, Ind);
      TobG.AddChampSupValeur('TYPE', CodSoc);
      TobG.AddChampSupValeur('CODE', NatC);
      TobG.AddChampSupValeur('NATURE'  , Nat);
      TobG.AddChampSupValeur('LIBELLE', LibLig);
      for n := 0 to NbColonne - 1 do {Ajout des colonnes dates}
        TobG.AddChampSupValeur(RetourneCol(n), '');
    end;

    {JP 19/03/04 : Le traitement de SommerBanque a été sorti pour pouvoir être
                   exécuté plus tard
    {-------------------------------------------------------------------------}
    procedure CreeTobTotal;
    {-------------------------------------------------------------------------}
    begin
      CreeTobDetail(CodeSoc, '+', GetLibelle(lSociete, na_Total), '');
    end;

    {-------------------------------------------------------------------------}
    procedure SetMontant(Chp : string; Val : Double);
    {-------------------------------------------------------------------------}
    begin
      TobG.SetDouble(Chp, Arrondi(Val, NbDec));
    end;

    {-------------------------------------------------------------------------}
    procedure CalculTotal(Nat : string);
    {-------------------------------------------------------------------------}
    var
      N   : Integer;
      S   : string;
      Mnt : Double;
    begin
      MntTmp := 0.0;

      {Cumule chaque colonne date pour le type flux courant}
      for N := 0 to NbColonne - 1 do begin
        S := RetourneCol(n);
        {Calcul des totaux par société}
        if CodeSoc <> '' then
          Mnt := TobGrid.Somme(S, ['TYPE', 'NATURE'], [CodeSoc, na_Total], False) {Total par société}
        else
          Mnt := TobGrid.Somme(S, ['CODE', 'NATURE'], [Nat, na_Total], False); {Total par nature}

        if Mnt <> 0 then SetMontant(S, Mnt)
                    else SetMontant(S, MntTmp);
        MntTmp := Mnt;
        SetMontant(S, MntTmp);
      end;
    end;

    {Création et calcul des totaux généraux 
    {-------------------------------------------------------------------------}
    procedure GererTotaux(Nat : string);
    {-------------------------------------------------------------------------}
    begin
      if Nat = na_Total then begin
        CreeTobDetail('+', Nat, TraduireMemoire('Total général'), na_Total);
        CodeSoc := '-';
        CalculTotal('');
      end
      else begin
        CodeSoc := '';  {Par précaution !}
        if TobGrid.FindFirst(['CODE'], [Nat], True) <> nil then begin
          CreeTobDetail('-', Nat, GetLibelle('Total', Nat), na_Total);
          CalculTotal(Nat);
        end;
      end;
    end;

    {-------------------------------------------------------------------------}
    procedure CalculeSoldes;
    {-------------------------------------------------------------------------}
    var
      N   : Integer;
      S   : string;
      Mnt : Double;
      UsD : string;
    begin
      if not Assigned(TobG) then Exit;
      MntTmp := 0;

      for N := 0 to NbColonne - 1 do begin
        S := RetourneCol(n);

        {Gestion du solde forcé au premier janvier, dans la mesure où l'on n'est pas sur un total banque}
        if IsColAvecSoldeInit(n) then begin 
          {On mémorise le montant des opérations du premier jour}
          Mnt := TobG.GetDouble(S);
          {Mémorisation des cellules mouvementés}
          if Mnt <> 0 then ListeRouge.Add(S + '|' + TobG.GetString('TYPE') + '|' + TobG.GetString('CODE')+ '|');

          if Periodicite = tp_7 then UsD := USDateTime(DebutAnnee(ColToDate(n + 1)))
                                else UsD := USDateTime(DebutAnnee(ColToDate(n)));

          {JP 28/11/05 : FQ 10317 : chaque type d'affichage demande une gestion spécifique du solde d'initialisation
                         Quotidien : on récupère le solde du 01/01 au matin et on ajoute les opérations du 01/01
                         Hebdomadaire : récupération du solde du 01/01 au Matin (ColToDate(n + 1) pour être sûr de
                                        prendre la bonne année auquel on ajoute les opérations comprises entre le
                                        01/01 et la fin de la semaine contenant le 01/01 (MntHebdo)
                         Mensuel : récupération du solde au 01/01 au matin auquel on rajoute toutes les opérations
                                   du mois de janvier}
          if Periodicite = tp_1 then
            MntTmp := GetSoldeNatureCompte(TobG.GetString('TYPE'), TobG.GetString('CODE'), DateToStr(ColToDate(n)), Nature, cbBanque.Value, not rbDateOpe.Checked, True)
          else if Periodicite = tp_7 then
            MntTmp := GetSoldeNatureCompte(TobG.GetString('TYPE'), TobG.GetString('CODE'), DateToStr(ColToDate(n + 1)), Nature, cbBanque.Value, not rbDateOpe.Checked, True)
          else
            MntTmp := GetSoldeNatureCompte(TobG.GetString('TYPE'), TobG.GetString('CODE'), DateToStr(ColToDate(n)), Nature, cbBanque.Value, not rbDateOpe.Checked, True);
          {Conversion en devise d'affichage}
          MntTmp := MntTmp * txConv;

          {Sauf en affichage mensuel, on force le solde ainsi calculé}
          if Periodicite = tp_7 then begin
            MntTmp := MntTmp + GetlOpeDebAnnee(TobG.GetString('TYPE') + '|' + TobG.GetString('CODE'));
            SetMontant(S, MntTmp);
          end
          else begin
            MntTmp := MntTmp + Mnt;
            SetMontant(S, MntTmp);
          end;
        end
        else if N = 0 then begin
          {On mémorise le montant des opérations du premier jour}
          Mnt := TobG.GetDouble(S);
          {Mémorisation des cellules mouvementés}
          if Mnt <> 0 then ListeRouge.Add(S + '|' + TobG.GetString('TYPE') + '|' + TobG.GetString('CODE')+ '|');
          {Calcul du solde initial}                                                                              {***}
          MntTmp := GetSoldeNatureCompte(TobG.GetString('TYPE'), TobG.GetString('CODE'), DateToStr(DateDepart - 1), Nature, cbBanque.Value, not rbDateOpe.Checked, False);
          {Conversion en devise d'affichage}
          MntTmp := MntTmp * txConv;
          {Ajout les opérations du jour au solde calculé}
          MntTmp := MntTmp + Mnt;
          {Mise à jour du montant}
          SetMontant(S, MntTmp);
        end
        else begin
          {Récupération des opérations du jour}
          Mnt := TobG.GetDouble(S);
          {Mémorisation des cellules mouvementés}
          if Mnt <> 0 then ListeRouge.Add(S + '|' + TobG.GetString('TYPE') + '|' + TobG.GetString('CODE')+ '|');
          {Ajout les opérations de la période au solde de la période précédente}
          MntTmp := Mnt + MntTmp;
          SetMontant(S, MntTmp);
        end;
      end;
    end;


    {Gestion des nature non mouvementées
    {-------------------------------------------------------------------------}
    procedure GererDossiersAbsents;
    {-------------------------------------------------------------------------}
    var
      F : TOB;
      T : TOB;

          {---------------------------------------------------------------}
          procedure CreeLigneSolde(Typ : string; Ind : Integer = -1);
          {---------------------------------------------------------------}
          var
            s : string;
          begin
            s := cbBanque.Value;
            {On s'assure que le filtre sur les banques rend utile la création de la ligne}
            if s <> '' then begin
              if s[Length(s)] <> ';' then s := s + ';';
              {Si la banque des comptes courants ne fait pas partie de la sélection ...}
              if (Typ = tcb_Courant) and (Pos(CODECOURANTS, s) = 0) then Exit;
              {Si la banque des comptes titres ne fait pas partie de la sélection ...}
              if (Typ = tcb_Titre  ) and (Pos(CODETITRES  , s) = 0) then Exit;
              {Si seuls les comptes titres ou / et courants sont dans le filtre, on
               n'affiche pas de ligne pour les comptes bancaires}
              if (Typ = tcb_bancaire) and
                 ((s = CODECOURANTS + ';') or (s = CODETITRES + ';') or
                  (s = CODECOURANTS + ';' + CODETITRES + ';') or
                  (s = CODETITRES + ';' + CODECOURANTS + ';')) then Exit;
            end;
            CreeTobDetail(CodeSoc, Typ, GetLibelle('', Typ), na_Total, Ind);
            CalculeSoldes;
          end;

          {---------------------------------------------------------------}
          function GetIndexTob(NC : string) : Integer;
          {---------------------------------------------------------------}
          begin
            Result := -1;
            F := TobGrid.FindFirst(['LIBELLE'], [lSociete], True);
            {Ne devrait théoriquement pas arrivé}
            if Assigned(F) then begin
                   if NC = tcb_Bancaire then Result := F.GetIndex + 1
              else if NC = tcb_Courant  then Result := F.GetIndex + 2
                                        else Result := F.GetIndex + 3
            end;
          end;

    begin
      if not Assigned(ListeCpte) then Exit;

      T := ListeCpte.FindFirst(['CODDOSSIER'], [CodeSoc], True);
      while T <> nil do begin
        CreeLigneSolde(T.GetString('BQ_NATURECPTE'));
        {Création de la tob qui va contenir les soldes du compte non mouvementé}
        //CreeTobDetail(CodeSoc, T.GetString('COMPTE'), RetourneNomLib(T.GetString('LIBELLE'), na_Total), T.GetString('BQ_NATURECPTE'));
        {28/11/05 : FQ 10317 : Récupération et affectation du solde de départ}
        ///GereSoldeDepart(TobG, aTob.GetString('COMPTE'), aTob.GetString('ADEVISE'), txConv);
        {On supprime ce compte, car il est possible que des banques n'est aucun compte mouvementé.
         A la fin de la boucle, on bouclera sur les comptes qui restent la Tob ListeCpte}
        SupprimeLignePresente(CodeSoc, T.GetString('BQ_NATURECPTE'));
        T := ListeCpte.FindNext(['CODDOSSIER'], [CodeSoc], True);
      end;
      (*
      for n := 0 to ListeCpte.Detail.Count - 1 do begin
        F := ListeCpte.Detail[n];
        if not Assigned(F) then begin
          lSociete := Q.FindField('DOS_LIBELLE').AsString;
          CreeTobDetail('', '', lSociete, '');
          {Création de la ligne des comptes bancaires}
          CreeLigneSolde(tcb_Bancaire);
          {Création de la ligne des comptes courants}
          CreeLigneSolde(tcb_Courant);
          {Création de la ligne des comptes Titres
          CreeLigneSolde(tcb_Titre);}
          {Création de la tob contenant le solde total du dossier}
          CreeTobTotal;
          {Création d'une tob fille vide pour séparer les banques}
          CreeTobDetail('*', '', '', '');
          Q.Next;
          Continue;
        end;
        *)
        {
        lSociete := Q.FindField('DOS_LIBELLE').AsString;

        F := TobGrid.FindFirst(['TYPE', 'CODE'], [Q.FindField('DOS_NODOSSIER').AsString, tcb_Bancaire], True);
        if not Assigned(F) then
          CreeLigneSolde(tcb_Bancaire, GetIndexTob(tcb_Bancaire));

        F := TobGrid.FindFirst(['TYPE', 'CODE'], [Q.FindField('DOS_NODOSSIER').AsString, tcb_Courant], True);
        if not Assigned(F) then
          CreeLigneSolde(tcb_Courant, GetIndexTob(tcb_Courant));

        {F := TobGrid.FindFirst(['TYPE', 'CODE'], [Q.FindField('DOS_NODOSSIER').AsString, tcb_Titre], True);
        if not Assigned(F) then
          CreeLigneSolde(tcb_Titre, GetIndexTob(tcb_Titre));
        
      end;}
    end;

var
  I       : Integer;
  NatCpte : string;
  S, Code : string;
  Montant : Double;
begin
  {On vide la liste des soldes vides}
  ListeRouge.Clear;
  ColRouge := -1;

  Code    := '';
  NatCpte := '';
  lSociete := '';
  TobL    := nil;

  {16/01/06 : FQ 10328 : On ne gère le taux de conversion que si on travaille avec TE_MONTANT}
  txConv := RetContreTaux(V_PGI.DateEntree, V_PGI.DevisePivot, Devise);
  NbDec := CalcDecimaleDevise(Devise);

  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];

    S := TobL.GetString('DOS_LIBELLE');
    {Changement de société}
    if S <> lSociete then begin
      if lSociete <> '' then begin
        GererDossiersAbsents;
        {Calcule les soldes à partir des montants mis dans la grille}
        CalculeSoldes;
        {Création de la tob contenant le solde total du dossier}
        CreeTobTotal;
        {Création d'une tob fille vide pour séparer les banques}
        CreeTobDetail('*', '', '', '');
      end;

      {Mise à jour du code de la banque pour GererComptesAbsents}
      CodeSoc := TobL.GetString('BQ_NODOSSIER');

      lSociete := S;
      {Création de la tob contenant le libellé de la banque}
      CreeTobDetail('', '', lSociete, '');
      NatCpte := '';
    end;

    S := TobL.GetString('BQ_NATURECPTE');

    {Nouvelle nature de compte}
    if NatCpte <> S then begin
      {Si on n'est pas sur la première ligne ...}
      if NatCpte <> '' then begin
        {Calcule les soldes à partir des montants mis dans la grille}
        CalculeSoldes;
      end;

      {Mise à jour du compte bancaire et de son libellé}
      NatCpte := S;
      {Création de la ligne de cumul des opérations pour le nouveau compte et la nature en cours}
      CreeTobDetail(CodeSoc, NatCpte, GetLibelle('', NatCpte), na_Total);
      {Suppression des lignes traitées}
      SupprimeLignePresente(CodeSoc, NatCpte);
    end;

    S := GetTitreFromDate(TobL.GetDateTime(sDateOpe));
    {Dans la fourchette d'affichage}
    if (TobL.GetDateTime(sDateOpe) >= DateDepart) and (TobL.GetDateTime(sDateOpe) <= DateFin) then begin
      Montant := TobL.GetDouble('MONTANT') * txConv;
      SetMontant(S, Montant + TobG.GetDouble(S));
      SetlOpeDebAnnee(CodeSoc + '|' + NatCpte, Montant, TobL.GetDateTime(sDateOpe));
    end;
  end;{Boucle For}

  if lSociete <> '' then begin
    {Calcule les soldes à partir des montants mis dans la grille}
    CalculeSoldes;
    {Mise à jour du code de la banque pour GererComptesAbsents}
    CodeSoc := TobL.GetString('BQ_NODOSSIER');
    GererDossiersAbsents;
    {Pour la dernière banque}
    CreeTobTotal;
    {Création d'une tob fille vide pour séparer les banques}
    CreeTobDetail('*', '', '', '');
  end;
/////////////////////
  {Ajout des sociétés non mouvementées}
  if Assigned(ListeCpte)  then
    {12/05/04 : ajout des banques n'ayant que des comptes non mouvementés}
    while ListeCpte.Detail.Count > 0 do begin
      CodeSoc  := ListeCpte.Detail[0].GetString('CODDOSSIER');
      lSociete := ListeCpte.Detail[0].GetString('LIBDOSSIER');

      {Création d'une tob fille vide pour séparer les Sociétés}
      CreeTobDetail('*', '', '', '');
      {Création de la tob contenant le libellé de la société}
      CreeTobDetail('', '', ListeCpte.Detail[0].GetString('LIBDOSSIER'), '');
      {On ajoute les comptes de la banque}
      GererDossiersAbsents;
      {Pour la dernière banque}
      CreeTobDetail('+', CodeSoc, 'Total ' + lSociete, '');
    end;

  {16/04/04 : Gestion des soldes des comptes, on part du solde initial auquel on ajoute les opérations du jours}
  CodeSoc := '';
  for I := 1 to TobGrid.Detail.Count - 1 do begin
    TobG := TobGrid.Detail[I];
    if (TobG.GetString('CODE') <> '+') or (TobG.GetString('TYPE') = '') then Continue;
    CodeSoc := TobG.GetString('TYPE');
    CalculTotal('');
  end;

  {Totaux généraux}
  CreeTobDetail('', '', TraduireMemoire('Totaux généraux'), '');
  GererTotaux(tcb_Bancaire);
  GererTotaux(tcb_Courant);
  GererTotaux(tcb_Titre);
  GererTotaux(na_Total);
//  TobDebug(TobGrid);
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVIGROUPE.GetLibelle(Lib, Nat: string): string;
{---------------------------------------------------------------------------------------}
begin
       if Nat = tcb_Titre    then Result := Lib + ' Cptes titres'
  else if Nat = tcb_Bancaire then Result := Lib + ' Cptes bancaires'
  else if Nat = tcb_Courant  then Result := Lib + ' Cptes courants'
  else if Nat = na_Total     then Result := 'Total ' + Lib
  else Result := Lib;
end;


{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.bGraphClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  (*
var
  Titre  : string;
  sSolde : string;
  sDev   : string;
  n, p   : Integer;
  T, D   : TOB;
  sGraph : string;
  Ch     : Char;
begin
  (*
  if TypeDesc = dsSolde then
    inherited
  else begin
    if cbCompte.NbSelected <> 1 then begin
      HShowMessage('0;' + Ecran.Caption + ';Veuillez ne sélectionner qu''un seul compte.;W;O;O;O;', '', '');
      Exit;
    end;
    sSolde := cbCompte.Value;
    sSolde := RechDom('TRBANQUECP', ReadTokenSt(sSolde), False);
    Titre := TraduireMemoire('Évolution du rapprochement bancaire : ' + sSolde);

    if rbDateOpe.Checked then sSolde := 'Les montants sont calculés en date d''opération'
                         else sSolde := 'Les montants sont calculés en date de valeur';
    sDev  := 'Les montants sont éxprimés en ' + cbDevise.Items[cbDevise.ItemIndex];

    {Constitution de la chaine qui servira aux paramètres de "LanceGraph"}
    sGraph := 'DATE;ESTIME;POINTE';

    {Constitution de la tob des flux}
    T := Tob.Create('$$$', nil, -1);

    for n := COL_DATEDE to Grid.ColCount - 1 do begin
      D := Tob.Create('$$$', T, -1);
      D.AddChampSup('DATE', False);
      D.AddChampSup('ESTIME', False);
      D.AddChampSup('POINTE', False);
      D.PutValue('DATE', Grid.Cells[n, 0]);
      for p := 1 to Grid.RowCount - 1 do begin
        Ch := StrToChr(Grid.Cells[COL_NATURE, p]);
        {Les lignes contenant les soldes des comptes}
        if Ch = na_Estime then
          D.PutValue('ESTIME', Grid.Cells[n, p])
        else if Ch = na_Pointe then
          D.PutValue('POINTE', Grid.Cells[n, p]);
      end;
    end;

    {!!! Mettre sGraph en fin à cause des ";" qui traine dans la chaine}
    TRLanceFiche_TRGRAPHBANCAIRE(Titre + ';' + sDev + ';' + sSolde + ';' + sGraph, T);
    T.ClearDetail;
    FreeAndNil(T);

  end;
  *)
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.DessinerCells(Sender : TObject; Col, Row : Integer; Rect : TRect; State : TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  S     : string;
  sType : string;
  sCode : string;
begin
  inherited;
  S := Grid.Cells[Col, Row];
  sType := Grid.Cells[COL_TYPE, Row];
  sCode := Grid.Cells[COL_CODE, Row];
  with Grid.Canvas do begin
    {Ligne sélectionnée}
    if gdSelected in State then begin
      Brush.Color := clHighLight;
      Font.Color := clHighLightText;
    end

    {Ligne de libellé des dossiers}
    else if sType = '' then begin
      Font.Style := [fsBold];
      Brush.Color := clBleuLight;//clUltraLight;
    end

    {Ligne de séparation}
    else if sType = '*' then
      Brush.Color := clCremePale

    {On est sur des soldes}
    else if (sType = '+') or (sCode = '+') then begin
      if V_PGI.NumAltCol = 0 then Brush.Color := clInfoBk
                             else Brush.Color := AltColors[V_PGI.NumAltCol];
      Font.Style := [fsBold];
    end

    {Ligne de solde ou total pour une nature}
    else begin
      Brush.Color := clUltraCreme;
      {Gestion de la couleur de la fonte en fonction du sens du solde}
      if (Col >= COL_DATEDE) then begin
             if (Valeur(S) > 0) then Font.Color := clGreen {Solde positif}
        else if (Valeur(S) < 0) then Font.Color := clRed; {Solde négatif}
      end;
      {Gestion des céllules mouvementées}
      sCode := Grid.Cells[Col, 0] + '|' + Grid.Cells[COL_TYPE, Row] + '|' + Grid.Cells[COL_CODE, Row] + '|';
      if (ListeRouge.IndexOf(sCode) > -1) then
        Font.Style := [fsBold];
      {Gestion des montants de réinitialisation}
      if IsColAvecSoldeInit(Col - COL_DATEDE) then
        Font.Color := clBlue;
    end
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.FormOnKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
var
  I : Integer;
  {-----------------------------------------------------------------------}
  function Reached(Col, Row: Integer): Boolean;
  {-----------------------------------------------------------------------}
  begin
    {True si on est sur une ligne de solde par nature de compte}
    Result := (Grid.Cells[Col, Row] <> '') and
              (Grid.Cells[COL_TYPE, Row] <> '*') and
              (Grid.Cells[COL_TYPE, Row] <> '+') and
              (Grid.Cells[COL_CODE, Row] <> '+') and
              (Grid.Cells[COL_TYPE, Row] <> '-');
    Result := Result and (ListeRouge.IndexOf(Grid.Cells[Col, 0] + '|' +
                                             Grid.Cells[COL_TYPE, Row] + '|' +
                                             Grid.Cells[COL_CODE, Row] + '|') > -1);

  end;

begin
  if (Shift = [ssCtrl]) then begin
    case Key of
      VK_RIGHT : for I := Grid.Col + 1 to Grid.ColCount - 1 do
                   {Si la date appartient à la liste c'est qu'il y a des écritures}
                   if Reached(I, Grid.Row) then begin
                     Grid.Col := I;
                     Key := 0;
                     Exit;
                   end;

      VK_LEFT : for I := Grid.Col - 1 downto COL_DATEDE do
                  {Si la date appartient à la liste c'est qu'il y a des écritures}
                   if Reached(I, Grid.Row) then begin
                    Grid.Col := I;
                    Key := 0;
                    Exit;
                  end;
    end;
  end;

  FormKeyDown(Sender, Key, Shift); // Pour touches standard AGL
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  Devise := cbDevise.Value;
  if Devise = '' then Devise := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), Devise);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVIGROUPE.SupprimeLignePresente(Soc, Nat : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  if not Assigned(ListeCpte) then Exit;
  T := ListeCpte.FindFirst(['CODDOSSIER', 'BQ_NATURECPTE'], [Soc, Nat], True);
  if T <> nil then FreeAndNil(T);
end;

initialization
  RegisterClasses([TOF_TRSUIVIGROUPE]);

end.

