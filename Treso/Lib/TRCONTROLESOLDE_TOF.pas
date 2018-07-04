{-------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
7.09.001.002  23/11/06   JP   Création du contrôle des soldes avec la compta et EEXBQ
8.10.001.013  08/10/07   JP   FQ 10522 : accès aux détail des flux
--------------------------------------------------------------------------------------}

unit TRCONTROLESOLDE_TOF ;

interface

uses
  StdCtrls, Controls, Classes, UFicTableau,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, UTOB, Grids, Windows;

type
  TOF_TRCONTROLESOLDE = class(FICTAB)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
  private
    FNomBase   : string;
    FNoDossier : string;
    FDevise    : string;
    FLibelle   : string;
    FGeneral   : string;
    FWhEexBq   : string;  

    function  ValideSelection : Boolean;
    procedure Affichage;
  protected

    {Evènements}
    function  IsCelulleReinit(Col, Row : Integer) : Boolean; override;
    procedure DessinerCells(Sender: TObject; Col, Row: Integer; Rect: TRect; State: TGridDrawState); override;
    procedure GeneralOnChange(Sender : TObject);
  public
    function  CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean; override;
    function  GetCumulCompte(DateC : TDateTime; Cpt : string; NomBase : string = '') : Double;
    function  IsDebExo(Col : Integer; AvecSolde : Boolean; var aSolde : Double) : Boolean;
  end;

function TRLanceFiche_TRCONTROLESOLDE(Arguments : string) : string;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Graphics, Commun, UProcGen, Constantes, HMsgBox, UtilPgi, UProcSolde, uTobDebug, Ent1,
  ULibExercice, ULibPointage;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRCONTROLESOLDE(Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := FICTAB.FicLance('TR', 'TRCONTROLESOLDE', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLESOLDE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  TypeDesc := dsControle;
  inherited;
  Ecran.HelpContext := 50000159;

  LargeurCol := 100;
  InArgument := True;

  ListeRouge.Sorted := True;
  ListeRouge.Duplicates := dupIgnore;
  THEdit(GetControl('GENERAL')).OnChange := GeneralOnChange;
end;

{---------------------------------------------------------------------------------------}
function  TOF_TRCONTROLESOLDE.CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Lib : string;
begin
  Result := inherited CommandeDetail(CurRow, TT);
  {Si on est sur une ligne de totalisation, on sort}
  if (CurRow[COL_NATURE] = '') then Exit;
  {A cette date, il n'y a pas d'écriture, on n'affiche donc pas le détail}
  if (CurRow[PtClick.X] = '0') or (CurRow[PtClick.X] = '') then Exit;

  Lib := Grid.Cells[Grid.Col, 0] + '|' + CurRow[COL_TYPE] + '|';
  if (ListeRouge.IndexOf(Lib) = -1) then Exit;

  TT.SetString('GENERAL', FGeneral);
  TT.SetString('CODE', GetControlText('GENERAL'));
  TT.SetString('REGROUP', FNoDossier);
  TT.SetString('NOMBASE', FNomBase);
  TT.SetString('DATEOPE', sDateOpe);
  TT.SetInteger('APPEL', Ord(TypeDesc));
  TT.SetInteger('PERIODICITE', Ord(Periodicite));
  TT.SetString('LIBNAT', LibNature);
  TT.SetDateTime('DATEDEB', ColToDate(PtClick.X - COL_DATEDE));
  TT.SetDateTime('DATEFIN', ColToDate(PtClick.X - COL_DATEDE));

  TT.SetString('NATURE', 'AND TE_NATURE = "' + na_Realise + '"');

  {La date est gérée dans UFicTableau.GridOnDblClick}
  Lib := ' ' + GetControlText('GENERAL');

  {Sur TRECRITURE}
  if CurRow[COL_TYPE] = 'T' then begin
    TT.SetString('TYPE', 'T');
    Lib := 'Opérations de trésorerie du compte ' + Lib;
    if rbDateOpe.Checked then TT.SetString('DATEAFF', 'TE_DATEVALEUR')
                         else TT.SetString('DATEAFF', 'TE_DATECOMPTABLE');
  end

  {Sur ECRITURE}
  else if rbDateOpe.Checked then begin
    TT.SetString('TYPE', 'C');
    Lib := 'Opérations comptables du compte ' + Lib;
  end

  {Sur EEXBQ}
  else begin
    TT.SetString('TYPE', 'B');{EEXBQ}
    Lib := 'Opérations bancaires du compte ' + Lib;
  end;

  TT.SetString('LIBELLE', Lib);
  Result := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLESOLDE.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  Q     : TQuery;
  SQL   : string;
  I     : Integer;
  sCpt,
  wDate : string;
begin
  {Valisation de la sélection et récupération des infos bancaires et sociétés du compte sélectionné}
  if not ValideSelection then Exit;

  inherited;

  {REQUÊTE SUR LES ÉCRITURES DE TRÉSORERIE}
  Nature := GetWhereDossier;
  if Nature <> '' then Nature := ' AND ' + Nature;
  Nature := Nature + ' AND TE_NATURE = "' + na_realise + '" ';

  wDate  := sDateOpe + ' BETWEEN "' + UsDateTime(DateDepart) + '" AND "' + UsDateTime(DateFin) + '" ';

  {Constitution du filtre sur les comptes}
  sCpt := ' AND TE_GENERAL = "' + GetControlText('GENERAL') + '" ';

  SQL := 'SELECT SUM(TE_MONTANTDEV) AMONTANT, TE_GENERAL ACOMPTE, ' + sDateOpe + ' ADATE, ' +
         'TE_DEVISE ADEVISE FROM TRECRITURE ' +
         'WHERE ' + wDate + Nature + sCpt + ' ' +
         'GROUP BY TE_GENERAL, TE_DEVISE, ' + sDateOpe;

  Q := OpenSQL(SQL, True);
  TobListe.LoadDetailDB('1', '', '', Q, False, True);
  Ferme(Q);

  if rbDateOpe.Checked then begin
    {REQUÊTE SUR LES ÉCRITURES DE COMPTABILITÉ}
    wDate  := 'E_DATECOMPTABLE BETWEEN "' + UsDateTime(DateDepart) + '" AND "' + UsDateTime(DateFin) + '" ';

    {Constitution du filtre sur les comptes}
    sCpt := 'E_GENERAL = "' + FGeneral + '" ';

    SQL := 'SELECT SUM(E_DEBIT - E_CREDIT) AMONTANT, BQ_CODE ACOMPTE, E_DATECOMPTABLE ADATE, ' +
           '"" ADEVISE FROM ' + GetTableDossier(FNomBase, 'ECRITURE') + ' ' +
           'LEFT JOIN BANQUECP ON BQ_GENERAL = E_GENERAL AND BQ_NODOSSIER = "' + FNoDossier + '" ' +
           'WHERE ' + wDate + ' AND E_QUALIFPIECE = "N" AND E_ECRANOUVEAU <> "OAN" AND ' + sCpt +
           'GROUP BY BQ_CODE, E_DATECOMPTABLE ';
  end
  else begin
    {REQUÊTE SUR LES RELEVÉS}
    wDate  := 'CEL_DATEVALEUR BETWEEN "' + UsDateTime(DateDepart) + '" AND "' + UsDateTime(DateFin) + '" ';

    {Constitution du filtre sur les comptes}
    if EstPointageSurTreso then begin
      sCpt := ' CEL_GENERAL = "' + GetControlText('GENERAL') + '" ';
      FNomBase := '';
    end
    else
      sCpt := ' CEL_GENERAL = "' + FGeneral + '" ';

    FWhEexBq := sCpt;

    {09/10/07 : FQ 10522 : Maintenant on travaille sur EEXBQLIG et non plus sur EEXBQ}
    SQL := 'SELECT IIF(CEL_DEVISE = "' + V_PGI.DevisePivot + '", ' +
                      'SUM(CEL_CREDITEURO-CEL_DEBITEURO), SUM(CEL_CREDITDEV-CEL_DEBITDEV)) AMONTANT, ' +
           'CEL_GENERAL ACOMPTE, CEL_DATEVALEUR ADATE, CEL_DEVISE ADEVISE FROM ' + GetTableDossier(FNomBase, 'EEXBQLIG') +
           ' WHERE ' + FWhEexBq + ' AND ' + wDate +
           ' GROUP BY CEL_GENERAL, CEL_DATEVALEUR, CEL_DEVISE' +
           ' ORDER BY CEL_DATEVALEUR';
  end;

  Q := OpenSQL(SQL, True);
  TobRappro.LoadDetailDB('2', '', '', Q, False, True);
  Ferme(Q);

  Affichage;

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

  SQL := 'Suivi des soldes ';
  if not MultiNature then begin
         if Pos('"' + na_Realise    + '"', Nature) > 0 then LibNature := ' réalisés'
    else if Pos('"' + na_Simulation + '"', Nature) > 0 then LibNature := ' de simulation'
                                                       else LibNature := ' prévisionnels';
    SQL := SQL + ' (flux' + LibNature + ')';
  end
  else
    LibNature := '';

  if TypeDesc = dsSolde then
    Ecran.Caption := TraduireMemoire(SQL)
  else
    Ecran.Caption := TraduireMemoire('Fiche de suivi bancaire');
  UpdateCaption(Ecran);

  Grid.Col := COL_DATEDE;
  Grid.Row := Grid.FixedRows;
  Grid.SetFocus;

end;

{---------------------------------------------------------------------------------------}
function TOF_TRCONTROLESOLDE.IsCelulleReinit(Col, Row : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited IsCelulleReinit(Col, Row);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLESOLDE.Affichage;
{---------------------------------------------------------------------------------------}
var
  TobL,
  TobC, {Tob comptable/relevé}
  TobG : TOB; {Tob Tréso}
  txConv   : Double;
  NbDec    : Integer;
  {LÉGENDE DES TYPES : ""           = Ligne contenant le libellé du compte
                       "T"          = Solde de Trésorerie
                       "$"          = Solde Bancaire / comptable
                       "<"          = Ècart
  }

    {------------------------------------------------------------------}
    procedure CreeTobDetail(TypeFlux, CodeFlux, LibelleFlux, Nat : string);
    {------------------------------------------------------------------}
    var
      n : Integer;
    begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('TYPE', TypeFlux);
      TobG.AddChampSupValeur('CODE', CodeFlux);
      TobG.AddChampSupValeur('NATURE'  , Nat);
      TobG.AddChampSupValeur('LIBELLE', LibelleFlux);
      for n := 0 to NbColonne - 1 do {Ajout des colonnes dates}
        TobG.AddChampSupValeur(RetourneCol(n), '');
    end;

    {-------------------------------------------------------------------------}
    procedure Sommer;
    {-------------------------------------------------------------------------}
    var
      N   : Integer;
      S   : string;
      Mnt : Double;
      UsD : string;
      MntTmp : Double;
    begin
      Mnt    := 0.0;

      {Cumule chaque colonne date pour le type flux courant}
      for n := 0 to NbColonne - 1 do begin
        S := RetourneCol(n);

        {Gestion du solde forcé au premier janvier, dans la mesure où l'on est sur la ligne de Trésorerie}
        if IsColAvecSoldeInit(n) and
          ((TobG.GetString('TYPE') = 'T') or ((TobG.GetString('TYPE') = '$') and not rbDateOpe.Checked)) then begin
          UsD := USDateTime(DebutAnnee(ColToDate(n)));

          if ExisteSQL('SELECT TE_GENERAL FROM TRECRITURE WHERE TE_DATEVALEUR = "' + UsD +
                       '" AND TE_GENERAL = "' + GetControlText('GENERAL') + '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"') then begin
            {02/08/07 : FQ 10504 : Il faut récupérer le soldes au matin et non au soir}
            Mnt := GetSoldeMillesime(GetControlText('GENERAL'), DateToStr(ColToDate(n)), Nature, not rbDateOpe.Checked, True, False);

            MntTmp := Arrondi(Mnt, NbDec);
            Mnt := TobG.GetDouble(S);
            Mnt := MntTmp + Mnt;
            TobG.SetDouble(S, Arrondi(Mnt, NbDec));
          end
          else begin
            MntTmp := TobG.GetDouble(S);
            Mnt := MntTmp + Mnt;
          end;
        end

        {Récupération du solde initial comptable}
        else if (n = 0) and (TobG.GetString('TYPE') = 'T') then begin
          Mnt := GetSoldeInit(GetControlText('GENERAL'), DateToStr(ColToDate(n)), Nature, not rbDateOpe.Checked);
          MntTmp := Arrondi(Mnt, NbDec);
          Mnt := TobG.GetDouble(S);
          Mnt := MntTmp + Mnt;
          TobG.SetDouble(S, Arrondi(Mnt, NbDec));
        end

        else if (n = 0) and (TobG.GetString('TYPE') = '$') and rbDateOpe.Checked then begin
          Mnt := GetCumulCompte(ColToDate(n), FGeneral, FNomBase) * txConv;
          MntTmp := Arrondi(Mnt, NbDec);
          Mnt := TobG.GetDouble(S);
          Mnt := MntTmp + Mnt;
          TobG.SetDouble(S, Arrondi(Mnt, NbDec));
        end

        {Recherche d'un éventuel début d'exercice comptable}
        else if (TobG.GetString('TYPE') = '$') and rbDateOpe.Checked and IsDebExo(n, True, Mnt) then begin
          Mnt := Mnt * txConv;
          {Dans ce cas, le solde calculé dans IsDebExo s'impose dans la grille}
          TobG.SetDouble(S, Arrondi(Mnt, NbDec));
        end
        
        {09/10/07 : FQ 10522 : Récupération du solde initial bancaire}
        else if (TobG.GetString('TYPE') = '$') and not rbDateOpe.Checked and (n = 0) then begin
          Mnt := GetSoldeRelevesBQE(GetControlText('GENERAL'), ColToDate(n), FNomBase);
          MntTmp := Arrondi(Mnt, NbDec);
          Mnt := TobG.GetDouble(S);
          Mnt := MntTmp + Mnt;
          TobG.SetDouble(S, Arrondi(Mnt, NbDec));
        end

        {Calcul des écarts}
        else if (TobG.GetString('TYPE') = '<') then begin
          Mnt := Arrondi(TobGrid.Detail[1].GetDouble(S) - TobGrid.Detail[2].GetDouble(S), NbDec);
          TobG.SetDouble(S, Mnt);
        end

        {Ajout du solde précédent aux flux des opérations en cours pour les flux de Tréso et de compta}
        else begin
          MntTmp := TobG.GetDouble(S);
          Mnt := MntTmp + Mnt;
          TobG.SetDouble(S, Mnt);
        end;
      end;
    end;

var
  I    : Integer;
  S       : string;
  Montant : Double;
begin
  {On vide la liste des soldes vides}
  ListeRouge.Clear;
  ColRouge := -1;
  NbDec := CalcDecimaleDevise(FDevise);

  {Gestion des devises :
   1/ les soldes de Trésorerie (Tob1) sont dans la devise du compte (FDevise)
   2.1/ en Opération : on travaille sur ECRITURE en devise Pivot => il faut les convertir dans FDevise
   2.2/ en Valeur : on travaille sur EEXBQ : il faut convertir de la devise du relevé vers FDevise}
  txConv := 1;

  {Création des tobs Filles}
  CreeTobDetail('' , '' , FLibelle + ' (Dev. : ' + FDevise + ')', '');
  if rbDateOpe.Checked then begin
    CreeTobDetail('T', 'T', TraduireMemoire('Solde Tréso. (Opé.)'), 'S');
    CreeTobDetail('$', '$', TraduireMemoire('Solde Compta.'), 'S');
    txConv := RetContreTaux(V_PGI.DateEntree, V_PGI.DevisePivot, FDevise);
  end else begin
    CreeTobDetail('T', 'T', TraduireMemoire('Solde Tréso. (Val.)'), 'S');
    CreeTobDetail('$', '$', TraduireMemoire('Solde Banque'), 'S');
  end;
  CreeTobDetail('<', '>', TraduireMemoire('Écart'), '');

  TobG := TobGrid.Detail[1];
  TobC := TobGrid.Detail[2];

  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];

    S := GetTitreFromDate(TobL.GetDateTime('ADATE'));
    {Dans la fourchette d'affichage}
    if (TobL.GetDateTime('ADATE') >= DateDepart) and (TobL.GetDateTime('ADATE') <= DateFin) then begin
      Montant := Arrondi(TobL.GetDouble('AMONTANT'), NbDec);
      TobG.SetDouble(S, Montant);
      ListeRouge.Add(S + '|T|');
    end;
  end;{Boucle For}

  for I := 0 to TobRappro.Detail.Count - 1 do begin
    TobL := TobRappro.Detail[I];

    S := GetTitreFromDate(TobL.GetDateTime('ADATE'));
    {Dans la fourchette d'affichage}
    if (TobL.GetDateTime('ADATE') >= DateDepart) and (TobL.GetDateTime('ADATE') <= DateFin) then begin
      {Pour la gestion des devises voir le commentaire en début de fonction
       08/10/07 : FQ 10522 : correction de la gestion de ListeRouge pour le détail}
      if rbDateOpe.Checked then 
        Montant := Arrondi(TobL.GetDouble('AMONTANT') * txConv, NbDec)
      else
        Montant := Arrondi(TobL.GetDouble('AMONTANT'), NbDec);

      ListeRouge.Add(S + '|$|');
      TobC.SetDouble(S, Montant);
    end;
  end;{Boucle For}

  {16/04/04 : Gestion des soldes des comptes, on part du solde initial auquel on ajoute les opérations du jours}
  for I := 1 to 3 do begin
    TobG := TobGrid.Detail[I];
    Sommer;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLESOLDE.DessinerCells(Sender: TObject; Col, Row: Integer; Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  S, sType : string;
  D : Double;
begin
  {Soit détail des soldes en Tréso multi sociétés, soit Suivi des soldes en mono société,
  on reprend le code de UFicTableau}
  inherited;
  D := 0;
  S := Grid.Cells[Col, Row];
  sType := Grid.Cells[COL_TYPE, Row];

  with Grid.Canvas do begin
    {On est sur des totaux}
    if (sType = '<') then begin
      if V_PGI.NumAltCol = 0 then Brush.Color := clInfoBk
                             else Brush.Color := AltColors[V_PGI.NumAltCol];
      Font.Style := [fsBold];
    end

    {On est sur le solde comptable / bancaire ou le solde de Trésorerie}
    else if (sType = '$') or (sType = 'T') then begin

      if sType = '$' then Brush.Color := clBleuLight
                     else Brush.Color := clOrangePale;
      if (Col >= COL_DATEDE) then begin
        {Gestion du signe du solde}
             if (Valeur(S) > 0) then Font.Color := clGreen {Solde positif}
        else if (Valeur(S) < 0) then Font.Color := clRed; {Solde négatif}
        {Gestion des céllules mouvementées}
        S := Grid.Cells[Col, 0] + '|T|';
        if (ListeRouge.IndexOf(S) > -1) and (sType = 'T') then
          Font.Style := [fsBold];
        {Gestion des céllules mouvementées}
        S := Grid.Cells[Col, 0] + '|$|';
        if (ListeRouge.IndexOf(S) > -1) and (sType = '$') then
          Font.Style := [fsBold];
        {Gestion des montants de réinitialisation}
        if (IsColAvecSoldeInit(Col - COL_DATEDE) and (sType = 'T')) or
           (IsDebExo(Col - COL_DATEDE, False, D) and (sType = '$')) then
          Font.Color := clBlue;
      end;
    end;
  end; {With Canvas}
end;

{---------------------------------------------------------------------------------------}
function TOF_TRCONTROLESOLDE.ValideSelection : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Result := False;
  if GetControlText('GENERAL') = '' then begin
    {On n'affiche pas le message au chargement de la fonction}
    if not InArgument then
      HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner un compte.;W;O;O;O;', '', '');
    InArgument := False;
    Exit;
  end;

  Q := OpenSQL('SELECT BQ_NODOSSIER, BQ_GENERAL, DOS_NOMBASE, BQ_NATURECPTE, BQ_DEVISE, BQ_LIBELLE FROM BANQUECP ' +
               'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
               'WHERE BQ_CODE = "' + GetControlText('GENERAL') + '"', True);
  try
    if not Q.Eof then begin
      FNomBase   := Q.FindField('DOS_NOMBASE' ).AsString;
      FNoDossier := Q.FindField('BQ_NODOSSIER').AsString;
      FDevise    := Q.FindField('BQ_DEVISE'   ).AsString;
      FLibelle   := Q.FindField('BQ_LIBELLE'  ).AsString;
      FGeneral   := Q.FindField('BQ_GENERAL'  ).AsString;
      if ((Q.FindField('BQ_NATURECPTE').AsString = tcb_Courant) or
          (Q.FindField('BQ_NATURECPTE').AsString = tcb_Titre)) and rbDateOpe.Checked then begin
        HShowMessage('1;' + Ecran.Caption + ';La fiche de contrôle en date de valeur ne concerne que les comptes courants.'#13 +
                     'Les soldes seront calculés en date d''opération.;I;O;O;O;', '', '');
        rbDateOpe.Checked := True;
      end;
      Result := True;
    end
    else
      HShowMessage('0;' + Ecran.Caption + ';Impossible de récupérer les informations bancaires;W;O;O;O;', '', '');
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRCONTROLESOLDE.GetCumulCompte(DateC : TDateTime; Cpt : string; NomBase : string = '') : Double;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  E : TExoDate;
begin
  Result := 0;
  {On travaille sur une exercice ouvert}
  if GetEncours.Deb < DateC then begin
    Q := OpenSQL('SELECT SUM(E_DEBIT) - SUM(E_CREDIT) FROM ' + GetTableDossier(NomBase, 'ECRITURE') +
                 ' WHERE E_GENERAL = "' + Cpt + '" AND E_DATECOMPTABLE BETWEEN "' + UsDateTime(GetEncours.Deb) +
                 '" AND "' + UsDateTime(DateC - 1) + '" AND E_QUALIFPIECE = "N" AND (E_ECRANOUVEAU IN ("N", "H") OR ' +
                 '(E_ECRANOUVEAU = "OAN" AND E_DATECOMPTABLE = "' + UsDateTime(GetEncours.Deb) + '"))', True);
    if not Q.EOF then
      Result := Q.Fields[0].AsFloat;
    Ferme(Q);
  end
  else begin
    {On est sur un exercie clos}
    if CQuelExercice(DateC - 1, E) then begin
      Q := OpenSQL('SELECT SUM(E_DEBIT) - SUM(E_CREDIT) FROM ' + GetTableDossier(NomBase, 'ECRITURE') +
                   ' WHERE E_GENERAL = "' + Cpt + '" AND E_DATECOMPTABLE BETWEEN "' + UsDateTime(E.Deb) +
                   '" AND "' + UsDateTime(DateC - 1) + '" AND E_QUALIFPIECE = "N" AND (E_ECRANOUVEAU IN ("N", "H") OR ' +
                   '(E_ECRANOUVEAU = "OAN" AND E_DATECOMPTABLE = "' + UsDateTime(E.Deb) + '"))', True);
      if not Q.EOF then
        Result := Q.Fields[0].AsFloat;
      Ferme(Q);
    end
    {La date ne correspond à aucun exercice}
    else begin
      HShowMessage('0;' + Ecran.Caption + ';Le solde comptable ne sera qu''une approximation car la date'#13 +
                   'ne correspond à aucun exercice ouvert.;I;O;O;O;', '', '');
      Q := OpenSQL('SELECT SUM(E_DEBIT) - SUM(E_CREDIT) FROM ' + GetTableDossier(NomBase, 'ECRITURE') +
                   ' WHERE E_GENERAL = "' + Cpt + '" AND E_DATECOMPTABLE <= "' + UsDateTime(DateC - 1) +
                   '" AND E_QUALIFPIECE = "N" ', True);
      if not Q.EOF then
        Result := Q.Fields[0].AsFloat;
      Ferme(Q);
    end;
  end
end;

{---------------------------------------------------------------------------------------}
function TOF_TRCONTROLESOLDE.IsDebExo(Col : Integer; AvecSolde : Boolean; var aSolde : Double) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  E : TExoDate;
  D : TDateTime;
begin
  D := ColToDate(Col);
  Result := D = GetEnCours.Deb;
  if not Result then
    if CQuelExercice(D, E) then Result := D = E.Deb;

  if Result and AvecSolde then begin
    {Si on est sur un début d'exercice, on reprend toutes les écritures du jour, OAN comprises}
    Q := OpenSQL('SELECT SUM(E_DEBIT) - SUM(E_CREDIT) FROM ' + GetTableDossier(FNomBase, 'ECRITURE') +
                 'WHERE E_QUALIFPIECE = "N" AND E_DATECOMPTABLE = "' + UsDateTime(D) + '"', True);
    if not Q.EOF then aSolde := Q.Fields[0].AsFloat;
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCONTROLESOLDE.GeneralOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  MajAffichageDevise(GetControl('IDEV'), GetControl('DEV'), GetControlText('GENERAL'), sd_Compte);
end;

initialization
  RegisterClasses([TOF_TRCONTROLESOLDE]);

end.
