{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  22/11/05   JP   Création de l'unité : La vignette Indicateur d'En-Cours
--------------------------------------------------------------------------------------}
unit CPSuiviBudgetaire;

interface

uses
  Classes, UTob,
  {$IFNDEF DBXPRESS}
  dbtables,
  {$ELSE}
  uDbxDataSet,
  {$ENDIF}
  uCPVignettePlugIn, HCtrls;

type
  TObjetSuiviBudget = class(TCPVignettePlugIn)
  private
    Budget  : string;
    Section : string;
    Etab    : string;
    Exercice: string;
    DateDeb : TDateTime;
    DateFin : TDateTime;

    function  CreeTob : TOB;
    procedure SetRealise;
    procedure SetTotaux;
  protected
    function  GetInterface(NomGrille : string = '') : Boolean;    override;
    procedure RecupDonnees ;             override;
    procedure GetClauseWhere;            override;
    procedure DrawGrid(Grille : string); override;
  public

  end;

implementation

uses
  SysUtils, HEnt1, CalcOle, Ent1;

{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBudget.DrawGrid(Grille: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ddWriteLN('PGICPSuiviBudget : DRAWGRID');
  SetTitreCol('GRID' , 'COMPTE'  , 'Compte');
  SetTitreCol('GRID' , 'LIBELLE' , 'Libellé');
  SetTitreCol('GRID' , 'REALISE' , 'Réalisé');
  SetTitreCol('GRID' , 'BUDGETE' , 'Budgété');
  SetTitreCol('GRID' , 'ECART'   , 'Écart');
  SetTitreCol('GRID' , 'POURCENT', 'En %');
  SetFormatCol('GRID', 'COMPTE'  , 'G.0 ---');
  SetFormatCol('GRID', 'LIBELLE' , 'G.0 ---');
  SetFormatCol('GRID', 'REALISE' , 'D.2 ---');
  SetFormatCol('GRID', 'BUDGETE' , 'D.2 ---');
  SetFormatCol('GRID', 'ECART'   , 'D.2 ---');
  SetFormatCol('GRID', 'POURCENT', 'D.4 ---');
  SetWidthCol('GRID' , 'COMPTE'  , 10);
  SetWidthCol('GRID' , 'LIBELLE' , 30);
  SetWidthCol('GRID' , 'REALISE' , 20);
  SetWidthCol('GRID' , 'BUDGETE' , 20);
  SetWidthCol('GRID' , 'ECART'   , 20);
  SetWidthCol('GRID' , 'POURCENT', 20);
end;

{Calcul des montants réalisés : je reprends le traitement de CalcOle : GetCumul
{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBudget.SetRealise;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  Q : TQuery;
  RS : string;
  RC : string;
  C1 : string;
  C2 : string;
  Lib : string;
  TD : TabloExt;

    {Récupération des informations d'un compte budgétaire
    {-------------------------------------------------------------------}
    function SetLibRC(Compte : string) : Boolean;
    {-------------------------------------------------------------------}
    begin
      Q := OpenSelect('SELECT BG_RUB, BG_LIBELLE FROM BUDGENE WHERE BG_BUDGENE = "' + Compte + '"');
      if not Q.EOF then begin
        RC  := Q.FindField('BG_RUB').AsString;
        Lib := Q.FindField('BG_LIBELLE').AsString;
        Result := True;;
      end
      else begin
        MessageErreur := 'Compte non trouvé';
        Result := False;
      end;
      Ferme(Q);
    end;

var
  Reali : Double;
  Delta : Double;
begin
  {Récupération des informations de la section budgétaire sélectionnée}
  Q := OpenSelect('SELECT BS_RUB, BS_LIBELLE FROM BUDSECT WHERE BS_BUDSECT = "' + Section + '"');
  if not Q.EOF then begin
    RS := Q.FindField('BS_RUB').AsString;
    SetControlValue('LBLIBELLE', 'Suivi budgétaire : section ' + Q.FindField('BS_RUB').AsString + ' (' + Section + ')');
    Ferme(Q);
  end
  else begin
    MessageErreur := 'Section non trouvée';
    Ferme(Q);
    Exit;
  end;

  ddWriteLN('CPSuiviBudgetaire : Récupération des cumuls');

  {Le Calcul sur les montants budgétés s'est fait sur les bases d'un regroupement
   par compte budgétaire}
  for n := 0 to TobDonnees.Detail.Count - 1 do begin
    {Récupération des informations d'un compte budgétaire}
    if not SetLibRC(TobDonnees.Detail[n].GetString('COMPTE')) then Exit;
    {Constitution du code de la rubrique en fonction de la section et du compte bugétaire courant :
     Les rubriques bugétaire section par général permettent de faire le lien entre le Budget et la
     comptabilité générale}
    C1 := 'S/G' + Budget + RS + ':' + RC;
    {Calcul du cumul dans la compta générale}
    GetCumul('RUBREA', C1, C2, 'SAN', Etab, 'EUR', Exercice, DateDeb, DateFin, False, True, nil, TD, False);

    {TD[3] = Debit et TD[2] = Crédit}
    Reali := TD[3] - TD[2];
    TobDonnees.Detail[n].SetString('REALISE', Format('%.2n', [Reali]));
    {Calcul de l'écart}
    Delta := Reali - TobDonnees.Detail[n].GetDouble('BUDGETE');
    {Calcul du pourcentage  : Ecart ramené à ce qui est bugété (Repris de l'état "Ecart budgétaire" de la compta}
    if TobDonnees.Detail[n].GetDouble('BUDGETE') <> 0 then Reali := Delta / TobDonnees.Detail[n].GetDouble('BUDGETE') * 100
                                                      else Reali := 0;

    {Affectation de la Tob et formatage des montants}
    TobDonnees.Detail[n].SetString('LIBELLE', Lib);
    TobDonnees.Detail[n].SetString('ECART', Format('%.2n', [Delta]));
    TobDonnees.Detail[n].SetString('POURCENT', Format('%.4n', [Abs(Reali)]) + '%');
    TobDonnees.Detail[n].SetString('BUDGETE', Format('%.2n', [TobDonnees.Detail[n].GetDouble('BUDGETE')]));
  end;
end;

{Génération de la lignes des Totaux
{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBudget.SetTotaux;
{---------------------------------------------------------------------------------------}
var
  MntR : Double;
  MntB : Double;
  F    : TOB;
  n    : Integer;
begin
  MntR := 0;
  MntB := 0;

  {On boucle sur les comptes budgétaires pour calculer le total sur la section}
  for n := 0 to TobDonnees.Detail.Count - 1 do begin
    MntR := MntR + Valeur(TobDonnees.Detail[n].GetString('REALISE'));
    MntB := MntB + Valeur(TobDonnees.Detail[n].GetString('BUDGETE'));
  end;

  F := CreeTob;
  F.SetString('COMPTE'  , '*******');
  F.SetString('LIBELLE' , 'Totaux');
  F.SetString('POURCENT', '*********');
  F.SetString('REALISE' , Format('%.2n', [MntR]));
  F.SetString('BUDGETE' , Format('%.2n', [MntB]));
  F.SetString('ECART'   , Format('%.2n', [MntR - MntB]));
end;

{Constitution de la clause where sur la Table BudEcr, pour le calcul des montants budgétés
{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBudget.GetClauseWhere;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetClause(' WHERE BE_BUDJAL = "' + Budget + '"');
  SetClause('BE_BUDSECT = "' + Section + '"');

  if Etab <> '' then
    SetClause('BE_ETABLISSEMENT = "' + Etab + '"');

  if DateDeb > iDate1900 then
    SetClause('BE_DATECOMPTABLE >= "' + USDateTime(DateDeb) + '"');
  if DateFin > iDate1900 then
    SetClause('BE_DATECOMPTABLE <= "' + USDateTime(DateFin) + '"');
end;

{Affectation des critères de traitement à partir des critères de sélection de la vignette
{---------------------------------------------------------------------------------------}
function TObjetSuiviBudget.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}

      {On s'assure que les Fourchettes de sélection sont dans le bon ordre
      {--------------------------------------------------------------------}
      procedure VerifCoherence(var Deb, Fin : string; Chp : string);
      {--------------------------------------------------------------------}
      var
        Tmp : Variant;
      begin
        if Deb > Fin then begin
          Tmp := Deb;
          Deb := Fin;
          Fin := Tmp;
          SetCritereValue(Chp + '1', Deb);
          SetCritereValue(Chp + '2', Fin);
        end;
      end;

var
  Q : TQuery;
  D : TDateTime;
begin
  Result := inherited GetInterface;

  Budget  := GetString('FJOURNAL');
  Section := GetString('FSECTION');
  Etab    := GetString('ETABLISSEMENT');

  {Il faut Choisir un budget}
  if Budget = '' then begin
    MessageErreur := 'Veuillez choisir un budget';
    Result := False;
  end

  {Il faut Choisir une section budgétaire}
  else if Section = '' then begin
    MessageErreur := 'Veuillez choisir une section';
    Result := False;
  end

  else begin
    DateDeb := GetDateTime('FDATE');
    DateFin := GetDateTime('FDATE_');

    {Contrôle des dates de début et de fin en fonction du budget}
    Q := OpenSelect('SELECT BJ_PERDEB, BJ_PERFIN, BJ_EXODEB FROM BUDJAL WHERE BJ_BUDJAL = "' + Budget + '"');
    try
      if not Q.EOF then begin
        Exercice := Q.FindField('BJ_EXODEB').AsString;
//        ddWriteLN('PGICPVIGNETTES : DEB 1/ ' + DateToStr(DateDeb));
  //      ddWriteLN('PGICPVIGNETTES : FIN 1/ ' + DateToStr(DateFin));

        {Cohérence des dates entre elles}
        if DateFin < DateDeb then begin
          D := DateDeb;
          DateDeb := DateFin;
          DateFin := D;
          DateDeb := DebutDeMois(DateDeb);
          DateFin := FinDeMois(DateFin);
          SetCritereValue('FDATE', DateDeb);
          SetCritereValue('FDATE_', DateFin);
        end;

        {Cohérence de la date de début avec la date de début du budget}
        if (DateDeb < Q.FindField('BJ_PERDEB').AsDateTime) or
           (DateDeb > Q.FindField('BJ_PERFIN').AsDateTime) then begin
          DateDeb := Q.FindField('BJ_PERDEB').AsDateTime;
          SetCritereValue('FDATE', DateDeb);
        end;

        {Cohérence de la date de fin avec la date de fin du budget}
        if (DateFin > Q.FindField('BJ_PERFIN').AsDateTime) or
           (DateFin < Q.FindField('BJ_PERDEB').AsDateTime) then begin
          DateFin := Q.FindField('BJ_PERFIN').AsDateTime;
          SetCritereValue('FDATE_', DateFin);
        end;

      end

      else begin
        MessageErreur := 'Budget introuvable';
        Result := False;
      end
    finally
      Ferme(Q);
    end;
  end;
end;

{Création d'une tob Fille à TobDonnees
{---------------------------------------------------------------------------------------}
function TObjetSuiviBudget.CreeTob : TOB;
{---------------------------------------------------------------------------------------}
begin
  Result := TOB.Create('£BUDGET', TobDonnees, -1);
  Result.AddChampSup('COMPTE'  , False);
  Result.AddChampSup('LIBELLE' , False);
  Result.AddChampSup('REALISE' , False);
  Result.AddChampSup('BUDGETE' , False);
  Result.AddChampSup('ECART'   , False);
  Result.AddChampSup('POURCENT', False);
end;

{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBudget.RecupDonnees;
{---------------------------------------------------------------------------------------}
const
  s = 'SELECT SUM(BE_DEBIT - BE_CREDIT) AS MNT, BE_BUDGENE FROM BUDECR ';
  g = ' GROUP BY BE_BUDGENE ORDER BY BE_BUDGENE';
var
  Q : TQuery;
  T : TOB;
begin
  inherited;
  ddWriteLN('CPSuiviBudgetaire : Debut des traitements du suivi budgétaire');
  try

//    ddWriteLN('CPSuiviBudgetaire : requete budget'#13 + s + ClauseWhere + g);
    Q := OpenSelect(s + ClauseWhere + g);
    try
      Q.First;                     
      while not Q.EOF do begin
        T := CreeTob;
        T.SetString('COMPTE' , Q.FindField('BE_BUDGENE').AsString);
        T.SetDouble('BUDGETE', Q.FindField('MNT').AsFloat);
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;
    {Calcul du réalisé par compte}
    SetRealise;
    {Calcul des totaux}
    SetTotaux;
  except
    on E : Exception do
      MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
  end;
end;

end.


