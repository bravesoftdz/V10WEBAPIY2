{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  13/04/06   JP   Création de l'unité : Scoring clients
 7.03.001.006  05/10/06   JP   Initialisation de la grille si pas de données cf.DawGrid
--------------------------------------------------------------------------------------}
unit CPVignetteScoringClient;

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
  TObjScoringClient = class(TCPVignettePlugIn)
  private
    EtatOk : Boolean;

    procedure RemplitTobDonnees;
    procedure RemplitTobDonneesEtat;

  protected
    procedure TraiteParam(Param : string); override;
    procedure RecupDonnees ;               override;
    procedure DrawGrid(Grille : string);   override;
  public

  end;

implementation

uses
  SysUtils, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TObjScoringClient.DrawGrid(Grille : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited;
  {05/10/06 : Pour forcer le portail à réinitialiser de la grille}
  if TobDonnees.Detail.Count = 0 then begin
    T := TOB.Create('_SCORING', TobDonnees, -1);;
    T.AddChampSupValeur('TIERS', '');
    T.AddChampSupValeur('SCORE', '');
    PutGridDetail('GRID', TobDonnees);
  end;
  ddWriteLN('PGICPSCORINGCLIENT : DRAWGRID');

  SetTitreCol('GRID' , 'TIERS', 'Tiers');
  SetTitreCol('GRID' , 'SCORE', 'Score');
  SetFormatCol('GRID', 'TIERS', 'G.0O ---');
  SetFormatCol('GRID', 'SCORE', 'C.0O ---');
  SetWidthCol('GRID', 'TIERS', 58);
  SetWidthCol('GRID', 'SCORE', 12);
end;

{---------------------------------------------------------------------------------------}
procedure TObjScoringClient.RecupDonnees;
{---------------------------------------------------------------------------------------}
begin
  if GetInteger('CBNOMBRE', False) = 0 then
    SetControlValue('CBNOMBRE', '3');

  if EtatOk then begin
    {Pour ne pas lancer le SetInterface}
    SansInterface := True;
    ddwriteln('PGICPSCORINGCLIENT : RecupDonnees = ETAT');
    RemplitTobDonneesEtat;
    FTip := 'E';
    FNat := 'VCP';
    FModele := 'VCS';
    FLanguePrinc := 'FRA';
    Rapport;
  end
  else
    RemplitTobDonnees;
end;

{---------------------------------------------------------------------------------------}
procedure TObjScoringClient.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  Q   : TOB;
  QQ  : TOB;
  SQL : string;
  wh  : string;
  n   : Integer;
  k   : Integer;
  T   : TOB;
  ScoreRel : Integer;
  TotNiv   : Integer;
begin
  k := GetInteger('CBNOMBRE', False);
  if GetString('CBMONTANT', False) <> '' then
    Wh := 'AND E_DEBIT > ' + GetString('CBMONTANT', False) + ' '
  else
    SetControlValue('CBMONTANT', '0');

  SQL := 'SELECT COUNT(E_NIVEAURELANCE) RELANCEQTE, SUM(E_NIVEAURELANCE) RELANCETOTAL, T_LIBELLE ' +
         'FROM ECRITURE LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
         'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
         'WHERE E_ECHE = "X" AND E_NUMECHE >= 0 AND E_ETATLETTRAGE = "TL" AND T_NATUREAUXI = "CLI" AND ' +
         'T_LETTRABLE = "X" AND E_DEBIT > 0 AND NOT (E_GENERAL LIKE "416%") AND J_NATUREJAL = "VTE" ' + Wh +
         'GROUP BY T_LIBELLE' ;

  ddwriteln('PGICPSCORINGCLIENT : SQL = ' + SQL);

  Q := OpenSelectInCache(SQL);
  try
    for n := 0 to Q.Detail.count - 1 do begin
      QQ := Q.Detail[n];
      if ((QQ.GetInteger('RELANCEQTE') < k) or
          (QQ.GetInteger('RELANCEQTE') = 0 )) then
         Continue;

      T := TOB.Create('§§§', TobDonnees, -1);;
      T.AddChampSupValeur('TIERS', QQ.GetString('T_LIBELLE'));

      TotNiv   := QQ.GetInteger('RELANCETOTAL') + QQ.GetInteger('RELANCEQTE');
      ScoreRel := ((TotNiv - 1) div QQ.GetInteger('RELANCEQTE')) + 1;

      T.AddChampSupValeur('SCORE', ScoreRel);
    end;

    TobDonnees.Detail.Sort('SCORE');

    if TobDonnees.Detail.Count > 10 then
      while TobDonnees.Detail.Count > 9 do begin
        if Assigned(TobDonnees.Detail[0]) then begin
          T := TobDonnees.Detail[0];
          FreeAndNil(T);
        end
        else
          Break;
      end;

    {Pourmettre les plus gros score en haut}
    for n := 0 to TobDonnees.Detail.Count - 1  do
      TobDonnees.Detail[n].ChangeParent(TobDonnees, 0);

  finally
    FreeAndNil(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjScoringClient.RemplitTobDonneesEtat;
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
  SQL : string;
  wh  : string;
  n   : Integer;
  T   : TOB;
  ScoreRel : Integer;
  TotNiv   : Integer;
begin
  n := GetInteger('CBNOMBRE', False);

  if GetString('CBMONTANT', False) <> '' then
    Wh := 'AND E_DEBIT > ' + GetString('CBMONTANT', False) + ' '
  else
    SetControlValue('CBMONTANT', '0');

  SQL := 'SELECT COUNT(E_NIVEAURELANCE) RELANCEQTE, SUM(E_NIVEAURELANCE) RELANCETOTAL, T_LIBELLE, E_AUXILIAIRE, ' +
         'SUM(E_DEBIT) MONTANT FROM ECRITURE LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
         'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
         'WHERE E_ECHE = "X" AND E_NUMECHE >= 0 AND E_ETATLETTRAGE = "TL" AND T_NATUREAUXI = "CLI" AND ' +
         'T_LETTRABLE = "X" AND E_DEBIT > 0 AND NOT (E_GENERAL LIKE "416%") AND J_NATUREJAL = "VTE" ' + Wh +
         'GROUP BY T_LIBELLE, E_AUXILIAIRE' ;

  ddwriteln('PGICPSCORINGCLIENT : SQL = ' + SQL);

  Q := OpenSelect(SQL);
  try
    while not Q.EOF do begin
      if ((Q.FindField('RELANCEQTE').AsInteger < n) or
          (Q.FindField('RELANCEQTE').AsInteger = 0 )) then begin
         Q.Next;
         Continue;
       end;

      T := TOB.Create('§§§', TobDonnees, -1);;
      T.AddChampSupValeur('TIERS', Q.FindField('T_LIBELLE').AsString);
      T.AddChampSupValeur('CODET', Q.FindField('E_AUXILIAIRE').AsString);

      TotNiv   := Q.FindField('RELANCETOTAL').AsInteger + Q.FindField('RELANCEQTE').AsInteger;
      ScoreRel := ((TotNiv - 1) div Q.FindField('RELANCEQTE').AsInteger ) + 1;

      T.AddChampSupValeur('E_SCORE'  , IntToStr(ScoreRel));
      T.AddChampSupValeur('E_MONTANT', FormatFloat('#,##0.00', Q.FindField('MONTANT').AsFloat));
      T.AddChampSupValeur('E_NOMBRE', Q.FindField('RELANCEQTE').AsString);

      Q.Next;
    end;

    TobDonnees.Detail.Sort('E_SCORE;E_MONTANT;');
    if TobDonnees.Detail.Count > 100 then
      while TobDonnees.Detail.Count > 100 do begin
        if Assigned(TobDonnees.Detail[0]) then begin
          T := TobDonnees.Detail[0];
          FreeAndNil(T);
        end
        else
          Break;
      end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjScoringClient.TraiteParam(Param : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  EtatOk := Pos('ETAT', Param) > 0;
  ddwriteln('PGICPSCORINGCLIENT : TraiteParam = ' + FParam);
end;

end.


