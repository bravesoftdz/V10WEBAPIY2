{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  14/04/06   JP   Création de l'unité : Retards clients
 7.03.001.006  05/10/06   JP   Initialisation de la grille si pas de données cf.DawGrid
--------------------------------------------------------------------------------------}
unit CPVignetteRetardClient;

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
  TObjRetardClient = class(TCPVignettePlugIn)
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
procedure TObjRetardClient.DrawGrid(Grille : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited;
  {05/10/06 : Pour forcer le portail à réinitialiser de la grille}
  if TobDonnees.Detail.Count = 0 then begin
    T := TOB.Create('RETARD', TobDonnees, -1);;
    T.AddChampSupValeur('TIERS', '');
    T.AddChampSupValeur('ECHE', '');
    T.AddChampSupValeur('MONTANT', '');
    PutGridDetail('GRID', TobDonnees);
  end;

  ddWriteLN('PGICPRETARDCLIENT : DRAWGRID');

  SetTitreCol('GRID' , 'TIERS'  , 'Tiers');
  SetTitreCol('GRID' , 'ECHE'   , '1ère éché');
  SetTitreCol('GRID' , 'MONTANT', 'Montant');
  SetFormatCol('GRID', 'TIERS'  , 'G.0 ---');
  SetFormatCol('GRID', 'ECHE'   , 'C.0 ---');
  SetFormatCol('GRID', 'MONTANT', 'D.2 ---');
  SetWidthCol('GRID' , 'TIERS'  , 45);
  SetWidthCol('GRID' , 'ECHE'   , 20);
  SetWidthCol('GRID' , 'MONTANT', 25);
end;

{---------------------------------------------------------------------------------------}
procedure TObjRetardClient.RecupDonnees;
{---------------------------------------------------------------------------------------}
begin
  if EtatOk then begin
    {Pour ne pas lancer le SetInterface}
    SansInterface := True;
    ddwriteln('PGICPRETARDCLIENT : RecupDonnees = ETAT');
    RemplitTobDonneesEtat;
    FTip := 'E';
    FNat := 'VCP';
    FModele := 'VCR';
    FLanguePrinc := 'FRA';
    Rapport;
  end
  else
    RemplitTobDonnees;
end;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/10/2006
Modifié le ... :   /  /    
Description .. : - LG  - 09/10/2006 - FB 10002 - on recupere les ecritures 
Suite ........ : typée H et on supprime on garde le signe de 
Suite ........ : e_debit-e_credit pour prendre en compte les avoir
Mots clefs ... : 
*****************************************************************}
procedure TObjRetardClient.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  Q   : TOB;
  QQ  : TOB;
  SQL : string;
  wh  : string;
  n   : Integer;
  T   : TOB;
begin
  if GetString('CBMONTANT', False) <> '' then
    Wh := 'AND ABS(E_DEBIT - E_CREDIT) - E_COUVERTURE > ' + GetString('CBMONTANT', False) + ' '
  else
    SetControlValue('CBMONTANT', '0');

  SQL := 'SELECT T_LIBELLE, MIN(E_DATEECHEANCE) ECHE, ((SUM(E_DEBIT - E_CREDIT)) - SUM(E_COUVERTURE)) MONTANT ' +
         'FROM ECRITURE LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
         'WHERE E_QUALIFPIECE = "N" AND (  E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H" ) AND E_DATEECHEANCE < "' + UsDateTime(DateRef) + '" ' +
         'AND E_NUMECHE >= 0 AND E_ETATLETTRAGE IN ("AL", "PL") AND NOT (E_GENERAL LIKE "416%") ' +
         'AND T_NATUREAUXI = "CLI" AND T_LETTRABLE = "X" ' + Wh +
         'GROUP BY T_LIBELLE ORDER BY MONTANT DESC, ECHE';

  ddwriteln('PGICPRETARDCLIENT : SQL = ' + SQL);

  Q := OpenSelectInCache(SQL);
  try
    for n := 0 to Q.Detail.count - 1 do begin
      if N > 9 then Break;
      QQ := Q.Detail[n];
      T := TOB.Create('§§§', TobDonnees, -1);;
      T.AddChampSupValeur('TIERS'  , QQ.GetString('T_LIBELLE'));
      T.AddChampSupValeur('ECHE'   , QQ.GetString('ECHE'));
      T.AddChampSupValeur('MONTANT', Format('%.2n', [QQ.GetDouble('MONTANT')]));
    end;
  finally
    FreeAndNil(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : J. Pasteris
Créé le ...... : 06/10/2006
Modifié le ... :   /  /    
Description .. : - FB 10002 - LG - corection de l'affichage de la date, on 
Suite ........ : formate directement la zone
Mots clefs ... : 
*****************************************************************}
procedure TObjRetardClient.RemplitTobDonneesEtat;
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
  SQL : string;
  wh  : string;
  n   : Integer;
  T   : TOB;
begin
  if GetString('CBMONTANT', False) <> '' then
    Wh := 'AND ABS(E_DEBIT - E_CREDIT) - E_COUVERTURE > ' + GetString('CBMONTANT', False) + ' '
  else
    SetControlValue('CBMONTANT', '0');

  SQL := 'SELECT T_LIBELLE, MIN(E_DATEECHEANCE) MINECHE, ((SUM(E_DEBIT - E_CREDIT)) - SUM(E_COUVERTURE)) MONTANT, ' +
         'E_AUXILIAIRE, MAX(E_DATEECHEANCE) MAXECHE FROM ECRITURE LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
         'WHERE E_QUALIFPIECE = "N" AND (  E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H" ) AND E_DATEECHEANCE < "' + UsDateTime(DateRef) + '" ' +
         'AND E_NUMECHE >= 0 AND E_ETATLETTRAGE IN ("AL", "PL") AND NOT (E_GENERAL LIKE "416%") ' +
         'AND T_NATUREAUXI = "CLI" AND T_LETTRABLE = "X" ' + Wh +
         'GROUP BY T_LIBELLE, E_AUXILIAIRE ORDER BY MONTANT DESC, MINECHE';

  ddwriteln('PGICPRETARDCLIENT : SQL = ' + SQL);

  n := 1;
  Q := OpenSelect(SQL);
  try
    while (not Q.EOF) and (n <= 100) do begin
      Inc(n);
      T := TOB.Create('§§§', TobDonnees, -1);;
      T.AddChampSupValeur('TIERS'  , Q.FindField('T_LIBELLE').AsString);
      T.AddChampSupValeur('MINECHE', Q.FindField('MINECHE').AsString);
      T.AddChampSupValeur('MAXECHE', Q.FindField('MAXECHE').AsString);
      T.AddChampSupValeur('CODET'  , Q.FindField('E_AUXILIAIRE').AsString);
      T.AddChampSupValeur('E_MONTANT', FormatFloat('#,##0.00', Q.FindField('MONTANT').AsFloat));
      Q.Next;
    end;

  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjRetardClient.TraiteParam(Param : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  EtatOk := Pos('ETAT', Param) > 0;
  ddwriteln('PGICPRETARDCLIENT : TraiteParam = ' + FParam);
end;

end.


