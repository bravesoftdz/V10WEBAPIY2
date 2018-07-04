{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  22/11/05   JP   Création de l'unité : La vignette Indicateur d'En-Cours
--------------------------------------------------------------------------------------}
unit CPVIGNETTETEST;

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
  TObjetTest = class(TCPVignettePlugIn)
  private

  protected
    procedure RecupDonnees ;             override;
    procedure GetClauseWhere;            override;
    procedure DrawGrid(Grille : string); override;
  public

  end;

implementation

uses
  SysUtils, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TObjetTest.DrawGrid(Grille: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ddWriteLN('PGICPTEST : DRAWGRID');

  SetTitreCol('GRID' , 'SYSDOSSIER', 'Dossier');
  SetTitreCol('GRID' , 'SO_SOCIETE', 'Société');
  SetTitreCol('GRID' , 'SO_LIBELLE', 'Libellé');
  SetTitreCol('GRID' , 'SO_ADRESSE2', 'Adresse');

  SetFormatCol('GRID', 'SYSDOSSIER', 'C.0 ---');
  SetFormatCol('GRID', 'SO_SOCIETE', 'C.0 ---');
  SetFormatCol('GRID', 'SO_LIBELLE', 'G.0 ---');
  SetFormatCol('GRID', 'SO_ADRESSE2', 'G.0 ---');

  SetWidthCol('GRID', 'SYSDOSSIER', 10);
  SetWidthCol('GRID', 'SO_SOCIETE', 10);
  SetWidthCol('GRID', 'SO_LIBELLE', 30);
  SetWidthCol('GRID', 'SO_ADRESSE2', 30);

  if not Assigned(TobDonnees.Detail[0]) or
     (TobDonnees.Detail[0].GetString('SYSDOSSIER') = '') then
    SetVisibleCol('GRID', 'SYSDOSSIER', False);
end;

{---------------------------------------------------------------------------------------}
procedure TObjetTest.GetClauseWhere;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TObjetTest.RecupDonnees;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  T : TOB;
  s : string;
begin
  inherited;
  try
    ddWriteLN('PGICPTEST : Debut des traitements TEST');
    s := Regroupement;
    if s <> '' then
      Q := Openselect('SELECT SYSDOSSIER, SO_SOCIETE, SO_LIBELLE, SO_ADRESSE2  FROM SOCIETE')
    else
      Q := Openselect('SELECT  SO_SOCIETE, SO_LIBELLE, SO_ADRESSE2  FROM SOCIETE');
    while not Q.EOF do begin
      T := TOB.Create('µµµµµµµ', TobDonnees, -1);
      if s <> '' then T.AddChampSupValeur('SYSDOSSIER', Q.FindField('SYSDOSSIER').AsString)
                 else T.AddChampSupValeur('SYSDOSSIER', '');
      T.AddChampSupValeur('SO_SOCIETE' , Q.FindField('SO_SOCIETE' ).AsString);
      T.AddChampSupValeur('SO_LIBELLE' , Q.FindField('SO_LIBELLE' ).AsString);
      T.AddChampSupValeur('SO_ADRESSE2', Q.FindField('SO_ADRESSE2').AsString);
      ddWriteLN('PGICPTEST : Une Fille');
      Q.Next;
    end;
    Ferme(Q);

    ddWriteLN('PGICPTEST : Fin des traitements TEST');
    
  except
    on E : Exception do
      MessageErreur := 'Erreur lors du traitement des données avec le message :'#13#13 + E.Message;
  end;
end;

end.


