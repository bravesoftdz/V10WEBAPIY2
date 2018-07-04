{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  22/12/05   JP   Création de l'unité : Récupération de données et action
--------------------------------------------------------------------------------------}
unit CPVIGNETTEACTION;

interface

uses
  Classes, UTob, DBTables, uCPVignettePlugIn, HCtrls;

type
  TObjetAction = class(TCPVignettePlugIn)
  private
    procedure TraitementFaux;
    procedure TraitementVrai;
  protected
    function  GetInterface(NomGrille : string = '') : Boolean; override;
    procedure DrawGrid(Grille : string); override;
    procedure RecupDonnees;              override;
  public

  end;

implementation

uses
  SysUtils, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TObjetAction.DrawGrid(Grille: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ddWriteLN('PGICPACTION : DRAWGRID');
//  TobResponse.SaveToXMLFile('C:\Documents and Settings\Pasteris\Bureau\Draw.txt', True, True);

(*
  if not Assigned(TobDonnees.Detail[0]) or
     (TobDonnees.Detail[0].GetString('SYSDOSSIER') = '') then
    SetVisibleCol('GRID', 'SYSDOSSIER', False);
    *)
end;


{---------------------------------------------------------------------------------------}
function TObjetAction.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {On passe en paramètres le nom de la grille pour récupérer les données sélectionnées du mul}
  Result := inherited GetInterface('FListe');
  if Result then begin
    Result := ((FParam = 'VRAI') and (TobSelection.Detail.Count > 0)) or (FParam <> 'VRAI');
    if not Result then MessageErreur := 'Il n''y a pas de données à traiter';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetAction.RecupDonnees;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  try
    ddWriteLN('PGICPACTION : Debut des traitements ACTION');
    if FParam = 'VRAI' then TraitementVrai
                       else TraitementFaux;
    ddWriteLN('PGICPACTION : Fin des traitements ACTION');
  except
    on E : Exception do
      MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetAction.TraitementFaux;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  inherited;
  try
    for n := 0 to 10 do begin
      T := TOB.Create('µµµµµ', TobDonnees, -1);
      T.AddChampSupValeur('Premier', FParam, CSTString);
      if (n mod 2) = 0 then T.AddChampSupValeur('Deuxième', True, CSTBoolean)
                       else T.AddChampSupValeur('Deuxième', False, CSTBoolean);
      T.AddChampSupValeur('Troisième', 'OUI', CSTString);
      T.AddChampSupValeur('Quatrième', 'FIN', CSTString);
    end;
  except
    raise;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetAction.TraitementVrai;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  inherited;
  try
    for n := 0 to TobSelection.Detail.Count - 1 do begin
      T := TOB.Create('µµµµµ', TobDonnees, -1);
      T.Dupliquer(TobSelection.Detail[n], False, True);
      if (n mod 2) = 0 then T.AddChampSupValeur('Acceptée', True, CSTBoolean)
                       else T.AddChampSupValeur('Acceptée', False, CSTBoolean);
    end;
  except
    raise;
  end;
end;

end.


