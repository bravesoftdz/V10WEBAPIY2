{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  22/11/05   JP   Création de l'unité : Ancêtre des vignettes Compta 
--------------------------------------------------------------------------------------}
unit uCPVignettePlugIn;

interface

uses
  Classes, UTob, DBTables, HEnt1, uAncetreVignettePlugIn;

type
  TCPVignettePlugIn = class(TAncetreVignettePlugIn)
  private
    {Fait office de RegisterClasses}
    class function GetObjet(TobIn, TobOut : TOB; ClassName : string) : TCPVignettePlugIn;
  protected
    {Récupération des controls/valeurs depuis la TobRequest}
    function  GetInterface(NomGrille : string = '') : Boolean; override;
    {Affectation des controls/valeurs dans la TobRespons}
    function  SetInterface : Boolean; override;
  public
    class function NouvelleInstance(Param : string; RequestTob : TOB; ClassName : string) : TOB; override;

    constructor Create(TobIn, TobOut : Tob); override;
    destructor  Destroy; override;
  end;

implementation

uses
  SysUtils,
  CPEnCoursTiers, CPEnCours, CPSuiviBudgetaire, CPVignetteTest, CPVignetteAction,
  CPVignetteValidBAP, CPVignetteGraphEcheances, CPVignetteGraphImpayes,
  CPVignetteGraphMBVA, CPVignetteGraphTypeEcr, CPVignetteGraphEffets, CPVignetteGraphVolumetrie,
  CPVignetteScoringClient, CPVignetteRetardClient, CPVigSuiviBap{14/05/08 : Demande SIC}; 

{Fonction de création de l'objet et lancement du traitement métier
{---------------------------------------------------------------------------------------}
class function TCPVignettePlugIn.NouvelleInstance(Param : string; RequestTob : TOB; ClassName : string) : TOB;
{---------------------------------------------------------------------------------------}
var
  ResponseTOB : TOB;
  MonObjet    : TCPVignettePlugIn;
begin
  ddWriteLN('PGICPVIGNETTES : NouvelleInstance');

  ResponseTOB := TOB.Create('OUT', nil, -1);
  ResponseTOB.AddChampSupValeur('ERROR','');

  MonObjet := TCPVignettePlugIn.GetObjet(RequestTob, ResponseTOB, ClassName);
  try
    ddWriteLN('PGICPVIGNETTES : Début de traitement');
    MonObjet.LanceTraitement(Param);
  finally
    ddWriteLN('PGICPVIGNETTES : Fin de traitement');
    if Assigned(MonObjet) then FreeAndNil(monObjet);
  end;

  Result := ResponseTOB;
  if Result = nil then
    ddWriteLN('PGICPVIGNETTES : RESPONSETOB VIDE');
//  Result.SaveToXMLFile('C:\Documents and Settings\Pasteris\Bureau\tob.txt', True, True);
end;

{---------------------------------------------------------------------------------------}
function TCPVignettePlugIn.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetInterface(NomGrille);
end;

{---------------------------------------------------------------------------------------}
function TCPVignettePlugIn.SetInterface : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited SetInterface;
end;

{---------------------------------------------------------------------------------------}
constructor TCPVignettePlugIn.Create(TobIn, TobOut : Tob);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(TobIn, TobOut);
end;

{---------------------------------------------------------------------------------------}
destructor TCPVignettePlugIn.Destroy;
{---------------------------------------------------------------------------------------}
begin
  inherited Destroy;
end;

{Cette fonction fait office de RegisterClass
{---------------------------------------------------------------------------------------}
class function TCPVignettePlugIn.GetObjet(TobIn, TobOut: TOB; ClassName: string): TCPVignettePlugIn;
{---------------------------------------------------------------------------------------}
begin
       if ClassName = 'TObjetEnCours'      then Result := TObjetEnCours     .Create(TobIn, TobOut)
  else if ClassName = 'TObjetEnCoursTiers' then Result := TObjetEnCoursTiers.Create(TobIn, TobOut)
  else if ClassName = 'TObjetSuiviBudget'  then Result := TObjetSuiviBudget .Create(TobIn, TobOut)
  else if ClassName = 'TObjetTest'         then Result := TObjetTest        .Create(TobIn, TobOut)
  else if ClassName = 'TObjetAction'       then Result := TObjetAction      .Create(TobIn, TobOut)
  else if ClassName = 'TObjetValidBap'     then Result := TObjetValidBap    .Create(TobIn, TobOut)
  else if ClassName = 'TObjGraphEche'      then Result := TObjGraphEche     .Create(TobIn, TobOut)
  else if ClassName = 'TObjGraphImpaye'    then Result := TObjGraphImpaye   .Create(TobIn, TobOut)
  else if ClassName = 'TObjGraphMBVA'      then Result := TObjGraphMBVA     .Create(TobIn, TobOut)
  else if ClassName = 'TObjGraphTypeEcr'   then Result := TObjGraphTypeEcr  .Create(TobIn, TobOut)
  else if ClassName = 'TObjGraphEffets'    then Result := TObjGraphEffets   .Create(TobIn, TobOut)
  else if ClassName = 'TObjGraphVolume'    then Result := TObjGraphVolume   .Create(TobIn, TobOut)
  else if ClassName = 'TObjScoringClient'  then Result := TObjScoringClient .Create(TobIn, TobOut)
  else if ClassName = 'TObjRetardClient'   then Result := TObjRetardClient  .Create(TobIn, TobOut)
  else if ClassName = 'TObjetSuiviBap'     then Result := TObjetSuiviBap    .Create(TobIn, TobOut)
                                           else Result := TCPVignettePlugIn .Create(TobIn, TobOut);
end;

end.

