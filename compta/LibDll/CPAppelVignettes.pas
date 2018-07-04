{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  21/11/05   JP   Création de l'unité : Unité d'appel des vignettes de la compta
--------------------------------------------------------------------------------------}
unit CPAppelVignettes;

interface

uses
  Classes, UTOB,
  eSession
;

  procedure InitApplication;
  function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;

implementation

uses
  SysUtils, HEnt1, eDispatch,
  uCPVignettePlugIn;

function VignettesDispatch(Action, Param : string; RequestTOB : TOB) : TOB; forward;
function GetVignetteDictionary : TOB; forward;

{---------------------------------------------------------------------------------------}
procedure InitApplication;
{---------------------------------------------------------------------------------------}
begin
  ddWriteLN('PGICPVIGNETTES Initialisation...');
end;

{---------------------------------------------------------------------------------------}
function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;
{---------------------------------------------------------------------------------------}
begin
 ddwriteln('PGICPVIGNETTES : Dispatch');
 Action := UpperCase(Action);
 Result := VignettesDispatch(Action, Param, RequestTOB);
end;

{---------------------------------------------------------------------------------------}
function VignettesDispatch(Action, Param : string; RequestTOB : TOB) : TOB;
{---------------------------------------------------------------------------------------}
begin
  Result := nil;
  Action := UpperCase(Action);

  ddWriteLN('PGICPVIGNETTES : VignettesDispatch (' + Action + ')');

  if not Assigned(RequestTob) then begin
    ddWriteLN('PGICPVIGNETTES : RequestTOB non assignée');
    Exit;
  end;
  //lasession := lookupcurrentsession (TISession ??)
  if Action = 'ENCOURS' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetEnCours')
  else if Action = 'SUIVIBUDGETAIRE' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetSuiviBudget')
  else if Action = 'ENCOURSTIERS' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetEnCoursTiers')
  else if Action = 'TEST' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetTest')
  else if Action = 'VERT' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetAction')
  else if Action = 'ROUGE' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetAction')
  {Vignettes effectivement proposées}
  else if Action = 'VALIDBAP' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetValidBap')
  else if Action = 'GRAPHECHE' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjGraphEche')
  else if Action = 'ETATECHE' then
    Result := TCPVignettePlugIn.NouvelleInstance('ETAT', RequestTOB, 'TObjGraphEche')
  else if Action = 'GRAPHIMPAYE' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjGraphImpaye')
  else if Action = 'ETATIMPAYE' then
    Result := TCPVignettePlugIn.NouvelleInstance('ETAT', RequestTOB, 'TObjGraphImpaye')
  else if Action = 'GRAPHMBVA' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjGraphMBVA')
  else if Action = 'GRAPHTYPECR' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjGraphTypeEcr')
  else if Action = 'GRAPHEFFET' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjGraphEffets')
  else if Action = 'ETATEFFET' then
    Result := TCPVignettePlugIn.NouvelleInstance('ETAT', RequestTOB, 'TObjGraphEffets')
  else if Action = 'GRAPHVOLUME' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjGraphVolume')
  else if Action = 'VIERSCORING' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjScoringClient')
  else if Action = 'ETATSCORING' then
    Result := TCPVignettePlugIn.NouvelleInstance('ETAT', RequestTOB, 'TObjScoringClient')
  else if Action = 'VIERRETARD' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjRetardClient')
  else if Action = 'ETATRETARD' then
    Result := TCPVignettePlugIn.NouvelleInstance('ETAT', RequestTOB, 'TObjRetardClient')
  else if Action = 'SUIVIBAP' then
    Result := TCPVignettePlugIn.NouvelleInstance(Param, RequestTOB, 'TObjetSuiviBap')
  else if Action = 'GETVIGNETTEDICTIONARY' then
    Result := GetVignetteDictionary

  else begin
    Result := TOB.Create('OUT', nil, -1);
    Result.AddChampSupValeur('ERROR','');
    ddWriteLN('PGICPVIGNETTES : Action non trouvée');
  end;
  //Result.SaveToXMLFile('C:\Documents and Settings\Pasteris\Bureau\tob.txt', True, True);
end;

{---------------------------------------------------------------------------------------}
function GetVignetteDictionary : TOB;
{---------------------------------------------------------------------------------------}
const
  BOUQUET_GRATUIT  = '00485080';
  BOUQUET_TEST     = '00561000';


   {--------------------------------------------------------------------------------}
   procedure add(AName: string; AModule: string = BOUQUET_TEST); 
   {--------------------------------------------------------------------------------}
   var
     T : tob;
    begin
      T := TOB.Create('dictionary_item', result,-1);
      T.AddChampSupValeur('CODE',AName);
//      T.AddChampSupValeur('SERIA_MODULE',AModule);
    T.AddChampSupValeur('SERIA_MODULE',BOUQUET_GRATUIT); // Tout est gratuit : on est chez ED !!!
    end;

begin
  Result:= TOB.Create ('DICTIONARY', nil, -1);

  Add('CPVIGEFFETS');
  Add('CPVIGIMPAYES');
  Add('CPVIGMBVA');
  Add('CPVIGNETTEECHE');
  Add('CPVIGRETARDCLIENT');
  Add('CPVIGSCORINGCLI');
  Add('CPVIGTYPEECRITURE');
  Add('CPVIGVOLUMETRIE');
  Add('CPVIGNETTEMULBAP');
  Add('CPVIGNETTEMULBAP');
  Add('CPVIGSUIVIBAP');
end;

end.

