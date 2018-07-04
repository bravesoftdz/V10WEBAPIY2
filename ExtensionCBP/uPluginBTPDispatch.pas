unit uPluginBTPDispatch;

interface

uses
  uTOB;

function Dispatch(Action, Param:string; RequestTOB:TOB):TOB ; stdcall;

implementation

uses
  Sysutils,
  BTPGetStructure,
  BTPGetVersions,
  uBTPVerrouilleDossier,
  uBTPGetDocument;

const
  // Mettre ici le nom de votre Dll de plugin
  PLUGIN_NAME = 'PluginBTPS';
  PLUGIN_VERSION = '7.3.9.720';

//Fonction de dispatch du plugin
function Dispatch(Action, Param:string; RequestTOB:TOB):TOB ; stdcall;
begin
  // Créer systématiquement la tob résultat
  Result := TOB.Create ('Out', nil, -1);
  try
    // Ajouter systématiquement ce champ afin de rendre les tests génériques
    Result.AddChampSupValeur('ERROR', '');

    if Action = 'GetStructure' then
    begin
    	// -----------
    	if (not RequestTOB.fieldExists('SQL')) then
      begin
      	result.putValue ('ERROR',PLUGIN_NAME+'.'+Action+' requète d''entrée non définie');
      end else if RequestTOB.getValue('SQL')='' then
      begin
      	result.putValue ('ERROR',PLUGIN_NAME+'.'+Action+' requète d''entrée Vide');
      end else begin
    		BTPGetFieldStructure (RequestTOB.getValue('SQL'),result);
      end;
      // -----------
    end else if Action = 'ISBaseDisponible' then
    begin
      ISBaseDisponible (result);
    end else if Action = 'GetVersionRef' then
    begin
    	BTPGetVersionRef (Result);
    end else if Action = 'GetDocument' then
    begin
    	if (RequestTOB.fieldExists('CLEDOC')) and (RequestTOB.GetString('CLEDOC')<>'') then
      begin
    		BTPGetDocument (RequestTOB,Result);
      end else
      begin
      	result.putValue ('ERROR',PLUGIN_NAME+'.'+Action+' DOCUMENT non définie');
      end;
    end else
    begin
      //Provoquera le soulèvement d'une exception par le client
      Result.Free ;
      Result := nil;
    end;
  except
    Result.free;
    //Toujours remonter les erreurs au serveur
    raise;
  end;
end;

end.
