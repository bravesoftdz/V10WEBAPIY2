unit ComEdServer;

interface

uses classes, UTOB, UImportCom, UExportCom, Ent1, ULibExercice, UWA, galSystem,
galMainSauveDossier,UGedFiles, Paramsoc;

procedure InitApplication;
function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;

implementation

uses
  sysutils, DBTables,
  HEnt1, HCtrls,
  ULibCpContexte,
  eSession, UtilTrans, windows;

function ComsxDispatch(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function ComsxEdInit(Action, Param: string; RequestTOB: TOB): TOB;  forward;

function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;
begin
 ddwriteln('dispatch ....................................');
 // PB devise pivot n'est pas renseignée
Try
 V_PGI.DevisePivot:=GetParamSocSecur('SO_DEVISEPRINC','') ;
 Action := UpperCase(Action);
 result := ComsxDispatch(Action, Param, RequestTOB);
Except
      On e:exception Do
      begin
         ddwriteln(e.Message) ;
      end;
End ;
end;

function ComsxDispatch(Action, Param: string; RequestTOB: TOB): TOB;
begin
    Result := ComsxEdInit(Action, Param, RequestTOB)
end;

procedure InitApplication;
begin
  ddWriteLN('PGIComsxAPI loading...');
end;


function ComsxEdInit(Action, Param : string ; RequestTOB : TOB) : TOB;
var
 ResponseTOB : TOB ;
 PExport     : TExport;
 PImport     : TImportCom;
begin
 ResponseTOB := TOB.Create('', nil, -1) ;
 if not ResponseTOB.FieldExists ('ERROR') then
    ResponseTOB.AddChampSupValeur ('ERROR', '');

 // A OPTIMISER : on devrait recharger systématiquement les paramsocs utilisés
 if ((Action = 'IMPORT') or (Action = 'EXPORT')) then  ChargeTOBSOC;

 if V_GedFiles = nil then
 begin
    if V_PGI.RunFromLanceur then
      InitializeGedFiles(V_PGI.DefaultSectionDbName, nil)
    else
      InitializeGedFiles(V_PGI.DbName, nil);
 end;

 if Action='IMPORT' then
  begin
      PImport:= nil;
      if LanceImport (RequestTOB, '' , nil, PImport) then
          DDWriteLn ('Ok import') ;
      ResponseTOB.Dupliquer(RequestTOB, True, True);
  end
 else
 if Action='EXPORT' then
  begin

       PExport := nil;
       if LanceExport(RequestTOB, '', nil, PExport) then
          DDWriteLn ('OK Export') ;
       ResponseTOB.Dupliquer(RequestTOB, True, True);
  end
  else
  if Action='ZIPPE' then
  begin
       PExport := nil;
       if LanceZippeFile(RequestTOB, '', nil, PExport) then
          DDWriteLn ('OK Zippe') ;
       ResponseTOB.Dupliquer(RequestTOB, True, True);
  end
  else
  if Action='SUPPDIR' then
     LanceRemoveDirectory(RequestTOB.GetValue('APPLICATION'))
  else
  if Action='CREATEDIR' then
     LanceCreatDir(RequestTOB.GetValue('APPLICATION'));
  Result := ResponseTOB;

end;



end.

