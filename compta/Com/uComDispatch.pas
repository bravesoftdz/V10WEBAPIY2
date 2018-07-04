unit uComDispatch;

interface

uses
  classes, Windows,
  UTOB,
  HMsgBox,
  HEnt1, ent1, inifiles;

procedure InitApplication;
function ComDispatch(Action, Param : String ; RequestTOB : TOB) : TOB ;
function myProcCalcEdt(sf, sp : string) : variant;

implementation

uses
  sysutils,
//DBTables, HEnt1,HCtrls,
  eSession, eDispatch, UExportCom, UImportCom;
//var
//FichierIE : TextFile;


procedure InitApplication;
begin

  (*while true do
  begin end;*)

  ddWriteLN('PGICOMSXAPI loading...');
//  AssignFile(FichierIE, 'C:\CWS\Scripts\XXX.txt') ;
//  Rewrite(FichierIE) ;

  ProcDispatch := ComDispatch;
//  ProcCalcEdt := myProcCalcEdt;
end;

function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;
begin
  DDWriteLn ('AAAAAAAAAAAAAAA : ' + ACTION);
  Action := UpperCase(Action);
  result := ComDispatch(Action, Param, RequestTOB);
end;

function ComDispatch(Action, Param : String ; RequestTOB : TOB) : TOB ;
var
stName, stDossier : string;
begin

  InitLaVariableHalley;
  ChargeMagHalley;

  result := TOB.Create('le result', nil, -1);
  result.AddChampSupValeur('ERROR', ''); // créer au moins (au plus ?) ce champ
  // result doit être virtuelle, mais elle peut avoir des filles réelles
  // traitement ici :
  // ...

//  Writeln (FichierIE, Action);
//  CloseFile(FichierIE);


  // Exemple
  if Action='IMPORT' then
  begin
       if LanceImport (RequestTOB, '' , nil) then
          DDWriteLn ('Ok import')
       else
          result.PutValue('ERROR', 'ERROR');
  end ;

  if Action='EXPORT' then
  begin
//  RequestTOB.SaveToFile('C:\CWS\Scripts\journal.txt',True,True,True);
    if LanceExport(RequestTOB, '', nil) then
       DDWriteLn ('OK Export')
    else
          result.PutValue('ERROR', 'ERROR');
  end ;

  // ne pas abuser des ERROR <> '' : elles sont transmises à l'administrateur...
end;

function myProcCalcEdt(sf, sp : string) : variant;
begin
  // init, au cas où on n'a pas la fonction
  result := UnAssigned; // et pas autre chose
  // traitement ici :
  // ...
  // result := ''; si et seulement si le champ doit être blanc
end;


end.

