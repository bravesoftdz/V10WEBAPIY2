unit udbssDispatch;

interface

uses
  classes,
  UTOB, Hent1;

procedure InitApplication;

function DispatchEx(Action, Param : String ; RequestTOB : TOB; var ResultStream : TMemoryStream) : TOB ; stdcall;

implementation

uses
  windows, sysutils,
  eSession, eDispatch,
  BasesUtil, MsdeUtil, OutilWA ;


procedure InitApplication;
begin
  ddWriteLN('---> dbssPGI loading...');
end;

function GBDSSDispatch(Action, Param : string; RequestTOB : TOB;var ResultStream : TMemoryStream) : TOB ;
var
  ResponseTOB : TOB;
begin
ResponseTOB := TOB.Create('le test', nil, -1);
ddWriteLN('dbssPGI doing '+Action+', '+Param+' ...');

if v_pgi.OkOuvert then
  begin
  try
    if Action = 'MSDE' then
      begin
      OutilMSDE.ActionMSDE(RequestTOB,ResponseTOB) ;
      end else
    if Action = 'BASE' then
      begin
      OutilBase.ActionBase(RequestTOB,ResponseTOB) ;
      end ;
  except
    on E:Exception do ResponseTOB.AddChampSupValeur('ERROR', 'Erreur: '+E.Message)
    end ;
  end
else
  begin
  ResponseTOB.AddChampSupValeur('ERROR','pgi pas ok ouvert') ;
  end ;

Result := ResponseTOB;
end;

function DispatchEx(Action, Param : String ; RequestTOB : TOB; var ResultStream : TMemoryStream) : TOB ; stdcall;
begin
  result := GBDSSDispatch(Action, Param, RequestTOB, ResultStream);
end;

function GBDSSProcCalcEdt(sf, sp : string) : variant;
begin
  // init, au cas où on n'a pas la fonction
  result := UnAssigned; // et pas autre chose
  // traitement ici :
  // ...
  // result := ''; si et seulement si le champ doit être blanc
end;


end.

