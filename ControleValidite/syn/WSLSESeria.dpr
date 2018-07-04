library WSLSESeria;

uses
  ActiveX,
  ComObj,
  WebBroker,
  ISAPIThreadPool,
  ISAPIApp,
  IISLSESeria in '..\lib\WebServiceClient\IISLSESeria.pas' {WebModule1: TWebModule},
  IISLSESeriaImpl in '..\lib\WebServiceClient\IISLSESeriaImpl.pas',
  IISLSESeriaIntf in '..\lib\WebServiceClient\IISLSESeriaIntf.pas',
  UAccesDatabase in '..\lib\Commun\UAccesDatabase.pas';

{$R *.res}

exports
  GetExtensionVersion,
  HttpExtensionProc,
  TerminateExtension;

begin
  CoInitFlags := COINIT_MULTITHREADED;
  Application.Initialize;
  Application.CreateForm(TWebModule1, WebModule1);
  Application.Run;
end.
