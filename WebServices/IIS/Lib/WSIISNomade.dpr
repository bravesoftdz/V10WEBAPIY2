library WSIISNomade;

uses
  ActiveX,
  ComObj,
  WebBroker,
  ISAPIThreadPool,
  ISAPIApp,
  WSIISLSENomade in 'WSIISLSENomade.pas' {WebModule1: TWebModule},
  WSIISLSENomadeImpl in 'WSIISLSENomadeImpl.pas',
  WSIISLSENomadeIntf in 'WSIISLSENomadeIntf.pas',
  WSNomadeDecl in 'WSNomadeDecl.pas',
  Udefinitions in '..\..\Commun\Lib\Udefinitions.pas',
  Ulog in 'Ulog.pas',
  USql in 'USql.pas';

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
