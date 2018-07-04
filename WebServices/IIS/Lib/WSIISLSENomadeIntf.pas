{ Interface invocable IWSIISLSENomade }

unit WSIISLSENomadeIntf;

interface

uses InvokeRegistry, Types, XSBuiltIns, WSNomadeDecl;

type

  { Les interfaces invocables doivent dériver de IInvokable }
  IWSIISLSENomade = interface(IInvokable)
  ['{F4541998-87D2-4994-9035-69121E22E481}']
    { Les méthodes de l'interface invocable ne doivent pas utiliser la convention }
    { d'appel par défaut ; stdcall est conseillée }
    function IsValideConnect ( ParamIn : TWSUserParam) : WideString; stdcall;
    function GetAppels (identificateur : WideString; Depuis : Widestring) : WideString; stdcall;
    function GetAppel (identificateur : WideString; CodeAppel : Widestring) : WideString; stdcall;
  end;

implementation

initialization
  { Les interfaces invocables doivent être recensées }
  InvRegistry.RegisterInterface(TypeInfo(IWSIISLSENomade));

end.
 