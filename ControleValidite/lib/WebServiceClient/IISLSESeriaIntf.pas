{ Interface invocable IIISLSESeria }

unit IISLSESeriaIntf;

interface

uses InvokeRegistry, Types, XSBuiltIns;

type


  TSeriaDB = class(TRemotable)
  private
    FClient : AnsiString;
    FMethode : AnsiString;
    FDB : AnsiString;
    FDBID: integer;
    FDateAuth: TDateTime;
  published
    property ClientName: AnsiString read FClient write FClient;
    property Methode : AnsiString read FMethode write FMethode;
    property DB: AnsiString read FDB write FDB;
    property DBID: Integer read FDBID write FDBID;
    property DateAuth: TDateTime read FDateAuth write FDateAuth;
  end;

  TSeriaDBArray = array of TSeriaDB;

  { Les interfaces invocables doivent dériver de IInvokable }
  IIISLSESeria = interface(IInvokable)
  ['{684CFF55-C7B7-45D1-BFE8-39325DD88BF1}']

    { Les méthodes de l'interface invocable ne doivent pas utiliser la convention }
    { d'appel par défaut ; stdcall est conseillée }
    function GetSeriaTiers(const CodeTiers : String): TSeriaDBArray; stdcall;
  end;

implementation

initialization
  { Les interfaces invocables doivent être recensées }
  InvRegistry.RegisterInterface(TypeInfo(IIISLSESeria));

end.
 