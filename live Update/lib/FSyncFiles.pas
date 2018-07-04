unit fSyncFiles;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,
  UControleUAC,
  UfronctionBase,
  ShellAPI;

type
  TFSyncFiles = class(TForm)
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    exePath : string;

  public
    { Déclarations publiques }
  end;

var
  FUpdUpdate: TFSyncFiles;

implementation

{$R *.dfm}

procedure TFSyncFiles.FormCreate(Sender: TObject);
begin
  exePath := ExtractFilePath(Application.exename);
end;

procedure TFSyncFiles.FormDestroy(Sender: TObject);
begin
//
end;

procedure TFSyncFiles.FormShow(Sender: TObject);
begin
//
  ExeRunning('LSEClientMaj',true);
  if UpdateUpdProducts then
  begin
    RegenConfFiles;
  end;
  Sleep(2000);
  exePath := ExtractFilePath(Application.exename);
  //
  ShellExecute(Handle, 'runas', PansiChar(IncludeTrailingBackslash (exePath)+'LSEClientMaj.exe'), nil, nil, SW_SHOWNORMAL);
  close;
end;

end.
