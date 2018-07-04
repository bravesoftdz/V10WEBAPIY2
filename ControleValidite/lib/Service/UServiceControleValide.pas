unit UServiceControleValide;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, SvcMgr, Dialogs,QExtCtrls,
  ExtCtrls, Menus,ADODB,UConnectionData,Registry;

type

  TModeAutorise = (TmaOK,TmaParam,TmaBloque);

  TLSEAutorise = class(TService)
    TT: TTimer;
    procedure ServiceExecute(Sender: TService);
    procedure ServiceContinue(Sender: TService; var Continued: Boolean);
    procedure ServicePause(Sender: TService; var Paused: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceShutdown(Sender: TService);
    procedure TTTimer(Sender: TObject);
    procedure ServiceAfterInstall(Sender: TService);
  private
    fStatusProtec : TModeAutorise;
    fClient : string;
    fAdresse : string;
    flastDate : TDateTime;
    fCounter : Integer;
    fCRC : double;
    fNomLog : string;
    { Déclarations privées }
    procedure GetInfosConnection;
    procedure VerifSerialisations;
    procedure RecupSerialisations;
  public
    function GetServiceController: TServiceController; override;
    { Déclarations publiques }
  end;

var
  LSEAutorise: TLSEAutorise;

implementation

uses DateUtils;

{$R *.DFM}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  LSEAutorise.Controller(CtrlCode);
end;

function TLSEAutorise.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TLSEAutorise.ServiceExecute(Sender: TService);
begin
  TT.Enabled := true;
  while not Terminated do ServiceThread.ProcessRequests(True);
  TT.Enabled := false;
end;

procedure TLSEAutorise.ServiceContinue(Sender: TService;var Continued: Boolean);
begin
  Continued := true;
end;

procedure TLSEAutorise.ServicePause(Sender: TService; var Paused: Boolean);
begin
	Paused := true;
end;

procedure TLSEAutorise.ServiceStart(Sender: TService;var Started: Boolean);
begin
  fNomLog := '';
  CreateInfosregistry ;
  GetInfosConnection;
  if fStatusProtec = TmaOk then RecupSerialisations;
  VerifSerialisations;
  Started := True;
end;

procedure TLSEAutorise.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
	Stopped := True;
end;

procedure TLSEAutorise.ServiceShutdown(Sender: TService);
begin
  //
end;

procedure TLSEAutorise.TTTimer(Sender: TObject);
(*var QQ : TadoQuery;
		MDBData : string;
*)
begin
  if IsSameDay(Now,flastDate) then Exit;

  (*
  MDBData := IncludeTrailingPathDelimiter(GetAppData)+'CEGID\LSE Controle\Controle.mdb';
  // On controle si on doit se remettre a jour ou pas par rapport aux dates des fonctionnalités
  // si liaison possible --> on récup les infos d'autorisations et on mets à jour les infos
  QQ := TADOQuery.Create(self);
  QQ.ConnectionString :=ACCESConnectionString(MDBDATA,'Admin');
  QQ.SQL.Add('SELECT * FROM DETAILS');
  QQ.Open;
  while not QQ.Eof do
  begin
    //
  end;
  QQ.Close;
  QQ.Free;
  *)
  //
  if fStatusProtec = TmaOk then RecupSerialisations;
  VerifSerialisations;
end;

procedure TLSEAutorise.VerifSerialisations;
begin
end;

procedure TLSEAutorise.GetInfosConnection;
begin
  fStatusProtec := TmaParam;

  fClient := GetInfosRegString('CodeClient');
  if fClient = '' then
  begin
  	LogMessage('Merci de définir les paramètres de connexion',EVENTLOG_WARNING_TYPE);
    Exit;
  end;
	fAdresse := GetInfosRegString('LSEAdress');
  if fAdresse = '' then
  begin
  	LogMessage('Merci de définir les paramètres de connexion',EVENTLOG_WARNING_TYPE);
    Exit;
  end;
	flastDate := GetInfosRegDate('LastDate');
  fcounter := getInfosRegInteger ('Counter');
  fCRC := getInfosRegdWord ('CRC');
  fStatusProtec := TmaOK;
  if flastDate <> StrToDate('01/01/1900') then
  begin
    if CalculCRC (fLastDate,fCounter) <> fCRC then
    begin
      LogMessage('Paramètres incorrects. BAD CRC',EVENTLOG_WARNING_TYPE);
      fStatusProtec := TmaBloque;
      Exit;
    end;
  end;
end;

procedure TLSEAutorise.ServiceAfterInstall(Sender: TService);
begin
  CreateInfosregistry ;
end;

procedure TLSEAutorise.RecupSerialisations;
begin
	
end;

end.
