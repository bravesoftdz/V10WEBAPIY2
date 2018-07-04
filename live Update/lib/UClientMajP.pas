unit UClientMajP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, CoolTrayIcon, Menus, Buttons, ExtCtrls, StdCtrls, ComCtrls,
  FileCtrl, Spin, ImgList,UControleUAC, Mask, UEncryptage, UbrowseRepert,Uregistry,
  UConstantes,USystemInfo,commctrl,ShellApi,UfronctionBase,UCommonFuncs;

type
  TFlanceurMajLse = class(TForm)
    TrayIcon: TCoolTrayIcon;
    TimeUpdate: TTimer;
    ImageList7: TImageList;
    ImageList1: TImageList;
    PopupM: TPopupMenu;
    MnChangeParams: TMenuItem;
    N1: TMenuItem;
    MnFindUpdate: TMenuItem;
    N2: TMenuItem;
    MnExit: TMenuItem;
    ImpParams: TMenuItem;
    ExpParams: TMenuItem;
    N3: TMenuItem;
    procedure TrayIconStartup(Sender: TObject; var ShowMainForm: Boolean);
    procedure MnExitClick(Sender: TObject);
    procedure MnFindUpdateClick(Sender: TObject);
    procedure MnChangeParamsClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TimeUpdateTimer(Sender: TObject);
    procedure TrayIconDblClick(Sender: TObject);
    procedure TrayIconBalloonHintClick(Sender: TObject);
    procedure ExpParamsClick(Sender: TObject);
    procedure ImpParamsClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Déclarations privées }
    //
    TheListExe : TStringList;
    RepertStockage : string;
    Intervalle : integer;
    OutilsVisible : string;
    MajDispo : boolean;
    CegidInstallApp : string;
    LSEInstallApp : string;
    fWaitUpdate : boolean;
    exePath : string;
    MessageOkVersion : boolean;
    procedure IsMajDisponible ;
    procedure GetInfos;
    procedure SetInfoTimer;
    function UpdateVers: boolean;
    function GetConfigFile : boolean;
    procedure UpdateProducts;
  public
    { Déclarations publiques }
  end;

var
  FlanceurMajLse: TFlanceurMajLse;

implementation
uses RunElevatedSupport,Registry;

{$R *.dfm}

procedure TFlanceurMajLse.TrayIconStartup(Sender: TObject;var ShowMainForm: Boolean);
begin
	ShowmainForm := false;
end;

procedure TFlanceurMajLse.MnExitClick(Sender: TObject);
begin
	close;
end;

procedure TFlanceurMajLse.MnFindUpdateClick(Sender: TObject);
begin
	IsMajDisponible;
  if MajDispo then
  begin
    TrayIcon.ShowBalloonHint('LSE LiveUpdate Information','Une mise à jour des produits est disponible.Cliquez ici pour l''installer',bitInfo, 10);
    TrayIcon.CycleInterval := 150;
    TrayIcon.CycleIcons := true;
    fWaitUpdate := true;
  end else
  begin
    TrayIcon.ShowBalloonHint('LSE LiveUpdate Information','Vos produits sont à jour',bitInfo, 10);
  end;
end;

procedure TFlanceurMajLse.MnChangeParamsClick(Sender: TObject);
begin
  ShellExecute(Handle, 'runas', PansiChar(IncludeTrailingBackslash (exePath)+'SetParams.exe'), nil, nil, SW_SHOWNORMAL);
  while  ExeRunning('SetParams',false) do Sleep(100) ;
  GetInfos;
	SetInfoTimer;
end;


procedure TFlanceurMajLse.IsMajDisponible;
var ResultGet : boolean;
begin
  MajDispo := false;
  TrayIcon.IconList := ImageList7;
  TrayIcon.CycleInterval := 150;
  TrayIcon.CycleIcons := True;
  //
  ResultGet := GetConfigFile;
  if ResultGet then
  begin
    if ISVersionsDIff (TheListExe) then MajDispo := true;
  end;
  //
  TrayIcon.IconList := ImageList1;
  TrayIcon.CycleInterval := 1;
  TrayIcon.CycleIcons := false;
end;

procedure TFlanceurMajLse.FormCreate(Sender: TObject);
begin
  TheListExe := TStringList.Create;
  MessageOkVersion := false;
  exePath := ExtractFilePath(Application.exename); 
  fWaitUpdate := false;
  MajDispo := false;
	CegidInstallApp := IncludeTrailingBackSlash(GetCegidInstallFolder)+'APP';
	LSEInstallApp := IncludeTrailingBackSlash(GetLSEInstallFolder)+'APP';
  //
  TrayIcon.IconList := ImageList1;
  TrayIcon.CycleInterval := 1;
  TrayIcon.CycleIcons := false;
  GetInfos;
	SetInfoTimer;
//  SetButtonElevated(MnChangeParams.Handle);
end;

function  TFlanceurMajLse.UpdateVers : boolean;
var IsUpdateUpd : Boolean;

  function OneExeRunning (var UpdToUpdate : Boolean) : boolean;
  var II : Integer;
      Theexe : string;
  begin
    Result := false;
    UpdToUpdate := false;
    for II := 0 to TheListExe.Count -1 do
    begin
      Theexe := ExtractFileName(TheListExe.Strings [II]);
      Theexe := Copy(Theexe,1,Length(Theexe)-4);
      if ExeRunning(TheExe,false) then
      begin
        if TheExe = 'LseClientMaj' then UpdToUpdate := True;
        Result := true;
        exit;
      end;
    end;
  end;

begin
	result := false;
  if OneExeRunning (IsUpdateUpd) then
  begin
    TrayIcon.ShowBalloonHint('LSE LiveUpdate Information','Merci de sortir des produits à mettre à jour',bitInfo, 10);
    exit;
  end else
  begin
    UpdateProducts;
    result := true;
  end;
end;

procedure TFlanceurMajLse.GetInfos;
var RR : string;
begin
	RepertStockage:= GetInfoStocke ('ServerRepository');
  RR := getInfoStocke('Timer');
  if RR <> '' then Intervalle := StrToInt(RR) else Intervalle := 60;
  RR := GetInfoStocke ('OutilsVisible');
  if RR <> '' then OutilsVisible := RR else OutilsVisible := 'OUI';
  if OutilsVisible = 'NON' then
  begin
    N1.visible := false;
    N2.visible := false;
    MnChangeParams.visible := false;
    ExpParams.visible := false;
    ImpParams.visible := false;
  end;
end;

procedure TFlanceurMajLse.TimeUpdateTimer(Sender: TObject);
begin
	IsMajDisponible;
  if MajDispo then
  begin
    TrayIcon.ShowBalloonHint('LSE LiveUpdate Information','Une mise à jour des produits est disponible.Cliquez ici pour l''installer',bitInfo, 10);
    TrayIcon.CycleInterval := 150;
    TrayIcon.CycleIcons := true;
    fWaitUpdate := true;
  end;
end;

procedure TFlanceurMajLse.SetInfoTimer;
begin
  TimeUpdate.Interval := (60000 * Intervalle);
  TimeUpdate.Enabled := true;
end;

procedure TFlanceurMajLse.TrayIconDblClick(Sender: TObject);
begin
  MnFindUpdateClick(self);
end;


function TFlanceurMajLse.GetConfigFile: boolean;
begin

  result := false;
  if RepertStockage = '' then
  begin
    exit;
  end;
  if not DirectoryExists(RepertStockage) then
  begin
    TrayIcon.ShowBalloonHint('LSE LiveUpdate Information','Le répertoire de stockage du repository est incorrect',bitInfo, 10);
    exit;
  end;

  GetServerFiles (RepertStockage);
  //----
  result := true;
end;

procedure TFlanceurMajLse.TrayIconBalloonHintClick(Sender: TObject);
begin
  if fWaitUpdate then UpdateVers;
end;


procedure TFlanceurMajLse.ExpParamsClick(Sender: TObject);
begin
  ShellExecute(Handle, 'runas', PansiChar(IncludeTrailingBackslash (exePath)+'ExportParams.exe'), nil, nil, SW_SHOWNORMAL);
  while  ExeRunning('ExportParams',false) do Sleep(100) ;
	SetInfoTimer;
end;

procedure TFlanceurMajLse.ImpParamsClick(Sender: TObject);
begin
  ShellExecute(Handle, 'runas', PansiChar(IncludeTrailingBackslash (exePath)+'ImportParams.exe'), nil, nil, SW_SHOWNORMAL);
  while  ExeRunning('ImportParams',false) do Sleep(100) ;
  GetInfos;
	SetInfoTimer;
end;

procedure TFlanceurMajLse.FormDestroy(Sender: TObject);
begin
  TheListExe.free;

end;

procedure TFlanceurMajLse.UpdateProducts;
var exefile  :string;
begin
  TrayIcon.IconList := ImageList7;
  TrayIcon.CycleInterval := 150;
  TrayIcon.CycleIcons := True;
  exefile :=  IncludeTrailingBackslash(GetTempDir)+'UpdVersions.exe';
  if FileExists(ExeFile) then DeleteFile(exefile);
  // --
  CopyFile(PansiChar(IncludeTrailingBackslash (exePath)+'UpdVersions.exe'),PAnsiChar(exefile),false);
  //
  ShellExecute(Handle, 'runas', PansiChar(Exefile), nil, nil, SW_SHOWNORMAL);
  PostMessage(Handle, WM_CLOSE, 0, 0); //on sort de la saisie

  {
  // ---
  fWaitUpdate := false;
  TrayIcon.IconList := ImageList1;
  TrayIcon.CycleInterval := 1;
  TrayIcon.CycleIcons := false;
  TrayIcon.ShowBalloonHint('LSE LiveUpdate Information','Les produits ont été mis à jour',bitInfo, 10);
  TheListExe.clear;
  }
end;

end.
