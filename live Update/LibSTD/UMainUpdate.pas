unit UMainUpdate;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, URegistry, Registry,UCommonFuncs,Uconstantes,UEncryptage,Ulog,UsystemInfo,ShellAPI,
  CoolTrayIcon;

type

  K2KThread = Class(TThread)
  private
     ProgressBar : TProgressBar ;
  protected
     procedure Execute ; override ;
  public
     constructor Create (CreateSuspended:Boolean ; PBar : TProgressBar);
  end;

  TfMainUpdate = class(TForm)
    LNAMESPACE: TLabel;
    LFILE: TLabel;
    PGBRANCHE: TProgressBar;
    PGFile: TProgressBar;
    TrayIcon: TCoolTrayIcon;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    isOneTreated : Boolean;
    Th : K2KThread ;
    { Déclarations privées }
    TheInstallDirCegid : string;
    TheInstallDirLSE : string;
    TheProgramData : string;
    TheInstallDirSTD : string;
    TheInstallDirUPD : string;
    TheShare,TheRepository,TheRepoExeCegid,TheRepoExeLSe,TheRepoData,TheRepoStd,TheRepoUpd : string;
    RepertStockage : string;
    TheBatchCmd : string;

    procedure TraiteDir (NameSpace : string);
    function TraitDetail (NameSpace : string; TTL,TTS : TProductList) : boolean;
    procedure StartDefile;
    procedure StopDefile;
    procedure MoveDefile;
  public
    { Déclarations publiques }
    procedure LanceTraitement;
  end;


implementation
uses UControleUAC,UfronctionBase;

{$R *.dfm}

{ K2KThread }

constructor K2KThread.Create(CreateSuspended: Boolean; PBar: TProgressBar);
begin
  inherited Create(CreateSuspended);
  Priority := tpNormal	;
  FreeOnTerminate := False ;
  ProgressBar := PBar ;
end;

procedure K2KThread.Execute;
begin
  if ProgressBar = Nil then exit ;
  ProgressBar.Position := ProgressBar.Min ;
  while (not Terminated) do
  begin
    if ProgressBar.Position < ProgressBar.Max then
       ProgressBar.Position := ProgressBar.Position + 1
    else
       ProgressBar.Position := ProgressBar.Min ;
    ProgressBar.Invalidate ;
    Sleep(50);
  end;
  ProgressBar.Position := ProgressBar.Max ;
end;

procedure TfMainUpdate.FormCreate(Sender: TObject);
var Treg : TRegistry;
begin
  isOneTreated := False;
  TheShare := '';
  TheRepository := '';
  TheRepoData := '';
  TheRepoSTD := '';
  TheInstallDirCegid := '';
  TheInstallDirLSE := '';
  TheInstallDirSTD := '';
  TheInstallDirUPD := '';
  RepertStockage:= GetInfoStocke ('ServerRepository');

  Treg := TRegistry.Create;
  TRY
    with Treg do
    begin
      RootKey := HKEY_LOCAL_MACHINE;
      if not OpenKey ('SOFTWARE\Wow6432Node\Cegid\Cegid Business',False) then Exit;
      TheShare := ReadString('DIRCOPY');
      TheInstallDirCegid := ReadString('INSTALLDIR');
      TheInstallDirLSE := ReadString('INSTDIRLSE');
      TheInstallDirSTD := 'C:\PGI00\STD';
      CloseKey;
      if not OpenKey ('SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion',False) then Exit;
      TheInstallDirUPD := IncludeTrailingBackslash(readString('ProgramFilesDir'))+'CEGID\LSE Live Update\APP';
      CloseKey;
      if not OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',false) then Exit;
      TheProgramData := ReadString('Common AppData');
    end;
  finally
    Treg.Free;
  end;

  if TheShare = '' then Exit;
  if TheProgramData = '' then Exit;
  //
  if not DirectoryExists (TheShare) then
  begin
    if not CreateDir (TheShare) then Exit;
  end;
  //
  TheRepository := IncludeTrailingBackslash (TheShare)+'REPOSITORY';
  if not DirectoryExists (TheRepository) then
  begin
    if not CreateDir (TheRepository) then Exit;
  end;
  TheRepoExeCEGID := IncludeTrailingBackslash(TheRepository)+'CEGID';
  if not DirectoryExists (TheRepoExeCEGID) then
  begin
    if not CreateDir (TheRepoExeCEGID) then Exit;
  end;
  TheRepoExeLSE := IncludeTrailingBackslash(TheRepository)+'LSE';
  if not DirectoryExists (TheRepoExeLSE) then
  begin
    if not CreateDir (TheRepoExeLSE) then Exit;
  end;
  TheRepoData := IncludeTrailingBackslash(TheRepository)+'DATA';
  if not DirectoryExists (TheRepoData) then
  begin
    if not CreateDir (TheRepoData) then Exit;
  end;
  TheRepoSTD := IncludeTrailingBackslash(TheRepository)+'STD';
  if not DirectoryExists (TheRepoSTD) then
  begin
    if not CreateDir (TheRepoSTD) then Exit;
  end;
  TheRepoUPD := IncludeTrailingBackslash(TheRepository)+'UPDATE';
  if not DirectoryExists (TheRepoUPD) then
  begin
    if not CreateDir (TheRepoUPD) then Exit;
  end;

  Th := K2KThread.Create(True,PGFile) ;

  
end;

procedure TfMainUpdate.FormShow(Sender: TObject);
begin
//
end;


procedure TfMainUpdate.LanceTraitement;
var  ExeLance : string;
begin

  if TheShare = '' then close;
  if RepertStockage = '' then close;
  GetServerFiles (RepertStockage);
  //
  ExeLance := '"'+IncludeTrailingBackslash(TheInstallDirUPD)+'\LSEClientMaj.exe"';
  StartDefile;
  TraiteDir ('UPDATE');
  TraiteDir ('CEGID');
  TraiteDir ('LSE');
  TraiteDir ('DATA');
  TraiteDir ('STD');
  //
  if isOneTreated then
  begin
    Sleep(2000);
    RegenConfFiles;
    TrayIcon.ShowBalloonHint('LSE LiveUpdate Information','Vos produits sont à jour',bitInfo, 10);
    Sleep(2000);
  end;
  ShellExecute(Handle, 'runas', PansiChar(ExeLance), nil, nil, SW_SHOWNORMAL);
  close;
end;

procedure TfMainUpdate.StartDefile;
begin
  if not PGFile.Visible then
  begin
     PGFile.Visible := True ;
     PGFile.Position := PGFile.min;
  end;
  Th.Resume ;
end;

procedure TfMainUpdate.StopDefile;
begin
  if PGFile.Visible then
     PGFile.Visible := False ;
  Th.Terminate ;
end;

procedure TfMainUpdate.MoveDefile;
begin
    if PGFile.Position < PGFile.Max then
       PGFile.Position := PGFile.Position + 1
    else
       PGFile.Position := PGFile.Min ;

   Self.Invalidate;
   Application.ProcessMessages;
end;

function TfMainUpdate.TraitDetail(NameSpace: string; TTL,TTS: TProductList) : boolean;
var II : Integer;
    From,TODest,ToTemp,TheName,cmd  : string;
    ServerFile,TempFile,DestFile : string;
    LT,DT : TElement;
    OneShot : Boolean;
begin
  OneShot := True;
  result := false;
  LNAMESPACE.Caption := NameSpace; LNAMESPACE.Visible := True;
  PGBRANCHE.Visible := True;  PGBRANCHE.StepIt;
  Self.paint;
  From := IncludeTrailingBackslash(IncludeTrailingBackslash(RepertStockage)+NameSpace);
  ToTemp := IncludeTrailingBackslash(IncludeTrailingBackslash(TheRepository)+NameSpace);
  if NameSpace = 'LSE' then TODest := IncludeTrailingBackslash(TheInstallDirLSE)+'APP'
  else if NameSpace = 'CEGID' then ToDest := IncludeTrailingBackslash(TheInstallDirCegid)+'APP'
  else if NameSpace = 'STD' then ToDest := TheInstallDirSTD 
  else if NameSpace = 'DATA' then ToDest := IncludeTrailingBackslash(TheProgramData)+'CEGID'
  else if NameSpace = 'UPDATE' then ToDest := TheInstallDirUPD;

  for II := 0 to TTL.Count -1 do
  begin
    LT := TTL.items[II];
    TheName := LT.Element;
    DT := TTS.Find(TheName);
    if DT <> nil then
    begin
      if DT.Signature <> LT.Signature then
      begin
        result := True;
        if not isOneTreated then isOneTreated := True;
        LFILE.Caption := TheName;
        LFILE.Visible := True;
        Self.Refresh;
        //
        ServerFile := IncludeTrailingBackslash(From)+TheName;
        TempFile := IncludeTrailingBackslash(ToTemp)+TheName;
        DestFile := IncludeTrailingBackslash(ToDest)+TheName;
        ShellExecAndWait('c:\windows\system32\xcopy', '"'+ServerFile+'"  "'+TempFile+'" /C/Y',SW_HIDE);
        ShellExecAndWait('c:\windows\system32\xcopy', '"'+Tempfile+'"  "'+DestFile+'" /C/Y',SW_HIDE);
        //
        LFILE.Visible := false;
      end;
    end;
  end;
  LNAMESPACE.Visible := false;
  PGBRANCHE.Visible := false;
  LFILE.Visible := false;
end;

procedure TfMainUpdate.TraiteDir(NameSpace: string);
var TTL,TTS : TProductList;
    LLOC,LTemp : string;
    ATraiter : boolean;
begin

  TTL := TProductList.Create;
  TTS := TProductList.Create;
  TRY
    if NameSpace='CEGID' then
    begin
      LLOc := IncludeTrailingBackslash (TheRepository)+CONFCEGIDFILE ;
      if FileExists(LLOC) then
      begin
        ChargeListeProducts (LLOc,TTL);
        LTemp := IncludeTrailingBackslash (GetTempDir)+CONFCEGIDFILE;
        if FileExists (LTemp) then
        begin
          ChargeListeProducts (LTemp,TTS);
          ATraiter := TraitDetail ('CEGID',TTL,TTS);
          if ATraiter then DeleteFile(LLOC);
        end;
      end;
    end else if NameSpace='LSE' then
    begin
      LLOc := IncludeTrailingBackslash (TheRepository)+CONFLSEFILE;
      if FileExists(LLOC) then
      begin
        ChargeListeProducts (LLOc,TTL);
        LTemp := IncludeTrailingBackslash (GetTempDir)+CONFLSEFILE;
        if FileExists (LTemp) then
        begin
          ChargeListeProducts (LTemp,TTS);
          ATraiter := TraitDetail ('LSE',TTL,TTS);
          if ATraiter then DeleteFile(LLOC);
        end;
      end;
    end else if NameSpace='STD' then
    begin
      LLOc := IncludeTrailingBackslash (TheRepository)+CONFSTDFILE;
      if FileExists(LLOC) then
      begin
        ChargeListeProducts (LLOc,TTL);
        LTemp := IncludeTrailingBackslash (GetTempDir)+CONFSTDFILE;
        if FileExists (LTemp) then
        begin
          ChargeListeProducts (LTemp,TTS);
          ATraiter := TraitDetail ('STD',TTL,TTS);
          if ATraiter then DeleteFile(LLOC);
        end;
      end;
    end else if NameSpace ='DATA' then
    begin
      LLOc := IncludeTrailingBackslash (TheRepository)+PARAMUPDATE ;
      if FileExists(LLOC) then
      begin
        ChargeListeProducts (LLOc,TTL);
        LTemp := IncludeTrailingBackslash (GetTempDir)+PARAMUPDATE;
        if FileExists (LTemp) then
        begin
          ChargeListeProducts (LTemp,TTS);
          Atraiter := TraitDetail ('DATA',TTL,TTS);
          if ATraiter then DeleteFile(LLOC);
        end;
      end;
    end else if NameSpace ='UPDATE' then
    begin
      LLOc := IncludeTrailingBackslash (TheRepository)+CONFUPDFILE ;
      if FileExists(LLOC) then
      begin
        ChargeListeProducts (LLOc,TTL);
        LTemp := IncludeTrailingBackslash (GetTempDir)+CONFUPDFILE;
        if FileExists (LTemp) then
        begin
          ChargeListeProducts (LTemp,TTS);
          Atraiter := TraitDetail ('UPDATE',TTL,TTS);
          if ATraiter then DeleteFile(LLOC);
        end;
      end;
    end;
  FINALLY
    TTL.free;
    TTS.free;
  END;
end;

procedure TfMainUpdate.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not Th.Suspended then
     Th.Terminate ;
  Th.Free ;

end;

end.
