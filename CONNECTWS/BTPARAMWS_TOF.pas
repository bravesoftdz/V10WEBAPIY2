{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMWS ()
Mots clefs ... : TOF;BTPARAMWS
*****************************************************************}
Unit BTPARAMWS_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , mul
  , FE_main
  , uTob
  {$ELSE EAGLCLIENT}
  , eMul
  , MaineAGL
  {$ENDIF EAGLCLIENT}
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HEnt1
  , HMsgBox
  , UTOF
  , UConnectWSCEGID
  , HTB97
  ;

function BTLanceFicheParamWSCegid(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_BTPARAMWS = Class (TOF)
  private
    ServerName : THEdit;
    Port       : THEdit;
    FolderName : THValComboBox;
    WSCegid    : TconnectCEGID;
    TOBFolderName : TOB;
    LoadFolderNameResponse : WideString;
    SearchFolders : TToolbarButton97;

    procedure ChangeValues(Sender : TObject);
    procedure SetFloderList(Sender : TObject);

    procedure TstWApi_Onclick(Sender : TObject);

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses
  ParamSoc
  , ed_Tools
  , UConnectWSConst
  , uWSDataService
//*********** POUR TEST
  , uExecuteService
  , CommonTools
//*********** Fin POUR TEST
  ;

function BTLanceFicheParamWSCegid(Nat, Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (Nat <> '') and (Cod <> '') then
  begin
    V_PGI.ZoomOle := True;
    Result := AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
    V_PGI.ZoomOle := False;
  end else
    Result := '';
end;

procedure TOF_BTPARAMWS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMWS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMWS.OnUpdate ;
begin
  Inherited ;
  SetParamSoc(WSCDS_SocServer, ServerName.Text);
  SetParamSoc(WSCDS_SocNumPort, Port.Text);
  SetParamSoc(WSCDS_SocCegidDos, FolderName.Value);
end ;

procedure TOF_BTPARAMWS.OnLoad ;
begin
  Inherited ;
  ServerName.Text  := TGetParamWSCEGID.GetPSoc(wspsServer);
  Port.Text        := TGetParamWSCEGID.GetPSoc(wspsPort);
  FolderName.Value := TGetParamWSCEGID.GetPSoc(wspsFolder); 
  ChangeValues(nil);
  SetFloderList(nil);
end ;

procedure TOF_BTPARAMWS.OnArgument (S : String ) ;
begin
  Inherited ;
  ServerName    := THEdit(GetControl('SERVEUR'));
  Port          := THEdit(GetControl('PORT'));
  FolderName    := THValComboBox(GetControl('DOSSIER'));
  SearchFolders := TToolbarButton97(GetControl('SEARCHFOLDERS'));
  TOBFolderName := TOB.Create('_FOLDERNAME', nil, -1);
  WSCegid       := TconnectCEGID.create;
  SearchFolders.OnClick := SetFloderList;
  ServerName.OnChange   := ChangeValues;
  Port.OnChange         := ChangeValues;

  TToolbarButton97(GetControl('TSTWEBAPI')).OnClick := TstWApi_Onclick;

end ;

procedure TOF_BTPARAMWS.OnClose ;
begin
  Inherited ;
  FreeAndNil(TOBFolderName);
  WSCegid.Free;
end ;

procedure TOF_BTPARAMWS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMWS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMWS.ChangeValues(Sender: TObject);
begin
  SearchFolders.Enabled := ((ServerName.Text <> '') and (Port.Text <> ''));
  if FolderName.Items.Count > 0 then
  begin
    FolderName.Items.Clear;
    TOBFolderName.ClearDetail;
  end;
end;

procedure TOF_BTPARAMWS.SetFloderList(Sender : TObject);
{$IF not defined(APPSRV)}
var
  Cpt            : integer;
  TOBFolderNameL : TOB;
  FolderValue    : string;
  ItemIndex      : integer;
{$IFEND !APPSRV}
begin
  {$IF not defined(APPSRV)}
  if (ServerName.Text = '') or (Port.Text = '') then
  begin
    PGIError('Le nom du serveur et le port de communication doivent être renseignés', Ecran.Caption);
  end else
  begin
    InitMoveProgressForm(nil, 'Recherche des dossiers ...', Ecran.Caption, 0, False, True);
    try
      WSCegid.CEGIDServer := ServerName.Text;
      WSCegid.CEGIDPORT   := Port.Text;
      FolderValue         := TGetParamWSCEGID.GetPSoc(wspsFolder);
      { La liste des dossiers n'est pas encore montée }
      if FolderName.Items.Count = 0 then
      begin
        TOBFolderName.ClearDetail;
        WSCegid.GetDossiers(TOBFolderName, LoadFolderNameResponse);
        { Construit le combo avec le contenu de la TOB }
        if TOBFolderName.Detail.Count > 0 then
        begin
          ItemIndex := 0;
          for Cpt := 0 to pred(TOBFolderName.Detail.Count) do
          begin
            TOBFolderNameL := TOBFolderName.Detail[Cpt];
            FolderName.Items.Insert(Cpt, TOBFolderNameL.GetString('DESCRIPTION'));
            FolderName.Values.Insert(Cpt, TOBFolderNameL.GetString('ID'));
            if TOBFolderNameL.GetString('ID') = FolderValue then
              ItemIndex := Cpt;
          end;
          FolderName.ItemIndex := ItemIndex;          
        end;
      end;
    finally
      FiniMoveProgressForm;
    end;
  end;
  {$IFEND !APPSRV}
end;

procedure TOF_BTPARAMWS.TstWApi_Onclick;
var
  BTPY2Exec : TSvcSyncBTPY2Execute;
  AppName   : string;
begin
  { Test du service }
  AppName   := ExtractFilePath(Application.ExeName) + 'SvcSynBTPY2.exe';
  BTPY2Exec := TSvcSyncBTPY2Execute.Create;
  try
    BTPY2Exec.ApplicationName := AppName;
    BTPY2Exec.CreateObjects;
    try
      BTPY2Exec.InitApplication;
      try
       BTPY2Exec.ServiceExecute;
      finally
      end;
    finally
      BTPY2Exec.FreeObjects;
//      WriteLog(ssbylWindows, 'Fin d''exécution du service SvcSynBTPY2', 0);
    end;
  finally
    BTPY2Exec.Free;
  end;
end;
  
Initialization
  registerclasses ( [ TOF_BTPARAMWS ] ) ;
end.

