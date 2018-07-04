unit UNewCollab;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,UgestionINI, HTB97, ExtCtrls, StdCtrls, Hctrls,USQLServer,
  HSysMenu,Udefinitions;

type

  TfNewID = class(TForm)
    PBAS: TPanel;
    BFERME: TToolbarButton97;
    BVALIDE: TToolbarButton97;
    ID: TEdit97;
    LID: THLabel;
    LIBELLE: TEdit97;
    LLIBELLE: THLabel;
    LSERVER: THLabel;
    HLabel1: THLabel;
    LRESSOURCE: THLabel;
    MRESSOURCE: THValComboBox;
    USER: TEdit97;
    SERVER: THValComboBox;
    DATABASE: THValComboBox;
    RESSOURCE: TEdit97;
    HmTrad: THSystemMenu;
    BDELETE: TToolbarButton97;
    PASSWORD: TEdit97;
    LPASSWORD: THLabel;
    procedure BFERMEClick(Sender: TObject);
    procedure BVALIDEClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure IDExit(Sender: TObject);
    procedure SERVERChange(Sender: TObject);
    procedure DATABASEChange(Sender: TObject);
    procedure RESSOURCEChange(Sender: TObject);
    procedure BDELETEClick(Sender: TObject);
  private
    { Déclarations privées }
    fModeTrait : TActionFiche;
    fEnreg : TEnreg;
    fOkValide : Boolean;
    fSupprime : boolean;
    fgestionIni : TGestionINI;
    procedure  EnregToEcran;
    procedure  EcranToEnreg;
    procedure ChargeListeServer;
    procedure ChargeListeDatabases;
    procedure ChargeListeUsers;
    procedure SetEvents (Etat : Boolean);
    function GetUserAssocie : string;
    function VerifOK : Boolean;
  public
    { Déclarations publiques }
    property ModeTrait : TActionFiche read fModeTrait write fModeTrait;
    property Enreg : TEnreg read fEnreg write fEnreg;
    property OkValide : boolean read fOkValide write fOkValide;
    property GestionIni : TGestionINI read fgestionIni write fgestionIni; 
    property Supprime : boolean read fSupprime write fSupprime;
  end;


procedure AppelAjout (var NewEnreg : Tenreg; GestionIni : TGestionINI);
function AppelModif (ID : string; GestionIni : TGestionINI) : Boolean;

implementation

{$R *.dfm}

procedure AppelAjout (var NewEnreg : Tenreg;GestionIni : TGestionINI);
var XX: TfNewID;
		NewEnr : TEnreg;
begin
	XX := TfNewID.Create(Application);
  NewEnr := TEnreg.Create;
  XX.ModeTrait := taCreat;
  XX.GestionIni := GestionIni;
  XX.Enreg := NewEnr;
  TRY
  	XX.ShowModal;
  finally
    if XX.OkValide then
    begin
    	NewEnreg := newEnr
    end else
    begin
    	NewEnr.Free;
    end;
    XX.Free;
  end;
end;

function AppelModif (ID : string; GestionIni : TGestionINI) : Boolean;
var XX: TfNewID;
		OneEnreg : TEnreg;
begin
  Result := true;
  OneEnreg := GestionIni.FindId(ID);
  if OneEnreg = nil then
  begin
  	Application.MessageBox('Ce collaborateur n''existe plus',PAnsiChar(application.MainForm.Caption));
    Exit;
  end;
  //
	XX := TfNewID.Create(Application);
  XX.ModeTrait := taModif;
  XX.GestionIni := GestionIni;
  XX.Enreg := OneEnreg;
  TRY
  	XX.ShowModal
  finally
    if XX.Supprime then Result := false;
    XX.Free;
  end;
end;

procedure TfNewID.BDELETEClick(Sender: TObject);
begin
	if MessageDlg('Confirmez-vous la suppression cet appareil?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
  	GestionIni.SupprimeEnreg (fEnreg);
    fSupprime := True;
    close;
  end;
end;

procedure TfNewID.BFERMEClick(Sender: TObject);
begin
	close;
end;

procedure TfNewID.BVALIDEClick(Sender: TObject);
begin
	if not VerifOK then Exit;
  EcranToEnreg;
  GestionIni.EcritEnreg (fEnreg,fModeTrait);
  fOkValide := True;
  Close;
end;

procedure TfNewID.EcranToEnreg;
begin
  fEnreg.ID := ID.Text;
  fEnreg.LIBELLE := LIBELLE.Text;
  fEnreg.Server  := SERVER.Text;
  fEnreg.Database  := DATABASE.Text;
  fEnreg.User  := USER.Text;
  fEnreg.Ressource  := RESSOURCE.Text;
  fEnreg.Password := PASSWORD.Text;
end;

procedure TfNewID.EnregToEcran;
  function GetEntryServeur (NomServeur : string) : Integer;
  var ii : Integer;
  begin
    ii := -1;
		for ii := 0 to SERVER.Items.Count -1 do
    begin
      if SERVER.Items [ii] = NomServeur then
      begin
        result := Ii;
        break;
      end;
    end;
	end;

  function GetEntryDatabase (NomBDD : string) : Integer;
  var ii : Integer;
  begin
    ii := -1;
		for ii := 0 to DATABASE.Items.Count -1 do
    begin
      if DATABASE.Items [ii] = NomBDD then
      begin
        result := Ii;
        break;
      end;
    end;
	end;

  function GetEntryRessource (NomRess : string) : Integer;
  var ii : Integer;
  begin
    ii := -1;
		for ii := 0 to MRESSOURCE.Items.Count -1 do
    begin
      if MRESSOURCE.values [ii] = NomRess then
      begin
        result := Ii;
        break;
      end;
    end;
	end;

begin
  //
  //
  ChargeListeServer;
  //
	if fEnreg.Server <> '' then
  begin
  	SERVER.ItemIndex  := getEntryServeur(fEnreg.Server);
    ChargeListeDatabases;
  	DATABASE.ItemIndex := GetEntryDatabase(fEnreg.Database);
  end;
  if fEnreg.Database <> '' then
  begin
    ChargeListeUsers;
    MRESSOURCE.ItemIndex := GetEntryRessource(fEnreg.Ressource);
  end;
  ID.Text := fEnreg.ID;
  LIBELLE.Text := fEnreg.Libelle;
  USER.Text := fEnreg.User;
  PASSWORD.Text := fEnreg.Password;
end;

procedure TfNewID.ChargeListeDatabases;
var DatabaseList : TStringList;
		Indice : Integer;
begin
  DatabaseList := TStringList.Create;
  ListDatabasesOnServer (SERVER.Text,DatabaseList);
  for Indice := 0 to DatabaseList.Count -1 do
  begin
    DATABASE.AddItem(DatabaseList.Strings [Indice],nil)
  end;
  DatabaseList.Free;
end;

procedure TfNewID.ChargeListeServer;
var ServerList : TStringList;
		Indice : Integer;
begin
  ServerList := TStringList.Create;
  ListAvailableSQLServers (ServerList);
  for Indice := 0 to ServerList.Count -1 do
  begin
    SERVER.AddItem(ServerList.Strings [Indice],nil);
  end;
  ServerList.Free;
end;

procedure TfNewID.ChargeListeUsers;
var UserList,UserValue : TStringList;
		Indice : Integer;
begin
  UserList := TStringList.Create;
  UserValue := TStringList.Create;
  ListAvailableUsers (SERVER.Text,DATABASE.Text,UserList,Uservalue);
  for Indice := 0 to UserList.Count -1 do
  begin
  	MRESSOURCE.Items.add (UserList.Strings[Indice]);
    MRESSOURCE.Values.Add(Uservalue.Strings[Indice]);
  end;
  UserList.Free;
  UserValue.Free;
end;

procedure TfNewID.FormShow(Sender: TObject);
begin
  EnregToEcran;

	if fModeTrait = taModif then
  begin
    ID.Enabled := false;
    LIBELLE.SetFocus;
    BDELETE.Visible := True;
  end else
  begin
    DATABASE.Enabled := false;
    MRESSOURCE.Enabled := false;
  end;
  SetEvents(True);
end;

procedure TfNewID.IDExit(Sender: TObject);
begin
  if ID.Text = fEnreg.ID then Exit;
	if GestionIni.FindId(ID.Text) <> nil then
  begin
  	Application.MessageBox('Cet appareil est déjà référencé',PAnsiChar(self.Caption));
    ID.SetFocus;
  end;
end;

procedure TfNewID.SERVERChange(Sender: TObject);
begin
  if SERVER.Text <> fEnreg.Server then
  begin
    DATABASE.Clear;
    DATABASE.Text := '';
    DATABASE.Enabled := True;
    MRESSOURCE.Clear;
    MRESSOURCE.Text := '';
    MRESSOURCE.Enabled := false;
    ChargeListeDatabases;
  end;
end;

procedure TfNewID.DATABASEChange(Sender: TObject);
begin
	if DATABASE.Text <> fEnreg.Database then
  begin
    MRESSOURCE.Clear;
    MRESSOURCE.Text := '';
    MRESSOURCE.Enabled := true;
    ChargeListeUsers;
  end;
end;

procedure TfNewID.SetEvents(Etat: Boolean);
begin
  if Etat then
  begin
    SERVER.OnChange := SERVERChange;
    DATABASE.OnChange := DATABASEChange;
    MRESSOURCE.OnChange := RESSOURCEChange;
  end else
  begin
    SERVER.OnChange := nil;
    DATABASE.OnChange := nil;
    MRESSOURCE.OnChange := nil;
  end;
end;

procedure TfNewID.RESSOURCEChange(Sender: TObject);
begin
	RESSOURCE.Text := MRESSOURCE.Value;
  USER.Text := GetUserAssocie;
  if User.text = '' then
  begin
  	Application.MessageBox('Ce collaborateur n''est plus associé à un utilisateur identifiable',PAnsiChar(self.Caption) );
    MRESSOURCE.ClearSelection;
    MRESSOURCE.SetFocus;
    RESSOURCE.Text := '';
  end;
end;

function TfNewID.GetUserAssocie: string;
begin
	Result := getUserFromSql (SERVER.Text,DATABASE.Text,RESSOURCE.Text);
end;

function TfNewID.VerifOK: Boolean;
begin
	Result := false;
  if ID.Text = '' then
  begin
  	Application.MessageBox('L''identifiant de l''appareil est obligatoire',PAnsiChar(self.Caption) );
    ID.SetFocus;
    Exit;
  end;
  if LIBELLE.Text = '' then
  begin
  	Application.MessageBox('Merci de renseigner une désignation',PAnsiChar(self.Caption) );
    LIBELLE.SetFocus;
    Exit;
  end;
  if SERVER.Text = '' then
  begin
  	Application.MessageBox('Merci de renseigner le serveur de BDD à contacter',PAnsiChar(self.Caption) );
    LIBELLE.SetFocus;
    Exit;
  end;
  if DATABASE.Text = '' then
  begin
  	Application.MessageBox('Merci de renseigner la BDD à contacter',PAnsiChar(self.Caption) );
    LIBELLE.SetFocus;
    Exit;
  end;
  if USER.Text = '' then
  begin
  	Application.MessageBox('Merci un utilisateur identifié',PAnsiChar(self.Caption) );
    LIBELLE.SetFocus;
    Exit;
  end;
  if PASSWORD.Text = '' then
  begin
  	Application.MessageBox('Merci de renseigner le mot de passe',PAnsiChar(self.Caption) );
  	PASSWORD.SetFocus;
    Exit;
  end;
  Result := True;
end;

end.
