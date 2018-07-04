unit UADeseria;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls, Mask, ExtCtrls,UUtilSeria,UtilEnvoieMail;

type
  TTFDeseria = class(TForm)
    P: TPageControl;
    lEtape: TLabel;
    BFIn: TBitBtn;
    BSuivant: TBitBtn;
    BPrecedent: TBitBtn;
    BAnnuler: TBitBtn;
    TabSheet1: TTabSheet;
    MDefinition: TMemo;
    TabSheet2: TTabSheet;
    CodeClient: TMaskEdit;
    Label2: TLabel;
    Label3: TLabel;
    Responsable: TMaskEdit;
    Label4: TLabel;
    EmailResp: TMaskEdit;
    Label5: TLabel;
    Societe: TMaskEdit;
    Adress: TLabel;
    Adr1: TMaskEdit;
    Adr2: TMaskEdit;
    Adr3: TMaskEdit;
    CodePostal: TMaskEdit;
    Ville: TMaskEdit;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    TabSheet3: TTabSheet;
    Label13: TLabel;
    SMTPAdress: TMaskEdit;
    Label14: TLabel;
    Label15: TLabel;
    SMTPPort: TMaskEdit;
    CBAuthentification: TCheckBox;
    Luser: TLabel;
    SMTPUser: TMaskEdit;
    LPAssword: TLabel;
    SMTPPassword: TMaskEdit;
    Label18: TLabel;
    TabSheet4: TTabSheet;
    Mresult: TMemo;
    Label16: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Plan: TPanel;
    Ccontrols: TListBox;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PChange(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BSuivantClick(Sender: TObject);
    procedure BPrecedentClick(Sender: TObject);
    procedure BFInClick(Sender: TObject);
    procedure CBAuthentificationClick(Sender: TObject);
  private
    { Déclarations privées }
    function FirstPage: TTabSheet; virtual;
    function PreviousPage: TTabSheet; virtual;
    function NextPage: TTabSheet; virtual;
    function PageNumber: Integer; virtual;
    function PageCount: Integer; virtual;
    function GetPageName: string;
    function GetPage: TTabSheet;
    procedure RestorePage;
    function FindPage(s: string): TTabSheet;
    procedure DisplayStep;
    procedure SetTabs;
		procedure ConstitueDocument;
		function EnvoieDocument: boolean;
		procedure SupprimeAcces;
    function CheckValidEmail(const Value: string): Boolean;
  public
    { Déclarations publiques }
  end;

var
  TFDeseria: TTFDeseria;

implementation

{$R *.dfm}

procedure TTFDeseria.FormShow(Sender: TObject);
begin

  if not SeriaServerpresent then
  begin
    MessageBox(Self.Handle,'Aucun serveur de sérialisation trouvé','ATTENTION',MB_OK);
    SendMessage(Self.Handle, WM_CLOSE, 0, 0);
  end;

  if not isConnexionOk then
  begin
    MessageBox(Self.Handle,'Aucune connexion à Internet n''est possible depuis ce poste','OOOPS',MB_OK);
    SendMessage(Self.Handle, WM_CLOSE, 0, 0);
  end;

  P.ActivePage := FirstPage;
  PChange(nil);

end;

procedure TTFDeseria.FormCreate(Sender: TObject);
begin
  Plan.Top := P.Top;
  Plan.left := P.left;
  Plan.ClientWidth := P.ClientWidth;
  Plan.ClientHeight := P.ClientHeight;
end;

function TTFDeseria.FirstPage: TTabSheet;
begin
  result := P.Pages[0];
end;

function TTFDeseria.GetPage: TTabSheet;
begin
  result := P.ActivePage;
end;

function TTFDeseria.GetPageName: string;
begin
  result := P.ActivePage.Caption;
end;

function TTFDeseria.NextPage: TTabSheet;
begin
  if P.ActivePage.PageIndex < P.PageCount - 1 then result := (P.Pages[P.ActivePage.PageIndex + 1]) else result := nil;
end;

function TTFDeseria.PageCount: Integer;
begin
  Result := P.PageCount;
end;

function TTFDeseria.PageNumber: Integer;
begin
  result := P.ActivePage.PageIndex + 1;
end;

function TTFDeseria.PreviousPage: TTabSheet;
begin
  if P.ActivePage.PageIndex > 0 then result := P.Pages[P.ActivePage.PageIndex - 1] else result := nil;
end;

procedure TTFDeseria.DisplayStep;
begin
  lEtape.Caption := 'Etape' + ' ' + IntToStr(PageNumber) + '/' + IntToStr(PageCount);
end;

function TTFDeseria.FindPage(s: string): TTabSheet;
var
  i: Integer;
begin
  result := nil;
  for i := 0 to P.ControlCount - 1 do
  begin
    if Uppercase(TTabSheet(P.Controls[i]).Caption) = Uppercase(s) then
    begin
      result := TTabSheet(P.Controls[i]);
      exit;
    end;
  end;
end;

procedure TTFDeseria.RestorePage;
var
  i: Integer;
begin
  // Sauvegarde des ItemIndex
  for i := 0 to ComponentCount - 1 do
  begin
    if Components[i] is TComboBox then TComboBox(Components[i]).Tag := TComboBox(Components[i]).ItemIndex;
  end;
  // Changement de parent
  while Plan.ControlCount > 0 do Plan.Controls[0].Parent := P.ActivePage;
  // Restauration des ItemIndex
  for i := 0 to ComponentCount - 1 do
  begin
    if Components[i] is TComboBox then TComboBox(Components[i]).ItemIndex := TComboBox(Components[i]).Tag;
  end;
end;

procedure TTFDeseria.SetTabs;
var
  i: Integer;
  C: TWinControl;
begin
  cControls.Items.Clear;
  for i := 0 to Plan.ControlCount - 1 do
  begin
    if (Plan.Controls[i] is TWinControl) then
    begin
      if (Plan.Controls[i].Left < 0) or (Plan.Controls[i].Top < 0) then continue;
      cControls.Items.Add(FormatFloat('0000', Plan.Controls[i].Top) + FormatFloat('0000', Plan.Controls[i].Left) + copy(Plan.Controls[i].Name,1, 20));
    end;
  end;
  for i := 0 to cControls.Items.Count - 1 do
  begin
    C := TWinControl(FindComponent(Trim(Copy(cControls.Items[i], 9, 20))));
    if C <> nil then C.TabOrder := i;
  end;
end;

procedure TTFDeseria.PChange(Sender: TObject);
var
  i: Integer;
  T: TTabSheet;
  fs, fp: Boolean;
begin
  T := P.ActivePage;
  if T = nil then exit;
  while T.ControlCount > 0 do
  begin
    if T.Controls[0] is TComboBox then
    begin
      TComboBox(T.Controls[0]).Tag := TComboBox(T.Controls[0]).ItemIndex; // Sauvegarde de ItemIndex dans Tag
    end;
    T.Controls[0].Parent := Plan;
  end;
  for i := 0 to Plan.ControlCount - 1 do // Restauration des ItemIndex
  begin
    if Plan.Controls[i] is TComboBox then TComboBox(Plan.Controls[i]).ItemIndex := TComboBox(Plan.Controls[i]).Tag;
  end;
  SetTabs;
  if PreviousPage <> nil then fp := (T.PageIndex > 0) else fp := FALSE;
  if NextPage <> nil then fs := (T.PageIndex < P.ControlCount - 1) else fs := FALSE;
  bPrecedent.Enabled := fp;
  bSuivant.Enabled := fs;
  bSuivant.Default := bSuivant.Enabled;
  bFin.Default := not bSuivant.Default;
  BFIn.Enabled := (P.ActivePage.PageIndex = P.PageCount - 1); 
  DisplayStep;
end;

procedure TTFDeseria.BAnnulerClick(Sender: TObject);
begin
	Close;
end;

procedure TTFDeseria.BSuivantClick(Sender: TObject);
var
  t: TTabSheet;
begin
  if P.ActivePage = TabSheet3 then
  begin
    if (CodeClient.Text = '') or  (Societe.Text = '') or  (Adr1.Text='') or  (CodePostal.Text='') or (Ville.Text = '') then
    begin
      RestorePage;
      P.ActivePage := TabSheet2;
      PChange(nil);
      MessageBox(Self.Handle,'Vous devez renseigner les zones obligatoires','Attention',MB_OK);
      Exit;
    end else if (SMTPAdress.Text = '') or  (StrToInt(SMTPPort.Text)  = 0)  or
    						((CBAuthentification.Checked) and (((SMTPUser.Text ='')) or (SMTPPassword.Text ='')))  then
    begin
      RestorePage;
      P.ActivePage := TabSheet3;
      PChange(nil);
      MessageBox(Self.Handle,'Vous devez renseigner les zones obligatoires','Attention',MB_OK);
      Exit;
    end;
    if not CheckValidEmail ( EmailResp.Text) then
    begin
      MessageBox(Self.Handle,'Merci de renseigner une adresse Mail Valide','Attention',MB_OK);
      Exit;
    end;
    ConstitueDocument;
  end;
  RestorePage;
  t := NextPage;
  if t = nil then P.SelectNextPage(True) else
  begin
    P.ActivePage := t;
    PChange(nil);
  end;
end;

procedure TTFDeseria.BPrecedentClick(Sender: TObject);
var
  t: TTabSheet;
begin
  RestorePage;
  t := PreviousPage;
  if t = nil then P.SelectNextPage(False) else
  begin
    P.ActivePage := t;
    PChange(nil);
  end;
end;

procedure TTFDeseria.BFInClick(Sender: TObject);
begin
    if EnvoieDocument then
    begin
			SupprimeAcces;
      close;
    end;
end;

procedure TTFDeseria.ConstitueDocument;
var OneLig : string;
begin
	Mresult.Clear;
  //
	OneLig := '<h1>Arrêt d''utilisation des produits CEGID/L.S.E Business</h1>';
  Mresult.Lines.Add(OneLig);
  Mresult.Lines.Add('<BR>');
  Mresult.Lines.Add('<BR>');
  OneLig := 'Je sous signé Mr/Mme '+Responsable.text+'<BR>';
  Mresult.Lines.Add(OneLig);
  OneLig := 'Email '+EmailResp.text+'<BR>';
  Mresult.Lines.Add(OneLig);
  Mresult.Lines.Add('<BR>');
  Mresult.Lines.Add('<BR>');
  OneLig := 'Responsable de la société '+Societe.Text+' Code Client : '+CodeClient.text+'<BR>';
  Mresult.Lines.Add(OneLig);
  OneLig := Adr1.Text+'<BR>';
  Mresult.Lines.Add(OneLig);
  if Adr2.Text <> '' then
  begin
    OneLig := Adr2.Text+'<BR>';
    Mresult.Lines.Add(OneLig);
  end;
  if Adr2.Text <> '' then
  begin
    OneLig := Adr3.Text+'<BR>';
    Mresult.Lines.Add(OneLig);
  end;
  OneLig := CodePostal.text+' '+Ville.text+'<BR>';
  Mresult.Lines.Add(OneLig);
  Mresult.Lines.Add('<BR>');
  OneLig := 'Atteste avoir accepté la désinstallation des produits CEGID/L.S.E Business et qu''ils ne seront plus utilisés depuis le matériel suivant :<BR>';
  Mresult.Lines.Add(OneLig);
  OneLig := 'A la date du '+DateToStr(Now)+'<BR>';
  Mresult.Lines.Add(OneLig);
  OneLig := 'Nom du serveur de sérialisation : '+ComputerName+'<BR>';
  Mresult.Lines.Add(OneLig);
  OneLig := 'Adresse du serveur de sérialisation : '+GetAdrSeria+'<BR>';
  Mresult.Lines.Add(OneLig);
end;

function TTFDeseria.EnvoieDocument: boolean;
var Destinataire : string;
begin
	Result := false;
  Destinataire := 'l.santucci@lse.fr';
  Result := SMTPSendMail (EmailResp.Text,Destinataire,'',Mresult,SMTPAdress.Text,SMTPPort.Text,SMTPUser.Text,SMTPPassword.Text,CBAuthentification.Checked);
  if Not result then MessageBox(Application.Handle,'Document non envoyé','Attention',MB_OK);
end;

function TTFDeseria.CheckValidEmail(const Value: string): Boolean;
    function CheckAllowed(const s: string): Boolean;
    var i: Integer;
    begin
      Result:= false;
      for i:= 1 to Length(s) do
      if not (s[i] in ['a'..'z',
                       'A'..'Z',
                       '0'..'9',
                       '_',
                       '-',
                       '.']) then Exit;
      Result:= true;
    end;
var
  i: Integer;
  NamePart, ServerPart: string;
begin
  Result:= False;
  i:=Pos('@', Value);
  if i=0 then Exit;
  NamePart:=Copy(Value, 1, i-1);
  ServerPart:=Copy(Value, i+1, Length(Value));
  if (Length(NamePart)=0) or ((Length(ServerPart)<5)) then Exit;
  i:=Pos('.', ServerPart);
  if (i=0) or (i>(Length(serverPart)-2)) then Exit;
  Result:= CheckAllowed(NamePart) and CheckAllowed(ServerPart);
end;

procedure TTFDeseria.SupprimeAcces;
begin
	DeleteSeria;
end;

procedure TTFDeseria.CBAuthentificationClick(Sender: TObject);
begin
	SMTPUser.Visible := CBAuthentification.Checked;
	SMTPPassword.Visible := CBAuthentification.Checked;
  Luser.Visible := CBAuthentification.Checked;
  LPAssword.Visible := CBAuthentification.Checked;
end;

end.
