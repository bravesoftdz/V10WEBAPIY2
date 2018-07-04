unit Main;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, ImgList;

type
  TMainForm = class(TForm)
    ed_inifile: TEdit;
    btn_browse: TButton;
    Label1: TLabel;
    cbx_bases: TComboBox;
    Label2: TLabel;
    btn_GO: TButton;
    progress: TProgressBar;
    Label3: TLabel;
    OpenDialog1: TOpenDialog;
    iml_progress: TImageList;
    list: TListView;
    chk_zipped: TCheckBox;
    procedure ed_inifileChange(Sender: TObject);
    procedure cbx_basesChange(Sender: TObject);
    procedure btn_GOClick(Sender: TObject);
    procedure btn_browseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
  public
    procedure SetStatus(caption : string; ImageIdx : integer; Msg : String);
    procedure InitListView(FirstTime : boolean);
  end;

var
  MainForm: TMainForm;

implementation
uses IniFiles, EBizUtil;

{$R *.DFM}

procedure TMainForm.ed_inifileChange(Sender: TObject);
var bf : TIniFile;
    SL : TStringList;
begin
if not FileExists(ed_inifile.text) then
   begin cbx_bases.Clear;
         btn_go.Enabled := false;
         exit; end;

SL := TStringList.Create;
bf := TIniFile.Create(ed_inifile.text);
bf.ReadSections(SL);
cbx_bases.Items.Assign(SL);
bf.Free;
SL.Free;
end;

procedure TMainForm.SetStatus(caption : string; ImageIdx : integer; Msg : String);
begin
with list.FindCaption(0, caption, false, true, true) do
  begin
  ImageIndex := ImageIdx;
  SubItems[0] := Msg;
  MakeVisible(false);
  end;
Application.ProcessMessages;
end;

const PGI2E : Array[1..25] of String =
              ('ADRESSES','EARTICLE', 'ARTICLELIE', 'CHANCELL', 'CHOIXCOD', 'COMMERCIAL',
               'LIENSOLE', 'CONTACT', 'DECOMBOS', 'DEVISE', 'DIMENSION', 'DIMMASQUE', 'ETIERS',
               'JUTYPECIVIL',  'MEA', 'MODEPAIE', 'MODEREGL', 'PARAMSOC', 'PAYS', 'TARIF',
               'TRADTABLETTE', 'GTRADARTICLE', 'TRADDICO', 'TXCPTTVA', 'COMMUN');
      ExpFolder : String = 'Generation\';

procedure TMainForm.InitListView(FirstTime : boolean);
var i : integer;
begin
list.Items.BeginUpdate;
if FirstTime then list.Items.Clear;
for i := Low(PGI2E) to High(PGI2E) do
 if FirstTime then
  with list.Items.Add do
    begin
     Caption := PGI2E[i];
     SubItems.Add('En attente');
     ImageIndex := 0;
     Checked := true;
    end
 else
  with list.FindCaption(0, PGI2E[i], false, true, true) do
    begin
    ImageIndex := 0;
    SubItems[0] := 'En attente';
    end;
list.Items.EndUpdate;
end;

procedure TMainForm.cbx_basesChange(Sender: TObject);
begin
btn_GO.Enabled := (cbx_bases.ItemIndex > -1);
end;

procedure TMainForm.btn_GOClick(Sender: TObject);
var i : integer;
    F, basedir : string;
begin
btn_GO.Enabled := false;
ed_inifile.Enabled := false;
ed_inifile.Color := clBtnFace;
btn_browse.Enabled := false;
cbx_bases.Enabled := false;
cbx_bases.Color := clBtnFace;
chk_zipped.Enabled := false;

if paramcount <> 1 then basedir := extractfilepath(Paramstr(0))
                   else basedir := paramstr(1);

InitListView(false);

Label3.Caption := 'Connexion en cours...'; Application.ProcessMessages;
TAConnecte(ed_inifile.text, cbx_bases.Text);

progress.Max := Length(PGI2E);
for i := Low(PGI2E) to High(PGI2E) do
  begin
  if not list.FindCaption(0, PGI2E[i], false, true, true).Checked then Continue;
  F := basedir+ExpFolder+'Export_'+PGI2E[i]+'.xml';
  if chk_zipped.Checked then F := F+'.hz';
  Label3.Caption := 'Intégration de '+PGI2E[i]; Application.ProcessMessages;
  SetStatus(PGI2E[i], 1, 'Intégration...');
  if not FileExists(F)
    then SetStatus(PGI2E[i], 3, 'Fichier non trouvé ('+F+')')
    else try
           TAIntegreDataFichier(F, chk_zipped.Checked);
           SetStatus(PGI2E[i], 2, 'OK');
         except
           on E : Exception do SetStatus(PGI2E[i], 3, E.Message);
         end;
  progress.position := i+1;
  end;

Label3.Caption := 'Déconnexion...'; Application.ProcessMessages;
TADeConnecte;
Label3.Caption := 'Terminé';
progress.Position := 0;

btn_GO.Enabled := true;
ed_inifile.Enabled := true;
ed_inifile.Color := clWindow;
btn_browse.Enabled := true;
cbx_bases.Enabled := true;
cbx_bases.Color := clWindow;
chk_zipped.Enabled := true;
end;

procedure TMainForm.btn_browseClick(Sender: TObject);
begin
//opendialog1.filename := ed_inifile.text;
if OpenDialog1.Execute then ed_inifile.text := OpenDialog1.FileName;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var s : string;
begin
setlength(s, 255);
GetWindowsDirectory(PChar(s), 255);
opendialog1.initialdir := s;
InitListView(true);
end;

end.
