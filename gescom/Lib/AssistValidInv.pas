unit AssistValidInv;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, ComCtrls, ImgList, HSysMenu, hmsgbox, StdCtrls, HTB97, ExtCtrls,
  Hctrls, UTOB, HPanel, TntStdCtrls, TntComCtrls, TntExtCtrls;

type
  TValidateProc = procedure(CodeListe : String) of object;
  TFAssistValidInv = class(TFAssist)
    TabSheet1: TTabSheet;
    WelcomePanel: TPanel;
    HLabel3: THLabel;
    TabSheet2: TTabSheet;
    GroupBox2: TGroupBox;
    cbx_prix1: THValComboBox;
    cbx_prix2: THValComboBox;
    cbx_prix3: THValComboBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    lbx_prix: TListBox;
    Label1: TLabel;
    Bevel1: TBevel;
    CB_REMISEAZERO: TCheckBox;
    Label2: TLabel;
    CHKUPDATEPMAP: TCheckBox;
    procedure OnComboChange(Sender : TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
  private
    FTOBSelection : TOB;

  public
    Validated : boolean;
    procedure Init(TOBSel : TOB);

  end;

function EntreeAssistValidInv(TOBSelection : TOB) : boolean;

implementation

uses ValidateInv;

{$R *.DFM}


function EntreeAssistValidInv(TOBSelection : TOB) : boolean;
begin
with TFAssistValidInv.Create(Application) do
  try
    Init(TOBSelection);
    ShowModal;
  finally
    result := Validated;
    Release;
  end;
end;

procedure TFAssistValidInv.Init(TOBSel : TOB);
begin
FTOBSelection := TOBSel;
Validated := false;
end;


procedure TFAssistValidInv.OnComboChange(Sender : TObject);
var FCombo, FSameCombo : THValComboBox;
    i_cbx, i_idx : integer;
    b : boolean;
begin
FCombo := (Sender as THValComboBox);

for i_cbx := 1 to 3 do
  begin
  FSameCombo := THValComboBox(FindComponent('cbx_prix'+inttostr(i_cbx)));
  if (FSameCombo <> FCombo) and (FSameCombo.ItemIndex = FCombo.ItemIndex) then break
                                                                          else FSameCombo := nil;
  end;
if FSameCombo = nil then exit;

for i_idx := 0 to 2 do
  begin
  b := true;
  for i_cbx := 1 to 3 do
   if THValComboBox(FindComponent('cbx_prix'+inttostr(i_cbx))).ItemIndex = i_idx then b := false;
  if b then
    begin
    FSameCombo.ItemIndex := i_idx;
    break;
    end;
  end;
end;

procedure TFAssistValidInv.bFinClick(Sender: TObject);
begin
inherited;

ValideLesListes(Self, FTOBSelection,
                THValComboBox(FindComponent('cbx_prix1')).Value,
                THValComboBox(FindComponent('cbx_prix2')).Value,
                THValComboBox(FindComponent('cbx_prix3')).Value,
                CB_REMISEAZERO.checked,
                CHKUPDATEPMAP.checked);

Validated := true;
Close;
end;

procedure TFAssistValidInv.FormShow(Sender: TObject);
var i : integer;
begin
inherited;
for i := 0 to 2 do
 THValComboBox(FindComponent('cbx_prix'+inttostr(i+1))).ItemIndex := i;
bFin.Enabled := false;
end;

procedure TFAssistValidInv.bSuivantClick(Sender: TObject);
begin
inherited;
bFin.Enabled := true;
lbx_prix.Clear;
lbx_prix.Items.Add(Msg.Mess[2]);
lbx_prix.Items.Add('  1. '+cbx_prix1.Text);
lbx_prix.Items.Add('  2. '+cbx_prix2.Text);
lbx_prix.Items.Add('  3. '+cbx_prix3.Text);
lbx_prix.Items.Add('');
if CB_REMISEAZERO.Checked=True
   then lbx_prix.Items.Add('Remise à zéro du stock des articles non inventoriés')
   else lbx_prix.Items.Add('Pas de remise à zéro du stock des articles non inventoriés');
end;

procedure TFAssistValidInv.bPrecedentClick(Sender: TObject);
begin
inherited;
bFin.Enabled := false;
end;

end.
