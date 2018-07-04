unit UVar1;

interface

uses Classes, Forms, Controls, StdCtrls, HTB97;

type

  TVariableDlg = class(TForm)
    Label1: TLabel;
	Label2: TLabel;
	edtValue: TEdit;
	edtName: TEdit;
	CheckBox1: TCheckBox;
	Button1: TButton;
	ComboBox1: TComboBox;
	Label3: TLabel;
    edtIntitule: TEdit;
    Label4: TLabel;
    ListBox1: TListBox;
    Label5: TLabel;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
	procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
	{ Private declarations }
	FDetail : Boolean;

	procedure SetDetail(AValue:Boolean);

	property Detail : Boolean read FDetail write SetDetail;
  public
	{ Public declarations }
  end;

var
  VariableDlg: TVariableDlg;

implementation

{$R *.DFM}

procedure TVariableDlg.SetDetail(AValue:Boolean);
begin
	FDetail := AValue;
	if FDetail then
	begin
		ClientHeight := ComboBox1.Top + ComboBox1.Height + 4 + Button1.Height + 8;
		ComboBox1.Show;
		Label3.Show;
	end
	else begin
		ClientHeight := ComboBox1.Top + Button1.Height + 4;
		ComboBox1.Hide;
		Label3.Hide;
	end;
end;

procedure TVariableDlg.FormShow(Sender: TObject);
begin
	if edtValue.ReadOnly then
		edtName.SetFocus
	else edtValue.SetFocus;

	SetDetail(FDetail);
end;

procedure TVariableDlg.Button1Click(Sender: TObject);
begin
	Detail := not Detail;
	if FDetail then
		Button1.Caption := 'Moins de détails <<<'
	else Button1.Caption := 'Plus de détails >>>';
end;

procedure TVariableDlg.BValiderClick(Sender: TObject);
begin
ModalResult := mrOK;
end;

end.
