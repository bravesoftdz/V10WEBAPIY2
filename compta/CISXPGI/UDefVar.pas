{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 21/10/2002
Modifié le ... :   /  /    
Description .. : Unité des définitions des variables qui seront utilisés au 
Suite ........ : niveau du script
Mots clefs ... : 
*****************************************************************}

unit UDefVar;

interface

uses
  Windows, SysUtils, Classes, Controls, Forms,
  Vierge, StdCtrls, ExtCtrls, UScript, HSysMenu, HTB97;

type
  TFDefvariable = class(TFVierge)
    Bevel1: TBevel;
    ListBox1: TListBox;
    procedure BinsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure ListBox1DblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
    FListVar : TVariableDefList;
	procedure SetListVar(Avalue : TVariableDefList);
	function GetListVar : TVariableDefList;
  public
    { Déclarations publiques }
	constructor Create(AOwner : TComponent); override;
	destructor Destroy; override;
	property ListVar : TVariableDefList read GetListVar write SetListVar;

  end;

var
  FDefvariable: TFDefvariable;

Procedure DefinitionVar;

implementation

uses UVar1;

{$R *.DFM}

Procedure DefinitionVar;
BEGIN
     FDefvariable :=TFDefvariable.Create(Application) ;
     FDefvariable.ShowModal ;
     FDefvariable.Free ;
END ;

procedure TFDefvariable.SetListVar(Avalue : TVariableDefList);
var I : Integer;
begin
	FListVar.Assign(AValue);
	ListBox1.Clear;
	for I:=0 to AValue.Count-1 do
		ListBox1.Items.AddObject(AValue[I].Name+'='+AValue[I].Text, FListVar[I]);
end;

function TFDefvariable.GetListVar : TVariableDefList;
begin
	Result := FListVar;
end;


constructor TFDefvariable.Create(AOwner : TComponent);
begin
	inherited Create(AOwner);
	FListVar := TVariableDefList.Create;
end;

destructor TFDefvariable.Destroy;
begin
	FListVar.Free;
	inherited Destroy;
end;


procedure TFDefvariable.BinsertClick(Sender: TObject);
var AVarDef : TVariableDef;
begin
  inherited;
	VariableDlg := TVariableDlg.Create(self);
	VariableDlg.edtValue.ReadOnly := False;
	VariableDlg.edtValue.Text := '';
	VariableDlg.edtName.Text := '';
	VariableDlg.edtIntitule.Text := '';
    VariableDlg.ComboBox1.ItemIndex := 1;
	if VariableDlg.ShowModal = mrOK then  // Uvar1
	begin
		AVarDef := FListVar.AddObject('Nouveau');
		AVarDef.Name := VariableDlg.edtValue.Text;
		AVarDef.Text := VariableDlg.edtName.Text;
		AVarDef.Libelle := VariableDlg.edtIntitule.Text;
		AVarDef.Demandable := VariableDlg.CheckBox1.Checked;
        AVarDef.TypeVar := VariableDlg.ComboBox1.ItemIndex;
		if Trim(VariableDlg.edtValue.Text) <> '' then
			ListBox1.Items.AddObject(Trim(VariableDlg.edtValue.Text)+'='+VariableDlg.edtName.Text, AVarDef);
	end;
	VariableDlg.Free;

end;

procedure TFDefvariable.BDeleteClick(Sender: TObject);
begin
  inherited;
	if ListBox1.ItemIndex < 0 then exit;
    FlistVar.Delete(FListVar.IndexOf(Listbox1.items.names[Listbox1.ItemIndex]));
	ListBox1.Items.Delete(ListBox1.ItemIndex);
end;

procedure TFDefvariable.ListBox1DblClick(Sender: TObject);
var
	S : String; P, index : Integer;
begin
  inherited;
	index := ListBox1.ItemIndex;
	if index < 0 then exit;
	VariableDlg := TVariableDlg.Create(self);
	S := 	ListBox1.Items[index];
	P := Pos('=', S);
	if P = 0 then begin
		VariableDlg.edtValue.Text := S;
		VariableDlg.edtName.Text := '';
	end else begin
		VariableDlg.edtValue.Text := Copy(S, 1, P-1);
		VariableDlg.edtName.Text := Copy(S, P+1, 255);
	end;
	VariableDlg.edtIntitule.Text := TVariableDef(ListBox1.Items.Objects[index]).Libelle;
	VariableDlg.CheckBox1.Checked := TVariableDef(ListBox1.Items.Objects[index]).Demandable;
	VariableDlg.ComboBox1.ItemIndex := TVariableDef(ListBox1.Items.Objects[index]).TypeVar;
	VariableDlg.edtValue.ReadOnly := True;
	if VariableDlg.ShowModal = mrOK then
	begin
		TVariableDef(ListBox1.Items.Objects[index]).Libelle := VariableDlg.edtIntitule.Text;
		TVariableDef(ListBox1.Items.Objects[index]).Text := VariableDlg.edtName.Text;
		TVariableDef(ListBox1.Items.Objects[index]).TypeVar := VariableDlg.ComboBox1.ItemIndex;
		ListBox1.Items[index] := VariableDlg.edtValue.Text+'='+VariableDlg.edtName.Text;
	end;
	VariableDlg.Free;
end;

procedure TFDefvariable.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
	if (Shift = []) and (Key = VK_F2) then
              ListBox1DblClick (Sender)
	else if (Shift = []) and (Key = VK_DELETE) then
              BDeleteClick (Sender)
	else if (Shift = []) and (Key = VK_INSERT) then
		      BinsertClick(Sender);
end;

procedure TFDefvariable.BValiderClick(Sender: TObject);
begin
  inherited;
  ModalResult := mrOK;
end;

end.
