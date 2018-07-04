unit UOption;

interface

uses
  Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, ComCtrls, HSysMenu, HTB97, StdCtrls, Mask, Hctrls;

type
  TOptionsDlg = class(TFVierge)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    GroupBox2: TGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    cbSepMil: TComboBox;
    cbSepDec: TComboBox;
    LTable: TLabel;
    Table: THCritMaskEdit;
    procedure cbSepMilChange(Sender: TObject);
    procedure cbSepDecChange(Sender: TObject);
    procedure TableElipsisClick(Sender: TObject);
  private
    { Déclarations privées }
		Fmillier : char;
		fDecimal : char;
		function GetScriptSuivant : String;
		procedure SetScriptSuivant(AValue:String);

  public
    { Déclarations publiques }
		property Millier : Char read FMillier write FMillier;
		property Decimal : Char read FDecimal write FMillier;
		property ScriptSuivant : string read GetScriptSuivant write SetScriptSuivant;
  end;

var
  OptionsDlg: TOptionsDlg;

implementation
{$IFNDEF EAGLCLIENT}
uses USelectScript;
{$ENDIF}
{$R *.DFM}

procedure TOptionsDlg.cbSepMilChange(Sender: TObject);
begin
  inherited;
	case cbSepMil.itemindex of
		0 : Fmillier := '.';
		1 : Fmillier := ',';
		2 : Fmillier := ' ';
		else Fmillier := ' ';
	end;

end;

procedure TOptionsDlg.cbSepDecChange(Sender: TObject);
begin
  inherited;
	 case cbsepdec.ItemIndex of
		  0 : Fdecimal := '.';
		  1 : Fdecimal := ',';
		  2 : Fdecimal := ' ';
		else FDecimal := ' ';
	 end;
end;

procedure TOptionsDlg.TableElipsisClick(Sender: TObject);
Var
DD : string;
begin
  inherited;
{$IFNDEF EAGLCLIENT}
Table.Text := SelectScript(DD);
{$ENDIF}
end;

function TOptionsDlg.GetScriptSuivant : String;
begin
	Result := Table.Text
end;

procedure TOptionsDlg.SetScriptSuivant(AValue:String);
begin
	Table.Text := AValue
end;


end.
