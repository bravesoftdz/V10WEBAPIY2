unit ftest;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Hctrls, Mask;

type
  TForm1 = class(TForm)
    HValComboBox1: THValComboBox;
    RESULTVALUE: THCritMaskEdit;
    RESULTEXT: THCritMaskEdit;
    procedure HValComboBox1Change(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure Ouvretest;

implementation

uses TntStdCtrls;

{$R *.dfm}

procedure Ouvretest;
var Form1: TForm1;
begin
  Form1 := TForm1.Create(application);
  TRY
	  Form1.ShowModal;
  FINALLY
  	Form1.Free;
  end;
end;


procedure TForm1.HValComboBox1Change(Sender: TObject);
begin
	RESULTVALUE.Text := HValComboBox1.Value;
	RESULTEXT.Text := HValComboBox1.Text;
end;

end.
