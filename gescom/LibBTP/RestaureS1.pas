unit RestaureS1;

interface

uses Windows,
     Messages,
     SysUtils,
     Classes,
     Graphics,
     Controls,
     Forms,
     Dialogs,
		 HTB97,
     ComCtrls,
     StdCtrls;
//   Controls;

type
  TRestaure_S1 = class(TForm)
    Label1: TLabel;
    Trestaure: TLabel;
    ComboBox1: TComboBox;
    ODGetInfosBAK: TOpenDialog;
    Animate1: TAnimate;
    Label2: TLabel;
    BValide: TToolbarButton97;
    BAnnule: TToolbarButton97;
    procedure BAnnuleClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  Restaure_S1: TRestaure_S1;

implementation

{$R *.DFM}

procedure TRestaure_S1.BAnnuleClick(Sender: TObject);
begin

  ComboBox1.Clear;

	Close;

end;

procedure TRestaure_S1.BValideClick(Sender: TObject);
begin

  Combobox1.Enabled := False;

	Close;

end;

end.
