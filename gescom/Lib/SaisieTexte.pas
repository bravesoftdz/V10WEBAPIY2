unit SaisieTexte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TFSaisieTexte = class(TForm)
    Edit1: TEdit;
    Button1: TButton;
    Button2: TButton;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FSaisieTexte: TFSaisieTexte;

implementation

{$R *.DFM}

procedure TFSaisieTexte.Button2Click(Sender: TObject);
begin
    Edit1.Text := '';
    Close;
end;

procedure TFSaisieTexte.Button1Click(Sender: TObject);
begin
    Close;
end;

end.
