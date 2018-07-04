unit SaisieTexteValeur;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, Hctrls;

type
  TFSaisieTexteValeur = class(TForm)
    Saisie: TEdit;
    bOk: TToolbarButton97;
    bAnnuler: TToolbarButton97;
    Libelle: THLabel;
    GroupBox1: TGroupBox;
    procedure bOkClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
  end;

var
  FSaisieTexteValeur: TFSaisieTexteValeur;

implementation

{$R *.DFM}

procedure TFSaisieTexteValeur.bOkClick(Sender: TObject);
begin
  Close;
end;

procedure TFSaisieTexteValeur.bAnnulerClick(Sender: TObject);
begin
  Saisie.Text := '';
  Close;
end;

end.
