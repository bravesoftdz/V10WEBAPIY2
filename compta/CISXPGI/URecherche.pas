unit URecherche;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, HSysMenu, HTB97;

type
  TFRecherche = class(TFVierge)
    RechText: TEdit;
    Label1: TLabel;
    RespecterMaj: TCheckBox;
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FRecherche: TFRecherche;

implementation

{$R *.DFM}

procedure TFRecherche.BValiderClick(Sender: TObject);
begin
  inherited;
  ModalResult := mrOK;
end;

end.
