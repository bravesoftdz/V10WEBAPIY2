unit UAcceuil;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, HTB97, TntExtCtrls, HPanel, GraphicEx, Hctrls,HmsgBox,UCalcProtect,
  StdCtrls;

type
  TFprincipale = class(TForm)
    Panel1: TPanel;
    BABANDON: TToolbarButton97;
    PCALCULPROT: THPanel;
    Label1: TLabel;
    Panel2: TPanel;
    IMGLSE: THImage;
    procedure BABANDONClick(Sender: TObject);
    procedure PCALCULPROTClick(Sender: TObject);
    procedure PMAJPROTClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  Fprincipale: TFprincipale;

implementation

{$R *.dfm}

procedure TFprincipale.BABANDONClick(Sender: TObject);
begin
  close;
end;

procedure TFprincipale.PCALCULPROTClick(Sender: TObject);
begin
  OuvreFicheCalc;
end;

procedure TFprincipale.PMAJPROTClick(Sender: TObject);
begin
//
PgiInfo('MAJ protection');
end;

end.
