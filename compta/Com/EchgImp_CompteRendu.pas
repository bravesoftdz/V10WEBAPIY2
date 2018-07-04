unit EchgImp_CompteRendu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, ColMemo, HTB97, Hctrls, uiutil,
  Echg_Code;

type
  TfEchgImp_CompteRendu = class(TForm)
    ColorMemo1: TColorMemo;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BFerme: TToolbarButton97;
    ToolbarButton972: TToolbarButton97;
    bSauve: TToolbarButton97;
    SaveDialog1: TSaveDialog;
    HOpenDialog1: THOpenDialog;
    procedure BFermeClick(Sender: TObject);
    procedure bSauveClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  fEchgImp_CompteRendu: TfEchgImp_CompteRendu;

implementation

{$R *.DFM}

procedure TfEchgImp_CompteRendu.BFermeClick(Sender: TObject);
begin
  Close;
  if IsInside (Self) then
    CloseInsidePanel (Self) ;
end;

procedure TfEchgImp_CompteRendu.bSauveClick(Sender: TObject);
begin
  CompteRenduComSauvegarde (SaveDialog1, ColorMemo1);
end;

end.
