unit ImVisuEcr;
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, hmsgbox, HTB97, Menus, Grids, StdCtrls, Buttons,
  UiUtil,UTob, Hctrls, ExtCtrls, HPanel, Outils, Hent1 ;

type
  TImVisuEcr = class(TForm)
    HM: THMsgBox;
    HPanel1: THPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bExport: TToolbarButton97;
    BImprimer: TToolbarButton97;
    HPanel4: THPanel;
    HPanel3: THPanel;
    GridEcriture: THGrid;
    HMTrad: THSystemMenu;
    CreditProg: THNumEdit;
    DebitProg: THNumEdit;
    CreditGen: THNumEdit;
    DebitGen: THNumEdit;
    HLabel1: THLabel;
    HLabel3: THLabel;
    BVentil: TToolbarButton97;
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    LaTob: Tob ;
    procedure PutLaTobSurLaGrille(UneTob: Tob) ;
  public
    { Déclarations publiques }
  end;

procedure ImVisuEcriture(E_ : Tob) ;

implementation

{$R *.DFM}
procedure  ImVisuEcriture(E_ : Tob) ;
var
  ImVisuEcr: TImVisuEcr;
  PP : THpanel;
begin
  ImVisuEcr:=TImVisuEcr.Create(Application);
  ImVisuEcr.LaTob:=E_ ;
  PP:=FindInsidePanel ;
  if PP=Nil then
  begin
    try
      ImVisuEcr.ShowModal;
    finally
      ImVisuEcr.Free;
    end;
  end else
  begin
    InitInside(ImVisuEcr,PP) ;
    ImVisuEcr.Show ;
  end;
end;

procedure TImVisuEcr.FormShow(Sender: TObject);
begin
  GridEcriture.ColAligns[0]:= taCenter;
  GridEcriture.ColAligns[1]:= taCenter;
  GridEcriture.ColAligns[2]:= taCenter;
  GridEcriture.ColAligns[3]:= taLeftJustify;
  GridEcriture.ColAligns[4]:= taRightJustify;
  GridEcriture.ColAligns[5]:= taRightJustify;
  if LaTob<>nil then PutLaTobSurLaGrille(LaTob) ;
  GridEcriture.RowCount:=GridEcriture.RowCount-1 ;
end;

procedure TImVisuEcr.PutLaTobSurLaGrille(UneTob: Tob) ;
const Tab: array[0..5] of string=('E_DATECOMPTABLE','E_JOURNAL','E_GENERAL','E_LIBELLE','E_DEBITDEV','E_CREDITDEV') ;
var i,j: integer ; T1: Tob ; Okok: boolean ;
begin
  for i:=0 to UneTob.Detail.Count-1 do
  begin
    T1:=UneTob.Detail[i] ;
    Okok:=true ;
    for j:=low(Tab) to High(Tab) do if not T1.FieldExists(Tab[j]) then begin Okok:=false; break ; end ;
    if Okok then
    begin
      for j:=low(Tab) to High(Tab) do
      begin
        if (Tab[j]='E_DEBITDEV') or (Tab[j]='E_DEBITDEV') then
          GridEcriture.Cells[j,GridEcriture.RowCount-1]:=FormatFloat('#0.00',arrondi(T1.GetValue(Tab[j]),2))
        else
          GridEcriture.Cells[j,GridEcriture.RowCount-1]:=T1.GetValue(Tab[j])
      end ;
      GridEcriture.RowCount:=GridEcriture.RowCount+1 ;
    end
    else
    begin
      PutLaTobSurLaGrille(T1) ;
    end ;
  end ;
end ;



end.

