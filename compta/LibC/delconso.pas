unit delconso;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs
   ,Hctrls,HPanel, Uiutil, hmsgbox, Hent1, HTB97, StdCtrls, Mask, ExtCtrls
   ,Utob, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, ent1, HFLabel, TImpFic, HStatus;

procedure DelAffaire ;

type
  TDelConso = class(TForm)
    Panel1: TPanel;
    Dock971: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    MsgBox: THMsgBox;
    Aff: THCritMaskEdit;
    Label4: TLabel;
    CEcr: TCheckBox;
    Patience: TLabel;
    CAna: TCheckBox;
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}


procedure DelAffaire ;
var ExpBa: TDelConso ;
    PP : THPanel ;
BEGIN
ExpBa:=TDelConso.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
  begin
  try
    ExpBa.ShowModal ;
  finally
    ExpBa.Free ;
  end ;
  Screen.Cursor:=SyncrDefault ;
  end
  else begin
  InitInside(ExpBa,PP) ;
  ExpBa.Show ;
  end ;
end ;

procedure TDelConso.BValiderClick(Sender: TObject);
Var
    St : String ;
begin
If MsgBox.Execute(0,Caption,'')<>mrYes Then Exit ;
Patience.Visible:=TRUE ; Application.ProcessMessages ;
St:='Delete from ECRITURE WHERE E_LIBRETEXTE9="'+AFF.Text+'" ' ;
ExecuteSQL(St) ;
St:='Delete from ANALYTIQ WHERE Y_LIBRETEXTE9="'+AFF.Text+'" ' ;
ExecuteSQL(St) ;
Patience.Visible:=FALSE ; Application.ProcessMessages ;
end;


end.
