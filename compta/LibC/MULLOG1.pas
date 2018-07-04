unit MULLOG1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB, {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  ExtCtrls, ComCtrls, Buttons, Mask, HRichEdt, Ent1, HSysMenu, HDB, HEnt1,
  HTB97, ColMemo, HMsgBox, HPanel, UiUtil, HRichOLE, ADODB ;

Procedure VisuLog1 ;


type
  TFMulLog = class(TFMul)
    TRB_FAMILLERUB: THLabel;
    LG_UTILISATEUR: THValComboBox;
    TE_DATECREATION: THLabel;
    LG_DATE: THCritMaskEdit;
    TE_DATECREATION_: THLabel;
    LG_DATE_: THCritMaskEdit;
    HLabel1: THLabel;
    LG_TEMPS: THCritMaskEdit;
    HLabel2: THLabel;
    LG_TEMPS_: THCritMaskEdit;
    BPurge: TToolbarButton97;
    CapPurge: TEdit;
    procedure FormShow(Sender: TObject);
    procedure BPurgeClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}
Uses Purge ;

Procedure VisuLog1 ;
var FMulLog:TFMulLog ;
    PP   : THPanel ;
BEGIN
if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FMulLog:=TFMulLog.Create(Application) ;
FMulLog.FNomFiltre:='MULLOG' ;
FMulLog.Q.Liste:='MULLOG' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   Try
    FMulLog.ShowModal ;
   Finally
    FMulLog.Free ;
   End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FMulLog,PP) ;
   FMulLog.Show ;
   END ;
END ;


procedure TFMulLog.FormShow(Sender: TObject);
begin
LG_DATE.text:=StDate1900 ; LG_DATE_.text:=StDate2099 ;
  inherited;
BOuvrir.Enabled:=False ;
Pages.Pages[1].TabVisible:=FALSE ;
//08/12/2006 YMO Norme NF
BPurge.Visible:=False;
end;

procedure TFMulLog.BPurgeClick(Sender: TObject);
Var dd : TDateTime ;
begin
  inherited;
if PurgeOui('',CapPurge.Text,dd) then
   BEGIN
   ExecuteSql('Delete From LOG Where LG_DATE<="'+UsDateTime(dd+1)+'"') ;
   BChercheClick(Nil) ;
   END ;
end;

procedure TFMulLog.BAideClick(Sender: TObject);
begin
  inherited;
CallHelpTopic(Self) ;
end;

procedure TFMulLog.QAfterOpen(DataSet: TDataSet);
Var F : TField ;
begin
  inherited;
F:=Q.FindField('LG_TEMPS') ;
if F<>nil then TDateTimeField(F).DisplayFormat:='nn:ss' ;
end;

end.
