unit MulAbo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Mask, Hctrls, StdCtrls, Menus, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, Grids, DBGrids,
  ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, SaisUtil,
  //XMG 05/04/04 début
  //{$IFDEF ESP}
  CONTABON_TOM,
  (*{$ELSE}
  ContAbon, // A FAIRE
  {$ENDIF ESP}
  //XMG 05/04/04 fin*)
  HRichEdt,
  HSysMenu, HDB, HTB97, ColMemo, HmsgBox, HPanel, UiUtil, HRichOLE, ADODB ;

procedure ListeAbonnements ;

type
  TFMulAbo = class(TFMul)
    TG_CONTRAT: THLabel;
    CB_CONTRAT: TEdit;
    CB_RECONDUCTION: THValComboBox;
    TCB_RECONDUCTION: THLabel;
    CB_COMPTABLE: TCheckBox;
    TCB_ARRONDI: THLabel;
    CB_ARRONDI: THValComboBox;
    TCB_DEJAGENERE: THLabel;
    CB_DEJAGENERE: THCritMaskEdit;
    CB_NBREPETITION: THCritMaskEdit;
    TG_NOMBREREPETITION: THLabel;
    HLabel3: THLabel;
    CB_LIBELLE: TEdit;
    HLabel4: THLabel;
    CB_DATECREATION: THCritMaskEdit;
    HLabel6: THLabel;
    CB_DATECREATION_: THCritMaskEdit;
    CB_DATECONTRAT_: THCritMaskEdit;
    HLabel7: THLabel;
    CB_DATECONTRAT: THCritMaskEdit;
    HLabel8: THLabel;
    TG_DATEDERNGENERE: THLabel;
    CB_DATEDERNGENERE: THCritMaskEdit;
    HLabel5: THLabel;
    CB_DATEDERNGENERE_: THCritMaskEdit;
    CB_DATEMODIF_: THCritMaskEdit;
    HLabel9: THLabel;
    CB_DATEMODIF: THCritMaskEdit;
    HLabel10: THLabel;
    TCB_QUALIFPIECE: THLabel;
    CB_QUALIFPIECE: THValComboBox;
    procedure BChercheClick(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
  private
  public
  end;

implementation

{$R *.DFM}
Uses UtilPGI ;

procedure ListeAbonnements ;
Var FAbo : TFMulAbo ;
    PP   : THPanel ;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FAbo:=TFMulAbo.Create(Application) ;
FAbo.FNomFiltre:='MULABO' ;
FAbo.Q.Liste:='MULABO' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FAbo.ShowModal ;
    Finally
     FAbo.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FAbo,PP) ;
   FAbo.Show ;
   END ;
END ;

procedure TFMulAbo.BChercheClick(Sender: TObject);
begin
  inherited;
GereSelectionsGrid(FListe,Q) ;
end;

procedure TFMulAbo.FListeDblClick(Sender: TObject);
begin
  inherited;
if (Q.Eof)And(Q.Bof) then Exit ;
ParamAbonnement(True,Q.FindField('CB_CONTRAT').AsString,taConsult) ;
end;

procedure TFMulAbo.FormShow(Sender: TObject);
begin
CB_DATECREATION.text:=StDate1900 ; CB_DATECREATION_.text:=StDate2099 ;
CB_DATECONTRAT.text:=StDate1900 ; CB_DATECONTRAT_.text:=StDate2099 ;
CB_DATEDERNGENERE.text:=StDate1900 ; CB_DATEDERNGENERE_.text:=StDate2099 ;
CB_DATEMODIF.text:=StDate1900 ; CB_DATEMODIF_.text:=StDate2099 ;
CB_RECONDUCTION.ItemIndex:=0 ;
  inherited;
CB_COMPTABLE.Visible:=False ;
end;

procedure TFMulAbo.BOuvrirClick(Sender: TObject);
begin
FListeDblClick(Nil) ;
end;

procedure TFMulAbo.FormCreate(Sender: TObject);
begin
  // Pour passer dans le create.
  inherited;
end;

end.
