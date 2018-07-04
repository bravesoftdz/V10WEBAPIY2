{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par MULBUDJ_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit Mulbudj;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Mask, Hctrls, StdCtrls, hmsgbox, HSysMenu, Menus, DB, DBTables,
  Hqry, Grids, DBGrids, HDB, ComCtrls, HRichEdt, ExtCtrls, Buttons, Ent1, Hent1,
  HTB97, ColMemo, HPanel, UiUtil, HRichOLE ;

Procedure MulticritereBudjal(Comment : TActionFiche) ;

type
  TFMulBudJ = class(TFMul)
    HM: THMsgBox;
    TBJ_BUDJAL: THLabel;
    BJ_BUDJAL: TEdit;
    TBJ_LIBELLE: THLabel;
    BJ_LIBELLE: TEdit;
    TBJ_AXE: TLabel;
    BJ_AXE: THValComboBox;
    TBJ_DATEMODIF: THLabel;
    BJ_DATEMODIF: THCritMaskEdit;
    TBJ_DATEMODIF_: THLabel;
    BJ_DATEMODIF_: THCritMaskEdit;
    HLabel1: THLabel;
    BJ_DATECREATION: THCritMaskEdit;
    HLabel2: THLabel;
    BJ_DATECREATION_: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BinsertClick(Sender: TObject);
  private
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses Budjal, MZSUtil ;

Procedure MulticritereBudjal(Comment : TActionFiche) ;
var FMulBudJ : TFMulBudJ ;
    PP : THPanel ;
begin
FMulBudJ:=TFMulBudJ.Create(Application) ;
FMulBudJ.TypeAction:=Comment ;
Case Comment Of
  taConsult      : begin
                   FMulBudJ.Caption:=FMulBudJ.HM.Mess[0] ;
                   FMulBudJ.FNomFiltre:='MULVBUDJ' ;
                   FMulBudJ.Q.Liste:='MULVBUDJ' ;
                   FMulBudJ.HelpContext:=15151000 ;
                   end ;
  taModif        : begin
                   FMulBudJ.Caption:=FMulBudJ.HM.Mess[1];
                   FMulBudJ.FNomFiltre:='MULMBUDJ' ;
                   FMulBudJ.Q.Liste:='MULMBUDJ' ;
                   FMulBudJ.HelpContext:=15155000 ;
                   end ;
  taModifEnSerie : begin
                   FMulBudJ.Caption:=FMulBudJ.HM.Mess[2] ;
                   FMulBudJ.FNomFiltre:='MULMBUDJ' ;
                   FMulBudJ.Q.Liste:='MULMBUDJ' ;
                   FMulBudJ.HelpContext:=15157000 ;
                   end ;
  end ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FMulBudJ.Caption:=FMulBudJ.HM.Mess[9] ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulBudJ.ShowModal ;
    finally
     FMulBudJ.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulBudJ,PP) ;
   FMulBudJ.Show ;
   END ;
end ;

procedure TFMulBudJ.FormShow(Sender: TObject);
begin
Pages.Pages[1].TabVisible:=False ;
if FCompte<>'' then BJ_BUDJAL.text:=FCompte ;
if FLibelle<>'' then BJ_LIBELLE.text:=FLibelle ;
BJ_DATEMODIF.Text:=StDate1900    ; BJ_DATEMODIF_.Text:=StDate2099 ;
BJ_DATECREATION.Text:=StDate1900 ; BJ_DATECREATION_.Text:=StDate2099 ;
BJ_AXE.ItemIndex:=0 ;
  inherited;
if TypeAction=taModifEnSerie then
   BEGIN
   FListe.MultiSelection := True ;
   BOuvrir.Hint:=HM.Mess[3] ;
   bSelectAll.Visible:=True ;
   END else
   BEGIN
   FListe.MultiSelection := False ;
   END;
// MODIF PACK AVANCE
if ((TypeAction<>taConsult)) then BInsert.Visible:=True ;
end;

procedure TFMulBudJ.FListeDblClick(Sender: TObject);
begin
if(Q.Eof)And(Q.Bof) then Exit ;
  inherited;
if TypeAction<>taModifEnSerie then
   BEGIN
   FicheBudjal(Q,Q.FindField('BJ_AXE').AsString,Q.FindField('BJ_BUDJAL').AsString,TypeAction,0) ;
   if Typeaction<>taConsult then BChercheClick(Nil) ;
   END
   else if TControl(Sender).Name='FListe' then
           BEGIN
           FicheBudjal(Q,Q.FindField('BJ_AXE').AsString,Q.FindField('BJ_BUDJAL').AsString,TypeAction,0) ;
           Fliste.ClearSelected ;
           END else
           BEGIN
           if (Fliste.NbSelected>0) or (FListe.AllSelected) then
              BEGIN
              ModifieEnSerie('BUDJAL',Q.FindField('BJ_AXE').AsString,FListe,Q) ;
              ChercheClick ;
              END ;
           END;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulBudJ.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulBudJ.BinsertClick(Sender: TObject);
begin
  inherited;
// MODIF PACK AVANCE
  FicheBudjal(Nil,'','',taCreatEnSerie,0) ;
end;

end.
