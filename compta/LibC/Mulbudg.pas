{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par MULBUDG_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit MulBudg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Mask, Hctrls, Menus, DB, DBTables, Hqry, Grids, DBGrids,
  ExtCtrls, ComCtrls, Buttons, hmsgbox, HEnt1, Ent1, Budgene, Spin, MZSUtil,
  HRichEdt, HSysMenu, HDB, Hcompte, HTB97, ColMemo, HPanel, UiUtil,
  HRichOLE ;

Procedure MulticritereBudgene(Comment : TActionFiche) ;

type
  TFMulBudG = class(TFMul)
    TBG_BUDGENE: THLabel;
    BG_BUDGENE: TEdit;
    NePasVirer: TComboBox;
    BG_LIBELLE: TEdit;
    TBG_LIBELLE: THLabel;
    HM: THMsgBox;
    TBG_DATEMODIFICATION: THLabel;
    BG_DATEMODIF: THCritMaskEdit;
    TBG_DATEMODIFICATION2: THLabel;
    BG_DATEMODIF_: THCritMaskEdit;
    TBG_SENS: THLabel;
    BG_SENS: THValComboBox;
    PzLibre: TTabSheet;
    Bevel5: TBevel;
    TBG_TABLE0: TLabel;
    TBG_TABLE1: TLabel;
    TBG_TABLE2: TLabel;
    TBG_TABLE3: TLabel;
    TBG_TABLE4: TLabel;
    TBG_TABLE5: TLabel;
    TBG_TABLE6: TLabel;
    TBG_TABLE7: TLabel;
    TBG_TABLE8: TLabel;
    TBG_TABLE9: TLabel;
    BG_TABLE0: THCpteEdit;
    BG_TABLE1: THCpteEdit;
    BG_TABLE2: THCpteEdit;
    BG_TABLE3: THCpteEdit;
    BG_TABLE4: THCpteEdit;
    BG_TABLE5: THCpteEdit;
    BG_TABLE6: THCpteEdit;
    BG_TABLE7: THCpteEdit;
    BG_TABLE8: THCpteEdit;
    BG_TABLE9: THCpteEdit;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation


{$R *.DFM}

Procedure MulticritereBudgene(Comment : TActionFiche) ;
var FMulBudG : TFMulBudG;
    PP : THPanel ;
begin
FMulBudG:=TFMulBudG.Create(Application) ;
FMulBudG.TypeAction:=Comment ;
Case Comment Of
  taConsult      : begin
                   FMulBudG.Caption:=FMulBudG.HM.Mess[0] ;
                   FMulBudG.FNomFiltre:='MULVBUDG' ;
                   FMulBudG.Q.Liste:='MULVBUDG' ;
                   FMulBudG.HelpContext:=15111000 ;
                   end ;
  taModif        : begin
                   FMulBudG.Caption:=FMulBudG.HM.Mess[1];
                   FMulBudG.FNomFiltre:='MULMBUDG' ;
                   FMulBudG.Q.Liste:='MULMBUDG' ;
                   FMulBudG.HelpContext:=15115000 ;
                   end ;
  taModifEnSerie : begin
                   FMulBudG.Caption:=FMulBudG.HM.Mess[2] ;
                   FMulBudG.FNomFiltre:='MULMBUDG' ;
                   FMulBudG.Q.Liste:='MULMBUDG' ;
                   FMulBudG.HelpContext:=15117000 ;
                   end ;
  end ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FMulBudG.Caption:=FMulBudG.HM.Mess[9] ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulBudG.ShowModal ;
    finally
     FMulBudG.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulBudG,PP) ;
   FMulBudG.Show ;
   END ;
end ;

procedure TFMulBudG.FormShow(Sender: TObject);
begin
Pages.Pages[2].TabVisible:=False ;
if FCompte<>'' then BG_BUDGENE.text:=FCompte ;
if FLibelle<>'' then BG_Libelle.text:=FLibelle ;
BG_DATEMODIF.text:=StDate1900 ; BG_DATEMODIF_.text:=StDate2099 ;
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

procedure TFMulBudG.FListeDblClick(Sender: TObject);
begin
if(Q.Eof)And(Q.Bof) then Exit ;
  inherited;
if TypeAction<>taModifEnSerie then
   BEGIN
   FicheBudgene(Q,'',Q.FindField('BG_BUDGENE').AsString,TypeAction,0) ;
   if Typeaction<>taConsult then BChercheClick(Nil) ;
   END
   else if TControl(Sender).Name='FListe' then
           BEGIN
           FicheBudgene(Q,'',Q.FindField('BG_BUDGENE').AsString,TypeAction,0) ;
           Fliste.ClearSelected ;
           END else
           BEGIN
           if (Fliste.NbSelected>0) or (FListe.AllSelected) then
              BEGIN
              ModifieEnSerie('BUDGENE','',FListe,Q) ;
              ChercheClick ;
              END ;
           END;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulBudG.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulBudG.HMTradBeforeTraduc(Sender: TObject);
begin
  inherited;
LibellesTableLibre(PzLibre,'TBG_TABLE','BG_TABLE','B') ;
end;

procedure TFMulBudG.BinsertClick(Sender: TObject);
begin
  inherited;
// MODIF PACK AVANCE
  FicheBudgene(Nil,'','',taCreatEnSerie,0) ;
end;

end.
