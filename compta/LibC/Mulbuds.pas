{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par MULBUDS_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit Mulbuds;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Mask, Hctrls, StdCtrls, hmsgbox, HSysMenu, Menus, DB, DBTables,
  Hqry, Grids, DBGrids, HDB, ComCtrls, HRichEdt, ExtCtrls, Buttons, Ent1, Hent1,
  Hcompte, HTB97, ColMemo, HPanel, UiUtil, HRichOLE ;

Procedure MulticritereBudsect(Comment : TActionFiche) ;

type
  TFMulBudS = class(TFMul)
    HM: THMsgBox;
    TBS_BUDSECT: THLabel;
    BS_BUDSECT: TEdit;
    TBS_LIBELLE: THLabel;
    BS_LIBELLE: TEdit;
    TBS_AXE: TLabel;
    BS_AXE: THValComboBox;
    TBS_DATEMODIF: THLabel;
    BS_DATEMODIF: THCritMaskEdit;
    TBS_DATEMODIF_: THLabel;
    BS_DATEMODIF_: THCritMaskEdit;
    PzLibre: TTabSheet;
    Bevel5: TBevel;
    TBS_TABLE0: TLabel;
    TBS_TABLE1: TLabel;
    TBS_TABLE2: TLabel;
    TBS_TABLE3: TLabel;
    TBS_TABLE4: TLabel;
    TBS_TABLE5: TLabel;
    TBS_TABLE6: TLabel;
    TBS_TABLE7: TLabel;
    TBS_TABLE8: TLabel;
    TBS_TABLE9: TLabel;
    BS_TABLE0: THCpteEdit;
    BS_TABLE1: THCpteEdit;
    BS_TABLE2: THCpteEdit;
    BS_TABLE3: THCpteEdit;
    BS_TABLE4: THCpteEdit;
    BS_TABLE5: THCpteEdit;
    BS_TABLE6: THCpteEdit;
    BS_TABLE7: THCpteEdit;
    BS_TABLE8: THCpteEdit;
    BS_TABLE9: THCpteEdit;
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

Uses Budsect, MZSUtil ;

Procedure MulticritereBudsect(Comment : TActionFiche) ;
var FMulBudS : TFMulBudS ;
    PP : THPanel ;
begin
FMulBudS:=TFMulBudS.Create(Application) ;
FMulBudS.TypeAction:=Comment ;
Case Comment Of
  taConsult      : begin
                   FMulBudS.Caption:=FMulBudS.HM.Mess[0] ;
                   FMulBudS.FNomFiltre:='MULVBUDS' ;
                   FMulBudS.Q.Liste:='MULVBUDS' ;
                   FMulBudS.HelpContext:=15131000 ;
                   end ;
  taModif        : begin
                   FMulBudS.Caption:=FMulBudS.HM.Mess[1];
                   FMulBudS.FNomFiltre:='MULMBUDS' ;
                   FMulBudS.Q.Liste:='MULMBUDS' ;
                   FMulBudS.HelpContext:=15135000 ;
                   end ;
  taModifEnSerie : begin
                   FMulBudS.Caption:=FMulBudS.HM.Mess[2] ;
                   FMulBudS.FNomFiltre:='MULMBUDS' ;
                   FMulBudS.Q.Liste:='MULMBUDS' ;
                   FMulBudS.HelpContext:=15137000 ;
                   end ;
  end ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FMulBudS.Caption:=FMulBudS.HM.Mess[9] ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulBudS.ShowModal ;
    finally
     FMulBudS.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulBudS,PP) ;
   FMulBudS.Show ;
   END ;
end ;

procedure TFMulBudS.FormShow(Sender: TObject);
begin
Pages.Pages[2].TabVisible:=False ;
if FCompte<>'' then BS_BUDSECT.text:=FCompte ;
if FLibelle<>'' then BS_LIBELLE.text:=FLibelle ;
BS_DATEMODIF.text:=StDate1900 ; BS_DATEMODIF_.text:=StDate2099 ;
BS_AXE.ItemIndex:=0 ;
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

procedure TFMulBudS.FListeDblClick(Sender: TObject);
begin
if(Q.Eof)And(Q.Bof) then Exit ;
  inherited;
if TypeAction<>taModifEnSerie then
   BEGIN
   FicheBudsect(Q,Q.FindField('BS_AXE').AsString,Q.FindField('BS_BUDSECT').AsString,TypeAction,0) ;
   if Typeaction<>taConsult then BChercheClick(Nil) ;
   END
   else if TControl(Sender).Name='FListe' then
           BEGIN
           FicheBudsect(Q,Q.FindField('BS_AXE').AsString,Q.FindField('BS_BUDSECT').AsString,TypeAction,0) ;
           Fliste.ClearSelected ;
           END else
           BEGIN
           if (Fliste.NbSelected>0) or (FListe.AllSelected) then
              BEGIN
              ModifieEnSerie('BUDSECT',Q.FindField('BS_AXE').AsString,FListe,Q) ;
              ChercheClick ;
              END ;
           END;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulBudS.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulBudS.HMTradBeforeTraduc(Sender: TObject);
begin
  inherited;
LibellesTableLibre(PzLibre,'TBS_TABLE','BS_TABLE','D') ;
end;

procedure TFMulBudS.BinsertClick(Sender: TObject);
begin
  inherited;
// MODIF PACK AVANCE
  FicheBudsect(Nil,'','',taCreatEnSerie,0) ;
end;

end.
