unit MulGene;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  ExtCtrls, ComCtrls, Buttons, Mask, Hent1,
  hmsgbox,
{$IFNDEF PGIIMMO}
{$IFNDEF CCMP}
{$IFNDEF IMP}
{$IFNDEF CCS3}
  MZSUtil,
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
  Spin,
  HRichEdt, Ent1, HSysMenu, HDB, Hcompte, HTB97, ColMemo, HPanel,UiUtil,
  HRichOLE,UtilPGI, ADODB ;

Procedure MultiCritereCpteGene(Comment : TActionFiche);

type
  TFMulgene = class(TFMul)
    TG_GENERAL: THLabel;
    G_GENERAL: TEdit;
    TG_LIBELLE: THLabel;
    G_LIBELLE: TEdit;
    TG_NATUREGENE: THLabel;
    G_NATUREGENE: THValComboBox;
    G_VENTILABLE: TComboBox;
    TG_VENTILABLE: THLabel;
    G_POINTABLE: TCheckBox;
    G_LETTRABLE: TCheckBox;
    G_FERME: TCheckBox;
    G_COLLECTIF: TCheckBox;
    TG_SENS: THLabel;
    G_SENS: THValComboBox;
    TG_ABREGE: THLabel;
    TG_DATEDERNMVT: THLabel;
    TG_DATEMODIFICATION: THLabel;
    G_DATEMODIF: THCritMaskEdit;
    G_DATEDERNMVT: THCritMaskEdit;
    G_DATEDERNMVT_: THCritMaskEdit;
    G_DATEMODIF_: THCritMaskEdit;
    TG_DATEMODIFICATION2: THLabel;
    TG_DATEDERNMVT2: THLabel;
    HM: THMsgBox;
    G_ABREGE: TEdit;
    Pzlibre: TTabSheet;
    G_TABLE0: THCpteEdit;
    G_TABLE1: THCpteEdit;
    G_TABLE2: THCpteEdit;
    G_TABLE3: THCpteEdit;
    G_TABLE4: THCpteEdit;
    G_TABLE5: THCpteEdit;
    G_TABLE6: THCpteEdit;
    G_TABLE7: THCpteEdit;
    G_TABLE8: THCpteEdit;
    G_TABLE9: THCpteEdit;
    TG_TABLE0: TLabel;
    Bevel5: TBevel;
    TG_TABLE1: TLabel;
    TG_TABLE2: TLabel;
    TG_TABLE3: TLabel;
    TG_TABLE4: TLabel;
    TG_TABLE5: TLabel;
    TG_TABLE6: TLabel;
    TG_TABLE7: TLabel;
    TG_TABLE8: TLabel;
    TG_TABLE9: TLabel;
    VentilS5: TComboBox;
    VentilS3: TCheckBox;
    procedure FListeDblClick(Sender: TObject); override;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BParamListeClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure VentilS5Change(Sender: TObject);
    procedure VentilS3Click(Sender: TObject);
  private
  public
  end;

implementation

uses CPGeneraux_TOM,UFonctionsCBP;

{$R *.DFM}

Procedure MultiCritereCpteGene(Comment : TActionFiche);
var FMulGene : TFMulGene;
    PP       : THPanel ;
begin
if Comment<>taConsult then if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FMulGene:=TFMulGene.Create(Application) ;
FMulGene.TypeAction:=Comment ;
Case Comment Of
  taConsult      : begin
                   FMulGene.Caption:=FMulGene.HM.Mess[0] ;
                   FMulGene.FNomFiltre:='MULVGENE' ;
                   FMulGene.Q.Liste:='MULVGENE' ;
                   FMulGene.HelpContext := 7106000 ;
                   end ;
  taModif        : begin
                   FMulGene.Caption:=FMulGene.HM.Mess[1];
                   FMulGene.FNomFiltre:='MULMGENE' ;
                   FMulGene.Q.Liste:='MULMGENE' ;
                   FMulGene.HelpContext := 7112000 ;
                   end ;
  taModifEnSerie : begin
                   FMulGene.Caption:=FMulGene.HM.Mess[2] ;
                   FMulGene.FNomFiltre:='MULMGENE' ;
                   FMulGene.Q.Liste:='MULMGENE' ;
                   FMulGene.HelpContext := 7115000 ;
                   end ;
  end ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FMulGene.Caption:=FMulGene.HM.Mess[9] ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   try
    FMulGene.ShowModal ;
   finally
    FMulGene.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(FMulGene,PP) ;
  FMulGene.Show ;
  END ;
end ;

procedure TFMulgene.FListeDblClick(Sender: TObject);
var Q1 : TQuery ;
    AA : TActionFiche ;
begin
if(Q.Eof) And (Q.Bof) then Exit ;
  inherited;
AA:=TypeAction ;
if TypeAction<>taModifEnSerie then
   BEGIN
   if (TypeAction=taModif) and ((G_Ferme.State=cbGrayed) or (G_Ferme.State=cbChecked)) then
     BEGIN
     Q1:=OpenSQL('SELECT G_FERME FROM GENERAUX WHERE G_GENERAL="'+Q.FindField('G_GENERAL').AsString+'"',True) ;
     if Q1.Fields[0].AsString='X'then HM.Execute(7,'','') ;
     Ferme(Q1) ;
     END ;
   if ((TypeAction=taModif) and (Not ExJaiLeDroitConcept(TConcept(ccGenModif),False))) then AA:=taConsult ;
   if V_PGI.MonoFiche then FicheGene(nil,'',Q.FindField('G_GENERAL').AsString,AA,0)
                      else FicheGene(Q,'',Q.FindField('G_GENERAL').AsString,AA,0) ;
   if TypeAction<>taConsult then BChercheClick(Nil) ;
   END
   else if TControl(Sender).Name='FListe' then
           BEGIN
           FicheGene(Q,'',Q.FindField('G_GENERAL').AsString,taModif,0) ;
           Fliste.ClearSelected ;
           END else
           BEGIN
           if (Fliste.NbSelected>0) or (FListe.AllSelected) then
              BEGIN
{$IFNDEF PGIIMMO}
{$IFNDEF CCMP}
{$IFNDEF IMP}
{$IFNDEF CCS3}
              ModifieEnSerie('GENERAUX','',FListe,Q) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
              ChercheClick ;
              END ;
           END;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulgene.FormShow(Sender: TObject);
begin
if FCompte<>'' then G_GENERAL.text:=FCompte ;
if FLibelle<>'' then G_Libelle.text:=FLibelle ;
G_DATEMODIF.text:=StDate1900 ; G_DATEMODIF_.text:=StDate2099 ;
G_DATEDERNMVT.text:=StDate1900 ; G_DATEDERNMVT_.text:=StDate2099 ;
  inherited;
if TypeAction=taModifEnSerie then
   BEGIN
   FListe.MultiSelection := True ; BOuvrir.Hint:=HM.Mess[3] ;
   bSelectAll.Visible:=True ;
   END else
   BEGIN
   FListe.MultiSelection := False ;
   END;
// MODIF PACK AVANCE
if ((ExJaiLeDroitConcept(TConcept(ccGenCreat),False)) and (TypeAction<>taConsult)) then BInsert.Visible:=True ;
// JLD5Axes
(**
if EstSerie(S5) then
   BEGIN
   VentilS5.Visible:=True ; G_Ventilable.Visible:=False ; VentilS3.Visible:=False ;
   TG_Ventilable.FocusControl:=VentilS5 ;
   END ;
**)
if EstSerie(S3) then
   BEGIN
   VentilS5.Visible:=False ; G_Ventilable.Visible:=False ; if not estcomptasansana then VentilS3.Visible:=True ;
   TG_Ventilable.Visible:=False ; 
   END ;
end;

procedure TFMulgene.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulgene.BParamListeClick(Sender: TObject);
begin
  inherited;
Fliste.ClearSelected ;
end;

procedure TFMulgene.HMTradBeforeTraduc(Sender: TObject);
begin
inherited;
LibellesTableLibre(PzLibre,'TG_TABLE','G_TABLE','G') ;
end;

procedure TFMulgene.BinsertClick(Sender: TObject);
begin
  inherited;
// MODIF PACK AVANCE
FicheGene(Nil,'','',taCreatEnSerie,0) ;
end;

procedure TFMulgene.VentilS5Change(Sender: TObject);
begin
Case VentilS5.ItemIndex of
    0..2 : G_Ventilable.ItemIndex:=VentilS5.ItemIndex ;
    else G_Ventilable.ItemIndex:=VentilS5.ItemIndex+3 ;
    END ;
end;

procedure TFMulgene.VentilS3Click(Sender: TObject);
begin
  inherited;
{$IFDEF CCS3}
Case VentilS3.State of
   cbChecked   : G_Ventilable.ItemIndex:=1 ;
   cbUnChecked : G_Ventilable.ItemIndex:=6 ;
   cbGrayed    : G_Ventilable.ItemIndex:=0 ;
   END ;
{$ENDIF}
end;

end.
