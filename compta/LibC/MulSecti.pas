unit MulSecti;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, Grids, DBGrids, StdCtrls, Hctrls, Ent1,
  ExtCtrls, ComCtrls, Buttons, HEnt1, CPSection_TOM, hmsgbox,
 {$IFNDEF CCMP}
 {$IFNDEF IMP}
 {$IFNDEF CCS3}
  MZSUtil,
 {$ENDIF}
 {$ENDIF}
 {$ENDIF}
  HRichEdt, HSysMenu, HDB, Hcompte, HTB97, ColMemo, HPanel, UiUtil, HRichOLE ;

Procedure MulticritereSection(Comment : TActionFiche);

type
  TFMulSection = class(TFMul)
    TS_SECTION: THLabel;
    S_SECTION: TEdit;
    TS_LIBELLE: THLabel;
    S_LIBELLE: TEdit;
    S_FERME: TCheckBox;
    S_AXE: THValComboBox;
    TS_AXE: THLabel;
    NePasVirer: TComboBox;
    HM: THMsgBox;
    TS_MAITREOEUVRE: THLabel;
    S_MAITREOEUVRE: TEdit;
    TS_CHANTIER: THLabel;
    S_CHANTIER: TEdit;
    TG_SENS: THLabel;
    S_SENS: THValComboBox;
    Pzlibre: TTabSheet;
    S_TABLE0: THCpteEdit;
    S_TABLE1: THCpteEdit;
    S_TABLE2: THCpteEdit;
    S_TABLE3: THCpteEdit;
    S_TABLE4: THCpteEdit;
    S_TABLE5: THCpteEdit;
    S_TABLE6: THCpteEdit;
    S_TABLE7: THCpteEdit;
    S_TABLE8: THCpteEdit;
    S_TABLE9: THCpteEdit;
    TS_TABLE0: TLabel;
    Bevel5: TBevel;
    TS_TABLE1: TLabel;
    TS_TABLE2: TLabel;
    TS_TABLE3: TLabel;
    TS_TABLE4: TLabel;
    TS_TABLE5: TLabel;
    TS_TABLE6: TLabel;
    TS_TABLE7: TLabel;
    TS_TABLE8: TLabel;
    TS_TABLE9: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BParamListeClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
  private
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses UtilPGI,UFonctionsCBP ; 

Procedure MulticritereSection(Comment : TActionFiche);
var FMulSection: TFMulSection;
    PP       : THPanel ;
begin
if Comment<>taConsult then if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FMulSection:=TFMulSection.Create(Application) ;
FMulSection.TypeAction:=Comment ;
FMulSection.Zoom:=TRUE ;
Case Comment Of
  taConsult :      begin
                   FMulSection.Caption:=FMulSection.HM.Mess[0] ;
                   FMulSection.FNomFiltre:='MULVSEC' ;
                   FMulSection.Q.Liste:='MULVSECTION' ;
                   FMulSection.HelpContext:=7172000 ;
                   end ;
  taModif :        begin
                   FMulSection.Caption:=FMulSection.HM.Mess[1] ;
                   FMulSection.FNomFiltre:='MULMSEC' ;
                   FMulSection.Q.Liste:='MULMSECTION' ;
                   FMulSection.HelpContext:=7178000 ;
                   end ;
  taModifEnSerie : begin
                   FMulSection.Caption:=FMulSection.HM.Mess[3] ;
                   FMulSection.FNomFiltre:='MULMSEC' ;
                   FMulSection.Q.Liste:='MULMSECTION' ;
                   FMulSection.HelpContext:=7181000 ;
                   end ;
  end ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FMulSection.Caption:=FMulSection.HM.Mess[10] ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
    FMulSection.ShowModal ;
   finally
    FMulSection.Free ;
   end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulSection,PP) ;
   FMulSection.Show ;
   END ;
end ;


procedure TFMulSection.FormShow(Sender: TObject);
begin
Pages.Pages[2].TabVisible:=FALSE ;
if FCompte<>'' then S_SECTION.text:=FCompte ;
if FLibelle<>'' then S_Libelle.text:=FLibelle ;
if TypeAction=taModifEnSerie then
   BEGIN
   FListe.MultiSelection := True ;
   BOuvrir.Hint:=HM.Mess[4] ;
   S_AXE.Vide:=False ;
   S_AXE.Reload  ;
   S_AXE.ItemIndex:=0 ;
   bSelectAll.Visible:=True ;
   END else
   BEGIN
   FListe.MultiSelection := False ;
   END;
if S_SECTION.CanFocus then S_SECTION.SetFocus ;
  inherited;
// MODIF PACK AVANCE
if ((ExJaiLeDroitConcept(TConcept(ccSecCreat),False)) and (TypeAction<>taConsult)) then BInsert.Visible:=True ;
{$IFDEF CCS3}
S_AXE.Visible:=False ; TS_AXE.Visible:=False ; 
S_CHANTIER.Visible:=False ; TS_CHANTIER.Visible:=False ;
S_MAITREOEUVRE.Visible:=False ; TS_MAITREOEUVRE.Visible:=False ;
{$ENDIF}
end;

procedure TFMulSection.FListeDblClick(Sender: TObject);
var Q1 : TQuery ;
    AA : TActionFiche ;
begin
if(Q.Eof) And (Q.Bof) then Exit ;
  inherited;
AA:=TypeAction ;
if TypeAction<>taModifEnSerie then
   BEGIN
   if (TypeAction=taModif) and ((S_Ferme.State=cbGrayed) or (S_Ferme.State=cbChecked)) then
     BEGIN
     Q1:=OpenSQL('SELECT S_FERME FROM SECTION WHERE S_SECTION="'+Q.FindField('S_SECTION').AsString+'" AND S_AXE="'+Q.FindField('S_AXE').AsString+'"',True) ;
     if Q1.Fields[0].AsString='X'then HM.Execute(8,'','') ;
     Ferme(Q1) ;
     END ;
   if ((TypeAction=taModif) and (Not ExJaiLeDroitConcept(TConcept(ccSecModif),False))) then AA:=taConsult ; 
   if V_PGI.MonoFiche then FicheSection(nil,Q.FindField('S_AXE').AsString,Q.FindField('S_SECTION').AsString,AA,0)
                   else FicheSection(Q,Q.FindField('S_AXE').AsString,Q.FindField('S_SECTION').AsString,AA,0) ;
   if typeaction<>taConsult then BChercheClick(Nil) ;
   END else
   if TControl(Sender).Name='FListe' then
      BEGIN
      FicheSection(Q,Q.FindField('S_AXE').AsString,Q.FindField('S_SECTION').AsString,taModif,0) ;
      Fliste.ClearSelected ;
      END else
      BEGIN
      if (Fliste.NbSelected>0) or (FListe.AllSelected) then
         BEGIN
        {$IFNDEF IMP}
        {$IFNDEF CCMP}
        {$IFNDEF CCS3}
         ModifieEnSerie('SECTION',Q.FindField('S_AXE').AsString,FListe,Q) ;
        {$ENDIF}
        {$ENDIF}
        {$ENDIF}
         ChercheClick ;
         END ;
      END ;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulSection.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulSection.BParamListeClick(Sender: TObject);
begin
  inherited;
Fliste.ClearSelected ; 
end;

procedure TFMulSection.HMTradBeforeTraduc(Sender: TObject);
begin
  inherited;
LibellesTableLibre(PzLibre,'TS_TABLE','S_TABLE','S') ;
end;

procedure TFMulSection.BinsertClick(Sender: TObject);
begin
  inherited;
// MODIF PACK AVANCE
  FicheSection(Nil,'A1','',taCreatEnSerie,0) ;
end;

end.
