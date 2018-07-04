unit RUPTURE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBCtrls, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Grids, DBGrids, HDB, StdCtrls, Buttons, ExtCtrls,
  hmsgbox ,Ent1, hcompte, HCtrls, HEnt1, HSysMenu, Hqry, HPanel, UiUtil,
  HTB97, ADODB ;

Procedure PlanRupture(LeCode : String) ;

type
  TFRupture = class(TForm)
    BDefaire: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    Fliste   : THDBGrid;
    SChoixcod: TDataSource;
    TChoixcod: THTable;
    DBNav    : TDBNavigator;
    BDplanRup: TToolbarButton97;
    HPB: TToolWindow97;
    MsgBox   : THMsgBox;
    TChoixcodCC_TYPE   : TStringField;
    TChoixcodCC_CODE   : TStringField;
    TChoixcodCC_LIBELLE: TStringField;
    TChoixcodCC_LIBRE  : TStringField;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock: TDock97;
    procedure BDplanRupClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BDefaireClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SChoixcodDataChange(Sender: TObject; Field: TField);
    procedure FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure TChoixcodNewRecord(DataSet: TDataSet);
    procedure TChoixcodBeforeDelete(DataSet: TDataSet);
    procedure BImprimerClick(Sender: TObject);
    procedure FlisteDblClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TChoixcodPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure SChoixcodUpdateData(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
   Lequel   : String ;
   Acreer   : Boolean ;
   Agenerer : Boolean ;
   WMinX,WMinY : Integer ;
   procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
   Function  Bouge(Button: TNavigateBtn) : boolean ;
   Function  OnSauve : boolean ;
   Procedure NewEnreg ;
   Procedure ChargeEnreg ;
   Function  Supprime : Boolean ;
   Function  EnregOK : boolean ;
   Function  VerifiSiExiste : Boolean ;
   Procedure CreationRupture ;
   Procedure FaitLeCaption ;
   Procedure GeleLesBoutons ;
   Procedure GenereRupture ;
   Procedure AppelAvertirTable ;
   Procedure BrancheHelpContext ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses printDBG, RupAnal, DRupture, UtilPgi ;

Procedure TFRupture.GeleLesBoutons ;
BEGIN
BInsert.Enabled:=(Not TChoixCod.Modified)And(TChoixCod.State<>dsInsert) ;
BDelete.Enabled:=(Not TChoixCod.Modified)And(TChoixCod.State<>dsInsert) ;
BDplanRup.Enabled:=(Not TChoixCod.Modified)And(TChoixCod.State<>dsInsert) ;
if TChoixCod.Eof And TChoixCod.Bof then BEGIN BDelete.Enabled:=False ; BDplanRup.Enabled:=False ; END ;
if TChoixCod.State=dsBrowse then FListe.Columns.Items[0].ReadOnly:=True ;
END ;

procedure TFRupture.SChoixcodDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
GeleLesBoutons ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TChoixCod.BOF;
   DnEnable := Enabled and not TChoixCod.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
//   if ((Field.FieldName='CC_LIBELLE') and (CC_ABREGE.Field.AsString='')) then
//   CC_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFRupture.FlisteKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if((ssCtrl in Shift)AND(Key=VK_DELETE)AND(TChoixCod.Eof)AND(TChoixCod.Bof)) then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift)AND(Key=VK_DELETE))then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFRupture.BDplanRupClick(Sender: TObject);
begin DetailPlanrupture(Lequel,TChoixcodCC_CODE.AsString,taModif) ; end;

procedure TFRupture.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFRupture.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFRupture.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFRupture.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFRupture.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFRupture.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFRupture.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil,Caption,''); end;

procedure TFRupture.BDefaireClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFRupture.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFRupture.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFRupture.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

Function TFRupture.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
     nblast,nbprior,nbnext,nbfirst,
     nbinsert : if Not OnSauve  then Exit ;
     nbPost   : if Not EnregOK  then Exit ;
     nbDelete : if Not Supprime then Exit ;
     end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.Mess[6]) ;
if Button = nbPost then BEGIN CreationRupture ; Acreer:=False ; END ;
GenereRupture ; Result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFRupture.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TChoixCod.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.Execute(0,caption,'') ;
   END else Rep:=321 ;
Case rep of
 mrYes : if not Bouge(nbPost) then exit ;
 mrNo  : if not Bouge(nbCancel) then exit ;
 mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function  TFRupture.Supprime : Boolean ;
BEGIN
Result:=False ;
if MsgBox.Execute(1,caption,'')=mrYES then BEGIN if NOT(TChoixCod.EOF And TChoixCod.BOF)then Result:=True END ;
END ;

procedure TFRupture.TChoixcodBeforeDelete(DataSet: TDataSet);
begin
ExecuteSQL('Delete from RUPTURE Where RU_NATURERUPT="'+Lequel+'" '+
           'AND RU_PLANRUPT="'+TChoixcodCC_CODE.AsString+'"') ;
end;

Procedure TFRupture.ChargeEnreg ;
BEGIN
(*************)
END ;

Procedure TFRupture.NewEnreg ;
BEGIN
InitNew(TChoixCod) ;
TChoixcodCC_TYPE.AsString:=Lequel ; Acreer:=True ;
FListe.Columns.Items[0].ReadOnly:=False ;
FListe.SelectedIndex:=0 ; FListe.SetFocus ;
END ;

Function TFRupture.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TChoixCod.state in [dsEdit,dsInsert]=False then Exit ;
if TChoixCod.state in [dsEdit,dsInsert] then
   BEGIN
   if TChoixcodCC_CODE.AsString='' then
      BEGIN MsgBox.Execute(2,caption,'') ; FListe.SetFocus ; FListe.SelectedIndex:=0 ; Exit ; END ;
   if TChoixcodCC_LIBELLE.AsString='' then
      BEGIN MsgBox.Execute(3,caption,'') ; FListe.SetFocus ; FListe.SelectedIndex:=1 ;Exit ; END ;
   END ;
if TChoixCod.state in [dsInsert] then
   BEGIN
    if VerifiSiExiste then
       BEGIN
       MsgBox.Execute(4,caption,'') ; FListe.SetFocus ;
       FListe.SelectedIndex:=0 ; Exit ;
       END ;
   END ;
Result:=TRUE  ;
END ;

Function  TFRupture.VerifiSiExiste : Boolean ;
BEGIN
Result:=ExisteSQL('Select CC_TYPE,CC_CODE from CHOIXCOD '+
                  ' Where CC_TYPE="'+LeQuel+'" AND CC_CODE="'+TChoixcodCC_CODE.AsString+'"') ;
END ;

procedure TFRupture.TChoixcodNewRecord(DataSet: TDataSet);
begin NewEnreg ; end ;

Procedure TFRupture.FaitLeCaption ;
BEGIN
if Lequel<>'' then
   BEGIN
   Case Lequel[3] of
        'G' : Caption:=Caption+MsgBox.Mess[8]  ;
        'T' : Caption:=Caption+MsgBox.Mess[9]  ;
        '1' : Caption:=Caption+MsgBox.Mess[10] ;
        '2' : Caption:=Caption+MsgBox.Mess[11] ;
        '3' : Caption:=Caption+MsgBox.Mess[12] ;
        '4' : Caption:=Caption+MsgBox.Mess[13] ;
        '5' : Caption:=Caption+MsgBox.Mess[14] ;
        'B' : Caption:=Caption+MsgBox.Mess[18] ;
     End ;
   END ;
UpdateCaption(Self) ;    
END ;

procedure TFRupture.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; BrancheHelpContext ;
TChoixcod.Open ; if Lequel<>''then TChoixcod.SetRange([Lequel],[Lequel]) ;
FaitLeCaption ; Acreer:=False ; Agenerer:=False ;
end;

procedure TFRupture.FlisteDblClick(Sender: TObject);
Var Q : TQuery ;
    Clibre,CLibreMemo : String ;
begin
if (Sender<>NIL)AND(TChoixCod.state in [dsEdit,dsInsert]) then Exit ;
if (Lequel='RUT') OR (Lequel='RUG') then
    BEGIN
    if ACreer then BEGIN Agenerer:=True ; GenereRupture ; END ;
    END else
    BEGIN
    Q:=OpenSql('Select CC_LIBRE from CHOIXCOD Where CC_TYPE="'+Lequel+'" AND '+
               'CC_CODE="'+TChoixcodCC_CODE.AsString+'"',True) ;
    CLibreMemo:=Q.Fields[0].AsString ; CLibre:=CLibreMemo ;
    Ferme(Q) ;
    RuptureAnalytique(Lequel,TChoixcodCC_CODE.AsString,Clibre) ;
    if CLibre<>CLibreMemo then
       BEGIN
       Agenerer:=True ;
       if Acreer then Acreer:=False ;
       if TChoixCod.state in [dsBrowse] then TChoixCod.Edit ;
       TChoixcodCC_LIBRE.AsString:=CLibre ;
       if Agenerer then Bouge(nbPost) ;
       END ;
    END ;
end;

Procedure TFRupture.CreationRupture ;
BEGIN
if Acreer then
   BEGIN
    if((Lequel='RU1')Or(Lequel='RU2')Or(Lequel='RU3')Or
      (Lequel='RU4')Or(Lequel='RU5'))then
      if MsgBox.Execute(5,caption,'')<>mrYes then BEGIN Acreer:=False ; Exit ; END ;
    if Lequel='RUG' then
       if MsgBox.Execute(16,caption,'')<>mrYes then BEGIN Acreer:=False ; Exit ; END ;
    if Lequel='RUT' then
       if MsgBox.Execute(17,caption,'')<>mrYes then BEGIN Acreer:=False ; Exit ; END ;
    FlisteDblClick(Nil) ;
 //   if(Lequel<>'RUG')And(Lequel<>'RUT')then TChoixCod.Edit ;
   END ;
END ;

Procedure PlanRupture(LeCode : String) ;
var FRupture: TFRupture;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FRupture:=TFRupture.Create(Application) ;
FRupture.Lequel:=LeCode ;
if LeCode='RUG' then FRupture.HelpContext := 1370000 else
if LeCode='RUT' then FRupture.HelpContext := 1380000 else FRupture.HelpContext := 139500 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     if FRupture.Lequel<>'' then FRupture.ShowModal ;
    Finally
     FRupture.Free ;
    End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FRupture,PP) ;
   FRupture.Show ;
   END ;
END ;

procedure TFRupture.FormClose(Sender: TObject; var Action: TCloseAction);
begin
AppelAvertirTable ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFRupture.TChoixcodPostError(DataSet: TDataSet;
  E: EDatabaseError; var Action: TDataAction);
begin
if TChoixCod.State=dsInsert then
   BEGIN MsgBox.Execute(4,caption,'') ; FListe.SelectedIndex:=0 ; Action:=daAbort ; END ;
end;

procedure TFRupture.SChoixcodUpdateData(Sender: TObject);
begin
if(Trim(TChoixCodCC_CODE.AsString)='') And
  (Trim(TChoixCodCC_LIBELLE.AsString)='') then BEGIN TChoixCod.Cancel ; Exit ; END ;
if(Trim(TChoixCodCC_CODE.AsString)='') And (Trim(TChoixCodCC_LIBELLE.AsString)<>'') then
  BEGIN MsgBox.Execute(2,caption,'') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TChoixCodCC_CODE.AsString)<>'') And (Trim(TChoixCodCC_LIBELLE.AsString)='') then
  BEGIN MsgBox.Execute(3,caption,'') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
end;

Procedure TFRupture.GenereRupture ;
BEGIN
if Agenerer then
   BEGIN
   DetailPlanRupture(Lequel,TChoixcodCC_CODE.AsString ,taCreat) ;
   Agenerer:=False ;
   END ;
END ;

procedure TFRupture.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFRupture.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

Procedure TFRupture.AppelAvertirTable ;
BEGIN
Case Lequel[3] of
     'G' : AvertirTable('ttRuptGene') ;
     'T' : AvertirTable('ttRuptTiers') ;
     '1' : AvertirTable('ttRuptSect1') ;
     '2' : AvertirTable('ttRuptSect2') ;
     '3' : AvertirTable('ttRuptSect3') ;
     '4' : AvertirTable('ttRuptSect4') ;
     '5' : AvertirTable('ttRuptSect5') ;
   End ;
END ;

Procedure TFRupture.BrancheHelpContext ;
BEGIN
Case Lequel[3] of
     'G'      : HelpContext:=1370000 ;
     'T'      : HelpContext:=1380000 ;
     '1'..'5' : HelpContext:=1395000 ;
   End ;
END ;

procedure TFRupture.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

end.
