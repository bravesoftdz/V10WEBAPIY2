// Remplacé par UTOTVENTILTYPE.PAS
unit VENTTYPE;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons, Sysutils,
     StdCtrls, ExtCtrls, Grids, DBGrids, Hctrls, DB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Hqry, DBIErrs,
     Dialogs, DBItypes, DBIProcs, HDB, DBCtrls, hmsgbox, ent1, HEnt1, Messages,
     HSysMenu, HTB97, HPanel, UiUtil, ADODB ;

Procedure ParamVentilType ;

type
  TFVentilType = class(TForm)
    SChoixCod: TDataSource;
    TChoixCod: THTable;
    FListe   : THDBGrid;
    DBNav    : TDBNavigator;
    HPB: TToolWindow97;
    BDefaire: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    MsgBox   : THMsgBox;
    BVentil: TToolbarButton97;
    TChoixCodCC_TYPE   : TStringField;
    TChoixCodCC_CODE   : TStringField;
    TChoixCodCC_LIBELLE: TStringField;
    TChoixCodCC_ABREGE : TStringField;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    Dock: TDock97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
    procedure SChoixCodDataChange(Sender: TObject; Field: TField);
    procedure TChoixCodNewRecord(DataSet: TDataSet);
    procedure BDefaireClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure TChoixCodAfterPost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TChoixCodAfterDelete(DataSet: TDataSet);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TChoixCodPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure SChoixCodUpdateData(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private  { Private declarations }
    FAvertir: Boolean ;
    ACreer,AGenerer : Boolean ;
    WMinX,WMinY : Integer ;
    MemoCpte : String ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure NewEnreg ;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure ChargeEnreg ;
    Function  VerifiSiExiste : Boolean ;
    Function  Supprime : Boolean ;
    Procedure Detruit ;
  public   { Public declarations }
    FQuoi : String ;
  end;


implementation

uses VENTIL, PrintDBG, UtilPgi;

{$R *.DFM}

Procedure ParamVentilType ;
var FVentilType: TFVentilType ;
    PP : THpanel ;
begin
if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FVentilType:=TFVentilType.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FVentilType.ShowModal ;
    finally
     FVentilType.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FVentilType,PP) ;
   FVentilType.Show ;
   END ;
END ;


procedure TFVentilType.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFVentilType.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFVentilType.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFVentilType.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFVentilType.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFVentilType.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFVentilType.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFVentilType.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFVentilType.NewEnreg ;
BEGIN
TChoixCod.FindField('CC_TYPE').AsString:='VTY' ;
FListe.Columns.Items[0].ReadOnly:=False ;
FListe.SelectedIndex:=0 ; FListe.SetFocus ;
END ;

procedure TFVentilType.FormCloseQuery(Sender: TObject;var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

Function TFVentilType.EnregOK : boolean ;
BEGIN
Result:=FALSE  ;
if TChoixCod.state in [dsEdit,dsInsert]=False then Exit ;
if TChoixCod.state in [dsEdit,dsInsert] then
   BEGIN
   if TChoixCodCC_CODE.AsString='' then
      BEGIN MsgBox.Execute(2,'','') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ; END ;
   if TChoixCodCC_LIBELLE.AsString='' then
      BEGIN MsgBox.Execute(3,'','') ; FListe.SelectedIndex:=1 ; FListe.SetFocus ; Exit ; END ;
   END ;
if TChoixCod.state in [dsInsert] then
   if VerifiSiExiste then
      BEGIN MsgBox.Execute(4,'','') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ; END ;
Result:=TRUE  ;
END ;

Function TFVentilType.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TChoixCod.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFVentilType.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : if Not EnregOK then Exit ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if(TChoixCod.State=dsInsert)And(Button=nbPost)then ACreer:=True else ACreer:=False ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.Mess[5]) ;
if ACreer And (Not AGenerer) then
   BEGIN
   ParamVentil('TY',TChoixCodCC_Code.AsString,'12345',taModif,True) ; Acreer:=False ; AGenerer:=False ;
   END ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Procedure TFVentilType.ChargeEnreg ;
BEGIN
//if ACreer then BEGIN ParamVentil('TY',TChoixCodCC_Code.AsString,'12345',taModif,True) ; Acreer:=False ; END ;
END ;

procedure TFVentilType.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TChoixCod.Open ; FAvertir:=FALSE ; ACreer:=False ; AGenerer:=False ;
TChoixCod.SetRange(['VTY'],['VTY']) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 18/06/2002
Modifié le ... :   /  /    
Description .. : redimmension de la grille
Mots clefs ... : 
*****************************************************************}
procedure TFVentilType.SChoixCodDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=(Not(TChoixCod.State in [dsEdit,dsInsert])) ;
BDelete.Enabled:=(Not(TChoixCod.State in [dsEdit,dsInsert])) ;
BVentil.Enabled:=(Not(TChoixCod.State in [dsEdit,dsInsert])) ;
if TChoixCod.State=dsBrowse then FListe.Columns.Items[0].ReadOnly:=True ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TChoixCod.BOF;
   DnEnable := Enabled and not TChoixCod.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
   if ((Copy(Field.FieldName,3,8)='_LIBELLE') and (TChoixCod.FindField('CC_ABREGE').AsString='')) then
      TChoixCod.FindField('CC_ABREGE').AsString:=Copy(Field.AsString,1,17) ;
   END ;
HMTrad.ResizeDBGridColumns(FListe) ;
end;

procedure TFVentilType.TChoixCodNewRecord(DataSet: TDataSet);
begin Newenreg ; end;

procedure TFVentilType.BDefaireClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFVentilType.BVentilClick(Sender: TObject);
begin ParamVentil('TY',TChoixCodCC_Code.AsString,'12345',taModif,True) ; end;

procedure TFVentilType.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil,Caption,'') ; end;

procedure TFVentilType.TChoixCodAfterPost(DataSet: TDataSet);
begin
if AGenerer and (Not ACreer) then
   BEGIN
   if MemoCpte<>'' then
   ParamVentil('TY',TChoixCodCC_Code.AsString,'12345',taModif,True) ; Acreer:=False ; AGenerer:=False ;
   END ;
FAvertir:=TRUE ;
end;

procedure TFVentilType.FormClose(Sender: TObject;var Action: TCloseAction);
begin
if FAvertir then AvertirTable ('ttVentilType') ;
if IsInside(Self) then Action:=caFree ;
end;

procedure TFVentilType.TChoixCodAfterDelete(DataSet: TDataSet);
begin FAvertir:=TRUE ; end;

procedure TFVentilType.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if((ssCtrl in Shift)AND(Key=VK_DELETE)AND(TChoixCod.Eof)AND(TChoixCod.Bof))then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift)AND(Key=VK_DELETE))then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFVentilType.TChoixCodPostError(DataSet: TDataSet;
  E: EDatabaseError; var Action: TDataAction);
begin
if TChoixCod.State=dsInsert then
   BEGIN MsgBox.Execute(4,'','') ; FListe.SelectedIndex:=0 ; Action:=daAbort ; END ;
end;

Function TFVentilType.VerifiSiExiste : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CC_CODE from CHOIXCOD Where CC_TYPE="VTY" And '+
           'CC_CODE="'+TChoixCodCC_CODE.AsString+'"',True) ;
Result:=Not Q.Eof ; Ferme(Q) ;
END ;

procedure TFVentilType.SChoixCodUpdateData(Sender: TObject);
begin
if(Trim(TChoixCodCC_CODE.AsString)='')And
  (Trim(TChoixCodCC_LIBELLE.AsString)='')And
  (Trim(TChoixCodCC_ABREGE.AsString)='') then BEGIN TChoixCod.Cancel ; Exit ; END ;
if(Trim(TChoixCodCC_CODE.AsString)='') And (Trim(TChoixCodCC_LIBELLE.AsString)<>'') then
  BEGIN MsgBox.Execute(2,'','') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TChoixCodCC_CODE.AsString)<>'') And (Trim(TChoixCodCC_LIBELLE.AsString)='') then
  BEGIN MsgBox.Execute(3,'','') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TChoixCodCC_CODE.AsString)='') And (Trim(TChoixCodCC_ABREGE.AsString)<>'') then
  BEGIN MsgBox.Execute(2,'','') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if TChoixCod.State=dsInsert then
   BEGIN
   if ACreer then AGenerer:=False else BEGIN MemoCpte:=TChoixCodCC_CODE.AsString ; AGenerer:=True ; END ;
   END ;
end;

procedure TFVentilType.FListeRowEnter(Sender: TObject);
begin if TChoixCod.State=dsInsert then FListe.SelectedIndex:=0 ; end;

Procedure TFVentilType.Detruit ;
BEGIN
ExecuteSql('Delete From Ventil Where V_NATURE Like "TY%" And V_COMPTE="'+TChoixCodCC_CODE.AsString+'"') ;
END ;

Function TFVentilType.Supprime : Boolean ;
BEGIN
Result:=False ;
if MsgBox.Execute(1,'','')<>mrYes then Exit ;
Detruit ; Result:=True ;
END ;

procedure TFVentilType.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFVentilType.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=316 ;
end;

procedure TFVentilType.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.

 
