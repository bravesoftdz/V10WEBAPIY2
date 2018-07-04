unit CODEAFB;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons, Sysutils,
     StdCtrls, ExtCtrls, Grids, DBGrids, Hctrls, DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Hqry, DBIErrs,
     Dialogs, DBItypes, DBIProcs, HDB, DBCtrls, hmsgbox, ent1, HEnt1, Messages,
     HSysMenu, HPanel, UiUtil, HTB97, ADODB ;

Procedure ParamCodeAFB ;

type
  TFCodeAFB = class(TForm)
    SCodeAFB  : TDataSource;
    TCodeAFB: THTable;
    FListe    : THDBGrid;
    HPB: TToolWindow97;
    BAnnuler: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    DBNav     : TDBNavigator;
    MsgBox    : THMsgBox;
    TCodeAFBAF_CODEAFB     : TStringField;
    TCodeAFBAF_LIBELLE     : TStringField;
    TCodeAFBAF_MODEPAIEMENT: TStringField;
    TModePaie: THTable;
    TCodeAFBAF_MPLibelle : TStringField;
    TModePaieMP_MODEPAIE : TStringField;
    TModePaieMP_LIBELLE  : TStringField;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock: TDock97;
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
    procedure SCodeAFBDataChange(Sender: TObject; Field: TField);
    procedure TCodeAFBNewRecord(DataSet: TDataSet);
    procedure BAnnulerClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TCodeAFBPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure SCodeAFBStateChange(Sender: TObject);
    procedure SCodeAFBUpdateData(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure TCodeAFBAfterDelete(DataSet: TDataSet);
    procedure TCodeAFBAfterPost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private  { Private declarations }
    Modifier : Boolean ;
    FAvertir : boolean ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure NewEnreg ;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure ChargeEnreg ;
    Procedure ChargePickList ;
    Function  VerifiSiExiste : Boolean ;
  public   { Public declarations }

  end;


implementation

uses PrintDBG, UtilPgi;

{$R *.DFM}


procedure TFCodeAFB.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCodeAFB.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCodeAFB.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCodeAFB.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCodeAFB.BInsertClick(Sender: TObject);
begin
  Bouge(nbInsert) ;
end;

procedure TFCodeAFB.BDeleteClick(Sender: TObject);
begin
  if Not TCodeAFB.EOF then Bouge(nbDelete) ;
end;

procedure TFCodeAFB.BValiderClick(Sender: TObject);
begin Modifier:=False ; Bouge(nbPost) ; end;

procedure TFCodeAFB.BFermeClick(Sender: TObject);
begin
  Close ;
  {b FP FQ16067}
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
  {e FP FQ16067}
end;

procedure TFCodeAFB.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFCodeAFB.TCodeAFBNewRecord(DataSet: TDataSet);
begin newenreg ; end;

Procedure TFCodeAFB.NewEnreg ;
BEGIN
//InitNew(TSection) ;
FListe.SetFocus ;
END ;

procedure TFCodeAFB.FormCloseQuery(Sender: TObject;var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

Function  TFCodeAFB.VerifiSiExiste : Boolean ;
BEGIN
Result:=False ;
if Presence('MODEPAIE','MP_MODEPAIE',TCodeAFBAF_MODEPAIEMENT.AsString) then Result:=True ;
END ;

Function TFCodeAFB.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TCodeAFB.state in [dsEdit,dsInsert]=False then Exit ;
if TCodeAFB.state in [dsEdit,dsInsert] then
   BEGIN
   if TCodeAFBAF_CODEAFB.AsString='' then BEGIN MsgBox.Execute(2,'','') ; FListe.SetFocus ; Exit ; END ;
   if TCodeAFBAF_LIBELLE.AsString='' then BEGIN MsgBox.Execute(3,'','') ; FListe.SetFocus ; Exit ; END ;
   if TCodeAFBAF_MODEPAIEMENT.AsString<>''then
      BEGIN
      if Not VerifiSiExiste then BEGIN MsgBox.Execute(5,'','') ; FListe.SetFocus ; Exit ; END ;
      END ;
   END ;
if TCodeAFB.state in [dsInsert] then
   BEGIN
    if Presence('CODEAFB','AF_CODEAFB',TCodeAFBAF_CODEAFB.AsString) then
       BEGIN MsgBox.Execute(4,'','') ; Exit ; END ;
   END ;
result:=TRUE  ;
END ;

Function TFCodeAFB.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier:=False ;
if TCodeAFB.Modified then
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

Function TFCodeAFB.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,nbfirst,nbinsert :
     BEGIN
     if Not OnSauve then Exit ;
     END ;
   nbPost : if Not EnregOK then exit ;
   nbDelete : if MsgBox.execute(1,'','')<>mrYes then exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[6]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Procedure TFCodeAFB.ChargeEnreg ;
BEGIN
// code ...
END ;

procedure TFCodeAFB.SCodeAFBDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not TCodeAFB.Modified ; BDelete.Enabled:= Not TCodeAFB.Modified ;
if(TCodeAFB.Eof)And(TCodeAFB.Bof) then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TCodeAFB.BOF;
   DnEnable := Enabled and not TCodeAFB.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
   END ;
end;

procedure TFCodeAFB.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if((ssCtrl in Shift)and(Key=Vk_Delete)And(TCodeAFB.Eof))then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift)and(Key=Vk_Delete)) then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFCodeAFB.TCodeAFBPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
if TCodeAFB.State=dsInsert then
   BEGIN MsgBox.Execute(7,'','') ; Action:=daAbort ; END ;
end;

procedure TFCodeAFB.SCodeAFBUpdateData(Sender: TObject);
begin
if Modifier then BEGIN Modifier:=False ; OnSauve ; END ;
end;

procedure TFCodeAFB.SCodeAFBStateChange(Sender: TObject);
begin Modifier:=True ; end;

Procedure TFCodeAFB.ChargePickList ;
BEGIN
With FListe.Columns.Items[2] do
     BEGIN
      PickList.Clear ; TModePaie.First ;
      While Not TModePaie.EOF do
         BEGIN
         Picklist.Add(TModePaie.FindField('MP_MODEPAIE').AsString ); TModePaie.Next  ;
         END ;
     END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 18/06/2002
Modifié le ... :   /  /    
Description .. : redimensionnemetn de la grille
Mots clefs ... : 
*****************************************************************}
procedure TFCodeAFB.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FAvertir:=False ;
TCodeAFB.Open ; TModePaie.Open ; ChargePickList ;
HMTrad.ResizeDBGridColumns(FListe) ;
end;

Procedure ParamCodeAFB ;
var FCodeAFB : TFCodeAFB ;
    PP : THPanel ;
begin
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FCodeAFB:=TFCodeAFB.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FCodeAFB.ShowModal ;
    finally
     FCodeAFB.Free ;
    end ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FCodeAFB,PP) ;
   FCodeAFB.Show ;
   END ;
END ;

procedure TFCodeAFB.BImprimerClick(Sender: TObject);
begin
PrintDBGrid (FListe,Nil,Caption,'') ;
end;

procedure TFCodeAFB.TCodeAFBAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCodeAFB.TCodeAFBAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCodeAFB.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then AvertirTable('ttAFB') ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFCodeAFB.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCodeAFB.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFCodeAFB.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
