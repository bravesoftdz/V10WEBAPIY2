unit CORRESP;

interface

uses SysUtils, WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons,Dialogs,
     StdCtrls, ExtCtrls, Hctrls, {$IFNDEF DBXPRESS}dbtables, HSysMenu,
  hmsgbox, Db, Hqry, ComCtrls, DBCtrls, HTB97, Grids, DBGrids, HDB{$ELSE}uDbxDataSet{$ENDIF}, DB, Grids, DBGrids, Mask, DBCtrls,
     Ent1, Spin, HmsgBox, Hqry, HEnt1, ComCtrls, HDB, Messages, HSysMenu, HPanel, UiUtil,
  HTB97 ;

procedure ParamCorresp ( Quoi : String ) ;
Procedure ZoomCorresp  (Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);

type
  TFCorresp = class(TForm)
    FListe: THDBGrid;
    TCORRESP: THTable;
    SCORRESP : TDataSource;
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
    HPlan    : TPanel;
    Plan     : TTabControl;
    TCORRESPCR_TYPE   : TStringField;
    TCORRESPCR_CORRESP: TStringField;
    TCORRESPCR_LIBELLE: TStringField;
    TCORRESPCR_ABREGE : TStringField;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock: TDock97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure TCorrespNewRecord(DataSet: TDataset);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BDefaireClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure SCORRESPDataChange(Sender: TObject; Field: TField);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BImprimerClick(Sender: TObject);
    procedure TCORRESPAfterDelete(DataSet: TDataSet);
    procedure TCORRESPAfterPost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TCORRESPPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure PlanChange(Sender: TObject);
    procedure SCORRESPUpdateData(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    LePlan : integer ;
    FType : string ;
    FAvertir : boolean ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure NewEnreg ;
    Procedure ChargeEnreg ;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure FaitCaption ;
    Procedure Initialisation ;
    Function  Supprime : Boolean ;
    Function  VerifiSiExiste : Boolean ;
  public
  end;


implementation

uses PrintDBG, UtilPgi;
{$R *.DFM}


{---Procédure d'appel de la fiche CORRESP---------------------------------}
Procedure ZoomCorresp  (Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
var FCorresp: TFCorresp;
    PP : THPanel ;
begin
if _Blocage(['nrCloture'],False,'nrBatch') then exit ;
FCorresp:=TFCorresp.Create(Application) ;
FCorresp.FType:=Axe ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FCorresp.ShowModal ;
    finally
     FCorresp.Free ;
     _Bloqueur('nrBatch',False) ;
    end ;
   END else
   BEGIN
   InitInside(FCorresp,PP) ;
   FCorresp.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure ParamCorresp ( Quoi : String ) ;
var FCorresp: TFCorresp;
    PP : THPanel ;
begin
if _Blocage(['nrCloture'],False,'nrBatch') then exit ;
FCorresp:=TFCorresp.Create(Application) ;
FCorresp.FType:=Quoi ;
if Quoi='GE' then FCorresp.HelpContext := 1325000 ;
if Quoi='AU' then FCorresp.HelpContext := 1330000 ;
if (Quoi='A1') or (Quoi='A2') or (Quoi='A3') or (Quoi='A4') or (Quoi='A5') then FCorresp.HelpContext := 1134500 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FCorresp.ShowModal ;
    finally
     FCorresp.Free ;
     _Bloqueur('nrBatch',False) ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FCorresp,PP) ;
   FCorresp.Show ;
   END ;
end;

Procedure TFCorresp.FaitCaption ;
BEGIN
if FType='GE' then Caption:=Caption+' '+MsgBox.Mess[8] else
   if FType='AU' then Caption:=Caption+' '+MsgBox.Mess[9] else
      if FType='A1' then Caption:=Caption+' '+MsgBox.Mess[10] else
         if FType='A2' then Caption:=Caption+' '+MsgBox.Mess[11] else
            if FType='A3' then Caption:=Caption+' '+MsgBox.Mess[12] else
               if FType='A4' then Caption:=Caption+' '+MsgBox.Mess[13] else
                  if FType='A5' then Caption:=Caption+' '+MsgBox.Mess[14] else
                     if FType='BU' then Caption:=Caption+' '+MsgBox.Mess[15] ;
UpdateCaption(Self) ;
END ;

Procedure TFCorresp.Initialisation ;
BEGIN
LePlan:=Plan.TabIndex+1 ;
TCorresp.SetRange([FType+IntToStr(LePlan)],[FType+IntToStr(LePlan)]) ;
FAvertir:=False ;
if(TCorresp.EOF)And(TCorresp.BOF) then Bouge(nbInsert) ;
END ;

procedure TFCorresp.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TCorresp.Open ;
Initialisation ; FaitCaption ;
FListe.SetFocus ;
if ((EstSerie(S3)) or (EstSerie(S5))) then Plan.Tabs.Delete(1) ;
HMTrad.ResizeDBGridColumns(Fliste);
end;


procedure TFCorresp.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFCorresp.TCorrespNewRecord(DataSet: TDataset);
begin NewEnreg ; end;

procedure TFCorresp.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCorresp.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCorresp.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCorresp.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCorresp.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFCorresp.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFCorresp.BDefaireClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFCorresp.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFCorresp.NewEnreg ;
BEGIN
InitNew(TCorresp) ;
TCORRESPCR_TYPE.AsString:=FType+IntToStr(Plan.TabIndex+1) ;
FListe.Columns.Items[0].ReadOnly:=False ;
FListe.SelectedIndex:=0 ; FListe.SetFocus ;
END ;

Procedure TFCorresp.ChargeEnreg ;
BEGIN
// code ...
END ;

Function TFCorresp.VerifiSiExiste : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select CR_TYPE,CR_CORRESP From CORRESP Where CR_TYPE="'+TCORRESPCR_TYPE.AsString+'" And '+
              'CR_CORRESP="'+TCORRESPCR_CORRESP.AsString+'"',True) ;
Result:=(Not QLoc.Eof) ; Ferme(QLoc) ;
END ;

Function TFCorresp.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TCORRESP.state in [dsEdit,dsInsert]=False then Exit ;
if TCORRESP.state in [dsEdit,dsInsert] then
   BEGIN
   if TCORRESPCR_CORRESP.AsString='' then
      BEGIN MsgBox.Execute(2,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=0 ; Exit ; END ;
   if TCORRESPCR_LIBELLE.AsString='' then
      BEGIN MsgBox.Execute(3,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=1 ; Exit ; END ;
   END else exit ;
if TCORRESP.state in [dsInsert] then
   BEGIN
   if VerifiSiExiste then
      BEGIN MsgBox.Execute(4,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=0 ; Exit ; END ;
   END ;
Result:=TRUE  ;
END ;

Function TFCorresp.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TCORRESP.Modified then
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

Function TFCorresp.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.Mess[6]);
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

procedure TFCorresp.SCORRESPDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not((TCORRESP.State in [dsInsert,dsEdit])) ;
BDelete.Enabled:=Not((TCORRESP.State in [dsInsert,dsEdit])) ;
if(TCORRESP.Eof)And(TCORRESP.Bof) then BDelete.Enabled:=False ;
if TCORRESP.State=dsBrowse then Fliste.Columns.Items[0].ReadOnly:=True ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TCORRESP.BOF;
   DnEnable := Enabled and not TCORRESP.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
   if ((Field.FieldName='CR_LIBELLE') and (TCORRESPCR_ABREGE.AsString='')) then
      TCORRESPCR_ABREGE.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFCorresp.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

procedure TFCorresp.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil,Caption,'') ; end;

procedure TFCorresp.TCORRESPAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorresp.TCORRESPAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorresp.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
  BEGIN
  END ;
if Parent is THPanel then
   BEGIN
   _Bloqueur('nrBatch',False) ;
   Action:=caFree ;
   END ;
end;

procedure TFCorresp.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if(ssCtrl in Shift)And(Key=VK_DELETE)And(TCORRESP.EOF)And(TCORRESP.BOF) then
  BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)And(Key=VK_DELETE)then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFCorresp.TCORRESPPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
if TCORRESP.State=dsInsert then
   BEGIN MsgBox.Execute(4,'','') ; Fliste.SelectedIndex:=0 ; Action:=daAbort ; END ;
end;

procedure TFCorresp.PlanChange(Sender: TObject);
begin
if TCORRESP.State in [dsEdit,dsInsert] then BEGIN Plan.TabIndex:=LePlan-1 ; Exit ; END ;
Initialisation ;
end;

Function TFCorresp.Supprime : Boolean ;
Var zz,St : String ;
    i : Byte ;
BEGIN
Result:=False ;
if (MsgBox.execute(1,'','')<>mrYes) then Exit ;
Case MsgBox.execute(5,'','') of
   mrCancel : Exit ;
   mrNone  : BEGIN Result:=True ; Exit ; END ;
   END ;
St:=TCORRESPCR_CORRESP.AsString ; i:=0 ;
zz:=IntToStr(LePlan) ;
if FType='GE' then i:=1 else
if FType='AU' then i:=2 else
if FType='BU' then i:=3 else
if(FType='A1')Or(FType='A2')Or(FType='A3')Or(FType='A4')Or(FType='A5') then i:=4 ;
Case i of
     1:ExecuteSql('Update GENERAUX SET G_CORRESP'+zz+'="" Where G_CORRESP'+zz+'="'+St+'"') ;
     2:ExecuteSql('Update TIERS SET T_CORRESP'+zz+'="" Where T_CORRESP'+zz+'="'+St+'"') ;
//Simon     3:ExecuteSql('Update BUDGET SET B_CORRESP'+zz+'="" Where B_CORRESP'+zz+'="'+St+'"') ;
     4:ExecuteSql('Update SECTION SET S_CORRESP'+zz+'="" Where S_AXE="'+FType+'" AND S_CORRESP'+zz+'="'+St+'"') ;
     End ;
Result:=True ; Screen.Cursor:=SyncrDefault ;
END ;

procedure TFCorresp.SCORRESPUpdateData(Sender: TObject);
begin
if(Trim(TCORRESPCR_CORRESP.AsString)='') And (Trim(TCORRESPCR_LIBELLE.AsString)='') And
  (Trim(TCORRESPCR_ABREGE.AsString)='') then
  BEGIN TCORRESP.Cancel ; Exit ; END ;
if(Trim(TCORRESPCR_LIBELLE.AsString)<>'') And (Trim(TCORRESPCR_CORRESP.AsString)='') then
  BEGIN MsgBox.Execute(2,'','') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TCORRESPCR_CORRESP.AsString)<>'') And (Trim(TCORRESPCR_LIBELLE.AsString)='') then
  BEGIN MsgBox.Execute(3,'','') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TCORRESPCR_ABREGE.AsString)<>'') And (Trim(TCORRESPCR_CORRESP.AsString)='') then
  BEGIN MsgBox.Execute(2,'','') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
end ;

procedure TFCorresp.FListeRowEnter(Sender: TObject);
begin if TCORRESP.State=dsInsert then Fliste.SelectedIndex:=0 ; end;

procedure TFCorresp.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCorresp.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFCorresp.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
