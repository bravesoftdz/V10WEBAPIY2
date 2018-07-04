unit CORIMPSP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, StdCtrls, Ent1, Grids, DBGrids, DB, {$IFNDEF DBXPRESS}dbtables,
  HSysMenu, Hqry, HTB97, DBCtrls, ComCtrls, HDB, ExtCtrls{$ELSE}uDbxDataSet{$ENDIF}, Hctrls, ComCtrls, SaisComm,
  ExtCtrls, Buttons, DBCtrls, HCompte, HEnt1, HDB, HSysMenu, Hqry, HTB97, HPanel, UiUtil ;

Procedure ParamCorrespPlanAnal(Faxe : String) ;

type
  TFCorrespSP = class(TForm)
    TPlan: THTable;                    
    SPlan  : TDataSource;
    FListe1: THDBGrid;
    TPlanSS_AXE         : TStringField;
    TPlanSS_SOUSSECTION : TStringField;
    TPlanSS_LIBELLE     : TStringField;
    TPlanSS_CONTROLE    : TStringField;
    TPlanSS_DEBUT: TIntegerField;
    TPlanSS_LONGUEUR: TIntegerField;
    TSSPlan: THTable;
    SSSPlan : TDataSource;
    FListe2 : THDBGrid;
    TSSPlanPS_AXE         : TStringField;
    TSSPlanPS_SOUSSECTION : TStringField;
    TSSPlanPS_CODE        : TStringField;
    TSSPlanPS_LIBELLE     : TStringField;
    TSSPlanPS_ABREGE      : TStringField;
    PAxe     : TTabControl;
    Panel2   : TPanel;
    Panel3   : TPanel;
    HPB: TToolWindow97;
    Panel5   : TPanel;
    PPlan    : TPanel;
    BValider2: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide2: TToolbarButton97;
    FAutoSave : TCheckBox;
    MsgBox    : THMsgBox;
    MsgBox1   : THMsgBox;
    MsgBox2   : THMsgBox;
    TSSPlanPS_SOCIETE: TStringField;
    TPlanSS_SOCIETE: TStringField;
    Lgaxe: TLabel;
    HMTrad: THSystemMenu;
    DBNav1: TDBNavigator;
    DBNav2: TDBNavigator;
    Dock: TDock97;
    TSSPlanPS_CORRESP: TStringField;
    procedure FormShow(Sender: TObject);
    procedure PAxeChange(Sender: TObject);
    procedure BFerme1Click(Sender: TObject);
    procedure BValider2Click(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SSSPlanStateChange(Sender: TObject);
    procedure SSSPlanUpdateData(Sender: TObject);
    procedure SPlanDataChange(Sender: TObject; Field: TField);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TPlanAfterDelete(DataSet: TDataSet);
    procedure TPlanAfterPost(DataSet: TDataSet);
    procedure FListe1RowEnter(Sender: TObject);
    procedure FListe2RowEnter(Sender: TObject);
    procedure BAide2Click(Sender: TObject);
    procedure FListe1DblClick(Sender: TObject);
  private { Déclarations privées }
    FAxe : String ;
    Modifier1,Modifier2 : Boolean ;
    LgSect : Byte ;
    MemoLg : Byte ;
    FAvertir : Boolean ;
    OnlyOne : Boolean ;
    Function  Bouge2(Button: TNavigateBtn) : boolean ;
    Function  EnregOK2 : boolean ;
    Procedure LongueurSection ;
    Procedure CodeEditMask ;
    Function  SommeLongeur : Byte ;
    Function  OnSauve2 : boolean ;
  public  { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses Section, HStatus, GeSecAna ;

Procedure ParamCorrespPlanAnal(Faxe : String) ;
var FCorrespSP : TFCorrespSP;
    PP : THPanel ;
BEGIN
if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FCorrespSP:=TFCorrespSP.Create(Application) ;
FCorrespSP.Faxe:=Faxe ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FCorrespSP.ShowModal ;
    Finally
     FCorrespSP.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FCorrespSP,PP) ;
   FCorrespSP.Show ;
   END ;
END ;

procedure TFCorrespSP.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FAvertir:=False ;
TPlan.Open ;
if Faxe='' then FAxe:='A1' else PAxe.TabIndex:=StrToInt(Copy(FAxe,2,1))-1 ;
LongueurSection ;
TPlan.SetRange([FAxe],[FAxe]) ;
TSSPlan.Open ;
DelTabsSerie(PAxe) ;
end;

procedure TFCorrespSP.PAxeChange(Sender: TObject);
Var OldAxe : Integer ;
    NewAxe : String ;
    i : Byte ;
begin
OldAxe:=StrToInt(Copy(FAxe,2,1)) ; NewAxe:='A'+IntToStr(PAxe.TabIndex+1) ;
if (TPlan.State in [dsEdit,dsInsert])Or(TSSPlan.State in [dsEdit,dsInsert]) then
   BEGIN PAxe.TabIndex:=OldAxe-1 ; Exit ; END ;
if VH^.Cpta[AxeToFb(NewAxe)].Structure then
   BEGIN Faxe:=NewAxe ; TPlan.SetRange([FAxe],[FAxe]) ; END
else
   BEGIN MsgBox.Execute(0,'','') ; PAxe.TabIndex:=OldAxe-1 ; END ;
LongueurSection ;
end;

Procedure TFCorrespSP.LongueurSection ;
BEGIN
LgSect:=VH^.Cpta[AxeToFb(Faxe)].Lg ;
if LgSect<=0 then Lgaxe.Visible:=False
             else Lgaxe.Visible:=True ;
Case Faxe[2] of
   '1' : BEGIN Lgaxe.Left:=5   ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '2' : BEGIN Lgaxe.Left:=50  ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '3' : BEGIN Lgaxe.Left:=100 ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '4' : BEGIN Lgaxe.Left:=150 ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '5' : BEGIN Lgaxe.Left:=200 ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   End ;
END ;

procedure TFCorrespSP.BFerme1Click(Sender: TObject);
begin Close ; end;

procedure TFCorrespSP.BValider2Click(Sender: TObject);
begin Modifier2:=False ; Bouge2(nbPost) ; end;

procedure TFCorrespSP.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFCorrespSP.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
ChargeSousPlanAxe ;
end;

Function TFCorrespSP.Bouge2(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve2 then Exit ;
   nbPost           : if Not EnregOK2 then Exit ;
   nbDelete         : if MsgBox2.execute(1,'','')<>mrYes then Exit ;
   end ;
if not TransacNav(DBNav2.BtnClick,Button,10) then MessageAlerte(Msgbox.mess[1]) ;
result:=TRUE ;
END ;


Procedure TFCorrespSP.CodeEditMask ;
Var i : Byte ;
    St : String ;
BEGIN
St:='' ;
for i:=1 to TPlanSS_LONGUEUR.AsInteger do
   if Not OnlyOne then St:=St+'a' else St:=St+'c' ;
TSSPlanPS_CODE.EditMask:='>'+St+';0; ' ;
END ;

Function TFCorrespSP.EnregOK2 : boolean ;
BEGIN
result:=FALSE  ; Modifier2:=True ;
if TSSPlan.state in [dsinsert,dsedit]=False then Exit ;
if TSSPlan.state in [dsinsert,dsedit] then
   BEGIN
   if TSSPlanPS_CODE.asString='' then
      BEGIN
      MsgBox2.execute(2,'','') ; Fliste2.SelectedIndex:=0 ; FListe2.SetFocus ; Exit ;
      END ;
   if TSSPlanPS_LIBELLE.asString='' then
      BEGIN
      MsgBox2.execute(3,'','') ; Fliste2.SelectedIndex:=1 ; FListe2.SetFocus ; Exit ;
      END ;
   END ;
result:=TRUE  ; Modifier2:=False ;
END ;

Function TFCorrespSP.OnSauve2 : boolean ;
Var Rep : Integer ;
BEGIN
Result:=FALSE  ;
Modifier2:=True ;
if TSSPlan.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox2.execute(0,'','') ;
   END else rep:=321 ;
Case rep of
  mrYes : if not Bouge2(nbPost)   then Exit ;
  mrNo  : if not Bouge2(nbCancel) then Exit ;
  mrCancel : BEGIN Abort ; Exit ; END ;
  end ;
Modifier1:=False ;
result:=TRUE  ;
end ;

procedure TFCorrespSP.SSSPlanStateChange(Sender: TObject);
begin
Modifier2:=True ;
end;

procedure TFCorrespSP.SSSPlanUpdateData(Sender: TObject);
begin
if(Trim(TSSPlanPS_CODE.AsString)='') And (Trim(TSSPlanPS_LIBELLE.AsString)='') then
   BEGIN TSSPlan.Cancel ; Exit ; END ;
if(Trim(TSSPlanPS_CODE.AsString)<>'')And(Trim(TSSPlanPS_LIBELLE.AsString)='') then
  BEGIN
  MSgBox2.Execute(3,'','') ; FListe2.SelectedIndex:=1 ; Fliste2.SetFocus ; SysUtils.Abort ; Exit ;
  END ;
if(Trim(TSSPlanPS_LIBELLE.AsString)<>'')And(Trim(TSSPlanPS_CODE.AsString)='')then
  BEGIN
  MsgBox2.Execute(2,'','') ;  FListe2.SelectedIndex:=0 ; Fliste2.SetFocus ; SysUtils.Abort ; Exit ;
  END ;
end;


procedure TFCorrespSP.SPlanDataChange(Sender: TObject; Field: TField);
begin
if TPlan.State=dsBrowse then FListe1.Columns.Items[0].ReadOnly:=True ;
if Field=Nil then MemoLg:=TPlanSS_LONGUEUR.AsInteger ;
PPlan.Caption:=Msgbox.Mess[2]+Fliste1.Fields[1].AsString ;
Panel5.Caption:=Msgbox.Mess[4]+Fliste1.Fields[0].AsString ;
if TPlanSS_CONTROLE.AsString='X' then OnlyOne:=True else OnlyOne:=False ;
CodeEditMask ;
end;

Function TFCorrespSP.SommeLongeur : Byte ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select Sum(SS_LONGUEUR) from STRUCRSE where SS_AXE="'+Faxe+'"',True) ;
Result:=Q.Fields[0].AsInteger ; Ferme(Q) ;
END ;


procedure TFCorrespSP.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
   BEGIN
   ChargeStructureUnique ;
   AvertirTable('ttStrucrse1') ;
   AvertirTable('ttStrucrse2') ;
   AvertirTable('ttStrucrse3') ;
   AvertirTable('ttStrucrse4') ;
   AvertirTable('ttStrucrse5') ;
   END ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFCorrespSP.TPlanAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrespSP.TPlanAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrespSP.FListe1RowEnter(Sender: TObject);
begin
if TPlan.State=dsInsert then FListe1.SelectedIndex:=0 ;
end;

procedure TFCorrespSP.FListe2RowEnter(Sender: TObject);
begin
if TSSPlan.State=dsInsert then FListe2.SelectedIndex:=0 ;
end;

procedure TFCorrespSP.BAide2Click(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFCorrespSP.FListe1DblClick(Sender: TObject);
begin
if FListe1.SelectedIndex<>2 then Exit ;
end;


end.
