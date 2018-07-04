unit PlanEdt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, DB, DBTables, DBCtrls, StdCtrls, Buttons, Grids,
  DBGrids, HDB, Spin, Hctrls, Mask, ExtCtrls, HEnt1, Ent1, ComCtrls, Hqry,
  MajTable;

Procedure FicheEntetePlanEdt(Nature,Plan : String ; Action : TActionFiche) ;

type
  TFPlanEdt = class(TForm)
    Pappli: TPanel;
    FListe: THDBGrid;
    PBouton: TPanel;
    Binsert: TBitBtn;
    BImprimer: TBitBtn;
    FAutoSave: TCheckBox;
    BValider: TBitBtn;
    BAnnuler: TBitBtn;
    HelpBtn: TBitBtn;
    BFerme: TBitBtn;
    BDelete: TBitBtn;
    Panel1: TPanel;
    BFirst: TBitBtn;
    BPrev: TBitBtn;
    BNext: TBitBtn;
    BLast: TBitBtn;
    DBNav: TDBNavigator;
    Ta: THTable;
    TaPD_NATUREEDT: TStringField;
    TaPD_PLANEDT: TStringField;
    TaPD_LIBELLE: TStringField;
    TaPD_FAMILLES: TStringField;
    TaPD_UTILISATEUR: TStringField;
    TaPD_SOCIETE: TStringField;
    TaPD_NBDATEDEF: TIntegerField;
    TaPD_DATEDEF: TStringField;
    STa: TDataSource;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    HM2: THMsgBox;
    TaPD_BUDJAL: TStringField;
    TaPD_TOTALISATION: TStringField;
    TaPD_TYPEDETAIL: TStringField;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TPD_PLANEDT: TLabel;
    TPD_LIBELLE: TLabel;
    TPD_FAMILLES: TLabel;
    TPD_BUDJAL: TLabel;
    TPD_NATUREEDT: TLabel;
    PD_PLANEDT: TDBEdit;
    PD_LIBELLE: TDBEdit;
    PD_FAMILLES: THDBValComboBox;
    PD_BUDJAL: THDBValComboBox;
    PD_NATUREEDT: THDBValComboBox;
    GbDate: TGroupBox;
    TPD_NBDATEDEF: TLabel;
    TD1: TLabel;
    TD2: TLabel;
    TD3: TLabel;
    TD4: TLabel;
    TD5: TLabel;
    TD6: TLabel;
    TD7: TLabel;
    TD8: TLabel;
    TD9: TLabel;
    TD10: TLabel;
    TD11: TLabel;
    TD12: TLabel;
    PD_NBDATEDEF: THDBSpinEdit;
    D1: TEdit;
    D2: TEdit;
    D3: TEdit;
    D4: TEdit;
    D5: TEdit;
    D6: TEdit;
    D7: TEdit;
    D8: TEdit;
    D9: TEdit;
    D10: TEdit;
    D11: TEdit;
    D12: TEdit;
    PDateDef: TPanel;
    TPD_DATEDEF: TLabel;
    PD_DATEDEF: TDBEdit;
    GbCompo: TGroupBox;
    PD_TOTALISATION: TDBCheckBox;
    CbDet1: TCheckBox;
    CbDet2: TCheckBox;
    procedure BAnnulerClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure STaDataChange(Sender: TObject; Field: TField);
    procedure STaStateChange(Sender: TObject);
    procedure STaUpdateData(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure PD_NATUREEDTChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TaBeforeDelete(DataSet: TDataSet);
    procedure PD_NBDATEDEFChange(Sender: TObject);
    procedure D1DblClick(Sender: TObject);
    procedure PD_FAMILLESChange(Sender: TObject);
  private
    Nature,Plan : String ;
    Action : TActionFiche ;
    Modifier : Boolean ;
    FNomPrint : String ;
    MemoNat,MemoPlan : String ;
    LigneExiste : Boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Procedure ChargeEnreg ;
    Procedure NewEnreg ;
    Procedure ChercheLigneAssocie ;
    Function  OnSupprime : Boolean ;
    Procedure FaitCaption ;
    Procedure LitDateDefaut ;
    Procedure InitDateDefaut(St : String) ;
    Function  ChangeChampDateDef(LaVal : String) : String ;
    Function  ExisteCle : Boolean ;
    Procedure LitDetail ;
 public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses PrintDBG, DateDef ;

Procedure FicheEntetePlanEdt(Nature,Plan : String ; Action : TActionFiche) ;
var FPlanEdt : TFPlanEdt ;
BEGIN
FPlanEdt:=TFPlanEdt.Create(Application) ;
 Try
  FPlanEdt.Action:=Action ;
  FPlanEdt.Nature:=Nature ;
  FPlanEdt.Plan:=Plan ;
//  FPlanEdt.FNomPrint:='PRT_PLANEDT' ; Dés que Liste.exe est opérationnel créer la liste
  FPlanEdt.ShowModal ;
 Finally
  FPlanEdt.Free ;
 End ;
SourisNormale ;
END ;

procedure TFPlanEdt.FormCreate(Sender: TObject);
begin MemoNat:='' ; MemoPlan:='' ; end;

procedure TFPlanEdt.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFPlanEdt.BinsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFPlanEdt.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFPlanEdt.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFPlanEdt.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFPlanEdt.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFPlanEdt.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFPlanEdt.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFPlanEdt.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFPlanEdt.BImprimerClick(Sender: TObject);
var MyBookmark: TBookmark;
begin
if FNomPrint='' then exit ;
MyBookmark :=Ta.GetBookmark ;
PrintDBGrid (Nil,Nil,Copy(Caption,1,Pos(':',Caption)-1),FNomPrint) ;
Ta.GotoBookmark(MyBookmark) ;
Ta.FreeBookmark(MyBookmark);
end;

procedure TFPlanEdt.HelpBtnClick(Sender: TObject);
begin CallHelpTopic(Self) ; end ;

Function TFPlanEdt.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : if Not EnregOK then Exit ;
   nbDelete         : if Not OnSupprime then Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(HM.Mess[5]);
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFPlanEdt.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier:=False ;
if Ta.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=HM.execute(0,Caption,'') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : BEGIN Modifier:=True ; Exit ; END ;
  end ;
result:=TRUE  ;
end ;

Function TFPlanEdt.EnregOK : boolean ;
BEGIN
result:=FALSE  ; Modifier:=True ;
if Ta.state in [dsEdit,dsInsert]=False then Exit ;
if Ta.state in [dsEdit,dsInsert] then
   BEGIN
   if TaPD_PLANEDT.AsString=''then
      BEGIN HM.Execute(2,Caption,'') ; if PD_PLANEDT.CanFocus then PD_PLANEDT.SetFocus ; Exit ; END ;
   if TaPD_LIBELLE.AsString='' then
      BEGIN HM.Execute(3,Caption,'') ; if PD_LIBELLE.CanFocus then PD_LIBELLE.SetFocus ; Exit ; END ;
   if TaPD_FAMILLES.AsString='' then
      BEGIN HM.Execute(6,Caption,'') ; if PD_FAMILLES.CanFocus then PD_FAMILLES.SetFocus ; Exit ; END ;
   if (TaPD_NATUREEDT.AsString='BUD') And (TaPD_BUDJAL.AsString='') then
      BEGIN HM.Execute(8,Caption,'') ; if PD_BUDJAL.CanFocus then PD_BUDJAL.SetFocus ; Exit ; END ;
   END ;
if Ta.state in [dsInsert] then
   BEGIN
   if ExisteCle then
      BEGIN HM.Execute(4,Caption,'') ; if PD_PLANEDT.CanFocus then PD_PLANEDT.SetFocus ; Exit ; END ;
   END ;
DateModification(Ta,'PD') ;
result:=TRUE  ; Modifier:=False ;
END ;

Procedure TFPlanEdt.NewEnreg ;
BEGIN
InitNew(Ta) ;
PD_PLANEDT.Enabled:=True ; PD_NATUREEDT.Enabled:=True ; PD_FAMILLES.Enabled:=True ;
TaPD_NATUREEDT.AsString:=PD_NATUREEDT.Values[0] ;
PD_NATUREEDTChange(Nil) ;
TaPD_FAMILLES.AsString:=PD_FAMILLES.Values[0] ;
PD_NBDATEDEF.Value:=0 ; InitDateDefaut('') ;
LigneExiste:=False ; FaitCaption ;
PD_PLANEDT.SetFocus ;
END ;

Procedure TFPlanEdt.ChercheLigneAssocie ;
BEGIN
LigneExiste:=PresenceComplexe('LIGNEEDT',['LD_NATLIGNEEDT','LD_PLANEDT'],['=','='],[TaPD_NATUREEDT.AsString,TaPD_PLANEDT.AsString],['S','S']) ;
END ;

Procedure TFPlanEdt.FaitCaption ;
BEGIN
if TaPD_NATUREEDT.AsString='CPT' then Caption:=HM2.Mess[0]
                                 else Caption:=HM2.Mess[1] ;
InitCaption(Self,TaPD_PLANEDT.AsString,TaPD_LIBELLE.AsString) ;
END ;

Procedure TFPlanEdt.ChargeEnreg ;
BEGIN
FaitCaption ;
if Action=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
if (Ta.State<>dsInsert) then
   BEGIN
   if (MemoNat<>TaPD_NATUREEDT.AsString) or (MemoPlan<>TaPD_PLANEDT.AsString) then ChercheLigneAssocie ;
   LitDateDefaut ;
   END ;
MemoPlan:=TaPD_PLANEDT.AsString ;
if MemoNat<>TaPD_NATUREEDT.AsString then PD_NATUREEDTChange(Nil) ;
MemoNat:=TaPD_NATUREEDT.AsString ;
PD_PLANEDT.Enabled:=False ; PD_NATUREEDT.Enabled:=False ;
if Action=taModif then PD_FAMILLES.Enabled:=Not LigneExiste ;
END ;

procedure TFPlanEdt.STaDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
Binsert.Enabled:=(Not(Ta.State in[dsEdit,dsInsert])) ;
BDelete.Enabled:=(Not(Ta.State in[dsEdit,dsInsert])) ;
if(Ta.Eof)And(Ta.Bof) then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not Ta.BOF;
   DnEnable := Enabled and not Ta.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
   END ;
end;

procedure TFPlanEdt.STaStateChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFPlanEdt.STaUpdateData(Sender: TObject);
begin
if Not Ta.Modified then Exit ;
if Modifier then BEGIN Modifier:=False ; if Not OnSauve then if Ta.Modified then SysUtils.Abort ; END ;
end;

procedure TFPlanEdt.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin if (ssCtrl in Shift) AND (Key=VK_DELETE) then BEGIN Key:=0 ; Exit ; END ; end;

procedure TFPlanEdt.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFPlanEdt.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Ta.Open ;
if (Nature<>'') And (Plan<>'') then
   if Not FindLakey(Ta,[Nature,Plan]) then BEGIN MessageAlerte(HM.Mess[5]) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; END ;
if (Nature<>'') And (Plan='') then Ta.SetRange([Nature],[Nature]) else
   if (Nature='') And (Plan<>'') then Ta.SetRange(['AAA',Plan],['zzz',Plan]) ;
Case Action of
     taConsult : BEGIN FicheReadOnly(Self) ; Exit ; END ;
     taCreat..taCreatOne : BEGIN
                           BinsertClick(Nil) ; PD_PLANEDT.Text:=Plan ;
                           PD_NATUREEDT.Value:=Nature ; BAnnuler.Enabled:=False ;
                           END ;
   End ;
if(Ta.Eof) And (Ta.Bof) And (Action<>taConsult)then BinsertClick(Nil) ;
end;

procedure TFPlanEdt.PD_NATUREEDTChange(Sender: TObject);
Var QLoc : TQuery ;
    St : String ;
begin
if PD_NATUREEDT.Value='CPT' then
   BEGIN
   PD_FAMILLES.DataType:='ttRubFamille' ;
   TPD_BUDJAL.Enabled:=False ; PD_BUDJAL.Enabled:=False ;
   if Ta.State<>dsBrowse then TaPD_BUDJAL.AsString:='' ;
   END else
   BEGIN
   TPD_BUDJAL.Enabled:=True ; PD_BUDJAL.Enabled:=True ;
   if (PD_BUDJAL.Values.Count>0) And (TaPD_BUDJAL.AsString='') And (Ta.State<>dsBrowse) then
      TaPD_BUDJAL.AsString:=PD_BUDJAL.Values[0] ;
   PD_FAMILLES.DataType:='' ;
   PD_FAMILLES.Items.Clear ; PD_FAMILLES.Values.Clear ;
   if Ta.State<>dsBrowse then TaPD_FAMILLES.AsString:='' ;
   QLoc:=OpenSql('Select CO_CODE,CO_LIBELLE From COMMUN Where CO_TYPE="RBB" Order by CO_CODE ',True) ;
   While Not QLoc.Eof do
       BEGIN
       St:=QLoc.Fields[0].AsString ;
       if St='ANA' then St:='CBS' else
          if St='A/G' then St:='S/G' else
             if St='GEN' then St:='CBG' else
                 if St='G/A' then St:='G/S' ;
       PD_FAMILLES.Values.Add(St) ; PD_FAMILLES.Items.Add(QLoc.Fields[1].AsString) ;
       QLoc.Next ;
       END ;
   if (Ta.State=dsBrowse) and (TaPD_FAMILLES.AsString<>'') then PD_FAMILLES.ItemIndex:=PD_FAMILLES.Values.IndexOf(TaPD_FAMILLES.AsString) ;
   END ;
end;

Function TFPlanEdt.OnSupprime : Boolean ;
BEGIN
Result:=False ;
if HM.Execute(1,Caption,'')<>mrYes then Exit ;
if LigneExiste then
   if HM.Execute(7,Caption,'')<>mrYes then Exit ;
Result:=True ;
END ;

procedure TFPlanEdt.TaBeforeDelete(DataSet: TDataSet);
begin
if LigneExiste then
   ExecuteSql('Delete From LIGNEEDT Where LD_NATLIGNEEDT="'+TaPD_NATUREEDT.AsString+'" And '+
              'LD_PLANEDT="'+TaPD_PLANEDT.AsString+'"') ;
end;

Function TFPlanEdt.ChangeChampDateDef(LaVal : String) : String ;
Var i,j : Integer ;
    St,St1 : String ;
BEGIN
St:=TaPD_DATEDEF.AsString ;
j:=0 ; St1:='' ;
for i:=1 to Length(St) do
    BEGIN
    if j=PD_NBDATEDEF.Value then Break ;
    St1:=St1+St[i] ;
    if (St[i]=';') then Inc(j) ;
    END ;
if j<PD_NBDATEDEF.Value then St1:=St1+LaVal ;
if (St1<>'') or (j=PD_NBDATEDEF.Value) then TaPD_DATEDEF.AsString:=St1 ;
Result:=St1 ;
END ;

procedure TFPlanEdt.PD_NBDATEDEFChange(Sender: TObject);
Var St : String ;
begin
if Ta.State=dsBrowse then Exit ;
St:=ChangeChampDateDef(HM2.Mess[2]+';') ;
InitDateDefaut(St) ;
end;

Procedure TFPlanEdt.InitDateDefaut(St : String) ;
Var i : Integer ;
BEGIN
for i:=1 to 12 do
   BEGIN
   TEdit(FindComponent('D'+IntToStr(i))).Text:=ReadTokenSt(St) ;
   TEdit(FindComponent('D'+IntToStr(i))).Enabled:=(i<=PD_NBDATEDEF.Value) ;
   TLabel(FindComponent('TD'+IntToStr(i))).Enabled:=(i<=PD_NBDATEDEF.Value) ;
   END ;
END ;

Procedure TFPlanEdt.LitDateDefaut ;
Var St : String ;
BEGIN St:=TaPD_DATEDEF.AsString ; InitDateDefaut(St) ; END ;

procedure TFPlanEdt.D1DblClick(Sender: TObject);
Var St,St1,St2,StOk : String ;
    i,LaPos : Integer ;
begin
St:=TEdit(Sender).Text ;
i:=StrToInt(Copy(TEdit(Sender).Name,2,2)) ;
ParamDateDefPlEdt(St) ;
if St=TEdit(Sender).Text then Exit ;
if Ta.State=dsBrowse then Ta.Edit ;
TEdit(Sender).Text:=St ;
St1:=TaPD_DATEDEF.AsString ; LaPos:=0 ; StOk:='' ;
While St1<>'' do
    BEGIN
    St2:=ReadTokenSt(St1) ; Inc(LaPos) ;
    if i=LaPos then Break else StOk:=StOk+St2+';' ;
    END ;
StOk:=StOk+St+';'+St1 ;
TaPD_DATEDEF.AsString:=StOk ;
end;

Function TFPlanEdt.ExisteCle : Boolean ;
BEGIN
Result:=PresenceComplexe('PLANEDT',['PD_PLANEDT','PD_NATUREEDT'],['=','='],[TaPD_PLANEDT.AsString,TaPD_NATUREEDT.AsString],['S','S']) ;
END ;

procedure TFPlanEdt.PD_FAMILLESChange(Sender: TObject);
begin
GbCompo.Enabled:=Pos('/',PD_FAMILLES.Value)>0 ;
end;

Procedure TFPlanEdt.LitDetail ;
BEGIN
if TaPD_TYPEDETAIL.AsString='' then
   BEGIN
   CbDet1.State:=cbUnchecked ; CbDet2.State:=cbUnchecked ;
   END else
   BEGIN
   CbDet1.Checked:=(Copy(TaPD_TYPEDETAIL.AsString,4,1)='X') ;
   CbDet2.Checked:=(Copy(TaPD_TYPEDETAIL.AsString,5,1)='X') ;
   END ;
END ;


end.
