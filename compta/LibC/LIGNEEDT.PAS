unit LigneEdt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, DB, {$IFNDEF DBXPRESS}dbtables, Hqry, StdCtrls,
  Hcompte, HRegCpte, DBCtrls, Spin, HDB, Hctrls, Mask, Grids, DBGrids,
  Buttons, ExtCtrls{$ELSE}uDbxDataSet{$ENDIF}, DBCtrls, Grids, DBGrids, HDB, StdCtrls,
  Buttons, ExtCtrls, HCtrls, HEnt1, Ent1, Mask, Spin, Hcompte, Hqry,
  HRegCpte;

Procedure ParamLigneEdt(Nature,Plan,CodLig : String ; Action : TActionFiche) ;

Type TInfoPlanEdt = Class
         Nat : String ;
         Fam : String ;
         Jal : String ;
      End ;

type
  TFLigneEdt = class(TForm)
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
    FListe: THDBGrid;
    DBNav: TDBNavigator;
    Ta: THTable;
    STa: TDataSource;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    TaLD_NATLIGNEEDT: TStringField;
    TaLD_PLANEDT: TStringField;
    TaLD_LIGNEEDT: TStringField;
    TaLD_CHRONO: TIntegerField;
    TaLD_LIBELLE: TStringField;
    TaLD_RUBRIQUE: TStringField;
    TaLD_TYPEDETAIL: TStringField;
    TaLD_TOTALISATION: TStringField;
    TaLD_POURBUD: TStringField;
    TaLD_FORMAT: TStringField;
    TLD_LIGNEEDT: TLabel;
    LD_LIGNEEDT: TDBEdit;
    TLD_LIBELLE: TLabel;
    LD_LIBELLE: TDBEdit;
    TLD_PLANEDT: TLabel;
    LD_PLANEDT: THDBValComboBox;
    TLD_CHRONO: TLabel;
    LD_CHRONO: THDBSpinEdit;
    TLD_RUBRIQUE: TLabel;
    TLD_POURBUD: TLabel;
    LD_POURBUD: THDBValComboBox;
    GbCompo: TGroupBox;
    LD_TOTALISATION: TDBCheckBox;
    CbDet1: TCheckBox;
    CbDet2: TCheckBox;
    PNat: TPanel;
    TLD_NATLIGNEEDT: TLabel;
    LD_NATLIGNEEDT: TDBEdit;
    LD_RUBRIQUE: THDBCpteEdit;
    CbDet3: TCheckBox;
    procedure BAnnulerClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure STaDataChange(Sender: TObject; Field: TField);
    procedure STaStateChange(Sender: TObject);
    procedure STaUpdateData(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure LD_PLANEDTChange(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
  private
    Nature,Plan,CodLig : String ;
    Action : TActionFiche ;
    Modifier : Boolean ;
    MemoPlan : String ;
    QuelPlan,QuelNat : String ;
    FNomPrint : String ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Procedure ChargeEnreg ;
    Procedure NewEnreg ;
    Function  OnSupprime : Boolean ;
    Procedure ChargeCbPlan ;
    Procedure VideCbPlan ;
    Procedure ChargeCbRubrique ;
    Procedure MemoChampParDefaut ;
    Function  ChampsParDefautPourCreatenSerie : Boolean ;
    Procedure ClearZones ;
    Function  PourBudOk : Boolean ;
    Procedure LitDetail ;
    Procedure EcritDetail ;
    Function  ExisteCle : Boolean ;
    Function  ChronoPasOk : Boolean ;
    Procedure ChargeDefauts ;
    Procedure PositionnePlan ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses PrintDBG ;

Procedure ParamLigneEdt(Nature,Plan,CodLig : String ; Action : TActionFiche) ;
var FLigneEdt : TFLigneEdt ;
BEGIN
FLigneEdt:=TFLigneEdt.Create(Application) ;
 Try
//  FLigneEdt.FNomPrint:='PRT_LIGNEEDT' ; Dés que Liste.exe est opérationnel créer la liste
  FLigneEdt.Nature:=Nature ;
  FLigneEdt.Plan:=Plan ;
  FLigneEdt.CodLig:=CodLig ;
  FLigneEdt.Action:=Action ;
  FLigneEdt.ShowModal ;
 Finally
  FLigneEdt.Free ;
 End ;
SourisNormale ;
END ;

procedure TFLigneEdt.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFLigneEdt.BinsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFLigneEdt.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFLigneEdt.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFLigneEdt.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFLigneEdt.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFLigneEdt.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFLigneEdt.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFLigneEdt.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

Function TFLigneEdt.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=False  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : if Not EnregOK then Exit ;
   nbDelete         : if Not OnSupprime then Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(HM.Mess[5]);
result:=True ;
if Button=NbInsert then NewEnreg ;
if Button=nbCancel then ChargeCbRubrique ;
END ;

Function TFLigneEdt.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=False ; Modifier:=False ;
if Ta.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=HM.execute(0,Caption,'') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : BEGIN Modifier:=True ; Exit ; END ;
  end ;
result:=True  ;
end ;

Procedure TFLigneEdt.ClearZones ;
BEGIN
if LD_PLANEDT.Value='' then Exit ;
if TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Nat<>'BUD' then
   BEGIN TaLD_TOTALISATION.AsString:='-' ; TaLD_POURBUD.AsString:='' ; END ;
END ;

Function TFLigneEdt.PourBudOk : Boolean ;
BEGIN
Result:=True ;
if TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Nat<>'BUD' then Exit ;
if TaLD_POURBUD.AsString='' then Result:=False ;
END ;

Function TFLigneEdt.ChronoPasOk : Boolean ;
BEGIN
Result:=PresenceComplexe('LIGNEEDT',['LD_CHRONO','LD_LIGNEEDT','LD_PLANEDT','LD_NATLIGNEEDT'],['=','<>','=','='],[TaLD_CHRONO.AsString,TaLD_LIGNEEDT.AsString,TaLD_PLANEDT.AsString,TaLD_NATLIGNEEDT.AsString],['I','S','S','S']) ;
END ;

Function TFLigneEdt.EnregOK : boolean ;
BEGIN
result:=FALSE  ; Modifier:=True ;
if Ta.state in [dsEdit,dsInsert]=False then Exit ;
if Ta.state in [dsEdit,dsInsert] then
   BEGIN
   if TaLD_LIGNEEDT.AsString=''then
      BEGIN HM.Execute(2,Caption,'') ; if LD_LIGNEEDT.CanFocus then LD_LIGNEEDT.SetFocus ; Exit ; END ;
   if TaLD_LIBELLE.AsString='' then
      BEGIN HM.Execute(3,Caption,'') ; if LD_LIBELLE.CanFocus then LD_LIBELLE.SetFocus ; Exit ; END ;
   if TaLD_RUBRIQUE.AsString='' then
      BEGIN HM.Execute(6,Caption,'') ; if LD_RUBRIQUE.CanFocus then LD_RUBRIQUE.SetFocus ; Exit ; END ;
   if Not PourBudOk then
      BEGIN HM.Execute(7,Caption,'') ; if LD_POURBUD.CanFocus then LD_POURBUD.SetFocus ; Exit ; END ;
   if ChronoPasOk then
      BEGIN HM.Execute(8,Caption,'') ; if LD_CHRONO.CanFocus then LD_CHRONO.SetFocus ; Exit ; END ;
   END ;
if Ta.state in [dsInsert] then
   BEGIN
   if ExisteCle then
      BEGIN HM.Execute(4,Caption,'') ; if LD_LIGNEEDT.CanFocus then LD_LIGNEEDT.SetFocus ; Exit ; END ;
   END ;
DateModification(Ta,'LD') ;
result:=TRUE  ; Modifier:=False ; MemoChampParDefaut ; ClearZones ; EcritDetail ;
END ;

Procedure TFLigneEdt.PositionnePlan ;
Var i : Integer ;
BEGIN
for i:=0 to LD_PLANEDT.Values.Count-1 do
   if (TInfoPlanEdt(LD_PLANEDT.Values.Objects[i]).Nat=TaLD_NATLIGNEEDT.AsString) And
      (LD_PLANEDT.Values[i]=TaLD_PLANEDT.AsString) then BEGIN LD_PLANEDT.ItemIndex:=i ; Break ; END ;
END ;

Procedure TFLigneEdt.ChargeEnreg ;
BEGIN
InitCaption(Self,TaLD_LIGNEEDT.AsString,TaLD_LIBELLE.AsString) ;
LD_LIGNEEDT.Enabled:=False ; LitDetail ;
PositionnePlan ;
if (QuelPlan<>TaLD_PLANEDT.AsString) or (QuelNat<>TaLD_NATLIGNEEDT.AsString) then
   BEGIN
   QuelPlan:=TaLD_PLANEDT.AsString ;
   QuelNat:=TaLD_NATLIGNEEDT.AsString ;
   if Ta.State<>dsInsert then LD_PLANEDTChange(Nil) else ChargeCbRubrique ;
   END ;
//if Action=taConsult then FicheReadOnly(Self) ;
END ;

Procedure TFLigneEdt.NewEnreg ;
BEGIN
InitNew(Ta) ;
if Not ChampsParDefautPourCreatenSerie then
   if LD_PLANEDT.Values.Count>0 then
      BEGIN
      TaLD_PLANEDT.AsString:=LD_PLANEDT.Values[0] ; LD_PLANEDTChange(Nil) ;
      END ;
ChargeDefauts ;
LD_LIGNEEDT.Enabled:=True ; LD_LIGNEEDT.SetFocus ;
END ;

Function TFLigneEdt.OnSupprime : Boolean ;
BEGIN
Result:=False ;
if HM.Execute(1,Caption,'')<>mrYes then Exit ;
Result:=True ;
END ;

procedure TFLigneEdt.STaDataChange(Sender: TObject; Field: TField);
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

procedure TFLigneEdt.STaStateChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFLigneEdt.STaUpdateData(Sender: TObject);
begin
if Not Ta.Modified then Exit ;
if Modifier then BEGIN Modifier:=False ; if Not OnSauve then if Ta.Modified then SysUtils.Abort ; END ;
end;

procedure TFLigneEdt.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin if (ssCtrl in Shift) AND (Key=VK_DELETE) then BEGIN Key:=0 ; Exit ; END ; end;

procedure TFLigneEdt.FormClose(Sender: TObject; var Action: TCloseAction);
begin VideCbPlan ; end;

procedure TFLigneEdt.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFLigneEdt.FormShow(Sender: TObject);
begin
MemoPlan:='' ; QuelPlan:='' ; QuelNat:='' ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; ChargeCBPlan ;
Ta.Open ;
if (Nature<>'') And (Plan<>'') And (CodLig<>'') then
   if Not FindLakey(Ta,[Nature,Plan,CodLig]) then BEGIN MessageAlerte(HM.Mess[5]) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; END ;
Case Action of
//     taConsult : BEGIN FicheReadOnly(Self) ; Exit ; END ;
     taCreat..taCreatOne : BinsertClick(Nil) ;
   End ;
if(Ta.Eof) And (Ta.Bof) And (Action<>taConsult)then BinsertClick(Nil) ;
end;

Procedure TFLigneEdt.ChargeCbPlan ;
Var QLoc : TQuery ;
    X : TInfoPlanEdt ;
BEGIN
LD_PLANEDT.Values.Clear ; LD_PLANEDT.Items.Clear ;
QLoc:=OpenSql('Select * From PLANEDT',True) ;
While Not QLoc.Eof do
   BEGIN
   X:=TInfoPlanEdt.Create ;
   X.Nat:=QLoc.FindField('PD_NATUREEDT').AsString ;
   X.Fam:=QLoc.FindField('PD_FAMILLES').AsString ;
   X.Jal:=QLoc.FindField('PD_BUDJAL').AsString ;
   LD_PLANEDT.Values.AddObject(QLoc.FindField('PD_PLANEDT').AsString,X) ;
   LD_PLANEDT.Items.Add(QLoc.FindField('PD_LIBELLE').AsString) ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

Procedure TFLigneEdt.VideCbPlan ;
Var i : Integer ;
BEGIN
for i:=0 to LD_PLANEDT.Values.Count-1 do
   if LD_PLANEDT.Values.Objects[i]<>Nil then
      BEGIN TObject(LD_PLANEDT.Values.Objects[i]).Free ; LD_PLANEDT.Values.Objects[i]:=Nil ; END ;
END ;

procedure TFLigneEdt.LD_PLANEDTChange(Sender: TObject);
begin
if LD_PLANEDT.Value='' then Exit ;
ChargeCbRubrique ;
if Ta.State<>dsBrowse then TaLD_NATLIGNEEDT.AsString:=TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Nat ;
GbCompo.Enabled:=Pos('/',TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Fam)>0 ;
TLD_POURBUD.Enabled:=TaLD_NATLIGNEEDT.AsString='BUD' ;
LD_POURBUD.Enabled:=TaLD_NATLIGNEEDT.AsString='BUD' ;
ChargeDefauts ;
end;

Procedure TFLigneEdt.ChargeCbRubrique ;
Var LaFam,LaNat,LeJal : String ;
BEGIN
if LD_PLANEDT.Value='' then Exit ;
if Ta.State<>dsBrowse then TaLD_RUBRIQUE.AsString:='' ;
LaNat:=TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Nat ;
LaFam:=TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Fam+';' ;
LeJal:=TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Jal ;
if LaNat<>'BUD' then
   BEGIN
   LD_RUBRIQUE.ZoomTable:=tzRubCPTA ; LD_RUBRIQUE.SynPlus:=LaFam ;
   END else
   BEGIN
   if LaFam='CBG;' then LD_RUBRIQUE.ZoomTable:=tzRubBUDG else
   if LaFam='CBS;' then LD_RUBRIQUE.ZoomTable:=tzRubBUDS else
   if LaFam='G/S;' then LD_RUBRIQUE.ZoomTable:=tzRubBUDGS else
   if LaFam='S/G;' then LD_RUBRIQUE.ZoomTable:=tzRubBUDSG ;
   LD_RUBRIQUE.SynPlus:=LeJal ;
   END ;
if Ta.State<>dsInsert then LD_RUBRIQUE.Text:=TaLD_RUBRIQUE.AsString ;
END ;

Procedure TFLigneEdt.MemoChampParDefaut ;
BEGIN
if Action<>taCreatEnSerie then Exit ;
MemoPlan:=TaLD_PLANEDT.AsString ;
END ;

Function TFLigneEdt.ChampsParDefautPourCreatenSerie : Boolean ;
BEGIN
Result:=False ;
if Action<>taCreatEnSerie then Exit ;
if MemoPlan='' then Exit ;
TaLD_PLANEDT.AsString:=MemoPlan ; LD_PLANEDT.Value:=MemoPlan ;
Result:=True ;
END ;

Procedure TFLigneEdt.LitDetail ;
BEGIN
if TaLD_TYPEDETAIL.AsString='' then
   BEGIN
   CbDet1.State:=cbUnchecked ; CbDet2.State:=cbUnchecked ; CbDet3.State:=cbUnchecked ;
   END else
   BEGIN
   CbDet1.Checked:=(Copy(TaLD_TYPEDETAIL.AsString,4,1)='X') ;
   CbDet2.Checked:=(Copy(TaLD_TYPEDETAIL.AsString,5,1)='X') ;
   CbDet3.Checked:=(Copy(TaLD_TYPEDETAIL.AsString,6,1)='X') ;
   END ;
END ;

Procedure TFLigneEdt.EcritDetail ;
Var i : Integer ;
    St : String ;
BEGIN
i:=Pos('/',TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Fam) ;
if i<=0 then TaLD_TYPEDETAIL.AsString:='' else
   BEGIN
   St:=TInfoPlanEdt(LD_PLANEDT.Values.Objects[LD_PLANEDT.ItemIndex]).Fam ;
   if CbDet1.Checked then St:=St+'X' else St:=St+'-' ;
   if CbDet2.Checked then St:=St+'X' else St:=St+'-' ;
   if CbDet3.Checked then St:=St+'X' else St:=St+'-' ;
   TaLD_TYPEDETAIL.AsString:=St ;
   END ;
END ;

Function TFLigneEdt.ExisteCle : Boolean ;
BEGIN
Result:=PresenceComplexe('LIGNEEDT',['LD_LIGNEEDT','LD_PLANEDT','LD_NATLIGNEEDT'],['=','=','='],[TaLD_LIGNEEDT.AsString,TaLD_PLANEDT.AsString,TaLD_NATLIGNEEDT.AsString],['S','S','S']) ;
END ;

procedure TFLigneEdt.HelpBtnClick(Sender: TObject);
begin CallHelpTopic(Self) ; end ;

procedure TFLigneEdt.BImprimerClick(Sender: TObject);
var MyBookmark: TBookmark;
begin
if FNomPrint='' then exit ;
MyBookmark :=Ta.GetBookmark ;
PrintDBGrid (Nil,Nil,Copy(Caption,1,Pos(':',Caption)-1),FNomPrint) ;
Ta.GotoBookmark(MyBookmark) ;
Ta.FreeBookmark(MyBookmark);
end;

Procedure TFLigneEdt.ChargeDefauts ;
Var QLoc : TQuery ;
    i : Integer ;
BEGIN
if Ta.State<>dsInsert then Exit ;
QLoc:=OpenSql('Select LD_CHRONO From LIGNEEDT Where LD_PLANEDT="'+LD_PLANEDT.Value+'" '+
              ' And LD_NATLIGNEEDT="'+LD_NATLIGNEEDT.Text+'" Order by LD_CHRONO',True) ;
i:=1 ;
While Not QLoc.Eof do
   BEGIN
   if i<QLoc.Fields[0].AsInteger then Break else Inc(i) ;
   QLoc.Next ;
   END ;
TaLD_CHRONO.AsInteger:=i ;
Ferme(QLoc) ;
QLoc:=OpenSql('Select PD_TOTALISATION, PD_TYPEDETAIL From PLANEDT Where PD_PLANEDT="'+LD_PLANEDT.Value+'" '+
              ' And PD_NATUREEDT="'+LD_NATLIGNEEDT.Text+'"',True) ;
if not QLoc.Eof then
  BEGIN
  LD_TOTALISATION.Checked:=(QLoc.Fields[0].AsString='X') ;
  CbDet1.Checked:=(Copy(QLoc.Fields[1].AsString,4,1)='X') ;
  CbDet2.Checked:=(Copy(QLoc.Fields[1].AsString,5,1)='X') ;
  CbDet3.Checked:=(Copy(QLoc.Fields[1].AsString,6,1)='X') ;
  END ;
Ferme(QLoc) ;
END ;

end.
