unit  SISCOCOR ;

interface

uses SysUtils, WinProcs, Classes, Forms, Controls,
     StdCtrls, Hctrls, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} DB, DBGrids, DBCtrls,
     Ent1, HmsgBox, Hqry, HEnt1, HDB, Messages, HSysMenu, HPanel, UiUtil,
     HTB97,Hcompte,HStatus, ExtCtrls, Grids, ADODB, TntDBGrids ;

procedure ParamAuxSISCO ;
procedure ParamAuxRECUPSISCO(StFichier : String) ;
procedure MajSISCOCOR(StFichier : String) ;

type
  TFCorrSISCO = class(TForm)
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
    TCORRESPCR_TYPE   : TStringField;
    TCORRESPCR_CORRESP: TStringField;
    TCORRESPCR_LIBELLE: TStringField;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock: TDock97;
    Cache: THCpteEdit;
    TCORRESPCR_ABREGE: TStringField;
    BRecupSISCO: TToolbarButton97;
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
    procedure SCORRESPUpdateData(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BRecupSISCOClick(Sender: TObject);
  private
    FType : string ;
    FAvertir : boolean ;
    WMinX,WMinY : Integer ;
    StFichier : String ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure NewEnreg ;
    Procedure ChargeEnreg ;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure Initialisation ;
    Function  Supprime : Boolean ;
    Function  VerifiSiExiste : Boolean ;
    procedure ClickZoom ;
    Function  OkZoom (Var A : TActionFiche) : Boolean ;
  public
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  PrintDBG ;

procedure ParamAuxSISCO ;
var FCorrSISCO: TFCorrSISCO;
    PP : THPanel ;
begin
FCorrSISCO:=TFCorrSISCO.Create(Application) ;
FCorrSISCO.FType:='SIS' ;
FCorrSISCO.StFichier:='' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FCorrSISCO.ShowModal ;
    finally
     FCorrSISCO.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FCorrSISCO,PP) ;
   FCorrSISCO.Show ;
   END ;
end;

procedure ParamAuxRECUPSISCO(StFichier : String) ;
var FCorrSISCO: TFCorrSISCO;
    PP : THPanel ;
begin
FCorrSISCO:=TFCorrSISCO.Create(Application) ;
FCorrSISCO.FType:='SIS' ;
FCorrSISCO.StFichier:=StFichier ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FCorrSISCO.ShowModal ;
    finally
     FCorrSISCO.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FCorrSISCO,PP) ;
   FCorrSISCO.Show ;
   END ;
end;

procedure MajSISCOCOR(StFichier : String) ;
Var Q : TQuery ;
    Fichier : TextFile ;
    St : String ;
    OkOk : Boolean ;
BEGIN
ExecuteSQL('DELETE FROM CORRESP WHERE CR_TYPE="SIS"') ;
Q:=OpenSQL('SELECT * FROM CORRESP WHERE CR_TYPE="SIS"',FALSE) ;
OkOk:=Not Q.Eof ;
If OkOk Then BEGIN Ferme(Q) ; Exit END Else
  BEGIN
  AssignFile(Fichier,StFichier) ; {$i-} Reset(Fichier) ; {$i+}
  If Ioresult<>0 Then BEGIN Ferme(Q) ; Exit ; END ;
  If Not Eof(Fichier) THen
    BEGIN
    Readln(Fichier,St) ; If Copy(St,1,11)<>'***DEBUT***' Then BEGIN Ferme(Q) ; Exit ; END ;
    END ;
  InitMove(100,'') ;
  OkOk:=FALSE ;
  While Not Eof(Fichier) Do
    BEGIN
    MoveCur(FALSE) ;
    ReadLn(Fichier,St) ;
    If (Copy(St,1,2)='08') And ((Copy(St,33,1)='C') Or (Copy(St,33,1)='F')) Then
      BEGIN
      OkOk:=TRUE ;
      Q.Insert ;
      InitNew(Q) ;
      Q.FindField('CR_TYPE').AsString:='SIS' ;
      Q.FindField('CR_CORRESP').AsString:=BourreEtLess(Trim(Copy(St,3,10)),fbGene) ;
      Q.FindField('CR_LIBELLE').AsString:=BourreEtLess(Trim(Copy(St,13,10)),fbAux) ;
      Q.FindField('CR_ABREGE').AsString:=BourreEtLess(Trim(Copy(St,23,10)),fbAux) ;
      Q.Post ;
      END Else If (Copy(St,1,2)<>'08') And OkOk Then Break ;
    END ;
  Close(Fichier) ;
  FiniMove ;
  END ;
Ferme(Q) ;
END ;


Procedure TFCorrSISCO.Initialisation ;
BEGIN
If StFichier<>'' Then BRecupSISCO.Visible:=TRUE ;
TCorresp.SetRange([FType],[FType]) ;
FAvertir:=False ;
if(TCorresp.EOF)And(TCorresp.BOF) then Bouge(nbInsert) ;
END ;

procedure TFCorrSISCO.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TCorresp.Open ;
Initialisation ;
FListe.SetFocus ;
end;


procedure TFCorrSISCO.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFCorrSISCO.TCorrespNewRecord(DataSet: TDataset);
begin NewEnreg ; end;

procedure TFCorrSISCO.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCorrSISCO.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCorrSISCO.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCorrSISCO.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCorrSISCO.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFCorrSISCO.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFCorrSISCO.BDefaireClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFCorrSISCO.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFCorrSISCO.NewEnreg ;
BEGIN
InitNew(TCorresp) ;
TCORRESPCR_TYPE.AsString:=FType ;
FListe.Columns.Items[0].ReadOnly:=False ;
FListe.SelectedIndex:=0 ; FListe.SetFocus ;
END ;

Procedure TFCorrSISCO.ChargeEnreg ;
BEGIN
// code ...
END ;

Function TFCorrSISCO.VerifiSiExiste : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select CR_TYPE,CR_CORRESP From CORRESP Where CR_TYPE="'+TCORRESPCR_TYPE.AsString+'" And '+
              'CR_CORRESP="'+TCORRESPCR_CORRESP.AsString+'"',True) ;
Result:=(Not QLoc.Eof) ; Ferme(QLoc) ;
END ;

Function TFCorrSISCO.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TCORRESP.state in [dsEdit,dsInsert]=False then Exit ;
if TCORRESP.state in [dsEdit,dsInsert] then
   BEGIN
   if TCORRESPCR_CORRESP.AsString='' then
      BEGIN MsgBox.Execute(2,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=0 ; Exit ; END ;
   if TCORRESPCR_LIBELLE.AsString='' then
      BEGIN MsgBox.Execute(3,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=1 ; Exit ; END ;
   if TCORRESPCR_ABREGE.AsString='' then
      BEGIN MsgBox.Execute(3,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=1 ; Exit ; END ;
   END else exit ;
if TCORRESP.state in [dsInsert] then
   BEGIN
   if VerifiSiExiste then
      BEGIN MsgBox.Execute(4,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=0 ; Exit ; END ;
   END ;
Result:=TRUE  ;
END ;

Function TFCorrSISCO.OnSauve : boolean ;
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

Function TFCorrSISCO.Bouge(Button: TNavigateBtn) : boolean ;
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

procedure TFCorrSISCO.SCORRESPDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not((TCORRESP.State in [dsInsert,dsEdit])) ;
BDelete.Enabled:=Not((TCORRESP.State in [dsInsert,dsEdit])) ;
if(TCORRESP.Eof)And(TCORRESP.Bof) then BDelete.Enabled:=False ;
//if TCORRESP.State=dsBrowse then Fliste.Columns.Items[0].ReadOnly:=True ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TCORRESP.BOF;
   DnEnable := Enabled and not TCORRESP.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END ;
end;

procedure TFCorrSISCO.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

procedure TFCorrSISCO.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil,Caption,'') ; end;

procedure TFCorrSISCO.TCORRESPAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrSISCO.TCORRESPAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrSISCO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
  BEGIN
  END ;
if Parent is THPanel then
   BEGIN
   Action:=caFree ;
   END ;
end;

procedure TFCorrSISCO.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : Boolean ;
begin
if(ssCtrl in Shift)And(Key=VK_DELETE)And(TCORRESP.EOF)And(TCORRESP.BOF) then
  BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)And(Key=VK_DELETE)then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
Vide:=(Shift=[]) ;
Case Key of
   VK_F5     : if (Vide) then BEGIN Key:=0 ; ClickZoom ; END ;
   END ;
end;

procedure TFCorrSISCO.TCORRESPPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
if TCORRESP.State=dsInsert then
   BEGIN MsgBox.Execute(4,'','') ; Fliste.SelectedIndex:=0 ; Action:=daAbort ; END ;
end;

Function TFCorrSISCO.Supprime : Boolean ;
BEGIN
Result:=False ; if (MsgBox.execute(1,'','')<>mrYes) then Exit ;
Result:=True ;
END ;

procedure TFCorrSISCO.SCORRESPUpdateData(Sender: TObject);
begin
if(Trim(TCORRESPCR_CORRESP.AsString)='') And (Trim(TCORRESPCR_LIBELLE.AsString)='')
   then BEGIN TCORRESP.Cancel ; Exit ; END ;
if(Trim(TCORRESPCR_LIBELLE.AsString)<>'') And (Trim(TCORRESPCR_CORRESP.AsString)='') then
  BEGIN MsgBox.Execute(2,'','') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TCORRESPCR_CORRESP.AsString)<>'') And ((Trim(TCORRESPCR_LIBELLE.AsString)='') Or (Trim(TCORRESPCR_ABREGE.AsString)=''))then
  BEGIN MsgBox.Execute(3,'','') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
end ;

procedure TFCorrSISCO.FListeRowEnter(Sender: TObject);
begin if TCORRESP.State=dsInsert then Fliste.SelectedIndex:=0 ; end;

procedure TFCorrSISCO.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCorrSISCO.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFCorrSISCO.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

Function TFCorrSISCO.OkZoom (Var A : TActionFiche) : Boolean ;
Var CCo : TConcept ;
BEGIN
Result:=FALSE ;
If FListe.SelectedField.FieldName<>'CR_CORRESP' Then Exit ;
CCo:=TConcept(ccGenModif) ; Cache.ZoomTable:=tzGCollectif ;
if Not ExJaiLeDroitConcept(CCo,False) then A:=taConsult else A:=taModif ;
Result:=TRUE ;
END ;

procedure TFCorrSISCO.ClickZoom ;
Var A   : TActionFiche ;
begin
If Not OkZoom(A) Then Exit ;
Cache.Text:=Trim(FListe.SelectedField.Text) ;
if GChercheCompte(Cache,Nil) then
  BEGIN
  If TCorresp.State=dsBrowse Then TCorresp.Edit ;
  FListe.Fields[FListe.SelectedIndex].AsString:=Cache.Text ;
  END ;
end;

procedure TFCorrSISCO.FListeDblClick(Sender: TObject);
begin
ClickZoom ;
end;


procedure TFCorrSISCO.BRecupSISCOClick(Sender: TObject);
begin
If MsgBox.Execute(8,'','')<>mrYes Then Exit ;
MajSISCOCOR(StFichier) ;
TCorresp.Close ; TCorresp.Open ;
TCorresp.SetRange([FType],[FType]) ;
end;

end.
