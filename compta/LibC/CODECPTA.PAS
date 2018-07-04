unit CODECPTA;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons,Dialogs,
     StdCtrls, ExtCtrls, Hctrls, {$IFNDEF DBXPRESS}dbtables, HSysMenu,
  hmsgbox, Db, Hqry, DBCtrls, Hcompte, Mask, Grids, DBGrids, HDB{$ELSE}uDbxDataSet{$ENDIF}, DB, Grids, DBGrids, Mask, DBCtrls,
     Ent1,mulgene, HDB, hmsgbox, Hcompte, sysutils, General, HEnt1,
     HSysMenu, HQry ;

Procedure ParamCodeCompta ;

type
  TFCodeCompta = class(TForm)
    TCodeCompta: THTable;
    SCodeCompta: TDataSource;
    TCC: THTable;
    SCC: TDataSource;
    TRegime: THTable;
    TCodeComptaCP_CODECOMPTA: TStringField;
    TCodeComptaCP_REGIMETVA: TStringField;
    TCodeComptaCP_CPTEGENEACH: TStringField;
    TCodeComptaCP_CPTEGENEVTE: TStringField;
    Panel1: TPanel;
    FListe: THDBGrid;
    HPanel2: TPanel;
    TCC_CODE: THLabel;
    TCC_LIBELLE: THLabel;
    CC_CODE: TDBEdit;
    CC_LIBELLE: TDBEdit;
    FListe2: THDBGrid;
    DBNav: TDBNavigator;
    TCCCC_TYPE: TStringField;
    TCCCC_CODE: TStringField;
    TCCCC_LIBELLE: TStringField;
    TCCCC_ABREGE: TStringField;
    MsgBox: THMsgBox;
    Cache: THCpteEdit;
    HPB: TPanel;
    FAutoSave: TCheckBox;
    BAide: TBitBtn;
    BAnnuler: TBitBtn;
    BValider: TBitBtn;
    BImprimer: TBitBtn;
    BFirst: TBitBtn;
    BPrev: TBitBtn;
    BNext: TBitBtn;
    BLast: TBitBtn;
    BInsert: TBitBtn;
    BDelete: TBitBtn;
    BFerme: TBitBtn;
    Bevel1: TBevel;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure TCCNewRecord(DataSet: TDataset);
    procedure TCCAfterPost(DataSet: TDataset);
    procedure FListeColExit(Sender: TObject);
    procedure TCCBeforeDelete(DataSet: TDataset);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SCCDataChange(Sender: TObject; Field: TField);
    procedure BImprimerClick(Sender: TObject);
    procedure TCodeComptaAfterDelete(DataSet: TDataSet);
    procedure TCodeComptaAfterPost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListe2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SCodeComptaDataChange(Sender: TObject; Field: TField);
    procedure SCodeComptaUpdateData(Sender: TObject);
    procedure TCodeComptaNewRecord(DataSet: TDataSet);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private { Private declarations }
    FAvertir,Insere,UnNouveau : boolean ;
    procedure InitRegime ;
    Function ChercheCpte( T : TField ; Vide : Boolean ; ZoomTable : TZoomTable) : byte ;
    Procedure ChargeEnreg ;
    Procedure NewEnreg ;
    Function EnregOK : boolean ;
    Function OnSauve : boolean ;
    Function Bouge(Button: TNavigateBtn) : boolean ;
    Procedure SupprimeRegime ;
    Function  SiExisteCode : Boolean ;
    Procedure GeleLesBoutons ;
    Function  MajRegimeSurInsert : Boolean ;
    Function  EnregOk1 : Boolean ;
  public  { Public declarations }
  end;


implementation

uses PrintDBG;

{$R *.DFM}

procedure TFCodeCompta.FormShow(Sender: TObject);
Var i : Byte ;
    St : String ;
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FAvertir:=False ; Insere:=False ;
TCC.Open ; TCodeCompta.Open ;
TCC.SetRange(['CCA'],['CCA']) ; St:='' ;
for i:=1 to VH^.Cpta[fbGene].Lg  do St:=St+'a' ;
TCodeComptaCP_CPTEGENEACH.EditMask:='>'+St+';0; ';
TCodeComptaCP_CPTEGENEVTE.EditMask:='>'+St+';0; ';
if (TCodeCompta.Eof) And (TCodeCompta.Bof) then Bouge(nbInsert) ;
end;

procedure TFCodeCompta.InitRegime  ;
Var StCode,StREG : string[3] ;
BEGIN
if TCCCC_TYPE.AsString<>'CCA' then Exit ;
if Not TRegime.Active then TRegime.Open ;
TRegime.SetRange(['RTV'],['RTV']) ; TRegime.First ;
StCode:=TCCCC_CODE.AsString ;
if (StCode<>'') or UnNouveau then
   BEGIN
   While Not TRegime.Eof do
      BEGIN
      StREG:=TRegime.FindField('CC_CODE').AsString ;
      if Not FindLaKey(TCodeCompta,[StCode,StREG]) then
         BEGIN
         Insere:=True ; TCodeCompta.Insert ; InitNew(TCodeCompta) ;
         TCodeComptaCP_CODECOMPTA.AsString:=StCode ;
         TCodeComptaCP_REGIMETVA.AsString:=StREG ;
         TCodeCompta.Post ;
         END else Insere:=False ;
      TRegime.Next ;
      END ;
   END ;
Insere:=False ; TCodeCompta.First ; TRegime.Close ;
END ;

procedure TFCodeCompta.TCCNewRecord(DataSet: TDataset);
begin Newenreg ; end;

procedure TFCodeCompta.TCCAfterPost(DataSet: TDataset);
begin InitRegime ; end;

Function TFCodeCompta.ChercheCpte( T : TField ; Vide : Boolean ; ZoomTable : TZoomTable) : byte ;
Var St    : String ;
BEGIN
ChercheCpte:=0 ;
St:=uppercase(T.AsString) ; Cache.Text:=St ; Cache.ZoomTable:=ZoomTable ;
if ((Vide) and (St='')) then exit ;
if GChercheCompte(Cache,Nil) then
   BEGIN
   if St<>Cache.Text then
      BEGIN
      if TCodeCompta.State in [dsEdit] then else TCodeCompta.edit ;
      T.AsString:=Cache.Text ;
      ChercheCpte:=1 ;
      END ;
   ChercheCpte:=2 ;
   END ;
END ;

procedure TFCodeCompta.FListeColExit(Sender: TObject);
Var i : Byte ;
    St : String ;
begin
if(FListe.SelectedField.FieldName='CP_CPTEGENEACH') Or
  (FListe.SelectedField.FieldName='CP_CPTEGENEVTE') then
  BEGIN
  FListe.SelectedIndex:=FListe.SelectedIndex ;
  if FListe.SelectedField.FieldName='CP_CPTEGENEACH' then i:=ChercheCpte(FListe.SelectedField,TRUE,tzGCharge)
                                                     else i:=ChercheCpte(FListe.SelectedField,TRUE,tzGProduit) ;
  St:=TCodeComptaCP_CPTEGENEACH.AsString ;
  if St<>'' then
     if Length(St)<VH^.Cpta[fbGene].Lg then
         BEGIN St:=BourreLaDonc(St,fbGene) ; TCodeComptaCP_CPTEGENEACH.AsString:=St ; END ;
  St:=TCodeComptaCP_CPTEGENEVTE.AsString ;
  if St<>'' then
     if Length(St)<VH^.Cpta[fbGene].Lg then
         BEGIN St:=BourreLaDonc(St,fbGene) ; TCodeComptaCP_CPTEGENEVTE.AsString:=St ; END ;
  END ;
end;

procedure TFCodeCompta.TCCBeforeDelete(DataSet: TDataset);
BEGIN SupprimeRegime ; END ;

Procedure ParamCodeCompta;
var FCodeCompta: TFCodeCompta;
begin
  if Blocage(['nrCloture'],False,'nrAucun') then exit ;
  FCodeCompta:= TFCodeCompta.Create(Application) ;
  try
    FCodeCompta.ShowModal ;
    finally
    FCodeCompta.Free ;
    end ;
Screen.Cursor:=crDefault ;
end ;

procedure TFCodeCompta.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCodeCompta.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCodeCompta.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCodeCompta.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCodeCompta.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFCodeCompta.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFCodeCompta.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFCodeCompta.BAnnulerClick(Sender: TObject);
begin
if TCC.State<>dsInsert then TCodeCompta.Cancel ;
Bouge(nbCancel) ;
end;

procedure TFCodeCompta.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFCodeCompta.NewEnreg ;
BEGIN
TCCCC_TYPE.AsString:='CCA' ;
UnNouveau:=True ;  InitRegime ; UnNouveau:=False ;
CC_CODE.Enabled:=True ; CC_CODE.SetFocus ;
END ;

procedure TFCodeCompta.FormCloseQuery(Sender: TObject;  var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

Procedure TFCodeCompta.ChargeEnreg ;
BEGIN
CC_CODE.Enabled:=False ;
if TCodeCompta.Modified then
   if Not EnregOk1 then
      BEGIN TCodeCompta.Cancel ; MsgBox.Execute(8,'','') ; END ;
InitRegime ;
END ;

Function TFCodeCompta.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TCC.state in [dsEdit,dsInsert]=False then Exit ;
if TCC.state in [dsEdit,dsInsert] then
   BEGIN
   if TCCCC_CODE.AsString='' then BEGIN MsgBox.Execute(2,'','') ; CC_CODE.SetFocus ; Exit ; END ;
   if TCCCC_LIBELLE.AsString='' then BEGIN MsgBox.Execute(3,'','') ; CC_LIBELLE.SetFocus ; Exit ; END ;
   END ;
if TCC.state in [dsInsert] then
   BEGIN
   if SiExisteCode then BEGIN MsgBox.Execute(5,'','') ; CC_CODE.SetFocus ; Exit ; END ;
   END ;
result:=TRUE  ;
END ;

Function TFCodeCompta.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TCC.Modified then
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

Function TFCodeCompta.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : BEGIN
                      if Not MajRegimeSurInsert then Exit ;
                      if Not EnregOK then Exit ;
                      END ;
   nbDelete         : if MsgBox.execute(1,'','')<>mrYes then exit ;
   end ;
if(Button=nbCancel) And (TCC.State=dsInsert)then SupprimeRegime ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[4]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Procedure TFCodeCompta.GeleLesBoutons ;
BEGIN
BInsert.Enabled:=(Not(TCC.State In [dsEdit,dsInsert])) And (Not(TCodeCompta.State In [dsEdit,dsInsert])) ;
BDelete.Enabled:=(Not(TCC.State In [dsEdit,dsInsert])) And (Not(TCodeCompta.State In [dsEdit,dsInsert])) ;
if(TCC.EOF)And(TCC.BOF) then BDelete.Enabled:=False ;
END ;

procedure TFCodeCompta.SCodeComptaDataChange(Sender: TObject; Field: TField);
begin GeleLesBoutons ; end;

procedure TFCodeCompta.SCCDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
GeleLesBoutons ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TCC.BOF;
   DnEnable := Enabled and not TCC.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
   if ((Field.FieldName='CC_LIBELLE') and (TCCCC_ABREGE.AsString='')) then
      TCCCC_ABREGE.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFCodeCompta.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,HPanel2,Caption,'') ; end;

procedure TFCodeCompta.TCodeComptaAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCodeCompta.TCodeComptaAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCodeCompta.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin if FAvertir then AvertirTable('ttCodeCompta') ; end;

procedure TFCodeCompta.FListe2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if(ssCtrl in Shift)And(Key=VK_DELETE) then BEGIN Bouge(nbDelete) ; Key:=0 END ;
end;

Procedure TFCodeCompta.SupprimeRegime ;
BEGIN
ExecuteSql('Delete From CODECPTA Where CP_CODECOMPTA="" Or CP_CODECOMPTA="'+TCCCC_CODE.AsString+'"') ;
END ;

Function TFCodeCompta.SiExisteCode : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CC_CODE From CHOIXCOD Where CC_TYPE="CCA" And '+
           'CC_CODE="'+TCCCC_CODE.AsString+'"',True) ;
Result:=Not Q.Eof ; Ferme(Q) ;
END ;

Function TFCodeCompta.MajRegimeSurInsert : Boolean ;
BEGIN
Result:=False ;
if(TCC.State<>dsInsert) then BEGIN Result:=True ; Exit ; END ;
if TCC.State=dsInsert then
   if Not EnregOK then Exit ;
if TCodeCompta.State=dsEdit then TCodeCompta.Post ;
if TCC.State=dsInsert then
   ExecuteSql('Update CODECPTA Set CP_CODECOMPTA="'+TCCCC_CODE.AsString+'" Where CP_CODECOMPTA=""') ;
Result:=True ;
END ;

procedure TFCodeCompta.SCodeComptaUpdateData(Sender: TObject);
begin if Not Enregok1 then SysUtils.Abort ; end;

Function TFCodeCompta.EnregOk1 : Boolean ;
Var QLoc : TQuery ;
BEGIN
Result:=False ;
if TCodeComptaCP_CPTEGENEACH.AsString<>'' then
   BEGIN
   QLoc:=OpenSql('Select G_GENERAL From GENERAUX Where G_NATUREGENE="CHA" and '+
                 'G_GENERAL="'+BourreLaDonc(TCodeComptaCP_CPTEGENEACH.AsString,fbGene)+'"',True) ;
   if QLoc.Eof then
      BEGIN FListe.SelectedIndex:=1 ; MsgBox.Execute(6,'','') ; Ferme(QLoc) ; Exit ; END ;
   Ferme(QLoc) ;
   END ;
if TCodeComptaCP_CPTEGENEVTE.AsString<>'' then
   BEGIN
   QLoc:=OpenSql('Select G_GENERAL From GENERAUX Where G_NATUREGENE="PRO" and '+
                 'G_GENERAL="'+BourreLaDonc(TCodeComptaCP_CPTEGENEVTE.AsString,fbGene)+'"',True) ;
   if QLoc.Eof then
      BEGIN FListe.SelectedIndex:=2 ; MsgBox.Execute(7,'','') ; Ferme(QLoc) ; Exit ; END ;
   Ferme(QLoc) ;
   END ;
Result:=True ;
END ;

procedure TFCodeCompta.TCodeComptaNewRecord(DataSet: TDataSet);
begin if Not Insere then TCodeCompta.Cancel ; end;

procedure TFCodeCompta.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if Not(Key in [VK_UP,VK_DOWN,VK_PRIOR,VK_NEXT,VK_TAB])then Exit ;
if(Key in [VK_TAB])And(Fliste.SelectedIndex<>3) then Exit ;
FListe.SelectedIndex:=FListe.SelectedIndex ;
if Not TCodeCompta.Modified then Exit ;
if Not Enregok1 then BEGIN Key:=0 ; Exit ; END ;
end;

end.
