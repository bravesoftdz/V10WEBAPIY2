{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 26/04/2002
Modifié le ... : 26/04/2002
Description .. : Depuis 04/2002, passage en E-AGL :
Suite ........ :  - Passage de la fiche dans Décla "CPTABLIBRELIB",
Suite ........ :  - Création d'une unité TOF : CPTABLIBRELIB_TOF,
Mots clefs ... : TABLES LIBRES
*****************************************************************}
unit LibTabLi;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, DBCtrls, HSysMenu, hmsgbox, Grids, DBGrids, HDB,
  Buttons, ExtCtrls,
{$IFNDEF CCS3}
  LibChpLi,
{$ENDIF}
  Ent1, Hent1, Hctrls,PrintDBG,CPTABLIBRELIB_TOF ;

Procedure ModifLibelleTableLibre ;
Procedure AGLModifLibelleTableLibre ;

type
  TFLibTabLi = class(TForm)
    HPB: TPanel;
    BAnnuler: THBitBtn;
    BFirst: THBitBtn;
    BPrev: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    FListe: THDBGrid;
    MsgBox: THMsgBox;
    SCommun: TDataSource;
    HMTrad: THSystemMenu;
    DBNav: TDBNavigator;
    FAutoSave: TCheckBox;
    PTop: TPanel;
    Tl: THValComboBox;
    TTl: TLabel;
    QCommun: TQuery;
    HMCombo: THMsgBox;
    bChamps: THBitBtn;
    procedure BFermeClick(Sender: TObject);
    procedure TlChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SCommunDataChange(Sender: TObject; Field: TField);
    procedure SCommunUpdateData(Sender: TObject);
    procedure QCommunNewRecord(DataSet: TDataSet);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeColEnter(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure bChampsClick(Sender: TObject);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    WMinX,WMinY : Integer ;
    CoLib : String ;
    Procedure ChargeCombos ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    procedure InverseSelection ;
    procedure MajParamLib ;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Procedure ModifLibelleTableLibre ;
var FLibTabLi : TFLibTabLi ;
BEGIN
FLibTabLi:=TFLibTabLi.Create(Application) ;
 Try
  FLibTabLi.ShowModal ;
 Finally
  FLibTabLi.Free ;
 End ;
SourisNormale ;
END ;

procedure TFLibTabLi.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFLibTabLi.ChargeCombos ;
Var i : Integer ;
    St : String ;
    it,va : String ;
BEGIN
Tl.Values.Clear ; Tl.Items.Clear ;
for i:=0 to HMCombo.Mess.Count-1 do
   BEGIN
   St:=HMCombo.Mess[i] ;
   va:=ReadTokenSt(St) ; it:=ReadTokenSt(St) ;
   if ((EstSerie(S5)) or (EstSerie(S3))) then
      BEGIN
      if ((*(va='E') or*) (va='A') or (va='U')) then Continue ;
      if EstSerie(S3) then if ((va='B') or (va='D')) then Continue ;
      END ;
   Tl.Values.Add(va) ; Tl.Items.Add(it) ;
   END ;
END ;

procedure TFLibTabLi.TlChange(Sender: TObject);
begin
OnSauve ;
QCommun.Close ; QCommun.SQl.Clear ;
If TL.Value='E' Then
  BEGIN
  if (EstSerie(S3)) then
     BEGIN
     QCommun.Sql.Add('Select * From COMMUN Where CO_TYPE="NAT" And CO_CODE Like "'+Tl.Value+'%" AND CO_CODE<="'+Tl.Value+'00" ORDER BY CO_CODE') ;
     END else
  if (EstSerie(S5)) then
     BEGIN
     QCommun.Sql.Add('Select * From COMMUN Where CO_TYPE="NAT" And CO_CODE Like "'+Tl.Value+'%" AND CO_CODE<="'+Tl.Value+'01" ORDER BY CO_CODE') ;
     END else
     BEGIN
     QCommun.Sql.Add('Select * From COMMUN Where CO_TYPE="NAT" And CO_CODE Like "'+Tl.Value+'%" ORDER BY CO_CODE') ;
     END ;
  END Else
  BEGIN
  if EstSerie(S3) then
     BEGIN
     QCommun.Sql.Add('Select * From COMMUN Where CO_TYPE="NAT" And CO_CODE Like "'+Tl.Value+'%" AND CO_CODE<="'+Tl.Value+'02" ORDER BY CO_CODE') ;
     END else
     BEGIN
     QCommun.Sql.Add('Select * From COMMUN Where CO_TYPE="NAT" And CO_CODE Like "'+Tl.Value+'%" ORDER BY CO_CODE') ;
     END ;
  END ;
ChangeSql(QCommun) ;
FListe.Columns.Items[1].FieldName:='CO_'+coLib ;
QCommun.Open ;
end;

procedure TFLibTabLi.InverseSelection ;
begin
if QCommun.State<>dsEdit then QCommun.Edit ;
if QCommun.FindField('CO_ABREGE').AsString='-' then QCommun.FindField('CO_ABREGE').AsString:='X'
                                               else QCommun.FindField('CO_ABREGE').AsString:='-' ;
end ;

procedure TFLibTabLi.FormShow(Sender: TObject);
begin
Colib:='LIBELLE' ;
ChargeCombos ; Tl.Value:=Tl.Values[0] ;
if EstSerie(S3) then bChamps.Visible:=False ;
end;

procedure TFLibTabLi.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFLibTabLi.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; end;

procedure TFLibTabLi.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) end;

procedure TFLibTabLi.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFLibTabLi.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFLibTabLi.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFLibTabLi.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFLibTabLi.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFLibTabLi.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
BFerme.SetFocus ;
CanClose:=OnSauve ;
end;

Function TFLibTabLi.EnregOK : boolean ;
BEGIN
Result:=False ;
if QCommun.state in [dsEdit,dsInsert]=False then Exit ;
if QCommun.state in [dsEdit] then
   if QCommun.FindField('CO_'+CoLib).AsString='' then
      BEGIN MsgBox.Execute(1,Caption,'') ; FListe.SelectedIndex:=1 ; FListe.SetFocus ; Exit ; END ;
Result:=True ;
END ;

Function TFLibTabLi.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if QCommun.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,caption,'') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFLibTabLi.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[2]) ;
Result:=TRUE ;
END ;

procedure TFLibTabLi.SCommunDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
BEGIN
bChamps.Enabled := Not (QCommun.State In [dsEdit,dsInsert]) ;
if QCommun.State=dsBrowse then bChamps.Enabled := QCommun.FindField('CO_ABREGE').AsString='X' ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not QCommun.BOF;
   DnEnable := Enabled and not QCommun.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   END ;
END ;

procedure TFLibTabLi.MajParamLib ;
var Q : TQuery ;
    Code,TypeTable : String ;
    n : Integer ;
BEGIN
Code:=(QCommun.FindField('CO_CODE').AsString) ;
n:=StrToInt(Copy(Code,Length(Code),1)) ;
TypeTable:=Tl.Value ;
if Tl.Value='A' then TypeTable:='Y' else
  if Tl.Value='U' then TypeTable:='BE' ;
Q:=OpenSQL('SELECT * FROM PARAMLIB WHERE PL_TABLE="'+TypeTable+'" AND PL_CHAMP="'+TypeTable+'_TABLE'+IntToStr(n)+'"',False) ;
if not Q.Eof then
  BEGIN
  Q.Edit ; //vt à modifier traduction ?
  Q.FindField('PL_'+CoLib).AsString:=QCommun.FindField('CO_'+CoLib).AsString ;
  Q.FindField('PL_VISIBLE').AsString:=QCommun.FindField('CO_ABREGE').AsString ;
  Q.Post ;
  END ;
Ferme(Q) ;
END ;

procedure TFLibTabLi.SCommunUpdateData(Sender: TObject);
begin
if(Trim(QCommun.FindField('CO_CODE').AsString)='') And
  (Trim(QCommun.FindField('CO_'+CoLib).AsString)='') Then
  BEGIN QCommun.Cancel ; Exit ; END ;
if(Trim(QCommun.FindField('CO_CODE').AsString)<>'') And
  (Trim(QCommun.FindField('CO_'+CoLib).AsString)='')then
  BEGIN MsgBox.Execute(1,Caption,'') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
MajParamLib ;
end;

procedure TFLibTabLi.QCommunNewRecord(DataSet: TDataSet);
begin Bouge(nbCancel) ; end;

procedure TFLibTabLi.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin if((ssCtrl in Shift) and (Key=VK_DELETE)) then Key:=0 ; end;

procedure TFLibTabLi.FListeColEnter(Sender: TObject);
begin if Fliste.SelectedIndex=0 then Fliste.SelectedIndex:=1  ; end;

procedure TFLibTabLi.BImprimerClick(Sender: TObject);
begin PrintDBGrid(FListe,Nil,VireSouligne(TTL.Caption+' '+TL.Text),''); end;

procedure TFLibTabLi.bChampsClick(Sender: TObject);
begin
{$IFNDEF CCS3}
AGLModifLibelleChampLibre(QCommun.FindField('CO_CODE').AsString,QCommun.FindField('CO_'+CoLib).AsString) ;
{$ENDIF}
end;

procedure TFLibTabLi.FListeKeyPress(Sender: TObject; var Key: Char);
begin
if FListe.SelectedField<>QCommun.FindField('CO_ABREGE') then exit ;
if Key=' ' then InverseSelection ;
if (Key<>'-') and (Key<>'X') and (Key<>'x') then Key:=#0 ;
if Key='x' then Key:='X' ;
end;

procedure TFLibTabLi.FListeDblClick(Sender: TObject);
begin
InverseSelection ;
end;

procedure TFLibTabLi.FListeRowEnter(Sender: TObject);
begin
bChamps.Enabled := QCommun.FindField('CO_ABREGE').AsString='X' ;
end;

procedure TFLibTabLi.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

// =============================================================================
procedure AGLModifLibelleTableLibre;
var Rep : Integer;
begin
	CPLanceFiche_ParamTablesLibres;
end;
// =============================================================================

end.
