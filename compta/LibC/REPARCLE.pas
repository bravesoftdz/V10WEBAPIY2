unit REPARCLE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, HDB, DBCtrls, StdCtrls, Hctrls, Mask, Buttons, ExtCtrls,
  hmsgbox, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Ent1, Hcompte, ComCtrls, Ventil, HEnt1, HSysMenu, ParamDBG,
  Hqry, HPanel, UiUtil, SaisComm, HTB97, ADODB ;

Procedure ParamRepartCle ;

type
  TFRepartcle = class(TForm)
    SRepartCle : TDataSource;
    TRepartCle: THTable;
    DBNav      : TDBNavigator;
    FListe     : THDBGrid;
    Pappli     : TPanel;
    MsgBox     : THMsgBox;
    RE_CLE             : TDBEdit;
    RE_LIBELLE         : TDBEdit;
    RE_ABREGE          : TDBEdit;
    RE_MODEREPARTITION : THDBValComboBox;
    RE_QUALIFQTE       : THDBValComboBox;
    TRE_CLE            : TLabel;
    TRE_LIBELLE        : THLabel;
    TRE_ABREGE         : TLabel;
    TRE_MODEREPARTITION: TLabel;
    TRE_QUALIFQTE      : TLabel;
    PAxe: TTabControl;
    TRepartCleRE_AXE: TStringField;
    TRepartCleRE_CLE: TStringField;
    TRepartCleRE_LIBELLE: TStringField;
    TRepartCleRE_ABREGE: TStringField;
    TRepartCleRE_MODEREPARTITION: TStringField;
    TRepartCleRE_QUALIFQTE: TStringField;
    HPB: TToolWindow97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BFerme: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BVentil: TToolbarButton97;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    Dock: TDock97;
    TRepartCleRE_COMPTES: TStringField;
    RE_COMPTES: TDBEdit;
    TRE_COMPTES: TLabel;
    procedure BFermeClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
    procedure SRepartCleDataChange(Sender: TObject; Field: TField);
    procedure SRepartCleStateChange(Sender: TObject);
    procedure SRepartCleUpdateData(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure PAxeChange(Sender: TObject);
    procedure Autorise(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TRepartCleAfterDelete(DataSet: TDataSet);
    procedure TRepartCleAfterPost(DataSet: TDataSet);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
  private
   Modifier : Boolean ;
   FAxe     : String ;
   FAvertir : Boolean ;
   Acreer   : Boolean ;
   Procedure ChargeEnreg ;
   Procedure NewEnreg ;
   Function  OnSauve : boolean ;
   Function  EnregOK : boolean ;
   Function  Bouge(Button: TNavigateBtn) : boolean ;
   Function  Supprime:Boolean ;
   Function  VerifisiExiste:Boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses printDBG, UtilPgi ;

procedure TFRepartcle.BImprimerClick(Sender: TObject);
begin
PrintDBGrid (Nil,Nil,Caption,'PRT_CLEREPAR') ;
end;

procedure TFRepartcle.SRepartCleDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not(TRepartCle.State in [dsEdit,dsInsert]) ;
BDelete.Enabled:=Not(TRepartCle.State in [dsEdit,dsInsert]) ;
BVentil.Enabled:=Not(TRepartCle.State in [dsEdit,dsInsert]) ;
if (TRepartCle.Eof) and (TRepartCle.Bof) then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TRepartCle.BOF;
   DnEnable := Enabled and not TRepartCle.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
   if ((Field.FieldName='RE_LIBELLE') and (RE_ABREGE.Field.AsString='')) then
      RE_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFRepartcle.SRepartCleStateChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFRepartcle.SRepartCleUpdateData(Sender: TObject);
begin
if Modifier then BEGIN Modifier:=False ; if Not OnSauve then Bouge(nbCancel) ;  END ;
end;

procedure TFRepartcle.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFRepartcle.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFRepartcle.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFRepartcle.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFRepartcle.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFRepartcle.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFRepartcle.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFRepartcle.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFRepartcle.BValiderClick(Sender: TObject);
begin Modifier:=False ; Bouge(nbPost) ; end;

procedure TFRepartcle.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

procedure TFRepartcle.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FAvertir:=False ; Acreer:=False ;
TRepartCle.Open ; FAxe:='A1' ;
TRepartCle.SetRange([FAxe],[FAxe]) ;
if (TRepartCle.BOF) and (TRepartCle.EOF) then Bouge(nbInsert) ;
DelTabsSerie(PAxe) ; 
end;

Procedure TFRepartcle.ChargeEnreg ;
BEGIN
RE_CLE.Enabled:=False ;
InitCaption(Self,RE_CLE.text,RE_LIBELLE.text) ;
END ;

Procedure TFRepartcle.NewEnreg ;
BEGIN
InitNew(TRepartCle) ;
RE_MODEREPARTITION.Value:=RE_MODEREPARTITION.Values[0] ;
RE_QUALIFQTE.Value:=RE_QUALIFQTE.Values[0] ;
TRepartCleRE_AXE.AsString:=FAxe ;
TRepartCleRE_MODEREPARTITION.AsString:=RE_MODEREPARTITION.Value ;
RE_CLE.Enabled:=True ; RE_CLE.SetFocus ;
END ;

Function TFRepartcle.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier:=False ;
if TRepartCle.Modified then
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

Function TFRepartcle.EnregOK : boolean ;
BEGIN
Result:=FALSE  ;
Modifier:=True ;
if TRepartCle.state in [dsEdit,dsInsert]=False then Exit ;
if TRepartCle.state in [dsEdit,dsInsert] then
   BEGIN
   if TRepartCleRE_CLE.AsString='' then BEGIN MsgBox.Execute(2,'','') ; RE_CLE.SetFocus ; Exit ; END ;
   if TRepartCleRE_LIBELLE.AsString='' then BEGIN MsgBox.Execute(3,'','') ; RE_LIBELLE.SetFocus ; Exit ; END ;
   if TRepartCleRE_MODEREPARTITION.AsString='' then BEGIN MsgBox.Execute(8,'','') ; RE_MODEREPARTITION.SetFocus ; Exit ; END ;
   END ;
if TRepartCle.state in [dsInsert] then
   BEGIN
   if Not VerifisiExiste then BEGIN MsgBox.Execute(4,'','') ; RE_CLE.SetFocus ; Exit ; END ;
   END ;
Modifier:=False ;
result:=TRUE  ;
END ;

Function TFRepartcle.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
     nblast,nbprior,nbnext,nbfirst,
     nbinsert : if Not OnSauve then Exit ;
     nbPost   : if Not EnregOK then Exit ;
     nbDelete : if Not Supprime then Exit ;
    end ;
if(TRepartCle.state in [dsInsert])And(Button=nbPost)then Acreer:=True else Acreer:=False ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[11]) ;
if Acreer then BEGIN BVentilClick(Nil) ; Acreer:=False ; END ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFRepartcle.Supprime:Boolean ;
BEGIN
Result:=False ;
if MsgBox.execute(1,'','')<>mrYes then Exit ;
ExecuteSql('Delete From Ventil Where V_NATURE="CL'+IntToStr(PAxe.TabIndex+1)+'" And V_COMPTE="'+TRepartCleRE_CLE.AsString+'"') ;
Result:=True ;
END ;

Function  TFRepartcle.VerifisiExiste:Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select RE_AXE,RE_CLE from CLEREPAR where RE_AXE="'+FAxe+'" AND RE_CLE="'+TRepartCleRE_CLE.AsString+'"',True) ;
Result:=(Q.EOF) ; Ferme(Q) ;
END ;

Procedure ParamRepartCle ;
var FRepartcle : TFRepartcle ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FRepartcle:=TFRepartcle.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   Try
    FRepartcle.ShowModal ;
   Finally
    FRepartcle.Free ;
   End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FRepartcle,PP) ;
   FRepartcle.Show ;
   END ;
END ;

procedure TFRepartcle.PAxeChange(Sender: TObject);
begin
if TRepartCle.Modified then
  BEGIN PAxe.TabIndex:=StrToInt(Copy(FAxe,2,1))-1 ; Exit ; END ;
FAxe:='A'+IntToStr(PAxe.TabIndex+1) ;
TRepartCle.SetRange([FAxe],[FAxe]) ;
if TRepartCle.EOF then Bouge(nbInsert) ;
end;

procedure TFRepartcle.Autorise(Sender: TObject);
BEGIN
(*St:=THDBcpteEdit(Sender).Name ;
i:=StrToInt(Copy(St,Length(St)-1,2)) ;
St1:=Copy(St,1,Length(St)-2) ;
if i<=10 then St2:=St1+'0'+IntToStr(i-1) else St2:=St1+IntToStr(i-1) ;
if i<10 then StL:='TT'+St1+'0'+IntToStr(i) else StL:='TT'+St1+IntToStr(i) ;

if(THDBcpteEdit(Sender).Text)<>''then
  BEGIN
   C:=FindComponent(St2) ; CL:=FindComponent(StL) ;
   if THDBcpteEdit(C).Text='' then
      BEGIN
       MsgBox.Execute(10,'','') ; St:=THDBcpteEdit(Sender).Text ;
       THDBcpteEdit(Sender).Text:='' ; TLabel(CL).Caption:='' ;
       for i:=1 to 24 do
         BEGIN
           if i<10 then St2:=St1+'0'+IntToStr(i)
                   else St2:=St1+IntToStr(i) ;
           C:=FindComponent(St2) ;
           if THDBcpteEdit(C).Text='' then
              BEGIN THDBcpteEdit(C).Text:=St ;THDBcpteEdit(C).SetFocus ; Exit ; END ;
         END ;
      END ;
  END ;
*)
END ;

procedure TFRepartcle.BVentilClick(Sender: TObject);
Var St : String ;
begin
if RE_CLE.Text='' then Exit ;
St:=IntToStr(PAxe.TabIndex+1) ; ParamVentil('CL',RE_CLE.Text,St,taModif,True) ;
end;

procedure TFRepartcle.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
  BEGIN
  AvertirTable('ttCleRepart1') ;
  AvertirTable('ttCleRepart2') ;
  AvertirTable('ttCleRepart3') ;
  AvertirTable('ttCleRepart4') ;
  AvertirTable('ttCleRepart5') ;
  END ;
  //SG6 23.02.05
  if IsInside(Self) then CloseInsidePanel(Self);
end;

procedure TFRepartcle.TRepartCleAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFRepartcle.TRepartCleAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFRepartcle.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if(ssCtrl in Shift)AND(Key=VK_DELETE)AND(TRepartCle.EOF)then BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)AND(Key=VK_DELETE)then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFRepartcle.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

end.
