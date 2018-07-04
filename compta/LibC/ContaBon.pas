unit ContaBon;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,Dialogs,
  DBTables, DB, DBCtrls, StdCtrls, Buttons, ExtCtrls, Grids, DBGrids, Spin,
  HDB, Hctrls, Mask, hmsgbox, Ent1, HEnt1, Menus, ComCtrls, HRichEdt,
  HSysMenu, HRichOLE,MajTable, HTB97, HQry, HPanel, UiUtil ;

Procedure ParamAbonnement(EstCompta : Boolean ; Lequel : String ; Mode : TActionFiche) ;

type
  TFContaBon = class(TForm)
    FListe    : TDBGrid;
    Pappli    : TPanel;
    SContaBon : TDataSource;
    TContaBon: THTable;
    BGuide: TToolbarButton97;
    DBNav     : TDBNavigator;
    MsgBox    : THMsgBox;
    GBBlocNote: TGroupBox;
    TCB_CONTRAT         : TLabel;
    TCB_LIBELLE         : TLabel;
    TCB_DATECONTRAT     : TLabel;
    TCB_SEPAREPAR       : TLabel;
    TCB_NOMBREREPETITION: TLabel;
    TCB_RECONDUCTION    : TLabel;
    TCB_DATECREATION    : TLabel;
    TCB_DATEMODIFICATION: TLabel;
    TCB_DEJAGENERE      : TLabel;
    TCB_DATEDERNGENERE  : TLabel;
    TCB_GUIDE           : TLabel;
    CB_CONTRAT         : TDBEdit;
    CB_LIBELLE         : TDBEdit;
    CB_DATECONTRAT     : TDBEdit;
    CB_SEPAREPAR       : THDBValComboBox;
    CB_NBREPETITION: THDBSpinEdit;
    CB_RECONDUCTION    : THDBValComboBox;
    CB_DEJAGENERE      : TDBEdit;
    CB_DATECREATION    : TDBEdit;
    CB_DATEMODIF: TDBEdit;
    CB_DATEDERNGENERE  : TDBEdit;
    CB_GUIDE           : THDBValComboBox;
    CB_BLOCNOTE: THDBRichEditOLE;
    TContaBonCB_CONTRAT         : TStringField;
    TContaBonCB_COMPTABLE       : TStringField;
    TContaBonCB_LIBELLE         : TStringField;
    TContaBonCB_DATECONTRAT     : TDateTimeField;
    TContaBonCB_SEPAREPAR       : TStringField;
    TContaBonCB_RECONDUCTION    : TStringField;
    TContaBonCB_DEJAGENERE: TIntegerField;
    TContaBonCB_DATEDERNGENERE  : TDateTimeField;
    TContaBonCB_GUIDE           : TStringField;
    TContaBonCB_DATECREATION    : TDateTimeField;
    TContaBonCB_UTILISATEUR     : TStringField;
    TContaBonCB_BLOCNOTE        : TMemoField;
    TContaBonCB_SOCIETE         : TStringField;
    CB_ARRONDI: THDBValComboBox;
    TCB_ARRONDI: THLabel;
    TContaBonCB_ARRONDI: TStringField;
    HMTrad: THSystemMenu;
    TContaBonCB_NBREPETITION: TIntegerField;
    TContaBonCB_DATEMODIF: TDateTimeField;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BAnnuler: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    FAutoSave: TCheckBox;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    TCB_QUALIFPIECE: TLabel;
    CB_QUALIFPIECE: THDBValComboBox;
    TContaBonCB_QUALIFPIECE: TStringField;
    procedure BFermeClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure SContaBonDataChange(Sender: TObject; Field: TField);
    procedure SContaBonStateChange(Sender: TObject);
    procedure SContaBonUpdateData(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
    procedure BGuideClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure CB_SEPAREPARChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    EstCompta : Boolean ;
    Lequel    : String ;
    Modifier  : Boolean ;
    Mode      : TActionFiche ;
    OkRes     : Boolean ;
    Procedure ChargeEnreg ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  Supprime :Boolean ;
    Function  VerifiSiExiste : Boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses printDBG, CPGUIDE_TOM,UtilPGI ;

Procedure ParamAbonnement(EstCompta : Boolean ; Lequel : String ; Mode : TActionFiche) ;
var FContaBon : TFContaBon ;
    OkRes : Boolean ;
    PP : THPanel ;
BEGIN
OkRes:=True ;
if Mode=taConsult then OkRes:=False ;
if ((Mode=taModif) and (Lequel<>'')) then OkRes:=False ; {Appel depuis Génère abo}
if OkRes then if _Blocage(['nrCloture','nrBatch'],True,'nrBatch') then Exit ;
FContaBon:=TFContaBon.Create(Application) ;
FContaBon.EstCompta:=EstCompta ;
FContaBon.OkRes:=OkRes ;
FContaBon.Lequel:=Lequel ;
FContaBon.Mode:=Mode ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   Try
    FContaBon.ShowModal ;
   Finally
    FContaBon.Free ;
    if OkRes then _Bloqueur('nrBatch',False) ;
   End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FContaBon,PP) ;
   FContaBon.Show ;
   END ;
END ;

procedure TFContaBon.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(FListe,Nil,MsgBox.Mess[7],'') ;
end;

procedure TFContaBon.BGuideClick(Sender: TObject);
var St : String;
begin
St:=CB_GUIDE.Value;
St:=ParamGuide(St,'ABO',Mode) ;
if Mode=taConsult then Exit ;
CB_GUIDE.Reload ;
//CB_GUIDE.Value:=St ; TContaBonCB_GUIDE.AsString:=St ;
end;

procedure TFContaBon.SContaBonDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not TContaBon.Modified ; BDelete.Enabled:= Not TContaBon.Modified  ;
if(TContaBon.Eof)And(TContaBon.Bof)then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TContaBon.BOF;
   DnEnable := Enabled and not TContaBon.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// if ((Field.FieldName='RE_LIBELLE') and (RE_ABREGE.Field.AsString='')) then
//    RE_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFContaBon.SContaBonStateChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFContaBon.SContaBonUpdateData(Sender: TObject);
begin
if Modifier then BEGIN Modifier:=False ; if Not OnSauve then Bouge(nbCancel) ;  END ;
end;

procedure TFContaBon.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFContaBon.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFContaBon.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFContaBon.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFContaBon.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFContaBon.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFContaBon.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFContaBon.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFContaBon.BValiderClick(Sender: TObject);
begin Modifier:=False ; Bouge(nbPost) ; end;

procedure TFContaBon.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

procedure TFContaBon.FormShow(Sender: TObject);
var St : String ;
begin
ChangeSizeMemo(TContaBonCB_BLOCNOTE) ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TContaBon.Open ;
if EstCompta then St:='X' else St:='-' ;
TContaBon.SetRange([St],[St]) ;
if Lequel<>'' then FindLaKey(TContaBon,[St,Lequel]) ;
BDelete.Enabled:=NOT(TContaBon.EOF) ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
if TContaBon.EOF then Bouge(nbInsert) ;
end;

Procedure TFContaBon.ChargeEnreg ;
BEGIN
InitCaption(Self,TContaBonCB_CONTRAT.AsString,TContaBonCB_LIBELLE.AsString) ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
CB_CONTRAT.Enabled:=False ;
if TContaBonCB_DEJAGENERE.AsInteger<=0 then
   BEGIN
   CB_DATECONTRAT.Color:=clWindow  ; CB_DATECONTRAT.Enabled:=True  ;
   CB_QUALIFPIECE.Color:=clWindow  ; CB_QUALIFPIECE.Enabled:=True  ;
   END else
   BEGIN
   CB_DATECONTRAT.Color:=clBtnFace ; CB_DATECONTRAT.Enabled:=False ;
   CB_QUALIFPIECE.Color:=clBtnFace ; CB_QUALIFPIECE.Enabled:=False ;
   END ;
CB_NBREPETITION.MinValue:=TContaBonCB_DEJAGENERE.AsInteger ;
END ;

Procedure TFContaBon.NewEnreg ;
BEGIN
InitNew(TContaBon) ;
CB_CONTRAT.Enabled:=True ; CB_CONTRAT.SetFocus ;
if EstCompta then TContaBonCB_COMPTABLE.AsString:='X' else TContaBonCB_COMPTABLE.AsString:='-' ;
TContaBonCB_RECONDUCTION.AsString:='TAC' ;
TContaBonCB_SEPAREPAR.AsString:='1M' ;
TContaBonCB_ARRONDI.AsString:='PAS' ;
TContaBonCB_QUALIFPIECE.AsString:='N' ;
END ;

Function TFContaBon.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier:=False ; NextControl(Self) ;
if TContaBon.Modified then
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

Function TFContaBon.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
Modifier:=True ; NextControl(Self) ;
if TContaBon.state in [dsEdit,dsInsert]=False then Exit ;
if TContaBon.state in [dsEdit,dsInsert] then
   BEGIN
   if TContaBonCB_CONTRAT.AsString='' then BEGIN MsgBox.Execute(2,'','') ; CB_CONTRAT.SetFocus ; Exit ; END ;
   if TContaBonCB_LIBELLE.AsString='' then BEGIN MsgBox.Execute(3,'','') ; CB_LIBELLE.SetFocus ; Exit ; END ;
   if TContaBonCB_GUIDE.AsString='' then BEGIN MsgBox.Execute(6,'','') ; CB_GUIDE.SetFocus ; Exit ; END ;
   END ;
if TContaBon.state in [dsInsert] then
   BEGIN
   if Not VerifisiExiste then BEGIN MsgBox.Execute(4,'','') ; CB_CONTRAT.SetFocus ; Exit ; END ;
   if TContaBonCB_DATECONTRAT.AsDateTime<VH^.Encours.Deb then BEGIN MsgBox.Execute(9,'','') ; CB_DATECONTRAT.SetFocus ; Exit ; END ;
//   TContaBonCB_DATEDERNGENERE.AsDateTime:=TContaBonCB_DATECONTRAT.AsDateTime ;
   TContaBonCB_DATEDERNGENERE.AsDateTime:=iDate1900;
   END ;
if TContaBon.state in [dsEdit] then
   BEGIN
   if TContaBonCB_DEJAGENERE.AsInteger<=0 then TContaBonCB_DATEDERNGENERE.AsDateTime:=TContaBonCB_DATECONTRAT.AsDateTime ;
   END ;
TContaBonCB_DATEMODIF.AsDateTime:=Date ;
Modifier:=False ;
result:=TRUE  ;
END ;

Function TFContaBon.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
     nblast,nbprior,nbnext,nbfirst,
     nbinsert : if Not OnSauve then Exit ;
     nbPost   : if Not EnregOK then Exit ;
     nbDelete : if Not Supprime then Exit ;
    end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[5]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFContaBon.Supprime:Boolean ;
BEGIN
Result:=(MsgBox.execute(1,'','')=mrYes) AND (Not TContaBon.EOF) ;
END ;

Function TFContaBon.VerifiSiExiste : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CB_COMPTABLE,CB_CONTRAT from CONTABON Where CB_COMPTABLE="'+TContaBonCB_COMPTABLE.AsString+'" '+
           'And CB_CONTRAT="'+TContaBonCB_CONTRAT.AsString+'"',True) ;
Result:=(Q.EOF) ; Ferme(Q) ;
END ;

procedure TFContaBon.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if(ssCtrl in Shift)And(Key=VK_DELETE)And(TContaBon.Eof) then BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)And(Key=VK_DELETE)then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFContaBon.CB_SEPAREPARChange(Sender: TObject);
begin
if (CB_SEPAREPAR.Value='SEM') or (CB_SEPAREPAR.Value='QUI') then
  BEGIN
  CB_ARRONDI.Value:='PAS' ; CB_ARRONDI.Enabled:=False ;
  END else CB_ARRONDI.Enabled:=True ;
end;

procedure TFContaBon.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

procedure TFContaBon.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then
   BEGIN
   Action:=caFree ; if OkRes then _Bloqueur('nrBatch',False) ;
   END ;
end;

end.
