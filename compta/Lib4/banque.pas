unit banque;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Forms, Dialogs, DBCtrls, StdCtrls, Buttons, ExtCtrls, Spin, HDB, Mask,
     Hctrls, DBTables, DB, Grids, DBGrids, TabNotBk, CodePost, ComCtrls,
     hmsgbox, ENT1, hCompte, HEnt1, Hqry, HSysMenu, ParamDBG,majtable, HPanel, UiUtil,
  HTB97 ;

Procedure FicheBanque(CodeBanque : String ; Comment : TActionFiche ; QuellePage : Integer) ;

type
  TFbanque = class(TForm)
    SBanque  : TDataSource;
    TBanque: THTable;
    DBNav    : TDBNavigator;
    BAide: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BValider: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BFerme: TToolbarButton97;
    HPB: TToolWindow97;
    BPages      : TPageControl;
    PIdent      : TTabSheet;
    TPQ_BANQUE  : THLabel;
    TPQ_LIBELLE : THLabel;
    PQ_BANQUE   : TDBEdit;
    PQ_LIBELLE  : TDBEdit;
    PCondition  : TTabSheet;
    MRemise     : TGroupBox;
    MRestitu    : TGroupBox;
    MpLib       : TGroupBox;
    PQ_REMCB    : THDBValComboBox;
    PQ_RESTCB   : THDBValComboBox;
    TCB         : THLabel;
    MsgBox      : THMsgBox;
    PQ_REMCHQ  : THDBValComboBox;
    PQ_REMESP  : THDBValComboBox;
    PQ_REMLCR  : THDBValComboBox;
    PQ_REMPRE  : THDBValComboBox;
    PQ_REMTEP  : THDBValComboBox;
    PQ_REMTIP  : THDBValComboBox;
    PQ_REMTRI  : THDBValComboBox;
    PQ_REMVIC  : THDBValComboBox;
    PQ_REMVIR  : THDBValComboBox;
    PQ_REMVIT  : THDBValComboBox;
    PQ_RESTCHQ : THDBValComboBox;
    PQ_RESTESP : THDBValComboBox;
    PQ_RESTLCR : THDBValComboBox;
    PQ_RESTPRE : THDBValComboBox;
    PQ_RESTTEP : THDBValComboBox;
    PQ_RESTTIP : THDBValComboBox;
    PQ_RESTTRI : THDBValComboBox;
    PQ_RESTVIC : THDBValComboBox;
    PQ_RESTVIR : THDBValComboBox;
    PQ_RESTVIT : THDBValComboBox;
    TCH   : THLabel;
    TESP  : THLabel;
    TLCR  : THLabel;
    TPRE  : THLabel;
    TTEP  : THLabel;
    TTIP  : THLabel;
    TTRI  : THLabel;
    TVIC  : THLabel;
    TVIR  : THLabel;
    TVIT  : THLabel;
    GbDecouvert : TGroupBox;
    TPQ_DE_MODE : THLabel;
    PQ_DE_MODE  : THDBValComboBox;
    PQ_DE_DATETAUX     : TDBEdit;
    TPQ_DE_DATETAUX    : THLabel;
    TPQ_DE_DENOMINATEUR: THLabel;
    PQ_DE_TAUX         : TDBEdit;
    PQ_DE_PLAFOND      : TDBEdit;
    TPQ_DE_PLAFOND     : THLabel;
    TPQ_DE_NUMERATEUR  : THLabel;
    GbCommdecouv       : TGroupBox;
    TPQ_PD_MODE        : THLabel;
    PQ_PD_MODE         : THDBValComboBox;
    PQ_PD_TXREF        : THDBValComboBox;
    PQ_PD_SAISIE       : TDBEdit;
    PQ_PD_CORRECTION   : TDBEdit;
    PQ_PD_DATETAUX     : TDBEdit;
    TPQ_PD_DATETAUX    : THLabel;
    PQ_PD_DENOMINATEUR : TDBEdit;
    TPQ_PD_DENOMINATEUR: THLabel;
    PQ_PD_TAUX         : TDBEdit;
    PQ_PD_PLAFOND      : TDBEdit;
    TPQ_PD_PLAFOND     : THLabel;
    TPQ_PD_NUMERATEUR  : THLabel;
    PQ_PD_NUMERATEUR   : TDBEdit;
    GbCredit           : TGroupBox;
    TPQ_CR_MODE        : THLabel;
    PQ_CR_MODE         : THDBValComboBox;
    PQ_CR_TXREF        : THDBValComboBox;
    PQ_CR_SAISIE       : TDBEdit;
    PQ_CR_CORRECTION   : TDBEdit;
    PQ_CR_DATETAUX     : TDBEdit;
    PQ_CR_TYPEPLAFOND  : THDBValComboBox;
    TPQ_CR_TYPEPLAFOND : THLabel;
    TPQ_CR_DATETAUX    : THLabel;
    PQ_CR_DENOMINATEUR : TDBEdit;
    PQ_CR_TAUX         : TDBEdit;
    TPQ_CR_DENOMINATEUR: THLabel;
    PQ_CR_NUMERATEUR   : TDBEdit;
    PQ_CR_PLAFOND      : TDBEdit;
    TPQ_CR_PLAFOND     : THLabel;
    TPQ_CR_NUMERATEUR  : THLabel;
    GbCommcpte         : TGroupBox;
    TPQ_CO_MODE        : THLabel;
    PQ_CO_MODE         : THDBValComboBox;
    PQ_CO_TXREF        : THDBValComboBox;
    PQ_CO_SAISIE       : TDBEdit;
    PQ_CO_CORRECTION   : TDBEdit;
    PQ_CO_DATETAUX     : TDBEdit;
    TPQ_CO_DATETAUX    : THLabel;
    PQ_CO_DENOMINATEUR : TDBEdit;
    TPQ_CO_DENOMINATEUR: THLabel;
    PQ_CO_TAUX         : TDBEdit;
    PQ_CO_NUMERATEUR   : TDBEdit;
    TPQ_CO_NUMERATEUR  : THLabel;
    PQ_DE_SAISIE       : TDBEdit;
    PQ_DE_NUMERATEUR   : TDBEdit;
    PQ_DE_DENOMINATEUR : TDBEdit;
    PQ_DE_TXREF        : THDBValComboBox;
    PQ_DE_CORRECTION   : TDBEdit;
    TTPQ_PD_PLAFOND    : THLabel;
    TPQ_CO_PLAFOND     : THLabel;
    PQ_CO_FRAIS        : TDBEdit;
    Gbtaux             : TGroupBox;
    TPQ_BB_TAUX        : THLabel;
    TPQ_BB_DATETAUX    : THLabel;
    PQ_BB_TAUX         : TDBEdit;
    PQ_BB_DATETAUX     : TDBEdit;
    TBanquePQ_BANQUE          : TStringField;
    TBanquePQ_LIBELLE         : TStringField;
    TBanquePQ_RESTCB          : TStringField;
    TBanquePQ_REMCB           : TStringField;
    TBanquePQ_RESTCHQ         : TStringField;
    TBanquePQ_REMCHQ          : TStringField;
    TBanquePQ_RESTESP         : TStringField;
    TBanquePQ_REMESP          : TStringField;
    TBanquePQ_RESTLCR         : TStringField;
    TBanquePQ_REMLCR          : TStringField;
    TBanquePQ_RESTPRE         : TStringField;
    TBanquePQ_REMPRE          : TStringField;
    TBanquePQ_RESTTRI         : TStringField;
    TBanquePQ_REMTRI          : TStringField;
    TBanquePQ_RESTVIR         : TStringField;
    TBanquePQ_REMVIR          : TStringField;
    TBanquePQ_RESTVIT         : TStringField;
    TBanquePQ_REMVIT          : TStringField;
    TBanquePQ_RESTTIP         : TStringField;
    TBanquePQ_REMTIP          : TStringField;
    TBanquePQ_RESTTEP         : TStringField;
    TBanquePQ_REMTEP          : TStringField;
    TBanquePQ_RESTVIC         : TStringField;
    TBanquePQ_REMVIC          : TStringField;
    TBanquePQ_DE_MODE         : TStringField;
    TBanquePQ_DE_TXREF        : TStringField;
    TBanquePQ_DE_CORRECTION   : TFloatField;
    TBanquePQ_DE_SAISIE       : TFloatField;
    TBanquePQ_DE_NUMERATEUR: TIntegerField;
    TBanquePQ_DE_DENOMINATEUR: TIntegerField;
    TBanquePQ_DE_TAUX         : TFloatField;
    TBanquePQ_DE_DATETAUX     : TDateTimeField;
    TBanquePQ_DE_PLAFOND      : TFloatField;
    TBanquePQ_PD_MODE         : TStringField;
    TBanquePQ_PD_TXREF        : TStringField;
    TBanquePQ_PD_CORRECTION   : TFloatField;
    TBanquePQ_PD_SAISIE       : TFloatField;
    TBanquePQ_PD_NUMERATEUR: TIntegerField;
    TBanquePQ_PD_DENOMINATEUR: TIntegerField;
    TBanquePQ_PD_TAUX         : TFloatField;
    TBanquePQ_PD_DATETAUX     : TDateTimeField;
    TBanquePQ_PD_PLAFOND      : TFloatField;
    TBanquePQ_CR_MODE         : TStringField;
    TBanquePQ_CR_TXREF        : TStringField;
    TBanquePQ_CR_CORRECTION   : TFloatField;
    TBanquePQ_CR_SAISIE       : TFloatField;
    TBanquePQ_CR_NUMERATEUR: TIntegerField;
    TBanquePQ_CR_DENOMINATEUR: TIntegerField;
    TBanquePQ_CR_TAUX         : TFloatField;
    TBanquePQ_CR_DATETAUX     : TDateTimeField;
    TBanquePQ_CR_PLAFOND      : TFloatField;
    TBanquePQ_CR_TYPEPLAFOND  : TStringField;
    TBanquePQ_CO_MODE         : TStringField;
    TBanquePQ_CO_TXREF        : TStringField;
    TBanquePQ_CO_CORRECTION   : TFloatField;
    TBanquePQ_CO_SAISIE       : TFloatField;
    TBanquePQ_CO_NUMERATEUR: TIntegerField;
    TBanquePQ_CO_DENOMINATEUR: TIntegerField;
    TBanquePQ_CO_TAUX         : TFloatField;
    TBanquePQ_CO_DATETAUX     : TDateTimeField;
    TBanquePQ_CO_FRAIS        : TFloatField;
    TBanquePQ_BB_TAUX         : TFloatField;
    TBanquePQ_BB_DATETAUX     : TDateTimeField;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    Dock: TDock97;
    procedure FormShow(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure SBanqueDataChange(Sender: TObject; Field: TField);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BAnnulerClick(Sender: TObject);
    procedure PQ_DE_MODEChange(Sender: TObject);
    procedure PQ_PD_MODEChange(Sender: TObject);
    procedure PQ_CR_MODEChange(Sender: TObject);
    procedure PQ_CO_MODEChange(Sender: TObject);
    procedure PQ_BB_TAUXClick(Sender: TObject);
    procedure PQ_DE_SAISIEClick(Sender: TObject);
    procedure PQ_DE_NUMERATEURClick(Sender: TObject);
    procedure PQ_DE_CORRECTIONClick(Sender: TObject);
    procedure PQ_DE_DENOMINATEURClick(Sender: TObject);
    procedure PQ_DE_PLAFONDClick(Sender: TObject);
    procedure PQ_PD_SAISIEClick(Sender: TObject);
    procedure PQ_PD_NUMERATEURClick(Sender: TObject);
    procedure PQ_PD_CORRECTIONClick(Sender: TObject);
    procedure PQ_PD_DENOMINATEURClick(Sender: TObject);
    procedure PQ_PD_PLAFONDClick(Sender: TObject);
    procedure PQ_CR_SAISIEClick(Sender: TObject);
    procedure PQ_CR_NUMERATEURClick(Sender: TObject);
    procedure PQ_CR_CORRECTIONClick(Sender: TObject);
    procedure PQ_CR_DENOMINATEURClick(Sender: TObject);
    procedure PQ_CR_PLAFONDClick(Sender: TObject);
    procedure PQ_CO_SAISIEClick(Sender: TObject);
    procedure PQ_CO_NUMERATEURClick(Sender: TObject);
    procedure PQ_CO_CORRECTIONClick(Sender: TObject);
    procedure PQ_CO_DENOMINATEURClick(Sender: TObject);
    procedure PQ_CO_FRAISClick(Sender: TObject);
    procedure PQ_BB_TAUXChange(Sender: TObject);
    procedure PQ_BB_DATETAUXChange(Sender: TObject);
    procedure PQ_BB_DATETAUXKeyPress(Sender: TObject; var Key: Char);
    procedure PQ_DE_DATETAUXKeyPress(Sender: TObject; var Key: Char);
    procedure PQ_PD_DATETAUXKeyPress(Sender: TObject; var Key: Char);
    procedure PQ_CR_DATETAUXKeyPress(Sender: TObject; var Key: Char);
    procedure PQ_CO_DATETAUXKeyPress(Sender: TObject; var Key: Char);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
  private
    LeQuel : String ;
    Mode   : TActionFiche ;
    LaPage : Integer ;
    Modifier : Boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Procedure CurseurAuLibelle ;
    Procedure CurseurAuCode ;
    Function  EnregOK : boolean ;
    Procedure ChargeEnreg ;
    Function  ExisteCode:Boolean ;
    Function  LePrefixe(Sender : TObject):String ;
    Procedure CacheLesChamps(Sender : TObject) ;
    Procedure Initialisation ;
    Function  DenominateurOk(Sender : TObject):Boolean ;
    Procedure RempliTaux(Sender : TObject) ;
    Procedure TbbDateModifier ;
    Function  Supprime : Boolean ;
  public

  end;


implementation

{$R *.DFM}

Uses ParamDat, PrintDBG ;

procedure TFbanque.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TBanque.Open ; Modifier:=False ;
BPages.ActivePage:=BPages.Pages[LaPage] ;
if Mode=taConsult then FicheReadOnly(Self) ;
if Mode in [taCreat..taCreatOne] then Bouge(nbInsert) ;
if LeQuel<>'' then
  BEGIN
  if Not TBanque.FindKey([LeQuel]) then
     BEGIN
     MessageAlerte(MsgBox.Mess[6]) ;
     if Not V_PGI.OutLook then BEGIN PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; END ;
     END ; 
  END ;
if TBanque.Eof then Bouge(nbInsert) ;
end;

procedure TFbanque.SBanqueDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not TBanque.Modified ; BDelete.Enabled:=Not TBanque.Modified ;
if(TBanque.Eof)And(TBanque.Bof) then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TBanque.BOF;
   DnEnable := Enabled and not TBanque.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
//   if ((Field.FieldName='BQ_LIBELLE') and (BQ_ABREGE.Field.AsString='')) then
//      BQ_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

Function TFbanque.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[6]) ;
result:=TRUE ; TbbDateModifier ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFbanque.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TBanque.Modified then
   BEGIN
   if Mode in [taCreat..taCreatOne]then Rep:=mrNo
   else if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFbanque.EnregOK : boolean ;
BEGIN
Result:=FALSE  ;
if TBanque.state in [dsEdit,dsInsert]=False then Exit ;
if TBanque.state in [dsEdit,dsInsert] then
   BEGIN
   if TBanquePQ_BANQUE.AsString='' then
      BEGIN CurseurAuCode ; MsgBox.Execute(2,'','') ;  Exit ; END ;
   if TBanquePQ_LIBELLE.AsString='' then
      BEGIN CurseurAuLibelle ; MsgBox.Execute(3,'','') ;  Exit ; END ;
   if Not DenominateurOk(PQ_DE_DENOMINATEUR) then Exit ;
   if Not DenominateurOk(PQ_PD_DENOMINATEUR) then Exit ;
   if Not DenominateurOk(PQ_CR_DENOMINATEUR) then Exit ;
   if Not DenominateurOk(PQ_CO_DENOMINATEUR) then Exit ;
   END ;
if TBanque.state in [dsInsert] then
   BEGIN
   if ExisteCode then
      BEGIN CurseurAuCode ; MsgBox.Execute(4,'','') ;  Exit ; END ;
   END ;
RempliTaux(PQ_DE_MODE) ; RempliTaux(PQ_PD_MODE) ;
RempliTaux(PQ_CR_MODE) ; RempliTaux(PQ_CO_MODE) ;
Result:=TRUE  ;
END ;

Function TFbanque.ExisteCode : Boolean ;
BEGIN Result:=Presence('BANQUES','PQ_BANQUE',PQ_BANQUE.Text) ; END ;

Procedure TFbanque.ChargeEnreg ;
BEGIN
InitCaption(Self,PQ_BANQUE.text,PQ_LIBELLE.text) ;
PQ_BANQUE.Enabled:=False ;
CacheLesChamps(PQ_DE_MODE) ; CacheLesChamps(PQ_PD_MODE) ;
CacheLesChamps(PQ_CR_MODE) ; CacheLesChamps(PQ_CO_MODE) ;
END ;

Procedure TFbanque.NewEnreg ;
BEGIN
InitNew(TBanque) ; Initialisation ; PQ_BANQUE.Enabled:=True ; CurseurAuCode ;
END ;

Function TFbanque.Supprime : Boolean ;
Var Q : TQuery ;
BEGIN
Result:=False ;
if MsgBox.execute(1,'','')<>mrYes then Exit ;
Q:=OpenSql('Select BQ_GENERAL from BANQUECP Where BQ_BANQUE="'+TBanquePQ_BANQUE.AsString+'"',True) ;
if Not Q.Eof then BEGIN MsgBox.Execute(7,'','') ; Ferme(Q) ; Exit ; END ;
Ferme(Q) ; Result:=True ;
END ;

procedure TFbanque.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFbanque.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFbanque.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFbanque.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFbanque.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFbanque.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFbanque.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFbanque.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost)then
   BEGIN
   if Mode=taCreatEnSerie then Bouge(nbInsert) ;
   if Mode=taCreatOne then Close ;
   END ;
end;

procedure TFbanque.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFbanque.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

Procedure TFbanque.CurseurAuLibelle ;
BEGIN BPages.ActivePage:=PIdent ; PQ_LIBELLE.SetFocus ; END ;

Procedure TFbanque.CurseurAuCode ;
BEGIN BPages.ActivePage:=PIdent ; PQ_BANQUE.SetFocus ; END ;

procedure TFbanque.PQ_DE_MODEChange(Sender: TObject);
begin CacheLesChamps(Sender) ;end;

procedure TFbanque.PQ_PD_MODEChange(Sender: TObject);
begin CacheLesChamps(Sender) ; end;

procedure TFbanque.PQ_CR_MODEChange(Sender: TObject);
begin CacheLesChamps(Sender) ; end;

procedure TFbanque.PQ_CO_MODEChange(Sender: TObject);
begin CacheLesChamps(Sender) ; end;

procedure TFbanque.PQ_BB_TAUXClick(Sender: TObject);
begin PQ_BB_TAUX.SelectAll; end;

procedure TFbanque.PQ_DE_SAISIEClick(Sender: TObject);
begin PQ_DE_SAISIE.SelectAll ; end;

procedure TFbanque.PQ_DE_NUMERATEURClick(Sender: TObject);
begin PQ_DE_NUMERATEUR.SelectAll ; end;

procedure TFbanque.PQ_DE_CORRECTIONClick(Sender: TObject);
begin PQ_DE_CORRECTION.SelectAll ; end;

procedure TFbanque.PQ_DE_DENOMINATEURClick(Sender: TObject);
begin PQ_DE_DENOMINATEUR.SelectAll ; end;

procedure TFbanque.PQ_DE_PLAFONDClick(Sender: TObject);
begin PQ_DE_PLAFOND.SelectAll ; end;

procedure TFbanque.PQ_PD_SAISIEClick(Sender: TObject);
begin PQ_PD_SAISIE.SelectAll ; end;

procedure TFbanque.PQ_PD_NUMERATEURClick(Sender: TObject);
begin PQ_PD_NUMERATEUR.SelectAll ; end;

procedure TFbanque.PQ_PD_CORRECTIONClick(Sender: TObject);
begin PQ_PD_CORRECTION.SelectAll ; end;

procedure TFbanque.PQ_PD_DENOMINATEURClick(Sender: TObject);
begin PQ_PD_DENOMINATEUR.SelectAll ; end;

procedure TFbanque.PQ_PD_PLAFONDClick(Sender: TObject);
begin PQ_PD_PLAFOND.SelectAll ; end;

procedure TFbanque.PQ_CR_SAISIEClick(Sender: TObject);
begin PQ_CR_SAISIE.SelectAll ; end;

procedure TFbanque.PQ_CR_NUMERATEURClick(Sender: TObject);
begin PQ_CR_NUMERATEUR.SelectAll ; end;

procedure TFbanque.PQ_CR_CORRECTIONClick(Sender: TObject);
begin PQ_CR_CORRECTION.SelectAll ; end;

procedure TFbanque.PQ_CR_DENOMINATEURClick(Sender: TObject);
begin PQ_CR_DENOMINATEUR.SelectAll ; end;

procedure TFbanque.PQ_CR_PLAFONDClick(Sender: TObject);
begin PQ_CR_PLAFOND.SelectAll ; end;

procedure TFbanque.PQ_CO_SAISIEClick(Sender: TObject);
begin PQ_CO_SAISIE.SelectAll ; end;

procedure TFbanque.PQ_CO_NUMERATEURClick(Sender: TObject);
begin PQ_CO_NUMERATEUR.SelectAll ; end;

procedure TFbanque.PQ_CO_CORRECTIONClick(Sender: TObject);
begin PQ_CO_CORRECTION.SelectAll ; end;

procedure TFbanque.PQ_CO_DENOMINATEURClick(Sender: TObject);
begin PQ_CO_DENOMINATEUR.SelectAll ; end;

procedure TFbanque.PQ_CO_FRAISClick(Sender: TObject);
begin PQ_CO_FRAIS.SelectAll ; end;

procedure TFbanque.PQ_BB_TAUXChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFbanque.PQ_BB_DATETAUXChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFbanque.PQ_BB_DATETAUXKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFbanque.PQ_DE_DATETAUXKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFbanque.PQ_PD_DATETAUXKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFbanque.PQ_CR_DATETAUXKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFbanque.PQ_CO_DATETAUXKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

Function TFbanque.LePrefixe(Sender : TObject):String ;
BEGIN
if THDBValComboBox(Sender).Name='PQ_DE_MODE' then Result:='DE' else
if THDBValComboBox(Sender).Name='PQ_PD_MODE' then Result:='PD' else
if THDBValComboBox(Sender).Name='PQ_CR_MODE' then Result:='CR' else
if THDBValComboBox(Sender).Name='PQ_CO_MODE' then Result:='CO' ;
END ;

Procedure TFbanque.CacheLesChamps(Sender : TObject) ;
Var Pref : String ;
BEGIN
Pref:=LePrefixe(Sender) ;
if THDBValComboBox(Sender).Value='TX' then
   BEGIN
   TControl(FindComponent('PQ_'+Pref+'_SAISIE')).Visible:=True ;
   TControl(FindComponent('PQ_'+Pref+'_TXREF')).Visible:=False ;
   TControl(FindComponent('PQ_'+Pref+'_NUMERATEUR')).Visible:=False ;
   TControl(FindComponent('TPQ_'+Pref+'_NUMERATEUR')).Visible:=False ;
   TControl(FindComponent('TPQ_'+Pref+'_DENOMINATEUR')).Visible:=False ;
   TControl(FindComponent('PQ_'+Pref+'_DENOMINATEUR')).Visible:=False ;
   TControl(FindComponent('PQ_'+Pref+'_CORRECTION')).Visible:=False ;
   if TBanque.state in [dsEdit,dsInsert] then
      BEGIN
       TDBEdit(FindComponent('PQ_'+Pref+'_NUMERATEUR')).Field.AsFloat:=0 ;
       TDBEdit(FindComponent('PQ_'+Pref+'_DENOMINATEUR')).Field.AsFloat:=0 ;
       TDBEdit(FindComponent('PQ_'+Pref+'_CORRECTION')).Field.AsFloat:=0 ;
      END ;
   END else
if THDBValComboBox(Sender).Value='TF' then
   BEGIN
   TControl(FindComponent('PQ_'+Pref+'_SAISIE')).Visible:=True ;
   TControl(FindComponent('PQ_'+Pref+'_NUMERATEUR')).Visible:=True ;
   TControl(FindComponent('TPQ_'+Pref+'_NUMERATEUR')).Visible:=True ;
   TControl(FindComponent('TPQ_'+Pref+'_DENOMINATEUR')).Visible:=True ;
   TControl(FindComponent('PQ_'+Pref+'_DENOMINATEUR')).Visible:=True ;
   TControl(FindComponent('PQ_'+Pref+'_TXREF')).Visible:=False ;
   TControl(FindComponent('PQ_'+Pref+'_CORRECTION')).Visible:=False ;
   if TBanque.state in [dsEdit,dsInsert] then
      TDBEdit(FindComponent('PQ_'+Pref+'_CORRECTION')).Field.AsFloat:=0 ;
   END else
if THDBValComboBox(Sender).Value='TT' then
   BEGIN
   TControl(FindComponent('PQ_'+Pref+'_TXREF')).Visible:=True ;
   TControl(FindComponent('PQ_'+Pref+'_CORRECTION')).Visible:=True ;
   TControl(FindComponent('PQ_'+Pref+'_SAISIE')).Visible:=False ;
   TControl(FindComponent('PQ_'+Pref+'_NUMERATEUR')).Visible:=False ;
   TControl(FindComponent('TPQ_'+Pref+'_NUMERATEUR')).Visible:=False ;
   TControl(FindComponent('TPQ_'+Pref+'_DENOMINATEUR')).Visible:=False ;
   TControl(FindComponent('PQ_'+Pref+'_DENOMINATEUR')).Visible:=False ;
   if TBanque.state in [dsEdit,dsInsert] then
      BEGIN
       TDBEdit(FindComponent('PQ_'+Pref+'_NUMERATEUR')).Field.AsFloat:=0 ;
       TDBEdit(FindComponent('PQ_'+Pref+'_DENOMINATEUR')).Field.AsFloat:=0 ;
       TDBEdit(FindComponent('PQ_'+Pref+'_SAISIE')).Field.AsFloat:=0 ;
      END ;
   END ;
END ;

Procedure TFbanque.Initialisation ;
Var i : Integer ;
    CH : THDbValComboBox ;
BEGIN
PQ_DE_MODE.Value:='TX' ; PQ_PD_MODE.Value:='TX' ;
PQ_CR_MODE.Value:='TX' ; PQ_CO_MODE.Value:='TX' ;
PQ_DE_TXREF.Value:='TBB' ; PQ_PD_TXREF.Value:='TBB' ;
PQ_CR_TXREF.Value:='TBB' ; PQ_CO_TXREF.Value:='TBB' ;
TBanquePQ_BB_DATETAUX.AsDateTime:=Date ; TBanquePQ_DE_DATETAUX.AsDateTime:=Date ;
TBanquePQ_PD_DATETAUX.AsDateTime:=Date ; TBanquePQ_CR_DATETAUX.AsDateTime:=Date ;
TBanquePQ_CO_DATETAUX.AsDateTime:=Date ; PQ_CR_TYPEPLAFOND.ItemIndex:=0 ;
for i:=0 to MRemise.ControlCount-1 do
   if MRemise.Controls[i] is THDBValComboBox then
   BEGIN
   CH:=THDBValComboBox(MRemise.Controls[i]) ;
   if CH.Values.Count>0 then CH.Value:=CH.Values[0] ;
   END ;
for i:=0 to MRestitu.ControlCount-1 do
    if MRestitu.Controls[i] is THDBValComboBox then
    BEGIN
    CH:=THDBValComboBox(MRestitu.Controls[i]) ;
    if CH.values.Count>0 then CH.Value:=CH.Values[0] ;
    END ;
END ;

Function TFbanque.DenominateurOk(Sender : TObject):Boolean ;
BEGIN
Result:=True ;
if Not TDBEdit(Sender).Visible then Exit ;
if TDBEdit(Sender).Text<>'0' then Exit
else BEGIN
     BPages.ActivePage:=PCondition ; MsgBox.Execute(5,'','') ;
     Result:=False ; TDBEdit(Sender).SetFocus ;
     END ;
END ;

Procedure TFbanque.RempliTaux(Sender : TObject) ;
Var Pref,Tau : String ;
    Sai,Num,Den,Cor,Tbb : Extended ;
    Q : TQuery ;
BEGIN
Pref:=LePrefixe(Sender) ;
Sai:=Tbanque.FindField('PQ_'+Pref+'_SAISIE').AsFloat ;
Num:=Tbanque.FindField('PQ_'+Pref+'_NUMERATEUR').AsFloat ;
Den:=Tbanque.FindField('PQ_'+Pref+'_DENOMINATEUR').AsFloat ;
Cor:=Tbanque.FindField('PQ_'+Pref+'_CORRECTION').AsFloat ;
if THDBValComboBox(Sender).Value='TX' then
  BEGIN TDBEdit(FindComponent('PQ_'+Pref+'_TAUX')).Field.AsFloat:=Sai ; END ;
if THDBValComboBox(Sender).Value='TF' then
  BEGIN TDBEdit(FindComponent('PQ_'+Pref+'_TAUX')).Field.AsFloat:=Sai+(Num/Den) ; END ;
if THDBValComboBox(Sender).Value='TT' then
  BEGIN
  Tau:=THDBValComboBox(FindComponent('PQ_'+Pref+'_TXREF')).Value ;
  Q:=OpenSql('Select BT_TAUX from BQTAUX Where BT_TYPETAUX="'+Tau+'" And '+
             'BT_DATE<="'+UsDateTime(TBanquePQ_BB_DATETAUX.AsDateTime)+'" order by BT_DATE DESC',True) ;
  Tbb:=Q.Fields[0].AsFloat ; Ferme(Q) ;
  TDBEdit(FindComponent('PQ_'+Pref+'_TAUX')).Field.AsFloat:=Tbb+Cor ;
  END ;
END ;

Procedure TFbanque.TbbDateModifier ;
Var Q : TQuery ;
BEGIN
if Not Modifier then Exit ;
Q:=OpenSql('Select BT_DATE From BQTAUX Where BT_DATE="'+UsDateTime(TBanquePQ_BB_DATETAUX.AsDateTime)+'" '+
           'And BT_TYPETAUX="TBB" And BT_BANQUE="'+TBanquePQ_BANQUE.AsString+'"',True) ;
if Q.Eof then
  BEGIN
  ExecuteSql('INSERT INTO BQTAUX (BT_BANQUE, BT_TYPETAUX, BT_DATE, BT_TAUX) '+
             'VALUES ("'+TBanquePQ_BANQUE.AsString+'", "TBB", "'+UsDateTime(TBanquePQ_BB_DATETAUX.AsDateTime)+'", '+
             ''+StrFPoint(TBanquePQ_BB_TAUX.AsFloat)+')') ;
  END else
  ExecuteSql('UPDATE BQTAUX SET BT_DATE="'+UsDateTime(TBanquePQ_BB_DATETAUX.AsDateTime)+'", '+
             'BT_TAUX='+StrFPoint(TBanquePQ_BB_TAUX.AsFloat)+' Where '+
             'BT_BANQUE="'+TBanquePQ_BANQUE.AsString+'" '+
             'And BT_DATE="'+UsDateTime(TBanquePQ_BB_DATETAUX.AsDateTime)+'"') ;
Ferme(Q) ;
END ;

Procedure FicheBanque(CodeBanque : String ; Comment : TActionFiche ; QuellePage : Integer) ;
var FBanque: TFBanque;
    PP : THPanel ;
begin
if Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FBanque:=TFBanque.Create(Application) ;
FBanque.LeQuel:=CodeBanque ; FBanque.Mode:=Comment ; FBanque.LaPage:=QuellePage ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FBanque.ShowModal ;
    finally
     FBanque.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FBanque,PP) ;
   FBanque.Show ;
   END ;
end ;

procedure TFbanque.FormClose(Sender: TObject; var Action: TCloseAction);
begin
AvertirTable('ttBanque') ;
if Parent is THPanel then Action:=caFree ;
END;

procedure TFbanque.BImprimerClick(Sender: TObject);
begin PrintDBGrid(Nil,Nil,Caption,'PRT_BQE') ; end;

procedure TFbanque.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

end.
