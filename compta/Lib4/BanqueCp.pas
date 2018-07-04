unit banqueCP;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Dialogs, DBCtrls, StdCtrls, Buttons, ExtCtrls, Spin, HDB, Mask,
     Hctrls, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB, Grids, DBGrids, TabNotBk, CodePost, ComCtrls,
     hmsgbox, ENT1, hCompte, HEnt1, HRichEdt, HSysMenu, HRichOLE, Region,MajTable,
     HTB97, Hqry, HRegCpte, UtilPGI, 
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCS3}
     EtbUser,
{$ENDIF}
{$ENDIF}
{$ENDIF}
     Forms, ADODB, TntComCtrls, TntStdCtrls, TntDBCtrls, TntButtons;

Procedure FicheBanqueCP(CpteGeneral : String ; Comment : TActionFiche ; QuellePage : Integer) ;

type
  TFbanqueCP = class(TForm)
    SBanqueCP: TDataSource;
    TBanqueCP: THTable;
    MsgBox   : THMsgBox;
    DBNav    : TDBNavigator;
    BAide: THBitBtn;
    BAnnuler: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    BFirst: THBitBtn;
    BPrev: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    BInsert: THBitBtn;
    BDelete: THBitBtn;
    BFerme: THBitBtn;
    BMCbanq: TToolbarButton97;
    HPB      : TPanel;
    FAutoSave: TCheckBox;
    Pbanque  : TPanel;
    BPages        : TPageControl;
    PCoordo       : TTabSheet;
    PCommunic     : TTabSheet;
    PInfo         : TTabSheet;
    PDelai        : TTabSheet;
    PIdent        : TTabSheet;
    HGroupBox2    : TGroupBox;
    HGroupBox4    : TGroupBox;
    HGroupBox5    : TGroupBox;
    HGroupBox6    : TGroupBox;
    HGBRIB        : TGroupBox;
    TBQ_CONTACT              : THLabel;
    TBQ_ADRESSE              : THLabel;
    TBQ_CODEPOSTAL           : THLabel;
    TBQ_DIVTERRIT            : THLabel;
    TBQ_TELEPHONE            : THLabel;
    TBQ_FAX                  : THLabel;
    TBQ_VILLE                : THLabel;
    TBQ_TELEX                : THLabel;
    TBQ_REPVIR            : THLabel;
    TBQ_REPPRELEV         : THLabel;
    TBQ_REPLCR            : THLabel;
    TBQ_REPBONAPAYER      : THLabel;
    TBQ_DELAIVIRORD    : THLabel;
    TBQ_DELAIVIRCHAUD        : THLabel;
    TBQ_DELAIVIRBRULANT      : THLabel;
    TBQ_DELAILCR             : THLabel;
    TBQ_DOMICILIATION        : THLabel;
    TBQ_CODEBIC              : THLabel;
    TBQ_DELAIPRELVORD : THLabel;
    TBQ_DELAIPRELVACC     : THLabel;
    TBQ_ADRESSE2             : THLabel;
    TBQ_ADRESSE3             : THLabel;
    TBQ_ETABBQ               : THLabel;
    TBQ_GUICHET              : THLabel;
    TBQ_NUMEROCOMPTE         : THLabel;
    TBQ_CLERIB               : THLabel;
    TBQ_LIBELLE              : THLabel;
    TBQ_COMMENTAIRE          : THLabel;
    TBQ_LANGUE               : THLabel;
    TBQ_BLOCNOTE             : TGroupBox;
    TBQ_NUMEMETLCR           : THLabel;
    BQ_CONTACT               : TDBEdit;
    BQ_ADRESSE1              : TDBEdit;
    BQ_ADRESSE2              : TDBEdit;
    BQ_ADRESSE3              : TDBEdit;
    BQ_CODEPOSTAL            : TDBEdit;
    BQ_DIVTERRIT             : TDBEdit;
    BQ_TELEPHONE             : TDBEdit;
    BQ_VILLE                 : TDBEdit;
    BQ_FAX                   : TDBEdit;
    BQ_TELEX                 : TDBEdit;
    BQ_REPVIR             : TDBEdit;
    BQ_REPPRELEV          : TDBEdit;
    BQ_REPLCR             : TDBEdit;
    BQ_REPBONAPAYER       : TDBEdit;
    BQ_DELAIVIRORD     : THDBSpinEdit;
    BQ_DELAIVIRCHAUD         : THDBSpinEdit;
    BQ_DELAIVIRBRULANT       : THDBSpinEdit;
    BQ_ECHEREPPRELEV         : TDBCheckBox;
    BQ_DELAIPRELVORD  : THDBSpinEdit;
    BQ_DELAIPRELVACC      : THDBSpinEdit;
    BQ_ECHEREPLCR            : TDBCheckBox;
    BQ_DELAILCR              : THDBSpinEdit;
    BQ_LIBELLE               : TDBEdit;
    BQ_DOMICILIATION         : TDBEdit;
    BQ_CODEBIC               : TDBEdit;
    BQ_COMMENTAIRE           : TDBEdit;
    BQ_LANGUE                : THDBValComboBox;
    BQ_BLOCNOTE: THDBRichEditOLE;
    BQ_NUMEMETLCR            : TDBEdit;
    TBQ_BANQUE               : THLabel;
    BQ_BANQUE                : THDBValComboBox;
    GbJfer                   : TGroupBox;
    Cbj1                     : TCheckBox;
    Cbj2                     : TCheckBox;
    Cbj3                     : TCheckBox;
    Cbj4                     : TCheckBox;
    Cbj5                     : TCheckBox;
    Cbj6                     : TCheckBox;
    Cbj7                     : TCheckBox;
    GbTi                     : TGroupBox;
    TBQ_DELAITRANSINT        : THLabel;
    BQ_DELAITRANSINT         : THDBSpinEdit;
    TBQ_COMPTEFRAIS          : THLabel;
    BQ_COMPTEFRAIS           : THDBCpteEdit;
    TBQ_TYPEREMTRANS         : THLabel;
    BQ_TYPEREMTRANS          : THDBValComboBox;
    TBQ_INDREMTRANS          : THLabel;
    BQ_INDREMTRANS           : THDBValComboBox;
    GbSold                   : TGroupBox;
    TBQ_DATEDERNSOLDE        : THLabel;
    BQ_DATEDERNSOLDE         : TDBEdit;
    TBQ_DERNSOLDEFRS         : THLabel;
    TBQ_DERNSOLDEDEV         : THLabel;
    GbModele                 : TGroupBox;
    TBQ_LETTREVIR            : THLabel;
    BQ_LETTREVIR             : THDBValComboBox;
    BQ_LETTREPRELV           : THDBValComboBox;
    TBQ_LETTREPRELV          : THLabel;
    BQ_LETTRELCR             : THDBValComboBox;
    TBQ_LETTRELCR            : THLabel;
    GbBap                    : TGroupBox;
    TBQ_DELAIBAPLCR          : THLabel;
    BQ_RAPPROAUTOLCR         : TDBCheckBox;
    BQ_DELAIBAPLCR           : THDBSpinEdit;
    GbDiv                    : TGroupBox;
    BQ_RELEVEETRANGER        : TDBCheckBox;
    TBQ_NUMEMETVIR           : THLabel;
    TBQ_GUIDECOMPATBLE       : THLabel;
    BQ_GUIDECOMPATBLE        : TDBEdit;
    BQ_RAPPAUTOREL           : TDBCheckBox;
    BQ_NUMEMETVIR            : TDBEdit;
    TBQ_CONVENTIONLCR        : THLabel;
    BQ_CONVENTIONLCR         : TDBEdit;
    TBQ_ENCOURSLCR           : THLabel;
    BQ_ENCOURSLCR            : TDBEdit;
    TBQ_PLAFONDLCR           : THLabel;
    BQ_PLAFONDLCR            : TDBEdit;
    TBQ_PAYS                 : THLabel;
    BQ_PAYS                  : THDBValComboBox;
    BQ_DERNSOLDEFRS          : THNumEdit;
    BQ_DERNSOLDEDEV          : THNumEdit;
    TBQ_CODE                 : THLabel;
    BQ_CODE                  : TDBEdit;
    TBQ_GENERAL              : THLabel;
    TBQ_DEVISE               : THLabel;
    BQ_DEVISE                : THDBValComboBox;
    TEdev                    : TLabel;
    EDev                     : TEdit;
    BQ_GENERAL               : THDBCpteEdit;
    HMTrad                   : THSystemMenu;
    TBQ_LETTRECHQ            : THLabel;
    BQ_LETTRECHQ             : THDBValComboBox;
    TBQ_DESTINATAIRE         : THLabel;
    BQ_DESTINATAIRE          : TDBEdit;
    HLabel1                  : THLabel;
    TBanqueCPBQ_CODE         : TStringField;
    TBanqueCPBQ_GENERAL      : TStringField;
    TBanqueCPBQ_LIBELLE      : TStringField;
    TBanqueCPBQ_DOMICILIATION: TStringField;
    TBanqueCPBQ_ADRESSE1     : TStringField;
    TBanqueCPBQ_ADRESSE2     : TStringField;
    TBanqueCPBQ_ADRESSE3     : TStringField;
    TBanqueCPBQ_CODEPOSTAL   : TStringField;
    TBanqueCPBQ_VILLE        : TStringField;
    TBanqueCPBQ_DIVTERRIT    : TStringField;
    TBanqueCPBQ_PAYS         : TStringField;
    TBanqueCPBQ_LANGUE       : TStringField;
    TBanqueCPBQ_DEVISE       : TStringField;
    TBanqueCPBQ_TELEPHONE    : TStringField;
    TBanqueCPBQ_FAX          : TStringField;
    TBanqueCPBQ_TELEX        : TStringField;
    TBanqueCPBQ_CONTACT      : TStringField;
    TBanqueCPBQ_ETABBQ       : TStringField;
    TBanqueCPBQ_NUMEROCOMPTE : TStringField;
    TBanqueCPBQ_CLERIB       : TStringField;
    TBanqueCPBQ_GUICHET      : TStringField;
    TBanqueCPBQ_CODEBIC      : TStringField;
    TBanqueCPBQ_NUMEMETLCR   : TStringField;
    TBanqueCPBQ_CONVENTIONLCR: TStringField;
    TBanqueCPBQ_NUMEMETVIR   : TStringField;
    TBanqueCPBQ_JOURFERMETUE : TStringField;
    TBanqueCPBQ_REPRELEVE    : TStringField;
    TBanqueCPBQ_REPLCR       : TStringField;
    TBanqueCPBQ_REPLCRFOURN  : TStringField;
    TBanqueCPBQ_REPVIR       : TStringField;
    TBanqueCPBQ_REPPRELEV    : TStringField;
    TBanqueCPBQ_REPBONAPAYER : TStringField;
    TBanqueCPBQ_REPIMPAYELCR : TStringField;
    TBanqueCPBQ_DELAIVIRORD  : TIntegerField;
    TBanqueCPBQ_DELAIVIRCHAUD: TIntegerField;
    TBanqueCPBQ_DELAIVIRBRULANT: TIntegerField;
    TBanqueCPBQ_DELAIPRELVORD: TIntegerField;
    TBanqueCPBQ_DELAIPRELVACC: TIntegerField;
    TBanqueCPBQ_DELAILCR     : TIntegerField;
    TBanqueCPBQ_GUIDECOMPATBLE: TStringField;
    TBanqueCPBQ_DERNSOLDEFRS : TFloatField;
    TBanqueCPBQ_DERNSOLDEDEV : TFloatField;
    TBanqueCPBQ_DATEDERNSOLDE: TDateTimeField;
    TBanqueCPBQ_RELEVEETRANGER: TStringField;
    TBanqueCPBQ_CALENDRIER   : TStringField;
    TBanqueCPBQ_RAPPAUTOREL  : TStringField;
    TBanqueCPBQ_RAPPROAUTOLCR: TStringField;
    TBanqueCPBQ_LETTREVIR    : TStringField;
    TBanqueCPBQ_LETTREPRELV  : TStringField;
    TBanqueCPBQ_LETTRELCR    : TStringField;
    TBanqueCPBQ_ENCOURSLCR   : TFloatField;
    TBanqueCPBQ_PLAFONDLCR   : TFloatField;
    TBanqueCPBQ_REPIMPAYEPRELV: TStringField;
    TBanqueCPBQ_ECHEREPPRELEV: TStringField;
    TBanqueCPBQ_ECHEREPLCR   : TStringField;
    TBanqueCPBQ_SOCIETE      : TStringField;
    TBanqueCPBQ_BANQUE       : TStringField;
    TBanqueCPBQ_BLOCNOTE     : TMemoField;
    TBanqueCPBQ_DELAIBAPLCR  : TIntegerField;
    TBanqueCPBQ_DELAITRANSINT: TIntegerField;
    TBanqueCPBQ_COMPTEFRAIS  : TStringField;
    TBanqueCPBQ_TYPEREMTRANS : TStringField;
    TBanqueCPBQ_INDREMTRANS  : TStringField;
    TBanqueCPBQ_COMMENTAIRE  : TStringField;
    TBanqueCPBQ_LETTRECHQ    : TStringField;
    TBanqueCPBQ_NUMEMETPRE   : TStringField;
    TBanqueCPBQ_DESTINATAIRE : TStringField;
    TBanqueCPBQ_MULTIDEVISE  : TStringField;
    TBanqueCPBQ_CODEIBAN     : TStringField;
    BQ_CODEIBAN              : THDBEdit;
    BQ_ETABBQ                : THDBEdit;
    BQ_GUICHET               : THDBEdit;
    BQ_NUMEROCOMPTE          : THDBEdit;
    BQ_CLERIB                : THDBEdit;
    procedure FormShow(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BAnnulerClick(Sender: TObject);
    procedure BMCbanqClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TBanqueCPAfterDelete(DataSet: TDataSet);
    procedure TBanqueCPAfterPost(DataSet: TDataSet);
    procedure Cbj1MouseDown(Sender: TObject; Button: TMouseButton;  Shift: TShiftState; X, Y: Integer);
    procedure BQ_CODEPOSTALDblClick(Sender: TObject);
    procedure BQ_VILLEDblClick(Sender: TObject);
    procedure BQ_DEVISEChange(Sender: TObject);
    procedure BQ_DIVTERRITDblClick(Sender: TObject);
    procedure BQ_PAYSChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BQ_CODEIBANEnter(Sender: TObject);
    procedure BQ_NUMEROCOMPTEKeyPress(Sender: TObject; var Key: Char);
  private
    LeQuel     : String ;
    CodeISO    : String ;
    Mode       : TActionFiche ;
    LaPage     : Integer ;
    FAvertir   : boolean ;
    AvecMvt    : Boolean ;
    OldCodeDev : String ;
    OkMajTotPointe : Boolean ;
    gbNoIban     : Boolean;
    gbChargement : boolean;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  CodExist : Boolean ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Procedure CurseurAuCode ;
    Procedure CurseurAuLibelle ;
    Procedure CurseurAuBanque ;
    Procedure CurseurAuGuichet ;
    Procedure CurseurAuNumCpte ;
    Procedure CurseurAuCleRib ;
    Procedure CurseurAuBic ;
    Procedure CurseurADevise ;
    Function  EnregOK : boolean ;
    Procedure ChargeEnreg ;
    Procedure LitJourFermeture ;
    Procedure EcritJourFermeture ;
    Function  VerifiRibBic : Boolean ;
    Procedure IsMouvementer ;
    Function  OkPourChangementDevise(OldCode,NewCode : String) : Boolean ;
    Procedure GererBoutons ;
    Function  RIBRenseigner: Boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses FichComm,PrintDBG, ParamDBG, CPTESAV ;

procedure TFbanqueCP.FormShow(Sender: TObject);
begin
ChangeSizeMemo(TBanqueCPBQ_BLOCNOTE) ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FAvertir:=False ; AvecMvt:=False ; OldCodeDev:='' ;
TBanqueCP.Open ;
BPages.ActivePage:=BPages.Pages[LaPage] ;
ChangeMask(BQ_DERNSOLDEFRS,V_PGI.OkDecV,V_PGI.SymbolePivot) ;
Case Mode Of
  taConsult : FicheReadOnly(Self) ;
  taCreat..taCreatOne : Bouge(nbInsert) ;
  end ;
if Not(Mode in [taCreat..taCreatOne])then TBanqueCP.FindKey([Lequel]) ;
{$IFDEF CCS3}
BQ_LANGUE.Visible:=False ; TBQ_LANGUE.Visible:=False ;
{$ENDIF}
  // Chargement
  ChargeEnreg ;
  // Boutons de navigation
  GererBoutons ;
end;

Function TFbanqueCP.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
  result:=FALSE  ; OkMajTotPointe:=FALSE ;
  Case Button of
     nblast,nbprior,nbnext,
     nbfirst,nbinsert : if Not OnSauve  then Exit ;
     nbPost :           if Not EnregOK  then Exit ;
     nbDelete :         if MsgBox.execute(1,'','')<>mrYes then Exit ;
     end ;
  if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[9]) ;
  If OkMajTotPointe Then RecalculTotPointeNew1(Lequel) ; ;
  result:=TRUE ;
  // Attention à la mise à jour des zones
  if Button=NbInsert
    then NewEnreg
    else ChargeEnreg ;
  // Maj des boutons de naviguation
  GererBoutons ;
END ;

Function TFbanqueCP.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TBanqueCP.Modified then
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

Function TFbanqueCP.OkPourChangementDevise(OldCode,NewCode : String) : Boolean ;
Var St : String ;
BEGIN
Result:=TRUE ;
If (NewCode=V_PGI.DevisePivot) Or (NewCode=V_PGI.DeviseFongible) Then Exit ;
If (OldCode<>V_PGI.DevisePivot) And (OldCode<>V_PGI.DeviseFongible) Then BEGIN Result:=FALSE ; Exit ; END ;
St:='Select e_general from ecriture Where e_general="'+TBanqueCPBQ_GENERAL.AsString+'" AND E_DEVISE<>"'+OldCode+'" AND E_DEVISE<>"'+NewCode+'" ' ;
If ExisteSQL(St)Then Result:=FALSE ;
END ;

Function TFbanqueCP.EnregOK : boolean ;
Var RibCalcul : String ;
BEGIN
result:=FALSE  ;
if TBanqueCP.state in [dsEdit,dsInsert]=False then Exit ;

if TBanqueCP.state in [dsEdit,dsInsert] then begin
   if TBanqueCPBQ_CODE.AsString='' then BEGIN MsgBox.Execute(2,'','') ; CurseurAuCode ; Exit ; END ;
   if TBanqueCPBQ_LIBELLE.AsString='' then BEGIN MsgBox.Execute(3,'','') ; CurseurAuLibelle ; Exit ; END ;
   if ExisteCarInter(TBanqueCPBQ_DOMICILIATION.AsString) then begin
      MsgBox.Execute(13,'','') ;
      if BQ_DOMICILIATION.CanFocus then BQ_DOMICILIATION.SetFocus ;
      Exit ;
   end;
   if ((VH^.CtrlRIB) and (codeISO='FR') and (RIBRenseigner)) then begin
     RibCalcul:=VerifRib(TBanqueCPBQ_ETABBQ.AsString,TBanqueCPBQ_GUICHET.AsString,TBanqueCPBQ_NUMEROCOMPTE.AsString) ;
     if (RibCalcul<>Trim(TBanqueCPBQ_CLERIB.AsString)) Then begin
       if MsgBox.Execute(11,'','')<>mrYes Then begin
         if BQ_CLERIB.CanFocus then BQ_CLERIB.SetFocus ;
         BQ_CODEIBAN.Field.Value := ''; Exit;
         end
       else begin
         TBanqueCP.Edit;
         BQ_CLERIB.Field.Value :=RibCalcul;
         BQ_CODEIBAN.Field.Value := ''; BQ_CODEIBANEnter(nil);
       end;
     END ;
   END ;
END;
if TBanqueCP.state in [dsInsert] then
   BEGIN
   if CodExist then Exit ;
   if Not VerifiRibBic then Exit ;
   END Else If TBanqueCP.state in [dsEdit] Then
   BEGIN
   if AvecMvt And (OldCodeDev<>BQ_DEVISE.Value) Then
     BEGIN
     OkMajTotPointe:=FALSE ;
     If OkPourChangementDevise(OldCodeDev,BQ_DEVISE.Value) Then OkMajTotPointe:=TRUE Else
       BEGIN
       Result:=FALSE ; MsgBox.Execute(12,'','') ; CurseurADevise ; Exit ;
       END ;
     END ;
   END ;
EcritJourFermeture ;
Result:=TRUE  ;
END ;

Function TFbanqueCP.CodExist : Boolean ;
BEGIN
Result:=False ;
if Presence('BANQUECP','BQ_CODE',TBanqueCPBQ_CODE.AsString) then
   BEGIN
   MsgBox.Execute(10,'','')  ; CurseurauCode ;
   Result:=True ;
   END ;
END ;

Procedure TFbanqueCP.ChargeEnreg ;
BEGIN
  InitCaption(Self,TBanqueCPBQ_CODE.AsString,TBanqueCPBQ_LIBELLE.AsString) ;

  EDev.Text             :=  BQ_DEVISE.Text ;
  BQ_DERNSOLDEFRS.Value :=  TBanqueCPBQ_DERNSOLDEFRS.AsFloat ;
  BQ_DERNSOLDEDEV.Value :=  TBanqueCPBQ_DERNSOLDEDEV.AsFloat ;
  BQ_GENERAL.Enabled    :=  ((Lequel='') and (TBanqueCPBQ_GENERAL.AsString=''));

  LitJourFermeture ;

  if Mode=taConsult then
    begin
    FicheReadOnly(Self) ;
    Exit ;
    end ;

  IsMouvementer ;
  BQ_BANQUE.Enabled := ((Not AvecMvt) or (TBanqueCPBQ_Banque.AsString='')) ;
  BQ_CODE.Enabled   := Not AvecMvt ;
  // Permet le non calcul de l'iban
  gbNoIban := True;

  // Gestion des erreurs dues au pays non renseigné
  if TBanqueCPBQ_PAYS.AsString = '' then
    if RibRenseigner then begin
      TBanqueCP.Edit ;
//      TBanqueCPBQ_PAYS.AsString := 'FRA' ;
//      BQ_PAYS.Value := 'FRA' ;
      BQ_PAYS.Field.Value := CodePaysDeIso('FR') ;
    end ;

  gbChargement:= True;
  BQ_PAYSChange(nil);
  gbChargement:= False;
  BQ_CODEIBANEnter(nil);
END ;

procedure TFbanqueCP.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFbanqueCP.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFbanqueCP.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFbanqueCP.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFbanqueCP.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFbanqueCP.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFbanqueCP.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFbanqueCP.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then Bouge(nbInsert) ;
   if Mode=taCreatOne then Close ;
   END ;
end;

procedure TFbanqueCP.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFbanqueCP.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

Procedure TFbanqueCP.CurseurAuCode ;
BEGIN BPages.ActivePage:=PIdent ; BQ_CODE.SetFocus ; END ;

Procedure TFbanqueCP.CurseurAuLibelle ;
BEGIN BPages.ActivePage:=PIdent ; BQ_LIBELLE.SetFocus ; END ;

Procedure TFbanqueCP.CurseurAuBanque ;
BEGIN BPages.ActivePage:=PIdent ; BQ_ETABBQ.SetFocus ; END ;

Procedure TFbanqueCP.CurseurAuGuichet ;
BEGIN BPages.ActivePage:=PIdent ; BQ_GUICHET.SetFocus ; END ;

Procedure TFbanqueCP.CurseurAuNumCpte ;
BEGIN BPages.ActivePage:=PIdent ; BQ_NUMEROCOMPTE.SetFocus ; END ;

Procedure TFbanqueCP.CurseurAuCleRib ;
BEGIN BPages.ActivePage:=PIdent ; BQ_CLERIB.SetFocus ; END ;

Procedure TFbanqueCP.CurseurAuBic ;
BEGIN BPages.ActivePage:=PIdent ; BQ_CODEBIC.SetFocus ; END ;

Procedure TFbanqueCP.CurseurADevise ;
BEGIN BPages.ActivePage:=PIdent ; BQ_DEVISE.SetFocus ; END ;

Procedure TFbanqueCP.NewEnreg ;
BEGIN
  InitNew(TBanqueCP) ;
  BQ_BANQUE.Enabled   :=True ;
  BQ_CODE.Enabled     :=True ;

  TBanqueCPBQ_GENERAL.AsString:=Lequel ;
  BQ_GENERAL.ExisteH ;
  TBanqueCPBQ_JOURFERMETUE.AsString:='000000X' ;

  BQ_PAYS.Value             := codepaysdeiso('FR');
  BQ_LANGUE.Value           :=V_PGI.LanguePrinc ;
  BQ_DEVISE.Value           :=V_PGI.DevisePivot ;
  BQ_RAPPAUTOREL.Checked    :=True ;
  BQ_RAPPROAUTOLCR.Checked  :=True ;
  BQ_LETTRELCR.ItemIndex    :=0 ;
  BQ_ECHEREPPRELEV.Checked  :=True ;
  BQ_ECHEREPLCR.Checked     :=True ;

  LitJourFermeture ;
END ;

Procedure TFbanqueCP.LitJourFermeture ;
Var i : Byte ;
    C : TComponent ;
    St : String[7] ;
BEGIN
St:=TBanqueCPBQ_JOURFERMETUE.AsString ;
for i:=1 to 7 do
   BEGIN
   C:=FindComponent('Cbj'+IntToStr(i)) ;
   TCheckBox(C).Checked:=(St[i]='X') ;
   END ;
END ;

Procedure TFbanqueCP.EcritJourFermeture ;
Var i : Byte ;
    C : TComponent ;
    St : String[7] ;
BEGIN
St:='' ;
for i:=1 to 7 do
   BEGIN
   C:=FindComponent('Cbj'+IntToStr(i)) ;
   if TCheckBox(C).Checked then St:=St+'X' else St:=St+'0' ;
   END ;
TBanqueCPBQ_JOURFERMETUE.AsString:=St ;
END ;

Function TFbanqueCP.VerifiRibBic : Boolean ;
BEGIN
Result:=False ;
if codeISO='FR' then
  BEGIN
   if TBanqueCPBQ_ETABBQ.AsString='' then
      BEGIN MsgBox.Execute(4,'','') ; CurseurAuBanque ; Exit ; END ;
   if TBanqueCPBQ_GUICHET.AsString='' then
      BEGIN MsgBox.Execute(5,'','') ; CurseurAuGuichet ; Exit ; END ;
   if TBanqueCPBQ_NUMEROCOMPTE.AsString='' then
      BEGIN MsgBox.Execute(6,'','') ; CurseurAuNumCpte ; Exit ; END ;
   if TBanqueCPBQ_CLERIB.AsString='' then
      BEGIN MsgBox.Execute(7,'','') ; CurseurAuCleRib ; Exit ; END ;
  END else
  BEGIN
   if TBanqueCPBQ_CODEBIC.AsString<>''then BEGIN Result:=True ; Exit ; END
   else
   if ((TBanqueCPBQ_ETABBQ.AsString='')OR(TBanqueCPBQ_GUICHET.AsString='')OR
      (TBanqueCPBQ_NUMEROCOMPTE.AsString='')OR(TBanqueCPBQ_CLERIB.AsString='')) then
      BEGIN MsgBox.Execute(8,'','') ; CurseurAuBic ; Exit ; END ;
  END ;
// if Not ControleSaisieRib then BEGIN     Exit ; END ;
Result:=True ;
END ;

procedure TFbanqueCP.BMCbanqClick(Sender: TObject);
Var Laction : TActionFiche ;
begin
if BQ_CODE.CanFocus then BQ_CODE.SetFocus ; 
if BQ_BANQUE.Value='' then Laction:=taCreatOne else Laction:=taModif ;
FicheBanque_AGL(BQ_BANQUE.Value,Laction,0) ;
BQ_BANQUE.Reload ;
end;

Procedure FicheBanqueCP(CpteGeneral : String ; Comment : TActionFiche ; QuellePage : Integer) ;
var FBanqueCP: TFBanqueCP;
begin
if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FBanqueCP:=TFBanqueCP.Create(Application) ;
try
  FBanqueCP.LeQuel:=CpteGeneral ;
  FBanqueCP.Mode:=Comment ;
  FBanqueCP.LaPage:=QuellePage ;
  FBanqueCP.ShowModal ;
  finally
  FBanqueCP.Free ;
  end ;
Screen.Cursor:=crDefault ;
end ;

procedure TFbanqueCP.FormClose(Sender: TObject; var Action: TCloseAction);
begin if FAvertir then AvertirTable('ttBanqueCP') ; end;

procedure TFbanqueCP.TBanqueCPAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFbanqueCP.TBanqueCPAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFbanqueCP.Cbj1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if TBanqueCP.State in [dsEdit,dsInsert] then Exit ;
TBanqueCP.Edit ;
TBanqueCPBQ_CONTACT.AsString:=TBanqueCPBQ_CONTACT.AsString ;
end;

procedure TFbanqueCP.BQ_CODEPOSTALDblClick(Sender: TObject);
begin VerifCodePostal(TBanqueCP,THDBEdit(BQ_CODEPOSTAL),THDBEdit(BQ_VILLE),TRUE) ; end;

procedure TFbanqueCP.BQ_VILLEDblClick(Sender: TObject);
begin VerifCodePostal(TBanqueCP,THDBEdit(BQ_CODEPOSTAL),THDBEdit(BQ_VILLE),FALSE) ; end;

Procedure TFbanqueCP.IsMouvementer ;
BEGIN
if TBanqueCP.State=dsInsert then AvecMvt:=False else
  BEGIN
  AvecMvt:=ExisteSQL('Select G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+TBanqueCPBQ_GENERAL.AsString+'" '+
                     'And Exists(Select E_GENERAL FROM ECRITURE Where E_GENERAL="'+TBanqueCPBQ_GENERAL.AsString+'")') ;
  OldCodeDev:=TBanqueCPBQ_DEVISE.AsString ;
  END ;
END ;

procedure TFbanqueCP.BQ_DEVISEChange(Sender: TObject);
begin EDev.Text:=BQ_DEVISE.Text ; end;

procedure TFbanqueCP.BQ_DIVTERRITDblClick(Sender: TObject);
begin
PaysRegion(BQ_PAYS,BQ_DIVTERRIT,True) ;
end;

procedure TFbanqueCP.BQ_PAYSChange(Sender: TObject);
//var
//  Q : TQuery;
//  szISO2 : String;
begin
codeISO := codeISODuPays(BQ_PAYS.Value);
  if (codeISO='FR') or (codeISO = '') then begin
    BQ_ETABBQ.Enabled       := True;
    BQ_GUICHET.Enabled      := True;
    BQ_NUMEROCOMPTE.Enabled := True;
    BQ_CLERIB.Enabled       := True;
    BQ_CODEIBAN.Enabled     := False;
    if not(gbChargement) then begin
      TBanqueCP.Edit;
//      BQ_CODEIBAN.Field.Value := codepaysdeiso('FR');
      BQ_CODEIBAN.Field.Value := codeIso;  // CA - 02/10/2003 - FQ 12547
    end;
    end
  else begin
    if not(gbChargement) then begin
      TBanqueCP.Edit;
      BQ_ETABBQ.Field.Value       := '';
      BQ_GUICHET.Field.Value      := '';
      BQ_NUMEROCOMPTE.Field.Value := '';
      BQ_CLERIB.Field.Value       := '';
    end;
    BQ_ETABBQ.Enabled       := False;
    BQ_GUICHET.Enabled      := False;
    BQ_NUMEROCOMPTE.Enabled := False;
    BQ_CLERIB.Enabled       := False;
    BQ_CODEIBAN.Enabled     := True;

    if ((SBanqueCP.State in [dsInsert,dsEdit]) and  not(gbChargement)) then begin
      {Q:=OpenSQL('SELECT PY_CODEISO2 FROM PAYS WHERE PY_PAYS="'+BQ_PAYS.Value+'"',True);
      if not Q.EOF then begin
        szISO2:=Q.Fields[0].AsString;}
        TBanqueCPBQ_CODEIBAN.AsString := codeISO;
      {end;
      Ferme(Q);}
    end;
  end;
end;

procedure TFbanqueCP.BImprimerClick(Sender: TObject);
begin PrintDBGrid(Nil,Nil,Caption,'PRT_BQECP') ; end;

procedure TFbanqueCP.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFbanqueCP.BQ_CODEIBANEnter(Sender: TObject);
var
  szPays, szIban : string;
begin
  // Au chargement : Ne fait rien
  if gbNoIban then begin gbNoIban := False; Exit; end;

	szPays :=  codeISO ;
  if szPays = 'FRA' then szPays := 'FR'
  else if szPays = 'ESP' then szPays := 'ES'
  else szPays:=Copy(szPays,1,2);

  // Pas de code iban ou iban incorrect : Le calcul
  szIban := calcIBAN(szPays, calcRIB(szPays,BQ_ETABBQ.Text,BQ_GUICHET.Text,BQ_NUMEROCOMPTE.Text,BQ_CLERIB.Text));
//  if ((BQ_CODEIBAN.Text = '') or ((szPays = 'FR') and (szIban<>BQ_CODEIBAN.Text) )) then begin
  if ((Trim(BQ_CODEIBAN.Text) = '') or ((szPays = 'FR') and (szIban<>Trim(BQ_CODEIBAN.Text)) )) then begin
    TBanqueCP.Edit;
    BQ_CODEIBAN.Field.Value := szIban;
  end;
end;

procedure TFbanqueCP.BQ_NUMEROCOMPTEKeyPress(Sender: TObject; var Key: Char);
begin
// Authorise uniquement les caractères A-Z et 0-9 et Backspace
if not(Key in ['a'..'z','A'..'Z','0'..'9',#8]) then Key:=#0 ;
end;

procedure TFbanqueCP.GererBoutons;
var UpEnable, DnEnable: Boolean;
begin
  // Maj des boutons de navigation
   UpEnable := Enabled and not TBanqueCP.BOF;
   DnEnable := Enabled and not TBanqueCP.EOF;
   BFirst.Enabled := UpEnable;
   BPrev.Enabled  := UpEnable;
   BNext.Enabled  := DnEnable;
   BLast.Enabled  := DnEnable;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 21/01/2003
Modifié le ... :   /  /
Description .. : Retourne True si le Rib est renseigné dans l'enregistrement
Suite ........ : courant.
Suite ........ :
Suite ........ : Formule à vérifier
Mots clefs ... :
*****************************************************************}
function TFbanqueCP.RIBRenseigner: Boolean ;
begin
  Result := (Trim(TBanqueCPBQ_ETABBQ.AsString) <> '')
            and (Trim(TBanqueCPBQ_GUICHET.AsString) <> '')
            and (Trim(TBanqueCPBQ_NUMEROCOMPTE.AsString) <> '') ;
end;

end.
