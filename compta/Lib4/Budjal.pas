{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par BUDJAL_TOM.PAS
Mots clefs ... : 
*****************************************************************}
unit Budjal;

interface

uses
    WinTypes, WinProcs, Classes, Graphics, Forms, Controls, StdCtrls, Tabs,
    Buttons, ExtCtrls, Grids, Mask,  DB, DBTables, DBCtrls,HEnt1,
    Hctrls, HDB, DBGrids, Hqry, TabNotBk, Dialogs, Spin, SysUtils, Ent1,
    ComCtrls, hmsgbox,hcompte, Messages, Menus, HRichEdt, HRichOLE,
    HSysMenu,MajTable, HTB97, HRegCpte;

Procedure FicheBudjal(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
Procedure FicheBudjalMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);

type
  TFbudjal = class(TForm)
    BPages             : TPageControl;
    HPB                : TPanel;
    BAide              : TBitBtn;
    BAnnuler           : TBitBtn;
    BValider           : TBitBtn;
    BImprimer          : TBitBtn;
    BFirst             : TBitBtn;
    BPrev              : TBitBtn;
    BNext              : TBitBtn;
    BLast              : TBitBtn;
    BInsert            : TBitBtn;
    BFerme             : TBitBtn;
    FAutoSave          : TCheckBox;
    PCaract            : TTabSheet;
    PComplement        : TTabSheet;
    PInfo              : TTabSheet;
    HGBDates           : TGroupBox;
    BJ_FERME: TDBCheckBox;
    TBJ_BLOCNOTE: TGroupBox;
    BJ_BLOCNOTE: THDBRichEditOLE;
    SBudjal: TDataSource;
    TBJ_DATECREATION: THLabel;
    BJ_DATECREATION: TDBEdit;
    TBJ_DATEMODIF: THLabel;
    BJ_DATEMODIF: TDBEdit;
    TBJ_DATEOUVERTURE: THLabel;
    BJ_DATEOUVERTURE: TDBEdit;
    TBJ_DATEFERMETURE: THLabel;
    BJ_DATEFERMETURE: TDBEdit;
    DBNav              : TDBNavigator;
    MsgBox             : THMsgBox;
    TBJ_BUDJAL: THLabel;
    BJ_BUDJAL: TDBEdit;
    TBJ_LIBELLE: THLabel;
    BJ_LIBELLE: TDBEdit;
    TBJ_ABREGE: THLabel;
    BJ_ABREGE: TDBEdit;
    QBudjal: TQuery;
    HMTrad: THSystemMenu;
    TBJ_AXE: TLabel;
    BJ_AXE: THDBValComboBox;
    QBudjalBJ_BUDJAL: TStringField;
    QBudjalBJ_LIBELLE: TStringField;
    QBudjalBJ_ABREGE: TStringField;
    QBudjalBJ_DATECREATION: TDateTimeField;
    QBudjalBJ_DATEMODIF: TDateTimeField;
    QBudjalBJ_DATEOUVERTURE: TDateTimeField;
    QBudjalBJ_DATEFERMETURE: TDateTimeField;
    QBudjalBJ_FERME: TStringField;
    QBudjalBJ_BLOCNOTE: TMemoField;
    QBudjalBJ_COMPTEURNORMAL: TStringField;
    QBudjalBJ_COMPTEURSIMUL: TStringField;
    QBudjalBJ_UTILISATEUR: TStringField;
    QBudjalBJ_EXODEB: TStringField;
    QBudjalBJ_EXOFIN: TStringField;
    QBudjalBJ_PERDEB: TDateTimeField;
    QBudjalBJ_PERFIN: TDateTimeField;
    QBudjalBJ_AXE: TStringField;
    QBudjalBJ_GENEATTENTE: TStringField;
    QBudjalBJ_SECTATTENTE: TStringField;
    GroupBox1: TGroupBox;
    BJ_EXODEB: THDBValComboBox;
    TBJ_EXOFIN: TLabel;
    BJ_EXOFIN: THDBValComboBox;
    Bevel1: TBevel;
    GroupBox3: TGroupBox;
    TBJ_GENEATTENTE: TLabel;
    TBJ_SECTATTENTE: TLabel;
    TBJ_PERFIN: TLabel;
    DS: TDataSource;
    GbCpte: TGroupBox;
    FListeG: THGrid;
    FListeS: THGrid;
    Panel1: TPanel;
    BUp: TBitBtn;
    BDown: TBitBtn;
    GBTab: TGroupBox;
    BVal: TBitBtn;
    BFer: TBitBtn;
    FListe: THDBGrid;
    QTab: TQuery;
    BJ_PERDEB: THDBValComboBox;
    BJ_PERFIN: THDBValComboBox;
    QBudjalBJ_BUDGENES: TMemoField;
    QBudjalBJ_BUDSECTS: TMemoField;
    TBJ_COMPTEURNORMAL: THLabel;
    BJ_COMPTEURNORMAL: THDBValComboBox;
    BJ_COMPTEURSIMUL: THDBValComboBox;
    Label3: TLabel;
    Label4: TLabel;
    Bajout: TBitBtn;
    BDel: TBitBtn;
    Label1: TLabel;
    BJ_SENS: THDBValComboBox;
    QBudjalBJ_SENS: TStringField;
    TBJ_NATJAL: TLabel;
    QBudjalBJ_NATJAL: TStringField;
    BJ_NATJAL: THDBValComboBox;
    QBudjalBJ_CATEGORIE: TStringField;
    QBudjalBJ_SOUSPLAN: TStringField;
    TBJ_CATEGORIE: TLabel;
    BJ_CATEGORIE: THDBValComboBox;
    TBJ_SOUSPLAN: THLabel;
    BJ_SOUSPLAN: TDBEdit;
    BZoomSousPlan: TToolbarButton97;
    BCopieComplement: TBitBtn;
    BCopierJal: TBitBtn;
    QBudjalBJ_BUDGENES2: TMemoField;
    QBudjalBJ_BUDSECTS2: TMemoField;
    PopZ: TPopupMenu;
    BCopieCpt: TBitBtn;
    BCopieSect: TBitBtn;
    BCopieDeux: TBitBtn;
    BJ_CONTROLESENS: TDBCheckBox;
    QBudjalBJ_CONTROLESENS: TStringField;
    BJ_GENEATTENTE: THDBCpteEdit;
    BJ_SECTATTENTE: THDBCpteEdit;
    procedure FormCreate(Sender: TObject);
    procedure SBudjalDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BAjoutClick(Sender: TObject);
    procedure BFerClick(Sender: TObject);
    procedure FListeGExit(Sender: TObject);
    procedure FListeSExit(Sender: TObject);
    procedure BValClick(Sender: TObject);
    procedure BDelClick(Sender: TObject);
    procedure BUpClick(Sender: TObject);
    procedure BDownClick(Sender: TObject);
    procedure BJ_EXODEBChange(Sender: TObject);
    procedure BJ_EXOFINChange(Sender: TObject);
    procedure BJ_AXEChange(Sender: TObject);
    procedure BJ_AXEClick(Sender: TObject);
    procedure BJ_PERDEBChange(Sender: TObject);
    procedure BJ_PERFINChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure QTabAfterPost(DataSet: TDataSet);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure GbCpteEnter(Sender: TObject);
    procedure QBudjalAfterPost(DataSet: TDataSet);
    procedure QBudjalAfterDelete(DataSet: TDataSet);
    procedure BJ_NATJALChange(Sender: TObject);
    procedure BZoomSousPlanClick(Sender: TObject);
    procedure BJ_SOUSPLANKeyPress(Sender: TObject; var Key: Char);
    procedure BCopieComplementClick(Sender: TObject);
    procedure BCopierJalClick(Sender: TObject);
    procedure BJ_CATEGORIEChange(Sender: TObject);
    procedure BCopieCptClick(Sender: TObject);
    procedure BCopieSectClick(Sender: TObject);
    procedure BCopieDeuxClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FListeGSorted(Sender: TObject);
    procedure FListeSSorted(Sender: TObject);
  private { Private declarations }
    Q : TQuery ;
    Mode : TActionFiche ;
    LaPage : Integer ;
    Lequel,LeCompte : String ;
    LesModif : String ;
    QuelAxe : String ;
    DeGen,DeSect : Boolean ;
    Modifier,BGModifier,BSModifier : Boolean ;
    MemoCod : String ;
    IsMouvementer : Boolean ;
    Lefb : TFichierBase ;
    FAvertir : Boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Procedure CurseurAuCode ;
    Procedure CurseurAuLibelle ;
    Function  EnregOK : boolean ;
    Procedure ChargeEnreg ;
    Function  CodeValide : Boolean ;
    Procedure AffecteLe(Champ, Valeur : string) ;
    Procedure InitModifEnSerie(StModif : string) ;
    Procedure LitBudgensBudsects(Stc : String ; G : THgrid ; PremFois : Boolean) ;
    Function  EcritBudgensBudsects(G : THgrid ; Var Ind : Integer) : String ;
    Function  CpteAttenteOk ( FicBase : TFichierBase) : Boolean ;
    Procedure GriseDegriseControl ;
    Function  DoublonExiste(St : String ; G : THGrid) : Boolean ;
    Function  ExoOuv : Boolean ;
    Procedure PositionnePerDeb ;
    Procedure PositionnePerFin ;
    Function  TrouveCpteMvt(G : THGrid) : Boolean ;
    Function  CoherenceSection : Boolean ;
    procedure GenereCroisement(CodeACreer,CodeChoisi : String) ;
    Procedure DesactiveLesControles ;
    Procedure CopieCatSurJournaux ;
    procedure BCopieListe(Qui : Integer) ;
  public  { Public declarations }
  end;

  Function EstMouvementeBudjal(Lejal : String) : Boolean ;

implementation

{$R *.DFM}

uses
{$IFNDEF IMP}
     QRBudJal,
{$ENDIF}
     BudJalSP,Choix,Hstatus,Filtre,UtilPGI ;

procedure TFbudjal.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; Q:=NIL ; FAvertir:=False ;
LeCompte:=W_W ;
end;

Procedure LibelleCatBud(CodeCat,SousPlanBud : String ; Var LibBJ,AbrBJ : String) ;
Var k,l : Integer ;
    St,StSousPlan,St2,Code,Lib : String ;
    SousPlan : TSousPlan ;
    SPC : TSousPlanCat ;
    Cat : TUneCatBud ;
BEGIN
LibBJ:='' ; AbrBJ:='' ; SPC:=SousPlanCat(CodeCat,TRUE) ; Cat:=QuelleCatBud(CodeCat) ;
If Cat.Code='' Then Exit ;
For l:=1 To MaxSousPlan Do
  BEGIN
  St:=Cat.SurJal[l] ; SousPlan:=SPC[l] ;
  If St<>'' Then
     BEGIN
     StSousPlan:=ReadTokenSt(SousPlanBud) ;
     If StSousPlan<>'' Then
        For k:=0 To SousPlan.ListeSP.Count-1 Do
          BEGIN
          St2:=SousPlan.ListeSP.Strings[k] ;
          Code:=ReadTokenSt(St2) ; Lib:=ReadTokenSt(St2) ;
          If Code=StSousPlan Then
            BEGIN
            LibBJ:=LibBJ+Lib ; AbrBJ:=AbrBJ+Code ; Break ;
            END ;
          END ;
     END ;
  END ;
END ;

procedure TFbudjal.SBudjalDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
//    Lib,Abr : String ;
begin
BInsert.Enabled:=Not(QBudjal.State in [dsEdit,dsInsert]) ;
BCopierJal.Enabled:=(QBudjal.State in [dsInsert]) ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not QBudjal.BOF;
   DnEnable := Enabled and not QBudjal.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
//   if ((*(Field.FieldName='BJ_SOUSPLAN')*) and (BJ_LIBELLE.Field.AsString='') And (BJ_SOUSPLAN.Field.AsString<>'')) then
//      BEGIN
//      LibelleCatBud(BJ_CATEGORIE.Field.AsString,BJ_SOUSPLAN.Field.AsString,Lib,Abr) ;
//      BJ_LIBELLE.Field.AsString:=Lib ;
//      BJ_ABREGE.Field.AsString:=Abr ;
//      END ;
   if ((Field.FieldName='BJ_LIBELLE') and (BJ_ABREGE.Field.AsString='')) then
      BJ_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

Procedure TFbudjal.DesactiveLesControles ;
BEGIN
GbCpte.Enabled:=True ; FListeG.Enabled:=True ;
FListeS.Enabled:=True ; Panel1.Enabled:=False ;
BUp.Enabled:=False ; BDown.Enabled:=False ; BAjout.Enabled:=False ; BDel.Enabled:=False ;
BCopieComplement.Enabled:=False ;
FlisteG.SortEnabled:=False ; FlisteS.SortEnabled:=False ;
END ;

procedure TFbudjal.FormShow(Sender: TObject);
begin
MakeZoomOLE(Handle) ;
GBTab.Parent:=PComplement ; GBTab.Align:=alClient ; GBTab.Visible:=False ; DeGen:=False ;
DeSect:=False ; MemoCod:='' ;    FListe.MultiSelection := True ;
RecupWhereSQL(Q,QBudjal) ;
if ((Q=NIL) and (Lequel<>'')) then QBudjal.SQL.Add('Where BJ_BUDJAL="'+Lequel+'"') ;
if ((Q=NIL) and (Lequel<>'')) then ChangeSQL(QBudjal) ;
ChangeSizeMemo(QBudjalBJ_BLOCNOTE) ;
ChangeSizeLongChar(QBudjalBJ_BUDGENES) ;
ChangeSizeLongChar(QBudjalBJ_BUDSECTS) ;
ChangeSizeLongChar(QBudjalBJ_BUDGENES2) ;
ChangeSizeLongChar(QBudjalBJ_BUDSECTS2) ;
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
   BJ_CATEGORIE.Visible:=False ; TBJ_CATEGORIE.Visible:=False ;
   BJ_SOUSPLAN.Visible:=False  ; TBJ_SOUSPLAN.Visible:=False ;
   BZoomSousPlan.Visible:=False ;
   END ;
QBudjal.Open ;
if(Lequel<>'')And((Mode in [taCreat..taCreatOne])=False) then
   BEGIN
   if Not QBudjal.Locate('BJ_BUDJAL',Lequel,[]) then
     BEGIN MessageAlerte(MsgBox.Mess[5]) ; PostMessage(Handle,WM_CLOSE,0,0);  Exit ; END ;
   END ;
BPages.ActivePage:=BPages.Pages[LaPage] ;
Case Mode Of
     taConsult           : BEGIN FicheReadOnly(Self) ; DesactiveLesControles ; Exit ; END ;
     taCreat..taCreatOne : BEGIN Bouge(nbInsert) ; BJ_BUDJAL.Text:=Lequel ; BAnnuler.Enabled:=False ; END ;
     taModif             : BEGIN END ;
     taModifEnSerie      : InitModifEnSerie(LesModif);
     end ;
if(QBudjal.Eof) then Bouge(nbInsert) ;
end;

Procedure TFbudjal.ChargeEnreg ;
Var Lib,Abr : String ;
BEGIN
FListeS.SortedCol:=-1 ; FListeG.SortedCol:=-1 ; 
if LeCompte=QBudjalBJ_BUDJAL.AsString then
   BEGIN
   if Mode=taConsult then BEGIN FicheReadOnly(Self) ; DesactiveLesControles ; END ;
   Exit ;
   END ;
if (QBudjalBJ_LIBELLE.AsString='') And (QBudjalBJ_SOUSPLAN.AsString<>'') then
   BEGIN
   LibelleCatBud(QBudjalBJ_CATEGORIE.AsString,QBudjalBJ_SOUSPLAN.AsString,Lib,Abr) ;
   QBudjalBJ_LIBELLE.AsString:=Lib ;
   QBudjalBJ_ABREGE.AsString:=Abr ;
   END ;
LeCompte:=QBudjalBJ_BUDJAL.AsString ;
InitCaption(Self,BJ_BUDJAL.text,BJ_LIBELLE.text) ;
if QBudjal.State<>dsInsert then
   BEGIN
   if MemoCod<>QBudjalBJ_BUDJAL.AsString then
      BEGIN
      IsMouvementer:=EstMouvementeBudjal(QBudjalBJ_BUDJAL.AsString) ;
      LitBudgensBudsects(Trim(QBudjalBJ_BUDGENES.AsString),FListeG,TRUE) ;
      LitBudgensBudsects(Trim(QBudjalBJ_BUDGENES2.AsString),FListeG,FALSE) ;
      LitBudgensBudsects(Trim(QBudjalBJ_BUDSECTS.AsString),FListeS,TRUE) ;
      LitBudgensBudsects(Trim(QBudjalBJ_BUDSECTS2.AsString),FListeS,FALSE) ;
      ListePeriode(BJ_EXODEB.Value,BJ_PERDEB.Items,BJ_PERDEB.Values,TRUE) ;
      ListePeriode(BJ_EXOFIN.Value,BJ_PERFIN.Items,BJ_PERFIN.Values,FALSE) ;
      BJ_PERDEB.ItemIndex:=BJ_PERDEB.Values.IndexOf(DateToStr(QBudjalBJ_PERDEB.AsDateTime)) ;
      BJ_PERFIN.ItemIndex:=BJ_PERFIN.Values.IndexOf(DateToStr(QBudjalBJ_PERFIN.AsDateTime)) ;
      MemoCod:=QBudjalBJ_BUDJAL.AsString ;
      END ;
   END ;
BJ_AXEChange(Nil) ; GriseDegriseControl ; BJ_BUDJAL.Enabled:=False ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; DesactiveLesControles ; Exit ; END ;
END ;

Function TFbudjal.Bouge(Button: TNavigateBtn) : boolean ;
Var
  Ind : Integer ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : if Not EnregOK then Exit ;
   nbDelete         :  ;
   end ;
if Button = nbPost then
   BEGIN
   if BGModifier or BSModifier then CopieCatSurJournaux ;
   if QBudJal.State=dsBrowse then QBudJal.Edit ;
   DateModification(QBudjal,'BJ') ;
   Ind:=1 ;
   QBudjalBJ_BUDGENES.AsString:=EcritBudgensBudsects(FListeG,Ind) ;
   If Ind>0 Then QBudjalBJ_BUDGENES2.AsString:=EcritBudgensBudsects(FListeG,Ind) ;
   Ind:=1 ;
   QBudjalBJ_BUDSECTS.AsString:=EcritBudgensBudsects(FListeS,Ind) ;
   If Ind>0 Then QBudjalBJ_BUDSECTS2.AsString:=EcritBudgensBudsects(FListeS,Ind) ;

   // 13343
   if QBudjalBJ_BUDGENES.AsString = '' then QBudjalBJ_BUDGENES.AsString := ' ';
   if QBudjalBJ_BUDGENES2.AsString = '' then QBudjalBJ_BUDGENES2.AsString := ' ';
   if QBudjalBJ_BUDSECTS.AsString = '' then QBudjalBJ_BUDSECTS.AsString := ' ';
   if QBudjalBJ_BUDSECTS2.AsString = '' then QBudjalBJ_BUDSECTS2.AsString := ' ';
   END ;

if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[5]) ;
if Button=nbCancel then BJ_AXEChange(Nil) ;
Result:=TRUE ; Modifier:=False ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFbudjal.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
Result:=FALSE  ;
if (QBudjal.Modified) or Modifier then
   BEGIN
   if (Mode in [taCreat..taCreatOne]) And
      (QBudjalBJ_BUDJAL.asString='') then Rep:=mrNo
   else if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
     mrYes : if not Bouge(nbPost) then Exit ;
     mrNo :  if not Bouge(nbCancel) then Exit ;
     mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFbudjal.EnregOK : boolean ;
BEGIN
Result:=False ;
if (QBudjal.state in [dsEdit,dsInsert]=False) And (Not Modifier) then Exit ;
if Not CodeValide then Exit ;
if QBudjal.state in [dsEdit,dsInsert] then
   BEGIN
   if QBudjalBJ_LIBELLE.AsString='' then BEGIN CurseurAuLibelle ; Msgbox.Execute(3,'','') ; Exit ; END ;
   if QBudjalBJ_COMPTEURNORMAL.AsString='' then BEGIN Msgbox.Execute(25,'','') ; Exit ; END ;
   if Not CpteAttenteOk(fbBudGen) then Exit ;
   if Not CpteAttenteOk(Lefb) then Exit ;
   if Not CoherenceSection then Exit ;
   if Not ExoOuv then
      BEGIN
      BPages.ActivePage:=PCaract ; If BJ_EXODEB.Focused Then BJ_EXODEB.SetFocus ; Msgbox.Execute(12,'','') ; Exit ;
      END ;
   END ;
Result:=True ;
END ;

Procedure TFbudjal.CurseurAuCode ;
BEGIN BPages.ActivePage:=PCaract ; BJ_BUDJAL.SetFocus ; END ;

Procedure TFbudjal.CurseurAuLibelle ;
BEGIN BPages.ActivePage:=PCaract ; BJ_LIBELLE.SetFocus ; END ;

procedure TFbudjal.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFbudjal.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFbudjal.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFbudjal.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFbudjal.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFbudjal.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFbudjal.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then Bouge(nbInsert) ;
   if ((Mode=taCreatOne) or (Mode=taModifEnSerie)) then Close ;
   END ;
end;

procedure TFbudjal.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFbudjal.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
If Not HPB.Enabled Then BEGIN CanClose:=FALSE ; Exit ; END ;
BFerme.SetFocus ; CanClose:=OnSauve ;
end;

procedure TFbudjal.BImprimerClick(Sender: TObject);
begin
{$IFNDEF IMP}
PlanBudJal(QBudjalBJ_AXE.AsString, QBudjalBJ_BUDJAL.AsString,True)
{$ENDIF}
end;

Procedure TFbudjal.NewEnreg ;
BEGIN
InitNew(QBudjal) ; IsMouvementer:=False ; MemoCod:='' ; GriseDegriseControl ;
FListeG.VidePile(False) ; FListeS.VidePile(False) ;
QBudjalBJ_AXE.AsString:='A1' ; BJ_AXEChange(Nil) ;
QBudJalBJ_NATJAL.AsString:='CHA' ; BJ_NATJALChange(Nil) ;
BJ_EXODEB.Value:=VH^.Encours.Code ; BJ_EXOFIN.Value:=VH^.Encours.Code ;
QBudjalBJ_EXODEB.AsString:=VH^.Encours.Code ;
QBudjalBJ_EXOFIN.AsString:=VH^.Encours.Code ;
PositionnePerDeb ; PositionnePerFin ;
if BJ_COMPTEURNORMAL.Values.Count>0 then BJ_COMPTEURNORMAL.Value:=BJ_COMPTEURNORMAL.Values[0] ;
BJ_BUDJAL.Enabled:=True ;
BPages.ActivePage:=PCaract ; BJ_BUDJAL.SetFocus ; DateCreation(QBudjal,'BJ') ;
BcopierJal.Enabled:=TRUE ;
END ;

Procedure FicheBudjal(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
var Fbudjal: TFbudjal ;
begin
(*Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not JaiLeDroit(ccGenCreat,True) then Exit ;
              taModif,taModifEnSerie : if Not JaiLeDroit(ccGenModif,True) then Exit ;
   END ;*)
if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
Fbudjal:=TFbudjal.Create(Application) ;
try
  Fbudjal.Q:=Q ;
  Fbudjal.QuelAxe:=Axe ;
  Fbudjal.Lequel:=Compte ;
  Fbudjal.Mode:=Comment ;
  Fbudjal.LaPage:=QuellePage ;
  Fbudjal.ShowModal ;
  finally
  Fbudjal.Free ;
  end ;
Screen.Cursor:=crDefault ;
end ;

Procedure FicheBudjalMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
var Fbudjal: TFbudjal;
begin
(*Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not JaiLeDroit(ccGenCreat,True) then Exit ;
              taModif,taModifEnSerie : if Not JaiLeDroit(ccGenModif,True) then Exit ;
   END ;*)
if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
Fbudjal:=TFbudjal.Create(Application) ;
try
  Fbudjal.QuelAxe:=Axe ;
  Fbudjal.Lequel:=Lequel ;
  Fbudjal.Mode:=Comment ;
  Fbudjal.LaPage:=QuellePage ;
  Fbudjal.LesModif:=LesModif ;
  Fbudjal.ShowModal ;
  finally
  Fbudjal.Free ;
  end ;
Screen.Cursor:=crDefault ;
end ;

Function TFbudjal.CodeValide : Boolean ;
BEGIN
Result:=False ;
if QBudjal.state in [dsInsert] then
   BEGIN
   if QBudjalBJ_BUDJAL.AsString='' then
      BEGIN CurseurAuCode ; Msgbox.Execute(2,'','') ; Exit ; END ;
   if Presence('BUDJAL','BJ_BUDJAL',BJ_BUDJAL.Text) then
      BEGIN CurseurAuCode ; MsgBox.Execute(4,'','') ; Exit ; END ;
   END ;
Result:=True ;
END ;

Procedure TFbudjal.InitModifEnSerie(StModif : string) ;
var St,Champ, Valeur : string;
    i             : integer;
    B             : TBitBtn;
BEGIN
if QBudjal.State=dsBrowse then QBudjal.Edit ;
While StModif<>'' do
   BEGIN
   St:=ReadTokenSt(StModif);
   i:=Pos('=',St); if i>0 then Champ:=Trim(Copy(St,1,i-1));
   i:=Pos('"',St); if i>0 then St:=Trim(Copy(St,i+1,Length(St)));
   i:=Pos('"',St); if i>0 then Valeur:=Trim(Copy(St,1,i-1));
   AffecteLe(Champ,Valeur);
   END;
For i:=0 to HPB.ControlCount-1 do
   if HPB.Controls[i] is TBitBtn then
      BEGIN
      B:=TBitBtn(HPB.Controls[i]);
      if ((UpperCase(B.Name)<>'BVALIDER') and
          (UpperCase(B.Name)<>'BFERME') and
          (UpperCase(B.Name)<>'BAIDE')) then B.Enabled:=false;
      END;
END;

Procedure TFbudjal.AffecteLe(Champ, Valeur : string) ;
var C : TControl;
BEGIN
C:=TControl(FindComponent(Champ)) ;
if(C is TDBCheckBox)Or(C is THDBValComboBox)Or(C is TDBEdit)Or(C is THDBCpteEdit)then
   BEGIN
   QBudjal.FindField(Champ).AsString:=Valeur ; TEdit(C).Font.Color:=clRed;
   END else if C is THDBSpinEdit then
   BEGIN
   QBudjal.FindField(Champ).AsInteger:=StrToInt(Valeur) ; THDBSpinEdit(C).Font.Color:=clRed;
   END ;
END;

procedure TFbudjal.BAjoutClick(Sender: TObject);
begin
if (Not Degen) And (Not DeSect) then Exit ;
QTab.Close ; QTab.Sql.Clear ;
if Degen then
   BEGIN
   FListe.MultiFieds:='BG_BUDGENE;' ;
   QTab.Sql.Add('Select BG_BUDGENE, BG_LIBELLE From BUDGENE Where BG_ATTENTE="-" Order By BG_BUDGENE') ;
   ChangeSql(QTab) ; QTab.Open ;
   FListe.Columns.Items[0].Width:=90 ;
   FListe.Columns.Items[1].Width:=275 ;
   BGModifier:=True ;
   END else
   BEGIN
   FListe.MultiFieds:='BS_AXE;BS_BUDSECT;' ;
   QTab.Sql.Add('Select BS_BUDSECT,BS_LIBELLE,BS_AXE From BUDSECT '+
                'Where BS_AXE="'+QBudjalBJ_AXE.AsString+'" And BS_ATTENTE="-" Order By BS_BUDSECT') ;
   ChangeSql(QTab) ; QTab.Open ;
   FListe.Columns.Items[2].Title.Caption:=MsgBox.Mess[8] ;
   FListe.Columns.Items[2].Field:=FListe.Fields[2] ;
   FListe.Columns.Items[2].Title.Alignment:=taCenter ;
   FListe.Columns.Items[0].Width:=90 ;
   FListe.Columns.Items[1].Width:=240 ;
   FListe.Columns.Items[2].Width:=35 ;
   BSModifier:=True ;
   END ;
FListe.Columns.Items[0].Field:=FListe.Fields[0] ;
FListe.Columns.Items[1].Field:=FListe.Fields[1] ;
FListe.Columns.Items[0].Title.Caption:=MsgBox.Mess[6] ;
FListe.Columns.Items[1].Title.Caption:=MsgBox.Mess[7] ;
FListe.Columns.Items[0].Title.Alignment:=taCenter ;
FListe.Columns.Items[1].Title.Alignment:=taCenter ;
GbCpte.Visible:=False ; GBtab.Visible:=True ; HPB.Enabled:=False ;
PCaract.TabVisible:=FALSE ; PInfo.TabVisible:=FALSE ;
end;

procedure TFbudjal.BFerClick(Sender: TObject);
begin
QTab.Close ; GBTab.Visible:=False ; GbCpte.Visible:=True ;
Degen:=False ; DeSect:=False ; HPB.Enabled:=True ;
PCaract.TabVisible:=TRUE ; PInfo.TabVisible:=TRUE ; BPages.ActivePage:=PComplement ;
end;

procedure TFbudjal.BValClick(Sender: TObject);
Var i,NbLig : Integer ;
    St,St1 : String ;
    G : THGrid ;
begin
NbLig:=Fliste.NbSelected ;
if NbLig<=0 then BEGIN BFerClick(Nil) ; Exit ; END ;
St:='' ; Modifier:=False ;
for i:=0 to NbLig-1 do
    BEGIN
    Fliste.GotoLeBookMark(i) ;
    if Desect then St:=St+FListe.Fields[2].AsString+FListe.Fields[0].AsString+';'
              else St:=St+FListe.Fields[0].AsString+';' ;
    END ;
if Desect then G:=FListeS else G:=FListeG ;
While St<>'' do
    BEGIN
    St1:=ReadTokenSt(St) ;
    if G=FListeS then
       if Copy(St1,1,2)<>QBudjalBJ_AXE.AsString then Continue ;
    if DoublonExiste(St1,G) then Continue ;
    if G.Cells[0,G.RowCount-1]<>'' then G.RowCount:=G.RowCount+1 ;
    if G=FListeS then BEGIN G.Cells[1,G.RowCount-1]:=Copy(St1,1,2) ; Delete(St1,1,2) ; END ;
    G.Cells[0,G.RowCount-1]:=St1 ; Modifier:=True ;
    END ;
BFerClick(Nil) ;
end;

procedure TFbudjal.FListeGExit(Sender: TObject);
begin
if (BAjout.Focused) or (BUp.Focused) or (BDown.Focused) or (BDel.Focused) then
    BEGIN DeGen:=True ; DeSect:=False ; END else
    BEGIN DeGen:=False ; DeSect:=False ; END ;
end;

procedure TFbudjal.FListeSExit(Sender: TObject);
begin
if (BAjout.Focused) or (BUp.Focused) or (BDown.Focused) or (BDel.Focused) then
   BEGIN DeSect:=True ; DeGen:=False ; END else
   BEGIN DeSect:=False ; DeGen:=False ; END ;
end;

Procedure TFbudjal.LitBudgensBudsects(Stc : String ; G : THgrid ; PremFois : Boolean) ;
BEGIN
If PremFois Then G.VidePile(False) Else If Stc='' Then Exit ;
While Stc<>'' do
   BEGIN
   if G.Cells[0,G.RowCount-1]<>'' then G.RowCount:=G.RowCount+1 ;
   if G.Name='FListeS' then G.Cells[1,G.RowCount-1]:=QBudjalBJ_AXE.AsString ;
   G.Cells[0,G.RowCount-1]:=ReadTokenSt(Stc) ;
//   G.RowCount:=G.RowCount+1 ;
   END ;
//if G.RowCount>2 then G.RowCount:=G.RowCount-1 ;
//if (G.RowCount>2) And (Trim(G.Cells[0,G.RowCount-1])='') then G.RowCount:=G.RowCount-1 ;
END ;

Function TFbudjal.EcritBudgensBudsects(G : THgrid ; Var Ind : Integer) : String ;
Var St,St1 : String ;
    i,j : Integer ;
BEGIN
St:='' ; j:=0 ;
If Ind>0 Then
  for i:=Ind to G.RowCount-1 do
     BEGIN
     if (i=Ind) And (G.Cells[0,i]='') then Break ;
     St1:=St+G.Cells[0,i]+';' ;
     If Length(St1)<=2000 Then St:=St1 Else BEGIN j:=i ; Break ; END ;
     END ;
Result:=St ; Ind:=j ;
END ;

Function TFbudjal.TrouveCpteMvt(G : THGrid) : Boolean ;
Var QLoc : TQuery ;
    Champ,St : String ;
BEGIN
Result:=False ;
if Not IsMouvementer then Exit ;
if G.Name='FListeG' then Champ:='BE_BUDGENE' else Champ:='BE_BUDSECT' ;
QLoc:=OpenSql('Select '+Champ+' From BUDECR Where BE_BUDJAL="'+QBudjalBJ_BUDJAL.AsString+'" '+
              'And '+Champ+'="'+G.Cells[0,G.Row]+'"',True) ;
St:=QLoc.Fields[0].AsString ; Ferme(QLoc) ; SourisNormale ;
if St='' then Exit ;
Result:=True ;
END ;

procedure TFbudjal.BDelClick(Sender: TObject);
Var G : THGrid ;
    Cpte : String ;
begin
if (Not Degen) And (Not DeSect) then Exit ;
//if (G.Row=1) And (G.Cells[0,G.Row]='') then Exit ;
//if MsgBox.Execute(9,'','')<>mrYes then Exit ;
if Degen then BEGIN BGModifier:=True ; G:=FListeG ; Cpte:='CX_COMPTE' ; END
         else BEGIN BSModifier:=True ; G:=FListeS ; Cpte:='CX_SECTION'  ; END ;
if (G.Row=1) And (G.Cells[0,G.Row]='') then BEGIN BGModifier:=FALSE ; BSModifier:=FALSE ; Exit ; END ;
if MsgBox.Execute(9,'','')<>mrYes then BEGIN BGModifier:=FALSE ; BSModifier:=FALSE ; Exit ; END ;
If TrouveCpteMvt(G) Then BEGIN MsgBox.Execute(13,'','') ; Exit ; END ;
Modifier:=True ;
ExecuteSql('Delete From CROISCPT Where CX_JAL="'+QBudJalBJ_BUDJAL.AsString+'" And '+
           'CX_TYPE="BUD" And '+Cpte+'="'+G.Cells[0,G.Row]+'"') ;
If QBudJalBJ_CATEGORIE.AsString<>'' Then
  ExecuteSql('Delete From CROISCPT Where CX_JAL="'+QBudJalBJ_CATEGORIE.AsString+'" And '+
             'CX_TYPE="BUD" And '+Cpte+'="'+G.Cells[0,G.Row]+'"') ;

G.DeleteRow(G.Row) ;
if G.RowCount=1 then BEGIN G.RowCount:=G.RowCount+1 ; G.FixedRows:=1 ; G.Row:=G.RowCount-1 ; END ;
Degen:=False ; DeSect:=False ;
end;

procedure TFbudjal.BUpClick(Sender: TObject);
Var G : THGrid ;
begin
if (Not Degen) And (Not DeSect) then Exit ;
Modifier:=True ;
if Degen then G:=FListeG else G:=FListeS ;
if G.Row=1 then BEGIN G.SetFocus ; Exit ; END ;
G.ExchangeRow(G.Row,G.Row-1) ; G.Row:=G.Row-1 ; G.SetFocus ;
end;

procedure TFbudjal.BDownClick(Sender: TObject);
Var G : THGrid ;
begin
if (Not Degen) And (Not DeSect) then Exit ;
Modifier:=True ;
if Degen then G:=FListeG else G:=FListeS ;
if G.Row=G.RowCount-1 then BEGIN G.SetFocus ; Exit ; END ;
G.ExchangeRow(G.Row,G.Row+1) ; G.Row:=G.Row+1 ; G.SetFocus ;
end;

Function TFbudjal.CpteAttenteOk ( FicBase : TFichierBase) : Boolean ;
Var St : String ;
    C  : THDBCpteEdit ;
BEGIN
Result:=False ; C:=NIL ;
Case FicBase of
     fbBudGen : BEGIN St:=QBudjalBJ_GENEATTENTE.AsString ; C:=BJ_GENEATTENTE ; END ;
     fbBudSec1..fbBudSec5 : BEGIN St:=QBudjalBJ_SECTATTENTE.AsString ; C:=BJ_SECTATTENTE ; END ;
   End ;
if St='' then
   BEGIN
   BPages.ActivePage:=PCaract ; C.SetFocus ;
   if FicBase=fbBudGen then MsgBox.Execute(10,'','')
                       else MsgBox.Execute(14,'','') ;
   Exit ;
   END ;
if C.Existeh<=0 then
   BEGIN
   BPages.ActivePage:=PCaract ; C.SetFocus ;
   if FicBase=fbBudGen then MsgBox.Execute(11,'','')
                       else MsgBox.Execute(15,'','') ;
   Exit ;
   END ;
Result:=True ;
END ;

Function EstMouvementeBudjal(Lejal : String) : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select BJ_BUDJAL From BUDJAL Where BJ_BUDJAL="'+Lejal+'" And '+
              '(Exists(Select BE_BUDJAL From BUDECR WHere BE_BUDJAL="'+Lejal+'"))',True) ;
Result:=Not QLoc.Eof ; Ferme(QLoc) ;
END ;

Procedure TFbudjal.GriseDegriseControl ;
BEGIN
BJ_AXE.Enabled:=Not IsMouvementer ; BJ_COMPTEURNORMAL.Enabled:=Not IsMouvementer ;
GroupBox3.Enabled:=Not IsMouvementer ; GroupBox1.Enabled:=Not IsMouvementer ;
BCopieComplement.Enabled:=Not IsMouvementer ;
//BDel.Enabled:=Not IsMouvementer ;
END ;

Function TFbudjal.DoublonExiste(St : String ; G : THGrid) : Boolean ;
Var i : Integer ;
    StTemp : String ;
BEGIN
Result:=False ;
if G.Cells[0,1]='' then Exit ;
if G=FListeS then StTemp:=Copy(St,3,Length(St))
             else StTemp:=St ;
for i:=1 to G.RowCount-1 do
    if G.Cells[0,i]=StTemp then BEGIN Result:=True ; Exit ; END ;
END ;

procedure TFbudjal.BJ_EXODEBChange(Sender: TObject);
begin
if BJ_EXODEB.Value>BJ_EXOFIN.Value then BJ_EXOFIN.Value:=BJ_EXODEB.Value ;
ListePeriode(BJ_EXODEB.Value,BJ_PERDEB.Items,BJ_PERDEB.Values,TRUE) ;
PositionnePerDeb ;
end;

procedure TFbudjal.BJ_EXOFINChange(Sender: TObject);
begin
if BJ_EXOFIN.Value<BJ_EXODEB.Value then BJ_EXODEB.Value:=BJ_EXOFIN.Value ;
ListePeriode(BJ_EXOFIN.Value,BJ_PERFIN.Items,BJ_PERFIN.Values,FALSE) ;
PositionnePerFin ;
end;

procedure TFbudjal.BJ_AXEChange(Sender: TObject);
Var UnFb : TFichierBase ;
begin
Lefb:=AxeTofbBud(BJ_AXE.Value) ;
UnFb:=AxeToFb(BJ_AXE.Value) ;
Case Lefb of
     fbBudsec1 : BEGIN BJ_SECTATTENTE.ZoomTable:=tzBudSecAtt1 ; BJ_CATEGORIE.DataType:='ttCatJalBud1' ; END ;
     fbBudsec2 : BEGIN BJ_SECTATTENTE.ZoomTable:=tzBudSecAtt2 ; BJ_CATEGORIE.DataType:='ttCatJalBud2' ; END ;
     fbBudsec3 : BEGIN BJ_SECTATTENTE.ZoomTable:=tzBudSecAtt3 ; BJ_CATEGORIE.DataType:='ttCatJalBud3' ; END ;
     fbBudsec4 : BEGIN BJ_SECTATTENTE.ZoomTable:=tzBudSecAtt4 ; BJ_CATEGORIE.DataType:='ttCatJalBud4' ; END ;
     fbBudsec5 : BEGIN BJ_SECTATTENTE.ZoomTable:=tzBudSecAtt5 ; BJ_CATEGORIE.DataType:='ttCatJalBud5' ; END ;
   End ;
If Mode<>taConsult Then
   BEGIN
   TBJ_Categorie.Enabled:=VH^.Cpta[UnFb].Structure ; BJ_Categorie.Enabled:=VH^.Cpta[UnFb].Structure ;
   TBJ_SousPlan.Enabled:=VH^.Cpta[UnFb].Structure ; BJ_SousPlan.Enabled:=VH^.Cpta[UnFb].Structure ;
   BZoomSousPlan.Enabled:=VH^.Cpta[UnFb].Structure ;
   END ;
end;

procedure TFbudjal.BJ_AXEClick(Sender: TObject);
begin
if QBudjal.State<>dsInsert then
   BEGIN
   QBudjalBJ_BUDSECTS.AsString:='' ;
   QBudjalBJ_BUDSECTS2.AsString:='' ;
   FListeS.VidePile(False) ;
   END ;
end;

procedure TFbudjal.BJ_PERDEBChange(Sender: TObject);
begin
if StrToDate(BJ_PERDEB.Value)>StrToDate(BJ_PERFIN.Value) then BJ_PERFIN.Value:=DateToStr(FinDeMois(StrToDate(BJ_PERDEB.Value))) ;
end;

procedure TFbudjal.BJ_PERFINChange(Sender: TObject);
begin
if StrToDate(BJ_PERFIN.Value)<StrToDate(BJ_PERDEB.Value) then BJ_PERDEB.Value:=DateToStr(DebutDeMois(StrToDate(BJ_PERFIN.Value))) ;
end;

Function TFbudjal.ExoOuv : Boolean ;
Var QLoc : TQuery ;
BEGIN
Result:=False ;
QLoc:=OpenSql('Select EX_EXERCICE From EXERCICE Where EX_ETATBUDGET="OUV"',True) ;
if QLoc.Eof then BEGIN Ferme(QLoc) ; Exit ; END ;
While Not QLoc.Eof do
    BEGIN
    if QLoc.Fields[0].AsString=QBudjalBJ_EXODEB.AsString then
       BEGIN Ferme(QLoc) ; Result:=True ; Exit ; END ;
    QLoc.Next ;
    END ;
Ferme(QLoc) ;
END ;

Procedure TFbudjal.PositionnePerDeb ;
BEGIN
BJ_PERDEB.ItemIndex:=0 ;
QBudjalBJ_PERDEB.AsString:=BJ_PERDEB.Value ;
END ;

Procedure TFbudjal.PositionnePerFin ;
BEGIN
BJ_PERFIN.ItemIndex:=BJ_PERFIN.Items.Count-1 ;
QBudjalBJ_PERFIN.AsString:=BJ_PERFIN.Value ;
END ;

procedure TFbudjal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
   BEGIN
   AvertirTable('ttBudJal') ; AvertirTable('ttBudJalSais') ;
   AvertirTable('TTBUDJALSANSCAT') ; AvertirTable('TTCATJALBUD') ;
   END ;
end;

procedure TFbudjal.QTabAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFbudjal.FormKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ;
if Vide then
   BEGIN
   Case Key of
        VK_F3  : BPrevClick(Nil) ;
        VK_F4  : BNextClick(Nil) ;
        VK_F10 : BValiderClick(Nil) ;
        VK_RETURN :BEGIN
                   if ActiveControl is TCustomMemo then Exit ;
                   FindNextControl(ActiveControl,True,True,False).SetFocus ;
                   END ;
     END ;
   END ;
end;

procedure TFbudjal.GbCpteEnter(Sender: TObject);
begin
if (Not Degen) And (Not DeSect) then BEGIN DeGen:=TRUE ; FListeG.SetFocus ; END ;
end;

procedure TFbudjal.QBudjalAfterPost(DataSet: TDataSet);
begin Favertir:=True ; end;

procedure TFbudjal.QBudjalAfterDelete(DataSet: TDataSet);
begin Favertir:=True ; end;

procedure TFbudjal.BJ_NATJALChange(Sender: TObject);
begin
Case BJ_NATJAL.Value[1] of
     'C' : QBudjalBJ_SENS.AsString:='D' ;
     'P' : QBudjalBJ_SENS.AsString:='C' ;
   End ;
end;

Function TFbudjal.CoherenceSection : Boolean ;
Var i : Integer ;
    QLoc : TQuery ;
    Sql : String ;
BEGIN
Result:=False ;
if FListeS.Cells[0,FListeS.RowCount-1]='' then BEGIN Result:=True ; Exit ; END ;
Sql:='Select BS_BUDSECT From BUDSECT Where BS_BUDSECT=:CpteSec And BS_AXE=:Ax' ;
QLoc:=PrepareSQL(Sql,TRUE) ;
for i:=1 to FListeS.RowCount-1 do
    BEGIN
    QLoc.Close ;
    QLoc.Params[0].AsString:=FListeS.Cells[0,i] ;
    QLoc.Params[1].AsString:=BJ_AXE.Value ;
    QLoc.Open ;
    if QLoc.Eof then
       BEGIN
       Ferme(QLoc) ; BPages.ActivePage:=PComplement ;
       if FListeS.CanFocus then FListeS.Focused ;
       MsgBox.Execute(16,'','') ; Exit ;
       END ;
    END ;
Ferme(QLoc) ; Result:=True ;
END ;

procedure TFbudjal.BZoomSousPlanClick(Sender: TObject);
Var LaCatBud : TUneCatBud ;
    St,Lib,Abr : String ;
begin
BJ_SOUSPLAN.SetFocus ; LaCatBud:=QuelleCatBud(BJ_CATEGORIE.Value) ;
If LaCatBud.Code<>'' Then
   BEGIN
   St:=ChoisirSousSectionBudget(LaCatBud,BJ_SOUSPLAN.Text) ;
   If (St<>BJ_SOUSPLAN.Text) And (St<>'') Then
      BEGIN
      if QBudjal.State=dsBrowse then QBudjal.Edit ;
      QBudJalBJ_SOUSPLAN.AsString:=St ;
      LibelleCatBud(QBudJalBJ_CATEGORIE.AsString,QBudJalBJ_SOUSPLAN.AsString,Lib,Abr) ;
      QBudJalBJ_LIBELLE.AsString:=Lib ;
      QBudJalBJ_ABREGE.AsString:=Abr ;
      END ;
   END ;
end;

procedure TFbudjal.BJ_SOUSPLANKeyPress(Sender: TObject; var Key: Char);
begin
If key<>#8 Then Key:=#0 ;
end;

procedure TFbudjal.BCopieComplementClick(Sender: TObject);
begin
PopZoom(BCopieComplement,POPZ) ;
end;

procedure TFbudjal.BCopieListe(Qui : Integer) ;
Var St : String ;
    Q : TQuery ;
begin
If BJ_AXE.Value='' Then Exit ;
St:=Choisir(MsgBox.Mess[17],'BUDJAL','BJ_LIBELLE','BJ_BUDJAL',
            'BJ_AXE="'+BJ_AXE.Value+'" AND BJ_BUDJAL<>"'+QBudJalBJ_BUDJAL.AsString+'"','BJ_BUDJAL');
if St='' then exit;
if MsgBox.Execute(18,'','')<>mrYes then Exit ;
Modifier:=True ;
ExecuteSql('Delete From CROISCPT Where CX_JAL="'+QBudJalBJ_BUDJAL.AsString+'" And CX_TYPE="BUD"') ;
If (Qui=0) Or (Qui=2) Then FListeG.VidePile(False) ;
If (Qui=1) Or (Qui=2) Then FListeS.VidePile(False) ;
Q:=OpenSQL('Select BJ_CATEGORIE,BJ_BUDGENES, BJ_BUDSECTS, BJ_BUDGENES2, BJ_BUDSECTS2 From BUDJAL Where BJ_BUDJAL="'+St+'"',TRUE) ;
If (QBudJalBJ_CATEGORIE.AsString<>'') And
   (QBudJalBJ_CATEGORIE.AsString<>Q.FindField('BJ_CATEGORIE').AsString) Then
   ExecuteSql('Delete From CROISCPT Where CX_JAL="'+QBudJalBJ_CATEGORIE.AsString+'" And CX_TYPE="BUD"') ;
If Not Q.Eof Then
   BEGIN
   If (Qui=0) Or (Qui=2) Then
      BEGIN
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES').AsString),FListeG,TRUE) ;
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES2').AsString),FListeG,FALSE) ;
      END ;
   If (Qui=1) Or (Qui=2) Then
      BEGIN
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS').AsString),FListeS,TRUE) ;
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS2').AsString),FListeS,FALSE) ;
      END ;
   END ;
Ferme(Q) ;
If (Qui=0) Or (Qui=2) Then
   BEGIN
   if FListeG.RowCount=1 then BEGIN FListeG.RowCount:=FListeG.RowCount+1 ; FListeG.FixedRows:=1 ; FListeG.Row:=FListeG.RowCount-1 ; END ;
   END ;
If (Qui=1) Or (Qui=2) Then
   BEGIN
   if FListeS.RowCount=1 then BEGIN FListeS.RowCount:=FListeS.RowCount+1 ; FListeS.FixedRows:=1 ; FListeS.Row:=FListeS.RowCount-1 ; END ;
   END ;
end;

procedure TFbudjal.GenereCroisement(CodeACreer,CodeChoisi : String) ;
Var Q,Q1 : TQuery ;
BEGIN
EnableControls(Self,False) ;
BeginTrans ;
Q:=OpenSQL('SELECT * FROM CROISCPT WHERE CX_TYPE="BUD" AND CX_JAL="'+CodeChoisi+'"',TRUE) ;
Q1:=OpenSQL('SELECT * FROM CROISCPT WHERE CX_TYPE="Ed#"',False) ;
InitMove(500,MsgBox.Mess[21]) ;
While Not Q.Eof Do
   BEGIN
   MoveCur(FALSE) ;
   Q1.Insert ; InitNew(Q1) ;
   Q1.FindField('CX_TYPE').AsString:=Q.FindField('CX_TYPE').AsString ;
   Q1.FindField('CX_JAL').AsString:=CodeACreer ;
   Q1.FindField('CX_COMPTE').AsString:=Q.FindField('CX_COMPTE').AsString ;
   Q1.FindField('CX_SECTION').AsString:=Q.FindField('CX_SECTION').AsString ;
   Q1.FindField('CX_INFO').AsString:=Q.FindField('CX_INFO').AsString ;
   Q1.Post ;
   Q.Next ;
   END ;
Ferme(Q1) ; Ferme(Q) ; FiniMove ;
CommitTrans ;
EnableControls(Self,TRUE) ;
END ;

procedure TFbudjal.BCopierJalClick(Sender: TObject);
Var St : String ;
    Q : TQuery ;
begin
If Trim(BJ_BUDJAL.Text)='' Then
   BEGIN
   MsgBox.Execute(20,'','') ; BJ_BUDJAL.SetFocus ; Exit ;
   END ;
St:=Choisir(MsgBox.Mess[17],'BUDJAL','BJ_LIBELLE','BJ_BUDJAL','BJ_BUDJAL<>"'+QBudJalBJ_BUDJAL.AsString+'"','BJ_BUDJAL');
if St='' then exit;
if MsgBox.Execute(19,'','')<>mrYes then Exit ;
FListeG.VidePile(False) ; FListeS.VidePile(False) ;
Q:=OpenSQL('Select * From BUDJAL Where BJ_BUDJAL="'+St+'"',TRUE) ;
If Not Q.Eof Then
   BEGIN
   QBudJalBJ_AXE.AsString:=Q.FindField('BJ_AXE').AsString ;
   QBudJalBJ_NATJAL.AsString:=Q.FindField('BJ_NATJAL').AsString ;
   QBudJalBJ_CATEGORIE.AsString:=Q.FindField('BJ_CATEGORIE').AsString ;
   QBudJalBJ_SENS.AsString:=Q.FindField('BJ_SENS').AsString ;
   QBudJalBJ_COMPTEURNORMAL.AsString:=Q.FindField('BJ_COMPTEURNORMAL').AsString ;
   QBudJalBJ_GENEATTENTE.AsString:=Q.FindField('BJ_GENEATTENTE').AsString ;
   QBudJalBJ_SECTATTENTE.AsString:=Q.FindField('BJ_SECTATTENTE').AsString ;
   QBudJalBJ_EXODEB.AsString:=Q.FindField('BJ_EXODEB').AsString ;
   QBudJalBJ_EXOFIN.AsString:=Q.FindField('BJ_EXOFIN').AsString ;
   BJ_EXODEBChange(Nil) ; BJ_EXOFinChange(Nil) ;
   QBudJalBJ_PERDEB.AsString:=Q.FindField('BJ_PERDEB').AsString ;
   QBudJalBJ_PERFIN.AsString:=Q.FindField('BJ_PERFIN').AsString ;
(*
   LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES').AsString),FListeG) ;
   LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS').AsString),FListeS) ;
*)
   LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES').AsString),FListeG,TRUE) ;
   LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES2').AsString),FListeG,FALSE) ;
   LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS').AsString),FListeS,TRUE) ;
   LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS2').AsString),FListeS,FALSE) ;
   If QBudJalBJ_CATEGORIE.AsString='' Then GenereCroisement(BJ_BUDJAL.Text,St)  ;
   END ;
Ferme(Q) ;
if FListeG.RowCount=1 then BEGIN FListeG.RowCount:=FListeG.RowCount+1 ; FListeG.FixedRows:=1 ; FListeG.Row:=FListeG.RowCount-1 ; END ;
if FListeS.RowCount=1 then BEGIN FListeS.RowCount:=FListeS.RowCount+1 ; FListeS.FixedRows:=1 ; FListeS.Row:=FListeS.RowCount-1 ; END ;
end;

procedure TFbudjal.BJ_CATEGORIEChange(Sender: TObject);
begin
QBudjalBJ_SOUSPLAN.AsString:='' ;
end;

Procedure TFbudjal.CopieCatSurJournaux ;
var St,StW,ListeC,BudG,BudG2,BudS,BudS2 : String ;
    Ind : Integer ;
    First : boolean ;
BEGIN
if QBudjalBJ_CATEGORIE.AsString='' then Exit ;
if not BGModifier and not BSModifier then Exit ;
if MsgBox.Execute(22,'','')<>mrYes then Exit ;
StW:='' ;
StW:='BJ_CATEGORIE="'+QBudjalBJ_CATEGORIE.AsString+'" AND BJ_BUDJAL<>"'+QBudJalBJ_BUDJAL.AsString+'"' ; ;
ListeC:=ChoisirMulti (MsgBox.Mess[23],'BUDJAL','BJ_BUDJAL','BJ_LIBELLE',StW,'BJ_BUDJAL','') ;
Ind:=1 ; BudG2:='' ;
BudG:=EcritBudgensBudsects(FListeG,Ind) ;
If Ind>0 Then BudG2:=EcritBudgensBudsects(FListeG,Ind) ;
Ind:=1 ; BudS2:='' ;
BudS:=EcritBudgensBudsects(FListeS,Ind) ;
If Ind>0 Then BudS2:=EcritBudgensBudsects(FListeS,Ind) ;
St:='UPDATE BUDJAL SET ' ;
if BGModifier then St:=St+'BJ_BUDGENES="'+BudG+'",BJ_BUDGENES2="'+BudG2+'"' ;
if BSModifier then
  BEGIN
  if BGModifier then St:=St+',' ;
  St:=St+'BJ_BUDSECTS="'+BudS+'",BJ_BUDSECTS2="'+BudS2+'"' ;
  END ;
St:=St+' WHERE BJ_BUDJAL<>"'+QBudJalBJ_BUDJAL.AsString+'" AND BJ_CATEGORIE="'+QBudjalBJ_CATEGORIE.AsString+'"' ;
StW:=ReadTokenSt(ListeC) ; First:=True ;
While StW<>'' do
  BEGIN
  if First then St:=St+' AND ('
           else St:=St+' OR ' ;
  First:=False ;
  St:=St+' BJ_BUDJAL="'+StW+'"' ;
  StW:=ReadTokenSt(ListeC) ;
  END ;
if not First then St:=St+')' ;
ExecuteSQL(St) ;
BSModifier:=False ; BGModifier:=False ;
END ;

procedure TFbudjal.BCopieCptClick(Sender: TObject);
begin
BCopieListe(0) ;
end;

procedure TFbudjal.BCopieSectClick(Sender: TObject);
begin
BCopieListe(1) ;
end;

procedure TFbudjal.BCopieDeuxClick(Sender: TObject);
begin
BCopieListe(2) ;
end;

procedure TFbudjal.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;

end;

procedure TFbudjal.FListeGSorted(Sender: TObject);
begin
FListeG.SortedCol:=0 ; Modifier:=True ;
end;

procedure TFbudjal.FListeSSorted(Sender: TObject);
begin
FListeS.SortedCol:=0 ; Modifier:=True ;
end;

end.
