{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par BUDSECT_TOM.PAS
Mots clefs ... : 
*****************************************************************}
unit Budsect;

interface

uses
    WinTypes, WinProcs, Classes, Graphics, Forms, Controls, StdCtrls, Tabs,
    Buttons, ExtCtrls, Grids, Mask,  DB, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} DBCtrls,HEnt1,
    Hctrls, HDB, DBGrids, Hqry, TabNotBk, Dialogs, Spin, SysUtils, Ent1,
    ComCtrls, hmsgbox,hcompte, Messages, Menus, HRichEdt, HRichOLE,
    HSysMenu,MajTable, HRegCpte, ParamSoc,UtilPGI   ,UentCommun, ADODB;

Procedure FicheBudsect(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
Procedure FicheBudsectMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
Function SectionRetrieBud(CodRupt : String ; Axe : String ; ListeCodes : TStringList) : TSectRetri ;

type
  TFbudsect = class(TForm)
    BPages             : TPageControl;
    HPB                : TPanel;
    BAide: THBitBtn;
    BAnnuler: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    BFirst: THBitBtn;
    BPrev: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    BInsert: THBitBtn;
    BFerme: THBitBtn;
    FAutoSave          : TCheckBox;
    PCaract            : TTabSheet;
    PComplement        : TTabSheet;
    PInfo              : TTabSheet;
    HGBDates           : TGroupBox;
    BS_FERME: TDBCheckBox;
    TBS_BLOCNOTE: TGroupBox;
    BS_BLOCNOTE: THDBRichEditOLE;
    SBudsect: TDataSource;
    TBS_DATECREATION: THLabel;
    BS_DATECREATION: TDBEdit;
    TBS_DATEMODIF: THLabel;
    BS_DATEMODIF: TDBEdit;
    TBS_DATEOUVERTURE: THLabel;
    BS_DATEOUVERTURE: TDBEdit;
    TBS_DATEFERMETURE: THLabel;
    BS_DATEFERMETURE: TDBEdit;
    DBNav              : TDBNavigator;
    MsgBox             : THMsgBox;
    TBS_BUDSECT: THLabel;
    BS_BUDSECT: TDBEdit;
    TBS_LIBELLE: THLabel;
    BS_LIBELLE: TDBEdit;
    TBS_SENS: THLabel;
    BS_SENS: THDBValComboBox;
    TBS_ABREGE: THLabel;
    BS_ABREGE: TDBEdit;
    QBudsect: TQuery;
    HMTrad: THSystemMenu;
    ZL: TTabSheet;
    TBS_TABLE0: THLabel;
    TBS_TABLE1: THLabel;
    TBS_TABLE2: THLabel;
    TBS_TABLE3: THLabel;
    TBS_TABLE4: THLabel;
    TBS_TABLE5: THLabel;
    TBS_TABLE6: THLabel;
    TBS_TABLE7: THLabel;
    TBS_TABLE8: THLabel;
    TBS_TABLE9: THLabel;
    BS_TABLE0: THDBCpteEdit;
    BS_TABLE1: THDBCpteEdit;
    BS_TABLE2: THDBCpteEdit;
    BS_TABLE3: THDBCpteEdit;
    BS_TABLE4: THDBCpteEdit;
    BS_TABLE5: THDBCpteEdit;
    BS_TABLE6: THDBCpteEdit;
    BS_TABLE7: THDBCpteEdit;
    BS_TABLE8: THDBCpteEdit;
    BS_TABLE9: THDBCpteEdit;
    BS_SIGNE: THDBValComboBox;
    TBS_SIGNE: TLabel;
    BS_REPORTDISPO: THDBValComboBox;
    BS_PARENT: TDBEdit;
    TBS_PARENT: TLabel;
    BS_ATTENTE: TDBCheckBox;
    BS_HT: TDBCheckBox;
    TBS_REPORTDISPO: TLabel;
    BS_FORMULE: TDBEdit;
    TBS_FORMULE: TLabel;
    BS_CONFIDENTIEL: TDBCheckBox;
    FListe: THGrid;
    QBudsectBS_BUDSECT: TStringField;
    QBudsectBS_LIBELLE: TStringField;
    QBudsectBS_ABREGE: TStringField;
    QBudsectBS_SECTIONRUB: TStringField;
    QBudsectBS_EXCLURUB: TStringField;
    QBudsectBS_SIGNE: TStringField;
    QBudsectBS_REPORTDISPO: TStringField;
    QBudsectBS_PARENT: TStringField;
    QBudsectBS_ATTENTE: TStringField;
    QBudsectBS_FORMULE: TStringField;
    QBudsectBS_DATECREATION: TDateTimeField;
    QBudsectBS_DATEMODIF: TDateTimeField;
    QBudsectBS_DATEOUVERTURE: TDateTimeField;
    QBudsectBS_DATEFERMETURE: TDateTimeField;
    QBudsectBS_FERME: TStringField;
    QBudsectBS_SENS: TStringField;
    QBudsectBS_BLOCNOTE: TMemoField;
    QBudsectBS_CONFIDENTIEL: TStringField;
    QBudsectBS_HT: TStringField;
    QBudsectBS_UTILISATEUR: TStringField;
    QBudsectBS_SOCIETE: TStringField;
    QBudsectBS_BLOQUANT: TStringField;
    QBudsectBS_CREERPAR: TStringField;
    QBudsectBS_EXPORTE: TStringField;
    QBudsectBS_TABLE0: TStringField;
    QBudsectBS_TABLE1: TStringField;
    QBudsectBS_TABLE2: TStringField;
    QBudsectBS_TABLE3: TStringField;
    QBudsectBS_TABLE4: TStringField;
    QBudsectBS_TABLE5: TStringField;
    QBudsectBS_TABLE6: TStringField;
    QBudsectBS_TABLE7: TStringField;
    QBudsectBS_TABLE8: TStringField;
    QBudsectBS_TABLE9: TStringField;
    QBudsectBS_AXE: TStringField;
    TBS_AXE: TLabel;
    BS_AXE: THDBValComboBox;
    QBudsectBS_RUB: TStringField;
    TBS_RUB: THLabel;
    BS_RUB: TDBEdit;
    procedure FormCreate(Sender: TObject);
    procedure SBudsectDataChange(Sender: TObject; Field: TField);
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
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint;
      var Cancel: Boolean);
    procedure BS_AXEChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private { Private declarations }
    Q : TQuery ;
    Mode : TActionFiche ;
    LaPage : Integer ;
    Lequel,LeCompte,LeAxe : String ;
    LesModif : String ;
    AxeSerie,QuelAxe : String ;
    MemoAxe : String ;
    SensDefaut,SigneDefaut,AxeDefaut : String ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Procedure CurseurAuCode ;
    Procedure CurseurAuCodeRub ;
    Procedure CurseurAuLibelle ;
    Function  EnregOK : boolean ;
    Procedure ChargeEnreg ;
    Function  CodeValide : Boolean ;
    Function  Supprime : Boolean ;
    Procedure AffecteLe(Champ, Valeur : string) ;
    Procedure InitModifEnSerie(StModif : string) ;
    Function  OkConfidentiel : Boolean ;
    Function  Grid2St ( ACol : Integer ) : String ;
    Procedure LitCompteExclu(St : String ; ACol : Integer) ;
    Procedure EcritCompteExclu ;
    Procedure RempliGrid ;
    Function  PresenceRub : Boolean ;
    Function  VerifCoherenceTL : Boolean ;
  public  { Public declarations }
  end;

    Function  EstMouvementeBudsect(LeCpte,Laxe : String)   : Boolean ;  //ECRITURE

implementation

{$R *.DFM}

{$IFNDEF IMP}
uses QRBudSec ;
{$ENDIF}

Procedure FicheBudsect(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
var FBudsect: TFbudsect ;
begin
(*Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not JaiLeDroit(ccGenCreat,True) then Exit ;
              taModif,taModifEnSerie : if Not JaiLeDroit(ccGenModif,True) then Exit ;
   END ;*)
if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
FBudsect:=TFbudsect.Create(Application) ;
try
  FBudsect.Q:=Q ;
  FBudsect.QuelAxe:=Axe ;
  FBudsect.Lequel:=Compte ;
  FBudsect.Mode:=Comment ;
  FBudsect.LaPage:=QuellePage ;
  FBudsect.ShowModal ;
  finally
  FBudsect.Free ;
  end ;
Screen.Cursor:=crDefault ;
end ;

Procedure FicheBudsectMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
var FBudsect: TFbudsect;
begin
(*Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not JaiLeDroit(ccGenCreat,True) then Exit ;
              taModif,taModifEnSerie : if Not JaiLeDroit(ccGenModif,True) then Exit ;
   END ;*)
if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
FBudsect:=TFbudsect.Create(Application) ;
try
  FBudsect.QuelAxe:=Axe ;
  FBudsect.Lequel:=Lequel ;
  FBudsect.Mode:=Comment ;
  FBudsect.LaPage:=QuellePage ;
  FBudsect.LesModif:=LesModif ;
  FBudsect.ShowModal ;
  finally
  FBudsect.Free ;
  end ;
Screen.Cursor:=crDefault ;
end ;

procedure TFbudsect.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; Q:=NIL ;
LeCompte:='aze' ; LeAxe:='aze' ;
end;

procedure TFbudsect.SBudsectDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not(QBudsect.State in [dsEdit,dsInsert]) ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not QBudsect.BOF;
   DnEnable := Enabled and not QBudsect.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
   if ((Field.FieldName='BS_LIBELLE') and (BS_ABREGE.Field.AsString='')) then
      BS_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   if ((Field.FieldName='BS_BUDSECT') and (BS_RUB.Field.AsString='')) then
      BS_RUB.Field.AsString:=Copy(Field.AsString,1,5) ;
   END ;
end;

procedure TFbudsect.FormShow(Sender: TObject);
begin
MakeZoomOLE(Handle) ;
AxeSerie:='' ; RecupWhereSQL(Q,QBudsect) ; MemoAxe:='' ;
if ((Q=NIL) and (Lequel<>'')) then QBudsect.SQL.Add('Where BS_BUDSECT="'+Lequel+'"') ;
if ((Q=NIL) and (Lequel<>'') and (QuelAxe<>'')) then QBudsect.SQL.Add('And BS_AXE="'+QuelAxe+'" ') ;
if ((Q=NIL) and (Lequel<>'')) then ChangeSQL(QBudsect) ;
ChangeSizeMemo(QBudsectBS_BLOCNOTE) ;
QBudsect.Open ;
if(Lequel<>'')And((Mode in [taCreat..taCreatOne])=False) then
   BEGIN
   if Not QBudsect.Locate('BS_BUDSECT',Lequel,[]) then
     BEGIN MessageAlerte(MsgBox.Mess[5]) ; PostMessage(Handle,WM_CLOSE,0,0);  Exit ; END ;
   END ;
if Not OkConfidentiel then Exit ;
BPages.ActivePage:=BPages.Pages[LaPage] ;
Case Mode Of
     taConsult           : BEGIN
                           FicheReadOnly(Self) ;
                           FListe.Enabled:=True ; FListe.Options:=FListe.Options-[goEditing]+[goRowSelect] ;
                           Exit ;
                           END ;
     taCreat..taCreatOne : BEGIN Bouge(nbInsert) ; BS_BUDSECT.Text:=Lequel ; BAnnuler.Enabled:=False ; END ;
     taModif             : BEGIN END ;
     taModifEnSerie      : InitModifEnSerie(LesModif);
     end ;
SensDefaut:='' ; SigneDefaut:='' ; AxeDefaut:='A1' ;
if(QBudsect.Eof) then Bouge(nbInsert) ;
end;

Procedure TFbudsect.ChargeEnreg ;
BEGIN
if ((LeCompte=QBudsectBS_BUDSECT.AsString) and (LeAxe=QBudsectBS_AXE.AsString)) then
   BEGIN
   if Mode=taConsult then
      BEGIN
      FicheReadOnly(Self) ;
      FListe.Enabled:=True ; FListe.Options:=FListe.Options-[goEditing]+[goRowSelect] ;
      END ;
   Exit ;
   END ;
LeCompte:=QBudsectBS_BUDSECT.AsString ; LeAxe:=QBudsectBS_AXE.AsString ;

InitCaption(Self,BS_BUDSECT.text,BS_LIBELLE.text) ;
RempliGrid ;
if Mode=taConsult then
   BEGIN
   FicheReadOnly(Self) ;
   FListe.Enabled:=True ; FListe.Options:=FListe.Options-[goEditing]+[goRowSelect] ;
   Exit ;
   END ;
BS_BUDSECT.Enabled:=FALSE ; MemoAxe:=QBudsectBS_AXE.AsString ;
END ;

Function TFbudsect.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : BEGIN
                      if Not EnregOK then Exit ;
                      If mode=taCreatEnSerie Then
                         BEGIN
                         SensDefaut:=QBudsectBS_SENS.AsString ;
                         SigneDefaut:=QBudsectBS_SIGNE.AsString ;
                         AxeDefaut:=QBudsectBS_AXE.AsString ;
                         END ;
                      END ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if Mode in [taCreat..taCreatOne]then AxeSerie:=QBudsectBS_AXE.AsString ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[5]) ;
Result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFbudsect.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
Result:=FALSE  ;
if QBudsect.Modified then
   BEGIN
   if (Mode in [taCreat..taCreatOne]) And
      (QBudsectBS_BUDSECT.asString='') then Rep:=mrNo
   else if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
     mrYes : if not Bouge(nbPost) then Exit ;
     mrNo :  if not Bouge(nbCancel) then Exit ;
     mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFbudsect.EnregOK : boolean ;
BEGIN
Result:=False ;
if QBudsect.state in [dsEdit,dsInsert]=False then Exit ;
if Not CodeValide then Exit ;
if QBudsect.state in [dsEdit,dsInsert] then
   BEGIN
   if QBudsectBS_LIBELLE.asString='' then
      BEGIN Msgbox.Execute(3,'','') ; CurseurAuLibelle ; Exit ; END ;
   if QBudsectBS_RUB.asString='' then
      BEGIN Msgbox.Execute(11,'','') ; CurseurAuCodeRub ; Exit ; END ;
   if FListe.Cells[0,1]='' then
      if MsgBox.Execute(12,'','')<>mrYes then
         BEGIN BPages.ActivePage:=PComplement ; FListe.SetFocus ; Exit ; END ;
   if PresenceRub then if MsgBox.Execute(13,'','')<>mrYes then BEGIN CurseurAuCodeRub ; Exit ; END ;
   if Not VerifCoherenceTL then Exit ;
   END ;
DateModification(QBudsect,'BS') ;
EcritCompteExclu ; Result:=True ;
END ;

Function TFbudsect.PresenceRub : Boolean ;
BEGIN
if Length(QBudsectBS_RUB.AsString)>5 then QBudsectBS_RUB.AsString:=Copy(QBudsectBS_RUB.AsString,1,5) ;
Result:=PresenceComplexe('BUDSECT',['BS_BUDSECT','BS_RUB'],['<>','='],[BS_BUDSECT.Text,BS_RUB.Text],['S','S']) ;
END ;

Procedure TFbudsect.CurseurAuCode ;
BEGIN BPages.ActivePage:=PCaract ; BS_BUDSECT.SetFocus ; END ;

Procedure TFbudsect.CurseurAuCodeRub ;
BEGIN BPages.ActivePage:=PCaract ; BS_RUB.SetFocus ; END ;

Procedure TFbudsect.CurseurAuLibelle ;
BEGIN BPages.ActivePage:=PCaract ; BS_LIBELLE.SetFocus ; END ;

procedure TFbudsect.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFbudsect.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFbudsect.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFbudsect.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFbudsect.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFbudsect.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFbudsect.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then Bouge(nbInsert) ;
   if ((Mode=taCreatOne) or (Mode=taModifEnSerie)) then Close ;
   END ;
end;

procedure TFbudsect.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFbudsect.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFbudsect.BImprimerClick(Sender: TObject);
begin
{$IFNDEF IMP}
PlanBudSec(QBudsectBS_AXE.AsString, QBudsectBS_BUDSECT.AsString,True)
{$ENDIF}
end;

Procedure TFbudsect.NewEnreg ;
BEGIN
InitNew(QBudsect) ;
QBudsectBS_SENS.AsString:='M' ;
if Mode in [taCreat..taCreatOne] then
   BEGIN
   if AxeSerie<>'' then QBudsectBS_AXE.AsString:=AxeSerie
                   else QBudsectBS_AXE.AsString:='A1' ;
   END else
   BEGIN
   if Mode in [taModif] then
      BEGIN
      if QuelAxe<>'' then QBudsectBS_AXE.AsString:=QuelAxe
                     else QBudsectBS_AXE.AsString:='A1' ;
      END ;
   END ;
BS_HT.Checked:=False ; BS_BUDSECT.Enabled:=True ; BS_AXE.Enabled:=True ;
BPages.ActivePage:=PCaract ; BS_BUDSECT.SetFocus ; DateCreation(QBudsect,'BS') ;
If Mode=taCreatEnSerie Then
   BEGIN
   If SensDefaut<>'' Then QBudsectBS_SENS.AsString:=SensDefaut ;
   If SigneDefaut<>'' Then QBudsectBS_SIGNE.AsString:=SigneDefaut ;
   If AxeDefaut<>'' Then QBudsectBS_Axe.AsString:=AxeDefaut ;
   END ;
END ;

Function TFbudsect.Supprime : Boolean ;
BEGIN
Result:=False ;
if Msgbox.Execute(1,'','')<>mrYes then Exit ;
if EstMouvementeBudsect(BS_BUDSECT.Text,QBudsectBS_AXE.AsString)then
   BEGIN MsgBox.Execute(6,'','') ; Exit ; END ;
Result:=True ;
END ;

Function TFbudsect.CodeValide : Boolean ;
Var fb : TFichierBase ;
BEGIN
Result:=False ;
if QBudsect.state in [dsInsert] then
   BEGIN
   if QBudsectBS_BUDSECT.asString='' then
      BEGIN CurseurAuCode ; Msgbox.Execute(2,'','') ; Exit ; END ;
   fb:=TFichierBase(Ord(AxeToFb(QBudsectBS_AXE.AsString))-Ord(fbAxe1)+Ord(fbBudSec1)) ;
   if ExisteCpte(QBudsectBS_BUDSECT.AsString,fb) then
      BEGIN CurseurAuCode ; MsgBox.Execute(4,'','') ; Exit ; END ;
   END ;
Result:=True ;
END ;

Function TFbudsect.OkConfidentiel : Boolean ;
BEGIN
Result:=False ;
if (V_PGI.Confidentiel='0') and (BS_Confidentiel.Checked) then
  BEGIN
  MessageAlerte(MsgBox.Mess[7]) ;
  PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
  END ;
BS_CONFIDENTIEL.Visible:=(V_PGI.Confidentiel='1') ;
BS_CONFIDENTIEL.Enabled:=V_PGI.Superviseur ;
Result:=True ;
END ;

Function EstMouvementeBudsect(LeCpte,Laxe : String) : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select BS_BUDSECT From BUDSECT Where BS_BUDSECT="'+LeCpte+'" '+
              'And BS_AXE="'+Laxe+'" And '+
              '(Exists(Select BE_BUDSECT From BUDECR '+
              'Where BE_BUDSECT="'+LeCpte+'" And BE_AXE="'+Laxe+'"))',True) ;
Result:=Not QLoc.Eof ; Ferme(QLoc) ;
END ;

Procedure TFbudsect.InitModifEnSerie(StModif : string) ;
var St,Champ, Valeur : string;
    i             : integer;
    B             : TBitBtn;
BEGIN
if QBudsect.State=dsBrowse then QBudsect.Edit ;
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

Procedure TFbudsect.AffecteLe(Champ, Valeur : string) ;
var C : TControl;
BEGIN
C:=TControl(FindComponent(Champ)) ;
if(C is TDBCheckBox)Or(C is THDBValComboBox)Or(C is TDBEdit)Or(C is THDBCpteEdit)then
   BEGIN
   QBudsect.FindField(Champ).AsString:=Valeur ; TEdit(C).Font.Color:=clRed;
   END else if C is THDBSpinEdit then
   BEGIN
   QBudsect.FindField(Champ).AsInteger:=StrToInt(Valeur) ; THDBSpinEdit(C).Font.Color:=clRed;
   END ;
END;

Function TFbudsect.Grid2St ( ACol : Integer) : String ;
Var i : Integer ;
    St : String ;
BEGIN
St:='' ;
for i:=1 to FListe.RowCount-1 do
    St:=St+FListe.Cells[ACol,i]+';' ;
Result:=St ; if Length(Result)>250 then MsgBox.Execute(8,'','') ;
END ;

Procedure TFbudsect.LitCompteExclu(St : String ; ACol : Integer) ;
Var i : Integer ;
    St1 : String ;
BEGIN
i:=1 ;
While St <> '' do
  BEGIN
  St1:=ReadTokenSt(St) ;
  if St1 <> '' then
     BEGIN
     FListe.Cells[ACol,i]:=St1 ; Inc(i) ;
     END else
     BEGIN
     if ACol=1 then Inc(i) ;
     END ;
  if i > Fliste.RowCount-1 then FListe.RowCount:=FListe.RowCount+1 ;
  END ;
if FListe.Cells[0,FListe.RowCount-2]='' then FListe.RowCount:=FListe.RowCount-1 ;
END ;

Procedure TFbudsect.EcritCompteExclu ;
BEGIN
QBudsectBS_SECTIONRUB.AsString:=Grid2St(0) ;
QBudsectBS_EXCLURUB.AsString:=Grid2St(1) ;
END ;

Procedure TFbudsect.RempliGrid ;
BEGIN
FListe.VidePile(False) ;
LitCompteExclu(QBudsectBS_SECTIONRUB.AsString,0) ;
LitCompteExclu(QBudsectBS_EXCLURUB.AsString,1) ;
END ;

procedure TFbudsect.FListeKeyPress(Sender: TObject; var Key: Char);
begin
if Mode=taConsult then Exit ;
if Not (QBudsect.State in [dsEdit,dsInsert]) then
   BEGIN
   QBudsect.Edit ; QBudsectBS_LIBELLE.AsString:=QBudsectBS_LIBELLE.AsString ;
   END ;
Key:=UpCase(Key) ;
end;

procedure TFbudsect.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
//Var i : Byte ;
begin
if Mode=taConsult then Exit ;
if Not(ssCtrl in Shift) then Exit ;
if Key<>VK_DELETE then Exit ;
if FListe.RowCount<=2 then Exit ;
//for i:=0 to FListe.ColCount-1 do FListe.Cells[i,FListe.Row]:='' ;
FListe.DeleteRow(FListe.Row) ;
if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount:=FListe.RowCount+1 ;
if Not (QBudsect.Modified) then
   BEGIN
   if Not (QBudsect.State in [dsEdit,dsInsert]) then QBudsect.Edit ;
   QBudsectBS_LIBELLE.AsString:=QBudsectBS_LIBELLE.AsString ;
   END ;
end;

procedure TFbudsect.FListeCellExit(Sender: TObject; var ACol, ARow: Longint;
  var Cancel: Boolean);
begin
if Mode = taConsult then Exit ;
if QBudsect.Modified then
   if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount:=FListe.RowCount+1 ;
end;

procedure TFbudsect.BS_AXEChange(Sender: TObject);
begin
if QBudsect.State<>dsEdit then Exit ;
if MemoAxe<>QBudsectBS_AXE.AsString then
   if EstMouvementeBudsect(QBudsectBS_BUDSECT.AsString,QBudsectBS_AXE.AsString) then
      BEGIN MsgBox.Execute(6,'','') ; QBudsectBS_AXE.AsString:=MemoAxe ; END ;
end;

procedure TFbudsect.FormKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
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

Function SectionRetrieBud(CodRupt : String ; Axe : String ; ListeCodes : TStringList) : TSectRetri ;
Var Q : TQuery ;
    Deb,Lon : TabByte ;
    St,StTemp,Ru,Sql : String ;
    i,Lg : Byte   ;
    Trie,Execute : Boolean ;
BEGIN
Result:=srOk ;
if Not VH^.Cpta[AxeToFb(Axe)].Structure then BEGIN Result:=srNonStruct ; Exit ; END ;
Ru:='RU'+Axe[2] ;
Q:=TQuery.Create(Application) ; Q.DataBaseName:='SOC' ; Q.Close ; Q.Sql.Clear ;
Sql:='Select CC_LIBRE from CHOIXCOD Where CC_TYPE="'+Ru+'" and CC_CODE="'+CodRupt+'"' ;
Q.Sql.Add(Sql) ; ChangeSql(Q) ; Q.Open ;
if(Q.Fields[0].AsString='') Or (IsFieldNull(Q,'CC_LIBRE')) Or (Q.Eof) then
   BEGIN Result:=srPasEnchainement ; Exit ; END else StTemp:=Q.Fields[0].AsString ;
St:='' ;
Sql:='Select SS_DEBUT,SS_LONGUEUR from STRUCRSE Where SS_AXE="'+Axe+'" and SS_SOUSSECTION=:COD' ;
Q.Close ; Q.Sql.Clear ; Q.Sql.Add(Sql) ; ChangeSql(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
FillChar(Deb,SizeOF(Deb),#0) ; FillChar(Lon,SizeOF(Lon),#0) ; i:=1 ;
While StTemp<>'' do
      BEGIN
      St:=ReadTokenSt(StTemp) ;
      Q.Close ; Q.ParamByName('COD').AsString:=St ; Q.Open ;
      Deb[i]:=Q.Fields[0].AsInteger ; Lon[i]:=Q.Fields[1].AsInteger ;
      Inc(i) ;
      END ;
Lg:=VH^.Cpta[AxeToFb(Axe)].Lg ;
Sql:='Select BS_BUDSECT,BS_SECTIONTRIE from BUDSECT Where BS_AXE="'+Axe+'"' ;
Q.Close ; Q.Sql.Clear ; Q.Sql.Add(Sql) ; ChangeSql(Q) ; Q.Open ;
Trie:=False ; Execute:=False ;
While(Not Q.EOF)AND(Not Trie)do
     BEGIN
     St:='' ;
     for i:=1 to Lg do
        BEGIN
        StTemp:=Copy(Q.Fields[0].AsString,Deb[i],Lon[i]) ;
        if StTemp='' then Break ;
        St:=St+StTemp ;
//        if ListeCodes.IndexOf(St+'x')<0 then St:='' ;
        END ;
        if Not Execute then
           BEGIN
           if (Q.Fields[1].AsString='') or IsFieldNull(Q,'BS_SECTIONTRIE') then
              BEGIN Trie:=False ; Execute:=True ; END  else
              if Q.Fields[0].AsString<>VH^.Cpta[AxeToFb(Axe)].Attente then
                 BEGIN
                 Execute:=True ;
                 if Q.Fields[1].AsString<>St then Trie:=False else Trie:=True ;
                 END ;
           END ;
     if Not Trie then
        BEGIN
        ExecuteSql('UPDATE BUDSECT SET BS_SECTIONTRIE="'+St+'" Where BS_BUDSECT="'+Q.Fields[0].AsString+'" And BS_AXE="'+Axe+'"') ;
        END ;
     Q.Next ;
     END ;
Ferme(Q) ;
END ;

procedure TFbudsect.HMTradBeforeTraduc(Sender: TObject);
begin
LibellesTableLibre(ZL,'TBS_TABLE','BS_TABLE','D') ;
end;

Function TFbudsect.VerifCoherenceTL : Boolean ;
Var i : Integer ;
    C : TComponent ;
    Alerte : Boolean ;
BEGIN
Result:=True ;
if Not ZL.TabVisible then BEGIN Result:=True ; Exit ; END ;
Alerte:=False ;
for i:=0 to 9 do
  BEGIN
  C:=FindComponent('BS_TABLE'+IntToStr(i)) ;
  if Not(THDBCpteEdit(C).Enabled) then Continue ;
  if THDBCpteEdit(C).Text='' then Continue ;
  if THDBCpteEdit(C).ExisteH<=0 then Alerte:=True ;
  END ;
if Alerte then if MsgBox.Execute(14,'','')<>mrYes then Result:=False ;
END ;

procedure TFbudsect.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
