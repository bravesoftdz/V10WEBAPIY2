{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par BUDGENE_TOM.PAS
Mots clefs ... : 
*****************************************************************}
unit Budgene;

interface

uses
    WinTypes, WinProcs, Classes, Graphics, Forms, Controls, StdCtrls, Tabs,
    Buttons, ExtCtrls, Grids, Mask,  DB, DBTables, DBCtrls,HEnt1,
    Hctrls, HDB, DBGrids, Hqry, TabNotBk, Dialogs, Spin, SysUtils, Ent1,
    ComCtrls, hmsgbox,hcompte, Messages, Menus, HRichEdt, HRichOLE,
    HSysMenu,MajTable, HRegCpte, ParamSoc,UtilPGI ;

Procedure FicheBudgene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
Procedure FicheBudgeneMZS(Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);

type
  TFbudgene = class(TForm)
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
    BG_FERME: TDBCheckBox;
    TBG_BLOCNOTE: TGroupBox;
    BG_BLOCNOTE: THDBRichEditOLE;
    SBudgene: TDataSource;
    TBG_DATECREATION: THLabel;
    BG_DATECREATION: TDBEdit;
    TBG_DATEMODIF: THLabel;
    BG_DATEMODIF: TDBEdit;
    TBG_DATEOUVERTURE: THLabel;
    BG_DATEOUVERTURE: TDBEdit;
    TBG_DATEFERMETURE: THLabel;
    BG_DATEFERMETURE: TDBEdit;
    DBNav              : TDBNavigator;
    MsgBox             : THMsgBox;
    TB_BUDGET          : THLabel;
    BG_BUDGENE: TDBEdit;
    TB_LIBELLE         : THLabel;
    BG_LIBELLE: TDBEdit;
    TB_SENS            : THLabel;
    BG_SENS: THDBValComboBox;
    TB_ABREGE          : THLabel;
    BG_ABREGE: TDBEdit;
    QBudgene: TQuery;
    HMTrad: THSystemMenu;
    ZL: TTabSheet;
    TB_TABLE0: THLabel;
    TB_TABLE1: THLabel;
    TB_TABLE2: THLabel;
    TB_TABLE3: THLabel;
    TB_TABLE4: THLabel;
    TB_TABLE5: THLabel;
    TB_TABLE6: THLabel;
    TB_TABLE7: THLabel;
    TB_TABLE8: THLabel;
    TB_TABLE9: THLabel;
    QBudgeneBG_BUDGENE: TStringField;
    QBudgeneBG_LIBELLE: TStringField;
    QBudgeneBG_ABREGE: TStringField;
    QBudgeneBG_SIGNE: TStringField;
    QBudgeneBG_REPORTDISPO: TStringField;
    QBudgeneBG_PARENT: TStringField;
    QBudgeneBG_ATTENTE: TStringField;
    QBudgeneBG_FORMULE: TStringField;
    QBudgeneBG_COMPTERUB: TStringField;
    QBudgeneBG_EXCLURUB: TStringField;
    QBudgeneBG_DATECREATION: TDateTimeField;
    QBudgeneBG_DATEMODIF: TDateTimeField;
    QBudgeneBG_DATEOUVERTURE: TDateTimeField;
    QBudgeneBG_DATEFERMETURE: TDateTimeField;
    QBudgeneBG_FERME: TStringField;
    QBudgeneBG_SENS: TStringField;
    QBudgeneBG_BLOCNOTE: TMemoField;
    QBudgeneBG_CONFIDENTIEL: TStringField;
    QBudgeneBG_HT: TStringField;
    QBudgeneBG_UTILISATEUR: TStringField;
    QBudgeneBG_SOCIETE: TStringField;
    QBudgeneBG_BLOQUANT: TStringField;
    QBudgeneBG_CREERPAR: TStringField;
    QBudgeneBG_EXPORTE: TStringField;
    QBudgeneBG_TABLE0: TStringField;
    QBudgeneBG_TABLE1: TStringField;
    QBudgeneBG_TABLE2: TStringField;
    QBudgeneBG_TABLE3: TStringField;
    QBudgeneBG_TABLE4: TStringField;
    QBudgeneBG_TABLE5: TStringField;
    QBudgeneBG_TABLE6: TStringField;
    QBudgeneBG_TABLE7: TStringField;
    QBudgeneBG_TABLE8: TStringField;
    QBudgeneBG_TABLE9: TStringField;
    BG_SIGNE: THDBValComboBox;
    TBG_SIGNE: TLabel;
    BG_REPORTDISPO: THDBValComboBox;
    BG_PARENT: TDBEdit;
    TBG_PARENT: TLabel;
    BG_ATTENTE: TDBCheckBox;
    BG_HT: TDBCheckBox;
    TBG_REPORTDISPO: TLabel;
    BG_FORMULE: TDBEdit;
    TBG_FORMULE: TLabel;
    BG_CONFIDENTIEL: TDBCheckBox;
    FListe: THGrid;
    QBudgeneBG_RUB: TStringField;
    TBG_RUB: THLabel;
    BG_RUB: TDBEdit;
    BG_TABLE0: THDBCpteEdit;
    BG_TABLE1: THDBCpteEdit;
    BG_TABLE2: THDBCpteEdit;
    BG_TABLE3: THDBCpteEdit;
    BG_TABLE4: THDBCpteEdit;
    BG_TABLE5: THDBCpteEdit;
    BG_TABLE6: THDBCpteEdit;
    BG_TABLE7: THDBCpteEdit;
    BG_TABLE8: THDBCpteEdit;
    BG_TABLE9: THDBCpteEdit;
    procedure FormCreate(Sender: TObject);
    procedure SBudgeneDataChange(Sender: TObject; Field: TField);
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
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BG_BUDGENEExit(Sender: TObject);
  private { Private declarations }
    Q : TQuery ;
    Mode : TActionFiche ;
    LaPage : Integer ;
    Lequel,LeCompte : String ;
    LesModif : String ;
    LgMin,LgMax : Integer ;
    SensDefaut,SigneDefaut : String ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Procedure CurseurAuCode ;
    Procedure CurseurAuLibelle ;
    Procedure CurseurAuCodeRub ;
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
    Procedure ChargeLgMinMax ;
    Function  PresenceRub : Boolean ;
    Function  VerifCoherenceTL : Boolean ;
  public  { Public declarations }
  end;

    Function  EstMouvementeBudgen(LeCpte : String)   : Boolean ;  //ECRITURE

implementation

{$R *.DFM}

{$IFNDEF IMP}
uses QRBudgen ;
{$ENDIF}

procedure TFbudgene.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; Q:=NIL ;
LeCompte:='aze' ;
end;

procedure TFbudgene.SBudgeneDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not(QBudgene.State in [dsEdit,dsInsert]) ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not QBudgene.BOF;
   DnEnable := Enabled and not QBudgene.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
// code pour gerer les champ +- automatique
   if ((Field.FieldName='BG_LIBELLE') and (BG_ABREGE.Field.AsString='')) then
      BG_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   if ((Field.FieldName='BG_BUDGENE') and (BG_RUB.Field.AsString='')) then
      BG_RUB.Field.AsString:=Copy(Field.AsString,1,5) ;
   END ;
end;

Procedure TFbudgene.ChargeLgMinMax ;
{$IFDEF SPEC302}
Var QLoc : TQuery ;
{$ENDIF}
BEGIN
{$IFDEF SPEC302}
QLoc:=OpenSql('Select SO_LGMAXBUDGET,SO_LGMINBUDGET From SOCIETE',True) ;
LgMax:=QLoc.Fields[0].AsInteger ; LgMin:=QLoc.Fields[1].AsInteger ;
Ferme(QLoc) ;
{$ELSE}
  LgMax := GetParamSocSecur('SO_LGMAXBUDGET', 0);
  LgMin := GetParamSocSecur('SO_LGMINBUDGET', 0);
{$ENDIF}
if LgMax=0 then LgMax:=17 ;
END ;

procedure TFbudgene.FormShow(Sender: TObject);
begin
MakeZoomOLE(Handle) ;
ChargeLgMinMax ; RecupWhereSQL(Q,QBudgene) ;
if ((Q=NIL) and (Lequel<>'')) then QBudgene.SQL.Add('Where BG_BUDGENE="'+Lequel+'"') ;
if ((Q=NIL) and (Lequel<>'')) then ChangeSQL(QBudgene) ;
ChangeSizeMemo(QBudgeneBG_BLOCNOTE) ;
QBudgene.Open ;
if(Lequel<>'')And((Mode in [taCreat..taCreatOne])=False) then
   BEGIN
   if Not QBudgene.Locate('BG_BUDGENE',Lequel,[]) then
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
     taCreat..taCreatOne : BEGIN Bouge(nbInsert) ; BG_BUDGENE.Text:=Lequel ; BAnnuler.Enabled:=False ; END ;
     taModif             : BEGIN END ;
     taModifEnSerie      : InitModifEnSerie(LesModif);
     end ;
SensDefaut:='' ; SigneDefaut:='' ;
if(QBudgene.Eof) then Bouge(nbInsert) ;
end;

Procedure TFbudgene.ChargeEnreg ;
BEGIN
if LeCompte=QBudgeneBG_BUDGENE.AsString then
   BEGIN
   if Mode=taConsult then
      BEGIN
      FicheReadOnly(Self) ;
      FListe.Enabled:=True ; FListe.Options:=FListe.Options-[goEditing]+[goRowSelect] ;
      END ;
   Exit ;
   END ;
LeCompte:=QBudgeneBG_BUDGENE.AsString ;

InitCaption(Self,BG_BUDGENE.text,BG_LIBELLE.text) ;
RempliGrid ;
if Mode=taConsult then
   BEGIN
   FicheReadOnly(Self) ;
   FListe.Enabled:=True ; FListe.Options:=FListe.Options-[goEditing]+[goRowSelect] ;
   Exit ;
   END ;
BG_BUDGENE.Enabled:=FALSE ;
END ;

Function TFbudgene.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : BEGIN
                      if Not EnregOK then Exit ;
                      If mode=taCreatEnSerie Then
                         BEGIN
                         SensDefaut:=QBudgeneBG_SENS.AsString ;
                         SigneDefaut:=QBudgeneBG_SIGNE.AsString ;
                         END ;
                      END ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[5]) ;
Result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFbudgene.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
Result:=FALSE  ;
if QBudgene.Modified then
   BEGIN
   if (Mode in [taCreat..taCreatOne]) And
      (QBudgeneBG_BUDGENE.asString='') then Rep:=mrNo
   else if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
     mrYes : if not Bouge(nbPost) then Exit ;
     mrNo :  if not Bouge(nbCancel) then Exit ;
     mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFbudgene.EnregOK : boolean ;
BEGIN
Result:=False ;
if QBudgene.state in [dsEdit,dsInsert]=False then Exit ;
if Not CodeValide then Exit ;
if QBudgene.state in [dsEdit,dsInsert] then
   BEGIN
   if QBudgeneBG_LIBELLE.asString='' then
      BEGIN Msgbox.Execute(3,'','') ; CurseurAuLibelle ; Exit ; END ;
   if QBudgeneBG_RUB.asString='' then
      BEGIN Msgbox.Execute(12,'','') ; CurseurAuCodeRub ; Exit ; END ;
   if FListe.Cells[0,1]='' then
      if MsgBox.Execute(13,'','')<>mrYes then
         BEGIN BPages.ActivePage:=PComplement ; FListe.SetFocus ; Exit ; END ;
   if PresenceRub then if MsgBox.Execute(14,'','')<>mrYes then BEGIN CurseurAuCodeRub ; Exit ; END ;
   if Not VerifCoherenceTL then Exit ;
   END ;
DateModification(QBudgene,'BG') ;
EcritCompteExclu ; Result:=True ;
END ;

Function TFbudgene.PresenceRub : Boolean ;
BEGIN
if Length(QBudgeneBG_RUB.AsString)>5 then QBudgeneBG_RUB.AsString:=Copy(QBudgeneBG_RUB.AsString,1,5) ;
Result:=PresenceComplexe('BUDGENE',['BG_BUDGENE','BG_RUB'],['<>','='],[BG_BUDGENE.Text,BG_RUB.Text],['S','S']) ;
END ;

Procedure TFbudgene.CurseurAuCode ;
BEGIN BPages.ActivePage:=PCaract ; BG_BUDGENE.SetFocus ; END ;

Procedure TFbudgene.CurseurAuLibelle ;
BEGIN BPages.ActivePage:=PCaract ; BG_LIBELLE.SetFocus ; END ;

Procedure TFbudgene.CurseurAuCodeRub ;
BEGIN BPages.ActivePage:=PCaract ; BG_RUB.SetFocus ; END ;

procedure TFbudgene.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFbudgene.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFbudgene.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFbudgene.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFbudgene.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFbudgene.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFbudgene.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then Bouge(nbInsert) ;
   if ((Mode=taCreatOne) or (Mode=taModifEnSerie)) then Close ;
   END ;
end;

procedure TFbudgene.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFbudgene.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFbudgene.BImprimerClick(Sender: TObject);
begin
{$IFNDEF IMP}
PlanBudget(QBudgeneBG_BUDGENE.AsString,True)
{$ENDIF}
end;

Procedure TFbudgene.NewEnreg ;
BEGIN
InitNew(QBudgene) ;
QBudgeneBG_SENS.AsString:='M' ;
BG_HT.Checked:=False ;
BG_BUDGENE.Enabled:=True ; BPages.ActivePage:=PCaract ;
BG_BUDGENE.SetFocus ; DateCreation(QBudgene,'BG') ;
If Mode=taCreatEnSerie Then
   BEGIN
   If SensDefaut<>'' Then QBudgeneBG_SENS.AsString:=SensDefaut ;
   If SigneDefaut<>'' Then QBudgeneBG_SIGNE.AsString:=SigneDefaut ;
   END ;
END ;

Procedure FicheBudgene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
var FBudgene: TFBudgene;
begin
(*Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not JaiLeDroit(ccGenCreat,True) then Exit ;
              taModif,taModifEnSerie : if Not JaiLeDroit(ccGenModif,True) then Exit ;
   END ;*)
if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
FBudgene:=TFBudgene.Create(Application) ;
try
  FBudgene.Q:=Q ;
  FBudgene.Lequel:=Compte ;
  FBudgene.Mode:=Comment ;
  FBudgene.LaPage:=QuellePage ;
  FBudgene.ShowModal ;
  finally
  FBudgene.Free ;
  end ;
Screen.Cursor:=crDefault ;
end ;

Procedure FicheBudgeneMZS(Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
var FBudgene: TFBudgene;
begin
(*Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not JaiLeDroit(ccGenCreat,True) then Exit ;
              taModif,taModifEnSerie : if Not JaiLeDroit(ccGenModif,True) then Exit ;
   END ;*)
if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
FBudgene:=TFBudgene.Create(Application) ;
try
  FBudgene.Lequel:=Lequel ;
  FBudgene.Mode:=Comment ;
  FBudgene.LaPage:=QuellePage ;
  FBudgene.LesModif:=LesModif ;
  FBudgene.ShowModal ;
  finally
  FBudgene.Free ;
  end ;
Screen.Cursor:=crDefault ;
end ;

Function TFbudgene.Supprime : Boolean ;
BEGIN
Result:=False ;
if Msgbox.Execute(1,'','')<>mrYes then Exit ;
if EstMouvementeBudgen(BG_BUDGENE.Text)then BEGIN MsgBox.Execute(6,'','') ; Exit ; END ;
Result:=True ;
END ;

Function TFbudgene.CodeValide : Boolean ;
BEGIN
Result:=False ;
if QBudgene.state in [dsInsert] then
   BEGIN
   if QBudgeneBG_BUDGENE.asString='' then
      BEGIN CurseurAuCode ; Msgbox.Execute(2,'','') ; Exit ; END ;
   if Length(QBudgeneBG_BUDGENE.asString)<LgMin then
      BEGIN CurseurAuCode ; Msgbox.Execute(10,'','') ; Exit ; END ;
   if Length(QBudgeneBG_BUDGENE.asString)>LgMax then
      BEGIN CurseurAuCode ; Msgbox.Execute(11,'','') ; Exit ; END ;
   if Presence('BUDGENE','BG_BUDGENE',BG_BUDGENE.Text)  then
      BEGIN CurseurAuCode ; MsgBox.Execute(4,'','') ; Exit ; END ;
   END ;
Result:=True ;
END ;

Function TFbudgene.OkConfidentiel : Boolean ;
BEGIN
Result:=False ;
if (V_PGI.Confidentiel='0') and (BG_Confidentiel.Checked) then
  BEGIN
  MessageAlerte(MsgBox.Mess[8]) ;
  PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
  END ;
BG_CONFIDENTIEL.Visible:=(V_PGI.Confidentiel='1') ;
BG_CONFIDENTIEL.Enabled:=V_PGI.Superviseur ;
Result:=True ;
END ;

Function EstMouvementeBudgen(LeCpte : String) : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select BG_BUDGENE From BUDGENE Where BG_BUDGENE="'+LeCpte+'" '+
              'And (Exists(Select BE_BUDGENE From BUDECR Where BE_BUDGENE="'+LeCpte+'"))',True) ;
Result:=Not QLoc.Eof ; Ferme(QLoc) ;
END ;

Procedure TFbudgene.InitModifEnSerie(StModif : string) ;
var St,Champ, Valeur : string;
    i             : integer;
    B             : TBitBtn;
BEGIN
if QBudgene.State=dsBrowse then QBudgene.Edit ;
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

Procedure TFbudgene.AffecteLe(Champ, Valeur : string) ;
var C : TControl;
BEGIN
C:=TControl(FindComponent(Champ)) ;
if(C is TDBCheckBox)Or(C is THDBValComboBox)Or(C is TDBEdit)Or(C is THDBCpteEdit)then
   BEGIN
   QBudgene.FindField(Champ).AsString:=Valeur ; TEdit(C).Font.Color:=clRed;
   END else if C is THDBSpinEdit then
   BEGIN
   QBudgene.FindField(Champ).AsInteger:=StrToInt(Valeur) ; THDBSpinEdit(C).Font.Color:=clRed;
   END ;
END;

Function TFbudgene.Grid2St ( ACol : Integer) : String ;
Var i : Integer ;
    St : String ;
BEGIN
St:='' ;
for i:=1 to FListe.RowCount-1 do
    St:=St+FListe.Cells[ACol,i]+';' ;
Result:=St ; if Length(Result)>250 then MsgBox.Execute(9,'','') ;
END ;

Procedure TFbudgene.LitCompteExclu(St : String ; ACol : Integer) ;
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

Procedure TFbudgene.EcritCompteExclu ;
BEGIN
QBudgeneBG_COMPTERUB.AsString:=Grid2St(0) ;
QBudgeneBG_EXCLURUB.AsString:=Grid2St(1) ;
END ;

Procedure TFbudgene.RempliGrid ;
BEGIN
FListe.VidePile(False) ;
LitCompteExclu(QBudgeneBG_COMPTERUB.AsString,0) ;
LitCompteExclu(QBudgeneBG_EXCLURUB.AsString,1) ;
END ;

procedure TFbudgene.FListeKeyPress(Sender: TObject; var Key: Char);
begin
if Mode=taConsult then Exit ;
if Not (QBudgene.State in [dsEdit,dsInsert]) then
   BEGIN
   QBudgene.Edit ; QBudgeneBG_LIBELLE.AsString:=QBudgeneBG_LIBELLE.AsString ;
   END ;
Key:=UpCase(Key) ;
end;

procedure TFbudgene.FListeKeyDown(Sender: TObject; var Key: Word;
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
if Not (QBudgene.Modified) then
   BEGIN
   if Not (QBudgene.State in [dsEdit,dsInsert]) then QBudgene.Edit ;
   QBudgeneBG_LIBELLE.AsString:=QBudgeneBG_LIBELLE.AsString ;
   END ;
end;

procedure TFbudgene.FListeCellExit(Sender: TObject; var ACol, ARow: Longint;
  var Cancel: Boolean);
begin
if Mode = taConsult then Exit ;
if QBudgene.Modified then
   if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount:=FListe.RowCount+1 ;
end;

procedure TFbudgene.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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

procedure TFbudgene.HMTradBeforeTraduc(Sender: TObject);
begin
LibellesTableLibre(ZL,'TB_TABLE','BG_TABLE','B') ;
end;

Function TFbudgene.VerifCoherenceTL : Boolean ;
Var i : Integer ;
    C : TComponent ;
    Alerte : Boolean ;
BEGIN
Result:=True ;
if Not ZL.TabVisible then BEGIN Result:=True ; Exit ; END ;
Alerte:=False ;
for i:=0 to 9 do
  BEGIN
  C:=FindComponent('BG_TABLE'+IntToStr(i)) ;
  if Not(THDBCpteEdit(C).Enabled) then Continue ;
  if THDBCpteEdit(C).Text='' then Continue ;
  if THDBCpteEdit(C).ExisteH<=0 then Alerte:=True ;
  END ;
if Alerte then if MsgBox.Execute(15,'','')<>mrYes then Result:=False ;
END ;

procedure TFbudgene.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;

end;

procedure TFbudgene.BG_BUDGENEExit(Sender: TObject);
begin
if QBudGene.State=dsInsert then QBudgeneBG_BUDGENE.AsString:=Trim(QBudgeneBG_BUDGENE.AsString) ;
end;

end.
