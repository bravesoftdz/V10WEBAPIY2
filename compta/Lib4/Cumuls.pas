unit Cumuls;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Mask,
  HCtrls,
  Ent1,
  hmsgbox,
  Grids,
  Hqry,
{$IFDEF EAGLCLIENT}

{$ELSE}
  DBCtrls,
  DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  PrintDbg,
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
  HSysMenu, ADODB;

Procedure ParametrageCumuls(Constante : Boolean ; LeQuel,Lib : String) ;

type
  TFCumuls = class(TForm)
    DBnav      : TDBNavigator;
    QCumuls    : THQuery;
    QSuite     : THQuery;
    HPB        : TPanel;
    MsgBox     : THMsgBox;
    FAutoSave  : TCheckBox;
    Pappli     : TPanel;
    CU_SUITE   : TDBCheckBox;
    QCumulsCU_TYPE          : TStringField;
    QCumulsCU_COMPTE1       : TStringField;
    QCumulsCU_COMPTE2       : TStringField;
    QCumulsCU_EXERCICE      : TStringField;
    QCumulsCU_SUITE         : TStringField;
    QCumulsCU_ETABLISSEMENT : TStringField;
    QCumulsCU_DEVQTE        : TStringField;
    QCumulsCU_QUALIFPIECE   : TStringField;
    QCumulsCU_DEBIT1        : TFloatField;
    QCumulsCU_DEBIT2        : TFloatField;
    QCumulsCU_DEBIT3        : TFloatField;
    QCumulsCU_DEBIT4        : TFloatField;
    QCumulsCU_DEBIT5        : TFloatField;
    QCumulsCU_DEBIT6        : TFloatField;
    QCumulsCU_DEBIT7        : TFloatField;
    QCumulsCU_DEBIT8        : TFloatField;
    QCumulsCU_DEBIT9        : TFloatField;
    QCumulsCU_DEBIT10       : TFloatField;
    QCumulsCU_DEBIT11       : TFloatField;
    QCumulsCU_DEBIT12       : TFloatField;
    QCumulsCU_CREDIT1       : TFloatField;
    QCumulsCU_CREDIT2       : TFloatField;
    QCumulsCU_CREDIT3       : TFloatField;
    QCumulsCU_CREDIT4       : TFloatField;
    QCumulsCU_CREDIT5       : TFloatField;
    QCumulsCU_CREDIT6       : TFloatField;
    QCumulsCU_CREDIT7       : TFloatField;
    QCumulsCU_CREDIT8       : TFloatField;
    QCumulsCU_CREDIT9       : TFloatField;
    QCumulsCU_CREDIT10      : TFloatField;
    QCumulsCU_CREDIT11      : TFloatField;
    QCumulsCU_CREDIT12      : TFloatField;
    QCumulsCU_CREDITAN      : TFloatField;
    QCumulsCU_DEBITAN       : TFloatField;
    QCumulsCU_QUALIFQTEDIFF : TStringField;
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    BAnnuler: THBitBtn;
    BImprimer: THBitBtn;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    XX_WHERE: TPanel;
    Fliste: THGrid;
    GbCar: TGroupBox;
    TCU_ETABLISSEMENT: THLabel;
    TCU_EXERCICE: THLabel;
    CU_EXERCICE: THValComboBox;
    CU_ETABLISSEMENT: THValComboBox;
    CU_COMPTE1: TEdit;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FlisteKeyPress(Sender: TObject; var Key: Char);
    procedure FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure CU_ETABLISSEMENTChange(Sender: TObject);
    procedure QCumulsBeforePost(DataSet: TDataSet);
    procedure CU_EXERCICEChange(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    LeQuel,Lib : String ;
    Exo : TExoDate ;
    PremMois,PremAnnee,NombreMois : Word ;
    Constante : Boolean ;
    EcrirSuite : Boolean ;
    Insertion : Boolean ;
    Cpte1,Lexo,Letab : String ;
    WMinX,WMinY : Integer ;
    AFermer : Boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure ChargeEnreg ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure RempliGridPeriode ;
    Procedure EffaceLeGrid ;
    Procedure EcritDebCred ;
    Procedure LitDebCred ;
    Procedure RemplirLaSuite ;
    Procedure LireLaSuite ;
    Function  ChercheSuite : Boolean ;
    Procedure InitGrid ;
    Procedure PositionneQCumuls ;
    Procedure ReloadQCumuls ;
    Function  PresenceCle : Boolean ;
    Function  ChercheLequel : Boolean ;
    Function  ChercheEnreg : Boolean ;
    Procedure PositionneExo ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  HEnt1;

Procedure ParametrageCumuls(Constante : Boolean ; LeQuel,Lib : String);
var FCumuls : TFCumuls ;
BEGIN
FCumuls:=TFCumuls.Create(Application) ;
 Try
  FCumuls.Constante:=Constante ;
  FCumuls.LeQuel:=LeQuel ;
  FCumuls.Lib:=Lib ;
  FCumuls.ShowModal ;
 Finally
  FCumuls.Free ;
 End ;
END ;

procedure TFCumuls.BFermeClick(Sender: TObject);
begin Close ; end;

Function TFCumuls.ChercheLequel : Boolean ;
BEGIN
Result:=False ;
if Lequel='' then Exit ;
While Not QCumuls.Eof do
   BEGIN
   if(QCumulsCU_TYPE.AsString='CON')And(QCumulsCU_COMPTE1.AsString=Lequel)And
     (QCumulsCU_COMPTE2.AsString='')And(QCumulsCU_EXERCICE.AsString=VH^.Entree.Code)
      then BEGIN Result:=True ; Exit ; END else QCumuls.Next ;
   END ;
END ;

procedure TFCumuls.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
PremMois:=0 ; PremAnnee:=0 ; NombreMois:=0 ; AFermer:=False ;
Cpte1:='' ; Lexo:='' ;
InitCaption(Self,Lequel,Lib) ;
ChangeSQL(QCumuls) ; QCumuls.Open ;
if((QCumuls.EOF)And(QCumuls.BOF))Or((Not ChercheLeQuel)And(Lequel<>''))then
  BEGIN
  CU_COMPTE1.Text:=Lequel ; CU_ETABLISSEMENT.ItemIndex:=0 ;
  CU_EXERCICE.Value:=VH^.Entree.Code ;
  END else ChargeEnreg ;
if AFermer then BEGIN PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;  END ;
end;

Procedure TFCumuls.ChargeEnreg ;
BEGIN
if (QCumuls.Eof) and (QCumuls.Bof) then Exit ;
CU_COMPTE1.Text:=QCumulsCU_COMPTE1.AsString ;
if CU_COMPTE1.Text='' then CU_COMPTE1.Text:=Lequel ;
CU_ETABLISSEMENT.ItemIndex:=CU_ETABLISSEMENT.Values.IndexOf(QCumulsCU_ETABLISSEMENT.AsString) ;
if CU_ETABLISSEMENT.ItemIndex=-1 then CU_ETABLISSEMENT.ItemIndex:=0 ;
CU_EXERCICE.ItemIndex:=CU_EXERCICE.Values.IndexOf(QCumulsCU_EXERCICE.AsString) ;
if CU_EXERCICE.ItemIndex=-1 then CU_EXERCICE.ItemIndex:=CU_EXERCICE.Values.IndexOf(VH^.Entree.Code) ;
ChercheSuite ;
if(QCumuls.State=dsBrowse)And(Not EcrirSuite) then
  BEGIN CU_EXERCICEChange(Nil) ; LitDebCred ; END ;
InitCaption(Self,Lequel,Lib) ;
END ;

procedure TFCumuls.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; ChargeEnreg ; end;

procedure TFCumuls.BValiderClick(Sender: TObject);
begin if Bouge(nbPost) then ChargeEnreg ; end;

procedure TFCumuls.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

Procedure TFCumuls.NewEnreg ;
BEGIN
InitNew(QCumuls) ;
QCumulsCU_TYPE.AsString:='CON' ; QCumulsCU_EXERCICE.AsString:=CU_EXERCICE.Value ;
QCumulsCU_ETABLISSEMENT.AsString:=CU_ETABLISSEMENT.Value ;
QCumulsCU_COMPTE1.AsString:=CU_COMPTE1.Text ; InitGrid ;
END ;

Procedure TFCumuls.InitGrid ;
Var i : Byte ;
BEGIN
for i:=1 to Fliste.RowCount-1 do
   Fliste.Cells[1,i]:=IntToStr(0) ;
END ;

Function TFCumuls.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if QCumuls.Modified then
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

Function TFCumuls.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if QCumuls.state in [dsEdit,dsInsert]=False then Exit ;
if QCumuls.state in [dsInsert] then
   BEGIN
   if CU_ETABLISSEMENT.Value='' then
      BEGIN MsgBox.Execute(12,'','') ; CU_ETABLISSEMENT.SetFocus ; Exit ; END ;
   if NombreMois>12 then QCumulsCU_SUITE.AsString:='X'
                    else QCumulsCU_SUITE.AsString:='-' ;
//   if PresenceCle then BEGIN MsgBox.Execute(4,'','') ; CU_ETABLISSEMENT.SetFocus ; Exit ; END ;
   END ;
Result:=TRUE  ;
END ;

Function TFCumuls.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
     nblast,nbprior,nbnext,nbfirst,
     nbinsert : if Not OnSauve then Exit ;
     nbPost   : if Not EnregOK then Exit ;
     end ;
if(QCumuls.State in [dsEdit,dsInsert])And(Button=nbPost)then EcritDebCred ;
if QCumuls.State=dsInsert then Insertion:=True else Insertion:=False ;
if Button=nbPost then QCumuls.Post else
   if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[7]) ;
result:=TRUE ;
if EcrirSuite then RemplirLaSuite ;
if Insertion then ReloadQCumuls ;
if Button=NbInsert then NewEnreg ;
END ;

Procedure TFCumuls.RempliGridPeriode ;
Var i : Integer ;
    DebDate : TDateTime ;
    Md : String ;
BEGIN
EffaceLeGrid ;
DebDate:=Exo.Deb ;
for i:=1 to NombreMois+1 do
    BEGIN
    if Constante then if i=NombreMois+1 then Break ;
    Md:=FormatDateTime('mmmm yyyy',DebDate) ; Md[1]:=UpCase(Md[1]) ;
    FListe.Cells[0,i]:=Md ;
    DebDate:=PlusMois(DebDate,1) ;
    Fliste.RowCount:=Fliste.RowCount+1 ;
    END ;
Fliste.RowCount:=Fliste.RowCount-1 ;
END ;

Procedure TFCumuls.EffaceLeGrid ;
Var i,j : Byte ;
BEGIN
for i:=0 to Fliste.ColCount-1 do
    for j:=1 to Fliste.RowCount-1 do
        if i=0 then Fliste.Cells[i,j]:=''
               else Fliste.Cells[i,j]:=IntToStr(0) ;
Fliste.RowCount:=2 ;
END ;

procedure TFCumuls.FlisteKeyPress(Sender: TObject; var Key: Char);
begin
if Key='.' then BEGIN Key:=',' ; Exit End ;
if Not(Key in ['0'..'9',',',#8,#9,#13])then Key:=#0 ;
end;

procedure TFCumuls.FlisteKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if Key in [VK_PRIOR,VK_NEXT,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT] then Exit ;
if Not QCumuls.Modified then
   BEGIN
   if QCumuls.State=dsBrowse then QCumuls.Edit ;
   QCumulsCU_SUITE.AsString:=QCumulsCU_SUITE.AsString ;
   END ;
end;

Function TFCumuls.ChercheSuite : Boolean ;
Var St : String ;
BEGIN
St:='Select Count(CU_SUITE) From CUMULS Where CU_TYPE="CON" And '+
    'CU_COMPTE1="'+QCumulsCU_COMPTE1.AsString+'" And CU_COMPTE2="" And '+
    'CU_EXERCICE="'+QCumulsCU_EXERCICE.AsString+'" And CU_ETABLISSEMENT="'+CU_ETABLISSEMENT.Value+'" ' ;
QSuite.Close ; QSuite.Sql.Clear ; QSuite.Sql.Add(St) ; ChangeSql(QSuite) ; QSuite.Open ;
if ((Not QSuite.Eof) And (QSuite.Fields[0].AsInteger>1)) then
   BEGIN
   Result:=True ; PositionneQCumuls ;
   END else Result:=False ;
QSuite.Close ;
END ;

Procedure TFCumuls.PositionneQCumuls ;
Var LeCpte1,Lexocice,Etab : String ;
    St : String ;
BEGIN
LeCpte1:=QCumulsCU_COMPTE1.AsString ;
Lexocice:=QCumulsCU_EXERCICE.AsString ;
Etab:=QCumulsCU_ETABLISSEMENT.AsString ;
St:='Select * From CUMULS Where CU_SUITE="X" And CU_TYPE="CON" And CU_COMPTE1="'+LeCpte1+'" And '+
    'CU_COMPTE2="" And CU_EXERCICE="'+Lexocice+'" And CU_ETABLISSEMENT="'+Etab+'" ' ;
QCumuls.Close ; QCumuls.Sql.Clear ; QCumuls.Sql.Add(St) ;
ChangeSql(QCumuls) ; QCumuls.Open ;
END ;

Procedure TFCumuls.LitDebCred ;
Var i : Byte ;
    St : String ;
    LaSuite : Boolean ;
BEGIN
LaSuite:=False ;
St:='Select * From CUMULS Where CU_TYPE="CON" And '+
    'CU_COMPTE1="'+QCumulsCU_COMPTE1.AsString+'" And CU_COMPTE2="" And '+
    'CU_EXERCICE="'+QCumulsCU_EXERCICE.AsString+'" And CU_ETABLISSEMENT="'+QCumulsCU_ETABLISSEMENT.AsString+'" ';
St:=St+'Order by CU_SUITE DESC' ;
QSuite.Close ; QSuite.Sql.Clear ; QSuite.Sql.Add(St) ; ChangeSql(QSuite) ; QSuite.Open ;
QSuite.First ;
for i:=1 to Fliste.RowCount-1 do
    BEGIN
    if i=13 then BEGIN LaSuite:=True ; Break ; END ;
    Fliste.Cells[1,i]:=QSuite.FindField('CU_DEBIT'+IntToStr(i)).AsString ;
    END ;
if LaSuite then LireLaSuite ;
END ;

Procedure TFCumuls.LireLaSuite ;
Var i,j : Byte ;
    Q : TQuery ;
    St : String ;
BEGIN
i:=13 ;
St:='Select * from CUMULS Where CU_TYPE="CON" '+
    'And CU_EXERCICE="'+QCumulsCU_EXERCICE.AsString+'" And CU_COMPTE1="'+QCumulsCU_COMPTE1.AsString+'" '+
    'And CU_COMPTE2="" And CU_SUITE="-" ';
Q:=OpenSql(St,True) ;
if Not Q.Eof then
   BEGIN
   for j:=1 to 12 do
       BEGIN
       Fliste.Cells[1,i]:=Q.FindField('CU_DEBIT'+IntToStr(j)).AsString ;
       Inc(i) ;
       END ;
   END ;
Ferme(Q) ;
END ;

Procedure TFCumuls.EcritDebCred ;
Var i : Byte ;
BEGIN
EcrirSuite:=False ;
for i:=1 to Fliste.RowCount-1 do
    BEGIN
    if i>12 then BEGIN EcrirSuite:=True ; Break ; END ;
    if Fliste.Cells[1,i]=''then QCumuls.FindField('CU_DEBIT'+IntToStr(i)).AsFloat:=0
                           else QCumuls.FindField('CU_DEBIT'+IntToStr(i)).AsFloat:=StrToFloat(Fliste.Cells[1,i]) ;
    END ;
END ;

Procedure TFCumuls.RemplirLaSuite ;
Type TabDouble=Array[1..12] of Double ;
Var i,j : Byte ;
    TabDeb,TabCred : TabDouble ;
BEGIN
j:=13 ;
FillChar(TabDeb,Sizeof(Tabdeb),0) ; FillChar(TabCred,Sizeof(TabCred),0) ;
for i:=j to Fliste.RowCount-1 do
    TabDeb[(i-j)+1]:=StrToFloat(Fliste.Cells[1,i]) ;
if Insertion then
   ExecuteSql('INSERT INTO CUMULS (CU_TYPE,CU_COMPTE1,CU_COMPTE2,CU_EXERCICE,CU_SUITE,CU_ETABLISSEMENT,'+
              'CU_DEVQTE,CU_QUALIFPIECE,CU_DEBIT1,CU_DEBIT2,CU_DEBIT3,CU_DEBIT4,CU_DEBIT5,CU_DEBIT6,'+
              'CU_DEBIT7,CU_DEBIT8,CU_DEBIT9,CU_DEBIT10,CU_DEBIT11,CU_DEBIT12,CU_CREDIT1,CU_CREDIT2,'+
              'CU_CREDIT3,CU_CREDIT4,CU_CREDIT5,CU_CREDIT6,CU_CREDIT7,CU_CREDIT8,CU_CREDIT9,CU_CREDIT10,'+
              'CU_CREDIT11,CU_CREDIT12,CU_CREDITAN,CU_DEBITAN,CU_QUALIFQTEDIFF)'+
              'VALUES ("CON", "'+QCumulsCU_COMPTE1.AsString+'", '+
              '"", "'+CU_EXERCICE.Value+'", "-", "'+CU_ETABLISSEMENT.Value+'", "", '+
              '"", '+StrfPoint(TabDeb[1])+', '+
              ''+StrfPoint(TabDeb[2])+', '+StrfPoint(TabDeb[3])+', '+StrfPoint(TabDeb[4])+', '+
              ''+StrfPoint(TabDeb[5])+', '+StrfPoint(TabDeb[6])+', '+StrfPoint(TabDeb[7])+', '+
              ''+StrfPoint(TabDeb[8])+', '+StrfPoint(TabDeb[9])+', '+StrfPoint(TabDeb[10])+', '+
              ''+StrfPoint(TabDeb[11])+', '+StrfPoint(TabDeb[12])+', '+StrfPoint(TabCred[1])+', '+
              ''+StrfPoint(TabCred[2])+', '+StrfPoint(TabCred[3])+', '+StrfPoint(TabCred[4])+', '+
              ''+StrfPoint(TabCred[5])+', '+StrfPoint(TabCred[6])+', '+StrfPoint(TabCred[7])+', '+
              ''+StrfPoint(TabCred[8])+', '+StrfPoint(TabCred[9])+', '+StrfPoint(TabCred[10])+', '+
              ''+StrfPoint(TabCred[11])+', '+StrfPoint(TabCred[12])+', '+StrfPoint(0)+', '+
              ''+StrfPoint(0)+', "'+QCumulsCU_QUALIFQTEDIFF.AsString+'")')
else
   ExecuteSql('UPDATE CUMULS SET CU_ETABLISSEMENT="'+CU_ETABLISSEMENT.Value+'",'+
              'CU_DEVQTE="",CU_QUALIFPIECE="",'+
              'CU_DEBIT1='+StrfPoint(TabDeb[1])+',CU_DEBIT2='+StrfPoint(TabDeb[2])+','+
              'CU_DEBIT3='+StrfPoint(TabDeb[3])+',CU_DEBIT4='+StrfPoint(TabDeb[4])+','+
              'CU_DEBIT5='+StrfPoint(TabDeb[5])+',CU_DEBIT6='+StrfPoint(TabDeb[6])+','+
              'CU_DEBIT7='+StrfPoint(TabDeb[7])+',CU_DEBIT8='+StrfPoint(TabDeb[8])+','+
              'CU_DEBIT9='+StrfPoint(TabDeb[9])+',CU_DEBIT10='+StrfPoint(TabDeb[10])+','+
              'CU_DEBIT11='+StrfPoint(TabDeb[11])+',CU_DEBIT12='+StrfPoint(TabDeb[12])+','+
              'CU_CREDIT1='+StrfPoint(TabCred[1])+',CU_CREDIT2='+StrfPoint(TabCred[2])+','+
              'CU_CREDIT3='+StrfPoint(TabCred[3])+',CU_CREDIT4='+StrfPoint(TabCred[4])+','+
              'CU_CREDIT5='+StrfPoint(TabCred[5])+',CU_CREDIT6='+StrfPoint(TabCred[6])+','+
              'CU_CREDIT7='+StrfPoint(TabCred[7])+',CU_CREDIT8='+StrfPoint(TabCred[8])+','+
              'CU_CREDIT9='+StrfPoint(TabCred[9])+',CU_CREDIT10='+StrfPoint(TabCred[10])+','+
              'CU_CREDIT11='+StrfPoint(TabCred[11])+',CU_CREDIT12='+StrfPoint(TabCred[12])+','+
              'CU_CREDITAN='+StrfPoint(0)+',CU_DEBITAN='+StrfPoint(0)+','+
              'CU_QUALIFQTEDIFF= "" Where '+
              'CU_TYPE="CON" And CU_COMPTE1="'+QCumulsCU_COMPTE1.AsString+'" And '+
              'CU_COMPTE2="" And CU_EXERCICE="'+CU_EXERCICE.Value+'" And '+
              'CU_SUITE="-"') ;

EcrirSuite:=False ;
END ;

Procedure TFCumuls.ReloadQCumuls ;
Var St : String ;
BEGIN
St:='Select * From CUMULS Where CU_TYPE="CON" And CU_COMPTE1="'+Cpte1+'" And '+
    'CU_COMPTE2="" And CU_EXERCICE="'+Lexo+'" And CU_ETABLISSEMENT="'+Letab+'" ' ;
QCumuls.Close ; QCumuls.Sql.Clear ; QCumuls.Sql.Add(St) ;
ChangeSql(QCumuls) ; QCumuls.Open ;
ChercheSuite ;
END ;

Function TFCumuls.PresenceCle : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CU_TYPE,CU_COMPTE1,CU_COMPTE2,CU_EXERCICE,CU_ETABLISSEMENT from CUMULS Where '+
           'CU_TYPE="CON" And CU_COMPTE1="'+CU_COMPTE1.Text+'" And '+
           'CU_COMPTE2="" And CU_EXERCICE="'+CU_EXERCICE.Value+'" And '+
           'CU_ETABLISSEMENT="'+CU_ETABLISSEMENT.Value+'"',True) ;
Result:=Not Q.Eof ; Ferme(Q) ;
END ;

Function TFCumuls.ChercheEnreg : Boolean;
Var Q : TQuery ;
    Stetab,StExo,St : String ;
    Rep : Integer ;
BEGIN
Result:=True ;
if CU_ETABLISSEMENT.Text='' then SteTab:=QCumulsCU_ETABLISSEMENT.AsString
                            else SteTab:=CU_ETABLISSEMENT.Value ;
if CU_EXERCICE.Text='' then StExo:=QCumulsCU_EXERCICE.AsString
                       else StExo:=CU_EXERCICE.Value ;
St:='Select CU_TYPE,CU_COMPTE1,CU_COMPTE2,CU_EXERCICE,CU_ETABLISSEMENT from CUMULS Where '+
    'CU_TYPE="CON" And CU_COMPTE1="'+CU_COMPTE1.Text+'" And '+
    'CU_COMPTE2="" And CU_EXERCICE="'+StExo+'" '+
    'And CU_ETABLISSEMENT="'+Stetab+'"' ;
Q:=OpenSQL(St,TRUE) ;
if Not Q.Eof then
   BEGIN
   Cpte1:=Q.Fields[1].AsString ;
   Lexo:=Q.Fields[3].AsString ;
   Letab:=Q.Fields[4].AsString ;
   ReloadQCumuls ; LitDebCred ;
   InitCaption(Self,Lequel,Lib) ;
   END else
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(11,'','') ;
   Case rep of
       mrYes : BEGIN
               QCumuls.Insert ; NewEnreg ; //Bouge(nbInsert) ;
               END ;
       mrNo,mrCancel : Result:=False ;
      End ;
   END ;
Ferme(Q) ;
END ;

procedure TFCumuls.CU_ETABLISSEMENTChange(Sender: TObject);
begin if Not ChercheEnreg then Close ; end;

procedure TFCumuls.QCumulsBeforePost(DataSet: TDataSet);
begin
Cpte1:=QCumulsCU_COMPTE1.AsString ; Lexo:=QCumulsCU_EXERCICE.AsString ;
Letab:=QCumulsCU_ETABLISSEMENT.AsString ;
end;

Procedure TFCumuls.PositionneExo ;
Var i,j : Integer ;
begin
Exo.Code:=CU_EXERCICE.Value ;
if Exo.Code<>'' then
   BEGIN
   RempliExoDate(Exo) ; NombreMois:=Exo.NombrePeriode ; RempliGridPeriode ;
   if (NombreMois>12)then
       BEGIN
       j:=13 ;
       for i:=j to Fliste.RowCount-1 do
           if Fliste.Cells[1,i]='' then Fliste.Cells[1,i]:=IntToStr(0) ;
       END ;
   if Not ChercheEnreg then Afermer:=True
                       else Afermer:=False ;
   END else EffaceLeGrid ;
END ;

procedure TFCumuls.CU_EXERCICEChange(Sender: TObject);
BEGIN PositionneExo ; if AFermer then Close ; end;

procedure TFCumuls.BChercheClick(Sender: TObject);
begin if Not ChercheEnreg then Close ; end;

procedure TFCumuls.FormClose(Sender: TObject; var Action: TCloseAction);
begin FListe.VidePile(False) ; end;

procedure TFCumuls.WMGetMinMaxInfo(var MSG: Tmessage) ;
BEGIN
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
END ;

procedure TFCumuls.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; end;

procedure TFCumuls.BImprimerClick(Sender: TObject);
begin
XX_WHERE.Hint:=' Where CU_TYPE="CON" And CU_COMPTE1="'+Lequel+'" And CU_EXERCICE="'+QCumulsCU_EXERCICE.AsString+'" '+
               ' And CU_ETABLISSEMENT="'+QCumulsCU_ETABLISSEMENT.AsString+'"' ;
PrintDBGrid(Nil,XX_WHERE,Caption,'PRT_CONSTANTE') ;
end;

procedure TFCumuls.FormResize(Sender: TObject);
begin
Fliste.Colwidths[1]:=FListe.Width-178 ;
end;

procedure TFCumuls.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
