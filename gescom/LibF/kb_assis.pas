unit kb_assis;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
{$IFDEF EAGLCLIENT}                  //NA debut 25/11/02
  UtileAGL,
{$ELSE}
  db, dbtables, MajTable,
{$ENDIF}                             //NA fin 25/11/02
{$IFDEF PGIS5}
  MC_Lib, FOUtil,    //NA debut 17/11/03
{$ELSE}
  Ut,
{$ENDIF}
  Hent1, HFLabel, KB_Ecran, HPanel ; //NA fin 17/11/03

Function CopieClavierFromSoc (Appellant : TForm ; TypesDef : String )  : Boolean ;
Function RecopieClavierEcran ( DBSource : TDataBase ; FromCaisse,TypesBtn : String  ; PnlBtn : TClavierEcran) : Boolean;

type
  TFCopieClavierSoc = class(TFAssist)
    PSoc: TTabSheet;
    Label6: TLabel;
    FSocSource: THValComboBox;
    Label12: TLabel;
    Label3: TLabel;
    Bevel2: TBevel;
    pCaisse: TTabSheet;
    pTypes: TTabSheet;
    grTypes: TGroupBox;
    Label13: TLabel;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    DBSource: TDatabase;
    pnlCaisse1: TPanel;
    Label1: TLabel;
    Bevel1: TBevel;
    TCbCaisse: TLabel;
    Label4: TLabel;
    cbCaisse: THValComboBox;
    PnlCaisse2: TPanel;
    LblAvis: TFlashingLabel;
    procedure bAnnulerClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure PChange(Sender: TObject); override ;
    procedure FSocSourceChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cbCaisseChange(Sender: TObject);
    procedure clickCheck(Sender: TObject);
  private
    OkRecharge : Boolean ;
    Procedure ChargeTypes ;
    Function  ChargeCaisses : boolean ;
    procedure setTypes( ATypes : String ) ;
    Function  AuMoinsUn : Boolean ;
    Procedure MiseajourPnls ;
  public
    Appellant : TForm ;
    TypesDef  : Array of String ;
  end;


implementation

{$R *.DFM}
// *******************************************
// JTR - eQualité 11771 - Ajout du HelpContext
// *******************************************


Function CopieClavierFromSoc (Appellant : TForm ; TypesDef : String )  : Boolean ;
var X : TFCopieClavierSoc ;
Begin
X:=TFCopieClavierSoc.Create(Application) ;
try
   X.Appellant:=Appellant ;
   X.SetTypes(TypesDef) ;
   Result:=(X.ShowModal=mrOk);
 finally
   X.Free ;
   SourisNormale ;
 End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
//OpenSQLDB à été déplacée à Ut.pas  //XMG 03/04/01
// et elle a été toute remplacée par la ligne suivante :

//Result:=DBOpenSQL(DB,SQL,StToDriver(DB.DriverName,FALSE)) ;

/////////////////////////////////////////////////////////////////////////////////////////
Function RecopieClavierEcran ( DBSource : TDataBase ; FromCaisse,TypesBtn : String  ; PnlBtn : TClavierEcran) : Boolean;
Var Q        : TQuery ;
    Where,St : String ;
    i,nump   : Integer ;
Begin
{$IFDEF PGIS5}
Q:=OpenSQLDB('select GPK_PARAMSCE from PARCAISSE where GPK_CAISSE="'+FromCaisse+'"',DBSource) ;
{$ELSE}
Q:=OpenSQLDB('select PC_PARAMSCE from PCAISSE where PC_CAISSE="'+FromCaisse+'"',DBSource) ;
{$ENDIF}
St:=Q.fields[0].AsString ;
Ferme(Q) ;
if (trim(St)='') or (NbCarInString(St,';')<>3) then St:=inttostr(CENbrBtnWidthDef)+';'+inttostr(CENbrBtnHeightDef)+';'+inttostr(ord(pcLeft))+';' ;
PnlBtn.NbrBtnWidth:=valeuri(readtokenst(St))  ;
PnlBtn.NbrBtnHeight:=Valeuri(ReadTokenst(St)) ;
PnlBtn.ClcPosition:=tPosClc(Valeuri(readtokenst(St))) ;
St:='' ; Where:='' ; Nump:=nbCarInString(TypesBtn,';') ;
if NumP>=1 then
   for i:=1 to NumP do
     Begin
     St:=gtfs(TypesBtn,';',i) ;
     if trim(St)<>'' then
        Begin
        if trim(where)<>'' then Where:=Where+' or ' ;
        Where:=Where+'(CE_TEXTE like "'+St+';%")' ;
        End ;
     End ;
if Trim(Where)<>'' then Where:=' and ('+Where+')' ;
Where:=' where CE_CAISSE="'+FromCaisse+'"'+Where ;
Q:=OpenSQLDB('select * from CLAVIERECRAN'+where,DBSource) ;
Result:=PnlBtn.ChargeFromQ(Q) ;
Ferme(Q) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.bAnnulerClick(Sender: TObject);
begin
  inherited;
ModalResult:=mrCancel ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.bFinClick(Sender: TObject);
Var Ok     : Boolean ;
    St,Nom : String ;
    PnlBtn : TClavierEcran ;
    i,nump : Integer ;
    Pages  : TPageControl ;
begin
Ok:=FALSE ;
if PGIAsk('Confirmez-vous la charge du pavé depuis la caisse choisie?',Caption)=mrYes then
   Begin
   Buzy:=TRUE ;
   try
       inherited;
       PnlBtn:=TClavierEcran(Appellant.FindCOmponent('KB_ECRAN')) ;
       Pages:=TPageControl(Appellant.FindComponent('pages')) ;
       if (Assigned(PnlBtn)) and (Assigned(Pages)) then
          Begin
          St:='' ;
          for i:=0 to GRTypes.ControlCount-1 do
            if (GRTypes.Controls[i] is TCheckBox) and (TCheckBox(GRTypes.Controls[i]).Checked) then
               Begin
               Nom:=GRTypes.Controls[i].Name ;
               NumP:=Valeuri(Copy(Nom,3,length(Nom))) ;
               St:=St+Pages.pages[NumP].Hint+';' ;
               End ;
          Ok:=RecopieClavierEcran(DBSource,cbCaisse.Value,St,PnlBtn) ;
          End ;
     Finally
      Buzy:=FALSE ;
     End ;
   End ;
if Ok then ModalResult:=MrOk ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.FormShow(Sender: TObject);
Var i    : integer ;
Begin
inherited ; 
ChargeDossier(FSocSource.Items, false) ;
i:=FSocSource.Items.indexof(V_PGI.CurrentAlias) ;
if i>-1 then FSocSource.Items.delete(i) ;
i:=FSocSource.Items.indexof(GetSocRef) ;
if i>-1 then FSocSource.ItemIndex:=i ;
ChargeTypes ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure TFCopieClavierSoc.ChargeTypes ;
Var i,CkParCol,j : Integer ;
    cb           : THValComboBox ;
    ck           : TCheckBox ;
Const ckHeight = 17 ;
      ckMargeT = 3 ;
      ckMargeB = 3 ;
      ckMargeL = 10 ;
      ckMargeR = 10 ;
begin
  inherited;
While GrTypes.ControlCount>0 do GrTypes.Controls[0].Free ;
if assigned(Appellant) then
   Begin
   CB:=THValComboBox(Appellant.FindComponent('CBConcept')) ;
   if assigned(CB) then
      Begin
      CkparCol:=(GRTypes.Height-ckMargeB) div (ckHeight+ckMargeT) ;
      j:=0 ;
      For i:=0 to CB.Items.Count-1 do
        if valeuri(CB.Values[i])<>0 then
           Begin
           inc(j) ;
           ck:=TCheckBox.Create(Self) ;
           ck.Parent:=GRTypes ;
           ck.Left:=CkMargeL+((GRTypes.Width-ckMargeL-ckMargeR) div 2) * ord(j>ckParCol) ;
           ck.Top:=(CkMargeT+ckHeight)*(j-(ckParCOl*ord(j>ckParCol))) ; 
           ck.Height:=ckHeight ;
           ck.Width:=(GRTypes.Width-ckMargeL-ckMargeR) div 2 ;
           ck.Name:='CK'+inttostr(i) ;
           ck.Caption:=CB.Items[i] ;
           ck.Tag:=i ;
           ck.Checked:=inString(CB.Values[i],TypesDef) ;
           ck.OnClick:=ClickCheck ;
           End ;
      End ;
   End ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
Function TFCopieClavierSoc.AuMoinsUn : Boolean ;
var i : integer ;
Begin
Result:=FALSE ;
i:=0 ;
While (i<=GRTypes.ControlCount-1) and (not Result)  do
   Begin
   if GRTypes.Controls[i] is TCheckBox then Result:=TCheckBox(GRTypes.Controls[i]).Checked ;
   inc(i) ;
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure TFCopieClavierSoc.MiseajourPnls ;
Begin
PnlCaisse1.Visible:=not OkRecharge ;
PnlCaisse2.Visible:=not PnlCaisse1.Visible ;
Application.ProcessMessages ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.PChange(Sender: TObject);
var OldDriver : TDBDriver ;
    OldOdbc,Ok        : Boolean ;
begin
MiseajourPnls ;
  inherited;
BSuivant.Visible:=(GetPage<>PTypes) ;
BFin.Visible:=(GetPage=PTypes) and (auMoinsUn) ;
BSuivant.Enabled:=(GetPage<>PCaisse) or ((GetPage=PCaisse) and (cbCaisse.Itemindex>-1)) ;
if (GetPage=PCaisse) and (OkRecharge) then
   Begin
   OldDriver:=V_PGI.Driver ;
   OldOdbc:=V_PGI.ODBC ;
   try
       Application.ProcessMessages ;
       Ok:=ConnecteDB(FSocSource.Items[FSocSource.ItemIndex],DBSource,'DBSource') ;
       Application.ProcessMessages ;
      Finally
       V_PGI.Driver:=OldDriver ;
       V_PGI.ODBC:=OldOdbc ;
      End ;
   if not Ok then
      Begin
      PgiBox('Impossible de se connecter à la société source '+ FSocSource.Items[FSocSource.ItemIndex],caption) ;
      BPrecedent.click ;
      End else
      if not ChargeCaisses then
         Begin
         PgiBox('Cette société n''a aucun pavé paramétré',Caption) ;
         BPrecedent.click ;
         End ;
   OkRecharge:=FALSE ;
   miseajourPnls ;
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.setTypes( ATypes : String ) ;
var i,j : integer ;
    St  : String ;
Begin
i:=nbCarinString(ATypes,';') ;
TypesDef:=Copy(TypesDef,0,0) ;
For j:=1 to i do
  Begin
  St:=gtfs(ATypes,';',j) ;
  if trim(St)<>'' then
     Begin
     SetLength(TypesDef,Length(TypesDef)+1) ;
     TypesDef[High(TypesDef)]:=St ;
     End ;
  End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Function TFCopieClavierSoc.ChargeCaisses : boolean ;
Var Q  : TQuery ;
    i  : integer ;
Begin
cbCaisse.Items.BeginUpdate ;
cbCaisse.Values.BeginUpdate ;
cbCaisse.Items.Clear ;
cbCaisse.Values.clear ;
{$IFDEF PGIS5}
Q:=OpenSQLDB('Select GPK_CAISSE, GPK_LIBELLE from PARCAISSE',DBSource) ;
{$ELSE}
Q:=OpenSQLDB('Select PC_CAISSE, PC_LIBELLE from PCAISSE',DBSource) ;
{$ENDIF}
while Not Q.Eof do
  Begin
  if trim(Q.Fields[0].AsString)<>'' then
     Begin
     cbCaisse.Values.Add(Q.Fields[0].AsString) ;
     cbCaisse.Items.add('<<VIDE>>') ;
     End ;
  Q.next ;
  End ;
Ferme(Q) ;
{$IFDEF PGIS5}   //NA 25/11/02
Q:=OpenSQLDB('Select GPK_CAISSE, GPK_LIBELLE from PARCAISSE',DBSource) ;
{$ELSE}
Q:=OpenSQLDB('Select PC_CAISSE, PC_LIBELLE from PCAISSE',DBSource) ;
{$ENDIF}        //NA 25/11/02
while Not Q.Eof do
  Begin
  i:=CBCaisse.Values.indexof(Q.Fields[0].AsString) ;
  if i>-1 then cbCaisse.Items[i]:=Q.Fields[1].AsString ;
  Q.next ;
  End ;
Ferme(Q) ;
repeat
  i:=cbCaisse.items.indexof('<<VIDE>>') ;
  if i>-1 then
     Begin
     cbCaisse.items.delete(i) ;
     cbCaisse.Values.delete(i) ;
     End ;
until i<0 ;
cbCaisse.Items.EndUpdate ;
cbCaisse.Values.EndUpdate ;
Result:=cbCaisse.Values.Count>0 ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.FSocSourceChange(Sender: TObject);
begin
  inherited;
okRecharge:=TRUE ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.FormCreate(Sender: TObject);
begin
  inherited;
OkRecharge:=TRUE ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.cbCaisseChange(Sender: TObject);
begin
  inherited;
BSuivant.Enabled:=CbCaisse.ItemIndex>-1 ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFCopieClavierSoc.clickCheck(Sender: TObject);
begin
  inherited;
BFin.Visible:=(GetPage=PTypes) and (auMoinsUn) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
end.
