unit testPCL ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mask, HTB97, Hcompte, StdCtrls, ExtCtrls, Hctrls, ComCtrls, CritEdt,
  Menus, HSysMenu, hmsgbox, HPanel, UiUtil, Filtre, Ent1, HEnt1, ParamDat,
  Spin ;

Procedure EtatPCL ;

type
  TEtatPCL = class(TForm)
    Pages: TPageControl;
    Standards: TTabSheet;
    Dock971: TDock97;
    PanelFiltre: TToolWindow97;
    FFiltres: TComboBox;
    HPB: TToolWindow97;
    BParamListe: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    GB1: TGroupBox;
    HLabel4: THLabel;
    Exo1: THValComboBox;
    HLabel6: THLabel;
    FD11: TMaskEdit;
    Label7: TLabel;
    FD12: TMaskEdit;
    OkCol1: TCheckBox;
    FAvecPourcent: TCheckBox;
    FAvecDetail: TCheckBox;
    LFile: THLabel;
    FMaq: TEdit;
    RechFile: TToolbarButton97;
    Complements: TTabSheet;
    HLabel7: THLabel;
    HLabel8: THLabel;
    FDevises: THValComboBox;
    FEtab: THValComboBox;
    FAffiche: TRadioGroup;
    Sit1: TCheckBox;
    GB2: TGroupBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    Label1: TLabel;
    Exo2: THValComboBox;
    FD21: TMaskEdit;
    FD22: TMaskEdit;
    OkCol2: TCheckBox;
    Sit2: TCheckBox;
    GB3: TGroupBox;
    HLabel3: THLabel;
    HLabel9: THLabel;
    Label2: TLabel;
    Exo3: THValComboBox;
    FD31: TMaskEdit;
    FD32: TMaskEdit;
    OkCol3: TCheckBox;
    Sit3: TCheckBox;
    GB4: TGroupBox;
    HLabel10: THLabel;
    HLabel11: THLabel;
    Label3: TLabel;
    Exo4: THValComboBox;
    FD41: TMaskEdit;
    FD42: TMaskEdit;
    OkCol4: TCheckBox;
    Sit4: TCheckBox;
    FAvec: TGroupBox;
    FSimu: TCheckBox;
    FSitu: TCheckBox;
    FRevi: TCheckBox;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    Sauve: TSaveDialog;
    FVoirNom: TCheckBox;
    Type1: THValComboBox;
    Type2: THValComboBox;
    Type3: THValComboBox;
    Type4: THValComboBox;
    Formule1: TEdit;
    Formule2: TEdit;
    Formule3: TEdit;
    Formule4: TEdit;
    TFRESOL: THLabel;
    FRESOL: THValComboBox;
    HLabel5: THLabel;
    FFormat: TEdit;
    FVariation: TRadioGroup;
    FValVar1: TSpinEdit;
    FValVar2: TSpinEdit;
    FValVar3: TSpinEdit;
    FValVar4: TSpinEdit;
    procedure FFiltresChange(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OkCol1Click(Sender: TObject);
    procedure Exo1Change(Sender: TObject);
    procedure RechFileClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure FD11KeyPress(Sender: TObject; var Key: Char);
    procedure Type1Change(Sender: TObject);
    procedure Formule1Enter(Sender: TObject);
    procedure Formule1Exit(Sender: TObject);
    procedure FVariationClick(Sender: TObject);
  private
    { Déclarations privées }
    Crit : TCritEdtPCL ;
    FNomFiltre : String ;
    IndF : Integer ;
    Procedure ActiveCol(GB : tGroupBox ; OkOk : Boolean) ;
    procedure RecupCrit ;
    Procedure SwapCol(GB : tGroupBox ; OnRub : Boolean) ;
    procedure AjoutePourcentCrit ;

  public
    { Déclarations publiques }
  end;

implementation

uses GridPCL, UtilEdt ;

{$R *.DFM}

Procedure EtatPCL ;
var XX : TEtatPCL ;
    PP : THPanel ;
BEGIN
XX:=TEtatPCL.Create(Application) ;
XX.FNomFiltre:='ETATPCL' ;
//XX.HelpContext:=7802000 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     XX.ShowModal ;
    Finally
     XX.Free ;
    End ;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;
END ;

Procedure TEtatPCL.ActiveCol(GB : tGroupBox ; OkOk : Boolean) ;
var I : Integer;
    ChildControl: TControl;
begin
If GB=Nil Then Exit ;
for I:= 0 to GB.ControlCount -1 do
  BEGIN
  ChildControl:=GB.Controls[I];
  If ChildControl<>NIl Then
    BEGIN
    If ChildControl.Tag=0 Then ChildControl.Enabled:=OkOk ;
    END ;
  END ;
end;

Procedure TEtatPCL.SwapCol(GB : tGroupBox ; OnRub : Boolean) ;
var I : Integer;
    ChildControl: TControl;
begin
If GB=Nil Then Exit ;
for I:= 0 to GB.ControlCount -1 do
  BEGIN
  ChildControl:=GB.Controls[I];
  If ChildControl<>NIl Then
    BEGIN
    If (Pos('Type',ChildControl.Name)=0) And (ChildControl.Tag=0) Then
      BEGIN
      If Pos('Formule',ChildControl.Name)>0 Then ChildControl.Visible:=Not onRub Else ChildControl.Visible:=OnRub ;
      END ;
    END ;
  END ;
end;

procedure TEtatPCL.FFiltresChange(Sender: TObject);
begin
LoadFiltre(FNomFiltre,FFiltres,Pages) ;
end;

procedure TEtatPCL.BCreerFiltreClick(Sender: TObject);
begin NewFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TEtatPCL.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre(FNomFiltre,FFiltres,Pages) ; end;

procedure TEtatPCL.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre(FNomFiltre,FFiltres) ; end;

procedure TEtatPCL.BRenFiltreClick(Sender: TObject);
begin RenameFiltre(FNomFiltre,FFiltres) ; end;

procedure TEtatPCL.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres,Pages) ; end;



procedure TEtatPCL.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TEtatPCL.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TEtatPCL.FormShow(Sender: TObject);
Var i : Integer ;
    HIV : tHvalComboBox ;
    SP : TSpinEdit ;
begin
Pages.ActivePage:=Standards ;
InitResolution(FRESOL) ;
OkCol1.Checked:=TRUE ; OkCol2.Checked:=FALSE ; OkCol3.Checked:=FALSE ; OkCol4.Checked:=FALSE ;
OkCol1Click(OkCol1) ; OkCol1Click(OkCol2) ;
OkCol1Click(OkCol3) ; OkCol1Click(OkCol4) ;

For i:=1 To 4 Do
  BEGIN
  HIV:=THvalComboBox(FindComponent('Type'+IntToStr(i))) ;
  If HIV<>NIL Then HIV.Value:='RUB' ;
  SP:=TSpinEdit(FindComponent('FValVar'+IntToStr(i))) ;
  If SP<>NIL Then SP.Value:=-1 ;
  END ;
ChargeFiltre(FNomFiltre,FFiltres,Pages) ;
end;

procedure TEtatPCL.OkCol1Click(Sender: TObject);
Var St : String ;
    CB : tCheckBox ;
    GB : TGroupBox ;
begin
CB:=TCheckBox(Sender) ; If CB=NIL Then Exit ;
If CB.Tag>0 Then
  BEGIN
  GB:=TGroupBox(FindComponent('GB'+IntToStr(CB.Tag))) ; If GB<>NIL Then ActiveCol(GB,CB.Checked) ;
  END ;
end;

procedure TEtatPCL.Exo1Change(Sender: TObject);
Var HIV : THValComboBox ;
    GB : tGroupBox ;
    FD1,FD2 : tMaskEdit ;
begin
HIV:=THValComboBox(Sender) ; If HIV=NIL Then Exit ;
GB:=TGroupBox(HIV.Parent) ; If GB=NIL Then Exit ;
FD1:=tMaskEdit(FindComponent('FD'+IntToStr(GB.Tag)+'1')) ; If FD1=NIL Then Exit ;
FD2:=tMaskEdit(FindComponent('FD'+IntToStr(GB.Tag)+'2')) ; If FD2=NIL Then Exit ;
ExoToDates (HIV.Value,FD1,FD2) ;
end;

procedure TEtatPCL.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FMaq.Text) ;
if Sauve.Execute then FMaq.Text:=Sauve.FileName ;
end ;

Function ConvStUneDate(d : tDateTime) : String ;
Var YY,MM,DD : Word ;
BEGIN
DecodeDate(d,YY,MM,DD) ;
Result:='('+FormatFloat('00',DD)+'-'+FormatFloat('00',MM)+'-'+FormatFloat('0000',YY)+')' ;
END ;

Function QuelTitre(Date1,Date2 : tDateTime) : String ;
Var YY2,MM2,DD2,YY1,MM1,DD1 : Word ;
    Deb,Fin : Boolean ;
    St : String ;
BEGIN
DecodeDate(Date2,YY2,MM2,DD2) ; DecodeDate(Date1,YY1,MM1,DD1) ; St:='' ;
Deb:=(DebutDeMois(Date1)=Date1) ;
Fin:=(FinDeMois(Date2)=Date2) ;
If (MM1=MM2) And (YY1=YY2) Then
  BEGIN
  If Deb And Fin Then
    BEGIN
    St:=FormatDateTime('MMMM YYYY',Date1) ;
    END Else
    BEGIN
    St:=FormatDateTime('MMM YYYY',Date1)+'('+FormatFloat('00',DD1)+'-'+FormatFloat('00',DD2) ;
    END ;
  END Else
  BEGIN
  If Deb And Fin Then
    BEGIN
    St:=FormatDateTime('MMM YYYY',Date1)+'-'+FormatDateTime('MMM YYYY',Date2) ;
    END Else
    BEGIN
    St:=FormatDateTime('DD-MMM-YY',Date1)+' à '+FormatDateTime('DD-MMM-YY',Date2) ;
    END ;
  END ;
Result:=St ;
END ;

Function ConvStDate(Var C : tColPCL) : String ;
Var i,j,k : Integer ;
    st : String ;
BEGIN
St:='-' ;
If (C.Exo.Deb=C.Date1) And (C.Exo.Fin=C.Date2) Then
  BEGIN
  If C.Exo.Code=VH^.EnCours.Code Then BEGIN Result:='N' ; C.StTitre:='N' ; END Else
   If C.Exo.Code=VH^.Precedent.Code Then BEGIN Result:='N-' ; C.StTitre:='N-1' ; END Else
    If C.Exo.Code=VH^.Suivant.Code Then BEGIN Result:='N+' ; C.StTitre:='N+1' ; END Else
       For i:=1 To 5 Do If C.Exo.Code=VH^.ExoClo[i].Code then
         BEGIN
         k:=1 ;
         For j:=5 DownTo i Do If VH^.ExoClo[j].code<>'' Then BEGIN St:=St+'-' ; Inc(k) ; END ;
         Result:='N'+St ; C.StTitre:='N-'+IntToStr(k) ;
         END ;
  END Else
  BEGIN
  St:=ConvStUneDate(C.Date1) ; St:=St+ConvStUneDate(C.Date2) ; Result:=St ;
  C.StTitre:=QuelTitre(C.Date1,C.Date2) ;
  END ;

END ;

procedure TEtatPCL.AjoutePourcentCrit ;
Var CritMirror : tCritEdtPCL ;
    i : Integer ; 
BEGIN
CritMirror:=Crit ;
For i:=1 To 4 Do If CritMirror.Col[i-1].Actif Then
  BEGIN
  Crit.Col[2*i-2]:=CritMirror.Col[i-1] ;
  Crit.Col[2*i-1]:=CritMirror.Col[i-1] ;
  With Crit.Col[2*i-1] Do
    BEGIN
    Actif:=TRUE ; Fillchar(Exo,SizeOf(Exo),#0) ;
    Date1:=0 ; Date2:=0 ;
    AvecHisto:=FALSE ;
    IdentHisto:='' ;
    StFormule:='=100*[COL'+IntToStr(2*i-1)+']/[BASE'+IntToStr(2*i-1)+']' ;
    StTitre:='Pourcentage' ;
    END ;
  Inc(Crit.NbColActif) ;
  END ;
END ;

procedure TEtatPCL.RecupCrit ;
Var i : Integer ;
    St : String ;
    CodeExo : String ;
    HIV : tHValComboBox ;
    OnRub : Boolean ;
BEGIN
Fillchar(Crit,SizeOf(Crit),#0) ;
For i:=1 To 4 Do With Crit Do
  BEGIN
  Col[i-1].Actif:=TCheckBox(FindComponent('OkCol'+IntToStr(i))).Checked ;
  If Col[i-1].Actif Then
    BEGIN
    HIV:=THvalComboBox(FindComponent('Type'+IntToStr(i))) ;
    If HIV<>NIL Then OnRub:=HIV.Value='RUB' ;
    If OnRub Then
      BEGIN
      CodeExo:=THvalComboBox(FindComponent('Exo'+IntToStr(i))).Value ;
      QuelDateDeExo(CodeExo,Col[i-1].Exo) ;
      Col[i-1].Date1:=StrToDate(TMaskEdit(FindComponent('FD'+IntToStr(i)+'1')).Text) ;
      Col[i-1].Date2:=StrToDate(TMaskEdit(FindComponent('FD'+IntToStr(i)+'2')).Text) ;
      Col[i-1].AvecHisto:=TCheckBox(FindComponent('Sit'+IntToStr(i))).Checked ;
      END Else
      BEGIN
      Col[i-1].StFormule:=TEdit(FindComponent('Formule'+IntToStr(i))).Text ;
      END ;
    END ;
  END ;
With Crit Do
  BEGIN
  AvecPourcent:=FAvecPourcent.Checked ;
  AvecDetail:=FAvecDetail.Checked ;
  Modele:=FMaq.Text ;
  If FEtab.ItemIndex>0 Then ETab:=FEtab.Value ;
  If FDevises.ItemIndex>0 Then Devise:=FDevises.Value ;
  TypEcr[Reel]:=TRUE ; TypEcr[Simu]:=FSimu.Checked ; TypEcr[Situ]:=FSitu.Checked ;
  TypEcr[Revi]:=FRevi.Checked ;
  StTypEcr:='' ;
  If TypEcr[Reel] Then StTypEcr:=StTypEcr+'N' ;
  If TypEcr[Simu] Then StTypEcr:=StTypEcr+'S' ;
  If TypEcr[Situ] Then StTypEcr:=StTypEcr+'U' ;
  If TypEcr[Revi] Then StTypEcr:=StTypEcr+'R' ;
  Resolution:=FResol.Value ;
  END ;
For i:=1 To 4 Do With Crit Do If Col[i-1].Actif Then
  BEGIN
  Inc(NbColActif) ;
  If Col[i-1].StFormule='' Then Col[i-1].StFormule:=ConvStDate(Col[i-1]) ;
  If FFormat.Text<>'' Then Col[i-1].StFormat:=FFormat.Text Else Col[i-1].StFormat:='#,##0.00' ;
  END ;
If Crit.AvecPourcent Then AjoutePourcentCrit ;
END ;

procedure TEtatPCL.BValiderClick(Sender: TObject);
Var SFormule,SFormat : Array Of String ;
    i : integer ;
begin
RecupCrit ;
SetLength(SFormule,Crit.NbColActif) ; SetLength(SFormat,Crit.NbColActif) ;
For i:=0 To Crit.NbColActif-1 Do If Crit.Col[i].Actif Then
  BEGIN
  SFormule[i]:=Crit.Col[i].StFormule ; SFormat[i]:=Crit.Col[i].StFormat ;
  END ;
LanceLiasse(FMaq.Text,FALSE,FVoirNom.Checked,sFormule,sFormat,Crit,FALSE) ;
SFormule:=Nil ; SFormat:=Nil ;
end;

procedure TEtatPCL.BParamListeClick(Sender: TObject);
begin
LanceLiasse(FMaq.Text,FALSE,FVoirNom.Checked,['N'],['#,##0.00'],Crit,TRUE) ;

end;

procedure TEtatPCL.FD11KeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TEtatPCL.Type1Change(Sender: TObject);
Var HIV : THValComboBox ;
    GB : tGroupBox ;
    FD1,FD2 : tMaskEdit ;
begin
HIV:=THValComboBox(Sender) ; If HIV=NIL Then Exit ;
GB:=TGroupBox(HIV.Parent) ; If GB=NIL Then Exit ;
SwapCol(GB,HIV.Value='RUB') ;
end;

procedure TEtatPCL.Formule1Enter(Sender: TObject);
Var Ed : TEdit ;
    GB : tGroupBox ;
    LeLeft,LeTop : Integer ;
    SP : TSpinEdit ;
begin
Ed:=TEdit(Sender) ; If Ed=Nil Then Exit ;
IndF:=StrToInt(Copy(Ed.Name,Length(Ed.Name),1)) ;
GB:=tGroupBox(ED.Parent) ;
LeTop:=GB.Top+Ed.Top+Ed.Height+Standards.Top+2 ; LeLeft:=Ed.Left+GB.Left+Standards.Left ;
If LeTop+FVariation.Height>Pages.Height Then LeTop:=GB.Top+Ed.Top+Standards.Top-FVariation.Height-2 ;


FVariation.Top:=LeTop ; FVariation.Left:=LeLeft ;
FVariation.Visible:=TRUE ;
//TEdit(FindComponent('Formule'+IntToStr(IndF))).SetFocus ;
SP:=TSpinEdit(FindComponent('FValVar'+IntToStr(IndF))) ;
If SP<>NIL Then FVariation.ItemIndex:=SP.Value ;
end;


procedure TEtatPCL.Formule1Exit(Sender: TObject);
Var Ac :TWincontrol ;
begin
AC:=Self.ActiveControl ; If Ac=NIL Then Exit ;
If (AC.Parent<>Nil) And (AC.Parent.Name=FVariation.Name) Then Else FVariation.Visible:=FALSE ;
end;

Function NBP(St1,St2 : String) : Word ;
Var Ex : tExoDate ;
    XX,YY,ZZ : Word ;
    Date1,Date2 : tDateTime ;
BEGIN
Date1:=StrToDate(St1) ;
Date2:=StrToDate(St2) ;
FillChar(Ex,SizeOf(Ex),#0) ; EX.Deb:=Date1 ; EX.Fin:=Date2 ;
NOMBREPEREXO (Ex,XX,YY,ZZ) ; Result:=ZZ ;
END ;

procedure TEtatPCL.FVariationClick(Sender: TObject);
Var Ed : TEdit ;
    i1,i2 : Integer ;
    j : Integer ;
    HIV : THValComboBox ;
    OnRub : Boolean ;
    NbP1,nbP2 : Word ;
    NbpMin,NbpMax : Word ;
    St : String ;
    SP : TSpinEdit ;
begin
Ed:=TEdit(FindComponent('Formule'+IntToStr(IndF))) ;
i1:=0 ; i2:=0 ; j:=1 ;
Repeat
  HIV:=THvalComboBox(FindComponent('Type'+IntToStr(j))) ;
  If HIV<>NIL Then OnRub:=HIV.Value='RUB' ;
  If OnRub Then BEGIN If i1=0 Then i1:=j Else i2:=j ; END ;
  Inc(J) ;
Until ((i1>0) And (i2>0)) Or (j>4) ;
If (i1=0) Or (i2=0) Then Exit ;
nbP1:=NBP(TMaskEdit(FindComponent('FD'+IntToStr(i1)+'1')).Text,TMaskEdit(FindComponent('FD'+IntToStr(i1)+'2')).Text) ;
nbP2:=NBP(TMaskEdit(FindComponent('FD'+IntToStr(i2)+'1')).Text,TMaskEdit(FindComponent('FD'+IntToStr(i2)+'2')).Text) ;
NbPMin:=Nbp1 ; NbPMax:=Nbp2 ; If NbP1>NbP2 Then BEGIN NbPMin:=Nbp2 ; NbPMax:=Nbp1 ; END ;
Case FVariation.ItemIndex of
  0 : BEGIN
      St:='=[COL'+IntToStr(i1)+']-[COL'+IntToStr(i2)+']' ;
      END ;
  1 : BEGIN
      St:='=(12*[COL'+IntToStr(i1)+']/'+IntToStr(Nbp1)+')-(12*[COL'+IntToStr(i2)+']/'+IntToStr(Nbp2)+')' ;
      END ;
  2 : BEGIN
      St:='=('+IntToStr(NbpMin)+'*[COL'+IntToStr(i1)+']/'+IntToStr(Nbp1)+')-('+IntToStr(NbpMin)+'*[COL'+IntToStr(i2)+']/'+IntToStr(Nbp2)+')' ;
      END ;
  3 : BEGIN
      St:='=('+IntToStr(NbpMax)+'*[COL'+IntToStr(i1)+']/'+IntToStr(Nbp1)+')-('+IntToStr(NbpMax)+'*[COL'+IntToStr(i2)+']/'+IntToStr(Nbp2)+')' ;
      END ;
  END ;
Ed.Text:=St ;
SP:=TSpinEdit(FindComponent('FValVar'+IntToStr(IndF))) ;
If SP<>NIL Then SP.Value:=FVariation.ItemIndex ;
end;

end.
