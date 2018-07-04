unit ParaSpBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, hmsgbox, StdCtrls, Hctrls, ComCtrls, Buttons, ExtCtrls, HEnt1,
  ImgList, HImgList;

Procedure ParamSousPlanBudget(Cod,Lib : String ; Var Axe,SousPlan : String) ;

type
  TFParaSpBu = class(TForm)
    HPB: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    Panel2: TPanel;
    TCC_CODE: TLabel;
    Cod: TEdit;
    TCC_LIBELLE: TLabel;
    Lib: TEdit;
    TvSect: TTreeView;
    TvJal: TTreeView;
    TSectBud: TLabel;
    SectBud: TEdit;
    JalBud: TEdit;
    TCbAxe: TLabel;
    CbAxe: THValComboBox;
    TJalBud: TLabel;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    Image: THImageList;
    BRazSec: THBitBtn;
    BRazJal: THBitBtn;
    procedure FormShow(Sender: TObject);
    procedure BRazSecClick(Sender: TObject);
    procedure BRazJalClick(Sender: TObject);
    procedure CbAxeChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure TvSectDblClick(Sender: TObject);
    procedure TvSectChange(Sender: TObject; Node: TTreeNode);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Axe,SousPlan : String ;
    Procedure ChargeUnTv(UnTv : TTreeView) ;
    Procedure ChargeLesEdit(Sender : TObject) ;
    Procedure NeVaPasAuTitre(Sender : TObject) ;
    Function  ControleSimple(St : String) : Boolean ;
    Function  ControleDouble(St,St1 : String) : Boolean ;
    Function  ControlMaximum(St : String) : Boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  Ent1;

Procedure ParamSousPlanBudget(Cod,Lib : String ; Var Axe,SousPlan : String) ;
var FParaSpBu : TFParaSpBu ;
BEGIN
FParaSpBu:=TFParaSpBu.Create(Application) ;
 Try
  FParaSpBu.Cod.Text:=Cod ;
  FParaSpBu.Lib.Text:=Lib ;
  FParaSpBu.Axe:=Axe ;
  FParaSpBu.SousPlan:=SousPlan ;
  FParaSpBu.ShowModal ;
  Axe:=FParaSpBu.Axe ;
  SousPlan:=FParaSpBu.SousPlan ;
 Finally
  FParaSpBu.Free ;
 End ;
SourisNormale ;
END ;

procedure TFParaSpBu.BRazSecClick(Sender: TObject);
begin SectBud.Text:='' ; end;

procedure TFParaSpBu.BRazJalClick(Sender: TObject);
begin JalBud.Text:='' ; end;

procedure TFParaSpBu.FormShow(Sender: TObject);
begin
if Axe<>'' then CbAxe.Value:=Axe else CbAxe.Value:=CbAxe.Values[0] ;
SectBud.Text:=Copy(SousPlan,1,Pos('/',SousPlan)-1) ;
JalBud.Text:=Copy(SousPlan,Pos('/',SousPlan)+1,Length(SousPlan)) ;
end;

procedure TFParaSpBu.CbAxeChange(Sender: TObject);
begin
ChargeUnTv(TvSect) ; ChargeUnTv(TvJal) ;
end;

Procedure TFParaSpBu.ChargeUnTv(UnTv : TTreeView) ;
Var T1,T2 : TTreeNode ;
    UnFb : TFichierBase ;
    i : Integer ;
BEGIN
UnTv.Items.Clear ;
T1:=Nil ; //T2:=Nil ;
T1:=UnTv.Items.Add(T1,HM.Mess[0]) ;
T1.ImageIndex:=0 ; T1.SelectedIndex:=T1.ImageIndex ;
UnFb:=AxeToFb(CbAxe.Value) ;
if Not VH^.Cpta[UnFb].Structure then Exit ;
for i:=1 to MaxSousPlan do
   BEGIN
   if VH^.SousPlanAxe[UnFb,i].Code<>'' then
      BEGIN
      T2:=UnTv.Items.AddChild(T1,VH^.SousPlanAxe[UnFb,i].Code+'  '+VH^.SousPlanAxe[UnFb,i].Lib) ;
      T2.ImageIndex:=1 ; T2.SelectedIndex:=T2.ImageIndex ;
      END ;
   END ;
UnTv.Items.GetFirstNode.Expand(False) ; UnTv.Selected:=UnTv.Items.GetFirstNode.GetNext ;
END ;

procedure TFParaSpBu.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFParaSpBu.ChargeLesEdit(Sender : TObject) ;
Var Edit : TEdit ;
    St : String ;
BEGIN
if TTreeView(Sender).Selected.Text=HM.Mess[0] then Exit ;
if TTreeView(Sender).Name='TvSect' then Edit:=SectBud else Edit:=JalBud ;
St:=TTreeView(Sender).Selected.Text ;
St:=Copy(St,1,Pos(' ',St)-1) ;
Edit.Text:=Edit.Text+St+';' ;
END ;

procedure TFParaSpBu.TvSectDblClick(Sender: TObject);
begin ChargeLesEdit(Sender) ; end;

procedure TFParaSpBu.TvSectChange(Sender: TObject; Node: TTreeNode);
begin NeVaPasAuTitre(Sender) ; end;

Procedure TFParaSpBu.NeVaPasAuTitre(Sender : TObject) ;
BEGIN
if TTreeView(Sender).Selected.Text=HM.Mess[0] then
   if TTreeView(Sender).Items.GetFirstNode.GetNext<>Nil then TTreeView(Sender).Selected:=TTreeView(Sender).Items.GetFirstNode.GetNext ;
END ;

Function TFParaSpBu.ControleSimple(St : String) : Boolean ;
Var St1 : String ;
BEGIN
Result:=False ;
if St='' then BEGIN Result:=True ; Exit ; END ;
While St<>'' do BEGIN St1:=ReadTokenSt(St) ; if Pos(St1,St)>0 then Exit ; END ;
Result:=True ;
END ;

Function TFParaSpBu.ControleDouble(St,St1 : String) : Boolean ;
Var St2 : String ;
BEGIN
Result:=False ;
if (St='') or (St1='') then BEGIN Result:=True ; Exit ; END ;
While St<>'' do BEGIN St2:=ReadTokenSt(St) ; if Pos(St2,St1)>0 then Exit ; END ;
Result:=True ;
END ;

Function TFParaSpBu.ControlMaximum(St : String) : Boolean ;
Var i,j : Integer ;
BEGIN
Result:=False ; j:=0 ;
for i:=1 to Length(St) do if St[i]=';' then Inc(j) ;
if j>MaxSousPlan then Exit ;
Result:=True ;
END ;

procedure TFParaSpBu.BValiderClick(Sender: TObject);
begin
//if HM.Execute(2,'','')<>mrYes then Exit ;
if Not ControlMaximum(SectBud.Text) then BEGIN HM.Execute(4,'',IntToStr(MaxSousPlan)) ; TvSect.SetFocus ; Exit ; END ;
if Not ControlMaximum(JalBud.Text)  then BEGIN HM.Execute(4,'',IntToStr(MaxSousPlan)) ; TvJal.SetFocus ; Exit ; END ;
if Not ControleSimple(SectBud.Text) then BEGIN HM.Execute(3,'','') ; TvSect.SetFocus ; Exit ; END ;
if Not ControleSimple(JalBud.Text)  then BEGIN HM.Execute(3,'','') ; TvJal.SetFocus ; Exit ; END ;
if Not ControleDouble(SectBud.Text,JalBud.Text) then BEGIN HM.Execute(3,'','') ; TvJal.SetFocus ; Exit ; END ;
SousPlan:=SectBud.Text+'/'+JalBud.Text ; Axe:=CbAxe.Value ; Close ;
end;

procedure TFParaSpBu.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
