{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 18/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit BudJalSP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, hmsgbox, HSysMenu, Buttons, ExtCtrls, ComCtrls,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HCtrls, HEnt1, Ent1, ImgList, HImgList  ,UentCommun ;

Function ChoisirSousSectionBudget(LaCatBud : TUneCatBud; ValSousPlan : String) : String ;

type
  TBudJalSP = class(TForm)
    Tv: TTreeView;
    HPB: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    Image: THImageList;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TvChange(Sender: TObject; Node: TTreeNode);
    procedure TvDblClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    LaCatBud : TUneCatBud ;
    ValSousPlan : String ;
    LeResultat : String ;
    WMinX,WMinY    : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure ChargeSousPlanBudget ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Function ChoisirSousSectionBudget(LaCatBud : TUneCatBud; ValSousPlan : String) : String ;
var FBudJalSP : TBudJalSP ;
BEGIN
Result:='' ;
//if Not(V_PGI.Cpta[UnFb].Structure) then Exit ;
FBudJalSP:=TBudJalSP.Create(Application) ;
 Try
  FBudJalSP.LaCatBud:=LaCatBud ;
  FBudJalSP.ValSousPlan:=ValSousPlan ;
  FBudJalSP.ShowModal ;
  Result:=FBudJalSP.LeResultat ;
 Finally
  FBudJalSP.Free ;
 End ;
SourisNormale ;
END ;

procedure TBudJalSP.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; END ;

procedure TBudJalSP.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; end;

procedure TBudJalSP.FormShow(Sender: TObject);
begin
ChargeSousPlanBudget ;
end;

Procedure TBudJalSP.ChargeSousPlanBudget ;
Var Tn,Tn1,Tn2 : TTReeNode ;
    i,j,k,l : Integer ;
    St,St2,Code,Lib,StSousPlan : String ;
    SousPlan : TSousPlan ;
BEGIN
Tn:=Nil ; Tn:=Tv.Items.AddObject(Tn,HM.Mess[0],Nil) ;
For l:=1 To MaxSousPlan Do
  BEGIN
  St:=LaCatBud.SurJal[l] ;
  If St<>'' Then
     BEGIN
     j:=0 ;
     For i:=1 To MaxSousPlan Do If VH^.SousPlanAxe[LaCatBud.fb,i].Code=St Then j:=i ;
     If j<>0 Then
        BEGIN
        SousPlan:=VH^.SousPlanAxe[LaCatBud.fb,j] ;
        Tn1:=Tv.Items.AddChildObject(Tn,SousPlan.Lib,Nil) ;
        Tn1.ImageIndex:=1 ; Tn1.SelectedIndex:=Tn1.ImageIndex ;
        StSousPlan:=ReadTokenSt(ValSousPlan) ;
        For k:=0 To SousPlan.ListeSP.Count-1 Do
          BEGIN
          St2:=SousPlan.ListeSP.Strings[k] ;
          Code:=ReadTokenSt(St2) ; Lib:=ReadTokenSt(St2) ;
          Tn2:=Tv.Items.AddChildObject(Tn1,Code+' - '+Lib,Nil) ;
          If Code=StSousPlan Then Tn2.ImageIndex:=3 Else Tn2.ImageIndex:=2 ;
          Tn2.SelectedIndex:=Tn2.ImageIndex ;
          END ;
        END ;
     END ;
  END ;
Tv.Items.GetFirstNode.Expand(True) ; Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
END ;

procedure TBudJalSP.TvChange(Sender: TObject; Node: TTreeNode);
begin
if Tv.Selected.Text=HM.Mess[0] then
   if Tv.Items.GetFirstNode.GetNext<>Nil then Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
end;

procedure TBudJalSP.TvDblClick(Sender: TObject);
Var Code : String ;
    Tn1,Tn2  : TTreeNode ;
    i,NoImage : Integer ;
begin
If (Tv.Selected.Level=2) Then
   BEGIN
   Code:=Tv.Selected.Text ;
   Tn1:=Tv.Selected.Parent ;
   For i:=0 To Tn1.Count-1 Do
      BEGIN
      Tn2:=Tn1.Item[i] ;
      If Tn2.Text<>Code Then NoImage:=2 Else NoImage:=3 ;
      Tn2.ImageIndex:=NoImage ; Tn2.SelectedIndex:=Tn2.ImageIndex ;
      END ;
   END ;
Tv.Refresh ;
end;

procedure TBudJalSP.BValiderClick(Sender: TObject);
Var St,St1 : String ;
    Tn : TTreeNode ;
    i,j,k,Lg : Integer ;
begin
Tn:=tv.Items.GetFirstNode ; If Tn<>NIL Then Tn:=Tn.GetNext ; St:='' ; i:=0 ;
While (Tn<>NIL) do
  BEGIN
  If Tn.Level=1 Then Inc(i) ;
  if i>MaxSousPlan then Break ;
  If (Tn.Level=2) And (Tn.ImageIndex=3) Then
     BEGIN
     j:=0 ; St1:=LaCatBud.SurJal[i] ;
     For k:=1 To MaxSousPlan Do If VH^.SousPlanAxe[LaCatBud.fb,k].Code=St1 Then j:=k ;
     If j<>0 Then
        BEGIN
        Lg:=VH^.SousPlanAxe[LaCatBud.fb,j].Longueur ; St:=St+Copy(Tn.Text,1,Lg)+';' ;
        END ;
     END ;
  Tn:=Tn.GetNext ;
  END ;
LeResultat:=St ;
end;

procedure TBudJalSP.BFermeClick(Sender: TObject);
begin
LeResultat:='' ;
end;

end.
