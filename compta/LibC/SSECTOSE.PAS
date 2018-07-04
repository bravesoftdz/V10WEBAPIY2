unit SSecToSe;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DB, DBTables, hmsgbox, HSysMenu, Buttons, ExtCtrls, ComCtrls,
  HCtrls, HEnt1, Ent1, ImgList ;

Function ChoisirSSectionDeTv(UnFb : TFichierBase ; UneSection : String ; OnJoker : Boolean) : String ;

Type TInfoSSect = Class
       PlanSec : String ;
       CodSouPlan : String ;
       LibSouPlan : String ;
       Deb,Lon : Integer ;
     End ;

type
  TFSSecToSe = class(TForm)
    Tv: TTreeView;
    HPB: TPanel;
    Panel1: TPanel;
    BAide: TBitBtn;
    BFerme: TBitBtn;
    BValider: TBitBtn;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    Image: TImageList;
    Pinfo: TPanel;
    TSlplan: TLabel;
    TSPlan: TLabel;
    TPlan: TLabel;
    Slplan: TEdit;
    SPlan: TEdit;
    Plan: TEdit;
    TCabr: TLabel;
    CodIn: TEdit;
    Label1: TLabel;
    CodOut: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TvChange(Sender: TObject; Node: TTreeNode);
    procedure TvDblClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    UnFb : TFichierBase ;
    UneSection : String ;
    OnJoker : Boolean ;
    LeResultat : String ;
    UnAxe : String ;
    WMinX,WMinY    : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure ChargeLeTv ;
    Function  ChercheUneSection(St : String) : Boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Function ChoisirSSectionDeTv(UnFb : TFichierBase ; UneSection : String ; OnJoker : Boolean) : String ;
var FSSecToSe : TFSSecToSe ;
BEGIN
Result:='' ;
if Not(V_PGI.Cpta[UnFb].Structure) then Exit ;
FSSecToSe:=TFSSecToSe.Create(Application) ;
 Try
  FSSecToSe.UnFb:=UnFb ;
  FSSecToSe.UneSection:=UneSection ;
  FSSecToSe.OnJoker:=OnJoker ;
  FSSecToSe.ShowModal ;
  Result:=FSSecToSe.LeResultat ;
 Finally
  FSSecToSe.Free ;
 End ;
SourisNormale ;
END ;

procedure TFSSecToSe.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; END ;

procedure TFSSecToSe.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; end;

procedure TFSSecToSe.FormResize(Sender: TObject);
begin
Tv.Width:=(Width Div 2)-10 ; Pinfo.Left:=(Width Div 2)+2 ; Pinfo.Width:=(Width Div 2)-10 ;
end;

procedure TFSSecToSe.FormShow(Sender: TObject);
begin
UnAxe:=FbToAxe(UnFb) ; CodIn.Text:=UneSection ; CodOut.Text:=UneSection ;
LeResultat:=UneSection ; ChargeLeTv ;
end;

Procedure TFSSecToSe.ChargeLeTv ;
Var QLoc,QLoc1 : TQuery ;
    Tn,Tn1,Tn2 : TTReeNode ;
    X : TInfoSSect ;
    Sql : String ;
BEGIN
Tn:=Nil ; Tn1:=Nil ; Tn2:=Nil ;
Tn:=Tv.Items.AddObject(Tn,HM.Mess[0],Nil) ;
QLoc:=OpenSql('Select * From STRUCRSE Where SS_AXE="'+UnAxe+'"',True) ;
Sql:='Select PS_CODE,PS_LIBELLE From SSSTRUCR Where PS_AXE="'+UnAxe+'" And '+
     'PS_SOUSSECTION=:SSECT' ;
QLoc1:=PrepareSQL(Sql) ;
While Not QLoc.Eof do
  BEGIN
  X:=TInfoSSect.Create ;
  X.PlanSec:=QLoc.FindField('SS_SOUSSECTION').AsString ;
  X.CodSouPlan:='' ; X.LibSouPlan:='' ; X.Deb:=0 ; X.Lon:=0 ;
  Tn1:=Tv.Items.AddChildObject(Tn,QLoc.FindField('SS_LIBELLE').AsString,X) ;
  Tn1.ImageIndex:=1 ; Tn1.SelectedIndex:=Tn1.ImageIndex ;
  QLoc1.Close ;
  QLoc1.Params[0].AsString:=QLoc.FindField('SS_SOUSSECTION').AsString ;
  QLoc1.Open ;
  While Not QLoc1.Eof do
      BEGIN
      X:=TInfoSSect.Create ;
      X.PlanSec:=QLoc.FindField('SS_SOUSSECTION').AsString ;
      X.CodSouPlan:=QLoc1.Fields[0].AsString ; X.LibSouPlan:=QLoc1.Fields[1].AsString ;
      X.Deb:=QLoc.FindField('SS_DEBUT').AsInteger ;
      X.Lon:=QLoc.FindField('SS_LONGUEUR').AsInteger ;
      Tn2:=Tv.Items.AddChildObject(Tn1,X.CodSouPlan+'  '+X.LibSouPlan,X) ;
      if OnJoker then
         BEGIN
         if Copy(CodIn.Text,X.Deb,1)='?' then Tn2.ImageIndex:=2
                                         else Tn2.ImageIndex:=3 ;
         END else Tn2.ImageIndex:=2 ;
      Tn2.SelectedIndex:=Tn2.ImageIndex ;
      QLoc1.Next ;
      END ;
  QLoc.Next ;
  END ;
Ferme(QLoc) ; Ferme(QLoc1) ;
Tv.Items.GetFirstNode.Expand(False) ; Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
END ;

procedure TFSSecToSe.TvChange(Sender: TObject; Node: TTreeNode);
begin
if Tv.Selected.Text=HM.Mess[0] then
   if Tv.Items.GetFirstNode.GetNext<>Nil then Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
if Tv.Selected.Data<>Nil then
   BEGIN
   Plan.Text:=TInfoSSect(Tv.Selected.Data).PlanSec ;
   SPlan.Text:=TInfoSSect(Tv.Selected.Data).CodSouPlan ;
   Slplan.Text:=TInfoSSect(Tv.Selected.Data).LibSouPlan ;
   END else
   BEGIN
   Plan.Text:='' ; SPlan.Text:='' ; Slplan.Text:='' ;
   END ;
end;

procedure TFSSecToSe.TvDblClick(Sender: TObject);
Var St : String ;
    Debu,Long : Integer ;
begin
if (Tv.Selected.Data=Nil) or (TInfoSSect(Tv.Selected.Data).CodSouPlan='') then Exit ;
St:=CodOut.Text ;
Debu:=TInfoSSect(Tv.Selected.Data).Deb ; Long:=TInfoSSect(Tv.Selected.Data).Lon ;
if OnJoker then if Copy(CodIn.Text,Debu,1)<>'?' then Exit ;
if St<>'' then
   BEGIN
   Delete(St,Debu,Long) ; Insert(TInfoSSect(Tv.Selected.Data).CodSouPlan,St,Debu) ;
   CodOut.Text:=St ;
   END ;
end;

Function TFSSecToSe.ChercheUneSection(St : String) : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select S_SECTION From SECTION Where S_SECTION="'+St+'" And S_AXE="'+UnAxe+'"',True) ;
Result:=Not QLoc.Eof ; Ferme(QLoc) ;
END ;

procedure TFSSecToSe.BValiderClick(Sender: TObject);
Var St : String ;
begin
St:=CodOut.Text ;
if Not ChercheUneSection(St) then
   BEGIN
   Case HM.Execute(1,'','') of
     mrYes : BEGIN
             if Assigned(ProcZoomSection) then ProcZoomSection(Nil,UnAxe,St,taCreatOne,0) ;
             if Not ChercheUneSection(St) then BEGIN HM.Execute(2,'','') ; Exit ; END ;
             END ;
     mrNo,mrCancel : Exit ;
     end ;
   END ;
LeResultat:=CodOut.Text ; Close ;
end;

end.
