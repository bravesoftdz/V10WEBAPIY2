{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CabrSsec;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  hmsgbox, HSysMenu, StdCtrls, Buttons, ExtCtrls, ComCtrls,
  Ent1, HEnt1, HCtrls, ed_tools, ImgList, HImgList;

Type TcodeAbr = Class
       UnAbr : String ;
       PlanSec : String ;
       CodSouPlan : String ;
       LibSouPlan : String ;
       Deb,Lon : Integer ;
     End ;

Function ParamCodeAbrege(UnAxe : String ; LaListe : HTStrings ; Lequel : String) : Boolean ;

type
  TFCabrSsec = class(TForm)
    Tv: TTreeView;
    HPB: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    Image: THImageList;
    Pinfo: TPanel;
    TSlplan: TLabel;
    Slplan: TEdit;
    TSPlan: TLabel;
    SPlan: TEdit;
    TPlan: TLabel;
    Plan: TEdit;
    Pabr: TPanel;
    TCabr: TLabel;
    Cabr: TEdit;
    BRapatri: THBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TvChange(Sender: TObject; Node: TTreeNode);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CabrChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure BRapatriClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    WMinX,WMinY    : Integer ;
    UnAxe : String ;
    LaListe : HTStrings ;
    LeQuel : String ;
    OkListe : Boolean ;
    OnFerme : Boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure ChargeLeTv ;
    Function  UnWherSup : String ;
    Procedure RapatrieLesSousPlan ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Function ParamCodeAbrege(UnAxe : String ; LaListe : HTStrings ; LeQuel : String) : Boolean ;
var FCabrSsec : TFCabrSsec ;
BEGIN
FCabrSsec:=TFCabrSsec.Create(Application) ;
  Try
   FCabrSsec.UnAxe:=UnAxe ;
   FCabrSsec.LaListe:=LaListe ;
   FCabrSsec.LeQuel:=LeQuel ;
   FCabrSsec.ShowModal ;
   Result:=FCabrSsec.OkListe ;
  Finally
   FCabrSsec.Free ;
  End ;
SourisNormale ;
END ;

procedure TFCabrSsec.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; END ;

procedure TFCabrSsec.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; OkListe:=False ; OnFerme:=False ; end;

procedure TFCabrSsec.FormClose(Sender: TObject; var Action: TCloseAction);
begin  OnFerme:=True ; VideTree(Tv) ; end;

procedure TFCabrSsec.FormShow(Sender: TObject);
begin ChargeLeTv ; end;

Function TFCabrSsec.UnWherSup : String ;
Var St,Sql : String ;
BEGIN
Sql:='' ;
While LeQuel<>'' do
   BEGIN
   St:=ReadTokenSt(LeQuel) ;
   if St<>'' then Sql:=Sql+'Or SS_SOUSSECTION="'+St+'" ';
   END ;
Delete(Sql,1,2) ; Sql:='And '+Sql ;
Result:=Sql ;
END ;

Procedure TFCabrSsec.ChargeLeTv ;
Var QLoc,QLoc1 : TQuery ;
    Tn,Tn1,Tn2 : TTReeNode ;
    X : TcodeAbr ;
    Sql1,PlusWhere, SsSect : String ;
BEGIN
Tn:=Nil ; Sql1:='' ; PlusWhere:='' ;
Tn:=Tv.Items.AddObject(Tn,HM.Mess[0],Nil) ;
if LeQuel<>'' then PlusWhere:=UnWherSup ;
Sql1:='Select * From STRUCRSE Where SS_AXE="'+UnAxe+'"'+PlusWhere ;
QLoc:=OpenSql(Sql1,True) ;

While Not QLoc.Eof do
  BEGIN
  X:=TcodeAbr.Create ;
  X.UnAbr:='' ; X.PlanSec:=QLoc.FindField('SS_SOUSSECTION').AsString ;
  X.CodSouPlan:='' ; X.LibSouPlan:='' ; X.Deb:=0 ; X.Lon:=0 ;
  Tn1:=Tv.Items.AddChildObject(Tn,QLoc.FindField('SS_LIBELLE').AsString,X) ;
  Tn1.ImageIndex:=1 ; Tn1.SelectedIndex:=Tn1.ImageIndex ;
  SsSect := QLoc.FindField('SS_SOUSSECTION').AsString ;
  QLoc1 := OpenSQL('SELECT PS_CODE,PS_LIBELLE FROM SSSTRUCR WHERE PS_AXE="'+UnAxe+'" AND PS_SOUSSECTION="'+SsSect+'"', True);
  While Not QLoc1.Eof do
      BEGIN
      X:=TcodeAbr.Create ;
      X.UnAbr:='' ; X.PlanSec:=QLoc.FindField('SS_SOUSSECTION').AsString ;
      X.CodSouPlan:=QLoc1.Fields[0].AsString ;
      X.LibSouPlan:=QLoc1.Fields[1].AsString ;
      X.Deb:=QLoc.FindField('SS_DEBUT').AsInteger ;
      X.Lon:=QLoc.FindField('SS_LONGUEUR').AsInteger ;
      Tn2:=Tv.Items.AddChildObject(Tn1,QLoc1.Fields[0].AsString,X) ;
      Tn2.ImageIndex:=2 ; Tn2.SelectedIndex:=Tn2.ImageIndex ;
      QLoc1.Next ;
      END ;
  QLoc.Next ;
  END ;
Ferme(QLoc) ; Ferme(QLoc1) ;
Tv.Items.GetFirstNode.Expand(False) ; Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
END ;

procedure TFCabrSsec.TvChange(Sender: TObject; Node: TTreeNode);
begin
if OnFerme then Exit ;
if Tv.Selected.Text=HM.Mess[0] then
   if Tv.Items.GetFirstNode.GetNext<>Nil then Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
if (Tv.Selected.ImageIndex=1) or (Tv.Selected.Data=Nil) then Pabr.Enabled:=False else Pabr.Enabled:=True ;
if Tv.Selected.Data<>Nil then
   BEGIN
   Plan.Text:=TcodeAbr(Tv.Selected.Data).PlanSec ;
   SPlan.Text:=TcodeAbr(Tv.Selected.Data).CodSouPlan ;
   Slplan.Text:=TcodeAbr(Tv.Selected.Data).LibSouPlan ;
   Cabr.Text:=TcodeAbr(Tv.Selected.Data).UnAbr ;
   END else
   BEGIN
   Plan.Text:='' ; SPlan.Text:='' ; Slplan.Text:='' ; Cabr.Text:='' ;
   END ;
end;

procedure TFCabrSsec.CabrChange(Sender: TObject);
begin
if (Tv.Selected.Data<>Nil) and (Tv.Selected.ImageIndex>1) then
   BEGIN
   TcodeAbr(Tv.Selected.Data).UnAbr:=Cabr.Text ;
   if Cabr.Text='' then Tv.Selected.ImageIndex:=2 else Tv.Selected.ImageIndex:=3 ;
   Tv.Selected.SelectedIndex:=Tv.Selected.ImageIndex ;
   END ;
end;

procedure TFCabrSsec.BValiderClick(Sender: TObject);
Var i : Integer ;
    St : String ;
begin
if HM.Execute(1,'','')<>mrYes then Exit ;
LaListe.Clear ;
for i:=0 to Tv.Items.Count-1 do
   BEGIN
   St:='' ;
   if(Tv.Items[i].Data<>Nil) then
     BEGIN
     if (TcodeAbr(Tv.Items[i].Data).UnAbr='') and (Tv.Items[i].ImageIndex>1) then BEGIN HM.Execute(2,'','') ; OkListe:=False ; Exit ; END ;
     St:=IntToStr(TcodeAbr(Tv.Items[i].Data).Deb) ;
     St:=St+';'+IntToStr(TcodeAbr(Tv.Items[i].Data).Lon) ;
     St:=St+';'+TcodeAbr(Tv.Items[i].Data).CodSouPlan ;
     St:=St+';'+TcodeAbr(Tv.Items[i].Data).UnAbr ;
     LaListe.Add(St) ;
     END ;
   END ;
OkListe:=True ; Close ;
end;

procedure TFCabrSsec.FormResize(Sender: TObject);
begin
Tv.Width:=(Width Div 2)-10 ; Pinfo.Left:=(Width Div 2)-10 ; Pabr.Left:=Pinfo.Left ;
Pinfo.Height:=(Height Div 2)-40 ; Pabr.Height:=Pinfo.Height ; Pabr.Top:=Pabr.Height+15 ;
end;

Procedure TFCabrSsec.RapatrieLesSousPlan ;
Var i : Integer ;
BEGIN
for i:=0 to Tv.Items.Count-1 do
  BEGIN
  if (Tv.Items[i].ImageIndex<>2) And (Tv.Items[i].ImageIndex<>3) then Continue ;
  TcodeAbr(Tv.Items[i].Data).UnAbr:=Tv.Items[i].Text ;
  Tv.Items[i].ImageIndex:=3 ; Tv.Items[i].SelectedIndex:=Tv.Items[i].ImageIndex ;
  END ;
if (Tv.Selected.Data<>Nil) and (Tv.Selected.ImageIndex>1) then TvChange(Nil,Tv.Selected) ;
Tv.Refresh ;
END ;

procedure TFCabrSsec.BRapatriClick(Sender: TObject);
begin RapatrieLesSousPlan ; end;

procedure TFCabrSsec.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;   
end;

end.
