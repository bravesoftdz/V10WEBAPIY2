unit AtChComp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, hmsgbox, StdCtrls, ComCtrls, Buttons, ExtCtrls, Ent1, HEnt1, HCtrls,
  HStatus, Hcompte, ed_tools, ImgList, HImgList ;

Type TTInfoComp = Class
       Origine : String ;
       Champ   : String ;
       OblG    : String ;
       Modif   : String ;
       ValDef  : String ;
       ValTl   : String ;
       Acces   : String ;
       EstCocher : String ;
     End ;

type
  TFAtChComp = class(TForm)
    PBottom: TPanel;
    FAutoSave: TCheckBox;
    PBouton: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    Tv: TTreeView;
    Image: THImageList;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    PRigth: TPanel;
    CbObligatoire: TCheckBox;
    CbModifiable: TCheckBox;
    TValdef: TLabel;
    Valdef: TEdit;
    TZoomTl: TLabel;
    BAssist: THBitBtn;
    ZoomTl: THValComboBox;
    ZoomTabli: THCpteEdit;
    CbValDef: TCheckBox;
    CbDate: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure TvChange(Sender: TObject; Node: TTreeNode);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAssistClick(Sender: TObject);
    procedure CbObligatoireClick(Sender: TObject);
    procedure CbModifiableClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure ZoomTabliExit(Sender: TObject);
    procedure CbValDefClick(Sender: TObject);
    procedure ZoomTlChange(Sender: TObject);
    procedure ValdefChange(Sender: TObject);
    procedure CbDateChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    LesAttribs : HTStrings ;
    OnFerme,AFermer : Boolean ;
    Modifier : Boolean ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure FaitNodes ;
    Function  FaitObjet(St : String) : TTInfoComp ;
    Function  ChercheItem(St : String ; Var Titre : String) : Byte ;
    Procedure EcritLaListe ;
    Procedure GenereListeClean ;
    Procedure SwapControl ;
  public
    { Déclarations publiques }
  end;

Procedure AttributComplementaires(LesAttribs : HTStrings) ;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  GuidTool ;

Procedure AttributComplementaires(LesAttribs : HTStrings) ;
var FAtChComp : TFAtChComp ;
BEGIN
FAtChComp:=TFAtChComp.Create(Application) ;
  Try
   FAtChComp.LesAttribs:=LesAttribs ;
   FAtChComp.ShowModal ;
  Finally
   FAtChComp.Free ;
  End ;
SourisNormale ;
END ;

procedure TFAtChComp.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFAtChComp.WMGetMinMaxInfo(var MSG: Tmessage);
begin with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; end;

procedure TFAtChComp.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; OnFerme:=False ; Modifier:=False ; AFermer:=True ; end;

procedure TFAtChComp.FormResize(Sender: TObject);
begin
Tv.Width:=(Width div 2)-5 ; PRigth.Left:=Tv.Width+5 ;
if Width Mod 2 =0 then PRigth.Width:=Tv.Width else PRigth.Width:=Tv.Width+1 ;
end;

procedure TFAtChComp.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if Modifier then
   BEGIN
   AFermer:=True ;
   Case HM.Execute(3,'','')of
        mrYes    : EcritLaListe ;
        mrCancel : AFermer:=False ;
    End ;
   END ;
CanClose:=AFermer ;
end;

procedure TFAtChComp.FormClose(Sender: TObject; var Action: TCloseAction);
begin OnFerme:=True ; VideTree(Tv) ; end;

procedure TFAtChComp.FormShow(Sender: TObject);
begin FaitNodes ; GenereListeClean ; end;

Function TFAtChComp.FaitObjet(St : String) : TTInfoComp ;
Var X : TTInfoComp ;
    St1 : String ;
BEGIN
X:=TTInfoComp.Create ; St1:=ReadTokenSt(St) ; X.Origine:=Copy(St1,1,Pos(':',St1)-1) ;
X.Champ:=ReadTokenSt(St) ; X.OblG:=ReadTokenSt(St) ; X.Modif:=ReadTokenSt(St) ;
X.ValDef:=ReadTokenSt(St) ; X.ValTl:=ReadTokenSt(St) ;
X.Acces:=ReadTokenSt(St) ; X.EstCocher:=ReadTokenSt(St) ;
Result:=X ;
END ;

Function TFAtChComp.ChercheItem(St : String ; Var Titre : String) : Byte ;
Var StTemp,St1 : String ;
    j : Integer ;
BEGIN
Result:=0 ;
StTemp:=ReadTokenSt(St) ;
Titre:=Copy(StTemp,Pos(':',StTemp)+1,Length(StTemp)) ; StTemp:=ReadTokenSt(St) ;
if (Pos('LE',StTemp)>0) or (Pos('SC_',StTemp)>0) then BEGIN Result:=8 ; Exit ; END ;
St1:='' ;
for j:=1 to Length(StTemp) do if StTemp[j] in ['0'..'9'] then St1:=St1+StTemp[j] ;
j:=StrToInt(St1) ;
if (j>=1) and (j<=10) then Result:=3 else
   if (j>=11) and (j<=14) then Result:=4 else
      if (j>=15) and (j<=18) then Result:=5 else
         if (j>=19) and (j<=20) then Result:=6 else
            if j=21 then Result:=7 ;
END ;

Procedure TFAtChComp.FaitNodes ;
Var T1,T2,T3,T4,T5 : TTreeNode ;
    i,j : Integer ;
    X : TTInfoComp ;
    St,StTemp,Titre : String ;
BEGIN
T1:=Nil ; T2:=Nil ; T3:=Nil ;
T1:=Tv.Items.Add(T1,HM.Mess[0]) ;
T1.ImageIndex:=0 ; T1.SelectedIndex:=T1.ImageIndex ;
for i:=0 to LesAttribs.Count-1 do
   BEGIN
   St:=LesAttribs.Strings[i] ;
   if Copy(St,Length(St)-1,1)='-' then Continue ;
   X:=FaitObjet(St) ; Titre:='' ;
   if Pos('SC_LIBREENTETE',St)=1 then
      BEGIN
      if T2=Nil then
         BEGIN
         T2:=Tv.Items.AddChild(T1,HM.Mess[1]) ; T2.ImageIndex:=1 ; T2.SelectedIndex:=T2.ImageIndex ;
         END ;
      j:=ChercheItem(St,Titre) ;
      T4:=Tv.Items.AddChildObject(T2,Titre,X) ;
      T4.ImageIndex:=j ; T4.SelectedIndex:=T4.ImageIndex ;
      END else
      if Pos('SC_RADICAL',St)=1 then
         BEGIN
         StTemp:=Copy(St,11,Pos(':',St)-11) ;
         StTemp:='R'+StTemp ;
         if T3=Nil then
            BEGIN
            T3:=Tv.Items.AddChild(T1,HM.Mess[2]) ; T3.ImageIndex:=2 ; T3.SelectedIndex:=T3.ImageIndex ;
            END ;
         j:=ChercheItem(St,Titre) ;
         T5:=Tv.Items.AddChildObject(T3,StTemp+' '+Titre,X) ;
         T5.ImageIndex:=j ; T5.SelectedIndex:=T5.ImageIndex ;
         END ;
   END ;
Tv.Items.GetFirstNode.Expand(False) ; Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
Tv.SetFocus ;
END ;

Procedure TFAtChComp.SwapControl ;
BEGIN
BAssist.Enabled:=False ;
TZoomTl.Enabled:=False ; ZoomTl.Enabled:=False ;
Case Tv.Selected.ImageIndex of
    3 : BEGIN
        TValdef.Visible:=True ; Valdef.Visible:=True ; CbDate.Visible:=False ;
        ZoomTabli.Visible:=False ; CbValDef.Visible:=False ;
        TZoomTl.Enabled:=True ; ZoomTl.Enabled:=True ;
        END ;
    4 : BEGIN
        CbValDef.Visible:=False ; Valdef.Visible:=False ; TZoomTl.Enabled:=False ;
        ZoomTl.Enabled:=False ; ZoomTabli.Visible:=True ; TValdef.Visible:=True ;
        CbDate.Visible:=False ;
        Case TTInfoComp(Tv.Selected.Data).Champ[Length(TTInfoComp(Tv.Selected.Data).Champ)] of
            '1' : ZoomTabli.ZoomTable:=tzNatEcrE0 ;
            '2' : ZoomTabli.ZoomTable:=tzNatEcrE1 ;
            '3' : ZoomTabli.ZoomTable:=tzNatEcrE2 ;
            '4' : ZoomTabli.ZoomTable:=tzNatEcrE3 ;
           End ;
        END ;
    5 : BEGIN
        ZoomTabli.Visible:=False ; CbValDef.Visible:=False ;
        TValdef.Visible:=True ; ValDef.Visible:=True ; CbDate.Visible:=False ;
        END ;
    6 : BEGIN
        ZoomTabli.Visible:=False ; ValDef.Visible:=False ; CbDate.Visible:=False ;
        TValdef.Visible:=False ; CbValDef.Visible:=True ;
        CbValDef.Checked:=TTInfoComp(Tv.Selected.Data).ValDef='X' ;
        END ;
    7 : BEGIN
        CbDate.Visible:=True ; ZoomTabli.Visible:=False ; ValDef.Visible:=False ;
        TValdef.Visible:=True ; CbValDef.Visible:=False ;
        ZoomTabli.Visible:=False ;
        END ;
    8 : BEGIN
        CbValDef.Visible:=False ; TZoomTl.Enabled:=False ; ZoomTl.Enabled:=False ;
        ZoomTabli.Visible:=False ; ValDef.Visible:=True ;
        TValDef.Visible:=True ; CbDate.Visible:=False ;
        END ;
    End ;
CbObligatoire.Enabled:=TTInfoComp(Tv.Selected.Data).Acces='X' ;
CbModifiable.Enabled:=TTInfoComp(Tv.Selected.Data).Acces='X' ;
TValdef.Enabled:=TTInfoComp(Tv.Selected.Data).Acces='X' ;
CbValDef.Enabled:=TTInfoComp(Tv.Selected.Data).Acces='X' ;
ZoomTabli.Enabled:=TTInfoComp(Tv.Selected.Data).Acces='X' ;
ValDef.Enabled:=TTInfoComp(Tv.Selected.Data).Acces='X' ;
CbDate.Enabled:=TTInfoComp(Tv.Selected.Data).Acces='X' ;
if (TTInfoComp(Tv.Selected.Data).Acces='X') and (Tv.Selected.ImageIndex=3) then
   BEGIN TZoomTl.Enabled:=True ; ZoomTl.Enabled:=True ; END else
   BEGIN TZoomTl.Enabled:=False ; ZoomTl.Enabled:=False ; END ;
if (TTInfoComp(Tv.Selected.Data).Acces='X') and ((Tv.Selected.ImageIndex=3) or
   (Tv.Selected.ImageIndex=5) or (Tv.Selected.ImageIndex=8)) then BAssist.Enabled:=True
                                                             else BAssist.Enabled:=False ;
END ;

procedure TFAtChComp.TvChange(Sender: TObject; Node: TTreeNode);
begin
if OnFerme then Exit ;
if Tv.Selected.Text=HM.Mess[0] then
   if Tv.Items.GetFirstNode.GetNext<>Nil then Tv.Selected:=Tv.Items.GetFirstNode.GetNext ;
if Tv.Selected.Data<>Nil then
   BEGIN
   PRigth.Visible:=True ;
   SwapControl ;
   Case Tv.Selected.ImageIndex of
        3,5,8 : Valdef.Text:=TTInfoComp(Tv.Selected.Data).ValDef ;
        4 : ZoomTabli.Text:=TTInfoComp(Tv.Selected.Data).ValDef ;
        6 : CbValDef.Checked:=TTInfoComp(Tv.Selected.Data).ValDef='X' ;
        7 : BEGIN
            if TTInfoComp(Tv.Selected.Data).ValDef='' then CbDate.ItemIndex:=0
                                                      else CbDate.Value:=TTInfoComp(Tv.Selected.Data).ValDef ;
            END ;
      End ;
   if TTInfoComp(Tv.Selected.Data).ValTl='' then ZoomTl.ItemIndex:=0
                                            else ZoomTl.Value:=TTInfoComp(Tv.Selected.Data).ValTl ;
   if TTInfoComp(Tv.Selected.Data).OblG='X' then CbObligatoire.State:=cbChecked
                                            else CbObligatoire.State:=cbUnChecked ;
   if TTInfoComp(Tv.Selected.Data).Modif='X' then CbModifiable.State:=cbChecked
                                             else CbModifiable.State:=cbUnChecked ;
   if TTInfoComp(Tv.Selected.Data).Champ='LE9' then BEGIN TValDef.Enabled:=False ; ValDef.Enabled:=False ; END ;
   END else BEGIN PRigth.Visible:=False ; BAssist.Enabled:=False ; END ;
end;

procedure TFAtChComp.BAssistClick(Sender: TObject);
Var St : String ;
begin
St:=ChoixChampZone(0,'LIB') ;
Valdef.Text:=St ;
end;

procedure TFAtChComp.CbObligatoireClick(Sender: TObject);
begin
if Tv.Selected.ImageIndex=6 then CbObligatoire.Checked:=False ;
if CbObligatoire.Checked then
   BEGIN
   TTInfoComp(Tv.Selected.Data).OblG:='X' ;
   TTInfoComp(Tv.Selected.Data).Modif:='X' ; CbModifiable.State:=cbChecked ;
   END else
   BEGIN
   TTInfoComp(Tv.Selected.Data).OblG:='-' ;
   END ;
Modifier:=True ;
end;

procedure TFAtChComp.CbModifiableClick(Sender: TObject);
begin
if CbModifiable.Checked then
   BEGIN
   TTInfoComp(Tv.Selected.Data).Modif:='X' ;
   END else
   BEGIN
   TTInfoComp(Tv.Selected.Data).Modif:='-' ;
   TTInfoComp(Tv.Selected.Data).OblG:='-' ; CbObligatoire.State:=cbUnchecked ;
   END ;
Modifier:=True ;
end;

procedure TFAtChComp.CbValDefClick(Sender: TObject);
begin
if Not CbValDef.Visible then Exit ;
if CbValDef.Checked then TTInfoComp(Tv.Selected.Data).ValDef:='X'
                    else TTInfoComp(Tv.Selected.Data).ValDef:='-' ;
Modifier:=True ;
end;

procedure TFAtChComp.ValdefChange(Sender: TObject);
begin
TTInfoComp(Tv.Selected.Data).ValDef:=Valdef.Text ;
Modifier:=True ;
end;

procedure TFAtChComp.ZoomTlChange(Sender: TObject);
begin
TTInfoComp(Tv.Selected.Data).ValTl:=ZoomTl.Value ;
Modifier:=True ;
end;

procedure TFAtChComp.CbDateChange(Sender: TObject);
begin
TTInfoComp(Tv.Selected.Data).ValDef:=CbDate.Value ;
Modifier:=True ;
end;

procedure TFAtChComp.ZoomTabliExit(Sender: TObject);
begin
TTInfoComp(Tv.Selected.Data).ValDef:=ZoomTabli.Text ;
Modifier:=True ;
end;

Procedure TFAtChComp.EcritLaListe ;
Var i : Integer ;
    St,St1,StTemp : String ;
    OkNonSais : boolean ;
    T : TTreeNode ;
BEGIN
//LesAttribs.Clear ;
OkNonSais:=False ;
T:=Tv.Items.GetFirstNode ;
InitMove(Tv.Items.Count-1,'') ;
while T<>nil do
   begin
   MoveCur(False) ;
   if T.Data<>nil then
      BEGIN
      St:='' ; StTemp:='' ;
      St:=St+TTInfoComp(T.Data).Origine+';' ;
      St:=St+TTInfoComp(T.Data).Champ+';' ;
      StTemp:=St ;
      St:=St+TTInfoComp(T.Data).OblG+';' ;
      St1:=TTInfoComp(T.Data).Modif ; if St1='-' then OkNonSais:=True ;
      St:=St+St1+';' ;
      St:=St+TTInfoComp(T.Data).ValDef+';' ;
      St:=St+TTInfoComp(T.Data).ValTl+';' ;
      for i:=0 to LesAttribs.Count-1 do
          if Pos(StTemp,LesAttribs.Strings[i])>0 then
             BEGIN LesAttribs.Strings[i]:=St ; Break ; END ;
      END ;
   T:=T.GetNext ;
   end ;
Modifier:=False ; FiniMove ;
if OkNonSais then HM.Execute(4,Caption,'') ;
END ;

procedure TFAtChComp.BValiderClick(Sender: TObject);
begin EcritLaListe ; end;

Procedure TFAtChComp.GenereListeClean ;
Var St,St1,St2 : String ;
    i : Integer ;
BEGIN
for i:=0 to LesAttribs.Count-1 do
  BEGIN
  St:=LesAttribs.Strings[i] ;
  St1:=Copy(St,1,Pos(':',St)-1) ;
  St2:=Copy(St,Pos(';',St)+1,(Length(St)-2)-(Pos(';',St))) ;
  St:=St1+';'+St2 ;
  LesAttribs.Strings[i]:=St ;
  END ;
END ;

end.
