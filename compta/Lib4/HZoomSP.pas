unit HZoomSP ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Hctrls, ExtCtrls, hmsgbox, HSysMenu,
{$IFDEF EAGLCLIENT}
  uTOB,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,
{$ENDIF}
  Hent1, ComCtrls,Ent1, Hcompte
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
  ,uEntCommun, TntStdCtrls, TntButtons
   ;

Type TValSousPlan = Array[1..MaxSousPlan] Of String ;
Type Tab6Integer = Array[1..6] Of Integer ;
Type TSPB = Record
            Entier,Deb,Long : Tab6Integer ;
            End ;

Function ChoisirValCat(CodeCat,StSPJ : String ; Var Jal,Sect : String ; AvecJoker : Boolean ; Action : TActionFiche) : Boolean ;
Function ChoisirSousPlan(Axe : TFichierBase ; Var Sect : String ; AvecJoker : Boolean ; Action : TActionFiche) : Boolean ;
Function MiseEnFormeValCat(CodeCat,Jal,Sect,StSPJ : String) : TValSousPlan ;
Function MiseEnFormeSec(fb : TFichierBase ; Sect : String) : TValSousPlan ;
Procedure RestitueJalSect(CodeCat : String ; ValSousPlan : TValSousPlan ; Var Jal,Sect : String) ;
Function  RestitueSec(fb : TFichierBase ; ValSousPlan : TValSousPlan ; AvecJoker : Boolean) : String ;
Function CatExiste : Boolean ;

type
  TFZoomSP = class(TForm)
    PBouton: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    Panel2: TPanel;
    TSC1: TLabel;
    SC1: THCpteEdit;
    LibSC1: THLabel;
    SC2: THCpteEdit;
    TSC2: TLabel;
    LibSC2: THLabel;
    TSC3: TLabel;
    SC3: THCpteEdit;
    LibSC3: THLabel;
    TSC4: TLabel;
    SC4: THCpteEdit;
    LibSC4: THLabel;
    TSC5: TLabel;
    SC5: THCpteEdit;
    LibSC5: THLabel;
    TSC6: TLabel;
    SC6: THCpteEdit;
    LibSC6: THLabel;
    Panel3: TPanel;
    FCat: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BFermeClick(Sender: TObject);
    procedure FCatChange(Sender: TObject);
  private
    Action : TActionFiche ;
    AvecJoker : Boolean ;
    CodeCat,Jal,Sect,StSPJ : String ;
    OkSection : Boolean ;
    ValSousPlan : TValSousPlan ;
    CatBud : TUneCatBud ;
    OnBud  : Boolean ;
    IndSP : TSPB ;
    NbVisible : Integer ;
    Function  QuelIndiceBud : TSPB ;
    Function  QuelIndiceSec : TSPB ;
    Procedure InitLesControles(i : Integer) ;
    public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function ChoisirValCat(CodeCat,StSPJ : String ; Var Jal,Sect : String ; AvecJoker : Boolean ; Action : TActionFiche) : Boolean ;
var
  FZoomSP: TFZoomSP ;
BEGIN
Result:=FALSE ;
FZoomSP:=TFZoomSP.Create(Application) ;
try
  FZoomSP.CodeCat:=CodeCat ;
  FZoomSP.Jal:=Jal ;
  FZoomSP.Sect:=Sect ;
  FZoomSP.StSPJ:=StSPJ ;
  FZoomSP.OkSection:=FALSE ;
  FZoomSP.AvecJoker:=AvecJoker ;
  FZoomSP.Action:=Action ;
  FZoomSP.OnBud:=TRUE ;
  FZoomSP.ShowModal ;
  finally
  If Action in [taModif,taCreat] Then
     BEGIN
     if FZoomSP.OkSection then
        BEGIN
        RestitueJalSect(FZoomSP.CodeCat,FZoomSP.ValSousPlan,Jal,Sect) ;
        Result:=TRUE ;
        END ;
     END ;
  FZoomSP.free ;
  end ;
//SourisNormale ;
Screen.Cursor:=SyncrDefault ;
END ;

Function ChoisirSousPlan(Axe : TFichierBase ; Var Sect : String ; AvecJoker : Boolean ; Action : TActionFiche) : Boolean ;
var
  FZoomSP: TFZoomSP ;
BEGIN
Result:=FALSE ;
FZoomSP:=TFZoomSP.Create(Application) ;
try
  FZoomSP.CodeCat:='' ;
  FZoomSP.Jal:='' ;
  FZoomSP.Sect:=Sect ;
  FZoomSP.StSPJ:='' ;
  FZoomSP.OkSection:=FALSE ;
  FZoomSP.AvecJoker:=AvecJoker ;
  FZoomSP.Action:=Action ;
  FZoomSP.OnBud:=FALSE ;
  FZoomSP.CatBud.Fb:=Axe ;
  FZoomSP.ShowModal ;
  finally
  If Action in [taModif,taCreat] Then
     BEGIN
     if FZoomSP.OkSection then
        BEGIN
        Sect:=RestitueSec(FZoomSP.CatBud.fb,FZoomSP.ValSousPlan,FZoomSP.AvecJoker) ;
        Result:=TRUE ;
        END ;
     END ;
  FZoomSP.free ;
  end ;
//SourisNormale ;
Screen.Cursor:=SyncrDefault ;
END ;


procedure TFZoomSP.FormShow(Sender: TObject);
//Var i : Integer ;
begin
If OnBud Then
   BEGIN
   Panel3.Visible:=TRUE ;
   Caption:=MsgBox.Mess[1] ;
   If CodeCat='' Then BEGIN FCat.ItemIndex:=0 ; FCatChange(Nil) ; END Else FCat.Value:=CodeCat ;
   END Else
   BEGIN
   Panel3.Visible:=FALSE ;
   Caption:=MsgBox.Mess[2] ;
   FCat.Visible:=FALSE ; FCatChange(Nil) ;
   END ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
SC1.SetFocus ;
UpdateCaption(Self) ;
end;


procedure TFZoomSP.BValiderClick(Sender: TObject);
Var i : Integer ;
    HCpt : TComponent ;
    OkFerme : Boolean ;
begin
OkFerme:=TRUE ; If Action=taConsult Then Exit ; OkSection:=TRUE ;
For i:=1 To NbVisible Do
   BEGIN
   HCpt:=Self.FindComponent('SC'+IntToStr(i)) ;
   If HCpt<>NIL Then ValSousPlan[i]:=THCpteEdit(HCpt).Text ;
   END ;
If OkFerme Then Close ;
end;

procedure TFZoomSP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ;
if Vide then
   BEGIN
   Case Key of
        VK_F10 : BEGIN Key:=0 ; BValiderClick(Nil) ; END ;
        VK_RETURN :BEGIN
                   Key:=0 ;
                   FindNextControl(ActiveControl,True,True,False).SetFocus ;
                   END ;
     END ;
   END ;
end;

procedure TFZoomSP.BFermeClick(Sender: TObject);
begin
OkSection:=FALSE ;
end;

Function TFZoomSP.QuelIndiceBud : TSPB ;
Var i,j,l,Ind,k : Integer ;
    St : String ;
    Tab6I : TSPB ;
    SurJal : Boolean ;
BEGIN
Fillchar(Tab6I,SizeOf(Tab6I),#0) ; Ind:=0 ;
If CatBud.Code<>'' Then
   BEGIN
   For k:=1 To 2 Do
     BEGIN
     If k=1 Then SurJal:=FALSE Else SurJal:=TRUE ;
     For l:=1 To MaxSousPlan Do
       BEGIN
       If SurJal Then St:=CatBud.SurJal[l] Else St:=CatBud.SurSect[l] ;
       { Prend le l° code sous plan de la catégorie trouvée }
       If St<>'' Then
          BEGIN
          j:=0 ;
          For i:=1 To MaxSousPlan Do If VH^.SousPlanAxe[CatBud.fb,i].Code=St Then j:=i ;
          { retrouve les caractéristiques du sous plan concerné }
          If j<>0 Then
             BEGIN
             Inc(Ind) ; Tab6I.Entier[Ind]:=j ;
             Tab6I.Long[Ind]:=VH^.SousPlanAxe[CatBud.fb,j].Longueur ;
             END ;
          END ;
       END ;
     END ;
   END ;
Result:=Tab6I ;
END ;

Function TFZoomSP.QuelIndiceSec : TSPB ;
Var i,Ind : Integer ;
    Tab6I : TSPB ;
BEGIN
Fillchar(Tab6I,SizeOf(Tab6I),#0) ; Ind:=0 ;
For i:=1 To MaxSousPlan Do If VH^.SousPlanAxe[CatBud.fb,i].Code<>'' Then
  BEGIN
  Inc(Ind) ; Tab6I.Entier[Ind]:=i ;
  Tab6I.Long[Ind]:=VH^.SousPlanAxe[CatBud.fb,i].Longueur ;
  END  ;
Result:=Tab6I ;
END ;

Procedure TFZoomSP.InitLesControles(i : Integer) ;
Var SC : THCpteEdit ;
    TSC : TLabel ;
    SL : THLabel ;
    TC : TComponent ;
BEGIN
TC:=Self.FindComponent('SC'+IntToStr(i))    ; If TC=NIL Then SC:=NIL Else SC:=THCpteEdit(TC) ;
TC:=Self.FindComponent('LibSC'+IntToStr(i)) ; If TC=NIL Then SL:=NIL Else SL:=THLabel(TC) ;
TC:=Self.FindComponent('TSC'+IntToStr(i))   ; If TC=NIL Then TSC:=NIL Else TSC:=TLabel(TC) ;
If SL<>NIL Then SL.Caption:='' ;
If IndSP.Entier[i]<>0 Then
   BEGIN
   If (TSC<>NIL) Then BEGIN TSC.Caption:='&'+VH^.SousPlanAxe[CatBud.fb,IndSP.Entier[i]].Lib ; TSC.Visible:=TRUE ; END ;
   If (SC<>NIL) Then BEGIN SC.ZoomTable:=TZoomTable(Ord(tzAxe1SP1)+IndSP.Entier[i]-1+((Ord(CatBud.fb)-Ord(fbAxe1))*6)) ; SC.Visible:=TRUE ; END ;
   If (SL<>NIL) Then SL.Visible:=TRUE ;
   Inc(NbVisible) ;
   END Else
   BEGIN
   If TSC<>NIL Then TSC.Caption:='' ;
   If (SL<>NIL) Then SL.Visible:=FALSE ; If (TSC<>NIL) Then TSC.Visible:=FALSE ; If (SC<>NIL) Then SC.Visible:=FALSE ;
   END ;
END ;

procedure TFZoomSP.FCatChange(Sender: TObject);
Var i : Integer ;
    HCpt : TComponent ;
begin
If OnBud Then
   BEGIN
   CatBud:=QuelleCatBud(FCat.Value) ; CodeCat:=FCat.Value ;
   ValSousPlan:=MiseEnFormeValCat(CodeCat,Jal,Sect,StSPJ) ;
   IndSP:=QuelIndiceBud ;
   END Else
   BEGIN
   IndSP:=QuelIndiceSec ;
   ValSousPlan:=MiseEnFormeSec(CatBud.fb,Sect) ;
   END ;
NbVisible:=0 ;
For i:=1 To 6 Do
   InitLesControles(i) ;
If OnBud Then Self.Height:=PBouton.Height+Panel3.Height+(NbVisible+1)*(SC1.Height+SC2.Top-(SC1.Top+SC1.Height))+12
         Else Self.Height:=PBouton.Height+(NbVisible+1)*(SC1.Height+SC2.Top-(SC1.Top+SC1.Height))+12 ;
For i:=1 To NbVisible Do If IndSP.Entier[i]<>0 Then
  BEGIN
  HCpt:=Self.FindComponent('SC'+IntToStr(i)) ;
  If HCpt<>NIL Then
     BEGIN
     THCpteEdit(HCpt).Text:=ValSousPlan[i] ;
     THCpteEdit(HCpt).ExisteH ;
     If Action=TaConsult Then THCpteEdit(HCpt).Enabled:=FALSE ;
     END ;
  END ;

end;

Function MiseEnFormeValCat(CodeCat,Jal,Sect,StSPJ : String) : TValSousPlan ;
Var VSP : TValSousPlan ;
    CatBud : TUneCatBud ;
    Q : TQuery ;
    SousPlanCatS : TSousPlanCat ;
    Ind,l : Integer ;
    St,St1 : String ;
BEGIN
Fillchar(VSP,SizeOf(VSP),#0) ;
Result:=VSP ;
If (CodeCat='') And (Jal='') Then Exit ;
If ((CodeCat='') Or (StSPJ='')) And (Jal<>'') Then
   BEGIN
   Q:=OpenSQL('SELECT BJ_CATEGORIE,BJ_SOUSPLAN FROM BUDJAL WHERE BJ_BUDJAL="'+Jal+'"',TRUE,-1, '', True) ;
   If Not Q.Eof Then BEGIN CodeCat:=Q.FindField('BJ_CATEGORIE').AsString ; StSPJ:=Q.FindField('BJ_SOUSPLAN').AsString ; END ;
   Ferme(Q) ;
   END ;
If CodeCat='' Then Exit ;
CatBud:=QuelleCatBud(CodeCat) ;
If CatBud.Code<>'' Then
   BEGIN
{ Alimentation des valeurs de sous plan de la section budgétaire }
   Ind:=0 ;
   SousPlanCatS:=SousPlanCat(CatBud.code,FALSE) ;
   For l:=1 To MaxSousPlan Do If SousPlanCatS[l].Code<>'' Then
      BEGIN
      Inc(Ind) ; VSP[Ind]:=Copy(Sect,SousPlanCatS[l].Debut,SousPlanCatS[l].Longueur) ;
      END ;
{ Alimentation des valeurs de sous plan du journal budgétaire }
   St:=StSPJ ;
   While St<>'' Do
     BEGIN
     St1:=ReadTokenSt(St) ;
     If St1<>'' Then BEGIN Inc(Ind) ; VSP[Ind]:=St1 ; END ;
     END ;
   END ;
Result:=VSP ;
END ;

Function MiseEnFormeSec(fb : TFichierBase ; Sect : String) : TValSousPlan ;
Var VSP : TValSousPlan ;
    Ind,l : Integer ;
BEGIN
Fillchar(VSP,SizeOf(VSP),#0) ; Result:=VSP ; Ind:=0 ; If Sect='' Then Exit ;
For l:=1 To MaxSousPlan Do If VH^.SousPlanAxe[fb,l].Code<>'' Then
   BEGIN
   Inc(Ind) ; VSP[Ind]:=Copy(Sect,VH^.SousPlanAxe[fb,l].Debut,VH^.SousPlanAxe[fb,l].Longueur) ;
   If Pos('?',VSP[Ind])<>0 Then VSP[Ind]:='' ;
   END ;
Result:=VSP ;
END ;

Function  RestitueSec(fb : TFichierBase ; ValSousPlan : TValSousPlan ; AvecJoker : Boolean) : String ;
Var Sect : String ;
    C : Char ;
    j,l : Integer ;
BEGIN
Result:='' ; Sect:='' ;
For l:=1 To MaxSousPlan Do If VH^.SousPlanAxe[fb,l].Code<>'' Then
   BEGIN
   If Trim(ValSousPlan[l])='' Then
      BEGIN
      ValSousPlan[l]:='' ; If AvecJoker Then C:='?' Else C:=' ' ;
      For j:=1 To VH^.SousPlanAxe[fb,l].Longueur Do ValSousPlan[l]:=ValSousPlan[l]+C ;
      END ;
   Sect:=Sect+ValSousPlan[l] ;
   END ;
Result:=Sect ;
END ;

Procedure RestitueJalSect(CodeCat : String ; ValSousPlan : TValSousPlan ; Var Jal,Sect : String) ;
Var CatBud : TUneCatBud ;
    Q : TQuery ;
    SousPlanCatS,SousPlanCatJ : TSousPlanCat ;
    Ind,l,i,Lg : Integer ;
    St,Pt,Code : String ;
    OkOk : Boolean ;
BEGIN
Jal:='' ; Sect:='' ; If (CodeCat='') Then Exit ; Pt:='............' ;
CatBud:=QuelleCatBud(CodeCat) ;
If CatBud.Code<>'' Then
   BEGIN
   //Ind:=0 ;
   SousPlanCatS:=SousPlanCat(CatBud.code,FALSE) ;
   SousPlanCatJ:=SousPlanCat(CatBud.code,TRUE) ;
   Ind:=0 ;
   For l:=1 To MaxSousPlan Do
     BEGIN
     Code:=VH^.SousPlanAxe[CatBud.Fb,l].Code ;
     OkOk:=FALSE ;
     If Code<>'' Then
       BEGIN
       Lg:=1 ;
       For i:=1 To MaxSousPlan Do If SousPlanCatS[i].Code=Code Then
         BEGIN
         Inc(Ind) ; OkOk:=TRUE ;
         END Else If SousPlanCatJ[i].Code=Code Then Lg:=SousPlanCatJ[i].Longueur ;

       If OkOk Then St:=ValSousPlan[Ind] Else St:=Copy(Pt,1,Lg) ;
       Sect:=Sect+St ;
       END ;
     END ;
   While Sect[Length(Sect)]='.' Do Delete(Sect,Length(Sect),1) ;
   St:='' ;
   For i:=Ind+1 To 6 Do If ValSousPlan[i]<>'' Then
     BEGIN
     St:=St+ValSousPlan[i]+';' ;
     END ;
   If St<>'' Then
     BEGIN
     Q:=OpenSQL('SELECT BJ_BUDJAL FROM BUDJAL WHERE BJ_SOUSPLAN="'+St+'"',TRUE,-1, '', True) ;
     If Not Q.Eof Then Jal:=Q.FindField('BJ_BUDJAL').AsString ;
     Ferme(Q) ;
     END ;
   END ;
END ;

Function CatExiste : Boolean ;
Var i : Integer ;
BEGIN
Result:=FALSE ;
For i:=1 To MaxCatBud Do If VH^.CatBud[i].Code<>'' Then Result:=TRUE ;
END ;

end.
