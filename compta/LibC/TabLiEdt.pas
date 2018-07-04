{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 10/02/2005
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
// QTri  Select * From NatCpte Where NT_TYPECPTE=:Nat
unit TabLiEdt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  Hent1, StdCtrls, Buttons, ExtCtrls, Hctrls, ComCtrls, HSysMenu,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  DB, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  hmsgbox  ,UentCommun;

Procedure ChoixTableLibre(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;
Procedure ChoixTableLibreSur(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;

// VL 150305 Migration budget  
// Ajout de ces procédures pour récupérer les info des tables libres automatiquement
Procedure ChoixTableLibreInfo(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;
Procedure ChoixTableLibreSurInfo(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;


type
  TFTabLiEdt = class(TForm)
    PBouton: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    PanelTrie: TPanel;
    CTri: TEdit;
    PanelChoix: TPanel;
    LT0: TLabel;
    LT1: TLabel;
    LT2: TLabel;
    LT3: TLabel;
    LT4: TLabel;
    Label1: TLabel;
    CB0: TCheckBox;
    CB1: TCheckBox;
    CB2: TCheckBox;
    CB3: TCheckBox;
    CB4: TCheckBox;
    CbMemo: THValComboBox;
    PanelCombos1: TPanel;
    LT00: TLabel;
    LT11: TLabel;
    LT22: TLabel;
    LT33: TLabel;
    LT44: TLabel;
    T0: THValComboBox;
    T1: THValComboBox;
    T2: THValComboBox;
    T3: THValComboBox;
    T4: THValComboBox;
    T00: THValComboBox;
    T11: THValComboBox;
    T22: THValComboBox;
    T33: THValComboBox;
    T44: THValComboBox;
    LT5: TLabel;
    LT6: TLabel;
    LT7: TLabel;
    LT8: TLabel;
    LT9: TLabel;
    CB5: TCheckBox;
    CB6: TCheckBox;
    CB7: TCheckBox;
    CB8: TCheckBox;
    CB9: TCheckBox;
    LT55: TLabel;
    LT66: TLabel;
    LT77: TLabel;
    LT88: TLabel;
    LT99: TLabel;
    T5: THValComboBox;
    T6: THValComboBox;
    T7: THValComboBox;
    T8: THValComboBox;
    T9: THValComboBox;
    T55: THValComboBox;
    T66: THValComboBox;
    T77: THValComboBox;
    T88: THValComboBox;
    T99: THValComboBox;
    CSortia: TEdit;
    CEntre: TEdit;
    CSorti: TEdit;
    Centrea: TEdit;
    LLT0: TLabel;
    LLT1: TLabel;
    LLT2: TLabel;
    LLT3: TLabel;
    LLT4: TLabel;
    LLT5: TLabel;
    LLT6: TLabel;
    LLT7: TLabel;
    LLT8: TLabel;
    LLT9: TLabel;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure T0Change(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure T0Click(Sender: TObject);
    procedure CB0MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure CTriDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
  private
    Pref : Char ;
    Cod,Cod1,Tri : String ;
    Ind : Integer ;
    Lefb : TFichierBase ;
    LLibelTable :HTStringList ;
    CodeRupture : Boolean ;
    Procedure ChargeLesCombos ;
    Procedure PositionneLesCombos ;
    Procedure GriseDegriseLesCombos(T,T1 : THValComboBox ; Cb : TCheckBox) ;
    Procedure InterpreteLeTri ;
    Procedure AffecteLibelTable ;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;
    Procedure StartLaForm(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String ; SurRupture : Boolean; OnlyInfo : Boolean = False) ;


implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  TriTabLi ;

Procedure ChoixTableLibre(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;
BEGIN
StartLaForm(Lefb,Tri,Cod,Cod1,False) ;
END ;

Procedure ChoixTableLibreSur(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;
BEGIN
StartLaForm(Lefb,Tri,Cod,Cod1,True) ;
END ;

Procedure ChoixTableLibreInfo(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;
begin
  StartLaForm(Lefb,Tri,Cod,Cod1,False, True);
end;

Procedure ChoixTableLibreSurInfo(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String) ;
begin
  StartLaForm(Lefb,Tri,Cod,Cod1,True, True);
end;

Procedure StartLaForm(Lefb : TFichierBase ; Var Tri,Cod,Cod1 : String ; SurRupture : Boolean; OnlyInfo : Boolean = False) ;
var FTabLiEdt : TFTabLiEdt ;
BEGIN
FTabLiEdt:=TFTabLiEdt.Create(Application) ;
  Try
   FTabLiEdt.Lefb:=Lefb ; FTabLiEdt.CodeRupture:=SurRupture ;
   Case Lefb of
        fbGene         : BEGIN
                         FTabLiEdt.Pref:='G' ; FTabLiEdt.Ind:=1 ;
                         FTabLiEdt.Caption:=FTabLiEdt.MsgBox.Mess[1] ;
                         END ;
        fbAux          : BEGIN
                         FTabLiEdt.Pref:='T' ; FTabLiEdt.Ind:=2 ;
                         FTabLiEdt.Caption:=FTabLiEdt.MsgBox.Mess[2] ;
                         END ;
        fbBudgen       : BEGIN
                         FTabLiEdt.Pref:='B' ; FTabLiEdt.Ind:=4 ;
                         FTabLiEdt.Caption:=FTabLiEdt.MsgBox.Mess[3] ;
                         END ;
        fbAxe1..fbAxe5 : BEGIN
                         FTabLiEdt.Pref:='S' ; FTabLiEdt.Ind:=3 ;
                         FTabLiEdt.Caption:=FTabLiEdt.MsgBox.Mess[4] ;
                         END ;
        fbBudSec1..fbBudSec5 : BEGIN
                               FTabLiEdt.Pref:='D' ; FTabLiEdt.Ind:=5 ;
                               FTabLiEdt.Caption:=FTabLiEdt.MsgBox.Mess[5] ;
                               END ;
        fbImmo         : BEGIN
                         FTabLiEdt.Pref:='I' ; FTabLiEdt.Ind:=9 ;
                         FTabLiEdt.Caption:=FTabLiEdt.MsgBox.Mess[6] ;
                         END ;
     End ;
   FTabLiEdt.Tri:=Tri ;
   if Pos('<>',Cod)>0 then
      BEGIN
      FTabLiEdt.Cod:=Copy(Cod,1,Pos('<>',Cod)-1) ;
      FTabLiEdt.Cod1:=Copy(Cod,Pos('<>',Cod)+2,Length(Cod)) ;
      END else
      BEGIN
      FTabLiEdt.Cod:=Cod ;
      FTabLiEdt.Cod1:=Cod1 ;
      END ;
   if not OnlyInfo then FTabLiEdt.ShowModal
   else begin
     FTabLiEdt.FormShow(nil);
     FTabLiEdt.CB0.Checked := True;
     FTabLiEdt.CB0MouseUp(FTabLiEdt.CB0, mbLeft, [], 0,0);
   end;
   Tri:=FTabLiEdt.Tri ;
   if (OnlyInfo) or (FTabLiEdt.ModalResult=mrOk) then
      BEGIN
      if Pos('<>',Cod)>0 then
         BEGIN Cod:=FTabLiEdt.CSorti.Text+'<>'+FTabLiEdt.CSortia.Text ; END else
         BEGIN Cod:=FTabLiEdt.CSorti.Text ; Cod1:=FTabLiEdt.CSortia.Text ; END ;
      END else
      BEGIN
      if Pos('<>',Cod)>0 then
         BEGIN Cod:=FTabLiEdt.CEntre.Text+'<>'+FTabLiEdt.CEntrea.Text ; END else
         BEGIN Cod:=FTabLiEdt.CEntre.Text ; Cod1:=FTabLiEdt.CEntrea.Text ; END ;
      END ;
  Finally
   FTabLiEdt.Free ;
  End ;
SourisNormale ;
END ;

procedure TFTabLiEdt.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFTabLiEdt.FormShow(Sender: TObject);
Var St : String ;
    i,j : Integer ;
    C,C1 : THValComboBox ;
{$IFDEF CCS3}
    TLB : TLabel ;
{$ENDIF}
    Cb : TCheckBox ;
begin
//ClientHeight:=250 ;
If CodeRupture then ClientWidth:=540 else ClientWidth:=280 ;
//Pages.ActivePage:=Pages.Pages[0] ;
ChargeLesCombos ; InterpreteLeTri ;
AffecteLibelTable ;
if Cod='' then
   BEGIN
   CEntre.Text:='' ; CEntrea.Text:='' ;
   for i:=0 to 9 do
      BEGIN
      St:='' ;
      if THValComboBox(FindComponent('T'+IntToStr(i))).Enabled=False then
         BEGIN
         CSorti.Text:=CSorti.Text+'#;' ; CSortia.Text:=CSortia.Text+'#;'
         END else
         BEGIN
         for j:=1 to VH^.LgTableLibre[Ind,i+1] do St:=St+'-' ;
         St:=St+';' ; CSorti.Text:=CSorti.Text+St ; CSortia.Text:=CSortia.Text+St ;
         END ;
      END ;
   END else
   BEGIN
   CEntre.Text:=Cod ; CEntrea.Text:=Cod1 ; CSorti.Text:=Cod ; CSortia.Text:=Cod1 ;
   END ;
if Cod='' then
   BEGIN
   for i:=0 to 9 do
      BEGIN
      C:=THValComboBox(FindComponent('T'+IntToStr(i))) ;
      C1:=THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))) ;
      Cb:=TCheckBox(FindComponent('CB'+IntToStr(i))) ;
      GriseDegriseLesCombos(C,C1,Cb) ;
{$IFDEF CCS3}
      If I>3 Then
        BEGIN
        C.Visible:=FALSE ; C1.Visible:=FALSE ; CB.Visible:=FALSE ;
        TLB:=TLabel(FindComponent('LLT'+IntToStr(i))) ;  TLB.Visible:=FALSE ;
        TLB:=TLabel(FindComponent('LT'+IntToStr(i)+IntToStr(i))) ;  TLB.Visible:=FALSE ;
        TLB:=TLabel(FindComponent('LT'+IntToStr(i))) ;  TLB.Visible:=FALSE ;
        END ;
{$ENDIF}
      END ;
   END else PositionneLesCombos ;
PanelCombos1.Visible:=CodeRupture ;
PanelTrie.Visible:=false ;
{$IFDEF CCS3}
Self.Height:=Self.Height-(CB9.Top-CB3.Top) ;
{$ENDIF}
end;

Procedure TFTabLiEdt.ChargeLesCombos ;
Var Sql : String ;
    i : Integer ;
    QLoc : TQuery;
BEGIN
for i:=0 to 9 do BEGIN
  QLoc := OpenSQL('SELECT NT_NATURE,NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE="'+Pref+'0'+IntToStr(i)+'" ORDER BY NT_NATURE', True);
  if QLoc.Eof then
     BEGIN
     TLabel(FindComponent('LT'+IntToStr(i))).Enabled:=False ;
     TLabel(FindComponent('LT'+IntToStr(i)+IntToStr(i))).Enabled:=False ;
     THValComboBox(FindComponent('T'+IntToStr(i))).Enabled:=False ;
     THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Enabled:=False ;
     TCheckBox(FindComponent('CB'+IntToStr(i))).Enabled:=False ;
     END else
     BEGIN
     THValComboBox(FindComponent('T'+IntToStr(i))).Values.Clear ;
     THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Values.Clear ;
     THValComboBox(FindComponent('T'+IntToStr(i))).Items.Clear ;
     THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Items.Clear ;
     THValComboBox(FindComponent('T'+IntToStr(i))).Values.Add('') ;
     THValComboBox(FindComponent('T'+IntToStr(i))).Items.Add(Traduirememoire('<<Tous>>')) ;
     While Not QLoc.Eof do
         BEGIN
         THValComboBox(FindComponent('T'+IntToStr(i))).Values.Add(QLoc.Fields[0].AsString) ;
         THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Values.Add(QLoc.Fields[0].AsString) ;
         THValComboBox(FindComponent('T'+IntToStr(i))).Items.Add(QLoc.Fields[1].AsString) ;
         THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Items.Add(QLoc.Fields[1].AsString) ;
         QLoc.Next ;
         END ;
     END ;
  END ;
  Ferme(QLoc);
Sql:='Select CC_LIBELLE,CC_CODE From CHOIXCOD Where CC_TYPE="NAT" And CC_CODE Like"'+Pref+'0%"' ;
QLoc := OpenSQL(Sql, True);
i:=0 ;
CbMemo.Values.Clear ; CbMemo.Items.Clear ;
While Not QLoc.Eof do
   BEGIN
   TLabel(FindComponent('LT'+IntToStr(i))).Caption:=QLoc.Fields[0].AsString ;
   CbMemo.Items.Add(QLoc.Fields[0].AsString) ; CbMemo.Values.Add(QLoc.Fields[1].AsString) ;
   Inc(i) ; QLoc.Next ;
   END ;
Ferme(QLoc);
END ;

Procedure TFTabLiEdt.PositionneLesCombos ;
Var Stde,Sta,St1,St2 : String ;
    i : Integer ;
    C,C1 : THValComboBox ;
BEGIN
if Cod='' then Exit ; Stde:=Cod ; Sta:=Cod1 ;
for i:=0 to 9 do
    BEGIN
    C:=THValComboBox(FindComponent('T'+IntToStr(i))) ;
    C1:=THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))) ;
    St1:=ReadTokenSt(Stde) ; St2:=ReadTokenSt(Sta) ;
    if (St1[1]='#') Or (St1[1]='-') then
       BEGIN
       GriseDegriseLesCombos(C,C1,TCheckBox(FindComponent('CB'+IntToStr(i)))) ;
       Continue ;
       END ;
    if St1[1]='*' then
       BEGIN
       TLabel(FindComponent('LT'+IntToStr(i)+IntToStr(i))).Enabled:=False ;
       C1.Enabled:=False ; C.ItemIndex:=0 ;
       END else
       BEGIN
       C.Value:=C.Values[C.Values.IndexOf(Copy(St1,1,VH^.LgTableLibre[Ind,i+1]))] ;
       C1.Value:=C1.Values[C1.Values.IndexOf(Copy(St2,1,VH^.LgTableLibre[Ind,i+1]))] ;
       END ;
    TCheckBox(FindComponent('CB'+IntToStr(i))).State:=cbChecked ;
    END ;
END ;

procedure TFTabLiEdt.T0Change(Sender: TObject);
Var Ch,St,St1 : String ;
    i,j,k : Integer ;
begin
Ch:=Copy(THValComboBox(Sender).Name,Length(THValComboBox(Sender).Name),1) ;
if TCheckBox(FindComponent('CB'+Ch)).Enabled=False then Exit ;
j:=0 ;
if Length(THValComboBox(Sender).Name)=2 then St:=CSorti.Text
                                        else St:=CSortia.Text ;
k:=1 ;
for i:=1 to Length(St) do
    BEGIN
    if j<StrToInt(Ch) then
       BEGIN
       if St[i]=';' then Inc(j) ;
       END else BEGIN k:=i ; Break ; END ;
    END ;
St1:=THValComboBox(Sender).Value ;
if (St1='') And (THValComboBox(FindComponent('T'+Ch)).Text<>Traduirememoire('<<Tous>>'))then
   for i:=1 to VH^.LgTableLibre[Ind,StrToInt(Ch)+1] do St1:=St1+'-' ;
if THValComboBox(FindComponent('T'+Ch)).Text=Traduirememoire('<<Tous>>') then
   for i:=1 to VH^.LgTableLibre[Ind,StrToInt(Ch)+1] do St1:=St1+'*' ;
if St[k]='#' then Delete(St,k,1)
             else Delete(St,k,VH^.LgTableLibre[Ind,StrToInt(Ch)+1]) ;
Insert(St1,St,k) ;
if Length(THValComboBox(Sender).Name)=2 then CSorti.Text:=St
                                        else CSortia.Text:=St ;
if Length(THValComboBox(Sender).Name)=2 then
   BEGIN
   if THValComboBox(Sender).ItemIndex > THValComboBox(FindComponent('T'+Ch+Ch)).ItemIndex+1 then
      THValComboBox(FindComponent('T'+Ch+Ch)).Value:=THValComboBox(Sender).Value ;
   END else
   BEGIN
   if THValComboBox(Sender).ItemIndex+1 < THValComboBox(FindComponent('T'+Ch)).ItemIndex then
      THValComboBox(FindComponent('T'+Ch)).Value:=THValComboBox(Sender).Value ;
   END ;
end;

procedure TFTabLiEdt.BValiderClick(Sender: TObject);
begin ModalResult:=mrOk ; end;

procedure TFTabLiEdt.T0Click(Sender: TObject);
Var Ch : String ;
begin
Ch:=Copy(THValComboBox(Sender).Name,Length(THValComboBox(Sender).Name),1) ;
if Length(THValComboBox(Sender).Name)=2 then
   BEGIN
   if THValComboBox(FindComponent('T'+Ch)).Value<>'' then
      BEGIN
      if Not(THValComboBox(FindComponent('T'+Ch+Ch)).Enabled) then
         BEGIN
         THValComboBox(FindComponent('T'+Ch+Ch)).Enabled:=True ;
         TLabel(FindComponent('LT'+Ch+Ch)).Enabled:=True ;
         END ;
      if THValComboBox(FindComponent('T'+Ch+Ch)).Value='' then
         THValComboBox(FindComponent('T'+Ch+Ch)).Value:=THValComboBox(FindComponent('T'+Ch)).Value ;
      END else
      BEGIN
      THValComboBox(FindComponent('T'+Ch+Ch)).Value:='' ;
      THValComboBox(FindComponent('T'+Ch+Ch)).Enabled:=False ;
      TLabel(FindComponent('LT'+Ch+Ch)).Enabled:=False ;
      END
   END else
   BEGIN
   if THValComboBox(FindComponent('T'+Ch+Ch)).Value<>'' then
      BEGIN
      if THValComboBox(FindComponent('T'+Ch)).Value='' then
         THValComboBox(FindComponent('T'+Ch)).Value:=THValComboBox(FindComponent('T'+Ch+Ch)).Value ;
      END else
      BEGIN
      THValComboBox(FindComponent('T'+Ch)).Value:='' ;
      THValComboBox(FindComponent('T'+Ch)).Text:=Traduirememoire('<<Tous>>') ;
      THValComboBox(FindComponent('T'+Ch+Ch)).Enabled:=False ;
      TLabel(FindComponent('LT'+Ch+Ch)).Enabled:=False ;
      END ;
   END ;
end;

procedure TFTabLiEdt.CB0MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
Var C,C1 : THValComboBox ;
begin
C:=THValComboBox(FindComponent('T'+Copy(TCheckBox(Sender).Name,3,1))) ;
C1:=THValComboBox(FindComponent('T'+Copy(TCheckBox(Sender).Name,3,1)+Copy(TCheckBox(Sender).Name,3,1))) ;
if TCheckBox(Sender).State=cbUnchecked then
   BEGIN
   C.Value:='' ; C1.Value:='' ;
   END else
   BEGIN
   C.ItemIndex:=0 ; T0Change(C) ; T0Change(C1) ; T0Click(C) ;
   END ;
GriseDegriseLesCombos(C,C1,TCheckBox(Sender)) ;
end;

Procedure TFTabLiEdt.GriseDegriseLesCombos(T,T1 : THValComboBox ; Cb : TCheckBox) ;
BEGIN
If Not CodeRupture then Exit ;
if Cb.State=cbUnChecked then
   BEGIN
   T.Enabled:=False ; T1.Enabled:=False ;
   TLabel(FindComponent('LT'+Copy(Cb.Name,3,1))).Enabled:=False ;
   TLabel(FindComponent('LT'+Copy(Cb.Name,3,1)+Copy(Cb.Name,3,1))).Enabled:=False ;
   END else
   BEGIN
   T.Enabled:=True ;
   TLabel(FindComponent('LT'+Copy(Cb.Name,3,1))).Enabled:=True ;
   if T.ItemIndex=0 then
      BEGIN
      T1.Enabled:=False ;
      TLabel(FindComponent('LT'+Copy(Cb.Name,3,1)+Copy(Cb.Name,3,1))).Enabled:=False ;
      END else
      BEGIN
      T1.Enabled:=True ;
      TLabel(FindComponent('LT'+Copy(Cb.Name,3,1)+Copy(Cb.Name,3,1))).Enabled:=True
      END ;
   END ;
END ;

Procedure TFTabLiEdt.InterpreteLeTri ;
Var St : String ;
BEGIN
St:=Tri ; CTri.Text:='' ;
While St<>'' do
   CTri.Text:=CTri.Text+CbMemo.Items[CbMemo.Values.IndexOf(ReadTokenSt(St))]+';' ;
END ;

procedure TFTabLiEdt.CTriDblClick(Sender: TObject);
begin ChoixTriTableLibre(LeFb,Tri,True,'') ; InterpreteLeTri ; BValider.SetFocus ; end ;

Procedure TFTabLiEdt.AffecteLibelTable ;
Var StLib : String ;
    i     : Byte ;
BEGIN
GetLibelleTableLibre(Pref,LLibelTable) ;
if LLibelTable.Count=0 then exit;
For i:=0 to LLibelTable.Count-1 do
    BEGIN
    StLib:=LLibelTable[i] ;
    TLabel(FindComponent('LT'+IntToStr(i))).Caption:=ReadTokenSt(StLib) ;
    END ;
END ;

procedure TFTabLiEdt.FormCreate(Sender: TObject);
begin
LLibelTable:=HTStringList.Create ;
end;

procedure TFTabLiEdt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
LLibelTable.Free ;
end;

procedure TFTabLiEdt.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
