{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit TriTabLi;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1, HEnt1,
  hmsgbox,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  StdCtrls, Buttons, ExtCtrls, HSysMenu, HCtrls   ,UentCommun;

Procedure ChoixTriTableLibre(LeFb : TFichierBase ; Var Tri : String ; Asc : Boolean ; ChoixLibre : String) ;

type
  TFChoiTri = class(TForm)
    PBouton: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    Ptop: TPanel;
    SBox: TScrollBox;
    TCLib: TLabel;
    CLib: TEdit;
    CTri: TEdit;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    MemoTri: THValComboBox;
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Tri : String ;
    LeFb : TFichierBase ;
    Combien : Integer ;
    Pref : String ;
    Asc : Boolean ;
    StChoixLibre : String ;
    Procedure CreationDesCombos ;
    Procedure ComboChanger(Sender : TObject) ;
    Function  QuelCombo(Sender : TObject) : Byte ;
    Procedure InitCombo ;
    Function  CherchePosition(Quel : Byte ; St : String) : Integer ;
    Function  FabriqueChaine(St,St1 : String ; j : Byte) : String ;
    Function  EnregOk : Boolean ;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


implementation

uses CPTEUTIL;

{$R *.DFM}

Procedure ChoixTriTableLibre(LeFb : TFichierBase ; Var Tri : String ; Asc : Boolean ; ChoixLibre : String) ;
var FChoiTri : TFChoiTri ;
BEGIN
FChoiTri:=TFChoiTri.Create(Application) ;
  Try
   FChoiTri.Tri:=Tri ;
   FChoiTri.LeFb:=LeFb ;
   FChoiTri.Asc:=Asc ;
   FChoiTri.StChoixLibre:=ChoixLibre ;
   Case LeFb of
      fbGene :         BEGIN FChoiTri.Caption:=FChoiTri.MsgBox.Mess[0] ; FChoiTri.Pref:='G0' ; END ;
      fbAux :          BEGIN FChoiTri.Caption:=FChoiTri.MsgBox.Mess[1] ; FChoiTri.Pref:='T0' ; END ;
      fbBudgen :       BEGIN FChoiTri.Caption:=FChoiTri.MsgBox.Mess[2] ; FChoiTri.Pref:='B0' ; END ;
      fbAxe1..fbAxe5 : BEGIN FChoiTri.Caption:=FChoiTri.MsgBox.Mess[3] ; FChoiTri.Pref:='S0' ; END ;
      fbBudSec1..fbBudSec5 : BEGIN FChoiTri.Caption:=FChoiTri.MsgBox.Mess[8] ; FChoiTri.Pref:='D0' ; END ;
      fbImmo :         BEGIN FChoiTri.Caption:=FChoiTri.MsgBox.Mess[9] ; FChoiTri.Pref:='I0' ; END ;
    End ;
   if FChoiTri.ShowModal=mrOk then Tri:=FChoiTri.CTri.Text ;
  Finally
   FChoiTri.Free ;
  End ;
SourisNormale ;
END ;

procedure TFChoiTri.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFChoiTri.BValiderClick(Sender: TObject);
begin if EnregOk then ModalResult:=mrOk ; end;

procedure TFChoiTri.FormShow(Sender: TObject);
Var St : String ;
begin
Combien:=0 ; St:=Tri ;
CreationDesCombos ; InitCombo ;
if Combien>0 then THValComboBox(FindComponent('Cb1')).SetFocus ;
end;

Procedure QuelleTable(St : String ; Var NatOk : TOKNatLibre) ;
Const C1='#################' ;
      C2='-----------------' ;
Var i,l : Integer ;
    StV : String ;

BEGIN
For i:=1 TO 10 Do NatOk[i]:=TRUE ; If St='' Then Exit ;
If St[Length(St)]<>';' Then St:=St+';' ; i:=0 ;
While St<>'' Do
  BEGIN
  StV:=ReadTokenSt(St) ; Inc(i) ; l:=Length(StV) ;
  If (StV=Copy(C1,1,l)) Or (StV=Copy(C2,1,l)) Then NatOk[i]:=FALSE ;
  END ;
END ;

Procedure TFChoiTri.CreationDesCombos ;
Const CLar = 250 ;
      CBLar = 90 ;
      CGau = 50 ;
      LGau = 4 ;
      Haut = 28 ;
Var i : Byte ;
    C : THValComboBox ;
    L : TLabel ;
    B : TCheckBox ;
    Sql : String ;
    TabVid : Array[0..9] of Boolean ;
    NatOk :TOKNatLibre ;
    Existe : Boolean;
    QTri : TQuery;
BEGIN
QuelleTable(StChoixLibre,NatOk) ;
for i:=0 to 9 do
   BEGIN
   Existe := ExisteSQL('SELECT * FROM NATCPTE WHERE NT_TYPECPTE="'+Pref+IntToStr(i)+'"');
   if Existe And NatOk[i+1] then BEGIN TabVid[i]:=True ; Inc(Combien) ; END
                            else TabVid[i]:=False ;
   END ;
Sql:='Select CC_CODE,CC_LIBELLE From CHOIXCOD Where CC_TYPE="NAT" ' ;
for i:=0 to 9 do
   BEGIN
   if TabVid[i] then
     if Pos('AND',Sql)>0 then Sql:=Sql+'OR CC_CODE="'+Pref+IntToStr(i)+'" '
                         else Sql:=Sql+'AND (CC_CODE="'+Pref+IntToStr(i)+'" ' ;
   END ;
   if Pos('AND',Sql)>0 then Sql:=SQl+')';
QTri := OpenSQL(Sql, True);
for i:=1 to Combien do
    BEGIN
    C:=THValComboBox.Create(Self); C.Parent:=SBox; C.Name:='Cb'+IntToStr(i) ;
    C.Left:=CGau ; C.Width:=CLar ; C.Style:=csDropDownList ;
    if i>1 then C.Top:=4+Haut*(i-1) else C.Top:=4 ;
    C.Items.Clear ; C.Values.Clear ; C.Enabled:=False ; C.OnChange:=ComboChanger ;

    L:=TLabel.Create(Self) ; L.Parent:=SBox; L.Name:='TCb'+IntToStr(i) ;
    L.Left:=LGau ; if i>1 then L.Top:=8+Haut*(i-1) else L.Top:=8 ;
    L.Caption:=MsgBox.Mess[4]+IntToStr(i) ;

    if Asc then
       BEGIN
       B:=TCheckbox.Create(Self) ; B.Parent:=SBox; B.Name:='Cbb'+IntToStr(i) ;
       B.Left:=CGau+CLar+10 ; if i>1 then B.Top:=8+Haut*(i-1) else B.Top:=8 ;
       B.Caption:=MsgBox.Mess[6] ; B.Alignment:=taLeftJustify ; //Simon B.OnClick:=CheckChanger ;
       B.Visible:=False ; //Rony 27/10/97 Pas géré pour l'instant
       END ;

    QTri.First ;
    While Not QTri.Eof do
         BEGIN
         C.Values.Add(QTri.Fields[0].AsString) ; C.Items.Add(QTri.Fields[1].AsString) ;
         QTri.Next ;
         END ;
    END ;
QTri.First ; MemoTri.Values.Clear ; MemoTri.Items.Clear ;
While Not QTri.Eof do
     BEGIN
     MemoTri.Values.Add(QTri.Fields[0].AsString) ; MemoTri.Items.Add(QTri.Fields[1].AsString) ;
     QTri.Next ;
     END ;
Ferme(QTri);
END ;

Procedure TFChoiTri.InitCombo ;
Var i : Byte ;
    StTemp : String ;
BEGIN
if Tri='' then
   BEGIN
   for i:=1 to Combien do
     BEGIN
     if i=1 then THValComboBox(FindComponent('Cb'+IntToStr(i))).Enabled:=True
            else THValComboBox(FindComponent('Cb'+IntToStr(i))).Enabled:=False ;
//Simon     TCheckBox(FindComponent('Cbb'+IntToStr(i)+'')).Enabled:=False ;
     END ;
   END else
   BEGIN
   StTemp:=Tri ; i:=1 ;
   While StTemp<>'' do
     BEGIN
     THValComboBox(FindComponent('Cb'+IntToStr(i))).Enabled:=True ;
     THValComboBox(FindComponent('Cb'+IntToStr(i))).Value:=ReadTokenSt(StTemp) ;
//Simon     TCheckBox(FindComponent('Cbb'+IntToStr(i)+'')).Checked:=(ReadTokenSt(StTemp)='ASC') ;
     Inc(i) ;
     END ;
   END ;
END ;

Function TFChoiTri.QuelCombo(Sender : TObject) : Byte ;
Var i : Byte ;
BEGIN
Result:=0 ;
for i:=1 to Combien do
    if THValComboBox(Sender).Name='Cb'+IntToStr(i) then BEGIN Result:=i ; Break ; END ;
END ;

Function  TFChoiTri.CherchePosition(Quel : Byte ; St : String) : Integer ;
Var i,j : Integer ;
BEGIN
Result:=0 ;
if Result=Quel-1 then Exit ;
j:=0 ;
for i:=1 to Length(St) do
    BEGIN
    if St[i]=';'then
       BEGIN
       Inc(j) ; if j=Quel-1 then BEGIN j:=i ; Break ; END ;
       END ;
    END ;
Result:=j ;
END ;

Function  TFChoiTri.FabriqueChaine(St,St1 : String ; j : Byte) : String ;
Var S : String ;
    i,k : Integer ;
BEGIN
if j=0 then
   BEGIN
   Delete(St,1,Pos(';',St)-1) ; Result:=St1+St ;
   END else
   BEGIN
   S:=Copy(St,1,j) ; k:=0 ;
   for i:=j+1 to Length(St) do
       if St[i]=';' then Break else Inc(k) ;
   Delete(St,1,j+k) ; Result:=S+St1+St ;
   END ;
END ;

Procedure TFChoiTri.ComboChanger(Sender : TObject) ;
Var Quel : Byte ;
    S,SLib : String ;
BEGIN
Quel:=QuelCombo(Sender) ; S:=THValComboBox(Sender).Value ;
if THValComboBox(FindComponent('Cb'+IntToStr(Quel+1)))<>Nil then
   if Not THValComboBox(FindComponent('Cb'+IntToStr(Quel+1))).Enabled then
      BEGIN
      THValComboBox(FindComponent('Cb'+IntToStr(Quel+1))).Enabled:=True ;
      CTri.Text:=CTri.Text+S+';' ;
      CLib.Text:=CLib.Text+MemoTri.Items[MemoTri.Values.IndexOf(S)]+';' ;
      Exit ;
      END ;
CTri.Text:=FabriqueChaine(CTri.Text,S,CherchePosition(Quel,CTri.Text)) ;
if CTri.Text<>'' then if CTri.Text[Length(CTri.Text)]<>';' then CTri.Text:=CTri.Text+';' ;
SLib:=MemoTri.Items[MemoTri.Values.IndexOf(S)] ;
CLib.Text:=FabriqueChaine(CLib.Text,SLib,CherchePosition(Quel,CLib.Text)) ;
if CTri.Text<>'' then if CLib.Text[Length(CLib.Text)]<>';' then CLib.Text:=CLib.Text+';' ;
END ;

Function TFChoiTri.EnregOk : Boolean ;
Var St : String ;
BEGIN
Result:=True ; St:=CTri.Text ;
While (St<>'') do
   if Pos(ReadTokenSt(St),St)>0 then BEGIN MsgBox.Execute(5,'','') ; Result:=False ; Exit ; END ;
END ;

procedure TFChoiTri.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
