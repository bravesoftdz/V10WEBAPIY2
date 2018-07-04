unit Rupanal;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  Buttons,
  ExtCtrls,
  Hctrls,
{$IFDEF EAGLCLIENT}
  uTob,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  Hmsgbox,
  Ent1,
  HCompte,
  HSysMenu ;

Procedure RuptureAnalytique(TypRup,CodRup : String ; Var Clibre : String) ;

type
  TFRupanal = class(TForm)
    PBouton  : TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    Pappli   : TPanel;
    Ptop     : TPanel;
    EClibre  : TEdit;
    TEClibre : THLabel;
    Sb1      : TScrollBox;
    TTri1    : THLabel;
    Tri1     : THValComboBox;
    TTri2    : THLabel;
    Tri2     : THValComboBox;
    TTri3    : THLabel;
    TTri4    : THLabel;
    Tri4     : THValComboBox;
    TTri5    : THLabel;
    Tri5     : THValComboBox;
    TTri6    : THLabel;
    Tri6     : THValComboBox;
    TTri7    : THLabel;
    Tri7     : THValComboBox;
    Tri3     : THValComboBox;
    TTri8    : THLabel;
    Tri8     : THValComboBox;
    TTri9    : THLabel;
    Tri9     : THValComboBox;
    TTri10   : THLabel;
    Tri10    : THValComboBox;
    TTri11   : THLabel;
    Tri11    : THValComboBox;
    TTri12   : THLabel;
    Tri12    : THValComboBox;
    TTri13   : THLabel;
    Tri13    : THValComboBox;
    TTri14   : THLabel;
    Tri14    : THValComboBox;
    TTri15   : THLabel;
    Tri15    : THValComboBox;
    TTri16   : THLabel;
    Tri16    : THValComboBox;
    TTri17   : THLabel;
    Tri17    : THValComboBox;
    TCodPlan : THLabel;
    TLibPlan : THLabel;
    CodPlan  : TEdit;
    LibPlan  : TEdit;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure Tri1Click(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
   TypRup,CodRup,Clibre : String ;
   Axe : String ;
   DataType : String ;
   Combien : Byte ;
   Procedure Queltt ;
   Procedure Libelle ;
   Procedure AfficheCombo ;
   Procedure ChangementCombo(Sender : Tobject) ;
   Function  QuelCombo(Sender: TObject) : Byte ;
   Function  EnregOk : Boolean ;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HEnt1;

Procedure TFRupanal.Queltt ;
Var i: Byte ;
BEGIN
i:=StrToInt(Axe[2]) ;
Case i of
   1:DataType:='ttStrucrse1' ;
   2:DataType:='ttStrucrse2' ;
   3:DataType:='ttStrucrse3' ;
   4:DataType:='ttStrucrse4' ;
   5:DataType:='ttStrucrse5' ;
 End ;
END ;

Procedure TFRupanal.AfficheCombo ;
Var Q : TQuery ;
    i : Byte ;
    St,StTemp : String ;
BEGIN
Q:=OpenSql('Select Count(SS_SOUSSECTION) from STRUCRSE Where SS_AXE="'+Axe+'"',True) ;
ComBien:=Q.Fields[0].AsInteger ; Ferme(Q) ;
for i:=1 to 17 do
   BEGIN
    TControl(FindComponent('Tri'+IntToStr(i)+'')).Visible:=(i<=Combien) ;
    TControl(FindComponent('TTri'+IntToStr(i)+'')).Visible:=(i<=Combien) ;
    THValComboBox(FindComponent('Tri'+IntToStr(i)+'')).DataType:=DataType ;
   END ;
if Clibre<>'' then
  BEGIN
   EClibre.Text:=Clibre ; StTemp:=Clibre ; i:=1 ;
   While StTemp<>'' do
     BEGIN
     St:=ReadTokenSt(StTemp) ;
     THValComboBox(FindComponent('Tri'+IntToStr(i)+'')).Value:=St ; Inc(i) ;
     END ;
  END else
  BEGIN
    for i:=2 to Combien do
        THValComboBox(FindComponent('Tri'+IntToStr(i)+'')).Enabled:=False ;
  END ;
END ;

Procedure TFRupanal.Libelle ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CC_LIBELLE from CHOIXCOD Where CC_TYPE="'+TypRup+'" And '+
           'CC_CODE="'+CodRup+'"',True) ;
CodPlan.Text:=CodRup ; LibPlan.Text:=Q.Fields[0].AsString ; Ferme(Q) ;
END ;

procedure TFRupanal.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Libelle ; Queltt ; AfficheCombo ;
end;

procedure TFRupanal.BFermeClick(Sender: TObject);
begin Close ; end;

Function TFRupanal.QuelCombo(Sender: TObject) : Byte ;
Var i : Byte ;
BEGIN
Result:=0 ;
for i:=1 to Combien do
   if THValComboBox(Sender).Name='Tri'+IntToStr(i) then BEGIN Result:=i ; Break ; END ;
END ;

procedure TFRupanal.Tri1Click(Sender: TObject);
begin ChangementCombo(Sender) ; end;

Procedure TFRupanal.ChangementCombo(Sender : Tobject) ;
Var Lequel : Byte ;
    i,j : Byte ;
    S,S1,S2 : String ;
BEGIN
Lequel:=QuelCombo(Sender) ; S:=THValComboBox(Sender).Value ;
if Not THValComboBox(FindComponent('Tri'+IntToStr(Lequel+1)+'')).Enabled then
   BEGIN
   THValComboBox(FindComponent('Tri'+IntToStr(Lequel+1)+'')).Enabled:=True ;
   EClibre.Text:=EClibre.Text+S+';' ; Exit ;
   END ;
j:=0 ;
if j=Lequel-1 then j:=0
else BEGIN
      for i:=1 to Length(EClibre.Text) do
       BEGIN
        if EClibre.Text[i]=';'then
           BEGIN
            Inc(j) ; if j=Lequel-1 then BEGIN j:=i ; Break ; END ;
           END ;
       END ;
     END ;
S1:=Copy(EClibre.Text,1,j) ;
S2:=Copy(EClibre.Text,Length(S1)+(Length(S)+1),Length(EClibre.Text)) ;
EClibre.Text:=S1+S+S2 ;
if EClibre.Text[Length(EClibre.Text)]<>';' then EClibre.Text:=EClibre.Text+';' ;
END ;

Function TFRupanal.EnregOk : Boolean ;
Var Q : TQuery ;
    St,St1 : String ;
    Ok : Boolean ;
BEGIN
Result:=True ;
St:=EClibre.Text ; Ok:=True ;
While (St<>'')AND Ok do
   BEGIN
   St1:=ReadTokenSt(St) ;
   if Pos(St1,St)>0 then BEGIN MsgBox.Execute(0,'','') ; Result:=False ; Exit ; END ;
   Q:=OpenSql('Select SS_SOUSSECTION from STRUCRSE Where SS_AXE="'+Axe+'" AND '+
              'SS_SOUSSECTION="'+St1+'"',True) ;
   if Q.Eof then BEGIN MsgBox.Execute(1,'','') ; Ok:=False ; Result:=False ; END ;
   Ferme(Q) ;
   END ;
END ;

procedure TFRupanal.BValiderClick(Sender: TObject);
begin if EnregOk then BEGIN Clibre:=EClibre.Text ; Close ; END ; end;

Procedure RuptureAnalytique(TypRup,CodRup : String ; Var Clibre : String) ;
var FRupanal : TFRupanal ;
BEGIN
 If Blocage(['nrCloture'],False,'nrAucun') then Exit ;
 FRupanal:=TFRupanal.Create(Application) ;
  Try
   if VH^.Cpta[AxeToFb(FRupanal.Axe)].Structure then
      BEGIN
       FRupanal.Axe:='A'+TypRup[3] ;
       FRupanal.TypRup:=TypRup ;
       FRupanal.CodRup:=CodRup ;
       FRupanal.Clibre:=Clibre ;
       FRupanal.ShowModal ;
      END ;
  Finally
   Clibre:=FRupanal.Clibre ;
   FRupanal.Free ;
  End ;
Screen.Cursor:=crDefault ;
END ;

procedure TFRupanal.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

end.
