unit TabLiRub;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  StdCtrls, Buttons, ExtCtrls, Hctrls, ComCtrls, HSysMenu,
{$IFDEF EAGLCLIENT}
  uTOB,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons, 
  {$ENDIF MODENT1}
  hmsgbox, HCompte;

Procedure ChoixTableLibrePourRub(Lefb : TFichierBase ; Ax : String ; Var Comp,Cod,Cod1 : String) ;

type
  TFTabLiRub = class(TForm)
    PBouton: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    PanelChoix: TPanel;
    LT0: TLabel;
    LT1: TLabel;
    LT2: TLabel;
    LT3: TLabel;
    LT4: TLabel;
    CB0: TCheckBox;
    CB1: TCheckBox;
    CB2: TCheckBox;
    CB3: TCheckBox;
    CB4: TCheckBox;
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
    CTri: TEdit;
    CbMemo: THValComboBox;
    CompEt: TEdit;
    TComplement: TLabel;
    Label1: TLabel;
    CompOu: TEdit;
    HM: THMsgBox;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure T0Change(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure T0Click(Sender: TObject);
    procedure CB0MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BAideClick(Sender: TObject);
  private
    Pref : Char ;
    Cod,Cod1,Comp : String ;
    Ind : Integer ;
    Lefb : TFichierBase ;
    CarBour : Char;
    Procedure ChargeLesCombos ;
    Procedure PositionneLesCombos ;
    Procedure GriseDegriseLesCombos(T,T1 : THValComboBox ; Cb : TCheckBox) ;
    Procedure AffecteLibelTable ;
    Procedure PositionneComplement ;
    Function  AnalyseCompOk : Boolean ;
    Procedure FabriqueComplement ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  Hent1;

Procedure ChoixTableLibrePourRub(Lefb : TFichierBase ; Ax : String ; Var Comp,Cod,Cod1 : String) ;
Var FTabLiRub : TFTabLiRub ;
BEGIN
FTabLiRub:=TFTabLiRub.Create(Application) ;
  Try
   FTabLiRub.Lefb:=Lefb ;
   Case Lefb of
        fbGene         : BEGIN
                         FTabLiRub.Pref:='G' ; FTabLiRub.Ind:=1 ;
                         FTabLiRub.Caption:=FTabLiRub.MsgBox.Mess[1] ;
                         FTabLiRub.CarBour:=VH^.Cpta[fbGene].Cb ;
                         END ;
        fbAux          : BEGIN
                         FTabLiRub.Pref:='T' ; FTabLiRub.Ind:=2 ;
                         FTabLiRub.Caption:=FTabLiRub.MsgBox.Mess[2] ;
                         FTabLiRub.CarBour:=VH^.Cpta[fbAux].Cb ;
                         END ;
        fbBudgen       : BEGIN
                         FTabLiRub.Pref:='B' ; FTabLiRub.Ind:=4 ;
                         FTabLiRub.Caption:=FTabLiRub.MsgBox.Mess[3] ;
                         FTabLiRub.CarBour:=#0 ;
                         END ;
        fbAxe1..fbAxe5 : BEGIN
                         FTabLiRub.Pref:='S' ; FTabLiRub.Ind:=3 ;
                         FTabLiRub.Caption:=FTabLiRub.MsgBox.Mess[4] ;
                         FTabLiRub.CarBour:=VH^.Cpta[AxeToFb(Ax)].Cb ;
                         END ;
        fbBudSec1..fbBudSec5 : BEGIN
                               FTabLiRub.Pref:='D' ; FTabLiRub.Ind:=5 ;
                               FTabLiRub.Caption:=FTabLiRub.MsgBox.Mess[5] ;
                               FTabLiRub.CarBour:=#0 ;
                               END ;
     End ;
   FTabLiRub.Comp:=Comp ;
   if Pos('<>',Cod)>0 then
      BEGIN
      FTabLiRub.Cod:=Copy(Cod,1,Pos('<>',Cod)-1) ;
      FTabLiRub.Cod1:=Copy(Cod,Pos('<>',Cod)+2,Length(Cod)) ;
      END else
      BEGIN
      FTabLiRub.Cod:=Cod ;
      FTabLiRub.Cod1:=Cod1 ;
      END ;
   FTabLiRub.ShowModal ;
   Comp:=FTabLiRub.Comp ;
   if FTabLiRub.ModalResult=mrOk then
      BEGIN
      if Pos('<>',Cod)>0 then
         BEGIN Cod:=FTabLiRub.CSorti.Text+'<>'+FTabLiRub.CSortia.Text ; END else
         BEGIN Cod:=FTabLiRub.CSorti.Text ; Cod1:=FTabLiRub.CSortia.Text ; END ;
      END else
      BEGIN
      if Pos('<>',Cod)>0 then
         BEGIN Cod:=FTabLiRub.CEntre.Text+'<>'+FTabLiRub.CEntrea.Text ; END else
         BEGIN Cod:=FTabLiRub.CEntre.Text ; Cod1:=FTabLiRub.CEntrea.Text ; END ;
      END ;
  Finally
   FTabLiRub.Free ;
  End ;
SourisNormale ;
END ;

procedure TFTabLiRub.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFTabLiRub.FormShow(Sender: TObject);
Var St : String ;
    i,j : Integer ;
    C,C1 : THValComboBox ;
    Cb : TCheckBox ;
begin
AffecteLibelTable ; ChargeLesCombos ;
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
      END ;
   END else PositionneLesCombos ;
PositionneComplement ;
end;

Procedure TFTabLiRub.PositionneComplement ;
Var i,j : Integer ;
    StAnd,StOr,St : String ;
BEGIN
if Comp='' then Exit ;
i:=Pos('&',Comp) ; j:=Pos('|',Comp) ; StAnd:='' ; StOr:='' ; St:='' ;
if (i>0) And (j<=0) then StAnd:=Copy(Comp,1,Length(Comp)) else
if (j>0) And (i<=0) then StOr:=Copy(Comp,1,Length(Comp)) else
if (i>0) And (j>i)  then BEGIN StAnd:=Copy(Comp,1,j-1) ; StOr:=Copy(Comp,j,Length(Comp)) ; END else
if (j>0) And (j<i)  then BEGIN StOr:=Copy(Comp,1,i-1) ; StAnd:=Copy(Comp,i,Length(Comp)) ; END ;
if StAnd<>'' then
   BEGIN
   for i:=1 to Length(StAnd) do if StAnd[i]<>'&' then St:=St+StAnd[i] ;
   if St[1]=',' then St:=Copy(St,2,Length(St)) ;
   if St[Length(St)]=',' then St:=Copy(St,1,Length(St)-1) ;
   StAnd:=St ;
   END ;
St:='' ;
if StOr<>'' then
   BEGIN
   for i:=1 to Length(StOr) do if StOr[i]<>'|' then St:=St+StOr[i] ;
   if St[1]=',' then St:=Copy(St,2,Length(St)) ;
   if St[Length(St)]=',' then St:=Copy(St,1,Length(St)-1) ;
   StOr:=St ;
   END ;
CompEt.Text:=StAnd ; CompOu.Text:=StOr ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/03/2003
Modifié le ... :   /  /    
Description .. : Chargement des combos tables libres
Suite ........ : Réécriture en eAGL : on ne passe plus par des requêtes 
Suite ........ : paramétrées.
Mots clefs ... : 
*****************************************************************}
Procedure TFTabLiRub.ChargeLesCombos ;
Var Sql : String ;
    i : Integer ;
    QLoc : TQuery;
BEGIN
  for i := 0 to 9 do
  begin
    Sql := 'Select NT_NATURE,NT_LIBELLE From NATCPTE Where NT_TYPECPTE="'+Pref+'0'+IntToStr(i)+'"';
    QLoc := OpenSQL ( SQL, True );
    try
      if QLoc.Eof then
      begin
        TLabel(FindComponent('LT'+IntToStr(i))).Enabled:=False ;
        TLabel(FindComponent('LT'+IntToStr(i)+IntToStr(i))).Enabled:=False ;
        THValComboBox(FindComponent('T'+IntToStr(i))).Enabled:=False ;
        THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Enabled:=False ;
        TCheckBox(FindComponent('CB'+IntToStr(i))).Enabled:=False ;
      end else
      begin
        THValComboBox(FindComponent('T'+IntToStr(i))).Values.Clear ;
        THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Values.Clear ;
        THValComboBox(FindComponent('T'+IntToStr(i))).Items.Clear ;
        THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Items.Clear ;
        THValComboBox(FindComponent('T'+IntToStr(i))).Values.Add('') ;
        THValComboBox(FindComponent('T'+IntToStr(i))).Items.Add(Traduirememoire('<<Tous>>')) ;
        while not QLoc.Eof do
        begin
          THValComboBox(FindComponent('T'+IntToStr(i))).Values.Add(QLoc.Fields[0].AsString) ;
          THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Values.Add(QLoc.Fields[0].AsString) ;
          THValComboBox(FindComponent('T'+IntToStr(i))).Items.Add(QLoc.Fields[1].AsString) ;
          THValComboBox(FindComponent('T'+IntToStr(i)+IntToStr(i))).Items.Add(QLoc.Fields[1].AsString) ;
          QLoc.Next ;
        end;
      end;
    finally
      Ferme ( QLoc );
    end;
  end;
end;

Procedure TFTabLiRub.PositionneLesCombos ;
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

procedure TFTabLiRub.T0Change(Sender: TObject);
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

Function TFTabLiRub.AnalyseCompOk : Boolean ;
Var i : Integer ;
    St : String ;
    TypChar : Set of Char ;
BEGIN
Result:=False ;
TypChar:=['0'..'9','A'..'Z',CarBour,','] ;
St:=CompEt.Text ;
if St<>'' then
   BEGIN
   for i:=1 to Length(St) do if Not (St[i] in TypChar) then
       BEGIN HM.Execute(0,'','') ; CompEt.SetFocus ; Exit ; END ;
   END ;
St:=CompOu.Text ;
if St<>'' then
   BEGIN
   for i:=1 to Length(St) do if Not (St[i] in TypChar) then
       BEGIN HM.Execute(0,'','') ; CompOu.SetFocus ; Exit ; END ;
   END ;
Result:=True ;
END ;

Procedure TFTabLiRub.FabriqueComplement ;
Var i : Integer ;
    St,St1 : String ;
BEGIN
St:=CompEt.Text ; St1:='' ; Comp:='' ;
if CSorti.Text='' then Exit ;
if St<>'' then
   BEGIN
   if St[1]=',' then St:=Copy(St,2,Length(St)) ;
   if St[Length(St)]=',' then St:=Copy(St,1,Length(St)-1) ;
   St1:='&' ;
   for i:=1 to Length(St) do
      if St[i]=',' then St1:=St1+St[i]+'&' else St1:=St1+St[i] ;
   Comp:=St1 ;
   END ;
St:=CompOu.Text ; St1:='' ;
if St<>'' then
   BEGIN
   if St[1]=',' then St:=Copy(St,2,Length(St)) ;
   if St[Length(St)]=',' then St:=Copy(St,1,Length(St)-1) ;
   St1:='|' ;
   for i:=1 to Length(St) do
      if St[i]=',' then St1:=St1+St[i]+'|' else St1:=St1+St[i] ;
   Comp:=Comp+','+St1 ;
   END ;
END ;

procedure TFTabLiRub.BValiderClick(Sender: TObject);
begin
if Not AnalyseCompOk then Exit ;
FabriqueComplement ;
ModalResult:=mrOk ;
end;

procedure TFTabLiRub.T0Click(Sender: TObject);
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

procedure TFTabLiRub.CB0MouseUp(Sender: TObject; Button: TMouseButton;
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

Procedure TFTabLiRub.GriseDegriseLesCombos(T,T1 : THValComboBox ; Cb : TCheckBox) ;
BEGIN
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

Procedure TFTabLiRub.AffecteLibelTable ;
Var StLib : String ;
    i     : Byte ;
    LLibelTable : HTStringList ;
BEGIN
LLibelTable:=HTStringList.Create ; GetLibelleTableLibre(Pref,LLibelTable) ;
For i:=0 to LLibelTable.Count-1 do
    BEGIN
    StLib:=LLibelTable[i] ;
    TLabel(FindComponent('LT'+IntToStr(i))).Caption:=ReadTokenSt(StLib) ;
    END ;
LLibelTable.Free ;
END ;

procedure TFTabLiRub.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
