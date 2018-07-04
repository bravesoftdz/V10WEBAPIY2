unit DateEdt;

interface

uses                          
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, StdCtrls, Spin, Buttons, ExtCtrls, Ent1, HEnt1, Mask, HCtrls,
  ParamDat, UtilEdt1;

Procedure ChoixDateEdt(var NbDates : Integer ; var ListePer,ListeDates : String) ;

procedure GetCorrespDates(St : String ; var StD1,StD2 : String) ;
procedure DecodePeriode(var StP,Per : String ; var OkDebPer,OkFinPer : boolean ; var NumPer : Integer) ;

type
  TFDateEdt = class(TForm)
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    PBouton: TPanel;
    BValider: TBitBtn;
    HelpBtn: TBitBtn;
    BFerme: TBitBtn;
    GbDate: TGroupBox;
    TPD_NBDATEDEF: TLabel;
    TD1: TLabel;
    TD2: TLabel;
    TD3: TLabel;
    TD4: TLabel;
    TD5: TLabel;
    TD6: TLabel;
    TD7: TLabel;
    TD8: TLabel;
    TD9: TLabel;
    TD10: TLabel;
    TD11: TLabel;
    TD12: TLabel;
    PD_NBDATEDEF: TSpinEdit;
    PDateDef: TPanel;
    TPD_DATEDEF: TLabel;
    PD_DATEDEF: TEdit;
    D1: TMaskEdit;
    D2: TMaskEdit;
    D3: TMaskEdit;
    D4: TMaskEdit;
    D5: TMaskEdit;
    D6: TMaskEdit;
    D13: TMaskEdit;
    D14: TMaskEdit;
    D15: TMaskEdit;
    D16: TMaskEdit;
    D17: TMaskEdit;
    D18: TMaskEdit;
    TD13: TLabel;
    TD14: TLabel;
    TD15: TLabel;
    TD16: TLabel;
    TD17: TLabel;
    TD18: TLabel;
    TD19: TLabel;
    TD20: TLabel;
    TD21: TLabel;
    TD22: TLabel;
    TD23: TLabel;
    TD24: TLabel;
    D7: TMaskEdit;
    D8: TMaskEdit;
    D9: TMaskEdit;
    D10: TMaskEdit;
    D11: TMaskEdit;
    D12: TMaskEdit;
    D19: TMaskEdit;
    D20: TMaskEdit;
    D21: TMaskEdit;
    D22: TMaskEdit;
    D23: TMaskEdit;
    D24: TMaskEdit;
    FPeriodicite: THValComboBox;
    TFPeriodicite: TLabel;
    procedure FormShow(Sender: TObject);
    procedure PD_NBDATEDEFChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure D1KeyPress(Sender: TObject; var Key: Char);
    procedure D1DblClick(Sender: TObject);
    procedure FPeriodiciteChange(Sender: TObject);
  private
    ListePer,ListeDates : String ;
    NbDates : integer ;
    Loading    : boolean ;
    Procedure InitDates(St : String) ;
    Function  ChangeChampDateDef(LaVal : String) : String ;
    procedure RempliListeDates ;
    procedure RempliLesDates ;
    procedure PeriodiciteChange(Iperiodicite,DerDate : Integer ;FP12,FP1 : TMaskEdit ; Var TabD : TTabDate24) ;
  public
    { Déclarations publiques }
  end;

implementation

uses DateDef ;

{$R *.DFM}

Procedure ChoixDateEdt(var NbDates : Integer ; var ListePer,ListeDates : String) ;
var FDateEdt : TFDateEdt ;
BEGIN
FDateEdt:=TFDateEdt.Create(Application) ;
 Try
 FDateEdt.NbDates:=NbDates ;
 FDateEdt.ListePer:=ListePer ;
 FDateEdt.ListeDates:=ListeDates ;
 FDateEdt.ShowModal ;
 if FDateEdt.ModalResult=mrOk then
   BEGIN
   NbDates:=FDateEdt.NbDates ;
   ListePer:=FDateEdt.ListePer ;
   ListeDates:=FDateEdt.ListeDates ;
   END ;
 Finally
 FDateEdt.Free ;
 End ;
SourisNormale ;
END ;

procedure QuelExoPer (var St : String ; var DD1,DD2 : TDateTime) ;
BEGIN
DD1:=VH^.Encours.Deb ; DD2:=VH^.Encours.Fin ;
   //Exercice N-3
   if Pos('---',St)>0 then BEGIN DD1:=VH^.Encours.Deb ; DD2:=VH^.Encours.Fin ; Delete(St,Pos('---',St),3) END else
   //Exercice N-2
      if Pos('--',St)>0 then BEGIN DD1:=VH^.Encours.Deb ; DD2:=VH^.Encours.Fin ; Delete(St,Pos('--',St),2) END else
   //Exercice N-1
          if Pos('-',St)>0 then BEGIN DD1:=VH^.Precedent.Deb ; DD2:=VH^.Precedent.Fin ; Delete(St,Pos('-',St),1) ; END else
   //Exercice N+1
              if Pos('+',St)>0 then BEGIN DD1:=VH^.Suivant.Deb ; DD2:=VH^.Suivant.Fin ; Delete(St,Pos('+',St),1) ; END ;
END ;

procedure LitPeriode (var DD1 : TDateTime ; Per : String ; NumPer : Integer ; Deb,Prem : boolean) ;
var  a,m,j : Word ;
begin
DecodeDate(DD1,a,m,j) ;
Case Per[1] of
    //Exercice entier
    'N' :  ;
    //Mois
    'M' : if Deb then DD1:=EncodeDate(a,NumPer,1) ;
    //Bimestre
    'B' : BEGIN
          if Prem then DD1:=EncodeDate(a,(NumPer*2)-1,1) else DD1:=EncodeDate(a,(NumPer*2),1) ;
          END ;
    //Trimestre
    'T' : BEGIN
          if Prem then DD1:=EncodeDate(a,(NumPer*3)-2,1) else DD1:=EncodeDate(a,(NumPer*3),1) ;
          END ;
    //Quadrimestre
    'Q' : BEGIN
          if Prem then DD1:=EncodeDate(a,(NumPer*4)-3,1) else DD1:=EncodeDate(a,(NumPer*4),1) ;
          END ;
    //Semestre
    'S' : BEGIN
          if Prem then DD1:=EncodeDate(a,(NumPer*6)-5,1) else DD1:=EncodeDate(a,(NumPer*6),1) ;
          END ;
    End ;
end;

procedure DecodePeriode(var StP : String ; var Per : String ; var OkDebPer,OkFinPer : boolean ; var NumPer : Integer) ;
var St : string ;
BEGIN
St:=StP ;
   Per:=Copy(St,1,1) ; Delete(St,1,1) ; StP:=St ;
//Modif sur la période sélectionnée
OkDebPer:=True ; OkFinPer:=True ;
//Du début de l'exercice jusqu'à la fin de la période (incluse)
   if Pos('<',St)>0 then BEGIN OkDebPer:=False ; Delete(St,Pos('<',St),1) END else
//Du début de la période jusqu'à la fin de l'exercice
    if Pos('>',St)>0 then BEGIN OkFinPer:=False ; Delete(St,Pos('>',St),1) END ;
NumPer:=0 ;
   if St<>'' then NumPer:=StrToInt(St) ;
END ;

procedure GetCorrespDates(St : String ; var StD1,StD2 : String) ;
var DD1,DD2  : TDateTime ;
    OkDebPer,OkFinPer : boolean ;
    Per : String ;
    NumPer : Integer ;
BEGIN
DD1:=iDate1900 ; DD2:=iDate2099 ; NumPer:=0 ;
if (St='NP') then
   BEGIN
   DD1:=VH^.Encours.Deb ; DD2:=VH^.EnCours.Fin ;
   END else
if (St<>'') then
   BEGIN
   QuelExoPer(St,DD1,DD2) ;
   DecodePeriode(St,Per,OkDebPer,OkFinPer,NumPer) ;
   LitPeriode(DD1,Per,NumPer,OkDebPer,TRUE) ;
   LitPeriode(DD2,Per,NumPer,OkFinPer,FALSE) ;
   DD2:=FinDeMois(DD2) ;
   END ;
StD1:=DateToStr(DD1) ;
StD2:=DateToStr(DD2) ;
//if (QuelPer.ItemIndex=-1) and (QuelPer.Items.Count>0) then QuelPer.ItemIndex:=0 ;
END ;

Procedure TFDateEdt.InitDates(St : String) ;
Var i : Integer ;
    Per,StD1,StD2 : String ;

BEGIN
i:=1 ;
while i<=23 do
   BEGIN
   Per:=ReadTokenSt(St) ;
   GetCorrespDates(Per,StD1,StD2) ;
   TEdit(FindComponent('D'+IntToStr(i))).Text:=StD1 ;
   // impaire -> +1
   TEdit(FindComponent('D'+IntToStr(i))).Enabled:=(((i div 2)+1)<=PD_NBDATEDEF.Value) ;
   TLabel(FindComponent('TD'+IntToStr(i))).Enabled:=(((i div 2)+1)<=PD_NBDATEDEF.Value) ;
   inc(i) ;
   TEdit(FindComponent('D'+IntToStr(i))).Text:=StD2 ;
   TEdit(FindComponent('D'+IntToStr(i))).Enabled:=((i div 2)<=PD_NBDATEDEF.Value) ;
   TLabel(FindComponent('TD'+IntToStr(i))).Enabled:=((i div 2)<=PD_NBDATEDEF.Value) ;
   inc(i) ;
   END ;
END ;


procedure TFDateEdt.FormShow(Sender: TObject);
begin
Loading:=True ;
PD_DATEDEF.Text:=ListePer ;
PD_NBDATEDEF.Value:=NbDates ;
InitDates(ListePer) ;
RempliLesDates ;
Loading:=False ;
end;

Function TFDateEdt.ChangeChampDateDef(LaVal : String) : String ;
Var i,j : Integer ;
    St,St1 : String ;
BEGIN
St:=PD_DATEDEF.Text ;
j:=0 ; St1:='' ;
for i:=1 to Length(St) do
    BEGIN
    if j=PD_NBDATEDEF.Value then Break ;
    St1:=St1+St[i] ;
    if (St[i]=';') then Inc(j) ;
    END ;
if j<PD_NBDATEDEF.Value then St1:=St1+LaVal ;
if (St1<>'') or (j=PD_NBDATEDEF.Value) then PD_DATEDEF.Text:=St1 ;
Result:=St1 ;
END ;

procedure TFDateEdt.PD_NBDATEDEFChange(Sender: TObject);
var St1,St : String ;
begin
if Loading then Exit ;
St1:=PD_DATEDEF.Text ;
while St1<>'' do
  BEGIN
  St:=ReadTokenSt(St1) ;
  END ;
St:=ChangeChampDateDef(HM.Mess[15]+';') ;
InitDates(St) ;
end;

procedure TFDateEdt.BValiderClick(Sender: TObject);
begin
ListePer:=PD_DATEDEF.Text ;
NbDates:=Round(PD_NbDATEDEF.Value) ;
RempliListeDates ;
end;

procedure TFDateEdt.D1KeyPress(Sender: TObject; var Key: Char);
begin
PARAMDATE (Self,Sender,Key) ;
end;

procedure TFDateEdt.RempliListeDates ;
var S : String ;
    i : integer ;
begin
S:='' ;
for i:=1 to 24 do
  if TMaskEdit(FindComponent('D'+IntToStr(i))).Enabled then S:=S+TMaskEdit(FindComponent('D'+IntToStr(i))).Text+';' ;
ListeDates:=S ;
end ;

procedure TFDateEdt.RempliLesDates ;
var S,SD : String ;
    i : integer ;
begin
S:=ListeDates ;
SD:=ReadTokenSt(S) ; i:=1 ;
while SD<>'' do
  BEGIN
  TMaskEdit(Findcomponent('D'+IntToStr(i))).Text:=SD ;
  SD:=ReadTokenSt(S) ;
  inc(i) ;
  END ;
end ;

procedure TFDateEdt.D1DblClick(Sender: TObject);
var StD1,StD2,St,S,Dat,StPer : String ;
    TD,TD1 : TMaskEdit ;
    i,CeluiLa,LAutre,LeDernier : integer ;
begin
TD:=TMaskEdit(Sender) ;
CeluiLa:=StrToInt(Copy(TD.Name,2,Length(TD.Name)-1)) ;
if (CeluiLa mod 2<>0) then
   BEGIN
   LAutre:=CeluiLa+1 ;
   TD1:=TMaskEdit(FindComponent('D'+IntToStr(LAutre))) ;
   END else
   BEGIN
   LAutre:=CeluiLa-1 ;
   TD1:=TD ;
   TD:=TMaskEdit(FindComponent('D'+IntToStr(LAutre))) ;
   END ;
St:=PD_DATEDEF.Text ; StPer:='' ;
if CeluiLa>LAutre then LeDernier:=CeluiLa
                  else LeDernier:=LAutre ;
for i:=1 to ((LeDernier) div 2) do
  BEGIN
  S:=ReadTokenSt(St) ;
  if (S<>'') and (i<((LeDernier) div 2)) then StPer:=StPer+S+';' ;
  END ;
Dat:=S ;
ParamDateDefPlEdt(Dat) ;
GetCorrespDates(Dat,StD1,StD2) ;
TD.Text:=StD1 ; TD1.Text:=StD2 ;
StPer:=StPer+Dat+';' ;
While S<>'' do
  BEGIN
  S:=ReadTokenSt(St) ;
  if S<>'' then StPer:=StPer+S+';' ;
  END ;
PD_DATEDEF.Text:=StPer ;
end;

procedure TFDateEdt.FPeriodiciteChange(Sender: TObject);
var Der,i : integer ;
     T : TMaskEdit ;
     TabD : TTabDate24 ;
begin
T:=D24 ; Der:=24 ;
for i:=1 to 24 do
  if not TMaskEdit(FindComponent('D'+IntToStr(i))).Enabled then
    BEGIN
    T:=TMaskEdit(FindComponent('D'+IntToStr(i))) ;
    Der:=i-1 ;
    Break ;
    END ;
PeriodiciteChange(Fperiodicite.ItemIndex,Der,T,D1,TabD) ;
For i:=1 to Der do TMaskEdit(FindComponent('D'+IntToStr(i))).Text:=DateToStr(TabD[i]) ;
end;

procedure TFDateEdt.PeriodiciteChange(Iperiodicite,DerDate : Integer ;FP12,FP1 : TMaskEdit ; Var TabD : TTabDate24) ;
Var Choix,i : Integer ;
    a,m,j, JMax         : Word ;
begin
Choix:=IPeriodicite ;
If Choix<=-1 Then Choix:=0 ;
i:=3 ;
TabD[1]:=StrToDate(FP1.Text) ;   { Calcul à partir de la date d'arrêtée }
Case IPeriodicite of   { 0..5 : Période Mensuel ; 6 : Quinzaine ; 7 ; Hebdomadaire }
     0,1,2,3,4,5 : BEGIN
                   while i<=24 Do { Pour chaque Fourchette de dates, en mensuel }
                       BEGIN        { Date d'arrivée = (Date départ + nb Mois) - 1 jour }
                       TabD[i]:=PlusMois(TabD[i-2], (Choix+1)) ;
                       if i<24 then TabD[i-1]:=TabD[i]-1 ;
                       inc(i,2) ;
                       END ;
                   END ;
     6           : BEGIN
                   while i<=24 Do { Pour chaque Fourchette de dates, en Quinzaine }
                       BEGIN
                       DecodeDate(TabD[i-2],a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       If J<15 then
                          BEGIN    { Date d'arrivée = (Date départ + 15 jours) - 1 jour, si date avant le 15 du mois }
                          TabD[i]:=TabD[i-2]+(15-1) ;
                          END ELse
                          BEGIN    { Date d'arrivée = (Date départ + (Nb jours Max du mois - 15 jours) ) - 1 jour, si date aprés le 15 du mois }
                          TabD[i]:=TabD[i-2]+(JMax-15) ;
                          END ;
                       if i<24 then TabD[i-1]:=TabD[i]-1 ;
                       inc(i,2) ;
                       END ;
                   END ;
     7           : BEGIN
                   while i<=24 Do { Pour chaque Fourchette de dates, en Hebdo }
                       BEGIN        { Date d'arrivée = (Date départ + (7 jours - 1 jour)  }
                       TabD[i]:=TabD[i-2]+(7) ;
                       if i<24 then TabD[i-1]:=TabD[i]-1 ;
                       inc(i,2) ;
                       END ;
                   END ;
     End ;
end;

end.
