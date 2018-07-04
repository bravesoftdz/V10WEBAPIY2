unit DateDef;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, hmsgbox, StdCtrls, Hctrls, Buttons, ExtCtrls, HEnt1, Ent1;

Procedure ParamDateDefPlEdt(Var Dat : String) ;

type
  TFDateDef = class(TForm)
    PBouton: TPanel;
    BValider: TBitBtn;
    HelpBtn: TBitBtn;
    BFerme: TBitBtn;
    TExo: THLabel;
    Exo: THValComboBox;
    TPer: THLabel;
    Per: THValComboBox;
    QuelPer: THValComboBox;
    TDetPer: THLabel;
    DetPer: THValComboBox;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    NumPer: THNumEdit;
    procedure FormShow(Sender: TObject);
    procedure PerChange(Sender: TObject);
    procedure QuelPerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    Dat : String ;
    Function  GetMois(i : Integer) : String ;
    Function  Get2Mois(i,j : Integer) : String ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Procedure ParamDateDefPlEdt(Var Dat : String) ;
var FDateDef : TFDateDef ;
BEGIN
FDateDef:=TFDateDef.Create(Application) ;
 Try
 FDateDef.Dat:=Dat ;
 FDateDef.ShowModal ;
 Dat:=FDateDef.Dat ;
 Finally
 FDateDef.Free ;
 End ;
SourisNormale ;
END ;

procedure TFDateDef.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFDateDef.FormShow(Sender: TObject);
Var St : String ;
begin
St:=Dat ;
Exo.ItemIndex:=0 ; Per.ItemIndex:=0 ; DetPer.ItemIndex:=0 ;
if St='NP' then
   BEGIN
   QuelPer.Enabled:=False ; DetPer.Enabled:=False ; TDetPer.Enabled:=False ;
   END else
   BEGIN
   Per.Value:=Copy(St,1,1) ; Delete(St,1,1) ;
   if Pos('---',St)>0 then BEGIN Exo.ItemIndex:=3 ; Delete(St,Pos('---',St),3) END else
      if Pos('--',St)>0 then BEGIN Exo.ItemIndex:=2 ; Delete(St,Pos('--',St),2) END else
          if Pos('-',St)>0 then BEGIN Exo.ItemIndex:=1 ; Delete(St,Pos('-',St),1) END else
              if Pos('+',St)>0 then BEGIN Exo.ItemIndex:=4 ; Delete(St,Pos('+',St),1) END ;
   DetPer.ItemIndex:=0 ;
   if Pos('<',St)>0 then BEGIN DetPer.ItemIndex:=1 ; Delete(St,Pos('<',St),1) END else
      if Pos('>',St)>0 then BEGIN DetPer.ItemIndex:=2 ; Delete(St,Pos('>',St),1) END ;
   if St<>'' then
      BEGIN
      QuelPer.ItemIndex:=StrToInt(St)-1 ;
      NumPer.Value:=StrToInt(St) ;
      END ;
   END ;
if (QuelPer.ItemIndex=-1) and (QuelPer.Items.Count>0) then QuelPer.ItemIndex:=0 ;
end;

Function TFDateDef.GetMois(i : Integer) : String ;
BEGIN Result:=HM.Mess[i-1] ; END ;

Function TFDateDef.Get2Mois(i,j : Integer) : String ;
BEGIN Result:=HM.Mess[12]+' '+LowerCase(GetMois(i))+' '+HM.Mess[13]+' '+LowerCase(GetMois(j)) ; END ;

procedure TFDateDef.PerChange(Sender: TObject);
Var i : Integer ;
begin
QuelPer.Enabled:=(Per.Value<>'N') ; TDetPer.Enabled:=(Per.Value<>'N') ; DetPer.Enabled:=(Per.Value<>'N') ;
QuelPer.Items.Clear ;
Case Per.ItemIndex of
     0 : BEGIN DetPer.ItemIndex:=0 ; NumPer.Text:='0' ; END ;
     1 : BEGIN for i:=1 to 12 do QuelPer.Items.Add(GetMois(i)) ; NumPer.Max:=12 ; END ;
     2 : BEGIN
         i:=1 ;
         While i<12 do BEGIN QuelPer.Items.Add(Get2Mois(i,i+1)) ; i:=i+2 ; END ;
         NumPer.Max:=6 ;
         END ;
     3 : BEGIN
         i:=1 ;
         While i<12 do BEGIN QuelPer.Items.Add(Get2Mois(i,i+2)) ; i:=i+3 ; END ;
         NumPer.Max:=4 ;
         END ;
     4 : BEGIN
         i:=1 ;
         While i<12 do BEGIN QuelPer.Items.Add(Get2Mois(i,i+3)) ; i:=i+4 ; END ;
         NumPer.Max:=3 ;
         END ;
     5 : BEGIN
         i:=1 ;
         While i<12 do BEGIN QuelPer.Items.Add(Get2Mois(i,i+5)) ; i:=i+6 ; END ;
         NumPer.Max:=2 ;
         END ;
     End ;
if QuelPer.Items.Count>0 then BEGIN QuelPer.ItemIndex:=0 ; QuelPerClick(Nil) ; END ;
DetPer.Enabled:=(Per.ItemIndex=1) ;
end;

procedure TFDateDef.QuelPerClick(Sender: TObject);
begin NumPer.Value:=QuelPer.ItemIndex+1 ; end;

procedure TFDateDef.BValiderClick(Sender: TObject);
begin
if HM.Execute(14,'','')<>mrYes then Exit ;
Dat:='' ;
if NumPer.Text<>'0' then Dat:=Per.Value+NumPer.Text+Exo.Value+DetPer.Value
                    else Dat:=Per.Value+Exo.Value+DetPer.Value ;
Close ;
end;

end.
