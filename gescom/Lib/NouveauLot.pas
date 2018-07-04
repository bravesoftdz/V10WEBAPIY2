unit NouveauLot;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Mask, HTB97, ExtCtrls, HPanel, UTof, UTob, hmsgbox, SaisUtil,
  HEnt1;

function CreatLot(TobDispoLot : TOB ; Art, Dep : string) : boolean;

type
  TFNouveauLot = class(TForm)
    PGeneral: THPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    baide: TToolbarButton97;
    bfermer: TToolbarButton97;
    bvalider: TToolbarButton97;
    NumeroLot: THCritMaskEdit;
    DateLot: THCritMaskEdit;
    Physique: THCritMaskEdit;
    TNumeroLot: THLabel;
    TDateLot: THLabel;
    TPhysique: THLabel;
    TArticle: THLabel;
    TDepot: THLabel;
    Article: THLabel;
    Depot: THLabel;
    MsgBox: THMsgBox;

    procedure bvaliderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure PhysiqueExit(Sender: TObject);

  private
    { Déclarations privées }


    LotArt : string;
    LotDep : string;
    TOBDesLots: TOB;
    
    function  NumLotExist : boolean;
    procedure TobLotPutValeur;
    procedure InitForm;
    function  ControlFormOK : boolean;

  public
    { Déclarations publiques }
  end;

var   FNouveauLot: TFNouveauLot;
      ValideOK : boolean;

implementation

{$R *.DFM}

function CreatLot(TobDispoLot : TOB ; Art, Dep : string) : boolean;
begin
ValideOK := false;
Result := false;
FNouveauLot := TFNouveauLot.Create(Application);
FNouveauLot.LotArt := Art;
FNouveauLot.LotDep := Dep;
FNouveauLot.TOBDesLots := TobDispoLot;
FNouveauLot.ShowModal ;
if ValideOK then Result := true;
end;


procedure TFNouveauLot.FormShow(Sender: TObject);
begin
Article.Caption := RechDom('GCARTICLE', LotArt ,false);
Depot.Caption := RechDom('GCDEPOT', LotDep ,false);
InitForm;
end;

procedure TFNouveauLot.InitForm;
begin
NumeroLot.SetFocus;
NumeroLot.Text := '';
DateLot.Text := '01/01/1900';
Physique.Text := '0,00';
end;

procedure TFNouveauLot.bvaliderClick(Sender: TObject);
begin
if ControlFormOK then
   begin
   TobLotPutValeur;
   ValideOK := true
   end else ModalResult := 0;
end;

function TFNouveauLot.ControlFormOK : boolean;
begin
Result := false;
if NumeroLot.Text = '' then
   begin
   MsgBox.Execute(0,caption,'');
   NumeroLot.SetFocus;
   exit;
   end;

if NumLotExist then
   begin
   MsgBox.Execute(2,caption,'');
   NumeroLot.SetFocus;
   exit;
   end;

   //(not IsValidDate(DateLot.Text) and (DateLot.Text<>'')
//if StrToDate(DateLot.Text) > Date then
if StrToDate(DateLot.Text) > V_PGI.DateEntree then
   begin
   MsgBox.Execute(1,caption,'');
   DateLot.SetFocus;
   exit;
   end;

Result := true;
end;

function TFNouveauLot.NumLotExist : boolean;
var i_ind : integer;
begin
Result := false;
for i_ind:=0 to TOBDesLots.Detail.Count-1 do
    if TOBDesLots.Detail[i_ind].GetValue('GQL_NUMEROLOT')=NumeroLot.Text then
       begin
       Result := true;
       exit;
       end;
end;

procedure TFNouveauLot.TobLotPutValeur;
var TobLD : TOB;
begin
TobLD := TOB.Create('DISPOLOT',TOBDesLots,-1);
TobLD.AddChampSup('QTESAISIE',false);
TobLD.PutValue('GQL_ARTICLE', LotArt);
TobLD.PutValue('GQL_DEPOT', LotDep);
TobLD.PutValue('GQL_NUMEROLOT', NumeroLot.Text);
TobLD.PutValue('GQL_DATELOT', StrToDate(DateLot.Text));
TobLD.PutValue('QTESAISIE', Valeur(Physique.Text));
end;


procedure TFNouveauLot.PhysiqueExit(Sender: TObject);
begin
if Valeur(Physique.Text) = 0 then Physique.Text := '0,00'
else Physique.Text := StrS0(Valeur(Physique.Text));
end;

end.
