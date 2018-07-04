unit SaisTaux;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, ExtCtrls, Buttons, hmsgbox, Ent1, HEnt1, Saisutil,
  HSysMenu ;

type
  TFSaisTaux = class(TForm)
    BValider: TBitBtn;
    BFerme: TBitBtn;
    BAide: TBitBtn;
    HDevise: TLabel;
    HQUOTITE: TLabel;
    HDATECHANCEL: TLabel;
    HTAUXCHANCEL: TLabel;
    TDevise: TLabel;
    TQUOTITE: TLabel;
    TDATECHANCEL: TLabel;
    TTAUXCHANCEL: TLabel;
    Bevel1: TBevel;
    Label2: TLabel;
    NewTaux: THNumEdit;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    ISigneEuro: TImage;
    HTitre: TLabel;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
  public
    DEV : RDEVISE ;
    DateCpt : TDateTime ;
    GereEuro  : boolean ;
  end;

function SaisieNewTaux ( Var DEV : RDEVISE ; DateCpt : TDateTime ; GereEuro : boolean ) : boolean ;

implementation

{$R *.DFM}

Uses SaisTaux1 ;

function SaisieNewTauxOld ( Var DEV : RDEVISE ; DateCpt : TDateTime ; GereEuro : boolean ) : boolean ;
Var X  : TFSaisTaux ;
    ii : integer ;
BEGIN
Result:=False ;
if EstMonnaieIN(DEV.Code) then Exit ;
X:=TFSaisTaux.Create(Application) ;
 Try
  X.DEV:=DEV ; X.DateCpt:=DateCpt ; X.GereEuro:=GereEuro ;
  ii:=X.ShowModal ;
  if ((ii=mrOk) and (X.DEV.Taux>0)) then BEGIN DEV:=X.DEV ; Result:=True ; END ;
 Finally
  X.Free ;
 End ;
END ;

function SaisieNewTaux ( Var DEV : RDEVISE ; DateCpt : TDateTime ; GereEuro : boolean ) : boolean ;
BEGIN
If DateCpt<=V_PGI.DateDebutEuro Then Result:=SaisieNewTauxOld(DEV,DateCpt,GereEuro)
                                Else Result:=SaisieNewTaux2000(DEV,DateCpt) ;
END ;

procedure TFSaisTaux.BValiderClick(Sender: TObject);
begin
if NewTaux.Value<=0 then BEGIN HM.Execute(0,'','') ; Exit ; END ;
if ActiveControl=NewTaux then BValider.SetFocus ;
if ((GereEuro) and (DateCpt>=V_PGI.DateDebutEuro)) then DEV.Taux:=NewTaux.Value*V_PGI.TauxEuro
                                                else DEV.Taux:=NewTaux.Value ;
ModalResult:=mrOk ;
end;

procedure TFSaisTaux.FormShow(Sender: TObject);
begin
TDEVISE.Caption:=RechDom('ttDevisetoutes',DEV.Code,False) ;
TQUOTITE.Caption:=FloatToStr(DEV.Quotite) ;
TDATECHANCEL.Caption:=DateToStr(DEV.DateTaux) ;
if ((GereEuro) and (DateCpt>=V_PGI.DateDebutEuro)) then NewTaux.Value:=DEV.Taux/V_PGI.TauxEuro
                                                else NewTaux.Value:=DEV.Taux ;
TTAUXCHANCEL.Caption:=FloatToStr(NewTaux.Value) ;
if ((GereEuro) and (DateCpt>=V_PGI.DateDebutEuro)) then
   BEGIN
   iSigneEuro.Visible:=True ;
   HTitre.Caption:=HM.Mess[1]+' '+TDevise.Caption+' '+HM.Mess[2] ;
   HTitre.Visible:=True ;
   END ;
end;

end.
