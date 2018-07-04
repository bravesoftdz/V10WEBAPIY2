unit MPEscompte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTOB,
  StdCtrls, Mask, Hctrls, HSysMenu ;

type
  TFMPEscompte = class(TForm)
    HMTrad: THSystemMenu;
    ESCOMPTABLE: TCheckBox;
    COMPTEHT: THCritMaskEdit;
    COMPTETVA: THCritMaskEdit;
    TAUXTVA: THCritMaskEdit;
    E_NUMEROPIECE: THCritMaskEdit;
    Button1: TButton;
    TAUXESC: THCritMaskEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
  public
    TOBE : TOB ;
  end;



Procedure SaisieMPEscompte ( TOBE : TOB ) ;

implementation

{$R *.DFM}

Procedure SaisieMPEscompte ( TOBE : TOB ) ;
Var X : TFMPEscompte ;
BEGIN
X:=TFMPEscompte.Create(Application) ;
X.TOBE:=TOBE ;
 Try
   X.ShowModal ;
 Finally
   X.Free ;
 End ;
END ;


procedure TFMPEscompte.FormShow(Sender: TObject);
begin
TOBE.PutEcran(Self) ;
end;

procedure TFMPEscompte.Button1Click(Sender: TObject);
begin
TOBE.GetEcran(Self) ;
end;

end.
