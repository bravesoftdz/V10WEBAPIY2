unit BTWebIE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  OleCtrls, SHDocVw_TLB, StdCtrls, Hctrls, HTB97, ExtCtrls,Hpanel,UiUtil;

type
  TFBTWebIE = class(TForm)
    PBas: TPanel;
    Label1: TLabel;
    BValide: TToolbarButton97;
    BQUIT: TToolbarButton97;
    CBADRESSE: THValComboBox;
    Webfinestra: TWebBrowser_V1;
    precedent: TToolbarButton97;
    suivant: TToolbarButton97;
    procedure BValideClick(Sender: TObject);
    procedure BQUITClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure precedentClick(Sender: TObject);
    procedure suivantClick(Sender: TObject);
  private
    { Déclarations privées }
    ExistPrevPages,ExistnextPages : boolean;
  public
    { Déclarations publiques }
  end;

procedure OuvertureWebBrowser;

implementation

{$R *.DFM}

procedure OuvertureWebBrowser;
var  XX: TFBTWebIE;
		 PP : THPanel;
Begin

  PP := FindInsidePanel;
	XX := TFBTWebIe.Create (application);

  if PP = nil then
     begin
     try
       XX.ShowModal;
     FINALLY
       XX.Free;
     END;
     end
  else
     begin
     XX.ClientWidth := PP.ClientWidth;
     XX.ClientHeight  := PP.ClientHeight;
     InitInside(XX, PP);
     XX.Show;
     end;
     
end;

procedure TFBTWebIE.BValideClick(Sender: TObject);
var flags : OleVariant;
begin
//
flags := 0;
WebFinestra.Navigate (WideString(CBADRESSE.Text),Flags,flags,flags,flags);
end;

procedure TFBTWebIE.BQUITClick(Sender: TObject);
begin
	WebFInestra.Stop;
	Close;
end;

procedure TFBTWebIE.FormShow(Sender: TObject);
begin
  if IsInside(self) then BQUIT.Visible := false;

	BValide.Click;
end;

procedure TFBTWebIE.precedentClick(Sender: TObject);
begin
	TRY
		WebFInestra.GoBack ;
  EXCEPT
  END;
end;

procedure TFBTWebIE.suivantClick(Sender: TObject);
begin
	TRY
		WebFInestra.GoForward;
  EXCEPT
  END;
end;

end.
