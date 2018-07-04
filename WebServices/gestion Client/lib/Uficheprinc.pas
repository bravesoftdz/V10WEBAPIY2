unit Uficheprinc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HSysMenu, HTB97, ExtCtrls, HPanel, StdCtrls, Hctrls,WSDL;

type
  TFPrinc = class(TForm)
    PBAS: THPanel;
    BQUIT: TToolbarButton97;
    BVALIDE: TToolbarButton97;
    HmTrad: THSystemMenu;
    USER: TEdit97;
    PASSW: TEdit97;
    LUSER: THLabel;
    LPASS: THLabel;
    MTEXT: THMemo;
    procedure BQUITClick(Sender: TObject);
    procedure BVALIDEClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FPrinc: TFPrinc;

implementation

{$R *.dfm}

procedure TFPrinc.BQUITClick(Sender: TObject);
begin
	close;
end;

procedure TFPrinc.BVALIDEClick(Sender: TObject);
var TWSLSENomade : IWSIISLSENomade;
		WSResp : TWSUser;
    WSIN : TWSUserParam;
begin
  WSIN := TWSUserParam.Create;
  WSIN.UniqueID := 'AAAAAAAAAAA';
  WSIN.User  := USER.Text;
  WSIN.Password   := PASSW.text;
  TWSLSENomade := GetIWSIISLSENomade ();
  TRY
    WSResp := TWSLSENomade.IsValideConnect(WSIN);
    MTEXT.Clear;
    MTEXT.Lines.Add('Code retour :' +InttoStr(WSREsp.CodeErreur));
    MTEXT.Lines.Add('ID :' +WSResp.InternalID);
    MTEXT.Lines.Add('user :' +WSResp.Codeuser );
    MTEXT.Lines.Add('SERVER :' +WSResp.DataBaseServer);
    MTEXT.Lines.Add('BATABASE :' +WSResp.Database);
    MTEXT.Lines.Add('message :' +WSResp.LibErreur);
  finally
    WSIN.Free;
  end;
//
end;

end.
