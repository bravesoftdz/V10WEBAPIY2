unit URapport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, ExtCtrls, StdCtrls, TntStdCtrls, Hctrls;

type
  TFBTRapport = class(TForm)
    PBAS: TPanel;
    BClose: TToolbarButton97;
    MemoRapport: THMemo;
    procedure BCloseClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    Property Memo : THMemo read MemoRapport write MemoRapport;
  end;

var
  FBTRapport: TFBTRapport;

implementation

{$R *.dfm}

procedure TFBTRapport.BCloseClick(Sender: TObject);
begin
  close;
end;

end.
