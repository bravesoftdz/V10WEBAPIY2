unit UlanceurUpdate;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,UMainUpdate, CoolTrayIcon;

type
  TForm3 = class(TForm)
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.FormShow(Sender: TObject);
var fMainUpdate: TfMainUpdate;
begin
  fMainUpdate := TfMainUpdate.Create(self);
  fMainUpdate.Show;
  fMainUpdate.LanceTraitement;
  fMainUpdate.Free;
  close;
end;

end.
