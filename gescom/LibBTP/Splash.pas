unit Splash;

interface

uses
  SysUtils, Windows, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, ComCtrls;

type
  TFsplashScreen = class(TForm)
    Label1: TLabel;
    Animate1: TAnimate;
    Label2: TLabel;
    TRestaure: TLabel;
    ComboFile: TComboBox;
    LblEncours: TLabel;
  end;

var
  FsplashScreen: TFsplashScreen;

implementation

{$R *.DFM}


end.
