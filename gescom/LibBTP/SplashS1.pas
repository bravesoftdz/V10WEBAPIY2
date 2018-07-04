unit SplashS1;

interface

uses Forms,
     SysUtils,
     Windows,
     Messages,
     Classes,
     Graphics,
     Controls,
     StdCtrls,
     ExtCtrls,
     ComCtrls,
     HTB97,
     Dialogs;

type
  TSplash_S1 = class(TForm)
     Label1: TLabel;
     Animate1: TAnimate;
     Label2: TLabel;
  end;

var
  Splash_S1: TSplash_S1;

implementation

{$R *.DFM}

end.
