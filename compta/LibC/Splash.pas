unit Splash;

interface

uses WinTypes,WinProcs,Classes,Graphics,Forms,Controls,Buttons,StdCtrls,ExtCtrls,
     HEnt1, HSysMenu, Ent1, HTB97, HPanel ;

type
  TFSplashScreen = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    BMPPanel1: THPanel;
    FVersion: TLabel;
    RCopyright: TLabel;
    HLogo: TToolbarButton97;
    iHalley: TToolbarButton97;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private    { Private declarations }
  public    { Public declarations }
  end;


Var FSplashScreen : TFSplashScreen ;
implementation

{$R *.DFM}

procedure TFSplashScreen.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Screen.Cursor:=crDefault ;
end;

procedure TFSplashScreen.FormShow(Sender: TObject);
begin
RCopyright.Caption := Copyright ;
FVersion.Caption := FVersion.Caption + ' ' + V_PGI.NumVersion ;
end;

end.
