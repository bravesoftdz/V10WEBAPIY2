unit Splash35;

interface

uses WinTypes,WinProcs,Classes,Graphics,Forms,Controls,Buttons,StdCtrls,ExtCtrls,
     HEnt1, HSysMenu, HTB97, HPanel ;

type
  TSplashScreen = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    BMPPanel1: THPanel;
    IVersion: TLabel;
    ICopyright: TLabel;
    HLogo: TToolbarButton97;
    iHalley: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private    { Private declarations }
  public    { Public declarations }
  end;


Var SplashScreen : TSplashScreen ;
implementation

{$R *.DFM}

procedure TSplashScreen.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Screen.Cursor:=crDefault ;
end;

procedure TSplashScreen.FormShow(Sender: TObject);
begin
iHalley.Caption:=NomHalley ;
//iHalley.Caption:='Compta S5' ; 
ICopyright.Caption := Copyright ;
IVersion.Caption := IVersion.Caption + ' ' + V_PGI.NumVersion ;
end;

end.
