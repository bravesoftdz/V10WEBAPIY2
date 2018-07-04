unit CegidIEU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  WinSock, StdCtrls,
{$IFDEF SANSOCX}
  MenuOLX,
{$ELSE}
  eAGLX_TLB,
{$ENDIF}

  ExtCtrls;

type
  TForm1 = class(TForm)
    LePanel: TPanel;
    bConnecte: TButton;
    HostN: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Module: TComboBox;
    procedure bConnecteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private    { Déclarations privées }
    eAGL: TeAGL;
  public     { Déclarations publiques }
  end;

var
  WSAData: TWSAData;
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.bConnecteClick(Sender: TObject);
begin
Caption:='CEGID eAGL sur '+HostN.text ;
//obsolete if FMultiServeur.Checked then Caption:=Caption+' Multi-process' ;
LePanel.Visible:=FALSE ;
Height:=600 ; Width:=800 ; Top:=10 ; left:=10 ;
eAGL:=TeAGL.Create(Self) ; eAGL.Parent:=Self ; eAGL.Align:=alClient ;
With eAGL do
   BEGIN
   Host:=HostN.text ;
   Domaine:=Module.text ;
   //Obsolete   if FMultiServeur.Checked then MultiProcess:='X' else MultiProcess:='-' ;
{$IFDEF SANSOCX}
   Show ;
{$ENDIF}
   END ;
Caption:=Caption+' - module '+eAGL.NomModule ;
LePanel.enabled:=FALSE ;
end ;


procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if eAGL<>Nil then eAGL.Free ;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
eAGL:=Nil ;
end;

procedure TForm1.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if Assigned(eagl) then eAGL.OnKeyDown(Sender,Key,Shift) ;
if key=vk_escape then Close ;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if eAGL<>Nil then CanClose:=(eAGL.CanQuit=1) ;
end;

Initialization
  WSAStartup($0101, WSAData);
Finalization
  WSACleanup;
end.
