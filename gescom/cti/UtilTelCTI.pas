unit UtilTelCTI;

interface
uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  OleCtrls,
  Buttons,
  Mask,
  AGLInitGC,
  ParamSoc,
  CBPPath,
{$IFDEF EAGLCLIENT}
  eFiche,
  Maineagl,
{$ELSE}
  Fiche,
  FE_main,
{$ENDIF}
  AppelsUtil,
	PCb_TLB,
  UIUtil,
  StdCtrls,
  HMsgBox,
  HTB97,
  HEnt1 ;

Const
  Avance : integer = 1;
  Simple : integer = 0;
  ModeAppelSortant : integer = 1 ;
  ModeAppelEntrant : integer = 2 ;

type
	TUtilTelCTI = class
	private
  	fform : Tform;
  	fControlPCB : TCOntrolPCB;
    fLogCTI : TextFile;
    fNumtel : string;
    LOG: TMemo;
    fBtnAppeler : TToolbarButton97;
    fBtnAttente : TToolbarButton97;
    CtiErreurConnexion: Boolean;

    procedure SetNumTel(const Value: string);
  	procedure AssigneEvenement;
    procedure ControlPCB1AppelAbouti(Sender: TObject; RefAppel: Integer;
      const appele: WideString);
    procedure ControlPCB1ErreurConnexion(Sender: Tobject;
      RefAppel: Integer);
    procedure ControlPCB1EtatConnectionPCB(Sender: TObject;
      EtatPCB: WordBool);

  public
  	constructor Create (AOwner : TForm);
    destructor destroy; override;
    //
    property Numtel : string read fNumTel write SetNumTel;
    Property BtnAppeler : TToolbarButton97 read fBtnAppeler write fBtnAppeler;
    Property BtnAttente : TToolbarButton97 read fBtnAttente write fBtnAttente;
    //
    procedure AddLog (StLog : String);
    procedure AppeleNumero (Numero : string);
	end;

implementation

{ TUtilTelCTI }

procedure TUtilTelCTI.AddLog(StLog: String);
begin

	Log.Lines.add(StLog);
	writeln (FLogCti,StLog);

end;

procedure TUtilTelCTI.AppeleNumero(Numero: string);
begin
  NumTel := Numero;
	if NumTel <> '' then
  begin
  	if GetParamSoc('SO_RTCTILIGNEEXT') <> '' then
    		NumTel:=GetParamSoc('SO_RTCTILIGNEEXT')+NumTel;
    AddLog('--> Appel du no '+NumTel);
    AddLog(' ') ;
    fControlPCB.AppelUnClient(NumTel) ;
    if CtiErreurConnexion then
    begin
      fNumTel := '';
    end;
  end;

end;

procedure TUtilTelCTI.AssigneEvenement;
begin
  fControlPCB.OnCstaFailed := ControlPCB1ErreurConnexion;
  fControlPCB.OnCstaDelivered := ControlPCB1AppelAbouti;
  fControlPCB.OnEtatConnectionPCB := ControlPCB1EtatConnectionPCB;
end;


procedure TUtilTelCTI.ControlPCB1EtatConnectionPCB(Sender: TObject;
  EtatPCB: WordBool);
begin
AddLog('--> Connexion Téléphonie établie ');
AddLog(' ') ;
end;

procedure TUtilTelCTI.ControlPCB1ErreurConnexion(Sender: Tobject;  RefAppel: Integer);
begin
  AddLog('--> Erreur sur la connexion');
  AddLog('    Id Appel   : '+IntToStr(RefAppel));
  CtiErreurConnexion:=True;
end;

procedure TUtilTelCTI.ControlPCB1AppelAbouti(Sender: TObject; RefAppel: Integer;
 const appele: WideString);
begin
  AddLog('--> Appel Sortant Abouti');
  AddLog('    Id Appel   : '+IntToStr(RefAppel));
  AddLog('    Appelé   : '+appele);
end;

constructor TUtilTelCTI.Create(AOwner: TForm);
Var JournalCti : String;
begin
	fform := TForm(AOwner);

  JournalCti := IncludeTrailingBackSlash(TCBPPath.GetLocalAppData) + 'JournalCti.log';
  //
	AssignFile (FLogCti, JournalCti);
	Rewrite (FLogCti);
  //
  fControlPCB := TcontrolPCB.create (AOwner);
  fControlPCB.Parent := AOwner;
  fControlPCB.StartMode (Avance);
  CtiErreurConnexion:=False;
  AssigneEvenement;
end;

destructor TUtilTelCTI.destroy;
begin
  inherited;
	CloseFile (FLogCti);
	fControlPCB.quit;
  fControlPCB.Free;
end;

procedure TUtilTelCTI.SetNumTel(const Value: string);
begin
  fNumTel := FormatNoTel (value);
end;

end.
