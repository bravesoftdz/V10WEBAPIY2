unit ASCIIV;
{
			-Fichier-------------------------------------------
			|
			|FIgnored
	 PtJ[0]	|--------
			|
 [iLineBuf]	|-Buffer------------------------------------  0
			|aLine[NBLBUF][<=LngLineMax]
			|
			|
			|				|HorzScrollBar.Position
			|	 VertPosWin -Fenêtre-----------------
			|				|	nbColWin			|
			|				|nbRowWin				|
			|				|						|     NBLMAX
			|				|			CurColWin	|
	CurRow	|				+ CurRowWin |-Curseur	|
  CurLine[]	|				|						|
			|				|<=NBLMAX				|
			|				------------+------------
			|						CurCol
			|
			|-------------------------------------------  NBLBUF
			|
  LineCount	---------------------------------------------------
}

{$B-,Q-,R-,W-}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  ExtCtrls, StdCtrls, Clipbrd;

const
	foFromTop =	$00000000;
	foFromCaret =	$00000001;
	foIgnoreCase =	$00000002;
	// TFindOption = ();	TFindOptions = set of TFindOption; ??

	NBLMAX = 50;	// Nombre maxi de lignes affichables
	NBLBUF = NBLMAX*2;	// Nombre de lignes dans buffer mémoire
	LNGLMAX = 4096;	// Longueur maxi d'une ligne

type
  TASCIILine = array[0..LNGLMAX] of Char;

	{ ASCIIView }
  TGetLineEvent = procedure (Sender: TObject; var ALine: TASCIILine; var AColor: TColor; var AGlyph: Integer; ALineN: Integer) of object;

  TAJ = array[0..65535] of Longint; // Tableau dynamique Jalons

  TASCIIView = class(TCustomPanel)
  private
	StdFont: TFont;
	HorzScrollBar: TScrollBar;
	VertScrollBar: TScrollBar;
	bMouseDown: Boolean;
	ASCIIFile: Integer;	// Handle fichier, -1 si pas ouvert
///	ThreadID: DWORD;
///	ThreadIndex: THandle;
	Timer: TTimer;
	Jalon: Integer;	// Index courant pour l'indexation
	PtJ: ^TAJ;		// Repérage de débuts de lignes dans le fichier
	NbJ: Integer;	// Nombre de jalons alloués
	iFileSize: Longint;	// Taille du fichier sans Ignored ni dernière page pour calcul position verticale
	Bord: Integer;		// Epaisseur de la bordure
	HCar: Integer;		// Hauteur d'un caractère
	WCar: Integer;		// Largeur d'un caractère
	nbRowWin: Integer;	// Nombre de lignes d'affichage
	nbColWin: Integer;	// Nombre de colonnes d'affichage (peut déborder à droite)
	LngLineMax: Integer;	// Longueur maxi d'une ligne chargée, au moins nbColWin, 0 si pas de texte
	aLine: array[0..NBLBUF] of string;	// Lignes pour buffer, #0 = fin de fichier
	aColor: array[0..NBLBUF-1] of TColor;	// Couleurs pour Lignes
	aGlyph: array[0..NBLBUF-1] of Integer;	// Codes glyphs pour Lignes
	iLineBuf: Integer;	// Index dans Jalons de la 1ere ligne du buffer dans le fichier
	VertPosWin: Integer;// Offset vertical de la fenêtre dans le buffer
	CurRowWin: Integer;	// Positions par rapport à la fenêtre, en caractères
	CurColWin: Integer;
        FSelColor: TColor;      // Couleur de fond pour texte selectionné
	FIgnored: Integer;	// Nombre de lignes ignorées au début du fichier
	FLineCount: Integer;// Nombre de lignes chargées, hors Ignored (32767 maxi ?), 0 si pas encore connu
	FLCol: Integer;		// Positions colonnes marquées, -1 si pas de sélection ?
	FRCol: Integer;		//
	FLSel: Integer;		// Positions par rapport à la ligne, -1 si pas de sélection
	FRSel: Integer;		//
        FTopLine: Integer ;
	FMargin: Integer;	// Largeur de la marge (avec gouttière de marquage)
	FFileName: TFileName;
	FOnGetLine: TGetLineEvent;
	FOnCaretMove: TNotifyEvent;
	procedure EnableTimer;
	procedure IndexExec(ASender: TObject);
	procedure WMSize(var Msg: TWMSize); message WM_SIZE;
	procedure WMLButtonDown(var Msg: TWMLButtonDown); message WM_LBUTTONDOWN;
	procedure WMLButtonUp(var Msg: TWMLButtonUp); message WM_LBUTTONUP;
	procedure WMMouseMove(var Msg: TWMMouseMove); message WM_MOUSEMOVE;
	procedure WMGetDlgCode(var Msg: TWMGetDlgCode); message WM_GETDLGCODE;
	procedure CMEnter(var Msg: TCMGotFocus); message CM_ENTER;
	procedure CMExit(var Msg: TCMExit); message CM_EXIT;
//	procedure WMCopy(var Msg: TMessage); message WM_COPY;
	function GetJ(AI: Integer) : Longint;
	procedure SetJ(AI: Integer; AJ: Longint);
	procedure ClearJ(AInit: Longint);
	procedure CaretMove;
	function GetVersion: String;
	procedure SetVersion(AV: String);
	procedure SetFileName(AV: TFileName);
	function GetBevelInner: TPanelBevel;
	procedure SetBevelInner(AV: TPanelBevel);
	function GetBevelOuter: TPanelBevel;
	procedure SetBevelOuter(AV: TPanelBevel);
	procedure SetSelColor(AV: TColor);
	procedure SetCurCol(AV: Integer);
	function GetCurCol : Integer;
	procedure SetCurRow(AV: Integer);
	function GetCurRow : Integer;
	function GetCurLine : String;
	function GetLineCount : Integer;
	procedure SetLCol(AV: Integer);
	procedure SetRCol(AV: Integer);
	procedure SetLSel(AV: Integer);
	procedure SetRSel(AV: Integer);
	procedure SetTopLine(AV: Integer);
	procedure SetIgnored(AV: Integer);
	function GetMargin : Boolean;
	procedure SetMargin(AV: Boolean);
	procedure ScrollHChange(Sender: TObject);
	procedure ScrollV(Sender: TObject; ScrollCode: TScrollCode; var ScrollPos: Integer);
	procedure ChangeBord;
	procedure AjustVScroll;
	procedure AjustHScroll;
	function VScrollPos : Longint;
	procedure BufUp;
	procedure BufDown;
	procedure NextLine;
	procedure PrevLine;
	procedure NextPage;
	procedure PrevPage;
	procedure ReadLines(ABloc: Integer; AIndex: Longint);
	procedure UpdateTextSel;
	procedure UpdateText;
	procedure UpdateBar;
	procedure Size;
        procedure InverseColor(Rc : TRect; clBack : TColor);
  protected
///	Running: Integer; // <> 0 si le thread est lancé, 4 pour l'arréter, 2 si accés fichier
	//property LngLineMax: Integer read FLngLineMax write SetLngLineMax stored False;
	procedure Paint; override;
	procedure KeyDown(var Key: Word; Shift: TShiftState); override;
//	procedure KeyPress(var Key: Char); override;
  public
	property SelColor: TColor read FSelColor write SetSelColor;
	property CurCol: Integer read GetCurCol write SetCurCol;
	property CurLine: String read GetCurLine;
	property CurRow: Integer read GetCurRow write SetCurRow;
	property Ignored: Integer read FIgnored write SetIgnored;
	property LeftCol: Integer read FLCol write SetLCol;
	property LeftSel: Integer read FLSel write SetLSel;
	property LineCount: Integer read GetLineCount;
	property RightCol: Integer read FRCol write SetRCol;
	property RightSel: Integer read FRSel write SetRSel;
	property TopLine: Integer read FTopLine write SetTopLine ;
        property LineMax: Integer read LngLineMax write LngLineMax; // Longueu max ligne
	procedure UpdateLines;
	function Find(AStr: string; AOption: Integer; Ajout: integer=0; Suivant : integer=0): Boolean;
	constructor Create(AOwner : TComponent); override;
	destructor Destroy; override;
	procedure Loaded; override;
        function SelectText : string;
        procedure CloseMemo;
  published
	property Align;
	property BevelInner: TPanelBevel read GetBevelInner write SetBevelInner;
	property BevelOuter: TPanelBevel read GetBevelOuter write SetBevelOuter;
//	property BevelWidth;
//	property BorderWidth;
//	property BorderStyle;
//	property Caption;
	property Color;	// Couleur de fond panel
	property Ctl3D;
	property DragCursor;
	property DragMode;
	property Enabled;
	property Font; // SetFont; HCar := 0; Size; ??
//	property FullRepaint;
//	property Locked;
	property Margin: Boolean read GetMargin write SetMargin;
	property ParentColor;
	property ParentCtl3D;
//	property ParentFont;
	property ParentShowHint;
	property PopupMenu;
	property ShowHint;
	property TabOrder;
	property TabStop;
	property Version: String read GetVersion write SetVersion stored False;
	property Visible;
	property OnClick;
	property OnDblClick;
	property OnDragDrop;
	property OnDragOver;
	property OnEndDrag;
	property OnEnter;
	property OnExit;
	property OnGetLine: TGetLineEvent read FOnGetLine write FOnGetLine;
	property OnMouseDown;
	property OnMouseMove;
	property OnMouseUp;
	property OnResize;
	property OnCaretMove: TNotifyEvent read FOnCaretMove write FOnCaretMove;
	property OnStartDrag;
	property FileName: TFileName read FFileName write SetFileName; // Nécessite l'initialisation d'autres propriétés
  end;

procedure Register;


implementation


const
	HBAR = 18;	// Hauteur de la barre graduée : selon hauteur texte ??
	HTICK = 5;	// Hauteur d'un trait
	MARGE = 1;	// Marge gauche et bas pour la zone texte
	GOUTTIERE = 16;	// Marge gauche pour marquage


	{--Timer--------------}
procedure TASCIIView.EnableTimer;
begin	{ Premier lancement }
	iFileSize := FileSeek(ASCIIFile, - GetJ(0), soFromEnd); // Sera ajusté quand fin rencontrée
// if = HFILE_ERROR ??
	FLineCount := 0;
	Jalon := 0;
	Timer.Enabled := True;
	iLineBuf := -1;	// Provoque chargement par SetCurRow
	CurRow := 0; // iLineBuf,VertPosWin,CurRowWin := 0, ReadLines, UpdateTextSel
end;


///function ThreadIndexExec(AASCIIView: TASCIIView): Integer; stdcall;
procedure TASCIIView.IndexExec(ASender: TObject);
var	TmpBuf: TASCIILine;
	I, L, N: Integer;
	Offset: Longint;
begin	 // Compte lignes sans charger aLine ni appeler GetLine
//Showmessage('Thread lancé !'); // ???
	// Faire ClearJ avant
	Offset := GetJ(Jalon);
///	iFileSize := FileSeek(ASCIIFile, - Offset, soFromEnd); // Sera ajusté quand fin rencontrée
///	FLineCount := 0;
///	J := 0;
///	repeat
		for I := 0 to NBLMAX-1 do
		  try
			 // Lit une ligne // Eviter repositionnements -> read 1 caractère ??
			// if Running <> 4 then begin Running := 2 ... ?
///			Inc(Running); // EnterCriticalSection !
			FileSeek(ASCIIFile, Offset, soFromBeginning);
			L := FileRead(ASCIIFile, TmpBuf, LNGLMAX);
///			Dec(Running);

			if (L = 0) or (TmpBuf[0] = #26) then	// Erreur ou fin de fichier ???
			begin
				L := 0;
				Break;
			end;

			for N := 0 to L-1 do	// Pos( Scan ??
			begin	 // Recherche fin de ligne $0D0A, $0D ou $0A
				if TmpBuf[N] = #$0D then
				begin
					if TmpBuf[N+1] = #$0A then	Inc(Offset);
					break;
				end;
				if TmpBuf[N] = #$0A then
					break;
			end;
			Inc(Offset, N+1); // FileSeek(ASCIIFile, N+1-L, soFromCurrent);
			Inc(FLineCount);
		  except
//			Showmessage('Probleme read'); // ???
///			Running := 0;
			Timer.Enabled := False;
			Raise;
		  end;

		Inc(Jalon);
		SetJ(Jalon, Offset); //SetJ(Jalon, FileSeek(ASCIIFile, 0, soFromCurrent));	// Position après la dernière ligne

		AjustVScroll;

///	until (L = 0) or (Running > 2); //Terminated;

	if L = 0 then //Terminated
		Timer.Enabled := False;
///  end;
///Showmessage('Thread Terminé'); // ???
end;


	{--Jalons--------------}
function TASCIIView.GetJ(AI: Integer) : Longint;
begin
	if AI > 0 then	// [0] Toujours disponible
	begin
		repeat
			if AI < NbJ then	// Nb incrémenté par Thread !
			begin	// Index disponible
				Result := PtJ[AI];
				if Result <> 0 then	// Valeur disponible
					Exit;
			end;
			 // Valeur pas disponible
///			Sleep(0); // and Running = 0 ?
			IndexExec(Nil); // Application.HandleMessage; ?
///		until Running = 0;	// Thread terminé
		until not Timer.Enabled;

		if AI >= NbJ then // Si valeur > maxi rendre maxi !
		begin
			AI := NbJ;
			repeat
				Dec(AI);
			until PtJ[AI] > 0;
		end
	end;

	Result := PtJ[AI];
end;


procedure TASCIIView.SetJ(AI: Integer; AJ: Longint);
begin
// CriticalSection ??
	if AI >= NbJ then
	begin	// Agrandi le tableau
		ReAllocMem(PtJ, (AI + 8)*SizeOf(Longint));
		// if PtJ = Nil then Exit ??
		ZeroMemory(@PtJ[NbJ], (AI + 8 - NbJ)*SizeOf(Longint)); // FillChar(PtJ[NbJ], (AI + 8 - NbJ)*SizeOf(Longint), 0);
		NbJ := AI + 8;
	end;
	PtJ[AI] := AJ;
end;

	{ Efface les jalons suite à nouveau fichier ou lignes à ignorer }
procedure TASCIIView.ClearJ(AInit: Longint);
begin
	if AInit = 0 then
	begin
		NbJ := 8;
		ReAllocMem(PtJ, 8*SizeOf(Longint));
		// PtJ = Nil pas possible !?
	end;
	ZeroMemory(PtJ, NbJ*SizeOf(Longint)); // FillChar(PtJ^, 8*SizeOf(Longint), 0);
	PtJ[0] := AInit; // Offset de départ
end;


	{--ASCIIView--------------}
function TASCIIView.GetVersion: String;
begin
	Result := '1.0b';
end;

procedure TASCIIView.SetVersion(AV: String);
begin	// Sans effet
end;


procedure TASCIIView.SetFileName(AV: TFileName);
var	I: Integer;
begin
	FFileName := AV;
	if (HCar = 0) or (csLoading in ComponentState) then
		Exit;	// HCar, VCar, nbRowWin, nbColWin nécessaires : WMSize

{	if Running <> 0 then // Si Thread en route (not Suspended)
	begin
		ThreadIndex.Terminate;
		ThreadIndex.WaitFor;
	end;
}
	if ASCIIFile >= 0 then
	begin
///		if Running <> 0 then
///		begin
///			Running := 4;
///			Sleep(1); // Attend la fin ???
			Timer.Enabled := False;
///		end;
		FileClose(ASCIIFile);
		ASCIIFile := -1;
	end;

	FIgnored := 0;
	FLCol := -1;
	FLSel := -1;
	ClearJ(0);
	if AV <> '' then
	begin
		ASCIIFile := FileOpen(AV, fmOpenRead or fmShareDenyWrite);
		if ASCIIFile >= 0 then
		begin	// try ??
			VertScrollBar.Max := High(Smallint);
			LngLineMax := nbColWin;
			if csDesigning in ComponentState then
			begin
				// iLineBuf,VertPosWin,CurRowWin := 0
				ReadLines(0, 0);
				UpdateText;
			end
			else
			begin
///				ThreadIndex := CreateThread(Nil, 0, @ThreadIndexExec, Self, 0, ThreadID);
				EnableTimer;
// ?			if not (csLoading in ComponentState) then // Diffère lecture des données calculées par Thread ??
///					ThreadRun;
			end;
			//HorzScrollBar.Enabled := True;
			//HorzScrollBar.Max := LngLineMax - nbColWin;
			//VertScrollBar.Enabled := True;
		end
		else
		begin
			if csDesigning in ComponentState then
 //ajout me				ShowMsg(6, [AV]);
			FileName := ''; // -> vide
		end;
	end
	else
	begin // Retirer le focus ?
		VertPosWin := 0;
		FLineCount := 0;
		LngLineMax := 0; // iFileSize := 0 ?
		for I:=0 to NBLBUF-1 do
			aLine[I] := #0;
		HorzScrollBar.Max := 0;	//HorzScrollBar.Enabled := False;
		VertScrollBar.Max := 0; //VertScrollBar.Enabled := False;
		AjustVScroll; // ?
		AjustHScroll;
		UpdateTextSel;
	end;

	HorzScrollBar.Position := 0;
end;


function TASCIIView.GetBevelInner: TPanelBevel;
begin
	Result := inherited BevelInner;
end;

procedure TASCIIView.SetBevelInner(AV: TPanelBevel);
begin
	inherited BevelInner := AV;
	ChangeBord;
end;


function TASCIIView.GetBevelOuter: TPanelBevel;
begin
	Result := inherited BevelOuter;
end;

procedure TASCIIView.SetBevelOuter(AV: TPanelBevel);
begin
	inherited BevelOuter := AV;
	ChangeBord;
end;

procedure TASCIIView.SetSelColor(AV: TColor);
begin
	FSelColor := AV;
end;


procedure TASCIIView.SetCurCol(AV: Integer);
var	bUpdateText: Boolean;
	N: Integer;
begin
	//if AV > LngLineMax then  AV := fin de ligne ?
	if AV < 0 then	AV := 0;
	if (AV <> CurCol) and (LngLineMax <> 0) then
	begin
		bUpdateText := False;
		N := HorzScrollBar.Position;
		if (AV < N) or (AV >= N+nbColWin) then
		begin	// Scroll
			if AV >= LngLineMax then
			begin	// Au bout
				AV := LngLineMax; // -1 ??
				N := HorzScrollBar.Max;
			end
			else
			begin
				if bMouseDown then
				begin	// Souris en dehors -> décalage d'un caractère
					if AV < N then
						Dec(N)
					else
						Inc(N);
				end
				else
					N := (AV div nbColWin) * nbColWin;
			end;
			HorzScrollBar.Position := N; // UpdateText
//			if N >= HorzScrollBar.Max then	// Au bout
//				N := HorzScrollBar.Max-1;	// Si scroll en sélection ??
			N := HorzScrollBar.Position; //
//			bUpdateText := True;
		end;
		CurColWin := AV - N;
		if bMouseDown then
			UpdateText	// Marque la nouvelle sélection
		else if (FLSel >= 0) or bUpdateText then
			UpdateTextSel;
		CaretMove;
	end
end;


function TASCIIView.GetCurCol : Integer;
begin
	Result := HorzScrollBar.Position + CurColWin;
end;


	{ Ligne du curseur dans le fichier (à partir de 0) }
procedure TASCIIView.SetCurRow(AV: Integer);
var	LineBuf: Integer;
begin	// Positionne la ligne courante (0 = première ligne non Ignored)
	if (AV >= 0) and (LngLineMax <> 0) then
	try
		Screen.Cursor := crHourGlass;
		if FLineCount <> 0 then	// Pas de problème pour début !
		begin
			// Tant que Thread en cours attend valeur disponible
			while Timer.Enabled and (AV >= FLineCount) do
				IndexExec(Nil); // Application.HandleMessage; ?

			if AV >= FLineCount then
				AV := FLineCount-1;
		end;
		LineBuf := AV div NBLMAX; // Index du bloc contenant ligne AV
		VertPosWin := AV - LineBuf * NBLMAX;
//exit; // AVOIR
		if (LineBuf > 0) and (VertPosWin + nbRowWin < NBLMAX) then
		begin	// Lit plutôt au dessus au cas où fin de fichier atteinte
			Dec(LineBuf);
			Inc(VertPosWin, NBLMAX);
		end;
		if LineBuf <> iLineBuf then
			ReadLines(2, LineBuf); // Peut modifier VertPosWin
		CurRowWin := AV - (iLineBuf*NBLMAX + VertPosWin);
		VertScrollBar.Position := VScrollPos;
	finally
		UpdateTextSel;
		CaretMove;
		Screen.Cursor := crDefault;
	end
end;


function TASCIIView.GetCurRow : Integer;
begin
	Result := iLineBuf*NBLMAX + VertPosWin + CurRowWin;
end;


function TASCIIView.GetCurLine : String;
begin
	if LngLineMax <> 0 then	// Il y a un texte
		Result := aLine[VertPosWin + CurRowWin]
	else
		Result := '';
end;


function TASCIIView.GetLineCount : Integer;
begin
	//if pas (indexation suite à SetIgnored) then ??  optimisation
///	while Running <> 0 do
	while Timer.Enabled do
///		Sleep(0);	// Attend Thread
		Application.ProcessMessages;
	Result := FLineCount;
end;


procedure TASCIIView.SetLCol(AV: Integer);
begin
	if (AV <> FLCol) and (LngLineMax <> 0) then	// and >= -1 (and < LngLineMax) ??
	begin
		FLCol := AV;
		if FLCol <= FRCol then
		begin
			UpdateBar; UpdateText; // Invalidate; ??
		end
	end
end;

//VT 26/2/2003
procedure TASCIIView.SetTopLine(AV: Integer);
begin
	if (AV <> FLCol) then	// and >= -1 (and < LngLineMax) ??
	begin
		FTopLine := AV;
		UpdateText; // Invalidate; ??
	end
end;



procedure TASCIIView.SetRCol(AV: Integer);
begin
	if (AV <> FRCol) and (LngLineMax <> 0) then	// (and >= 0) and < LngLineMax ??
	begin
		FRCol := AV;
		if FRCol >= FLCol then
		begin
			UpdateBar; UpdateText; // Invalidate; ??
		end
	end
end;


procedure TASCIIView.SetLSel(AV: Integer);
begin
	if (AV <> FLSel) and (LngLineMax <> 0) then	// and >= -1 (and < LngLineMax) ??
	begin
		FLSel := AV;
		if FLSel <= FRSel then
			UpdateText;
		//SelChange;
	end
end;


procedure TASCIIView.SetRSel(AV: Integer);
begin
	if (AV <> FRSel) and (LngLineMax <> 0) then	// (and >= 0) and < LngLineMax ??
	begin
		FRSel := AV;
		if FRSel >= FLSel then
			UpdateText;
		//SelChange;
	end
end;


procedure TASCIIView.SetIgnored(AV: Integer);
var	I, L: Integer;
	Offset: Longint;
	TmpBuf: Char;
	bDouble: Boolean;
begin
	if (AV <> FIgnored) and (LngLineMax <> 0) then
//	and AV < nbLines then ??
	begin
		Screen.Cursor := crHourGlass;
		 // Détermine offset de ligne FIgnored
		if AV < FIgnored then
		begin
			I := AV;
			Offset := 0;	// Repartir du début
		end
		else
		begin
			I := (AV-FIgnored) div NBLMAX;
			Offset := GetJ(I); // Index du bloc contenant ligne AV
			I := (AV-FIgnored) - I * NBLMAX; // Nombre de lignes restantes
		end;

{		if Running <> 0 then	// Ou while Running Sleep ??
		begin
			ThreadIndex.Terminate;
			ThreadIndex.WaitFor;
		end; }
//		if not Running then	// Optimisation
//			Inc(FLineCount, FIgnored - AV); // Si LineCount connu avant, calcule nouvelle valeur tout de suite !
//			et ne pas recalculer dans Thread ??
///		Running := 4;
///		Sleep(1); // Attend la fin ???
		Timer.Enabled := False;

		FIgnored := AV;

		bDouble := False;
		FileSeek(ASCIIFile, Offset, soFromBeginning);
		while I > 0 do
		begin
		  try
			L := FileRead(ASCIIFile, TmpBuf, 1);
			Inc(Offset);
			if L = 0 then	// Erreur ou fin de fichier ???
				// Dec(FIgnored, I); ??
				break;

			if (TmpBuf = #$0A) or bDouble then
			begin
				Dec(I);
				if (I = 0) and (TmpBuf <> #$0A) then
					Dec(Offset); // Recule
			end;
			bDouble := (TmpBuf = #$0D);

		  except
//			Showmessage('Probleme read'); // ??
			Raise;
		  end;

		end;

		ClearJ(Offset); // Ne pas faire Realloc même si Offset = 0 ??
		//	if not (csDesigning in ComponentState) then
///		ThreadIndex := CreateThread(Nil, 0, @ThreadIndexExec, Self, 0, ThreadID);
//		ThreadIndex := TThreadIndex.Create(Self);
//		ThreadIndex.Resume; // Execute ?
///		ThreadRun; // Garder si possible position d'avant ??
		EnableTimer;
	end
end;


function TASCIIView.GetMargin : Boolean;
begin
	Result := (FMargin > MARGE);
end;


procedure TASCIIView.SetMargin(AV: Boolean);
begin
	if AV <> Margin then
	begin
		if AV then
			FMargin := MARGE+GOUTTIERE
		else
			FMargin := MARGE;
		ChangeBord;
	end
end;


	{-------------}
constructor TASCIIView.Create(AOwner : TComponent);
begin
	inherited Create(AOwner);

	ControlStyle := ControlStyle - [csSetCaption];

	FLSel := -1;
	FLCol := -1;
	Width := 181;
	Height := 97;
	//LngLineMax := 0;
	StdFont := TFont.Create;
	StdFont.Assign(Font);
	Font.Name := 'FixedSys';
	ASCIIFile := -1;
	TabStop := True;
	Bord := 1; // BevelOuter Raised
	HorzScrollBar := TScrollBar.Create(Self);
	HorzScrollBar.Left := Bord {+FMargin};
	HorzScrollBar.LargeChange := 10;
	HorzScrollBar.TabStop := False;
	HorzScrollBar.ControlStyle := []; // -csFramed
	HorzScrollBar.OnChange := ScrollHChange;
	HorzScrollBar.Parent := Self;

	VertScrollBar := TScrollBar.Create(Self);
	VertScrollBar.Top := Bord+HBAR;
	VertScrollBar.Kind := sbVertical;
		//LargeChange := 10;
	VertScrollBar.TabStop := False;
//	VertScrollBar.ControlStyle := []; // ?
	VertScrollBar.OnScroll := ScrollV;
	VertScrollBar.Parent := Self;

	Timer := TTimer.Create(Self); // Essayer OnIdle / HandleMessage ??
	Timer.Enabled := False;
	Timer.Interval := 50;
	Timer.OnTimer := IndexExec;

	 // Canvas.Text... nécessite Parent !
	{Parent := TWinControl(AOwner); // Provoque WMSize en Design !
	if HCar = 0 then
		Size;
	Parent := Nil;}
end;


destructor TASCIIView.Destroy;
begin
	Timer.Free;
	if ASCIIFile >= 0 then
	begin
		// if Running then
		FileClose(ASCIIFile);	//ASCIIFile := 0;
	end;
	if PtJ <> Nil then
		FreeMem(PtJ);
	StdFont.Free;

	inherited Destroy;
end;


procedure TASCIIView.Loaded;
begin
	inherited Loaded;
//	if not (csDesigning in ComponentState) then
//?		CurRow := 0; // iLineBuf,VertPosWin,CurRowWin := 0, ReadLines, UpdateTextSel
	Size;	// Une fois bord calculé
	SetFileName(FileName); // Différer après affectation de Margin, OnGetLine
end;


procedure TASCIIView.CaretMove;
begin
	if Focused then
	begin
		SetCaretPos(Bord+FMargin + CurColWin*WCar, Bord+HBAR+1 + CurRowWin*HCar);
		if Assigned(OnCaretMove) then OnCaretMove(Self);
	end
end;


	{ Diffère traitement pour que le thread soit en route
procedure TASCIIView.UMThreadOK(var Msg: TMessage);
begin
	CurRow := 0; // iLineBuf,VertPosWin,CurRowWin := 0, ReadLines, UpdateTextSel
end;
}

procedure TASCIIView.WMSize(var Msg: TWMSize);
begin
  inherited;
{	if HCar = 0 then
	begin
		Size; // Modifie HCar
		SetFileName(FFileName); // Charge le fichier
	end
	else   }
		Size;
end;


procedure TASCIIView.WMLButtonDown(var Msg: TWMLButtonDown);
var	P: TPoint;
	N: Integer;
	bModif: Boolean;
begin
  inherited;
  if LngLineMax <> 0 then	// Fait rien si pas de texte ?!
{	if not Focused then
		SetFocus
	else }
	begin
		SetFocus;
		P.X := Msg.XPos; // TSmallPoint !
		P.Y := Msg.YPos;
		if PtInRect(Rect(Bord+FMargin, Bord+HBAR, VertScrollBar.Left -1, HorzScrollBar.Top -MARGE), P) then	// Clic dans le texte
		begin
			bModif := False;
			N := (P.Y - (Bord+HBAR) -1) div HCar;
			if CurRowWin <> N then
			begin
				bModif := True;
				CurRowWin := N;
				VertScrollBar.Position := VScrollPos;
			end;

			N := ((P.X - (Bord+FMargin)) div WCar) + HorzScrollBar.Position;
			if CurCol <> N then
				CurCol := N // CaretMove, UpdateTextSel, HorzScrollBar
			else if bModif then
			begin
				if FLSel >= 0 then	// Retire sélection
					UpdateTextSel;
				CaretMove;
			end;

			bMouseDown := True; // Pour marquage bloc
			SetCapture(Handle);
		end
	end
end;


procedure TASCIIView.WMLButtonUp(var Msg: TWMLButtonUp);
begin
  inherited;
  bMouseDown := False;
  ReleaseCapture;
end;


procedure TASCIIView.WMMouseMove(var Msg: TWMMouseMove);
var	N: Integer;
begin
  inherited;
  if bMouseDown then
  begin
	N := ((Msg.XPos - (Bord+FMargin)) div WCar) + HorzScrollBar.Position;
	if CurCol <> N then
	begin
		if N < 0 then	N := 0;
		if N > LngLineMax then	N := LngLineMax;

		if FLSel < 0 then	// Nouvelle sélection
		begin
			if CurCol < N then
			begin
				FLSel := CurCol;
				FRSel := N;
			end
			else
			begin
				FLSel := N;
				FRSel := CurCol;
			end
		end
		else
		begin
			if FRSel = CurCol then
				FRSel := N
			else
				FLSel := N;

			if FRSel = FLSel then
				FLSel := -1;	// Défait sélection
		end;

		CurCol := N; // CaretMove, UpdateTextSel, HorzScrollBar
	end
  end
end;


procedure TASCIIView.WMGetDlgCode(var Msg: TWMGetDlgCode);
begin
  Msg.Result := DLGC_WANTARROWS or DLGC_WANTCHARS;
end;


procedure TASCIIView.CMEnter(var Msg: TCMGotFocus);
begin
	inherited;
	if LngLineMax <> 0 then	// Il y a un texte
	begin
		CreateCaret(HAndle, 0, 2, HCar);
		ShowCaret(HAndle);
		CaretMove; // ?? SetCaretPos
	end
end;


procedure TASCIIView.CMExit(var Msg: TCMExit);
begin
	//if LngLineMax <> 0 then	// Il y a un texte
	DestroyCaret;
	inherited;
end;


procedure TASCIIView.KeyDown(var Key: Word; Shift: TShiftState);
var	iCurCol, L: Integer;
	P: PChar;
	bUpdateText: Boolean;
	bCaretMove: Boolean;
//        Rc: TRect;
//        BordMarge,Y,ilig,iCol: Integer;
//        clBack : TColor;
//	bSelChange: Boolean;
begin
  inherited KeyDown(Key, Shift);

  if (Key = 16) or (LngLineMax = 0) then
	Exit;

  if Shift = [ssCtrl] then
  begin
	case Key of
	VK_HOME, VK_END :	// Début, Fin fichier
	  begin
		CurColWin := 0;	// CurCol := 0 sans UpdateText/CaretMove
		HorzScrollBar.Position := 0;
		if Key = VK_HOME then
			CurRow := 0
		else
			CurRow := 99999;	// Sera changé en LineCount !
		//VertScrollBar.Position := 0;
	  end;
	VK_LEFT, VK_RIGHT :
	  begin
		iCurCol := CurCol;
		P := PChar(aLine[VertPosWin+CurRowWin]); // Optimisation
		L := StrLen(P);
		if (iCurCol > L) or ((iCurCol = L) and (Key = VK_RIGHT)) then
			CurCol := L	// Retire sélection !
		else
		begin
			if Key = VK_RIGHT then
			begin
				 // Passe le mot
				while (iCurCol < L) and (P[iCurCol] <> ' ') do
					Inc(iCurCol);
				 // Passe les blancs
				while (iCurCol < L) and (P[iCurCol] = ' ') do
					Inc(iCurCol);
			end
			else
			begin
				 // Passe les blancs
				while (iCurCol > 0) and (P[iCurCol-1] = ' ') do
					Dec(iCurCol);
				 // Passe le mot
				while (iCurCol > 0) and (P[iCurCol-1] <> ' ') do
					Dec(iCurCol);
			end;
			CurCol := iCurCol;	// Retire sélection !
		end;

	  end;

	Ord('C'), VK_INSERT :  	// Ctrl C ou Ctrl Ins
	  begin
		if FLSel >= 0 then                                        // ajout me FSel+1
			ClipBoard.AsText := Copy(aLine[VertPosWin+CurRowWin], FLSel+1, FRSel-FLSel);
	  end
	end;

	//Exit;
  end; // else


  if (Shift = []) or (Shift = [ssShift]) then
  begin
	bUpdateText := False;
	bCaretMove := False;
	iCurCol := CurCol;

	case Key of
        VK_F5 : ;
	VK_UP :
	  begin
		if CurRow = 0 then
			Exit;

		if CurRowWin > 0 then
		begin
			Dec(CurRowWin);
			bCaretMove := True; // VertScrollBar.Position := VScrollPos; si fichier petit ?
		end
		else
		begin
			PrevLine;
			VertScrollBar.Position := VScrollPos;
			bUpdateText := True;
		end
	  end;
	VK_DOWN :
	  begin
		if aLine[VertPosWin+CurRowWin+1] = #0 then
			Exit;

		if CurRowWin < nbRowWin-1 then
		begin
			Inc(CurRowWin);
			bCaretMove := True; // VertScrollBar.Position := VScrollPos; si fichier petit ?
		end
		else
		begin
			NextLine;
			VertScrollBar.Position := VScrollPos;
			bUpdateText := True;
		end
	  end;
	VK_PRIOR :
	  begin
		PrevPage;
		VertScrollBar.Position := VScrollPos;
		bUpdateText := True;
	  end;
	VK_NEXT :
	  begin
		NextPage;
		VertScrollBar.Position := VScrollPos;
		bUpdateText := True;
	  end;
	VK_LEFT :
		//	CurCol := CurCol -1; ??
	  begin
		if CurColWin > 0 then
		begin
			Dec(CurColWin);
			bCaretMove := True;
		end
		else
		begin
			if HorzScrollBar.Position = 0 then
				Exit;

			HorzScrollBar.Position := HorzScrollBar.Position -1;
			//bUpdateText := True; Fait par ScrollHChange
		end
	  end;
	VK_RIGHT :
		//	CurCol := CurCol +1; ??
	  begin
		if CurColWin < nbColWin-1 then
		begin
			Inc(CurColWin);
			bCaretMove := True;
		end
		else
		begin
			if HorzScrollBar.Position = HorzScrollBar.Max then
				Exit;

			HorzScrollBar.Position := HorzScrollBar.Position +1;
			//bUpdateText := True; Fait par ScrollHChange
		end
	  end;
	VK_HOME :
	  begin
		CurCol := 0;	// Retire sélection !
		//HorzScrollBar.Position := 0; SetCaretPos
		//CaretMove;
		Exit;	//bUpdateText := True;
	  end;
	VK_END :
	  begin
		CurCol := Length(aLine[VertPosWin+CurRowWin]);	// Retire sélection !
		//HorzScrollBar.Position := ; SetCaretPos
		//CaretMove;
		Exit;	//bUpdateText := True;
	  end
	else
	  Exit;
	end;


	if (Shift = [ssShift]) and (CurCol <> iCurCol) then
	begin	// Selection change
		case Key of
		VK_LEFT :
		  begin
			if FLSel < 0 then	// Nouvelle sélection
			begin
				FLSel := CurCol;
				FRSel := iCurCol;
			end
			else
			begin
				if FRSel = iCurCol then
					Dec(FRSel)
				else
					Dec(FLSel);

				if FRSel = FLSel then
					FLSel := -1;	// Défait sélection
			end
		  end;
		VK_RIGHT :
		  begin
			if FLSel < 0 then	// Nouvelle sélection
			begin
				FLSel := iCurCol;
				FRSel := CurCol;
			end
			else
			begin
				if FRsel = iCurCol then
					Inc(FRSel)
				else
					Inc(FLSel);

				if FRSel = FLSel then
					FLSel := -1;	// Défait sélection
			end
		  end;
		end;

		bUpdateText := True;
	end
	else if (FLSel >= 0) and (Key <> VK_F5) then
	begin
		FLSel := -1;	// Défait sélection
		//bCaretMove := True;
		bUpdateText := True;
	end;


	if (bUpdateText = True) {or (bSelChange = True)} then
		UpdateText;
	if (bUpdateText = True) or (bCaretMove = True) then
		CaretMove;
  end
end;


procedure TASCIIView.ScrollHChange(Sender: TObject);
begin	// HorzScrollBar.Position mis à jour
	UpdateBar;
	UpdateText;
	CaretMove; // SetCaretPos ?
end;


procedure TASCIIView.ScrollV(Sender: TObject; ScrollCode: TScrollCode; var ScrollPos: Integer);
var	I: Integer;
	J1, J2, OffPos: Longint;
begin
  inherited;

  if LngLineMax = 0 then	// Fait rien si pas de texte
	Exit;

	case ScrollCode of
	scLineDown :
	  begin
		NextLine; // VertSrollBar.Position pas modifiable ???
		ScrollPos := VScrollPos;
	  end;
	scLineUp :
	  begin
		PrevLine;
		ScrollPos := VScrollPos;
	  end;
	scPageDown :
	  begin
		NextPage;
		ScrollPos := VScrollPos;
	  end;
	scPageUp :
	  begin
		PrevPage;
		ScrollPos := VScrollPos;
	  end;
	scPosition :
	  begin
//		ScrollPos à 0 si > 32767 !	-> iFileSize
//		ScrollPos := GetScrollPos(TScrollBar(Sender).Handle, SB_CTL); marche pas mieux
		OffPos := Muldiv(ScrollPos, iFileSize, VertScrollBar.Max - (FLineCount div nbRowWin));
// Détecter Fin car scBottom pas appelé !??
///		if (Running = 0) and (FLineCount < NBLMAX) then
		if not Timer.Enabled and (FLineCount < NBLMAX) then
			CurRow := MulDiv(OffPos, FLineCount, iFileSize) // Cas fin de petit fichier
		else
		begin
			Screen.Cursor := crHourGlass;
			Inc(OffPos, GetJ(0));  // Offset debut du fichier entier
			I := 0;
			J2 := 0;
			repeat // Rechercher dans Jalons un encadrement de l'offset ScrollPos
				J1 := J2;
				Inc(I);
				J2 := GetJ(I);
				if J2 = J1 then // Si J2 change pas : offset maxi
				begin
					Dec(I);
					break;
				end;
			until OffPos <= J2; // Si J2 change pas : offset maxi
			J1 := GetJ(I-1);
//			if J2 > iFileSize then
//				J2 := iFileSize // Cas fin de fichier
			VertPosWin := Muldiv(OffPos - J1, NBLMAX, J2 - J1);
			if VertPosWin > NBLMAX then // Sécurité
				VertPosWin := NBLMAX div 2; // Approximatif ??
			CurRow := NBLMAX*(I-1) + VertPosWin; // Contraire de VScrollPos
			//Screen.Cursor := crDefault;  Fait par SetCurRow
		end;
		Exit;
	  end;
{	scTop :	L'utilisateur a déplacé le curseur de défilement à l'extrémité haute ou gauche de la barre de défilement
	  begin
showmessage('Top');	//ScrollPos := 0;
//		ReadLines(0);
	  end;
	scBottom :	L'utilisateur a déplacé le curseur de défilement à l'extrémité basse ou droite de la barre de défilement
	  begin
showmessage('Bottom'); //Remonter d'un buffer ??
		//ReadLines(??);
		  CurRow := 99999;
	  end; }
	else	// scEndScroll scTrack
		Exit;
	end;

	UpdateTextSel;
	CaretMove; // SetCaretPos ?

{
scPosition	L'utilisateur a positionné le curseur de défilement et l'a relâché
scTrack	L'utilisateur déplace le curseur de défilement
scEndScroll	L'utilisateur a effectué le déplacement du curseur de défilement sur la barre de défilement
}
end;


	{ Retaille suite à changement de bordure }
procedure TASCIIView.ChangeBord;
begin
	Bord := Integer(BevelOuter <> bvNone) + Integer(BevelInner <> bvNone);
	HorzScrollBar.Left := Bord+FMargin-MARGE;
	VertScrollBar.Top := Bord+HBAR;
	Size;
end;

	{ Ajuste le thumb suite à modif de FLineCount ou nbRowWin }
procedure TASCIIView.AjustVScroll;
var	SI: TScrollInfo;
begin
	Assert(nbRowWin <> 0, 'nbRowWin = 0');
	SI.cbSize := SizeOf(SI);
	SI.fMask  := SIF_PAGE;   // Thumb size
	SI.nPage  := FLineCount div nbRowWin;
	SetScrollInfo(VertScrollBar.Handle, SB_CTL, SI, True);
end;

	{ Ajuste le thumb suite à modif de LngLineMax ou nbColWin }
procedure TASCIIView.AjustHScroll;
var	SI: TScrollInfo;
begin
	Assert(nbColWin <> 0, 'nbColWin = 0');
	SI.cbSize := SizeOf(SI);
	SI.fMask := SIF_PAGE;   // Thumb size
	SI.nPage := LngLineMax div nbColWin;
	SetScrollInfo(HorzScrollBar.Handle, SB_CTL, SI, True);
end;


	{ Rend position VertScrollBar (approximatif) de la ligne courante dans le fichier (-Ignored) }
function TASCIIView.VScrollPos : Longint;
var	I: Integer;  // I,I+1 encadrent ligne courante
	N: Integer;  // Nombre de caractères depuis le début du bloc
begin
	 // Longueur du dernier bloc contenant le curseur
	N := CurCol; // Dernière ligne
	for I := ((VertPosWin + CurRowWin) div NBLMAX) * NBLMAX // 0 ou NBLMAX
				to VertPosWin + CurRowWin -1 do  // Sauf dernière ligne
		Inc(N, Length(aLine[I]) + 2);

	I := iLineBuf;
	if VertPosWin >= NBLMAX then	Inc(I); // Deuxième bloc
	Result := Muldiv(GetJ(I) - GetJ(0) + N,
				VertScrollBar.Max - (FLineCount div nbRowWin), iFileSize);
//	Result := GetJ(I);  // Offset du jalon précédent ligne courante
//	Result := Muldiv(Result - GetJ(0) + ( (VertPosWin Mod NBLMAX)*(GetJ(I+1) - Result) ) div NBLMAX,
//				VertScrollBar.Max - (FLineCount div nbRowWin), iFileSize);
end;


	{ Recopie Un demi buffer }
procedure TASCIIView.BufUp;
var	I: Integer;
begin
	for I := 0 to NBLMAX-1 do
	begin
		aLine[I] := aLine[I+NBLMAX];
		aColor[I] := aColor[I+NBLMAX]; // Move ??
		aGlyph[I] := aGlyph[I+NBLMAX]; // Move ??
	end
end;

procedure TASCIIView.BufDown;
var	I: Integer;
begin
	for I := 0 to NBLMAX-1 do
	begin
		aLine[I+NBLMAX] := aLine[I];
		aColor[I+NBLMAX] := aColor[I]; // Move ??
		aGlyph[I+NBLMAX] := aGlyph[I]; // Move ??
	end;
	aLine[NBLBUF] := aLine[NBLMAX]; // Copie une ligne de plus au cas ou Fin de fichier (#0)
end;


	{ Scroll d'une ligne vers le bas }
procedure TASCIIView.NextLine;
begin
	if aLine[VertPosWin+nbRowWin] = #0 then
		Exit;	// Fin de fichier

	if VertPosWin < NBLBUF-nbRowWin then
		Inc(VertPosWin)	// Scroll dans buffer
	else
	begin	// Scolle dans fichier
		BufUp;
		Dec(VertPosWin, NBLMAX-1);
		Inc(iLineBuf);
		ReadLines(1, iLineBuf+1);
	end;

//	UpdateText;
end;


	{ Scroll d'une ligne vers le haut }
procedure TASCIIView.PrevLine;
begin
	if VertPosWin > 0 then
		Dec(VertPosWin)	// Scrolle dans buffer
	else if iLineBuf > 0 then	// Pas en début de fichier
	begin	// Scolle dans fichier
		BufDown;
		Inc(VertPosWin, NBLMAX-1);
		ReadLines(0, iLineBuf-1);
	end;

//	UpdateText;
end;


	{ Scroll d'une page vers le bas }
procedure TASCIIView.NextPage;
begin
	if aLine[VertPosWin+nbRowWin] = #0 then
		Exit;	// Fin de fichier

	if VertPosWin < NBLBUF-nbRowWin-nbRowWin then
	begin
		Inc(VertPosWin, nbRowWin);	// Scroll dans buffer
		while (VertPosWin > 0) and (aLine[VertPosWin+nbRowWin-1] = #0) do	// Recale en fin
			Dec(VertPosWin);
	end
	else
	begin	// Scolle dans fichier
		BufUp;
		Dec(VertPosWin, NBLMAX-nbRowWin);
		Inc(iLineBuf);
		ReadLines(1, iLineBuf+1);
	end;

//	UpdateText;
end;


	{ Scroll d'une page vers le haut }
procedure TASCIIView.PrevPage;
begin
	if VertPosWin >= nbRowWin then
		Dec(VertPosWin, nbRowWin)	// Scroll dans buffer
	else
	begin	// Scroll dans fichier
          if iLineBuf > 0 then	// Pas début de fichier
	  begin
	    BufDown;
	    Inc(VertPosWin, NBLMAX-nbRowWin);
	    ReadLines(0, iLineBuf-1);
	  end
	  else
	    VertPosWin := 0;
	end;

//	UpdateText;
end;


{$warnings off}	// N peut-être pas initialisé
	{ Charge le buffer à partir d'une ligne de départ }
procedure TASCIIView.ReadLines(ABloc: Integer; AIndex: Longint);
var	TmpBuf: TASCIILine;
	I, L, N: Integer;
	iDeb, iFin: Integer;
//	bMax: Boolean;	// True si LngLineMax change
	J: Longint;
begin	// ABloc 0 : 1er bloc, 1 : 2eme bloc, 2 : les deux
		// VertPosWin doit être valide pour modif si fin de fichier
	iDeb := NBLMAX;
	if ABloc <> 1 then
	begin
		iDeb := 0;
		iLineBuf := AIndex; // Nouvelle position du buffer
	end;

	iFin := NBLMAX;
	if ABloc > 0 then
		iFin := NBLBUF; // NBLMAX*2

	J := GetJ(AIndex); // Attend disponibilité
{///	if Running <> 0 then
	begin	// Accès concurent au fichier
		while Running > 1 do
			Sleep(0); // Ne pas interrompre entre Seek et Read !
		//ThreadIndex.Suspend; // Critical section
		SuspendThread(ThreadIndex);
		Sleep(0); // ??
	end;
}
	FileSeek(ASCIIFile, J, soFromBeginning);	// ASCIIStream.Position := Offset ??
//	bMax := False;
	for I := iDeb to iFin-1 do
	begin
	  try	 // Lit une ligne	Eviter repositionnements ??
		N := 0;
		L := FileRead(ASCIIFile, TmpBuf, LNGLMAX);
		///L := ASCIIStream.Read(TmpBuf, LNGLMAX);
		if (L <= 0) or (TmpBuf[0] = #26) then	// Erreur ou fin de fichier ???
		begin
			for L := I to iFin do	// Vide le reste + Fin fichier
				aLine[L] := #0;

			if (I <> iDeb) and	// Quelquechose a été chargé
				not (csDesigning in ComponentState) then
			begin // Ajustement de la position de fin pour VertScrollBar
				iFileSize := FileSeek(ASCIIFile, - GetJ(0), soFromEnd);
				J := I;
				for L := 1 to nbRowWin do	// Retire dernière page (approximatif !) pour VertScrollBar
				begin
					Dec(J);
					if J < 0 then
						Break;
					Dec(iFileSize, Length(aLine[J])+2);
				end
(*
			VertScrollBar.Max := High(Smallint)-page; ?

			begin // Ajustement de la position de fin
				L := VertPosWin;
				VertPosWin := I - nbRowWin; // Dernière page
				if VertPosWin < 0 then	VertPosWin := 0;
				VertScrollBar.Max := //VScrollPos; // Ajuste Thumb=haut de page, Nécessite Jalons à jour
				if L < VertPosWin then // Restaure si possible
					VertPosWin := L;
*)
			end;

			break;	// Fin rencontrée
		end;

		repeat
			if TmpBuf[N] = #0 then	// Remplace les #0 par espaces
				TmpBuf[N] := ' '
			else	 // Recherche fin de ligne $0D0A, $0D ou $0A
			begin
				if TmpBuf[N] = #$0D then
				begin
					if TmpBuf[N+1] = #$0A then	Dec(L);
					break;
				end;
				if TmpBuf[N] = #$0A then
					break;
			end;
			Inc(N);
		until N = L;

		TmpBuf[N] := #0;
		FileSeek(ASCIIFile, N+1-L, soFromCurrent);
		///ASCIIStream.Seek(N+1-L, soFromCurrent);

		aColor[I] := clWindow; // Par défaut
		aGlyph[I] := 0;
		if Assigned(OnGetLine) then	// Permet la modification
			OnGetLine(Self, TmpBuf, aColor[I], aGlyph[I], AIndex*NBLMAX +I);
		aLine[I] := TmpBuf;
		L := StrLen(TmpBuf);
		if L > LngLineMax then
		begin
			LngLineMax := L;
//			bMax := True;
			HorzScrollBar.Max := LngLineMax - nbColWin +1; // +1 pour pouvoir sélectionner fin de ligne !
			AjustHScroll;
		end;
	  except
//Showmessage('Probleme readlines'); ??
		Raise;
	  end
	end;
{///
	if Running <> 0 then
	begin
		ResumeThread(ThreadIndex);
	end;
}
	if (N <> 0) and (aLine[iFin] = #0) then
		aLine[iFin] := ''; // Pas fin de fichier

{	if bMax then	// LngLineMax a augmenté
		HorzScrollBar.Max := LngLineMax - nbColWin +1; // +1 pour pouvoir sélectionner fin de ligne ! }
end;
{$warnings on}


procedure TASCIIView.Paint;
var	I, X, Y, D, N: Integer;
	Rc: TRect;
begin
	inherited Paint; // Dessine bordure

	with Canvas do
	begin
		X := Bord+FMargin;
		Y := Bord+HBAR;
		N := HorzScrollBar.Position;
		 { Barre graduée }
		if FLCol >= 0 then	// Marquage colonne
		begin
			Brush.Color := clInfoBk 	;
			I := X +1 + (FLCol-N)*WCar;
			D := I+(FRCol-FLCol+1)*WCar;
			if D > X then
			begin
				if I < X then
					I := X;
				FillRect(Rect(I, Bord+1, D, Y));
			end;
		end;

		Pen.Color := clBtnShadow; // Ligne de base
		MoveTo(Bord, Y-1);
		LineTo(Width-Bord, Y-1);

		Font := StdFont;	// Remet fonte standard
		Brush.Style := bsClear;
		Rc := Rect(0, Y+1 - HTICK, 9999, Y);
//	CanvasDC := Canvas.Handle;
		for I := nbColWin downto 1 do
		begin
			Inc(X, WCar);
			Inc(N);
			if (N mod 5) = 0 then
			begin
				D := TextWidth(IntToStr(N)+'1') shr 1; // Décalage pour centrage, 1 à 4 chiffres
				TextOut(X - D, Bord+1, IntToStr(N)); // DrawText centré ?
			end;
//			MoveTo(X, Y);
//			LineTo(X, Y+1 - HTICK);  =
			Rc.Left := X;
			DrawEdge(Handle, Rc, EDGE_ETCHED, BF_LEFT); // Mieux que LineTo ?
		end;
{
		Pen.Color := clBtnHighlight;  // 2eme trait -> DrawEdge
		X := Bord+1+FMargin;
		for I := nbColWin downto 1 do
		begin
			Inc(X, WCar);
			MoveTo(X, Y);
			LineTo(X, Y+1 - HTICK);
		end;
}
		Brush.Style := bsSolid;
		//Pen.Color := clBlack;
		//if Margin then  efface gouttière
		 Brush.Color := clBtnFace;	// Couleur Panel
		 FillRect(Rect(Bord, Bord+HBAR, Bord+FMargin-MARGE, HorzScrollBar.Top - 2));
	end;

	UpdateText;
end;


	{ Retire la sélection et redessine le texte }
procedure TASCIIView.UpdateTextSel;
begin
	if FLSel >= 0 then
		LeftSel := -1	// Défait sélection
	else
		UpdateText;
end;

{$warnings off}	// xBitmap, XGlyph et YGlyph peut-être pas initialisés
	{ Redessine la zone de texte (Affiche une page sans effacer le fond) }
procedure TASCIIView.UpdateText;
const WGlyph = 12;
	  HGlyph = 10;
var	I, Y: Integer;
	BordMarge: Integer;
	Base: Integer;
	sTmp: String;
	Rc: TRect;
	xBitmap: TBitmap;
	XGlyph, YGlyph: Integer;
//        clBack : TColor;
//        iLig, iCol : integer;
        SommeRGB : integer;  // utilisé pour le changement de couleur de pinceau
begin
//	if csLoading in ComponentState then
//		Exit; // ?? Rien visible (creation/chargement composant) showing , cliprect ?
	try
	  HideCaret(Handle);
	  with Canvas do
	  begin
		ExcludeClipRect(Handle, VertScrollBar.Left-1, 0, 999, 999); // Au cas où nbColWin déborde à droite
		 { Bordure autour du texte (bas et gauche et gouttière) }
		//Brush.Style := bs;
		BordMarge := Bord+FMargin-MARGE;
		Base := HorzScrollBar.Top - 2;
		if Margin then
		begin
		//	Brush.Color := Color;	// Couleur Panel
		//	FillRect(Rect(Bord, Bord+HBAR, BordMarge, Base)); moins propre que glyph par glyph
			xBitmap := TBitmap.Create;
	//	 si Catégorie 	xBitmap.Transparent := True;	// voir SVLoadBitmap( ?
	//	 Toolbar		xBitmap.TransparentColor := clFuchsia;
			XGlyph := Bord+ (GOUTTIERE-WGlyph) div 2;
			YGlyph := (HCar-HGlyph) div 2;
		end;
		Pen.Color := clWindow;
		MoveTo(BordMarge, Bord+HBAR);
		LineTo(BordMarge, Base);
		LineTo(VertScrollBar.Left, Base);

		Font := Self.Font;
		for I:=0 to nbRowWin-1 do
		begin
			Y := Bord+HBAR + I*HCar;
			if aLine[VertPosWin+I] = #0 then
			begin	// Après dernière ligne
				Brush.Color := Color;	// Couleur Panel
				FillRect(Rect(Bord, Y, VertScrollBar.Left-1, Base));
				Break;
			end;

			if Margin then
				if aGlyph[VertPosWin+I] <> 0 then // and hLib <> 0 !
				begin	// Affiche glyph (12x10 Olive, codes 5000+)
					// voir SVLoadBitmap( ?  mettre en 1800 au lieu de 5000 ??
					//hLib := _LoadSVResDll;
					//if hLib = 0 then
					//	Exit;
//CR 2/10					xBitmap.Handle := LoadBitmap(_LoadSVResDll, MAKEINTRESOURCE(aGlyph[VertPosWin+I]));
//CR 2/10					_ChangeBmpColor(xBitmap, clFuchsia);
//CR 2/10					Draw(XGlyph, Y+YGlyph, xBitmap);
				end
				else
				begin
                                        Brush.Color := clBtnFace;
					FillRect(Rect(XGlyph, Y+YGlyph, XGlyph+WGlyph, Y+YGlyph+HGlyph));
				end;

			Brush.Color := aColor[VertPosWin+I];
			sTmp := Copy(aLine[VertPosWin+I], HorzScrollBar.Position+1, nbColWin);
			if Length(sTmp) < nbColWin then	// Complète à blanc
				sTmp := sTmp + StringOfChar(' ', nbColWin-Length(sTmp));
			TextOut(BordMarge+MARGE, Y, sTmp);

			if (FLSel >= 0) and (I = CurRowWin) then	// Sélection
			begin
                          Rc := Bounds(BordMarge+MARGE + (FLSel-HorzScrollBar.Position) * WCar,
							Y, (FRSel-FLSel) * WCar, HCar +1);
                          if Rc.Left < Bord+FMargin then Rc.Left := Bord+FMargin;
                          if Rc.Right > VertScrollBar.Left then	Rc.Right := VertScrollBar.Left;
			  if Rc.Right > Rc.Left then
                          begin
//  BBY : Gestion des stabylos de couleur
                            if SelColor <> clBlack then
                            begin
                              //clBack := Pixels[Rc.Left + 1, Rc.Top + 1];
                              Brush.Color := SelColor;
                              //if (SelColor > 12000000) or ((SelColor < 7000000) and (SelColor > 50000))
                              //then Font.Color := clBlack
                              //else Font.Color := clWhite;
                              SommeRGB := GetRValue(SelColor) + GetGValue(SelColor) + GetBValue(SelColor);
                              if ( SommeRGB > 385) or (GetGValue(SelColor) = 255)
                              then Font.Color := clBlack else Font.Color := clWhite;
                              TextOut(BordMarge+MARGE + (FLSel-HorzScrollBar.Position) * WCar,Y,
                                                      copy(aLine[VertPosWin+I],FLSel+1,(FRSel-FLSel)));
                              // Eviter que le texte surligné écrase la marge gauche lors d'un scrolling droit
                              Brush.Color := clWhite;
                              Font.Color := clWhite;
                              TextOut(0,Y,'.');
                              Brush.Color := clBtnFace;
                              TextOut(1,Y,'  ');
                              Font.Color := clBlack;

                              // Ancienne méthode - consommant beaucoup de ressources CPU 
                              {for ilig := RC.Top to Rc.Bottom do
                                for iCol := Rc.Left to Rc.Right do
                                  if Pixels[iCol, iLig] = clBack then
                                    Pixels[iCol, iLig] := SelColor
                                  else
                                    Pixels[iCol, iLig] := not (SelColor);}
                            end
                            else
                            begin
//  fin modifs BBY
			      InvertRect(Handle, Rc);
                            end;
                          end;
			end
		end;

		if Margin then
                  xBitmap.Free;

		if FLCol >= 0 then	// Marquage colonne
		begin
			Inc(Base);
			Pen.Color := clBlack;
			Pen.Style := psDot;
			I := BordMarge+MARGE + (FLCol-HorzScrollBar.Position)*WCar;
			if I > BordMarge then // and < VertScrollBar.Left
			begin
				MoveTo(I, Bord+1);
				LineTo(I, Base);
			end;
			Inc(I, (FRCol-FLCol+1)*WCar);
			if I > BordMarge then // and < VertScrollBar.Left
			begin
				MoveTo(I, Bord+1);
				LineTo(I, Base);
			end;
		   	Pen.Style := psSolid;
		end;

		if FTopLine >= 0 then	// Marquage ligne
		begin
			Inc(Base);
			Pen.Color := clBlack;
			Pen.Style := psDot;
                        Y := Bord+HBAR + FTopLine*HCar;
			MoveTo(BordMarge+MARGE, Y);
			LineTo(Width-Bord, Y);
		   	Pen.Style := psSolid;
		end;

	  end;
	finally
	  ShowCaret(Handle);
	end;
end;


procedure TASCIIView.UpdateBar;
var	R: TRect;
begin
	R := Rect(Bord+MARGE, Bord, Width-Bord, Bord+HBAR {-HTICK});	// Graduations
	InvalidateRect(Handle, @R, True);
end;


procedure TASCIIView.Size;
const bSize : Boolean = False;
var	N: Integer;
	CarSize: TSize;
begin
	if bSize or (csLoading in componentState) then
		Exit;
	bSize := True;

	if HCar = 0 then
	begin
		Canvas.Font := Font;
		CarSize := Canvas.TextExtent('A');
		HCar := CarSize.cY; // TextHeight('A')
		WCar := CarSize.cX; // TextWidth('A')
	end;

	 // Nombre entier de lignes pour ScrollWin
	N := Bord +HBAR + MARGE;
	nbRowWin := (Height - N -HorzScrollBar.Height -Bord) div HCar;
	if nbRowWin <= 0 then	nbRowWin := 1;
	if nbRowWin > NBLMAX then	nbRowWin := NBLMAX;
	// reporter marges plutot que redimensionner ??
	Inc(N, nbRowWin * HCar +1); // +1 car erreur sur Scrollbar.Height ! ???
	Height := N + HorzScrollBar.Height +Bord;
	HorzScrollBar.Top := N;
	VertScrollBar.Height := N - (Bord+HBAR);

	N := Width - VertScrollBar.Width -Bord;
	VertScrollBar.Left := N;
	HorzScrollBar.Width := N -(Bord+FMargin);

	nbColWin := (N -(Bord+FMargin) +WCar-1) div WCar;
	if nbColWin <= 0 then	nbColWin := 1;
	if LngLineMax <> 0 then	// Adapte positions du texte
	begin
		//if NewnbRowWin > nbRowWin then
		while (VertPosWin > 0) and (aLine[VertPosWin+nbRowWin-1] = #0) do	// Recale en fin
			Dec(VertPosWin);

		if nbColWin >= LngLineMax then  // Largeur max de ligne affichée entièrement
			HorzScrollBar.Max := 0
		else
			HorzScrollBar.Max := LngLineMax - nbColWin +1; // HorzScrollBar.Position se recadre si > max

		if CurColWin >= nbColWin then
		begin	// Adapte position du curseur
			CurColWin := nbColWin -1; // ?
			CaretMove;
		end;

		Invalidate;
	end;

	bSize := False;
	AjustVScroll;
	AjustHScroll;
end;


	{ Déclanche le Thread pour pouvoir lire jalons et positionner au début }
(*///procedure TASCIIView.ThreadRun;
begin
//	ThreadIndex := CreateThread(Nil, 0, @ThreadIndexExec, Self, 0, ThreadID);
{	with TForm.CreateNew(TComponent(0)) do
	begin
//		WindowState := wsMinimized;
	??  BorderStyle := bsNone;
		Show;
		Release;
	end; ??? }
	Application.ProcessMessages; // Pour effacer fenêtre tout de suite !
	//		Showmessage('Attend thread'); // ???
//	WindowList := DisableTaskWindows(0);
//    EnableTaskWindows(WindowList);
//					PostMessage(Handle, UM_THREADOK, 0, 0);
	CurRow := 0; // iLineBuf,VertPosWin,CurRowWin := 0, ReadLines, UpdateTextSel
end;
*)


 { Recharge le buffer pour appeler OnGetLine suite à modifs }
procedure TASCIIView.UpdateLines;
begin
	ReadLines(2, iLineBuf);
	UpdateTextSel;
end;


	{ Recherche une chaine et positionne le curseur, rend True si trouvé }
function TASCIIView.Find(AStr: string; AOption: Integer; Ajout : integer=0; Suivant : integer=0): Boolean;
var	Posit: Integer;
	Lig, LineBuf: Integer;
	S: string;
begin
	if LngLineMax <> 0 then		// if FLSel >= 0 then	// Sélection
	begin
		Result := True; // Trouvé

		S := Copy(aLine[VertPosWin+CurRowWin], CurCol+1, 9999);
		if ((AOption and foIgnoreCase) <> 0) or
                (AOption = foFromTop) then
		begin
			AStr := UpperCase(AStr);
			S := UpperCase(S);
		end;

		LineBuf := iLineBuf;
		if (AOption = foFromTop) then
		begin
			Lig := 0;
			if iLineBuf <> 0 then	// Pas dans le premier bloc
			begin
				iLineBuf := -2;
				Lig := NBLBUF;
			end
		end
		else
		begin
                 // Recherche sur la ligne
                 if suivant = 0 then
                 begin
			Posit := Pos(AStr, S);
			if Posit > 0 then
			begin
				CurCol := CurCol + Posit -1;
				Exit;
			end;
                 end;
			 // Recherche dans buffer d'abord
			Lig := VertPosWin+CurRowWin+1+suivant;
		end;

		Screen.Cursor := crHourGlass;
		repeat	 // Recherche dans fichier
			while Lig < NBLBUF do
			begin
				S := aLine[Lig];
				if S <> '' then
				begin
					if S[1] = #0 then // Dernière bloc
					begin	// Pas trouvé
						if iLineBuf <> LineBuf then
							ReadLines(2, LineBuf); // Restaure
						Screen.Cursor := crDefault;
						Result := False;
						Exit;
					end;

					if (AOption and foIgnoreCase) <> 0 then
						S := UpperCase(S);
					Posit := Pos(AStr, S);
					if Posit > 0 then
					begin	// Trouvé
						CurRow := iLineBuf*NBLMAX + Lig+Ajout;
						CurCol := Posit -1;
						Screen.Cursor := crDefault;
						Exit;
					end;
				end;

				Inc(Lig);
			end;
			Lig := 0;	// Bloc suivant
			ReadLines(2, iLineBuf+2);
		until False;
	end;
	Result := False;
end;

procedure TASCIIView.InverseColor(Rc : TRect; clBack : TColor);
var iLig, iCol : integer;
begin
for ilig := RC.Top to Rc.Bottom do
    for iCol := Rc.Left to Rc.Right do
        if Canvas.Pixels[iCol, iLig] = clBack then
            Canvas.Pixels[iCol, iLig] := SelColor
            else
            Canvas.Pixels[iCol, iLig] := not (SelColor);
end;


function TASCIIView.SelectText : string;
begin
Result := Copy(aLine[VertPosWin+CurRowWin], FLSel+1, FRSel-FLSel);
end;

procedure TASCIIView.CloseMemo;
begin
      if ASCIIFile >= 0 then
      begin
                      FileClose(ASCIIFile);
                      ASCIIFile := -1;
      end;
end;


// ajout me
procedure Register;
begin
  RegisterComponents('Exemples', [TASCIIView]);
end;

end.







