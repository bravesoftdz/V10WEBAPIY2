{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 21/10/2002
Modifié le ... :   /  /
Description .. : Unité de paramétrage des tables de correspondance
Mots clefs ... :
*****************************************************************}

unit UCorresp;

interface

uses
  Windows, SysUtils, Classes, Controls, Forms, Dialogs,
  StdCtrls, Buttons,Grids, Uscript, ExtCtrls,
  Variants,
  HTB97, Mask, Hctrls, HEnt1, HMsgBox, Menus, uDbxDataSet, AglInit;

type
  TCorrespDlg = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Timer1: TTimer;
    Compteur: TLabel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    cbOpTblCorr: TComboBox;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    btnOK: TToolbarButton97;
    btnFermer: TToolbarButton97;
    bDefaire: TToolbarButton97;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    btnScruter: THBitBtn;
    Label16: TLabel;
    edNomTable: THCritMaskEdit;
    Label3: TLabel;
    StringGrid3: TStringGrid;
    StringGrid4: TStringGrid;
    Label8: TLabel;
    Profile: THCritMaskEdit;
    Label12: TLabel;
    Famille: THCritMaskEdit;
    POPZ: TPopupMenu;
    Breprise: TToolbarButton97;
    Memo1: THCritMaskEdit;
    TDOMAINE: TLabel;
    Domaine: THValComboBox;
    StringGrid2: THGrid;
    StringGrid1: THGrid;
    procedure FormCreate(Sender: TObject);
	  procedure StringGrid1SetEditText(Sender: TObject; ACol, ARow: Longint;
	  const Value: string);
    procedure btnOKClick(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure btnScruterClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure btnSupprimerClick(Sender: TObject);
    procedure StringGrid1Enter(Sender: TObject);
    procedure SelectCell(Sender: TObject; Col, Row: Longint;
	  var CanSelect: Boolean);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure QSort1Compare(Sender: TObject; e1, e2: Word;
	  var Action: Integer);
	  procedure QSort1Swap(Sender: TObject; e1, e2: Word);
    procedure FormShow(Sender: TObject);
    procedure StringGrid2TopLeftChanged(Sender: TObject);
    procedure cbOpTblCorrChange(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure StringGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure StringGrid2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FamilleMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure HelpBtnClick(Sender: TObject);
    procedure StringGrid4DblClick(Sender: TObject);
  private
    { Private declarations }
    FModifier : Boolean;
    FTable : PTableCorrRec;
    FScript : TScript;
    FChamp : TChamp;
    FOpTableCorr : TOpTblCorr;
    NumEnr, PrevNumEnr : Longint;
    bAbort : Boolean;
    FocusCtrl : TComponent;
    CurRow, CurCol : Integer;
    function Callback(ANumEnr:Longint; AEtape:Integer) : integer; stdcall;
    procedure SetTable(Value : PTableCorrRec);
    procedure SetFopTblCorr(value : TOpTblcorr);

    procedure OnArreterClick(Sender: TObject);
    procedure chargeTableExt;
    procedure BMenuZoomClick(Sender: TObject);
  public
    procedure BtnOkQuit(Sender: TObject);
    procedure SetScruteList;
    procedure ShowMAJ(Strl : TstringList);
          procedure Libere;
    { Public declarations }
    property Modifier : Boolean read FModifier;
    property Table    : PTableCorrRec read FTable write SetTable;
    property Champ    : TChamp write FChamp;
    property Script   : TScript write FScript;
    property OpTblCorr : TOpTblCorr read FOpTableCorr write SetFOpTblCorr;
  end;

var
  CorrespDlg: TCorrespDlg;

implementation

uses UDMIMP, UPDomaine;

{$R *.DFM}

procedure TCorrespDlg.FormCreate(Sender: TObject);
begin
	FModifier := False; 
  if FScript = nil then   // cas de liste de correspondance
  begin
      if not Assigned(FTable) then
      begin (* TList *)
            New(FTable);
            with FTable^ do
            begin
                 FEntree := TStringList.Create;
                 FSortie := TStringList.Create;
            end;
      end;
  end;
{$IFNDEF CISXPGI}
  ChargementComboDomaine (Domaine, FALSE);
{$ELSE}
  Domaine.dataType  := 'CPZIMPDOMAINE';
{$ENDIF}
end;

procedure TCorrespDlg.SetTable(Value : PTableCorrRec);
var
	N : Integer;
begin { TStrings String }
	if FTable <> nil then
	begin
		FTable.FEntree.Clear;
		FTable.FSortie.Clear;
	end
    else
    begin
		New(FTable);
		FTable.FEntree := TStringList.Create;
		FTable.FSortie := TStringList.Create;
	end;

	FTable.FEntree.Assign(Value.FEntree);
	FTable.FSortie.Assign(Value.FSortie);

	StringGrid1.RowCount := FTable.FEntree.Count+1;
	for N:=0 to FTable.FEntree.Count-1 do
	begin
		StringGrid1.Cells[0, N] := FTable.FEntree.Strings[N];
		StringGrid1.Cells[1, N] := FTable.FSortie.Strings[N];
	end;
end;

procedure TCorrespDlg.SetFopTblCorr(value : TOpTblcorr);
begin
	 cbOpTblCorr.ItemIndex := Integer(value);
end;

procedure TCorrespDlg.StringGrid1SetEditText(Sender: TObject; ACol,
  ARow: Longint; const Value: string);
begin
	if ARow = StringGrid1.RowCount-1 then
		StringGrid1.RowCount := StringGrid1.RowCount + 1;
	FModifier := True;
end;

var
	slAtrie :  TStringList;

function TriEntree(Item1, Item2 : Pointer) : Integer;
begin
	result := AnsiCompareStr(slATrie[Integer(item1)],slATrie[Integer(item2)]);
end;

procedure TCorrespDlg.btnOKClick(Sender: TObject);
var
	MB : integer;
	N : Integer;
	L : TList;
	SL1, SL2 : TStringList;
begin
	(* sauvegarde de la table de correspondance *)
	if Modifier then
        begin
		(*MB := Application.MessageBox('Voulez-vous enregistrer cette table de correspondance',
									 'Enregistrement',
									 MB_YESNOCANCEL or MB_ICONQUESTION);
		case MB of
		ID_CANCEL: begin ModalResult := mrNone; exit; end;
		end;
                *)
                if (PGIAsk('Voulez-vous enregistrer cette table de correspondance', 'Enregistrement')<> mryes) then
                begin
                     ModalResult := mrNone; exit;
                end;
	end;
	FTable.FEntree.Clear;
	FTable.FSortie.Clear;
	for N:=0 to StringGrid1.RowCount-1 do
	begin
		if (StringGrid1.Cells[0, N] = '') or (StringGrid1.Cells[1, N] = '') then continue;
		FTable.FEntree.Add(StringGrid1.Cells[0, N]);
		FTable.FSortie.Add(StringGrid1.Cells[1, N]);
	end;
	L := TList.Create;
	slAtrie := FTable.FEntree;
	for N:=0 to slATrie.Count-1 do
		L.Add(Pointer(N));
  {$IFNDEF CISXPGI} // fiche 21492
	L.Sort(TriEntree);
  {$ENDIF}
	SL1 := TStringList.Create;
	SL2:= TStringList.Create;
	SL1.Capacity := slATrie.Count;
	SL2.Capacity := slATrie.Count;
	MB := 1;	// pour avoir une seule ligne vide
	for N:=0 to slATrie.Count-1 do
	begin
		if (FTable.FEntree[Integer(L[N])] = '') and
		   (FTable.FSortie[Integer(L[N])] = '') then
		begin
			if MB = 0 then continue;
			Dec(MB);
		end;
		SL1.Add(FTable.FEntree[Integer(L[N])]);
		SL2.Add(FTable.FSortie[Integer(L[N])]);
	end;
 	FTable.FEntree.Assign(SL1);
 	FTable.FSortie.Assign(SL2);
	SL1.Free;
	SL2.Free;
	L.Free;
end;

procedure TCorrespDlg.StringGrid1Click(Sender: TObject);
begin
	;
end;

function TCorrespDlg.Callback(ANumEnr:Longint; AEtape:Integer) : integer; stdcall;
begin
	Application.ProcessMessages;
	NumEnr := ANumEnr;

	if bAbort then Result := -1 else Result := 1;
end;

procedure TCorrespDlg.OnArreterClick(Sender: TObject);
begin
	 bAbort := True;
end;

procedure TCorrespDlg.BtnOkQuit(Sender: TObject);
var N : integer;
begin
	FTable.FEntree.Clear;
	FTable.FSortie.Clear;
	for N:=0 to StringGrid1.RowCount-1 do
	begin
		FTable.FEntree.Add(StringGrid1.Cells[0, N]);
		FTable.FSortie.Add(StringGrid1.Cells[1, N]);
	end;
	 modalResult := IDOK;
end;

procedure TCorrespDlg.Timer1Timer(Sender: TObject);
begin
	if NumEnr <> PrevNumEnr then
	begin
		Compteur.Caption := IntToStr(NumEnr);
		Compteur.Update;
		PrevNumEnr := NumEnr;
	end;
end;

type
	PCreateScruteList =  function : TScruterList; stdcall;

procedure TCorrespDlg.chargeTableExt;
var
  Ft: TextFile;
  TempStr: string;
  FTableCorr: PTableCorrRec;
  TCR : TTableCorrRec;
  function GetEntree(str: string): string;
  var
    i: integer;
  begin
    Result := '';
    for i := 1 to length(str) do
    begin
      if str[i] = ';' then
        break;
      Result := Result + str[i];
    end;
  end;

  function GetSortie(str: string): string;
  var
    i: integer;
  begin
    Result := '';
    for i := pos(';', str) to (length(str) - 1) do
      Result := Result + str[i + 1];
  end;
begin
  if EdNomTable.Text = '' then
  begin
       PGIInfo ('Il faut choisir un fichier', 'Tables de correspondances');
       exit;
  end;
  New(FTableCorr);
  if Assigned(FTableCorr) then
  begin (* TList *)
    New(FTableCorr);
    with FTableCorr^ do
    begin
      FEntree := TStringList.Create;
      FSortie := TStringList.Create;
      AssignFile(Ft, RecomposeChemin(InterpreteVar(EdNomTable.Text,nil)));
      Reset(Ft);
      while not EOF(Ft) do
      begin
        Readln(Ft, TempStr);
        FEntree.add(GetEntree(TempStr));
        FSortie.add(GetSortie(TempStr));
      end;
      TCR.FEntree := FEntree;
      TCR.FSortie := FSortie;
    end;
    closeFile(Ft);
    Table := @TCR;
    FtableCorr.FEntree.free;
    FtableCorr.FSortie.free;
    FtableCorr.FEntree := nil;
    FtableCorr.FSortie := nil;
    Dispose(FtableCorr);
  end;

end;

procedure TCorrespDlg.btnScruterClick(Sender: TObject);
var
	S, TempCond : string;
	f : TextFile;
	Count, N : integer;
	prevOnclick : TNotifyEvent;
	ScrollW, style1 : Integer;

	SL : TScruterList;
	CreateScruteList : PCreateScruteList;
	szFileName : array[0..260] of char;
  AScript    : Tscript;
  NameEnrg   : string;
  SR :TScriptRecord;
begin
        if FScript = nil then
        begin
             chargeTableExt;
             exit;
        end;
	SL := nil;	(* -- ne pas supprimer *)
(* A VOIR	if (pos('$', FScript.Options.FileName) <> 0) then
	begin
	   ShowMessage('Le nom du fichier est incorrecte ('+FScript.Options.FileName+').');
	   exit;
	end;
*)

	bAbort := False;
	btnOK.Enabled := False;
	prevOnclick := btnFermer.OnClick;
	btnFermer.ModalResult := mrNone;
	btnFermer.OnClick := OnArreterClick;
	btnScruter.Enabled := False;
	Timer1.Enabled := True;

	try
		@CreateScruteList := GetProcAddress(DMImport.HdlImpfic, 'CreerScruterList');
		if not Assigned(@CreateScruteList) then
			raise Exception.Create('fonction "CreerScruterList" non trouvée');


		if Fchamp.TypeInfo = tiReference then
		begin  // remplacement de la condition pour les references
                      TempCond := FScript.Options.Condition;
                      FScript.Options.condition := Fchamp.Reference.Condition;
		end
    else
    begin
          AScript  := FScript.CloneScript;
          NameEnrg := FChamp.Name;
          NameEnrg := ReadTokenPipe (NameEnrg, '_');
          AScript.Assign(FScript, NameEnrg);
          if AScript.champ[0].ConditionChamp <> '' then
           FChamp.conditionchamp := AScript.champ[0].ConditionChamp;
          // Pb écrasement de condition général du script
          // FScript.options.condition := AScript.champ[0].ConditionChamp;
    end;
    With FScript do
    begin
          // Initialisation de record pour la Dll
//          SR.FSignature := Signature;
          SR.FOptions   := Options;
          SR.FOptionCpt := OptionsComptable;
          SR.FChamps    := Champ;
 //         SR.FTableCorr := TableCorr;
          SR.FName      := Name;
          SR.FFileType  := FileType; (* 0:base 1:ASCII *)
          SR.FASCIIMODE := ASCIIMODE; (* 0:Fixe 1:Delimite *)
          SR.FDestTable := DestTable; (* 0:Dossier 1:Utilisateur *)
          SR.FScriptSuivant:= ScriptSuivant;
          SR.FParName   := ParName; (* CLE PARAMETRE  *)
          SR.FVariable  := Variable;
          SR.FVariableList:= Variables;
          SR.FParOrig   := ParOrig;
    end;

		SL := CreateScruteList;
		try
                   szFileName := '\0';
		   SL.Scrute(szFileName, Integer(Self), @SR, FChamp, Callback);
		except
			if (itrace > 0) then ShowMessage(Exception(ExceptObject).message);
		end;

		 if Fchamp.TypeInfo = tiReference then
			FScript.options.condition := TempCond;

		Timer1Timer(Timer1);
		Timer1.Enabled := False;

		AssignFile(f, szFileName); Reset(f);
		Count := 0;
		while not eof(f) do
			begin readln(f, s); Inc(Count); end;
		StringGrid2.RowCount := Count+1;

		Reset(f);
		N := 0;
		while not eof(f) do
		begin
			readln(f, s);
			S := Trim(S);
			if StringGrid1.Cols[0].IndexOf(S) < 0 then
			begin
				StringGrid2.Cells[0,N] := S;
				StringGrid2.Cells[1,N] := S;
				Inc(N);
			end;
		end;
		StringGrid2.Cells[0,N] := '';
		StringGrid2.Cells[1,N] := '';
		StringGrid2.RowCount := N+1;

		style1 := GetWindowLong(StringGrid2.Handle, GWL_STYLE);
		if (style1 and WS_VSCROLL) = WS_VSCROLL then
			ScrollW := GetSystemMetrics(SM_CXVSCROLL)
		else ScrollW := 0;
		StringGrid2.ColWidths[1] := StringGrid2.ColWidths[1] - ScrollW - 2;
		CloseFile(f);
	finally
		FModifier := True;
		btnFermer.ModalResult := mrCancel;
		btnOK.Enabled := True;
		btnScruter.Enabled := True;
		btnFermer.OnClick := prevOnClick;

		if Assigned(SL) then
			SL.Free;
	end;
end;

procedure TCorrespDlg.SetScruteList;
var
	S : string;
	f : TextFile;
	Count, N : integer;
	prevOnclick : TNotifyEvent;
	ScrollW, style1 : Integer;
	SL : TScruterList;
	CreateScruteList : PCreateScruteList;
	szFileName : array[0..260] of char;
  SR :TScriptRecord;

begin
	SL := nil;	(* -- ne pas supprimer *)

	bAbort := False;
	prevOnclick := btnFermer.OnClick;
	btnFermer.OnClick := OnArreterClick;
	Timer1.Enabled := True;
  With FScript do
  begin
        // Initialisation de record pour la Dll
 //       SR.FSignature := Signature;
        SR.FOptions   := Options;
        SR.FOptionCpt := OptionsComptable;
        SR.FChamps    := Champ;
 //       SR.FTableCorr := TableCorr;
        SR.FName      := Name;
        SR.FFileType  := FileType; (* 0:base 1:ASCII *)
        SR.FASCIIMODE := ASCIIMODE; (* 0:Fixe 1:Delimite *)
        SR.FDestTable := DestTable; (* 0:Dossier 1:Utilisateur *)
        SR.FScriptSuivant:= ScriptSuivant;
        SR.FParName   := ParName; (* CLE PARAMETRE  *)
        SR.FVariable  := Variable;
        SR.FVariableList:= Variables;
        SR.FParOrig   := ParOrig;
  end;


	try
		@CreateScruteList := GetProcAddress(DMImport.HdlImpfic, 'CreerScruterList');
		if not Assigned(@CreateScruteList) then
			raise Exception.Create('fonction "CreerScruterList" non trouvée');

		SL := CreateScruteList;
                szFileName := '\0';
		SL.Scrute(szFileName, Integer(Self), @SR, FChamp, Callback);

		Timer1Timer(Timer1);
		Timer1.Enabled := False;
		AssignFile(f, szFileName); Reset(f);
		Count := 0;
		while not eof(f) do
			begin readln(f, s); Inc(Count); end;
		StringGrid2.RowCount := Count+1;

		Reset(f);
		N := 0;
		while not eof(f) do
		begin
			readln(f, s);
			if StringGrid1.Cols[0].IndexOf(S) < 0 then begin
				StringGrid2.Cells[0,N] := S;
				StringGrid2.Cells[1,N] := S;
				Inc(N);
			end;
		end;
		StringGrid2.Cells[0,N] := '';
		StringGrid2.Cells[1,N] := '';
		StringGrid2.RowCount := N+1;

		style1 := GetWindowLong(StringGrid2.Handle, GWL_STYLE);
		if (style1 and WS_VSCROLL) = WS_VSCROLL then
			ScrollW := GetSystemMetrics(SM_CXVSCROLL)
		else ScrollW := 0;
		StringGrid2.ColWidths[1] := StringGrid2.ColWidths[1] - ScrollW - 2;
		CloseFile(f);
	finally
		if Assigned(SL) then
			SL.Free;
		FModifier := True;
		btnFermer.OnClick := prevOnClick;
	end;
end;


procedure TCorrespDlg.btnSupprimerClick(Sender: TObject);
{$IFNDEF CISXPGI} var N : integer; {$ENDIF}
begin
{$IFNDEF CISXPGI}
	if FocusCtrl is TStringGrid then
		with FocusCtrl as TStringGrid do
        begin
			for N:=CurRow+1 to RowCount-1 do
				Rows[N-1] := Rows[N];
			RowCount := RowCount - 1;
		end;
{$ELSE}
  StringGrid1.Cells[CurCol, CurRow] := '';
{$ENDIF}
end;

procedure TCorrespDlg.StringGrid1Enter(Sender: TObject);
begin
	FocusCtrl := TComponent(Sender);
end;

procedure TCorrespDlg.SelectCell(Sender: TObject; Col,
	Row: Longint; var CanSelect: Boolean);
begin
	CurRow := Row;
  CurCol := Col;
end;

procedure TCorrespDlg.Button1Click(Sender: TObject);
var L1, L2, N, M : Integer;
begin { -- Copier les elements selectionnes }
	L1 := StringGrid2.Selection.Top;
	L2 := StringGrid2.Selection.Bottom;

	if L1 < 0 then exit;

	if (StringGrid2.Cells[0,L1] = '') or (StringGrid2.Cells[0,L1] = '') then
		exit;

	M := StringGrid1.RowCount-1;
	StringGrid1.RowCount := StringGrid1.RowCount + (L2-L1+1);
	for N:=L1 to L2 do
		begin StringGrid1.Rows[M] := StringGrid2.Rows[N];	Inc(M); end;
	M := L1;
	for N:=L2+1 to StringGrid2.RowCount-1 do
		begin StringGrid2.Rows[M] := StringGrid2.Rows[N];	Inc(M); end;
	StringGrid2.RowCount := StringGrid2.RowCount - (L2-L1+1);
end;

procedure TCorrespDlg.Button2Click(Sender: TObject);
var L1, L2, N, M : Integer;
begin { -- Copier les elements selectionnes }
	L1 := StringGrid1.Selection.Top;
	L2 := StringGrid1.Selection.Bottom;

	if L1 < 0 then exit;
	if (StringGrid1.Cells[0,L1] = '') or (StringGrid1.Cells[0,L1] = '') then
		exit;

	M := StringGrid2.RowCount-1;
	StringGrid2.RowCount := StringGrid2.RowCount + (L2-L1+1);
	for N:=L1 to L2 do
		begin StringGrid2.Rows[M] := StringGrid1.Rows[N];	Inc(M); end;
	M := L1;
	for N:=L2+1 to StringGrid1.RowCount-1 do
		begin StringGrid1.Rows[M] := StringGrid1.Rows[N];	Inc(M); end;
	StringGrid1.RowCount := StringGrid1.RowCount - (L2-L1+1);
end;

(* ------------------------------------------------------------------ *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)

procedure TCorrespDlg.QSort1Compare(Sender: TObject; e1, e2: Word;
  var Action: Integer);
begin
	Action := AnsiCompareStr(FTable.FEntree[e1-1], FTable.FEntree[e2-1]);
end;

procedure TCorrespDlg.QSort1Swap(Sender: TObject; e1, e2: Word);
var S : String;
begin
	S := FTable.FEntree[e1-1];
	FTable.FEntree[e1-1] := FTable.FEntree[e2-1];
	FTable.FEntree[e2-1] := S;

	S := FTable.FSortie[e1-1];
	FTable.FSortie[e1-1] := FTable.FSortie[e2-1];
	FTable.FSortie[e2-1] := S;
end;


procedure TCorrespDlg.FormShow(Sender: TObject);
begin
	StringGrid2.ColWidths[0] := StringGrid2.ClientWidth div 2 -2;
	StringGrid2.ColWidths[1] := StringGrid2.ClientWidth div 2 -2;
	StringGrid1.ColWidths[0] := StringGrid1.ClientWidth div 2 -2;
	StringGrid1.ColWidths[1] := StringGrid1.ClientWidth div 2 -2;
        StringGrid3.Cells[0, 0] := TraduireMemoire('Entrée');
        StringGrid3.Cells[1, 0] := TraduireMemoire('Sortie');
        StringGrid4.Cells[0, 0] := TraduireMemoire('Entrée');
        StringGrid4.Cells[1, 0] := TraduireMemoire('Sortie');
        if FScript <> nil then
        begin
             edNomTable.Enabled := FALSE;
             Profile.Enabled    := FALSE;
             Famille.Enabled    := FALSE;
             Label16.Enabled    := FALSE;
             Label8.Enabled     := FALSE;
             Label12.Enabled    := FALSE;
             Domaine.Enabled    := FALSE;
             TDomaine.Enabled   := FALSE;
        end
        else
        begin
             StringGrid2.Enabled := FALSE;
             Button1.Enabled     := FALSE;
             Button2.Enabled     := FALSE;
             Button3.Enabled     := FALSE;
             Button4.Enabled     := FALSE;
             StringGrid3.Enabled := FALSE;
             Label2.Enabled      := FALSE;
             Breprise.Enabled    := FALSE;
        end;
             HelpContext        := 1210;
end;


procedure TCorrespDlg.StringGrid2TopLeftChanged(Sender: TObject);
begin
	;
end;

procedure TCorrespDlg.cbOpTblCorrChange(Sender: TObject);
begin
	 FopTableCorr := TopTblCorr(cbOpTblCorr.itemIndex);
end;

procedure TCorrespDlg.ShowMAJ(strl : TStringList);
var
	Count, N, I : integer;
	prevOnclick : TNotifyEvent;
	ScrollW, style1 : Integer;
begin
	bAbort := False;
	btnOK.Enabled := False;
	prevOnclick := btnFermer.OnClick;
	btnFermer.ModalResult := mrNone;
	btnFermer.OnClick := OnArreterClick;
	btnScruter.Enabled := False;

    try
		Count := 0;
		N := 0;

        for I := 0 to Strl.count-1 do
        begin
		  if StringGrid1.Cols[0].IndexOf(Strl.strings[I]) < 0 then
                  begin
                        StringGrid2.RowCount := Count+1;
			StringGrid2.Cells[0,N] := Strl.strings[I];
			StringGrid2.Cells[1,N] := Strl.strings[I];
			Inc(N);
                  end;
        end;

		StringGrid2.Cells[0,N] := '';
		StringGrid2.Cells[1,N] := '';
		StringGrid2.RowCount := N+1;

		style1 := GetWindowLong(StringGrid2.Handle, GWL_STYLE);
		if (style1 and WS_VSCROLL) = WS_VSCROLL then
			ScrollW := GetSystemMetrics(SM_CXVSCROLL)
		else ScrollW := 0;
		StringGrid2.ColWidths[1] := StringGrid2.ColWidths[1] - ScrollW - 2;
	finally
		FModifier := True;

		btnFermer.ModalResult := mrCancel;
		btnOK.Enabled := True; btnScruter.Enabled := True;
		btnFermer.OnClick := prevOnClick;
	end;

end;


procedure TCorrespDlg.Button3Click(Sender: TObject);
var L1, L2, N, M : Integer;
begin
	L1 := 0;
	L2 := StringGrid2.RowCount-1;

	if L1 < 0 then exit;

	if (StringGrid2.Cells[0,L1] = '') or (StringGrid2.Cells[0,L1] = '') then
		exit;

	M := StringGrid1.RowCount-1;
	StringGrid1.RowCount := StringGrid1.RowCount + (L2-L1+1);
	for N:=L1 to L2 do
		begin StringGrid1.Rows[M] := StringGrid2.Rows[N];	Inc(M); end;
	M := L1;
	for N:=L2 to StringGrid2.RowCount-1 do
		begin StringGrid2.Rows[M] := StringGrid2.Rows[N];	Inc(M); end;
	StringGrid2.RowCount := StringGrid2.RowCount - (L2-L1+1);
end;

procedure TCorrespDlg.Button4Click(Sender: TObject);
var L1, L2, N, M : Integer;
begin
	L1 := 0;//StringGrid1.Selection.Top;
	L2 := StringGrid1.rowCount-1;

	if L1 < 0 then exit;
	if (StringGrid1.Cells[0,L1] = '') or (StringGrid1.Cells[0,L1] = '') then
		exit;

	M := StringGrid2.RowCount-1;
	StringGrid2.RowCount := StringGrid2.RowCount + (L2-L1+1);
	for N:=L1 to L2 do
		begin StringGrid2.Rows[M] := StringGrid1.Rows[N];	Inc(M); end;
	M := L1;
	for N:=L2 to StringGrid1.RowCount-1 do
		begin StringGrid1.Rows[M] := StringGrid1.Rows[N];	Inc(M); end;
	StringGrid1.RowCount := StringGrid1.RowCount - (L2-L1+1);
end;

procedure TCorrespDlg.StringGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var N : integer;
begin
Inherited ;
     if key <> 46 then exit; (* Suppr *)

  case Key of
     VK_DELETE     : if Shift=[ssAlt] then
                    begin
                    Key:=0 ;
                    end
                     else
                      if Shift=[ssCtrl] then
                       begin
                           with StringGrid1 do
                           begin
                              for N:=CurRow+1 to RowCount-1 do
                                  Rows[N-1] := Rows[N];
                              RowCount := RowCount - 1;
                           end;
                       end ;
    end;
end;

procedure TCorrespDlg.StringGrid2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var N : integer;
begin
     if key <> 46 then exit; (* Suppr *)
     with StringGrid2 do
     begin
        for N:=CurRow+1 to RowCount-1 do
            Rows[N-1] := Rows[N];
        RowCount := RowCount - 1;
     end;
end;

procedure TCorrespDlg.Libere;
begin
	if FTable <> nil then
	begin
            FTable.FEntree.Free;
            FTable.FEntree := nil;
            FTable.FSortie.free;
            FTable.FSortie := nil;
            Dispose(FTable);
	end;
end;

function CreerLigPop ( Name : string ; Owner : TComponent) : TMenuItem ;
Var
    T   : TMenuItem ;
BEGIN
T:=TMenuItem.Create(nil) ; T.Name:=Name ;
T.Caption:=Name ;
Result:=T ;
END ;

procedure TCorrespDlg.BMenuZoomClick(Sender: TObject);
var
  MI               : TMenuItem;
begin
  MI := Sender as TMenuItem;
  Famille.text  := MI.Name;

end;

procedure TCorrespDlg.FamilleMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
QP         : TQuery;
T          : TMenuItem ;
Index      : integer;
pt         : TPoint;

      function IndexOf (Name : string; PP: TPopupMenu): Boolean;
      var
      d : integer;
      begin
                 Result := FALSE;
                 Index := -1;
                 for d := 0 to  POPZ.Items.count-1 do
                 begin
                            if PP.Items[d].Name = Name then
                            begin Result := TRUE; Index := d; break; end;
                 end;
      end;
begin
  inherited;
      PurgePopup(POPZ);  T := nil;
      QP := OpenSQLADO ('Select FamCorresp from '+DMImport.GzImpPar.TableName+
                        ' Where FamCorresp<>""', DMImport.DBGlobal.ConnectionString);

      While  not (QP.Eof) do
      begin
           if not (IndexOf (QP.FindField('FamCorresp').asstring, POPZ)) then
           begin
                      T := CreerLigPop(QP.FindField('FamCorresp').asstring, T);
                      T.OnClick := BMenuZoomClick;
                      POPZ.Items.Add(T) ;
           end;

           QP.Next;
      end;
      QP.Free;
      if ssRight in Shift then
      begin
            pt.X := X; pt.Y := Y;
            pt := Famille.ClientToScreen(pt);
            POPZ.Popup(pt.X, pt.Y);
      end;
end;


procedure TCorrespDlg.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TCorrespDlg.StringGrid4DblClick(Sender: TObject);
var
	MB : integer;
	N : Integer;
	L : TList;
begin
        (* sauvegarde de la table de correspondance *)
        FTable.FEntree.Clear;
        FTable.FSortie.Clear;
        for N:=0 to StringGrid1.RowCount-1 do
        begin
          if (StringGrid1.Cells[0, N] = '') or (StringGrid1.Cells[1, N] = '') then continue;
          FTable.FEntree.Add(StringGrid1.Cells[0, N]);
          FTable.FSortie.Add(StringGrid1.Cells[1, N]);
        end;
        L := TList.Create;
        slAtrie := FTable.FEntree;
        for N:=0 to slATrie.Count-1 do
          L.Add(Pointer(N));
        L.Sort(TriEntree);
        MB := 1;	// pour avoir une seule ligne vide
        for N:=0 to slATrie.Count-1 do
        begin
          if (FTable.FEntree[Integer(L[N])] = '') and
             (FTable.FSortie[Integer(L[N])] = '') then
          begin
            if MB = 0 then continue;
            Dec(MB);
          end;
          StringGrid1.Cells[0, N] :=  FTable.FEntree[Integer(L[N])];
          StringGrid1.Cells[1, N] :=  FTable.FSortie[Integer(L[N])];
        end;
        L.Free;
        FTable.FEntree.Clear;
        FTable.FSortie.Clear;
end;

end.
