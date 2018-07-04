{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 21/10/2002
Modifié le ... :   /  /
Description .. : Unité permettant de visualiser le résultat de chaque champ
Suite ........ : du script
Mots clefs ... : 
*****************************************************************}

unit USCRUTER;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons,Grids, Uscript, Qsort, ExtCtrls, ComCtrls,
  HTB97,
{$IFNDEF EAGLCLIENT}
  QRGrid,
{$ENDIF}
  Hmsgbox;

type
  TStrutDlg = class(TForm)
	StringGrid2: TStringGrid;
	Label2: TLabel;
        Timer1: TTimer;
        Compteur: TLabel;
        Dock971: TDock97;
        PBouton: TToolWindow97;
        btnOK: TToolbarButton97;
        btnFermer: TToolbarButton97;
        bDefaire: TToolbarButton97;
        Binsert: TToolbarButton97;
        BImprimer: TToolbarButton97;
        HelpBtn: TToolbarButton97;
        procedure FormCreate(Sender: TObject);
	procedure btnOKClick(Sender: TObject);
	procedure StringGrid1Click(Sender: TObject);
	procedure Timer1Timer(Sender: TObject);
	procedure StringGrid1Enter(Sender: TObject);
	procedure SelectCell(Sender: TObject; Col, Row: Longint;
	  var CanSelect: Boolean);
        procedure HeaderControl1Resize(Sender: TObject);
        procedure FormShow(Sender: TObject);
        procedure HeaderControl1SectionResize(HeaderControl: THeaderControl;
          Section: THeaderSection);
        procedure StringGrid2TopLeftChanged(Sender: TObject);
        procedure StringGrid2KeyDown(Sender: TObject; var Key: Word;
          Shift: TShiftState);
        procedure btnFermerClick(Sender: TObject);
        procedure BImprimerClick(Sender: TObject);
  private
	{ Private declarations }
	FModifier : Boolean;
	FScript : TScript;
	FChamp : TChamp;
	NumEnr, PrevNumEnr : Longint;
	bAbort : Boolean;
	FocusCtrl : TComponent;
	CurRow : Integer;
        CurrentFile : String;

	procedure OnArreterClick(Sender: TObject);
        procedure SaveFile;
        procedure ScruterClick;
  public
	{ Public declarations }
	property Modifier : Boolean read FModifier;
	property Champ : TChamp write FChamp;
	property Script : TScript write FScript;
  end;

var
  StrutDlg: TStrutDlg;

Procedure LectureScrute (Filename : string);

implementation


{$R *.DFM}

Procedure LectureScrute (Filename : string);
begin
          StrutDlg             := TStrutDlg.Create(Application);
          StrutDlg.CurrentFile := Filename;
          StrutDlg.Caption     := 'Résultat : ' + Filename;
          StrutDlg.Show;
end;

procedure TStrutDlg.FormCreate(Sender: TObject);
begin
	FModifier := False;
end;


var
	slAtrie :  TStringList;



function TriEntree(Item1, Item2 : Pointer) : Integer;
begin
	result := AnsiCompareStr(slATrie[Integer(item1)],slATrie[Integer(item2)]);
end;

procedure TStrutDlg.btnOKClick(Sender: TObject);
begin
SaveFile;
end;

procedure TStrutDlg.StringGrid1Click(Sender: TObject);
begin
	;
end;


procedure TStrutDlg.OnArreterClick(Sender: TObject);
begin
	 bAbort := True;
end;


procedure TStrutDlg.Timer1Timer(Sender: TObject);
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


procedure TStrutDlg.ScruterClick;
var
	S : string;
	f : TextFile;
	Count, N : integer;
	(*ScrollW,*) //style1 : Integer;
begin
	AssignFile(f, CurrentFile); Reset(f);
	Count := 0;
        while not eof(f) do
        begin
          readln(f, s); Inc(Count);
        end;
		StringGrid2.RowCount := Count+1;

		Reset(f);
		N := 0;
		while not eof(f) do
		begin
		  readln(f, s);
		  S := Trim(S);
		  if StringGrid2.Cols[0].IndexOf(S) < 0 then
		  begin
		    StringGrid2.Cells[0,N] := S;
		    Inc(N);
		  end;
		end;
		StringGrid2.Cells[0,N] := '';
		StringGrid2.RowCount := N+1;

		//style1 :=
                GetWindowLong(StringGrid2.Handle, GWL_STYLE);
		(*if (style1 and WS_VSCROLL) = WS_VSCROLL then
			ScrollW := GetSystemMetrics(SM_CXVSCROLL)
		else ScrollW := 0;
                *)
		CloseFile(f);
                DeleteFile(CurrentFile);
end;


procedure TStrutDlg.StringGrid1Enter(Sender: TObject);
begin
	FocusCtrl := TComponent(Sender);
end;

procedure TStrutDlg.SelectCell(Sender: TObject; Col,
	Row: Longint; var CanSelect: Boolean);
begin
	CurRow := Row;
end;


(* ------------------------------------------------------------------ *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)

procedure TStrutDlg.HeaderControl1Resize(Sender: TObject);
begin
	;
end;

procedure TStrutDlg.FormShow(Sender: TObject);
begin	// THeaderSection
	StringGrid2.ColWidths[0] := StringGrid2.ClientWidth div 2 -2;
        ScruterClick;
end;

procedure TStrutDlg.HeaderControl1SectionResize(
  HeaderControl: THeaderControl; Section: THeaderSection);
var //ScrollW : Integer;
	//style1 : Integer;
	aStringGrid : TStringGrid;
begin
        aStringGrid := nil;
	case HeaderControl.Tag of
		2: aStringGrid := StringGrid2;
	end;
	//style1 :=
        GetWindowLong(aStringGrid.Handle, GWL_STYLE);
	(*if (style1 and WS_VSCROLL) = WS_VSCROLL then
		ScrollW := GetSystemMetrics(SM_CXVSCROLL)
	else ScrollW := 0;
        *)
	aStringGrid.ColWidths[0] := Section.Width-2;
	HeaderControl.Sections.Items[1].Width := aStringGrid.Width - Section.Width;
end;

procedure TStrutDlg.StringGrid2TopLeftChanged(Sender: TObject);
begin
	;
end;


procedure TStrutDlg.StringGrid2KeyDown(Sender: TObject; var Key: Word;
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

procedure TStrutDlg.btnFermerClick(Sender: TObject);
var
  MB : integer;
begin
  MB := PGIAsk('Voulez-vous enregistrer le fichier :' + CurrentFile, 'Enregistrement');
  if MB = mryes then
    SaveFile;
  Close;
end;

procedure TStrutDlg.SaveFile;
var
F          : TextFile;
N          : integer;
CurrentDir : String;
begin
     if (CurrentFile = '') then
     begin
       with Topendialog.create(Self) do
       begin
         if Execute then
           CurrentFile := FileName;
         Free;
       end;
       SetCurrentDirectory(PChar(CurrentDir));
    end;

    AssignFile(F,CurrentFile) ;
    Rewrite(F) ;
    if IoResult<>0 then
      begin
        PgiBox('Impossible d''écrire dans le fichier #10'+CurrentFile,'Export') ;
        Exit ;
      end ;
    try
    	for N:=0 to StringGrid2.RowCount-1 do
            writeln(F, StringGrid2.Cells[0, N]);
    finally
        CloseFile(F);
    end ;

end;

procedure TStrutDlg.BImprimerClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
     PrintGrid([StringGrid2],caption);
{$ENDIF}
end;

end.
