{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit GeSecAna;

interface

uses
    Windows,
    Messages,
    SysUtils, // IntToStr
    Classes,
    Graphics, // TCanvas, fsItalic
    Controls,
    Forms, // TForm, TCloseAction, ShowModal, Width, Height
    Grids,
    Hctrls,
    UTob,
{$IFDEF EAGLCLIENT}
{$ELSE}
    DB,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    hmsgbox,
    HSysMenu,
    StdCtrls,
    ExtCtrls,
    Buttons,
    Ent1, // BourreLaDonc, AxeToFb
    HEnt1 // SourisNormale, NowH, BeginTran, CommitTrans
    ;

type
    TCodSSect = class
        UnTab : array[1..17] of HTStrings;
        constructor Create;
        destructor Destroy; override;
    end;

procedure GenereSectionsAna(Faxe, Plan, SSect, SLib : string; Complet : Boolean);

type
    TFGeSecAna = class(TForm)
    BTag: THBitBtn;
    Bdetag: THBitBtn;
        PBouton2 : TPanel;
        PRight : TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
        HMTrad : THSystemMenu;
        HM : THMsgBox;
        FListe : THGrid;
        Nb1 : TLabel;
        procedure FormShow(Sender : TObject);
        procedure FormCreate(Sender : TObject);
        procedure FListeKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
        procedure FListeMouseDown(Sender : TObject; Button : TMouseButton; Shift : TShiftState; X, Y : Integer);
        procedure BTagClick(Sender : TObject);
        procedure BdetagClick(Sender : TObject);
        procedure FormClose(Sender : TObject; var Action : TCloseAction);
        procedure BAideClick(Sender : TObject);
        procedure BValiderClick(Sender : TObject);
    private
        Faxe, Plan, SSect, SLib : string;
        Complet : Boolean;
        Tablo : TCodSSect;
        LgLib : Integer;
        Cod : array[1..17] of string;
        Lib : array[1..17] of string;
        WMinX, WMinY : Integer;
        procedure WMGetMinMaxInfo(var MSG : Tmessage); message WM_GetMinMaxInfo;
        procedure TagDetag(Avec : Boolean);
        procedure CompteElemSelectionner;
        procedure InverseSelection;
        function ListeVide : Boolean;
        procedure GenereCode(j : Integer);
        procedure AfficheCode(StC, StL : string);
        procedure GetCellCanvas(Acol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);
        procedure EcritLenreg(Q : TOB; SCod, SLib : string; DatModif : TDateTime);
    public
        { Déclarations publiques }
    end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus;

constructor TCodSSect.Create;
var
    i : Integer;
begin
    inherited Create;
    for i := 1 to 17 do
        UnTab[i] := nil;
end;

destructor TCodSSect.Destroy;
var
    i : Integer;
begin
    for i := 1 to 17 do
        if UnTab[i] <> nil then
        begin
            UnTab[i].Free;
            UnTab[i] := nil;
        end;
    inherited Destroy;
end;

procedure GenereSectionsAna(Faxe, Plan, SSect, SLib : string; Complet : Boolean);
var
    FGeSecAna : TFGeSecAna;
begin
    FGeSecAna := TFGeSecAna.Create(Application);
    try
        FGeSecAna.Faxe := Faxe;
        FGeSecAna.Plan := Plan;
        FGeSecAna.SSect := SSect;
        FGeSecAna.SLib := SLib;
        FGeSecAna.Complet := Complet;
        FGeSecAna.ShowModal;
    finally
        FGeSecAna.Free;
    end;
    SourisNormale;
end;

procedure TFGeSecAna.BAideClick(Sender : TObject);
begin
    CallHelpTopic(Self);
end;

procedure TFGeSecAna.FormCreate(Sender : TObject);
begin
    WMinX := Width;
    WMinY := Height;
    Tablo := TCodSSect.Create;
    FillChar(Cod, SizeOf(Cod), #0);
    FillChar(Lib, SizeOf(Lib), #0);
end;

procedure TFGeSecAna.FormClose(Sender : TObject; var Action : TCloseAction);
begin
    Tablo.Destroy;
end;

procedure TFGeSecAna.WMGetMinMaxInfo(var MSG : Tmessage);
begin
    with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
    begin
        X := WMinX;
        Y := WMinY;
    end;
end;

procedure TFGeSecAna.TagDetag(Avec : Boolean);
var
    i : Integer;
begin
    for i := 1 to FListe.RowCount - 1 do
        if Avec then FListe.Cells[FListe.ColCount - 1, i] := '*'
        else FListe.Cells[FListe.ColCount - 1, i] := '';
    FListe.Invalidate;
    FListe.SetFocus;
    Bdetag.Visible := Avec;
    BTag.Visible := not Avec;
    CompteElemSelectionner;
end;

procedure TFGeSecAna.BdetagClick(Sender : TObject);
begin
    TagDetag(False);
end;

procedure TFGeSecAna.BTagClick(Sender : TObject);
begin
    TagDetag(True);
end;

procedure TFGeSecAna.CompteElemSelectionner;
var
    i, j : Integer;
begin
    j := 0;
    if not ListeVide then
    begin
        for i := 1 to FListe.RowCount - 1 do
            if FListe.Cells[FListe.ColCount - 1, i] = '*' then Inc(j);
    end;

    case j of
        0, 1 :
            begin
                Nb1.Caption := IntTostr(j) + ' ' + HM.Mess[0];
            end;
    else
        begin
            Nb1.Caption := IntTostr(j) + ' ' + HM.Mess[1];
        end;
    end;
end;

procedure TFGeSecAna.InverseSelection;
begin
    if ListeVide then Exit;
    if FListe.Cells[FListe.ColCount - 1, FListe.Row] = '*' then FListe.Cells[FListe.ColCount - 1, FListe.Row] := ''
    else FListe.Cells[FListe.ColCount - 1, FListe.Row] := '*';
    CompteElemSelectionner;
    FListe.Invalidate;
end;

function TFGeSecAna.ListeVide : Boolean;
begin
    Result := FListe.Cells[0, 1] = '';
end;

procedure TFGeSecAna.FListeKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
begin
    if ((ssShift in Shift) and (Key = VK_DOWN)) or ((ssShift in Shift) and (Key = VK_UP)) then
        InverseSelection
    else
        if (Shift = []) and (Key = VK_SPACE) then
        begin
            InverseSelection;
            if ((FListe.Row < FListe.RowCount - 1) and (Key <> VK_SPACE)) then FListe.Row := FListe.Row + 1;
        end;
end;

procedure TFGeSecAna.FListeMouseDown(Sender : TObject; Button : TMouseButton; Shift : TShiftState; X, Y : Integer);
begin
    if (ssCtrl in Shift) and (Button = mbLeft) then InverseSelection;
end;

procedure TFGeSecAna.GetCellCanvas(Acol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);
begin
    if FListe.Cells[FListe.ColCount - 1, ARow] = '*' then FListe.Canvas.Font.Style := FListe.Canvas.Font.Style + [fsItalic]
    else FListe.Canvas.Font.Style := FListe.Canvas.Font.Style - [fsItalic];
end;

procedure TFGeSecAna.FormShow(Sender : TObject);
var
    i : Integer;
    QLoc, QLocCount : TQuery;
    QSSec : TQuery;
begin
    i := 1;
    FListe.ColWidths[FListe.ColCount - 1] := 0;
    FListe.GetCellCanvas := GetCellCanvas;
    QLocCount := OpenSql('SELECT COUNT(SS_SOUSSECTION) FROM STRUCRSE WHERE SS_AXE="' + FAxe + '"', True);
    LgLib := QLocCount.Fields[0].AsInteger;
    Ferme(QLocCount);
    QLoc := OpenSql('SELECT SS_SOUSSECTION, SS_DEBUT FROM STRUCRSE WHERE SS_AXE="' + Faxe + '" ORDER BY SS_DEBUT ASC', True);

    while not QLoc.Eof do
    begin
        if (Plan <> '') and (Plan = Qloc.Fields[0].AsString) then
        begin
            if Tablo.UnTab[i] = nil then Tablo.UnTab[i] := HTStringList.Create;
            Tablo.UnTab[i].Add(SSect + ';' + SLib);
        end
        else
        begin
            Ferme(QSSec);
            QSSec := OpenSQL('SELECT PS_CODE,PS_LIBELLE FROM SSSTRUCR WHERE PS_AXE="' + FAxe + '" AND PS_SOUSSECTION="' + Qloc.Fields[0].AsString + '"', True);

            while not QSSec.Eof do
            begin
                if Tablo.UnTab[i] = nil then Tablo.UnTab[i] := HTStringList.Create;
                Tablo.UnTab[i].Add(QSSec.Fields[0].AsString + ';' + QSSec.Fields[1].AsString);
                QSSec.Next;
            end;
        end;
        QLoc.Next;
        Inc(i);
    end;
    Ferme(QLoc);
    Ferme(QSSec);
    GenereCode(1);
    CompteElemSelectionner;
    if FListe.RowCount > 2 then FListe.RowCount := FListe.RowCount - 1;
end;

procedure TFGeSecAna.GenereCode(j : Integer);
var
    i, k : Integer;
    St, StC, StL : string;
begin
    if (j <= 17) and (Tablo.UnTab[j] <> nil) then
    begin
        for i := 0 to Tablo.UnTab[j].Count - 1 do
        begin
            St := Tablo.UnTab[j][i];
            Cod[j] := ReadTokenSt(St);
            Lib[j] := Copy(ReadTokenSt(St), 1, 35 div LgLib);
            StC := '';
            StL := '';
            for k := 1 to 17 do
            begin
                StC := StC + Cod[k];
                StL := StL + Lib[k];
            end;
            if Complet then AfficheCode(StC, StL);
            GenereCode(j + 1);
            if (not Complet) and (Cod[LgLib] <> '') then AfficheCode(StC, StL);
        end;
        Cod[j] := '';
        Lib[j] := '';
    end;
end;

procedure TFGeSecAna.AfficheCode(StC, StL : string);
begin
    FListe.Cells[0, FListe.RowCount - 1] := BourreLaDonc(StC, AxeToFb(Faxe));
    FListe.Cells[1, FListe.RowCount - 1] := StL;
    FListe.Cells[FListe.ColCount - 1, FListe.RowCount - 1] := '*';
    FListe.RowCount := FListe.RowCount + 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 29/11/2004
Modifié le ... : 29/11/2004
Description .. : - BPY le 29/11/2004 - fiche n° 14945 - utilisation de TOB a
Suite ........ : la place de TQuery qui ne fonctionne pas en eagl !
Mots clefs ... :
*****************************************************************}
procedure TFGeSecAna.BValiderClick(Sender : TObject);
var
    i : Integer;
    Trouver : Boolean;
    QLoc : TOB;
    StC, StL : string;
    NowFutur : TDateTime;
begin
    Trouver := false;

    for i := 1 to FListe.RowCount - 1 do
    begin
        if (FListe.Cells[FListe.ColCount - 1, i] = '*') then
        begin
            Trouver := true;
            Break;
        end;
    end;

    if (Trouver) then if (HM.Execute(2, '', '') <> mrYes) then Exit;

    NowFutur := NowH;

    try
        BeginTrans;

        QLoc := TOB.Create('', nil, -1);

        InitMove(FListe.RowCount - 1, '');
        for i := 1 to FListe.RowCount - 1 do
        begin
            MoveCur(False);
            if (FListe.Cells[FListe.ColCount - 1, i] = '*') then
            begin
                StC := FListe.Cells[0, i];
                StL := FListe.Cells[1, i];
                if (not PresenceComplexe('SECTION', ['S_SECTION', 'S_AXE'], ['=', '='], [StC, Faxe], ['S', 'S'])) then EcritLenreg(QLoc, StC, StL, NowFutur);
            end;
        end;

        CommitTrans;
        Close;
    except
        HM.Execute(3, '', '');
        rollback;
    end;

    FiniMove;
    FreeAndNil(QLoc);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 29/11/2004
Modifié le ... : 29/11/2004
Description .. : - BPY le 29/11/2004 - fiche n° 14945 - utilisation de TOB a
Suite ........ : la place de TQuery qui ne fonctionne pas en eagl !
Mots clefs ... :
*****************************************************************}
procedure TFGeSecAna.EcritLenreg(Q : TOB; SCod, SLib : string; DatModif : TDateTime);
var
    Qsup : TOB;
begin
    qsup := TOB.Create('SECTION', Q, -1);

    qsup.InitValeurs;
    qsup.SetString('S_SECTION', scod);
    qsup.SetString('S_LIBELLE', Copy(SLib, 1, 35));
    qsup.SetString('S_ABREGE', Copy(SLib, 1, 17));
    qsup.SetString('S_SENS', 'M');
    qsup.SetString('S_AXE', Faxe);

    qsup.InsertDB(nil);
end;

end.
