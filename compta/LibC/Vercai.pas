{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 24/09/2004
Modifié le ... : 24/09/2004
Description .. : - BPY le 24/09/2004 - Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit VerCai;

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
    StdCtrls,
    Buttons,
    ExtCtrls,
    hmsgbox,
    HSysMenu,
    Ent1,
    HEnt1,
    HStatus,
{$IFDEF EAGLCLIENT}
    UTOB,
{$ELSE}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
    Hctrls,
    cpteutil,
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ELSE}
    tCalcCum,
    {$ENDIF MODENT1}
    Rapsuppr;

function ControleCaisse : Boolean;
function CtrlCaiClo : TTraitement;

type
    TFVerCai = class(TForm)
        HPB : TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
        Panel1 : TPanel;
        HMTrad : THSystemMenu;
        FGene : TEdit;
        TFGene : TLabel;
        TFPiece : TLabel;
        FPiece : TEdit;
        TFLigne : TLabel;
        FLigne : TEdit;
        Label1 : TLabel;
        FDateCpta : TEdit;
    BStop: THBitBtn;
        MsgRien : THMsgBox;
        MError : THMsgBox;
        PourClotureOk : TCheckBox;
        Bevel1 : TBevel;
        Bevel2 : TBevel;
        Bevel3 : TBevel;
        Bevel4 : TBevel;
    BAide: THBitBtn;
        procedure BValiderClick(Sender : TObject);
        procedure FormShow(Sender : TObject);
        procedure BStopClick(Sender : TObject);
        procedure FormClose(Sender : TObject;var Action : TCloseAction);
        procedure FormCreate(Sender : TObject);
        procedure BAideClick(Sender : TObject);
    private
    { Déclarations privées }
        OkVerif, StopVerif : Boolean;
        NbError : Integer;
        LaListe : TList;
        procedure LanceControle;
        procedure TuFaisQuoi(LeQ : TQuery);
        function TestBreak : Boolean;
        function GoListe(Q : TQuery;Libel : string;Montants : TabTot) : DelInfo;
    public
    { Déclarations publiques }
    end;

function LeSolde(D, C : Double) : Double;

implementation

{$R *.DFM}

uses UtilPgi;

function ControleCaisse : Boolean;
var
    CCai : TFVerCai;
begin
    Result := True;
    if not _BlocageMonoPoste(True) then Exit;
    CCai := TFVerCai.Create(Application);
    try
        CCai.ShowModal;
        Result := CCai.OkVerif;
    finally
        CCai.Free;
        _DeblocageMonoPoste(True);
    end;
    Screen.Cursor := SyncrDefault;
end;

(*                 *)

function CtrlCaiClo : TTraitement;
var
    CCai : TFVerCai;
begin
    Result := cPasFait;
    if not _BlocageMonoPoste(True) then Exit;
    CCai := TFVerCai.Create(Application);
    try
        CCai.ShowModal;
        case CCai.PourClotureOk.State of
            cbGrayed : Result := cPasFait;
            cbChecked : Result := cOk;
            cbUnChecked : Result := cPasOk;
        end;
    finally
        CCai.Free;
        _DeblocageMonoPoste(True);
    end;
    Screen.Cursor := SyncrDefault;
end;
(**)

procedure TFVerCai.BValiderClick(Sender : TObject);
begin
    EnableControls(Self, False);
    LanceControle;
    EnableControls(Self, True);
//If Not OkVerif then MsgRien.Execute(1,'',IntToStr(NbError)) ;
    if not OkVerif then
    begin
        RapportdeSuppression(Laliste, 7);
        PourClotureOk.State := cbUnChecked;
    end
    else
    begin
        MsgRien.Execute(2, caption, ''); // sinon message tout est ok
        PourClotureOk.State := cbChecked;
    end;
end;

procedure TFVerCai.FormShow(Sender : TObject);
begin
    PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
    OkVerif := True;
    StopVerif := False;
    NbError := 0;
    PourClotureOk.State := cbGrayed;
    FGene.Text := '';
    FDateCpta.Text := '';
    FPiece.Text := '';
    FLigne.Text := '';
end;

function TFVerCai.TestBreak : Boolean;
begin
    Application.ProcessMessages;
    if StopVerif then
    begin
        if MsgRien.Execute(0, '', '') <> mryes then StopVerif := False;
    end;
    Result := StopVerif;
end;

procedure TFVerCai.TuFaisQuoi(LeQ : TQuery);
begin
    FGene.Text := LeQ.FindField('E_GENERAL').AsString;
    FPiece.Text := IntToStr(LeQ.FindField('E_NUMEROPIECE').AsInteger);
    FLigne.Text := IntToStr(LeQ.FindField('E_NUMLIGNE').AsInteger);
    FDateCpta.Text := DateToStr(LeQ.FindField('E_DATECOMPTABLE').AsDateTime);
end;

function TFVerCai.GoListe(Q : TQuery;Libel : string;Montants : TabTot) : DelInfo;
var
    X : DelInfo;
begin
//Inc(NbError) ;
    X := DelInfo.Create;
    X.LeCod := Q.findField('E_GENERAL').AsString;
    X.LeLib := Libel;
    X.LeMess := FloatToStr(Arrondi(Montants[0].TotDebit, V_PGI.OkdecV));
    X.LeMess2 := FloatToStr(Arrondi(Montants[0].TotCredit, V_PGI.OkdecV));
    Result := X;
end;

function LeSolde(D, C : Double) : Double;
begin
    Result := 0;
    if D = C then
        result := abs(D) - abs(C)
    else
        if Abs(D) >= Abs(C) then
        result := abs(D) - abs(C)
    else
        if Abs(D) < Abs(C) then
        result := (abs(C) - abs(D)) * -1;
end;

procedure TFVerCai.LanceControle;
var
    Q1, Q2 : TQuery;
    Monts : TABTOT;
    NbMvt : Integer;
begin
    LaListe.Clear;
    NbError := 0;
    OkVerif := True;
    InitMove(100, '');
    NbMvt := 0;
    Q2 := nil;
    Q1 := OpenSql(' Select G_GENERAL, G_LIBELLE from GENERAUX where G_NATUREGENE="CAI" ', True);
    while not Q1.Eof do
    begin
        Fillchar(Monts, SizeOf(Monts), #0);
        MoveCur(FALSE);
        Q2 := OpenSql(' Select * from ECRITURE where E_GENERAL="' + Q1.Fields[0].AsString + '" ' +
            ' And E_QUALIFPIECE="N" and E_EXERCICE="' + VH^.Encours.Code + '" ' +
            ' And E_ECRANOUVEAU<>"CLO" ' +
            ' Order by E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ', True);
        while not Q2.Eof do
        begin
            TuFaisQuoi(Q2);
            Monts[0].TotDebit := Arrondi(Monts[0].TotDebit + Q2.FindField('E_DEBIT').AsFloat, V_PGI.OkdecV);
            Monts[0].TotCredit := Arrondi(Monts[0].TotCredit + Q2.FindField('E_CREDIT').AsFloat, V_PGI.OkdecV);
            if TestBreak then Break;
            Q2.Next;
            if NbMvt >= 100 then
            begin
                NbMvt := 0;
                FiniMove;
                InitMove(100, '');
            end;
        end;
        if Arrondi(Monts[0].TotCredit, V_PGI.OkdecV) > Arrondi(Monts[0].TotDebit, V_PGI.OkdecV) then
        begin
            OkVerif := False; //Inc(NbError) ;
            LaListe.Add(GoListe(Q2, Q1.findfield('G_LIBELLE').AsString, Monts));
        end;
        if TestBreak then Break;
        Q1.Next;
    end;
    Ferme(Q1);
    if Q2 <> nil then Ferme(Q2);
    FiniMove;
end;

procedure TFVerCai.BStopClick(Sender : TObject);
begin
    StopVerif := True;
end;

procedure TFVerCai.FormClose(Sender : TObject;var Action : TCloseAction);
begin
    LaListe.Free;
end;

procedure TFVerCai.FormCreate(Sender : TObject);
begin
    LaListe := TList.Create;
end;

procedure TFVerCai.BAideClick(Sender : TObject);
begin
    CallHelpTopic(Self);
end;

end.
