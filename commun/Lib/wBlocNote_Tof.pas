{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 21/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : WBLOCNOTE ()
Mots clefs ... : TOF;WBLOCNOTE
*****************************************************************}
Unit WBLOCNOTE_TOF ;

Interface

Uses
	StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  HTB97,
  UTOF,
  uTob,
  Dialogs,
  uTom,
  wCommuns,
  wTof
  ;

Type
  TBNProperties = Set of (bnReadOnly, bnObligatoire, bnByString, bnByTob, bnBySQL, bnByFile, bnZipped, bnSection);

Type
  TOF_WBLOCNOTE = Class (tWTOF)
  	procedure OnNew                ; override;
  	procedure OnDelete             ; override;
  	procedure OnUpdate             ; override;
  	procedure OnLoad               ; override;
   	procedure OnArgument(S: String); override;
   	procedure OnClose              ; override;
  private
    function GetCoords: String;
    procedure SetCoords(const Value: String);
    function GetBlocNoteSize: Integer;
    procedure SetBlocNoteSize(const Value: Integer);
	private
   	Prefixe, Where, EtatRev, BlobFieldSuffixe, FilePath, LaString: string;
    TobST: Tob;
    Properties: TBNProperties;
    FSearch: TMemoFindDialog;
    FMemoOrigine: String;
    procedure BVALIDER_OnClick(Sender: TObject);
    procedure BTSearch_OnClick(Sender: TObject);
    procedure Form_KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Memo_Click(Sender: TObject);
    procedure Memo_KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Memo_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SetMemoCoord(Sender: TObject);

    property Coords: String read GetCoords write SetCoords;
    property BlocNoteSize: Integer read GetBlocNoteSize write SetBlocNoteSize;
  end;

const
  ErrOblig = 1;

Implementation

Uses
  Vierge,
  Windows,
  Messages,
  Math,
  ExtCtrls,
  UtilPGI
  ;

procedure TOF_WBLOCNOTE.OnNew;
begin
  Inherited;
end;

procedure TOF_WBLOCNOTE.OnDelete;
begin
  Inherited;
end;

procedure TOF_WBLOCNOTE.OnUpdate;
begin
  Inherited;
end;

procedure TOF_WBLOCNOTE.OnLoad;
begin
  Inherited;

  { Rend inaccessible le memo si on n'est pas en modif. }
  if Etatrev <> '' then
  begin
    SetControlEnabled('BN1', EtatRev = 'MOD');
    SetControlProperty('BN1', 'READONLY', EtatRev <> 'MOD');
  end;
end;

procedure TOF_WBLOCNOTE.OnArgument(S: String);

  function GetLaString: String;
  var
    iPosDebut: Integer;
  begin
    iPosDebut := Pos('[%DEB%]', S);
    if iPosDebut > 0 then
      Result := Copy(S, iPosDebut + 7, Pos('[%FIN%]', S) - (iPosDebut + 7))
    else
      Result := GetArgumentValue(S, 'LASTRING')
  end;

var
  PropFont: Boolean;  
begin
  Properties := [];

	Inherited;
  EtatRev := GetArgumentValue(S, 'ETATREV');
  Prefixe := GetArgumentValue(S, 'PREFIXE');
  Where := GetArgumentValue(S, 'WHERE');
  TobSt   := Tob(GetArgumentInteger(S, 'TOBST'));
  if Where <> '' then
    Properties := Properties + [bnBySQL];
  BlobFieldSuffixe := GetArgumentValue(S, 'BLOBFIELDSUFFIXE');
  if GetArgumentBoolean(S, 'RO') or (Action = 'CONSULTATION') then
    Properties := Properties + [bnReadOnly];
  FilePath := GetArgumentValue(S, 'FILEPATH');
  if (FilePath <> '') or (Pos('FILEPATH', S) <> 0) then
    Properties := Properties + [bnByFile];
  LaString := GetLaString();
  if (LaString <> '') or GetArgumentBoolean(S, 'FROMST') then
    Properties := Properties + [bnByString];
  if Assigned(TobSt) then
    Properties := Properties + [bnByTob];
  if GetArgumentBoolean(S, 'OBLIGATOIRE') then
    Properties := Properties + [bnObligatoire];
  if GetArgumentBoolean(S, 'ZIPPED') then
    Properties := Properties + [bnZipped];
  if GetArgumentString(S, 'SECTION') <> '' then
    Properties := Properties + [bnSection];

  LastError := 0;

  PropFont := GetArgumentBoolean(S, 'PROPFONT');
  if PropFont then
  begin
    with TMemo(GetControl('BN1')) do
    begin
      Font.Name := 'Courier New';
      OnClick := Memo_Click;
      OnKeyDown := Memo_KeyDown;
      OnKeyUp := Memo_KeyUp;
    end
  end;

  if (BlobFieldSuffixe <> '') and not (bnByFile in Properties) then
    fBlobSuffixe := BlobFieldSuffixe;

  if bnBySQL in Properties then
    TMemo(GetControl('BN1')).Text := wGetSqlFieldValue(Prefixe + '_' + fBlobSuffixe, PrefixeToTable(Prefixe), Where)
  else if bnByString in Properties then
    TMemo(GetControl('BN1')).Text := LaString
  else if bnByFile in Properties then
  begin
    if FileExists(FilePath) then
      TMemo(GetControl('BN1')).Lines.LoadFromFile(FilePath)
    else
      PGIError(Format(TraduireMemoire('Le fichier %s est introuvable.'), [FilePath]));
  end
  else if bnByTob in Properties then
  begin
    if Assigned(TobST) and TobST.FieldExists(Prefixe + iif(Prefixe <> '', '_', '') + fBlobSuffixe) then
      TMemo(GetControl('BN1')).Text := TobST.G(Prefixe + iif(Prefixe <> '', '_', '') + fBlobSuffixe)
  end;

  if bnZipped in Properties then
    tMemo(GetControl('BN1')).Text := wDezipString(TMemo(GetControl('BN1')).Text);

  { Extraction d'une section }
  if bnSection in Properties then
  begin
    FMemoOrigine := tMemo(GetControl('BN1')).Text;
    tMemo(GetControl('BN1')).Text := wExtractSectionFromMemo(GetArgumentString(S, 'SECTION'), FMemoOrigine, True);
  end;

  TMemo(GetControl('BN1')).ReadOnly := bnReadOnly in Properties;
  SetControlVisible('BVALIDER', not (bnReadOnly in Properties));

  { Evénements }
  if GetControl('BVALIDER') <> nil then
    TToolBarButton97(GetControl('BVALIDER')).OnClick := BVALIDER_OnClick;
  if Assigned(GetControl('BTSearch')) then
  begin
    TToolBarButton97(GetControl('BTSearch')).OnClick := BTSearch_OnClick;
    Ecran.OnKeyDown := Form_KeyDown
  end;
  FSearch := TMemoFindDialog.Create(Ecran);
  FSearch.OnFound := SetMemoCoord;
  FSearch.Memo := TMemo(GetControl('BN1'));
  FSearch.Ecran := Ecran;

  { Visibilité du panel infos. }
  SetControlVisible('PNINFO', PropFont);
  if PropFont then
    BlocNoteSize := Length(TMemo(GetControl('BN1')).Text)
end;

procedure TOF_WBLOCNOTE.OnClose ;
begin
  Inherited;
end;

procedure TOF_WBLOCNOTE.BVALIDER_OnCLick(Sender: TObject);
var
	Sql: string;
	theTob: tob;

  procedure SetMemo;

    function GetMemoSection: String;
    begin
      Result := '[' + GetArgumentString(StArgument, 'SECTION') + ']'+#13#10
              + TMemo(GetControl('BN1')).Text + #13#10#13#10 + Trim(FMemoOrigine)
    end;

  begin
    if bnSection in Properties then
      TMemo(GetControl('BN1')).Text := GetMemoSection();
    if bnZipped in Properties then
      TMemo(GetControl('BN1')).Text := wZipString(TMemo(GetControl('BN1')).Text);
  end;

  procedure wSetTobMemo(t: Tob);
  begin
    SetMemo();
    t.SetString(Prefixe + '_' + fBlobSuffixe, TMemo(GetControl('BN1')).Text);
    t.SetBoolean(Prefixe + '_WBMEMO', t.GetString(Prefixe + '_' + fBlobSuffixe) <> '');
    t.UpDateDb;
  end;

begin
  if (bnObligatoire in Properties) and (Trim(TMemo(GetControl('BN1')).Text) = '') then
  begin
    LastError := ErrOblig;
    if LastError > 0 then
    begin
      case LastError of
        ErrOblig: PGIError('Vous devez saisir un bloc-notes.', Ecran.Caption);
      end;
    end;
  end;

  if not (bnReadOnly in Properties) and (LastError = 0) then
  begin
    if bnByFile in Properties then
    begin
      if FileExists(FilePath) then
      begin
        SetMemo();
        TMemo(GetControl('BN1')).Lines.SaveToFile(FilePath)
      end
    end
    else if bnByString in Properties then
    begin
      SetMemo();
      TFVierge(Ecran).Retour := TMemo(GetControl('BN1')).Text
    end
    else if bnBySQL in Properties then
    begin
      Sql := 'SELECT *'
           + ' FROM ' + PrefixeToTable(Prefixe)
           + ' WHERE ' + Where
           ;
      theTob := Tob.Create('BLOCNOTE', nil, -1);
      try
        if TheTob.LoadDetailDBFromSql(PrefixeToTable(Prefixe), Sql) then
          wSetTobMemo(TheTob.Detail[0]);
      finally
        theTob.free;
      end
    end
    else if bnByTob in Properties then
    begin
      if Assigned(TobSt) and TobSt.FieldExists(Prefixe + iif(Prefixe <> '', '_', '') + fBlobSuffixe) then
      begin
        SetMemo();
        TobSt.SetString(Prefixe + iif(Prefixe <> '', '_', '') + fBlobSuffixe, TMemo(GetControl('BN1')).Text)
      end
    end
  end
end;

procedure TOF_WBLOCNOTE.BTSearch_OnClick(Sender: TObject);
begin
  FSearch.Execute
end;

procedure TOF_WBLOCNOTE.Form_KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = vk_Cherche) and (ssCtrl	in Shift) then
    TToolBarButton97(GetControl('BTSearch')).Click
  else if (Key = VK_F3) and (FSearch.StringToFind <> '') then
  begin
    FSearch.FindText := FSearch.StringToFind;
    FSearch.FindNext()
  end
  else
    TFVierge(Ecran).FormKeyDown(Sender, Key, Shift)
end;

procedure TOF_WBLOCNOTE.SetMemoCoord(Sender: TObject);
begin
  with TMemo(Sender) do
  begin
    Coords := 'Lig. '  + IntToStr(CaretPos.Y + 1) + ' : ' + IntToStr(Lines.Count + 1)
            + '  Col. ' + IntToStr(CaretPos.X + 1)
            + '  Sél. ' + IntToStr(SelLength);
    BlocNoteSize := Length(Lines.Text)
  end
end;

procedure TOF_WBLOCNOTE.Memo_Click(Sender: TObject);
begin
  SetMemoCoord(Sender)
end;

procedure TOF_WBLOCNOTE.Memo_KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  SetMemoCoord(Sender)
end;

procedure TOF_WBLOCNOTE.Memo_KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  SetMemoCoord(Sender)
end;

function TOF_WBLOCNOTE.GetCoords: String;
begin
  Result := TPanel(GetControl('PNCOORD')).Caption
end;

procedure TOF_WBLOCNOTE.SetCoords(const Value: String);
begin
  TPanel(GetControl('PNCOORD')).Caption := Value
end;

function TOF_WBLOCNOTE.GetBlocNoteSize: Integer;
begin
  Result := Length(TMemo(GetControl('BN1')).Text)
end;

procedure TOF_WBLOCNOTE.SetBlocNoteSize(const Value: Integer);
begin
  TPanel(GetControl('PNSIZE')).Caption := wFileSizeToStr(Value, 2)
end;

Initialization
  RegisterClasses([TOF_WBLOCNOTE]);
end.
