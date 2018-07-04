{***********UNITE*************************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : WMULRECHERCHE ()
Mots clefs ... : TOF;WMULRECHERCHE
*****************************************************************}
Unit WMULRECHERCHE_TOF ;

Interface

Uses
  StdCtrls
  ,Controls
  ,Classes
  {$IFNDEF EAGLCLIENT}
    ,Mul
    ,db
    {$IFNDEF DBXPRESS},dbtables{BDE}{$ELSE},uDbxDataSet{$ENDIF}
    ,Fe_Main
  {$ELSE}
    ,eMul
    ,MainEAgl
  {$ENDIF}
  ,forms
  ,sysutils
  ,ComCtrls
  ,HCtrls
  ,HEnt1
  ,HMsgBox
  ,wCommuns
  ,UTOF
  ,WTOF
  ,uTob
  ;

Type
  TOF_WMULRECHERCHE = Class(twTOF)
    procedure OnArgument(S: String); override;
    procedure OnDisplay            ; override;
    procedure OnNew                ; override;
    procedure OnLoad               ; override;
    procedure OnUpdate             ; override;
    procedure OnDelete             ; override;
    procedure OnCancel             ; override;
    procedure OnClose              ; override;
  private
    procedure BtValider_OnClick(Sender: TObject);
    procedure Mul_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListe_OnDblClick(Sender: TObject);
  public

  end;

  TMulRecherche = class
  private
    FFieldsToReturn : Array of String;
    FTobResult      : Tob;
    FDBListe, FWhere,
    FArguments,
    FsResult        : String;
    FPerso          : Boolean;
  public
    constructor Create(const DBListe: String; const FieldsToReturn: Array of String; const Caption: String; const Where: String = '');
    destructor Destroy; override;
    procedure Resize(const AHeight, AWidth: Integer);
    procedure Position(const ATop, ALeft: Integer);
    function Execute(const MultiSelection: Boolean = False): Boolean;
    property TobResult: Tob read FTobResult;
    property sResult: String read FsResult write FsResult;
    property Personalisable: Boolean read FPerso write FPerso default True;
  end;

Implementation

uses
  HTB97
  ,Windows
  ,UtilPGI
  ;

procedure TOF_WMULRECHERCHE.OnArgument(S: String);
var
  STmp: String;
  MultiSelection: Boolean;
begin
  Inherited;

  { Caption }
  Ecran.Caption := Tob(GetArgumentInteger(StArgument, 'TOBRESULT')).GetString('CAPTION');

  { XX_Where }
  SetControlText('XX_WHERE', GetArgumentValue(S, 'XX_WHERE'));

  { Multi sélection }
  MultiSelection := GetArgumentBoolean(S, 'MULTISELECTION');
  {$IFDEF EAGLCLIENT}
    TFMul(Ecran).FListe.MultiSelect := MultiSelection;
  {$ELSE}
    TFMul(Ecran).FListe.MultiSelection := MultiSelection;
  {$ENDIF EAGLCLIENT}
  SetControlVisible('BSELECTALL', MultiSelection);
  SetControlVisible('DOCK971', GetArgumentBoolean(S, 'PERSO'));

  { Liste }
  wChangeDBListe(TFMul(Ecran), GetArgumentString(S, 'DBLISTE'));

  { Resize }
  STmp := GetArgumentString(S, 'RESIZE');
  if STmp <> '' then
  begin
    TFMul(Ecran).Height := ValeurI(ReadTokenPipe(STmp, '~'));
    TFMul(Ecran).Width  := ValeurI(ReadTokenPipe(STmp, '~'));
  end;

  { Position }
  STmp := GetArgumentString(S, 'POSITION');
  if STmp <> '' then
  begin
    TFMul(Ecran).Top  := ValeurI(ReadTokenPipe(STmp, '~'));
    TFMul(Ecran).Left := ValeurI(ReadTokenPipe(STmp, '~'));
  end;

  { Evénements }
  if Assigned(GetControl('BOuvrir')) then
    TToolBarButton97(GetControl('BOuvrir')).OnClick := BtValider_OnClick;
  if Assigned(TFMul(Ecran).FListe) then
  begin
    TFMul(Ecran).FListe.OnDblClick := FListe_OnDblClick;
    TFMul(Ecran).FListe.OnKeyDown := Mul_OnKeyDown
  end
end;

procedure TOF_WMULRECHERCHE.OnDisplay();
begin
  Inherited;
end;

procedure TOF_WMULRECHERCHE.OnNew;
begin
  Inherited;
end;

procedure TOF_WMULRECHERCHE.OnLoad;
begin
  Inherited;

  TFMul(Ecran).Pages.ActivePage := TFMul(Ecran).PAvance;
end;

procedure TOF_WMULRECHERCHE.OnUpdate;
begin
  Inherited;
end;

procedure TOF_WMULRECHERCHE.OnDelete;
begin
  Inherited;
end;

procedure TOF_WMULRECHERCHE.OnCancel();
begin
  Inherited;
end;

procedure TOF_WMULRECHERCHE.OnClose;
begin
  Inherited;
end;

{ Composition de la tob de retour }
procedure TOF_WMULRECHERCHE.BtValider_OnClick(Sender: TObject);
var
  s, sField: String;
  Q: TQuery;
  {$IFDEF EAGLCLIENT}
    wBookMark: Integer;
  {$ELSE}
    wBookMark: TBookMark;
  {$ENDIF EAGLCLIENT}
  i: Integer;

  procedure AddToTob;
  begin
    s := GetArgumentString(StArgument, 'FIELDS');
    with Tob.Create('_Result_', Tob(GetArgumentInteger(StArgument, 'TOBRESULT')), -1) do
    begin
      while s <> '' do
      begin
        sField := ReadTokenPipe(s, '~');
        if not wMultiSelected(TFMul(Ecran).FListe) then
          AddChampSupValeur(sField, Self.GetString(sField))
        else
          AddChampSupValeur(sField, Q.FindField(sField).AsString)
      end
    end
  end;

begin
  if wMultiSelected(TFMul(Ecran).FListe) then
  begin
    {$IFNDEF EAGLCLIENT}
      Q := TFMul(Ecran).Q;
    {$ELSE}
      Q := TFMul(Ecran).Q.TQ;
    {$ENDIF NEAGLCLIENT}

    Q.DisableControls;
    try
      {$IFDEF EAGLCLIENT}
        wBookMark := TFMul(Ecran).FListe.Row;
        if TFMul(Ecran).FListe.AllSelected then
        begin
          Q.First;
          for i := 1 to Q.recordCount do
          begin
            AddToTob;
            Q.next;
          end;
        end
        else
        begin
          Q.First;
          for i := 1 to Q.recordCount do
          begin
            if TFMul(Ecran).FListe.IsSelected(i) then
              AddToTob;
            Q.next;
          end;
        end;
        TFMul(Ecran).FListe.Row := wBookMark;
        if TFMul(Ecran).FListe.Row>0 then
          Q.Seek(TFMul(Ecran).FListe.Row - 1)
        else
          Q.First;
      {$ELSE}
        wBookMark := TFMul(Ecran).FListe.DataSource.DataSet.GetBookmark;
        if TFMul(Ecran).FListe.AllSelected then
        begin
          Q.First;
          for i := 1 to Q.recordCount do
          begin
            AddToTob;
            Q.next;
          end;
        end
        else
        begin
          for i := 0 to TFMul(Ecran).FListe.NbSelected - 1 do
          begin
            TFMul(Ecran).FListe.GotoLeBOOKMARK(i);
            AddToTob;
          end;
        end;
        Q.GotoBookmark(wBookMark);
      {$ENDIF EAGLCLIENT}
    finally
      Q.EnableControls;
    end
  end
  else
    AddToTob;

  TToolBarButton97(GetControl('BAnnuler')).Click
end;

{ Validation ligne par touche ENTREE }
procedure TOF_WMULRECHERCHE.Mul_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_RETURN: TToolBarButton97(GetControl('BOuvrir')).Click
  end
end;

{ Validation ligne par double clique }
procedure TOF_WMULRECHERCHE.FListe_OnDblClick(Sender: TObject);
begin
  TToolBarButton97(GetControl('BOuvrir')).Click
end;

{ TMulRecherche }

constructor TMulRecherche.Create(const DBListe: String; const FieldsToReturn: Array of String; const Caption: String; const Where: String = '');
var
  i: Integer;
begin
  FTobResult := Tob.Create('_MulRecherche_', nil, -1);
  FTobResult.AddChampSupValeur('CAPTION', Caption);
  FDBListe := DBListe;
  FWhere := Where;
  FsResult := '';
  FPerso := True;
  SetLength(FFieldsToReturn, Length(FieldsToReturn));
  for i := Low(FieldsToReturn) to High(FieldsToReturn) do
    FFieldsToReturn[i] := FieldsToReturn[i];
end;

destructor TMulRecherche.Destroy;
begin
  inherited;

  FTobResult.Free
end;

function TMulRecherche.Execute(const MultiSelection: Boolean = False): Boolean;
var
  sFields: String;
  i, j   : Integer;
begin
  if Assigned(FTobResult) then
  begin
    sFields := '';
    for i := Low(FFieldsToReturn) to High(FFieldsToReturn) do
      sFields := sFields + iif(sFields <> '', '~', '') + FFieldsToReturn[i];

    AGLLanceFiche('W', 'WMULRECHERCHE_MUL', '', '', iif(FWhere = '', '', 'XX_WHERE=' + FWhere + ';')
                                                  + 'ACTION=MODIFICATION;DBLISTE='
                                                  + FDBListe + ';MULTISELECTION='
                                                  + booltostr_(MultiSelection)
                                                  + ';FIELDS=' + sFields
                                                  + ';TOBRESULT=' + IntToStr(LongInt(FTobResult))
                                                  + ';PERSO=' + booltostr_(FPerso)
                                                  + ';' + FArguments);

  end;
  Result := FTobResult.Detail.Count > 0;
  if Result and not MultiSelection then
    with FTobResult do
      for i := 0 to Pred(Detail.Count) do
        for j := Low(FFieldsToReturn) to High(FFieldsToReturn) do
          FsResult := FsResult + iif(FsResult <> '', ';', '') + Detail[i].GetString(FFieldsToReturn[j]);
end;

procedure TMulRecherche.Position(const ATop, ALeft: Integer);
begin
  FArguments := FArguments  + 'POSITION=' + IntToStr(ATop) + '~' + IntToStr(ALeft) + ';'
end;

procedure TMulRecherche.Resize(const AHeight, AWidth: Integer);
begin
  FArguments := FArguments  + 'RESIZE=' + IntToStr(AHeight) + '~' + IntToStr(AWidth) + ';'
end;

Initialization
  RegisterClasses([TOF_WMULRECHERCHE]);
end.
