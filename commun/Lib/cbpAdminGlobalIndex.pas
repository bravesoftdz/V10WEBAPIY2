unit cbpAdminGlobalIndex;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, HEnt1, ComCtrls, ImgList, jpeg, HPanel,
  HTB97, Hctrls, CheckLst, uGlobalIndexPropertyEditor, Buttons, Menus,
  HSysMenu, Spin, HImgList, CommCtrl, HPdfviewer, cbpGlobalIndex, GraphicEx, hmsgbox;

type
  TOutTextEvent = procedure(Msg: string) of object;

  TFAdminGlobalIndex = class(TForm)
    EditFileName: TEdit;
    ButtonParser: TButton;
    Memo: TMemo;
    TreeView: TTreeView;
    LabelCount: TLabel;
    BtnGlobalIndex: TButton;
    SaveDialog: TSaveDialog;
    BtnSave: TButton;
    EditSavePath: TEdit;
    Informations: TGroupBox;
    BtnSaveDll: TButton;
    ListForms: TListBox;
    Label1: TLabel;
    ImageGlyph: TImage;
    Button1: TButton;
    MemoKeyWord: TMemo;
    clbKeyWords: THCheckListBox;
    Label3: TLabel;
    Image1: TImage;
    HMTrad: THSystemMenu;
    EditLot: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    HSpinEdit: THSpinEdit;
    OpenDialog: TOpenDialog;
    GroupBox1: TGroupBox;
    BtnSaveDB: TButton;
    ButtonLoadDB: TButton;
    GroupBox2: TGroupBox;
    BtnSaveFile: TButton;
    BtnLoadFile: TButton;
    MemoForms: TMemo;
    ImageList1: TImageList;
    BtnLoadImage: TButton;
    BtnSaveDllVista: TButton;
    ToolbarButton971: TToolbarButton97;
    Panel1: TPanel;
    HSpeedButton1: THSpeedButton;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Button2: TButton;
    HImageList1: THImageList;

    procedure ButtonParserClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure ButtonLoadDBClick(Sender: TObject);
    procedure BtnGroupClick(Sender: TObject);
    procedure DfmGly(Sender: TCustomTreeView;
      Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure RefreshTreeView();
    procedure BtnSaveDBClick(Sender: TObject);
    procedure BtnClearDBClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BtnGlobalIndexClick(Sender: TObject);
    procedure BtnSaveDllClick(Sender: TObject);
    procedure TreeViewChange(Sender: TObject; Node: TTreeNode);
    procedure Button1Click(Sender: TObject);
    procedure MemoKeyWordKeyPress(Sender: TObject; var Key: Char);
    procedure MemoKeyWordExit(Sender: TObject);
    procedure clbKeyWordsClickCheck(Sender: TObject);
    procedure btnUpdateKeyWordsClick(Sender: TObject);
    procedure tbbTestClick(Sender: TObject);
    procedure HSpinEditChange(Sender: TObject);
    procedure EditLotChange(Sender: TObject);
    procedure HBitBtn1Click(Sender: TObject);
    procedure BtnSaveFileClick(Sender: TObject);
    procedure BtnLoadFileClick(Sender: TObject);
    procedure HSpeedButton1Click(Sender: TObject);
    procedure btn256Click(Sender: TObject);
    procedure BtnLoadImageClick(Sender: TObject);
    procedure BtnSaveDllVistaClick(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);

  private
    FDfmGlyphs: TDfmGlyphs;
    FImageLists: TMyImageLists;
    FGlobalKeyWords: TGlobalKeyWords;

    FIndex: Integer;
    FTreeChangeEnabled: Boolean;

    procedure OutText(str: string);
    function CopyFile(srcFileName, destFileName: string): Boolean;

    procedure UpdateKeyWords(KeyWords: string; UpdateList: Boolean = False);
    procedure BuildKeyWords;
  public

  end;

var
  FParseDfm: TFAdminGlobalIndex;

implementation

{$R *.dfm}

uses HDebug, uTOB,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
DB,
HQry, FileResourceMgr, uGedFiles, HDrawXP, HZStream, MajTable, ADODB, CegidPath,
TeamParseDfm, Graphics32, StrUtils;


//function ReadToken(var Buffer: string): string;
function ReadToken(Stream: TStream): string;
var
  State: Integer;
  c, c1: Char;
  Pos: Integer;
  WideStr: string;
begin
  State := 0;
  Pos := 1;
  Result := '';

  while (State <> 10) and (Stream.Position < Stream.Size) do
  begin
    Stream.ReadBuffer(c, Sizeof(c));

    case State of
      0 :
        begin
          if c in [' ', #13, #10] then
          else if c = '''' then
          begin
            State := 2;
            Result := c;
          end
          else
          begin
            Result := c;
            State := 1;
          end;
        end;

      1 :
        begin
          if c in [' ', #13, #10] then
          begin
            State := 10;
          end
          else
            Result := Result + c;
        end;

      2:
        begin
          if Stream.Position < Stream.Size then
          begin
            Stream.ReadBuffer(c1, Sizeof(c1));
            Stream.Position := Stream.Position - 1;
          end
          else
            c1 := '-';

          if (c = '''') and (c1 = '#') then
          begin
            State := 3;
            WideStr := '';
            Stream.Position := Stream.Position + 1;
          end
          else if (c = '''') and (c1 = '''')  then
          begin
            Result := Result + c;
            Stream.Position := Stream.Position + 1;
          end
          else if c = '''' then
          begin
            State := 10;
            Result := Result + c;
          end
          else
            Result := Result + c;
        end;

      3:
        begin
          if (c = '''') or (c = #13) then
          begin
            Result := Result + Char(ValeurI(WideStr));
            State := 2;
          end
          else if c in ['0'..'9'] then
            WideStr := WideStr + c
          else
            State := 2;
        end;
    end;
  end;
end;

procedure TFAdminGlobalIndex.ButtonParserClick(Sender: TObject);
begin
  Screen.Cursor := crHourGlass;

  {$IFNDEF EAGL}
  FDfmGlyphs.ParseDfmArguments(EditFileName.Text);
  {$ENDIF}
  FDfmGlyphs.SortByGlobalIndex;
  RefreshTreeView();

  Screen.Cursor := crDefault;
end;

procedure TFAdminGlobalIndex.FormCreate(Sender: TObject);
begin
  FDfmGlyphs := TDfmGlyphs.Create(Self);
  FDfmGlyphs.OnOutText := OutText;

//  FImageLists := TMyImageLists.Create;
  FGlobalKeyWords := TGlobalKeyWords.Create;

  FTreeChangeEnabled := True;
end;

function StringToVarArray(const Value: string): OleVariant;
var
  PData: Pointer;
  Size: Integer;
begin
  Size := Length(Value);
  Result := VarArrayCreate([0, Size-1], varByte);
  PData := VarArrayLock(Result);
  try
    Move(Pointer(Value)^, PData^, Size);
  finally
    VarArrayUnlock(Result);
  end;
end;

procedure TFAdminGlobalIndex.BtnSaveClick(Sender: TObject);
var
  i: Integer;
begin
{
  Screen.Cursor := crHourGlass;

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if FDfmGlyphs.Items[i].FRefGUID = '' then
    begin
      FDfmGlyphs.Items[i].Bitmap.SaveToFile(EditSavePath.Text + '\' + FDfmGlyphs.Items[i].FGlobalIndex + '.bmp');
    end;
  end;

  Screen.Cursor := crDefault; }
end;

procedure TFAdminGlobalIndex.BuildKeyWords();
var
  i: Integer;
begin
  FGlobalKeyWords.Clear;

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if (FDfmGlyphs[i].RefGUID = '') and (FDfmGlyphs[i].KeyWords <> '')  then
    begin
      FGlobalKeyWords.AddKeyWords(FDfmGlyphs[i].KeyWords, FDfmGlyphs[i].GUID, FDfmGlyphs[i].GlobalIndex);
    end;
  end;
end;

procedure TFAdminGlobalIndex.ButtonLoadDBClick(Sender: TObject);
begin
  Screen.Cursor := crHourGlass;
  FDfmGlyphs.LoadDB;
  BuildKeyWords;
  RefreshTreeView;
  Screen.Cursor := crDefault;
end;

procedure TFAdminGlobalIndex.OutText(str: string);
begin
  Str := DateTimeToStr(Now) + ' : ' + Str;
  Memo.Lines.Add(str);
  Application.ProcessMessages;
end;

procedure TFAdminGlobalIndex.BtnGroupClick(Sender: TObject);
var
  i, j: Integer;
  b1, b2: TBitmap;
  p1, p2 :PByteArray;
  Variance: Integer;
  r, c: Integer;
begin
{
  Screen.Cursor := crHourGlass;

  OutText('Début Grouper les images');
  // efface toutes les référencers

  FDfmGlyphs.Sort(SortGlobalIndex);

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    FDfmGlyphs[i].FRefGUID := '';
  end;

  // recherche des images identitiques
  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if FDfmGlyphs[i].FRefGUID = '' then
    begin
      FDfmGlyphs[i].FCountChild := 0;

      for j := i + 1 to FDfmGlyphs.Count - 1 do
      begin
        if (FDfmGlyphs[j].FRefGUID = '') and (FDfmGlyphs[i].FHCode = FDfmGlyphs[j].FHCode) and (FDfmGlyphs[i].FImgHex = FDfmGlyphs[j].FImgHex) then
        begin
          FDfmGlyphs[j].FRefGUID := FDfmGlyphs[i].FGUID;
          FDfmGlyphs[i].FCountChild := FDfmGlyphs[i].FCountChild + 1;
        end;
      end;
    end;

    if (i mod 10000) = 9999 then
      OutText('Grouper les images identiques à ' + IntToStr(i * 100 div FDfmGlyphs.Count) + '%');
  end;

  // recherche des images semblable
  b1 := TBitmap.Create;
  b2 := TBitmap.Create;

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if FDfmGlyphs[i].FRefGUID = '' then
    begin
      b1.Assign(FDfmGlyphs[i].Bitmap);
      b1.PixelFormat := pf24bit;

      for j := i + 1 to FDfmGlyphs.Count - 1 do
      begin
        if (FDfmGlyphs[j].FRefGUID = '') and (FDfmGlyphs[j].Width = FDfmGlyphs[i].Width) and (FDfmGlyphs[j].Height = FDfmGlyphs[i].Height) then
        begin
          b2.Assign(FDfmGlyphs[j].Bitmap);
          b2.PixelFormat := pf24bit;
          Variance := 0;

          for r := 0 to FDfmGlyphs[i].Height - 1 do
          begin
            p1 := b1.ScanLine[r];
            p2 := b2.ScanLine[r];

            for c := 0 to (FDfmGlyphs[i].Width - 1) * 3 do
              Variance := variance + (p1[c] xor p2[c]);
          end;

          if Variance < 5000 then
          begin
            FDfmGlyphs[j].FRefGUID := FDfmGlyphs[i].FGUID;
            FDfmGlyphs[i].FCountChild := FDfmGlyphs[i].FCountChild + 1;
          end;
        end;
      end;
    end
    else
      FDfmGlyphs[i].ReleaseBitmap; //pour libérer la ressource image sur les doublons

    if (i mod 10000) = 9999 then
      OutText('Grouper les images semblables à ' + IntToStr(i * 100 div FDfmGlyphs.Count) + '%');
  end;

  b1.Free;
  b2.Free;

  OutText('Nb : ' + IntToStr(FDfmGlyphs.Count) + ' Distinct : ' + IntToStr(FDfmGlyphs.CountDistinct));

  Screen.Cursor := crDefault;

  OutText('Fin Grouper les images');
}
end;

function PixelFormatToStr(PixelFormat: Integer): string;
begin
  Result := '';

  case PixelFormat of
    0 : Result := 'pfDevice';
    1 : Result := 'pf1bit';
    2 : Result := 'pf4bit';
    3 : Result := 'pf8bit';
    4 : Result := 'pf15bit';
    5 : Result := 'pf16bit';
    6 : Result := 'pf24bit';
    7 : Result := 'pf32bit';
    8 : Result := 'pfCustom';
  end;
end;


procedure TFAdminGlobalIndex.DfmGly(Sender: TCustomTreeView;
  Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
var
  BmpRect, NodeRect: TRect;
  DfmGlyph: TDfmGlyph;
  b32: TBitmap32;
begin

  DefaultDraw := False;
  NodeRect := Node.DisplayRect(False);
//  TreeView.Canvas.Rectangle(NodeRect.Left, NodeRect.Top, NodeRect.Right, NodeRect.Bottom);
  BmpRect := NodeRect;
  BmpRect.Right := BmpRect.Left + 32;
  BmpRect.Bottom := BmpRect.Top + 32;

  try
    DfmGlyph := TDfmGlyph(Node.Data);

    if DfmGlyph.ImageType = 'BMP' then
    begin
      DfmGlyph.Bitmap.Transparent := True;
      Sender.Canvas.StretchDraw(BmpRect, DfmGlyph.Bitmap);
    end
    else if DfmGlyph.ImageType = 'PNG' then
    begin
      b32 := TBitmap32.Create;
      b32.Assign(DfmGlyph.Bitmap);
      b32.PreMultAlpha;
      b32.DrawBlend(sender.Canvas.Handle, BmpRect);
      b32.Free;
    end
    else
      Sender.Canvas.StretchDraw(BmpRect, DfmGlyph.Bitmap);

    NodeRect.Left := NodeRect.Left + 36;
    Sender.Canvas.TextRect(NodeRect, NodeRect.Left, NodeRect.Top,
    DfmGlyph.GlobalIndex + ' T:' + DfmGlyph.ImageType + ' S:' + IntToStr(DfmGlyph.Bitmap.Height) + ' P:' + PixelFormatToStr(DfmGlyph.PixelFormat) + ' C:' + IntToStr(DfmGlyph.CountChild));
  except
  end;
end;

procedure TFAdminGlobalIndex.RefreshTreeView();
var
  i, j: Integer;
begin
  Screen.Cursor := crHourGlass;

  FDfmGlyphs.SortByGlobalIndex;

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin

    if FDfmGlyphs.Items[i].RefGUID = '' then
    begin
      FDfmGlyphs.Items[i].CountChild := 0;

      for j := i + 1 to FDfmGlyphs.Count - 1 do
      begin
        if FDfmGlyphs.Items[j].RefGUID = FDfmGlyphs.Items[i].GUID then
        begin
          FDfmGlyphs.Items[i].CountChild := FDfmGlyphs.Items[i].CountChild + 1;
        end
        else
          Break;
      end;

//      TreeView.Items.AddChildObject(nil, FDfmGlyphs.Items[i].FGUID, Pointer(FDfmGlyphs.Items[i]));
    end;
  end;

  TreeView.Items.Clear;
  TreeView.Perform(TVM_SETITEMHEIGHT, 34, 0);
  TreeView.Canvas.Font.Size := 8;

  TreeView.Items.BeginUpdate;

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if FDfmGlyphs.Items[i].RefGUID = '' then
    begin
      TreeView.Items.AddChildObject(nil, FDfmGlyphs.Items[i].GUID, Pointer(FDfmGlyphs.Items[i]));
    end;
  end;

  TreeView.Items.EndUpdate;

  LabelCount.Caption := 'Count : ' + IntToStr(TreeView.Items.Count);

  Screen.Cursor := CrDefault;
end;

procedure TFAdminGlobalIndex.BtnSaveDBClick(Sender: TObject);
begin

//  ParserDfmFile('C:\AglEncours5\TRAVAIL\PGR\AGL7\MAC7\uParseDfm.dfm');

  Screen.Cursor := crHourGlass;
  FDfmGlyphs.SaveDB;
  Screen.Cursor := crDefault;
end;

procedure TFAdminGlobalIndex.BtnClearDBClick(Sender: TObject);
begin
//  FDfmGlyphs.Clear;
end;

procedure TFAdminGlobalIndex.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FTreeChangeEnabled := False;

  TreeView.Items.Clear;

  FDfmGlyphs.Free;
//  FImageLists.Free;
  FGlobalKeyWords.Free;

  AglDepileFiche ;
  if IsInside(Self) then Action:=caFree;
end;

procedure TFAdminGlobalIndex.BtnGlobalIndexClick(Sender: TObject);
var
  i:  Integer;
  Index: Integer;
  NumGlyph: Integer;
begin
{
  Screen.Cursor := crHourGlass;

  Index := 1;

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if (FDfmGlyphs[i].FRefGUID = '') and ((FDfmGlyphs[i].FGlobalIndex = '') or (FDfmGlyphs[i].FGlobalIndex[1] = 'Z')) then
    begin
      if FDfmGlyphs[i].Height < 24 then
      begin
        if FDfmGlyphs[i].Height <> 0 then
          NumGlyph := FDfmGlyphs[i].Width div FDfmGlyphs[i].Height
        else
          NumGlyph := 1;

        FDfmGlyphs[i].FGlobalIndex := 'Z' + FindEtReplace(format('%0:4d', [Index]), ' ', '0', True) + '_S16G'+ IntToStr(NumGlyph);
        Index := Index + 1;
      end
      else if FDfmGlyphs[i].Height < 32 then
      begin
        if FDfmGlyphs[i].Height <> 0 then
          NumGlyph := FDfmGlyphs[i].Width div FDfmGlyphs[i].Height
        else
          NumGlyph := 1;

        FDfmGlyphs[i].FGlobalIndex := 'Z' + FindEtReplace(format('%0:4d', [Index]), ' ', '0', True) + '_S24G'+ IntToStr(NumGlyph);
        Index := Index + 1;
      end
      else if FDfmGlyphs[i].Height < 48 then
      begin
        if FDfmGlyphs[i].Height <> 0 then
          NumGlyph := FDfmGlyphs[i].Width div FDfmGlyphs[i].Height
        else
          NumGlyph := 1;

        FDfmGlyphs[i].FGlobalIndex := 'Z' + FindEtReplace(format('%0:4d', [Index]), ' ', '0', True) + '_S32G'+ IntToStr(NumGlyph);
        Index := Index + 1;
      end
      else
      begin
        NumGlyph := 1;
        FDfmGlyphs[i].FGlobalIndex := 'Z' + FindEtReplace(format('%0:4d', [Index]), ' ', '0', True) +'_SOG'+ IntToStr(NumGlyph);
        Index := Index + 1;
      end
    end;
  end;

  TreeView.Refresh;

  Screen.Cursor := crDefault;
}
end;

function TFAdminGlobalIndex.CopyFile(srcFileName, destFileName: string): Boolean;
begin
  if not windows.CopyFile(PChar(srcFileName), PChar(destFileName), True) then
  begin
    Result := False;
  end
  else
    Result := True;
end;

procedure TFAdminGlobalIndex.BtnSaveDllClick(Sender: TObject);
var
  Path: String;
  i: Integer;
  FrMgr: TFileResourceMgr;
  ms: TMemoryStream;
begin
  Screen.Cursor := crHourGlass;

  TGlobalIndex.FreeFileGlobalImageRes;

  Path := ExtractFilePath(Application.ExeName);

  if FileExists(Path + 'AglRes2000.dll') then
    DeleteFile(Path + 'AglRes2000.dll');

  CopyFile(Path + 'AglResTemplate.dll', Path + 'AglRes2000.dll');

  FrMgr := TFileResourceMgr.Create(Path + 'AglRes2000.dll');

  try
    for i := 0 to FDfmGlyphs.Count - 1 do
    begin
      if (FDfmGlyphs.Items[i].RefGUID = '') and (FDfmGlyphs.Items[i].ImageType = 'OLD') then
      begin
        ms := TMemoryStream.Create;

        try
          FDfmGlyphs.Items[i].Bitmap.SaveToStream(ms);
          ms.Position := 0;

          if FrMgr.WriteResourceTryed(10, rtRCData, FDfmGlyphs.Items[i].GlobalIndex, ms) <> 0 then
          begin
            PGIInfo('Echec de la génération !');
            Exit;
          end;
        finally
          ms.Free;
        end;
      end;
    end;

    //Sauvegarde des mots clefs

    ms := TMemoryStream.Create;
    FGlobalKeyWords.SaveToStream(ms);
    ms.Position := 0;
    FrMgr.WriteResource(rtRCData, 'GLOBALKEYWORDS', ms);
    ms.Free;
  finally
    FrMgr.Free;
  End;

  Screen.Cursor := crDefault;
end;

procedure TFAdminGlobalIndex.TreeViewChange(Sender: TObject; Node: TTreeNode);
var
  DfmGlyph: TDfmGlyph;
  ListForms: TStringList;
  i: Integer;

  procedure AddItem(const ADfmGlyph: TDfmGlyph);
  begin
//  ListForms.Items.Add(ADfmGlyph.FFileName + ';   ' + ADfmGlyph.FObjectName + ';   ' + ADfmGlyph.FObjectType);
    ListForms.Add(ADfmGlyph.FileName + ';   ' + ADfmGlyph.ObjectName + ';   ' + ADfmGlyph.ObjectType);
  end;

begin

  if not FTreeChangeEnabled then Exit;

  DfmGlyph := TDfmGlyph(Node.Data);

  Screen.Cursor := crHourGlass;

  ListForms := TStringList.Create;
  AddItem(DfmGlyph);

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if FDfmGlyphs[i].RefGUID = DfmGlyph.GUID then
      AddItem(FDfmGlyphs[i]);
  end;

  ListForms.Sort;

  MemoForms.Lines := ListForms;
  ListForms.Free;

//  MemoForms.Lines.Sorted := true;

  ImageGlyph.Picture.Bitmap := DfmGlyph.Bitmap;
  UpdateKeyWords(DfmGlyph.KeyWords);

  HSpinEdit.Value := DfmGlyph.Level;
  EditLot.Text := DfmGlyph.Lot;

  Screen.Cursor := crDefault;
end;

procedure TFAdminGlobalIndex.Button1Click(Sender: TObject);
var
  Path: string;
begin
  Path := ExtractFilePath(Application.ExeName) + 'ParseDfm.log';

  if FileExists(Path) then
    DeleteFile(Path);

  Memo.Lines.SaveToFile(Path);
end;

procedure TFAdminGlobalIndex.MemoKeyWordKeyPress(Sender: TObject; var Key: Char);
begin
  if Key in ['a'..'z'] then
    Dec(Key, 32)
  else if Key in ['A'..'Z', ';', #8] then
  else if Key = #13 then
  begin
    MemoKeyWordExit(Sender);
    Key := #0
  end
  else
  begin
    Key := #0;
    Beep;
  end;
end;

procedure TFAdminGlobalIndex.UpdateKeyWords(KeyWords: string; UpdateList: Boolean = False);
var
  iKey: Integer;
begin
  MemoKeyWord.Text := KeyWords;

  if UpdateList or (clbKeyWords.Count = 0) or (FGlobalKeyWords.Count <> clbKeyWords.Count) then
  begin
    clbKeyWords.Clear;
    FGlobalKeyWords.Sort;

    for iKey := 0 to FGlobalKeyWords.Count - 1 do
    begin
      clbKeyWords.Items.Add(FGlobalKeyWords[iKey].Key);

      if Pos(FGlobalKeyWords[iKey].Key, KeyWords) > 0 then
        clbKeyWords.Checked[clbKeyWords.Count-1] := True;
    end;
  end
  else
  begin
    for iKey := 0 to FGlobalKeyWords.Count - 1 do
      clbKeyWords.Checked[iKey] := (Pos(FGlobalKeyWords[iKey].Key, KeyWords) > 0);
  end;
end;

procedure TFAdminGlobalIndex.MemoKeyWordExit(Sender: TObject);
var
  str, Token: string;
  Separator, KeyWords: string;
  DfmGlyph: TDfmGlyph;
begin
{
  if TreeView.Selected = nil then Exit;

  DfmGlyph := TDfmGlyph(TreeView.Selected.Data);

  FGlobalKeyWords.DeleteGUID(DfmGlyph.FGUID);

  Str := MemoKeyWord.Text;
  Separator := '';
  KeyWords := '';

  while str <> '' do
  begin
    Token := ReadTokenSt(Str);
    Token := Trim(Token);

    KeyWords := KeyWords + Separator + Token;
    Separator := ';';

    FGlobalKeyWords.AddKeyWord(Token, DfmGlyph.FGUID, DfmGlyph.FGlobalIndex);
  end;

  UpDateKeyWords(KeyWords, True);
  DFmGlyph.FKeyWords := KeyWords;
}
end;

procedure TFAdminGlobalIndex.clbKeyWordsClickCheck(Sender: TObject);
var
  i: Integer;
  Separator, KeyWords: string;
  DfmGlyph: TDfmGlyph;
begin
{
  if TreeView.Selected = nil then Exit;

  DfmGlyph := TDfmGlyph(TreeView.Selected.Data);

  Separator := '';
  KeyWords := '';

  for i := 0  to clbKeyWords.Count - 1 do
  begin
    if clbKeyWords.Checked[i] then
    begin
      KeyWords := KeyWords + Separator + clbKeyWords.Items[i];
      Separator := ';';
    end;
  end;

  UpDateKeyWords(KeyWords);
  DFmGlyph.FKeyWords := KeyWords;
}
end;

procedure TFAdminGlobalIndex.btnUpdateKeyWordsClick(Sender: TObject);
var
  i:  Integer;
  KeyWords, Str, Token: string;

  function ReadKeyToken(var Buffer: string): string;
  var
    State: Integer;
    c: Char;
    Pos: Integer;
  begin
    State := 0;
    Pos := 1;
    Result := '';

    while (State <> 10) and (Pos <= Length(Buffer)) do
    begin
      c := Buffer[Pos];

      if c in ['A'..'Z'] then
        Result := Result + c
      else
      begin
        if Length(Result) <= 3 then
          Result := ''
        else
          Break;
      end;

      Pos := Pos + 1;
    end;

    Buffer := Copy(Buffer, Pos + 1, Length(Buffer) - Pos);
  end;

begin
{
  Screen.Cursor := crHourGlass;

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if FDfmGlyphs[i].FObjectHint <> '' then
    begin
      if FDfmGlyphs[i].FRefGUID = '' then
        KeyWords := FDfmGlyphs[i].FKeyWords
      else
        KeyWords := FDfmGlyphs[FDfmGlyphs.IndexOfGUID(FDfmGlyphs[i].FRefGUID)].FKeyWords;

      Str := FDfmGlyphs[i].FObjectHint;

      while Str <> '' do
      begin
        Token := ReadKeyToken(Str);

        if (ToKen <> '') and (Pos(Token, KeyWords) = 0) then
        begin
          if KeyWords = '' then
            KeyWords := Token
          else
            KeyWords := KeyWords + ';' + Token;
        end;
      end;

      if FDfmGlyphs[i].FRefGUID = '' then
        FDfmGlyphs[i].FKeyWords := KeyWords
      else
        FDfmGlyphs[FDfmGlyphs.IndexOfGUID(FDfmGlyphs[i].FRefGUID)].FKeyWords := KeyWords;
    end;

      if (i mod 10000) = 9999 then
        OutText('Calcul des mots Clefs à ' + IntToStr(i * 100 div FDfmGlyphs.Count) + '%');
  end;

  BuildKeyWords();

  Screen.Cursor := crDefault;
}
end;

procedure TFAdminGlobalIndex.tbbTestClick(Sender: TObject);
var
//  bt: THBitbtn;
  bt: THSpeedButton;
  F: TFGlobalIndexPropertyEditor;
  FileNames: TStringList;
begin

{  bt := THBitbtn.Create(Self);
  bt.Parent := Self;
  bt.Caption := 'Coucou';
  bt.GlobalIndexImage := tbbTest.GlobalIndexImage;

  bt := THSpeedButton.Create(Self);
  bt.Parent := Self;
  bt.Caption := 'Coucou';
  bt.Width := 100;

  bt.GlobalIndexImage := tbbTest.GlobalIndexImage;
//  bt.Glyph :=  Image1.Picture.Bitmap;
 }

(*
  F := TFGlobalIndexPropertyEditor.Create(Self, FGlobalKeyWords, True);
  F.GlobalIndexImage := tbbTest.GlobalIndexImage;

  if F.ShowModal = mrOk then
  begin
//    tbbTest.GlobalIndexImage := F.GlobalIndexImage;
  end;

  F.Free;
*)

  FileNames := TStringList.Create;
  FileNames.Add('C:\AglEncours5\TRAVAIL\PGR\AGL7\IM\about.dfm');
  OpenTeamParseDfm(Self, FileNames);
  FileNames.Free;

end;

procedure TFAdminGlobalIndex.HSpinEditChange(Sender: TObject);
var
  DfmGlyph: TDfmGlyph;
begin
{
  if TreeView.Selected = nil then Exit;

  DfmGlyph := TDfmGlyph(TreeView.Selected.Data);

  DfmGlyph.FLevel := HSpinEdit.Value;
}
end;

procedure TFAdminGlobalIndex.EditLotChange(Sender: TObject);
var
  DfmGlyph: TDfmGlyph;
begin
{
  if TreeView.Selected = nil then Exit;

  DfmGlyph := TDfmGlyph(TreeView.Selected.Data);

  DfmGlyph.FLot := EditLot.Text;
}
end;

procedure TFAdminGlobalIndex.HBitBtn1Click(Sender: TObject);
//var
//  AglGlobalKeyWords: TGlobalKeyWords;
//  ms: TMemoryStream;
begin
{
  AglGlobalKeyWords := TGlobalKeyWords.Create;

  try
    ms := TMemoryStream.Create;

    try
      LoadDataFromAglRes('GLOBALKEYWORDS', ms);

      if ms.Size > 0 then
      begin
        ms.Position := 0;
        AglGlobalKeyWords.LoadFromStream(ms);
      end
      else
        Application.MessageBox('Pas de mots clefs', 'GlobalIndexProperty');
    finally
      ms.Free;
    end;
  except
    on E: Exception do
      Application.MessageBox(PChar(E.Message), 'GlobalIndexProperty');
  end;

  AglGlobalKeyWords.Free;
}
end;

procedure TFAdminGlobalIndex.BtnSaveFileClick(Sender: TObject);
begin
  SaveDialog.DefaultExt := '.Dat';
  SaveDialog.Filter := 'Fichiers Data (*.Dat)|*.DAT';
  SaveDialog.Title := 'Sauver un fichier de données';

  if SaveDialog.Execute then
    FDfmGlyphs.SaveToFile(SaveDialog.FileName);
end;

procedure TFAdminGlobalIndex.BtnLoadFileClick(Sender: TObject);
var
  i: Integer;
begin
  OpenDialog.DefaultExt := '.Dat';
  OpenDialog.Filter := 'Fichiers Data (*.Dat)|*.DAT';
  OpenDialog.Title := 'Ouvir un fichier de données';

  if OpenDialog.Execute then
  begin
    Screen.Cursor := crHourGlass;
    FDfmGlyphs.LoadFromFile(OpenDialog.FileName);
    BuildKeyWords;
    RefreshTreeView();
    Screen.Cursor := crDefault;
  end;
end;

procedure TFAdminGlobalIndex.HSpeedButton1Click(Sender: TObject);
var
  Ret: string;
begin
  Ret := TCegidPath.GetAppData();
//  HImageList1.GlobalIndexImages.Add('Z0120_S16G1');
end;

procedure TFAdminGlobalIndex.btn256Click(Sender: TObject);
var
  iDfmGlyph: Integer;
begin
{
  for iDfmGlyph := 0 to FDfmGlyphs.Count - 1 do
  begin
    if FDfmGlyphs[iDfmGlyph].FGlobalIndex = '' then
    begin
      try
        if FDfmGlyphs.FindGlobalIndex(FDfmGlyphs[iDfmGlyph].ImgHex) = '' then
          OutText(FDfmGlyphs[iDfmGlyph].FObjectName + ' : ' + FDfmGlyphs[iDfmGlyph].FFileName);
      except
        OutText('Erreur Image : ' + FDfmGlyphs[iDfmGlyph].FObjectName + ' : ' + FDfmGlyphs[iDfmGlyph].FFileName);
      end;
    end;
  end;
}
end;

procedure TFAdminGlobalIndex.BtnLoadImageClick(Sender: TObject);
var
  i: Integer;
  DfmGlyph: TDfmGlyph;
begin
  Screen.Cursor := crHourGlass;

  FDfmGlyphs.LoadImages(EditSavePath.Text);
  RefreshTreeView();


  Screen.Cursor := crDefault;
end;

procedure TFAdminGlobalIndex.BtnSaveDllVistaClick(Sender: TObject);
var
  Path: String;
  i: Integer;
  FrMgr: TFileResourceMgr;
  ms: TMemoryStream;
begin
  Screen.Cursor := crHourGlass;

  TGlobalIndex.FreeFileGlobalImageRes;

  Path := ExtractFilePath(Application.ExeName);

  if FileExists(Path + 'AglResVista.dll') then
    DeleteFile(Path + 'AglResVista.dll');

  CopyFile(Path + 'AglResTemplate.dll', Path + 'AglResVista.dll');

  FrMgr := TFileResourceMgr.Create(Path + 'AglResVista.dll');

  for i := 0 to FDfmGlyphs.Count - 1 do
  begin
    if (FDfmGlyphs.Items[i].RefGUID = '') and (FDfmGlyphs.Items[i].ImageType = 'PNG') then
    begin

      ms := TMemoryStream.Create;

      try
        FDfmGlyphs.Items[i].Bitmap.SaveToStream(ms);
        ms.Position := 0;

        if FrMgr.WriteResourceTryed(10, rtRCData, FDfmGlyphs.Items[i].GlobalIndex, ms) <> 0 then
        begin
          PGIInfo('Echec de la génération !');
          Exit;
        end;
      finally
        ms.Free;
      end;
    end;
  end;

  //Sauvegarde des mots clefs

{  ms := TMemoryStream.Create;
  FGlobalKeyWords.SaveToStream(ms);
  ms.Position := 0;
  FrMgr.WriteResource(rtRCData, 'GLOBALKEYWORDS', ms);
  ms.Free; }

  FrMgr.Free;

  Screen.Cursor := crDefault;
end;


procedure TFAdminGlobalIndex.ToolbarButton971Click(Sender: TObject);
var
  pic: TPicture;
begin
{  Pic := TPicture.Create;
  Pic.LoadFromFile('C:\Users\gronchi\Documents\PGI Aero\Icones modifiées priorité1 PNG\png\Z0027_S32G1.png');
  HSpeedButton1.Glyph.Assign(Pic);
  Pic.Free;
}
  Pic := TPicture.Create;
  Pic.LoadFromFile('C:\Users\gronchi\Documents\PGI Aero\Icones modifiées priorité1 PNG\png\Z0027_S32G1.png');
  imageList1.Add(Pic.Bitmap, nil);
  Pic.Free;


end;

procedure TFAdminGlobalIndex.Button2Click(Sender: TObject);
var
  F: TSearchRec;
  Ret: Integer;

  FileName: string;
  Index: Integer;
  GlobalIndex: string;
begin
  Ret := FindFirst(EditFileName.Text, faDirectory, F);

  while Ret = 0 do
  begin
    if (F.Name <> '.') and (F.Name <> '..') and ((F.Attr and faDirectory) = 0) then
    begin
      FileName := ExtractFileName(F.Name);
      FileName := LeftStr(FileName, Length(FileName) - 4);

        Index := StrToInt(FileName);

      GlobalIndex := TGlobalIndex.FormatGlobalIndex('O', Index, 16);

      RenameFile(ExtractFilePath(EditFileName.Text) + '\' + F.Name, ExtractFilePath(EditFileName.Text) + '\' + GlobalIndex + '.png');
    end;

    Ret := FindNext(F);
  end;
end;

end.
