unit USaveRestore;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, ComCtrls, Db,
  UTOB,
  HPanel, HEnt1, ImgList, TeEngine, Series, TeeProcs, Chart, HTB97,HmsgBox,
  Hctrls,
{$IFNDEF EAGLCLIENT}
  UPDomaine,
{$ENDIF}
{$IFDEF CISXPGI}
  uYFILESTD,
  ULibCpContexte,
  {$IFDEF EAGLCLIENT}
    UScriptTob,
  {$ENDIF}
{$ENDIF}
 uDbxDataSet, Variants, ADODB, 
 Uscript, HImgList;

type
  TFSaveRestore = class(TForm)
    Splitter1: TSplitter;
    Panel3: TPanel;
    Panel4: TPanel;
    EdFicName: TEdit;
    BitBtn1: THBitBtn;
    LSave: TLabel;
    LRest: TLabel;
    Panel5: TPanel;
    OpenDialog: TOpenDialog;
    ImageList1: THImageList;
    Panel6: TPanel;
    bOuvrir: TToolbarButton97;
    BAide: TToolbarButton97;
    Chart: TChart;
    Series1: TPieSeries;
    CBTables: TComboBox;
    Label1: TLabel;
    RGSR: TRadioGroup;
    Panel1: TPanel;
    TV: TTreeView;
    Panel2: TPanel;
    Label2: TLabel;
    lblTraitement: TLabel;
    Ldomaine: TLabel;
    Domaine: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RGSRClick(Sender: TObject);
    procedure CBTablesChange(Sender: TObject);
    procedure TVDblClick(Sender: TObject);
    procedure TVChange(Sender: TObject; Node: TTreeNode);
    procedure BitBtn1Click(Sender: TObject);
    procedure bOuvrirClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    TobGen      : TOB;
    TobTV       : TOB;
    TobCourante : TOB;
    TobDomaine  : TOB;
{$IFNDEF CISXPGI}
    TBase : TADOTable;
{$ELSE}
    TBase : TQuery;
{$ENDIF}
    procedure InitTobDomaine;
{$IFNDEF EAGLCLIENT}
    procedure InsereDansTobTable;
{$ENDIF}
    function  EtudieTobTable(TobLig : TOB) : integer;
    procedure TraiteTobTable(TobLig, TobSave : TOB; NbMax : integer; var NbTraite : integer);
    procedure EnregTobSave(TobLig, TobSave : TOB; NbMax : integer; var NbTraite : integer);
    procedure EnregTobTable(TobLig : TOB; NbMax : integer; var NbTraite : integer);
    procedure EnregTobRest(TobLig : TOB; NbMax : integer; var NbTraite : integer);
    function  SearchNodeFromTob(TobL : TOB) : TTreeNode;
//    function  SearchTobFromNode(TN : TTreeNode) : TOB;
    procedure ChargeFichierSave(TobRest : TOB);
    procedure EnregistreTob(TobF : TOB);
    procedure AnalyseTV;
  public
    { Déclarations publiques }
  end;

type TOB_2 = class(TOB)
     public
       procedure LoadFromXMLFile(FName : String);
       function  YouzTob(T : TOB) : integer;
     end;

var
  FSaveRestore: TFSaveRestore;

procedure SaveRestore(Save : boolean; PP : THPanel);
{$IFNDEF EAGLCLIENT}
function  BlobToString(Blob: TBlobField): string;
procedure StringToBlob(Value: string; var Field : TBlobField);
{$ENDIF}

{$IFDEF CISXPGI}
function LanceRestoreCisx (Cmd : string) : Boolean;
{$ENDIF}

implementation

uses UDMIMP, UIUtil; // , voirtob;

{$R *.DFM}

{$IFDEF CISXPGI}
function LanceRestoreCisx (Cmd : string) : Boolean;
begin
   TCPContexte.GetCurrent.VarCisx.CHARGECIX(Cmd);
   SaveRestore(FALSE, nil);
   TCPContexte.Release;
   Result := TRUE;
end;
{$ENDIF}


procedure SaveRestore(Save : boolean; PP : THPanel);
var XX  : TFSaveRestore ;
TobTemp      : TOB_2;
ind1         : integer;
TOBS         : TOB;
BEGIN
XX:=TFSaveRestore.Create(Application);
With XX do
begin
      if Save then
      begin
          Caption := 'Sauvegarde des paramètrages';
          RGSR.ItemIndex := 0;
          Domaine.Enabled := TRUE; Ldomaine.Enabled := TRUE;
{$IFNDEF CISXPGI}
         ChargementComboDomaine (Domaine);
{$ENDIF}
      end
      else
      begin
          RGSR.ItemIndex := 1;
          Caption := 'Restauration des paramètrages';
          Domaine.visible := FALSE; Ldomaine.Enabled := FALSE;
      end;
      if  (UpperCase(GetInfoVHCX.Mode) = 'RESTOR') then
      begin
        EdFicName.text := GetInfoVHCX.ListeFichier;
        if (EdFicName.text = '') or (not FileExists(EdFicName.text)) then
        begin
                    PGIInfo ('Le fichier '+ EdFicName.text+ ' n''existe pas','');
                    XX.free;
                    exit;
        end;
        TobTemp := TOB_2.Create('', nil, -1);
        TobTemp.LoadFromXMLFile(EdFicName.Text);
        TOBS := TobTemp.Detail[0];
        ChargeFichierSave(TOBS);
        for ind1 := 0 to TOBS.Detail.Count - 1 do
        begin
             EnregistreTob(TOBS.Detail[ind1]);
        end;
        if TOBS.Detail.Count <> 0 then
           TOBS.SaveToFile(GetInfoVHCX.Directory+'\'+GetInfoVHCX.NomFichier,True,True,True);
        TobTemp.free;
        close;
      end
      else
      begin
{$IFDEF CISXPGI}
                  Try
                   ShowModal ;
                  Finally
                   Free ;
                  End ;
{$ELSE}
              if PP=Nil then
              BEGIN
                  Try
                   ShowModal ;
                  Finally
                   Free ;
                  End ;
              END else
              BEGIN
                      InitInside(XX,PP) ;
                      Show ;
              END ;
{$ENDIF}
      end;
end;
end;

//==============================================================================
{$IFNDEF EAGLCLIENT}
function BlobToString(Blob: TBlobField): string;
var BinStream : TBlobStream;
    StrStream : TStringStream;
    s : string;

begin
BinStream := TBlobStream.Create(Blob, bmRead);
try
    StrStream := TStringStream.Create(s);
    try
        BinStream.Seek(0, soFromBeginning);
        StrStream.CopyFrom(BinStream, BinStream.Size);
        StrStream.Seek(0, soFromBeginning);
        Result := StrStream.DataString;
        StrStream.Free;
    except
        Result := '';
    end;
finally
    BinStream.Free
end;
end;

procedure StringToBlob(Value: string; var Field : TBlobField);
var StrStream : TStringStream;
    BinStream : TBlobStream;
begin
StrStream := TStringStream.Create(Value);
try
    BinStream := TBlobStream.Create(Field, bmWrite);
    try
        StrStream.Seek(0, soFromBeginning);
        BinStream.CopyFrom(StrStream, StrStream.Size);
    finally
        BinStream.Free;
    end;
finally
    StrStream.Free;
end;
end;
{$ENDIF}

//==============================================================================
procedure TFSaveRestore.FormShow(Sender: TObject);
begin
if RGSR.ItemIndex = 0 then
    begin
    LSave.Visible := True;
    LRest.Visible := False;
    HelpContext   := 1500;
    end
    else
    begin
    LSave.Visible := False;
    LRest.Visible := True;
    Panel2.Visible := False;
    Label1.Visible := False;
    CBTables.Visible := False;
    HelpContext   := 1600;
    end;
if TobGen <> nil then
    TobGen.Free;
TobGen := TOB.Create('Paramètres', nil, -1);
//TobGen.AddChampSupValeur('NomTable', 'Paramètres', False);
if TobTV <> nil then
    TobTV.Free;
TobTV := TOB.Create('Paramètres', nil, -1);
//TobTV.AddChampSupValeur('NomTable', 'Paramètres', False);
InitTobDomaine;
lblTraitement.Caption := '';
end;

procedure TFSaveRestore.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if TBase <> nil then
begin
    if TBase.active then TBase.Close;
    TBase.free; 
end;
if TobGen <> nil then
begin
    TobGen.Free; TobGen := nil;
end;
if TobTV <> nil then
begin
    TobTV.Free; TobTV := nil;
end;
if TobDomaine <> nil then
begin
    TobDomaine.Free;  TobDomaine := nil;
end;
end;

procedure TFSaveRestore.RGSRClick(Sender: TObject);
begin
LSave.Visible := (RGSR.ItemIndex = 0);
LRest.Visible := not LSave.Visible;
EdFicName.Text := '';
end;

procedure TFSaveRestore.CBTablesChange(Sender: TObject);
var
    Count         : integer;
    Indice        : integer;
    Dom           : string;
begin
{$IFNDEF EAGLCLIENT}
Screen.Cursor := crHourGlass;
case RGSR.ItemIndex of
0 : begin
{$IFNDEF CISXPGI}
  if not assigned(TBase) then TBase := TADOTable.Create(Application);
  TBase.Connection := DMImport.Db as TADOConnection;
    if CBTables.Items.Strings[CBTables.ItemIndex] = 'Domaine' then
        TBase.TableName := DMImport.Gzimpdomaine.TableName
    else if CBTables.Items.Strings[CBTables.ItemIndex] = 'Dictionnaire' then
        TBase.TableName := DMImport.GzImpPar.TableName
    else if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Scripts') or
    (CBTables.Items.Strings[CBTables.ItemIndex] = 'Conception') then
        TBase.TableName := DMImport.GzImpReq.TableName
    else if CBTables.Items.Strings[CBTables.ItemIndex] = 'Correspondance' then
        TBase.TableName := DMImport.GzImpCorresp.TableName;
{$ELSE}
        TBase := OpenSQl ('SELECT * FROM CPGZIMPREQ', TRUE);
{$ENDIF}
    if TobGen <> nil then
        TobGen.Free;
    TobGen := TOB.Create('Paramètres', nil, -1);
    TobGen.AddChampSupValeur('NomTable', 'Paramètres', False);

    if TobTV <> nil then
        TobTV.Free;

    Indice := 0; Count := 1000;
    Dom := RendDomaine(Domaine.Text);
    TobTV := TOB.Create('Paramètres', nil, -1);
    TobTV.AddChampSupValeur('NomTable', CBTables.Items.Strings[CBTables.ItemIndex], False);
{$IFNDEF CISXPGI}
    if not TBase.active then TBase.Open;  // ajout me
    with TBase do
{$ELSE}
    if not TBase.Eof then
    with TBase do
{$ENDIF}
        begin
        First;
        while not Eof do
            begin
            // ajout me pour interface
            if  (ParamCount > 0) and (GetInfoVHCX.Mode = 'I') and (GetInfoVHCX.Domaine <> '') then
            begin
                 if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Domaine') and (FindField('Domaine').asstring <> GetInfoVHCX.Domaine) then
                 begin Next; Continue; end
                 else if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Dictionnaire') and (FindField('Domaine').asstring <> GetInfoVHCX.Domaine) then
                 begin Next; Continue; end
                 else if ((CBTables.Items.Strings[CBTables.ItemIndex] = 'Scripts')
                 or (CBTables.Items.Strings[CBTables.ItemIndex] = 'Conception'))
                 and (FindField('Domaine').asstring <> GetInfoVHCX.Domaine) then
                 begin Next; Continue; end
                 else if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Correspondance') and (FindField('Domaine').asstring <> GetInfoVHCX.Domaine)
                 and (FindField('Domaine').asstring <> '') then
                 begin
                      Next;
                      Continue;
                 end;
            end
            else
            begin
                   if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Scripts') and (Dom <> '') and (Dom <> FindField('Domaine').asstring)  then
                   begin
                            Next;
                            Continue;
                   end;
            end;
            InsereDansTobTable;
            Next;
            // pour éviter en cas de désindexation de boucler
            if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Scripts') or (CBTables.Items.Strings[CBTables.ItemIndex] = 'Conception') then
            begin
                 inc(Indice);
                 if Indice > Count then break;
            end;
            end;
        end;
    TobTV.PutTreeView(TV, nil, 'NomTable');
    AnalyseTV;
    end;
1 : begin
    end;
end;
Screen.Cursor := crDefault;
{$ENDIF}
end;

procedure TFSaveRestore.TVDblClick(Sender: TObject);
    procedure TraiteFlipFlop(TobLigne : TOB; FieldOn : boolean);
    var ind1 : integer;
        TobS : TOB;
    begin
    if FieldOn then
        TobLigne.AddChampSup('ATraiter', False)
        else
        TobLigne.DelChampSup('ATraiter', False);
    for ind1 := 0 to TobLigne.Detail.Count - 1 do
        begin
        TobS := TobLigne.Detail[ind1];
        TraiteFlipFlop(TobS, FieldOn);
        end;
    end;

begin
if TV.Selected.ImageIndex < 10 then
    begin
    TV.Selected.ImageIndex := 10 + RGSR.ItemIndex;
    TV.Selected.SelectedIndex := 10 + RGSR.ItemIndex;
    TV.Selected.Expand(False);
    TraiteFlipFlop(TOB(TV.Selected.Data), True);
    end
    else
    begin
    TV.Selected.ImageIndex := TOB(TV.Selected.Data).GetValue('ImageIndex');
    TV.Selected.SelectedIndex := TOB(TV.Selected.Data).GetValue('ImageIndex');
    TraiteFlipFlop(TOB(TV.Selected.Data), False);
    end;
end;

procedure TFSaveRestore.TVChange(Sender: TObject; Node: TTreeNode);
begin
if Node.Data = nil then Exit;
TobCourante := TOB(Node.Data);
end;

procedure TFSaveRestore.BitBtn1Click(Sender: TObject);
var TobTemp : TOB_2;
begin
if OpenDialog.Execute then
    EdFicName.Text := OpenDialog.FileName;
if Pos('.CIX', UpperCase(EdFicName.Text)) = 0 then
    EdFicName.Text := EdFicName.Text + '.cix';
if (RGSR.ItemIndex = 1) and (EdFicName.Text <> '') then
    begin
    TobTemp := TOB_2.Create('', nil, -1);
    TobTemp.LoadFromXMLFile(EdFicName.Text);
//    TobDebug(TobTemp);
    ChargeFichierSave(TobTemp.Detail[0]);
    TobTemp.free;
    end;
end;

procedure TFSaveRestore.bOuvrirClick(Sender: TObject);
var ind1, NbTraite : integer;
    TobSave : TOB;
begin
  if EdFicName.Text <> '' then
  begin
    Chart.Visible := True;
    case RGSR.ItemIndex of
      0 : begin
        lblTraitement.Caption := 'Traitement en cours ...';
        ind1 := EtudieTobTable(TobTV);
        NbTraite := 0;
        Chart.Series[0].Clear;
        Chart.Series[0].Add(0);
        Chart.Series[0].Add(ind1);
        TobSave := TOB.Create(CBTables.Items.Strings[CBTables.ItemIndex], nil, -1);
        TraiteTobTable(TobTV, TobSave, ind1, NbTraite);
        TobSave.SaveToXMLFile(EdFicName.Text, True, True);
        //TobSave.SaveToFile(EdFicName.Text+'.TOB',True,True,True);
        TobSave.Free;
        lblTraitement.Caption := 'Traitement terminé.';
        end;
      1 : begin
        lblTraitement.Caption := 'Traitement en cours ...';
        ind1 := EtudieTobTable(TobTV);
         NbTraite := 0;
        Chart.Series[0].Clear;
        Chart.Series[0].Add(0);
        Chart.Series[0].Add(ind1);
        EnregTobTable(TobTV, ind1, NbTraite);
        lblTraitement.Caption := 'Traitement terminé.';
        end;
    end;
  end
  else // Champ nom de fichier est vide
  begin PGIInfo('Veuillez renseigner le nom de fichier.','Traitement'); end;
end;

//==============================================================================
procedure TFSaveRestore.InitTobDomaine;
{$IFNDEF CISXPGI}
var TobL : TOB;
TDomaine : TADOTable;
{$ENDIF}
begin
{$IFNDEF CISXPGI}
TDomaine := TADOTable.Create(Application);
TDomaine.Connection := DMImport.Db as TADOConnection;
TDomaine.TableName := DMImport.GzImpDomaine.TableName;
TobDomaine := TOB.Create('', nil, -1);
with TDomaine do
    begin
    Open;
    First;
    while not Eof do
        begin
        TobL := TOB.Create('', TobDomaine, -1);
        TobL.AddChampSupValeur('Domaine', TDomaine.FindField('Domaine').AsString, False);
        TobL.AddChampSupValeur('Libelle', TDomaine.FindField('Libelle').AsString, False);
        Next;
        end;
    Close;
    end;
{$ENDIF}
end;
{$IFNDEF EAGLCLIENT}
{$IFNDEF CISXPGI}
procedure TFSaveRestore.InsereDansTobTable;
var ind1                                   : integer;
    TobL, TobChamp                         : TOB;
    TobTVDomaine, TobTVTable               : TOB;
    FNumero, FLocDomaine, FLocTable, FNom, FNomChamp : TStringField;
    stBlob                                 : string;
    TDelim                                 : TADOTable;
begin
TobL := nil; TobTVDomaine:=nil;
if CBTables.Items.Strings[CBTables.ItemIndex] = 'Domaine' then
    begin
    for ind1 := 0 to TBase.FieldCount - 1 do
        begin
        if ind1 = 0 then
            begin
            TobTVDomaine := TOB.Create(TBase.Fields[ind1].AsString, TobTV, -1);
            TobTVDomaine.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            TobL := TOB.Create(TBase.Fields[ind1].AsString, TobGen, -1);
            TobL.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'Domaine' then
            begin
            TobTVDomaine.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            TobL.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'Libelle' then
            begin
            TobTVDomaine.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
            TobL.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            end;
        TobL.AddChampSupValeur(TBase.Fields[ind1].FieldName, TBase.Fields[ind1].AsString, False);
        end;
    end
else if CBTables.Items.Strings[CBTables.ItemIndex] = 'Correspondance' then
    begin
//    FNumero     := TStringField(TBase.FindField('Ident'));
    FLocDomaine := TStringField(TBase.FindField('Domaine'));
//    FLocTable   := TStringField(TBase.FindField('TableName'));
    FNom        := TStringField(TBase.FindField('Profile'));
    FNomChamp   := TStringField(TBase.FindField('ChampCode'));
    TobTVDomaine := TobTV.FindFirst(['NomTable'], [FNom.AsString + '-' +FNomChamp.AsString], False);
    if TobTVDomaine = nil then
        begin
        TobTVDomaine := TOB.Create(FNom.AsString + '-' +FNomChamp.AsString, TobTV, -1);
        TobTVDomaine.AddChampSupValeur('NomTable', FNom.AsString + '-' +FNomChamp.AsString, False);
        TobTVDomaine.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
        TobTVDomaine.AddChampSupValeur('LaTable'   , FNomChamp.AsString, False);
        TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
        end;
    TobChamp := TOB.Create('', TobGen, -1);
    TobChamp.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
    TobChamp.AddChampSupValeur('LaTable'   , FNomChamp.AsString, False);

    for ind1 := 0 to TBase.FieldCount - 1 do
        TobChamp.AddChampSupValeur(TBase.Fields[ind1].FieldName, TBase.Fields[ind1].AsString, False);
    end
else if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Scripts') or (CBTables.Items.Strings[CBTables.ItemIndex] = 'Conception') then
begin
    TobChamp := nil;
    // ajout me pour sauvegarde delimitée
    TDelim := TADOTable.Create(Application);
    TDelim.Connection := DMImport.Db as TADOConnection;
    TDelim.Tablename := DMImport.GzImpDelim.TableName;
    with TDelim do
    begin
         if not active then Open;
         if TDelim.Locate('CLE', TBase.FieldByName('CLE').AsString , [loCaseInsensitive]) then
         begin
            TobChamp := TOB.Create(TDelim.FieldByName('CLE').AsString, nil, -1);
            for ind1 := 0 to TDelim.FieldCount - 1 do
            if (TDelim.Fields[ind1].DataType = ftBlob) or (TDelim.Fields[ind1].DataType =  ftMemo) then
            begin
                    stBlob := BlobToString(TBlobField(TDelim.Fields[ind1]));
                    TobChamp.AddChampSupValeur(TDelim.Fields[ind1].FieldName, stBlob, False);
            end
            else
                    TobChamp.AddChampSupValeur(TDelim.Fields[ind1].FieldName, TDelim.Fields[ind1].AsString, False);

         end;
    end;
    for ind1 := 0 to TBase.FieldCount - 1 do
        begin
        if ind1 = 0 then
            begin
            TobTVDomaine := TOB.Create(TBase.Fields[ind1].AsString, TobTV, -1);
            TobTVDomaine.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            TobL := TOB.Create(TBase.Fields[ind1].AsString, TobGen, -1);
            TobL.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'DOMAINE' then
            begin
            TobTVDomaine.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            TobL.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'CLE' then
            begin
            TobTVDomaine.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
            TobL.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            end;
        if (TBase.Fields[ind1].DataType = ftBlob) or (TBase.Fields[ind1].DataType =  ftMemo) then
            begin
            stBlob := BlobToString(TBlobField(TBase.Fields[ind1]));
            TobL.AddChampSupValeur(TBase.Fields[ind1].FieldName, stBlob, False);
            end
            else
            TobL.AddChampSupValeur(TBase.Fields[ind1].FieldName, TBase.Fields[ind1].AsString, False);
        end;
        if TobChamp <> nil then
        begin
            for ind1 := 0 to TDelim.FieldCount - 1 do
                TobL.AddChampSupValeur('Delim-'+TDelim.Fields[ind1].FieldName,
                                       TobChamp.GetValue(TDelim.Fields[ind1].FieldName), False);
            TobChamp.free;
        end;
        TDelim.close; TDelim.free;
end
else if CBTables.Items.Strings[CBTables.ItemIndex] = 'Dictionnaire' then
    begin
    FNumero  := TStringField(TBase.FindField('Numero'));
    FLocDomaine := TStringField(TBase.FindField('Domaine'));
    FLocTable   := TStringField(TBase.FindField('TableName'));
    FNom     := TStringField(TBase.FindField('Nom'));
    TobTVDomaine := TobTV.FindFirst(['NomTable'], [FLocDomaine.AsString], False);
    if TobTVDomaine = nil then
        begin
        TobTVDomaine := TOB.Create(FLocDomaine.AsString, TobTV, -1);
        TobTVDomaine.AddChampSupValeur('NomTable', FLocDomaine.AsString, False);
        TobTVDomaine.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
        TobTVDomaine.AddChampSupValeur('LaTable'   , '', False);
        TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
        end;
    TobTVTable := TobTVDomaine.FindFirst(['NomTable'], [FLocTable.AsString], False);
    if TobTVTable = nil then
        begin
        TobTVTable := TOB.Create(FLocTable.AsString, TobTVDomaine, -1);
        TobTVTable.AddChampSupValeur('NomTable', FLocTable.AsString, False);
        TobTVTable.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
        TobTVTable.AddChampSupValeur('LaTable'   , FLocTable.AsString, False);
        TobTVTable.AddChampSupValeur('ImageIndex', 1, False);
        TobTVTable.AddChampSupValeur('NbFilles', 0, False);
        end;
    TobTVTable.PutValue('NbFilles', TobTVTable.GetValue('NbFilles') + 1);
    TobChamp := TOB.Create('', TobGen, -1);
    TobChamp.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
    TobChamp.AddChampSupValeur('LaTable'   , FLocTable.AsString, False);
    TobChamp.AddChampSupValeur(FNumero.FieldName  , FNumero.AsString, False);
    TobChamp.AddChampSupValeur(FLocDomaine.FieldName , FLocDomaine.AsString, False);
    TobChamp.AddChampSupValeur(FLocTable.FieldName   , FLocTable.AsString, False);
    TobChamp.AddChampSupValeur(FNom.FieldName     , FNom.AsString, False);
    for ind1 := 4 to TBase.FieldCount - 1 do
        TobChamp.AddChampSupValeur(TBase.Fields[ind1].FieldName, TBase.Fields[ind1].AsString, False);
    end;
end;
{$ELSE}
procedure TFSaveRestore.InsereDansTobTable;
var ind1                                   : integer;
    TobL, TobChamp                         : TOB;
    TobTVDomaine, TobTVTable               : TOB;
    FNumero, FLocDomaine, FLocTable, FNom, FNomChamp : TStringField;
    stBlob                                 : string;
    TDelim                                 : TTable;
    NomField,Tmp                           : string;
begin
TobL := nil; TobTVDomaine:=nil;
if CBTables.Items.Strings[CBTables.ItemIndex] = 'Domaine' then
    begin
    for ind1 := 0 to TBase.FieldCount - 1 do
        begin
        if ind1 = 0 then
            begin
            TobTVDomaine := TOB.Create(TBase.Fields[ind1].AsString, TobTV, -1);
            TobTVDomaine.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            TobL := TOB.Create(TBase.Fields[ind1].AsString, TobGen, -1);
            TobL.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'Domaine' then
            begin
            TobTVDomaine.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            TobL.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'Libelle' then
            begin
            TobTVDomaine.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
            TobL.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            end;
        TobL.AddChampSupValeur(TBase.Fields[ind1].FieldName, TBase.Fields[ind1].AsString, False);
        end;
    end
else if CBTables.Items.Strings[CBTables.ItemIndex] = 'Correspondance' then
    begin
    FLocDomaine := TStringField(TBase.FindField('Domaine'));
    FNom        := TStringField(TBase.FindField('Profile'));
    FNomChamp   := TStringField(TBase.FindField('ChampCode'));
    TobTVDomaine := TobTV.FindFirst(['NomTable'], [FNom.AsString + '-' +FNomChamp.AsString], False);
    if TobTVDomaine = nil then
        begin
        TobTVDomaine := TOB.Create(FNom.AsString + '-' +FNomChamp.AsString, TobTV, -1);
        TobTVDomaine.AddChampSupValeur('NomTable', FNom.AsString + '-' +FNomChamp.AsString, False);
        TobTVDomaine.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
        TobTVDomaine.AddChampSupValeur('LaTable'   , FNomChamp.AsString, False);
        TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
        end;
    TobChamp := TOB.Create('', TobGen, -1);
    TobChamp.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
    TobChamp.AddChampSupValeur('LaTable'   , FNomChamp.AsString, False);

    for ind1 := 0 to TBase.FieldCount - 1 do
        TobChamp.AddChampSupValeur(TBase.Fields[ind1].FieldName, TBase.Fields[ind1].AsString, False);
    end
else if (CBTables.Items.Strings[CBTables.ItemIndex] = 'Scripts') or (CBTables.Items.Strings[CBTables.ItemIndex] = 'Conception') then
    begin
    TobChamp := nil;
    for ind1 := 0 to TBase.FieldCount - 1 do
        begin
        NomField := TBase.Fields[ind1].FieldName;
        Tmp := ReadTokenPipe(NomField,'CIS_');
        if ind1 = 0 then
            begin
            TobTVDomaine := TOB.Create(TBase.Fields[ind1].AsString, TobTV, -1);
            TobTVDomaine.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            TobL := TOB.Create(TBase.Fields[ind1].AsString, TobGen, -1);
            TobL.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'CIS_DOMAINE' then
            begin
            TobTVDomaine.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            TobL.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
            end;
        if TBase.Fields[ind1].FieldName = 'CIS_CLE' then
            begin
            TobTVDomaine.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
            TobL.AddChampSupValeur('LaTable'   , TBase.Fields[ind1].AsString, False);
            end;
        if (TBase.Fields[ind1].DataType = ftBlob) or (TBase.Fields[ind1].DataType =  ftMemo) then
            begin
            stBlob := BlobToString(TBlobField(TBase.Fields[ind1]));
            TobL.AddChampSupValeur(NomField, stBlob, False);
            end
            else
        if TBase.Fields[ind1].FieldName = 'CIS_DATEMODIF' then
            TobL.AddChampSupValeur('DATEDEMODIF', Copy(TBase.Fields[ind1].AsString,0,10), False)
            else
            TobL.AddChampSupValeur(NomField, TBase.Fields[ind1].AsString, False);
        end;
    if TobChamp <> nil then
    begin
        for ind1 := 0 to TDelim.FieldCount - 1 do
            TobL.AddChampSupValeur('Delim-'+TDelim.Fields[ind1].FieldName,
                                   TobChamp.GetValue(TDelim.Fields[ind1].FieldName), False);
    end;
    end
else if CBTables.Items.Strings[CBTables.ItemIndex] = 'Dictionnaire' then
    begin
    FNumero  := TStringField(TBase.FindField('Numero'));
    FLocDomaine := TStringField(TBase.FindField('Domaine'));
    FLocTable   := TStringField(TBase.FindField('TableName'));
    FNom     := TStringField(TBase.FindField('Nom'));
    TobTVDomaine := TobTV.FindFirst(['NomTable'], [FLocDomaine.AsString], False);
    if TobTVDomaine = nil then
        begin
        TobTVDomaine := TOB.Create(FLocDomaine.AsString, TobTV, -1);
        TobTVDomaine.AddChampSupValeur('NomTable', FLocDomaine.AsString, False);
        TobTVDomaine.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
        TobTVDomaine.AddChampSupValeur('LaTable'   , '', False);
        TobTVDomaine.AddChampSupValeur('ImageIndex', 7, False);
        end;
    TobTVTable := TobTVDomaine.FindFirst(['NomTable'], [FLocTable.AsString], False);
    if TobTVTable = nil then
        begin
        TobTVTable := TOB.Create(FLocTable.AsString, TobTVDomaine, -1);
        TobTVTable.AddChampSupValeur('NomTable', FLocTable.AsString, False);
        TobTVTable.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
        TobTVTable.AddChampSupValeur('LaTable'   , FLocTable.AsString, False);
        TobTVTable.AddChampSupValeur('ImageIndex', 1, False);
        TobTVTable.AddChampSupValeur('NbFilles', 0, False);
        end;
    TobTVTable.PutValue('NbFilles', TobTVTable.GetValue('NbFilles') + 1);
    TobChamp := TOB.Create('', TobGen, -1);
    TobChamp.AddChampSupValeur('LeDomaine' , FLocDomaine.AsString, False);
    TobChamp.AddChampSupValeur('LaTable'   , FLocTable.AsString, False);
    TobChamp.AddChampSupValeur(FNumero.FieldName  , FNumero.AsString, False);
    TobChamp.AddChampSupValeur(FLocDomaine.FieldName , FLocDomaine.AsString, False);
    TobChamp.AddChampSupValeur(FLocTable.FieldName   , FLocTable.AsString, False);
    TobChamp.AddChampSupValeur(FNom.FieldName     , FNom.AsString, False);
    for ind1 := 4 to TBase.FieldCount - 1 do
        TobChamp.AddChampSupValeur(TBase.Fields[ind1].FieldName, TBase.Fields[ind1].AsString, False);
    end;
end;
{$ENDIF}
{$ENDIF}

function TFSaveRestore.EtudieTobTable(TobLig : TOB) : integer;
var ind1, ind2 : integer;
    TobL : TOB;
begin
Result := 0;
for ind1 := 0 to TobLig.Detail.Count - 1 do
    begin
    TobL := TobLig.Detail[ind1];
    if TobL.FieldExists('ATraiter') then
        begin
        if TobL.FieldExists('NbFilles') then
            Result := Result + TobL.GetValue('NbFilles')
            else
            begin
            ind2 := EtudieTobTable(TobL);
            if ind2 > 0 then
                Result := Result + ind2
                else
                Result := Result + 1;
            end;
        end
        else
        begin
        Result := Result + EtudieTobTable(TobL);
        end;
    end;
end;

procedure TFSaveRestore.TraiteTobTable(TobLig, TobSave : TOB; NbMax : integer; var NbTraite : integer);
var ind1 : integer;
    TobL : TOB;
begin
  if TobLig.FieldExists('ATraiter') then
    EnregTobSave(TobLig, TobSave, NbMax, NbTraite)
  else
  begin
    for ind1 := 0 to TobLig.Detail.Count - 1 do
        begin
        TobL := TobLig.Detail[ind1];
        if TobL.FieldExists('ATraiter') then
            EnregTobSave(TobL, TobSave, NbMax, NbTraite)
            else
            TraiteTobTable(TobL, TobSave, NbMax, NbTraite);
        end;
    end;

end;

procedure TFSaveRestore.EnregTobSave(TobLig, TobSave : TOB; NbMax : integer; var NbTraite : integer);
var TobL, TobS : TOB;
    TN1 : TTreeNode;
    ind1 : integer;
    LeDomaine, LaTable : string;
begin
LeDomaine := TobLig.GetValue('LeDomaine');
LaTable := TobLig.GetValue('LaTable');
LeDomaine := Trim(LeDomaine);
LaTable := Trim(LaTable);  ind1 := 0;
if LeDomaine = '' then
    begin
    ind1 := 0;
    TobL := TobGen.Detail[ind1];
    end
else if LaTable = '' then
    TobL := TobGen.FindFirst(['LeDomaine'], [LeDomaine], True)
    else
    TobL := TobGen.FindFirst(['LeDomaine', 'LaTable'], [LeDomaine, LaTable], True);
while TobL <> nil do
    begin
    Application.ProcessMessages;
    TobS := TOB.Create('Ligne', TobSave, -1);
    TobS.Dupliquer(TobL, False, True, True);
    TobS.DelChampSup('LeDomaine', False);
    TobS.DelChampSup('LaTable', False);
    TN1 := SearchNodeFromTob(TobLig);
    if (TN1 <> nil) and (TN1.ImageIndex <> 12) then
        begin
        TN1.ImageIndex := 12;
        TN1.SelectedIndex := 12;
        TV.Repaint;
        end;
    Inc(NbTraite);
    Chart.Series[0].Clear;
    Chart.Series[0].Add(NbTraite);
    Chart.Series[0].Add(NbMax - NbTraite);
    Chart.Repaint;
    sleep(50);
    if LeDomaine = '' then
        begin
        Inc(ind1);
        if ind1 < TobGen.Detail.Count then
            TobL := TobGen.Detail[ind1]
            else
            TobL := nil;
        end
    else if LaTable = '' then
        TobL := TobGen.FindNext(['LeDomaine'], [LeDomaine], True)
        else
        TobL := TobGen.FindNext(['LeDomaine', 'LaTable'], [LeDomaine, LaTable], True);
    end;
end;

procedure TFSaveRestore.EnregTobTable(TobLig : TOB; NbMax : integer; var NbTraite : integer);
var ind1 : integer;
    TobL : TOB;
{$IFDEF CISXPGI}
    FFile : string;
{$ENDIF}
begin
{$IFDEF CISXPGI}
if Toblig.Getvalue ('NOMTABLE') = 'Dictionnaire' then
begin
     for ind1 := 0 to TobLig.Detail.Count - 1 do
     begin
          TobL := TobLig.Detail[ind1];
          if TobL.FieldExists('ATraiter')  then
          begin
               FFile := ExtractFileName(EdFicName.Text);
               AGL_YFILESTD_IMPORT (EdFicName.Text, 'CISX', FFile, '.CIX', TobL.GetValue('LeDomaine'), 'COMPTA', 'PARAM');
          end;
     end;
     exit;
end;
{$ENDIF}
for ind1 := 0 to TobLig.Detail.Count - 1 do
    begin
    TobL := TobLig.Detail[ind1];
    if TobL.FieldExists('ATraiter') then
        EnregTobRest(TobL, NbMax, NbTraite)
    else
        EnregTobTable(TobL, NbMax, NbTraite);
    end;
end;

procedure TFSaveRestore.EnregTobRest(TobLig : TOB; NbMax : integer; var NbTraite : integer);
var TobL       : TOB;
    TN1        : TTreeNode;
    ind1       : integer;
    LeDomaine, LaTable : string;
begin
Application.ProcessMessages;
LeDomaine := TobLig.GetValue('LeDomaine');
LaTable := TobLig.GetValue('LaTable');
LeDomaine := Trim(LeDomaine);
LaTable := Trim(LaTable);  ind1 := 0;
if LeDomaine = '' then
    begin
    ind1 := 0;
    TobL := TobGen.Detail[ind1];
    end
else if LaTable = '' then
    TobL := TobGen.FindFirst(['LeDomaine'], [LeDomaine], True)
    else
    TobL := TobGen.FindFirst(['LeDomaine', 'LaTable'], [LeDomaine, LaTable], True);
while TobL <> nil do
    begin
    EnregistreTob(TobL);
    TN1 := SearchNodeFromTob(TobL);
    if (TN1 <> nil) and (TN1.ImageIndex <> 12) then
        begin
        TN1.ImageIndex := 12;
        TN1.SelectedIndex := 12;
        TV.Repaint;
        end;
    Inc(NbTraite);
    Chart.Series[0].Clear;
    Chart.Series[0].Add(NbTraite);
    Chart.Series[0].Add(NbMax - NbTraite);
    Chart.Repaint;
    sleep(50);
    if LeDomaine = '' then
        begin
        Inc(ind1);
        if ind1 < TobGen.Detail.Count then
            TobL := TobGen.Detail[ind1]
            else
            TobL := nil;
        end
    else if LaTable = '' then
        TobL := TobGen.FindNext(['LeDomaine'], [LeDomaine], True)
        else
        TobL := TobGen.FindNext(['LeDomaine', 'LaTable'], [LeDomaine, LaTable], True);
    end;
end;
(*
function TFSaveRestore.SearchTobFromNode(TN : TTreeNode) : TOB;

    function TraiteLigne(TobLoc : TOB) : TOB;
    var ind1 : integer;
        TobL : TOB;

    begin
    for ind1 := 0 to TobLoc.Detail.Count - 1 do
        begin
        TobL := TobLoc.Detail[ind1];
        if (TobL.GetValue('Domaine') = TOB(TN.Data).GetValue('Domaine')) and
           (TobL.GetValue('TableName') = TOB(TN.Data).GetValue('TableName')) then
            begin
            Result := TobL;
            Exit;
            end;
        if TobL.Detail.Count > 0 then
            begin
            Result := TraiteLigne(TobL);
            if Result <> nil then
                Exit;
            end;
        end;
    end;

begin
Result := nil;
if TN = nil then Exit;
Result := TraiteLigne(TobGen);
end;
*)

function TFSaveRestore.SearchNodeFromTob(TobL : TOB) : TTreeNode;
begin
Result := nil;
if TV.Items.Count <= 0 then Exit;
Result := TV.Items.GetFirstNode;
while Result <> nil do
    begin
    if (TOB(Result.Data).GetValue('LeDomaine') = TobL.GetValue('Domaine')) and
       (TOB(Result.Data).GetValue('LaTable') = TobL.GetValue('TableName')) then
        Exit;
    Result := Result.GetNext;
    end;
end;


procedure TFSaveRestore.ChargeFichierSave(TobRest : TOB);
var
    ind1                            : integer;
    TobLevel1, TobLevel2, TobLevel3 : TOB;
    TobTemp                         : TOB;
begin
if Trim(EdFicName.Text) = '' then Exit;
if TobGen <> nil then
    TobGen.Free;

TobGen := TOB.Create('Paramètres', nil, -1);
TobGen.AddChampSupValeur('NomTable', 'Paramètres', False);
if TobTV <> nil then
    TobTV.Free;
TobTV := TOB.Create('Paramètres', nil, -1);
TobTV.AddChampSupValeur('NomTable', TobRest.NomTable, False);

{$IFNDEF CISXPGI}
  if not assigned(TBase) then TBase := TADOTable.Create(Application);
  TBase.Connection := DMImport.Db as TADOConnection;
if TobRest.NomTable = 'Domaine' then
    TBase.TableName := DMImport.Gzimpdomaine.TableName
else if TobRest.NomTable = 'Dictionnaire' then
    TBase.TableName := DMImport.GzImpPar.TableName
else if (TobRest.NomTable = 'Scripts') or (TobRest.NomTable = 'Conception') then
    TBase.TableName := DMImport.GzImpReq.TableName
else if TobRest.NomTable = 'Correspondance' then
    TBase.TableName := DMImport.GzImpCorresp.TableName;
if TBase.Active then TBase.Close;
TBase.Open;

{$ENDIF}

for ind1 := 0 to TobRest.Detail.Count - 1 do
    begin
    TobTemp := TobRest.Detail[ind1];
    if TobRest.NomTable = 'Domaine' then
        begin
        TobLevel1 := TOB.Create(TobTemp.GetValue('Domaine'), TobGen, -1);
        TobLevel1.Dupliquer(TobTemp, True, True, True);
        TobLevel1.AddChampSupValeur('NomTable', TobTemp.GetValue('Domaine'), False);
        TobLevel1.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('Domaine'), False);
        TobLevel1.AddChampSupValeur('LaTable'   , TobTemp.GetValue('TableName'), False);
        TobLevel2 := TOB.Create(TobTemp.GetValue('Domaine'), TobTV, -1);
        TobLevel2.AddChampSupValeur('NomTable', TobTemp.GetValue('Domaine'), False);
        TobLevel2.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('Domaine'), False);
        TobLevel2.AddChampSupValeur('LaTable'   , TobTemp.GetValue('TableName'), False);
        TobLevel2.AddChampSupValeur('ImageIndex', 7, False);
        end
    else if (TobRest.NomTable = 'Scripts') or (TobRest.NomTable = 'Conception') then
        begin
        TobLevel1 := TOB.Create(TobTemp.GetValue('NomTable'), TobGen, -1);
        TobLevel1.Dupliquer(TobTemp, True, True, True);
        TobLevel1.AddChampSupValeur('NomTable', TobTemp.GetValue('NomTable'), False);
        TobLevel1.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('DOMAINE'), False);
        TobLevel1.AddChampSupValeur('LaTable'   , TobTemp.GetValue('CLE'), False);
        TobLevel2 := TOB.Create(TobTemp.GetValue('NomTable'), TobTV, -1);
        TobLevel2.AddChampSupValeur('NomTable', TobTemp.GetValue('NomTable'), False);
        TobLevel2.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('DOMAINE'), False);
        TobLevel2.AddChampSupValeur('LaTable'   , TobTemp.GetValue('CLE'), False);
        TobLevel2.AddChampSupValeur('ImageIndex', 7, False);
        end
    else if TobRest.NomTable = 'Correspondance' then
        begin
        TobLevel1 := TOB.Create(TobTemp.GetValue('NomTable'), TobGen, -1);
        TobLevel1.Dupliquer(TobTemp, True, True, True);
        TobLevel1.AddChampSupValeur('NomTable', TobTemp.GetValue('NomTable'), False);
        TobLevel1.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('Domaine'), False);
        TobLevel1.AddChampSupValeur('LaTable'   , TobTemp.GetValue('TableName'), False);
        TobLevel2 := TOB.Create(TobTemp.GetValue('Profile') + '-' + TobTemp.GetValue('ChampCode'), TobTV, -1);
        TobLevel2.AddChampSupValeur('NomTable', TobTemp.GetValue('NomTable'), False);
        TobLevel2.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('Domaine'), False);
        TobLevel2.AddChampSupValeur('LaTable'   , TobTemp.GetValue('TableName'), False);
        TobLevel2.AddChampSupValeur('ImageIndex', 7, False);
        end
    else if TobRest.NomTable = 'Dictionnaire' then
        begin
        TobLevel1 := TobTV.FindFirst(['LeDomaine'], [TobTemp.GetValue('Domaine')], False);
        if TobLevel1 = nil then
            begin
            TobLevel1 := TOB.Create(TobTemp.GetValue('Domaine'), TobTV, -1);
            TobLevel1.AddChampSupValeur('NomTable', TobTemp.GetValue('Domaine'), False);
            TobLevel1.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('Domaine'), False);
            TobLevel1.AddChampSupValeur('LaTable'   , '', False);
            TobLevel1.AddChampSupValeur('ImageIndex', 7, False);
            end;
        TobLevel2 := TobLevel1.FindFirst(['LaTable'], [TobTemp.GetValue('TableName')], False);
        if TobLevel2 = nil then
            begin
            TobLevel2 := TOB.Create(TobTemp.GetValue('TableName'), TobLevel1, -1);
            TobLevel2.AddChampSupValeur('NomTable', TobTemp.GetValue('TableName'), False);
            TobLevel2.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('Domaine'), False);
            TobLevel2.AddChampSupValeur('LaTable'   , TobTemp.GetValue('TableName'), False);
            TobLevel2.AddChampSupValeur('ImageIndex', 1, False);
            TobLevel2.AddChampSupValeur('NbFilles', 0, False);
            end;
        TobLevel2.PutValue('NbFilles', TobLevel2.GetValue('NbFilles') + 1);
        TobLevel3 := TOB.Create('', TobGen, -1);
        TobLevel3.Dupliquer(TobTemp, True, True, True);
        TobLevel3.AddChampSupValeur('LeDomaine' , TobTemp.GetValue('Domaine'), False);
        TobLevel3.AddChampSupValeur('LaTable'   , TobTemp.GetValue('TableName'), False);
        end;
    end;
TobTV.PutTreeView(TV, nil, 'NomTable');
AnalyseTV;
end;

{$IFNDEF CISXPGI}
procedure TFSaveRestore.EnregistreTob(TobF : TOB);
var RecordExist : boolean;
    ind1 : integer;
    stTemp : string;
    BField : TBlobField;
    TDelim : TADOTable;
begin
RecordExist := False;

if not assigned(TBase) then TBase := TADOTable.Create(Application);
TBase.Connection := DMImport.Db as TADOConnection;

if TBase.TableName = 'PGZIMPCORRESP' then
begin
    TBase.TableName := DMImport.GzImpCorresp.TableName;
    RecordExist := TBase.Locate('Ident;TableName;Profile', VarArrayOf([TobF.GetValue('Ident'), TobF.GetValue('TableName'), TobF.GetValue('Profile')]), [loCaseInsensitive])
end else if TBase.TableName = 'PGZIMPDOMAINE' then
begin
    TBase.TableName := DMImport.Gzimpdomaine.TableName;
    RecordExist := TBase.Locate('Domaine', TobF.GetValue('Domaine'), [loCaseInsensitive])
end else if TBase.TableName = 'PGZIMPPAR' then
begin
    TBase.TableName := DMImport.GzImpPar.TableName;
    RecordExist := TBase.Locate('Numero;Domaine;TableName', VarArrayOf([TobF.GetValue('Numero'), TobF.GetValue('Domaine'), TobF.GetValue('TableName')]), [loCaseInsensitive])
end else if TBase.TableName = 'PGZIMPREQ' then
begin
    TBase.TableName := DMImport.GzImpReq.TableName;
    RecordExist := TBase.Locate('CLE', TobF.GetValue('Cle'), [loCaseInsensitive]);
end;

    // ajout me car on supprimer les champs des ecritures si on remonte tiers par exemple
if (TBase.TableName = DMImport.GzImpPar.TableName) and (TBase.FieldCount >=2) and RecordExist then
begin
         if (TBase.Fields[2].FieldName = 'TableName') and
          (TBase.Fields[2].Value <> TobF.GetValue(TBase.Fields[2].FieldName)) then
          exit;
end;

if not TBase.active then TBase.Open;
if not RecordExist then
    TBase.Insert
    else
    TBase.Edit;
for ind1 := 0 to TBase.FieldCount - 1 do
    begin
    if (TBase.Fields[ind1].DataType = ftBlob) or (TBase.Fields[ind1].DataType =  ftMemo) then
        begin
        BField := TBlobField(TBase.Fields[ind1]);
        stTemp := TobF.GetValue(TBase.Fields[ind1].FieldName);
        stTemp := StringReplace(stTemp, '&#xA0;', ' ', [rfReplaceAll]);
        stTemp := StringReplace(stTemp, '&', ' ', [rfReplaceAll]);
        stTemp := StringReplace(stTemp, '&apos;', '''', [rfReplaceAll]);
        StringToBlob(stTemp, BField);
        end
    else if TobF.GetValue(TBase.Fields[ind1].FieldName) = '' then
        begin
              if (TBase.Fields[ind1].DataType = ftSmallint) or
                 (TBase.Fields[ind1].DataType = ftInteger) or
                 (TBase.Fields[ind1].DataType = ftWord) or
                 (TBase.Fields[ind1].DataType = ftBytes) or
                 (TBase.Fields[ind1].DataType = ftFloat) then
                  TobF.PutValue(TBase.Fields[ind1].FieldName, 0)
                  //TBase.FieldValues[TBase.Fields[ind1].FieldName] := 0
              else // ajout me
                  TBase.FieldValues[TBase.Fields[ind1].FieldName] := TobF.GetValue(TBase.Fields[ind1].FieldName);
        end
        else
        TBase.FieldValues[TBase.Fields[ind1].FieldName] := TobF.GetValue(TBase.Fields[ind1].FieldName);
    end;
    TBase.Post;
//debug TobF.SaveToFile('XXXX.tmp',True,True,True);
if TBase.TableName = 'PGZIMPREQ' then
begin
    if not TobF.FieldExists('DELIM-CLE') then Exit;
    if (TobF.GetValue('DELIM-CLE')) <> (TBase.FieldByName('CLE').AsString) then exit;
    TDelim := TADOTable.Create(Application);
    TDelim.Connection := DMImport.Db as TADOConnection;
    TDelim.Tablename := DMImport.GzImpDelim.TableName;
    if not TDelim.active then TDelim.Open;
    RecordExist := TDelim.Locate('CLE', TobF.GetValue('DELIM-CLE'), [loCaseInsensitive]);
    if not RecordExist then
        TDelim.Insert
    else
    begin
        // suppression avant import pb cle n'était pas mis sur script délimité
        DMImport.DB.Execute ('Delete from ' + DMImport.GzImpDelim.TableName + ' where CLE="'+TBase.FieldByName('CLE').AsString+'"'
        + ' and DOMAINE="'+TBase.FieldByName('DOMAINE').AsString+'"');
        TDelim.Insert;
    end;
    for ind1 := 0 to TDelim.FieldCount - 1 do
    begin
        if (TDelim.Fields[ind1].DataType = ftBlob) or (TDelim.Fields[ind1].DataType =  ftMemo) then
        begin
            BField := TBlobField(TDelim.Fields[ind1]);
            stTemp := TobF.GetValue('Delim-'+TDelim.Fields[ind1].FieldName);
            stTemp := StringReplace(stTemp, '&#xA0;', ' ', [rfReplaceAll]);
            stTemp := StringReplace(stTemp, '&', ' ', [rfReplaceAll]);
            stTemp := StringReplace(stTemp, '&apos;', '''', [rfReplaceAll]);
            StringToBlob(stTemp, BField);
        end
        else if TobF.GetValue('Delim-'+TDelim.Fields[ind1].FieldName) = '' then
        begin
                  if (TDelim.Fields[ind1].DataType = ftSmallint) or
                     (TDelim.Fields[ind1].DataType = ftInteger) or
                     (TDelim.Fields[ind1].DataType = ftWord) or
                     (TDelim.Fields[ind1].DataType = ftBytes) or
                     (TDelim.Fields[ind1].DataType = ftFloat) then
                      TobF.PutValue(TDelim.Fields[ind1].FieldName, 0)
                      //TDelim.FieldValues[TDelim.Fields[ind1].FieldName] := 0
                  else // ajout me
                      TDelim.FieldValues[TDelim.Fields[ind1].FieldName] := TobF.GetValue('Delim-'+TDelim.Fields[ind1].FieldName);
        end
        else
            TDelim.FieldValues[TDelim.Fields[ind1].FieldName] := TobF.GetValue('Delim-'+TDelim.Fields[ind1].FieldName);
    end;
    TDelim.Post; TDelim.close; TDelim.free;
end;
end;
{$ELSE}
procedure TFSaveRestore.EnregistreTob(TobF : TOB);
var
T1        : TOB;
begin
             T1 := TOB.Create('CPGZIMPREQ', nil, -1);
             T1.putValue ('CIS_CLE', TobF.getvalue('CLE'));
             T1.putValue ('CIS_COMMENT', TobF.getvalue('COMMENT'));
             T1.putValue ('CIS_PARAMETRES', TobF.getvalue('PARAMETRES'));

             T1.putValue ('CIS_CLEPAR', TobF.getvalue('CLEPAR'));
             T1.putValue ('CIS_DOMAINE', TobF.getvalue('DOMAINE'));
             T1.putValue ('CIS_DATEMODIF',FormatDateTime('dd/mm/yyyy', now));
             T1.putValue ('CIS_TABLE1', TobF.getvalue('TABLE1'));
             T1.putValue ('CIS_TABLE2', TobF.getvalue('TABLE0'));
             T1.putValue ('CIS_TABLE3', TobF.getvalue('TABLE3'));
             T1.putValue ('CIS_TABLE4', TobF.getvalue('TABLE4'));
             T1.putValue ('CIS_TABLE5', TobF.getvalue('TABLE5'));
             T1.putValue ('CIS_TABLE6', TobF.getvalue('TABLE6'));
             T1.putValue ('CIS_TABLE7', TobF.getvalue('TABLE7'));
             T1.putValue ('CIS_TABLE8', TobF.getvalue('TABLE8'));
             T1.putValue ('CIS_TABLE9', TobF.getvalue('TABLE9'));
             T1.putValue ('CIS_MODIFIABLE', TobF.getvalue('MODIFIABLE'));
             T1.SetAllModifie(True);
             T1.InsertOrUpdateDB(TRUE);
             if TobF.FieldExists('DELIM-CLE') then
             begin
                  T1.putValue ('CIS_PARAMETRES', TobF.getvalue('DELIM-PARAMETRES'));
                  T1.putValue ('CIS_NATURE', 'X');
                  T1.SetAllModifie(True);
                  T1.InsertOrUpdateDB(TRUE);
             end;
             T1.Free;
             T1 := TOB.Create('CPGZIMPCORRESP', nil, -1);
             T1.putValue ('CIC_TABLEBLOB', 'CIS');
             T1.putValue ('CIC_IDENTIFIANT', TobF.getvalue('CLE'));
             T1.PutValue('CIC_QUALIFIANTBLOB', 'DATA');
             T1.putValue ('CIC_LIBELLE', TobF.getvalue('COMMENT'));
             T1.putValue ('CIC_RANGBLOB', 1);
             T1.putValue ('CIC_OBJET', TobF.getvalue('TBLCOR'));
             T1.PutValue('CIC_PRIVE', '-');
             T1.putValue ('CIC_DATEMODIF', FormatDateTime('dd/mm/yyyy', now));
             if TObF.FieldExists ('MODIFIABLE') and (TobF.getvalue('MODIFIABLE') = 0) then
                T1.putValue ('CIC_CREATEUR', 'CEG')
             else
                T1.putValue ('CIC_CREATEUR', 'STD');
             T1.SetAllModifie(True);
             T1.InsertOrUpdateDB(TRUE);
             T1.Free;

end;
{$ENDIF}

procedure TFSaveRestore.AnalyseTV;
var TN1 : TTreeNode;
    VInt : Variant;
    TobL : TOB;
begin
if TobDomaine = nil then exit;
TN1 := TV.TopItem;
TN1.ImageIndex := 9;
TN1.SelectedIndex := 9;
TN1 := TN1.GetNext;
while TN1 <> nil do
    begin
    if TN1.Level = 1 then
        begin
        TobL := TobDomaine.FindFirst(['Domaine'], [TN1.Text], True);
        if TobL <> nil then
            TN1.Text := TOB(TN1.Data).GetValue('NomTable') + ' (' +
                        TobL.GetValue('Libelle') + ')';
        end;
    if TOB(TN1.Data).FieldExists('ImageIndex') then
        VInt := TOB(TN1.Data).GetValue('ImageIndex')
        else
        VInt := 4;
    TN1.ImageIndex := VarAsType(VInt,VarInteger);
    TN1 := TN1.GetNext;
    end;
end;

//==============================================================================
procedure TOB_2.LoadFromXMLFile(FName : String);
begin
    TOBLoadFromXMLFile(FName, YouzTob);
end;

function TOB_2.YouzTob(T : TOB) : integer;
begin
    T.ChangeParent(Self, -1);
    T.SetAllModifie(true);
    result := 0;
end;


procedure TFSaveRestore.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.

