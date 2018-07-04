unit UMulScript;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls,
  DB, Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  hmsgbox, ExtCtrls, HSysMenu, HDB, HTB97,
  HPanel, UiUtil, HEnt1, UPDomaine, HStatus, HRichEdt, ColMemo, ComCtrls,
  HRichOLE, Menus, Dialogs, PrintDbg, Mask, Mul,
  uDbxDataSet, Variants, ADODB, Forms, UTob;

function MulticritereScript (PP: THPanel; Var Dom : string) : String;

type
  TFMulScript = class(TFMul)
    HM: THMsgBox;
    QDOMAINE: TStringField;
    QCLE: TStringField;
    QCOMMENT: TStringField;
    QCLEPAR: TStringField;
    CLE: TEdit;
    XX_WHERE: TEdit;
    BDelete: TToolbarButton97;
    TDOMAINE: TLabel;
    QDATEDEMODIF: TDateField;
    Label2: TLabel;
    Label3: TLabel;
    CLEPAR: TEdit;
    Label4: TLabel;
    Comment: TEdit;
    BDelim: TToolbarButton97;
    QTable0: TStringField;
    QTable1: TStringField;
    Label1: TLabel;
    Label5: TLabel;
    Editeur: TEdit;
    Nature: TComboBox;
    ToolbarButton971: TToolbarButton97;
    HFindDialog1: THFindDialog;
    Domaine: THValComboBox;
    BSauve: TToolbarButton97;
    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure VisibleClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BDelimClick(Sender: TObject);
    procedure DOMAINEChange(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BSauveClick(Sender: TObject);
  private
    { Déclarations privées }
    CodeCle, DPar : string;
    ModeSelect    : Boolean;
    OkOuv         : Boolean;
    procedure SuppressionScript;
    procedure InsereDansTobTable (TobGen : TOB);
    function  BlobToString(Blob: TBlobField): string;
    procedure StringToBlob(Value: string; var Field : TBlobField);
  public
    { Déclarations publiques }
  end;
var
  FMuls: TFMulScript;

implementation

{$R *.DFM}
uses UConcept, Uscript, UDMIMP, UAssistConcept;


function MulticritereScript (PP: THPanel; Var Dom : string) : String;
begin
 Result := '';
 FMuls := TFMulScript.Create(Application);
 FMuls.ModeSelect  := FALSE;
 FMuls.OkOuv       := FALSE;
 FMuls.DPar        := '';
 if (Dom <> '') and (Dom <> '0') then FMuls.DPar   := Dom;
 PP:=FindInsidePanel ;
 if PP=Nil then
 BEGIN
    FMuls.ModeSelect := TRUE;
    Try
     FMuls.ShowModal ;
    Finally
          if FMuls.OkOuv then
          begin
               Result := FMuls.Q.FindField ('CLE').asstring;
               Dom    := FMuls.Q.FindField ('DOMAINE').asstring;
          end;
     FMuls.Free ;
    End ;
 END else
 BEGIN
   InitInside(FMuls,PP) ;
   FMuls.Show ;
 END ;
end;

procedure TFMulScript.BOuvrirClick(Sender: TObject);
begin
  inherited;
  OkOuv := TRUE;
  FListeDblClick(Sender);
end;

procedure TFMulScript.FListeDblClick(Sender: TObject);
var
  Table, Dm, Nat: string;
  lib,Nt,edt    : string;
begin
  inherited;
  if ModeSelect then
  begin
       ModalResult := MrOk;
       OkOuv       := TRUE;
       exit;
  end;
  Table := Q.FindField('CLE').AsString;
  Dm    := Q.FindField('Domaine').AsString;
  Nat   := Q.FindField('CLEPAR').AsString;
  lib   := RendLibelleDomaine(Dm);
  Nt   := Q.FindField('Table0').AsString;
  Edt   := Q.FindField('Table1').AsString;
  PExtConcept(nil, Lib, Table, Nat, Nt, Edt, taModif);

end;

procedure TFMulScript.BinsertClick(Sender: TObject);
begin
  inherited;
    if (ParamCount > 0) and ((UpperCase(GetInfoVHCX.Mode) = 'EXPORT') or (UpperCase(GetInfoVHCX.Mode) = 'IMPORT')) then   exit;
  CreateScript (FALSE, tacreat);
  BChercheClick(Sender);
end;

procedure TFMulScript.FormCreate(Sender: TObject);
begin
  inherited;
  ChargementComboDomaine (Domaine);
  BImprimer.Enabled := FALSE;
end;

procedure TFMulScript.BChercheClick(Sender: TObject);
var
  SelectQL : string;
  SavePlace: TBookmark;
  St,DD    : string;
begin
//  SelectQL := ' SELECT Domaine,CLE,COMMENT,CLEPAR,DATEDEMODIF from '+ DMImport.GzImpReq.TableName;

  SelectQL := ' SELECT Table1,CLE,COMMENT,CLEPAR,DATEDEMODIF,Table0,Domaine from '+ DMImport.GzImpReq.TableName;
  Q.Close;
  Q.SQL.Clear;
  St := ' Where (CLEPAR<>"SQL" or CLEPAR="") ';
  if DPar <> '' then
  begin
     DD          := DPar;
     if Domaine.text = '' then
          Domaine.itemindex  := Domaine.items.IndexOf(RendLibelleDomaine(Dpar));
     if GetInfoVHCX.complement <> '' then Clepar.Text := GetInfoVHCX.complement;
     if GetInfoVHCX.Nature <> '' then
          Nature.Itemindex :=  Nature.items.IndexOf(GetInfoVHCX.Nature);
  end
  else
     DD :=  RendDomaine(Domaine.Text);
  if Cle.Text <> ''     then St := St + ' and CLE like "' + Cle.Text + '%" ';
  if (DD <> '') and (Domaine.Text <> '<Tous>') then St := St + ' and Domaine="' + DD + '" ';
  if Comment.Text <> '' then St := St + ' and Comment like "%' + Comment.Text + '%" ';
  if Clepar.Text <> ''  then St := St + ' and CLEPAR like "%' + Clepar.Text + '%" ';
  if (Nature.Text <> '') and (Nature.Text <> '<Tous>')  then St := St + ' and Table0 like "%' + Nature.Text + '%" ';
  if Editeur.Text <> '' then St := St + ' and Table1 like "' + Editeur.Text + '%" ';

  Q.SQL.Add(SelectQL + St + ' ORDER BY Table1,CLE');
  Q.SqlConnection := DMImport.Db ; // XP 07.01.2008
  //ChangeSQL(Q);
  Q.Open;
  Q.UpdateCumul(PCumul, FListe);
  Z_SQL.Lines.Clear;
  Z_SQL.Lines.Add(Q.SQL[0]);
  if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then SavePlace:=Q.GetBookmark else SavePlace:=Nil ;
  if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then if SavePlace<>Nil then
   BEGIN
       Try
       Q.GotoBookmark(SavePlace) ;
       Q.FreeBookmark(SavePlace) ;
       Except
       End ;
   END ;
  HMTrad.ResizeDBGridColumns(FListe);
  CentreDBGrid(FListe);
end;


procedure TFMulScript.SuppressionScript;
var
   tbcharger: TADOTable;
begin
  tbCharger := TADOTable.Create(Application);
  tbCharger.Connection := DMImport.Db as TADOConnection;
  tbCharger.TableName := DMImport.GzImpReq.TableName;

  try
    tbCharger.open;
    if tbCharger.Locate('CLE', CodeCle, [loCaseInsensitive]) then
    begin
      tbCharger.Delete;
    end;
  except
    PGIBox(Exception(ExceptObject).message, '');
    exit;
  end;
  tbcharger.Close; tbcharger.free;

  tbcharger := TADOTable.Create(Application);
  tbcharger.Connection := DMImport.Db as TADOConnection;
  tbcharger.TableName := DMImport.GzImpDelim.TableName;

  try
    tbCharger.open;
    if tbCharger.Locate('CLE', CodeCle, [loCaseInsensitive]) then
    begin
      tbCharger.Delete;
    end;
  except
  end;
  tbcharger.Close;  tbcharger.free;
end;

procedure TFMulScript.BDeleteClick(Sender: TObject);
var
Texte : string;
i     : integer;
begin
  inherited;

  if (ParamCount > 0) and ((UpperCase(GetInfoVHCX.Mode) = 'EXPORT') or (UpperCase(GetInfoVHCX.Mode) = 'IMPORT')) then   exit;

  if (Fliste.NbSelected=0) and (not Fliste.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;
  Texte:='Vous allez supprimer définitivement les Scripts.#13#10Confirmez vous l''opération ?';
  if HShowMessage('0;Suppression;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;
  if Fliste.AllSelected then
  BEGIN
    Q.First;
    while Not Q.EOF do
    BEGIN
      MoveCur(False);
      CodeCle := Q.FindField('Cle').AsString;
      SuppressionScript;
      Q.Next;
    END;
    Fliste.AllSelected:=False;
  END
  ELSE
  BEGIN
    InitMove(Fliste.NbSelected,'');
    for i:=0 to Fliste.NbSelected-1 do
    BEGIN
      MoveCur(False);
      Fliste.GotoLeBookmark(i);
      CodeCle := Q.FindField('Cle').AsString;
      SuppressionScript;
    END;
    Fliste.ClearSelected;
  END;
  FiniMove;
  BChercheClick(Sender);

end;



procedure TFMulScript.VisibleClick(Sender: TObject);
begin
  inherited;
  BChercheClick(Sender);
end;

procedure TFMulScript.FormShow(Sender: TObject);
begin
  inherited;
  PComplement.TabVisible:=FALSE;
  if (Codeseria = '') or ((ParamCount > 0) and ((UpperCase(GetInfoVHCX.Mode) = 'EXPORT') or (UpperCase(GetInfoVHCX.Mode) = 'IMPORT'))) then
  begin
       Binsert.visible := FALSE;
       BDelete.visible := FALSE;
       BDelim.visible  := FALSE;
       BImprimer.visible := FALSE;
  end
  else
  begin
       Binsert.visible := TRUE;
       BDelete.visible := TRUE;
       BDelim.visible  := TRUE;
       BImprimer.visible := TRUE;
  end;
  if  (ParamCount > 0) and (GetInfoVHCX.Domaine <> '') then
  begin
          DOMAINEChange (Sender);
          Domaine.Enabled := FALSE;
          if GetInfoVHCX.complement <> '' then Clepar.Enabled := FALSE;
          if GetInfoVHCX.Nature <> '' then Nature.Enabled := FALSE;
  end;
end;

procedure TFMulScript.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
    VK_DELETE : BDeleteClick (Sender);
    VK_INSERT : BinsertClick (Sender);
  end;

end;


procedure TFMulScript.BDelimClick(Sender: TObject);
var
  Table, Nat    : string;
  Nt,edt,Dm     : string;
begin
  inherited;
  Table := Q.FindField('CLE').AsString;
  Nat   := Q.FindField('CLEPAR').AsString;
  Nt    := Q.FindField('Table0').AsString;
  Edt   := Q.FindField('Table1').AsString;
  Dm    := Q.FindField('Domaine').AsString;

  ModifScriptDelim(Table, Nat, Nt, Edt, Dm);
end;

procedure TFMulScript.DOMAINEChange(Sender: TObject);
begin
  inherited;
    if (RendDomaine(Domaine.Text) = 'O') then
    begin
      Nature.clear;
      Nature.Items.add ('<Tous>');
      Nature.Items.add ('Balance');
      Nature.Items.add ('Ecriture');
    end
    else
    begin
       Nature.clear;
       Nature.Items.add ('<Tous>');
       Nature.Items.add ('Dossier');
       Nature.Items.add ('Journal');
       Nature.Items.add ('Balance');
       Nature.Items.add ('Synchronisation');
    end;

end;

procedure TFMulScript.BAnnulerClick(Sender: TObject);
begin
  inherited;
  Close;
  if IsInside(Self) then THPanel(parent).CloseInside;
end;

procedure TFMulScript.BSauveClick(Sender: TObject);
var
    i       : integer;
    FName, Fna : string;
    SD      : TSaveDialog;
    s       : TMemoryStream;
    TobSave : TOB;
begin
  inherited;

    // si rien de selectionné !
    if ((FListe.NbSelected = 0) and (not FListe.AllSelected)) then
    begin
        MessageAlerte('Aucun élément sélectionné.');
        exit;
    end;

    // message d'avertisement !
    SD := TSaveDialog.Create(Application);
    SD.DefaultExt := '*.xls';
    SD.Title := 'Extraction des rapports';
    SD.Filter := 'Fichier Texte (*.cix)|*.cix';
    if SD.execute then
      Fna := SD.FileName;
    SD.free;
    FName := ReadTokenPipe (Fna, '.')+ '.cix';


    TobSave := TOB.Create('Scripts', nil, -1);

    // destruction
    if (FListe.AllSelected) then // si tout ??
    begin
        Q.First;

        while (Not Q.EOF) do
        begin
            InsereDansTobTable (TobSave);
            Q.Next;
        end;

        FListe.AllSelected:=False;
    end
    else // si pas tout ....
    begin
        for i := 0 to FListe.NbSelected-1 do
        begin
            FListe.GotoLeBookMark(i);
            {$IFDEF EAGLCLIENT}
            FListe.Q.TQ.Seek(L.Row-1);
            {$ENDIF}
            InsereDansTobTable (TobSave);
        end;

        FListe.ClearSelected;
    end;
    TobSave.SaveToXMLFile(FName, True, True);
    TobSave.free;
    BChercheClick(nil) ;
end;

procedure TFMulScript.InsereDansTobTable (TobGen : TOB);
var ind1                                   : integer;
    TobL, TobChamp                         : TOB;
    FNumero, FLocDomaine, FLocTable, FNom, FNomChamp : TStringField;
    stBlob                                 : string;
    TDelim, TBase                          : TADOTable;
begin
TobL := nil;
    TobChamp := nil;
    // sauvegarde delimitée
    TDelim := TADOTable.Create(Application);
    TDelim.Connection := DMImport.Db as TADOConnection;
    TDelim.TableName := DMImport.GzImpDelim.TableName;

    with TDelim do
    begin
         Open;
         if TDelim.Locate('CLE', Q.FieldByName('CLE').AsString , [loCaseInsensitive]) then
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
    TDelim.close; TDelim.free;

    TBase := TADOTable.Create(Application);
    TBase.Connection := DMImport.Db as TADOConnection;
    TBase.Tablename := DMImport.GzImpReq.TableName;
    TBase.Open;
    if TBase.Locate('CLE', Q.FieldByName('CLE').AsString , [loCaseInsensitive]) then
          for ind1 := 0 to TBase.FieldCount - 1 do
              begin
              if ind1 = 0 then
                  begin
                  TobL := TOB.Create(TBase.Fields[ind1].AsString, TobGen, -1);
                  TobL.AddChampSupValeur('NomTable', TBase.Fields[ind1].AsString, False);
                  end;
              if TBase.Fields[ind1].FieldName = 'DOMAINE' then
                  begin
                  TobL.AddChampSupValeur('LeDomaine' , TBase.Fields[ind1].AsString, False);
                  end;
              if TBase.Fields[ind1].FieldName = 'CLE' then
                  begin
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
          end;
          TBase.close; TBase.free;
end;

function TFMulScript.BlobToString(Blob: TBlobField): string;
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

procedure TFMulScript.StringToBlob(Value: string; var Field : TBlobField);
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

end.

