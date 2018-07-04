unit AssistPieceEpure;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, ComCtrls, ExtCtrls, HTB97, Hctrls,
  UTOB, HPanel, Grids, EntGC, HEnt1, UtilPGI, HStatus, HRichEdt, HRichOLE;
{$IFDEF EAGLCLIENT}
{$ELSE}
{$ENDIF}

type
  TFAssistantPieceEpure = class(TFAssist)
    TabGrid: TTabSheet;
    TabSheet2: TTabSheet;
    TLBRecapitulatif: TListBox;
    PGENERAL: THPanel;
    PTITRE: THPanel;
    GListePieces: THGrid;
    TDescriptif: TToolWindow97;
    Descriptif: THRichEditOLE;
    bNote: TToolbarButton97;
    procedure bFinClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bNoteClick(Sender: TObject);
    procedure GListePiecesRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure TDescriptifClose(Sender: TObject);
    procedure GListePiecesDblClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    TobPieceAssist, TobJournal : Tob;
    stNaturePieceG : string;
    bFinFait : boolean;
    procedure EpureLesPieces;
    function ConstruitWhere (stPrefixe : string) : string;
  end;

var
  FAssistantPieceEpure: TFAssistantPieceEpure;

implementation
{$R *.DFM}

function TFAssistantPieceEpure.ConstruitWhere (stPrefixe : string) : string;
var iInd : integer;
    TobLoc : Tob;
    stWhere : string;
    stCleAna : string;
begin
stWhere := '';
if stPrefixe = 'GP_' then
    begin
    for iInd := 0 to TobPieceAssist.Detail.Count - 1 do
        begin
        TobLoc := TobPieceAssist.Detail[iInd];
        if not TobLoc.GetValue('EPURABLE') then continue;
        if stWhere <> '' then stWhere := stWhere + ' OR ';
        stWhere := stWhere + '(' + stPrefixe + 'SOUCHE="'  + TobLoc.GetValue ('GP_SOUCHE') + '"' +
                           ' AND ' + stPrefixe + 'NUMERO=' + IntToStr (TobLoc.GetValue ('GP_NUMERO')) + ')';
        end;
    end else
    begin
    for iInd := 0 to TobPieceAssist.Detail.Count - 1 do
        begin
        TobLoc := TobPieceAssist.Detail[iInd];
        if not TobLoc.GetValue('EPURABLE') then continue;
        //stCleAna := EncodeRefCPGescom (TobLoc);
        stCleAna := TobLoc.GetValue('GP_NATUREPIECEG')+';'+TobLoc.GetValue('GP_SOUCHE')+';'
                 +FormatDateTime('ddmmyyyy',TobLoc.GetValue('GP_DATEPIECE'))+';'
                 +IntToStr(TobLoc.GetValue('GP_NUMERO'))+';';

        if stWhere <> '' then stWhere := stWhere + ' OR ';
        stWhere := stWhere + 'YVA_IDENTIFIANT LIKE "' + stCleAna + '%"';
        end;

    if stWhere <> '' then stWhere := '(' + stWhere + ')';
    end;
Result := stWhere;
end;

procedure TFAssistantPieceEpure.EpureLesPieces;
var stWhere, stClause : String;
    iNbEnreg : integer;
begin
  stWhere := ConstruitWhere ('GP_');
  stClause := 'DELETE FROM PIECE WHERE ' +
                  'GP_NATUREPIECEG="' + stNaturePieceG + '"';
  if stWhere <> '' then
  begin
    stClause := stClause + ' AND (' + stWhere + ')';
    InitMove (9, '');
    iNbEnreg := ExecuteSql (stClause);
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table PIECE supprimé(s)');
    MoveCur (False);
    stClause := StringReplace (stClause, 'GP_', 'GL_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'PIECE ', 'LIGNE ', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause);
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table LIGNE supprimé(s)');
    MoveCur (False);

    stClause := StringReplace (stClause, 'GL_', 'GLF_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNE', 'LIGNEFORMULE', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // LIGNELOT
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table LIGNEFORMULE supprimé(s)');
    MoveCur (False);

    stClause := StringReplace (stClause, 'GLF_', 'GLL_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNEFORMULE', 'LIGNELOT', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // LIGNELOT
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table LIGNELOT supprimé(s)');
    MoveCur (False);

    stClause := StringReplace (stClause, 'GLL_', 'GLS_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNELOT', 'LIGNESERIE', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // LIGNESERIE
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table LIGNESERIE supprimé(s)');
    MoveCur (False);

    stClause := StringReplace (stClause, 'GLS_', 'GLN_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNESERIE', 'LIGNENOMEN', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // LIGNENOMEN
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table LIGNENOMEN supprimé(s)');
    MoveCur (False);
    stClause := StringReplace (stClause, 'GLN_', 'GPA_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'LIGNENOMEN', 'PIECEADRESSE', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // PIECEADRESSE
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table LIGNENOMEN supprimé(s)');
    MoveCur (False);
    stClause := StringReplace (stClause, 'GPA_', 'GPB_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'PIECEADRESSE', 'PIEDBASE', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // PIEDBASE
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table PIEDBASE supprimé(s)');
    MoveCur (False);
    stClause := StringReplace (stClause, 'GPB_', 'GPE_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'PIEDBASE', 'PIEDECHE', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // PIEDECHE
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table PIEDECHE supprimé(s)');
    MoveCur (False);
    stClause := StringReplace (stClause, 'GPE_', 'GPT_', [rfReplaceAll]);
    stClause := StringReplace (stClause, 'PIEDECHE', 'PIEDPORT', [rfReplaceAll]);
    iNbEnreg := ExecuteSql (stClause); // PIEDPORT
    if iNbEnreg <> 0 then
        TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table PIEDPORT supprimé(s)');
    MoveCur (False);
    if (GetInfoParpiece (stNaturePieceG,
                         'GPP_COMPANALLIGNE') <> 'SAN') or
       (GetInfoParPiece (stNaturePieceG,
                         'GPP_COMPANALPIED') <> 'SAN') then
        begin
          stWhere := ConstruitWhere ('YVA_');
          if stWhere <> '' then
          begin
          stClause := 'DELETE FROM VENTANA WHERE ' + stWhere;
           iNbEnreg := ExecuteSql (stClause); // VENTANA
          if iNbEnreg <> 0 then
            TLBRecapitulatif.Items.Add (IntToStr (iNbEnreg) + ' Enregistrement(s) de la table VENTANA supprimé(s)');
          MoveCur (False);
          end;
        end;
    FiniMove;
  end;
end;

procedure TFAssistantPieceEpure.bFinClick(Sender: TObject);
var ioerr : TIOErr ;
begin
  inherited;
if not bFinFait then
    begin
    //if TobPieceAssist.Detail.Count > 0 then
    if TobPieceAssist.FindFirst(['EPURABLE'],[True],False) <> nil then
        begin
        ioErr := Transactions (EpureLesPieces, 1);
        TLBRecapitulatif.Items.Add('');
        Case ioerr of
                oeOk  : begin
                        TOBJournal.PutValue('GEV_ETATEVENT', 'OK');
                        TLBRecapitulatif.Items.Add(
                            FloatToStr(TobPieceAssist.Somme('EPURABLE',['EPURABLE'],[True],False,True)) + ' "' +
                            GetInfoParPiece (stNaturePieceG, 'GPP_LIBELLE') + '" épurés(ées)');
                        end;
            oeUnknown : begin
                        TOBJournal.PutValue('GEV_ETATEVENT', 'ERR');
                        TLBRecapitulatif.Items.Add ('Erreur lors de l''épuration des pièces');
                        end;
            end;
        end else
        begin
        TLBRecapitulatif.Items.Add ('Aucune pièce à épurer');
        end;
    bFinFait := True;
    bFin.Caption := 'Terminer';
    bAnnuler.Enabled := False;
    end else
    begin
    Close;
    end;
end;

procedure TFAssistantPieceEpure.bAnnulerClick(Sender: TObject);
begin
  inherited;
if not bFinFait then
    begin
    TLBRecapitulatif.Items.Add ('Abandon de l''épuration par l''utilisateur');
    TOBJournal.PutValue('GEV_ETATEVENT', 'ERR');
    bFinFait := True;
    end;
end;

procedure TFAssistantPieceEpure.FormCreate(Sender: TObject);
begin
  inherited;
bFinFait := False;
end;

procedure TFAssistantPieceEpure.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
if not bFinFait then
    begin
    TLBRecapitulatif.Items.Add ('Abandon de l''épuration par l''utilisateur');
    TOBJournal.PutValue('GEV_ETATEVENT', 'ERR');
    bFinFait := True;
    end;
end;

procedure TFAssistantPieceEpure.FormShow(Sender: TObject);
var  Cancel : boolean;
begin
  inherited;
PTITRE.Caption := 'Epuration des "' +
                    GetInfoParPiece (stNaturePieceG, 'GPP_LIBELLE') + '"';
bFin.Enabled := False;
GListePiecesRowEnter(nil,1,Cancel,False);
end;

procedure TFAssistantPieceEpure.bSuivantClick(Sender: TObject);
begin
  inherited;
  if bFinFait then bAnnuler.Enabled := False;
  if (bSuivant.Enabled) then bFin.Enabled := False
  else bFin.Enabled := True;
  if P.ActivePage <> TabGrid then
  begin
    bNote.Enabled := False;
    bNote.Down := False;
    TDescriptif.Visible := False;
  end else bNote.Enabled := True;
end;

procedure TFAssistantPieceEpure.bPrecedentClick(Sender: TObject);
begin
  inherited;
  if bFinFait then bAnnuler.Enabled := False;
  if (bSuivant.Enabled) and (not bFinFait) then bFin.Enabled := False
  else bFin.Enabled := True;
  if P.ActivePage <> TabGrid then
  begin
    bNote.Enabled := False;
    bNote.Down := False;
    TDescriptif.Visible := False;
  end else bNote.Enabled := True;
end;


procedure TFAssistantPieceEpure.bNoteClick(Sender: TObject);
begin
  inherited;
  TDescriptif.Visible := bNote.Down;
end;

procedure TFAssistantPieceEpure.GListePiecesRowEnter(Sender: TObject;
  Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var TOBP : TOB;
    Msg,Lg : string;
begin
  inherited;
  TOBP := TOB(GListePieces.Objects[0,Ou]);
  if TOBP = nil then exit;
  Descriptif.Text := '';
  Msg := TOBP.GetValue('BLOCNOTE');
  if pos('|',Msg) > 0 then
  begin
    Repeat
      Lg := ReadTokenPipe(Msg,'|');
      if Lg <> '' then Descriptif.Lines.Add(Lg);
    Until Msg = '';
  end else StringToRich(Descriptif,Msg);
end;

procedure TFAssistantPieceEpure.TDescriptifClose(Sender: TObject);
begin
  inherited;
  bNote.Down := TDescriptif.Visible;
end;

procedure TFAssistantPieceEpure.GListePiecesDblClick(Sender: TObject);
begin
  inherited;
  bNote.Down := true;
  bNoteClick(nil);
end;

end.
