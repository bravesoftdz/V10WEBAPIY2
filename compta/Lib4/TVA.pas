unit TVA;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons, Dialogs,
  StdCtrls, ExtCtrls, Hctrls,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB, Grids, DBGrids, Mask, DBCtrls,
  Ent1, HDB, hmsgbox, Hcompte, sysutils, PrintDBG, HEnt1, HSysMenu, Hqry, Spin,
  HPanel, UiUtil, HTB97, ToolWin, ParamSoc
  {$IFNDEF GCGC}
  {$IFNDEF COMPTAPAIE}
  ,CPGeneraux_TOM, ADODB
  {$ENDIF}
  {$ENDIF}
  ,UentCommun
  ;

procedure ParamTVATPF(OkTva: boolean);

type
  TFTva = class(TForm)
    TTaux: THTable;
    STaux: TDataSource;
    TTVA: THTable;
    STVA: TDataSource;
    FListe2: THDBGrid;
    DBNav: TDBNavigator;
    FListe: THDBGrid;
    Panel1: TPanel;
    HB1: TPanel;
    TCC_CODE: THLabel;
    TCC_LIBELLE: THLabel;
    CC_CODE: TDBEdit;
    CC_LIBELLE: TDBEdit;
    MsgBox: THMsgBox;
    Cache: THCpteEdit;
    HT: THMsgBox;
    TTauxTV_TVAOUTPF: TStringField;
    TTauxTV_CODETAUX: TStringField;
    TTauxTV_REGIME: TStringField;
    TTauxTV_TAUXACH: TFloatField;
    TTauxTV_TAUXVTE: TFloatField;
    TTauxTV_CPTEACH: TStringField;
    TTauxTV_CPTEVTE: TStringField;
    TTVACC_TYPE: TStringField;
    TTVACC_CODE: TStringField;
    TTVACC_LIBELLE: TStringField;
    TTVACC_ABREGE: TStringField;
    HMTrad: THSystemMenu;
    TTauxTV_ENCAISACH: TStringField;
    TTauxTV_ENCAISVTE: TStringField;
    CC_LIBRE: TDBEdit;
    TCC_LIBRE: THLabel;
    TTVACC_LIBRE: TStringField;
    FAutoSave: TCheckBox;
    Dock: TDock97;
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BValider: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BFerme: TToolbarButton97;
    TTauxTV_CPTVTERG: TStringField;
    BTVAEtab: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure TTVABeforeDelete(DataSet: TDataset);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BTVAEtabClick(Sender: TOBject);
    procedure STVADataChange(Sender: TObject; Field: TField);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BImprimerClick(Sender: TObject);
    procedure TTauxAfterDelete(DataSet: TDataSet);
    procedure TTauxAfterPost(DataSet: TDataSet);
    procedure TTVAAfterPost(DataSet: TDataSet);
    procedure TTVAAfterDelete(DataSet: TDataSet);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure STVAUpdateData(Sender: TObject);
    procedure STVAStateChange(Sender: TObject);
    procedure FListe2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TTauxNewRecord(DataSet: TDataSet);
    procedure STauxStateChange(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure STauxDataChange(Sender: TObject; Field: TField);
    procedure STauxUpdateData(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure CC_LIBREExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private { Private declarations }
    FCodeTva, W_W2: string;
    FAvertir, Modifier, Modifier1: Boolean;
    Insere, UnNouveau: Boolean;
    LiRegime: HTStrings;
    procedure InitRegime;
    function ChercheCpte(T: TField; Vide: Boolean): byte;
    procedure ChargeEnreg;
    function EnregOK: boolean;
    function EnregOK1: boolean;
    function OnSauve: boolean;
    function Bouge(Button: TNavigateBtn): boolean;
//    function OnSauve1: boolean;
    procedure NewEnreg;
    function MajCodeTauxSurInsert: Boolean;
    function SiExisteCode: Boolean;
    function SiExisteIndice: boolean;
    procedure SupprimeTaux;
    procedure GeleLesBoutons;
    function Supprime: Boolean;
    function FabReqTva(Cpt: string): string;
//    procedure CacheZonesAchat;
  public { Public declarations }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  UtilPgi, TVAEtab;

procedure TFTva.GeleLesBoutons;
begin
  BInsert.Enabled := (not (TTVA.State in [dsEdit, dsInsert])) and (not (TTaux.State in [dsEdit, dsInsert]));
  BDelete.Enabled := (not (TTVA.State in [dsEdit, dsInsert])) and (not (TTaux.State in [dsEdit, dsInsert]));
  if (TTVA.Eof) and (TTVA.Bof) then BDelete.Enabled := False;
end;

(*
procedure TFTva.CacheZonesAchat;
var i: integer;
begin
  for i := 0 to FListe.Columns.Count - 1 do
  begin
    if (FListe.Columns[i].FieldName = 'TV_TAUXACH') or
      (FListe.Columns[i].FieldName = 'TV_CPTEACH') or
      (FListe.Columns[i].FieldName = 'TV_ENCAISACH') then FListe.Columns[i].Visible := False;
  end;
end;
*)

procedure TFTva.FormShow(Sender: TObject);
var i: Byte;
  St: string;
  QLoc: TQuery;
begin
  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
  FAvertir := False;
  Insere := False;
  Modifier1 := False;
  if FCodeTVA = VH^.DefCatTPF then
  begin
    FListe.Columns[3].Title.Caption := HT.Mess[0];
    FListe.Columns[4].Title.Caption := HT.Mess[1];
    FListe.Columns[5].Title.Caption := HT.Mess[2];
    FListe.Columns[6].Title.Caption := HT.Mess[3];
    HB1.Height := 36;
    CC_LIBRE.Visible := False;
    TCC_LIBRE.Visible := False;
    Caption := HT.Mess[4];
  end;
  {$IFDEF ESP}
  FListe.Columns[5].Visible:=FALSE ; //TV_ENCAISACH
  FListe.Columns[6].Visible:=FALSE ; //TV_ENCAISVTE
  HMtrad.ResizeDBGridColumns(Fliste) ;
  {$ENDIF} //XMG 26/11/03
  TTaux.Open;
  TTva.Open;
  QLoc := OpenSql('Select CC_CODE From CHOIXCOD Where CC_TYPE="RTV" Order by CC_CODE', True,-1,'',true);
  LiRegime := HTStringList.Create;
  LiRegime.Clear;
  while not QLoc.Eof do
  begin
    LiRegime.Add(QLoc.Fields[0].AsString);
    QLoc.Next;
  end;
  Ferme(QLoc);
  TTva.SetRange([FCodeTVA], [FCodeTVA]);
  // FQ 12915
  {$IFDEF MODE}
  BTVAEtab.Visible := True;
  {$ENDIF}
  St := '';
  for i := 1 to VH^.Cpta[fbGene].Lg do St := St + 'a';
  TTauxTV_CPTEACH.EditMask := '>' + St + ';0; ';
  TTauxTV_CPTEVTE.EditMask := '>' + St + ';0; ';
  if (TTVA.Eof) and (TTVA.Bof) then Bouge(nbInsert);
  //if ctxScot in V_PGI.PGIContexte then CacheZonesAchat;
end;

procedure TFTVA.InitRegime;
var StQuoi, StTVA, StREG, MemoStQuoi, MemoStTVA: string[3];
  i: Integer;
begin
  if ((TTVACC_TYPE.AsString <> VH^.DefCatTVA) and (TTVACC_TYPE.AsString <> VH^.DefCatTPF)) then Exit;
  StQuoi := TTVACC_TYPE.AsString;
  StTVA := TTVACC_CODE.AsString;
  if (StTVA <> '') or (UnNouveau) then
  begin
    if (TTVA.State = dsInsert) and (StTVA = '') then StTVA := W_W2;
    TTaux.SetRange([StQuoi, StTVA], [StQuoi, StTVA]);
    MemoStQuoi := StQuoi;
    MemoStTVA := StTVA;
    for i := 0 to LiRegime.Count - 1 do
    begin
      StREG := LiRegime.Strings[i];
      if not FindLaKey(TTaux, [StQuoi, StTVA, StREG]) then
      begin
        Insere := True;
        TTaux.Insert;
        InitNew(TTaux);
        TTauxTV_TVAOUTPF.AsString := StQuoi;
        TTauxTV_CODETAUX.AsString := StTVA;
        TTauxTV_REGIME.AsString := StREG;
        TTaux.Post;
      end else Insere := False;
    end;
  end;
  TTaux.First;
  Insere := False;
end;

function TFTva.ChercheCpte(T: TField; Vide: Boolean): byte;
var St: string;
begin
  ChercheCpte := 0;
  St := UpperCase(T.AsString);
  Cache.Text := St;
  if ((Vide) and (St = '')) then exit;
  {$IFNDEF GCGC}
  {$IFNDEF COMPTAPAIE}
  if GChercheCompte(Cache, FicheGene) then
  {$ELSE}
  if GChercheCompte(Cache, nil) then
  {$ENDIF}
  {$ELSE}
  if GChercheCompte(Cache, nil) then
  {$ENDIF}
  begin
    if St <> Cache.Text then
    begin
      if not (TTaux.State in [dsEdit]) then TTaux.Edit;
      T.AsString := Cache.Text;
    end;
    ChercheCpte := 2;
  end;
end;

procedure TFTva.TTVABeforeDelete(DataSet: TDataset);
begin
  SupprimeTaux;
end;

procedure TFTva.SupprimeTaux;
begin
  if TTaux.State <> dsBrowse then TTaux.Cancel;
  ExecuteSql('Delete From TXCPTTVA Where TV_CODETAUX="' + W_W2 + '" Or TV_CODETAUX="' + TTVACC_CODE.AsString + '"');
end;

procedure TFTva.BFirstClick(Sender: TObject);
begin
  Bouge(nbFirst);
end;

procedure TFTva.BPrevClick(Sender: TObject);
begin
  Bouge(nbPrior);
end;

procedure TFTva.BNextClick(Sender: TObject);
begin
  Bouge(nbNext);
end;

procedure TFTva.BLastClick(Sender: TObject);
begin
  Bouge(nbLast);
end;

procedure TFTva.BInsertClick(Sender: TObject);
begin
  fliste.Enabled := false;
  Bouge(nbInsert);
end;

procedure TFTva.BDeleteClick(Sender: TObject);
begin
  Bouge(nbDelete);
end;

procedure TFTva.BValiderClick(Sender: TObject);
begin
  Modifier := False;
  Bouge(nbPost);
  if (not fliste.enabled) then fliste.Enabled := true;
end;

procedure TFTva.BAnnulerClick(Sender: TObject);
begin
  if TTva.State <> dsInsert then TTaux.Cancel;
  Bouge(nbCancel);
end;

procedure TFTva.BFermeClick(Sender: TObject);
begin
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

procedure TFTva.BTVAEtabClick(Sender: TOBject);
begin
  ParamTVAEtab(CC_CODE.Text, FCodeTVA);
end;

procedure TFTva.STVADataChange(Sender: TObject; Field: TField);
var UpEnable, DnEnable: Boolean;
begin
  GeleLesBoutons;
  if Field = nil then
  begin
    UpEnable := Enabled and not TTva.BOF;
    DnEnable := Enabled and not TTva.EOF;
    BFirst.Enabled := UpEnable;
    BPrev.Enabled := UpEnable;
    BNext.Enabled := DnEnable;
    BLast.Enabled := DnEnable;
    ChargeEnreg;
  end else
  begin
    // code pour gerer les champ +- automatique
    if ((Field.FieldName = 'CC_LIBELLE') and (TTVACC_ABREGE.AsString = '')) then
      TTVACC_ABREGE.AsString := Copy(Field.AsString, 1, 17);
  end;
end;

procedure TFTva.ChargeEnreg;
begin
  InitCaption(Self, TTVACC_CODE.AsString, TTVACC_LIBELLE.AsString);
  CC_CODE.Enabled := False;
  if TTaux.Modified then
    if not Enregok1 then
    begin
      TTaux.Cancel;
      MsgBox.Execute(12, caption, '');
    end;
  InitRegime;
  Modifier1 := False;
end;

function TFTva.EnregOK: boolean;
begin
  result := FALSE;
  Modifier := True;
  if TTva.state in [dsEdit, dsInsert] = False then Exit;
  if TTva.state in [dsInsert] then
  begin
    if TTVACC_CODE.AsString = '' then
    begin
      MsgBox.Execute(2, caption, '');
      CC_CODE.SetFocus;
      Exit;
    end;
    if SiExisteCode then
    begin
      MsgBox.Execute(4, caption, '');
      CC_CODE.SetFocus;
      Exit;
    end;
  end;
  if TTva.state in [dsEdit, dsInsert] then
  begin
    if TTVACC_LIBELLE.AsString = '' then
    begin
      MsgBox.Execute(3, caption, '');
      CC_LIBELLE.SetFocus;
      Exit;
    end;
    if SiExisteIndice then
    begin
      MsgBox.Execute(16, Caption, '');
      CC_LIBRE.SetFocus;
      Exit;
    end;
  end;
  Result := TRUE;
  Modifier := False;
end;

function TFTva.OnSauve: boolean;
var Rep: Integer;
begin
  result := FALSE;
  Modifier := False;
  if TTva.Modified then
  begin
    if FAutoSave.Checked then Rep := mrYes else Rep := MsgBox.execute(0, caption, '');
  end else Rep := 321;
  case rep of
    mrYes: if not Bouge(nbPost) then exit;
    mrNo: if not Bouge(nbCancel) then exit;
    mrCancel: Exit;
  end;
  result := TRUE;
end;

function TFTva.Bouge(Button: TNavigateBtn): boolean;
begin
  result := FALSE;
  case Button of
    nblast, nbprior, nbnext,
      nbfirst, nbinsert: if not OnSauve then Exit;
    nbPost:
      begin
        if not MajCodeTauxSurInsert then Exit;
        if not EnregOk then Exit;
      end;
    nbDelete: if not Supprime then Exit;
  end;
  if (Button = nbCancel) and (TTva.State = dsInsert) then SupprimeTaux;
  if (Button = nbPost) and (not (TTva.state in [dsEdit, dsInsert])) then Exit;
  if not TransacNav(DBNav.BtnClick, Button, 10) then MessageAlerte(Msgbox.Mess[5]);
  Result := TRUE;
  if Button = NbInsert then NewEnreg;
end;

function TFTva.FabReqTva(Cpt: string): string;
begin
  {
  Result:='Select G_GENERAL from GENERAUX Where G_GENERAL="'+BourrelaDonc(Cpt,fbGene)+'"'
         +' AND G_NATUREGENE<>"COC" And G_NATUREGENE<>"COD" And G_NATUREGENE<>"COF" And G_NATUREGENE<>"COS" '
         +' AND G_COLLECTIF="-" AND G_NATUREGENE<>"IMO" AND G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"' ;
  }
  Result := 'Select G_GENERAL from GENERAUX Where G_GENERAL="' + Cpt + '"'
    + ' AND G_NATUREGENE<>"COC" And G_NATUREGENE<>"COD" And G_NATUREGENE<>"COF" And G_NATUREGENE<>"COS" '
    + ' AND G_COLLECTIF="-" AND G_NATUREGENE<>"IMO" AND G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"';
end;

function TFTva.EnregOK1: boolean;
var Q: TQuery;
  SQL: string;
begin
  Result := FALSE;
  Modifier1 := True;
  if IsFieldNull(TTaux, 'TV_TAUXACH') then if TTaux.State in [dsEdit, dsInsert] then TTauxTV_TAUXACH.AsFloat := 0;
  if IsFieldNull(TTaux, 'TV_TAUXVTE') then if TTaux.State in [dsEdit, dsInsert] then TTauxTV_TAUXVTE.AsFloat := 0;
  if (TTauxTV_TAUXACH.AsFloat <> 0) and (TTauxTV_CPTEACH.AsString = '') then
  begin
    MsgBox.Execute(10, caption, '');
    Fliste.SelectedIndex := 3;
    FListe.SetFocus;
    Exit;
  end;
  if (TTauxTV_TAUXVTE.AsFloat <> 0) and (TTauxTV_CPTEVTE.AsString = '') then
  begin
    MsgBox.Execute(11, caption, '');
    Fliste.SelectedIndex := 4;
    FListe.SetFocus;
    Exit;
  end;
  if TTauxTV_CPTEACH.AsString <> '' then
  begin
    SQL := FabReqTva(TTauxTV_CPTEACH.AsString);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      MsgBox.Execute(6, caption, '');
      Fliste.SelectedIndex := 3;
      Ferme(Q);
      Exit;
    end;
    Ferme(Q);
  end;
  if TTauxTV_CPTEVTE.AsString <> '' then
  begin
    SQL := FabReqTva(TTauxTV_CPTEVTE.AsString);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      MsgBox.Execute(7, caption, '');
      Fliste.SelectedIndex := 4;
      Ferme(Q);
      Exit;
    end;
    Ferme(Q);
  end;
  if TTauxTV_ENCAISACH.AsString <> '' then
  begin
    SQL := FabReqTva(TTauxTV_ENCAISACH.AsString);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      MsgBox.Execute(14, caption, '');
      Fliste.SelectedIndex := 5;
      Ferme(Q);
      Exit;
    end;
    Ferme(Q);
  end;
  if TTauxTV_ENCAISVTE.AsString <> '' then
  begin
    SQL := FabReqTva(TTauxTV_ENCAISVTE.AsString);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      MsgBox.Execute(15, caption, '');
      Fliste.SelectedIndex := 7;
      Ferme(Q);
      Exit;
    end;
    Ferme(Q);
  end;
  if TTauxTV_CPTVTERG.AsString <> '' then
  begin
    SQL := FabReqTva(TTauxTV_CPTVTERG.AsString);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      MsgBox.Execute(17, caption, '');
      Fliste.SelectedIndex := 6;
      Ferme(Q);
      Exit;
    end;
    Ferme(Q);
  end;

  Result := True;
  Modifier1 := False;
end;

(*
function TFTva.OnSauve1: boolean;
var Rep: Integer;
begin
  result := FALSE;
  if TTaux.Modified then
  begin
    if FAutoSave.Checked then Rep := mrYes else Rep := MsgBox.execute(0, caption, '');
  end else Rep := 321;
  case rep of
    mrYes: if not EnregOk1 then Exit;
    mrNo: TTaux.Cancel;
    mrCancel: Exit;
  end;
  result := TRUE;
end;
*)

procedure TFTva.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := OnSauve;
end;

procedure TFTva.NewEnreg;
begin
  InitNew(TTVA);
  TTVACC_TYPE.AsString := FCodeTVA;
  TTVACC_LIBRE.AsString := '0';
  UnNouveau := True;
  InitRegime;
  UnNouveau := False;
  CC_CODE.Enabled := True;
  CC_CODE.SetFocus;
end;

procedure ParamTVATPF(OkTva: boolean);
var FTVA: TFTVA;
  PP: THPanel;
begin
  if _Blocage(['nrCloture', 'nrBatch', 'nrSaisieModif', 'nrEnca', 'nrDeca'], True, 'nrBatch') then Exit;
  FTVA := TFTVA.Create(application);
  if OkTVA then
  begin
    FTVA.FCodeTVA := VH^.DefCatTVA;
    FTVA.HelpContext := 1170000;
  end else
  begin
    FTVA.FCodeTVA := VH^.DefCatTPF;
    FTVA.HelpContext := 1175000;
  end;
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      FTVA.ShowModal;
    finally
      FTVA.Free;
      _Bloqueur('nrBatch', False);
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(FTVA, PP);
    FTVA.Show;
  end;
end;

procedure TFTva.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ChargeTVATPF;
  if FAvertir then
  begin
    AvertirTable('ttRegimeTVA');
    AvertirTable('ttTvaEncaissement');
    AvertirTable('ttTva');
    AvertirTable('ttTpf');
    AvertirTable('ttControleTVA');
  end;
  LiRegime.Clear;
  LiRegime.Free;
  if Parent is THPanel then
  begin
    _Bloqueur('nrBatch', False);
    Action := caFree;
  end;
end;

procedure TFTva.BImprimerClick(Sender: TObject);
begin
  if (FCodeTVA = VH^.DefCatTVA) then PrintDBGrid(FListe, HB1, Caption, '')
  else PrintDBGrid(FListe, HB1, Caption, '');
end;

procedure TFTva.TTauxAfterDelete(DataSet: TDataSet);
begin
  FAvertir := True;
end;

procedure TFTva.TTauxAfterPost(DataSet: TDataSet);
begin
  FAvertir := True;
end;

procedure TFTva.TTVAAfterPost(DataSet: TDataSet);
begin
  FAvertir := True;
end;

procedure TFTva.TTVAAfterDelete(DataSet: TDataSet);
begin
  FAvertir := True;
end;

procedure TFTva.FListeKeyPress(Sender: TObject; var Key: Char);
var Nam: string;
begin
  Nam := FListe.SelectedField.FieldName;
  if (Nam = 'TV_TAUXACH') or (Nam = 'TV_TAUXVTE') then
  begin
    if (Key = 'E') or (Key = 'e') then
    begin
      Key := #0;
      Beep;
    end;
  end;
  if ((Nam = 'TV_CPTEACH') or (Nam = 'TV_CPTEVTE') or (Nam = 'TV_ENCAISACH') or (Nam = 'TV_ENCAISVTE')) then
  begin
    if Key in ['a'..'z'] then Key := upcase(Key);
  end;
end;

procedure TFTva.STVAUpdateData(Sender: TObject);
begin
  if Modifier then
  begin
    Modifier := False;
    if not OnSauve then Bouge(nbCancel);
  end;
end;

procedure TFTva.STVAStateChange(Sender: TObject);
begin
  Modifier := True;
end;

procedure TFTva.FListe2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (ssCtrl in Shift) and (Key = VK_DELETE) then
  begin
    Bouge(nbDelete);
    Key := 0;
  end;
end;

procedure TFTva.TTauxNewRecord(DataSet: TDataSet);
begin
  if not Insere then TTaux.Cancel;
end;

function TFTva.MajCodeTauxSurInsert: Boolean;
begin
  Result := False;
  if (TTaux.State = dsBrowse) and (TTVA.State <> dsInsert) then
  begin
    Result := True;
    Exit;
  end;
  if not EnregOk1 then Exit;
  if TTVA.State = dsInsert then
    if not EnregOK then Exit;
  if TTaux.State = dsEdit then TTaux.Post;
  if TTVA.State = dsInsert then
  begin
    ExecuteSql('Update TXCPTTVA Set TV_CODETAUX="' + TTVACC_CODE.AsString + '" Where TV_CODETAUX="' + W_W2 + '"');
    TTVA.Refresh;
  end;
  Modifier1 := False;
  Result := True;
end;

function TFTva.SiExisteCode: Boolean;
var Q: TQuery;
begin
  Q := OpenSql('Select CC_CODE From CHOIXCOD Where CC_TYPE="' + FCodeTVA + '" And ' +
    'CC_CODE="' + TTVACC_CODE.AsString + '"', True,-1,'',true);
  Result := not Q.Eof;
  Ferme(Q);
end;

function TFTva.SiExisteIndice: boolean;
var Q: TQuery;
begin
  Result := False;
  {#TVAENC}
  if ((FCodeTVA = VH^.DefCatTVA) and (CC_LIBRE.Text <> '') and (CC_LIBRE.Text <> '0')) then
  begin
    Q := OpenSQL('Select Count(*) from CHOIXCOD Where CC_TYPE="' + VH^.DefCatTVA + '" AND CC_CODE<>"' + CC_CODE.Text + '" AND CC_LIBRE="' + CC_LIBRE.Text +
      '"', True,-1,'',true);
    if not Q.EOF then if Q.Fields[0].AsInteger > 0 then Result := True;
    Ferme(Q);
  end;
end;

procedure TFTva.STauxStateChange(Sender: TObject);
begin
  Modifier1 := True;
end;

procedure TFTva.STauxDataChange(Sender: TObject; Field: TField);
begin
  GeleLesBoutons;
end;

procedure TFTva.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var i: Byte;
begin
  if not (Key in [VK_UP, VK_DOWN, VK_PRIOR, VK_NEXT, VK_TAB]) then Exit;
  if (Key in [VK_TAB]) and (Fliste.SelectedIndex <> 6) then Exit;
  i := FListe.SelectedIndex;
  FListe.SelectedIndex := i;
  if not TTaux.Modified then Exit;
  if not Enregok1 then
  begin
    Key := 0;
    Exit;
  end;
end;

procedure TFTva.STauxUpdateData(Sender: TObject);
begin
  if not Enregok1 then SysUtils.Abort;
end;

function TFTva.Supprime: Boolean;
var Q: TQuery;
begin
  Result := False;
  if MsgBox.Execute(1, caption, '') <> mrYes then Exit;
  if FCodeTva = VH^.DefCatTVA then Q := OpenSql('Select G_TVA From GENERAUX Where G_TVA="' + TTVACC_CODE.AsString + '"', True,-1,'',true) else
    if FCodeTva = VH^.DefCatTPF then Q := OpenSql('Select G_TPF From GENERAUX Where G_TPF="' + TTVACC_CODE.AsString + '"', True,-1,'',true) else
  begin
    Result := True;
    Exit;
  end;
  if Q.Eof then Result := True else MsgBox.Execute(13, caption, '');
  Ferme(Q);
end;

procedure TFTva.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFTva.FListeDblClick(Sender: TObject);
var
    St: string;
begin
  //if (TTaux.State in [dsEdit,dsInsert])=False then Exit ;
  if ((FListe.SelectedField.FieldName = 'TV_CPTEACH') or
    (FListe.SelectedField.FieldName = 'TV_CPTEVTE') or
    (FListe.SelectedField.FieldName = 'TV_ENCAISACH') or
    (FListe.SelectedField.FieldName = 'TV_CPTVTERG') or
    (FListe.SelectedField.FieldName = 'TV_ENCAISVTE')) then
  begin
    FListe.SelectedIndex := FListe.SelectedIndex;
    ChercheCpte(FListe.SelectedField, False);
    FListe.SelectedIndex := FListe.SelectedIndex;
    St := TTauxTV_CPTEACH.AsString;
    if St <> '' then if Length(St) < VH^.Cpta[fbGene].Lg then
      begin
        St := BourreLaDonc(St, fbGene);
        TTauxTV_CPTEACH.AsString := St;
      end;
    St := TTauxTV_CPTEVTE.AsString;
    if St <> '' then if Length(St) < VH^.Cpta[fbGene].Lg then
      begin
        St := BourreLaDonc(St, fbGene);
        TTauxTV_CPTEVTE.AsString := St;
      end;
    St := TTauxTV_ENCAISACH.AsString;
    if St <> '' then if Length(St) < VH^.Cpta[fbGene].Lg then
      begin
        St := BourreLaDonc(St, fbGene);
        TTauxTV_ENCAISACH.AsString := St;
      end;
    St := TTauxTV_ENCAISVTE.AsString;
    if St <> '' then if Length(St) < VH^.Cpta[fbGene].Lg then
      begin
        St := BourreLaDonc(St, fbGene);
        TTauxTV_ENCAISVTE.AsString := St;
      end;
    St := TTauxTV_CPTVTERG.AsString;
    if St <> '' then if Length(St) < VH^.Cpta[fbGene].Lg then
      begin
        St := BourreLaDonc(St, fbGene);
        TTauxTV_CPTVTERG.AsString := St;
      end;
  end;
end;

procedure TFTva.CC_LIBREExit(Sender: TObject);
var C: Char;
begin
  if CC_LIBRE.Text = '' then CC_LIBRE.Text := '0' else
    if Length(CC_LIBRE.Text) <> 1 then CC_LIBRE.Text := '0' else
  begin
    C := CC_LIBRE.Text[1];
    if (C in ['0'..'4']) = False then CC_LIBRE.Text := '0';
  end;
end;

procedure TFTva.FormCreate(Sender: TObject);
var Indice : integer;
    TheColonne : TColumn;
begin
  if V_PGI.Driver = dbMSACCESS then W_W2 := 'zzz' else W_W2 := W_W;
  // fiche qualitée N° 10834
  if (not GetParamSocSecur('SO_GCVENTILRG', False)) or
     (not GetParamSocSecur('SO_RETGARANTIE', False)) then
  begin
    // Supression dans la grid des champs associées
    for Indice := 0 To Fliste.Columns.Count -1 do
    begin
      TheColonne := Fliste.Columns[Indice];
      if (TheColonne.FieldName = 'TV_CPTVTERG') or
         (TheColonne.FieldName  = 'TV_CPTACHRG') then TheColonne.Visible := false;
    end;
  end;
  //uniquement en line
  {*
    for Indice := 0 To Fliste.Columns.Count -1 do
    begin
      TheColonne := Fliste.Columns[Indice];
      if (TheColonne.FieldName  = 'TV_CPTACHRG') then TheColonne.Visible := false;
      if (TheColonne.FieldName  = 'TV_TAUXACH') then TheColonne.Visible := false;
      if (TheColonne.FieldName  = 'TV_CPTEACH') then TheColonne.Visible := false;
      if (TheColonne.FieldName  = 'TV_ENCAISACH') then TheColonne.Visible := false;
    end;
  *}
end;

end.
