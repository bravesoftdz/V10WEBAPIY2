unit TVAEtab;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons, Dialogs,
  StdCtrls, ExtCtrls, Hctrls,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB, Grids, DBGrids, Mask, DBCtrls,
  Ent1, HDB, hmsgbox, Hcompte, sysutils, HEnt1, HSysMenu, Hqry, Spin,
  HPanel, UiUtil, HTB97, ToolWin, LookUp,
  {$IFNDEF CCADM}
  {$IFNDEF GCGC}
  {$IFNDEF COMPTAPAIE}
   CPGeneraux_TOM,
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
  {$IFNDEF EAGLCLIENT}
	PrintDBG,
  {$ENDIF}
  utom, ADODB,uEntCommun;

procedure ParamTVAEtab(CodeTaux, CodeTvaOuTpf: string);

type
  TFTvaEtab = class(TForm)
    TTauxCompl: THTable;
    STauxCompl: TDataSource;
    TEtab: THTable;
    STEtab: TDataSource;
    FListe2: THDBGrid;
    DBNav: TDBNavigator;
    FListe: THDBGrid;
    Panel1: TPanel;
    HB1: TPanel;
    TCC_CODE: THLabel;
    TCC_LIBELLE: THLabel;
    ET_ETABLISSEMENT: TDBEdit;
    ET_LIBELLE: TDBEdit;
    MsgBox: THMsgBox;
    Cache: THCpteEdit;
    TTauxComplTVC_TVAOUTPF: TStringField;
    TTauxComplTVC_CODETAUX: TStringField;
    TTauxComplTVC_REGIME: TStringField;
    TTauxComplTVC_ETABLISSEMENT: TStringField;
    TTauxComplTVC_CPTEACH: TStringField;
    TTauxComplTVC_CPTEVTE: TStringField;
    HMTrad: THSystemMenu;
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
    TEtabET_ETABLISSEMENT: TStringField;
    TEtabET_ABREGE: TStringField;
    TEtabET_LIBELLE: TStringField;
    TTVC_CODETAUX: THLabel;
    Taux: TEdit;
    HT: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure STEtabDataChange(Sender: TObject; Field: TField);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BImprimerClick(Sender: TObject);
    procedure TTauxComplAfterDelete(DataSet: TDataSet);
    procedure TTauxComplAfterPost(DataSet: TDataSet);
    procedure TEtabAfterPost(DataSet: TDataSet);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure STEtabStateChange(Sender: TObject);
    procedure FListe2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TTauxComplNewRecord(DataSet: TDataSet);
    procedure STauxComplStateChange(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure STauxComplDataChange(Sender: TObject; Field: TField);
    procedure STauxComplUpdateData(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure FListe2CellClick(Column: TColumn);
    procedure TTauxComplTVC_REGIMEChange(Sender: TField);

  private { Private declarations }
    CodeTaux, NvRegime, AncRegime, TvaOuTpf: string;
    FAvertir, Modifier, Modifier1, Init: Boolean;
    Insere, UnNouveau: Boolean;
    function ChercheCpte(T: TField; Vide: Boolean): byte;
    procedure ChargeEnreg;
    function EnregOK1: boolean;
    function OnSauve: boolean;
    function Bouge(Button: TNavigateBtn): boolean;
    function MajCodeTauxSurInsert: Boolean;
//    function SiExisteCode: Boolean;
//    function SiExisteIndice: boolean;
    procedure SupprimeTaux(StTVA, StEtab, StReg: string);
    procedure GeleLesBoutons;
    function Supprime: Boolean;
//    procedure CacheZonesAchat;
    function FabReqTva(Cpt: string): string;
    procedure InitRegime;
    procedure EnregCptOk;
  public { Public declarations }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  UtilPgi;

procedure TFTvaEtab.GeleLesBoutons;
begin
  BInsert.Enabled := (not (TEtab.State in [dsEdit, dsInsert])) and (not (TTauxCompl.State in [dsEdit, dsInsert]));
  BDelete.Enabled := (not (TEtab.State in [dsEdit, dsInsert])) and (not (TTauxCompl.State in [dsEdit, dsInsert]));
  if (TEtab.Eof) and (TEtab.Bof) then BDelete.Enabled := False;
end;

(*
procedure TFTvaEtab.CacheZonesAchat;
var i: integer;
begin
  for i := 0 to FListe.Columns.Count - 1 do
  begin
    if (FListe.Columns[i].FieldName = 'TVC_CPTEACH') then FListe.Columns[i].Visible := False;
  end;
end;
*)

procedure TFTvaEtab.FormShow(Sender: TObject);
var i: Byte;
  St: string;
begin
  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
  FAvertir := False;
  Insere := False;
  Modifier1 := False;
  TTauxCompl.Open;
  TEtab.Open;
  St := '';
  for i := 1 to VH^.Cpta[fbGene].Lg do St := St + 'a';
  if (TEtab.Eof) and (TEtab.Bof) then Bouge(nbInsert);
  Taux.Text := CodeTaux;
  if TvaOuTpf = VH^.DefCatTPF then
  begin
    FListe.Columns[4].Title.Caption := HT.Mess[0];
    FListe.Columns[5].Title.Caption := HT.Mess[1];
  end;
end;

function TFTvaEtab.ChercheCpte(T: TField; Vide: Boolean): byte;
var St: string;
begin
  ChercheCpte := 0;
  St := UpperCase(T.AsString);
  Cache.Text := St;
  if ((Vide) and (St = '')) then exit;
  {$IFNDEF CCADM}
  {$IFNDEF GCGC}
  {$IFNDEF COMPTAPAIE}
  if GChercheCompte(Cache, FicheGene) then
  {$ELSE}
  if GChercheCompte(Cache, nil) then
  {$ENDIF}
  {$ELSE}
  if GChercheCompte(Cache, nil) then
  {$ENDIF}
  {$ENDIF}
  begin
    if St <> Cache.Text then
    begin
      if not (TTauxCompl.State in [dsEdit]) then TTauxCompl.Edit;
      T.AsString := Cache.Text;
    end;
    ChercheCpte := 2;
  end;
end;

procedure TFTvaEtab.SupprimeTaux(StTVA, StEtab, StReg: string);
begin
  if TTauxCompl.State <> dsBrowse then TTauxCompl.Cancel;
  ExecuteSql('Delete From TXCPTTVACOMPL Where TVC_CODETAUX="' + StTVA + '" and TVC_REGIME="' + StReg + '" and TVC_ETABLISSEMENT="' + StEtab + '"');
end;

procedure TFTvaEtab.BFirstClick(Sender: TObject);
begin
  Bouge(nbFirst);
end;

procedure TFTvaEtab.BPrevClick(Sender: TObject);
begin
  Bouge(nbPrior);
end;

procedure TFTvaEtab.BNextClick(Sender: TObject);
begin
  Bouge(nbNext);
end;

procedure TFTvaEtab.BLastClick(Sender: TObject);
begin
  Bouge(nbLast);
end;

procedure TFTvaEtab.BInsertClick(Sender: TObject);
begin
  UnNouveau := True;
  InitRegime
end;

procedure TFTvaEtab.BDeleteClick(Sender: TObject);
var StQuoi, StTVA, StEtab, StReg: string;
begin
  if MsgBox.Execute(1, caption, '') <> mrYes then Exit;
  StQuoi := TTauxComplTVC_TVAOUTPF.AsString;
  ;
  StTVA := TTauxComplTVC_CODETAUX.AsString;
  StEtab := TTauxComplTVC_ETABLISSEMENT.AsString;
  StReg := TTauxComplTVC_REGIME.AsString;
  if FindLaKey(TTauxCompl, [StQuoi, StTVA, StEtab, StReg]) then
    SupprimeTaux(StTVA, StEtab, StReg);
  TTauxCompl.Close;
  TTauxCompl.Open;
  InitRegime;
end;

procedure TFTvaEtab.BValiderClick(Sender: TObject);
begin
  Modifier := False;
  Bouge(nbPost);
  Init := False;
end;

procedure TFTvaEtab.BAnnulerClick(Sender: TObject);
begin
  if TEtab.State <> dsInsert then TTauxCompl.Cancel;
  Bouge(nbCancel);
end;

procedure TFTvaEtab.BFermeClick(Sender: TObject);
begin
  EnregCptOk;
  Close;
end;

procedure TFTvaEtab.STEtabDataChange(Sender: TObject; Field: TField);
var UpEnable, DnEnable: Boolean;
begin
  GeleLesBoutons;
  EnregCptOk;
  if Field = nil then
  begin
    UpEnable := Enabled and not TEtab.BOF;
    DnEnable := Enabled and not TEtab.EOF;
    BFirst.Enabled := UpEnable;
    BPrev.Enabled := UpEnable;
    BNext.Enabled := DnEnable;
    BLast.Enabled := DnEnable;
    ChargeEnreg;
  end;
end;

procedure TFTvaEtab.ChargeEnreg;
begin
  if TTauxCompl.Modified then
    if not Enregok1 then
    begin
      TTauxCompl.Cancel;
      MsgBox.Execute(12, caption, '');
    end;
  InitRegime;
  Modifier1 := False;
end;

procedure TFTvaEtab.InitRegime;
var StQuoi, StEtab, StTVA, MemoStEtab, MemoStTVA: string[3];
begin
  //StQuoi:='TX1' ;
  StQuoi := TvaOuTpf;
  StEtab := TEtabET_ETABLISSEMENT.AsString;
  StTVA := CodeTaux;
  if (StTVA <> '') or (UnNouveau) then
  begin
    TTauxCompl.SetRange([StQuoi, StTVA, StEtab], [StQuoi, StTVA, StEtab]);
    MemoStEtab := StEtab;
    MemoStTVA := StTVA;
    if (not FindLaKey(TTauxCompl, [StQuoi, StTVA, StEtab])) or (UnNouveau) then
    begin
      Insere := True;
      Init := True;
      TTauxCompl.Insert;
      InitNew(TTauxCompl);
      TTauxComplTVC_TVAOUTPF.AsString := StQuoi;
      TTauxComplTVC_ETABLISSEMENT.AsString := StEtab;
      TTauxComplTVC_CODETAUX.AsString := StTVA;
      //TTauxComplTVC_REGIME.AsString:=StREG ;
      //TTauxCompl.Post ;
      UnNouveau := False;
    end else Insere := False;
  end;
  //TTauxCompl.First ; Insere:=False ;
  Init := False;
end;

function TFTvaEtab.OnSauve: boolean;
var Rep: Integer;
begin
  result := FALSE;
  Modifier := False;
  if (TTauxComplTVC_REGIME.AsString = '') and (TTauxCompl.State = DsInsert) then TTauxCompl.Cancel;
  if TEtab.Modified then
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

function TFTvaEtab.Bouge(Button: TNavigateBtn): boolean;
begin
  result := FALSE;
  case Button of
    nblast, nbprior, nbnext,
      nbfirst, nbinsert: if not OnSauve then Exit;
    nbPost: if not MajCodeTauxSurInsert then Exit;
    nbDelete: if not Supprime then Exit;
  end;
  //if(Button=nbCancel) And (TEtab.State=dsInsert)then SupprimeTaux ;
  if (Button = nbPost) and (not (TEtab.state in [dsEdit, dsInsert])) then TEtab.Edit;
  if not TransacNav(DBNav.BtnClick, Button, 10) then MessageAlerte(Msgbox.Mess[5]);
  Result := TRUE;
end;

function TFTvaEtab.FabReqTva(Cpt: string): string;
begin
  Result := 'Select G_GENERAL from GENERAUX Where G_GENERAL="' + Cpt + '"'
    + ' AND G_NATUREGENE<>"COC" And G_NATUREGENE<>"COD" And G_NATUREGENE<>"COF" And G_NATUREGENE<>"COS" '
    + ' AND G_COLLECTIF="-" AND G_NATUREGENE<>"IMO" AND G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"';
end;

function TFTvaEtab.EnregOK1: boolean;
var Q: TQuery;
  SQL: string;
begin
  Result := FALSE;
  Modifier1 := True;
  if Init then
  begin
    result := True;
    Exit
  end;
  if (TTauxComplTVC_REGIME.AsString = '') and ((TTauxComplTVC_CPTEACH.AsString <> '') or (TTauxComplTVC_CPTEACH.AsString <> '')) then
  begin
    MsgBox.Execute(8, caption, '');
    Fliste.SelectedIndex := 3;
    Exit;
  end;
  if TTauxComplTVC_CPTEACH.AsString <> '' then
  begin
    SQL := FabReqTva(TTauxComplTVC_CPTEACH.AsString);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      MsgBox.Execute(6, caption, '');
      Fliste.SelectedIndex := 4;
      Ferme(Q);
      Exit;
    end;
    Ferme(Q);
  end;
  if TTauxComplTVC_CPTEVTE.AsString <> '' then
  begin
    SQL := FabReqTva(TTauxComplTVC_CPTEVTE.AsString);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      MsgBox.Execute(7, caption, '');
      Fliste.SelectedIndex := 5;
      Ferme(Q);
      Exit;
    end;
    Ferme(Q);
  end;
  Result := True;
  Modifier1 := False;
end;

procedure TFTvaEtab.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := OnSauve;
end;

procedure ParamTVAEtab(CodeTaux, CodeTvaOuTpf: string);
var FTvaEtab: TFTvaEtab;
  PP: THPanel;
begin
  if _Blocage(['nrCloture', 'nrBatch', 'nrSaisieModif', 'nrEnca', 'nrDeca'], True, 'nrBatch') then Exit;
  FTvaEtab := TFTvaEtab.Create(application);
  FTvaEtab.CodeTaux := CodeTaux;
  FTvaEtab.TvaOuTpf := CodeTvaOuTpf;
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      FTvaEtab.ShowModal;
    finally
      FTvaEtab.Free;
      _Bloqueur('nrBatch', False);
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(FTvaEtab, PP);
    FTvaEtab.Show;
  end;
end;

procedure TFTvaEtab.FormClose(Sender: TObject; var Action: TCloseAction);
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
  if Parent is THPanel then
  begin
    _Bloqueur('nrBatch', False);
    Action := caFree;
  end;
end;

procedure TFTvaEtab.BImprimerClick(Sender: TObject);
begin
  //if (FCodeTVA=VH^.DefCatTVA) then PrintDBGrid (FListe,HB1,Caption,'')
                              //else
  PrintDBGrid(FListe, HB1, Caption, '');
end;

procedure TFTvaEtab.TTauxComplAfterDelete(DataSet: TDataSet);
begin
  FAvertir := True;
end;

procedure TFTvaEtab.TTauxComplAfterPost(DataSet: TDataSet);
begin
  FAvertir := True;
end;

procedure TFTvaEtab.TEtabAfterPost(DataSet: TDataSet);
begin
  FAvertir := True;
end;

procedure TFTvaEtab.FListeKeyPress(Sender: TObject; var Key: Char);
var Nam: string;
begin
  Nam := FListe.SelectedField.FieldName;
  if ((Nam = 'TVC_CPTEACH') or (Nam = 'TVC_CPTEVTE')) then
  begin
    if Key in ['a'..'z'] then Key := upcase(Key);
  end;
end;

procedure TFTvaEtab.STEtabStateChange(Sender: TObject);
begin
  Modifier := True;
end;

procedure TFTvaEtab.FListe2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (TTauxComplTVC_REGIME.AsString = '') and (TTauxCompl.State = DsInsert) then TTauxCompl.Cancel;
  //if (ssCtrl in Shift) And (Key=VK_DELETE) then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFTvaEtab.FListe2CellClick(Column: TColumn);
begin
  if (TTauxCompl.State in [dsEdit, dsInsert]) then TTauxCompl.Cancel;
end;

procedure TFTvaEtab.TTauxComplNewRecord(DataSet: TDataSet);
begin
  if not Insere then TTauxCompl.Cancel;
end;

function TFTvaEtab.MajCodeTauxSurInsert: Boolean;
var regime: string;
begin
  Result := False;
  if (TTauxCompl.State = dsBrowse) then
  begin
    Result := True;
    Exit;
  end; //And (TEtab.State<>dsInsert) then BEGIN Result:=True ; Exit ; END ;
  if not EnregOk1 then Exit;
  //TTauxComplTVC_TVAOUTPF.AsString:='TX1';
  TTauxComplTVC_TVAOUTPF.AsString := TvaOuTpf;
  TTauxComplTVC_ETABLISSEMENT.AsString := ET_ETABLISSEMENT.Text;
  TTauxComplTVC_CODETAUX.AsString := CodeTaux;
  Regime := TTauxComplTVC_REGIME.AsString;
  if TTauxCompl.State = dsEdit then TTauxCompl.Post;
  Modifier1 := False;
  Result := True;
end;

(*
function TFTvaEtab.SiExisteCode: Boolean;
begin
  Result := False;
end;
*)

(*
function TFTvaEtab.SiExisteIndice: boolean;
begin
  Result := False;
end;
*)
procedure TFTvaEtab.STauxComplStateChange(Sender: TObject);
begin
  Modifier1 := True;
end;

procedure TFTvaEtab.STauxComplDataChange(Sender: TObject; Field: TField);
begin
  GeleLesBoutons;
end;

procedure TFTvaEtab.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var i: Byte;
begin
  if not (Key in [VK_UP, VK_DOWN, VK_PRIOR, VK_NEXT, VK_TAB]) then Exit;
  if (Key in [VK_TAB]) and (Fliste.SelectedIndex <> 6) then Exit;
  i := FListe.SelectedIndex;
  FListe.SelectedIndex := i;
  if not TTauxCompl.Modified then Exit;
  if not Enregok1 then
  begin
    Key := 0;
    Exit;
  end;
end;

procedure TFTvaEtab.STauxComplUpdateData(Sender: TObject);
begin
  if not Enregok1 then SysUtils.Abort;
end;

function TFTvaEtab.Supprime: Boolean;
begin
  Result := False;
  if MsgBox.Execute(1, caption, '') <> mrYes then Exit;
  Result := True;
end;

procedure TFTvaEtab.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFTvaEtab.FListeDblClick(Sender: TObject);
var
  St: string;
  REGIME: THCritMaskEdit;
begin
  if not (TTauxCompl.State in [dsEdit, dsInsert]) then TTauxCompl.Edit;
  if ((FListe.SelectedField.FieldName = 'TVC_CPTEACH') or
    (FListe.SelectedField.FieldName = 'TVC_CPTEVTE')) then
  begin
    FListe.SelectedIndex := FListe.SelectedIndex;
    ChercheCpte(FListe.SelectedField, False);
    FListe.SelectedIndex := FListe.SelectedIndex;
    St := TTauxComplTVC_CPTEACH.AsString;
    if St <> '' then if Length(St) < VH^.Cpta[fbGene].Lg then
      begin
        St := BourreLaDonc(St, fbGene);
        TTauxComplTVC_CPTEACH.AsString := St;
      end;
    St := TTauxComplTVC_CPTEVTE.AsString;
    if St <> '' then if Length(St) < VH^.Cpta[fbGene].Lg then
      begin
        St := BourreLaDonc(St, fbGene);
        TTauxComplTVC_CPTEVTE.AsString := St;
      end;
  end
  else
  begin
    AncRegime := TTauxComplTVC_REGIME.AsString;
    REGIME := THCritMaskEdit.Create(FListe);
    REGIME.Parent := FListe;
    REGIME.Width := 3;
    REGIME.Visible := False;
    REGIME.OpeType := otString;
    REGIME.DATATYPE := 'TTREGIMETVA';
    LookUpCombo(REGIME);
    if REGIME.Text <> '' then
    begin
      NvRegime := REGIME.Text;
      TTauxComplTVC_REGIME.AsString := REGIME.Text;
    end;
    REGIME.Destroy;
  end;
end;

procedure TFTvaEtab.FormCreate(Sender: TObject);
begin
  //if V_PGI.Driver=dbMSACCESS then W_W2:='zzz' else W_W2:=W_W ;
end;

procedure TFTvaEtab.FListeEnter(Sender: TObject);
begin
  Insere := True;
  Modifier1 := True;
end;

procedure TFTvaEtab.EnregCptOk;
var SQL: string;
  Q: TQuery;
begin
  SQL := 'SELECT * FROM TXCPTTVACOMPL WHERE TVC_TVAOUTPF="' + TTauxComplTVC_TVAOUTPF.AsString + '" and  TVC_CODETAUX="' + TTauxComplTVC_CODETAUX.AsString + '"'
    +
    ' AND TVC_ETABLISSEMENT="' + TTauxComplTVC_ETABLISSEMENT.AsString + '"';
  Q := OpenSQL(SQL, True,-1,'',true);
  while not Q.EOF do
  begin
    if (Q.FindField('TVC_CPTEACH').AsString = '') and (Q.FindField('TVC_CPTEVTE').AsString = '') then
    begin
      SQL := 'DELETE FROM TXCPTTVACOMPL WHERE TVC_TVAOUTPF="' + TTauxComplTVC_TVAOUTPF.AsString + '" and  TVC_CODETAUX="' + TTauxComplTVC_CODETAUX.AsString +
        '"' +
        ' AND TVC_REGIME="' + Q.FindField('TVC_REGIME').AsString + '" AND TVC_ETABLISSEMENT="' + TTauxComplTVC_ETABLISSEMENT.AsString + '"';
      EXECUTESQL(SQL);
    end;
    Q.Next;
  end;
  ferme(Q);
end;

procedure TFTvaEtab.TTauxComplTVC_REGIMEChange(Sender: TField);
begin
  if NvRegime <> AncRegime then
  begin
    if ExisteSQL('Select TVC_TVAOUTPF from TXCPTTVACompl where TVC_TVAOUTPF="' + TTauxComplTVC_TVAOUTPF.AsString + '" and TVC_CODETAUX="' +
      TTauxComplTVC_CODETAUX.AsString + '" and '
      + 'TVC_ETABLISSEMENT="' + TTauxComplTVC_ETABLISSEMENT.AsString + '" and TVC_REGIME="' + NvRegime + '"') then
    begin
      MsgBox.Execute(17, caption, '');
      Fliste.SelectedIndex := 3;
      NvRegime := AncRegime;
      TTauxComplTVC_REGIME.AsString := NvRegime;
      TTauxCompl.Post;
    end;
  end;
end;

end.
