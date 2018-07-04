unit UTofPiedPort;

interface
uses StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAgl,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} grids, FE_main,
  {$ENDIF}
  forms, sysutils,
  ComCtrls, Hpanel, Math,
  HCtrls, HEnt1, HMsgBox, UTOF, Vierge, UTOB, AglInit, LookUp, SaisUtil, EntGC, graphics,
  windows, M3FP, HTB97, Dialogs, AGLInitGC, ExtCtrls, Hqry, Factutil,
  UtilPgi, Ent1;

function AppelPiedPort(Action: TActionFiche; TOBPieces, TOBPorcs, TOBBasesP: TOB; TarifTTC: Boolean = False): boolean;

type
  TOF_PiedPort = class(TOF)
  private
    Format: string;
    EnHT: Boolean;
    G_Port: THGRID;
    PCUMUL, PPIED: THPanel;
    CUMMONTANTHT: THNumEdit;
    CUMMONTANTTC: THNumEdit;
    GPT_FRANCO: TRadioButton;
    GPT_MINIMUM: THNumEdit;
    TGPT_MINIMUM: THLabel;
    GPT_MONTANTMINI: THNumEdit;
    TGPT_MONTANTMINI: THLabel;
    TTOTALHTMAR: THLabel;
    TOTALHTMAR: THNumEdit;
    TOTALTTCMAR: THNumEdit;
    TOTALHT: double;
    TOTALTTC: double;
    BZOOM: TToolbarButton97;
    BNEWLINE: TToolbarButton97;
    BDELLINE: TToolbarButton97;
    BAFFECT: TToolbarButton97;
    BVERROU: TToolbarButton97;
    TOBPort: TOB;
    //
    fListe,fColonnes : string;
    FNomTable, FLien, FSortBy, stCols : string;
    Ftitre : Hstring;
    FLargeur, FAlignement, FParams, FPerso: string;
    title, NC : Hstring;
    OkTri, OkNumCol : boolean;
    nbcolsinliste : integer;
    StCellCur : string;
    DEV : Rdevise;
    // Actions liées au Grid
    procedure EtudieColsListe;
    procedure ChargeGrille;
    function GetTOBPLigne(ARow: integer): TOB;
    procedure InitRow(Row: integer);
    procedure CreerTOBPLigne(ARow: integer);
    function LigVide(Row: integer): Boolean;
    procedure InsertLigne(ARow: Integer);
    procedure SupprimeLigne(ARow: Integer);
    procedure AfficheLaligne(ARow: Integer);
    procedure ChercherPort;
    procedure InitLaLigne(ARow: integer; TOBP, TOBPL: TOB);
    procedure ControlColGrille(ARow: Integer);
    function SortDeLaLigne: boolean;
    // Evenement du Grid
    procedure G_PortElipsisClick(Sender: TObject);
    procedure G_PortCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_PortCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_PortRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_PortRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_PortColumnWidthsChanged(Sender: TObject);
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    procedure DessineCellPort(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    {$ENDIF}
    // Manipulation des Champs GRID
    procedure TraiterPort(ACol, ARow: integer; var Cancel: Boolean);
    procedure TraiterLibelle(ACol, ARow: integer; var Cancel: Boolean);
    procedure TraiterBase(ACol, ARow: integer; var Cancel: Boolean);
    procedure TraiterPourcent(ACol, ARow: integer; var Cancel: Boolean);
    procedure TraiterMontant(ACol, ARow: integer; var Cancel: Boolean);
    procedure CalculMontant(TOBPL: TOB; ARow: integer);
    procedure CalculTotal;
    // Evenement Pied
    procedure GPT_FRANCOClick(Sender: TObject);
    procedure GPT_MINIMUMExit(Sender: TObject);
    procedure GPT_MONTANTMINIExit(Sender: TObject);
    // Manipulation du pied
    procedure AffichePied(ARow: integer);
    // Evenement lié aux Boutons
    procedure BZOOMClick(Sender: TObject);
    procedure BNEWLINEClick(Sender: TObject);
    procedure BDELLINEClick(Sender: TObject);
    procedure BAFFECTClick(Sender: TObject);
    { MenuItems }
    procedure ChangeBloqueFrais(Valeur: string; ImgIndex: Integer);
    procedure PMVERROU_OnPopup(Sender: TObject);
    procedure MnVer_OnClick(Sender: TObject);
    procedure MnVem_OnClick(Sender: TObject);
    procedure MnNon_OnClick(Sender: TObject);
    procedure ZoneSuivanteOuOk(Grille: THGrid; var ACol, ARow: integer;
      var Cancel: boolean);
    function ZoneAccessible(Grille: THgrid; var ACol,
      ARow: integer): boolean;
    procedure SetparamGrille;
    procedure TraiterMontantTTC(ACol, ARow: integer; var Cancel: Boolean);
    procedure AddlesChampsSupport(TT: TOB);
    procedure TraiterBaseTTC(ACol, ARow: integer; var Cancel: Boolean);
    procedure Raye(Grid: THgrid; Canvas: TCanvas; ACol, ARow: integer);
    procedure TraiteLaLigne(Acol, Arow: integer);
    procedure CalculMontants;
    procedure TraiterMini(ACol, ARow: integer; var Cancel: Boolean);
    procedure TraiterTVA(ACol, ARow: integer; var Cancel: Boolean);
  public
    Action: TActionFiche;
    // Evenement de la Form
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    // Evenement de la TOF
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;

  end;

var TOBPiece, TOBPied,TOBBases: TOB;
//  DEV: RDEVISE;

const colRang = 1;
  NbRowsInit = 50;
  NbRowsPlus = 20;

const PT_Port: integer = 0;
  PT_Libel: integer = 0;
  PT_BaseHT: integer = 0;
  PT_Mini: integer = 0;
  PT_BaseTTC: integer = 0;
  PT_Pour: integer = 0;
  PT_MontHT: integer = 0;
  PT_MontTTC: integer = 0;
  PT_Bloque: integer = 0;
  PT_TVA: integer = 0;
  IcoVER = 68;
  IcoVEM = 22; //A VOIR
  IcoNON = -1;
  // libellés des messages
  TexteMsg: array[1..1] of string = (
    {1}'Le code blocage n''existe pas. Vérifiez dans la liste.'
    );

procedure ReaffecterLesPorcs(TOBPorcs: TOB);

implementation

uses
	BTPUtil,
  Menus,
  FactRetenues;

function AppelPiedPort(Action: TActionFiche; TOBPieces, TOBPorcs,TOBBasesP: TOB; TarifTTC: Boolean = False): boolean;
var Arg, Retour: string;
		DEV : Rdevise;
begin
  Result := False;
  if (TOBPieces = nil) or (TOBPorcs = nil) then exit;
  TOBPiece := TOBPieces;
  TOBPied := TOB.Create('', nil, -1);
  TOBPied.Dupliquer(TOBPorcs, True, True);
  TOBBases := TOBBasesP;
  DEV.Code := TOBPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.GetValue('GP_TAUXDEV');
  DEV.DateTaux := TOBPiece.GetValue('GP_DATETAUXDEV');
  Arg := ActionToString(Action);
  Retour := AglLanceFiche('GC', 'GCPIEDPORT', '', '', Arg);
  if Retour = 'VALIDE' then
  begin
    if Action <> taconsult then
    begin
      ReaffecterLesPorcs(TOBPorcs);
      Result := True;
    end;
  end;
  TOBPied.Free;
end;

procedure ReaffecterLesPorcs(TOBPorcs: TOB);
var i_ind: integer;
begin
  if TOBPied = nil then exit;
  for i_ind := TOBPorcs.detail.count - 1 downto 0 do TOBPorcs.detail[i_ind].Free;
  TOBPorcs.Dupliquer(TOBPied, True, True);
  for i_ind := 0 to TOBPorcs.detail.count - 1 do TOBPorcs.detail[i_ind].PutValue('GPT_NUMPORT', i_ind + 1);
end;

{==============================================================================================}
{================================= Evenement de la TOF ========================================}
{==============================================================================================}

procedure TOF_PiedPort.OnArgument(Arguments: string);
var St: string;
  i: integer;
begin
  inherited;
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.GetValue('GP_TAUXDEV');
  DEV.DateTaux := TOBPiece.GetValue('GP_DATETAUXDEV');
  //
  St := Arguments;
  i := Pos('ACTION=', St);
  if i > 0 then
  begin
    System.Delete(St, 1, i + 6);
    St := uppercase(ReadTokenSt(St));
    if St = 'CREATION' then
    begin
      Action := taCreat;
    end;
    if St = 'MODIFICATION' then
    begin
      Action := taModif;
    end;
    if St = 'CONSULTATION' then
    begin
      Action := taConsult;
    end;
  end;
  TOBPort := TOB.Create('', nil, -1);
  Format := '#,##0';
  case DEV.Decimale of
    0: ;
    1: Format := Format + '.0';
    2: Format := Format + '.00';
    3: Format := Format + '.000';
    4: Format := Format + '.0000';
    5: Format := Format + '.00000';
  end;
  EnHT := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');

  BZOOM := TToolbarButton97(GetControl('BZOOM'));
  BZOOM.OnClick := BZOOMCLick;
  BNEWLINE := TToolbarButton97(GetControl('BNEWLINE'));
  BNEWLINE.OnClick := BNEWLINECLick;
  BDELLINE := TToolbarButton97(GetControl('BDELLINE'));
  BDELLINE.OnClick := BDELLINECLick;
  BAFFECT := TToolbarButton97(GetControl('BAFFECT'));
  BAFFECT.OnClick := BAFFECTCLick;
  BVERROU := TToolbarButton97(GetControl('BVERROU'));

  CUMMONTANTHT := THNumEdit(GetControl('CUMMONTANTHT'));
  CUMMONTANTTC := THNumEdit(GetControl('CUMMONTANTTTC'));
  CUMMONTANTHT.Masks.PositiveMask := Format;
  CUMMONTANTTC.Masks.PositiveMask := Format;
  PCUMUL := THPanel(GetControl('PCUMUL'));
  PPIED := THPanel(GetControl('PPIED'));
  GPT_FRANCO := TRadioButton(GetControl('GPT_FRANCO'));
  GPT_FRANCO.OnClick := GPT_FRANCOClick;
  GPT_MINIMUM := THNumEdit(GetControl('GPT_MINIMUM'));
  GPT_MINIMUM.OnExit := GPT_MINIMUMExit;
  GPT_MINIMUM.Masks.PositiveMask := Format;
  TGPT_MINIMUM := THLabel(GetControl('TGPT_MINIMUM'));
  GPT_MONTANTMINI := THNumEdit(GetControl('GPT_MONTANTMINI'));
  GPT_MONTANTMINI.OnExit := GPT_MONTANTMINIExit;
  GPT_MONTANTMINI.Masks.PositiveMask := Format;
  TGPT_MONTANTMINI := THLabel(GetControl('TGPT_MONTANTMINI'));
  TOTALHTMAR := THNumEdit(GetControl('TOTALHTMAR'));
  TOTALTTCMAR := THNumEdit(GetControl('TOTALTTCMAR'));
  TOTALHTMAR.Masks.PositiveMask := Format;
  TOTALTTCMAR.Masks.PositiveMask := Format;
  TotalHT := TOBPiece.Somme('GL_TOTALHTDEV', ['GL_TYPELIGNE'], ['ART'], False) + TOBPiece.GetValue('GP_TOTALESCDEV');
  TotalTTC := TOBPiece.Somme('GL_TOTALTTCDEV', ['GL_TYPELIGNE'], ['ART'], False) + TOBPiece.GetValue('GP_TOTESCTTCDEV');
  TOTALHTMAR.Value := TOBPiece.Somme('GL_TOTALHTDEV', ['GL_TYPELIGNE'], ['ART'], False) + TOBPiece.GetValue('GP_TOTALESCDEV');
  TOTALTTCMAR.Value := TOBPiece.Somme('GL_TOTALTTCDEV', ['GL_TYPELIGNE'], ['ART'], False) + TOBPiece.GetValue('GP_TOTESCTTCDEV');

  TOTALHTMAR.Enabled := False;
  TOTALTTCMAR.Enabled := False;

  G_Port := THGRID(GetControl('G_PORT'));
//  G_Port.Tag:=1 ; // Pour ListeParam
//  G_Port.ListeParam := 'GCPIEDPORT';
	EtudieColsListe;
  G_Port.OnElipsisClick := G_PortElipsisClick;
  G_Port.OnDblClick := G_PortElipsisClick;
  G_Port.OnCellEnter := G_PortCellEnter;
  G_Port.OnCellExit := G_PortCellExit;
  G_Port.OnRowEnter := G_PortRowEnter;
  G_Port.OnRowExit := G_PortRowExit;
  G_Port.OnColumnWidthsChanged := G_PortColumnWidthsChanged;
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  G_Port.PostDrawCell := DessineCellPort;
  {$ENDIF}
  G_Port.RowCount := NbRowsInit;
  G_Port.ColWidths[0] := 18;
  AffecteGrid(G_Port, Action);
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_Port);
  TFVierge(Ecran).OnKeyDown := FormKeyDown;
  TFVierge(Ecran).Retour := '';
  if Assigned(GetControl('PMVERROU')) then
    TPopupMenu(GetControl('PMVERROU')).OnPopup := PMVERROU_OnPopup;
  if Assigned(GetControl('MnVER')) then
    TMenuItem(GetControl('MnVER')).OnClick := MnVER_OnClick;
  if Assigned(GetControl('MnVEM')) then
    TMenuItem(GetControl('MnVEM')).OnClick := MnVEM_OnClick;
  if Assigned(GetControl('MnNON')) then
    TMenuItem(GetControl('MnNON')).OnClick := MnNON_OnClick;

end;

procedure TOF_PiedPort.OnLoad;
begin
  inherited;
  ChargeGrille;
end;

procedure TOF_PiedPort.OnUpdate;
var i_ind: integer;
  verrou : string;
begin
  inherited;
  if not SortDeLaLigne then
  begin
    TFVierge(Ecran).ModalResult := 0;
    ;
    Exit;
  end;
  for i_ind := TOBPied.detail.count - 1 downto 0 do
  begin
    if TOBPied.detail[i_ind].GetValue('GPT_CODEPORT') = '' then
    begin
      TOBPied.detail[i_ind].Free;
      continue;
    end;
    verrou := TOBPied.detail[i_ind].getString('GPT_BLOQUEFRAIS');
    if verrou = 'VEM' then begin TOBPied.detail[i_ind].SetString('GPT_BLOQUEFRAIS','VER') end;
  end;
  TFVierge(Ecran).Retour := 'VALIDE';
end;

procedure TOF_PiedPort.OnClose;
begin
  G_Port.VidePile(False);
  TOBPort.free;
  inherited;
end;

{==============================================================================================}
{================================ Evènements de la Form =======================================}
{==============================================================================================}

procedure TOF_PiedPort.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid: Boolean;
  ARow: Longint;
begin
  if Action = taConsult then Exit;
  FocusGrid := False;
  ARow := 0;
  if (Screen.ActiveControl = G_Port) then
  begin
    FocusGrid := True;
    ARow := G_Port.Row;
  end;
  case Key of
    VK_F5: if FocusGrid then G_Port.OnElipsisClick(Sender);
    VK_RETURN: Key := VK_TAB;
    VK_INSERT:
      begin
        if FocusGrid then
        begin
          Key := 0;
          InsertLigne(ARow);
        end;
      end;
    VK_DELETE:
      begin
        if ((FocusGrid) and (Shift = [ssCtrl])) then
        begin
          Key := 0;
          SupprimeLigne(ARow);
        end;
      end;
  end;
end;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
procedure TOF_PiedPort.SetparamGrille;
var lelement,laliste : string;
begin
  // récupération du paramétrage général des grilles
  fliste := 'GCPIEDPORT';
  ChargeHListe (fListe, FNomTable, FLien, FSortBy, stCols, FTitre,
                FLargeur, FAlignement, FParams, title, NC, FPerso, OkTri, OkNumCol);
  fColonnes  := 'SEL;'+stCols;

  laliste := fColonnes;
  nbcolsInListe := 0;
  repeat
    lelement := READTOKENST (laliste);
    if lelement <> '' then
    begin
      inc(nbcolsInListe);
    end;
  until lelement = '';

end;

procedure TOF_PiedPort.EtudieColsListe;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,TheLibelle : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec : integer;
    indice : integer;
    FFQTE,FFLQTE : string;
begin
  SetparamGrille;
	FFQTE := '###';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQTE := FFQTE+'0.';
    for indice := 1 to V_PGI.OkDecQ do
    begin
      FFQTE := FFQTE + '0';
    end;
  end;
  st := fColonnes;
  (*
  if not EnHT then
  begin
    st := FindEtReplace(st, 'GPT_BASEHTDEV', 'GPT_BASETTCDEV', False);
    st := FindEtReplace(st, 'GPT_TOTALHTDEV', 'GPT_TOTALTTCDEV', False);
  end;
  *)
  G_Port.ColCount := NbColsInListe;

  lesalignements := Falignement;
  lestitres := FTitre;
  leslargeurs := flargeur;

  for indice := 0 to nbcolsInListe -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    if Nam = 'SEL' then
    begin
      G_Port.ColNames [Indice] := Nam;
      G_Port.Cells[Indice,0]:='';
      G_Port.ColWidths[indice] := 15;
    end else
    begin
      alignement := ReadTokenSt(lesalignements);
      lalargeur := readtokenst(leslargeurs);
      letitre := readtokenst(lestitres);
      TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
      Thelibelle := Nam;
      G_Port.cells[Indice,0] := leTitre;
      G_Port.ColNames [Indice] := Nam;
      if copy(Alignement,1,1)='G' then G_Port.ColAligns[indice] := taLeftJustify
      else if copy(Alignement,1,1)='D' then G_Port.ColAligns[indice] := taRightJustify
      else if copy(Alignement,1,1)='C' then G_Port.ColAligns[indice] := taCenter;
      okImg := (copy(Alignement,8,1)='X');
      if OkVisu then
      begin
        G_Port.ColWidths[indice] := strtoint(lalargeur);
      end else
      begin
        G_Port.ColWidths[indice] := -1;
      end;

      if (OkLib) or (okImg) then
      begin
        G_Port.ColFormats[indice] := 'CB=' + Get_Join(Nam);
        if OkImg then G_Port.ColDrawingModes[Indice]:= 'IMAGE';
      end
      else if (Dec<>0) or (Sep) then
      begin
        if OkNulle then
        begin
          FFLQte := FFQTE+';-'+FFQTE+'; ;';
          G_Port.ColFormats[indice] := FFLQTE ;
        end else
        begin
          G_Port.ColFormats[indice] := FF ;
        end;
      end;

      if Nam = 'GPT_CODEPORT' then
      begin
        PT_Port := Indice;
      end else if Nam = 'GPT_LIBELLE' then
      begin
        PT_Libel := Indice;
      end else if (Nam = 'GPT_BASEHTDEV')  then
      begin
        PT_BaseHT := Indice;
      end else if (Nam = 'GPT_BASETTCDEV') then
      begin
        PT_BaseTTC := Indice;
      end else if Nam = 'GPT_POURCENT' then
      begin
        PT_Pour := Indice
      end else if Nam = 'GPT_MONTANTMINI' then
      begin
        PT_Mini := Indice
      end else if (Nam = 'GPT_TOTALHTDEV')  then
      begin
        PT_MontHT := Indice;
      end else if (Nam = 'GPT_TOTALTTCDEV') then
      begin
        PT_MontTTC := Indice;
      end else if (Nam = 'GPT_VALEURHTDEV') or (Nam = 'GPT_TIERSFRAIS') then
      begin
        G_Port.ColEditables[Indice] := False;
        G_Port.ColLengths[Indice] := -1;
        G_Port.ColWidths[indice] := -1;
      end else if (Nam = 'GPT_BLOQUEFRAIS') then
      begin
        G_Port.ColLengths[indice] := -1;
      end else if (Nam = 'GPT_FAMILLETAXE1') then
      begin
        PT_TVA := indice;
        G_Port.ColFormats [Indice] := 'CB=GCFAMILLETAXE1';
      end;

    end;
  end ;
end;

procedure TOF_PiedPort.TraiteLaLigne ( Acol,Arow : integer);
begin
  TOBPied.PutLigneGrid(G_Port,Arow,False,false,fColonnes);
  CalculMontants;
  ControlColGrille (Arow);
  AffichePied (Arow);
  CalculTotal;
end;

procedure TOF_PiedPort.CalculMontants;
var i : Integer;
begin
  for i := 0 to TOBPied.Detail.Count - 1 do
  begin
    CalculMontant(TOBPied.Detail[i], i + 1);
  end;
end;

procedure TOF_PiedPort.ChargeGrille;
begin
  TOBPied.PutGridDetail(G_Port, False, False, fColonnes, True);
  G_Port.RowCount := Max(NbRowsInit, G_Port.RowCount + 1);
  if TOBPied.Detail.Count = 0 then CreerTOBPLigne(1)
                              else CalculMontants;
  ControlColGrille(1);
  AffichePied(1);
  CalculTotal;
end;

function TOF_PiedPort.GetTOBPLigne(ARow: integer): TOB;
begin
  Result := nil;
  if ((ARow <= 0) or (ARow > TOBPied.Detail.Count)) then Exit;
  Result := TOBPied.Detail[ARow - 1];
end;

procedure TOF_PiedPort.InitRow(Row: integer);
var Col: integer;
  TOBPL: TOB;
begin
  TOBPL := GetTOBPLigne(Row);
  if TOBPL <> nil then TOBPL.InitValeurs;
  for Col := 0 to G_Port.ColCount do G_Port.cells[Col, Row] := '';
end;

procedure TOF_PiedPort.CreerTOBPLigne(ARow: integer);
var TT : TOB;
begin
  if ARow <> TOBPied.Detail.Count + 1 then exit;
  TT := TOB.Create('PIEDPORT', TOBPied, ARow - 1);
  AddlesChampsSupport(TT);
  InitRow(ARow);
end;

function TOF_PiedPort.LigVide(Row: integer): Boolean;
begin
  Result := True;
  if (G_Port.Cells[PT_Port, Row] <> '') then result := False;
end;

procedure TOF_PiedPort.AddlesChampsSupport(TT : TOB);
begin
  TT.AddChampSupValeur('SEL','');
end;

procedure TOF_PiedPort.InsertLigne(ARow: Integer);
var TT : TOB;
begin
  if Action = taConsult then Exit;
  if ARow < 1 then Exit;
  if LigVide(ARow) then exit;
  if (ARow > TOBPied.Detail.Count) then Exit;
  G_Port.CacheEdit;
  G_Port.SynEnabled := False;
  TT := TOB.Create('PIEDPORT', TOBPied, ARow - 1);
  AddlesChampsSupport(TT);
  G_Port.InsertRow(ARow);
  G_Port.Row := ARow;
  G_PORT.Col := PT_Port;
  InitRow(ARow);
  G_Port.ElipsisButton := True;
  G_Port.MontreEdit;
  G_Port.SynEnabled := True;
  ControlColGrille(ARow);
  BDELLINE.Enabled := True;
end;

procedure TOF_PiedPort.SupprimeLigne(ARow: Integer);
begin
  if Action = taConsult then Exit;
  if ARow < 1 then Exit;
  if (ARow > TOBPied.Detail.Count) then Exit;
  G_Port.CacheEdit;
  G_Port.SynEnabled := False;
  G_Port.DeleteRow(ARow);
  if (ARow = TOBPied.Detail.Count) then CreerTOBPLigne(ARow + 1);
  TOBPied.Detail[ARow - 1].Free;
  if G_Port.RowCount < NbRowsInit then G_Port.RowCount := NbRowsInit;
  G_Port.MontreEdit;
  G_Port.SynEnabled := True;
  ControlColGrille(ARow);
  AffichePied(ARow);
end;

procedure TOF_PiedPort.AfficheLaligne(ARow: Integer);
var TOBPL: TOB;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  TOBPL.PutLigneGrid(G_Port, ARow, False, False, fColonnes);
end;

procedure TOF_PiedPort.ChercherPort;
var CODE: THCritMaskEdit;
  Coord: TRect;
begin
  Coord := G_Port.CellRect(G_Port.Col, G_Port.Row);
  CODE := THCritMaskEdit.Create(ECRAN);
  CODE.Parent := G_Port;
  CODE.Top := Coord.Top;
  CODE.Left := Coord.Left;
  CODE.Width := 3;
  CODE.Visible := False;
  CODE.Text := G_Port.Cells[G_Port.Col, G_Port.Row];
  CODE.DataType := 'GCPORT';
  CODE.Plus := ' AND (GPO_TYPEFRAIS <> "B00")';
  //
  LookupCombo(CODE);
  //
  if CODE.Text <> '' then G_Port.Cells[G_Port.Col, G_Port.Row] := CODE.Text;
  //
  CODE.Destroy;
end;

procedure TOF_PiedPort.InitLaLigne(ARow: integer; TOBP, TOBPL: TOB);
begin
  if TOBPL = nil then Exit;
  if TOBP = nil then Exit;
  ChargeTobPiedPort(TOBPL, TOBP, TOBPiece);

  CalculMontant(TOBPL, ARow);
  CalculTotal;
  AffichePied(ARow);
  ControlColGrille(ARow);
end;

procedure TOF_PiedPort.ControlColGrille(ARow: Integer);
var TOBPL: TOB;
  TypePort: string;
  Franco: Boolean;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  TypePort := TOBPL.GetValue('GPT_TYPEPORT');
  if TOBPL.GetValue('GPT_FRANCO') = 'X' then Franco := True else Franco := False;
  (*
  if (TypePort = 'MT') or (TypePort = '') then
  begin
    G_Port.ColLengths[PT_BaseHT] := -1;
    G_Port.ColLengths[PT_Pour] := -1;
    G_Port.ColLengths[PT_MontHT] := 0;
  end else
  begin
    G_Port.ColLengths[PT_Base] := 0;
    G_Port.ColLengths[PT_Pour] := 0;
    G_Port.ColLengths[PT_MontHT] := -1;
  end;
  *)
  if TypePort = 'MI' then
  begin
    GPT_MONTANTMINI.Visible := True;
    GPT_MONTANTMINI.TabStop := True;
    TGPT_MONTANTMINI.Visible := True;
  end else
  begin
    GPT_MONTANTMINI.Visible := False;
    GPT_MONTANTMINI.TabStop := False;
    TGPT_MONTANTMINI.Visible := False;
  end;
  if Franco then
  begin
    GPT_MINIMUM.Visible := True;
    GPT_MINIMUM.TabStop := True;
    tGPT_MINIMUM.Visible := True;
  end else
  begin
    GPT_MINIMUM.Visible := False;
    GPT_MINIMUM.TabStop := False;
    TGPT_MINIMUM.Visible := False;
  end;
  if TOBPL.GetValue('GPT_CODEPORT') <> '' then
  begin
//    GPT_FRANCO.Visible := True;
    BZOOM.Enabled := True;
    if Action <> taConsult then
    begin
      BNEWLINE.Enabled := True;
      BDELLINE.Enabled := True;
      if TypePort = 'MT' then
      begin
        BAFFECT.Enabled := False;
        BVERROU.Visible := False;
      end
      else
      begin
        BAFFECT.Enabled := True;
        BVERROU.Visible := True;
      end
    end else
    begin
      BNEWLINE.Enabled := False;
      BDELLINE.Enabled := False;
      BAFFECT.Enabled := False;
    end;
  end else
  begin
//    GPT_FRANCO.Visible := False;
    BZOOM.Enabled := False;
    BNEWLINE.Enabled := False;
    BDELLINE.Enabled := False;
    BAFFECT.Enabled := False;
  end;
  (*
  if TOBPL.Getvalue('GPT_FRAISREPARTIS') = 'X' then
  begin
//    GPT_FRANCO.visible := False;
    GPT_MINIMUM.Visible := False;
    GPT_MINIMUM.TabStop := False;
    TGPT_MINIMUM.Visible := False;
  end;
  *)
    GPT_FRANCO.visible := false;
end;

function TOF_PiedPort.SortDeLaLigne: boolean;
var ACol, ARow: integer;
  Cancel: boolean;
begin
  Result := False;
  ACol := G_Port.Col;
  ARow := G_Port.Row;
  Cancel := False;
  G_PortCellExit(nil, ACol, ARow, Cancel);
  if Cancel then Exit;
  G_PortRowExit(nil, ACol, Cancel, False);
  if Cancel then Exit;
  Result := True;
end;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}

procedure TOF_PiedPort.G_PortElipsisClick(Sender: TObject);
begin
  if Action = taConsult then Exit;
  if G_Port.Col = PT_Port then
  begin
    ChercherPort;
  end;
end;


function TOF_PiedPort.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
var TOBPL : TOB;
begin
  TOBPL := nil;
  if Arow <= TOBPied.detail.count then TOBPL := GetTOBPLigne(ARow);
  result := true;
  if TOBPL = nil then exit; 
  if Grille.ColWidths[acol] = 0 then BEGIN result := false; exit; END;
  if Grille.ColLengths[acol] = -1 then BEGIN result := false; exit; END;
  if (TOBPL.GetBoolean('GPT_RETENUEDIVERSE')) then
  begin
    if (Acol=PT_TVA)  then
    begin
      result := false;
      exit;
    end;
  end;
  if (TOBPL.GetValue('GPT_FRAISREPARTIS') = 'X')  then
    begin
//    if ((ACol = PT_Pour) or( Acol=PT_BaseHT) or (Acol=PT_BaseTTC)  ) then
    if ((Acol=PT_MontTTC) or (Acol=PT_BaseTTC) or (ACol=PT_Mini) or (Acol = PT_MontHT) ) then
    begin
      result := false;
      exit;
    end;
  end else if (TOBPL.GetString('GPT_TYPEPORT')='HT') OR (TOBPL.GetString('GPT_TYPEPORT')='') then
  begin
    // Pourcentage du HT
    if ((Acol=PT_MontTTC) or (Acol=PT_BaseTTC) or (ACol=PT_Mini) or (Acol = PT_MontHT) ) then
    begin
      result := false;
      exit;
    end;
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MI') then
  begin
    // Pourcentage HT avec mini
    if ((Acol=PT_MontTTC) or (Acol=PT_BaseTTC) or (Acol = PT_MontHT) ) then
    begin
      result := false;
      exit;
    end;
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MIC') then
  begin
    // Pourcentage TTC avec mini
    if ((Acol=PT_MontHT) or (Acol=PT_BaseHT) or (Acol = PT_MontTTC) ) then
    begin
      result := false;
      exit;
    end;
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MT') then
  begin
    // Montant fixe HT
    if ((Acol=PT_BASEHT) or (ACol = PT_Pour) or (Acol=PT_BaseTTC) or (Acol=PT_MontTTC) or (ACol=PT_Mini) ) then
    begin
      result := false;
      exit;
    end;
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MTC') then
  begin
    // Montant fixe TTC
    if ((Acol=PT_BASEHT) or (ACol = PT_Pour) or (Acol=PT_BaseTTC) or (Acol=PT_MontHT) or (ACol=PT_Mini) ) then
    begin
      result := false;
      exit;
    end;
  end else if (TOBPL.GetString('GPT_TYPEPORT')='PT') then
  begin
    // Pourcentage du TTC
    if ((Acol=PT_MontHT) or (Acol=PT_BaseHT) or (ACol=PT_Mini) or (Acol=PT_MontTTC) ) then
    begin
      result := false;
      exit;
    end;
  end;
end;

procedure TOF_PiedPort.ZoneSuivanteOuOk(Grille : THGrid;var ACol, ARow : integer; var Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  ChgLig := (Grille.Row <> ARow);
  ChgSens := False;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol <= Grille.Col)) then Sens := 1;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  while not ZoneAccessible(Grille,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := Grille.RowCount ;
      // ---
      if ((ACol = Grille.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          // Ajout d'une ligne
          break;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then
      begin
        Inc(ACol);
      end else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.ColCount;
        ChgLig := False;
      end;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

procedure TOF_PiedPort.G_PortCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(G_PORT,ACol, ARow, Cancel);
  if not Cancel then
  begin
    if (G_Port.Col <> PT_Port) and (G_Port.Cells[PT_Port, G_Port.Row] = '') then
    begin
      G_Port.Col := PT_Port;
    end;
    G_Port.ElipsisButton := (G_Port.Col = PT_Port);
    StCellCur := G_Port.Cells[G_Port.Col, G_Port.Row];
  end;
end;

procedure TOF_PiedPort.G_PortCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  if ACol = PT_Port then TraiterPort(ACol, ARow, Cancel) else
    if ACol = PT_Libel then TraiterLibelle(ACol, ARow, Cancel) else
    if ACol = PT_BaseHT then TraiterBase(ACol, ARow, Cancel) else
    if ACol = PT_BaseTTC then TraiterBaseTTC(ACol, ARow, Cancel) else
    if ACol = PT_Pour then TraiterPourcent(ACol, ARow, Cancel) else
    if ACol = PT_Mini then TraiterMini(ACol, ARow, Cancel);
    if ACol = PT_MontHT then TraiterMontant(ACol, ARow, Cancel);
    if ACol = PT_MontTTC then TraiterMontantTTC(ACol, ARow, Cancel);
    if ACol = PT_TVA then TraiterTVA(ACol, ARow, Cancel);
end;

procedure TOF_PiedPort.G_PortRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow: Integer;
begin
  G_Port.InvalidateRow(ou);
  if Ou >= G_Port.RowCount - 1 then G_Port.RowCount := G_Port.RowCount + NbRowsPlus;
  ;
  ARow := Min(Ou, TOBPied.detail.count + 1);
  if (ARow = TOBPied.detail.count + 1) and (not LigVide(ARow - 1)) then
  begin
    CreerTOBPLigne(ARow);
  end;
  if Ou > TOBPied.detail.count then
  begin
    G_Port.Row := TOBPied.detail.count;
  end;
  AffichePied(ARow);
  ControlColGrille(ARow);
end;

procedure TOF_PiedPort.G_PortRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  G_Port.InvalidateRow(ou);
  if LigVide(Ou) then G_Port.Row := Min(G_Port.Row, Ou);
end;

procedure TOF_PiedPort.G_PortColumnWidthsChanged(Sender: TObject);
var Coord: TRect;
begin
  if PCUMUL.ControlCount <= 0 then exit;
  Coord := G_Port.CellRect(PT_MontHT, 0);
  CUMMONTANTHT.Left := Coord.Left + 1;
  CUMMONTANTHT.Width := G_Port.ColWidths[PT_MontHT] + 1;
  Coord := G_Port.CellRect(PT_MontTTC, 0);
  CUMMONTANTTC.Left := Coord.Left + 1;
  CUMMONTANTTC.Width := G_Port.ColWidths[PT_MontTTC] + 1;
end;

{$IFDEF EAGLCLIENT}
{$ELSE}

procedure TOF_PiedPort.Raye(Grid : THgrid; Canvas : TCanvas; ACol,ARow : integer);
var  Arect: Trect;
begin
  Arect := G_Port.CellRect(Acol, Arow);
  if (Acol <> PT_MontTTC) and (ACol <> PT_MontHT) then
  begin
    Canvas.Brush.Color := G_Port.FixedColor;
    Canvas.FillRect(ARect);
  end;
  //
  Canvas.Brush.Color := clActiveCaption;
  Canvas.Brush.Style := bsBDiagonal;
  Canvas.Pen.Color := clActiveCaption;
  Canvas.Pen.Mode := pmCopy;
  Canvas.Pen.Style := psClear;
  Canvas.Pen.Width := 1;
  Canvas.Rectangle(Arect.Left, Arect.Top, Arect.Right + 1, Arect.Bottom + 1);
end;


procedure TOF_PiedPort.DessineCellPort(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var Triangle: array[0..2] of TPoint;
  Arect: Trect;
  TOBPL: TOB;
begin
  if Arow < G_Port.Fixedrows then exit;
  if (gdFixed in AState) and (ACol = 0) then
  begin
    Arect := G_Port.CellRect(Acol, Arow);
    Canvas.Brush.Color := G_Port.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = G_Port.row) then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Triangle[1].X := ((ARect.Left + ARect.Right) div 2);
      Triangle[1].Y := ((ARect.Top + ARect.Bottom) div 2);
      Triangle[0].X := Triangle[1].X - 5;
      Triangle[0].Y := Triangle[1].Y - 5;
      Triangle[2].X := Triangle[1].X - 5;
      Triangle[2].Y := Triangle[1].Y + 5;
      if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle);
    end;
  end;
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then Exit;
  if (TOBPL.GetValue('GPT_FRAISREPARTIS') = 'X')  then
  begin
//    if ((ACol = PT_Pour) or( Acol=PT_BaseHT) or (Acol=PT_BaseTTC) or (ACol=PT_Mini)) then
    if ((Acol=PT_MontTTC) or (Acol=PT_MontHT) or (Acol=PT_BaseTTC) or (ACol=PT_Mini) ) then
      raye(G_Port,Canvas,ACol,ARow);
  end else if (TOBPL.GetString('GPT_TYPEPORT')='HT') OR (TOBPL.GetString('GPT_TYPEPORT')='') then
  begin
    // Pourcentage du HT
    if ((Acol=PT_MontTTC) or (Acol=PT_BaseTTC) or (ACol=PT_Mini) ) then
      raye(G_Port,Canvas,ACol,ARow);
    if ((Acol=PT_TVA) and (TOBPL.GetBoolean('GPT_RETENUEDIVERSE'))) then
      raye(G_Port,Canvas,ACol,ARow);
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MI') then
  begin
    // Pourcentage HT avec mini
    if ((Acol=PT_MontTTC) or (Acol=PT_BaseTTC)) then
      raye(G_Port,Canvas,ACol,ARow);
    if ((Acol=PT_TVA) and (TOBPL.GetBoolean('GPT_RETENUEDIVERSE'))) then
      raye(G_Port,Canvas,ACol,ARow);
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MIC') then
  begin
    // Pourcentage TTC avec mini
    if ((Acol=PT_MontHT) or (Acol=PT_BaseHT)) then
      raye(G_Port,Canvas,ACol,ARow);
    if ((Acol=PT_TVA) and (TOBPL.GetBoolean('GPT_RETENUEDIVERSE'))) then
      raye(G_Port,Canvas,ACol,ARow);
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MT') then
  begin
    // Montant fixe HT
    if ((Acol=PT_BASEHT) or (ACol = PT_Pour) or (Acol=PT_BaseTTC) or (Acol=PT_MontTTC) or (ACol=PT_Mini) ) then
      raye(G_Port,Canvas,ACol,ARow);
    if ((Acol=PT_TVA) and (TOBPL.GetBoolean('GPT_RETENUEDIVERSE'))) then
      raye(G_Port,Canvas,ACol,ARow);
  end else if (TOBPL.GetString('GPT_TYPEPORT')='MTC') then
  begin
    // Montant fixe TTC
    if ((Acol=PT_BASEHT) or (ACol = PT_Pour) or (Acol=PT_BaseTTC) or (Acol=PT_MontHT) or (ACol=PT_Mini)) then
      raye(G_Port,Canvas,ACol,ARow);
    if ((Acol=PT_TVA) and (TOBPL.GetBoolean('GPT_RETENUEDIVERSE'))) then
      raye(G_Port,Canvas,ACol,ARow);
  end else if (TOBPL.GetString('GPT_TYPEPORT')='PT') then
  begin
    // Pourcentage du TTC
    if ((Acol=PT_MontHT) or (Acol=PT_BaseHT) or (ACol=PT_Mini) ) then
      raye(G_Port,Canvas,ACol,ARow);
    if ((Acol=PT_TVA) and (TOBPL.GetBoolean('GPT_RETENUEDIVERSE'))) then
      raye(G_Port,Canvas,ACol,ARow);
  end;

end;
{$ENDIF}

{==============================================================================================}
{============================ Manipulation des Champs GRID ====================================}
{==============================================================================================}

procedure TOF_PiedPort.TraiterPort(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL, TOBP: TOB;
  St, St1: string;
  QQ: TQuery;
  DD: TDateTime;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  St := UpperCase(G_Port.Cells[ACol, ARow]);
  G_Port.Cells[ACol, ARow] := St;
  if St = '' then
  begin
    G_Port.Cells[ACol, ARow] := TOBPL.GetValue('GPT_CODEPORT');
    exit;
  end;

  if TOBPL.GetValue('GPT_CODEPORT') <> St then
  begin
    TOBP := TOBPort.FindFirst(['GPO_CODEPORT'], [St], False);
    if TOBP = nil then
    begin
      QQ := OpenSQL('Select * from PORT Where GPO_CODEPORT="' + St + '"', True);
      if not QQ.EOF then
      begin
        TOBP := TOB.Create('PORT', TOBPort, -1);
        TOBP.SelectDB('', QQ);
        Ferme(QQ);
      end else // JT - eQualité 10932
      begin
        Ferme(QQ);
        St1 := '0;?caption?;Ce code n''existe pas.;X;O;O;O;';
        HShowMessage(St1, Ecran.Caption, '');
        Cancel := True;
        Exit;
      end; // Fin JT
    end;

    DD := TOBP.GetValue('GPO_DATESUP');
    if ((DD < V_PGI.DateEntree) and (DD > iDate1900)) then
    begin
      St1 := '0;?caption?;Ce code est supprimé;X;O;O;O;';
      HShowMessage(St1, Ecran.Caption, '');
      Cancel := True;
      Exit;
    end;
    if (TOBP.GetValue('GPO_FRAISREPARTIS')='X') and ( not isPieceGerableFraisDetail (TOBPiece.GetValue('GP_NATUREPIECEG'))) then
    begin
      St1 := '0;?caption?;Ce type de frais n''est pas utilisable dans ce document;X;O;O;O;';
      HShowMessage(St1, Ecran.Caption, '');
      Cancel := True;
      Exit;
    end;
    if (TOBP.GetValue('GPO_FERME') = 'X') then
    begin
      St1 := '0;?caption?;Ce code est fermé. Confirmez vous ce choix ?;Q;YN;Y;N;';
      if HShowMessage(St1, Ecran.Caption, '') <> mrYes then
      begin
        Cancel := True;
        Exit;
      end;
    end;
    if TOBP <> nil then
    begin
      TOBPL.PutValue('GPT_CODEPORT', St);
      InitLaLigne(ARow, TOBP, TOBPL);
      TraiteLaLigne (Acol,Arow);
    end else
    begin
      G_Port.Col := ACol;
      G_Port.Row := ARow;
      ChercherPort;
      Cancel := True;
    end;
  end;

end;

procedure TOF_PiedPort.TraiterLibelle(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
  St,verrou: string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  St := G_Port.Cells[ACol, ARow];
  TOBPL.PutValue('GPT_LIBELLE', St);
end;

procedure TOF_PiedPort.TraiterBase(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
  Base: Double;
  Verrou : string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  Base := Valeur(G_Port.Cells[ACol, ARow]);
  if Base >= 0 then
  begin
    if EnHT then TOBPL.PutValue('GPT_BASEHTDEV', Base) else TOBPL.PutValue('GPT_BASETTCDEV', Base);
  end else Cancel := True;

  if TOBPL.GetValue('GPT_TYPEPORT') <> 'MT' then 
    CalculBasePiedPort(TOBPL, EnHT);

  CalculMontant(TOBPL, ARow);
  CalculTotal;
end;

procedure TOF_PiedPort.TraiterBaseTTC(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
  Base: Double;
  verrou : string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  Base := Valeur(G_Port.Cells[ACol, ARow]);
  if Base >= 0 then
  begin
    TOBPL.PutValue('GPT_BASETTCDEV', Base);
  end else Cancel := True;

  if TOBPL.GetValue('GPT_TYPEPORT') <> 'MT' then
    CalculBasePiedPort(TOBPL, EnHT);

  CalculMontant(TOBPL, ARow);
  CalculTotal;
end;

procedure TOF_PiedPort.TraiterPourcent(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
  Pourcent: Double;
  verrou : string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  Pourcent := Valeur(G_Port.Cells[ACol, ARow]);
//  if Pourcent >= 0 then TOBPL.PutValue('GPT_POURCENT', Pourcent) else Cancel := True;
  TOBPL.PutValue('GPT_POURCENT', Pourcent);
  CalculMontant(TOBPL, ARow);
  CalculTotal;
end;

procedure TOF_PiedPort.TraiterMini(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
  Mont: Double;
  verrou : string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  Mont := Valeur(G_Port.Cells[ACol, ARow]);
  TOBPL.PutValue('GPT_MONTANTMINI', Mont);
  CalculMontant(TOBPL, ARow);
  CalculTotal;
end;

procedure TOF_PiedPort.TraiterMontant(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
  Mont: Double;
  Typeport : string;
  verrou : string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  TypePort := TOBPL.GetString('GPT_TYPEPORT');
  Mont := Valeur(G_Port.Cells[ACol, ARow]);
  TOBPL.PutValue('GPT_TOTALHTDEV', Mont);
  if TypePort = 'MT' then TOBPL.PutValue('GPT_BASEHTDEV', Mont);
  CalculMontant(TOBPL, ARow);
  CalculTotal;
end;

procedure TOF_PiedPort.TraiterTVA(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
    ZZ : string;
  verrou : string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  ZZ := G_Port.CellValues[ACol, ARow];
  TOBPL.PutValue('GPT_FAMILLETAXE1', ZZ);
  CalculMontant(TOBPL, ARow);
  CalculTotal;
end;

procedure TOF_PiedPort.TraiterMontantTTC(ACol, ARow: integer; var Cancel: Boolean);
var TOBPL: TOB;
  Mont: Double;
  TypePort : string;
  verrou : string;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then exit;
  TypePort := TOBPL.GetValue('GPT_TYPEPORT');
  verrou := TOBPL.getString('GPT_BLOQUEFRAIS');
  if verrou = 'VER' then begin G_Port.Cells[ACol, ARow] := StCellCur; exit; end;
  Mont := Valeur(G_Port.Cells[ACol, ARow]);
  TOBPL.PutValue('GPT_TOTALTTCDEV', Mont);
  if TypePort = 'MTC' then TOBPL.PutValue('GPT_BASETTCDEV', Mont);
  CalculMontant(TOBPL, ARow);
  CalculTotal;
end;

procedure TOF_PiedPort.CalculMontant(TOBPL: TOB; ARow: integer);
begin
  CalculMontantsPiedPort(TobPiece, TOBPL,TOBBases);
  AfficheLaligne(ARow);
end;

procedure TOF_PiedPort.CalculTotal;
var i_row: Integer;
  Mont,Mont1: Double;
  TOBPL: TOB;
begin
  Mont := 0; Mont1 := 0;
  for i_row := 1 to G_Port.RowCount - 1 do
  begin
    TOBPL := GetTOBPLigne(i_Row);
    if TOBPL = nil then continue;
    if TOBPL.GetValue('GPT_FRAISREPARTIS') <> 'X' then
    begin
      Mont := Mont + Valeur(G_Port.Cells[PT_MontHT, i_row]);
      Mont1 := Mont1 + Valeur(G_Port.Cells[PT_MontTTC, i_row]);
  end;
  end;
  CUMMONTANTHT.Value := Mont;
  CUMMONTANTTC.Value := Mont1;
end;

{==============================================================================================}
{================================= Evènements Pied ============================================}
{==============================================================================================}

procedure TOF_PiedPort.GPT_FRANCOClick(Sender: TObject);
var TOBPL: TOB;
  Franco: Boolean;
begin
  TOBPL := GetTOBPLigne(G_Port.Row);
  if TOBPL = nil then exit;
  if GPT_FRANCO.Checked then Franco := True else Franco := False;
  if Franco then
  begin
    GPT_MINIMUM.Visible := True;
    GPT_MINIMUM.TabStop := True;
    tGPT_MINIMUM.Visible := True;
  end else
  begin
    GPT_MINIMUM.Visible := False;
    GPT_MINIMUM.TabStop := False;
    TGPT_MINIMUM.Visible := False;
  end;
  if Franco then TOBPL.PutValue('GPT_FRANCO', 'X')
  else TOBPL.PutValue('GPT_FRANCO', '-');
  CalculMontant(TOBPL, G_Port.Row);
  CalculTotal;
end;

procedure TOF_PiedPort.GPT_MINIMUMExit(Sender: TObject);
var TOBPL: TOB;
  Mini: Double;
begin
  TOBPL := GetTOBPLigne(G_Port.Row);
  if TOBPL = nil then exit;
  Mini := GPT_MINIMUM.Value;
  if Mini < 0 then
  begin
    GPT_MINIMUM.Value := TOBPL.GetValue('GPT_MINIMUM');
    Exit;
  end;
  TOBPL.PutValue('GPT_MINIMUM', Mini);
  CalculMontant(TOBPL, G_Port.Row);
  CalculTotal;
end;

procedure TOF_PiedPort.GPT_MONTANTMINIExit(Sender: TObject);
var TOBPL: TOB;
  Mini: Double;
begin
  TOBPL := GetTOBPLigne(G_Port.Row);
  if TOBPL = nil then exit;
  Mini := GPT_MONTANTMINI.Value;
  if Mini < 0 then
  begin
    GPT_MONTANTMINI.Value := TOBPL.GetValue('GPT_MONTANTMINI');
    Exit;
  end;
  TOBPL.PutValue('GPT_MONTANTMINI', Mini);
  CalculTotal;
end;

{==============================================================================================}
{============================= Manipulation du pied ===========================================}
{==============================================================================================}

procedure TOF_PiedPort.AffichePied(ARow: integer);
var TOBPL: TOB;
begin
  TOBPL := GetTOBPLigne(ARow);
  if TOBPL = nil then Exit;
  TOBPL.PutEcran(Ecran, PPIED);
end;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}

procedure TOF_PiedPort.BZOOMClick(Sender: TObject);
var TOBPL: TOB;
begin
  TOBPL := GetTOBPLigne(G_Port.Row);
  if TOBPL = nil then Exit;
  AglLanceFiche('GC', 'GCPORT', '', TOBPL.GetValue('GPT_CODEPORT'), 'ACTION=CONSULTATION');
end;

procedure TOF_PiedPort.BNEWLINEClick(Sender: TObject);
begin
  if Action = taConsult then Exit;
  InsertLigne(G_Port.Row);
end;

procedure TOF_PiedPort.BDELLINEClick(Sender: TObject);
begin
  if Action = taConsult then Exit;
  SupprimeLigne(G_Port.Row);
end;

procedure TOF_PiedPort.BAFFECTClick(Sender: TObject);
var Cancel: Boolean;
begin
  Cancel := False;
  if Action = taConsult then Exit;
  G_port.Cells[G_Port.col, G_Port.Row] := FloatToStr(TOTALHT);
  TraiterBase(G_Port.col, G_Port.Row, Cancel);
end;

procedure TOF_PiedPort.ChangeBloqueFrais(Valeur: string; ImgIndex: Integer);
begin
  TobPied.Detail[G_Port.Row - 1].P('GPT_BLOQUEFRAIS', Valeur);
  AfficheLaligne (G_Port.Row);
end;

procedure TOF_PiedPort.PMVERROU_OnPopup(Sender: TObject);
var
  i: Integer;
begin
  ChargeImageList;
  for i := 0 to TPopupMenu(Sender).Items.Count - 1 do
  begin
    if UpperCase(TPopupMenu(Sender).Items[i].Name) = 'MNVER' then
    begin
      TPopupMenu(Sender).Items[i].Visible := TobPied.Detail[G_Port.Row - 1].G('GPT_BLOQUEFRAIS') <> 'VER';
      V_PGI.GraphList.GetBitmap(IcoVER - 1, TPopupMenu(Sender).Items[i].BitMap);
    end
    else if UpperCase(TPopupMenu(Sender).Items[i].Name) = 'MNVEM' then
    begin
      TPopupMenu(Sender).Items[i].Visible := TobPied.Detail[G_Port.Row - 1].G('GPT_BLOQUEFRAIS') <> 'VEM';
      V_PGI.GraphList.GetBitmap(IcoVEM - 1, TPopupMenu(Sender).Items[i].BitMap);
    end
    else if UpperCase(TPopupMenu(Sender).Items[i].Name) = 'MNNON' then
    begin
      TPopupMenu(Sender).Items[i].Visible := TobPied.Detail[G_Port.Row - 1].G('GPT_BLOQUEFRAIS') <> 'NON';
      V_PGI.GraphList.GetBitmap(IcoNON - 1, TPopupMenu(Sender).Items[i].BitMap);
    end;
  end;
end;

procedure TOF_PiedPort.MnNon_OnClick(Sender: TObject);
begin
  ChangeBloqueFrais('NON', IcoNON);
end;

procedure TOF_PiedPort.MnVem_OnClick(Sender: TObject);
begin
  ChangeBloqueFrais('VEM', IcoVEM);
end;

procedure TOF_PiedPort.MnVer_OnClick(Sender: TObject);
begin
  ChangeBloqueFrais('VER', IcoVER);
end;

initialization
  registerclasses([TOF_PiedPort]);

end.
