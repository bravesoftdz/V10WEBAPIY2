{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EXTOURNE ()
Mots clefs ... : TOF;EXTOURNE
*****************************************************************}

unit CPANOECR;

//================================================================================
// Interface
//================================================================================
interface

uses
  StdCtrls,
  Windows, // Pour VK_SPACE
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  fe_main,
  PrintDBG, // pour le prindDB
{$ENDIF}
  Saisie,
  SaisBor,
  forms,
  sysutils,
  Graphics,
  Dialogs,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  HTB97,
  HDB,
  HQry,
  uTOB,
  HStatus,
  HSysMenu,
  ed_tools,
  Vierge,
  Ent1,
  HPanel,
  SaisUtil,
  UtilSais,
  DelVisuE,
  AGLInit,
  UtilSoc,
  HXLSPAS,
  menus,
  UObjFiltres; {Pour la gestion des filtres}

//==================================================
// Definition des Constante
//==================================================
const
  NOM_FILTRE = 'ANOECR';
  // L'accès au paramétrage est interdit. Ces colonnes sont donc fixes.
  COL_NATURE = 1;
  COL_DATE = 2;
  COL_PIECE = 3;
  COL_GENERAL = 4;
  COL_AUXILIAIRE = 5;
  COL_JOURNAL = 6;
  COL_REFINTERNE = 7;
  COL_LIBELLE = 8;
  COL_DEBIT = 9;
  COL_CREDIT = 10;
  COL_SOLDE = 11;
  COL_PERIODE = 12;
  COL_LIGNE = 13;
  COL_SELECTION = 14;
  COL_QUALIFPIECE = 15;
  COL_MODESAISIE = 16;
  COL_NOMBRE = 17;

procedure CPLanceFiche_CPANOECR;

//==================================================
// Definition de class
//==================================================
type
  TOF_ANOECR = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    BExport: TToolBarButton97;
    BImprimer: TToolBarButton97;
    BValider: TToolBarButton97;
    BSelectAll: TToolBarButton97;
    SD: TSaveDialog;
    FHSystemMenu: THSystemMenu;
    FGrille: THGrid;
    FPages: TPageControl;
    FListeEcr: TOB;
    FListePiece: TOB;
    FJournal: THValComboBox;
    FExercice: THValComboBox;
    FModeSaisie: string;
    FRapport: TList;
    FTotalLigne: integer;
    FTotalDebit: double;
    FTotalCredit: double;
    FSelectionLigne: integer;
    FSelectionDebit: double;
    FSelectionCredit: double;
    procedure OnEcranResize(Sender: TObject);
    procedure OnChercheClick(Sender: TOBject);
    procedure OnToutSelectionnerClick(Sender: TOBject);
    procedure FListeDblClick(Sender: TObject);
    procedure OnChangeLargeurColonneGrille(Sender: TObject);
    procedure ChargeLesEcritures;
    procedure ReFormeListe(T: integer);
    procedure BExportClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure RazSelection;
    procedure ToutSelectionner(DepuisLigne: integer);
    procedure OnGridFlipSelection(Sender: TObject);
    procedure OnE_EXERCICEChange(Sender: TObject);
    procedure InitMultiCriteres;
  protected
    ObjFiltre     : TObjFiltre;
end;

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  paramsoc;


procedure CPLanceFiche_CPANOECR;
begin
  AGLLanceFiche('CP', 'CPANOECR', '', '', '');
end;

procedure TOF_ANOECR.OnNew;
begin
  inherited;
end;

procedure TOF_ANOECR.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (csDestroying in Ecran.ComponentState) then
    Exit;
  if (key = VK_cherche) then
    OnChercheClick(nil)
  else if (key = VK_choixmul) then
    FListeDblClick(nil)
  else if (key = vk_liste) then
  begin
    if FGrille.Focused then
    begin
      TPageControl(GetControl('PAGES')).SetFocus;
      Hent1.NextControl(Ecran);
    end
    else
      FGrille.SetFocus;
  end
  else if (key = vk_applique) then
  begin
    Key := 0;
    OnChercheClick(nil);
    if FGrille.CanFocus then
      FGrille.SetFocus;
  end
  else if (key = vk_valide) then
  begin
    Key := 0;
  end
  else if FGrille.Focused then
  begin
    if ((key = Ord('A')) and (ssCtrl in Shift)) then
    begin
      Key := 0;
      TToolBarButton97(GetControl('BSelectAll')).Click;
      // Pièce précédente
    end
    else if (key = Ord('P')) then
    begin
      Key := 0;
      // Pièce suivante
    end
    else if (key = Ord('S')) then
    begin
      Key := 0;
    end
    else if (key = Ord('D')) then
    begin
      Key := 0;
    end
    else if ((key = Ord('N')) and (ssCtrl in Shift)) then
    begin
      Key := 0;
      TToolBarButton97(GetControl('BSELECTIONFIN')).Click;
    end
    else if (key = VK_F11) then
    begin
      Key := 0;
      TPopupMenu(GetControl('POPF11')).Popup(Mouse.CursorPos.x,
        Mouse.CursorPos.y);
    end;
  end;
end;

procedure TOF_ANOECR.OnE_EXERCICEChange(Sender: TObject);
begin
  inherited;
  if GetControlText('E_EXERCICE')<>'' then
    ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE'),
      GetControl('E_DATECOMPTABLE_'))
  else
  begin
    SetControlText('E_DATECOMPTABLE', StDate1900) ;
    SetControlText('E_DATECOMPTABLE_',StDate2099) ;
  end;
  

end;


procedure TOF_ANOECR.OnDelete;
begin
  inherited;
end;

procedure TOF_ANOECR.OnUpdate;
begin
  inherited;
  ChargeLesEcritures;
end;

procedure TOF_ANOECR.OnLoad;
begin
  inherited;
  ObjFiltre.Charger;
  ChargeLesEcritures;
end;

procedure TOF_ANOECR.BImprimerClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  PrintDBGrid(FGrille, nil, Ecran.Caption, '');
{$ENDIF}
end;

procedure TOF_ANOECR.BValiderClick(Sender: TObject);
var
  i : integer;
  Piece: integer;
  Journal, QualifPiece: string;
  DateComptable: TDateTime;
  St: string;

begin
  if not ((TCheckBox(GetControl('ANOLETT')).Checked = true)
    or (TCheckBox(GetControl('INCOPOI')).Checked = true)) then
    exit;
  if FListeEcr.Detail.Count = 0 then
    exit;
  begintrans;
  for i := 0 to FGrille.nbSelected - 1 do
  begin
    FGrille.GotoLeBookMark(i);
    Piece := StrToInt(FGrille.Cells[COL_PIECE, FGrille.Row]);
    Journal := FGrille.Cells[COL_JOURNAL, FGrille.Row];
    QualifPiece := FGrille.Cells[COL_QUALIFPIECE, FGrille.Row];
    DateComptable := StrToDate(FGrille.Cells[COL_DATE, FGrille.Row]);
    St := 'UPDATE ECRITURE SET E_TRESOLETTRE="-", E_QUALIFPIECE="N"'
      + ' WHERE E_DATECOMPTABLE="' + USDateTime(DateComptable) + '"'
      + ' AND E_JOURNAL="' + Journal + '" AND E_NUMEROPIECE=' + IntToStr(Piece);
    try
      executeSQL(St);
      CommitTrans;
    except
      V_PGI.IoError := oeUnknown;
      Rollback;
    end;
  end;
  ChargeLesEcritures;
  RazSelection;
end;

procedure TOF_ANOECR.RazSelection;
var
  i: integer;
begin
  inherited;
  for i := 1 to FGrille.RowCount - 1 do
  begin
    if i <= FListeEcr.Detail.Count then
      FListeEcr.Detail[i - 1].AddChampSupValeur('SELECTION', '-');
    FGrille.Cells[COL_SELECTION, i] := '-';
    FGrille.OnFlipSelection := nil;
    if FGrille.IsSelected(i) then
      FGrille.FlipSelection(i);
    //FGrille.OnFlipSelection := OnGridFlipSelection;
  end;
end;

procedure TOF_ANOECR.OnToutSelectionnerClick(Sender: TOBject);
begin
  inherited;
  ToutSelectionner(0);
end;

procedure TOF_ANOECR.ToutSelectionner(DepuisLigne: integer);
var
  i: integer;
  cNew : Char;
begin
  inherited;
  if DepuisLigne = 0 then
    DepuisLigne := 1;
  FGrille.OnFlipSelection := nil;
  if FGrille.Row > (FListeEcr.Detail.Count - 1) then
    exit;
  cNew := '-';
  if FListeEcr.Detail[FGrille.Row - 1].GetValue('SELECTION') = '+' then
  else
    cNew := '+';
  for i := DepuisLigne to FGrille.RowCount - 1 do
  begin
    (*if FListeEcr.Detail[i-1].GetValue('SELECTION')=cOld then
    begin*)
    FListeEcr.Detail[i - 1].PutValue('SELECTION', cNew);
    FGrille.Cells[COL_SELECTION, i] := cNew;
    FGrille.FlipSelection(i);
    //end;
  end;
end;

procedure TOF_ANOECR.OnGridFlipSelection(Sender: TObject);
var
   iRef: integer;
begin
  inherited;
  // Mémorisation de la ligne de référence
  iRef := FGrille.Row;
  // Mise à jour de la ligne de référence dans la TOB
  if FListeEcr.Detail[iRef - 1].GetValue('SELECTION') = '+' then
  begin
    FListeEcr.Detail[iRef - 1].PutValue('SELECTION', '-');
    FGrille.Cells[COL_SELECTION, iRef] := '-';
  end
  else
  begin
    FListeEcr.Detail[iRef - 1].PutValue('SELECTION', '+');
    FGrille.Cells[COL_SELECTION, iRef] := '+';
  end;
end;

procedure TOF_ANOECR.BExportClick(Sender: Tobject);
var
  lStHint: string;
begin
  if not ExJaiLeDroitConcept(ccExportListe, True) then
    exit;
  if SD.Execute then
  begin
    if SD.FilterIndex = 5 then //html
    begin
      lStHint := FGrille.Hint;
      FGrille.Hint := Ecran.Caption;
      ExportGrid(FGrille, nil, SD.FileName, SD.FilterIndex, True);
      FGrille.Hint := lStHint;
    end
    else
      ExportGrid(FGrille, nil, SD.FileName, SD.FilterIndex, True);
  end;
end;

procedure TOF_ANOECR.ReFormeListe(T: integer);
begin
  FGrille.ColCount := COL_NOMBRE;
  FGrille.ColWidths[COL_PERIODE] := -1;
  FGrille.ColLengths[COL_PERIODE] := -1;
  FGrille.ColWidths[COL_LIGNE] := -1;
  FGrille.ColLengths[COL_LIGNE] := -1;
  FGrille.ColWidths[COL_SELECTION] := -1;
  FGrille.ColLengths[COL_SELECTION] := -1;
  FGrille.ColWidths[COL_SOLDE] := -1;
  FGrille.ColLengths[COL_SOLDE] := -1;
  FGrille.ColWidths[0] := 20;
  FGrille.ColWidths[COL_NATURE] := 50;
  FGrille.ColWidths[COL_DATE] := 80;
  FGrille.ColWidths[COL_PIECE] := 40;
  FGrille.ColWidths[COL_GENERAL] := 100;
  FGrille.ColWidths[COL_AUXILIAIRE] := 100;
  FGrille.ColWidths[COL_JOURNAL] := 100;
  FGrille.ColLengths[COL_JOURNAL] := 100;
  FGrille.ColWidths[COL_REFINTERNE] := 100;
  FGrille.ColWidths[COL_LIBELLE] := 160;
  FGrille.ColWidths[COL_DEBIT] := 100;
  FGrille.ColWidths[COL_CREDIT] := 100;
  FGrille.ColWidths[COL_QUALIFPIECE] := 20;
  FGrille.ColLengths[COL_QUALIFPIECE] := 20;

  FGrille.ColWidths[COL_SELECTION] := -1;
  FGrille.ColLengths[COL_SELECTION] := -1;

  FGrille.Cells[COL_REFINTERNE, 0] := TraduireMemoire('Référence interne');
  FGrille.Cells[COL_LIBELLE, 0] := TraduireMemoire('Libellé');
  if T = 1 then
  begin
    FGrille.ColWidths[COL_NATURE] := -1;
    FGrille.ColLengths[COL_NATURE] := -1;
    FGrille.ColWidths[COL_PIECE] := -1;
    FGrille.ColLengths[COL_PIECE] := -1;
    FGrille.ColWidths[COL_DATE] := -1;
    FGrille.ColLengths[COL_DATE] := -1;
    FGrille.ColWidths[COL_GENERAL] := -1;
    FGrille.ColLengths[COL_GENERAL] := -1;
    FGrille.ColWidths[COL_AUXILIAIRE] := -1;
    FGrille.ColLengths[COL_AUXILIAIRE] := -1;
    FGrille.ColWidths[COL_JOURNAL] := -1;
    FGrille.ColLengths[COL_JOURNAL] := -1;
    FGrille.ColWidths[COL_LIBELLE] := -1;
    FGrille.ColLengths[COL_LIBELLE] := -1;
    FGrille.ColWidths[COL_REFINTERNE] := -1;
    FGrille.ColLengths[COL_REFINTERNE] := -1;
    FGrille.ColWidths[COL_DEBIT] := -1;
    FGrille.ColLengths[COL_DEBIT] := -1;
    FGrille.ColWidths[COL_CREDIT] := -1;
    FGrille.ColLengths[COL_SELECTION] := -1;
    FGrille.ColLengths[COL_SOLDE] := -1;
    FGrille.ColWidths[COL_QUALIFPIECE] := 20;
    FGrille.ColLengths[COL_QUALIFPIECE] := 20;
  end;
  if (T = 2) or (T = 4) then
  begin
    FGrille.ColWidths[COL_NATURE] := 20;
    FGrille.ColLengths[COL_NATURE] := 20;
    FGrille.ColWidths[COL_PIECE] := 20;
    FGrille.ColLengths[COL_PIECE] := 20;
    FGrille.ColWidths[COL_DATE] := 20;
    FGrille.ColLengths[COL_DATE] := 20;
    FGrille.ColWidths[COL_GENERAL] := 20;
    FGrille.ColLengths[COL_GENERAL] := 20;
    FGrille.ColWidths[COL_AUXILIAIRE] := 20;
    FGrille.ColLengths[COL_AUXILIAIRE] := 20;
    FGrille.ColWidths[COL_JOURNAL] := 20;
    FGrille.ColLengths[COL_JOURNAL] := 20;
    FGrille.ColWidths[COL_LIBELLE] := 20;
    FGrille.ColLengths[COL_LIBELLE] := 20;
    FGrille.ColWidths[COL_REFINTERNE] := 20;
    FGrille.ColLengths[COL_REFINTERNE] := 20;
    FGrille.ColWidths[COL_DEBIT] := 20;
    FGrille.ColLengths[COL_DEBIT] := 20;
    FGrille.ColWidths[COL_CREDIT] := 20;
    FGrille.ColLengths[COL_CREDIT] := 20;
    FGrille.ColWidths[COL_QUALIFPIECE] := 20;
    FGrille.ColLengths[COL_QUALIFPIECE] := 20;

  end;
  if T = 3 then
  begin
    FGrille.ColWidths[COL_NATURE] := -1;
    FGrille.ColLengths[COL_NATURE] := -1;
    FGrille.ColWidths[COL_PIECE] := 20;
    FGrille.ColLengths[COL_PIECE] := 20;
    FGrille.ColWidths[COL_DATE] := -1;
    FGrille.ColLengths[COL_DATE] := -1;
    FGrille.ColWidths[COL_GENERAL] := -1;
    FGrille.ColLengths[COL_GENERAL] := -1;
    FGrille.ColWidths[COL_AUXILIAIRE] := -1;
    FGrille.ColLengths[COL_AUXILIAIRE] := -1;
    FGrille.ColWidths[COL_JOURNAL] := 20;
    FGrille.ColLengths[COL_JOURNAL] := 20;
    FGrille.ColWidths[COL_LIBELLE] := -1;
    FGrille.ColLengths[COL_LIBELLE] := -1;
    FGrille.ColWidths[COL_REFINTERNE] := -1;
    FGrille.ColLengths[COL_REFINTERNE] := -1;
    FGrille.ColWidths[COL_DEBIT] := 20;
    FGrille.ColLengths[COL_DEBIT] := 20;
    FGrille.ColWidths[COL_CREDIT] := 20;
    FGrille.ColLengths[COL_CREDIT] := 20;
    FGrille.ColWidths[COL_SOLDE] := 20;
    FGrille.ColLengths[COL_SOLDE] := 20;
    FGrille.ColWidths[COL_QUALIFPIECE] := 20;
    FGrille.ColLengths[COL_QUALIFPIECE] := 20;
  end;
  if T = 5 then
  begin
    FGrille.Cells[COL_REFINTERNE, 0] := TraduireMemoire('Nom Paramsoc');
    FGrille.Cells[COL_LIBELLE, 0] := TraduireMemoire('Valeur');

    FGrille.ColWidths[COL_REFINTERNE] := 20;
    FGrille.ColLengths[COL_REFINTERNE] := 20;
    FGrille.ColWidths[COL_LIBELLE] := 20;
    FGrille.ColLengths[COL_LIBELLE] := 20;

    FGrille.ColWidths[COL_NATURE] := -1;
    FGrille.ColLengths[COL_NATURE] := -1;
    FGrille.ColWidths[COL_PIECE] := -1;
    FGrille.ColLengths[COL_PIECE] := -1;
    FGrille.ColWidths[COL_DATE] := -1;
    FGrille.ColLengths[COL_DATE] := -1;
    FGrille.ColWidths[COL_GENERAL] := -1;
    FGrille.ColLengths[COL_GENERAL] := -1;
    FGrille.ColWidths[COL_AUXILIAIRE] := -1;
    FGrille.ColLengths[COL_AUXILIAIRE] := -1;
    FGrille.ColWidths[COL_JOURNAL] := -1;
    FGrille.ColLengths[COL_JOURNAL] := -1;
    FGrille.ColWidths[COL_QUALIFPIECE] := -1;
    FGrille.ColLengths[COL_QUALIFPIECE] := -1;
    FGrille.ColWidths[COL_DEBIT] := -1;
    FGrille.ColLengths[COL_DEBIT] := -1;
    FGrille.ColWidths[COL_CREDIT] := -1;
    FGrille.ColLengths[COL_CREDIT] := -1;
    FGrille.ColWidths[COL_SELECTION] := -1;
    FGrille.ColLengths[COL_SELECTION] := -1;
    FGrille.ColWidths[COL_SOLDE] := -1;
    FGrille.ColLengths[COL_SOLDE] := -1;
  end;
end;

procedure TOF_ANOECR.OnArgument(S: string);
var
  Composants : TControlFiltre;
begin
  inherited;
  FTotalLigne := 0;
  FTotalDebit := 0;
  FTotalCredit := 0;
  FSelectionLigne := 0;
  FSelectionDebit := 0;
  FSelectionCredit := 0;
  Ecran.OnKeyDown := FormKeyDown;  
  Ecran.OnResize := OnEcranResize;
  FHSystemMenu := THSystemMenu(TFVierge(ECRAN).HMTrad);


  BExport := TToolBarButton97(GetControl('BExport', True));
  BExport.OnClick := BExportClick;
  BExport.Enabled := ExJaiLeDroitConcept(ccExportListe, False);

  BImprimer := TToolBarButton97(GetControl('BImprimer', True));
  BImprimer.OnClick := BImprimerClick;

  BValider := TToolBarButton97(GetControl('BValider', True));
  BValider.OnClick := BValiderClick;

  BSelectAll := TToolBarButton97(GetControl('BSELECTALL', true));
  BSelectAll.OnClick := OnToutSelectionnerClick;

  FJournal := THValComboBox(GetControl('E_JOURNAL', true)) ;
  FExercice := THValComboBox(GetControl('E_EXERCICE', true));
  THValComboBox(GetControl('E_EXERCICE')).OnChange := OnE_EXERCICEChange;


  SD := TSaveDialog.Create(Ecran);
  SD.Filter :=
    'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
  SD.DefaultExt := 'XLS';
  SD.FilterIndex := 1;
  SD.Options := SD.Options + [ofOverwritePrompt, ofPathMustExist,
    ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];

  FPages := TPageControl(GetControl('PAGES'));

  FGrille := THGrid(GetControl('FLISTE'));
  FGrille.OnFlipSelection := OnGridFlipSelection;
  FGrille.SynEnabled := False;
  FGrille.OnColumnWidthsChanged := OnChangeLargeurColonneGrille;
  FGrille.ColCount := COL_NOMBRE;
  FGrille.ColWidths[COL_PERIODE] := -1;
  FGrille.ColLengths[COL_PERIODE] := -1;
  FGrille.ColWidths[COL_LIGNE] := -1;
  FGrille.ColLengths[COL_LIGNE] := -1;
  FGrille.ColWidths[COL_SELECTION] := -1;
  FGrille.ColLengths[COL_SELECTION] := -1;

  FGrille.ColWidths[0] := 20;
  FGrille.ColWidths[COL_NATURE] := 50;
  FGrille.ColWidths[COL_PIECE] := 40;
  FGrille.ColWidths[COL_DATE] := 80;
  FGrille.ColWidths[COL_GENERAL] := 100;
  FGrille.ColWidths[COL_AUXILIAIRE] := 100;
  FGrille.ColWidths[COL_JOURNAL] := 100;
  FGrille.ColWidths[COL_LIBELLE] := 160;
  FGrille.ColWidths[COL_REFINTERNE] := 100;
  FGrille.ColWidths[COL_DEBIT] := 100;
  FGrille.ColWidths[COL_CREDIT] := 100;
  FGrille.ColWidths[COL_QUALIFPIECE] := 50;
  FGrille.ColAligns[COL_NATURE] := taCenter;
  FGrille.ColAligns[COL_PIECE] := taCenter;
  FGrille.ColAligns[COL_DATE] := taCenter;
  FGrille.ColAligns[COL_LIBELLE] := taLeftJustify;
  FGrille.ColAligns[COL_GENERAL] := taLeftJustify;
  FGrille.ColAligns[COL_AUXILIAIRE] := taLeftJustify;
  FGrille.ColAligns[COL_JOURNAL] := taLeftJustify;
  FGrille.ColAligns[COL_DEBIT] := taRightJustify;
  FGrille.ColAligns[COL_CREDIT] := taRightJustify;
  FGrille.ColAligns[COL_SOLDE] := taRightJustify;

  FGrille.ColAligns[COL_QUALIFPIECE] := taCenter;

  FGrille.ColFormats[COL_DEBIT] := '#,##0.00;#,##0.00; ;';
  FGrille.ColFormats[COL_CREDIT] := '#,##0.00;#,##0.00; ;';
  FGrille.ColFormats[COL_SOLDE] := '#,##0.00;#,##0.00; ;';

  FGrille.Cells[COL_NATURE, 0] := TraduireMemoire('Nature');
  FGrille.Cells[COL_PIECE, 0] := TraduireMemoire('N° Pièce');
  FGrille.Cells[COL_DATE, 0] := TraduireMemoire('Date');
  FGrille.Cells[COL_GENERAL, 0] := TraduireMemoire('Général');
  FGrille.Cells[COL_AUXILIAIRE, 0] := TraduireMemoire('Auxiliaire');
  FGrille.Cells[COL_JOURNAL, 0] := TraduireMemoire('Journal');
  FGrille.Cells[COL_LIBELLE, 0] := TraduireMemoire('Libellé');
  FGrille.Cells[COL_REFINTERNE, 0] := TraduireMemoire('Référence');
  FGrille.Cells[COL_DEBIT, 0] := TraduireMemoire('Débit');
  FGrille.Cells[COL_CREDIT, 0] := TraduireMemoire('Crédit');
  FGrille.Cells[COL_SOLDE, 0] := TraduireMemoire('Solde');
  FGrille.Cells[COL_QUALIFPIECE, 0] := TraduireMemoire('QualifPièce');

  FListeEcr := TOB.Create('', nil, -1);
  FListePiece := TOB.Create('', nil, -1);
  FRapport := TList.Create;
  {
  TToolBarButton97(GetControl('BZOOM')).OnClick := OnZoomEcritureClick;
  TToolBarButton97(GetControl('BSelectAll')).OnClick := OnToutSelectionnerClick;
  }
  TToolBarButton97(GetControl('BCHERCHE')).OnClick := OnChercheClick;
  THGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick;
  FGrille.SynEnabled := True;

  Composants.PopupF   := TPopUpMenu      (Getcontrol('POPF'));
  Composants.Filtres  := THValComboBox   (Getcontrol('FFILTRES'));
  Composants.Filtre   := TToolBarButton97(Getcontrol('BFILTRE'));
  Composants.PageCtrl := TPageControl    (Getcontrol('PAGECONTROL'));
  ObjFiltre := TObjFiltre.Create(Composants, 'CPANOECR');

  InitMultiCriteres ;

end;

procedure TOF_ANOECR.OnClose;
begin
  VideListe(FRapport);
  FRapport.Free;
  FListePiece.Free;
  FListeEcr.Free;
  if assigned(SD) then
    SD.Free;
  inherited;
  if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
end;

procedure TOF_ANOECR.FListeDblClick(Sender: TObject);
var
  P: RParFolio;
  z, zz: integer;
  Periode, Piece: integer;
  Journal, QualifPiece: string;
  DateComptable: TDateTime;
  Q: TQuery;
  St: string;
begin
  if not ((TCheckBox(GetControl('ANOLETT')).Checked = true) or
    (TCheckBox(GetControl('INCOPOI')).Checked = true)) then
    exit;
  if FListeEcr.Detail.Count = 0 then
    exit;
  Periode := StrToInt(FGrille.Cells[COL_PERIODE, FGrille.Row]);
  Piece := StrToInt(FGrille.Cells[COL_PIECE, FGrille.Row]);
  Journal := FGrille.Cells[COL_JOURNAL, FGrille.Row];
  QualifPiece := FGrille.Cells[COL_QUALIFPIECE, FGrille.Row];
  DateComptable := StrToDate(FGrille.Cells[COL_DATE, FGrille.Row]);
  FModeSaisie := FGrille.Cells[COL_MODESAISIE, FGrille.Row];
  if FModeSaisie = '-' then // mode pièce
  begin
    St := 'SELECT E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE FROM ECRITURE '
      + ' WHERE E_DATECOMPTABLE="' + USDateTime(DateComptable) + '"'
      + ' AND E_JOURNAL="' + Journal + '" AND E_NUMEROPIECE=' + IntToStr(Piece)
      + ' AND E_NUMLIGNE=1 AND E_NUMECHE<=1'
      + ' AND E_QUALIFPIECE="' + QualifPiece + '"'
      + ' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, '
      + 'E_NUMECHE, E_QUALIFPIECE';
    ;
    Q := OpenSQL(St, True);
    TrouveEtLanceSaisie(Q, taConsult, QualifPiece);
    Ferme(Q);
  end
  else // mode bordereau ou libre
  begin
    FillChar(P, Sizeof(P), #0);
    z := Trunc(Periode / 100);
    zz := Periode - (z * 100);
    P.ParPeriode := DateToStr(DebutDeMois(EncodeDate(z, zz, 1)));
    P.ParCodeJal := Journal;
    P.ParNumFolio := IntToStr(Piece);
    P.ParNumLigne := StrToInt(FGrille.Cells[COL_LIGNE, FGrille.Row]);
    ChargeSaisieFolio(P, taConsult);
  end;
end;

procedure TOF_ANOECR.ChargeLesEcritures;
var
  Q: TQuery;
  St : string;
  r: integer;
  wJournal, wExercice : string ;

begin
  FGrille.SynEnabled := False;
  FListeEcr.ClearDetail;
  FGrille.VidePile(False);
  St := '';
  wJournal := '';
  wExercice:='';
  r := 0;
  if GetControlText('E_JOURNAL')<>'' then
    wJournal := ' E_JOURNAL="' + GetControlText('E_JOURNAL') + '" ' ;
  if GetControlText('E_EXERCICE')<>'' then
    wExercice := ' E_EXERCICE="' + GetControlText('E_EXERCICE') + '" ' ;

  if TCheckBox(GetControl('ANOPIE')).Checked then
  begin
    St := 'Select distinct e_qualifpiece from ecriture';
    r := 1;
  end
  else if TCheckBox(GetControl('ANOLETT')).Checked then
  begin
    r := 2;
    St := 'Select E_NATUREPIECE,E_DATECOMPTABLE,E_NUMEROPIECE,E_GENERAL,E_AUXILIAIRE,E_JOURNAL,E_REFINTERNE,E_LIBELLE,E_DEBIT,E_CREDIT,E_PERIODE,E_NUMLIGNE,E_QUALIFPIECE,E_MODESAISIE from ecriture';
  end
  else if TCheckBox(GetControl('ANOBAL')).Checked then
  begin
    R := 3;
    St := 'select e_qualifpiece,e_journal, e_numeropiece,sum(e_debit) as debit,';
    St := St + 'sum(e_credit) as credit , sum(e_debit)-sum(e_credit) as solde from ecriture ';
  end
  else if TCheckBox(GetControl('INCOPOI')).Checked then
  begin
    r := 4;
    St := 'Select E_NATUREPIECE,E_DATECOMPTABLE,E_NUMEROPIECE,E_GENERAL,E_AUXILIAIRE,E_JOURNAL,E_REFINTERNE,E_LIBELLE,E_DEBIT,E_CREDIT,E_PERIODE,E_NUMLIGNE,E_QUALIFPIECE,E_MODESAISIE ';
    St := St + 'from ecriture' ;

  end
  else if TCheckBox(GetControl('EXOV8')).Checked then
  begin
    r := 5;
    St :=
      'Select Soc_nom, soc_data from paramsoc where soc_nom like "SO_EXOV8"';
  end;

  if r<> 5 then
  begin
    st := st + ' where ' ;
    // Critère exercice
    if length(trim(wExercice))<> 0 then
      St := St + wExercice + 'AND ';
    // Critère journal
    if length(trim(wJournal))<> 0 then
      St := St + wJournal + 'AND ' ;
    // Critère date comtpable
    St := St + ' E_DATECOMPTABLE>="' + USDate(THCritMaskEdit(GetControl('E_DATECOMPTABLE'))) + '" ';
    St := St + ' AND E_DATECOMPTABLE<="' + USDate(THCritMaskEdit(GetControl('E_DATECOMPTABLE_'))) + '" ';
    if r=2 then
    begin
      St := st + ' AND e_tresolettre="X"' ;
      St := St + ' ORDER BY E_NUMEROPIECE';
    end
    else
    if r=4 then
    begin
      St := st + ' AND e_refpointage<>"" and e_datepointage="01/01/1900"' ;
      St := St + ' ORDER BY E_NUMEROPIECE';
    end
    else
    if r=3 then St := St + ' group by e_qualifpiece,e_journal, e_numeropiece having sum(e_debit)-sum(e_credit)<>0 ORDER BY E_NUMEROPIECE';

  end ;


  ReFormeListe(r);
  if St <> '' then
  begin
    Q := OpenSQL(St, True);
    if r = 5 then
      FListeEcr.LoadDetailDB('PARAMSOC1', '', '', Q, False)
    else
      FListeEcr.LoadDetailDB('ECRITURE1', '', '', Q, False);
    Ferme(Q);
    if r = 1 then
      FListeEcr.PutGridDetail(THGrid(GetControl('FLISTE')), False, False,
        ';;;;;;;;;;;;;;E_QUALIFPIECE')
    else if (r = 2) or (r = 4) then
    begin
      FListeEcr.PutGridDetail(THGrid(GetControl('FLISTE')), False, False,
        'E_NATUREPIECE;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_JOURNAL;E_REFINTERNE;E_LIBELLE;E_DEBIT;E_CREDIT;SOLDE;E_PERIODE;E_NUMLIGNE;;E_QUALIFPIECE;E_MODESAISIE');
    end
    else if r = 3 then
      FListeEcr.PutGridDetail(THGrid(GetControl('FLISTE')), False, False,
        'E_NATUREPIECE;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_JOURNAL;E_REFINTERNE;E_LIBELLE;DEBIT;CREDIT;SOLDE;E_PERIODE;E_NUMLIGNE;;E_QUALIFPIECE')
    else if r = 5 then
      FListeEcr.PutGridDetail(THGrid(GetControl('FLISTE')), False, False,
        ';;;;;;SOC_NOM;SOC_DATA');

    //if FListeEcr.Detail.Count > 0 then
    //  FListeEcr.Detail[0].AddChampSupValeur('SELECTION', '-', True);
    // Redimensionnement des colonnes
    TFVierge(Ecran).FormResize := True;
    FHSystemMenu.ResizeGridColumns(FGrille);
    // Recalcul des totaux
    FTotalLigne := FListeEcr.Detail.Count;
    FGrille.SynEnabled := True;
  end;
end;

procedure TOF_ANOECR.OnChercheClick(Sender: TOBject);
begin
  ChargeLesEcritures;
end;

procedure TOF_ANOECR.InitMultiCriteres;
begin
  if VH^.CPExoRef.Code <> '' then
  begin
    SetControlText('E_EXERCICE', VH^.CPExoRef.Code);
    // Préselection dernier mois exercice
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.CPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.CPExoRef.Fin));
  end
  else
  begin
    SetControlText('E_EXERCICE', VH^.Encours.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.Encours.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.Encours.Fin));
  end;
  // On se positionne sur le premier journal
  THValComboBox(GetControl('E_JOURNAL')).ItemIndex := 0;
end;


{******************************************************************************}
{ * Gestion des filtres                                                       *}
{******************************************************************************}



procedure TOF_ANOECR.OnEcranResize(Sender: TObject);
begin
  inherited;
end;

procedure TOF_ANOECR.OnChangeLargeurColonneGrille(Sender: TObject);
begin
  inherited;
end;

initialization
  registerclasses([TOF_ANOECR]);
end.

