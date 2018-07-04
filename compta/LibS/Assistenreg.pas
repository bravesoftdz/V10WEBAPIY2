unit Assistenreg;

interface

uses
  Windows, Messages, Graphics, Controls, Forms, Dialogs,
  dbtables, classes, sysutils, UTOF, HMsgBox, HTB97, ParamSoc, Grids, HCtrls,
  UTob, AGLInitPlus, AssistPL, HEnt1, assist, StdCtrls, Spin, ComCtrls,
  HSysMenu, ExtCtrls,
  DBGrids, Buttons, Ent1, General, HStatus, RapSuppr, Lookup,
  Tablette, FE_Main, EdtEtat;

type
  TFASSISTENREG = class(TFAssist)
    TabSheet1: TTabSheet;
    NUMSTD: TSpinEdit;
    LabelNUMSTD: TLabel;
    Labelibelle: TLabel;
    LIBELLESTD: TEdit;
    GD: THGrid;
    BTNUMSTD: TToolbarButton97;
    procedure FormCreate(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure BTNUMSTDClick(Sender: TObject);
    procedure NUMSTDChange(Sender: TObject);
  private
    procedure DessineCell(ACol, ARow: Longint; Canvas: TCanvas; State:
      TGridDrawState);
    procedure SaveAsStandard;
    procedure SaveInfoStandard(NumStd: integer; LibelleStd: string);
    function ExistStandardCompta(NumStd: integer; TableStd: string): Boolean;
    procedure SaveInfoListe(NumStd: integer; LIST: string);

    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FASSISTENREG: TFASSISTENREG;
function LanceAssistantEnreg: boolean;

implementation
//uses TOMStdcpta;
uses uTOFStdChoixPlan, uLibStdCpta;

{$R *.DFM}

function LanceAssistantEnreg: boolean;
var
  bRetour: boolean;
begin
  bRetour := false;

  FASSISTENREG := TFASSISTENREG.Create(Application);
  try
    FASSISTENREG.ShowModal;
  finally
    FASSISTENREG.Free;
  end;
end;

procedure TFASSISTENREG.DessineCell(ACol, ARow: Longint; Canvas: TCanvas; State:
  TGridDrawState);
var
  Texte: string;
  Rect: TRect;
begin
  Rect := Gd.CellRect(ACol, ARow);
  // traitements de la barre de titre
  if ARow = 0 then
  begin
    case ACol of
      //      0: Texte := 'No dossier';
      0: Texte := 'Libellé';
      1: Texte := 'Marqué';
    end;
    with Gd.Canvas do
      // centrage, comme pour la case à cocher
      TextOut((Rect.Left + Rect.Right) div 2 - TextWidth(Texte) div 2,
        (Rect.Top + Rect.Bottom) div 2 - TextHeight(Texte) div 2, Texte);
    exit;
  end;
  // si autre que la colonne des coches : on ne fait rien
  if (ACol <> 1) then
    exit;
  // case à cocher (marquage)
  if Gd.Cells[ACol, Arow] = 'X' then
    Texte := 'þ' // coche
  else
    Texte := '¨'; // case vide
  with Gd.Canvas do
  begin
    FillRect(Rect); // redessine
    Font.Name := 'Wingdings';
    Font.Size := 10;
    // centrage de la coche
    TextOut((Rect.Left + Rect.Right) div 2 - TextWidth(Texte) div 2,
      (Rect.Top + Rect.Bottom) div 2 - TextHeight(Texte) div 2, Texte);
  end;
end;

procedure TFASSISTENREG.FormCreate(Sender: TObject);
var
  libelle: string;
  Tb, TF1, TF2, TF3, TF4, TF5, TF6, TF7: Tob;
  TF8, TF9, TF10, TF11, TF12, TF13: Tob;
  i: integer;
  Q1: TQuery;
begin
  inherited;

  if FileExists('STDCEGID.DAT') then
    NUMSTD.MinValue := 1;

  if NumPlanCompte <> 0 then
  begin

    if (NUMSTD.MinValue = 21) and (NumPlanCompte < 21) then
    begin
      Q1 := OpenSQL('SELECT max(STC_NUMPLAN) as maxno FROM STDCPTA', True);
      if not Q1.EOF then
        NUMSTD.value := Q1.Fields[0].AsInteger + 1;
    end
    else
    begin
      Q1 := OpenSql('SELECT STC_LIBELLE FROM STDCPTA WHERE STC_NUMPLAN=' +
        IntToStr(NumPlanCompte), TRUE);
      if not Q1.EOF then
        LIBELLESTD.Text := Q1.Fields[0].asstring;
      NUMSTD.value := NumPlanCompte;
    end;
    Ferme(Q1);
  end;
  Tb := TOB.Create('Enregistrement', nil, -1);
  TF1 := TOB.Create('fille1', Tb, -1);

  // ajoute une colonne de case à cocher (pour chaque enreg)

  TF1.AddChampSupValeur('Libelle', 'Standard');
  TF1.AddChampSupValeur('Coche', ' ');

  TF2 := TOB.Create('fille2', Tb, -1);
  TF2.AddChampSupValeur('Libelle', 'Généraux');
  TF2.AddChampSupValeur('Coche', ' ');

  TF3 := TOB.Create('fille3', Tb, -1);
  TF3.AddChampSupValeur('Libelle', 'Tiers');
  TF3.AddChampSupValeur('Coche', ' ');

  TF4 := TOB.Create('fille3', Tb, -1);
  TF4.AddChampSupValeur('Libelle', 'Journaux');
  TF4.AddChampSupValeur('Coche', ' ');

  TF5 := TOB.Create('fille3', Tb, -1);
  TF5.AddChampSupValeur('Libelle', 'Guides');
  TF5.AddChampSupValeur('Coche', ' ');

  TF6 := TOB.Create('fille3', Tb, -1);
  TF6.AddChampSupValeur('Libelle', 'Analytique');
  TF6.AddChampSupValeur('Coche', ' ');

  TF7 := TOB.Create('fille3', Tb, -1);
  TF7.AddChampSupValeur('Libelle', 'Libellés Standards');
  TF7.AddChampSupValeur('Coche', ' ');

  TF8 := TOB.Create('fille3', Tb, -1);
  TF8.AddChampSupValeur('Libelle', 'Correspondants');
  TF8.AddChampSupValeur('Coche', ' ');

  TF9 := TOB.Create('fille3', Tb, -1);
  TF9.AddChampSupValeur('Libelle', 'Ruptures');
  TF9.AddChampSupValeur('Coche', ' ');

  TF10 := TOB.Create('fille3', Tb, -1);
  TF10.AddChampSupValeur('Libelle', 'Libellés Standards');
  TF10.AddChampSupValeur('Coche', ' ');

  TF11 := TOB.Create('fille3', Tb, -1);
  TF11.AddChampSupValeur('Libelle', 'Axes analytiques');
  TF11.AddChampSupValeur('Coche', ' ');

  TF12 := TOB.Create('fille3', Tb, -1);
  TF12.AddChampSupValeur('Libelle', 'Sections analytiques');
  TF12.AddChampSupValeur('Coche', ' ');

  TF13 := TOB.Create('fille3', Tb, -1);
  TF13.AddChampSupValeur('Libelle', 'Ventilations types');
  TF13.AddChampSupValeur('Coche', ' ');

  // affichage
  Tb.PutGridDetail(GD, False, False, 'Libelle;Coche');
  Tb.Free;

  // dessin des cellules
  GD.PostDrawCell := DessineCell;
end;

procedure TFASSISTENREG.SaveAsStandard;
var
  NStd: integer;
  libelle: string;
  Tb, TF1, TF2: Tob;
  i: integer;
begin

  NStd := NUMSTD.Value;
  // Enregistrement des paramètres dossier
//  SaveParamAsStandardCompta (NStd);
  EnregistreParamSocStandard(NStd);
  GD.Cells[1, 1] := 'X';
  GD.PostDrawCell := DessineCell;

  // Enregistrement du plan comptable
  SaveAsStandardCompta(NStd, 'GENERAUX', 'GENERAUXREF');
  GD.Cells[1, 2] := 'X';
  GD.PostDrawCell := DessineCell;

  // Enregistrement des tiers
  SaveAsStandardCompta(NStd, 'TIERS', 'TIERSREF');
  GD.Cells[1, 3] := 'X';
  GD.PostDrawCell := DessineCell;

  // Enregistrement des journaux
  SaveAsStandardCompta(NStd, 'JOURNAL', 'JALREF');
  GD.Cells[1, 4] := 'X';
  GD.PostDrawCell := DessineCell;

  // Enregistrement des guides
  SaveAsStandardCompta(NStd, 'GUIDE', 'GUIDEREF');
  SaveAsStandardCompta(NStd, 'ECRGUI', 'ECRGUIREF');
  SaveAsStandardCompta(NStd, 'FILTRES', 'FILTRESREF');

  // CA - 02/04/2002
  SaveAsStandardCompta(NStd, 'MODEPAIE', 'MODEPAIEREF');
  SaveAsStandardCompta(NStd, 'MODEREGL', 'MODEREGLREF');
  // Fin CA - 02/04/2002

  SaveInfoListe(NStd, 'MULMMVTS');
  SaveInfoListe(NStd, 'MULMANAL');
  SaveInfoListe(NStd, 'MULVMVTS');
  GD.Cells[1, 5] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveAsStandardCompta(NStd, 'ANAGUI', 'ANAGUIREF');
  GD.Cells[1, 6] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveInfoStandard(NStd, LIBELLESTD.Text);
  GD.Cells[1, 7] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveAsStandardCompta(NStd, 'CORRESP', 'CORRESPREF');
  GD.Cells[1, 8] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveAsStandardCompta(NStd, 'RUPTURE', 'RUPTUREREF');
  SaveAsStandardCompta(NStd, 'NATCPTE', 'NATCPTEREF');
  GD.Cells[1, 9] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveAsStandardCompta(NStd, 'REFAUTO', 'REFAUTOREF');
  GD.Cells[1, 10] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveAsStandardCompta(NStd, 'AXE', 'AXEREF');
  GD.Cells[1, 11] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveAsStandardCompta(NStd, 'SECTION', 'SECTIONREF');
  GD.Cells[1, 12] := 'X';
  GD.PostDrawCell := DessineCell;

  SaveAsStandardCompta(NStd, 'CHOIXCOD', 'CHOIXCODREF');
  SaveAsStandardCompta(NStd, 'VENTIL', 'VENTILREF');
  GD.Cells[1, 13] := 'X';
  GD.PostDrawCell := DessineCell;

  SetParamSoc('SO_NUMPLANREF', NStd);
  NumPlanCompte := NUMSTD.value;
end;

procedure TFASSISTENREG.SaveInfoStandard(NumStd: integer; LibelleStd: string);
var
  QPlanRef: TQuery;
begin
  QPlanRef := OpenSQL('SELECT * FROM STDCPTA WHERE STC_NUMPLAN=' +
    IntToStr(NumStd), False);
  if QPlanRef.Eof then
  begin
    QPlanRef.Insert;
    InitNew(QPlanRef);
    QPlanRef.FindField('STC_NUMPLAN').AsInteger := NumStd;
    QPlanRef.FindField('STC_LIBELLE').AsString := LibelleStd;
    QPlanRef.FindField('STC_ABREGE').AsString := Copy(LibelleStd, 1, 17);
    if NumStd < 21 then
      QPlanRef.FindField('STC_PREDEFINI').AsString := 'CEG'
    else
      QPlanRef.FindField('STC_PREDEFINI').AsString := 'STD';
    QPlanRef.Post;
  end;
  Ferme(QPlanRef);
end;

procedure TFASSISTENREG.SaveInfoListe(NumStd: integer; LIST: string);
var
  Q1, QPlanRef: TQuery;
begin
  Q1 := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="' + LIST + '" '
    + 'AND LI_UTILISATEUR="---"', FALSE);

  if not Q1.Eof then
  begin
    QPlanRef := OpenSQL('SELECT * FROM LISTEREF WHERE LIR_NUMPLAN=' +
      IntToStr(NumStd) + ' AND LIR_LISTE="' + LIST + '" '
      + 'AND LIR_UTILISATEUR="---"', FALSE);

    if QPlanRef.Eof then
    begin
      QPlanRef.Insert;
      InitNew(QPlanRef);
      QPlanRef.FindField('LIR_NUMPLAN').AsInteger := NumStd;
      QPlanRef.FindField('LIR_LISTE').AsString := LIST;
      //            QPlanRef.FindField('LIR_UTILISATEUR').AsString := Q1.findField ('LI_UTILISATEUR').asstring;
      QPlanRef.FindField('LIR_UTILISATEUR').AsString := '---';
      QPlanRef.FindField('LIR_LIBELLE').AsString :=
        Q1.findField('LI_LIBELLE').asstring;
      QPlanRef.FindField('LIR_SOCIETE').AsString :=
        Q1.findField('LI_SOCIETE').asstring;
      QPlanRef.FindField('LIR_NUMOK').AsVariant :=
        Q1.findField('LI_NUMOK').AsVariant;
      QPlanRef.FindField('LIR_TRIOK').AsVariant :=
        Q1.findField('LI_TRIOK').AsVariant;
      QPlanRef.FindField('LIR_LANGUE').AsString :=
        Q1.findField('LI_LANGUE').asstring;
      QPlanRef.FindField('LIR_DATA').AsVariant :=
        Q1.findField('LI_DATA').AsVariant;
      if NumStd < 21 then
        QPlanRef.FindField('LIR_PREDEFINI').Asstring := 'CEG'
      else
        QPlanRef.FindField('LIR_PREDEFINI').Asstring := 'STD';
      QPlanRef.Post;
    end
    else
    begin
      QPlanRef.Edit;
      QPlanRef.FindField('LIR_DATA').Asstring :=
        Q1.findField('LI_DATA').Asstring;
      QPlanRef.Post;
    end;
    Ferme(QPlanRef);
  end;
  Ferme(Q1);
end;

procedure TFASSISTENREG.bFinClick(Sender: TObject);
var
  Texte: string;
begin
  inherited;
  if LIBELLESTD.Text = '' then
  begin
    PGIBox('Renseignez le libellé', 'Libellé du standard');
    exit;
  end;
  if ExistStandardCompta(NUMSTD.Value, 'STDCPTA') then
  begin
    Texte := 'Le plan existe déjà voulez-vous le remplacer';
    if HShowMessage('0;Chargement du dossier type;' + Texte + ';Q;YN;N;N;', '',
      '') <> mrYes then
      exit;
    SupprimeDonneStd('PARSOCREF', NUMSTD.Value);
    SupprimeDonneStd('GENERAUXREF', NUMSTD.Value);
    SupprimeDonneStd('JALREF', NUMSTD.Value);
    SupprimeDonneStd('GUIDEREF', NUMSTD.Value);
    SupprimeDonneStd('ECRGUIREF', NUMSTD.Value);
    SupprimeDonneStd('ANAGUIREF', NUMSTD.Value);
    SupprimeDonneStd('TIERSREF', NUMSTD.Value);
    SupprimeDonneStd('CORRESPREF', NUMSTD.Value);
    SupprimeDonneStd('RUPTUREREF', NUMSTD.Value);
    SupprimeDonneStd('REFAUTOREF', NUMSTD.Value);
    SupprimeDonneStd('AXEREF', NUMSTD.Value);
    SupprimeDonneStd('SECTIONREF', NUMSTD.Value);
    SupprimeDonneStd('CHOIXCODREF', NUMSTD.Value);
    SupprimeDonneStd('VENTILREF', NUMSTD.Value);
    SupprimeDonneStd('NATCPTEREF', NUMSTD.Value);
    SupprimeDonneStd('FILTRESREF', NUMSTD.Value);
    SupprimeDonneStd('LISTEREF', NUMSTD.Value);
    ExecuteSQL('UPDATE  STDCPTA SET STC_LIBELLE="' + LIBELLESTD.Text +
      '" where STC_NUMPLAN=' + IntTostr(NUMSTD.Value));
    close;
  end;
  SaveAsStandard;
  ModalResult := MROK;
end;

function TFASSISTENREG.ExistStandardCompta(NumStd: integer; TableStd: string):
  Boolean;
var
  PrefStd: string;
  OK: Boolean;
  QStd: TQuery;
begin
  OK := TRUE;
  PrefStd := TableToPrefixe(TableStd);
  QStd := OpenSQL('SELECT * FROM ' + TableStd + ' WHERE ' + PrefStd + '_NUMPLAN='
    + IntToStr(NumStd), False);
  if QStd.Eof then
    OK := FALSE;
  Ferme(QStd);
  Result := OK;
end;

procedure TFASSISTENREG.BTNUMSTDClick(Sender: TObject);
begin
  inherited;
  LookUpList(TControl(Sender), 'Liste des dossiers types', 'STDCPTA',
    'STC_NUMPLAN', 'STC_LIBELLE', '', 'STC_NUMPLAN', True, 0);
end;

procedure TFASSISTENREG.NUMSTDChange(Sender: TObject);
var
  Q1: TQuery;
begin
  inherited;
  Q1 := OpenSql('SELECT STC_LIBELLE FROM STDCPTA WHERE STC_NUMPLAN=' +
    IntToStr(NUMSTD.Value), TRUE);
  if not Q1.EOF then
    LIBELLESTD.Text := Q1.Fields[0].asstring;
  ferme(Q1);
end;

end.

