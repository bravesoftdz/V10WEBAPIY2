{***********UNITE*************************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Gestion des tarifs de référence d'un article
Mots clefs ... : TARIF;ARTICLE
*****************************************************************}
unit UTofMBOTarifSaisiBase;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  {$IFNDEF EAGLCLIENT}
  DBTables, Fe_Main, HDB,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, vierge, UTOB, AglInit, LookUp, EntGC, SaisUtil, graphics
  , grids, windows, TarifUtil, M3FP, UtilArticle, Paramsoc,
  ent1;

type
  TOF_MBOTarifSaisiBase = class(TOF)
  private
    TarifBase: THGRID;
    TobBase, TobBaseLigne, TOBTarif, TobTarifDim, TOBTarifExistant: Tob;
    Col_Mov, colType, colDev, colPrix, colArrondi, colCreer: Integer;
    LesColonnes, CodeArticle, Etat: string;
    TarifCreer: Boolean;

    procedure GSElipsisClick(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ChargeTarifBase;
    procedure NouveauTarifBase;
    procedure MAJPrixArrondi(Coef: Double);
    procedure FormateColSaisie(ACol, ARow: Longint);
    procedure FormatePrix;

  public
    Action: TActionFiche;
    DEV: RDEVISE;
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;

    // MAJ Table
    procedure MAJTarif(Arow, i: Integer; CodeTarifMode: string);
    procedure CreerNouveauTarif(ARow, i, MaxTarif: Integer; Q: TQuery; Etab, CodeTarifMode: string);
    procedure MAJTarifDim;
    function MAJTarifMode(Arow: Integer): string;
    procedure ValideTarif;

    procedure ColTriangle(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure CodeGras(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
  end;

implementation

procedure TOF_MBOTarifSaisiBase.OnArgument(Arguments: string);
var St, S, NomCol, Critere, ChampMul, ValMul: string;
  x, i: Integer;
begin
  inherited;
  St := Arguments;
  repeat
    Critere := Trim(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'CodeArticle' then CodeArticle := ValMul;
        if ChampMul = 'Etat' then Etat := ValMul;
      end;
    end;
  until Critere = '';
  // DCA - FQ MODE 10785
  if Etat = 'Consult' then Action := taConsult
  else Action := taModif;
  Ecran.Caption := 'Tarif de base de l''article: ' + CodeArticle;
  LesColonnes := 'MOV;_TYPE;_DEV;_PRIX;_ARRONDI;_CREER';
  TarifBase := THGRID(GetControl('TARIFBASE'));
  TarifBase.OnCellEnter := GSCellEnter;
  TarifBase.OnCellExit := GSCellExit;
  TarifBase.OnRowExit := GSRowExit;
  TarifBase.OnRowEnter := GSRowEnter;
  TarifBase.OnEnter := GSEnter;
  TarifBase.OnElipsisClick := GSElipsisClick;
  TarifBase.OnDblClick := GSElipsisClick;
  //TarifBase.OnClick:=GCBoiteCoche ;
  TarifBase.GetCellCanvas := CodeGras;
  TarifBase.PostDrawCell := ColTriangle;
  TarifBase.ColCount := 1;
  i := 0;
  S := LesColonnes;
  Col_Mov := -1;
  colType := -1;
  colPrix := -1;
  colArrondi := -1;
  colCreer := -1;
  repeat
    NomCol := ReadTokenSt(S);
    if NomCol <> '' then
    begin
      if NomCol = 'MOV' then
      begin
        if i <> 0 then TarifBase.ColCount := TarifBase.ColCount + 1;
        Col_Mov := i;
        TarifBase.ColWidths[Col_Mov] := 10;
      end
      else
        if NomCol = '_TYPE' then
      begin
        if i <> 0 then TarifBase.ColCount := TarifBase.ColCount + 1;
        colType := i;
        TarifBase.ColWidths[colType] := 150;
        TarifBase.ColLengths[colType] := -1;
      end
      else
        if NomCol = '_DEV' then
      begin
        if i <> 0 then TarifBase.ColCount := TarifBase.ColCount + 1;
        colDev := i;
        TarifBase.ColWidths[colDev] := 50;
        TarifBase.ColLengths[colDev] := -1;
        TarifBase.ColAligns[colDev] := taCenter;
      end
      else
        if NomCol = '_PRIX' then
      begin
        if i <> 0 then TarifBase.ColCount := TarifBase.ColCount + 1;
        colPrix := i;
        TarifBase.ColWidths[colPrix] := 50;
        TarifBase.ColLengths[colPrix] := 30;
        TarifBase.ColAligns[colPrix] := taRightJustify;
      end
      else
        if NomCol = '_ARRONDI' then
      begin
        if i <> 0 then TarifBase.colCount := TarifBase.ColCount + 1;
        colArrondi := i;
        TarifBase.ColWidths[colArrondi] := 50;
        TarifBase.ColAligns[colArrondi] := taCenter;
      end
      else
        if NomCol = '_CREER' then
      begin
        if i <> 0 then TarifBase.colCount := TarifBase.ColCount + 1;
        colCreer := i;
        TarifBase.ColWidths[colCreer] := 100;
        TarifBase.ColFormats[colCreer] := 'CB=PGOUINON|''|';
        TarifBase.ColLengths[colCreer] := 1;
      end;
      Inc(i);
    end;
  until ((St = '') or (NomCol = ''));
  if Col_Mov <> -1 then TarifBase.FixedCols := 1;

  AffecteGrid(TarifBase, taModif);
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(TarifBase);
  TFVierge(Ecran).OnKeyDown := FormKeyDown;
end;

procedure TOF_MBOTarifSaisiBase.OnLoad;
begin
  TarifCreer := False;
  TOBTarifExistant := TOB.Create('_TARIF', nil, -1);
  if Etat = 'nouveau' then NouveauTarifBase else ChargeTarifBase;
  // Permet d'afficher en gras le code de la première ligne du Grid
  TarifBase.SetFocus;
  FormatePrix;
  TobTarif := TOB.Create('', nil, -1);
  TobTarifDim := TOB.Create('TARIF', nil, -1);
end;

procedure TOF_MBOTarifSaisiBase.OnUpdate;
var CodeTarifMode: string;
  i, j: Integer;
begin
  // DCA - FQ MODE 10785
  if Action = taConsult then exit;
  // Vérifier si tarif existant ou pas => modif ou création
  TOBBase.GetGridDetail(TarifBase, TarifBase.RowCount - 1, 'les Tarifs', LesColonnes);
  TarifCreer := True;
  for i := 1 to TarifBase.rowcount - 1 do
  begin
    if TarifBase.CellValues[ColCREER, i] = 'OUI' then
      //if TarifBase.isSelected(i) then
    begin
      if (TarifBase.Cells[ColPrix, i] = '') or (TarifBase.Cells[ColPrix, i] = '0') then continue;
      CodeTarifMode := MAJTarifMode(i);
      j := TobTarif.Detail.count;
      MAJTarif(i, j, CodeTarifMode);
      ValideTarif;
    end;
  end;
end;

procedure TOF_MBOTarifSaisiBase.OnClose;
begin
  TobBase.free;
  TobBase := nil;
  TOBTarif.free;
  TobTarif := nil;
  TOBTarifDim.free;
  TobTarifDim := nil;
  TOBTarifExistant.free;
  TOBTarifExistant := nil;
end;

{==============================================================================================}
{=============================== MAJ des tables ===============================================}
{==============================================================================================}

procedure TOF_MBOTarifSaisiBase.MAJTarif(Arow, i: Integer; CodeTarifMode: string);
var Q, Q2: TQuery;
  Etablissement: string;
  MaxTarif, j: Integer;
  TobF: TOB;
begin
  Q := OpenSQL('SELECT MAX(GF_TARIF) FROM TARIF', TRUE);
  if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1;
  Ferme(Q);
  Q2 := OpenSQL('Select * from TarifMode where GFM_TARFMODE="' + CodeTarifMode + '"', True);
  if not Q2.EOF then
  begin
    // Tarif toute boutique
    Etablissement := '';
    if TobTarifExistant.Detail.count = 0 then
    begin
      CreerNouveauTarif(ARow, i, MaxTarif, Q2, Etablissement, CodeTarifMode);
    end else
    begin
      TOBF := TOBTarifExistant.FindFirst(['GF_DEPOT', 'GF_TARFMODE'], [Etablissement, CodeTarifMode], False);
      if TOBF = nil then
        CreerNouveauTarif(ARow, TOBTarif.Detail.count, MaxTarif, Q2, Etablissement, CodeTarifMode)
      else
      begin
        j := TOBTarif.Detail.count;
        TOB.Create('TARIF', TobTarif, j);
        TobTarif.Detail[j].Dupliquer(TOBF, False, True);
        TobTarif.Detail[j].PutValue('GF_DEVISE', TobBase.Detail[Arow - 1].GetValue('_DEV'));
        TobTarif.Detail[j].PutValue('GF_PRIXUNITAIRE', TobBase.Detail[Arow - 1].GetValue('_PRIX'));
        TobTarif.Detail[j].Putvalue('GF_ARRONDI', TobBase.Detail[Arow - 1].GetValue('_ARRONDI'));
        CalcPriorite(TobTarif.Detail[j]);
      end;
    end;
  end;
  Ferme(Q2);
end;

procedure TOF_MBOTarifSaisiBase.CreerNouveauTarif(ARow, i, MaxTarif: Integer; Q: TQuery; Etab, CodeTarifMode: string);
var libelle: string;
begin
  TOB.Create('TARIF', TobTarif, i);
  TobTarif.Detail[i].PutValue('GF_TARIF', MaxTarif);
  TobTarif.Detail[i].PutValue('GF_TARFMODE', CodeTarifMode);
  TobTarif.Detail[i].PutValue('GF_ARTICLE', CodeArticle);
  TobTarif.Detail[i].PutValue('GF_DEVISE', TobBase.Detail[Arow - 1].GetValue('_DEV'));
  TobTarif.Detail[i].PutValue('GF_DATEDEBUT', Q.FindField('GFM_DATEDEBUT').AsDateTime);
  TobTarif.Detail[i].PutValue('GF_DATEFIN', Q.FindField('GFM_DATEFIN').AsDateTime);
  TobTarif.Detail[i].PutValue('GF_PRIXUNITAIRE', TobBase.Detail[Arow - 1].GetValue('_PRIX'));
  TobTarif.Detail[i].PutValue('GF_BORNEINF', -999999);
  TobTarif.Detail[i].PutValue('GF_BORNESUP', 999999);
  TobTarif.Detail[i].PutValue('GF_QUANTITATIF', '-');
  TobTarif.Detail[i].PutValue('GF_REMISE', 0);
  TobTarif.Detail[i].Putvalue('GF_ARRONDI', TobBase.Detail[Arow - 1].GetValue('_ARRONDI'));
  TobTarif.Detail[i].PutValue('GF_DEMARQUE', Q.FindField('GFM_DEMARQUE').AsString);
  TobTarif.Detail[i].PutValue('GF_CALCULREMISE', '');
  TobTarif.Detail[i].PutValue('GF_MODECREATION', 'MAN');
  TobTarif.Detail[i].PutValue('GF_CASCADEREMISE', 'MIE');
  TobTarif.Detail[i].PutValue('GF_QUALIFPRIX', 'GRP');
  TobTarif.Detail[i].PutValue('GF_REGIMEPRIX', 'TTC');
  TobTarif.Detail[i].PutValue('GF_NATUREAUXI', 'CLI');
  libelle := RechDom('GCTARIFTYPE1VTE', Q.FindField('GFM_TYPETARIF').AsString, False) + ' - ' + RechDom('GCTARIFPERIODE1',
    Q.FindField('GFM_PERTARIF').AsString, False);
  TobTarif.Detail[i].PutValue('GF_LIBELLE', copy(libelle, 1, 35));
  TobTarif.Detail[i].PutValue('GF_DEPOT', Etab);
  TobTarif.Detail[i].PutValue('GF_SOCIETE', Etab);
  CalcPriorite(TOBTarif.Detail[i]);
end;

procedure TOF_MBOTarifSaisiBase.MAJTarifDim; // AC
var MaxTarif, j: Integer;
  QArtDim: TQuery;
begin
  j := 0;
  TobTarifDim := TOB.Create('', nil, -1);
  MaxTarif := TOBTarif.GetValue('GF_TARIF') + 1; // + TOBTarfArt.Detail.count;
  begin
    QArtDim := OpenSql('Select GA_ARTICLE from Article where GA_CODEARTICLE="' + TRIM(copy(CodeArticle, 1, 18)) + '" And GA_STATUTART="DIM"', True);
    while not QArtDim.EOF do
    begin
      TOB.Create('TARIF', TobTarifDim, j);
      TobTarifDim.Detail[j].Dupliquer(TOBTarif, False, True);
      TobTarifDim.Detail[j].PutValue('GF_ARTICLE', QArtDim.FindField('GA_ARTICLE').AsString);
      TobTarifDim.Detail[j].PutValue('GF_TARIF', MaxTarif);
      QArtDim.next;
      MaxTarif := MaxTarif + 1;
      j := j + 1;
      CalcPriorite(TobTarifDim.Detail[j]);
    end;
  end;
  ferme(QArtDim);
end;

function TOF_MBOTarifSaisiBase.MAJTarifMode(Arow: Integer): string;
var Q: TQuery;
  CodeMaxInt: Integer;
  TobType, TobPer, TobTarifMode: TOB;
  CodeMax, LibTarif, TypeTarif: string;
begin
  CodeMaxInt := 1;
  TobTarifMode := TOB.Create('TARIFMODE', nil, -1);
  TobType := TOB.Create('TARIFTYPMODE', nil, -1);
  TobPer := TOB.Create('TARIFPER', nil, -1);
  LibTarif := TarifBase.Cells[ColType, ARow];
  TypeTarif := ReadTokenPipe(LibTarif, '-');
  TOBType.SelectDB('"' + TypeTarif + '";"VTE"', nil);
  TobPer.SelectDB('"...";"' + GetParamSoc('SO_GCPERBASETARIF') + '"', nil);
  Q := OpenSql('Select * from TARIFMODE where GFM_TYPETARIF="' + TypeTarif + '"and GFM_NATURETYPE="VTE" and GFM_PERTARIF="' + GetParamSoc('SO_GCPERBASETARIF')
    +
    '"', True);
  if not Q.EOF then
  begin
    TobTarifMode.SelectDB('', Q);
    Result := TobTarifMode.GetValue('GFM_TARFMODE');
  end else
  begin
    Q := OpenSQL('SELECT MAX(GFM_TARFMODE) FROM TARIFMODE', True);
    if not Q.EOF then
    begin
      CodeMax := Q.Fields[0].AsString;
      if CodeMax <> '' then CodeMaxInt := StrToInt(CodeMax) + 1;
    end;
    Result := IntToStr(CodeMaxInt);
    TobTarifMode.PutValue('GFM_TARFMODE', IntToStr(CodeMaxInt));
    TobTarifMode.PutValue('GFM_TYPETARIF', TobType.GetValue('GFT_CODETYPE'));
    TobTarifMode.PutValue('GFM_NATURETYPE', TobType.GetValue('GFT_NATURETYPE'));
    TobTarifMode.PutValue('GFM_PERTARIF', TobPer.GetValue('GFP_CODEPERIODE'));
    TobTarifMode.PutValue('GFM_LIBELLE', Copy((TobType.GetValue('GFT_LIBELLE') + '-' + TobPer.GetValue('GFP_LIBELLE')), 1, 35));
    TobTarifMode.PutValue('GFM_DATEDEBUT', TobPer.GetValue('GFP_DATEDEBUT'));
    TobTarifMode.PutValue('GFM_DATEFIN', TobPer.GetValue('GFP_DATEFIN'));
    TobTarifMode.PutValue('GFM_DEMARQUE', TobPer.GetValue('GFP_DEMARQUE'));
    TobTarifMode.PutValue('GFM_ARRONDI', TobPer.GetValue('GFP_ARRONDI'));
    TobTarifMode.PutValue('GFM_PROMO', TobPer.GetValue('GFP_PROMO'));
    TobTarifMode.PutValue('GFM_COEF', TobType.GetValue('GFT_COEF'));
    TobTarifMode.PutValue('GFM_DEVISE', TobType.GetValue('GFT_DEVISE'));
    TobTarifMode.PutValue('GFM_ETABLISREF', TobType.GetValue('GFT_ETABLISREF'));
    TobTarifMode.InsertOrUpdateDBTable;
  end;
  TobType.Free;
  TobPer.Free;
  TobTarifMode.Free;
  Ferme(Q);
end;

procedure TOF_MBOTarifSaisiBase.ValideTarif;
begin
  TOBTarif.InsertOrUpdateDBTable(False);
  if LaTob.GetValue('GA_DIMMASQUE') <> '' then TOBTarifDim.InsertOrUpdateDBTable(False);
end;

{==============================================================================================}
{=============================== Evenement du Grid ========================================}
{==============================================================================================}

procedure TOF_MBOTarifSaisiBase.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  TarifBase.ElipsisButton := (TarifBase.Col = colArrondi);
  //If TarifBase.Col=colCreer then BoiteCoche(Sender) ;
end;

procedure TOF_MBOTarifSaisiBase.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var ValArrondi: string;
  Prix: Double;
begin
  if (ACol = ColPrix) or (ACol = ColArrondi) then TarifBase.CellValues[ColCREER, ARow] := 'OUI';
  Prix := Valeur(TarifBase.Cells[ColPrix, Arow]);
  ValArrondi := TarifBase.CellValues[ColArrondi, Arow];
  if ACol = ColPrix then TarifBase.Cells[ACol, Arow] := FloatToStr(ArrondirPrix(ValArrondi, Prix));
  if ACol = ColArrondi then TarifBase.Cells[ColPrix, Arow] := FloatToStr(ArrondirPrix(ValArrondi, Prix));
  DEV.Code := TOBBase.Detail[ARow - 1].GetValue('_DEV');
  GetInfosDevise(DEV);
  FormateColSaisie(Acol, Arow);
end;

procedure TOF_MBOTarifSaisiBase.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  TarifBase.InvalidateRow(Ou);
end;

procedure TOF_MBOTarifSaisiBase.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  TarifBase.InvalidateRow(Ou);
  //BoiteCoche(Sender) ;
end;

procedure TOF_MBOTarifSaisiBase.GSEnter(Sender: TObject);
var ACol, ARow: integer;
  Temp: Boolean;
begin
  if Action = taConsult then Exit;
  ACol := TarifBase.Col;
  ARow := TarifBase.Row;
  //CreerBoite (Sender) ;
  TarifBase.InvalidateRow(ARow);
  GSCellEnter(TarifBase, ACol, ARow, Temp);
end;

procedure TOF_MBOTarifSaisiBase.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
end;

procedure TOF_MBOTarifSaisiBase.GSElipsisClick(Sender: TObject);
var ARRONDI: THCritMaskEdit;
  Coord: TRect;
begin
  inherited;
  if (TarifBase.Col = colArrondi) then
  begin
    Coord := TarifBase.CellRect(TarifBase.Col, TarifBase.Row);
    ARRONDI := THCritMaskEdit.Create(TarifBase);
    ARRONDI.Parent := TarifBase;
    ARRONDI.Top := Coord.Top;
    ARRONDI.left := Coord.Left;
    ARRONDI.Width := 3;
    ARRONDI.Visible := False;
    ARRONDI.OpeType := otString;
    ARRONDI.DATATYPE := 'GCCODEARRONDI';
    LookUpCombo(ARRONDI);
    if ARRONDI.Text <> '' then TarifBase.Cells[TarifBase.Col, TarifBase.Row] := ARRONDI.Text;
    ARRONDI.Destroy;
  end;
end;

{==============================================================================================}
{=============================== Chargement des Tarifs de base ================================}
{==============================================================================================}

procedure TOF_MBOTarifSaisiBase.NouveauTarifBase;
var QType: TQuery;
  i: Integer;
  Exist: Boolean;
begin
  Exist := False;
  if TOBBase = nil then TOBBase := TOB.Create('les Tarifs', nil, -1);
  QType := OpenSql('Select * from TARIFTYPMODE order by GFT_CODETYPE', True);
  if not QType.EOF then
  begin
    while not QType.EOF do
    begin
      for i := 0 to TOBBase.Detail.Count - 1 do
      begin
        if TOBBase.Detail[i].GetValue('_TYPE') = QType.Findfield('GFT_CODETYPE').AsString + '-' + QType.Findfield('GFT_LIBELLE').AsString then
        begin
          Exist := True;
        end;
      end;
      if not Exist then
      begin
        if QType.Findfield('GFT_CODETYPE').AsString <> '...' then
        begin
          TOBBaseLigne := Tob.create('un tarif', TOBBase, -1);
          TOBBaseLigne.AddChampSup('MOV', False);
          TOBBaseLigne.AddChampSup('_TYPE', False);
          TOBBaseLigne.PutValue('_TYPE', QType.Findfield('GFT_CODETYPE').AsString + '-' + QType.Findfield('GFT_LIBELLE').AsString);
          TOBBaseLigne.AddChampSup('_DEV', False);
          TOBBaseLigne.PutValue('_DEV', QType.Findfield('GFT_DEVISE').AsString);
          DEV.Code := QType.Findfield('GFT_DEVISE').AsString;
          GetInfosDevise(DEV);
          TOBBaseLigne.AddChampSup('_CREER', False);
          TOBBaseLigne.PutValue('_CREER', 'OUI');
          // Recup des prix et arrondi
          TOBBaseLigne.AddChampSup('_PRIX', False);
          TOBBaseLigne.AddChampSup('_ARRONDI', False);
          MAJPrixArrondi(QType.FindField('GFT_COEF').AsFloat);
          TarifBase.RowCount := TarifBase.RowCount + 1;
        end;
      end;
      QType.next;
      Exist := False;
    end;
  end;
  ferme(QType);
  if TOBBase.Detail.Count > 0 then TOBBase.PutGridDetail(TarifBase, False, False, LesColonnes, True);
end;

procedure TOF_MBOTarifSaisiBase.ChargeTarifBase;
var QTarifMode, QTarif, QtarifExist: TQuery;
  SQL: string;
begin
  QTarifExist := OpenSQL('Select * from TARIF Where GF_ARTICLE="' + LaTob.GetValue('GA_ARTICLE') + '" order by GF_TARFMODE,GF_DEPOT', True);
  if not QTarifExist.EOF then TOBTarifExistant.LoadDetailDB('TARIF', '', '', QTarifExist, false, true);
  Ferme(QTarifExist);
  SQL := 'Select DISTINCT GF_TARFMODE, GF_PRIXUNITAIRE, GF_REMISE, GF_ARRONDI, GF_DEVISE from TARIF Where GF_ARTICLE="' + LaTob.GetValue('GA_ARTICLE') + '"';
  if ExisteSQL(SQL) then
  begin
    TOBBase := TOB.Create('les Tarifs', nil, -1);
    QTarif := OpenSQL(SQL, True);
    if not QTarif.EOF then
    begin
      while not QTarif.EOF do
      begin
        if QTarif.FindField('GF_TARFMODE').AsString <> '0' then
        begin
          QTarifMode := OpenSql('Select * from TARIFMODE where GFM_TARFMODE=' + QTarif.FindField('GF_TARFMODE').AsString, True);
          if not QTarifMode.EOF then
          begin
            if QTarifMode.FindField('GFM_PERTARIF').AsString = GetParamSoc('SO_GCPERBASETARIF') then
            begin
              if (QTarifMode.Findfield('GFM_TYPETARIF').AsString <> '...') then
              begin
                TOBBaseLigne := Tob.create('un tarif', TOBBase, -1);
                TOBBaseLigne.AddChampSup('MOV', False);
                TOBBaseLigne.AddChampSup('_TYPE', False);
                TOBBaseLigne.PutValue('_TYPE', QTarifMode.Findfield('GFM_TYPETARIF').AsString + '-' + RechDom('GCTARIFTYPE1VTE',
                  QTarifMode.Findfield('GFM_TYPETARIF').AsString, False));
                TOBBaseLigne.AddChampSup('_DEV', False);
                TOBBaseLigne.PutValue('_DEV', QTarif.Findfield('GF_DEVISE').AsString);
                DEV.Code := QTarif.Findfield('GF_DEVISE').AsString;
                GetInfosDevise(DEV);
                TOBBaseLigne.AddChampSup('_CREER', False);
                TOBBaseLigne.PutValue('_CREER', 'NON');
                TOBBaseLigne.AddChampSup('_PRIX', False);
                TOBBaseLigne.AddChampSup('_ARRONDI', False);
                TOBBaseLigne.PutValue('_PRIX', QTarif.Findfield('GF_PRIXUNITAIRE').AsFloat);
                TOBBaseLigne.PutValue('_ARRONDI', QTarif.Findfield('GF_ARRONDI').AsString);
                TarifBase.RowCount := TarifBase.RowCount + 1;
              end;
            end;
          end;
          ferme(QTarifMode);
          QTarif.Next;
        end else
          QTarif.Next;
      end;
      ferme(QTarif);
    end;
    if TOBBase.Detail.Count > 0 then TOBBase.PutGridDetail(TarifBase, False, False, LesColonnes, True);
  end;
  NouveauTarifBase;
end;

procedure TOF_MBOTarifSaisiBase.MAJPrixArrondi(Coef: Double);
var PrixArticle, Prix: Double;
  ValArrondi: string;
begin
  if LaTob <> nil then
  begin
    CodeArticle := LaTob.GetValue('GA_ARTICLE');
    ValArrondi := LaTob.GetValue('GA_ARRONDIPRIXTTC');
    PrixArticle := LaTob.GetValue('GA_PVTTC');
    if Coef <> 0 then Prix := PrixArticle * Coef else Prix := PrixArticle;
    Prix := ArrondirPrix(ValArrondi, Prix);
    TOBBaseLigne.PutValue('_PRIX', Prix);
    TOBBaseLigne.PutValue('_ARRONDI', ValArrondi);
  end;
end;
{==============================================================================================}
{=============================== Actions liées au grid ========================================}
{==============================================================================================}

procedure TOF_MBOTarifSaisiBase.FormateColSaisie(ACol, ARow: Longint);
var st, Stc: string;
begin
  St := TarifBase.Cells[ACol, ARow];
  StC := St;
  if ACol = ColPrix then StC := StrF00(Valeur(St), DEV.Decimale);
  TarifBase.Cells[ACol, ARow] := StC;
end;

procedure TOF_MBOTarifSaisiBase.FormatePrix;
var i: Integer;
begin
  for i := 1 to TarifBase.RowCount - 1 do
  begin
    DEV.Code := TOBBase.Detail[i - 1].GetValue('_DEV');
    GetInfosDevise(DEV);
    FormateColSaisie(ColPrix, i);
  end;
end;

procedure TOF_MBOTarifSaisiBase.ColTriangle(ACol, ARow: Longint; Canvas: TCanvas;
  AState: TGridDrawState);
var Triangle: array[0..2] of TPoint;
  Arect: Trect;
begin
  if Arow < TarifBase.Fixedrows then exit;
  if (gdFixed in AState) and (ACol = Col_Mov) then
  begin
    Arect := TarifBase.CellRect(Acol, Arow);
    Canvas.Brush.Color := TarifBase.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = TarifBase.row) then
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
end;

procedure TOF_MBOTarifSaisiBase.CodeGras(ACol, ARow: Longint; Canvas: TCanvas;
  AState: TGridDrawState);
begin
  if (ACol = colType) and (ARow > 0) then
  begin
    Canvas.Font.Style := [fsBold];
  end;
end;

initialization
  registerclasses([TOF_MBOTarifSaisiBase]);
end.
