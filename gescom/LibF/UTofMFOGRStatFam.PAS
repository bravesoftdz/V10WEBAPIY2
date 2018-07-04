{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/04/2001
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : GRSTATFAM
Suite ........ : Statistiques par famille article sur les tickets du Front Office
Mots clefs ... : TOF;UTOFMFOGRSTATFAM;FO
*****************************************************************}
unit UTofMFOGRStatFam;

interface
uses
  Controls, Classes, TeEngine, Chart, Series, Graphics, sysutils, forms,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, dbtables,
  {$ENDIF}
  HCtrls, HEnt1, GRS1, HMsgBox, M3FP, UTOF, UTOB, ParamSoc;

const
  TexteMessage: array[1..2] of string = (
    {1}'Vous devez sélectionner un niveau de famille !',
    {2}'Aucune donnée ne correspond aux critères de recherche !'
    );
type
  TOF_MFOGRSTATFAM = class(TOF)
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    NaturePiece: string; // nature de pièce des tickets
    TOBGraph: TOB; // TOB des lignes sélectionnées
    CurSelect: string; // select utilisé pour générer la TOBGraph
    CurSelect2: string; // select utilisé pour générer la TOBGraph
    TitresCrit: string; // Titre du critère choisi
    TitresDate: string; // Titre de la borne de date
    LookOrli: boolean; // Présentation ORLI (8 niveaux de familles)
    FamilleOrli: THMultiValComboBox;
    function VerifieIntegrite: boolean;
    function FamilleNivIsSelected(Indice: integer): boolean;
    function MakeSelect(Serie2: boolean): string;
    procedure MergeTob(TOBSerie1, TOBSerie2: TOB);
    procedure AddLibelleTob;
    function EstUnComparatif: Boolean;
    procedure FChartGetLegendText(Sender: TCustomAxisPanel; LegendStyle: TLegendStyle; Index: Integer; var LegendText: string);
    procedure FListeDblClick(Sender: TObject);
    procedure FChartClickSeries(Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure TraiteZoom(ValeurIndex, NoSerie: Integer; FromGrid: Boolean);
    procedure DateSelection(var Chaine: string; Serie2, PourWhere: Boolean);
  end;

implementation
uses
  GraphUtil, FODefi, FOUtil;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieIntegrite : vérifie la cohérence des choix de l'utilisateur
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOGRSTATFAM.VerifieIntegrite: boolean;
var Ind: integer;
begin
  Result := False;
  // un niveau de famille doît être choisi
  if LookOrli then
  begin
    if FamilleOrli.Text <> '' then Result := True;
  end else
  begin
    for Ind := 1 to 3 do
      if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
      begin
        Result := True;
        Break;
      end;
  end;
  if not Result then
  begin
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FamilleNivIsSelected : indique si un niveau de familles est sélectionné
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOGRSTATFAM.FamilleNivIsSelected(Indice: integer): boolean;
var Stg: string;
  Ind: integer;
begin
  Result := False;
  if LookOrli then
  begin
    if (Indice < 1) or (Indice > 8) then Exit;
    Stg := FamilleOrli.Text;
    if Stg = '' then Exit;
    Result := FOStrCmp('LF' + IntToStr(Indice), Stg, True);
    if Result then
    begin
      // on élimine les niveaux dont le libellé n'est pas défini
      Ind := FamilleOrli.Values.IndexOf('LF' + IntToStr(Indice));
      if (Ind >= 0) and (Ind < FamilleOrli.Items.Count) then
        Result := (FamilleOrli.Items[Ind] <> '.-')
      else
        Result := False;
    end;
  end else
  begin
    if (Indice < 1) or (Indice > 3) then Exit;
    Result := (GetControlText('FAMILLENIV' + IntToStr(Indice)) = 'X');
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MakeSelect : fabrique la clause select à partir des choix de l'utilisateur
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOGRSTATFAM.MakeSelect(Serie2: boolean): string;
var
  sCode, Stg: string;
  Ind: integer;
  ArtComplOk: boolean;
begin
  // Constitution du select
  sCode := '';
  ArtComplOk := False;
  for Ind := 1 to 3 do if FamilleNivIsSelected(Ind) then
    begin
      if sCode <> '' then sCode := sCode + '||';
      sCode := sCode + 'GL_FAMILLENIV' + IntToStr(Ind);
    end;
  if LookOrli then
  begin
    for Ind := 4 to 8 do if FamilleNivIsSelected(Ind) then
      begin
        if sCode <> '' then sCode := sCode + '||';
        sCode := sCode + 'ISNULL(GA2_FAMILLENIV' + IntToStr(Ind) + ',"   ")';
        ArtComplOk := True;
      end;
  end;
  if sCode = '' then Exit;
  Result := 'SELECT ' + sCode + ' AS CODE,SUM(GL_QTEFACT) AS QTE,SUM(GL_TOTALTTC) AS CA FROM LIGNE';
  if ArtComplOk then
    Result := Result + ' LEFT JOIN ARTICLECOMPL ON GL_ARTICLE=GA2_ARTICLE';
  Result := Result + ' WHERE GL_NATUREPIECEG="' + NaturePiece + '"'
    + ' AND GL_TYPELIGNE="ART" AND GL_TYPEARTICLE IN ("MAR","NOM")';
  Stg := GetControlText('GL_CAISSE');
  if Stg <> '' then Result := Result + ' AND GL_CAISSE="' + Stg + '"';
  Stg := GetControlText('GL_ETABLISSEMENT');
  if Stg <> '' then Result := Result + ' AND GL_ETABLISSEMENT="' + Stg + '"';
  DateSelection(Result, Serie2, True);
  Result := Result + ' GROUP BY ' + sCode;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MergeTob : fusion de la TOB de la 2ème série dans la TOB de la 1ère série
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.MergeTob(TOBSerie1, TOBSerie2: TOB);
var TOBS1, TOBS2: TOB;
  Ind: Integer;
  ValeurCle: Variant;
begin
  if (TOBSerie1 = nil) or (TOBSerie2 = nil) then Exit;
  // ajout des colonnes de la TOB de la 2ème série dans la TOB de la 1ère série
  for Ind := 0 to TOBSerie1.Detail.Count - 1 do
  begin
    TOBS1 := TOBSerie1.Detail[Ind];
    if TOBS1.FieldExists('CA2') then TOBS1.PutValue('CA2', 0)
    else TOBS1.AddChampSupValeur('CA2', 0);
    if TOBS1.FieldExists('QTE2') then TOBS1.PutValue('QTE2', 0)
    else TOBS1.AddChampSupValeur('QTE2', 0);
  end;
  for Ind := 0 to TOBSerie2.Detail.Count - 1 do
  begin
    TOBS2 := TOBSerie2.Detail[Ind];
    // recherche de la TOB de la 1ère série équivalente
    ValeurCle := TOBS2.GetValue('CODE');
    TOBS1 := TOBSerie1.FindFirst(['CODE'], [ValeurCle], False);
    if TOBS1 = nil then
    begin
      TOBS1 := TOB.Create(TOBSerie1.NomTable, TOBSerie1, -1);
      TOBS1.AddChampSupValeur('CODE', ValeurCle);
      TOBS1.AddChampSupValeur('CA', 0);
      TOBS1.AddChampSupValeur('QTE', 0);
      TOBS1.AddChampSupValeur('CA2', 0);
      TOBS1.AddChampSupValeur('QTE2', 0);
    end;
    // ajout des champs de la TOB e la 2ème série inexistants dans la TOB de la 1ère série
    TOBS1.PutValue('CA2', TOBS2.GetValue('CA'));
    TOBS1.PutValue('QTE2', TOBS2.GetValue('QTE'));
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AddLibelleTob : Ajout des colonnes libellé et n°
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.AddLibelleTob;
var Ind, Nl: integer;
  TOBL: TOB;
  sCode, sLib: string;
  Val: Variant;
begin
  for Nl := 0 to TOBGraph.Detail.Count - 1 do
  begin
    TOBL := TOBGraph.Detail[Nl];
    TOBL.AddChampSupValeur('NUMERO', Nl, False);
    Val := TOBL.GetValue('CODE');
    if VarIsNull(Val) then
    begin
      sCode := '';
    end else
    begin
      sCode := VarAsType(Val, VarString);
      if sCode = #0 then sCode := '';
    end;
    for Ind := 1 to 8 do
      if FamilleNivIsSelected(Ind) then
      begin
        sLib := RechDom('GCFAMILLENIV' + IntToStr(Ind), FOExtract(sCode, 1, 3), False);
        TOBL.AddChampSupValeur('FAMILLENIV' + IntToStr(Ind), sLib, False);
      end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EstUnComparatif : vérifie si on compare 2 séries
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOGRSTATFAM.EstUnComparatif: Boolean;
begin
  Result := (GetControlText('DATEPIECE') <> GetControlText('DATEPIECE2'))
    or (GetControlText('DATEPIECE_') <> GetControlText('DATEPIECE2_'));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnUpdate :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.OnUpdate;
var sSelect, sColGraph, sTitre, sColonnes, sTitresCol, sTitresDate2: string;
  Ok: Boolean;
  tstTitre1: Tstrings;
  Ind, Nl, NbNiv: Integer;
  QQ: TQuery;
  TOBG: TOB;
begin
  inherited;
  // Contrôles des choix utilisateurs
  VerifieIntegrite;
  // Constitution du select
  sSelect := MakeSelect(False);
  // sélection des données
  if sSelect <> CurSelect then
  begin
    QQ := OpenSql(sSelect, True);
    if QQ.EOF then
    begin
      Ferme(QQ);
      LastError := 2;
      LastErrorMsg := TexteMessage[LastError];
      Exit;
    end;
    if TOBGraph <> nil then TOBGraph.Free;
    TOBGraph := TOB.Create('', nil, -1);
    TOBGraph.LoadDetailDB('', '', '', QQ, False);
    Ferme(QQ);
    ///TOBGraph.LoadDetailFromSQL(sSelect) ;
    if (TOBGraph = nil) or (TOBGraph.Detail.Count <= 0) then
    begin
      if TOBGraph <> nil then TOBGraph.Free;
      LastError := 2;
      LastErrorMsg := TexteMessage[LastError];
      Exit;
    end;
    CurSelect := sSelect;
    CurSelect2 := '';
  end;
  // Constitution du select de la 2ème série
  sSelect := MakeSelect(True);
  if sSelect <> CurSelect2 then
  begin
    TOBG := TOB.Create('', nil, -1);
    // sélection des données de la 2ème série
    if EstUnComparatif then
    begin
      QQ := OpenSql(sSelect, True);
      if not QQ.EOF then TOBG.LoadDetailDB('', '', '', QQ, False);
      Ferme(QQ);
    end;
    ///if EstUnComparatif then TOBG.LoadDetailFromSQL(sSelect) ;
    MergeTob(TOBGraph, TOBG);
    TOBG.Free;
    CurSelect2 := sSelect;
  end;
  // Ajout des colonnes libellé et n°
  AddLibelleTob;
  // Choix de la colonne du graphe
  if GetControlText('BQTE') = 'X' then sColGraph := 'QTE' else sColGraph := 'CA';
  sColonnes := 'NUMERO;CODE';
  sTitresCol := TraduireMemoire('N°') + ';' + TraduireMemoire('Code');
  Nl := 2;
  TFGRS1(Ecran).FListe.ColCount := Nl;
  TFGRS1(Ecran).FListe.ColAligns[(Nl - 2)] := taCenter;
  TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taLeftJustify;
  for Ind := 1 to 8 do if FamilleNivIsSelected(Ind) then
    begin
      sColonnes := sColonnes + ';FAMILLENIV' + IntToStr(Ind);
      sTitresCol := sTitresCol + ';' + RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Ind), False);
      Inc(Nl);
      TFGRS1(Ecran).FListe.ColCount := Nl;
      TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taLeftJustify;
    end;
  sColonnes := sColonnes + ';QTE;CA';
  sTitresCol := sTitresCol + ';' + TraduireMemoire('Quantité')
    + ';' + TraduireMemoire('Chiffre d''affaires');
  Inc(Nl, 2);
  TFGRS1(Ecran).FListe.ColCount := Nl;
  TFGRS1(Ecran).FListe.ColAligns[(Nl - 2)] := taRightJustify;
  TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taRightJustify;
  if EstUnComparatif then
  begin
    if GetControlText('BQTE') = 'X' then sColGraph := sColGraph + ';QTE2' else sColGraph := sColGraph + ';CA2';
    sColonnes := sColonnes + ';QTE2;CA2';
    sTitresCol := sTitresCol + ';' + TraduireMemoire('Quantité n°2')
      + ';' + TraduireMemoire('C.A. n°2');
    Inc(Nl, 2);
    TFGRS1(Ecran).FListe.ColCount := Nl;
    TFGRS1(Ecran).FListe.ColAligns[(Nl - 2)] := taRightJustify;
    TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taRightJustify;
  end;
  // Constitution du titre
  tstTitre1 := TStringList.Create;
  sTitre := TraduireMemoire('Statistiques en') + ' ';
  if sColGraph = 'QTE' then sTitre := sTitre + TraduireMemoire('Quantité')
  else sTitre := sTitre + RechDom('TTDEVISE', V_PGI.DevisePivot, False);
  sTitre := sTitre + ' ' + TraduireMemoire('par') + ' ';
  TitresCrit := '';
  for Ind := 1 to 8 do if FamilleNivIsSelected(Ind) then
    begin
      if TitresCrit <> '' then TitresCrit := TitresCrit + ', ';
      TitresCrit := TitresCrit + RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Ind), False);
    end;
  tstTitre1.Add(sTitre + TitresCrit);
  TitresDate := TraduireMemoire('du') + ' ' + GetControlText('DATEPIECE') + ' '
    + TraduireMemoire('au') + ' ' + GetControlText('DATEPIECE_');
  sTitresDate2 := '';
  if EstUnComparatif then
  begin
    sTitresDate2 := ' ' + TraduireMemoire('et') + ' '
      + TraduireMemoire('du') + ' ' + GetControlText('DATEPIECE2') + ' '
      + TraduireMemoire('au') + ' ' + GetControlText('DATEPIECE2_');
  end;
  tstTitre1.Add(TitresDate + sTitresDate2);
  // lancement du graphe
  LanceGraph(TFGRS1(Ecran), TOBGraph, 'LIGNE', sColonnes, '', sTitresCol,
    sColGraph, '', tstTitre1, nil, TBarSeries, 'CODE', False);
  // Libération des objets créés
  tstTitre1.Free;
  // Branchement du zoom
  TFGRS1(Ecran).FListe.OnDblClick := FListeDblClick;
  TFGRS1(Ecran).FChart1.OnClickSeries := FChartClickSeries;
  TFGRS1(Ecran).FChart1.OnGetLegendText := FChartGetLegendText;
  // Définition des étiquettes
  TFGRS1(Ecran).FChart1.Series[0].ColorEachPoint := True;
  TFGRS1(Ecran).FChart1.Legend.LegendStyle := lsValues;
  TFGRS1(Ecran).FChart1.Legend.TextStyle := ltsPlain;
  // On affiche le code sous l'axe sauf si les 3 niveaux de famille sont choisis on affiche alors le rang
  //if (GetControlText('FAMILLENIV1') = 'X') and (GetControlText('FAMILLENIV2') = 'X') and (GetControlText('FAMILLENIV3') = 'X') then
  // On affiche le code sous l'axe sauf si plus de 2 niveaux de famille sont choisis on affiche alors le rang
  NbNiv := 0;
  for Ind := 1 to 8 do if FamilleNivIsSelected(Ind) then Inc(NbNiv);
  if NbNiv > 2 then
  begin
    TFGRS1(Ecran).FChart1.BottomAxis.LabelStyle := talValue;
  end else TFGRS1(Ecran).FChart1.BottomAxis.LabelStyle := talText;
  if EstUnComparatif then TFGRS1(Ecran).FChart1.BottomAxis.Title.Caption := TraduireMemoire('Période n°1 - Période n°2')
  else TFGRS1(Ecran).FChart1.BottomAxis.Title.Caption := '';
  TFGRS1(Ecran).FChart1.Series[0].ValueFormat := '# ##0';
  TFGRS1(Ecran).FChart1.Series[0].PercentFormat := '##0.00%';
  if GetControlText('BVALEUR') = 'X' then TFGRS1(Ecran).FChart1.Series[0].Marks.Style := smsValue
  else TFGRS1(Ecran).FChart1.Series[0].Marks.Style := smsPercent;
  //if GetControlText('BVALEUR') = 'X' then TFGRS1(Ecran).FChart1.Legend.TextStyle := ltsLeftValue
  //                                   else TFGRS1(Ecran).FChart1.Legend.TextStyle := ltsLeftPercent ;
  // Définition des étiquettes de la 2ème série
  if EstUnComparatif then
  begin
    TFGRS1(Ecran).FChart1.Series[1].ColorEachPoint := True;
    if TFGRS1(Ecran).FChart1.Series[1] is TBarSeries then
      TBarSeries(TFGRS1(Ecran).FChart1.Series[1]).BarBrush.Style := bsDiagCross;
    TFGRS1(Ecran).FChart1.Series[1].ValueFormat := '# ##0';
    TFGRS1(Ecran).FChart1.Series[1].PercentFormat := '##0.00%';
    if GetControlText('BVALEUR') = 'X' then TFGRS1(Ecran).FChart1.Series[1].Marks.Style := smsValue
    else TFGRS1(Ecran).FChart1.Series[1].Marks.Style := smsPercent;
  end;
  // Rend toutes les colonnes de la grille visible
  Ok := False;
  if TFGRS1(Ecran).FListe.ColWidths[0] <> 19 then
  begin
    TFGRS1(Ecran).FListe.ColWidths[0] := 19;
    Ok := True;
  end;
  for Ind := 1 to TFGRS1(Ecran).FListe.ColCount - 1 do if TFGRS1(Ecran).FListe.ColWidths[Ind] <> TFGRS1(Ecran).FListe.DefaultColWidth then
    begin
      TFGRS1(Ecran).FListe.ColWidths[Ind] := TFGRS1(Ecran).FListe.DefaultColWidth;
      Ok := True;
    end;
  if Ok then TFGRS1(Ecran).HMTrad.ResizeGridColumns(TFGRS1(Ecran).FListe);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnLoad :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.OnLoad;
begin
  inherited;
  SetControlProperty('bAffGraph', 'Down', True);
  SetControlProperty('FListe', 'Visible', False);
  // Création de la TOB des lignes à traiter
  TOBGraph := TOB.Create('', nil, -1);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnArgument :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.OnArgument(S: string);
var Ind: integer;
  Stg, Champ, Valeur: string;
  CC: TControl;
begin
  inherited;
  Stg := ReadTokenSt(S);
  while Stg <> '' do
  begin
    Ind := Pos('=', Stg);
    if Ind > 0 then
    begin
      Champ := Copy(Stg, 1, Ind - 1);
      Valeur := Copy(Stg, Ind + 1, length(Stg));
      if Champ = 'NATUREPIECE' then
        NaturePiece := Valeur
      else
        SetControlText(Champ, Valeur);
    end;
    Stg := ReadTokenSt(S);
  end;
  CC := GetControl('FAMILLEORLI');
  if (CC <> nil) and (CC is THMultiValComboBox) then
  begin
    FamilleOrli := THMultiValComboBox(CC);
    LookOrli := ((V_PGI.LaSerie > S3) and GetParamSoc('SO_ARTLOOKORLI'));
  end else
  begin
    FamilleOrli := nil;
    LookOrli := False;
  end;
  for Ind := 1 to 3 do SetControlVisible('FAMILLENIV' + IntToStr(Ind), not LookOrli);
  if FamilleOrli <> nil then FamilleOrli.Visible := LookOrli;
  {$IFDEF EAGLCLIENT}
  SetControlVisible('BGraph', False);
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnClose :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.OnClose;
begin
  inherited;
  // Suppression de la TOB des lignes à traiter
  if TOBGraph <> nil then TOBGraph.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FChartGetLegendText : indique le texte à afficher comme légende
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.FChartGetLegendText(Sender: TCustomAxisPanel; LegendStyle: TLegendStyle; Index: Integer; var LegendText: string);
var TOBL: TOB;
  Stg: string;
  Ind, NbF, NoFam: Integer;
begin
  inherited;
  Stg := '';
  // si un seul niveau de famille est choisi on affiche le libellé sion on affiche le code
  NbF := 0;
  NoFam := 0;
  for Ind := 1 to 8 do if FamilleNivIsSelected(Ind) then
    begin
      NoFam := Ind;
      Inc(NbF);
    end;
  if NbF = 1 then
  begin
    TOBL := TOBGraph.FindFirst(['CODE'], [LegendText], False);
    if TOBL <> nil then Stg := Trim(TOBL.GetValue('FAMILLENIV' + IntToStr(NoFam)));
  end;
  if Stg <> '' then LegendText := Stg;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FListeDblClick : double click dans la grille
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.FListeDblClick(Sender: TObject);
begin
  TraiteZoom(TFGRS1(Ecran).Fliste.Row - TFGRS1(Ecran).FLIste.FixedRows, 0, TRUE);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FChartClickSeries : double click dans le graphe
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.FChartClickSeries(Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y:
  Integer);
var Stg, sCode, sTitre, sFam: string;
  Ind, NoSerie: Integer;
begin
  sCode := Series.Name;
  if (Length(sCode) <= 2) or (Copy(sCode, 1, 2) <> 'SS') then Exit;
  NoSerie := ValeurI(Copy(sCode, 3, Length(sCode)));
  if ssRight in shift then
  begin
    sCode := TOBGraph.Detail[ValueIndex].GetValue('CODE');
    Stg := TraduireMemoire('Code') + ' : ' + sCode;
    for Ind := 1 to 8 do if FamilleNivIsSelected(Ind) then
      begin
        sTitre := RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Ind), False);
        sFam := RechDom('GCFAMILLENIV' + IntToStr(Ind), FOExtract(sCode, 1, 3), False);
        Stg := Stg + '#10 ' + sTitre + '  : ' + sFam;
      end;
    Stg := Stg + '#10 ' + TraduireMemoire('Quantité') + '  : '
      + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('QTE'), 12, V_PGI.OkDecQ, '', True);
    Stg := Stg + '#10 ' + TraduireMemoire('C.A.') + '  : '
      + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('CA'), 12, V_PGI.OkDecV, V_PGI.SymbolePivot, True);
    if NoSerie = 0 then Stg := Stg + '#10 ' + TraduireMemoire('Soit') + '  : ' + Series.MarkPercent(ValueIndex, False);
    if EstUnComparatif then
    begin
      Stg := Stg + '#10 ' + TraduireMemoire('Quantité n°2') + '  : '
        + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('QTE2'), 12, V_PGI.OkDecQ, '', True);
      Stg := Stg + '#10 ' + TraduireMemoire('C.A. n°2') + '  : '
        + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('CA2'), 12, V_PGI.OkDecV, V_PGI.SymbolePivot, True);
      if NoSerie = 1 then Stg := Stg + '#10 ' + TraduireMemoire('Soit') + '  : ' + Series.MarkPercent(ValueIndex, False);
    end;
    PGIInfo(Stg, ' ' + TraduireMemoire('Serie n°') + IntToStr(ValueIndex));
  end else TraiteZoom(ValueIndex, NoSerie, FALSE);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteZoom : zoom sur les lignes constituant la série
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.TraiteZoom(ValeurIndex, NoSerie: integer; FromGrid: boolean);
var
  Ind: integer;
  sRange, sFam, sWhere, sOpeVide, sCaisse: string;
begin
  if (ValeurIndex < 0) or (ValeurIndex > (TFGRS1(Ecran).FListe.RowCount - TFGRS1(Ecran).FListe.FixedRows)) then
    Exit;
  {$IFDEF EAGLCLIENT}
  sOpeVide := ' IN ("")';
  {$ELSE}
  sOpeVide := '##""';
  {$ENDIF}
  // Recherche des codes familles
  sFam := '';
  if FromGrid then
    sFam := TFGRS1(Ecran).FListe.Cells[1, ValeurIndex + TFGRS1(Ecran).Fliste.FixedRows]
  else
    sFam := TOBGraph.Detail[ValeurIndex].GetValue('CODE');
  // Initialisation des critères de recherche
  sRange := 'GL_NATUREPIECEG='+ NaturePiece +';GL_TYPELIGNE=ART';
  sWhere := ';XX_WHERE=GL_TYPEARTICLE IN ("MAR","NOM")';
  sCaisse := GetControlText('GL_CAISSE');
  if sCaisse <> '' then sRange := sRange + ';GL_CAISSE=' + sCaisse;
  for Ind := 1 to 3 do if FamilleNivIsSelected(Ind) then
    begin
      if sFam = '' then
        sWhere := sWhere + ' AND GL_FAMILLENIV' + IntToStr(Ind) + sOpeVide
      else
        sRange := sRange + ';GL_FAMILLENIV' + IntToStr(Ind) + '=' + FOExtract(sFam, 1, 3);
    end;
  for Ind := 4 to 8 do if FamilleNivIsSelected(Ind) then
    begin
      if sFam = '' then
        sWhere := sWhere + ' AND (GA2_FAMILLENIV' + IntToStr(Ind) + sOpeVide
          + ' OR GA2_FAMILLENIV' + IntToStr(Ind) + ' IS NULL)'
      else
        sRange := sRange + ';GA2_FAMILLENIV' + IntToStr(Ind) + '=' + FOExtract(sFam, 1, 3);
    end;
  sRange := sRange + sWhere;
  DateSelection(sRange, (NoSerie = 1), False);
  // Lancement de la liste des lignes
  if LookOrli then
    AGLLanceFiche('MFO', 'PIECARTS5_MUL', sRange, '', 'STATISTIQUES')
  else
    AGLLanceFiche('MFO', 'PIECART_MUL', sRange, '', 'STATISTIQUES');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DateSelection : prend en compte les dates de sélection pour une clause where ou une clause range
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATFAM.DateSelection(var Chaine: string; Serie2, PourWhere: Boolean);
var sBorneInf, sBorneSup: string;
begin
  if Serie2 then
  begin
    sBorneInf := GetControlText('DATEPIECE2');
    sBorneSup := GetControlText('DATEPIECE2_');
  end else
  begin
    sBorneInf := GetControlText('DATEPIECE');
    sBorneSup := GetControlText('DATEPIECE_');
  end;
  if PourWhere then
  begin
    if Trim(UpperCase(Chaine)) <> 'SELECT' then Chaine := Chaine + ' and ';
    Chaine := Chaine + 'GL_DATEPIECE>="' + USDateTime(StrToDate(sBorneInf)) + '" and '
      + 'GL_DATEPIECE<="' + USDateTime(StrToDate(sBorneSup)) + '"';
  end else
  begin
    if Chaine <> '' then Chaine := Chaine + ';';
    Chaine := Chaine + 'GL_DATEPIECE=' + sBorneInf + ';GL_DATEPIECE_=' + sBorneSup;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOImprimeStatFam = lance l'impression des statistiques sur
Suite ........ : par famille en mode ticket depuis le script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure ImprimeStatFam(Parms: array of variant; Nb: integer);
var FF: TForm;
  TOBG, TOBL: TOB;
begin
  FF := TForm(Longint(Parms[0]));
  if not (FF is TFGRS1) then Exit;
  if not (TFGRS1(FF).LaTOF is TOF_MFOGRSTATFAM) then Exit;
  TOBG := TOF_MFOGRSTATFAM(TFGRS1(FF).LaTOF).TOBGraph;
  if (TOBG = nil) or (TOBG.Detail.Count <= 0) then Exit;
  TOBL := TOBG.Detail[0];
  if TOBL = nil then Exit;
  // Le titre de l'état est ajouté au 1er élément de la TOB
  TOBL.AddChampSupValeur('TITRE', TOF_MFOGRSTATFAM(TFGRS1(FF).LaTOF).TitresCrit);
  TOBL.AddChampSupValeur('BORNEDATE', TOF_MFOGRSTATFAM(TFGRS1(FF).LaTOF).TitresDate);
  // Lancement de l'impression
  FOLanceImprimeLP(efoStatFam, '', False, TOBG);
end;

initialization
  RegisterClasses([TOF_MFOGRSTATFAM]);
  RegisterAglProc('FOImprimeStatFam', TRUE, 0, ImprimeStatFam);
end.
