{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/04/2001
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : GRSTATREM
Suite ........ : Statistiques sur les remises consenties dans les tickets
Suite ........ : du Front Office
Mots clefs ... : TOF;UTOFMFOGRSTATREM;FO
*****************************************************************}
unit UTofMFOGRStatRem;

interface
uses
  Controls, Classes, TeEngine, Chart, Series, Graphics, sysutils, forms,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, dbtables,
  {$ENDIF}
  HCtrls, HEnt1, GRS1, HMsgBox, M3FP, UTOF, UTOB;

const
  TexteMessage: array[1..2] of string = (
    {1}'Vous devez sélectionner un niveau de famille !',
    {2}'Aucune donnée ne correspond aux critères de recherche !'
    );
type
  TOF_MFOGRSTATREM = class(TOF)
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    NaturePiece: string; // nature de pièce des tickets
    TOBGraph: TOB; // TOB des lignes sélectionnées
    CurSelect: string; // select utilisé pour générer la TOBGraph
    Comparatif: Boolean; // mode domparaison entre remise et non remisé
    TitresCrit: string; // Titre du critère choisi
    TitresDate: string; // Titre de la borne de date
    function VerifieIntegrite: Boolean;
    function MakeSelect(Serie2: Boolean): string;
    procedure MergeTob(TOBSerie1, TOBSerie2: TOB);
    procedure AddLibelleTob;
    procedure FChartGetLegendText(Sender: TCustomAxisPanel; LegendStyle: TLegendStyle; Index: Integer; var LegendText: string);
    procedure FListeDblClick(Sender: TObject);
    procedure FChartClickSeries(Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure TraiteZoom(ValeurIndex, NoSerie: Integer; FromGrid: Boolean);
    procedure DateSelection(var Chaine: string; PourWhere: Boolean);
  end;

implementation
uses
  GraphUtil, FODefi, FOUtil;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieIntegrite : vérifie la cohérence des choix de l'utilisateur
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOGRSTATREM.VerifieIntegrite: Boolean;
var Ind: Integer;
begin
  Result := True;
  // Si on travaille sur les familles, un niveau doît être choisi
  if GetControlText('BFAMILLE') = 'X' then
  begin
    Result := False;
    for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
      begin
        Result := True;
        Break;
      end;
    if not Result then
    begin
      LastError := 1;
      LastErrorMsg := TexteMessage[LastError];
      Exit;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MakeSelect : fabrique la clause select à partir des choix de l'utilisateur
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOGRSTATREM.MakeSelect(Serie2: Boolean): string;
var sCode, sOrder: string;
  Ind: Integer;
begin
  inherited;
  sCode := '';
  sOrder := '';
  if GetControlText('BFAMILLE') = 'X' then
  begin
    for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
      begin
        if sCode <> '' then sCode := sCode + '||';
        sCode := sCode + 'GL_FAMILLENIV' + IntToStr(Ind);
        if sOrder <> '' then sOrder := sOrder + ',';
        sOrder := sOrder + 'GL_FAMILLENIV' + IntToStr(Ind);
      end;
  end else
    if GetControlText('BCOLLECTION') = 'X' then
  begin
    sCode := 'GL_COLLECTION';
    sOrder := 'GL_COLLECTION';
  end else
    if GetControlText('BVENDEUR') = 'X' then
  begin
    sCode := 'GL_REPRESENTANT';
    sOrder := 'GL_REPRESENTANT';
  end else
    if GetControlText('BDEMARQUE') = 'X' then
  begin
    sCode := 'GL_TYPEREMISE';
    sOrder := 'GL_TYPEREMISE';
  end;
  if sCode = '' then Exit else sCode := sCode + ' AS CODE';
  Result := 'SELECT ' + sCode;
  if Serie2 then Result := Result + ',SUM(GL_MONTANTTTC) AS CABRUT, SUM(GL_QTEFACT) AS QTBRUT'
  else Result := Result + ',SUM(GL_MONTANTTTC) AS CA, SUM(((GL_PUTTC*GL_QTEFACT)-GL_MONTANTTTC)) AS MNT'
    + ',AVG(GL_REMISELIGNE) AS POURCENT, SUM(GL_QTEFACT) AS QTE';
  Result := Result + ' from LIGNE';
  Result := Result + ' where GL_NATUREPIECEG="'+ NaturePiece + '"'
          + ' AND GL_TYPELIGNE="ART" AND GL_TYPEARTICLE IN ("MAR","NOM")';
  sCode := GetControlText('GL_CAISSE');
  if sCode <> '' then Result := Result + ' AND GL_CAISSE="' + sCode + '"';
  sCode := GetControlText('GL_ETABLISSEMENT');
  if sCode <> '' then Result := Result + ' AND GL_ETABLISSEMENT="' + sCode + '"';
  if Serie2 then
  begin
    Result := Result + ' AND GL_TOTREMLIGNE=0';
  end else
  begin
    if (GetControlText('BREMONLY') = 'X') or (GetControlText('BCOMPARE') = 'X') then
      Result := Result + ' AND GL_TOTREMLIGNE<>0';
  end;
  DateSelection(Result, True);
  if sOrder <> '' then Result := Result + ' GROUP BY ' + sOrder;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MergeTob : fusion de la TOB de la 2ème série dans la TOB de la 1ère série
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.MergeTob(TOBSerie1, TOBSerie2: TOB);
var TOBS1, TOBS2: TOB;
  Ind: Integer;
  ValeurCle: Variant;
begin
  if (TOBSerie1 = nil) or (TOBSerie2 = nil) then Exit;
  // ajout des colonnes de la TOB de la 2ème série dans la TOB de la 1ère série
  for Ind := 0 to TOBSerie1.Detail.Count - 1 do
  begin
    TOBS1 := TOBSerie1.Detail[Ind];
    TOBS1.AddChampSupValeur('CABRUT', 0);
    TOBS1.AddChampSupValeur('QTBRUT', 0);
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
      TOBS1.AddChampSupValeur('MNT', 0);
      TOBS1.AddChampSupValeur('POURCENT', 0);
      TOBS1.AddChampSupValeur('QTE', 0);
      TOBS1.AddChampSupValeur('CABRUT', 0);
      TOBS1.AddChampSupValeur('QTBRUT', 0);
    end;
    // ajout des champs de la TOB e la 2ème série inexistants dans la TOB de la 1ère série
    TOBS1.PutValue('CABRUT', TOBS2.GetValue('CABRUT'));
    TOBS1.PutValue('QTBRUT', TOBS2.GetValue('QTBRUT'));
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AddLibelleTob : Ajout des colonnes libellé et n°
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.AddLibelleTob;
var Ind, Nl: Integer;
  TOBL: TOB;
  sCode, sLib: string;
begin
  for Nl := 0 to TOBGraph.Detail.Count - 1 do
  begin
    TOBL := TOBGraph.Detail[Nl];
    TOBL.AddChampSupValeur('NUMERO', Nl, False);
    sCode := TOBL.GetValue('CODE');
    if GetControlText('BFAMILLE') = 'X' then
    begin
      for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
        begin
          TOBL.AddChampSupValeur('FAMILLENIV' + IntToStr(Ind), RechDom('GCFAMILLENIV' + IntToStr(Ind), FOExtract(sCode, 1, 3), False), False);
        end;
    end else
      if GetControlText('BCOLLECTION') = 'X' then
    begin
      if sCode = '' then sLib := TraduireMemoire('<<Aucun>>') else sLib := RechDom('GCCOLLECTION', sCode, False);
      TOBL.AddChampSupValeur('LIBELLE', sLib, False);
    end else
      if GetControlText('BVENDEUR') = 'X' then
    begin
      if sCode = '' then sLib := TraduireMemoire('<<Aucun>>') else sLib := RechDom('GCCOMMERCIAL', sCode, False);
      TOBL.AddChampSupValeur('LIBELLE', sLib, False);
    end else
      if GetControlText('BDEMARQUE') = 'X' then
    begin
      if sCode = '' then sLib := TraduireMemoire('<<Aucun>>') else sLib := RechDom('GCTYPEREMISE', sCode, False);
      TOBL.AddChampSupValeur('LIBELLE', sLib, False);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnUpdate :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.OnUpdate;
var sSelect, sColGraph, sTitre, sColonnes, sTitresCol: string;
  tstTitre1: Tstrings;
  Ind, Nl, Nbc: Integer;
  QQ: TQuery;
  Ok, doSelect: Boolean;
  TOBG: TOB;
begin
  inherited;
  // Contrôles des choix utilisateurs
  if not VerifieIntegrite then Exit;
  // Constitution du select de la 1ère série
  sSelect := MakeSelect(False);
  doSelect := False;
  // sélection des données de la 1ère série
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
    doSelect := True;
  end;
  // Constitution du select de la 2ème série
  if ((doSelect) or (not Comparatif)) and (GetControlText('BCOMPARE') = 'X') then
  begin
    sSelect := MakeSelect(True);
    // sélection des données de la 2ème série
    QQ := OpenSql(sSelect, True);
    if not QQ.EOF then
    begin
      TOBG := TOB.Create('', nil, -1);
      TOBG.LoadDetailDB('', '', '', QQ, False);
      MergeTob(TOBGraph, TOBG);
      TOBG.Free;
    end;
    Ferme(QQ);
    ///TOBG := TOB.Create('', Nil, -1) ;
    ///TOBG.LoadDetailFromSQL(sSelect) ;
    ///MergeTob(TOBGraph, TOBG) ;
    ///TOBG.Free ;
  end;
  Comparatif := (GetControlText('BCOMPARE') = 'X');
  // Ajout des colonnes libellé et n°
  AddLibelleTob;
  // Choix de la colonne du graphe
  sColGraph := 'MNT';
  if GetControlText('BCA') = 'X' then sColGraph := 'CA' else
    if GetControlText('BMONTANT') = 'X' then sColGraph := 'MNT' else
    if GetControlText('BMOYPOURCENT') = 'X' then sColGraph := 'POURCENT' else ;
  if GetControlText('BQTE') = 'X' then sColGraph := 'QTE' else ;
  if GetControlText('BCOMPARE') = 'X' then
  begin
    if GetControlText('BCA') = 'X' then sColGraph := sColGraph + ';CABRUT' else
      if GetControlText('BQTE') = 'X' then sColGraph := sColGraph + ';QTBRUT' else ;
  end;
  sColonnes := 'NUMERO;CODE';
  sTitresCol := TraduireMemoire('N°') + ';' + TraduireMemoire('Code');
  Nl := 2;
  TFGRS1(Ecran).FListe.ColCount := Nl;
  TFGRS1(Ecran).FListe.ColAligns[(Nl - 2)] := taCenter;
  TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taLeftJustify;
  if GetControlText('BFAMILLE') = 'X' then
  begin
    for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
      begin
        sColonnes := sColonnes + ';FAMILLENIV' + IntToStr(Ind);
        sTitresCol := sTitresCol + ';' + RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Ind), False);
        Inc(Nl);
        TFGRS1(Ecran).FListe.ColCount := Nl;
        TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taLeftJustify;
      end;
  end else
  begin
    sColonnes := sColonnes + ';LIBELLE';
    if GetControlText('BCOLLECTION') = 'X' then sTitresCol := sTitresCol + ';' + TraduireMemoire('Collection') else
      if GetControlText('BVENDEUR') = 'X' then sTitresCol := sTitresCol + ';' + TraduireMemoire('Vendeur') else
      if GetControlText('BDEMARQUE') = 'X' then sTitresCol := sTitresCol + ';' + TraduireMemoire('Démarque');
    Inc(Nl);
    TFGRS1(Ecran).FListe.ColCount := Nl;
    TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taLeftJustify;
  end;
  sColonnes := sColonnes + ';CA;MNT;POURCENT;QTE';
  sTitresCol := sTitresCol + ';' + TraduireMemoire('Chiffre d''affaires')
    + ';' + TraduireMemoire('Montant')
    + ';' + TraduireMemoire('Pourcentage')
    + ';' + TraduireMemoire('Quantité');
  Nbc := 4;
  if GetControlText('BCOMPARE') = 'X' then
  begin
    sColonnes := sColonnes + ';CABRUT;QTBRUT';
    sTitresCol := sTitresCol + ';' + TraduireMemoire('CA sans remise')
      + ';' + TraduireMemoire('Qté sans remise');
    Inc(Nbc, 2);
  end;
  for Ind := 1 to Nbc do
  begin
    Inc(Nl);
    TFGRS1(Ecran).FListe.ColCount := Nl;
    TFGRS1(Ecran).FListe.ColAligns[(Nl - 1)] := taRightJustify;
  end;
  // Constitution du titre
  tstTitre1 := TStringList.Create;
  sTitre := TraduireMemoire('Statistiques en') + ' ';
  if FOStrCmp('POURCENT', sColGraph) then sTitre := sTitre + TraduireMemoire('Pourcentage') else
    if FOStrCmp('QTE', sColGraph) then sTitre := sTitre + TraduireMemoire('Quantité')
  else sTitre := sTitre + RechDom('TTDEVISE', V_PGI.DevisePivot, False);
  sTitre := sTitre + ' ' + TraduireMemoire('par') + ' ';
  TitresCrit := '';
  if GetControlText('BFAMILLE') = 'X' then
  begin
    for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
      begin
        if TitresCrit <> '' then TitresCrit := TitresCrit + ', ';
        TitresCrit := TitresCrit + RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Ind), False);
      end;
  end else
    if GetControlText('BCOLLECTION') = 'X' then
  begin
    TitresCrit := TitresCrit + TraduireMemoire('Collection');
  end else
    if GetControlText('BVENDEUR') = 'X' then
  begin
    TitresCrit := TitresCrit + TraduireMemoire('Vendeur');
  end else
    if GetControlText('BDEMARQUE') = 'X' then
  begin
    TitresCrit := TitresCrit + TraduireMemoire('Démarque');
  end;
  tstTitre1.Add(sTitre + TitresCrit);
  TitresDate := TraduireMemoire('du') + ' ' + GetControlText('DATEPIECE') + ' '
    + TraduireMemoire('au') + ' ' + GetControlText('DATEPIECE_');
  tstTitre1.Add(TitresDate);
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
  TFGRS1(Ecran).FChart1.Legend.LegendStyle := lsValues;
  TFGRS1(Ecran).FChart1.Legend.TextStyle := ltsPlain;
  // On affiche le code sous l'axe sauf si les 3 niveaux de famille sont choisis on affiche alors le rang
  if (GetControlText('FAMILLENIV1') = 'X') and (GetControlText('FAMILLENIV2') = 'X') and (GetControlText('FAMILLENIV3') = 'X') then
  begin
    TFGRS1(Ecran).FChart1.BottomAxis.LabelStyle := talValue;
  end else TFGRS1(Ecran).FChart1.BottomAxis.LabelStyle := talText;
  TFGRS1(Ecran).FChart1.BottomAxis.Title.Caption := '';
  if GetControlText('BCOMPARE') = 'X' then
    TFGRS1(Ecran).FChart1.BottomAxis.Title.Caption := TraduireMemoire('Avec Remise - Sans Remise');
  TFGRS1(Ecran).FChart1.Series[0].ColorEachPoint := True;
  TFGRS1(Ecran).FChart1.Series[0].ValueFormat := '# ##0';
  TFGRS1(Ecran).FChart1.Series[0].PercentFormat := '##0.00%';
  if GetControlText('BVALEUR') = 'X' then TFGRS1(Ecran).FChart1.Series[0].Marks.Style := smsValue
  else TFGRS1(Ecran).FChart1.Series[0].Marks.Style := smsPercent;
  // Définition des étiquettes de la 2ème série
  if GetControlText('BCOMPARE') = 'X' then
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

procedure TOF_MFOGRSTATREM.OnLoad;
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

procedure TOF_MFOGRSTATREM.OnArgument(S: string);
var Ind: Integer;
  Stg, Champ, Valeur: string;
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
  {$IFDEF EAGLCLIENT}
  SetControlVisible('BGraph', False);
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnClose :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.OnClose;
begin
  inherited;
  // Suppression de la TOB des lignes à traiter
  if TOBGraph <> nil then TOBGraph.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FChartGetLegendText : indique le texte à afficher comme légende
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.FChartGetLegendText(Sender: TCustomAxisPanel; LegendStyle: TLegendStyle; Index: Integer; var LegendText: string);
var TOBL: TOB;
  Stg: string;
  Ind, NbF, NoFam: Integer;
begin
  inherited;
  Stg := '';
  if GetControlText('BFAMILLE') = 'X' then
  begin
    // si un seul niveau de famille est choisi on affiche le libellé sion on affiche le code
    NbF := 0;
    NoFam := 0;
    for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
      begin
        NoFam := Ind;
        Inc(NbF);
      end;
    if NbF = 1 then
    begin
      TOBL := TOBGraph.FindFirst(['CODE'], [LegendText], False);
      if TOBL <> nil then Stg := Trim(TOBL.GetValue('FAMILLENIV' + IntToStr(NoFam)));
    end;
  end else
  begin
    TOBL := TOBGraph.FindFirst(['CODE'], [LegendText], False);
    if TOBL <> nil then Stg := Trim(TOBL.GetValue('LIBELLE'));
  end;
  if Stg <> '' then LegendText := Stg;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FListeDblClick : double click dans la grille
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.FListeDblClick(Sender: TObject);
var ValeurIndex, NoSerie: Integer;
  GS: THGrid;
begin
  GS := TFGRS1(Ecran).Fliste;
  ValeurIndex := GS.Row - GS.FixedRows;
  NoSerie := 0;
  if (GetControlText('BCOMPARE') = 'X') and (GS.Col > (GS.ColCount - 2)) then NoSerie := 1;
  TraiteZoom(ValeurIndex, NoSerie, TRUE);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FChartClickSeries : double click dans le graphe
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.FChartClickSeries(Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y:
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
    Stg := 'Code : ' + sCode;
    if GetControlText('BFAMILLE') = 'X' then
    begin
      for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
        begin
          sTitre := RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Ind), False);
          sFam := RechDom('GCFAMILLENIV' + IntToStr(Ind), FOExtract(sCode, 1, 3), False);
          Stg := Stg + '#10 ' + sTitre + '  : ' + sFam;
        end;
    end else
      if GetControlText('BCOLLECTION') = 'X' then
    begin
      if sCode = '' then sTitre := TraduireMemoire('<<Aucun>>')
      else sTitre := RechDom('GCCOLLECTION', sCode, False);
      Stg := Stg + '#10 ' + TraduireMemoire('Collection') + '  : ' + sTitre;
    end else
      if GetControlText('BVENDEUR') = 'X' then
    begin
      if sCode = '' then sTitre := TraduireMemoire('<<Aucun>>')
      else sTitre := RechDom('GCCOMMERCIAL', sCode, False);
      Stg := Stg + '#10 ' + TraduireMemoire('Vendeur') + '  : ' + sTitre;
    end else
      if GetControlText('BDEMARQUE') = 'X' then
    begin
      if sCode = '' then sTitre := TraduireMemoire('<<Aucun>>')
      else sTitre := RechDom('GCTYPEREMISE', sCode, False);
      Stg := Stg + '#10 ' + TraduireMemoire('Démarque') + '  : ' + sTitre;
    end;
    Stg := Stg + '#10 ' + TraduireMemoire('Chiffre d''affaires') + '  : '
      + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('CA'), 12, V_PGI.OkDecV, V_PGI.SymbolePivot, True);
    Stg := Stg + '#10 ' + TraduireMemoire('Montant de la remise') + '  : '
      + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('MNT'), 12, V_PGI.OkDecV, V_PGI.SymbolePivot, True);
    Stg := Stg + '#10 ' + TraduireMemoire('Pourcentage moyen') + '  : '
      + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('POURCENT'), 12, V_PGI.OkDecP, '', True) + '%';
    Stg := Stg + '#10 ' + TraduireMemoire('Quantité') + '  : '
      + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('QTE'), 12, V_PGI.OkDecQ, '', True);
    if NoSerie = 0 then Stg := Stg + '#10 ' + TraduireMemoire('Soit') + '  : ' + Series.MarkPercent(ValueIndex, False);
    if GetControlText('BCOMPARE') = 'X' then
    begin
      Stg := Stg + '#10 ' + TraduireMemoire('CA sans remise') + '  : '
        + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('CABRUT'), 12, V_PGI.OkDecV, V_PGI.SymbolePivot, True);
      Stg := Stg + '#10 ' + TraduireMemoire('Qté sans remise') + '  : '
        + StrfMontant(TOBGraph.Detail[ValueIndex].GetValue('QTBRUT'), 12, V_PGI.OkDecQ, '', True);
      if NoSerie = 1 then Stg := Stg + '#10 ' + TraduireMemoire('Soit') + '  : '
        + Series.MarkPercent(ValueIndex, False);
      if NoSerie = 0 then sTitre := TraduireMemoire('(avec remise)')
      else sTitre := TraduireMemoire('(sans remise)');
    end else sTitre := '';
    PGIInfo(Stg, ' ' + TraduireMemoire('Serie n°') + IntToStr(ValueIndex) + ' ' + sTitre);
  end else TraiteZoom(ValueIndex, NoSerie, FALSE);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteZoom : zoom sur les lignes constituant la série
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.TraiteZoom(ValeurIndex, NoSerie: Integer; FromGrid: Boolean);
var
  Ind: Integer;
  sRange, sCode, sArgs, sWhere, sCaisse: string;
  ///////////////////////////////////////////////////////////////////////////////////////
  procedure AjoutXXWhere(Stg: string);
  begin
    if sWhere <> '' then sWhere := sWhere + ' and ';
    sWhere := sWhere + Stg;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
begin
  if (ValeurIndex < 0) or (ValeurIndex > (TFGRS1(Ecran).FListe.RowCount - TFGRS1(Ecran).FListe.FixedRows)) then Exit;
  sRange := '';
  sArgs := '';
  sWhere := '';
  // Recherche des codes familles, collection, démarques ou vendeur
  sCode := '';
  if FromGrid then sCode := TFGRS1(Ecran).FListe.Cells[1, ValeurIndex + TFGRS1(Ecran).Fliste.FixedRows]
  else sCode := TOBGraph.Detail[ValeurIndex].GetValue('CODE');
  // Initialisation des critères de recherche
  sRange := 'GL_NATUREPIECEG='+ NaturePiece +';GL_TYPELIGNE=ART';
  AjoutXXWhere('GL_TYPEARTICLE IN ("MAR","NOM")');
  sCaisse := GetControlText('GL_CAISSE');
  if sCaisse <> '' then sRange := sRange + ';GL_CAISSE=' + sCaisse;
  if GetControlText('BFAMILLE') = 'X' then
  begin
    for Ind := 1 to 3 do if GetControlText('FAMILLENIV' + IntToStr(Ind)) = 'X' then
      begin
        if sCode = '' then
          {$IFDEF EAGLCLIENT}
          AjoutXXWhere('GL_FAMILLENIV' + IntToStr(Ind) + ' IN ("")')
            {$ELSE}
          AjoutXXWhere('GL_FAMILLENIV' + IntToStr(Ind) + '##""')
            {$ENDIF}
        else
          sRange := sRange + ';GL_FAMILLENIV' + IntToStr(Ind) + '=' + FOExtract(sCode, 1, 3);
      end;
  end else
    if GetControlText('BCOLLECTION') = 'X' then
  begin
    if sCode = '' then
      {$IFDEF EAGLCLIENT}
      AjoutXXWhere('GL_COLLECTION IN ("")')
        {$ELSE}
      AjoutXXWhere('GL_COLLECTION##""')
        {$ENDIF}
    else
      sRange := sRange + ';GL_COLLECTION=' + sCode;
  end else
    if GetControlText('BVENDEUR') = 'X' then
  begin
    if sCode = '' then
    begin
      {$IFDEF EAGLCLIENT}
      AjoutXXWhere('GL_REPRESENTANT IN ("")');
      {$ELSE}
      AjoutXXWhere('GL_REPRESENTANT##""');
      {$ENDIF}
      sArgs := 'AUCUNVENDEUR';
    end else sRange := sRange + ';GL_REPRESENTANT=' + sCode;
  end else
    if GetControlText('BDEMARQUE') = 'X' then
  begin
    if sCode = '' then
    begin
      {$IFDEF EAGLCLIENT}
      AjoutXXWhere('GL_TYPEREMISE IN ("")');
      {$ELSE}
      AjoutXXWhere('GL_TYPEREMISE##""');
      {$ENDIF}
      sArgs := 'AUCUNEDEMARQUE';
    end else sRange := sRange + ';GL_TYPEREMISE=' + sCode;
  end;
  if NoSerie = 0 then
  begin
    if (GetControlText('BREMONLY') = 'X') or (GetControlText('BCOMPARE') = 'X') then
    begin
      AjoutXXWhere('GL_TOTREMLIGNE<>0');
    end;
  end else
  begin
    {$IFDEF EAGLCLIENT}
    AjoutXXWhere('GL_TOTREMLIGNE IN (0)');
    {$ELSE}
    AjoutXXWhere('GL_TOTREMLIGNE##0');
    {$ENDIF}
  end;
  DateSelection(sRange, False);
  if sWhere <> '' then sRange := sRange + ';XX_WHERE=' + sWhere;
  // Lancement de la liste des lignes
  AGLLanceFiche('MFO', 'REMLIGNE_MUL', sRange, '', sArgs);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DateSelection : prend en compte les dates de sélection pour une clause where ou une clause range
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGRSTATREM.DateSelection(var Chaine: string; PourWhere: Boolean);
begin
  if PourWhere then
  begin
    if Trim(UpperCase(Chaine)) <> 'SELECT' then Chaine := Chaine + ' and ';
    Chaine := Chaine + 'GL_DATEPIECE>="' + USDateTime(StrToDate(GetControlText('DATEPIECE'))) + '" and '
      + 'GL_DATEPIECE<="' + USDateTime(StrToDate(GetControlText('DATEPIECE_'))) + '"';
  end else
  begin
    if Chaine <> '' then Chaine := Chaine + ';';
    Chaine := Chaine + 'GL_DATEPIECE=' + GetControlText('DATEPIECE')
      + ';GL_DATEPIECE_=' + GetControlText('DATEPIECE_');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOImprimeStatRem = lance l'impression des statistiques sur
Suite ........ : les remises en mode ticket depuis le script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure ImprimeStatRem(Parms: array of variant; Nb: integer);
var FF: TForm;
  TOBG, TOBL: TOB;
begin
  FF := TForm(Longint(Parms[0]));
  if not (FF is TFGRS1) then Exit;
  if not (TFGRS1(FF).LaTOF is TOF_MFOGRSTATREM) then Exit;
  TOBG := TOF_MFOGRSTATREM(TFGRS1(FF).LaTOF).TOBGraph;
  if (TOBG = nil) or (TOBG.Detail.Count <= 0) then Exit;
  TOBL := TOBG.Detail[0];
  if TOBL = nil then Exit;
  // Le titre de l'état est ajouté au 1er élément de la TOB
  TOBL.AddChampSupValeur('TITRE', TOF_MFOGRSTATREM(TFGRS1(FF).LaTOF).TitresCrit);
  TOBL.AddChampSupValeur('BORNEDATE', TOF_MFOGRSTATREM(TFGRS1(FF).LaTOF).TitresDate);
  // Lancement de l'impression
  FOLanceImprimeLP(efoStatRem, '', False, TOBG);
end;

initialization
  RegisterClasses([TOF_MFOGRSTATREM]);
  RegisterAglProc('FOImprimeStatRem', TRUE, 0, ImprimeStatRem);
end.
