{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Statistiques Répartition par mode de paiement
Mots clefs ... : FO
*****************************************************************}
unit GrStaMR;

interface

uses
  GRS1, Hqry, StdCtrls, Hctrls, Mask, Dialogs, HSysMenu, Classes,
  hmsgbox, Menus, ComCtrls, TeeProcs, TeEngine, Chart, Grids, ExtCtrls,
  SysUtils, Forms, Series, Controls,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  FE_Main, db, dbtables,
  {$ENDIF}
  HTB97, HPanel, UTOB, HEnt1, HStatus, UiUtil;

procedure StatparModeReglement(Inside: THPanel; NaturePiece: string);

type
  TFGRStatMR = class(TFGRS1)
    TGP_DATEPIECE: THLabel;
    TGP_DATEPIECE_: THLabel;
    GP_DATEPIECE: THCritMaskEdit;
    GP_DATEPIECE_: THCritMaskEdit;
    Q1: THQuery;
    GP_CAISSE: THValComboBox;
    TGP_CAISSE: THLabel;
    XX_WHERE: TEdit;
    TGP_NUMZCAISSE: THLabel;
    GP_NUMZCAISSE: THCritMaskEdit;
    TGP_NUMZCAISSE_: THLabel;
    GP_NUMZCAISSE_: THCritMaskEdit;
    GVOIR: TGroupBox;
    BMONTANT: TRadioButton;
    BNOMBRE: TRadioButton;
    GETIQ: TGroupBox;
    BPOURCENT: TRadioButton;
    BVALEUR: TRadioButton;
    GP_ETABLISSEMENT: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject); override;
    procedure FChartClickSeries(Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GP_DATEPIECEChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FChart1GetLegendText(Sender: TCustomAxisPanel;
      LegendStyle: TLegendStyle; Index: Integer; var LegendText: string);
  private
    FCriteresChange: boolean;
    FCodes: string;
    FNaturePiece: string;
    procedure DoGraph;
    procedure TraiteZoom(ValeurIndex, Col: integer; FromGrid: boolean);
  public
    { Déclarations publiques }
  end;

implementation

uses
  FOUtil, MC_Lib;

const
  I_MR = 0;
  I_LIBELLE = 1;
  I_ENCNBR = 2;
  I_ENCMNT = 3;

  {$R *.DFM}
  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 11/12/2002
  Modifié le ... : 11/12/2002
  Description .. : Statistiques Répartition par mode de paiement
  Mots clefs ... :
  *****************************************************************}

procedure StatparModeReglement(Inside: THPanel; NaturePiece: string);
var XX: TFGRStatMR;
begin
  XX := TFGRStatMR.Create(Application);
  XX.InitGR('FOGRSTAMR');
  XX.FNaturePiece := NaturePiece;
  if Inside = nil then
  begin
    try
      XX.ShowModal;
    finally
      XX.Free;
    end;
  end else
  begin
    InitInside(XX, Inside);
    XX.Show;
  end;
  SourisNormale;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : FormShow
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.FormShow(Sender: TObject);
begin
  inherited;
  GP_CAISSE.Value := FOCaisseCourante;
  FOPositionneCaisseUser(GP_CAISSE, GP_ETABLISSEMENT);
  FListe.FixedRows := 1;
  FListe.VidePile(True);
  FListe.Cells[I_MR, 0] := TraduireMemoire('Mode de paiement');
  FListe.Cells[I_LIBELLE, 0] := TraduireMemoire('Libellé');
  Fliste.Cells[I_ENCNBR, 0] := TraduireMemoire('Nombre d''encaissements');
  FListe.ColAligns[I_ENCNBR] := taRightJustify;
  Fliste.Cells[I_ENCMNT, 0] := TraduireMemoire('Montant encaissé');
  FListe.ColAligns[I_ENCMNT] := taRightJustify;
  FCriteresChange := True;
  HMTrad.ResizeGridColumns(FListe);
  {$IFDEF EAGLCLIENT}
  BGraph.Visible := False;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : BValiderClick
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.BValiderClick(Sender: TObject);
var col1: integer;
  sTitre: string;
begin
  if FCriteresChange then DoGraph;
  FCriteresChange := False;
  Col1 := I_ENCNBR + ((I_ENCMNT - I_ENCNBR) * Ord(BMONTANT.Checked));
  UpdateGraph(I_MR, 0, 1, FListe.RowCount - 1, False, [Col1], [], TBarSeries);
  // Légende Titre et ordonnée
  if BMONTANT.Checked then
    sTitre := RechDom('TTDEVISE', V_PGI.DevisePivot, False)
  else
    sTitre := BNOMBRE.Caption;
  FChart1.Title.Text[0] := Caption + ' ' + Trim(sTitre);
  FChart1.Title.Text.Add(TraduireMemoire('du') + ' ' + GP_DATEPIECE.Text
    + ' ' + TraduireMemoire('au') + ' ' + GP_DATEPIECE_.Text);
  FChart1.LeftAxis.Title.Font.Color := FChart1.Title.Font.Color;
  FChart1.Legend.visible := True;
  FChart1.AxisVisible := True;
  // Définition des étiquettes
  FChart1.Series[0].ColorEachPoint := True;
  FChart1.Legend.LegendStyle := lsValues;
  FChart1.Legend.TextStyle := ltsPlain;
  FChart1.BottomAxis.LabelStyle := talText;
  FChart1.Series[0].ValueFormat := '# ##0';
  FChart1.Series[0].PercentFormat := '##0.00%';
  if BVALEUR.Checked then
    FChart1.Series[0].Marks.Style := smsValue
  else
    FChart1.Series[0].Marks.Style := smsPercent;
  baffGraph.Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Fabrication du graphe
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.DoGraph;
var Ind: integer;
  SQL, FmtQ, FmtV, MR: string;
  QQ: TQuery;
begin
  FListe.VidePile(TRUE);
  FmtQ := StrfMask(0, '', TRUE);
  FmtV := StrfMask(V_PGI.OkDecV, '', TRUE);
  SQL := 'SELECT GPE_MODEPAIE,SUM(GPE_MONTANTECHE) AS MONT,COUNT(GPE_MONTANTECHE) AS NBRE'
    + ' FROM GCREGLEMENTFO ';
  Q1.UpdateCriteres;
  if Q1.Criteres <> '()' then SQL := SQL + ' WHERE ' + Q1.Criteres;
  SQL := SQL + ' GROUP BY GPE_MODEPAIE';
  QQ := OpenSQL(SQL, TRUE);
  Initmove(RecordsCount(QQ), '');
  FCodes := '';
  while not QQ.Eof do
  begin
    if Movecur(FALSE) then ;
    MR := QQ.FindField('GPE_MODEPAIE').asString;
    MR := Format_String(MR, 3);
    Ind := Pos(';' + MR + ';', ';' + FCodes);
    if Ind <= 0 then
    begin
      if Trim(FListe.Cells[I_MR, FListe.RowCount - 1]) <> '' then
        Fliste.RowCount := FListe.RowCount + 1;
      FCodes := FCodes + MR + ';';
      Ind := FListe.RowCount - 1;
      FListe.Cells[I_MR, Ind] := Trim(MR);
      FListe.Cells[I_LIBELLE, Ind] := RechDom('GCMODEPAIE', Trim(MR), FALSE);
    end else
    begin
      Ind := (Ind - 1) div 4 + Fliste.FixedRows;
    end;
    FListe.Cells[I_ENCNBR, Ind] := FormatFloat(FmtQ, ValeurI(Fliste.Cells[I_ENCNBR, Ind]) + QQ.FindField('NBRE').asInteger);
    FListe.Cells[I_ENCMNT, Ind] := FormatFloat(FmtV, Valeur(Fliste.Cells[I_ENCMNT, Ind]) + QQ.FindField('MONT').asFloat);
    QQ.Next;
  end;
  FiniMove;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : BValiderClick
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.FChartClickSeries(Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y:
  Integer);
var Ind: integer;
  SerieName, sCode, Stg: string;
begin
  SerieName := Series.Name;
  if (Length(SerieName) <= 2) or (Copy(SerieName, 1, 2) <> 'SS') then Exit;
  Ind := ValeurI(Copy(SerieName, 3, Length(SerieName)));
  if ssRight in shift then
  begin
    sCode := Trim(gtfs(FCodes, ';', ValueIndex + 1));
    Stg := TraduireMemoire('Code') + ' : ' + sCode;
    Stg := Stg + '#10 ' + TraduireMemoire('Mode de paiement')
      + '  : ' + RechDom('GCMODEPAIE', Trim(sCode), FALSE);
    for Ind := FListe.FixedRows to FListe.RowCount do
      if FListe.cells[I_MR, Ind] = sCode then
      begin
        Stg := Stg + '#10 ' + TraduireMemoire('Nombre') + '  : '
          + StrfMontant(Valeur(FListe.Cells[I_ENCNBR, Ind]), 12, V_PGI.OkDecQ, '', True);
        Stg := Stg + '#10 ' + TraduireMemoire('Montant') + '  : '
          + StrfMontant(Valeur(FListe.Cells[I_ENCMNT, Ind]), 12, V_PGI.OkDecV, V_PGI.SymbolePivot, True);
        Break;
      end;
    Stg := Stg + '#10 ' + TraduireMemoire('Soit') + '  : ' + Series.MarkPercent(ValueIndex, False);
    PGIInfo(Stg, ' ' + TraduireMemoire('Serie n°') + IntToStr(ValueIndex));
  end else
  begin
    TraiteZoom(ValueIndex, Ind + 1, FALSE);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Zoom sur un élément du graphe
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.TraiteZoom(ValeurIndex, Col: integer; FromGrid: boolean);
var St, Mr: string;
begin
  if (ValeurIndex < 0) or (ValeurIndex > FListe.RowCount - FListe.FixedRows) then Exit;
  if FromGrid then
    Mr := FListe.Cells[I_MR, ValeurIndex + Fliste.FixedRows]
  else
    Mr := Trim(gtfs(FCodes, ';', ValeurIndex + 1));
  // Critères
  St := 'GP_NATUREPIECEG=' + FNaturePiece
      + ';XX_WHERE=GP_NUMZCAISSE>0;GP_CAISSE=' + GP_CAISSE.Value
      + ';GP_DATEPIECE=' + GP_DATEPIECE.Text
      + ';GP_DATEPIECE_=' + GP_DATEPIECE_.Text
      + ';GPE_MODEPAIE=' + Mr
      + ';GP_NUMZCAISSE=' + GP_NUMZCAISSE.Text
      + ';GP_NUMZCAISSE_=' + GP_NUMZCAISSE_.Text;
  AGLLanceFiche('MFO', 'PIEDECHE_MUL', St, '', 'STATISTIQUES');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : GP_DATEPIECEChange
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.GP_DATEPIECEChange(Sender: TObject);
begin
  inherited;
  FCriteresChange := TRUE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : FListeDblClick
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.FListeDblClick(Sender: TObject);
begin
  inherited;
  TraiteZoom(FListe.Row - FListe.FixedRows, -1, TRUE);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : FChart1GetLegendText
Mots clefs ... :
*****************************************************************}

procedure TFGRStatMR.FChart1GetLegendText(Sender: TCustomAxisPanel;
  LegendStyle: TLegendStyle; Index: Integer; var LegendText: string);
begin
  inherited;
  LegendText := RechDom('GCMODEPAIE', trim(LegendText), FALSE);
end;

end.
