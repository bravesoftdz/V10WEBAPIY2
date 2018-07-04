{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Statistiques Meilleure journée / Meilleure heure
Mots clefs ... : FO
*****************************************************************}
unit GrJH;

interface

uses
  GRS1, Hqry, StdCtrls, Hctrls, ExtCtrls, Mask, Grids, Controls, HTB97, Classes,
  Dialogs, HSysMenu, hmsgbox, Menus, ComCtrls, TeeProcs, TeEngine, Chart,
  Forms, Series, SysUtils,
  {$IFDEF EAGLCLIENT}
  Maineagl, UtileAGL,
  {$ELSE}
  FE_Main, Db, DBTables,
  {$ENDIF}
  Hent1, UiUtil, HStatus, HPanel, UTOB;

procedure MeuilleurJouretHour(Inside: THPanel; NaturePiece: string);

type
  TFGRJ_H = class(TFGRS1)
    TGP_DATEPIECE_: THLabel;
    TGP_DATEPIECE: THLabel;
    TGP_CAISSE: THLabel;
    TGP_REPRESENTANT: THLabel;
    GP_DATEPIECE_: THCritMaskEdit;
    GP_DATEPIECE: THCritMaskEdit;
    GRJOUR: TRadioGroup;
    Q1: THQuery;
    XX_WHERE: TEdit;
    GP_REPRESENTANT: THValComboBox;
    GP_CAISSE: THValComboBox;
    FTypePalmares: THValComboBox;
    TFTypePalmares: THLabel;
    GRCHOIXJOUR: TGroupBox;
    CBLUNDI: TCheckBox;
    CBMARDI: TCheckBox;
    CBMERCREDI: TCheckBox;
    CBJEUDI: TCheckBox;
    CBVENDREDI: TCheckBox;
    CBSAMEDI: TCheckBox;
    CBDIMANCHE: TCheckBox;
    GP_ETABLISSEMENT: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject); override;
    procedure GP_DATEPIECEChange(Sender: TObject);
    procedure FChart1ClickSeries(Sender: TCustomChart;
      Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FListeDblClick(Sender: TObject);
    procedure CHOIXJOURClick(Sender: TObject);
  private
    FChangeCriteres: boolean;
    FCodes: string;
    FNaturePiece: string;
    procedure DoGraph;
    function DonneValeurChoisie(ValueIndex: integer; FromGrid: boolean): integer;
    procedure TraiteZoom(ValueIndex: integer; FromGrid: boolean);
    function VerifieChoixDuJour(TOBLigne: TOB): boolean;
  public
    { Déclarations publiques }
  end;

implementation
{$R *.DFM}

uses
  FOUtil, MC_Lib;

const
  I_CODE = 0;
  I_LIBE = 1;
  I_NBRE = 2;
  I_CAFF = 3;
  I_QTE = 4;
  I_PCA = 5;
  I_PQT = 6;

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 11/12/2002
  Modifié le ... : 11/12/2002
  Description .. : Statistiques Meilleure journée / Meilleure heure
  Mots clefs ... :
  *****************************************************************}

procedure MeuilleurJouretHour(Inside: THPanel; NaturePiece: string);
var XX: TFGRJ_H;
begin
  XX := TFGRJ_H.Create(Application);
  XX.InitGR('FOGRJH');
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
Description .. : Retourne la caption d'un radio bouton
Mots clefs ... :
*****************************************************************}

function LireItemRadioBox(RG: TCustomControl): string;
var j: Integer;
  RB: TRadioButton;
begin
  Result := '';
  if RG is TGroupBox then
  begin
    for j := 0 to RG.ControlCount - 1 do
      if RG.Controls[j] is TRadioButton then
      begin
        RB := TRadioButton(RG.Controls[j]);
        if RB.Checked then
        begin
          Result := RB.Caption;
          Break;
        end;
      end;
  end else Result := TRadioGroup(RG).Items[TRadioGroup(RG).ItemIndex];
  Result := VireSouligne(Result);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : FormShow
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.FormShow(Sender: TObject);
begin
  inherited;
  XX_WHERE.Text := 'GP_NATUREPIECEG=' + FOGetNatureTicket(False, True)
    + ' AND GP_NUMZCAISSE>0';
  GP_CAISSE.Value := FOCaisseCourante;
  FOPositionneCaisseUser(GP_CAISSE, GP_ETABLISSEMENT);
  FListe.FixedCols := 1;
  FListe.FixedRows := 1;
  FListe.VidePile(True);
  FLIste.Cells[I_CODE, 0] := '';
  FListe.ColWidths[I_CODE] := 0;
  FListe.Cells[I_LIBE, 0] := LireItemRadioBox(GRJour);
  FListe.ColAligns[I_LIBE] := taLeftJustify;
  Fliste.Cells[I_NBRE, 0] := TraduireMemoire('Nombre de pièces');
  FListe.ColAligns[I_NBRE] := taRightJustify;
  Fliste.Cells[I_CAFF, 0] := TraduireMemoire('Chiffre d''affaires');
  FListe.ColAligns[I_CAFF] := taRightJustify;
  Fliste.Cells[I_QTE, 0] := TraduireMemoire('Quantité');
  FListe.ColAligns[I_QTE] := taRightJustify;
  Fliste.Cells[I_PCA, 0] := TraduireMemoire('Panier moyen (Montant)');
  FListe.ColAligns[I_PCA] := taRightJustify;
  Fliste.Cells[I_PQT, 0] := TraduireMemoire('Panier moyen (Quantité)');
  FListe.ColAligns[I_PQT] := taRightJustify;
  HMTrad.ResizeGridColumns(FListe);
  GP_REPRESENTANT.ItemIndex := 0;
  FTypePalmares.ItemIndex := 0;
  FCHangeCriteres := True;
  {$IFDEF EAGLCLIENT}
  BGraph.Visible := False;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Ajoute le libellé d'un CheckBox au titre du graphe
Mots clefs ... :
*****************************************************************}

procedure AjoutTitre(Chkbx: TCheckBox; var Titre: string);
begin
  if Chkbx.Checked then
  begin
    if Titre = '' then Titre := ' (' else Titre := Titre + ', ';
    Titre := Titre + Chkbx.Caption;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Ajoute le libellé d'un CheckBox au titre du graphe
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.BValiderClick(Sender: TObject);
var Col, Ind: integer;
  Stg: string;
begin
  if FChangeCriteres then DoGraph;
  FListe.Cells[I_LIBE, 0] := LireItemRadioBox(GRJour);
  FChangeCriteres := False;
  Col := I_QTE;
  if FTypePalmares.value = 'NBR' then Col := I_NBRE else
    if FTypePalmares.value = 'CAF' then Col := I_CAFF else
    if FTypePalmares.value = 'QTE' then Col := I_QTE else
    if FTypePalmares.value = 'PCA' then Col := I_PCA else
    if FTypePalmares.value = 'PQT' then Col := I_PQT;
  UpdateGraph(I_LIBE, 0, 1, FListe.RowCount - 1, False, [Col], [], TBarSeries);
  if not ParamUtilisateur then
    for Ind := 0 to FChart1.SeriesList.Count - 1 do
      FChart1.Series[Ind].marks.Visible := False;
  // Légende Titre et ordonnée
  Stg := '';
  if GRJour.ItemIndex = 1 then
  begin
    if not (CBLUNDI.Checked) or not (CBMARDI.Checked) or not (CBMERCREDI.Checked) or
      not (CBJEUDI.Checked) or not (CBVENDREDI.Checked) or not (CBSAMEDI.Checked) or
      not (CBDIMANCHE.Checked) then
    begin
      AjoutTitre(CBLUNDI, Stg);
      AjoutTitre(CBMARDI, Stg);
      AjoutTitre(CBMERCREDI, Stg);
      AjoutTitre(CBJEUDI, Stg);
      AjoutTitre(CBVENDREDI, Stg);
      AjoutTitre(CBSAMEDI, Stg);
      AjoutTitre(CBDIMANCHE, Stg);
      if Stg <> '' then Stg := Stg + ')';
    end;
  end;
  FChart1.Title.Text[0] := LireItemRadioBox(GRJour) + Stg;
  FChart1.LeftAxis.Title.Font.Color := FChart1.Title.Font.Color;
  FChart1.Legend.visible := True;
  FChart1.AxisVisible := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Retourne le code d'une heure
Mots clefs ... :
*****************************************************************}

function ConvHeureToCode(Heure: TDateTime): integer;
var NH, H, M, S, MS: word;
  Val: integer;
begin
  DecodeTime(Heure, H, M, S, MS);
  Val := Round(Frac(EncodeTime(H, 0, 0, 0)) * 100);
  DecodeTime(Val / 100, NH, M, S, MS);
  if NH < H then Inc(Val);
  Result := Val;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/07/2003
Modifié le ... : 03/07/2003
Description .. : Recherche la ligne de la grille qui contient une valeur dans
Suite ........ : une colonne
Mots clefs ... : 
*****************************************************************}

function ChercheRow(Val: string; Fliste: THgrid; ColCode: Integer): Integer;
var Trouve: Boolean;
begin
  Result := Fliste.FixedRows;
  Trouve := FALSE;
  val := trim(Val);
  while (Result <= Fliste.rowcount - 1) and (not Trouve) do
  begin
    Trouve := ((trim(FLIste.Cells[COlCOde, Result]) = Val) or (trim(FLIste.Cells[ColCode, Result]) = ''));
    if not trouve then inc(Result);
  end;
  if not Trouve then
  begin
    FListe.RowCount := FListe.RowCount + 1;
    Result := Fliste.Rowcount - 1;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Fabrication du graphe
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.DoGraph;
var OkJour, Last: boolean;
  OldVal, Val, Row, Ind, No: integer;
  Y, M, D, H, S, MS: word;
  FmtV, FmtQ, Sql, Ch: string;
  Ca, Qt, Coef: double;
  TOBprinc, TOBL: TOB;
begin
  Fliste.VidePile(True);
  FmtQ := StrfMask(V_PGI.OkDecQ, '', True);
  FmtV := StrfMask(V_PGI.OkDecV, '', True);
  OldVal := -1;
  No := 0;
  Ca := 0;
  Qt := 0;
  OkJour := (GRJour.ItemIndex = 0);
  if OkJour then Ch := 'GP_DATEPIECE' else Ch := 'GP_HEURECREATION';
  // Select
  Sql := 'SELECT ' + Ch + ',GP_NATUREPIECEG,GP_CAISSE,GP_NUMERO,';
  Sql := Sql + 'GP_TOTALTTCDEV AS CAFFAIRE,GP_TOTALQTEFACT AS QTE';
  if not OkJour then Sql := Sql + ',GP_DATEPIECE';
  // From
  Sql := Sql + ' FROM PIECE ';
  // Where
  Q1.UpdateCriteres;
  if Q1.Criteres <> '()' then Sql := Sql + ' WHERE ' + Q1.Criteres;
  // Order By
  Sql := Sql + ' ORDER BY ' + Ch + ',GP_NATUREPIECEG,GP_CAISSE';
  // Sélection
  TOBPrinc := TOB.Create('', nil, -1);
  TOBPrinc.LoadDetailFromSQL(Sql);
  InitMove(TOBPrinc.Detail.Count, '');
  for Ind := 0 to TOBPrinc.Detail.Count do
  begin
    MoveCur(False);
    Last := (Ind = TOBPrinc.Detail.Count);
    Val := OldVal;
    TOBL := nil;
    if not Last then
    begin
      TOBL := TOBPrinc.Detail[Ind];
      if OkJour then
      begin
        DecodeDate(vDate(TOBL.GetValue('GP_DATEPIECE')), Y, M, D);
        Val := Round(DayOfWeek(EncodeDate(Y, M, D)));
        if Val = 1 then Val := 8;
      end else
      begin
        Val := ConvHeureToCode(vDate(TOBL.GetValue('GP_HEURECREATION')));
        if not VerifieChoixDuJour(TOBL) then Continue;
      end;
    end;
    if (OldVal <> Val) or (Last) then
    begin
      if OldVal <> -1 then
      begin
        Row := ChercheRow(IntToStr(OldVal), FListe, I_CODE);
        FListe.Cells[I_CODE, Row] := IntToStr(OldVal);
        FListe.Cells[I_NBRE, Row] := FormatFloat('#,##0', ValeurI(FListe.Cells[I_NBRE, Row]) + No);
        ;
        FListe.Cells[I_CAFF, Row] := FormatFloat(FmtV, Valeur(FListe.Cells[I_CAFF, Row]) + Ca);
        FListe.Cells[I_QTE, Row] := FormatFloat(FmtQ, Valeur(FListe.Cells[I_QTE, Row]) + Qt);
      end;
      OldVal := Val;
      No := 0;
      Ca := 0;
      Qt := 0;
    end;
    if (TOBL <> nil) and (not Last) then
    begin
      Coef := 1;
      Inc(No);
      Ca := Ca + (vDouble(TOBL.GetValue('CAFFAIRE')) * Coef);
      Qt := Qt + (vDouble(TOBL.GetValue('QTE')) * Coef);
    end;
  end;
  TOBPrinc.Free;
  FiniMove;
  if OldVal <> -1 then
  begin
    Fliste.SortGrid(I_CODE, False);
    FCodes := '';
    for Row := Fliste.FixedRows to Fliste.RowCount - 1 do
    begin
      Val := ValeurI(FListe.Cells[I_CODE, Row]);
      FCodes := FCodes + IntToStr(Val) + ';';
      if OkJour then
      begin
        if Val = 8 then Val := 1;
        FListe.Cells[I_LIBE, Row] := LongDayNames[Val];
      end else
      begin
        DecodeTime(Val / 100, H, M, S, ms);
        FListe.Cells[I_LIBE, Row] := FormatFloat('00', H) + ':00 - '
          + FormatFloat('00', H + 1 - 24 * Ord(H = 23)) + ':00';
      end;
      No := ValeurI(FListe.Cells[I_NBRE, Row]);
      if No <> 0 then
      begin
        FListe.Cells[I_PCA, Row] := FormatFloat(FmtV, Valeur(FListe.Cells[I_CAFF, Row]) / No);
        FListe.Cells[I_PQT, Row] := FormatFloat(FmtQ, Valeur(FListe.Cells[I_QTE, Row]) / No);
      end else
      begin
        FListe.Cells[I_PCA, Row] := FormatFloat(FmtV, 0);
        FListe.Cells[I_PQT, Row] := FormatFloat(FmtQ, 0);
      end;
    end;
  end;
  bAffGraph.Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Vérifie si la date d'une pièce fait partie des jours de la
Suite ........ : semaine à traiter
Mots clefs ... :
*****************************************************************}

function TFGRJ_H.VerifieChoixDuJour(TOBLigne: TOB): boolean;
var Val: integer;
  Y, M, D: Word;
begin
  Result := False;
  DecodeDate(vDate(TOBLigne.GetValue('GP_DATEPIECE')), Y, M, D);
  Val := Round(DayOfWeek(EncodeDate(Y, M, D)));
  case Val of
    1: Result := CBDIMANCHE.Checked;
    2: Result := CBLUNDI.Checked;
    3: Result := CBMARDI.Checked;
    4: Result := CBMERCREDI.Checked;
    5: Result := CBJEUDI.Checked;
    6: Result := CBVENDREDI.Checked;
    7: Result := CBSAMEDI.Checked;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : GP_DATEPIECEChange
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.GP_DATEPIECEChange(Sender: TObject);
begin
  inherited;
  FChangeCriteres := True;
  GRCHOIXJOUR.Visible := (GRJOUR.ItemIndex = 1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : CHOIXJOURClick
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.CHOIXJOURClick(Sender: TObject);
begin
  inherited;
  FChangeCriteres := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : FChart1ClickSeries
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.FChart1ClickSeries(Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  TraiteZoom(ValueIndex, False);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Retourne la valeur choisie pour le zoom
Mots clefs ... :
*****************************************************************}

function TFGRJ_H.DonneValeurChoisie(ValueIndex: integer; FromGrid: boolean): integer;
begin
  Result := 0;
  if (ValueIndex < 0) or (ValueIndex > FListe.RowCount - FListe.FixedRows) then Exit;
  if FromGrid then
    Result := ValeurI(FListe.Cells[I_CODE, ValueIndex + Fliste.FixedRows])
  else
    Result := ValeurI(gtfs(FCodes, ';', ValueIndex + 1));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : Zoom sur un élément du graphe
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.TraiteZoom(ValueIndex: integer; FromGrid: boolean);
var Val: integer;
  St, StDay, StClause: string;
  OkJour: boolean;
  H, H1, M, S, ms: word;
  //////////////////////////////////////////////////////////////////////////
  procedure AjouteStDay(Texte, Separateur: string);
  begin
    if Texte <> '' then
    begin
      if StDay <> '' then StDay := StDay + Separateur;
      StDay := StDay + Texte;
    end;
  end;
  //////////////////////////////////////////////////////////////////////////
begin
  if (ValueIndex < 0) or (ValueIndex > FListe.RowCount - FListe.FixedRows) then Exit;
  Val := DonneValeurChoisie(ValueIndex, FromGrid);
  OkJour := (GRJour.ItemIndex = 0);
  // Conditions speciales
  St := 'GP_NATUREPIECEG=' + FNaturePiece
      + ';GP_CAISSE=' + GP_CAISSE.Value + ';GP_DATEPIECE=' + GP_DATEPIECE.Text
      + ';GP_DATEPIECE_=' + GP_DATEPIECE_.Text;
  if GP_REPRESENTANT.ItemIndex > 0 then
    St := St + ';GP_REPRESENTANT=' + GP_REPRESENTANT.Value;
  if OkJour then
  begin
    {$IFDEF EAGLCLIENT}
    St := St + ';XX_WHERE=WDAY(GP_DATEPIECE) IN (' + IntToStr(Val) + ')';
    {$ELSE}
    St := St + ';XX_WHERE=WDAY(GP_DATEPIECE)##' + IntToStr(Val);
    {$ENDIF}
  end else
  begin
    DecodeTime(Val / 100, H, M, S, ms);
    if (ValueIndex + 1) <= (FListe.RowCount - FListe.FixedRows) then
    begin
      Val := DonneValeurChoisie(ValueIndex + 1, FromGrid);
      DecodeTime(Val / 100, H1, M, S, ms);
      Dec(H1);
    end else
    begin
      H1 := 24;
    end;
    // jours de  la semaine choisis
    StDay := '';
    if not (CBDIMANCHE.Checked) or not (CBLUNDI.Checked) or
      not (CBMARDI.Checked) or not (CBMERCREDI.Checked) or
      not (CBJEUDI.Checked) or not (CBVENDREDI.Checked) or
      not (CBVENDREDI.Checked) then
    begin
      if CBSAMEDI.Checked then
        if CBDIMANCHE.Checked then AjouteStDay('1', ',');
      if CBLUNDI.Checked then AjouteStDay('2', ',');
      if CBMARDI.Checked then AjouteStDay('3', ',');
      if CBMERCREDI.Checked then AjouteStDay('4', ',');
      if CBJEUDI.Checked then AjouteStDay('5', ',');
      if CBVENDREDI.Checked then AjouteStDay('6', ',');
      if CBSAMEDI.Checked then AjouteStDay('7', ',');
      if StDay <> '' then StDay := 'WDAY(GP_DATEPIECE) IN (' + StDay + ')';
    end;
    // heures choisies
    StClause := '';
    case V_PGI.Driver of
      dbINTRBASE: ;
      dbMSSQL: StClause := 'CONVERT(CHAR(2), GP_HEURECREATION, 8)';
      dbORACLE7,
        dbORACLE8: StClause := 'TO_CHAR(GP_HEURECREATION, "HH24")';
      dbDB2: ;
      dbINFORMIX: ;
      dbMSACCESS: StClause := 'HOUR(GP_HEURECREATION)';
      dbPARADOX: ;
      dbSQLANY: ;
      dbSQLBASE: ;
      dbPOL: ;
      dbSYBASE: StClause := 'DATEPART(HOUR, GP_HEURECREATION)';
      dbMySQL: ;
      dbPROGRESS: ;
    end;
    if StClause <> '' then
    begin
      StClause := 'CAST(' + StClause + ' AS INTEGER)';
      {$IFDEF EAGLCLIENT}
      StClause := StClause + ' IN (' + IntToStr(H) + ')';
      {$ELSE}
      StClause := StClause + '##' + IntToStr(H);
      {$ENDIF}
      AjouteStDay(StClause, ' AND ');
    end;
    if StDay <> '' then St := St + ';XX_WHERE=' + StDay;
  end;
  AGLLanceFiche('MFO', 'PIECE_MUL', St, '', 'TICKET');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/12/2002
Modifié le ... : 11/12/2002
Description .. : FListeDblClick
Mots clefs ... :
*****************************************************************}

procedure TFGRJ_H.FListeDblClick(Sender: TObject);
begin
  TraiteZoom(Fliste.Row - FLIste.FixedRows, True);
end;

end.
