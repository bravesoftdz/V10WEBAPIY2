{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 21/08/2002
Modifié le ... :   /  /
Description .. : Unit de consultation des rémunérations des paies
Suite ........ : précédantes dans la saisie des primes
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
unit UTofSaisPrim_Rem;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}HDB, DBCtrls, Fe_Main, DBGrids,
{$ELSE}
  MaineAgl,
{$ENDIF}
  uPaieRemunerations, ParamSoc,
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Vierge, P5Util, P5Def, AGLInit;
type
  TOF_PGSaisPrim_Rem = class(TOF)
  private
    CodeSal, Etab, DateD, DateF: string;
    DD, DF: TDateTime;
    Grille: THGrid;
    VCbxLRub: THValComboBox;
    procedure RechHisto(Sender: TObject);
//    procedure GetCellCanvas(Acol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation

{procedure TOF_PGSaisPrim_Rem.GetCellCanvas(Acol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = Grille.RowCount - 2 then
  begin
    Grille.RowHeights[ARow] := 2;
    exit;
  end;
  if ARow <> Grille.RowCount - 1 then exit; // On ne fait rien car on n'est pas positionné sur la dernière ligne affichant les cumuls
  Grille.canvas.Font.Height := 13;
  if ACol = 1 then Grille.canvas.Font.Style := Grille.canvas.Font.Style + [fsItalic] + [fsBold]
  else Grille.canvas.Font.Style := Grille.canvas.Font.Style + [fsItalic];
  Grille.Canvas.Brush.Color := Grille.FixedColor;
  Grille.Canvas.Font.Color := Grille.Font.Color;
end;              }

procedure TOF_PGSaisPrim_Rem.OnArgument(Arguments: string);
var
  i, ind: Integer;
  T1: TOB;
  st, Nom, lib: string;
  CEG: Boolean;
  VCbxLann: THValComboBox;
begin
  inherited;
  st := Trim(Arguments);
  CodeSal := ReadTokenSt(st); // Recup code Salarie
  Etab := ReadTokenSt(st); // Recup Code Etablissement
  DateD := ReadTokenSt(st); // Recup Date debut et date de fin de la session de paie
  DateF := ReadTokenSt(st);

  if st <> '' then Nom := ReadTokenSt(st);
  if Nom <> '' then
  begin
    Ecran.Caption := Ecran.Caption + Nom;
    UpdateCaption(Ecran);
  end;
  Lib := ReadTokenSt(st); // Recup du libelle de la colonne
  //RendDateExerSocial (StrToDate(DateD),StrToDate(DateF), DD,DF);
  DD := StrToDate(DateD);
  DF := StrToDate(DateF);
  Grille := THGrid(GetControl('GRILLEREM'));
  if Grille <> nil then
  begin
    Grille.CacheEdit;
  end
  else Grille.Enabled := FALSE;
  VCbxLRub := THValComboBox(GetControl('VALCBXRUB')); // Controle sur la combo contenant les codes des rubriques de bases
  if VCbxLRub <> nil then
  begin // il convient de récupèrer LaTob pour alimenter les combos Au maxi 5 à 6 Rubriques
    VCbxLRub.clear;
    ind := 0;
    for I := 0 to LaTOB.Detail.Count - 1 do
    begin
      T1 := LaTOB.Detail[I];
      VCbxLRub.Items.Add(T1.GetValue('LIBELBASE'));
      VCbxLRub.Values.Add(T1.GetValue('CODEBASE'));
      if lib = T1.GetValue('LIBELBASE') then ind := i;
    end;
    Ceg := GetParamSocSecur('SO_IFDEFCEGID', FALSE);
    if CEG then
    begin
      VCbxLRub.Items.Add('Prime Exceptionnelle');
      VCbxLRub.Values.Add('2305');
    end;
    VCbxLRub.ItemIndex := Ind;
    VCbxLRub.OnChange := RechHisto;
  end;
  VCbxLAnn := THValComboBox(GetControl('VALCBXANN')); // Controle sur la combo contenant les années
  if VCbxLAnn <> nil then
  begin
    VCbxLAnn.ItemIndex := 0;
    VCbxLAnn.OnChange := RechHisto;
  end;
  Grille.ColAligns[0] := taCenter;
  for i := 1 to Grille.ColCount do Grille.ColAligns[i] := taRightJustify;
  RechHisto(nil);
end;

procedure TOF_PGSaisPrim_Rem.RechHisto(Sender: TObject);
var
  Q: TQuery;
  St, Rub: string;
  i, j, NbDec, NbTaux, NbCoeff, NbMt: Integer;
  T2: TOB;
  CC: Double;
  Lbl4: THLabel;
  LaDateD, LaDateF: TDateTime;
begin
  Rub := VCbxLRub.Value;
  CC := 0;

{  St := 'SELECT PHB_DATEDEBUT,PHB_DATEFIN,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,PHB_MTREM' +
    ' FROM HISTOBULLETIN WHERE PHB_SALARIE="' + CodeSal + '" AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="' + Rub +
    '" AND PHB_DATEDEBUT>="' + UsDateTime(DD) + '" AND PHB_DATEFIN<="' + UsDateTime(DF) + '" ORDER BY PHB_DATEFIN';
  Q := OpenSQL(St, TRUE);
  if Q.Eof then
  begin // @@@@@ On se repositionne sur Année -1.
    FERME(Q);}
  if (GetControlText('VALCBXANN') = '1') then
  begin
    LaDateD := PLUSMOIS(DD, -12);
    LaDateF := PLUSMOIS(DF, -12);
  end
  else
  begin
    LaDateD := DD;
    LaDateF := DF;
  end;
  St := 'SELECT PHB_DATEDEBUT,PHB_DATEFIN,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,PHB_MTREM' +
    ' FROM HISTOBULLETIN WHERE PHB_SALARIE="' + CodeSal + '" AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="' + Rub +
    '" AND PHB_DATEDEBUT>="' + UsDateTime(LaDateD) + '" AND PHB_DATEFIN<="' + UsDateTime(LaDateF) + '" ORDER BY PHB_DATEFIN DESC';
  Q := OpenSQL(St, TRUE);
//  end;

  T2 := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], TRUE);
  // A mettre les bonnes valeurs du nbre de décimales des taux si besoin
  NbDec := 2;
  NbTaux := 2;
  NbCoeff := 2;
  NbMt := 2;

  // RAZ de la grille à chaque changement de rubrique
  for i := 1 to Grille.RowCount - 1 do
  begin
    for j := 0 to Grille.ColCount - 1 do
      Grille.Cells[j, i] := '';
  end;

  Grille.RowCount := 2;
  if T2 <> nil then
  begin
    NbDec := T2.GetValue('PRM_DECBASE');
    NbTaux := T2.GetValue('PRM_DECTAUX');
    NbCoeff := T2.GetValue('PRM_DECCOEFF');
    NbMt := T2.GetValue('PRM_DECMONTANT');
  end;
  i := 1;
  if Q.EOF then
  begin
    for i := 0 to 3 do Grille.Cells[i, 1] := '';
    Lbl4 := THLabel(GetControl('Lbl4'));
    if Lbl4 <> nil then Lbl4.Caption := DoubleToCell(0, NbMt);
    Grille.Cells[4, 1] := 'Aucune rémunération';
    Ferme(Q);
    exit;
  end;

  while not Q.EOF do
  begin
    Grille.Cells[0, i] := FormatDatetime('mmmm yyyy', Q.FindField('PHB_DATEFIN').AsFloat);
    CC := CC + Q.FindField('PHB_MTREM').AsFloat;
    if Q.FindField('PHB_BASEREM').AsFloat <> 1 then
    begin
      Grille.Cells[1, i] := DoubleToCell(Q.FindField('PHB_BASEREM').AsFloat, NbDec); // Base
      Grille.Cells[2, i] := DoubleToCell(Q.FindField('PHB_TAUXREM').AsFloat, NbTaux); // Taux
    end
    else
    begin
      Grille.Cells[1, i] := '';
      Grille.Cells[2, i] := '';
    end;
    Grille.Cells[3, i] := DoubleToCell(Q.FindField('PHB_COEFFREM').AsFloat, NbCoeff); // Coeff
    Grille.Cells[4, i] := DoubleToCell(Q.FindField('PHB_MTREM').AsFloat, NbMt); // Montant
    Q.Next;
    i := i + 1;
    Grille.RowCount := Grille.RowCount + 1;
  end;
  Ferme(Q);
  Lbl4 := THLabel(GetControl('Lbl4'));
  if Lbl4 <> nil then Lbl4.Caption := DoubleToCell(CC, NbMt);
end;

initialization
  registerclasses([TOF_PGSaisPrim_Rem]);
end.

