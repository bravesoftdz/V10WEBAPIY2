unit UTofTarifMAJMode;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, Hdb, M3FP, UtilGC, HTB97, UTOB,
  {$IFDEF EAGLCLIENT}
  emul, MaineAgl,
  {$ELSE}
  dbTables, db, DBGrids, mul, Fe_Main,
  {$ENDIF}
  HDimension, Entgc, utilPGI, utilArticle, utilDimArticle, AglInit, HStatus, Hqry,
  AssistMajTarfMode;
type
  TOF_TarifMAJMode = class(TOF)
    procedure OnArgument(st: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure MajTarifArt;
    procedure ChargeLesArticles;
  end;

var
  bModifCritAuto: boolean;
  bOldStatutGen, bOldStatutUni, bOldStatutDim: boolean; // Anciennes valeurs des checkBox
  TOBArticle: TOB;
  NatureType: string; // Utilisé dans la mise à jour des tarif ACH ouVTE

const
  // libellés des messages
  TexteMessage: array[1..3] of string = (
    {1}'Les articles dimensionnés n''étant pas sélectionnés, les critères dimensions sont ignorés'
    {2}, 'Attention'
    {3}, 'Aucun élément sélectionné'
    );

procedure AGLTarfArtMaj_mul_Maj(parms: array of variant; nb: integer);

implementation

{Les critères de type "GA_FOURNPRINC=001" passé en parametre après le mot clef EXCLUSIF
 deviennent disable}

procedure TOF_TarifMAJMode.OnArgument(st: string);
var i_ind, iCol: integer;
  THLIB: THLabel;
  BOUT: TToolbarButton97;
  FF: TFMUL;
  Nbr: integer;
  Critere, ChampMul, ValMul: string;
  CC: TComponent;
  bExclusif: boolean;
  {$IFDEF EAGLCLIENT}
  TypeArticle: THValComboBox;
  {$ELSE}
  TypeArticle: THDBValComboBox;
  {$ENDIF}

begin
  inherited;

  bModifCritAuto := False;

  // Paramétrage des libellés des familles, stat. article et dimensions
  for iCol := 1 to 3 do ChangeLibre2('TGA_FAMILLENIV' + InttoStr(iCol), Ecran);
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
  begin
    for iCol := 4 to 8 do ChangeLibre2('TGA2_FAMILLENIV' + InttoStr(iCol), Ecran);
    for iCol := 1 to 2 do ChangeLibre2('TGA2_STATART' + InttoStr(iCol), Ecran);
  end;
  // Paramétrage des libellés des tables libres
  Nbr := 0;
  if (GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False);

  if (ctxMode in V_PGI.PGIContexte) or True then
  begin
    FF := TFMul(Ecran);
    repeat
      Critere := uppercase(Trim(ReadTokenSt(St)));
      if (Critere = 'GCARTICLE') then
      begin
        {$IFDEF EAGLCLIENT}
        TypeArticle := THValComboBox(GetControl('GA_TYPEARTICLE'));
        {$ELSE}
        TypeArticle := THDBValComboBox(GetControl('GA_TYPEARTICLE'));
        {$ENDIF}
        TypeArticle.Plus := 'AND ((CO_CODE <>"PRE") AND (CO_CODE<>"FI"))';
        TypeArticle.Value := 'MAR'; // Initialisation par défaut
      end;
      if Critere <> '' then
      begin
        i_ind := pos('=', Critere);
        if i_ind > 0 then
        begin
          ChampMul := copy(Critere, 1, i_ind - 1);
          ValMul := copy(Critere, i_ind + 1, length(Critere));
          if ChampMul = 'TYPE' then NatureType := ValMul else
          begin
            CC := FF.FindComponent(ChampMul);
            if (CC is THEdit) then THEdit(CC).Text := ValMul;
            if (CC is THValComboBox) then THValComboBox(CC).Value := ValMul;
            if (CC is TCheckBox) then
            begin
              if (ValMul = 'X') then TCheckBox(CC).state := cbChecked;
              if (ValMul = '-') then TCheckBox(CC).state := cbUnChecked;
              if (ValMul = ' ') then TCheckBox(CC).state := cbGrayed;
            end;
            if (CC is THLabel) then THLabel(CC).caption := ValMul;
            SetControlProperty(CC.Name, 'Tag', 0);
          end;
        end;
      end;
    until Critere = '';

    // Mise en forme des libellés des dimensions CT:29/08/02
    for iCol := 1 to MaxDimension do
    begin
      THLIB := THLabel(GetControl('DIMENSION' + InttoStr(iCol)));
      THLIB.Caption := RechDom('GCCATEGORIEDIM', 'DI' + InttoStr(iCol), False);
      if THLIB.Caption = '.-'
        then
      begin
        THLIB.Visible := False;
        THLIB.FocusControl.Visible := False;
      end
        // Annule l'alignement automatique des champs effectués par l'AGL
      else
        if (THLIB.FocusControl <> nil) and (THLIB.Left = (THLIB.FocusControl).Left) then THLIB.Top := (THLIB.FocusControl).Top - 17;
    end;

    // Mise en forme des libellés des dates, booléans libres et montants libres
    if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr);
    if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr);
    if (GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr);
    if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False);
  end; // if (ctxMode in V_PGI.PGIContexte) then

  if (ctxFO in V_PGI.PGIContexte) then
  begin
    // Suppression des boutons d'insertion et de duplication
    SetControlVisible('BINSERT', False);
    SetControlVisible('BINSERTNOMENC', False);
    SetControlVisible('B_DUPLICATION', False);
  end;
end;

procedure TOF_TarifMAJMode.OnLoad;
var F: TFMul;
  CC: TCheckBox;
  CB: TCheckBox;
  iControl: integer;
  bAffiche: boolean;
  ctrl, xx_where: string;
begin
  inherited;
  if not (Ecran is TFMul) then exit;
  if not (ctxMode in V_PGI.PGIContexte) then exit;

  xx_where := '';
  F := TFMul(Ecran);

  F.Q.Manuel := True; // Evite l'exécution de la requête lors de la maj de Q.Liste
  if TCheckBox(F.FindComponent('STATUT_DIM')).State = cbChecked
    then F.Q.Liste := 'GCMULARTDIM_MODE'
  else F.Q.Liste := 'GCMULARTICLE_MODE';

  /// Gestion des checkBox du statut article
  for iControl := 1 to 3 do
  begin
    if iControl = 1 then ctrl := 'GEN'
    else
      if iControl = 2 then ctrl := 'UNI'
    else ctrl := 'DIM';
    CC := TCheckBox(TFMul(F).FindComponent('STATUT_' + ctrl));
    if (CC <> nil) and (CC.State = cbChecked) then
    begin
      if xx_where <> '' then xx_where := xx_where + ' or ' else xx_where := '(';
      xx_where := xx_where + 'GA_STATUTART="' + ctrl + '"';
    end;
  end;
  if xx_where <> '' then xx_where := xx_where + ')';

  // Gestion des checkBox : tenu en stock, commissionnable, remise ligne et remise pied
  for iControl := 1 to 4 do
  begin
    if iControl = 1 then ctrl := 'GA_TENUESTOCK'
    else
      if iControl = 2 then ctrl := 'GA_COMMISSIONNABLE'
    else
      if iControl = 3 then ctrl := 'GA_REMISELIGNE'
    else ctrl := 'GA_REMISEPIED';
    CC := TCheckBox(TFMul(F).FindComponent(ctrl));
    if (CC <> nil) and (CC.State = cbChecked) then
    begin
      if xx_where <> '' then xx_where := xx_where + ' AND ';
      xx_where := xx_where + ctrl + '="X"';
    end
    else
      if (CC <> nil) and (CC.State = cbUnChecked) then
    begin
      if xx_where <> '' then xx_where := xx_where + ' AND ';
      xx_where := xx_where + ctrl + '="-"';
    end;
  end;

  // Gestion des checkBox : booléens libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');

  // Gestion des dates libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');

  // Gestion des montants libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');

  SetControlText('XX_WHERE', xx_where);

  CB := TCheckBox(TFMul(F).FindComponent('STATUT_DIM'));
  bAffiche := boolean((CB <> nil) and (CB.State = cbChecked));

  // Annule la sélection de critères dimension si les articles dimensionnés ne sont pas sélectionnés.
  if not bAffiche then
  begin
    reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE')));
    reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_1')));
    reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_2')));
    reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_3')));
    reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_4')));
  end;

  F.Q.Manuel := False;

  if NatureType = 'VTE' then
    TFmul(Ecran).Caption := 'Mise à jour des tarifs de vente'
  else TFmul(Ecran).Caption := 'Mise à jour des tarifs d''achat';
  UpdateCaption(Ecran);
end;

procedure TOF_TarifMAJMode.OnUpdate;
var iCol: integer;
  bAffiche: boolean;
  F: TFMul;
  TLIB: THLabel;
  CB: TCheckBox;
  stIndice, stLabel: string;
begin
  inherited;
  if not (Ecran is TFMul) then exit;
  F := TFMul(Ecran);

  CB := TCheckBox(TFMul(F).FindComponent('STATUT_DIM'));
  bAffiche := boolean((CB <> nil) and (CB.State = cbChecked));

  // Maj de la liste : affichage ou non des colonnes dimensions et modif des titres

{$IFDEF EAGLCLIENT}
  // Mise en place des libellés dans les colonnes
  for iCol := 0 to F.FListe.ColCount - 1 do
  begin
    if copy(F.FListe.Cells[iCol, 0], 1, 7) = 'Famille' then
    begin // Mise en place des libellés des familles
      stIndice := copy(F.FListe.Cells[iCol, 0], length(F.FListe.Cells[iCol, 0]), 1);
      if (stIndice > '3') and (stIndice < '9') then stLabel := 'TGA2_FAMILLENIV' else stLabel := 'TGA_FAMILLENIV';
      TLIB := THLabel(F.FindComponent(stLabel + stIndice));
      if TLIB <> nil then
      begin
        if TLIB.Caption = '.-' then F.Fliste.colwidths[iCol] := 0
        else F.Fliste.cells[iCol, 0] := TLIB.Caption;
      end;
    end
    else
      if (bAffiche) and (copy(F.FListe.Cells[iCol, 0], 1, 3) = 'DIM') then
    begin // Mise en place des libellés des dimensions
      stIndice := copy(F.FListe.Cells[iCol, 0], length(F.FListe.Cells[iCol, 0]), 1);
      TLIB := THLabel(F.FindComponent('DIMENSION' + stIndice));
      if TLIB <> nil then
      begin
        if TLIB.Caption = '.-' then F.Fliste.colwidths[iCol] := 0
        else F.Fliste.cells[iCol, 0] := TLIB.Caption;
      end;
    end
    else
      if (copy(F.FListe.Cells[iCol, 0], 1, 19) = 'Statistique article') then
    begin // Mise en place des libellés des statistiques articles
      stIndice := copy(F.FListe.Cells[iCol, 0], length(F.FListe.Cells[iCol, 0]), 1);
      TLIB := THLabel(F.FindComponent('TGA2_STATART' + stIndice));
      if TLIB <> nil then
      begin
        if TLIB.Caption = '.-' then F.Fliste.colwidths[iCol] := 0
        else F.Fliste.cells[iCol, 0] := TLIB.Caption;
      end;
    end;
  end;
{$ELSE}
  // Mise en place des libellés dans les colonnes
  for iCol := 0 to F.FListe.Columns.Count - 1 do
  begin
    if copy(F.FListe.Columns[iCol].Title.caption, 1, 7) = 'Famille' then
    begin // Mise en place des libellés des familles
      stIndice := copy(F.FListe.Columns[iCol].Title.caption, length(F.FListe.Columns[iCol].Title.caption), 1);
      if (stIndice > '3') and (stIndice < '9') then stLabel := 'TGA2_FAMILLENIV' else stLabel := 'TGA_FAMILLENIV';
      TLIB := THLabel(F.FindComponent(stLabel + stIndice));
      if TLIB <> nil then
      begin
        if TLIB.Caption = '.-' then F.Fliste.columns[iCol].visible := False
        else F.Fliste.columns[iCol].Field.DisplayLabel := TLIB.Caption;
      end;
    end
    else
      if (bAffiche) and (copy(F.FListe.Columns[iCol].Title.caption, 1, 3) = 'DIM') then
    begin // Mise en place des libellés des dimensions
      stIndice := copy(F.FListe.Columns[iCol].Title.caption, length(F.FListe.Columns[iCol].Title.caption), 1);
      TLIB := THLabel(F.FindComponent('DIMENSION' + stIndice));
      if TLIB <> nil then
      begin
        if TLIB.Caption = '.-' then F.Fliste.columns[iCol].visible := False
        else F.Fliste.columns[iCol].Field.DisplayLabel := TLIB.Caption;
      end;
    end
    else
      if (copy(F.FListe.Columns[iCol].Title.caption, 1, 19) = 'Statistique article') then
    begin // Mise en place des libellés des statistiques articles
      stIndice := copy(F.FListe.Columns[iCol].Title.caption, length(F.FListe.Columns[iCol].Title.caption), 1);
      TLIB := THLabel(F.FindComponent('TGA2_STATART' + stIndice));
      if TLIB <> nil then
      begin
        if TLIB.Caption = '.-' then F.Fliste.columns[iCol].visible := False
        else F.Fliste.columns[iCol].Field.DisplayLabel := TLIB.Caption;
      end;
    end;
  end;
{$ENDIF}
end;

/////////////// Mise à jour des tarifs articles //////////////

procedure TOF_TarifMAJMode.MajTarifArt;
var F: TFMul;
  i: integer;
  Statut: string;
begin
  TOBArticle := TOB.Create('_ArticlesSelect', nil, -1);
  F := TFMul(Ecran);
  if (F.FListe.NbSelected = 0) and (not F.FListe.AllSelected) then
  begin
    //MessageAlerte('Aucun élément sélectionné');
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    if VAlerte <> nil then VAlerte.Visible := FALSE;
    {$ENDIF}
    HShowMessage('0;' + F.Caption + ';' + TexteMessage[3] + ';W;O;O;O;', '', '');
    exit;
  end;
  with TFMul(Ecran) do
  begin
    if FListe.AllSelected then
    begin
      Q.First;
      while not Q.EOF do
      begin
        ChargeLesArticles;
        Q.NEXT;
      end;
      FListe.AllSelected := False;
    end else
    begin
      for i := 0 to FListe.NbSelected - 1 do
      begin
        FListe.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row - 1);
        {$ELSE}
        {$ENDIF}
        ChargeLesArticles;
      end;
      FListe.ClearSelected;
    end;
  end;
  // Assistant de mise à jour
  if F.Q.Liste = 'GCMULARTDIM_MODE' then Statut := 'DIM' else Statut := 'GEN';
  MiseAJourTarifMode(TOBArticle, Statut, NatureType);
  if F.FListe.AllSelected then F.FListe.AllSelected := False else F.FListe.ClearSelected;
  F.bSelectAll.Down := False;
  TOBArticle.free;
  TOBArticle := nil;
end;

procedure TOF_TarifMAJMode.ChargeLesArticles;
var TOBArt: TOB;
  Article, CodeArticle: string;
  QQ: Tquery;
begin
  Article := TFmul(Ecran).Q.FindField('GA_ARTICLE').asstring;
  CodeArticle := TFmul(Ecran).Q.FindField('GA_CODEARTICLE').asstring;
  TOBArt := TOB.Create('_ARTICLE', TOBArticle, -1);
  QQ := OpenSql('SELECT GA_ARTICLE,GA_CODEARTICLE,GA_STATUTART,GA_PRIXUNIQUE,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5' +
    ',GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5 From ARTICLE Where GA_ARTICLE="' + Article + '"', TRUE);
  if not QQ.EOF then TOBArt.SelectDB('', QQ);
  ferme(QQ);
end;

/////////////// Procedure appellé par le bouton Validation //////////////

procedure AGLTarfArtMaj_mul_Maj(parms: array of variant; nb: integer);
var F: TForm;
  MaTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then MaTOF := TFMul(F).LaTOF else exit;
  if (MaTOF is TOF_TarifMAJMode) then TOF_TarifMAJMode(MaTOF).MajTarifArt else exit;
end;

procedure ChangeFieldMulTarif(Parms: array of variant; nb: integer);
var F: TForm;
  FieldName, stName: string;
  iName: integer;
  bTest: boolean;
  CodeDim: THValComboBox;
begin
  F := TForm(Longint(Parms[0]));
  FieldName := Parms[1];

  // Modif automatique des critères : statut article et critères dimension
  if (copy(FieldName, 1, 11) = 'GDI_LIBELLE') then
  begin
    CodeDim := THValComboBox(F.FindComponent(FieldName));
    if (CodeDim.Text <> '') and (CodeDim.Text <> TraduireMemoire('<<Tous>>')) then
    begin
      GetSetStatut(F, 'STATUT_GEN', True, bModifCritAuto, bOldStatutGen);
      GetSetStatut(F, 'STATUT_UNI', True, bModifCritAuto, bOldStatutUni);
      GetSetStatut(F, 'STATUT_DIM', True, bModifCritAuto, bOldStatutDim);
      bModifCritAuto := True;
    end
    else
      if (bModifCritAuto) then // Modif critères automatique
    begin
      bTest := True;
      for iName := 1 to 5 do
      begin
        if iName = 1 then stName := 'GDI_LIBELLE'
        else
          if iName = 2 then stName := 'GDI_LIBELLE_1'
        else
          if iName = 3 then stName := 'GDI_LIBELLE_2'
        else
          if iName = 4 then stName := 'GDI_LIBELLE_3'
        else stName := 'GDI_LIBELLE_4';
        CodeDim := THValComboBox(F.FindComponent(stName));
        bTest := bTest and ((CodeDim.Text = '') or (CodeDim.Text = TraduireMemoire('<<Tous>>')));
      end;
      if bTest then // Toutes les combos "vides" -> réinit statut
      begin
        bModifCritAuto := False;
        GetSetStatut(F, 'STATUT_GEN', False, bModifCritAuto, bOldStatutGen);
        GetSetStatut(F, 'STATUT_UNI', False, bModifCritAuto, bOldStatutUni);
        GetSetStatut(F, 'STATUT_DIM', False, bModifCritAuto, bOldStatutDim);
      end;
    end;
  end;

  if (copy(FieldName, 1, 7) = 'STATUT_') then
  begin
    bModifCritAuto := False;
    if ((FieldName = 'STATUT_GEN') or (FieldName = 'STATUT_UNI')) and
      (TCheckBox(F.FindComponent(FieldName)).checked) then
    begin
      // réinit combos dimension
      reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE')));
      reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_1')));
      reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_2')));
      reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_3')));
      reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_4')));
    end;
  end;

end;

initialization
  registerclasses([TOF_TarifMAJMode]);
  RegisterAglProc('TarfArtMaj_mul_Maj', TRUE, 1, AGLTarfArtMaj_mul_Maj);
  RegisterAglProc('ArticleModeChangeDimGen', True, 1, ChangeFieldMulTarif);
end.
