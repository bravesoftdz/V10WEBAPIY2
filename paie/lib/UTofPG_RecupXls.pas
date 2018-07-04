{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 14/02/2002
Modifié le ... : 14/02/2002
Description .. : Unit de récupération de données à partir d'une
Suite ........ : feuille EXCEL
Mots clefs ... : PAIE;
*****************************************************************}
{
PT1   : 02/07/2002 PH V582 prise en compte creation enrg
---- PH 10/08/2005 Suppression directive de compil $IFDEF AGL550B ----
PT2   : 22/05/2007 VG V_72 Intégration du planning unifié
PT3   : 13/10/2008 SJ FQ n°15288 Création automatique du compte tiers lorsque le matricule est alpha-numérique
}
unit UTofPG_RECUPXLS;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, Comctrls, HTB97,
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, HSysMenu, UTOF, UTOB, UTOBXLS, Vierge, HStatus,
{$IFNDEF EAGLCLIENT}
  DBCtrls, HDB, HQry, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HQuickRP, QRGrid, PrintDbg,
{$ELSE}
  HQry, UtileAgl,
{$ENDIF}
  AGLInit, ed_tools, pgoutils, P5Def, P5Util, PgOutils2,
  YRessource;
  
type
  TOF_PG_RECUPXLS = class(TOF)
  private
    LaGrille: THGrid; // Grilles de saisie des cumuls et des bases
    TSal, TPar: TOB; // TOB des salaries et du paramètrage import
//       PT1  : 02/07/2002 V582 PH prise en compte creation enrg
    ParCol: array[1..100] of string; // Colonnes
    ParPre: array[1..100] of string; // Prefixe
    ParSuf: array[1..100] of string; // suffixe
    ParTab: array[1..100] of string; // tablette
    ParTyp: array[1..100] of string; // type
    ParEcr: array[1..100] of string; // Autorisation ecriture dans la table
    ParVal: array[1..100] of string; // Valeurs origines récupérés
    OkEcrit, UneErr: Boolean; // Top OK ecriture, erreur sur la ligne
    NbreCol: Integer;
    HMTrad: THSystemMenu;
    TOB_Val: TOB; // TOB des Valeurs origines
    TOB_Etab: TOB; // TOB des etablissements
    TRib: TOB;
    OkEtab, OkMat: Boolean; // Indicateur si champ Etab, et matricule présents
    OkRib: Boolean; // Indicateur si champ de la table RIB
    IndEtab: Integer; // Indice correspondant à la colonne etablissement
    TypeTrait: string; // Type de traiement I=Import fichier XLS sinon Copier/coller
    function OnSauve: boolean;
    procedure ValiderClick(Sender: TObject);
    procedure FermeClick(Sender: TObject);
    procedure GrilleCopierColler(Sender: TObject);
    procedure ActiveBtn(Sender: TObject);
    function IdentSalarie(var Salarie: string): Integer;
    procedure ImpClik(Sender: TObject);
    procedure GridColorCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function VerifBic(TS: TOB): Boolean;
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;

implementation

uses TiersUtil;

procedure TOF_PG_RECUPXLS.OnArgument(Arguments: string);
var BtnValid, BtnFerme, BtnImp, BtnRecop: TToolbarButton97;
  Q: TQuery;
begin
  inherited;
  TSal := nil;
  TypeTrait := '';
  BtnImp := TToolbarButton97(GetControl('Bimprimer'));
  if BtnImp <> nil then BtnImp.OnClick := ImpClik;
// recuperation de la query du multicritere
  BtnValid := TToolbarButton97(GetControl('BVALIDER'));
  if BtnValid <> nil then BtnValid.OnClick := ValiderClick;
  BtnFerme := TToolbarButton97(GetControl('BFERME'));
  if BtnFerme <> nil then BtnFerme.OnClick := FermeClick;
  BtnRecop := TToolbarButton97(GetControl('RECOPIER'));
  if BtnRecop <> nil then BtnRecop.OnClick := ActiveBtn;
  SetControlEnabled('BVALIDER', FALSE);
  TPar := TOB.Create('Les param', nil, -1);
  TPar.loaddetaildb('PAIEPARIM', '', '', nil, FALSE);
  LaGrille := THGrid(Getcontrol('GRILLE'));
  if LaGrille <> nil then
  begin
    LaGrille.GetCellCanvas := GridColorCell;
    LaGrille.OnKeyDown := KeyDown;
  end;
  Tsal := TOB.Create('Les Salaries', nil, -1);
  Q := OpenSql('SELECT PSA_SALARIE FROM SALARIES', TRUE);
  Tsal.LoadDetailDB('SALARIES', '', '', Q, FALSE, False);
  FERME(Q);
  TOB_Etab := TOB.Create('Les Etab', nil, -1);
  Q := OpenSql('SELECT * FROM ETABCOMPL', TRUE);
  TOB_Etab.LoadDetailDB('ETABCOMPL', '', '', Q, FALSE, False);
  FERME(Q);

end;

// Fonction de validation de la saisie

procedure TOF_PG_RECUPXLS.ValiderClick(Sender: TObject);
var
  rep: Integer;
begin
  inherited;
  Rep := PGIAsk('Voulez vous sauvegarder votre saisie ?', Ecran.Caption);
  if rep = mrNo then exit;
  if rep = mrCancel then exit;
  if rep = mryes then OnSauve;
end;
// fermeture de la forme

procedure TOF_PG_RECUPXLS.FermeClick(Sender: TObject);
begin
  ValiderClick(Sender);
  Close;
end;

{ fonction d'ecriture des elements recuperer de la feuille excel dans la grille de Saisie
Il s'agit en fait de construire la tob qui mettra à jour la table salarié
}

function TOF_PG_RECUPXLS.OnSauve: boolean;
var LeSal, st, LeChamp: string;
  i, j: Integer;
  T_sal, TS, TVal, TR: TOB;
  Q: TQuery;
  InfTiers: Info_Tiers;
  CodeAuxi, LeRapport: string;
  OkCreat: Boolean;
begin
  result := TRUE;
  if not OkEcrit then
  begin
    PGIBox('Vous avez des erreurs, les modifications ne peuvent pas être prises en compte', Ecran.caption);
    exit;
  end;
  try
    st := 'in (';
    for i := 1 to LaGrille.Rowcount - 1 do
    begin
      LeSal := LaGrille.cells[0, i]; // Recup numero salarie
      if LeSal <> '' then
      begin
        if i <> 1 then st := st + ',';
        st := st + '"' + LeSal + '"';
      end;
    end;
    St := St + ') ORDER BY PSA_SALARIE';
    T_sal := TOB.Create('Les Salaries', nil, -1);
    Trib := TOB.Create('Les RIB Salaries', nil, -1);
    Q := OpenSql('SELECT * FROM SALARIES WHERE PSA_SALARIE ' + st, TRUE);
    T_sal.LoadDetailDB('SALARIES', '', '', Q, FALSE, False);
    FERME(Q);
    BeginTrans;
    for i := 1 to LaGrille.Rowcount - 1 do
    begin
      LeSal := LaGrille.cells[0, i]; // Recup numero salarie
      TS := T_Sal.FindFirst(['PSA_SALARIE'], [LeSal], FALSE);
// PT1  : 02/07/2002 V582 PH prise en compte creation enrg
      OkCreat := FALSE;
      if (TS = nil) and (OkEtab) then // (VH_Paie.PgEAI) AND
      begin
        TS := TOB.create('SALARIES', TSal, -1);
        if TS <> nil then InitSalDef(TS, TOB_Etab, ParVal[IndEtab]);
        OkCreat := TRUE;
      end;
      if TS <> nil then
      begin
        for j := 1 to NbreCol do ParVal[j] := '';
        TVal := TOB_Val.FindFirst(['LeSalarie'], [LeSal], FALSE);
        st := '';
        if TVal <> nil then st := TVal.GetValue('Tableau');
        for j := 1 to NbreCol do ParVal[j] := readtokenst(st);
        for j := 2 to LaGrille.colCount do
        begin
          if (ParPre[j] <> 'PSA') then continue;
          if ((ParEcr[j] <> 'X')) then continue; //  AND NOT VH_paie.PgEAI
          LeChamp := ParPre[j] + '_' + ParSuf[j];
          if ParTyp[j] = 'F' then TS.Putvalue(LeChamp, Valeur(ParVal[j]));
          if ParTyp[j] = 'D' then
            TS.Putvalue(LeChamp, StrToDate(LaGrille.cells[j - 1, i]));
          if ParTyp[j] = 'I' then TS.Putvalue(LeChamp, VALEURI(ParVal[j]));
          if (ParTyp[j] = 'T') or (ParTyp[j] = 'B') then
          begin
            TS.Putvalue(LeChamp, ParVal[j]);
          end;
          if (ParTyp[j] = 'S') then TS.Putvalue(LeChamp, ParVal[j]);
        end;
      end
      else ; // Grosse Erreur  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      TR := nil;
      if (TS <> nil) and (OkEtab) and (OkCreat) then // (VH_Paie.PgEAI) AND
      begin
        for j := 2 to LaGrille.colCount do
        begin
          if (ParPre[j] = 'R') then
          begin
            if TR = nil then TR := Tob.Create('RIB', TRib, -1);
            LeChamp := ParPre[j] + '_' + ParSuf[j];
            if (ParTyp[j] = 'S') then TR.Putvalue(LeChamp, ParVal[j]);
          end;
        end;
      if (VH_Paie.PgTiersAuxiAuto = TRUE) then//PT3
        begin
        with InfTiers do
        begin
          Libelle := TS.getValue('PSA_LIBELLE');
          Adresse1 := TS.getValue('PSA_ADRESSE1');
          Adresse2 := TS.getValue('PSA_ADRESSE2');
          Adresse3 := TS.getValue('PSA_ADRESSE3');
          Ville := TS.getValue('PSA_VILLE');
          Telephone := TS.getValue('PSA_TELEPHONE');
          CodePostal := TS.getValue('PSA_CODEPOSTAL');
          Pays := '';
        end;
        CodeAuxi := CreationTiers(InfTiers, LeRapport, 'SAL', TS.getValue('PSA_SALARIE')); // Creation du compte de tiers en automatique et recup du numéro
        if CodeAuxi <> '' then
        begin
          TS.Putvalue('PSA_AUXILIAIRE', CodeAuxi);
          if TR <> nil then
          begin
            TR.PutValue('R_AUXILIAIRE', CodeAuxi);
            TR.PutValue('R_NUMERORIB', 1);
            TR.PutValue('R_PRINCIPAL', 'X');
            TR.PutValue('R_SOCIETE', '000');
            TR.PutValue('R_SALAIRE', 'X');
            TR.PutValue('R_ACOMPTE', 'X');
            TR.PutValue('R_FRAISPROF', 'X');
            TR.PutValue('R_PAYS', 'FRA');
          end
          else TR.Free;
          if (not VerifBic(TR)) then TR.free;
        end;
        CreateYRS (TS.getValue ('PSA_SALARIE'), '', '');       //PT2
      end;//PT3
      end;
    end;
    T_sal.InsertOrUpdateDB(TRUE);
    if TRib.Detail.count > 0 then TRib.InsertOrUpdateDB(TRUE);
    CommitTrans;
  except
    result := FALSE;
    Rollback;
    PGIBox('Une erreur est survenue lors de la validation de la saisie', Ecran.Caption);
  end;
  if T_sal <> nil then T_sal.free;
end;

procedure TOF_PG_RECUPXLS.OnClose;
begin
  if TSal <> nil then begin TSal.Free; TSal := nil; end;
  if TPAr <> nil then begin TPAr.free; TPAr := nil; end;
  if TOB_Val <> nil then begin TOB_Val.free; TOB_Val := nil; end;
  if TOB_Etab <> nil then begin TOB_Etab.free; TOB_Etab := nil; end;
  if TRib <> nil then begin TRib.Free; TRib := nil; end;
end;

{ Cette fonction recupère les données de la feuille EXCEL dans une TOB
Elle alimente la grille pour voir les données qui seront recupérées.
Les salariés inexistants ne sont pas pris et idem pour les colonnes.
Les entetes de colonne sont acceptées en fonction de l'existance
dans la table paramètrage de l'import.
}

procedure TOF_PG_RECUPXLS.GrilleCopierColler(Sender: TObject);
var T, T1, TT, TVal: TOB;
  i, j, k, LstErrorL, LstErrorP: integer;
  St, Salarie, FileN: string;
  LeVar: Variant;
  TraceErr: TListBox;
  zz, lig: Integer;
begin
  if LaGrille = nil then
  begin
    PgiBox('Attention, grille non identifiée', Ecran.Caption);
    exit;
  end;

  LstErrorL := 0;
  LstErrorP := 0;
  if TypeTrait = '' then
  begin
    SourisSablier;
    T := TOBLoadFromClipBoard;
    SourisNormale;
    if T.Detail.Count <= 1 then
    begin
      PgiBox('Attention, vous n''avez rien sélectionné dans votre feuille EXCEL', Ecran.Caption);
      exit;
    end;
  end
  else
  begin
    FileN := GetControlText('FICHIER');
    if FileN = '' then
    begin
      PgiBox('Attention, vous n''avez pas sélectionné de fichier EXCEL', Ecran.Caption);
      exit;
    end;
    if not FileExists(FileN) then
    begin
      PgiBox('Attention, votre fichier EXCEL n''existe pas', Ecran.Caption);
      exit;
    end;
    T := TOB.Create('Ma tob', nil, -1);
    SourisSablier;
    ImportTOBFromXLS(T, FileN);
    SourisNormale;
  end;

  TOB_Val := TOB.Create('Les valeurs origines', nil, -1);
  TraceErr := TListBox(GetControl('LSTBXERROR'));
  if TraceErr = nil then
  begin
    PgiBox('Attention, composant trace non trouvé, abandon du traitement', Ecran.Caption);
    exit;
  end;
  TraceErr.Refresh;
  if TypeTrait = '' then LaGrille.RowCount := T.Detail.Count - 1 // determination du nombre de ligne de la grille
  else LaGrille.RowCount := T.Detail.Count;
// récupération du nombre de champs saisis dans la feuille EXCEL
  if TypeTrait = '' then TT := T.Detail[1]
  else TT := T.Detail[0];
  InitMoveProgressForm(nil, 'Traitement en cours', 'Veuillez patienter SVP ...', T.Detail.Count, FALSE, TRUE);
  NbreCol := T.Detail[1].ChampsSup.Count;
  LaGrille.ColCount := NbreCol;
  i := 1;
  IndEtab := 0;
  for j := 0 to T.Detail[1].ChampsSup.Count - 1 do
  begin
// Identification des champs de la feuille EXCEL par rapport au paramètrage mis en place
// remplissage des tableaux afin d'accèder plus rapidement aux infos sans parcourir
// la TOB autant de fois qu'il y a de ligne (et/ou de colonnes) et de ne reconnaitre que
// les champs présents dans la feuille excel
    st := TCS(T.Detail[1].Champssup[j]).Nom;
    T1 := TPar.FindFirst(['PAI_COLONNE'], [st], FALSE);
    if T1 <> nil then
    begin
      ParCol[i] := T1.GetValue('PAI_COLONNE'); // Nom de la colonne
      ParPre[i] := T1.GetValue('PAI_PREFIX'); // Prefixe
      ParSuf[i] := T1.GetValue('PAI_SUFFIX'); // suffixe
      ParTab[i] := T1.GetValue('PAI_LIENASSOC'); // tablette
      ParTyp[i] := T1.GetValue('PAI_LETYPE'); // type de données
      ParEcr[i] := T1.GetValue('PAI_CHOIX'); // Autorisation ecriture dans la base
      LaGrille.Cells[i - 1, 0] := TT.GetValue(st);
      LaGrille.ColAligns[i - 1] := taLeftJustify;
//PT1  : 02/07/2002 V582 PH prise en compte creation enrg
      if ParCol[i] = 'ETAB' then
      begin
        OkEtab := True;
        IndEtab := i;
      end
      else OkEtab := FALSE;
      if ParCol[i] = 'MATRICULE' then OkMat := True else Okmat := FALSE;
      if ParPre[i] = 'R' then OkRib := TRUE;

      if ParTyp[i] = 'B' then // Boolean
      begin
        LaGrille.ColTypes[i] := 'B';
        LaGrille.ColAligns[i] := taCenter;
// IntTostr (Integer(CsCoche));
        LaGrille.ColFormats[i] := IntToStr(Ord(csCheckBox));
      end;
      if ParTyp[i] = 'D' then //  Date
      begin
        LaGrille.ColTypes[i] := 'D';
        LaGrille.ColAligns[i] := taCenter;
        LaGrille.ColFormats[i] := ShortDateFormat;
      end;
      if ParTyp[i] = 'I' then //  Integer
      begin
        LaGrille.ColTypes[i] := 'I';
        LaGrille.ColAligns[i] := taCenter;
        LaGrille.ColFormats[i] := '##0';
      end;
      if ParTyp[i] = 'F' then //  Float ou double
      begin
        LaGrille.ColAligns[i] := taRightJustify;
        LaGrille.ColFormats[i] := '# ##0.00';
      end;
      i := i + 1;
    end;
  end;
  LaGrille.FixedCols := 2;
  if i <= 2 then exit;
  if TypeTrait = '' then zz := 2
  else zz := 1;

  for i := zz to T.Detail.Count - 1 do
  begin
    TT := T.Detail[i];
    LeVar := TT.GetValue('Matricule');
    Salarie := VarToStr(LeVar);
    k := IdentSalarie(Salarie); // recherche si le salarie existe
    MoveCurProgressForm('Salarié : ' + Salarie + ' en cours de traitement'); // PORTAGECWAS
    if k = -1 then
    begin
      LstErrorL := LstErrorL + 1;
      TraceErr.Items.Add('Le Salarié ' + Salarie + ' est inconnu');
      LaGrille.Cells[0, i - 1] := Salarie;
      LaGrille.Cells[1, i - 1] := 'Le salarié est inconnu';
    end;
    if (k > 0) then
    begin // Boucle sur les lignes de la TOB
      if TypeTrait = '' then lig := i - 1
      else lig := i;
      for j := 1 to NbreCol do ParVal[j] := '';
      for j := 1 to NbreCol do
      begin // Boucle sur les colonnes de la grille qui correspondent aux colonnes de la feuilles EXCEL
        LeVar := TT.GetValue(ParCol[j]);
        if (ParCol[j] = 'MATRICULE') and (VH_PAie.PgTypeNumSal = 'NUM') then LeVar := ColleZerodevant(LeVar, 10);
        UneErr := FALSE;
// type de champ traité
        if ParTyp[j] = 'D' then
        begin
          if IsNumeric(VarToStr(LeVar)) then LaGrille.Cells[j - 1, lig] := DateToStr(LeVar)
          else
          begin
            TraceErr.Items.Add('Pour le Salarié ' + Salarie + ' le champ ' + LaGrille.Cells[j - 1, 0] + ' n''est pas numérique');
            LstErrorP := LstErrorP + 1;
            UneErr := TRUE;
          end;
        end;
        if (ParTyp[j] = 'I') or (ParTyp[j] = 'F') then
        begin
          if IsNumeric(VarToStr(LeVar)) then LaGrille.Cells[j - 1, lig] := DoubleToCell(Valeur(VarToStr(LeVar)), 0)
          else
          begin
            TraceErr.Items.Add('Pour le Salarié ' + Salarie + ' le champ ' + LaGrille.Cells[j - 1, 0] + ' n''est pas numérique');
            LstErrorP := LstErrorP + 1;
            UneErr := TRUE;
          end;
        end;
        if ParTyp[j] = 'T' then
        begin
          st := RechDom(ParTab[j], LeVar, FALSE);
          if (st = '') or (St = 'Error') then // 2 eme tentative zone numerique alors on complète avec 0
          begin
            if IsNumeric(VarToStr(LeVar)) then
            begin
              LeVar := ColleZerodevant(LeVar, 3);
              st := RechDom(ParTab[j], LeVar, FALSE);
            end;
          end;
          if (st = '') or (St = 'Error') then
          begin
            TraceErr.Items.Add('Pour le Salarié ' + Salarie + ' le champ ' + LaGrille.Cells[j - 1, 0] + ' n''a pas de valeur connue');
            if (st = '') or (St = 'Error') then st := 'Erreur de codification';
            LstErrorP := LstErrorP + 1;
            UneErr := TRUE;
          end;
          LaGrille.Cells[j - 1, lig] := st;
        end;
        if (ParTyp[j] = 'B') or (ParTyp[j] = 'S') then
        begin
          LaGrille.Cells[j - 1, lig] := LeVar;
        end;
        ParVal[j] := LeVar; // pour stocker les valeurs initiales de la feuille EXCEL
      end; // fin de la boucle sur les colonnes
      TVal := TOB.Create('Mes Valeurs', TOB_Val, -1);
      TVal.AddChampSup('LeSalarie', FALSE);
      TVal.AddChampSup('Tableau', FALSE);
      TVal.PutValue('LeSalarie', Salarie);
      st := '';
      for j := 1 to NbreCol do
      begin
        St := st + ParVal[j] + ';'; // Confection d'une chaine
      end;
      TVal.PutValue('Tableau', st);
    end; // si salarie identifie
  end; // Fin de la boucle sur la TOB
  T.Free;
  FiniMoveProgressForm(); // PORTAGECWAS
  HMTrad.ResizeGridColumns(LaGrille);
  St := 'Attention, ';
  if LstErrorL > 0 then St := St + IntTOstr(LstErrorL) + ' ligne(s) n''ont pas été intégrées';
  if LstErrorP > 0 then St := St + IntTOstr(LstErrorP) + ' cellule(s) colonne n''ont pas été intégrées';
  if (LstErrorL > 0) or (LstErrorP > 0) then PGIBox(St, Ecran.Caption);
  if (LstErrorP > 0) or (LstErrorL > 0) then OkEcrit := FALSE else
  begin
    okEcrit := TRUE;
    SetControlEnabled('BVALIDER', true);
  end;
end;

function TOF_PG_RECUPXLS.IdentSalarie(var Salarie: string): Integer;
var TS: TOB;
  LeSal: string;
begin
  result := -1;
  if VH_PAie.PgTypeNumSal = 'NUM' then LeSal := ColleZerodevant(StrtoInt(Salarie), 10)
  else LeSAl := Salarie;

  TS := Tsal.FindFirst(['PSA_SALARIE'], [LeSal], FALSE);
  if TS <> nil then
  begin
    result := 1;
    Salarie := LeSal;
  end
  else
// PT1  : 02/07/2002 V582 PH prise en compte creation enrg
  begin
    if (not OkMat) or (not OkEtab) then exit; // (NOT VH_Paie.PgEAI) OR
    result := 9;
    Salarie := LeSal;
  end;
end;

procedure TOF_PG_RECUPXLS.ImpClik(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  PrintDBGrid(TCustomGrid(LaGrille), nil, Ecran.Caption, '');
{$ELSE}
  PrintDBGrid('GRILLE', '', Ecran.Caption, '');
{$ENDIF}
end;

procedure TOF_PG_RECUPXLS.GridColorCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if Copy(LaGrille.Cells[ACol, ARow], 1, 6) = 'Erreur' then Canvas.Font.Color := clRed;
  if LaGrille.Cells[ACol, ARow] = 'Le salarié est inconnu' then Canvas.Font.Color := clBlue;
end;

procedure TOF_PG_RECUPXLS.KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 86) and (ssCtrl in Shift) then
    GrilleCopierColler(Sender);
end;

procedure TOF_PG_RECUPXLS.ActiveBtn(Sender: TObject);
begin
  TypeTrait := 'I';
  GrilleCopierColler(Sender);
end;

function TOF_PG_RECUPXLS.VerifBic(TS: TOB): Boolean;
begin
  result := FALSE;
  if TS.getValue('R_ETABBQ') = '' then exit;
  if TS.getValue('R_GUICHET') = '' then exit;
  if TS.getValue('R_NUMEROCOMPTE') = '' then exit;
  if TS.getValue('R_CLERIB') = '' then exit;
  if TS.getValue('R_CODEBIC') <> '' then begin Result := True; Exit; end else
    if ((TS.getValue('R_ETABBQ') = '') or (TS.getValue('R_GUICHET') = '') or
      (TS.getValue('R_NUMEROCOMPTE') = '') or (TS.getValue('R_CLERIB') = ''))
      then exit;
  result := TRUE;
end;

initialization
  registerclasses([TOF_PG_RECUPXLS]);
end.

