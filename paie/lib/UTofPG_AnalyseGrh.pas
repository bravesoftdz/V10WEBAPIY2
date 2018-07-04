{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Unit permettant de changer les listes sur les THEDIT des
Suite ........ : fiches d'analyses (TOBVIEWER et CUBES) de la GRH.
Suite ........ : Permet aussi de n'afficher que les zones libres gérées
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 12/02/2002 PH V571 Traitement des cubes avec comparatifs
PT2   : 19/03/2002 PH V571 Prise en compte des zones combos libres et du
                           ThMulValComboBox
PT3   : 31/05/2002 VG V582 Version S3
PT4   : 27/06/2002 PH V582 Prise en compte table report
PT5   : 13/08/2002 PH V585 Différenciation des rubiques de bases et de cotisations
PT6   : 31/01/2003 PH V591 Affectation de bonnes tablettes sections en fonction de l'axe
PT7   : 17/02/2003 PH V_42 Tobviewer sur historique salariés
PT8   : 19/08/2003 PH V_421 FQ 10328 Affection des bonnes tablettes sur les rubriques
PT9   : 10/02/2004 VG V_50 Annulation des spécificités S3 - FQ N°10911
PT10  : 11/05/2004 PH V_50 FQ 10659 Edition pour les cubes au mode paysage
PT11  : 09/06/2004 PH V_50 Prise en compte des champs de la table PAIEENCOURS
PT12  : 10/03/2005 PH V_602 FQ 11924 Prise en compte des champs à NULL sous ORACLE + FQ12029
                            Sélection d'un paramétrage obligatoire
PT13  : 22/06/2005 PH V_60 FQ 12049,12211 Extraction RIB et Ville
PT14  : 08/07/2005 PH V_60 FQ 12432 padding à gauche des matricules salariés
PT15  : 29/08/2005 PH V_60 Tablette PGREMUN pour avoir le code et le libellé des rémunérations
PT16  : 24/04/2006 GGR V_7 FQ 12899 Suspension paie
PT17  : 29/12/2006 FCO V_8 FQ 13525 Supprimer la case à cocher liste d'exportation
PT18  : 14/12/2007 GGU V_8 FQ 14065 paramétrage des extractions : pouvoir choisir le TAUX d'une rémunération
PT19  : 04/03/2008 FC V_80 FQ 15206 Gestion des confidentiels
PT20  : 24/04/2008 VG V_80 Anomalies sur le comparatif des rubriques - Le nombre
                           de mois n'était plus pris en compte
}
unit UTofPG_ANALYSEGRH;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  Spin,
  Hqry,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HDB,
  mul,
  FE_Main,
  EdtREtat,
  {$ELSE}
  emul,
  MaineAgl,
  UtilEAGL,
  {$ENDIF}
  HTB97,
  Grids,
  HCtrls,
  HEnt1,
  vierge,
  EntPaie,
  HMsgBox,
  UTOF,
  UTOB,
  UTOM,
  stat,
  Cube,
  AGLInit,
  ParamDat;

type
  TOF_PGANALYSEGRH = class(TOF)
  private
    BtnNature: THValComboBox;
    LeType: string;
    LaCheck: TCheckBox;      //Pt16
    procedure NatureOnExit(Sender: TObject);
    procedure RendLaDate(D1: TDateTime; var XX_WHERE: string; MoisComplet: TCheckBox);
    // PT6   : 31/01/2003 PH V591 Affectation de bonnes tablettes sections en fonction de l'axe
    procedure ChangeSection(Sender: TObject);
    procedure ChangeLeMode(Sender: TObject);
    procedure LanceExtract(Sender: TObject);
    function  RendSQLAgregat(ETALIM, ETCOL, ETLIB, Presentation: string; DateDebut, DateFin: TDateTime; Numero: string; NumCh: Boolean): string;
    function  RendConditionOR(ETCOL, LeChamp: string): string;
    procedure LancerEtatAnalyse(Requete: string);
    Procedure PGRechSal (FF : TForm) ; // PT14
    procedure ExitEdit(Sender: TObject); // PT14
    procedure ExitCheck(Sender: TObject); //PT16
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

uses  PgOutils,
      PgOutils2,
      P5Def;

// DEB PT14
Procedure TOF_PGANALYSEGRH.PGRechSal (FF : TForm) ;
var i : Integer;
LeControl : TControl;
begin
  for i := 0 to FF.ComponentCount - 1 do
  begin
    LeControl := TControl(FF.Components[i]);
    if LeControl is THLabel then continue;
    if (LeControl is THEdit) AND ((pos('SALARIE', LeControl.Name) > 0) OR (pos('SALARIE_', LeControl.Name) > 0)) then
      begin
      THEdit(LeControl).OnExit := ExitEdit;
      end
      else Continue;
  end;
end;

procedure TOF_PGANALYSEGRH.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

// FIN PT14
// PT6   : 31/01/2003 PH V591 Affectation de bonnes tablettes sections en fonction de l'axe
procedure TOF_PGANALYSEGRH.ChangeSection(Sender: TObject);
var
  Axe, Section: THValComboBox;
  St: string;
begin
  Axe := THValComboBox(GetControl('PHA_AXE'));
  Section := THValComboBox(GetControl('PHA_SECTION'));
  if (AXE <> nil) and (SECTION <> nil) then
  begin
    St := Copy(Axe.value, 2, 1);
    if St <> '1' then
      SetControlProperty('PHA_SECTION', 'DataType', 'TZSECTION' + St)
    else
      SetControlProperty('PHA_SECTION', 'DataType', 'TZSECTION');
  end;
end;

procedure TOF_PGANALYSEGRH.LanceExtract(Sender: TObject);
var
  LeWhere, St, St1, St2, TitreParamEx, StRupt: string;
  ParamEx, nbchp, nbcol: Integer;
  Q: TQuery;
  T_PARIM, T1: TOB;
  NumChamp, i: Integer;
  Presentation: string;
  DD, DateD, DateF: TDateTime;
  Abandon, NumCh, OkRIB : boolean; // PT13
begin
  if LeType <> 'EX' then exit;
// DEB PT12
  //Début PT17
{  if (LeType = 'EX') AND (GetControlText('LISTEEXPORT') = 'X') then  SetControlText ('TRAIT', '1')
  else SetControlText ('TRAIT', '2');}
  if (GetControlText('TRAIT') = '') then SetControlText ('TRAIT', '2');
  //Fin PT17

  LeWhere := RecupWhereCritere(TFMul(Ecran).Pages);
  if (GetControlText('PARAMEXTPAIE') = '') then
  begin
    PgiBox ('Sélectionnez un paramétrage', Ecran.caption);
    exit;
  end;
// FIN PT12
  ParamEx := StrToInt(GetControlText('PARAMEXTPAIE'));
  if GetControlText('TRAIT') = '1' then NumCh := TRUE
  else NumCh := FALSE;
  if ParamEx >= 0 then // PT12 pour traiter le cas du numéro de paramétrage = 0
  begin
    T_PARIM := TOB.Create('MaTob Param', nil, -1);
    Q := OpenSql('SELECT * FROM PAIEPARIM', TRUE);
    T_PARIM.LoadDetailDB('PAIEPARIM', '', '', Q, FALSE, FALSE);
    Ferme(Q);
    OkRIB := FALSE; // PT13
    Q := OpenSql('SELECT * FROM PARAMEXTPAIE WHERE ##PPE_PREDEFINI## PPE_NUMORDRE =' + IntToStr(ParamEx), TRUE);
    if not Q.EOF then
    begin
      TitreParamEx := Q.FindField('PPE_LIBELLE').AsString;
      for nbchp := 1 to 10 do
      begin
        NumChamp := Q.FindField('PPE_ZON' + IntToStr(nbchp)).AsInteger;
        if NumChamp > 0 then
        begin
          T1 := T_PARIM.FindFirst(['PAI_IDENT'], [NumChamp], TRUE);
          if T1 <> nil then
          begin
          // DEB PT13
            if (T1.GetValue('PAI_PREFIX') = 'R') AND (T1.GetValue('PAI_SUFFIX')= 'RIB') then
            begin
              st1 := St1 + ',R_ETABBQ||R_GUICHET||R_NUMEROCOMPTE||R_CLERIB RIBSALARIE';
              OkRIB := TRUE;
            end
            else st1 := St1 + ',' + T1.GetValue('PAI_PREFIX') + '_' + T1.GetValue('PAI_SUFFIX');
          // FIN PT13
          end;
        end;
      end;
      Presentation := Q.FindField('PPE_PGPRESENTATION').AsString;
      if Presentation = 'PER' then
      begin // Chaque colonne représente un agrégat
        for nbcol := 1 to 12 do
        begin
          if Q.FindField('PPE_ETALIM' + IntToStr(nbcol)).AsString <> '' then
            St2 := St2 +
              RendSQLAgregat(Q.FindField('PPE_ETALIM' + IntToStr(nbcol)).AsString, Q.FindField('PPE_ETCOL' + IntToStr(nbcol)).AsString,
              Q.FindField('PPE_ETLIB' + IntToStr(nbcol)).AsString, Presentation, StrToDate(GetControlText('DATEDEBUT')), StrToDate(GetControlText('DATEFIN')), IntToStr(nbcol),
              Numch);
        end;
      end
      else
      begin
        DD := StrToDate(GetControlText('DATEDEBUT'));
        DateD := DD;
 //       DF := StrToDate(GetControlText('DATEFIN'));
        for nbCol := 1 to 12 do
        begin
          if nbCol > 1 then DateD := PLUSMOIS(DateD, 1);
          DateF := FINDEMOIS(DateD);
          if Q.FindField('PPE_ETALIM1').AsString <> '' then
            St2 := St2 +
              RendSQLAgregat(Q.FindField('PPE_ETALIM1').AsString, Q.FindField('PPE_ETCOL1').AsString,
              Q.FindField('PPE_ETLIB1').AsString, Presentation, DATED, DATEF, IntToStr(nbcol), Numch);
        end;
      end;
      //Gestion Rupture pour édition
      StRupt := '';
      if (GetControltext('TRAIT') = '1') then
      begin
        for i := 1 to 4 do
        begin
          if TRadioButton(GetControl('R_TRAVAILN' + IntToStr(i))).Checked = True then StRupt := 'PSA_TRAVAILN' + IntToStr(i);
          if TRadioButton(GetControl('R_LIBREPCMB' + IntToStr(i))).Checked = True then StRupt := 'PSA_LIBREPCMB' + IntToStr(i);
        end;
        if TRadioButton(GetControl('R_CODESTAT')).Checked = True then StRupt := 'PSA_CODESTAT';
        if StRupt <> '' then
        begin
          for nbchp := 1 to 10 do
          begin
            NumChamp := Q.FindField('PPE_ZON' + IntToStr(nbchp)).AsInteger;
            if NumChamp > 0 then
            begin
              T1 := T_PARIM.FindFirst(['PAI_IDENT'], [NumChamp], TRUE);
              if T1 <> NIL then
              begin
                if (T1.GetValue('PAI_PREFIX') + '_' + T1.GetValue('PAI_SUFFIX')) = StRupt then StRupt := '';
              end;
            end;
          end;
        end;
        if StRupt <> '' then StRupt := ',' + StRupt;
      end;
      // DEB PT13
      if NOT OkRIB then  St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM' + StRupt + st1 + st2 + ' FROM SALARIES '
      else St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM ' + StRupt + st1 + st2 + ' FROM SALARIES LEFT JOIN RIB ON R_AUXILIAIRE=PSA_AUXILIAIRE AND R_PRINCIPAL="X" ';
      // FIN PT13
      if LeWhere <> '' then St := st + LeWHERE;
      Abandon := FALSE;
    end
    else Abandon := TRUE;
    if not Abandon then
    begin
      if (GetControltext('TRAIT') = '1') then LancerEtatAnalyse(st);
      if (GetControltext('TRAIT') = '2') then
        AglLanceFiche('PAY', 'EXT_STA', '', '', 'STA-EX;' + st + ';' + TitreParamEx)
      else
        if (GetControltext('TRAIT') = '3') then
        AglLanceFiche('PAY', 'EXT_CUB', '', '', 'CUB-EX;' + st + ';' + TitreParamEx);

    end;
    FreeAndNil (T_PARIM);
    Ferme(Q);
  end;
end;

function TOF_PGANALYSEGRH.RendSQLAgregat(ETALIM, ETCOL, ETLIB, Presentation: string; DateDebut, DateFin: TDateTime; Numero: string; NumCh: Boolean): string;
var
  st, st1: string;
  i: Integer;
begin
  i := StrToInt(ETALIM);
  st := ',(SELECT SUM(';
  case i of
    1:
      begin
        St := St + 'PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_SALARIE=PSA_SALARIE AND ';
        st := St + 'PHC_DATEFIN >="' + UsDateTime(DateDebut) + '" AND PHC_DATEFIN <="' + UsDateTime(DateFin) + '"';
        St := St + RendConditionOR(ETCOL, 'PHC_CUMULPAIE');
      end;
    5..6:
      begin
        if i = 5 then st1 := 'PHB_BASEREM'
        else st1 := 'PHB_MTREM';
        St := St + St1 + ') FROM HISTOBULLETIN WHERE PHB_SALARIE=PSA_SALARIE AND PHB_NATURERUB="AAA" AND ';
        st := St + 'PHB_DATEFIN >="' + UsDateTime(DateDebut) + '" AND PHB_DATEFIN <="' + UsDateTime(DateFin) + '"';
        St := St + RendConditionOR(ETCOL, 'PHB_RUBRIQUE');
      end;
    310:  //PT18
      begin
        st1 := 'PHB_TAUXREM';
        St := St + St1 + ') FROM HISTOBULLETIN WHERE PHB_SALARIE=PSA_SALARIE AND PHB_NATURERUB="AAA" AND ';
        st := St + 'PHB_DATEFIN >="' + UsDateTime(DateDebut) + '" AND PHB_DATEFIN <="' + UsDateTime(DateFin) + '"';
        St := St + RendConditionOR(ETCOL, 'PHB_RUBRIQUE');
      end;
    10..12:
      begin
        if i = 10 then st1 := 'PHB_BASECOT'
        else if i = 11 then st1 := 'PHB_MTSALARIAL'
        else St1 := 'PHB_MTPATRONAL';
        St := St + St1 + ') FROM HISTOBULLETIN WHERE PHB_SALARIE=PSA_SALARIE AND PHB_NATURERUB="COT" AND ';
        st := St + 'PHB_DATEFIN >="' + UsDateTime(DateDebut) + '" AND PHB_DATEFIN <="' + UsDateTime(DateFin) + '"';
        St := St + RendConditionOR(ETCOL, 'PHB_RUBRIQUE');
      end;
    20..23:
      begin
        if i = 20 then st1 := 'PHB_BASECOT'
        else if i = 21 then st1 := 'PHB_TRANCHE1'
        else if i = 22 then st1 := 'PHB_TRANCHE2'
        else St1 := 'PHB_TRANCHE3';
        St := St + St1 + ') FROM HISTOBULLETIN WHERE PHB_SALARIE=PSA_SALARIE AND PHB_NATURERUB="BAS" AND ';
        st := St + 'PHB_DATEFIN >="' + UsDateTime(DateDebut) + '" AND PHB_DATEFIN <="' + UsDateTime(DateFin) + '"';
        St := St + RendConditionOR(ETCOL, 'PHB_RUBRIQUE');
      end;
  end;
  st := st + ') ';

  if NumCh then st := st + 'CH' + Numero
  else
  begin
    if Presentation = 'MEN' then
      st := st + FormatDateTime('mmmmyyyy', DateFin)
    else st := st + '"' +FindEtReplace(ETLIB, ' ', '', TRUE)+ '"';
  end;
  result := st;
end;

function TOF_PGANALYSEGRH.RendConditionOR(ETCOL, LeChamp: string): string;
var
  st, St1: string;
  nb: Integer;
begin
  st := ' AND (';
  Nb := 0;
  St1 := readtokenst(ETCOL);
  while st1 <> '' do
  begin
    if nb > 0 then st := st + ' OR ';
    st := st + LeChamp + '="' + St1 + '"';
    St1 := readtokenst(ETCOL);
    Nb := Nb + 1;
  end;
  result := st + ') ';
end;

procedure TOF_PGANALYSEGRH.NatureOnExit(Sender: TObject);
var
  st: string;
  R1, R2: THEdit;
begin
  if BtnNature <> nil then
  begin
    st := BtnNature.Value;
    R1 := THEdit(GetControl('PHB_RUBRIQUE'));
    R2 := THEdit(GetControl('PHB_RUBRIQUE_'));
    if (R1 = nil) then // OR (R2 = NIL)
    begin
      if (R1 = nil) and (R2 = nil) then
      begin
        R1 := THEdit(GetControl('PHA_RUBRIQUE'));
        R2 := THEdit(GetControl('PHA_RUBRIQUE_'));
      end;
      if (R1 = nil) then exit; // OR (R2 = NIL)
    end;
    if St = 'AAA' then
    begin
      // PT8   : 19/08/2003 PH V_421 FQ 10328 Affection des bonnes tablettes sur les rubriques
      SetControlProperty(R1.Name, 'DataType', 'PGREMUN'); // PT15
      if R2 <> nil then SetControlProperty(R2.Name, 'DataType', 'PGREMUN'); // PT15
    end
    else
    begin
      // PT5   : 13/08/2002 PH V585 Différenciation des rubiques de bases et de cotisations
      if St = 'BAS' then
      begin
        SetControlProperty(R1.Name, 'DataType', 'PGBASECOTISATION');
        if R2 <> nil then SetControlProperty(R2.Name, 'DataType', 'PGBASECOTISATION');
      end
      else
      begin
        SetControlProperty(R1.Name, 'DataType', 'PGCOTIS');
        if R2 <> nil then SetControlProperty(R2.Name, 'DataType', 'PGCOTIS');
        // FIN PT8
      end;
    end;
  end;
end;

procedure TOF_PGANALYSEGRH.OnArgument(Arguments: string);
var
  Arg, Ch1, Ch2, Ch3, Cs1, Cs2, Cs3: string;
  Num: Integer;
  Pref: string;
  Axe: THValComboBox;
  BtnLance: TToolbarButton97;
  StSQl, TitreParamExt: string;
  MPaysage: TCheckBox;
begin
  inherited;
  Arg := Arguments;
  Arg := Trim(Arg);
  LeType := Arg;

//PT16
  LaCheck := TCheckBox(GetControl('PSA_SUSPENSIONPAIE'));
  if LaCheck <> nil then
  begin
    LaCheck.OnClick := ExitCheck;
    if LaCheck.Checked then SetControlEnabled('PSA_MOTIFSUSPPAIE', TRUE)
    else SetControlEnabled('PSA_MOTIFSUSPPAIE', FALSE);
  end;
// FIN PT16
  if Arg = '' then exit;
  if Copy(Arg, 1, 6) = 'STA-EX' then
  begin
    Arg := readtokenst(LeType);
    StSql := readtokenst(LeType);
    TitreParamExt := readtokenst(LeType);
    TFStat(Ecran).Caption := TFStat(Ecran).Caption + TitreParamExt;
    UpdateCaption(TFStat(Ecran));
    Num := 0;
    Ch1 := readtokenpipe(StSQl, ',');
    while ch1 <> '' do
    begin
      if Num > 0 then ch1 := ',' + ch1;
      TFStat(Ecran).SQL.Lines.add(ch1);
      Num := Num + 1;
      Ch1 := readtokenpipe(StSQl, ',');
    end;
    exit;
  end;
  if Copy(Arg, 1, 6) = 'CUB-EX' then
  begin
    Arg := readtokenst(LeType);
    StSql := readtokenst(LeType);
    TitreParamExt := readtokenst(LeType);
    TFCube(Ecran).Caption := TFCube(Ecran).Caption + TitreParamExt;
    UpdateCaption(TFCube(Ecran));
    TFCube(Ecran).StSQL := StSql;
    exit;
  end;
  BtnLance := TToolbarButton97(GetControl('BTNLANCE'));
  if BtnLance <> nil then BtnLance.OnClick := LanceExtract;
  Ch1 := '';
  Ch2 := '';
  // PT4   : 27/06/2002 PH V582 Prise en compte table report
  if Arg = 'D' then Pref := 'PHB' else
    if Arg = 'A' then Pref := 'PHA' else
    if Arg = 'C' then Pref := 'PHC' else
    if Arg = 'M' then Pref := 'PTM' else
    // PT7   : 17/02/2003 PH V_42 Tobviewer sur historique salariés
    if Arg = 'HS' then Pref := 'PSA' else
    if Arg = 'SA' then Pref := 'PSA' else
    if Arg = 'SI' then Pref := 'PSI' else
    if Arg = 'BG' then Pref := 'PBG' else
    if Arg = 'PU' then Pref := 'PPU' else  // PT11
    if Arg = 'EX' then Pref := 'PSA'
  else Pref := 'PSA';
  if Pref <> '' then
  begin
    Cs1 := Pref + '_CODESTAT';
    Cs2 := 'T' + Pref + '_CODESTAT';
    Cs3 := 'R_CODESTAT';
    VisibiliteStat(GetControl(Cs1), GetControl(Cs2), GetControl(Cs3));
    VisibiliteStat(GetControl(Cs1 + '_'), GetControl(Cs2 + '_'));
    VisibiliteStat(GetControl(Cs1 + '__'), GetControl(Cs2 + '__'));
  end;
  for Num := 1 to 4 do
  begin
    if Num > 4 then Break;
    Ch1 := Pref + '_TRAVAILN' + IntToStr(Num);
    Ch2 := 'T' + Pref + '_TRAVAILN' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_TRAVAILN' + IntToStr(Num);
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampSalarie(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;

  for Num := 1 to 4 do
  begin
    if Num > 4 then Break;
    Ch1 := Pref + '_LIBREPCMB' + IntToStr(Num);
    Ch2 := 'T' + Pref + '_LIBREPCMB' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      Ch3 := 'R_LIBREPCMB' + IntToStr(Num);
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;
  for Num := 1 to 4 do
  begin
    if Num > 4 then Break;
    Ch1 := Pref + '_BOOLLIBRE' + IntToStr(Num);
    Ch2 := 'T' + Pref + '_BOOLLIBRE' + IntToStr(Num);
    if (Ch1 <> '') and (Ch2 <> '') then
    begin
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1), GetControl(Ch2));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
      VisibiliteChampLibreSal(IntToStr(Num), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
    end;
  end;
  // FIN PT4
  if Arg = 'A' then BtnNature := THValComboBox(GetControl('PHA_NATURERUB'))
  else BtnNature := THValComboBox(GetControl('PHB_NATURERUB'));
  if BtnNature <> nil then BtnNature.OnExit := NatureOnExit;
  // PT6   : 31/01/2003 PH V591 Affectation de bonnes tablettes sections en fonction de l'axe
  if Arg = 'A' then
  begin
    Axe := THValComboBox(GetControl('PHA_AXE'));
    if Axe <> nil then
      Axe.OnChange := ChangeSection;
  end;
  //PT9
//PT10  : 11/05/2004 PH V_50 FQ 10659 Edition pour les cubes au mode paysage
  MPaysage := TCheckBox(GetControl('MODEPAYSAGE'));
  if MPaysage <> NIL then MPaysage.OnClick := ChangeLeMode;

  PGRechSal (Ecran); // PT14
end;

// PT1 12/02/2002 V571 PH Traitement des cubes avec comparatifs
procedure TOF_PGANALYSEGRH.OnLoad;
var
LeWhere, LaDat: THEdit;
LaDate, XX_WHERE, st: string;
D1, D2: TDateTime;
Nbre, i: Integer;
NbreMois: TSpinEdit;
Glissant, MoisComplet: TCheckBox;
begin
inherited;
D2:= IDate1900;
LeWhere:= THEdit (Getcontrol ('XX_WHERE'));
LaDat:= THEdit (Getcontrol ('DATEFIN'));
NbreMois:= TSpinEdit (Getcontrol ('NBREMOIS'));
Nbre:= 0;
MoisComplet:= TCheckBox (GetControl ('MOISCOMPLET'));
if (NbreMois<>nil) then
   Nbre:= NbreMois.value;
if ((LeWhere<>nil) and (LaDat<>nil)) then
   begin
   LaDate:= GetcontrolText ('DATEFIN');
   D1:= StrToDate (LaDate);
   if (Nbre=0) then
      begin // Comparatif annuel
//      RendLaDate (D1, XX_WHERE,MoisComplet);
      st:= 'PHC_DATEFIN="'+UsDateTime (D1)+'" OR'+
           ' PHC_DATEFIN="'+UsDateTime (D2)+'"';
      if (MoisComplet<>nil) then
         begin
         if (MoisComplet.Checked) then
            st:= '(PHC_DATEFIN>="'+UsDateTime (DEBUTDEMOIS (D1))+'" AND'+
                 ' PHC_DATEFIN<="'+UsDateTime (FINDEMOIS (D1))+'") OR'+
                 ' (PHC_DATEFIN>="'+UsDateTime (DEBUTDEMOIS (D2))+'" AND'+
                 ' PHC_DATEFIN<="'+UsDateTime (FINDEMOIS (D2))+'"';
         end;
      if (LeType='D') then
         st:= StringReplace (st, 'PHC', 'PHB', [rfReplaceAll]);
      SetControlText ('XX_WHERE', st);
      end
   else
      begin // comparatif mensuel sur x mois
      Glissant:= TCheckBox (GetControl ('MOISGLISSANT'));
      XX_WHERE:= 'PHC_DATEFIN="'+UsDateTime (D1)+'"';
      if (MoisComplet.Checked) then
         XX_WHERE:= '(PHC_DATEFIN>="'+UsDateTime (DEBUTDEMOIS (D1))+'" AND'+
                    ' PHC_DATEFIN<="'+UsDateTime (FINDEMOIS (D1))+'")';
      if ((Glissant.Checked) and (Nbre<=12)) then
         RendLaDate (D1, XX_WHERE, MoisComplet);
      for i:= 1 to Nbre-1 do
          begin
          D1:= PLUSMOIS (D1, -1);
          D1:= FINDEMOIS (D1);
          XX_WHERE:= XX_WHERE+' OR PHC_DATEFIN="'+UsDateTime (D1)+'"';
          if (MoisComplet.Checked) then
             XX_WHERE:= XX_WHERE+' OR'+
                        ' (PHC_DATEFIN>="'+UsDateTime (DEBUTDEMOIS (D1))+'" AND'+
                        ' PHC_DATEFIN<="'+UsDateTime (FINDEMOIS (D1))+'")';
          if ((Glissant.Checked) and (Nbre<=12)) then
// Gestion glissant si nbre de mois <= 12 pour comparer si 6 mois
             RendLaDate (D1, XX_WHERE, MoisComplet);
          end;
      if (LeType='D') then
         XX_WHERE:= StringReplace (XX_WHERE, 'PHC', 'PHB', [rfReplaceAll]);
      SetControlText ('XX_WHERE', XX_WHERE);
      end;
   end;

//DEB PT19
if ((LeWhere<>nil) and (LeType='D')) then
{PT20
   SetControlText ('XX_WHERE', SQLConf ('SALARIES'));
}
   begin
   St:= SQLConf ('SALARIES');
   if (St<>'') then
      begin
      if (XX_WHERE<>'') then
         XX_WHERE:= XX_WHERE+' AND '+St
      else
         XX_WHERE:= St;
      SetControlText ('XX_WHERE', XX_WHERE);
      end;
   end;
//FIN PT20
//FIN PT19
end;

procedure TOF_PGANALYSEGRH.RendLaDate(D1: TDateTime; var XX_WHERE: string; MoisComplet: TCheckBox);
var
  D2: TDateTime;
  AA, MM, JJ: word;
begin
  DecodeDate(D1, AA, MM, JJ);
  AA := AA - 1; // Annee N-1
  JJ := 1; // Pour se mettre en debut de mois pour trouver la fin de mois pour fevrier
  D2 := EncodeDate(AA, MM, JJ);
  D2 := FindeMois(D2);
  if MoisComplet.Checked then
    XX_WHERE := XX_WHERE + ' OR (PHC_DATEFIN >= "' + UsDateTime(DEBUTDEMOIS(D2)) + '" AND PHC_DATEFIN <= "' + UsDateTime(FINDEMOIS(D2)) + '")'
  else XX_WHERE := XX_WHERE + ' OR PHC_DATEFIN = "' + UsDateTime(D2) + '"';
  if LeType = 'D' then XX_WHERE := StringReplace(XX_WHERE, 'PHC', 'PHB', [rfReplaceAll]);
end;
// FIN PT1
procedure TOF_PGANALYSEGRH.LancerEtatAnalyse(Requete: string);
var
  TobEtat, T: Tob;
  Q, QTablette: TQuery;
  i, a: Integer; // NbChampSal,
  CodeChamp, TypeChamp, LeChamp, ValeurChamp, Presentation, LaRupture, TabletteRupt: string;
  TabChamp, TabTablette: array[1..10] of string;
  Pages: TPageControl;
  DateMois: TDateTime;
begin
  //gestion rupture
  LaRupture := '';
  SetcontrolText('LIBELLERUPTURE', '');
  for i := 1 to 4 do
  begin
    if TRadioButton(GetControl('R_TRAVAILN' + IntToStr(i))).Checked = True then
    begin
      LaRupture := 'PSA_TRAVAILN' + IntToStr(i);
      SetControltext('LIBELLERUPTURE', THLabel(GetControl('TPSA_TRAVAILN' + IntToStr(i))).Caption);
      TabletteRupt := 'PGTRAVAILN' + IntToStr(i);
    end;
    if TRadioButton(GetControl('R_LIBREPCMB' + IntToStr(i))).Checked = True then
    begin
      LaRupture := 'PSA_LIBREPCMB' + IntToStr(i);
      SetControltext('LIBELLERUPTURE', THLabel(GetControl('TPSA_LIBREPCMB' + IntToStr(i))).Caption);
      TabletteRupt := 'PGLIBREPCMB' + IntToStr(i);
    end;
  end;
  if TRadioButton(GetControl('R_CODESTAT')).Checked = True then
  begin
    LaRupture := 'PSA_CODESTAT';
    SetControltext('LIBELLERUPTURE', THLabel(GetControl('TPSA_CODESTAT')).Caption);
    TabletteRupt := 'PGCODESTAT';
  end;
  Pages := TPageControl(GetControl('PAGES'));
//  NbChampSal := 0;
  //Affichage des libellés des colonnes et champs salariés
  Q := OpenSql('SELECT * FROM PARAMEXTPAIE WHERE ##PPE_PREDEFINI## PPE_NUMORDRE =' + GetControlText('PARAMEXTPAIE') + '', TRUE); // DB2
  if not Q.Eof then
  begin
    Presentation := Q.FindField('PPE_PGPRESENTATION').AsString;
    for i := 1 to 10 do
    begin
      LeChamp := '';
      CodeChamp := IntToStr(Q.FindField('PPE_ZON' + IntToStr(i)).AsInteger); // DB2
      if CodeChamp <> '' then
      begin
//        NbChampSal := NbChampSal + 1;
        QTablette := OpenSQL('SELECT * FROM PAIEPARIM WHERE PAI_IDENT=' + CodeChamp + '', true); // DB2
        if not QTablette.Eof then
        begin
          LeChamp := QTablette.FindField('PAI_PREFIX').AsString + '_' + QTablette.FindField('PAI_SUFFIX').AsString;
          TypeChamp := QTablette.FindField('PAI_LETYPE').AsString;
          if typeChamp = 'T' then TabTablette[i] := QTablette.FindField('PAI_LIENASSOC').AsString
          else TabTablette[i] := '';
          SetControlText('INFOSSAL' + IntToStr(i), QTablette.FindField('PAI_LIBELLE').AsString);
        end;
        Ferme(QTablette);
      end
      else SetControlText('INFOSSAL' + IntToStr(i), '');
      TabChamp[i] := LeChamp;
    end;
    if Presentation = 'PER' then
    begin
      for i := 1 to 12 do
      begin
        SetControlText('LIBCOL' + IntToStr(i), Q.FindField('PPE_ETLIB' + IntToStr(i)).AsString);
      end;
    end
    else
    begin
      DateMois := StrToDate(GetControlText('DATEDEBUT'));
      for i := 1 to 12 do
      begin
        SetControlText('LIBCOL' + IntToStr(i), FormatDateTime('mm/yyyy', DateMois));
        DateMois := PlusMois(DateMois, 1);
      end;

    end;
  end;
  Ferme(Q);
  //Préparation de la TOB de l'édition
  Q := OpenSQL(Requete, True);
  TobEtat := Tob.Create('Edition', nil, -1);
  TobEtat.LoadDetailDB('Edition', '', '', Q, False);
  Ferme(Q);
  for i := 0 to TobEtat.Detail.Count - 1 do
  begin
    T := TobEtat.Detail[i];
    for a := 1 to 10 do
    begin
      if TabChamp[a] <> '' then
      begin
// DEB PT12
        if (T.GetValue(TabChamp[a]) <> NULL) then ValeurChamp := T.GetValue(TabChamp[a])
        else ValeurChamp := '';
// FIN PT12
        if TabTablette[a] <> '' then T.AddchampSupValeur('LIBELLE' + IntToStr(a), RechDom(TabTablette[a], ValeurChamp, False))
        else T.AddchampSupValeur('LIBELLE' + IntToStr(a), ValeurChamp, False);
      end
      else T.AddchampSupValeur('LIBELLE' + IntToStr(a), '');
    end;
    for a := 1 to 12 do
    begin
      if not (TobEtat.Detail[i].FieldExists('CH' + IntToStr(a))) then T.AddchampSupValeur('CH' + IntToStr(a), 0);
    end;
    if LaRupture <> '' then
    begin
      T.AddchampSupValeur('RUPTURE1', TobEtat.Detail[i].GetValue(LaRupture), False);
      T.AddchampSupValeur('LIBRUPTURE1', RechDom(tabletteRupt, TobEtat.Detail[i].GetValue(LaRupture), False), False);
    end
    else
    begin
      T.AddchampSupValeur('RUPTURE1', '', False);
      T.AddchampSupValeur('LIBRUPTURE1', '', False);
    end;
  end;
  TobEtat.Detail.Sort('RUPTURE1;PSA_SALARIE');
  //Début PT17
  //  if GetControlText('LISTEEXPORT') = 'X' then LanceEtatTOB('E', 'PAG', 'PAG', TobEtat, True, True, False, Pages, '', '', False)
  //  else LanceEtatTOB('E', 'PAG', 'PAG', TobEtat, True, False, False, Pages, '', '', False);
  LanceEtatTOB('E', 'PAG', 'PAG', TobEtat, True, False, False, Pages, '', '', False);
  //Fin PT17

  TobEtat.Free;
end;
// PT10  : 11/05/2004 PH V_50 FQ 10659 Edition pour les cubes au mode paysage
procedure TOF_PGANALYSEGRH.ChangeLeMode(Sender: TObject);
begin
  if GetControlText('MODEPAYSAGE') = 'X' then
  begin
    TFCube(Ecran).NatureEtat := 'PAY';
    TFCube(Ecran).CodeEtat := 'PSG';
  end
  else
  begin
    TFCube(Ecran).NatureEtat := '';
    TFCube(Ecran).CodeEtat := '';
  end;
end;
procedure TOF_PGANALYSEGRH.ExitCheck(Sender: TObject);
begin
  if LaCheck <> nil then
  begin
    if LaCheck.Checked then SetControlEnabled('PSA_MOTIFSUSPPAIE', TRUE)
    else
    begin
      SetControlText('PSA_MOTIFSUSPPAIE','');                //PT10
      SetControlEnabled('PSA_MOTIFSUSPPAIE', FALSE);
    end;
  end;
end;

initialization
  registerclasses([TOF_PGANALYSEGRH]);
end.

