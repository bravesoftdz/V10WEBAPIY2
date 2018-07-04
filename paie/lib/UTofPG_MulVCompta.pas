
{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion du multi critère de lancement de la
Suite ........ : génération des écritures comptables
Mots clefs ... : PAIE;GENERCOMPTA
*****************************************************************}
{
 PT1 11/10/01 PH Raffraichissement de la liste avant lancement du traitement
 PT2 20/11/01 PH Bornes Etablissements dans le cas du multi
 PT3 13/06/03 V4_21 PH FQ 10674 Non traitement des paies négatives si non traitées en compta
 PT4 13/08/03 V4_21 PH FQ 10692 Test si période comptable ouverte
 PT5 12/01/04 V4_21  PH On pré-coche par défaut l'autorisation des paies négatives
 PT6 12/01/04 V5_00  PH Rajout test paies déjà comptabilisées FQ 11256
 PT7 02/09/04 V5_00  PH FQ 11575 ergonomie
 PT8 04/10/04 V5_00  PH FQ 11662 Rajout du code statistique comme  critère de sélection
 PT9 15/04/04 V6_00  PH FQ 11769 Ergonomie
 PT9 18/05/05 V6_00  PH FQ 11247 Initialisation de la combo
 PT10 18/05/05 V6_00 PH FQ 11791 gestion de la liste vide
 PT11 29/09/05 V6_00 PH FQ 12382 Passage en paramètre du critère salarie du XX_WHERE
 PT12 29/09/05 V6_00 PH FQ 11979 Récup de la période en cours au lieu du mois de la date système
 PT13 15/01/07 V7_00 FCO Mise en place filtrage des habilitations/poupulations
 PT14 29/06/07 V7_00 PH FQ 14401 AccessVIO sur la Query/TOB en WEBACCESS.
 PT15 02/07/07 V7_00 PH FQ 13701 Ergonmie titre écran dynamque en fonction de la gestion analytique
 PT16 10/07/07 V7_20 FC FQ 14527 Liste des salariés vides, si des habilitations sont paramétrées

}
unit UTofPG_MulVCompta;

interface
uses StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
  Hqry,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HDB,
  DBCtrls,
  Mul,
  Fe_Main,
  DBGrids,
  {$ELSE}
  MaineAgl, eMul,
  {$ENDIF}
  Grids,
  HCtrls,
  HEnt1,
  Ent1,
  EntPaie,
  HMsgBox,
  UTOF,
  UTOB,
  UTOM,
  Vierge,
  P5Util,
  P5Def,
  AGLInit;

type
  TOF_PGMULVCOMPTA = class(TOF)
  private
    CbxMul: THValComboBox;
    LeTitre, CritSal: string; // PT11
    Date1, Date2: THEdit;
    DebExer, FinExer: TDateTime;
    procedure ActiveWhere(Sender: TObject);
    procedure LanceFichePrep(Sender: TObject);
    procedure DateDebutExit(Sender: TObject);
    procedure DateFinExit(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override ; //PT16
  end;

implementation

uses PgOutils2;

procedure TOF_PGMULVCOMPTA.ActiveWhere(Sender: TObject);
var
  DD, DF, WW: THEdit;
  Dat1, Dat2: string;
  Where2: THEdit;     //PT13
begin
  LeTitre := Ecran.Caption;
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  WW := THEdit(GetControl('XX_WHERE'));
  Dat1 := UsDateTime(StrToDate(DD.Text));
  Dat2 := UsDateTime(StrToDate(DF.Text));
  if (DD <> nil) and (WW <> nil) then
  begin
    WW.Text := '(PPU_DATEDEBUT >="' + Dat1 + '") AND (PPU_DATEFIN <="' + Dat2 + '")';
  end else exit;
  //  PT3 13/06/03 V4_21 PH Non traitement des paies négatives si non traitées en compta
  if (not VH^.MontantNegatif) and (GetControlText('DROITNEG') = '-') then WW.Text := WW.Text + '  AND PPU_CNETAPAYER >= 0 ';
  //  PT2 20/11/01 PH Bornes Etablissements dans le cas du multi
  if GetControlText('VCBXMULTI') <> 'MON' then
  begin
    if GetControlText('ETABLISSEMENT') <> '' then
      WW.Text := WW.Text + ' AND PPU_ETABLISSEMENT >= "' + GetControlText('ETABLISSEMENT') + '"';
    if GetControlText('ETABLISSEMENT_') <> '' then
      WW.Text := WW.Text + ' AND PPU_ETABLISSEMENT <= "' + GetControlText('ETABLISSEMENT_') + '"';
  end
  else // Cas un seul etablissement
  begin
    if GetControlText('ETABLISSEMENT') <> '' then
    begin
      WW.Text := WW.Text + ' AND PPU_ETABLISSEMENT >= "' + GetControlText('ETABLISSEMENT') + '"';
      WW.Text := WW.Text + ' AND PPU_ETABLISSEMENT <= "' + GetControlText('ETABLISSEMENT') + '"';
    end;
  end;
//  PT5 12/01/04 V5_00  PH Rajout test paies déjà comptabilisées FQ 11256
  if GetControlText ('TOPGENECR') = 'X' then
      WW.Text := WW.Text + ' AND PPU_TOPGENECR <> "X" ';
// DEB PT9 Sélection des salariés sortis ou non
  if (GetControltext ('CBXSALSORTIS') <> 'SAL') then
  begin
  // DEB PT11
    if GetControltext ('CBXSALSORTIS') = 'SAN' then // Sans les salariés sortis
        CritSal := ' AND (PSA_DATESORTIE >"'+Dat2+'" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+UsDateTime(Idate1900)+'")'
    else // Uniquement les salariés sotis dans la période
        CritSal := ' AND (PSA_DATESORTIE >="'+Dat1+'" AND PSA_DATESORTIE <="'+Dat2+'")';
    WW.Text := WW.Text + CritSal;
  // FIN PT11
  end;
// FIN PT9

  //DEB PT13
  Where2 := THEdit(GetControl('XX_WHERE2'));
  //DEB PT16
  if Assigned(MonHabilitation) then
    if ('PSA_SALARIE' <> copy(MonHabilitation.LeSQL, 1, 11)) and (MonHabilitation.LeSQL <> '') then
    begin
      SetControlText('LEPREFIXE', 'PSA');
      MonHabilitation.LeSQL := '';
      PGAffecteEtabByUser(TFMul(Ecran));
    end;
  //FIN PT16

  if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL <> '') then
    if Where2 <> nil then SetControlText('XX_WHERE2',MonHabilitation.LeSQL);
  //FIN PT13
end;

procedure TOF_PGMULVCOMPTA.DateDebutExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de début est supérieure à la date de fin', LeTitre);
      SetFocusControl('DATEDEBUT');
    end;
    if StrToDate(DD.Text) < DebExer then
    begin // PT9
      PGIBox('Attention, la date de début est inférieure à la date de début d''exercice', LeTitre);
      SetFocusControl('DATEFIN');
    end;
  end;
end;

procedure TOF_PGMULVCOMPTA.DateFinExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de fin est inférieure à la date de début', LeTitre);
      SetFocusControl('DATEDEBUT');
    end;
    if StrToDate(DF.Text) > FinExer then
    begin  // PT9
      PGIBox('Attention, la date de fin est supérieure à la date de fin d''exercice', LeTitre);
      SetFocusControl('DATEFIN');
    end;
  end;
end;

procedure TOF_PGMULVCOMPTA.LanceFichePrep(Sender: TObject);
var
  st, St1, LeWhere: string;
  DD, DF: THEdit;
  LadateD, LadateF, DateDeb, DateFin: TDateTime;
  rep: Integer;
  Q: TQuery;
  Etab: THValComboBox;
  BtnCherche: TToolbarButton97;
begin
  // PT1 11/10/01 PH Raffraichissement de la liste avant lancement du traitement
  BtnCherche := TToolbarButton97(GetControl('BCherche'));
  if BtnCherche <> nil then BtnCherche.Click;

  LeWhere := RecupWhereCritere(TFMul(Ecran).Pages);
  Etab := THValComboBox(GetControl('ETABLISSEMENT'));
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) and (Etab <> nil) and (CbxMul <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de début est supérieure à la date de fin', LeTitre);
      SetFocusControl('DATEDEBUT');
      exit;
    end;
    LaDateD := DEBUTDEMOIS(DateDeb);
    LaDateF := FINDEMOIS(DateFin);
    rep := mrYes;
    if (FINDEMOIS(LaDateD) <> LaDateF) then
      rep := PGIAsk('Attention, la période de paie est supérieure à 1 mois ou à cheval sur 2 mois', LeTitre);
    // Controle des paies pour voir si elles sont déjà passées en compta
    if rep = mrYes then
    begin // Test si les écritures ont déjà été ventilées
      st := 'Select PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEDEBUT>="' + USDateTime(StrToDate(DD.Text)) + '" AND PPU_DATEFIN<="' + USDateTime(StrToDate(DF.Text)) +
        '" AND PPU_TOPGENECR="X"';
      Q := OpenSql(st, FALSE);
      if not Q.Eof then
        rep := PGIAsk('Attention, vous avez déjà comptabilisé des paies sur la même période, voulez vous continuer ?', LeTitre);
      Ferme(Q);
    end;
    if rep = mrYes then
    begin
      if (CbxMul.Value = 'MON') and (Etab.value = '') then
      begin
        PGIBox('Vous devez saisir un établissement !', Ecran.Caption);
        rep := mrNo;
      end;
      if (CbxMul.Value = 'MUL') then
      begin
        if VH^.EtablisDefaut = '' then
        begin
          PGIBox('Vous devez obligatoirement renseigner un établissement principal dans la comptabilité', Ecran.Caption);
          rep := mrNo;
        end;
        // PT7 ergonomie dans contenu du message
        if rep = mrYes then rep := PGIAsk('Vous allez générer vos ODs pour tous vos établissements, établissement par établissement !', Ecran.Caption);
      end;
      if (CbxMul.Value = 'PRI') then rep := PGIAsk('Vous allez générer vos ODs pour tous vos établissements sur l''établissement principal !', Ecran.Caption);
    end;
    if rep = mrYes then
    begin
      st := DD.Text + ';' + DF.Text + ';' + Etab.Value + ';' + CbxMul.Value + ';'+ LeWhere;
      if CritSal <> '' then st := st +';'+CritSal; // PT11
      //  PT4 13/08/03 V4_21 PH FQ 10692 Test si période comptable ouverte
      if VH_Paie.PGIntegODPaie <> 'ECP' then
      begin
        St1 := 'SELECT EX_ETATCPTA FROM EXERCICE WHERE EX_DATEDEBUT <= "' + UsDateTime(StrToDate(DD.Text)) + '" AND EX_DATEFIN >="' + UsDateTime(StrToDate(DF.text)) +
          '" AND EX_ETATCPTA="OUV"';
        Q := OpenSql(St1, true);
        if Q.Eof then
        begin
          ferme(Q);
          PgiBox('La période de paie ne correspond pas à une période comptable disponible, #13#10 Traitement impossible ', Ecran.Caption);
          exit;
        end;
      end;
      {$IFDEF EAGLCLIENT}
      if TFMul(Ecran).Fetchlestous then
          TheMulQ:=TOB(Ecran.FindComponent('Q'))
      else
      begin
        PgiBox('Vous n''avez pas de ligne total dans votre liste, #13#10 Traitement impossible ', Ecran.Caption);
        exit;
      end;
      {$ELSE}
      TheMulQ := THQuery(Ecran.FindComponent('Q'));
      {$ENDIF}
// DEV PT10  PT14
      {$IFDEF EAGLCLIENT}
      if TOB (TFMul(Ecran).Q.TQ).RecordCount > 0
      {$ELSE}
      if NOT TheMulq.EOF
      {$ENDIF}
      then AglLanceFiche('PAY', 'GENECOMPTA', '', '', st)
      else PGIBOX ('Vous n''avez sélectionné aucune paie', Ecran.Caption);
// FIN PT10 PT14
      // FIN PT4
      TheMulQ := nil;
    end;
  end;
end;

procedure TOF_PGMULVCOMPTA.OnArgument(Arguments: string);
var
  Num: Integer;
  BtnValidMul: TToolbarButton97;
  MoisE, AnneeE, ComboExer: string;
  CbxEtab: THValComboBox;
  DateDeb, DateFin : TDateTime; // PT12
begin
  inherited;
  // DEB PT15
  if not VH_Paie.PGAnalytique then
  begin
    Ecran.Caption := TraduireMemoire ('Génération des écritures comptables');
    UpdateCaption(Ecran);
  end;
  // FIN PT15
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PPU_TRAVAILN' + IntToStr(Num)), GetControl('TPPU_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PPU_CODESTAT'), GetControl('TPPU_CODESTAT'));       // PT8
  //  PT3 13/06/03 V4_21 PH Non traitement des paies négatives si non traitées en compta
  if not VH^.MontantNegatif then
  begin
    SetControlVisible('MESSNEG', TRUE);
    SetControlVisible('DROITNEG', TRUE);
    //  PT5 12/01/04 V_50  PH On pré-coche par défaut l'autorisation des paies négatives
    SetControlChecked('DROITNEG', FALSE);
  end;

  BtnValidMul := TToolbarButton97(GetControl('BOuvrir'));
  if BtnValidMul <> nil then BtnValidMul.OnClick := LanceFichePrep;
  Date1 := THEdit(GetControl('DATEDEBUT'));
  if Date1 <> nil then Date1.OnClick := DateDebutExit;
  Date2 := THEdit(GetControl('DATEFIN'));
  if Date2 <> nil then Date2.OnClick := DateFinExit;
  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
  // DEB PT12
  DateDeb := EncodeDate (StrToInt(AnneeE),StrToInt(MoisE),01);
  DateFin := FinDeMois (DateDeb);
  Date2.text := DateToStr(DateFin);
  Date1.text := DateToStr(DateDeb);
  // FIN PT12
  if StrToDate(Date2.Text) > FinExer then
  begin
    Date2.Text := DateToStr(FinExer);
    Date1.Text := DateToSTr(DEBUTDEMOIS(FinExer));
  end;
  CbxMul := THValComboBox(GetControl('VCBXMULTI'));
  if CbxMul <> nil then CbxMul.Value := VH_Paie.PgVentilMulEtab;
  CbxEtab := THValComboBox(GetControl('ETABLISSEMENT'));
  if CbxEtab <> nil then
  begin
    CbxEtab.Vide := TRUE;
    CbxEtab.VideString := 'Tous';
    CbxEtab.DataType := 'TTETABLISSEMENT';
  end;
  SetControlVisible('BParamListe', TRUE);
  SetControltext ('CBXSALSORTIS', 'SAL'); // PT9 Initialisation à tous par défaut
end;


procedure TOF_PGMULVCOMPTA.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;

//DEB PT16
procedure TOF_PGMULVCOMPTA.OnUpdate;
begin
  inherited;
  ActiveWhere(nil);
end;
//FIN PT16

initialization
  registerclasses([TOF_PGMULVCOMPTA]);
end.

