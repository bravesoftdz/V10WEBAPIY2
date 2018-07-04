{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/05/2004
Modifié le ... : 16/06/2004
Description .. : Source TOF de la FICHE : PG_CALINDICBS ()
Suite ........ :
Suite ........ : Fonction de calculs des indicateurs sélectionnés dans
Suite ........ : une présentation
Mots clefs ... : TOF;PG_CALINDICBS
*****************************************************************}
{
PT1 10/02/2006 SB V650 FQ 12905 Refonte indicateur absence
PT2 10/02/2006 SB V650 FQ 12902 Refonte date sortie sur Tranche Age et Anc
PT3 02/05/2006 SB V_65 FQ 12480 Nouveaux indicateurs
PT4 17/05/2006 PH V_65 FQ 12965 respect de la syntaxe SQL pour la traduction des fonctions SQL
PT5 19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT6 28/09/2006 SB V_65 FQ 13493 & FQ 13494 Correction anomalie <> indicateur
PT7 12/10/2006 SB V_65 Refonte calcul indicateurs effectif (30,31,32,33,40)
PT8 16/10/2006 SB V_65 FQ 12903 Calcul de l'indicateur Nbre d'embauches moins 25 ans
}
unit UTOFPG_CALINDICBS;

interface

uses StdCtrls,
  Controls,
  Classes,
  Windows,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  HSysMenu,
  HTB97,
  UTOB,
  UTOF,
  Ed_tools,
  Vierge,
  HDebug;

type
  TChampRupt = record
    LIBELLEEMPLOI, QUALIFICATION, COEFFICIENT, CODESTAT, TRAVAILN1,
      TRAVAILN2, TRAVAILN3, TRAVAILN4, LIBREPCMB1, LIBREPCMB2,
      LIBREPCMB3, LIBREPCMB4, ETABLISSEMENT, CATBILAN: string; // Champs ruptures possibles
  end;

type
  TTypeIndDate = (TNombre, TMoyenne);
  TTypeIndRem  = (TMoins, TPlus);
  TTypeIndTrav = (TNbInt, TNbMoyenInt, TDureeContInt, TNbMoyenContInt, TNbStag);
  TTypeCalcul  = (TCompte, TSomme, TMoyen, TDecembre );  { PT1 PT3 }
  TTypeSelect  = (TMotif, TTypeAbs);  { PT1 }

type
  TOF_PG_CALINDICBS = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    TOBInd, TOBIndSel  {$IFDEF DEBUGBILAN}, Tob_temp {$ENDIF}: TOB;
    Presentation, Theme: string;
    DateDeb, DateFin: TDateTime;
    TablS: array[1..12] of string;
    ListeError: TListBox;
    GblRupt: TChampRupt;
    procedure ChangePresentation(Sender: TObject);
    procedure CalculIndic(Sender: TObject);
    procedure CalculIndicateur(Indic: string; TOBCalcul: TOB);
    procedure CalculIndCumul(TypeSql : TTypeCalcul; Quoi, Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string); { PT3 }
    procedure CalculIndRub(TypeSql : TTypeCalcul; LeChamp, Quoi, Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string); { PT3 }
    procedure CalculIndAbs(Quoi, LesVal, Indic, Sexe, Nat: string;  TypeSql : TTypeCalcul; TOBCalcul, Ti: TOB; Racine: string; TypeSelect : TTypeSelect); { PT1 }
    procedure CalculIndEffFinPer(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB);
    procedure CalculIndEffMoyen(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB);
    procedure CalculIndNbMoyenDate(TypeTrait: TTypeIndDate; TypeCal, Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
    procedure CalculDeparts(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
    procedure CalculNbMutEtab(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
    procedure Calcul10RemPlusMoinsElev(TypeTrait: TTypeIndRem; Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB; RuptEtab: Boolean);

    procedure CalculTravExt(TypeCalTrav: TTypeIndTrav; Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
    procedure CalculNbInt(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
    procedure CalculMoyenInt(TypeCalTrav: TTypeIndTrav; Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
    procedure CalculDureeContInt(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
    procedure CalculNbStag(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);

    procedure RempliDef(Indic: string; T1: TOB);
    function  RendSexeNat(Sexe, Nat: string; Pref: string = 'PSA'): string;
    function  RendRupture(Indicat: TOB; Racine: string): string;
    procedure RempliCompl(Indic: string; TQ, T1, Ti: TOB; Origine, TheRacine: string );
    function  TestRupture(Racine: string; T1: TOB; var Rupt: TChampRupt): boolean;
    function  CompareRuptOk(T1: TOB; var Rupt: TChampRupt; LeChamp: string; i: integer): Boolean;
    procedure AffecteChampRupt(T1: TOB; var Rupt: TChampRupt; LeChamp: string; i: integer; Valdef: Boolean);
    procedure RempliComplAvecRupt(Indic: string; var Rupt: TChampRupt; T1: TOB; Nbre: Double; NbMois, NbArrondi: Integer);
    procedure CalculIndContrat(Quoi, Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine, TypInd: string; Moins25 : Boolean = False); { PT8 }
    procedure CalculIndVisiteMed(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string);
    procedure CalculNbreHandicape(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string);
    procedure CalculMTCompta(Quoi, Indic: string; TOBCalcul, Indicat: TOB; TypeMT: string);
    procedure RempliIndCompta(Indic: string; Q: TQuery; T1: TOB; Origine: string);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function  RendNbPeriode(DD, DF: TDateTime): integer;
    function  RendListeSal (TheRacine,Sexe, Nat,StChamp : String; Compte,WithSal,Distinct :Boolean) :  String; { PT7 }
  end;


implementation

uses P5Def, pgoutils2;

var
  HistS: array[1..12] of string = (
    'LIBELLEEMPLOI', 'QUALIFICATION', 'COEFFICIENT', 'CODESTAT', 'TRAVAILN1',
    'TRAVAILN2', 'TRAVAILN3', 'TRAVAILN4', 'LIBREPCMB1', 'LIBREPCMB2',
    'LIBREPCMB3', 'LIBREPCMB4');

procedure TOF_PG_CALINDICBS.OnNew;
begin
  inherited;
end;

procedure TOF_PG_CALINDICBS.OnDelete;
begin
  inherited;
end;

procedure TOF_PG_CALINDICBS.OnUpdate;
begin
  inherited;
end;

procedure TOF_PG_CALINDICBS.OnLoad;
begin
  inherited;
end;

procedure TOF_PG_CALINDICBS.OnArgument(S: string);
var
  BTNLANCE: TToolbarButton97;
  Presentat: THValComboBox;
begin
  inherited;
  TOBInd := TOB.Create('TOB DES INDICATEURS', nil, -1);
  TOBInd.LoadDetailDB('PBSINDICATEURS', '', 'PBI_INDICATEURBS', nil, FALSE, False);
  BTNLANCE := TToolbarButton97(GetControl('BTNLANCE'));
  if BTNLANCE <> nil then BTNLANCE.OnClick := CalculIndic
  else SetControlVisible('BTNLANCE', FALSE);
  Presentat := THValComboBox(GetControl('PRESENTATION'));
  if Presentat <> nil then Presentat.OnChange := ChangePresentation;
  TFVierge(Ecran).OnKeyDown := FormKeyDown;
end;

procedure TOF_PG_CALINDICBS.OnClose;
begin
  inherited;
  FreeAndNil(TOBInd);
  FreeAndNil(TOBIndSel);
end;

procedure TOF_PG_CALINDICBS.OnDisplay();
begin
  inherited;
end;

procedure TOF_PG_CALINDICBS.OnCancel();
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 16/06/2004
Modifié le ... :   /  /
Description .. : Fonction de lancement du calcul des indicateurs
Suite ........ :
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndic(Sender: TObject);
var
  nb, i, rep, compteur: Integer;
  Indic, MIndic, St, StDelete: string;
  GrilleSelect: THGRID;
  TOBCalcul, T1, Tob_Exist: TOB;
  Q: TQuery;
  ToutSelect, NonModifiable: Boolean;
begin
  ToutSelect := False;
  StDelete := '';
  GrilleSelect := THGRID(GetControl('GRILLESELECT'));
  if (GrilleSelect = nil) then exit;
  ListeError := TListBox(GetControl('LBANOMALIE'));
  if ListeError = nil then exit;
  ListeError.Items.Clear;
  rep := PgiAsk('Voulez-vous calculer les indicateurs de la présentation sélectionnée?', Ecran.Caption);
  if rep <> mrYes then exit;
  InitMoveProgressForm(nil, 'Traitement en cours. Veuillez patienter..', Ecran.Caption, GrilleSelect.RowCount, False, True);
  DateDeb := StrtoDate(GetControltext('DATEDEB'));
  DateFin := StrtoDate(GetControltext('DATEDEB_'));
  TOBCalcul := TOB.create('MON BILAN SOCIAL', nil, -1);
  { On récupère historiquement les indicateurs de type Non calculable, afin de ne pas les recalculer }
  St := 'SELECT DISTINCT PBC_INDICATEURBS,PBC_PGINDICATION FROM BILANSOCIAL ' +
    'WHERE PBC_BSPRESENTATION="' + Presentation + '" ' +
    'AND PBC_DATEDEBUT="' + USDateTime(DateDeb) + '" ' +
    'AND PBC_DATEFIN="' + USDateTime(DateFin) + '" ';
  Q := OpenSql(st, True);
  //    'AND PBC_PGINDICATION="NON"';
  if not Q.eof then
  begin
    Tob_Exist := TOB.create('BILAN SOCIAL DEJA CALCULE', nil, -1);
    Tob_Exist.LoadDetailDB('BILAN SOCIAL DEJA CALCULE', '', '', Q, False);
  end
  else
    FreeAndNil(Tob_Exist);
  Ferme(Q);

  { contrôle nombre de salariés présents }
  St := 'SELECT COUNT(*) NB FROM SALARIES WHERE PSA_CATBILAN=""';
  Q := OpenSql(st, True);
  if not Q.eof then
  begin
    nb := Q.Findfield('NB').AsInteger;
    if nb <> 0 then
      if nb = 1 then ListeError.items.Add('Un salarié n''est pas affecté à une catégorie bilan social')
      else ListeError.items.Add(IntToStr(nb) + ' salariés ne sont pas affectés à une catégorie bilan social');
  end;
  Ferme(Q);
  {$IFDEF DEBUGBILAN}
  TOb_Temp := Tob.create('Test',nil,-1);
  {$ENDIF}
  { Calcul des indicateurs }
  if GrilleSelect.nbSelected > 0 then
  begin
    for i := 1 to GrilleSelect.RowCount - 1 do
      if GrilleSelect.IsSelected(i) then
      begin
        NonModifiable := False;
        //GrilleSelect.GotoLeBookmark(i);
        { Recherche si indicateur présauvegardé existe et est de type "Non Calculable" }
        if Assigned(Tob_Exist) then
        begin
          T1 := Tob_Exist.FindFirst(['PBC_INDICATEURBS'], [GrilleSelect.Cells[1, i]], False);
          if Assigned(T1) then
            NonModifiable := (T1.GetValue('PBC_PGINDICATION') <> 'CAL');
        end
        else FreeAndNil(T1);
        { Ne traite pas les indicateurs "Non calculable" }
        if NonModifiable then
        begin
          ListeError.items.Add('L''indicateur ' + GrilleSelect.Cells[1, i] + ' ne peut être recalculé, il n''est pas définit comme "calculable"');
          Continue;
        end;
        StDelete := StDelete + ',"' + GrilleSelect.Cells[1, i] + '"';
        CalculIndicateur(GrilleSelect.Cells[1, i], TOBCalcul); // Fonction de calcul
        MoveCurProgressForm(GrilleSelect.Cells[2, i]);
      end;
  end
  else
    for i := 1 to GrilleSelect.RowCount - 1 do
    begin
      ToutSelect := True;
      NonModifiable := False;
      { Recherche si indicateur présauvegardé existe et est de type "Non Calculable" }
      if Assigned(Tob_Exist) then
      begin
        T1 := Tob_Exist.FindFirst(['PBC_INDICATEURBS'], [GrilleSelect.Cells[1, i]], False);
        if Assigned(T1) then
          NonModifiable := (T1.GetValue('PBC_PGINDICATION') <> 'CAL');
      end
      else FreeAndNil(T1);
      { Ne traite pas les indicateurs "Non calculable" }
      if NonModifiable then
      begin
        ListeError.items.Add('L''indicateur ' + GrilleSelect.Cells[1, i] + ' ne peut être recalculé, il n''est pas définit comme "calculable"');
        Continue;
      end;
      CalculIndicateur(GrilleSelect.Cells[1, i], TOBCalcul); // Fonction de calcul
      MoveCurProgressForm(GrilleSelect.Cells[2, i]);
    end;
  {$IFDEF DEBUGBILAN}
  Tob_Temp.Detail.Sort('SALARIE');
  Tob_Temp.SaveToFile('C:\Test.xls',False,False,True);
  FreeAndNil(Tob_Temp);
  {$ENDIF}

  rep := PgiAsk('Voulez-vous sauvegarder les indicateurs calculés?', Ecran.Caption);
  if rep <> mrYes then
  begin
    FreeAndNil(TOBCalcul);
    FiniMoveProgressForm;
    exit;
  end;
  { Boucle de reaffection de l'increment qui sert de dédoublonnage }
  Compteur := 0;
  for i := 0 to TOBCalcul.detail.count - 1 do
  begin
    T1 := TOBCalCul.Detail[i];
    Indic := T1.GetValue('PBC_INDICATEURBS');
    if (Indic = Mindic) and (i > 0) then Compteur := Compteur + 1
    else
    begin
      Compteur := 0;
      Mindic := Indic;
    end;
    T1.PutValue('PBC_INCREM', Compteur); // Compteur de dédoublonnage
    T1.PutValue('PBC_NODOSSIER', V_PGI.NODOSSIER); // Affectation du numéro de dossier ou de soc
  end;
  //  PGVisuUnObjet(TOBCalcul, '', '');

  FreeAndNil(Tob_Exist);

  { Mise à jour des données }
  try
    BeginTrans;
    MoveCurProgressForm('Mise à jour des données. Veuillez patienter..');
    st := 'DELETE FROM BILANSOCIAL WHERE PBC_DATEDEBUT="' + USDateTime(DateDeb) + '" AND PBC_DATEFIN="' +
      USDateTime(DateFin) + '" AND PBC_BSPRESENTATION="' + Presentation + '" AND PBC_NODOSSIER="' + PgRendNoDossier() + '" ' +
      'AND PBC_PGINDICATION="CAL"';
    if not ToutSelect then
      St := St + ' AND PBC_INDICATEURBS IN (' + Copy(StDelete, 2, length(StDelete)) + ') ';
    ExecuteSQL(st);
    TOBCalcul.InsertDB(nil, FALSE);
    CommitTrans;
    PGIBox('Le calcul et la validation des indicateurs sont terminés.', Ecran.Caption);
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la validation des indicateurs.', Ecran.Caption);
  end;
  FreeAndNil(TOBCalcul);
  FiniMoveProgressForm;
  GrilleSelect.ClearSelected;


  if ListeError.Items.Count > 0 then
  begin
    TTabSheet(GetControl('TBANOMALIE')).TabVisible := True;
    TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('TBANOMALIE'));
  end;

end;

procedure TOF_PG_CALINDICBS.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  TFVierge(Ecran).FormKeyDown(Sender, Key, Shift);
  case Key of
    VK_F9: CalculIndic(nil);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 16/06/2004
Modifié le ... :   /  /
Description .. : Fonction d'affichage des indicateurs qui vont être calculés
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.ChangePresentation(Sender: TObject);
var
  Q: TQuery;
  st: string;
  GrilleSelect: THGrid;
  i: Integer;
begin
  GrilleSelect := THGrid(GetControl('GRILLESELECT'));
  if not Assigned(GrilleSelect) then exit;
  for i := GrilleSelect.RowCount - 1 downto 1 do GrilleSelect.DeleteRow(i);
  Presentation := GetControlText('PRESENTATION');
  FreeAndNil(TOBIndSel);
  st := 'SELECT * FROM PBSINDDETSEL,PBSINDICATEURS ' +
    'WHERE ##PIL_PREDEFINI## PIL_BSPRESENTATION="' + Presentation + '" ' +
    'AND ##PBI_PREDEFINI##  PIL_INDICATEURBS=PBI_INDICATEURBS';
  Q := OpenSql(st, TRUE);
  TOBIndSel := TOB.Create('TOB DES INDICATEURS SELECT', nil, -1);
  TOBIndSel.LoadDetailDB('Indicateurs sélectionnés', '', '', Q, FALSE);
  FERME(Q);
  TOBIndSel.PutGridDetail(GrilleSelect, False, True, 'PIL_PREDEFINI;PBI_INDICATEURBS;PBI_LIBELLE;PBI_FOURCHETTEMINI;PBI_FOURCHETTEMAXI', False);
  //  HMTrad.ResizeGridColumns(GrilleSelect);

end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 16/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction d'analyse de l'indicateur qui ventile les calculs
Suite ........ : à faire en fonction du type d'indicateurs
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndicateur(Indic: string; TOBCalcul: TOB);
var
  Ti: TOB;
  LesVal, TypInd, Sexe, Nat: string;
  Cas: Integer;
begin
  Ti := TOBInd.FindFirst(['PBI_INDICATEURBS'], [Indic], FALSE);
  if Ti = nil then
  begin
    PgiError('Attention, l''indicateur ' + Indic + ' a été supprimé, il ne peut pas être calculé!', Ecran.Caption);
    Exit;
  end;
  TypInd := Ti.GetValue('PBI_TYPINDICATBS');
  LesVal := Ti.GetValue('PBI_LESVALEURS');
  Sexe := Ti.GetValue('PBI_SEXE');
  Nat := Ti.GetValue('PBI_NATIONALITE');
  Theme := Ti.GetValue('PBI_BSINDTHEME');
  if TypInd = '' then exit;
  Cas := StrToInt(TypInd);
  case cas of
    1: CalculIndCumul(TSomme,LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHC_');               // Cumuls : Somme
    2: CalculIndCumul(TMoyen,LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHC_');               // Cumuls : Moyenne { PT3 }
    3: CalculIndCumul(TDecembre,LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHC_');            // Cumuls : 31/12 { PT3 }
    5: CalculIndRub(TSomme,'BR', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_');           // Calcul de rémunérations
    6: CalculIndRub(TSomme,'MR', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_');           // Rémunérations : Somme
    7: CalculIndRub(TDecembre,'MR', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_');        // Rémunérations : au 31/12 { PT3 }
    8: CalculIndRub(TMoyen,'MR', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_');           // Rémunérations : Moyen { PT3 }
    10: CalculIndRub(TSomme,'BCOT', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_'); // Calcul de cotisations
    11: CalculIndRub(TSomme,'MSCOT', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_'); // Calcul de cotisations
    12: CalculIndRub(TSomme,'MPCOT', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_'); // Calcul de cotisations
    20: CalculIndRub(TSomme,'BBAS', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_'); // Calcul de bases de cotisations
    21: CalculIndRub(TSomme,'BTR1', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_'); // Calcul de bases de cotisations
    22: CalculIndRub(TSomme,'BTR2', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_'); // Calcul de bases de cotisations
    23: CalculIndRub(TSomme,'BTR3', LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PHB_'); // Calcul de bases de cotisations
    24: CalculIndAbs('S', LesVal, Indic, Sexe, Nat, TCompte, TOBCalcul, Ti, 'PCN_',TMotif); //24 Abs : Nbre de salariés par motif   { PT1 }
    25: CalculIndAbs('S', LesVal, Indic, Sexe, Nat, TCompte, TOBCalcul, Ti, 'PCN_',TTypeAbs); //25 Abs : Nbre de salariés par type  { PT1 }
    26: CalculIndAbs('A', LesVal, Indic, Sexe, Nat, TCompte, TOBCalcul, Ti, 'PCN_',TMotif); //26 Abs : Nbre d'absences par motif    { PT1 }
    27: CalculIndAbs('A', LesVal, Indic, Sexe, Nat, TCompte, TOBCalcul, Ti, 'PCN_',TTypeAbs); //27 Abs : Nbre d'absence par type    { PT1 }
    30: CalculIndNbMoyenDate(TNombre, 'AGE', Indic, Sexe, Nat, TOBCalcul, Ti); // Age
    31: CalculIndNbMoyenDate(TNombre, 'ANC', Indic, Sexe, Nat, TOBCalcul, Ti); // Ancienneté
    32: CalculIndNbMoyenDate(TMoyenne, 'AGE', Indic, Sexe, Nat, TOBCalcul, Ti); // Age
    33: CalculIndNbMoyenDate(TMoyenne, 'ANC', Indic, Sexe, Nat, TOBCalcul, Ti); // Ancienneté
    35: CalculIndAbs('J', LesVal, Indic, Sexe, Nat, TSomme, TOBCalcul, Ti, 'PCN_',TMotif); // Calcul somme jours absences par motif;    { PT1 }
    36: CalculIndAbs('H', LesVal, Indic, Sexe, Nat, TSomme, TOBCalcul, Ti, 'PCN_',TMotif); // Calcul somme heures absences par motif;   { PT1 }
    37: CalculIndAbs('J', LesVal, Indic, Sexe, Nat, TSomme, TOBCalcul, Ti, 'PCN_',TTypeAbs); // Calcul somme jours absences par Type;   { PT1 }
    38: CalculIndAbs('H', LesVal, Indic, Sexe, Nat, TSomme, TOBCalcul, Ti, 'PCN_',TTypeAbs); // Calcul somme heures absences par Type ; { PT1 }
    40: CalculIndEffFinPer(Indic, Sexe, Nat, TOBCalcul, Ti); // Effectif en fin de période
    41: CalculIndEffMoyen(Indic, Sexe, Nat, TOBCalcul, Ti); // Effectif moyen = Effectif totaux des 12 mois de la période / 12
    42: CalculIndContrat('CDI', Indic, Sexe, Nat, TOBCalcul, Ti, 'PCI_', 'EFF'); // Effectif permament  = Nbre de Salariés en CDI sur la totalité de le période
    45: CalculTravExt(TNbInt, Indic, Sexe, Nat, TOBCalcul, Ti); // Nombre d'intérimaires qui ont un emploi dans la période
    46: CalculTravExt(TNbMoyenInt, Indic, Sexe, Nat, TOBCalcul, Ti); // Nombre moyen d'intérimaires
    47: CalculTravExt(TNbMoyenContInt, Indic, Sexe, Nat, TOBCalcul, Ti); // durée moyenne des missions d'intérim
    48: CalculTravExt(TNbStag, Indic, Sexe, Nat, TOBCalcul, Ti); // nombre de stagiaires
    49: CalculTravExt(TDureeContInt, Indic, Sexe, Nat, TOBCalcul, Ti); // Nbre d'intérim. durée moy. contrat
    50: CalculDeparts(Indic, Sexe, Nat, TobCalcul, Ti); // nombre de départs
    51: ListeError.items.Add(indic + ' : Fonction non disponible'); // Nbre de décès
    60: CalculIndContrat(LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PCI_', ''); // Nombre de contrats
    61: ListeError.items.Add(indic + ' : Fonction non disponible'); // Nbre de temps partiel
    65: CalculIndContrat(LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PCI_', 'EMB'); // Nombres d'embauches par type de contrat
    66: CalculIndContrat(LesVal, Indic, Sexe, Nat, TOBCalcul, Ti, 'PCI_', 'EMB',True); // Nbre d'embauches de - 25 ans { PT8 }
    70: CalculNbMutEtab(Indic, Sexe, Nat, TOBCalcul, Ti); // Nombre de mutations d'1 établissement à 1 autre
    100: CalculNbreHandicape(Indic, Sexe, Nat, TOBCalcul, Ti, 'PSA_'); // Nombre d'handicapés au 31/03
    101: ListeError.items.Add(indic + ' : Fonction non disponible'); // Nbre handicapés dûs à un accident du travail dans l'entreprise
    200: Calcul10RemPlusMoinsElev(TPlus, Indic, Sexe, Nat, TOBCalcul, Ti, False); // Montant des 10 rémunérations les + élévées
    202: Calcul10RemPlusMoinsElev(TMoins, Indic, Sexe, Nat, TOBCalcul, Ti, False); // Montant des 10 rémunérations les - élévées
    204: Calcul10RemPlusMoinsElev(TPlus, Indic, Sexe, Nat, TOBCalcul, Ti, True); // Montant des 10 rémunérations les + élévées par etablissement
    206: Calcul10RemPlusMoinsElev(TMoins, Indic, Sexe, Nat, TOBCalcul, Ti, True); // Montant des 10 rémunérations les - élévées par etablissement
    300: CalculIndVisiteMed(Indic, Sexe, Nat, TOBCalcul, Ti, 'PSA_'); // Nombre de visites médicales
    400: CalculMTCompta(LesVal, Indic, TOBCalcul, Ti, 'CRE'); // Montant créditeur
    401: CalculMTCompta(LesVal, Indic, TOBCalcul, Ti, 'DEB'); // Montant débiteur
    998: ListeError.items.Add(indic + ' : Fonction non disponible'); // Chapitre ==> 1 ligne par établissement
    999: ListeError.items.Add(indic + ' : Fonction non disponible'); // Précalculés ?????
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction qui calcule le montant d'un cumul
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndCumul(TypeSql : TTypeCalcul; Quoi, Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string);
var
  T1: TOB;
  st, Scum, SC: string;
  Q: TQuery;
  TheRacine: string;
  i: Integer;
  TMere, TFille : Tob;
  DD,DF : TDateTime;
  YY,MM : WORD;
begin
  TheRacine := RendRupture(Indicat, Racine);
  St := 'SUM'; { DEB PT3 }
  DD := DateDeb;
  DF := DateFin;
  case TypeSql of
     TMoyen : St := 'AVG';
     TDecembre : Begin DecodeDate(DF,YY,MM,MM);
                       DF := EncodeDate(YY,12,31);
                       DD := DebutDeMois(DF);
                 End;
  End;  { FIN PT3 }

  st := 'SELECT '+St+' (PHC_MONTANT) MT,PHC_ETABLISSEMENT,PSA_CATBILAN'; { PT3 }
  st := st + TheRacine +
    ' FROM HISTOCUMSAL LEFT JOIN SALARIES ON PSA_SALARIE=PHC_SALARIE WHERE' +
    ' PHC_DATEFIN <= "' + UsDateTime(DF) + '" AND PHC_DATEDEBUT >= "' + UsDateTime(DD) + '"'; { PT3 }
  Sc := Quoi;
  if (Sc <> '') and (Sc[1] <> '!') then
  begin
    i := 0;
    Scum := ReadTokenst(Sc);
    while (SCum <> '') AND (SCum <> '<<Tous>>') do  { PT6 16/10/06 }
    begin
      if i = 0 then st := st + ' AND ( '  { PT6 }
      else st := st + ' OR ';
      st := st + ' PHC_CUMULPAIE="' + Scum + '"';
      Scum := ReadTokenst(Sc);
      i := i + 1;
      If SCum = '' then St := St + ' ) ';   { PT6 }
    end;
  end
  else st := st + ' AND PHC_CUMULPAIE <> "' + Sc + '"';
  st := st + RendSexeNat(Sexe, Nat);
  st := st + ' GROUP BY PHC_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;
  Q := OPENSQL(st, TRUE);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Indicat, 'PHC_', TheRacine); // initialisation des valeurs en fonction des ruptures variables
    end;
  FreeAndNil(TMere);
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction qui initialise par défaut les champs clés de
Suite ........ : l'enregistrement.
Suite ........ : Ne prend pas en compte l'incrément automatique qui sera
Suite ........ : imputé lors de la validation pour tenir compte de l'ensemble
Suite ........ : des indicateurs calculés et qui sert en fait de clé de
Suite ........ : dédoublonnage.
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.RempliDef(Indic: string; T1: TOB);
begin
  T1.PutValue('PBC_INDICATEURBS', Indic);
  T1.PutValue('PBC_NODOSSIER', PgRendNoDossier());
  T1.PutValue('PBC_DATEDEBUT', DateDeb);
  T1.PutValue('PBC_DATEFIN', DateFin);
  T1.PutValue('PBC_BSPRESENTATION', Presentation);
  T1.PutValue('PBC_PGINDICATION', 'CAL');
  T1.PutValue('PBC_BSINDTHEME', Theme);
  T1.PutValue('PBC_CATBILAN', '000');
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction qui l'agrégat SQL en fonction du sexe et de la
Suite ........ : nationalité
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

function TOF_PG_CALINDICBS.RendSexeNat(Sexe, Nat: string; Pref: string = 'PSA'): string;
var
  Scum, SC, StPlus, StWhere, ope: string;
  i: Integer;
begin
  StPlus := '';
  StWhere := '';
  if Pref = '' then Pref := 'PSA';
  if Trim(Sexe) <> '' then StWhere := ' AND ' + Pref + '_SEXE="' + Sexe + '" ';
  if (Trim(Nat) <> '') and (Trim(Nat) <> '<<Tous>>') then
  begin
    if Nat[1] = '!' then
    begin
      ope := '<>';
      Nat := Copy(Nat, 2, Length(nat));
    end
    else ope := '=';
    Sc := Nat;
    i := 0;
    SCum := ReadTokenSt(Sc);
    while (SCum <> '') AND (SCum <> '<<Tous>>') do  { PT6 16/10/06 }
    begin
      if i = 0 then StPlus := StPlus + ' AND ('
      else
        if ope = '=' then StPlus := StPlus + ' OR ' else StPlus := StPlus + ' AND ';
      stplus := stplus + ' ' + Pref + '_NATIONALITE ' + ope + ' "' + Scum + '" ';
      Scum := ReadTokenst(Sc);
      i := i + 1;
    end;
    if (StPlus <> '') then StPlus := Stplus + ') ';
  end;
  result := StPlus + StWhere;
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : fonction de calcul d'un élément concernant une rubrique
Suite ........ : quelque soit sa nature
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndRub(TypeSql : TTypeCalcul; LeChamp, Quoi, Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string);
var
  T1: TOB;
  st, SRem, SC: string;
  Q: TQuery;
  TheRacine: string;
  I: Integer;
   TMere, TFille : Tob;
  DD,DF : TDateTime;
  YY,MM : WORD;

begin
  TheRacine := RendRupture(Indicat, Racine);
  DD := DateDeb;  { DEB PT3 }
  DF := DateFin;
  St := 'SELECT SUM';
  case TypeSql of
     TMoyen : St := 'SELECT AVG';
     TDecembre : Begin DecodeDate(DF,YY,MM,MM);
                       DF := EncodeDate(YY,12,31);
                       DD := DebutDeMois(DF);
                 End;

  End;
  if LeChamp = 'MR' then st := St + '(PHB_MTREM)'
  else  if LeChamp = 'BR'     then St := St + '(PHB_BASEREM)'
  else  if LeChamp = 'BCOT'   then St := St + '(PHB_BASECOT)'
  else  if LeChamp = 'MSCOT'  then St := St + '(PHB_MTSALARIAL)'
  else  if LeChamp = 'MPCOT'  then St := St + '(PHB_MTPATRONAL)'
  else  if LeChamp = 'BBAT'   then St := St + '(PHB_BASECOT)'
  else  if LeChamp = 'BTR1'   then St := St + '(PHB_TRANCHE1)'
  else  if LeChamp = 'BTR2'   then St := St + '(PHB_TRANCHE2)'
  else  if LeChamp = 'BTR3'   then St := St + '(PHB_TRANCHE3)';
			{ FIN PT3 }
  st := st + '  MT,PHB_ETABLISSEMENT,PSA_CATBILAN';
  st := st + TheRacine + ' FROM HISTOBULLETIN LEFT JOIN SALARIES ON PSA_SALARIE=PHB_SALARIE WHERE' +
    ' PHB_DATEFIN <= "' + UsDateTime(DF) + '" AND PHB_DATEDEBUT >= "' + UsDateTime(DD) + '"'; 			{ PT3 }
  Sc := Quoi;
  if Sc[1] <> '!' then
  begin
    i := 0;
    Srem := ReadTokenst(Sc);
    while (Srem <> '') AND (Srem <> '<<Tous>>') do  { PT6 16/10/06 }
    begin
      if i = 0 then st := st + ' AND ( '   { PT6 }
      else st := st + ' OR ';
      st := st + ' PHB_RUBRIQUE="' + Srem + '"';
      Srem := ReadTokenst(Sc);
      i := i + 1;
      If Srem = '' then Srem := Srem + ' ) ';  { PT6 }
    end;
  end
  else st := st + ' AND PHB_RUBRIQUE <> "' + Sc + '"';

  if (LeChamp = 'MR') or (LeChamp = 'BR') then
    st := st + ' AND PHB_NATURERUB="AAA"'
  else
    if (LeChamp = 'BCOT') or (LeChamp = 'MSCOT') or (LeChamp = 'MPCOT') then
    st := st + ' AND PHB_NATURERUB="COT"'
  else
    st := st + ' AND PHB_NATURERUB="BAS"';

  st := st + RendSexeNat(Sexe, Nat);
  st := st + ' GROUP BY PHB_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;
  Q := OPENSQL(st, TRUE);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Indicat, 'PHB_', TheRacine); // initialisation des valeurs en fonction des ruptures variables
    end;
    FreeAndNil(TMere);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction qui rend la clause SQL en fonction des ruptures
Suite ........ : saisies au niveau de l'indicateur en traitement
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

function TOF_PG_CALINDICBS.RendRupture(Indicat: TOB; Racine: string): string;
var
  i, j: Integer;
begin
  if Indicat = nil then exit;
  for i := 1 to 12 do TablS[i] := '';
  for i := 1 to 12 do
  begin
    j := Indicat.GetValue('PBI_B' + HistS[i]);
    if j > 0 then
    begin // Cas de la table PAIEENCOURS, Ts les champs sont connus à l'exception de la catégorie bilan social
      if (i > 3) or (Racine = 'PPU_') then TablS[i] := ',' + Racine + HistS[i]
      else TablS[i] := ',' + 'PSA_' + HistS[i];
    end;
  end;
  for i := 1 to 12 do
    if TablS[i] <> '' then result := result + TablS[i];
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction de calcul du nombre de jours d'absences
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndAbs(Quoi, LesVal, Indic, Sexe, Nat: string; TypeSql : TTypeCalcul; TOBCalcul, Ti: TOB; Racine: string; TypeSelect : TTypeSelect);
var
  T1: TOB;
  st, St2, Motif, SC, StTypeAbs : string;
  Q: TQuery;
  TheRacine: string;
  i: Integer;
  LeChamp: string;
    TMere, TFille : Tob;
  begin
  if Quoi = 'J' then LeChamp := 'PCN_JOURS'
  else if Quoi = 'H' then LeChamp := 'PCN_HEURES'
  else if Quoi = 'S' then LeChamp := 'DISTINCT PCN_SALARIE'  { PT1 }
  else LeChamp := 'PCN_TYPECONGE';                           { PT1 }
  St2 := '';
  TheRacine := RendRupture(Ti, Racine);
  St := 'SUM';    { PT1 }
  Sc := LesVal;
  { DEB PT1 }
  case TypeSql of
     TCompte : St := 'COUNT';
     TSomme  : St := 'SUM';
  End;

  Case TypeSelect Of
       TMotif : Begin     { Dans le cas où le critère est le motif d'absence }
                st := 'SELECT '+St+'(' + LeChamp + ') MT,PCN_ETABLISSEMENT,PSA_CATBILAN';
                st := st + TheRacine +
                ' FROM ABSENCESALARIE LEFT JOIN SALARIES ON PSA_SALARIE=PCN_SALARIE '+
                'WHERE ((PCN_DATEDEBUTABS >="' + USDateTime(DateDeb) + '" ) '+
                'AND (PCN_DATEFINABS <="' + USDateTime(DateFin) + '")) AND PCN_ETATPOSTPAIE <> "NAN"'; { PT5 }
                if (Pos('Tous', Sc) > 0) Or (Trim(Sc) = '') then   //Ds le cas où tous est séléctionné
                st := st + ' AND (PCN_TYPEMVT = "ABS"  OR (PCN_TYPECONGE="PRI" AND PCN_MVTDUPLIQUE<>"X")) AND PCN_SENSABS="-" '
                else                      //Ds le cas où seulement PRI est sélectionné
                  if (POS('PRI', Sc) > 0) and (Length(Sc) < 5) then
                  st := st + ' AND PCN_TYPECONGE="PRI" AND PCN_MVTDUPLIQUE<>"X" AND PCN_SENSABS="-" '
                  else
                    if (POS('PRI', Sc) > 0) and (Length(Sc)>3) then
                       begin      //Ds le cas où PRI et ABS selectionné
                       Motif := ReadTokenst(Sc);
                       while (Motif <> '') AND (Motif <> '<<Tous>>') do  { PT6 16/10/06 }
                         Begin
                         if motif = 'PRI' then begin Motif := ReadTokenst(Sc); continue; end;
                         St2 := St2 + ' PCN_TYPECONGE="'+Motif+'" OR ';
                         Motif := ReadTokenst(Sc);
                         end;
                       If St2 <> '' then St := St + ' AND PCN_SENSABS="-" AND ( (PCN_TYPEMVT = "ABS" AND ('+Copy(St2,1,length(St2)-4)+')) '+
                             'OR (PCN_TYPECONGE="PRI" AND PCN_MVTDUPLIQUE<>"X")) '
                       else
                       st := st + ' AND PCN_TYPECONGE="PRI" AND PCN_MVTDUPLIQUE<>"X" AND PCN_SENSABS="-" '
                       end
                    else
                       if (POS('PRI', Sc) = 0) and (Sc<>'') then
                       begin      //Ds le cas où ABS selectionné
                       Motif := ReadTokenst(Sc);
                       while (Motif <> '') AND (Motif <> '<<Tous>>') do  { PT6 16/10/06 }
                         Begin
                         St2 := St2 + ' PCN_TYPECONGE="'+Motif+'" OR ';
                         Motif := ReadTokenst(Sc);
                         end;
                       If St2 <> '' then St := St + ' AND PCN_SENSABS="-" AND PCN_TYPEMVT = "ABS" AND ('+Copy(St2,1,length(St2)-4)+') '
                       end;
                End;
       TTypeAbs : Begin     { Dans le cas où le critère est le type d'absence }
                  st := 'SELECT '+St+'(' + LeChamp + ') MT,PCN_ETABLISSEMENT,PSA_CATBILAN';
                  st := st + TheRacine +
                  ' FROM ABSENCESALARIE '+
                  'LEFT JOIN SALARIES ON PSA_SALARIE=PCN_SALARIE '+
                  'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PCN_TYPECONGE=PMA_MOTIFABSENCE '+
                  'WHERE (PCN_TYPEMVT = "ABS" OR (PCN_TYPECONGE="PRI" AND PCN_MVTDUPLIQUE<>"X")) AND PCN_SENSABS="-" '+
                  'AND PCN_DATEDEBUTABS >="' + USDateTime(DateDeb) + '"  '+
                  'AND PCN_DATEFINABS <="' + USDateTime(DateFin) + '" AND PCN_ETATPOSTPAIE <> "NAN"'; { PT5 }
                  StTypeAbs := ReadTokenst(Sc);
                  while (StTypeAbs <> '')  AND (Pos('Tous',StTypeAbs) = 0) do
                    Begin
                    St2 := St2 + ' PMA_TYPEABS="'+StTypeAbs+'" OR ';
                    StTypeAbs := ReadTokenst(Sc);
                    end;
                  If St2 <> '' then St := St + ' AND ('+Copy(St2,1,length(St2)-4)+') ';
                  End;
  end;
  { FIN PT1 }
  st := st + RendSexeNat(Sexe, Nat);
  st := st + ' GROUP BY PCN_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;
  Q := OPENSQL(st, TRUE);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Ti, 'PCN_', TheRacine); // initialisation des valeurs en fonction des ruptures variables
    end;
      FreeAndNil(TMere); 
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction qui calcule l'effectif
Suite ........ : total à une date exemple 31/12 donc en fin de période
Suite ........ :
Suite ........ : On prend en compte le nombre de paies au 31/12
Suite ........ : et les salariés qui sont en suspension de paie mais qui ne
Suite ........ : sont pas en sortis. Cela représente les salariés en congés
Suite ........ : sabatiques, arrêt longue maladie ...  mais qui n'ont pas de
Suite ........ : paie dans la période considérée
Suite ........ :
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndEffFinPer(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB);
var
  T1: TOB;
  st : string;
  Q: TQuery;
  TheRacine: string;
   TMere, TFille : Tob;
   i : integer;
begin
  TheRacine := RendRupture(Indicat, 'PSA_');
  St := RendListeSal(TheRacine,Sexe,Nat,'',True,False,False);  { PT7 }
  st := st + ' GROUP BY PSA_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;
  Q := OPENSQL(st, TRUE);
  TheRacine := RendRupture(Indicat, 'PSA_'); { PT7 }
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Indicat, 'PSA_', TheRacine); // initialisation des valeurs en fonction des ruptures variables  { PT7 }
    end;
      FreeAndNil(TMere);

  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2004
Modifié le ... : 07/07/2004
Description .. : Fonction qui calcule l'effectif moyen
Suite ........ : C'est l'effectif mensuel totalisé et divisé par le nombre
Suite ........ : de mois de la période
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndEffMoyen(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB);
var
  TOBMere, T1, T2: TOB;
  st: string;
  Q: TQuery;
  TheRacine: string;
  i: Integer;
  Nbre: Double;
  Rupt: TChampRupt;
  NombreMois: Word;
  OkOk: Boolean;
begin
  TheRacine := RendRupture(Indicat, 'PPU_');
  st := 'SELECT Count (DISTINCT (PPU_SALARIE)) MT,MONTH(PPU_DATEFIN) MOIS, PPU_ETABLISSEMENT,PSA_CATBILAN';// PT4
  st := st + TheRacine +
    ' FROM PAIEENCOURS LEFT JOIN SALARIES ON PSA_SALARIE=PPU_SALARIE WHERE' +
    ' PPU_DATEFIN <= "' + UsDateTime(DateFin) + '" AND PPU_DATEFIN >= "' + UsDateTime(DateDeb) + '"';
  st := st + RendSexeNat(Sexe, Nat);
  st := st + ' GROUP BY PPU_ETABLISSEMENT,PSA_CATBILAN' + TheRacine + ',MONTH(PPU_DATEFIN)'; // PT4
  TOBMere := TOB.Create('MA TOB', nil, -1);
  Q := OPENSQL(st, TRUE);
  TOBMere.LoadDetailDB('MES_PAIEENCOURS', '', '', Q, FALSE);
  Ferme(Q);
  Nbre := 0;

  NombreMois := RendNbPeriode(DateDeb, DateFin);

  for i := 0 to TOBMere.detail.count - 1 do
  begin
    T1 := TOBMere.Detail[i];
    Okok := TestRupture('PPU_', T1, Rupt);
    if (Nbre <> 0) and (OkOk) then // rupture donc Alimentation de la TOBCalcul
    begin
      T2 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      RempliComplAvecRupt(Indic, GblRupt, T2, Nbre, NombreMois, 0); // Arrondi à l'entier
      Nbre := 0;
    end;
    Nbre := Nbre + T1.GetValue('MT'); // ON cumule à rupture identique
  If (i > 0) and (i = TOBMere.detail.count -1 ) then   { PT6 }
      Begin
      T2 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      RempliComplAvecRupt(Indic, Rupt, T2, Nbre, NombreMois, 0); // Arrondi à l'entier
      End;
  end;
  FreeAndNil(TOBMere);
end;



procedure TOF_PG_CALINDICBS.CalculIndNbMoyenDate(TypeTrait: TTypeIndDate;
  TypeCal, Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
var
  Q: TQuery;
  St, StRupt : string;
  Tob_IndMoyen, Tob_Rupt, TM, TR: Tob;
  i, j, Tot: integer;
  D1: TDatetime;
  PMois, PAnnee, Nb: Word;
  Fmini, Fmaxi: Double;
  NomChamp: array of string;
  ValChamp: array of Variant;
  {$IFDEF DEBUGBILAN}
  TTemp : Tob;
  {$ENDIF}
begin
  PMois := 0;
  PAnnee := 0;
  Tot := 0;
  StRupt := RendRupture(Ti, 'PSA_');
  FMini := Ti.getValue('PBI_FOURCHETTEMINI');
  FMaxi := Ti.getValue('PBI_FOURCHETTEMAXI');
  St := RendListeSal(StRupt,Sexe,Nat,',PSA_DATENAISSANCE,PSA_DATEANCIENNETE,PSA_DATEENTREE',False,True,False);  { PT7 }
  St := St + ' ORDER BY PSA_ETABLISSEMENT,PSA_CATBILAN' + StRupt;
  Q := OpenSql(st, True);
  if not Q.eof then
  begin
    Tob_IndMoyen := Tob.Create('Calcul moyen date', nil, -1);
    Tob_IndMoyen.LoadDetailDB('Calcul moyen date', '', '', Q, False);
  end;
  Ferme(Q);            
  St := RendListeSal(StRupt,Sexe,Nat,'',False,False,True); { PT7 }
  St := St +  ' ORDER BY PSA_ETABLISSEMENT,PSA_CATBILAN' + StRupt;
  Q := OpenSql(st, True);
  if not Q.eof then
  begin
    Tob_Rupt := Tob.Create('liste regrpmt', nil, -1);
    Tob_Rupt.LoadDetailDB('liste regrpmt', '', '', Q, False);
  end;
  Ferme(Q);

  { Balayege de la tob pour calcul des ages sur tranches paramétrées }
  if Assigned(Tob_IndMoyen) then
  begin
    Debug('PAIE PGI : Calcul type indicateur : '+TypeCal);
    Debug('PAIE PGI : Tranche indicateur : '+FloatToStr(FMini)+' - '+FloatToStr(FMaxi));
    TM := Tob_IndMoyen.FindFirst([''], [''], False);
    while Assigned(TM) do
    begin
      D1 := idate1900;
      if TypeCal = 'AGE' then D1 := TM.GetValue('PSA_DATENAISSANCE')
      else
        if TypeCal = 'ANC' then
      begin
        D1 := TM.GetValue('PSA_DATEANCIENNETE');
        if (D1 = idate1900) then D1 := TM.GetValue('PSA_DATEENTREE');
      end;
      if (D1 > idate1900) then
      begin
        NOMBREMOIS(D1, DateFin, PMois, PAnnee, Nb);
        Nb := Nb div 12;
        if ((FMini = 0) and (FMaxi = 0))
          or ((FMini <> 0) and (FMaxi = 0) and (Nb >= FMini))
          or ((FMini = 0) and (FMaxi <> 0) and (Nb < FMaxi))
          or ((FMini <> 0) and (FMaxi <> 0) and (Nb >= FMini) and (Nb < FMaxi)) then
            Begin
            TM.AddChampSupValeur('ANNEE', Nb);
            Debug('PAIE PGI : Salarié pris en compte : '+TM.GetValue('PSA_SALARIE'));
            Debug('PAIE PGI : Date pris en compte    : '+DateToStr(D1));
            Debug('PAIE PGI : Nombre d''année         : '+IntToStr(Nb));
            {$IFDEF DEBUGBILAN}
            TTemp := Tob.Create('test_test', Tob_temp , -1);
            TTemp.AddChampSupValeur('SALARIE', TM.GetValue('PSA_SALARIE'));
            {$ENDIF}           
            End
        else
          TM.Free;
      end else
      begin
        ListeError.items.Add(indic + ' : Le date du salarié ' + TM.GetValue('PSA_SALARIE') + ' n''est pas valide.');
        TM.Free;
      end;
      TM := Tob_IndMoyen.FindNext([''], [''], False);
    end;

    { Regroupement par rupture }
    Debug('PAIE PGI : Mise en tob'); 
    if Assigned(Tob_Rupt) then
    begin
      TR := Tob_Rupt.FindFirst([''], [''], True);
      while Assigned(TR) do
      begin
        Tot := TR.ChampsSup.Count;
        SetLength(NomChamp, Tot);
        SetLength(ValChamp, Tot);
        for i := 0 to Tot - 1 do
          NomChamp[i] := TR.GetNomChamp(1000 + i);
        for i := 0 to Tot - 1 do
          ValChamp[i] := TR.GetValue(NomChamp[i]);
        TR.AddChampSupValeur('SOMMEANNEE', 0);
        TR.AddChampSupValeur('NBELEMENT', 0);
        TM := Tob_IndMoyen.FindFirst(NomChamp, ValChamp, True);
        while Assigned(TM) do
        begin
          TR.PutValue('SOMMEANNEE', TR.GetValue('SOMMEANNEE') + TM.GetValue('ANNEE'));
          TR.PutValue('NBELEMENT', TR.GetValue('NBELEMENT') + 1);
          TM.free;
          TM := Tob_IndMoyen.FindNext(NomChamp, ValChamp, True);
        end;
        if TR.GetValue('NBELEMENT') = 0 then TR.free;
        TR := Tob_Rupt.FindNext([''], [''], True);
      end;

      { Remplissage de la tob Bilansocial }
      for i := 0 to Tob_Rupt.detail.count - 1 do
      begin
        TM := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
        RempliDef(indic, TM);
        for j := 0 to Tot - 1 do
        begin
          St := Copy(NomChamp[j], 5, length(NomChamp[j]));
          TM.PutValue('PBC_' + St, Tob_Rupt.detail[i].GetValue(NomChamp[j]));
        end;
        case TypeTrait of
          TNombre:
            begin
              TM.PutValue('PBC_VALCAT', Tob_Rupt.detail[i].GetValue('NBELEMENT'));
              TM.PutValue('PBC_NBELEMENT', 0);
            end;
          TMoyenne:
            begin
              TM.PutValue('PBC_VALCAT', Arrondi(Tob_Rupt.detail[i].GetValue('SOMMEANNEE') / Tob_Rupt.detail[i].GetValue('NBELEMENT'), Ti.GetValue('PBI_ARRONDI')));
              TM.PutValue('PBC_NBELEMENT', Tob_Rupt.detail[i].GetValue('NBELEMENT'));
            end;
        end;
        Debug('PAIE PGI : Indicateurs : '+indic+'  Valeurs     : '+FloatToStr(TM.GetValue('PBC_VALCAT')));
      end;

    end;

  end;
  FreeAndNil(Tob_IndMoyen);
  FreeAndNil(Tob_Rupt);

  {  Moyen := Tob_IndMoyen.Somme('ANNEE',NomChamp,ValChamp,True,True);
    if Moyen <> 0 Then
      Begin
      Moyen := Tob_IndMoyen.Somme('ANNEE',NomChamp,ValChamp,True) / Moyen ;
      if Moyen = 0 then TR.free
      else              TR.AddChampSupValeur('SOMMEANNEE',Moyen);
      End
    Else
      TR.free; }

end;



{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 16/06/2004
Modifié le ... :   /  /
Description .. : Fonction qui rempli les champs ruptures de l'indicateur
Suite ........ : calculé
Mots clefs ... : PÄIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.RempliCompl(Indic: string; TQ, T1, Ti: TOB; Origine, TheRacine: string);
var
  LeChamp: string;
  i: Integer;
begin
  RempliDef(Indic, T1); // initialisation des valeurs par défaut des champs fixes
  if Assigned(TQ) then
  begin
    T1.PutValue('PBC_ETABLISSEMENT', TQ.GetValue(Origine + 'ETABLISSEMENT'));
    If TQ.FieldExists('PSA_CATBILAN') then
      Begin
      if TQ.GetValue('PSA_CATBILAN')='' then TQ.PutValue('PSA_CATBILAN','000');
      End
    else
      TQ.AddChampSupValeur('PSA_CATBILAN','000');
    T1.PutValue('PBC_CATBILAN', TQ.GetValue('PSA_CATBILAN'));
    T1.PutValue('PBC_VALCAT', Arrondi(TQ.GetValue('MT'), Ti.GetValue('PBI_ARRONDI')));
    if TQ.FieldExists('NBELEMENT') then
      T1.PutValue('PBC_NBELEMENT', TQ.GetValue('NBELEMENT'));
    for i := 1 to 12 do
    begin
      if TablS[i] <> '' then
      begin
        if i < 4 then LeChamp := 'PSA_' + HistS[i]
        else LeChamp := Origine + HistS[i];
        T1.PutValue('PBC_' + HistS[i], TQ.GetValue(LeChamp));
      end;
    end;

  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 07/07/2004
Modifié le ... :   /  /
Description .. : Fonction de test s'il y a rupture en fonction des différentes
Suite ........ : ruptures gérées au niveau de l'indicateur.
Suite ........ :
Suite ........ : En fait, on cumule tous les mois pour un critère identique
Mots clefs ... :
*****************************************************************}

function TOF_PG_CALINDICBS.TestRupture(Racine: string; T1: TOB; var Rupt: TChampRupt): boolean;
var
  LeChamp: string;
  i: Integer;
  OkOk: Boolean;
begin
  OkOk := TRUE;
  for i := 1 to 12 do
  begin
    if TablS[i] <> '' then
    begin
      if i < 4 then LeChamp := 'PSA_' + HistS[i]
      else LeChamp := Racine + HistS[i];
      if T1.FieldExists(LeChamp) then
      begin
        OkOk := CompareRuptOk(T1, Rupt, LeChamp, i);
        if not OkOk then Break; // Fonction de comparaison des ruptures
      end;
    end;
  end;

  // Prise en compte aussi Etablissement et catégorie bilan social
  if OkOk then
  begin
    OkOk := CompareRuptOk(T1, Rupt, Racine + 'ETABLISSEMENT', 13);
    if not OkOk then OkOk := CompareRuptOk(T1, Rupt, 'PSA_CATBILAN', 14);
  end;
  if not OkOk then
  begin
    GblRupt := Rupt;
    for i := 1 to 12 do
    begin
      if TablS[i] <> '' then
      begin
        if i < 4 then LeChamp := 'PSA_' + HistS[i]
        else LeChamp := Racine + HistS[i];
        if T1.FieldExists(LeChamp) then
          AffecteChampRupt(T1, Rupt, Lechamp, i, FALSE) // Fonction Affectation par défaut des nouveaux champs ruptures troouvés
        else AffecteChampRupt(T1, Rupt, Lechamp, i, TRUE)
      end;
    end;
  end;
  AffecteChampRupt(T1, Rupt, Racine + 'ETABLISSEMENT', 13, FALSE);
  AffecteChampRupt(T1, Rupt, 'PSA_CATBILAN', 14, FALSE);
  result := OkOK;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Fonction de comparaison champ à champ des ruptures
Suite ........ : possibles
Mots clefs ... :
*****************************************************************}

function TOF_PG_CALINDICBS.CompareRuptOk(T1: TOB; var Rupt: TChampRupt; LeChamp: string; i: integer): Boolean;
var
  St: string;
begin
  Result := False;
  St := T1.GetValue(LeChamp); // Récup de la donnée
  case i of
    1: if st <> Rupt.LIBELLEEMPLOI then result := TRUE;
    2: if st <> Rupt.QUALIFICATION then result := TRUE;
    3: if st <> Rupt.COEFFICIENT then result := TRUE;
    4: if st <> Rupt.CODESTAT then result := TRUE;
    5: if st <> Rupt.TRAVAILN1 then result := TRUE;
    6: if st <> Rupt.TRAVAILN2 then result := TRUE;
    7: if st <> Rupt.TRAVAILN3 then result := TRUE;
    8: if st <> Rupt.TRAVAILN4 then result := TRUE;
    9: if st <> Rupt.LIBREPCMB1 then result := TRUE;
    10: if st <> Rupt.LIBREPCMB2 then result := TRUE;
    11: if st <> Rupt.LIBREPCMB3 then result := TRUE;
    12: if st <> Rupt.LIBREPCMB4 then result := TRUE;
    13: if st <> Rupt.ETABLISSEMENT then result := TRUE;
    14: if st <> Rupt.CATBILAN then result := TRUE;
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Fonction d'affectation champ à champ des ruptures
Suite ........ : possibles
Mots clefs ... :
*****************************************************************}

procedure TOF_PG_CALINDICBS.AffecteChampRupt(T1: TOB; var Rupt: TChampRupt; LeChamp: string; i: integer; Valdef: Boolean);
var
  St: string;
begin
  if not ValDef then St := T1.GetValue(LeChamp)
  else st := '';
  case i of
    1: Rupt.LIBELLEEMPLOI := st;
    2: Rupt.QUALIFICATION := st;
    3: Rupt.COEFFICIENT := st;
    4: Rupt.CODESTAT := st;
    5: Rupt.TRAVAILN1 := st;
    6: Rupt.TRAVAILN2 := st;
    7: Rupt.TRAVAILN3 := st;
    8: Rupt.TRAVAILN4 := st;
    9: Rupt.LIBREPCMB1 := st;
    10: Rupt.LIBREPCMB2 := st;
    11: Rupt.LIBREPCMB3 := st;
    12: Rupt.LIBREPCMB4 := st;
    13: Rupt.ETABLISSEMENT := st;
    14: Rupt.CATBILAN := st;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Fonction remplissage de la TOB dans le cas de traitement
Suite ........ : d'indicateur avec rupture cad  cumul de montants ou effectif
Suite ........ : avant d'écrire.
Suite ........ :
Suite ........ : Cas de calcul d'effectifs mensuels divisés par le nombre de
Suite ........ : mois de la période pour avoir un effetif moyen
Mots clefs ... :
*****************************************************************}

procedure TOF_PG_CALINDICBS.RempliComplAvecRupt(Indic: string; var Rupt: TChampRupt; T1: TOB;
  Nbre: Double; NbMois, NbArrondi: Integer);
begin
  RempliDef(Indic, T1); // initialisation des valeurs par défaut des champs fixes
  T1.PutValue('PBC_CATBILAN', Rupt.CATBILAN);
  T1.PutValue('PBC_LIBELLEEMPLOI', Rupt.LIBELLEEMPLOI);
  T1.PutValue('PBC_QUALIFICATION', Rupt.QUALIFICATION);
  T1.PutValue('PBC_COEFFICIENT', Rupt.COEFFICIENT);
  T1.PutValue('PBC_CODESTAT', Rupt.CODESTAT);
  T1.PutValue('PBC_TRAVAILN1', Rupt.TRAVAILN1);
  T1.PutValue('PBC_TRAVAILN2', Rupt.TRAVAILN2);
  T1.PutValue('PBC_TRAVAILN3', Rupt.TRAVAILN3);
  T1.PutValue('PBC_TRAVAILN4', Rupt.TRAVAILN4);
  T1.PutValue('PBC_LIBREPCMB1', Rupt.LIBREPCMB1);
  T1.PutValue('PBC_LIBREPCMB2', Rupt.LIBREPCMB2);
  T1.PutValue('PBC_LIBREPCMB3', Rupt.LIBREPCMB3);
  T1.PutValue('PBC_LIBREPCMB4', Rupt.LIBREPCMB4);
  T1.PutValue('PBC_ETABLISSEMENT', Rupt.ETABLISSEMENT);
  if NbMois <> 0 then  { PT6 }
  T1.PutValue('PBC_VALCAT', ARRONDI(Nbre / NbMois, NbArrondi)) // Arrondi de l'effectif à l'entier
  else
  T1.PutValue('PBC_VALCAT', 0);  { PT6 }
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 09/07/2004
Modifié le ... : 09/07/2004
Description .. : Fonction qui calcule le nombre de contrat de travail
Suite ........ :
Suite ........ : Utilisée pour calculer le nombre de contrats par type de
Suite ........ : contrat mais aussi pour l'effectif permanent qui represente
Suite ........ : le nombre de salariés en CDI présent sur l'ensemble de la
Suite ........ : période
Mots clefs ... : PAIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndContrat(Quoi, Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine, TypInd: string; Moins25 : Boolean = False);
var
  T1: TOB;
  st, Scontrat, Sc: string;
  Q: TQuery;
  TheRacine: string;
  i: Integer;
  TMere, TFille : Tob;
begin
  TheRacine := RendRupture(Indicat, Racine);
  st := 'SELECT COUNT(*) MT,PCI_ETABLISSEMENT,PSA_CATBILAN';
  st := st + TheRacine +
    ' FROM CONTRATTRAVAIL LEFT JOIN SALARIES ON PSA_SALARIE=PCI_SALARIE WHERE';
  if TypInd = 'EFF' then // Pour le calcul de l'effectif permanent
  begin
    st := st + ' PCI_DEBUTCONTRAT <= "' + UsDateTime(DateDeb) + '" AND (PCI_FINCONTRAT >= "' +
      UsDateTime(DateFin) + '" OR PCI_FINCONTRAT="' + UsDateTime(IDate1900) + '" OR PCI_FINCONTRAT IS NULL)';
  end
  else
  begin
    if TypInd = 'EMB' then // Nombre d'embauche par contrat dans la période
    begin
      st := st + ' PCI_DEBUTCONTRAT >= "' + UsDateTime(DateDeb) + '" AND PCI_DEBUTCONTRAT <= "' +
        UsDateTime(DateFin) + '" ';
    end
    else
    begin // Pour le calcul du nombre de contrats dans la période
      st := st + ' (PCI_DEBUTCONTRAT <= "' + UsDateTime(DateFin) + '" AND (PCI_FINCONTRAT > "' +   { PT6 }
        UsDateTime(DateFin) + '" OR PCI_FINCONTRAT="' + UsDateTime(IDate1900) + '" OR PCI_FINCONTRAT IS NULL))';   { PT6 }
    end;
  end;
  Sc := Quoi;
  if Sc[1] <> '!' then
  begin
    Scontrat := ReadTokenst(Sc);
    i := 0;
    while (Scontrat <> '') AND  (Scontrat <> '<<Tous>>') do  { PT6 16/10/06 }
    begin
      if i = 0 then st := st + ' AND ( '   { PT6 }
      else st := st + ' OR ';
      st := st + ' PCI_TYPECONTRAT="' + Scontrat + '"';
      Scontrat := ReadTokenst(Sc);
      i := i + 1;
      if SContrat = '' then St := St +' ) ';   { PT6 }
    end;
  end
  else st := st + ' AND PCI_TYPECONTRAT <> "' + Sc + '"';
  { DEB PT8 }
  if Moins25 then
    Begin
    St := St + 'AND (( YEAR(PCI_DEBUTCONTRAT)-YEAR(PSA_DATENAISSANCE)<25 ) '+
               'OR ( (YEAR(PCI_DEBUTCONTRAT)-YEAR(PSA_DATENAISSANCE))=25 AND (MONTH(PCI_DEBUTCONTRAT)- MONTH(PSA_DATENAISSANCE))=0 '+
               '      AND (DAY(PCI_DEBUTCONTRAT)-DAY(PSA_DATENAISSANCE))<0  ) '+
               ' OR ( (YEAR(PCI_DEBUTCONTRAT)-YEAR(PSA_DATENAISSANCE))=25 AND (MONTH(PCI_DEBUTCONTRAT)- MONTH(PSA_DATENAISSANCE))<0)) ';
   End;
  { FIN PT8 }
  st := st + RendSexeNat(Sexe, Nat);
  st := st + ' GROUP BY PCI_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;
  Q := OPENSQL(st, TRUE);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Indicat, 'PCI_', TheRacine); // initialisation des valeurs en fonction des ruptures variables
    end;
   FreeAndNil(TMere);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 19/10/2004
Modifié le ... : 19/10/2004
Description .. : Fonction qui calcule le nombre de visites médicales
Suite ........ :
Mots clefs ... : PAIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculIndVisiteMed(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string);
var
  T1: TOB;
  st: string;
  Q: TQuery;
  TheRacine: string;
  i : integer;
  TMere,TFille : TOB;
begin
  TheRacine := RendRupture(Indicat, Racine);
  st := 'SELECT Count (DISTINCT (PVM_SALARIE)) MT,PSA_ETABLISSEMENT,PSA_CATBILAN';
  st := st + TheRacine +
    ' FROM VISITEMEDTRAV LEFT JOIN SALARIES ON PSA_SALARIE=PVM_SALARIE WHERE' +
    ' PVM_DATEVISITE <= "' + UsDateTime(DateFin) + '" AND PVM_DATEVISITE >= "' + UsDateTime(DateDeb) + '"';
  st := st + RendSexeNat(Sexe, Nat);
  st := st + ' GROUP BY PSA_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;
  Q := OPENSQL(st, TRUE);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Indicat, 'PSA_', TheRacine); // initialisation des valeurs en fonction des ruptures variables
    end;
      FreeAndNil(TMere);

  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 19/10/2004
Modifié le ... : 19/10/2004
Description .. : Fonction qui calcule le nombre de salariés handicapés
Suite ........ : On regarde les salariés handicapés et s'ils ont eu des paies
Suite ........ : dans la période traitée soit au 31/03 de l'année concernée
Mots clefs ... : PAIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculNbreHandicape(Indic, Sexe, Nat: string; TOBCalcul, Indicat: TOB; Racine: string);
var
  T1: TOB;
  st: string;
  Q: TQuery;
  TheRacine: string;
  LaDate: TDateTime;
  Jour, Mois, Annee: WORD;
  TMere, TFille : Tob;
  I : integer;
begin
  TheRacine := RendRupture(Indicat, Racine);
  DecodeDate(DateFin, Annee, Mois, Jour);
  LaDate := EncodeDate(Annee, 3, 31);
  st := 'SELECT COUNT (DISTINCT (PGH_SALARIE)) MT,PSA_ETABLISSEMENT,PSA_CATBILAN ';
  st := st + TheRacine +
    ' FROM HANDICAPE LEFT JOIN SALARIES ON PSA_SALARIE=PGH_SALARIE WHERE ' +
    'EXISTS (SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_SALARIE=PGH_SALARIE AND PPU_DATEDEBUT >="' + UsDateTime(DateDeb) +
    '" AND PPU_DATEFIN<="' + UsDateTime(LaDate) + '")';
  st := st + RendSexeNat(Sexe, Nat);
  st := st + ' GROUP BY PSA_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;
  Q := OPENSQL(st, TRUE);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Indicat, 'PSA_', TheRacine); // initialisation des valeurs en fonction des ruptures variables
    end;
    FreeAndNil(TMere);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 21/10/2004
Modifié le ... : 21/10/2004
Description .. : Fonction qui va rechercher un montant  débiteur ou
Suite ........ : créditeur dans la comptabilité PGI
Suite ........ : Correspond au chapitre VII du bilan social
Suite ........ : Autres conditions de vie relevant de l'entreprise
Suite ........ : Il n'ya que l'établissement comme rupture possible
Mots clefs ... : PAIEBILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.CalculMTCompta(Quoi, Indic: string; TOBCalcul, Indicat: TOB; TypeMT: string);
var
  st, sc, Scompte: string;
  Q: TQuery;
  T1: TOB;
  i: Integer;
begin
  if TypeMT = '' then exit;
  if TypeMT = 'CRE' then st := 'SELECT SUM(E_CREDIT) MT'
  else if TypeMT = 'DEB' then st := 'SELECT SUM(E_DEDIT) MT';
  st := st + ',E_ETABLISSEMENT FROM ECRITURE WHERE E_DATECOMPTABLE >="' + UsDateTime(DateDeb) + '" AND E_DATECOMPTABLE <="' + UsDateTime(DateFin) + '"';
  Sc := Quoi;
  if Sc[1] <> '!' then
  begin
    i := 0;
    Scompte := ReadTokenst(Sc);
    while (SCompte <> '') and  (SCompte <> '<<Tous>>') do  { PT6 16/10/06 }
    begin
      if i = 0 then st := st + ' AND ( '  { PT6 }
      else st := st + ' OR ';
      st := st + ' E_GENERAL ="' + Scompte + '"';
      Scompte := ReadTokenst(Sc);
      i := i + 1;
      IF SCompte = '' then St := St + ' ) ';   { PT6 }
    end;
  end
  else st := st + ' AND E_GENERAL <> "' + Sc + '"';
  st := st + ' GROUP BY E_ETABLISSEMENT';
  Q := OPENSQL(st, TRUE);
  while not Q.EOF do
  begin
    T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
    RempliIndCompta(Indic, Q, T1, 'E'); // initialisation des valeurs en fonction des ruptures variables
    Q.Next;
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : Fonction de remplissage des enregistrements du bilan social
Suite ........ : de type éléments extérieurs à la paie et aux salariés.
Suite ........ : On ne prend en compte que l'établissement par exemple
Suite ........ : dans le cas de la compta
Mots clefs ... : PAIE;BILAN
*****************************************************************}

procedure TOF_PG_CALINDICBS.RempliIndCompta(Indic: string; Q: TQuery; T1: TOB; Origine: string);
var
  i: Integer;
begin
  RempliDef(Indic, T1); // initialisation des valeurs par défaut des champs fixes
  T1.PutValue('PBC_ETABLISSEMENT', Q.findField(Origine + 'ETABLISSEMENT').AsString);
  T1.PutValue('PBC_CATBILAN', '');
  T1.PutValue('PBC_VALCAT', Q.findField('MT').AsFloat);
  for i := 1 to 12 do T1.PutValue('PBC_' + HistS[i], '');
end;




procedure TOF_PG_CALINDICBS.CalculDeparts(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
var
  St, StWhere, ListeTypeDep, TheRacine: string;
  Q: TQuery;
  T1: Tob;
    i: Integer;
  TMere, TFille : Tob;
begin
  {11: ; { Lic. Redress./Liquid. judiciaire }
  {12: ; { Lic. Fermeture définitive }
  {14: ; { Lic. Motif économique }
  {15: ; { Lic. Fin de chantier }
  {20: ; { Autre motif de licenciement }
  {25: ; { Autre rupture motif économique }
  {31: ; { Fin de CDD }
  {32: ; { Fin de mission d'intérim }
  {34: ; { Fin essai employeur }
  {35: ; { Fin essai Salarié }
  {36: ; { Rupture CDD employeur }
  {37: ; { Rupture CDD salarié }
  {38: ; { Départ retraite employeur }
  {39: ; { Départ retraite salarié }
  {59: ; { Démission }
  {60: ; { Autre motif }
  {81: ; { Fin de contrat d'apprentissage }
  {82: ; { Résiliation judiciaire }
  {83: ; { Force majeure/Fait du prince }
  {84: ; { Commun accord CDD ou apprentissage }
  StWhere := '';
  ListeTypeDep := Ti.GetValue('PBI_LESVALEURS');
  while (ListeTypeDep <> '') AND (ListeTypeDep <> '<<Tous>>') do  { PT6 16/10/06 }
  begin
    St := ReadTokenst(ListeTypeDep);
    if IsNumeric(St) then StWhere := StWhere + 'PAS_MOTRUPCONT="' + St + '" OR ';
  end;
  if StWhere <> '' then StWhere := 'AND (' + Copy(StWhere, 1, Length(StWhere) - 4) + ')';

  StWhere := StWhere + RendSexeNat(Sexe, Nat);
  TheRacine := RendRupture(Ti, 'PSA_');

  St := 'SELECT COUNT (*) MT, PSA_ETABLISSEMENT,PSA_CATBILAN' + TheRacine +
    ' FROM ATTESTATIONS,SALARIES ' +
    'WHERE PAS_SALARIE=PSA_SALARIE AND PAS_TYPEATTEST="ASS" ' +
    'AND PAS_DERNIERJOUR>="' + USDateTime(DateDeb) + '" ' +
    'AND PAS_DERNIERJOUR<="' + USDateTime(DateFin) + '" ' + StWhere +
    'GROUP BY PSA_ETABLISSEMENT,PSA_CATBILAN' + TheRacine;

  Q := OpenSql(St, True);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Ti, 'PSA_', TheRacine); // initialisation des valeurs en fonction des ruptures variables
    end;
    FreeAndNil(TMere);
  end;
end;

procedure TOF_PG_CALINDICBS.CalculNbMutEtab(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
var
  St, StAutre, StRupt, Etab: string;
  Q: TQuery;
  Tob_Etab, Tob_liste, TM, T1: Tob;
  i, j: Integer;
begin
  { Liste des établissements }
  Tob_Etab := Tob.Create('ETABCOMPL', nil, -1);
  Tob_Etab.LoadDetailDB('ETABCOMPL', '', '', nil, False);
  StAutre := RendSexeNat(sexe, nat);

  for i := 0 to Tob_Etab.detail.count - 1 do
  begin
    FreeAndNil(Tob_liste);
    Etab := Tob_Etab.detail[i].GetValue('ETB_ETABLISSEMENT');
    StRupt := RendRupture(Ti, 'PSA_');
    { Chargement de la liste des mutations }
    St := 'SELECT DISTINCT PPU_ETABLISSEMENT,COUNT(DISTINCT PPU_SALARIE) NB,PSA_CATBILAN' + StRupt +
      ' FROM PAIEENCOURS,SALARIES ' +
      ' WHERE PSA_SALARIE = PPU_SALARIE ' +
      'AND PPU_ETABLISSEMENT = "' + Etab + '" ' + StAutre +
      'AND PPU_SALARIE IN (SELECT PP.PSA_SALARIE FROM SALARIES PP ' +
      'WHERE PP.PSA_SALARIE = PPU_SALARIE AND PP.PSA_ETABLISSEMENT<>"' + Etab + '") ' +
      'AND PPU_DATEDEBUT>="' + USDateTime(DateDeb) + '" ' +
      'AND PPU_DATEFIN<="' + USDateTime(DateFin) + '" ' +
      'GROUP BY PPU_ETABLISSEMENT,PSA_CATBILAN' + StRupt;
    Q := OpenSql(st, true);
    if not Q.eof then
    begin
      Tob_liste := TOB.Create('Liste des établissements', nil, -1);
      Tob_liste.LoadDetailDB('Liste des établissements', '', '', Q, False);
    end;
    Ferme(Q);

    if not assigned(Tob_liste) then continue;
    { Balayage de la tob pour retrouver les nb salarié > 1 }
    for j := 0 to Tob_liste.Detail.count - 1 do
    begin
      TM := Tob_liste.Detail[j];
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      RempliDef(Indic, T1); // initialisation des valeurs par défaut des champs fixes
      T1.PutValue('PBC_ETABLISSEMENT', Etab);
      T1.PutValue('PBC_CATBILAN', TM.GetValue('PSA_CATBILAN'));
      T1.PutValue('PBC_VALCAT', TM.GetValue('NB'));
      ReadTokenPipe(StRupt, ','); //pour supprimer la 1ère virgule
      St := ReadTokenPipe(StRupt, ',');
      while St <> '' do
      begin
        T1.PutValue('PBC_' + Copy(st, 5, length(st)), TM.GetValue(St));
        St := ReadTokenPipe(StRupt, ',');
      end;
    end;
  end;
  FreeAndNil(Tob_liste);
  FreeAndNil(Tob_Etab);
end;


procedure TOF_PG_CALINDICBS.Calcul10RemPlusMoinsElev(TypeTrait: TTypeIndRem;
  Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB; RuptEtab: Boolean);
var
  St, StAutre, StWhere, StCumul, Etab: string;
  Tob_Etab, TEtab, T1: Tob;
  Q: TQuery;
begin
  FreeAndNil(Tob_Etab);
  St := Ti.GetValue('PBI_LESVALEURS');
  StCumul := ReadTokenSt(St); //Ne tient compte que du premier cumul sélectionné
  StAutre := RendSexeNat(sexe, nat);
  if StAutre <> '' then StAutre := ', SALARIES WHERE PHC_SALARIE = PSA_SALARIE ' + StAutre +
    ' AND PHC_CUMULPAIE = "' + StCumul + '" '
  else StAutre := ' WHERE PHC_CUMULPAIE = "' + StCumul + '" ';
  StAutre := StAutre + 'AND PHC_DATEDEBUT>="' + USDateTime(DateDeb) + '" ' +
    'AND PHC_DATEFIN<="' + USDateTime(DateFin) + '" ';
  if StCumul = '' then exit;

  if RuptEtab then
  begin
    Tob_Etab := TOB.Create('Les établissements', nil, -1);
    Tob_Etab.LoadDetailDB('ETABCOMPL', '', '', nil, False);
    if Assigned(Tob_Etab) then TEtab := Tob_Etab.FindFirst([''], [''], False) else FreeAndNil(TEtab);
    if Assigned(TEtab) then Etab := TEtab.GetValue('ETB_ETABLISSEMENT');
  end
  else
    Etab := '...'; //Pour entrer boucle et traiter le cas de non rupture par étab.

  { Balayage des établissements }
  while Etab <> '' do
  begin
    if Etab <> '...' then StWhere := ' AND PHC_ETABLISSEMENT = "' + Etab + '" ' else StWhere := '';
    St := 'SELECT ##TOP 10## PHC_MONTANT FROM HISTOCUMSAL ' + StAutre + StWhere +
      'ORDER BY PHC_MONTANT';

    case TypeTrait of
      TPlus: St := St + ' DESC';
    end;

    Q := OpenSql(St, True);
    while not Q.EOF do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      RempliDef(Indic, T1);
      if Etab <> '...' then T1.PutValue('PBC_ETABLISSEMENT', Etab);
      //    T1.PutValue('PBC_CATBILAN', '***');
      T1.PutValue('PBC_VALCAT', Arrondi(Q.findField('PHC_MONTANT').AsFloat, Ti.GetValue('PBI_ARRONDI')));
      Q.Next;
    end;
    Ferme(Q);
    if Assigned(Tob_Etab) then TEtab := Tob_Etab.FindNext([''], [''], False);
    if Assigned(TEtab) then Etab := TEtab.GetValue('ETB_ETABLISSEMENT')
    else Etab := '';
  end;


end;

procedure TOF_PG_CALINDICBS.CalculTravExt(TypeCalTrav: TTypeIndTrav; Indic, Sexe, Nat: string;
  TOBCalcul, Ti: TOB);
begin
  case TypeCalTrav of
    TNbInt: CalculNbInt(Indic, Sexe, Nat, TOBCalcul, Ti); //45
    TNbMoyenInt: CalculMoyenInt(TNbMoyenInt, Indic, Sexe, Nat, TOBCalcul, Ti); //46
    TNbMoyenContInt: CalculMoyenInt(TNbMoyenContInt, Indic, Sexe, Nat, TOBCalcul, Ti); //47
    TNbStag: CalculNbStag(Indic, Sexe, Nat, TOBCalcul, Ti); //48
    TDureeContInt: CalculDureeContInt(Indic, Sexe, Nat, TOBCalcul, Ti); //49
  end;
end;

procedure TOF_PG_CALINDICBS.CalculDureeContInt(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
var
  Sql, Etab, StWhere: string;
  Q: TQuery;
  TCont, T, T1: TOB;
  i, Cpt: integer;
  DE, FE: TDateTime;
  NbJour: Double;
begin
  StWhere := RendSexeNat(Sexe, Nat, 'PSI');
  SQL := 'SELECT PEI_DEBUTEMPLOI,PEI_FINEMPLOI,PEI_ETABLISSEMENT FROM EMPLOIINTERIM ' +
    'LEFT JOIN INTERIMAIRES ON PEI_INTERIMAIRE=PSI_INTERIMAIRE ' +
    'WHERE PSI_TYPEINTERIM="INT"  ' + StWhere +
    'AND ((PEI_DEBUTEMPLOI>="' + USDateTime(DateDeb) + '" ' +
    'AND PEI_DEBUTEMPLOI<="' + USDateTime(DateFin) + '") ' +
    'OR (PEI_FINEMPLOI>="' + USDateTime(DateDeb) + '"  ' +
    'AND PEI_FINEMPLOI<="' + USDateTime(DateFin) + '")) ORDER BY PEI_ETABLISSEMENT';

  Q := OpenSql(Sql, True);
  if not Q.EOF then
  begin
    TCont := Tob.Create('Liste des contrat', nil, -1);
    TCont.LoadDetailDB('Liste des contrat', '', '', Q, False);
  end
  else
    FreeAndNil(TCont);
  Ferme(Q);

  Cpt := 0;
  NbJour := 0;
  Etab := '';

  if Assigned(TCont) then
  begin
    for i := 0 to TCont.Detail.Count - 1 do
    begin
      T := TCont.Detail[i];
      if (Etab <> '') and (Etab <> T.GetValue('PEI_ETABLISSEMENT')) then
      begin { Changement d'établissement }
        T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
        RempliDef(Indic, T1);
        T1.PutValue('PBC_ETABLISSEMENT', Etab);
        T1.PutValue('PBC_VALCAT', Arrondi(NbJour / Cpt, Ti.GetValue('PBI_ARRONDI')));
        T1.PutValue('PBC_NBELEMENT', Cpt);
        Cpt := 0;
        NbJour := 0;
      end;
      Etab := T.GetValue('PEI_ETABLISSEMENT');
      DE := T.GetValue('PEI_DEBUTEMPLOI');
      FE := T.GetValue('PEI_FINEMPLOI');
      if (DE > Idate1900) and (FE > Idate1900) then
      begin
        if DE < DateDeb then DE := DateDeb;
        if FE > DateFin then FE := DateFin;
        NbJour := NbJour + (FE - DE);
        Inc(Cpt);
      end;
    end;
    T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
    RempliDef(Indic, T1);
    T1.PutValue('PBC_ETABLISSEMENT', Etab);
    T1.PutValue('PBC_VALCAT', Arrondi(NbJour / Cpt, Ti.GetValue('PBI_ARRONDI')));
    T1.PutValue('PBC_NBELEMENT', Cpt);
  end;
  FreeAndNil(TCont);
end;




{procedure TOF_PG_CALINDICBS.CalculNbMoyenContInt(Indic, Sexe, Nat: string;
  TOBCalcul, Ti: TOB);
Var
   StWhere, Sql : String;
   Q : TQuery;
   T1 : Tob;
begin
  StWhere := RendSexeNat(Sexe,Nat,'PSI');
  SQL := 'SELECT COUNT (PEI_INTERIMAIRE) MT ,PEI_ETABLISSEMENT '+
       'FROM INTERIMAIRES '+
       'LEFT JOIN EMPLOIINTERIM ON PSI_INTERIMAIRE=PEI_INTERIMAIRE '+
       'WHERE PSI_TYPEINTERIM="INT"  '+StWhere+
       'AND ((PEI_DEBUTEMPLOI>="'+USDateTime(DateDeb)+'" '+
       'AND PEI_DEBUTEMPLOI<="'+USDateTime(DateFin)+'") '+
       'OR (PEI_FINEMPLOI>="'+USDateTime(DateDeb)+'"  '+
       'AND PEI_FINEMPLOI<="'+USDateTime(DateFin)+'")) GROUP BY PEI_ETABLISSEMENT';

  Q := OpenSql(SQL,True);
  while not Q.EOF do
  begin
    T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
    RempliCompl(Indic, Q, nil, T1, Ti, 'PEI_', '');
    Q.Next;
  end;
  Ferme(Q);
end;  }

procedure TOF_PG_CALINDICBS.CalculNbInt(Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
var
  Sql, StWhere: string;
  Q: TQuery;
  TMere, TFille,T1: Tob;
  i: Integer;
begin
  StWhere := RendSexeNat(Sexe, Nat, 'PSI');

  SQL := 'SELECT COUNT (DISTINCT(PEI_INTERIMAIRE)) MT ,PEI_ETABLISSEMENT ' +
    'FROM INTERIMAIRES ' +
    'LEFT JOIN EMPLOIINTERIM ON PSI_INTERIMAIRE=PEI_INTERIMAIRE ' +
    'WHERE PSI_TYPEINTERIM="INT" ' + StWhere +
    'AND ((PEI_DEBUTEMPLOI>="' + USDateTime(DateDeb) + '" ' +
    'AND PEI_DEBUTEMPLOI<="' + USDateTime(DateFin) + '") ' +
    'OR (PEI_FINEMPLOI>="' + USDateTime(DateDeb) + '" ' +
    'AND PEI_FINEMPLOI<="' + USDateTime(DateFin) + '")) ' +
    'GROUP BY PEI_ETABLISSEMENT';

  Q := OpenSql(SQL, True);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Ti, 'PEI_', ''); // initialisation des valeurs en fonction des ruptures variables
    end;
      FreeAndNil(TMere);

  end;
end;

procedure TOF_PG_CALINDICBS.CalculMoyenInt(TypeCalTrav: TTypeIndTrav; Indic, Sexe, Nat: string; TOBCalcul, Ti: TOB);
var
  Sql, StWhere, Etab: string;
  Q: TQuery;
  T1, TMere, TMereTQ, TQ: Tob;
  i: Integer;
  Mt, Cpt: Double;
begin
  StWhere := RendSexeNat(Sexe, Nat, 'PSI');
  if TypeCalTrav = TNbMoyenInt then SQL := 'DISTINCT'
  else if TypeCalTrav = TNbMoyenContInt then SQL := '';

  SQL := 'SELECT COUNT (' + SQL + '(PEI_INTERIMAIRE)) MT,PEI_ETABLISSEMENT,' +
    'MONTH(PEI_DEBUTEMPLOI),MONTH(PEI_FINEMPLOI) ' +
    'FROM INTERIMAIRES ' +
    'LEFT JOIN EMPLOIINTERIM ON PSI_INTERIMAIRE=PEI_INTERIMAIRE ' +
    'WHERE PSI_TYPEINTERIM="INT" ' + StWhere +
    'AND ((PEI_DEBUTEMPLOI>="' + USDateTime(DateDeb) + '" ' +
    'AND PEI_DEBUTEMPLOI<="' + USDateTime(DateFin) + '") ' +
    'OR (PEI_FINEMPLOI>="' + USDateTime(DateDeb) + '" ' +
    'AND PEI_FINEMPLOI<="' + USDateTime(DateFin) + '")) ' +
    'GROUP BY PEI_ETABLISSEMENT,MONTH(PEI_DEBUTEMPLOI),MONTH(PEI_FINEMPLOI)';

  FreeAndNil(TMereTQ);
  Q := OpenSql(SQL, True);
  if not Q.eof then
  begin
    TMere := TOB.Create('BILAN_SOCIAL', nil, -1);
    TMere.LoadDetailDB('BILAN_SOCIAL', '', '', Q, False);
    Etab := '';
    for i := 0 to TMere.Detail.count - 1 do
    begin
      if Etab = TMere.Detail[i].GetValue('PEI_ETABLISSEMENT') then continue;
      Etab := TMere.Detail[i].GetValue('PEI_ETABLISSEMENT');
      Mt := TMere.somme('MT', ['PEI_ETABLISSEMENT'], [Etab], False);
      Cpt := TMere.somme('MT', ['PEI_ETABLISSEMENT'], [Etab], False, True);
      TMereTQ := TOB.Create('BILAN_SOC', nil, -1);
      TQ := TOB.Create('BILAN_SOC', TMereTQ, -1);
      TQ.AddChampSupValeur('PEI_ETABLISSEMENT', Etab);
      TQ.AddChampSupValeur('MT', Mt / Cpt);
      TQ.AddChampSupValeur('NBELEMENT', Cpt);
    end;
  end;
  Ferme(Q);
  if Assigned(TMereTQ) then
    for i := 0 to TMereTQ.Detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TQ := TMereTQ.Detail[i];
      RempliCompl(Indic, TQ, T1, Ti, 'PEI_', '');
    end;
  FreeAndNil(TMere);
  FreeAndNil(TMereTQ);
end;

function TOF_PG_CALINDICBS.RendNbPeriode(DD, DF: TDateTime): integer;
var
  AA, MM1, MM2, JJ, PremMois, PremAnnee, Nb: Word;
begin
  nb := 0;
  DecodeDate(DD, AA, MM1, JJ);
  DecodeDate(DF, AA, MM2, JJ);
  if (MM2 > MM1) then Nb := MM2 - MM1 + 1
  else if MM2 < MM1 then Nb := 13 - MM1 + MM2
  else if (MM2 = MM1) then AglNombreDeMoisComplet(DD, DF, PremMois, PremAnnee, Nb); // recup nombre de mois de la periode
  result := nb;
end;

procedure TOF_PG_CALINDICBS.CalculNbStag(Indic, Sexe, Nat: string;
  TOBCalcul, Ti: TOB);
var
  Sql, StWhere: string;
  Q: TQuery;
  TMere, TFille, T1: Tob;
  i : integer;

begin
  StWhere := RendSexeNat(Sexe, Nat, 'PSI');

  SQL := 'SELECT COUNT (DISTINCT(PEI_INTERIMAIRE)) MT ,PEI_ETABLISSEMENT ' +
    'FROM INTERIMAIRES ' +
    'LEFT JOIN EMPLOIINTERIM ON PSI_INTERIMAIRE=PEI_INTERIMAIRE ' +
    'WHERE PSI_TYPEINTERIM="STA" ' + StWhere +
    'AND ((PEI_DEBUTEMPLOI>="' + USDateTime(DateDeb) + '" ' +
    'AND PEI_DEBUTEMPLOI<="' + USDateTime(DateFin) + '") ' +
    'OR (PEI_FINEMPLOI>="' + USDateTime(DateDeb) + '" ' +
    'AND PEI_FINEMPLOI<="' + USDateTime(DateFin) + '")) ' +
    'GROUP BY PEI_ETABLISSEMENT';

  Q := OpenSql(SQL, True);
  if not Q.eof then
  begin
    TMere := Tob.Create('Les données', nil, -1);
    TMere.LoadDetailDB('Les données', '', '', Q, False);
    Ferme(Q);
    for i := 0 to TMere.detail.count - 1 do
    begin
      T1 := TOB.Create('BILANSOCIAL', TOBCalcul, -1);
      TFille := TMere.detail[i];
      RempliCompl(Indic, TFille, T1, Ti, 'PEI_', ''); // initialisation des valeurs en fonction des ruptures variables
    end;
     FreeAndNil(TMere);
  end;
end;
{ DEB PT7 }
function TOF_PG_CALINDICBS.RendListeSal (TheRacine,Sexe, Nat,StChamp : String; Compte,WithSal,Distinct :Boolean) :  String;
Var
 St : String;
begin
  If Compte then
     St := ' Count (DISTINCT (PSA_SALARIE)) MT,'
  else
    if WithSal then
       St := ' DISTINCT PSA_SALARIE,'
  else
     if Distinct then
       St := ' DISTINCT ';
  st := 'SELECT '+St+'PSA_ETABLISSEMENT,PSA_CATBILAN'+StChamp;
  st := st + TheRacine +
    ' FROM PAIEENCOURS,SALARIES WHERE PSA_SALARIE=PPU_SALARIE  ' +
    'AND ( (PPU_DATEFIN = "' + UsDateTime(DateFin) + '" '+
    'AND (PPU_DATESORTIE IS NULL OR PPU_DATESORTIE="' + USDateTime(Idate1900) + '" ' +  { PT6 }
    'OR (PPU_DATESORTIE > "' + USDateTime(DateFin) + '")) ) '+   { PT6 }
    'OR (PSA_DATEENTREE <= "' + UsDateTime(DateFin) + '" '+  { PT6 }
    'AND ((PSA_SUSPENSIONPAIE = "X" AND PSA_SORTIEDEFINIT <> "X") '+
    'OR (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="' + USDateTime(Idate1900) + '" '+
    'OR PSA_DATESORTIE > "' + USDateTime(DateFin) + '"))))';
  st := st + RendSexeNat(Sexe, Nat);
  result := St;
end;
{ FIN PT7 }

initialization
  registerclasses([TOF_PG_CALINDICBS]);
end.

