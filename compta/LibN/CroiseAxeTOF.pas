{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 26/03/2004
Description .. : BPY le 26/03/004
Suite ........ : * ajout d'un champs CCA_NUMEROPIECE pour pouvoir
Suite ........ : retrouver plus facilement une piece ! et ajout d'une case a
Suite ........ : cocher pour ajouté ce champs dans le traitement !
Suite ........ : * ajout d'une case a cocher TPURGE pour programmer une
Suite ........ : purge avant traitement
Mots clefs ... : TOF;CROISEAXE
*****************************************************************}

unit CroiseAxeTOF;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HSysMenu,
  HCtrls, HEnt1, HMsgBox, UTOF, vierge, UTOB, AglInit, LookUp, Graphics,
  windows, utilPGI, ent1, extctrls, paramsoc, //SG6
  {$IFDEF EAGLCLIENT}
  MaineAGL, eQRS1,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Fe_Main, QRS1, majTable,
  {$ENDIF}
  Filtre, // Filtres
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  M3FP, Messages, HQry, Cube, Stat;

{======================= CROISEMENT DES AXES ==========================}
type
  TCroiseAxe_S = array[1..MaxAxe] of string;
  TCroiseAxe_I = array[1..MaxAxe] of integer;
  TCroiseAxe_D = array[1..2] of double;

  TOF_CPGENERECROISEAXE = class(TOF)
  private
    procedure RemplirCroiseAxe(TOBAnals: TOB; General: string; Periode: integer);
    procedure InitTOBParAxe(TOBParAxe: TOB);
    procedure CreerLesCroisements(TOBParAxe, TOBCCA: TOB; var Compteur: integer);
    function IncLesIndices(var CurInd, MaxInd: TCroiseAxe_I; MaxNiv: integer): boolean;
    procedure ExtraitTranches(Section: string; NumAxe: integer; var T1, T2: string);
    procedure ExerciceOnChange(Sender: TObject);
    procedure OnVentilableChange (Sender : TObject);
  public
    {TControl car cela peut être un THValComboBox ou un TMultiHValComboBox}
    ComboEtab : TControl;

    procedure OnArgument(s: string); override;
    procedure OnUpdate; override;
    {JP 28/06/06 : FQ 16149 : pour appeler ControlEtab à la fin du chargement des filtres}
    procedure AfterSelectFiltre;
  end;

  TOF_CROISEAXE = class(TOF)
  private
    FBoTableToView : Boolean ;   // Gestion MultiSoc SBO 17/08/2005
    procedure RemplirCombi;
    procedure PositionnePeriodes;
  public
    {TControl car cela peut être un THValComboBox ou un TMultiHValComboBox}
    ComboEtab : TControl;

    procedure OnArgument(s: string); override;
    procedure OnLoad; override;
    procedure OnClose; override;
    {JP 28/06/06 : FQ 16149 : pour appeler ControlEtab à la fin du chargement des filtres}
    procedure AfterSelectFiltre;
  end;


  {======================= ANALYSES CUBES ET TOBVIEWER ======================}
  TOF_COMPTAANALYSES = class(TOF)
    procedure OnArgument(s: string); override;
    procedure GereSerie(Pref: string);
    procedure GereDefaut(Pref: string);
    procedure ExerciceOnChange(Sender: TObject);
    procedure OnClose; override;
  private
    c: Char;
    FBoTableToView : Boolean ;   // Gestion MultiSoc SBO 17/08/2005
  public
    {TControl car cela peut être un THValComboBox ou un TMultiHValComboBox}
    ComboEtab : TControl;

    {JP 28/06/06 : FQ 16149 : pour appeler ControlEtab à la fin du chargement des filtres}
    procedure AfterSelectFiltre;
  end;

  TOF_COMPTACUBE = class(TOF_COMPTAANALYSES)
  private
    procedure GereLibres(Pref: string);
  public
  end;

  TOF_ECRCUBE = class(TOF_COMPTACUBE)
    procedure OnArgument(s: string); override;
    procedure GereArgument(s: string);
    procedure CacheFromNam(Nam: string; Cacher: Boolean);
  end;

  TOF_ANACUBE = class(TOF_COMPTACUBE)
    procedure OnArgument(s: string); override;
  end;

  TOF_BALAGEECUBE = class(TOF_ECRCUBE)
    procedure OnArgument(s: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure AuxiElipsisClick(Sender : TObject);
  end;

  TOF_COMPTATOBV = class(TOF_COMPTAANALYSES)
  private
    procedure GereLibres(Pref: string);
  public
  end;

  TOF_ECRTOBV = class(TOF_COMPTATOBV)
    procedure OnArgument(s: string); override;
    procedure AfterShow;
  end;

  TOF_ANATOBV = class(TOF_COMPTATOBV)
    procedure OnArgument(s: string); override;
    procedure AfterShow;
  end;

  TOF_ANAGENETOBV = class(TOF_COMPTATOBV)
    procedure OnArgument(s: string); override;
    public
      procedure OnChangeAxe ( Sender : Tobject);
  end;

  {Anticiper de futures divergences possibles de comportement entre les fiches}
  TOF_CROISEAXECUBE = class(TOF_CROISEAXE)
    procedure OnArgument(s: string); override;
  end;

  TOF_CROISEAXESTAT = class(TOF_CROISEAXE)
    procedure OnArgument(s: string); override;
  end;

  TOF_CROISEAXEETAT = class(TOF_CROISEAXE)
    procedure OnArgument(s: string); override;
  end;

procedure CPLanceFiche_BalAgeeCube(vStRange, vStLequel, vStArgs: string);
procedure CPLanceFiche_BalAgeeCubeM(vStRange, vStLequel, vStArgs: string);

implementation


uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  ULibAnalytique, // AlloueAxe
  UTofMulParamGen; {23/04/07 YMO F5 sur Auxiliaire }

{JP 28/06/06 : Récupération du combo gérant les établissements}
function GetComboEtab(Ecran : TForm) : TControl; forward;
{JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
procedure GereEtablissement(LaTOF : TOF; ComboEtab : TControl); forward;
{JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
procedure ControlEtab(LaTOF : TOF; ComboEtab : TControl); forward;

{==================================================================================================}
{===================== ECRAN DE LANCEMENT DU CUBE DES AXES CROISES ================================}
{==================================================================================================}
procedure ChargeCombiPeriodes(Ecran: TFOrm);
var
  MoisDeb, MoisFin, AnneeDeb, AnneeFin: THValComboBox;
  aDep, aFin: Word;
  i: integer;
  DD, DDep, DFin: TDateTime;
  sPer: string;
  QQ: TQuery;
  JJ, MM: Word;
begin
  AnneeDeb := THValComboBox(Ecran.FindComponent('ANNEEDEB'));
  AnneeFin := THValComboBox(Ecran.FindComponent('ANNEEFIN'));
  MoisDeb := THValComboBox(Ecran.FindComponent('MOISDEB'));
  MoisFin := THValComboBox(Ecran.FindComponent('MOISFIN'));
  // A Changer quand Jocelyne et Philippe seront "harmonisés"
  for i := 1 to 12 do
  begin
    DD := EncodeDate(2000, i, 01);
    sPer := FirstMajuscule(FormatDateTime('mmmm', DD));
    MoisDeb.Items.Add(sPer);
    MoisDeb.Values.Add(IntToStr(i));
    MoisFin.Items.Add(sPer);
    MoisFin.Values.Add(IntToStr(i));
  end;
  DDep := VH^.Entree.Deb;
  DFin := VH^.Entree.Fin;
  QQ := OpenSQL('SELECT MIN(EX_DATEDEBUT), MAX(EX_DATEFIN) FROM EXERCICE', True);
  if not QQ.EOF then
  begin
    DDep := QQ.Fields[0].AsDateTime;
    DFin := QQ.Fields[1].AsDateTime;
  end;
  Ferme(QQ);
  DecodeDate(DDep, aDep, MM, JJ);
  DecodeDate(DFin, aFin, MM, JJ);
  for i := aDep to aFin do
  begin
    sPer := IntToStr(i);
    AnneeDeb.Items.Add(sPer);
    AnneeDeb.Values.Add(sPer);
    AnneeFin.Items.Add(sPer);
    AnneeFin.Values.Add(sPer);
  end;
  AnneeDeb.Value := AnneeDeb.Values[0];
  AnneeFin.Value := AnneeFin.Values[AnneeFin.Values.Count - 1];
  MoisDeb.Value := MoisDeb.Values[0];
  MoisFin.Value := MoisFin.Values[MoisFin.Values.Count - 1];
end;

procedure CombiPeriodes(Ecran: TForm);
var
  PerMin, PerMax, sMoisFin, sMoisDeb: string;
  XX_WHERE: THEdit;
  MoisDeb, MoisFin, AnneeDeb, AnneeFin: THValComboBox;
begin
  AnneeDeb := THValComboBox(Ecran.FindComponent('ANNEEDEB'));
  AnneeFin := THValComboBox(Ecran.FindComponent('ANNEEFIN'));
  MoisDeb := THValComboBox(Ecran.FindComponent('MOISDEB'));
  MoisFin := THValComboBox(Ecran.FindComponent('MOISFIN'));
  sMoisDeb := MoisDeb.Value;
  if Length(sMoisDeb) < 2 then sMoisDeb := '0' + sMoisDeb;
  PerMin := AnneeDeb.Value + sMoisDeb;
  sMoisFin := MoisFin.Value;
  if Length(sMoisFin) < 2 then sMoisFin := '0' + sMoisFin;
  PerMax := AnneeFin.Value + sMoisFin;
  XX_WHERE := THEdit(Ecran.FindComponent('XX_WHEREPERIODE'));
  XX_WHERE.Text := 'CCA_PERIODE>=' + PerMin + ' AND CCA_PERIODE<=' + PerMax;
end;

procedure TOF_CROISEAXE.RemplirCombi;
begin
  ChargeCombiPeriodes(Ecran);
end;

procedure TOF_CROISEAXECUBE.OnArgument(s: string);
begin
  Ecran.HelpContext:=160300200;
  inherited;
end;

procedure TOF_CROISEAXESTAT.OnArgument(s: string);
begin
  Ecran.HelpContext:=160300300;
  inherited;
end;

procedure TOF_CROISEAXEETAT.OnArgument(s: string);
begin
  Ecran.HelpContext:=160300400;
  inherited;
end;

procedure TOF_CROISEAXE.OnArgument(s: string);
begin
  inherited;
  // Gestion MultiSoc SBO 17/08/2005
  if EstMultiSoc then
    begin
    FBoTableToView := V_PGI.EnableTableToView ;
    V_PGI.EnableTableToView := True ;
    end ;
  RemplirCombi;

  {JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements}
  ComboEtab := GetComboEtab(Ecran);
  GereEtablissement(TOF(Self), ComboEtab);
end;

procedure TOF_CROISEAXE.OnLoad;
begin
  inherited;
  PositionnePeriodes;
end;

procedure TOF_CROISEAXE.PositionnePeriodes;
begin
  CombiPeriodes(Ecran);
end;

{==================================================================================================}
{=================== CALCUL DES CROISEMENTS D'AXE POUR TABLE CROISEAXE ============================}
{==================================================================================================}
procedure TOF_CPGENERECROISEAXE.InitTOBParAxe(TOBParAxe: TOB);
begin
  {Créer un TOB avec MaxAxe sous-niveaux qui auront les analytiques en détail pour chauqe Axe}
  TOBParAxe.ClearDetail;
  AlloueAxe( TOBParAxe ) ; // SBO 25/01/2006
end;

function TOF_CPGENERECROISEAXE.IncLesIndices(var CurInd, MaxInd: TCroiseAxe_I; MaxNiv: integer): boolean;
var
  ia, ii: integer;
begin
  Result := False;
  for ia := MaxNiv downto 1 do
  begin
    if CurInd[ia] < MaxInd[ia] then
    begin
      Inc(CurInd[ia]);
      for ii := ia + 1 to MaxNiv do CurInd[ii] := 0;
      Result := True;
      Break;
    end;
  end;
end;

procedure TOF_CPGENERECROISEAXE.ExtraitTranches(Section: string; NumAxe: integer; var T1, T2: string);
var
  LeFb: TFichierBase;
  Lg: integer;
begin
  T1 := Section;
  T2 := '';
  LeFb := TFichierBase(Ord(fbAxe1) + NumAxe - 1);
  if VH^.Cpta[LeFb].Structure then
  begin
    Lg := VH^.SousPlanAxe[LeFb, 1].Longueur;
    T1 := Copy(Section, 1, Lg);
    T2 := Copy(Section, Lg + 1, 16);
  end;
end;

procedure TOF_CPGENERECROISEAXE.ExerciceOnChange(Sender: TObject);
begin
  ExoToDates(GetControlText('Y_EXERCICE'), THEdit(GetControl('Y_DATECOMPTABLE')), THEdit(GetControl('Y_DATECOMPTABLE_')));
end;

procedure TOF_CPGENERECROISEAXE.OnVentilableChange(Sender: TObject);
var
  stAxe : string;
  stPlus : string;
begin
  stAxe := GetControlText('AXE');
  stPlus := '';
  if (Pos('A1',stAxe) > 0 ) then stPlus := stPlus + ' AND G_VENTILABLE1="X"';
  if (Pos('A2',stAxe) > 0 ) then stPlus := stPlus + ' AND G_VENTILABLE2="X"';
  if (Pos('A3',stAxe) > 0 ) then stPlus := stPlus + ' AND G_VENTILABLE3="X"';
  if (Pos('A4',stAxe) > 0 ) then stPlus := stPlus + ' AND G_VENTILABLE4="X"';
  if (Pos('A5',stAxe) > 0 ) then stPlus := stPlus + ' AND G_VENTILABLE5="X"';
  SetControlProperty('Y_GENERAL','Plus',stPlus);
  SetControlProperty('Y_GENERAL_','Plus',stPlus);
end;

procedure TOF_CPGENERECROISEAXE.CreerLesCroisements(TOBParAxe, TOBCCA: TOB; var Compteur: integer);
var
  isect, ia, NumAxe, MaxNiv, CurNiv, Periode, numpieces: integer;
  Ax, General, Journal, Exercice, Etablissement, Section, Tranche1, Tranche2: string;
  QualifPiece : string;
  Debit: Boolean;
  Pourc, TotP, X: Double;
  TOBAx, TOBAAA, TOBInit, TOBTemp: TOB;
  MaxInd, CurInd: TCroiseAxe_I;
  TabSection: TCroiseAxe_S;
  TabMontant: TCroiseAxe_D;
begin
  {Eliminer les niveaux sur des axes non affectés}
  for ia := TOBParAxe.Detail.Count - 1 downto 0 do
  begin
    TOBAx := TOBParAxe.Detail[ia];
    if (TOBAx.Detail.Count <= 0) then TOBAx.Free;
  end;

  MaxNiv := TOBParAxe.Detail.Count;
  FillChar(MaxInd, Sizeof(MaxInd), #0);
  FillChar(CurInd, Sizeof(CurInd), #0);
  for ia := 1 to MaxNiv do MaxInd[ia] := TOBParAxe.Detail[ia - 1].Detail.Count - 1;

  {Capturer les informations de la ligne de compta générale}
  TOBInit := TOBParAxe.Detail[0].Detail[0];
  TotP := TOBInit.GetValue('Y_TOTALECRITURE');
  Debit := (Arrondi(TOBInit.GetValue('Y_DEBIT'), 6) <> 0);
  General := TOBInit.GetValue('Y_GENERAL');
  Journal := TOBInit.GetValue('Y_JOURNAL');
  Etablissement := TOBInit.GetValue('Y_ETABLISSEMENT');
  Exercice := TOBInit.GetValue('Y_EXERCICE');
  Periode := TOBInit.GetValue('Y_PERIODE');
  numpieces := TOBInit.GetValue('Y_NUMEROPIECE');
  QualifPiece := TOBInit.GetValue('Y_QUALIFPIECE');
  FillChar(TabMontant, Sizeof(TabMontant), #0);

  {Balayer tous les croisements posibles}
  repeat
    Pourc := 1.0;
    FillChar(TabSection, Sizeof(TabSection), #0);

    {Constituer la combinaison des sections, calculer les % résultant}
    for CurNiv := 1 to MaxNiv do
    begin
      TOBTemp := TOBParAxe.Detail[CurNiv - 1].Detail[CurInd[CurNiv]];
      Ax := TOBTemp.GetValue('Y_AXE');
      NumAxe := Ord(Ax[2]) - 48;
      Pourc := Pourc * TOBTemp.GetValue('Y_POURCENTAGE') / 100;
      TabSection[NumAxe] := TOBTemp.GetValue('Y_SECTION');
    end;

    TabMontant[2 - Ord(Debit)] := Arrondi(TotP * Pourc, V_PGI.OkDecV);

    if (GetCheckBoxState('TNUMPIECE') = cbChecked) then
      TOBAAA := TOBCCA.FindFirst(['CCA_JOURNAL', 'CCA_ETABLISSEMENT', 'CCA_NUMEROPIECE', 'CCA_QUALIFPIECE', 'CCA_SECTIONA1', 'CCA_SECTIONA2', 'CCA_SECTIONA3', 'CCA_SECTIONA4', 'CCA_SECTIONA5'],
        [Journal, Etablissement, numpieces, QualifPiece, TabSection[1], TabSection[2], TabSection[3], TabSection[4], TabSection[5]], True)
    else
      TOBAAA := TOBCCA.FindFirst(['CCA_JOURNAL', 'CCA_ETABLISSEMENT','CCA_QUALIFPIECE', 'CCA_SECTIONA1', 'CCA_SECTIONA2', 'CCA_SECTIONA3', 'CCA_SECTIONA4', 'CCA_SECTIONA5'],
        [Journal, Etablissement, QualifPiece, TabSection[1], TabSection[2], TabSection[3], TabSection[4], TabSection[5]], True);

    {création nouvelle TOB}
    if TOBAAA = nil then
    begin
      {Initialisation de la TOB}
      TOBAAA := TOB.Create('CROISEAXE', TOBCCA, -1);
      TOBAAA.PutValue('CCA_GENERAL', General);
      TOBAAA.PutValue('CCA_JOURNAL', Journal);
      TOBAAA.PutValue('CCA_EXERCICE', Exercice);
      TOBAAA.PutValue('CCA_PERIODE', Periode);
      TOBAAA.PutValue('CCA_ETABLISSEMENT', Etablissement);
      if (GetCheckBoxState('TNUMPIECE') = cbChecked) then TOBAAA.PutValue('CCA_NUMEROPIECE', numpieces)
      else TOBAAA.PutValue('CCA_NUMEROPIECE', 0);
      for isect := 1 to 5 do
      begin
        Section := TabSection[iSect];
        ExtraitTranches(Section, iSect, Tranche1, Tranche2);
        TOBAAA.PutValue('CCA_SECTIONA' + IntToStr(iSect), Section);
        TOBAAA.PutValue('CCA_TRANCHE1A' + IntToStr(iSect), Tranche1);
        TOBAAA.PutValue('CCA_TRANCHE2A' + IntToStr(iSect), Tranche2);
      end;
      Inc(Compteur);
      TOBAAA.PutValue('CCA_COMPTEUR', Compteur);
      TOBAAA.PutValue('CCA_QUALIFPIECE', QualifPiece);
    end;

    {Cumul dans TOB (forcement existante maintenant !)}
    if (Debit) then
    begin
      X := TOBAAA.GetValue('CCA_DEBIT');
      X := Arrondi(X + TabMontant[1], V_PGI.OkDecV);
      TOBAAA.PutValue('CCA_DEBIT', X);
    end
    else
    begin
      X := TOBAAA.GetValue('CCA_CREDIT');
      X := Arrondi(X + TabMontant[2], V_PGI.OkDecV);
      TOBAAA.PutValue('CCA_CREDIT', X);
    end;

    if not IncLesIndices(CurInd, MaxInd, MaxNiv) then Break;
  until False;
end;

procedure TOF_CPGENERECROISEAXE.RemplirCroiseAxe(TOBAnals: TOB; General: string; Periode: integer);
var
  TOBCCA, TOBDetA, TOBParAxe, TOBPADet: TOB;
  i, NumAxe, Compteur: integer;
  Rupt: boolean;
  OldJal, Jal, Ax: string;
  OldNum, OldLigne, Num, Ligne: integer;
begin
  {Vider la table}
  ExecuteSQL('DELETE FROM CROISEAXE WHERE CCA_GENERAL="' + General + '" AND CCA_PERIODE=' + IntToStr(Periode));

  {Traitement des analytiques pour un compte général et une période donnée}
  TOBCCA := TOB.Create('', nil, -1);
  TOBParAxe := TOB.Create('', nil, -1);

  InitTOBParAxe(TOBParAxe);
  OldJal := '';
  OldNum := -1;
  OldLigne := -1;
  Compteur := 0;

  for i := 0 to TOBAnals.Detail.Count - 1 do
  begin
    TOBDetA := TOBAnals.Detail[i];
    Jal := TOBDetA.GetValue('Y_JOURNAL');
    Num := TOBDetA.GetValue('Y_NUMEROPIECE');
    Ligne := TOBDetA.GetValue('Y_NUMLIGNE');

    if (i <= 0) then Rupt := False
    else Rupt := ((Jal <> OldJal) or (Num <> OldNum) or (Ligne <> OldLigne));

    if (Rupt) then
    begin
      CreerLesCroisements(TOBParAxe, TOBCCA, Compteur);
      InitTOBParAxe(TOBParAxe);
    end;

    Ax := TOBDetA.GetValue('Y_AXE');
    NumAxe := Ord(Ax[2]) - 48;

    TOBPADet := TOB.Create('', TOBParAxe.Detail[NumAxe - 1], -1);
    TOBPADet.Dupliquer(TOBDetA, False, True);

    OldJal := Jal;
    OldNum := Num;
    OldLigne := Ligne;
  end;

  {Dernière rupture}
  CreerLesCroisements(TOBParAxe, TOBCCA, Compteur);
  InitTOBParAxe(TOBParAxe);

  {Enregistrement}
  TOBCCA.InsertDB(nil);

  {Libérations}
  TOBCCA.Free;
  TOBParAxe.Free;
end;

procedure TOF_CPGENERECROISEAXE.OnArgument(s: string);
begin
  Ecran.HelpContext:=160300100;
  {JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements}
  ComboEtab := GetComboEtab(Ecran);
  GereEtablissement(TOF(Self), ComboEtab);

  THValComboBox(GetControl('Y_EXERCICE')).OnChange := ExerciceOnChange;
  THValComboBox(GetControl('AXE')).OnChange := OnVentilableChange;
  OnVentilableChange(nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENERECROISEAXE.AfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  ControlEtab(TOF(Self), ComboEtab);
end;

procedure TOF_CPGENERECROISEAXE.OnUpdate;
var
  C: TPageControl;
  sWhere, SQL, SQLDepart, General: string;
  i, Periode: integer;
  TOBDist, TOBDistU, TOBAnals: TOB;
begin
  inherited;

  if HShowMessage('0;Génération des cumuls analytiques croisés;Confirmez-vous le traitement de génération ?;Q;YN;Y;N;', '', '') <> mrYes then Exit;
  C := TPageControl(Ecran.FindComponent('PAGES'));
  if (C = nil) then Exit;

  {suppression du contenu de la table temp !}
  if (GetCheckBoxState('TPURGE') = cbChecked) then
  begin
    if (HShowMessage('0;Génération des cumuls analytiques croisés;ATTENTION :' + #10 + #13 + 'Vous avez choisi de purger les informations précédentes.' + #10 + #13 +
      'Voulez-vous continuer ?;Q;YN;Y;N;', '', '') = mrYes) then
      ExecuteSQL('DELETE FROM CROISEAXE')
    else
      SetControlChecked('TPURGE', false);
  end;

  {Etude des critères}
  sWhere := RecupWhereCritere(C);
(* CA - 04/05/2007 - On gère désormais la qualif pièce dans les critères
  if (sWhere <> '') then sWhere := sWhere + ' AND Y_QUALIFPIECE="N"'
  else sWhere := 'WHERE Y_QUALIFPIECE="N"';
*)
  {Distinct des combinaisons Section/Axe}
  SQL := 'SELECT DISTINCT Y_GENERAL, Y_PERIODE FROM ANALYTIQ ' + sWhere;
  TOBDist := TOB.Create('', nil, -1);
  TOBDist.LoadDetailFromSQL(SQl, false, true);

  {Traitement détaillé pour chaque combinaison}
  TOBAnals := TOB.Create('', nil, -1);
  SQLDepart := 'SELECT Y_GENERAL, Y_AXE, Y_SECTION, Y_ETABLISSEMENT, Y_JOURNAL, Y_EXERCICE, Y_PERIODE, Y_DATECOMPTABLE, '
    + 'Y_DEBIT, Y_CREDIT, Y_DEBITDEV, Y_CREDITDEV, Y_NUMEROPIECE, Y_NUMLIGNE, Y_TOTALECRITURE, Y_POURCENTAGE, Y_QUALIFPIECE  '
    + 'FROM ANALYTIQ ' + sWhere;

  for i := 0 to TOBDist.Detail.Count - 1 do
  begin
    TOBDistU := TOBDist.Detail[i];
    General := TOBDistU.GetValue('Y_GENERAL');
    Periode := TOBDistU.GetValue('Y_PERIODE');
    SQL := SQLDepart + ' AND Y_GENERAL="' + General + '" AND Y_PERIODE=' + IntToStr(Periode)
      + ' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE';
    TOBAnals.LoadDetailFromSQL(SQl, false, true);
    RemplirCroiseAxe(TOBAnals, General, Periode);
  end;

  TOBDist.Free;
  TOBAnals.Free;
  HShowMessage('0;Génération des cumuls analytiques croisés;Le traitement s''est correctement effectué;E;O;O;O;', '', '');
end;

{==================================================================================================}
{======================== GESTION DES CUBES ECRITURES ET ANALYTIQUES ==============================}
{==================================================================================================}
procedure TOF_COMPTAANALYSES.OnArgument(s: string);
begin
  // Gestion MultiSoc SBO 17/08/2005
  if EstMultiSoc then
    begin
    FBoTableToView := V_PGI.EnableTableToView ;
    V_PGI.EnableTableToView := True ;
    end ;
  if C <> '' then
    THValComboBox(GetControl(C + '_EXERCICE')).OnChange := ExerciceOnChange;

  {JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements}
  ComboEtab := GetComboEtab(Ecran);
  GereEtablissement(TOF(Self), ComboEtab);

  if Ecran is TFCube then
    TFCube(Ecran).OnAfterSelectFiltre := AfterSelectFiltre
  else if Ecran is TFStat then
    TFStat(Ecran).OnAfterSelectFiltre := AfterSelectFiltre;
end;

procedure TOF_COMPTAANALYSES.ExerciceOnChange(Sender: TObject);
begin
  ExoToDates(GetControlText(C + '_EXERCICE'), THEdit(GetControl(C + '_DATECOMPTABLE')), THEdit(GetControl(C + '_DATECOMPTABLE_')));
end;

procedure TOF_COMPTAANALYSES.GereSerie(Pref: string);
var
  BLibre: TBevel;
  i, MaxTL: integer;
  THVC: THMultiValComboBox;
  THL: THLabel;
  stNom: string;
begin
  if Pref = 'Y' then
  begin
    {Pas de tables libres mouvements sur analytiques en S3 et S5}
    if EstSerie(S3) or EstSerie(S5) then
    begin
      BLibre := TBevel(Ecran.FindComponent('BLIBREY'));
      if BLibre <> nil then BLibre.Visible := False;
      THL := THLabel(Ecran.FindComponent('TBLIBREY'));
      if THL <> nil then THL.Visible := False;
      for i := 0 to 9 do
      begin
        stNom := 'Y_TABLE' + IntToStr(i);
        THVC := THMultiValComboBox(Ecran.FindComponent(stNom));
        if THVC <> nil then THVC.Visible := False;
        THL := THLabel(Ecran.FindComponent('T' + stNom));
        if THL <> nil then THL.Visible := False;
      end;
    end;
  end
  else
    if Pref = 'E' then
  begin
    {1 table libre Ecriture en S3, 2 en S5, 4 en S7}
    if EstSerie(S3) then MaxTL := 1
    else if EstSerie(S5) then MaxTL := 2
    else MaxTL := 4;
    for i := MaxTL to 3 do
    begin
      stNom := 'E_TABLE' + IntToStr(i);
      THVC := THMultiValComboBox(Ecran.FindComponent(stNom));
      if THVC <> nil then THVC.Visible := False;
      THL := THLabel(Ecran.FindComponent('T' + stNom));
      if THL <> nil then THL.Visible := False;
    end;
  end
  else
    if EstSerie(S3) then
  begin
    {3 tables libres maximum pour les fiches de base en S3}
    for i := 3 to 9 do
    begin
      stNom := Pref + '_TABLE' + IntToStr(i);
      THVC := THMultiValComboBox(Ecran.FindComponent(stNom));
      if THVC <> nil then THVC.Visible := False;
      THL := THLabel(Ecran.FindComponent('T' + stNom));
      if THL <> nil then THL.Visible := False;
    end;
  end;
end;

procedure TOF_COMPTAANALYSES.GereDefaut(Pref: string);
var
  Exo, QualifPiece: THValComboBox;
  DateDeb, DateFin: THCritMaskEdit;
begin
  Exo := THValComboBox(Ecran.FindComponent(Pref + '_EXERCICE'));
  QualifPiece := THValComboBox(Ecran.FindComponent(Pref + '_QUALIFPIECE'));
  DateDeb := THCritMaskEdit(Ecran.FindComponent(Pref + '_DATECOMPTABLE'));
  DateFin := THCritMaskEdit(Ecran.FindComponent(Pref + '_DATECOMPTABLE_'));
  if VH^.CPExoRef.Code <> '' then
  begin
    Exo.Value := VH^.CPExoRef.Code;
    DateDeb.Text := DateToStr(VH^.CPExoRef.Deb);
    DateFin.Text := DateToStr(VH^.CPExoRef.Fin);
  end else
  begin
    Exo.Value := VH^.Entree.Code;
    DateDeb.Text := DateToStr(V_PGI.DateEntree);
    DateFin.Text := DateToStr(V_PGI.DateEntree);
  end;
  QualifPiece.Value := 'N';
  if Pref = 'Y' then
  begin
    //Ax:=THValComboBox(Ecran.FindComponent('Y_AXE')) ; Ax.Value:='A1' ;  //SG6
  end;
end;

procedure TOF_COMPTACUBE.GereLibres(Pref: string);
var
  ZL: TTabSheet;
  Suff, stNom: string;
  blAvecGene, blAvecPref, blAvecSuff: Boolean;
  i: Integer;
  THVC: THMultiValComboBox;
  THL: THLabel;
begin
  { Détermination des blocs (Suffs) (Tiers ou Sections) }
  if Pref = 'E'
    then Suff := 'T'
  else Suff := 'S';
  ZL := TTabSheet(Ecran.FindComponent('PLIBRES'));
  { Gestion accès au bloc 'Généraux' }
  LibellesTableLibre(ZL, 'TG_TABLE', 'G_TABLE', 'G');
  blAvecGene := ZL.TabVisible;
  if not ZL.TabVisible then
  begin
    ZL.TabVisible := True;
    for i := 0 to 9 do
    begin
      stNom := 'G_TABLE' + IntToStr(i);
      THVC := THMultiValComboBox(Ecran.FindComponent(stNom));
      if THVC <> nil then THVC.Enabled := False;
      THL := THLabel(Ecran.FindComponent('T' + stNom));
      if THL <> nil then THL.Enabled := False;
    end;
  end;
  { Gestion accès au bloc 'Suff' }
  LibellesTableLibre(ZL, 'T' + Suff + '_TABLE', Suff + '_TABLE', Suff);
  blAvecSuff := ZL.TabVisible;
  if not ZL.TabVisible then
  begin
    ZL.TabVisible := True;
    for i := 0 to 9 do
    begin
      stNom := Suff + '_TABLE' + IntToStr(i);
      THVC := THMultiValComboBox(Ecran.FindComponent(stNom));
      if THVC <> nil then THVC.Enabled := False;
      THL := THLabel(Ecran.FindComponent('T' + stNom));
      if THL <> nil then THL.Enabled := False;
    end;
  end;
  { Gestion accès au bloc 'Ecriture' de type 'Pref' }
  LibellesTableLibre(ZL, 'T' + Pref + '_TABLE', Pref + '_TABLE', Pref);
  blAvecPref := ZL.TabVisible;
  if not ZL.TabVisible then
  begin
    ZL.TabVisible := True;
    for i := 0 to 3 do
    begin
      stNom := Pref + '_TABLE' + IntToStr(i);
      THVC := THMultiValComboBox(Ecran.FindComponent(stNom));
      if THVC <> nil then THVC.Enabled := False;
      THL := THLabel(Ecran.FindComponent('T' + stNom));
      if THL <> nil then THL.Enabled := False;
    end;
  end;
  { Si rien visible, on cache la page }
  if not (blAvecGene or blAvecPref or blAvecSuff) then ZL.TabVisible := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_COMPTAANALYSES.AfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  ControlEtab(TOF(Self), ComboEtab);
end;

procedure TOF_COMPTATOBV.GereLibres(Pref: string);
var
  ZL: TTabSheet;
  Suff: string;
begin
  if Pref = 'E' then Suff := 'T' else Suff := 'S'; {Tiers ou Sections}
  ZL := TTabSheet(Ecran.FindComponent('GLIBRES'));
  LibellesTableLibre(ZL, 'TG_TABLE', 'G_TABLE', 'G');
  ZL := TTabSheet(Ecran.FindComponent(Suff + 'LIBRES'));
  LibellesTableLibre(ZL, 'T' + Suff + '_TABLE', Suff + '_TABLE', Suff);
end;

procedure TOF_ECRCUBE.CacheFromNam(Nam: string; Cacher: boolean);
var
  THL: THLabel;
  THC: THCritMaskEdit;
begin
  THC := THCritMaskEdit(Ecran.FindComponent(Nam));
  if THC <> nil then THC.Visible := not Cacher;
  THL := THLabel(Ecran.FindComponent('T' + Nam));
  if THL <> nil then THL.Visible := not Cacher;
  THC := THCritMaskEdit(Ecran.FindComponent(Nam + '_'));
  if THC <> nil then THC.Visible := not Cacher;
  THL := THLabel(Ecran.FindComponent('T' + Nam + '_'));
  if THL <> nil then THL.Visible := not Cacher;
end;

procedure TOF_ECRCUBE.GereArgument(s: string);
var
  SQL: string;
begin
  if s = '' then Exit;
  if s = 'CONTREPARTIE' then
  begin
    {Rendre les fourchettes de contrepartie visibles}
    CacheFromNam('E_CONTREPARTIEAUX', False);
    CacheFromNam('E_CONTREPARTIEGEN', False);
    //RR on force ce mode pour le look 2003
    CacheFromNam('E_AUXILIAIRE', True);
    CacheFromNam('E_GENERAL', True);
    (* VL le 20/03/2002
    CacheFromNam('E_CONTREPARTIEGEN',FAlse) ;
    *)
    {Manipulation SQL pour jointure sur contrepartie}
    SQL := TFCube(Ecran).FromSQL;
    SQL := FindEtReplace(SQL, 'E_AUXILIAIRE=T_AUXILIAIRE', 'E_CONTREPARTIEAUX=T_AUXILIAIRE', True);
    SQL := FindEtReplace(SQL, 'E_GENERAL=G_GENERAL', 'E_CONTREPARTIEGEN=G_GENERAL', True);
    TFCube(Ecran).FromSQL := SQL;
    {Filtres, titres, caption}
    (* VL le 20/03/2002
    TFCube(Ecran).FNomFiltre:=TFCube(Ecran).Name+'CONTREP' ;
    *)
    TFCube(Ecran).FNomFiltre := Copy(TFCube(Ecran).Name + 'CONTREP', 1, 17); // Modification
    TFCube(Ecran).Caption := TFCube(Ecran).Caption + TraduireMemoire(' sur contreparties');
    UpdateCaption(TFCube(Ecran));
    TGroupBox(Ecran.FindComponent('GBG')).Caption := TGroupBox(Ecran.FindComponent('GBG')).Caption + TraduireMemoire(' de contrepartie');
    TGroupBox(Ecran.FindComponent('GBT')).Caption := TGroupBox(Ecran.FindComponent('GBT')).Caption + TraduireMemoire(' de contrepartie');
  end;
end;

procedure TOF_ECRCUBE.OnArgument(s: string);
begin
  C := 'E';
  inherited;
  if S = 'CONTREPARTIE' then Ecran.HelpContext := 999999977 else Ecran.HelpContext := 999999976;
  GereSerie('E');
  GereLibres('E');
  GereDefaut('E');
  GereArgument(s);
end;

procedure TOF_ANACUBE.OnArgument(s: string);
var
  i: integer;
  axeventil: boolean;
begin
  C := 'Y';

  //SG6
  for i:=1 to 5 do
  begin
    axeventil := GetParamSocSecur('SO_VENTILA' + IntToStr(i),False);
    SetControlVisible('TY_AXE' + IntToStr(i) + '_', axeventil);
    SetControlVisible('TY_AXE' + IntToStr(i), axeventil);
    SetControlVisible('Y_SOUSPLAN' + IntToStr(i), axeventil);
    SetControlVisible('Y_SOUSPLAN' + IntToStr(i) + '_', axeventil);
  end;

  if VH^.AnaCroisaxe then
  begin
    SetControlVisible('TY_AXE', False);
    SetControlVisible('Y_AXE', False);
    SetControlVisible('TY_SECTION', False);
    SetControlVisible('TY_SECTION_', False);
    SetControlVisible('Y_SECTION', False);
    SetControlVisible('Y_SECTION_', False);
  end;



  inherited;
  Ecran.helpContext := 999999978;
  GereSerie('Y');
  GereLibres('Y');
  GereDefaut('Y');
end;


procedure TOF_ECRTOBV.OnArgument(s: string);
begin
  C := 'E';
  inherited;
  Ecran.HelpContext := 999999109;
  GereSerie('E');
  GereLibres('E');
  GereDefaut('E');
  TFStat(Ecran).OnAfterFormShow := AfterShow;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ECRTOBV.AfterShow;
begin
  {$IFDEF CCSTD}
  TFStat(Ecran).ListeFiltre.ForceAccessibilite := faPublic;
  {$ENDIF}
end;
////////////////////////////////////////////////////////////////////////////////

procedure TOF_ANATOBV.OnArgument(s: string);
var i :integer;
begin
 //SG6 16/12/2004 Ventilation Multi Axxe
  if VH^.AnaCroisaxe then
  begin
    SetControlVisible('TY_AXE', False);
    SetControlVisible('Y_AXE', False);
    SetcontrolVisible('TY_SECTION', False);
    SetControlVisible('TY_SECTION_', False);
    SetControlVisible('S_SECTION', False);
    SetControlVisible('S_SECTION_', False);
    for i:=1 to 5 do
    begin
      SetControlVisible('TY_SECTION'+IntToStr(i),GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
      SetControlVisible('TY_SECTION'+IntToStr(i)+'_',GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
      SetControlVisible('Y_SOUSPLAN'+IntToStr(i),GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
      SetControlVisible('Y_SOUSPLAN'+IntToStr(i)+'_',GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
    end;
  end
  else
  begin
    TTabSheet(getcontrol('PMultiAxe', true)).Tabvisible := False;
  end;


  C := 'Y';
  inherited;
  Ecran.HelpContext := 999999986;
  GereSerie('Y');
  GereLibres('Y');
  GereDefaut('Y');
  TFStat(Ecran).OnAfterFormShow := AfterShow;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ANATOBV.AfterShow;
begin
  {$IFDEF CCSTD}
  TFStat(Ecran).ListeFiltre.ForceAccessibilite := faPublic;
  {$ENDIF}

end;
////////////////////////////////////////////////////////////////////////////////

{ TOF_BALAGEECUBE }

procedure TOF_BALAGEECUBE.OnArgument(s: string);
begin
  C := 'E';
  inherited;
  // Initialisation des critères
  THValComboBox(GetControl('E_QUALIFPIECE')).Value := 'N';
  THValComboBox(GetControl('E_DEVISE')).ItemIndex := 0;
  THValComboBox(GetControl('E_MODEPAIE')).ItemIndex := 0;
  THValComboBox(GetControl('E_EXERCICE')).ItemIndex := 0;

  {JP 28/06/06 : FQ 16149 : Nouvelle gestion des établissements
  THMultiValComboBox(GetControl('E_ETABLISSEMENT')).value := '';}

  THMultiValComboBox(GetControl('E_NATUREPIECE')).value := '';
  THMultiValComboBox(GetControl('E_ETATLETTRAGE')).value := '';
  SetControlText('E_DATECOMPTABLE', DateToStr(idate1900));
  SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.EnCours.Fin));
  SetControlText('DATESITUATION', DateToStr(VH^.EnCours.Fin));
  THNumEdit(GetControl('MONTANTMIN')).value := 0;
  THNumEdit(GetControl('MONTANTMAX')).value := 99999999;

  // si mode IAS14 : Ajout de la table ecriture
  if pos('IAS14', s) > 0 then
  begin
    TFCube(Ecran).FromSQL := 'ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' +
      'LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' +
      'LEFT JOIN ANALYTIQ ON ( E_NUMEROPIECE=Y_NUMEROPIECE AND ' +
      'E_JOURNAL=Y_JOURNAL AND ' +
      'E_NUMLIGNE=Y_NUMLIGNE AND ' +
      'E_EXERCICE=Y_EXERCICE AND ' +
      'E_QUALIFPIECE=Y_QUALIFPIECE )';

    Ecran.caption := 'Balance ventilée IAS 14'; // 13805
    Ecran.HelpContext := 999999990; // 14678

    UpdateCaption(Ecran);
  end;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('E_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_BALAGEECUBE.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


procedure TOF_BALAGEECUBE.OnLoad;
begin
  inherited;
  // Test Date situation obligatoire
  if GetControlText('DATESITUATION') = '' then
    SetControlText('DATESITUATION', DateToStr(VH^.EnCours.Fin));
  // Test Etat lettrage
  if GetControlText('E_ETATLETTRAGE') = '' then
    THMultiValComboBox(GetControl('E_ETATLETTRAGE')).value := '';
end;

procedure TOF_BALAGEECUBE.OnUpdate;
var
  stWhereEtatLett, stWhere, stDateSitu: string;
  MontantMax, MontantMin: Real;
begin
  inherited;
  // Critères auto
  stWhere := 'E_ETATLETTRAGE<>"RI" AND E_ECRANOUVEAU<>"CLO" AND E_ECRANOUVEAU<>"OAN"';
  // Critère date suivant exoV8
  if (VH^.ExoV8.Code <> '') then
    stWhere := stWhere + ' AND E_DATECOMPTABLE >= "' + USDateTime(VH^.ExoV8.Deb) + '"';
  // Critères si qualifpiece = "tous"
  if THValComboBox(GetControl('E_QUALIFPIECE')).value = '' then
    stWhere := stWhere + ' AND E_QUALIFPIECE<>"C" AND E_QUALIFPIECE<>"R"';
  // Critères sur etat lettrage + date situation
  stDateSitu := USDateTime(StrToDate(GetControlText('DATESITUATION')));
  if THMultiValComboBox(GetControl('E_ETATLETTRAGE')).Tous then
  begin
    stWhereEtatLett := THMultiValComboBox(GetControl('E_ETATLETTRAGE')).GetSQLValue;
    stWhere := stWhere + ' AND (E_ETATLETTRAGE <> "TL" OR'
      + ' (E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX>"' + stDateSitu + '"))'
  end
  else
    if Pos('E_ETATLETTRAGE="TL"', TFCube(Ecran).WhereCritere) > 0 then
    TFCube(Ecran).WhereCritere := FindEtReplace(TFCube(Ecran).WhereCritere,
      'E_ETATLETTRAGE="TL"',
      '(E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX>"' + stDateSitu + '")',
      False);
  // Critères sur les montants
  MontantMin := THNumEdit(GetControl('MONTANTMIN')).value;
  MontantMax := THNumEdit(GetControl('MONTANTMAX')).value;
  {$IFDEF EAGLCLIENT}
  if (MontantMin > 0) or (MontantMax < 99999999) then
    stWhere := stWhere + ' AND ((E_DEBIT=0 AND '
      + 'ABS(E_CREDIT) >= ' + StrFPoint(MontantMin) + ' AND '
      + 'ABS(E_CREDIT) <=' + StrFPoint(MontantMax)
      + ') OR (E_CREDIT=0 AND '
      + 'ABS(E_DEBIT) >=' + StrFPoint(MontantMin) + ' AND '
      + 'ABS(E_DEBIT) <=' + StrFPoint(MontantMax) + '))';
  {$ELSE}
  if (MontantMin > 0) or (MontantMax < 99999999) then
    stWhere := stWhere + ' AND ((E_DEBIT=0 AND '
      + DB_ABS('E_CREDIT') + '>=' + StrFPoint(MontantMin) + ' AND '
      + DB_ABS('E_CREDIT') + '<=' + StrFPoint(MontantMax)
      + ') OR (E_CREDIT=0 AND '
      + DB_ABS('E_DEBIT') + '>=' + StrFPoint(MontantMin) + ' AND '
      + DB_ABS('E_DEBIT') + '<=' + StrFPoint(MontantMax) + '))';
  {$ENDIF}
  // Mise à jour de la requête
  TFCube(Ecran).WhereCritere := TFCube(Ecran).WhereCritere + ' AND ' + stWhere;

end;

// ==== PROCEDURE DE LANCEMENT DU CUBE TYPE BALANCE AGEE
procedure CPLanceFiche_BalAgeeCube(vStRange, vStLequel, vStArgs: string);
begin
  AGLLanceFiche('CP', 'CPBALAGEE_CUBE', vStRange, vStLequel, vStArgs);
end;

procedure CPLanceFiche_BalAgeeCubeM(vStRange, vStLequel, vStArgs: string);
begin
  AGLLanceFiche('CP', 'CPBALAGEE_CUBEM', vStRange, vStLequel, vStArgs);
end;


{ TOF_ANAGENETOBV }

procedure TOF_ANAGENETOBV.OnArgument(s: string);
var i : integer;
begin
  //SG6 16/12/2004 Ventilation Multi Axxe
  if VH^.AnaCroisaxe then
  begin
    SetControlVisible('TY_AXE', False);
    SetControlVisible('Y_AXE', False);
    SetcontrolVisible('TY_SECTION', False);
    SetControlVisible('TY_SECTION_', False);
    SetControlVisible('Y_SECTION', False);
    SetControlVisible('Y_SECTION_', False);
    for i:=1 to 5 do
    begin
      SetControlVisible('TY_SECTION'+IntToStr(i),GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
      SetControlVisible('TY_SECTION'+IntToStr(i)+'_',GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
      SetControlVisible('Y_SOUSPLAN'+IntToStr(i),GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
      SetControlVisible('Y_SOUSPLAN'+IntToStr(i)+'_',GetParamSocSecur('SO_VENTILA'+IntToStr(i),False));
    end;
  end
  else
  begin
    TTabSheet(getcontrol('PMultiAxe', true)).Tabvisible := False;
    THValComboBox(Getcontrol('Y_AXE', true)).OnChange := OnChangeAxe;
    SetControlText ('Y_AXE','A1');
    OnChangeAxe(nil);
  end;
  C := 'Y';
  inherited;
  GereSerie('Y');
  GereLibres('Y');
  GereDefaut('Y');
end;

procedure TOF_COMPTAANALYSES.OnClose;
begin
  inherited;
  // Gestion MultiSoc SBO 17/08/2005
  if EstMultiSoc then
    V_PGI.EnableTableToView := FBoTableToView ;
end;

procedure TOF_CROISEAXE.OnClose;
begin
  inherited;
  // Gestion MultiSoc SBO 17/08/2005
  if EstMultiSoc then
    V_PGI.EnableTableToView := FBoTableToView ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CROISEAXE.AfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  ControlEtab(TOF(Self), ComboEtab);
end;

{JP 28/06/06 : Récupération du combo gérant les établissements
{---------------------------------------------------------------------------------------}
function GetComboEtab(Ecran : TForm) : TControl;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := nil;
  if not Assigned(Ecran) then Exit;

  for n := 0 to Ecran.ComponentCount - 1 do
    if (Pos('ETABLISSEMENT', Ecran.Components[n].Name) > 0) and
       ((Ecran.Components[n] is THValComboBox) or
        (Ecran.Components[n] is THMultiValComboBox)) then begin
      Result := TControl(Ecran.Components[n]);
      Break;
    end;
end;

{---------------------------------------------------------------------------------------}
procedure GereEtablissement(LaTOF : TOF; ComboEtab : TControl);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ComboEtab) and Assigned(LaTof) then
    with LaTOF do begin
      {Si l'on ne gère pas les établissement ...}
      if not VH^.EtablisCpta  then begin
        {... on affiche l'établissement par défaut}
        SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
        {... on désactive la zone}
        SetControlEnabled(ComboEtab.Name, False);
      end

      {On gère l'établisement, donc ...}
      else begin
        {... On commence par regarder les restrictions utilisateur}
        PositionneEtabUser(ComboEtab);
        {... s'il n'y a pas de restrictions, on reprend le paramSoc
         JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
        if GetControlText(ComboEtab.Name) = '' then begin
          {... on affiche l'établissement par défaut
          SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
          {... on active la zone
          SetControlEnabled(ComboEtab.Name, True);
        end;}
      end;
    end;
end;

{---------------------------------------------------------------------------------------}
procedure ControlEtab(LaTOF : TOF; ComboEtab : TControl);
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(ComboEtab) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;

  with LaTOF do begin
    Eta := EtabForce;
    {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
    if (Eta <> '') and (Eta <> GetControlText(ComboEtab.Name)) then begin
      {... on affiche l'établissement des restrictions}
      SetControlText(ComboEtab.Name, Eta);
      {... on désactive la zone}
      SetControlEnabled(ComboEtab.Name, False);
    end;
  end;
end;

procedure TOF_ANAGENETOBV.OnChangeAxe(Sender: Tobject);
var
  stAxe  : string;
  lTFB   : TFichierBase;
  CBAxe : THValComboBox;
begin
  CBAxe := THValComboBox(Getcontrol('Y_AXE', true));
  if CBAxe <> nil then
  begin
    stAxe := copy( CBAxe.Value, 2, 1 ) ; ;
    SetControlText('S_SECTION','');
    SetControlText('S_SECTION_','');
    if stAxe = '1' then
    begin
      SetControlProperty('S_SECTION','DataType','TZSECTION');
      SetControlProperty('S_SECTION_','DataType','TZSECTION');
    end else
    begin
      SetControlProperty('S_SECTION','DataType','TZSECTION' + stAxe);
      SetControlProperty('S_SECTION_','DataType','TZSECTION' + stAxe);
    end;
    lTFB := AxeToFb('A' + stAxe);
    SetControlProperty('S_SECTION','MaxLength',VH^.Cpta[lTFB].Lg);
    SetControlProperty('S_SECTION_','MaxLength',VH^.Cpta[lTFB].Lg);
  end;
end;

initialization
  RegisterClasses([TOF_CPGENERECROISEAXE, TOF_CROISEAXECUBE, TOF_CROISEAXESTAT, TOF_CROISEAXEETAT,
    TOF_COMPTACUBE, TOF_ECRCUBE, TOF_ANACUBE,
      TOF_COMPTATOBV, TOF_ECRTOBV, TOF_ANATOBV, TOF_ANAGENETOBV,
      TOF_BALAGEECUBE
      ]);

end.

