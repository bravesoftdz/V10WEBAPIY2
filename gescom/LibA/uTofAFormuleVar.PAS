{***********UNITE*************************************************
Auteur  ...... : pascal.boudon@laposte.net
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVFORMULE ()
Mots clfs ... : TOF;uTofAFormuleVar
*****************************************************************}
unit uTofAFormuleVar;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MainEagl,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF, Utob, HTB97, windows, messages, extctrls, dicobtp, Graphics, vierge;

type

  TOF_AFormuleVar = class(TOF)
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    FormuleOk, FormuleBloque: Boolean;
    Action: TActionFiche;
    Clause_where: string;
    TobFormule: Tob;
    TobFormuledetail: Tob;
    TobInsert: Tob;
    QFormule: Tquery;
    FrappeClavier, OldFrappeClavier: string;
    AVF_QTETARIFFOR: TRadioButton;
    AVF_QTETARIFFINALE: TRadioButton;
    AVF_QTETARIFPRINC: TRadioButton;
    LBFRAPPECLAVIER: THlabel;
    BtnAjouteUnite: TToolbarButton97;
    BtnRetrancheUnite: TToolbarButton97;
    BTNPAROUV: TToolbarButton97;
    BTNPARFER: TToolbarButton97;
    BTNPOINTVIRGULE: TToolbarButton97;
    Btn1: TToolbarButton97;
    Btn2: TToolbarButton97;
    Btn3: TToolbarButton97;
    Btn4: TToolbarButton97;
    Btn5: TToolbarButton97;
    Btn6: TToolbarButton97;
    Btn7: TToolbarButton97;
    Btn8: TToolbarButton97;
    Btn9: TToolbarButton97;
    Btn0: TToolbarButton97;
    BEFF: TToolbarButton97;
    BEFFMESURE: TToolbarButton97;
    BTNDECIMALSEPARATOR: TToolbarButton97;
    BTNPLUS: TToolbarButton97;
    BTNMOINS: TToolbarButton97;
    BTNFOIS: TToolbarButton97;
    BTNDIV: TToolbarButton97;
    bValider: TToolbarButton97;
    bFerme: TToolbarButton97;
    Binsert: TToolbarButton97;
    Bdelete: TToolbarButton97;
    Bdefaire: TToolbarButton97;
    BVerification: TToolbarButton97;
    Bduplication: TToolbarButton97;
    MonTimer: TTimer;
    GridUnite: THGrid;
    Griddetail: THGrid;
    MemoFormule: TMemo;
    procedure RempliListeUnite;
    procedure BtnAjouteUniteClick(sender: Tobject);
    procedure BtnRetrancheUniteClick(sender: Tobject);
    procedure BTNPAROUVClick(sender: Tobject);
    procedure BTNPARFERClick(sender: Tobject);
    procedure Btn1Click(sender: Tobject);
    procedure Btn2Click(sender: Tobject);
    procedure Btn3Click(sender: Tobject);
    procedure Btn4Click(sender: Tobject);
    procedure Btn5Click(sender: Tobject);
    procedure Btn6Click(sender: Tobject);
    procedure Btn7Click(sender: Tobject);
    procedure Btn8Click(sender: Tobject);
    procedure Btn9Click(sender: Tobject);
    procedure Btn0Click(sender: Tobject);
    procedure BEFFClick(sender: Tobject);
    procedure BEFFMESUREClick(sender: Tobject);
    procedure BTNDECIMALSEPARATORClick(sender: Tobject);
    procedure BTNPLUSClick(sender: Tobject);
    procedure BTNMOINSClick(sender: Tobject);
    procedure BTNFOISClick(sender: Tobject);
    procedure BTNDIVClick(sender: Tobject);
    procedure bValiderClick(sender: Tobject);
    procedure bFermeClick(sender: Tobject);
    procedure BinsertClick(sender: Tobject);
    procedure BDeleteClick(sender: Tobject);
    procedure BdefaireClick(sender: Tobject);
    procedure BVerificationClick(sender: Tobject);
    procedure BDuplicationClick(sender: Tobject);
    procedure RadioBouttonClick(sender: Tobject);
    procedure GridUniteKeyPress(Sender: TObject; var Key: Char);
    procedure MemoFormuleKeyPress(Sender: TObject; var Key: Char);
    procedure GridUniteDblClick(Sender: TObject);
    procedure VideFrappeClavier(Sender: Tobject);
    procedure EcritFormuleLib;
    procedure EcritFormuleDetail;    
    procedure VideChamps;
    procedure ChargeChamps;
    procedure MontreBouton;
    procedure FabriqueLesTobs;
    procedure FabriqueLaTobDetail;
    function VerifUnePaireCrochets(StUnite: string): boolean;
    function VerifContenuCrochets(St: string): integer;
    function VerifParenthese(St: string): integer;
    function VerifOperateur(St: string): integer;
    function TrouveUneLigne(formule, unite: string; Rang: integer): integer;
    procedure RempliUneLigne(StUnite, StLibelle: string);
    procedure Remplilagrille(StUnite: string);
    procedure FabriqueGrille;
    function VerifieFormule(FormuleTarifaire: boolean): integer;
    function VerifAuMoinsUnOperateur(St: string): integer;
    function VerifFormulecompteur(St: string): integer;
    function verificationGrille: integer;
    procedure MajDetailFormule;
    function OkPourSupp: boolean;
    procedure Enregistre(EnSilence: boolean);
    function VerifVide(St: string): integer;
    function OkFormuleTarifaire: boolean;
    function OkPremiereQte: boolean;
    function TrouvePremiereUnite: string;
    procedure BlocageChamps;
  end;

const
  TexteMessage: array[1..2] of string = (
    {1}'Confirmez-vous la Suppression de cette formule ? ',
    {2}'Vérification effectuée avec succès !'
    );

  TextedesErreurs: array[1..22] of string = (
    {1}'il manque un ].',
    {2}'il manque un [.',
    {3}'Le code de la formule est obligatoire.',
    {4}'Le libellé de la formule est obligatoire.',
    {5}'Formule Tarifaire : ',
    {6}'Seules les valeurs O,N sont acceptées dans la grille.',
    {7}'Une Unité n''est pas reconnu dans la formule.',
    {8}'Il manque une ( .',
    {9}'Il manque une ) .',
    {10}'Il y a une erreur dans les parenthèses.',
    {11}'Une Formule ne peut se terminer par un opérateur! ',
    {12}'Une Formule ne peut être vide.',
    {13}'Il y a une erreur d''opérateur.',
    {14}'Les libellés doivent etre uniques!',
    {15}'Les libellés sont obligatoires !',
    {16}'Cette formule existe déja. ',
    {17}'Une formule de type "compteur" doit obligatoirement commencer par [XXX]-[YYY] ',
    {18}'Suppression impossible : un article est relié à cette formule.',
    {19}'Suppression impossible : une ligne  est reliée à cette formule.',
    {20}'Suppression impossible : une quantité variable  est reliée à cette formule.',
    {21}'Modification impossible:une ligne  est reliée à cette formule.',
    {22}'Modification impossible:une quantité variable est reliée à cette formule.'
    );

function AFLanceFiche_FormuleVariable(cle, Action: string): string; //AB-0503

implementation

procedure TOF_AFormuleVar.BDuplicationClick(sender: Tobject);
begin
  Action := taCreat;
  setcontroltext('AVF_FORMULEVAR', getcontroltext('AVF_FORMULEVAR') + '_');
  MontreBouton;
  BlocageChamps;
end;

function TOF_AFormuleVar.VerifUnePaireCrochets(StUnite: string): boolean;
var i: integer;
begin
  i := 0;
  while ((i < GridUnite.rowCount) and ((uppercase(GridUnite.Cells[0, i]) <> StUnite))) do inc(i);
  result := (i < GridUnite.rowCount);
end;

function TOF_AFormuleVar.VerifContenuCrochets(St: string): integer;
var StUnite: string;
  PosOuv, PosFer: integer;
begin
  result := 0;
  while (pos('[', st) > 0) and (result = 0) do
  begin
    PosOuv := pos('[', st);
    PosFer := pos(']', st);
    if PosFer > 0 then
    begin
      StUnite := copy(St, PosOuv + 1, PosFer - posOuv - 1);
      St[PosFer] := '|';
      if not (VerifUnePaireCrochets(uppercase(StUnite))) then result := -7;
    end
    else result := -1; //il manque un crochet fermant
    St[PosOuv] := '|';
  end;
  if result = 0 then if pos(']', st) > 0 then result := -2; //il manque un crochet ouvrant
end;

function TOF_AFormuleVar.VerifParenthese(St: string): integer;
var PosParOuv, PosParFer: integer;
begin
  result := 0;
  while (pos('(', st) > 0) and (result = 0) do
  begin
    PosParOuv := pos('(', st);
    PosParFer := pos(')', st);
    if (PosParFer > 0) then St[PosParFer] := '.' else result := -9;
    if (result = 0) and (PosParFer < PosParOuv) then result := -10;
    St[PosParOuv] := '|';
  end;
  if (result = 0) then
    if (pos(')', st) > 0) then result := -8;
end;

function TOF_AFormuleVar.VerifVide(St: string): integer;
begin
  result := 0;
  if trim(st) = '' then result := -12;
end;

function TOF_AFormuleVar.VerifOperateur(St: string): integer;
var i_ind1: integer;
begin
  result := 0;
  i_ind1 := 1;
  while i_ind1 <= Length(St) - 1 do
  begin
    if (St[i_ind1] = ']') and ((St[i_ind1 + 1] = '[') or (St[i_ind1 + 1] = '('))
      or (St[i_ind1] = ')') and ((St[i_ind1 + 1] = '[') or (St[i_ind1 + 1] = '(')) then
      result := -13;
    inc(i_ind1);
  end;
  if (result = 0) and (trim(st) <> '') then
    if St[length(st)] in ['*', '+', '-', '/'] then result := -11;
end;

function compte(st, stin: string): integer;
begin
  result := 0;
  while pos(stin, st) > 0 do
  begin
    st[pos(stin, st)] := '?'; // caractere quelconque
    inc(result);
  end;
end;

function TOF_AFormuleVar.VerifAuMoinsUnOperateur(St: string): integer;
var i: integer;
begin
  result := 0;
  if compte(St, '[') > 1 then // on verifie seulement si plusieurs valeurs
  begin
    result := -13;
    for i := 1 to length(st) do
      if St[i] in ['*', '+', '-', '/'] then result := 0;
  end;
end;

function TOF_AFormuleVar.VerifFormulecompteur(St: string): integer;
var posCrochet: integer;
begin
  result := 0;
  posCrochet := pos(']', St);
  if (posCrochet > 0) and (posCrochet <> length(st)) then
  begin
    if st[posCrochet + 1] <> '-' then result := -17;
  end;
end;

function TOF_AFormuleVar.verifieFormule(FormuleTarifaire: boolean): integer;
var St: string;
  UneErreur: integer;
begin
  if FormuleTarifaire then
    st := getcontrolText('AVF_FORMULETARIF')
  else
    st := MemoFormule.Lines.Text;
  UneErreur := VerifContenuCrochets(St);
  if UneErreur = 0 then UneErreur := VerifParenthese(St);
  if UneErreur = 0 then UneErreur := VerifOperateur(St);
  if UneErreur = 0 then UneErreur := VerifVide(St);
  if UneErreur = 0 then UneErreur := VerifAuMoinsUnOperateur(St);
  if UneErreur = 0 then
    if getcontrolText('AVF_FORCOMPTEUR') = 'X' then
      UneErreur := VerifFormulecompteur(St);
  result := UneErreur;
end;

procedure TOF_AFormuleVar.BVerificationClick(sender: Tobject);
begin
  if OkPremiereQte and (GetControltext('AVF_MESURE') = '') then SetControlText('AVF_MESURE', TrouvePremiereUnite);
  enregistre(False); // j'enregistre sans sortir avec message d'alerte
  TobFormuleDetail.free; // je detruis la tobFille
  FabriqueLaTobDetail; // je relie les valeurs => base vers la tob
  FabriqueGrille; // je les affiche => text+tob => grille
end;

procedure TOF_AFormuleVar.VideChamps;
begin
  setcontroltext('AVF_FORMULEVAR', '');
  setcontroltext('AVF_FORMVARLIB', '');
  setcontroltext('AVF_FORMVARDEF', '');
  setcontroltext('AVF_FORMVARACTIVE', 'X');
  setcontroltext('AVF_FORCOMPTEUR', '');
  setcontroltext('AVF_MESURE', '');
  setcontroltext('AVF_QTETARIFPRINC', 'X');
  setcontroltext('AVF_QTETARIFFINALE', '');
  setcontroltext('AVF_QTETARIFFOR', '');
  setcontroltext('AVF_FORMULETARIF', '');
  setcontroltext('AVF_QTEARRONDI', '-');
end;

//Ne pas modifier la structure de la formule si elle est déjà utilisée
procedure TOF_AFormuleVar.BlocageChamps;
begin
  FormuleBloque := false;
  if (action = taModif) and
    ExisteSQL('Select AVV_FORMULEVAR  FROM AFORMULEVARQTE WHERE AVV_FORMULEVAR="' + GetControlText('AVF_FORMULEVAR') + '"') then
    FormuleBloque := true;

  SetControlEnabled('AVF_FORMVARLIB', not FormuleBloque);
  SetControlEnabled('AVF_FORMVARDEF', not FormuleBloque);
  SetControlEnabled('AVF_FORCOMPTEUR', not FormuleBloque);
  SetControlEnabled('AVF_MESURE', not FormuleBloque);
  SetControlEnabled('AVF_QTETARIFPRINC', not FormuleBloque);
  SetControlEnabled('AVF_QTETARIFFINALE', not FormuleBloque);
  SetControlEnabled('AVF_QTETARIFFOR', not FormuleBloque);
  SetControlEnabled('AVF_FORMULETARIF', not FormuleBloque);
  MemoFormule.ReadOnly := FormuleBloque;
  Beff.Enabled := not FormuleBloque;
  BeffMesure.Enabled := not FormuleBloque;
  Bdelete.enabled := not FormuleBloque;
  if FormuleBloque then
    LBFRAPPECLAVIER.caption := TextedesErreurs[22]
  else
    LBFRAPPECLAVIER.caption := '';
end;

procedure TOF_AFormuleVar.BinsertClick(sender: Tobject);
begin
  VideChamps;
  Action := taCreat;
  MontreBouton;
  BlocageChamps;
  TobFormuleDetail.free; // je detruis la tobFille
  FabriqueLaTobDetail; // je relie les valeurs
  FabriqueGrille; // je les affiche
end;

function TOF_AFormuleVar.OkPourSupp: boolean;
var erreur: integer;
begin
  erreur := 0;
  if ExisteSQL('Select GA_FORMULEVAR  FROM ARTICLE WHERE GA_FORMULEVAR="' + GetControlText('AVF_FORMULEVAR') + '"') then
    erreur := 18 else
    if ExisteSQL('Select AVV_FORMULEVAR  FROM AFORMULEVARQTE WHERE AVV_FORMULEVAR="' + GetControlText('AVF_FORMULEVAR') + '"') then
    erreur := 20 else
    if ExisteSQL('Select GL_FORMULEVAR  FROM LIGNE WHERE GL_FORMULEVAR="' + GetControlText('AVF_FORMULEVAR') + '"') then
    erreur := 19;

  result := erreur = 0;
  if not result then pgiBoxAF(TextedesErreurs[erreur], Ecran.Caption);
end;

procedure TOF_AFormuleVar.BDeleteClick(sender: Tobject);
begin
  if OkPourSupp then
  begin
    if pgiAskCancelAf(TexteMessage[1], ecran.Caption) <> mrYes then Exit;
    ExecuteSQL('DELETE FROM AFORMULEVAR WHERE AVF_FORMULEVAR="' + GetControlText('AVF_FORMULEVAR') + '"');
    ExecuteSQL('DELETE FROM AFORMULEVARDET WHERE AVD_FORMULEVAR="' + GetControlText('AVF_FORMULEVAR') + '"');
    BinsertClick(nil);
  end;
end;

procedure TOF_AFormuleVar.BdefaireClick(sender: Tobject);
begin
  ChargeChamps;
end;

function TOF_AFormuleVar.TrouveUneLigne(formule, unite: string; Rang: integer): integer;
var i: integer;
begin
  // lors de l'initialisation je vais chercher dans les fichiers les valeur des aformulevardets
  // si rien dans TobFormuleDetail alors renvoie -1
  result := -1;

  if (Rang < tobFormuledetail.Detail.Count) then // je vais essayer de trouver la meme unite au meme rang
  begin
    if (tobFormuledetail.Detail[Rang].GetValue('AVD_FORMULEVAR') = formule) and
      (tobFormuledetail.Detail[Rang].GetValue('AVD_MESURE') = unite) then result := Rang;
  end;
  i := 0;
  // au pire si au meme rang je ne trouve pas de valeur je vais voir autre par.
  while ((result < 0) and (i < tobFormuledetail.Detail.Count)) do
  begin
    if (tobFormuledetail.Detail[i].GetValue('AVD_FORMULEVAR') = formule) and
      (tobFormuledetail.Detail[i].GetValue('AVD_MESURE') = unite) then result := i;
    inc(i);
  end;
end;

function PositionneBooleenGridToBase(st: string): string;
begin
  st := uppercase(st);
  result := '-';
  if (st = 'O') or (st = 'OUI') or (st = 'X') then result := 'X';
end;

function PositionneBooleenBaseToGrid(st: string): string;
begin
  st := uppercase(st);
  result := 'Non';
  if (st = 'X') then result := 'Oui';
end;

procedure TOF_AFormuleVar.RempliUneLigne(StUnite, StLibelle: string);
var i: integer;
begin
  (* Je considere que la formule a la priorité sur aformlevardet c'est pourquoi je donne priorité a
  l'occurence de l'unite dans la formule plutot qu'a son rang dans aformulevardet
  Fonctionnellement c'est plus ergonomique => si user deplace une unité dans formule ses saisies restent valides.
  *)
  GridDetail.Cells[0, Griddetail.RowCount - 1] := StUnite;
  GridDetail.Cells[1, Griddetail.RowCount - 1] := StLibelle;
  i := TrouveUneLigne(GetControlText('AVF_FORMULEVAR'), StUnite, Griddetail.RowCount - 2); // indice commence a Zero et pas a 2 comme la grille
  if i >= 0 then
  begin // au chargement je mets ma Tob dans la grille ouille ! ca fait mal
    GridDetail.Cells[1, Griddetail.RowCount - 1] := tobFormuledetail.Detail[i].GetValue('AVD_LIBMESURE');
    GridDetail.Cells[2, Griddetail.RowCount - 1] := tobFormuledetail.Detail[i].GetValue('AVD_LIBMESURECOURT');
    GridDetail.Cells[3, Griddetail.RowCount - 1] := tobFormuledetail.Detail[i].GetValue('AVD_VALDEF');
    GridDetail.Cells[4, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_LPVIS'));
    GridDetail.Cells[5, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_LPENABLE'));
    GridDetail.Cells[6, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_ACTVIS'));
    GridDetail.Cells[7, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_ACTENABLE'));
    GridDetail.Cells[8, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_PLAVIS'));
    GridDetail.Cells[9, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_PLAENABLE'));
  end else
  begin
    GridDetail.Cells[2, Griddetail.RowCount - 1] := '';
    GridDetail.Cells[3, Griddetail.RowCount - 1] := '0';
    GridDetail.Cells[4, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid('X');
    GridDetail.Cells[5, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid('X');
    GridDetail.Cells[6, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid('X');
    GridDetail.Cells[7, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid('X');
    GridDetail.Cells[8, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid('X');
    GridDetail.Cells[9, Griddetail.RowCount - 1] := PositionneBooleenBaseToGrid('X');
  end;
  Griddetail.RowCount := Griddetail.RowCount + 1;
end;

procedure TOF_AFormuleVar.Remplilagrille(StUnite: string);
var i: integer;
begin
  i := 0;
  while ((i < GridUnite.rowCount) and (uppercase(GridUnite.Cells[0, i]) <> StUnite)) do inc(i);
  if i < GridUnite.rowCount then
    RempliUneLigne(GridUnite.Cells[0, i], GridUnite.Cells[1, i]);
end;

procedure TOF_AFormuleVar.FabriqueGrille;
var st, StUnite: string;
  i, PosOuv, PosFer: integer;
begin
  Griddetail.RowCount := 2;
  Griddetail.ColEditables[0] := False;
  for i := 0 to 2 do Griddetail.cells[i, 1] := ''; //initialisation sinon ancienne valeur pour insert
  for i := 3 to Griddetail.ColCount - 1 do Griddetail.cells[i, 1] := 'Oui'; //initialisation sinon ancienne valeur pour insert

  st := MemoFormule.Lines.CommaText;
  while (pos('[', st) > 0) do
  begin
    PosOuv := pos('[', st);
    PosFer := pos(']', st);
    if PosFer > 0 then
    begin
      StUnite := copy(St, PosOuv + 1, PosFer - posOuv - 1);
      St[PosFer] := '|';
      Remplilagrille(uppercase(StUnite));
    end;
    St[PosOuv] := '|';
  end;
  // une ligne est rajoutée en trop lors du dernier passage mais si on ne passe jamais (formule vide) il ne faut pas l'enlever
  if Griddetail.RowCount > 2 then Griddetail.RowCount := Griddetail.RowCount - 1;
end;

procedure TOF_AFormuleVar.EcritFormuleLib;
var i: integer;
  st: string;
begin
  st := MemoFormule.text;
  i := 0;
  while (i < GridUnite.rowCount) do
  begin
    St := StringReplace(St, '[' + GridUnite.Cells[0, i] + ']', ' ' + GridUnite.Cells[1, i] + ' ', [rfReplaceAll, rfIgnoreCase]);
    inc(i);
  end;
  setControlText('AFE_FORMULELIB', st);
end;

procedure TOF_AFormuleVar.EcritFormuleDetail;
var i: integer;
  st: string;
begin
  st := MemoFormule.text;
  for i := 0 to Griddetail.rowcount-1 do
  begin
    St := FindEtReplace(St,'['+GridDetail.Cells[0, i]+']', '['+ GridDetail.Cells[1,i]+']', false);
  end;
  SetControlText('AFE_FORMULELIB', st);
end;

procedure TOF_AFormuleVar.BtnAjouteUniteClick(sender: Tobject);
begin
  if FormuleBloque then exit;
  MemoFormule.Seltext := '[' + GridUnite.Cells[0, GridUnite.Row] + ']';
  EcritFormuleLib;
end;

procedure TOF_AFormuleVar.GridUniteDblClick(Sender: TObject);
begin
  if FormuleBloque then exit;
  MemoFormule.Seltext := '[' + GridUnite.Cells[0, GridUnite.Row] + ']';
  EcritFormuleLib;
end;

procedure TOF_AFormuleVar.BtnRetrancheUniteClick(sender: Tobject);
var St: string;
begin
  if FormuleBloque then exit;
  St := MemoFormule.text;
  St := StringReplace(St, '[' + GridUnite.Cells[0, GridUnite.Row] + ']', '', [rfReplaceAll, rfIgnoreCase]);
  MemoFormule.text := St;
  EcritFormuleLib;
end;

procedure TOF_AFormuleVar.BTNPAROUVClick(sender: Tobject);
begin
  MemoFormule.seltext := '(';
end;

procedure TOF_AFormuleVar.BTNPARFERClick(sender: Tobject);
begin
  MemoFormule.seltext := ')';
end;

procedure TOF_AFormuleVar.Btn1Click(sender: Tobject);
begin
  MemoFormule.Seltext := '1';
end;

procedure TOF_AFormuleVar.Btn2Click(sender: Tobject);
begin
  MemoFormule.Seltext := '2';
end;

procedure TOF_AFormuleVar.Btn3Click(sender: Tobject);
begin
  MemoFormule.Seltext := '3';
end;

procedure TOF_AFormuleVar.Btn4Click(sender: Tobject);
begin
  MemoFormule.Seltext := '4';
end;

procedure TOF_AFormuleVar.Btn5Click(sender: Tobject);
begin
  MemoFormule.Seltext := '5';
end;

procedure TOF_AFormuleVar.Btn6Click(sender: Tobject);
begin
  MemoFormule.Seltext := '6';
end;

procedure TOF_AFormuleVar.Btn7Click(sender: Tobject);
begin
  MemoFormule.Seltext := '7';
end;

procedure TOF_AFormuleVar.Btn8Click(sender: Tobject);
begin
  MemoFormule.Seltext := '8';
end;

procedure TOF_AFormuleVar.Btn9Click(sender: Tobject);
begin
  MemoFormule.Seltext := '9';
end;

procedure TOF_AFormuleVar.Btn0Click(sender: Tobject);
begin
  MemoFormule.Seltext := '0';
end;

procedure TOF_AFormuleVar.BTNDECIMALSEPARATORClick(sender: Tobject);
begin
  MemoFormule.Seltext := DecimalSeparator;
end;

procedure TOF_AFormuleVar.BTNPLUSClick(sender: Tobject);
begin
  MemoFormule.Seltext := '+';
end;

procedure TOF_AFormuleVar.BTNMOINSClick(sender: Tobject);
begin
  MemoFormule.Seltext := '-';
end;

procedure TOF_AFormuleVar.BTNFOISClick(sender: Tobject);
begin
  MemoFormule.Seltext := '*';
end;

procedure TOF_AFormuleVar.BTNDIVClick(sender: Tobject);
begin
  MemoFormule.Seltext := '/';
end;

procedure TOF_AFormuleVar.BEFFClick(sender: Tobject);
begin
  MemoFormule.text := '';
  setControlText('AFE_FORMULELIB', '');
end;

procedure TOF_AFormuleVar.BEFFMESUREClick(sender: Tobject);
begin
  SetControlText('AVF_FORMULETARIF', '');
end;

procedure TOF_AFormuleVar.MemoFormuleKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in [#22, #3, '(', '.', ',', ')', ' ', ';', #8, '0'..'9', '*', '+', '-', '/']) then key := #0;
end;

procedure TOF_AFormuleVar.GridUniteKeyPress(Sender: TObject; var Key: Char);
var i: integer;
begin
  i := 0;
  if FormuleBloque then Exit;
  FrappeClavier := Uppercase(FrappeClavier + Key);
  while (i < GridUnite.RowCount) and (copy(uppercase(GridUnite.Cells[0, i]), 1, length(FrappeClavier)) <> FrappeClavier) do i := i + 1;
  if i >= GridUnite.RowCount then i := 0;
  GridUnite.GotoRow(i);
  LBFRAPPECLAVIER.caption := FrappeClavier;
end;

procedure TOF_AFormuleVar.MontreBouton;
begin
  Bdelete.Enabled := Action = taModif;
  Binsert.Enabled := Action = taModif;
  Bduplication.Enabled := Action = taModif;
  SetControlEnabled('AVF_FORMULEVAR', Action = taCreat);
end;

procedure TOF_AFormuleVar.FabriqueLaTobDetail;
var stDetail: string;
  QFormuleDetail: Tquery;
begin
  TobFormuledetail := TOB.Create('Mes details de Formule', nil, -1);
  try
    stDetail := 'SELECT AVD_FORMULEVAR, AVD_RANG, AVD_MESURE,AVD_LIBMESURE, AVD_LIBMESURECOURT,';
    stDetail := stDetail + ' AVD_VALDEF,AVD_LPVIS, AVD_LPENABLE, AVD_ACTVIS, AVD_ACTENABLE,AVD_PLAVIS,';
    stDetail := stDetail + ' AVD_PLAENABLE FROM AFORMULEVARDET left join  AFORMULEVAR on AVF_FORMULEVAR=AVD_FORMULEVAR ';
    stDetail := stDetail + ' WHERE AVD_FORMULEVAR="' + getcontroltext('AVF_FORMULEVAR') + '"';
    QFormuleDetail := nil;
    QFormuleDetail := OpenSQL(stDetail, TRUE);
    TobFormuledetail.LoadDetailDB('AFORMULEVARDET', '', '', QFormuleDetail, false);
  finally
    Ferme(QFormuleDetail);
  end; // Try

end;

procedure TOF_AFormuleVar.FabriqueLesTobs;
var StQformule: string;
begin
  TobFormule := TOB.Create('Ma Formule', nil, -1);
  if Action = taCreat then
  begin
    TobInsert := TOB.Create('AFORMULEVAR', TobFormule, -1);
  end else
  begin
    try
      StQformule := 'SELECT AVF_FORMULEVAR,AVF_FORMULETARIF,AVF_QTETARIFFOR,AVF_QTETARIFFINALE,';
      StQformule := StQformule + 'AVF_QTETARIFPRINC,AVF_MESURE,AVF_FORMVARACTIVE,AVF_FORCOMPTEUR,';
      StQformule := StQformule + 'AVF_FORMVARDEF,AVF_FORMVARLIB,AVF_QTEARRONDI ';
      StQformule := StQformule + ' FROM AFORMULEVAR  ' + clause_where; //'FROM AFORMULEVAR  WHERE AVF_FORMULEVAR="'+getcontroltext('AVF_FORMULEVAR')+'"' ;
      QFormule := nil;
      QFormule := OpenSQL(StQformule, TRUE);
      TobFormule.LoadDetailDB('AFORMULEVAR', '', '', QFormule, false);
      ChargeChamps;
      FabriqueLaTobDetail;
    finally
      Ferme(QFormule);
    end; // Try
  end; // if TaCrea

end;

procedure TOF_AFormuleVar.ChargeChamps;
begin
  if TobFormule.Detail.count > 0 then
  begin
    setcontroltext('AVF_FORMULEVAR', TobFormule.Detail[0].GetValue('AVF_FORMULEVAR'));
    setcontroltext('AVF_FORMVARLIB', TobFormule.Detail[0].GetValue('AVF_FORMVARLIB'));
    setcontroltext('AVF_FORMVARDEF', TobFormule.Detail[0].GetValue('AVF_FORMVARDEF'));
    SetControlChecked('AVF_FORMVARACTIVE', TobFormule.Detail[0].GetValue('AVF_FORMVARACTIVE') = 'X');
    SetControlChecked('AVF_FORCOMPTEUR', TobFormule.Detail[0].GetValue('AVF_FORCOMPTEUR') = 'X');
    TobFormule.Detail[0].GetValue('AVF_MESURE');
    setcontroltext('AVF_MESURE', TobFormule.Detail[0].GetValue('AVF_MESURE'));
    SetControlChecked('AVF_QTETARIFPRINC', TobFormule.Detail[0].GetValue('AVF_QTETARIFPRINC') = 'X');
    SetControlChecked('AVF_QTETARIFFINALE', TobFormule.Detail[0].GetValue('AVF_QTETARIFFINALE') = 'X');
    SetControlChecked('AVF_QTETARIFFOR', TobFormule.Detail[0].GetValue('AVF_QTETARIFFOR') = 'X');
    setcontroltext('AVF_FORMULETARIF', TobFormule.Detail[0].GetValue('AVF_FORMULETARIF'));
    SetControlChecked('AVF_QTEARRONDI', TobFormule.Detail[0].GetValue('AVF_QTEARRONDI') = 'X');
  end; // if Count
end;

procedure TOF_AFormuleVar.OnLoad;
begin
  inherited;
  FormuleOk := True;
  MontreBouton;
  FrappeClavier := '';
  OldFrappeClavier := '';
  RempliListeUnite;
  MonTimer := TTimer.Create(Ecran);
  MonTimer.Interval := 500;
  MonTimer.OnTimer := VideFrappeClavier;
  FabriqueLesTobs;
  verifieFormule(false);
  FabriqueGrille;
  BlocageChamps;
  EcritFormuleDetail;
  if action <> tacreat then
    ecran.caption := ecran.caption + ' : ' + GetControltext('AVF_FORMULEVAR') + ' ' + GetControltext('AVF_FORMVARLIB');
end;

function valeur(st: string): string;
var code: integer;
  f: double;
begin
  val(st, f, code);
  result := floattostr(f);
end;

procedure TOF_AFormuleVar.MajDetailFormule;
var st, stValues, Formule, mesure: string;
  colonne, ligne: integer;
  libMesure, libMesureCourt: string;
begin
  //formule:=TobFormule.Detail[0].GetValue('AVF_FORMULEVAR') ;
  formule := Getcontroltext('AVF_FORMULEVAR');
  st := 'delete from  AFORMULEVARDET where AVD_FORMULEVAR = "' + formule + '"';
  executesql(st);
  st := 'insert into AFORMULEVARDET (AVD_FORMULEVAR, AVD_RANG, AVD_MESURE,AVD_LIBMESURE, AVD_LIBMESURECOURT, AVD_VALDEF,';
  st := st +
    'AVD_LPVIS, AVD_LPENABLE, AVD_ACTVIS, AVD_ACTENABLE,AVD_PLAVIS, AVD_PLAENABLE,AVD_DATECREATION,AVD_DATEMODIF,AVD_CREATEUR,AVD_UTILISATEUR) values (';
  for ligne := 1 to GridDetail.rowcount - 1 do
  begin
    mesure := GridDetail.Cells[0, ligne];
    libMesure := GridDetail.Cells[1, ligne];
    libMesureCourt := GridDetail.Cells[2, ligne];
    stValues := '"' + formule + '",' + inttostr(ligne) + ',"' + mesure + '","' + libMesure + '","' + libMesureCourt + '",' + valeur(GridDetail.Cells[3,
      ligne]);
    for colonne := 4 to 9 do
      stValues := stValues + ',"' + PositionneBooleenGridToBase(GridDetail.Cells[colonne, ligne]) + '"';
    stValues := stValues + ',"' + usdatetime(V_PGI.dateEntree) + '","' + usdatetime(V_PGI.dateEntree) + '","' + V_PGI.User + '","' + V_PGI.User + '"';
    executesql(st + stValues + ')');
  end;
end;

function TOF_AFormuleVar.verificationGrille: integer;

var ligne, ligne2, Colonne: integer;
  erreur: integer;
begin
  erreur := 0;
  //libelle obligatoire ;
  for ligne := 1 to GridDetail.rowcount - 1 do
  begin
    if (trim(GridDetail.Cells[0, ligne]) <> '') and (trim(GridDetail.Cells[1, ligne]) = '') then // libelle ;
      erreur := -15;
  end;
  // unicité des lib
  for ligne := 1 to GridDetail.rowcount - 1 do
  begin
    for ligne2 := 2 to GridDetail.rowcount - 1 do
    begin
      if (uppercase(trim(GridDetail.Cells[1, ligne])) = uppercase(trim(GridDetail.Cells[1, ligne2]))) and (ligne <> Ligne2) then // libelle ;
        erreur := -14;
    end;
  end;
  if erreur = 0 then
  begin
    for ligne := 1 to GridDetail.rowcount - 1 do
      for colonne := 4 to Griddetail.Colcount - 1 do
        if (uppercase(GridDetail.Cells[colonne, ligne]) <> 'O') and
          (uppercase(GridDetail.Cells[colonne, ligne]) <> 'OUI') and
          (uppercase(GridDetail.Cells[colonne, ligne]) <> 'N') and
          (uppercase(GridDetail.Cells[colonne, ligne]) <> 'NON') then erreur := -6;
  end;
  result := erreur;
end;

procedure TOF_AFormuleVar.Enregistre(EnSilence: boolean);
var Erreur: integer;
begin
  if trim(GetControltext('AVF_FORMULEVAR')) = '' then
  begin
    PgiinfoAF(TextedesErreurs[3], Ecran.Caption);
    SetfocusControl('AVF_FORMULEVAR');
    exit;
  end;
  if trim(GetControltext('AVF_FORMVARLIB')) = '' then
  begin
    PgiinfoAF(TextedesErreurs[4], Ecran.Caption);
    SetfocusControl('AVF_FORMVARLIB');
    exit;
  end;

  MajDetailFormule;
  TobFormuleDetail.free; // je detruis la tobFille
  FabriqueLaTobDetail; // je relie les valeurs => base vers la tob
  FabriqueGrille; // je les affiche => text+tob => grille

  if (Action = taCreat) and ExisteSQL('select AVF_FORMULEVAR from AFORMULEVAR  where AVF_FORMULEVAR="' + GetControltext('AVF_FORMULEVAR') + '"') then
  begin
    pgiBoxAF(TextedesErreurs[16], Ecran.Caption);
  end
  else
  begin
    erreur := verifieFormule(false);
    if erreur <> 0 then
    begin
      pgiBoxAF(TextedesErreurs[-erreur], Ecran.Caption);
      formuleOK := false;
      SetFocusControl('AVF_FORMVARDEF');
      Exit;
    end else
      if OkFormuleTarifaire then
    begin
      erreur := verifieFormule(true);
      if erreur <> 0 then
      begin
        pgiBoxAF(TextedesErreurs[5] + TextedesErreurs[-erreur], Ecran.Caption);
        formuleOK := false;
        SetFocusControl('AVF_FORMULETARIF');
        Exit;
      end;
    end;
    erreur := verificationGrille;
    if erreur <> 0 then
    begin
      pgiBoxAF(TextedesErreurs[-erreur], Ecran.Caption);
      formuleOK := false;
      SetfocusControl('GRIDDETAIL');
      Exit;
    end;
    formuleOK := true;
    EcritFormuleDetail;
    if (not Ensilence) and (erreur = 0) then PgiinfoAF(TexteMessage[2], Ecran.Caption);
    if Ensilence and (Erreur = 0) then
    begin
      TobFormule.Detail[0].PutValue('AVF_FORMULEVAR', GetControltext('AVF_FORMULEVAR'));
      TobFormule.Detail[0].PutValue('AVF_FORMVARLIB', GetControltext('AVF_FORMVARLIB'));
      TobFormule.Detail[0].PutValue('AVF_FORMVARDEF', GetControltext('AVF_FORMVARDEF'));
      TobFormule.Detail[0].PutValue('AVF_FORMVARACTIVE', GetControltext('AVF_FORMVARACTIVE'));
      TobFormule.Detail[0].PutValue('AVF_FORCOMPTEUR', GetControltext('AVF_FORCOMPTEUR'));
      TobFormule.Detail[0].PutValue('AVF_MESURE', GetControltext('AVF_MESURE'));
      TobFormule.Detail[0].PutValue('AVF_QTETARIFPRINC', GetControltext('AVF_QTETARIFPRINC'));
      TobFormule.Detail[0].PutValue('AVF_QTETARIFFINALE', GetControltext('AVF_QTETARIFFINALE'));
      TobFormule.Detail[0].PutValue('AVF_QTETARIFFOR', GetControltext('AVF_QTETARIFFOR'));
      TobFormule.Detail[0].PutValue('AVF_FORMULETARIF', GetControltext('AVF_FORMULETARIF'));
      TobFormule.Detail[0].PutValue('AVF_QTEARRONDI', GetControltext('AVF_QTEARRONDI'));
      if Action = taModif then TobFormule.UpdateDB(False)
      else
      begin
        TobFormule.insertDB(nil, false);
        TFVierge(Ecran).retour := GetControltext('AVF_FORMULEVAR'); //AB-0503
      end;
      MajDetailFormule;
      Action := taModif;
      MontreBouton;
    end;
  end;
end;

procedure TOF_AFormuleVar.bValiderClick(sender: Tobject);
begin
  enregistre(true);
end;

procedure TOF_AFormuleVar.bFermeClick(sender: Tobject);
begin
  FormuleOk := True;
end;

procedure TOF_AFormuleVar.VideFrappeClavier(Sender: Tobject);
begin
  if FormuleBloque then exit;
  if (OldFrappeClavier = FrappeClavier) then FrappeClavier := '';
  OldFrappeClavier := FrappeClavier;
  LBFRAPPECLAVIER.caption := FrappeClavier;
end;

procedure TOF_AFormuleVar.RempliListeUnite;
var StRequete: string;
  TobUnite: TOB;
  MaRequete: TQuery;
  i: integer;

begin
  StRequete := 'SELECT GME_MESURE,GME_LIBELLE FROM MEA';
  MaRequete := nil;
  TobUnite := TOB.Create('', nil, -1);
  i := 0;
  try
    MaRequete := OpenSQL(StRequete, TRUE);
    TobUnite.LoadDetailDB('', '', '', MaRequete, false);
    GridUnite.RowCount := TobUnite.Detail.count;
    while (i < TobUnite.Detail.count) do
    begin
      GridUnite.Cells[0, i] := TobUnite.Detail[i].GetValue('GME_MESURE');
      GridUnite.Cells[1, i] := TobUnite.Detail[i].GetValue('GME_LIBELLE');
      i := i + 1;
    end;
  finally
    Ferme(MaRequete);
    TobUnite.Free;
  end;
end;

function TOF_AFormuleVar.OkFormuleTarifaire: boolean;
begin
  result := GetControltext('AVF_QTETARIFFOR') = 'X';
end;

function TOF_AFormuleVar.OkPremiereQte: boolean;
begin
  result := GetControltext('AVF_QTETARIFPRINC') = 'X';
end;

function TOF_AFormuleVar.TrouvePremiereUnite: string;
var StUnite, St: string;
  PosOuv, PosFer: integer;
begin
  st := MemoFormule.Lines.CommaText;
  PosOuv := pos('[', st);
  PosFer := pos(']', st);
  StUnite := copy(St, PosOuv + 1, PosFer - posOuv - 1);
  if VerifUnePaireCrochets(uppercase(StUnite)) then result := Stunite else result := '';
end;

procedure TOF_AFormuleVar.RadioBouttonClick(sender: Tobject);
var OkMontre: boolean;
begin
  OkMontre := OkFormuleTarifaire;
  setcontrolEnabled('AVF_FORMULETARIF', OKMontre);
  setcontrolEnabled('BEFFMESURE', OKMontre);
  if okmontre and (GetControltext('AVF_FORMULETARIF') = '') then setControltext('AVF_FORMULETARIF', MemoFormule.Lines.CommaText);
  if OkPremiereQte and (GetControltext('AVF_MESURE') = '') then SetControlText('AVF_MESURE', TrouvePremiereUnite);

  if THedit(getcontrol('AVF_FORMULETARIF')).enabled then
    THedit(getcontrol('AVF_FORMULETARIF')).color := ClWindow
  else
  begin
    THedit(getcontrol('AVF_FORMULETARIF')).color := clBtnFace;
    SetcontrolText('AVF_FORMULETARIF', '');
  end;
end;

procedure TOF_AFormuleVar.OnArgument(S: string);
var Critere, Champ, valeur: string;
  X: integer;
begin
  inherited;
  Action := taCreat;
  Critere := (Trim(ReadTokenSt(S)));
  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos('=', Critere);
      if x <> 0 then
      begin
        Champ := copy(Critere, 1, X - 1);
        Valeur := Copy(Critere, X + 1, length(Critere) - X);
      end;
      if Champ = 'AVF_FORMULEVAR' then Clause_where := 'WHERE AVF_FORMULEVAR="' + Valeur + '"';
      if (Champ = 'ACTION') and (valeur = 'CREATION') then Action := taCreat;
      if (Champ = 'ACTION') and (valeur = 'MODIFICATION') then Action := taModif;
      if (Champ = 'ACTION') and (valeur = 'CONSULTATION') then Action := taConsult;
    end;
    Critere := (Trim(ReadTokenSt(S)));
  end;

  BtnAjouteUnite := TToolbarButton97(Getcontrol('BTNAJOUTEUNITE'));
  BtnRetrancheUnite := TToolbarButton97(Getcontrol('BTNRETRANCHEUNITE'));
  BTNPAROUV := TToolbarButton97(Getcontrol('BTNPAROUV'));
  BTNPARFER := TToolbarButton97(Getcontrol('BTNPARFER'));
  Btn1 := TToolbarButton97(Getcontrol('BTN1'));
  Btn2 := TToolbarButton97(Getcontrol('BTN2'));
  Btn3 := TToolbarButton97(Getcontrol('BTN3'));
  Btn4 := TToolbarButton97(Getcontrol('BTN4'));
  Btn5 := TToolbarButton97(Getcontrol('BTN5'));
  Btn6 := TToolbarButton97(Getcontrol('BTN6'));
  Btn7 := TToolbarButton97(Getcontrol('BTN7'));
  Btn8 := TToolbarButton97(Getcontrol('BTN8'));
  Btn9 := TToolbarButton97(Getcontrol('BTN9'));
  Btn0 := TToolbarButton97(Getcontrol('BTN0'));
  BDelete := TToolbarButton97(Getcontrol('BDELETE'));
  Binsert := TToolbarButton97(Getcontrol('BINSERT'));
  Bdefaire := TToolbarButton97(Getcontrol('BDEFAIRE'));
  Binsert.OnClick := BinsertClick;
  BDelete.OnClick := BDeleteClick;
  Bdefaire.OnClick := BdefaireClick;

  BVerification := TToolbarButton97(Getcontrol('BVERIFICATION'));
  BVerification.OnClick := BVerificationClick;

  BTNDECIMALSEPARATOR := TToolbarButton97(Getcontrol('BTNDECIMALSEPARATOR'));
  BTNPOINTVIRGULE := TToolbarButton97(Getcontrol('BTNPOINTVIRGULE'));

  BTNPLUS := TToolbarButton97(Getcontrol('BTNPLUS'));
  BTNMOINS := TToolbarButton97(Getcontrol('BTNMOINS'));
  BTNFOIS := TToolbarButton97(Getcontrol('BTNFOIS'));
  BTNDIV := TToolbarButton97(Getcontrol('BTNDIV'));
  BEFF := TToolbarButton97(Getcontrol('BEFF'));
  BEFFMESURE := TToolbarButton97(Getcontrol('BEFFMESURE'));
  bValider := TToolbarButton97(GetControl('bValider'));
  bFerme := TToolbarButton97(GetControl('bFerme'));
  BDuplication := TToolbarButton97(Getcontrol('BDUPLICATION'));
  BDuplication.OnClick := BDuplicationClick;
  BtnAjouteUnite.OnClick := BtnAjouteUniteClick;
  BtnRetrancheUnite.OnClick := BtnRetrancheUniteClick;

  BTNPAROUV.OnClick := BTNPAROUVClick;
  BTNPARFER.OnClick := BTNPARFERClick;

  BEFF.OnClick := BEFFClick;
  BEFFMESURE.OnClick := BEFFMESUREClick;

  Btn1.OnClick := Btn1Click;
  Btn2.OnClick := Btn2Click;
  Btn3.OnClick := Btn3Click;
  Btn4.OnClick := Btn4Click;
  Btn5.OnClick := Btn5Click;
  Btn6.OnClick := Btn6Click;
  Btn7.OnClick := Btn7Click;
  Btn8.OnClick := Btn8Click;
  Btn9.OnClick := Btn9Click;
  Btn0.OnClick := Btn0Click;
  BTNDECIMALSEPARATOR.OnClick := BTNDECIMALSEPARATORClick;
  BTNPLUS.OnClick := BTNPLUSClick;
  BTNMOINS.OnClick := BTNMOINSClick;
  BTNFOIS.OnClick := BTNFOISClick;
  BTNDIV.OnClick := BTNDIVClick;
  bValider.OnClick := bValiderClick;
  bFerme.OnClick := bFermeClick;
  LBFRAPPECLAVIER := THLabel(GetControl('LBFRAPPECLAVIER'));
  LBFRAPPECLAVIER.caption := '';
  GridUnite := THGrid(getcontrol('GRIDUNITE'));
  GridDetail := THGrid(getcontrol('GRIDDETAIL'));

  // provisoire en attendant que le planning soit développé
  Griddetail.Colcount := 8;
  // Fin du provisoire en attendant que le planning soit développé

  GridUnite.OnKeyPress := GridUniteKeyPress;
  GridUnite.OnDblClick := GridUniteDblClick;
  MemoFormule := Tmemo(getControl('AVF_FORMVARDEF'));

  AVF_QTETARIFFOR := TRadioButton(getcontrol('AVF_QTETARIFFOR'));
  AVF_QTETARIFFINALE := TRadioButton(getcontrol('AVF_QTETARIFFINALE'));
  AVF_QTETARIFPRINC := TRadioButton(getcontrol('AVF_QTETARIFPRINC'));
  AVF_QTETARIFFOR.OnClick := RadioBouttonClick;
  AVF_QTETARIFFINALE.OnClick := RadioBouttonClick;
  AVF_QTETARIFPRINC.OnClick := RadioBouttonClick;

  MemoFormule.OnKeyPress := MemoFormuleKeyPress;

end;

procedure TOF_AFormuleVar.OnClose;
begin
  inherited;
  if formuleOK then
  begin
    TobFormule.free;
    TobFormuleDetail.free;
  end else
    lastError := -1;

  AvertirTable('AFORMULEVAR');
end;

procedure TOF_AFormuleVar.OnDisplay();
begin
  inherited;
end;

procedure TOF_AFormuleVar.OnCancel();
begin
  inherited;
end;

function AFLanceFiche_FormuleVariable(cle, Action: string): string; //AB-0503
begin
  Result := AglLanceFiche('AFF', 'AFORMULEVAR', '', cle, Action);
end;

initialization
  registerclasses([TOF_AFormuleVar]);
end.
