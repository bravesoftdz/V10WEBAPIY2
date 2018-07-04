{ Unité : Source TOF de la fiche liée à la TABLE : CPCIRCUITBAP
--------------------------------------------------------------------------------------
    Version   |   Date | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  02/02/06  JP   Création de l'unité
 7.01.001.011  07/06/06  JP   FQ 18296 : Ajout d'un test d'existence
 8.00.001.018  29/05/07  JP   Mise en place des rôles à la place des groupes pour filtrer les utilisateurs
 8.00.001.018  08/06/07  JP   Demande de Sic de ne pas interdire l'utilisation du même viseur sur plusieurs étapes
 8.00.001.019  13/06/07  JP   A la demande de SIC, on "debloque" tous les contrôles sur les utilisateurs
 8.01.001.021  21/06/07  JP   FQ 20776 : Incrémentation automatique du code
--------------------------------------------------------------------------------------}
unit CPCIRCUITBAP_TOF;

interface

uses
  Controls, Classes, Vierge, 
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
  {$ENDIF}
  Grids, SysUtils, HCtrls, UTob, HEnt1, UTOF, Windows, Graphics, Forms, UObjGen;

type
  TOF_CPCIRCUITBAP = class(TOF)
    FListe   : THGrid;

    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
  private
    Action     : TActionFiche;
    Code       : string;
    tDonnees   : TOB;
    lViseurs   : TStringList;
    lUtilisat  : TStringList;
    ValCell    : string;
    InDelete   : Boolean;
    PeutFermer : Boolean;
    ObjetCode  : TObjCodeCombo; {21/06/07 : FQ 20776}

    function  IsUpdating   : Boolean;
    function  TesteCode    : Boolean;
    function  ValideSaisie : Boolean;
    function  TesteUser(C, R : Byte) : Boolean; {FQ 18296}

    {21/06/07 : FQ 20776 : Initialisation du code circuit en création}
    procedure InitCodeCircuit;
    procedure Detruire;
    procedure Annuler;
    procedure Inserer(ParFleche : Boolean);
    procedure ChargerGrille;
    procedure GereCancel(Actif : Boolean);
    procedure RempliCodeEtLigne;
    procedure RemplirListeViseurs;
    procedure RemplirLigneViseurs(Col, Row : Integer);
    function  ValideLigne(AvecMaj : Boolean; Row: Integer) : Boolean;
  public
    FFermerClick  : TNotifyEvent;
    FFormCanClose : TCloseQueryEvent;

    procedure CodeCircuitOnExit(Sender : TObject);
    procedure FListeEllipsis   (Sender : TObject);
    procedure FListeOnEnter    (Sender : TObject);
    procedure FListeOnExit     (Sender : TObject);
    procedure BAnnuleOnClick   (Sender : TObject);
    procedure BInsertOnClick   (Sender : TObject);
    procedure BDeleteOnClick   (Sender : TObject);
    procedure BFermerOnClick   (Sender : TObject);
    procedure FlisteOnKeyDown  (Sender : TObject; var Key: Word; Shift: TShiftState);
    procedure FormOnKeyDown    (Sender : TObject; var Key: Word; Shift: TShiftState);
    procedure FListeOnRowExit  (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
    procedure FListeOnRowEnter (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
    procedure FListeCellEnter  (Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
    procedure FListeCellExit   (Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
    procedure FormCanCloseQuery(Sender : TObject; var CanClose : Boolean);
    procedure FListeDrawCell   (ACol, ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure FListeOnDbleClick(Sender : TObject);
  end;

procedure CpLanceFiche_CircuitFiche(Code, Action : string);

implementation

uses
  HMsgBox, UProcGen, UtilPGI, HTB97, LookUp, ParamSoc, ULibBonAPayer;

type
  TObjViseur = class
    Col : Integer;
    Row : Integer;
  end;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_CircuitFiche(Code, Action : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPCIRCUITBAP', '', '', Code + ';' + Action);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.Detruire;
{---------------------------------------------------------------------------------------}
begin
  if Action = taConsult then Exit;
  {Pour ne pas exécuter le RowExit et les tests de validation}
  InDelete := True;
  FListe.DeleteRow(Fliste.Row);
  {Il faut regérer la liste des viseurs car les index risquent fort d'être décalés}
  RemplirListeViseurs;
  {Initialise le code et le numéro d'ordre de la ligne}
  RempliCodeEtLigne;
  {On active le bouton Defaire}
  GereCancel(True);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.Inserer(ParFleche: Boolean);
{---------------------------------------------------------------------------------------}
begin
  if Action = taConsult then Exit;
  if (not ParFleche) or (Fliste.Row = FListe.RowCount - 1) then begin
    FListe.InsertRow(Fliste.Row + 1);
    {On active le bouton Defaire}
    GereCancel(True);
    {Il faut regérer la liste des viseurs car les index risquent fort d'être décalés}
    RemplirListeViseurs;
    {Initialise le code et le numéro d'ordre de la ligne}
    RempliCodeEtLigne;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.Annuler;
{---------------------------------------------------------------------------------------}
begin
  Fliste.VidePile(False);
  tDonnees.PutGridDetail(FListe, False, False, 'CCI_NUMEROORDRE;CCI_CIRCUITBAP;CCI_VISEUR1;CCI_VISEUR2;CCI_NBJOUR');
  RemplirListeViseurs;
  GereCancel(False);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.GereCancel(Actif : Boolean);
{---------------------------------------------------------------------------------------}
begin
  SetControlProperty('BDEFAIRE', 'ENABLED', Actif);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCIRCUITBAP.IsUpdating : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := GetControlEnabled('BDEFAIRE');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.ChargerGrille;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if (Action = taCreat) or (Code = '') then Exit;
  tDonnees.ClearDetail;
  Q := OpenSelect('SELECT * FROM CPCIRCUIT WHERE CCI_CIRCUITBAP = "' + Code + '" ORDER BY CCI_NUMEROORDRE');
  try
    tDonnees.LoadDetailDB('', '', '', Q, True);
    tDonnees.PutGridDetail(FListe, False, False, 'CCI_NUMEROORDRE;CCI_CIRCUITBAP;CCI_VISEUR1;CCI_VISEUR2;CCI_NBJOUR');
    SetControlText('CODECIRCUIT', Code);
    if tDonnees.Detail.Count > 0 then
      SetControlText('LIBCIRCUIT', tDonnees.Detail[0].GetString('CCI_LIBELLE'));
    {Mémorise les viseurs pour éviter les doublons}
    RemplirListeViseurs;
    {Désactive le bouton Défaire}
    GereCancel(False);
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.RempliCodeEtLigne;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 1 to Fliste.RowCount - 1 do begin
    FListe.Cells[0, n] := IntToStr(n);
    FListe.CellValues[1, n] := GetControlText('CODECIRCUIT');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.RemplirListeViseurs;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  LibereListe(lViseurs, False);

  for n := 1 to FListe.RowCount - 1 do begin
    RemplirLigneViseurs(2, n);
    RemplirLigneViseurs(3, n);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.RemplirLigneViseurs(Col, Row: Integer);
{---------------------------------------------------------------------------------------}
var
  Vis : TObjViseur;
begin
  Vis := TObjViseur.Create;
  Vis.Col := Col;
  Vis.Row := Row;

  lViseurs.AddObject(FListe.CellValues[Col, Row], Vis);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FlisteOnKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_DELETE : if ssCtrl in Shift then Detruire
                                   else if Fliste.Col > 0 then FListe.Cells[FListe.Col, FListe.Row] := '';
    VK_INSERT : Inserer(False);
    VK_DOWN   : Inserer(True);
    VK_F5     : FlisteEllipsis(Sender);
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FormOnKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F10 : TFVierge(Ecran).BValider.Click;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeOnRowEnter(Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if (Action = taCreat) and (Fliste.CellValues[1, Fliste.Row] = '') then
    {Initialise le code et le numéro d'ordre de la ligne}
    RempliCodeEtLigne;
  if (Action = taCreat) or IsUpdating then
    FListe.Col := 2;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeOnRowExit(Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
{---------------------------------------------------------------------------------------}
begin
  {Je ne sais pas à quoi sert ce booléen, mais, parfois, FListeOnRowExit est exécuté deux fois
   consécutivement, la première fois avec Chg à False et l'autre à True}
  if Chg then Exit;

  if (Action = taConsult) or not IsUpdating or InDelete then begin
    InDelete := False;
    Exit;
  end;

  Cancel := not ValideLigne(True, Ou);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeCellEnter(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  ValCell := FListe.CellValues[ACol, ARow];
  FListe.ElipsisButton := (FListe.Col in [2, 3]);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if ValCell <> FListe.CellValues[ACol, ARow] then begin
    GereCancel(True);
    {07/06/06 : FQ 18296 : on vérifie que le User existe bien dans UTILISAT}
    Cancel := not TesteUser(ACol, ARow);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.BFermerOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if (Action <> taConsult) and IsUpdating then begin
    if HShowMessage('3;' + Ecran.Caption + ';Êtes-vous sûr de vouloir abandonner la saisie ?;Q;YN;N;N;', '', '') = mrYes then begin
      GereCancel(False);
      InDelete := True;
      PeutFermer := True;
      FFermerClick(Sender);
    end
    else
      PeutFermer := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeDrawCell(ACol, ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  s : string;
  R : TRect;
begin
  if ACol in [2, 3] then begin
    s := RechDom('TTUTILISATEUR', FListe.CellValues[ACol, ARow], False);
    if s <> '' then begin
      R := FListe.CellRect(ACol, ARow);
      Canvas.FillRect(R);
      Canvas.TextRect(R, R.Left + 2, R.Top + 2, s);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.BAnnuleOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Annuler;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.BDeleteOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Detruire;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.BInsertOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Inserer(False);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeOnEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if GetControlText('CODECIRCUIT') = '' then begin
    HShowMessage('1;' + Ecran.Caption + ';Veuillez choisir un circuit.;W;O;O;O;', '', '');
    SetFocusControl('CODECIRCUIT');
  end;
end;

{21/06/07 : Dans le cas où l'on sortirait de la grille pour aller dans le libellé
{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeOnExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  c : Boolean;
begin
  if not TesteUser(FListe.Col, FListe.Row) then
    FListe.SetFocus

  else begin
    FListeOnRowExit(FListe, FListe.Row, C, False);
    if C then 
      FListe.SetFocus;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FlisteEllipsis(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {JP 29/05/07 : Mise en place des rôles
  Wh := 'US_GROUPE = "' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') + '" AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)';
  LookupList(Fliste, TraduireMemoire('Liste des viseurs'), 'UTILISAT', 'US_UTILISATEUR', 'US_LIBELLE', Wh, 'US_UTILISATEUR', False, 3170);}
  LookUpUtilisateur(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.CodeCircuitOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if Action = taCreat then begin
    if TesteCode then RempliCodeEtLigne;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  a : Char;
  Q : TQuery;
begin
  inherited;
  Ecran.HelpContext := 7509025;

  Code := ReadTokenSt(S);
  a    := StrToChr(ReadTokenSt(S));
  if Code = '' then a := 'C';

  FListe := THGrid(GetControl('FLISTE'));
  case a of
    'M' : Action := taModif;
    'C' : Action := taCreat;
    else begin
      Action := taConsult;
      Fliste.Options := Fliste.Options - [goEditing];
      SetControlEnabled('LIBCIRCUIT', False);
      SetControlEnabled('BVALIDER'  , False);
      SetControlEnabled('BDEFAIRE'  , False);
      SetControlEnabled('BDELETE'   , False);
      SetControlEnabled('BINSERT'   , False);
      FListe.Enabled := False;
      GereCancel(False);
    end;
  end;

  tDonnees  := TOB.Create('$CPCIRCUIT', nil, -1);
  lViseurs  := TStringList.Create;
  {08/06/06 : FQ 18296 : Chargement des utilisateur}
  lUtilisat := TStringList.Create;
  {JP 29/05/07 : Mise en place des rôles
  Q := OpenSQL('SELECT US_UTILISATEUR FROM UTILISAT WHERE US_GROUPE = "' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') +
               '" AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)', True);}
  Q := OpenSQL('SELECT US_UTILISATEUR FROM UTILISAT WHERE (US_GRPSDELEGUES LIKE "%' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') +
               ';%" OR US_GRPSDELEGUES = "" OR US_GRPSDELEGUES IS NULL OR US_GRPSDELEGUES LIKE "<<%>>") AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)', True);
  try
    while not Q.EOF do begin
      lUtilisat.Add(Q.FindField('US_UTILISATEUR').AsString);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

  Fliste.ColTypes  [4] := 'I';
  FListe.ColEditables[1] := False;
  FListe.OnKeyDown   := FlisteOnKeyDown;
  FListe.OnRowExit   := FListeOnRowExit;
  FListe.OnRowEnter  := FListeOnRowEnter;
  FListe.OnCellEnter := FListeCellEnter;
  FListe.OnCellExit  := FListeCellExit;
  FListe.OnEnter     := FListeOnEnter;
  FListe.OnExit      := FListeOnExit;
  FListe.PostDrawCell  := FListeDrawCell;
  Fliste.OnElipsisClick := FlisteEllipsis;
  Fliste.OnDblClick := FListeOnDbleClick;

  Ecran.OnKeyDown := FormOnKeyDown;
  FFermerClick    := TToolbarButton97(GetControl('BFERME'  )).OnClick;

  TToolbarButton97(GetControl('BDEFAIRE')).OnClick := BAnnuleOnClick;
  TToolbarButton97(GetControl('BDELETE' )).OnClick := BDeleteOnClick;
  TToolbarButton97(GetControl('BINSERT' )).OnClick := BInsertOnClick;
  TToolbarButton97(GetControl('BFERME'  )).OnClick := BFermerOnClick;

  ChargerGrille;

  if Action = taCreat then
    THEdit(GetControl('CODECIRCUIT')).OnExit := CodeCircuitOnExit
  else
    SetControlEnabled('CODECIRCUIT', False);

  InDelete := False;
  if GetControlText('LIBCIRCUIT') <> '' then begin
    Ecran.Caption := 'Cicuit : ' + GetControlText('LIBCIRCUIT') + ' (' + GetControlText('CODECIRCUIT') + ')';
    UpdateCaption(Ecran);
  end;

  FFormCanClose := Ecran.OnCloseQuery;
  Ecran.OnCloseQuery := FormCanCloseQuery;
  PeutFermer := True;

  {21/06/07 : FQ 20776 : Initialisation du code circuit en création}
  InitCodeCircuit;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tDonnees ) then FreeAndNil(tDonnees);
  if Assigned(lViseurs ) then LibereListe(lViseurs, True);
  if Assigned(lUtilisat) then FreeAndNil(lUtilisat);
  if Assigned(ObjetCode) then FreeAndNil(ObjetCode); {21/06/07 : FQ 20776}
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
  C : Boolean;
begin
  {30/05/07 : En Cas de F10, le CellExit n'est pas exécuté}
  if not TesteUser(FListe.Col, FListe.Row) then begin
    FListe.SetFocus;
    Ecran.ModalResult := mrNone;
    Exit;
  end;

  {30/05/07 : En Cas de F10, le RowExit n'est pas exécuté}
  FListeOnRowExit(FListe, FListe.Row, C, False);
  if C then begin
    FListe.SetFocus;
    Ecran.ModalResult := mrNone;
    Exit;
  end;

  if not ValideSaisie then begin
    Ecran.ModalResult := mrNone;
    Exit;
  end;

  inherited;
  GereCancel(False);
  {Maj de la table}
  BeginTrans;
  try
    tDonnees.ClearDetail;
    {On commence par supprimer ce qu'il y a en base, pour gérer le cas où l'on
     aurait supprimer des lignes}
    if Action = taModif then
      ExecuteSQL('DELETE FROM CPCIRCUIT WHERE CCI_CIRCUITBAP = "' + GetControlText('CODECIRCUIT') + '"');

    {Le GetGridDetail me crée une tob virtuelle si je ne travaille pas dans
     l'ordre des champs dans la Table, ce qui est le cas ici}
    for n := 1 to Fliste.RowCount - 1 do begin
      T := TOB.Create('CPCIRCUIT', tDonnees, -1);
      T.PutValue('CCI_NUMEROORDRE', FListe.Cells[0, n]);
      T.PutValue('CCI_CIRCUITBAP', GetControlText('CODECIRCUIT'));
      T.PutValue('CCI_LIBELLE', GetControlText('LIBCIRCUIT'));
      T.PutValue('CCI_VISEUR1', FListe.CellValues[2, n]);
      T.PutValue('CCI_VISEUR2', FListe.CellValues[3, n]);
      T.PutValue('CCI_NBJOUR', FListe.Cells[4, n]);
    end;

    tDonnees.InsertDb(nil);
    CommitTrans;
  except
    RollBack;
  end;
  {Si jamais on a créé un circuit}
  AvertirTable   ('CPCIRCUITBAP');
  MajInfoTablette('CPCIRCUITBAP');
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCIRCUITBAP.TesteCode : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  if Action = taCreat then
    Result := not ExisteSQL('SELECT CCI_CIRCUITBAP FROM CPCIRCUIT WHERE CCI_CIRCUITBAP = "' +
                            GetControlText('CODECIRCUIT') + '" ORDER BY CCI_NUMEROORDRE');
  if not Result then
    HShowMessage('0;' + Ecran.Caption + ';Le code existe déjà.;W;O;O;O;', '', '');
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCIRCUITBAP.ValideSaisie: Boolean;
{---------------------------------------------------------------------------------------}

    {13/06/07 : cette fonction était auparavant dans ValideLigne
    {-------------------------------------------------------------------}
    function _TesteViseur(Ind, Row, Col : Integer) : Boolean;
    {-------------------------------------------------------------------}
    var
      Vis : TObjViseur;
    begin
      Result := True;

      if (Ind > -1)then begin
        Vis := TObjViseur(lViseurs.Objects[Ind]);
        Result := (Vis.Row = Row) {and (Vis.Col = Col)};
      end;
    end;

    {---------------------------------------------------------------------}
    function ValideGrille : Boolean;
    {---------------------------------------------------------------------}
    var
      n : Integer;
    begin
      Result := True;
      for n := 1 to FListe.RowCount - 1 do
        if not ValideLigne(False, n) then begin
          Result := False;
          Break;
        end;
    end;

var
  n    : Integer;
  MsgC : Boolean;
  MsgR : Boolean;
  MsgA : string;
  Ref  : Integer;
begin
  Result := False;
  if not TesteCode then
    SetFocusControl('CODECIRCUIT')
  else if GetControlText('LIBCIRCUIT') = '' then begin
    HShowMessage('0;' + Ecran.Caption + ';Vous devez saisir un libellé.;W;O;O;O;', '', '');
    SetFocusControl('LIBCIRCUIT');
  end
  else if not ValideGrille then
    SetFocusControl('FLISTE')
  else
    Result := True;

  {13/06/07 : Demande de SIC d'un simple message de confirmation de l'utilisation d'un même viseur
              dans plusieurs lignes ou colonnes}
  MsgC := False;
  MsgR := False;

  for n := 1 to FListe.RowCount - 1 do begin
    if FListe.CellValues[2, n] = FListe.CellValues[3, n] then MsgR := True;
    if not MsgC then begin
      Ref := lViseurs.IndexOf(FListe.CellValues[2, n]);
      MsgC := not _TesteViseur(Ref, n, 2);
      if not MsgC then begin
        Ref := lViseurs.IndexOf(FListe.CellValues[3, n]);
        MsgC := not _TesteViseur(Ref, n, 3);
      end;
    end;
    if MsgC and MsgR then Break;
  end;

  if MsgC or MsgR then begin
    Msga := TraduireMemoire('Certains viseurs sont utilisés plusieurs fois dans le circuit : ');
    if MsgR then MsgA := MsgA + #13 + TraduireMemoire('Certains viseurs sont présents deux fois dans la même étape');
    if MsgC then MsgA := MsgA + #13 + TraduireMemoire('Certains viseurs sont présents au moins dans deux étapes');
    Result := PgiAsk(MsgA, Ecran.Caption) = mrYes;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCIRCUITBAP.ValideLigne(AvecMaj : Boolean; Row: Integer) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Ref : Integer;
begin
  Result := not ((Fliste.CellValues[0, Row] = '') or (Fliste.CellValues[1, Row] = '') or
                 (Fliste.CellValues[2, Row] = '') or (Fliste.CellValues[3, Row] = '') or
                 (Fliste.CellValues[4, Row] = ''));

  if not Result then
    PgiError(TraduireMemoire('Veuillez renseigner tous les champs'), Ecran.Caption)

  else if AvecMaj then begin
    Ref := lViseurs.IndexOf(FListe.CellValues[2, Row]);
    if Ref = -1 then RemplirLigneViseurs(2, Row);
    Ref := lViseurs.IndexOf(FListe.CellValues[3, Row]);
    if Ref = -1 then RemplirLigneViseurs(3, Row);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FormCanCloseQuery(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := PeutFermer;
  if PeutFermer then FFormCanClose(Sender, CanClose);
  PeutFermer := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.FListeOnDbleClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FListe.Col in [2..3] then
    FlisteEllipsis(Sender);
end;

{FQ 18296 : On s'assure que le viseur saisi est valide
{---------------------------------------------------------------------------------------}
function TOF_CPCIRCUITBAP.TesteUser(C, R : Byte) : Boolean;
{---------------------------------------------------------------------------------------}
var
  u : string;
begin
  if C in [2..3] then begin
    {Si la ligne est vide}
    if (FListe.CellValues[2, R] = '') and (FListe.CellValues[3, R] = '') and (FListe.CellValues[4, R] = '') then begin
      Result := True;
      if FListe.Row <> R then FListe.Row := R; 
      Detruire;
    end
    else begin
      u := FListe.CellValues[C, R];
      if lUtilisat.IndexOf(u) = -1 then begin
        HShowMessage('0;' + Ecran.Caption + ';Ce viseur n''est pas autorisé.;W;O;O;O;', '', '');
        Result := False;
      end
      else begin
        Result := True;
        RemplirListeViseurs;{JP 30/05/07}
      end;
    end;
  end
  else
    Result := True;
end;

{21/06/07 : FQ 20776 : Initialisation du code circuit en création
{---------------------------------------------------------------------------------------}
procedure TOF_CPCIRCUITBAP.InitCodeCircuit;
{---------------------------------------------------------------------------------------}
begin
  if Action <> taCreat then Exit;
  {21/06/07 : FQ 20777 : génération automatique du Code}
  if not Assigned(ObjetCode) then begin
    ObjetCode := TObjCodeCombo.Create('CPCIRCUIT', 'CCI_CIRCUITBAP');
    ObjetCode.LastCode := '';
  end;
  SetControlText('CODECIRCUIT', ObjetCode.GetNewCode);
end;


initialization
  RegisterClasses([TOF_CPCIRCUITBAP]);

end.


