{***********UNITE*************************************************
Auteur  ...... : Guillon Stéphane
Créé le ...... : 28/01/2005
Modifié le ... : 18/01/2006
Description .. : Source TOF de la FICHE : CPVENTILTYPECROIS ()
Suite ........ : 
Suite ........ : JP 28/04/05 : FQ 15781 et 15783
Suite ........ : 
Suite ........ : JP 18/01/06 : FQ 17268 : gestion du F6
Mots clefs ... : TOF;CPVENTILTYPECROIS
*****************************************************************}
unit CPVENTILTYPECROIS_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  SaisUtil,
  AglInit, //ActionToString
  Ent1,
  ParamSoc, //GetParamSocSecur
  Graphics,
  ULibWindows, //SetGridgrise
  Grids, // TGridDrawState
  LookUp, //Lookuplist
  UtilPGI, //_BLoquage, _Bloqueur
  UTOB,
  HTB97, //TToolBarButton97
  windows, //VK_...
  CPSECTION_TOM, //FicheSection
  menus, //TPopupMenu
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  fe_main,
  mul,
  {$ELSE}
  eMul,
  MaineAGL,
  {$ENDIF}
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
  SousSect,              {FP 19/10/2005 FQ15725 Pour ChoisirSSection}
{$ENDIF}
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UlibAnalytique ,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UTOF  ,UentCommun;
 // CPMODRESTANA_TOM;          {FP 29/12/2005}

type
  TOF_CPVENTILTYPECROIS = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    //Controles
    GSA: THGrid;
    TempControl: THEdit;
    GX, GY: LongInt;

    //Paramètres du point d'entrée
    reseau: boolean;
    Nature, Compte: string;
    Action: TActionFiche;

    //Autres
    Axes: array[1..MaxAxe] of boolean;
    dernier_axe: integer;
    premier_axe: integer;
    SensParcours: boolean;
    vTobVentil: TOB;
    bModifier : boolean;
    SurGeneraux : Boolean; {JP 28/04/05 : FQ 15781}
    sVentilType : string;  {JP 28/04/05 : FQ 15781}
    FRestriction: TRestrictionAnalytique;        // Modèlde de restriction ana FP 29/12/2005}
    //Evts
    procedure GSAPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSASelectCell(Sender: TObject; ACol, ARow: Longint; var CanSelect: Boolean);
    procedure GSAElipsisClick(Sender: TObject);
    procedure GSACellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSACellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSARowExit(Sender: TObject; ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSARowEnter(Sender: TObject; ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure EcranFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValiderClick(Sender: TObject);
    procedure BAddRowClick(Sender: TObject);
    procedure BDelRowClick(Sender: TObject);
    procedure GSAMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSADblClick(Sender: TObject); {FP 19/10/2005 FQ15725}
    procedure POPSPopup(Sender : TObject);
    procedure EcranCloseQuery(Sender: TObject; var CanClose: Boolean);
    {JP 28/04/05 : FQ 15781 : pour charger la ventilation type si besoin est}
    procedure CbVentilTypeClick(Sender : TObject);

    //Autres
    procedure ChargVentilGrid(Nat : string = ''; Cpt : string = ''); {JP 28/04/05 : FQ 15781 : Ajout des paramètres}
    function ReturnAxeParam(ACol: integer; typecol: string): integer;
    function ReturnIndiceArray(tableau: array of integer; Valeur: integer): integer;
    procedure NewLigne(ARow: LongInt);
    function LigneCorrecte(ARow: LongInt; affiche: boolean): boolean;
    procedure MAJMontants;
    procedure ClickZoom(AfficheFicheSection: Boolean = True);   {FP 19/10/2005 FQ15725}
    procedure ClickZoomSousplan;                                {FP 19/10/2005 FQ15725}
    procedure DoSolde(Grid : THGrid); {JP 17/01/06 : FQ 17268}
    function LookUpSectionList(Ctrl: TControl; Axe: String; ARow, NumTag: Integer): Boolean;    {FP 29/12/2005}
    function ExisteSQLSection(Axe, Section: String; ARow: Integer): Boolean;                    {FP 29/12/2005}
    function GetClauseRestrictionAna(Axe: String; ARow: Integer): String; {FP 19/04/2006 FQ17728}
  end;


procedure ParamVentilCroise(Nature, Compte: string; Comment: TActionFiche; Reseau: boolean);

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  Messages {VK_TAB}
  ,UFonctionsCBP;

//Définition de constantes

const
  AN_MNumL: integer = 0;
  AN_MSect: array[1..5] of integer = (1, 3, 5, 7, 9);
  AN_MLib: array[1..5] of integer = (2, 4, 6, 8, 10);
  AN_MPourcent: integer = 11;
  AN_MPourQt1: integer = 12;
  AN_MPourQt2: integer = 13;


procedure ParamVentilCroise(Nature, Compte: string; Comment: TActionFiche; Reseau: boolean);
var
  Str_reseau: string;
begin
  if Reseau then if _Blocage(['nrCloture', 'nrBatch'], True, 'nrBatch') then Exit;
  if Reseau then Str_reseau := '1' else Str_reseau := '0';
  AGlLanceFiche('CP', 'CPVENTILTYPECROIS', '', '', ActionToString(Comment) + ';' + Nature + ';' + Compte + ';' + Str_reseau + ';');
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 10/12/2004
Modifié le ... :   /  /
Description .. : Gestion de l'entrée dans une cellule
Mots clefs ... : CELLENTER
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.GSACellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
  if GSA.ColEditables[GSA.Col] = False then
  begin
    if SensParcours then GSA.Col := GSA.Col + 1 else GSA.Col := GSA.Col - 1;
  end;

  if ReturnIndiceArray(AN_MLib, GSA.Col) <> -1 then
  begin
    if SensParcours then GSA.Col := GSA.Col + 1 else GSA.Col := GSA.Col - 1;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 10/12/2004
Modifié le ... :   /  /
Description .. : Gestion de la sortie d'une cellule
Mots clefs ... : CELLEXIT
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.GSACellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
var
  iAxe: integer;
  sAxe: string;
begin
  if csDestroying in Ecran.ComponentState then exit;
  SensParcours := (ACol < GSA.Col);
  if Action = taConsult then exit;
  if GSA.ColEditables[GSA.Col] = False then exit;

  //Si Section renseigné on vérifie qu'elle existe > autrement Cancel et lookup
  if ReturnIndiceArray(AN_MSect, ACol) <> -1 then
  begin
    if (ARow = GSA.RowCount - 1) and (ARow > 2) and (ARow <> GSA.Row) then Exit;

    //Calcul de l'axe
    if ((ARow <> GSA.Row) and (GSA.Cells[AN_MPourcent, ARow] = '')) then
    begin
      Cancel := true;
      Exit;
    end;
    iAxe := ReturnAxeParam(ACol, 'tSection');
    {b FP 29/12/2005 if not (ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="' + GSA.Cells[ACol, Arow] + '" AND S_AXE="A' + IntToStr(iAxe) + '"')) then}
    if not ExisteSQLSection('A'+IntToStr(iAxe), GSA.Cells[ACol, Arow], ARow) then
    {e FP 29/12/2005}
    begin
      bModifier := True;
      TempControl.Text := '';
      {b FP 29/12/2005 if LookUpList(TempControl, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE="A' + IntToStr(iAxe) + '"', 'S_SECTION', True, 0) then}
      if LookUpSectionList(TempControl, 'A' + IntToStr(iAxe), ARow, 0) then
      {e FP 29/12/2005}
      begin
        GSA.Cells[ACol, ARow] := TempControl.Text;
      end
      else
      begin
        Cancel := true;
        exit;
      end;
    end;
    sAxe := InttoStr(iAxe);
    if sAxe = '1' then sAxe := '';
    GSA.Cells[ACol + 1, ARow] := RechDom('TZSECTION' + sAxe, GSA.Cells[ACol, ARow], false);

    //si dernier axe, on cree un nouvelle ligne
    if iAxe = dernier_axe then
    begin
      if (GSA.Row = GSA.RowCount - 1) and (GSA.Row = ARow) then
      begin
        GSA.RowCount := GSA.RowCount + 1;
        NewLigne(GSA.RowCount - 1);
      end;
    end;
  end;

  //Si cellule pourcentage (valeur , qte1 ou qt2)
  if (ACol = AN_MPourcent) or (ACol = AN_MPourQt1) or (ACol = AN_MPourQt2) then
  begin
    if GSA.Cells[ACol, ARow] = '' then
    begin
      bModifier := True;
      GSA.Cells[ACol,ARow] := '0.000';
    end
    else if not (IsNumeric(GSA.Cells[ACol, ARow])) then
    begin
      bModifier := True;
      GSA.Cells[ACol, ARow] := '';
      if (ARow <> GSA.Row) then
      begin
        Cancel := True;
        Exit;
      end;
    end
    else
    begin
      bModifier := True;
      GSA.Cells[ACol, ARow] := strfmontant(Valeur(GSA.Cells[ACol, ARow]), 15, 3, '', False);
      MAJMontants;
    end;
    if (ACol = AN_MPourcent) then
    begin
      if GSA.Cells[AN_MPourQt1,ARow] = '' then GSA.Cells[AN_MPourQt1,ARow]:= '0.000';
      if GSA.Cells[AN_MPourQt2,ARow] = '' then GSA.Cells[AN_MPourQt2,ARow]:= '0.000';
    end;
  end;

end;


procedure TOF_CPVENTILTYPECROIS.OnNew;
begin
  inherited;
end;

procedure TOF_CPVENTILTYPECROIS.OnDelete;
begin
  inherited;
end;

procedure TOF_CPVENTILTYPECROIS.OnUpdate;
begin
  inherited;
end;

procedure TOF_CPVENTILTYPECROIS.OnLoad;
begin
  inherited;
end;

procedure TOF_CPVENTILTYPECROIS.OnArgument(S: string);
var
  i: integer;
  StArgs: string;
  BAddRow, BDelRow, BValider: TToolBarButton97;
  POPS : TPopupMenu;
  CB : THValComboBox;{JP 28/04/05 : FQ 15781}
begin
  inherited;
  GX := 0;
  GY := 0;
  Ecran.HelpContext := 1460100;
  StArgs := S;
  FRestriction := TRestrictionAnalytique.Create;          {FP 29/12/2005}
  //Récupération des paramètres
  Action := StringToAction(ReadToKenPipe(StArgs, ';'));
  Nature := ReadTokenPipe(StArgs, ';');
  Compte := ReadTokenPipe(StArgs, ';');
  if ReadTokenPipe(StArgs, ';') = '1' then reseau := True else reseau := False;

  {JP 28/04/05 : FQ 15781 : Gestion des ventilations types sur les généraux}
  SurGeneraux := Nature = 'GE';
  SetControlVisible('LBVENTILETYPE', SurGeneraux);
  SetControlVisible('CBVENTILETYPE', SurGeneraux);
  SetControlEnabled('LBVENTILETYPE', Action <> taConsult);
  SetControlEnabled('CBVENTILETYPE', Action <> taConsult);
  CB := THValComboBox(GetControl('CBVENTILETYPE'));
  if Assigned(CB) then CB.OnClick := CbVentilTypeClick;

  //Bloquage
  if Reseau then _Bloqueur('nrBatch', False);

  //Récupération des contrôles
  GSA := THGrid(GetControl('GSA', true));
  TempControl := THEdit(GetControl('TempControl', true));
  BValider := TToolBarButton97(GetControl('BValider', true));
  BAddRow := TToolBarButton97(getcontrol('BAddRow', true));
  BDelRow := TToolBarButton97(getcontrol('BDelRow', true));
  POPS := TPopupMenu(getcontrol('POPS',true));

  //Gestion evts
  GSA.PostDrawCell := GSAPostDrawCell;
  GSA.OnSelectCell := GSASelectCell;
  GSA.OnElipsisClick := GSAElipsisClick;
  GSA.OnCellExit := GSACellExit;
  GSA.OnCellEnter := GSACellEnter;
  GSA.OnRowExit := GSARowExit;
  GSA.OnRowEnter := GSARowEnter;
  Ecran.OnKeyDown := EcranFormKeyDown;
  BValider.OnClick := BValiderClick;
  BAddRow.OnClick := BAddRowClick;
  BDelRow.OnClick := BDelRowClick;
  GSA.OnMouseDown := GSAMouseDown;
  GSA.OnDblClick  := GSADblClick; {FP 19/10/2005 FQ15725}
  POPS.OnPopup := POPSPopup;
  Ecran.OnCloseQuery := EcranCloseQuery;


  AffecteGrid(GSA, Action);


  //Initialisation des variables
  bModifier := False;

  vTobVentil := TOB.Create('$VENTIL', nil, -1);

  dernier_axe := 0;
  premier_axe := 0;
  for i := 1 to MaxAxe do
  begin
    Axes[i] := GetParamSocSecur('SO_VENTILA' + IntToStr(i),False);
    if not Axes[i] then
    begin
      GSA.ColEditables[i * 2 - 1] := False;
      GSA.ColEditables[i * 2] := False;
    end
    else
    begin
      dernier_axe := i;
      if premier_axe = 0 then premier_axe := i;
    end;
  end;

  i := 1;
  while (i < 11) do
  begin
    GSA.Cells[i, 1] := 'Section';
    GSA.Cells[i + 1, 1] := 'Intitulé';
    Inc(i, 2);
  end;

  GSA.ColFormats[AN_MPourcent] := '0.000';
  GSA.ColFormats[AN_MPourQt1] := '0.000';
  GSA.ColFormats[AN_MPourQt2] := '0.000';


  Focuscontrole(GSA, true);

  ChargVentilGrid;
end;

procedure TOF_CPVENTILTYPECROIS.OnClose;
begin
  inherited;
  FreeAndNil(vTobVentil);
  FreeAndNil(FRestriction);                   {FP 29/12/2005}
  //On recharge la tablette
  AvertirTable('ttVentilType') ;
end;

procedure TOF_CPVENTILTYPECROIS.OnDisplay();
begin
  inherited;
end;

procedure TOF_CPVENTILTYPECROIS.OnCancel();
begin
  inherited;
end;

procedure TOF_CPVENTILTYPECROIS.GSAPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if GSA.ColEditables[ACol] = False then
  begin
    GSA.PostDrawCell := nil;
    SetGridGrise(ACol, ARow, GSA);
    GSA.PostDrawCell := GSAPostDrawCell;
  end;
end;

procedure TOF_CPVENTILTYPECROIS.GSASelectCell(Sender: TObject; ACol, ARow: Longint; var CanSelect: Boolean);
begin
  GSA.ElipsisButton := (ReturnIndiceArray(AN_MSect, ACol) <> -1);
end;

procedure TOF_CPVENTILTYPECROIS.GSAElipsisClick(Sender: TObject);
begin
  {b FP 29/12/2005 LookUpList(GSA, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE="A' + IntToStr(ReturnAxeParam(GSA.Col, 'tSection')) + '"', 'S_SECTION', True, 0);}
  LookUpSectionList(GSA, 'A' + IntToStr(ReturnAxeParam(GSA.Col, 'tSection')), GSA.Row, 0);
  {e FP 29/12/2005}
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 08/12/2004
Modifié le ... : 10/12/2004
Description .. : Fonction qui retourne le numéro d'axe par rapport au
Suite ........ : numéro de colonne du grid (dans le cas de colonne de
Suite ........ : Code Section ou libelle)
Suite ........ : param : tSection pour colonne de section
Suite ........ : param : tLibelle pour colonne de libelle
Mots clefs ... : CODE SECTION, LIBELLE, COLONNE, NUMERO AXE
*****************************************************************}
function TOF_CPVENTILTYPECROIS.ReturnAxeParam(ACol: integer; typecol: string): integer;
begin
  result := -1;
  if typecol = 'tSection' then
  begin
    if ACol = AN_MSect[1] then
    begin
      result := 1;
    end
    else if ACol = AN_MSect[2] then
    begin
      result := 2;
    end
    else if ACol = AN_MSect[3] then
    begin
      result := 3;
    end
    else if ACol = AN_MSect[4] then
    begin
      result := 4;
    end
    else if ACol = AN_MSect[5] then
    begin
      result := 5;
    end;
  end
  else if typecol = 'tLibelle' then
  begin
    if ACol = AN_Mlib[1] then
    begin
      result := 1;
    end
    else if ACol = AN_MLib[2] then
    begin
      result := 2;
    end
    else if ACol = AN_MLib[3] then
    begin
      result := 3;
    end
    else if ACol = AN_MLib[4] then
    begin
      result := 4;
    end
    else if ACol = AN_MLib[5] then
    begin
      result := 5;
    end;
  end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 13/12/2004
Modifié le ... :   /  /
Description .. : Fonction qui recherche la valeur dans un tableau et
Suite ........ : retourne l'indice de la valeur dans le tableau
Suite ........ : retourne -1 si la valeur n'est pas dans le tableau
Mots clefs ... : INDICE; ARRAY; VALEUR
*****************************************************************}
function TOF_CPVENTILTYPECROIS.ReturnIndiceArray(tableau: array of integer; Valeur: integer): integer;
var
  i: integer;
begin
  result := -1;
  for i := Low(tableau) to High(tableau) do
  begin
    if (tableau[i] = Valeur) then
    begin
      result := i;
      break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 31/01/2005
Modifié le ... : 28/04/2005
Description .. : Procedure qui permet de charger le grid
Suite ........ : JP : FQ 15781 : Ajout des paramètres ()
Mots clefs ... : CHARGEMENT GRID
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.ChargVentilGrid(Nat : string = ''; Cpt : string = '');
var
  iAxe, iRow : integer;
  sAxe : string;
  CompteAna:  array[1..MaxAxe] of String;         {FP 29/12/2005}
  {b FP 19/04/2006 FQ17725}
  SQL:             String;
  QrySansFiltre:   String;
  vTobVentilAll:   TOB;
  {e FP 19/04/2006}
begin
  {JP 28/04/05 : Dans certains cas, on ne peut se contenter des paramètres généraux}
  if Nat = '' then Nat := Nature;
  if Cpt = '' then Cpt := Compte;
  FillChar(CompteAna, sizeof(CompteAna), #0);     {FP 29/12/2005}

  {b FP 19/04/2006 FQ17725: Requête sans filtre sur la restriction analytique}
  vTobVentilAll := TOB.Create('$VENTIL', nil, -1);

  QrySansFiltre := 'SELECT * FROM VENTIL WHERE V_NATURE = "' + Nat + IntToStr(premier_axe) +'"'+
       ' AND V_COMPTE = "' + Cpt + '"';
  SQL := QrySansFiltre
       {b FP 29/12/2005}
       +' AND ' + FRestriction.GetClauseCompteAutorise(
         Compte, 'A'+IntToStr(premier_axe), 'VENTIL', CompteAna);
       {e FP 29/12/2005}
  if (Nat = 'TY') and (not SurGeneraux) then  {Paramétrage des Ventilation type}
    begin
    vTobVentil.LoadDetailDBFromSQL('VENTIL',QrySansFiltre);
    end
  else
    begin
    vTobVentil.LoadDetailDBFromSQL('VENTIL',SQL);
    vTobVentilAll.LoadDetailDBFromSQL('VENTIL',QrySansFiltre);
    FRestriction.VerifModelVentil(Compte, vTobVentil, vTobVentilAll, 'VENTIL', 'A'+IntToStr(premier_axe), Cpt<>Compte);
    end;
  if vTobVentil.Detail.Count = 0 then
{e FP 19/04/2006}
  begin
    vTobVentil.ClearDetail ; {FP 29/12/2005+19/04/2006 FQ17725}
    NewLigne(2);
  end
  else
  begin
  {FQ 15781 : Si on charge unventilation type sur un compte général, il faut changer la Nature, sinon on
              va droit vers une violation de clef !!}
    if SurGeneraux and (Nat <> Nature) then {Nat = 'TY' et Nature = 'GE'}
      for iRow := 0 to vTobVentil.Detail.Count - 1 do begin
        vTobVentil.Detail[iRow].SetString('V_NATURE', Nature + IntToStr(premier_axe));
        vTobVentil.Detail[iRow].SetString('V_COMPTE', Compte);
      end;
    {b FP 29/12/2005}
    for iRow := 0 to vTobVentil.Detail.Count - 1 do begin
      vTobVentil.Detail[iRow].SetString('V_NUMEROVENTIL', IntToStr(iRow+1));
    end;
    {e FP 29/12/2005}

    vTobVentil.PutGridDetail(GSA, False, False, 'V_NUMEROVENTIL;V_SOUSPLAN1;;V_SOUSPLAN2;;V_SOUSPLAN3;;V_SOUSPLAN4;;V_SOUSPLAN5;;V_TAUXMONTANT;V_TAUXQTE1;V_TAUXQTE2', True);

    //Mise à jour des libellés
    for iRow := 2 to GSA.RowCount - 1 do
    begin
      for iAxe :=1 to MaxAxe do
      begin
        if Axes[iAxe] then
        begin
          if iAxe = 1 then sAxe := '' else sAxe := IntToStr(iAxe);
          GSA.Cells[AN_MLib[iAxe],iRow] := RechDom('TZSECTION' + sAxe, GSA.Cells[AN_MSect[iAxe],iRow] ,False);
        end;
      end;
    end;


    MAJMontants;
  end;

  FreeAndNil(vTobVentilAll);  {FP 19/04/2006 FQ17725}
end;

procedure TOF_CPVENTILTYPECROIS.NewLigne(ARow: LongInt);
var
  vTobVentilFille: TOB;
  i: integer;
begin
  if ARow < 2 then Exit;
  vTobVentilFille := TOB.Create('VENTIL', vTobVentil, ARow - 2);

  vTobVentilFille.PutValue('V_NATURE', Nature + IntToStr(premier_axe));
  vTobVentilFille.PutValue('V_COMPTE', Compte);
  vTobVentilFille.PutValue('V_SECTION', '');
  vTobVentilFille.PutValue('V_NUMEROVENTIL', ARow - 1);
  vTobVentilFille.PutValue('V_MONTANT', 0.00);
  vTobVentilFille.PutValue('V_SOCIETE', V_PGI.CodeSociete);

  //On renumerote les tob fille qui suivent (pour V_NUMEROVENTIL)
  //On renumerote le grid
  for i := ARow - 1 to vTobVentil.Detail.count - 1 do
  begin
    vTobVentil.Detail[i].PutValue('V_NUMEROVENTIL', i + 1);
    GSA.Cells[AN_MNuml,i + 2] := IntToStr(i+1);
  end;

  //Mise à jour
  GSA.Cells[AN_MPourcent, ARow] := '';
  GSA.Cells[AN_MPourQt1, ARow] := '';
  GSA.Cells[AN_MPourQt2, ARow] := '';
  GSA.Cells[AN_MNuml, ARow] := IntToStr(ARow - 1);
end;

procedure TOF_CPVENTILTYPECROIS.GSARowExit(Sender: TObject; ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
  if not ((ou = GSA.RowCount - 1) and (ou > 2)) then
  begin
    if not LigneCorrecte(ou, True) then
    begin
      Cancel := True;
      Exit;
    end;
  end;

  //Mise à jour tob
  vTobVentil.DeTail[ou - 2].GetLigneGrid(GSA, ou, ';V_SOUSPLAN1;;V_SOUSPLAN2;;V_SOUSPLAN3;;V_SOUSPLAN4;;V_SOUSPLAN5;;V_TAUXMONTANT;V_TAUXQTE1;V_TAUXQTE2');
  vTobVentil.DeTail[ou - 2].PutValue('V_SECTION',vTobVentil.DeTail[ou - 2].GetValue('V_SOUSPLAN'+IntToStr(premier_axe)));
end;

procedure TOF_CPVENTILTYPECROIS.GSARowEnter(Sender: TObject; ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
  if not LigneCorrecte(GSA.Row, False) then
  begin
    SensParcours := True;
    GSA.Col := AN_MSect[premier_axe];
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : Fonction qui permet de savoir si une ligne du grid est
Suite ........ : correcte ou pas et si non, permet d'afficher un message
Suite ........ : d'erreur (a la demande)
Mots clefs ... : LIGNE CORRECTE
*****************************************************************}
function TOF_CPVENTILTYPECROIS.LigneCorrecte(ARow: LongInt; Affiche: boolean): boolean;
var
  i: Integer;
begin
  for i := 1 to MaxAxe do
  begin
    if Axes[i] then
    begin
      if GSA.Cells[2 * i - 1, ARow] = '' then
      begin
        if Affiche then PGIInfo('Vous devez sélectionnez une section pour l''axe ' + IntToStr(i), 'Erreur sur la ligne');
        GSA.Col := AN_MSect[i];
        bModifier := True;
        result := False;
        Exit;
      end;
    end;
  end;

  if not (IsNumeric(GSA.Cells[AN_MPourcent, ARow])) or not (IsNumeric(GSA.Cells[AN_MPourQt1, ARow])) or not (IsNumeric(GSA.Cells[AN_MPourQt2, ARow])) then
  begin
    result := False;
    if affiche then PGIInfo('Les pourcentages doivent être de type numérique', 'Erreur sur la ligne');
    GSA.Col := AN_MPourcent;
    bModifier := True;
    Exit;
  end;

  {b FP 21/06/2006 FQ18422
  if Valeur(GSA.Cells[AN_MPourcent, ARow]) = 0 then
  begin
    result := False;
    if affiche then PGIInfo('Le pourcentage de la valeur ne doit pas être nul', 'Erreur sur la ligne');
    GSA.Col := AN_MPourcent;
    bModifier := True;
    Exit;
  end;
  e FP 21/06/2006 FQ18422}

  result := True;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 01/02/2005
Modifié le ... :   /  /
Description .. : Mise à jour de montant
Mots clefs ... : MAJ MONTANT
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.MAJMontants;
var
  i: integer;
  totpourVal, totpourQt1, totpourQt2: double;
begin
  totpourVal := 0;
  totpourQt1 := 0;
  totpourQt2 := 0;
  //Calcul
  for i := 2 to GSA.RowCount - 1 do
  begin
    if GSA.Cells[AN_MPourcent, i] <> '' then totpourVal := totpourVal + Valeur(GSA.Cells[AN_MPourcent, i]);
    if GSA.Cells[AN_MPourQt1, i] <> '' then totpourQt1 := totpourQt1 + Valeur(GSA.Cells[AN_MPourQt1, i]);
    if GSA.Cells[AN_MPourQt2, i] <> '' then totpourQt2 := totpourQt2 + Valeur(GSA.Cells[AN_MPourQt2, i]);
  end;

  //Mise à jour des controles
  SetControlText('FTOTVAL', strfmontant(totpourval, 15, 3, '', False));
  SetControlText('FTOTQTE1', strfmontant(totpourqt1, 15, 3, '', False));
  SetControlText('FTOTQTE2', strfmontant(totpourqt2, 15, 3, '', False));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 01/02/2005
Modifié le ... :   /  /
Description .. : Gestion des triggers claviers
Mots clefs ... : GESTION CLAVIER
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.EcranFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  vide, OkGrid: boolean;            {FP 19/10/2005 FQ15725}
begin
  OkGrid := (Ecran.ActiveControl = GSA);
  vide   := (shift = []);           {FP 19/10/2005 FQ15725}

  case Key of
    VK_ESCAPE: Ecran.Close;
    VK_F10: BValiderClick(nil);
    {b FP 19/10/2005 FQ15725 Provient de CPVENTIL_TOF}
    //VK_F5: ClickZoom;
    VK_F5: if ((OkGrid) and (vide)) then
      begin
        ClickZoom;
        Key := 0;
      end
      else if ((OkGrid) and ({FP FQ15725 (shift = [ssCtrl]) or} (shift = [ssAlt]))) then
      begin
        ClickZoomSousplan;
      end
      {b FP FQ15725}
      else if (OkGrid) and (shift = [ssCtrl]) then
      begin
        shift := [];
        GSAElipsisClick(GSA);
      end;
      {e FP FQ15725}
    {JP 17/01/06 : FQ 17268 : Gestion du calcul automatique du solde}
    VK_F6 : if (OkGrid) and (vide) then begin
              Key := 0;
              DoSolde(GSA);
            end;
    {e FP 19/10/2005 FQ15725}
    VK_RETURN: Key := VK_TAB;
    VK_DELETE: if ((OkGrid) and (shift = [ssCtrl])) then begin BDelRowClick(nil); key :=0; end;
    VK_INSERT: BAddRowClick(nil);
  end;
end;

procedure TOF_CPVENTILTYPECROIS.BValiderClick(Sender: TObject);
var
  Bbool : boolean;
  i     : Integer; 
begin
  if Ecran.ActiveControl is THGrid then begin
    {JP 28/04/05 : FQ 15783 : Les SpeedButtons ne prenant pas le Focus, on simule un changement de ligne
                   (vers le haut, pour contourner le test "GSA.Row <> GSA.Rowcount - 1") afin d'être sûr
                   que les CellExit RoxExit sont bien exécutés}
    PostMessage(Ecran.ActiveControl.Handle, WM_KEYDOWN, VK_UP, 0);
    Application.ProcessMessages;
  end;

  //On vérifie la ligne en cours
  Bbool := False;
  if GSA.Row <> GSA.Rowcount - 1 then GSARowExit(nil,GSA.Row,Bbool,False);
  if Bbool then Exit;
  {b FP 21/06/2006 FQ18422}
  //On vérifié la derniere ligne
  //if GSA.Cells[AN_MPourcent,GSA.RowCount - 1] = '' then vTobVentil.Detail[GSA.RowCount - 3].Free;
  if GSA.Cells[AN_MSect[1],GSA.RowCount - 1] = '' then vTobVentil.Detail[GSA.RowCount - 3].Free;
  {e FP 21/06/2006 FQ18422}
  {b FP 29/12/2005: Vérifie la validité de la ligne}
  for i := 0 to vTobVentil.Detail.Count-1 do
    begin
    if not ExisteSQLSection('A'+IntToStr(dernier_axe),
             vTobVentil.Detail[i].GetString('V_SOUSPLAN'+IntToStr(dernier_axe)),
             i+2) then
      begin
      PGIError(TraduireMemoire('Ligne de ventilation incorrecte'), Ecran.Caption);
      GSA.Row := i+2;
      Exit;
      end;
    end;
  {e FP 29/12/2005}
  //Enregistrement
  try
    begintrans;

    ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE="' + Nature + IntToStr(premier_axe) + '" AND V_COMPTE="' + Compte + '"');

    vTobVentil.InsertDB(nil);

    committrans;
  except
    on e : exception do
    begin
      rollback;
      PGIInfo(e.message,'Erreur');
      Exit;
    end;
  end;

  BModifier := False;

  Ecran.Close;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 13/12/2004
Modifié le ... :   /  /
Description .. : zoom suivant la cell séléctionné du grid
Mots clefs ... : ZOOM
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.ClickZoom(AfficheFicheSection: Boolean = True);   {FP FQ15725}
var
  cell_section: boolean;
  iAxe: integer;
  ACol: longint;
  ARow: longint;
begin
  //Sections
  //si section existe on ouvre la fénétre de la section autrement lookuplist sur section
  if Action = taConsult then
  begin
    GSA.MouseToCell(GX, GY, ACol, ARow);
  end
  else
  begin
    ARow := GSA.Row;
    ACol := GSA.Col;
  end;

  //SG6 27.01.05 On est dans les ligne de titres du grille !
  if ARow < 2 then Exit;

  cell_section := ReturnIndiceArray(AN_MSect, ACol) <> -1;
  if cell_section then
  begin
    iAxe := ReturnAxeParam(ACol, 'tSection');
    {b FP 29/12/2005 if not (ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="' + GSA.Cells[Acol, ARow] + '" AND S_AXE="A' + IntToStr(iAxe) + '"')) then}
    if not ExisteSQLSection('A'+IntToStr(iAxe), GSA.Cells[ACol, Arow], ARow) then
    {e FP 29/12/2005}
    begin
      TempControl.Text := '';
{b FP 29/12/2005 if LookUpList(TempControl, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE="A' + IntToStr(iAxe) + '"', 'S_SECTION', True, 0) then}
      if LookUpSectionList(TempControl, 'A' + IntToStr(iAxe), ARow, 0) then
      {e FP 29/12/2005}
      begin
        GSA.Cells[ACol, ARow] := TempControl.Text;
      end
    end
    else if AfficheFicheSection then   {FP FQ15725}
    begin
      FicheSection(nil, 'A' + IntToStr(iAxe), GSA.Cells[ACol, ARow], taConsult, 0);
    end;
  end;
end;

{b FP 19/10/2005 FQ15725: Provient de CPVENTIL_TOF}
procedure TOF_CPVENTILTYPECROIS.ClickZoomSousplan;
var
   iAxe: integer;
   Sect: string;
   fb: TFichierBase;
begin
  if ReturnIndiceArray(AN_MSect, GSA.Col) = -1 then exit;
  iAxe := ReturnAxeParam(GSA.Col, 'tSection');
  fb := AxeTofb('A' + IntToStr(iAxe));
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
  Sect := ChoisirSSection(fb, GSA.Cells[GSA.Col, GSA.Row], FALSE, Action);
{$ENDIF}
{$ENDIF}
  if Sect <> '' then GSA.Cells[GSA.Col, GSA.Row] := sect;
  ClickZoom(False);
end;
{e FP 19/10/2005 FQ15725}

{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 01/02/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet d'insérer une ligne dans le grid
Mots clefs ... : INSERER LIGNE
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.BAddRowClick(Sender: TObject);
begin
  if LigneCorrecte(GSA.Row,True) then
  begin
    //Mise à jour tob
    vTobVentil.DeTail[GSA.Row - 2].GetLigneGrid(GSA, GSA.Row, ';V_SOUSPLAN1;;V_SOUSPLAN2;;V_SOUSPLAN3;;V_SOUSPLAN4;;V_SOUSPLAN5;;V_TAUXMONTANT;V_TAUXQTE1;V_TAUXQTE2');
    GSA.InsertRow(GSA.Row);
    NewLigne(GSA.Row - 1);
    GSA.Row := GSA.Row - 1;
    GSA.Col := AN_MSect[premier_axe];
  end;
  bModifier := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 01/02/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet d'effacer un ligne du grid
Mots clefs ... : DEL ROW
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.BDelRowClick(Sender: TObject);
var
  i :integer;
begin
  //Vérification
  if (GSA.RowCount <= 3) or (GSA.Row = GSA.RowCount - 1) then
  begin
    GSA.Col := AN_MSect[premier_axe];
    Exit;
  end;
  //Suppression tob ventil
  vTobVentil.Detail[GSA.Row - 2].Free;
  GSA.DeleteRow(GSA.Row);
  //On renumerote les tob fille qui suivent (pour V_NUMEROVENTIL)
  for i := GSA.Row - 2 to vTobVentil.Detail.count - 1 do
  begin
    vTobVentil.Detail[i].PutValue('V_NUMEROVENTIL', i + 1);
    GSA.Cells[AN_MNumL,i + 2] := IntToStr(i + 1);
  end;

  bModifier := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 01/02/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet de récupere les coordonées de la
Suite ........ : souris dans la form
Mots clefs ... : COORDONNÉEES SOURIS
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.GSAMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  GX := X;
  GY := Y;
end;

{b FP 19/10/2005 FQ15725}
procedure TOF_CPVENTILTYPECROIS.GSADblClick(Sender: TObject);
begin
  ClickZoom;
end;
{e FP 19/10/2005 FQ15725}

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Gestion du PopUp menu
Mots clefs ... : POPUP
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.POPSPopup(Sender: TOBject);
begin
  InitPopUp(Ecran);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 02/02/2005
Modifié le ... :   /  /    
Description .. : Procédure pour gérer la sortie de la fiche
Mots clefs ... : CLOSE QUERY ECRAN
*****************************************************************}
procedure TOF_CPVENTILTYPECROIS.EcranCloseQuery(Sender: TObject; var CanClose: Boolean);
Var
  Rep : Integer;
begin
  CanClose := True;
  if bModifier then
  begin
    Rep:=PGIAskCancel('Voulez-vous enregistrer les modifications ?','Ventilations analytiques');
    Case Rep of
      mrYes    : BValiderClick(Nil) ;
      mrNo     : Exit ;
      mrCancel : CanClose := False ;
    end ;
  end;
end;

{JP 28/04/05 : FQ 15781 : pour charger la ventilation type si besoin est
{---------------------------------------------------------------------------------------}
procedure TOF_CPVENTILTYPECROIS.CbVentilTypeClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ven : string;
begin
  Ven := GetControlText('CBVENTILETYPE');
  {Si on a changé de ventilation type, on la charge}
  if SurGeneraux and (sVentilType <> Ven) then
    ChargVentilGrid('TY', Ven);
  sVentilType := Ven;
end;

{JP 17/01/06 : FQ 17268 : Gestion du calcul automatique du solde
{---------------------------------------------------------------------------------------}
procedure TOF_CPVENTILTYPECROIS.DoSolde(Grid: THGrid);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Double;
begin
  {On s'assure que l'on a le droit au concept}
  if not ExJaiLeDroitConcept(TConcept(ccSaisSolde), True) then Exit;

  if Grid.RowCount > 1 then begin
    p := 0;
    {Calcul du total de la saisie}
    for n := 1 to Grid.RowCount - 1 do
      p := p + Valeur(Grid.Cells[11, n]);

    {Si la ligne courante n'a pas de montant, on met le reste (100 - p), dans le cas contraire
     il faut rajouter la valeur de la ligne car elle a été comptabilisée dans p}
    if Valeur(Grid.Cells[11, Grid.Row]) = 0 then p := 100 - p
                                            else p := 100 - p + Valeur(Grid.Cells[11, Grid.Row]);
    {Mise à jour du montant, au bon format}
    Grid.Cells[11, Grid.Row] := StrFMontant(p, 15, 4, '', True);
  end;
end;
{FP 29/12/2005 Provient de CPVENTIL_TOF}
function TOF_CPVENTILTYPECROIS.ExisteSQLSection(Axe, Section: String; ARow: Integer): Boolean;
var
  Clause:       String;
begin
  {b FP 29/12/2005}
  Clause := GetClauseRestrictionAna(Axe, ARow);   {FP 19/04/2006 FQ17728}
  Result := ExisteSQL('SELECT S_SECTION FROM SECTION'+
    ' WHERE S_SECTION="'+Section+'"'+
    '   AND '+Clause);
  {e FP 29/12/2005}
end;

{FP 29/12/2005 Provient de CPVENTIL_TOF}
function TOF_CPVENTILTYPECROIS.LookUpSectionList(Ctrl: TControl; Axe: String; ARow, NumTag: Integer): Boolean;
var
  Clause:       String;
begin
  {b FP 29/12/2005}
  Clause := GetClauseRestrictionAna(Axe, ARow);   {FP 19/04/2006 FQ17728}
  Result := LookUpList(Ctrl, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', Clause, 'S_SECTION', True, NumTag);
  {e FP 29/12/2005}
end;

{b FP 19/04/2006 FQ17728}
function TOF_CPVENTILTYPECROIS.GetClauseRestrictionAna(Axe: String; ARow: Integer): String;
var
  CompteAna:    array[1..MaxAxe] of String;
  i:            Integer;
  s:            String;
  CompteTrouve: Boolean;
begin
  FillChar(CompteAna, sizeof(CompteAna), #0);

  for i:=1 to MAxAxe do   {Vérifie que le dernier axe croisé}
    begin
    s := Trim(GSA.Cells[AN_MSect[i], ARow]);
    if Axes[i] and (s<>'') and ('A'+IntToStr(i)<>Axe) then
      begin
      CompteAna[i] := '"'+s+'"';
      end;
    end;

  {b FP 19/04/2006 FQ17728}
  if FRestriction.IsModelExclu(Self.Compte, 'A'+IntToStr(premier_axe)) then
    begin
    {Vérifie que le dernier axe croisé}
    CompteTrouve := Axe='A'+IntToStr(dernier_axe);
    end
  else
    CompteTrouve := True;
  {e FP 19/04/2006 FQ17728}

  Result := 'S_AXE="' + Axe + '"';
  if CompteTrouve then
    Result := Result +
      ' AND ' + FRestriction.GetClauseCompteAutorise(Self.Compte, Axe, 'SECTION', CompteAna);
end;
{e FP 19/04/2006}

initialization
  registerclasses([TOF_CPVENTILTYPECROIS]);
end.

