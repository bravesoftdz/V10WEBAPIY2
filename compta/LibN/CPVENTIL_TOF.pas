{***********UNITE*************************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 06/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPVENTIL ()
Mots clefs ... : TOF;CPVENTIL;VENTILATIONS
*****************************************************************}
unit CPVENTIL_TOF;

interface

uses StdCtrls,
  Controls, dialogs,
  Classes, SAISUTIL, HCompte, aglinit, vierge, HTB97, ULibAnalytique, SaisComm, Ent1, ParamSoc, lookup, Windows,
  CPSECTION_TOM, ULibEcriture, eSaisAnal, buttons {BitBtn}, SaisComp, hpanel,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul,
  fe_main,
  {$ELSE}
  MaineAGL,
  eMul,
  {$ENDIF}
{$IFNDEF PGIIMMO}
  SousSect,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HMsgBox,
  UTOF,
  menus,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UTOB;

  //CPMODRESTANA_TOM;          {FP 29/12/2005}




type
  ARG_CLASS = class //Pour pouvoir récuperer un ARG_ANA grace a l'objet TheData
  public
    Arguments: ARG_ANA;
  end;

type
  TOF_CPVENTIL = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    //Var
    GSA: THGrid;
    G_GENERAL, E_REFINTERNE, E_LIBELLE: THLabel;
    E_MONTANTECR, S_SOLDE, E_TOTAL, E_RESTE, E_TOTPOURC, E_RESTEPOURC: THNumEdit;
    CVType: THValComboBox;
    BSolde, BVentilType, BZoom, BComplement, BValider, BFerme, HelpBtn, BSauveVentil, BValiderVentilType: TToolBarButton97;
    TempControl, THCode, THLib : THEdit;
    POPS: TPopupMenu;
    BInsertRow, BDeleteRow: TBitBtn;
    PPied, PVentilType, PEntete, PGA : THPanel    ;
    ObjEcr: TOB;
    Arguments: ARG_ANA;
    Pf: string[2]; // Préfixe de la table Y ou ...? TOUJOURS Y !!!
    MontantEcrP, MontantEcrD: double; // Montant écriture
    Sens: Byte; // Ecriture au débit (=1) ou au crédit
    ModeVentil: boolean; // Enregistrement d'une ventilation type en cours
    ModeDevise: boolean; // JAMAIS UTILISE
    ModeOppose: boolean; //saisie en euro
    Axes: array[1..MaxAxe] of Boolean; // Axes ventilables ?
    SensParcours: boolean; //sens de parcours du grid | True de gauche a droite, false pour le reste
    dernier_axe: integer; //numéro du dernier axe ventilable
    premier_axe: integer; //numéro du premier axe ventilable
    //SG6 26.01.05
    GX, GY : integer; //Coordonnées souris ds la form
    EtatBoutonBValider : boolean;
    FRestriction: TRestrictionAnalytique;        // Modèlde de restriction ana FP 29/12/2005}
    //Proc
    procedure DefautEntete;
    procedure ChargeGrid;
    procedure LectureSeule;
    procedure MAJMontant;
    procedure ClickZoom(AfficheFicheSection: Boolean = True);   {FP FQ15725}
    procedure ClickSolde;
    procedure DelRow;
    procedure CreationToBGrid(var tob_result: TOB);
    procedure ChargementToBGrid(vTob : TOB);
    procedure NouvelleLigne;
    procedure ClickZoomSousplan;
    procedure ClickComplement(Lig: integer);
    procedure Preventil(premier_axe : integer ; Cpte : String ; VT : boolean);
    procedure FinVentil;


    //Func
    function PasDerniereLigneVide: boolean;
    function ReturnAxeParam(ACol: integer; typecol: string): integer;
    function ReturnIndiceArray(tableau: array of integer; Valeur: integer): integer;
    function LigneRempli(row: integer): boolean;
    function LookUpSectionList(Ctrl: TControl; Axe: String; ARow, NumTag: Integer): Boolean;    {FP 29/12/2005}
    function ExisteSQLSection(Axe, Section: String; ARow: Integer): Boolean;                     {FP 29/12/2005}
    function GetClauseRestrictionAna(Axe: String; ARow: Integer): String; {FP 19/04/2006 FQ17728}
    //Evts
    procedure BSoldeClick(Sender: TObject);
    procedure BVentilTypeClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure BSauveVentilClick(Sender: TObject);
    procedure CVTypeChange(Sender: TObject);
    procedure GSACellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSACellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSASelectCell(Sender: TObject; ACol, ARow: Longint; var CanSelect: Boolean);
    procedure GSAElipsisClick(Sender: TObject);
    procedure GSADblClick(Sender: TObject);
    procedure EcranKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EcranKeyPress(Sender: TObject; var Key: Char);
    procedure POPSPopup(Sender: TOBject);
    procedure BDeleteRowClick(Sender: TObject);
    procedure BInsertRowClick(Sender: TObject);
    procedure BValiderVentilTypeClick(Sender : TObject);
    //SG6 26.01.05
    procedure GSAMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  end;


procedure eSaisieAnalCroise(var ObjEcr: TOB; Arguments: ARG_ANA);

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  HEnt1;

//Définition de constantes

const
  AN_MNumL: integer = 0;
  AN_MSect: array[1..5] of integer = (1, 3, 5, 7, 9);
  AN_MLib: array[1..5] of integer = (2, 4, 6, 8, 10);
  AN_MPourcent: integer = 11;
  AN_MMontant: integer = 12;



  {***********A.G.L.***********************************************
  Auteur  ...... : SG6
  Créé le ...... : 06/12/2004
  Modifié le ... :   /  /
  Description .. : Point d'entrée de la fiche CPVENTIL
  Mots clefs ... : POINT ENTREE;VENTILATIONS
  *****************************************************************}
procedure eSaisieAnalCroise(var ObjEcr: TOB; Arguments: ARG_ANA);
var
  arguments_class: ARG_CLASS;
begin
  arguments_class := ARG_CLASS.Create;
  arguments_class.Arguments := Arguments;

  TheData := arguments_class;
  TheTob := ObjEcr;

  AGlLanceFiche('CP', 'CPVENTIL', '', '', '');

  Arguments := ARG_CLASS(TheData).Arguments;

end;


procedure TOF_CPVENTIL.OnNew;
begin
  inherited;
end;


procedure TOF_CPVENTIL.OnDelete;
begin
  inherited;
end;


procedure TOF_CPVENTIL.OnUpdate;
begin
  inherited;
end;


procedure TOF_CPVENTIL.OnLoad;
begin
  inherited;
end;


procedure TOF_CPVENTIL.OnArgument(S: string);
var
  i: integer;
  Cpte: string;
begin
  inherited;
  GX := 0;
  GY := 0;
  ModeVentil := False;
  FRestriction := TRestrictionAnalytique.Create;          {FP 29/12/2005}
  ObjEcr := TFVierge(Ecran).LaTof.LaTOB;
  Arguments := ARG_CLASS(TheData).Arguments;
  EtatBoutonBValider := True;

  //Initialisation des contrôles
  GSA := THGrid(getcontrol('GSA', true));
  G_GENERAL := THLabel(getcontrol('G_GENERAL', true));
  E_REFINTERNE := THLabel(getcontrol('E_REFINTERNE', true));
  E_LIBELLE := THLabel(getcontrol('E_LIBELLE', true));
  E_MONTANTECR := THNumEdit(getcontrol('E_MONTANTECR', true));
  S_SOLDE := THNumEdit(getcontrol('S_SOLDE', true));
  E_TOTAL := THNumEdit(getcontrol('E_TOTAL', true));
  E_RESTE := THNumEdit(getcontrol('E_RESTE', true));
  E_TOTPOURC := THNumEdit(getcontrol('E_TOTPOURC', true));
  E_RESTEPOURC := THNumEdit(getcontrol('E_RESTEPOURC', true));
  BSolde := TToolBarButton97(getcontrol('BSolde', true));
  BVentilType := TToolBarButton97(getcontrol('BVentilType', true));
  BZoom := TToolBarButton97(getcontrol('BZoom', true));
  BComplement := TToolBarButton97(getcontrol('BComplement', true));
  BValider := TToolBarButton97(getcontrol('BValider', true));
  BFerme := TToolBarButton97(getcontrol('BFerme', true));
  HelpBtn := TToolBarButton97(getcontrol('HelpBtn', true));
  BSauveVentil := TToolBarButton97(getcontrol('BSauveVentil', true));
  CVType := THValComboBox(getcontrol('CVType', true));
  TempControl := THEdit(getcontrol('Tempcontrol', true));
  POPS := TPopupMenu(getcontrol('POPS', true));
  BInsertRow := TBitBtn(getcontrol('BInsertRow', true));
  BDeleteRow := TBitBtn(getcontrol('BDeleteRow', true));
  //SG6 07.02.05 Ventilation types
  BValiderVentilType := TToolBarButton97(getcontrol('BValiderVentilType', true));
  PPied := THPanel(getcontrol('PPied',true));
  PEntete := THPanel(getcontrol('PEntete',true));
  PVentiltype := THPanel(getcontrol('PVentilType',true));
  THCode := THEdit(getcontrol('THCode',True));
  THLib := THEdit(getcontrol('THLib', True));
  PGA := THPanel(getcontrol('PGA',True));


  //définition des evts sur les contrôles
  BSolde.OnClick := BSoldeClick;
  BVentilType.OnClick := BVentilTypeClick;
  BZoom.OnClick := BZoomClick;
  BComplement.OnClick := BComplementClick;
  BValider.OnClick := BValiderClick;
  BFerme.OnClick := BFermeClick;
  HelpBtn.OnClick := HelpBtnClick;
  BSauveVentil.OnClick := BSauveVentilClick;
  CVType.OnChange := CVTypeChange;
  GSA.OnCellEnter := GSACellEnter;
  GSA.OnCellExit := GSACellExit;
  GSA.OnSelectCell := GSASelectCell;
  GSA.OnElipsisClick := GSAElipsisClick;
  GSA.OnDblClick := GSADblClick;
  Ecran.OnKeyDown := EcranKeyDown;
  Ecran.OnKeyPress := EcranKeyPress;
  POPS.OnPopup := POPSPopup;
  BInsertRow.OnClick := BInsertRowClick;
  BDeleteRow.OnClick := BDeleteRowClick;
  //SG6 26.01.05
  GSA.OnMouseDown := GSAMouseDown;
  //SG6 07.02.05 Ventilations types
  BValiderVentilType.OnClick := BValiderVentilTypeClick;


  //définition de propriétés sur les contrôles
  Ecran.HelpContext := 7244100;
  i := 1;
  while (i < 11) do
  begin
    GSA.Cells[i, 1] := 'Section';
    GSA.Cells[i + 1, 1] := 'Intitulé';
    inc(i, 2);
  end;
  GSA.Cells[0, 2] := '1';
  GSA.TwoColors := true;
  Focuscontrole(GSA, true);

  //Paramètrage des axes ventilables et calcul du dernier axe ventilable
  Sens := GetSens(ObjEcr);
  dernier_axe := 0;
  premier_axe := 0;
  for i := 1 to MaxAxe do
  begin
    Axes[i] := GetParamSocSecur('SO_VENTILA' + IntToStr(i),False);
    if Axes[i] then
    begin
      dernier_axe := i;
      if (premier_axe = 0) then premier_axe := i;
    end;
  end;

  if Arguments.QuelEcr = EcrGen then
  begin
    Pf := 'Y';
    GSA.TypeSais := tsAnal;
    Cpte := Arguments.CC.General;
    MontantEcrP := GetMontant(ObjEcr);
    MontantEcrD := GetMontantDev(ObjEcr);
    Sens := GetSens(ObjEcr);
  end;
  DefautEntete;
  ModeVentil := False;
  ColorOpposeEuro(GSA, ModeOppose, ModeDevise);
  AffecteGrid(GSA, Arguments.Action);

  //Préventilsation
  //Scénario

  ChargeGrid;
  case Arguments.Action of
    taConsult: LectureSeule;
  end;

  //calcule et mise à jour des montants
  MAJMontant;
end;


procedure TOF_CPVENTIL.OnClose;
var
  arguments_class: ARG_CLASS;
begin
  arguments_class := ARG_CLASS.Create;
  arguments_class.Arguments := Arguments;

  TheData := arguments_class;
  //Free des objets
  //FreeAndNil(TheData);
  FreeAndNil(FRestriction);                   {FP 29/12/2005}
  Ecran.Close;
  inherited;
end;


procedure TOF_CPVENTIL.OnDisplay();
begin
  inherited;
end;


procedure TOF_CPVENTIL.OnCancel();
begin
  inherited;
  Close;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BSolde
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.BSoldeClick(Sender: TObject);
begin
  ClickSolde;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BVentilType
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.BVentilTypeClick(Sender: TObject);
begin
  if Arguments.Action = taConsult then Exit;
  if not CVType.Enabled then Exit;
  if Arguments.DernVentilType = '' then Exit;
  CVType.Value := Arguments.DernVentilType;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BZoom
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.BZoomClick(Sender: TObject);
begin
  ClickZoom;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BComplement
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.BComplementClick(Sender: TObject);
begin
  ClickComplement(GSA.Row);
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BValide
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.BValiderClick(Sender: TObject);
var
  i, j: Integer;
begin
  MajMontant;

  if not (BValider.Enabled) then Exit;
  {b FP 29/12/2005: Vérifie la validité de la ligne}
  for i := 2 to GSA.Rowcount - 1 do
    begin
    if LigneRempli(i) then
    //YMO FQ13244 23/06/06 Vérification des sections à la sortie
     for j := 1 to MAXAXE do
     if Axes[j] and (not ExisteSQLSection('A'+IntToStr(j), GSA.Cells[AN_MSect[j], i], i)) then
     begin
      PGIError(TraduireMemoire('Ligne de ventilation incorrecte'), Ecran.Caption);
      GSA.Row := i;
      GSA.Col:=AN_MSect[j];
      Exit;
     end;
    end;
  {e FP 29/12/2005}
  CreationToBGrid(ObjEcr);
  TFVierge(Ecran).LaTof.LaTob := ObjEcr;
  Close;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BAbandon
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.BFermeClick(Sender: TObject);
begin
  Close;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BAide
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Ecran);
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnClick sur le bouton BSauveVentil
Mots clefs ... : CLICK
*****************************************************************}
procedure TOF_CPVENTIL.BSauveVentilClick(Sender: TObject);
begin
  if Arguments.Action=taConsult then Exit ;
  if GSA.RowCount<=2 then Exit ;
  if E_TOTPOURC.Value=0 then Exit ;
  if Not ExJaiLeDroitConcept(TConcept(ccSaisCreatVentil),True) then Exit ;

  // Affichage du panel d'enregistrement de la ventilatin type
  PVentilType.Left := GSA.Left+(GSA.Width-PVentilType.Width) div 2 ;
  ModeVentil := True ;
  PEntete.Enabled := False ;
  PPied.Enabled := False ;
  GSA.Enabled := False ;
//  PGA.Enabled := False ;
  BValider.Enabled := False;
  BFerme.Enabled := False;
  HelpBtn.Enabled := False;
  BZoom.Enabled := False;
  BComplement.Enabled := False;
  BVentilType.Enabled := False;
  BSolde.Enabled := False;
  PVentilType.Visible := True;
  THCode.SetFocus ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure Evenementielle
Suite ........ : Evt OnChange sur le controle CVType
Mots clefs ... : CHANGe
*****************************************************************}
procedure TOF_CPVENTIL.CVTypeChange(Sender: TObject);
begin
  if (CVType.Value = '') then exit ;
  Preventil(premier_axe,CVType.Value,True);
  Arguments.DernVentilType := CVType.Value
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Permet d'initilisaliser les certains contrôles de la fiche lors de
Suite ........ : son ouverture
Mots clefs ... : INIT, CONTROLE
*****************************************************************}
procedure TOF_CPVENTIL.DefautEntete;
begin
  if Arguments.QuelEcr = EcrGen then G_GENERAL.Caption := ObjEcr.GetValue('E_GENERAL');
  E_REFINTERNE.Caption := ObjEcr.GetValue('E_REFINTERNE');
  E_LIBELLE.Caption := ObjEcr.GetValue('E_LIBELLE');
  E_MONTANTECR.Value := GetMontantDev(ObjEcr);
  E_TOTAL.Value := 0;
  E_RESTE.Value := E_MONTANTECR.Value;
  E_RESTEPOURC.Value := 1;
  E_TOTPOURC.Value := 0;
  E_MONTANTECR.DEBIT := (Sens = 1);
  ChangeFormatDevise(Ecran, Arguments.DEV.Decimale, Arguments.DEV.Symbole);
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure qui permet de paramètrer les controles lorsque
Suite ........ : l'on est en lecture seule
Mots clefs ... : READONLY; FICHE
*****************************************************************}
procedure TOF_CPVENTIL.LectureSeule;
begin
  E_MONTANTECR.Enabled := False;
  BSolde.Enabled := False;
  BVentilType.Enabled := False;
  BSauveVentil.Enabled := False;
  CVType.Enabled := False;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Procédure qui permet de charger le grid de ventilation
Mots clefs ... : CHARGER; GRID
*****************************************************************}
procedure TOF_CPVENTIL.ChargeGrid;
var
  i, j: integer;
  tobtmp: TOB;
begin
  //On cache les axes à cacher
  j := 1;
  for i := 1 to MaxAxe do
  begin
    if not Axes[i] then
    begin
      GSA.ColWidths[j] := -1;
      GSA.ColWidths[j + 1] := -1;
      GSA.ColLengths[j] := -1;
      GSA.ColLengths[j + 1] := -1;
    end;
    Inc(j, 2);
  end;

{  if MaxAxe < 5 then
  begin
    for i := MaxAxe + 1 to 5 do
    begin
      if not Axes[i] then
      begin
        GSA.ColWidths[j] := -1;
        GSA.ColWidths[j + 1] := -1;
        GSA.ColLengths[j] := -1;
        GSA.ColLengths[j + 1] := -1;
      end;
      Inc(j, 2);
    end;
  end;} //SG6 22/12/2004 ventilation croisaxe pour ccs5 pack avancé et dans ce cas MaxAxe = 5 !

  if ObjEcr.Detail[premier_axe - 1].Detail.Count = 0 then
  begin
    tobtmp := Tob.Create('ANALYTIQ', nil, -1);
    CPutDefautAna(tobtmp);
    InitCommunObjAnalNEW(ObjEcr, tobtmp);
    tobtmp.AddChampSupValeur('OLDDEBIT', 0);
    tobtmp.AddChampSupValeur('OLDCREDIT', 0);
    tobtmp.AddChampSupValeur('RATIO', 0);
    tobtmp.AddChampSupValeur('CONVERTFRANC', 0);
    tobtmp.AddChampSupValeur('CONVERTEURO', 0);
    GSA.Objects[0, 2] := tobtmp;
  end
  else
  begin
    ChargementToBGrid(ObjEcr);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 10/12/2004
Modifié le ... :   /  /
Description .. : Gestion de l'entrée dans une cellule
Mots clefs ... : ROWENTER
*****************************************************************}
procedure TOF_CPVENTIL.GSACellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
  if GSA.ColLengths[GSA.Col] = -1 then exit;
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
procedure TOF_CPVENTIL.GSACellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
var
  iAxe   : integer;
  sAxe   : string;
  NumTag : Integer;
begin
  if csDestroying in Ecran.ComponentState then exit;
  SensParcours := (ACol < GSA.Col);
  if Arguments.Action = taConsult then exit;
  if GSA.ColLengths[GSA.Col] = -1 then exit;

  //Si Section renseigné on vérifie qu'elle existe > autrement Cancel et lookup
  if ReturnIndiceArray(AN_MSect, ACol) <> -1 then
  begin
    //Calcul de l'axe
    if ((ARow <> GSA.Row) and ((GSA.Cells[AN_MMontant, ARow] = '') or (GSA.Cells[AN_MPourcent, ARow] = ''))) then
    begin
      Cancel := true;
      Exit;
    end;
    iAxe := ReturnAxeParam(ACol, 'tSection');
    {b FP 29/12/2005 if not (ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="' + GSA.Cells[ACol, Arow] + '" AND S_AXE="A' + IntToStr(iAxe) + '"')) then}
    if not ExisteSQLSection('A'+IntToStr(iAxe), GSA.Cells[ACol, Arow], ARow) then
    {e FP 29/12/2005}
        begin
      TempControl.Text := GSA.Cells[ACol, ARow];
      {JP 18/01/06 : FQ 17261 : gestion du tag DispatchTT pour pouvoir créer des sections sur le LookUp}
      NumTag := 71780 + iAxe;
      {b FP 29/12/2005 if LookUpList(TempControl, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE="A' + IntToStr(iAxe) + '"', 'S_SECTION', True, NumTag) then}
      if LookUpSectionList(TempControl, 'A' + IntToStr(iAxe), ARow, NumTag) then
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
      if GSA.Row = GSA.RowCount - 1 then
      begin
        NouvelleLigne;
      end;
    end;


  end;

  //Si cellule pourcentage
  if (ACol = AN_MPourcent) then
  begin
    if not (IsNumeric(GSA.Cells[ACol, ARow])) then
    begin
      GSA.Cells[ACol, ARow] := '';
      GSA.Cells[ACol + 1, ARow] := '';
      if (ARow <> GSA.Row) then
      begin
        Cancel := True;
        Exit;
      end;
    end
    else
    begin
      GSA.Cells[AN_MMontant, ARow] := strfmontant(Arrondi(Valeur(GSA.Cells[ACol, ARow]) * MontantEcrD / 100, Arguments.DEV.Decimale), 15, 2, '', true);
      GSA.Cells[AN_MPourcent, ARow] := strfmontant(Arrondi(Valeur(GSA.Cells[AN_MMontant, ARow]) * 100 / MontantEcrD, Arguments.DEV.Decimale), 15, 4, '', true);
    end;
    MAJMontant;
  end;

  //si cellule montant
  if (ACol = AN_MMontant) then
  begin
    if not (IsNumeric(GSA.Cells[ACol, ARow])) then
    begin
      GSA.Cells[ACol, ARow] := '';
      if (ARow <> GSA.Row) or SensParcours then
      begin
        Cancel := true;
        exit;
      end;
    end
    else
    begin
      GSA.Cells[AN_MPourcent, ARow] := strfmontant(Arrondi(Valeur(GSA.Cells[ACol, ARow]) * 100 / MontantEcrD, Arguments.DEV.Decimale), 15, 4, '', true);
    end;
    MAJMontant;
  end;
end;


procedure TOF_CPVENTIL.GSASelectCell(Sender: TObject; ACol, ARow: Longint; var CanSelect: Boolean);
begin
  GSA.ElipsisButton := (ReturnIndiceArray(AN_MSect, ACol) <> -1);
end;

procedure TOF_CPVENTIL.GSAElipsisClick(Sender: TObject);
var
  iAxe   : integer;
  NumTag : Integer;
begin
  {JP 21/02/06 : FQ 17261 : gestion du tag DispatchTT pour pouvoir créer des sections sur le LookUp}
  iAxe := ReturnAxeParam(GSA.Col, 'tSection');
  NumTag := 71780 + iAxe;
  {b FP 29/12/2005 LookUpList(GSA, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE="A' + IntToStr(ReturnAxeParam(GSA.Col, 'tSection')) + '"', 'S_SECTION', True, 0);}
  LookUpSectionList(GSA, 'A'+IntToStr(ReturnAxeParam(GSA.Col, 'tSection')), GSA.Row, NumTag);
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
function TOF_CPVENTIL.ReturnAxeParam(ACol: integer; typecol: string): integer;
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
Créé le ...... : 15/12/2004
Modifié le ... :   /  /
Description .. : Gestion du double click dans le grid
Mots clefs ... : DBL CLICK
*****************************************************************}
procedure TOF_CPVENTIL.GSADblClick(Sender: TObject);
begin
  ClickZoom;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 06/12/2004
Modifié le ... : 30/03/2007
Description .. : 
Suite ........ : SBO 30/03/2007 : FQ19865
Suite ........ : Pb d'arrondi sur calcul du solde >> impossible de valider 
Suite ........ : dans certains cas.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPVENTIL.MAJMontant;
var
  i: integer;
  totmtr, totpourc: double;
  lBoSoldeOk : Boolean ;
begin
  totpourc := 0;
  totmtr := 0;

  for i := 2 to GSA.RowCount - 1 do
  begin
    if GSA.Cells[AN_MMontant, i] <> '' then
      totmtr := Arrondi( totmtr + Valeur(GSA.Cells[AN_MMontant, i]), Arguments.DEV.Decimale );
    if GSA.Cells[AN_MPourcent, i] <> '' then
      totpourc :=  Arrondi( totpourc + Valeur(GSA.Cells[AN_MPourcent, i]), 4 );
  end;
  E_TOTAL.Value      := totmtr;
  E_TOTPOURC.Value   := totpourc / 100;
  E_RESTE.Value      := Arrondi( E_MontantEcr.Value - totmtr, Arguments.DEV.Decimale );
  E_RESTEPOURC.Value := 1 - totpourc / 100;

  lBoSoldeOk         := ( Arrondi(totmtr - E_MONTANTECR.Value, Arguments.DEV.Decimale ) = 0 )
                        and ( Arrondi(100 - totpourc, 4) = 0 ) ;
  BValider.Enabled   := lBoSoldeOk;
  EtatBoutonBValider := lBoSoldeOk;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 13/12/2004
Modifié le ... :   /  /
Description .. : Gestion des raccourci claviers
Mots clefs ... : RACCOURCI CLAVIER
*****************************************************************}
procedure TOF_CPVENTIL.EcranKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  vide, OkGrid: boolean;
begin
  //SG6 07.02.05 Gestion ventilations types
  if ((ModeVentil) and (Key<>VK_F10) and (Key<>VK_ESCAPE)) then Exit ;

  OkGrid := (Ecran.ActiveControl = GSA);
  vide := (shift = []);
  case Key of
    VK_F5: if ((OkGrid) and (vide)) then
      begin
        ClickZoom;
        Key := 0;
      end
      else if ((OkGrid) and ((shift = [ssCtrl]) or (shift = [ssAlt]))) then
      begin
        ClickZoomSousplan;
      end;

    VK_F6: if ((OkGrid) and (vide)) then
      begin
        ClickSolde;
        Key := 0;
      end;
    VK_ESCAPE:
      begin
        if ModeVentil then FinVentil
        else Close;
        Key := 0;
      end;
    VK_RETURN: Key := VK_TAB;
    VK_DELETE: if ((OkGrid) and (shift = [ssCtrl])) then
      begin
        DelRow;
        Key := 0;
      end;
    VK_F10:
      begin
        if ModeVentil then BValiderVentilTypeClick(nil)
        else BValider.Click;  //BValiderClick(nil);  {FP 16/06/2006 FQ17337}
        Key := 0;                                    {FP 16/06/2006 FQ17337}
      end;
    {=}187: if ((OkGrid) and (Vide)) then
      begin
        Key := 0;
        ClickSolde;
      end;
    VK_INSERT: if ((OkGrid) and (Vide)) then
      begin
        Key := 0;
        BInsertRowClick(nil);
      end;
    {AC}67:
      begin
        if Shift = [ssAlt] then
        begin
          Key := 0;
          ClickComplement(GSA.Row);
        end;
      end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 13/12/2004
Modifié le ... :   /  /
Description .. : zoom suivant la cell séléctionné du grid
Mots clefs ... : ZOOM
*****************************************************************}
procedure TOF_CPVENTIL.ClickZoom(AfficheFicheSection: Boolean = True);   {FP FQ15725}
var
  cell_section: boolean;
  iAxe: integer;
  ACol: longint;
  ARow: longint;
  NumTag : Integer;
begin
  //Sections
  //si section existe on ouvre la fénétre de la section autrement lookuplist sur section
  if Arguments.Action = taConsult then
  begin
    GSA.MouseToCell(GX,GY,ACol,ARow);
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
    if not ExisteSQLSection('A'+IntToStr(iAxe), GSA.Cells[Acol, ARow], ARow) then
    {e FP 29/12/2005}
    begin
      //SG6 03.03.05 FQ 15383
      TempControl.Text := GSA.Cells[ACol, ARow];
      {JP 21/02/06 : FQ 17261 : gestion du tag DispatchTT pour pouvoir créer des sections sur le LookUp}
      NumTag := 71780 + iAxe;
      {b FP 29/12/2005 if LookUpList(TempControl, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', 'S_AXE="A' + IntToStr(iAxe) + '"', 'S_SECTION', True, 0) then}
      if LookUpSectionList(TempControl, 'A' + IntToStr(iAxe), ARow, NumTag) then
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

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 13/12/2004
Modifié le ... :   /  /
Description .. : Calcul du solde
Mots clefs ... : SOLDE
*****************************************************************}
procedure TOF_CPVENTIL.ClickSolde;
var
  i: integer;
begin
  //Checkage si toutes les sections pour la ligne en cours renseigné et si ell existent
  GSA.Cells[AN_MMontant, GSA.Row] := strfmontant(0, 15, 2, '', true);
  GSA.Cells[AN_MPourcent, GSA.Row] := strfmontant(0, 15, 4, '', true);
  MajMontant;

  for i := 1 to MaxAxe do
  begin
    if Axes[i] then
    begin
      {b FP 29/12/2005 if not (ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="' + GSA.Cells[AN_MSect[i], GSA.Row] + '" AND S_AXE="A' + IntToStr(i) + '"')) then}
      if not ExisteSQLSection('A'+IntToStr(i), GSA.Cells[AN_MSect[i], GSA.Row], GSA.Row) then
      begin
        exit;
      end;
    end;
  end;
  GSA.Cells[AN_MMontant, GSA.Row] := strfmontant(E_RESTE.Value, 15, 2, '', true);
  GSA.Cells[AN_MPourcent, GSA.Row] := strfmontant(E_RESTEPOURC.Value * 100, 15, 4, '', true);
  MajMontant;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 13/12/2004
Modifié le ... :   /  /
Description .. : Procédure qui permet d'effacer une ligne entière
Mots clefs ... : SUPPRESSION LIGNE
*****************************************************************}
procedure TOF_CPVENTIL.DelRow;
var
  i: integer;
  tobanal: TOB;
begin
  if Arguments.Action = taConsult then exit;

  if (GSA.Row < 2) then exit;
  GSA.CacheEdit;
  if (GSA.Row = 2) and (GSA.RowCount - 1 = GSA.Row) then
  begin
    for i := 1 to GSA.ColCount - 1 do
    begin
      GSA.Cells[i, 2] := '';
      GSA.Objects[0, 2].Free;
      tobanal := TOB.Create('ANALYTIQ', nil, -1);
      CPutDefautAna(tobanal);
      InitCommunObjAnalNEW(ObjEcr, tobanal);
      GSA.Objects[0, 2] := tobanal;
    end;
  end
  else
  begin
    GSA.Objects[0, GSA.Row].Free;
    GSA.DeleteRow(GSA.Row);
    //Recalcul des numeros de lignes
    for i := 2 to GSA.RowCount - 1 do
    begin
      GSA.Cells[0, i] := IntToStr(i - 1);
    end;

    //positionnement dans la première colonne non caché
    for i := 1 to GSA.ColCount - 1 do
    begin
      if GSA.ColLengths[i] <> -1 then
      begin
        GSA.Col := i;
        break;
      end
    end;
  end;
  GSA.MontreEdit;
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
function TOF_CPVENTIL.ReturnIndiceArray(tableau: array of integer; Valeur: integer): integer;
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
Créé le ...... : 13/12/2004
Modifié le ... :   /  /
Description .. : Procédure qui permet de créer la tob à partir du grid des
Suite ........ : ventilations croisaxes
Mots clefs ... : TOB, GRID
*****************************************************************}
procedure TOF_CPVENTIL.CreationToBGrid(var tob_result: TOB);
var
  tobfille: TOB;
  tobaxe: TOB;
  i, j: integer;
  num_ventil: integer;
  XC, XD: double;
begin
  num_ventil := 1;
  tobaxe := tob_result.Detail[premier_axe - 1];
  //Effacement pour tous les axes des ecritures analytiques eventuelles
  for i := 1 to MaxAxe do
  begin
    while tob_result.Detail[i - 1].Detail.Count <> 0 do
    begin
      tob_result.Detail[i - 1].Detail[0].Free;
    end;
  end;

  for i := 2 to GSA.Rowcount - 1 do
  begin
    if LigneRempli(i) then
    begin
      //On recup la tob
      tobfille := TOB(GSA.Objects[0, i]);
      tobfille.ChangeParent(tobaxe, -1);

      with tobfille do
      begin
        AddChampSupValeur('OLDDEBIT', 0);
        AddChampSupValeur('OLDCREDIT', 0);
        AddChampSupValeur('RATIO', 0);
        AddChampSupValeur('CONVERTFRANC', 0);
        AddChampSupValeur('CONVERTEURO', 0);
        PutValue('Y_AXE', 'A' + IntToStr(premier_axe));
        PutValue('Y_SECTION', GSA.Cells[AN_MSect[premier_axe], i]);
        PutValue('Y_NUMVENTIL', num_ventil);
        Inc(num_ventil);
        PutValue('Y_POURCENTAGE', Valeur(GSA.Cells[AN_MPourcent, i]));

        //Renseignements des sous plans
        for j := 1 to MaxAxe do
        begin
          if Axes[j] then
          begin
            PutValue('Y_SOUSPLAN' + InttoStr(j), GSA.Cells[AN_MSect[j], i]);
          end;
        end;

        if Sens = 1 then
        begin
          XD := Valeur(GSA.Cells[AN_MMontant, i]);
          XC := 0;
        end
        else
        begin
          XD := 0;
          XC := Valeur(GSA.Cells[AN_MMontant, i]);
        end;
        CSetMontants(tobfille, XD, XC, Arguments.DEV, False);
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 13/12/2004
Modifié le ... : 03/08/2005
Description .. : Fonction qui renvoie True si toutes les cellules pour une
Suite ........ : ligne donnée sont non vides (ils verifient pas les colonnes
Suite ........ : avec ColWidth à -1)
Suite ........ : 
Suite ........ : SBO 03/08/2005 : le test sur le ColWidth un peu cavalier a 
Suite ........ : été remplacé par un test sur le tableau de booléens Axes 
Suite ........ : représentant la paramétrage de ventilation CROISAXE...
Mots clefs ... : LIGNE; VIDE
*****************************************************************}
function TOF_CPVENTIL.LigneRempli(row: integer): boolean;
var
  i: integer;
begin
  result := true;

  // Vérification des sections
  for i := 1 to MAXAXE do
    begin
    if Axes[i] and (GSA.Cells[AN_MSect[1], row] = '') then
      begin
      result := false;
      Exit ;
      end
    end;

  // Vérification des montants
  if ((Valeur(GSA.Cells[AN_MMontant, row]) = 0) or (Valeur(GSA.Cells[AN_MPourcent, row]) = 0)) then
    result := false;

end;



{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 14/12/2004
Modifié le ... :   /  /
Description .. : Procedure qui permet de charger dans la tob les écritures
Suite ........ : analytiques
Mots clefs ... : CHARGEMENT, TOB, GRID
*****************************************************************}
procedure TOF_CPVENTIL.ChargementToBGrid(vTob : TOB);
var
  tobaxe: TOB;
  tobanal: TOB;
  row, i, j: integer;
  sTmp: string;

begin

  tobaxe := vTob.Detail[premier_axe - 1];
  if tobaxe = nil then exit;
  row := 1;
  for i := 0 to tobaxe.Detail.Count - 1 do
  begin
    // A Utiliser si calcul solde sur section (comme ventilations mono axe)
    //tobaxe.Detail[i].PutValue ('OLDDEBIT', tobaxe.Detail[i].GetValue( 'Y_DEBIT' ) ) ;
    //tobaxe.Detail[i].PutValue ('OLDCREDIT', tobaxe.Detail[i].GetValue( 'Y_CREDIT' ) ) ;

    row := row + 1;
    if GSA.RowCount = row then GSA.RowCount := GSA.RowCount + 1;
    tobanal := TOB.Create('', nil, -1);
    tobanal.Dupliquer(tobaxe.Detail[i], true, true);
    GSA.Objects[0, row] := tobanal;
    GSA.Cells[0, row] := IntToStr(row - 1);

    //Remplissage des sections
    for j := 1 to MaxAxe do
    begin
      if Axes[j] then
      begin
        if j = 1 then sTmp := '' else sTmp := IntToStr(j);
        GSA.Cells[AN_MSect[j], row] := tobanal.GetValue('Y_SOUSPLAN' + IntToStr(j));
        GSA.Cells[AN_MLib[j], row] := rechdom('TZSECTION' + sTmp, tobanal.GetValue('Y_SOUSPLAN' + IntToStr(j)), false);
      end;
    end;

    //Remplissage montant et pourcentage
    if Sens = 1 then sTmp := strfmontant(tobanal.GetValue('Y_DEBITDEV'), 15, 2, '', true) else stmp := strfmontant(tobanal.Getvalue('Y_CREDITDEV'), 15, 2, '', true);

    GSA.Cells[AN_MMontant, row] := sTmp;
    GSA.Cells[AN_MPourcent, row] := strfmontant(tobanal.Getvalue('Y_POURCENTAGE'), 15, 4, '', true);
  end;

  //NouvelleLigne;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 14/12/2004
Modifié le ... :   /  /
Description .. : Création d'une nouvelle ligne dans le grid
Mots clefs ... : NEW LIGNE
*****************************************************************}
procedure TOF_CPVENTIL.NouvelleLigne;
var
  tobfille: TOB;
begin
  GSA.RowCount := GSA.RowCount + 1;
  GSA.Cells[0, GSA.RowCount - 1] := IntToStr(GSA.RowCount - 2);
  tobfille := TOB.Create('ANALYTIQ', nil, -1);
  CPutDefautAna(tobfille);
  InitCommunObjAnalNEW(ObjEcr, tobfille);
   // Pour Calcul Solde (qd  calcul mis en place)s
  {  tobfille.AddChampSupValeur('OLDDEBIT', 0);
    tobfille.AddChampSupValeur('OLDCREDIT', 0);
    tobfille.AddChampSupValeur('RATIO', 0);
    tobfille.AddChampSupValeur('CONVERTFRANC', 0);
    tobfille.AddChampSupValeur('CONVERTEURO', 0);}
  GSA.Objects[0, GSA.Rowcount - 1] := tobfille;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Gestion du PopUp menu
Mots clefs ... : POPUP
*****************************************************************}
procedure TOF_CPVENTIL.POPSPopup(Sender: TOBject);
begin
  InitPopUp(Ecran);
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Gère la suppression d'une ligne du grid (pour popup menu)
Mots clefs ... : SUPPRESSION LIGNE GRID
*****************************************************************}
procedure TOF_CPVENTIL.BDeleteRowClick(Sender: TObject);
begin
  DelRow;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Gère l'ajout d'une ligne du grid (pour popup menu)
Mots clefs ... : AJOUT LIGNE GRID
*****************************************************************}
procedure TOF_CPVENTIL.BInsertRowClick(Sender: TObject);
begin
  if Arguments.Action = taConsult then exit;
  if (LigneRempli(GSA.Row) and PasDerniereLigneVide) then
  begin
    NouvelleLigne;
    GSA.Col := 1;
    GSA.Row := GSA.RowCount - 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Fonction qui permet de voir si la dernière du ligne du Grid
Suite ........ : est vide
Mots clefs ... : DERNIER LIGNE VIDE
*****************************************************************}
function TOF_CPVENTIL.PasDerniereLigneVide: boolean;
var
  i: integer;
begin
  result := True;
  for i := 1 to GSA.Colcount - 1 do
  begin
    if (GSA.ColLengths[i] <> -1) and (GSA.Cells[i, GSA.RowCount - 1] = '') then
    begin
      result := result and True;
    end
    else if (GSA.ColLengths[i] <> -1) and (GSA.Cells[i, GSA.RowCount - 1] <> '') then
    begin
      result := result and False;
      break;
    end;
  end;
  result := not (result);
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Gestion des raccourci claviers
Mots clefs ... : ECRAN KEYPRESS
*****************************************************************}
procedure TOF_CPVENTIL.EcranKeyPress(Sender: TObject; var Key: Char);
begin
  if not GSA.SynEnabled then Key := #0 else
    if Key = #127 then Key := #0 else
    if Key = '=' then Key := #0;
end;


procedure TOF_CPVENTIL.ClickZoomSousplan;
var
   iAxe: integer;
   Sect: string;
   fb: TFichierBase;
begin
  if ReturnIndiceArray(AN_MSect, GSA.Col) = -1 then exit;
  iAxe := ReturnAxeParam(GSA.Col, 'tSection');
  fb := AxeTofb('A' + IntToStr(iAxe));
  Sect := ChoisirSSection(fb, GSA.Cells[GSA.Col, GSA.Row], FALSE, Arguments.Action);
  if Sect <> '' then GSA.Cells[GSA.Col, GSA.Row] := sect;
  ClickZoom(False);             {FP FQ15725}
end;


procedure TOF_CPVENTIL.ClickComplement(Lig: integer);
{$IFNDEF GCGC}
var
  ModBN: Boolean;
  LeEcr: TTypeEcr;
  RC: R_COMP;
  {$ENDIF}
begin
  {$IFNDEF GCGC}
  if not LigneRempli(Lig) then Exit;
  LeEcr := EcrAna;
  RC.StLibre := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
  RC.StComporte := '--XXXXXXXX';
  RC.Conso := True;
  RC.Attributs := False;
  RC.MemoComp := nil;
  RC.Origine := -1;
  RC.DateC := 0;
  RC.TOBCompl := nil ;  
  SaisieComplement(TobM(GSA.Objects[0,Lig]), LeEcr, Arguments.Action, ModBN, RC);
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 26/01/2005
Modifié le ... :   /  /    
Description .. : Procedure qui permet de récupere les coordonées de la 
Suite ........ : souris dans la form
Mots clefs ... : COORDONNÉEES SOURIS
*****************************************************************}
procedure TOF_CPVENTIL.GSAMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  GX := X;
  GY := Y;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 07/02/2005
Modifié le ... :   /  /    
Description .. : Procédure qui permet de ventiler suivant une ventil type
Mots clefs ... : VENTILATIONS TYPES
*****************************************************************}
procedure TOF_CPVENTIL.Preventil(premier_axe : integer ; Cpte : String ; VT : boolean);
var
  Q : TQuery;
  i, NumV : integer;
  vTobAna, vTob : TOB;
  Ax, StV : string;
  PM, P1, P2, XP, XD, TotP, TotMP, TotMD, DiffP, DiffD : double;
  {b FP 29/12/2005}
  CompteAna:  array[1..MaxAxe] of String;
  {e FP 29/12/2005}
  {b FP 19/04/2006 FQ17725}
   QrySansFiltre:   String;
   OrderBy:         String;
  {e FP 19/04/2006}
begin
  if Arguments.Action = taConsult then Exit;

  if ((Not VT) and (GSA.RowCount>2) and LigneRempli(2)) then Exit;

  FillChar(CompteAna, sizeof(CompteAna), #0);       {FP 29/12/2005}
  QrySansFiltre := '';                              {FP 19/04/2006 FQ17725}
  if ((Arguments.QuelEcr = EcrGen) and (Not VT)) then
  begin
{b FP 19/04/2006 FQ17725: Requête sans filtre sur la restriction analytique}
    QrySansFiltre := 'SELECT * FROM VENTIL WHERE V_NATURE="GE'+IntToStr(premier_axe)+'" AND V_COMPTE="'+Cpte+'" ';
    StV := QrySansFiltre
      {b FP 29/12/2005}
       +' AND ' + FRestriction.GetClauseCompteAutorise(
         G_GENERAL.Caption, 'A'+IntToStr(premier_axe), 'VENTIL', CompteAna);
      {e FP 29/12/2005}
    OrderBy := ' ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
    QrySansFiltre := QrySansFiltre + OrderBy;
    StV           := StV           + OrderBy;
{e FP 19/04/2006}
  end
  else if ((Arguments.QuelEcr = EcrBud) and (Not VT)) then
  begin
    StV:='SELECT * FROM VENTIL WHERE V_NATURE="BU'+IntToStr(premier_axe)+'" AND V_COMPTE="'+Cpte+'" '
         +'ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
  end
  else
  begin
{b FP 19/04/2006 FQ17725: Requête sans filtre sur la restriction analytique}
    QrySansFiltre:='SELECT * FROM VENTIL WHERE V_NATURE="TY'+IntToStr(premier_axe)+'" AND V_COMPTE="'+Cpte+'" ';
    StV := QrySansFiltre
        {b FP 29/12/2005}
         +' AND ' + FRestriction.GetClauseCompteAutorise(
           G_GENERAL.Caption, 'A'+IntToStr(premier_axe), 'VENTIL', CompteAna);
        {e FP 29/12/2005}
    OrderBy := ' ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
    QrySansFiltre := QrySansFiltre + OrderBy;
    StV           := StV           + OrderBy;
{e FP 19/04/2006}
  end;

  Q := OpenSQL(StV, True);

  //init var
  TotP := 0 ;
  TotMP := 0 ;
  TotMD := 0 ;

  //Création tob mère
  vTob := TOB.Create('',nil,-1);
  TOB.Create('A1',vTob,-1);
  TOB.Create('A2',vTob,-1);
  TOB.Create('A3',vTob,-1);
  TOB.Create('A4',vTob,-1);
  TOB.Create('A5',vTob,-1);

  FRestriction.VerifModelVentil(G_GENERAL.Caption, Q, 'VENTIL', 'A'+IntToStr(premier_axe), QrySansFiltre, VT); {FP 29/12/2005+19/04/2006 FQ17725}
  if (not Q.eof) {or (not OkVentil) FP 29/12/2005+19/04/2006 FQ17725} then
  begin
    for i := 2 to GSA.RowCount - 1 do
    begin
      GSA.Objects[0,i].Free;
    end;
    GSA.RowCount := 3;

  end;
  NumV:=0; //FP 29/12/2005
  while (not Q.Eof) {and OkVentil FP 29/12/2005+19/04/2006 FQ17725}  do
  begin
    //Calcul préliminaire
    Ax := 'A' + Chr(48 + premier_axe);
    Inc(NumV); //FP 29/12/2005 NumV := Q.FindField('V_NUMEROVENTIL').AsInteger;
    PM := Q.FindField('V_TAUXMONTANT').AsFloat;
    P1 := Q.FindField('V_TAUXQTE1').AsFloat;
    P2 := Q.FindField('V_TAUXQTE2').Asfloat;
    XP := Arrondi(PM * MontantEcrP / 100.0, V_PGI.OkDecV);
    XD := Arrondi(PM * MontantEcrD / 100.0, Arguments.DEV.Decimale);
    TotP := TotP + PM;
    TotMP := TotMP + XP;
    TotMD := TotMD + XD;

    //Gestion Tob ana
    vTobAna := TOB.Create('ANALYTIQ',vTob.Detail[premier_axe - 1],-1);
    CPutDefautAna(vTobAna);
    InitCommunObjAnalNEW(ObjEcr, vTobAna);

    vTobAna.PutValue('Y_TOTALQTE1', ObjEcr.GetValue('E_QTE1'));
    vTobAna.PutValue('Y_TOTALQTE2', ObjEcr.GetValue('E_QTE2'));
    vTobAna.PutValue('Y_SECTION', Q.FindField('V_SECTION').AsString);
    vTobAna.PutValue('Y_NUMVENTIL', NumV);
    vTobAna.PutValue('Y_POURCENTAGE', PM);
    vTobAna.PutValue('Y_POURCENTQTE1', P1);
    vTobAna.PutValue('Y_QTE1', Arrondi(ObjEcr.GetValue('E_QTE1') * P1 / 100.0, V_PGI.OkDecV));
    vTobAna.PutValue('Y_POURCENTQTE2', P2);
    vTobAna.PutValue('Y_QTE2', Arrondi(ObjEcr.GetValue('E_QTE2') * P2 / 100.0, V_PGI.OkDecV));

    for i := 1 to MaxAxe do
    begin
      vTobAna.PutValue('Y_SOUSPLAN' + IntToStr(i), Q.FindField('V_SOUSPLAN' + IntToStr(i)).AsString);
    end;

    if Sens = 1 then
    begin
      vTobAna.PutValue('Y_DEBIT', XP);
      vTobAna.PutValue('Y_DEBITDEV', XD);
      vTobAna.PutValue('Y_CREDIT', 0);
      vTobAna.PutValue('Y_CREDITDEV', 0);
    end
    else
    begin
      vTobAna.PutValue('Y_DEBIT', 0);
      vTobAna.PutValue('Y_DEBITDEV', 0);
      vTobAna.PutValue('Y_CREDIT', XP);
      vTobAna.PutValue('Y_CREDITDEV', XD);
    end;
    Q.Next;

    //si derniere //si somme de 100 % alors on arrondi le total sur la derniere ligne
    if Q.Eof then
    begin
      DiffP := Arrondi(MontantEcrP - TotMP, V_PGI.OkDecV);
      DiffD := Arrondi(MontantEcrD - TotMD, Arguments.Dev.Decimale);
      if (Arrondi(TotP - 100.0, ADecimP) = 0) and ((DiffP <> 0) or (DiffD <> 0)) then
      begin
        if ((XP + DiffP <> 0) and (XD + DiffD <> 0)) then
        begin
          XP := XP + DiffP;
          XD := XD + DiffD;
          if Sens = 1 then
          begin
            vTobAna.PutValue('Y_DEBIT', XP);
            vtobana.PutValue('Y_DEBITDEV', XD);
          end
          else
          begin
            vTobAna.PutValue('Y_CREDIT', XP);
            vtobAna.PutValue('Y_CREDITDEV', XD);
          end;
        end;
      end;
    end;

  end;

  //Gestion Grid
  if vTob.Detail[premier_axe - 1].Detail.Count <> 0 then
  begin
    ChargementToBGrid(vTob);
    MAJMontant;
  end;

  FreeAndNil(vTob);

  Ferme(Q);
end;

procedure TOF_CPVENTIL.BValiderVentilTypeClick(Sender : TObject);
var
  vTobCC, vTobVentil, vTobAna : TOB;
  i, j : integer;
begin
  //Vérification
  if (THCode.Text = '') or (THLib.Text = '') then
  begin
    PGIInfo('Vous devez saisir un code et un libellé.','Ventilations analytiques');
    Exit
  end;
  if ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="VTY" AND (CC_CODE="' + THCode.Text + '" OR UPPER(CC_LIBELLE)="' + uppercase(THLib.Text) + '")') then
  begin
    PGIInfo('Le code ou le libellé existe déjà.','Ventilations analytiques');
    Exit;
  end;

  vTobCC := TOB.Create('CHOIXCOD',nil,-1);
  vTobCC.PutValue('CC_TYPE','VTY');
  vTobCC.PutValue('CC_CODE',THCode.Text);
  vTobCC.PutValue('CC_LIBELLE',THLib.Text);
  vTobCC.InsertDB(nil);
  FreeAndNil(vTobCC);

  for i := 2 to GSA.RowCount - 1 do
  begin
    if LigneRempli(i) then
    begin
      vTobAna := TOB(GSA.Objects[0,i]);
      vTobVentil := TOB.Create('VENTIL',nil,-1);
      vTobVentil.PutValue('V_NATURE','TY'+IntToStr(premier_axe));
      vTobVentil.PutValue('V_COMPTE',THCode.text);
      vTobVentil.PutValue('V_SECTION',GSA.Cells[AN_MSect[premier_axe],i]);
      vTobVentil.PutValue('V_TAUXMONTANT',GSA.Cells[AN_MPourcent,i]);
      vTobVentil.PutValue('V_TAUXQTE1',vTobAna.GetValue('Y_POURCENTQTE1'));
      vTobVentil.PutValue('V_TAUXQTE2',vTobAna.GetValue('Y_POURCENTQTE2'));
      vTobVentil.PutValue('V_NUMEROVENTIL',i-1);
      for j := 1 to MaxAxe do
      begin
        if Axes[j] then
        begin
          vTobVentil.PutValue('V_SOUSPLAN' + IntToStr(j), GSA.Cells[AN_MSect[j], i]);
        end;
      end;

      vTobVentil.InsertDB(nil);
      FreeAndNil(vTobVentil);
    end;
  end;

  //Fin Ventil
  FinVentil;
  AvertirTable('ttVentilType') ;
  CVType.DataType:='' ;
  CVType.DataType:='ttVentilType';
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 07/02/2005
Modifié le ... :   /  /    
Description .. : Procedure qui permet d'initialiser les controles à la fin de 
Suite ........ : l'enregistrement d'un ventilation type
Mots clefs ... : VENTILATION TYPE FIN
*****************************************************************}
procedure TOF_CPVENTIL.FinVentil;
begin
  PEntete.Enabled := True;
  PGA.Enabled := True ;
  PPied.Enabled:=True ;
  GSA.Enabled:=True ;
  BValider.Enabled := EtatBoutonBValider;
  BFerme.Enabled := True;
  HelpBtn.Enabled := True;
  BZoom.Enabled := True;
  BComplement.Enabled := True;
  BVentilType.Enabled := True;
  BSolde.Enabled := True;
  CVType.SetFocus ;
  PVentilType.Visible:=False ;
  ModeVentil:=False ;

end;

function TOF_CPVENTIL.LookUpSectionList(Ctrl: TControl; Axe: String; ARow, NumTag: Integer): Boolean;
var
  Clause:       String;
begin
  {b FP 29/12/2005}
  Clause := GetClauseRestrictionAna(Axe, ARow);   {FP 19/04/2006 FQ17728}
  Result := LookUpList(Ctrl, 'Sections analytiques', 'SECTION', 'S_SECTION', 'S_LIBELLE', Clause, 'S_SECTION', True, NumTag);
  {e FP 29/12/2005}
end;

function TOF_CPVENTIL.ExisteSQLSection(Axe, Section: String; ARow: Integer): Boolean;
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

{b FP 19/04/2006 FQ17728}
function TOF_CPVENTIL.GetClauseRestrictionAna(Axe: String; ARow: Integer): String;
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
  if FRestriction.IsModelExclu(G_GENERAL.Caption, 'A'+IntToStr(premier_axe)) then
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
      ' AND S_FERME="-"' +  //YMO 20/06/2006 FQ18223 Uniquement les sections ouvertes
      ' AND ' + FRestriction.GetClauseCompteAutorise(G_GENERAL.Caption, Axe, 'SECTION', CompteAna);
end;
{e FP 19/04/2006}

initialization
  registerclasses([TOF_CPVENTIL]);
end.

