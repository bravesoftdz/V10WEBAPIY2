{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/12/2000
Modifié le ... : 25/11/2002
Description .. : Saisie des échéances pour le Front Office
Mots clefs ... : FO                                                                  
*****************************************************************}
unit EcheanceFO;

interface

uses
  Windows, Forms, Menus, hmsgbox, HSysMenu, StdCtrls, Hctrls, HFLabel, ExtCtrls,
  Math, SysUtils, HTB97, LCD_Lab, Controls, HPanel, Classes, Grids, LookUp,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, dbtables,
  {$ENDIF}
  AGLInit, UTob, Hent1, UIUtil, SaisUtil, EntGC;

function GereEcheancesFO(FF: TForm; CleDoc: R_CleDoc; var TOBPiece, TOBEches, TOBAcomptes: TOB; Action: TActionFiche): Boolean;
function SaisirEcheanceFO(CleDoc: R_CleDoc; Action: TActionFiche = taModif): Boolean;

// Choix du contrôle où se positionner dans la fenêtre de saisie de la remise globale
type TChoixCtrlRemTic = (rtPourcent, rtMontant, rtTotal, rtMotif);

type
  TFEcheanceFO = class(TForm)
    GSReg: THGrid;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    PPied: THPanel;
    TFAPayer: TLabel;
    TFEncaisse: TLabel;
    FReste: THNumEdit;
    FEncaisse: THNumEdit;
    FTotalBrut: THNumEdit;
    DockBottom: TDock97;
    Outils97: TToolbar97;
    BMenuZoom: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    bNewligne: TToolbarButton97;
    bSolde: TToolbarButton97;
    bDelLigne: TToolbarButton97;
    LFAPayer: THLabel;
    LFEncaisse: THLabel;
    LFReste: TLCDLabel;
    POPZ: TPopupMenu;
    DetailLigne: TMenuItem;
    N1: TMenuItem;
    VoirModePaie: TMenuItem;
    VoirDevise: TMenuItem;
    RemiseTicket: TToolWindow97;
    BValiderRemise: TToolbarButton97;
    BCancelRemise: TToolbarButton97;
    TFTotalBrut: THLabel;
    LFTotalBrut: THLabel;
    TGP_REMISEPOURPIED: THLabel;
    GP_REMISEPOURPIED: THNumEdit;
    TT_PI_TAUXREMISE: THLabel;
    GP_REMISETTCDEV: THNumEdit;
    FTotAPAyer: THNumEdit;
    TFTotAPAyer: THLabel;
    Bevel1: TBevel;
    TFTotRegle: THLabel;
    FTotRegle: TFlashingLabel;
    FTotARendre: TFlashingLabel;
    TFTotARendre: THLabel;
    GP_TYPEREMISE: THValComboBox;
    FAPayer: THNumEdit;
    BRemise: TToolbarButton97;
    LDETAXE: TLabel;
    TSOLDETIERS: THLabel;
    SOLDETIERS: TFlashingLabel;
    FBaseRemise: THNumEdit;
    LFBaseRemise: THLabel;
    TLFBaseRemise: THLabel;
    procedure FormShow(Sender: TObject);
    procedure GSRegCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSRegCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure GSRegRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GSRegEnter(Sender: TObject);
    procedure GSRegRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GSRegElipsisClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure bDelLigneClick(Sender: TObject);
    procedure ClickDel(ARow: Integer);
    procedure bNewligneClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure GSRegDblClick(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure AfficheZonePied(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure DetailLigneClick(Sender: TObject);
    procedure VoirModePaieClick(Sender: TObject);
    procedure VoirDeviseClick(Sender: TObject);
    procedure GP_REMISEPOURPIEDExit(Sender: TObject);
    procedure GP_REMISETTCDEVExit(Sender: TObject);
    procedure FTotAPAyerExit(Sender: TObject);
    procedure BCancelRemiseClick(Sender: TObject);
    procedure BValiderRemiseClick(Sender: TObject);
    procedure RemiseTicketEnter(Sender: TObject);
    procedure BRemiseClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Déclarations privées }
    StCellCur: string; // Valeur d'une cellule avant modification
    LesColonnes: string; // Identifiant des colonnes du Grid
    TOBMode: TOB; // Liste des modes de paiements dispo en caisse
    TOBDev: TOB; // Liste des devises utilisées
    MntEcart: Double; // Montant écart de change ou d'arrondi
    MntRef: Double; // Montant de référence pour la conversion en devise
    FormTicket: TForm; // Forme de saisie des tickets
    PourRemCur: string; // Pourcentage de remise ticket courant
    MntRemCur: string; // Montant de remise ticket courant
    TotRemCur: string; // Total à payer après remise ticket courant
    ValiderOk: Boolean; // La saisie est valide
    RenduEnCours: Boolean; // La saisie est valide
    // Présentation de la fiche
    {$IFDEF FOS5}
    procedure LieTicketFO(FF: TForm);
    {$ENDIF}
    // Gestion des TOBS
    procedure CreeLesTobs;
    procedure LibereLesTobs;
    // Gestion des TOB Echéances
    procedure CreeTOBEche(ARow: integer);
    procedure InsereTOBEche(ARow: integer);
    procedure InitTOBEche(TOBLigne: TOB; NoLigne: Integer; AffCell: Boolean);
    procedure RazLigne(ARow: Integer; TOBLigne: TOB; RazCellCur: Boolean);
    procedure RenumeroteLigne;
    procedure DepileTOBEche(ARow, NewRow: integer);
    procedure EmpileTOBEche(NewRow: Integer);
    function GetTobEcheance(ARow: Integer): TOB;
    procedure ChargeTobEcheance;
    {$IFDEF FOS5}
    procedure VerifAction;
    {$ENDIF}
    procedure EnregistreSaisie;
    // Gestion de la TOB Pièce
    procedure ChargeTobPiec;
    procedure ChargeDevisePiece;
    // Gestion des TOB Devises
    procedure AjoutDevise(CodeDev: string);
    function GetTobDevise(ARow: Integer): TOB;
    function GetNbDecDevise(ARow: Integer): Integer;
    function GetSymboleDevise(CodeDev: string; ARow: Integer): string;
    // Gestion des TOB Modes de paiements
    procedure ChargeModePaie;
    function GetTobModePaie(ARow: Integer): TOB;
    function GetLibelleModePaie(CodeMode: string): string;
    function GetTypeModePaie(CodeMode: string): string;
    // Traitement sur les champs
    procedure PrepareModePaie(var ACol, ARow: integer; var Cancel: boolean);
    procedure TraiteModePaie(var ACol, ARow: integer; var Cancel: boolean);
    function ModePaieAutorise(TOBMdp: TOB): boolean;
    function VerifClientObligatoire(TOBMdp: TOB): Boolean;
    function VerifModePaie(ModePaie: string; Final: Boolean): Integer;
    procedure PrepareMontant(var ACol, ARow: integer; var Cancel: boolean);
    procedure TraiteMontant(var ACol, ARow: integer; var Cancel: boolean);
    function VerifMontant(MontantEch: double; ModePaie: string; VerifLimiteModePaie: boolean): integer;
    procedure TraiteDateEch(var ACol, ARow: integer; var Cancel: boolean);
    function VerifDateEch(DateEch: TDateTime; MdpDiffere: Boolean): Integer;
    // Conversion des montants
    procedure ConvertMontant(MntDevSaisi: Double; CodeDev: string; var MontantP, MontantD : Double);
    function ArrondiMontantLigne(ARow: integer; MntLigne: Double): Double;
    function CalculeLigne(ARow: integer): Double;
    function CalculeMontantLigne(ARow: Integer; Montant: Double): Double;
    function LigneVautReste(ARow: Integer): Boolean;
    // Manipulation des lignes
    procedure AfficheLigne(ARow: integer);
    procedure CompleteLesLignes;
    procedure GenereLigneEcart;
    procedure InitLigneEche(TOBMode, TOBLigne: TOB; ForceRaz: boolean);
    function LigneVide(ARow: Integer): Boolean;
    procedure InitInfoComplement(TOBL, TOBM: TOB);
    procedure MarqueUneLigneRemiseBnq(TOBL, TOBM: TOB);
    procedure MarqueRemiseBnq;
    // Actions liées au GRID
    procedure EtudieColsGrid(GS: THGrid);
    procedure InitLesCols;
    procedure FormateCols;
    procedure FormateZoneSaisie(ACol, ARow: Longint);
    function ZoneAccessible(ACol, ARow: Longint): boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function VerifieTOBEche(NoLigne: Integer): Integer;
    function SaisieCorrecte: Boolean;
    function CellIsModified(ACol, ARow: Integer): Boolean;
    // Calculs et affichage des totaux
    procedure InitTotaux;
    procedure AfficheTotal;
    function CalculReste: Double;
    // Actions sur la forme
    procedure BloqueFiche(ActionSaisie: TActionFiche; Bloc: boolean);
    procedure GereEnabled(ARow: integer);
    // Remise sur le ticket
    procedure ArrondiRemisePied;
    procedure SauveMontantRemiseCur;
    procedure CalculTotalRemise;
    function VerifTotalRemise: Boolean;
    {$IFDEF FOS5}
    procedure MiseAJourPiece;
    {$ENDIF}
    // Impression des chèques, bons et validation par le TPE
    {$IFDEF FOS5}
    procedure ImpressionCheque;
    function ValideCB: Boolean;
    {$ENDIF}
    // Liaisons des réglements
    function RechercheReglementDispo(ARow: integer; MntReste: double; var Montant: double): boolean;
  public
    { Déclarations publiques }
    Action: TActionFiche; // type de saisie
    CleDoc: R_CleDoc; // Identifiant de la pièce associée
    DEV: RDEVISE; // Caractéristique de la devise de saisie de la piece
    DepuisPiece: Boolean; // Appel depuis la saisie d'une pièce
    RemTicketCtrl: TWinControl; // Contrôle actif de la saisie de la remise du ticket
    TOBEche: TOB; // Liste des échéances saisies
    TOBPiec: TOB; // Pièce associée
    TOBAcc: TOB; // Liste des acomptes
    // Rang des colonnes
    SG_NL: integer; // Numéro de ligne
    SG_Mode: integer; // Mode de paiement (présence obligatoire)
    SG_Lib: integer; // Libellé du mode de paiement
    SG_Mont: integer; // Montant de l'échéance (présence obligatoire)
    SG_Dev: integer; // Devise de l'échéance
    SG_Ech: integer; // Date d'échéance (présence obligatoire)
    // Actions sur la forme
    procedure EnvoieToucheGrid;
    function ClickValide: Boolean;
    procedure ClickAbandon;
    procedure ZoomModePaie;
    procedure ZoomDevise;
    {$IFDEF FOS5}
    procedure AfficheSoldeClient;
    {$ENDIF}
    // Manipulation des lignes
    procedure TOBEcheLigneExiste;
    procedure AfficheLigneLCD(ARow: integer);
    procedure SupprimeLigne;
    // Remise sur le ticket
    procedure SaisieRemiseTicket(Pourcent: Double; CodeDem: string);
    procedure AllerSurChpRemise(NoChamp: TChoixCtrlRemTic);
    // Détaxe
    procedure TraiteDetaxe;
  end;

implementation
{$R *.DFM}

uses
  Ent1, UtilPGI, UtilCB,
  {$IFNDEF SANSCOMPTA}
  Devise_tom,
  {$ENDIF}
  TarifUtil, FactUtil, FactCalc, FactTOB,
  {$IFDEF FOS5}
  FOAutCB, TicketFO, MFODETAILCB_TOF,
  {$ENDIF}
  TickUtilFO, FODefi, FOUtil;

const NbRowsInit = 13;
  NbRowsPlus = 5;

  {==============================================================================================}
  {=============================== FONCTIONS PRINCIPALES ========================================}
  {==============================================================================================}
  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 27/12/2000
  Modifié le ... : 23/07/2001
  Description .. : Consultation ou modification des échéances d'un ticket du
  Suite ........ : Front Office
  Mots clefs ... : FO
  *****************************************************************}

function SaisirEcheanceFO(CleDoc: R_CleDoc; Action: TActionFiche = taModif): Boolean;
var X: TFEcheanceFO;
  PP: THPanel;
begin
  Result := False;
  if (Action = taModif) and (FOComptaEstActive(CleDoc.NaturePiece)) then
  begin
    PGIBox('Opération impossible car la comptabilité est alimentée depuis la caisse ', 'Modification des règlements');
    Exit;
  end;
  SourisSablier;
  X := TFEcheanceFO.Create(Application);
  X.Action := Action;
  X.CleDoc := CleDoc;
  X.DepuisPiece := False;
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      Result := (X.ShowModal = mrOK);
    finally
      X.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/12/2000
Modifié le ... : 23/07/2001
Description .. : Saisie des échéances d'un ticlet du Front Office
Mots clefs ... : FO
*****************************************************************}

function GereEcheancesFO(FF: TForm; CleDoc: R_CleDoc; var TOBPiece, TOBEches, TOBAcomptes: TOB; Action: TActionFiche): Boolean;
var X: TFEcheanceFO;
  PP: THPanel;
  {$IFNDEF FOS5}
  Stg: string;
  {$ENDIF}
begin
  Result := False;
  SourisSablier;
  {$IFNDEF FOS5}
  Stg := TOBPiece.GetValue('GP_CAISSE');
  if Stg = '' then Stg := FOCaisseEtab(TOBPiece.GetValue('GP_ETABLISSEMENT'));
  if Stg <> FOCaisseCourante then FOChargeVHGCCaisse(Stg);
  {$ENDIF}
  X := TFEcheanceFO.Create(Application);
  X.Action := Action;
  X.DepuisPiece := True;
  X.TOBPiec.Dupliquer(TOBPiece, True, True, True);
  X.TOBEche.Dupliquer(TOBEches, True, True, True);
  X.CleDoc := CleDoc;
  X.TOBAcc := TOBAcomptes;
  PP := nil;
  if FF <> nil then PP := THPanel(FF.FindComponent('PnlCorps'));
  if PP = nil then
  begin
    try
      {$IFDEF FOS5}
      Result := (X.ShowModal = mrOK);
      {$ELSE}
      X.ShowModal;
      {$ENDIF}
    finally
      {$IFNDEF FOS5}
      if X.ValiderOk then
      begin
        ///X.MiseAJourPiece ;
        ///FORecopieTob(X.TOBPiec, TOBPiece) ;
        ///FORecopieTob(X.TOBEche, TOBEches) ;
        TOBEches.Dupliquer(X.TOBEche, True, True, True);
      end;
      Result := X.ValiderOk;
      {$ENDIF}
      X.LibereLesTobs;
      X.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
  {$IFDEF FOS5}
  // On conserve un lien sur la forme de saisie des tickets
  if (FF <> nil) and (FF is TFTicketFO) then X.LieTicketFO(FF);
  {$ENDIF}
end;

{==============================================================================================}
{=============================== PRESENTATION DE LA FICHE =====================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  LieTicketFO : établit un lien entre la fiche de saisie des tickets et la fiche de saisie des échéances
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF FOS5}

procedure TFEcheanceFO.LieTicketFO(FF: TForm);
begin
  if FF = nil then Exit;
  if not (FF is TFTicketFO) then Exit;
  // On conserve un lien sur la forme de saisie des tickets
  FormTicket := TFTicketFO(FF);
  TFTicketFO(FF).FormEcheance := Self;
  // On reprend certaines propriétés d'affichage de la fiche de saisie des tickets
  GSReg.Ctl3D := TFTicketFO(FF).GS.Ctl3D;
  GSReg.FixedColor := TFTicketFO(FF).GS.FixedColor;
  GSReg.TwoColors := TFTicketFO(FF).GS.TwoColors;
  GSReg.Font := TFTicketFO(FF).GS.Font;
  PPied.ColorStart := TFTicketFO(FF).PPied.ColorStart;
  PPied.ColorEnd := TFTicketFO(FF).PPied.ColorEnd;
  PPied.BackGroundEffect := TFTicketFO(FF).PPied.BackGroundEffect;
  PPied.ColorNB := TFTicketFO(FF).PPied.ColorNB;
  LFReste.BackGround := TFTicketFO(FF).PI_NETAPAYERDEV.BackGround;
  LFReste.PixelOn := TFTicketFO(FF).PI_NETAPAYERDEV.PixelOn;
  LFReste.PixelOff := TFTicketFO(FF).PI_NETAPAYERDEV.PixelOff;
  FOAppliqueColCarac(GSReg, TFTicketFO(FF).ColCarac, 20);
  HMTrad.ResizeGridColumns(GSReg);
  LDETAXE.Visible := TFTicketFO(FormTicket).DemandeDetaxe;
  // affichage du solde du client
  SOLDETIERS.Caption := TFTicketFO(FormTicket).SOLDETIERS.Caption;
  SOLDETIERS.Flashing := TFTicketFO(FormTicket).SOLDETIERS.Flashing;
  SOLDETIERS.Font.Color := TFTicketFO(FormTicket).SOLDETIERS.Font.Color;
  SOLDETIERS.Visible := TFTicketFO(FormTicket).SOLDETIERS.Visible;
  TSOLDETIERS.Visible := TFTicketFO(FormTicket).TSOLDETIERS.Visible;
  // incrustation et positionnement des barres de boutons dans l'écran principal
  if TFTicketFO(FF).DockBottom.Visible then
  begin
    if FOGetParamCaisse('GPK_TOOLBAR') = 'X' then FoShowEtiquette(Self, True);
    TFTicketFO(FF).Valide97.Visible := False;
    TFTicketFO(FF).Outils97.Visible := False;
    Valide97.DockedTo := TFTicketFO(FF).DockBottom;
    Valide97.DragHandleStyle := dhNone;
    Valide97.BorderStyle := bsNone;
    Valide97.DockPos := TFTicketFO(FF).DockBottom.Width - Valide97.Width;
    Outils97.DockedTo := TFTicketFO(FF).DockBottom;
    Outils97.DragHandleStyle := dhNone;
    Outils97.BorderStyle := bsNone;
    Outils97.DockPos := 0;
  end;
  // On rend invisible la barre de boutons et on affiche le reste à payer dans l'afficheur
  DockBottom.Visible := False;
  AfficheLigneLCD(1);
  if GSReg.CanFocus then GSReg.SetFocus;
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheSoldeClient : affichage du solde du client
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF FOS5}

procedure TFEcheanceFO.AfficheSoldeClient;
begin
  if FormTicket = nil then Exit;
  if not (FormTicket is TFTicketFO) then Exit;
  SOLDETIERS.Caption := TFTicketFO(FormTicket).SOLDETIERS.Caption;
  SOLDETIERS.Flashing := TFTicketFO(FormTicket).SOLDETIERS.Flashing;
  SOLDETIERS.Font.Color := TFTicketFO(FormTicket).SOLDETIERS.Font.Color;
  SOLDETIERS.Visible := TFTicketFO(FormTicket).SOLDETIERS.Visible;
  TSOLDETIERS.Visible := TFTicketFO(FormTicket).TSOLDETIERS.Visible;
end;
{$ENDIF}

{==============================================================================================}
{=============================== GESTION DES TOBS =============================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  CreeLesTobs : création des differentes TOB
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.CreeLesTobs;
begin
  TOBMode := TOB.Create('Les MODEPAIES', nil, -1);
  TOBDev := TOB.Create('Les DEVISES', nil, -1);
  TOBEche := TOB.Create('Les ECHEANCES', nil, -1);
  TOBPiec := TOB.Create('PIECE', nil, -1);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LibereLesTobs : libération des differentes TOB
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.LibereLesTobs;
begin
  if TOBMode <> nil then TOBMode.Free;
  TOBMode := nil;
  if TOBDev <> nil then TOBDev.Free;
  TOBDev := nil;
  if TOBEche <> nil then TOBEche.Free;
  TOBEche := nil;
  if TOBPiec <> nil then TOBPiec.Free;
  TOBPiec := nil;
end;

{==============================================================================================}
{====================================== ECHEANCES =============================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  CreeTOBEche : création d'une TOB échéances
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.CreeTOBEche(ARow: integer);
var TOBL: TOB;
begin
  if TOBEche.Detail.Count < ARow then
  begin
    TOBL := TOB.Create('PIEDECHE', TOBEche, -1);
    InitTOBEche(TOBL, TOBEche.Detail.Count, True);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InsereTOBEche : insertion d'une TOB échéances
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.InsereTOBEche(ARow: integer);
var TOBL: TOB;
  NumL, Indice: integer;
begin
  if TOBEche.Detail.Count < ARow then
  begin
    // ajout en fin de liste
    NumL := TOBEche.Detail.Count;
    Indice := -1;
  end else
  begin
    // insertion dans la liste
    NumL := ARow;
    Indice := ARow - 1;
  end;
  TOBL := TOB.Create('PIEDECHE', TOBEche, Indice);
  InitTOBEche(TOBL, NumL, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InitTOBEche : initialisation d'une TOB échéances
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.InitTOBEche(TOBLigne: TOB; NoLigne: Integer; AffCell: Boolean);
begin
  if TOBLigne <> nil then
  begin
    TOBLigne.InitValeurs;
    if not TOBLigne.FieldExists('GPE_LIBELLE') then TOBLigne.AddChampSupValeur('GPE_LIBELLE', '');
    if not TOBLigne.FieldExists('TYPEMODEPAIE') then TOBLigne.AddChampSupValeur('TYPEMODEPAIE', '');
    TOBLigne.PutValue('GPE_NUMECHE', NoLigne);
    if (AffCell) and (SG_NL <> -1) then GSReg.Cells[SG_NL, NoLigne] := IntToStr(NoLigne);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RazLigne : met à blanc une ligne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.RazLigne(ARow: Integer; TOBLigne: TOB; RazCellCur: Boolean);
var Ind: integer;
begin
  for Ind := 0 to GSReg.ColCount - 1 do GSReg.Cells[Ind, ARow] := '';
  if SG_NL <> -1 then GSReg.Cells[SG_NL, ARow] := IntToStr(ARow);
  if RazCellCur then StCellCur := '';
  if TOBLigne <> nil then InitTOBEche(TOBLigne, ARow, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RenumeroteLigne : renumérote les lignes
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.RenumeroteLigne;
var TOBL: TOB;
  Ind: Integer;
begin
  for Ind := 1 to TOBEche.Detail.Count do
  begin
    TOBL := TOBEche.Detail[(Ind - 1)];
    TOBL.PutValue('GPE_NUMECHE', Ind);
    if SG_NL <> -1 then GSReg.Cells[SG_NL, Ind] := IntToStr(Ind);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DepileTOBEche : supprime les TOB échéances vides
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.DepileTOBEche(ARow, NewRow: integer);
var Ind, Limite, MaxL, Col: integer;
begin
  if NewRow > ARow then Exit;
  Limite := TOBEche.Detail.Count + 1;
  MaxL := Max(Limite, ARow);
  for Ind := MaxL downto NewRow + 1 do
  begin
    if LigneVide(Ind) then Limite := Ind else Break;
  end;
  for Ind := TOBEche.Detail.Count - 1 downto Limite - 1 do
  begin
    TOBEche.Detail[Ind].Free;
    if SG_NL <> -1 then for Col := GSReg.FixedCols to GSReg.ColCount - 1 do GSReg.Cells[SG_NL, Ind + 1] := '';
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EmpileTOBEche : ajoute les TOB échéances vides
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.EmpileTOBEche(NewRow: Integer);
var Ind: Integer;
begin
  for Ind := TOBEche.Detail.Count + 1 to NewRow do CreeTOBEche(Ind);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetTobEcheance : retourne la TOB d'une ligne d'échéance
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.GetTobEcheance(ARow: Integer): TOB;
begin
  if ((ARow > 0) and (ARow <= TOBEche.Detail.Count)) then Result := TOBEche.Detail[ARow - 1] else Result := nil;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeTobEcheance : charge en TOB les échéances d'une pièce
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ChargeTobEcheance;
var sSql: string;
  QQ: TQuery;
begin
  if DepuisPiece then Exit;
  // Lecture des echéances
  sSql := 'SELECT * FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False) + ' ORDER BY GPE_NUMECHE';
  QQ := OpenSQL(sSql, True);
  TOBEche.LoadDetailDB('PIEDECHE', '', '', QQ, False);
  Ferme(QQ);
  //TOBEche.LoadDetailFromSQL(sSql) ;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifAction : vérifie l'action en fonction de l'état de la journée de la pièce
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF FOS5}

procedure TFEcheanceFO.VerifAction;
var NumZ: Integer;
  Caisse: string;
begin
  if action <> taModif then Exit;
  if TOBPiec = nil then Exit;
  Caisse := TOBPiec.GetValue('GP_CAISSE');
  if Caisse = '' then Exit;
  NumZ := TOBPiec.GetValue('GP_NUMZCAISSE');
  if FOEtatJournee(Caisse, NumZ) <> ETATJOURCAISSE[2] then Action := taConsult;
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  EnregistreSaisie : enregistre les données saisies
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.EnregistreSaisie;
var Nb: Integer;
begin
  if DepuisPiece then Exit;
  // Mise à jour de la date de modification de la piece
  TOBPiec.UpdateDateModif;
  if not TOBPiec.UpdateDB(False) then
  begin
    V_PGI.IoError := oeSaisie;
    Exit;
  end;
  // Suppression des échéances saisies précédemment
  Nb := ExecuteSQL('DELETE FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False));
  if Nb < 0 then
  begin
    V_PGI.IoError := oeUnknown;
    Exit;
  end;
  // Enregistre la TOB des échéances saisies
  TOBEche.SetAllModifie(True);
  if not TOBEche.InsertDB(nil, False) then
  begin
    V_PGI.IoError := oeUnknown;
    Exit;
  end;
  {$IFDEF FOS5}
  FOIncrJoursCaisse(jfoNbModifMdp, FOFabriqueCommentairePiece(TOBPiec, ''));
  {$ENDIF}
end;

{==============================================================================================}
{====================================== PIECE =================================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  ChargeTobPiec : charge en TOB une pièce
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ChargeTobPiec;
var QQ: TQuery;
begin
  if DepuisPiece then Exit;
  // Lecture de la pièce associée
  QQ := OpenSQL('SELECT * FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False), True);
  TOBPiec.SelectDB('', QQ);
  Ferme(QQ);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeDevisePiece : charge les caractéristiques de la devise de saisie de la pièce
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ChargeDevisePiece;
begin
  if TOBPiec = nil then Exit;

  DEV.Code := TOBPiec.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  if FOTestCodeEuro(DEV.Code) then DEV.Symbole := SIGLEEURO;
  // Ajout dans le liste des devises de la devise de la piéce et de la devise de tenue.
  AjoutDevise(DEV.Code);
  AjoutDevise(V_PGI.DevisePivot);
end;

{==============================================================================================}
{====================================== DEVISES ===============================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  AjoutDevise : ajoute en TOB une devise utilisée par un mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.AjoutDevise(CodeDev: string);
var TOBM: TOB;
begin
  TOBM := TOBDev.FindFirst(['D_DEVISE'], [CodeDev], False);
  if TOBM = nil then
  begin
    TOBM := TOB.Create('DEVISE', TOBDev, -1);
    TOBM.SelectDB('"' + CodeDev + '"', nil, False);
    if FOTestCodeEuro(CodeDev) then TOBM.PutValue('D_SYMBOLE', SIGLEEURO);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetTobDevise : retourne la TOB de la devise d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.GetTobDevise(ARow: Integer): TOB;
var TOBM: TOB;
  CodeDev: string;
begin
  CodeDev := GSReg.CellValues[SG_Dev, ARow];
  if CodeDev = '' then
  begin
    TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [GSReg.Cells[SG_Mode, ARow]], False);
    if TOBM <> nil then CodeDev := TOBM.GetValue('MP_DEVISEFO')
    else CodeDev := DEV.Code;
  end;
  Result := TOBDev.FindFirst(['D_DEVISE'], [CodeDev], False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetNbDecDevise : retourne le nombre de décimales d'une devise d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.GetNbDecDevise(ARow: Integer): Integer;
var TOBD: TOB;
begin
  TOBD := GetTobDevise(ARow);
  if TOBD = nil then Result := V_PGI.OkDecV else Result := TOBD.GetValue('D_DECIMALE');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetSymboleDevise : retourne le symbolle d'une devise
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.GetSymboleDevise(CodeDev: string; ARow: Integer): string;
var TOBD: TOB;
begin
  Result := CodeDev;
  if CodeDev = '' then TOBD := GetTobDevise(ARow)
  else TOBD := TOBDev.FindFirst(['D_DEVISE'], [CodeDev], False);
  if TOBD <> nil then Result := TOBD.GetValue('D_SYMBOLE');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ZoomDevise : Affichage de la devise
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ZoomDevise;
{$IFNDEF SANSCOMPTA}
var TOBD: TOB;
  {$ENDIF}
begin
  {$IFNDEF SANSCOMPTA}
  TOBD := GetTobDevise(GSReg.Row);
  if TOBD = nil then exit;
  if not FOJaiLeDroit(16) then Exit;
  FicheDevise(TOBD.GetValue('D_DEVISE'), taConsult, False);
  {$ENDIF}
end;

{==============================================================================================}
{=============================== MODES DE PAIEMENTS ===========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  ChargeModePaie : chargement en TOB des modes de paiements disponible pour la caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ChargeModePaie;
var MaxMnt, MinMnt, TxDev: double;
  //    QQ                    : TQUERY ;
  TOBM: TOB;
  Ind: integer;
  CodeD {, sSql}: string;
begin
  // lecture des modes de paiement
  TOBMode.LoadDetailDB('MODEPAIE', '', 'MP_MODEPAIE', nil, False, True);
  //  sSql := 'SELECT * FROM MODEPAIE ORDER BY MP_MODEPAIE' ;
  //  QQ := OpenSQL (sSql, True) ;
  //  TOBMode.LoadDetailDB ('MODEPAIE', '', '', QQ, False, True) ;
  //  Ferme(QQ) ;
    // rechecherche des caractéristiques de la devise de la pièce
  TxDev := TOBPiec.GetValue('GP_TAUXDEV');
  for Ind := 0 to TOBMode.Detail.Count - 1 do
  begin
    TOBM := TOBMode.Detail[Ind];
    // si la devise n'est pas renseignée, on utilise la devise du document
    CodeD := TOBM.GetValue('MP_DEVISEFO');
    if CodeD = '' then
    begin
      CodeD := DEV.Code;
      TOBM.PutValue('MP_DEVISEFO', CodeD);
    end;
    AjoutDevise(CodeD);
    // conversion du montant maximum en devise de saisie
    MaxMnt := 0;
    MinMnt := 0;
    if TOBM.GetValue('MP_CONDITION') = 'X' then
    begin
      MaxMnt := TOBM.GetValue('MP_MONTANTMAX');
      if TOBM.FieldExists('MP_MONTANTMIN') then MinMnt := TOBM.GetValue('MP_MONTANTMIN');
      if DEV.Code <> V_PGI.DevisePivot then
      begin
        if VH^.TenueEuro then
        begin
          MaxMnt := EuroToDevise(MaxMnt, TxDev, DEV.Quotite, DEV.Decimale);
          MinMnt := EuroToDevise(MinMnt, TxDev, DEV.Quotite, DEV.Decimale);
        end else
          if TOBPiec.GetValue('GP_SAISIECONTRE') = 'X' then
        begin
          MaxMnt := PivotToEuro(MaxMnt);
          MinMnt := PivotToEuro(MinMnt);
        end else
        begin
          MaxMnt := PivotToDevise(MaxMnt, TxDev, DEV.Quotite, DEV.Decimale);
          MinMnt := PivotToDevise(MinMnt, TxDev, DEV.Quotite, DEV.Decimale);
        end;
      end;
    end;
    TOBM.PutValue('MP_MONTANTMAX', MaxMnt);
    if TOBM.FieldExists('MP_MONTANTMIN') then
      TOBM.PutValue('MP_MONTANTMIN', MinMnt)
    else
      TOBM.AddChampSupValeur('MP_MONTANTMIN', MinMnt);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetTobModePaie : retourne la TOB du mode de paiement d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.GetTobModePaie(ARow: Integer): TOB;
begin
  Result := TOBMode.FindFirst(['MP_MODEPAIE'], [GSReg.Cells[SG_Mode, ARow]], False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetLibelleModePaie : retourne le libelle d'un mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.GetLibelleModePaie(CodeMode: string): string;
var TOBM: TOB;
begin
  Result := CodeMode;
  TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [CodeMode], False);
  if TOBM <> nil then Result := TOBM.GetValue('MP_LIBELLE');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetTypeModePaie : retourne le type d'un mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.GetTypeModePaie(CodeMode: string): string;
var TOBM: TOB;
begin
  Result := '';
  TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [CodeMode], False);
  if TOBM <> nil then Result := TOBM.GetValue('MP_TYPEMODEPAIE');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ModePaieUtilFO : vérifie si un mode de paiement est utilisable en caisse
///////////////////////////////////////////////////////////////////////////////////////

function ModePaieUtilFO(TOBM: TOB): boolean;
begin
  Result := ((TOBM <> nil) and (TOBM.GetValue('MP_UTILFO') = 'X'));
end;


{==============================================================================================}
{=============================== TRAITEMENT SUR LES CHAMPS ====================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  PrepareModePaie : action avant la saisie du mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.PrepareModePaie(var ACol, ARow: integer; var Cancel: boolean);
var sRendu: string;
begin
  // Génération du rendu monnaie
  if (FReste.Value < 0) and (FTotalBrut.Value > 0) and (LigneVide(ARow)) then
  begin
    sRendu := FODonneRendu;
    if sRendu <> '' then
    begin
      GSReg.Cells[SG_Mode, ARow] := sRendu;
      TraiteModePaie(ACol, ARow, Cancel);
      // On se positionne sur le montant.
      RenduEnCours := True;
      GSReg.Col := SG_Mont;
      if GSReg.CanFocus then GSReg.SetFocus;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteModePaie : action suite à la saisie du mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.TraiteModePaie(var ACol, ARow: integer; var Cancel: boolean);
var TOBM, TOBL: TOB;
  NoErr: Integer;
  ForceRaz: boolean;
  TypeMode: string;
begin
  NoErr := 0;
  TOBL := GetTobEcheance(ARow);
  if TOBL = nil then exit;
  // Modification du mode de paiement
  if GSReg.Cells[SG_Mode, ARow] <> TOBL.GetValue('GPE_MODEPAIE') then
  begin
    if TOBL.FieldExists('ACCORDTPE') then
      NoErr := 19 // Autorisation accordée par le TPE
    else if FOChampSupValeurNonVide(TOBL, 'GOC_NUMPIECELIEN') then
      NoErr := 28; // Réglement lié
    if NoErr <> 0 then
    begin
      MsgBox.Execute(NoErr, '', '');
      GSReg.Cells[SG_Mode, ARow] := TOBL.GetValue('GPE_MODEPAIE');
      Cancel := True;
      if GSReg.CanFocus then GSReg.SetFocus;
      Exit;
    end;
  end;
  TOBL.PutValue('GPE_MODEPAIE', GSReg.Cells[SG_Mode, ARow]);
  if LigneVide(ARow) then
  begin
    RazLigne(ARow, TOBL, True);
    AfficheTotal;
  end else
  begin
    // Mode de paiement
    TOBM := GetTobModePaie(ARow);
    if TOBM = nil then NoErr := 5 else
      if not ModePaieUtilFO(TOBM) then NoErr := 29 else
      if GSReg.Cells[SG_Mode, ARow] = FOGetParamCaisse('GPK_MDPECART') then NoErr := 23 else
      if not ModePaieAutorise(TOBM) then NoErr := -1 else
      if not VerifClientObligatoire(TOBM) then NoErr := -24;
    if NoErr <> 0 then
    begin
      RazLigne(ARow, TOBL, True);
      AfficheTotal;
      if NoErr > 0 then MsgBox.Execute(NoErr, '', '');
      Cancel := True;
      if GSReg.CanFocus then GSReg.SetFocus;
      Exit;
    end;
    // Alerte si le nombre maximum d'échéances n'est pas dépassé.
    if (ARow - GSReg.FixedRows) >= MaxEche then MsgBox.Execute(25, '', '');
    // Initialisation de la TOB échéance à partir de la pièce
    TypeMode := TOBM.GetValue('MP_TYPEMODEPAIE');
    if (TypeMode = TYPEPAIEAVOIR) or (TypeMode = TYPEPAIEBONACHAT) or (TypeMode = TYPEPAIEARRHES) then
      ForceRaz := True
    else
      ForceRaz := False;
    InitLigneEche(TOBM, TOBL, ForceRaz);
    MarqueUneLigneRemiseBnq(TOBL, TOBM);
    InitInfoComplement(TOBL, TOBM);
    AfficheLigne(ARow);
  end;
  StCellCur := GSReg.Cells[ACol, ARow];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ModePaieAutorise : vérifie si le mode de paiement est autorisé pour l'utilisateur
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.ModePaieAutorise(TOBMdp: TOB): boolean;
var sTypeMdp: string;
  CodeEvt: integer;
begin
  Result := True;
  if TOBMdp = nil then Exit;
  sTypeMdp := TOBMdp.GetValue('MP_TYPEMODEPAIE');
  if sTypeMdp = '001' then CodeEvt := 72 // Espèce
  else if sTypeMdp = '002' then CodeEvt := 73 // Avoir
  else if sTypeMdp = '003' then CodeEvt := 74 // Chèque
  else if sTypeMdp = '004' then CodeEvt := 75 // Chèque différé
  else if sTypeMdp = '005' then CodeEvt := 76 // Carte bancaire
  else if sTypeMdp = '006' then CodeEvt := 77 // Arrhes déjà versées
  else if sTypeMdp = '007' then CodeEvt := 78 // Reste dû
  else if sTypeMdp = '008' then CodeEvt := 79 // Bon d'achat
  else if sTypeMdp = '009' then CodeEvt := 80 // Autre mode de règlement
  else CodeEvt := 0;

  if CodeEvt > 0 then Result := FOJaiLeDroit(CodeEvt)
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifClientObligatoire : vérifie si la saisie du client est obligatoire pour le mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.VerifClientObligatoire(TOBMdp: TOB): Boolean;
var TOBP: TOB;
begin
  Result := True;
  if TOBMdp = nil then Exit;
  if FOGetParamCaisse('GPK_CLISAISIE') <> 'X' then Exit;
  if (not TOBMdp.FieldExists('MP_CLIOBLIGFO')) or (TOBMdp.GetValue('MP_CLIOBLIGFO') <> 'X') then Exit;
  {$IFDEF FOS5}
  if (FormTicket <> nil) and (FormTicket is TFTicketFO) then
  begin
    TOBP := TFTicketFO(FormTicket).TOBPiece;
    if (TOBP.GetValue('GP_TIERS') = FOGetParamCaisse('GPK_TIERS')) and
      (MsgBox.Execute(24, Caption, '') = mrYes) then
    begin
      TFTicketFO(FormTicket).MCChoixClientClick(nil);
    end;
  end else
  begin
    TOBP := TOBPiec;
  end;
  {$ELSE}
  TOBP := TOBPiec;
  {$ENDIF}
  if TOBP.GetValue('GP_TIERS') = FOGetParamCaisse('GPK_TIERS') then Result := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifModePaie : vérifie un mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.VerifModePaie(ModePaie: string; Final: Boolean): Integer;
var
  TOBM: TOB;
begin
  Result := 0;
  if ModePaie = '' then
  begin
    Result := 5;
  end else
  begin
    TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [ModePaie], False);
    if TOBM = nil then
      Result := 5
    else if not ModePaieUtilFO(TOBM) then
      Result := 29;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetModePaieLookUp : Recherche du mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

function GetModePaieLookUp(CC: TControl): Boolean;
begin
  Result := LookupList(CC, 'Mode de paiement', 'MODEPAIE', 'MP_MODEPAIE', 'MP_LIBELLE,MP_DEVISEFO',
    'MP_UTILFO="X"', 'MP_MODEPAIE', False, 0, '', tlLocate);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ZoomModePaie : Affichage du mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ZoomModePaie;
var TOBL: TOB;
begin
  TOBL := GetTobEcheance(GSReg.Row);
  if TOBL = nil then exit;
  if not FOJaiLeDroit(51) then Exit;
  AGLLanceFiche('YY', 'YYMODEPAIE', '', GSReg.Cells[SG_Mode, GSReg.Row], ActionToString(taConsult));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PrepareMontant : action avant la saisie du montant
///////////////////////////////////////////////////////////////////////////////////////
procedure TFEcheanceFO.PrepareMontant(var ACol, ARow: integer; var Cancel: boolean);
var Montant, MntReste: Double;
  ChoixOk: boolean;
begin
  ChoixOk := False;
  if Valeur(GSReg.Cells[SG_Mont, ARow]) = 0 then
  begin
    //MntReste := FReste.Value ;
    MntReste := CalculReste;
    Montant := CalculeMontantLigne(ARow, MntReste);
    MntRef := Montant;

    // recherche des règlements liés
    if not RechercheReglementDispo(ARow, MntReste, Montant) then
    begin
      ACol := SG_Mode;
      Cancel := True;
      Exit;
    end;
    if Montant = 0 then
      Montant := MntRef
    else
      ChoixOk := True;

    GSReg.Cells[SG_Mont, ARow] := StrF00(Montant, GetNbDecDevise(ARow));
    AfficheLigneLCD(ARow);
  end;
  if ChoixOK then EnvoieToucheGrid;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteMontant : action suite à la saisie du montant
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.TraiteMontant(var ACol, ARow: integer; var Cancel: boolean);
var TOBL: TOB;
  ErrNo: Integer;
  MntSaisi, MontantP, MontantD: Double;
begin
  ErrNo := 0;
  if RenduEnCours then
  begin
    // lors de l'insertion automatique de la ligne de rendu, TraiteMontant sur la ligne précédente est appelée 2 fois
    RenduEnCours := False;
    Exit;
  end;
  // Arrondi du montant en fonction du mode de paiement
  MntSaisi := CalculeLigne(ARow);
  TOBL := GetTobEcheance(ARow);
  if TOBL = nil then exit;
  // Modification du montant
  if MntSaisi <> TOBL.GetValue('GPE_MONTANTENCAIS') then
  begin
    if TOBL.FieldExists('ACCORDTPE') then
      ErrNo := 19 // Autorisation accordée par le TPE
    else if FOChampSupValeurNonVide(TOBL, 'GOC_NUMPIECELIEN') then
      ErrNo := 28; // Réglement lié
  end;
  // Verification du montant
  if ErrNo = 0 then
    ErrNo := VerifMontant(MntSaisi, GSReg.Cells[SG_Mode, ARow], True);
  if ErrNo <> 0 then
  begin
    if ErrNo > 0 then MsgBox.Execute(ErrNo, '', '');
    StCellCur := '';
    Cancel := True;
    if GSReg.CanFocus then GSReg.SetFocus;
    Exit;
  end;
  TOBL.PutValue('GPE_MONTANTENCAIS', MntSaisi);
  StCellCur := GSReg.Cells[ACol, ARow];
  ConvertMontant(MntSaisi, TOBL.GetValue('GPE_DEVISEESP'), MontantP, MontantD);
  // Montant en devise de tenu
  TOBL.PutValue('GPE_MONTANTECHE', MontantP);
  // Montant en devise de la pièce
  TOBL.PutValue('GPE_MONTANTDEV', MontantD);
  AfficheLigne(ARow);
  AfficheTotal;
  // Pour ignorer les écarts d'arrondi et de change
  if MntSaisi = MntRef then
  begin
    MntEcart := FReste.Value;
    FReste.Value := 0;
  end else MntEcart := 0;
  MntRef := 0;
  // Pour saisir la date d'échéance
  if (Action = taCreat) and (SG_Ech <> -1) and
    (TOBL.GetValue('TYPEMODEPAIE') = TYPEPAIECHQDIFF) and
    (TOBL.GetValue('GPE_DATEECHE') <= TOBL.GetValue('GPE_DATEPIECE')) then
  begin
    Cancel := True;
    ACol := SG_Ech;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifMontant : vérifie le montant d'une échéance
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.VerifMontant(MontantEch: double; ModePaie: string; VerifLimiteModePaie: boolean): integer;
var TOBM, TOBD: TOB;
  MtMax, MtMin: double;
  Stg, Symbole, Mess: string;
  Nbdec, NoMsg, Ind: integer;
begin
  Result := 0;
  if MontantEch = 0 then
  begin
    Result := 6; // Vous devez renseigner le montant
  end else
    if (VerifLimiteModePaie) and (ModePaie <> '') then
  begin
    // Vérification des montants maximum et minimum du mode de paiement
    MtMax := 0;
    MtMin := 0;
    Symbole := '';
    Nbdec := V_PGI.OkDecV;
    TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [ModePaie], False);
    if (TOBM <> nil) and (TOBM.GetValue('MP_CONDITION') = 'X') then
    begin
      MtMax := TOBM.GetValue('MP_MONTANTMAX');
      if TOBM.FieldExists('MP_MONTANTMIN') then
        MtMin := TOBM.GetValue('MP_MONTANTMIN');
      TOBD := TOBDev.FindFirst(['D_DEVISE'], [TOBM.GetValue('MP_DEVISEFO')], False);
      if TOBD <> nil then
      begin
        Symbole := TOBD.GetValue('D_SYMBOLE');
        NbDec := TOBD.GetValue('D_DECIMALE');
      end;
    end;
    NoMsg := 0;
    if (MtMin > 0) and (MontantEch < MtMin) then
    begin
      Stg := StrfMontant(MtMin, 12, NbDec, Symbole, True);
      NoMsg := 27;
    end else
      if (MtMax > 0) and (MontantEch > MtMax) then
    begin
      Stg := StrfMontant(MtMax, 12, NbDec, Symbole, True);
      NoMsg := 1;
    end;
    if NoMsg > 0 then
    begin
      if FOJaiLeDroit(55, False, False) then
      begin
        if MsgBox.Execute(NoMsg, Stg, '') = mrNo then Result := -1;
      end else
      begin
        Mess := FOMsgBoxExec(MsgBox, NoMsg, Stg, '');
        Ind := Pos('#10', Mess);
        if Ind > 0 then Delete(Mess, Ind, MaxInt);
        if not FOJaiLeDroit(55, True, True, Mess) then Result := -1;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteDateEch : action suite à la saisie de la date d'échéance
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.TraiteDateEch(var ACol, ARow: integer; var Cancel: boolean);
var TOBL: TOB;
  DateEch: TDateTime;
  ErrNo: Integer;
begin
  DateEch := StrToDate(GSReg.Cells[SG_Ech, ARow]);
  ErrNo := VerifDateEch(DateEch, True);
  if ErrNo > 0 then
  begin
    MsgBox.Execute(ErrNo, '', '');
    Cancel := True;
    if GSReg.CanFocus then GSReg.SetFocus;
    Exit;
  end;
  TOBL := GetTobEcheance(ARow);
  if TOBL = nil then exit;
  TOBL.PutValue('GPE_DATEECHE', StrToDate(GSReg.Cells[SG_Ech, ARow]));
  StCellCur := GSReg.Cells[ACol, ARow];
  AfficheLigne(ARow);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifDateEch : vérifie une date d'échéance
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.VerifDateEch(DateEch: TDateTime; MdpDiffere: Boolean): Integer;
begin
  Result := 0;
  if MdpDiffere then
  begin
    if CleDoc.DatePiece >= DateEch then Result := 21;
  end else
  begin
    if CleDoc.DatePiece > DateEch then Result := 20;
  end;
  if (Result = 0) and not (NbJoursOK(CleDoc.DatePiece, DateEch)) then Result := 3;
end;

{==============================================================================================}
{=============================== CONVERSION DES MONTANTS ======================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  ConvertMontant : convertit le montant saisi dans les différentes devises
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ConvertMontant(MntDevSaisi: Double; CodeDev: string; var MontantP, MontantD: Double);
begin
  FOConvtMntEcheance(MntDevSaisi, CodeDev, TOBPiec, TOBDev, DEV, MontantP, MontantD);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetCotation : retourne le taux d'une devise par rapport à l'Euro (1,00 EUR = xx,xxxx DEV)
///////////////////////////////////////////////////////////////////////////////////////

function GetCotation(Code: String3; DateC: TDateTime): Double;
var QQ: TQuery;
  SQL: string;
  LeTaux, LaCote: double;
begin
  Result := 1;
  if ((Code = V_PGI.DevisePivot) or (Code = V_PGI.DeviseFongible)) then Exit;
  if DateC < V_PGI.DateDebutEuro then
  begin
    // Si date < date début --> nécessairement par rapport à ancienne chancellerie
    SQL := 'Select H_DATECOURS, H_TAUXREEL FROM CHANCELL WHERE H_DEVISE="' + Code + '" '
      + 'AND H_DATECOURS<="' + UsDateTime(DateC) + '" AND H_DATECOURS<"' + UsDateTime(V_PGI.DateDebutEuro) + '"';
    SQL := SQL + 'ORDER BY H_DATECOURS DESC';
    QQ := OpenSQL(SQL, True);
    if not QQ.EOF then Result := QQ.Fields[1].AsFloat;
    Ferme(QQ);
  end else
  begin
    if EstMonnaieIN(Code) then
    begin
      // Monnaie IN pour date > début euro --> parité fixe de la devise
      QQ := OpenSQL('Select D_PARITEEURO from DEVISE Where D_DEVISE="' + Code + '"', True);
      if not QQ.EOF then Result := QQ.Fields[0].AsFloat;
      Ferme(QQ);
      if Result = 0 then Result := 1;
    end else
    begin
      // Monnaie Out pour date > début euro --> chancellerie par rapport à Euro
      SQL := 'Select H_DATECOURS, H_TAUXREEL, H_COTATION FROM CHANCELL WHERE H_DEVISE="' + Code + '" '
        + 'AND H_DATECOURS<="' + UsDateTime(DateC) + '" AND H_DATECOURS>="' + UsDateTime(V_PGI.DateDebutEuro) + '"';
      SQL := SQL + 'ORDER BY H_DATECOURS DESC';
      QQ := OpenSQL(SQL, True);
      if not QQ.EOF then
      begin
        LeTaux := Arrondi(QQ.Fields[1].AsFloat, 9);
        LaCote := Arrondi(QQ.Fields[2].AsFloat, 9);
        if LaCote <> 0 then Result := LaCote else
          if LeTaux <> 0 then Result := 1 / LeTaux else Result := 1;
      end;
      Ferme(QQ);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ArrondiMontantLigne : Arrondi le montant d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.ArrondiMontantLigne(ARow: integer; MntLigne: Double): Double;
var TOBM: TOB;
  CodeArr: string;
begin
  Result := MntLigne;
  TOBM := GetTobModePaie(ARow);
  if TOBM <> nil then
  begin
    CodeArr := TOBM.GetValue('MP_ARRONDIFO');
    if CodeArr <> '' then Result := ArrondirPrix(CodeArr, MntLigne);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculeLigne : Calcule le montant d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.CalculeLigne(ARow: integer): Double;
var MntSaisi: Double;
begin
  MntSaisi := Valeur(GSReg.Cells[SG_Mont, ARow]);
  Result := ArrondiMontantLigne(ARow, MntSaisi);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculeMontantLigne : convertit et arrondit le montant d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.CalculeMontantLigne(ARow: Integer; Montant: Double): Double;
var TOBD: TOB;
  MntDev: Double;
begin
  Result := 0;
  if Montant <> 0 then
  begin
    TOBD := GetTobDevise(ARow);
    MntDev := ToxConvToDev(Montant, DEV.Code, TOBD.GetValue('D_DEVISE'), TOBPiec.GetValue('GP_DATEPIECE'), TOBDev);
    Result := ArrondiMontantLigne(ARow, MntDev);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LigneVautReste : vérifie si le montant d'une ligne vaut le reste à payer
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.LigneVautReste(ARow: Integer): Boolean;
var Mnt, Reste: Double;
begin
  Mnt := Valeur(GSReg.Cells[SG_Mont, ARow]);
  Reste := CalculeMontantLigne(ARow, FReste.Value);
  Result := FOCompareMontant(Mnt, Reste);
end;

{==============================================================================================}
{=============================== MANIPULATION DES LIGNES ======================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  TOBEcheLigneExiste : Crée si besoin la TOB de la ligne courante de la grille
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.TOBEcheLigneExiste;
begin
  if Action = taConsult then Exit;
  CreeTOBEche(GSReg.Row);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheLigne : Affichage d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.AfficheLigne(ARow: integer);
var TOBL: TOB;
  Ind: integer;
begin
  TOBL := GetTobEcheance(ARow);
  if TOBL = nil then Exit;
  TOBL.PutLigneGrid(GSReg, ARow, False, False, LesColonnes);
  for Ind := 1 to GSReg.ColCount - 1 do FormateZoneSaisie(Ind, ARow);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheLigneLCD : Affichage d'une ligne sur le LCD
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.AfficheLigneLCD(ARow: integer);
var Texte: string;
  Montant: double;
  {$IFDEF FOS5}
  NbDec: Integer;
  Sigle: string;
  {$ENDIF}
begin
  if FormTicket = nil then Exit;
  Texte := '';
  Montant := 0;
  if GSReg.Cells[SG_Mode, ARow] = '' then
  begin
    // Affichage du reste à payer
    if FReste.Value > 0 then
    begin
      Texte := MsgBox.Mess[15];
      Montant := FReste.Value;
    end;
    if FReste.Value < 0 then
    begin
      Texte := MsgBox.Mess[18];
      Montant := (FReste.Value * -1);
    end;
    if (FReste.Value = 0) and (ARow > 0) then
    begin
      // Affichage du montant de la ligne précédente
      if SG_Lib <> -1 then Texte := GSReg.Cells[SG_Lib, ARow - 1];
      if SG_Mont <> -1 then Montant := Valeur(GSReg.Cells[SG_Mont, ARow - 1]);
      if Montant < 0 then
      begin
        Texte := MsgBox.Mess[18];
        {$IFDEF FOS5}
        Montant := (Montant * -1);
        {$ENDIF}
      end;
      {$IFDEF FOS5}
      Dec(ARow);
      {$ENDIF}
    end;
  end else
  begin
    // Affichage du montant de la ligne
    if SG_Lib <> -1 then Texte := GSReg.Cells[SG_Lib, ARow];
    {$IFDEF FOS5}
    if SG_Mont <> -1 then Montant := Valeur(GSReg.Cells[SG_Mont, ARow]);
    {$ENDIF}
  end;
  {$IFDEF FOS5}
  Sigle := GetSymboleDevise('', ARow);
  NbDec := GetNbDecDevise(ARow);
  TFTicketFO(FormTicket).AffichageLCD(Texte, 0, Montant, NbDec, Sigle, [ofInterne, ofClient], qaLibreTexte, False);
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CompleteLesLignes : Compléte et affiche les lignes d'échéances
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.CompleteLesLignes;
var TOBL: TOB;
  Stg: string;
  Ind: integer;
begin
  for Ind := 0 to TOBEche.Detail.Count - 1 do
  begin
    TOBL := TOBEche.Detail[Ind];
    if TOBL.GetValue('GPE_MODEPAIE') <> '' then
    begin
      if not TOBL.FieldExists('GPE_LIBELLE') then
      begin
        Stg := GetLibelleModePaie(TOBL.GetValue('GPE_MODEPAIE'));
        TOBL.AddChampSupValeur('GPE_LIBELLE', Stg);
      end;
      if not TOBL.FieldExists('TYPEMODEPAIE') then
      begin
        Stg := GetTypeModePaie(TOBL.GetValue('GPE_MODEPAIE'));
        TOBL.AddChampSupValeur('TYPEMODEPAIE', Stg);
      end;
      if (Ind + 1) >= GSReg.RowCount then GSReg.RowCount := GSReg.RowCount + NbRowsPlus;
      AfficheLigne(Ind + 1);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// GenereLigneEcart : Génération automatique d'une échéance pour les écarts
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GenereLigneEcart;
var TOBM, TOBL: TOB;
  CodeMode: string;
  ARow: Integer;
  MontantP, MontantD: Double;
begin
  if MntEcart <> 0 then
  begin
    // Mode de paiement
    CodeMode := FODonneEcartChange;
    TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [CodeMode], False);
    if TOBM = nil then Exit;
    // Insertion en fin de liste
    DepileTOBEche(GSReg.RowCount, 1);
    ARow := TOBEche.Detail.Count + 1;
    InsereTOBEche(ARow);
    TOBL := GetTobEcheance(ARow);
    if TOBL = nil then exit;
    TOBL.PutValue('GPE_NUMECHE', ARow);
    TOBL.PutValue('GPE_MODEPAIE', CodeMode);
    // Initialisation de la TOB échéance à partir de la pièce
    InitLigneEche(TOBM, TOBL, False);
    // Mise à jour du montant de l'échéance
    TOBL.PutValue('GPE_MONTANTENCAIS', MntEcart);
    ConvertMontant(MntEcart, V_PGI.DevisePivot, MontantP, MontantD);
    // Montant en devise de tenu
    TOBL.PutValue('GPE_MONTANTECHE', MontantP);
    // Montant en devise de la pièce
    TOBL.PutValue('GPE_MONTANTDEV', MontantD);
    MntEcart := 0;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InitLigneEche : initialisation d'une ligne d'échéance
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.InitLigneEche(TOBMode, TOBLigne: TOB; ForceRaz: boolean);
begin
  if (TOBMode = nil) or (TOBLigne = nil) then exit;
  // Initialisation de la TOB échéance à partir de la pièce
  if TOBLigne.GetValue('GPE_NATUREPIECEG') = '' then
  begin
    TOBLigne.PutValue('GPE_NATUREPIECEG', TOBPiec.GetValue('GP_NATUREPIECEG'));
    TOBLigne.PutValue('GPE_DATEPIECE', TOBPiec.GetValue('GP_DATEPIECE'));
    TOBLigne.PutValue('GPE_SOUCHE', TOBPiec.GetValue('GP_SOUCHE'));
    TOBLigne.PutValue('GPE_NUMERO', TOBPiec.GetValue('GP_NUMERO'));
    TOBLigne.PutValue('GPE_INDICEG', TOBPiec.GetValue('GP_INDICEG'));
    TOBLigne.PutValue('GPE_DEVISE', TOBPiec.GetValue('GP_DEVISE'));
    TOBLigne.PutValue('GPE_TAUXDEV', TOBPiec.GetValue('GP_TAUXDEV'));
    TOBLigne.PutValue('GPE_COTATION', TOBPiec.GetValue('GP_COTATION'));
    TOBLigne.PutValue('GPE_DATETAUXDEV', TOBPiec.GetValue('GP_DATETAUXDEV'));
    TOBLigne.PutValue('GPE_SAISIECONTRE', TOBPiec.GetValue('GP_SAISIECONTRE'));
    TOBLigne.PutValue('GPE_NUMZCAISSE', TOBPiec.GetValue('GP_NUMZCAISSE'));
    TOBLigne.PutValue('GPE_TIERS', TOBPiec.GetValue('GP_TIERS'));
  end;
  if TOBLigne.GetValue('GPE_CAISSE') = '' then TOBLigne.PutValue('GPE_CAISSE', FOCaisseCourante);
  // Libellé du mode de paiement
  TOBLigne.PutValue('GPE_LIBELLE', TOBMode.GetValue('MP_LIBELLE'));
  TOBLigne.PutValue('TYPEMODEPAIE', TOBMode.GetValue('MP_TYPEMODEPAIE'));
  // Devise du mode de paiement
  if (ForceRaz) or (TOBLigne.GetValue('GPE_DEVISEESP') = '') or
    (TOBLigne.GetValue('GPE_DEVISEESP') <> TOBMode.GetValue('MP_DEVISEFO')) then
  begin
    TOBLigne.PutValue('GPE_DEVISEESP', TOBMode.GetValue('MP_DEVISEFO'));
    TOBLigne.PutValue('GPE_MONTANTENCAIS', 0);
    TOBLigne.PutValue('GPE_MONTANTECHE', 0);
    TOBLigne.PutValue('GPE_MONTANTDEV', 0);
  end;
  // Date d'échéance
  if (TOBLigne.GetValue('GPE_DATEECHE') = IDate1900) or (TOBMode.GetValue('MP_TYPEMODEPAIE') <> TYPEPAIECHQDIFF) then
  begin
    TOBLigne.PutValue('GPE_DATEECHE', TOBPiec.GetValue('GP_DATEPIECE'));
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LigneVide : indique si une ligne est renseignée ou non
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.LigneVide(ARow: Integer): Boolean;
begin
  Result := (Trim(GSReg.Cells[SG_Mode, ARow]) = '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InitInfoComplement : initialise les informations complémentaires d'une échéance en fonction du type de mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.InitInfoComplement(TOBL, TOBM: TOB);
var
  Montant: Double;
  TypeMode: string;
  RazNumCB, RazNumCHQ, RazNumAuto: boolean;
begin
  if TOBL = nil then Exit;

  RazNumCB := True;
  RazNumCHQ := True;
  RazNumAuto := True;
  if TOBM <> nil then
  begin
    if TOBM.GetValue('MP_AVECINFOCOMPL') = 'X' then
    begin
      TypeMode := TOBM.GetValue('MP_TYPEMODEPAIE');
      Montant := TOBL.GetValue('GPE_MONTANTENCAIS');
      if ((TypeMode = TYPEPAIECHEQUE) or (TypeMode = TYPEPAIECHQDIFF)) and (Montant > 0) then
        RazNumCHQ := False;
      if TypeMode = TYPEPAIECB then
        RazNumCB := False;
    end;
    if TOBM.GetValue('MP_AVECNUMAUTOR') = 'X' then RazNumAuto := False;
  end;

  if RazNumCB then
  begin
    TOBL.PutValue('GPE_CBINTERNET', '');
    TOBL.PutValue('GPE_CBNUMCTRL', '');
    TOBL.PutValue('GPE_DATEEXPIRE', '');
    TOBL.PutValue('GPE_TYPECARTE', '');
    TOBL.PutValue('GPE_CBLIBELLE', '');
  end;
  if RazNumCHQ then TOBL.PutValue('GPE_NUMCHEQUE', '');
  if RazNumAuto then TOBL.PutValue('GPE_CBNUMAUTOR', '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MarqueUneLigneRemiseBnq : marque une échéance comme remises en banque en fonction du type de mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.MarqueUneLigneRemiseBnq(TOBL, TOBM: TOB);
var
  Montant: Double;
  TypeMode, ChqDifUtil: string;
begin
  if TOBL = nil then Exit;
  ChqDifUtil := 'X';
  TypeMode := '';
  Montant := TOBL.GetValue('GPE_MONTANTENCAIS');
  if TOBM <> nil then TypeMode := TOBM.GetValue('MP_TYPEMODEPAIE');
  if ((TypeMode = TYPEPAIECHEQUE) or (TypeMode = TYPEPAIECHQDIFF)) and (Montant > 0) then ChqDifUtil := '-';
  TOBL.PutValue('GPE_CHQDIFUTIL', ChqDifUtil);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MarqueRemiseBnq : marque les échéances comme remises en banque en fonction du type de mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.MarqueRemiseBnq;
var
  TOBL, TOBM: TOB;
  Ind: Integer;
  CodeMode: string;
begin
  if Action <> taCreat then Exit;
  for Ind := 1 to TOBEche.Detail.Count do
  begin
    TOBL := TOBEche.Detail[(Ind - 1)];
    if TOBL <> nil then
    begin
      CodeMode := TOBL.GetValue('GPE_MODEPAIE');
      TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [CodeMode], False);
      MarqueUneLigneRemiseBnq(TOBL, TOBM);
    end;
  end;
end;

{==============================================================================================}
{=============================== ACTIONS LIEES AU GRID ========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  QuelChampLigne : retourne le rang d'un champ
///////////////////////////////////////////////////////////////////////////////////////

function QuelChampLigne(iTableLigne: integer; NomChamp: string): integer;
var Ind: integer;
begin
  Result := -1;
  for Ind := 1 to High(V_PGI.Dechamps[iTableLigne]) do if V_PGI.DEChamps[iTableLigne, Ind].Nom = NomChamp then
    begin
      Result := Ind;
      Break;
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EtudieColsGrid : recherche le champ associé à chaque colonne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.EtudieColsGrid(GS: THGrid);
var NomCol, LesCols: string;
  icol, ichamp, iTableLigne: integer;
begin
  LesCols := GS.Titres[0];
  icol := 0;
  iTableLigne := PrefixeToNum('GPE');
  repeat
    NomCol := Uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol <> '' then
    begin
      ///ichamp := QuelChampLigne(iTableLigne,NomCol) ; *** pb en eAgl ***
      ichamp := ChampToNum(NomCol);
      if ichamp >= 0 then
      begin
        if Pos('X', V_PGI.Dechamps[iTableLigne, ichamp].Control) > 0 then
        begin
          // "X" interdit la saisie sauf si aussi "Y" pour FrontOffice
          if ((Pos('Y', V_PGI.Dechamps[iTableLigne, ichamp].Control) > 0) and (ctxFO in V_PGI.PGIContexte))
            then else GS.ColLengths[icol] := -1;
        end;
        if NomCol = 'GPE_NUMECHE' then SG_NL := icol else
          if NomCol = 'GPE_MODEPAIE' then SG_Mode := icol else
          if NomCol = 'GPE_MONTANTENCAIS' then SG_Mont := icol else
          if NomCol = 'GPE_DEVISEESP' then SG_Dev := icol else
          if NomCol = 'GPE_DATEECHE' then SG_Ech := icol else
          ;
      end else
      begin
        if NomCol = '(GPE_LIBELLE)' then SG_Lib := icol else
      end;
    end;
    Inc(icol);
  until ((LesCols = '') or (NomCol = ''));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InitLesCols : initialise la numero de colonne associé aux champs qui subissent un traitement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.InitLesCols;
begin
  SG_NL := -1;
  SG_Mode := -1;
  SG_Lib := -1;
  SG_Mont := -1;
  SG_Dev := -1;
  SG_Ech := -1;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormateCols : définit les caractéristiques de chaque colonne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormateCols;
var Ind: Integer;
begin
  GSReg.ListeParam := 'GCREGLEFO';
  LesColonnes := GSReg.Titres[0];
  LesColonnes := FindEtReplace(LesColonnes, '(GPE_LIBELLE)', 'GPE_LIBELLE', False);
  InitLesCols;
  EtudieColsGrid(GSReg);
  // Code devise
  if SG_Dev <> -1 then GSReg.ColFormats[SG_Dev] := 'CB=TTDEVISETOUTES';
  // Date d'échéance
  if SG_Ech <> -1 then
  begin
    GSReg.ColTypes[SG_Ech] := 'D';
    GSReg.ColFormats[SG_Ech] := ShortDateFormat;
  end;
  // Uniquement le mode de paiement, le montant et la date d'échéance sont saisissables
  for Ind := 0 to GSReg.ColCount - 1 do
  begin
    if (Ind <> SG_Mode) and (Ind <> SG_Mont) and (Ind <> SG_Ech) then GSReg.ColLengths[Ind] := -1;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormateZoneSaisie : formate une cellules en fonction du type de données
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormateZoneSaisie(ACol, ARow: Longint);
var St, StC: string;
  Nbdec: Integer;
  TOBD: TOB;
begin
  St := GSReg.Cells[ACol, ARow];
  StC := St;
  if ACol = SG_Mode then StC := Uppercase(Trim(St)) else
    if ACol = SG_Mont then
  begin
    TOBD := GetTobDevise(ARow);
    if TOBD = nil then Nbdec := V_PGI.OkDecV else NbDec := TOBD.GetValue('D_DECIMALE');
    ////StC := StrfMontant(Valeur(St), 12, NbDec, TOBD.GetValue('D_SYMBOLE'), True) ;
    StC := StrF00(Valeur(St), NbDec);
  end else
    if ACol = SG_Ech then
  begin
    if (St = '') or not (IsValidDate(St)) then StC := DateToStr(V_PGI.DateEntree);
  end else ;
  GSReg.Cells[ACol, ARow] := StC;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ZoneAccessible : vérifie si une cellule est saisissable
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.ZoneAccessible(ACol, ARow: Longint): boolean;
var TOBL, TOBM: TOB;
begin
  Result := True;
  // si le mode de paiement n'est pas défini, les autres colonnes ne sont pas saisissable.
  if (ACol <> SG_Mode) and (LigneVide(ARow)) then
  begin
    Result := False;
    Exit;
  end;
  TOBL := GetTobEcheance(ARow);
  if TOBL <> nil then
  begin
    if TOBL.FieldExists('ACCORDTPE') then
    begin
      // Si l'autorisation du TPE est accordé, la ligne n'est pas modifiable
      Result := False;
      Exit;
    end;
    if FOChampSupValeurNonVide(TOBL, 'GOC_NUMPIECELIEN') then
    begin
      // Si le réglement est lié, la ligne n'est pas modifiable
      Result := False;
      Exit;
    end;
  end;
  // la saisie de l'échéance n'est nécessaire que pour les chèques différés
  if ACol = SG_Ech then
  begin
    TOBM := GetTobModePaie(ARow);
    if (TOBM = nil) or (TOBM.GetValue('MP_TYPEMODEPAIE') <> TYPEPAIECHQDIFF) then Result := False;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ZoneSuivanteOuOk : recherche la zone suivante saisissable
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var Sens, ii: integer;
  OldEna, ChgLig: boolean;
begin
  OldEna := GSReg.SynEnabled;
  GSReg.SynEnabled := False;
  Sens := -1;
  ChgLig := (GSReg.Row <> ARow);
  if GSReg.Row > ARow then Sens := 1 else if ((GSReg.Row = ARow) and (ACol < GSReg.Col)) then Sens := 1;
  ACol := GSReg.Col;
  ARow := GSReg.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if (Sens = 1) or ((ACol = GSReg.FixedCols) and (ARow = 1)) then
    begin
      Sens := 1;
      if ((ACol = GSReg.ColCount - 1) and (ARow = GSReg.RowCount - 1)) then Break;
      if ChgLig then
      begin
        ACol := GSReg.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GSReg.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GSReg.FixedCols;
      end;
    end else
    begin
      if ((ACol = GSReg.FixedCols) and (ARow = 1)) then Break;
      if ChgLig then
      begin
        if ARow <= GSReg.FixedRows then ACol := GSReg.ColCount else ACol := GSReg.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol > GSReg.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GSReg.ColCount - 1;
      end;
    end;
  end;
  GSReg.SynEnabled := OldEna;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieTOBEche : vérifie l'intégrité d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.VerifieTOBEche(NoLigne: Integer): Integer;
var TOBL: TOB;
  MdpDiffere: boolean;
  ModePaie: string;
  MontantE: Double;
  DateEch: TDateTime;
begin
  Result := -1;
  TOBL := TOBEche.Detail[NoLigne];
  if TOBL <> nil then
  begin
    ModePaie := TOBL.GetValue('GPE_MODEPAIE');
    Result := VerifModePaie(ModePaie, True);
    if Result <> 0 then Exit;
    MontantE := TOBL.GetValue('GPE_MONTANTENCAIS');
    Result := VerifMontant(MontantE, ModePaie, False);
    if Result <> 0 then Exit;
    DateEch := TOBL.GetValue('GPE_DATEECHE');
    MdpDiffere := (GetTypeModePaie(ModePaie) = TYPEPAIECHQDIFF);
    Result := VerifDateEch(DateEch, MdpDiffere);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieCorrecte : vérifie l'intégrité de la saisie
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.SaisieCorrecte: Boolean;
var ErrNo, Ind: Integer;
begin
  ErrNo := 0;
  // Contrôle de chaque ligne
  for Ind := 0 to TOBEche.Detail.Count - 1 do if ErrNo = 0 then ErrNo := VerifieTOBEche(Ind) else Break;
  // Génération automatique d'une échéance pour les écarts
  if MntEcart <> 0 then GenereLigneEcart;
  // Vérifie si le nombre maximum d'échéances n'est pas dépassé.
  if (Errno = 0) and (TOBEche.Detail.Count > MaxEche) then Errno := 8;
  // Le reste à payer doit être nulle
  if ErrNo = 0 then
  begin
    AfficheTotal;
    if FReste.Value <> 0 then ErrNo := 7;
  end;
  {**
  // Vérification de la cohérence des règlements liés
  if ErrNo = 0 then
  begin
    ErrNo := FOVerifReglementLie(TOBPiece, TOBEches, AnnulPiece, NoLig);
    if NoErr <> 0 then
    begin
      if NoErr = 2 then Result := erRegleLie
      else Result := erRegleDejaLie;
      Inc(NoLig, GS.FixedRows);
    end;
  end;
  **}
  // Message d'erreur
  if ErrNo > 0 then
  begin
    MsgBox.Execute(ErrNo, '', '');
    if GSReg.CanFocus then GSReg.SetFocus;
  end;
  Result := (ErrNo = 0);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CellIsModified : vérifie si la cellule courante a été modifiée
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.CellIsModified(ACol, ARow: Integer): Boolean;
begin
  if ACol = SG_Mont then
    Result := (Valeur(GSReg.Cells[ACol, ARow]) <> Valeur(StCellCur))
  else if ACol = SG_Ech then
    Result := ((not IsValidDate(GSReg.Cells[ACol, ARow])) or
      (not IsValidDate(StCellCur)) or
      (StrToDate(GSReg.Cells[ACol, ARow]) <> StrToDate(StCellCur)))
  else
    Result := (GSReg.Cells[ACol, ARow] <> StCellCur);
end;

{==============================================================================================}
{============================ CALCULS ET AFFICHAGE DES TOTAUX =================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  InitTotaux : initialise les totaux de bas d'écran
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.InitTotaux;
begin
  FTotalBrut.Value := TOBPiec.GetValue('GP_TOTALTTCDEV');
  ChangeMask(FTotalBrut, DEV.Decimale, DEV.Symbole);
  FAPayer.Value := TOBPiec.GetValue('GP_TOTALTTCDEV');
  ChangeMask(FApayer, DEV.Decimale, DEV.Symbole);
  FReste.Value := FAPayer.Value;
  ChangeMask(FReste, DEV.Decimale, DEV.Symbole);
  FEncaisse.Value := 0;
  ChangeMask(FEncaisse, DEV.Decimale, DEV.Symbole);
  if FTotalBrut.Value < 0 then
  begin
    TFAPayer.Caption := MsgBox.Mess[16];
    TFEncaisse.Caption := MsgBox.Mess[17];
    TFTotAPayer.Caption := MsgBox.Mess[16];
    TFTotRegle.Caption := MsgBox.Mess[17];
    TFTotARendre.Caption := MsgBox.Mess[18];
  end else
  begin
    TFAPayer.Caption := MsgBox.Mess[13];
    TFEncaisse.Caption := MsgBox.Mess[14];
    TFTotAPayer.Caption := MsgBox.Mess[13];
    TFTotRegle.Caption := MsgBox.Mess[14];
    TFTotARendre.Caption := MsgBox.Mess[15];
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheTotal : calcule et affiche les totaux en bas d'écran
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.AfficheTotal;
var Ind: integer;
  Total: Double;
  TOBL: TOB;
begin
  // Total déjà encaisser
  Total := 0;
  for Ind := 0 to TOBEche.Detail.Count - 1 do
  begin
    TOBL := TOBEche.Detail[Ind];
    Total := Total + TOBL.GetValue('GPE_MONTANTDEV');
  end;
  FEncaisse.Value := Arrondi(Total, 6);
  // Reste à Encaisser
  FReste.Value := Arrondi((FAPayer.Value - FEncaisse.Value), 6);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculReste : calcule le reste à payer
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.CalculReste: Double;
var Ind: integer;
  Total: Double;
  TOBL: TOB;
begin
  // Total déjà encaisser
  Total := 0;
  for Ind := 0 to TOBEche.Detail.Count - 1 do
  begin
    TOBL := TOBEche.Detail[Ind];
    Total := Total + TOBL.GetValue('GPE_MONTANTDEV');
  end;
  // Reste à Encaisser
  Result := Arrondi((FAPayer.Value - Total), 6);
  ;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheZonePied : affichage des Labels correspondants aux totaux de bas d'écran
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.AfficheZonePied(Sender: TObject);
//////////////////////////////////////////////////////////////////////////////////////////
  procedure CopieText(Orig: TObject; Dest: TComponent);
  var Stg: string;
    LCD: TLCDLabel;
  begin
    if Dest <> nil then
    begin
      if Dest is TLCDLabel then
      begin
        LCD := TLCDLabel(Dest);
        Stg := THNumEdit(Orig).Text;
        // On supprime le symbole monétaire
        Delete(Stg, Pos(DEV.Symbole, Stg), Length(DEV.Symbole));
        Stg := Trim(Stg);
        // On compléte par des espaces
        while length(Stg) < LCD.NoOfChars do Stg := ' ' + Stg;
        LCD.Caption := Stg;
      end;
      if Dest is THLabel then THLabel(Dest).Caption := THNumEdit(Orig).Text;
      if Dest is TFlashingLabel then TFlashingLabel(Dest).Caption := THNumEdit(Orig).Text;
    end;
  end;
  //////////////////////////////////////////////////////////////////////////////////////////
begin
  if Sender = nil then Exit;
  CopieText(Sender, FindComponent('L' + THNumEdit(Sender).Name));
end;

{==============================================================================================}
{=============================== EVENEMENTS DE LE GRID ========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  GSRegCellEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GSRegCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  if not Cancel then
  begin
    CreeTOBEche(ARow);
    StCellCur := GSReg.Cells[GSReg.Col, GSReg.Row];
    if ACol = SG_Mode then PrepareModePaie(ACol, ARow, Cancel) else
      if ACol = SG_Mont then PrepareMontant(ACol, ARow, Cancel) else
      ;
    GSReg.ElipsisButton := (GSReg.Col = SG_Mode);
    ////   StCellCur := GSReg.Cells[GSReg.Col, GSReg.Row] ;
    {$IFDEF FOS5}
    if (FormTicket <> nil) and (FormTicket is TFTicketFO) then TFTicketFO(FormTicket).AffichePaveChoisi(20 + GSReg.Col);
    {$ENDIF}
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRegCellExit
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GSRegCellExit(Sender: TObject; var ACol, ARow: Integer;
  var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  //if (Action <> taConsult) and (GSReg.Cells[ACol, ARow] <> StCellCur) then
  if (Action <> taConsult) and (CellIsModified(ACol, ARow)) then
  begin
    CreeTOBEche(ARow);
    FormateZoneSaisie(ACol, ARow);
    if ACol = SG_Mode then TraiteModePaie(ACol, ARow, Cancel) else
      if ACol = SG_Mont then TraiteMontant(ACol, ARow, Cancel) else
      if ACol = SG_Ech then TraiteDateEch(ACol, ARow, Cancel) else
      ;
  end;
  AfficheLigneLCD(GSReg.Row);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRowEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GSRegRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if GSReg.Row >= GSReg.RowCount - 1 then GSReg.RowCount := GSReg.RowCount + NbRowsPlus;
  if Action <> taConsult then CreeTOBEche(GSReg.Row);
  GereEnabled(Ou);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRowExit
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GSRegRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  if Action = taConsult then Exit;
  DepileTOBEche(Ou, GSReg.Row);
  if (GSReg.Row > 0) and (GSReg.Row > Ou) and (LigneVide(GSReg.Row - 1)) then
  begin
    Cancel := True;
  end else
    if (GSReg.Cells[SG_Mode, Ou] <> '') and (Valeur(GSReg.Cells[SG_Mont, Ou]) = 0) then
  begin
    MsgBox.Execute(6, '', '');
    Cancel := True;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GSRegEnter(Sender: TObject);
var bc, Cancel: boolean;
  ACol, ARow: integer;
begin
  bc := False;
  Cancel := False;
  ACol := GSReg.Col;
  ARow := GSReg.Row;
  GSRegRowEnter(GSReg, GSReg.Row, bc, False);
  GSRegCellEnter(GSReg, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSElipsisClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GSRegElipsisClick(Sender: TObject);
var ChoixOK: Boolean;
  Key: Char;
begin
  if Action = taConsult then Exit;
  ChoixOK := False;
  CreeTOBEche(GSReg.Row);
  if (GSReg.Col = SG_Mode) and (Sender is TControl) then
  begin
    ChoixOK := GetModePaieLookUp(TControl(Sender));
  end else
    if (GSReg.Col = SG_Ech) and (Sender is TControl) then
  begin
    Key := '*';
    V_PGI.ParamDateKeyProc(TControl(Sender), Key);
    ChoixOK := True;
  end;
  if ChoixOK then EnvoieToucheGrid;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GSRegDblClick(Sender: TObject);
begin
  GSRegElipsisClick(Sender);
  {****
  if GSReg.Col = SG_Mode then ZoomOuChoixModePaie(Sender) else
   if GSReg.Col = SG_Ech then GSRegElipsisClick(Sender) else ;
  ****}
end;

{==============================================================================================}
{=============================== ACTIONS SUR LA FORME =========================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  BloqueFiche : désactive ou active certains contrôles de la fiche selon le mode de saisie
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BloqueFiche(ActionSaisie: TActionFiche; Bloc: boolean);
begin
  if ActionSaisie = taConsult then
  begin
    BValider.Enabled := not Bloc;
    bNewligne.Enabled := not Bloc;
    bDelLigne.Enabled := not Bloc;
    bSolde.Enabled := not Bloc;
    bRemise.Enabled := not Bloc;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GereEnabled : active les boutons en fonction du contenu de la ligne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GereEnabled(ARow: integer);
var TOBL: TOB;
  EcheOK: Boolean;
begin
  TOBL := GetTobEcheance(ARow);
  EcheOK := ((TOBL <> nil) and (not LigneVide(ARow)));
  BMenuZoom.Enabled := EcheOK;
  DetailLigne.Enabled := EcheOK;
  VoirModePaie.Enabled := EcheOK;
  {$IFDEF EAGLCLIENT}
  VoirDevise.Enabled := False;
  {$ELSE}
  VoirDevise.Enabled := EcheOK;
  {$ENDIF}
  if Action = taConsult then Exit;
  bDelLigne.Enabled := EcheOK;
  bSolde.Enabled := EcheOK;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EnvoieToucheGrid : simule la frappe de la touche TAB dans le Grid
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.EnvoieToucheGrid;
begin
  if Action = taConsult then Exit;
  FOSimuleClavier(VK_TAB);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ClickValide : traitement de la touche Valide
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.ClickValide: Boolean;
var io: TIOErr;
  bc, cancel: boolean;
  ACol, ARow: integer;
begin
  // Tests et actions préalables
  Result := True;
  if (Action = taConsult) or (not TOBEche.IsOneModifie) then
  begin
    ModalResult := mrOk;
    Exit;
  end;
  DepileTOBEche(GSReg.Row, 0);
  bc := False;
  Cancel := False;
  ACol := GSReg.Col;
  ARow := GSReg.Row;
  GSRegCellExit(GSReg, ACol, ARow, Cancel);
  if Cancel then
  begin
    GSReg.Col := ACol;
    GSReg.Row := ARow;
    EmpileTOBEche(GSReg.Row);
    Cancel := False;
    GSRegCellEnter(GSReg, ACol, ARow, Cancel);
    Result := False;
    Exit;
  end;
  GSRegRowExit(GSReg, ARow, bc, False);
  // Contrôle d'intégrité
  if not SaisieCorrecte then
  begin
    EmpileTOBEche(GSReg.Row);
    Result := False;
    Exit;
  end;
  // Marque les échéances comme remises en banque en fonction du type de mode de paiement
  MarqueRemiseBnq;
  // Autorisation du TPE
  {$IFDEF FOS5}
  if not ValideCB then
  begin
    EmpileTOBEche(GSReg.Row);
    Result := False;
    Exit;
  end;
    {$ENDIF}
  // Mise à jour de la base
  if DepuisPiece then
  begin
    // Impression des chèques
    {$IFDEF FOS5}
    if Action = taCreat then ImpressionCheque;
    {$ENDIF}
    ModalResult := mrOk;
    ValiderOk := True;
    {$IFDEF FOS5}
    if (FormTicket <> nil) and (FormTicket is TFTicketFO) then
    begin
      ////FORecopieTob(TobPiec, TFTicketFO(FormTicket).TOBPiece) ;
      MiseAJourPiece;
      FORecopieTob(TOBEche, TFTicketFO(FormTicket).TOBEches);
    end;
    {$ENDIF}
    Close;
  end else
  begin
    BValider.Enabled := False;
    io := Transactions(EnregistreSaisie, 2);
    BValider.Enabled := True;
    case io of
      oeOk: ;
      oeUnknown:
        begin
          MessageAlerte(MsgBox.Mess[9]);
          Result := False;
          EmpileTOBEche(GSReg.Row);
        end;
      oeSaisie:
        begin
          MessageAlerte(MsgBox.Mess[10]);
          Result := False;
          EmpileTOBEche(GSReg.Row);
        end;
    end;
    Close;
  end;
  if Result then ModalResult := mrOk;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ClickAbandon : traitement de la touche Abandon
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ClickAbandon;
begin
  {$IFDEF FOS5}
  if (FormTicket <> nil) and (FormTicket is TFTicketFO) then TFTicketFO(FormTicket).DepuisEcheance := True;
  {$ENDIF}
  Close;
end;

{==============================================================================================}
{=============================== EVENEMENTS DE LA FORME =======================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormCreate(Sender: TObject);
begin
  CreeLesTobs;
  FormTicket := nil;
  ValiderOk := False;
  RenduEnCours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormShow(Sender: TObject);
var Ind: Integer;
begin
  FormateCols;
  case Action of
    taCreat:
      begin
        GSReg.RowCount := NbRowsInit;
        for Ind := 1 to NbRowsInit do CreeTOBEche(Ind);
        ChargeDevisePiece;
        ChargeModePaie;
        CompleteLesLignes;
      end;
    taConsult,
      taModif:
        begin
        ChargeTobPiec;
        ChargeDevisePiece;
        ChargeModePaie;
        ChargeTobEcheance;
        CompleteLesLignes;
        {$IFDEF FOS5}
        VerifAction;
        {$ENDIF}
      end;
  end;
  {$IFDEF FOS5}
  BRemise.Visible := (DepuisPiece);
  BRemise.Enabled := (TOBPiec.GetValue('GP_FACTUREHT') = '-');
  {$ELSE}
  BRemise.Visible := False;
  BRemise.Enabled := False;
  {$ENDIF}
  if action = taConsult then
  begin
    BloqueFiche(Action, True);
    BValider.Enabled := False;
  end;
  HMTrad.ResizeGridColumns(GSReg);
  AffecteGrid(GSReg, Action);
  InitTotaux;
  AfficheTotal;
  // Appel de la fonction d'empilage dans la liste des fiches
  if not DepuisPiece then AglEmpileFiche(Self);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormResize
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormResize(Sender: TObject);
begin
  LFReste.Left := pPied.width - LFReste.Width - 4;
  LDETAXE.Left := LFReste.Left + LFReste.Width - LDETAXE.Width - 16;
  SOLDETIERS.Left := LFEncaisse.Left;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormDestroy
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  if not DepuisPiece then AglDepileFiche;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormClose
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  {$IFDEF FOS5}
  {if not DepuisPiece then} LibereLesTobs;
  // On détruit un lien sur la forme de saisie des tickets
  if (FormTicket <> nil) and (FormTicket is TFTicketFO) then
  begin
    if TFTicketFO(FormTicket).DockBottom.Visible then
    begin
      Valide97.DockedTo := DockBottom;
      Outils97.DockedTo := DockBottom;
      TFTicketFO(FormTicket).Valide97.Visible := True;
      TFTicketFO(FormTicket).Outils97.Visible := True;
    end;
    TFTicketFO(FormTicket).FormEcheance := nil;
    if ModalResult = mrOk then TFTicketFO(FormTicket).ClickValide
    else if TFTicketFO(FormTicket).GS.CanFocus then TFTicketFO(FormTicket).GS.SetFocus;
    FormTicket := nil;
  end;
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormKeyDown
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var OkG: Boolean;
begin
  OkG := (Screen.ActiveControl = GSReg);
  if Key = VK_RECHERCHE then
  begin
    if ((OkG) and (Shift = []) and ((GSReg.Col = SG_Mode) or (GSReg.Col = SG_Ech))) then
    begin
      Key := 0;
      GSRegElipsisClick(GSReg);
    end;
  end else
    case Key of
      VK_RETURN: if RemiseTicket.Visible then
        begin
          FOSimuleClavier(VK_TAB);
          Key := 0;
        end else Key := VK_TAB;
      VK_SPACE: if ((OkG) and (GSReg.Col = SG_Ech) and (Shift = [])) then
        begin
          Key := 0;
          GSRegElipsisClick(GSReg);
        end;
      VK_INSERT: if ((OkG) and (Shift = [])) then
        begin
          Key := 0;
          bNewligneClick(nil);
        end;
      VK_DELETE: if ((OkG) and (Shift = [ssCtrl])) then
        begin
          Key := 0;
          bDelLigneClick(nil);
        end;
      VK_F10: if Shift = [] then
        begin
          Key := 0;
          if RemiseTicket.Visible then BValiderRemiseClick(nil) else BValiderClick(nil);
        end;
      VK_ESCAPE: if Shift = [] then
        begin
          Key := 0;
          if RemiseTicket.Visible then BCancelRemiseClick(nil) else BAbandonClick(nil);
        end;
    end;
end;

{==============================================================================================}
{=============================== ACTIONS SUR LES BOUTONS ======================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  bNewligneClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.bNewligneClick(Sender: TObject);
var ARow: Integer;
begin
  if Action = taConsult then Exit;
  if ((GSReg.Row < 1) or (GSReg.Row > TOBEche.Detail.Count)) then Exit;
  ARow := GSReg.Row;
  GSReg.CacheEdit;
  GSReg.SynEnabled := False;
  // Insertion d'une ligne
  GSReg.InsertRow(ARow);
  InsereTOBEche(ARow);
  GSReg.Col := SG_Mode;
  GSReg.Row := ARow;
  GSReg.MontreEdit;
  GSReg.SynEnabled := True;
  RenumeroteLigne;
  AfficheLigneLCD(GSReg.Row);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bDelLigneClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.bDelLigneClick(Sender: TObject);
var ARow: integer;
begin
  if Action = taConsult then Exit;
  if ((GSReg.Row < 1) or (GSReg.Row > TOBEche.Detail.Count)) then Exit;
  ARow := GSReg.Row;
  ClickDel(ARow);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ClickDel : suppression d'une ligne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ClickDel(ARow: Integer);
var bc, cancel: boolean;
  ACol, NoLig, ErrNo: integer;
  TOBL: TOB;
begin
  if Action = taConsult then Exit;
  if ((ARow < 1) or (ARow > TOBEche.Detail.Count)) then Exit;
  ErrNo := 0;
  TOBL := GetTobEcheance(ARow);
  if TOBL = nil then exit;
  if TOBL.FieldExists('ACCORDTPE') then
    ErrNo := 19 // Autorisation accordée par le TPE
  else if FOChampSupValeurNonVide(TOBL, 'GOC_NUMPIECELIEN') then
    ErrNo := 28; // Réglement lié
  if ErrNo <> 0 then
  begin
    MsgBox.Execute(ErrNo, '', '');
    Exit;
  end;
  // Suppression de la ligne
  GSReg.CacheEdit;
  GSReg.SynEnabled := False;
  NoLig := GSReg.TopRow;
  GSReg.DeleteRow(ARow);
  if NoLig > 1 then GSReg.TopRow := (NoLig - 1) else GSReg.TopRow := 1;
  if (TOBEche.Detail.Count > 0) and (ARow > 0) and (ARow < TOBEche.Detail.Count) then
    TOBEche.Detail[ARow - 1].Free
  else
    InitTOBEche(TOBEche.Detail[ARow - 1], ARow, True);
  if GSReg.RowCount < NbRowsInit then GSReg.RowCount := GSReg.RowCount + NbRowsPlus;
  GSReg.MontreEdit;
  GSReg.SynEnabled := True;
  RenumeroteLigne;
  AfficheTotal;
  AfficheLigneLCD(GSReg.Row);
  bc := False;
  Cancel := False;
  GSReg.Col := SG_Mode;
  if GSReg.Row > GSReg.FixedRows then
  begin
    ARow := GSReg.Row;
    Dec(ARow);
    GSReg.Row := ARow;
  end;
  ACol := GSReg.Col;
  GSRegRowEnter(GSReg, ARow, bc, False);
  GSRegCellEnter(GSReg, ACol, ARow, Cancel);
  if Cancel then
  begin
    GSReg.Col := ACol;
    GSReg.Row := ARow;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SupprimeLigne : suppression de la ligne courante ou de la précédente si la courante est vide
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.SupprimeLigne;
var ARow: integer;
begin
  if Action = taConsult then Exit;
  if ((GSReg.Row < 1) or (GSReg.Row > TOBEche.Detail.Count)) then Exit;
  // Choix de la ligne à supprimer
  for ARow := GSReg.Row downto 0 do if not LigneVide(ARow) then Break;
  if ARow > 0 then ClickDel(ARow);
  if (ARow >= GSReg.FixedRows) and (ARow < GSReg.Row) then GSReg.Row := ARow;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BSoldeClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BSoldeClick(Sender: TObject);
var MntSolde: Double;
  TOBD: TOB;
  NbDec, ACol, ARow: Integer;
  Cancel: boolean;
begin
  if Action = taConsult then Exit;
  if GoRowSelect in GSReg.Options then Exit;
  if GSReg.Cells[SG_Mode, GSReg.Row] = '' then Exit;
  if not FOJaiLeDroit(53) then Exit;
  TOBD := GetTobDevise(GSReg.Row);
  if TOBD = nil then NbDec := V_PGI.OkDecV else NbDec := TOBD.GetValue('D_DECIMALE');
  MntSolde := Valeur(GSReg.Cells[SG_MONT, GSReg.Row])
            + ToxConvToDev(FReste.Value, DEV.Code, TOBD.GetValue('D_DEVISE'), TOBPiec.GetValue('GP_DATEPIECE'), TOBDev);
  MntSolde := ArrondiMontantLigne(GSReg.Row, MntSolde);
  MntRef := MntSolde;
  GSReg.Cells[SG_Mont, GSReg.Row] := StrF00(MntSolde, NbDec);
  GSReg.Col := SG_Mont;
  ACol := SG_Mont;
  ARow := GSReg.Row;
  Cancel := True;
  TraiteMontant(ACol, ARow, Cancel);
  if GSReg.CanFocus then GSReg.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BRemiseClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BRemiseClick(Sender: TObject);
begin
  if not BRemise.Enabled then Exit;
  SaisieRemiseTicket(0, '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BAbandonClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BAbandonClick(Sender: TObject);
begin
  if not BAbandon.Enabled then Exit;
  ClickAbandon;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValiderClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BValiderClick(Sender: TObject);
begin
  if not BValider.Enabled then Exit;
  ClickValide;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BAideClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DetailLigneClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.DetailLigneClick(Sender: TObject);
var
  TOBMain, TOBL, TOBM: TOB;
  TauxDev: double;
  Stg, CodeDev, CodeDevPiece, CodeMode: string;
  ModeUS: boolean;
begin
  // Détail de l'échéance en mémoire
  TOBL := GetTobEcheance(GSReg.Row);
  if TOBL = nil then Exit;
  if not FOJaiLeDroit(52) then Exit;
  TOBMain := TOB.Create('PIEDECHE', nil, -1);
  TOBMain.Dupliquer(TOBL, False, True);
  // Détermination de la devise de saisie de la pièce
  if ((TOBPiec.GetValue('GP_DEVISE') = V_PGI.DevisePivot) and
    ((TOBPiec.GetValue('GP_SAISIECONTRE') = '-') and (VH^.TenueEuro)) or
    ((TOBPiec.GetValue('GP_SAISIECONTRE') = 'X') and not (VH^.TenueEuro))) then CodeDevPiece := FODonneCodeEuro
  else CodeDevPiece := TOBPiec.GetValue('GP_DEVISE');
  // Ajout des codes devises piece, pivots et contre-valeurs
  TOBMain.AddChampSupValeur('GPE_DEVISEPIE', CodeDevPiece);
  TOBMain.AddChampSupValeur('GPE_DEVISEECH', V_PGI.DevisePivot);
  // Ajout du taux de conversion de la devise d'encaissement
  Stg := '';
  CodeDev := TOBL.GetValue('GPE_DEVISEESP');
  if ((CodeDev <> '') and (not FOTestCodeEuro(CodeDev)) and (CodeDev <> V_PGI.DevisePivot) and (CodeDev <> V_PGI.DeviseFongible)) then
  begin
    TauxDev := GetCotation(CodeDev, TOBPiec.GetValue('GP_DATEPIECE'));
    Stg := '1,00 ' + SIGLEEURO + ' = ' + StrfMontant(TauxDev, 12, 5, GetSymboleDevise(CodeDev, -1), True);
  end;
  TOBMain.AddChampSupValeur('GPE_TAUXDEVISEENC', Stg);
  // Ajout du taux de conversion de la devise de la pièce
  Stg := '';
  if ((DEV.Code <> '') and (not FOTestCodeEuro(DEV.Code)) and (DEV.Code <> V_PGI.DevisePivot) and (DEV.Code <> V_PGI.DeviseFongible)) then
  begin
    TauxDev := GetCotation(DEV.Code, TOBPiec.GetValue('GP_DATEPIECE'));
    Stg := '1,00 ' + SIGLEEURO + ' = ' + StrfMontant(TauxDev, 12, 5, DEV.Symbole, True);
  end;
  TOBMain.AddChampSupValeur('GPE_TAUXDEVISEECH', Stg);
  // Informations complémentaires
  CodeMode := TOBL.GetValue('GPE_MODEPAIE');
  TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [CodeMode], False);
  if TOBM <> nil then
  begin
    TOBMain.AddChampSupValeur('MP_TYPEMODEPAIE', TOBM.GetValue('MP_TYPEMODEPAIE'));
    TOBMain.AddChampSupValeur('MP_AVECINFOCOMPL', TOBM.GetValue('MP_AVECINFOCOMPL'));
    TOBMain.AddChampSupValeur('MP_AVECNUMAUTOR', TOBM.GetValue('MP_AVECNUMAUTOR'));
  end;
  Stg := TOBMain.GetValue('GPE_CBINTERNET');
  ModeUS := (TOBM.GetValue('MP_AFFICHNUMCBUS') = 'X');
  if Stg <> '' then TOBMain.PutValue('GPE_CBINTERNET', DeCrypteNoCarteCB(Stg, ModeUS));
  // Affichage du détail de l'échéance
  TheTob := TOBMain;
  AglLanceFiche('MFO', 'PIEDECHE', '', '', ActionToString(taConsult));
  if TOBMain <> nil then TOBMain.Free;
  TheTob := nil;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VoirModePaieClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.VoirModePaieClick(Sender: TObject);
begin
  ZoomModePaie;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VoirDeviseClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.VoirDeviseClick(Sender: TObject);
begin
  ZoomDevise;
end;

{==============================================================================================}
{=============================== SAISIE DE LA REMISE GLOBALE ==================================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  SaisieRemiseTicket : saisie d'une remise sur le ticket
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.SaisieRemiseTicket(Pourcent: Double; CodeDem: string);
var
  Ok: boolean;
begin
  if not DepuisPiece then Exit;
  if not BRemise.Enabled then Exit;
  if not FOJaiLeDroit(54) then Exit;
  Ok := False;
  FBaseRemise.Visible := not Ok;
  LFBaseRemise.Visible := not Ok;
  FBaseRemise.Value := FOCalculBaseRemiseGlobale(TOBPiec, Ok);
  ChangeMask(FBaseRemise, DEV.Decimale, DEV.Symbole);
  GP_TYPEREMISE.Visible := FOGereDemarque(False);
  ChangeMask(GP_REMISEPOURPIED, V_PGI.OkDecP, '');
  ChangeMask(GP_REMISETTCDEV, DEV.Decimale, '');
  ChangeMask(FTotAPayer, DEV.Decimale, '');
  FTotRegle.Caption := FEncaisse.Text;
  FTotAPayer.Value := FAPayer.Value;
  CalculTotalRemise;
  if not TOBPiec.FieldExists('GP_REMISEPOURPIED') then TOBPiec.AddChampSupValeur('GP_REMISEPOURPIED', 0);
  GP_REMISEPOURPIED.Value := TOBPiec.GetValue('GP_REMISEPOURPIED');
  if not TOBPiec.FieldExists('GP_REMISETTCDEV') then TOBPiec.AddChampSupValeur('GP_REMISETTCDEV', 0);
  GP_REMISETTCDEV.Value := TOBPiec.GetValue('GP_REMISETTCDEV');
  SauveMontantRemiseCur;
  if not RemiseTicket.Visible then
  begin
    if FOExisteClavierEcran then RemiseTicket.Top := 70;
    RemiseTicket.Visible := True;
  end;
  if CodeDem <> '' then GP_TYPEREMISE.Value := CodeDem;
  if Pourcent <> 0 then
  begin
    GP_REMISEPOURPIED.Value := Pourcent;
    GP_REMISEPOURPIEDExit(nil);
  end;
  if GP_REMISEPOURPIED.CanFocus then GP_REMISEPOURPIED.SetFocus else
    if GP_REMISETTCDEV.CanFocus then GP_REMISETTCDEV.SetFocus else
    if FTotAPayer.CanFocus then FTotAPayer.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculTotalRemise : calcul du montant restant à encaisser après application de la remise
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.CalculTotalRemise;
begin
  FTotARendre.Caption := StrFMontant((FTotAPayer.Value - FEncaisse.Value), 12, DEV.Decimale, DEV.Symbole, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifTotalRemise : verifie le montant de la remise
///////////////////////////////////////////////////////////////////////////////////////

function TFEcheanceFO.VerifTotalRemise: Boolean;
var Ind, NoErr: Integer;
  TOBL: TOB;
  MaxMontant: Double;
  MaxRem: Integer;
  CodeDemValide: string;
  {$IFDEF FOS5}
  MsgErr: string;
  {$ENDIF}
begin
  Result := True;
  // Pourcentage maximum de remise autorisé
  CodeDemValide := '';
  NoErr := FOVerifMaxRemise(GP_REMISEPOURPIED.Value, GP_TYPEREMISE.Value, False, MaxRem, CodeDemValide);
  if NoErr <> 0 then
  begin
    Result := False;
    Exit;
  end else
    // Remise inférieure à la somme des lignes remisable
    MaxMontant := 0;
  for Ind := 0 to TOBPiec.Detail.Count - 1 do
  begin
    TOBL := TOBPiec.Detail[Ind];
    if (TOBL <> nil) and (TOBL.GetValue('GL_ARTICLE') <> '') and (FOLigneRemisable(TOBL, True)) then
    begin
      MaxMontant := MaxMontant + TOBL.GetValue('GL_MONTANTTTCDEV');
    end;
  end;
  if FOCompareMontant(GP_REMISETTCDEV.Value, MaxMontant, '>') then
  begin
    Result := False;
    MsgBox.Execute(12, StrFMontant(GP_REMISETTCDEV.Value, 12, DEV.Decimale, DEV.Symbole, True),
      StrFMontant(MaxMontant, 12, DEV.Decimale, DEV.Symbole, True));
  end;

  {$IFDEF FOS5}
  if Not TFTicketFO(FormTicket).FideliteCli.AutoriseRemiseGlobale(GP_REMISEPOURPIED.Value, GP_REMISETTCDEV.Value, MsgErr, GP_TYPEREMISE.Value) then
  begin
    Result := False;
    PGIError(MsgErr, Caption);
  end;
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauveMontantRemiseCur : conserve la valeur des champs
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.SauveMontantRemiseCur;
begin
  PourRemCur := GP_REMISEPOURPIED.Text;
  MntRemCur := GP_REMISETTCDEV.Text;
  TotRemCur := FTotAPayer.Text;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AllerSurChpRemise : se positionne sur un champ de saisie de la remise ticket
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.AllerSurChpRemise(NoChamp: TChoixCtrlRemTic);
begin
  if RemTicketCtrl = nil then Exit;
  // On quitte le champ courant
  if RemTicketCtrl is THNumEdit then
  begin
    if Assigned(THNumEdit(RemTicketCtrl).OnExit) then THNumEdit(RemTicketCtrl).OnExit(nil);
  end else
    if RemTicketCtrl is THValComboBox then
  begin
    if Assigned(THValComboBox(RemTicketCtrl).OnExit) then THValComboBox(RemTicketCtrl).OnExit(nil);
  end;
  // On entre dans le nouveau champ
  case NoChamp of
    rtPourcent: // aller sur le % de remise
      begin
        if Assigned(GP_REMISEPOURPIED.OnEnter) then GP_REMISEPOURPIED.OnEnter(GP_REMISEPOURPIED);
        if GP_REMISEPOURPIED.CanFocus then GP_REMISEPOURPIED.SetFocus;
      end;
    rtMontant: // aller sur le montant de remise
      begin
        if Assigned(GP_REMISETTCDEV.OnEnter) then GP_REMISETTCDEV.OnEnter(GP_REMISETTCDEV);
        if GP_REMISETTCDEV.CanFocus then GP_REMISETTCDEV.SetFocus;
      end;
    rtTotal: // aller sur le total après remise
      begin
        if Assigned(FTotAPayer.OnEnter) then FTotAPayer.OnEnter(FTotAPayer);
        if FTotAPayer.CanFocus then FTotAPayer.SetFocus;
      end;
    rtMotif: // aller sur le motif de remise
      begin
        if Assigned(GP_TYPEREMISE.OnEnter) then GP_TYPEREMISE.OnEnter(GP_TYPEREMISE);
        if GP_TYPEREMISE.CanFocus then GP_TYPEREMISE.SetFocus;
      end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ArrondiRemisePied : Applique l'arrondi sur le montant de la remise de pied
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.ArrondiRemisePied;
var MtRem: Double;
begin
  MtRem := GP_REMISETTCDEV.Value;
  if MtRem <> 0 then MtRem := FOArrondiRemise(MtRem, FBaseRemise.Value);
  GP_REMISETTCDEV.Value := MtRem;
  GP_REMISETTCDEV.UpdateValue;
  GP_REMISEPOURPIED.Value := (GP_REMISETTCDEV.Value * 100) / FBaseRemise.Value;
  FTotAPayer.Value := FTotalBrut.Value - GP_REMISETTCDEV.Value;
  CalculTotalRemise;
  SauveMontantRemiseCur;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GP_REMISEPOURPIEDExit
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GP_REMISEPOURPIEDExit(Sender: TObject);
begin
  if PourRemCur = GP_REMISEPOURPIED.Text then Exit;
  GP_REMISETTCDEV.Value := FBaseRemise.Value * GP_REMISEPOURPIED.Value / 100;
  ArrondiRemisePied;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GP_REMISETTCDEVExit
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.GP_REMISETTCDEVExit(Sender: TObject);
begin
  if MntRemCur = GP_REMISETTCDEV.Text then Exit;
  ArrondiRemisePied;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FTotAPAyerExit
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.FTotAPAyerExit(Sender: TObject);
begin
  if TotRemCur = FTotAPayer.Text then Exit;
  GP_REMISETTCDEV.Value := FBaseRemise.Value - FTotAPayer.Value;
  ArrondiRemisePied;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCancelRemiseClick : abandon de la remise ticket saisie
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BCancelRemiseClick(Sender: TObject);
begin
  RemiseTicket.Visible := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValiderRemiseClick : validation de la remise ticket saisie
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.BValiderRemiseClick(Sender: TObject);
var Ok, Cancel: Boolean;
  ACol, ARow: Integer;
begin
  if (Screen.ActiveControl is THNumEdit) and (Assigned(THNumEdit(Screen.ActiveControl).OnExit)) then
  begin
    THNumEdit(Screen.ActiveControl).UpdateValue;
    THNumEdit(Screen.ActiveControl).OnExit(Sender);
  end;
  if GP_REMISETTCDEV.Value = 0 then
  begin
    // aucune remise saisie
    RemiseTicket.Visible := False;
    if GSReg.CanFocus then GSReg.SetFocus;
    AfficheLigneLCD(GSReg.Row);
    Exit;
  end;
  // Contrôle de cohérence
  if not VerifTotalRemise then
  begin
    if GP_REMISETTCDEV.CanFocus then GP_REMISETTCDEV.SetFocus else
      if RemTicketCtrl.CanFocus then RemTicketCtrl.SetFocus;
    Exit;
  end;
  // Contrôle du code démarque
  if (FOGereDemarque(True)) and (GP_REMISETTCDEV.Value <> 0) and (GP_TYPEREMISE.Value = '') then
  begin
    MsgBox.Execute(11, '', '');
    if GP_TYPEREMISE.CanFocus then GP_TYPEREMISE.SetFocus;
    Exit;
  end;
  // Contrôle du client obligatoire
  if (FOGetParamCaisse('GPK_CLISAISIE') = 'X') and (FODemarqueClientOblig(GP_TYPEREMISE.Value)) then
  begin
    {$IFDEF FOS5}
    if (TOBPiec.GetValue('GP_TIERS') = FOGetParamCaisse('GPK_TIERS')) and
      (MsgBox.Execute(26, Caption, '') = mrYes) then
    begin
      if (FormTicket <> nil) and (FormTicket is TFTicketFO) then
        TFTicketFO(FormTicket).MCChoixClientClick(nil);
    end;
    {$ENDIF}
    if TOBPiec.GetValue('GP_TIERS') = FOGetParamCaisse('GPK_TIERS') then
    begin
      if GP_TYPEREMISE.CanFocus then GP_TYPEREMISE.SetFocus;
      Exit;
    end;
  end;
  // Vérifie si le montant de la ligne doit être recalculé.
  ARow := GSReg.Row;
  //Ok := ((GSReg.Col = SG_Mont) and (Valeur(GSReg.Cells[SG_Mont, GSReg.Row]) = CalculeMontantLigne(ARow, FReste.Value))) ;
  Ok := ((GSReg.Col = SG_Mont) and (LigneVautReste(ARow)));
  // Mise à jour de la piece
  if GP_REMISETTCDEV.Value = 0 then GP_TYPEREMISE.Value := '';
  TOBPiec.PutValue('GP_TYPEREMISE', GP_TYPEREMISE.Value);
  TOBPiec.PutValue('GP_REMISEPOURPIED', GP_REMISEPOURPIED.Value);
  TOBPiec.PutValue('GP_REMISETTCDEV', GP_REMISETTCDEV.Value);
  FAPayer.Value := FTotAPayer.Value;
  RemiseTicket.Visible := False;
  AfficheTotal;
  // Mise à jour du montant de la ligne
  if Ok then
  begin
    ACol := SG_Mont;
    ARow := GSReg.Row;
    Cancel := False;
    GSReg.Cells[SG_Mont, ARow] := '';
    PrepareMontant(ACol, ARow, Cancel);
    if GSReg.CanFocus then GSReg.SetFocus;
  end;
  AfficheLigneLCD(GSReg.Row);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RemiseTicketEnter : OnEnter sur les contôles de la saisie de la remise ticket
///////////////////////////////////////////////////////////////////////////////////////

procedure TFEcheanceFO.RemiseTicketEnter(Sender: TObject);
begin
  // On conserve le contrôle actif
  if Sender is TWinControl then RemTicketCtrl := TWinControl(Sender) else RemTicketCtrl := nil;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MiseAJourPiece : mise à jour et recalcul de la pièce
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF FOS5}

procedure TFEcheanceFO.MiseAJourPiece;
begin
  if (FormTicket = nil) or not (FormTicket is TFTicketFO) then Exit;
  // Répartition du montant de la remise sur les lignes
  if TOBPiec.FieldExists('GP_REMISETTCDEV') then
    FOAppliqueRemiseGlobale(TFTicketFO(FormTicket).TOBPiece, TOBPiec.GetValue('GP_REMISETTCDEV'), TOBPiec.GetValue('GP_TYPEREMISE'), DEV);
  // Recalcul de la pièce
  CalculFacture(TFTicketFO(FormTicket).PieceContainer);
end;
{$ENDIF}

{==============================================================================================}
{========== IMPRESSION DES CHEQUES, DES BONS ET VALIDATION DU TPE =============================}
{==============================================================================================}
///////////////////////////////////////////////////////////////////////////////////////
//  ImpressionCheque : impression des chèques et des bons
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF FOS5}

procedure TFEcheanceFO.ImpressionCheque;
begin
  if Action = taConsult then Exit;
  FOImpressionCheque(TOBEche, TOBMode);
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  ValideCB : validation des cartes bancaires et des chèques par le TPE
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF FOS5}

function TFEcheanceFO.ValideCB: Boolean;
var
  TOBL, TOBM, TOBD: TOB;
  Ind, Err: integer;
  Montant: double;
  CodeMode, TypeMode, sLib: string;
  OkTpe, SuiteCB: boolean;
begin
  Result := True;
  if Action = taConsult then Exit;
  OkTpe := ((Action = taCreat) and (DepuisPiece) and (FOExisteTPE));
  for Ind := 1 to TOBEche.Detail.Count do
  begin
    TOBL := TOBEche.Detail[(Ind - 1)];
    if (TOBL <> nil) and (not FOTesteChampSupValeur(TOBL, 'ACCORDTPE', 'X')) then
    begin
      CodeMode := TOBL.GetValue('GPE_MODEPAIE');
      Montant := TOBL.GetValue('GPE_MONTANTENCAIS');
      TOBM := TOBMode.FindFirst(['MP_MODEPAIE'], [CodeMode], False);
      if TOBM <> nil then
      begin
        SuiteCB := True;
        TypeMode := TOBM.GetValue('MP_TYPEMODEPAIE');
        // envoi du montant au TPE
        if (OkTpe) and (TOBM.GetValue('MP_ENVOITPEFO') = 'X') and (Montant <> 0) then
        begin
          if (TypeMode = TYPEPAIECB) or ((TypeMode = TYPEPAIECHEQUE) or (TypeMode = TYPEPAIECHQDIFF) and (Montant > 0)) then
          begin
            TOBD := TOBDev.FindFirst(['D_DEVISE'], [TOBM.GetValue('MP_DEVISEFO')], False);
            if TypeMode = TYPEPAIECB then FODecodeLectureCBFromKB(TOBEche, TOBL, TOBM, True, '');

            Err := FOLanceAcceptationCB(TOBL, TOBD, TOBM);
            case Err of
              mrOk :
              begin
                // transaction acceptée par le TPE
                FOMajChampSupValeur(TOBL, 'ACCORDTPE', 'X');
                SuiteCB := (FOVerifDetailCB(TOBL, TOBM) <> 0);
              end;
              mrYes :
              begin
                // transaction forcée par l'utilisateur
                SuiteCB := True;
              end;
            else
              begin
                // transaction refusée
                Result := False;
                GSReg.Row := Ind;                                                                                                                    
                if SG_Mode <> -1 then GSReg.Col := SG_Mode;
                Exit;
              end;
            end;
          end;
        end;

        if SuiteCB then
        begin
          // Saisie du détail de la carte bancaire
          if (FormTicket <> nil) and (FormTicket is TFTicketFO) then
            sLib := TFTicketFO(FormTicket).LIBELLETIERS.Caption
          else
            sLib := FOMakeNomClient(nil, TOBL.GetValue('GPE_TIERS'));

          Result := FOSaisieDetailCB(TOBEche, TOBL, TOBM, sLib);
        end;
      end;
    end;
  end;
end;
{$ENDIF}

{==============================================================================================}
{===================================== DETAXE =================================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 15/04/2003
Modifié le ... : 15/04/2003
Description .. : Détaxe demandée par l'utilisateur
Mots clefs ... :
*****************************************************************}

procedure TFEcheanceFO.TraiteDetaxe;
begin
  if Action <> taCreat then Exit;
  {$IFDEF FOS5}
  if (FormTicket <> nil) and (FormTicket is TFTicketFO) then
  begin
    if not TFTicketFO(FormTicket).DetaxeActive then Exit;
    TFTicketFO(FormTicket).TraiteDemandeDetaxe;
    LDETAXE.Visible := TFTicketFO(FormTicket).DemandeDetaxe;
  end;
  {$ENDIF}
end;

{==============================================================================================}
{=============================== LIAISON DES REGLEMENTS =======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 15/04/2003
Modifié le ... : 15/04/2003
Description .. : Recherche des règlements disponibles
Mots clefs ... :
*****************************************************************}

function TFEcheanceFO.RechercheReglementDispo(ARow: integer; MntReste: double; var Montant: double): boolean;
var
  TypeMode: string;
  TOBL: TOB;
begin
  Result := False;
  Montant := 0;
  TypeMode := '';
  TOBL := GetTobEcheance(ARow);
  if TOBL = nil then Exit;
  if TOBL.FieldExists('TYPEMODEPAIE') then
    TypeMode := TOBL.GetValue('TYPEMODEPAIE');
  if TypeMode = '' then
    TypeMode := GetTypeModePaie(TOBL.GetValue('GPE_MODEPAIE'));

  Result := FORechercheOpCaisseDispo(TOBL, TOBPiec, TOBEche, TypeMode, MntReste, Montant);
  if Result then TOBL.PutValue('GPE_MONTANTENCAIS', Montant);
end;

end.
