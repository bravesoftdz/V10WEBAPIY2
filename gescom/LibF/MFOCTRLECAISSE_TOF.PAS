{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/01/2002
Modifié le ... : 28/01/2002
Description .. : Source TOF de la FICHE : MFOCTRLECAISSE()
Mots clefs ... : TOF;MFOCTRLECAISSE
*****************************************************************}
unit MFOCTRLECAISSE_TOF;

interface

uses
  Forms, Classes, StdCtrls, ComCtrls, Controls, Graphics, Grids, HCtrls, HPanel,
  Windows, sysutils, extctrls, imglist, Vierge, HTB97, HMsgBox, HEnt1, HDebug,
  {$IFDEF EAGLCLIENT}
  MaineAgl,
  {$ELSE}
  FE_Main, dbtables,
  {$ENDIF}
  UTOF, UTOB, Kb_Ecran;

type
  TOF_MFOCTRLECAISSE = class(TOF)
  private
    FCalculette: TCalculatriceEcran; // calculette tactile
    ImagesBtn: TImageList; // liste d'image
    Caisse: string; // code de la caisse
    NumZ: Integer; // n° de clôture
    Vendeur: string; // code du vendeur qui fait la fermeture
    TOBEch: TOB; // TOB des échéances (cumuls + détail)
    TOBBillet: TOB; // TOB des billets (montants par type de billets)
    TOBPiece: TOB; // TOB des pièces (montants par type de pièces)
    TOBBilMdp: TOB; // TOB des billets du mode de paiement courant
    TOBPieMdp: TOB; // TOB des pièces du mode de paiement courant
    GCum: THGrid; // grille des cumuls
    GDet: THGrid; // grille des détails
    GBillet: THGrid; // grille des billets
    GPiece: THGrid; // grille des pièces
    PDetail: THPanel; // Panel du détail
    PPiecBil: THPanel; // Panel des pièces et billers
    BoutonEncours: Boolean; // traitement d'un bouton en cours
    ApplicEncours: Boolean; // recopie en cours du total (détail ou pièces/billet) sur la grille des cumuls
    ValideEnCours: Boolean; // la validation de la saisie est en cours
    PiecBilEnCours: Boolean; // saisie des pièces et billets en cours
    AffichageDetail: boolean; // affichage du détail des tickets autorisé
    CurLigDet: Integer; // ligne de détail en cours
    StCellCur: string; // valeur de la cellule courante de la grille des cumuls
    StPiecBilCur: string; // valeur de la cellule courante de la grille des pièces ou billets
    COLC_IMAGE: Integer; // définition des colonnes de la grille des cumuls
    COLC_MODEPAIE: Integer;
    COLC_DEVISE: Integer;
    COLC_TOTCALCULE: Integer;
    COLC_ETAT: Integer;
    COLC_TOTSAISI: Integer;
    COLC_TOTFDCAIS: Integer;
    COLC_TOTENCAIS: Integer;
    COLD_IMAGE: Integer; // définition des colonnes de la grille des détails
    COLD_NUMERO: Integer;
    COLD_TIERS: Integer;
    COLD_MONTANT: Integer;
    COLD_DEVISE: Integer;
    COLB_LIBELLE: Integer; // définition des colonnes des grilles des billets et pièces
    COLB_QUANTITE: Integer;
    COLB_MONTANT: Integer;
    procedure CreeListeImage;
    procedure InitColNumber;
    procedure GCumMasqueCol;
    procedure GCumColAligns;
    procedure GCumPutGrid(TOBE: TOB; ARow: Integer);
    procedure GPiecBilColAligns(GS: THGrid);
    procedure GPiecBilPutGrid(GS: THGrid; TOBPB: TOB);
    procedure GDetMasqueCol;
    procedure GDetColAligns;
    procedure GDetPutGrid(TOBD: TOB);
    procedure AlloueTOB;
    function GetTOBLigne(Row: Integer): TOB;
    procedure AfficheDetail(Row: Integer; ForceAffich, ForceDetail: Boolean);
    function ZoneAccessible(ACol, ARow: Longint): Boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: Boolean);
    procedure TraiteTotSaisi(var ACol, ARow: Integer; var Cancel: Boolean);
    procedure TraiteTotFdCais(var ACol, ARow: Integer; var Cancel: Boolean);
    procedure TraiteTotEncais(var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCumEnter(Sender: TObject);
    procedure GCumExit(Sender: TObject);
    procedure GCumCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCumCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCumRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GCumRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GCumDessineCell(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
    procedure GCumPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GDetRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GDetRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GDetEnter(Sender: TObject);
    procedure GDetExit(Sender: TObject);
    procedure GDetFlipSelection(Sender: TObject);
    procedure GDetPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GBilletEnter(Sender: TObject);
    procedure GBilletExit(Sender: TObject);
    procedure GBilletCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GBilletCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GBilletColumnWidthsChanged(Sender: TObject);
    procedure GPieceEnter(Sender: TObject);
    procedure GPieceExit(Sender: TObject);
    procedure GPieceCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GPieceCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GPieceColumnWidthsChanged(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure CalculMontantPieceBillet(GS: THGrid; TOBParent: TOB; ACol, ARow: Longint; var Cancel: Boolean);
    function CalculTotalPieceBillet(TOBBil, TOBPie: TOB; Affiche: Boolean): Double;
    function VerifiePieceBillet(ModePaie: string; TotSaisi: Double): Boolean;
    procedure FormateZoneSaisie(GS: THGrid; ACol, ARow: Longint);
    procedure GotoEntete;
    procedure GotoPied;
    procedure BoutonCalculetteClick(Val: string);
    procedure SaisieCalculette(Val: Double);
    procedure AppliqueTotal;
    procedure BLeftClick(Sender: TObject);
    procedure BRightClick(Sender: TObject);
    procedure BUpClick(Sender: TObject);
    procedure BDownClick(Sender: TObject);
    procedure BEnteteClick(Sender: TObject);
    procedure BDetailClick(Sender: TObject);
    procedure BBilletClick(Sender: TObject);
    procedure BVoirTicClick(Sender: TObject);
    procedure BSelectClick(Sender: TObject);
    procedure BSuivantClick(Sender: TObject);
    procedure BPrecedentClick(Sender: TObject);
  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  end;

procedure FOChargeTOBPiecBil(TOBEch, TOBBillet, TOBPiece: TOB; ModePaie, Devise, Prefixe: string);
function FOMAJTotalPieceBillet(TOBBillet, TOBPiece: TOB; ModePaie, Prefixe: string): string;
procedure FOChargeTablesCtrlCaisse(TOBPrinc: TOB; Caisse: string; NumZ: Integer);
procedure FOMAJTablesCtrlCais(TOBPrinc: TOB; Caisse, Vendeur, Titre: string; NumZ: Integer; PieceEcart: boolean);
function FOSelectCtrlCaisse: string;

implementation
uses
  EntGC, SaisUtil, FactUtil, FactTOB, TickUtilFO, FODefi, FOUtil;

///////////////////////////////////////////////////////////////////////////////////////
//  OnNew :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.OnNew;
begin
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnDelete :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.OnDelete;
begin
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnUpdate :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.OnUpdate;
var Ind: Integer;
  TOBL: TOB;
  sEtat: string;
begin
  inherited;
  if TOBEch = nil then Exit;
  if ValideEnCours then Exit;
  ValideEnCours := True;
  BoutonEncours := True;
  // tous les cumuls doivent être saisis
  for Ind := 0 to TOBEch.Detail.Count - 1 do
  begin
    TOBL := TOBEch.Detail[Ind];
    if TOBL.FieldExists('CODEETAT') then sEtat := TOBL.GetValue('CODEETAT') else sEtat := '';
    if sEtat = 'E' then
    begin
      PGIBox('Il reste une ligne en erreur !', Ecran.Caption);
      LastError := -1; // la fiche ne sera pas fermée
      Break;
    end;
    if (sEtat <> 'O') and (sEtat <> 'F') and (TOBL.GetValue('TOTCALCULE') <> 0) then
    begin
      PGIBox('Vous devez saisir toutes les lignes !', Ecran.Caption);
      LastError := -1; // la fiche ne sera pas fermée
      Break;
    end;
    if not VerifiePieceBillet(TOBL.GetValue('MP_MODEPAIE'), TOBL.GetValue('TOTSAISI')) then
    begin
      PGIBox('Le détail des pièces et de billets est incorrect !', Ecran.Caption);
      LastError := -1; // la fiche ne sera pas fermée
      Break;
    end;
  end;
  if LastError = 0 then TFVierge(Ecran).Retour := 'OK' else TFVierge(Ecran).Retour := '';
  ValideEnCours := False;
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnLoad :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.OnLoad;
begin
  inherited;
  if GCum <> nil then
  begin
    if GCum.CanFocus then GCum.SetFocus;
    if Assigned(GCum.OnEnter) then GCum.OnEnter(GCum);
    GCum.Row := GCum.FixedRows;
    GCum.Col := COLC_TOTSAISI;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnArgument :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.OnArgument(S: string);
///////////////////////////////////////////////////////////////////////////////////////
  procedure BrancheClickButton(BtName: string; ClickFct: TNotifyEvent);
  var Ctrl: TControl;
  begin
    // Branchement des événements des boutons
    Ctrl := GetControl(BtName);
    if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then TToolbarButton97(Ctrl).OnClick := ClickFct;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
var Ctrl: TControl;
  Ind, iWidth: Integer;
begin
  inherited;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then
  begin
    // GESCOM 
    Ecran.Width := Screen.Width;
    Ecran.Height := Screen.Height;
  end else
  if not V_PGI.Maximize then
  begin
    Ecran.WindowState := wsNormal;
    Ecran.Width := Application.MainForm.Width;
    Ecran.Height := Application.MainForm.Height;
  end;
  BoutonEncours := False;
  ApplicEncours := False;
  ValideEnCours := False;
  PiecBilEnCours := False;
  AffichageDetail := True;
  CurLigDet := -1;
  InitColNumber;
  AlloueTOB;
  // Caisse et numéro de clôture à traiter
  Caisse := FOCaisseCourante;
  NumZ := FOGetNumZCaisse(Caisse);
  // Vendeur qui ferme la journée
  Ind := Pos('VENDEUR=', S);
  if Ind > 0 then Vendeur := Copy(S, (Ind + 1), (Length(S) - Ind))
  else Vendeur := '';
  // Branchement des événements des boutons
  BrancheClickButton('BLEFT', BLeftClick);
  BrancheClickButton('BRIGHT', BRightClick);
  BrancheClickButton('BUP', BUpClick);
  BrancheClickButton('BDOWN', BDownClick);
  BrancheClickButton('BENTETE', BEnteteClick);
  BrancheClickButton('BDETAIL', BDetailClick);
  BrancheClickButton('BBILLET', BBilletClick);
  BrancheClickButton('BVOIRTIC', BVoirTicClick);
  BrancheClickButton('BSELECT', BSelectClick);
  BrancheClickButton('BSUIVANT', BSuivantClick);
  BrancheClickButton('BPRECEDENT', BPrecedentClick);
  SetControlVisible('BFERME', False);
  Ecran.OnKeyDown := FormKeyDown;
  // Affichage de la calculette
  Ctrl := GetControl('PCALC');
  if (Ctrl <> nil) and (Ctrl is THPanel) then
  begin
    FCalculette := TCalculatriceEcran.Create(Ecran);
    FCalculette.Parent := THPanel(Ctrl);
    FCalculette.Name := 'CALCULETTE';
    FCalculette.OnActionClc := BoutonCalculetteClick;
    FCalculette.OnSaisieClc := SaisieCalculette;
  end;
  // Alignement des panels détail et billets
  Ctrl := GetControl('PDETAIL');
  if (Ctrl <> nil) and (Ctrl is THPanel) then PDetail := THPanel(Ctrl) else PDetail := nil;
  Ctrl := GetControl('PPIECBIL');
  if (Ctrl <> nil) and (Ctrl is THPanel) then PPiecBil := THPanel(Ctrl) else PPiecBil := nil;
  if (PPiecBil <> nil) and (PDetail <> nil) then
  begin
    PPiecBil.Top := PDetail.Top;
    PPiecBil.Left := PDetail.Left;
    PPiecBil.Height := PDetail.Height;
    PPiecBil.Width := PDetail.Width;
    PPiecBil.Visible := False;
    PDetail.Visible := True;
  end;
  // Alignement du panel des pièces et de celui des billets
  if PPiecBil <> nil then
  begin
    iWidth := PPiecBil.Width div 2;
    Ctrl := GetControl('PBILLET');
    if (Ctrl <> nil) and (Ctrl is THPanel) then THPanel(Ctrl).Width := iWidth;
    Ctrl := GetControl('PPIECE');
    if (Ctrl <> nil) and (Ctrl is THPanel) then THPanel(Ctrl).Width := PPiecBil.Width - iWidth;
  end;
  // Grille des cumuls
  Ctrl := GetControl('GCUM');
  if (Ctrl <> nil) and (Ctrl is THGrid) then
  begin
    GCum := THGrid(Ctrl);
    // branchement des événements
    GCum.OnEnter := GCumEnter;
    GCum.OnExit := GCumExit;
    GCum.OnCellEnter := GCumCellEnter;
    GCum.OnCellExit := GCumCellExit;
    GCum.OnRowEnter := GCumRowEnter;
    GCum.OnRowExit := GCumRowExit;
    GCum.GetCellCanvas := GCumDessineCell;
    GCum.PostDrawCell := GCumPostDrawCell;
    // colonnes inaccessibles
    GCum.ColLengths[COLC_IMAGE] := -1;
    GCum.ColLengths[COLC_MODEPAIE] := -1;
    GCum.ColLengths[COLC_DEVISE] := -1;
    GCum.ColLengths[COLC_TOTCALCULE] := -1;
    GCum.ColLengths[COLC_ETAT] := -1;
    if FOGetParamCaisse('GPK_MODIFFDCAIS') <> 'X' then
    begin
      GCum.ColLengths[COLC_TOTFDCAIS] := -1;
      GCum.ColLengths[COLC_TOTENCAIS] := -1;
    end;
    // colonnes masquées
    GCumMasqueCol;
    // format de saisie
    GCum.ColFormats[COLC_TOTSAISI] := StrfMask(V_PGI.OkDecV, '', True);
    // alignement des données
    GCumColAligns;
    GCum.Col := COLC_TOTSAISI;
  end else GDet := nil;
  // Grille des details
  Ctrl := GetControl('GDET');
  if (Ctrl <> nil) and (Ctrl is THGrid) then
  begin
    GDet := THGrid(Ctrl);
    // branchement des événements
    GDet.OnRowEnter := GDetRowEnter;
    GDet.OnRowExit := GDetRowExit;
    GDet.OnEnter := GDetEnter;
    GDet.OnExit := GDetExit;
    GDet.OnFlipSelection := GDetFlipSelection;
    GDet.PostDrawCell := GDetPostDrawCell;
    // colonnes masquées
    GDetMasqueCol;
    // alignement des données
    GDetColAligns;
  end else GDet := nil;
  // Grille des billets
  Ctrl := GetControl('GBILLET');
  if (Ctrl <> nil) and (Ctrl is THGrid) then
  begin
    GBillet := THGrid(Ctrl);
    // branchement des événements
    GBillet.OnEnter := GBilletEnter;
    GBillet.OnExit := GBilletExit;
    GBillet.OnCellEnter := GBilletCellEnter;
    GBillet.OnCellExit := GBilletCellExit;
    GBillet.OnColumnWidthsChanged := GBilletColumnWidthsChanged;
    // colonnes inaccessibles
    GBillet.ColLengths[COLB_LIBELLE] := -1;
    GBillet.ColLengths[COLB_MONTANT] := -1;
    // format de saisie
    GBillet.ColFormats[COLB_QUANTITE] := '###0';
    // alignement des données
    GPiecBilColAligns(GBillet);
  end else GBillet := nil;
  // Grille des pièces
  Ctrl := GetControl('GPIECE');
  if (Ctrl <> nil) and (Ctrl is THGrid) then
  begin
    GPiece := THGrid(Ctrl);
    // branchement des événements
    GPiece.OnEnter := GPieceEnter;
    GPiece.OnExit := GPieceExit;
    GPiece.OnCellEnter := GPieceCellEnter;
    GPiece.OnCellExit := GPieceCellExit;
    GPiece.OnColumnWidthsChanged := GPieceColumnWidthsChanged;
    // colonnes inaccessibles
    GPiece.ColLengths[COLB_LIBELLE] := -1;
    GPiece.ColLengths[COLB_MONTANT] := -1;
    // format de saisie
    GPiece.ColFormats[COLB_QUANTITE] := '###0';
    // alignement des données
    GPiecBilColAligns(GPiece);
  end else GPiece := nil;
  // Création de la liste d'image
  CreeListeImage;
  // Affichage du détail des tickets interdit
  if not FOJaiLeDroit(84, False, False) then
  begin
    AffichageDetail := False;
    PDetail.Visible := False;
    SetControlVisible('BDETAIL', False);
    SetControlVisible('BVOIRTIC', False);
    SetControlVisible('BSELECT', False);
  end;
  // Affichage des cumuls et du détail du 1er cumul
  if GCum <> nil then
  begin
    GCumPutGrid(nil, -1);
    AfficheDetail(GCum.FixedRows, False, False);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnClose :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.OnClose;
begin
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CreeListeImage : met en liste les images du fichier des ressources
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.CreeListeImage;
///////////////////////////////////////////////////////////////////////////////////////
  procedure AddImageToList(Lst: TImageList; CtrlName: string);
  var Ctrl: TControl;
  begin
    Ctrl := GetControl(CtrlName);
    if (Ctrl <> nil) and (Ctrl is TImage) then
    begin
      Lst.AddMasked(TImage(Ctrl).Picture.Bitmap, clFuchsia);
    end;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
begin
  ImagesBtn := TImageList.Create(Ecran);
  ImagesBtn.AllocBy := 8;
  AddImageToList(ImagesBtn, 'IMESP'); // 0 = espèce
  AddImageToList(ImagesBtn, 'IMCHQ'); // 1 = chèque
  AddImageToList(ImagesBtn, 'IMCB'); // 2 = carte bancaire
  AddImageToList(ImagesBtn, 'IMBAC'); // 3 = bon d'achat
  AddImageToList(ImagesBtn, 'IMARR'); // 4 = arrhes
  AddImageToList(ImagesBtn, 'IMAV'); // 5 = avoir
  AddImageToList(ImagesBtn, 'IMRD'); // 6 = reste dû
  AddImageToList(ImagesBtn, 'IMAUT'); // 7 = autres
  //ImagesBtn.Masked := True ;
  //ImagesBtn.BkColor := clNone ;
  //ImagesBtn.ResourceLoad(rtBitmap, 'BES', clFuchsia) ;   // 0 = espèce
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumMasqueCol : initialise les numéros des colonnes des grilles
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.InitColNumber;
begin
  // définition des colonnes de la grille des cumuls
  COLC_IMAGE := 0;
  COLC_MODEPAIE := 1;
  COLC_DEVISE := 2;
  COLC_TOTCALCULE := 3;
  COLC_ETAT := 4;
  COLC_TOTSAISI := 5;
  COLC_TOTFDCAIS := 6;
  COLC_TOTENCAIS := 7;
  // définition des colonnes de la grille des détails
  COLD_IMAGE := 0;
  COLD_NUMERO := 1;
  COLD_TIERS := 2;
  COLD_MONTANT := 3;
  COLD_DEVISE := 4;
  // définition des colonnes des grilles des billets et pièces
  COLB_LIBELLE := 0;
  COLB_QUANTITE := 1;
  COLB_MONTANT := 2;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumMasqueCol : masque les colonnes de la grille des cumuls
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumMasqueCol;
var Ind: Integer;
  OkTitre: Boolean;
begin
  if GCum = nil then Exit;
  OkTitre := False;
  if FOGetParamCaisse('GPK_GEREFONDCAISSE') <> 'X' then
  begin
    GCum.Titres.Delete(COLC_TOTENCAIS);
    GCum.Titres.Delete(COLC_TOTFDCAIS);
    OkTitre := True;
    Ind := GCum.ColCount;
    Dec(Ind, 2);
    GCum.ColCount := Ind;
    COLC_TOTFDCAIS := -1;
    COLC_TOTENCAIS := -1;
  end;
  // décalage des caractèristiques des colonnes pour remplacer la colonne total calculé
  if FOGetParamCaisse('GPK_CTRLAVEUGLE') = 'X' then
  begin
    for Ind := COLC_TOTCALCULE to GCum.ColCount - 2 do
    begin
      GCum.ColLengths[Ind] := GCum.ColLengths[Ind + 1];
      GCum.ColWidths[Ind] := GCum.ColWidths[Ind + 1];
    end;
    GCum.Titres.Delete(COLC_TOTCALCULE);
    OkTitre := True;
    Ind := GCum.ColCount;
    Dec(Ind);
    GCum.ColCount := Ind;
    COLC_TOTCALCULE := -1;
    Dec(COLC_ETAT);
    Dec(COLC_TOTSAISI);
    Dec(COLC_TOTFDCAIS);
    Dec(COLC_TOTENCAIS);
  end;
  if OkTitre then GCum.UpdateTitres;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumColAligns : définition de l'alignement des données de la grille des cumuls
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumColAligns;
begin
  if COLC_DEVISE >= 0 then GCum.ColAligns[COLC_DEVISE] := taCenter;
  if COLC_ETAT >= 0 then GCum.ColAligns[COLC_ETAT] := taCenter;
  if COLC_TOTCALCULE >= 0 then GCum.ColAligns[COLC_TOTCALCULE] := taRightJustify;
  if COLC_TOTSAISI >= 0 then GCum.ColAligns[COLC_TOTSAISI] := taRightJustify;
  if COLC_TOTFDCAIS >= 0 then GCum.ColAligns[COLC_TOTFDCAIS] := taRightJustify;
  if COLC_TOTENCAIS >= 0 then GCum.ColAligns[COLC_TOTENCAIS] := taRightJustify;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumPutGrid : affichage des lignes de la TOB dans la grille des cumuls
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumPutGrid(TOBE: TOB; ARow: Integer);
var sColName: string;
begin
  if TOBEch = nil then Exit;
  if GCum = nil then Exit;
  sColName := ';MP_LIBELLE;D_SYMBOLE';
  if FOGetParamCaisse('GPK_CTRLAVEUGLE') <> 'X' then sColName := sColName + ';TOTCALCULE';
  sColName := sColName + ';ETAT;TOTSAISI';
  if FOGetParamCaisse('GPK_GEREFONDCAISSE') = 'X' then sColName := sColName + ';TOTFDCAIS;TOTENCAIS';
  if (TOBE = nil) or (ARow < GCum.FixedRows) then
  begin
    // affichage de toutes les lignes
    TOBEch.PutGridDetail(GCum, False, False, sColName, True);
    TFVierge(Ecran).HMTrad.ResizeGridColumns(GCum);
  end else
  begin
    // affichage d'une ligne
    TOBE.PutLigneGrid(GCum, ARow, False, False, sColName);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPiecBilColAligns : définition de l'alignement des données des pièces et billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GPiecBilColAligns(GS: THGrid);
begin
  if COLB_QUANTITE >= 0 then GS.ColAligns[COLB_QUANTITE] := taRightJustify;
  if COLB_MONTANT >= 0 then GS.ColAligns[COLB_MONTANT] := taRightJustify;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPiecBilPutGrid : affichage des lignes de la TOB dans la grille des pièces ou dans celle de billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GPiecBilPutGrid(GS: THGrid; TOBPB: TOB);
var sColName: string;
begin
  if TOBPB = nil then Exit;
  if GS = nil then Exit;
  sColName := 'GPI_LIBELLE;QUANTITE;MONTANT';
  TOBPB.PutGridDetail(GS, False, False, sColName, True);
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GS);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetMasqueCol : masque les colonnes de la grille du détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetMasqueCol;
var Ind: Integer;
begin
  if GDet = nil then Exit;
  if FOGetParamCaisse('GPK_CLISAISIE') = 'X' then Exit;
  if COLD_TIERS = -1 then Exit;
  // décalage des caractèristiques des colonnes pour remplacer la colonne tiers
  for Ind := COLD_TIERS to GDet.ColCount - 2 do
  begin
    GDet.ColLengths[Ind] := GDet.ColLengths[Ind + 1];
    GDet.ColWidths[Ind] := GDet.ColWidths[Ind + 1];
  end;
  GDet.Titres.Delete(COLD_TIERS);
  Ind := GDet.ColCount;
  Dec(Ind);
  GDet.ColCount := Ind;
  COLD_TIERS := -1;
  Dec(COLD_MONTANT);
  Dec(COLD_DEVISE);
  GDet.UpdateTitres;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetColAligns : définition de l'alignement des données de la grille des détails
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetColAligns;
begin
  GDet.ColAligns[COLD_DEVISE] := taCenter;
  GDet.ColAligns[COLD_MONTANT] := taRightJustify;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetPutGrid : affichage des lignes de la TOB dans la grille du détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetPutGrid(TOBD: TOB);
var sColName: string;
begin
  if TOBD = nil then Exit;
  if GDet = nil then Exit;
  if FOGetParamCaisse('GPK_CLISAISIE') = 'X' then sColName := 'xx;GPE_NUMERO;GPE_TIERS;GPE_MONTANTENCAIS;D_SYMBOLE'
  else sColName := 'xx;GPE_NUMERO;GPE_MONTANTENCAIS;D_SYMBOLE';
  TOBD.PutGridDetail(GDet, False, False, sColName, True);
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GDet);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AlloueTOB : Création des TOB
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.AlloueTOB;
var TOBPrinc: TOB;
begin
  TOBPrinc := TFVierge(Ecran).LaTOF.LaTOB;
  if TOBPrinc.Detail.Count < 1 then TOBEch := TOB.Create('Les echeances', TOBPrinc, -1)
  else TOBEch := TOBPrinc.Detail[0];
  if TOBPrinc.Detail.Count < 2 then TOBBillet := TOB.Create('Les billets', TOBPrinc, -1)
  else TOBBillet := TOBPrinc.Detail[1];
  if TOBPrinc.Detail.Count < 3 then TOBPiece := TOB.Create('Les pieces', TOBPrinc, -1)
  else TOBPiece := TOBPrinc.Detail[2];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetTOBLigne : retourne la TOB d'une ligne de cumul
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOCTRLECAISSE.GetTOBLigne(Row: Integer): TOB;
var ARow: Integer;
begin
  Result := nil;
  if TOBEch = nil then Exit;
  ARow := Row - GCum.FixedRows;
  if (ARow >= 0) and (ARow < TOBEch.Detail.Count) then Result := TOBEch.Detail[ARow];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheDetail : affiche le détail d'une ligne de cumul
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.AfficheDetail(Row: Integer; ForceAffich, ForceDetail: Boolean);
var TOBC: TOB;
  sTitre, sMdp: string;
begin
  if GDet = nil then Exit;
  if (Row = CurLigDet) and (not ForceAffich) then Exit;
  TOBBilMdp := nil;
  TOBPieMdp := nil;
  TOBC := GetTOBLigne(Row);
  if TOBC <> nil then
  begin
    CurLigDet := Row;
    sTitre := TOBC.GetValue('MP_LIBELLE');
    sMdp := TOBC.GetValue('MP_MODEPAIE');
    if TOBC.GetValue('MP_TYPEMODEPAIE') = TYPEPAIEESPECE then
    begin
      TOBBilMdp := TOBBillet.FindFirst(['MP_MODEPAIE'], [sMdp], False);
      TOBPieMdp := TOBPiece.FindFirst(['MP_MODEPAIE'], [sMdp], False);
    end;
    if (not ForceDetail) and ((TOBBilMdp <> nil) or (TOBPieMdp <> nil)) then
    begin
      PDetail.Visible := False;
      PPiecBil.Visible := True;
      SetControlProperty('PCUMUL', 'Caption', sTitre);
      SetControlProperty('PCUMUL', 'TextEffect', 1);
      if TOBBilMdp <> nil then GPiecBilPutGrid(GBillet, TOBBilMdp);
      if TOBPieMdp <> nil then GPiecBilPutGrid(GPiece, TOBPieMdp);
      CalculTotalPieceBillet(TOBBilMdp, TOBPieMdp, True);
      SetControlEnabled('BENTETE', True);
      SetControlEnabled('BBILLET', True);
      SetControlEnabled('BVOIRTIC', False);
      SetControlEnabled('BSELECT', False);
    end else
    begin
      if AffichageDetail then
      begin
        PDetail.Visible := True;
        SetControlEnabled('BENTETE', True);
        SetControlEnabled('BVOIRTIC', True);
        SetControlEnabled('BSELECT', True);
      end else
      begin
        PDetail.Visible := False;
        SetControlEnabled('BENTETE', False);
        SetControlEnabled('BVOIRTIC', False);
        SetControlEnabled('BSELECT', False);
      end;
      PPiecBil.Visible := False;
      SetControlProperty('PCUMDETAIL', 'Caption', sTitre);
      SetControlProperty('PCUMDETAILL', 'TextEffect', 1);
      GDet.ClearSelected;
      GDetPutGrid(TOBC);
      SetControlEnabled('BBILLET', ((TOBBilMdp <> nil) or (TOBPieMdp <> nil)));
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ZoneAccessible : vérifie si une cellule est saisissable
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOCTRLECAISSE.ZoneAccessible(ACol, ARow: Longint): Boolean;
var TOBL: TOB;
begin
  Result := False;
  if ACol = COLC_TOTSAISI then
  begin
    Result := True;
  end else if (ACol = COLC_TOTFDCAIS) or (ACol = COLC_TOTENCAIS) then
  begin
    TOBL := GetTOBLigne(ARow);
    if (TOBL <> nil) and (TOBL.FieldExists('GEREFDCAIS')) and (TOBL.GetValue('GEREFDCAIS') = 'X') then Result := True;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ZoneSuivanteOuOk : se positionne sur la cellule saisissable suivante
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: Boolean);
var Sens, Ind, OldCol, OldRow: Integer;
  OldEna, ChgLig: Boolean;
begin
  OldEna := GCum.SynEnabled;
  GCum.SynEnabled := False;
  Sens := -1;
  ChgLig := (GCum.Row <> ARow);
  OldCol := ACol;
  OldRow := ARow;
  if GCum.Row > ARow then Sens := 1 else if ((GCum.Row = ARow) and (ACol < GCum.Col)) then Sens := 1;
  ACol := GCum.Col;
  ARow := GCum.Row;
  Ind := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    Inc(Ind);
    if Ind > GCum.RowCount then Break;
    if Sens = 1 then
    begin
      if ((ACol = GCum.ColCount - 1) and (ARow = GCum.RowCount - 1)) then
      begin
        ACol := OldCol;
        ARow := OldRow;
        Break;
      end;
      if ChgLig then
      begin
        ACol := GCum.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GCum.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GCum.FixedCols;
      end;
    end else
    begin
      if ((ACol = GCum.FixedCols) and (ARow = 1)) then Break;
      if ChgLig then
      begin
        ACol := GCum.ColCount;
        ChgLig := False;
      end;
      if ACol > GCum.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GCum.ColCount - 1;
      end;
    end;
  end;
  GCum.SynEnabled := OldEna;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteTotSaisi : traitement du total saisi
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.TraiteTotSaisi(var ACol, ARow: Integer; var Cancel: Boolean);
var TOBE: TOB;
  dVal, dFdc, dMt: Double;
  sChoix: string;
  iRep: Integer;
begin
  TOBE := GetTOBLigne(ARow);
  if TOBE = nil then
  begin
    Cancel := True;
    Exit;
  end;
  FormateZoneSaisie(GCum, ACol, ARow);
  dVal := Valeur(GCum.Cells[ACol, ARow]);
  dVal := Arrondi(dVal, V_PGI.okDecV);
  if dVal < 0 then
  begin
    iRep := PGIAsk('Confirmez-vous la saisie d''un montant négatif ?', Ecran.Caption);
    if iRep <> mrYes then
    begin
      Cancel := True;
      Exit;
    end;
  end;
  dMt := TOBE.GetValue('TOTCALCULE');
  dMt := Arrondi(dMt, V_PGI.okDecV);
  if dVal <> dMt then
  begin
    TOBE.PutValue('CODEETAT', 'E'); // ERREUR
    sChoix := AGLLanceFiche('MFO', 'ECARTCAISSE', '', '', 'TOTO');
    if sChoix = 'ANNULER' then
    begin
      StCellCur := GCum.Cells[ACol, ARow];
      BoutonCalculetteClick('CLEAR');
      Cancel := True;
      Exit;
    end else
      if sChoix = 'FORCER' then
    begin
      StCellCur := GCum.Cells[ACol, ARow];
      TOBE.PutValue('ETAT', 'u');
      TOBE.PutValue('CODEETAT', 'F'); // FORCER
    end else
    begin
      TOBE.PutValue('ETAT', 'u');
      TOBE.PutValue('CODEETAT', 'E'); // ERREUR
      Cancel := True;
    end;
  end else
  begin
    StCellCur := GCum.Cells[ACol, ARow];
    TOBE.PutValue('ETAT', 'u');
    TOBE.PutValue('CODEETAT', 'O'); // OK
  end;
  TOBE.PutValue('TOTSAISI', dVal);
  if not Cancel then
  begin
    dFdc := TOBE.GetValue('TOTFDCAIS');
    if dFdc = 0 then
    begin
      if TOBE.FieldExists('TOTFDCAISCALC') then dFdc := TOBE.GetValue('TOTFDCAISCALC')
      else dFdc := 0;
    end;
    if dVal < dFdc then dFdc := dVal;
    TOBE.PutValue('TOTFDCAIS', dFdc);
    TOBE.PutValue('TOTENCAIS', (dVal - dFdc));
  end;
  GCumPutGrid(TOBE, ARow);
  if GCum.CanFocus then GCum.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteTotFdCais : traitement du fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.TraiteTotFdCais(var ACol, ARow: Integer; var Cancel: Boolean);
var TOBE: TOB;
  dTot, dFdc: Double;
begin
  TOBE := GetTOBLigne(ARow);
  if TOBE = nil then
  begin
    Cancel := True;
    Exit;
  end;
  FormateZoneSaisie(GCum, ACol, ARow);
  dFdc := Valeur(GCum.Cells[ACol, ARow]);
  if dFdc < 0 then
  begin
    PGIBox('Vous devez saisir un montant positif !', Ecran.Caption);
    Cancel := True;
    Exit;
  end;
  if dFdc <> TOBE.GetValue('TOTFDCAIS') then
  begin
    dTot := TOBE.GetValue('TOTSAISI');
    if dTot < dFdc then
    begin
      dFdc := dTot;
      PGIInfo('Le montant du fond de caisse est supérieur au montant disponible en caisse !', Ecran.Caption);
    end;
    TOBE.PutValue('TOTFDCAIS', dFdc);
    TOBE.PutValue('TOTENCAIS', (dTot - dFdc));
    GCumPutGrid(TOBE, ARow);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteTotEncais : traitement du total encaissé
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.TraiteTotEncais(var ACol, ARow: Integer; var Cancel: Boolean);
var TOBE: TOB;
  dVal, dTot: Double;
begin
  TOBE := GetTOBLigne(ARow);
  if TOBE = nil then
  begin
    Cancel := True;
    Exit;
  end;
  FormateZoneSaisie(GCum, ACol, ARow);
  dVal := Valeur(GCum.Cells[ACol, ARow]);
  if dVal < 0 then
  begin
    PGIBox('Vous devez saisir un montant positif !', Ecran.Caption);
    Cancel := True;
    Exit;
  end;
  if dVal <> TOBE.GetValue('TOTENCAIS') then
  begin
    dTot := TOBE.GetValue('TOTSAISI');
    if dTot < dVal then
    begin
      dVal := dTot;
      PGIInfo('Le montant encaissé est supérieur au montant disponible en caisse !', Ecran.Caption);
    end;
    TOBE.PutValue('TOTFDCAIS', (dTot - dVal));
    TOBE.PutValue('TOTENCAIS', dVal);
    GCumPutGrid(TOBE, ARow);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumEnter : OnEnter sur la grille cumul
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumEnter(Sender: TObject);
var Bc, Cancel: Boolean;
  ACol, ARow: Integer;
begin
  GCum.TitleBold := True;
  GCumColAligns;
  Bc := False;
  Cancel := False;
  /////GCum.Col := COLC_TOTSAISI ;
  ACol := COLC_TOTSAISI;
  ARow := GCum.Row;
  GCumRowEnter(GCum, ARow, Bc, False);
  GCumCellEnter(GCum, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumExit : OnExit sur la grille cumul
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumExit(Sender: TObject);
begin
  GCum.TitleBold := False;
  GCumColAligns;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumCellEnter : OnCellEnter sur la grille cumul
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var TOBE: TOB;
  sEtat: string;
begin
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  if not Cancel then
  begin
    if not ApplicEncours then StCellCur := GCum.Cells[GCum.Col, GCum.Row];
    if GCum.Col = COLC_TOTSAISI then
    begin
      TOBE := GetTOBLigne(ARow);
      if TOBE <> nil then sEtat := TOBE.GetValue('CODEETAT') else sEtat := '';
      if sEtat = '' then StCellCur := '';
      if not PiecBilEnCours then
      begin
        if (sEtat = '') and (PPiecBil.Visible) and
          (FOGetParamCaisse('GPK_CTRLPIECBIL') = 'X') then GotoPied;
      end else PiecBilEnCours := False;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumCellExit : OnCellExit sur la grille cumul
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var TOBE: TOB;
  sEtat: string;
begin
  //if csDestroying in ComponentState then Exit ;
  TOBE := GetTOBLigne(ARow);
  if (TOBE <> nil) then sEtat := TOBE.GetValue('CODEETAT') else sEtat := '';
  if (ACol = COLC_TOTSAISI) and (GCum.Row < ARow) and (GCum.Row < (GCum.RowCount - 1)) and
    (sEtat = '') and (GCum.Cells[ACol, ARow] = '') then
  begin
    // retour sur une ligne précédente
    GCum.Cells[ACol, ARow] := StCellCur;
    Exit;
  end;
  if (sEtat = 'O') or (sEtat = 'F') then
  begin
    if GCum.Cells[ACol, ARow] = StCellCur then Exit;
  end;
  if ACol = COLC_TOTSAISI then TraiteTotSaisi(ACol, ARow, Cancel) else
    if ACol = COLC_TOTFDCAIS then TraiteTotFdCais(ACol, ARow, Cancel) else
    if ACol = COLC_TOTENCAIS then TraiteTotEncais(ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumRowEnter : OnRowEnter sur la grille cumul
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  AfficheDetail(GCum.Row, False, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumRowExit : OnRowExit sur la grille cumul
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  //if csDestroying in ComponentState then Exit ;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumDessineCell : dessine une cellule de la grille des cumuls
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumDessineCell(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
var Color: TColor;
  TOBE: TOB;
  sEtat: string;
begin
  Color := 0;
  if (ACol = COLC_ETAT) and (ARow >= GCum.FixedRows) then
  begin
    TOBE := GetTOBLigne(ARow);
    if TOBE <> nil then sEtat := TOBE.GetValue('CODEETAT') else sEtat := '';
    if sEtat = 'O' then Color := $0000FF00 else // OK     = Vert
      if sEtat = 'F' then Color := $000F87FF else // FORCER = Orange
      if sEtat = 'E' then Color := $000000FF; // ERREUR = Rouge
    if Color <> 0 then
    begin
      Canvas.Font.Name := 'Wingdings';
      Canvas.Font.Size := 18;
      Canvas.Font.Style := [fsBold];
      Canvas.Font.Color := Color;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GCumPostDrawCell : affiche une image dans la colonne
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GCumPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
  TOBL: TOB;
  Ind: Integer;
  sType: string;
  XX, YY: Integer;
begin
  if GCum.RowHeights[ARow] <= 0 then Exit;
  if (ARow < GCum.FixedRows) or (ARow > GCum.TopRow + GCum.VisibleRowCount) then Exit;
  if (Acol = COLC_IMAGE) and (ImagesBtn <> nil) then
  begin
    TOBL := GetTOBLigne(ARow);
    if TOBL = nil then Exit;
    sType := TOBL.GetValue('MP_TYPEMODEPAIE');
    if sType = TYPEPAIEESPECE then Ind := 0 else
      if (sType = TYPEPAIECHEQUE) or (sType = TYPEPAIECHQDIFF) then Ind := 1 else
      if sType = TYPEPAIECB then Ind := 2 else
      if sType = TYPEPAIEBONACHAT then Ind := 3 else
      if sType = TYPEPAIEARRHES then Ind := 4 else
      if sType = TYPEPAIEAVOIR then Ind := 5 else
      if sType = TYPEPAIERESTEDU then Ind := 6 else Ind := 7;
    if (Ind >= 0) and (Ind < ImagesBtn.Count) then
    begin
      ARect := GCum.CellRect(ACol, ARow);
      Canvas.FillRect(ARect);
      XX := ((ARect.Right - ARect.Left - ImagesBtn.Width) div 2) + ARect.Left;
      if XX < ARect.Left then XX := ARect.Left;
      YY := ((ARect.Bottom - ARect.Top - ImagesBtn.Height) div 2) + ARect.Top;
      if YY < ARect.Top then YY := ARect.Top;
      ImagesBtn.DrawingStyle := dsTransparent;
      ImagesBtn.Draw(CanVas, XX, YY, Ind);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetRowEnter : OnRowEnter sur la grille détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  GDet.InvalidateRow(Ou);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetRowExit : OnRowExit sur la grille détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  GDet.InvalidateRow(Ou);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetEnter : OnEnter sur la grille détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetEnter(Sender: TObject);
begin
  GDet.TitleBold := True;
  GDetColAligns;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetExit : OnExit sur la grille détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetExit(Sender: TObject);
begin
  GDet.TitleBold := False;
  GDetColAligns;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetFlipSelection : OnFlipSelection sur la grille détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetFlipSelection(Sender: TObject);
var Ind: Integer;
  dTot: Double;
begin
  GDet.InvalidateRow(Gdet.Row);
  dTot := 0;
  for Ind := GDet.FixedRows to Gdet.RowCount - 1 do
  begin
    if (GDet.IsSelected(Ind)) and (GDet.Cells[COLD_MONTANT, Ind] <> '') then
    begin
      dTot := dTot + Valeur(GDet.Cells[COLD_MONTANT, Ind]);
    end;
  end;
  SetControlProperty('CUMMTDETAIL', 'Value', dTot);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GDetPostDrawCell : OnPostDrawCell sur la grille détail
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GDetPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var Triangle: array[0..2] of TPoint;
  ARect: TRect;
  OkSel: Boolean;
begin
  if Arow < Gdet.FixedRows then Exit;
  if (gdFixed in AState) and (ACol = COLD_IMAGE) then
  begin
    ARect := GDet.CellRect(ACol, ARow);
    Canvas.Brush.Color := GDet.FixedColor;
    Canvas.FillRect(ARect);
    if GDet.IsSelected(ARow) then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Canvas.Ellipse(((ARect.Left + ARect.Right) div 2) - 3, ((ARect.Top + ARect.Bottom) div 2) - 2,
        ((ARect.Left + ARect.Right) div 2) + 1, ((ARect.Top + ARect.Bottom) div 2) + 2);
      OkSel := True;
    end else OkSel := False;
    if ARow = GDet.Row then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Triangle[1].X := ((ARect.Left + ARect.Right) div 2);
      Triangle[1].Y := ((ARect.Top + ARect.Bottom) div 2);
      Triangle[0].X := Triangle[1].X - 5;
      Triangle[0].Y := Triangle[1].Y - 5;
      Triangle[2].X := Triangle[1].X - 5;
      Triangle[2].Y := Triangle[1].Y + 5;
      if OkSel then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletEnter : OnEnter sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GBilletEnter(Sender: TObject);
var Cancel: Boolean;
  ACol, ARow: Integer;
begin
  PiecBilEnCours := True;
  GPiece.TitleBold := True;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := True;
  GPiecBilColAligns(GBillet);
  Cancel := False;
  ACol := GBillet.Col;
  ARow := GBillet.Row;
  GBilletCellEnter(nil, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletExit : OnExit sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GBilletExit(Sender: TObject);
var ACol, ARow: Integer;
  Cancel: Boolean;
begin
  Cancel := False;
  ACol := GBillet.Col;
  ARow := GBillet.Row;
  GBilletCellExit(nil, ACol, ARow, Cancel);
  GPiece.TitleBold := False;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := False;
  GPiecBilColAligns(GBillet);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletCellEnter : OnCellEnter sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GBilletCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var OldEna: Boolean;
begin
  if GBillet.Col <> COLB_QUANTITE then
  begin
    OldEna := GBillet.SynEnabled;
    GBillet.SynEnabled := False;
    GBillet.Col := COLB_QUANTITE;
    GBillet.SynEnabled := OldEna;
  end;
  StPiecBilCur := GBillet.Cells[GBillet.Col, GBillet.Row];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletCellExit : OnCellExit sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GBilletCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  //if csDestroying in ComponentState then Exit ;
  if GBillet.Cells[ACol, ARow] = StPiecBilCur then Exit;
  if ACol = COLB_QUANTITE then
  begin
    FormateZoneSaisie(GBillet, ACol, ARow);
    CalculMontantPieceBillet(GBillet, TOBBilMdp, ACol, ARow, Cancel);
    if not Cancel then CalculTotalPieceBillet(TOBBilMdp, TOBPieMdp, True);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletColumnWidthsChanged : OnColumnWidthsChanged sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GBilletColumnWidthsChanged(Sender: TObject);
var Coord: TRect;
  ALeft, AWidth: Integer;
begin
  Coord := GBillet.CellRect(COLB_MONTANT, 0);
  ALeft := Coord.Left + 1;
  AWidth := GBillet.ColWidths[COLB_MONTANT] + 1;
  SetControlProperty('CUMMTBILLET', 'Left', ALeft);
  SetControlProperty('CUMMTBILLET', 'Width', AWidth);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceEnter : OnEnter sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GPieceEnter(Sender: TObject);
var Cancel: Boolean;
  ACol, ARow: Integer;
begin
  PiecBilEnCours := True;
  GPiece.TitleBold := True;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := True;
  GPiecBilColAligns(GBillet);
  Cancel := False;
  ACol := GPiece.Col;
  ARow := GPiece.Row;
  GPieceCellEnter(nil, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceExit : OnExit sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GPieceExit(Sender: TObject);
var ACol, ARow: Integer;
  Cancel: Boolean;
begin
  Cancel := False;
  ACol := GPiece.Col;
  ARow := GPiece.Row;
  GPieceCellExit(nil, ACol, ARow, Cancel);
  GPiece.TitleBold := False;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := False;
  GPiecBilColAligns(GBillet);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceCellEnter : OnCellEnter sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GPieceCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var OldEna: Boolean;
begin
  if GPiece.Col <> COLB_QUANTITE then
  begin
    OldEna := GPiece.SynEnabled;
    GPiece.SynEnabled := False;
    GPiece.Col := COLB_QUANTITE;
    GPiece.SynEnabled := OldEna;
  end;
  StPiecBilCur := GPiece.Cells[GPiece.Col, GPiece.Row];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceCellExit : OnCellExit sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GPieceCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  //if csDestroying in ComponentState then Exit ;
  if GPiece.Cells[ACol, ARow] = StPiecBilCur then Exit;
  if ACol = COLB_QUANTITE then
  begin
    FormateZoneSaisie(GPiece, ACol, ARow);
    CalculMontantPieceBillet(GPiece, TOBPieMdp, ACol, ARow, Cancel);
    if not Cancel then CalculTotalPieceBillet(TOBBilMdp, TOBPieMdp, True);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceColumnWidthsChanged : OnColumnWidthsChanged sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GPieceColumnWidthsChanged(Sender: TObject);
var Coord: TRect;
  ALeft, AWidth: Integer;
  Ctrl: TControl;
begin
  Coord := GPiece.CellRect(COLB_MONTANT, 0);
  ALeft := Coord.Left + 1;
  AWidth := GPiece.ColWidths[COLB_MONTANT] + 1;
  SetControlProperty('CUMMTPIECE', 'Left', ALeft);
  SetControlProperty('CUMMTPIECE', 'Width', AWidth);
  Ctrl := GetControl('PPIECE');
  if Ctrl <> nil then
  begin
    ALeft := Ctrl.Left + 1;
    AWidth := Coord.Right;
    SetControlProperty('CUMMONTANT', 'Left', ALeft);
    SetControlProperty('CUMMONTANT', 'Width', AWidth);
    SetControlProperty('CUMMTDETAIL', 'Left', ALeft);
    SetControlProperty('CUMMTDETAIL', 'Width', AWidth);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormKeyDown :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_RETURN: Key := VK_TAB;
    VK_F10: if Shift = [] then
      begin
        Key := 0;
        BoutonCalculetteClick('ENTER');
      end;
    VK_HOME: if Shift = [ssCtrl] then
      begin
        Key := 0;
        GotoEntete;
      end;
    VK_END: if Shift = [ssCtrl] then
      begin
        Key := 0;
        GotoPied;
      end;
    VK_ESCAPE: if Shift = [] then
      begin
        Key := 0;
        BoutonCalculetteClick('CLEAR');
      end;
  end;
  if (Key = VK_DOWN) or ((Key = VK_TAB) and (Shift = [])) then
  begin
    if (Ecran.ActiveControl = GBillet) and (GBillet.Row = GBillet.RowCount - 1) then
    begin
      if GPiece.CanFocus then
      begin
        GPiece.Row := GPiece.FixedRows;
        GPiece.Col := COLB_QUANTITE;
        GPiece.SetFocus;
      end;
    end;
    if (Ecran.ActiveControl = GPiece) and (GPiece.Row = GPiece.RowCount - 1) then
    begin
      if GBillet.CanFocus then
      begin
        GBillet.Row := GBillet.FixedRows;
        GBillet.Col := COLB_QUANTITE;
        GBillet.SetFocus;
      end;
    end;
  end;
  if (Key = VK_UP) or ((Key = VK_TAB) and (Shift = [ssShift])) then
  begin
    if (Ecran.ActiveControl = GBillet) and (GBillet.Row = GBillet.FixedRows) then
    begin
      if GPiece.CanFocus then
      begin
        GPiece.Row := GPiece.RowCount - 1;
        GPiece.Col := COLB_QUANTITE;
        GPiece.SetFocus;
      end;
    end;
    if (Ecran.ActiveControl = GPiece) and (GPiece.Row = GPiece.FixedRows) then
    begin
      if GBillet.CanFocus then
      begin
        GBillet.Row := GBillet.RowCount - 1;
        GBillet.Col := COLB_QUANTITE;
        GBillet.SetFocus;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculMontantPieceBillet : calcul du montant d'une ligne de pièces et de billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.CalculMontantPieceBillet(GS: THGrid; TOBParent: TOB; ACol, ARow: Longint; var Cancel: Boolean);
var TOBL, TOBE: TOB;
  dQte, dMt: Double;
  iLig, iNbdec: Integer;
begin
  TOBL := nil;
  iLig := ARow - GS.FixedRows;
  TOBE := GetTOBLigne(GCum.Row);
  if TOBE <> nil then iNbDec := TOBE.GetValue('D_DECIMALE') else iNbDec := V_PGI.OkDecV;
  if (iLig >= 0) and (iLig < TOBParent.Detail.Count) then TOBL := TOBParent.Detail[iLig];
  if TOBL <> nil then dQte := Valeur(GS.Cells[ACol, ARow]) else dQte := 0;
  if dQte < 0 then
  begin
    PGIBox('Vous devez saisir une quantité positive !', Ecran.Caption);
    Cancel := True;
    Exit;
  end;
  TOBL.PutValue('QUANTITE', dQte);
  dMt := dQte * TOBL.GetValue('GPI_VALEUR');
  TOBL.PutValue('MONTANT', dMt);
  GS.Cells[COLB_MONTANT, ARow] := StrS(dMt, iNbdec);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculTotalPieceBillet : calcul le total des pièces et billets
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOCTRLECAISSE.CalculTotalPieceBillet(TOBBil, TOBPie: TOB; Affiche: Boolean): Double;
var TOBL: TOB;
  Ind: Integer;
  dBil, dPiec, dQte, dVal: Double;
begin
  dBil := 0;
  if TOBBil <> nil then
  begin
    for Ind := 0 to TOBBil.Detail.Count - 1 do
    begin
      TOBL := TOBBil.Detail[Ind];
      dQte := TOBL.GetValue('QUANTITE');
      dVal := TOBL.GetValue('GPI_VALEUR');
      dBil := dBil + Arrondi((dQte * dVal), V_PGI.okDecV);
      Debug('Quantité=' + FloatToStr(dQte) + ' Valeur=' + FloatToStr(dVal) + ' Montant=' + FloatToStr((dQte * dVal)) + ' Total billets=' + FloatToStr(dBil));
    end;
  end;
  if Affiche then SetControlProperty('CUMMTBILLET', 'Value', dBil);
  dPiec := 0;
  if TOBPie <> nil then
  begin
    for Ind := 0 to TOBPie.Detail.Count - 1 do
    begin
      TOBL := TOBPie.Detail[Ind];
      dQte := TOBL.GetValue('QUANTITE');
      dVal := TOBL.GetValue('GPI_VALEUR');
      dPiec := dPiec + Arrondi((dQte * dVal), V_PGI.okDecV);
      Debug('Quantité=' + FloatToStr(dQte) + ' Valeur=' + FloatToStr(dVal) + ' Montant=' + FloatToStr((dQte * dVal)) + ' Total pièces=' + FloatToStr(dPiec));
    end;
  end;
  if Affiche then SetControlProperty('CUMMTPIECE', 'Value', dPiec);
  Result := dBil + dPiec;
  Debug('Total billets et pièces=' + FloatToStr((dBil + dPiec)));
  if Affiche then SetControlProperty('CUMMONTANT', 'Value', Result);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormateZoneSaisie : formatage de la cellule
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOCTRLECAISSE.VerifiePieceBillet(ModePaie: string; TotSaisi: Double): Boolean;
var TOBB, TOBP: TOB;
  dTot: Double;
begin
  Result := True;
  if TOBBillet <> nil then
    TOBB := TOBBillet.FindFirst(['MP_MODEPAIE'], [ModePaie], False)
  else TOBB := nil;

  if TOBPiece <> nil then
    TOBP := TOBPiece.FindFirst(['MP_MODEPAIE'], [ModePaie], False)
  else TOBP := nil;

  if (TOBB <> nil) or (TOBP <> nil) then
  begin
    dTot := CalculTotalPieceBillet(TOBB, TOBP, False);
    Debug('Saisi=' + FloatToStr(TotSaisi) + ' Calculé=' + FloatToStr(dTot) + ' Ecart=' + FloatToStr((TotSaisi - dTot)));
    if (dTot <> 0) or (FOGetParamCaisse('GPK_CTRLPIECBIL') = 'X') then
    begin
      Result := (Arrondi(dTot, V_PGI.okDecV) = Arrondi(TotSaisi, V_PGI.okDecV));
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormateZoneSaisie : formatage de la cellule
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.FormateZoneSaisie(GS: THGrid; ACol, ARow: Longint);
var Stg: string;
  iNbDec: Integer;
  TOBL: TOB;
begin
  iNbDec := 0;
  if GS = GCum then
  begin
    TOBL := GetTOBLigne(GCum.Row);
    if TOBL <> nil then iNbDec := TOBL.GetValue('D_DECIMALE') else iNbDec := V_PGI.OkDecV;
  end;
  Stg := StrS(Valeur(GS.Cells[ACol, ARow]), iNbDec);
  GS.Cells[ACol, ARow] := Stg;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GotoEntete : se positionne sur la grille (cumul) de l'en-tête
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GotoEntete;
begin
  if GCum.CanFocus then GCum.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GotoPied : se positionne sur la grille (détail ou billet) de pied
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.GotoPied;
begin
  if PPiecBil.Visible then
  begin
    if GBillet.CanFocus then
    begin
      GBillet.Row := GBillet.FixedRows;
      GBillet.Col := COLB_QUANTITE;
      GBillet.SetFocus;
    end;
  end else
  begin
    if GDet.CanFocus then
    begin
      GDet.Row := GBillet.FixedRows;
      GDet.SetFocus;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BoutonCalculetteClick : interprète les boutons de la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BoutonCalculetteClick(Val: string);
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  if Val = 'ENTER' then
  begin
    if Ecran.ActiveControl = GCum then
      BSuivantClick(nil)
    else
      if GCum.Col = COLC_TOTSAISI then AppliqueTotal else FOSimuleClavier(VK_TAB);
  end else
    if Val = 'CLEAR' then
  begin
    if PGIAsk('Confirmez-vous l''abandon de la saisie ?', Ecran.Caption) = mrYes then
    begin
      TFVierge(Ecran).BFermeClick(Self);
      Ecran.Close;
    end;
  end;
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieCalculette : prend en compte un chiffre saisi sur la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.SaisieCalculette(Val: Double);
var ACol, ARow: Integer;
  Texte: string;
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  Texte := FloatToStr(Val);
  if Ecran.ActiveControl = GCum then
  begin
    ACol := GCum.Col;
    ARow := GCum.Row;
    GCum.Cells[ACol, ARow] := Texte;
  end else
    if Ecran.ActiveControl = GBillet then
  begin
    ACol := GBillet.Col;
    ARow := GBillet.Row;
    GBillet.Cells[ACol, ARow] := Texte;
  end else
    if Ecran.ActiveControl = GPiece then
  begin
    ACol := GPiece.Col;
    ARow := GPiece.Row;
    GPiece.Cells[ACol, ARow] := Texte;
  end;
  FOSimuleClavier(VK_TAB);
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AppliqueTotal : recopie du total (détail ou pièces/billet) sur la grille des cumuls
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.AppliqueTotal;
var ACol, ARow: Integer;
  NomChamp: string;
  dMt: Double;
  Ctrl: TControl;
begin
  if ValideEnCours then Exit;
  if ApplicEncours then Exit;
  ApplicEncours := True;
  dMt := 0;
  NomChamp := '';
  if Ecran.ActiveControl = GDet then NomChamp := 'CUMMTDETAIL' else
    if (Ecran.ActiveControl = GBillet) or (Ecran.ActiveControl = GPiece) then NomChamp := 'CUMMONTANT';
  if NomChamp <> '' then
  begin
    Ctrl := GetControl(NomChamp);
    if (Ctrl <> nil) and (Ctrl is THNumEdit) then dMt := THNumEdit(Ctrl).Value;
  end;
  if GCum.CanFocus then
  begin
    ACol := GCum.Col;
    ARow := GCum.Row;
    GCum.Cells[ACol, ARow] := StrS(dMt, V_PGI.OkDecV);
    GCum.SetFocus;
    FOSimuleClavier(VK_TAB);
  end;
  ApplicEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BLeftClick : OnClick du bouton BLEFT
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BLeftClick(Sender: TObject);
var Key: Word;
  Shift: Boolean;
begin
  if ValideEnCours then Exit;
  if Ecran.ActiveControl = GDet then
  begin
    Key := VK_UP;
    Shift := False;
  end else
  begin
    Key := VK_TAB;
    Shift := True;
  end;
  FOSimuleClavier(Key, Shift);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BRightClick : OnClick du bouton BRIGHT
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BRightClick(Sender: TObject);
var Key: Word;
begin
  if ValideEnCours then Exit;
  if Ecran.ActiveControl = GDet then Key := VK_DOWN else Key := VK_TAB;
  FOSimuleClavier(Key);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BUpClick : OnClick du bouton BUP
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BUpClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  FOSimuleClavier(VK_UP);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BDownClick : OnClick du bouton BDOWN
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BDownClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  FOSimuleClavier(VK_DOWN);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BEnteteClick : OnClick du bouton BENTETE
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BEnteteClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  GotoEntete;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BDetailClick : OnClick du bouton BDETAIL
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BDetailClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  AfficheDetail(GCum.Row, True, True);
  if (PDetail.Visible) and (GDet.CanFocus) then GDet.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BBilletClick : OnClick du bouton BBILLET
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BBilletClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  AfficheDetail(GCum.Row, True, False);
  if (PPiecBil.Visible) and (GBillet.CanFocus) then GBillet.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BVoirTicClick : OnClick du bouton BVOIRTIC
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BVoirTicClick(Sender: TObject);
var TOBE, TOBL: TOB;
  Ind: Integer;
  CleDoc: R_CleDoc;
begin
  if ValideEnCours then Exit;
  if not PDetail.Visible then Exit;
  TOBE := GetTOBLigne(GCum.Row);
  if TOBE = nil then Exit;
  Ind := GDet.Row - GDet.FixedRows;
  if (Ind < 0) or (Ind >= TOBE.Detail.Count) then Exit;
  TOBL := TOBE.Detail[Ind];
  CleDoc := TOB2CleDoc(TOBL);
  FOImprimerLaPiece(nil, nil, nil, nil, CleDoc, True, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BSelectClick : OnClick du bouton BSELECT
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BSelectClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  if not PDetail.Visible then Exit;
  if GDet.CanFocus then GDet.SetFocus;
  if Ecran.ActiveControl = GDet then
  begin
    FOSimuleClavier(VK_SPACE);
    FOSimuleClavier(VK_DOWN, False, 10);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BSuivantClick : OnClick du bouton BSUIVANT
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BSuivantClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  TFVierge(Ecran).BValiderClick(Self);
  if LastError = 0 then Ecran.Close;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BPrecedentClick : OnClick du bouton BPRECEDENT
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOCTRLECAISSE.BPrecedentClick(Sender: TObject);
begin
  if ValideEnCours then Exit;
  TFVierge(Ecran).BFermeClick(Self);
  TFVierge(Ecran).Retour := 'PREV';
  Ecran.Close;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/06/2003
Modifié le ... : 03/06/2003
Description .. : Extrait le nombre saisi pour un code pièce ou billet
Mots clefs ... : FO
*****************************************************************}

function ExtractPieceBillet(TOBPiecBil: TOB; DetailPiecBil: string): double;
var
  St1, St2, St3: string;
  Code: integer;
begin
  Result := 0;
  if TOBPiecBil = nil then Exit;
  Code := TOBPiecBil.GetValue('GPI_PIECEBILLET');
  St2 := TOBPiecBil.GetValue('GPI_TYPE');
  St1 := Copy(St2, 1, 1) + IntToStr(Code) + '=';
  while DetailPiecBil <> '' do
  begin
    St2 := ReadTokenSt(DetailPiecBil);
    if St1 = Copy(St2, 1, Length(St1)) then
    begin
      St3 := Copy(St2, Length(St1) + 1, MaxInt);
      Result := StrTofloat(St3);
      break;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/06/2003
Modifié le ... : 03/06/2003
Description .. : charge les TOB des pieces et billet pour un mode de paiement
Mots clefs ... :
*****************************************************************}

procedure ChargeQuantitePiecBil(TOBPrinc, TOBLigEch: TOB; ModePaie: string);
var
  Stg: string;
  TOBP: TOB;
  ii, jj: integer;
  dQt, dTot: double;
begin
  Stg := TOBLigEch.GetValue('GJM_PIECBILTOT');
  if Stg = '' then Exit;

  ii := 1;
  while (ii < TOBPrinc.Detail.Count) and (ii < 3) do
  begin
    TOBP := TOBPrinc.Detail[ii].FindFirst(['MP_MODEPAIE'], [ModePaie], False);
    if TOBP <> nil then
    begin
      for jj := 0 to TOBP.Detail.Count - 1 do
      begin
        dQt := ExtractPieceBillet(TOBP.Detail[jj], Stg);
        if dQt <> 0 then
        begin
          TOBP.Detail[jj].PutValue('QUANTITE', dQt);
          dTot := dQt * TOBP.Detail[jj].GetValue('GPI_VALEUR');
          TOBP.Detail[jj].PutValue('MONTANT', dTot);
        end;
      end;
    end;
    Inc(ii);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeMtCtrlCaisse : Chargement des montants du contrôle caisse par mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure ChargeMtCtrlCaisse(TOBPrinc: TOB; Caisse: string; NumZ: Integer);
var sSql, sMdp: string;
  QQ: TQuery;
  TOBEch, TOBE, TOBF: TOB;
  Ind: Integer;
begin
  if (TOBPrinc = nil) or (TOBPrinc.Detail.Count <= 0) then Exit;
  TOBEch := TOBPrinc.Detail[0];
  TOBF := TOB.Create('Les montants de caisse', nil, -1);
  sSql := 'select * from CTRLCAISMT'
    + ' where GJM_CAISSE="' + Caisse + '" and GJM_NUMZCAISSE="' + IntToStr(NumZ) + '" '
    + FOFabriqueListeMDP('GPK_MDPCTRLCAIS', 'GJM_MODEPAIE');
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then
    TOBF.LoadDetailDB('CTRLCAISMT', '', '', QQ, False);
  Ferme(QQ);
  for Ind := 0 to TOBF.Detail.Count - 1 do
  begin
    sMdp := TOBF.Detail[Ind].GetValue('GJM_MODEPAIE');
    TOBE := TOBEch.FindFirst(['MP_MODEPAIE'], [sMdp], False);
    if TOBE <> nil then
    begin
      TOBE.AddChampSupValeur('GJM_TOTALDEV', TOBF.Detail[Ind].GetValue('GJM_TOTALDEV'));
      TOBE.AddChampSupValeur('GJM_ECARTDEV', TOBF.Detail[Ind].GetValue('GJM_ECARTDEV'));
      TOBE.AddChampSupValeur('GJM_FDCAISOUV', TOBF.Detail[Ind].GetValue('GJM_FDCAISOUV'));
      TOBE.AddChampSupValeur('GJM_PIECBILOUV', TOBF.Detail[Ind].GetValue('GJM_PIECBILOUV'));
      TOBE.AddChampSupValeur('GJM_PIECBILTOT', TOBF.Detail[Ind].GetValue('GJM_PIECBILTOT'));
      TOBE.PutValue('TOTSAISI', TOBF.Detail[Ind].GetValue('GJM_TOTALDEV'));
      ChargeQuantitePiecBil(TOBPrinc, TOBE, sMdP);
    end;
  end;
  TOBF.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeFondCaisse : Chargement du fond de caisse par mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure ChargeFondCaisse(TOBEch: TOB; Caisse: string; NumZ: Integer);
var sSql, sMdp: string;
  QQ: TQuery;
  TOBE, TOBF: TOB;
  Ind: Integer;
  dTot: Double;
begin
  sSql := 'select GPE_MODEPAIE,GPE_MONTANTENCAIS from LIGNE,PIEDECHE'
    + ' where GL_NATUREPIECEG=' + FOGetNatureTicket(False, True) + ' and GL_CAISSE="' + Caisse + '"'
    + ' and GL_DATEPIECE="' + USDateTime(FOGetDateOuv(Caisse, NumZ)) + '"'
    + ' and GL_CODEARTICLE="' + FOGetParamCaisse('GPK_FDCAISSE') + '"'
    + ' and GPE_NATUREPIECEG=GL_NATUREPIECEG and GPE_SOUCHE=GL_SOUCHE'
    + ' and GPE_NUMERO=GL_NUMERO and GPE_INDICEG=GL_INDICEG'
    + ' and GPE_NUMZCAISSE="' + IntToStr(NumZ) + '" '
    + FOFabriqueListeMDP('GPK_MDPCTRLCAIS', 'GPE_MODEPAIE');
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then
  begin
    TOBF := TOB.Create('Les fonds de caisse', nil, -1);
    TOBF.LoadDetailDB('PIEDECHE', '', '', QQ, False);
    for Ind := 0 to TOBF.Detail.Count - 1 do
    begin
      sMdp := TOBF.Detail[Ind].GetValue('GPE_MODEPAIE');
      TOBE := TOBEch.FindFirst(['MP_MODEPAIE'], [sMdp], False);
      if TOBE <> nil then
      begin
        if TOBE.FieldExists('TOTFDCAISCALC') then
        begin
          dTot := TOBE.GetValue('TOTFDCAISCALC');
        end else
        begin
          TOBE.AddChampSup('TOTFDCAISCALC', False);
          dTot := 0;
        end;
        dTot := dTot + TOBF.Detail[Ind].GetValue('GPE_MONTANTENCAIS');
        TOBE.PutValue('TOTFDCAISCALC', dTot);
      end;
    end;
    TOBF.Free;
  end;
  Ferme(QQ);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeDetail : Chargement des lignes de détail
///////////////////////////////////////////////////////////////////////////////////////

procedure ChargeDetail(TOBEch: TOB; Caisse: string; NumZ: Integer);
var sSql, sMdp, sCurMdp: string;
  TOBLig, TOBL, TOBE, TOBC: TOB;
  QQ: TQuery;
  Ind: Integer;
  dTot: Double;
begin
  TOBLig := nil;
  TOBE := nil;
  sCurMdp := '';
  // Chargement des échéances
  sSql := FOSelectCtrlCaisse;
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then
  begin
    TOBLig := TOB.Create('Les echeances', nil, -1);
    TOBLig.LoadDetailDB('PIEDECHE', '', '', QQ, False);
  end;
  Ferme(QQ);
  // fusion de la TOB détail dans la TOB des cumuls
  if TOBLig <> nil then
  begin
    for Ind := 0 to TOBLig.Detail.Count - 1 do
    begin
      TOBL := TOBLig.Detail[Ind];
      sMdp := TOBL.GetValue('GPE_MODEPAIE');
      if (TOBE = nil) or (sMdp <> sCurMdp) then TOBE := TOBEch.FindFirst(['MP_MODEPAIE'], [sMdp], False);
      if TOBE <> nil then
      begin
        sCurMdp := sMdp;
        TOBC := TOB.Create('Cumul', TOBE, -1);
        TOBC.Dupliquer(TOBL, True, True);
        TOBC.AddChampSupValeur('D_SYMBOLE', TOBE.GetValue('D_SYMBOLE'));
        dTot := TOBE.GetValue('TOTCALCULE') + TOBC.GetValue('GPE_MONTANTENCAIS');
        TOBE.PutValue('TOTCALCULE', dTot);
        {*********** à voir **************************************************
        dMt := TOBE.GetValue('TOTSAISI') ;
        if dMt = dTot then
           BEGIN
           TOBE.PutValue('ETAT', 'u') ;
           TOBE.PutValue('CODEETAT', 'O') ;      // OK
           END else
           BEGIN
           TOBE.PutValue('ETAT', 'u') ;
           TOBE.PutValue('CODEETAT', 'E') ;      // ERREUR
           END ;
        *********** à voir **************************************************}
      end;
    end;
    TOBLig.Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOChargeTOBPiecBil : charge les TOB des pieces et billet pour un mode de paiement
///////////////////////////////////////////////////////////////////////////////////////

procedure FOChargeTOBPiecBil(TOBEch, TOBBillet, TOBPiece: TOB; ModePaie, Devise, Prefixe: string);
var sSql: string;
  TOBC: TOB;
  QQ: TQuery;
  dTot: Double;
begin
  if TOBEch = nil then Exit;
  dTot := 0;
  // billets
  TOBC := TOBBillet.FindFirst([Prefixe + '_MODEPAIE'], [ModePaie], False);
  if TOBC = nil then
  begin
    sSql := 'select GPI_PIECEBILLET,GPI_LIBELLE,GPI_VALEUR,GPI_TYPE from PARPIECBIL '
      + 'where GPI_TYPE="BIL" and GPI_DEVISE="' + Devise + '" '
      + 'order by GPI_VALEUR desc';
    QQ := OpenSQL(sSql, True);
    if not QQ.EOF then
    begin
      TOBC := TOB.Create('Cumul', TOBBillet, -1);
      TOBC.Dupliquer(TOBEch, True, True);
      TOBC.LoadDetailDB('PARPIECBIL', '', '', QQ, False);
      if TOBC.Detail.Count > 0 then
      begin
        TOBC.Detail[0].AddChampSupValeur('QUANTITE', dTot, True);
        TOBC.Detail[0].AddChampSupValeur('MONTANT', dTot, True);
      end;
    end;
    Ferme(QQ);
  end;
  // pièces
  TOBC := TOBPiece.FindFirst([Prefixe + '_MODEPAIE'], [ModePaie], False);
  if TOBC = nil then
  begin
    sSql := 'select GPI_PIECEBILLET,GPI_LIBELLE,GPI_VALEUR,GPI_TYPE from PARPIECBIL '
      + 'where GPI_TYPE="PIE" and GPI_DEVISE="' + Devise + '" '
      + 'order by GPI_VALEUR desc';
    QQ := OpenSQL(sSql, True);
    if not QQ.EOF then
    begin
      TOBC := TOB.Create('Cumul', TOBPiece, -1);
      TOBC.Dupliquer(TOBEch, True, True);
      TOBC.LoadDetailDB('PARPIECBIL', '', '', QQ, False);
      if TOBC.Detail.Count > 0 then
      begin
        TOBC.Detail[0].AddChampSupValeur('QUANTITE', dTot, True);
        TOBC.Detail[0].AddChampSupValeur('MONTANT', dTot, True);
      end;
    end;
    Ferme(QQ);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOChargeTablesCtrlCaisse : Chargement en TOB des tables du contrôle caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure FOChargeTablesCtrlCaisse(TOBPrinc: TOB; Caisse: string; NumZ: Integer);
var sSql, sType, sMdp, sDev, sRang, sFdCais, sLstMdpFd: string;
  TOBEch, TOBBillet, TOBPiece, TOBE: TOB;
  QQ: TQuery;
  Ind: Integer;
  dTot: Double;
begin
  // Création des TOB
  // ----------------
  // La TOB principale du contrôle caisse est constitutée de 3 TOB filles
  //  * rang 0: la TOB des modes de paiement qui doivent subir le contrôle caisse.
  //    Au 1er niveau de l'arborescence on trouve les modes de paiement et les cumuls
  //    du contrôle caisse (MP_MODEPAIE,MP_LIBELLE,MP_TYPEMODEPAIE,MP_DEVISEFO,
  //    D_SYMBOLE,D_DECIMALE,GJM_TOTALDEV,GJM_ECARTDEV,GJM_FDCAISSEDEV,GJM_FDCAISOUV,
  //    GJM_PIECBILOUV,GJM_PIECBILTOT,TOTCALCULE,TOTSAISI,TOTENCAIS,TOTFDCAIS,
  //    TOTFDCAISCALC,ETAT,CODEETAT,GEREFDCAIS),
  //    puis en fille la liste des échéances de la journée pour le mode de paiement
  //    (GPE_NATUREPIECEG,GPE_SOUCHE,GPE_NUMERO,GPE_INDICEG,GPE_NUMECHE,GPE_TIERS,
  //     GPE_MONTANTENCAIS,D_SYMBOLE).
  //  * rang 1: la TOB des billets. Au 1er niveau de l'arborescence on trouve les modes
  //    de paiement (MP_MODEPAIE,GPI_DEVISE),
  //    puis en fille la quantité saisie  par type de billet
  //    (GPI_LIBELLE,GPI_VALEUR,QUANTITE,MONTANT).
  //  * rang 2: la TOB des pièces qui a une structure identique à la TOB des billets.
  //
  if TOBPrinc = nil then Exit;
  if TOBPrinc.Detail.Count < 1 then TOBEch := TOB.Create('Les echeances', TOBPrinc, -1)
  else TOBEch := TOBPrinc.Detail[0];
  if TOBPrinc.Detail.Count < 2 then TOBBillet := TOB.Create('Les billets', TOBPrinc, -1)
  else TOBBillet := TOBPrinc.Detail[1];
  if TOBPrinc.Detail.Count < 3 then TOBPiece := TOB.Create('Les pieces', TOBPrinc, -1)
  else TOBPiece := TOBPrinc.Detail[2];
  dTot := 0;
  sLstMdpFd := FOGetParamCaisse('GPK_MDPFDCAIS');
  // Chargement des modes de paiements
  sSql := 'select MP_MODEPAIE,MP_LIBELLE,MP_TYPEMODEPAIE,MP_DEVISEFO,D_SYMBOLE,D_DECIMALE'
    + ' from MODEPAIE,DEVISE '
    + ' where MP_UTILFO="X" and ' + FOFabriqueListeMDP('GPK_MDPCTRLCAIS', 'MP_MODEPAIE')
    + ' and MP_DEVISEFO=D_DEVISE'
    + ' order by MP_TYPEMODEPAIE,MP_MODEPAIE';
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then TOBEch.LoadDetailDB('MODEPAIE', '', '', QQ, False);
  Ferme(QQ);
  for Ind := 0 to TOBEch.Detail.Count - 1 do
  begin
    TOBE := TOBEch.Detail[Ind];
    sMdp := TOBE.GetValue('MP_MODEPAIE');
    if FOTestCodeEuro(TOBE.GetValue('MP_DEVISEFO')) then TOBE.PutValue('D_SYMBOLE', SIGLEEURO);
    // définition de l'ordre de présentation
    sType := TOBE.GetValue('MP_TYPEMODEPAIE');
    if sType = TYPEPAIEESPECE then sRang := '1' else
      if (sType = TYPEPAIECHEQUE) or (sType = TYPEPAIECHQDIFF) then sRang := '2' else
      if sType = TYPEPAIECB then sRang := '3' else sRang := '4';
    sDev := TOBE.GetValue('MP_DEVISEFO');
    if sDev = V_PGI.DevisePivot then sRang := sRang + '1' else
      if sDev = V_PGI.DeviseFongible then sRang := sRang + '2' else sRang := sRang + '3';
    sRang := sRang + TOBE.GetValue('MP_TYPEMODEPAIE') + sMdp;
    TOBE.AddChampSupValeur('RANG', sRang);
    TOBE.AddChampSupValeur('TOTCALCULE', dTot);
    TOBE.AddChampSupValeur('TOTSAISI', dTot);
    TOBE.AddChampSupValeur('TOTENCAIS', dTot);
    TOBE.AddChampSupValeur('TOTFDCAIS', dTot);
    TOBE.AddChampSupValeur('ETAT', '');
    TOBE.AddChampSupValeur('CODEETAT', '');
    if sType = TYPEPAIEESPECE then
    begin
      // le fond caisse est-il géré pour ce mode de paiement ?
      if FOStrCmp(TOBE.GetValue('MP_MODEPAIE'), sLstMdpFd, True) then sFdCais := 'X'
      else sFdCais := '-';
      TOBE.AddChampSupValeur('GEREFDCAIS', sFdCais);
      FOChargeTOBPiecBil(TOBE, TOBBillet, TOBPiece, sMdp, sDev, 'MP');
    end;
  end;
  TOBEch.Detail.Sort('RANG');
  // Chargement du fond de caisse  et du détail des échéances
  ChargeMtCtrlCaisse(TOBPrinc, Caisse, NumZ);
  ChargeFondCaisse(TOBEch, Caisse, NumZ);
  ChargeDetail(TOBEch, Caisse, NumZ);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJEcart : création d'un ticket d'écart
///////////////////////////////////////////////////////////////////////////////////////

procedure MAJEcart(TOBLig: TOB; Vendeur, Titre: string);
var dMtEcartDev: Double;
  sArt, sDev, sMdp: string;
begin
  sArt := FOGetParamCaisse('GPK_CTRLECART');
  sMdp := TOBLig.GetValue('MP_MODEPAIE');
  sDev := TOBLig.GetValue('MP_DEVISEFO');
  dMtEcartDev := TOBLig.GetValue('TOTSAISI') - TOBLig.GetValue('TOTCALCULE');
  dMtEcartDev := Arrondi(dMtEcartDev, V_PGI.OkDecV);
  if dMtEcartDev <> 0 then FOGenereTicket(sArt, Vendeur, sMdp, sDev, '', Titre, dMtEcartDev, True, True, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOMAJTotalPieceBillet : mise à jours du détail en pièces et billets
///////////////////////////////////////////////////////////////////////////////////////

function FOMAJTotalPieceBillet(TOBBillet, TOBPiece: TOB; ModePaie, Prefixe: string): string;
var TOBC, TOBL: TOB;
  Ind: Integer;
  dQte: Double;
begin
  Result := '';
  if TOBBillet <> nil then TOBC := TOBBillet.FindFirst([Prefixe + '_MODEPAIE'], [ModePaie], False)
  else TOBC := nil;
  if TOBC <> nil then
  begin
    for Ind := 0 to TOBC.Detail.Count - 1 do
    begin
      TOBL := TOBC.Detail[Ind];
      dQte := TOBL.GetValue('QUANTITE');
      if dQte > 0 then Result := Result + 'B'
        + IntToStr(TOBL.GetValue('GPI_PIECEBILLET')) + '='
          + StrS(dQte, 0) + ';';
    end;
  end;
  if TOBPiece <> nil then TOBC := TOBPiece.FindFirst([Prefixe + '_MODEPAIE'], [ModePaie], False)
  else TOBC := nil;
  if TOBC <> nil then
  begin
    for Ind := 0 to TOBC.Detail.Count - 1 do
    begin
      TOBL := TOBC.Detail[Ind];
      dQte := TOBL.GetValue('QUANTITE');
      if dQte > 0 then Result := Result + 'P'
        + IntToStr(TOBL.GetValue('GPI_PIECEBILLET')) + '='
          + StrS(dQte, 0) + ';';
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJCtrlMtCais : mise à jours des montants du contrôle caisse
///////////////////////////////////////////////////////////////////////////////////////

function MAJCtrlMtCais(TOBCtr, TOBLig, TOBBillet, TOBPiece: TOB; Caisse: string; NumZ: Integer): Boolean;
var TOBM: TOB;
  dTot, dEcart, dFdc: Double;
  sMdp, sBuf: string;
begin
  Result := False;
  if (TOBCtr = nil) or (TOBLig = nil) then Exit;
  dTot := Arrondi(TOBLig.GetValue('TOTSAISI'), V_PGI.OkDecV);
  dEcart := dTot - TOBLig.GetValue('TOTCALCULE');
  dEcart := Arrondi(dEcart, V_PGI.OkDecV);
  dFdc := Arrondi(TOBLig.GetValue('TOTFDCAIS'), V_PGI.OkDecV);
  if (dTot <> 0) or (dEcart <> 0) or (dFdc <> 0) then
  begin
    sMdp := TOBLig.GetValue('MP_MODEPAIE');
    TOBM := TOB.Create('CTRLCAISMT', TOBCtr, -1);
    TOBM.InitValeurs;
    TOBM.PutValue('GJM_CAISSE', Caisse);
    TOBM.PutValue('GJM_NUMZCAISSE', NumZ);
    TOBM.PutValue('GJM_MODEPAIE', sMdp);
    TOBM.PutValue('GJM_TOTALDEV', dTot);
    TOBM.PutValue('GJM_DEVISE', TOBLig.GetValue('MP_DEVISEFO'));
    TOBM.PutValue('GJM_ECARTDEV', dEcart);
    TOBM.PutValue('GJM_FDCAISSEDEV', dFdc);
    if TOBLig.FieldExists('GJM_FDCAISOUV') then dFdc := TOBLig.GetValue('GJM_FDCAISOUV') else dFdc := 0;
    TOBM.PutValue('GJM_FDCAISOUV', dFdc);
    if TOBLig.FieldExists('GJM_PIECBILOUV') then sBuf := TOBLig.GetValue('GJM_PIECBILOUV') else sBuf := '';
    TOBM.PutValue('GJM_PIECBILOUV', sBuf);
    TOBM.PutValue('GJM_PIECBILTOT', FOMAJTotalPieceBillet(TOBBillet, TOBPiece, sMdp, 'MP'));
    TOBM.SetAllModifie(True);
    Result := True;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOMAJTablesCtrlCais : mise à jours des tables du contrôle caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure FOMAJTablesCtrlCais(TOBPrinc: TOB; Caisse, Vendeur, Titre: string; NumZ: Integer; PieceEcart: boolean);
var TOBEch, TOBBillet, TOBPiece, TOBL, TOBC: TOB;
  Ind: Integer;
  OkMaj: Boolean;
  sEtat: string;
begin
  if (TOBPrinc = nil) or (TOBPrinc.Detail.Count < 1) then Exit;
  TOBEch := TOBPrinc.Detail[0];
  if TOBPrinc.Detail.Count >= 2 then TOBBillet := TOBPrinc.Detail[1] else TOBBillet := nil;
  if TOBPrinc.Detail.Count >= 3 then TOBPiece := TOBPrinc.Detail[2] else TOBPiece := nil;
  // création des pièces d'écarts
  for Ind := 0 to TOBEch.Detail.Count - 1 do
  begin
    TOBL := TOBEch.Detail[Ind];
    if TOBL.FieldExists('CODEETAT') then sEtat := TOBL.GetValue('CODEETAT') else sEtat := '';
    if (PieceEcart) and (sEtat = 'F') then MAJEcart(TOBL, Vendeur, Titre);
  end;
  // mise à jours des montants du contrôle caisse
  OkMaj := False;
  TOBC := TOB.Create('Les montants', nil, -1);
  for Ind := 0 to TOBEch.Detail.Count - 1 do
  begin
    TOBL := TOBEch.Detail[Ind];
    if MAJCtrlMtCais(TOBC, TOBL, TOBBillet, TOBPiece, Caisse, NumZ) then OkMaj := True;
  end;
  if OkMaj then TOBC.InsertOrUpdateDB;
  TOBC.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOSelectCtrlCaisse : Constitue la requête de sélection pour le contrôle caisse
///////////////////////////////////////////////////////////////////////////////////////

function FOSelectCtrlCaisse: string;
begin
  {*****
  Result := 'select GPE_MODEPAIE,GPE_NUMERO,GPE_TIERS,GPE_MONTANTENCAIS,GPE_NATUREPIECEG,GPE_SOUCHE,GPE_INDICEG,'
          + ' GPE_NUMECHE,GPE_DATEPIECE from PIEDECHE'
          + ' where ' + FOMakeWhereTicketX('GPE') + ' and ' + FOFabriqueListeMDP('GPK_MDPCTRLCAIS', 'GPE_MODEPAIE')
          + ' order by GPE_NUMERO' ;
  *****}
  Result := 'SELECT GPE_MODEPAIE,GPE_NUMERO,GPE_TIERS,GPE_MONTANTENCAIS,GPE_NATUREPIECEG,'
    + ' GPE_SOUCHE,GPE_INDICEG,GPE_NUMECHE,GPE_DATEPIECE FROM PIEDECHE'
    + ' LEFT JOIN MODEPAIE ON MP_MODEPAIE=GPE_MODEPAIE'
    + ' LEFT JOIN PIECE ON PIECE.GP_NATUREPIECEG = PIEDECHE.GPE_NATUREPIECEG'
    + '  AND PIECE.GP_SOUCHE = PIEDECHE.GPE_SOUCHE'
    + '  AND PIECE.GP_INDICEG = PIEDECHE.GPE_INDICEG'
    + '  AND PIECE.GP_NUMERO = PIEDECHE.GPE_NUMERO'
    + ' WHERE ' + FOMakeWhereTicketX('GP')
    + FOFabriqueListeMDP('GPK_MDPCTRLCAIS', 'GPE_MODEPAIE')
    + ' AND MP_TYPEMODEPAIE<>"007"'
    + ' ORDER BY GPE_NUMERO';
end;

initialization
  registerclasses([TOF_MFOCTRLECAISSE]);
end.
