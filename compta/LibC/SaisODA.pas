{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... : 12/07/2007
Description .. : Saisie des OD Analityques
Suite ........ : GCO - 02/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Suite ........ : GCO - 06/04/2004
Suite ........ : -> Passage en CWAS ( Suppression des requêtes 
Suite ........ : paramétrées )
Suite ........ : BVE 12.07.07
Suite ........ : Suppression des evenements sur le bouton BValide
Mots clefs ... : 
*****************************************************************}
unit SaisODA;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Hctrls, Grids, StdCtrls, Buttons, Hcompte, Mask, ExtCtrls,
  HEnt1, SaisUtil, SaisComm, Ent1, SaisComp, Filtre,Math {SG6 28/12/04 pour Ceil},
  Menus, hmsgbox, Paramsoc {SG6 22/12/04},
  {$IFDEF EAGLCLIENT}

  {$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  uTob,
  Saisie, HQry, Choix,
  HStatus, HLines,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HSysMenu, HTB97, ed_tools, HPanel, UiUtil, UtilSais, FichComm, UtilPGI,
  TntStdCtrls, TntButtons, TntGrids, TntExtCtrls;

type
  TFSaisODA = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    PEntete: THPanel;
    H_JOURNAL: THLabel;
    H_DATECOMPTABLE: THLabel;
    H_NUMEROPIECE: THLabel;
    H_GENERAL: THLabel;
    H_AXE: THLabel;
    H_LIBELLE: THLabel;
    H_ETABLISSEMENT: THLabel;
    Y_JOURNAL: THValComboBox;
    Y_DATECOMPTABLE: TMaskEdit;
    Y_NUMEROPIECE: THLabel;
    Y_GENERAL: THCpteEdit;
    Y_ETABLISSEMENT: THValComboBox;
    Cache: THCpteEdit;
    PPied: THPanel;
    H_SOLDE: THLabel;
    S_LIBELLE: THLabel;
    GS_LIBELLE: THLabel;
    OA_SOLDES: THNumEdit;
    OA_SOLDEGS: THNumEdit;
    OA_TOTALDEBIT: THNumEdit;
    OA_TOTALCREDIT: THNumEdit;
    OA_SOLDE: THNumEdit;
    H_MONTANTECR: THLabel;
    OA_MONTANTECR: THNumEdit;
    NONCALC: THLabel;
    BInsert: THBitBtn;
    BSDel: THBitBtn;
    GSO: THGrid;
    POPS: TPopupMenu;
    HMessLigne: THMsgBox;
    HMessPiece: THMsgBox;
    HTitres: THMsgBox;
    HDiv: THMsgBox;
    FindSais: TFindDialog;
    Bevel1: TBevel;
    BevCache: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    POPZ: TPopupMenu;
    BZoomPiece: THBitBtn;
    BZoom: THBitBtn;
    BZoomEtabl: THBitBtn;
    HMTrad: THSystemMenu;
    Outils97: TToolbar97;
    BSolde: TToolbarButton97;
    BVentilType: TToolbarButton97;
    BComplement: TToolbarButton97;
    BSoldeGS: TToolbarButton97;
    BChercher: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    HConf: TToolbarButton97;
    LOA_SoldeS: THLabel;
    LOA_SoldeGS: THLabel;
    LOA_TotalDebit: THLabel;
    LOA_TotalCredit: THLabel;
    LOA_MontantEcr: THLabel;
    LOA_Solde: THLabel;
    procedure Y_JOURNALChange(Sender: TObject);
    procedure H_JOURNALDblClick(Sender: TObject);
    procedure Y_GENERALExit(Sender: TObject);
    procedure Y_GENERALChange(Sender: TObject);
    procedure H_GENERALDblClick(Sender: TObject);
    procedure Y_DATECOMPTABLEChange(Sender: TObject);
    procedure Y_DATECOMPTABLEExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure GSOSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
    procedure GSORowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSORowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSOCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSOCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSODblClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BZoomEtablClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BSDelClick(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BSoldeGSClick(Sender: TObject);
    procedure BVentilTypeClick(Sender: TObject);
    procedure GSOEnter(Sender: TObject);
    procedure GSOExit(Sender: TObject);
    procedure GSOKeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure GereAffSolde(Sender: TObject);
  private
    SAJAL: TSAJAL;
    DEV: RDEVISE;
    NumPieceInt, NbLigAna: Longint;
    TAX: HTStrings;
    OI_TotDebit, OI_TotCredit: double;
    OI_TotQ1, OI_TotQ2: double;
    PieceModifiee, Revision: boolean;
    FindFirst, PureAnal: boolean;
    GeneTypeExo: TTypeExo;
    CpteGeneChange: boolean;
    NumLigGene, SensGene: integer;
    TotPGene, TotDGene, TotEGene: Double;
    TS, TSOLDEGS: TList;
    GX, GY, NbLOrig: integer;
    AbregeGen: string;
    NowFutur: TDateTime;
    TDELANA: TList;
    PieceConf: boolean;
    WMinX, WMinY: Integer;
    TEcrODA: Tob;

    //Mode Croisaxe //SG6 22/12/2004
    Axes: array[1..MaxAxe] of Boolean; // Axes ventilables ?
    nbre_axe_ventil : integer; //Nbre d'axes ventilables en mode croisaxe(=1 en mode monoaxe)

    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    // Init et defaut
    procedure PosLesSoldes;
    procedure DefautEntete;
    procedure DefautDebut;
    procedure DefautPied;
    procedure DefautLigne(Lig: integer; Init: boolean);
    procedure InitEnteteJal;
    procedure InitGrid;
    procedure ReInitPiece;
    procedure AttribLeTitre;
    procedure GereEnabled(Lig: integer);
    procedure BloqueDebut(Bloquer: boolean);
    // ALLOCATIONS, SOLDE
    procedure DesalloueA(Lig: Longint);
    procedure SoldelaLigne(Lig: integer);
    function DejaCalcule(Sect: string): boolean;
    // OBJET + DIVERS
    procedure InitOBM(OBM: TOBM; Lig: integer);
    procedure AlimObjetMvt(Lig: integer);
    function FermerSaisie: boolean;
    procedure GereANouveau;
    procedure LectureSeule;
    procedure IncrementeNumero(Provisoire : Boolean);
    procedure CloseFen;
    // CHARGEMENTS
    procedure ChargeLignes;
    procedure ChargeSections(Force: boolean);
    procedure ChargeEcritures;
    procedure ChargeEcrituresCroisaxe;
    procedure ChargeSoldesComptes;
    // VALIDATIONS
    procedure ValideLaPiece;
    procedure ValideLaPieceTrans; //SG6 27/12/2004 Validation des pièces en mode croisaxe
    procedure InverseSolde;
    procedure MiseAJourODA(Lig : integer); //SG6 27/12/2004 Mise à jour tob et base en modif et revision = false
    procedure MAJCptesSection;
    procedure ValideLesSections;
    procedure ValideLeJournal;
    procedure ValideLignesODA;
    procedure RecupTronc(Lig: Integer);
    procedure RecupFromGrid(Lig: integer);
    procedure GetODA(Lig: Integer);
    procedure InsertCroisaxe(Lig : longint);
    procedure GetODAGrid(Lig: Integer);
    procedure DetruitAncien;
    // Traitements ENTETE / PIED
    procedure ChercheJAL(Jal: String3);
    procedure AffichePied;
    procedure AfficheAxe;
    // Traitements LIGNES
    function SommeQte(var QF1, QF2: string): byte;
    procedure DetruitLigne(Lig: integer);
    procedure DetruitLigneCroisaxe(Lig : integer); //SG6 27/12/2004 Gestion en mode croisaxe
    procedure InsereLigne(Lig: integer);
    procedure CalculDebitCredit;
    procedure TraiteMontant(Lig: integer; Calc: boolean);
    procedure ChercheMontant(ACol, ARow: longint);
    procedure ChercheMontantCroisAxe(Acol, ARow : longint ; montant : string); //SG6 22/12/2004 Gestion du mode croisaxe
    procedure CalculSoldeCompte(Sect: string; var TDS, TCS: Double);
    function DateMvt(Lig: integer): TDateTime;
    function Return_Bornesup:longint; //SG6 23/12/2004 Retourne la borne supérieur
    // Barre OUTILS
    procedure ClickValide;
    procedure ClickAbandon;
    procedure ClickComplement;
    procedure ClickZoom;
    procedure ClickZoomPiece;
    procedure ClickCherche;
    procedure ClickEtabl;
    procedure ClickInsert;
    procedure ClickDel;
    procedure ClickSolde(Egal: boolean);
    procedure ClickSoldeGS;
    procedure ClickVentilType;
    // RECHERCHES
    function ChercheSect(L: integer; Force: boolean): byte;
    function RechercheAxe(Lig : integer;var numero_axe : byte):byte; //Sg6 22/12/2004 Fonction qui renvoie l'axe  en fonction de la ligne en cours (en mode croisaxe)
    // CONTROLES
    procedure ErreurSaisie(Err: integer);
    function LigneCorrecte(Lig: integer; Alim: boolean): Boolean;
    function PieceCorrecte: Boolean;
    function PieceVide: Boolean;
    procedure ControleLignes;
    function OkBudget: boolean;
    function OkJal : boolean ;
  public
    Action: TActionFiche;
    M: RMVT;
    {JP 04/08/05 : FQ 15859/15651 : Les BitButtons ne prenant pas le focus, on force la sortie de la cellule}
    procedure LostFocusBeforeClick;

    {JP 07/10/05 : FQ 15881 : Gestion des ventilations types en CroiseAxe}
    procedure VentileTypeCroiseAxe(Code : string; Col : Integer);
    procedure VentileTypeClassique(Code : string; Col : Integer);
    {Remplit les lignes à partir de la ventilation type, en CroisAxe ou non}
    function RemplitLignes(Q : TQuery; var OBM : TOBM; ii : Integer; QF1, QF2 : string) : Double;
    {Calcul et affectation  des montants sur les ventilations types}
    procedure AffecteMontant(var OBM : TOBM; Col, Lig : Integer; XP : Double);

    {Gestion des arrondis, après une ventilation type}
    procedure GereArrondi(Lig, Col : Integer; var OBM : TOBM; XP, TotP, TotM : Double);
    {Termine la pièce}
    procedure FinitPiece(ii : Integer);
    {Renvoie le numéro de ventile en fonction du numéro de ligne dans la grille :
     Si 3 axes de paramétrer en CroisAxe => 3 Lignes dans la grille et 1 lignes dans ANALYTIQ}
    function GetNumLigneCroisAxe(Ligne : Integer) : Integer;
  end;

procedure LanceSaisieODA(Q: TQuery; Action: TActionFiche; M: RMVT);
function TrouveSaisieODA(Q: TQuery; var M: RMVT): boolean;
function TrouveEtLanceSaisieODA(Q: TQuery; TypeAction: TActionFiche): Boolean;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  CPGeneraux_TOM, // FicheGene
  CPJournal_TOM, // FicheJournal
  CPSection_TOM
  {$IFNDEF IMP}
  {$IFNDEF CCMP}
  , SaisBor
  {$ENDIF}
  {$ENDIF}
  ;

procedure LanceSaisieODA(Q: TQuery; Action: TActionFiche; M: RMVT);
var
  X: TFSaisODA;
  PP: THPanel;
begin
  case Action of
    taCreat:
      begin
        if PasCreerDate(V_PGI.DateEntree) then Exit;
        if DepasseLimiteDemo then Exit;
        if _Blocage(['nrCloture'], True, 'nrSaisieCreat') then Exit;
      end;
    taModif:
      begin
        if RevisionActive(M.DateC) then Exit;
        if _Blocage(['nrCloture', 'nrBatch'], True, 'nrSaisieModif') then Exit;
      end;
  end;
  PP := FindInsidePanel;
  X := TFSaisODA.Create(Application);
  X.Action := Action;
  X.M := M;
  if PP = nil then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
      case Action of
        taCreat: _Bloqueur('nrSaisieCreat', False);
        taModif: _Bloqueur('nrSaisieModif', False);
      end;
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

function TrouveSaisieODA(Q: TQuery; var M: RMVT): boolean;
var
  Q1: TQuery;
  Trouv: boolean;
label
  0;
begin
  0: TrouveSaisieODA := False;
  if (Q.EOF) and (Q.Bof) then Exit; //n°1825
  Q1 := OpenSQL('SELECT ' + SQLForIdent(fbSect) + ' FROM ANALYTIQ WHERE Y_JOURNAL="' + Q.FindField('Y_JOURNAL').AsString + '"'
    + ' AND Y_DATECOMPTABLE="' + USDATETIME(Q.FindField('Y_DATECOMPTABLE').AsDateTime) + '"'
    + ' AND Y_NUMEROPIECE=' + Q.FindField('Y_NUMEROPIECE').AsString
    + ' AND Y_NUMLIGNE=' + Q.FindField('Y_NUMLIGNE').AsString
    + ' AND Y_AXE="' + Q.FindField('Y_AXE').AsString + '"'
    + ' AND Y_GENERAL="' + Q.FindField('Y_GENERAL').AsString + '"', True);
  Trouv := not Q1.EOF;
  if Trouv then
    M := MvtToIdent(Q1, fbSect, False);
  Ferme(Q1);
  TrouveSaisieODA := Trouv;
end;

function TrouveEtLanceSaisieODA(Q: TQuery; TypeAction: TActionFiche): Boolean;
var
  M: RMVT;
begin
  Result := TrouveSaisieODA(Q, M);
  if Result then LanceSaisieODA(Q, TypeAction, M);
end;

{=================================== Init et Defauts ==================================}
procedure TFSaisODA.GereEnabled(Lig: integer);
var
  Visu   : Boolean;
  NumAxe : Byte;
begin
  Visu := (Action = taConsult);
  {JP 03/08/05 : FQ 15854 : On cache le zoom si l'on est en train de créer la pièce}
  BZoomPiece.Enabled := (not PureAnal) and (Action <> taCreat);
  BVentilType.Enabled := ((PureAnal) and (OA_Solde.Value <> 0) and (not Visu)); //Sg6 27/12/2004 Gestion croisaxe

  {JP 04/10/05 : FQ 15881 : En cours, en CroiseAxe, le bouton n'est actif que si on est sur la dernière ligne}
  if VH^.AnaCroisaxe then begin
    RechercheAxe(GSO.Row, NumAxe);
    BVentilType.Enabled := BVentilType.Enabled and (NumAxe = nbre_axe_ventil);
  end;

  BComplement.Enabled := EstRempli(GSO, Lig);
  BSoldeGS.Enabled := ((not PieceVide) and (Y_GENERAL.Text <> '') and (not Visu));
  BInsert.Enabled := ((EstRempli(GSO, Lig)) and (not Visu));
  BSDel.Enabled := ((GSO.RowCount > 2) and (not Visu));
end;

procedure TFSaisODA.AttribLeTitre;
var
  i: integer;
begin
  i := 1;
  case Action of
    taConsult:
      begin
        i := 1;
        HelpContext := 7368100;
      end;
    taModif:
      begin
        i := 2;
        HelpContext := 7368200;
      end;
    taCreat:
      begin
        i := 3;
        HelpContext := 7367000;
      end;
  end;
  if M.Simul <> 'N' then i := i + 3;
  if M.Anouveau then
    case Action of
      taConsult:
        begin
          i := 8;
          HelpContext := 7368100;
        end;
      taModif:
        begin
          i := 9;
          HelpContext := 7368200;
        end;
      taCreat:
        begin
          i := 10;
          HelpContext := 7367000;
        end;
    end;
  Caption := HTitres.Mess[i - 1];
  UpdateCaption(Self);
end;

procedure TFSaisODA.BloqueDebut(Bloquer: boolean);
begin
  if Action <> taCreat then Exit;
  BSolde.Enabled := (not Bloquer);
  {JP 04/08/05 : FQ 15881 : La ventiliation type ne doit être active que lorsque l'on vient de saisir une 
                 ligne (cela est géré dans GereEnabled). Par contre ici, que ce soit lors de l'initialisation
                 d'une nouvelle pièce ou au changement de journal, elle doit être désactivée}
  BVentilType.Enabled := False;
  BComplement.Enabled := (not Bloquer);
  BSoldeGS.Enabled := (not Bloquer);
  BChercher.Enabled := (not Bloquer);
  BMenuZoom.Enabled := (not Bloquer);
  BValide.Enabled := (not Bloquer);
end;

procedure TFSaisODA.ReInitPiece;
begin
  GSO.VidePile(True);
  if Action <> taCreat then Exit;

  //SG6 29/12/2004 Gestion en mode croisaxe
  if VH^.AnaCroisaxe then GSO.RowCount := nbre_axe_ventil+1;

  BValide.Enabled := True;
  DefautPied;
  Y_JOURNAL.Enabled := True;
//  Y_JOURNAL.Value := ''; // FQ18743 pas de reinit de la zone
  Y_GENERAL.Text := '' ;   // Reinit du compte
  H_LIBELLE.Caption := '' ;
  Y_GENERALChange(nil);
  PieceModifiee := False;
  Cursor := crDefault;
  Y_JOURNAL.SetFocus;
  PieceConf := False;
  HConf.Visible := False;
  InitEnteteJal;
  BloqueDebut(False);
end;

procedure TFSaisODA.DefautDebut;
begin
  Y_NumeroPiece.Caption := '';
  GSO.Enabled := False;
  if ((not PureAnal) or (not Y_GENERAL.CanFocus) or (Y_GENERAL.Text <> '')) then Y_GENERAL.Enabled := False;
  BloqueDebut(True);
end;

procedure TFSaisODA.DefautEntete;
begin
  Y_JOURNAL.Value := M.JAL;
  Y_DATECOMPTABLE.Text := DateToStr(M.DateC);
  Y_ETABLISSEMENT.Value := M.Etabl;
  Y_ETABLISSEMENT.Enabled := VH^.EtablisCpta;
  if Action = taCreat then PositionneEtabUser(Y_ETABLISSEMENT);
  if Action <> taCreat then SAJAL.Axe := M.Axe;
  Y_GENERAL.Text := M.General;
  Y_GENERAL.ExisteH;
  Y_GENERALExit(nil);
  case Action of
    taCreat: DefautDebut;
    taModif:
      begin
        Y_JOURNAL.Enabled := False;
        Y_DATECOMPTABLE.Enabled := False;
        Y_NumeroPiece.Caption := InttoStr(M.Num);
        NumPieceInt := M.Num;
        GSO.Enabled := True;
        GSO.SetFocus;
        if M.CodeD <> V_PGI.DevisePivot then
        begin
          DEV.Code := M.CodeD;
          GetInfosDevise(DEV);
          DEV.Taux := M.TauxD;
          DEV.DateTaux := M.DateTaux;
        end;
      end;
    taConsult:
      begin
        Y_NumeroPiece.Caption := InttoStr(M.Num);
        NumPieceInt := M.Num;
        GSO.Enabled := True;
        PEntete.Enabled := False;
      end;
  end;
end;

procedure TFSaisODA.DefautPied;
begin
  OI_TotDebit := 0;
  OI_TotCredit := 0;
  ZeroBlanc(PPied);
end;

procedure TFSaisODA.DefautLigne(Lig: integer; Init: boolean);
var
  y, m, d: Word;
  DD: TDateTime;
  S: string;
  i : integer; //SG6 22/12/2004
  numero_axe: byte;   {FP FQ16213}
begin
//SG6 22/12/2004 Mode Croisaxe
  if not(VH^.AnaCroisaxe) then
  begin
    GSO.Cells[OA_NumP, Lig] := IntToStr(NumPieceInt);
    DD := StrToDate(Y_DATECOMPTABLE.Text);
    DecodeDate(DD, y, m, d);
    if Lig > 1 then
    begin
      GSO.Cells[OA_DateC, Lig] := GSO.Cells[OA_DateC, Lig - 1];
      GSO.Cells[OA_RefI, Lig] := GSO.Cells[OA_RefI, Lig - 1];
      GSO.Cells[OA_Lib, Lig] := GSO.Cells[OA_Lib, Lig - 1];
    end else
    begin
      S := IntToStr(d);
      if d < 10 then S := '0' + S;
      GSO.Cells[OA_DateC, Lig] := S;
    end;

    GSO.Cells[OA_NumL, Lig] := IntToStr(Lig);
    GSO.Cells[OA_Exo, Lig] := QuelExoDT(DD);
    AlloueMvt(GSO, EcrAna, Lig, Init);
  end
  else
  begin
   for i := Lig to Lig + nbre_axe_ventil - 1 do
   begin
     GSO.Cells[OA_NumP, i] := IntToStr(NumPieceInt);
     DD := StrToDate(Y_DATECOMPTABLE.Text);
     DecodeDate(DD, y, m, d);
     {if Lig > 1 then
     begin
       GSO.Cells[OA_DateC, Lig] := GSO.Cells[OA_DateC, Lig - 1];
       GSO.Cells[OA_RefI, Lig] := GSO.Cells[OA_RefI, Lig - 1];
       GSO.Cells[OA_Lib, Lig] := GSO.Cells[OA_Lib, Lig - 1];
     end else
     begin}
     S := IntToStr(d);
     if d < 10 then S := '0' + S;
     GSO.Cells[OA_DateC, i] := S;
     {end;}
     GSO.Cells[OA_NumL, Lig] := IntToStr(GetNumLigneCroisAxe(Lig));//IntToStr(Lig);
     GSO.Cells[OA_Axe, Lig]  := 'A'+IntToStr(RechercheAxe(Lig,numero_axe));   {FP FQ16213}
     GSO.Cells[OA_Exo, i] := QuelExoDT(DD);
     AlloueMvt(GSO, EcrAna, i, Init);
   end;
  end;

end;

procedure TFSaisODA.IncrementeNumero(Provisoire : Boolean);
var
  Facturier : String3;
  MM        : String17; {FQ 16848 }
  DD        : TDateTime;
begin
  if M.Simul <> 'N' then Facturier := SAJAL.COMPTEURSIMUL
                    else Facturier := SAJAL.COMPTEURNORMAL;
  DD := StrToDate(Y_DATECOMPTABLE.Text);
  if Action <> taCreat then Exit;
  {JP 10/10/05 : FQ 16848 : Attribution d'un numéro provisoire en début d'écriture}
  if Provisoire then NumPieceInt := GetNum(EcrAna, Facturier, MM, DD)
                else SetIncNum(EcrAna, Facturier, NumPieceInt, DD);
end;

procedure TFSaisODA.InitEnteteJal;
begin
  if Action <> taCreat then Exit;
  if SAJAL = nil then Exit;
  {JP 10/10/05 : FQ 16848 : Lors de l'initialisation de l'entête, le numéro est provisoire
   if Transactions(IncrementeNumero, 10) <> oeOK then MessageAlerte(HDiv.Mess[2]);}
  IncrementeNumero(True);
  if NumPieceInt > 0 then Y_NUMEROPIECE.Caption := IntToStr(NumPieceInt);
  InitGrid;
end;

procedure TFSaisODA.InitGrid;

begin
  if Action <> taCreat then Exit;
  GSO.Enabled := True;
//  DefautLigne(GSO.RowCount - 1, True); //SG6 22/12/2004
  DefautLigne(1,True);
  GSO.Col := OA_Sect;
  GSO.Row := 1;
  if ((PureAnal) and (Y_GENERAL.CanFocus) and (Y_GENERAL.Text = '')) then Y_GENERAL.SetFocus
  else
  begin
    //SG6 25/01/05 FQ 15322
    if Y_DATECOMPTABLE.CanFocus then Y_DATECOMPTABLE.SetFocus; // Fiche 5601 GSO.SetFocus ;
  end;
end;

{================================== Traitements LIGNES ===================================}
function TFSaisODA.SommeQte(var QF1, QF2: string): byte;
var
  i, i1, i2, borne_sup: integer; //SG6 23/12/2004 Mode Croisaxe
  OF1, OF2: string;
  O: TOBM;
begin
  OF1 := '';
  OF2 := '';
  QF1 := '';
  QF2 := '';
  i1 := 0;
  i2 := 0;
  OI_TotQ1 := 0;
  OI_TotQ2 := 0;
 //SG6 23/21/2004 Gestion du mode croisaxe
  borne_sup:=Return_Bornesup;
  for i := 1 to borne_sup do
  begin
    O := GetO(GSO, i);
    if O = nil then Continue;
    QF1 := O.GetMvt('Y_QUALIFQTE1');
    QF2 := O.GetMvt('Y_QUALIFQTE2');
    if O.GetMvt('Y_DEBIT') <> 0 then
    begin
      OI_TotQ1 := OI_TotQ1 + O.GetMvt('Y_QTE1');
      OI_TotQ2 := OI_TotQ2 + O.GetMvt('Y_QTE2');
    end else
    begin
      OI_TotQ1 := OI_TotQ1 - O.GetMvt('Y_QTE1');
      OI_TotQ2 := OI_TotQ2 - O.GetMvt('Y_QTE2');
    end;
    if ((OF1 <> '') and (OF1 <> QF1) and (i1 <= 0)) then i1 := 1;
    if ((OF2 <> '') and (OF2 <> QF2) and (i2 <= 0)) then i2 := 1;
    OF1 := QF1;
    OF2 := QF2;
  end;
  if OI_TotQ1 = 0 then i1 := 0;
  if OI_TotQ2 = 0 then i2 := 0;
  Result := i1 + i2;
end;

procedure TFSaisODA.DetruitLigne(Lig: integer);
var
  R: integer;
begin
  GSO.CacheEdit;
  R := Lig;
  GSO.DeleteRow(Lig);
  CalculDebitCredit;

  NumeroteLignes(GSO);

  GSO.SetFocus;
  if R >= GSO.RowCount then R := GSO.RowCount - 1;
  if R > 1 then GSO.Row := R else GSO.Row := 1;
  GSO.Col := OA_Sect;
  GereNewLigne(GSO);
  GSO.Invalidate;
  GSO.MontreEdit;
end;

procedure TFSaisODA.InsereLigne(Lig: integer);
var
  i : integer;
  numero_axe : byte;
begin
  if Action = taConsult then Exit;
  if Lig < GSO.RowCount - 1 then
  begin
    //sg6 29/12/2004 Gestion en mode croisaxe (en mode monoaxe nbre_axe_ventil = 1)
    for i:=1 to nbre_axe_ventil do
    begin
      GSO.InsertRow(Lig);
      DefautLigne(Lig, True);
    end;
  end;

  //SG6 28/12/2004 Gestion en mode csoisaxe
  if not VH^.AnaCroisaxe then
  begin
    NumeroteLignes(GSO);
  end
  else
  begin
    for i:=1 to GSO.RowCount - 1 do
    begin
      GSO.Cells[OA_NumL,i] := IntToStr(GetNumLigneCroisAxe(i)); //IntToStr(i);
      {b FP FQ16213}
      GSO.Cells[OA_Axe, i]  := 'A'+IntToStr(RechercheAxe(i,numero_axe));
      //RechercheAxe(i,numero_axe);
      {e FP FQ16213}
      if numero_axe = 0 then exit;
      if GetO(GSO,i)<>nil then TOBM(GetO(GSO,i)).PutMvt('Y_NUMVENTIL',IntToStr(GetNumLigneCroisAxe(i - numero_axe + 1)))
    end;
  end;

  AffichePied;
end;

function TFSaisODA.DateMvt(Lig: integer): TDateTime;
begin
  Result := StrToDate(Y_Datecomptable.Text);
end;

procedure TFSaisODA.CalculSoldeCompte(Sect: string; var TDS, TCS: Double);
var
  i: integer;
begin
  TDS := 0;
  TCS := 0;
  for i := 1 to GSO.RowCount - 1 do
  begin
    if ((Sect <> '') and (GSO.Cells[OA_Sect, i] = Sect)) then
    begin
      TDS := TDS + ValD(GSO, i) - GetO(GSO, i).GetMvt('OLDDEBIT');
      TCS := TCS + ValC(GSO, i) - GetO(GSO, i).GetMvt('OLDCREDIT');
    end;
  end;
end;

procedure TFSaisODA.TraiteMontant(Lig: integer; Calc: boolean);
var
  XC, XD, SD, SC: Double;
  OBM: TOBM;
  TotEcr, Pourc: Double;
begin
  OBM := GetO(GSO, Lig);
  if OBM = nil then Exit;
  XD := ValD(GSO, Lig);
  XC := ValC(GSO, Lig);
  SD := OBM.GetMvt('Y_DEBITDEV');
  SC := OBM.GetMvt('Y_CREDITDEV');
  if ((SD = XD) and (SC = XC) and (PureAnal)) then Exit;
  OBM.PutMvt('Y_DEBITDEV', XD);
  OBM.PutMvt('Y_CREDITDEV', XC);
  OBM.PutMvt('Y_DEBIT', XD);
  OBM.PutMvt('Y_CREDIT', XC);
  TotEcr := OBM.GetMvt('Y_TOTALECRITURE');
  if ((TotEcr <> 0) and (not PureAnal)) then
  begin
    Pourc := Arrondi(100.0 * (XC + XD) / TotEcr, ADecimP);
    OBM.PutMvt('Y_POURCENTAGE', Pourc);
  end;
  if Calc then CalculDebitCredit;
end;

procedure TFSaisODA.ChercheMontant(ACol, ARow: longint);
begin
  if ACol = OA_Debit then
  begin
    if ValD(GSO, ARow) <> 0 then GSO.Cells[OA_Credit, ARow] := '' else GSO.Cells[ACol, ARow] := '';
  end
  else
  begin
    if ValC(GSO, ARow) <> 0 then GSO.Cells[OA_Debit, ARow] := '' else GSO.Cells[ACol, ARow] := '';
  end;
  FormatMontant(GSO, ACol, ARow, V_PGI.OkDecV);
  TraiteMontant(ARow, True);
end;

procedure TFSaisODA.CalculDebitCredit;
var
  i: integer;
  ModeConf: string[1];
  CSect: TGSection;
  O: TOBM;
begin
  OI_TotDebit := 0;
  OI_TotCredit := 0;
  ModeConf := '0';

  //SG6 28/12/2004 Gestion croisaxe
  i := 1;
  while (i<GSO.RowCount-nbre_axe_ventil+1) do
  begin
    O := GetO(GSO, i);
    if O = nil then
    begin
      Inc(i,nbre_axe_ventil);
      Continue;
    end;
    OI_TotDebit := OI_TotDebit + ValD(GSO, i);
    OI_TotCredit := OI_TotCredit + ValC(GSO, i);
    CSect := TGSection(GSO.Objects[OA_Sect, i]);
    if ((CSect <> nil) and (CSect.Confidentiel > ModeConf)) then ModeConf := CSect.Confidentiel;
    Inc(i,nbre_axe_ventil);
  end;

  OA_TotalDebit.Value := OI_TotDebit;
  OA_TotalCredit.Value := OI_TotCredit;
  AfficheLeSolde(OA_Solde, OI_TotDebit, OI_TotCredit);
  if ((Action <> taConsult) and (PureAnal)) then HConf.Visible := (ModeConf > '0');
end;

{================================== ENTETE / PIED ====================================}
procedure TFSaisODA.AffichePied;
var
  CSect: TGSection;
  Sect: string;
  TDS, TCS, DIS, CIS, GSD, GSC: double;
  T: T_SC;
  i: integer;
  Okok, OkVoir: boolean;
begin
  Sect := '';
  DIS := 0;
  CIS := 0;
  GSD := 0;
  GSC := 0;
  S_Libelle.Caption := '';
  GS_Libelle.Caption := '';
  Okok := False;
  OkVoir := False;
  if GSO.Objects[OA_Sect, GSO.Row] <> nil then
  begin
    CSect := TGSection(GSO.Objects[OA_Sect, GSO.Row]);
    OkVoir := True;
    Sect := CSect.Sect;
    DIS := CSect.TotalDebit;
    CIS := CSect.TotalCredit;
    S_Libelle.Caption := CSect.Libelle;
    if Y_GENERAL.Text <> '' then
    begin
      GS_Libelle.Caption := AbregeGen + ' / ' + CSect.Abrege;
      for i := 0 to TSoldeGS.Count - 1 do
      begin
        T := T_SC(TSoldeGS[i]);
        if T.Cpte = Sect then
        begin
          GSD := T.Debit;
          GSC := T.Credit;
          Okok := True;
          Break;
        end;
      end;
    end;
  end;
  CalculSoldeCompte(Sect, TDS, TCS);
  AfficheLeSolde(OA_SoldeS, TDS + DIS, TCS + CIS);
  if Okok then
  begin
    NonCalc.Visible := False;
    LOA_SoldeGS.Visible := TRUE;
    AfficheLeSolde(OA_SoldeGS, TDS + GSD, TCS + GSC);
  end
  else
  begin
    LOA_SoldeGS.Visible := False;
    NonCalc.Visible := True;
  end;
  if not OkVoir then
  begin
    LOA_SOLDES.Visible := False;
    LOA_SOLDEGS.Visible := False;
    NonCalc.Visible := False;
  end else
  begin
    LOA_SOLDES.Visible := True;
  end;
end;

procedure TFSaisODA.AfficheAxe;
var
  Q: TQuery;
  NumA: integer;
  i : integer;
begin
  if SAJAL.Journal = '' then Exit;

  //SG6 22/12/04 Mode Croisaxe Axe fonction non plus du journal mais d'ou l'on se trouve dans le grid
  if VH^.AnaCroisaxe then
  begin
    H_AXE.Caption := 'Axe ';
    for i:=1 to MaxAxe do
    begin
      if Axes[i] then H_AXE.Caption := H_AXE.Caption + IntToStr(i) + ', ';
    end;
    H_AXE.Caption := Copy(H_AXE.Caption,1,Length(H_AXE.Caption)-2);
  end
  else
  begin
    H_AXE.Caption := '';
    if SAJAL.Axe <> '' then
    begin
      Q := OpenSQL('Select X_LIBELLE from AXE Where X_AXE="' + SAJAL.Axe + '"', True);
      if not Q.EOF then H_AXE.Caption := Q.Fields[0].AsString;
      Ferme(Q);
      NumA := Ord(SAJAL.Axe[2]) - 48;
    end else
    begin
      NumA := 0;
    end;
    if NumA > 0 then PositionneTz(Cache, NumA) else Cache.ZoomTable := AxeToTz(M.Axe);
    if Action = taCreat then
    begin
      case NumA of
        1: Y_GENERAL.ZoomTable := tzGVentil1;
        2: Y_GENERAL.ZoomTable := tzGVentil2;
        3: Y_GENERAL.ZoomTable := tzGVentil3;
        4: Y_GENERAL.ZoomTable := tzGVentil4;
        5: Y_GENERAL.ZoomTable := tzGVentil5;
      end;
    end;
  end;
end;

procedure TFSaisODA.ChercheJAL(Jal: String3);
begin
  SAJAL.Free;
  SAJAL := TSAJAL.Create(Jal, FALSE);
  if Action = taCreat then InitEnteteJal;
  PieceModifiee := True;
end;

procedure TFSaisODA.Y_DATECOMPTABLEChange(Sender: TObject);
begin
  PieceModifiee := True;
end;

procedure TFSaisODA.Y_DATECOMPTABLEExit(Sender: TObject);
var
  Err: integer;
  Lig: integer;
  OBM: TOBM;
  y, m, d: Word;
  DD, LaDate: TDateTime;
  StD: string;
begin
  if csDestroying in ComponentState then Exit;
  Err := ControleDate(Y_DATECOMPTABLE.EditText);
  if Err > 0 then
  begin
    HMessPiece.Execute(5 + Err, caption, '');
    Y_DATECOMPTABLE.SetFocus;
    Exit;
  end;
  if ((Action = taCreat) and (RevisionActive(StrToDate(Y_DATECOMPTABLE.Text)))) then
  begin
    Y_DATECOMPTABLE.SetFocus;
    Exit;
  end;
  DD := StrToDate(Y_DATECOMPTABLE.Text);
  DecodeDate(DD, y, m, d);
  StD := IntToStr(d);
  if d < 10 then StD := '0' + StD;
  for Lig := 1 to GSO.RowCount - 1 do
  begin
    GSO.Cells[OA_DateC, Lig] := StD;
    OBM := GetO(GSO, Lig);
    if OBM <> nil then
    begin
      LaDate := StrToDate(Y_DATECOMPTABLE.Text);
      OBM.PutMvt('Y_DATECOMPTABLE', LaDate);
      {$IFNDEF SPEC302}
      OBM.PutMvt('Y_PERIODE', GetPeriode(LaDate));
      OBM.PutMvt('Y_SEMAINE', NumSemaine(LaDate));
      {$ENDIF}
    end;
  end;
end;

procedure TFSaisODA.H_JOURNALDblClick(Sender: TObject);
var
  a: TActionFiche;
begin
  if Action = taConsult then a := taConsult else a := taModif;
  if not ExJaiLeDroitConcept(TConcept(ccJalModif), False) then A := taConsult;
  FicheJournal(nil, '', Y_JOURNAL.Value, a, 0);
  if a <> taConsult then ChercheJal(Y_JOURNAL.Value);
end;

procedure TFSaisODA.Y_JOURNALChange(Sender: TObject);
var
  Jal: string;
begin
  Jal := Y_JOURNAL.Value;
  if Jal = '' then
  begin
    SAJAL.Free;
    SAJAL := nil;
    Exit;
  end;
  if SAJAL <> nil then if SAJAL.JOURNAL = Jal then Exit;
  ChercheJal(Jal);
  if ((Action = taCreat) and (NumPieceInt <= 0)) then
  begin
    if SAJAL <> nil then
    begin
      SAJAL.Free;
      SAJAL := nil;
    end;
    Y_JOURNAL.Value := '';
    Y_JOURNAL.SetFocus;
    GSO.Enabled := False;
    HMessPiece.Execute(13, caption, '');
    Exit;
  end;
  if ((VH^.JalAutorises <> '') and (Pos(';' + Jal + ';', VH^.JalAutorises) <= 0) and (Action = taCreat)) then
  begin
    if SAJAL <> nil then
    begin
      SAJAL.Free;
      SAJAL := nil;
    end;
    Y_JOURNAL.Value := '';
    Y_JOURNAL.SetFocus;
    GSO.Enabled := False;
    HMessPiece.Execute(19, caption, '');
    Exit;
  end;
  AfficheAxe;
  if Action = taCreat then
  begin
    Y_GENERAL.Enabled := True;
    BloqueDebut(False);
  end;
end;

procedure TFSaisODA.Y_GENERALChange(Sender: TObject);
begin
  PieceModifiee := True;
  Revision := False;
  CpteGeneChange := True;
end;

procedure TFSaisODA.H_GENERALDblClick(Sender: TObject);
var
  A: TActionFiche;
begin
  if Y_GENERAL.Text = '' then Exit;
  Y_GENERALExit(nil);
  if CpteGeneChange then Exit;
  //if Not JaiLeDroitConcept(TConcept(ccGenModif),False) then A:=taConsult ;
  //if Action=taConsult then A:=taConsult else A:=taModif ;
  if not ExJaiLeDroitConcept(TConcept(ccGenModif), False) then A := taConsult
  else if Action = taConsult then A := taConsult else A := taModif;
  FicheGene(nil, '', Y_GENERAL.Text, A, 0);
end;

procedure TFSaisODA.Y_GENERALExit(Sender: TObject);
var
  Q: TQuery;
  OldA: string;
  ia: integer;
  Compat: boolean;
begin
  if csDestroying in ComponentState then Exit;
  if not CpteGeneChange then Exit;
  AbregeGen := '';
  if Y_GENERAL.Text = '' then
    begin
    if not Assigned(SAJAL) then
      Y_JOURNAL.SetFocus ;
    Exit;
    end ;
  Y_GENERAL.Text := BourreLaDonc(Y_GENERAL.Text, fbGene);
  Q := OpenSQL('Select G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, G_ABREGE, G_CONFIDENTIEL, G_LIBELLE '
    + 'from GENERAUX where G_GENERAL="' + Y_GENERAL.Text + '"', True);
  if Q.EOF then
  begin
    HMessPiece.Execute(3, caption, '');
    Y_GENERAL.SetFocus;
  end else
  if Assigned(SAJAL) then // FQ 18743 SBO 07/09/2006 : Pb si journal non sélectionné
  begin
    OldA := SAJAL.Axe;
    ia := Ord(OldA[2]) - 48;
    Compat := Q.FindField('G_VENTILABLE' + IntToStr(ia)).AsString = 'X';
    if not Compat then
    begin
      HMessPiece.Execute(2, caption, '');
      Y_GENERAL.SetFocus;
    end else if EstConfidentiel(Q.FindField('G_CONFIDENTIEL').AsString) then
    begin
      HMessPiece.Execute(12, caption, '');
      Y_GENERAL.Text := '';
      Y_GENERAL.SetFocus;
    end else
    begin
      AbregeGen := Q.FindField('G_ABREGE').AsString;
      H_LIBELLE.Caption := Q.FindField('G_LIBELLE').AsString ;
      CpteGeneChange := False;
    end;
  end
  else
    begin
    // FQ 18743 SBO 07/09/2006 : Pb si journal non sélectionné
    PGIError('Vous devez sélectionner un journal.', Caption ) ;
    Y_GENERAL.Text := '' ;
    H_LIBELLE.Caption := '' ;
    Y_JOURNAL.SetFocus ;
    end;
  Ferme(Q);
end;

{================================= OBJET / DIVERS =========================================}
procedure TFSaisODA.GereANouveau;
begin
  if Action <> taCreat then Exit;
  if not M.ANouveau then Exit;
  Y_DATECOMPTABLE.EditText := DateToStr(VH^.Encours.Deb);
  if VH^.Entree.Code = VH^.Suivant.Code then Y_DATECOMPTABLE.EditText := DateToStr(VH^.Suivant.Deb);
  Y_DATECOMPTABLE.Enabled := False;
end;

procedure TFSaisODA.LectureSeule;
begin
  BInsert.Enabled := False;
  BSDel.Enabled := False;
  BSolde.Enabled := False;
end;

procedure TFSaisODA.InitOBM(OBM: TOBM; Lig: integer);
var
  LaDate: TDateTime;
begin
  if OBM = nil then Exit;
  LaDate := StrToDate(Y_DATECOMPTABLE.Text);
  OBM.PutMvt('Y_GENERAL', Y_GENERAL.Text);
  OBM.PutMvt('Y_EXERCICE', QuelExo(Y_DATECOMPTABLE.Text));
  OBM.PutMvt('Y_NUMEROPIECE', NumPieceInt);
  OBM.PutMvt('Y_DATECOMPTABLE', LaDate);
  {$IFNDEF SPEC302}
  OBM.PutMvt('Y_PERIODE', GetPeriode(LaDate));
  OBM.PutMvt('Y_SEMAINE', NumSemaine(LaDate));
  {$ENDIF}
  OBM.PutMvt('Y_NUMLIGNE', NumLigGene);
  OBM.PutMvt('Y_UTILISATEUR', V_PGI.User);
  OBM.PutMvt('Y_SOCIETE', V_PGI.CodeSociete);
  OBM.PutMvt('Y_ETABLISSEMENT', Y_Etablissement.Value);
  OBM.PutMvt('Y_TOTALECRITURE', TotPGene);
  OBM.PutMvt('Y_TOTALDEVISE', TotDGene);
  OBM.PutMvt('Y_JOURNAL', Y_JOURNAL.Value);
  OBM.PutMvt('Y_QUALIFPIECE', '');
  OBM.PutMvt('Y_TYPEMVT', 'AE');
  OBM.PutMvt('Y_REFINTERNE', Copy(GSO.Cells[OA_RefI, Lig], 1, 35));
  OBM.PutMvt('Y_LIBELLE', Copy(GSO.Cells[OA_Lib, Lig], 1, 35));
  if ((not PureAnal) and (TotPGene <> 0)) then
  begin
    ChercheMontant(OA_Debit, Lig);
    ChercheMontant(OA_Credit, Lig);
  end;
end;

procedure TFSaisODA.AlimObjetMvt(Lig: integer);
var
  OBM: TOBM;
  numventil : integer;
begin
  OBM := GetO(GSO, Lig);
  if OBM = nil then
  begin
    AlloueMvt(GSO, EcrAna, Lig, True);
    OBM := GetO(GSO, Lig);
  end;
  InitOBM(OBM, Lig);
  //SG6 28/12/2004 Gestion du champ Y_NUMVENTIL cn mode croisaxe
  if VH^.AnaCroisaxe then
  begin
    numventil := GetNumLigneCroisAxe(Lig);//Ceil(Lig/nbre_axe_ventil);
    OBM.PutMvt('Y_NUMVENTIL',numventil);
  end
  else
  begin
    OBM.PutMvt('Y_NUMVENTIL',Lig);
  end;

  OBM.PutMvt('Y_AXE', SAJAL.Axe);
  OBM.PutMvt('Y_SECTION', GSO.Cells[OA_Sect, Lig]);
  OBM.MajLesDates(Action);
end;

function TFSaisODA.FermerSaisie: boolean;
begin
  Result := True;
  if ((Action = taConsult) or (not PieceModifiee)) then Exit;
  if HMessPiece.Execute(1, caption, '') <> mrYes then Result := False;
end;

{=================================== CONTROLES ========================================}
procedure TFSaisODA.ControleLignes;
var
  Lig: integer;
begin
  for Lig := 1 to GSO.RowCount - 1 do if not EstRempli(GSO, Lig) then DefautLigne(Lig, True);
end;

function TFSaisODA.PieceVide: boolean;
begin
  PieceVide := ((GSO.RowCount <= 2) and (not EstRempli(GSO, 1)));
end;

procedure TFSaisODA.ErreurSaisie(Err: integer);
begin
  if Err < 100 then
  begin
    HMessLigne.Execute(Err - 1, caption, '');
  end else if Err < 200 then
  begin
    HMessPiece.Execute(Err - 101, caption, '');
  end;
end;

function TFSaisODA.LigneCorrecte(Lig: integer; Alim: boolean): Boolean;
var
  CSect: TGSection;
  Err, Col: integer;
begin
  if Action = taConsult then
  begin
    Result := True;
    Exit;
  end;
  Err := 0;
  Col := -1;
  CSect := GetGSect(GSO, Lig);
  if ((Err = 0) and (CSect = nil)) then
  begin
    Err := 1;
    Col := OA_Sect;
  end;
  if ((Err = 0) and (CSect.Sect <> GSO.Cells[OA_Sect, Lig])) then
  begin
    Err := 1;
    Col := OA_Sect;
  end;
  if ((Err = 0) and (ValD(GSO, Lig) = 0) and (ValC(GSO, Lig) = 0)) then
  begin
    Err := 2;
    if SensGene = 2 then Col := OA_Credit else Col := OA_Debit;
  end;
  Result := (Err = 0);
  if not Result then
  begin
    ErreurSaisie(Err);
    if Col > 0 then
    begin
      GSO.Col := Col;
      if GSO.CanFocus then GSO.SetFocus;
    end;
  end else
  begin
  if Alim then AlimObjetMvt(Lig);
  end;
end;

function TFSaisODA.PieceCorrecte: Boolean;
var
  i, Err, borne_sup: integer; //SG6 23/12/2004 Mode Croisaxe
begin
  if Action = taConsult then
  begin
    Result := True;
    exit;
  end;
  Err := 0;
  if ((Err = 0) and (GSO.RowCount < 3)) then Err := 101;
  if ((Err = 0) and (PureAnal)) then
  begin
    if Y_GENERAL.Text = '' then Err := 116 else
      if Y_GENERAL.ExisteH <= 0 then Err := 116;
  end;
  if Err = 0 then
  begin
    if SensGene = 1 then
    begin
      if not PureAnal then
      begin
        if Arrondi(OI_TotDebit - OI_TotCredit - TotPGene, V_PGI.OkDecV) <> 0 then Err := 106;
      end else if ((VH^.EquilAnaODA) and (not M.ANouveau)) then
      begin
        if Arrondi(OI_TotDebit - OI_TotCredit, V_PGI.OkDecV) <> 0 then Err := 122;
      end;
    end else
    begin
      if not PureAnal then
      begin
        if Arrondi(OI_TotCredit - OI_TotDebit - TotPGene, V_PGI.OkDecV) <> 0 then Err := 106;
      end else if ((VH^.EquilAnaODA) and (not M.ANouveau)) then
      begin
        if Arrondi(OI_TotDebit - OI_TotCredit, V_PGI.OkDecV) <> 0 then Err := 122;
      end;
    end;
  end;
  //SG6 23/12/2004 Mode Croisaxe
  borne_sup:=Return_Bornesup;
  if Err = 0 then for i := 1 to borne_sup do
    begin
      if not LigneCorrecte(i, False) then
      begin
        Err := 1000;
        Break;
      end;
    end;
  Result := (Err = 0);
  if not Result then
    begin
    ErreurSaisie(Err);
    if Err=116 then Y_GENERAL.Setfocus ;
    end ;
end;

{=================================== Barre OUTILS ========================================}
procedure TFSaisODA.ClickVentilType;
var
  Code : string;
  Col  : Integer;
begin
  if Action = taConsult then Exit;
  if not PureAnal then Exit;
  if not BVentilType.Enabled then Exit;

  if OA_Solde.Value = 0 then Exit;

  {Gestion des quantités}
  if OA_Solde.Debit then
    Col := OA_Credit
  else begin
    Col := OA_Debit;
    OI_TotQ1 := -OI_TotQ1;
    OI_TotQ2 := -OI_TotQ2;
  end;

  Code := Choisir(HTitres.Mess[6], 'CHOIXCOD', 'CC_LIBELLE', 'CC_CODE', 'CC_TYPE="VTY"', '');
  if Code = '' then Exit;

  {JP 04/10/05 : FQ 15881 : Gestion du CroiseAxe sur les ventilations types}
  if VH^.AnaCroisaxe then VentileTypeCroiseAxe(Code, Col)
                     else VentileTypeClassique(Code, Col);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 18/08/2004
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices
Suite ........ : decalees
Mots clefs ... :
*****************************************************************}
procedure TFSaisODA.ClickZoomPiece;
var
  Q: TQuery;
  Trouv: boolean;
  St: string;
  ME: RMVT;
  {$IFNDEF IMP}
  {$IFNDEF CCMP}
  P: RParFolio;
  {$ENDIF}
  {$ENDIF}
begin
  if PureAnal then Exit;
  St := 'SELECT ' + SQLForIdent( fbGene ) + ' FROM ECRITURE WHERE ' + WhereEcriture(tsGene, M, False);
  Q := OpenSQL(St, True);
  Trouv := not Q.EOF;
  if Trouv then
  begin
    ME := MvtToIdent(Q, fbGene, False);
    ME.General := Y_GENERAL.Text;
    ME.AXE := SAJAL.Axe;
  end;
  Ferme(Q);
  if Trouv then
  begin
    if ((ME.ModeSaisieJal <> '-') and (ME.ModeSaisieJal <> '')) then
    begin
      {$IFNDEF IMP}
      {$IFNDEF CCMP}
      FillChar(P, Sizeof(P), #0);
      P.ParPeriode := DateToStr(ME.DateC);
      P.ParCodeJal := ME.Jal;
      P.ParNumFolio := IntToStr(ME.Num);
      P.ParNumLigne := ME.NumLigne;
      ChargeSaisieFolio(P, taConsult);
      {$ENDIF}
      {$ENDIF}
    end else
    begin
      LanceSaisie(nil, taConsult, ME);
    end;
  end;
end;

procedure TFSaisODA.ClickZoom;
var
  C, R: Longint;
  A: TActionFiche;
  iAxe,numero_axe : byte;
  CSect : TGSection ;
begin
  if Action = taConsult then
  begin
    R := GSO.Row;
    C := OA_Sect;
    A := taConsult;
  end else
  begin
    R := GSO.Row;
    C := GSO.Col;
    A := taModif;
  end;
  if R < 1 then Exit;
  if C <> OA_Sect then Exit;
  if ((Action = taConsult) and (GSO.Cells[C, R] = '')) then Exit;
  if not ExJaiLeDroitConcept(TConcept(ccSecModif), False) then A := taConsult;

  //SG6 22/12/2004 On determine sur quel axe on se trouve
  iAxe := 0;
  if VH^.AnaCroisaxe then
  begin
    iAxe := RechercheAxe(GSO.Row,numero_axe);
    if iAxe = 0 then exit;
    Cache.ZoomTable := AxeToTz('A'+IntToStr(iAxe));
  end;

  // Si la compte est déja chargé, on ne refait pas la recherche en base...
  CSect := TGSection(GSO.Objects[OA_Sect, R]);
  if (CSect <> nil) and ( GSO.Cells[OA_Sect, GSO.Row] = CSect.Sect ) then
    begin
    //SG6 27/12/2004 Gestion du mode croisaxe
    if not VH^.AnaCroisaxe
      then FicheSection(nil, SAJAL.Axe, GSO.Cells[OA_Sect, GSO.Row], A, 0)
      else FicheSection(nil,'A'+IntToStr(iAxe), GSO.Cells[OA_Sect, GSO.Row], A, 0);
    end
  else if Action <> taConsult then
         ChercheSect(R, True);
end;

procedure TFSaisODA.ClickCherche;
begin
  if not BChercher.Enabled then Exit;
  if GSO.RowCount <= 1 then Exit;
  FindFirst := True;
  FindSais.Execute;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lek
Créé le ...... : 23/09/2005
Modifié le ... :   /  /
Description .. : La varible M.Axe n'est pas alimentée.
Suite ........ : Je l'ai remplacé par SAJAL.Axe
Mots clefs ... :
*****************************************************************}
function TFSaisODA.OkBudget: boolean;
var
  Q: TQuery;
  Nb, i, borne_sup: integer;   
  Trouv, OkAx: boolean;
  Sect: string;
begin
  Result := True;
  Nb := 0;
  Trouv := True;
  OkAx := True;
  if Action = taConsult then Exit;
  {si pas de jalbud pour ctrl dans Bo --> exit}
  if VH^.JalCtrlBud = '' then Exit;
  {si pas de général --> exit}
  if Y_GENERAL.Text = '' then Exit;
  {si pas de croisements budgets définis --> exit}
  Q := OpenSQL('Select Count(*) from CROISCPT where CX_TYPE="BUD" AND CX_JAL="' + VH^.JalCtrlBud + '"', True);
  if not Q.EOF then Nb := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nb <= 0 then Exit;
  {si axe de l'écriture <> Axe du jal de BO --> exit}
  Q := OpenSQL('Select * from BUDJAL Where BJ_BUDJAL="'
       + VH^.JalCtrlBud + '" AND BJ_AXE="' +
       {M.Axe} SAJAL.Axe + '"', True); {Lek 23/09/05 FQ16571}
  if Q.EOF then OkAx := False;
  Ferme(Q);
  if not OkAx then Exit;

  //SG6 23/12/2004 Gestion croisaxe
  borne_sup:=Return_Bornesup;
  for i := 1 to borne_sup do
  begin
    Sect := GSO.Cells[OA_Sect, i];
    if Sect = '' then Break;
    Q := OpenSQL('Select CX_TYPE from CROISCPT Where CX_TYPE="GEN" AND CX_JAL="' + VH^.JalCtrlBud + '" AND CX_COMPTE="' + Y_GENERAL.Text + '" AND CX_SECTION="' + Sect + '"',
      True);
    if Q.EOF then Trouv := False;
    Ferme(Q);
    if not Trouv then Break;
  end;
  if not Trouv then
  begin
    if HMessPiece.Execute(16, '', '') <> mrYes then Result := False;
  end;
end;

procedure TFSaisODA.ClickValide;
var
  StE: String3;
  io: TIOErr;
  numero_axe : byte; //SG6 29/12/2004 Gestion en mode croisaxe
  i : integer;
  OkVerif:    Boolean;     {FP FQ15861 20/10/2005}
begin
  if not GSO.SynEnabled then Exit;
  if Action = taConsult then
  begin
    CloseFen;
    Exit;
  end;
  if not BValide.Enabled then Exit else if Outils97.CanFocus then Outils97.SetFocus;
  GSOExit(nil);
  //SG6 29/12/2004 Gestion du mode croisaxe
  if VH^.AnaCroisaxe then
  begin
    RechercheAxe(GSO.Row,numero_axe);
    if numero_axe = 0 then Exit;
    {b FP FQ15861 20/10/2005: Pas de vérif si aucun compte est renseigné pour l'écriture}
    OkVerif := False;
    for i := GSO.Row - numero_axe + 1 to GSO.Row - numero_axe + nbre_axe_ventil do begin
      if EstRempli(GSO, i) then
        OkVerif := True;
    end;
    if OkVerif then begin
    {e FP FQ15861 20/10/2005}
      for i:= GSO.Row-numero_axe + 1 to GSO.Row - numero_axe + nbre_axe_ventil do
      begin
        if not((EstRempli(GSO, i))) then
        begin
          ErreurSaisie(1);
          exit;
        end;
        if not(LigneCorrecte(i, True)) then Exit;
      end;
    end;     {FP FQ15861 20/10/2005}
  end
  else
  begin
    if ((EstRempli(GSO, GSO.Row)) and (not LigneCorrecte(GSO.Row, True))) then Exit;
  end;

  if not PieceCorrecte then Exit;
  if not OkBudget then Exit;
  if GSO.RowCount <> NbLOrig then Revision := False;
  GSO.Row := 1;
  GSO.Col := OA_Sect;
  GSO.SetFocus;
  StE := QuelExoDT(StrToDate(Y_DATECOMPTABLE.Text));
  if StE = VH^.Encours.Code then GeneTypeExo := teEncours else
    if StE = VH^.Suivant.Code then GeneTypeExo := teSuivant else GeneTypeExo := tePrecedent;
  NowFutur := NowH;
  GSO.Enabled := False;

  {JP 10/10/05 : FQ 16848 : Lors de l'initialisation de l'entête, le numéro est provisoire.
                 On attribue maintenant un nuémro provisoire}
  IncrementeNumero(False);

  //SG6 27/12/2004
  if not VH^.AnaCroisaxe then
  begin
    io := Transactions(ValideLaPiece, 3);
    GSO.Enabled := True;
    case io of
      oeOK: ;
      oeSaisie: MessageAlerte(HDiv.Mess[0]);
      oeUnknown: MessageAlerte(HDiv.Mess[1]);
    end;
  end
  else
  begin
    //initialisation
    V_PGI.IOError := oeOK;
    ValideLaPieceTrans;
  end;

  PieceModifiee := False;
  if Action = taCreat then ReInitPiece else CloseFen;
end;

function TFSaisODA.DejaCalcule(Sect: string): boolean;
var
  i: integer;
  T: T_SC;
begin
  Result := False;
  for i := 0 to TSoldeGS.Count - 1 do
  begin
    T := T_SC(TSoldeGS[i]);
    if T.Cpte = Sect then
    begin
      Result := True;
      Break;
    end;
  end;
end;

procedure TFSaisODA.ClickSoldeGS;
var
  i, borne_sup: integer; 
  Q: TQuery;
  St, Sect: string;
  TD, TC: double;
begin
  if PieceVide then Exit;
  if Y_GENERAL.Text = '' then Exit;
  if not BSoldeGS.Enabled then Exit;
  if ((EstRempli(GSO, GSO.Row)) and (not LigneCorrecte(GSO.Row, True))) then Exit;
  VideListe(TSoldeGS);
  //St:='Select Sum(Y_DEBIT), Sum(Y_CREDIT) from ANALYTIQ Where Y_GENERAl="'+Y_GENERAL.Text+'" '
  //   +'AND Y_SECTION=:SECT AND Y_AXE="'+SAJAL.Axe+'" AND Y_DATECOMPTABLE>="'+USDATETIME(VH^.Encours.Deb)+'" '
  //   +'AND Y_QUALIFPIECE="N" AND Y_TYPEANOUVEAU<>"ASA" AND Y_TYPEANOUVEAU<>"AOD" AND Y_TYPEANOUVEAU<>"ACP" ' ;
  //Q:=TQuery.Create(Application) ; Q.DataBaseName:=DBSOC.DataBaseName ;
  //Q.SQL.Add(St) ; ChangeSQL(Q) ; //Q.Prepare ;
  //PrepareSQLODBC(Q) ;

  //SG6 23/12/2004 Mode Croisaxe
  borne_sup:=Return_Bornesup;
  for i := 1 to borne_sup do
  begin
    Sect := GSO.Cells[OA_Sect, i];
    if ((Sect <> '') and (not DejaCalcule(Sect))) then
    begin
      St := 'SELECT SUM(Y_DEBIT), SUM(Y_CREDIT) FROM ANALYTIQ WHERE ' +
        'Y_GENERAl = "' + Y_GENERAL.Text + '" AND ' +
        'Y_SECTION = "' + Sect + '" AND ' +
        'Y_AXE = "' + SAJAL.Axe + '" AND ' +
        'Y_DATECOMPTABLE >= "' + USDATETIME(VH^.Encours.Deb) + '" AND ' +
        'Y_QUALIFPIECE = "N" AND Y_TYPEANOUVEAU <> "ASA" AND ' +
        'Y_TYPEANOUVEAU <> "AOD" AND Y_TYPEANOUVEAU <> "ACP" ';

      Q := OpenSql(St, True);
      //Q.ParamByName('SECT').AsString:=Sect ; Q.Open ;
      if not Q.EOF then
      begin
        TD := Q.Fields[0].AsFloat;
        TC := Q.Fields[1].AsFloat;
      end
      else
      begin
        TD := 0;
        TC := 0;
      end;
      Ajoute(TSoldeGS, SAJAL.Axe, Sect, TD, TC,-1,-1,iDate1900);
      //Q.Close ;
      Ferme(Q);
    end;
  end;
  //Q.Close ;
  //Q.Free ;
  AffichePied;
end;

procedure TFSaisODA.ClickAbandon;
begin
  CloseFen;
end;

procedure TFSaisODA.ClickComplement;
var
  ModBN: boolean;
  RC: R_COMP;
  i : integer;
  numero_axe : byte;
begin
  if not BComplement.Enabled then Exit;

  //SG6 28/12/2004 Gestion du Complement en mode croisawe
  if VH^.AnaCroisaxe then
  begin
    RechercheAxe(GSO.Row,numero_axe);
    if numero_axe = 0 then exit;
    for i:= GSO.Row-numero_axe + 1 to GSO.Row - numero_axe + nbre_axe_ventil do
    begin
      if not EstRempli(GSO, i) then Exit;
      if not LigneCorrecte(i, True) then Exit;
    end;
  end
  else
  begin
    if not EstRempli(GSO, GSO.Row) then Exit;
    if not LigneCorrecte(GSO.Row, True) then Exit;
  end;


  RC.StComporte := '--XXXXXXXX';
  RC.StLibre := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
  RC.Conso := True;
  RC.Attributs := False;
  RC.MemoComp := nil;
  RC.Origine := -1;
  RC.DateC := 0;
  RC.TOBCompl := nil ; //FQ 15715  
  //SG6 28/12/2004 en mode croisaxe, on stock juste les infos sur la première ligne
  // (puique c'est la premier ligne qui sera enregisgtré dans la DB)
  if not VH^.AnaCroisaxe then
  begin
    if not SaisieComplement(GetO(GSO, GSO.Row), EcrAna, Action, ModBN, RC) then Exit;
  end
  else
  begin
    if not SaisieComplement(GetO(GSO, GSO.Row-numero_axe + 1), EcrAna, Action, ModBN, RC) then Exit;
  end;

  if GSO.CanFocus then GSO.SetFocus;
  if ModBN then Revision := False;
end;

procedure TFSaisODA.ClickInsert;
var
  R: integer;
  numero_axe : byte;
begin
  if Action = taConsult then exit;
  if not EstRempli(GSO, GSO.Row) then Exit;
  if not LigneCorrecte(GSO.Row, True) then Exit;
  GSO.CacheEdit;
  //SG6 29/12/2004 Gestion en mode croisaxe
  if not VH^.AnaCroisaxe then
  begin
    R := GSO.Row;
  end
  else
  begin
    RechercheAxe(GSO.Row,numero_axe);
    if numero_axe = 0 then Exit;
    GSO.Row := GSO.Row - numero_axe + 1;
    R := GSO.Row;
  end;

  InsereLigne(GSO.Row);
  GSO.Row := R;
  GSO.Col := OA_Sect;
  GSO.Cells[GSO.Col, GSO.Row] := '';
  GSO.SetFocus;
  if not VH^.AnaCroisaxe then GereNewLigne(GSO);
  GSO.MontreEdit;
  Revision := False;
end;

procedure TFSaisODA.ClickDel;
begin
  if Action = taConsult then exit;
  if (GSO.RowCount <= 1 + nbre_axe_ventil) then exit;

  if ((GSO.Row = GSO.RowCount - 1) and (not EstRempli(GSO, GSO.Row))) then
  begin
    GSO.SetFocus;
    Exit;
  end;
  if not VH^.AnaCroisaxe then DetruitLigne(GSO.Row)
  else DetruitLigneCroisaxe(GSO.Row);

  Revision := False;
  ControleLignes;
end;

procedure TFSaisODA.ClickEtabl;
begin
  FicheEtablissement_AGL(taConsult);
end;

procedure TFSaisODA.ClickSolde(Egal: boolean);
var
  sDebit,sCredit : string;
  i : integer;
  numero_axe : byte;
begin
  if Action = taConsult then Exit;
  if not BSolde.Enabled then Exit;
  if not EstRempli(GSO, GSO.Row) then Exit;
  if ((Egal) and (GSO.Col <> OA_Sect) and (GSO.Col <> OA_Debit) and (GSO.Col <> OA_Credit)) then Exit;

  if not VH^.AnaCroisaxe then GereNewLigne(GSO)
  else  if EstRempli(GSO,GSO.RowCount - nbre_axe_ventil) then GSO.RowCount := GSO.RowCount + nbre_axe_ventil;

  if ((GSO.Col = OA_Debit) or (GSO.Col = OA_Credit)) then
  begin
    FormatMontant(GSO, OA_Debit, GSO.Row, V_PGI.OkDecV);
    FormatMontant(GSO, OA_Credit, GSO.Row, V_PGI.OkDecV);
    TraiteMontant(GSO.Row, True);

    //Mise a jour pour en mode croisaxe des autres lignes SG6 28/12/2004
    sDebit := GSO.Cells[OA_Debit,GSO.Row];
    sCredit := GSO.Cells[OA_Credit,GSO.Row];
    RechercheAxe(GSO.Row,numero_axe);
    for i:= GSO.Row-numero_axe + 1 to GSO.Row - numero_axe + nbre_axe_ventil do
    begin
       GSO.Cells[OA_Debit,i] := sDebit;
       GSO.Cells[OA_Credit,i] := sCredit;
    end;

  end;
  SoldeLaLigne(GSO.Row);

end;

procedure TFSaisODA.BVentilTypeClick(Sender: TObject);
begin
  ClickVentilType;
  if GSO.Enabled then GSO.SetFocus;
end;

procedure TFSaisODA.BSoldeGSClick(Sender: TObject);
begin
  {JP 04/08/05 : FQ 15859/15651 : Les BitButtons ne prenant pas le focus, on force la sortie de la cellule}
  LostFocusBeforeClick;

  ClickSoldeGS;
end;

procedure TFSaisODA.BSoldeClick(Sender: TObject);
begin
  {JP 04/08/05 : FQ 15859/15651 : Les BitButtons ne prenant pas le focus, on force la sortie de la cellule}
  LostFocusBeforeClick;

  ClickSolde(False);
end;

procedure TFSaisODA.BSDelClick(Sender: TObject);
begin
  ClickDel;
end;

procedure TFSaisODA.BInsertClick(Sender: TObject);
begin
  ClickInsert;
end;

procedure TFSaisODA.BComplementClick(Sender: TObject);
begin
  ClickComplement;
end;

procedure TFSaisODA.BAbandonClick(Sender: TObject);
begin
  ClickAbandon;
end;

procedure TFSaisODA.BValideClick(Sender: TObject);
begin
  {JP 04/08/05 : FQ 15859/15651 : Les BitButtons ne prenant pas le focus, on force la sortie de la cellule}
  LostFocusBeforeClick;

  ClickValide;
end;

procedure TFSaisODA.BChercherClick(Sender: TObject);
begin
  ClickCherche;
end;

procedure TFSaisODA.BZoomClick(Sender: TObject);
begin
  ClickZoom;
end;

procedure TFSaisODA.BZoomPieceClick(Sender: TObject);
begin
  {JP 03/08/05 : FQ 15854 : On cache le zoom si l'on est en train de créer la pièce}
  if Action <> taCreat then
    ClickZoomPiece;
end;

procedure TFSaisODA.BZoomEtablClick(Sender: TObject);
begin
  ClickEtabl;
end;

{====================================== CHARGEMENTS ======================================}
procedure TFSaisODA.ChargeEcritures;
var
  Lig: integer;
  OBM: TOBM;
  XD, XC: Double;
  JD, JC: Double;
  QODA: TQuery;
  XDEL: TDELMODIF;
begin
  if Action = taCreat then Exit;
  BeginTrans;
  NbLigAna := 0;
  InitMove(50, '');
  // SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
  QODA := OpenSQL('SELECT ' + GetSelectAll('Y', True) + ', Y_BLOCNOTE FROM ANALYTIQ WHERE ' + WhereEcriture(tsODA, M, False)
    + ' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE,'
    + ' Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL', True);
  Lig := 0;
  JD := 0;
  JC := 0;
  VideListe(TDELANA);
  while not QODA.EOF do
  begin
    MoveCur(FALSE);
    {Sauvegarde des dates modif}
    XDEL := TDELMODIF.Create;
    XDEL.NumLigne := QODA.FindField('Y_NUMLIGNE').AsInteger;
    XDEL.Ax := QODA.FindField('Y_AXE').AsString;
    XDEL.NumEcheVent := QODA.FindField('Y_NUMVENTIL').AsInteger;
    XDEL.DateModification := QODA.FindField('Y_DATEMODIF').AsDateTime;
    TDELANA.Add(XDEL);
    {Lecture des infos}
    Inc(Lig);
    DefautLigne(Lig, False);
    OBM := GetO(GSO, Lig);
    OBM.ChargeMvt(QODA);
    OBM.HistoMontants;
    if OBM.GetMvt('Y_CONFIDENTIEL') > '0' then
    begin
      PieceConf := True;
      HConf.Visible := True;
    end;
    {$IFDEF EAGLCLIENT}
    // A FAIRE
    {$ELSE}
    OBM.M.Assign(TMemoField(QODA.FindField('Y_BLOCNOTE')));
    {$ENDIF}
    GSO.Cells[OA_Sect, Lig] := QODA.FindField('Y_SECTION').AsString;
    GSO.Cells[OA_RefI, Lig] := QODA.FindField('Y_REFINTERNE').AsString;
    GSO.Cells[OA_Lib, Lig] := QODA.FindField('Y_LIBELLE').AsString;
    XD := QODA.FindField('Y_DEBIT').AsFloat;
    XC := QODA.FindField('Y_CREDIT').AsFloat;
    JD := JD + XD;
    JC := JC + XC;
    GSO.Cells[OA_Debit, Lig] := StrS(XD, V_PGI.OkDecV);
    GSO.Cells[OA_Credit, Lig] := StrS(XC, V_PGI.OkDecV);
    FormatMontant(GSO, OA_Debit, Lig, V_PGI.OkDecV);
    FormatMontant(GSO, OA_Credit, Lig, V_PGI.OkDecV);
    if Lig = 1 then
    begin
      NumLigGene := QODA.FindField('Y_NUMLIGNE').AsInteger;
      TotPGene := QODA.FindField('Y_TOTALECRITURE').AsFloat;
      TotDGene := QODA.FindField('Y_TOTALDEVISE').AsFloat;
      PureAnal := (NumLigGene <= 0);
      if not PureAnal then
      begin
        if XD <> 0 then SensGene := 1 else SensGene := 2;
      end;
    end;
    GereNewLigne(GSO);
    QODA.Next;
    Inc(NbLigAna);
  end;
  Ferme(QODA);
  SAJAL.OldDebit := JD;
  SAJAL.OldCredit := JC;
  if not PureAnal then Y_GENERAL.Enabled := False;
  FiniMove;
  CommitTrans;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/04/2004
Modifié le ... :   /  /
Description .. : Passage en EAGL
Mots clefs ... :
*****************************************************************}
procedure TFSaisODA.ChargeSections(Force: boolean);
var
  lQuery: TQuery;
  i: integer;
  CSect: TGSection;
  sAxe : string;
  iAxe : integer ; //SG6 27/12/2004 Gestion en mode croisaxe
  numero_axe : byte;
begin
  if ((Action = taCreat) and (not Force)) then Exit;
  BeginTrans;
  //SG6 23/12/2004 Mode Croisaxe
//  borne_sup := Return_Bornesup; FQ 15881
  for i := 1 to GSO.RowCount - 1 {borne_sup} do
    if GSO.Cells[OA_Sect, i] <> '' then
    begin
      CSect := nil;
      if not VH^.AnaCroisaxe then
      begin
        sAxe := SAJAL.AXE;
      end
      else
      begin
        iAxe := RechercheAxe(i,numero_axe);
        if iAxe = 0 then exit;
        sAxe := 'A'+IntTostr(iAxe);
      end;

      lQuery := OpenSql(SelectSection + ' FROM SECTION WHERE ' +
        'S_AXE = "' + sAxe + '" AND ' +
        'S_SECTION = "' + GSO.Cells[OA_Sect, i] + '"', True);
      if not lQuery.EOF then
      begin
        CSect := TGSection.Create('', tzSection);
        CSect.QueryToSection(lQuery);
      end;
      Ferme(lQuery);
      GSO.Objects[OA_Sect, i] := TObject(CSect);
    end;
  CommitTrans;
end;

procedure TFSaisODA.ChargeLignes;
begin
  if Action = taCreat then Exit;
  if not VH^.AnaCroisaxe then
    ChargeEcritures
  else
    ChargeEcrituresCroisaxe;
    
  ChargeSections(False);
  CalculDebitCredit;
  GSO.Row := 1;
  GSO.Col := OA_Sect;
  GSO.SetFocus;
  AffichePied;
  NbLOrig := GSO.RowCount;
end;

procedure TFSaisODA.ChargeSoldesComptes;
var
  i, borne_sup: integer; //SG6 23/12/2004 Gestion croisaxe 
  OBM: TOBM;
  Sect: String17;
  D, C: Double;
begin
  if Action <> taModif then exit;
  //SG6 23/12/2004 Gestion du mode croisaxe
  borne_sup:=Return_Bornesup;
  for i := 1 to borne_sup do
  begin
    OBM := GetO(GSO, i);
    if OBM = nil then Continue;
    Sect := OBM.GetMvt('Y_SECTION');
    D := OBM.GetMvt('Y_DEBIT');
    C := OBM.GetMvt('Y_CREDIT');
    Ajoute(TS, SAJAL.Axe, Sect, D, C,-1,-1,iDate1900,'','');
  end;
end;

{=============================== ALLOCATIONS, SOLDE =================================}
procedure TFSaisODA.DesalloueA(Lig: Longint);
begin
  if GSO.Objects[OA_Sect, Lig] = nil then Exit;
  TGSection(GSO.Objects[OA_Sect, Lig]).Free;
  GSO.Objects[OA_Sect, Lig] := nil;
end;

procedure TFSaisODA.SoldelaLigne(Lig: integer);
var
  Diff, XD: Double;
  Col: integer;
  b, bb: boolean;
begin
  Diff := OI_TotDebit - OI_TotCredit;
  b := False;
  bb := False;
  if not PureAnal then
  begin
    if SensGene = 1 then Diff := Diff - TotPGene else Diff := Diff + TotPGene;
  end;
  XD := -1 * ValD(GSO, Lig);
  if XD = 0 then XD := ValC(GSO, Lig);
  Diff := Diff + XD;
  GSO.Cells[OA_Debit, Lig] := '';
  GSO.Cells[OA_Credit, Lig] := '';
  if Diff > 0 then
  begin
    if SensGene <> 1 then
    begin
      Col := OA_Credit;
      GSO.Cells[Col, Lig] := StrS(Diff, V_PGI.OkDecV);
    end else
    begin
      Col := OA_Debit;
      GSO.Cells[Col, Lig] := StrS(-Diff, V_PGI.OkDecV);
    end;
  end else
  begin
    if SensGene <> 2 then
    begin
      Col := OA_Debit;
      GSO.Cells[Col, Lig] := StrS(-Diff, V_PGI.OkDecV);
    end else
    begin
      Col := OA_Credit;
      GSO.Cells[Col, Lig] := StrS(Diff, V_PGI.OkDecV);
    end;
  end;
  ChercheMontant(Col, Lig);
  if VH^.AnaCroisaxe then GSOCellExit(nil,Col,Lig,bb);

  GSORowExit(nil, Lig, b, bb);

end;

{================================== RECHERCHES ==================================}
function TFSaisODA.ChercheSect(L: integer; Force: boolean): byte;
var
  St: string;
  CSect: TGSection;
  Changed, Idem: boolean;
  begin
  Result := 0;
  Changed := False;
  St := uppercase(ConvertJoker(GSO.Cells[OA_Sect, L]));
  Cache.Text := St;
  if GChercheCompte(Cache, FicheSection) then
  begin
    if GSO.Objects[OA_Sect, L] <> nil then if TGSection(GSO.Objects[OA_Sect, L]).Sect <> St then Changed := True;
    Idem := ((St = Cache.Text) and (GSO.Objects[OA_Sect, L] <> nil) and (not Changed));
    if ((not Idem) or (Force)) then
    begin
      GSO.Cells[OA_Sect, L] := Cache.Text;
      DesalloueA(L);
      CSect := TGSection.Create(Cache.Text, Cache.ZoomTable);
      GSO.Objects[OA_Sect, L] := TObject(CSect);
      if ((not PureAnal) and (Action = taModif) and (not PieceConf) and (CSect.Confidentiel > '0')) then
      begin
        DesalloueA(L);
        GSO.Cells[OA_Sect, L] := '';
        HMessLigne.Execute(3, caption, '');
        Exit;
      end;
      if CSect.Ferme then
      begin
        if Arrondi(CSect.TotalDebit - CSect.TotalCredit, V_PGI.OkDecV) = 0 then
        begin
          DesalloueA(L);
          GSO.Cells[OA_Sect, L] := '';
          ErreurSaisie(5);
          Exit;
        end else
        begin
          ErreurSaisie(3);
        end;
      end;
      Result := 1;
    end else
    begin
      {CSect:=TGSection(GSO.Objects[OA_Sect,L]) ;}Result := 2;
    end;
  end;
end;

procedure TFSaisODA.FindSaisFind(Sender: TObject);
begin
  Rechercher(GSO, FindSais, FindFirst);
end;

{====================================== VALIDATIONS ====================================}
procedure TFSaisODA.DetruitAncien;
var
  i: integer;
  StW, StLig, sAxe: string; //SG6 27/12/2004 Gestion croisaxe
  X: TDELMODIF;
begin
  if Action = taConsult then Exit;
  if Revision then Exit;
  StW := WhereEcriture(tsODA, M, False);
  for i := 0 to TDELANA.Count - 1 do
  begin
    X := TDELMODIF(TDELANA[i]);
    sAxe := ' AND Y_AXE="' + X.Ax + '"';

    StLig := ' AND Y_NUMLIGNE=' + IntToStr(X.NumLigne) + sAxe + ' AND Y_NUMVENTIL=' + InttoStr(X.NumEcheVent) + ' AND Y_DATEMODIF="' + UsTime(X.DateModification) + '"';
    if ExecuteSQL('Delete from Analytiq where ' + StW + StLig) <= 0 then
    begin
      V_PGI.IOError := oeSaisie;
      Break;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/04/2004
Modifié le ... :   /  /
Description .. : Passage en EAGL
Mots clefs ... :
*****************************************************************}
procedure TFSaisODA.InverseSolde;
var
  i: integer;
  T: T_SC;
  lFRM: TFRM;
begin
  if Action = taConsult then Exit;
  // Sections
  for i := 0 to TS.Count - 1 do
  begin
    T := T_SC(TS[i]);
    lFRM.Cpt := T.Cpte;
    lFRM.Axe := T.Identi;
    lFRM.Deb := T.Debit;
    lFRM.Cre := T.Credit;

    AttribParamsNew(lFRM, lFRM.Deb, lFRM.Cre, GeneTypeExo);

    if ExecuteSQL('UPDATE SECTION SET ' +
      'S_TOTALDEBIT  = S_TOTALDEBIT  - ' + StrFPoint(lFRM.Deb) + ', ' +
      'S_TOTALCREDIT = S_TOTALCREDIT - ' + StrFPoint(lFRM.Cre) + ', ' +
      'S_TOTDEBE     = S_TOTDEBE - ' + StrFPoint(lFRM.DE) + ', ' +
      'S_TOTCREE     = S_TOTCREE - ' + StrFPoint(lFRM.CE) + ', ' +
      'S_TOTDEBS     = S_TOTDEBS - ' + StrFPoint(lFRM.DS) + ', ' +
      'S_TOTCRES     = S_TOTCRES - ' + StrFPoint(lFRM.CS) + ', ' +
      'S_TOTDEBP     = S_TOTDEBP - ' + StrFPoint(lFRM.DP) + ', ' +
      'S_TOTCREP     = S_TOTCREP - ' + StrFPoint(lFRM.CP) + ' WHERE ' +
      'S_AXE = "' + lFRM.Axe + '" AND ' +
      'S_SECTION = "' + lFRM.Cpt + '"') <> 1 then V_PGI.IoError := oeSaisie;
  end;

  // Journal
  lFRM.Cpt := SAJAL.JOURNAL;
  lFRM.Deb := SAJAL.OldDebit;
  lFRM.Cre := SAJAL.OldCredit;
  AttribParamsNew(lFRM, lFRM.Deb, lFRM.Cre, GeneTypeExo);

  if ExecuteSQL('UPDATE JOURNAL SET ' +
    'J_TOTALDEBIT  = J_TOTALDEBIT  - ' + StrFPoint(lFRM.Deb) + ', ' +
    'J_TOTALCREDIT = J_TOTALCREDIT - ' + StrFPoint(lFRM.Cre) + ', ' +
    'J_TOTDEBE     = J_TOTDEBE     - ' + StrFPoint(lFRM.DE) + ', ' +
    'J_TOTCREE     = J_TOTCREE     - ' + StrFPoint(lFRM.CE) + ', ' +
    'J_TOTDEBS     = J_TOTDEBS     - ' + StrFPoint(lFRM.DS) + ', ' +
    'J_TOTCRES     = J_TOTCRES     - ' + StrFPoint(lFRM.CS) + ', ' +
    'J_TOTDEBP     = J_TOTDEBP     - ' + StrFPoint(lFRM.DP) + ', ' +
    'J_TOTCREP     = J_TOTCREP     - ' + StrFPoint(lFRM.CP) + ' WHERE ' +
    'J_JOURNAL = "' + lFRM.Cpt + '"') <> 1 then V_PGI.IoError := oeSaisie;

end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 06/04/2004
Modifié le ... :   /  /
Description .. : Passage en EAGL
Mots clefs ... :
*****************************************************************}
procedure TFSaisODA.MajCptesSection;
var
  XD, XC: Double;
  Lig, borne_sup,iAxe: integer; //SG6 23/12/2004 Gestion croisaxe
  lFRM: TFRM;
  numero_axe : byte;
begin
  if Action = taConsult then Exit;
  borne_sup:=Return_Bornesup;
  for Lig := 1 to borne_sup do
  begin
    if ((GSO.Cells[OA_Sect, Lig] <> '') and (GSO.Objects[OA_Sect, Lig] <> nil)) then
    begin
      lFRM.Cpt := GSO.Cells[OA_Sect, Lig];
      //Sg6 29/12/2004 Gestion du mode croisaxe
      if VH^.AnaCroisaxe then
      begin
        iAxe := RechercheAxe(Lig,numero_axe);
        if iAxe = 0 then exit;
        lFRM.Axe := 'A'+IntToStr(iAxe);
      end
      else
      begin
        lFRM.Axe := SAJAL.AXE;
      end;
      XD := ValD(GSO, Lig);
      XC := ValC(GSO, Lig);

      if not M.ANouveau then
      begin
        lFRM.NumD := NumPieceInt;
        lFRM.DateD := DateMvt(Lig);
        lFRM.LigD := Lig;
        AttribParamsNew(lFRM, XD, XC, GeneTypeExo);
      end
      else
      begin
        lFRM.Deb := XD;
        lFRM.Cre := XC;
      end;

      if ExecReqMAJ(fbSect, M.ANouveau, False, lFRM) <> 1 then
        V_PGI.IoError := oeSaisie;

    end;
  end;
end;

procedure TFSaisODA.ValideLesSections;
begin
  if Action = taConsult then exit;
  if M.Simul <> 'N' then Exit;
  MajCptesSection;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 06/04/2004
Modifié le ... :   /  /
Description .. : Passage en EAGL
Mots clefs ... :
*****************************************************************}
procedure TFSaisODA.ValideLeJournal;
var
  lFRM: TFRM;
begin
  if Action = taConsult then Exit;
  if M.Simul <> 'N' then Exit;

  lFRM.Cpt := SAJAL.Journal;
  lFRM.NumD := NumPieceInt;
  lFRM.DateD := DateMvt(1);

  AttribParamsNew(lFRM, OI_TotDebit, OI_TotCredit, GeneTypeExo);

  if ExecReqMAJ(fbJal, M.ANouveau, False, lFRM) <> 1 then
    V_PGI.IoError := oeSaisie;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 06/04/2004
Modifié le ... :   /  /
Description .. : Passage en EAGL
Mots clefs ... :
*****************************************************************}
procedure TFSaisODA.RecupTronc(Lig: Integer);
var
  CSect: TGSection;
  OBM: TOBM;
  StConf: string;
begin
  if Action = taConsult then Exit;
  CSect := GetGSect(GSO, Lig);
  if CSect = nil then Exit;
  OBM := GetO(GSO, Lig);
  if OBM = nil then Exit;
  StConf := '0';
  if HConf.Visible then StConf := '1';
  TEcrODA.Dupliquer(OBM, True, True, True);
  TEcrODA.PutValue('Y_GENERAL', Y_GENERAL.Text);
  if Lig = 1 then
    TEcrODA.PutValue('Y_TYPEMVT', 'AE')
  else
    TEcrODA.PutValue('Y_TYPEMVT', 'AL');

  //TMemoField(TEcrODA.FindField('Y_BLOCNOTE')).Assign(OBM.M) ;
  if PureAnal then
  begin
    TEcrODA.PutValue('Y_AXE', SAJAL.Axe);
    TEcrODA.PutValue('Y_TYPEANALYTIQUE', 'X');
    TEcrODA.PutValue('Y_NATUREPIECE', 'OD');
    if M.ANouveau then
    begin
      TEcrODA.PutValue('Y_ECRANOUVEAU', 'OAN');
      TEcrODA.PutValue('Y_VALIDE', 'X');
      if TEcrODA.GetValue('Y_TYPEANOUVEAU') = '-' then
        TEcrODA.PutValue('Y_TYPEANOUVEAU', 'ASA');
      //TEcrODA.FindField('Y_TYPEANALYTIQUE').AsString:='X' ;
    end;
  end
  else
  begin
    TEcrODA.PutValue('Y_TYPEANALYTIQUE', '-');
    TEcrODA.PutValue('Y_NATUREPIECE', M.Nature);
  end;
  TEcrODA.PutValue('Y_QUALIFPIECE', M.Simul);
  TEcrODA.PutValue('Y_DATEMODIF', NowFutur);
  TEcrODA.PutValue('Y_CONFIDENTIEL', StConf);
end;

procedure TFSaisODA.RecupFromGrid(Lig: integer);
var
  OBM: TOBM;
  St: string;
  //    Sens : integer ;
  SQL, StConf: string;
  DD: TDateTime;
begin
  OBM := GetO(GSO, Lig);
  StConf := '0';
  if HConf.Visible then StConf := '1';
  DD := OBM.GetMvt('Y_DATEMODIF');
  //if ValD(GSO,Lig)<>0 then Sens:=1 else Sens:=2 ;
  OBM.PutMvt('Y_GENERAL', Y_GENERAL.Text);
  OBM.PutMvt('Y_QUALIFPIECE', M.Simul);
  OBM.PutMvt('Y_NATUREPIECE', M.Nature);
  if PureAnal then
  begin
    OBM.PutMvt('Y_AXE', SAJAL.Axe);
    OBM.PutMvt('Y_TYPEANALYTIQUE', 'X');
    if M.ANouveau then
    begin
      OBM.PutMvt('Y_ECRANOUVEAU', 'OAN');
      OBM.PutMvt('Y_VALIDE', 'X');
      if OBM.GetMvt('Y_TYPEANOUVEAU') = '-' then OBM.PutMvt('Y_TYPEANOUVEAU', 'ASA');
    end;
  end else
  begin
    OBM.PutMvt('Y_TYPEANALYTIQUE', '-');
  end;
  OBM.PutMvt('Y_TYPEMVT', 'JLD'); {Forcer modif}
  OBM.PutMvt('Y_CONFIDENTIEL', StConf);
  if Lig = 1 then OBM.PutMvt('Y_TYPEMVT', 'AE') else OBM.PutMvt('Y_TYPEMVT', 'AL');
  St := OBM.StPourUpdate;
  if St = '' then Exit;
  SQL := 'UPDATE ANALYTIQ SET ' + St + ', Y_DATEMODIF="' + UsTime(NowFutur) + '" Where ' + WhereEcriture(tsODA, M, False);
  SQL := SQL + ' AND Y_NUMLIGNE=' + IntToStr(NumLigGene) + ' AND Y_NUMVENTIL=' + IntToStr(Lig)
    + ' AND Y_DATEMODIF="' + UsTime(DD) + '"';
  if ExecuteSQL(SQL) <= 0 then V_PGI.IOError := oeSaisie;
end;

procedure TFSaisODA.GetODA(Lig: Integer);
var
  CSect: TGSection;
begin
  if Action = taConsult then exit;
  CSect := GetGSect(GSO, Lig);
  if CSect = nil then Exit;
  RecupTronc(Lig);
  //SG6 23/21/04 Gestion du mode analytique (croisaxe ou mono axe)
  if not VH^.AnaCroisaxe then  TEcrODA.InsertDB(nil)
  else InsertCroisaxe(Lig); {FQ 15881}
end;

procedure TFSaisODA.GetODAGrid(Lig: Integer);
var
  CSect: TGSection;
begin
  if Action = taConsult then exit;
  CSect := GetGSect(GSO, Lig);
  if CSect = nil then Exit;
  RecupFromGrid(Lig);
end;

procedure TFSaisODA.ValideLignesODA;
var
  i, borne_sup: integer; //SG6 23/12/2004 Gestion croisaxe
begin
  if Action = taConsult then exit;
  borne_sup:=Return_Bornesup;
  initmove(borne_sup, '');
  if Revision then
  begin
    //SG6 23/12/2004 Mode croisaxe
    if not VH^.AnaCroisaxe then
    begin
      for i := 1 to borne_sup do
      begin
        MoveCur(FALSE);
        GetODAGrid(i);
        if V_PGI.IOError <> oeOk then Break;
      end;
    end
    else
    begin
      i:=1;
      while(i<borne_sup+1) do
      begin
        MiseAJourODA(i);
        Inc(i,nbre_axe_ventil);
      end;
    end;
  end else
  begin
    TEcrODA := Tob.Create('ANALYTIQ', nil, -1);
    //SG6 23/12/2004 Mode Croisaxe
    if not VH^.AnaCroisaxe then
    begin
      for i := 1 to borne_sup do
      begin
        MoveCur(FALSE);
        GetODA(i);
        if V_PGI.IOError <> oeOk then Break;
      end;
    end
    else
    begin
      i:=1;
      while(i<borne_sup+1) do
      begin
        GetODA(i);
        Inc(i,nbre_axe_ventil);
      end;
    end;
  end;
  FreeAndNil(TEcrODA);
  FiniMove;
end;

procedure TFSaisODA.CloseFen;
begin
  //if not IsInside(Self) then Close;
  //SG6 13/01/05 FQ 15158 Fermeture de la saisie analytique
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFSaisODA.ValidelaPiece;
begin
  if not GSO.SynEnabled then Exit;
  GridEna(GSO, False);
  if Action = taModif then
  begin
    DetruitAncien;
    if V_PGI.IOError = oeOK then InverseSolde;
  end;
  if V_PGI.IOError = oeOK then ValideLignesODA;
  if V_PGI.IOError = oeOK then ValideLesSections;
  if V_PGI.IOError = oeOK then ValideLeJournal;
  GridEna(GSO, True);
end;

{================================== Méthodes de la FORM ==================================}
procedure TFSaisODA.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := FermerSaisie;
end;

procedure TFSaisODA.POPSPopup(Sender: TObject);
begin
  InitPopup(Self);
end;

procedure TFSaisODA.PosLesSoldes;
begin
  LOA_SoldeS.SetBounds(OA_SoldeS.Left, OA_SoldeS.Top, OA_SoldeS.Width, OA_SoldeS.Height);
  LOA_SoldeGS.SetBounds(OA_SoldeGS.Left, OA_SoldeGS.Top, OA_SoldeGS.Width, OA_SoldeGS.Height);
  LOA_Solde.SetBounds(OA_Solde.Left, OA_Solde.Top, OA_Solde.Width, OA_Solde.Height);
  LOA_TotalDebit.SetBounds(OA_TotalDebit.Left, OA_TotalDebit.Top, OA_TotalDebit.Width, OA_TotalDebit.Height);
  LOA_TotalCredit.SetBounds(OA_TotalCredit.Left, OA_TotalCredit.Top, OA_TotalCredit.Width, OA_TotalCredit.Height);
  LOA_MontantEcr.SetBounds(OA_MontantEcr.Left, OA_MontantEcr.Top, OA_MontantEcr.Width, OA_MontantEcr.Height);
  OA_SoldeS.Visible := False;
  OA_SoldeGS.Visible := False;
  OA_Solde.Visible := False;
  OA_TotalDebit.Visible := False;
  OA_TotalCredit.Visible := False;
  OA_MontantEcr.Visible := False;
end;

procedure TFSaisODA.FormCreate(Sender: TObject);
var i : integer;  //SG6 2/12/2004
begin
  GSO.RowCount := 2;

  //SG6 22/12/04
  nbre_axe_ventil := 1;
  if VH^.AnaCroisaxe then
  begin
    nbre_axe_ventil := 0;
    for i := 1 to MaxAxe do
    begin
      Axes[i] := GetParamSoc('SO_VENTILA' + IntToStr(i));
      if Axes[i] then Inc(nbre_axe_ventil);
    end;

    //Nbre de lignes en fonction
    GSO.RowCount := nbre_axe_ventil+1;
  end;



  WMinX := Width;
  WMinY := Height_ODA;
  GSO.TypeSais := tsODA;
  GSO.ListeParam := 'SSAISODA';

  {b FP FQ16213}
  {JP 07/10/05 : FQ 16213 : Gestion de la largeur des colonnes}
  GSO.ColWidths[OA_EXO  ] := -1;
  GSO.ColWidths[OA_NUMP ] := -1;
  GSO.ColWidths[OA_DATEC] := -1;
  GSO.ColWidths[OA_NUML ] := 23;

  if  VH^.AnaCroisaxe then
  begin
    GSO.FixedCols := GSO.FixedCols + 1; {La colonne Axe est fixe}
    GSO.ColWidths[OA_AXE] := 23;
  end
  else
  begin
    GSO.ColWidths[OA_AXE]    := -1;        {Masque la colonne axe}
    GSO.ColLengths[OA_AXE]   := -1;
    {JP 10/07/07 : FQ 20990 : pour une raison que j'ignore la colonne Crédit est invisible,
                   on la renseigne avec les valeurs de la colonne debit}
    GSO.ColWidths[OA_Credit]  := GSO.ColWidths[OA_Debit];
    GSO.ColLengths[OA_Credit] := GSO.ColLengths[OA_Debit];
    GSO.ColEditables[OA_AXE] := False;
  end;

  {e FP FQ16213}
  SAJAL := nil;
  TS := TList.Create;
  TSOLDEGS := TList.Create;
  TAX := HTStringList.Create;
  TDELANA := TList.Create;
  PieceModifiee := False;
  Revision := False;
  AbregeGen := '';
  CpteGeneChange := False;
  NumLigGene := 0;
  TotPGene := 0;
  TotDGene := 0;
  TotEGene := 0;
  PureAnal := True;
  SensGene := 0;
  FillChar(DEV, Sizeof(DEV), #0);
  DEV.Code := V_PGI.DevisePivot;
  GetInfosDevise(DEV);
  OI_TotQ1 := 0;
  OI_TotQ2 := 0;
  PieceConf := False;
  HConf.Visible := False;
  RegLoadToolbarPos(Self, 'SaisODA');
  PosLesSoldes;
  GSO.ColLengths[OA_RefI] := 35;
  GSO.ColLengths[OA_Lib] := 35;
end;

procedure TFSaisODA.FormShow(Sender: TObject);
var
  OKM, bc, OkDEV, OkPasJal: Boolean;
begin
  {JP 03/08/05 : FQ 15854 : On cache le zoom si l'on est en train de créer la pièce}
  BZoomPiece.Enabled := Action <> taCreat;
  //FQ 14806 : SG6 2/11/2004
  H_AXE.Width := H_DATECOMPTABLE.Left - H_AXE.Left - 5;
  //Fin FQ 14806
  LookLesDocks(Self);
  Revision := False; {OkM:=((M.Valide) and (Action=taModif)) ;}
  OKM := False;
  OkDEV := False;
  if Action = taCreat then
  begin
    if M.ANouveau then Y_JOURNAL.DataType := 'TTJALANALAN' else Y_JOURNAL.DataType := 'ttJalAnalytique';
  end else Y_JOURNAL.DataType := 'TTJOURNAUX';
  ChangeFormatPivot(Self);
  NbLigAna := 0;
  DefautEntete;
  DefautPied;
  ChargeLignes;
  if DEV.Code <> V_PGI.DevisePivot then OkDev := True;
  OkPasJal := ((VH^.JalAutorises <> '') and (Pos(';' + Y_JOURNAL.Value + ';', VH^.JalAutorises) <= 0) and (Action = taModif));
  if ((OkM) or (OkDEV) or (OkPasJal)) then Action := taConsult;
  case Action of
    taCreat:
      begin
        if M.ANouveau then GereANouveau;
        // FQ 18743 SBO 07/09/2006 : En creation, sélection du 1er journal ou msg d'avertissement / sortie ou création
        if not OkJal then
          begin
          PieceModifiee := False ;
          CloseFen ;
          Exit ;
          end ;
        Y_JOURNAL.SetFocus; //Y_GENERAL.Enabled:=False ;
      end;
    taModif:
      begin
        ChargeSoldesComptes;
        Revision := True;
        if SAJAL.NatureJal = 'ANA' then
        begin
          M.ANouveau := True;
          GereANouveau;
        end;
      end;
    taConsult:
      begin
        LectureSeule;
        if SAJAL.NatureJal = 'ANA' then
        begin
          M.ANouveau := True;
          GereANouveau;
        end;
      end;
  end;
  (*
  if Action<>taCreat then
    BEGIN
    If M.ANouveau Then Y_JOURNAL.DataType:='TTJALANALAN' Else Y_JOURNAL.DataType:='ttJalAnalytique' ;
    Y_JOURNAL.Value:=M.Jal ; Y_JOURNAL.Refresh ;
    END ;
  *)
  AffecteGrid(GSO, Action);
  PieceModifiee := False;
  AttribLeTitre;
  if PureAnal then
  begin
    LOA_TotalDebit.Visible := True;
    LOA_TotalCredit.Visible := True;
    H_MontantEcr.Visible := False;
    LOA_MontantEcr.Visible := False;
  end else
  begin
    LOA_TotalDebit.Visible := False;
    LOA_TotalCredit.Visible := False;
    BevCache.Visible := False;
    H_MontantEcr.Visible := True;
    LOA_MontantEcr.Visible := True;
    if SensGene = 1 then AfficheLeSolde(OA_MontantEcr, TotPGene, 0) else AfficheleSolde(OA_MontantEcr, 0, TotPGene);
    // GP le 22/05/2002 N°10008
    Y_ETABLISSEMENT.Enabled := FALSE;
    H_ETABLISSEMENT.Enabled := FALSE;
  end;
  if OkM then HMessPiece.Execute(4, caption, '') else
    if OkPasJal then HMessPiece.Execute(20, caption, '') else
    if OkDEV then HMessPiece.Execute(14, caption, '');
  if Action <> taCreat then GSORowEnter(nil, GSO.Row, bc, False);
  Cursor := crDefault;
  CpteGeneChange := False;
  //GP - 07/03/2002  force le redimensionnement de la grille
  {JP 13/11/07 : FQ 21320 : je ne sais pourquoi le ResizeGridColumns vait été mis en commentaires !!}
  HMTrad.ResizeGridColumns(GSO);
  //HMTrad.Resize(Self);
end;

procedure TFSaisODA.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  GSO.VidePile(True);
  SAJAL.Free;
  SAJAL := nil;
  TS.Free;
  TS := nil;
  TSOLDEGS.Free;
  TSOLDEGS := nil;
  TAX.Free;
  TAX := nil;
  PurgePopup(POPS);
  VideListe(TDELANA);
  TDELANA.Free;
  TDELANA := nil;
  RegSaveToolbarPos(Self, 'SaisODA');
  if Parent is THPanel then
  begin
    case Self.Action of
      taCreat: _Bloqueur('nrSaisieCreat', False);
      taModif: _Bloqueur('nrSaisieModif', False);
    end;
    Action := caFree;
  end;
end;

procedure TFSaisODA.FormMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  GX := X;
  GY := Y;
end;

procedure TFSaisODA.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if not GSO.SynEnabled then
  begin
    Key := #0;
    Exit;
  end;
  if Key = #127 then Key := #0 else
    if ((Key = '=') and ((GSO.Col = OA_Debit) or (GSO.Col = OA_Credit) or (GSO.Col = OA_Sect))) then Key := #0;
end;

procedure TFSaisODA.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  OkG, Vide: Boolean;
begin
  if not GSO.SynEnabled then
  begin
    Key := 0;
    Exit;
  end;
  OkG := (Screen.ActiveControl = GSO);
  Vide := (Shift = []);
  case Key of
    VK_F5:
      begin
        if ((OkG) and (Vide)) then
        begin
          Key := 0;
          ClickZoom;
        end;
        if Shift = [ssShift] then
        begin
          Key := 0;
          ClickZoomPiece;
        end;
      end;
    VK_F6: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        ClickSolde(False);
      end;
    VK_F10: if Vide then
      begin
        Key := 0;
        ClickValide;
      end;
    VK_RETURN: if ((OkG) and (Vide)) then KEY := VK_TAB;
    VK_INSERT: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        ClickInsert;
      end;
    VK_DELETE: if ((OkG) and (Shift = [ssCtrl])) then
      begin
        Key := 0;
        ClickDel;
      end;
    //SG6 13/01/05 Gestion de la fermeture de la saisie oda FQ 15158
    VK_ESCAPE: if Vide then begin key :=0; ClickAbandon; end;
    VK_BACK: if ((OkG) and (Shift = [ssCtrl])) then
      begin
        Key := 0;
        VideZone(GSO);
      end;
    {=}187: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        ClickSolde(True);
      end;
    {AC}67: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickComplement;
      end;
    {AS}83: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickSoldeGS;
      end;
    {AT}84: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickEtabl;
      end;
    {AV}86: if Shift = [ssAlt] then
      begin
        Key := 0;
        ClickVentilType;
      end;
    // {^R}     82 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickRepart ;END ;
  end;
end;

{================================== Méthodes de le GRID ==================================}
procedure TFSaisODA.GSOSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
begin
  PieceModifiee := True;
end;

procedure TFSaisODA.GSORowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
var
  OkL : Boolean;
  {b FP FQ15861 20/10/2005}
  CptPresent: Boolean;
  i:          Integer;
  numero_axe: byte;
  {e FP FQ15861 20/10/2005}
begin
  if csDestroying in ComponentState then Exit;
  if Action = taConsult then exit;

  {b FP FQ15861 20/10/2005: Dans le cas des axes croisés, on génère par avance autant de lignes qu'il y a d'axes croisés}
  if VH^.AnaCroisaxe and (Ou > 0) then begin
    CptPresent := False;
    RechercheAxe(Ou,numero_axe);
    for i := Ou - numero_axe + 1 to Ou - numero_axe + nbre_axe_ventil do begin
      if EstRempli(GSO, i) then
        CptPresent := True;
    end;
    if (not CptPresent) and (Ou - numero_axe + nbre_axe_ventil >= GSO.RowCount - 1) then Exit;
  end
  else begin
    if ((Ou >= GSO.RowCount - 1) or (Ou <= 0)) then Exit;
  end;
  {e FP FQ15861 20/10/2005}
  GridEna(GSO, False);

  OkL := LigneCorrecte(Ou, True);

  if not OkL then Cancel := True;
  GridEna(GSO, True);
end;

procedure TFSaisODA.GSORowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
  if ((not EstRempli(GSO, Ou)) and (Action <> taConsult)) then DefautLigne(Ou, True);
  Y_Journal.Enabled := ((GSO.RowCount <= 2) and (Action = taCreat));
  AffichePied;
  GereEnabled(Ou);
end;

procedure TFSaisODA.GSOCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
var
  NumAxe : Byte;{FQ 15881} 
begin
  if ((Action <> taConsult) and (not EstRempli(GSO, GSO.Row)) and (GSO.Col <> OA_Sect))
    then
  begin
    ACol := OA_Sect;
    Cancel := True;
    Exit;
  end;

  //SG6 22/12/04 Gestion du mode croisaxe
  if VH^.AnaCroisaxe then
  begin
    if EstRempli(GSO,GSO.RowCount-1) then GSO.RowCount := GSO.RowCount + nbre_axe_ventil;
    {JP 10/10/05 : FQ 15881 : On autorsie la ventilation type qu'une fois saisi le dernier axe ventilable}
    RechercheAxe(GSO.Row, NumAxe);
    BVentilType.Enabled := BVentilType.Enabled and (NumAxe = nbre_axe_ventil);
  end
  else
  begin
    GereNewLigne(GSO);
  end;

  //GereZoom ;
  if Action = taConsult then Exit;
  if GSO.Col = OA_Credit then
  begin
    if ((GSO.Cells[OA_Debit, GSO.Row] <> '') or (SensGene = 1)) then
    begin
      if ((GSO.Row = ARow) and (GSO.Cells[OA_Debit, GSO.Row] <> '')) then
      begin
        if ACol = OA_Debit then
        begin
          ARow := ARow + 1;
          ACol := OA_Sect;
        end else
        begin
          ACol := OA_Debit;
          ARow := GSO.Row;
        end;
      end else
      begin
        ACol := OA_Debit;
        ARow := GSO.Row;
      end;
      Cancel := True;
    end;
  end else if GSO.Col = OA_Debit then
  begin
    if ((GSO.Cells[OA_Credit, GSO.Row] <> '') or (Sensgene = 2)) then
    begin
      if ((GSO.Row = ARow) or (GSO.Cells[OA_Credit, GSO.Row] <> '')) then
      begin
        if ACol <= OA_Lib then ACol := OA_Credit else if ACol >= OA_Credit then ACol := OA_Lib;
      end else
      begin
        ACol := OA_Credit;
        ARow := GSO.Row;
      end;
      Cancel := TRUE;
    end;
  end;
end;

procedure TFSaisODA.GSOCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
var
  iAxe,numero_axe : byte; //SG6 22/21/2004 Gestion croisaxe
  i : integer;
  sTmp : string;
begin
  if csDestroying in ComponentState then Exit;
  if Action = taConsult then Exit;
  if GSO.Row = ARow then
  begin
    if ACol = OA_Sect then
    begin
      if LeMeme(GSO, ACol, ARow) then Exit;
      //SG6 22/12/2004
      if VH^.AnaCroisaxe then
      begin
        iAxe := RechercheAxe(GSO.Row,numero_axe);
        if iAxe = 0 then exit;
        Cache.ZoomTable := AxeToTz('A' + IntToStr(iAxe));
      end;

      if ChercheSect(ARow, False) <= 0 then
      begin
        Cancel := True;
        Exit;
      end;
    end;
  end;
  //SG6 22/12/2004 Gestion du mode croisaxe pour les debits
  if ((ACol = OA_Debit) or (ACol = OA_Credit)) then
  begin
    if VH^.AnaCroisaxe then ChercheMontantCroisAxe(ACol, ARow, GSO.Cells[Acol,ARow])
                       else ChercheMontant(Acol, ARow);
  end;

  //SG6 Gestion du libelle et reference en mode croisaxe
  if (VH^.AnaCroisaxe and ((ACol = OA_RefI) or (ACol = OA_Lib)) and (GSO.Cells[ACol,ARow]<>'')) then begin
    sTmp := GSO.Cells[ACol, ARow];
    RechercheAxe(ARow,numero_axe);
    for i := ARow - numero_axe + 1 to ARow - numero_axe + nbre_axe_ventil do begin
      GSO.Cells[ACol,i] := sTmp;
      {JP 11/10/05 : FQ 15881 : Maj de l'OBM}
      AlimObjetMvt(i);
    end;
  end;
end;

procedure TFSaisODA.GSODblClick(Sender: TObject);
begin
  ClickZoom;
end;

procedure TFSaisODA.GSOEnter(Sender: TObject);
begin
  GereEnabled(1);
end;

procedure TFSaisODA.GSOExit(Sender: TObject);
var
  b, bb: boolean;
  C, R: longint;
begin
  if csDestroying in ComponentState then Exit;
  if Action = taConsult then Exit;
  if Screen.ActiveControl = nil then Exit;
  { FQ 21060 BVE 12.07.07 }
  // if Valide97.Tag <> 1 then Exit; {Le bouton bValide n'est pas sélectionné}
  { END FQ 21060 }
  if ((Screen.ActiveControl.Parent = PEntete) and (not EstRempli(GSO, GSO.Row))) then Exit;
  if ((Screen.ActiveControl.Parent = PEntete) and (EstRempli(GSO, 1))) then
  begin
    if Screen.ActiveControl = Y_Journal then ActiveControl := GSO;
    Y_Journal.Enabled := False;
  end;
  if ((Screen.ActiveControl <> GSO) and (EstRempli(GSO, GSO.Row))) then
  begin
    C := GSO.Col;
    R := GSO.Row;
    b := False;
    bb := False;
    GSOCellExit(nil, C, R, b);
    GSORowExit(nil, R, b, bb);
  end;
end;

procedure TFSaisODA.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
  begin
    X := WMinX;
    Y := WMinY;
  end;
end;

procedure TFSaisODA.GSOKeyPress(Sender: TObject; var Key: Char);
begin
  if not GSO.SynEnabled then Key := #0;
end;

procedure TFSaisODA.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFSaisODA.BMenuZoomMouseEnter(Sender: TObject);
begin
  if not BMenuZoom.Enabled then Exit;
  PopZoom97(BMenuZoom, POPZ);
end;

procedure TFSaisODA.GereAffSolde(Sender: TObject);
var
  Nam: string;
  C: THLabel;
begin
  Nam := THNumEdit(Sender).Name;
  Nam := 'L' + Nam;
  C := THLabel(FindComponent(Nam));
  if C <> nil then C.Caption := THNumEdit(Sender).Text;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. : Fonction qui renvoie le numéro de l'axe correspondant a la
Suite ........ : ligne en cours
Mots clefs ... : NUMERO AXE, ROW, GRID
*****************************************************************}
function TFSaisODA.RechercheAxe(Lig : integer;var numero_axe : byte):byte;
var
  i, j : integer;
begin
  result := 0;
  //Determination de la position de cette axe permi tous les axes ventilés
  numero_axe := 0;
  for i := 1 to nbre_axe_ventil do
  begin
    if round((Lig-i)/(nbre_axe_ventil)) = (Lig-i)/(nbre_axe_ventil) then
    begin
      numero_axe := i;
      break;
    end;
  end;
  //Determination de l'axe en fonction du numero de l'axe dans la liste des axes ventilés
  j := 0;
  for i := 1 to MaxAxe do
  begin
    if Axes[i] then
    begin
      Inc(j);
    end;
    if j = numero_axe then
    begin
      result := i;
      break;
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 22/12/2004
Modifié le ... :   /  /    
Description .. : Fonction qui permet de la gestion de montant en mode 
Suite ........ : croisaxe
Mots clefs ... : GESTION DES MONTANTS EN MODE CROISAXE
*****************************************************************}
procedure TFSaisODA.ChercheMontantCroisAxe(Acol, ARow : longint ; montant : string);
var
  i : integer;
  numero_axe : byte;
begin
  RechercheAxe(ARow,numero_axe);
  for i:=ARow-numero_axe+1 to ARow-numero_axe + nbre_axe_ventil do
  begin
    GSO.Cells[ACol,i] := montant;
    ChercheMontant(ACol,i);
  end;
end;

procedure TFSaisODA.InsertCroisaxe(Lig : longint);
var
  i :integer;
  iAxe, numero_axe : byte;
begin
  for i := Lig to Lig - 1 + nbre_axe_ventil do
  begin
    iAxe := RechercheAxe(i,numero_axe);
    TEcrODA.PutValue('Y_SOUSPLAN'+IntToStr(iAxe),GSO.Cells[OA_Sect,i]);
  end;
  TEcrODA.InsertDB(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 23/12/2004
Modifié le ... :   /  /    
Description .. : Fonction qui permet de renvoyer la borne sup lors de 
Suite ........ : boucle
Mots clefs ... : BORNE SUP
*****************************************************************}
function TFSaisODA.Return_Bornesup:longint;
begin
  //Init de borne_sup //SG6 23/12/04
  if not VH^.AnaCroisaxe then result := GSO.Rowcount - 2
  else result := GSO.RowCount - 1 - nbre_axe_ventil;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 27/12/2004
Modifié le ... :   /  /    
Description .. : Validation des pièces en mode croisaxe
Mots clefs ... : VALIDE PIECE CROISAXE
*****************************************************************}
procedure TFSaisODA.ValideLaPieceTrans;
begin
  try
    begintrans;
    ValideLaPiece;
    Committrans;
  except
    on e : exception do
    begin
      rollback;
      showmessage(e.message);
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 27/12/2004
Modifié le ... :   /  /    
Description .. : Procédure qui permet de charger les écritures en mode 
Suite ........ : croisaxe
Mots clefs ... : CHARGEMENT ECRITURE CROISAXE
*****************************************************************}
procedure TFSaisODA.ChargeEcrituresCroisaxe;
var
  Lig : integer;
  OBM : TOBM;
  XD, XC : Double;
  JD, JC : Double;
  QODA : TQuery;
  XDEL : TDELMODIF;
  sousplan : integer;
begin
  if Action = taCreat then exit;

  BeginTrans;
  NbLigAna := 0;
  InitMove(50,'');
  Lig := 0;
  JD := 0;
  JC := 0;
  VideListe(TDELANA);
  QODA := OpenSQL('SELECT ' + GetSelectAll('Y', True) + ', Y_BLOCNOTE FROM ANALYTIQ WHERE ' + WhereEcriture(tsODA, M, False)
    + ' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE,'
    + ' Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL', True);

  while not QODA.eof do
  begin
    MoveCur(FALSE);
    {Sauvegarde des dates modif}
    XDEL := TDELMODIF.Create;
    XDEL.NumLigne := QODA.FindField('Y_NUMLIGNE').AsInteger;
    XDEL.Ax := QODA.FindField('Y_AXE').AsString;
    XDEL.NumEcheVent := QODA.FindField('Y_NUMVENTIL').AsInteger;
    XDEL.DateModification := QODA.FindField('Y_DATEMODIF').AsDateTime;
    TDELANA.Add(XDEL);

    for sousplan := 1 to 5 do
    begin
      if QODA.FindField('Y_SOUSPLAN'+IntToStr(sousplan)).AsString<>'' then
      begin

      {Lecture des infos}
      Inc(Lig);
      DefautLigne(Lig, False);
      OBM := GetO(GSO, Lig);
      OBM.ChargeMvt(QODA);
      OBM.HistoMontants;
      if OBM.GetMvt('Y_CONFIDENTIEL') > '0' then
      begin
        PieceConf := True;
        HConf.Visible := True;
      end;
      {$IFDEF EAGLCLIENT}
      // A FAIRE
      {$ELSE}
      OBM.M.Assign(TMemoField(QODA.FindField('Y_BLOCNOTE')));
      {$ENDIF}
      GSO.Cells[OA_Sect, Lig] := QODA.FindField('Y_SOUSPLAN'+IntToStr(sousplan)).AsString;
      GSO.Cells[OA_RefI, Lig] := QODA.FindField('Y_REFINTERNE').AsString;
      GSO.Cells[OA_Lib, Lig] := QODA.FindField('Y_LIBELLE').AsString;
      XD := QODA.FindField('Y_DEBIT').AsFloat;
      XC := QODA.FindField('Y_CREDIT').AsFloat;
      JD := JD + XD;
      JC := JC + XC;
      GSO.Cells[OA_Debit, Lig] := StrS(XD, V_PGI.OkDecV);
      GSO.Cells[OA_Credit, Lig] := StrS(XC, V_PGI.OkDecV);
      FormatMontant(GSO, OA_Debit, Lig, V_PGI.OkDecV);
      FormatMontant(GSO, OA_Credit, Lig, V_PGI.OkDecV);
      if Lig = 1 then
      begin
        NumLigGene := QODA.FindField('Y_NUMLIGNE').AsInteger;
        TotPGene := QODA.FindField('Y_TOTALECRITURE').AsFloat;
        TotDGene := QODA.FindField('Y_TOTALDEVISE').AsFloat;
        PureAnal := (NumLigGene <= 0);
        if not PureAnal then
        begin
          if XD <> 0 then SensGene := 1 else SensGene := 2;
        end;
      end;
      Inc(NbLigAna);

      end;
    end;
    QODA.Next;
    GSO.RowCount := GSO.RowCount + nbre_axe_ventil;
  end;
  Ferme(QODA);
  SAJAL.OldDebit := JD;
  SAJAL.OldCredit := JC;
  if not PureAnal then Y_GENERAL.Enabled := False;
  FiniMove;
  CommitTrans;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 27/12/2004
Modifié le ... :   /  /    
Description .. : Fonction qui met à jour la base en mode croisaxe pour les 
Suite ........ : écriture OD analytiques
Mots clefs ... : MAJ BD ECRITURES OD ANALYTIQUES
*****************************************************************}
procedure TFSaisODA.MiseAJourODA(Lig : integer);
var
  iAxe, i: integer;
  OBM : TOBM;
  numero_axe : byte;
  St: string;
  //    Sens : integer ;
  SQL, StConf: string;
  DD: TDateTime;
begin

  OBM := GetO(GSO, Lig);
  StConf := '0';
  if HConf.Visible then StConf := '1';
  DD := OBM.GetMvt('Y_DATEMODIF');
  //if ValD(GSO,Lig)<>0 then Sens:=1 else Sens:=2 ;
  OBM.PutMvt('Y_GENERAL', Y_GENERAL.Text);
  OBM.PutMvt('Y_QUALIFPIECE', M.Simul);
  OBM.PutMvt('Y_NATUREPIECE', M.Nature);
  if PureAnal then
  begin
    iAxe := RechercheAxe(Lig,numero_axe);
    if iAxe = 0 then Exit;


    OBM.PutMvt('Y_AXE','A'+IntToStr(iAxe));
    OBM.PutMvt('Y_TYPEANALYTIQUE', 'X');
    if M.ANouveau then
    begin
      OBM.PutMvt('Y_ECRANOUVEAU', 'OAN');
      OBM.PutMvt('Y_VALIDE', 'X');
      if OBM.GetMvt('Y_TYPEANOUVEAU') = '-' then OBM.PutMvt('Y_TYPEANOUVEAU', 'ASA');
    end;
  end else
  begin
    OBM.PutMvt('Y_TYPEANALYTIQUE', '-');
  end;
  OBM.PutMvt('Y_TYPEMVT', 'JLD'); {Forcer modif}
  OBM.PutMvt('Y_CONFIDENTIEL', StConf);
  if Lig = 1 then OBM.PutMvt('Y_TYPEMVT', 'AE')
             else OBM.PutMvt('Y_TYPEMVT', 'AL');


  //Sous plan
  for i := Lig to Lig + nbre_axe_ventil - 1 do
  begin
    iAxe := RechercheAxe(i,numero_axe);
    if iAxe = 0 then Exit;
    OBM.PutMvt('Y_SOUSPLAN'+IntToStr(iAxe),GSO.Cells[OA_Sect,i]);
  end;

  St := OBM.StPourUpdate;
  if St = '' then Exit;

  SQL := 'UPDATE ANALYTIQ SET ' + St + ', Y_DATEMODIF="' + UsTime(NowFutur) + '" Where ' + WhereEcriture(tsODA, M, False);
  SQL := SQL + ' AND Y_NUMLIGNE=' + IntToStr(NumLigGene) + ' AND Y_NUMVENTIL=' + IntToStr(GetNumLigneCroisAxe(Lig)) {FQ 15881}
    + ' AND Y_DATEMODIF="' + UsTime(DD) + '"';
  if ExecuteSQL(SQL) <= 0 then V_PGI.IOError := oeSaisie;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 27/12/2004
Modifié le ... :   /  /    
Description .. : Procédure qui permet de supprimer une ligne en mode 
Suite ........ : croisaxe
Mots clefs ... : SUPPRESSION LIGNE CROISAXE
*****************************************************************}
procedure TFSaisODA.DetruitLigneCroisaxe(Lig : integer);
var
  numero_axe : byte;
  i : integer;
  OBM : TOBM;
begin
  GSO.CacheEdit;
  if RechercheAxe(Lig,numero_axe) = 0 then Exit;
  for i := Lig-numero_axe + nbre_axe_ventil downto Lig - numero_axe + 1 do
  begin
    //On free l'object
    OBM := GetO(GSO, i);
    if OBM<> nil then FreeAndNil(OBM);
    GSO.DeleteRow(i);
  end;

  CalculDebitCredit;

  //Numeration des lignes
  for i:=1 to GSO.RowCount - 1 do
  begin
    GSO.Cells[OA_NumL,i]:= IntToStr(GetNumLigneCroisAxe(i)); //IntToStr(i);
    {b FP FQ16213}
    GSO.Cells[OA_Axe, i]  := 'A'+IntToStr(RechercheAxe(i,numero_axe));
    //RechercheAxe(i,numero_axe);
    {e FP FQ16213}
    if numero_axe = 0 then exit;
    if GetO(GSO,i)<>nil then TOBM(GetO(GSO,i)).PutMvt('Y_NUMVENTIL',IntToStr(GetNumLigneCroisAxe(i - numero_axe + 1)))
  end;

  GSO.SetFocus;
  GSO.Col := OA_Sect;
  GSO.Invalidate;
  GSO.MontreEdit;

end;

{JP 04/08/05 : FQ 15859/15651 : Les BitButtons ne prenant pas le focus, on force la sortie de la cellule
{---------------------------------------------------------------------------------------}
procedure TFSaisODA.LostFocusBeforeClick;
{---------------------------------------------------------------------------------------}
var
  C, R : Integer;
  BOK  : Boolean;
begin
  if Screen.ActiveControl = GSO then begin
    C := GSO.Col;
    R := GSO.Row;
    BOk := False;
    GSOCellExit(GSO, C, R, BOk);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TFSaisODA.VentileTypeClassique(Code : string; Col : Integer);
{---------------------------------------------------------------------------------------}
var
  Ax : string;
  Solde, XP, Pourc, TotP, TotM: Double;
  Lig, ii, borne_sup: integer; //SG6 23/12/2004 Gestion croisxe
  OBM: TOBM;
  QF1, QF2: string;
  Q : TQuery;
begin
  ii := SommeQte(QF1, QF2);
  Solde := OA_Solde.Value;
  Xp := 0;

  //SG6 23/12/2004 Mode croisaxe
  borne_sup := Return_Bornesup;
  Lig := borne_sup;

  Ax := SAJAL.Axe;
  OBM := nil;
  TotP := 0;
  TotM := 0;

  Q := OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE="TY' + SAJAL.AXE[2] + '" '
    + 'AND V_COMPTE="' + Code + '" ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL', True);
  try
    while not Q.EOF do
    begin
      Inc(Lig);
      DefautLigne(Lig, True);
      AlimObjetMvt(Lig);
      OBM := GetO(GSO, Lig);

      GSO.Cells[OA_Sect, Lig] := Q.FindField('V_SECTION').AsString;
      OBM.PutMvt('Y_SECTION', GSO.Cells[OA_Sect, Lig]);

      {Récupération et affecation des principaux champs du TOBM}
      Pourc := RemplitLignes(Q, OBM, ii, QF1, QF2);

      {Calcul du montant}
      XP := Arrondi(Pourc * Solde / 100.0, V_PGI.OkDecV);
      TotM := TotM + XP;
      TotP := TotP + Pourc;

      {Affectation du montant de la ventilation}
      AffecteMontant(OBM, Col, Lig, XP);

      GereNewLigne(GSO);
      Q.Next;
    end;

    {Gestion des arrondis}
    GereArrondi(Lig, Col, OBM, XP, TotP, TotM);
    {Termine pièce}
    FinitPiece(ii);
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TFSaisODA.VentileTypeCroiseAxe(Code : string; Col : Integer);
{---------------------------------------------------------------------------------------}
var
  Solde, XP,
  Pourc, TotP,
  TotM     : Double;
  Lig, ii  : Integer;
  OBM      : TOBM;
  QF1, QF2 : string;
  Q        : TQuery;
  n        : Integer;
  NumAxe   : Byte;
begin
  if GSO.Cells[OA_Sect, GSO.Row] = '' then begin
    HShowMessage('0;' + Caption + ';Veuillez renseigner la section.;W;O;O;O;', '', '');
    Exit;
  end;

  ii  := 0;
  Lig := GSO.Row;

  {On recherche la ligne de départ sur laquelle on va mettre les ventilatin types en contreparties}
  for n := GSO.RowCount - 1 downto GSO.Row do begin
    if GSO.Cells[OA_Sect, n] = '' then
      {Permettra de calculer le nombre de ligne vide à rajouter}
      Inc(ii)
    else begin
      {La ligne de départ est donc la ligne suivante, la première vide}
      Lig := n + 1;
      Break;
    end;
  end;

  {Ajout des éventuelles lignes vides nécessaires en fonction du nombre d'axes ventilés}
  GSO.RowCount := GSO.RowCount + nbre_axe_ventil - ii;

  ii := SommeQte(QF1, QF2);
  Solde := OA_Solde.Value;
  Xp := 0;
  OBM := nil;
  TotP := 0;
  TotM := 0;

  Q := OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE = "TY1" AND V_COMPTE = "' + Code + '" ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL', True);
  try
    {On boucle sur les lignes de la ventilation type choisie}
    while not Q.EOF do begin
      {Pour chaque ligne, on boucle sur le nombre d'axes paramétrés}
      for n := 1 to nbre_axe_ventil do begin
        {Initialisation de la ligne}
        DefautLigne(Lig, True);
        AlimObjetMvt(Lig);
        OBM := GetO(GSO, Lig);

        {Récupération du libellé et de la référence sur les premières ligne}
        GSO.Cells[OA_RefI, Lig] := GSO.Cells[OA_RefI, GSO.Row];
        OBM.PutMvt('Y_REFINTERNE', GSO.Cells[OA_RefI, GSO.Row]);
        GSO.Cells[OA_Lib, Lig] := GSO.Cells[OA_Lib, GSO.Row];
        OBM.PutMvt('Y_LIBELLE', GSO.Cells[OA_Lib, GSO.Row]);

        {Récupération de la section stockée dans V_SOUSPLAN + Numéro d'axe}
        GSO.Cells[OA_Sect, Lig] := Q.FindField('V_SOUSPLAN' + IntToStr(RechercheAxe(Lig, NumAxe))).AsString;
        OBM.PutMvt('Y_SECTION', GSO.Cells[OA_Sect, Lig]);

        {Récupération et affecation des principaux champs du TOBM}
        Pourc := RemplitLignes(Q, OBM, ii, QF1, QF2);

        {Calcul du montant}
        XP := Arrondi(Pourc * Solde / 100.0, V_PGI.OkDecV);
        TotM := TotM + XP;
        TotP := TotP + Pourc;

        {Affectation du montant de la ventilation}
        AffecteMontant(OBM, Col, Lig, XP);

        Inc(Lig);

        {Ajout des lignes vides}
        if GSO.RowCount - 1 = Lig then
          GSO.RowCount := GSO.RowCount + nbre_axe_ventil;
      end; {For : boucle sur le nombre d'axes}

      {Ventilation suivante}
      Q.Next;
    end; {While sur les lignes de ventilations types}

    {Gestion des arrondis}
    GereArrondi(Lig, Col, OBM, XP, TotP, TotM);
    {Termine pièce}
    FinitPiece(ii);
  finally
    Ferme(Q);
  end;
end;

{Renvoie le numéro de ventile en fonction du numéro de ligne dans la grille :
 Si 3 axes de paramétrer en CroisAxe => 3 Lignes dans la grille et 1 lignes dans ANALYTIQ
{---------------------------------------------------------------------------------------}
function TFSaisODA.GetNumLigneCroisAxe(Ligne : Integer) : Integer;
{---------------------------------------------------------------------------------------}
begin
  Result := Round(Int((Ligne -1) / nbre_axe_ventil)) + 1;
end;

{Gestion des arrondis, après une ventilation type
{---------------------------------------------------------------------------------------}
procedure TFSaisODA.GereArrondi(Lig, Col : Integer; var OBM : TOBM; XP, TotP, TotM : Double);
{---------------------------------------------------------------------------------------}
var
  Solde : Double;
begin
  Solde := OA_Solde.Value;

  {Si 100% alors arrondi sur dernière ligne sauf si 0}
  if ((OBM <> nil) and (Arrondi(TotP - 100.0, ADecimP) = 0) and (Arrondi(Solde - TotM, V_PGI.OkDecV) <> 0)) then
    if XP + Arrondi(Solde - TotM, V_PGI.OkDecV) <> 0 then begin
      XP := XP + Arrondi(Solde - TotM, V_PGI.OkDecV);
      if Col = OA_Debit then begin
        OBM.PutMvt('Y_DEBIT', XP);
        GSO.Cells[OA_Debit, Lig] := StrS(XP, V_PGI.OkDecV);
        FormatMontant(GSO, OA_Debit, Lig, V_PGI.OkDecV);
      end
      else begin
        OBM.PutMvt('Y_CREDIT', XP);
        GSO.Cells[OA_Credit, Lig] := StrS(XP, V_PGI.OkDecV);
        FormatMontant(GSO, OA_Credit, Lig, V_PGI.OkDecV);
      end;
    end;
end;

{Remplit les lignes à partir de la ventilation type, en CroisAxe ou non
{---------------------------------------------------------------------------------------}
function TFSaisODA.RemplitLignes(Q : TQuery; var OBM : TOBM; ii : Integer; QF1, QF2 : string) : Double;
{---------------------------------------------------------------------------------------}
var
  PQ1 : Double;
  PQ2 : Double;
begin
  {Récupération de la répartition de la ligne de ventilation}
  Result := Q.FindField('V_TAUXMONTANT').AsFloat;
  OBM.PutMvt('Y_POURCENTAGE', Result);
  PQ1 := Q.FindField('V_TAUXQTE1').AsFloat;
  OBM.PutMvt('Y_POURCENTQTE1', PQ1);
  PQ2 := Q.FindField('V_TAUXQTE2').AsFloat;
  OBM.PutMvt('Y_POURCENTQTE2', PQ2);

  {Gestion de la quantité 1}
  if ii in [0, 2] then begin
    OBM.PutMvt('Y_QTE1', Arrondi(OI_TotQ1 * PQ1 / 100.0, V_PGI.OkDecQ));
    OBM.PutMvt('Y_QUALIFQTE1', QF1);
  end;

  {Gestion de la quantité 2}
  if ii in [0, 1] then  begin
    OBM.PutMvt('Y_QTE2', Arrondi(OI_TotQ2 * PQ2 / 100.0, V_PGI.OkDecQ));
    OBM.PutMvt('Y_QUALIFQTE2', QF2);
  end;

end;

{Termine la pièce
{---------------------------------------------------------------------------------------}
procedure TFSaisODA.FinitPiece(ii : Integer);
{---------------------------------------------------------------------------------------}
begin
  ChargeSections(True);
  CalculDebitCredit;
  if ii > 0 then HMessPiece.Execute(11, caption, '');
end;

{Calcul et affectation  des montants sur les ventilations types
{---------------------------------------------------------------------------------------}
procedure TFSaisODA.AffecteMontant(var OBM : TOBM; Col, Lig: Integer; XP: Double);
{---------------------------------------------------------------------------------------}
begin
  if Col = OA_Debit then begin
    OBM.PutMvt('Y_DEBIT', XP);
    OBM.PutMvt('Y_CREDIT', 0);
    GSO.Cells[OA_Debit, Lig] := StrS(XP, V_PGI.OkDecV);
    FormatMontant(GSO, OA_Debit, Lig, V_PGI.OkDecV);
  end
  else begin
    OBM.PutMvt('Y_CREDIT', XP);
    OBM.PutMvt('Y_DEBIT', 0);
    GSO.Cells[OA_Credit, Lig] := StrS(XP, V_PGI.OkDecV);
    FormatMontant(GSO, OA_Credit, Lig, V_PGI.OkDecV);
  end;
end;

function TFSaisODA.OkJal: boolean;
begin
  // FQ 18743 SBO 18/09/2006 : Test présence d'un journal d'OD Analytique
  // Sinon propotion de création
  result := Y_JOURNAL.Values.Count > 0 ;
  if not result then
    begin
    if PgiAsk('Aucun journal d''OD Analytique n''a été paramétré. Voulez-vous le créer maintenant ?', Caption) = mrYes then
      begin
      FicheJournal(nil, '', '', taCreat, 0) ;
      Y_JOURNAL.ReLoad ;
      result := Y_JOURNAL.Values.Count > 0 ;
      end ;
    end ;
  // Si journaux présents, sélection auto du 1er de la liste
  if result then
    begin
    Y_JOURNAL.ItemIndex := 0 ;
    Y_JOURNALChange( nil ) ;
    end ;
end;

end.

