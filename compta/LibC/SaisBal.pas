unit SaisBal ;

// LockedCtrls
//   BCreateExo
//   BValEntete
//   BDelBalance
//=======================================================
// Modifications
// ~~~~~~~~~~~~~
// (x) Gestion du multi-tiers
// REPRISE DE BALANCE N
// (x) Interdire la reprise d'une balance si des écritures sont déjà saisies.
// (x) Même remarque que la suivante sur la touche ECHAP.
//     Sortie alors que balance pas équilibrée.
// (x) Le calcul du résultat et F6 ne fonctionnement pas avec le solde final
//     mais avec le solde progressif.
// BALANCE DE SITUATION
// (x) Aération des zones en décalant la date de début vers le bouton de création.
// (x) Agrandir la zone où est référencé la liste des balances de situation actives.
// (x) La touche ECHAP. même lorsqu'on se trouve dans le truc de recherche d'un compte (les jumelles)
//     sort complétement de la saisie. Pas de validation.
// (x) Possibilité d'avoir une balance déséquilibré !
// (x) Temps de chargement est long !
// (x) L'utilisation de la touche F6 ne prend que le solde progressif et pas le solde final.
// (x) Reprise des données en provenance de SISCO II au niveau des balances de situation.

//=======================================================
//=========== Questions sur la saisie balance ===========
//=======================================================
//  1. Quand on ajoute un compte, la liste est-elle triée ?  OUI
//  2. La saisie est-elle uniquement en Pivot => OUI
//  3. Après validation de la balance, peut-on ajouter, supprimer des comptes ?
//  4. Balances de situation : Peut-on saisir sur N-10 ? OUI
//  5. Voir maquette
//  6. Comptes 120 et 129 présents ? Si généré auto NON sinon OUI
//  7. Collectifs ? OUI à la fin des tiers
//  8. Vérifier le pied de la balance (Résultat en pied). Tout progressif
//  9. Création de comptes ?
// 10. Pas de plan comptable ?

//=======================================================
//======== +++++++ === +++ ??? +++ === +++++++ ==========
//=======================================================
// Balance N-1 avec AN
// 1. Si pas de Balance et pas d'AN => OK
// 2. Si pas de Balance et AN => NON
// 3. Si balance et AN => Vérification date de saisie

//=======================================================
//======== +++++++ ===   LEGENDE    === +++++++ =========
//=======================================================
// ( ) Reste à faire
// (o) Implémenté mais non testé
// (x) Implémenté et testé
//=======================================================
//======== +++++++ === +++ TODO +++ === +++++++ =========
//=======================================================
// ( ) Critères HB_ETABLISSEMENT + HB_SOCIETE

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Hctrls, Mask, ExtCtrls, ComCtrls,
  Buttons, Hspliter, Ent1, HCompte, HEnt1,
  hmsgbox, HQry, Menus, Lookup, HTB97,
{$IFDEF EAGLCLIENT}
  UtileAGL, // PrintDBGrid
{$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  PrintDBG, // PrintDBGrid
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
  CPGeneraux_TOM,
  SaisUtil,
  SaisComm,
  HStatus,
  HSysMenu,
  HFLabel,
  HPanel,
  ed_tools,
  UiUtil,
  UtilSoc,
  HXlsPas,
  TZ,         // TZF
  ZTypes,     // ROpt
  ZBalance,   // TZBalance
  ZPlan,      // TZPlan
  ZCompte,    // TZCompte
  ZTiers,     // TZTiers
  SaisTier,   // SaisieBalAux
  UtilPGI, TntStdCtrls, TntExtCtrls, TntGrids ;   // DeviseToEuro

type RParBal = record
     ParExercice,              // '005'
     ParDebPeriode,            // '01/01/1999'
     ParFinPeriode : string ;  // '31/12/1999'
     end ;

procedure SaisieBalance(TypeBal : string; Action : TActionFiche) ;
procedure ChargeSaisieBalance(Params : RParBal; Action : TActionFiche) ;

const SB_ETABL  : integer = 0 ;
      SB_NEW    : integer = 1 ;
      SB_COLL   : integer = 2 ;
      SB_TYPE   : integer = 3 ;
      SB_LETTRE : integer = 4 ;
      SB_MODER  : integer = 5 ;
      SB_GEN    : integer = 6 ;
      SB_AUX    : integer = 7 ;
      SB_LIB    : integer = 8 ;
      SB_DEBIT  : integer = 9 ;
      SB_CREDIT : integer = 10;
      SB_FIRST  : integer = 6 ;
      SB_LAST   : integer = 10;

const RC_BALNONSOLDE      =  0 ;
      RC_BADWRITE         =  1 ;
      RC_BLOQUERESEAU     =  2 ;
      RC_CUMULCLASSE      =  3 ;
      RC_CHARGEPLAN       =  4 ;
      RC_BALANCEAU        =  5 ;
      RC_PERTE            =  6 ;
      RC_BENEFICE         =  7 ;
      RC_GENINTERDIT      =  8 ;
      RC_AUXINTERDIT      =  9 ;
      RC_EXO              = 10 ;
      RC_BALN1CEXISTE     = 11 ;
      RC_NOEXO            = 12 ;
      RC_ANMODIFIE        = 13 ;
      RC_BALN0AEXISTE     = 14 ;
      RC_BALN1AEXISTE     = 15 ;
      RC_ECRSAISIES       = 16 ;
      RC_DATEINCORRECTE   = 17 ;
      RC_AVECAUX          = 18 ;
      RC_CHOIXBALANCE     = 19 ;
      RC_NOBALANCE        = 20 ;
      RC_DELBALANCE       = 21 ;
      RC_NOGENBALANCE     = 22 ;
      RC_ATTEND           = 23 ;
      RC_JALRBMODIFIE     = 24 ;
      RC_BALBDSEXIST      = 25 ;
      RC_BALN1A           = 26 ;
      RC_BALN1C           = 27 ;
      RC_BALN0A           = 28 ;
      RC_BALBDS           = 29 ;
      RC_NONEGATIF        = 30 ;
      RC_NOGRPMONTANT     = 31 ;
      RC_BALN1ANOJAL      = 32 ;
      RC_BALN0ANOJAL      = 33 ;
      RC_CREATION         = 34 ;
      RC_MODIFICATION     = 35 ;
      RC_GENERATION       = 36 ;
      RC_BALEXIT          = 37 ;
      RC_EXOV8            = 38 ;
      RC_AVECRES          = 39 ;

const NRC_BALN1CEXISTE    = '11;Balance N-1;Vous ne pouvez pas saisir de balance N-1 car il existe une balance N-1 pour comparatif;E;O;O;O;' ;
      NRC_NOEXO           = '12;Balance N-1;Aucun exercice de référence pour saisir la balance;E;O;O;O;' ;
      NRC_ANMODIFIE       = '13;Balance N-1;Les A Nouveaux ont été modifiés (Lettrage et/ou Pointage);E;O;O;O;' ;
      NRC_BALN0AEXISTE    = '14;Balance N-1;Vous ne pouvez pas saisir de balance N-1 car il existe une balance N;E;O;O;O;' ;
      NRC_BALN1AEXISTE    = '15;Balance N;Vous ne pouvez pas saisir de balance car il existe une balance N-1;E;O;O;O;' ;
      NRC_ECRSAISIES      = '16;Balance N;Des écritures ont déjà été saisies sur l''exercice;E;O;O;O;' ;
      NRC_ATTEND          = '23;Balances;Veuillez paramétrer les comptes d''attente;E;O;O;O;' ;
      NRC_BALN0MODIF      = '24;Balance N;Le journal de reprise a été modifié. Confirmez la visualisation de la balance ?;Q;YN;Y;C;' ;
      NRC_BALN1ANOJAL     = '32;Balance N-1 avec A Nouveaux;Veuillez paramétrer le journal d''ouverture;E;O;O;O;' ;
      NRC_BALN0ANOJAL     = '33;Balance N;Veuillez paramétrer le journal de reprise de balance;E;O;O;O;' ;
      NRC_NOACTIVE        = '0;Comptabilité;Vous ne pouvez pas accéder à cette fonction;W;O;O;O;' ;
      NRC_ECRSAISIESN1    = '0;Balance N-1;Des écritures ont déjà été saisies sur l''exercice;E;O;O;O;' ;
      NRC_BALN0JALNOTEXISTE = '34;Balance N;Le journal de reprise de balance n''existe pas;E;O;O;O;' ;

const MAX_ROW             = 500 ;

const WM_SAISIEBALANCE    = WM_USER+77 ;

type
  TFSaisieBalance = class(TForm)
    GS              : THGrid;
    PEntete: THPanel;
    HB_DEVISE: THValComboBox;
    H_DEVISE        : THLabel;
    PPied: THPanel;
    H_SOLDE         : THLabel;
    SA_CUMULDEBIT: THNumEdit;
    SA_TOTALCREDIT  : THNumEdit;
    SA_SOLDE        : THNumEdit;
    HBalance: THMsgBox;
    FindSais        : TFindDialog;
    Bevel2          : TBevel;
    Bevel3          : TBevel;
    Bevel4          : TBevel;
    HMTrad: THSystemMenu;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Outils: TToolbar97;
    BSolde: TToolbarButton97;
    BChercher: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    Sep97: TToolbarSep97;
    HConf: TToolbarButton97;
    LSA_SOLDE: THLabel;
    LSA_TOTALCREDIT: THLabel;
    LSA_TOTALDEBIT: THLabel;
    SA_TOTALDEBIT: THNumEdit;
    Bevel1: TBevel;
    SA_CUMULCREDIT: THNumEdit;
    Bevel5: TBevel;
    LSA_CUMULDEBIT: THLabel;
    LSA_CUMULCREDIT: THLabel;
    HLCUMUL: THLabel;
    HLabel2: THLabel;
    FLASHDEVISE: TFlashingLabel;
    DATE1: TMaskEdit;
    Bevel6: TBevel;
    HLRESULTAT: THLabel;
    LSA_RESULTAT: THLabel;
    SA_RESULTAT: THNumEdit;
    Bevel7: TBevel;
    SA_BALANCE: THNumEdit;
    HLabel1: THLabel;
    LSA_BALANCE: THLabel;
    BNew: TToolbarButton97;
    HLBAL: THLabel;
    BHide: TButton;
    BImprimer: TToolbarButton97;
    GSPrint: THGrid;
    bExport: TToolbarButton97;
    SD: TSaveDialog;
    BHideFocus: TButton;
    BValEntete: TToolbarButton97;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure HB_DEVISEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GereAffSolde(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSKeyPress(Sender: TObject; var Key: Char);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BValideClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BSoldeClick(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSExit(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BNewClick(Sender: TObject);
    procedure DATE1KeyPress(Sender: TObject; var Key: Char);
    procedure BValEnteteClick(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
    procedure PFENMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BHideFocusEnter(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure bExportClick(Sender: TObject);
  private
    memSaisie    : RDevise ;
    memPivot     : RDevise ;
    memOptions   : ROpt ;
    Comptes      : TZCompte ;
    Tiers        : TZTiers ;
    ObjBalance   : TZBalance ;
    BalCourante  : TZBalance ;
    Plan         : TZPlan ;
    PlanTiers    : TZPlan ;
    //Exos         : TZExo ;
    bFirst, bExo : Boolean ;
    GridX, GridY : Integer ;
    bFilling     : Boolean ;
    bReading     : Boolean ;
    bUsing       : Boolean ;
    bGenAuto     : Boolean ;
    iRowPerte    : LongInt ;
    iRowBenef    : LongInt ;
    bClosing     : Boolean ;
    FindFirst    : Boolean ;
    bInit        : Boolean ;
    bQuestion    : Boolean ;
    SoldeBalD    : Double ;
    SoldeBalC    : Double ;
    InitParams   : RParBal ;
    WMinX, WMinY : Integer ;

    // Fonctions ressource
    function  GetMessageRC(MessageID : integer) : string ;
    function  PrintMessageRC(MessageID : Integer) : Integer ;
    // Messages
    procedure WMGetMinMaxInfo(var Msg: TMessage) ; message WM_GETMINMAXINFO ;
    procedure GoBalance(var Message: TMessage) ; message WM_SAISIEBALANCE ;
    // Fonctions utilitaires
    procedure InitBalance ;
    procedure PosLesSoldes ;
    procedure InitPivot ;
    procedure InitSaisie ;
    function  CanCloseSaisie : Boolean ;
    procedure CloseSaisie ;
    procedure AlimenteEntete(ParExercice, ParDebPeriode, ParFinPeriode : string) ;
    procedure InitEnteteBal ;
//    procedure InitEntete ;
    procedure InitPied ;
    procedure InitGrid ;
    procedure InitTotaux ;
    procedure InfosPied ;
    procedure EnableButtons ;
    // Fontions TZF
    function  SetRowTZF(Row : LongInt) : TZF ;
    function  GetRowTZF(Row, RowCur : LongInt) : Boolean ;
    procedure SetRowBad(Row : LongInt) ;
    // Calculs
    function  GetCalcRes : Boolean ;
    procedure CalculSoldeBalance ;
    function  IsSoldeBalance : boolean ;
    // Fonctions Grid
    procedure UpdateScrollBars ;
    procedure CreateRow(Row : LongInt) ;
    procedure DeleteRow(Row : LongInt) ;
    function  NextRow(Row : LongInt) : Boolean ;
    function  IsRowValid(Row : integer; var ACol : integer) : Boolean ;
    function  GetGridSens(ACol, ARow : Integer) : Integer ;
    procedure GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
    //procedure SetGridSep(ACol, ARow : Integer ; Canvas : TCanvas ; bHaut : Boolean) ;
    //procedure SetGridBold(ACol, ARow : Integer ; Canvas : TCanvas) ;
    function  IsRowOk(Row : LongInt) : Boolean ;
    procedure SetGridEditing(bOn : Boolean) ;
    procedure SetGridOptions(Row : LongInt) ;
    procedure PreparePrintGrid(bPrint : Boolean) ;
    // SQL
    procedure WriteBalance ;
    procedure ReadBalance ;
    procedure PutGrid(Row : LongInt; bColl : Boolean) ;
    // Click boutons
    function  ValideBalance : Boolean ;
    procedure SearchClick ;
    procedure ValClick ;
    procedure DelClick ;
    procedure NewClick ;
    procedure ByeClick ;
    procedure SoldeClick(Row : longint) ;
    procedure DATE1Click ;
    //procedure HB_DATE2Click ;
    procedure SetMontants(LeTZ : TZF; XD, XC: double; Dev: RDEVISE; ModeOppose,
      Force: boolean);
  public
    nbLignes : LongInt ;
    FAction  : TActionFiche ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamDat,
  ParamSoc;


//=======================================================
//======== Point d'entrée dans la saisie balance ========
//=======================================================
function CanOpenBalance(TypeBal : string; var bVisu : Boolean; var NumPiece : Integer) : Boolean ;
var ObjBal : TZBalance ; Tiers: TZTiers ;
    //Exos : TZExo ;
    Opt : ROpt ; Dev : RDevise ;
    bExist, bExistEcrs, bExistAn : Boolean ; kt : Integer ; Jal, NumAux : string ;
begin
Result:=TRUE ;
if not VH^.ZACTIVEPFU then begin HShowMessage(NRC_NOACTIVE, '', '') ; Result:=FALSE ; Exit ; end ;
FillChar(Dev ,Sizeof(Dev), #0) ;
// Les comptes d'attente ne sont pas paramétrés correctement
if (VH^.TiersDefCli='') or (VH^.TiersDefFou='') or
   (VH^.TiersDefSal='') or (VH^.TiersDefDiv='') then
   begin HShowMessage(NRC_ATTEND, '', '') ; Result:=FALSE; Exit ; end ;
Tiers:=TZTiers.Create ;
NumAux:=VH^.TiersDefCli ;
kt:=Tiers.GetCompte(NumAux) ;
NumAux:=VH^.TiersDefFou ;
if kt>=0 then kt:=Tiers.GetCompte(NumAux) ;
NumAux:=VH^.TiersDefSal ;
if kt>=0 then kt:=Tiers.GetCompte(NumAux) ;
NumAux:=VH^.TiersDefDiv ;
if kt>=0 then kt:=Tiers.GetCompte(NumAux) ;
Tiers.Free ;
if kt<0 then begin HShowMessage(NRC_ATTEND, '', '') ; Result:=FALSE; Exit ; end ;
// Tests sur la balance N-1 avec A Nouveaux
if TypeBal='N1A' then
   begin
   // Journal des AN = Journal d'ouverture
   {$IFDEF SPEC302}
   Q:=OpenSQL('SELECT SO_JALOUVRE FROM SOCIETE', TRUE) ;
   Jal:=Q.Fields[0].AsString ;
   Ferme(Q) ;
   {$ELSE}
   Jal:=GetParamSoc('SO_JALOUVRE') ;
   {$ENDIF}
   if Jal='' then
      begin HShowMessage(NRC_BALN1ANOJAL, '', '') ; Result:=FALSE; Exit ; end ;
   // Une balance N existe-t-elle ?
   Opt.TypeBal:='N0A' ; Opt.ExoBal:=VH^.Encours.Code ;
   ObjBal:=TZBalance.Create(Dev, Opt) ; bExist:=ObjBal.Exist(FALSE) ; ObjBal.Free ;
   if bExist then
      begin HShowMessage(NRC_BALN0AEXISTE, '', '') ; Result:=FALSE; Exit ; end ;

   // L'exercice N-1 existe-t-il ?
   if VH^.Precedent.Code='' then
      begin
      //Exos:=TZExo.Create ; Exos.Load ; NbExos:=Exos.Count ; Exos.Free ;
      //if not SaisieZExo(NbExos, VH^.Encours.Deb-1) then Result:=FALSE ;
      PgiInfo('Fonction : Création d''exercice non disponible','Saisie de balance');
      Result := False;
      Exit ;
      end ;

   // Une balance N-1 pour comparatif existe-t-elle ?
   Opt.TypeBal:='N1C' ; Opt.ExoBal:=VH^.Precedent.Code ;
   ObjBal:=TZBalance.Create(Dev, Opt) ; bExist:=ObjBal.Exist(FALSE) ; ObjBal.Free ;
   if bExist then
      begin HShowMessage(NRC_BALN1CEXISTE, '', '') ; Result:=FALSE; Exit ; end ;
   // Existe-t-il un exercice N-1 sans balance N1A ?
   Opt.TypeBal:='N1A' ; Opt.ExoBal:=VH^.Precedent.Code ;
   ObjBal:=TZBalance.Create(Dev, Opt) ; bExist:=ObjBal.Exist(FALSE) ; ObjBal.Free ;
   if not bExist then
      begin
      // Des écritures existent-elles sur N-1 ?
      // (Si l'exercice a été créé par la reprise de balance, la table
      // Ecriture doit être vide !!!)
      Opt.TypeBal:='N1A' ; Opt.ExoBal:=VH^.Precedent.Code ;
      ObjBal:=TZBalance.Create(Dev, Opt) ;
      bExistEcrs:=ObjBal.ExistEcrs ; ObjBal.Free ;
      if bExistEcrs then
         begin HShowMessage(NRC_ECRSAISIESN1, '', '') ; Result:=FALSE; Exit ; end ;
      end ;
   // Les A Nouveaux générés par la balance sont-ils pointés et/ou lettrés ?
   if bExist then
      begin
      Opt.TypeBal:='N1A' ; Opt.ExoBal:=VH^.Precedent.Code ;
      ObjBal:=TZBalance.Create(Dev, Opt) ; bExistAn:=ObjBal.AnOk ; ObjBal.Free ;
      if not bExistAn then
         begin HShowMessage(NRC_ANMODIFIE, '', '') ; bVisu:=TRUE; Exit ; end ;
      end ;
   end ;
// Tests sur la balance N-1 pour comparatif
if TypeBal='N1C' then
   begin
   // L'exercice N-1 existe-t-il ?
   if VH^.Precedent.Code='' then
      begin
      //Exos:=TZExo.Create ; Exos.Load ; NbExos:=Exos.Count ; Exos.Free ;
      //if not SaisieZExo(NbExos, VH^.Encours.Deb-1) then Result:=FALSE ;
      PgiInfo('Fonction : Création d''exercice non disponible','Saisie de balance');
      Result := False;
      Exit ;
      end ;
   // Une balance N-1 pour A Nouveaux existe-t-elle ?
   Opt.TypeBal:='N1A' ; Opt.ExoBal:=VH^.Precedent.Code ;
   ObjBal:=TZBalance.Create(Dev, Opt) ; bExist:=ObjBal.Exist(FALSE) ; ObjBal.Free ;
   if bExist then
      begin HShowMessage(NRC_BALN1AEXISTE, '', '') ; Result:=FALSE; Exit ; end ;
   // Existe-t-il un exercice N-1 sans balance N1C ?
   Opt.TypeBal:='N1C' ; Opt.ExoBal:=VH^.Precedent.Code ;
   ObjBal:=TZBalance.Create(Dev, Opt) ; bExist:=ObjBal.Exist(FALSE) ; ObjBal.Free ;
   if not bExist then
      begin
      // Des écritures existent-elles sur N-1 ?
      // (Si l'exercice a été créé par la reprise de balance, la table
      // Ecriture doit être vide !!!)
      Opt.TypeBal:='N1C' ; Opt.ExoBal:=VH^.Precedent.Code ;
      ObjBal:=TZBalance.Create(Dev, Opt) ;
      bExistEcrs:=ObjBal.ExistEcrs ; ObjBal.Free ;
      if bExistEcrs then
         begin HShowMessage(NRC_ECRSAISIESN1, '', '') ; Result:=FALSE; Exit ; end ;
      end ;
   end ;
// Tests sur la balance N
if (TypeBal='N0A') then
   begin
   if VH^.JalRepBalAN='' then
      begin HShowMessage(NRC_BALN0ANOJAL, '', '') ; Result:=FALSE; Exit ; end ;
   // Est-ce que le journal de reprise balance existe bien ?
   if not Presence('JOURNAL','J_JOURNAL',VH^.JalRepBalAN) then
      begin HShowMessage(NRC_BALN0JALNOTEXISTE, '', '') ; Result:=FALSE; Exit ; end ;
   // Une balance N-1 pour A Nouveaux existe-t-elle ?
   if (VH^.Precedent.Code<>'') then
     begin
     Opt.TypeBal:='N1A' ; Opt.ExoBal:=VH^.Precedent.Code ;
     ObjBal:=TZBalance.Create(Dev, Opt) ; bExist:=ObjBal.Exist(FALSE) ; ObjBal.Free ;
     if bExist then
        begin HShowMessage(NRC_BALN1AEXISTE, '', '') ; Result:=FALSE; Exit ; end ;
     end ;
   // Existe-t-il des écritures saisies sur N ?
   Opt.TypeBal:='N0A' ; Opt.ExoBal:=VH^.Encours.Code ;
   ObjBal:=TZBalance.Create(Dev, Opt) ;
   bExistEcrs:=ObjBal.ExistEcrs ; ObjBal.Free ;
   if bExistEcrs then
      begin HShowMessage(NRC_ECRSAISIES, '', '') ; Result:=FALSE; Exit ; end ;
   // Les ecritures générées par la balance N ont-elles été modifiées en saisie ?
   Opt.TypeBal:='N0A' ; Opt.ExoBal:=VH^.Encours.Code ;
   ObjBal:=TZBalance.Create(Dev, Opt) ;
   NumPiece:=ObjBal.EcrsModif ; ObjBal.Free ;
   if NumPiece>0 then
     begin
     if (HShowMessage(NRC_BALN0MODIF, '', '')=mrYes) then bVisu:=TRUE else Result:=FALSE ;
     Exit ;
     end ;
   end ;
end ;

procedure SaisieBalance(TypeBal : string; Action : TActionFiche) ;
var Balance : TFSaisieBalance ; PP : THPanel ; bVisu : Boolean ; NumPiece : Integer ;
    ObjBal : TZBalance;
    Opt : ROpt ; bExist : Boolean ;
    RDev    : RDevise ;
begin
bVisu:=FALSE ; NumPiece:=-1 ;
if (ctxPCL in V_PGI.PGIContexte) and ((TypeBal='N1A') or (TypeBal='N1C')) then
begin
  Opt.TypeBal:=TypeBal ;
  Opt.ExoBal:=VH^.Precedent.Code ;
  ObjBal:=TZBalance.Create(RDev, Opt) ;
  bExist:=ObjBal.Exist(TRUE) ;
  ObjBal.Free ;
  if not bExist then
  begin
    PGIInfo ('La création de balance N-1 est impossible.#10#13Veuillez-vous référer à la fiche version 4.1.0.','Saisie balance');
    exit;
  end;
end;
if not CanOpenBalance(TypeBal, bVisu, NumPiece) then Exit ;
Balance:=TFSaisieBalance.Create(Application) ;
Balance.memOptions.TypeBal:=TypeBal ;
Balance.memOptions.bVisu:=bVisu ;
Balance.memOptions.NumPiece:=NumPiece ;
Balance.bInit:=FALSE ;
Balance.FAction:=Action ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     Balance.ShowModal ;
   finally
     Balance.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(Balance, PP) ;
  Balance.Show ;
  end ;
end ;

procedure ChargeSaisieBalance(Params : RParBal; Action : TActionFiche) ;
var Balance : TFSaisieBalance ; PP : THPanel ;
begin
Balance:=TFSaisieBalance.Create(Application) ;
Balance.memOptions.TypeBal:='BDS' ;
Balance.memOptions.bVisu:=FALSE ;
Balance.memOptions.NumPiece:=-1 ;
Balance.bInit:=TRUE ;
Balance.FAction:=Action ;
FillChar(Balance.InitParams, Sizeof(Balance.InitParams), #0) ;
Balance.InitParams:=Params ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     Balance.ShowModal ;
   finally
     Balance.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(Balance, PP) ;
  Balance.Show ;
  end ;
end ;

//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFSaisieBalance.GetMessageRC(MessageID : Integer) : string ;
begin
Result:=HBalance.Mess[MessageID] ;
end ;

function TFSaisieBalance.PrintMessageRC(MessageID : Integer) : Integer ;
begin
Result:=HBalance.Execute(MessageID, Caption, '') ;
end ;

//=======================================================
//================== Fonctions Message ==================
//=======================================================
procedure TFSaisieBalance.WMGetMinMaxInfo(var Msg: TMessage) ;
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X:=WMinX ; Y:=WMinY ; end ;
end;

procedure TFSaisieBalance.GoBalance(var Message: TMessage) ;
begin
if bClosing then Exit ;
BValEntete.Enabled:=FALSE ;
Application.ProcessMessages ;
ReadBalance ;
if (memOptions.TypeBal<>'N0A') and (Message.LParam=1) then memOptions.bCalcRes:=GetCalcRes ;
GS.Enabled:=TRUE ; GS.SetFocus ; GSEnter(nil) ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TFSaisieBalance.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height div 2 ;
ObjBalance:=nil ; BalCourante:=nil ;
Plan:=nil ;       PlanTiers:=nil ;
bClosing:=FALSE ; bQuestion:=TRUE ;
//Exos:=nil ; Exos:=TZExo.Create ; Exos.Load ;
bFilling:=FALSE ; bReading:=FALSE ; bUsing:=FALSE ;
iRowPerte:=-1 ; iRowBenef:=-1 ;
memOptions.PlanBal:='BAL' ; bGenAuto:=FALSE ;
memOptions.bLettrage:=FALSE ;
// Paramètres de saisie
//memParams.bEquilibre:=FALSE ; memParams.bAnalytique:=FALSE ;
//memParams.bEcheance:=FALSE ;  memParams.bLibre:=FALSE ;
// Paramètres DEBUG
//memParams.bPiece:=FALSE ;     memParams.bSoldeProg:=FALSE ;
InitPivot ;
InitGrid ;
PosLesSoldes ;
end;

procedure TFSaisieBalance.FormShow(Sender: TObject);
var ObjBal : TZBalance ; Opt : ROpt ; bExist : Boolean ;
begin
//InitSaisie(FALSE) ;
bExist:=FALSE ;
LookLesDocks(Self) ;
InitEnteteBal ;
if memOptions.TypeBal='N0A' then
  begin
  Opt.TypeBal:='N0A' ;
  Opt.ExoBal:=VH^.Encours.Code ;
  ObjBal:=TZBalance.Create(memSaisie, Opt) ;
  bExist:=ObjBal.Exist(TRUE) ;
  if bExist then
    begin
    DecodeDate(ObjBal.GetDateDeb, memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour);
    DecodeDate(ObjBal.GetDateFin, memOptions.Year,    memOptions.Month,    memOptions.MaxJour);
    DATE1.Text:=DateToStr(ObjBal.GetDateDeb) ; DATE1.Enabled:=FALSE ;
    (* PFU: DEBUT BAL *)
    //memOptions.WithAux:=ObjBal.GetAux ;
    memOptions.WithAux:=FALSE ;
    (* PFU: FIN BAL *)
    if memOptions.WithAux then memOptions.PlanBal:='AUX'
                          else memOptions.PlanBal:='BAL' ;
    end else
    begin
    (* PFU: DEBUT BAL *)
    //memOptions.WithAux:=(PrintMessageRC(RC_AVECAUX)=mrYes) ;
    memOptions.WithAux:=FALSE ;
    (* PFU: FIN BAL *)
//    memOptions.bCalcRes:=(PrintMessageRC(RC_AVECRES)=mrYes) ;
    memOptions.bCalcRes:=FALSE ;
    end ;
  if memOptions.WithAux then memOptions.PlanBal:='AUX'
                        else memOptions.PlanBal:='BAL' ;
  ObjBal.Free ;
  end ;
if (memOptions.TypeBal='N1A') or (memOptions.TypeBal='N1C') then
  begin
  Opt.TypeBal:=memOptions.TypeBal ;
  Opt.ExoBal:=VH^.Precedent.Code ;
  ObjBal:=TZBalance.Create(memSaisie, Opt) ;
  bExist:=ObjBal.Exist(TRUE) ;
  if bExist then (* PFU: DEBUT BAL *)
                 //memOptions.WithAux:=ObjBal.GetAux
                 memOptions.WithAux:=FALSE
                 (* PFU: FIN BAL *)
            else begin
            (* PFU: DEBUT BAL *)
            //memOptions.WithAux:=(PrintMessageRC(RC_AVECAUX)=mrYes) ;
            memOptions.WithAux:=FALSE ;
            (* PFU: FIN BAL *)
            memOptions.bCalcRes:=(PrintMessageRC(RC_AVECRES)=mrYes) ;
            end ;
  if memOptions.WithAux then memOptions.PlanBal:='AUX'
                        else memOptions.PlanBal:='BAL' ;
  ObjBal.Free ;
  end ;
HB_DEVISE.Enabled:=FALSE ;
if bInit then AlimenteEntete(InitParams.ParExercice, InitParams.ParDebPeriode, InitParams.ParFinPeriode) ;
if (bInit) or (memOptions.TypeBal='N1A') or (memOptions.TypeBal='N1C') or (bExist) then
   PostMessage(Handle, WM_SAISIEBALANCE, 0, LParam(bExist)) ;
end;

procedure TFSaisieBalance.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//if not CanCloseSaisie then begin Action:=caNone ; Exit ; end ;
if bReading then begin Action:=caNone ; Exit ; end ;
CloseSaisie ;
//if Exos<>nil  then begin Exos.Free ;  Exos:=nil ;  end ;
if Parent is THPanel then Action:=caFree ;
if FindSais.Handle<>0 then FindSais.CloseDialog ;
ObjBalance.Free ;
bClosing:=TRUE ;
end;

procedure TFSaisieBalance.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if FAction=taConsult then begin CanClose:=TRUE ; Exit ; end ;
CanClose:=CanCloseSaisie ;
if (CanClose) and (bQuestion) then
  if (PrintMessageRC(RC_BALEXIT)<>mrYes)
    then begin CanClose:=FALSE ; Exit ; end
    else Exit ;
(*
if (CanClose) and (not IsSoldeBalance)
  if (PrintMessageRC(RC_BALNONSOLDEEXIT)<>mrYes)
    then begin CanClose:=FALSE ; Exit ; end
    else Exit ;
*)    
end;

procedure TFSaisieBalance.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Vide : Boolean ;
begin
if bReading then begin Key:=0 ; Exit ; end ;
if not GS.SynEnabled then begin Key:=0 ; Exit ; end ;
(* PFU : Pour tester une fonction *)
if (Key=VK_F10) and (ssAlt in Shift) and (ssCtrl in Shift) then
   begin
   end ;
if not (Screen.ActiveControl=GS) then Exit ;
Vide:=(Shift=[]) ;
case Key of
  VK_RETURN : if (Vide) then Key:=VK_TAB ;
     VK_END : if Shift=[ssCtrl] then
                 begin
                 Key:=0 ; SetGridOptions(nbLignes) ;
                 GS.Row:=nbLignes ; GS.Col:=SB_CREDIT ;
                 PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
                 PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
                 end ;
  VK_INSERT : if (Vide) then begin Key:=0; end ;
  VK_DELETE : if Shift=[ssCtrl] then begin Key:=0 ; DelClick ; end ;
      VK_F5 : if (Vide) then begin Key:=0 ; GSDblClick(nil); end ;
      VK_F6 : if (Vide) then begin Key:=0 ; SoldeClick(-1) ; end ;
  VK_ESCAPE,
     VK_F10 : if (Vide) then
                 begin
                 Key:=0 ;
                 if FindSais.Handle<>0 then begin FindSais.CloseDialog ; Exit ; end ;
                 if BValEntete.Enabled then BValEnteteClick(nil) else ValClick ;
                 end ;
  {CTRL+F} 70 : if Shift=[ssCtrl] then begin if bReading then Exit ; SearchClick ; end ;
  end ;
end;

//=======================================================
//================ Evénements des contrôles =============
//=======================================================
procedure TFSaisieBalance.HB_DEVISEChange(Sender: TObject);
begin
memSaisie.Code:=HB_DEVISE.Value ;
if memSaisie.Code<>'' then GetInfosDevise(memSaisie) ;
memSaisie.Taux:=GetTaux(memSaisie.Code,memSaisie.DateTaux,EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
ChangeMask(SA_RESULTAT,    memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_TOTALDEBIT,  memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_TOTALCREDIT, memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_CUMULDEBIT,  memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_CUMULCREDIT, memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_SOLDE,       memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_BALANCE,     memSaisie.Decimale, memSaisie.Symbole) ;
end;

procedure TFSaisieBalance.DATE1KeyPress(Sender: TObject; var Key: Char) ;
begin ParamDate(Self,Sender,Key) ; end ;

procedure TFSaisieBalance.GereAffSolde(Sender: TObject);
var Nam : string ; c : THLabel ;
begin
Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
c:=THLabel(FindComponent(Nam)) ;
if c<>nil then c.Caption:=THNumEdit(Sender).Text ;
end;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
procedure TFSaisieBalance.InitBalance ;
begin
CloseSaisie ;
InitPivot ;
InitGrid ;
PosLesSoldes ;
InitSaisie ;
end ;

procedure TFSaisieBalance.InitSaisie ;
begin
nbLignes:=0 ; bExo:=FALSE ; iRowPerte:=-1 ; iRowBenef:=-1 ;
Comptes:=TZCompte.Create ; Tiers:=TZTiers.Create ;
Plan:=TZPlan.Create ; PlanTiers:=TZPlan.Create ;
//InitEntete ;
InitPied ;
end ;

function TFSaisieBalance.CanCloseSaisie : Boolean ;
var Cancel, Chg : Boolean ; ACol, ARow : LongInt ;
begin
Result:=TRUE ;
if bReading then begin Result:=FALSE ; Exit ; end ;
if Screen.ActiveControl=GS then
  begin
  ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ; Chg:=FALSE ;
  GSCellExit(GS, ACol, ARow, Cancel) ;
  if Cancel then begin Result:=FALSE ; Exit ; end ;
  GSRowExit(GS, ARow, Cancel, Chg) ;
  if Cancel then begin Result:=FALSE ; Exit ; end ;
  end ;
//if not IsSoldeBalance then
//   begin PrintMessageRC(RC_BALNONSOLDE) ; Result:=FALSE ; end ;
end ;

procedure TFSaisieBalance.CloseSaisie ;
begin
if Comptes<>nil   then begin Comptes.Free ;   Comptes:=nil ;   end ;
if Tiers<>nil     then begin Tiers.Free ;     Tiers:=nil ;     end ;
if Plan<>nil      then begin Plan.Free ;      Plan:=nil ;      end ;
if PlanTiers<>nil then begin PlanTiers.Free ; PlanTiers:=nil ; end ;
end ;

procedure TFSaisieBalance.AlimenteEntete(ParExercice, ParDebPeriode, ParFinPeriode : string) ;
begin
// Renseigner les options
memOptions.DecalAno:=0 ;
memOptions.WithAux:=FALSE ;
memOptions.WithRes:=TRUE ;
memOptions.bNoBilan:=FALSE ;
memOptions.bCalcRes:=FALSE ;
end ;

procedure TFSaisieBalance.InitEnteteBal ;
var ObjBal : TZBalance ; Opt : ROpt ;
begin
if bClosing then Exit ;
// Les boutons fonctionnels
BSolde.Enabled:=FALSE ;  BChercher.Enabled:=FALSE ; BNew.Enabled:=FALSE ;
BValide.Enabled:=FALSE ; BImprimer.Enabled:=FALSE ; BExport.Enabled:=FALSE ;
// Remplir la combo Période
memOptions.Exo:=VH^.Entree.Code ;
// Initialisations suivant le type de balance
if memOptions.TypeBal='N1A' then
   begin
   DecodeDate(VH^.Precedent.Fin, memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour) ;
   DecodeDate(VH^.Precedent.Fin, memOptions.Year,    memOptions.Month,    memOptions.MaxJour) ;
   memOptions.ExoBal:=VH^.Precedent.Code ;
   memOptions.ExoAno:=VH^.Encours.Code ;
   memOptions.DecalAno:=1 ;
(* PFU: DEBUT BAL *)
   //memOptions.WithAux:=TRUE ;
   memOptions.WithAux:=FALSE ;
(* PFU: FIN BAL *)
   memOptions.WithRes:=TRUE ;
   memOptions.bNoBilan:=FALSE ;
   memOptions.bCalcRes:=TRUE ;
   // Contrôles
   PEntete.Visible:=TRUE ;
   DATE1.Text:=DateToStr(VH^.Precedent.Fin) ; DATE1.Enabled:=FALSE ;
   Caption:=GetMessageRC(RC_BALN1A) ;
   HelpContext:=7710000 ;
   end ;
if memOptions.TypeBal='N1C' then
   begin
(* PFU: DEBUT BAL *)
   //memOptions.WithAux:=TRUE ;
   memOptions.WithAux:=FALSE ;
(* PFU: FIN BAL *)
   memOptions.WithRes:=FALSE ;
   // Si l'exercice précédent est clôs et mouvementé, modif. Charges/Produits
   Opt.TypeBal:='N1C' ; Opt.ExoBal:=VH^.Precedent.Code ;
   ObjBal:=TZBalance.Create(memSaisie, memOptions) ;
   if ObjBal.ExistEcrs then memOptions.bNoBilan:=TRUE
                       else memOptions.bNoBilan:=FALSE ;
   ObjBal.Free ;
   DecodeDate(VH^.Precedent.Fin, memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour) ;
   DecodeDate(VH^.Precedent.Fin, memOptions.Year,    memOptions.Month,    memOptions.MaxJour) ;
   memOptions.ExoBal:=VH^.Precedent.Code ;
   memOptions.ExoAno:=VH^.Encours.Code ;
   memOptions.DecalAno:=0 ;
   memOptions.bCalcRes:=FALSE ;
   // Contrôles
   PEntete.Visible:=TRUE ;
   DATE1.Text:=DateToStr(VH^.Precedent.Fin) ; DATE1.Enabled:=FALSE ;
   Caption:=(*GetMessageRC(RC_MODIFICATION)+' '+*)GetMessageRC(RC_BALN1C) ;
   HelpContext:=7711000 ;
   end ;
if memOptions.TypeBal='N0A' then
   begin
   DecodeDate(VH^.Encours.Deb, memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour) ;
   DecodeDate(VH^.Encours.Deb, memOptions.Year,    memOptions.Month,    memOptions.MaxJour) ;
(* PFU: DEBUT BAL *)
   //memOptions.WithAux:=TRUE ;
   memOptions.WithAux:=FALSE ;
(* PFU: FIN BAL *)
   memOptions.WithRes:=TRUE ;
   memOptions.bNoBilan:=FALSE ;
   memOptions.ExoBal:=VH^.Encours.Code ;
   memOptions.ExoAno:=VH^.Encours.Code ;
   memOptions.DecalAno:=0 ;
   memOptions.bCalcRes:=FALSE ;
   // Contrôles
   PEntete.Visible:=TRUE ;
   // Bouton spécial
   BHide.Visible:=TRUE ; BHide.Width:=0 ;
   DATE1.Text:=DateToStr(VH^.Encours.Deb) ;
   DATE1.SetFocus ;
   Caption:=(*GetMessageRC(RC_MODIFICATION)+' '+*)GetMessageRC(RC_BALN0A) ;
   HelpContext:=7716000 ;
   end ;

// Quel Exo ?
if memOptions.ExoBal=VH^.Encours.Code then memOptions.TypeExo:=teEncours else
   if memOptions.ExoBal=VH^.Suivant.Code then memOptions.TypeExo:=teSuivant else
      memOptions.TypeExo:=tePrecedent ;
// Positionnement par défaut sur la devise pivot
HB_DEVISE.Value:=V_PGI.DevisePivot ;
// Caption
UpdateCaption(Self) ;
end ;

procedure TFSaisieBalance.InitTotaux ;
begin
end ;

procedure TFSaisieBalance.InitPied ;
begin
InfosPied ;
InitTotaux ;
ZeroBlanc(PPied) ;
end ;

procedure TFSaisieBalance.PosLesSoldes ;
begin
LSA_RESULTAT.SetBounds(SA_RESULTAT.Left,SA_RESULTAT.Top,SA_RESULTAT.Width,SA_RESULTAT.Height) ;
LSA_SOLDE.SetBounds(SA_SOLDE.Left,SA_SOLDE.Top,SA_SOLDE.Width,SA_SOLDE.Height) ;
LSA_BALANCE.SetBounds(SA_BALANCE.Left,SA_BALANCE.Top,SA_BALANCE.Width,SA_BALANCE.Height) ;
LSA_TOTALDEBIT.SetBounds(SA_TOTALDEBIT.Left,SA_TOTALDEBIT.Top,SA_TOTALDEBIT.Width,SA_TOTALDEBIT.Height) ;
LSA_TOTALCREDIT.SetBounds(SA_TOTALCREDIT.Left,SA_TOTALCREDIT.Top,SA_TOTALCREDIT.Width,SA_TOTALCREDIT.Height) ;
LSA_CUMULDEBIT.SetBounds(SA_CUMULDEBIT.Left,SA_CUMULDEBIT.Top,SA_CUMULDEBIT.Width,SA_CUMULDEBIT.Height) ;
LSA_CUMULCREDIT.SetBounds(SA_CUMULCREDIT.Left,SA_CUMULCREDIT.Top,SA_CUMULCREDIT.Width,SA_CUMULCREDIT.Height) ;
SA_RESULTAT.Visible:=FALSE ;   SA_SOLDE.Visible:=FALSE ;       SA_BALANCE.Visible:=FALSE ;
SA_TOTALDEBIT.Visible:=FALSE ; SA_TOTALCREDIT.Visible:=FALSE ;
SA_CUMULDEBIT.Visible:=FALSE ; SA_CUMULCREDIT.Visible:=FALSE ;
end ;

procedure TFSaisieBalance.InitGrid ;
begin
bFirst:=TRUE ;
// Avant le VidePile pour bien placer le curseur
GS.Row:=GS.FixedRows ; GS.Col:=SB_DEBIT ;
GS.VidePile(FALSE)  ;
GS.RowCount:=MAX_ROW ; GS.TypeSais:=tsFolio ; GSPrint.RowCount:=2 ;
GS.Enabled:=FALSE ; UpdateScrollBars ;
GS.ColWidths[SB_ETABL]:=0 ;
GS.ColWidths[SB_NEW]:=0 ;
GS.ColWidths[SB_COLL]:=0 ;
GS.ColWidths[SB_TYPE]:=0 ;
GS.ColWidths[SB_LETTRE]:=0 ;
GS.ColWidths[SB_MODER]:=0 ;
GS.ColAligns[SB_DEBIT]:=taRightJustify ;
GS.ColAligns[SB_CREDIT]:=taRightJustify ;
GSPrint.ColAligns[SB_DEBIT-SB_FIRST]:=taRightJustify ;
GSPrint.ColAligns[SB_CREDIT-SB_FIRST]:=taRightJustify ;
GSPrint.ColWidths[SB_AUX-SB_FIRST]:=GSPrint.ColWidths[SB_AUX]-10 ;
GSPrint.ColWidths[SB_GEN-SB_FIRST]:=GSPrint.ColWidths[SB_GEN]-10 ;
GSPrint.ColWidths[SB_LIB-SB_FIRST]:=GSPrint.ColWidths[SB_LIB]+20 ;
GS.GetCellCanvas:=GetCellCanvas ;
GS.PostDrawCell:=PostDrawCell ;
GS.Refresh ;
end ;

procedure TFSaisieBalance.InitPivot ;
begin
memPivot.Code:=V_PGI.DevisePivot ;
if memPivot.Code<>'' then GetInfosDevise(memPivot) ;
end ;

procedure TFSaisieBalance.EnableButtons ;
var bEnabled : Boolean ;
begin
BNew.Enabled:=TRUE ;
// La balance peut être validée ?
bEnabled:=IsSoldeBalance ;
// Entête
//HB_DATE1.Enabled:=bEnabled ;

// GCO - 09/09/2004 - FQ 13264
BValide.Enabled:=bEnabled and (not memOptions.bVisu);

// Boutons
BSolde.Enabled:=(not bEnabled) and ((GS.Cells[SB_COLL, GS.Row]<>'O') or (not memOptions.WithAux))
                and (GoEditing in GS.Options) ;
// Bouton rechercher
BChercher.Enabled:=(nbLignes<>0) ;
// Bouton imprimer
BImprimer.Enabled:=bEnabled ;
BExport.Enabled:=bEnabled ;
end ;

procedure TFSaisieBalance.InfosPied ;
begin
end ;

procedure TFSaisieBalance.SetGridEditing(bOn : Boolean) ;
begin
  if bOn then
  begin
    GS.Options:=GS.Options - [GoRowSelect] ;
    GS.Options:=GS.Options + [GoEditing,GoTabs,GoAlwaysShowEditor] ;
  end
  else
  begin
    GS.Options:=GS.Options - [GoEditing,GoTabs,GoAlwaysShowEditor] ;
    GS.Options:=GS.Options + [GoRowSelect] ;
  end ;
end ;

procedure TFSaisieBalance.SetGridOptions(Row : LongInt) ;
var c : string ;
begin
  if memOptions.bVisu then
  begin
    SetGridEditing(FALSE) ;
    Exit ;
  end ;

  if ((Valeur(GS.Cells[SB_DEBIT, Row])<>0) and
     (Valeur(GS.Cells[SB_DEBIT, Row])<VH^.GrpMontantMin) or
     (Valeur(GS.Cells[SB_DEBIT, Row])>VH^.GrpMontantMax))
     or
     ((Valeur(GS.Cells[SB_CREDIT, Row])<>0) and
     (Valeur(GS.Cells[SB_CREDIT, Row])<VH^.GrpMontantMin) or
     (Valeur(GS.Cells[SB_CREDIT, Row])>VH^.GrpMontantMax)) then
  begin
    SetGridEditing(FALSE) ;
    Exit ;
  end ;

if (memOptions.TypeBal<>'N0A') and (memOptions.TypeBal<>'BDS') then
  begin
  if memOptions.bCalcRes then
     begin
     SetGridEditing(TRUE) ;
     if GS.Cells[SB_GEN, Row]=Plan.OuvrePerte then
        begin SetGridEditing(FALSE) ; Exit ; end ;
     if GS.Cells[SB_GEN, Row]=Plan.OuvreBenef then
        begin SetGridEditing(FALSE) ; Exit ; end ;
  //   if GS.Cells[SB_GEN, Row]=Plan.FermePerte then
  //      begin SetGridEditing(FALSE) ; Exit ; end ;
  //   if GS.Cells[SB_GEN, Row]=Plan.FermeBenef then
  //      begin SetGridEditing(FALSE) ; Exit ; end ;
     end else
     begin
     SetGridEditing(TRUE) ;
     if (GS.Cells[SB_TYPE, Row]='CHA') or (GS.Cells[SB_TYPE, Row]='PRO') then
        begin SetGridEditing(FALSE) ; Exit ; end ;
     end ;
  end ;
  
if memOptions.bNoBilan then
   begin
   c:=GS.Cells[SB_TYPE, Row] ;
   if (c<>'CHA') and (c<>'PRO') and (c<>'EXT') then SetGridEditing(FALSE)
                                               else SetGridEditing(TRUE) ;
   end ;
if GS.Cells[SB_COLL, Row]='O' then
  begin
  if memOptions.WithAux then SetGridEditing(FALSE)
                        else SetGridEditing(TRUE) ;
  end else SetGridEditing(TRUE) ;
end ;

//=======================================================
//=================== Gestion des TZF ===================
//=======================================================
procedure TFSaisieBalance.SetRowBad(Row : LongInt) ;
var Ecr : TZF ;
begin
Ecr:=ObjBalance.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then Exit ;
Ecr.PutValue('BADROW', 'X') ;
end ;

function TFSaisieBalance.SetRowTZF(Row : longint) : TZF ;
var Ecr : TZF ;
begin
Ecr:=ObjBalance.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then
begin
  Result:=nil ;
  Exit ;
end ;
// Entête / Commun
Ecr.PutValue('HB_TYPEBAL',       memOptions.TypeBal) ;
Ecr.PutValue('HB_PLAN',       memOptions.PlanBal) ;
Ecr.PutValue('HB_EXERCICE',   memOptions.ExoBal) ;
if Ecr.GetValue('HB_DEVISE')='' then Ecr.PutValue('HB_DEVISE', memPivot.Code) ;
SetMontants(Ecr, Valeur(GS.Cells[SB_DEBIT, Row]), Valeur(GS.Cells[SB_CREDIT, Row]),
                       memSaisie, FALSE, FALSE) ;
Ecr.PutValue('HB_DATE1', EncodeDate(memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour)) ;
Ecr.PutValue('HB_DATE2', EncodeDate(memOptions.Year,    memOptions.Month,    memOptions.MaxJour)) ;
// Corps
Ecr.PutValue('HB_COMPTE1',       GS.Cells[SB_GEN,    Row]) ;
Ecr.PutValue('HB_COMPTE2',       GS.Cells[SB_AUX,    Row]) ;
Ecr.PutValue('HB_ETABLISSEMENT', GS.Cells[SB_ETABL,  Row]) ;
// Champs sup
Ecr.PutValue('TYPEGENE',         GS.Cells[SB_TYPE,   Row]) ;
Ecr.PutValue('LIBELLE',          GS.Cells[SB_LIB,    Row]) ;
Ecr.PutValue('LETTRABLE',        GS.Cells[SB_LETTRE, Row]) ;
Ecr.PutValue('MODEREGL',         GS.Cells[SB_MODER,  Row]) ;
Result:=Ecr ;
end ;

function TFSaisieBalance.GetRowTZF(Row, RowCur : LongInt) : Boolean ;
var Ecr : TZF ; i, kt : integer ; NumCompte, NumAux : string ;
    EcrDebit, EcrCredit : Double ;
begin
if RowCur<0 then Ecr:=ObjBalance.GetRow(Row)
            else Ecr:=BalCourante.GetRow(RowCur) ;
if Ecr=nil then begin Result:=FALSE ; Exit ; end ;
// Ajustement
Row:=Row+GS.FixedRows ;
// Entête / Commun
// Corps
if (RowCur<0) or (not BalCourante.Ecr) then
  begin
  NumCompte:=Ecr.GetValue('HB_COMPTE1') ;
  GS.Cells[SB_GEN, Row]:=NumCompte ; i:=Plan.FindCompte(NumCompte) ;
  GS.Cells[SB_LIB, Row]:=Plan.GetValue('G_LIBELLE', i) ;
  GS.Cells[SB_TYPE, Row]:=Plan.GetValue('G_NATUREGENE', i) ;
  if (Plan.GetValue('G_LETTRABLE', i)='X')
    then GS.Cells[SB_LETTRE, Row]:='X'
    else GS.Cells[SB_LETTRE, Row]:='-' ;
  if (Plan.GetValue('G_POINTABLE', i)='X')
    then GS.Cells[SB_LETTRE, Row]:=GS.Cells[SB_LETTRE, Row]+'X'
    else GS.Cells[SB_LETTRE, Row]:=GS.Cells[SB_LETTRE, Row]+'-' ;
  if VarType(Plan.GetValue('G_MODEREGLE', i))<>VarNull
    then GS.Cells[SB_MODER, Row]:=Plan.GetValue('G_MODEREGLE', i)
    else GS.Cells[SB_MODER, Row]:='' ;
  NumAux:=Ecr.GetValue('HB_COMPTE2') ;
  if NumAux<>'' then
    begin
    kt:=Tiers.GetCompte(NumAux) ;
    if (kt>=0) and (not memOptions.WithAux) then
      begin
      if (NumAux=VH^.TiersDefCli) or (NumAux=VH^.TiersDefFou) or
         (NumAux=VH^.TiersDefSal) or (NumAux=VH^.TiersDefDiv) then
        begin
        if (Tiers.GetValue('T_LETTRABLE', kt)='X')
          then GS.Cells[SB_LETTRE, Row]:='X'
          else GS.Cells[SB_LETTRE, Row]:='-' ;
        if VarType(Tiers.GetValue('T_MODEREGLE', kt))<>VarNull
          then GS.Cells[SB_MODER, Row]:=Tiers.GetValue('T_MODEREGLE', kt)
          else GS.Cells[SB_MODER, Row]:='' ;
        end ;
      end ;
    GS.Cells[SB_AUX, Row]:=NumAux ;
    end ;
  EcrDebit:=Arrondi(Ecr.GetValue('HB_DEBIT'), memSaisie.Decimale) ;
  EcrCredit:=Arrondi(Ecr.GetValue('HB_CREDIT'), memSaisie.Decimale) ;
  if EcrDebit<>0 then GS.Cells[SB_DEBIT, Row]:=StrFMontant(EcrDebit,15,memSaisie.Decimale,'',TRUE)
                 else GS.Cells[SB_DEBIT, Row]:='' ;
  if EcrCredit<>0 then GS.Cells[SB_CREDIT, Row]:=StrFMontant(EcrCredit,15,memSaisie.Decimale,'',TRUE)
                  else GS.Cells[SB_CREDIT, Row]:='' ;
  GS.Cells[SB_ETABL,Row]:=Ecr.GetValue('HB_ETABLISSEMENT') ;
  end else
  begin
  NumCompte:=Ecr.GetValue('E_GENERAL') ;
  GS.Cells[SB_GEN, Row]:=NumCompte ; i:=Plan.FindCompte(NumCompte) ;
  GS.Cells[SB_LIB, Row]:=Plan.GetValue('G_LIBELLE', i) ;
  GS.Cells[SB_TYPE, Row]:=Plan.GetValue('G_NATUREGENE', i) ;
  NumAux:=Ecr.GetValue('E_AUXILIAIRE') ;
  if NumAux<>'' then begin Tiers.GetCompte(NumAux) ; GS.Cells[SB_AUX, Row]:=NumAux ; end ;
  EcrDebit:=Arrondi(Ecr.GetValue('E_DEBIT'), memSaisie.Decimale) ;
  EcrCredit:=Arrondi(Ecr.GetValue('E_CREDIT'), memSaisie.Decimale) ;
  if EcrDebit<>0 then GS.Cells[SB_DEBIT, Row]:=StrFMontant(EcrDebit,15,memSaisie.Decimale,'',TRUE) ;
  if EcrCredit<>0 then GS.Cells[SB_CREDIT, Row]:=StrFMontant(EcrCredit,15,memSaisie.Decimale,'',TRUE) ;
  GS.Cells[SB_ETABL,Row]:=Ecr.GetValue('E_ETABLISSEMENT') ;
  end ;
Result:=TRUE ;
end ;

//=======================================================
//==================== Fonctions SQL ====================
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : P.FUGIER
Créé le ...... :   /  /
Modifié le ... : 09/09/2004
Description .. :
Suite ........ : GCO - 09/09/2004 - FQ 13264
Mots clefs ... :
*****************************************************************}
procedure TFSaisieBalance.WriteBalance ;
var i : integer ;
begin
  CalculSoldeBalance ;
  // GCO - 09/09/2004 - FQ 13264
  if not memOptions.NewObj then
  begin
    ObjBalance.Del ;
    if (memOptions.TypeBal='N1A') then ObjBalance.AnDel ;
    if (memOptions.TypeBal='N0A') then ObjBalance.RbDel ;
  end ;

  if V_PGI.IOError=oeOK then
  begin
    for i:=1 to nbLignes do
    begin
      if not IsRowOk(i) then
      begin
        SetRowBad(i) ;
        Continue ;
      end ;
      if SetRowTZF(i) = nil then
        Continue
      else
        nblignes := nblignes;
    end ;

    if (nbLignes>1) then
    begin
      if (memOptions.TypeBal='N1A') then
      begin
         if memOptions.bCalcRes then
           ObjBalance.Write(Plan.OuvrePerte, Plan.OuvreBenef)
         else
           ObjBalance.Write
      end
      else
        ObjBalance.Write ;

     if (memOptions.TypeBal='N1A') then
       ObjBalance.AnWrite(memOptions.bLettrage) ;
     if (memOptions.TypeBal='N0A') then
       ObjBalance.RbWrite ;
    end ; // if (nbLignes>1) then
  end ;

  if V_PGI.IOError=oeOK then
    MarquerPublifi(TRUE) ;
end ;

procedure TFSaisieBalance.PutGrid(Row : LongInt; bColl : Boolean) ;
var EcrRef, Ecr, Source : TZF ; NumCompte, NumAux : string ; i, kFind : Integer ;
    TotalD, TotalC : Double ;
begin
NumCompte:='' ; NumAux:='' ;
NumCompte:=GS.Cells[SB_GEN, Row] ; NumAux:=GS.Cells[SB_AUX, Row] ;
kFind:=BalCourante.FindCompte(NumCompte, Numaux) ;
if (kFind>=0) and ((NumAux='') or ((NumAux<>'') and (not memOptions.WithAux))) then GetRowTZF(Row-1, kFind) ;
if bColl then
  begin
  TotalD:=0 ; TotalC:=0 ;
  EcrRef:=ObjBalance.GetRow(Row-GS.FixedRows) ;
  for i:=0 to PlanTiers.Count-1 do
    begin
    NumAux:=PlanTiers.GetValue('T_AUXILIAIRE', i) ;
    kFind:=BalCourante.FindCompte(NumCompte, NumAux) ;
    if kFind>=0 then
      begin
      Ecr:=TZF.Create('HISTOBAL', EcrRef, -1) ;
      if memOptions.NumPiece>0 then
        begin
        Source:=BalCourante.GetRow(kFind) ;
        Ecr.PutValue('HB_TYPEBAL',       memOptions.TypeBal) ;
        Ecr.PutValue('HB_PLAN',          memOptions.PlanBal) ;
        Ecr.PutValue('HB_EXERCICE',      memOptions.ExoBal) ;
        Ecr.PutValue('HB_DEVISE',        memPivot.Code) ;
        Ecr.PutValue('HB_DATE1',         EncodeDate(memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour)) ;
        Ecr.PutValue('HB_DATE2',         EncodeDate(memOptions.Year,    memOptions.Month,    memOptions.MaxJour)) ;
        Ecr.PutValue('HB_COMPTE1',       NumCompte) ;
        Ecr.PutValue('HB_COMPTE2',       NumAux) ;
        Ecr.PutValue('HB_DEBIT',         Source.GetValue('E_DEBIT')) ;
        Ecr.PutValue('HB_CREDIT',        Source.GetValue('E_CREDIT')) ;
        Ecr.PutValue('HB_DEBITDEV',      Source.GetValue('E_DEBITDEV')) ;
        Ecr.PutValue('HB_CREDITDEV',     Source.GetValue('E_CREDITDEV')) ;
        Ecr.PutValue('HB_ETABLISSEMENT', Source.GetValue('E_ETABLISSEMENT')) ;
        Ecr.PutValue('TYPEGENE',         GS.Cells[SB_TYPE,   Row]) ;
        Ecr.PutValue('LIBELLE',          PlanTiers.GetValue('T_LIBELLE', i)) ;
        Ecr.PutValue('LETTRABLE',        PlanTiers.GetValue('T_LETTRABLE', i)) ;
        Ecr.PutValue('MODEREGL',         PlanTiers.GetValue('T_MODEREGLE', i)) ;
        TotalD:=TotalD+Source.GetValue('E_DEBIT') ;
        TotalC:=TotalC+Source.GetValue('E_CREDIT') ;
        end else
        begin
        Ecr.Assign(BalCourante.GetRow(kFind)) ; ;
        Ecr.PutValue('TYPEGENE',         GS.Cells[SB_TYPE,   Row]) ;
        Ecr.PutValue('LIBELLE',          PlanTiers.GetValue('T_LIBELLE', i)) ;
        Ecr.PutValue('LETTRABLE',        PlanTiers.GetValue('T_LETTRABLE', i)) ;
        Ecr.PutValue('MODEREGL',         PlanTiers.GetValue('T_MODEREGLE', i)) ;
        TotalD:=TotalD+Ecr.GetValue('HB_DEBIT') ;
        TotalC:=TotalC+Ecr.GetValue('HB_CREDIT') ;
        end ;
      end ;
    end ;
(*
    Solde:=Arrondi(TotalD-TotalC, memSaisie.Decimale) ;
    if Solde<>0 then
    begin
    if Solde>0 then GS.Cells[SB_DEBIT,  Row]:=StrFMontant(Solde,15,memSaisie.Decimale,'',TRUE)
               else GS.Cells[SB_CREDIT, Row]:=StrFMontant(Abs(Solde),15,memSaisie.Decimale,'',TRUE) ;
    end ;
*)
  if TotalD<>0 then GS.Cells[SB_DEBIT, Row]:=StrFMontant(TotalD,15,memSaisie.Decimale,'',TRUE) ;
  if TotalC<>0 then GS.Cells[SB_CREDIT, Row]:=StrFMontant(TotalC,15,memSaisie.Decimale,'',TRUE) ;
  end ;
end ;

procedure TFSaisieBalance.ReadBalance ;
var Collectif, CollectifLib, NumAux, p : string ;
    k, kRef, kDecal, kt : Integer ;
begin
bReading:=TRUE ;
Collectif:='' ; CollectifLib:='' ; kDecal:=0 ;
if ObjBalance<>nil then ObjBalance.Free ;
InitBalance ;
ObjBalance:=TZBalance.Create(memSaisie, memOptions) ;
if memOptions.NumPiece>0 then memOptions.bEcr:=TRUE ;
BalCourante:=TZBalance.Create(memSaisie, memOptions) ;
memOptions.bEcr:=FALSE ;
// Lecture de la balance
memOptions.NewObj:=not BalCourante.Read ;
// Chargement du plan comptable général
Plan.Load(FALSE, memOptions.WithRes, FALSE) ;
// Chargement du plan des Tiers
if memOptions.WithAux then PlanTiers.Load(TRUE, TRUE, FALSE) ;
GS.RowCount:=Plan.Count+1 ;
InitMove(Plan.Count+PlanTiers.Count, GetMessageRC(RC_CHARGEPLAN)) ;
for k:=1 to Plan.Count do
  begin
  MoveCur(FALSE) ; kRef:=k-1 ;
  if Plan.GetValue('G_COLLECTIF', kRef)='X' then
    begin  // Tiers d'attente
    CreateRow(k+kDecal) ;
    GS.Cells[SB_COLL, k+kDecal]:='O' ;
    GS.Cells[SB_GEN,  k+kDecal]:=Plan.GetValue('G_GENERAL', kRef) ;
    GS.Cells[SB_LIB,  k+kDecal]:=Plan.GetValue('G_LIBELLE', kRef) ;
    GS.Cells[SB_TYPE, k+kDecal]:=Plan.GetValue('G_NATUREGENE', kRef) ;
    if Plan.GetValue('G_NATUREGENE', kRef)='COC' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefCli else
      if Plan.GetValue('G_NATUREGENE', kRef)='COF' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefFou else
        if Plan.GetValue('G_NATUREGENE', kRef)='COS' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefSal else
          GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefDiv ;
    if not memOptions.WithAux then
      begin
      NumAux:=GS.Cells[SB_AUX, k+kDecal] ;
      kt:=Tiers.GetCompte(NumAux) ; if kt>=0 then p:='T_' else p:='G_' ;
      if (Tiers.GetValue(p+'LETTRABLE', kt)='X')
        then GS.Cells[SB_LETTRE, k+kDecal]:='X'
        else GS.Cells[SB_LETTRE, k+kDecal]:='-' ;
      if (Plan.GetValue('G_POINTABLE', kRef)='X')
        then GS.Cells[SB_LETTRE, k+kDecal]:=GS.Cells[SB_LETTRE, k+kDecal]+'X'
        else GS.Cells[SB_LETTRE, k+kDecal]:=GS.Cells[SB_LETTRE, k+kDecal]+'-' ;
      if VarType(Tiers.GetValue(p+'MODEREGLE', kt))<>VarNull
        then GS.Cells[SB_MODER, k+kDecal]:=Tiers.GetValue(p+'MODEREGLE', kt)
        else GS.Cells[SB_MODER, k+kDecal]:='' ;
      end else
      begin
      if (Plan.GetValue('G_LETTRABLE', kRef)='X')
        then GS.Cells[SB_LETTRE, k+kDecal]:='X'
        else GS.Cells[SB_LETTRE, k+kDecal]:='-' ;
      if (Plan.GetValue('G_POINTABLE', kRef)='X')
        then GS.Cells[SB_LETTRE, k+kDecal]:=GS.Cells[SB_LETTRE, k+kDecal]+'X'
        else GS.Cells[SB_LETTRE, k+kDecal]:=GS.Cells[SB_LETTRE, k+kDecal]+'-' ;
      if VarType(Plan.GetValue('G_MODEREGLE', kRef))<>VarNull
        then GS.Cells[SB_MODER, k+kDecal]:=Plan.GetValue('G_MODEREGLE', kRef)
        else GS.Cells[SB_MODER, k+kDecal]:='' ;
      end ;
    PutGrid(k+kDecal, TRUE) ;
    end else
    begin
    CreateRow(k+kDecal) ;
    GS.Cells[SB_GEN,  k+kDecal]:=Plan.GetValue('G_GENERAL', kRef) ;
    GS.Cells[SB_TYPE, k+kDecal]:=Plan.GetValue('G_NATUREGENE', kRef) ;
    GS.Cells[SB_LIB,  k+kDecal]:=Plan.GetValue('G_LIBELLE', kRef) ;
    if (Plan.GetValue('G_LETTRABLE', kRef)='X')
      then GS.Cells[SB_LETTRE, k+kDecal]:='X'
      else GS.Cells[SB_LETTRE, k+kDecal]:='-' ;
    if (Plan.GetValue('G_POINTABLE', kRef)='X')
      then GS.Cells[SB_LETTRE, k+kDecal]:=GS.Cells[SB_LETTRE, k+kDecal]+'X'
      else GS.Cells[SB_LETTRE, k+kDecal]:=GS.Cells[SB_LETTRE, k+kDecal]+'-' ;
    if VarType(Plan.GetValue('G_MODEREGLE', kRef))<>VarNull
      then GS.Cells[SB_MODER, k+kDecal]:=Plan.GetValue('G_MODEREGLE', kRef)
      else GS.Cells[SB_MODER, k+kDecal]:='' ;
    if memOptions.bCalcRes then
      begin
      if Plan.OuvrePerte=Plan.GetValue('G_GENERAL', kRef) then iRowPerte:=k+kDecal ;
      if Plan.OuvreBenef=Plan.GetValue('G_GENERAL', kRef) then iRowBenef:=k+kDecal ;
      end ;
    PutGrid(k+kDecal, FALSE) ;
    end ;
  end ;
FiniMove ;
BalCourante.Free ;  BalCourante:=nil ;
GS.Invalidate ;
bReading:=FALSE ;
end ;

(* ANCIEN CODE 2
procedure TFSaisieBalance.ReadBalance ;
var Collectif, CollectifLib, NumCompte, NumAux : string ;
    k, kRef, kDecal, kFind, l : integer ; bFind : Boolean ;
begin
bReading:=TRUE ;
Collectif:='' ; CollectifLib:='' ; kDecal:=0 ;
if ObjBalance<>nil then ObjBalance.Free ;
InitBalance ;
ObjBalance:=TZBalance.Create(memSaisie, memOptions) ;
BalCourante:=TZBalance.Create(memSaisie, memOptions) ;
memOptions.NewObj:=not BalCourante.Read ;
// Chargement du plan comptable général
Plan.Load(FALSE, memOptions.WithRes) ;
// Chargement du plan des Tiers
if memOptions.WithAux then PlanTiers.Load(TRUE, TRUE) ;
GS.RowCount:=Plan.Count+PlanTiers.Count+10 ;
InitMove(Plan.Count+PlanTiers.Count, GetMessageRC(RC_CHARGEPLAN)) ;
for k:=1 to Plan.Count do
    begin
    MoveCur(FALSE) ; kRef:=k-1 ;
    if Plan.GetValue('G_COLLECTIF', kRef)='X' then
      begin
      if (memOptions.WithAux) then
        begin
        bFind:=FALSE ;
        for l:=0 to PlanTiers.Count-1 do
          begin
          if PlanTiers.GetValue('T_AUXILIAIRE', l)=VH^.TiersDefCli then Continue ;
          if PlanTiers.GetValue('T_AUXILIAIRE', l)=VH^.TiersDefFou then Continue ;
          if PlanTiers.GetValue('T_AUXILIAIRE', l)=VH^.TiersDefSal then Continue ;
          if PlanTiers.GetValue('T_AUXILIAIRE', l)=VH^.TiersDefDiv then Continue ;
          if Plan.GetValue('G_GENERAL', kRef)<>PlanTiers.GetValue('T_COLLECTIF', l) then
            if bFind then Break else Continue ;
          // Ajout des Tiers
          if bFind then kDecal:=kDecal+1 ;
          CreateRow(k+kDecal) ;
          GS.Cells[SB_COLL, k+kDecal]:='O' ;
          GS.Cells[SB_GEN,  k+kDecal]:=Plan.GetValue('G_GENERAL', kRef) ;
          GS.Cells[SB_LIB,  k+kDecal]:=PlanTiers.GetValue('T_LIBELLE', l) ;
          GS.Cells[SB_AUX,  k+kDecal]:=PlanTiers.GetValue('T_AUXILIAIRE', l) ;
          GS.Cells[SB_TYPE, k+kDecal]:=Plan.GetValue('G_NATUREGENE', kRef) ;
          PutGrid(k+kDecal) ;
          bFind:=TRUE ;
          end ;
          if not bFind then
            begin
            CreateRow(k+kDecal) ;
            GS.Cells[SB_COLL, k+kDecal]:='O' ;
            GS.Cells[SB_GEN,  k+kDecal]:=Plan.GetValue('G_GENERAL', kRef) ;
            GS.Cells[SB_LIB,  k+kDecal]:=Plan.GetValue('G_LIBELLE', kRef) ;
            GS.Cells[SB_TYPE, k+kDecal]:=Plan.GetValue('G_NATUREGENE', kRef) ;
            if Plan.GetValue('G_NATUREGENE', kRef)='COC' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefCli else
              if Plan.GetValue('G_NATUREGENE', kRef)='COF' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefFou else
                if Plan.GetValue('G_NATUREGENE', kRef)='COS' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefSal else
                  GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefDiv ;
            PutGrid(k+kDecal) ;
            end ;
        end else
        begin  // Tiers d'attente
        CreateRow(k+kDecal) ;
        GS.Cells[SB_COLL, k+kDecal]:='O' ;
        GS.Cells[SB_GEN,  k+kDecal]:=Plan.GetValue('G_GENERAL', kRef) ;
        GS.Cells[SB_LIB,  k+kDecal]:=Plan.GetValue('G_LIBELLE', kRef) ;
        GS.Cells[SB_TYPE, k+kDecal]:=Plan.GetValue('G_NATUREGENE', kRef) ;
        if Plan.GetValue('G_NATUREGENE', kRef)='COC' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefCli else
          if Plan.GetValue('G_NATUREGENE', kRef)='COF' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefFou else
            if Plan.GetValue('G_NATUREGENE', kRef)='COS' then GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefSal else
              GS.Cells[SB_AUX, k+kDecal]:=VH^.TiersDefDiv ;
        PutGrid(k+kDecal) ;
        end ;
      end else
      begin
      CreateRow(k+kDecal) ;
      GS.Cells[SB_GEN,  k+kDecal]:=Plan.GetValue('G_GENERAL', kRef) ;
      GS.Cells[SB_TYPE, k+kDecal]:=Plan.GetValue('G_NATUREGENE', kRef) ;
      GS.Cells[SB_LIB,  k+kDecal]:=Plan.GetValue('G_LIBELLE', kRef) ;
      if memOptions.bCalcRes then
        begin
        if Plan.OuvrePerte=Plan.GetValue('G_GENERAL', kRef) then iRowPerte:=k+kDecal ;
        if Plan.OuvreBenef=Plan.GetValue('G_GENERAL', kRef) then iRowBenef:=k+kDecal ;
        end ;
      PutGrid(k+kDecal) ;
      end ;
    end ;
FiniMove ;
BalCourante.Free ;  BalCourante:=nil ;
GS.Invalidate ;
bReading:=FALSE ;
end ;
*)

(* ANCIEN CODE
procedure TFSaisieBalance.ReadBalance ;
var Q : TQuery ; Collectif, CollectifLib, NumCompte, NumAux : string ;
    k, kRef, kDecal, kFind : integer ;
begin
bReading:=TRUE ;
Collectif:='' ; CollectifLib:='' ; kDecal:=0 ;
if ObjBalance<>nil then ObjBalance.Free ;
InitBalance ;
ObjBalance:=TZBalance.Create(memSaisie, memOptions) ;
BalCourante:=TZBalance.Create(memSaisie, memOptions) ;
memOptions.NewObj:=not BalCourante.Read ;
// Chargement du plan comptable général
Plan.Load(memOptions.WithAux, memOptions.WithRes) ;
InitMove(Plan.Count, GetMessageRC(RC_CHARGEPLAN)) ;
for k:=1 to Plan.Count do
    begin
    MoveCur(FALSE) ; kRef:=k-1 ;
    if (Collectif<>'') and (Collectif<>Plan.GetValue('G_GENERAL', kRef)) then
       begin
       CreateRow(k+kDecal) ;
       GS.Cells[SB_GEN,  k+kDecal]:=Collectif ;    Collectif:='' ;
       GS.Cells[SB_LIB,  k+kDecal]:=CollectifLib ; CollectifLib:='' ;
       GS.Cells[SB_COLL, k+kDecal]:='X' ;
       kDecal:=kDecal+1 ;
       end ;
    CreateRow(k+kDecal) ;
    GS.Cells[SB_GEN,  k+kDecal]:=Plan.GetValue('G_GENERAL', kRef) ;
    GS.Cells[SB_TYPE, k+kDecal]:=Plan.GetValue('G_NATUREGENE', kRef) ;
    if memOptions.bCalcRes then
       begin
       if Plan.OuvrePerte=Plan.GetValue('G_GENERAL', kRef) then iRowPerte:=k+kDecal ;
       if Plan.OuvreBenef=Plan.GetValue('G_GENERAL', kRef) then iRowBenef:=k+kDecal ;
       end ;
    if (memOptions.WithAux) and (Plan.GetValue('G_COLLECTIF', kRef)='X') then
       if (VarType(Plan.GetValue('T_AUXILIAIRE', kRef))<>VarNull) then
          begin
          Collectif:=Plan.GetValue('G_GENERAL', kRef) ;
          CollectifLib:=Plan.GetValue('G_LIBELLE', kRef) ;
          end else
          begin
          Collectif:='' ; CollectifLib:='' ;
          GS.Cells[SB_COLL, k+kDecal]:='O' ;
          end ;
    if (memOptions.WithAux) and (VarType(Plan.GetValue('T_AUXILIAIRE', kRef))<>VarNull) then
       begin
       GS.Cells[SB_AUX, k+kDecal]:=Plan.GetValue('T_AUXILIAIRE', kRef) ;
       GS.Cells[SB_LIB, k+kDecal]:=Plan.GetValue('T_LIBELLE', kRef) ;
       end else
       begin
       GS.Cells[SB_LIB, k+kDecal]:=Plan.GetValue('G_LIBELLE', kRef) ;
       end ;
    NumCompte:='' ; NumAux:='' ;
    NumCompte:=Plan.GetValue('G_GENERAL', kRef) ;
    if (memOptions.WithAux) and (VarType(Plan.GetValue('T_AUXILIAIRE', kRef))<>VarNull) then
      NumAux:=Plan.GetValue('T_AUXILIAIRE', kRef) ;
    kFind:=BalCourante.FindCompte(NumCompte, Numaux) ;
    if kFind>=0 then GetRowTZF(k-1+kDecal, kFind) ;
    end ;
FiniMove ;
BalCourante.Free ;  BalCourante:=nil ;
GS.Invalidate ;
bReading:=FALSE ;
end ;
*)

//=======================================================
//================ Saisie complémentaires ===============
//=======================================================

//=======================================================
//================= Fonctions de calcul =================
//=======================================================
function TFSaisieBalance.GetCalcRes : Boolean ;
var i : Integer ;
begin
Result:=FALSE ;
if ObjBalance=nil then Exit ;
for i:=GS.FixedRows to nbLignes do
  begin
  if ((GS.Cells[SB_TYPE, i]='CHA') or (GS.Cells[SB_TYPE, i]='PRO')) and
     ((GS.Cells[SB_DEBIT, i]<>'') or (GS.Cells[SB_CREDIT, i]<>''))then
    begin Result:=TRUE ; Break ;end ;
  end ;
end ;

procedure TFSaisieBalance.CalculSoldeBalance ;
var CumulD, CumulC, TotalD, TotalC, Resultat : Double ;
    CollD, CollC : Double ; i : integer ; Classe : string ;
begin
if GS.Row>nbLignes then Exit ;
CumulD:=0 ;    CumulC:=0 ;    TotalD:=0;    TotalC:=0 ;
SoldeBalD:=0 ; SoldeBalC:=0 ; Resultat:=0 ; CollD:=0 ;  CollC:=0 ;
Classe:=Copy(GS.Cells[SB_GEN, GS.Row], 1, 1);
for i:=GS.FixedRows to nbLignes do
    begin
    if (memOptions.bCalcRes) and (Plan<>nil) and
       ((GS.Cells[SB_GEN,i]=Plan.OuvrePerte) or (GS.Cells[SB_GEN,i]=Plan.OuvreBenef)) then Continue ;
    if GS.Cells[SB_COLL, i]='X' then Continue ;
    // Solde progressif de la classe courante
    if (i<=GS.Row) and (Copy(GS.Cells[SB_GEN, i], 1, 1)=Classe) then
       begin
       CumulD:=CumulD+Valeur(GS.Cells[SB_DEBIT,  i]) ;
       CumulC:=CumulC+Valeur(GS.Cells[SB_CREDIT, i]) ;
       end ;
    // Totaux des comptes progressif
    if i<=GS.Row then
       begin
       TotalD:=TotalD+Valeur(GS.Cells[SB_DEBIT,  i]) ;
       TotalC:=TotalC+Valeur(GS.Cells[SB_CREDIT, i]) ;
       end ;
    // Calcul du résultat
    if (memOptions.TypeBal='BDS') and (Plan<>nil) and
       ((GS.Cells[SB_GEN,i]=Plan.OuvrePerte) or
        (GS.Cells[SB_GEN,i]=Plan.OuvreBenef)) then
       Resultat:=Resultat-Valeur(GS.Cells[SB_DEBIT, i])
                         +Valeur(GS.Cells[SB_CREDIT, i]) ;
    if GS.Cells[SB_TYPE, i]='CHA' then
       Resultat:=Resultat-Valeur(GS.Cells[SB_DEBIT, i])
                         +Valeur(GS.Cells[SB_CREDIT, i]) ;
    if GS.Cells[SB_TYPE, i]='PRO' then
       Resultat:=Resultat-Valeur(GS.Cells[SB_DEBIT, i])
                         +Valeur(GS.Cells[SB_CREDIT, i]) ;
    // Calcul des collectifs
    if (memOptions.WithAux) and (GS.Cells[SB_COLL, i]='X') then
       begin
       GS.Cells[SB_DEBIT,  i]:=StrFMontant(CollD, 15, memSaisie.Decimale, '', TRUE) ;
       GS.Cells[SB_CREDIT, i]:=StrFMontant(CollC, 15, memSaisie.Decimale, '', TRUE) ;
       CollD:=0 ; CollC:=0 ;
       end ;
    if (memOptions.WithAux) and (GS.Cells[SB_AUX, i]<>'') and (GS.Cells[SB_COLL, i]<>'X') then
       begin
       CollD:=CollD+Valeur(GS.Cells[SB_DEBIT,  i]) ;
       CollC:=CollC+Valeur(GS.Cells[SB_CREDIT, i]) ;
       end ;
    // Calcul du solde de la balance
    SoldeBalD:=SoldeBalD+Valeur(GS.Cells[SB_DEBIT,   i]) ;
    SoldeBalC:=SoldeBalC+Valeur(GS.Cells[SB_CREDIT,  i]) ;
    end ;
if Resultat>=0 then
   begin
   HLRESULTAT.Caption:=GetMessageRC(RC_BENEFICE) ; SA_RESULTAT.Tag:=1 ;
   if (memOptions.bCalcRes) and (iRowBenef<>-1) then
      begin
      Resultat:=Arrondi(Resultat, memSaisie.Decimale) ;
      if Resultat<>0 then GS.Cells[SB_CREDIT, iRowBenef]:=StrFMontant(Abs(Resultat),15,memSaisie.Decimale,'',TRUE)
                     else GS.Cells[SB_CREDIT, iRowBenef]:='' ;
      if (iRowPerte<>-1) then GS.Cells[SB_DEBIT, iRowPerte]:='' ;
      end ;
   end else
   begin
   HLRESULTAT.Caption:=GetMessageRC(RC_PERTE) ; SA_RESULTAT.Tag:=-1;
   if (memOptions.bCalcRes) and (iRowPerte<>-1) then
      begin
      Resultat:=Arrondi(Resultat, memSaisie.Decimale) ;
      if Resultat<>0 then GS.Cells[SB_DEBIT, iRowPerte]:=StrFMontant(Abs(Resultat),15,memSaisie.Decimale,'',TRUE)
                     else GS.Cells[SB_DEBIT, iRowPerte]:='' ;
      if (iRowBenef<>-1) then GS.Cells[SB_CREDIT, iRowBenef]:='' ;
      end ;
   end ;
SA_RESULTAT.Value:=Abs(Resultat) ;
HLCUMUL.Caption:=GetMessageRC(RC_CUMULCLASSE)+' '+Classe ;
SA_CUMULDEBIT.Value:=CumulD ;
SA_CUMULCREDIT.Value:=CumulC ;
SA_TOTALDEBIT.Value:=TotalD ;
SA_TOTALCREDIT.Value:=TotalC ;
AfficheLeSolde(SA_SOLDE, TotalD, TotalC) ;
AfficheLeSolde(SA_BALANCE, SoldeBalD, SoldeBalC) ;
end ;

function TFSaisieBalance.IsSoldeBalance : boolean ;
begin
Result:=Arrondi(SA_BALANCE.Value, memSaisie.Decimale)=0 ;
end ;

//=======================================================
//=================== Gestion du Grid ===================
//=======================================================
function TFSaisieBalance.GetGridSens(ACol, ARow : integer) : integer;
begin
// Sens de déplacement dans le Grid
if (GS.Row=ARow) then
   begin
   if (GS.Col>ACol) then Result:=1 else Result:=-1;
   end else if (GS.Row>ARow) then Result:=1 else Result:=-1;
end;

(*
procedure TFSaisieBalance.SetGridSep(ACol, ARow : integer ; Canvas : TCanvas; bHaut : boolean) ;
var R : TRect ;
begin
Canvas.Brush.Color := clRed ;
Canvas.Brush.Style := bsSolid ;
Canvas.Pen.Color   := clRed ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psSolid ;
Canvas.Pen.Width   := 1 ;
R:=GS.CellRect(ACol, ARow) ;
if bHaut then begin Canvas.MoveTo(R.Left, R.Top) ; Canvas.LineTo(R.Right+1, R.Top) end 
         else begin Canvas.MoveTo(R.Left, R.Bottom-1) ; Canvas.LineTo(R.Right+1, R.Bottom-1) end ;
end ;
*)

procedure TFSaisieBalance.SetGridGrise(ACol, ARow : integer ; Canvas : TCanvas) ;
var R : TRect ;
begin
Canvas.Brush.Color := GS.FixedColor ;
Canvas.Brush.Style := bsBDiagonal ;
Canvas.Pen.Color   := GS.FixedColor ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psClear ;
Canvas.Pen.Width   := 1 ;
R:=GS.CellRect(ACol, ARow) ;
Canvas.Rectangle(R.Left, R.Top, R.Right+1, R.Bottom+1) ;
end ;

(*
procedure TFSaisieBalance.SetGridBold(ACol, ARow : integer ; Canvas : TCanvas) ;
var R : TRect ; Text : array[0..255] of Char ;
begin
Canvas.Font.Style:=[fsBold]+[fsItalic] ;
Canvas.Font.Size:=15 ;
R:=GS.CellRect(ACol, ARow) ;
//Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom) ;
StrPCopy(Text, GS.Cells[ACol, ARow]);
ExtTextOut(Canvas.Handle, R.Left+2, R.Top+2,
           ETO_OPAQUE or ETO_CLIPPED, @Rect, Text, StrLen(Text), nil) ;
end ;
*)

procedure TFSaisieBalance.GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin
if (bReading) and (ARow>(GS.Row+GS.VisibleRowCount)) then Exit ;
if GS.Cells[SB_COLL, ARow]='X' then Canvas.Font.Style:=[fsBold] ;
if GS.Cells[SB_COLL, ARow]='O' then Canvas.Font.Style:=[fsBold] ;
end ;

procedure TFSaisieBalance.PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var bGrise : Boolean ;
begin
if (ARow<GS.FixedRows) or (ARow>nbLignes+1) then Exit ;
if (bReading) and (ARow>(GS.Row+GS.VisibleRowCount)) then Exit ;
bGrise:=FALSE ;
if GS.Cells[SB_COLL, ARow]='X' then bGrise:=TRUE ;
if bGrise then SetGridGrise(ACol, ARow, Canvas);
end ;

procedure TFSaisieBalance.UpdateScrollBars ;
begin
//if nbLignes>GS.VisibleRowCount-1 then GS.ScrollBars:=ssBoth else GS.ScrollBars:=ssHorizontal ;
end ;

procedure TFSaisieBalance.CreateRow(Row : longint) ;
//var bInsert : boolean ; RowRef : longint ;
begin
//RowRef:=Row-1 ; bInsert:=FALSE ;
  if Row <= nbLignes then
  begin
    //bInsert:=TRUE ;
    GS.InsertRow(Row) ;
    //if Row=1 then RowRef:=Row+1 ;
  end ;
  Inc(nbLignes) ;
  UpdateScrollBars ;
  if GS.RowCount-2 < nbLignes then
    GS.RowCount:=GS.RowCount+1 ;

  // Etablissement par défaut
  GS.Cells[SB_ETABL,Row]:=VH^.EtablisDefaut ;
  // Ligne créée en saisie
  GS.Cells[SB_NEW,Row]:='O' ;
  // Création de l'objet correspondant
  ObjBalance.CreateRow(nil, Row-GS.FixedRows) ;
end ;

procedure TFSaisieBalance.DeleteRow(Row : longint) ;
begin
ObjBalance.DeleteRow(Row-GS.FixedRows) ;
GS.DeleteRow(Row) ; Dec(nbLignes) ; UpdateScrollBars ;
if GS.RowCount>MAX_ROW then GS.RowCount:=GS.RowCount-1 ;
CalculSoldeBalance ;
if (Row=GS.FixedRows) and (nbLignes=0) then NextRow(Row) ;
end ;

function TFSaisieBalance.NextRow(Row : longint) : boolean ;
begin
if Row>nbLignes(*+1*) then begin Result:=FALSE ; Exit ; end ;
//if Row>nbLignes then CreateRow(Row) ;
Result:=TRUE ;
end ;

// En saisie, on considère que la ligne est valide si le compte général est renseigné
function TFSaisieBalance.IsRowValid(Row : integer; var ACol : integer) : boolean ;
begin
Result:=TRUE ;
if GS.Cells[SB_GEN, Row]='' then begin Result:=FALSE ; ACol:=SB_GEN ; end ;
end ;

function TFSaisieBalance.IsRowOk(Row : longint) : boolean ;
begin
Result:=TRUE ;
if GS.Cells[SB_GEN, Row]='' then Result:=FALSE ;
if (GS.Cells[SB_DEBIT, Row]='') and (GS.Cells[SB_CREDIT, Row]='') then Result:=FALSE ;
end ;

//=======================================================
//================= Evénements du Grid ==================
//=======================================================
procedure TFSaisieBalance.GSEnter(Sender: TObject);
begin
if bFirst then
   begin
   NextRow(GS.Row) ; CalculSoldeBalance ; SetGridOptions(GS.Row) ;
   GS.Col:=SB_DEBIT ;
   HB_DEVISE.Enabled:=FALSE ; EnableButtons ;
   PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
   PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
   end ;
bFirst:=FALSE ;
end;

procedure TFSaisieBalance.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var lc, lr, sens : integer ;
begin
// Bouton Elipsis
GS.ElipsisButton:=FALSE ;
if goRowSelect in GS.Options then Exit ;
lc:=GS.Col ; lr:=GS.Row ; sens:=GetGridSens(ACol, ARow) ;
// Vérifier que l'on est pas en dehors des zones de saisies autorisées
if lr>nbLignes then begin Cancel:=TRUE ; Exit ; end ;
// Attention l'ordre des tests est important
(*
if (ACol=SB_DEBIT) and (sens=-1) and (lr-1>=GS.FixedRows) then
   begin
   ACol:=SB_CREDIT ; ARow:=lr-1 ;
   Cancel:=TRUE ; Exit ;
   end ;
if (ACol=SB_CREDIT) and (sens=1) and (lr<=nbLignes) then
   begin
   ACol:=SB_DEBIT ; ARow:=lr ;
   Cancel:=TRUE ; Exit ;
   end ;
*)
if (ACol=SB_DEBIT) and (sens=-1) and (lr<GS.FixedRows) then
   begin Cancel:=TRUE ; Exit ; end ;
if (ACol=SB_CREDIT) and (sens=1) and (lr>nbLignes) then
   begin Cancel:=TRUE ; Exit ; end ;
if ((lc<>SB_DEBIT) and (lc<>SB_CREDIT)) then
   begin
   if sens=1 then begin if lc>SB_DEBIT  then ACol:=SB_CREDIT else ACol:=SB_DEBIT ;  end
             else begin if lc>SB_CREDIT then ACol:=SB_DEBIT  else ACol:=SB_CREDIT ; end ;
   ARow:=lr ;
   Cancel:=TRUE ;
   Exit ;
   (*
   if (ACol=SB_CREDIT) and (ARow>=nbLignes) then begin Cancel:=TRUE ; Exit ; end ;
   ACol:=lc+sens ; ARow:=lr ;
   if ACol<SB_GEN then
      begin
      ACol:=SB_CREDIT ; ARow:=lr-1 ;
      if ARow<GS.FixedRows then begin ACol:=SB_DEBIT ; ARow:=GS.FixedRows ; end ;
      end ;
   Cancel:=TRUE ; Exit ;
   *)
   end ;
if (GS.Cells[SB_COLL, lr]='X') then
   begin
   ARow:=lr+sens ;
   Cancel:=TRUE ; Exit ;
   end ;
end;

procedure TFSaisieBalance.GSCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
// Formatage du montant
if (ACol=SB_DEBIT) or (ACol=SB_CREDIT) then
   begin
   FormatMontant(GS, ACol, ARow, memSaisie.Decimale) ;
   if (not VH^.MontantNegatif) and (Valeur(GS.Cells[ACol, ARow])<0) then
     begin
     PrintMessageRC(RC_NONEGATIF) ; GS.Cells[ACol, ARow]:='' ;
     CalculSoldeBalance ;
     Cancel:=TRUE ; Exit ;
     end ;
   if (Valeur(GS.Cells[ACol, ARow])<>0) and
      (Valeur(GS.Cells[ACol, ARow])<VH^.GrpMontantMin) or
      (Valeur(GS.Cells[ACol, ARow])>VH^.GrpMontantMax) then
     begin
     PrintMessageRC(RC_NOGRPMONTANT) ; GS.Cells[ACol, ARow]:='' ;
     CalculSoldeBalance ;
     Cancel:=TRUE ; Exit ;
     end ;
   CalculSoldeBalance ;
   end ;
end;

procedure TFSaisieBalance.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var ACol : integer ;
begin
if not IsRowValid(Ou, ACol) then
   begin
   if (Ou=nbLignes) and (GetGridSens(GS.Col, Ou)=-1) then DeleteRow(Ou)
                                                     else Cancel:=TRUE ;
   end else
   begin
   SetRowTZF(Ou) ;
   if GS.Row>Ou then Cancel:=not NextRow(GS.Row) ;
   end ;
end;

procedure TFSaisieBalance.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
SetGridOptions(Ou) ;
EnableButtons ;
InfosPied ;
GS.Invalidate ;
end;

procedure TFSaisieBalance.GSExit(Sender: TObject) ;
var Cancel, Chg : Boolean ; ACol, ARow : LongInt ;
begin
ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ; Chg:=FALSE ;
GSCellExit(GS, ACol, ARow, Cancel) ; GSRowExit(GS, ARow, Cancel, Chg) ;
GS.Col:=ACol ; GS.Row:=ARow ;
end;

procedure TFSaisieBalance.GSKeyPress(Sender: TObject; var Key: Char);
begin
if not GS.SynEnabled then Key:=#0 else
   begin
//   if (Key='+') and ((GS.Col=SF_NAT) or (GS.Col=SF_GEN) or (GS.Col=SF_AUX)) then Key:=#0 ;
//   if (Key=' ') and (GS.Col=SF_NAT) then Key:=#0 ;
//   if (Key<>' ') and (GS.Col=SF_NAT) then Key:=#0 ;
   end ;
end;

procedure TFSaisieBalance.GSDblClick(Sender: TObject);
var EcrRef : TZF ; i : Integer ; TotalD, TotalC, DAvant, CAvant : Double ;
begin
if GS.Cells[SB_COLL, GS.Row]<>'O' then Exit ;
if not memOptions.WithAux then Exit ;
//Tmp:=TZF.Create('HISTOBAL', nil, -1) ;
EcrRef:=SetRowTZF(GS.Row) ; if EcrRef=nil then Exit ;
(*
for i:=0 to Ecr.Detail.Count-1 do
  begin
  EcrAux:=TZF(Ecr.Detail[i]) ; if EcrAux=nil then Break ;
  TmpD:=TZF.Create('HISTOBAL', Tmp, -1) ; TmpD.Assign(EcrAux) ;
  end ;
*)
DAvant:=SoldeBalD-Valeur(GS.Cells[SB_DEBIT, GS.Row]) ;
CAvant:=SoldeBalC-Valeur(GS.Cells[SB_CREDIT, GS.Row]) ;
SaisieBalAux(EcrRef, PlanTiers, GS.Cells[SB_TYPE, GS.Row], memSaisie, DAvant, CAvant) ;
TotalD:=0 ; TotalC:=0 ;
for i:=0 to EcrRef.Detail.Count-1 do
  begin
  TotalD:=TotalD+EcrRef.Detail[i].GetValue('HB_DEBIT') ;
  TotalC:=TotalC+EcrRef.Detail[i].GetValue('HB_CREDIT') ;
  EcrRef.Detail[i].PutValue('HB_TYPEBAL',       EcrRef.GetValue('HB_TYPEBAL')) ;
  EcrRef.Detail[i].PutValue('HB_PLAN',          EcrRef.GetValue('HB_PLAN')) ;
  EcrRef.Detail[i].PutValue('HB_EXERCICE',      EcrRef.GetValue('HB_EXERCICE')) ;
  EcrRef.Detail[i].PutValue('HB_DATE1',         EcrRef.GetValue('HB_DATE1')) ;
  EcrRef.Detail[i].PutValue('HB_DATE2',         EcrRef.GetValue('HB_DATE2')) ;
  EcrRef.Detail[i].PutValue('HB_COMPTE1',       EcrRef.GetValue('HB_COMPTE1')) ;
  EcrRef.Detail[i].PutValue('HB_ETABLISSEMENT', EcrRef.GetValue('HB_ETABLISSEMENT')) ;
  EcrRef.Detail[i].PutValue('TYPEGENE',         EcrRef.GetValue('TYPEGENE')) ;
  end ;
EcrRef.PutValue('HB_DEBIT',  TotalD) ;
EcrRef.PutValue('HB_CREDIT', TotalC) ;
(*
Solde:=Arrondi(TotalD-TotalC, memSaisie.Decimale) ;
if Solde>=0 then
  begin
  EcrRef.PutValue('HB_DEBIT',  Solde) ;
  EcrRef.PutValue('HB_CREDIT', 0) ;
  end else
  begin
  EcrRef.PutValue('HB_DEBIT',  0) ;
  EcrRef.PutValue('HB_CREDIT', Abs(Solde)) ;
  end ;
*)
GetRowTZF(GS.Row-1, -1);
CalculSoldeBalance ;
end;

procedure TFSaisieBalance.GSMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin GridX:=X ; GridY:=Y ; end ;

procedure TFSaisieBalance.PreparePrintGrid(bPrint : Boolean) ;
var c, r, z : Integer ; PrintCaption, Classe : String ; TDClasse, TCClasse, TD, TC : Double ;
begin
z:=1 ; GSPrint.RowCount:=2 ; TDClasse:=0 ; TCClasse:=0 ; TD:=0 ; TC:=0 ; Classe:='' ;
// Corps de la balance : Comptes mouvementés
for r:=GS.FixedRows to nbLignes do
  begin
  if (bPrint) and (Classe<>'') and (Copy(GS.Cells[SB_GEN, r], 1, 1)<>Classe) then
    begin
    GSPrint.Cells[SB_GEN-SB_FIRST, z]:=TraduireMemoire('*** CUMUL CLASSE '+Classe) ;
    if Arrondi(TDClasse, memSaisie.Decimale)<>0 then GSPrint.Cells[SB_DEBIT-SB_FIRST,  z]:=StrFMontant(TDClasse, 15, memSaisie.Decimale, '', TRUE) ;
    if Arrondi(TCClasse, memSaisie.Decimale)<>0 then GSPrint.Cells[SB_CREDIT-SB_FIRST, z]:=StrFMontant(TCClasse, 15, memSaisie.Decimale, '', TRUE) ;
    TDClasse:=0 ; TCClasse:=0 ; Classe:='' ;
    Inc(z) ; GSPrint.RowCount:=GSPrint.RowCount+1 ;
    end ;
  if (Valeur(GS.Cells[SB_DEBIT, r])=0) and (Valeur(GS.Cells[SB_CREDIT, r])=0) then Continue ;
  for c:=SB_GEN to SB_LAST do GSPrint.Cells[c-SB_FIRST, z]:=GS.Cells[c, r] ;
  Inc(z) ; GSPrint.RowCount:=GSPrint.RowCount+1 ;
  if bPrint then
    begin
    Classe:=Copy(GS.Cells[SB_GEN, r], 1, 1) ;
    TDClasse:=TDClasse+Valeur(GS.Cells[SB_DEBIT,  r]) ;
    TCClasse:=TCClasse+Valeur(GS.Cells[SB_CREDIT, r]) ;
    end ;
  TD:=TD+Valeur(GS.Cells[SB_DEBIT,  r]) ; TC:=TC+Valeur(GS.Cells[SB_CREDIT, r]) ;
  end ;
// Totaux
//Inc(z) ; GSPrint.RowCount:=GSPrint.RowCount+1 ;
if bPrint then
  begin
  GSPrint.Cells[SB_GEN-SB_FIRST, z]:=TraduireMemoire('*** TOTAUX') ;
  if Arrondi(TD, memSaisie.Decimale)<>0 then GSPrint.Cells[SB_DEBIT-SB_FIRST,  z]:=StrFMontant(TD, 15, memSaisie.Decimale, '', TRUE) ;
  if Arrondi(TC, memSaisie.Decimale)<>0 then GSPrint.Cells[SB_CREDIT-SB_FIRST, z]:=StrFMontant(TC, 15, memSaisie.Decimale, '', TRUE) ;
  // Inc(z) ;
  GSPrint.RowCount:=GSPrint.RowCount+1 ;
  end ;
// Le bon titre
if memOptions.TypeBal='N1A' then PrintCaption:=GetMessageRC(RC_BALN1A) else
  if memOptions.TypeBal='N1C' then PrintCaption:=GetMessageRC(RC_BALN1C) else
    if memOptions.TypeBal='N0A' then PrintCaption:=GetMessageRC(RC_BALN0A) else
      PrintCaption:=GetMessageRC(RC_BALBDS) ;
if (memOptions.TypeBal<>'BDS')
  then
    PrintCaption:=PrintCaption+Format(' au %.2d/%.2d/%.2d',
                                      [memOptions.DebJour, memOptions.DebMonth, memOptions.DebYear])
  else
    PrintCaption:=PrintCaption+Format(' du %.2d/%.2d/%.2d au %.2d/%.2d/%.2d',
                                      [memOptions.DebJour, memOptions.DebMonth, memOptions.DebYear,
                                       memOptions.MaxJour, memOptions.Month,    memOptions.Year]) ;
if bPrint then
  {$IFDEF EAGLCLIENT}
    PrintDBGrid( PrintCaption, GS.ListeParam, '', '');
  {$ELSE}
    PrintDBGrid(GSPrint, nil, PrintCaption, '') ;
  {$ENDIF}
end ;

//=======================================================
//=============== Evénements des boutons ================
//=======================================================
procedure TFSaisieBalance.BChercherClick(Sender: TObject);
begin if bReading then Exit ; SearchClick ; end ;

procedure TFSaisieBalance.BValideClick(Sender: TObject);
begin
 if bReading then Exit ;
 ValClick ;
end ;

procedure TFSaisieBalance.BNewClick(Sender: TObject);
begin if bReading then Exit ; NewClick ; end ;

procedure TFSaisieBalance.BAbandonClick(Sender: TObject);
begin ByeClick ; end ;

procedure TFSaisieBalance.BSoldeClick(Sender: TObject);
begin if bReading then Exit ; SoldeClick(-1) ; end ;

procedure TFSaisieBalance.SearchClick ;
begin
if bReading then Exit ;
FindFirst:=TRUE ; FindSais.Execute ;
end ;

procedure TFSaisieBalance.FindSaisFind(Sender: TObject);
var bCancel : Boolean ; ARow, ACol : LongInt ;
begin
ARow:=0 ; ACol:=GS.Col ;
if (GS.Row>GS.FixedRows) and (not (frDown in FindSais.Options)) then
  begin GS.Row:=GS.Row-1 ; ARow:=GS.Row ; end ;
Rechercher(GS, FindSais, FindFirst) ;
if ARow=GS.Row then GS.Row:=GS.Row-1 ;
//GS.Col:=SB_DEBIT ;
GSRowEnter(nil, GS.Row, bCancel, FALSE) ;
GS.Col:=ACol ; //SB_DEBIT ;
end ;

procedure TFSaisieBalance.ValClick ;
begin
CalculSoldeBalance ;
EnableButtons ;
if not BValide.Enabled then Exit ;
if (memOptions.TypeBal='N1A') and (PrintMessageRC(RC_EXOV8)=mrYes) then memOptions.bLettrage:=TRUE ;
if ValideBalance then
begin
  memOptions.bVisu := True;
  SetGridOptions( GS.Row );
  bQuestion:=FALSE ;
  BValide.Enabled := False;
  ByeClick;
end;
end;

function TFSaisieBalance.ValideBalance : Boolean ;
//var Start : LongInt ;
begin
if nbLignes<=0 then begin Result:=TRUE ; Exit ; end ;
//start := GetTickCount;
Result:=(Transactions(WriteBalance, 5)=oeOk) ;
if not Result then begin FiniMove ; PrintMessageRC(RC_BADWRITE) ; end ;
//InitBalance ;
//ShowMessage(Format('%f', [(GetTickCount-Start)/1000])) ;
end ;

procedure TFSaisieBalance.DelClick ;
var lr, ltop : longint ;
begin
lr:=GS.Row ; ltop:=GS.TopRow ; DeleteRow(lr) ;
if lr>nbLignes then GS.Row:=nbLignes else GS.Row:=lr ;
GS.TopRow:=ltop ; GS.ElipsisButton:=FALSE ;
end ;

procedure TFSaisieBalance.NewClick ;
var NumCompte : string ; i: Integer ;
begin
if bClosing then Exit ;
if not ExJaiLeDroitConcept(TConcept(ccGenCreat),FALSE) then
   begin PrintMessageRC(RC_GENINTERDIT) ; Exit ; end ;
NumCompte:=GS.Cells[SB_GEN, GS.Row] ;
FicheGene(nil, '', '', taCreatOne, 0) ;
ValideBalance ;
ReadBalance ;
for i:=GS.FixedRows to nbLignes do
  if GS.Cells[SB_GEN, i]=NumCompte then begin GS.Row:=i ; Break ; end ;
GS.Enabled:=TRUE ; GS.SetFocus ; GSEnter(nil) ;
end ;

procedure TFSaisieBalance.ByeClick ;
begin
  if not IsInside(Self) then
    Close ;
end;

procedure TFSaisieBalance.SoldeClick(Row : longint) ;
var lr : longint ; Solde, Debit, Credit : double ; bSoldeDebit : Boolean ;
begin
if BSolde.Enabled=FALSE then Exit ;
if Row<0 then lr:=GS.Row else lr:=Row ;
Solde:=SA_BALANCE.Value ; bSoldeDebit:=SA_BALANCE.Debit ;
Debit:=Valeur(GS.Cells[SB_DEBIT,lr]) ; Credit:=Valeur(GS.Cells[SB_CREDIT,lr]) ;
// Partie débit
if (Debit<>0) and (bSoldeDebit) then
   if (Debit-Solde)<0 then
      begin
      GS.Cells[SB_CREDIT,lr]:=StrFMontant(Abs(Debit-Solde),15,memSaisie.Decimale,'',TRUE) ;
      GS.Cells[SB_DEBIT,lr]:='' ;
      GS.Col:=SB_CREDIT ;
      end else GS.Cells[SB_DEBIT,lr]:=StrFMontant(Debit-Solde,15,memSaisie.Decimale,'',TRUE) ;
if (Debit<>0) and (not bSoldeDebit) then
   GS.Cells[SB_DEBIT,lr]:=StrFMontant(Debit+Solde,15,memSaisie.Decimale,'',TRUE) ;
// Partie crédit
if (Credit<>0) and (bSoldeDebit) then
   GS.Cells[SB_CREDIT,lr]:=StrFMontant(Credit+Solde,15,memSaisie.Decimale,'',TRUE) ;
if (Credit<>0) and (not bSoldeDebit) then
   if (Credit-Solde)<0 then
      begin
      GS.Cells[SB_DEBIT,lr]:=StrFMontant(Abs(Credit-Solde),15,memSaisie.Decimale,'',TRUE) ;
      GS.Cells[SB_CREDIT,lr]:='' ;
      GS.Col:=SB_DEBIT ;
      end else GS.Cells[SB_CREDIT,lr]:=StrFMontant(Credit-Solde,15,memSaisie.Decimale,'',TRUE) ;
if (Debit=0) and (Credit=0) then
   begin
   if bSoldeDebit then GS.Cells[SB_CREDIT,lr]:=StrFMontant(Solde,15,memSaisie.Decimale,'',TRUE)
                  else GS.Cells[SB_DEBIT,lr]:=StrFMontant(Solde,15,memSaisie.Decimale,'',TRUE) ;
   end ;
CalculSoldeBalance ;
EnableButtons ;
end ;

//=======================================================
//=========== Evénements des boutons d'entête ===========
//=======================================================
procedure TFSaisieBalance.BValEnteteClick(Sender: TObject);
var DateBal : TDateTime ;
begin
if (memOptions.TypeBal='N0A') then
  begin
  DateBal:=StrToDate(DATE1.Text) ;
  if (DateBal<VH^.Encours.Deb) or (DateBal>VH^.Encours.Fin) then
     begin PrintMessageRC(RC_DATEINCORRECTE) ; DATE1.SetFocus ; Exit ; end ;
  DATE1Click ;
  end ;
BHide.Enabled:=FALSE ;
BHideFocus.Enabled:=FALSE ;
BValEntete.Enabled:=FALSE ;
bUsing:=TRUE ;
EnableButtons ;
end ;

procedure TFSaisieBalance.DATE1Click ;
var DateBal : TDateTime ;
begin
if bClosing then Exit ;
DATE1.Enabled:=FALSE ;
DateBal:=StrToDate(DATE1.Text) ;
DecodeDate(DateBal, memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour) ;
DecodeDate(DateBal, memOptions.Year,    memOptions.Month,    memOptions.MaxJour) ;
ReadBalance ;
GS.Enabled:=TRUE ; GS.SetFocus ; GSEnter(nil) ;
end;

(*
procedure TFSaisieBalance.HB_DATE2Click ;
var bExist : Boolean ;
begin
if bClosing then Exit ;
// Une balance BDS existe-t-elle ?
if Objbalance<>nil then Objbalance.Free ;
ObjBalance:=TZBalance.Create(memSaisie, memOptions) ;
bExist:=ObjBalance.ExistBds(TRUE) ;
if bExist then
  memOptions.WithAux:=FALSE
else begin
  memOptions.WithAux:=FALSE ;

          //memOptions.bCalcRes:=(PrintMessageRC(RC_AVECRES)=mrYes) ;
          end ;
if memOptions.WithAux then memOptions.PlanBal:='AUX'
                      else memOptions.PlanBal:='BAL' ;
ObjBalance.SetOptions(memOptions) ;
if (bGenAuto) and (not bExist) then
   begin
   if not ObjBalance.GenAuto2 then PrintMessageRC(RC_NOGENBALANCE) ; bGenAuto:=FALSE ;
   end ;
ObjBalance.Free ; ObjBalance:=nil ;
ReadBalance ;
//if bExist then memOptions.bCalcRes:=GetCalcRes ;
GS.Enabled:=TRUE ; GS.SetFocus ; GSEnter(nil) ;
end;*)

procedure TFSaisieBalance.PFENMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var r : TRect ;
begin
if bUsing then Exit ;
bUsing:=TRUE ;
if not BValEntete.Enabled then Exit ;
if (not GS.Enabled) and (not bClosing) then
  begin
  r:=Rect(GS.Left, GS.Top, GS.Left+GS.Width, GS.Top+GS.Height) ;
  if PtInRect(r, Point(X, Y)) then BValEnteteClick(nil) ;
//  if bModeRO then PostMessage(GS.Handle, WM_LBUTTONUP, 0, 0) ;
  end ;
bUsing:=FALSE ;
end;

procedure TFSaisieBalance.BHideFocusEnter(Sender: TObject);
begin
if bUsing then Exit ;
bUsing:=TRUE ;
if BValEntete.Enabled then BValEnteteClick(nil) ;
bUsing:=FALSE ;
end;

procedure TFSaisieBalance.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end ;

procedure TFSaisieBalance.BImprimerClick(Sender: TObject);
begin PreparePrintGrid(TRUE) ; end ;

procedure TFSaisieBalance.bExportClick(Sender: TObject);
begin
if not ExJaiLeDroitConcept(ccExportListe, TRUE) then Exit ;
PreparePrintGrid(FALSE) ;
if SD.Execute then ExportGrid(GSPrint, nil, SD.FileName, SD.FilterIndex, TRUE) ;
end;

procedure TFSaisieBalance.SetMontants(LeTZ : TZF; XD, XC : double ; Dev : RDEVISE ; ModeOppose, Force : boolean) ;
var SD, SC : double ; P : string ;
begin
  P:=TableToPrefixe(LeTZ.NomTable) ;
  if Dev.Code=V_PGI.DevisePivot then
  begin
    { Saisie en monnaie de tenue }
    SD:=LeTZ.GetValue(P+'_DEBIT') ; SC:=LeTZ.GetValue(P+'_CREDIT') ;
    if ((SD=XD) and (SC=XC) and (not Force)) then Exit ;
    LeTZ.PutValue(P+'_DEBIT',XD)                   ; LeTZ.PutValue(P+'_CREDIT',XC) ;
    LeTZ.PutValue(P+'_DEBITDEV',XD)                ; LeTZ.PutValue(P+'_CREDITDEV',XC) ;
  end else
  begin
    {Saisie en Devise}
    SD:=LeTZ.GetValue(P+'_DEBITDEV') ; SC:=LeTZ.GetValue(P+'_CREDITDEV') ;
    if ((SD=XD) and (SC=XC) and (not Force)) then Exit ;
    LeTZ.PutValue(P+'_DEBIT',DeviseToEuro(XD,DEV.Taux,DEV.Quotite)) ;
    LeTZ.PutValue(P+'_CREDIT',DeviseToEuro(XC,DEV.Taux,DEV.Quotite)) ;
    LeTZ.PutValue(P+'_DEBITDEV',XD) ; LeTZ.PutValue(P+'_CREDITDEV',XC) ;
  end ;
end ;

end.

