unit ImpBds;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, Hctrls, StdCtrls, Buttons, ExtCtrls, DB, DBCtrls, HEnt1,
  Ent1, HStatus, FmtChoix, Filtre, ComCtrls, Menus, HSysMenu,MajTable, HTB97,
  Hqry, HPanel, UIUtil, UTob, Grids, PrintDBG,
  ImpFicU, SaisUtil, Sisco, TZ, ZTypes, ZBalance,udbxDataset ;

procedure ImportBds(Lequel : string) ;

const RC_BADFILENAME      = 0 ;
      RC_BADPATH          = 1 ;
      RC_BADFORMAT        = 2 ;
      RC_ARRET            = 3 ;
      RC_CPTGENERAL       = 4 ;
      RC_CPTAUX           = 5 ;
      RC_CREE             = 6 ;
      RC_TITREIMPORT      = 7 ;
      RC_VERIF            = 8 ;
      RC_IMPORT           = 9 ;
      RC_PARAMS           =10 ;
      RC_FOURCHETTES      =11 ;
      RC_BADCOMPTE        =12 ;
      RC_BADWRITE         =13 ;
      RC_STOP             =14 ;
      RC_BALEXIST         =15 ;
      RC_BALIMPORT        =16 ;
      RC_BADBALWRITE      =17 ;
      RC_BALBADDATE       =18 ;

type
  RBDSENT = record
    Bal     : TDateTime ;
    CodeBal : string ;
    Debut   : TDateTime ;
    Fin     : TDateTime ;
    Duree   : Integer ;
    bAux    : Boolean ;
    bAna    : Boolean ;
    Monnaie : string ;
    Ver     : Integer ;
    Exo     : string ;
  end ;
  PBDSENTETE = ^RBDSENT ;

  RBDSCPT = record
    Compte    : string[10] ;
    Code      : string[1] ;
    Libelle   : string[25] ;
    Typec     : string[1] ;      // nom original : type
  end ;
  PBDSCOMPTE = ^RBDSCPT ;

  RBDSLIG = record
    Compte    : string[10] ;
    DateC     : TDateTime ;
    Stat1     : string[4] ;      // date fin mois
    DateFin   : TDateTime ;      // date fin annee/mois
    Stat2     : string[4] ;
    Compte2   : string[10] ;
    Debit     : double ;
    Credit    : double ;
//    SldDebit  : double ;
//    SldCredit : double ;
    DebEur    : double ;         // euro
    CreEur    : double ;         // euro
    Exo       : string ;
  end ;
  PBDSLIGNE = ^RBDSLIG ;

const
  NB_R_S_DAD_ENT_GEN =  3 ;
  NB_R_S_DAD_ENT     = 16 ;
  NB_R_S_DAD_CPTE    =  7 ;
  NB_R_S_DAD_LIG     = 16 ;
  NB_R_S_DAD_DEBUT   =  2 ;
  NB_R_S_DAD_FIN     =  2 ;

type
  R_S_DAD_ENT_GEN = record
    code    : string[2] ;
    dossier : string[5] ;
    filler  : string[83] ;
  end ;
  PBALDOSSIER = ^R_S_DAD_ENT_GEN ;

  R_S_DAD_ENT = record
    code    : string[2] ;      // ' E' Entete
    data    : string[2] ;      // date fin annee
    datm    : string[2] ;      // date fin mois
    filler1 : string[32] ;
    cpnbc   : string[6] ;      // nbre de comptes
    cpnbg   : string[6] ;
    debDad  : string[6] ;
    finDad  : string[6] ;
    result  : string[15] ;
    duree   : string[2] ;
    analy   : string[1] ;
    aux     : string[1] ;
    version : string[2] ;
    maj     : string[1] ;
    monnaie : string[1] ;
    filler2 : string[5] ;
  end ;
  PBALENTETE = ^R_S_DAD_ENT ;

  R_S_DAD_CPTE = record
    compte    : string[10] ;
    code      : string[1] ;
    filler1   : string[27] ;
    libelle   : string[25] ;
    typec     : string[1] ;      // nom original : type
    precedent : string[10] ;
    filler2   : string[16] ;
  end ;
  PBALCOMPTE = ^R_S_DAD_CPTE ;

  R_S_DAD_LIG = record
    compte1   : string[10] ;
    code      : string[1] ;
    data      : string[2] ;      // date fin annee
    datm      : string[2] ;      // date fin mois
    stat1     : string[4] ;
    filler1   : string[1] ;
    dtfin     : string[4] ;      // date fin annee/mois
    stat2     : string[4] ;
    compte2   : string[10] ;
    debit     : string[15] ;
    credit    : string[15] ;
    slddebit  : string[15] ;
    sldcredit : string[15] ;
    debeur    : string[15] ;     // euro
    creeur    : string[15] ;     // euro
    filler2   : string[20] ;
  end ;
  PBALLIGNE = ^R_S_DAD_LIG ;

  R_S_DAD_DEBUT = record
    deb    : string[11] ;        // '***DEBUT***'
    filler : string[21] ;
  end ;
  PBALDEBUT = ^R_S_DAD_DEBUT ;

  R_S_DAD_FIN = record
    finsav : string[9] ;         // '***FIN***
    filler : string[20] ;
  end ;
  PBALFIN = ^R_S_DAD_FIN ;

type
  TFImpBds = class(TForm)
    HBds: THMsgBox;
    Sauve: TSaveDialog;
    Pages: TPageControl;
    TabImport: TTabSheet;
    Bevel1: TBevel;
    HLabel7: THLabel;
    HLabel2: THLabel;
    FFormat: THValComboBox;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    GModif: TGroupBox;
    Lib: TCheckBox;
    Car: TCheckBox;
    Image1: TImage;
    Image8: TImage;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    PTools: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BFormat: TToolbarButton97;
    BStop: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    GroupBox1: TGroupBox;
    FPositif: TCheckBox;
    FDelFile: TCheckBox;
    FExo: THValComboBox;
    HLabel1: THLabel;
    TabResult: TTabSheet;
    GS: THGrid;
    BImprimer: TToolbarButton97;
    procedure RechFileClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BFormatClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFormatChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
  private
    bArreter   : Boolean ;
    bLectureOk : Boolean ;
    NbErreur   : Integer ;
    NbImport   : Integer ;
    NbLu       : Integer ;
    NbLignes   : Integer ;
    CurNat     : string ;
    FiltreImp  : string ;
    Fichier    : TextFile ;
    Erreur     : TList ;
    Entete     : TList ;
    Row        : LongInt ;
    bPlan      : Boolean ;
    bPrint     : Boolean ;
    Sisco      : TInfoGeneSISCO ;
    Dev        : RDevise ;
    Opt        : ROpt ;
    iEntete    : Integer ;
    // Fonctions ressource
    function  GetMessageRC(MessageID : Integer) : string ;
    function  PrintMessageRC(MessageID : Integer; StAfter : string='') : Integer ;
    // Fonctions utilitaires
    procedure InitGrid ;
    procedure InitImport ;
    function  ImpStrToDate(St : string) : TDateTime;
    function  ImpStrToDouble(St : string) : Double ;
    function  IsOKImport(DebBal, FinBal : TDateTime) : string ;
    // Fonctions Importation
    procedure InitSisco ;
    procedure InitBal(ENT : PBDSENTETE) ;
    function  ExistGen(var NumCompte : string) : Boolean ;
    function  ExistAux(var NumAux : string) : Boolean ;
    procedure VerifCompte(IC : PBDSCOMPTE) ;
    function  FillEntete(St : string; E : PBALENTETE) : Integer ;
    procedure FillCompte(St : string; C : PBALCOMPTE) ;
    procedure FillLigne(St : string; L : PBALLIGNE; Version : Integer) ;
    function  ReadDAD(bVerifCompte : Boolean=TRUE; ENT : PBDSENTETE=nil) : Boolean ;
    procedure WriteBal ;
    procedure AjoutHistoBal(IL : PBDSLIGNE; ENT : PBDSENTETE) ;
    function  TestBreak : Boolean ;
  public
    Lequel : string ;
  end;

implementation

{$R *.DFM}

const ID_ENTETE     = ' E' ;
      ID_DOSSIER    = ' A' ;
      ID_DEBUT      = '***DEBUT***' ;
      ID_FIN        = '***FIN***' ;
      ID_MAJVERSION = 60 ;

const RE_CODE       = 0 ;
      RE_LIBELLE    = 1 ;
      RE_COMMENT    = 2 ;

const MAX_ROWS      = 500 ;      

//=======================================================
//========== Point d'entrée dans l'importation ==========
//=======================================================
procedure ImportBds(Lequel : string) ;
var FImpBds : TFImpBds ; PP : THPanel ;
begin
if Blocage(['nrCloture', 'nrBatch'], FALSE, 'nrBatch') then Exit ;
FImpBds:=TFImpBds.Create(Application) ;
FImpBds.Lequel:=Lequel ;
FImpBds.BFormat.Visible:=FALSE ;
PP:=FindInsidePanel ;
if PP=Nil then
  begin
  try
    FImpBds.ShowModal ;
  finally
    FImpBds.Free ;
    Bloqueur('nrBatch',False) ;
  end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(FImpBds,PP) ;
  FImpBds.Show ;
  end ;
end ;

//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFImpBds.GetMessageRC(MessageID : Integer) : string ;
begin
Result:=HBds.Mess[MessageID] ;
end ;

function TFImpBds.PrintMessageRC(MessageID : Integer; StAfter : string) : Integer ;
begin
Result:=HBds.Execute(MessageID, Caption, StAfter) ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TFImpBds.FormShow(Sender: TObject);
begin
Pages.ActivePage:=TabImport ;
TabResult.Enabled:=FALSE ;
FExo.ItemIndex:=0 ;
Erreur:=nil ;
Entete:=nil ;
if Lequel='BDS' then // Balance de situation
begin
  FiltreImp:='IMPBDS' ;
  CurNat:='BDS' ;
  Caption:=GetMessageRC(RC_TITREIMPORT) ;
  //   HelpContext:=6210000 ;
end ;
Sauve.Title:=Caption ;
ChargeFiltre(FiltreImp, FFiltres, Pages) ;
if (FFormat.ItemIndex=-1) then FFormatChange(nil) ;
UpdateCaption(Self) ;
InitGrid ;
end;

procedure TFImpBds.FormClose(Sender: TObject; var Action: TCloseAction);
begin
InitImport ;
if IsInside(Self) then
  begin
  Bloqueur('nrBatch', FALSE) ;
  Action:=caFree ;
  end ;
end;

//=======================================================
//================ Evénements des contrôles =============
//=======================================================
procedure TFImpBds.BCreerFiltreClick(Sender: TObject);
begin NewFiltre(FiltreImp, FFiltres, Pages) ; end;

procedure TFImpBds.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre(FiltreImp, FFiltres, Pages) ; end;

procedure TFImpBds.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre(FiltreImp, FFiltres) ; end;

procedure TFImpBds.BRenFiltreClick(Sender: TObject);
begin RenameFiltre(FiltreImp, FFiltres) ; end;

procedure TFImpBds.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres, Pages) ; end;

procedure TFImpBds.FFiltresChange(Sender: TObject);
begin LoadFiltre(FiltreImp, FFiltres, Pages) ; end;

procedure TFImpBds.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve, FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFImpBds.BFormatClick(Sender: TObject);
begin
ChoixFormatImpExp('X', 'FBA') ; // Format Balance
if not (ChoixFmt.OkSauve) then Exit ;
with ChoixFmt do
  begin
  if Format<>'' then FFormat.Value:=Format ;
  FFormat.Enabled:=not (FixeFmt) ;
  if Fichier<>'' then FileName.Text:=Fichier ;
  FileName.Enabled:=not (FixeFichier) ;
  RechFile.Enabled:=not (FixeFichier) ;
  end ;
end;

procedure TFImpBds.BStopClick(Sender: TObject);
begin bArreter:=TRUE ; end;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
procedure TFImpBds.InitGrid ;
begin
Row:=1 ;
GS.RowCount:=MAX_ROWS ;
GS.Refresh ;
end ;

procedure TFImpBds.InitImport ;
var Item : PBDSENTETE ;
begin
if Entete<>nil then
  begin
  while Entete.Count>0 do
    begin Item:=Entete.Items[0] ; Dispose(Item) ; Entete.Delete(0) ; end ;
  Entete.Free ; Entete:=nil ;
  end ;
if Erreur<>nil then begin Erreur.Free ; Erreur:=nil ; end ;
end ;

function TFImpBds.TestBreak : Boolean ;
begin
Application.ProcessMessages ;
if bArreter then if PrintMessageRC(RC_STOP, Caption)<>mrYes then bArreter:=FALSE ;
Result:=bArreter ;
end ;

// -------------------------------------------------------
// Format cas général importé DDMMYY
// -------------------------------------------------------
// Si DD='00' alors DD=Fin de Mois
// Si DD='99' alors DD=Fin de Mois et le format est DDYYMM
//--------------------------------------------------------
// L'année est ajustée sur le siècle correspondant :
// YY='00..50' => 2000
// YY='51..99' => 1900
//--------------------------------------------------------
function TFImpBds.ImpStrToDate(St : string) : TDateTime ;
var Y, M, D : Word ; i : Integer;
begin
// Controle de la coherence de la date au niveau des caractères
for i:=1 to Length(st) do
  if not (st[i] in ['0'..'9']) then
  begin
    St := '01011900';
    break;
  end;
if (St='') or (St='        ' ) then St:='01011900' ;
if St='01011900' then begin Result:=iDate1900 ; Exit ; end ;
D:=StrToInt(Copy(St, 1, 2)) ;
if D=99 then Y:=StrToInt(Copy(St, 3, 2)) else Y:=StrToInt(Copy(St, 5, 2)) ;
if Y<=50 then Y:=Y+2000 else Y:=Y+1900 ;
if D=99 then M:=StrToInt(Copy(St, 5, 2)) else M:=StrToInt(Copy(St, 3, 2)) ;
if (D=0) or (D=99) then begin Result:=FinDeMois(EncodeDate(Y, M, 1)) ; Exit ; end ;
Result:=EncodeDate(Y, M, D) ;
end ;

// -------------------------------------------------------
// Format cas général importé NNNNNNNNNNNNNNN
// -------------------------------------------------------
// Pas de point ou de virgule décimale
// -------------------------------------------------------
function TFImpBds.ImpStrToDouble(St : string) : Double ;
var i : Integer ; k64 : Int64 ;
begin
// Controle de la coherence du montant au niveau des caractères
for i:=1 to Length(st) do
  if not (st[i] in ['0'..'9']) then
  begin
    St := '000000000000000';
    break;
  end;
if (St='') or (St='               ' ) then St:='000000000000000' ;
k64:=StrToInt64(St) ;
Result:=k64/100.0 ;
end ;

// -------------------------------------------------------
// Existe-t-il un EXO contenant les dates de la balance ?
// La balance de situation existe déjà ?
// -------------------------------------------------------
function TFImpBds.IsOKImport(DebBal, FinBal : TDateTime) : string ;
var bEncours : Boolean ; ExoDeb, ExoFin : string ;
    ObjBal : TZBalance ; Opt : ROpt ; Dev : RDevise ;
begin
bEncours:=((DebBal>=VH^.EnCours.Deb) and (DebBal<=VH^.EnCours.Fin)) ;
ExoDeb:=QuelExoDt(DebBal) ;
if (not bEncours) and (ExoDeb=VH^.Encours.Code) then ExoDeb:='' ;
bEncours:=((FinBal>=VH^.EnCours.Deb) and (FinBal<=VH^.EnCours.Fin)) ;
ExoFin:=QuelExoDt(FinBal) ;
if (not bEncours) and (ExoFin=VH^.Encours.Code) then ExoFin:='' ;
if (ExoDeb<>ExoFin) then begin Result:='' ; Exit end ;
if (FExo.ItemIndex>0) and (ExoDeb<>FExo.Values[FExo.ItemIndex]) then ExoDeb:='' ;
if (FExo.ItemIndex>0) and (ExoFin<>FExo.Values[FExo.ItemIndex]) then ExoFin:='' ;
if (ExoDeb='') or (ExoFin='') then Result:='' else Result:=ExoDeb ;
if (ExoDeb<>'') and (ExoFin<>'') and (ExoDeb=ExoFin) then
  begin
//  QuelDateExo(DateBal, DateDeb, DateFin) ;
  Opt.TypeBal:='BDS' ; Opt.ExoBal:=ExoDeb ;
  DecodeDate(DebBal, Opt.DebYear, Opt.DebMonth, Opt.DebJour) ; Opt.DebJour:=1 ;
  DecodeDate(FinBal, Opt.Year,    Opt.Month,    Opt.MaxJour) ; Opt.MaxJour:=1 ;
  ObjBal:=TZBalance.Create(Dev, Opt) ;
  if ObjBal.ExistBds(FALSE) then Result:='PFU' ;
  ObjBal.Free ;
  end ;
end ;

procedure TFImpBds.InitSisco ;
begin
ChargeCptSISCO(Sisco.FourchetteSISCO) ;
ChargeFourchetteCompte('SIS', Sisco.FourchetteImport) ;
ChargeCharRemplace(Sisco) ;
end ;

function TFImpBds.ExistGen(var NumCompte : string) : Boolean ;
var Q : TQuery ;
begin
NumCompte:=BourreEtLess(NumCompte, fbGene) ;
NumCompte:=CptRemplace(NumCompte, Sisco.CharRG) ;
Q:=OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+NumCompte+'"', TRUE) ;
Result:=not Q.EOF ;
Ferme(Q) ;
end ;

function TFImpBds.ExistAux(var NumAux : string) : Boolean ;
var Q : TQuery ;
begin
NumAux:=BourreEtLess(NumAux, fbAux) ;
NumAux:=CptRemplace(NumAux, Sisco.CharRT) ;
Q:=OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+NumAux+'"', TRUE) ;
Result:=not Q.EOF ;
Ferme(Q) ;
end ;

procedure TFImpBds.VerifCompte(IC : PBDSCOMPTE) ;
var Q : TQuery ; NumCompte, NumAux, sComment, Nat : string ; bCreate, bAux : Boolean ;
begin
bCreate:=FALSE ; bAux:=FALSE ;
NumCompte:=AnsiUpperCase(Trim(IC^.Compte)) ;
if NumCompte[1]='}' then Exit ;
if not (NumCompte[1] in ['0'..'9', 'C', 'F']) then
  begin
  bPlan:=FALSE ; bPrint:=TRUE ;
  GS.Cells[RE_CODE,    Row]:=IC^.Compte ;
  GS.Cells[RE_LIBELLE, Row]:=IC^.Libelle ;
  GS.Cells[RE_COMMENT, Row]:=GetMessageRC(RC_BADCOMPTE) ;
  if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
  Row:=Row+1 ;
  Exit ;
  end ;
if NumCompte[1] in ['C', '9', 'F', '0'] then bAux:=TRUE ;
QuelCompte(NumCompte, Sisco.FourchetteSISCO, NumCompte, NumAux) ;
if (bAux) and (NumAux='') then
  begin
  bPlan:=FALSE ; bPrint:=TRUE ;
  GS.Cells[RE_CODE,    Row]:=IC^.Compte ;
  GS.Cells[RE_LIBELLE, Row]:=IC^.Libelle ;
  GS.Cells[RE_COMMENT, Row]:=GetMessageRC(RC_PARAMS) ;
  if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
  Row:=Row+1 ;
  Exit ;
  end ;
if (NumAux<>'') and (not ExistAux(NumAux)) then
  begin
  Q:=OpenSQL('SELECT * FROM TIERS WHERE T_AUXILIAIRE="'+W_W+'"', FALSE) ;
  Q.Insert ; InitNew(Q) ;
  Q.FindField('T_AUXILIAIRE').AsString:=NumAux ;
  if NumAux[1] in ['C','9'] then Q.FindField('T_NATUREAUXI').AsString:='CLI'
                            else Q.FindField('T_NATUREAUXI').AsString:='FOU' ;
  Q.FindField('T_LIBELLE').AsString:=Trim(IC^.Libelle) ;
  Q.FindField('T_COLLECTIF').AsString:=NumCompte ;
  Q.FindField('T_ABREGE').AsString:=Copy(Trim(IC^.Libelle), 1, 17) ;
  Q.FindField('T_CREERPAR').AsString:='BDS' ;
  Q.FindField('T_CODEIMPORT').AsString:=Copy(Trim(IC^.Compte), 1, 17) ;
  Q.Post ;
  Ferme(Q) ;
  bCreate:=TRUE ;
  end ;
if (not bAux) and (NumAux='') and (not ExistGen(NumCompte)) then
  begin
  Nat:=NatureCptImport(NumCompte, Sisco.FourchetteImport) ;
  if Nat='---' then Nat:=GetNatureDiv(NumCompte, Sisco.FourchetteSisco) ;
  if Nat='---' then
    begin
    bPlan:=FALSE ; bPrint:=TRUE ;
    GS.Cells[RE_CODE,    Row]:=IC^.Compte ;
    GS.Cells[RE_LIBELLE, Row]:=IC^.Libelle ;
    GS.Cells[RE_COMMENT, Row]:=GetMessageRC(RC_FOURCHETTES) ;
    if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
    Row:=Row+1 ;
    Exit ;
    end ;
  Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+W_W+'"', FALSE) ;
  Q.Insert ; InitNew(Q) ;
  Q.FindField('G_GENERAL').AsString:=NumCompte ;
  Q.FindField('G_NATUREGENE').AsString:=Nat ;
  Q.FindField('G_LIBELLE').AsString:=Trim(IC^.Libelle) ;
  Q.FindField('G_ABREGE').AsString:=Copy(Trim(IC^.Libelle), 1, 17) ;
  Q.FindField('G_CREERPAR').AsString:='BDS' ;
  Q.FindField('G_CODEIMPORT').AsString:=Copy(Trim(IC^.Compte), 1, 17) ;
  Q.Post ;
  Ferme(Q) ;
  bCreate:=TRUE ;
  end ;
if bCreate then
  begin
  bPrint:=TRUE ;
  GS.Cells[RE_CODE,    Row]:=IC^.Compte ;
  GS.Cells[RE_LIBELLE, Row]:=IC^.Libelle ;
  if bAux then sComment:=GetMessageRC(RC_CPTAUX)
          else sComment:=GetMessageRC(RC_CPTGENERAL) ;
  sComment:=sComment+GetMessageRC(RC_CREE) ;
  GS.Cells[RE_COMMENT, Row]:=sComment ;
  if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
  Row:=Row+1 ;
  end ;
end ;

procedure TFImpBds.InitBal(ENT : PBDSENTETE) ;
begin
// Options
Opt.TypeBal:='BDS' ;
DecodeDate(ENT^.Debut, Opt.DebYear, Opt.DebMonth, Opt.DebJour) ;
DecodeDate(ENT^.Fin,   Opt.Year,    Opt.Month,    Opt.MaxJour) ;
Opt.MaxJour:=1 ;
Opt.ExoBal:=ENT^.Exo ;
Opt.ExoAno:=ENT^.Exo ;
Opt.DecalAno:=0 ;
Opt.WithAux:=ENT^.bAux ;
if Opt.WithAux then Opt.PlanBal:='AUX'
               else Opt.PlanBal:='BAL' ;
Opt.WithRes:=TRUE ;
Opt.bNoBilan:=FALSE ;
Opt.bCalcRes:=FALSE ;
// Monnaie de saisie
Dev.Code:=V_PGI.DevisePivot ;
if Dev.Code<>'' then GetInfosDevise(Dev) ;
Dev.Taux:=GetTaux(Dev.Code, Dev.DateTaux, EncodeDate(Opt.Year, Opt.Month, 1)) ;
end ;

procedure TFImpBds.AjoutHistoBal(IL : PBDSLIGNE; ENT : PBDSENTETE) ;
var Q : TQuery ; NumCompte, NumAux : string ; Ecr : TZF ; bCreate : Boolean ;
begin
NumCompte:='' ; NumAux:='' ;
if IL^.Compte[1] in ['0', '9', 'C', 'F'] then
  begin
  NumAux:=BourreEtLess(IL^.Compte, fbAux) ;
  NumAux:=CptRemplace(NumAux, Sisco.CharRT) ;
  Q:=OpenSQL('SELECT T_COLLECTIF FROM TIERS WHERE T_AUXILIAIRE="'+NumAux+'"', TRUE) ;
  if not Q.EOF then NumCompte:=Q.Fields[0].AsString ;
  Ferme(Q) ;
  if NumCompte='' then Exit ;
  end else
  begin
  NumCompte:=BourreEtLess(IL^.Compte, fbGene) ;
  NumCompte:=CptRemplace(NumCompte, Sisco.CharRG) ;
  end ;
bCreate:=FALSE ;
Q:=OpenSQL('SELECT * FROM HISTOBAL WHERE HB_TYPE="'+Opt.TypeBal+'"'
          +' AND HB_PLAN="'+Opt.PlanBal+'"'
          +' AND HB_EXERCICE="'+Opt.ExoBal+'"'
          +' AND HB_DATE1="'+USDateTime(EncodeDate(Opt.DebYear, Opt.DebMonth, Opt.DebJour))+'"'
          +' AND HB_DATE2="'+USDateTime(EncodeDate(Opt.Year,    Opt.Month,    Opt.MaxJour))+'"'
          +' AND HB_COMPTE1="'+NumCompte+'"'
          +' AND HB_COMPTE2="'+NumAux+'"', FALSE) ;
if Q.EOF then bCreate:=TRUE else
  begin
  Q.Edit ;
  Ecr:=TZF.Create('HISTOBAL', nil, -1) ;
  //Ecr.PutValue('HB_DEBIT',  IL^.Debit) ;
  //Ecr.PutValue('HB_CREDIT', IL^.Credit) ;
  Ecr.SetMontants(IL^.Debit, IL^.Credit, Dev, FALSE) ;
  Q.FindField('HB_DEBIT').AsFloat          := Q.FindField('HB_DEBIT').AsFloat        + Ecr.GetValue('HB_DEBIT') ;
  Q.FindField('HB_CREDIT').AsFloat         := Q.FindField('HB_CREDIT').AsFloat       + Ecr.GetValue('HB_CREDIT') ;
  Q.FindField('HB_DEBITCONTRE').AsFloat    := Q.FindField('HB_DEBITCONTRE').AsFloat  + Ecr.GetValue('HB_DEBITCONTRE') ;
  Q.FindField('HB_CREDITCONTRE').AsFloat   := Q.FindField('HB_CREDITCONTRE').AsFloat + Ecr.GetValue('HB_CREDITCONTRE') ;
  Q.FindField('HB_DEBITDEV').AsFloat       := Q.FindField('HB_DEBITDEV').AsFloat     + Ecr.GetValue('HB_DEBITDEV') ;
  Q.FindField('HB_CREDITDEV').AsFloat      := Q.FindField('HB_CREDITDEV').AsFloat    + Ecr.GetValue('HB_CREDITDEV') ;
  Ecr.Free ;
  Q.Post ;
  end ;
Ferme(Q) ;
if bCreate then
  begin
  Q:=OpenSQL('SELECT * FROM HISTOBAL WHERE HB_TYPE="'+W_W+'"', FALSE) ;
  Q.Insert ; InitNew(Q) ;
  Q.FindField('HB_TYPE').AsString          := Opt.TypeBal ;
  Q.FindField('HB_PLAN').AsString          := Opt.PlanBal ;
  Q.FindField('HB_EXERCICE').AsString      := Opt.ExoBal ;
  Q.FindField('HB_DEVISE').AsString        := Dev.Code ;
  Q.FindField('HB_DATE1').AsDateTime       := EncodeDate(Opt.DebYear, Opt.DebMonth, Opt.DebJour) ;
  Q.FindField('HB_DATE2').AsDateTime       := EncodeDate(Opt.Year,    Opt.Month,    Opt.MaxJour) ;
  Q.FindField('HB_COMPTE1').AsString       := NumCompte ;
  Q.FindField('HB_COMPTE2').AsString       := NumAux ;
  Ecr:=TZF.Create('HISTOBAL', nil, -1) ;
  //Ecr.PutValue('HB_DEBIT',  IL^.Debit) ;
  //Ecr.PutValue('HB_CREDIT', IL^.Credit) ;
  Ecr.SetMontants(IL^.Debit, IL^.Credit, Dev, FALSE) ;
  Q.FindField('HB_DEBIT').AsFloat          := Ecr.GetValue('HB_DEBIT') ;
  Q.FindField('HB_CREDIT').AsFloat         := Ecr.GetValue('HB_CREDIT') ;
  Q.FindField('HB_DEBITCONTRE').AsFloat    := Ecr.GetValue('HB_DEBITCONTRE') ;
  Q.FindField('HB_CREDITCONTRE').AsFloat   := Ecr.GetValue('HB_CREDITCONTRE') ;
  Q.FindField('HB_DEBITDEV').AsFloat       := Ecr.GetValue('HB_DEBITDEV') ;
  Q.FindField('HB_CREDITDEV').AsFloat      := Ecr.GetValue('HB_CREDITDEV') ;
  Ecr.Free ;
  Q.FindField('HB_ETABLISSEMENT').AsString := VH^.EtablisDefaut ;
  Q.FindField('HB_SOCIETE').AsString       := V_PGI.CodeSociete ;
  Q.Post ;
  Ferme(Q) ;
  end ;  
end ;

function TFImpBds.FillEntete(St : string; E : PBALENTETE) : Integer ;
var k : Integer ;
begin
k:=1 ;
E^.code   :=Copy(St, k, SizeOf(E^.code)   -1) ; k:=k+SizeOf(E^.code)   -1 ;
E^.data   :=Copy(St, k, SizeOf(E^.data)   -1) ; k:=k+SizeOf(E^.data)   -1 ;
E^.datm   :=Copy(St, k, SizeOf(E^.datm)   -1) ; k:=k+SizeOf(E^.datm)   -1 ;
E^.filler1:=Copy(St, k, SizeOf(E^.filler1)-1) ; k:=k+SizeOf(E^.filler1)-1 ;
E^.cpnbc  :=Copy(St, k, SizeOf(E^.cpnbc)  -1) ; k:=k+SizeOf(E^.cpnbc)  -1 ;
E^.cpnbg  :=Copy(St, k, SizeOf(E^.cpnbg)  -1) ; k:=k+SizeOf(E^.cpnbg)  -1 ;
E^.debDad :=Copy(St, k, SizeOf(E^.debDad) -1) ; k:=k+SizeOf(E^.debDad) -1 ;
E^.finDad :=Copy(St, k, SizeOf(E^.finDad) -1) ; k:=k+SizeOf(E^.finDad) -1 ;
E^.result :=Copy(St, k, SizeOf(E^.result) -1) ; k:=k+SizeOf(E^.result) -1 ;
E^.duree  :=Copy(St, k, SizeOf(E^.duree)  -1) ; k:=k+SizeOf(E^.duree)  -1 ;
E^.analy  :=Copy(St, k, SizeOf(E^.analy)  -1) ; k:=k+SizeOf(E^.analy)  -1 ;
E^.aux    :=Copy(St, k, SizeOf(E^.aux)    -1) ; k:=k+SizeOf(E^.aux)    -1 ;
E^.version:=Copy(St, k, SizeOf(E^.version)-1) ; k:=k+SizeOf(E^.version)-1 ;
E^.maj    :=Copy(St, k, SizeOf(E^.maj)    -1) ; k:=k+SizeOf(E^.maj)    -1 ;
if StrToInt(E^.version)>=ID_MAJVERSION then
  begin
  E^.monnaie:=Copy(St, k, SizeOf(E^.monnaie)-1) ; k:=k+SizeOf(E^.monnaie)-1
  end else E^.monnaie:='' ;
E^.filler2:=Copy(St, k, SizeOf(E^.filler2)-1) ; k:=k+SizeOf(E^.filler2)-1 ;
Result:=StrToInt(E^.version) ;
end ;

procedure TFImpBds.FillCompte(St : string; C : PBALCOMPTE) ;
var k : Integer ;
begin
k:=1 ;
C^.compte   :=Copy(St, k, SizeOf(C^.compte)   -1) ; k:=k+SizeOf(C^.compte)   -1 ;
C^.code     :=Copy(St, k, SizeOf(C^.code)     -1) ; k:=k+SizeOf(C^.code)     -1 ;
C^.filler1  :=Copy(St, k, SizeOf(C^.filler1)  -1) ; k:=k+SizeOf(C^.filler1)  -1 ;
C^.libelle  :=Copy(St, k, SizeOf(C^.libelle)  -1) ; k:=k+SizeOf(C^.libelle)  -1 ;
C^.typec    :=Copy(St, k, SizeOf(C^.typec)    -1) ; k:=k+SizeOf(C^.typec)    -1 ;
C^.precedent:=Copy(St, k, SizeOf(C^.precedent)-1) ; k:=k+SizeOf(C^.precedent)-1 ;
C^.filler2  :=Copy(St, k, SizeOf(C^.filler2)  -1) ; k:=k+SizeOf(C^.filler2)  -1 ;
end ;

procedure TFImpBds.FillLigne(St : string; L : PBALLIGNE; Version : Integer) ;
var k : Integer ;
begin
k:=1 ;
L^.compte1   :=Copy(St, k, SizeOf(L^.compte1)  -1) ; k:=k+SizeOf(L^.compte1)  -1 ;
L^.code      :=Copy(St, k, SizeOf(L^.code)     -1) ; k:=k+SizeOf(L^.code)     -1 ;
L^.data      :=Copy(St, k, SizeOf(L^.data)     -1) ; k:=k+SizeOf(L^.data)     -1 ;
L^.datm      :=Copy(St, k, SizeOf(L^.datm)     -1) ; k:=k+SizeOf(L^.datm)     -1 ;
if Version<0 then Exit ;
L^.stat1     :=Copy(St, k, SizeOf(L^.stat1)    -1) ; k:=k+SizeOf(L^.stat1)    -1 ;
L^.filler1   :=Copy(St, k, SizeOf(L^.filler1)  -1) ; k:=k+SizeOf(L^.filler1)  -1 ;
L^.dtfin     :=Copy(St, k, SizeOf(L^.dtfin)    -1) ; k:=k+SizeOf(L^.dtfin)    -1 ;
L^.stat2     :=Copy(St, k, SizeOf(L^.stat2)    -1) ; k:=k+SizeOf(L^.stat2)    -1 ;
L^.compte2   :=Copy(St, k, SizeOf(L^.compte2)  -1) ; k:=k+SizeOf(L^.compte2)  -1 ;
L^.debit     :=Copy(St, k, SizeOf(L^.debit)    -1) ; k:=k+SizeOf(L^.debit)    -1 ;
L^.credit    :=Copy(St, k, SizeOf(L^.credit)   -1) ; k:=k+SizeOf(L^.credit)   -1 ;
L^.slddebit  :=Copy(St, k, SizeOf(L^.slddebit) -1) ; k:=k+SizeOf(L^.slddebit) -1 ;
L^.sldcredit :=Copy(St, k, SizeOf(L^.sldcredit)-1) ; k:=k+SizeOf(L^.sldcredit)-1 ;
if Version>=ID_MAJVERSION then
begin
L^.debeur    :=Copy(St, k, SizeOf(L^.debeur)   -1) ; k:=k+SizeOf(L^.debeur)   -1 ;
L^.creeur    :=Copy(St, k, SizeOf(L^.creeur)   -1) ; k:=k+SizeOf(L^.creeur)   -1 ;
end else
begin
L^.debeur    :='' ;
L^.creeur    :='' ;
end ;
L^.filler2   :=Copy(St, k, SizeOf(L^.filler2)  -1) ; k:=k+SizeOf(L^.filler2)  -1 ;
end ;

function TFImpBds.ReadDAD(bVerifCompte : Boolean; ENT : PBDSENTETE) : Boolean ;
var i : Integer ; St : string ; bDebut, bFin, bEntete, bCompte : Boolean ;
    IE : PBDSENTETE ; IC : PBDSCOMPTE ; IL : PBDSLIGNE ; F : TextFile ;
    TE : PBALENTETE ; TC : PBALCOMPTE ; TL : PBALLIGNE ;
begin
Result:=TRUE ; bDebut:=FALSE ; bFin:=FALSE ; bEntete:=FALSE ; bCompte:=FALSE ;
IE:=nil ; IC:=nil ; IL:=nil ;
AssignFile(F, FileName.Text) ;
{$I-} Reset (F) ; {$I+}
if IoResult<>0 then begin Result:=FALSE ; Exit ; end ;
InitSisco ;
if bVerifCompte then InitMove(NbLignes, GetMessageRC(RC_VERIF))
                else InitMove(NbLignes, GetMessageRC(RC_IMPORT)) ;
New(TE) ;
New(TC) ;
New(TL) ;
New(IL) ;
New(IC) ;
repeat
  if TestBreak then
    begin
    if not bVerifCompte then V_PGI.IoError:=oeSaisie ;
    Break ;
    end ;
  ReadLn(F, St) ;
  MoveCur(FALSE) ;
  if bFin then Continue ;
  if (Copy(St, 1, Length(ID_DEBUT))=ID_DEBUT) then
    begin
    //(Length(St)=SizeOf(R_S_DAD_DEBUT)-NB_R_S_DAD_DEBUT) and
    bDebut:=TRUE ;
    Continue ;
    end ;
  if (Copy(St, 1, Length(ID_FIN))=ID_FIN) then
    begin
    //(Length(St)=SizeOf(R_S_DAD_FIN)-NB_R_S_DAD_FIN) and
    bFin:=TRUE ;
    Continue ;
    end ;
  if (Copy(St, 1, Length(ID_ENTETE))=ID_ENTETE) then
    begin
    if not bVerifCompte then begin bEntete:=TRUE ; Continue ; end ;
    //(Length(St)=SizeOf(R_S_DAD_ENT)-NB_R_S_DAD_ENT) and
    if not bDebut then Continue ;
    bEntete:=TRUE ;
    New(IE) ;
    IE^.Ver    :=FillEntete(St, TE) ;
    IE^.Bal    :=ImpStrToDate('00'+TE^.datm+TE^.data) ;
    IE^.CodeBal:=TE^.data+TE^.datm ;
    IE^.Debut  :=ImpStrToDate(TE^.debdad) ;
    IE^.Fin    :=ImpStrToDate(TE^.findad) ;
    IE^.Duree  :=StrToInt(TE^.duree);
    IE^.bAux   :=(TE^.aux='O') ;
    IE^.bAna   :=(TE^.analy='O')  ;
    IE^.Monnaie:=TE^.monnaie ;
    IE^.Exo    :=IsOKImport(IE^.Debut, IE^.Fin) ;
    Entete.Add(IE) ;
    Continue ;
    end ;
  if (Length(St)+1=SizeOf(R_S_DAD_CPTE)-NB_R_S_DAD_CPTE) and (bEntete) then
    begin
    if not bVerifCompte then begin bCompte:=TRUE ; Continue ; end ;
    //if (bCompte) and (IC<>nil) then begin Dispose(IC) ; IC:=nil ; end ;
    bCompte:=TRUE ;
    FillCompte(St, TC) ;
    IC^.Compte :=TC^.compte ;
    IC^.Code   :=TC^.code ;
    IC^.Libelle:=TC^.libelle ;
    IC^.Typec  :=TC^.typec ;
    // Vérification Compte
    VerifCompte(IC) ;
    Continue ;
    end ;
  if (bCompte) then
    begin
    if bVerifCompte then Continue ;
    //(Length(St)=SizeOf(R_S_DAD_LIG)-NB_R_S_DAD_LIG) and
    FillLigne(St, TL, -1) ; // Ne charge que la partie utile (pour version)
    for i:=0 to Entete.Count-1 do
      begin
      if PBDSENTETE(Entete.Items[i])^.CodeBal=TL^.data+TL^.datm then
        if PBDSENTETE(Entete.Items[i])^.Exo='' then (*Pas d'exercice*)Continue else (*OK*)break ;
      end ;
    FillLigne(St, TL, PBDSENTETE(Entete.Items[i])^.Ver) ;
    IL^.Compte  :=TL^.compte1 ;
    IL^.DateC   :=ImpStrToDate('00'+TL^.datm+TL^.data) ;
    IL^.Stat1   :=TL^.stat1 ;
    IL^.DateFin :=ImpStrToDate('00'+TL^.dtfin) ;
    IL^.Stat2   :=TL^.stat2 ;
    IL^.Compte2 :=TL^.compte2 ;
    IL^.Debit   :=ImpStrToDouble(TL^.debit) ;
    IL^.Credit  :=ImpStrToDouble(TL^.credit) ;
    IL^.DebEur  :=ImpStrToDouble(TL^.debeur) ;
    IL^.CreEur  :=ImpStrToDouble(TL^.creeur) ;
    IL^.Exo     :=PBDSENTETE(Entete.Items[i])^.Exo ;
    // Ajout ligne
    AjoutHistoBal(IL, ENT) ;
    Continue ;
    end ;
until EOF(F) ;
Dispose(TE) ;
Dispose(TC) ;
Dispose(TL) ;
Dispose(IL) ;
Dispose(IC) ;
//if (bCompte) and (IC<>nil) then Dispose(IC) ;
CloseFile(F) ;
FiniMove ;
end ;

procedure TFImpBds.WriteBal ;
begin
InitBal(PBDSENTETE(Entete.Items[iEntete])) ;
ReadDAD(FALSE, PBDSENTETE(Entete.Items[iEntete])) ;
end ;

procedure TFImpBds.BValiderClick(Sender: TObject);
var F : TextFile ; St : string ; i : Integer ; bRes : Boolean ;
begin
bPlan:=TRUE ; bPrint:=FALSE ; BImprimer.Enabled:=FALSE ;
InitImport ;
Erreur:=TList.Create ;
Entete:=TList.Create ;
if (FileName.Text='') then begin PrintMessageRC(RC_BADFILENAME) ; Exit ; end ;
if (FFormat.ItemIndex=-1) then begin PrintMessageRC(RC_BADFORMAT) ; Exit ; end ;
TabResult.Enabled:=TRUE ;
Pages.ActivePage:=TabResult ;
AssignFile(F, FileName.Text) ;
{$I-} Reset (F) ; {$I+}
if IoResult<>0 then begin PrintMessageRC(RC_BADPATH) ; Exit ; end ;
ActivePanels(Self, FALSE, FALSE) ;
NbImport:=0 ; NbErreur:=0 ; NbLu:=0 ; NbLignes:=0 ;
bArreter:=FALSE ;
repeat
  if TestBreak then Break ;
  Readln(F,St) ; NbLignes:=NbLignes+1 ;
until EOF(F) ;
CloseFile(F) ;
if bArreter then Exit ;
ReadDAD ;
if bArreter then Exit ;
// Ecriture des balances
for i:=0 to Entete.Count-1 do
  begin
  if not bPlan then Break ;
  if PBDSENTETE(Entete.Items[i])^.Exo='PFU' then
    begin
    bPrint:=TRUE ;
    GS.Cells[RE_CODE,    Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Debut) ;
    GS.Cells[RE_LIBELLE, Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Fin) ;
    GS.Cells[RE_COMMENT, Row]:=GetMessageRC(RC_BALEXIST) ;
    if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
    Row:=Row+1 ;
    Continue ;
    end ;
  if PBDSENTETE(Entete.Items[i])^.Exo='' then
    begin
    bPrint:=TRUE ;
    GS.Cells[RE_CODE,    Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Debut) ;
    GS.Cells[RE_LIBELLE, Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Fin) ;
    GS.Cells[RE_COMMENT, Row]:=GetMessageRC(RC_BALBADDATE) ;
    if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
    Row:=Row+1 ;
    Continue ;
    end ;
  iEntete:=i ;
  bRes:=(Transactions(WriteBal, 3)=oeOk) ;
  if not bRes then
    begin
    FiniMove ;
    bPrint:=TRUE ;
    GS.Cells[RE_CODE,    Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Debut) ;
    GS.Cells[RE_LIBELLE, Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Fin) ;
    GS.Cells[RE_COMMENT, Row]:=GetMessageRC(RC_BADBALWRITE) ;
    if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
    Row:=Row+1 ;
    PrintMessageRC(RC_BADWRITE) ;
    Break ;
    end else
    begin
    bPrint:=TRUE ;
    GS.Cells[RE_CODE,    Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Debut) ;
    GS.Cells[RE_LIBELLE, Row]:=DateToStr(PBDSENTETE(Entete.Items[i])^.Fin) ;
    GS.Cells[RE_COMMENT, Row]:=GetMessageRC(RC_BALIMPORT) ;
    if GS.RowCount>MAX_ROWS then GS.RowCount:=GS.RowCount+1 ;
    Row:=Row+1 ;
    end ;
  end ;
if bPrint then
  begin
  GS.RowCount:=Row+1 ;
  bImprimer.Enabled:=TRUE ;
  end ;  
// C'est fini
ActivePanels(Self, TRUE, FALSE) ;
if (NbImport=0) then Exit ;
if FDelFile.Checked then DeleteFile(FileName.text) ;
Screen.cursor:=SynCrDefault ;
end ;

procedure TFImpBds.FFormatChange(Sender: TObject);
begin
if FFormat.Value='BDS' then Sauve.FilterIndex:=1 ;
//if FFormat.Value='CEG' then Sauve.FilterIndex:=2 ;
end;

procedure TFImpBds.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFImpBds.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TFImpBds.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(GS, nil, Caption, '') ;
end;

end.
