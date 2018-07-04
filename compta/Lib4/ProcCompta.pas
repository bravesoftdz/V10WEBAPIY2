unit ProcCompta;

interface

uses
  TypeEnt1, HEnt1, Controls;

(*
procedure InitPopup( F : TForm) ;
procedure PurgePopup( PP : TPopupMenu ) ;
procedure PopZoom ( BM : TBitBtn ; POPZ : TPopupMenu ) ;
procedure PopZoom97 ( BM : TToolbarButton97 ; POPZ : TPopupMenu ) ;
function  CreerLigPop ( B : TBitBtn ; Owner : TComponent ; ShortC : Boolean ; C : Char ) : TMenuItem ;
{$ENDIF}
Function  BourreEtLess ( St : String ; LeType : TFichierBase ; ForceLg : Integer = 0) : string ;
function BourreOuTronque(St : String ; fb : TFichierBase) : String ;
function  BourreLaDoncSurLesComptes(st : string ; c : string = '') : string;
function  CompteDansLeIntervalle(Quoi, Deb, Fin : string) : boolean;
// Fonctions de conversions
Function  AxeToTz ( Ax : String ) : TZoomTable ;
function  AxeToDataType(CodeAxe : string) : string; {<=> AxeToTz pour les DataType et non plus les ZoomTable}
Function  NatureToTz (Nat : String) : TZoomTable ;
function  NatureToDataType(CodeTable : string) : string;{<=> AxeToTz pour les DataType et non plus les ZoomTable}
Function  StringToTz(s : String) : TZoomTable ;
Function  tzToChampNature( tz : TZoomTable ; AvecPrefixe : Boolean) : String ;
Function  AxeToFbBud ( Axe : String ) : TFichierBase ;
function StrToTA(psz : String) : TActionFiche;
function TAToStr(TA : TActionFiche) : String;

Function  BourreLess ( St : String ; LeType : TFichierBase ) : string ;
Procedure AvertirMultiTable (FTypeTable : String ) ;
Function  AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
function  fileTemp(const aExt: String): String;
Procedure ChangeMask (C : THNumEdit ; Decim : Integer ; Symbole : String) ;
Procedure ATTRIBSOCIETE( T : TDataSet ; P : String) ;
{$IFNDEF EAGLSERVER}
Function  InitPath : String ;
{$ENDIF}
procedure AfficheLeSolde (T : THNumEdit ; TD,TC : Double) ;
function  INSERE( st,st1 : string ; deb,long : integer) : string ;
function  BISSEXE (Annee : Word ) : Boolean ;

FUNCTION  TVA2ENCAIS(ModeTVA,Tva : String3 ; Achat : Boolean; RG : boolean=false; FarFae : boolean=false) : String ;
FUNCTION  TVA2CPTE(ModeTVA,Tva : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
FUNCTION  TVA2TAUX(ModeTVA,Tva : String3 ; Achat : Boolean) : Real ;
FUNCTION  TPF2ENCAIS(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
FUNCTION  TPF2CPTE(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
FUNCTION  TPF2TAUX(ModeTVA,Tpf : String3 ; Achat : Boolean) : Real ;
{$ENDIF}
FUNCTION  TOTDIFFERENT(X1,X2 : Double) : BOOLEAN ;
{$IFNDEF NOVH}
FUNCTION  HT2TVA ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  HT2TPF ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  HT2TTC ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  TTC2HT ( TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  TTC2TPF ( TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  TTC2TVA (TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
{$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE}
Procedure ChgMaskChamp (C : TFloatField ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) ;
{$ENDIF}
Function  MontantToStr (Montant : Double ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) : string ;
Procedure CorrespToCombo(Var CB : THValComboBox ; Fb : TFichierBase );
Procedure CorrespToCodes(Plan : THValComboBox ; Var C1, C2 : TComboBox);
Procedure RuptureToCodes(Plan : THValComboBox ; Var C1, C2 : TComboBox ; Fb : TFichierBase );
Procedure CodesRuptSec(Var LCodes : TStringlist ; Plan,Axe : String) ;
Function  PrintSolde(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean) : String ;
Function  PrintSoldeFormate(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean ; FormMont : String) : String ;
Function  PrintSolde2(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean ; FormMont : THValCombobox) : String ;
Function  PrintEcart(TD,TC : DOUBLE ; Decim : Integer ; DebitPos : Boolean) : String ;
function  MoinsSymbole(LeSolde : String) : double ;
Function  SqlCptInterdit(LeChamp : String ; Var PourSql : String ; ZoneInterdit : TEdit) : Boolean ;
Function  Fourchette( Var St : String) : String ;
Procedure PremierDernier(fb : TFichierBase ; Var Cpt1,Cpt2 : String ) ;
Function  SQLPremierDernier(fb : TFichierbase ; Prem : Boolean) : String ;
Procedure PremierDernierRub(ZoomTable : TZoomTable ; SynPlus : String ; Var Cpt1,Cpt2 : String ) ;
Function  SQLPremierDernierRub(ZoomTable : TZoomTable ; SynPlus : String ; Prem : Boolean) : String ;
function  EstConfidentiel ( Conf : String ) : boolean ;
function  EstSQLConfidentiel ( StTable : String ; Cpte : String17 ) : boolean ;
{$IFNDEF NOVH}
function  DateBudget(CPER : THValComboBox) : String ;
Procedure ChargeComboTableLibre(Cod : String ; DesValues,DesItems : HTStrings) ;
Procedure GetLibelleTableLibre(Pref : String ; LesLib : HTStringList) ;
Function  GetLibelleTabLibCommun(Cod : String) : String ;
FUNCTION  QUELEXODTBUD(DD : TDateTime) : String ;
Function  ParamEuroOk : Boolean ;
Function  IncPeriode ( Periode : integer ) : integer;
Function  GetDateTimeFromPeriode ( Periode : integer ) : TDateTime;
procedure DirDefault(Sauve : TSaveDialog ; FileName : String) ;
Function  PieceSurFolio(Jal : String) : Boolean ;
Function  AuMoinsUneImmo : Boolean ;
Procedure SetTousCombo ( TH : THValComboBox ) ;
procedure MajBeforeDispatch(Num: Integer=0) ;
Function LienS1 : Boolean ;
Function LienS3 : Boolean ;
Function LienS1S3 : Boolean ;
{$IFNDEF SPEC302}
{$IFNDEF NOVH}
Procedure ChargeMPACC ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure CreerDeviseTenue(LaMonnaie : String) ;
{$ENDIF}
Function  CompareTL(St1,St2 : String) : Boolean ;
{$IFNDEF NOVH}
Function  CptDansProfil(Gen,Aux : String ; Ind : tProfilTraitement) : Boolean ;
Function  MonProfilOk(Qui : tProfilTraitement) : Boolean ;
{$ENDIF}
Function _ChampExiste(NomTable,NomChamp : String) : Boolean ;
{$IFDEF EURO}
Procedure EcritProcEuro(St : String) ;
{$ENDIF}
//Function OkSynchro : Boolean ;
FUNCTION GoINSERE( st,st1 : string ; deb,long : integer) : string ;
{$IFNDEF NOVH}
Function HalteAuFeu : Boolean ;
Procedure DemandeStop(St : String) ;
Function EstSpecif(St : String) : Boolean ;
{$ENDIF}
Function StopEcart (DD : tDateTime ; Dev : String = '') : Boolean ;
//Function StopDevise (DD : tDateTime ; Dev : String ; ModeOppose : Boolean=FALSE) : Boolean ;
// Fonctions communes avec la trésorerie
function calcIBAN	(pays : string ; RIB : string) : string ;
function calcRIB (pays : string ; banque : string ; guichet : string ; compte : string; cle : string) : string ;
{ Retraitement suite suppression contre-valeur }
{$IFDEF COMPTA}
//procedure SupprimeChampsEcartDansLesListes;
// function SupprimeEcartConversionEuro : boolean;
{$IFDEF COMPTA}
function CPEstTvaActivee : Boolean;
function CPGetMontantTVA(vStCodeTva : string ; vDateDebut, vDateFin : TDateTime) : Double;
Procedure MAJHistoparam (NomTable, Data :string);
{$ENDIF}
{$IFDEF COMSX}
Procedure MAJHistoparam (NomTable, Data :string);
{$ENDIF}
Function GetTenuEuro : Boolean;
Function GetRecupPCL : Boolean;
Function GetFromPCL : Boolean;
Function GetRecupComSx : Boolean;
Function GetRecupSISCOPGI : Boolean;
Function GetRecupLTL : Boolean;
Function GetRecupCegid : Boolean;
Function GetCPIFDEFCEGID : Boolean;




*)








implementation

uses
{$IFDEF EAGLSERVER}
{$IFDEF NOVH}
  ULibCpContexte ,
{$ENDIF}
{$ENDIF}
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  UTob, HCtrls,
  SysUtils, StdCtrls, Ent1;



end.
