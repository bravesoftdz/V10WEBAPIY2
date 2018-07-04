Unit SAISUTIL;

interface

uses
    SysUtils,
    Forms,
    Classes,
    HCtrls,
    HQry,
    Dialogs,
    Menus,
    Windows,
    Buttons,
    ExtCtrls,
    Grids,
    hent1,
    HMsgBox,
    Controls,
    ComCtrls,
    ed_tools,
    LicUtil,
{$IFNDEF EAGLSERVER}
    EdtEtat,
{$ENDIF}
{$IFDEF EAGLCLIENT}
    MaineAgl,
{$ELSE}
    HDB,
    DBGrids,
    DB,
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$IFNDEF EAGLSERVER}
    FE_Main,
    EdtDoc,
{$ENDIF}

    EdtRDoc,
    EdtREtat,
{$ENDIF}
    UTOB,
    UtilPGI,
{$IFDEF VER150}
    variants,
{$ENDIF}
    UtilSais,
    M3FP
  {$IFNDEF EAGLSERVER}
    ,AglInit
  {$ENDIF}
    ,ent1
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
    ,uEntCommun
    ;

Const MaxChampEcr = 150 ;

Const NMS : Boolean = True ; // New Maj Solde
      NMSB : Boolean = TRUE ; // New Maj Solde bordereau

Type TTypeEcr =(EcrGen,EcrBud,EcrGui,EcrAna,EcrAbo,EcrSais,EcrClo) ;
Type TModeSaisBud = (msbGene,msbSect,msbGeneSect,msbSectGene) ;
Type TSorteTva = (stvAchat,stvVente,stvDivers) ;
Type TExigeTva = (tvaDebit,tvaEncais,tvaMixte) ;
Type TSorteLettre = (tslAucun,tslCheque,tslTraite,tslBOR,tslVir,tslPre) ;
Type TEtat   = (Cree,Modifie,Inchange) ;
Type TStatus = (EnMemoire,DansTable,LuDansTable) ;
Type TChampsTiers = (ctRelance,ctRelFac,ctCoutTiers,ctExpMvt) ;
Type ttCtrPtr = (Ligne,PiedDC,PiedSolde,AuChoix) ;
Type TAppelLettrage = (tlMenu,tlSaisieCour,tlSaisieTreso) ;
Type TTypeDispo = (tdPMPA,tdPMA) ;
Type tModeSaisieEff = (OnEffet,OnChq,OnCB,OnBqe) ;
Type tModeSR = (srRien,srCli,srFou) ;

Type RDevise = RECORD
               Code     : String3 ;
               Libelle  : String35 ;
               Symbole  : String3 ;
               Decimale : integer ;
               DateTaux : TdateTime ;
               Taux,Quotite,MaxDebit,MaxCredit,Cotation : Double ;
               CptDebit,CptCredit              : String17 ;
               END ;

Type TLigTreso = RECORD
                 TypCtr    : ttCtrPtr ;
                 STypCtr   : String3 ;
                 Cpt       : String17 ;
                 Ventilable : Boolean ;
                 TotD,TotC : Double ;
                 DevBqe    : String3 ;
                 Collectif : Boolean ;
                 END ;

{Type TSuiviMP = (smpAucun,smpEncPreBqe,smpEncTraEdt,smpEncTraEnc,smpEncTraEsc,smpEncTraBqe,
                 smpEncTous,smpEncTraEdtNC,smpEncTraPor,smpEncPreEdt,smpEncChqPor,
                 smpEncCBPor,smpEncChqBqe,smpEncCBBqe,smpEncDiv,
                 smpDecVirBqe,(*smpDecVirBq2,smpDecVirBq3,*)smpDecVirEdt,smpDecVirEdtNC,smpDecChqEdt,smpDecChqEdtNC,smpDecBorEdt,smpDecBorEdtNC,
                 smpDecBorDec,smpDecBorEsc,smpDecTraBqe,smpDecTraPor,smpDecDiv,smpDecTous) ;}

{NE PAS MODIFIER L'ORDRE }
Type TSuiviMP = (smpAucun,
                 smpEncPreBqe  ,smpEncPreEdt  ,smpEncPreEdtNC  ,                                                        {Encaissement prélévement}
                 smpEncTraBqe  ,smpEncTraEdt  ,smpEncTraEnc,smpEncTraEsc,smpEncTraEdtNC  ,smpEncTraPor,              {Encaissement traite}
                 smpEncChqBqe  ,                                                          smpEncChqPor,              {Encaissement chèque}
                 smpEncCBBqe   ,                                                          smpEncCBPor ,              {Encaissement C.B}
                 smpEncTous    ,                                                                                     {Tous encaissement}
                 smpEncDiv     ,                                                                                     {Encaissement divers}
                 smpDecVirBqe  ,smpDecVirEdt  ,                          smpDecVirEdtNC  ,                           {Décaissement virement}
                 smpDecVirInBqe,smpDecVirInEdt,                          smpDecVirInEdtNC,                           {Décaissement virement internationale}
                 smpDecTraBqe  ,                                                          smpDecTraPor,              {Décaissement traite}
                 (*smpDecVirBq2,smpDecVirBq3,*)
                                smpDecChqEdt  ,                          smpDecChqEdtNC  ,                           {Décaisssemnt chèque}
                                smpDecBorEdt  ,smpDecBorEsc,             smpDecBorEdtNC  ,             smpDecBorDec, {Décaisssemnt B.O.R}
                 smpDecDiv,              {Décaissement divers}
                 smpDecTous,             {Tous décaissement}
                 smpCompenCli, smpCompenFou  {FP 21/02/2006 Compensation client, Compensation Fournisseur}
                 ) ;

Type TMSEncaDeca = Record
                   MultiSessionEncaDeca,Sessionfaite : Boolean ;
                   ModeleMultiSession : String ;
                   Spooler : Boolean ;
                   RepSpooler : String ;
                   RacineSpooler : String ;
                   SoucheSpooler : Integer ;
                   XFichierSpooler : Boolean ;
                   End ;
Type tFormuleRefCCMP = Record
                       Ref1,Lib1,Ref2,Lib2,RefExt1,RefLib1,RefExt2,RefLib2 : String ;
                       End ;


Type RMVT = RECORD
            Axe,Etabl,Jal,Exo,CodeD,Simul,Nature,LeGuide,TypeGuide : String3 ;
            General,Section : String17 ;
            DateC,DateTaux,LastDateEche,DateE : TDateTime ;
            Num,LastNumCreat : Longint ;
            TauxD   : Double ;
            Valide,EtapeRegle,FromGuide,ANouveau,SaisieGuidee : boolean ;
            NumLigne,NumEche,NumLigVisu : integer ;
            CodeLettrage : string;
            {utilisé uniquement en saisie de trésorerie, aucune incidence sinon}
            Treso,MajDirecte,Effet : Boolean ;
            Souche,FormatCFONB,Document,EnvoiTrans,ModeSaisieJal,GroupeEncadeca,MPGUnique : String3 ;
            MSED : TMSEncaDeca ;
            Indic  : String[1] ;
            TypeSaisie : String[2] ;
            ExportCFONB,Bordereau,Globalise : boolean ;
            SorteLettre                     : TSorteLettre ;
            smp                             : TSuiviMP ;
            TIDTIC                          : Boolean ;
            ForceModif : Boolean ;
            BloquePieceImport : Boolean ;
            NumEncaDeca : String ;
            ParCentral : boolean ;
            FormuleRefCCMP : tFormuleRefCCMP ;
            {$IFDEF GCGC}
            Historique : boolean ;
            {$ENDIF GCGC}
            DocType : string;
            DocNumber : Integer;
            End ;

Type   P_MV = Class
            R : RMVT ;
            End ;

Type RLETTR = RECORD
              General,Auxiliaire   : String17 ;
              DeviseMvt            : String3 ;
              CodeLettre           : String4 ;
              LettrageDevise,ToutSelDel : boolean ;
              Appel                : TAppelLettrage ;
{$IFDEF EAGLCLIENT}
              GL                   : THGrid ;
{$ELSE}
              GL                   : TDBGrid ;
{$ENDIF}
              CritMvt,CritDev,RefLettrage : String ;
              Ident                : RMVT ;
              NowStamp             : TDateTime ;
              SansLesPartiel,Distinguer : Boolean ;
              Confondre            : boolean ;
              LettrageEnSaisie     : boolean ;
              TOBResult            : TOB ;
              CritEtatLettrage     : string ;
              END ;

Type TT_V = Record
            V : Variant ;
            Etat : TEtat ;
            end ;

Type T_V  = Array[0..MaxChampEcr+1] of TT_V ;
     P_TV = Class
            F : T_V ;
            End ;

type TSigne = (sgEgal,sgOppose,sgNeutre) ;

 // type d'ecriture possible ( champs E_QUALIFORIGINE )
 // TESaisie : ecriture saisie E_QUALIFORIGINE = ''
 // TERegulLettrage : ecriture de regul de lettrage E_QUALIFORIGINE = 'REG'
 // TEEcartEuroLettrage : ecart euro en lettrage E_QUALIFORIGINE = 'CON'
 // TEEcartEuro : ecart euro en saisie = 'ECC'
 // TETreso : ecriture generee par la saisie de tresorerie = 'TRE'
 // TESimplifie : ecriture simplifie genere en lettrage = 'SIM'
 TTypeEcriture = ( TESaisie , TERegulLettrage , TEEcartEuroLettrage , TETreso , TEEcartEuro , TESimplifie ) ;

 TOBM = class(TOB)
  private
   SStatus          : TStatus;
   //SEtat            : TEtat;
   FM               : HTStrings;  // mettre une property sur le M
  // FTypeEcr         : TTypeEcriture;
   procedure SetStatus ( Value : TStatus ) ;
   procedure SetEtat ( Value : TEtat ) ;
   function  GetEtat : TEtat ;
  protected
   procedure PutDefautEcr;
   procedure PutDefautBud;
   procedure PutDefautAna ( Axe : String3 ) ;
   procedure PutDefaut ( Axe : String3 ) ;
   function  GetT ( Index : integer ) : TT_V;
   procedure SetT ( Index : integer ; Value : TT_V );
   function  GetDateComptable : TDateTime;
   procedure SetDateComptable( Value : TDateTime);
   function  GetModeSaisie : string;
   procedure SetModeSaisie (const Value : string);
   function  GetTypeEcr : TTypeEcriture;
  public
   Ident : TTypeEcr;
   NbC   : integer;
   NbDiv : integer;
   LC    : HTStrings;
   CompS : Boolean;
   CompE : boolean ;
 {$IFDEF GCGC}
   NatureG, TypeArticle : String3;
 {$ENDIF}

   constructor Create ( iecr : TTypeEcr; Axe : String3; Defaut : Boolean ; LeParent : TOB = nil ) ; reintroduce ;
   destructor  Destroy;override;

   function  GetMvt ( Nom : string ) : Variant;
   procedure PutMvt ( Nom : string; Value : Variant ) ;
   procedure ChargeMvt ( F : TDataSet ) ;
   procedure ChargeMvtP( F : TDataSet ; vStPrefixe : string = '' ) ;
   procedure HistoMontants;
   procedure EgalChamps ( TT : TDataSet ) ;
   procedure EgalChampsTob( TT : TOB ) ; {JP 24/10/07}
   procedure MajLesDates ( Action : TActionFiche ) ;
   procedure MajEtat ( Value : TEtat ) ;
   function  STPourUPDATE : string;
   // GC*
   function  UpdateDB : Boolean;
   function  OkPeriodeSemaine : Boolean;
   function  OkModeSaisie( vJModeSaisie : string ) : Boolean;
   function  OkEcrANouveau (vJNatureJal : string ) : Boolean;
   //
   property  Status : TStatus read SStatus write SetStatus;
   property  Etat : TEtat read GetEtat write SetEtat;
   procedure PutDefautDivers;
   procedure SetMontants ( XD, XC : Double; DEV : RDEVISE; Force : boolean ) ;
   procedure SetMontantsBUD ( XD, XC : Double ) ;
   procedure SetCotation ( DD : TDateTime ) ;
{$IFNDEF NOVH}
   procedure SetMPACC;
{$ENDIF}
   procedure TraiteLeEtat(PasEcart, AvecRegul, EcartChange, AvecConvert, ModeAuto: boolean; IndiceRegul: integer; ModeRappro: Char);
   function  GetDebitSaisi : double;
   function  GetCreditSaisi : double;

   property  M                       : HTStrings      read FM               write FM;
   property  T [ Index : integer ]   : TT_V           read GetT             write SetT ; default;
   property  DateComptable           : TDateTime      read GetDateComptable write SetDateComptable;
   property  ModeSaisie              : string         read GetModeSaisie    write SetModeSaisie;
   property  TypeEcr                 : TTypeEcriture  read GetTypeEcr;

  end;

//LG* sortie de la directive de compilation
// GP ANP
type
 tDC = record
  D, C, DE, CE, DD, CD : Double;
 end;

{$IFNDEF SANSCOMPTA}

Type TAA = Class
           Public
           L    : TList ;
           M    : TList ;
           Etat : TEtat ;
           Constructor Create ;
           Destructor Destroy ; override ;
           END ;

Type TObjAna = Class
               Sens                  : Byte ;
               General               : String17 ;
               Exercice              : String3 ;
               DateComptable         : TDateTime ;
               NumeroPiece           : Longint ;
               NumLigne              : integer ;
               NumLigneDecal         : Integer ;
               RefInterne            : String35 ;
               Libelle               : String35 ;
               RefLibre              : String35 ;
               NaturePiece           : String3 ;
               QualifPiece           : String3 ;
               RefExterne            : String35 ;
               DateRefExterne        : TDateTime ;
               Utilisateur           : String3 ;
               Controleur            : String3 ;
               Societe               : String3 ;
               Etablissement         : String3 ;
               TotalEcriture         : Double ;
               TauxDev               : Double ;
               DateTauxDev           : TDateTime ;
               Devise                : String3 ;
               QualifQte1,QualifQte2 : String3 ;
               Journal,Etabl         : String3 ;
               DernVentilType,GuideA : String3 ;
               Decimale              : byte ;
               ControleBudget        : boolean ;
               Affaire               : String ;
               ModeConf              : String[1] ;
               TotalDevise,TotalQte1,TotalQte2 : Double ;
               AA : Array[1..MaxAxe] of TAA ;
               Destructor destroy ; override ;
               END ;



{Type T_SCBOR = Class
               Identi : String3 ;
               Cpte   : String17 ;
               Debit  : Double ;
               Credit : Double ;
               NumP,NumL   : Integer ;
               DateP : tDateTime ;
               CptCegid : string ; // LG 18/04/2004
               YTCTiers : string ;
               end ;   }


Type tSoldeAN = Record
                Avant,Apres : tDC ;
                End ;

Type t_SCAN = Class
               (*
               L : tStringList ;
               EtabP,DevP,JalP : String ;
               NumP : Integer ;
               DateP : TDateTime ;
               TauxP,CotP : Double ;
               *)
               Identi : String3 ; { 'G', 'T', 'A1'..'A5', 'GS', 'B' }
               Cpte1   : String17 ;
               Cpte2   : String17 ;
               SoldeAN : tSoldeAN ;
               End ;
// FIN GG ANP

Type TDELMODIF = Class
                 NumLigne,NumEcheVent : integer ;
                 DateModification : TDateTime ;
                 Ax               : String ;
                 End ;

{$ENDIF}




Function  SENSENC ( D,C : Double ) : String3 ;
{$IFDEF EAGLCLIENT}
procedure GereSelectionsGrid ( G : THGrid ; Q : TQuery ) ;
{$ELSE}
  {$IFNDEF EAGLSERVER}
procedure GereSelectionsGrid ( G : THDBGrid ; Q : THQuery ) ;
  {$ENDIF}
{$ENDIF}
{$IFNDEF SANSCOMPTA}
procedure InitCommunObjAnal( QuelEcr : TTypeEcr ; OBM : TOBM ; OBA : TObjAna ) ;
function  RechercheLente ( St,Cpte : String ) : Variant ;
PROCEDURE RECALCULPRORATAANAL ( Pf : String ; OBA : TObjAna ; MtEcrP,MtEcrD : Double ; NumA : integer ; SensEcr : integer ) ;
  {$IFNDEF EAGLSERVER}
Function  InitVariant ( Nom : String ) : Variant ;
procedure EditEtatS5S7 (Tip,Nat,Modele : String ; AvecCreation : Boolean ; Spages : TPageControl ; StSQL,Titre : String );
  {$ENDIF}
Procedure PutMvtA ( AAi : TAA ; Var F : T_V ; Quoi : String ; Value : Variant) ;
Function  GetMvtA ( F : T_V ; Quoi : String ) : Variant ;
PROCEDURE CALCULTOTALA ( Pf : String ; OBA : TObjAna ; Var TotMP,TotMD,TotME,TotTaux : Double ; j,Lig : Integer) ;
	{$IFNDEF EAGLCLIENT}
    {$IFNDEF EAGLSERVER}
procedure EditDocumentS5S7 (Tip,Nat,Modele : String ; AvecCreation : Boolean);
  	{$ENDIF}
	{$ENDIF}
{$ENDIF}
function  GetNewNumJal ( Jal : String3; Normale : boolean; DD : TDateTime ; Cpt : string = '' ; ModeSaisie : string = '' ; vDossier : String = '' ) : Longint;
FUNCTION  GETNUM (TypeEcr : TTypeEcr ; Facturier : String3 ; Var MasqueNum : String17 ; DD : TDateTime ; vDossier : String = '' ) : LongInt ;
PROCEDURE SETNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; Num,OldN : Longint ; DD : TDateTime ; vDossier : String = '' ) ;
PROCEDURE SETINCNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; Var Num : LongInt ; DD : TDateTime ; vDossier : String = '' ) ;
PROCEDURE SETDECNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; DD : TDateTime ; vDossier : String = '' ) ;
PROCEDURE GETINFOSDEVISE ( Var DEV : RDEVISE ; vDossier : String = '') ;
Function  CaseNatJal ( NatJal : String3 ) : TZoomTable ;
Procedure ChangeFormatDevise ( FF : TForm ; Decim : integer ; Symbole : String ) ;
Procedure ChangeFormatPivot ( FF : TForm ) ;
FUNCTION  VALEURPIVOT(St : String ; Taux : Double ; Quotite : Double) : Double ;
FUNCTION GETTAUX ( Code : String3 ; Var DateTaux : TDateTime ; DateC : TDateTime ; vDossier : String = ''; WhereDateGC : string='') : Double ;
FUNCTION  QUELEXO(DateC : String) : String ;
FUNCTION  QUELEXODT(DD : TDateTime) : String ;
FUNCTION  QUELDATEEXO(DateC : TDateTime ; Var Date1,Date2 : TDateTime) : String ;
PROCEDURE RECALCULPRORATAECHE ( Var ModR : T_ModeRegl ; MtEcrP,MtEcrD : Double ) ;
//Function  TrouveIndice ( T : TStrings ; Nom : String ; Parle : boolean ) : integer ;
Function  TrouveIndice (Nom : String ; Parle : boolean ) : integer ;
Function  QuelSens ( NatPiece,NatGene : String3 ; Sens : byte ; CommeBQE : boolean = False ) : byte ;
Function  QuelAuto ( StAuto : String ; Cpte : String ) : integer ;
Function  EstInterdit( StInter : String ; Cpte : String ; Sens : byte ) : byte ;
Function  EstCptEcart ( Cpt : String17 ) : boolean ;
Function  EstOuvreBil( StCpte : String ) : boolean ;
Function  EstChaPro( NatGene : String ) : boolean ;
Function  EstCollectif( Cpte : String ) : boolean ;
Function  EstCptBQE( Cpte : String ) : Boolean ;
Function  EstCptSyn( Cpte, Jnl : string; NonEcc : Boolean; NomBase : string = '') : Boolean ;
Function  TrouveAuto ( StAuto : String ; indice : integer ) : String ;
function  ControleDate ( StD : String ) : integer ;
function  EnDevise ( Code : String3 ) : boolean ;
procedure AffecteGrid (GS : THGrid ; Action : TActionFiche ) ;
function  QuelleLigne ( Var St : hString ) : integer ;
Function  GetDeviseBanque ( Cpte : String ) : String ;
Function  PasDeviseBanque ( Cpte,CodeD : String ) : boolean ;
Function  PasOkDevBqe ( Devi,CodeD : String ) : boolean ;
function  PasCreerDate ( DD : TDateTime ) : boolean ;
function  PasCreerDateGC ( DD : TDateTime ) : boolean ;
function  ModeRevisionActive ( DD : TDateTime ) : boolean ;
function  RevisionActive ( DD : TDateTime ) : boolean ;
function  NbJoursOk ( DatePiece,DateEche : TDateTime ) : boolean ;
function  GetChampsTiers ( ct : TChampsTiers ) : String ;
Procedure GridEna(G : THGrid ; Ena : Boolean ) ;
function  GetNewNumPiece ( Souche : String3 ; Ecrit : boolean ; DD : TDateTime) : Longint ;
Procedure SetCotationDB ( DD : TDateTime ; DS : TDataSet) ;
function  DepasseLimiteDemo : boolean ;
function  QuellePeriode ( UneDate,DebExo : TDateTime ) : byte ;
Function  ADevalider(UnJal : String ; UneDate : TDateTime) : Boolean ;
procedure CVentilerAttente ( vTOBEcr : TOB ; vInAxe : integer ; vStSection : string = '' ) ;
function  CVentilerVentil ( vTOBEcr : TOB ; vInAxe : integer ; vNbDec : integer ; vStV_NATURE : string ; vStCodeVT : string ; vDossier : string ; vBoComplet : boolean ; vRestriction : TObject  ) : boolean ;
procedure CVentilerTOBTva( vTOBEcr : TOB ) ;
Function  VentilerTOB ( TOBEcr : TOB ; CodeG : String ; NumLig,NbDec : integer ; QueAttente : boolean ; vDossier : String = ''; vBoComplet : Boolean = True ; vStVT : String = '' ) : boolean ;
//SG6 28.01.05 Gestion mode croisaxe
Function VentilerTOBCroisaxe ( TOBEcr : TOB ; CodeG : String ; NumLig,NbDec : integer ; QueAttente : boolean; vBoComplet : Boolean = True) : boolean ;  {FP 12/06/2006 FQ18226 Ajout paramètre vBoComplet}

Function  VentilerTOBCreat ( TOBEcr : TOB ; NbDec : integer ; StAxe,Sect : String ; VerifGene : Boolean) : boolean ;
Procedure ProraterVentilsTOB ( TOBEcr : TOB ; NbDec : integer ) ;
Function  StrS( X : Double ; DD : integer ) : String ;
Function  StrF00 ( X : Double ; DD : integer ) : String ;
Function  StrSP( X : Double ) : String ;
Function  StrS0 ( X : Double ) : String ;
Function  StrP0 ( X : Double ) : String ;
Function  CaseFicDataType ( Fic : String ) : TFichierBase ;
Function  GetRIBPrincipal ( Cpt : String ) : String ;
Function GetRIBParticulier ( Cpt : String ; No : integer) : String ;
{$IFNDEF NOVH}
Procedure SetMPACCDB ( DS : TDataSet) ;
Function  MPTOACC ( MP : String ) : String ;
Function  MPTOCATEGORIE ( MP : String ) : String ;
{$ENDIF}
Function  EstTVATPF ( Compte : String17 ; TVA : Boolean ) : boolean ;
function  QuelTypeMvt ( Compte : String17 ; NC,NP : String3 ) : String ;

Function  isEncMP ( smp : TSuiviMP ) : Boolean ;
Procedure EcrVersAna ( TOBEcr,TOBAna : TOB ) ;
Procedure ArrondirAnaTOB ( TOBEcr : TOB ; NbDec : integer; vInAxe : integer = 0 ) ;
Procedure VentilLigneTOB ( TOBAna : TOB ; Section : String ; NumVentil,NbDec : integer ; Pourc : double ; Debit : boolean ; PourQte1 : double = 0.0 ; PourQte2 : double = 0.0 ; splan1 : string = '' ; splan2 : string = ''; splan3 : string = ''; splan4 : string = '' ; splan5 : string = '');

Function  GetNumLotTraChq ( STra : String ) : String ;
Function  IncNumLotTraChq ( STra : String ) : String ;
Function  NormaliserPieceSimu ( TOBL : TOB ; MajTOBL : Boolean ) : Boolean ;
Procedure SauveJalEffet(St,Jal : String) ;
Function  ChargeJalEffet(St : String) : String ;

//LG* 21/02/2002
procedure EgaliseOBM ( var OS, OD : TOBM ) ;
{$IFNDEF EAGLSERVER}
Function  MajIOCom(GS : THGrid) : Boolean ;
{$ENDIF}
//SB* 29/10/2002
PROCEDURE CALC(Var Deb,Cred,Tot,TauxC : Double ; Sens,Decim : Byte) ;
PROCEDURE RECALC(Var Deb,Cred,Tot,MtEcr : Double ; Sens : Byte) ;

function  TSorteLettreToStr(TSL : TSorteLettre) : String;
function  StrToTSorteLettre(psz : String) : TSorteLettre;

// BPY le 19/05/2003 : Getion des extourne
function  MemePiece(ListeEcriture : TOB ; mode : string ; iref,i : integer ; ForceBordereau : boolean = false) : boolean;
function  GenereLesExtournes(ListeEcriture,ListePiece : TOB ; bSelectionPiece : boolean ; All : boolean ; date : TDateTime ; tipe,negatif,libelle,refint,refext,mode : string; WithInsert : boolean=true) : TOB;
function ExtourneEcriture(ListeEcriture : TOB; WithEcran : boolean=True; LaDate : TdateTime=0; Qualif : string=''; Negatif : boolean=false; Lib : string=''; RefI : string=''; RefE : string=''; WithInsert : boolean=true) : TOB;
function TypeCptTVADefaut (CompteG, NatureG:string):string; //Lek FQ;001;26386
function ISmoisCloture (UneDate : TdateTime) : boolean;
function isMoisJournalCloture (Journal: string; UneDate : TdateTime) : Boolean;

Const CP_Jal  = 0 ; CP_Date = 1 ; CP_NumL = 2 ;
      CP_NB = 3   ; CP_RUBD = 4 ; CP_MONTANT = 5 ;
      CP_RUBC = 6 ; CP_REF = 7  ; CP_LIB = 8 ;

      ADecimP     : Integer = 4 ;

Const Height_Ecr  = 203 ;
      Height_Tres = 270 ;
      Height_ODA  = 203 ;

Type TMod = Class
            ModR : T_ModeRegl ;
            end ;

implementation

Uses SaisComm ,
 {$IFNDEF PGIIMMO}
  uLibEcriture,
  ULibAnalytique ,
 {$ENDIF}
 {$IFDEF EAGLSERVER}
 ULibCpContexte ,
 {$ENDIF}

 {$IFDEF MODENT1}
 CPProcGen,
 CPProcMetier, 
 {$ENDIF MODENT1}
//XMG 20/04/04 fin
  ParamSoc,UFonctionsCBP,
  CbpMCD,
  CbpEnumerator
  ;

(*======================================================================*)
Procedure GridEna(G : THGrid ; Ena : Boolean ) ;
BEGIN
G.SynEnabled:=Ena ;
END ;

(*======================================================================*)
procedure AffecteGrid (GS : THGrid ; Action : TActionFiche ) ;
Var O : Boolean ;
BEGIN

  O:=GS.Enabled ; GS.Enabled:=True ;

  case  Action of taConsult :
        BEGIN
          GS.Options := GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
          if GS.TypeSais <> tsBudget then GS.Options := GS.Options+[GoRowSelect] ;
               END ;
          else BEGIN
          if GS.TypeSais <> tsBudget then GS.Options := GS.Options-[GoRowSelect] ;
          GS.Options := GS.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
               END ;
   END ;

  GS.Enabled:=O ;

END ;


(*======================================================================*)
Procedure ChangeFormatDevise ( FF : TForm ; Decim : integer ; Symbole : String ) ;
Var i : integer ;
    T   : TComponent ;
    C   : THNumEdit ;
BEGIN
for i:=0 to FF.ComponentCount-1 do
    BEGIN
    T:=FF.Components[i] ;
    if T is THNumEdit then
       BEGIN
       C:=THNumEdit(T) ;
       Case T.Tag of
       1,3 : ChangeMask(C,Decim,Symbole) ; { Devise Ou pivot }
         2 : ChangeMask(C,V_PGI.OkDecV,V_PGI.SymbolePivot) ; { Pivot exclusivement }
         4 : ChangeMask(C,V_PGI.OkDecE,'') ; { Euro exclusivement }
         5 : ChangeMask(C,V_PGI.OkDecQ,'') ; { Qtés }
         End ;
       END ;
    END ;
END ;

(*======================================================================*)
Procedure ChangeFormatPivot ( FF : TForm ) ;
Var i : integer ;
    T   : TComponent ;
    C   : THNumEdit ;
BEGIN
for i:=0 to FF.ComponentCount-1 do
    BEGIN
    T:=FF.Components[i] ;
    if T is THNumEdit then BEGIN C:=THNumEdit(T) ; ChangeMask(C,V_PGI.OkDecV,V_PGI.SymbolePivot) ; END ;
    END ;
END ;

(*======================================================================*)
Function CaseNatJal ( NatJal : String3 ) : TZoomTable ;
BEGIN
if NatJal='ACH' then Result:=tzJAchat else
 if NatJal='VTE' then Result:=tzJVente else
  if NatJal='ECC' then Result:=tzJEcartChange else
   if ((NatJal='BQE') or (NatJal='CAI')) then Result:=tzJBanque else Result:=tzJOD ;
END ;                                               

(*======================================================================*)
Function OkSoucheN1(DD : TDateTime) : Boolean ;
BEGIN
Result:=(GetMultisouche) And (QuelExoDT(DD)=GetSuivant.Code) ;
END ;

(*======================================================================*)
FUNCTION GETNUM (TypeEcr : TTypeEcr ; Facturier : String3 ; Var MasqueNum : String17 ; DD : TDateTime ; vDossier : String = '' ) : LongInt ;
Var Q : TQuery ;
    SQL : String ;
begin
Result:=0 ; MasqueNum:='' ;
Case TypeEcr Of
  EcrGen,EcrAna,EcrClo :
       BEGIN
       Result:=CPIncrementerCompteur('CPT',Facturier,DD,MasqueNum,0);
       (*
       SQL:='SELECT SH_TYPE, SH_SOUCHE, SH_NUMDEPART, SH_SIMULATION, SH_MASQUENUM,SH_NUMDEPARTS, SH_SOUCHEEXO ';
       SQL:=SQL+'FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'"' ;
       Q:=OpenSelect(SQL,vDossier) ;
       if Not Q.Eof then
          BEGIN
          Result:=Q.FindField('SH_NUMDEPART').AsInteger ;
          MasqueNum:=Q.FindField('SH_MASQUENUM').AsString ;
          If OkSoucheN1(DD) And (Q.FindField('SH_SOUCHEEXO').AsString='X') Then Result:=Q.FindField('SH_NUMDEPARTS').AsInteger ;
          END ;
       Ferme(Q) ;
       *)
       END ;
  EcrBud :
       BEGIN
       Result:=CPIncrementerCompteur('BUD',Facturier,DD,MasqueNum,0);
(**       SQL:='SELECT SH_TYPE, SH_SOUCHE, SH_NUMDEPART, SH_SIMULATION, SH_MASQUENUM ';
       SQL:=SQL+'FROM SOUCHE WHERE SH_TYPE="BUD" AND SH_SOUCHE="'+Facturier+'"' ;
       Q:=OpenSelect(SQL,vDossier) ;
       if Not Q.Eof then
          BEGIN
          Result:=Q.FindField('SH_NUMDEPART').AsInteger ;
          MasqueNum:=Q.FindField('SH_MASQUENUM').AsString ;
          END ;
       Ferme(Q) ;
**)
       END ;
  end ;
end ;

(*======================================================================*)
PROCEDURE SETNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; Num,OldN : Longint ; DD : TDateTime ; vDossier : String = '' ) ;
Var Q : TQuery ;
    OkN1 : Boolean ;
begin
if Num<=0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ; OkN1:=FALSE ;
Case TypeEcr of
  EcrGen,EcrAna,EcrClo :
    BEGIN
    If OkSoucheN1(DD) Then
      BEGIN
      Q:=OpenSelect('SELECT SH_SOUCHEEXO FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" ',vDossier) ;
      If Not Q.Eof Then If Q.Fields[0].AsString='X' Then OkN1:=TRUE ; Ferme(Q) ;
      If OkN1 Then
        BEGIN
        If ExecuteSQL('UPDATE ' + GetTableDossier(vDossier, 'SOUCHE') + ' SET SH_NUMDEPARTS='+IntToStr(Num)+' WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPARTS='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;
        END
      Else
        BEGIN
        If ExecuteSQL('UPDATE ' + GetTableDossier(vDossier, 'SOUCHE') + ' SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPART='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;
        END ;
      END
    Else if ExecuteSQL('UPDATE ' + GetTableDossier(vDossier, 'SOUCHE') + ' SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPART='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;
    END ;

  EcrBud : if ExecuteSQL('UPDATE ' + GetTableDossier(vDossier, 'SOUCHE') + ' SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_TYPE="BUD" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPART='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;

  end ;
end ;

(*======================================================================*)
PROCEDURE SETINCNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; Var Num : LongInt ; DD : TDateTime ; vDossier : String = '' ) ;
Var MM   : String17 ;
//    OldN : Longint ;
    TypeSouche : string;
begin
  if TypeEcr=EcrBud then
    TypeSouche:='BUD'
  else TypeSouche:='CPT';
  Num:=CPIncrementerCompteur(TypeSouche,Facturier,DD,MM,1);
  if Num>0 then Dec(Num)
  else V_PGI.IoError:=oeUnknown ;
(**
Num:=IncrementerCompteur(
Num:=GetNum(TypeEcr,Facturier,MM,DD,vDossier) ;
OldN:=Num ; Inc(Num) ; SetNum(TypeEcr,Facturier,Num,OldN,DD,vDossier) ; Dec(Num) ;
**)
end ;

(*======================================================================*)
PROCEDURE SETDECNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; DD : TDateTime ; vDossier : String = '' ) ;
Var MM   : String17 ;
    Num,OldN  : Longint ;
begin
Num:=GetNum(TypeEcr,Facturier,MM,DD,vDossier) ;
OldN:=Num ; Dec(Num) ;  SetNum(TypeEcr,Facturier,Num,OldN,DD,vDossier) ;
end ;

(*======================================================================*)
PROCEDURE GETINFOSDEVISE ( Var DEV : RDEVISE ; vDossier : String = '') ;
Var Q : TQuery ;
    SQL : String ;
begin
DEV.Decimale:=V_PGI.OkDecV ; DEV.Symbole:=V_PGI.SymbolePivot ; DEV.Quotite:=1 ;
{$IFNDEF EAGLSERVER}
DEV.MaxDebit:=VH^.MaxDebitPivot ; DEV.MaxCredit:=VH^.MaxCreditPivot ;
DEV.CptDebit:=VH^.CptLettrDebit ; DEV.CptCredit:=VH^.CptLettrCredit ;
DEV.Libelle:=VH^.LibDevisePivot ;
{$ENDIF}
if ((DEV.Code='') or (DEV.Code=V_PGI.DevisePivot)) then Exit ;
SQL:='Select D_DEVISE, D_LIBELLE, D_SYMBOLE, D_DECIMALE, D_QUOTITE, D_MAXDEBIT, D_MAXCREDIT, '
    +'D_CPTLETTRDEBIT, D_CPTLETTRCREDIT FROM DEVISE WHERE D_DEVISE="'+DEV.Code+'"' ;
Q:=OpenSelect(SQL,vDossier) ;
if Not Q.Eof then
   BEGIN
   DEV.Libelle:=Q.FindField('D_LIBELLE').AsString ;
   DEV.Decimale:=Q.FindField('D_DECIMALE').AsInteger ;
   DEV.Symbole:=Q.FindField('D_SYMBOLE').AsString ; DEV.Quotite:=Q.FindField('D_QUOTITE').AsFloat ;
   DEV.MaxDebit:=Q.FindField('D_MAXDEBIT').AsFloat ; DEV.MaxCredit:=Q.FindField('D_MAXCREDIT').AsFloat ;
   DEV.CptDebit:=Q.FindField('D_CPTLETTRDEBIT').AsString ; DEV.CptCredit:=Q.FindField('D_CPTLETTRCREDIT').AsString ;
   END ;
Ferme(Q) ;
end ;

(*======================================================================*)
FUNCTION GETTAUX ( Code : String3 ; Var DateTaux : TDateTime ; DateC : TDateTime ; vDossier : String = ''; WhereDateGC : string='') : Double ;
Var Q : TQuery ;
    SQL : String ;
    LeTaux,LaCote : double ;
begin
  {Ne renvoie plus un taux mais un coef équivalent à un ratio entre devise et le Franc}
  Result := 1 ;
  DateTaux := DateC ;
  if ((Code = V_PGI.DevisePivot) or (Code = V_PGI.DeviseFongible)) then Exit ;
  if DateC < V_PGI.DateDebutEuro then
  begin
    {Si date < date début --> nécessairement par rapport à ancienne chancellerie}
    SQL := 'SELECT H_DATECOURS, H_TAUXREEL FROM CHANCELL WHERE H_DEVISE = "' + Code + '"';
    if WhereDateGC = '' then
      Sql := Sql + ' AND H_DATECOURS <= "' + UsDateTime(DateC) + '" AND H_DATECOURS < "' + UsDateTime(V_PGI.DateDebutEuro) + '"'
      else
      Sql := Sql + WhereDateGC;
    SQL := SQL + ' ORDER BY H_DATECOURS DESC' ;
    Q := OpenSelect(SQL, vDossier) ;
    if Not Q.EOF then
    begin
      Result := Q.Fields[1].AsFloat ;
      DateTaux := Q.Fields[0].AsDateTime ;
    end;
    Ferme(Q) ;
  end else
  begin
    {Monnaie Out pour date > début euro --> chancellerie par rapport à Euro}
    SQL := 'SELECT H_DATECOURS, H_TAUXREEL, H_COTATION FROM CHANCELL WHERE H_DEVISE = "' + Code + '"';
    if WhereDateGC = '' then
      Sql := Sql + ' AND H_DATECOURS <= "' + UsDateTime(DateC) + '" AND H_DATECOURS >= "' + UsDateTime(V_PGI.DateDebutEuro) + '"'
      else
      Sql := Sql + WhereDateGC;
    SQL := SQL + ' ORDER BY H_DATECOURS DESC' ;
    Q:=OpenSelect(SQL, vDossier) ;
    if Not Q.EOF then
    begin
      DateTaux := Q.Fields[0].AsDateTime ;
      LeTaux := Arrondi(Q.Fields[1].AsFloat,9) ;
      LaCote := Arrondi(Q.Fields[2].AsFloat,9) ;
      if ((LeTaux <> 0) and (LaCote <> 0)) then
        Result := LeTaux * V_PGI.TauxEuro
      else if ((LeTaux = 0) and (LaCote <> 0)) then
        Result := V_PGI.TauxEuro / LaCote
      else if ((LeTaux <> 0) and (LaCote = 0))
        then Result := LeTaux * V_PGI.TauxEuro
      else
      Result := 1 ;
    end;
    Ferme(Q) ;
  end;
end;

(*======================================================================*)
FUNCTION VALEURPIVOT(St : String ; Taux : Double ; Quotite : Double) : Double ;
Var X : Double ;
begin
if Quotite=0 then Quotite:=1 ;
X:=Valeur(St) ; Result:=X*Taux/Quotite ;
end ;

(*======================================================================*)
(*======================================================================*)
FUNCTION QUELEXO(DateC : String) : String ;
Var DD : TDateTime ;
begin
dd:=StrToDate(DateC) ; Result:=QuelExoDT(dd) ;
end ;

(*======================================================================*)
FUNCTION QUELEXODT(DD : TDateTime) : String ;
{$IFNDEF NOVH}
Var i : Integer ;
{$ENDIF}
{ DONE -olaurent -cdll serveur :  }
begin
{$IFNDEF NOVH}
Result:=VH^.EnCours.Code ;
If (dd>=VH^.EnCours.Deb) and (dd<=VH^.EnCours.Fin) then Result:=VH^.EnCours.Code else
   If (dd>=VH^.Suivant.Deb) and (dd<=VH^.Suivant.Fin) then Result:=VH^.Suivant.Code Else
      If (dd>=VH^.Precedent.Deb) and (dd<=VH^.Precedent.Fin) then Result:=VH^.Precedent.Code Else
      BEGIN
         For i:=1 To 5 Do
           BEGIN
           If (dd>=VH^.ExoClo[i].Deb) And (dd<=VH^.ExoClo[i].Fin)then BEGIN Result:=VH^.ExoClo[i].Code ; Exit ; END ;
           END ;
      END ;
{$ELSE}
{$IFDEF EAGLSERVER}
 result := TCPContexte.GetCurrent.Exercice.QUELEXODT(DD) ;
{$ENDIF}
{$ENDIF}
end ;

(*======================================================================*)
FUNCTION QUELDATEEXO(DateC : TDateTime ; Var Date1,Date2 : TDateTime) : String ;
{$IFNDEF NOVH}
Var DD : TDateTime ;
    i : Integer ;
{$ENDIF}
begin
{ DONE -olaurent -cdll serveur :  }
{$IFNDEF NOVH}
result:=GetSuivant.Code ; dd:=DateC ; Date1:=DD ; Date2:=DD ;
If (dd>=GetSuivant.Deb) and (dd<=GetSuivant.Fin) then
   BEGIN Result:=GetSuivant.Code ; Date1:=GetSuivant.Deb ; Date2:=GetSuivant.Fin ; END else
   If (dd>=GetSuivant.Deb) and (dd<=GetSuivant.Fin) then
      BEGIN Result:=GetSuivant.Code ; Date1:=GetSuivant.Deb ; Date2:=GetSuivant.Fin ; END Else
      If (dd>=GetPrecedent.Deb) and (dd<=GetPrecedent.Fin) then
         BEGIN Result:=GetPrecedent.Code ; Date1:=GetPrecedent.Deb ; Date2:=GetPrecedent.Fin ; END Else
            BEGIN
            For i:=1 To 5 Do
              BEGIN
              If (dd>=VH^.ExoClo[i].Deb) And (dd<=VH^.ExoClo[i].Fin)then
                 BEGIN
                 Date1:=VH^.ExoClo[i].Deb ; Date2:=VH^.ExoClo[i].Fin ;
                 Result:=VH^.ExoClo[i].Code ; Exit ;
                 END ;
              END ;
            END ;
{$ELSE}
 {$IFDEF EAGLSERVER}
 result := TCPContexte.GetCurrent.Exercice.QUELDATEEXO(DateC,Date1,Date2) ;
 {$ENDIF}
{$ENDIF}
end ;

(*======================================================================*)
PROCEDURE RECALCULPRORATAECHE ( Var ModR : T_ModeRegl ; MtEcrP,MtEcrD : Double ) ;
Var MtP,MtD{,MtE},Taux,TotP,TotD : Double ;
    i,k : Integer ;
    lBoOuiTvaEnc : boolean ;
BEGIN
MtP:=ModR.TotalAPayerP ; MtD:=ModR.TotalAPayerD ; //MtE:=ModR.TotalAPayerE ;
If MtP=0 then Exit ; Taux:=MtEcrP/MtP ;
if EnDevise(ModR.CodeDevise) Then BEGIN If MtD=0 then Exit ; Taux:=MtEcrD/MtD ; END ;
TotP:=0 ; TotD:=0 ;
With ModR do
  BEGIN
  For i:=1 to NbEche do
      BEGIN
      TabEche[i].MontantP:=Arrondi(Taux*TabEche[i].MontantP,MODR.Decimale) ;
      {#TVAENC}
      {$IFDEF ESP}
      lBoOuiTvaEnc:=GetParamSocSecur('SO_OUITVAENC', False) or (VH^.PaysISO=CodeISOES) ; //XMG 31/07/03
      {$ELSE}
      lBoOuiTvaEnc:=GetParamSocSecur('SO_OUITVAENC', False) ;
      {$ENDIF}
      if ((lBoOuiTvaEnc) and (ModifTva) and (Taux<>1)) then
         BEGIN
         for k:=1 to 5 do TabEche[i].TAV[k]:=Arrondi(Taux*TabEche[i].TAV[k],V_PGI.OkDecV) ;
         END ;
      END ;
  For i:=1 to NbEche do
      TotP:=TotP+TabEche[i].MontantP ;
  if TotP<>MtEcrP then TabEche[NbEche].MontantP:=TabEche[NbEche].MontantP+(MtEcrP-TotP) ;
  if EnDevise(ModR.CodeDevise) Then
     begin
     For i:=1 to NbEche do TabEche[i].MontantD:=Arrondi(Taux*TabEche[i].MontantD,MODR.Decimale) ;
     For i:=1 to NbEche do TotD:=TotD+TabEche[i].MontantD ;
     if TotD<>MtEcrD then TabEche[NbEche].MontantD:=TabEche[NbEche].MontantD+(MtEcrD-TotD) ;
     END else
     BEGIN
     For i:=1 to NbEche do TabEche[i].MontantD:=TabEche[i].MontantP ;
     END ;
  END ;
ModR.TotalAPayerP:=MtEcrP  ; ModR.TotalAPayerD:=MtEcrD ; 
END ;

{$IFNDEF SANSCOMPTA}
(*======================================================================*)
Destructor TObjAna.Destroy ;
Var i : Integer ;
begin
for i:=1 to MaxAxe do if AA[i]<>Nil then
    BEGIN
    if AA[i].L<>Nil then BEGIN VideListe(AA[i].L) ; AA[i].L.Free ; AA[i].L:=Nil ; END ;
    if AA[i].M<>Nil then BEGIN VideListe(AA[i].M) ; AA[i].M.Free ; AA[i].M:=Nil ; END ;
    AA[i].Free ;
    AA[i]:=Nil ;
    END ;
Inherited destroy ;
end ;

(*======================================================================*)
PROCEDURE CALCULTOTALANAL ( Pf : String ; Var MtP,MtD,TauxAN : Double ; T : TList ; CodeDev : String3 ; Sens : Byte ; Lig : Integer) ;
Var i : Integer  ;
    P : P_TV ;
BEGIN
Mtp:=0 ; MtD:=0 ; TauxAn:=0 ;
for i:=0 to T.Count-1 do if i<>Lig then
    BEGIN
    P:=P_TV(T.Items[i]) ;
    TauxAn:=TauxAn+GetMvtA(P.F,Pf+'_POURCENTAGE') ;
    If Sens=1 then
       BEGIN
       MtP:=MtP+GetMvtA(P.F,Pf+'_DEBIT')  ;
       MtD:=MtD+GetMvtA(P.F,Pf+'_DEBITDEV') ;
       END else
       BEGIN
       MtP:=MtP+GetMvtA(P.F,Pf+'_CREDIT') ;
       MtD:=MtD+GetMvtA(P.F,Pf+'_CREDITDEV') ;
       END ;
    END ;
END ;


(*======================================================================*)
PROCEDURE PRORATE1ANAL ( Pf : String ; Var MtP,MtD,MtEcrP,MtEcrD,TauxAN : Double ; AAi : TAA ; CodeDev : String3 ; Sens,Decim : Byte ) ;
Var LeTaux : Double ;
    TotP,TotD : Double ;
    FDebit,FCredit,FDebitDev,FCreditDev : Double ;
    P : P_TV ;
    i : Integer ;
    T : TList ;
BEGIN
If MtP=0 then Exit ;
T:=AAi.L ; If T=NIL Then Exit ;
LeTaux:=MtEcrP/MtP ;
if EnDevise(CodeDev) Then BEGIN If MtD=0 then Exit ; LeTaux:=MtEcrD/MtD ; END ;
TotP:=0 ; TotD:=0 ;
for i:=0 to T.Count-1 do
    BEGIN
    P:=P_TV(T.Items[i]) ;
    FDebit:=GetMvtA(P.F,Pf+'_DEBIT') ; FCredit:=GetMvtA(P.F,Pf+'_CREDIT') ;
    Calc(FDebit,FCredit,TotP,LeTaux,Sens,V_PGI.OkDecV) ;
    PutMvtA(AAi,P.F,Pf+'_DEBIT',FDebit) ; PutMvtA(AAi,P.F,Pf+'_CREDIT',FCredit) ;
    FDebitDev:=FDebit ; FCreditDev:=FCredit ;
    {Euro}
    {Devise}
    if EnDevise(CodeDev) then
       BEGIN
       FDebitDev:=GetMvtA(P.F,Pf+'_DEBITDEV') ; FCreditDev:=GetMvtA(P.F,Pf+'_CREDITDEV') ;
       Calc(FDebitDev,FCreditDev,TotD,LeTaux,Sens,Decim) ;
       END ;
    PutMvtA(AAi,P.F,Pf+'_DEBITDEV',FDebitDev) ; PutMvtA(AAi,P.F,Pf+'_CREDITDEV',FCreditDev) ;
    PutMvtA(AAi,P.F,Pf+'_TOTALECRITURE',MtEcrP) ;
    PutMvtA(AAi,P.F,Pf+'_TOTALDEVISE',MtEcrD) ;
    END ;
If T.Count>0 Then
   BEGIN
   i:=T.Count-1 ; P:=P_TV(T.Items[i]) ;
   if TotP<>MtEcrP then
      BEGIN
      FDebit:=GetMvtA(P.F,Pf+'_DEBIT') ; FCredit:=GetMvtA(P.F,Pf+'_CREDIT') ;
      ReCalc(FDebit,FCredit,TotP,MtEcrp,Sens) ;
      PutMvtA(AAi,P.F,Pf+'_DEBIT',FDebit) ; PutMvtA(AAi,P.F,Pf+'_CREDIT',FCredit) ;
      if Not EnDevise(CodeDev) then
         BEGIN
         FDebitDev:=FDebit ; FCreditDev:=FCredit ;
         END else
         BEGIN
         FDebitDev:=GetMvtA(P.F,Pf+'_DEBITDEV') ; FCreditDev:=GetMvtA(P.F,Pf+'_CREDITDEV') ;
         END ;
      END ;
   if ((EnDevise(CodeDev)) And (TotD<>MtEcrD)) then
      BEGIN
      FDebitDev:=GetMvtA(P.F,Pf+'_DEBITDEV') ; FCreditDev:=GetMvtA(P.F,Pf+'_CREDITDEV') ;
      ReCalc(FDebitDev,FCreditDev,TotD,MtEcrD,Sens) ;
      END ;
   PutMvtA(AAi,P.F,Pf+'_DEBITDEV',FDebitDev) ; PutMvtA(AAi,P.F,Pf+'_CREDITDEV',FCreditDev) ;
   END ;
END ;

(*======================================================================*)
Procedure InverseSensAna ( AAi : TAA ) ;
Var P : P_TV ;
    D,C : Double ;
    i   : integer ;
BEGIN
for i:=0 to AAi.L.Count-1 do
    BEGIN
    P:=P_TV(AAi.L[i]) ;
    D:=GetMvtA(P.F,'Y_DEBIT') ; C:=GetMvtA(P.F,'Y_CREDIT') ;
    PutMvtA(AAi,P.F,'Y_DEBIT',C) ; PutMvtA(AAi,P.F,'Y_CREDIT',D) ;
    D:=GetMvtA(P.F,'Y_DEBITDEV') ; C:=GetMvtA(P.F,'Y_CREDITDEV') ;
    PutMvtA(AAi,P.F,'Y_DEBITDEV',C) ; PutMvtA(AAi,P.F,'Y_CREDITDEV',D) ;
    END ;
END ;

PROCEDURE RECALCULPRORATAANAL ( Pf : String ; OBA : TObjAna ; MtEcrP,MtEcrD : Double ; NumA : integer ; SensEcr : integer ) ;
Var MtP,MtD,TauxAn : Double ;
    j : Integer ;
begin
for j:=1 To MaxAxe Do if ((NumA=0) or (j=NumA)) then
  begin
  If OBA.AA[j]<>Nil Then
     begin
     CalculTotalAnal(Pf,MtP,MtD,TauxAN,OBA.AA[j].L,OBA.Devise,OBA.Sens,-1) ;
     Prorate1Anal(Pf,MtP,MtD,MtEcrP,MtEcrD,TauxAN,OBA.AA[j],OBA.Devise,OBA.Sens,OBA.Decimale) ;
     OBA.TotalEcriture:=MtEcrP ; OBA.TotalDevise:=MtEcrD ;
     if SensEcr<>OBA.Sens then BEGIN InverseSensAna(OBA.AA[j]) ; OBA.Sens:=SensEcr ; END ;
     end ;
  end ;
end ;

(*======================================================================*)
PROCEDURE CALCULTOTALA ( Pf : String ; OBA : TObjAna ; Var TotMP,TotMD,TotME,TotTaux : Double ; j,Lig : Integer) ;
begin
TotMP:=0 ; TotMD:=0 ; TotME:=0 ; TotTaux:=0 ;
If OBA.AA[j]<>Nil Then CalculTotalAnal(Pf,TotMP,TotMD,TotTaux,OBA.AA[j].L,OBA.Devise,OBA.Sens,Lig) ;
end ;


{=====================================================================}
Function  GetMvtA ( F : T_V ; Quoi : String ) : Variant ;
Var i : integer ;
BEGIN
Result:=0 ;
if upcase(Quoi[1])='Y' then i:=TrouveIndice(Quoi,True) else Exit ;
try
 if i>=0 then Result:=F[i].V ;
except
 on E:Exception do messageAlerte('Erreur lors de la récupération de l''analytique sur le champs '+Quoi+#10#13+E.Message)
end;
END ;

{=====================================================================}
Procedure PutMvtA ( AAi : TAA ; Var F : T_V ; Quoi : String ; Value : Variant) ;
Var i : integer ;
BEGIN
if upcase(Quoi[1])='Y' then i:=TrouveIndice(Quoi,True) else Exit ;
try
if i>=0 then
  BEGIN
  If (VarType(F[i].V)<>VarEmpty) and (F[i].V<>Value) then BEGIN AAI.Etat:=Modifie ; F[i].Etat:=AAi.Etat ; END ;
  F[i].V:=Value ;
  END ;
except
 on E:Exception do messageAlerte('Erreur lors de l''affectation du champs '+Quoi+#10#13+E.Message)
end;
END ;

{$ENDIF}

{=====================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/03/2002
Modifié le ... :   /  /    
Description .. : LG* modif suite à la suppression des tstringlist DescriEcr, 
Suite ........ : DescriAna, DescriGui, 
Mots clefs ... : 
*****************************************************************}
Function TrouveIndice (Nom : String ; Parle : boolean ) : integer ;
BEGIN
result:=ChampToNum(Nom) ;
if ((Result=-1) and (Parle)) then MessageAlerte('pas trouvé '+Nom) ;
END ;

{$IFNDEF SANSCOMPTA}
{=====================================================================}
  {$IFNDEF EAGLSERVER}
Function InitVariant ( Nom : String ) : Variant ;
Var Mcd : IMCDServiceCOM;
		Tipe : string;
BEGIN
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
	Result:='' ;
	Tipe := (Mcd.GetField(Nom) as IFieldCom).Tipe;
  if ((Tipe='INTEGER') or (Tipe='DOUBLE') or (Tipe='RATE')) then Result:=0 else
  if Tipe='DATE' then Result:=iDate1900 else
  if Tipe='BOOLEAN' then Result:='-' ;
END ;
  {$ENDIF}
{$ENDIF}

{$IFNDEF SANSCOMPTA}
{=====================================================================}
Constructor TAA.Create ;
begin
Inherited Create ;
L:=TList.Create ; M:=TList.Create ; Etat:=Inchange ;
{Status:=EnMemoire ;}
end ;

Destructor TAA.Destroy ;
begin
VideListe(L) ; L.Free ; VideListe(M) ; M.Free ;
Inherited Destroy ;
end ;
{$ENDIF}


Function  QuelSens ( NatPiece,NatGene : String3 ; Sens : byte ; CommeBQE : boolean = False ) : byte ;
Var OkD : Boolean ;
    NatP : String ;
BEGIN
OkD:=False ; QuelSens:=Sens ; NatP:=NatPiece ;
if CommeBQE then
   BEGIN
   if ((NatGene='COC') or (NatGene='TID')) then NatP:='RC' else
    if ((NatGene='COF') or (NatGene='TIC')) then NatP:='RF' ;
   END ;
if ((NatP='AC') or (NatP='AF')) then OkD:=True ;
if ((NatP='RC') or (NatP='RF') or (NatP='OC') or (NatP='OF')) then
   BEGIN
   // FQ16898 Crédit / débit inversé sur Journaux Caisse et Banque
   if ((NatGene<>'CHA') and (NatGene<>'PRO') and
       (NatGene<>'BQE') and (NatGene<>'CAI')) then OkD:=True ;
   END ;
if ((OkD) and (Sens in [1,2])) then QuelSens:=3-Sens ;
END ;

Function QuelAuto ( StAuto : String ; Cpte : String ) : integer ;
Var StA,StC : String ;
    Find : boolean ;
    i    : integer ;
BEGIN
Result:=0 ; StA:=StAuto ; Find:=False ; i:=0 ;
Repeat
 StC:=ReadTokenSt(StA) ;
 if StC<>'' then BEGIN Inc(i) ; if Trim(StC)=Cpte then Find:=True ; END ;
Until ((StC='') or (Find)) ;
if Find then Result:=i ;
END ;

Function  TrouveAuto ( StAuto : String ; indice : integer ) : String ;
Var i : integer ;
    StC,StA : String ;
BEGIN
StA:=StAuto ; StC:='' ;
for i:=1 to Indice do StC:=ReadTokenSt(StA) ;
Result:=StC ;
END ;

Function  EstCollectif( Cpte : String ) : boolean ;
Var Q : TQuery ;
BEGIN
Result:=False ; if Cpte='' then Exit ;
Q:=OpenSQL('Select G_GENERAL from GENERAUX where G_GENERAL="'+Cpte+'" AND G_COLLECTIF="X"',True,-1,'',true) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function EstOuvreBil( StCpte : String ) : boolean ;
BEGIN
Result:=((GetParamSocSecur('SO_OUVREBIL', '')=StCpte) and (StCpte<>'')) ;
END ;

Function GetDeviseBanque ( Cpte : String ) : String ;
Var Q : TQuery ;
BEGIN
Result:=V_PGI.DevisePivot ;
Q:=OpenSQL('Select BQ_DEVISE from BANQUECP Where BQ_GENERAL="'+Cpte
          +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"',True,-1,'',true) ; // 19/10/2006 YMO Multisociétés
if Not Q.EOF then Result:=Q.Fields[0].AsString ;
Ferme(Q) ;
END ;

Function PasOkDevBqe ( Devi,CodeD : String ) : boolean ;
BEGIN
(*
if VH^.TenueEuro then
   BEGIN
   if ((Devi<>V_PGI.DevisePivot) and (Devi<>V_PGI.DeviseFongible))
      then Result:=(CodeD<>Devi)
      else Result:=((Devi<>V_PGI.DevisePivot) and (Devi<>V_PGI.DeviseFongible)) ;
   END else
   BEGIN
   Result:=(CodeD<>Devi) ;
   END ;
*)
Result:=FALSE ;
if GetParamSocSecur('SO_TENUEEURO', 'X') then
   BEGIN
   if ((Devi<>V_PGI.DevisePivot) and (Devi<>V_PGI.DeviseFongible)) then Result:=(CodeD<>Devi) And ((CodeD<>V_PGI.DevisePivot) and (CodeD<>V_PGI.DeviseFongible)) ;
   END else
   BEGIN
   If (Devi<>V_PGI.DevisePivot) Then Result:=(CodeD<>Devi) And (CodeD<>V_PGI.DevisePivot);
   END ;
END ;

Function PasDeviseBanque ( Cpte,CodeD : String ) : boolean ;
Var Devi : String3 ;
BEGIN
//Result:=False ;
Devi:=GetDeviseBanque(Cpte) ;
Result:=PasOkDevBqe(Devi,CodeD) ;
END ;

Function  EstChaPro( NatGene : String ) : boolean ;
BEGIN
EstChaPro:=((NatGene='CHA') or (NatGene='PRO')) ;
END ;

Function  EstTVATPF ( Compte : String17 ; TVA : Boolean ) : boolean ;
Var i : integer ;
    TOBT : TOB ;
BEGIN
Result:=False ;
{$IFNDEF NOVH}
for i:=0 to VH^.LaTOBTVA.Detail.Count-1 do
    BEGIN
    TOBT:=VH^.LaTOBTVA.Detail[i] ;
    if ((TVA) and (TOBT.GetValue('TV_TVAOUTPF')<>VH^.DefCatTVA)) then Continue ;
    if ((Not TVA) and (TOBT.GetValue('TV_TVAOUTPF')<>VH^.DefCatTPF)) then Continue ;
    if ((TOBT.GetValue('TV_CPTEACH')=Compte) or (TOBT.GetValue('TV_CPTEVTE')=Compte) or
        (TOBT.GetValue('TV_ENCAISACH')=Compte) or (TOBT.GetValue('TV_ENCAISVTE')=Compte)) then BEGIN Result:=True ; Break ; END ;
    END ;
{$ELSE}
for i:=0 to TCPContexte.GetCurrent.LaTOBTVA.Detail.Count-1 do
    BEGIN
    TOBT:=TCPContexte.GetCurrent.LaTOBTVA.Detail[i] ;
    if ((TVA) and (TOBT.GetValue('TV_TVAOUTPF')<>'TX1')) then Continue ;
    if ((Not TVA) and (TOBT.GetValue('TV_TVAOUTPF')<>'TX2')) then Continue ;
    if ((TOBT.GetValue('TV_CPTEACH')=Compte) or (TOBT.GetValue('TV_CPTEVTE')=Compte) or
        (TOBT.GetValue('TV_ENCAISACH')=Compte) or (TOBT.GetValue('TV_ENCAISVTE')=Compte)) then BEGIN Result:=True ; Break ; END ;
    END ;
{$ENDIF}
END ;


Function EstCptEcart ( Cpt : String17 ) : boolean ;
BEGIN
  Result := False;
END ;

Function EstInterdit( StInter : String ; Cpte : String ; Sens : byte ) : byte ;
Var StI,StC,CMin,CMax : String ;
    Find,PosP : byte ;
    SS   : Byte ;
    CS   : Char ;
    Okok : boolean ;
BEGIN
Result:=0 ; if ((Trim(StInter)='') or (Trim(Cpte)='')) then Exit ;
if EstCptEcart(Cpte) then Exit ; 
StI:=StInter ; Find:=0 ; Okok:=False ;
Repeat
 StC:=uppercase(ReadTokenSt(StI)) ;
 if StC<>'' then
    BEGIN
    SS:=0 ; CS:=StC[Length(Stc)] ; if CS='D' then SS:=1 else if CS='C' then SS:=2 ;
    if SS>0 then Delete(StC,Length(StC),1) ;
    PosP:=Pos(':',StC) ;
    if PosP>0 then
       BEGIN
       CMin:=BourreLaDonc(Trim(Copy(StC,1,PosP-1)),fbGene) ; CMax:=Trim(Copy(StC,PosP+1,GetInfoCpta(fbGene).Lg)) ;
       While Length(CMax) < GetInfoCpta(fbGene).Lg do CMax:=CMax+'9' ;
       Cpte:=BourreLaDonc(Cpte,fbGene) ;
       if ((CMin<=Cpte) and (CMax>=Cpte)) then Okok:=True ;
       END else
       BEGIN
       if Copy(Cpte,1,Length(StC))=StC then Okok:=True ;
       END ;
    if Okok then BEGIN if SS=0 then Find:=3 else if SS=Sens then Find:=Sens ; END ;
    END ;
Until ((StC='') or (Find>0)) ;
Result:=Find ;                              
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 21/03/2003
Modifié le ... : 09/04/2003
Description .. : 21/03/2003 : Modif vérification sur date de cloture 
Suite ........ : périodique : on force la relecture de la date pour 
Suite ........ : getsion du multi-utilisateur. 
Suite ........ : 09/04/2003 - LG - 2 versions pour le GetParamSocSecur pour ne 
Suite ........ : pas recharcher suystematiquement en mode non eAgl
Mots clefs ... : 
*****************************************************************}
function ControleDate ( StD : String ) : integer ;
Var DD  : TDateTime ;
    lDtCloPer : TDateTime ;
BEGIN
Result:=0 ;
if Not IsValidDate(StD) then BEGIN Result:=1 ; Exit ; END ;
{$IFDEF CONSOCERIC}
Exit ;
{$ENDIF}
DD:=StrToDate(StD) ;
if DD < GetEncours.Deb then BEGIN Result:=2 ; Exit ; END ;
if ((GetSuivant.Fin>0) and (DD > GetSuivant.Fin)) then BEGIN Result:=2 ; Exit ; END ;
if ((GetSuivant.Fin=0) and (DD > GetEncours.Fin)) then BEGIN Result:=2 ; Exit ; END ;
// Vérif Date de cloture périodique
lDtCloPer := GetParamSocSecur ('SO_DATECLOTUREPER', iDate1900); // PL le 04/10/04 : on doit recharger le paramsoc systématiquement car sinon la date a pu être changée en compta et on ne le saura pas en GI

if ((lDtCloPer>0) and (DD<=lDtCloPer)) then BEGIN Result:=3 ; Exit ; END ;

if ISmoisCloture(StrToDate(StD)) then begin Result := 3; Exit; END;
//
{$IFNDEF GCGC}

if ((GetParamSocSecur('SO_NBJECRAVANT', 0) > 0) and (DD<V_PGI.DateEntree) and (V_PGI.DateEntree-DD > GetParamSocSecur('SO_NBJECRAVANT', 0))) then BEGIN Result:=5 ; Exit ; END ;
if ((GetParamSocSecur('SO_NBJECRAPRES', 0) > 0) and (DD>V_PGI.DateEntree) and (DD-V_PGI.DateEntree > GetParamSocSecur('SO_NBJECRAPRES', 0))) then BEGIN Result:=5 ; Exit ; END ;

{$ENDIF}
END ;


function  QuelTypeMvt ( Compte : String17 ; NC,NP : String3 ) : String ;
BEGIN
QuelTypeMvt:='DIV' ;
if ((NC='') or (NP='')) then Exit ;
if (QuelPaysLocalisation<>CodeISOES) and
   ((NP<>'FC') and (NP<>'AC') and (NP<>'FF') and (NP<>'AF')) then Exit ; // XVI 24/02/2005
if ((NC='TID') or (NC='TIC') or (NC='CLI') or (NC='FOU') or (NC='AUD') or (NC='AUC')) then QuelTypeMvt:='TTC' else
 if EstChaPro(NC) then QuelTypeMvt:='HT' else
  if EstTVATPF(Compte,True) then QuelTypeMvt:='TVA' else
   if EstTVATPF(Compte,False) then QueltypeMvt:='TPF' else
    // Point 62 FFF
    if (NC='DIV') then QueltypeMvt:='TTC';
END ;


(*======================================================================*)
function EnDevise ( Code : String3 ) : boolean ;
BEGIN
Result:=(Code<>V_PGI.DevisePivot) ;
END ;

(*======================================================================*)
function QuelleLigne ( Var St : hString ) : integer ;
Var Pos1,Pos2 : integer ;
BEGIN
QuelleLigne:=-2 ;
Pos2:=Pos(':L',St) ; if Pos2>0 then System.Delete(St,Pos2+1,1) ;
Pos1:=Pos(':',St) ; if ((Pos1<=0) or (Pos1=Length(St))) then Exit ;
QuelleLigne:=Round(Valeur(Copy(St,Pos1+1,5))) ;
System.Delete(St,Pos1,5) ;
END ;

{$IFNDEF SANSCOMPTA}
(*======================================================================*)
procedure InitCommunObjAnal( QuelEcr : TTypeEcr ; OBM : TOBM ; OBA : TObjAna ) ;
Var Pf : String[2] ;
begin
if OBM=Nil then Exit ;
if QuelEcr=EcrGen then Pf:='Y' ;
if QuelEcr=EcrGen then
   BEGIN
   OBM.PutMvt('Y_GENERAL',OBA.General) ;
{$IFNDEF SPEC302}
   OBM.PutMvt('Y_PERIODE',GetPeriode(OBA.DateComptable)) ;
   OBM.PutMvt('Y_SEMAINE',NumSemaine(OBA.DateComptable)) ;
{$ENDIF}
   END ;
OBM.PutMvt(Pf+'_EXERCICE',OBA.EXERCICE) ;
OBM.PutMvt(Pf+'_DATECOMPTABLE',OBA.DateComptable) ; OBM.PutMvt(Pf+'_NUMEROPIECE',OBA.NumeroPiece) ;
OBM.PutMvt(Pf+'_NUMLIGNE',OBA.NumLigne)           ; OBM.PutMvt(Pf+'_REFINTERNE',Copy(OBA.RefInterne,1,35)) ;
OBM.PutMvt(Pf+'_LIBELLE',Copy(OBA.Libelle,1,35))  ; OBM.PutMvt(Pf+'_NATUREPIECE',OBA.NaturePiece) ;
if ((Pf='Y') and (OBA.RefLibre<>'')) then OBM.PutMvt(Pf+'_REFLIBRE',Copy(OBA.RefLibre,1,35)) ;
OBM.PutMvt(Pf+'_UTILISATEUR',OBA.Utilisateur)     ; OBM.PutMvt(Pf+'_CONTROLEUR',OBA.Controleur) ;
OBM.PutMvt(Pf+'_SOCIETE',OBA.Societe)             ; OBM.PutMvt(Pf+'_ETABLISSEMENT',OBA.Etablissement) ;
OBM.PutMvt(Pf+'_TOTALECRITURE',OBA.TotalEcriture) ;
OBM.PutMvt(Pf+'_TAUXDEV',OBA.TauxDev)             ; OBM.PutMvt(Pf+'_DATETAUXDEV',OBA.DateTauxDev) ;
OBM.PutMvt(Pf+'_DEVISE',OBA.Devise)               ;
OBM.PutMvt(Pf+'_JOURNAL',OBA.Journal)             ; OBM.PutMvt(Pf+'_TOTALDEVISE',OBA.TotalDevise) ;
OBM.PutMvt(Pf+'_QUALIFPIECE',OBA.QualifPiece) ;
OBM.PutMvt(Pf+'_TOTALQTE1',OBA.TotalQte1)         ; OBM.PutMvt(Pf+'_TOTALQTE2',OBA.TotalQte2) ;
{infos supplémentaires}
if VarAsType(OBM.GetMvt(Pf+'_QUALIFQTE1'),VarString)='' then OBM.PutMvt(Pf+'_QUALIFQTE1',OBA.QualifQte1) ;
if VarAsType(OBM.GetMvt(Pf+'_QUALIFQTE2'),VarString)='' then OBM.PutMvt(Pf+'_QUALIFQTE2',OBA.QualifQte2) ;
if QuelEcr in [EcrGen,EcrBud] then
   BEGIN
   OBM.PutMvt(Pf+'_QUALIFECRQTE1',OBA.QualifQte1) ; OBM.PutMvt(Pf+'_QUALIFECRQTE2',OBA.QualifQte2) ;
   END ;
if QuelEcr=EcrGen then
   BEGIN
   if VarAsType(OBM.GetMvt(Pf+'_REFEXTERNE'),VarString)='' then OBM.PutMvt(Pf+'_REFEXTERNE',OBA.RefExterne) ;
   if VarAsType(OBM.GetMvt(Pf+'_DATEREFEXTERNE'),VarDate)=IDate1900 then OBM.PutMvt(Pf+'_DATEREFEXTERNE',OBA.DateRefExterne) ;
   if VarAsType(OBM.GetMvt(Pf+'_AFFAIRE'),VarString)='' then OBM.PutMvt(Pf+'_AFFAIRE',OBA.Affaire) ;
   END ;
end ;

function  RechercheLente ( St,Cpte : String ) : Variant ;
Var C : Char ;
    FF,Code : String ;
    TT      : TField ;
    Q       : TQuery ;
BEGIN
Result:=#0 ; if ((Trim(St)='') or (Trim(Cpte)='')) then Exit ;
C:=St[1] ;
Case C of
   'J' : BEGIN FF:='JOURNAL' ; Code:='J_JOURNAL' ; END ;
   'G' : BEGIN FF:='GENERAUX' ; Code:='G_GENERAL' ; END ;
   'T' : BEGIN FF:='TIERS' ; Code:='T_AUXILIAIRE' ; END ;
   else Exit ;
   END ;
Q:=OpenSQL('Select * from '+FF+' Where '+Code+'="'+Cpte+'"',True,-1,'',true) ;
if Not Q.EOF then
   BEGIN
   TT:=Q.FindField(St) ; if TT<>Nil then Result:=TT.AsVariant ;
   END ;
Ferme(Q) ;
END ;
{$ENDIF}

function GetNewNumPiece ( Souche : String3 ; Ecrit : boolean ; DD : TDateTime) : Longint ;
Var Q : TQuery ;
    Num : Longint ;
    SQL : String ;
BEGIN
Num:=0 ;
Q:=OpenSQL('Select SH_NUMDEPART from SOUCHE Where SH_TYPE="GES" AND SH_SOUCHE="'+Souche+'"',True,-1,'',true) ;
if Not Q.EOF then Num:=Q.Fields[0].AsInteger else V_PGI.IoError:=oeUnknown ;
Ferme(Q) ;
if ((Num>0) and (Ecrit)) then
   BEGIN
   SQL:='UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(Num+1)+' Where SH_TYPE="GES" AND SH_SOUCHE="'+Souche+'" AND SH_NUMDEPART='+IntToStr(Num) ;
   if ExecuteSQL(SQL)<=0 then BEGIN Num:=0 ; V_PGI.IoError:=oeSaisie ; END ;
   END ;
Result:=Num ;
END ;

{function GetNewNumJal ( Jal : String3 ; Normale : boolean ; DD : TDateTime) : Longint ;
Var Q : TQuery ;
    Souche : String ;
BEGIN
Result:=0 ;
if Normale then Q:=OpenSQL('Select J_COMPTEURNORMAL from JOURNAL Where J_JOURNAL="'+Jal+'"',True)
           else Q:=OpenSQL('Select J_COMPTEURSIMUL from JOURNAL Where J_JOURNAL="'+Jal+'"',True) ;
if Not Q.EOF then
   BEGIN
   Souche:=Q.Fields[0].AsString ;
   if Souche<>'' then SetIncNum(EcrGen,Q.Fields[0].AsString,Result,DD) ;
   END ;
Ferme(Q) ;
END ;}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/04/2003
Modifié le ... :   /  /    
Description .. : - 24/04/2003  - prise en compte du e_qualiforigine
Mots clefs ... : 
*****************************************************************}
function GetNewNumJal ( Jal : String3; Normale : boolean; DD : TDateTime ; Cpt : string = '' ; ModeSaisie : string = '' ; vDossier : String = '' ) : Longint;
var
 Q                    : TQuery;
 Souche               : string;
 lYear                : Word ;
 lMonth               : Word ;
 lDay                 : Word ;
 lMaxJour             : integer;
 lStDateDebMois       : string;
 lStDateFinMois       : string;
 lStModeSaisie        : string;
begin

 Result := 0;

 if ( Cpt = '' ) or ( ModeSaisie = '' ) then
  begin // recherche des info du journal

  if Normale
    then Q := OpenSelect ( 'Select J_COMPTEURNORMAL, J_MODESAISIE from JOURNAL Where J_JOURNAL="' + Jal + '"', vDossier )
    else Q := OpenSelect ( 'Select J_COMPTEURSIMUL, J_MODESAISIE from JOURNAL Where J_JOURNAL="' + Jal + '"', vDossier ) ;

   if not Q.EOF then
    begin
     Souche        := Q.Fields[0].AsString;
     lStModeSaisie := Q.Fields[1].AsString;
    end
     else
      begin  // le journal n'existe pas -> on sort
       Ferme( Q);
       exit;
      end; // if

   Ferme( Q);

  end
   else
    begin
     // on recupere les param de la fct
     Souche        := Cpt;
     lStModeSaisie := ModeSaisie;
    end; // if


   if lStModeSaisie = '-' then
     begin // mode piece
      if Souche <> '' then
       SetIncNum ( EcrGen, Souche, Result, DD, vDossier ) ;
     end
      else
       begin // mode bordereau ou libre

        DecodeDate( DD , lYear , lMonth , lDay ) ;

        lMaxJour          := DaysPerMonth( lYear, lMonth) ;
        lStDateDebMois    := USDateTime( EncodeDate( lYear, lMonth, 1) );
        lStDateFinMois    := USDateTime( EncodeDate( lYear, lMonth, lMaxJour));

        Q := OpenSelect ( 'select MAX(E_NUMEROPIECE) as N from ECRITURE where E_EXERCICE = "' + QuelExoDt ( DD )  + '" '
                        + 'and E_JOURNAL = "' + Jal + '" and E_DATECOMPTABLE >="' +  lStDateDebMois + '" '
                        + 'and E_DATECOMPTABLE <="' +  lStDateFinMois + '" and E_QUALIFPIECE="N" ', vDossier );

        result            := Q.FindField('N').asInteger + 1;

        Ferme( Q );

       end; // if

end;


{$IFDEF EAGLCLIENT}
procedure GereSelectionsGrid ( G : THGrid ; Q : TQuery ) ;
BEGIN
G.ClearSelected ;
if (Q = nil) then Exit;
if Q.EOF then
   BEGIN
   if G.MultiSelect then G.MultiSelect := False ;
   END else
   BEGIN
   if Not G.MultiSelect then G.MultiSelect := True ;
   END ;
END ;
{$ELSE}
  {$IFNDEF EAGLSERVER}
procedure GereSelectionsGrid ( G : THDBGrid ; Q : THQuery ) ;
BEGIN
G.ClearSelected ;
if (Q = nil) then Exit;
if Q.EOF then
   BEGIN
   if G.MultiSelection then G.MultiSelection := False ;
   END else
   BEGIN
   if Not G.MultiSelection then G.MultiSelection := True ;
   END ;
END ;
  {$ENDIF}
{$ENDIF}

Function SENSENC ( D,C : Double ) : String3 ;
Var Solde : Double ;
BEGIN
SensEnc:='' ;
Solde:=D-C ; if Solde=0 then Exit ;
if Solde>0 then SensEnc:='ENC' else SensEnc:='DEC' ;
END ;

function ModeRevisionActive ( DD : TDateTime ) : boolean ;
BEGIN
Result:=((Not V_PGI.Controleur) and (GetParamSocSecur('SO_DATEREVISION', Idate1900)>IDate1900) and (DD<=GetParamSocSecur('SO_DATEREVISION', iDate1900))) ;
END ;

function RevisionActive ( DD : TDateTime ) : boolean ;
BEGIN
Result:=ModeRevisionActive(DD) ;
if Result then HShowMessage('0;Révision;Le mode révision est actif. Seul un contrôleur peut créer ou modifier des écritures avant le '+DateToStr(GetParamSocSecur('SO_DATEREVISION', iDate1900))+';W;O;O;O;','','') ;
END ;

function DepasseLimiteDemo : boolean ;
BEGIN
Result:=False ;
//  -------------------------------------------------------------------------------------------
//  En version 6, le nombre d'utilisation des saisies ne doit plus être limitée en version démo
//  -------------------------------------------------------------------------------------------
(* code conservé en prévision du contre-ordre ;-)
  if ((Not V_PGI.VersionDemo) or (V_PGI.SAV)) then Exit ;
  St:='' ; St:=GetSynRegKey('SCRC',St,TRUE) ; if St='123456' then Exit ;
  if St='' then St:=CryptageSt('1234'+'0') ;
  St:=DeCryptageSt(St) ; Delete(St,1,4) ;
  if St<>'' then Nb:=StrToInt(St) else Nb:=0 ;
  Inc(Nb) ;
  if Nb>100 then
     BEGIN
     HShowMessage('0;Version de démonstration;Vous avez dépassé le nombre de saisies autorisées en version de démonstration;W;O;O;O;','','') ;
     Result:=True ; Exit ;
     END ;
  St:=IntToStr(Nb) ; St:=CryptageSt('1234'+St) ;
  SaveSynRegKey('SCRC',St,TRUE) ;
*)
END ;

function PasCreerDate ( DD : TDateTime ) : boolean ;
Var ii : integer ;
BEGIN
if Not ExJaiLeDroitConcept(TConcept(ccSaisEcritures),True) then Result:=True else
   BEGIN
   ii:=ControleDate(DateToStr(DD)) ;
   Result:=(ii>0) ;
   if Result then HShowMessage('0;Comptabilité;La date d''entrée est incompatible avec la saisie. Vous ne pouvez pas créer de pièce au '+DateToStr(DD)+';W;O;O;O;','','')
             else Result:=RevisionActive(DD) ;
   END ;           
END ;

function PasCreerDateGC ( DD : TDateTime ) : boolean ;
Var ii : integer ;
BEGIN
ii:=ControleDate(DateToStr(DD)) ;
Result:=((ii>0) and (not GetParamSocSecur('SO_GCDESACTIVECOMPTA',False))) ;
if Result then HShowMessage('0;'+TitreHalley+';La date d''entrée est incompatible avec la saisie. Vous ne pouvez pas créer de pièce au '+DateToStr(DD)+';W;O;O;O;','','') ;
END ;

function  NbJoursOk ( DatePiece,DateEche : TDateTime ) : boolean ;
BEGIN
Result:=False ;
// FQ17165 SBO 25/04/2006 : Erreur sur le paramSoc utilisé...
if ((GetParamSocSecur('SO_NBJECHAVANT', 0) > 0) and (DateEche<DatePiece) and (DatePiece-DateEche > GetParamSocSecur('SO_NBJECHAVANT', 0))) then Exit ;
if ((GetParamSocSecur('SO_NBJECHAPRES', 0) > 0) and (DatePiece<DateEche) and (DateEche-DatePiece > GetParamSocSecur('SO_NBJECHAPRES', 0))) then Exit ;
Result:=True ;
END ;

function  GetChampsTiers ( ct : TChampsTiers ) : String ;
Var St : String ;
BEGIN
Result:='' ;
Case ct of
   ctRelance   : St:='T_AUXILIAIRE, T_RELANCETRAITE, T_RELANCEREGLEMENT, T_LIBELLE, T_PAYS, T_CODEPOSTAL, T_SCORERELANCE, T_EAN, T_ADRESSE1, T_SECTEUR, T_TELEPHONE ';
   ctCoutTiers : St:='T_AUXILIAIRE, T_LIBELLE, T_NATUREAUXI, T_EAN, T_ADRESSE1, T_SECTEUR, T_CODEPOSTAL, T_VILLE, T_TELEPHONE ' ;
   ctRelFac    : St:='T_AUXILIAIRE, T_LIBELLE, T_EAN, T_DATEDERNRELEVE, T_DERNLETTRAGE, T_FREQRELEVE, T_JOURRELEVE ' ;
   ctExpMvt    : St:='T_AUXILIAIRE, T_NATUREAUXI ' ;
   END ;
Result:=St ;
END ;

{$IFNDEF NOVH}
Procedure SetMPACCDB ( DS : TDataSet) ;
BEGIN
{$IFDEF EAGLCLIENT}
  DS.PutValue('E_CODEACCEPT', MPTOACC( DS.GetValue('E_MODEPAIE') ) ) ;
{$ELSE}
  DS.FindField('E_CODEACCEPT').AsString := MPTOACC(DS.FindField('E_MODEPAIE').AsString) ;
{$ENDIF}
END ;
{$ENDIF}

Procedure SetCotationDB ( DD : TDateTime ; DS : TDataSet) ;
Var Cote,Taux : Double ;
    Dev       : String3 ;
BEGIN
{$IFDEF EAGLCLIENT}
  if DD <= 0 then
    DD := DS.GetDateTime('E_DATECOMPTABLE') ; {JP 30/05/06 : FQ 18242}
  Taux := DS.GetDouble('E_TAUXDEV') ; {JP 30/05/06 : FQ 18242}
  if DD < V_PGI.DateDebutEuro
    then Cote:=Taux
    else
      BEGIN
      Dev := DS.GetString('E_DEVISE') ; {JP 30/05/06 : FQ 18242}
      if ((Dev = V_PGI.DevisePivot) or (Dev = V_PGI.DeviseFongible))
        then Cote := 1.0
        else if V_PGI.TauxEuro <> 0
               then Cote := Taux / V_PGI.TauxEuro
               else Cote := 1 ;
      END ;
  DS.PutValue('E_COTATION' , Cote ) ;
{$ELSE}
  if DD<=0 then DD:=DS.FindField('E_DATECOMPTABLE').AsDateTime ;
  Taux:=DS.FindField('E_TAUXDEV').AsFloat ;
  if DD<V_PGI.DateDebutEuro then Cote:=Taux else
    BEGIN
    Dev:=DS.FindField('E_DEVISE').AsString ;
    if ((Dev=V_PGI.DevisePivot) or (Dev=V_PGI.DeviseFongible)) then Cote:=1.0 else
    if V_PGI.TauxEuro<>0 then Cote:=Taux/V_PGI.TauxEuro else Cote:=1 ;
    END ;
  DS.FindField('E_COTATION').AsFloat:=Cote ;
{$ENDIF}
END ;

function QuellePeriode ( UneDate,DebExo : TDateTime ) : byte ;
Var i : byte ;
BEGIN
i:=0 ; While ((DebExo<=UneDate) and (i<24)) do BEGIN DebExo:=PlusMois(DebExo,1) ; Inc(i) ; END ;
QuellePeriode:=i ;
END ;

function ISmoisCloture (UneDate : TdateTime) : boolean;
var IMois : integer;
		QQ : TQuery;
    Exercice,Periode : string;
    DateDebutEx : TDateTime;
begin
	Exercice := QUELEXODT(UneDate);
  QQ := OpenSql('Select EX_DATEDEBUT,EX_VALIDEE From EXERCICE Where EX_EXERCICE="' + Exercice + '"',True);
  Periode := QQ.FindField('EX_VALIDEE').AsString;
  DateDebutEx := QQ.FindField('EX_DATEDEBUT').AsDateTime;
  Ferme(QQ);
  Result := (Periode[QuellePeriode(UneDate,DateDebutEx)]='X');
end;

function isMoisJournalCloture (Journal: string; UneDate : TdateTime) : Boolean;
var SQl : string;
		yy,mm,dd : Word;
    OneDate : TDateTime;
begin
  DecodeDate(UneDate,yy,mm,dd);
  OneDate := EncodeDate(yy,mm,1);
  Sql := 'SELECT 1 FROM CLOTPERJOU WHERE CCJ_JOURNAL="'+Journal+'" AND '+
  			 'CCJ_PERIODCLOTURE="'+USDateTime(OneDate)+'"';
  Result :=  ExisteSql(Sql);
end;

Function ADevalider ( UnJal : String ; UneDate : TDateTime ) : Boolean ;
Var DebExo : TDateTime ;
    LExo,Stj,Ste : String ;
    QuelPeriodJal : String ;
    i : Byte ;
    Q : TQuery ;
BEGIN
Result:=False ; if UnJal='' then Exit ;
if Not IsValidDate(DateToStr(UneDate)) then Exit ;
if (UneDate >= GetEncours.Deb) And (UneDate <= GetEncours.Fin) then
   BEGIN
   Lexo := GetEncours.Code ; DebExo := GetEnCours.Deb ; QuelPeriodJal:='J_VALIDEEN' ;
   END else if (UneDate >= GetSuivant.Deb) And (UneDate <= GetSuivant.Fin) then
   BEGIN
   Lexo := GetSuivant.Code ; DebExo := GetSuivant.Deb ; QuelPeriodJal:='J_VALIDEEN1' ;
   END else Exit ;
i:=QuellePeriode(UneDate,DebExo) ;
Q:=OpenSql('Select '+QuelPeriodJal+' from JOURNAL Where J_JOURNAL="'+UnJal+'"',True,-1,'',true) ;
Stj:=Q.Fields[0].AsString ;
Ferme(Q) ;
if Stj='' then Exit ;
if Stj[i]='-' then Exit else
 BEGIN
 Q:=OpenSql('Select EX_VALIDEE from EXERCICE Where EX_EXERCICE="'+Lexo+'"',True,-1,'',true) ;
 Ste:=Q.Fields[0].AsString ;
 Ferme(Q) ;
 if ((Ste<>'') and (Ste[i]='X')) then
    BEGIN
    Ste[i]:='-' ;
    ExecuteSql('UPDATE EXERCICE SET EX_VALIDEE="'+Ste+'" Where EX_EXERCICE="'+Lexo+'"') ;
    END ;
 Stj[i]:='-' ;
 ExecuteSql('UPDATE JOURNAL SET '+QuelPeriodJal+'="'+Stj+'" Where J_JOURNAL="'+UnJal+'"') ;
 END ;
Result:=True ;
END ;

{$IFNDEF SANSCOMPTA}
  {$IFNDEF EAGLSERVER}
procedure EditEtatS5S7 (Tip,Nat,Modele : String ; AvecCreation : Boolean ; Spages : TPageControl ; StSQL,Titre : String );
BEGIN
EditEtat(Tip,Nat,Modele,AvecCreation,Spages,StSQL,Titre) ;
if Nat='RLC' then AvertirTable('ttModeleRLC') else
if Nat='REL' then AvertirTable('ttModeleRelance') else
if Nat='LCH' then AvertirTable('ttModeleLettreCHQ') else
if Nat='LTR' then BEGIN AvertirTable('ttModeleLettreTRA') ; AvertirTable('ttModeleBOR') ; END else
if Nat='UCO' then AvertirTable('ttModeleDetail') else
if Nat='BOR' then AvertirTable('ttModeleBOR') else
if Nat='SAI' then AvertirTable('ttModeleSAI') else
if Nat='BAP' then AvertirTable('ttModeleBAP') ;
END ;
  {$ENDIF}
  {$IFNDEF EAGLCLIENT}
    {$IFNDEF EAGLSERVER}
procedure EditDocumentS5S7 (Tip,Nat,Modele : String ; AvecCreation : Boolean);
BEGIN
EditDocument(Tip,Nat,Modele,AvecCreation) ;
if Nat='RLC' then AvertirTable('ttModeleRLC') else
if Nat='REL' then AvertirTable('ttModeleRelance') else
if Nat='LCH' then AvertirTable('ttModeleLettreCHQ') else
if Nat='LTR' then BEGIN AvertirTable('ttModeleLettreTRA') ; AvertirTable('ttModeleBOR') ; END else
if Nat='UCO' then AvertirTable('ttModeleDetail') else
if Nat='BOR' then AvertirTable('ttModeleBOR') else
if Nat='SAI' then AvertirTable('ttModeleSAI') else
if Nat='BAP' then AvertirTable('ttModeleBAP') ;
END ;
	  {$ENDIF}
  {$ENDIF}
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : correction de l'affectation des champs suite au nouvel AGL
Mots clefs ... :
*****************************************************************}
(*
Procedure EcrVersAna ( TOBEcr,TOBAna : TOB ) ;
Var NomEcr,NomAna : String ;
    i,j           : integer ;
    TOBAxe        : TOB ;
BEGIN
for i:=1 to TOBEcr.NbChamps do
    BEGIN
    NomEcr:=TOBEcr.GetNomChamp(i) ;
    // le champs e_vision de la table ecriture est un varchar de 3 et celui
    // de la table analytique est un vachar(1) -> traitement particulier
    if NomEcr='E_VISION' then BEGIN TOBAna.PutValue('Y_VISION','-') ; continue ; END ;
    NomAna:=NomEcr ; NomAna[1]:='Y' ;
    j:=TOBAna.GetNumChamp(NomAna) ;
    if j>0 then TOBAna.PutValeur(j,TOBEcr.GetValeur(i)) ;
    END ;
TOBAxe:=TOBAna.Parent ; TOBAna.PutValue('Y_AXE',TOBAxe.NomTable) ;
TOBAna.PutValue('Y_TOTALECRITURE',TOBEcr.GetValue('E_DEBIT')+TOBEcr.GetValue('E_CREDIT')) ;
TOBAna.PutValue('Y_TOTALDEVISE',TOBEcr.GetValue('E_DEBITDEV')+TOBEcr.GetValue('E_CREDITDEV')) ;
TOBAna.PutValue('Y_TOTALEURO',TOBEcr.GetValue('E_DEBITEURO')+TOBEcr.GetValue('E_CREDITEURO')) ;
END ;
*)
Procedure EcrVersAna ( TOBEcr,TOBAna : TOB ) ;
Var NomEcr,NomAna : String ;
    i,j,FNumAna,FNumEcr : integer ;
    TOBAxe        : TOB ;
		Mcd : IMCDServiceCOM;
    TypeEcr,TypeAna : string;
BEGIN
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  for i:=1 to TOBEcr.NbChamps do
  BEGIN
    NomEcr:=TOBEcr.GetNomChamp(i) ;
    // le champs e_vision de la table ecriture est un varchar de 3 et celui
    // de la table analytique est un vachar(1) -> traitement particulier
    if NomEcr='E_VISION' then BEGIN TOBAna.PutValue('Y_VISION','-') ; continue ; END ;
    if NomEcr='E_TAUXDEV' then BEGIN TOBAna.PutValue('Y_TAUXDEV',TOBEcr.GetValue('E_TAUXDEV')) ; Continue ; END ;
    TypeEcr := (Mcd.getField(NomEcr) as IFieldCom).Tipe;

    NomAna:=NomEcr ; NomAna[1]:='Y' ;
    if Mcd.FieldExists(NomAna) then
    BEGIN
      TypeAna := (Mcd.getField(NomAna) as IFieldCom).Tipe;
      if TypeEcr=TypeAna then TOBAna.PutValue(NomAna,TOBEcr.GetValue(NomEcr)) ;
    END ;
  END ;
  TOBAxe:=TOBAna.Parent ;
  if TOBAxe.GetNumChamp('Y_AXE') < 0
  then TOBAna.PutValue('Y_AXE', 'A' + IntToStr( TOBAxe.GetIndex + 1 ) )
  else TOBAna.PutValue('Y_AXE',TOBAxe.GetString('Y_AXE') ) ;
  TOBAna.PutValue('Y_TOTALECRITURE',TOBEcr.GetValue('E_DEBIT')+TOBEcr.GetValue('E_CREDIT')) ;
  TOBAna.PutValue('Y_TOTALDEVISE',TOBEcr.GetValue('E_DEBITDEV')+TOBEcr.GetValue('E_CREDITDEV')) ;
  // V9 CEGID
  TOBAna.PutValue('Y_DATPER',iDate1900) ;
  TOBAna.PutValue('Y_ENTITY',0) ;
  TOBAna.PutValue('Y_REFGUID','') ;
END ;

Procedure VentilLigneTOB ( TOBAna : TOB ; Section : String ; NumVentil,NbDec : integer ; Pourc : double ; Debit : boolean ; PourQte1 : double = 0.0 ; PourQte2 : double = 0.0 ; splan1 : string = '' ; splan2 : string = ''; splan3 : string = '';
                  splan4 : string = '' ; splan5 : string = '') ;
Var XP,XD : double ;
    qte1, qte2 : double;
BEGIN
  TOBAna.PutValue('Y_SECTION',Section) ;
  TOBAna.PutValue('Y_NUMVENTIL',NumVentil) ;
  TOBAna.PutValue('Y_POURCENTAGE',Pourc) ;
  //SG6 Gestion mode croisaxe
  if VH^.LiaisonY2ViaShare then
  begin
    TOBAna.PutValue('Y_SOUSPLAN1', Section);
  end else
  begin
    TOBAna.PutValue('Y_SOUSPLAN1', splan1);
  end;
  TOBAna.PutValue('Y_SOUSPLAN2', splan2);
  TOBAna.PutValue('Y_SOUSPLAN3', splan3);
  TOBAna.PutValue('Y_SOUSPLAN4', splan4);
  TOBAna.PutValue('Y_SOUSPLAN5', splan5);
  //SG6 Prise en compte des quantités
  TOBAna.PutValue('Y_POURCENTQTE1',PourQte1);
  TOBAna.PutValue('Y_POURCENTQTE2',PourQte2);
  //Calcul dse quantites effectives
  qte1 := TOBAna.GetDouble('Y_TOTALQTE1');
  qte2 := TOBAna.GetDouble('Y_TOTALQTE2');
  TOBAna.PutValue('Y_QTE1', Arrondi(PourQte1 * qte1 / 100.0, V_PGI.OkDecQ));
  TOBAna.PutValue('Y_QTE2', Arrondi(PourQte2 * qte2 / 100.0, V_PGI.OkDecQ));
  XP:=Arrondi(Pourc*TOBAna.GetValue('Y_TOTALECRITURE')/100.0,V_PGI.OkDecV) ;
  XD:=Arrondi(Pourc*TOBAna.GetValue('Y_TOTALDEVISE')/100.0,NbDec) ;
  if Debit then
  BEGIN
    TOBAna.PutValue('Y_DEBIT',XP) ; TOBAna.PutValue('Y_DEBITDEV',XD) ;
    TOBAna.PutValue('Y_CREDIT',0) ; TOBAna.PutValue('Y_CREDITDEV',0) ;
  END else
  BEGIN
    TOBAna.PutValue('Y_DEBIT',0)   ; TOBAna.PutValue('Y_DEBITDEV',0)   ;
    TOBAna.PutValue('Y_CREDIT',XP) ; TOBAna.PutValue('Y_CREDITDEV',XD) ;
  END ;
  // V9 CEGID
  TOBAna.PutValue('Y_DATPER',iDate1900) ;
  TOBAna.PutValue('Y_ENTITY',0) ;
  TOBAna.PutValue('Y_REFGUID','') ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 01/07/2004
Modifié le ... : 01/06/2006
Description .. : FB 13748 - ME - ajout d'un parametre pour le parcourd des
Suite ........ : axes analytiques
Suite ........ :
Suite ........ : FQ 18180 - SBO - ajout solde des qte ana
Mots clefs ... :
*****************************************************************}
Procedure ArrondirAnaTOB ( TOBEcr : TOB ; NbDec : integer; vInAxe : integer = 0 ) ;
Var Pourc,XP,XD,TotP,TotD,TotPourc,TotalEcriture,TotalDevise  : double ;
    NumAxe,i : integer ;
    lInCpt   : integer;                 {FP 12/06/2006 FQ18226}
    fb       : TFichierBase ;
    TOBAxe,TOBAna : TOB ;
    PourcQte1 : double ;
    PourcQte2 : double ;
    TotQte1   : double ;
    TotQte2   : double ;
    lInDeb    : integer ;
    lInFin    : integer ;
BEGIN

  if vInAxe = 0 then
    begin
    lInDeb  := 1 ;
    lInFin  := MAXAXE ;
    end
  else
    begin
    lInDeb  := vInAxe ;
    lInFin  := vInAxe ;
    end ;

//Capturer les montants de la ligne géné
TotalEcriture:=TOBEcr.GetValue('E_DEBIT')+TOBEcr.GetValue('E_CREDIT') ;
TotalDevise:=TOBEcr.GetValue('E_DEBITDEV')+TOBEcr.GetValue('E_CREDITDEV') ;
for NumAxe:=lInDeb to lInFin do
  if TOBEcr.Detail[NumAxe-1].Detail.Count>0 then
    BEGIN
    // Prendre les axes ventilés
    fb:=AxeToFb('A'+IntToStr(NumAxe)) ;
    TotP:=0 ; TotD:=0 ; TotPourc:=0 ;
    PourcQte1 := 0 ; PourcQte2 := 0 ;
    TotQte1 := 0 ; TotQte2 := 0 ;
    TOBAxe:=TOBEcr.Detail[NumAxe-1] ;
    for i:=0 to TOBAxe.Detail.Count-1 do
        BEGIN
        // Sommations
        TOBAna    := TOBAxe.Detail[i] ;
        PourcQte1 := PourcQte1 + TOBAna.GetValue('Y_POURCENTQTE1') ;
        PourcQte2 := PourcQte1 + TOBAna.GetValue('Y_POURCENTQTE2') ;
        TotQte1   := TotQte1 + TOBAna.GetValue('Y_QTE1') ;
        TotQte2   := TotQte2 + TOBAna.GetValue('Y_QTE2') ;
        Pourc     := TOBAna.GetValue('Y_POURCENTAGE') ;
        TotPourc  := Arrondi(TotPourc+Pourc,ADecimP) ;
        if ((i<TOBAxe.Detail.Count-1) or (Arrondi(TotPourc-100.0,ADecimP)<>0)) then
           BEGIN
           // Si pas dernière ligne ou tot % <> 100 % alors cumuler
           XP:=TOBAna.GetValue('Y_DEBIT')+TOBAna.GetValue('Y_CREDIT') ;
           XD:=TOBAna.GetValue('Y_DEBITDEV')+TOBAna.GetValue('Y_CREDITDEV') ;
           TotP:=Arrondi(TotP+XP,V_PGI.OkDecV) ; TotD:=Arrondi(TotD+XD,NbDec) ;
           END else
           BEGIN
           // Si Dernière ligne et tot % = 100 % alors répercuter systématiquement le reste sur la dernière ligne
           XP:=Arrondi(TotalEcriture-TotP,V_PGI.OkDecV) ;
           XD:=Arrondi(TotalDevise-TotD,NbDec) ;
           if TOBEcr.GetValue('E_DEBIT')<>0 then
              BEGIN
              TOBAna.PutValue('Y_DEBIT',XP) ; TOBAna.PutValue('Y_DEBITDEV',XD) ;
              END else
              BEGIN
              TOBAna.PutValue('Y_CREDIT',XP) ; TOBAna.PutValue('Y_CREDITDEV',XD) ;
              END ;
           END ;
        END ;
     if (Arrondi(TotPourc-100,ADecimP)<>0) and (TotalEcriture<>TotP) then
        BEGIN
        // Si ventilation incomplète, ventiler le reste sur une nouvelle ligne avec section d'attente
        TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
        TobAna.InitValeurs ;
        EcrVersAna(TOBEcr,TOBAna) ;
        TOBAna.PutValue('Y_SECTION', GetInfoCpta(fb).Attente) ;
        {b FP 12/06/2006 FQ18226: Les champs sous-plan ne sont pas alimentés}
        if GetParamSocSecur('SO_CROISAXE', False) then
          for lInCpt := 1 to MaxAxe do
            begin
            if GetParamSocSecur('SO_VENTILA' + IntToStr(lInCpt), False) then
              TOBAna.PutValue('Y_SOUSPLAN'+IntToStr(lInCpt), GetInfoCpta(AxeToFb('A'+IntToStr(lInCpt))).Attente);
            end;
        {e FP 12/06/2006 FQ18226}
        TOBAna.PutValue('Y_NUMVENTIL',TOBAxe.Detail.Count) ;
        TOBAna.PutValue('Y_POURCENTAGE',Arrondi(100.0-TotPourc,ADecimP)) ;
        XP:=Arrondi(TotalEcriture-TotP,V_PGI.OkDecV) ;
        XD:=Arrondi(TotalDevise-TotD,NbDec) ;
        if TOBEcr.GetValue('E_DEBIT')<>0 then
           BEGIN
           TOBAna.PutValue('Y_DEBIT',XP) ; TOBAna.PutValue('Y_DEBITDEV',XD) ;
           END else
           BEGIN
           TOBAna.PutValue('Y_CREDIT',XP) ; TOBAna.PutValue('Y_CREDITDEV',XD) ;
           END ;
        // Gestion du solde des quantité : QTE1
        if (TOBEcr.GetValue('E_QTE1') <> 0) and (Arrondi(100-PourcQte1, ADecimP)<>0 ) then
          begin
          TobAna.PutValue('Y_POURCENTQTE1', Arrondi(100-PourcQte1, ADecimP) ) ;
          TobAna.PutValue('Y_QTE1',         TOBEcr.GetValue('E_QTE1') - TotQte1 ) ;
          end ;
        // Gestion du solde des quantité : QTE2
        if (TOBEcr.GetValue('E_QTE2') <> 0) and (Arrondi(100-PourcQte2, ADecimP)<>0 ) then
          begin
          TobAna.PutValue('Y_POURCENTQTE2', Arrondi(100-PourcQte2, ADecimP) ) ;
          TobAna.PutValue('Y_QTE2',         TOBEcr.GetValue('E_QTE2') - TotQte2 ) ;
          end ;
        END ;
    END ;
END ;

Procedure deleteAna(TobL  : Tob ; i : Integer) ;
var SQL : String ;
BEGIN
If TOBL=Nil Then Exit ;
If TobL.GetValue('E_LIBRETEXTE0')='' Then Exit ;
If TobL.GetValue('E_LIBRETEXTE1')='X' Then Exit ;
SQL:='DELETE FROM ANALYTIQ WHERE '+
     ' Y_JOURNAL="'+TOBL.getValue('E_JOURNAL')+'" AND Y_EXERCICE="'+TOBL.getValue('E_EXERCICE')+'" '+
     ' AND Y_DATECOMPTABLE="'+UsDateTime(TOBL.getValue('E_DATECOMPTABLE'))+'" '+
     ' AND Y_NUMEROPIECE='+InttoStr(TOBL.getValue('E_NUMEROPIECE'))+' AND Y_QUALIFPIECE="'+TOBL.getValue('E_QUALIFPIECE')+'" '+
     ' AND Y_AXE="A'+IntToStr(i)+'" AND Y_NUMLIGNE='+InttoStr(TOBL.getValue('E_NUMLIGNE')) ;
ExecuteSQL(SQL) ;
END ;

Function VentilerTOBCreat ( TOBEcr : TOB ; NbDec : integer ; StAxe,Sect : String ; VerifGene : Boolean) : boolean ;
Var CptGene,Section : String ;
    Q          : TQuery ;
    Vent       : Array[1..5] of boolean ;
    Trouv      : Boolean ;
    NumVentil,NumAxe : integer ;
    Pourcentage        : double ;
    TOBAxe,TOBAna : TOB ;
    fb : tFichierBase ;
BEGIN
Result:=False ;
if TOBEcr=Nil then Exit ;
CptGene:=TOBEcr.GetValue('E_GENERAL') ; FillChar(Vent,Sizeof(Vent),#0) ; Trouv:=False ;
// Rechercher les axes ventilables du compte
If VerifGene Then
  BEGIN
  Q:=OpenSQL('Select G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 from GENERAUX Where G_GENERAL="'+CptGene+'"AND G_VENTILABLE="X"',True,-1,'',true) ;
  if Not Q.EOF then
     BEGIN
     Trouv:=True ;
     for NumAxe:=1 to 5 do Vent[NumAxe]:=(Q.Fields[NumAxe-1].AsString='X') ;
     END ;
  Ferme(Q) ;
  if Not Trouv then Exit ;
  END Else
  BEGIN
  If Length(StAxe)<5 Then Exit ;
  for NumAxe:=1 to 5 do Vent[NumAxe]:=(StAxe[NumAxe]='X') ;
  END ;
// Section d'attente
for NumAxe:=1 to 5 do if Vent[NumAxe] then
    BEGIN
    //fb:=AxeToFb('A'+IntToStr(NumAxe)) ;
    TOBAxe:=TOBEcr.Detail[NumAxe-1] ;
    if TOBAxe.Detail.Count<=0 then
       BEGIN
       DeleteAna(TobEcr,NumAxe) ;
       TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
       TobAna.InitValeurs ;
       EcrVersAna(TOBEcr,TOBAna) ;
       Section:=Trim(Sect) ;
       fb:=AxeToFb('A'+IntToStr(NumAxe)) ;
       If Section='' Then Section :=GetInfoCpta(fb).Attente ;
       NumVentil:=1 ; Pourcentage:=100.0 ;
       VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0) ;
       END ;
    END ;
ArrondirAnaTOB(TOBEcr,NbDec) ;
Result:=True ;
END ;

procedure CVentilerTOBTva( vTOBEcr : TOB ) ;
begin
 CVentilerAttente(vTOBEcr,StrToInt(Copy(GetParamSocSecur('SO_CPPCLAXETVA', '  '),2,1)) ) ;
end ;

procedure CVentilerAttente ( vTOBEcr : TOB ; vInAxe : integer ; vStSection : string = '' ) ;
var
 lTOBAxe          : TOB ;
 lTOBAna          : TOB ;
 lStSection       : string ;
 lInNumVentil     : integer ;
 lRdPourcentage   : double ;
 lPourcQte1       : double ;
 lPourcQte2       : double ;
begin

 if vTOBEcr.Detail.Count = 0 then exit ;

  lTOBAxe := vTOBEcr.Detail[ vInAxe - 1 ] ;

  if lTOBAxe.Detail.Count <= 0 then
   begin

     lTOBAna := TOB.Create('ANALYTIQ',lTOBAxe, -1) ;
     lTobAna.InitValeurs ;
     EcrVersAna(vTOBEcr,lTOBAna) ;
     if vStSection<>'' then
       begin
       lstSection := vStSection ;
       lTOBAna.PutValue('Y_QTE1',          vTOBEcr.GetValue('E_QTE1'));
       lTOBAna.PutValue('Y_QTE2',          vTOBEcr.GetValue('E_QTE2'));
       lTOBAna.PutValue('Y_POURCENTQTE1',  100.0 );
       lTOBAna.PutValue('Y_POURCENTQTE2',  100.0 );
       end
     else if ( vInAxe = 1 ) and  ( vTOBEcr.FieldExists('SECTION') ) then
      lStSection := vTOBEcr.GetValue('SECTION')
     else
      lStSection := GetInfoCpta(AxeToFb('A'+IntToStr(vInAxe))).Attente ;
     // 
     lInNumVentil   := 1 ;
     lRdPourcentage := 100.0 ;
     lPourcQte1     := 100.0 ;
     lPourcQte2     := 100.0 ;
     lTOBAna.PutValue('Y_TOTALQTE1',     vTOBEcr.GetValue('E_QTE1'));
     lTOBAna.PutValue('Y_TOTALQTE2',     vTOBEcr.GetValue('E_QTE2'));
     lTOBAna.PutValue('Y_QUALIFECRQTE1', vTOBEcr.GetValue('E_QUALIFQTE1'));
     lTOBAna.PutValue('Y_QUALIFECRQTE2', vTOBEcr.GetValue('E_QUALIFQTE2'));
     lTOBAna.PutValue('Y_QUALIFQTE1',    vTOBEcr.GetValue('E_QUALIFQTE1'));
     lTOBAna.PutValue('Y_QUALIFQTE2',    vTOBEcr.GetValue('E_QUALIFQTE2'));
     if GetParamSocSecur('SO_CPPCLSAISIETVA',false) then
      begin
       if ( vInAxe = StrToInt(Copy(GetParamSocSecur('SO_CPPCLAXETVA', '  '),2,1) ))  then
        lTOBAna.PutValue('Y_DATEREFEXTERNE',  vTobEcr.GetValue('E_DATECOMPTABLE')) ;
      end ;
     VentilLigneTOB(lTOBAna,lStSection,lInNumVentil,V_PGI.OkDecV,lRdPourcentage,vTOBEcr.GetValue('E_DEBIT')<>0, lPourcQte1, lPourcQte2 ) ;
   end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 01/01/1900
Modifié le ... : 22/05/2007
Description .. : 
Suite ........ : SBO 22/05/2007 : FQ 19906 Plantage requête quand 
Suite ........ : des restrictions existent
Mots clefs ... : 
*****************************************************************}
function CVentilerVentil ( vTOBEcr : TOB ; vInAxe : integer ; vNbDec : integer ; vStV_NATURE : string ; vStCodeVT : string ; vDossier : string ; vBoComplet : boolean ; vRestriction : TObject ) : boolean ;
var
 lTOBAxe          : TOB ;
 lTOBAna          : TOB ;
 lStCptGene       : string ;
 lStOrderBy       : string ;
 lStQrySansFiltre : string ;
 lStSQL           : string ;
 lStSection       : string ;
 lInNumVentil     : integer ;
 lQ               : TQuery ;
 lRdPourcentage   : double ;
 lRdPourcQte1     : double ;
 lRdPourcQte2     : double ;
 {$IFNDEF PGIIMMO}
 lRestriction     : TRestrictionAnalytique ;
 lCompteAna       : array[1..MaxAxe] of string;
 {$ENDIF}
 lTotalPourc      : double ;
begin

 result := false ;

 if ( vInAxe = 1 ) and ( vTOBEcr.FieldExists('SECTION') ) then exit ;

 if vTOBEcr.Detail.Count = 0 then exit ;


 {$IFNDEF PGIIMMO}
 lRestriction := TRestrictionAnalytique(vRestriction) ;
 {$ELSE}
 vRestriction := nil ;
 {$ENDIF}

 lTOBAxe := vTOBEcr.Detail[vInAxe-1] ;

 if lTOBAxe.Detail.Count > 0 then exit ;

 lStCptGene := vTOBEcr.GetValue('E_GENERAL') ;
 {$IFNDEF PGIIMMO}
 FillChar( lCompteAna , sizeof(lCompteAna), #0) ;
 {$ENDIF}

 {b FP 19/04/2006 FQ17725: Requête sans filtre sur la restriction analytique}
 lStQrySansFiltre := 'SELECT * FROM ' + GetTableDossier(vDossier, 'VENTIL') + ' WHERE '+ vStV_NATURE + IntToStr(vInAxe) + '" AND V_COMPTE="' + vStCodeVT + '"' ;
 lStSQL         := lStQrySansFiltre
 {b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
 {$IFNDEF PGIIMMO}
  + ' AND ' + lRestriction.GetClauseCompteAutorise( lStCptGene , 'A' + IntToStr( vInAxe) , 'VENTIL', lCompteAna )
 {$ENDIF} ;
 {e FP 29/12/2005}
 lStOrderBy       := ' ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
 lStQrySansFiltre := lStQrySansFiltre + lStOrderBy;
 lStSQL           := lStSQL           + lStOrderBy;
 lQ               := OpenSQL(lStSQL, True,-1,'',true) ;
 {e FP 19/04/2006}
 {$IFNDEF PGIIMMO}
 lRestriction.VerifModelVentil(lStCptGene, lQ, 'VENTIL', 'A' + IntToStr(vInAxe), lStQrySansFiltre); {FP 29/12/2005+19/04/2006 FQ17725}
 {$ENDIF}
 result      := false ;
 lTotalPourc := 0 ;
 while ( not lQ.EOF) do
  begin
   // FQ 17237 : On ne met pas en place les lignes dont le pourcentage n'est pas renseigné
   lRdPourcentage := lQ.FindField('V_TAUXMONTANT').AsFloat ;
   if (lRdPourcentage <> 0) or (not vBoComplet) then
    begin
     lTotalPourc  := lTotalPourc + lRdPourcentage ;
     lTOBAna      := TOB.Create('ANALYTIQ',lTOBAxe,-1) ;
     lTobAna.InitValeurs ;
     EcrVersAna(vTOBEcr,lTOBAna) ;
     lStSection   := lQ.FindField('V_SECTION').AsString ;
     lInNumVentil := lQ.FindField('V_NUMEROVENTIL').AsInteger ;
     // Données Qtes
     lRdPourcQte1 := lQ.FindField('V_TAUXQTE1').AsFloat ;
     lRdPourcQte2 := lQ.FindField('V_TAUXQTE2').AsFloat ;
     lTOBAna.PutValue('Y_TOTALQTE1',     vTobEcr.GetValue('E_QTE1'));
     lTOBAna.PutValue('Y_TOTALQTE2',     vTobEcr.GetValue('E_QTE2'));
     lTOBAna.PutValue('Y_QUALIFECRQTE1', vTobEcr.GetValue('E_QUALIFQTE1'));
     lTOBAna.PutValue('Y_QUALIFECRQTE2', vTobEcr.GetValue('E_QUALIFQTE2'));
     lTOBAna.PutValue('Y_QUALIFQTE1',    vTobEcr.GetValue('E_QUALIFQTE1'));
     lTOBAna.PutValue('Y_QUALIFQTE2',    vTobEcr.GetValue('E_QUALIFQTE2'));
     lTOBAna.PutValue('Y_DATEREFEXTERNE',vTobEcr.GetValue('E_DATECOMPTABLE'));

     if not GetParamSocDossierMS('SO_CROISAXE', False, vDossier) then
       VentilLigneTOB(lTOBAna,lStSection,lInNumVentil,vNbDec,lRdPourcentage, ( vTOBEcr.GetValue('E_DEBIT') <> 0 ), lRdPourcQte1, lRdPourcQte2)
     else
       {JP 24/10/07 : ajout des sous-plans pour le croise axe}
       VentilLigneTOB(lTOBAna,lStSection,lInNumVentil,vNbDec,lRdPourcentage, ( vTOBEcr.GetValue('E_DEBIT') <> 0 ), lRdPourcQte1, lRdPourcQte2,
                    lQ.FindField('V_SOUSPLAN1').AsString, lQ.FindField('V_SOUSPLAN2').AsString, lQ.FindField('V_SOUSPLAN3').AsString,
                    lQ.FindField('V_SOUSPLAN4').AsString, lQ.FindField('V_SOUSPLAN5').AsString) ;
    end ; // if
   lQ.Next ;
   result := true ;
  end ; // while

  Ferme(lQ) ;

  // Equilibrage des montants si le pourcentage est à 100% 
  if Arrondi( lTotalPourc-100.0, ADecimP ) = 0 then
    ArrondirAnaTOB( vTOBEcr, vNbDec, vInAxe ) ;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 01/01/1900
Modifié le ... : 31/03/2006
Description .. :
Suite ........ : 31/03/2006 : Ajout de nouveaux paramètres d'appel :
Suite ........ :   - vDossier : dossier de connexion pour le chargement des
Suite ........ : paramètres en multi-dossier,
Suite ........ :   - vBoVT : indicateur de mode ventilation type ( le CodeG
Suite ........ : indique alors la ventilation type ),
Suite ........ :   - vBoComplet : indique si la ventilation doit être complétée
Suite ........ : par la section d'attente,
Suite ........ : REM : ces paramètres ne sont pas effectifs en mode
Suite ........ : croise-axe
Mots clefs ... :
*****************************************************************}
Function VentilerTOB ( TOBEcr : TOB ; CodeG : String ; NumLig,NbDec : integer ; QueAttente : boolean ; vDossier : String = ''; vBoComplet : Boolean = True ; vStVT : String = '' ) : boolean ;
Var CptGene,Ax,Section : String ;
    Q          : TQuery ;
    Vent       : Array[1..5] of boolean ;
    Trouv      : Boolean ;
    NumVentil,NumAxe : integer ;
    Pourcentage        : double ;
    TOBAxe,TOBAna : TOB ;
    {b FP 29/12/2005}
 {$IFNDEF PGIIMMO}
    Restriction:    TRestrictionAnalytique;        // Modèlde de restriction ana
 {$ENDIF}
    CompteAna:   array[1..MaxAxe] of String;
    lStV_NATURE : string ;
    {e FP 29/12/2005}
    {b FP 19/04/2006 FQ17725}
    SQL:           String;
    OrderBy:       String;
    QrySansFiltre: String;
    {e FP 19/04/2006}
    lStCodeVT   : string ;
    lPourcQte1  : double ;
    lPourcQte2  : double ;
BEGIN

Result:=False ;
if TOBEcr=Nil then Exit ;

//SG6 28.01.05 Gestion mode analytique croisaxe
  if GetParamSocSecur('SO_CROISAXE', False) then
  begin
    result := VentilerTOBCroisaxe(TobEcr,CodeG,NumLig,NbDec,QueAttente, vBoComplet);  {FP 12/06/2006 FQ18226}
    exit;
  end;

  // Compte général
  CptGene := TOBEcr.GetValue('E_GENERAL') ;

  // mode ventilation-type
  if vStVT <> '' then
    begin
    lStV_NATURE := 'V_NATURE="TY' ;
    lStCodeVT   := vStVT ;
    end
  // mode standard ==> ventilation par défaut des comptes généraux
  else begin
       lStV_NATURE := 'V_NATURE="GE' ;
       lStCodeVT   := CptGene ;
       end ;

  // Rechercher les axes ventilables du compte
  Trouv   := False ;
  FillChar(Vent,Sizeof(Vent),#0) ;
  Q:=OpenSelect('Select G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 from GENERAUX Where G_GENERAL="'+CptGene+'"AND G_VENTILABLE="X"',vDossier) ;
  if Not Q.EOF then
    BEGIN
    Trouv:=True ;
    for NumAxe:=1 to 5 do Vent[NumAxe]:=(Q.Fields[NumAxe-1].AsString='X') ;
    END ;
  Ferme(Q) ;
  if Not Trouv then Exit ;

{b FP 29/12/2005}
FillChar(CompteAna, sizeof(CompteAna), #0);
{$IFNDEF PGIIMMO}
Restriction := TRestrictionAnalytique.Create;
{$ENDIF}
{e FP 29/12/2005}

// Ventilation à partir du guide analytique
//Trouv:=False ;
Trouv := not vBoComplet ;
if ((CodeG<>'') and (not QueAttente)) and (Length(Trim(CodeG))<=3){FP 29/12/2005: Le code est sur 3 caractères} then
   BEGIN
   {b FP 29/12/2005: il faut lancer cette requête par axe}
   for NumAxe:=1 to 5 do
     begin
     {b FP 19/04/2006 FQ17725: Requête sans filtre sur la restriction analytique}
     QrySansFiltre := 'Select * from ANAGUI Where AG_TYPE="NOR" AND AG_GUIDE="'+CodeG+'"'
               +' AND AG_NUMLIGNE='+IntToStr(NumLig)
               +' AND AG_AXE="A'+IntToStr(NumAxe)+'"';
     SQL := QrySansFiltre
               {b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
               {$IFNDEF PGIIMMO}
               +' AND '  + Restriction.GetClauseCompteAutorise(
                           CptGene, 'A'+IntToStr(NumAxe), 'ANAGUI', CompteAna)
               {$ENDIF};
               {e FP 29/12/2005}
     OrderBy := ' Order By AG_AXE, AG_NUMVENTIL';
     QrySansFiltre := QrySansFiltre + OrderBy;
     SQL           := SQL           + OrderBy;
     Q := OpenSelect(SQL, vDossier) ;
     {e FP 19/04/2006}
     {$IFNDEF PGIIMMO}
     Restriction.VerifModelVentil(CptGene, Q, 'ANAGUI', 'A'+IntToStr(NumAxe), QrySansFiltre); {FP 29/12/2005+19/04/2006 FQ17725}
     {$ENDIF}
     While (Not Q.EOF) do
        BEGIN
        Ax:=Q.FindField('AG_AXE').AsString ; {FP 29/12/2005 NumAxe:=Ord(Ax[2])-48 ;}
        TOBAxe:=TOBEcr.Detail[NumAxe-1] ;
        TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
        TobAna.InitValeurs ;
        EcrVersAna(TOBEcr,TOBAna) ;
        Section     := Q.FindField('AG_SECTION').AsString ;
        NumVentil   := Q.FindField('AG_NUMVENTIL').AsInteger ;
        Pourcentage := Valeur(Q.FindField('AG_POURCENTAGE').AsString) ;
        // Données Qtes
        lPourcQte1  := Valeur(Q.FindField('AG_POURCENTQTE1').AsString) ;
        lPourcQte2  := Valeur(Q.FindField('AG_POURCENTQTE2').AsString) ;
        TOBAna.PutValue('Y_TOTALQTE1',     TobEcr.GetValue('E_QTE1'));
        TOBAna.PutValue('Y_TOTALQTE2',     TobEcr.GetValue('E_QTE2'));
        TOBAna.PutValue('Y_QUALIFECRQTE1', TobEcr.GetValue('E_QUALIFQTE1'));
        TOBAna.PutValue('Y_QUALIFECRQTE2', TobEcr.GetValue('E_QUALIFQTE2'));
        TOBAna.PutValue('Y_QUALIFQTE1',    TobEcr.GetValue('E_QUALIFQTE1'));
        TOBAna.PutValue('Y_QUALIFQTE2',    TobEcr.GetValue('E_QUALIFQTE2'));
        TOBAna.PutValue('Y_DATEREFEXTERNE',TobEcr.GetValue('E_DATECOMPTABLE'));
        VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,lPourcQte1,lPourcQte2) ;
        Trouv:=True ;
        Q.Next ;
        END ;
     Ferme(Q) ;
     END ;
   {e FP 29/12/2005}
   end;
// Préventilation du compte
if ((Not Trouv) or (CodeG='')) then
 if Not QueAttente then
  for NumAxe:=1 to 5 do if Vent[NumAxe] then
    begin
   {$IFNDEF PGIIMMO}
    CVentilerVentil ( TOBEcr,NumAxe,NbDec,lStV_NATURE,lStCodeVT,vDossier,vBoComplet,Restriction ) ;
   {$ELSE}
    CVentilerVentil ( TOBEcr,NumAxe,NbDec,lStV_NATURE,lStCodeVT,vDossier,vBoComplet,nil ) ;
    {$ENDIF}
    end ;
// Section d'attente
if vBoComplet or QueAttente then
  begin
  for NumAxe:=1 to 5 do if Vent[NumAxe] then
   CVentilerAttente(TOBEcr,NumAxe) ;
  ArrondirAnaTOB(TOBEcr,NbDec) ;
  end ;

Result:=True ;
{$IFNDEF PGIIMMO}
FreeAndNil(Restriction);             {FP 29/12/2005}
{$ENDIF}

END ;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 28/01/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet la ventilation en mode croisaxe
Mots clefs ... : VENTILATION CROISAXE
*****************************************************************}
Function VentilerTOBCroisaxe ( TOBEcr : TOB ; CodeG : String ; NumLig,NbDec : integer ; QueAttente : boolean; vBoComplet : Boolean = True) : boolean ;  {FP 12/06/2006 FQ18226 Ajout paramètre vBoComplet}
Var CptGene,Section : String ;
    Q          : TQuery ;
    Axes       : Array[1..MaxAxe] of boolean ;
    SecAttente : Array[1..MaxAxe] of string ;
    Trouv      : Boolean ;
    fb         : TFichierBase ;
    NumVentil  : integer ;
    Pourcentage        : double ;
    TOBAxe,TOBAna : TOB ;
    premier_axe : integer;
    lInCpt : Integer ;
    AVentilerAvecAttente : boolean;
   {b FP 29/12/2005}
{$IFNDEF PGIIMMO}
    Restriction:    TRestrictionAnalytique;        // Modèlde de restriction ana
{$ENDIF}
    CompteAna:      array[1..MaxAxe] of String;
   {e FP 29/12/2005}
   {b FP 19/04/2006 FQ17725}
   SQL:             String;
   QrySansFiltre:   String;
   OrderBy:         String;
   {e b FP 19/04/2006}
begin
  Result := False ;
  if TOBEcr = Nil then Exit ;

  AVentilerAvecAttente := True;

  CptGene := TOBEcr.GetValue('E_GENERAL') ;
  Trouv := False  ;

  // On regarde si le compte est ventilable
  Q:=OpenSQL('Select * from GENERAUX Where G_GENERAL="'+CptGene+'"AND G_VENTILABLE="X"',True,-1,'',true) ;
  if Not Q.EOF then
    begin
      Trouv := True ;
    end ;
  Ferme(Q) ;
  if Not Trouv then Exit ;

  //On recherche les axe ventilables et le premier axe ventilé
  FillChar(SecAttente,Sizeof(SecAttente),#0);
  premier_axe := 0 ;
  for lInCpt := 1 to MaxAxe do
  begin
    Axes[lInCpt] := GetParamSocSecur('SO_VENTILA' + IntToStr(lInCpt), False );
    if Axes[lInCpt] then
    begin
      if premier_axe = 0 then premier_axe := lInCpt;
      SecAttente[lInCpt] := GetInfoCpta (AxeToFb('A'+IntToStr(lInCpt)) ).Attente ;
    end
    else
    begin
      SecAttente[lInCpt] := '';
    end;
  end;

  {b FP 29/12/2005}
  FillChar(CompteAna, sizeof(CompteAna), #0);
  {$IFNDEF PGIIMMO}
  Restriction := TRestrictionAnalytique.Create;
  {$ENDIF}
  {e FP 29/12/2005}
  //Guide
  //A faire a mode croisaxe


  //Ventilation types par rapport au compte
  if (CptGene <> '') and Not(QueAttente) then
  begin
    TOBAxe:=TOBEcr.Detail[premier_axe-1] ;
    if TOBAxe.Detail.Count<=0 then
    begin
      {b FP 19/04/2006 FQ17725: Requête sans filtre sur la restriction analytique}
      QrySansFiltre := 'SELECT * FROM VENTIL WHERE V_NATURE="GE'+IntToStr(premier_axe)+'" AND V_COMPTE="'+CptGene+'"';
      SQL := QrySansFiltre
       {b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
       {$IFNDEF PGIIMMO}
       +' AND ' + Restriction.GetClauseCompteAutorise(
                    CptGene, 'A'+IntToStr(premier_axe), 'VENTIL', CompteAna)
       {$ENDIF};
       {e FP 29/12/2005}
      OrderBy := ' ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL';
      QrySansFiltre := QrySansFiltre + OrderBy;
      SQL           := SQL           + OrderBy;
      Q:=OpenSQL(SQL, True,-1,'',true) ;
      {e FP 19/04/2006}
      {$IFNDEF PGIIMMO}
      Restriction.VerifModelVentil(CptGene, Q, 'VENTIL', 'A'+IntToStr(premier_axe), QrySansFiltre);      {FP 29/12/2005+19/04/2006 FQ17725}
      {$ENDIF}
      AVentilerAvecAttente := QueAttente;    {FP 21/06/2006 FQ18421}
      while (Not Q.EOF) do
      begin
        TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
        TobAna.InitValeurs ;
        EcrVersAna(TOBEcr,TOBAna) ;
        Section:=Q.FindField('V_SECTION').AsString ;
        NumVentil:=Q.FindField('V_NUMEROVENTIL').AsInteger ;
        Pourcentage:=Q.FindField('V_TAUXMONTANT').AsFloat ;
        TOBAna.PutValue('Y_TOTALQTE1',TobEcr.GetValue('E_QTE1'));
        TOBAna.PutValue('Y_TOTALQTE2',TobEcr.GetValue('E_QTE2'));
        VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Q.FindField('V_TAUXQTE1').AsFloat,Q.FindField('V_TAUXQTE2').AsFloat,
             Q.FindField('V_SOUSPLAN1').AsString,Q.FindField('V_SOUSPLAN2').AsString,Q.FindField('V_SOUSPLAN3').AsString,Q.FindField('V_SOUSPLAN4').AsString,
             Q.FindField('V_SOUSPLAN5').AsString) ;
        Q.Next ;
        AVentilerAvecAttente := False;
      end ;
      Ferme(Q) ;
    end;
  end;

  //Section d'attente
  if AVentilerAvecAttente or vBoComplet then     {FP 12/06/2006 FQ18226}
  begin

    fb := AxeToFb('A'+IntToStr(premier_axe)) ;
    TOBAxe := TOBEcr.Detail[premier_axe-1] ;

    if TOBAxe.Detail.Count<=0 then
    begin
         TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
         TobAna.InitValeurs ;
         EcrVersAna(TOBEcr,TOBAna) ;

         if ((TOBEcr.FieldExists('SECTION'))) then Section:=TOBEcr.GetValue('SECTION')
                                              else Section := GetInfoCpta(fb).Attente ;
         NumVentil:=1 ; Pourcentage:=100.0 ;

         VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,0.0,0.0,SecAttente[1],SecAttente[2],SecAttente[3],SecAttente[4],SecAttente[5]) ;
    end ;
  ArrondirAnaTOB(TOBEcr,NbDec) ;      {FP 12/06/2006 FQ18226 A faire uniquement si SectionAttente comme pour les axes indépendants}
  end;


  {$IFNDEF PGIIMMO}
  FreeAndNil(Restriction);             {FP 29/12/2005}
  {$ENDIF}
  Result:=True ;

end;


Procedure ProraterVentilsTOB ( TOBEcr : TOB ; NbDec : integer ) ;
Var TOBAxe,TOBAna : TOB ;
    NumAxe,i      : integer ;
    OkDebit       : boolean ;
    TotEcrP,TotEcrD,Pourc : Double ;
    TotP,TotD : Double ;
BEGIN
if TOBEcr=Nil then Exit ;
if TOBEcr.GetValue('E_ANA')<>'X' then Exit ;
if TOBEcr.Detail.Count<5 then Exit ;
OkDebit:=(TOBEcr.GetValue('E_DEBIT')<>0) ;
TotEcrP:=Arrondi(TOBEcr.GetValue('E_DEBIT')+TOBEcr.GetValue('E_CREDIT'),V_PGI.OkDecV) ;
TotEcrD:=Arrondi(TOBEcr.GetValue('E_DEBITDEV')+TOBEcr.GetValue('E_CREDITDEV'),NbDec) ;
for NumAxe:=1 to 5 do
    BEGIN
    TOBAxe:=TOBEcr.Detail[NumAxe-1] ; TotP:=0 ; TotD:=0 ; 
    for i:=0 to TOBAxe.Detail.Count-1 do
        BEGIN
        TOBAna:=TOBAxe.Detail[i] ;
        TOBAna.PutValue('Y_TOTALECRITURE',TotEcrP); TOBAna.PutValue('Y_TOTALDEVISE',TotEcrD) ;
        if i=TOBAxe.Detail.Count-1 then
           BEGIN
           // Dernière ligne --> Montant = le reste à ventiler
           if OkDebit then
              BEGIN
              TOBAna.PutValue('Y_DEBIT',Arrondi(TotEcrP-TotP,V_PGI.OkDecV)) ;
              TOBAna.PutValue('Y_DEBITDEV',Arrondi(TotEcrD-TotD,NbDec)) ;
              END else
              BEGIN
              TOBAna.PutValue('Y_CREDIT',Arrondi(TotEcrP-TotP,V_PGI.OkDecV)) ;
              TOBAna.PutValue('Y_CREDITDEV',Arrondi(TotEcrD-TotD,NbDec)) ;
              END ;
           END else
           BEGIN
           // Ligne courante --> Montant = % x Montant de l'écriture
           Pourc:=TOBAna.GetValue('Y_POURCENTAGE')/100.0 ;
           if OkDebit then
              BEGIN
              TOBAna.PutValue('Y_DEBIT',Arrondi(TotEcrP*Pourc,V_PGI.OkDecV)) ;
              TOBAna.PutValue('Y_DEBITDEV',Arrondi(TotEcrD*Pourc,NbDec)) ;
              END else
              BEGIN
              TOBAna.PutValue('Y_CREDIT',Arrondi(TotEcrP*Pourc,V_PGI.OkDecV)) ;
              TOBAna.PutValue('Y_CREDITDEV',Arrondi(TotEcrD*Pourc,NbDec)) ;
              END ;
           END ;
        // Cumuls au fur et à mesure
        if i<TOBAxe.Detail.Count-1 then
           BEGIN
           TotP:=Arrondi(TotP+TOBAna.GetValue('Y_DEBIT')+TOBAna.GetValue('Y_CREDIT'),V_PGI.OkDecV) ;
           TotD:=Arrondi(TotD+TOBAna.GetValue('Y_DEBITDEV')+TOBAna.GetValue('Y_CREDITDEV'),NbDec) ;
           END ;
        END ;
    END ;
END ;

Function StrS( X : Double ; DD : integer ) : String ;
BEGIN Result:=StrfMontant(X,15,DD,'',TRUE) ; END ;

Function StrF00 ( X : Double ; DD : integer ) : String ;
BEGIN if X=0 then Result:='' else Result:=StrfMontant(X,15,DD,'',TRUE) ; END ;

Function StrSP( X : Double ) : String ;
BEGIN Result:=StrfMontant(X,15,V_PGI.OkDecV,'',TRUE) ; END ;

Function StrS0 ( X : Double ) : String ;
BEGIN if X=0 then Result:='' else Result:=StrfMontant(X,15,V_PGI.OkDecV,'',TRUE) ; END ;

Function StrP0 ( X : Double ) : String ;
begin if X=0 then Result:='' else Result:=StrfMontant(X,7,V_PGI.OkDecV,'',FALSE)+' %' ; end ;

Function CaseFicDataType ( Fic : String ) : TFichierBase ;
begin
Result:=fbGene ;
if Copy(Fic,1,3)='TZG'       then Result:=fbGene else
if Copy(Fic,1,3)='TZT'       then Result:=fbAux else
if Copy(Fic,1,3)='TZJ'       then Result:=fbJal else
if Copy(Fic,1,9)='TZCORRESP' then Result:=fbCorresp else
if Copy(Fic,1,5)='TZNAT'     then Result:=fbNatCpt else
if Fic='TZIMMO'              then Result:=fbImmo else
if Fic='TZBUDJAL'            then Result:=fbBudJal else
if Fic='TZSECTION'           then Result:=fbAxe1 else
if Fic='TZSECTION2'          then Result:=fbAxe2 else
if Fic='TZSECTION3'          then Result:=fbAxe3 else
if Fic='TZSECTION4'          then Result:=fbAxe4 else
if Fic='TZSECTION5'          then Result:=fbAxe5 else
if Fic='TZCSECTION'           then Result:=fbAxe1 else
if Fic='TZCSECTION2'          then Result:=fbAxe2 else
if Fic='TZCSECTION3'          then Result:=fbAxe3 else
if Fic='TZCSECTION4'          then Result:=fbAxe4 else
if Fic='TZCSECTION5'          then Result:=fbAxe5 else
if ((Fic='TZBUDGEN') or (Fic='TZBUDGENATT')) then Result:=fbBudGen else
if ((Fic='TZBUDSEC1') or (Fic='TZBUDSECATT1')) then Result:=fbBudSec1 else
if ((Fic='TZBUDSEC2') or (Fic='TZBUDSECATT2')) then Result:=fbBudSec2 else
if ((Fic='TZBUDSEC3') or (Fic='TZBUDSECATT3')) then Result:=fbBudSec3 else
if ((Fic='TZBUDSEC4') or (Fic='TZBUDSECATT4')) then Result:=fbBudSec4 else
if ((Fic='TZBUDSEC5') or (Fic='TZBUDSECATT5')) then Result:=fbBudSec5 ;
end ;

Function GetRIBPrincipal ( Cpt : String ) : String ;
Var Q : TQuery ;
    St : String ;
BEGIN
Result:='' ; St:='' ;
if Cpt='' then Exit ;
Q:=OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Cpt+'" AND R_PRINCIPAL="X"',True,-1,'',true) ;
if Not Q.EOF then
   BEGIN
   // Pays = FRANCE ou VIDE (Lors d'importation,le pays peut ne pas être renseigné)
   //XMG 20/04/04 début
   {$IFDEF ESP}
   St:=IBANtoE_RIB(Q.FindField('R_CODEIBAN').AsString) ;
   {$ELSE}
   //RR 19/02/2007
    //if ((codeisodupays(Q.FindField('R_PAYS').AsString) <> 'FR') and (Q.FindField('R_PAYS').AsString <> '')) then
    if ((codeisodupays(Q.FindField('R_PAYS').AsString) <> QuelPaysLocalisation ) and (Q.FindField('R_PAYS').AsString <> '')) then
     St:='*'+Q.FindField('R_CODEIBAN').AsString
   else
     St:=EncodeRIB(Q.FindField('R_ETABBQ').AsString,
                   Q.FindField('R_GUICHET').AsString,
                   Q.FindField('R_NUMEROCOMPTE').AsString,
                   Q.FindField('R_CLERIB').AsString,
                   Q.FindField('R_DOMICILIATION').AsString,
                   CodeIsoDuPays(Q.FindField('R_PAYS').AsString),
                   Q.FindField('R_TYPEPAYS').AsString) ;
   {$ENDIF ESP}
   //XMG 20/04/04 fin
   END
    else St:='' ;
Ferme(Q) ;
Result:=St ;
END ;

Function GetRIBParticulier ( Cpt : String ; No : integer) : String ;
Var Q : TQuery ;
    St : String ;
BEGIN
Result:='' ; St:='' ;
if Cpt='' then Exit ;
Q:=OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Cpt+'" AND R_NUMERORIB='+InttOStr(No),True,-1,'',true) ;
if Not Q.EOF then
   BEGIN
   // Pays = FRANCE ou VIDE (Lors d'importation,le pays peut ne pas être renseigné)
   //XMG 20/04/04 début
   {$IFDEF ESP}
   St:=IBANtoE_RIB(Q.FindField('R_CODEIBAN').AsString) ;
   {$ELSE}
   //RR 19/02/2007
    //if ((codeisodupays(Q.FindField('R_PAYS').AsString) <> 'FR') and (Q.FindField('R_PAYS').AsString <> '')) then
    if ((codeisodupays(Q.FindField('R_PAYS').AsString) <> QuelPaysLocalisation ) and (Q.FindField('R_PAYS').AsString <> '')) then
     St:='*'+Q.FindField('R_CODEIBAN').AsString
   else
     St:=EncodeRIB(Q.FindField('R_ETABBQ').AsString,
                   Q.FindField('R_GUICHET').AsString,
                   Q.FindField('R_NUMEROCOMPTE').AsString,
                   Q.FindField('R_CLERIB').AsString,
                   Q.FindField('R_DOMICILIATION').AsString,
                   CodeIsoDuPays(Q.FindField('R_PAYS').AsString),
                   Q.FindField('R_TYPEPAYS').AsString) ;
   {$ENDIF ESP}
   //XMG 20/04/04 fin
   END
    else St:='' ;
Ferme(Q) ;
Result:=St ;
END ;


{$IFNDEF NOVH}

procedure TOBM.SetMPACC;
begin
 if Ident = EcrGen then
  PutMvt ( 'E_CODEACCEPT', MPTOACC ( GetMvt ( 'E_MODEPAIE' ) ) ) ;
end;


Function MPTOACC ( MP : String ) : String ;
Var i : integer ;
    St,MPLu,Acc : String ;
BEGIN
Result:='NON' ;
{$IFNDEF SPEC302}
for i:=0 to VH^.MPACC.Count-1 do
    BEGIN
    St:=VH^.MPACC[i] ; MPLu:=ReadtokenSt(St) ;
    if MPLu=MP then BEGIN Acc:=ReadTokenSt(St) ; Result:=Acc ; Break ; END ;
    END ;
{$ENDIF}
END ;


Function MPTOCATEGORIE ( MP : String ) : String ;
Var i : integer ;
    St,MPLu,Acc,CatLu : String ;
BEGIN
Result:='NON' ;
{$IFNDEF SPEC302}
for i:=0 to VH^.MPACC.Count-1 do
    BEGIN
    St:=VH^.MPACC[i] ; MPLu:=ReadtokenSt(St) ; Acc:=ReadtokenSt(St) ;
    if MPLu=MP then BEGIN CatLu:=ReadTokenSt(St) ; Result:=CatLu ; Break ; END ;
    END ;
{$ENDIF}
END ;

{$ENDIF}

Function isEncMP ( smp : TSuiviMP ) : Boolean ;
BEGIN
Result:=smp in [smpEncPreBqe,smpEncTraEdt,smpEncTraEdtNC,smpEncTraEnc,smpEncTraEsc,
                smpEncTraBqe,smpEncChqBqe,smpEncCBBqe,smpEncTraPor,smpEncPreEdt,smpEncChqPor,smpEncCBPor,smpEncDiv,smpEncTous,smpEncPreEdtNC] ;
END ;

Function GetNumLotTraChq ( STra : String ) : String ;
Var i : integer ;
    NL,sNum : String ;
BEGIN
sNum:='' ; NL:=sTra ;
for i:=Length(NL) downto 1 do
    BEGIN
    if NL[i] in ['0'..'9'] then SNum:=NL[i]+SNum else Break ;
    END ;
Result:=sNum ;
END ;

Function IncNumLotTraChq ( STra : String ) : String ;
Var sNum,sf,NL : String ;
   ln,i,LastI : integer ;
BEGIN
NL:=sTra ; SNum:=GetNumLotTraChq(sTra) ; LastI:=Length(NL)-Length(sNum) ;
if SNum<>'' then
   BEGIN
   ln:=Length(SNum) ; SNum:=IntToStr(ValeurI(SNum)+1) ;
   if Length(SNum)<=ln then
      BEGIN
      sf:='' ; for i:=1 to ln do sf:=sf+'0' ;
      SNum:=FormatFloat(sf,ValeurI(sNum)) ;
      END ;
   NL:=Copy(NL,1,LastI)+SNum ;
   END ;
Result:=NL ;
END ;





Function NormaliserPieceSimu ( TOBL : TOB ; MajTOBL : Boolean ) : boolean ;
Var XX : RMVT ;
    TOBEcr,TOBE,TOBA : TOB ;
    QQ      : TQuery ;
    Jal,SQL : String ;
    NewNum,NumL : integer ;
    i,k     : integer ;
    DD      : TDateTime ;
    XXEcr   : RMVT ;  // Pour conserver le RMVT de l'écriture générale
{$IFDEF EAGLCLIENT}
    nbEnreg : LongInt ;
{$ENDIF}
BEGIN
Result:=False ;
if TOBL=Nil then Exit ;
XXEcr:=TOBToIdent(TOBL,False) ;
TOBEcr:=TOB.Create('',Nil,-1) ;
QQ:=OpenSQL('SELECT * FROM ECRITURE WHERE '+WhereEcriture(tsGene,XX,False),True,-1,'',true) ;
TOBEcr.LoadDetailDB('ECRITURE','','',QQ,False) ;
Ferme(QQ) ;
if TOBEcr.Detail.Count<=0 then BEGIN TOBEcr.Free ; Exit ; END ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBE:=TOBEcr.Detail[i] ;
    XX:=TOBToIdent(TOBE,True) ; NumL:=TOBE.GetValue('E_NUMLIGNE') ;
    SQL:='SELECT * FROM ANALYTIQ WHERE '+WhereEcriture(tsAnal,XX,True)+' AND Y_NUMLIGNE='+IntToStr(NumL) ;
    QQ:=OpenSQL(SQL,True,-1,'',true) ;
    TOBE.LoadDetailDB('ANALYTIQ','','',QQ,False) ;
    Ferme(QQ) ;
    END ;
Jal:=TOBEcr.Detail[0].GetValue('E_JOURNAL') ;
DD:=TOBEcr.Detail[0].GetValue('E_DATECOMPTABLE') ;
NewNum:=GetNewNumJal(Jal,True,DD) ; if NewNum<=0 then BEGIN TOBEcr.Free ; Exit ; END ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBE:=TOBEcr.Detail[i] ;
    TOBE.PutValue('E_QUALIFPIECE','N') ;
    TOBE.PutValue('E_NUMEROPIECE',NewNum) ;
    for k:=0 to TOBE.Detail.Count-1 do
        BEGIN
        TOBA:=TOBE.Detail[k] ;
        TOBA.PutValue('Y_QUALIFPIECE','N') ;
        TOBA.PutValue('Y_NUMEROPIECE',NewNum) ;
        END ;
    END ;
MajSoldesEcritureTOB(TOBEcr,True) ;

// Pb en eAGL, les lignes ne sont pas mise à jour // Fiche 12586 SBO
{$IFDEF EAGLCLIENT}
  // --> updateDB remplacer par executeSQL( UPDATE... )
  nbEnreg := ExecuteSQL ( 'UPDATE ECRITURE SET E_QUALIFPIECE="N", E_NUMEROPIECE=' + IntToStr(NewNum)
                         + ' WHERE ' + WhereEcriture(tsGene,XXEcr,False) ) ;
  Result := nbEnreg > 0 ;
  if Result then
    ExecuteSQL ( 'UPDATE ANALYTIQ SET Y_QUALIFPIECE="N", Y_NUMEROPIECE=' + IntToStr(NewNum)
                +' WHERE ' + WhereEcriture(tsAnal,XXEcr,False) ) ;
{$ELSE}
  if TOBEcr.UpdateDB(False,False) then
    Result:=True ;
{$ENDIF}

TOBEcr.Free ;
if MajTOBL then
   BEGIN
   TOBL.PutValue('E_NUMEROPIECE',NewNum) ;
   TOBL.PutValue('E_QUALIFPIECE','N') ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 12/09/2002
Modifié le ... :   /  /    
Description .. : - LG - 12/09/2002 - ajout du champ Supp _ETAT
Suite ........ : contenant l'etat du TOBM
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
Function AGLMPToCategorie ( Parms : array of variant ; nb : integer ) : Variant ;
BEGIN
Result:=MPToCategorie(Parms[0]) ;
END ;
{$ENDIF}

//  LG
// modif du creat pour ajout du nom de table
constructor TOBM.Create ( iecr : TTypeEcr; Axe : String3; Defaut : Boolean ; LeParent : TOB = nil ) ;
begin
 try
 case iecr of
  EcrGen     : inherited Create( 'ECRITURE'  , LeParent , -1 );
  EcrAna     : inherited Create( 'ANALYTIQ'  , LeParent , -1 );
  EcrGui     : inherited Create( 'ECRGUI'    , LeParent , -1 );
  EcrSais    : inherited Create( 'SUIVCPTA'  , LeParent , -1 );
  EcrAbo     : inherited Create( 'CONTABON'  , LeParent , -1 );
  EcrBud     : inherited Create( 'BUDECR'    , LeParent , -1 );
  else If V_PGI.SAV Then MessageAlerte('Type d''objet inconnue numéro '+intToStr(integer(iecr)));
 end; // case

 Ident       := iecr;
 Status      := EnMemoire;
 CompS       := False;
 CompE       := false ;

 LC          := HTStringList.Create;
 FM          := HTStringList.Create;
 // ajout du champs contenant l'etat de l'obm
 // fait avant le PutDefaut car celui ci fait des PutValues qui affecte l'etat de l'OBM
 AddChampSupValeur ('_ETAT' , integer(Inchange) );
 if Defaut then
  PutDefaut ( Axe ) ;

 {Champs particuliers}
 AddChampSupValeur ('OLDDEBIT'  , 0 );
 AddChampSupValeur ('OLDCREDIT' , 0 );
 AddChampSupValeur ('RATIO' , 0 );
 AddChampSupValeur ('CONVERTFRANC' , 0 );
 AddChampSupValeur ('CONVERTEURO' , 0 );

 NbC         := NbChamps + 1;

 except
  On E:Exception do
   begin
    If V_PGI.SAV Then MessageAlerte('Erreur dans la creation de l''objet' + #10#13 + E.Message );
    raise;
   end;
 end;
end;

destructor TOBM.Destroy;
begin
 FM.Clear;
 FM.Free;
 LC.Clear;
 LC.Free;
 inherited Destroy;
end;


procedure TOBM.ChargeMvt ( F : TDataSet ) ;
var
 i : Integer;
begin

 {$IFNDEF EAGLCLIENT}
 if ( not F.Active ) then exit ; //or F.EOF then
 {$ENDIF}

 if ( F is TQuery ) then
  SelectDB( '', ( F as TQuery ) )
   else
    for i := 0 to F.FieldCount - 1 do
     if ( ( F.Fields[i].DataType <> ftMemo ) and ( F.Fields[i].DataType <> ftBlob ) ) then
      PutValue( F.Fields[i].FieldName , F.Fields[i].AsVariant );

 // on remet la ligne à l'etat inchangé
 Etat := Inchange;

end;

procedure TOBM.ChargeMvtP( F : TDataSet ; vStPrefixe : string = '' ) ;
var
 i,lg    : integer;
 lStName : string;
begin
 if vStPrefixe='' then vStPrefixe:=TableToPrefixe(NomTable) ;
 for i := 0 to F.FieldCount - 1 do
  begin
   lStName := F.Fields[i].fieldName;
   Lg:=Length(UpperCase(vStPrefixe)+'_') ;
//    if UpperCase( Copy( lStName , 1 , 2 ) ) = UpperCase(vStPrefixe) + '_' then GP le 16/05/2002
    if UpperCase( Copy( lStName , 1 , Lg ) ) = UpperCase(vStPrefixe) + '_' then
     PutValue( lStName , F.Fields[i].AsVariant );
  end; // for
end;


function TOBM.GetMvt ( Nom : string ) : Variant;
begin
 result := GetValue( Nom );
end;

{=====================================================================}

procedure TOBM.PutMvt ( Nom : string; Value : Variant ) ;
begin
 if Value <> GetValue(Nom) then
  begin
   Etat := TEtat(1); // passage a l'etat du TOBM à l'etat modifié
   PutValue ( Nom , Value );
  end; // if
end;

{=====================================================================}

{***********A.G.L.***********************************************
Auteur  ...... : LG* Laurent GENDREAU
Créé le ...... : 12/09/2002
Modifié le ... :   /  /    
Description .. : - 12/09/2002 - ajout de l'initialisation des champs pour la 
Suite ........ : com
Mots clefs ... : 
*****************************************************************}
procedure TOBM.PutDefautEcr;
begin
{$IFNDEF SANSCOMPTA}
// PutMvt ( 'E_EXERCICE'          , VH^.Entree.Code ) ;
 PutMvt ( 'E_ETABLISSEMENT'     , GetParamSocSecur('SO_ETABLISDEFAUT', '') ) ;
 PutMvt ( 'E_PERIODE'           , GetPeriode ( V_PGI.DateEntree ) ) ;
 PutMvt ( 'E_SEMAINE'           , NumSemaine ( V_PGI.DateEntree ) ) ;
{$ENDIF}
 PutMvt ( 'E_DATECOMPTABLE'     , V_PGI.DateEntree ) ;
 PutMvt ( 'E_NATUREPIECE'       , 'OD' ) ;
 PutMvt ( 'E_QUALIFPIECE'       , 'N' ) ;
 PutMvt ( 'E_TYPEMVT'           , 'DIV' ) ;
 PutMvt ( 'E_VALIDE'            , '-' ) ;
 PutMvt ( 'E_ETAT'              , '0000000000' ) ;
 PutMvt ( 'E_UTILISATEUR'       , V_PGI.User ) ;
 PutMvt ( 'E_DATECREATION'      , Date ) ;
 PutMvt ( 'E_DATEMODIF'         , NowH ) ;
 PutMvt ( 'E_SOCIETE'           , V_PGI.CodeSociete ) ;
 PutMvt ( 'E_VISION'            , 'DEM' ) ;
 PutMvt ( 'E_TVAENCAISSEMENT'   , '-' ) ;
 PutMvt ( 'E_LETTRAGEDEV'       , '-' ) ;
 PutMvt ( 'E_DEVISE'            , V_PGI.DevisePivot ) ;
 PutMvt ( 'E_CONTROLE'          , '-' ) ;
 PutMvt ( 'E_TIERSPAYEUR'       , '' ) ;
 PutMvt ( 'E_QUALIFQTE1'        , '...' ) ;
 PutMvt ( 'E_QUALIFQTE2'        , '...' ) ;
 PutMvt ( 'E_ECRANOUVEAU'       , 'N' ) ;
 PutMvt ( 'E_DATEPAQUETMIN'     , V_PGI.DateEntree ) ;
 PutMvt ( 'E_DATEPAQUETMAX'     , V_PGI.DateEntree ) ;
 PutMvt ( 'E_ETATLETTRAGE'      , 'RI' ) ;
 PutMvt ( 'E_ENCAISSEMENT'      , 'RIE' ) ;
 PutMvt ( 'E_EMETTEURTVA'       , '-' ) ;
 PutMvt ( 'E_ANA'               , '-' ) ;
 PutMvt ( 'E_ECHE'              , '-' ) ;
 PutMvt ( 'E_FLAGECR'           , '' ) ;
 PutMvt ( 'E_DATETAUXDEV'       , V_PGI.DateEntree ) ;
 PutMvt ( 'E_CONTROLETVA'       , 'RIE' ) ;
 PutMvt ( 'E_CONFIDENTIEL'      , '0' ) ;
 PutMvt ( 'E_CREERPAR'          , 'SAI' ) ;
 PutMvt ( 'E_EXPORTE'           , '---' ) ;
 PutMvt ( 'E_TRESOLETTRE'       , '-' ) ;
 PutMvt ( 'E_CFONBOK'           , '-' ) ;
 PutMvt ( 'E_MODESAISIE'        , '-' ) ;
 PutMvt ( 'E_EQUILIBRE'         , '-' ) ;
 PutMvt ( 'E_AVOIRRBT'          , '-' ) ;
 PutMvt ( 'E_CODEACCEPT'        , 'NON' ) ;
 PutMvt ( 'E_EDITEETATTVA'      , '-' ) ;
 PutMvt ( 'E_ETATREVISION'      , '-' ) ;
 //GP COM
 PutMvt ( 'E_REFREVISION'       , 0 ) ;
 PutMvt ( 'E_PAQUETREVISION'    , 0 ) ;
 PutMvt ( 'E_IO'                , 'X') ;
 // SBO : Synchro Tréso
 PutMvt ( 'E_TRESOSYNCHRO'      , 'RIE') ;

end;

{=====================================================================}

procedure TOBM.PutDefautBud;
begin
 {$IFNDEF SANSCOMPTA}
 PutMvt ( 'BE_ETABLISSEMENT'    , GetParamSocSecur('SO_ETABLISDEFAUT', '') ) ;
 {$ENDIF}
 PutMvt ( 'BE_QUALIFPIECE'      , 'N' ) ;
 PutMvt ( 'BE_VALIDE'           , '-' ) ;
 PutMvt ( 'BE_UTILISATEUR'      , V_PGI.User ) ;
 PutMvt ( 'BE_DATECREATION'     , Date ) ;
 PutMvt ( 'BE_DATEMODIF'        , NowH ) ;
 PutMvt ( 'BE_CONTROLE'         , '-' ) ;
 PutMvt ( 'BE_SOCIETE'          , V_PGI.CodeSociete ) ;
 PutMvt ( 'BE_CONFIDENTIEL'     , '0' ) ;
 PutMvt ( 'BE_CREERPAR'         , 'SAI' ) ;
 PutMvt ( 'BE_TYPESAISIE'       , 'G' ) ;
 PutMvt ( 'BE_RESOLUTION'       , 'C' ) ;
end;

{=====================================================================}

procedure TOBM.PutDefautAna ( Axe : String3 ) ;
begin
 {$IFNDEF SANSCOMPTA}
// PutMvt ( 'Y_EXERCICE'          , VH^.Entree.Code ) ;
 PutMvt ( 'Y_ETABLISSEMENT'     , GetParamSocSecur('SO_ETABLISDEFAUT', '') ) ;
 {$IFNDEF SPEC302}
 PutMvt ( 'Y_PERIODE'           , GetPeriode ( V_PGI.DateEntree ) ) ;
 PutMvt ( 'Y_SEMAINE'           , NumSemaine ( V_PGI.DateEntree ) ) ;
 {$ENDIF}
 {$ENDIF}
 PutMvt ( 'Y_AXE'               , Axe ) ;
 PutMvt ( 'Y_DATECOMPTABLE'     , V_PGI.DateEntree ) ;
 PutMvt ( 'Y_NATUREPIECE'       , 'OD' ) ;
 PutMvt ( 'Y_QUALIFPIECE'       , 'N' ) ;
 PutMvt ( 'Y_ETAT'              , '0000000000' ) ;
 PutMvt ( 'Y_UTILISATEUR'       , V_PGI.User ) ;
 PutMvt ( 'Y_DATECREATION'      , Date ) ;
 PutMvt ( 'Y_DATEMODIF'         , NowH ) ;
 PutMvt ( 'Y_SOCIETE'           , V_PGI.CodeSociete ) ;
 PutMvt ( 'Y_DEVISE'            , V_PGI.DevisePivot ) ;
 PutMvt ( 'Y_TAUXDEV'           , 1 ) ;
 PutMvt ( 'Y_DATETAUXDEV'       , V_PGI.DateEntree ) ;
 PutMvt ( 'Y_CONTROLE'          , '-' ) ;
 PutMvt ( 'Y_QUALIFQTE1'        , '...' ) ;
 PutMvt ( 'Y_QUALIFQTE2'        , '...' ) ;
 PutMvt ( 'Y_QUALIFECRQTE1'     , '...' ) ;
 PutMvt ( 'Y_QUALIFECRQTE2'     , '...' ) ;
 PutMvt ( 'Y_ECRANOUVEAU'       , 'N' ) ;
 PutMvt ( 'Y_CREERPAR'          , 'SAI' ) ;
 PutMvt ( 'Y_EXPORTE'           , '---' ) ;
 PutMvt ( 'Y_TYPEANALYTIQUE'    , '-' ) ;
 PutMvt ( 'Y_JOURNAL'           , '' ) ;
 PutMvt ( 'Y_VALIDE'            , '-' ) ;
 PutMvt ( 'Y_CONFIDENTIEL'      , '0' ) ;
  // CEGID V9
  PutMvt('Y_DATPER',iDate1900) ;
  PutMvt('Y_ENTITY',0) ;
  PutMvt('Y_REFGUID','') ;
  // ----

end;


procedure TOBM.PutDefautDivers;
begin
 // verif utiliter de la fct
 InitValeurs(false);
end;


procedure TOBM.PutDefaut ( Axe : String3 ) ;
begin
 PutDefautDivers;
 case Ident of
  EcrGen : PutDefautEcr;
  EcrAna : PutDefautAna ( Axe ) ;
  EcrBud : PutDefautBud;
 end;
end;


procedure TOBM.HISTOMONTANTS;
begin
 case Ident of
  EcrGen :
   begin
    PutMvt ( 'OLDDEBIT', GetMvt ( 'E_DEBIT' ) ) ;
    PutMvt ( 'OLDCREDIT', GetMvt ( 'E_CREDIT' ) ) ;
   end;
  EcrBud :
   begin
    PutMvt ( 'OLDDEBIT', GetMvt ( 'BE_DEBIT' ) ) ;
    PutMvt ( 'OLDCREDIT', GetMvt ( 'BE_CREDIT' ) ) ;
   end;
  EcrAna :
   begin
    PutMvt ( 'OLDDEBIT', GetMvt ( 'Y_DEBIT' ) ) ;
    PutMvt ( 'OLDCREDIT', GetMvt ( 'Y_CREDIT' ) ) ;
   end;
 end;
end;

{JP 24/10/07 : La fonction EgalChamps n'est pas satisafaisante car elle ne gère que les tobs en CWAS
{---------------------------------------------------------------------------------------}
procedure TOBM.EgalChampsTOB(TT : TOB);
{---------------------------------------------------------------------------------------}
{$IFNDEF SANSCOMPTA}
var
  i, lIndex : Integer;
  iNumTable : Integer;
  szPrefixe, szNomChamp : String;
  MCD : IMCDServiceCOM;
	Table     : ITableCOM ;

{$ENDIF}
begin
  {$IFNDEF SANSCOMPTA}
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  lIndex:=0 ;
  try
      // Recherche du préfixe et du n° de la table
      szPrefixe := TableToPrefixe(TT.NomTable);
      iNumTable := PrefixeToNum(szPrefixe);
      Table := MCD.GetTable(TT.NomTable);
      if Not Assigned(Table) then Exit;
      for i := 1 to TT.NbChamps do
      begin
        TT.AddChampSup((Table.GetField(i) as IFieldCOM).Name ,False);
      end;

      for i := 1 to TT.NbChamps do
      begin
        if (((Table.GetField(i) as IFieldCOM).tipe <> 'MEMO') and ((Table.GetField(i) as IFieldCOM).tipe <> 'BLOB')) then
        begin
          szNomChamp := TT.GetNomChamp(i);
          TT.PutValue(szNomChamp, GetValue(szNomChamp));
        end;
      end;
  except
    on E:Exception do MessageAlerte('Erreur lors de l''appel de la fonction EgalChampsTOB' +#10#13+e.Message+' index :'+inttostr(lIndex) );
  end;
  {$ELSE}
    PGIInfo('Fonction non dispo','Erreur');
  {$ENDIF}
end;

procedure TOBM.EgalChamps ( TT : TDataSet ) ;
{$IFNDEF SANSCOMPTA}
{$IFNDEF EAGLCLIENT}
var
  i, lIndex : Integer;
{$ENDIF EAGLCLIENT}
{$ENDIF SANSCOMPTA}
begin
{$IFNDEF SANSCOMPTA}
  {$IFDEF EAGLCLIENT}
  {JP 24/10/07 : création de la fonction EgalChampsTOB, indépendant du fait que l'on soit en eAGL}
  EgalChampsTOB(TT);
  {$ELSE}
  lIndex:=0 ;
  try
    for i := 0 to TT.FieldCount - 1 do
      if ( ( TT.Fields[i].DataType <> ftMemo ) and ( TT.Fields[i].DataType <> ftBlob ) ) then begin
        lIndex:=i ;
        if TT.Fields[i].FieldName <> '' then
          TT.Fields[i].AsVariant := GetValue(TT.Fields[i].FieldName);
      end;
  except
    on E:Exception do MessageAlerte('Erreur lors de l''appel de la fonction EgalChamps' +#10#13+e.Message+' index :'+inttostr(lIndex) );
  end;
  {$ENDIF}
{$ELSE}
  PGIInfo('Fonction non dispo','Erreur');
{$ENDIF}
end;


procedure TOBM.MajLesDates ( Action : TActionFiche ) ;
var
 Pref : string;
begin

 if Action <> taCreat then
  Exit;

 case Ident of
  EcrGen : Pref := 'E_';
  EcrAna : Pref := 'Y_';
  EcrBud : Pref := 'BE_';
  else
   Exit;
 end; // case

 PutValue ( Pref + 'DATECREATION', Date ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 23/09/2004
Description .. : Construit la requet de mise à jour en fonction des champs 
Suite ........ : modifier
Suite ........ : 
Suite ........ : - LG - 09/04/2002 - ne pas prendre en compte le champs 
Suite ........ : E_BLOCNOTE
Suite ........ : -LG - 06/08/20004 - FB 13867 - ajout de e_tresosynchro a 
Suite ........ : la liste des champs a exclure ( affete manuellement ds la rq )
Suite ........ : JP 23/09/04 : tout compte fait cela pose moins de 
Suite ........ : problèmes d'affecter e_tresosynchro par la tob dans lettUtil 
Suite ........ : que d'exclure le champ de stPourUpdate
Mots clefs ... : 
*****************************************************************}
function TOBM.STPourUPDATE : string;
var
  k       : Integer;
  St,tmp  : string;
  Nom     : string;
begin
  St := '';

  for k := 0 to NbChamps - 1 do
  begin

    Nom := GetNomChamp ( k );
    if ( ( Nom <> 'OLDDEBIT' ) and ( Nom <> 'OLDCREDIT' ) and ( Nom <> 'E_DATEMODIF' ) and ( Nom <> 'E_BLOCNOTE' ) and  {( Nom <> 'E_TRESOSYNCHRO' ) and JP 23/09/04}
       ( Nom <> 'RATIO' ) )   and IsFieldModified( Nom ) then
    begin
      // BPY le 15/10/2004 : gestion des double quote
      if (TVarData(GetValeur(k)).VType = varString) then tmp := '"' + CheckdblQuote(GetValeur(k)) + '"'
      else tmp := VariantToSQL(GetValeur(k));
      St := St + ', ' + Nom + '=' + tmp;
      // Fin BPY
    end;

  end; // for

  if St <> '' then System.Delete ( St, 1, 2 ) ;

  Result := St;
end;

function TOBM.UpdateDB : Boolean;
var S : string;
    DD : TDateTime;
    Sql : string;
begin
 PutMvt('E_PAQUETREVISION',1) ;
 result:=true ;
 try
  S:= StPourUpdate ; DD:= GetMvt('E_DATEMODIF') ;
  if Trim(S)<>'' Then
   begin
    SQL := 'UPDATE ECRITURE SET '+S+', E_DATEMODIF="'+UsTime(NowH)+'"'
            +' Where '
            +' E_JOURNAL="'+ GetMvt('E_JOURNAL')+'" AND E_EXERCICE="'+ GetMvt('E_EXERCICE')+'"'
            +' AND E_DATECOMPTABLE="'+UsDateTime( GetMvt('E_DATECOMPTABLE'))+'" AND E_NUMEROPIECE='+InttoStr( GetMvt('E_NUMEROPIECE'))
            +' AND E_QUALIFPIECE="'+ GetMvt('E_QUALIFPIECE')+ '"'
            +' AND E_NUMLIGNE='+IntToStr(GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(GetMvt('E_NUMECHE'))
            +' AND E_DATEMODIF="'+UsTime(DD)+'"' ;

     Result:=(ExecuteSQL(SQL)=1) ;
   end; // if
 except
  on e : exception do MessageAlerte('Erreur lors de l''enregistrement en base.'+#10#13+E.Message)
 end;
end;


function TOBM.GetT ( Index : integer ) : TT_V;
begin

 if ( Index >= NbChamps) or ( Index = 0 ) then exit;

 try

  result.V    := GetValeur(Index);
  result.Etat := Etat;

 except
  on E:Exception do MessageAlerte('Erreur lors l''affectation des champs!' +#10#13+e.Message+' index :'+IntToStr(Index) );
 end;

end;

procedure TOBM.SetT ( Index : integer ; Value : TT_V );
begin

 if ( Index >= NbChamps) or ( Index = 0 ) then exit;

 try

 PutValeur( Index , Value.V ) ;
 Value.V := Modifie;

 except
  on E:Exception do MessageAlerte('Erreur lors l''affectation des champs!' +#10#13+e.Message+' index :'+IntToStr(Index) );
 end;

end;


procedure TOBM.SetStatus ( Value : TStatus ) ;
begin
 // utiliser dans SaisieTr
 SStatus := Value;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 30/05/2005
Modifié le ... :   /  /    
Description .. : - LG - 30/05/2005 - plantait lors d'un transtypage d'une 
Suite ........ : TOB en TOBM
Mots clefs ... : 
*****************************************************************}
procedure TOBM.SetEtat ( Value : TEtat ) ;
begin
 if GetNumChamp('_ETAT') <> - 1 then
  PutValue('_ETAT',integer(Value)) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/01/2004
Modifié le ... : 05/06/2007
Description .. : - 28/01/2004 - LG - pour determiner l'etat du TOBM,on 
Suite ........ : utilise le flag du TOB en parcourant les champs ( pb en 
Suite ........ : lettrage ds lettragefichier, le test etait tout le temps vraie )
Suite ........ : - LG - 05/06/2007 - FB 19943 - on ne tient plus compte du 
Suite ........ : champ e_controleur pour savoir sir la ligne est modifié. pb 
Suite ........ : en lettrage
Mots clefs ... : 
*****************************************************************}
function TOBM.GetEtat : TEtat ;
var
 k : integer ;
 Nom : string ;
begin

 result := Inchange ;

 for k := 0 to NbChamps - 1 do
  begin
    Nom := GetNomChamp ( k );
   if ( ( Nom <> 'OLDDEBIT' ) and ( Nom <> 'OLDCREDIT' ) and ( Nom <> 'E_DATEMODIF' ) and ( Nom <> 'E_BLOCNOTE' ) and
       ( Nom <> 'RATIO' ) and ( Nom <> 'E_CONTROLEUR' ) )   and IsFieldModified( Nom ) then
     begin
      result := TEtat(Modifie) ;
      exit ;
     end ; // if
  end; // for

end;

procedure TOBM.MajEtat ( Value : TEtat ) ;
begin
 Etat := Value;  // pas vraiment la meme chose que l'original
// for k := 0 to NbC do
//  T[k].Etat := Value;
end;

function TOBM.GetDebitSaisi : double;
var P : string;
begin
// faire une seule fct
 Result := 0;
 if Ident = EcrGen then P := 'E'
 else
  if Ident = EcrAna then P := 'Y'
  else
   if Ident = EcrBud then P := 'BE'
   else
    Exit;

  if Ident = EcrBud then
   Result := GetMvt ( P + '_DEBIT' )
  else
   Result := GetMvt ( P + '_DEBITDEV' ) ;

end;

function TOBM.GetCreditSaisi : double;
var P : string;
begin
 Result := 0;
 if Ident = EcrGen then
  P := 'E'
 else
  if Ident = EcrAna then
   P := 'Y'
  else
   if Ident = EcrBud then
    P := 'BE'
   else
    Exit;
  if Ident = EcrBud then
   Result := GetMvt ( P + '_CREDIT' )
  else
   Result := GetMvt ( P + '_CREDITDEV' ) ;
end;


procedure TOBM.SetCotation ( DD : TDateTime ) ;
var
 Cote, Taux : Double;
 Dev : String3;
begin
 if Ident <> EcrGen then
  Exit;
 if DD <= 0 then
  DD := GetMvt ( 'E_DATECOMPTABLE' ) ;
 Taux := GetMvt ( 'E_TAUXDEV' ) ;
 if DD < V_PGI.DateDebutEuro then
  Cote := Taux
 else
  begin
   Dev := GetMvt ( 'E_DEVISE' ) ;
   if ( ( Dev = V_PGI.DevisePivot ) or ( Dev = V_PGI.DeviseFongible ) ) then
    Cote := 1.0
   else
    if V_PGI.TauxEuro <> 0 then
     Cote := Taux / V_PGI.TauxEuro
    else
     Cote := 1;
  end;
 PutMvt ( 'E_COTATION', Cote ) ;
end;

procedure TOBM.SetMontants ( XD, XC : Double; DEV : RDEVISE; Force : boolean ) ;
var
 SD, SC : double;
 P : string;
begin
 if Ident = EcrGen then
  P := 'E'
 else
  if Ident = EcrAna then
   P := 'Y'
  else
   if Ident = EcrBud then
    P := 'BE'
   else
    Exit;

 // saisie en pivot
 if DEV.Code = V_PGI.DevisePivot then
  begin
   // saisie en pivot
   SD := GetMvt ( P + '_DEBIT' ) ;
   SC := GetMvt ( P + '_CREDIT' ) ;
   if ( ( SD = XD ) and ( SC = XC ) and ( not Force ) ) then
    Exit;
   PutMvt ( P + '_DEBIT', XD ) ;
   PutMvt ( P + '_CREDIT', XC ) ;
   PutMvt ( P + '_DEBITDEV', XD ) ;
   PutMvt ( P + '_CREDITDEV', XC ) ;
  end
   else
    begin
     // saisie en devise
     SD := GetMvt ( P + '_DEBITDEV' ) ;
     SC := GetMvt ( P + '_CREDITDEV' ) ;
     if ( ( SD = XD ) and ( SC = XC ) and ( not Force ) ) then
      Exit;
     PutMvt ( P + '_DEBIT', DeviseToEuro ( XD, DEV.Taux, DEV.Quotite ) ) ;
     PutMvt ( P + '_CREDIT', DeviseToEuro ( XC, DEV.Taux, DEV.Quotite ) ) ;
     PutMvt ( P + '_DEBITDEV', XD ) ;
     PutMvt ( P + '_CREDITDEV', XC ) ;
    end;

end;

procedure TOBM.SetMontantsBUD ( XD, XC : Double) ;
begin
 if Ident <> EcrBud then Exit;
 PutMvt ( 'BE_DEBIT', XD ) ;
 PutMvt ( 'BE_CREDIT', XC ) ;
end;

procedure TOBM.TraiteLeEtat ( PasEcart,AvecRegul,EcartChange,AvecConvert,ModeAuto : boolean; IndiceRegul : integer ; ModeRappro : Char ) ;
Var St: String ;
BEGIN
St:=GetMvt('E_ETAT') ; if Length(St)<6 then Exit ;
if AvecRegul then BEGIN St[1]:='X' ; St[4]:=IntToStr(IndiceRegul)[1] ; END else St[1]:='-' ;
if EcartChange then St[2]:='X' else
 if AvecConvert then St[2]:='#' else St[2]:='-' ;
if PasEcart then St[3]:='X' else St[3]:='-' ;
if ModeAuto then St[5]:='A' else St[5]:='M' ;
// M : rapprochement par montant R : Rapprochement par reference  S : Rapprochement par solde
St[6]:=ModeRappro ;
PutValue('E_ETAT',St) ;
END ;

function TOBM.GetDateComptable: TDateTime;
begin
 result:=GetMvt('E_DATECOMPTABLE');
end;

procedure TOBM.SetDateComptable(Value: TDateTime);
begin
 PutMvt('E_DATECOMPTABLE', Value);
 PutMvt('E_PERIODE'      , GetPeriode ( Value ) ) ;
 PutMvt('E_SEMAINE'      , NumSemaine ( Value ) ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : Vérifie le couple Période/Semaine par rapport à la date
Suite ........ : comptable
Mots clefs ... :
*****************************************************************}
function TOBM.OkPeriodeSemaine : Boolean;
begin                                        // Variant
  Result := (NumSemaine(DateComptable) = GetMvt('E_SEMAINE')) and
            (GetPeriode(DateComptable) = GetMvt('E_PERIODE')) ;
end;

function TOBM.OkEcrANouveau(vJNatureJal: string) : Boolean;
var lEcrANouveau : string;
begin
  lEcrANouveau := getMvt('E_ECRANOUVEAU');
  if (vJNatureJal = 'ANO') then
    Result := (lEcrANouveau = 'O') or (lEcrANouveau = 'H') or (lEcrANouveau = 'OAN')
  else
    Result := (lEcrANouveau = 'N');
end;

function TOBM.OkModeSaisie(vJModeSaisie : string): Boolean;
begin
  Result := (vJModeSaisie = GetMvt('E_MODESAISIE'));
end;

function TOBM.GetModeSaisie : string;
begin
  Result := GetMvt('E_MODESAISIE');
end;

procedure TOBM.SetModeSaisie(const Value : string);
begin
  PutMvt('E_MODESAISIE', Value );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/04/2002
Modifié le ... :   /  /
Description .. : Retourne le type d'une ecriture en fct de son qualiforigine
Suite ........ :
Suite ........ : REG : ecriture de regul de lettrage
Suite ........ : ECC : ecart de conversion en saisie
Suite ........ : SIM : ecriture simplifie
Suite ........ : TRE : ecriture genere par la saisie de tresorerie
Suite ........ : CON : ecart de conversion lettrage
Mots clefs ... :
*****************************************************************}
function TOBM.GetTypeEcr : TTypeEcriture;
var
 lStQualifOrigine : string;
begin
 lStQualifOrigine:=GetValue('E_QUALIFORIGINE') ;
 if lStQualifOrigine = 'REG' then result:=TERegulLettrage
  else if lStQualifOrigine = 'ECC' then result:=TEEcartEuro
   else if lStQualifOrigine = 'SIM' then result:=TESimplifie
    else if lStQualifOrigine = 'TRE' then result:=TETreso
     else if lStQualifOrigine = 'CON' then result:=TEEcartEuroLettrage
      else result:=TESaisie;
end;

//LG* hors directive compilation
procedure EgaliseOBM ( var OS, OD : TOBM ) ;
begin
 if OS = nil then
  Exit;
 if OD = nil then
  OD := TOBM.Create ( OS.Ident, '', False ) ;
 //OD.T := OS.T; !!!
 // LG*
 OD.Dupliquer ( OS , true , true );
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/09/2002
Modifié le ... : 04/08/2004
Description .. : -25/09/2002- Ajout de larecopie de e_refrevision pour
Suite ........ : l'ensemble de la piece
Suite ........ : - 04/08/2004 - FB 12694 - recriture de la fct : on recher le 
Suite ........ : prems e_refrevision diff de 0 et on le recopie sur l'ensemble 
Suite ........ : de la peice
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function MajIOCom(GS : THGrid) : Boolean ;
Var i    : integer ;
    O    : TOBM ;
    lInRefRevision : integer ;
BEGIN
result := false ;
lInRefRevision := 0 ;
for i := 1 to GS.RowCount - 2 do
 begin
  O := GetO(GS,i) ; if O = nil then Continue ;
  if O.GetMvt('E_REFREVISION')<>0 then
   begin
    lInRefRevision:=O.GetValue('E_REFREVISION') ;  // on recupere le code revision
    result := true ;
    break ;
   end ;
 end ;// for
if result Then
 for i := 1 to GS.RowCount-2 do
  begin
   O := GetO(GS,i) ; if O= nil then Continue ; O.PutMvt('E_IO','X') ;
   // affectation du code revision s'il est different de 0
   if (lInRefRevision <> 0) and (O.GetValue('E_REFREVISION') = 0) then O.PutValue('E_REFREVISION',lInRefRevision) ;
  end ; // for
END ;
{$ENDIF}

Procedure SauveJalEffet(St,Jal : String) ;
Var Q : TQuery ;
    St1 : String ;
BEGIN
If St='' then Exit ; If Jal='' Then Exit ;
Q:=OpenSQL('SELECT * FROM CHOIXEXT WHERE YX_TYPE="CSE" AND YX_CODE="'+V_PGI.User+';'+St+';" ',FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('YX_TYPE').AsString:='CSE' ;
  Q.FindField('YX_CODE').AsString:=V_PGI.User+';'+St+';' ;
  Q.FindField('YX_LIBELLE').AsString:=Jal ;
  Q.Post ;
  END Else
  BEGIN
  St1:=Q.FindField('YX_LIBELLE').AsString ;
  If St1<>Jal Then
    BEGIN
    Q.Edit ;
    Q.FindField('YX_LIBELLE').AsString:=Jal ;
    Q.Post ;
    END ;
  END ;
Ferme(Q) ;
END ;

Function  ChargeJalEffet(St : String) : String ;
Var Q : TQuery ;
BEGIN
Result:='' ;
If St='' then Exit ;
Q:=OpenSQL('SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="CSE" AND YX_CODE="'+V_PGI.User+';'+St+';" ',TRUE,-1,'',true) ;
If Not Q.Eof Then Result:=Q.FindField('YX_LIBELLE').AsString ;
Ferme(Q) ;
END ;

(*======================================================================*)
PROCEDURE CALC(Var Deb,Cred,Tot,TauxC : Double ; Sens,Decim : Byte) ;
BEGIN
If Sens=1 Then BEGIN Deb:=Arrondi(Deb*TauxC,Decim) ; Tot:=Tot+Deb ; END
          else BEGIN Cred:=Arrondi(Cred*TauxC,Decim) ; Tot:=Tot+Cred ; END ;
END ;

(*======================================================================*)
PROCEDURE RECALC(Var Deb,Cred,Tot,MtEcr : Double ; Sens : Byte) ;
BEGIN
If Sens=1 Then Deb:=Deb+(MtEcr-Tot) Else Cred:=Cred+(MtEcr-Tot) ;
END ;

function TSorteLettreToStr(TSL : TSorteLettre) : String;
begin
  if (TSL = tslCheque) then Result := 'TSLCHEQUE' else
  if (TSL = tslTraite) then Result := 'TSLTRAITE' else
  if (TSL = tslBOR)    then Result := 'TSLBOR'    else
  if (TSL = tslVir)    then Result := 'TSLVIR'    else
  if (TSL = tslPre)    then Result := 'TSLPRE'    else
                            Result := 'TSLAUCUN';
end;

function StrToTSorteLettre(psz : String) : TSorteLettre;
begin
  if (UpperCase(psz) = 'TSLCHEQUE') then Result := tslCheque else
  if (UpperCase(psz) = 'TSLTRAITE') then Result := tslTraite else
  if (UpperCase(psz) = 'TSLBOR')    then Result := tslBOR    else
  if (UpperCase(psz) = 'TSLVIR')    then Result := tslVir    else
  if (UpperCase(psz) = 'TSLPRE')    then Result := tslPre    else
                                         Result := tslAucun;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 27/11/2003
Modifié le ... :   /  /
Description .. : Retourne True si le compte dont le code est passé en
Suite ........ : paramètre est un compte de banque
Mots clefs ... :
*****************************************************************}
Function  EstCptBQE( Cpte : String ) : Boolean ;
begin
  if Trim(Cpte) = '' then
    begin
    Result:=False ;
    Exit ;
    end ;
  Result := ExisteSQL('Select G_GENERAL from GENERAUX where G_GENERAL="'+Cpte+'" AND G_NATUREGENE="BQE"') ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : JP
Créé le ...... : 26/04/2004
Modifié le ... : 09/08/2006
Description .. : Retourne True si le compte est synchronisable en
Suite ........ : Trésorerie
Suite ........ : 
Suite ........ : JP 09/08/06 : Gestion des comptes de groupes / courants
Mots clefs ... : 
*****************************************************************}
function  EstCptSyn( Cpte, Jnl : string; NonEcc : Boolean; NomBase : string = '') : Boolean ;
var
  Q : TQuery;

    {------------------------------------------------------------------}
    function IsJalOk(Facture : Boolean) : Boolean;
    {------------------------------------------------------------------}
    begin
      if Facture then
        Result := (Q.FindField('J_NATUREJAL').AsString = 'ACH') or
                  (Q.FindField('J_NATUREJAL').AsString = 'VTE') or
                  (Q.FindField('J_NATUREJAL').AsString = 'OD')
      else
        Result := EstMultiSoc and
                  ((Q.FindField('J_NATUREJAL').AsString = 'BQE') or
                   (Q.FindField('J_NATUREJAL').AsString = 'OD'));
    end;

    {------------------------------------------------------------------}
    function IsCptOk(Facture : Boolean) : Boolean;
    {------------------------------------------------------------------}
    begin
      if Facture then
        Result := (Q.FindField('G_NATUREGENE').AsString = 'COC') or
                  (Q.FindField('G_NATUREGENE').AsString = 'COD') or
                  (Q.FindField('G_NATUREGENE').AsString = 'TIC') or
                  (Q.FindField('G_NATUREGENE').AsString = 'TID') or
                  (Q.FindField('G_NATUREGENE').AsString = 'COF') or
                  (Q.FindField('G_NATUREGENE').AsString = 'COS')
      else
        Result := EstMultiSoc and
                  (Q.FindField('G_NATUREGENE').AsString = 'DIV') and
                  ExisteSQL('SELECT CLS_GENERAL FROM ' + GetTableDossier(NomBase, 'CLIENSSOC') +
                            ' WHERE CLS_GENERAL = "' + Cpte + '"');
    end;

begin
    Q := OpenSQL('SELECT G_NATUREGENE, J_NATUREJAL FROM JOURNAL, GENERAUX WHERE J_JOURNAL = "' + Jnl +
                 '" AND G_GENERAL="' + Cpte + '"', True,-1,'',true);
    try
      {1/ Compte bancaire, on synchronise}
      if Q.FindField('G_NATUREGENE').AsString = 'BQE' then
        Result := True
      {2/ Pas d'ecart de change, compte "Tiers" sur un journal de "facture}
      else if NonEcc and IsJalOk(True) and IsCptOk(True) then
        Result := True
      {3/ Compte divers appartenant à CLIENSSOC sur un journal de "règlement"}
      else if IsJalOk(False) and IsCptOk(False) then
        Result := True
      {... dans les autres cas l'écritures n'est pas synchronisable}
      else
        Result := False;
    finally
      Ferme(Q);
    end;
end;


////////////////////////////////////////////////////////////////////////////////
// BPY le 19/05/2003 : Gestion des extourne
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ExtourneEcriture(ListeEcriture : TOB; WithEcran : boolean=True; LaDate : TdateTime=0; Qualif : string='';
                          Negatif : boolean=false; Lib : string=''; RefI : string=''; RefE : string='';
                          WithInsert : boolean=True) : TOB;
var
    TInfo,ListePiece : TOB;

    Journal : string;
    ModeSaisie : string;
    nbLigne : integer;
    QualifPiece : string;
    SelectionDebit : double;
begin
    result := nil;
    ListePiece := TOB.Create('',nil,-1);

    // champs a fournir ....
    Journal := ListeEcriture.Detail[0].Getvalue('E_JOURNAL');
    QualifPiece := ListeEcriture.Detail[0].Getvalue('E_QUALIFPIECE');
    // champs ke je calcul
    ModeSaisie := GetColonneSQL('JOURNAL','J_MODESAISIE','J_JOURNAL="' + Journal + '"');
    nbLigne := ListeEcriture.Detail.Count;
    SelectionDebit := ListeEcriture.Somme('E_DEBIT',[''],[''],false);
    // creation de la TOB
    TInfo := TOB.Create('', nil, -1);
    try
      TInfo.AddChampSupValeur('JOURNAL',Journal);
      TInfo.AddChampSupValeur('MODESAISIE',ModeSaisie);
      TInfo.AddChampSupValeur('LIGNES',NbLigne);
      TInfo.AddChampSupValeur('QUALIFPIECE',QualifPiece);
      TInfo.AddChampSupValeur('MONTANT',Abs(SelectionDebit));
      {$IFDEF EAGLSERVER}
        WithEcran := False;
      {$ENDIF EAGLSERVER}
      if not WithEcran then
      begin
        TInfo.AddChampSupValeur('CONTRE_DATE', LaDate);
        TInfo.AddChampSupValeur('CONTRE_TYPE', Qualif);
  //      TInfo.AddChampSupValeur('CONTRE_NEGATIF', Negatif);
        TInfo.AddChampSup('CONTRE_NEGATIF', False);
        if Negatif then
          TInfo.PutValue('CONTRE_NEGATIF', 'X')
          else
          TInfo.PutValue('CONTRE_NEGATIF', '-');
        TInfo.AddChampSupValeur('CONTRE_LIBELLE', Lib);
        TInfo.AddChampSupValeur('CONTRE_REFINTERNE', RefI);
        TInfo.AddChampSupValeur('CONTRE_REFEXTERNE', RefE);
        result := GenereLesExtournes(ListeEcriture, ListePiece, true, true, TInfo.GetValue('CONTRE_DATE'),
                                     TInfo.GetValue('CONTRE_TYPE'), TInfo.GetValue('CONTRE_NEGATIF'),
                                     TInfo.GetValue('CONTRE_LIBELLE'), TInfo.GetValue('CONTRE_REFINTERNE'),
                                     TInfo.GetValue('CONTRE_REFEXTERNE'), ModeSaisie, WithInsert);
      end else
      begin
        {$IFNDEF EAGLSERVER}
          TheTOB := TInfo;
          if WithEcran then
            AGLLanceFiche('CP','EXTOURNE_PARAM','','','');
          if (TheTOB = nil) then exit;
          result := GenereLesExtournes(ListeEcriture, ListePiece, true, true, TheTOB.GetValue('CONTRE_DATE'),
                                       TheTOB.GetValue('CONTRE_TYPE'), TheTOB.GetValue('CONTRE_NEGATIF'),
                                       TheTOB.GetValue('CONTRE_LIBELLE'), TheTOB.GetValue('CONTRE_REFINTERNE'),
                                       TheTOB.GetValue('CONTRE_REFEXTERNE'), ModeSaisie, WithInsert);
        {$ENDIF !EAGLSERVER}
      end;
  finally
    TInfo.Free;
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function MemePiece(ListeEcriture : TOB ; mode : string ; iref,i : integer ; ForceBordereau : boolean = false) : boolean;
begin
    result := false;

    if (ListeEcriture.Detail.Count = 0) then exit;
    if (i >= ListeEcriture.Detail.Count) then exit;
    if (iref = -1) then exit;

    if (mode = '-') then result := ListeEcriture.Detail[iref].GetValue('E_NUMEROPIECE') = ListeEcriture.Detail[i].GetValue('E_NUMEROPIECE')
    else if ((mode = 'BOR') or (ForceBordereau)) then result := (ListeEcriture.Detail[iref].GetValue('E_NUMGROUPEECR') = ListeEcriture.Detail[i].GetValue('E_NUMGROUPEECR')) and (ListeEcriture.Detail[iref].GetValue('E_NUMEROPIECE') = ListeEcriture.Detail[i].GetValue('E_NUMEROPIECE')) and (ListeEcriture.Detail[iRef].GetValue('E_PERIODE') = ListeEcriture.Detail[i].GetValue('E_PERIODE'))
    else if (mode = 'LIB') then result := ListeEcriture.Detail[iref].GetValue('NUMGROUPEECRLIB') = ListeEcriture.Detail[i].GetValue('NUMGROUPEECRLIB');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TrouvePieceDejaCalculee(ListePiece : TOB ; T : TOB) : TOB;
var
    i: integer;
begin
    result := nil;

    for i := 0 to ListePiece.Detail.Count - 1 do
    begin
        if (ListePiece.Detail[i].GetValue('OLDPIECE') = T.GetValue('E_NUMEROPIECE')) and (ListePiece.Detail[i].GetValue('OLDPERIODE') = T.GetValue('E_PERIODE')) then
        begin
            result := ListePiece.Detail[i];
            break;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ChargeAnalytique(T : TOB);
var
    i : integer;
    TAxe : TOB;
    Q : TQuery;
    St : string;
begin
  T.ClearDetail ;
  for i := 1 to MAXAXE do
    begin
    TAxe := TOB.Create('$AXE', T, -1);
    TAxe.AddChampSupValeur('Y_AXE', 'A' + IntTostr(i) );
    St := 'SELECT * FROM ANALYTIQ WHERE Y_EXERCICE="' + T.GetValue('E_EXERCICE') + '" AND Y_JOURNAL="' + T.GetValue('E_JOURNAL') + '" AND Y_DATECOMPTABLE="' + USDateTime(StrToDate(T.GetValue('E_DATECOMPTABLE'))) + '" AND Y_NUMEROPIECE=' + IntToStr(T.GetValue('E_NUMEROPIECE')) + ' AND Y_NUMLIGNE=' + IntToStr(T.GetValue('E_NUMLIGNE')) + ' AND Y_QUALIFPIECE="' + T.GetValue('E_QUALIFPIECE') + '" AND Y_AXE="' + 'A' + IntToStr(i) + '" ORDER BY Y_NUMVENTIL';
    Q := OpenSQL(St,true,-1,'',true);
    TAxe.LoadDetailDB('ANALYTIQ','','',Q,true);
    Ferme(Q);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ExtourneLigneEcriture(T : TOB ; TInfoPiece : TOB ; date : TDateTime ; tipe,negatif,libelle,refint,refext,mode : string);
var
    EL : string;
    StRef : string;
    X : double;
    i,j : integer;
    TAxe,TAna : TOB;
begin
    if (T = nil) then exit;

    StRef := Copy(TraduireMemoire('Extourne') + ' ' + Inttostr(T.GetValue('E_NUMEROPIECE')) + '  ' + DateToStr(T.GetValue('E_DATECOMPTABLE')),1,35);

    if (libelle <> '') then T.PutValue('E_LIBELLE',libelle);
    if (refint = '') then T.PutValue('E_REFINTERNE',StRef)
    else T.PutValue('E_REFINTERNE',refint);
    if (refext <> '') then T.PutValue('E_REFEXTERNE',refext);

    T.PutValue('E_DATECOMPTABLE', date);
    T.PutValue('E_EXERCICE',QuelExoDt(date));
    T.PutValue('E_DATEECHEANCE', date);
    if (mode <> '-') then
    begin
        T.PutValue('E_NUMGROUPEECR',TInfoPiece.GetValue('NUMGROUPEECR'));
        T.PutValue('E_NUMLIGNE',TInfoPiece.GetValue('NUMLIGNE'));
        TInfoPiece.PutValue('NUMLIGNE',TInfoPiece.GetValue('NUMLIGNE') + 1);
    end;
{$IFNDEF SPEC302}
    T.PutValue('E_PERIODE',GetPeriode(date));
    T.PutValue('E_SEMAINE',NumSemaine(date));
{$ENDIF}
    T.PutValue('E_EXERCICE',QuelExoDT(date));
    T.PutValue('E_QUALIFPIECE',tipe);
    T.PutValue('E_NUMEROPIECE',TInfoPiece.GetValue('NEWPIECE'));
    T.PutValue('E_VALIDE','-');
    T.PutValue('E_LETTRAGE','');
    T.PutValue('E_TRACE','');
    T.PutValue('E_QUALIFORIGINE','EXT');
    EL := T.GetValue('E_ETATLETTRAGE');
    if ((EL = 'TL') or (EL = 'PL') or (EL = 'AL')) then T.PutValue('E_ETATLETTRAGE','AL')
    else T.PutValue('E_ETATLETTRAGE','RI');
    T.PutValue('E_LETTRAGEDEV','-');
    T.PutValue('E_COUVERTURE',0);
    T.PutValue('E_COUVERTUREDEV',0);
    T.PutValue('E_DATEPAQUETMIN',date);
    T.PutValue('E_DATEPAQUETMAX',date);
    T.PutValue('E_REFPOINTAGE','');
    T.PutValue('E_DATEPOINTAGE',IDate1900);
    T.PutValue('E_REFRELEVE','');
    T.PutValue('E_FLAGECR','');
    T.PutValue('E_ETAT','0000000000');
    T.PutValue('E_NIVEAURELANCE',0);
    T.PutValue('E_DATERELANCE',IDate1900);
    T.PutValue('E_DATECREATION',date);
    T.PutValue('E_DATEMODIF',NowH);
    T.PutValue('E_SUIVDEC','');
    T.PutValue('E_NOMLOT','');
    T.PutValue('E_EDITEETATTVA','-');
    T.PutValue('E_IO','X');
    if (negatif = 'X') then
    begin
        T.PutValue('E_DEBIT',-T.GetValue('E_DEBIT'));
        T.PutValue('E_CREDIT',-T.GetValue('E_CREDIT'));
        T.PutValue('E_DEBITDEV',-T.GetValue('E_DEBITDEV'));
        T.PutValue('E_CREDITDEV',-T.GetValue('E_CREDITDEV'));
        T.PutValue('E_QTE1',-T.GetValue('E_QTE1'));
        T.PutValue('E_QTE2',-T.GetValue('E_QTE2'));
    end
    else
    begin
        X := T.GetValue('E_DEBIT');
        T.PutValue('E_DEBIT',T.GetValue('E_CREDIT'));
        T.PutValue('E_CREDIT',X);
        X := T.GetValue('E_DEBITDEV');
        T.PutValue('E_DEBITDEV',T.GetValue('E_CREDITDEV'));
        T.PutValue('E_CREDITDEV',X);
    end;
    // Traitement de l'analytique
    for i := 0 to T.Detail.Count - 1 do
    begin
        TAxe := T.Detail[i];
        for j := 0 to TAxe.Detail.Count - 1 do
        begin
            TAna := TAxe.Detail[j];
            TAna.PutValue('Y_DATECOMPTABLE',date);
            TAna.PutValue('Y_NUMLIGNE',T.GetValue('E_NUMLIGNE'));
{$IFNDEF SPEC302}
            TAna.PutValue('Y_PERIODE',GetPeriode(date));
            TAna.PutValue('Y_SEMAINE',NumSemaine(date));
{$ENDIF}
            TAna.PutValue('Y_EXERCICE',QuelExoDT(date));
            TAna.PutValue('Y_QUALIFPIECE',tipe);
            TAna.PutValue('Y_NUMEROPIECE',TInfoPiece.GetValue('NEWPIECE'));
            TAna.PutValue('Y_TRACE','');
            TAna.PutValue('Y_VALIDE','-');
            TAna.PutValue('Y_DATECREATION',Date);
            TAna.PutValue('Y_DATEMODIF',NowH);
            // CEGID V9
            TAna.PutValue('Y_DATPER',iDate1900) ;
            TAna.PutValue('Y_ENTITY',0) ;
            TAna.PutValue('Y_REFGUID','') ;
            // ----
            if (negatif = 'X') then
            begin
                TAna.PutValue('Y_DEBIT',-TAna.GetValue('Y_DEBIT'));
                TAna.PutValue('Y_CREDIT',-TAna.GetValue('Y_CREDIT'));
                TAna.PutValue('Y_DEBITDEV',-TAna.GetValue('Y_DEBITDEV'));
                TAna.PutValue('Y_CREDITDEV',-TAna.GetValue('Y_CREDITDEV'));
                TAna.PutValue('Y_QTE1',-TAna.GetValue('Y_QTE1'));
                TAna.PutValue('Y_QTE2',-TAna.GetValue('Y_QTE2'));
                TAna.PutValue('Y_TOTALQTE1',-TAna.GetValue('Y_TOTALQTE1'));
                TAna.PutValue('Y_TOTALQTE2',-TAna.GetValue('Y_TOTALQTE2'));
                TAna.PutValue('Y_TOTALECRITURE',-TAna.GetValue('Y_TOTALECRITURE'));
                TAna.PutValue('Y_TOTALDEVISE',-TAna.GetValue('Y_TOTALDEVISE'));
            end
            else
            begin
                X := TAna.GetValue('Y_DEBIT');
                TAna.PutValue('Y_DEBIT',TAna.GetValue('Y_CREDIT'));
                TAna.PutValue('Y_CREDIT',X);
                X := TAna.GetValue('Y_DEBITDEV');
                TAna.PutValue('Y_DEBITDEV',TAna.GetValue('Y_CREDITDEV'));
                TAna.PutValue('Y_CREDITDEV',X);
            end;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function GenereLesExtournes(ListeEcriture,ListePiece : TOB ; bSelectionPiece : boolean ; All : boolean ; date : TDateTime ; tipe,negatif,libelle,refint,refext,mode : string; WithInsert : boolean=true) : TOB;
var
    i,iRef,NewPiece : integer;
    TEcr,TInsert,TNewEcr,TInfoPiece : TOB;
    bPieceUnique : boolean;
    result1,result2 : boolean;
    numerror : integer;
    liberror : string;
begin
    ListePiece.ClearDetail;
    result1 := true;
//    result2 := true;

    TInfoPiece := nil;
    bPieceUnique := false;
    i := 0;
    TInsert := TOB.Create('',nil,-1);

    // Suppression des enregistrements non sélectionnés
    if (not (All)) then
    begin
        while (i < ListeEcriture.Detail.Count) do
        begin
            if (ListeEcriture.Detail[i].GetValue('SELECTION') <> '+') then ListeEcriture.Detail[i].Free
            else Inc(i,1);
        end;
    end;

    // Calcul du numéro de pièce pour le cas de la sélection libre

    // Traitements des enregistrements sélectionnés
    iRef := -1;
    for i := 0 to ListeEcriture.Detail.Count - 1 do
    begin
        TEcr := ListeEcriture.Detail[i];
        TEcr.PutValue('E_QUALIFORIGINE', 'URN');

        // Détection de changement de pièce
        if ((not (bPieceUnique)) and (not (MemePiece(ListeEcriture,mode,iRef, i)))) then
        begin
            iRef := i;
            TInfoPiece := TrouvePieceDejaCalculee(ListePiece,TEcr); // TODO

            // Calcul du nouveau numéro de pièce
            if (TInfoPiece = nil) then
            begin
                // Si nouvelle pièce, on enregistre la dernière pièce pour ne pas
                // perturber le calcul du numéro de pièce
                if ((i <> 0) and (TInsert <> nil)) then
                begin
                    TInsert.SetAllModifie(true);
                    { JTR - Lorsqu'on vient de la GC et que l'on est en COMPTADIFF
                      il ne faut pas faire d'INSERT }
                    if WithInsert then
                    begin
                      TInsert.InsertDB(nil,true);
                      MajSoldesEcritureTOB(TInsert,true);
                    end;
                    TInsert.ClearDetail;
                end;
                // Calcul des infos de la nouvelle pièce
                NewPiece := GetNewNumJal(TEcr.GetValue('E_JOURNAL'),(TEcr.GetValue('E_QUALIFPIECE') = 'N'),date,'',mode);
                TInfoPiece := TOB.Create('',ListePiece, -1);
                TInfoPiece.AddChampSupValeur('OLDPIECE',TEcr.GetValue('E_NUMEROPIECE'));
                TInfoPiece.AddChampSupValeur('NEWPIECE',NewPiece);
                TInfoPiece.AddChampSupValeur('OLDPERIODE',TEcr.GetValue('E_PERIODE'));
                TInfoPiece.AddChampSupValeur('OLDDATE',TEcr.GetValue('E_DATECOMPTABLE'));
                TInfoPiece.AddChampSupValeur('NUMLIGNE',1);
                TInfoPiece.AddChampSupValeur('NUMGROUPEECR',1);
            end
            else if (mode <> '-') then TInfoPiece.PutValue('NUMGROUPEECR',TInfoPiece.GetValue('NUMGROUPEECR') + 1);

            if (not bSelectionPiece) then bPieceUnique := true;
        end;

        TNewEcr := TOB.Create('ECRITURE',TInsert,-1);
        // Création de la nouvelle écritureisteEcr,-1);
        TNewEcr.Dupliquer(TEcr,false,true);
        // Chargement de l'analytique sous l'écriture
        if (TNewEcr.GetValue('E_ANA') = 'X') then ChargeAnalytique(TNewEcr);
        ExtourneLigneEcriture(TNewEcr,TInfoPiece,date,tipe,negatif,libelle,refint,refext,mode);
    end;

    if WithInsert then
    begin
      if (TInsert <> nil) then
      begin
        TInsert.SetAllModifie(true);
        result1 := result1 and TInsert.InsertDB(nil, true);
        MajSoldesEcritureTOB(TInsert, true); // TODO
      end;
      result2 := ListeEcriture.UpdateDB(true)
    end else
      result2 := true;

    TInsert.AddChampSup('NumErreur',false);
    TInsert.AddChampSup('Erreur',false);

    if (result1 and result2) then
    begin
        // tt c bien passé ;)
        TInsert.PutValue('NumErreur',0);
        TInsert.PutValue('Erreur','Aucune');
    end
    else
    begin
        numerror := 0;
        liberror := '';

        if (not (result1)) then
        begin
            numerror := numerror + 1;
            liberror := liberror + 'Problème durant l''Insertion des nouvelles écritures. ';
        end;
        if (not (result2)) then
        begin
            numerror := numerror + 2;
            liberror := liberror + 'Problème durant la Mise a jour des écritures. ';
        end;

        TInsert.PutValue('NumErreur',numerror);
        TInsert.PutValue('Erreur',liberror);
    end;

    result := TInsert;
end;

function TypeCptTVADefaut (CompteG, NatureG:string):string; //Lek FQ;001;26386
begin
 if EstChaPro(NatureG) then result:='HT'
 else if NatureG='IMO' then result:='HT'
 else if EstTVATPF(CompteG,True) then result:='TVA'
 else if EstTVATPF(CompteG,False) then result:='TVA'
 else if (NatureG='COC') or (NatureG='COS') or
         (NatureG='COF') or (NatureG='COD') or
         (NatureG='TIC') or (NatureG='TID') then result:='TTC'
 else result:='DIV';
end;


////////////////////////////////////////////////////////////////////////////////



Initialization
{$IFNDEF NOVH}
RegisterAglFunc('MPTOCATEGORIE',False,1,AGLMPToCategorie) ;
{$ENDIF}

end.
