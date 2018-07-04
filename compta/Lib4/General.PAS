unit General;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB, DBTables, Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, DBCtrls, Mask,
  hmsgbox, ComCtrls, ent1, Hcompte, HDB, hctrls, Spin, HEnt1, HQry, Region, ed_tools,
  Menus, Filtre ,
{$IFNDEF PGIIMMO}
  CritEdt,
{$ENDIF}
  CodePost, HRichEdt, HSysMenu, HRichOLE,MajTable,HRegCpte, ParamSoc,
  ADODB, udbxDataset, TntComCtrls, TntStdCtrls;


Procedure FicheGeneMZS(Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
Procedure FicheGene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
// CA - 18/12/2001
function FicheGeneCreateOne : string ;
// Fin CA - 18/12/2001
Procedure MajLettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
Procedure MajDelettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;

type
  TFCpteGene = class(TForm)
    Sgene    : TDataSource;
    DBNav    : TDBNavigator;
    HPB      : TPanel;
    BFerme       : TBitBtn;
    BFirst       : TBitBtn;
    BPrev        : TBitBtn;
    BNext        : TBitBtn;
    BLast        : TBitBtn;
    BImprimer    : TBitBtn;
    BValider     : TBitBtn;
    BAide        : TBitBtn;
    BVentil      : TBitBtn;
    BInsert      : TBitBtn;
    BRIB         : TBitBtn;
    MsgBox       : THMsgBox;
    FAutoSave    : TCheckBox;
    Pages        : TPageControl;
    PComptable   : TTabSheet;
    PComplement  : TTabSheet;
    PInformations: TTabSheet;
    TG_GENERAL          : THLabel;
    G_GENERAL           : TDBEdit;
    TG_NATUREGENE       : THLabel;
    G_NATUREGENE        : THDBValComboBox;
    TG_SENS             : THLabel;
    G_SENS              : THDBValComboBox;
    TG_ABREGE           : THLabel;
    G_VENTILABLE        : TDBCheckBox;
    G_LIBELLE           : TDBEdit;
    TG_LIBELLE          : THLabel;
    HGBOptionaxes       : TGroupBox;
    G_VENTILABLE1       : TDBCheckBox;
    G_VENTILABLE2       : TDBCheckBox;
    G_VENTILABLE3       : TDBCheckBox;
    G_VENTILABLE4       : TDBCheckBox;
    G_VENTILABLE5       : TDBCheckBox;
    HGBDERNMOUV         : TGroupBox;
    TG_DEBITDERNMVT     : THLabel;
    TG_CREDITDERNMVT    : THLabel;
    TG_NUMDERNMVT       : THLabel;
    G_DATEDERNMVT       : TDBEdit;
    G_DEBITDERNMVT      : TDBEdit;
    G_CREDITDERNMVT     : TDBEdit;
    G_NUMDERNMVT        : TDBEdit;
    G_LIGNEDERNMVT      : TDBEdit;
    TG_CORRESP1         : THLabel;
    G_BUDGENE: THDBCpteEdit;
    TG_BUDGENE: THLabel;
    TG_CORRESP2         : THLabel;
    HGBOptionImpression : TGroupBox;
    G_SOLDEPROGRESSIF   : TDBCheckBox;
    G_TOTAUXMENSUELS    : TDBCheckBox;
    G_SAUTPAGE          : TDBCheckBox;
    G_CENTRALISABLE     : TDBCheckBox;
    HGBDates            : TGroupBox;
    TG_DATEOUVERTURE    : THLabel;
    TG_DATEFERMETURE    : THLabel;
    TG_DATEMODIFICATION : THLabel;
    TG_DATECREATION     : THLabel;
    G_DATEMODIF: TDBEdit;
    G_DATEOUVERTURE     : TDBEdit;
    G_DATEFERMETURE     : TDBEdit;
    G_DATECREATION      : TDBEdit;
    G_FERME             : TDBCheckBox;
    TG_BLOCNOTE         : TGroupBox;
    G_BLOCNOTE: THDBRichEditOLE;
    PAdresse            : TTabSheet;
    TG_ADRESSE1         : THLabel;
    G_ADRESSE1          : TDBEdit;
    G_ADRESSE2          : TDBEdit;
    G_ADRESSE3          : TDBEdit;
    TG_CODEPOSTAL       : THLabel;
    G_CODEPOSTAL        : TDBEdit;
    TG_VILLE            : THLabel;
    G_VILLE             : TDBEdit;
    G_PAYS              : THDBValComboBox;
    TG_PAYS             : THLabel;
    G_DIVTERRIT         : TDBEdit;
    TG_DIVTERRIT        : THLabel;
    TG_TELEPHONE        : THLabel;
    G_TELEPHONE         : TDBEdit;
    G_TELEX             : TDBEdit;
    TG_TELEX            : THLabel;
    TG_FAX              : THLabel;
    G_FAX               : TDBEdit;
    TG_QUALIFQTE1       : THLabel;
    G_QUALIFQTE1        : THDBValComboBox;
    TG_QUALIFQTE2       : THLabel;
    G_QUALIFQTE2        : THDBValComboBox;
    TG_LANGUE           : THLabel;
    G_LANGUE            : THDBValComboBox;
    PInfo               : TTabSheet;
    TG_REGIMETVA        : THLabel;
    G_REGIMETVA         : THDBValComboBox;
    G_SOUMISTPF         : TDBCheckBox;
    TG_JOURPAIEMENT1    : THLabel;
    G_JOURPAIEMENT1     : TDBEdit;
    TG_JOURPAIEMENT2    : THLabel;
    G_JOURPAIEMENT2     : TDBEdit;
    G_MODEREGLE: THDBValComboBox;
    TG_MODEREGLE: THLabel;
    TT_TRESORERIE       : TGroupBox;
    TG_RELANCEREGLEMENT : THLabel;
    TG_RELANCETRAITE    : THLabel;
    TG_LETTREPAIEMENT   : THLabel;
    TG_RESIDENTETRANGER : THLabel;
    TG_NATUREECONOMIQUE : THLabel;
    TG_PLAFOND          : THLabel;
    TG_RISQUE           : THLabel;
    G_RELANCEREGLEMENT  : THDBValComboBox;
    G_RELANCETRAITE     : THDBValComboBox;
    G_LETTREPAIEMENT    : THDBValComboBox;
    G_RESIDENTETRANGER  : THDBValComboBox;
    G_NATUREECONOMIQUE  : THDBValComboBox;
    G_PLAFOND           : TDBEdit;
    G_RISQUE            : TDBEdit;
    TG_NIF              : THLabel;
    G_NIF               : TDBEdit;
    G_SIRET             : TDBEdit;
    TG_SIRET            : THLabel;
    TG_APE              : THLabel;
    G_APE               : TDBEdit;
    TG_JURIDIQUE        : THLabel;
    G_JURIDIQUE         : THDBValComboBox;
    TG_MOTIFVIREMENT    : THLabel;
    G_MOTIFVIREMENT     : THDBValComboBox;
    TG_TVAENCAISSEMENT  : THLabel;
    G_TVAENCAISSEMENT   : THDBValComboBox;
    GBExo               : TGroupBox;
    HLabel1             : THLabel;
    HLabel3             : THLabel;
    HLabel2             : THLabel;
    HLabel4             : THLabel;
    GBTVA               : TGroupBox;
    TG_TPF              : THLabel;
    TG_TVA              : THLabel;
    G_TPF               : THDBValComboBox;
    G_TVASURENCAISS: TDBCheckBox;
    TG_ADRESSE2         : THLabel;
    TG_ADRESSE3         : THLabel;
    G_TVA               : THDBValComboBox;
    Label1              : TLabel;
    GBCaract            : TGroupBox;
    G_COLLECTIF         : TDBCheckBox;
    G_POINTABLE         : TDBCheckBox;
    G_PURGEABLE         : TDBCheckBox;
    G_MODELE            : TDBCheckBox;
    G_RISQUETIERS       : TDBCheckBox;
    G_SUIVITRESO        : THDBValComboBox;
    TG_SUIVITRESO       : THLabel;
    G_LETTRABLE         : TDBCheckBox;
    GroupBox2           : TGroupBox;
    G_DERNLETTRAGE      : TDBEdit;
    BAnnuler            : TBitBtn;
    HLabel5             : THLabel;
    G_TOTDEBP           : TDBEdit;
    G_TOTCREP           : TDBEdit;
    G_SOLCREP           : THNumEdit;
    HLabel6             : THLabel;
    G_TOTDEBE           : TDBEdit;
    G_TOTCREE           : TDBEdit;
    G_SOLCREE           : THNumEdit;
    HLabel7             : THLabel;
    G_TOTDEBS           : TDBEdit;
    G_TOTCRES           : TDBEdit;
    G_SOLCRES           : THNumEdit;
    HLabel8             : THLabel;
    BMenuZoom           : TBitBtn;
    PopZ                : TPopupMenu;
    BGL                 : TBitBtn;
    BBL                 : TBitBtn;
    BCumP               : TBitBtn;
    BZecrimvt           : TBitBtn;
    BCumul              : TBitBtn;
    BJustif             : TBitBtn;
    QGene: TQuery;
    QGeneG_GENERAL: TStringField;
    QGeneG_LIBELLE: TStringField;
    QGeneG_ABREGE: TStringField;
    QGeneG_NATUREGENE: TStringField;
    QGeneG_CENTRALISABLE: TStringField;
    QGeneG_PURGEABLE: TStringField;
    QGeneG_SOLDEPROGRESSIF: TStringField;
    QGeneG_SAUTPAGE: TStringField;
    QGeneG_TOTAUXMENSUELS: TStringField;
    QGeneG_COLLECTIF: TStringField;
    QGeneG_DATECREATION: TDateTimeField;
    QGeneG_DATEOUVERTURE: TDateTimeField;
    QGeneG_DATEFERMETURE: TDateTimeField;
    QGeneG_FERME: TStringField;
    QGeneG_DATEDERNMVT: TDateTimeField;
    QGeneG_DEBITDERNMVT: TFloatField;
    QGeneG_CREDITDERNMVT: TFloatField;
    QGeneG_NUMDERNMVT: TIntegerField;
    QGeneG_LIGNEDERNMVT: TIntegerField;
    QGeneG_BLOCNOTE: TMemoField;
    QGeneG_SENS: TStringField;
    QGeneG_TOTALDEBIT: TFloatField;
    QGeneG_TOTALCREDIT: TFloatField;
    QGeneG_LETTRABLE: TStringField;
    QGeneG_POINTABLE: TStringField;
    QGeneG_VENTILABLE: TStringField;
    QGeneG_CORRESP1: TStringField;
    QGeneG_CORRESP2: TStringField;
    QGeneG_TVA: TStringField;
    QGeneG_RISQUETIERS: TStringField;
    QGeneG_DERNLETTRAGE: TStringField;
    QGeneG_DEBNONPOINTE: TFloatField;
    QGeneG_CREDNONPOINTE: TFloatField;
    QGeneG_UTILISATEUR: TStringField;
    QGeneG_TPF: TStringField;
    QGeneG_SOCIETE: TStringField;
    QGeneG_VENTILABLE1: TStringField;
    QGeneG_VENTILABLE2: TStringField;
    QGeneG_VENTILABLE3: TStringField;
    QGeneG_VENTILABLE4: TStringField;
    QGeneG_VENTILABLE5: TStringField;
    QGeneG_MODELE: TStringField;
    QGeneG_QUALIFQTE1: TStringField;
    QGeneG_QUALIFQTE2: TStringField;
    QGeneG_JOURPAIEMENT1: TIntegerField;
    QGeneG_JOURPAIEMENT2: TIntegerField;
    QGeneG_REGIMETVA: TStringField;
    QGeneG_SOUMISTPF: TStringField;
    QGeneG_ADRESSE1: TStringField;
    QGeneG_ADRESSE2: TStringField;
    QGeneG_ADRESSE3: TStringField;
    QGeneG_CODEPOSTAL: TStringField;
    QGeneG_VILLE: TStringField;
    QGeneG_PAYS: TStringField;
    QGeneG_LANGUE: TStringField;
    QGeneG_NIF: TStringField;
    QGeneG_SIRET: TStringField;
    QGeneG_APE: TStringField;
    QGeneG_TELEPHONE: TStringField;
    QGeneG_FAX: TStringField;
    QGeneG_TELEX: TStringField;
    QGeneG_RISQUE: TFloatField;
    QGeneG_PLAFOND: TFloatField;
    QGeneG_RELANCEREGLEMENT: TStringField;
    QGeneG_RELANCETRAITE: TStringField;
    QGeneG_RESIDENTETRANGER: TStringField;
    QGeneG_NATUREECONOMIQUE: TStringField;
    QGeneG_MOTIFVIREMENT: TStringField;
    QGeneG_LETTREPAIEMENT: TStringField;
    QGeneG_JURIDIQUE: TStringField;
    QGeneG_ETABLISSEMENT: TStringField;
    QGeneG_TOTDEBP: TFloatField;
    QGeneG_TOTCREP: TFloatField;
    QGeneG_TOTDEBE: TFloatField;
    QGeneG_TOTCREE: TFloatField;
    QGeneG_TOTDEBS: TFloatField;
    QGeneG_TOTCRES: TFloatField;
    QGeneG_TOTDEBANO: TFloatField;
    QGeneG_TOTCREANO: TFloatField;
    QGeneG_SUIVITRESO: TStringField;
    QGeneG_CONFIDENTIEL: TStringField;
    QGeneG_TVAENCAISSEMENT: TStringField;
    BBalGenAna: TBitBtn;
    BBALGENAUX: TBitBtn;
    G_ABREGE: TDBEdit;
    QPRef: THQuery;
    GBConfidentiel: TGroupBox;
    G_CONFIDENTIEL: TDBCheckBox;
    Bevel1: TBevel;
    Bevel3: TBevel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    HMTrad: THSystemMenu;
    QGeneG_MODEREGLE: TStringField;
    QGeneG_DIVTERRIT: TStringField;
    G_CORRESP1: THDBCpteEdit;
    G_CORRESP2: THDBCpteEdit;
    QGeneG_DATEMODIF: TDateTimeField;
    QGeneG_TVASURENCAISS: TStringField;
    ZL: TTabSheet;
    TG_TABLE0: THLabel;
    G_TABLE0: THDBCpteEdit;
    QGeneG_CREERPAR: TStringField;
    QGeneG_EXPORTE: TStringField;
    QGeneG_TABLE0: TStringField;
    QGeneG_TABLE1: TStringField;
    QGeneG_TABLE2: TStringField;
    QGeneG_TABLE3: TStringField;
    QGeneG_TABLE4: TStringField;
    QGeneG_TABLE5: TStringField;
    QGeneG_TABLE6: TStringField;
    QGeneG_TABLE7: TStringField;
    QGeneG_TABLE8: TStringField;
    QGeneG_TABLE9: TStringField;
    G_TABLE1: THDBCpteEdit;
    TG_TABLE1: THLabel;
    TG_TABLE2: THLabel;
    G_TABLE2: THDBCpteEdit;
    TG_TABLE3: THLabel;
    G_TABLE3: THDBCpteEdit;
    TG_TABLE4: THLabel;
    G_TABLE4: THDBCpteEdit;
    TG_TABLE5: THLabel;
    G_TABLE5: THDBCpteEdit;
    TG_TABLE6: THLabel;
    G_TABLE6: THDBCpteEdit;
    TG_TABLE7: THLabel;
    G_TABLE7: THDBCpteEdit;
    TG_TABLE8: THLabel;
    G_TABLE8: THDBCpteEdit;
    TG_TABLE9: THLabel;
    G_TABLE9: THDBCpteEdit;
    QGeneG_BUDGENE: TStringField;
    TG_CONSO: THLabel;
    QGeneG_CONSO: TStringField;
    G_CONSO: TDBEdit;
    H_0: THLabel;
    H_1: THLabel;
    H_2: THLabel;
    H_3: THLabel;
    H_4: THLabel;
    H_5: THLabel;
    H_6: THLabel;
    H_7: THLabel;
    H_8: THLabel;
    H_9: THLabel;
    BImmo: TBitBtn;
    QGeneG_EFFET: TStringField;
    G_EFFET: TDBCheckBox;
    procedure BLastClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SgeneDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
    procedure G_NATUREGENEChange(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure G_GENERALExit(Sender: TObject);
    procedure BRIBClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BCumulClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BZecrimvtClick(Sender: TObject);
    procedure BJustifClick(Sender: TObject);
    Procedure ActiveBVentil(Sender : TObject) ;
    procedure BMenuZoomClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure G_LETTRABLEClick(Sender: TObject);
    procedure G_VILLEDblClick(Sender: TObject);
    procedure G_CODEPOSTALDblClick(Sender: TObject);
    procedure BBalGenAnaClick(Sender: TObject);
    procedure BBALGENAUXClick(Sender: TObject);
    procedure G_GENERALEnter(Sender: TObject);
    procedure G_DIVTERRITDblClick(Sender: TObject);
    procedure G_PAYSChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BImmoClick(Sender: TObject);
    procedure G_GENERALChange(Sender: TObject);
    procedure PagesChanging(Sender: TObject; var AllowChange: Boolean);
    procedure G_NATUREGENEEnter(Sender: TObject);
  private    { Déclarations privées }
    Q : TQuery ;
    Lequel,LesModif,MemoCpte,LeCompte  : String ;
    AvecMvt                            : Boolean ;
    Mode                               : TActionFiche ;
    LaPage                             : Integer ;
    LgCode                             : Byte ;
    VentiPoint                         : Boolean ;
    MA0,MA1,MA2,MA3,MA4,MA5,MP,ML,MNat,MColl : String3 ;
    EnSaisieCpta                       : Boolean ;
    Import                             : Boolean ;
    MajPlanRef                         : Boolean ;
    CptR3PourInsert                    : String ;
    SectAtt : Array[1..5] of Boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Procedure NewEnreg ;
    Procedure ChargeEnreg ;
    Function  IsMouvemente : Boolean ;
    Function  Supprime     : Boolean ;
    Procedure DeleteVentilation ;
    Procedure SwapVentilation ;
    Procedure InitModif ;
    Procedure VerifiNatureValide ;
    Procedure MajCollectif ;
    Procedure RepositionneVentilation ;
    procedure GeneDAttente ;
    Procedure SwapNatureGene ;
    Function  CollectifDansTableTiers : Boolean ;
    Function  CollectifChanger : Boolean ;
    Function  CollectifChangerChaPro : Boolean ;
    Function  BqeCaiChanger : Boolean ;
    Function  ContrePartieDansJournal : Boolean ;
    Procedure MajVentilation(LeCpte : String) ;
    Function  EstVentilable(Quelaxe : Byte) : Boolean ;
    Procedure OnEstEnSaisieCompta ;
    Function  CaracteristiqueIncompatible(Var j : Byte) : Boolean ;
    Procedure VentilationAutorise ;
    Procedure LettrageAutorise ;
    Procedure PointageAutorise ;
    Function  LettrageChanger : Boolean ;
    Function  VentilChanger : Boolean ;
    Procedure SupprimePreventilRib ;
    Function  CodeValide : Boolean ;
    Function  TvaTicTidOk : Boolean ;
    Procedure EnabledTva ;
    Function  OkRIB(Nat : String ; Var OkBqe : Boolean) : Boolean ;
    Procedure DelRibContact ;
{===== pour les modif en série ================================}
    Procedure InitModifEnSerie(StModif : string) ;
    Procedure AffecteLe(Champ, Valeur : string) ;
    Function  OkConfidentiel : Boolean ;
    Procedure CreerRibDefaut ;
    Procedure ChercheEtImportCptePlanRef ;
    Procedure MajDuPlanReference ;
    Procedure ChercheSectAttente ;
    Function  VentilationModifier : Boolean ;
    Procedure FormateLesMontants ;
    Function  VerifCoherenceTL : Boolean ;
    procedure Ergo35 ;
    procedure ActivationControl;

  public
    CptesCrees : String ;
  end;

    Function  EstEcrGuide(LeCpte : String)     : Boolean ;
    Function  EstAnalytiqPure(LeCpte : String) : Boolean ;
    Function  EstCpteAxe(LeCpte : String)      : Boolean ;
    Function  EstBanquaire(LeCpte : String)    : Boolean ;
    Function  EstCpteCorresp(LeCpte : String)  : Boolean ;
    Function  EstCpteDevise(LeCpte : String)   : Boolean ;
    Function  EstCpteJournal(LeCpte : String)  : Boolean ;
    Function  EstCpteModepaie(LeCpte : String) : Boolean ;
    Function  EstCpteSociete(LeCpte : String)  : Boolean ;
    Function  EstCpteTva(LeCpte : String)      : Boolean ;
    Function  EstCpteAmortissement(LeCpte : String) : boolean;

implementation


uses Ventil, FichComm,
 {$IFDEF TRSYNCHRO}
 BANQUECP_TOM,
 {$ELSE}
 banqueCP,
 {$ENDIF}
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
     QRPLANGE, CumMens, Zecrimvt, Justisol,QRGLGen, QRBalGen,
     {$IFNDEF CCS3} QRCumulG,{$ENDIF}
{$IFNDEF CCMP}
     QRBLGESE, MulImmo, Outils,
{$ENDIF}
     QRBAGEAU,
{$ENDIF}
{$ENDIF}
     Choix, SaisUtil, SaisComm,
     LetBatch,
{$IFNDEF PGIIMMO}
     UtilEdt,
     CpteSav,
     ImpPrefG,
{$ENDIF}
     UtilPGI ;

{$R *.DFM}


Procedure FicheGene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
var FCpteGene: TFCpteGene;
BEGIN
Case Comment of
   taCreat,taCreatEnSerie,
             taCreatOne : BEGIN
                          if Not ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then Exit ;
                          END ;
 taModif,taModifEnSerie : BEGIN
                          if Not ExJaiLeDroitConcept(TConcept(ccGenModif),True) then Exit ;
                          if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrEnca','nrDeca'],True,'nrBatch') then Exit ;
                          END ;
   END ;
FCpteGene:=TFCpteGene.Create(Application) ;
 try
   FCpteGene.Q:=Q ; FCpteGene.CptesCrees:='' ;
   FCpteGene.lequel:=Compte ;
   FCpteGene.Mode:=Comment ;
   FCpteGene.LaPage:=QuellePage ;
   FCpteGene.ShowModal ;
//   VH^.CptesCrees:=FCpteGene.CptesCrees ;
 Finally
   FCpteGene.free ;
   if Comment in [taModif,taModifEnSerie] then _Bloqueur('nrBatch',False) ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

// CA - 18/12/2001
function FicheGeneCreateOne : string ;
var FCpteGene: TFCpteGene;
begin
  if Not ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then Exit ;
  FCpteGene:=TFCpteGene.Create(Application) ;
  try
    FCpteGene.Q:=nil ; FCpteGene.CptesCrees:='' ;
    FCpteGene.lequel:='' ;
    FCpteGene.Mode:=taCreatOne ;
    FCpteGene.LaPage:=0 ;
    FCpteGene.ShowModal ;
  finally
    Result := FCpteGene.CptesCrees;
    FCpteGene.free ;
  end;
  Screen.Cursor:=SyncrDefault ;
END ;
// Fin CA - 18/12/2001

Procedure FicheGeneMZS(Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
var FCpteGene: TFCpteGene;
BEGIN
Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then Exit ;
              taModif,taModifEnSerie : if Not ExJaiLeDroitConcept(TConcept(ccGenModif),True) then Exit ;
   END ;
FCpteGene:=TFCpteGene.Create(Application) ;
 try
   FCpteGene.lequel:=Compte ;
   FCpteGene.Mode:=Comment ;
   FCpteGene.LaPage:=QuellePage ;
   FCpteGene.LesModif:=LesModif ;
   FCpteGene.ShowModal ;
 Finally
   FCpteGene.free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFCpteGene.BImprimerClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
PlanGeneral(G_GENERAL.Text,True) ;
{$ENDIF}
{$ENDIF}
end;

Function TFCpteGene.IsMouvemente : Boolean ;
BEGIN
if ((Mode=taConsult) or (QGene.State=dsInsert)) then BEGIN Result:=False ; Exit ; END ;
Result:=ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+G_GENERAL.Text+'" '+
                  ' AND ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+G_GENERAL.Text+'"))'+
                  ' Or (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL="'+G_GENERAL.Text+'")))') ;
END ;

Function TFCpteGene.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if QGene.Modified then
   BEGIN
   if (Mode in [taCreat..taCreatOne]) And
      (Bourreless(QGeneG_GENERAL.asString,fbGene)='') then Rep:=mrNo
   else if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then Exit ;
  mrNo :  BEGIN DelRibContact ; if not Bouge(nbCancel) then Exit ; END ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFCpteGene.CodeValide : Boolean ;
BEGIN
Result:=True ;
if QGene.state in [dsInsert] then
   BEGIN
   if QGeneG_GENERAL.AsString='' then
      BEGIN
      Pages.ActivePage:=PComptable ; G_GENERAL.SetFocus ;
      MsgBox.execute(2,'','') ; Result:=False ; Exit ;
      END ;
   if Presence('GENERAUX','G_GENERAL',QGeneG_GENERAL.AsString) then
      BEGIN
      Pages.ActivePage:=PComptable ; G_GENERAL.SetFocus ;
      MsgBox.execute(4,'','') ; Result:=False ; Exit ;
      END ;
   END ;
END ;

Procedure MajCollChaProEcriture(Cpte,Aux : String) ;
Var LeWhere,LeSet,Mp,Mr : String ;
    Q : TQuery ;
    OkLet : Boolean ;

BEGIN
if not VH^.CPIFDEFCEGID then exit ;
LeSet:='' ;
If Aux<>'' Then
  BEGIN
  OkLet:=FALSE ; Mr:='' ; Mp:='' ;
  Q:=OpenSQL('SELECT T_LETTRABLE, T_MODEREGLE FROM TIERS WHERE T_AUXILIAIRE="'+Aux+'" ',TRUE) ;
  If Not Q.Eof Then BEGIN OkLet:=Q.Fields[0].AsString='X' ; Mr:=Q.Fields[1].AsString ; END ;
  Ferme(Q) ;
  Q:=OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+MR+'" ',TRUE) ;
  If Not Q.Eof Then Mp:=Q.Fields[0].AsString ;
  Ferme(Q) ;
  If OkLet Then LeSet:=', E_MODEPAIE="'+Mp+'", E_DATEECHEANCE=E_DATECOMPTABLE, E_ETATLETTRAGE="AL", E_NUMECHE=1, E_ECHE="X" ' ;
  END Else LeSet:=', E_MODEPAIE="", E_ETATLETTRAGE="RI", E_NUMECHE=0, E_ECHE="-", E_COUVERTURE=0, E_COUVERTUREDEV=0, E_LETTRAGEDEV="-" ' ;
LeWhere:=' WHERE E_GENERAL="'+Cpte+'" ' ;
// GG COM
If OkSynChro Then LeSet:=LeSet+' , E_PAQUETREVISION=1 ' ;
BeginTrans ;
ExecuteSql('UPDATE ECRITURE SET E_AUXILIAIRE="'+Aux+'" '+LeSet+LeWhere) ;
CommitTrans ;
END ;

Function TFCpteGene.EnregOK : boolean ;
Var ModPaie : String ;
    j : Byte ;
    AuxDiv : String ;
BEGIN
Result:=FALSE  ;
if QGene.state in [dsInsert,dsEdit]=False then Exit ;
j:=0 ;
if Not CodeValide then Exit ;
if QGene.state in [dsInsert,dsEdit] then
   BEGIN
   if QGeneG_LIBELLE.asString='' then
      BEGIN
      Pages.ActivePage:=PComptable ; G_LIBELLE.SetFocus ;
      MsgBox.execute(3,'','') ; Exit ;
      END ;
   if QGeneG_NATUREGENE.asString='' then
      BEGIN
      Pages.ActivePage:=PComptable ; G_NATUREGENE.SetFocus ;
      MsgBox.execute(47,'','') ; Exit ;
      END ;
   if CaracteristiqueIncompatible(j) then
      BEGIN
      Case j of
           1:BEGIN MsgBox.Execute(25,'','') ; Exit ; END ; //Ventilable Pointable Lettrable
           2:BEGIN MsgBox.Execute(26,'','') ; Exit ; END ; //Ventilable Lettrable
           3:BEGIN MsgBox.Execute(27,'','') ; Exit ; END ; //Pointable Lettrable
           4:BEGIN MsgBox.Execute(28,'','') ; Exit ; END ; //Ventilable Pointable
        End ;
      END ;
   if Not TvaTicTidOk then Exit ;
   if Not CoherencePaysRegion('G_',QGene) then
      BEGIN
      Pages.ActivePage:=PAdresse ; G_PAYS.SetFocus ;
      MsgBox.execute(43,'','') ; Exit ;
      END ;
   if Not VerifCoherenceTL then Exit ;
   END ;
Case VerifCorrespondance(1,QGENEG_CORRESP1.AsString,QGENEG_CORRESP2.AsString) of
    0 : ;
    1 : BEGIN Msgbox.Execute(41,'','') ; Exit ; END ;
    2 : BEGIN Msgbox.Execute(42,'','') ; Exit ; END ;
    END ;
if QGene.state in [dsEdit] then
   BEGIN
   if CollectifChanger then
      BEGIN
      if CollectifDansTableTiers then
         BEGIN
         MsgBox.Execute(20,'','') ; QGeneG_NATUREGENE.AsString:=MNat ;
         END else
         BEGIN
         if AvecMvt then BEGIN MsgBox.Execute(22,'','') ; QGeneG_NATUREGENE.AsString:=MNat ; END ;
         END ;
      END ;
   if BqeCaiChanger then
      BEGIN
      if ContrePartieDansJournal then BEGIN MsgBox.Execute(21,'','') ; QGeneG_NATUREGENE.AsString:=MNat ; END ;
      END ;
   if LettrageChanger then
      BEGIN
      if ((Not _BlocCarFiche) and (MsgBox.Execute(18,'','')=mrYes)) then
         BEGIN
         MsgBox.Execute(23,'','') ; ModPaie:=Choisir(MsgBox.Mess[29],'MODEPAIE','MP_LIBELLE','MP_MODEPAIE','','') ;
         if ModPaie<>'' then MajLettrageEcriture(ModPaie,G_GENERAL.Text,fbGene) else
            BEGIN
            MsgBox.Execute(24,'','') ; QGeneG_LETTRABLE.AsString:=ML ;
            END ;
         END else QGeneG_LETTRABLE.AsString:=ML ;
      END ;
if VH^.CPIFDEFCEGID then
Begin
   If CollectifChangerChaPro Then
     BEGIN
     if ((Not _BlocCarFiche) and (MsgBox.Execute(50,'','')=mrYes)) then
        BEGIN
        AuxDiv:='' ;
        If Not G_COLLECTIF.Checked Then
          BEGIN
          MajCollChaProEcriture(G_GENERAL.Text,AuxDiv) ; If G_COLLECTIF.Checked Then MColl:='X' Else MColl:='-' ;
          END Else
          BEGIN
          MsgBox.Execute(51,'','') ; AuxDiv:=Choisir(MsgBox.Mess[52],'TIERS','T_AUXILIAIRE','T_AUXILIAIRE','','T_AUXILIAIRE') ;
          If AuxDiv<>'' Then
              BEGIN
              MajCollChaProEcriture(G_GENERAL.Text,AuxDiv) ; If G_COLLECTIF.Checked Then MColl:='X' Else MColl:='-' ;
              END else
              BEGIN
              MsgBox.Execute(53,'','') ; QGeneG_COLLECTIF.AsString:=MColl ;
              END ;
          END ;
        END else QGeneG_Collectif.AsString:=MColl ;
     END ;
END ;
   if VentilChanger then
      BEGIN
      if ((Not _BlocCarFiche) and (MsgBox.Execute(19,'','')=mrYes))
          then MajVentilation(G_GENERAL.Text) else RepositionneVentilation ;
      END ;
   END ;
DateModification(QGene,'G') ; MajCollectif ;
if(G_VENTILABLE1.Checked)Or(G_VENTILABLE2.Checked)Or(G_VENTILABLE3.Checked)Or
  (G_VENTILABLE4.Checked)Or(G_VENTILABLE5.Checked)then QGeneG_VENTILABLE.AsString:='X'
                                                  else QGeneG_VENTILABLE.AsString:='-' ;
Result:=TRUE  ;
END ;

Procedure TFCpteGene.SupprimePreventilRib ;
BEGIN
if AvecMvt then Exit ;
if QGene.State<>dsBrowse then Exit ;
if MemoCpte='' then Exit ;
if QGeneG_VENTILABLE.AsString='-'then DeleteVentilation ;
if QGeneG_NATUREGENE.AsString<>'BQE' then ExecuteSql('Delete From BANQUECP Where BQ_GENERAL="'+MemoCpte+'"') ;
END ;

Function TFCpteGene.LettrageChanger : Boolean ;
Var AvecMvtTemp,AeteModifier : Boolean ;
BEGIN
Result:=False ; AvecMvtTemp:=False ;
if ((ML='-') And (Not(G_LETTRABLE.Checked)) Or
   ((ML='X') And (G_LETTRABLE.Checked))) then AeteModifier:=False else AeteModifier:=True ;
if AeteModifier And (Not AvecMvt) then AvecMvtTemp:=IsMouvemente ;
if AvecMvtTemp then
   BEGIN
   if (ML='X') And (Not(G_LETTRABLE.Checked)) then
       BEGIN
       MsgBox.Execute(39,'','') ; AvecMvt:=True ; QGeneG_LETTRABLE.AsString:=ML ; Exit ;
       END ;
   AvecMvt:=True ;
   END ;
if Not AvecMvt then Exit ;
if QGene.State in [dsBrowse,dsInsert] then Exit ;
Result:=((ML='-')And(G_LETTRABLE.Checked)) ;
END ;

Function TFCpteGene.CollectifChangerChaPro : Boolean ;
Var AvecMvtTemp,AeteModifier : Boolean ;
BEGIN
if not VH^.CPIFDEFCEGID then exit ;
Result:=False ;
If (G_NATUREGENE.Value<>'CHA') And (G_NATUREGENE.Value<>'PRO') Then Exit ;
AvecMvtTemp:=False ;
if ((MColl='-') And (Not(G_COLLECTIF.Checked)) Or
   ((MColl='X') And (G_COLLECTIF.Checked))) then AeteModifier:=False else AeteModifier:=True ;
if AeteModifier And (Not AvecMvt) then AvecMvtTemp:=IsMouvemente ;
if AvecMvtTemp then AvecMvt:=True ;
if Not AvecMvt then Exit ;
if QGene.State in [dsBrowse,dsInsert] then Exit ;
Result:=((MColl='-')And(G_COLLECTIF.Checked)) Or ((MColl='X')And(Not G_COLLECTIF.Checked));
END ;

Procedure TFCpteGene.RepositionneVentilation ;
BEGIN
QGeneG_VENTILABLE.AsString:=MA0  ; QGeneG_VENTILABLE1.AsString:=MA1 ;
QGeneG_VENTILABLE2.AsString:=MA2 ; QGeneG_VENTILABLE3.AsString:=MA3 ;
QGeneG_VENTILABLE4.AsString:=MA4 ; QGeneG_VENTILABLE5.AsString:=MA5 ;
END ;

Procedure TFCpteGene.InitModif ;
BEGIN
MA0:=QGeneG_VENTILABLE.AsString  ; MA1:=QGeneG_VENTILABLE1.AsString ;
MA2:=QGeneG_VENTILABLE2.AsString ; MA3:=QGeneG_VENTILABLE3.AsString ;
MA4:=QGeneG_VENTILABLE4.AsString ; MA5:=QGeneG_VENTILABLE5.AsString ;
MP:=QGeneG_POINTABLE.AsString    ; ML:=QGeneG_LETTRABLE.AsString    ;
MColl:=QGeneG_COLLECTIF.AsString ;
MNat:=QGeneG_NATUREGENE.AsString ;
END ;

Procedure TFCpteGene.OnEstEnSaisieCompta ;
Var i : Integer ;
BEGIN
if Not EnSaisieCpta then Exit ;
G_NATUREGENE.Enabled:=False ; G_LETTRABLE.Enabled:=False ; G_POINTABLE.Enabled:=False ;
for i:=1 to 5 do TDBCheckBox(FindComponent('G_VENTILABLE'+IntToStr(i)+'')).Enabled:=False ;
BInsert.Enabled:=False ;
END ;

Procedure TFCpteGene.EnabledTva ;
Var l : Integer ;
BEGIN
For l:=0 To GBTVA.ControlCount-1 Do GBTVA.Controls[l].Enabled:=GBTVA.Enabled ;
END ;

Function TFCpteGene.OkRIB(Nat : String ; Var OkBqe : Boolean) : Boolean ;
BEGIN
OkBqe:=(Nat='BQE') ; Result:=(Nat='BQE') Or (Nat='TID') Or (Nat='TIC')  ;
END ;

Procedure TFCpteGene.ChargeEnreg ;
Var OkBqe : Boolean ;
BEGIN
if LeCompte=QGENEG_GENERAL.AsString then
   BEGIN
   if Mode=taConsult then FicheReadOnly(Self) ;
   Exit ;
   END ;
LeCompte:=QGENEG_GENERAL.AsString ;

InitCaption(Self,G_GENERAL.text,G_LIBELLE.text) ;
AfficheLeSolde(G_SOLCREP,QGeneG_TOTDEBP.AsFloat,QGeneG_TOTCREP.AsFloat) ;
AfficheLeSolde(G_SOLCREE,QGeneG_TOTDEBE.AsFloat,QGeneG_TOTCREE.AsFloat) ;
AfficheLeSolde(G_SOLCRES,QGeneG_TOTDEBS.AsFloat,QGeneG_TOTCRES.AsFloat) ;
BVentil.Enabled:=(G_Ventilable.Checked) ;
BJustif.Enabled:=(G_LETTRABLE.Checked) ;
{$IFNDEF IMP}
{$IFDEF SANSIMMO}
BImmo.Enabled:=FALSE ;
{$ELSE}
BImmo.Enabled:=((VH^.OkModImmo) and (G_NATUREGENE.Value='IMO')) ;
{$ENDIF}
//BImmo.Enabled:=False ;
{$ENDIF}
BBalGenAna.Enabled:=BVentil.Enabled ;
BBalGenAux.Enabled:=G_COLLECTIF.Checked ;
GBTVA.Enabled:=(G_NATUREGENE.Value='IMO')Or(G_NATUREGENE.Value='CHA')Or(G_NATUREGENE.Value='PRO') ;
EnabledTva ;
//GPRIB BRIB.Enabled:=(G_NATUREGENE.Value='BQE') ;
BRIB.Enabled:=OkRib(G_NATUREGENE.Value,OkBqe) ;
//PAdresse.TabVisible:=(QGeneG_NATUREGENE.asString='TID')or(QGeneG_NATUREGENE.asString='TIC') ;
PAdresse.TabVisible:=OkRib(QGeneG_NATUREGENE.AsString,OkBqe) ; If OkBqe Then PAdresse.TabVisible:=FALSE ;
PInfo.TabVisible:=(QGeneG_NATUREGENE.asString='TID')or(QGeneG_NATUREGENE.asString='TIC') ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
G_GENERAL.Enabled:=FALSE ;
AvecMvt:=IsMouvemente ; InitModif ; SwapNatureGene ; OnEstEnSaisieCompta ;
GeneDAttente ;
{ AJOUT ME LE 18/04/2001 }
//RR préparation version S3 Premium.
if ( ((ctxPCL in V_PGI.PGIContexte) or (EstComptaSansAna)) and (GetParamSoc('SO_CPPCLSANSANA')=TRUE)) then
begin
   HGBOptionaxes.enabled := FALSE;
   HGBOptionaxes.tabstop :=False;
   G_VENTILABLE1.enabled := FALSE;
   G_VENTILABLE2.enabled := FALSE;
   // CA - 13/02/2002 - suite ouverture des axes 3 à 5, ne pas permettre la modif. si ana pas ouvert.
   G_VENTILABLE3.enabled := FALSE;
   G_VENTILABLE4.enabled := FALSE;
   G_VENTILABLE5.enabled := FALSE;
end;
END ;

procedure TFCpteGene.GeneDAttente ;
Var i : Byte ;
BEGIN
if (QGeneG_GENERAL.AsString=VH^.Cpta[fbGene].Attente) then
   BEGIN
   G_NATUREGENE.Enabled:=False ; G_VENTILABLE.Enabled:=False ; G_LETTRABLE.Enabled:=False ;
   G_POINTABLE.Enabled:=False ;
   for i:=1 to 5 do TDBCheckBox(FindComponent('G_VENTILABLE'+IntToStr(i)+'')).Enabled:=False ;
   END ;
END ;

Procedure TFCpteGene.NewEnreg ;
BEGIN
InitNew(QGene) ;
AvecMvt:=False ;
G_NATUREGENE.Enabled:=True ;
// G_NATUREGENE.Value:='DIV' ;
QGeneG_SUIVITRESO.AsString:='RIE' ;
QGeneG_PURGEABLE.AsString:='X' ;
QGeneG_SOLDEPROGRESSIF.AsString:='X' ;
QGeneG_CONFIDENTIEL.AsInteger:=0 ;
QGeneG_SENS.AsString:='M' ;
QGeneG_MODELE.AsString:='-' ;
QGeneG_RISQUETIERS.AsString:='-' ;
Pages.ActivePage:=PComptable ;
G_GENERAL.Enabled:=TRUE ; G_GENERAL.Setfocus ;
LgCode:=VH^.Cpta[fbGene].Lg ;
//DateCreation(QGene,'G') ;
END ;

Function TFCpteGene.Bouge(Button: TNavigateBtn) : boolean ;
Var OkCreat : boolean ;
    OkBqe : Boolean ;
BEGIN
Result:=FALSE  ; OkCreat:=False ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if (QGene.State<>dsBrowse) And (OkRib(QGeneG_NATUREGENE.AsString,OkBqe)) And(Button<>nbCancel) And OkBqe then BribClick(Nil) ;
{ Modif CA le 13/03/2001 - Pas de maj de planref dans le cas PCL }
if (not (ctxPCL in V_PGI.PGIContexte)) And (QGene.State in [dsEdit,dsInsert]) And (MajPlanRef) And Not(Button in [nbCancel,nbDelete])then
  MajDuPlanReference ;
OkCreat:=(QGene.State in [dsInsert]) ;
if Not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[31]) else
   if OkCreat then
      BEGIN
      if ((QGeneG_GENERAL.AsString<>'') and (Button<>nbCancel) and (Pos(QGeneG_GENERAL.AsString+';',CptesCrees)<=0))
         then CptesCrees:=CptesCrees+QGeneG_GENERAL.AsString+';' ;
      END ;
Result:=TRUE ;
MemoCpte:=QGeneG_GENERAL.AsString ; SupprimePreventilRib ; MemoCpte:='' ;
if Button=NbInsert then NewEnreg ;
END ;

procedure TFCpteGene.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCpteGene.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCpteGene.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCpteGene.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCpteGene.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then begin Bouge(nbInsert) ; exit ; end ;
   if ((V_PGI.MonoFiche) or (Mode=taCreatOne) or (Mode=taModifEnSerie)) then Close ;
   END ;
end;

procedure TFCpteGene.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFCpteGene.BInsertClick(Sender: TObject);
begin
if ExJaiLeDroitConcept(TConcept(ccGenCreat),True) then Bouge(nbInsert) ;
end;

procedure TFCpteGene.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFCpteGene.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFCpteGene.SgeneDataChange(Sender: TObject; Field: TField);
var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=(Not(QGene.State in [dsEdit,dsInsert])) ;
BMenuZoom.Enabled:=(Not(QGene.State in [dsInsert])) ;
BImprimer.Enabled:=(Not(QGene.State in [dsInsert])) ;

// GCO - 17/04/2002
if ( CtxPcl in V_Pgi.PGIContexte ) then
  BEGIN
  G_NatureGene.Enabled := Not IsMouvemente;
  if G_NatureGene.Enabled then
    G_NatureGene.Color := ClWindow
  else
    G_NatureGene.Color :=   ClBtnFace;
  END ;
// FIn - GCO

if EnSaisieCpta then BInsert.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not QGene.BOF;
   DnEnable := Enabled and not QGene.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
   if ((Field.FieldName='G_LIBELLE') and (G_ABREGE.Field.AsString='')) then G_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFCpteGene.Ergo35 ;
Var i : integer ;
    CC : THCpteEdit ;
    LL : THLabel ;
BEGIN
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
//   G_VENTILABLE3.Visible:=False ; G_VENTILABLE4.Visible:=False ; G_VENTILABLE5.Visible:=False ;
   G_CORRESP2.Visible:=False ; TG_CORRESP2.Visible:=False ;
   if EstSerie(S3) then
      BEGIN
// JLD 5AXE déplacés en s3
      G_VENTILABLE3.Visible:=False ; G_VENTILABLE4.Visible:=False ; G_VENTILABLE5.Visible:=False ;
      G_VENTILABLE2.Visible:=False ;
      G_CORRESP1.Visible:=False ; TG_CORRESP1.Visible:=False ;
      G_QUALIFQTE1.Visible:=False ; TG_QUALIFQTE1.Visible:=False ;
      G_QUALIFQTE2.Visible:=False ; TG_QUALIFQTE2.Visible:=False ;
      G_CONSO.Visible:=False ; TG_CONSO.Visible:=False ;
      HGBOptionAxes.Caption:=MsgBox.Mess[48] ;
      for i:=2 to 5 do TCheckBox(FindComponent('G_VENTILABLE'+IntToStr(i))).Visible:=False ;
      G_VENTILABLE1.Caption:=MsgBox.Mess[49] ;
      G_VENTILABLE1.Width:=G_VENTILABLE3.Left-G_VENTILABLE1.Left ;
      G_TPF.Visible:=False ; TG_TPF.Visible:=False ;
      G_SOUMISTPF.Visible:=False ;
      G_LANGUE.Visible:=False ; TG_LANGUE.Visible:=False ;
      G_PURGEABLE.Visible:=False ; (*G_TVASURENCAISS.Visible:=False ; *)
      GBConfidentiel.Visible:=False ;   
      END ;
   END ;
if ZL.TabVisible then
   BEGIN
   for i:=0 to 9 do
       BEGIN
       CC:=THCpteEdit(FindComponent('G_TABLE'+IntTostr(i))) ;
       if CC<>Nil then
          BEGIN
          if CC.Text<>'' then CC.ExisteH else
             BEGIN
             LL:=THLabel(FindComponent('H_'+IntToStr(i))) ;
             if LL<>Nil then LL.Caption:='' ;
             END ;
          END ;
       END ;
   END ;
END ;

procedure TFCpteGene.FormShow(Sender: TObject);
begin
MakeZoomOLE(handle) ;
if VH^.CPIFDEFCEGID then
Begin
G_COLLECTIF.Top:=G_LETTRABLE.Top ; G_COLLECTIF.Left:=G_LETTRABLE.Left ;
G_COLLECTIF.Height:=G_LETTRABLE.Height ; G_COLLECTIF.Width:=G_LETTRABLE.Width ;
G_COLLECTIF.Color:=G_LETTRABLE.Color ; G_COLLECTIF.Font.Color:=G_LETTRABLE.Font.Color ;
end ;
ChercheSectAttente ; RecupWhereSQL(Q,QGene) ;
if ((Q=NIL) and (Lequel<>'')) then QGene.SQL.Add('Where G_GENERAL="'+Lequel+'"') ;
if (Q=NIL) then QGene.SQL.Add('Order by G_GENERAL') ;//lhe le 19/02/97 pour trier [tacreat..tacreatenserie] avec navigation
if ((Q=NIL) and (Lequel<>'')) then ChangeSQL(QGene) ;
CptR3PourInsert:='' ;
ChangeSizeMemo(QGeneG_BLOCNOTE) ;
QGene.Open ; FormateLesMontants ;
Pages.ActivePage:=Pages.Pages[LaPage] ;
VentiPoint:=True ; ChangeSQL(QPRef) ; //QPRef.Prepare ;
PrepareSQLODBC(QPRef) ;
Import:=False ; MajPlanRef:=True ;
if (not V_PGI.MonoFiche) and (Lequel<>'') And ((Mode in [taCreat..taCreatOne])=False) then
   BEGIN
   if Not QGene.Locate('G_GENERAL',Lequel,[]) then
       BEGIN MessageAlerte(MsgBox.Mess[31]) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; END ;
   END ;
if not OkConfidentiel then Exit ;
if Mode<>taConsult then BEGIN LgCode:=VH^.Cpta[fbGene].Lg ; END ;

EnSaisieCpta:=False ;
Case Mode Of
     taConsult           : BEGIN FicheReadOnly(Self) ; (* BVentil.Enabled:=False ; *) END ;
     taCreat..taCreatOne : BEGIN
                           Bouge(nbInsert) ; G_GENERAL.Text:=Lequel ;
                           BAnnuler.Enabled:=False ;
                           END ;
     taModif             : EnSaisieCpta:=SaisieLancer ;
     taModifEnSerie      : InitModifEnSerie(LesModif);
     end ;
OnEstEnSaisieCompta ;
BFirst.Visible := not V_PGI.MonoFiche ;
BPrev.Visible := not V_PGI.MonoFiche ;
BNext.Visible := not V_PGI.MonoFiche ;
BLast.Visible := not V_PGI.MonoFiche ;
Ergo35 ;
end;

Procedure TFCpteGene.VerifiNatureValide ;
BEGIN
if Not AvecMvt then Exit ;
if MNat=G_NATUREGENE.Value then Exit ;
if(MNat='COC')Or(MNat='COD')Or(MNat='COF')Or(MNat='COS')then
  if((G_NATUREGENE.Value='BQE')Or(G_NATUREGENE.Value='CAI')Or
     (G_NATUREGENE.Value='CHA')Or(G_NATUREGENE.Value='DIV')Or
     (G_NATUREGENE.Value='EXT')Or(G_NATUREGENE.Value='IMO')Or
     (G_NATUREGENE.Value='PRO')Or(G_NATUREGENE.Value='TIC')Or
     (G_NATUREGENE.Value='TID'))then BEGIN G_NATUREGENE.Value:=MNat ; Exit ; END ;
if(MNat<>'COC')And(MNat<>'COD')And(MNat<>'COF')And(MNat<>'COS')then
  if(G_NATUREGENE.Value<>'DIV') then BEGIN G_NATUREGENE.Value:=MNat ; Exit ; END ;
END ;

Procedure TFCpteGene.PointageAutorise ;
BEGIN
if(G_NATUREGENE.Value='DIV')Or(G_NATUREGENE.Value='EXT')Or
  (G_NATUREGENE.Value='BQE')Or(G_NATUREGENE.Value='CAI')then G_POINTABLE.Enabled:=True
else BEGIN
     G_POINTABLE.Enabled:=False ;
     if QGene.State in [dsEdit,dsInsert] then
        BEGIN QGeneG_POINTABLE.AsString:='-' ; G_POINTABLE.Checked:=False ; END ;
     END ;
if Not AvecMvt   then Exit ;
if Mode<>taModif then Exit ;
if EnSaisieCpta  then Exit ;
if G_POINTABLE.Checked then G_POINTABLE.Enabled:=False ;
END ;

Procedure TFCpteGene.LettrageAutorise ;
BEGIN
//if(G_NATUREGENE.Value='DIV')Or(G_NATUREGENE.Value='EXT')Or //Voir fiche de bug MPP N°869
if VH^.CPIFDEFCEGID then
Begin
G_COLLECTIF.Visible:=FALSE ; G_LETTRABLE.Visible:=TRUE ;
End ;
if(G_NATUREGENE.Value='TIC')Or(G_NATUREGENE.Value='TID')then G_LETTRABLE.Enabled:=True
else BEGIN
     G_LETTRABLE.Enabled:=False ;
     if QGene.State in [dsEdit,dsInsert] then
        BEGIN QGeneG_LETTRABLE.AsString:='-' ; G_LETTRABLE.Checked:=False ; END ;
     END ;
if VH^.CPIFDEFCEGID then
Begin
if(G_NATUREGENE.Value='PRO') Or (G_NATUREGENE.Value='CHA')then
  BEGIN
  If Not G_LETTRABLE.Checked Then
    BEGIN
    G_COLLECTIF.Visible:=TRUE ; G_LETTRABLE.Visible:=False ; G_COLLECTIF.Enabled:=TRUE ;
    END ;
  END ;
End ;
if Not AvecMvt   then Exit ;
if Mode<>taModif then Exit ;
if EnSaisieCpta  then Exit ;
if G_LETTRABLE.Checked then G_LETTRABLE.Enabled:=False ;
if VH^.CPIFDEFCEGID then
Begin
//If G_COLLECTIF.Checked Then G_COLLECTIF.Enabled:=False ;
end ;
END ;

Procedure TFCpteGene.VentilationAutorise ;
Var i : Byte ;
    C : TComponent ; 
BEGIN
if Not AvecMvt   then Exit ;
if Mode<>taModif then Exit ;
if EnSaisieCpta  then Exit ;
if G_VENTILABLE.Checked then
   BEGIN
   for i:=1 to 5 do TDBCheckBox(FindComponent('G_VENTILABLE'+IntToStr(i)+'')).Enabled:=(Not TDBCheckBox(FindComponent('G_VENTILABLE'+IntToStr(i)+'')).Checked) ;
   if ((G_NATUREGENE.Value='COC') or (G_NATUREGENE.Value='COD') or (G_NATUREGENE.Value='COF') or (G_NATUREGENE.Value='COS')) then
    if AvecMvt then
       BEGIN
       for i:=1 to 5 do BEGIN C:=FindComponent('G_VENTILABLE'+IntToStr(i)+'') ; TDBCheckBox(C).Enabled:=False ; END ;
       END ;
   END ;
END ;

Procedure TFCpteGene.SwapNatureGene ;
Var i : Byte ;
    OkBqe : Boolean ;
BEGIN
if(G_NATUREGENE.Value='')And(QGene.State=dsInsert) then
  BEGIN
  G_LETTRABLE.Enabled:=True ; G_POINTABLE.Enabled:=True ; G_VENTILABLE.Enabled:=True ;
  for i:=1 to 5 do TDBCheckBox(FindComponent('G_VENTILABLE'+IntToStr(i)+'')).Enabled:=True ;
  Exit ;
  END ;
VerifiNatureValide ;
GBTVA.Enabled:=(G_NATUREGENE.Value='IMO')Or(G_NATUREGENE.Value='CHA')Or(G_NATUREGENE.Value='PRO') ;
if Not GBTVA.Enabled then
  BEGIN
  if G_TVA.Value<>'' then G_TVA.Value:='' ;
  if G_TPF.Value<>'' then G_TPF.Value:='' ;
  if G_TVASURENCAISS.Checked then G_TVASURENCAISS.Checked:=False ;
  END ;
EnabledTVA ;
PAdresse.TabVisible:=OkRib(G_NATUREGENE.Value,OkBqe) ; If OkBqe Then PAdresse.TabVisible:=FALSE ;
PInfo.TabVisible:=(G_NATUREGENE.Value='TID')or(G_NATUREGENE.Value='TIC') ;
if (G_NATUREGENE.Value='TID')or(G_NATUREGENE.Value='TIC') then
  if (QGene.State in [dsInsert,dsEdit]) then QGeneG_LETTRABLE.AsString:='X' ;
SwapVentilation ; VentilationAutorise ; LettrageAutorise ; PointageAutorise ;
G_LETTRABLEClick(Nil) ; BRIB.Enabled:=OkRib(G_NATUREGENE.Value,OkBqe) ;
If (Not OkBqe) And (Not BRIB.Enabled) And (QGene.State<>dsBrowse) Then DelRibContact ;
END ;

procedure TFCpteGene.G_NATUREGENEChange(Sender: TObject);
begin
  SwapNatureGene ;
  ActivationControl;
end;

procedure TFCpteGene.BVentilClick(Sender: TObject);
var  Axes : String ;
begin
if QGene.State=dsInsert then if Not CodeValide then Exit ;
Axes:='' ;
if G_Ventilable1.Checked=True then Axes:=Axes+'1';
if G_Ventilable2.Checked=True then Axes:=Axes+'2' ;
if G_Ventilable3.Checked=True then Axes:=Axes+'3' ;
if G_Ventilable4.Checked=True then Axes:=Axes+'4' ;
if G_Ventilable5.Checked=True then Axes:=Axes+'5' ;
ParamVentil('GE',QGeneG_GENERAL.AsString,Axes,Mode,False) ;
if QGene.State=dsInsert then G_GENERAL.Enabled:=Not ExisteSql('Select V_COMPTE FROM VENTIL WHERE V_COMPTE="'+G_GENERAL.Text+'" And V_NATURE Like"GE%"') ;
end;

Function TFCpteGene.VentilationModifier : Boolean ;
BEGIN
Result:=True ;
if ((MA1='X') And (Not(G_VENTILABLE1.Checked))) Or
   ((MA1='-') And (G_VENTILABLE1.Checked)) then Exit ;
if ((MA2='X') And (Not(G_VENTILABLE2.Checked))) Or
   ((MA2='-') And (G_VENTILABLE2.Checked)) then Exit ;
if ((MA3='X') And (Not(G_VENTILABLE3.Checked))) Or
   ((MA3='-') And (G_VENTILABLE3.Checked)) then Exit ;
if ((MA4='X') And (Not(G_VENTILABLE4.Checked))) Or
   ((MA4='-') And (G_VENTILABLE4.Checked)) then Exit ;
if ((MA5='X') And (Not(G_VENTILABLE5.Checked))) Or
   ((MA5='-') And (G_VENTILABLE5.Checked)) then Exit ;
Result:=False ;
END ;

Function TFCpteGene.VentilChanger : Boolean ;
Var AvecMvtTemp : Boolean ;
    AeteModifier : Boolean ;
BEGIN
Result:=False ;
AeteModifier:=VentilationModifier ;
if (AeteModifier) And (Not AvecMvt) then
   BEGIN
   AvecMvtTemp:=IsMouvemente ;
   if (((MA1='X') And (Not G_VENTILABLE1.Checked)) Or
       ((MA2='X') And (Not G_VENTILABLE2.Checked)) Or
       ((MA3='X') And (Not G_VENTILABLE3.Checked)) Or
       ((MA4='X') And (Not G_VENTILABLE4.Checked)) Or
       ((MA5='X') And (Not G_VENTILABLE5.Checked))) And AvecMvtTemp then
      BEGIN
      MsgBox.Execute(40,'','') ; AvecMvt:=True ; RepositionneVentilation ; Exit ;
      END ;
   if AvecMvtTemp then AvecMvt:=True ;
   END ;
if Not AvecMvt then Exit ;
if QGene.State in [dsBrowse,dsInsert] then Exit ;
if (MA1='-') And (G_VENTILABLE1.Checked) then BEGIN Result:=True ; Exit ; END ;
if (MA2='-') And (G_VENTILABLE2.Checked) then BEGIN Result:=True ; Exit ; END ;
if (MA3='-') And (G_VENTILABLE3.Checked) then BEGIN Result:=True ; Exit ; END ;
if (MA4='-') And (G_VENTILABLE4.Checked) then BEGIN Result:=True ; Exit ; END ;
if (MA5='-') And (G_VENTILABLE5.Checked) then BEGIN Result:=True ; Exit ; END ;
END ;

Procedure TFCpteGene.ActiveBVentil(Sender : TObject) ;
Var i : Byte ;
begin
if ((Sender<>Nil) and (Sender is TDBCheckBox) and (TDBCheckBox(Sender).Checked)) then
   if TDBCheckBox(Sender).Name<>'G_VENTILABLE' then
      BEGIN
      i:=StrToInt(TDBCheckBox(Sender).Name[Length(TDBCheckBox(Sender).Name)]) ;
      if Not SectAtt[i] then
         BEGIN
         MsgBox.Execute(37,MsgBox.Mess[38]+' '+IntToStr(i)+' ','') ; QGene.FindField(TDBCheckBox(Sender).Name).AsString:='-' ; Exit ;
         END ;
      END ;
BVentil.Enabled:=((G_VENTILABLE.Checked)Or(G_VENTILABLE1.Checked)Or
                 (G_VENTILABLE2.Checked)Or(G_VENTILABLE3.Checked)Or
                 (G_VENTILABLE4.Checked)Or(G_VENTILABLE5.Checked)) ;
BBalGenAna.Enabled:=BVentil.Enabled ;
end;

procedure TFCpteGene.G_GENERALExit(Sender: TObject);
Var QQ : TQuery ;
    St,NewNat : String ;
begin
if(QGene.State=dsInsert)And(QGeneG_GENERAL.AsString<>'') then
   BEGIN
   QGeneG_GENERAL.AsString:=Trim(QGeneG_GENERAL.AsString) ;
   if Length(QGeneG_GENERAL.AsString)>LgCode then
      QGeneG_GENERAL.AsString:=Copy(QGeneG_GENERAL.AsString,1,LgCode) ;
   if Length(QGeneG_GENERAL.AsString)<LgCode then
      QGeneG_GENERAL.AsString:=BourreLaDonc(QGeneG_General.AsString,fbGene) ;
   if QGeneG_NATUREGENE.AsString='' then
      BEGIN
      St:=Copy(QGeneG_GENERAL.AsString,1,3) ; NewNat:='' ;
      QQ:=OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL LIKE "'+St+'%"',True) ;
      if Not QQ.EOF then NewNat:=QQ.Fields[0].AsString ;
      Ferme(QQ) ;
      if NewNat<>'' then
      begin
           QGeneG_NATUREGENE.AsString:=NewNat ;
           { AJOUT ME LE 18/04/2001 }
           if (QGeneG_NATUREGENE.AsString ='CHA') then G_SENS.value := 'D';
           if (QGeneG_NATUREGENE.AsString ='PRO') then G_SENS.value := 'C';
           // Fiche 10289 :
           SwapNatureGene;
      end;
      END ;
   END ;
  ChercheEtImportCptePlanRef ;
end;

Procedure TFCpteGene.DeleteVentilation ;
BEGIN
ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="'+MemoCpte+'" And V_NATURE Like"GE%"') ;
END ;

Procedure TFCpteGene.SwapVentilation ;
Var C : TComponent ;
    i : Byte ;
    Avec : Boolean ;
BEGIN
Avec:=TRUE ;
for i:=1 to 5 do BEGIN C:=FindComponent('G_VENTILABLE'+IntToStr(i)+'') ; TDBCheckBox(C).Enabled:=True ; END ;
{ GP : Tous les comptes peuvent être ventilables
if(G_NATUREGENE.Value='CHA')Or(G_NATUREGENE.Value='DIV')Or(G_NATUREGENE.Value='EXT')Or
  (G_NATUREGENE.Value='IMO')Or(G_NATUREGENE.Value='PRO')Or(((G_NATUREGENE.Value='BQE')Or
  (G_NATUREGENE.Value='CAI'))And Ventipoint) then Avec:=True else Avec:=False ;
}
if ((G_NATUREGENE.Value='COC') or (G_NATUREGENE.Value='COD') or (G_NATUREGENE.Value='COF') or (G_NATUREGENE.Value='COS')) then
 if ((AvecMvt) and (MA0='-')) then
    BEGIN
    for i:=1 to 5 do BEGIN C:=FindComponent('G_VENTILABLE'+IntToStr(i)+'') ; TDBCheckBox(C).Enabled:=False ; END ;
    END else
    BEGIN
    for i:=1 to 5 do
        BEGIN
        C:=FindComponent('G_VENTILABLE'+IntToStr(i)+'') ; TDBCheckBox(C).Enabled:=Avec ;
        if(Not Avec)And(QGene.State in [dsInsert,dsEdit]) then TDBCheckBox(C).Field.AsString:='-' ;
        END ;
    END ;
END ;

Procedure TFCpteGene.MajCollectif ;
BEGIN
if QGene.State in [dsBrowse]then Exit ;
if VH^.CPIFDEFCEGID then
Begin
if(G_NATUREGENE.Value='COC')Or(G_NATUREGENE.Value='COD')Or
  (G_NATUREGENE.Value='COF')Or(G_NATUREGENE.Value='COS')then QGeneG_COLLECTIF.AsString:='X' ;
Exit ;
end ;
if(G_NATUREGENE.Value='COC')Or(G_NATUREGENE.Value='COD')Or
  (G_NATUREGENE.Value='COF')Or(G_NATUREGENE.Value='COS')then
   QGeneG_COLLECTIF.AsString:='X' else QGeneG_COLLECTIF.AsString:='-' ;
END ;

Function TFCpteGene.CollectifChanger : Boolean ;
BEGIN
Result:=False ;
if QGene.State in [dsBrowse,dsInsert] then Exit ;
if(G_NATUREGENE.Value='COC')Or(G_NATUREGENE.Value='COD')Or
  (G_NATUREGENE.Value='COF')Or(G_NATUREGENE.Value='COS')then
  if G_NATUREGENE.Value<>MNat then Result:=True ;
END ;

Function TFCpteGene.CollectifDansTableTiers : Boolean ;
BEGIN Result:=Presence('TIERS','T_COLLECTIF',G_GENERAL.Text) ; END ;

Function TFCpteGene.BqeCaiChanger : Boolean ;
BEGIN
Result:=False ;
if QGene.State in [dsBrowse,dsInsert] then Exit ;
if(MNat='BQE')Or(MNat='CAI') then if(G_NATUREGENE.Value='DIV')then Result:=True ;
END ;

Function TFCpteGene.ContrePartieDansJournal : Boolean ;
BEGIN Result:=Presence('JOURNAL','J_CONTREPARTIE',G_GENERAL.Text) ; END ;

procedure TFCpteGene.BRIBClick(Sender: TObject);
Var okok : boolean ;
    LeMode : TActionFiche ;
    Existe : Boolean ;
begin
{$IFNDEF PGIIMMO}
if QGene.State=dsInsert then if Not CodeValide then Exit ;
LeMode:=Mode ;
If G_NATUREGENE.Value='BQE' Then
  BEGIN
  okok:=ExisteSQL('Select BQ_GENERAL from BANQUECP where BQ_GENERAL="'+G_GENERAL.text+'"') ;
  if((LeMode=taConsult)And(Not okok))then BEGIN Msgbox.Execute(5,'','') ; Exit ; END ;
  if((LeMode<>taConsult)And(Not okok)) then LeMode:=taCreatOne ;
  (*  if Msgbox.execute(6,'','')<>mrYes then BEGIN CreerRibDefaut ; Exit ; END
                                      else LeMode:=taCreatOne ;*) //Lhe le 06/05/1997
  if OkOk And (LeMode in [taCreat..taCreatOne]) then Exit ; //LeMode:=taModif ;
  FicheBanqueCP(G_GENERAL.text,LeMode,0) ;
  Existe:=ExisteSQL('Select BQ_GENERAL from BANQUECP where BQ_GENERAL="'+G_GENERAL.text+'"') ;
  G_GENERAL.Enabled:=Not Existe ;
  if Not Existe then CreerRibDefaut ;
  END Else
  BEGIN
  OkOk:=ExisteSQL('Select R_AUXILIAIRE from RIB where R_AUXILIAIRE="'+G_GENERAL.Text+'"') ;
  if((LeMode=taConsult)And(Not OkOk))then BEGIN Msgbox.execute(45,'','') ; Exit ; END ;
  if((LeMode<>taConsult)And(Not OkOk)) then if Msgbox.execute(46,'','')<>mrYes then Exit else LeMode:=taCreatOne ;
  if OkOk And (LeMode in [taCreat..taCreatOne]) then LeMode:=taModif ;
  FicheRIB_AGL(G_GENERAL.Text,LeMode,False,'',FALSE) ;
  G_GENERAL.Enabled:=Not(ExisteSQL('Select R_AUXILIAIRE from RIB where R_AUXILIAIRE="'+G_GENERAL.Text+'"')) ;
  END ;
{$ENDIF}
end;

procedure TFCpteGene.BCumulClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
CumulCpteMensuel(fbGene,QGeneG_GENERAL.AsString,QGeneG_LIBELLE.AsString,VH^.Entree) ;
{$ENDIF}
{$ENDIF}
end;

Function TFCpteGene.Supprime : Boolean ;
BEGIN
Result:=False ;
if MsgBox.Execute(1,'','')<>mrYes then Exit ;
Result:=True ;
END ;

Function EstEcrGuide(LeCpte : String)     : Boolean ;
BEGIN Result:=Presence('ECRGUI','EG_GENERAL',LeCpte) ; END ;

Function EstAnalytiqPure(LeCpte : String) : Boolean ;
BEGIN Result:=Presence('ANALYTIQ','Y_GENERAL',LeCpte) ; END ;

Function EstCpteAxe(LeCpte : String) : Boolean ;
Var i : Byte ;
BEGIN
Result:=False ;
for i:=1 to 5 do
   BEGIN
   if VH^.Cpta[AxeToFb('A'+IntToStr(i))].AxGenAttente=LeCpte then BEGIN Result:=True ; Exit ; END ;
   END ;
END ;

Function EstBanquaire(LeCpte : String) : Boolean ;
BEGIN Result:=Presence('BANQUECP','BQ_GENERAL',LeCpte) ; END ;

Function EstCpteCorresp(LeCpte : String) : Boolean ;
BEGIN Result:=Presence('CORRESP','CR_CORRESP',LeCpte) ; END ;

Function EstCpteDevise(LeCpte : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select D_DEVISE from DEVISE Where(D_CPTLETTRDEBIT="'+LeCpte+'") OR (D_CPTLETTRCREDIT="'+LeCpte+'") '+
           'OR (D_CPTPROVDEBIT="'+LeCpte+'") OR (D_CPTPROVCREDIT="'+LeCpte+'")',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function EstCpteJournal(LeCpte : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select J_JOURNAL from JOURNAL Where J_CONTREPARTIE="'+LeCpte+'" '+
           'AND ((J_NATUREJAL="BQE") OR (J_NATUREJAL="CAI") OR (J_NATUREJAL="REG"))',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function EstCpteModepaie(LeCpte : String) : Boolean ;
BEGIN Result:=Presence('MODEPAIE','MP_GENERAL',LeCpte) ; END ;

Function EstCpteSociete(LeCpte : String) : Boolean ;
Var Q : TQuery ;
    St : String ;
BEGIN
{$IFDEF SPEC302}
Q:=OpenSql('Select SO_SOCIETE from SOCIETE Where SO_GENATTEND="'+LeCpte+'" OR SO_FERMEBIL="'+LeCpte+'" '+
           ' OR SO_OUVREBIL="'+LeCpte+'" OR SO_RESULTAT="'+LeCpte+'" OR SO_FERMEPERTE="'+LeCpte+'" '+
           ' OR SO_OUVREPERTE="'+LeCpte+'" OR SO_FERMEBEN="'+LeCpte+'" OR SO_OUVREBEN="'+LeCpte+'"',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
{$ELSE}
Result:=True ;
St:=GetParamsoc('SO_GENATTEND') ; if St=LeCpte then Exit ;
St:=GetParamsoc('SO_FERMEBIL') ; if St=LeCpte then Exit ;
St:=GetParamsoc('SO_OUVREBIL') ; if St=LeCpte then Exit ;
St:=GetParamsoc('SO_RESULTAT') ; if St=LeCpte then Exit ;
St:=GetParamsoc('SO_FERMEPERTE') ; if St=LeCpte then Exit ;
St:=GetParamsoc('SO_OUVREPERTE') ; if St=LeCpte then Exit ;
St:=GetParamsoc('SO_FERMEBEN') ; if St=LeCpte then Exit ;
St:=GetParamsoc('SO_OUVREBEN') ; if St=LeCpte then Exit ;
Result:=False ;
{$ENDIF}
END ;

Function EstCpteTva(LeCpte : String) : Boolean ;
BEGIN
Result:=EstTvaTpf(LeCpte,True) ;
if Not Result then Result:=EstTvaTpf(LeCpte,False) ;
END ;

Function  EstCpteAmortissement(LeCpte : String) : boolean;
var Q : TQuery;
begin
  Result := ExisteSQL ( 'SELECT I_IMMO FROM IMMO WHERE '+
      'I_COMPTEIMMO="'+LeCpte+'" OR I_COMPTEAMORT="'+LeCpte+'" OR I_COMPTEDOTATION="'+LeCpte+
      '" OR I_COMPTEDEROG="'+LeCpte+'" OR I_REPRISEDEROG="'+LeCpte+'" OR I_PROVISDEROG="'+LeCpte+
      '" OR I_DOTATIONEXC="'+LeCpte+'" OR I_VACEDEE="'+LeCpte+'" OR I_AMORTCEDE="'+LeCpte+
      '" OR I_REPEXPLOIT="'+LeCpte+'" OR I_REPEXCEP="'+LeCpte+'" OR I_VAOACEDEE="'+LeCpte+'"');
end;

procedure TFCpteGene.BZecrimvtClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
ZoomEcritureMvt(G_GENERAL.Text, fbGene,'MULMMVTS' ) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteGene.BJustifClick(Sender: TObject);
begin (* rony 14/04/97 JustifSolde(G_GENERAL.Text,fbGene);*) end;

Function TFCpteGene.CaracteristiqueIncompatible(Var j : Byte) : Boolean ;
Var i : Byte ;
BEGIN
Result:=False ;
for i:=1 to 5 do if TDBCheckBox(FindComponent('G_VENTILABLE'+IntToStr(i)+'')).Checked then
    BEGIN G_VENTILABLE.Checked:=True ; Break ; END else G_VENTILABLE.Checked:=False ;
if (G_VENTILABLE.Checked)And(G_LETTRABLE.Checked)And(G_POINTABLE.Checked) then BEGIN Result:=true ; j:=1 ; Exit ; END ;
{ GP : Tous les comptes peuvent être ventilables
if (G_VENTILABLE.Checked)And(G_LETTRABLE.Checked) then BEGIN Result:=true ; j:=2 ; Exit ; END ;
}
if (G_LETTRABLE.Checked)And(G_POINTABLE.Checked) then BEGIN Result:=true ; j:=3 ; Exit ; END ;
if Not VentiPoint then
   BEGIN
   if(G_VENTILABLE.Checked)And(G_POINTABLE.Checked) then BEGIN Result:=True ; j:=4 ; Exit ; END ;
   END ;
END ;

Function TFCpteGene.EstVentilable(Quelaxe : Byte) : Boolean ;
BEGIN
Result:=False ;
Case QuelAxe of
     1: if(MA1='-')And(G_VENTILABLE1.Checked=True) then Result:=True ;
     2: if(MA2='-')And(G_VENTILABLE2.Checked=True) then Result:=True ;
     3: if(MA3='-')And(G_VENTILABLE3.Checked=True) then Result:=True ;
     4: if(MA4='-')And(G_VENTILABLE4.Checked=True) then Result:=True ;
     5: if(MA5='-')And(G_VENTILABLE5.Checked=True) then Result:=True ;
    End ;
END ;

Procedure TFCpteGene.MajVentilation(Lecpte : String) ;
Var Dev : RDevise ;
    Q,QAna : TQuery ;
    OEcr,OAna : TOBM ;
    TVentAna : TList ;
    NewAxe : String ;
    i ,k : Integer ;
    Okok : Boolean ;
BEGIN
Q:=OpenSQL('Select * from ECRITURE Where E_GENERAL="'+LeCpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")',True) ;
QAna:=OpenSQL('Select * from ANALYTIQ WHERE Y_SECTION="RerER"',False) ;
OEcr:=TOBM.Create(EcrGen,'',False) ; TVentAna:=TList.Create ; Okok:=False ;
While not Q.EOF do
   BEGIN
   OEcr.ChargeMvt(Q) ;
   Dev.Code:=Q.FindField('E_DEVISE').AsString ; GetInfosDevise(Dev) ;
   Dev.DateTaux:=Q.FindField('E_DATETAUXDEV').AsDateTime ;
   Dev.Taux:=Q.FindField('E_TAUXDEV').AsFloat ;
   for i:=1 to MaxAxe do if EstVentilable(i) then
       BEGIN
       VideListe(TVentAna) ; NewAxe:='A'+IntToStr(i) ; Okok:=True ;
       VentileGenerale(NewAxe,OEcr,Dev,TVentAna,True) ;
       for k:=0 to TVentAna.Count-1 do
           BEGIN
           QAna.Insert ;
           OAna:=TOBM(TVentAna[k]) ; OAna.EgalChamps(QAna) ;
           QAna.Post ;
           MajSoldeSection(QAna,True) ;
           END ;
       END ;
    Q.Next ;
   END ;
Ferme(Q) ; Ferme(QAna) ; VideListe(TVentAna) ; TVentAna.Free ; OEcr.Free ;
if Okok then ExecuteSQL('UPDATE ECRITURE SET E_ANA="X" WHERE E_GENERAL="'+LeCpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")') ; 
END ;

Procedure MajLettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
Var LeWhere : String ;
BEGIN
Case Lefb of
     fbGene :LeWhere:='WHERE E_GENERAL="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
     fbAux  :LeWhere:='WHERE E_AUXILIAIRE="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
    End ;
BeginTrans ;
(*
ExecuteSql('UPDATE ECRITURE SET E_MODEPAIE="'+Mp+'", E_DATEECHEANCE=E_DATECOMPTABLE, '+
           'E_ETATLETTRAGE="AL", E_NUMECHE=1, E_ECHE="X" '+LeWhere) ;
*)
// GG COM
If OkSynChro Then ExecuteSql('UPDATE ECRITURE SET E_MODEPAIE="'+Mp+'", E_DATEECHEANCE=E_DATECOMPTABLE, '+
                             'E_ETATLETTRAGE="AL", E_NUMECHE=1, E_ECHE="X", E_PAQUETREVISION=1 '+LeWhere)
             Else ExecuteSql('UPDATE ECRITURE SET E_MODEPAIE="'+Mp+'", E_DATEECHEANCE=E_DATECOMPTABLE, '+
                             'E_ETATLETTRAGE="AL", E_NUMECHE=1, E_ECHE="X" '+LeWhere) ;
CommitTrans ;
END ;

Procedure MajDelettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
Var LeWhere : String ;
BEGIN
Case Lefb of
     fbGene :LeWhere:='WHERE E_GENERAL="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
     fbAux  :LeWhere:='WHERE E_AUXILIAIRE="'+Cpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' ;
    End ;
BeginTrans ;
ExecuteSql('UPDATE ECRITURE SET E_MODEPAIE="'+Mp+'", E_DATEECHEANCE=E_DATECOMPTABLE, '+
           'E_ETATLETTRAGE="RI", E_NUMECHE=0, E_ECHE="-" '+LeWhere) ;
CommitTrans ;
END ;

Procedure TFCpteGene.InitModifEnSerie(StModif : string) ;
var St,Champ, Valeur : string;
    i             : integer;
    B             : TBitBtn;
BEGIN
if QGene.State=dsBrowse then QGene.Edit ;
While StModif<>'' do
   BEGIN
   St:=ReadTokenSt(StModif);
   i:=Pos('=',St); if i>0 then Champ:=Trim(Copy(St,1,i-1));
   i:=Pos('"',St); if i>0 then St:=Trim(Copy(St,i+1,Length(St)));
   i:=Pos('"',St); if i>0 then Valeur:=Trim(Copy(St,1,i-1));
   AffecteLe(Champ,Valeur);
   END;
For i:=0 to HPB.ControlCount-1 do
   if HPB.Controls[i] is TBitBtn then
      BEGIN
      B:=TBitBtn(HPB.Controls[i]);
      if ((B.Name<>'BValider') and (B.Name<>'BFerme') and (B.Name<>'BAide')) then B.Enabled:=false;
      END;
END;

Procedure TFCpteGene.AffecteLe(Champ, Valeur : string) ;
var C : TControl;
BEGIN
C:=TControl(FindComponent(Champ)) ;
if(C is TDBCheckBox)Or(C is THDBValComboBox)Or(C is TDBEdit)Or(C is THDBCpteEdit)then
   BEGIN
   QGene.FindField(Champ).AsString:=Valeur ; TEdit(C).Font.Color:=clRed;
   END else if C is THDBSpinEdit then
   BEGIN
   QGene.FindField(Champ).AsInteger:=StrToInt(Valeur) ; THDBSpinEdit(C).Font.Color:=clRed;
   END ;
END;

Function TFCpteGene.OkConfidentiel : Boolean ;
BEGIN
Result:=False ;
if (V_PGI.Confidentiel='0') and (G_Confidentiel.Checked) then
  BEGIN
  MessageAlerte(MsgBox.Mess[33]) ;
  PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
  END ;
GBConfidentiel.Visible:=(V_PGI.Confidentiel='1') ;
GBConfidentiel.Enabled:=V_PGI.Superviseur ;
Result:=True ;
END ;

procedure TFCpteGene.BMenuZoomClick(Sender: TObject);
BEGIN
PopZoom(BMenuZoom,POPZ) ;
end;

procedure TFCpteGene.FormCreate(Sender: TObject);
begin
{$IFNDEF IMP}
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
{$ENDIF}
Q:=NIL ;
LeCompte:='azeaze' ;
end;

Function TFCpteGene.TvaTicTidOk : Boolean ;
BEGIN
Result:=True ;
if((G_NATUREGENE.Value<>'TIC')And(G_NATUREGENE.Value<>'TID'))Or
  (((G_NATUREGENE.Value='TIC')Or(G_NATUREGENE.Value='TID'))And(Not G_LETTRABLE.Checked)) then
  BEGIN
  G_REGIMETVA.Value:='' ; G_TVAENCAISSEMENT.Value:='' ; G_MODEREGLE.Value:='' ; Exit ;
  END ;
if G_REGIMETVA.Text='' then
   BEGIN
   Pages.ActivePage:=PInfo ; G_REGIMETVA.SetFocus ; MsgBox.Execute(34,'','') ; Result:=False ;
   END ;
END ;

procedure TFCpteGene.G_LETTRABLEClick(Sender: TObject);
begin
if(G_NATUREGENE.Value<>'TIC')And(G_NATUREGENE.Value<>'TID') then Exit ;
if Not G_LETTRABLE.Checked then Exit ;
if G_TVAENCAISSEMENT.Text='' then G_TVAENCAISSEMENT.Value:=G_TVAENCAISSEMENT.Values[0] ;
if G_MODEREGLE.Text='' then G_MODEREGLE.Value:=G_MODEREGLE.Values[0] ;
end;

procedure TFCpteGene.G_CODEPOSTALDblClick(Sender: TObject);
begin VerifCodePostal(QGene,G_CODEPOSTAL,G_VILLE,TRUE) ; end;

procedure TFCpteGene.G_VILLEDblClick(Sender: TObject);
begin VerifCodePostal(QGene,G_CODEPOSTAL,G_VILLE,FALSE) ; end;

procedure TFCpteGene.BBalGenAnaClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit : TCritEdt ;
    D1,D2 : TdateTime ;
    Axe : String3 ;
    Etab : String ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF CCMP}
{$IFNDEF IMP}
Axe:='A1' ;
If G_VENTILABLE1.Checked Then Axe:='A1' Else If G_VENTILABLE2.Checked Then Axe:='A2' Else
If G_VENTILABLE3.Checked Then Axe:='A3' Else If G_VENTILABLE4.Checked Then Axe:='A4' Else
If G_VENTILABLE5.Checked Then Axe:='A5' Else Exit ;
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
If VH^.Entree.Code=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neBal ;
Crit.Bal.Axe:=Axe ;
InitCritEdt(Crit) ;
Crit.Bal.TypCpt:=1 ;
Crit.Cpt1:=G_General.text ;
Crit.Cpt2:=Crit.Cpt1 ;
Etab:=EtabForce ; if Etab<>'' then Crit.Etab:=Etab ;
BLGESEZoom(Crit) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end ;

Procedure TFCpteGene.CreerRibDefaut ;
Var QLoc : TQuery ;
    St : String ;
BEGIN
MsgBox.Execute(36,'','') ;
QLoc:=OpenSql('Select PQ_BANQUE From BANQUES',True) ;
St:='' ; if Not QLoc.Eof then St:=QLoc.Fields[0].AsString ; Ferme(QLoc) ;
ExecuteSql('INSERT INTO BANQUECP (BQ_CODE, BQ_GENERAL, BQ_BANQUE, BQ_LIBELLE, BQ_DEVISE, BQ_PAYS) '+
           'VALUES ("'+G_GENERAL.text+'", "'+G_GENERAL.text+'", "'+St+'", "'+MsgBox.Mess[35]+'", '+
           '"'+V_PGI.DevisePivot+'","FRA")') ;
G_GENERAL.Enabled:=False ;
END ;

procedure TFCpteGene.BBALGENAUXClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit : TCritEdt ;
    D1,D2 : TdateTime ;
    Etab : String ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
If VH^.Entree.Code=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neBal ;
InitCritEdt(Crit) ;
Crit.Bal.TypCpt:=1 ;
Crit.Cpt1:=G_General.text ;
Crit.Cpt2:=Crit.Cpt1 ;
Etab:=EtabForce ; if Etab<>'' then Crit.Etab:=Etab ;
BalGeAuZoom(Crit) ;
{$ENDIF}
{$ENDIF}
end ;

procedure TFCpteGene.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ;
if Vide then
   BEGIN
   Case Key of
        VK_F3  : BPrevClick(Nil) ;
        VK_F4  : BNextClick(Nil) ;
        VK_F10 : BEGIN Key:=0 ; BValiderClick(Nil) ; END ;
        VK_RETURN :BEGIN
                   if ActiveControl is TCustomMemo then Exit ;
                   FindNextControl(ActiveControl,True,True,False).SetFocus ;
                   END ;
     END ;
   END ;
end;

procedure TFCpteGene.G_GENERALEnter(Sender: TObject);
begin
  if G_GENERAL.Text='' then Import:=True else Import:=False ;
  ActivationControl;
end;

Procedure TFCpteGene.ChercheEtImportCptePlanRef ;
Var St : String ;
BEGIN
if VH^.NumPlanRef=0 then Exit ;
if Not Import then Exit ;
St:=Bourreless(QGeneG_GENERAL.asString,fbGene) ;
if St='' then Exit ;
if Presence('GENERAUX','G_GENERAL',QGeneG_GENERAL.AsString) then
   BEGIN
   Pages.ActivePage:=PComptable ; G_GENERAL.SetFocus ;
   MsgBox.execute(4,'','') ; Exit ;
   END ;
QPRef.Close ;
QPRef.ParamByName('N').AsInteger:=VH^.NumPlanRef ;
QPRef.ParamByName('C').AsString:=St ;
QPRef.Open ;
if QPRef.EOF then BEGIN QPRef.Close ; Exit ; END ;
{$IFNDEF PGIIMMO}
if FicheImportPlanRef(VH^.NumPlanRef,St)=1 then
   BEGIN
   QGeneG_LIBELLE.AsString:=QPRef.FindField('PR_LIBELLE').AsString ;
   QGeneG_ABREGE.AsString:=QPRef.FindField('PR_ABREGE').AsString ;
   QGeneG_CENTRALISABLE.AsString:=QPRef.FindField('PR_CENTRALISABLE').AsString ;
   QGeneG_SOLDEPROGRESSIF.AsString:=QPRef.FindField('PR_SOLDEPROGRESSIF').AsString ;
   QGeneG_SAUTPAGE.AsString:=QPRef.FindField('PR_SAUTPAGE').AsString ;
   QGeneG_TOTAUXMENSUELS.AsString:=QPRef.FindField('PR_TOTAUXMENSUELS').AsString ;
   QGeneG_COLLECTIF.AsString:=QPRef.FindField('PR_COLLECTIF').AsString ;
   QGeneG_BLOCNOTE.AsString:=QPRef.FindField('PR_BLOCNOTE').AsString ;
   QGeneG_SENS.AsString:=QPRef.FindField('PR_SENS').AsString ;
   QGeneG_LETTRABLE.AsString:=QPRef.FindField('PR_LETTRABLE').AsString ;
   QGeneG_POINTABLE.AsString:=QPRef.FindField('PR_POINTABLE').AsString ;
   QGeneG_VENTILABLE1.AsString:=QPRef.FindField('PR_VENTILABLE1').AsString ;
   QGeneG_VENTILABLE2.AsString:=QPRef.FindField('PR_VENTILABLE2').AsString ;
   QGeneG_VENTILABLE3.AsString:=QPRef.FindField('PR_VENTILABLE3').AsString ;
   QGeneG_VENTILABLE4.AsString:=QPRef.FindField('PR_VENTILABLE4').AsString ;
   QGeneG_VENTILABLE5.AsString:=QPRef.FindField('PR_VENTILABLE5').AsString ;
   G_NATUREGENE.Value:=QPRef.FindField('PR_NATUREGENE').AsString ;
   END ;
{$ENDIF}
QPRef.Close ;
END ;

Procedure TFCpteGene.MajDuPlanReference ;
Var St : String ;
BEGIN
if VH^.NumPlanRef=0 then Exit ;
if Not MajPlanRef then Exit ;
St:=BourreLess(QGeneG_GENERAL.AsString,fbGene) ;
QPRef.Close ;
QPRef.ParamByName('N').AsInteger:=VH^.NumPlanRef ;
QPRef.ParamByName('C').AsString:=St ;
BeginTrans ; QPRef.Open ;
if QPRef.EOF then
   BEGIN
   ExecuteSql('INSERT INTO PLANREF (PR_NUMPLAN, PR_COMPTE, PR_PREDEFINI, PR_LIBELLE, PR_ABREGE, PR_NATUREGENE, '+
              'PR_CENTRALISABLE, PR_SOLDEPROGRESSIF, PR_SAUTPAGE, PR_TOTAUXMENSUELS, PR_COLLECTIF, '+
              'PR_SENS, PR_LETTRABLE, PR_POINTABLE, PR_VENTILABLE1, PR_VENTILABLE2, '+
              'PR_VENTILABLE3, PR_VENTILABLE4, PR_VENTILABLE5) '+
              'VALUES '+
              '('+IntToStr(VH^.NumPlanRef)+', "'+St+'", '+'"CEG", '+
                '"'+QGeneG_LIBELLE.AsString+'", "'+QGeneG_ABREGE.AsString+'", "'+QGeneG_NATUREGENE.AsString+'", '+
                '"'+QGeneG_CENTRALISABLE.AsString+'", "'+QGeneG_SOLDEPROGRESSIF.AsString+'", '+
                '"'+QGeneG_SAUTPAGE.AsString+'", "'+QGeneG_TOTAUXMENSUELS.AsString+'", '+
                '"'+QGeneG_COLLECTIF.AsString+'", "'+QGeneG_SENS.AsString+'", '+
                '"'+QGeneG_LETTRABLE.AsString+'", "'+QGeneG_POINTABLE.AsString+'", '+
                '"'+QGeneG_VENTILABLE1.AsString+'", "'+QGeneG_VENTILABLE2.AsString+'", '+
                '"'+QGeneG_VENTILABLE3.AsString+'", "'+QGeneG_VENTILABLE4.AsString+'", "'+QGeneG_VENTILABLE5.AsString+'")') ;
   END else
   BEGIN
   ExecuteSql('UPDATE PLANREF SET PR_LIBELLE="'+QGeneG_LIBELLE.AsString+'", PR_ABREGE="'+QGeneG_ABREGE.AsString+'", '+
              'PR_NATUREGENE="'+QGeneG_NATUREGENE.AsString+'", PR_CENTRALISABLE="'+QGeneG_CENTRALISABLE.AsString+'", '+
              'PR_SOLDEPROGRESSIF="'+QGeneG_SOLDEPROGRESSIF.AsString+'", PR_SAUTPAGE="'+QGeneG_SAUTPAGE.AsString+'", '+
              'PR_TOTAUXMENSUELS="'+QGeneG_TOTAUXMENSUELS.AsString+'", PR_COLLECTIF="'+QGeneG_COLLECTIF.AsString+'", '+
              'PR_SENS="'+QGeneG_SENS.AsString+'", PR_LETTRABLE="'+QGeneG_LETTRABLE.AsString+'", '+
              'PR_POINTABLE="'+QGeneG_POINTABLE.AsString+'", '+
              'PR_VENTILABLE1="'+QGeneG_VENTILABLE1.AsString+'", PR_VENTILABLE2="'+QGeneG_VENTILABLE2.AsString+'", '+
              'PR_VENTILABLE3="'+QGeneG_VENTILABLE3.AsString+'", PR_VENTILABLE4="'+QGeneG_VENTILABLE4.AsString+'", '+
              'PR_VENTILABLE5="'+QGeneG_VENTILABLE5.AsString+'" WHERE PR_NUMPLAN='+IntToStr(VH^.NumPlanRef)+' AND '+
              'PR_COMPTE="'+St+'"') ;
   END ;
QPRef.Close ; CommitTrans ;
QPRef.ParamByName('N').AsInteger:=VH^.NumPlanRef ;
QPRef.ParamByName('C').AsString:=St ;
QPRef.Open ;
if QPRef.Eof then BEGIN QPRef.Close ; Exit ; END else
   BEGIN
   QPRef.Edit ;
   if Not TMemoField(QGene.FindField('G_BLOCNOTE')).IsNull
      then TMemoField(QPRef.FindField('PR_BLOCNOTE')).Assign(TMemoField(QGene.FindField('G_BLOCNOTE'))) ;
   QPRef.Post ;
   END ;
QPRef.Close ;         
END ;

Procedure TFCpteGene.ChercheSectAttente ;
Var i : Byte ;
BEGIN
for i:=1 to 5 do if VH^.Cpta[AxeToFb('A'+IntToStr(i))].Attente=''
    then SectAtt[i]:=False else SectAtt[i]:=True ;
END ;

procedure TFCpteGene.G_DIVTERRITDblClick(Sender: TObject);
begin
PaysRegion(G_PAYS,G_DIVTERRIT,True) ;
if G_DIVTERRIT.Text='' then QGeneG_DIVTERRIT.AsString:='' ;
if G_PAYS.Text<>'' then QGeneG_PAYS.AsString:=G_PAYS.Value ;
end;

procedure TFCpteGene.G_PAYSChange(Sender: TObject);
begin
PaysRegion(G_PAYS,G_DIVTERRIT,False) ;
if G_DIVTERRIT.Text='' then QGeneG_DIVTERRIT.AsString:='' ;
if G_PAYS.Text<>'' then QGeneG_PAYS.AsString:=G_PAYS.Value ;
end;

Procedure TFCpteGene.FormateLesMontants ;
BEGIN
QGeneG_DEBITDERNMVT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_CREDITDERNMVT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTDEBP.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTCREP.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTDEBE.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTCREE.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTDEBS.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTCRES.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTDEBANO.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTCREANO.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_RISQUE.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_PLAFOND.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTALDEBIT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_TOTALCREDIT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_DEBNONPOINTE.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QGeneG_CREDNONPOINTE.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
ChangeMask(G_SOLCREP,V_PGI.OkdecV,'') ;
ChangeMask(G_SOLCREE,V_PGI.OkdecV,'') ;
ChangeMask(G_SOLCRES,V_PGI.OkdecV,'') ;
END ;

procedure TFCpteGene.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

Function TFCpteGene.VerifCoherenceTL : Boolean ;
Var i : Integer ;
    C : TComponent ;                                            
    Alerte : Boolean ;
BEGIN
Result:=True ;
if Not ZL.TabVisible then BEGIN Result:=True ; Exit ; END ;
Alerte:=False ;
for i:=0 to 9 do
  BEGIN
  C:=FindComponent('G_TABLE'+IntToStr(i)) ;
  if Not(THDBCpteEdit(C).Enabled) then Continue ;
  if THDBCpteEdit(C).Text='' then Continue ;
  if THDBCpteEdit(C).ExisteH<=0 then Alerte:=True ;
  END ;
if Alerte then if MsgBox.Execute(44,'','')<>mrYes then Result:=False ;
END ;

procedure TFCpteGene.HMTradBeforeTraduc(Sender: TObject);
begin
LibellesTableLibre(ZL,'TG_TABLE','G_TABLE','G') ;
end;

procedure TFCpteGene.BImmoClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCMP}
ConsultationImmo(taConsult,toNone,True,G_GENERAL.Text) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

Procedure TFCpteGene.DelRibContact ;
BEGIN
if BourreLess(QGeneG_GENERAL.AsString,fbgene)='' then Exit ;
if Mode in [taCreat..taCreatOne] then
   BEGIN
   ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+QGeneG_GENERAL.AsString+'"') ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2002
Modifié le ... :   /  /
Description .. : Active ou Désactive les contrôles de la fiche lors de
Suite ........ : la création d'un compte général afin de forcer
Suite ........ : la saisie de  G_General et G_NatureGene
Mots clefs ... :
*****************************************************************}
procedure TFCpteGene.ActivationControl;
var lBoEnabled : Boolean;
    lStGeneral : String;
    lStNatureGene : String;
begin
  if not ( CtxPcl in V_Pgi.PGIContexte ) then Exit;

  lStGeneral := Trim(G_General.Text);
  lStNatureGene := Trim(G_NatureGene.Text);

  G_Sens.Enabled         := ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_Libelle.Enabled      := ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_Abrege.Enabled       := ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_Lettrable.Enabled    := ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_Pointable.Enabled    := ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_PurGeable.Enabled    := ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_Confidentiel.Enabled := ( lStGeneral <> '' ) and ( lStNatureGene <> '');

  if (lStGeneral <> '' ) and (lStNatureGene <> '') then
    SwapNatureGene;

  if G_Sens.Enabled then
    G_Sens.Color := ClWindow
  else
    G_Sens.Color := ClBtnFace;

  if G_Libelle.Enabled then
    G_Libelle.Color := ClWindow
  else
    G_Libelle.Color := ClBtnFace;

  if G_Abrege.Enabled then
    G_Abrege.Color := ClWindow
  else
    G_Abrege.Color := ClBtnFace;

  lBoEnabled := (GetParamSoc('SO_CPPCLSANSANA')=TRUE);

  G_VENTILABLE1.enabled := (not lBoEnabled ) and ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_VENTILABLE2.enabled := (not lBoEnabled ) and ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_VENTILABLE3.enabled := (not lBoEnabled ) and ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_VENTILABLE4.enabled := (not lBoEnabled ) and ( lStGeneral <> '' ) and ( lStNatureGene <> '');
  G_VENTILABLE5.enabled := (not lBoEnabled ) and ( lStGeneral <> '' ) and ( lStNatureGene <> '');

end;

procedure TFCpteGene.G_GENERALChange(Sender: TObject);
begin
  ActivationControl;
end;

procedure TFCpteGene.PagesChanging(Sender: TObject; var AllowChange: Boolean);
begin
  if not ( CtxPcl in V_Pgi.PGIContexte ) then Exit;

  AllowChange := (Trim(G_General.Text) <> '') and (Trim(G_NatureGene.Text) <> '');

  if not AllowChange then
  begin
    if Trim(G_General.Text) = '' then
      G_General.SetFocus
    else
      G_NatureGene.SetFocus;
  end;

end;

procedure TFCpteGene.G_NATUREGENEEnter(Sender: TObject);
begin
  ActivationControl;
end;

end.
