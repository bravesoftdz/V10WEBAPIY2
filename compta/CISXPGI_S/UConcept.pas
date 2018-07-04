{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 12/07/2002
Modifié le ... : 12/12/2006
Description .. : Unité de paramétrage de l'objet Script.
Suite ........ : base access
Mots clefs ... : Reprise à partir de External Import
*****************************************************************}

unit UConcept;

interface

uses
  Windows, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, db,
  StdCtrls, Buttons, Hctrls, ExtCtrls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, Variants, ADODB,  {$ENDIF}
  QRE,
{$ELSE}
  MainEAgl,
  UScriptTob,
  UscriptCwas,
{$ENDIF}
  hmsgbox, HEnt1,
  HPanel, UiUtil, ComCtrls, ASCIIV, Mask, ColMemo, Spin, HTB97, Menus,Clipbrd,
  HRichOLE, Uscript, UTOB, licUtil,
  UControlParam, HColor, HStatus, FileCtrl, UScriptDelim, Vierge,
  HRichEdt, ImgList;


type
    TExtPConcept = class;
    TNoeud = class
            FOwner : TExtPConcept;
    public
            procedure Show; virtual; abstract;
            constructor Create(AOwner:TForm);
            destructor Destroy; override;
    end;

    TChampNode = class (TNoeud)
    public
            FChamp : TChamp;
    end;

  TExtPConcept = class(TForm)
    Panel1: TPanel;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    bZoomPiece: TToolbarButton97;
    Msg: THMsgBox;
    bSelectAll: TToolbarButton97;
    bDown: TToolbarButton97;
    ImageList1: TImageList;
    PopupMenu1: TPopupMenu;
    Nouveau1: TMenuItem;
    Champnormal1: TMenuItem;
    Champ1: TMenuItem;
    Concat1: TMenuItem;
    Type1: TMenuItem;
    miNormal: TMenuItem;
    miConstante: TMenuItem;
    miCalcul: TMenuItem;
    miReference: TMenuItem;
    N6: TMenuItem;
    Supprimer1: TMenuItem;
    Copier1: TMenuItem;
    Couper1: TMenuItem;
    Coller1: TMenuItem;
    N1: TMenuItem;
    ViewFScript1: TMenuItem;
    PopupMenu2: TPopupMenu;
    Affecter1: TMenuItem;
    Scruter1: TMenuItem;
    Panel2: TPanel;
    Memo1: TASCIIView;
    Panel3: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Lbdeb: TLabel;
    lbLon: TLabel;
    lbFin: TLabel;
    btnAffecter: TToolbarButton97;
    PageControl1: TPageControl;
    tbDEFINITION: TTabSheet;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label23: TLabel;
    GroupBox1: TGroupBox;
    lblDebut: TLabel;
    Label7: TLabel;
    lblFin: TLabel;
    ChampDebut: TSpinEdit;
    ChampLon: TSpinEdit;
    edFin: TSpinEdit;
    edLongueur: TSpinEdit;
    ChampNom: TEdit;
    ComboBox1: TComboBox;
    GroupBox3: TGroupBox;
    cbTouteValeur: TRadioButton;
    cbValeurPositive: TRadioButton;
    cbValeurNegative: TRadioButton;
    cbFormatDate: TComboBox;
    cbType: TComboBox;
    tsComplement: TTabSheet;
    RadioGroup1: TRadioGroup;
    GroupBox2: TGroupBox;
    Label16: TLabel;
    cbTableExterne: TCheckBox;
    btnEditTbl: TButton;
    edNomTable: THCritMaskEdit;
    cbTableCorr: TCheckBox;
    GroupBox5: TGroupBox;
    Label8: TLabel;
    Label18: TLabel;
    cbComplement: TCheckBox;
    edComplLgn: TSpinEdit;
    edComplCar: TEdit;
    cbAlignLeft: TCheckBox;
    tsCALCUL: TTabSheet;
    Label17: TLabel;
    SpeedButton23: TSpeedButton;
    SpeedButton24: TSpeedButton;
    SpeedButton25: TSpeedButton;
    SpeedButton26: TSpeedButton;
    SpeedButton27: TSpeedButton;
    SpeedButton28: TSpeedButton;
    Label9: TLabel;
    SpeedButton29: TSpeedButton;
    SpeedButton31: TSpeedButton;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    Label26: TLabel;
    lbChampNum: TListBox;
    Button2: TButton;
    CalFormule: TColorMemo;
    CheckCalcul: TToolbarButton97;
    tsREFERENCE: TTabSheet;
    Label10: TLabel;
    lbRang: TLabel;
    RefPosterieur: TCheckBox;
    RefCondition: TEdit;
    Memo5: TMemo;
    GroupBox4: TGroupBox;
    Label11: TLabel;
    Label12: TLabel;
    RefPosition: TEdit;
    RefLongueur: TEdit;
    EdRang: TSpinEdit;
    Checkcondition: TToolbarButton97;
    tsCONSTANTE: TTabSheet;
    Valeur: TLabel;
    Pas: TLabel;
    Longueur: TLabel;
    CstVal: TEdit;
    CstPas: TSpinEdit;
    CstLon: TSpinEdit;
    Memo4: TMemo;
    TabConditioncomp: TTabSheet;
    tsCoherence: TTabSheet;
    MemoMsg: TMemo;
    Tabparam: TTabSheet;
    LFichier: TLabel;
    FILENAME: THCritMaskEdit;
    GroupBox6: TGroupBox;
    Label25: TLabel;
    Label19: TLabel;
    Label22: TLabel;
    Fichesauve: THCritMaskEdit;
    edcomment: TEdit;
    tsCHPCONDITION: TTabSheet;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton9: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton12: TSpeedButton;
    SpeedButton13: TSpeedButton;
    SpeedButton30: TSpeedButton;
    memCondition: THRichEditOLE;
    TabSortie: TTabSheet;
    Radiotable: TRadioGroup;
    RTypefichier: TRadioGroup;
    Label39: TLabel;
    EdNature: THCritMaskEdit;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    bCopier: TToolbarButton97;
    Panel4: TPanel;
    tvChamp: TTreeView;
    Splitter1: TSplitter;
    BColler: TToolbarButton97;
    Bcouper: TToolbarButton97;
    Bapprecu: TToolbarButton97;
    BColor: TPaletteButton97;
    Memo3: THRichEditOLE;
    SpeedButton16: TSpeedButton;
    SpeedButton17: TSpeedButton;
    SpeedButton18: TSpeedButton;
    SpeedButton19: TSpeedButton;
    SpeedButton20: TSpeedButton;
    SpeedButton21: TSpeedButton;
    SpeedButton22: TSpeedButton;
    SpeedButton32: TSpeedButton;
    SpeedButton33: TSpeedButton;
    SpeedButton34: TSpeedButton;
    SpeedButton35: TSpeedButton;
    SpeedButton36: TSpeedButton;
    bCondition: TCheckBox;
    EnregUnique: TCheckBox;
    Breprise: TToolbarButton97;
    POPZ: TPopupMenu;
    Label28: TLabel;
    cbDelimChamp: TComboBox;
    BRecherche: TToolbarButton97;
    tvchampinvisible: TTreeView;
    Checksyntaxe: TToolbarButton97;
    ModeParam: TRadioGroup;
    Rejet: TMenuItem;
    TabCorresp: TTabSheet;
    ListeProfile: TListBox;
    GroupBox7: TGroupBox;
    LTable: TLabel;
    Profile: THCritMaskEdit;
    bDefaire: TToolbarButton97;
    BValide: TToolbarButton97;
    SpeedButton37: TSpeedButton;
    SpeedButton38: TSpeedButton;
    SpeedButton39: TSpeedButton;
    SpeedButton40: TSpeedButton;
    SpeedButton41: TSpeedButton;
    SpeedButton42: TSpeedButton;
    Affecterligne1: TMenuItem;
    Label13: TLabel;
    Nbdecimales: TSpinEdit;
    Label21: TLabel;
    cblibellenul: TComboBox;
    Label20: TLabel;
    ComboChampNul: TComboBox;
    Combochampnul1: TComboBox;
    Cacher: TCheckBox;
    cbFormatDatesortie: TComboBox;
    Label27: TLabel;
    Label29: TLabel;
    TabLien: TTabSheet;
    LIENINTER: THRichEditOLE;
    Chargement: TToolbarButton97;
    Label24: TLabel;
    cbTypeCar: TComboBox;
    Memo2: TMemo;
    ORDER: TEdit;
    Label30: TLabel;
    bOptions: TToolbarButton97;
    Label31: TLabel;
    FicheSortie: THCritMaskEdit;
    Label32: TLabel;
    Table: THCritMaskEdit;
    Label33: TLabel;
    TableN: THValComboBox;
    Label34: TLabel;
    CodeChamp: THValComboBox;
    BValideLien: TToolbarButton97;
    BEffLien: TToolbarButton97;
    GroupBox11: TGroupBox;
    Label35: TLabel;
    ToolbarButton971: TToolbarButton97;
    EdShellExecute: THCritMaskEdit;
    Famille: TEdit;
    Label36: TLabel;
    Label37: TLabel;
    Typetransfert: THCritMaskEdit;
    Label38: TLabel;
    Editeur: THCritMaskEdit;
    CbLngRef: TCheckBox;
    SpeedButton43: TSpeedButton;
    SpeedButton44: TSpeedButton;
    SpeedButton45: TSpeedButton;
    SpeedButton46: TSpeedButton;
    SpeedButton47: TSpeedButton;
    SpeedButton48: TSpeedButton;
    SpeedButton49: TSpeedButton;
    SpeedButton50: TSpeedButton;
    SpeedButton51: TSpeedButton;
    SpeedButton52: TSpeedButton;
    SpeedButton53: TSpeedButton;
    SpeedButton54: TSpeedButton;
    Interne: TCheckBox;
    MemoComment: TMemo;
    Imprimer1: TMenuItem;
    SpeedButton55: TSpeedButton;
    SpeedButton56: TSpeedButton;
    ApercuConditionComp: TToolbarButton97;
    lblChampCourant: TLabel;
    lblEnregCourant: TLabel;
    btnArboresCheminCourant: TToolbarButton97;
    Label40: TLabel;
    ARRONDI: TCheckBox;
    Domaine: THValComboBox;
    TPAYS: TLabel;
    BPAYS: THValComboBox;
    HPanel1: THPanel;
    Label14: TLabel;
    Label15: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ChargeFichier;
    procedure MontreSel;
    procedure Memo1CaretMove(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure tvChampMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure miNouveauClick(Sender: TObject);
    procedure Supprimer1Click(Sender: TObject);
    procedure FILENAMEChange(Sender: TObject);
    procedure tvChampDblClick(Sender: TObject);
    procedure ChampNomChange(Sender: TObject);
    procedure btnAffecterClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure tvChampClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FichesauveChange(Sender: TObject);
    procedure tvChampChange(Sender: TObject; Node: TTreeNode);
    procedure tvChampEditing(Sender: TObject; Node: TTreeNode;
      var AllowEdit: Boolean);
    procedure tvChampEnter(Sender: TObject);
    procedure tvChampKeyPress(Sender: TObject; var Key: Char);
    procedure ChargementClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure miNormalClick(Sender: TObject);
    procedure FichesauveElipsisClick(Sender: TObject);
    procedure memConditionEnter(Sender: TObject);
    procedure CalFormuleEnter(Sender: TObject);
    procedure SpeedButton23Click(Sender: TObject);
    procedure SpeedButton25Click(Sender: TObject);
    procedure SpeedButton29Click(Sender: TObject);
    procedure SpeedButton31Click(Sender: TObject);
    procedure lbChampNumDblClick(Sender: TObject);
    procedure tsCALCULEnter(Sender: TObject);
    procedure CalFormuleChange(Sender: TObject);
    procedure cbTypeChange(Sender: TObject);
    procedure CstValChange(Sender: TObject);
    procedure CstPasChange(Sender: TObject);
    procedure CstLonChange(Sender: TObject);
    procedure tvChampChanging(Sender: TObject; Node: TTreeNode;
      var AllowChange: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RadioGroup1Click(Sender: TObject);
    procedure edComplCarChange(Sender: TObject);
    procedure edComplLgnChange(Sender: TObject);
    procedure cbAlignLeftClick(Sender: TObject);
    procedure cbTableCorrClick(Sender: TObject);
    procedure cbTableExterneClick(Sender: TObject);
    procedure RefConditionChange(Sender: TObject);
    procedure RefPosterieurClick(Sender: TObject);
    procedure RefPositionChange(Sender: TObject);
    procedure RefLongueurChange(Sender: TObject);
    procedure EdRangChange(Sender: TObject);
    procedure bZoomPieceClick(Sender: TObject);
    procedure Memo1GetLine(Sender: TObject; var ALine: TASCIILine;
      var AColor: TColor; var AGlyph: Integer; ALineN: Integer);
    procedure Memo1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Nouvellepartie1Click(Sender: TObject);
    procedure Champnormal1Click(Sender: TObject);
    procedure miCalculClick(Sender: TObject);
    procedure miReferenceClick(Sender: TObject);
    procedure miConstanteClick(Sender: TObject);
    procedure CalFormuleDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure CalFormuleDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure CalFormuleExit(Sender: TObject);
    procedure Memo1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormResize(Sender: TObject);
    procedure RefConditionEnter(Sender: TObject);
    procedure RefConditionExit(Sender: TObject);
    procedure edLongueurChange(Sender: TObject);
    procedure ChampDebutChange(Sender: TObject);
    procedure edFinChange(Sender: TObject);
    procedure ChampLonChange(Sender: TObject);
    procedure cbFormatDateChange(Sender: TObject);
    procedure cbTouteValeurClick(Sender: TObject);
    procedure cbValeurPositiveClick(Sender: TObject);
    procedure cbValeurNegativeClick(Sender: TObject);
    procedure cblibellenulChange(Sender: TObject);
    procedure Combochampnul1Change(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure btnEditTblClick(Sender: TObject);
    procedure Scruter1Click(Sender: TObject);
    procedure cbComplementClick(Sender: TObject);
    procedure Copier1Click(Sender: TObject);
    procedure Coller1Click(Sender: TObject);
    procedure ViewFScript1Click(Sender: TObject);
    procedure bDownClick(Sender: TObject);
    procedure tvChampKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RTypefichierClick(Sender: TObject);
    procedure FichesauveExit(Sender: TObject);
    procedure tvChampDeletion(Sender: TObject; Node: TTreeNode);
    procedure Couper1Click(Sender: TObject);
    procedure shColorMouseUp(Sender: TObject; Button: TMouseButton;
                             Shift: TShiftState; X, Y: Integer);
    procedure DOMAINEChange(Sender: TObject);
    procedure BColorChange(Sender: TObject);
    procedure Memo3Enter(Sender: TObject);
    procedure Memo3Change(Sender: TObject);
    procedure SpeedButton16Click(Sender: TObject);
    procedure bConditionClick(Sender: TObject);
    procedure EnregUniqueClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure ChecksyntaxeClick(Sender: TObject);
    procedure ModeParamClick(Sender: TObject);
    procedure RejetClick(Sender: TObject);
    procedure ProfileElipsisClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure bDefaireClick(Sender: TObject);
    procedure tvchampinvisibleDeletion(Sender: TObject; Node: TTreeNode);
    procedure Affecterligne1Click(Sender: TObject);
    procedure NbdecimalesChange(Sender: TObject);
    procedure edNomTableChange(Sender: TObject);
    procedure cbFormatDatesortieChange(Sender: TObject);
    procedure LIENINTERChange(Sender: TObject);
    procedure cbTypeCarChange(Sender: TObject);
    procedure ORDERChange(Sender: TObject);
    procedure bOptionsClick(Sender: TObject);
    procedure FicheSortieChange(Sender: TObject);
    procedure RadiotableClick(Sender: TObject);
    procedure TableElipsisClick(Sender: TObject);
    procedure TableChange(Sender: TObject);
    procedure TabLienEnter(Sender: TObject);
    procedure TableNExit(Sender: TObject);
    procedure BValideLienClick(Sender: TObject);
    procedure BEffLienClick(Sender: TObject);
    procedure EdShellExecuteChange(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
    procedure FamilleChange(Sender: TObject);
    procedure CbLngRefClick(Sender: TObject);
    procedure RefPositionEnter(Sender: TObject);
    procedure RefLongueurEnter(Sender: TObject);
    procedure SpeedButton43Click(Sender: TObject);
    procedure InterneClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure Imprimer1Click(Sender: TObject);
    procedure SpeedButton55Click(Sender: TObject);
    procedure SpeedButton56Click(Sender: TObject);
    procedure ApercuConditionCompClick(Sender: TObject);
    procedure ARRONDIClick(Sender: TObject);

  private
    { Déclarations privées }
    FocusCtrl              : TComponent;
    Fscript                : TScript;
    PrevNodeMaitre         : TTreeNode;
    curChamp               : TChamp;
    FRefresh               : Boolean;
    curInfo                : integer;
    NomTable               : string;
    Comment                : TActionFiche;
    ChampTemp              : TChamp;
    ParametrageCegid       : boolean;
    edtNature,TypeT        : string;
    TranfertT,Logiciel     : string;
    BOkValide              : Boolean;
    Domainepar             : string;
    FindText               : string;
    GestionAuxi            : Boolean;
    ModeVisu               : integer;
    Endommage              : Boolean;
    Codeimport,SaveFiche   : string;
    SaveCurRow             : integer;
    NbreExercice           : integer;
    Statistique            : Boolean;
    Analytique             : Boolean;
    FScriptDel             : TScriptDelimite;
    ChargeTree             : Boolean;
    AMemoPrint             : TRichEdit;
    PrevChMaitre           : string;

    procedure MajImageIndex(TN : TTreeNode);
    procedure InsererText(S : String; M : TMemo);
    procedure ChargementTree(param : Boolean);
    procedure renseignetree (Achamp : TChamp; TT : TTreeNode; var vChamp : TTreeview);
    procedure tvChampDragOver(Sender, Source: TObject; X, Y: Integer;
			State: TDragState; var Accept: Boolean);
    procedure tvChampDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure tvPartieDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure tvPartieDragOver(Sender, Source: TObject; X, Y: Integer;
			State: TDragState; var Accept: Boolean);
    procedure ShowInfoPanel(TNS: TTreeNode);
    procedure memConditionChange(Sender: TObject);
    procedure tsChpConditionEnter(Sender: TObject);
    procedure tsComplementEnter(Sender: TObject);
    function  RendRacineParent(TN : TTreeNode; var ParentName :string) : integer;
    Function  ControleValidite : Boolean;
    procedure EnabledChamp (Enb : Boolean);
    procedure BMenuZoomClick(Sender: TObject);
    function  CreerLigPop ( Name : string; Owner : TComponent) : TMenuItem ;
    procedure BrepriseClick (Sender: TObject);
    procedure MontreSelCondition (Txt : String; Key: Word=0);
    procedure ChargementInfo;
    procedure ChargementCombotableN;
    procedure ChargementCodechamp(Name : string);
    procedure ValidClickRichedit(Sender: TObject);
    function  CheckVariables(chaine : string) : integer;
    function  ExistVariable(chaine : string; debut : integer; fin : integer) : boolean;
    procedure AfficherErreurSyntaxeVariable (erreur : integer);
    function RendLORDER : string;
  public
    { Déclarations publiques }
        FModified : Boolean;
	property Script : TScript read FScript;
        function InitialiseScript : Boolean;
        function LoadScript(AScriptName:String) : TScript;


  end;

  Procedure PExtConcept (Scp : TScript; dom,tablex,nature,typetr,editer : string; CommentAct : TActionFiche; TPC : TSVControle=nil; CpteAuxi:Boolean= FALSE; NbExerc:integer=1; Stat:Boolean=FALSE; shellexec:string=''; Analytiq: Boolean=FALSE) ;
  Procedure AppelCorresp (dom,tablex,nature,typetr,editer : string; CommentAct : TActionFiche; TPC : TSVControle=nil; CpteAuxi:Boolean= FALSE; NbExerc:integer=1; Stat:Boolean=FALSE; shellexec:string=''; Analytiq: Boolean=FALSE) ;

implementation
{$R *.RES}
{$R *.DFM}
uses
UDMIMP, UExecute,
UView,
{$IFNDEF EAGLCLIENT}
USelectScript,
{$ENDIF}
UCorresp, USCRUTER, UPDomaine
,UDefVar
,URecherche
,USelectCorresp, UOption
;

const level = 1; level2 = 2;
Const ColonneCoche : Integer = 2;
const crMyCursor = 5;
const lgcorr = 2;

type
Tabled = record
             code     : string;
             domaine     : string;
end;
var
Correspimp : array [0..lgcorr] of Tabled =
   (
   (code : 'TRA' ; domaine : 'C'),
   (code : 'TRT' ; domaine : 'S'),
   (code : 'SAV' ; domaine : 'M')
   );


{------------------------- Evenements ---------------------------}
Procedure AppelCorresp (dom,tablex,nature,typetr,editer : string; CommentAct : TActionFiche; TPC : TSVControle=nil; CpteAuxi:Boolean= FALSE; NbExerc:integer=1; Stat:Boolean=FALSE; shellexec:string=''; Analytiq: Boolean=FALSE) ;
var XX            : TExtPConcept ;
    Idx           : integer;
    Stream        : TStream;
    N             : Integer;
    P, CurrentDir : String;
    TCR           : TTableCorrRec;
    slVariable    : TStringList;
    F             : TextFile;
begin
   if (GetInfoVHCX.FAMILLE = '') then exit;
   Stream := nil;
   XX                    := TExtPConcept.Create(Application) ;
   With XX do
   begin
              Comment            := CommentAct;
              Domaine.itemindex  := XX.Domaine.items.IndexOf(dom);
              NomTable           := copy(tablex,1,12);
              edtNature          := Nature;
              GestionAuxi        := CpteAuxi;
              if typetr = 'Dossier' then TypeT := 'DOS' else
              if typetr = 'Journal' then TypeT := 'JRL' else
              if typetr = 'Balance' then TypeT := 'BAL' else
              if typetr = 'Synchronisattion' then TypeT := 'SYN' else
              if typetr =  '<Tous>' then TypeT := 'DOS' else
              if typetr =  'Ecriture' then  TypeT := 'JRL';
              if typetr = 'Intra-Groupe' then TypeT := 'GRP';
              TranfertT          := typetr;
              Logiciel           := Editer;
              NbreExercice       := NbExerc;
              Statistique        := Stat;
              Analytique         := Analytiq;
              EdShellExecute.text:= shellexec;
              Fscript := LoadScript (NomTable) ;
              Idx := FScript.Champ.IndexOf (UpperCase(GetInfoVHCX.FAMILLE), 0, TRUE);
              if (Idx <> -1) then
              begin
                    curchamp := TChamp(FScript.Champ[idx]);
                    if curChamp.TableCorr = nil then
                    begin
                      curChamp.TableCorr := FScript.NewTableCorr;
                      curChamp.TableCorr.FAssociee := true;
                    end;

                    if curChamp.TableCorr <> nil then
                    begin
                         CorrespDlg := TCorrespDlg.Create(nil);
                         CorrespDlg.Script := FScript;
                         CorrespDlg.Champ := curChamp;
                         CorrespDlg.OpTblCorr := curchamp.OpTblCorr;
                         slVariable := TStringList.Create;

                   try
                    (* récupération de la table à partir d'un fichier*)
                             if GetInfoVHCX.NomFichier = '' then
                                Curchamp.NomtableExt := Interprete(Curchamp.NomTableExt)
                             else
                                curChamp.NomTableExt := GetInfoVHCX.Directory + '\' +GetInfoVHCX.NomFichier;
                             curChamp.TableExterne := TRUE;
                             if GetInfoVHCX.Listefichier <> '' then
                                FScript.Options.FileName := GetInfoVHCX.Listefichier;

                             if not FileExists(curChamp.NomTableExt) then
                             begin
                                     AssignFile(f, curChamp.NomTableExt);
                                     Rewrite(f);
                                     CloseFile(f);
                             end;
                             if GetInfoVHCX.NomFichier = '' then
                                CorrespDlg.edNomTable.Text := edNomTable.Text
                             else
                                CorrespDlg.EdNomtable.Text  := GetInfoVHCX.Directory + '\' + GetInfoVHCX.NomFichier;

                              curChamp.IniTableCorr(slvariable);
                              TCR.FEntree := curChamp.TableCorr^.FEntreeExt;
                              TCR.FSortie := curChamp.TableCorr^.FSortieExt;
                              CorrespDlg.Table := @TCR;
                              CorrespDlg.Domaine.visible  := FALSE;
                              CorrespDlg.TDomaine.visible := FALSE;
                              CorrespDlg.Label3.visible   := FALSE;
                              CorrespDlg.Memo1.visible    := FALSE;
                              CorrespDlg.Label8.visible   := FALSE;
                              CorrespDlg.Profile.visible  := FALSE;
                              CorrespDlg.Label12.visible  := FALSE;
                              CorrespDlg.Famille.visible  := FALSE;
                              CorrespDlg.BReprise.visible  := FALSE;
                              if UpperCase(GetInfoVHCX.Appelant) = 'ECHANGES PGI'  then
                              begin
                                 if UpperCase(GetInfoVHCX.FAMILLE) = 'CODEJOURNAL' then
                                    CorrespDlg.Caption := 'CISX : Table de correspondance journaux'
                                 else
                                    CorrespDlg.Caption := 'CISX : Table de correspondance ' + GetInfoVHCX.FAMILLE
                              end;
                              curChamp.NomTableExt := CorrespDlg.EdNomtable.Text;
                              if CorrespDlg.ShowModal = mrOK then
                              begin
                                      FModified := True;
                                      Curchamp.OpTblCorr := CorrespDlg.OpTblCorr;
                                      curChamp.TableCorr^.FEntree.Assign(CorrespDlg.Table.FEntree);
                                      curChamp.TableCorr^.FSortie.Assign(CorrespDlg.Table.FSortie);
                                      for N:= 0 to CorrespDlg.Table.FEntree.count-1 do
                                      begin
                                      curChamp.TableCorr^.FEntree.Strings[N] :=  CorrespDlg.Table.FEntree[N];
                                      curChamp.TableCorr^.FSortie.Strings[N] :=  CorrespDlg.Table.FSortie[N];
                                      end;
                                      curChamp.TableCorr^.FCount := curChamp.TableCorr^.FEntree.Count;
                                      curchamp.commentTablecorr := CorrespDlg.Memo1.text; // Ne pas oublier d'être en phase avec la Dll (déclaration des champs du script)
                                      try
                                       if (curChamp.NomTableExt = '') then
                                       begin
                                            with Topendialog.create(Application) do
                                            begin
                                                 if Execute then
                                                 begin
                                                     curChamp.NomTableExt := FileName;
                                                     edNomTable.Text := FileName;
                                                 end;
                                                 Free;
                                            end;
                                            SetCurrentDirectory(PChar(CurrentDir));
                                       end;
                                       Stream := TFileStream.Create(RecomposeChemin(InterpreteVar(curChamp.NomTableExt, slvariable)), fmCreate);
                                       except
                                               PGiBox(Exception(ExceptObject).message, 'Erreur');
                                       end;

                                        // si la création a réussie
                                        if (Stream <> nil) then
                                        begin
                                            for N:=0 to CorrespDlg.Table.FEntree.Count-1 do
                                            begin
                                              P := CorrespDlg.Table.FEntree.Strings[N]+';'+
                                                     CorrespDlg.Table.FSortie.Strings[N]+#13+#10;
                                              Stream.Write(Pointer(P)^, Length(P));
                                            end;
                                            Stream.Free;
                                        end;
                              end;
                              finally
                                     CorrespDlg.Libere;
                                     CorrespDlg.Free;
                                     slVariable.Free;
                              end;
                   end;
              end;
              FScript.destroy;
              free;
   end; //With
end;

procedure TExtPConcept.FormCreate(Sender: TObject);
begin
    FScript         := nil;
    TPControle      := nil;
    ChargeTree      := FALSE;
    ChargeFichier;
{$IFNDEF CISXPGI}
    ChargementComboDomaine (Domaine);
    Domainepar := RendDomaine(Domaine.Text);
{$ENDIF}
end;

procedure TExtPConcept.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
   MM :THRichEditOLE;
begin
Case Key of
   VK_F10    : ModalResult:=mrOk ;
   VK_F5     :  // select
   begin
        if (FocusCtrl is THRichEditOLE) then
        begin
             MM := THRichEditOLE(FocusCtrl);
             MontreSelCondition(MM.Text, Key);
        end;

   end;
   VK_F3 :
   begin
        if FindText <> '' then
        begin
              Memo1.Find(FindText, foIgnoreCase, 0, 1);
              Memo1.LeftSel := 1;
              Memo1.RightSel := LNGLMAX;
              MontreSel;
        end;
   end;
   {CTRL+F} 70 : if Shift=[ssCtrl] then begin Key:=0 ; BRechercheClick(Sender) ; end ;
  END ;
end;


procedure TExtPConcept.MontreSelCondition (Txt : String; Key: Word=0);
var
     Ligne: string;
     i      : integer;
     signe: string;
     St   : String;
     NbLigne   : String;
     ldeb,lfin : string;
begin
        Ligne := Txt;
        if pos('=', Ligne) <= 0  then exit;
        i := pos('F', Ligne);
        if (i <> 0)  and (Uppercase(Copy(Ligne, i, 6)) = 'FLIGNE') then
                i := pos('(', Ligne)
        else
        begin
                i := pos('L', Ligne);
                if (i <> 0)  and (Uppercase(Copy(Ligne, i, 5)) = 'LIGNE') then
                      i := pos('(', Ligne)
                else if Key <> VK_F5 then exit;
        end;
        if i <> 0 then
        begin
         i := pos('+', Ligne);
         if i <> 0 then signe := copy(Ligne, i,1);
         i := pos('-', Ligne);
         if i <> 0 then signe := copy(Ligne, i,1);
         St := ReadTokenPipe (Ligne, signe);
         if pos(')', Ligne) <> 0 then
         begin
               NbLigne := ReadTokenPipe (Ligne, ')');
               i := 0;
               if (NbLigne <> '') and (NbLigne[1] in ['1'..'9'])then
               begin
                   if signe = '-' then i := StrToInt(Nbligne)
                   else i := (-1)*StrToInt(Nbligne);
               end;
         end;
        end;
        St := ReadTokenPipe (Ligne, '[');
        ldeb := ReadTokenPipe (Ligne, ',');
        lfin := ReadTokenPipe (Ligne, ']');
        if (ldeb = '') or (lfin = '') then exit;
        St := ReadTokenPipe (Ligne, '''');
        St := ReadTokenPipe (Ligne, '''');
        if St <> '' then
        begin
             Memo1.Find(St, foFromTop, i);
             Memo1.LeftSel := 1;
             Memo1.RightSel := 4096;
             MontreSel;
        end;
end;

procedure TExtPConcept.ChargeFichier;
begin
        Memo1.FileName :=FILENAME.Text;
	Memo1.Ignored := 0;
	Memo1.Hint := FILENAME.Text;
        Memo1.UpdateLines;
end;

procedure TExtPConcept.MontreSel;
begin
	with Memo1 do begin
		if LeftSel >= 0 then begin
			lbDeb.Caption := IntToStr(LeftSel+1);
			lbLon.Caption := IntToStr(RightSel-LeftSel);
			lbFin.Caption := IntToStr(RightSel);
		end
		else begin
			lbDeb.Caption := '';
			lbFin.Caption := '';
			lbLon.Caption := '';
		end;
		Label14.Caption := 'L:'+IntToStr(curRow+1)+' C:'+IntToStr(curCol+1);
		Label15.Caption := 'L:'+IntToStr(curRow+1)+' C:'+IntToStr(curCol+1);
                TopLine := curRow;
	end;
end;


procedure TExtPConcept.Memo1CaretMove(Sender: TObject);
Begin
  MontreSel;
end;

Procedure PExtConcept (Scp : TScript; dom,tablex,nature,typetr,editer : string; CommentAct : TActionFiche; TPC : TSVControle=nil; CpteAuxi:Boolean= FALSE; NbExerc:integer=1; Stat:Boolean=FALSE; shellexec:string=''; Analytiq: Boolean=FALSE) ;
var XX : TExtPConcept ;
    PP : THPanel ;
    procedure Alimtypetransfert;
    begin
         With XX do
         begin
              Comment            := CommentAct;
{$IFDEF CISXPGI}
              Domaine.dataType  := 'CPZIMPDOMAINE';
{$ENDIF}
              Domaine.itemindex  := XX.Domaine.items.IndexOf(dom);
              NomTable           := copy(tablex,1,12);
              edtNature          := Nature;
              GestionAuxi        := CpteAuxi;
              if typetr = 'Dossier' then TypeT := 'DOS' else
              if typetr = 'Journal' then TypeT := 'JRL' else
              if typetr = 'Balance' then TypeT := 'BAL' else
              if typetr = 'Synchronisattion' then TypeT := 'SYN' else
              if typetr =  '<Tous>' then TypeT := 'DOS' else
              if typetr =  'Ecriture' then  TypeT := 'JRL';
              if typetr = 'Intra-Groupe' then TypeT := 'GRP';
              TranfertT          := typetr;
              Logiciel           := Editer;
              NbreExercice       := NbExerc;
              Statistique        := Stat;
              Analytique         := Analytiq;
              EdShellExecute.text:= shellexec;
         end;
    end;
BEGIN
XX                    := TExtPConcept.Create(Application) ;
Alimtypetransfert;

if TPC <> nil then TPControle := TPC;

if Scp <> nil then
begin
     XX.FScript       := Scp.CloneScript;
     XX.FScript.FComment := Scp.FComment;
     XX.FScript.Assign(Scp, '');
end;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     XX.ShowModal ;
    Finally
     XX.Free;
    End;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;

END ;

procedure TExtPConcept.PopupMenu1Popup(Sender: TObject);
var Node : TTreeNode;
begin
     Node := tvchamp.selected;
     concat1.enabled := (Node.data <> nil);
     type1.enabled := (Node.data <> nil);
     copier1.enabled := (Node.data <> nil);
     couper1.enabled := (Node.data <> nil);
     coller1.enabled := (Node.data <> nil);
     Imprimer1.Enabled := (Node.data <> nil);
     Supprimer1.Enabled := (Fscript.Modeparam = 0);
     Nouveau1.Enabled := (Fscript.Modeparam = 0);
end;

procedure TExtPConcept.tvChampMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
    TN    : TTreeNode;
    pt    : TPoint;
begin
	if ssRight in Shift then
	begin
            TN := tvChamp.GetNodeAt(X, Y);
            if TN <> nil then TN.Selected := True;
            pt.X := X; pt.Y := Y;
            pt := tvChamp.ClientToScreen(pt);
            PopupMenu1.Popup(pt.X, pt.Y);
	end;
end;

procedure TExtPConcept.miNouveauClick(Sender: TObject);
var
    AChamp    : TChamp;
    ANode     : TChampNode;
    TN        : TTreenode;
    idx       : integer;
    Name      : string;
    TNP       : TTreenode;
    ParentName: string;
begin
    if tvChamp.Selected.level = level then exit
//         Name := tvChamp.Selected.Text
    else
    begin
          ParentName := '';
          TNP := tvChamp.selected.Parent;
          // pour sauvegarder le nom avec Parent.Champ
          RendRacineParent(TNP, ParentName);

          FModified := True;

          if tvChamp.Selected.Data <> nil then Name := curchamp.name
          else
          if tvChamp.Selected.level = level then Name := ChampNom.Text;
    end;
    tvChamp.Items.BeginUpdate;

    Idx := FScript.Champ.IndexOf (Name);

    if (Idx <> -1) then
            AChamp := TChamp(FScript.Champ.Insert(idx))
    else
            AChamp := TChamp.Create(FScript.Champ);
    if Idx = 0 then
       Name := 'Parametrage'
    else
       Name :='Champ'+IntToStr(FScript.Champ.Count+1);

    AChamp.Name := ParentName+Name;
    if tvChamp.Selected.level = level then
       AChamp.Sel := 16
    else
       AChamp.Sel := 20;

    ANode := TChampNode.Create(Self);
    ANode.FChamp := AChamp;

    if tvChamp.Selected.level = level then
    begin
                  if (Idx <> -1) then
                     TN := tvChamp.Items.insertObject(tvChamp.Selected, Name, nil)
                  else
                     TN := tvChamp.Items.addObject(tvChamp.Selected, Name, nil);
                  TN :=tvChamp.Items.AddChildObject(TN, Name, ANode);
    end
    else
    begin
          if (Idx <> -1) then
             TN := tvChamp.Items.insertObject(tvChamp.Selected, Name, ANode)
          else
             TN := tvChamp.Items.addObject(tvChamp.Selected, Name, ANode);
    end;
    MajImageIndex (TN);
    tvChamp.Items.EndUpdate;
end;

procedure TExtPConcept.Supprimer1Click(Sender: TObject);
var
  Idx       : integer;
  TLec1,CC  : TChamp;
  Oksupp    : Boolean;
  TNprec    : TTreeNode;
  Name      : string;
	function idxsupof(str :  string):integer;
	var id,i : integer;
        NN   : string;
        ind  : integer;
	begin
	  result := -1;
          ind := 0;
	  for id := 0 to FScript.Champ.Count -1 do
	  begin
            NN := FScript.Champ.Items[id].Name;
            i := pos('_', FScript.Champ.Items[id].Name);
            if i <> 0 then NN := copy(FScript.Champ.Items[id].Name, i+1, Length(FScript.Champ.Items[id].Name));
            if NN = str then
            begin
	      ind := id;
              Result := ind;
	      break;
            end;
	  end;

        // Pour supprimer les enfants et la liste concat
        if (result = -1) then
        begin
             if ind <> 0 then id := ind
             else id := 0;
             Oksupp := FALSE;
             TLec1 := FScript.Champ[id];
             if (TLec1.TypeInfo = tiConcat) and (TLec1.Concat.count <> 0) and (TLec1.Concat <> nil) then
             begin
                  CC := TLec1.Concat[0];
                  i := 0;
                  while CC <> nil do
                  begin
                          TLec1.Concat.Delete(i);
                          if i < TLec1.Concat.Count-1 then  CC := TLec1.Concat[i]
                          else CC := nil;
                          inc (id);
                          Oksupp := TRUE;
                  end;
             end
             else
             while TLec1 <> nil  do
             begin
                  NN := FScript.Champ.Items[id].Name;
                  i := pos('_', FScript.Champ.Items[id].Name);
                  if i <> 0 then NN := copy(FScript.Champ.Items[id].Name, 1, i-1);
			      if NN = str then
                  begin
                     if FScript.Champ.Items[id].TableExterne then
                     begin
                          if FScript.Champ.Items[id].TableCorr^.FEntreeExt <> nil then   FScript.Champ.Items[id].TableCorr^.FentreeExt.Clear;
                          if FScript.Champ.Items[id].TableCorr^.FSortieExt <> nil then   FScript.Champ.Items[id].TableCorr^.FSortieExt.Clear;
                     end;
                     FScript.Champ.Delete(id); dec(id); Oksupp := TRUE;
                  end;
                  if id < FScript.Champ.Count-1 then
                  begin
                       if id < 0 then TLec1 := FScript.Champ[0]
                       else TLec1 := FScript.Champ[id];
                  end
                  else TLec1 := nil;
                  inc (id);
              end;
              if Oksupp then
                 ChampNom.Text := tvChamp.Items.Item[tvChamp.selected.absoluteIndex].Text;
        end;
    end;

begin
       if (tvChamp.Selected.Data <> nil)  and (tvChamp.Selected.level <> level) then
       begin
         if ((TchampNode(tvChamp.Selected.Data).Fchamp.sel and 4) = 0) then
         begin
                 PGIBox ('Suppression impossible', 'Champ obligatoire');
                 exit;
         end;
       end;
       TNprec := tvChamp.Selected.Getprev;
       if ChampNom.Text = '' then  ChampNom.Text := TChampNode(tvChamp.Selected.Data).FChamp.Name;
       if (not TvChamp.selected.HasChildren) or (tvChamp.Selected.level = level2) then
        Name := TChampNode(tvChamp.Selected.Data).FChamp.Name;

       if (tvChamp.Selected.level <> level2) or ((not TvChamp.selected.HasChildren) and (tvChamp.Selected.level = level2))then
          tvChamp.Selected.Delete;

       if (not TvChamp.selected.HasChildren) then
        Idx := FScript.IndexOfChamp(Name)
       else
        Idx := idxsupof(ChampNom.Text);

       if (tvChamp.Selected.level = level2) and (TvChamp.selected.HasChildren) then  // pour concat
          tvChamp.Selected.Delete;

       if Idx <> -1 then
       begin
                     if FScript.Champ.Items[Idx].TableExterne then
                     begin
                          if FScript.Champ.Items[Idx].TableCorr^.FEntreeExt <> nil then   FScript.Champ.Items[Idx].TableCorr^.FentreeExt.Clear;
                          if FScript.Champ.Items[Idx].TableCorr^.FSortieExt <> nil then   FScript.Champ.Items[Idx].TableCorr^.FSortieExt.Clear;
                     end;
                     FScript.Champ.Delete(Idx);
       end;
       tvchamp.Update;
       if (not TvChamp.selected.HasChildren) and (tvChamp.Selected.Data <> nil) then
//           TChampNode(tvChamp.Selected.Data).FChamp.TypeInfo := tiNull
       else
       begin
            if TNprec.parent <> nil then
            begin
                 tvChamp.Selected := tvChamp.Items[TNprec.parent.absoluteIndex];
                 ChampNom.Text := tvChamp.Items.Item[TNprec.parent.absoluteIndex].Text;
            end;
       end;
end;

procedure TExtPConcept.FILENAMEChange(Sender: TObject);
begin
     ChargeFichier;
end;

procedure TExtPConcept.MajImageIndex(TN : TTreeNode);
var
	TN1 : TTreeNode;

	procedure majParent(TNP : TTreeNode);
	var AllSel, AnySel, SomeSel : boolean;
		I : integer;
		TN2 : TTreeNode;
	begin
		TN2 := TNP.GetFirstChild;
		AllSel := True;
		SomeSel := False;
		anySel := True;
		for I := 0 to TNP.Count-1 do
		   if (TN2 <> nil) and assigned(TN2.Data) then
		   begin
			 if ((TchampNode(TN2.Data).Fchamp.sel and 4) = 0) then
			 begin
				SomeSel := true;
				AnySel := false;
			 end
			 else AllSel := False;
			 TN2 := TNP.GetNextChild(TN2);
		   end;
		if AllSel then
		begin
			 TNP.ImageIndex := 5;
             TNP.SelectedIndex := 5;
        end
        else if AnySel then
        begin
             TNP.ImageIndex := 7;
			 TNP.SelectedIndex := 7;
             if TNP.data <> nil then TchampNode(TNP.data).Fchamp.sel := 20;
        end
		else if Somesel then
        begin
             TNP.ImageIndex := 6;
             TNP.SelectedIndex := 6;
		end;

    end;

    procedure cocheT(TT : TTreeNode);
    var I : integer;
    begin
       if (TT.ImageIndex = 5) or (TT.ImageIndex = 6) then
       begin
          TN1 := TT.getfirstChild;
          for I:=0 to TT.count-1 do
          begin
              if TN1  <> nil then
              begin
                   TchampNode(TN1.data).Fchamp.sel := 20;
				   MajImageIndex(TN1);
              end;
              TN1 := TT.GetNextChild(TN1);
          end;
          TT.ImageIndex := 7; TT.SelectedIndex := 7;
       end
       else
       begin
          TN1 := TT.getfirstChild;
          for I:=0 to TN.count-1 do
          begin
			  if TN1  <> nil then
              begin
                   TchampNode(TN1.data).Fchamp.sel := 16;
                   MajImageIndex(TN1);
              end;
              TN1 := TT.GetNextChild(TN1);
          end;
          TT.ImageIndex := 5; TT.SelectedIndex := 5;
       end;
    end;
begin
    if TN.data <> nil then
    begin
        case (TChampNode(TN.Data).Fchamp.Sel and 12) of
            0: begin
				TN.ImageIndex := 0;
                TN.SelectedIndex := 0;
                end;
            4: begin
                TN.ImageIndex := 2;
                TN.SelectedIndex := 2;
                end;
            8: begin
                TN.ImageIndex := 3;
                TN.SelectedIndex := 3;
                end;
            12: begin
                TN.ImageIndex := 4;
                TN.SelectedIndex := 4;
                end;
        end;
    end
    else MajParent(TN);
    if TN.absoluteIndex = 0 then
    begin
              TN.ImageIndex := 9;
              TN.SelectedIndex := 9;
    end;

    if TN.Parent <> nil then
          MajParent(TN.Parent)
    else if TN.data = nil then
      CocheT(TN);
end;

procedure TExtPConcept.tvChampDblClick(Sender: TObject);
var Pt : TPoint;
index  : integer;
begin
	if FRefresh or not Assigned(curChamp) then exit;
	(* recherche la position de la souris *)
	GetCursorPos(pt);
	pt := tvChamp.ScreenToClient(pt);
	if Pt.x < 20 then exit;
    if tvchamp.Selected.data <> nil then
    begin
        if (curChamp.Sel and 4) <> 4 then begin
            curChamp.Sel := (curChamp.Sel or 4);
        end
        else begin
            curChamp.Sel := (curChamp.Sel and not 4);
        end;
    end;
    MajImageIndex(TvChamp.Selected);
    tvChamp.Invalidate;
    tvChamp.Update;
    if Tvchamp.selected.SelectedIndex = 2  then Memo1.Enabled := FALSE
    else Memo1.Enabled := TRUE;

    index := tvChamp.Selected.AbsoluteIndex;
    if index > 0 then
    ChampNom.Text := tvChamp.Items.Item[index].Text;

end;

procedure TExtPConcept.ChampNomChange(Sender: TObject);
var PN        : string;
    T1        : TOB;
    Fige      : Boolean;
    Comme     : string;
    function GetNamep(TT : TTreeNode) : string;
    var
       TN1    : TTreeNode;
       Name   : string;
    begin
          Name := '';
          TN1 := TT.Parent.GetFirstChild;
          if (TN1 <> nil) and (TN1.Data <> nil) then
          begin
             Name := (TchampNode(TN1.Data).Fchamp.Name);
             Name := ReadTokenPipe (Name, '_');
          end;
          Result := Name;
    end;
begin
    if FRefresh or not Assigned(curChamp) then exit;
    FModified := True;
    if (tvChamp.selected.Parent <> nil) and (tvChamp.Selected.level = level2) then
    begin
      PN := GetNamep(tvChamp.selected);
      if PN <> '' then
      begin
           lblChampCourant.Caption := ChampNom.Text;
           curChamp.Name := PN+'_'+ChampNom.Text;
           if TPControle.TOBParam <> nil then
           begin
               if Domainepar = 'X' then
                   T1 := TPControle.TOBParam.FindFirst(['TableName','Nom'], [PN, ChampNom.Text], FALSE)
               else
                   T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [Domainepar, PN, ChampNom.Text], FALSE);
               if T1 <> nil then
               begin
{$IFNDEF CISXPGI}
                        Fige := (T1.Getvalue ('Fige')=0);
{$ELSE}
                        Fige := TRUE;
{$ENDIF}
                         if Fige and (Tvchamp.selected.SelectedIndex = 2) then Fige := FALSE;
                         EnabledChamp (Fige);
                         if (T1.Getvalue ('Commentaire') <> '') then
                         Comme := 'Commentaire : ' + T1.Getvalue ('Commentaire');
                         if (T1.Getvalue ('Contenuoblig') <> '') then
                         begin
                              if Comme <> '' then Comme := Comme + '      Valeurs possibles : ' +  T1.Getvalue ('Contenuoblig')
                              else Comme := 'Valeurs possibles : ' +  T1.Getvalue ('Contenuoblig');
                         end;
                         MemoComment.Text := Comme;
               end;
           end;
      end;
    end
    else
    if tvChamp.selected.level <> level then
           curChamp.Name := ChampNom.Text;
    tvChamp.Selected.Text := ChampNom.Text;

end;

procedure TExtPConcept.btnAffecterClick(Sender: TObject);
var
	AChamp : TChamp;
	aPos : array [0..256] of integer;
	N : integer;
	p : pchar;

  function CalculDebut(iLongueur:Integer) : Integer;
  var uneLigne : String;
		I : Integer;
  begin
	 if FScript.ASCIIMODE = 0 then
	 begin
		Result := iLongueur;
	 end
	 else
	 begin	(* a faire *)
		uneLigne := Memo1.CurLine;
		p := PChar(uneLigne);
		aPos[0] := -1;
		N := 1;
		while P^ <> #0 do
		begin
			while (P^ <> #0) and (P^ <> FScript.Options.SepField) do inc(P);
			if P^ <> #0 then Inc(P);
			aPos[N] := integer(P-PChar(uneLigne));
			Inc(N);
		end;
		if iLongueur > Length(uneLigne) then
			Result := N-2
		else begin
			I := 0;
			while I < N do
			begin
				if aPos[I] > iLongueur then break;
				Inc(I)
			end;
			Result := I-1;
		end;
	 end;
  end;
begin
  if tvchamp.selected = nil then exit;
  if Tvchamp.selected.data = nil then
  begin
    if (FocusCtrl is TCustomMemo) or (FocusCtrl is TCustomEdit) or
       (FocusCtrl is TCustomRichEdit) or (FocusCtrl is THRichEditOLE) then
    begin
      Clipboard.Clear;
      try Clipboard.AsText := Format(' [%d,%d] ', [CalculDebut(StrToInt(lbDeb.Caption)), StrToInt(lbLon.Caption)]);
      except Clipboard.AsText := Format(' [%d,%d] ', [CalculDebut(Memo1.CurCol), 20]); end;
      TCustomEdit(FocusCtrl).PasteFromClipboard;
    end;
    Fmodified := true;
    exit;
  end;
  Fmodified := true;
  if (FocusCtrl is TCustomMemo) or (FocusCtrl is TCustomEdit) or
     (FocusCtrl is TCustomRichEdit) or (FocusCtrl is THRichEditOLE) then
  begin
  if (PageControl1.ActivePage = tsREFERENCE) and (FocusCtrl.Name <> 'RefCondition') then
  begin
    AChamp := TChampNode(tvChamp.Selected.Data).FChamp;
    AChamp.Deb       := StrToInt(lbDeb.Caption);
    AChamp.Lon       := StrToInt(lbLon.Caption);
    RefPosition.Text := lbDeb.Caption;
    RefLongueur.Text := lbLon.Caption;
  end
  else
  begin
    Clipboard.Clear;
    try
      Clipboard.AsText := Format(' [%d,%d] ', [CalculDebut(StrToInt(lbDeb.Caption)), StrToInt(lbLon.Caption)]);
    except
      Clipboard.AsText := Format(' [%d,%d] ', [CalculDebut(Memo1.CurCol), 20]); end;
      TCustomEdit(FocusCtrl).PasteFromClipboard;
    end ;
  end
  else
  begin
    if (PageControl1.ActivePage = tsREFERENCE) and (FocusCtrl.Name = 'RefCondition')then
    begin
      RefCondition.Text := Format(' [%d,%d] = ', [CalculDebut(StrToInt(lbDeb.Caption)), StrToInt(lbLon.Caption)]);
    end
    else
    begin
      AChamp := TChampNode(tvChamp.Selected.Data).FChamp;
      if lbDeb.Caption <> '' then
      begin
        if AChamp.Constante = nil then
        AChamp.Constante := TConstante.Create;
                  AChamp.Deb := CalculDebut(StrToInt(lbDeb.Caption)-1);
                  AChamp.Lon := StrToInt(lbLon.Caption);
              end
              else
              begin
                  AChamp.Deb := CalculDebut(Memo1.CurCol);
                  try AChamp.Lon := StrToInt(lbLon.Caption);
                  except AChamp.Lon := 20; end;
              end;
              if (PageControl1.ActivePage = tsREFERENCE) then
              begin
                  RefPosition.Text := lbDeb.Caption;
                  RefLongueur.Text := lbLon.Caption;
              end;
              ShowInfoPanel(tvChamp.Selected);
        end;
	end;
{ à voir si positionnement sur champ suivant
	if not (FocusCtrl is TCustomMemo) and not (FocusCtrl is TCustomEdit) and not (FocusCtrl is TCustomRichEdit) then
    begin
		index1 := tvChamp.Selected.AbsoluteIndex + 1;
		try tvChamp.Selected := tvChamp.Items[index1]; except; end;
	end;
}
end;


procedure TExtPConcept.FormShow(Sender: TObject);
var
Pass      : string;
begin
    tbDEFINITION.TabVisible             := False;
    tsCONSTANTE.TabVisible              := False;
    tsCALCUL.TabVisible                 := False;
    tsREFERENCE.TabVisible              := False;
    TabConditioncomp.TabVisible         := False;
    TabCorresp.Tabvisible               := False;
    TabLien.Tabvisible                  := False;
    tsCOMPLEMENT.TabVisible             := False;
    tsCoherence.TabVisible              := False;
    Tabparam.TabVisible                 := TRUE;
    BOkValide                           := FALSE;
    Endommage                           := FALSE;
    MemCondition.Text := '';
    PageControl1.ActivePage := Tabparam;
    {$IFDEF CISXPGI}
        Pass := CryptageSt(DayPass(Date));
    {$ELSE}
        Pass := DayPass(Date);
    {$ENDIF}
    if V_PGI.PassWord=Pass then ParametrageCegid := TRUE
    else ParametrageCegid := FALSE;


    ChargementTree(TRUE);
    if (FScript = nil) then Exit;
    if NomTable <> '' then
    begin
       Fichesauve.text    := NomTable;
       edNature.text      := edtNature;
       Typetransfert.text := TranfertT;
       Editeur.text       := Logiciel;

       if Comment= tacreat then
          Edcomment.Text := FScript.FComment;
        if FScript.Options.TypeCar =  tcOEM then
           cbTypeCar.ItemIndex := 0
        else
           cbTypeCar.ItemIndex := 1;
       ChargementInfo;
       //ChargementClick (Sender);
    end
    else exit;
    tvChamp.TopItem.Expand(False);

    Memo1.SelColor := StringToColor(GetSynRegKey('CouleurPinceau','',TRUE));  // couleur de sélection
    BColor.CurrentChoix := Memo1.SelColor;           // Couleur sur le bouton de sélection de couleur

    Screen.Cursors[crMyCursor] := LoadCursor(HInstance, 'STABAQUA');
    Memo1.Cursor := crMyCursor;

    BrepriseClick (Sender);
    bSelectAll.Enabled := (Fscript.Modeparam = 0);
    bInsert.Enabled := (Fscript.Modeparam = 0);
    bDelete.Enabled := (Fscript.Modeparam = 0);
    bReprise.Enabled := (Fscript.Modeparam = 0);
    if(Fscript.Modeparam = 1) then tvchamp.DragMode := dmManual;
    if (FScript.NiveauAcces = 0)then
         Caption :=   'STANDARD CEGID : '+ FScript.Name;
    FicheSortie.visible := (Radiotable.ItemIndex = 1);
    Label31.visible     := (Radiotable.ItemIndex = 1);
    Table.visible := (Radiotable.ItemIndex = 1);
    Label32.visible     := (Radiotable.ItemIndex = 1);
  if  (ParamCount > 0) and (GetInfoVHCX.Domaine <> '') then
          Domaine.Enabled := FALSE;

end;

{$IFNDEF CISXPGI}
function TExtPConcept.LoadScript(AScriptName:String) : TScript;
var
	Stream1, AStreamTable : TmemoryStream;
	ABlobField            : TField;
	ATblCorField          : TBlobField;
	AScript               : TScript;
{$IFDEF  DBXPRESS}
	tbCharger             : TADOTable;
  Cmd                   : TADOCommand;
{$ELSE}
	tbCharger             : TTable;
{$ENDIF}
begin
    Result := nil;
    Stream1 := nil;
    if not Assigned(DMImport) then DMImport := TDMImport.Create(Application);
    tbCharger := DMImport.GzImpReq;
{$IFDEF  DBXPRESS}
(*    Cmd := DMImport.Cmd;
    Cmd.ConnectionString := DMImport.DbGlobal.ConnectionString;
    Cmd.CommandText := 'SELECT * FROM PGZIMPREQ WHERE CLE="'+ AScriptName +'"';
    Cmd.Execute;
*)
    tbCharger.TableName:= 'PGZIMPREQ';
    tbCharger.Open;
{$ENDIF}
    with tbCharger do begin
{$IFDEF  DBXPRESS}
            if not Locate('CLE', AScriptName, [loCaseInsensitive]) then
{$ELSE}
            if not Active then Open;
            if not FindField([AScriptName]) then
{$ENDIF}
            begin
                showMessage('La requête '+aScriptName+' n''est pas une requête personnalisée');
            end;
            ABlobField := FieldByName('PARAMETRES');
            ATblCorField := FieldByName('TBLCOR') as TBlobField;
            try
                Stream1 := TmemoryStream.create;
                TBlobField(ABlobField).SaveToStream (Stream1);
                Stream1.Seek (0,0);
            except
                if (PGIAsk('Incohérences dans le Script, voulez-vous remonter la sauvegarde : ?',
                'Conception')= mryes) then
                begin
                      Stream1.free;
                      Stream1 := TmemoryStream.create;
                      Stream1.LoadFromFile (CurrentDossier+'\sauvegarde\'+AScriptName+'parametres.txt');
                      Stream1.Seek (0,0);
                      AScript := LoadScriptFromStream(Stream1, nil);
                      Result := Ascript;
                      Stream1.Free;
                      Close;
                      Endommage := TRUE;
                      exit;
                end
                else
                begin
                    Stream1.Free;
                    ShowMessage('Le script est endommagé.');
                    Close;
                    exit;
                end;
            end;
            AStreamTable := TmemoryStream.create;
            TBlobField(ATblCorField).SaveToStream (AStreamTable);
            AStreamTable.Seek (0,0);

            try AScript := LoadScriptFromStream(Stream1, AStreamTable);
            except begin
                    if (PGIAsk('Des incohérences dans le Script, voulez-vous remonter la sauvegarde : ?',
                    'Conception')= mryes) then
                    begin
                          Stream1.free;
                          Stream1 := TmemoryStream.create;
                          Stream1.LoadFromFile (CurrentDossier+'\sauvegarde\'+AScriptName+'parametres.txt');
                          Stream1.Seek (0,0);
                          AScript := LoadScriptFromStream(Stream1, nil);
                          Result := Ascript;
                          Stream1.Free;
                          AStreamTable.Free;
                          Close;
                          Endommage := TRUE;
                          exit;
                    end
                    else
                    begin
                          AStreamTable.Free;
                          Stream1.Free;
                          ShowMessage('Le script est endommagé.');
                          Close;
                          exit;
                    end;
                end;
            end;
            AScript.Destination := poGlobal;
            edtNature       := FieldByName('ClePar').asstring;
            TranfertT       := FieldByName('Table0').asstring;
            Logiciel        := FieldByName('Table1').asstring;
            AStreamTable.Free;
            Stream1.Free;

            AScript.FComment := FieldByName('Comment').AsString;
            try
                AScript.NiveauAcces := FieldByName('MODIFIABLE').AsInteger;
            except
                AScript.NiveauAcces := 0;    // non modifiable
            end;
            Close;
    end;
	Result := Ascript;
end;
{$ELSE}

function TExtPConcept.LoadScript(AScriptName:String) : TScript;
var
	Stream1, AStreamTable : TmemoryStream;
	ABlobField            : TField;
	ATblCorField          : TBlobField;
	AScript               : TScript;
	tbCharger,Q           : TQuery;
begin
    Result := nil;
    Stream1 := nil;
    tbCharger := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+AScriptName+'"', TRUE);
    if tbCharger.EOF then
                showMessage('La requête '+aScriptName+' n''est pas une requête personnalisée');
    with tbCharger do begin
{$IFNDEF EAGLCLIENT}
            ABlobField := FieldByName('CIS_PARAMETRES') as TBlobField;
{$ELSE}
            ABlobField := FindField('CIS_PARAMETRES');
{$ENDIF}
            //ATblCorField := FieldByName('CIS_TBLCOR') as TBlobField;
            try
                Stream1 := TmemoryStream.create;
                TBlobField(ABlobField).SaveToStream (Stream1);
                Stream1.Seek (0,0);
            except
                if (PGIAsk('Incohérences dans le Script, voulez-vous remonter la sauvegarde : ?',
                'Conception')= mryes) then
                begin
                      Stream1.free;
                      Stream1 := TmemoryStream.create;
                      Stream1.LoadFromFile (CurrentDossier+'\sauvegarde\'+AScriptName+'parametres.txt');
                      Stream1.Seek (0,0);
                      AScript := LoadScriptFromStream(Stream1, nil);
                      Result := Ascript;
                      Stream1.Free;
                      Ferme(tbCharger);
                      Endommage := TRUE;
                      exit;
                end
                else
                begin
                    Stream1.Free;
                    ShowMessage('Le script est endommagé.');
                     Ferme(tbCharger);
                    exit;
                end;
            end;
(*
            AStreamTable := TmemoryStream.create;
            TBlobField(ATblCorField).SaveToStream (AStreamTable);
            AStreamTable.Seek (0,0);
*)
            AStreamTable := TmemoryStream.create;
            Q := OpenSQl ('SELECT * FROM CPGZIMPCORRESP Where CIC_TABLEBLOB="CIS" and CIC_IDENTIFIANT="'+AScriptName+'"', FALSE);
            if not Q.EOF then
            begin
{$IFNDEF EAGLCLIENT}
                 ATblCorField := Q.Findfield('CIC_OBJET') as TBlobField;
                 TBlobField(ATblCorField).SaveToStream (AStreamTable);
{$ELSE}
                 ABlobField := Q.Findfield('CIC_OBJET');
                 TBlobField(ABlobField).SaveToStream (AStreamTable);
{$ENDIF}
                 AStreamTable.Seek (0,0);
            end;
            Ferme(Q);


            try AScript := LoadScriptFromStream(Stream1, AStreamTable);
            except begin
                    if (PGIAsk('Des incohérences dans le Script, voulez-vous remonter la sauvegarde : ?',
                    'Conception')= mryes) then
                    begin
                          Stream1.free;
                          Stream1 := TmemoryStream.create;
                          Stream1.LoadFromFile (CurrentDossier+'\sauvegarde\'+AScriptName+'parametres.txt');
                          Stream1.Seek (0,0);
                          AScript := LoadScriptFromStream(Stream1, nil);
                          Result := Ascript;
                          Stream1.Free;
                          AStreamTable.Free;
                          Close;
                          Endommage := TRUE;
                          exit;
                    end
                    else
                    begin
                          AStreamTable.Free;
                          Stream1.Free;
                          ShowMessage('Le script est endommagé.');
                          Ferme(tbCharger);
                          exit;
                    end;
                end;
            end;
            AScript.Destination := poGlobal;
            edtNature       := FindField('CIS_ClePar').asstring;
            TranfertT       := FindField('CIS_Table2').asstring;
            Logiciel        := FindField('CIS_Table1').asstring;
            AStreamTable.Free;
            Stream1.Free;

            AScript.FComment := tbCharger.FindField('CIS_COMMENT').AsString;
            AScript.NiveauAcces := FindField('CIS_MODIFIABLE').AsInteger;
            Close;
            Ferme(tbCharger);
    end;
	Result := Ascript;
end;
{$ENDIF}

function TExtPConcept.InitialiseScript : Boolean;
var
	N, index                     : Integer;
	AChamp                       : TChamp;
{$IFNDEF CISXPGI}
	FTable                       : TTable;
{$ENDIF}
        TN,TNV                       : TTreeNode;
        ANode                        : TNoeud;
begin
        if (Comment = taCreat) then begin Result := FALSE; exit; end;
        TNV := nil;
{$IFNDEF CISXPGI}
{$IFNDEF  DBXPRESS}
        FTable := TTable.Create(Application);
        FTable.DatabaseName :=  DMImport.DBGlobal.DatabaseName ;
        FTable.TableType := ttParadox;
        FTable.TableName := Fscript.options.asciiFilename;
{$ENDIF}
{$ENDIF}
        memCondition.Text := FScript.Options.Condition;
        FileName.Text := FScript.Options.FileName;
        edcomment.Text := FScript.FComment;
        if FScript.Options.TypeCar =  tcOEM then
           cbTypeCar.ItemIndex := 0
        else
           cbTypeCar.ItemIndex := 1;
        if FScript.Fields.Count = 0 then begin Result := FALSE; exit; end;
        ANode := nil;
        if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
        begin
           TN := tvChamp.Items.AddObject(nil, Codeimport, ANode);
           TNV := tvchampinvisible.Items.AddObject(nil, Codeimport, ANode)
        end
        else
           TN := tvChamp.Items.AddObject(nil, Codeimport, ANode);
        Modeparam.Itemindex := Fscript.Modeparam;
        InitMove(FScript.Champ.Count,'') ;
        for N:=0 to FScript.Champ.Count-1 do
        begin
                 AChamp := FScript.Champs[N];
                 MoveCur(FALSE);
                 if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
                 begin
                    renseignetree(Achamp, TNV, tvchampinvisible) ;
                    if (AChamp.Sel = 16) and (AChamp.Cache) then
                       renseignetree(Achamp, TN, tvchamp);
                 end
                 else
                    renseignetree(Achamp, TN, tvchamp);
                 Radiotable.ItemIndex := FScript.FileDest;
                 FicheSortie.Text     := FScript.Options.ASCIIFileName;
                 RTypefichier.ItemIndex := FScript.ASCIIMODE;
                 EdShellExecute.Text := FScript.Shellexec;
		            Table.Text := FScript.ScriptSuivant;

                 if assigned (FScript.Options) and (FScript.Options.SepField <> '') then
                 begin
                      for index :=0 to cbDelimChamp.Items.count-1 do
                      if  FScript.Options.SepField = CarSepField[index] then
                      begin cbDelimChamp.ItemIndex := index; break; end;
                 end;
        end;
{$IFNDEF CISXPGI}
{$IFNDEF  DBXPRESS}
        FTable.Free;
{$ENDIF}
{$ENDIF}
        FiniMove ;
        Result := TRUE;
end;

procedure TExtPConcept.ChargementTree(param : Boolean);
var
    Achamp                 : TChamp;
    N                      : Integer;
    TN,TNV                 : TTreeNode;
    ANode                  : TNoeud;
    Fige                   : Boolean;
    NT                     : string;
    procedure ChargementDictionnaire;
    var
    Val                 : string;
    i                   : integer;
    begin
        Anode := nil;
        Modeparam.Itemindex := Fscript.Modeparam;
        if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
        begin
           TN := tvChamp.Items.AddObject(nil, Codeimport, ANode);
           TNV := tvchampinvisible.Items.AddObject(nil, Codeimport, ANode)
        end
        else
           TN := tvChamp.Items.AddObject(nil, Codeimport, ANode);

        if TPControle.TOBParam = nil then begin PGIInfo ('Le Dictionnaire est Vide ');  exit; end;
        InitMove(TPControle.TOBParam.detail.count,'') ;
        for i := 0 to  TPControle.TOBParam.detail.count-1 do
        begin
             MoveCur(FALSE);
             with dmImport.GzImpPar do
             begin
               NT := (TPControle.TOBParam.detail[i].GetValue('TableName')) ;
               if (TPControle.TOBParam.detail[i].GetValue('Domaine') <> Domainepar) and
                  (Domainepar <>'X') then continue;
               if (Domainepar = 'G') then // Gescom
               begin
                   if  (NT <> 'Articles') and (NT <> 'Adresses') and (NT <> 'Catalogue') and
                   (NT <> 'Commercial') and (NT <> 'Contacts') and (NT <> 'Depots') and (NT <> 'MPaiement') and
                   (NT <> 'MReglt') and (NT <> 'Nomenclature') and (NT <> 'NomenLignes') and
                   (NT <> 'RIB') and (NT <> 'TablettesYX') and (NT <> 'TablettesCC') and
                   (NT <> 'Tarifs') and (NT <> 'TiersCP') and (NT <> 'TiersGC') and
                   (NT <> 'ComTiers') and (NT <> 'ExcepTiers') then continue;
               end;
               if (Domainepar = 'O') then // pour conso
               begin
                    if (TypeT = 'BAL') and (NT <> 'Obs') then continue;
                    if (TypeT = 'JRL') and (NT <> 'Oig') then continue;
                    if (TypeT = 'GRP') and (NT <> 'Oig') then continue;
               end;

               if (Domainepar= 'C') and ((NT= 'Entetepgi') or (*(NT= 'RIB') or*) (NT= 'Modepaiement') or
               (NT= 'Condregle') or (NT= 'Devise') or (NT= 'Regimetva') or (NT= 'Tauxtva') or (NT= 'Sectanalytiq'))
               then continue;
               if (not GestionAuxi) and (NT= 'Tiers') then continue;
               if (NbreExercice = 2) and (NT= 'Exercice2') Then continue;
               if (NbreExercice = 1) and ((NT= 'Exercice1') or (NT= 'Exercice2')) Then continue;
               if (Domainepar= 'M') and ((NT= 'Cess') or (NT= 'Leas') ) then continue;
               if (Domainepar= 'S') and (not Statistique) and  (NT= 'Arstat') then continue;
               if (Domainepar= 'C') and (not Analytique) and  (NT= 'Analytique') then continue;
               AChamp      := TChamp.Create(FScript.Champ);
               if NT <> '' then
                  AChamp.Name := NT+'_'+TPControle.TOBParam.detail[i].GetValue('Nom')
               else
                  AChamp.Name := TPControle.TOBParam.detail[i].GetValue('Nom');
               if (NT = 'Entete')  and
               (TPControle.TOBParam.detail[i].GetValue('Nom') = 'type') then
                    Val := TypeT
               else
                   Val  := TPControle.TOBParam.detail[i].GetValue('Contenu');
               if Val <> '' then
               begin
                  AChamp.TypeInfo := tiConstante;
                  AChamp.Constante.Valeur := Val;
                  AChamp.Constante.Pas := 0;
                  AChamp.Constante.Longueur := Length (Val);
               end;
               AChamp.Lon := TPControle.TOBParam.detail[i].GetValue('Longueur');
               AChamp.Siz := TPControle.TOBParam.detail[i].GetValue('Longueur');

               AChamp.Nbdecimal := TPControle.TOBParam.detail[i].GetValue('Nbdecimal');
               if TPControle.TOBParam.detail[i].GetValue('Obligatoire') = 'O' then
                  AChamp.Sel := 16
               else
                  AChamp.Sel := 20;

               if TPControle.TOBParam.detail[i].GetValue('calcul') <> '' then
               begin
                    if AChamp.Calcul = nil then
		       AChamp.Calcul := TCalcul.Create;
                    AChamp.Calcul.Formule := TPControle.TOBParam.detail[i].GetValue('calcul');
                    AChamp.TypeInfo := ticalcul;
               end;
{$IFNDEF CISXPGI}
               Fige := (TPControle.TOBParam.detail[i].GetValue('Fige')=1);
{$ELSE}
               Fige := FALSE;
{$ENDIF}
               AChamp.Cache :=  ((AChamp.Sel = 16) and (not Fige));

               if ((AChamp.sel and 4) <> 0)
               or ((Domainepar = 'O') and (AChamp.TypeInfo <> tiConstante) and (not Fige)) then // ajout me pour champ non sélectionné
               begin
                  AChamp.Deb := 0;
                  AChamp.Lon := 0;
               end;

               if (TPControle.TOBParam.detail[i].GetValue('TypeAlpha')='N') then
                  AChamp.Typ := 1;
               if (TPControle.TOBParam.detail[i].GetValue('TypeAlpha')='D') then
               begin
                  AChamp.Typ := 2;
                  AChamp.FormatDate := 3;
                  AChamp.FormatDateSortie := 3;
                  cbFormatDatesortie.Itemindex := 3; //JJMMAAAA
                  if Domainepar = 'S' then cbFormatDatesortie.Itemindex := 5; //JJMMAA

                  // cas imoII AAAAMMJJ
                  if (Domainepar= 'M') then
                  begin
                       if (NT= 'Exercice') and
                          (TPControle.TOBParam.detail[i].GetValue('Nom')='datefin') then
                       begin
                          AChamp.FormatDate := 4;
                          AChamp.FormatDateSortie := 4;
                       end;
                  end;
                  // cas datetri AAAAMMJJ
                  if (Domainepar= 'C') and (Domainepar= 'E')then
                  begin
                       if (NT= 'Ecriture') and
                          (TPControle.TOBParam.detail[i].GetValue('Nom')='datemvttri') then
                       begin
                          AChamp.FormatDate := 3;
                          AChamp.FormatDateSortie := 4;
                       end;
                  end;
               end;
               if (Domainepar= 'M') then
               begin
                       if (TPControle.TOBParam.detail[i].GetValue('Nom')='nocpte') then
                       begin
                               AChamp.Compl    := TRUE;
                               AChamp.ComplCar := '0';
                               AChamp.ComplLgn := 10;
                       end;
               end;

               Achamp.LOrder := TPControle.TOBParam.detail[i].GetValue('Ordretri');
               Achamp.FFamilleCorr := TPControle.TOBParam.detail[i].GetValue('Famcorresp');

               if (Achamp.LOrder = '') and  (Domainepar= 'C') then // si dictionnaire n'est pas à jour
               begin
                    if (NT = 'Ecriture') and
                          (TPControle.TOBParam.detail[i].GetValue('Nom')='codejournal') then
                                         Achamp.LOrder := 'Ecriture_codejournal, Ecriture_datemouvement, Ecriture_numeropiece, Ecriture_piece';
                    if (NT = 'Analytique') and
                          (TPControle.TOBParam.detail[i].GetValue('Nom')='codejournal') then
                                         Achamp.LOrder := 'Analytique_codejournal, Analytique_datemouvement, Analytique_piece ';
               end;
               // occurence
               if (DomainePar='C') and (TPControle.TOBParam.detail[i].GetValue('Nom')= 'fixe')and
               ((NT= 'Entete')or (NT= 'Param1')or (NT= 'Param2')or (NT= 'Param3')or (NT= 'Param4') or (NT= 'Param5')or (NT= 'Etabliss')) then
               Achamp.bUnenreg := TRUE;
               if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
               begin
                    renseignetree(Achamp, TNV, tvchampinvisible) ;
                    if (AChamp.Sel = 16) and (not Fige) then
                       renseignetree(Achamp, TN, tvchamp);
               end
               else
                    renseignetree(Achamp, TN, tvchamp);
             end;
        end;
        FiniMove ;
    end;

begin

    if Param then
    begin
         if FScript = nil then
              Fscript := LoadScript (NomTable)
         else
         begin
              Fichesauve.text := FScript.Name;
              FileName.Text   := FScript.Options.FileName;
              Edcomment.Text  := FScript.FComment;
              FScript.ParName := FScript.Name;
              Fscript.options.asciiFilename := FScript.Name;
              Radiotable.ItemIndex := FScript.FileDest;
              FicheSortie.Text := FScript.Options.ASCIIFileName;
              EdShellExecute.Text := Fscript.Shellexec;
              RTypefichier.ItemIndex := FScript.ASCIIMODE;
              Table.Text := FScript.ScriptSuivant;
         end;
    end
    else
    begin
         Domainepar := RendDomaine(Domaine.Text);
         for N := 0 to lgcorr do
         begin
                   if Correspimp[N].domaine = Domainepar then
                   begin Codeimport := Correspimp[N].code; break; end;
         end;
         if Codeimport = '' then  Codeimport := 'TRA';
         if TPControle = nil then
         begin
              TPControle := TSVControle.create;
              TPControle.ChargeTobParam(Domainepar);
         end
         else
         begin
              BPays.itemindex  := BPays.Values.IndexOf(TPControle.LePays);
         end;
         if not InitialiseScript then
         begin
                    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
                       tvchampinvisible.Items.Clear;
                    tvChamp.Items.Clear;
                    ChargementDictionnaire;
                    tvChamp.Items.EndUpdate;
         end; // initialise
    end; // param
end;

procedure TExtPConcept.InsererText(S : String; M : TMemo);
begin
	Clipboard.Clear;
	Clipboard.AsText := S;
	THRichEditOLE(M).PasteFromClipboard;
end;

procedure TExtPConcept.SpeedButton3Click(Sender: TObject);
begin
     InsererText(' '+TSpeedButton(Sender).Caption, TMEMO(memCondition));
end;

procedure TExtPConcept.tvChampClick(Sender: TObject);
var
index,I   : integer;
Node,TN2  : TTreeNode;
begin
    Node := nil;
    if  tvChamp.Items.Count = 0 then
    begin
         PGiInfo ('Chargez le fichier à importer', 'Paramétrage');
         exit;
    end;
    if (tvChamp.Selected.AbsoluteIndex >= 1)  then
    begin
         tbDEFINITION.TabVisible   := TRUE;
         Tabparam.TabVisible       := FALSE;
         tsCHPCONDITION.TabVisible := FALSE;
         TabSortie.TabVisible      := FALSE;
         PageControl1.ActivePage   := tbDEFINITION;
         ShowInfoPanel (tvChamp.Selected);
         if (tvChamp.Selected.level = level) and (tvChamp.Selected.Parent <> nil)
         and (TvChamp.selected.HasChildren) then
         begin
           // c'est un Enregistrement
                 TabConditioncomp.TabVisible := TRUE;
                 TabCorresp.Tabvisible       := TRUE;
                 //if (ParametrageCegid) then
                 TabLien.Tabvisible          := TRUE;

                 // Affichage du chemin en bas a droite de la fenetre
                 lblEnregCourant.Caption := 'TRA';
                 lblChampCourant.Caption := TvChamp.Selected.Text;
                 btnArboresCheminCourant.Visible := true;

                 if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
                 begin
                      Name := tvchamp.selected.text;
		      TN2 := tvchampinvisible.Items.GetFirstNode;
                      while TN2 <> nil do
                      begin
                           if TN2.Text = Name then
                           begin
                                Node := TN2.getFirstChild;
                                break;
                           end;
                           TN2 := TN2.GetNext;
                      end;
                 end
                 else
                      Node := tvChamp.selected.getFirstChild;
                 if Node = nil then exit;
                 Memo3.Text := TchampNode(Node.Data).Fchamp.ConditionChamp;
                 bCondition.checked := TchampNode(Node.Data).Fchamp.bCondition;
                 EnregUnique.checked := TchampNode(Node.Data).Fchamp.bUnenreg;
                 ListeProfile.clear;
                 for I := 0 to TchampNode(Node.Data).Fchamp.ListProfile.count - 1 do
                     ListeProfile.Items.Add(TchampNode(Node.Data).Fchamp.ListProfile.Strings[I]);
                 LienInter.Text := TchampNode(Node.Data).Fchamp.LienInter;
                 ORDER.Text     := TchampNode(Node.Data).Fchamp.LOrder;
         end
         else
         begin
           // un Champ est sélectionné
                 // Affichage du chemin en bas a droite de la fenetre
                 lblEnregCourant.Caption := TvChamp.Selected.Parent.Text;
                 lblChampCourant.Caption := TvChamp.Selected.Text;
                 btnArboresCheminCourant.Visible := true;

                 // Eviter les erreurs de champs avec un nom trop long
                 ChampNom.MaxLength := 25 - ( length(TvChamp.Selected.Parent.Text) + 1);

                 TabConditioncomp.TabVisible := False;
                 TabCorresp.Tabvisible       := False;
                 TabLien.Tabvisible          := False;
         end;
    end
    else
    if tvChamp.Selected.AbsoluteIndex = 0 then
    begin
      // TRA sélectionné
         tbDEFINITION.TabVisible    := False;
         tsCONSTANTE.TabVisible     := False;
         tsCALCUL.TabVisible        := False;
         tsREFERENCE.TabVisible     := False;
         tsCOMPLEMENT.TabVisible    := False;
         Tabparam.TabVisible        := TRUE;

         // Affichage du chemin en bas a droite de la fenetre
         lblEnregCourant.Caption := TvChamp.Selected.Text;
         lblChampCourant.Caption := '';
         btnArboresCheminCourant.Visible := false;

         if (not ParametrageCegid) and (FScript.NiveauAcces = 0) then
         begin
            tsCHPCONDITION.TabVisible  := FALSE;
            TabSortie.TabVisible      := FALSE;
         end
         else
         begin
            tsCHPCONDITION.TabVisible  := TRUE;
            TabSortie.TabVisible       := TRUE;
         end;
         TabConditioncomp.TabVisible := FALSE;
         TabCorresp.Tabvisible       := False;
         PageControl1.ActivePage     := Tabparam;
    end;
    index := tvChamp.Selected.AbsoluteIndex;
    if index > 0 then
      ChampNom.Text := tvChamp.Items.Item[index].Text;
      //ChampNom.Text := tvChamp.Selected.Text;
end;

Function TExtPConcept.ControleValidite : Boolean;
var
N         : integer;
AChamp    : TChamp;
Contenu   : variant;
begin
    Result := TRUE;

	for N:=0 to FScript.Champ.Count-1 do
	begin
		AChamp := FScript.Champ[N];
                if AChamp.TypeInfo = tiConstante then
                begin
                      Contenu := FScript.Champ[N].Constante.Valeur;
                      if not TPControle.Conroleparam(Contenu, Achamp.Name) then
                      begin
                           PgiBox ('Contenu du champ '+ Achamp.Name+ ' est incorrect, la valeur : '+TPControle.Valdefaut, 'Conception');
                           Result := FALSE;
                      end;
                end;
	end;

end;

{$IFDEF CISXPGI}
{$IFDEF EAGLCLIENT}
procedure TExtPConcept.BValiderClick(Sender: TObject);
var
    AStream, AStreamTable    : TMemoryStream;
    tbSauver,TLien           : TQuery;
    St                       : String;
    tbcharger                : TQuery;
    ABlobFieldD              : TField;
    Stream1                  : TmemoryStream;
    TOBSauve, TS             : TOB;
    s                        : TStringStream ;
    ABlobField, ATblCorField : TField;
    ABlobMemo                : TField;
begin
    // si niveau d'acces CEGID
    if (not ParametrageCegid) and (FScript.NiveauAcces = 0) then
    begin
         if (NomTable = FicheSauve.Text) then
         begin
               PGIBox ('Vous ne pouvez pas modifier le Standard CEGID', 'Conception');
               exit;
         end
         else
         if EdNature.Text = 'STD CEGID ' then EdNature.Text := '';
    end;

    // contrôle des champs constants
    if not ControleValidite then
    begin
     if (PGIAsk('Incohérences dans le Script, voulez-vous néanmoins enregistrer le Script : ?',
          'Conception')<> mryes) then exit;
    end;

    RefConditionExit (Sender);
    if not Assigned(dmImport) then
       dmImport := TdmImport.Create(Application);
  tbSauver := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+FScript.Name+'" AND CIS_NATURE<>"X"', FALSE);

  TS :=TOB.Create('CPGZIMPREQ',Nil,-1) ;
  TS.SelectDB('', tbSauver);
  Ferme(tbSauver);
  //	with tbSauver do
	begin
        FScript.Name := Copy (Fichesauve.text,1, 12);
        FScript.Options.FileName :=  FileName.Text;
        if ParametrageCegid then
           FScript.NiveauAcces := 0
        else
           FScript.NiveauAcces := 1;
        FScript.FComment := Edcomment.Text;
        FScript.FNouveau := False;
        FScript.FileType := 1; // pour ascii à voir
        FScript.ParName := FScript.Name;
        Fscript.options.asciiFilename := FScript.Name;
        FScript.FileDest := Radiotable.ItemIndex;
        FScript.Options.ASCIIFileName := FicheSortie.Text;
        Fscript.Shellexec := EdShellExecute.Text;
        FScript.ASCIIMODE := RTypefichier.ItemIndex;
        FScript.ConstruitVariable(TRUE);
        FScript.Variable.Clear;
        AjouteVariableALias(FScript.Variable);

        if RTypefichier.ItemIndex = 1 then
        begin
            if cbDelimChamp.ItemIndex <> -1 then
                    FScript.Options.SepField := CarSepField[cbDelimChamp.ItemIndex]
            else
            if  cbDelimChamp.Text <> '' then
                    FScript.Options.SepField := cbDelimChamp.Text[1];
        end;
        AStream := TMemoryStream.Create;
        AStreamTable := TMemoryStream.Create;
        FScript.SaveTo(AStream, AStreamTable);
        s := TStringStream.Create('');
        AStream.Seek(0, soFromBeginning);
        s.CopyFrom(AStream, AStream.Size);
        S.Seek(0, soFromBeginning);
        TS.putValue ('CIS_PARAMETRES', s.DataString);
        s.Free;
        if not Endommage then
        begin
             if not DirectoryExists(CurrentDossier+'\sauvegarde') then
                CreateDir(CurrentDossier+'\sauvegarde');
            AStream.SaveToFile (CurrentDossier+'\sauvegarde\'+FScript.Name+'parametres.txt');
        end;

        TS.putValue ('CIS_CLE', FScript.Name);
        TS.putValue ('CIS_COMMENT', FScript.FComment);
        TS.putValue ('CIS_CLEPAR', edNature.Text);
        TS.putValue('CIS_DOMAINE', RendDomaine(Domaine.Text));
        TS.putValue('CIS_DATEMODIF', now);
        TS.putValue('CIS_Table2', Typetransfert.text);
        TS.putValue('CIS_Table1', Editeur.text);
        TS.putValue('CIS_Table3', BPAYS.Value);
        TS.putValue('CIS_Table4', '');
        TS.putValue('CIS_Table5', '');
        TS.putValue('CIS_Table6', '');
        TS.putValue('CIS_Table7', '');
        TS.putValue('CIS_Table8', '');
        TS.putValue('CIS_Table9', '');
        TS.putValue('CIS_NATURE', '-');
        TS.InsertOrUpdateDB(True);
        TS.free;
        // table de correspondance
        TLien := OpenSQl ('SELECT * FROM CPGZIMPCORRESP Where CIC_TABLEBLOB="CIS" and CIC_IDENTIFIANT="'+FScript.Name+'"', FALSE);
        TS :=TOB.Create('CPGZIMPCORRESP',Nil,-1) ;
        TS.SelectDB('', TLien);
        Ferme(TLien);
                s := TStringStream.Create('');
                AStreamTable.Seek(0, soFromBeginning);
                s.CopyFrom(AStreamTable, AStreamTable.Size);
                S.Seek(0, soFromBeginning);
                TS.PutValue('CIC_OBJET', s.DataString);
                s.Free;
                AStreamTable.Free;

                TS.PutValue('CIC_TABLEBLOB', 'CIS') ;
                TS.PutValue('CIC_IDENTIFIANT', FScript.Name) ;
                TS.PutValue('CIC_LIBELLE', FScript.FComment) ;
                TS.PutValue('CIC_RANGBLOB', 1);
                TS.PutValue('CIC_DATEMODIF', now);
                TS.PutValue('CIC_QUALIFIANTBLOB', 'DATA');
                TS.PutValue('CIC_EMPLOIBLOB', '');
                TS.PutValue('CIC_PRIVE', '-');
                TS.PutValue('CIC_DATECREATION', now) ;
                TS.PutValue('CIC_DATEBLOB', now) ;
                if FScript.NiveauAcces = 0 then
                begin
                   TS.PutValue('CIC_UTILISATEUR', 'CEG') ;
                   TS.PutValue('CIC_CREATEUR', 'CEG') ;
                end
                else
                begin
                   TS.PutValue('CIC_UTILISATEUR', 'STD') ;
                   TS.PutValue('CIC_CREATEUR', 'STD') ;
                end;
                TS.InsertOrUpdateDB(True);
                TS.free;
		            FModified := False;
        BOkValide := TRUE;
	end;

     if ((Fscript.PreTrt.count <> 0) and (NomTable <> FScript.Name))
     or (ChargeTree)then
     begin
         if not (ChargeTree) then
         begin
            FScriptDel := nil;
            FScriptDel := TScriptDelimite.Create(nil);
            tbcharger := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+FScript.Name+'" AND CIS_NATURE="X"', TRUE);

            with tbcharger do
            begin
                   if not tbcharger.EOF then
                   begin
                    ABlobFieldD := FindField('CIS_PARAMETRES');
                    Stream1 := TmemoryStream.create;
                    TBlobField(ABlobFieldD).SaveToStream (Stream1);
                    Stream1.Seek (0,0);
                    FScriptDel := LoadScriptFromStreamDelim(Stream1);
                   end;
               Ferme (tbcharger);
            end;
          end;
        if FScriptDel = nil then exit;
        tbSauver := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+FScript.Name+'" AND CIS_NATURE="X"', FALSE);
        TS :=TOB.Create('CPGZIMPREQ',Nil,-1) ;
        TS.SelectDB('', tbSauver);
        Ferme (tbSauver);
        FScriptDel.Name := FScript.Name;
        TS.PutValue ('CIS_CLE', FScriptDel.Name);
        TS.PutValue ('CIS_COMMENT', FScriptDel.Libelle);
        TS.PutValue ('CIS_CLEPAR', '');
        TS.PutValue ('CIS_DATEMODIF', now);
        TS.PutValue ('CIS_Table2', Typetransfert.text);
        TS.PutValue ('CIS_Table1', Editeur.text);
        TS.PutValue ('CIS_Table3', BPAYS.Value);
        TS.PutValue ('CIS_MODIFIABLE', FScript.NiveauAcces);
        TS.PutValue('CIS_DOMAINE', Domainepar);
        TS.PutValue('CIS_DATEMODIF', now);
        TS.PutValue('CIS_NATURE', 'X');
        AStream := TMemoryStream.Create;
        FScriptDel.SaveDelimTo(AStream);
        s := TStringStream.Create('');
        TBlobField(AStream).SaveToStream(s);
        TS.PutValue('CIS_PARAMETRES', s.DataString);
        s.Free;
        TS.InsertOrUpdateDB(True);
        TS.free;
	  end;
        FScriptDel.free;

end;
{$ELSE}
procedure TExtPConcept.BValiderClick(Sender: TObject);
var
    AStream, AStreamTable    : TStream;
    tbSauver,TLien           : TQuery;
    St                       : String;
    tbcharger                : TQuery;
    ABlobFieldD              : TField;
    Stream1                  : TmemoryStream;
{$IFDEF EAGLCLIENT}
    s                        : TStringStream ;
    ABlobField, ATblCorField : TField;
    ABlobMemo                : TField;
{$ELSE}
    ABlobField, ATblCorField : TBlobField;
    ABlobMemo                : TMemoField;
{$ENDIF}

begin
    // si niveau d'acces CEGID
    if (not ParametrageCegid) and (FScript.NiveauAcces = 0) then
    begin
         if (NomTable = FicheSauve.Text) then
         begin
               PGIBox ('Vous ne pouvez pas modifier le Standard CEGID', 'Conception');
               exit;
         end
         else
         if EdNature.Text = 'STD CEGID ' then EdNature.Text := '';
    end;

    // contrôle des champs constants
    if not ControleValidite then
    begin
     if (PGIAsk('Incohérences dans le Script, voulez-vous néanmoins enregistrer le Script : ?',
          'Conception')<> mryes) then exit;
    end;

    RefConditionExit (Sender);
    if not Assigned(dmImport) then
       dmImport := TdmImport.Create(Application);
    tbSauver := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+FScript.Name+'" AND CIS_NATURE<>"X"', FALSE);

	with tbSauver do
	begin
		Open;
		if not tbSauver.EOF then
                begin
                    if FindField ('CIS_CLEPAR').Asstring = 'SQL' then
                    begin
                     PGIBox('Au niveau Extraction SGBD, une requête '+ FScript.Name + ' existe déjà, changez le nom de votre script', 'Conception') ;
                     Ferme (tbSauver);
                     exit;
                    end;

                    St := Format('Le Script %s existe déjà, confirmez-vous sa modification  ?', [FScript.Name]);
                    if (PGIAsk(St, 'Conception')<> mryes) then begin exit;  Close; end;
                    Edit;
                end
		else Insert;
                if ModeVisu <> Fscript.Modeparam then
                begin
                 if (PGIAsk('Voulez-vous modifier le mode de paramétrage ?',
                      'Conception')= mryes) then
                      Fscript.Modeparam := ModeVisu;
                      ModalResult := mrOK;
                end;

        FScript.Name := Copy (Fichesauve.text,1, 12);
        FScript.Options.FileName :=  FileName.Text;
        if ParametrageCegid then
           FScript.NiveauAcces := 0
        else
           FScript.NiveauAcces := 1;
        FScript.FComment := Edcomment.Text;
        FScript.FNouveau := False;
        FScript.FileType := 1; // pour ascii à voir
        FScript.ParName := FScript.Name;
        Fscript.options.asciiFilename := FScript.Name;
        FScript.FileDest := Radiotable.ItemIndex;
        FScript.Options.ASCIIFileName := FicheSortie.Text;
        Fscript.Shellexec := EdShellExecute.Text;
        FScript.ASCIIMODE := RTypefichier.ItemIndex;
        FScript.ConstruitVariable(TRUE);
        FScript.Variable.Clear;
        AjouteVariableALias(FScript.Variable);

        if RTypefichier.ItemIndex = 1 then
        begin
            if cbDelimChamp.ItemIndex <> -1 then
                    FScript.Options.SepField := CarSepField[cbDelimChamp.ItemIndex]
            else
            if  cbDelimChamp.Text <> '' then
                    FScript.Options.SepField := cbDelimChamp.Text[1];
        end;
                ABlobField := FindField('CIS_PARAMETRES') as TBlobField;
                ABlobMemo :=  FindField('CIS_PARAMETRES') as TMemoField;
                if not Endommage then
                begin
                     if not DirectoryExists(CurrentDossier+'\sauvegarde') then
                        CreateDir(CurrentDossier+'\sauvegarde');
                     TMemoField(ABlobMemo).SaveToFile (CurrentDossier+'\sauvegarde\'+FScript.Name+'parametres.txt');
                end;
		            FindField('CIS_CLE').AsString := FScript.Name;
		            FindField('CIS_COMMENT').AsString := FScript.FComment;
		            FindField('CIS_CLEPAR').AsString := edNature.Text;
                FindField('CIS_MODIFIABLE').asinteger := FScript.NiveauAcces;
                FindField('CIS_DOMAINE').asstring := RendDomaine(Domaine.Text);
                FindField('CIS_DATEMODIF').asdatetime := now;
                FindField('CIS_Table2').Asstring := Typetransfert.text;
                FindField('CIS_Table1').Asstring := Editeur.text;
                FindField('CIS_Table3').asstring := BPAYS.Value;
                FindField('CIS_Table4').asstring := '';
                FindField('CIS_Table5').asstring := '';
                FindField('CIS_Table6').asstring := '';
                FindField('CIS_Table7').asstring := '';
                FindField('CIS_Table8').asstring := '';
                FindField('CIS_Table9').asstring := '';
                FindField('CIS_NATURE').asstring := '-';

		//	AStream := tbSauver.CreateBlobStream(ABlobField, bmWrite);
                // table de correspondance

                TLien := OpenSQl ('SELECT * FROM CPGZIMPCORRESP Where CIC_TABLEBLOB="CIS" and CIC_IDENTIFIANT="'+FScript.Name+'"', FALSE);
                ATblCorField := TLien.FindField('CIC_OBJET') as TBlobField;
                if not TLien.EOF then TLien.Edit else TLien.Insert ;
                AStream := TBlobStream.Create(ABlobField, bmWrite);
		            AStreamTable := TBlobStream.Create(ATblCorField, bmWrite);
		            FScript.SaveTo(AStream, AStreamTable);
                AStreamTable.Free;
		            AStream.Free;

                TLien.FindField('CIC_TABLEBLOB').AsString   := 'CIS' ;
                TLien.FindField('CIC_IDENTIFIANT').AsString := FScript.Name ;
                TLien.FindField('CIC_LIBELLE').AsString     := FScript.FComment ;
                TLien.FindField('CIC_RANGBLOB').AsInteger   := 1;
                TLien.FindField('CIC_DATEMODIF').AsDateTime    := now ;
                TLien.FindField('CIC_QUALIFIANTBLOB').asstring := 'DATA';
                TLien.FindField('CIC_EMPLOIBLOB').asstring := '';
                TLien.FindField('CIC_PRIVE').asstring := '-';
                TLien.FindField('CIC_DATECREATION').AsDateTime    := now ;
                TLien.FindField('CIC_DATEBLOB').AsDateTime    := now ;
                if FScript.NiveauAcces = 0 then
                begin
                   TLien.FindField('CIC_UTILISATEUR').Asstring    := 'CEG' ;
                   TLien.FindField('CIC_CREATEUR').Asstring    := 'CEG' ;
                end
                else
                begin
                   TLien.FindField('CIC_UTILISATEUR').Asstring    := 'STD' ;
                   TLien.FindField('CIC_CREATEUR').Asstring    := 'STD' ;
                end;
                TLien.Post ;
                Ferme (TLien);

        // pour sauvegarder dans un fichier ascii ATblCorField.savetofile (FScript.Name+'.txt');
		            Post;

                Ferme (tbSauver);
		            FModified := False;
        BOkValide := TRUE;
	end;



     if ((Fscript.PreTrt.count <> 0) and (NomTable <> FScript.Name))
     or (ChargeTree)then
     begin
         if not (ChargeTree) then
         begin
            FScriptDel := nil;
            FScriptDel := TScriptDelimite.Create(nil);
            tbcharger := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+FScript.Name+'" AND CIS_NATURE="X"', TRUE);

            with tbcharger do
            begin
                Open;
                   if not tbcharger.EOF then
                   begin
                    ABlobFieldD := FindField('CIS_PARAMETRES');
                    Stream1 := TmemoryStream.create;
                    TBlobField(ABlobFieldD).SaveToStream (Stream1);
                    Stream1.Seek (0,0);
                    FScriptDel := LoadScriptFromStreamDelim(Stream1);
                   end;
               Ferme (tbcharger);
            end;
          end;
        if FScriptDel = nil then exit;
        tbSauver := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+FScript.Name+'" AND CIS_NATURE="X"', FALSE);
        with tbSauver do
        begin
		            Open;
                if tbsauver.EOF then
                begin
                      FScriptDel.Name    := FScript.Name;
                      FScriptDel.Libelle := FScript.FComment;

                      Insert;
                      ABlobField := FindField('CIS_PARAMETRES') as TBlobField;
		                  FindField('CIS_CLE').AsString             := FScriptDel.Name;
		                  FindField('CIS_COMMENT').AsString         := FScriptDel.Libelle;
		                  FindField('CIS_CLEPAR').AsString          := '';
                      FindField('CIS_DATEMODIF').asdatetime     := now;
                      FindField ('CIS_Table2').Asstring         := Typetransfert.text;
                      FindField ('CIS_Table1').Asstring         := Editeur.text;
                      FindField('CIS_Table3').asstring          := BPAYS.Value;
                      FindField('CIS_MODIFIABLE').asinteger     :=  FScript.NiveauAcces;
                      FindField('CIS_DOMAINE').asstring         := Domainepar;
                      FindField('CIS_DATEMODIF').asdatetime   := now;
                      FindField('CIS_NATURE').asstring          := 'X';
{$IFNDEF EAGLCLIENT}
		                  AStream := TBlobStream.Create(ABlobField, bmWrite);
{$ELSE}
		                  AStream := TMemoryStream.Create;
{$ENDIF}
		                  FScriptDel.SaveDelimTo(AStream);
{$IFDEF EAGLCLIENT}
                      s := TStringStream.Create(FindField('CIS_PARAMETRES').asstring);
                      TBlobField(AStream).SaveToStream(s);
                      FindField('CIS_PARAMETRES').asstring := s.DataString;
                      s.Free;
{$ENDIF}
                      Post;
                end;
               Ferme (tbSauver);
	end;
        FScriptDel.free;
     end;

end;
{$ENDIF EAGLCLIENT}
{$ELSE}

procedure TExtPConcept.BValiderClick(Sender: TObject);
var
    AStream, AStreamTable    : TStream;
    ABlobField, ATblCorField : TBlobField;
    ABlobMemo                : TMemoField;
    tbSauver,TbDelim         : TADOTable;
    tbcharger                : TADOTable;
    St                       : String;
    ABlobFieldD              : TField;
    Stream1                  : TmemoryStream;

begin
    // si niveau d'acces CEGID
    if (not ParametrageCegid) and (FScript.NiveauAcces = 0) then
    begin
         if (NomTable = FicheSauve.Text) then
         begin
               PGIBox ('Vous ne pouvez pas modifier le Standard CEGID', 'Conception');
               exit;
         end
         else
         if EdNature.Text = 'STD CEGID ' then EdNature.Text := '';
    end;

    // contrôle des champs constants
    if not ControleValidite then
    begin
     if (PGIAsk('Incohérences dans le Script, voulez-vous néanmoins enregistrer le Script : ?',
          'Conception')<> mryes) then exit;
    end;

    RefConditionExit (Sender);
    if not Assigned(dmImport) then
       dmImport := TdmImport.Create(Application);
    tbSauver := DMImport.GzImpReq;

	with tbSauver do
	begin
                if Active then
                            Close;
                Open;
                if Locate('CLE', FScript.Name, [loCaseInsensitive]) then
                begin
                    if FindField ('CLEPAR').asstring  = 'SQL' then
                    begin
                     PGIBox('Au niveau Extraction SGBD, une requête '+ FScript.Name + ' existe déjà, changez le nom de votre script', 'Conception') ;
                     Close;
                     exit;
                    end;

                    St := Format('Le Script %s existe déjà, confirmez-vous sa modification  ?', [FScript.Name]);
                    if (PGIAsk(St, 'Conception')<> mryes) then begin exit;  Close; end;
                    Edit;
                end
		            else Insert;
                if ModeVisu <> Fscript.Modeparam then
                begin
                 if (PGIAsk('Voulez-vous modifier le mode de paramétrage ?',
                      'Conception')= mryes) then
                      Fscript.Modeparam := ModeVisu;
                      ModalResult := mrOK;
                end;

                FScript.Name := Copy (Fichesauve.text,1, 12);
                FScript.Options.FileName :=  FileName.Text;
                if ParametrageCegid then
                   FScript.NiveauAcces := 0
                else
                   FScript.NiveauAcces := 1;
                FScript.FComment := Edcomment.Text;
                FScript.FNouveau := False;
                FScript.FileType := 1; // pour ascii à voir
                FScript.ParName := FScript.Name;
                Fscript.options.asciiFilename := FScript.Name;
                FScript.FileDest := Radiotable.ItemIndex;
                FScript.Options.ASCIIFileName := FicheSortie.Text;
                Fscript.Shellexec := EdShellExecute.Text;
                FScript.ASCIIMODE := RTypefichier.ItemIndex;
                FScript.ConstruitVariable(TRUE);
                FScript.Variable.Clear;
                AjouteVariableALias(FScript.Variable);

                if RTypefichier.ItemIndex = 1 then
                begin
                    if cbDelimChamp.ItemIndex <> -1 then
                            FScript.Options.SepField := CarSepField[cbDelimChamp.ItemIndex]
                    else
                    if  cbDelimChamp.Text <> '' then
                            FScript.Options.SepField := cbDelimChamp.Text[1];
                end;
                ABlobField := FieldByName('PARAMETRES') as TBlobField;
                ABlobMemo :=  FieldByName('PARAMETRES') as TMemoField;
                if not Endommage then
                begin
                     if not DirectoryExists(CurrentDossier+'\sauvegarde') then
                        CreateDir(CurrentDossier+'\sauvegarde');
                     TMemoField(ABlobMemo).SaveToFile (CurrentDossier+'\sauvegarde\'+FScript.Name+'parametres.txt');
                end;
                ATblCorField := FieldByName('TBLCOR') as TBlobField;
                FieldByName('CLE').AsString := FScript.Name;
                FieldByName('COMMENT').AsString := FScript.FComment;
                FieldByName('CLEPAR').AsString := edNature.Text;
                FieldByName('MODIFIABLE').AsInteger := FScript.NiveauAcces;
                FieldByName('DOMAINE').asstring := RendDomaine(Domaine.Text);
                FieldByName('DATEDEMODIF').asdatetime := now;
                FieldByName ('Table0').Asstring := Typetransfert.text;
                FieldByName ('Table1').Asstring := Editeur.text;
                FieldByName('Table2').asstring := '';
                FieldByName('Table3').asstring := '';
                FieldByName('Table4').asstring := '';
                FieldByName('Table5').asstring := '';
                FieldByName('Table6').asstring := '';
                FieldByName('Table7').asstring := '';
                FieldByName('Table8').asstring := '';
                FieldByName('Table9').asstring := '';
                AStream := TBlobStream.Create(ABlobField, bmWrite);
                //	AStream := tbSauver.CreateBlobStream(ABlobField, bmWrite);
                AStreamTable := TBlobStream.Create(ATblCorField, bmWrite);
                FScript.SaveTo(AStream, AStreamTable);
                AStreamTable.Free;
                AStream.Free;
                    // pour sauvegarder dans un fichier ascii ATblCorField.savetofile (FScript.Name+'.txt');
                Post;
                Close;
                FModified := False;
                BOkValide := TRUE;
	end;

     if ((Fscript.PreTrt.count <> 0) and (NomTable <> FScript.Name))
     or (ChargeTree)then
     begin
         if not (ChargeTree) then
         begin
            FScriptDel := nil;
            FScriptDel := TScriptDelimite.Create(nil);
            tbcharger := DMImport.GzImpDelim;
            with tbcharger do
            begin
                Open;
                if Locate('CLE', NomTable, [loCaseInsensitive]) then
                   begin
                    ABlobFieldD := FieldByName('PARAMETRES');
                    Stream1 := TmemoryStream.create;
                    TBlobField(ABlobFieldD).SaveToStream (Stream1);
                    Stream1.Seek (0,0);
                    FScriptDel := LoadScriptFromStreamDelim(Stream1);
                    Stream1.free;
                   end;
                close;
            end;
          end;
        if FScriptDel = nil then exit;
        TbDelim := DMImport.GzImpDelim;
        TbDelim.Connection := DMImport.DBGLOBAL;
        with TbDelim do
        begin
		            Open;
                if not Locate('CLE', FScript.Name, [loCaseInsensitive]) then
                begin
                      FScriptDel.Name := FScript.Name;
                      Insert;
                      ABlobField := FieldByName('PARAMETRES') as TBlobField;
                      FieldByName('CLE').AsString             := FScriptDel.Name;
                      FieldByName('COMMENT').AsString         := FScriptDel.Libelle;
                      FieldByName('CLEPAR').AsString          := '';
                      FieldByName('MODIFIABLE').AsInteger     := FScript.NiveauAcces;
                      FieldByName('DOMAINE').asstring         := Domainepar;
                      FieldByName('DATEDEMODIF').asdatetime   := now;
                      AStream := TBlobStream.Create(ABlobField, bmWrite);
                      FScriptDel.SaveDelimTo(AStream);
                      AStream.free;
                      Post;
                end;
                Close;
	end;

        FScriptDel.free;
     end;

end;
{$ENDIF}

procedure TExtPConcept.FichesauveChange(Sender: TObject);
begin
  SaveFiche    := FScript.Name;
  FScript.Name := Fichesauve.Text;

  if Fichesauve.text <> '' then
   caption := 'Paramétrage - ' + Fichesauve.text
  else caption := 'Paramétrage - Script sans nom';

  if FRefresh or not Assigned(curChamp) then exit;
  FModified := True;
  curChamp.Name := ChampNom.Text;
  if tvChamp.Selected <> nil then tvChamp.Selected.Text := curchamp.Name;
end;

procedure TExtPConcept.tvChampChange(Sender: TObject; Node: TTreeNode);
begin
   if Node.data <> nil then
   begin
		curChamp := TChampNode(Node.Data).FChamp;
	    if curChamp.Reference = nil then
		   curChamp.Reference := TReference.Create;
            if curChamp.Constante = nil then
		   curChamp.Constante := TConstante.Create;
            if curChamp.Calcul = nil then
		   curChamp.Calcul := TCalcul.Create;
   end;
    if Node.Level = level then
    begin
        tvChamp.OnDragOver := tvChampDragOver;
        tvChamp.OnDragDrop := tvChampDragDrop;
    end else begin
        tvChamp.OnDragOver := tvPartieDragOver;
        tvChamp.OnDragDrop := tvPartieDragDrop;
    end;

    ShowInfoPanel(Node);
end;

procedure TExtPConcept.tvChampEditing(Sender: TObject; Node: TTreeNode;
  var AllowEdit: Boolean);
begin
	AllowEdit := False;
end;

procedure TExtPConcept.tvChampEnter(Sender: TObject);
begin
FocusCtrl := TComponent(Sender);
end;

procedure TExtPConcept.tvChampKeyPress(Sender: TObject; var Key: Char);
begin
	if Key = ' ' then begin
		tvChampDblClick(Sender);
		Key := #0;
	end;
end;

(* ------------------------------------------------------------------ *)

constructor TNoeud.Create(AOwner:TForm);
begin
	FOwner := TExtPConcept(AOwner);
end;

destructor TNoeud.Destroy;
begin
  inherited Destroy;
end;

procedure TExtPConcept.ChargementInfo;
begin
        if (FileName.Text <> '') or (not Fmodified) then
        ChargementTree(FALSE);
end;

procedure TExtPConcept.ChargementClick(Sender: TObject);
var Stream1            : TmemoryStream;
begin
  if (Comment= tacreat)  then
  begin
   if FileExists(CurrentDossier+'\sauvegarde\'+FicheSauve.Text+'parametres.txt') then
   begin
        Stream1 := TmemoryStream.create;
        Stream1.LoadFromFile (CurrentDossier+'\sauvegarde\'+FicheSauve.Text+'parametres.txt');
        Stream1.Seek (0,0);
        Fscript := LoadScriptFromStream(Stream1, nil);
        Stream1.Free;
        Endommage := TRUE;
        Comment := TaModif;
        FScript.InitTableAndFields;
        ChargementTree(FALSE);
    end;
   if FileExists(CurrentDossier+'\sauvegarde\'+FicheSauve.Text+'Delimite.txt') then
   begin
                      NomTable := FScript.Name;
                      Stream1 := TmemoryStream.create;
                      Stream1.LoadFromFile (CurrentDossier+'\sauvegarde\'+FicheSauve.Text+'Delimite.txt');
                      Stream1.Seek (0,0);
                      FScriptDel := LoadScriptFromStreamDelim(Stream1);
   end;
   ChargeTree := TRUE;
  end;
end;

procedure TExtPConcept.renseignetree (Achamp : TChamp; TT : TTreeNode; var vChamp : TTreeview);
var
   TN, TN1, TN2     : TTreeNode;
   idx, iPoint,M, I : Integer;
   ANode            : TNoeud;
   ChMaitre, ChEnfant : string;
   allsel, anysel   : boolean;
	function idxof(str :  string):integer;
	var id : integer;
	begin
		result := -1;
		for id := 0 to vChamp.Items.Count -1 do
		begin
			 if vChamp.Items.Item[id].Text = str then
			 begin
				  result := id;
				  break;
			 end;
		end;
    end;
begin
        TN2 := nil; TN := nil;
        ANode := TChampNode.Create(Self);
        TChampNode(ANode).FChamp := AChamp;
{$IFNDEF CISXPGI}
        if Domaine.Text <> 'Reprise' then
           iPoint := pos('_', Achamp.Name)
        else
           iPoint := 0;
{$ELSE}
           iPoint := pos('_', Achamp.Name);
{$ENDIF}
        if (TChampNode(ANode).FChamp.Siz = 0) and (TChampNode(ANode).FChamp.Lon <> 0)  then
           TChampNode(ANode).FChamp.Siz := TChampNode(ANode).FChamp.Lon;

        if iPoint <> 0 then // il y a un champ à éclater
        begin
            chMaitre := copy(Achamp.name, 1, iPoint-1);
            chEnfant := copy(Achamp.name, iPoint+1, length(achamp.Name));
            if (PrevChMaitre <> ChMaitre)  or (Fscript.Modeparam = 1) then // pour réduire le temps de chargement de la forme
            begin // il faut rechercher le père
(*                  if (PrevChMaitre = 'Ecriture') and (Fscript.Modeparam <> 1) then  // Fiche 10102
                  begin
                        if (FScript.Champ.IndexOf ('Ecriture_datemvttri') = -1) then
                        begin
                              curChamp := TChamp.Create(FScript.Champ);
                              curchamp.Name := 'Ecriture_datemvttri';
                              curChamp.Lon := 8;  curChamp.Siz := 8;
                              curchamp.deb := 1;
                              curChamp.Typ := 2;
	                      if curChamp.Calcul = nil then curChamp.Calcul := TCalcul.Create;
                              curChamp.Calcul.Formule := 'Ecriture_datemouvement';
                              curChamp.TypeInfo := ticalcul;
                              curChamp.FormatDate := 3;
                              curChamp.FormatDateSortie := 4;
                              renseignetree(curchamp, TN, tvchamp);
                              idx := FScript.Champ.IndexOf ('Ecriture_codejournal');
                              FScript.Champ[idx].LOrder := FScript.Champ[idx].LOrder + ',Ecriture_datemvttri';
                        end;
                  end;
*)
                  idx := idxof(ChMaitre); // rechercher le champ père
                  if  idx < 0 then
                  TN2 := vChamp.Items.AddChildObject(TT, ChMaitre, nil) // on ne l'a pas trouvé donc on le crée
                  else TN2 := vChamp.Items.item[idx]; // sinon, on l'a trouvé
                  TN := vChamp.Items.AddChildObject(TN2, ChEnfant, ANode);
                  PrevChMaitre := ChMaitre;
                  PrevNodeMaitre := TN2;
            end
            else // le père est le même que le précédent
                  TN := vChamp.Items.AddChildObject(PrevNodeMaitre, ChEnfant, ANode);
        end
        else
		TN := vChamp.Items.AddChildObject(TT, AChamp.Name, ANode);


            case (AChamp.Sel and 12) of
                  0: begin
                      TN.ImageIndex := 0;   // selectionne
                      TN.SelectedIndex := 0;
                      end;
                  4: begin
                      TN.ImageIndex := 2;    // non selectionne
                      TN.SelectedIndex := 2;
                      end;
                  8: begin
                      TN.ImageIndex := 3;     // selectionne
                      TN.SelectedIndex := 3;
                      end;
                  12: begin
                      TN.ImageIndex := 4;    // non selectionne
                      TN.SelectedIndex := 4;
                      end;
            end;
            if TN.absoluteIndex = 0 then
            begin
                      TN.ImageIndex := 9;
                      TN.SelectedIndex := 9;
            end;
            if (TN2 <> nil) then
            begin
                      TN1 := TN2.GetFirstChild;

                      AllSel := True;
                      anySel := True;
                      for I := 0 to TN2.Count-1 do
                      begin
                         if TN1 <> nil then
                         begin
                           if ((TchampNode(TN1.Data).Fchamp.sel and 4) = 0) then
                           begin
                              TN2.ImageIndex := 6;
                              TN2.SelectedIndex := 6;
                              AnySel := false;
                           end
                           else
                              AllSel := False;
                           TN1 := TN2.GetNextChild(TN1);
                         end;
                      end;
                      if AllSel then
                      begin
                           TN2.ImageIndex := 5;
                           TN2.SelectedIndex := 5;
                      end
                      else if AnySel then
                      begin
                           TN2.ImageIndex := 7;
                           TN2.SelectedIndex := 7;
                      end;
            end;
            if AChamp.TypeInfo = tiConcat then
            begin
                      for M:=0 to AChamp.Concat.Count-1 do
                      begin
                          ANode := TChampNode.Create(Self);
                          TChampNode(ANode).FChamp := TChamp(AChamp.Concat[M]);
                          TN1 := vChamp.Items.AddChildObject(TN, TChampNode(ANode).FChamp.Name, ANode);
                          case (AChamp.Concat[M].Sel and 12) of
                          0: begin
                              TN1.ImageIndex := 0;
                              TN1.SelectedIndex := 0;
                              end;
                          4: begin
                              TN1.ImageIndex := 2;
                              TN1.SelectedIndex := 2;
                              end;
                          8: begin
                              TN1.ImageIndex := 3;
                              TN1.SelectedIndex := 3;
                              end;
                          12: begin
                              TN1.ImageIndex := 4;
                              TN1.SelectedIndex := 4;
                              end;
                          end;
                      end;
            end;
            if AChamp.Typ = 1 then lbChampNum.Items.Add(AChamp.Name);
end;


procedure TExtPConcept.tvChampDragDrop(Sender, Source: TObject; X,
	Y: Integer);
var
	TN, TN1 : TTreeNode;
        Idx,Idx2: integer;
        id      : integer;
begin
    FModified := True;
	with tvChamp do
	begin
		Items.BeginUpdate;
		TN1 := GetNodeAt(X, Y);
		TN := Selected;
		if TN1 = nil then
		begin
                      if TN.Data <> nil  then
                      begin
                          FScript.Champ.Move(TChampNode(TN.Data).FChamp.Index, FScript.Champ.Count-1);
                          TN.MoveTo(TN1, naAdd)
                      end;
                end
                else if (TN <> TN1) and (TN.Index <> TN1.Index-1) then
                begin
                              if (TN.Index > TN1.Index) and (tvchamp.selected.Data <> nil) and (TN1.Data <> nil) then
                                      FScript.Champ.Move(TChampNode(TN.Data).FChamp.Index, TChampNode(TN1.Data).FChamp.Index)
                              else if (tvchamp.selected.Data <> nil) and (TN1.data <> nil) then
                              FScript.Champ.Move(TChampNode(TN.Data).FChamp.Index, TChampNode(TN1.Data).FChamp.Index-1);
                              TN.MoveTo(TN1, naInsert);


(*                              Idx := FScript.Champ.IndexOf (TN.text);
                              Idx2 := FScript.Champ.IndexOf (TN1.text);
                              FScript.Champ.Move(Idx, Idx2);
*)
                              Idx := 0; Idx2 := 0;
		              for id := 0 to FScript.Champ.Count -1 do
		              begin
                                   Idx := FScript.Champ.IndexOf (TN.text, Idx);
                                   Idx2 := FScript.Champ.IndexOf (TN1.text, Idx2);
                                   if (Idx = -1) or (Idx2 = -1) then break;
                                   FScript.Champ.Move(Idx, Idx2);
                              end;

                              (*ATri := TQSort.Create(nil);
		              ATri.Compare := FScript.CompareTableList;
		              ATri.Swap := FScript.SwapTableCorrRec;
		              ATri.DoQSort(FScript, FScript.Champ.Count);
		              ATri.Free;
                              *)
                end;
                Items.EndUpdate;
	end;
end;

procedure TExtPConcept.tvChampDragOver(Sender, Source: TObject; X,
	Y: Integer; State: TDragState; var Accept: Boolean);
var
	TN1 : TTreeNode;
begin
	Accept := False;
	if not ( Source is TTreeView) then Exit;
	TN1 := tvChamp.GetNodeAt(X, Y);
	if TN1 = nil then Accept := True
	else Accept := TN1.Level = level;
end;

procedure TExtPConcept.tvPartieDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
var
	TN, TN1 : TTreeNode;
begin
	Accept := False;
	if not ( Source is TTreeView) then Exit;
	TN := tvChamp.Selected;
	TN1 := tvChamp.GetNodeAt(X, Y);
	if TN1 = nil then Accept := TN.Parent.GetNextSibling = nil
	else if TN1.Level = TN.Level then
	begin if TN.Parent = TN1.Parent  then Accept := True; end
	else if TN1 = TN.Parent.GetNextSibling then Accept := True;
end;

procedure TExtPConcept.tvPartieDragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
	TN, TN1 : TTreeNode;
begin
	FModified := True;
	TN1 := tvChamp.GetNodeAt(X, Y);
	TN := tvChamp.Selected;
        tvChamp.Items.BeginUpdate;
	if TN1 = nil then TN.MoveTo(TN.Parent, naAddChild)
	else if TN <> TN1 then
		if TN.Level = TN1.Level then TN.MoveTo(TN1, naInsert)
		else TN.MoveTo(TN.Parent, naAddChild);
	tvChamp.Items.EndUpdate;
end;

procedure TExtPConcept.ShowInfoPanel(TNS: TTreeNode);
var
	ANode : TChampNode;
	CH : TChamp;
begin
	if TNS = nil then exit;
        Label21.visible               := FALSE;
        cbLibelleNul.Visible          := FALSE;
        groupbox3.Visible             := TRUE;

	FRefresh := True;

      if (Tvchamp.selected <> nil) and (Tvchamp.Selected.data = nil) then
	begin   // groupe de champ "Groupe.Champ"
		 tsCOMPLEMENT.TabVisible         := False;
		 tsCONSTANTE.TabVisible          := False;
		 tsCALCUL.TabVisible             := False;
		 tsREFERENCE.TabVisible          := False;

		 groupbox3.Enabled               := false;
		 groupbox1.Enabled               := false;

		 champNom.enabled                := false;
		 edLongueur.enabled              := false;
		 Combobox1.enabled               := false;
		 Label4.enabled                  := false;
		 ChampDebut.enabled              := false;
		 edfin.enabled                   := false;
		 champlon.enabled                := false;
		 lbldebut.enabled                := false;
		 label5.enabled                  := false;
		 lblfin.enabled                  := false;
                 label6.enabled                  := false;
                 label7.enabled                  := false;
                 label13.enabled                 := false;
                 label23.enabled                 := false;
                 cbType.enabled                  := false;
                 Nbdecimales.enabled             := false;
		 cbtoutevaleur.enabled           := false;
		 cbvaleurpositive.enabled        := false;
		 cbvaleurnegative.enabled        := false;

		 // Contrôle des tiers
		 ComboChampnul1.Visible          := False;
		 ComboChampnul.Visible           := False;
		 Label20.Visible                 := False;

		 cbFormatDate.Visible            := false;
		 cbFormatDatesortie.Visible      := false;
                 Label27.visible                 := false;
                 Label29.visible                 := false;

		 GroupBox3.Visible               := true;
		 cbTouteValeur.Visible           := true;
		 cbValeurPositive.Visible        := true;
		 cbValeurNegative.Visible        := true;
                 groupbox3.visible               := False;
                 Interne.visible                 := False;
                 cbTableCorr.Enabled             := False;
                 cbTableExterne.Enabled          := False;
                 Label16.Enabled                 := False;
                 Famille.Enabled                 := False;
                 ARRONDI.Enabled                 := False;
		 exit;
	end
    else
    begin
                 groupbox3.Enabled               := true;//Sélection  (Val. positive, Val. négative...)
                 groupbox1.Enabled               := true; // Positionnement
                 groupbox1.visible               := true;
                 groupbox3.visible               := true;

                if (Comment = taCreat) or ((ParametrageCegid) and (FScript.NiveauAcces = 0))
                or  ((not ParametrageCegid) and (FScript.NiveauAcces = 1)) then
                begin
                         champNom.enabled                := true;
                         if edLongueur.value = 0 then edLongueur.enabled := true;
                         Combobox1.enabled               := true;
                         Label4.enabled                  := true;

                         ChampDebut.enabled              := true;
                         edfin.enabled                   := true;
                         champlon.enabled                := true;
                         lbldebut.enabled                := true;
                         label5.enabled                  := true;
                         lblfin.enabled                  := true;
                         Nbdecimales.enabled             := true;
                         label6.enabled                  := true;
                         label7.enabled                  := true;
                         label13.enabled                 := true;
                         label23.enabled                 := true;
                         cbType.enabled                  := true;

                         cbtoutevaleur.enabled           := true;
                         cbvaleurpositive.enabled        := true;
                         cbvaleurnegative.enabled        := true;

                         groupbox5.Enabled               := true;
                         cbComplement.enabled            := true;
                         cbAlignLeft.enabled             := true;
                         Interne.Enabled                 := true;
                         cbLngRef.Enabled                := true;
                         RadioGroup1.enabled             := true;
                         Interne.visible                 := true;
                         cbTableCorr.Enabled             := True;
                         cbTableExterne.Enabled          := True;
                         Label16.Enabled                 := True;
                         Famille.Enabled                 := True;
                         ARRONDI.Enabled                 := True;
                end;
    end;
	(* affectation des boites dans le deuxieme panneau *)
    ANode := TChampNode(TNS.Data);

    CH := ANode.FChamp;
    if CH.Typ = 2 then
    begin  // champ de type Date
         cbFormatDate.ItemIndex   := CH.FormatDate;
         cbFormatDatesortie.ItemIndex := CH.FormatDatesortie;
         cbFormatDate.Visible     := true;
         Label27.visible           := true;
         Label29.visible           := true;

         cbFormatDatesortie.Visible := true;
         GroupBox3.Visible        := false;
         cbTouteValeur.Visible    := false;
         cbValeurPositive.Visible := false;
         cbValeurNegative.Visible := false;
         Nbdecimales.Visible      := False;
         Label13.Visible          := false;
    end
    else begin
         cbFormatDate.Visible     := false;
         cbFormatDatesortie.Visible := false;
         Label27.visible          := false;
         Label29.visible          := false;
         GroupBox3.Visible        := true;
         cbTouteValeur.Visible    := true;
         cbValeurPositive.Visible := true;
         cbValeurNegative.Visible := true;
         Nbdecimales.Visible      := true;
         Label13.visible          := true;
    end;
    if (script.ParName = 'TIERS')  then
    begin // reduire les choix pour les champs de la table
         ComboChampnul1.Visible := False;
         ComboChampnul.Visible := False;
         if CH.fiche  = 0 then  // si ce sont les champs de la table
         begin
            CombochampNul1.Visible := True;
            Label20.Visible := True;
         end
         else
         begin // si ce sont les champs des memos
            CombochampNul.Visible := true;
            Label20.Visible := True;
         end;
    end
    else
    begin
         ComboChampNul.Visible := False;
         ComboChampNul1.Visible := False;
         Label20.Visible := False;
    end;

	with ANode.FChamp do
	begin
		if ((Sel and 16) = 16) and (Script.FileType = 0) then
                   ChampNom.ReadOnly := True
                else ChampNom.ReadOnly := False;

		ComboBox1.ItemIndex := Typ;
		GroupBox3.Enabled := Typ = 1;// selection (Val positive, Val négative ...)
		cbTableCorr.Checked := TableExist;
		edNomTable.Text := Interprete(NomTableExt);
                ComboChampNul.itemIndex := integer(opChampNull);
                ComboChAmpNul1.itemIndex := integer(opChampNull);
                cblibellenul.itemIndex := integer(opChampNull)-1;
		case (Sel and 3) of
			0: cbTouteValeur.Checked := True;
			1: cbValeurPositive.Checked := True;
			2: cbValeurNegative.Checked := True;
		end;

		cbComplement.Checked := Compl;
                cbComplementClick(cbComplement);
		edComplLgn.Text             := IntToStr(ComplLgn);
                Nbdecimales.Text            := IntToStr(Nbdecimal);
		edComplCar.Text             := ComplCar;
		cbAlignLeft.Checked         := AlignLeft;
                cbLngRef.Checked            := LngRef;
                Cacher.checked              := Cache;
		RadioGroup1.ItemIndex       := Integer(Transform);
                Famille.Text                := FFamilleCorr;
                Interne.checked             := bInterne;
                ARRONDI.Checked             := FArrondi;
	end;
	edLongueur.value := ANode.FChamp.Siz;
        if (ANode.FChamp.Siz = 0) and (ANode.FChamp.Lon <> 0)  then
        begin
	   edLongueur.value := ANode.FChamp.Lon;
           ANode.FChamp.Siz := ANode.FChamp.Lon;
        end;
	if ANode.FChamp.Lon > 0 then
	begin
		ChampDebut.Text := IntToStr(ANode.FChamp.Deb+1);
		ChampLon.Text := IntToStr(ANode.FChamp.Lon);
		edFin.Text := IntToStr(ANode.FChamp.Deb+ANode.FChamp.Lon);
	end
        else
	begin
		ChampDebut.Text := '';
		ChampLon.Text := '';
		edFin.Text := '';
	end;

	case ANode.FChamp.TypeInfo of
		tiNull :
		begin
			Memo1.LeftCol := ANode.FChamp.Deb;
			Memo1.RightCol := ANode.FChamp.Deb+ANode.FChamp.Lon-1;
			tsCOMPLEMENT.TabVisible := True;
			if curInfo > 2  then begin
				PageControl1.ActivePage := tbDEFINITION;
			end;
			tsCONSTANTE.TabVisible := False;
			tsCALCUL.TabVisible := False;
			tsREFERENCE.TabVisible := False;
                        cbType.ItemIndex := 0;
		end;
		tiConstante :
		begin
			tsCOMPLEMENT.TabVisible := True;
			tsCONSTANTE.TabVisible := True;
			tsCALCUL.TabVisible := False;
			tsREFERENCE.TabVisible := False;
			if (curInfo > 2) and (tvChamp.Selected.level <> level)then
					PageControl1.ActivePage := tsCONSTANTE;
			CstVal.Text := ANode.FChamp.Constante.Valeur;
			CstPas.Text := IntToStr(ANode.FChamp.Constante.Pas);
			CstLon.Text := IntToStr(ANode.FChamp.Constante.Longueur);
                        cbType.ItemIndex := 1;
		end;
		tiCalcul :
		begin
			tsCOMPLEMENT.TabVisible := True;
			tsCONSTANTE.TabVisible := False;
			tsCALCUL.TabVisible := True;
			tsREFERENCE.TabVisible := False;
			if (curInfo > 2) and (tvChamp.Selected.level <> level) then
				PageControl1.ActivePage := tsCALCUL;
                        cbType.ItemIndex := 2;
		end;
		tiReference :
		begin
			tsCOMPLEMENT.TabVisible := True;
			tsCONSTANTE.TabVisible := False;
			tsCALCUL.TabVisible := False;
			tsREFERENCE.TabVisible := True;
			if (curInfo > 2) and (tvChamp.Selected.level <> level) then
				PageControl1.ActivePage := tsREFERENCE;
			RefCondition.Text := ANode.FChamp.Reference.Condition;
			RefPosition.Text := IntToStr(ANode.FChamp.Reference.Pos+1);
			RefLongueur.Text := IntToStr(ANode.FChamp.Reference.Lon);
			edRang.text := IntTOStr(ANode.Fchamp.rang);
			RefPosterieur.Checked := ANode.FChamp.Reference.Post;
                        cbType.ItemIndex := 3;
		end;
		tiConcat :
		begin
			tsCOMPLEMENT.TabVisible := True;
			tsCONSTANTE.TabVisible := False;
			tsCALCUL.TabVisible := False;
			tsREFERENCE.TabVisible := False;
			if curInfo > 2 then
				PageControl1.ActivePage := tbDEFINITION;
                        cbType.ItemIndex := 4;
		end;
	end;
	if ANode.FChamp.TypeInfo = tiConcat then
	begin
		ChampDebut.Enabled := False;
		ChampLon.Enabled := False;
	end
	else begin
		ChampDebut.Enabled := True;
		ChampLon.Enabled := True;
	end;
	PageControl1Change(PageControl1);
	FRefresh := False;

    if (not ParametrageCegid) and (FScript.NiveauAcces = 0) then
       EnabledChamp (FALSE);
end;


procedure TExtPConcept.PageControl1Change(Sender: TObject);
begin
	curInfo := PageControl1.ActivePage.PageIndex;
	PageControl1.ActivePage.Visible := True;
	case PageControl1.ActivePage.PageIndex of
	 0: ;
     // Complement
	 1: tsCOMPLEMENTEnter(tsCOMPLEMENT);
	 2: tsCALCULEnter(tsCALCUL);
	 3: ;
	 4: ;
	 5: ;
	 6: ;
	 7: ;
     // condition
     8: tsChpConditionEnter(tsChpCondition) ;
	end;
end;

procedure TExtPConcept.ComboBox1Change(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
        FModified := True;
	curChamp.Typ := ComboBox1.ItemIndex;
        ShowInfoPanel(tvChamp.Selected);
end;

procedure TExtPConcept.miNormalClick(Sender: TObject);
begin
    if tvchamp.selected.Data = nil then exit;
	if FRefresh or not Assigned(curChamp) then exit;
    FModified := True;
	curChamp.TypeInfo := tiNull;
    if tvChamp.Selected.HasChildren then
       tvChamp.Selected.DeleteChildren;
	ShowInfoPanel(tvChamp.Selected);

end;

procedure TExtPConcept.FichesauveElipsisClick(Sender: TObject);
var
SS      : string;
DD      : string;
begin
{$IFNDEF EAGLCLIENT}
        SS := SelectScript(DD);
        if SS = '' then exit;
        Fichesauve.text := SS;
        InitialiseScript;
{$ENDIF}
end;

procedure TExtPConcept.memConditionEnter(Sender: TObject);
begin
FocusCtrl := TComponent(Sender);
end;

procedure TExtPConcept.memConditionChange(Sender: TObject);
begin
if Fscript <> nil then
 	FScript.Options.Condition := memCondition.Text;
end;

procedure TExtPConcept.tsChpConditionEnter(Sender: TObject);
begin
	memCondition.OnChange := memConditionChange;
end;

procedure TExtPConcept.tsComplementEnter(Sender: TObject);
var b : Boolean;
begin
	if not Assigned(CurChamp) then exit;
	b := cbComplement.Checked;
	curChamp.Compl     := b;
	Label8.Enabled     := b;
	Label18.Enabled    := b;
	edComplCar.Enabled := b;
	edComplLgn.Enabled := b;
        cbTableExterne.checked  := curchamp.TableExterne; // |
        cbTableExterne.Enabled := curchamp.TableExist;    // |
	edNomTable.Enabled := curchamp.TableExist;        // |
end;

procedure TExtPConcept.cbComplementClick(Sender: TObject);
var b : Boolean;
begin
	if FRefresh or not Assigned(CurChamp) then exit;
	b := cbComplement.Checked;
	curChamp.Compl := b;
	Label8.Enabled := b;
	Label18.Enabled := b;
	edComplCar.Enabled := b;
	edComplLgn.Enabled := b;
	FModified := True;
end;




procedure TExtPConcept.CalFormuleEnter(Sender: TObject);
begin
	FocusCtrl := TComponent(Sender);
end;

procedure TExtPConcept.SpeedButton23Click(Sender: TObject);
begin
	InsererText(' '+TSpeedButton(Sender).Caption+' ', TMemo(CalFormule));
end;


procedure TExtPConcept.SpeedButton25Click(Sender: TObject);
begin
InsererText(' * ', TMemo(CalFormule));
end;


procedure TExtPConcept.SpeedButton29Click(Sender: TObject);
begin
	InsererText(' ENTRE ( , )', TMemo(CalFormule));
end;

procedure TExtPConcept.SpeedButton31Click(Sender: TObject);
begin
InsererText(' DATE([ , ],''  /  /19  '')', TMemo(CalFormule));
end;

procedure TExtPConcept.lbChampNumDblClick(Sender: TObject);
begin
	with lbChampNum do begin
		Clipboard.Clear;
		Clipboard.AsText := Items[itemIndex];
		CalFormule.PasteFromClipboard;
	end;

end;

procedure TExtPConcept.tsCALCULEnter(Sender: TObject);
var AChamp : TChamp;
	N: integer;
begin
   IF (cbType.ItemIndex <> 2) and (CalFormule.Text <> '') then
   begin
        miCalculClick(miCalcul);
        cbType.ItemIndex := 2;
   end;
// à voir	FCalInited := True;
	FRefresh := True;
	Calformule.SetTextBuf(PChar(curChamp.Calcul.Formule));
	FRefresh := False;
	lbChampNum.Items.Clear;
	for N:=0 to FScript.Champ.Count-1 do
	begin
		AChamp := FScript.Champ[N];
		if AChamp.Typ = 1 then
			lbChampNum.Items.Add(AChamp.Name);
	end;

end;

procedure TExtPConcept.CalFormuleChange(Sender: TObject);
begin
// à voir	if not FCalInited or FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	curChamp.Calcul.Formule := CalFormule.Text;
end;

procedure TExtPConcept.cbTypeChange(Sender: TObject);
begin
	case TComboBox(Sender).ItemIndex of
	0: miNormalClick(miNormal);
	1: miConstanteClick(miConstante);
	2: miCalculClick(miCalcul);
	3: miReferenceClick(miReference);
	4: ;
  end;
end;

procedure TExtPConcept.miCalculClick(Sender: TObject);
begin
    if tvchamp.selected.Data = nil then exit;
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	curChamp.TypeInfo := tiCalcul;
	if curChamp.Calcul = nil then
		curChamp.Calcul := TCalcul.Create;
	curInfo := 3;	(* onglet CALCUL *)
    if tvChamp.Selected.HasChildren then
       tvChamp.Selected.DeleteChildren;
	ShowInfoPanel(tvChamp.Selected);
end;

procedure TExtPConcept.miReferenceClick(Sender: TObject);
begin
    if tvchamp.selected.Data = nil then exit;
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	curChamp.TypeInfo := tiReference;
	if curChamp.Reference = nil then
		curChamp.Reference := TReference.Create;
	curInfo := 4;	(* onglet REFERENCE *)
    if tvChamp.Selected.HasChildren then
       tvChamp.Selected.DeleteChildren;
	ShowInfoPanel(tvChamp.Selected);
end;

procedure TExtPConcept.miConstanteClick(Sender: TObject);
begin
    if tvchamp.selected.Data = nil then exit;
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	curChamp.TypeInfo := tiConstante;
	if curChamp.Constante = nil then
		curChamp.Constante := TConstante.Create;
	curInfo := 5;	(* onglet CONSTANTE *)
    if tvChamp.Selected.HasChildren then
       tvChamp.Selected.DeleteChildren;
	ShowInfoPanel(tvChamp.Selected);
end;


procedure TExtPConcept.CstValChange(Sender: TObject);
begin
    if not Assigned (curChamp.Constante) then
    begin
	     cbType.ItemIndex := 1;
         curChamp.Constante := TConstante.Create;
    end;
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	curChamp.Constante.Valeur := CstVal.Text;
	CstLon.Text := IntToStr(Length(CstVal.Text));

end;

procedure TExtPConcept.CstPasChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	try curChamp.Constante.Pas := StrToInt(CstPas.Text);
	except curChamp.Constante.Pas := 0; end;

end;

procedure TExtPConcept.CstLonChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	try curChamp.Constante.Longueur := StrToInt(CstLon.Text);
	except curChamp.Constante.Longueur := 0; end;

end;

procedure TExtPConcept.tvChampChanging(Sender: TObject; Node: TTreeNode;
  var AllowChange: Boolean);
begin
if Node = nil then exit;
end;

procedure TExtPConcept.FormClose(Sender: TObject; var Action: TCloseAction);
var
  reponse : integer;
begin
  reponse := mrno;
  if not BOkValide then
    reponse := HShowMessage('0;Conception;Voulez-vous enregistrer votre Script ?;Q;YNC;N;C;','','');

    if (reponse = mryes) or (reponse = mrno) then
    begin
      if (reponse = mryes) then BValiderClick(Sender);
      Action := caFree;
      if Fscript <> nil then FScript.Destroy;
      ListeProfile.Items.clear;

      if TPControle <> nil then
      begin
        TPControle.TOBParamFree;
        TPControle.free;
        TPControle := nil;
      end;

      // On enregistre la couleur du pinceau
      SaveSynRegKey('CouleurPinceau',ColorToString(Memo1.SelColor),TRUE);

    end else begin Action := caNone; end;  // reponse = mrcancel
end;

procedure TExtPConcept.RadioGroup1Click(Sender: TObject);
begin
  if FRefresh or not Assigned(curChamp) then exit;
  curChamp.Transform := TTransformMode(RadioGroup1.ItemIndex);
  FModified := True;
end;

procedure TExtPConcept.edComplCarChange(Sender: TObject);
var S : String;
begin
  if FRefresh or not Assigned(curChamp) then exit;
  S := edComplCar.Text;
  if Length(S) > 0 then curChamp.ComplCar := S[1];
end;

procedure TExtPConcept.edComplLgnChange(Sender: TObject);
begin
  if FRefresh then exit;
  try
    curChamp.ComplLgn := StrToInt(edComplLgn.Text);
  except
    curChamp.ComplLgn := 0;
  end;
end;

procedure TExtPConcept.cbAlignLeftClick(Sender: TObject);
begin
  if FRefresh or not Assigned(CurChamp) then exit;
  curChamp.AlignLeft := cbAlignLeft.Checked;
  FModified := True;
end;

procedure TExtPConcept.cbTableCorrClick(Sender: TObject);
begin
  with Sender as TCheckBox do
  begin
    if Checked then
    begin
      GroupBox2.Enabled := True;
      cbTableExterne.Enabled := True;
      btnEditTbl.Enabled := True;
      edNomTable.Enabled := true;

      if curChamp.TableCorr = nil then
      begin
        curChamp.TableCorr := FScript.NewTableCorr;
	curChamp.TableCorr.FAssociee := true;
      end;
    end else
    begin
(*
			GroupBox2.Enabled := False;
			cbTableExterne.Enabled := False;
			btnEditTbl.Enabled := False;
            edNomTable.Enabled := False;
*)
		end;
		if curChamp.TableCorr <> nil then
			curChamp.TableCorr.FAssociee := Checked;
		curChamp.TableExist := Checked;
	end;
end;

procedure TExtPConcept.cbTableExterneClick(Sender: TObject);
begin
	FModified := True;
	with Sender as TCheckBox do
	begin
		if Checked then
		begin
           if curchamp.TableCorr <> nil then
             if curChamp.TableCorr.FEntree.Count <> 0 then
			 begin
// à voir                  ShowMessage('ATTENTION'#10'L''activation d''une table externe provoquera l''annulation de la table interne');
				{ Voulez-vous écrasez le fichier %s ? }
// à voir				            ShowMsg($04060001, ['xxxxxxx']);
             end;
		   edNomTable.Enabled := True;
                   edNomTable.Text := Interprete(curChamp.NomTableExt);
		end
		else edNomTable.Enabled := False;

		curChamp.TableExterne := Checked;
           if curChamp.TableCorr <> nil then
		  curChamp.TableCorr.FAssociee := not Checked;
	end;

end;

procedure TExtPConcept.RefConditionChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	curChamp.Reference.Condition := RefCondition.Text;
end;

procedure TExtPConcept.RefPosterieurClick(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	curChamp.Reference.Post := RefPosterieur.Checked;
end;

procedure TExtPConcept.RefPositionChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	try curChamp.Reference.Pos := StrToInt(RefPosition.Text)-1;
	except curChamp.Reference.Pos := 0; end;
  curChamp.Deb := curChamp.Reference.Pos;
    Champdebut.Text := RefPosition.Text;

end;

procedure TExtPConcept.RefLongueurChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	try curChamp.Reference.Lon := StrToInt(RefLongueur.Text);
	except curChamp.Reference.Lon := 0;  end;
  curChamp.Lon := curChamp.Reference.Lon;
  champlon.Text := RefLongueur.Text;
end;

procedure TExtPConcept.EdRangChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	try curChamp.Rang := StrToInt(edRang.Text);
	except curChamp.Rang := 0; end;
end;

procedure TExtPConcept.bZoomPieceClick(Sender: TObject);
var
    AScript   : TScript;
    I,N       : integer;
    Okplus    : Boolean;
    NameEnrg  : string;
    NameE     : string;
    UnEnreg   : Boolean;
    Count     : integer;
    ListeEnreg: HTStringList;
    OkZoom    : Boolean;
    Tvch      : TTreeview;
    FichierIE : TextFile;
    ListeOrder : HTStringList;
    procedure MajNomTabe;
    var i : integer;
    begin
             for i := 1 to FScript.Champ.Count-1 do
             begin
                  NameEnrg :=  FScript.Champ[i].Name;
                  NameEnrg := ReadTokenPipe (NameEnrg, '_');
                  FScript.Champs[i].NomFichExt := FScript.Name+NameEnrg;
             end;
    end;
begin
    If (Char(Fichesauve.Text[1]) in ['0'..'9']) Then
    begin
         PGIBox('Code incorrect, il ne doit pas contenir des champs numériques', 'Conception');
         FicheSauve.SetFocus;
         exit;
    end;
    OkZoom := FALSE;
    OkPlus := FALSE;
    Script.ExecuteModeTest := TRUE; // à voir
    UnEnreg := TRUE;
    ListeEnreg := HTStringList.Create;
    ListeOrder := HTStringList.Create;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
       Tvch := tvchampinvisible
    else
       Tvch := tvchamp;

    for N := 0 to Tvch.Items.Count -1 do
       if (Tvch.Items[N].HasChildren) and
          (Tvch.Items[N].level <> 0) then begin UnEnreg := FALSE; break; end;
    if UnEnreg then Count := 2
    else
    begin
         Count := Tvch.Items.Count;
         OkPlus := TRUE;
         MajNomTabe;
    end;
    try
        CallBackDlg := TCallBackDlg.Create(Self);
        CallbackDlg.Label1.Caption := 'Vérification de la syntaxe...';
        OkZoom := FALSE;
        if (FScript.FileDest = 1) and (FicheSortie.Text <> '') then
        begin
            AssignFile(FichierIE,  CurrentDonnee+'\'+FicheSortie.Text) ;
            Rewrite(FichierIE) ;
        end;
        for N := 1 to Count -1 do
        begin
            if (not UnEnreg) and (Tvch.Items.Item[N].level <> level) then continue;
            (*création du script*)
            AScript  := Script.CloneScript;

            if UnEnreg then NameEnrg := '' else NameEnrg := Tvch.Items.Item[N].Text;
            if NameEnrg <> '' then ListeEnreg.add (NameEnrg);

            (*affectation des champs + TFieldRec avec la fonction InitTableAndFields*)
            AScript.Assign(FScript, NameEnrg);
            if AScript.Champ[0].LOrder <> '' then ListeOrder.add (AScript.Champ[0].LOrder);

            AScript.ExecuteModeTest := TRUE;

            AScript.FEcrAna := AScript.EcrVent;
            AScript.FEcrEch := AScript.EcrEch;
            Ascript.FileName := FileName.Text;
            Ascript.options.fileName := FileName.Text;
            if (Ascript.FileName = '') or not(FileExists(Ascript.options.fileName)) then
            begin
                PGiBox('Le fichier '+Ascript.options.fileName+' est manquant'+#10#13+' L''importateur ne peut continuer', 'Conception');
                Ascript.destroy;
                CallBackDlg.Free;
                if ListeEnreg <> nil then ListeEnreg.free;
                Exit;
            end;
            (* gestion des traitements comptables à brancher aussi sur ' E X E C I M P '*)
            if Ascript.OptionsComptable.Correspondance.Count > 0 then
                for I := 0 to Ascript.OptionsComptable.correspondance.count -1 do
                    with ascript.OptionsComptable do
                        DecoupeCritere(TInfoCpt(correspondance.Items[I]).Aux, TInfoCpt(correspondance.Items[I]).listFourchette);

            (* Verification de la syntaxe *)
            AScript.ConstruitVariable(false);
            AScript.InterpreteSyntax;
            // 0 résultat en tables paradox et 1 résultat en ASCII
            AScript.FileDest := Radiotable.ItemIndex;
            AScript.ASCIIMODE := RTypefichier.ItemIndex;
            if RTypefichier.ItemIndex = 1 then
            begin
                 if cbDelimChamp.ItemIndex <> -1 then
                    AScript.Options.SepField := CarSepField[cbDelimChamp.ItemIndex]
                 else AScript.Options.SepField := cbDelimChamp.Text[1];
            end;
            AScript.Variable.Clear;
            AjouteVariableALias(AScript.Variable);

            try AScript.Executer(CallBackDlg.LaCallback);
            except
                        PGIBox('L''importation ne s''est pas correctement déroulée ','');
                        Ascript.destroy;
                        CallBackDlg.Free;
                        if ListeEnreg <> nil then ListeEnreg.free;
                        Exit;
            end;

            NameE :=  AScript.Name;
            Ascript.destroy;
            OkZoom := TRUE;
            if (FScript.FileDest = 1) and (FicheSortie.Text <> '') then
                            ExportFichierText (FichierIE, CurrentDonnee+'\'+ AScript.Name+NameEnrg+'.txt');

        end;
    finally
           if (FScript.FileDest = 1) and (FicheSortie.Text <> '') then
                      CloseFile(FichierIE);
        if OkZoom then
        begin
              CallBackDlg.Hide;
              CallBackDlg.Free;
              DMImport.DBGlobalD.Connected:=FALSE ;
              if NameEnrg <> '' then
              begin
                   if Radiotable.ItemIndex = 0  then
                      ScriptVisu (NameE, DMImport.DBGlobalD.ConnectionString,ListeEnreg,Okplus, RTypefichier.ItemIndex,'', ListeOrder);
              end
              else
                      ScriptVisu (NameEnrg, DMImport.DBGlobalD.ConnectionString,nil, FALSE,RTypefichier.ItemIndex);
              ListeEnreg.free;
              ListeOrder.free;
        end;
    end;

end;


procedure TExtPConcept.Memo1GetLine(Sender: TObject; var ALine: TASCIILine;
  var AColor: TColor; var AGlyph: Integer; ALineN: Integer);
var I, CNT: integer;
	S : TASCIILine;
begin
	if FScript.Options.TypeCar = tcOEM then
		OemToChar(ALine, ALine);
	if FScript.Options.TabSize <> 0 then
	begin
		I := 0; CNT := 0;
		while I <= Length(ALine) do
		begin
			if ALine[I] = #9 then
			begin
				repeat
					S[CNT] := ' ';
					Inc(CNT);
				until (CNT mod FScript.Options.TabSize) = 0;
			end
			else
			begin
				S[CNT] := ALine[I];
				Inc(CNT);
			end;
			Inc(I);
		end;
		S[cnt] := #0;
		ALine := S;
	end;
end;

procedure TExtPConcept.Memo1MouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  if ssLeft in Shift then MontreSel;
end;

procedure TExtPConcept.Nouvellepartie1Click(Sender: TObject);
var
	AChamp, AChampParent : TChamp;
	ANode                : TChampNode;
	N, Count             : Integer;
	ATreeNode            : TTreeNode;
begin
    if tvChamp.Selected.data = nil then exit;
	FModified := True;
	if tvChamp.Selected = nil then
	begin
		ShowMessage('Aucun champ n''est sélectionné');
		Exit;
	end;
	ATreeNode := tvChamp.Selected;
//xxxx	while (ATreeNode.Level <> level) and (AtreeNode.Parent.data <> nil)do
//xxxx		ATreeNode := ATreeNode.Parent;
	AChampParent := TChampNode(ATreeNode.Data).FChamp;
	(* creation d'un nouveau champ *)
	if AChampParent.Concat = nil then
		AChampParent.Concat := TChampList.Create(TChamp);
  if AChampParent.Concat.Count = 0 then Count := 2
	else Count := 1;
	AChampParent.TypeInfo := tiConcat;

	(* On ajoute 2 champs la premiere fois *)
	for N:=1 to Count do
	begin
		ANode := TChampNode.Create(Self);
		AChamp := AChampParent.Concat.Add as TChamp;
		ANode.FChamp := AChamp;
		AChamp.Name := 'Partie'+IntToStr(AChampParent.Concat.Count+1);
		if ATreeNode <> tvChamp.Selected then
                             tvChamp.Items.InsertObject(tvChamp.Selected, AChamp.Name, ANode)
		else         tvChamp.Items.AddChildObject(ATreeNode, AChamp.Name, ANode);
	end;
	tvChamp.Selected.Expand(True);
	tvChamp.Invalidate;
end;

procedure TExtPConcept.Champnormal1Click(Sender: TObject);
var Achamp : Tchamp;
	Anode : TChampNode;
        TN    : TTreenode;
        ParentName: string;
        TNP         : TTreenode;
begin
    if tvChamp.selected.level = level then
    begin
        FModified := True;
        tvChamp.Items.BeginUpdate;
        (* creation d'un nouveau champ *)
        AChamp := TChamp.Create(FScript.Champ);
        ANode := TChampNode.Create(Self);
        ANode.FChamp := AChamp;
        AChamp.Name := 'Champ'+IntToStr(FScript.Champ.Count+1);
        AChamp.Sel := 16;
        tvChamp.Items.AddObject(nil, AChamp.Name, ANode);
        tvChamp.Items.EndUpdate;
    end
    else
    begin
        ParentName := '';
        TNP := tvChamp.selected.Parent;
        // pour sauvegarder le nom avec Parent.Champ
        RendRacineParent(TNP, ParentName);

        FModified := True;
        tvChamp.Items.BeginUpdate;
        (* creation d'un nouveau champ *)
        AChamp := TChamp.Create(FScript.Champ);

        AChamp.Name := ParentName+'Champ';
        AChamp.Sel := 20;
        ANode := TChampNode.Create(Self);
        ANode.FChamp := AChamp;
        TN := tvChamp.Items.addObject(tvChamp.Selected, 'Champ', ANode);
        if tvChamp.Selected.level = level then
             tvChamp.Items.AddChildObject(TN, 'Champ', ANode);
        MajImageIndex (TN);
        tvChamp.Items.EndUpdate;
    end;
end;

procedure TExtPConcept.CalFormuleDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
Accept := Source = lbChampNum;
end;

procedure TExtPConcept.CalFormuleDragDrop(Sender, Source: TObject; X,
  Y: Integer);
begin
    with Source as TListBox do
    begin
		Clipboard.Clear;
		Clipboard.AsText := Items[itemIndex];
		CalFormule.PasteFromClipboard;
    end;
end;

procedure TExtPConcept.CalFormuleExit(Sender: TObject);
var
  aSVParser          : TSVParser;
  CreerSVParser      : function : TSVParser;
  ResultSyntaxe      : integer;
  i, iPosComment : integer;
   SR            : TScriptRecord;
begin
  asvparser := nil;
  if Trim(CalFormule.Text) <> '' then
  begin
    for i := 0 to CalFormule.Lines.Count-1 do
    begin
      iPosComment := pos('//', CalFormule.Lines[i]);
      if ( iPosComment > 0 ) then
        ResultSyntaxe := CheckVariables( copy(CalFormule.Text,0,iPosComment-1) )
      else
        ResultSyntaxe := CheckVariables(CalFormule.Lines[i]);

      if ( ResultSyntaxe <> 0 ) then
      begin
        AfficherErreurSyntaxeVariable(ResultSyntaxe);
        exit;
      end;
    end;

    try
      @CreerSVParser := GetProcAddress(DMImport.HdlImpfic, 'CreerSVParser');
      if not Assigned(CreerSVParser) then
      	raise Exception.Create('la fonction "CreerSVParser" n''a pas été trouvée');
      aSVParser := CreerSVParser;
      FScript.AssignSr (SR);
 			aSVParser.Script := @SR;

//      aSVParser.Script := FScript;
      aSVParser.Text := Pchar(Trim(CalFormule.Text));
      aSVParser.Mode := 0;
      aSVParser.Analyse;
      if aSVParser.GetResult^ <> #0 then
        if ( pos('//',Calformule.text) <= 0 ) and ( POS('$',TRIM(CalFormule.text)) <> 1 ) then
          // La formule n'est pas commentée et ne contient pas seulement une variable
          Pgiinfo('Erreur de syntaxe hors variables',aSVParser.GetResult);
    finally
      aSVParser.Free;
    end;
  end;
end;

procedure TExtPConcept.Memo1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
    pt    : TPoint;
begin
	if ssRight in Shift then
	begin
            pt.X := X; pt.Y := Y;
            pt := Memo1.ClientToScreen(pt);
            PopupMenu2.Popup(pt.X, pt.Y);
	end;

end;

procedure TExtPConcept.FormResize(Sender: TObject);
begin
             (*ToolWindow971.Width := (Width-tvchamp.Width);
             ToolWindow972.Width := (Width-tvchamp.Width);
             *)
end;

procedure TExtPConcept.RefConditionEnter(Sender: TObject);
begin
FocusCtrl := TComponent(Sender);
end;

procedure TExtPConcept.RefConditionExit(Sender: TObject);
var
	aSVParser     : TSVParser;
	CreerSVParser : function  : TSVParser;
  SR            : TScriptRecord;
begin
        asvparser := nil;
	if Trim(RefCondition.Text) <> '' then
	begin
		try
			@CreerSVParser := GetProcAddress(DMImport.HdlImpfic, 'CreerSVParser');
			if not Assigned(CreerSVParser) then
				raise Exception.Create('la fonction "CreerSVParser" n''a pas été trouvée');
			aSVParser := CreerSVParser;
  // Initialisation de record pour la Dll
      FScript.AssignSr (SR);
 			aSVParser.Script := @SR;
			aSVParser.Text := Pchar(Trim(RefCondition.Text));
			aSVParser.Mode := 1;
			aSVParser.Analyse;
			if aSVParser.GetResult^ <> #0 then
            begin
                Pgiinfo('Erreur',aSVParser.GetResult);
                PageControl1.ActivePage := tsREFERENCE;
                RefCondition.SetFocus;
            end;
        finally
            asvparser.free;
        end;
    end;
end;

procedure TExtPConcept.edLongueurChange(Sender: TObject);
begin
    if FRefresh or not Assigned(curChamp) then exit;

    if edLongueur.value <> 0 then
    if (curChamp.Siz <> edLongueur.value) then
    begin
                 try curChamp.Siz := edLongueur.value;
                 except curChamp.Siz := 0;  end;
    end
    else
         edLongueur.value := curChamp.Siz;

    if ((curChamp.TypeInfo = tiChamp) and (curChamp.Lon < curChamp.Siz)) or (curchamp.TypeInfo = tiCalcul) or (curchamp.TypeInfo = tiConcat) then
		 curChamp.Lon := curChamp.Siz;

end;

procedure TExtPConcept.ChampDebutChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	if IsEmpty(ChampDebut.Text) then
		curChamp.Deb := 0
	else
	begin
		curChamp.Deb := StrToInt(ChampDebut.Text)-1;
		edFin.Text := IntToStr(curChamp.Deb+curChamp.Lon);
	end;

  if curChamp.TypeInfo = tiNULL then
  begin
{$ifdef MEMOSERVANT}
		Memo1.LeftCol := curChamp.Deb;
		Memo1.RightCol := curChamp.Deb+curChamp.Lon-1;
{$endif}
  end;
end;

procedure TExtPConcept.edFinChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	if IsEmpty(edFin.Text) then
		curChamp.Lon := 0
	else
	begin
		curChamp.Lon := StrToInt(edFin.Text)-curChamp.Deb;
		if curChamp.Lon <= 0 then begin
			edFin.Text := IntToStr(curChamp.Deb+1);
			curChamp.Lon := 1;
		end;
		ChampLon.Text := IntToStr(curChamp.Lon);
	end;
	if curChamp.TypeInfo = tiNULL then
	begin
{$ifdef MEMOSERVANT}
		Memo1.LeftCol := curChamp.Deb;
		Memo1.RightCol := curChamp.Deb+curChamp.Lon-1;
{$endif}
	end;

end;

procedure TExtPConcept.ChampLonChange(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	FModified := True;
	if IsEmpty(ChampLon.Text) then
	begin
		curChamp.Lon := 0; edFin.Text := '';
	end
	else
	begin
	    curchamp.Lon := StrToInt(ChampLon.Text);
		edFin.Text := IntToStr(curChamp.Deb+curChamp.Lon);
	end;
	if curChamp.TypeInfo = tiNULL then
	begin
{$ifdef MEMOSERVANT}
		Memo1.LeftCol := curChamp.Deb;
		Memo1.RightCol := curChamp.Deb+curChamp.Lon-1;
{$endif}
	end;
end;

procedure TExtPConcept.cbFormatDateChange(Sender: TObject);
begin
	if FRefresh then exit;
   	curChamp.FormatDate := cbFormatDate.ItemIndex;
        cbFormatDatesortie.Itemindex := 3; //JJMMAAAA
        if Domainepar = 'S' then cbFormatDatesortie.Itemindex := 5; //JJMMAA
        if (Domainepar= 'M') then
             if (curChamp.Name= 'Exercice_datefin') then  cbFormatDatesortie.Itemindex := 4;
        curChamp.FormatDateSortie := cbFormatDateSortie.ItemIndex;
end;

procedure TExtPConcept.cbTouteValeurClick(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	curChamp.Sel := (curChamp.Sel and not 3);
	FModified := True;
end;

procedure TExtPConcept.cbValeurPositiveClick(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	curChamp.Sel := (curChamp.Sel and not 3) or 1;
	FModified := True;
end;

procedure TExtPConcept.cbValeurNegativeClick(Sender: TObject);
begin
	if FRefresh or not Assigned(curChamp) then exit;
	curChamp.Sel := (curChamp.Sel and not 3) or 2;
	FModified := True;
end;

procedure TExtPConcept.cblibellenulChange(Sender: TObject);
begin
     TChampNode(tvChamp.selected.data).Fchamp.OpChampNull := TOperationChampNull(cbLibelleNul.ItemIndex+1);
end;

procedure TExtPConcept.Combochampnul1Change(Sender: TObject);
begin
     CurChamp.OpChampNull := TOperationChampNull(ComboChampNul1.itemindex+1);
end;


function TExtPConcept.RendRacineParent(TN : TTreeNode; var ParentName :string) : integer;
var
    I         : integer;
    TN2       : TTreeNode;
    PName,Name: string;
begin
     Result := -1;
     if TN = nil then exit;

		TN2 := TN.GetFirstChild;
		for I := 0 to TN.Count-1 do
		   if TN2 <> nil then
		   begin
                        Name := TchampNode(TN2.Data).Fchamp.Name;
                        PName :=  ReadTokenPipe (Name, '_');
                        if (Name = tvChamp.Selected.Text) then
                        begin
                             Result := I; ParentName := PName+'_'; exit;
                        end;
                        TN2 := TN.GetNextChild(TN2);
		   end;

end;

procedure TExtPConcept.BAnnulerClick(Sender: TObject);
begin
  Close;
end;

procedure TExtPConcept.btnEditTblClick(Sender: TObject);
var
	Stream : TStream;
	N : Integer;
	P, CurrentDir : String;
	TCR : TTableCorrRec;
	slVariable : TStringList;
begin
   if curChamp.TableCorr = nil then exit;
        Stream := nil;
	CorrespDlg := TCorrespDlg.Create(nil);
	CorrespDlg.Script := FScript;
	CorrespDlg.Champ := curChamp;
 	CorrespDlg.OpTblCorr := curchamp.OpTblCorr;
	slVariable := TStringList.Create;

   try
    (* récupération de la table à partir d'un fichier*)
	if curChamp.TableExterne then
	begin
		Curchamp.NomtableExt := Interprete(Curchamp.NomTableExt);
                curChamp.IniTableCorr(slvariable);
		TCR.FEntree := curChamp.TableCorr^.FEntreeExt;
		TCR.FSortie := curChamp.TableCorr^.FSortieExt;
		CorrespDlg.Table := @TCR;
                CorrespDlg.edNomTable.Text := edNomTable.Text;
	end
	else (*les tables internes sont deja chargées*)
        CorrespDlg.Table := curChamp.TableCorr;

	CorrespDlg.Memo1.Text := curchamp.CommentTableCorr;
        CorrespDlg.Domaine.itemindex  := CorrespDlg.Domaine.items.IndexOf(Domaine.text);
	if CorrespDlg.ShowModal = mrOK then
	begin
		FModified := True;
		Curchamp.OpTblCorr := CorrespDlg.OpTblCorr;
 		curChamp.TableCorr^.FEntree.Assign(CorrespDlg.Table.FEntree);
		curChamp.TableCorr^.FSortie.Assign(CorrespDlg.Table.FSortie);
                for N:= 0 to CorrespDlg.Table.FEntree.count-1 do
                begin
                curChamp.TableCorr^.FEntree.Strings[N] :=  CorrespDlg.Table.FEntree[N];
                curChamp.TableCorr^.FSortie.Strings[N] :=  CorrespDlg.Table.FSortie[N];
                end;
		curChamp.TableCorr^.FCount := curChamp.TableCorr^.FEntree.Count;
 		curchamp.commentTablecorr := CorrespDlg.Memo1.text; // Ne pas oublier d'être en phase avec la Dll (déclaration des champs du script)
 		if curChamp.TableExterne then
		begin  // sauvegarder la table sur fichier
			try
                         if (curChamp.NomTableExt = '') then
                         begin
                              with Topendialog.create(Self) do
                              begin
                                   if Execute then
                                   begin
                                       curChamp.NomTableExt := FileName;
                                       edNomTable.Text := FileName;
                                   end;
                                   Free;
                              end;
                              SetCurrentDirectory(PChar(CurrentDir));
                         end;
                            Stream := TFileStream.Create(RecomposeChemin(InterpreteVar(curChamp.NomTableExt, slvariable)), fmCreate);
                                  except
                                     PGiBox(Exception(ExceptObject).message, 'Erreur');
                                  end;

                          // si la création a réussie
                          if (Stream <> nil) then
                          begin
                              for N:=0 to CorrespDlg.Table.FEntree.Count-1 do
                              begin
                                P := CorrespDlg.Table.FEntree.Strings[N]+';'+
                                       CorrespDlg.Table.FSortie.Strings[N]+#13+#10;
                                Stream.Write(Pointer(P)^, Length(P));
                              end;
                              Stream.Free;
                          end;
		end;
	end;
	finally
               CorrespDlg.Libere;
	       CorrespDlg.Free;
               slVariable.Free;
	end;
end;

procedure TExtPConcept.Scruter1Click(Sender: TObject);
var
    N             : Integer;
    FileNamex     : string;
    AScript       : TScript;
    NameEnrg      : string;
    SavCondition  : string;
    ListeEnreg    : HTStringList;
    OKScrut       : Boolean;
    NameEnr       : string;
begin
    if tvchamp.Selected = nil then
    begin
         PGIINFO ('Sélectionnez un champ', 'Conception');
         exit;
    end;
    if tvChamp.Selected.level = 0 then exit;
    if tvChamp.Selected.level = level then
    begin
         Script.ExecuteModeTest := TRUE; // à voir
         OKScrut := FALSE;
         ListeEnreg := HTStringList.Create;
         try
            CallBackDlg := TCallBackDlg.Create(Self);
            CallbackDlg.Label1.Caption := 'Vérification de la syntaxe...';
            (*création du script*)
            AScript  := Script.CloneScript;
            NameEnrg := tvChamp.Items.Item[tvChamp.selected.absoluteIndex].Text;
            ListeEnreg.add (NameEnrg);

            (*affectation des champs + TFieldRec avec la fonction InitTableAndFields*)
            AScript.Assign(FScript, NameEnrg);
            AScript.ExecuteModeTest := TRUE;

            AScript.FEcrAna := AScript.EcrVent;
            AScript.FEcrEch := AScript.EcrEch;

            Ascript.FileName := FileName.Text;
            Ascript.options.fileName := FileName.Text;
            if (Ascript.FileName = '') or not(FileExists(Ascript.options.fileName)) then
            begin
                PGiBox('Le fichier '+Ascript.options.fileName+' est manquant'+#10#13+' L''importateur ne peut continuer', 'Conception');
                Ascript.destroy;
                CallBackDlg.Free;
                if ListeEnreg <> nil then ListeEnreg.free;
                Exit;
            end;
            (* gestion des traitements comptables à brancher aussi sur ' E X E C I M P '*)
            if Ascript.OptionsComptable.Correspondance.Count > 0 then
                for N := 0 to Ascript.OptionsComptable.correspondance.count -1 do
                    with ascript.OptionsComptable do
                        DecoupeCritere(TInfoCpt(correspondance.Items[N]).Aux, TInfoCpt(correspondance.Items[N]).listFourchette);

            (* Verification de la syntaxe *)
            AScript.ConstruitVariable(false);
            AScript.InterpreteSyntax;
            // 0 résultat en tables paradox et 1 résultat en ASCII
            AScript.FileDest := Radiotable.ItemIndex;
            AScript.ASCIIMODE := RTypefichier.ItemIndex;
            if RTypefichier.ItemIndex = 1 then
            begin
                 if cbDelimChamp.ItemIndex <> -1 then
                    AScript.Options.SepField := CarSepField[cbDelimChamp.ItemIndex]
                 else AScript.Options.SepField := cbDelimChamp.Text[1];
            end;
            AScript.Variable.Clear;
            AjouteVariableALias(AScript.Variable);

            try AScript.Executer(CallBackDlg.LaCallback);
            except
                        PGIBox('L''importation ne s''est pas correctement déroulée ','');
                        Ascript.destroy;
                        CallBackDlg.Hide;
                        CallBackDlg.Free;
                        if ListeEnreg <> nil then ListeEnreg.free;
                        Exit;
            end;

            try NameEnr :=  AScript.Name+NameEnrg;
            except;
                   PGIBox('Champ ' + Name + ' : ' + Exception(EXceptObject).Message,' Conception');
                   Ascript.destroy;
                   CallBackDlg.Hide;
                   CallBackDlg.Free;
                   if ListeEnreg <> nil then ListeEnreg.free;
                   exit;
            end;
            Ascript.destroy;
            OKScrut := TRUE;
          finally
            if OKScrut then
            begin
              CallBackDlg.Hide;
              CallBackDlg.Free;
              DMImport.DBGlobalD.Connected:=FALSE ;
              if Radiotable.ItemIndex = 0  then
                 ScriptVisu (NameEnr, DMImport.DBGlobalD.ConnectionString,ListeEnreg,FALSE, RTypefichier.ItemIndex, RendLORDER);
              ListeEnreg.free;
            end;
          end;
          exit;
    end;

    try
        CallBackDlg := TCallBackDlg.Create(Self);
        CallbackDlg.Label1.Caption := 'Vérification de la syntaxe...';

        AScript  := Script.CloneScript;
        NameEnrg := curchamp.Name;
        NameEnrg := ReadTokenPipe (NameEnrg, '_');
        AScript.Assign(FScript, NameEnrg);
        SavCondition := FScript.options.condition;
        FScript.options.condition := AScript.options.condition;

        FileNamex := FScript.Scruter(curChamp,Fscript,CallBackDlg.CallbackScrut);
        LectureScrute (Filenamex);
        FScript.options.condition := SavCondition;
        Ascript.destroy;
    finally
        CallBackDlg.Hide;
        CallBackDlg.Free;
    end;
end;


procedure TExtPConcept.Copier1Click(Sender: TObject);
begin
          if tvChamp.Selected.Data = nil then exit;
          Champtemp := TChamp.Create(nil);
          ChampTemp.Assign(TChampNode(tvChamp.Selected.Data).FChamp);
end;

procedure TExtPConcept.Coller1Click(Sender: TObject);
var str   : string;
    sel   : Smallint;
    M     : integer;
    ANode : TchampNode;
    TN    : TTreenode;
    size  : integer;
begin
    if tvChamp.Selected.Data = nil then exit;
    if not Assigned(ChampTemp) then
           exit;
    str := TChampNode(tvChamp.Selected.Data).FChamp.Name;
    sel := TChampNode(tvChamp.Selected.Data).FChamp.sel;
    size := TChampNode(tvChamp.Selected.Data).FChamp.Siz;
    tChampNode(tvChamp.Selected.Data).FChamp.Assign(ChampTemp);
    TChampNode(tvChamp.Selected.Data).FChamp.Name := Str;
    TChampNode(tvChamp.Selected.Data).FChamp.sel := Sel;
    TChampNode(tvChamp.Selected.Data).FChamp.siz := size;
    TvChamp.Update;
    if TChampNode(tvChamp.Selected.Data).FChamp.TypeInfo = tiConcat then
	begin
           TN := tvChamp.Selected;
	   for M:=0 to TChampNode(tvChamp.Selected.Data).FChamp.Concat.Count-1 do
	   begin
	   	  ANode := TChampNode.Create(Self);
		  TChampNode(ANode).FChamp := TChamp(TChampNode(tvChamp.Selected.Data).FChamp.Concat[M]);
		  tvChamp.Items.AddChildObject(TN, TChampNode(ANode).FChamp.Name, ANode);
       end;
    end;
    TvChamp.Update;
	ShowInfoPanel(tvChamp.Selected);
end;
(*
procedure TExtPConcept.ViewFScript1Click(Sender: TObject);
var
	AForm : TFVierge;
	AMemo : TMemo;
	ST1 : TMemoryStream;
	ST2 : TMemoryStream;
	WR : TWriter;
begin
    if not ParametrageCegid then  exit;
	AForm := TFVierge.Create(self);
	AForm.Width := 400;
	AForm.Height := 300;

	AMemo := TMemo.Create(AForm);
        AMemo.Parent := AForm;
	with AForm.ClientRect do
		AMemo.SetBounds(4, 4, Right-Left-8, Bottom-Top-8);
	AMemo.Visible := True;
        AMemo.Align := AlClient;
	ST1 := TMemoryStream.Create;
	WR := TWriter.Create(ST1, 4096);
	try WR.WriteRootComponent(FScript);
	except WR.Free; ST1.Free; Exit; end;
	WR.Free;

	ST1.Seek(0,0);
	ST2 := TMemoryStream.Create;
	try
	ObjectBinaryToText(ST1, ST2);
	finally
	ST1.Free;
	end;
	ST2.Seek(0,0);
	AMemo.Lines.LoadFromStream(ST2);
	ST2.Free;
	AForm.ShowModal;
	AForm.Free;
end;
*)
procedure TExtPConcept.ViewFScript1Click(Sender: TObject);
var
	ST1 : TMemoryStream;
	ST2 : TMemoryStream;
	WR  : TWriter;
        SS  : string;
begin
{$IFNDEF EAGLCLIENT}
        // uniquement utilisateur CEGID
        if not ParametrageCegid then  exit;
	ST1 := TMemoryStream.Create;
	WR := TWriter.Create(ST1, 4096);
	try WR.WriteRootComponent(FScript);
	except WR.Free; ST1.Free; Exit; end;
	WR.Free;

	ST1.Seek(0,0);
	ST2 := TMemoryStream.Create;
	try
	ObjectBinaryToText(ST1, ST2);
	finally
	ST1.Free;
	end;
	ST2.Seek(0,0);
        SS := CurrentDossier+'\sauvegarde\script.tmp';
        ST2.SaveToFile (SS);
	ST2.Free;
  DMImport.DBGlobalD.Connected:=FALSE ;
  ScriptVisu (SS, DMImport.DBGlobalD.ConnectionString,nil, FALSE, 1);
  DeleteFile(SS);
{$ENDIF}
end;

procedure TExtPConcept.bDownClick(Sender: TObject);
begin
	FDefvariable := TFDefvariable.Create(Self);
	FDefvariable.ListVar := FScript.Variables;
	if FDefvariable.ShowModal = mrOK then
	begin
		FScript.Variables.Assign(FDefvariable.ListVar);
		FScript.ConstruitVariable(TRUE);
	end;
	FDefvariable.Free;
end;


procedure TExtPConcept.tvChampKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_DELETE : begin  if(Fscript.Modeparam <> 1) then Supprimer1Click (Sender);  end;
    VK_INSERT : Champnormal1Click (Sender);
    VK_F10    : BValiderClick(nil) ;
    Ord('C')  :
          begin
            if Shift = [ssCtrl] then
            begin
              Copier1Click(Sender);
              Key := 0;
            end;
          end;

    Ord('V')  :
          begin
            if Shift = [ssCtrl] then
            begin
              Coller1Click(Sender);
              Key := 0;
            end;
          end;
  end;

end;


procedure TExtPConcept.RTypefichierClick(Sender: TObject);
begin
      FScript.ASCIIMODE :=  RTypefichier.ItemIndex;
      if (RTypefichier.ItemIndex = 1) then
      begin
         cbDelimChamp.visible := TRUE;
         Label28.visible := TRUE;
      end
      else
      begin
         cbDelimChamp.visible := FALSE;
         Label28.visible := FALSE;
      end;
end;

procedure TExtPConcept.FichesauveExit(Sender: TObject);
begin
if edcomment.text = '' then edcomment.text := Fichesauve.Text;
If (Fichesauve.Text <> '') and (Char(Fichesauve.Text[1]) in ['0'..'9']) Then
begin
     PGIBox('Code incorrect, il ne doit pas contenir des champs numériques', 'Conception');
     Fichesauve.Text := SaveFiche;
     exit;
end;
if (not BOkValide) and  (Comment = TaModif) and (SaveFiche <> Fichesauve.Text) then
begin
    if (PGIAsk('Voulez-vous enregistrer le Script '+ Fichesauve.text + ' ?',
      'Conception')= mryes) then
    begin
      FScript.Name := Fichesauve.Text;
      BValiderClick(Sender);
      BOkValide := FALSE;
    end
    else
    Fichesauve.Text := SaveFiche;
    exit;
end
else
FScript.Name := Fichesauve.Text;

end;


procedure TExtPConcept.tvChampDeletion(Sender: TObject; Node: TTreeNode);
begin
    if (TChampNode(Node.Data) <> nil) then
	   TChampNode(Node.Data).free;
end;

procedure TExtPConcept.Couper1Click(Sender: TObject);
begin
    if ActiveControl is TCustomEdit then
       TCustomEdit(ActiveControl).CutToClipboard;
end;

procedure TExtPConcept.EnabledChamp (Enb : Boolean);
begin
         tbDefinition.Enabled     := Enb;
         tsCOMPLEMENT.Enabled     := Enb;
         tsCONSTANTE.Enabled      := Enb;
         tsCALCUL.Enabled         := Enb;
         tsREFERENCE.Enabled      := Enb;
         Memo1.Enabled            := Enb;
         edNomTable.Enabled       := Enb;
         btnEditTbl.Enabled       := Enb;
         Label6.Enabled           := Enb;
         label7.enabled           := Enb;
         Valeur.Enabled           := Enb;
         CstVal.Enabled           := Enb;
         Pas.Enabled              := Enb;
         CstPas.Enabled           := Enb;
         Longueur.Enabled         := Enb;
         CstLon.Enabled           := Enb;

         if Enb = TRUE then exit;
           groupbox3.Enabled        := Enb;
           groupbox1.Enabled        := Enb;
           champNom.enabled         := Enb;
           edLongueur.enabled       := Enb;
           Combobox1.enabled        := Enb;
           Label4.enabled           := Enb;
           ComboBox1.enabled        := Enb;

           ChampDebut.enabled       := Enb;
           edfin.enabled            := Enb;
           champlon.enabled         := Enb;
           lbldebut.enabled         := Enb;
           label5.enabled           := Enb;
           lblfin.enabled           := Enb;
           label6.enabled           := Enb;
           label7.enabled           := Enb;
           label23.enabled          := Enb;

           cbtoutevaleur.enabled    := Enb;
           cbvaleurpositive.enabled := Enb;
           cbvaleurnegative.enabled := Enb;

           ComboChampnul1.Visible   := Enb;
           ComboChampnul.Visible    := Enb;
           Label20.Visible          := Enb;

           cbFormatDate.Visible     := Enb;
           cbFormatDateSortie.Visible:= Enb;
           Label27.visible           := Enb;
           Label29.visible           := Enb;

           GroupBox3.Visible        := Enb;
           cbTouteValeur.Visible    := Enb;
           cbValeurPositive.Visible := Enb;
           cbValeurNegative.Visible := Enb;
           RadioGroup1.Enabled      := Enb;
           cbComplement.Enabled     := Enb;
           Label18.Enabled          := Enb;
           edComplCar.Enabled       := Enb;
           edComplLgn.Enabled       := Enb;
           Label8.Enabled           := Enb;
           cbAlignLeft.Enabled      := Enb;
           cbLngRef.Enabled         := Enb;
           Nbdecimales.Enabled      := Enb;
           Label13.Enabled          := Enb;
           Interne.Enabled          := Enb;
           Label23.enabled          := Enb;
           cbType.enabled           := Enb;
           Label36.Enabled          := Enb;
           cbTableCorr.Enabled      := Enb;
           cbTableExterne.Enabled   := Enb;
           Label16.Enabled          := Enb;
           Famille.Enabled          := Enb;
           ARRONDI.Enabled          := Enb;
 end;

procedure TExtPConcept.shColorMouseUp(Sender: TObject; Button: TMouseButton;
                                      Shift: TShiftState; X, Y: Integer);
var ColorName : string;
begin
ColorName := Copy(TShape(Sender).Name, 3, 255);
Screen.Cursors[crMyCursor] := LoadCursor(HInstance, LPCTSTR('Stab' + ColorName));
Memo1.SelColor := StringToColor('cl' + ColorName);
end;

procedure TExtPConcept.DOMAINEChange(Sender: TObject);
begin
Domainepar := RendDomaine(Domaine.Text);
end;

procedure TExtPConcept.BColorChange(Sender: TObject);
var ColorName : string;
begin
  ColorName := Copy(ColorToString(BColor.CurrentChoix), 3, 255);

  if (ColorName[1] in ['A'..'Z']) then
  begin
    Memo1.SelColor := StringToColor('cl' + ColorName); // Couleurs préselectionnées
    Screen.Cursors[crMyCursor] := LoadCursor(HInstance, LPCTSTR('Stab' + ColorName));
  end
  else
  begin
    Memo1.SelColor := StringToColor('$' + ColorName); // Couleurs du nuancier
    Screen.Cursors[crMyCursor] := LoadCursor(HInstance,LPCTSTR('StabBlack'));
  end;
end;

procedure TExtPConcept.Memo3Enter(Sender: TObject);
begin
  FocusCtrl := TComponent(Sender);
end;

procedure TExtPConcept.Memo3Change(Sender: TObject);
var Node : TTreeNode;
    Name : string;
    TN2  : TTreeNode;

begin
    Node := nil;
    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
       Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;
    TchampNode(Node.Data).Fchamp.ConditionChamp := Memo3.Text;
    FModified := True;
    if Memo3.Text <> '' then MontreSelCondition(Memo3.Text);
end;

procedure TExtPConcept.SpeedButton16Click(Sender: TObject);
begin
      InsererText(' '+TSpeedButton(Sender).Caption, TMEMO(Memo3));
end;

procedure TExtPConcept.bConditionClick(Sender: TObject);
var Node,TN2 : TTreeNode;
begin
    Node := nil;
    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
    Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;
    TchampNode(Node.Data).Fchamp.bCondition:= bCondition.checked;
    FModified := True;
end;

procedure TExtPConcept.EnregUniqueClick(Sender: TObject);
var Node,TN2 : TTreeNode;
begin
    Node := nil;
    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
    Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;
    TchampNode(Node.Data).Fchamp.bUnenreg:= EnregUnique.checked;
    FModified := True;
end;

procedure TExtPConcept.BMenuZoomClick(Sender: TObject);
var
  MI               : TMenuItem;
  NameE,Nm         : string;
  AChamp           : TChamp;
  TN               : TTreenode;
  idx,i,id         : integer;
  Val              : string;
  Fige             : Boolean;
begin
  MI := Sender as TMenuItem;
  NameE := MI.Name;
  if tvChamp.Selected = nil then exit;
  if tvChamp.Selected.level <> level then
    begin PgiInfo ('Insertion Impossible',''); exit; end
  else
  if FScript.Champ.IndexOf (NameE) <> -1 then
  begin PgiInfo ('Insertion impossible, '+ NameE + ' existe déjà',''); exit; end
  else
  begin
    Nm := tvChamp.Selected.Text;
    tvChamp.Items.BeginUpdate;
    Idx := FScript.Champ.IndexOf (Nm);
    if (Idx <> -1) then
       TN := tvChamp.Items.insertObject(tvChamp.Selected, NameE, nil)
    else
       TN := tvChamp.Items.addObject(tvChamp.Selected, NameE, nil);

    for i := 0 to  TPControle.TOBParam.detail.count-1 do
    begin
             with dmImport.GzImpPar do
             begin
               if TPControle.TOBParam.detail[i].GetValue('Domaine') <> Domainepar then continue;
               if TPControle.TOBParam.detail[i].GetValue('TableName') <> NameE Then continue;
               AChamp := TChamp.Create(FScript.Champ);

               if TPControle.TOBParam.detail[i].GetValue('TableName') <> '' then
                  AChamp.Name := TPControle.TOBParam.detail[i].GetValue('TableName')+'_'+TPControle.TOBParam.detail[i].GetValue('Nom')
               else
                  AChamp.Name := TPControle.TOBParam.detail[i].GetValue('Nom');
               if (TPControle.TOBParam.detail[i].GetValue('TableName') = 'Entete')  and
               (TPControle.TOBParam.detail[i].GetValue('Nom') = 'TypeAlpha') then
                    Val := TypeT
               else
                   Val  := TPControle.TOBParam.detail[i].GetValue('Contenu');
               if Val <> '' then
               begin
                  AChamp.TypeInfo := tiConstante;
                  AChamp.Constante.Valeur := Val;
                  AChamp.Constante.Pas := 0;
                  AChamp.Constante.Longueur := Length (Val);
               end;
               AChamp.Lon := TPControle.TOBParam.detail[i].GetValue('Longueur');
               AChamp.Siz := TPControle.TOBParam.detail[i].GetValue('Longueur');
               if TPControle.TOBParam.detail[i].GetValue('Obligatoire') = 'O' then
                  AChamp.Sel := 16
               else
                  AChamp.Sel := 20;
               Fige := (TPControle.TOBParam.detail[i].GetValue('Fige')=1);
               AChamp.Cache :=  ((AChamp.Sel = 16) and (not Fige));
               renseignetree(Achamp, TN, tvchamp);
             end;
    end;
//i := FScript.Champ.IndexOf (NameE);
//  FScript.Champ.Move(i, idx);

    I := 0; Idx := 0;
    for id := 0 to FScript.Champ.Count -1 do
    begin
         i := FScript.Champ.IndexOf (NameE, i);
         Idx := FScript.Champ.IndexOf (Nm, idx);
         if (I = -1) or (Idx = -1) then break;
         FScript.Champ.Move(I, Idx);
    end;

    tvChamp.Items.EndUpdate;
  end;
end;

function TExtPConcept.CreerLigPop ( Name : string; Owner : TComponent) : TMenuItem ;
Var
    T   : TMenuItem ;
BEGIN
T         := TMenuItem.Create(Owner) ;
T.Name    := Name ;
T.Caption := T.Name;
T.OnClick := BMenuZoomClick;
Result    := T ;
END ;

procedure TExtPConcept.BrepriseClick (Sender: TObject);
var
   T,N   : TMenuItem ;
   i,d   : integer;
   Okadd : Boolean;
begin
  inherited;
      if not assigned (TPControle) then exit;
      While POPZ.Items.Count>0 do
         BEGIN
         T:=POPZ.Items[0] ;
         While T.Count>0 do BEGIN N:=T.Items[0] ; T.Remove(N) ; N.Free ; END ;
         POPZ.Items.Remove(T) ; T.Free ;
         END ;

      if TPControle.TOBParam <> nil then
          for i := 0 to  TPControle.TOBParam.detail.count-1 do
          begin
               Okadd := FALSE;
               for d := 0 to  POPZ.Items.count-1 do
               begin
                          if POPZ.Items[d].Name = TPControle.TOBParam.detail[i].GetValue('TableName') then
                          begin Okadd := TRUE; break; end;
               end;
               if not Okadd then
               begin
                            T:=CreerLigPop(TPControle.TOBParam.detail[i].GetValue('TableName'),POPZ);
                            POPZ.Items.Add(T) ;
               end;
          end;
end;

procedure TExtPConcept.BRechercheClick(Sender: TObject);
begin
      FRecherche := TFRecherche.create(application);
      if FindText <> '' then FRecherche.RechText.text := FindText;
      While FRecherche.ShowModal = mrOK do
      begin
           if not FRecherche.RespecterMaj.checked then
              Memo1.Find(FRecherche.RechText.text, foIgnoreCase, 0, 1)
           else
              Memo1.Find(FRecherche.RechText.text, foFromCaret, 0, 1);
           FindText := FRecherche.RechText.text;
           Memo1.LeftSel := 1;
           Memo1.RightSel := LNGLMAX;
           MontreSel;
      end;
      FRecherche.free;
end;

// Automate qui vérifie la syntaxe des variables $() dans un champ texte
// Result = 1   <=>  Erreur d'ouverture de parenthèse de variable
// Result = 2   <=>  Erreur de fermeture de parenthèse de variable
// Result = 3   <=>  Variable non déclarée ou mal orthographiée
function TExtPConcept.CheckVariables(chaine : string) : integer;
var
  compt, etat, debut    : integer;
begin
  etat := 1;
  result := 0; debut := 0;

  for compt := 1 to strlen(Pchar(chaine)) do
  begin
    case etat of
      1: if chaine[compt] = '$' then etat := 2;
      2: if chaine[compt] = '(' then
         begin
           etat := 3;
           debut := compt;
         end
         else
         begin
           result := 1;
           exit;
         end;
      3: begin if chaine[compt] = ')' then
               begin
                 etat := 1;
                 if ( ExistVariable(chaine,debut,compt) = false ) then
                 begin
                   result := 3;
                   exit;
                 end;
               end
               else if ( chaine[compt] ='(' ) or ( chaine[compt]=' ' ) then
               begin
                 result := 2;
                 exit;
               end;
         end;
    end;    // fin case
  end;     // fin for

  if etat = 2 then result := 1 else if etat = 3 then result := 2;
end;

// Vérifier si une variable entrée par un utilisateur
// dans un champ Condition ou Calcul est définie ou non
function TExtPConcept.ExistVariable(chaine : string; debut : integer; fin : integer) : boolean;
var
  compt : integer;
  chainetmp : string;
begin
  result := false;
  chainetmp := lowercase( copy(chaine,debut+1,fin-debut-1) );

  // La variable existe-t'elle ? (a-t'elle été déclarée ?)
  for compt := 0 to FScript.Variables.Count-1 do
  begin
    if ( lowercase(FSCript.Variables.Items[compt].Name) = chainetmp ) then
    begin
      result := true;
      exit;
    end;
  end;
end;

procedure TextPConcept.AfficherErreurSyntaxeVariable (erreur : integer);
begin
  if erreur = 1 then
  begin PGIinfo('Ouverture de parenthese','Erreur de syntaxe de variable'); end
  else if erreur = 2 then
  begin PGIinfo('Fermeture de parenthese','Erreur de syntaxe de variable'); end
  else if erreur = 3 then
  begin PGIinfo('Une variable n''est pas déclarée ou est mal orthographiée','Erreur de syntaxe de variable'); end;
end;

procedure TExtPConcept.ChecksyntaxeClick(Sender: TObject);
var
  aSVParser : TSVParser;
  CreerSVParser : function  : TSVParser;
  ResultSyntaxe : integer;
  compt : integer;
  SR            : TScriptRecord;
begin
  aSVParser := nil;
  if Trim(Memo3.Text) <> '' then
    // verification ligne par ligne
    for compt := 0 to Memo3.Lines.Count-1 do
    begin
      if ( pos('//', Memo3.Lines[compt]) > 0 ) then
      begin
        // On analyse que ce qu'il y a avant le commentaire
        ResultSyntaxe := CheckVariables( copy(Memo3.Text,0,pos('//', Memo3.Lines[compt])-1) );
      end else
      begin ResultSyntaxe := CheckVariables(Memo3.Lines[compt]); end;

      if (ResultSyntaxe <> 0) then
      begin
        AfficherErreurSyntaxeVariable(ResultSyntaxe);
        exit;
      end;

    end;

    try
      @CreerSVParser := GetProcAddress(DMImport.HdlImpfic, 'CreerSVParser');
      if not Assigned(CreerSVParser) then
        raise Exception.Create('la fonction "CreerSVParser" n''a pas été trouvée');
	aSVParser := CreerSVParser;
  FScript.AssignSr (SR);
  aSVParser.Script := @SR;

//	aSVParser.Script := FScript;
	aSVParser.Text := Pchar(Trim(Memo3.Text));
	aSVParser.Mode := 1;
	aSVParser.Analyse;
	if aSVParser.GetResult^ <> #0 then
		Pgiinfo('Erreur de syntaxe hors variables',aSVParser.GetResult);
    finally
      aSVParser.Free;
    end;
end;

procedure TExtPConcept.ModeParamClick(Sender: TObject);
begin
  ModeVisu := Modeparam.Itemindex;
end;

procedure TExtPConcept.RejetClick(Sender: TObject);
begin
 if (FocusCtrl is TEdit) and (PageControl1.ActivePage = tsREFERENCE) and (FocusCtrl.Name = 'RefCondition') then
 begin
    Refcondition.Text := '['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ Memo1.SelectText + ''''
 end
 else
 if (FocusCtrl is THRichEditOLE) and (FocusCtrl.Name = 'memCondition') then
 begin
  if MemCondition.text = '' then
     Memcondition.Text := '['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ Memo1.SelectText + ''''
  else
  begin
   if SaveCurRow = Memo1.curRow then
     Memcondition.Text := Memcondition.Text + ' ET ['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ (*ClipBoard.AsText*)Memo1.SelectText + ''''
   else
     Memcondition.Text := Memcondition.Text + ' OU ['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ Memo1.SelectText + ''''
  end;
  SaveCurRow := Memo1.curRow;
 end;
 if (FocusCtrl is THRichEditOLE) and (FocusCtrl.Name = 'Memo3')then
 begin
  if Memo3.text = '' then
     Memo3.Text := ' ['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ Memo1.SelectText + ''''
  else
     Memo3.Text := Memo3.Text + ' ET ['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ (*ClipBoard.AsText*)Memo1.SelectText + ''''
 end;
 if (FocusCtrl is TColorMemo) and (FocusCtrl.Name = 'CalFormule')then
 begin
  if CalFormule.text = '' then
     CalFormule.Text := ' ['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ Memo1.SelectText + ''''
  else
     CalFormule.Text := CalFormule.Text + ' ET ['+lbDeb.Caption+','+lbLon.Caption+']<>'''+ Memo1.SelectText + ''''
 end;

end;

procedure TExtPConcept.ProfileElipsisClick(Sender: TObject);
begin
Profile.Text := SelectCorresp( ChampNom.Text);
end;

procedure TExtPConcept.BValideClick(Sender: TObject);
var
i         : integer;
trouve    : Boolean;
Node,TN2    : TTreeNode;
begin
    trouve := FALSE;  Node := nil;
    for i := 0 to Listeprofile.Items.count-1 do
    begin
        if Listeprofile.Items.Strings[i]= (Profile.Text)  then
        begin trouve := TRUE; break; end;
    end;
    if not trouve then
       Listeprofile.Items.Add(Profile.Text);

    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
    Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;
    TchampNode(Node.Data).Fchamp.ListProfile.clear;
    for I := 0 to ListeProfile.Items.count - 1 do
      TchampNode(Node.Data).Fchamp.ListProfile.add(ListeProfile.Items.Strings[i]);
    FModified := True;

end;

procedure TExtPConcept.bDefaireClick(Sender: TObject);
var
Node,TN2    : TTreeNode;
begin
    Node := nil;
    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
    Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;
    TchampNode(Node.Data).Fchamp.ListProfile.clear;
    ListeProfile.Items.clear;
    FModified := True;
end;

procedure TExtPConcept.tvchampinvisibleDeletion(Sender: TObject;
  Node: TTreeNode);
begin
    if (TChampNode(Node.Data) <> nil) then
	   TChampNode(Node.Data).free;
end;

procedure TExtPConcept.Affecterligne1Click(Sender: TObject);
begin
 if (FocusCtrl is TEdit) and (PageControl1.ActivePage = tsREFERENCE) and (FocusCtrl.Name = 'RefCondition') then
 begin
    Refcondition.Text := '['+lbDeb.Caption+','+lbLon.Caption+']='''+ Memo1.SelectText + ''''
 end
 else
 if (FocusCtrl is THRichEditOLE) and (FocusCtrl.Name = 'memCondition')then
 begin
  if MemCondition.text = '' then
     Memcondition.Text := '['+lbDeb.Caption+','+lbLon.Caption+']='''+ Memo1.SelectText + ''''
  else
     Memcondition.Text := Memcondition.Text + ' ET ['+lbDeb.Caption+','+lbLon.Caption+']='''+ Memo1.SelectText + ''''
 end;
 if (FocusCtrl is THRichEditOLE) and (FocusCtrl.Name = 'Memo3')then
 begin
  if Memo3.text = '' then
     Memo3.Text := ' ['+lbDeb.Caption+','+lbLon.Caption+']='''+ Memo1.SelectText + ''''
  else
     Memo3.Text := Memo3.Text + ' ET ['+lbDeb.Caption+','+lbLon.Caption+']='''+ Memo1.SelectText + ''''
 end;
 if (FocusCtrl is TColorMemo) and (FocusCtrl.Name = 'CalFormule')then
 begin
  if CalFormule.text = '' then
     CalFormule.Text := ' ['+lbDeb.Caption+','+lbLon.Caption+']='''+ Memo1.SelectText + ''''
  else
     CalFormule.Text := CalFormule.Text + ' ET ['+lbDeb.Caption+','+lbLon.Caption+']='''+ Memo1.SelectText + ''''
 end;

end;

procedure TExtPConcept.NbdecimalesChange(Sender: TObject);
begin
	if FRefresh then exit;
        try
        	curChamp.NbDecimal := StrToInt(NbDecimales.Text);
        except
                curChamp.NbDecimal := 2;
        end;

end;

procedure TExtPConcept.edNomTableChange(Sender: TObject);
begin
	if FRefresh then exit;
        curChamp.NomTableExt := edNomTable.Text;
end;

procedure TExtPConcept.cbFormatDatesortieChange(Sender: TObject);
begin
 	if FRefresh then exit;
   	curChamp.FormatDateSortie := cbFormatDateSortie.ItemIndex;

end;

procedure TExtPConcept.LIENINTERChange(Sender: TObject);
var Node,TN2 : TTreeNode;
begin
    Node := nil;
    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
       Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;

    TchampNode(Node.Data).Fchamp.LienInter := InterpreteVarSyntaxSt(LienInter.Text);
    FModified := True;
end;

procedure TExtPConcept.cbTypeCarChange(Sender: TObject);
begin
if Fscript <> nil then
   FScript.Options.TypeCar := TTypeCar(cbTypeCar.ItemIndex);
end;

procedure TExtPConcept.ORDERChange(Sender: TObject);
var Node,TN2 : TTreeNode;
begin
    Node := nil;
    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
       Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;

    TchampNode(Node.Data).Fchamp.LOrder := ORDER.Text;
    FModified := True;
end;

function TExtPConcept.RendLORDER : string;
var Node,TN2 : TTreeNode;
begin
    Result := ''; Node := nil;
    if tvchamp.selected = nil then exit;
    if Fscript.Modeparam = 1 then  // MODE SIMPLIFIE
    begin
                Name := tvchamp.selected.text;
		TN2 := tvchampinvisible.Items.GetFirstNode;
                while TN2 <> nil do
                begin
                     if TN2.Text = Name then
                     begin
                        Node := TN2.getFirstChild;
                        break;
                     end;
                     TN2 := TN2.GetNext;
                end;
    end
    else
       Node := tvchamp.selected.getFirstChild;
    if Node = nil then exit;
    if TchampNode(Node.Data).Fchamp.LOrder <> '' then
       Result := ' ORDER BY ' + TchampNode(Node.Data).Fchamp.LOrder;
end;

procedure TExtPConcept.bOptionsClick(Sender: TObject);
begin
	OptionsDlg := TOptionsDlg.Create(Self);
	try
		OptionsDlg.ScriptSuivant := FScript.ScriptSuivant;
		//OptionsDlg.Millier := Fscript.options.millier;
		//OptionsDlg.Decimal := Fscript.options.Decimal;

		if OptionsDlg.ShowModal = mrOK then
		begin
                     FScript.ScriptSuivant := OptionsDlg.ScriptSuivant;

                end;
        finally
          OptionsDlg.Free;
       end;

end;

procedure TExtPConcept.FicheSortieChange(Sender: TObject);
begin
FScript.Options.ASCIIFileName := FicheSortie.Text;
end;

procedure TExtPConcept.RadiotableClick(Sender: TObject);
begin
FicheSortie.visible := (Radiotable.ItemIndex = 1);
Label31.visible     := (Radiotable.ItemIndex = 1);
Table.visible := (Radiotable.ItemIndex = 1);
Label32.visible     := (Radiotable.ItemIndex = 1);
end;

procedure TExtPConcept.TableElipsisClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
Table.Text := SelectScript(Domainepar);
{$ENDIF}
end;

procedure TExtPConcept.TableChange(Sender: TObject);
begin
FScript.ScriptSuivant := Table.Text;
end;

procedure TExtPConcept.ChargementCombotableN;
var
  Vales   : HTstringList;
  i,id    : integer;
  NN      : string;
begin
    TableN.clear;
    Vales := HTStringList.Create;
    for id := 0 to FScript.Champ.Count -1 do
    begin
         NN := FScript.Champ.Items[id].Name;
         i := pos('_', FScript.Champ.Items[id].Name);
         if i <> 0 then NN := copy(FScript.Champ.Items[id].Name, 1, i-1);
         if (Vales.IndexOf(NN) < 0) then
          Vales.add(NN);
    end;

    TableN.Items := Vales;
    TableN.Values := Vales;
    Vales.clear;
    Vales.Free;
end;

procedure TExtPConcept.ChargementCodechamp(Name : string);
var
  Vales   : HTstringList;
  i,id    : integer;
  NN      : string;
begin
    CodeChamp.clear;
    Vales := HTStringList.Create;
    for id := 0 to FScript.Champ.Count -1 do
    begin
         i := pos('_', FScript.Champ.Items[id].Name);
         if i <> 0 then
         NN := copy(FScript.Champ.Items[id].Name, 1, i-1)
         else
         NN := FScript.Champ.Items[id].Name;
         if NN <> Name then continue;
         if i <> 0 then NN := copy(FScript.Champ.Items[id].Name, i+1, length(FScript.Champ.Items[id].Name));
         if (Vales.IndexOf(NN) < 0) then
          Vales.add(NN);
    end;

    CodeChamp.Items := Vales;
    CodeChamp.Values := Vales;
    Vales.clear;
    Vales.Free;
end;


procedure TExtPConcept.TabLienEnter(Sender: TObject);
begin
ChargementCombotableN;
end;

procedure TExtPConcept.TableNExit(Sender: TObject);
begin
ChargementCodechamp(TableN.Text);
end;

procedure TExtPConcept.BValideLienClick(Sender: TObject);
begin
 if  TableN.text = '' then
 begin
      PgiInfo ('Saisir le nom de la table','Conception');
      exit;
 end;
 if LienInter.Text = '' then
  LienInter.Text := TableN.text + '/' +CodeChamp.Text
 else
  LienInter.Text := LienInter.Text + ',' + CodeChamp.Text;
end;

procedure TExtPConcept.BEffLienClick(Sender: TObject);
begin
   LienInter.Text := '';
end;

procedure TExtPConcept.EdShellExecuteChange(Sender: TObject);
begin
if Fscript = nil then exit;
  FScript.Shellexec := EdShellExecute.Text;
end;

procedure TExtPConcept.ToolbarButton971Click(Sender: TObject);
begin
  ExecuteShell(EdShellExecute.Text);
end;

procedure TExtPConcept.FamilleChange(Sender: TObject);
begin
	if FRefresh then exit;
        curChamp.FFamilleCorr := Famille.Text;
end;

procedure TExtPConcept.CbLngRefClick(Sender: TObject);
begin
 	if FRefresh or not Assigned(CurChamp) then exit;
	curChamp.LngRef := cbLngRef.Checked;
	FModified := True;
end;

procedure TExtPConcept.RefPositionEnter(Sender: TObject);
begin
FocusCtrl := TComponent(Sender);
end;

procedure TExtPConcept.RefLongueurEnter(Sender: TObject);
begin
FocusCtrl := TComponent(Sender);
end;

procedure TExtPConcept.SpeedButton43Click(Sender: TObject);
begin
	RefCondition.Text := RefCondition.Text + TSpeedButton(Sender).Caption;
end;

procedure TExtPConcept.InterneClick(Sender: TObject);
begin
	if FRefresh or not Assigned(CurChamp) then exit;
	curChamp.bInterne := Interne.Checked;
	FModified := True;
end;

procedure TExtPConcept.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

procedure TExtPConcept.ValidClickRichedit(Sender: TObject);
var
  PrinterSetupDialog1 : TPrinterSetupDialog;
begin
          PrinterSetupDialog1 := TPrinterSetupDialog.create(Application);
          if PrinterSetupDialog1.Execute then
          AMemoPrint.Print ('Conception');
          PrinterSetupDialog1.Free;
end;

procedure TExtPConcept.BImprimerClick(Sender: TObject);
var
	AForm             : TFVierge;
        I, Debut          : Integer;
        AChamp            : TChamp;
        St                : string;
        ChMaitre, ChEnfant, PrevChMaitre : string;
begin
	AForm := TFVierge.Create(self);
	AForm.Width := 400;
	AForm.Height := 300;

	AMemoPrint := TRichEdit.Create(AForm);
        with AMemoPrint do
        begin
                    Parent := AForm;
                    with AForm.ClientRect do
                            SetBounds(4, 4, Right-Left-8, Bottom-Top-8);
                    Visible := True;
                    Align := AlClient;
                    Paragraph.Numbering := nsBullet;
                    Lines.add ( 'Nom : ' + Fscript.Name);
                    Lines.add ( 'Fichier : ' + Fscript.Options.FileName);

                   if Fscript.Options.Condition <> '' then
                        Lines.add ( 'Condition : ' + Fscript.Options.Condition);
                   if Fscript.ASCIIMode = 1 then
                      Lines.add ( 'Type de table ASCII')
                   else
                      Lines.add ( 'Type de table Standard');

                    For I := 0 to  FScript.Champ.Count-1 do
                    begin
                         AChamp := FScript.Champ[I];
                         chMaitre := copy(Achamp.name, 1, pos('_', Achamp.Name)-1);
                         chEnfant := copy(Achamp.name, pos('_', Achamp.Name)+1, length(Achamp.Name));
                         if PrevChMaitre <> ChMaitre then // pour réduire le temps de chargement de la forme
                         begin
                             Paragraph.Numbering := nsNone;
                             Paragraph.Alignment := taLeftJustify;
                             SelAttributes.Style := [fsBold, fsUnderline];
                             Lines.add ( '');
                             Lines.add (chMaitre);
                             if AChamp.ConditionChamp <> '' then
                             begin
                                  Lines.add ( '');
                                  Lines.add ( 'Condition Complémentaire : ' + AChamp.ConditionChamp);
                             end;
                             PrevChMaitre := ChMaitre;
                         end;
                         if Achamp.Sel = 16 then
                         begin
                             Paragraph.Numbering := nsNone;
                             Paragraph.Alignment := taLeftJustify;
                             SelAttributes.Style := [fsBold];
                             Lines.add ( '');
                             Lines.add ( '     ' + chEnfant);
                             Paragraph.Numbering := nsNone;
                             Paragraph.Alignment := taCenter;
                             Lines.add ( '');
                             Debut := AChamp.Deb;
                             if Debut = 0 then Debut := 1;
                             Lines.add ( 'Type : ' +
                             ComboBox1.Items[Achamp.Typ]+ ', Longueur : '+
                             IntToStr(AChamp.Siz)+ ' Début : ' + IntToStr(Debut) + ' Fin :' +  IntToStr(AChamp.Lon)) ;

                             case AChamp.Transform of
                                   tmAUCUN :     Lines.add ('Transformation : Aucune   ');
                                   tmMAJUSCULE : Lines.add ('Transformation : Majuscule');
                                   tmMINUSCULE : Lines.add ('Transformation : Minuscule');
                             end;
                             if AChamp.Compl then
                             begin
                               St := 'Complèter la zone avec : ' + Achamp.ComplCar;
                               if Achamp.LngRef then
                               St := St + ' Longueur minimum : ' + IntToStr(AChamp.ComplLgn)
                               else
                               St := St + ' Longueur : ' + IntToStr(AChamp.ComplLgn);
                               if Achamp.AlignLeft then
                               St := St + ' Aligné à guauche';
                               Lines.add (St);
                             end;
                             case AChamp.TypeInfo of
                                       tiNull      : Lines.add ('Type de champ : Normal   ');
                                       tiConstante : Lines.add ('Type de champ : Constante');
                                       tiCalcul    : Lines.add ('Type de champ : Calcul   ');
                                       tiReference : Lines.add ('Type de champ : Référence');
                             end;
                             if (AChamp.TypeInfo = tiCalcul) and (AChamp.Calcul.Formule <> '') then
                                Lines.add ( 'Formule : ' + AChamp.Calcul.Formule);
                             if (AChamp.TypeInfo = tiReference) and (AChamp.Reference <> nil) then
                             begin
                                Lines.add ( 'Condition du référence : ' + AChamp.Reference.Condition);
                                Lines.add ( 'Pos : ' + IntTostr(AChamp.Reference.Pos+1));
                                Lines.add ( 'Longueur : ' + IntTostr(AChamp.Reference.Lon));
                                Lines.add ( 'Rang : ' + IntTostr(AChamp.rang));
                                if AChamp.Reference.Post then
                                Lines.add ( 'Référence postérieure');
                             end;
                             if (AChamp.TypeInfo = tiConstante) and (AChamp.Constante <> nil) then
                             begin
                                Lines.add ( 'Valeur du constante : ' + AChamp.Constante.Valeur);
                                Lines.add ( 'Longueur : ' + IntTostr(AChamp.Constante.Longueur));
                                if AChamp.Constante.Pas <> 0 then
                                Lines.add ( 'Pas : ' + IntTostr(AChamp.Constante.Pas));
                             end;
                         end;
                    end;
        end;
        AForm.Caption := Fscript.Name;
        AForm.BImprimer.Visible := TRUE;
        AForm.BValider.Visible := FALSE;
        AForm.BorderStyle := BsSizeable;
        AForm.BorderIcons := [biSystemMenu, biMaximize];
        AForm.BImprimer.OnClick :=  ValidClickRichedit;
	AForm.ShowModal;
	AForm.Free;
end;


procedure TExtPConcept.Imprimer1Click(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
PrintPageDeGarde(PageControl1, TRUE, TRUE, FALSE, Caption, 0);
{$ENDIF}
end;

procedure TExtPConcept.SpeedButton55Click(Sender: TObject);
begin
     InsererText('LIGNE(+/-X)', TMemo(CalFormule));
end;

procedure TExtPConcept.SpeedButton56Click(Sender: TObject);
begin
     InsererText(' <> ', TMemo(CalFormule));
end;

procedure TExtPConcept.ApercuConditionCompClick(Sender: TObject);
var
   NomEnreg   : string;
   ScriptTmp  : TScript;
   NomFichier : string;
begin
   ScriptTmp := nil;
   if(Memo3.text <> '') then
   begin
        try
           CallBackDlg := TCallBackDlg.Create(Self);
           CallbackDlg.Label1.Caption := 'Vérification de la syntaxe...';
           ScriptTmp  := Script.CloneScript;
           NomEnreg   := tvChamp.Items.Item[tvChamp.selected.absoluteIndex].Text;
           ScriptTmp.Assign(FScript, NomEnreg);
           ScriptTmp.Champ[0].Deb  := 0;        // Début au premier caractère
           ScriptTmp.Champ[0].Lon  := Memo1.LineMax;
           ScriptTmp.Champ[0].Siz  := Memo1.LineMax;
           ScriptTmp.Champ[0].Typ  := 0;        // Alphanumérique
           ScriptTmp.Champ[0].Sel  := 0;
           ScriptTmp.Champ[0].TypeInfo := tiNull;

           curchamp := ScriptTmp.Champ[0];
           NomFichier := ScriptTmp.Scruter(curchamp,ScriptTmp,CallBackDlg.CallbackScrut);
           LectureScrute (NomFichier);
        finally
           ScriptTmp.destroy;
           CallBackDlg.Hide;
           CallBackDlg.Free;
        end;
   end
   else begin
        // La condition est vide
        PGIinfo('Il n''y a pas de condition','Erreur');
   end;
end;

procedure TExtPConcept.ARRONDIClick(Sender: TObject);
begin
	if FRefresh or not Assigned(CurChamp) then exit;
	curChamp.FArrondi := ARRONDI.Checked;
	FModified := True;
end;

end.
