unit EtbMce;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls, Buttons, Hctrls, IniFiles, hmsgbox,
  HSysMenu, HPanel, Uiutil, HTB97, TntStdCtrls ;

procedure ParamTeletrans ;

const
SECTION_BQ		='BANQUE';
SECTION_MODELE		='MODELE';
SECTION_PORT		='PORT';
SECTION_MODEM		='MODEM';
SECTION_CARTE		='CARTE';
SECTION_TEL		='TELEPHONE';
SECTION_TRACE		='TRACE';

FILE_BQ			='BANQUE';
FILE_MODELE		='MODELE';
FILE_MODEM		='MODEM';
FILE_CARTE		='CARTE';
FILE_SCENARIO		='TRANS';
FILE_GEN		='GENERAL';
FILE_DEBUG		='DEBUG';
FILE_SIM		='$$EEPROM' ;

SIM_EXT			='.SIM' ;

FILENAME_DEBUG		='DEBUG.SYS';

MAX_ENTRY		=999;

ETEBAC_PATH		='ETEBAC\';
ETEBAC_DEBUG		='DEBUG';
ETEBAC_EXT		='.INI';
ETEBAC_PAR		='.PAR';
ETEBAC_LOG		='.LOG';
ETEBAC_SCN		='.SCN';
ETEBAC_DAT		='.DAT';
ETEBAC_SYS		='.SYS';
ETEBAC_ASYNC            = '0836064444' ;
ETEBAC_SYNC             = '0836063232' ;

ERR_LOGMISSING		=1000;
ERR_DLLMISSING		=1001;
ERR_PROCMISSING		=1002;
ERR_USERCANCEL		=1003;

DLL_NAME                ='W32L341.DLL';
DLL_DEBUGNAME           ='W32L341DBG.DLL';
DLL_PROCNAME            ='w32trx';

WORD_RETURN	        ='return(';
LEN_RETURN	        =7;
WORD_OUVRE	        ='P_ouvre(nom=com';
LEN_OUVRE	        =15;
WORD_FERME	        ='P_ferme(';
LEN_FERME	        =8;
WORD_MEM	        ='mem=';
LEN_MEM		        =4;
WORD_INIT	        ='P_init(';
LEN_INIT	        =7;
WORD_LIRE	        ='P_lire(';
LEN_LIRE	        =7;
WORD_ECRIRE	        ='P_ecrire(';
LEN_ECRIRE	        =9;
WORD_COMERR             ='GetCommError_Status';
LEN_COMERR	        =19;

type
  TFMce = class(TForm)
    PageControl: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    BAide: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BValider: TToolbarButton97;
    BConnect: TToolbarButton97;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    LBDest: TListBox;
    GroupBox1: TGroupBox;
    HLabel1: THLabel;
    ECode: TEdit;
    HLabel2: THLabel;
    EGuichet: TEdit;
    HLabel3: THLabel;
    ENomDest: TEdit;
    GroupBox2: TGroupBox;
    HLabel4: THLabel;
    EAdr: TEdit;
    HLabel5: THLabel;
    EExt: TEdit;
    CPcv: TCheckBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    ENomContact: TEdit;
    HLabel6: THLabel;
    ETelContact: TEdit;
    HLabel7: THLabel;
    HLabel8: THLabel;
    EFaxContact: TEdit;
    HLabel9: THLabel;
    ENomTech: TEdit;
    HLabel10: THLabel;
    ETelTech: TEdit;
    HLabel11: THLabel;
    EFaxTech: TEdit;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    GroupBox7: TGroupBox;
    LBDest2: TListBox;
    LBModele: TListBox;
    Label1: TLabel;
    LPos: TLabel;
    ECarte: TEdit;
    HLabel12: THLabel;
    GroupBox8: TGroupBox;
    RBCom1: TRadioButton;
    RBCom2: TRadioButton;
    RBCom3: TRadioButton;
    RBCom4: TRadioButton;
    RBCom5: TRadioButton;
    RBCom6: TRadioButton;
    RBCom7: TRadioButton;
    RBCom8: TRadioButton;
    GroupBox9: TGroupBox;
    RB300: TRadioButton;
    RB9600: TRadioButton;
    RB19200: TRadioButton;
    RB38400: TRadioButton;
    RB57600: TRadioButton;
    RB115200: TRadioButton;
    GroupBox10: TGroupBox;
    RB5b: TRadioButton;
    RB6b: TRadioButton;
    RB7b: TRadioButton;
    RB8b: TRadioButton;
    GroupBox11: TGroupBox;
    GroupBox12: TGroupBox;
    RBSans: TRadioButton;
    RBPaire: TRadioButton;
    RBImpaire: TRadioButton;
    RB1Stop: TRadioButton;
    RB2Stop: TRadioButton;
    ECRec: TEdit;
    HLabel13: THLabel;
    HLabel14: THLabel;
    HLabel15: THLabel;
    HLabel16: THLabel;
    HLabel17: THLabel;
    HLabel18: THLabel;
    HLabel19: THLabel;
    HLabel20: THLabel;
    HLabel21: THLabel;
    HLabel22: THLabel;
    HLabel23: THLabel;
    HLabel24: THLabel;
    ECNum: TEdit;
    ECRaz: TEdit;
    ECAut: TEdit;
    ECIni: TEdit;
    EORec: TEdit;
    EONum: TEdit;
    EORaz: TEdit;
    EOAut: TEdit;
    EOIni: TEdit;
    EODet: TEdit;
    EERec: TEdit;
    EENum: TEdit;
    EERaz: TEdit;
    EEAut: TEdit;
    EEIni: TEdit;
    EEDet: TEdit;
    EDRec: THNumEdit;
    EDNum: THNumEdit;
    EDRaz: THNumEdit;
    EDAut: THNumEdit;
    EDIni: THNumEdit;
    EDDet: THNumEdit;
    GroupBox13: TGroupBox;
    GroupBox14: TGroupBox;
    GroupBox15: TGroupBox;
    HLabel25: THLabel;
    EExterieure: TEdit;
    HLabel26: THLabel;
    ETelAsync: TEdit;
    GroupBox16: TGroupBox;
    CBSync: TCheckBox;
    HLabel27: THLabel;
    ETelSync: TEdit;
    HLabel28: THLabel;
    EIdent: TEdit;
    GroupBox17: TGroupBox;
    CBJalConnect: TCheckBox;
    RBEvSuppr: TRadioButton;
    RBEvAncien: TRadioButton;
    ENbJours: THNumEdit;
    HLabel29: THLabel;
    BVoirJal: TToolbarButton97;
    GroupBox18: TGroupBox;
    CBTrace: TCheckBox;
    BActualiser: TToolbarButton97;
    BCherche: TToolbarButton97;
    LBTrace: TListBox;
    BNouveau: TToolbarButton97;
    BSuppr: TToolbarButton97;
    Msgs: THMsgBox;
    Timer1: TTimer;
    FindDialog: TFindDialog;
    HMTrad: THSystemMenu;
    procedure BConnectClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure PageControlChanging(Sender: TObject; var AllowChange: Boolean);
    procedure LBDestClick(Sender: TObject);
    procedure OnChangeCtl(Sender: TObject);
    procedure BNouveauClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BSupprClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure LBModeleDblClick(Sender: TObject);
    procedure LBDest2Click(Sender: TObject);
    procedure LBModeleClick(Sender: TObject);
    procedure CBSyncClick(Sender: TObject);
    procedure CBJalConnectClick(Sender: TObject);
    procedure RBEvSupprClick(Sender: TObject);
    procedure RBEvAncienClick(Sender: TObject);
    procedure BActualiserClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure BVoirJalClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    bConnexion       : Boolean;
    bCreating        : Boolean;
    bModifying       : Boolean;
    bWaiting         : Boolean;
    bCarteModifying  : Boolean;
    bMem             : Boolean;
    i_Com            : Integer;
    m_pszProfileName : string ;
    WMinX, WMinY     : Integer ;
    // Fonctions Message
    procedure WMGetMinMaxInfo(var Msg: TMessage); message WM_GETMINMAXINFO ;
    procedure InitVars;
    procedure EnableButtons(bOn: Boolean);
    procedure SetModeSaisie(bOn: boolean);
    procedure ZFile_GetProfilePath(FileName: string);
    procedure ZFile_GetDebugPath;
    function Analyse(s: string): boolean;
    procedure ReadKeyBanque;
    procedure WriteKeyBanque(KeyVal: string);
    function GetKeyBanque: integer;
    function GetKeyFromDataBanque(KeyVal: string): integer;
    procedure ReadDataBanque(KeyVal: string);
    procedure WriteDataBanque(KeyVal: string);
    procedure ReadDataModele;
    procedure WriteDataModele;
    procedure ReadCarte;
    procedure WriteCarte;
    procedure InitDataPort;
    procedure ReadDataPort;
    procedure WriteDataPort;
    procedure ReadDataHayes;
    procedure WriteDataHayes;
    procedure ReadDataTel;
    procedure WriteDataTel;
    procedure CreateSim(FileName : string) ;
    procedure WriteDataSim ;
    procedure ReadDataTrace;
    procedure WriteDataTrace;
  public
    constructor Create(AOwner: TComponent); override;
  end;

implementation

{$R *.DFM}

procedure ParamTeletrans ;
var X : TFMce;
    PP : THPanel ;
begin
PP:=FindInsidePanel ;
X:=TFMce.Create(nil);
if PP=Nil then
   BEGIN
    try
     X.ShowModal;
    finally
     X.Free;
    end;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end;

procedure TFMce.ZFile_GetProfilePath(FileName: string) ;
var szBuff: string ;
begin
  szBuff:=ExtractFilePath(Application.ExeName) ;
  m_pszProfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FileName, ETEBAC_EXT]) ;
end;

procedure TFMce.ZFile_GetDebugPath ;
var szBuff: string ;
begin
  szBuff:=ExtractFilePath(Application.ExeName) ;
  m_pszProfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, ETEBAC_DEBUG, ETEBAC_SYS]) ;
end;

//=======================================================
//================== Fonctions Message ==================
//=======================================================
procedure TFMce.WMGetMinMaxInfo(var Msg: TMessage) ;
begin
with PMinMaxInfo(Msg.lparam)^.ptMinTrackSize do begin X:=WMinX ; Y:=WMinY ; end ;
end;

procedure TFMce.InitVars;
begin
  bWaiting:=FALSE ;
  ECode.Text:='' ;
  EGuichet.Text:='' ;
  ENomDest.Text:='' ;
  EAdr.Text:='' ;
  EExt.Text:='' ;
  CPcv.Checked:=TRUE ;
  ENomContact.Text:='' ;
  ETelContact.Text:='' ;
  EFaxContact.Text:='' ;
  ENomTech.Text:='' ;
  ETelTech.Text:='' ;
  EFaxTech.Text:='' ;
  bWaiting:=TRUE ;
end;

procedure TFMce.ReadKeyBanque ;
var zfile : TIniFile ; Sections : TStringList ; ic: Integer ; cRead : string ;
begin
  LBDest.Clear;
  LBDest2.Clear;
  ZFile_GetProfilePath(FILE_BQ);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    Sections:=TStringList.Create;
    try
      zfile.ReadSection(SECTION_BQ, Sections);
      for ic:=0 to Sections.Count-1 do begin
        cRead:=zfile.ReadString(SECTION_BQ, Sections.Strings[ic], '');
        if cRead<>'' then begin
          LBDest.Items.Add(cRead);
          LBDest2.Items.Add(cRead);
        end;
      end;
    finally
      Sections.Free;
    end;
  finally
    zfile.Free;
  end;
end;

procedure TFMce.WriteKeyBanque(KeyVal: string);
var zfile : TIniFile ; cEntry : string ; ic : integer ;
begin
  ZFile_GetProfilePath(FILE_BQ);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    ic:=GetKeyBanque;
    if ic>=0 then begin
      cEntry:=Format('CLE%.3d', [ic]);
      zfile.WriteString(SECTION_BQ, cEntry, KeyVal);
    end;
  finally
    zfile.Free;
  end;
end;

function TFMce.GetKeyBanque: integer;
var
   zfile: TIniFile;
   cRead: string;
  cEntry: string;
      ic: integer;
     val: integer;
begin
  val:=-1;
  ZFile_GetProfilePath(FILE_BQ);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    for ic:=0 to MAX_ENTRY do begin
      cEntry:=Format('CLE%.3d', [ic]);
      cRead:=zfile.ReadString(SECTION_BQ, cEntry, '');
      if cRead='' then begin
        val:=ic;
        break;
      end;
    end;
  finally
    zfile.Free;
  end;
  Result:=val;
end;

function TFMce.GetKeyFromDataBanque(KeyVal: string): integer;
var
     zfile: TIniFile;
     cRead: string;
        ic: integer;
       val: integer;
  Sections: TStringList;
begin
  val:=-1;
  ZFile_GetProfilePath(FILE_BQ);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    Sections:=TStringList.Create;
    try
      zfile.ReadSection(SECTION_BQ, Sections);
      for ic:=0 to Sections.Count-1 do begin
        cRead:=zfile.ReadString(SECTION_BQ, Sections.Strings[ic], '');
        if cRead<>'' then begin
          if CompareStr(cRead, KeyVal)=0 then begin
            val:=StrToInt(Copy(Sections.Strings[ic], 4, 3));
            break;
          end;
        end;
      end;
    finally
      Sections.Free;
    end;
  finally
    zfile.Free;
  end;
  Result:=val;
end;

procedure TFMce.ReadDataBanque(KeyVal: string);
var
   zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_BQ);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    bWaiting:=FALSE;
    ECode.Text:=KeyVal;
        EAdr.Text:=zfile.ReadString(KeyVal, 'ADR',     '');
    EGuichet.Text:=zfile.ReadString(KeyVal, 'GUICHET', '');
    if zfile.ReadInteger(KeyVal, 'PCV', 1)=0 then CPcv.Checked:=FALSE
    else CPcv.Checked:=TRUE;
           EExt.Text:=zfile.ReadString(KeyVal, 'EXT',  '');
       ENomDest.Text:=zfile.ReadString(KeyVal, 'NOM',  '');
    ENomContact.Text:=zfile.ReadString(KeyVal, 'NOM1', '');
       ENomTech.Text:=zfile.ReadString(KeyVal, 'NOM2', '');
    EFaxContact.Text:=zfile.ReadString(KeyVal, 'FAX1', '');
       EFaxTech.Text:=zfile.ReadString(KeyVal, 'FAX2', '');
    ETelContact.Text:=zfile.ReadString(KeyVal, 'TEL1', '');
       ETelTech.Text:=zfile.ReadString(KeyVal, 'TEL2', '');
    bWaiting:=TRUE;
  finally
    zfile.Free;
  end;
end;

procedure TFMce.WriteDataBanque(KeyVal: string);
var
   zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_BQ);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    zfile.WriteString(KeyVal, 'ADR', EAdr.Text);
    zfile.WriteString(KeyVal, 'GUICHET', EGuichet.Text);
    if CPcv.Checked=TRUE then
      zfile.WriteInteger(KeyVal, 'PCV', 1)
    else
      zfile.WriteInteger(KeyVal, 'PCV', 0);
    zfile.WriteString(KeyVal, 'EXT',  EExt.Text);
    zfile.WriteString(KeyVal, 'NOM',  ENomDest.Text);
    zfile.WriteString(KeyVal, 'NOM1', ENomContact.Text);
    zfile.WriteString(KeyVal, 'NOM2', ENomTech.Text);
    zfile.WriteString(KeyVal, 'FAX1', EFaxContact.Text);
    zfile.WriteString(KeyVal, 'FAX2', EFaxTech.Text);
    zfile.WriteString(KeyVal, 'TEL1', ETelContact.Text);
    zfile.WriteString(KeyVal, 'TEL2', ETelTech.Text);
  finally
    zfile.Free;
  end;
end;

procedure TFMce.ReadDataModele;
var
  cEntry: string;
   cRead: string;
      ic: integer;
   zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_MODELE);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    for ic:=0 to MAX_ENTRY do begin
      cEntry:=Format('E%.3d', [ic]);
      cRead:=zfile.ReadString(SECTION_MODELE, cEntry, '');
      if cRead='' then break;
      cRead:=cEntry+' '+cRead;
      LBModele.Items.Add(cRead);
    end;
    for ic:=0 to MAX_ENTRY do begin
      cEntry:=Format('R%.3d', [ic]);
      cRead:=zfile.ReadString(SECTION_MODELE, cEntry, '');
      if cRead='' then break;
      cRead:=cEntry+' '+cRead;
      LBModele.Items.Add(cRead);
    end;
  finally
    zfile.Free;
  end;
end;

procedure TFMce.WriteDataModele;
var
   zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_MODELE);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    zfile.WriteString(SECTION_MODELE, 'E000', 'Transmission des LCR');
    zfile.WriteString(SECTION_MODELE, 'E001', 'Transmission des virements à effectuer');
    zfile.WriteString(SECTION_MODELE, 'E002', 'Transmission des prélèvements émis');
    zfile.WriteString(SECTION_MODELE, 'E003', 'Transmission des LCC et BO clients');
    zfile.WriteString(SECTION_MODELE, 'R000', 'Relevés de comptes en Francs Français');
    zfile.WriteString(SECTION_MODELE, 'R001', 'Relevés de comptes en devises');
    zfile.WriteString(SECTION_MODELE, 'R002', 'LCR/BO fournisseurs arrivant à échéance');
    zfile.WriteString(SECTION_MODELE, 'R003', 'Impayés sur remises de LCC/LCR clients');
    zfile.WriteString(SECTION_MODELE, 'R004', 'Impayés sur prélèvements émis par le client');
  finally
    zfile.Free;
  end;
end;

procedure TFMce.ReadCarte;
var
  CarteLue: string;
        Bq: string;
    Modele: string;
     zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_CARTE);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    ECarte.Text:='';
    if LBDest2.ItemIndex<>-1 then Bq:=LBDest2.Items.Strings[LBDest2.ItemIndex];
    if LBModele.ItemIndex<>-1 then Modele:=Copy(LBModele.Items.Strings[LBModele.ItemIndex], 1, 4);
    CarteLue:=zfile.ReadString(Bq, Modele, '');
    if CarteLue<>'' then ECarte.Text:=TrimRight(Copy(CarteLue, 2, 80));
  finally
    zfile.Free;
  end;
end;

procedure TFMce.WriteCarte;
var
  CarteEcrite: string;
     CarteTmp: string;
           Bq: string;
       Modele: string;
           ic: integer;
        zfile: TIniFile;
begin
  if ECarte.Text<>'' then begin
    CarteTmp:=ECarte.Text;
    if Length(ECarte.Text)<80 then begin
      for ic:=1 to 80-Length(ECarte.Text) do
        CarteTmp:=CarteTmp+' ';
    end;
    if LBDest2.ItemIndex<>-1 then Bq:=LBDest2.Items.Strings[LBDest2.ItemIndex];
    if LBModele.ItemIndex<>-1 then Modele:=Copy(LBModele.Items.Strings[LBModele.ItemIndex], 1, 4);
    CarteEcrite:=Format('@%s@', [CarteTmp]);
    ZFile_GetProfilePath(FILE_CARTE);
    zfile:=TIniFile.Create(m_pszProfileName);
    try
	zfile.WriteString(Bq, Modele, CarteEcrite);
    finally
      zfile.Free;
    end;
  end;
end;

procedure TFMce.InitDataPort;
begin
  RBCom1.Checked:=FALSE;  RBCom2.Checked:=FALSE;  RBCom3.Checked:=FALSE;
  RBCom4.Checked:=FALSE;  RBCom5.Checked:=FALSE;  RBCom6.Checked:=FALSE;
  RBCom7.Checked:=FALSE;  RBCom8.Checked:=FALSE;
  RB300.Checked:=FALSE;   RB9600.Checked:=FALSE;  RB19200.Checked:=FALSE;
  RB38400.Checked:=FALSE; RB57600.Checked:=FALSE; RB115200.Checked:=FALSE;
  RB5b.Checked:=FALSE;    RB6b.Checked:=FALSE;    RB7b.Checked:=FALSE;
  RB8b.Checked:=FALSE;
  RBSans.Checked:=FALSE;  RBPaire.Checked:=FALSE; RBImpaire.Checked:=FALSE;
  RB1Stop.Checked:=FALSE; RB2Stop.Checked:=FALSE;
end;

procedure TFMce.ReadDataPort;
var
   zfile: TIniFile;
  i_read: integer;
begin
  ZFile_GetProfilePath(FILE_MODEM);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    InitDataPort;
    i_read:=zfile.ReadInteger(SECTION_PORT, 'SERIE', 0);
    case i_read of
      0: RBCom1.Checked:=TRUE;
      1: RBCom2.Checked:=TRUE;
      2: RBCom3.Checked:=TRUE;
      3: RBCom4.Checked:=TRUE;
      4: RBCom5.Checked:=TRUE;
      5: RBCom6.Checked:=TRUE;
      6: RBCom7.Checked:=TRUE;
      7: RBCom8.Checked:=TRUE;
    end;
    i_read:=zfile.ReadInteger(SECTION_PORT, 'VITESSE', 2);
    case i_read of
      0: RB300.Checked:=TRUE;
      1: RB9600.Checked:=TRUE;
      2: RB19200.Checked:=TRUE;
      3: RB38400.Checked:=TRUE;
      4: RB57600.Checked:=TRUE;
      5: RB115200.Checked:=TRUE;
    end;
    i_read:=zfile.ReadInteger(SECTION_PORT, 'BITS', 3);
    case i_read of
      0: RB5b.Checked:=TRUE;
      1: RB6b.Checked:=TRUE;
      2: RB7b.Checked:=TRUE;
      3: RB8b.Checked:=TRUE;
    end;
    i_read:=zfile.ReadInteger(SECTION_PORT, 'PARITE', 0);
    case i_read of
      0: RBSans.Checked:=TRUE;
      1: RBPaire.Checked:=TRUE;
      2: RBImpaire.Checked:=TRUE;
    end;
    i_read:=zfile.ReadInteger(SECTION_PORT, 'STOP', 0);
    case i_read of
      0: RB1Stop.Checked:=TRUE;
      1: RB2Stop.Checked:=TRUE;
    end;
  finally
    zfile.Free;
  end;
end;

procedure TFMce.WriteDataPort;
var
    zfile: TIniFile;
  i_write: integer;
begin
  ZFile_GetProfilePath(FILE_MODEM);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    i_write:=0;
    if RBCom1.Checked then i_write:=0;
    if RBCom2.Checked then i_write:=1;
    if RBCom3.Checked then i_write:=2;
    if RBCom4.Checked then i_write:=3;
    if RBCom5.Checked then i_write:=4;
    if RBCom6.Checked then i_write:=5;
    if RBCom7.Checked then i_write:=6;
    if RBCom8.Checked then i_write:=7;
    zfile.WriteInteger(SECTION_PORT, 'SERIE', i_write);
    i_write:=2;
    if RB300.Checked    then i_write:=0;
    if RB9600.Checked   then i_write:=1;
    if RB19200.Checked  then i_write:=2;
    if RB38400.Checked  then i_write:=3;
    if RB57600.Checked  then i_write:=4;
    if RB115200.Checked then i_write:=5;
    zfile.WriteInteger(SECTION_PORT, 'VITESSE', i_write);
    i_write:=3;
    if RB5b.Checked then i_write:=0;
    if RB6b.Checked then i_write:=1;
    if RB7b.Checked then i_write:=2;
    if RB8b.Checked then i_write:=3;
    zfile.WriteInteger(SECTION_PORT, 'BITS', i_write);
    i_write:=0;
    if RBSans.Checked    then i_write:=0;
    if RBPaire.Checked   then i_write:=1;
    if RBImpaire.Checked then i_write:=2;
    zfile.WriteInteger(SECTION_PORT, 'PARITE', i_write);
    i_write:=0;
    if RB1Stop.Checked then i_write:=0;
    if RB2Stop.Checked then i_write:=1;
    zfile.WriteInteger(SECTION_PORT, 'STOP', i_write);
  finally
    zfile.Free;
  end;
end;

procedure TFMce.ReadDataHayes;
var
    zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_MODEM);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    ECRec.Text:=zfile.ReadString(SECTION_MODEM, 'CMDREC', 'ATI');
    ECNum.Text:=zfile.ReadString(SECTION_MODEM, 'CMDNUM', 'ATDT');
    ECRaz.Text:=zfile.ReadString(SECTION_MODEM, 'CMDRAZ', 'ATZ');
    ECAut.Text:=zfile.ReadString(SECTION_MODEM, 'CMDAUT', 'ATS0=2');
    ECIni.Text:=zfile.ReadString(SECTION_MODEM, 'CMDINI', 'ATE1V1');
    EORec.Text:=zfile.ReadString(SECTION_MODEM, 'VOKREC', 'OK');
    EONum.Text:=zfile.ReadString(SECTION_MODEM, 'VOKNUM', 'CONNECT');
    EORaz.Text:=zfile.ReadString(SECTION_MODEM, 'VOKRAZ', 'OK');
    EOAut.Text:=zfile.ReadString(SECTION_MODEM, 'VOKAUT', 'OK');
    EOIni.Text:=zfile.ReadString(SECTION_MODEM, 'VOKINI', 'OK');
    EODet.Text:=zfile.ReadString(SECTION_MODEM, 'VOKDET', 'CONNECT');
    EERec.Text:=zfile.ReadString(SECTION_MODEM, 'VECREC', 'ERR');
    EENum.Text:=zfile.ReadString(SECTION_MODEM, 'VECNUM', 'ERR');
    EERaz.Text:=zfile.ReadString(SECTION_MODEM, 'VECRAZ', 'ERR');
    EEAut.Text:=zfile.ReadString(SECTION_MODEM, 'VECAUT', 'ERR');
    EEIni.Text:=zfile.ReadString(SECTION_MODEM, 'VECINI', 'ERR');
    EEDet.Text:=zfile.ReadString(SECTION_MODEM, 'VECDET', 'CARRIER');
    EDRec.Value:=zfile.ReadInteger(SECTION_MODEM, 'DATREC', 100);
    EDNum.Value:=zfile.ReadInteger(SECTION_MODEM, 'DATNUM', 1200);
    EDRaz.Value:=zfile.ReadInteger(SECTION_MODEM, 'DATRAZ', 100);
    EDAut.Value:=zfile.ReadInteger(SECTION_MODEM, 'DATAUT', 100);
    EDIni.Value:=zfile.ReadInteger(SECTION_MODEM, 'DATINI', 100);
    EDDet.Value:=zfile.ReadInteger(SECTION_MODEM, 'DATDET', 500);
  finally
    zfile.Free;
  end;
end;

procedure TFMce.WriteDataHayes;
var
    zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_MODEM);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    zfile.WriteString(SECTION_MODEM, 'CMDREC', ECRec.Text);
    zfile.WriteString(SECTION_MODEM, 'CMDNUM', ECNum.Text);
    zfile.WriteString(SECTION_MODEM, 'CMDRAZ', ECRaz.Text);
    zfile.WriteString(SECTION_MODEM, 'CMDAUT', ECAut.Text);
    zfile.WriteString(SECTION_MODEM, 'CMDINI', ECIni.Text);
    zfile.WriteString(SECTION_MODEM, 'VOKREC', EORec.Text);
    zfile.WriteString(SECTION_MODEM, 'VOKNUM', EONum.Text);
    zfile.WriteString(SECTION_MODEM, 'VOKRAZ', EORaz.Text);
    zfile.WriteString(SECTION_MODEM, 'VOKAUT', EOAut.Text);
    zfile.WriteString(SECTION_MODEM, 'VOKINI', EOIni.Text);
    zfile.WriteString(SECTION_MODEM, 'VOKDET', EODet.Text);
    zfile.WriteString(SECTION_MODEM, 'VECREC', EERec.Text);
    zfile.WriteString(SECTION_MODEM, 'VECNUM', EENum.Text);
    zfile.WriteString(SECTION_MODEM, 'VECRAZ', EERaz.Text);
    zfile.WriteString(SECTION_MODEM, 'VECAUT', EEAut.Text);
    zfile.WriteString(SECTION_MODEM, 'VECINI', EEIni.Text);
    zfile.WriteString(SECTION_MODEM, 'VECDET', EEDet.Text);
    zfile.WriteInteger(SECTION_MODEM, 'DATREC', Round(EDRec.Value));
    zfile.WriteInteger(SECTION_MODEM, 'DATNUM', Round(EDNum.Value));
    zfile.WriteInteger(SECTION_MODEM, 'DATRAZ', Round(EDRaz.Value));
    zfile.WriteInteger(SECTION_MODEM, 'DATAUT', Round(EDAut.Value));
    zfile.WriteInteger(SECTION_MODEM, 'DATINI', Round(EDIni.Value));
    zfile.WriteInteger(SECTION_MODEM, 'DATDET', Round(EDDet.Value));
  finally
    zfile.Free;
  end;
end;

procedure TFMce.ReadDataTel ;
var zfile : TIniFile ;
begin
  zfile_GetProfilePath(FILE_MODEM) ;
  zfile:=TIniFile.Create(m_pszProfileName) ;
  try
    EExterieure.Text:=zfile.ReadString(SECTION_TEL, 'TELEXT', '') ;
    ETelAsync.Text:=zfile.ReadString(SECTION_TEL, 'TASYNC', ETEBAC_ASYNC) ;
    if zfile.ReadInteger(SECTION_TEL, 'SYN', 0)=1 then CBSync.Checked:=TRUE
                                                  else CBSync.Checked:=FALSE ;
    ETelSync.Text:=zfile.ReadString(SECTION_TEL, 'TSYNC', ETEBAC_SYNC) ;
    EIdent.Text:=zfile.ReadString(SECTION_TEL, 'NUI', '') ;
  finally
    zfile.Free ;
  end;
  CBSync.OnClick(nil);
end;

procedure TFMce.WriteDataTel ;
var zfile : TIniFile ;
begin
  zfile_GetProfilePath(FILE_MODEM) ;
  zfile:=TIniFile.Create(m_pszProfileName) ;
  try
    zfile.WriteString(SECTION_TEL, 'TELEXT', EExterieure.Text) ;
    if ETelAsync.Text='' then ETelAsync.Text:=ETEBAC_ASYNC ;
    zfile.WriteString(SECTION_TEL, 'TASYNC', ETelAsync.Text) ;
    if CBSync.Checked then zfile.WriteInteger(SECTION_TEL, 'SYN', 1)
                      else zfile.WriteInteger(SECTION_TEL, 'SYN', 0) ;
    if ETelSync.Text='' then ETelSync.Text:=ETEBAC_SYNC ;
    zfile.WriteString(SECTION_TEL, 'TSYNC', ETelSync.Text) ;
    zfile.WriteString(SECTION_TEL, 'NUI', EIdent.Text) ;
    if (Length(EIdent.Text)>0) then WriteDataSim ;
  finally
    zfile.Free ;
  end;
end;

procedure TFMce.CreateSim(FileName : string) ;
var Stream : TextFile ; Buffer : string ;
begin
AssignFile(Stream, FileName) ;
ReWrite(Stream) ;
Writeln(Stream, '[COMPBLOC=N]') ;
Writeln(Stream, '[SEPARAT=59]') ;
Writeln(Stream, '[I_ATTRIBUTS=O]') ;
Writeln(Stream, '[I_AFFGRAPH=O]') ;
Writeln(Stream, '[I_ECHOFRONTAL=O]') ;
Writeln(Stream, '[L_SUITE=Tab]') ;
Writeln(Stream, '[L_RETOUR=SftTab]') ;
Writeln(Stream, '[L_CORRECTION=BkSpce]') ;
Writeln(Stream, '[L_ANNULATION=Suppr]') ;
Writeln(Stream, '[L_ENVOI=Entree]') ;
Writeln(Stream, '[L_SOMMAIRE=Echapp]') ;
Writeln(Stream, '[L_ARRET=Ctrl A]') ;
Writeln(Stream, '[C_CHARIOT=9]') ;
Writeln(Stream, '[C_VERT=255]') ;
Writeln(Stream, '[C_ARRIERE=8]') ;
Writeln(Stream, '[C_HORIZONT=11]') ;
Writeln(Stream, '[C_LIGNE=13]') ;
Writeln(Stream, '[C_ABANDON=27]') ;
Writeln(Stream, '[C_SUITE=80]') ;
Writeln(Stream, '[C_RETOUR=72]') ;
Writeln(Stream, '[C_CORRECTION=255]') ;
Writeln(Stream, '[C_ANNULATION=83]') ;
Writeln(Stream, '[C_ENVOI=81]') ;
Writeln(Stream, '[C_SOMMAIRE=73]') ;
Writeln(Stream, '[C_SONNERIE=7]') ;
Writeln(Stream, '[C_ECHAPPE=0]') ;
Writeln(Stream, '[D_ATTENTE=0]') ;
Writeln(Stream, '[ET3_ACT=1]') ;
Writeln(Stream, '[PES_ACT=1]') ;
Writeln(Stream, '[XMO_ACT=1]') ;
Writeln(Stream, '[P17_ACT=1]') ;
Writeln(Stream, '[ATL_ACT=1]') ;
Writeln(Stream, '[TED_ACT=1]') ;
Writeln(Stream, '[REG_ACT=1]') ;
Writeln(Stream, '[OFT_ACT=1]') ;
Writeln(Stream, '[TRA_ACT=1]') ;
Writeln(Stream, '[P7_ACT=1]') ;
Writeln(Stream, '[FAX_ACT=1]') ;
Writeln(Stream, '[ET5_ACT=1]') ;
Writeln(Stream, '[DGI_ACT=1]') ;
Writeln(Stream, '[TEDM_ACT=1]') ;
Buffer:=Format('[NUI=%s]', [EIdent.Text]);
Writeln(Stream, Buffer);
CloseFile(Stream) ;
end ;

procedure TFMce.WriteDataSim ;
var szBuff, Buffer, FileName2 : string ; Stream, Stream2 : TextFile ;
begin
szBuff:=ExtractFilePath(Application.ExeName) ;
m_pszProfileName:=Format('%s%s%s%s', [szBuff, ETEBAC_PATH, FILE_SIM, SIM_EXT]) ;
if FileExists(m_pszProfileName) then
  begin
  AssignFile(Stream, m_pszProfileName) ;
  Reset(Stream) ;
  FileName2:=m_pszProfileName ; FileName2[Length(FileName2)-1]:='Z' ;
  AssignFile(Stream2, FileName2) ;
  ReWrite(Stream2) ;
  while not Eof(Stream) do
    begin
    Readln(Stream, Buffer) ;
    if Copy(Buffer, 1, 5)='[NUI=' then Buffer:=Format('[NUI=%s]', [EIdent.Text]) ;
    Writeln(Stream2, Buffer) ;
    end ;
  CloseFile(Stream) ; CloseFile(Stream2) ;
  DeleteFile(m_pszProfileName) ; RenameFile(FileName2, m_pszProfileName) ;
  end else CreateSim(m_pszProfileName) ;
end ;

procedure TFMce.ReadDataTrace;
var
    zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_MODEM);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    if zfile.ReadInteger(SECTION_TRACE, 'CONNECT', 1)=1 then CBJalConnect.Checked:=TRUE
    else CBJalConnect.Checked:=FALSE;
    if zfile.ReadInteger(SECTION_TRACE, 'EVDEL', 0)=0 then RBEvSuppr.Checked:=TRUE
    else RBEvSuppr.Checked:=FALSE;
    if RBEvSuppr.Checked then RBEvAncien.Checked:=FALSE else RBEvAncien.Checked:=TRUE;
    ENbJours.Value:=zfile.ReadInteger(SECTION_TRACE, 'JOURS', 7);
    if zfile.ReadInteger(SECTION_TRACE, 'DEBUG', 0)=1 then CBTrace.Checked:=TRUE
    else CBTrace.Checked:=FALSE;
  finally
    zfile.Free;
  end;
end;

procedure TFMce.WriteDataTrace;
var
    zfile: TIniFile;
begin
  ZFile_GetProfilePath(FILE_MODEM);
  zfile:=TIniFile.Create(m_pszProfileName);
  try
    if CBJalConnect.Checked then zfile.WriteInteger(SECTION_TRACE, 'CONNECT', 1)
    else zfile.WriteInteger(SECTION_TRACE, 'CONNECT', 0);
    if RBEvSuppr.Checked then zfile.WriteInteger(SECTION_TRACE, 'EVDEL', 0)
    else zfile.WriteInteger(SECTION_TRACE, 'EVDEL', 1);
    zfile.WriteInteger(SECTION_TRACE, 'JOURS', Round(ENbJours.Value));
    if CBTrace.Checked then zfile.WriteInteger(SECTION_TRACE, 'DEBUG', 1)
    else zfile.WriteInteger(SECTION_TRACE, 'DEBUG', 0);
  finally
    zfile.Free;
  end;
end;

constructor TFMce.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);	{ Initialise les composants reçus en héritage }
  bConnexion:=TRUE ; //FALSE;
  bWaiting:=TRUE;
  bCreating:=FALSE;
  bModifying:=FALSE;
  bCarteModifying:=FALSE;
  TabSheet3.TabVisible:=TRUE ; //FALSE;
  TabSheet4.TabVisible:=TRUE ; //FALSE;
  TabSheet5.TabVisible:=TRUE ; //FALSE;
  TabSheet6.TabVisible:=TRUE ; //FALSE;
end;

procedure TFMce.EnableButtons(bOn: Boolean);
var
  bVisible: Boolean;
begin
  bVisible:=FALSE;
  if bOn then bVisible:=TRUE
  else if LBDest.Items.Count>0 then bVisible:=TRUE;
  LBDest.Enabled:=bVisible;
  ECode.Enabled:=bCreating;
  EAdr.Enabled:=bVisible;
  EGuichet.Enabled:=bVisible;
  CPcv.Enabled:=bVisible;
  EExt.Enabled:=bVisible;
  ENomDest.Enabled:=bVisible;
  ENomContact.Enabled:=bVisible;
  ENomTech.Enabled:=bVisible;
  EFaxContact.Enabled:=bVisible;
  EFaxTech.Enabled:=bVisible;
  ETelContact.Enabled:=bVisible;
  ETelTech.Enabled:=bVisible;
  BConnect.Enabled:=bVisible ;
  BSuppr.Enabled:=bVisible ;
end;

procedure TFMce.SetModeSaisie(bOn: boolean);
begin
  if bOn then begin
    BConnect.Enabled:=FALSE;
    BNouveau.Enabled:=FALSE;
    BSuppr.Enabled:=FALSE;
  end
  else begin
    BConnect.Enabled:=TRUE;
    BNouveau.Enabled:=TRUE;
    BSuppr.Enabled:=TRUE;
  end;
end;

procedure TFMce.BConnectClick(Sender: TObject);
begin
  if bConnexion then
  begin
    TabSheet3.TabVisible:=FALSE;
    TabSheet4.TabVisible:=FALSE;
    TabSheet5.TabVisible:=FALSE;
    TabSheet6.TabVisible:=FALSE;
    BNouveau.Visible:=TRUE;
    BSuppr.Visible:=TRUE;
    TabSheet1.Show;
    TabSheet1.SetFocus;
  end
  else
  begin
    TabSheet3.TabVisible:=TRUE;
    TabSheet4.TabVisible:=TRUE;
    TabSheet5.TabVisible:=TRUE;
    TabSheet6.TabVisible:=TRUE;
    BNouveau.Visible:=FALSE;
    BSuppr.Visible:=FALSE;
    TabSheet3.Show;
    TabSheet3.SetFocus;
  end;
  bConnexion:=not bConnexion;
end;

procedure TFMce.FormCreate(Sender: TObject);
var sPath: string ;
begin
WMinX:=Width ; WMinY:=Height ;
sPath:=ExtractFilePath(Application.ExeName);
sPath:=sPath+ETEBAC_PATH;
{$I-}
MkDir(sPath);
{$I+}
if IOResult=0 then ;
end;

procedure TFMce.FormShow(Sender: TObject);
begin
  ReadKeyBanque;
  ReadDataModele;
  if LBModele.Items.Count=0 then begin
    WriteDataModele;
    ReadDataModele;
  end;
  LBDest.ItemIndex:=0;
  LBDest.OnClick(nil);
  LBDest2.ItemIndex:=0;
  LBModele.ItemIndex:=0;
  ReadCarte;
  ECarte.Enabled:=FALSE;
  EnableButtons(FALSE);
  ReadDataPort;
  ReadDataHayes;
  ReadDataTel;
  ReadDataTrace;
  TabSheet1.Show;
  TabSheet1.SetFocus;
end;

procedure TFMce.PageControlChange(Sender: TObject);
begin
  if PageControl.ActivePage=TabSheet1 then begin
    BNouveau.Visible:=TRUE;
    BSuppr.Visible:=TRUE;
  end
  else begin
    BNouveau.Visible:=FALSE;
    BSuppr.Visible:=FALSE;
  end;
end;

procedure TFMce.PageControlChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
if (bCreating) or (bModifying) or (bCarteModifying) or (LBDest.Items.Count=0)
  then AllowChange:=FALSE;
end;

procedure TFMce.LBDestClick(Sender: TObject);
begin
  if LBDest.ItemIndex<>-1 then
    ReadDataBanque(LBDest.Items.Strings[LBDest.ItemIndex])
  else
    InitVars;
end;

procedure TFMce.OnChangeCtl(Sender: TObject);
begin
  if bWaiting and not bCreating and not bModifying then begin
    bModifying:=TRUE;
    SetModeSaisie(TRUE);
  end;
end;

procedure TFMce.BNouveauClick(Sender: TObject);
begin
  bCreating:=TRUE;
  EnableButtons(TRUE);
  LBDest.Enabled:=FALSE;
  InitVars;
  SetModeSaisie(TRUE);
  ECode.SetFocus;
end;

procedure TFMce.BValiderClick(Sender: TObject);
var
  bErreur: boolean;
begin
if Screen.ActiveControl is THNumEdit then THNumEdit(Screen.ActiveControl).UpdateValue ;
  bErreur:=FALSE;
  if bCreating or bModifying then begin
    // Sauvegarde des données banque
    if ECode.Text='' then begin
      Msgs.Execute(0, '', '');
      ECode.SetFocus;
      bErreur:=TRUE;
    end
    else if bCreating then begin
      if LBDest.Items.IndexOf(ECode.Text)<>-1 then begin
        Msgs.Execute(1, '', '');
        ECode.SetFocus;
        bWaiting:=FALSE;
        ECode.Text:='';
        bWaiting:=TRUE;
        bErreur:=TRUE;
      end
      else begin
        WriteKeyBanque(ECode.Text);
        WriteDataBanque(ECode.Text);
        LBDest.ItemIndex:=LBDest.Items.Add(ECode.Text);
        LBDest2.ItemIndex:=LBDest2.Items.Add(ECode.Text);
      end;
    end;
    if not bErreur then begin
      if bModifying then WriteDataBanque(ECode.Text);
      SetModeSaisie(FALSE);
      LBDest.Enabled:=TRUE;
      bCreating:=FALSE;
      bModifying:=FALSE;
      EnableButtons(FALSE);
    end;
  end
  else if bCarteModifying then begin
    WriteCarte;
    ECarte.Enabled:=FALSE;
    LBDest2.Enabled:=TRUE;
    LBModele.Enabled:=TRUE;
    bCarteModifying:=FALSE;
    BConnect.Enabled:=TRUE;
  end
  else begin
    WriteDataPort;
    WriteDataHayes;
    WriteDataTel;
    WriteDataTrace;
    ModalResult:=mrOk;
  end;
end;

procedure TFMce.BAnnulerClick(Sender: TObject);
begin
  if bCreating or bModifying then begin
    SetModeSaisie(FALSE);
    LBDest.Enabled:=TRUE;
    if LBDest.ItemIndex<>-1 then ReadDataBanque(LBDest.Items.Strings[LBDest.ItemIndex]);
    bCreating:=FALSE;
    bModifying:=FALSE;
    EnableButtons(FALSE);
  end
  else if bCarteModifying then begin
    ReadCarte;
    ECarte.Enabled:=FALSE;
    LBDest2.Enabled:=TRUE;
    LBModele.Enabled:=TRUE;
    bCarteModifying:=FALSE;
    BConnect.Enabled:=TRUE;
  end
  else begin
    WriteDataPort;
    WriteDataHayes;
    WriteDataTel;
    WriteDataTrace;
    ModalResult:=mrCancel;
  end;
end;

procedure TFMce.BSupprClick(Sender: TObject);
var
   zfile: TIniFile;
  cEntry: string;
      ic: integer;
begin
  if LBDest.ItemIndex<>-1 then begin
    if Msgs.Execute(2, '', '')=mrYes then begin
      ZFile_GetProfilePath(FILE_BQ);
      zfile:=TIniFile.Create(m_pszProfileName);
      try
        zfile.EraseSection(LBDest.Items.Strings[LBDest.ItemIndex]);
        ic:=GetKeyFromDataBanque(LBDest.Items.Strings[LBDest.ItemIndex]);
        if ic>=0 then begin
          cEntry:=Format('CLE%.3d', [ic]);
          zfile.DeleteKey(SECTION_BQ, cEntry);
        end;
      finally
        zfile.Free;
      end;
      ZFile_GetProfilePath(FILE_CARTE);
      zfile:=TIniFile.Create(m_pszProfileName);
      try
        zfile.EraseSection(LBDest.Items.Strings[LBDest.ItemIndex]);
        LBDest2.Items.Delete(LBDest2.Items.IndexOf(LBDest.Items.Strings[LBDest.ItemIndex]));
        LBDest.Items.Delete(LBDest.ItemIndex);
        LBDest.ItemIndex:=0;
        LBDest2.ItemIndex:=0;
        EnableButtons(FALSE);
        LBDest.OnClick(nil);
      finally
        zfile.Free;
      end;
    end;
  end;
end;

procedure TFMce.Timer1Timer(Sender: TObject);
var
  iz: integer;
begin
  iz:=SendMessage(ECarte.Handle, EM_GETSEL, 0, 0);
  LPos.Caption:=Format('%d', [Lo(iz)+1]);
end;

procedure TFMce.LBModeleDblClick(Sender: TObject);
begin
  if LBDest2.ItemIndex>=0 then begin
    bCarteModifying:=TRUE;
    ECarte.Enabled:=TRUE;
    ECarte.SetFocus;
    ECarte.SelStart:=0; ECarte.SelLength:=0;
    LBDest2.Enabled:=FALSE;
    LBModele.Enabled:=FALSE;
    BConnect.Enabled:=FALSE;
  end
  else begin
    Msgs.Execute(6, '', '');
    BNouveau.Visible:=TRUE;
    BSuppr.Visible:=TRUE;
    TabSheet1.Show;
    TabSheet1.SetFocus;
  end;
end;

procedure TFMce.LBDest2Click(Sender: TObject);
begin
  ReadCarte;
end;

procedure TFMce.LBModeleClick(Sender: TObject);
begin
  ReadCarte;
end;

procedure TFMce.CBSyncClick(Sender: TObject);
begin
  if CBSync.Checked then begin
    ETelAsync.Enabled:=FALSE;
    ETelSync.Enabled:=TRUE;
  end
  else begin
    ETelAsync.Enabled:=TRUE;
    ETelSync.Enabled:=FALSE;
  end;
end;

procedure TFMce.CBJalConnectClick(Sender: TObject);
begin
  if CBJalConnect.Checked then begin
    RBEvSuppr.Enabled:=TRUE;
    RBEvAncien.Enabled:=TRUE;
    ENbJours.Enabled:=TRUE;
  end
  else begin
    RBEvSuppr.Enabled:=FALSE;
    RBEvAncien.Enabled:=FALSE;
    ENbJours.Enabled:=FALSE;
  end
end;

procedure TFMce.RBEvSupprClick(Sender: TObject);
begin
  if RBEvSuppr.Checked then ENbJours.Enabled:=FALSE;
end;

procedure TFMce.RBEvAncienClick(Sender: TObject);
begin
  if RBEvAncien.Checked then ENbJours.Enabled:=TRUE;
end;

function TFMce.Analyse(s: string): boolean;
var
  i_scan: integer;
    stmp: string;
     Aff: string;
begin
  if bMem and (Pos(WORD_RETURN, s)=0) then begin
    LBTrace.Items.Add(s);
    Result:=FALSE;
  end
  else if Pos(WORD_OUVRE, s)>0 then begin
    i_Com:=0 ;
    if (s[LEN_OUVRE+1] in ['0'..'9']) then i_Com:=StrToInt(s[LEN_OUVRE+1]) ;
    if i_Com>0 then begin
      Aff:=Format('Ouverture de COM%d', [i_Com]);
      LBTrace.Items.Add(Aff);
    end;
    Result:=FALSE;
  end
  else if Pos(WORD_FERME, s)>0 then begin
    if i_Com>0 then begin
      Aff:=Format('Fermeture de COM%d', [i_Com]);
      LBTrace.Items.Add(Aff);
    end;
    Result:=FALSE;
  end
  else if Pos(WORD_RETURN, s)>0 then begin
    i_scan:=LEN_RETURN+1; while s[i_scan] in ['0'..'9'] do begin stmp:=stmp+s[i_scan]; i_scan:=i_scan+1; end;
    i_scan:=StrToInt(stmp);
    Aff:=Format('Nombre de caractère(s) transféré(s) : %d', [i_scan]);
    LBTrace.Items.Add(Aff);
    bMem:=FALSE;
    Result:=FALSE;
  end
  else if Pos(WORD_MEM, s)>0 then begin
    bMem:=TRUE;
    Result:=FALSE;
  end
  // Elimination du "bruit"
  else if (Pos(WORD_INIT, s)>0)   or (Pos(WORD_LIRE, s)>0) or
          (Pos(WORD_ECRIRE, s)>0) or (Pos(WORD_COMERR, s)>0) then
    Result:=FALSE
  else begin
    LBTrace.Items.Add('¿PFU?: '+s);
    Result:=FALSE;
  end;
end;

procedure TFMce.BActualiserClick(Sender: TObject);
var
  Stream: TextFile;
   sRead: string;
begin
  bMem:=FALSE; i_Com:=0;
  LBTrace.Clear;
  ZFile_GetDebugPath;
  AssignFile(Stream, m_pszProfileName);
  {$I-}
  Reset(Stream);
  {$I+}
  if IOResult<>2 then
  begin
    while not Eof(Stream) do
    begin
      Readln(Stream, sRead);
      if Analyse(sRead) then LBTrace.Items.Add(sRead);
    end;
    CloseFile(Stream);
    LBTrace.ItemIndex:=0;
    LBTrace.Enabled:=TRUE;
    BCherche.Visible:=TRUE;
  end
  else begin
    LBTrace.Enabled:=FALSE;
    BCherche.Visible:=FALSE;
    Msgs.Execute(3, '', '');
  end;
end;

procedure TFMce.BChercheClick(Sender: TObject);
begin
  FindDialog.Execute;
end;

procedure TFMce.FindDialogFind(Sender: TObject);
var
      i: integer;
  bFind: boolean;
begin
  bFind:=FALSE;
  if frDown in TFindDialog(Sender).Options then begin
    for i:=LBTrace.ItemIndex+1 to LBTrace.Items.Count do begin
      if (i<0) or (i>=LBTrace.Items.Count) then break;
      if (StrPos(StrUpper(PChar(LBTrace.Items.Strings[i])),
                 StrUpper(PChar(FindDialog.FindText)))<>nil) then
      begin
        LBTrace.ItemIndex:=i;
        bFind:=TRUE;
        break;
      end;
    end;
  end
  else begin
    for i:=LBTrace.ItemIndex-1 downto 0 do begin
      if (i<0) or (i>=LBTrace.Items.Count) then break;
      if (StrPos(StrUpper(PChar(LBTrace.Items.Strings[i])),
          StrUpper(PChar(FindDialog.FindText)))<>nil) then
      begin
        LBTrace.ItemIndex:=i;
        bFind:=TRUE;
        break;
      end;
    end;
  end;
  if not bFind then Msgs.Execute(4, '', '');
end;

procedure TFMce.BVoirJalClick(Sender: TObject);
var
  si: TStartupInfo;
  pi: TProcessInformation;
  st: string;
begin
    ZeroMemory(@si, SizeOf(si) );
    si.cb:=SizeOf(si);    // Start the child process.
    st:=ExtractFilePath(Application.ExeName);
    st:='NOTEPAD.EXE '+st+ETEBAC_PATH+FILE_GEN+ETEBAC_LOG;
    if not CreateProcess(nil, // No module name (use command line).
        PChar(st),       // Command line.
        nil,             // Process handle not inheritable.
        nil,             // Thread handle not inheritable.
        FALSE,           // Set handle inheritance to FALSE.
        0,               // No creation flags.
        nil,             // Use parent's environment block.
        nil,             // Use parent's starting directory.
        si,              // Pointer to STARTUPINFO structure.
        pi)              // Pointer to PROCESS_INFORMATION structure.
    then Msgs.Execute(5, '', '');
end;

procedure TFMce.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then Action:=caFree ; 
end;

procedure TFMce.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end ;

end.
