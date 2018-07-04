unit ScenaCom;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, HSysMenu, hmsgbox, DB, DBTables, DBCtrls, StdCtrls, Buttons,
  ExtCtrls, Grids, DBGrids, HDB, Mask, Hctrls, ComCtrls, Ent1, HEnt1, Spin,
  Hqry, HTB97, HPanel, Hcompte, HRegCpte, FileCtrl, ADODB, udbxDataset ;

Procedure ParamSupImport(Import,Nature,Format,Code : String ; Comment : TActionFiche ;
                         QuellePage : Integer ; FromImp : Boolean=FALSE);

type
  TFParamImpSup = class(TFFicheListe)
    TNT_NATURE: THLabel;
    TNT_LIBELLE: THLabel;
    FS_LIBELLE: TDBEdit;
    FS_CODE: TDBEdit;
    Pages: TPageControl;
    Tinfos1: TTabSheet;
    TInfos2: TTabSheet;
    TaFS_IMPORT: TStringField;
    TaFS_NATURE: TStringField;
    TaFS_FORMAT: TStringField;
    TaFS_CODE: TStringField;
    TaFS_TRESO: TStringField;
    TaFS_GENATTEND: TStringField;
    TaFS_CLIATTEND: TStringField;
    TaFS_FOUATTEND: TStringField;
    TaFS_SALATTEND: TStringField;
    TaFS_DIVATTEND: TStringField;
    TaFS_SECTION1: TStringField;
    TaFS_SECTION2: TStringField;
    TaFS_SECTION3: TStringField;
    TaFS_SECTION4: TStringField;
    TaFS_SECTION5: TStringField;
    TaFS_COLLCLI: TStringField;
    TaFS_CPTCOLLCLI: TStringField;
    TaFS_COLFOU: TStringField;
    TaFS_CPTCOLLFOU: TStringField;
    TaFS_CORRIMPG: TStringField;
    TaFS_CORRIMPT: TStringField;
    TaFS_CORRIMP1: TStringField;
    TaFS_CORRIMP2: TStringField;
    TaFS_CORRIMP3: TStringField;
    TaFS_CORRIMP4: TStringField;
    TaFS_CORRIMP5: TStringField;
    TaFS_MPDEFAUT: TStringField;
    TaFS_MRDEFAUT: TStringField;
    FS_TRESO: TDBCheckBox;
    Label1: THLabel;
    FS_GENATTEND: THDBCpteEdit;
    Label2: THLabel;
    FS_CLIATTEND: THDBCpteEdit;
    Label3: THLabel;
    FS_FOUATTEND: THDBCpteEdit;
    Label4: THLabel;
    FS_SALATTEND: THDBCpteEdit;
    Label5: THLabel;
    FS_DIVATTEND: THDBCpteEdit;
    Label6: THLabel;
    FS_SECTION1: THDBCpteEdit;
    TFS_SECTION2: THLabel;
    FS_SECTION2: THDBCpteEdit;
    TFS_SECTION3: THLabel;
    FS_SECTION3: THDBCpteEdit;
    TFS_SECTION4: THLabel;
    FS_SECTION4: THDBCpteEdit;
    TFS_SECTION5: THLabel;
    FS_SECTION5: THDBCpteEdit;
    FS_MPDEFAUT: THDBValComboBox;
    Label11: THLabel;
    TabSheet1: TTabSheet;
    FS_CORRIMPG: TDBCheckBox;
    FS_CORRIMPT: TDBCheckBox;
    FS_CORRIMP1: TDBCheckBox;
    FS_CORRIMP2: TDBCheckBox;
    FS_CORRIMP3: TDBCheckBox;
    FS_CORRIMP4: TDBCheckBox;
    FS_CORRIMP5: TDBCheckBox;
    TabSheet2: TTabSheet;
    TaFS_CHEMIN: TStringField;
    TaFS_PREFIXE: TStringField;
    TaFS_SUFFIXE: TStringField;
    TaFS_NOMINTER: TStringField;
    TaFS_NOMREJET: TStringField;
    Label14: THLabel;
    FS_CHEMIN: TDBEdit;
    Label15: THLabel;
    FS_PREFIXE: TDBEdit;
    Label16: THLabel;
    FS_SUFFIXE: TDBEdit;
    Label17: THLabel;
    FS_NOMINTER: TDBEdit;
    Label18: THLabel;
    FS_NOMREJET: TDBEdit;
    TaFS_TREGDEFAUT: TStringField;
    TaFS_TMRDEFAUT: TStringField;
    TaFS_FILTREGENE: TStringField;
    TaFS_NATUREGENE: TStringField;
    TaFS_FILTREAUX: TStringField;
    TaFS_NATUREAUX: TStringField;
    TaFS_NATMVT: TStringField;
    TaFS_TRAITETVA: TStringField;
    TaFS_TRAITECTR: TStringField;
    TaFS_LIBMP: TStringField;
    FS_TRAITETVA: TDBCheckBox;
    FS_TRAITECTR: TDBCheckBox;
    HLabel1: THLabel;
    FS_FORMAT: THDBValComboBox;
    TaFS_LIBELLE: TStringField;
    TaFS_RIBCLIENT: TStringField;
    FS_RIBCLIENT: TDBCheckBox;
    TaFS_DOUBLON: TStringField;
    TaFS_FORCEPIECE: TStringField;
    TaFS_ANOUVEAUD: TStringField;
    FS_ANOUVEAUD: TDBCheckBox;
    RechFile: TToolbarButton97;
    FS_DOUBLON: TDBCheckBox;
    FS_FORCEPIECE: TDBCheckBox;
    FS_FILTREGENE: TDBCheckBox;
    TabSheet3: TTabSheet;
    FS_COLLCLI: TDBCheckBox;
    Label12: THLabel;
    FS_CPTCOLLCLI: THDBCpteEdit;
    FS_COLFOU: TDBCheckBox;
    Label13: THLabel;
    FS_CPTCOLLFOU: THDBCpteEdit;
    TaFS_RIBFOUR: TStringField;
    TaFS_VALIDEECR: TStringField;
    TaFS_DOUBLONM: TStringField;
    TaFS_FORCEPIECEM: TStringField;
    TaFS_CORRIMPMP: TStringField;
    FS_RIBFOUR: TDBCheckBox;
    FS_CORRIMPMP: TDBCheckBox;
    FS_VALIDEECR: TDBCheckBox;
    TaFS_CORRIMPJAL: TStringField;
    FS_CORRIMPJAL: TDBCheckBox;
    GroupBox1: TGroupBox;
    Label8: THLabel;
    FS_MRDEFAUT: THDBValComboBox;
    Label9: THLabel;
    FS_TREGDEFAUT: THDBValComboBox;
    TaFS_TAUXDEVOUT: TStringField;
    FS_TAUXDEVOUT: TDBCheckBox;
    FS_FILTREAUX: TDBCheckBox;
    TaFS_ECCDIV: TStringField;
    FS_ECCDIV: TDBCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure RechFileClick(Sender: TObject);
    procedure FS_DOUBLONClick(Sender: TObject);
    procedure FS_MPDEFAUTKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FS_MRDEFAUTKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FS_TREGDEFAUTKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Déclarations privées }
    Import,Nature,Format,Code     : String ;
    FermeCa,FromImp : Boolean ;
    Procedure Sors ;
  public
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

// Substitution des caractères " " par des "." dans comptes tiers
Procedure ParamSupImport(Import,Nature,Format,Code : String ; Comment : TActionFiche ;
                         QuellePage : Integer ; FromImp : Boolean=FALSE);
Var FParamImpSup : TFParamImpSup ;
    NomPrint : string ;
BEGIN
NomPrint:='PRT_SCIMP' ;
FParamImpSup:=TFParamImpSup.Create(Application) ;
 try
   FParamImpSup.Import:=Import ;
   FParamImpSup.Nature:=Nature ;
   FParamImpSup.Format:=Format ;
   FParamImpSup.Code:=Code ;
   FParamImpSup.FromImp:=FromImp ;
   FParamImpSup.InitFL('FS',NomPrint,'','X',Comment,TRUE,FParamImpSup.TaFS_CODE,FParamImpSup.TaFS_LIBELLE,Nil,['']) ;
   FParamImpSup.ShowModal ;
 Finally
   FParamImpSup.free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END;

procedure TFParamImpSup.FormShow(Sender: TObject);
begin
  inherited;
If Import='X' Then Caption:=HM2.Mess[0] Else Caption:=HM2.Mess[1] ;
Pages.ActivePage:=Pages.Pages[0] ;
// JLD5Axes
if EstSerie(S3) then
   BEGIN
   FS_CORRIMP3.Visible:=False ; FS_CORRIMP4.Visible:=False ; FS_CORRIMP5.Visible:=False ;
   FS_SECTION3.Visible:=False ; FS_SECTION4.Visible:=False ; FS_SECTION5.Visible:=False ;
   TFS_SECTION3.Visible:=False ; TFS_SECTION4.Visible:=False ; TFS_SECTION5.Visible:=False ;
   END ;
if EstSerie(S3) then
   BEGIN
   FS_CORRIMP2.Visible:=False ; FS_SECTION2.Visible:=False ; TFS_SECTION2.Visible:=False ;
   END ;
FermeCa:=False ;

if FromImp then
   BEGIN
   if TA.State=dsBrowse then
      begin
      //TRIB.Locate('R_NUMEROCOMPTE',NumCompte,[]) ;
      while not TA.EOF do
        if (TA.FindField('FS_FORMAT').AsString=Format) And
           ((TA.FindField('FS_CODE').AsString=Code) Or (Code='')) And
           (TA.FindField('FS_IMPORT').AsString=Import) And
           (TA.FindField('FS_NATURE').AsString=Nature) then Break else TA.Next ;
      end ;
   END Else
   BEGIN
   if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
     BEGIN
     if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
     END ;
   END ;
//   if FTypeAction in [taCreat,taCreatEnSerie,taCreatOne] then BinsertClick(Nil) ;
end;

Function TFParamImpSup.EnregOK : boolean ;
BEGIN
result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ;
if ((result) and (Ta.state in [dsInsert])) then
   BEGIN
   Result:=False ;
   if PresenceComplexe('FMTSUP',['FS_IMPORT','FS_NATURE','FS_FORMAT','FS_CODE'],['=','=','=','='],[Import,Nature,Format,FS_CODE.Text],['S','S','S','S']) then
      BEGIN HM.Execute(4,'','') ; if FS_CODE.CanFocus then FS_CODE.SetFocus ; Exit ; END ;
   END ;
Result:=TRUE  ; Modifier:=False ;
END ;

Procedure TFParamImpSup.NewEnreg ;
BEGIN
Inherited ;
TaFS_IMPORT.AsString:=Import ;
TaFS_NATURE.AsString:=Nature ;
TaFS_FORMAT.AsString:=Format ;
FS_FORMAT.Value:=Format ; FS_FORMAT.Refresh ;
FS_FORCEPIECE.Enabled:=FALSE ;
END ;

Procedure TFParamImpSup.ChargeEnreg  ;
BEGIN
Inherited ;
FS_FORCEPIECE.ENABLED:=FS_DOUBLON.Checked ;
END ;

procedure TFParamImpSup.BCAbandonClick(Sender: TObject);
begin
  inherited;
Sors ;
end;

Procedure TFParamImpSup.Sors ;
BEGIN
FermeCa:=True ; Close ;
END ;

procedure TFParamImpSup.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if FermeCa then CanClose:=True Else inherited;
end;

procedure TFParamImpSup.RechFileClick(Sender: TObject);
var st : AnsiString ;
    i : integer ;
    Options: TSelectDirOpts ;
begin
Inherited ;
If directoryExists(FS_CHEMIN.Text) then St:=FS_CHEMIN.Text else st:='c:\' ;
i:=1 ; Options:=[sdAllowCreate,sdPerformCreate] ;
if selectdirectory(st,Options,i) then
  BEGIN
  if Ta.State=dsBrowse then Ta.Edit ;
  TaFS_CHEMIN.Text:=St ;
  END ;
end;

procedure TFParamImpSup.FS_DOUBLONClick(Sender: TObject);
begin
  inherited;
If Not FS_DOUBLON.Checked Then
  BEGIN
  FS_FORCEPIECE.Checked:=FALSE ;
  FS_FORCEPIECE.Enabled:=FALSE ;
  END Else FS_FORCEPIECE.Enabled:=TRUE ;
end;

procedure TFParamImpSup.FS_MPDEFAUTKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
  inherited;
if (Key=VK_DELETE) then BEGIN FS_MPDEFAUT.ItemIndex:=-1 ; FS_MPDEFAUT.Value:='' ; Key:=0 ; Exit ; END ;
end;

procedure TFParamImpSup.FS_MRDEFAUTKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
  inherited;
if (Key=VK_DELETE) then BEGIN FS_MRDEFAUT.ItemIndex:=-1 ; FS_MRDEFAUT.Value:='' ; Key:=0 ; Exit ; END ;
end;

procedure TFParamImpSup.FS_TREGDEFAUTKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
if (Key=VK_DELETE) then BEGIN FS_TREGDEFAUT.ItemIndex:=-1 ; FS_TREGDEFAUT.Value:='' ; Key:=0 ; Exit ; END ;
end;

end.
