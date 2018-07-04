unit AssImp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, Menus, Grids, DBGrids, HDB, StdCtrls, Buttons, Hctrls,
  ComCtrls, hmsgbox, ExtCtrls,Ent1, DB, Hqry, FmtChoix,HEnt1,Paramsoc,
  HSysMenu, Hcompte, ImpFic, HStatus, HTB97, TImpFic, ed_tools,RapSuppr,SISCO,
  Mask, HPanel, ADODB, udbxDataset ;

procedure LanceImport(Lequel : String) ;
procedure LanceImportMvtExt(Lequel : String) ;
procedure LanceImportRecupSISCO(Lequel : String ; StFichier : String ; LettrageSISCOEnImport : Boolean = FALSE ; BigVol : Boolean = FALSE) ;
procedure LanceImportRecupSERVANT(Lequel : String ; StFichier : String ; ShuntFichier : Boolean) ;
procedure LanceImportSU(Lequel : String ; StFichier : String ; ShuntFichier : Boolean) ;

type
  TFAssistImport = class(TFAssist)
    Fichier: TTabSheet;
    HLabel7: THLabel;
    FFormat: THValComboBox;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    QUALIFPIECE: THValComboBox;
    H_QUALIFPIECE: THLabel;
    HLabel2: THLabel;
    CBPositifs: TCheckBox;
    LgComptes: TCheckBox;
    Controle: TTabSheet;
    Integration: TTabSheet;
    EBFormat: TToolbarButton97;
    HTitre: THLabel;
    StepMenu: TPopupMenu;
    FinEtape: TMenuItem;
    N1: TMenuItem;
    bStep: TToolbarButton97;
    ip1: TToolbarButton97;
    ip2: TToolbarButton97;
    Sauve: TSaveDialog;
    CBDoublons: TCheckBox;
    EBControle: TToolbarButton97;
    GroupBox3: TGroupBox;
    HLabel5: THLabel;
    HLabel6: THLabel;
    GroupBox4: TGroupBox;
    CBLettrage: TCheckBox;
    EBImporter: TToolbarButton97;
    Panel1: TPanel;
    HM: THMsgBox;
    NatP: THLabel;
    JalP: THLabel;
    DateP: THLabel;
    HLabel3: THLabel;
    HLabel1: THLabel;
    ZTotDeb: THNumEdit;
    ZTotCred: THNumEdit;
    LNbPiece: THLabel;
    PieceP: THLabel;
    Slash: THLabel;
    LigneP: THLabel;
    EBIntegre: TToolbarButton97;
    Bevel1: TBevel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    HLabel12: THLabel;
    THSlash: THLabel;
    THLigne: THLabel;
    TSociete: THTable;
    GroupBox2: TGroupBox;
    PListe: TPanel;
    ECaption: TLabel;
    TJAL1: THLabel;
    TJAL2: THLabel;
    Bevel5: TBevel;
    Label4: TLabel;
    PFenGuide: TPanel;
    Panel2: TPanel;
    BCValide: TBitBtn;
    BCAide: TBitBtn;
    BCAbandon: TBitBtn;
    JALAN: THValComboBox;
    JALOD: THValComboBox;
    GC: THGrid;
    HCpte: THCpteEdit;
    Bevel2: TBevel;
    Bevel3: TBevel;
    HLabel17: THLabel;
    EBStop: TToolbarButton97;
    BSTOP: TToolbarButton97;
    HTitres: THMsgBox;
    ZNb: THLabel;
    ZNbLig: THLabel;
    JAL1: TEdit;
    JAL2: TEdit;
    Par1: THLabel;
    Par2: THLabel;
    BVoir: TButton;
    FChangeMethode: TCheckBox;
    FReqParam: TCheckBox;
    Panel3: TPanel;
    Bevel4: TBevel;
    Bevel6: TBevel;
    Bevel7: TBevel;
    HLabel14: THLabel;
    HLabel15: THLabel;
    HLabel16: THLabel;
    ZNbGenLu: THLabel;
    HLabel19: THLabel;
    ZNbGenLuFaux: THLabel;
    HLabel24: THLabel;
    HLabel25: THLabel;
    HLabel26: THLabel;
    HLabel27: THLabel;
    ZNbPieceLu: THLabel;
    ZNbLigneLu: THLabel;
    ZTotDLu: THNumEdit;
    ZTotCLu: THNumEdit;
    HLabel4: THLabel;
    HLabel11: THLabel;
    HLabel13: THLabel;
    HLabel21: THLabel;
    ZNbAuxLu: THLabel;
    HLabel23: THLabel;
    ZNbAuxLuFaux: THLabel;
    ZNbAnaLu: THLabel;
    HLabel30: THLabel;
    ZNbAnaLufaux: THLabel;
    ZNbJalLu: THLabel;
    HLabel33: THLabel;
    ZNbJalLuFaux: THLabel;
    VisuImpEcr: TToolbarButton97;
    CBForceNumPiece: TCheckBox;
    HLabel18: THLabel;
    ZNbDevLu: THLabel;
    HLabel22: THLabel;
    ZNbDevLuFaux: THLabel;
    HLabel20: THLabel;
    ZNbEtabLu: THLabel;
    HLabel29: THLabel;
    ZNbEtabLuFaux: THLabel;
    EBListeErreur: TToolbarButton97;
    EBListeCpt: TToolbarButton97;
    EBListeDoublon: TToolbarButton97;
    Status: THStatusBar;
    CBSn2ORLI: TCheckBox;
    TSCENARIOIMPORT: THLabel;
    SCENARIOIMPORT: THCritMaskEdit;
    TFORCEDEVISE: THLabel;
    FORCEDEVISE: THValComboBox;
    TSuffixeBigVol: THLabel;
    SuffixeBigVol: TEdit;
    FIgnorel1: TCheckBox;
    function  FirstPage : TTabSheet ; Override ;
    procedure PChange(Sender: TObject) ; Override ;
    procedure FinEtapeClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure EBFormatClick(Sender: TObject);
    procedure RechFileClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure StepMenuClick(Sender: TObject);
    procedure EBControleClick(Sender: TObject);
    procedure EBIntegreClick(Sender: TObject);
    procedure EBImporterClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BCValideClick(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure EBStopClick(Sender: TObject);
    procedure FFormatChange(Sender: TObject);
    procedure FileNameChange(Sender: TObject);
    procedure GCDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BVoirClick(Sender: TObject);
    procedure FChangeMethodeClick(Sender: TObject);
    procedure FReqParamClick(Sender: TObject);
    procedure VisuImpEcrClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EBListeCptClick(Sender: TObject);
    procedure EBListeErreurClick(Sender: TObject);
    procedure EBListeDoublonClick(Sender: TObject);
    procedure CBDoublonsClick(Sender: TObject);
    procedure SCENARIOIMPORTKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SCENARIOIMPORTExit(Sender: TObject);
   private
    Lequel : String ;
    Steps  : Array[1..20] of Boolean ;
    OkSauve : boolean ;
    QuelleEtape : integer ;
    NbEtapes : Integer ;
    OldReferences : TStrings ;
    FirstFind : boolean ;
    Prefixe : String3 ;
    ReprendreleFichier : Boolean ;
    InfoImp : TInfoImport ;
    FicSISCOConverti : Boolean ;
    PremierControle : Boolean ;
    IntegrationFaite : Boolean ;
    SemiAuto : Boolean ;
    StFichier : String ;
    ShuntFichier : Boolean ;
    LettrageSISCOEnImport,BigVol : Boolean ;
    Procedure AlimInfoImp ;
    procedure AfficheMessControle(Var InfoImp : TInfoImport) ;
    function  ImportTexte(Initialise : boolean) : boolean ;
    Procedure AffAide(N : Integer) ;
    procedure ChargettQualP ;
    procedure ChargeTitres ;
    procedure ChargeEtapes ;
    procedure AddMenuPops ;
    procedure SauveEtapes(Fic : Boolean) ;
    procedure ChangeStepBitmap ;
    function  VerifOk(Var DesErreurs : Boolean) : boolean ;
    function  EffectuerLesEtapes(Leq : integer ; Fin : boolean) : boolean;
    procedure CaptionsDesMsgBox ;
//    procedure ModifRefPointage(Sender : TObject) ;
    function  OkRefPointageSaisis : boolean ;
    procedure DelOldRef ;
  public
  end;


implementation

uses ImpUtil,ParamDBG,VerCPTA,ChoixRef, mulimpec,ImpFicU,RappType, ScenaCom ;

{$R *.DFM}

procedure LanceImport(Lequel : String) ;
Var X : TFAssistImport ;
BEGIN
X:=TFAssistImport.Create(Application) ;
 Try
 X.SemiAuto:=FALSE ;
 X.Lequel:=Lequel ;
 X.StFichier:='' ;
 X.ShuntFichier:=FALSE ;
 X.LettrageSISCOEnImport:=FALSE ;
 X.BigVol:=FALSE ;
{$IFDEF JACADI}
 X.ShuntFichier:=TRUE ;
{$ENDIF}
 X.ShowModal ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
END ;

procedure LanceImportMvtExt(Lequel : String) ;
Var X : TFAssistImport ;
BEGIN

X:=TFAssistImport.Create(Application) ;
 Try
 X.SemiAuto:=TRUE ;
 X.Lequel:=Lequel ;
 X.StFichier:='' ;
 X.ShuntFichier:=FALSE ;
 X.LettrageSISCOEnImport:=FALSE ;
 X.BigVol:=FALSE ;
 X.ShowModal ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
END ;

procedure LanceImportRecupSISCO(Lequel : String ; StFichier : String ; LettrageSISCOEnImport : Boolean = FALSE ; BigVol : Boolean = FALSE) ;
Var X : TFAssistImport ;
BEGIN
X:=TFAssistImport.Create(Application) ;
 Try
 X.SemiAuto:=FALSE ;
 X.Lequel:=Lequel ;
 X.StFichier:=StFichier ;
 X.ShuntFichier:=FALSE ;
 X.LettrageSISCOEnImport:=LettrageSISCOEnImport ;
 X.BigVol:=BigVol ;
 X.ShowModal ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
END ;

procedure LanceImportRecupSERVANT(Lequel : String ; StFichier : String ; ShuntFichier : Boolean) ;
Var X : TFAssistImport ;
BEGIN
X:=TFAssistImport.Create(Application) ;
 Try
 X.SemiAuto:=FALSE ;
 X.Lequel:=Lequel ;
 X.StFichier:=StFichier ;
 X.ShuntFichier:=ShuntFichier ;
 X.LettrageSISCOEnImport:=FALSE ;
 X.BigVol:=FALSE ;
 X.ShowModal ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
END ;

procedure LanceImportSU(Lequel : String ; StFichier : String ; ShuntFichier : Boolean) ;
Var X : TFAssistImport ;
BEGIN
X:=TFAssistImport.Create(Application) ;
 Try
 X.SemiAuto:=FALSE ;
 X.Lequel:=Lequel ;
 X.StFichier:=StFichier ;
 X.ShuntFichier:=ShuntFichier ;
 X.ShowModal ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
END ;


Procedure TFAssistImport.AlimInfoImp ;
Var SuffSup : String ;
BEGIN
SuffSup:='' ;
InfoImp.NomFic:=FileName.Text ; InfoImp.Format:=FFormat.Value ; InfoImp.FormatOrigine:='' ; InfoImp.NomFicOrigine:='' ;
InfoImp.ForceDevise:=ForceDevise.Value ;
If (InfoImp.ForceDevise<>'FRF') And (InfoImp.ForceDevise<>'EUR') then InfoImp.ForceDevise:='' ;
If InfoImp.Format='SIS' Then
   BEGIN
   If RecupSISCO Then
     BEGIN
     If (VH^.RecupLTL Or VH^.RecupSISCOPGI)  And BigVol Then SuffSup:=SuffixeBigVol.text ;
     InfoImp.NomFic:=NomFichierRecupSISCO(InfoImp.NomFic,SuffSup) ;
     InfoImp.LettrageSISCOEnImport:=LettrageSISCOEnImport ;
     END Else
     BEGIN
     InfoImp.FormatOrigine:='SIS' ; InfoImp.NomFicOrigine:=InfoImp.NomFic ;
     InfoImp.NomFic:=NewNomFic(InfoImp.NomFic,'CGN') ;
     END ;
   END ;
If ImportMvtSU Then
  BEGIN
  InfoImp.NomFicOrigine:=InfoImp.NomFic ;
  InfoImp.NomFic:=NomFichierImportSU(InfoImp.NomFic)
  END ;
InfoImp.ForceQualif:=QualifPiece.Value ;
InfoImp.Lequel:=Lequel ; InfoImp.ForcePositif:=CBPositifs.Checked ; InfoImp.ForceBourrage:=TRUE ;
InfoImp.CtrlDB:=CBDoublons.Checked ;
InfoImp.ForceNumPiece:=CBForceNumPiece.Checked ;
(*
InfoImp.NomFicRejet:=LePath+LeNom+'Err'+LExtension ;
InfoImp.NomFicDoublon:=LePath+LeNom+'Dbl'+LExtension ;
*)
InfoImp.NomFicRejet:=NewNomFicEtDir(InfoImp.NomFic,'Err','Rejets') ;
InfoImp.NomFicDoublon:=NewNomFicEtDir(InfoImp.NomFic,'Dbl','Doublons') ;
InfoImp.NomFicRapport:=NewNomFicEtDir(InfoImp.NomFic,'Rap','Rapports') ;
InfoImp.CodeFormat:=ScenarioImport.Text ;
Sn2Orli:=FALSE ; If (InfoImp.Format='SN2') And CBSn2Orli.Checked Then Sn2Orli:=TRUE ;
ChargeScenarioImport(InfoImp,FALSE) ;
CreateFicRapport(InfoImp) ;
If InfoImp.Format='SIS' Then
   BEGIN
   If RecupSISCO Then InfoImp.Format:='CGE' Else InfoImp.Format:='CGN' ;
   END ;
If ImportMvtSU Then InfoImp.Format:='CGE' ;
END ;

function TFAssistImport.FirstPage : TTabSheet  ;
var i : Integer ;
BEGIN
Result:=inherited FirstPage ;
// positionnement sur la première étape non terminée
i := 1 ; while Steps[i] do Inc(i) ;
if (i>1) and (i<=P.PageCount-1) then Result := TTabSheet(P.Pages[i-1]) ;
END ;

procedure TFAssistImport.PChange(Sender: TObject);
begin
  inherited;
FinEtape.Checked := Steps[P.ActivePage.PageIndex+1] ;
ChangeStepBitmap ;
end;

procedure TFAssistImport.ChangeStepBitmap ;
begin
if Steps[P.ActivePage.PageIndex + 1] then bStep.Glyph := ip2.Glyph
                                     else bStep.Glyph := ip1.Glyph ;
end ;

procedure TFAssistImport.FinEtapeClick(Sender: TObject);
begin
FinEtape.Checked := not FinEtape.Checked ;
Steps[P.ActivePage.PageIndex + 1] := FinEtape.Checked ;
ChangeStepBitmap ;
end;

function TFAssistImport.OkRefPointageSaisis : boolean ;
var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT DISTINCT IE_GENERAL,IE_REFPOINTAGE FROM IMPECR LEFT JOIN GENERAUX ON '
          +'IE_GENERAL=G_GENERAL WHERE G_POINTABLE="X" AND IE_REFPOINTAGE=""',True) ;
Result:=Q.Eof ;
Ferme(Q) ;
END ;

function AucunMvtControlesOk : boolean ;
var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT DISTINCT IE_OKCONTROLE FROM IMPECR WHERE IE_OKCONTROLE="X"',True) ;
Result:=Q.Eof ;
Ferme(Q) ;
END ;

procedure TFAssistImport.bSuivantClick(Sender: TObject);
Var DesErreurs : Boolean ;
begin
	// Pour être sur de passer par les évènements onExit des composant :
  NextControl(Self);
BFin.Enabled:=FALSE ; DesErreurs:=FALSE ;
if not Steps[P.ActivePage.PageIndex+1] then
  BEGIN
  if (P.ActivePage=Fichier) And Not ShuntFichier then if not ImportTexte(False) then Exit ;
  if (P.ActivePage=Controle) then
     BEGIN
     if (FFormat.Value<>'RAP') and (FFormat.Value<>'CRA') and (FFormat.Value<>'MP') then
       BEGIN
{$IFNDEF JACADI}
       EnableControls(Self,False) ;
       if not VerifOk(DesErreurs) then BEGIN EnableControls(Self,True) ; Exit ; END ; //Si aucun mouvement à intégrer
       EnableControls(Self,True) ;
       If DesErreurs Then
          BEGIN
          If HM.Execute(55,'','')<>mrYes Then BEGIN EnableControls(Self,True) ; Exit ; END ;
          END ;
{$ENDIF}
       END else
     if (FFormat.Value='RAP') or (FFormat.Value='CRA') then
       BEGIN
       if not OkRefPointageSaisis then BEGIN HM.Execute(49,'','') ; Exit ; END ;
       END ;
     END ;
  if P.ActivePage=Integration then BEGIN BFin.Enabled:=TRUE ; (*EBIntegreClick(nil) ;*) Exit ; END ;
  END ;
inherited;
if P.ActivePage=Integration then BEGIN BFin.Enabled:=TRUE ; (*EBIntegreClick(nil) ;*) Exit ; END ;
end;

procedure TFAssistImport.EBFormatClick(Sender: TObject);
begin
ChoixFormatImpExp('X' ,Lequel) ;
If not (ChoixFmt.OkSauve) then Exit ;
If RecupSISCO Or RecupSERVANT Or ImportMvtSU Then Exit ;
FileNameChange(nil) ;
With ChoixFmt do
  BEGIN
  if Format<>'' then BEGIN FFormat.Value:=Format ; FFormatChange(Nil) ; END ; 
  FFormat.Enabled:=not (FixeFmt) ;
  if Fichier<>'' then FileName.Text:=Fichier ;
  FileName.Enabled:=not (FixeFichier) ;
  RechFile.Enabled:=not (FixeFichier) ;
  END ;
end;

procedure TFAssistImport.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFAssistImport.CaptionsDesMsgBox ;
var j,i : integer ;
    St : String ;
BEGIN
For i:=0 to HM.Mess.Count-1 do
  BEGIN
  St:=HM.Mess[i] ;
  j:=Pos('?CAPTION?',UpperCase(St)) ;
  if j>0 then BEGIN System.Delete(St,j,9) ; Insert(Caption,St,j) ; END ;
  HM.Mess[i]:=St ;
  END ;
END ;

procedure TFAssistImport.AddMenuPops ;
var mi  : TMenuItem ;
    i : integer ;
BEGIN
for i:=0 to P.ControlCount-1 do
   BEGIN
   mi := TMenuItem.Create(Self);
   mi.Caption := TTabSheet(P.Controls[i]).Hint ;
   mi.OnClick := StepMenuClick ;
   StepMenu.Items.Add(mi);
   END ;

END ;

procedure TFAssistImport.ChargettQualP ;
var Q1 : TQuery ;
BEGIN
if (Lequel<>'BAL') then
  BEGIN
  if not V_PGI.Controleur then
    BEGIN
    QualifPiece.DataType:='' ;
    QualifPiece.Items.Clear ;
    QualifPiece.Values.Clear ;
    Q1:=OpenSQL('SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="QFP" AND CO_CODE<>"R"',True) ;
    While not Q1.Eof do
      BEGIN
      QualifPiece.Values.Add(Q1.FindField('CO_CODE').AsString) ;
      QualifPiece.Items.Add(Q1.FindField('CO_LIBELLE').AsString) ;
      Q1.Next ;
      END ;
    Ferme(Q1) ;
    END ;
  END ;
QualifPiece.Value:='N' ; FirstFind:=False ;
END ;

procedure TFAssistImport.ChargeTitres ;
BEGIN
Prefixe := 'E' ; QuelleEtape:=Low(Steps) ;
ScenarioImport.Enabled:=FALSE ; TScenarioImport.Enabled:=FALSE ;
if Lequel='FEC' then
  BEGIN
  Caption:=HTitres.Mess[0] ;
  ScenarioImport.Enabled:=TRUE ; TScenarioImport.Enabled:=TRUE ;
  END else
if Lequel='FBA' then
  BEGIN
  Caption:=HTitres.Mess[1] ;
  QuelleEtape:=NbEtapes ;
  END else
if Lequel='FBE' then
  BEGIN
  Prefixe := 'BE' ;
  Caption:=HTitres.Mess[2] ;
  CBDoublons.Checked:=False ; CBDoublons.enabled:=False ;
  END else
if Lequel='FOD' then
  BEGIN
  Prefixe := 'Y' ;
  Caption:=HTitres.Mess[3] ;
  END ;
HTitre.Caption:=HTitres.Mess[9]+LowerCase(Caption) ;
END ;

procedure TFAssistImport.FormCreate(Sender: TObject);
begin
  inherited;
// Liste des comptes à créer / Modifier

(* GP le 15/10/2002 : correction bug plantage dans CCIMPEXP : Dû à correction Sur récupSISCO
FillChar(InfoImp,SizeOf(InfoImp),#0) ;
InfoImp.LGenLu:=TStringList.Create ;
InfoImp.LAuxLu:=TStringList.Create ;
InfoImp.LAnaLu:=TStringList.Create ;
InfoImp.LJalLu:=TStringList.Create ;
InfoImp.LMP:=TStringList.Create ;
InfoImp.LMR:=TStringList.Create ;
InfoImp.LRGT:=TStringList.Create ;
InfoImp.ListeCptFaux:=TList.Create ;
InfoImp.ListePieceFausse:=TList.Create ;
InfoImp.ListePieceIntegre:=TStringList.Create ;
InfoImp.ListeEntetePieceFausse:=TStringList.Create ;
InfoImp.ListeEnteteDoublon:=TStringList.Create ;
InfoImp.CRListeEnteteDoublon:=TList.Create ;
InfoImp.ListePbAna:=TStringList.Create ;
InfoImp.LSoucheBOR:=TStringList.Create ;
*)
CreateListeImp(InfoImp) ;
end;

procedure TFAssistImport.FormShow(Sender: TObject);
begin
(*
if ctxPCL in V_PGI.PGIContexte then
  BEGIN
  FORCEDEVISE.Enabled:=FALSE ; TFORCEDEVISE.Enabled:=FALSE ; FORCEDEVISE.Value:='VFI' ;
  END ;
*)
If Not ParamEuroOk Then
  BEGIN
  HShowMessage('0;ATTENTION ! Paramétrage société incomplet;Vérifiez les paramètres EURO dans la fiche société;E;O;O;O;','','') ;
  PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
  END ;
If SemiAuto Then
  BEGIN
  VStatus:=Status ;
  Status.Caption := Copyright ;
  END Else Status.Visible:=FALSE ;

TSociete.Open ; NbEtapes:=P.PageCount+1 ; ReprendreLeFichier:=FALSE ;
OkSauve:=False ; AnnuleImport:=False ;
AddMenuPops ;
ChangeDataTypeFmt('X',Lequel,FFormat) ;
inherited ;
CaptionsDesMsgBox ;
GereFMTChoix('X',Lequel,FFormat,FileName,RechFile) ;
ChargeEtapes ; FinEtape.Checked := Steps[P.ActivePage.PageIndex+1] ; ChangeStepBitmap ;
ChargettQualP ;
ChargeTitres ;
SourisNormale ;
//CBForceNumPiece.Visible:=V_PGI.SAV ;
PremierControle:=TRUE ; IntegrationFaite:=FALSE ;
If Sn2Orli Then BEGIN FFormat.value:='SN2' ; CBSn2Orli.Checked:=TRUE ; END ;
If RecupSISCO Then BEGIN FFormat.value:='SIS' ; FileName.Text:=StFichier ; FicSISCOConverti:=TRUE ; END Else
  If RecupSERVANT Then BEGIN FFormat.value:='CGE' ; FileName.Text:=StFichier ; END Else
    If ImportMvtSU Then BEGIN FFormat.value:='CGE' ; FileName.Text:=StFichier ; END ;
If RecupSISCO And (VH^.RecupLTL Or VH^.RecupSISCOPGI) And BigVol Then
  BEGIN
  SuffixeBigVol.Visible:=TRUE ; TSuffixeBigVol.Visible:=TRUE ;
  END Else SuffixeBigVol.Text:='' ;
end;

procedure TFAssistImport.EBControleClick(Sender: TObject);
Var DesErreurs : Boolean ;
begin
EnableControls(Self,False) ;
If Not PremierControle Then MajOkControle ;
VerifOk(DesErreurs) ;
EnableControls(Self,True) ;
PremierControle:=FALSE ;
end;

function TFAssistImport.VerifOk(Var DesErreurs : Boolean) : boolean ;
var TypeV : TTypeVerif ;
    FicRap : TextFile ;
    OkRapport : Boolean ;
BEGIN
Result:=False ; DesErreurs:=FALSE ;
AffAide(9) ;
AffAide(5) ;
InfoImp.ListeEntetePieceFausse.Clear ;
InfoImp.ListeEntetePieceFausse.Sorted:=TRUE ;
InfoImp.ListeEntetePieceFausse.Duplicates:=DupIgnore ;
InfoImp.ListePbAna.Clear ;
InfoImp.ListePbAna.Sorted:=TRUE ;
InfoImp.ListePbAna.Duplicates:=DupIgnore ;
TypeV:=tvEcr ;
if Prefixe='BE' then TypeV:=tvEcrBud else
  if Prefixe='Y' then  TypeV:=tvEcrAna ;
VideListe(InfoImp.ListePieceFausse) ;
if not VerPourImp3(InfoImp.ListeEntetePieceFausse,InfoImp.ListePbAna,InfoImp.ListePieceFausse,False,TypeV,InfoImp.SC.ShuntPbAna,InfoImp.PbAna) then
  BEGIN
  DesErreurs:=TRUE ;
  MajImpErr(InfoImp.ListeEntetePieceFausse) ;
  END ;
EBListeErreur.Visible:=InfoImp.ListePieceFausse.Count>0 ;
if AucunMvtControlesOk then
  BEGIN
  If (Not ImpV3) Or (InfoImp.NbLigLettre<=0) Then HM.Execute(42,'','') ;
  EcrireRapportDebutIntegration(TRUE,InfoImp,FicRap,OkRapport) ;
  FaitFichierDoublon(InfoImp) ; FaitFichierRejet(InfoImp) ;
  If (Not ImpV3) Or (InfoImp.NbLigLettre<=0) Then Exit Else
     If InfoImp.NbLigLettre>0 Then LettrageImport(InfoImp) ;
  END ;
LAide.Caption:='' ;
SourisNormale ;
HM.Execute(3,'','') ;
if not FinEtape.Checked then FinEtapeClick(nil) ;
Result:=True ;
END ;

procedure TFAssistImport.DelOldRef ;
var i : integer ;
    StDate,StRef : String ;
    Gene : String17 ;
BEGIN
for i:=0 to OldReferences.Count-1 do
  BEGIN
  StDate:=OldReferences[i] ;
  Gene:=ReadTokenSt(StDate) ;
  StRef:=ReadTokenSt(StDate) ;
  ExecuteSQL('DELETE FROM EEXBQ WHERE EE_GENERAL="'+Gene+'" AND EE_REFPOINTAGE="'+StRef+'"'+
             ' AND NOT EXISTS(SELECT DISTINCT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+Gene+'"'+
             ' AND E_REFPOINTAGE="'+StRef+'" AND E_DATEPOINTAGE="'+StDate+'")') ;
  END ;
END ;

procedure TFAssistImport.EBIntegreClick(Sender: TObject);
var Q1 : TQuery ;
begin
if (FFormat.Value='RAP') or (FFormat.Value='CRA') then
  BEGIN
  Q1:=OpenSQL('SELECT IE_CHRONO FROM IMPECR WHERE IE_INTEGRE="X"',True) ;
  if not Q1.Eof then
    case HM.Execute(52,'','') of
      mrYes : ExecuteSQL('UPDATE IMPECR SET IE_INTEGRE="-" WHERE IE_INTEGRE="X"') ;
      else BEGIN Ferme(Q1) ; Exit ; END ;
      END ;
  Ferme(Q1) ;
  END ;
AnnuleImport:=False ;
EnableControls(Self,False) ; BStop.Visible:=TRUE ; EBIntegre.Visible:=FALSE ;
If Not InfoImp.ImportOk Then
   BEGIN
   AlimInfoImp ; RempliListeInfoImp(InfoImp) ;
   END ;
InfoImp.ForceNumPiece:=CBForceNumPiece.Checked ;
InfoImp.ListePieceIntegre.Clear ;
InfoImp.ListePieceIntegre.Sorted:=TRUE ;
InfoImp.ListePieceIntegre.Duplicates:=DupIgnore ;
IntegreEcr(Self,InfoImp) ;
IntegrationFaite:=TRUE ;
EnableControls(Self,True) ;  EBIntegre.Visible:=TRUE ; BStop.Visible:=FALSE ;
EBIntegre.Enabled:=TRUE ;
if AnnuleImport then BEGIN AnnuleImport:=False ; Exit ; END ;
AffAide(56) ;
FaitFichierDoublon(InfoImp) ; FaitFichierRejet(InfoImp) ;
LAide.Caption:='' ;
if (FFormat.Value='RAP') or (FFormat.Value='CRA') then
  BEGIN
  DelOldRef ;
  OldReferences.Clear ; OldReferences.Free ;
  HM.Execute(47,'','')
  END else HM.Execute(10,'','') ;
if not FinEtape.Checked then FinEtapeClick(nil) ;
If InfoImp.NbLigLettre>0 Then LettrageImport(InfoImp) ;
if ChoixFmt.Detruire then
  BEGIN
  DeleteFile(InfoImp.NomFic) ;
  If (InfoImp.FormatOrigine='SIS') And (InfoImp.NomFicOrigine<>'') And (Not RecupSISCO) Then DeleteFile(InfoImp.NomFicOrigine) ;
  If (InfoImp.FormatOrigine='CGE') And (InfoImp.NomFicOrigine<>'') And (ImportMvtSU) Then DeleteFile(InfoImp.NomFicOrigine) ;
  END ;
end;

procedure TFAssistImport.AfficheMessControle(Var InfoImp : TInfoImport) ;
BEGIN
ZNbPieceLu.Caption:=IntToStr(InfoImp.NbPiece) ;
ZNbLigneLu.Caption:=IntToStr(InfoImp.NbLigIntegre) ;
ZTotDLu.Text:=StrfMontant(Abs(InfoImp.TotDeb),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True) ;
ZTotCLu.Text:=StrfMontant(Abs(InfoImp.TotCred),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True) ;
ZNbGenLu.Caption:=IntToStr(InfoImp.LGenLu.Count) ;
ZNbAuxLu.Caption:=IntToStr(InfoImp.LAuxLu.Count) ;
ZNbAnaLu.Caption:=IntToStr(InfoImp.LAnaLu.Count) ;
ZNbJalLu.Caption:=IntToStr(InfoImp.LJalLu.Count) ;
ZNbGenLuFaux.Caption:=IntToStr(InfoImp.NbGenFaux) ;
ZNbAuxLuFaux.Caption:=IntToStr(InfoImp.NbAuxFaux) ;
ZNbAnaLuFaux.Caption:=IntToStr(InfoImp.NbAnaFaux) ;
ZNbJalLuFaux.Caption:=IntToStr(InfoImp.NbJalFaux) ;
EBListeCpt.Visible:=(InfoImp.NbAnaFaux<>0) Or
                    (InfoImp.NbAuxFaux<>0) Or
                    (InfoImp.NbGenFaux<>0) ;
END ;

function TFAssistImport.ImportTexte(Initialise : boolean) : boolean ;
var Q : TQuery ;
    i,j,k : integer ;
    PbSISCO : Boolean ;
BEGIN
PremierControle:=TRUE ; IntegrationFaite:=FALSE ;
Result:=False ; EBListeDoublon.Visible:=FALSE ;
AnnuleImport:=False ; if Initialise then
  BEGIN
  for i:=Low(Steps) to High(Steps) do Steps[i]:=False ;
  SauveEtapes(False) ; FinEtape.Checked:=False ;
  END ;
if (Lequel='FEC') and ((FFormat.Value='CLB') or (FFormat.Value='CPR') or (FFormat.Value='CRL')
//   or (FFormat.Value='CRA')
   or (FFormat.Value='CTR')
   or (FFormat.Value='CVI')) then BEGIN ShowMessage(Msg.Mess[9]) ; Exit ; END ;
//if Not FormatOk(FileName.Text,FFormat.Value,Lequel) then BEGIN HTitres.Execute(5,Caption,'') ; Exit ; END ;
AffAide(32) ;
EnableControls(Self,False) ;
If RecupSISCO And (FFormat.Value='SIS') Then
  BEGIN
  FicSISCOConverti:=TRUE ;
  END Else If (FFormat.Value='SIS') And (Not FicSISCOConverti) Then
  BEGIN
  PbSISCO:=FALSE ;
  if not FileExists(FileName.Text) then
     BEGIN
     HM.Execute(41,'','') ;
     PbSISCO:=TRUE ;
     END ;
  j:=TransfertSISCO(FileName.Text,FALSE,k) ;
  If j<>0 Then
    BEGIN
    PbSISCO:=TRUE ; HM.Execute(57,'','') ;
    END ;
  If PbSISCO Then
    BEGIN
    EnableControls(Self,True) ; LAide.Caption:='' ;
    AnnuleImport:=True ; EBStop.Visible:=FALSE ;
    Exit ;
    END ;
  FicSISCOConverti:=TRUE ;
  Application.ProcessMessages ; //delay(5000) ;
  END ;
EBImporter.Visible:=FALSE ; EBStop.Visible:=True ;
AlimInfoImp ; InfoImp.NbGenFaux:=0 ; InfoImp.NbAuxFaux:=0 ;
InfoImp.NbAnaFaux:=0 ; InfoImp.NbJalFaux:=0 ;
VideListeInfoImp(InfoImp,FALSE) ;
InfoImp.ListeEnteteDoublon.Sorted:=TRUE ;
InfoImp.ListeEnteteDoublon.Duplicates:=DupIgnore ;
if not ImporteLesEcritures(HM,InfoImp,FIgnorel1.Checked) then
  BEGIN
  EnableControls(Self,True) ; LAide.Caption:='' ;
  AnnuleImport:=True ; EBStop.Visible:=FALSE ;
  EBImporter.Visible:=TRUE ; EBImporter.Enabled:=TRUE ;
  Exit ;
  END ;
InfoImp.ImportOk:=TRUE ;
SauveEtapes(True) ;

AfficheMessControle(InfoImp) ;
EnableControls(Self,True) ; EBStop.Visible:=FALSE ; EBImporter.Visible:=TRUE ;
EBImporter.Enabled:=TRUE ;
Q:=OpenSQL('SELECT * FROM IMPECR WHERE IE_CHRONO<2',True) ;
if Q.Eof then
  BEGIN
  Ferme(Q) ;
  if (FFormat.Value='RAP') or (FFormat.Value='CRA') then HM.Execute(51,'','')
                                                    else HM.Execute(30,'','') ;
  AnnuleImport:=True ; EBStop.Visible:=FALSE ; EBImporter.Visible:=TRUE ;
  EBImporter.Enabled:=TRUE ;
  Exit ;
  END ;
Ferme(Q);
AffAide(9) ;
If InfoImp.CtrlDB And (InfoImp.CRListeEnteteDoublon.Count>0) Then
  RapportDeSuppression(InfoImp.CRListeEnteteDoublon,6) ;
LAide.Caption:='' ;
if not FinEtape.Checked then FinEtapeClick(nil) ;
Result:=True ;
ebStop.Visible:=False ;
EBListeErreur.Visible:=FALSE ;
EBListeDoublon.Visible:=InfoImp.CtrlDB And (InfoImp.CRListeEnteteDoublon.Count>0) ;
if (FFormat.Value='SIS') And (FicSISCOConverti) Then BVoir.Enabled:=TRUE ;
END ;

procedure TFAssistImport.EBImporterClick(Sender: TObject);
begin
ImportTexte(True) ;
end;

procedure TFAssistImport.bFinClick(Sender: TObject);
begin
If Not IntegrationFaite Then EBIntegreClick(Nil) ;
OkSauve:=True ; Close ;
end;

procedure TFAssistImport.SauveEtapes(Fic : Boolean) ;
var f,s   : String ;
    i,D : integer ;
    OkSauve : boolean ;
BEGIN
OkSauve:=False ;
for i:=low(Steps) to High(Steps) do if Steps[i] then BEGIN OkSauve:=True ; Break ; END ;
if not OkSauve then Exit ;
s := '' ; D:=0 ;
{$IFDEF SPEC302}
TSociete.Edit ;
s:=TSociete.FindField('SO_ETAPEIMPORT').AsString ;
if (s='') then while Length(s)<=High(Steps) do s:=s+'0' ;
f:=ReadtokenSt(s) ;
if not Fic then
  BEGIN
  if QuelleEtape=NbEtapes then D:=NbEtapes-1 ;
  for i:=QuelleEtape to QuelleEtape+NbEtapes+1 do
     if (i-D in [low(Steps)..High(Steps)]) then if (Steps[i-D]) Or VH^.Mugler then f[i]:='1' else f[i]:='0' ;
  END ;
if Fic then s:=f+';'+FileName.Text+';'+FFormat.Value+';'+Lequel else s:=f+';'+s ;
TSociete.FindField('SO_ETAPEIMPORT').AsString := s ;
TSociete.Post ;
{$ELSE}
s:=GetParamSoc('SO_ETAPEIMPORT') ;
if s='' then while Length(s)<=High(Steps) do s:=s+'0' ;
f:=ReadtokenSt(s) ;
if not Fic then
  BEGIN
  if QuelleEtape=NbEtapes then D:=NbEtapes-1 ;
  for i:=QuelleEtape to QuelleEtape+NbEtapes+1 do
     if (i-D in [low(Steps)..High(Steps)]) then if (Steps[i-D]) Or VH^.Mugler then f[i]:='1' else f[i]:='0' ;
  END ;
if Fic then s:=f+';'+FileName.Text+';'+FFormat.Value+';'+Lequel else s:=f+';'+s ;
If Length(S)>70 Then S:=Copy(S,1,70) ;
SetParamSoc('SO_ETAPEIMPORT',s) ;
{$ENDIF}
END ;

procedure TFAssistImport.ChargeEtapes ;
var Leq,etap,fic,s,fmt{,StM}   : String ;
    i{,D} : integer ;
BEGIN
for i:=Low(Steps) to High(Steps) do Steps[i] := False ;
{$IFDEF SPEC302}
s := TSociete.FindField('SO_ETAPEIMPORT').AsString ;
{$ELSE}
s:=GetParamSoc('SO_ETAPEIMPORT') ;
{$ENDIF}
etap:=Trim(ReadTokenSt(s)) ;
fic:=Trim(ReadTokenSt(s)) ;
fmt:=Trim(ReadTokenSt(s)) ;
Leq:=Trim(ReadTokenSt(s)) ;
// Type d'import différent
{if (Leq<>Lequel) then Exit ;
D:=0 ;
if QuelleEtape=NbEtapes then D:=NbEtapes-1 ;
StM:='' ;}
END ;

procedure TFAssistImport.StepMenuClick(Sender: TObject);
var j,i : Integer ;
begin
i:=TMenuItem(Sender).MenuIndex-2 ;
if (i=P.ActivePage.PageIndex) then Exit ;
if not EffectuerLesEtapes(i+1,False) then Exit ;
if (i>P.ActivePage.PageIndex) then
  BEGIN
  for j:=P.ActivePage.PageIndex to i-1 do
    if Steps[j-1] then bSuivantClick(nil) ;
  END else
  BEGIN
  RestorePage ;
  P.ActivePage := P.Pages[i] ;
  PChange(nil) ;
  END ;
END ;

function TFAssistImport.EffectuerLesEtapes(Leq : integer ; Fin : boolean) : boolean;
var j,i,k : integer ;
    Ok  : Boolean ;
    NbEtapes : Integer ;
BEGIN
Result:=False ;
Ok:=False ; NbEtapes:=0 ;
i:=P.ActivePage.PageIndex+1 ;
// Effectuer la dernière étape si FinClick ;
if Fin then k:=0 else k:=-1 ;
if Leq>=i+1 then
  for j:=i to Leq+k do if not Steps[j] then Inc(NbEtapes) ;
if Fin then
  BEGIN
  if NbEtapes>1 then
    BEGIN
    case HM.Execute(37,'','') of
        mrYes    : Ok:=True ;
        mrNo     : Ok:=False ;
        mrCancel : Exit ;
        END ;
    if ok then
      for j:=i to P.PageCount do if Steps[j-1] then bSuivantClick(nil) ;
    END else
    BEGIN
    if (NbEtapes>0) then
      case HM.Execute(38,'','') of
        mrYes    : EBIntegreClick(nil) ;
        mrNo     : ;
        mrCancel : Exit ;
        END ;
    END ;
  END else
  BEGIN
  if (NbEtapes>1) then
    BEGIN
    case HM.Execute(31,'','') of
      // mrYes    : Ok:=True ;
      mrNo     : Exit ;
      END ;
    END else
    BEGIN
    if (NbEtapes>0) then
      case HM.Execute(37,'','') of
        // mrYes    : Ok:=True ;
        mrNo     : Exit ;
        END ;
    END ;
  END ;
Result:=True ;
END ;

procedure TFAssistImport.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
If Not ParamEuroOk Then Exit ;
if Not AnnuleImport then if not EffectuerLesEtapes(P.PageCount,True) then BEGIN CanClose:=False ; Exit ; END else AnnuleImport:=False ;
inherited;
if (CanClose and OkSauve) then BEGIN SauveEtapes(False) ; SauveEtapes(True) ; TSociete.Close ; END ;
end;

Procedure TFAssistImport.AffAide(N : Integer) ;
BEGIN
LAide.Caption:=HM.Mess[N] ; Application.ProcessMessages ;
END ;

procedure TFAssistImport.BCValideClick(Sender: TObject);
var i,n : integer ;
    Tiers : String17 ;
    St,EL,Eche : String ;
    NumEche : Integer ;
    CTiers : TGTiers ;
    QJ : TQuery ;
begin
n:=GC.RowCount ; InitMove(n+3,'') ; MoveCur(False) ;
for i:=0 to n-1 do
  BEGIN
  Tiers:=Trim(UpperCase(GC.Cells[1,i])) ;
  if Tiers='' then BEGIN MoveCur(False) ; Continue ; END ;
  CTiers:=TGTiers.Create(Tiers) ;
  EL:='RI' ;  Eche:='-' ; NumEche:=0 ;
  if CTiers.Lettrable then BEGIN EL:='AL' ; Eche:='X' ; NumEche:=1 ; END ;
  ExecuteSQL('Update IMPECR SET IE_ETATLETTRAGE="'+EL+'",IE_ECHE="'+Eche+'",IE_NUMECHE="'+IntToStr(NumEche)+'",IE_AUXILIAIRE="'+Tiers+'" WHERE IE_GENERAL="'+Trim(UpperCase(GC.Cells[0,i]))+'"') ;
  MoveCur(False) ;
  CTiers.Free ;
  END ;
St:=',IE_ECRANOUVEAU="N"' ;
QJ:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+JALAN.Value+'"',True) ;
if QJ.Fields[0].AsString='ANO' then St:=',IE_ECRANOUVEAU="OAN"' ;
Ferme(QJ) ;
ExecuteSQL('Update IMPECR Set IE_JOURNAL="'+JALAN.Value+'"'+St+' WHERE IE_JOURNAL="'+JAL1.Text+'"') ;
MoveCur(False) ;
St:=',IE_ECRANOUVEAU="N"' ;
QJ:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+JALOD.Value+'"',True) ;
if QJ.Fields[0].AsString='ANO' then St:=',IE_ECRANOUVEAU="OAN"' ;
Ferme(QJ) ;
ExecuteSQL('Update IMPECR Set IE_JOURNAL="'+JALOD.Value+'"'+St+' WHERE IE_JOURNAL="'+JAL2.Text+'"') ;
MoveCur(False) ;
FiniMove ;
ActivePanels(Self,True,False) ; Application.ProcessMessages ;
end;

procedure TFAssistImport.BCAbandonClick(Sender: TObject);
begin ActivePanels(Self,True,False) ; end;


// Liste des comptes pour réf. de pointage
{procedure TFAssistImport.ModifRefPointage ;
var Q1 : TQuery ;
BEGIN
if OldReferences=nil then
  BEGIN
  OldReferences:=TStringList.Create ;
  Q1:=OpenSQL('SELECT DISTINCT IE_GENERAL,IE_REFPOINTAGE,IE_DATEPOINTAGE FROM IMPECR LEFT JOIN GENERAUX ON '
             +'IE_GENERAL=G_GENERAL WHERE G_POINTABLE="X"',True) ;
  While not Q1.Eof do
    BEGIN
    OldReferences.Add(Q1.Fields[0].AsString+';'+Q1.Fields[1].AsString+';'+USDateTime(Q1.Fields[2].AsDateTime)) ;
    Q1.Next ;
    END ;
  Ferme(Q1) ;
  END ;
ChoixDesReferences ;
END ;}

procedure TFAssistImport.bAnnulerClick(Sender: TObject);
begin
  inherited;
AnnuleImport:=True ;
end;

procedure TFAssistImport.EBStopClick(Sender: TObject);
begin AnnuleImport:=True ; end;

function ExisteLettrage : boolean ;
var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT DISTINCT IE_TYPEECR From IMPECR WHERE IE_TYPEECR="L"',True) ;
Result:=Not Q.Eof ;
Ferme(Q) ;
END ;

procedure TFAssistImport.FFormatChange(Sender: TObject);
Var Q : TQuery ;
begin
if FFormat.Value='EDI' then Sauve.FilterIndex:=4 else
  if (FFormat.Value='HAL') Or (FFormat.Value='HLI') Or
     (FFormat.Value='CGE') Or (FFormat.Value='CGN') then Sauve.FilterIndex:=1 else
    if FFormat.Value='SIS' then Sauve.FilterIndex:=5 else
      if (FFormat.Value='SAA') or (FFormat.Value='SN2') or (FFormat.Value='RAP')
      or (FFormat.Value='MP') then Sauve.FilterIndex:=3 ;
QualifPiece.Enabled:=((FFormat.Value<>'RAP') and (FFormat.Value<>'MP') and (FFormat.Value<>'CRA')) ;
H_QualifPiece.Enabled:=QualifPiece.Enabled ;
CbPositifs.Enabled:=QualifPiece.Enabled ;
CbDoublons.Enabled:=QualifPiece.Enabled ;
LgComptes.Enabled:=QualifPiece.Enabled ;
EBControle.Enabled:=QualifPiece.Enabled ;
//CbLettrage.Enabled:=ExisteLettrage ;
CbLettrage.Enabled:=FALSE ;
CbLettrage.Visible:=FALSE ;
if not QualifPiece.Enabled then
  BEGIN
  CbPositifs.Checked:=False ;
  CbDoublons.Checked:=False ;
  LgComptes.Checked:=False ;
  END ;
if (CbLettrage.Checked) and (not CbLettrage.Enabled) then CbLettrage.Checked:=False ;

if (FFormat.Value='RAP') or (FFormat.Value='CRA') then
  BEGIN
  EBIntegre.Caption:= HM.Mess[46] ;
  END else
  BEGIN
  EBIntegre.Caption:=HM.Mess[48] ;
  END ;
if (FFormat.Value='SIS') And (Not FicSISCOConverti) And (Not RecupSISCO) Then BVoir.Enabled:=FALSE ;
CBSn2Orli.Visible:=FFormat.Value='SN2' ; If Not CBSn2Orli.Visible Then CBSn2Orli.Checked:=FALSE ;
ScenarioImport.Plus:='FS_IMPORT="X" AND FS_NATURE="FEC" AND FS_FORMAT="'+FFormat.Value+'" ';
ScenarioImport.Text:='' ;
CBDoublons.checked := False ;
CBForceNumPiece.checked := False ;
// Chargement d'un scénario import
	if FFormat.Value<>'' then
  	begin
		Q := OpenSQL('SELECT MIN(FS_CODE) CODE FROM FMTSUP WHERE FS_IMPORT="X" AND FS_NATURE="FEC" AND FS_FORMAT="'+FFormat.Value+'" ', True) ;
		if not Q.Eof then
    	begin
      ScenarioImport.Text := Q.FindField('CODE').AsString ;
      SCENARIOIMPORTExit(nil) ;
      end ;
    Ferme(Q) ;
    end ;

end;

procedure TFAssistImport.FileNameChange(Sender: TObject);
begin
Steps[P.ActivePage.PageIndex+1]:=False ;
end;

procedure TFAssistImport.GCDblClick(Sender: TObject);
begin
  inherited;
HCpte.ZoomTable:=tzTiers ;
HCpte.Text:=GC.Cells[1,GC.Row] ;
HCpte.Bourre:=True ;
if GChercheCompte(HCpte,ProcZoomTiers) then GC.Cells[1,GC.Row]:=HCpte.Text ;
end;


//*******************************************************
procedure TFAssistImport.BVoirClick(Sender: TObject);
var StErr : String ;
    fmtFic : Integer ;
    NomFic : String ;
begin
NomFic:=FileName.text ;
StErr :='' ;
FmtFic:=0 ;
If (FFormat.Value='SIS') Then
  BEGIN
  If (Not FicSISCOConverti) Then Exit Else If RecupSISCO Then FmtFic:=4 Else FmtFic:=3 ;
  If RecupSISCO Then NomFic:=NomFichierRecupSISCO(NomFic)
                Else NomFic:=NewNomFic(NomFic,'CGN') ;
  END ;
If ImportMvtSU Then
  BEGIN
  NomFic:=NomFichierImportSU(NomFic) ;
  END ;
if FFormat.Value='HLI' then FmtFic:=1 else if FFormat.Value='HAL' then FmtFic:=2 else
if FFormat.Value='CGN' then FmtFic:=3 else if FFormat.Value='CGE' then FmtFic:=4 ;
if FileExists(NomFic) then VisuLignesErreurs(NomFic,StErr,FmtFic,LgComptes.Checked,FIgnorel1.Checked)
                             else HM.Execute(41,'','') ;
end;

procedure TFAssistImport.FChangeMethodeClick(Sender: TObject);
begin
  inherited;
Imp_Methode1:=Not FChangeMethode.Checked ;
EBStop.Enabled:=Imp_Methode1 ;
end;

procedure TFAssistImport.FReqParamClick(Sender: TObject);
begin
  inherited;
QAJParam:=FReqParam.Checked ;
end;

procedure TFAssistImport.VisuImpEcrClick(Sender: TObject);
begin
  inherited;
MultiCritereImpEcr(taModif,InfoImp.ListePieceFausse) ;
end;

procedure TFAssistImport.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
VideListeInfoImp(InfoImp,TRUE) ;
end;

procedure TFAssistImport.EBListeCptClick(Sender: TObject);
Var GGRien : Boolean ;
begin
  inherited;
If (InfoImp.NbGenFaux>0) Or (InfoImp.NbAuxFaux>0) Or (InfoImp.NbAnaFaux>0) Then
   BEGIN
   RapportdErreurMvt(InfoImP.ListeCptFaux,8,GGRien,False) ;
   END ;
end;

procedure TFAssistImport.EBListeErreurClick(Sender: TObject);
Var GGRien : Boolean ;
begin
  inherited;
GGrien:=FALSE ;
If InfoImp.ListePieceFausse.Count>0 Then
   RapportdErreurMvt(InfoImp.ListePieceFausse,3,GGRien,FALSE) ;

end;

procedure TFAssistImport.EBListeDoublonClick(Sender: TObject);
begin
  inherited;
If InfoImp.CtrlDB And (InfoImp.CRListeEnteteDoublon.Count>0) Then
  RapportDeSuppression(InfoImp.CRListeEnteteDoublon,6) ;
end;

procedure TFAssistImport.CBDoublonsClick(Sender: TObject);
begin
  inherited;
CBForceNumPiece.Enabled:=CBDoublons.Checked ;
If Not CBForceNumPiece.Enabled Then CBForceNumPiece.Checked:=FALSE ; 
end;

procedure TFAssistImport.SCENARIOIMPORTKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  inherited;
if ((SCENARIOIMPORT.Focused) and (Key=VK_F6)) Then
  BEGIN
  Key:=0 ;
  ParamSupImport('X','FEC',FFormat.Value,ScenarioImport.Text,tamodif,0,TRUE);
  END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 15/01/2003
Modifié le ... : 15/01/2003
Description .. : Initialise la coche "Controle des doublons" en fonction du
Suite ........ : scénario sélectionné
Mots clefs ... :
*****************************************************************}
procedure TFAssistImport.SCENARIOIMPORTExit(Sender: TObject);
Var Q : TQuery ;
begin
  if not  CBDoublons.Enabled then Exit ;
  if SCENARIOIMPORT.text = '' then
  	begin
		CBDoublons.checked := False ;
    CBForceNumPiece.checked := False ;
    Exit ;
    end ;
	// Récupération du paramètre "controle des doublons" à partir du scénario sélectionné
  Q := OpenSQL('SELECT FS_DOUBLON, FS_FORCEPIECE FROM FMTSUP '
  							+ 'WHERE FS_IMPORT="X" AND FS_NATURE="FEC" AND FS_FORMAT="'+FFormat.Value+'" '
                + 'AND FS_CODE="' + SCENARIOIMPORT.text + '"', True) ;
  if not Q.Eof then
    begin
    CBDoublons.checked := (Uppercase(Q.FindField('FS_DOUBLON').AsString) = 'X') ;
    CBForceNumPiece.checked := (Uppercase(Q.FindField('FS_FORCEPIECE').AsString) = 'X') ;
    end ;
  Ferme(Q);
end;

end.
