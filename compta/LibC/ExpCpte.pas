unit ExpCpte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hcompte, HDB, Hctrls, Buttons, ExtCtrls, HQry, DB, Ent1,
  Mask, hmsgbox, HEnt1, HStatus, EDI, {Filtre SG6 10/11/04 Gestion des filtres, FQ 14826} ComCtrls, ParamDat, Menus,
  HSysMenu, HTB97, HRegCpte, HPanel, UiUtil, RappType, ParamSoc, TImpFic,UObjFiltres, uDbxDataSet {SG6 10/11/04 Gestion des filtres FQ 14826} ;

procedure ExportCpte(Lequel : String) ;
procedure LanceExportCptExt(Lequel : String) ;

type PlanCpt = (cptGEN,cptAUX,cptANA,CptJalBUD,cptBUD,CptSEB) ;

type
  TFExpCpte = class(TForm)
    MsgBox: THMsgBox;
    Sauve: TSaveDialog;
    Pages: TPageControl;
    TabSheet1: TTabSheet;
    Bevel1: TBevel;
    LFile: THLabel;
    HLabel6: THLabel;
    Label7: THLabel;
    ENATUREGENE: THLabel;
    LFormat: THLabel;
    GComptes: TGroupBox;
    TCPTEDEBUT: THLabel;
    TCPTEFIN: THLabel;
    EDEPART: THLabel;
    EARRIVEE: THLabel;
    CPTEFIN: THDBCpteEdit;
    CPTEDEBUT: THDBCpteEdit;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    FDate1: TMaskEdit;
    FDate2: TMaskEdit;
    FFormat: THValComboBox;
    Nature: THValComboBox;
    FExport: TCheckBox;
    Msg: THMsgBox;
    HMTrad: THSystemMenu;
    HMTitre: THMsgBox;
    CAxe: THValComboBox;
    EAxe: THLabel;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    HPB: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BEDI: TToolbarButton97;
    BFormat: TToolbarButton97;
    BStop: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    FSousSection: TCheckBox;
    procedure BFermeClick(Sender: TObject);
    procedure RechFileClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BFormatClick(Sender: TObject);
    procedure BEDIClick(Sender: TObject);
    procedure FFormatChange(Sender: TObject);
    procedure BStopClick(Sender: TObject);
 {    procedure FFiltresChange(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);}

    procedure FDate1KeyPress(Sender: TObject; var Key: Char);
    procedure FDate2KeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure CAxeChange(Sender: TObject);
//    procedure POPFPopup(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    ObjFiltre:TObjFiltre; //SG6 10/11/04  Gestion des Filtres FQ 14826
    Lequel : String ;
    Arreter : boolean ;
    NbExport : LongInt ;
    CurNat : String ;
    Fichier : TextFile ;
    NumTypCpt : PlanCpt ;
    Quantite : integer ;
    Compte : String17 ;
    FiltreExp : String ;
    OkExp : Boolean ;
    procedure AfficheCompteRendu ;
    Procedure InitComptes ;
    Function ExporteCpteSAARI : Boolean ;
    Function ExporteTiers : Boolean ;
    Function ExporteFormatParametre : Boolean ;
   // EDI
    Function LanceExportPlanEDI : Boolean ;
    procedure Entete ;
    procedure Detail(Q : TQuery) ;
    procedure SousDetail(Str : string ; Q : Tquery) ;
    procedure Resume ;
    function ConstruitReq(OkUpdate : boolean) : boolean ;
    Procedure ReqSQL(OkUpdate : boolean ; Var St : String) ;
    function RecupChp(Leq : String ; Quoi : String) : String ;
    Function ExporteHalley : Boolean ;
    function RecupValChp(Q : Tquery) : String ;
    function TestBreak : Boolean ;
    Procedure TraiteExport ;
    Function ExporteCptCEGID : Boolean ;
    Procedure MajFlagExport ;
  public
  END ;

implementation

uses FmtChoix,ImporFmt,ImpUtil,ImpFicU,UtilTrans ;

{$R *.DFM}

procedure ExportCpte(Lequel : String) ;
var FExpCpte:TFExpCpte ;
    PP : THPanel ;
    Composants : TControlFiltre; //SG6   Gestion des Filtes 10/11/04   FQ 14826    
BEGIN
FExpCpte:=TFExpCpte.Create(Application) ;
FExpCpte.Lequel:=Lequel ;

//SG6 10/11/04 Gestion des Filtres      FQ 14826
Composants.PopupF   := FExpCpte.POPF;
Composants.Filtres  := FExpCpte.FFILTRES;
Composants.Filtre   := FExpCpte.BFILTRE;
Composants.PageCtrl := FExpCpte.Pages;
FExpCpte.ObjFiltre := TObjFiltre.Create(Composants, '');

PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FExpCpte.ShowModal ;
     finally
     FExpCpte.Free ;
     END ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FExpCpte,PP) ;
   FExpCpte.Show ;
   END ;
END ;

procedure LanceExportCptExt(Lequel : String) ;
var FExpCpte:TFExpCpte ;
BEGIN
FExpCpte:=TFExpCpte.Create(Application) ;
FExpCpte.Lequel:=Lequel ;
try
  FExpCpte.ShowModal ;
  finally
  FExpCpte.Free ;
  END ;
Screen.Cursor:=SyncrDefault ;
END ;


procedure TFExpCpte.BFermeClick(Sender: TObject);
BEGIN
  //SG6 13/01/05 FQ 15242
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
END;

procedure TFExpCpte.RechFileClick(Sender: TObject);
BEGIN
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
END ;

procedure TFExpCpte.FormShow(Sender: TObject);

BEGIN
FDate1.Text:=StDate1900 ; FDate2.Text:=StDate2099 ;
if lequel='G' then //Généraux
       BEGIN
       FiltreExp:='EXPFGE' ;
       CurNat:='FGE' ; NumTypCpt:=CptGEN ;
       Caption:=HMTitre.Mess[0] ;
       CPTEDEBUT.ZoomTable:=TzGeneral ;
       CPTEFIN.ZoomTable:=TzGeneral ;
       Nature.DataType:='ttNatGene' ;
       HelpContext:=6310000 ;
       END else
if lequel='A' then //Analytiques
       BEGIN
       FiltreExp:='EXPFSE' ;
       CurNat:='FSE' ; NumTypCpt:=CptANA ;
       Caption:= HMTitre.Mess[1] ;
       EAxe.Visible:=True ; CAxe.Visible:=True ;
       Nature.Visible:=False ; ENatureGene.Visible:=False ;
       FExport.Caption:=MsgBox.Mess[25] ;
       //FExport.Left:=ENatureGene.left-1 ;
       CPTEDEBUT.ZoomTable:=TzSection ; TCpteDebut.Caption:=MsgBox.Mess[26] ;
       CPTEFIN.ZoomTable:=TzSection ;   TCpteFin.Caption:=MsgBox.Mess[27] ;
       HelpContext:=6330000 ;
       GComptes.Caption:=MsgBox.Mess[30] ;
       END else
if lequel='X' then //Auxiliaires
       BEGIN
       FiltreExp:='EXPFTI' ;
       CurNat:='FTI' ; NumTypCpt:=CptAUX ;
       Caption:= HMTitre.Mess[2] ;
       CPTEDEBUT.ZoomTable:=TzTiers;
       CPTEFIN.ZoomTable:=Tztiers ;
       Nature.DataType:='ttNatTiers' ;
       FExport.Caption:=MsgBox.Mess[20] ;
       HelpContext:=6320000 ;
       END else
if lequel='FBJ' then // Journaux budgétaires
       BEGIN
       FiltreExp:='EXPFBJ' ;
       CurNat:='FBJ' ; NumTypCpt:=CptJalBUD ;
       Caption:= HMTitre.Mess[3] ;
       CPTEDEBUT.ZoomTable:=TzBudJal ; TCpteDebut.Caption:=MsgBox.Mess[28] ;
       CPTEFIN.ZoomTable:=TzBudJal ; TCpteFin.Caption:=MsgBox.Mess[29] ;
       EAxe.Visible:=True ; CAxe.Visible:=True ;
       Nature.Visible:=False ; ENatureGene.Visible:=False ;
       Nature.DataType:='ttNatJalBud' ;
       FExport.State:=cbGrayed ; FExport.Visible:=False ;
       //FExport.Caption:=MsgBox.Mess[24] ;
       HelpContext:=6331000 ;
       GComptes.Caption:=MsgBox.Mess[31] ;
       END else
if lequel='FBG' then //Comptes budgétaires
       BEGIN
       FiltreExp:='EXPFBG' ;
       CurNat:='FBG' ; NumTypCpt:=CptBUD ;
       Caption:= HMTitre.Mess[4] ;
       CPTEDEBUT.ZoomTable:=TzBudGen ;
       CPTEFIN.ZoomTable:=TzBudGen ;
       //EAxe.Visible:=True ; CAxe.Visible:=True ;
       Nature.Visible:=False ; ENatureGene.Visible:=False ;
       //Nature.DataType:=ttNatTiers ;
       //FExport.Caption:=MsgBox.Mess[16] ;
       HelpContext:=6332000 ;
       END else
if lequel='FBS' then  //Sections budgétaires
       BEGIN
       FiltreExp:='EXPFBS' ;
       CurNat:='FBS' ; NumTypCpt:=CptSEB ;
       Caption:= HMTitre.Mess[5] ;
       CPTEDEBUT.ZoomTable:=TzBudSec1 ; TCpteDebut.Caption:=MsgBox.Mess[26] ;
       CPTEFIN.ZoomTable:=TzBudSec1 ;   TCpteFin.Caption:=MsgBox.Mess[27] ;
       EAxe.Visible:=True ; CAxe.Visible:=True ; CAxe.ItemIndex:=0 ;
       ENatureGene.Visible:=False ; Nature.Visible:=False ;
       //Nature.DataType:=ttNatTiers ;
       FExport.Caption:=MsgBox.Mess[25] ;
       HelpContext:=6333000 ;
       GComptes.Caption:=MsgBox.Mess[30] ;
       END ;
InitComptes ;
Sauve.Title:=Caption ;
if Nature.visible then Nature.ItemIndex:=0 ;
ChangeDataTypeFmt('-',CurNat,FFormat) ;
GereFMTChoix('-',CurNat,FFormat,FileName,RechFile) ;

ObjFiltre.FFI_TABLE:=FiltreExp;
ObjFiltre.Charger;

BEDI.Enabled:=(FFormat.Value='EDI') ;
UpdateCaption(Self) ;
END ;

{procedure TFExpCpte.BCreerFiltreClick(Sender: TObject);
begin NewFiltre(FiltreExp,FFiltres,Pages) ; end;

procedure TFExpCpte.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre(FiltreExp,FFiltres,Pages) ; end;

procedure TFExpCpte.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre(FiltreExp,FFiltres) ; end;

procedure TFExpCpte.BRenFiltreClick(Sender: TObject);
begin RenameFiltre(FiltreExp,FFiltres) ; end;

procedure TFExpCpte.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres,Pages) ; end;

procedure TFExpCpte.FFiltresChange(Sender: TObject);
begin LoadFiltre(FiltreExp,FFiltres,Pages) ; end; }

Procedure TFExpCpte.InitComptes ;
var Q : TQuery ;
    Tbl,Chp : String ;
BEGIN
Case NumTypCpt of
  CptGEN : BEGIN Tbl:='GENERAUX';Chp:='G_GENERAL' ; END ;
  CptANA : BEGIN Tbl:='SECTION';Chp:='S_SECTION' ; END ;
  CptAUX : BEGIN Tbl:='TIERS';Chp:='T_AUXILIAIRE' ; END ;
  CptJalBUD : BEGIN Tbl:='BUDJAL';Chp:='BJ_BUDJAL' ; END ;
  CptBUD : BEGIN Tbl:='BUDGENE';Chp:='BG_BUDGENE' ; END ;
  CptSEB : BEGIN Tbl:='BUDSECT';Chp:='BS_BUDSECT' ; END ;
  END ;
Q:=OpenSQL('Select '+Chp+' from '+Tbl+' Order by '+Chp,True) ;
If Not Q.Eof then
  BEGIN
  Q.First;CPTEDEBUT.Text:=Q.Fields[0].AsString ;
  Q.Last;CPTEFIN.Text:=Q.Fields[0].AsString ;
  END ;
Ferme(Q) ;
CPTEDEBUT.ExisteH ;CPTEFIN.ExisteH  ;
END ;

procedure TFExpCpte.BFormatClick(Sender: TObject);
begin
ChoixFormatImpExp('-' ,CurNat) ;
If not (ChoixFmt.OkSauve) then Exit ;
With ChoixFmt do
  BEGIN
  if Format<>'' then FFormat.Value:=Format ;
  FFormat.Enabled:=not (FixeFmt) ;
  if Fichier<>'' then FileName.Text:=Fichier ;
  FileName.Enabled:=not (FixeFichier) ;
  RechFile.Enabled:=not (FixeFichier) ;
  END ;
end;

procedure TFExpCpte.FDate1KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFExpCpte.FDate2KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFExpCpte.BEDIClick(Sender: TObject);
begin ParametresEDI('-',CurNat); end;

procedure TFExpCpte.FFormatChange(Sender: TObject);
begin
BEDI.Enabled:=(FFormat.Value='EDI') ;
if FFormat.Value='EDI' then Sauve.FilterIndex:=4 else
  if FFormat.Value='HAL' then Sauve.FilterIndex:=1 else
    if FFormat.Value='SAA' then Sauve.FilterIndex:=3 else
      if FFormat.Value='SAB' then Sauve.FilterIndex:=3 ;
If (Lequel='A') And (FFormat.Value='CGN') Then FSousSection.Visible:=TRUE Else FSousSection.Visible:=FALSE ;   
end;

procedure TFExpCpte.BStopClick(Sender: TObject);
begin Arreter:=True ; end;

function TFExpCpte.TestBreak : Boolean ;
BEGIN
Application.ProcessMessages ;
if Arreter then if MsgBox.Execute(16,caption,'')<>mryes then Arreter:=False ;
Result:=Arreter ;
END ;

{================ Requête de sélection ======================}

Function TFExpCpte.ConstruitReq(OkUpdate : boolean) : boolean ;
var Cpte1,Cpte2 : String ;
    Table, Pref, Compte, St : String ;
BEGIN
Result:=False ;
Cpte1:= CPTEDEBUT.Text ;
Cpte2:= CPTEFIN.Text ;
Case NumTypCpt of
  CptGEN : BEGIN Table:='GENERAUX' ; Pref:='G' ; Compte:='_GENERAL' ; END ;
  CptANA : BEGIN Table:='SECTION' ; Pref:='S' ; Compte:='_SECTION' ; END ;
  CptAUX : BEGIN Table:='TIERS' ; Pref:='T' ; Compte:='_AUXILIAIRE' ; END ;
  CptJalBUD : BEGIN Table:='BUDJAL'; Pref:='BJ' ; Compte:='_BUDJAL' ; END ;
  CptBUD : BEGIN Table:='BUDGENE' ; Pref:='BG' ; Compte:='_BUDGENE' ; END ;
  CptSEB : BEGIN Table:='BUDSECT' ; Pref:='BS' ; Compte:='_BUDSECT' ; END ;
  END ;
if OkUpdate and (Pref='BJ') then BEGIN Result:=True ; Exit ; END ;
if OkUpdate then St:='UPDATE '+Table+' SET '+Pref+'_EXPORTE="X"' else St:='SELECT * FROM '+Table ;
St:=St+' WHERE '+Pref+'_DATEMODIF>="'+USDate(FDate1)+'" AND '+Pref+'_DATEMODIF<="'+USDate(FDate2)+'"' ;
if Cpte1 <>'' then St:=St+' AND '+Pref+Compte+'>="'+Cpte1+'"' ;
if Cpte2 <>'' then St:=St+' AND '+Pref+Compte+'<="'+Cpte2+'"' ;
if (Nature.ItemIndex>0) then
  if (NumTypCpt=CptGEN) then St:=St+' AND '+Pref+'_NATUREGENE="'+NATURE.Value+'"' else
   if (NumTypCpt=CptAUX) then St:=St+' AND '+Pref+'_NATUREAUXI="'+NATURE.Value+'"' else
    if (NumTypCpt=CptJalBUD) then St:=St+' AND '+Pref+'_NATJAL="'+NATURE.Value+'"' ;
if (CAxe.ItemIndex>=0) and ((NumTypCpt=CptJalBUD) or (NumTypCpt=CptSEB) or (NumTypCpt=CptANA)) then St:=St+' AND '+Pref+'_AXE="'+CAxe.Value+'"' ;
if not(FExport.State=cbGrayed) then
  if FExport.State=cbChecked then St:=St+' AND '+Pref+'_EXPORTE="X"' else St:=St+' AND '+Pref+'_EXPORTE<>"X"' ;

if not OkUpdate then St:=St+' ORDER BY '+Pref+Compte ;
(*
Query.Close ;
Query.Sql.Clear ;
Query.Sql.Add(St) ;
//Query.UniDirectional:=(Not OkUpdate) and (V_PGI.Driver<>dbMSACCESS) ;
ChangeSQL(Query) ; //Query.Prepare ;
//PrepareSQLODBC(Query) ;
Query.RequestLive:=False ;
if OkUpdate then Query.ExecSQL else
  BEGIN
  Query.Open ;
  if Query.EOF then
    BEGIN
    If (NumTypCpt<>CptANA) Or (FSousSection.State=cbUnChecked) Then BEGIN MsgBox.Execute(19,caption,'') ; Exit ; END ;
    END ;
  END ;
*)
Result:=True ;
END ;

Procedure TFExpCpte.ReqSQL(OkUpdate : boolean ; Var St : String) ;
var Cpte1,Cpte2 : String ;
    Table, Pref, Compte : String ;
BEGIN
Cpte1:= CPTEDEBUT.Text ;
Cpte2:= CPTEFIN.Text ;
Case NumTypCpt of
  CptGEN : BEGIN Table:='GENERAUX' ; Pref:='G' ; Compte:='_GENERAL' ; END ;
  CptANA : BEGIN Table:='SECTION' ; Pref:='S' ; Compte:='_SECTION' ; END ;
  CptAUX : BEGIN Table:='TIERS' ; Pref:='T' ; Compte:='_AUXILIAIRE' ; END ;
  CptJalBUD : BEGIN Table:='BUDJAL'; Pref:='BJ' ; Compte:='_BUDJAL' ; END ;
  CptBUD : BEGIN Table:='BUDGENE' ; Pref:='BG' ; Compte:='_BUDGENE' ; END ;
  CptSEB : BEGIN Table:='BUDSECT' ; Pref:='BS' ; Compte:='_BUDSECT' ; END ;
  END ;
//if OkUpdate and (Pref='BJ') then BEGIN Result:=True ; Exit ; END ;
if OkUpdate then St:='UPDATE '+Table+' SET '+Pref+'_EXPORTE="X"' else St:='SELECT * FROM '+Table ;
St:=St+' WHERE '+Pref+'_DATEMODIF>="'+USDate(FDate1)+'" AND '+Pref+'_DATEMODIF<="'+USDate(FDate2)+'"' ;
if Cpte1 <>'' then St:=St+' AND '+Pref+Compte+'>="'+Cpte1+'"' ;
if Cpte2 <>'' then St:=St+' AND '+Pref+Compte+'<="'+Cpte2+'"' ;
if (Nature.ItemIndex>0) then
  if (NumTypCpt=CptGEN) then St:=St+' AND '+Pref+'_NATUREGENE="'+NATURE.Value+'"' else
   if (NumTypCpt=CptAUX) then St:=St+' AND '+Pref+'_NATUREAUXI="'+NATURE.Value+'"' else
    if (NumTypCpt=CptJalBUD) then St:=St+' AND '+Pref+'_NATJAL="'+NATURE.Value+'"' ;
if (CAxe.ItemIndex>=0) and ((NumTypCpt=CptJalBUD) or (NumTypCpt=CptSEB) or (NumTypCpt=CptANA)) then St:=St+' AND '+Pref+'_AXE="'+CAxe.Value+'"' ;
if not(FExport.State=cbGrayed) then
  if FExport.State=cbChecked then St:=St+' AND '+Pref+'_EXPORTE="X"' else St:=St+' AND '+Pref+'_EXPORTE<>"X"' ;
if not OkUpdate then St:=St+' ORDER BY '+Pref+Compte ;
END ;

{================ Lancement de l'exportation ======================}

procedure TFExpCpte.AfficheCompteRendu ;
var Mss,MssNb :integer ;
    Reste,Index,Titre : String ;
BEGIN
if NbExport=0 then Exit ;
MssNb:=1 ; Mss:=1 ;
case NumTypCpt of
  CptGEN : ;
  CptAUX : BEGIN Mss:=2 ; MssNb:=3 ; END ;
  CptANA : BEGIN Mss:=3 ; MssNb:=5 ; END ;
  CptJalBUD : BEGIN Mss:=4 ; MssNb:=7 ; END ;
  CptBUD : BEGIN Mss:=5 ; MssNb:=9 ; END ;
  CptSEB : BEGIN Mss:=6 ; MssNb:=11 ; END ;
  END ;
if (NbExport>0) then
  if ChoixFmt.CompteRendu then
    BEGIN
    Reste:=Msg.Mess[0] ; Index:=ReadTokenSt(Reste) ;Titre:=ReadTokenSt(Reste) ;
    Msg.Mess[0]:=Index+';'+caption+';'+Reste;
    if (NbExport>1) then Inc(MssNb) ;
    if NbExport>0 then Msg.Execute(0,IntToStr(NbExport),' '+Msg.Mess[MssNb]) else MsgBox.Execute(19,Caption,'') ;
    END else MsgBox.Execute(Mss+6,Caption,'') ;
END ;

Function TFExpCpte.ExporteFormatParametre : Boolean ;
var Entete   : TFmtEntete ;
    Detail   : TTabFmtDetail ;
    Debut    : Boolean ;
    Fichier : TextFile ;
    Q : TQuery ;
    SQL : String ;
BEGIN
  ReqSQL(FALSE,SQL) ;
  Q:=OpenSQL(SQL,TRUE) ;
  try
    if Q.Eof then begin
      Result := False;
      {JP 18/07/05 : Ajout d'un message d'avertissment}
      HShowMessage('0;' + Caption + ';Aucune donnée à traiter.;I;O;O;O;', '', '');
      Exit;
    end;

    Result:=TRUE ;
    {JP 18/07/05 : ajout d'un bloc finally, car avant si Q était vide on faisait un CloseFile
                   alors ChargeFormat n'avait pas été exécuté => Erreur E/S 103}
    try
      if ChargeFormat(Fichier,Filename.Text,CurNat,'-',FFormat.Value,Entete,Detail,Debut) then
        BEGIN
        InitMove(RecordsCount(Q),'') ;
        While not Q.Eof do
          BEGIN
          EcrireFormat(Fichier,Entete,Detail,Debut,Q) ;
          Q.Next ;
          Inc(NbExport) ; MoveCur(False) ;
          END ;
        FiniMove ;
        END ;
    finally
      CloseFile(Fichier) ;
    end ;
  finally
    Ferme(Q);
  end;
END ;

Procedure TFExpCpte.MajFlagExport ;
Var SQL : String ;
    Q : TQuery ;
BEGIN
If NumTypCpt=CptJalBUD Then Exit ;
ReqSQL(TRUE,SQL) ;
ExecuteSQL(SQL) ;
END ;

Procedure TFExpCpte.TraiteExport ;
Var F : TextFile ;
    ExpParam : Boolean ;
BEGIN
OkExp:=FALSE ;
NbExport:=0 ;
if FFormat.Value='SAB' then Lequel:='T' ;
ExpParam:=((FFormat.Value<>'SAA') and (FFormat.Value<>'SAB') and (FFormat.Value<>'EDI')
          and (FFormat.Value<>'HAL') and (FFormat.Value<>'CLB')and (FFormat.Value<>'CPR')
          and (FFormat.Value<>'CGN') and (FFormat.Value<>'CVI'))
          or (Lequel='FBG') or (Lequel='FBS') or (Lequel='FBJ') ;
if not ConstruitReq(False) then exit ;
if ExpParam then
  BEGIN
  OkExp:=ExporteFormatParametre ;
  END else
  BEGIN
  AssignFile(F,FileName.Text) ;
  {$I-} ReWrite (F) ; {$I+}
  if IoResult<>0 then BEGIN Msgbox.Execute(4,caption,'') ; Exit ; END ;
  CloseFile(F) ;
  ActivePanels(Self,False,False) ;
  SourisSablier ;
  If FFormat.Value='EDI' then
    BEGIN
    OkExp:=LanceExportPlanEDI ;
    END else
    BEGIN
    if FFormat.Value='SAB' then OkExp:=ExporteTiers else
     if FFormat.Value='SAA' then OkExp:=ExporteCpteSAARI else
      if FFormat.Value='HAL' then OkExp:=ExporteHalley Else
        if FFormat.Value='CGN' then OkExp:=ExporteCptCEGID ;

    END ;
  END ;
END ;

procedure TFExpCpte.BValiderClick(Sender: TObject);
Var Ok,ExpParam : Boolean ;
BEGIN
if (FFormat.ItemIndex=-1) then BEGIN MsgBox.Execute(15,caption,'') ; Exit ; END ;
if not (IsValidDate(FDate1.Text) and IsValidDate(FDate2.Text)) then BEGIN Msgbox.Execute(13,caption,'') ; Exit ; END ;
if (CPTEDEBUT.Text='') or (CPTEFIN.Text='') then BEGIN Msgbox.Execute(14,caption,'') ; Exit ; END ;
if FileName.Text='' then BEGIN Msgbox.Execute(3,caption,'') ; Exit ; END ;
Ok:=TRUE ; OkExp:=FALSE ;
if (Transactions(TraiteExport,1)<>OeOk) then BEGIN MessageAlerte(MsgBox.Mess[17]) ; Ok:=False ; END ;
If OkExp Then Transactions(MajFlagExport,1) ;
(* GP le 23/10/98
if not ConstruitReq(False) then exit ;
Ok:=True ;
if ExpParam then
  BEGIN
  ExporteFormatParametre ;
  END else
  BEGIN
  AssignFile(F,FileName.Text) ;
  {$I-} ReWrite (F) ; {$I+}
  if IoResult<>0 then BEGIN Msgbox.Execute(4,caption,'') ; Exit ; END ;
  CloseFile(F) ;
  ActivePanels(Self,False,False) ;
  SourisSablier ;
  If FFormat.Value='EDI' then
    BEGIN
    if (Transactions(LanceExportPlanEDI,2)<>OeOk) then BEGIN MessageAlerte(MsgBox.Mess[17]) ; Ok:=False ; END ;
    END else
    BEGIN
    if FFormat.Value='SAB' then BEGIN if (Transactions(ExporteTiers,2)<>OeOk) then BEGIN Ok:=False ; MessageAlerte(MsgBox.Mess[14]) ; END ; END else
     if FFormat.Value='SAA' then BEGIN if (Transactions(ExporteCpteSAARI,2)<>OeOk) then BEGIN Ok:=False ; MessageAlerte(MsgBox.Mess[14]) ; END ; END else
      if FFormat.Value='HAL' then if (Transactions(ExporteHalley,2)<>OeOk) then BEGIN Ok:=False ; MessageAlerte(MsgBox.Mess[14]) ; END else
    END ;
  END ;
ConstruitReq(True) ;
*)
ActivePanels(Self,True,False) ;
SourisNormale ;
if Ok then AfficheCompteRendu ;
Screen.cursor:=SynCrDefault ;
END;

{================ Format SAARI  ======================}

Function TFExpCpte.ExporteCpteSAARI : Boolean ;
Var QSOC : TQuery ;
    St,Pref : String ;
    F       : TextFile ;
    Regroup :String ;
    Coll,Point,Cent,vent,Lett,Niv,Nat : String[1] ;
    NumCompte,Collectif,Sens :String ;
    TypeCpte : String ;
    Q : TQuery ;
    SQL : String ;
BEGIN
ReqSQL(FALSE,SQL) ;
Q:=OpenSQL(SQL,TRUE) ; If Q.Eof Then BEGIN MsgBox.Execute(19,caption,'') ; Result:=FALSE ; Exit ; END ;
Result:=TRUE ;
TypeCpte:=Lequel[1] ;//Lg:=0 ;
InitMove(RecordsCount(Q)+1,'') ;
Movecur(False) ;
AssignFile(F,FileName.Text) ;
{$I-} ReWrite (F) ; {$I+}
if IoResult<>0 then BEGIN CloseFile(F) ; Exit ; END ;
{$IFDEF SPEC302}
QSoc:=OpenSQL('Select SO_LIBELLE From SOCIETE',True) ;
if not QSOC.Eof then BEGIN QSOC.First ; St:=Format_String(QSOC.Fields[0].AsString,70) ; END ;
Ferme(QSOC) ;
{$ELSE}
St:=GetParamSoc('SO_LIBELLE') ; St:=Format_String(St,30) ;
{$ENDIF}
//St:=St+FFormat.Value+Lequel[1] ;
Writeln(F,St) ;
While not Q.Eof do
  BEGIN
  if TestBreak then Break ; Inc(NbExport) ;
  St:='' ;Point:='';Lett:='';Coll:='';Cent:='' ;Vent:='' ;
  NumCompte:=Format_String(Q.Fields[0].AsString,13);
  If NumTypCpt=CptAUX Then St:=TypeCpte+NumCompte+Format_String(Copy(Q.Fields[2].AsString,1,31),31)
                   Else St:=TypeCpte+NumCompte+Format_String(Copy(Q.Fields[1].AsString,1,31),31) ;
  Case NumTypCpt of
    CptGEN,CptANA :
      BEGIN
      if TypeCpte='G' then
        BEGIN
        Pref:='G' ;//Lg:=11 ;
        if Q.FindField('G_POINTABLE').AsString='X' then Point:='P' ;
        if Q.FindField('G_LETTRABLE').AsString='X' then Lett:='L' ;
        if Q.FindField('G_COLLECTIF').AsString='X' then Coll:='C' ;
        if Q.FindField('G_CENTRALISABLE').AsString='X' then Cent:='E' ;
        if Q.FindField('G_VENTILABLE').AsString='X' then Vent:='V' ;
        Collectif:=' ' ;
        Nat:=' ' ;
        if Q.FindField('G_CONFIDENTIEL').AsString='X' then Niv:='1' else Niv:='0' ;
        END Else
        BEGIN
        Pref:='S' ;//Lg:=7 ;
        Nat:=' ' ;
        if Q.FindField('S_CONFIDENTIEL').AsString='X' then Niv:='1' else Niv:='0' ;
        END ;
      Regroup:=Format_String(' ',4)+Format_String(Collectif,13)+Format_String(Q.FindField(Pref+'_SENS').AsString,1)+Format_String(Nat,1) ;
      St:=St+Regroup+Format_string(Coll+Point+Cent+vent+Lett,4)+Format_String(Niv,1) ;
      END ;
    CptAUX :
      BEGIN
      Pref:='T' ;//Lg:=5 ;
      Sens:=RecupChp(Q.Findfield('T_COLLECTIF').AsString,'G_SENS') ;
      if Q.Findfield('T_LETTRABLE').AsString='X' then Lett:='L' ;
      if RecupChp(Q.Findfield('T_COLLECTIF').AsString,'G_POINTABLE')='X' then Point:='P' ;
      Nat:=' ' ; //Q.Findfield('T_NATUREAUXI').AsString ;
      if (Q.FindField('T_CONFIDENTIEL').AsString='X') then  Niv:='1' else Niv:='0' ;
      Regroup:=Sens+Format_String(Lett,1)+Format_String(Point,1)+Format_String(Niv,1)+Format_String(Copy(Q.FindField('T_COLLECTIF').AsString,1,13),13)
       +Format_String(Copy(Q.FindField('T_ADRESSE1').AsString,1,21),21)+Format_String(Copy(Q.FindField('T_ADRESSE2').AsString,1,21),21)
       +Format_String(Copy(Q.FindField('T_CODEPOSTAL').AsString,1,6),6)+Format_String(Copy(Q.FindField('T_VILLE').AsString,1,21),21)+Format_String(Copy(Q.FindField('T_TELEPHONE').AsString,1,15),15) ;
      St:=St+Regroup ;
      END ;
    END ;
  if ChoixFmt.Ascii then St:=Ansi2Ascii(St) ;
  Writeln(F,St) ;
  //Q.Edit ; Q.findField(Pref+'_EXPORTE').AsString:='X' ; Q.Post ;
  Movecur(False) ;
  Q.Next ;
  END ;
CloseFile(F) ;
Ferme(Q) ;
FiniMove ;
END ;

function TFExpCpte.RecupChp(Leq : String ; Quoi : String) : String ;
var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT '+Quoi+' FROM GENERAUX WHERE G_GENERAL="'+Leq+'"',True) ;
If Not Q.EOF then Result:=Q.Fields[0].AsString ;
Ferme(Q) ;
END ;

{================ Format SAARI Banque ======================}

Function TFExpCpte.ExporteTiers : Boolean ;
Var Q1,QSoc:  TQuery ;
    F : TextFile ;
    St : String ;
    Q : tQuery ;
    SQL : String ;
BEGIN
// Il manque le code sélection, code pays, lieu de compens (non utilisé dans SV7 ?)
(* GP le 23/08/1998
St:='SELECT T_EAN, T_LIBELLE, T_AUXILIAIRE, T_COLLECTIF, T_ADRESSE1, T_ADRESSE2, T_CODEPOSTAL, T_VILLE, T_PAYS, '
   +' T_EXPORTE From TIERS WHERE '
   +' T_DATEMODIF>="'+USDate(FDate1)+'" AND T_DATEMODIF<="'+USDate(FDate2)+'"'
   +' AND T_AUXILIAIRE>="'+CPTEDEBUT.Text+'" AND T_AUXILIAIRE<="'+CPTEFIN.Text+'"' ;
if not(FExport.State=cbGrayed) then
  if FExport.State=cbChecked then St:=St+' AND T_EXPORTE="X"' else St:=St+' AND T_EXPORTE<>"X"' ;
if (Nature.ItemIndex>0) then St:=St+' AND T_NATUREAUXI="'+NATURE.Value+'"' ;
Query.Close ;
Query.SQL.Clear ;
Query.SQL.Add(St) ;
Query.UniDirectional:=False ;
ChangeSQL(Query) ; Query.Prepare ;
Query.RequestLive:=True ;
Query.Open ;
*)
ReqSQL(FALSE,SQL) ;
Q:=OpenSQL(SQL,TRUE) ; If Q.Eof Then BEGIN MsgBox.Execute(19,caption,'') ; Result:=FALSE ; Exit ; END ;
Result:=TRUE ;
InitMove(RecordsCount(Q)+1,'') ;
Movecur(False) ;
AssignFile(F,FileName.Text) ;
{$I-} ReWrite (F) ; {$I+}
if IoResult<>0 then exit ;
{$IFDEF SPEC302}
QSoc:=OpenSQL('Select SO_LIBELLE From SOCIETE',True) ;
if not QSOC.Eof then BEGIN QSOC.First ; St:=Format_String(QSOC.Fields[0].AsString,70) ; END ;
Ferme(QSOC) ;
{$ELSE}
St:=GetParamSoc('SO_LIBELLE') ; St:=Format_String(St,70) ;
{$ENDIF}
St:=St+FFormat.Value+Lequel[1] ;
Writeln(F,St) ;
While not Q.Eof do
  BEGIN
  if TestBreak then Break ; Inc(NbExport) ;
  Q1:=OpenSQL('SELECT R_DOMICILIATION , R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB FROM RIB WHERE R_AUXILIAIRE="'+Q.FindField('T_AUXILIAIRE').AsString+'" AND R_PRINCIPAL="X"',True) ;
  St:=Format_String(Copy(Q.FindField('T_EAN').AsString,1,13),13)+
      Format_String(Copy(Q.FindField('T_LIBELLE').AsString,1,24),24)+
      Format_String(Copy(Q.FindField('T_AUXILIAIRE').AsString,1,13),13)+
      Format_String(Copy(Q.FindField('T_COLLECTIF').AsString,1,13),13)+
      Format_String(Copy(Q.FindField('T_ADRESSE1').AsString,1,24),24)+
      Format_String(Copy(Q.FindField('T_ADRESSE2').AsString,1,24),24)+
      Format_String(Copy(Q.FindField('T_CODEPOSTAL').AsString,1,8),8)+
      Format_String(Copy(Q.FindField('T_VILLE').AsString,1,24),24)+'  '+
      Format_String(Copy(Q1.Fields[0].AsString,1,20),20)
    +Format_String(Copy(Q1.Fields[1].AsString,1,5),5)+Format_String(Copy(Q1.Fields[2].AsString,1,5),5)
    +Format_String(Copy(Q1.Fields[3].AsString,1,11),11)+Format_String(Copy(Q1.Fields[4].AsString,1,2),2)
    +' '+Format_String(Copy(Q.FindField('T_PAYS').AsString,1,3),3) ;
  if ChoixFmt.Ascii then St:=Ansi2Ascii(St) ;
  Ferme(Q1) ;
  Writeln(F,St) ;
  (*
  Query.Edit ; Query.FindField('T_EXPORTE').AsString:='X'+Copy(Query.FindField('T_EXPORTE').AsString,2,5) ; Query.Post ;
  *)
  Q.Next ;
  Movecur(False) ;
  END ;
CloseFile(F) ;
Ferme(Q) ;
FiniMove ;
END ;

{================= Exportation des comptes Halley =====================}

Function TFExpCpte.ExporteHalley : Boolean ;
var QSoc : TQuery ;
    St,Pref : String ;
    Q : TQuery ;
    SQL : String ;
BEGIN
ReqSQL(FALSE,SQL) ;
Q:=OpenSQL(SQL,TRUE) ; If Q.Eof Then BEGIN MsgBox.Execute(19,caption,'') ; Result:=FALSE ; Exit ; END ;
Result:=TRUE ;
Case NumTypCpt of
     CptGEN : Pref:='G' ;
     CptANA : Pref:='S' ;
     CptAUX : Pref:='T' ;
     CptBUD : Pref:='FBG' ;
     CptSEB : Pref:='FBS' ;
     END ;
InitMove(RecordsCount(Q)+1,'') ;
Movecur(False) ;
AssignFile(Fichier,FileName.Text) ;
{$I-} ReWrite (Fichier) ; {$I+}
if IoResult<>0 then BEGIN CloseFile(Fichier) ; Exit ; END ;
{$IFDEF SPEC302}
QSoc:=OpenSQL('Select SO_LIBELLE From SOCIETE',True) ;
if not QSOC.Eof then BEGIN QSOC.First ; St:=Format_String(QSOC.Fields[0].AsString,70) ; END ;
Ferme(QSOC) ;
{$ELSE}
St:=GetParamSoc('SO_LIBELLE') ; St:=Format_String(St,70) ;
{$ENDIF}
St:=St+FFormat.Value+Lequel[1] ;
Writeln(Fichier,St) ;
While not Q.Eof do
  BEGIN
  if TestBreak then Break ; Inc(NbExport) ;
  St:=RecupValChp(Q) ;
  if ChoixFmt.Ascii then St:=Ansi2Ascii(St) ;
  Writeln(Fichier,St) ;
  Movecur(False) ;
  Q.Next ;
  END ;
CloseFile(Fichier) ;
Ferme(Q) ;
FiniMove ;
END ;

function TFExpCpte.RecupValChp(Q : TQuery) : String ;
var i : integer ;
    St : string ;
    Val : String ;
    St1 : String ;
BEGIN
St:='' ;
for i:=0 to Q.FieldCount-1 do
  BEGIN
  Case Q.Fields[i].DataType of
    ftString  : Val:=Format_String(Copy(Q.Fields[i].AsString,1,Q.Fields[i].DataSize-1),Q.Fields[i].DataSize-1) ;
    ftBlob : Continue ;
    ftDate,ftDateTime : Val:=FormatDateTime('dd/mm/yyyy',Q.Fields[i].AsDateTime) ;
    ftTime : Val:=FormatDateTime('hh:mm:ss',Q.Fields[i].AsDateTime) ;
    //ftDateTime : Val:=FormatDateTime('dd/mm/yyyy hh:mm:ss',Q.Fields[i].AsDateTime) ;
    ftAutoInc,ftInteger,ftSmallInt,ftWord : Val:=AlignDroite(IntToStr(Q.Fields[i].AsInteger),9) ;
    ftBCD,ftCurrency,ftFloat : BEGIN
                               St1:=AlignDroite(StrfMontant(Q.Fields[i].AsFloat,15,2,'',False),15) ;
                               If Length(St1)>15 Then St1:=AlignDroite(StrfMontant(0,15,2,'',False),15) ;
                               Val:=AlignDroite(StrfMontant(Q.Fields[i].AsFloat,15,2,'',False),15) ;
                               END ;
    ftBoolean : Val:=Copy(Q.Fields[i].AsString,1,1) ;
    else continue ;
    END ;
  if (Length(St+Val)>1024) then
    BEGIN
    St:=St+Format_String(' ',1024-Length(St)-1)+SepLigneIE ;
    Writeln(Fichier,St) ;
    St:='' ;
    END ;
  St:=St+Val ;
  END ;
Result:=St ;
END ;

{================ Format EDIFICAS ======================}

Function TFExpCpte.LanceExportPlanEDI : Boolean ;
var FTMP : TextFile ;
    St : String ;
    Fic:String;
    Q : TQuery ;
    SQL : String ;
begin
ReqSQL(FALSE,SQL) ;
Q:=OpenSQL(SQL,TRUE) ; If Q.Eof Then BEGIN MsgBox.Execute(19,caption,'') ; Result:=FALSE ; Exit ; END ;
Result:=TRUE ;
If not (Dossier.OkSauve) then InitDossier('-',CurNat) ;
Fic:=FileName.Text ;
Application.ProcessMessages ;
AssignFile(Fichier,Fic) ;
{$I-} Rewrite(Fichier) ; {$I+}
Entete ; Detail(Q) ; Resume ;
CloseFile(Fichier) ;
// Conversion Ascii...
If ChoixFmt.Ascii then
   BEGIN
   AssignFile(Fichier,Fic) ;
   {$I-} Reset(Fichier) ; {$I+}
   AssignFile(FTMP,ExpandFileName(ExtractFileName(Fic)+'.$$$')) ;
   {$I-} ReWrite(FTMP) ; {$I+}
   While Not EOF(Fichier) do
     BEGIN
     Readln(Fichier,St) ;
     if ChoixFmt.Ascii then St:=ANSI2ASCII(St) ;
     WriteLn(FTMP,St) ;
     END ;
  CloseFile(Fichier) ;
  CloseFile(FTMP) ;
  Erase(Fichier) ;
  RenameFile(ExpandFileName(ExtractFileName(Fic)+'.$$$'),Fic) ;
  END ;
END ;

procedure TFExpCpte.Entete ;
var Q : TQuery ;
BEGIN
Q:=Nil ;
With Dossier do
  Writeln(Fichier, '02010'+'UNOC'+'1'+Format_String(EditEmet,35)+'OEC '+
                   Format_String(EditAdrEmet,14)+Format_String(EditRec,35)+
                   'OEC '+Format_String(EditAdrRec,14)+FormatDateTime('yymmddhhnn',NowH)+
                   Format_String(EditRefIntChg,14)+CBoxAccRec,CBoxEssai) ;
  Case NumTypCpt of
    cptGEN : Q:=OpenSQL('Select MIN(G_DATECREATION) AS DATCRE, MAX(G_DATEFERMETURE) AS DATFER '+
                           'from GENERAUX',True) ;
    cptAUX : Q:=OpenSQL('Select MIN(T_DATECREATION) AS DATCRE, MAX(T_DATEFERMETURE) AS DATFER '+
                           'from TIERS',True) ;
    cptANA : Q:=OpenSQL('Select MIN(S_DATECREATION) AS DATCRE, MAX(S_DATEFERMETURE) AS DATFER '+
                           'from SECTION',True) ;
    cptBUD : Q:=OpenSQL('Select MIN(BG_DATECREATION) AS DATCRE, MAX(BG_DATEFERMETURE) AS DATFER '+
                           'from BUDGENE',True) ;
    END ;
  With dossier do  Write(Fichier, '02030'+Format_String(EditRefIntChg,14)+'CHACCO'+'  2'+'  1'+'RT');
  Case NumTypCpt of
    cptGEN : Write(Fichier,'PCG') ;
    cptAUX : Write(Fichier,'PCD') ;
    cptANA : Write(Fichier,'PCA') ;
    cptBUD : Write(Fichier,'PCB') ;
    END ;
  With dossier Do
    BEGIN
    Writeln(Fichier,'CPT'+'PEE'+Format_String(EditIdentEnv,17)+Blanc(3)+'DPP'+
          QD(Q.FindField('DATCRE').asDateTime,'yyyymmdd')+QD(Q.FindField('DATFER').AsDateTime,'yyyymmdd')+'718'+'242'+
          Format_String(FormatDateTime('yyyymmddhhnn',NowH),16)+'203') ;
  //Writeln(Fichier,'02050'+'TPR'+'PC '+'PEE'+FILLER{5719}+'PC '+'CNC'+'TPV'{1154}) ;
    Writeln(Fichier,'02110'+'ACF'+Format_String(DEditSiret,17)+'100'+'107'+
          Format_String(DEditPerChrg,35)+Format_String(DEditRue1,35)+
          Format_String(DEditRue2,35)+Format_String(DEditVille,35)+
          Format_String(DEditDiv,9)+Format_String(DEditCodPost,9)+
          Format_String(DEditPays,2)) ;
    DEditModCom:='TE' ;
    Writeln(Fichier, '02115'+'ACF'+Format_String(DEditFctCont,3)+Format_String(DEditPerChrg,35)+
          Format_String(DEditNumCom,25)+Format_String(DEditModCom,3)+'AAA'+
          Format_String(EditRefDem,35));
    Writeln(Fichier, '02130'+'ACS'+Format_String(EEditSiret,17)+'100'+'107'+
          Format_String(EEditPerChrg,35)+Format_String(EEditRue1,35)+
          Format_String(EEditRue2,35)+Format_String(EEditVille,35)+
          Format_String(EEditDiv,9)+Format_String(EEditCodPost,9)+
          Format_String(EEditPays,2)) ;
    EEditModCom:='TE' ;
    Writeln(Fichier, '02135'+'ACS'+Format_String(EEditFctCont,3)+Format_String(EEditPerChrg,35)+
          Format_String(EEditNumCom,25)+Format_String(EEditModCom,3)+'ASV'+
          Format_String(EditNumDosEmt,35)) ;
    Writeln(Fichier, '02150'+'ACR'+Format_String(REditSiret,17)+'100'+'107'+
          Format_String(RComboPerChrg,35)+Format_String(REditRue1,35)+
          Format_String(REditRue2,35)+Format_String(REditVille,35)+
          Format_String(REditDiv,9)+Format_String(REditCodPost,9)+
          Format_String(REditPays,2)) ;
    REditModCom:='TE' ;
    Writeln(Fichier, '02155'+'ACR'+Format_String(REditFctCont,3)+Format_String(RComboPerChrg,35)+
          Format_String(REditNumCom,25)+Format_String(REditModCom,3)+'ARV'+
          Format_String(EditNumDosRec,35)) ;
    END ;
Ferme(Q) ;
END ;

procedure TFExpCpte.Detail(Q : TQuery) ;
var D,Pref : String ;
BEGIN
Quantite:=0;
if Q.EOF then Exit ;
InitMove(RecordsCount(Q)+1,'') ;
MoveCur(False) ;
while not Q.EOF do
  BEGIN
  if TestBreak then Break ; Inc(NbExport) ;
  Case NumTypCpt of
    cptGEN : BEGIN
             Pref:='G' ;
             if (QS(Q.FindField('G_GENERAL').AsString,17)>=Format_String(CPTEDEBUT.Text,17))
                   and (QS(Q.FindField('G_GENERAL').AsString,17)<=Format_String(CPTEFIN.Text,17)) then
                   BEGIN
                   Inc(Quantite) ;
                   Compte:=Copy(QS(Q.FindField('G_GENERAL').AsString,0),1,12) ;
                   Write(Fichier,'02200'+'ECG'+Format_String(Compte,17)+'PC '+'PEE'+QS(Q.FindField('G_ABREGE').AsString,17)+
                         QS(Q.FindField('G_LIBELLE').AsString,35)+QS(Q.FindField('G_LIBELLE').AsString,35)+Blanc(3)) ;
                   if QS(Q.FindField('G_FERME').AsString,0)='X' then D:='DOF'+QD(Q.FindField('G_DATECREATION').AsDateTime,'yyyymmdd')+QD(Q.FindField('G_DATEFERMETURE').AsDateTime,'yyyymmdd')+'718'
                                                                    else D:='DOC'+QD(Q.FindField('G_DATECREATION').AsDateTime,'yyyymmdd')+'102' ;
                   Write(Fichier,D);
                   Writeln(Fichier,Blanc(3)+Format_String(Dossier.EditTxt,35)) ;
                   Writeln(Fichier,'02250'+QS(Q.FindField('G_GENERAL').AsString,17)+'VAT'+Blanc(7)+Blanc(17)+Blanc(6)+
                           Blanc(3)+Blanc(3)+Blanc(6)+Blanc(3)+Blanc(3)+Blanc(6)+Blanc(3)+Blanc(3)) ;
                         //'FJE'+'107'+FILLER+FILLER{5718}+'CAE'+'107'+Blanc(6)+'TCE'+'GTC') ;
                   SousDetail(QS(Q.FindField('G_NATUREGENE').AsString,3),Q) ;
                   if QS(Q.FindField('G_CENTRALISABLE').AsString,1)='X' then SousDetail('CCE',Q) ;
                   if QS(Q.FindField('G_LETTRABLE').AsString,1)='X' then
                      BEGIN
                      SousDetail('CLE',Q) ;
                      //SousDetail('CEC') ;
                      END ;
                   if (QS(Q.FindField('G_POINTABLE').AsString,1)='X') then SousDetail('CPO',Q) ;
                   //if (QS(Q.FindField('G_VENTILABLE').AsString,1)='X') then SousDetail('CVE') ;
                   if (QS(Q.FindField('G_COLLECTIF').AsString,1)='X') then SousDetail('CCO',Q) ;
                   //if (QS(Q.FindField('G_SENS').AsString,1)='C') then SousDetail('CCS') else
                   //  if (QS(Q.FindField('G_SENS').AsString,1)='D') then SousDetail('CDS') else SousDetail('CCD') ;
                   //if QS(Q.FindField('G_BUDGENE').AsString,0)<>'' then SousDetail('CBU') ;
                   //if (QS(Q.FindField('G_CORRESP1').AsString,0)<>'')or(QS(Q.FindField('G_CORRESP2').AsString,0)<>'') then SousDetail('COR') ;
                   END ;
             END ;
    cptAUX : BEGIN
             Pref:='T' ;
             if (QS(Q.FindField('T_AUXILIAIRE').AsString,17)>=Format_String(CPTEDEBUT.Text,17))
                   and (QS(Q.FindField('T_AUXILIAIRE').AsString,17)<=Format_String(CPTEFIN.Text,17))then
                   BEGIN
                   Inc(Quantite) ;
                   Write(Fichier,'02200'+'ECD'+QS(Q.FindField('T_AUXILIAIRE').AsString,17)+'PC '+'PEE'+QS(Q.FindField('T_ABREGE').AsString,17)+
                         QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_DEVISE').AsString,3)) ;
                   if QS(Q.FindField('T_FERME').AsString,0)='X' then Write(Fichier, 'DOF') else Write(Fichier, 'DOC') ;
                   Write(Fichier,QD(Q.FindField('T_DATECREATION').AsDateTime,'yyyymmdd')+QD(Q.FindField('T_DATEFERMETURE').AsDateTime,'yyyymmdd'));
                   if QS(Q.FindField('T_FERME').AsString,0)='X' then Write(Fichier, '718') else Write(Fichier, '102') ;
                   Writeln(Fichier,Blanc(3)+Format_String(Dossier.EditTxt,35)) ;
                   Writeln(Fichier,'02250'+QS(Q.FindField('T_AUXILIAIRE').AsString,17)+'VAT'+Blanc(7)+QS(Q.FindField('T_ABREGE').AsString,17)+Blanc(6){5718}+
                           Blanc(3)+Blanc(3)+Blanc(6)+Blanc(3)+Blanc(3)+Blanc(6)+Blanc(3)+Blanc(3)) ;
                         //'FJE'+'107'+FILLER+FILLER{5718}+'CAE'+'107'+Blanc(6)+'TCE'+'GTC') ;
                   SousDetail(QS(Q.FindField('T_NATUREAUXI').AsString,3),Q) ;
                   if QS(Q.FindField('T_LETTRABLE').AsString,1)='X' then
                      BEGIN
                      SousDetail('CLE',Q) ;
                      //SousDetail('CEC') ;
                      END ;
                   //if (QS(Q.FindField('T_CORRESP1').AsString,0)<>'')or(QS(Q.FindField('T_CORRESP2').AsString,0)<>'') then SousDetail('COR') ;
                   if RecupChp(Q.Findfield('T_COLLECTIF').AsString,'G_POINTABLE')='X' then SousDetail('CPO',Q) ;
                   if RecupChp(Q.Findfield('T_COLLECTIF').AsString,'G_CENTRALISABLE')='X' then SousDetail('CCE',Q) ;
                   Writeln(Fichier,'02550'+QS(Q.FindField('T_AUXILIAIRE').AsString,17)+'CPT'+Format_String(Dossier.REditSiret,17)+
                           '100'+'107'+Format_String(Dossier.REditNom,35)+Format_String(Dossier.REditRue1,35)+
                           Format_String(Dossier.REditRue2,35)+Format_String(Dossier.REditVille,35)+
                           Format_String(Dossier.REditDiv,9)+Format_String(Dossier.REditCodPost,9)+
                           Format_String(Dossier.REditPays,2)) ;
                   Writeln(Fichier,'02600'+QS(Q.FindField('T_AUXILIAIRE').AsString,17)+Format_String(Dossier.REditFctCont,3)+
                           Format_String(Dossier.RComboPerChrg,35)+Format_String(Dossier.REditNumCom,25)+
                           Format_String(Dossier.REditModCom,3)+'FC ') ;
                   END ;
             END ;
    cptANA : BEGIN
             Pref:='S' ;
             if (QS(Q.FindField('S_SECTION').AsString,17)>=Format_String(CPTEDEBUT.Text,17))
                   and (QS(Q.FindField('S_SECTION').AsString,17)<=Format_String(CPTEFIN.Text,17))then
                  BEGIN
                  Inc(Quantite) ;
                  Write(Fichier,'02200'+'ECA'+QS(Q.FindField('S_SECTION').AsString,17)+'PC '+'PEE'+QS(Q.FindField('S_ABREGE').AsString,17)+
                        QS(Q.FindField('S_LIBELLE').AsString,35)+QS(Q.FindField('S_LIBELLE').AsString,35)+Blanc(3)) ;
                  if QS(Q.FindField('S_FERME').AsString,0)='X' then Write(Fichier, 'DOF') else Write(Fichier, 'DOC') ;
                  Write(Fichier, QD(Q.FindField('S_DATECREATION').AsDateTime,'yyyymmdd')+QD(Q.FindField('S_DATEFERMETURE').AsDateTime,'yyyymmdd')) ;
                  if QS(Q.FindField('S_FERME').AsString,0)='X' then Write(Fichier, '718') else Write(Fichier, '102') ;
                  Writeln(Fichier, Blanc(3)+Format_String(Dossier.EditTxt,35)) ;
                  if QS(Q.FindField('S_SENS').AsString,1)='C'
                     then SousDetail('CCS',Q)
                     else if QS(Q.FindField('S_SENS').AsString,1)='D'
                             then SousDetail('CDS',Q)
                             else SousDetail('CCD',Q) ;
                  if (QS(Q.FindField('S_CORRESP1').AsString,0)<>'')or(QS(Q.FindField('S_CORRESP2').AsString,0)<>'')
                     then SousDetail('COR',Q) ;
                  END ;
             END ;
    cptBUD : BEGIN
             Pref:='BG' ;
             if (QS(Q.FindField('BG_BUDGENE').AsString,17)>=Format_String(CPTEDEBUT.Text,17))
                   and (QS(Q.FindField('BG_BUDGENE').AsString,17)<=Format_String(CPTEFIN.Text,17))then
                  BEGIN
                  Inc(Quantite) ;
                  Write(Fichier,'02200'+'ECB'+QS(Q.FindField('BG_BUDGENE').AsString,17)+'PC '+'PEE'+QS(Q.FindField('BG_ABREGE').AsString,17)+
                        QS(Q.FindField('BG_LIBELLE').AsString,35)+QS(Q.FindField('BG_LIBELLE').AsString,35)+Blanc(3)) ;
                  if QS(Q.FindField('BG_FERME').AsString,0)='X' then Write(Fichier, 'DOF') else Write(Fichier, 'DOC') ;
                  Write(Fichier, QD(Q.FindField('BG_DATECREATION').AsDateTime,'yyyymmdd')+QD(Q.FindField('BG_DATEFERMETURE').AsDateTime,'yyyymmdd'));
                  if QS(Q.FindField('BG_FERME').AsString,0)='X' then Write(Fichier, '718') else Write(Fichier, '102') ;
                  Writeln(Fichier,Blanc(3)+Format_String(Dossier.EditTxt,35)) ;
                  Writeln(Fichier,'02250'+QS(Q.FindField('BG_BUDGENE').AsString,17)+'VAT'+Blanc(7)+QS(Q.FindField('BG_ABREGE').AsString,17)+Blanc(6){5718}+
                          Blanc(3)+Blanc(3)+Blanc(6)+Blanc(3)+Blanc(3)+Blanc(6)+Blanc(3)+Blanc(3)) ;
                        //'FJE'+'107'+FILLER+FILLER{5718}+'CAE'+'107'+Blanc(6)+'TCE'+'GTC') ;
//Simon                  if QS(Q.FindField('B_VENTILABLE',1)='O' then SousDetail('CVE') ;
                  if QS(Q.FindField('BG_SENS').AsString,1)='C'
                     then SousDetail('CCS',Q)
                     else if QS(Q.FindField('BG_SENS').AsString,1)='D'
                             then SousDetail('CDS',Q)
                             else SousDetail('CCD',Q) ;
//Simon                  if (QS(Q.FindField('B_CORRESP1',0)<>'')or(QS(Q.FindField('B_CORRESP2',0)<>'')
//                     then SousDetail('COR') ;
                  END ;
             END ;
     END ;
  //Q.Edit ; Q.FindField(Pref+'_EXPORTE').AsString:='X' ; Q.Post ;
  Q.Next ;
  MoveCur(False) ;
  END;
FiniMove ;
END;

function Caract(N : String3) : String3 ;
var St : String ;
BEGIN
if N='BQE' then Result:='BAN' else
if N='IMO' then Result:='IMM' else Result:=N ;
END ;

procedure TFExpCpte.SousDetail(Str : string ; Q : Tquery) ;
//var Clef :  string ;
BEGIN
Case NumTypCpt of
   CptGEN : BEGIN
            if str=QS(Q.FindField('G_NATUREGENE').AsString,3) then
               BEGIN
               Writeln(Fichier, '02300'+QS(Q.FindField('G_GENERAL').AsString,17)+'TYP'+QS(Q.FindField('G_NATUREGENE').AsString,3)+
                       'CC '+'PEE'+ 'ECG'+format_String(Compte,17)+'PC '+'PEE'+QS(Q.FindField('G_ABREGE').AsString,17)+
                       QS(Q.FindField('G_LIBELLE').AsString,35)+QS(Q.FindField('G_LIBELLE').AsString,35)+Blanc(3))
               END else
               BEGIN
               if str='COR' then
                  BEGIN
                  Writeln(Fichier, '02300'+QS(Q.FindField('G_GENERAL').AsString,17)+'TYP'+'COR'+{Niv Integer balance???}
                          'CC '+'PEE'+ 'ECG'+format_String(Compte,17)+'PC '+'PEE'+QS(Q.FindField('G_ABREGE').AsString,17)+
                          QS(Q.FindField('G_LIBELLE').AsString,35)+QS(Q.FindField('G_LIBELLE').AsString,35)+Blanc(3))
                  END else
                  BEGIN
                  Writeln(Fichier, '02300'+QS(Q.FindField('G_GENERAL').AsString,17)+'CDC'+str+
                          'CC '+'PEE'+ 'ECG'+format_String(Compte,17)+'PC '+'PEE'+QS(Q.FindField('G_ABREGE').AsString,17)+
                          QS(Q.FindField('G_LIBELLE').AsString,35)+QS(Q.FindField('G_LIBELLE').AsString,35)+Blanc(3)) ;
                  END ;
               END ;
            END ;
   CptAUX : BEGIN
            if str='COR' then
               BEGIN
               Writeln(Fichier, '02300'+QS(Q.FindField('T_AUXILIAIRE').AsString,17)+'TYP'+'COR'+{Niv Integer balance???}
                       'CC '+'PEE'+ 'ECG'+QS(Q.FindField('T_COLLECTIF').AsString,17)+'PC '+'PEE'+QS(Q.FindField('T_ABREGE').AsString,17)+
                       QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_DEVISE').AsString,3))
               END else
               BEGIN
               if Str=QS(Q.FindField('T_NATUREAUXI').AsString,3) then
                  BEGIN
                  Writeln(Fichier, '02300'+QS(Q.FindField('T_AUXILIAIRE').AsString,17)+'TYP'+Caract(QS(Q.FindField('T_NATUREAUXI').AsString,3))+
                          'CC '+'PEE'+ 'ECG'+QS(Q.FindField('T_COLLECTIF').AsString,17)+'PC '+'PEE'+QS(Q.FindField('T_ABREGE').AsString,17)+
                          QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_DEVISE').AsString,3))
                  END else
                  BEGIN
                  Writeln(Fichier,'02300'+QS(Q.FindField('T_AUXILIAIRE').AsString,17)+'CDC'+str+
                          'CC '+'PEE'+ 'ECG'+QS(Q.FindField('T_COLLECTIF').AsString,17)+'PC '+'PEE'+QS(Q.FindField('T_ABREGE').AsString,17)+
                          QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_LIBELLE').AsString,35)+QS(Q.FindField('T_DEVISE').AsString,3)) ;
                  END ;
               END ;
            (*Q.Close ;
            Q.sql.Clear ;
            Clef:=QS(Q.FindField('T_AUXILIAIRE',0) ;
//Simon "
            Q.sql.add('Select R_PRINCIPAL,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,'+
                          'R_CLERIB,R_DOMICILIATION from RIB '+
                          'where R_AUXILIAIRE = "'+Clef+'"') ;
            ChangeSQL(Q) ; Q.Prepare ;
            Q.Open ;
            merde toto4:=QS('R_DEVISE',3) ;
            merde toto8:=QS('R_PAYS',3) ;
            ShowMessage(toto1+#13+toto2+#13+toto3+#13+toto5+#13+toto6+#13+toto7) ;}
            While not(Q.EOF) DO
              BEGIN
              if QS(Q.FindField('R_PRINCIPAL',1)='X' then
                 BEGIN
                 Writeln(Fichier,'02325'+QS(Q.FindField('T_AUXILIAIRE',17)+'BK '+
                         QS(Q.FindField('R_NUMEROCOMPTE',11)+QS(Q.FindField('R_CLERIB',2)+Blanc(4)+QS(Q.FindField('T_LIBELLE',35)+
                         Blanc(35)+Blanc(3)+{QS(Q.FindField('R_DEVISE',3)}Blanc(5)+QS(Q.FindField('R_ETABBQ',5)+
                         Blanc(3)+Blanc(3)+QS(Q.FindField('R_GUICHET',5)+'25 '+'108'+QS(Q.FindField('R_DOMICILIATION',35)+
                         Blanc(35)+Blanc(2){QS(Q.FindField('R_PAYS',3)});
                 Writeln(Fichier,'02350'+QS(Q.FindField('T_AUXILIAIRE',17)) ;
                 END ;
              Q.Next ;
            *)
            END ;
   CptANA : BEGIN
            if str='COR' then
               BEGIN
               Writeln(Fichier,'02300'+QS(Q.FindField('S_SECTION').AsString,17)+'TYP'+'COR'+
                       'CC '+'PEE'+ 'ECA'+QS(Q.FindField('S_SECTION').AsString,17)+'PC '+'PEE'+QS(Q.FindField('S_ABREGE').AsString,17)+
                       QS(Q.FindField('S_LIBELLE').AsString,35)+QS(Q.FindField('S_LIBELLE').AsString,35)+Blanc(3)) ;
               END else
               BEGIN
               Writeln(Fichier,'02300'+QS(Q.FindField('S_SECTION').AsString,17)+'CDC'+str+
                       'CC '+'PEE'+ 'ECA'+QS(Q.FindField('S_SECTION').AsString,17)+'PC '+'PEE'+QS(Q.FindField('S_ABREGE').AsString,17)+
                       QS(Q.FindField('S_LIBELLE').AsString,35)+QS(Q.FindField('S_LIBELLE').AsString,35)+Blanc(3)) ;
               END ;
            END ;
   cptBUD : BEGIN
            if str='COR' then
              BEGIN
              Writeln(Fichier, '02300'+QS(Q.FindField('BG_BUDGENE').AsString,17)+'TYP'+'COR'+{Niv Integer balance???}
                      'CC '+'PEE'+ 'ECB'+QS(Q.FindField('BG_BUDGENE').AsString,17)+'PC '+'PEE'+QS(Q.FindField('BG_ABREGE').AsString,17)+
                      QS(Q.FindField('BG_LIBELLE').AsString,35)+QS(Q.FindField('BG_LIBELLE').AsString,35)+Blanc(3)) ;
              END else
              BEGIN
              Writeln(Fichier, '02300'+QS(Q.FindField('BG_BUDGENE').AsString,17)+'CDC'+str+
                      'CC '+'PEE'+ 'ECB'+QS(Q.FindField('BG_BUDGENE').AsString,17)+'PC '+'PEE'+QS(Q.FindField('BG_ABREGE').AsString,17)+
                      QS(Q.FindField('BG_LIBELLE').AsString,35)+QS(Q.FindField('BG_LIBELLE').AsString,35)+Blanc(3)) ;
              END ;
            END ;
   END ;
END ;

procedure TFExpCpte.Resume ;
var Qte : string[15] ;
BEGIN
Str(Quantite:15,Qte) ;
Writeln(Fichier, '02700'+'CPT'+Qte+'001') ;
Writeln(Fichier, '02800'+'ACC'+Format_String(Dossier.CombFctTxt,3)+Blanc(3)+Format_String(Dossier.EditRefTxt,35)+
                 Format_String(Dossier.EditTxt,35)+Blanc(143)) ;
END;

{================ Format CEGID ======================}

Function TFExpCpte.ExporteCptCEGID : Boolean ;
Var Q1,QSoc:  TQuery ;
    F : TextFile ;
    St : String ;
    Qui : Char ;
    QueryEof : Boolean ;
    Q : TQuery ;
    SQL : String ;
BEGIN
Result:=FALSE ;
Case NumTypCpt of
  CptGEN : Qui:='G' ;
  CptANA : Qui:='S' ;
  CptAUX : Qui:='X' ;
  Else Exit ;
  END ;
ReqSQL(FALSE,SQL) ;
Q:=OpenSQL(SQL,TRUE) ;
If Q.Eof Then
  BEGIN
  Result:=FALSE ;
  If (Qui<>'S') Or (FSousSection.State=cbUnChecked) Then BEGIN MsgBox.Execute(19,caption,'') ; EXIT ; END ;
  END ;
Result:=TRUE ;
//if Q then If (Qui<>'S') Or (FSousSection.State=cbUnChecked) Then BEGIN MsgBox.Execute(19,caption,'') ; exit ; END ;
InitMove(RecordsCount(Q)+1,'') ;
Movecur(False) ;
AssignFile(F,FileName.Text) ;
{$I-} ReWrite (F) ; {$I+}
if IoResult<>0 then exit ;
{$IFDEF SPEC302}
QSoc:=OpenSQL('Select SO_LIBELLE From SOCIETE',True) ;
if not QSOC.Eof then BEGIN QSOC.First ; St:=Format_String(QSOC.Fields[0].AsString,70) ; END ;
Ferme(QSOC) ;
{$ELSE}
St:=GetParamSoc('SO_LIBELLE') ; St:=Format_String(St,70) ;
{$ENDIF}
St:=St+FFormat.Value+Lequel[1] ;
Writeln(F,St) ;
If ((Qui<>'S') Or (FSousSection.State<>cbChecked))  Then
 BEGIN
 While not Q.Eof do
   BEGIN
   if TestBreak then Break ; Inc(NbExport) ;
   St:=FaitStCPTCEGID(Qui,Q) ;
   if ChoixFmt.Ascii then St:=Ansi2Ascii(St) ;
   If St<>'' Then Writeln(F,St) ;
   Q.Next ;
   Movecur(False) ;
   END ;
 End ;
Ferme(Q) ;
If (Qui='S') And (FSousSection.State<>cbUnChecked) Then ExporteSousSection(F,'',FALSE,TRUE) ;
CloseFile(F) ;
FiniMove ;
END ;

procedure TFExpCpte.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFExpCpte.CAxeChange(Sender: TObject);
begin
if CurNat='FSE' then
  BEGIN
  case CAxe.Value[2] of
    '1': CPTEDEBUT.ZoomTable:=TzSection ;
    '2': CPTEDEBUT.ZoomTable:=TzSection2 ;
    '3': CPTEDEBUT.ZoomTable:=TzSection3 ;
    '4': CPTEDEBUT.ZoomTable:=TzSection4 ;
    '5': CPTEDEBUT.ZoomTable:=TzSection5 ;
    END ;
  END else
if CurNat='FBS' then
  BEGIN
  case CAxe.Value[2] of
    '1': CPTEDEBUT.ZoomTable:=TzBudSec1 ;
    '2': CPTEDEBUT.ZoomTable:=TzBudSec2 ;
    '3': CPTEDEBUT.ZoomTable:=TzBudSec3 ;
    '4': CPTEDEBUT.ZoomTable:=TzBudSec4 ;
    '5': CPTEDEBUT.ZoomTable:=TzBudSec5 ;
    END ;
  END ;
CPTEFIN.ZoomTable:=CPTEDEBUT.ZoomTable ;
end;


procedure TFExpCpte.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FreeAndNil(ObjFiltre); //SG6 10/11/04 Gestion des Filtres FQ 14826
//SG6 13/01/05 FQ 15242
//if IsInside(Self) then Action:=caFree ;
end;



procedure TFExpCpte.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //SG6 13/01/05 FQ 15242
  if key = VK_ESCAPE then BFermeClick(nil);
end;

end.

