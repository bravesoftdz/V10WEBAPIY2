unit ImpCegid;

// 05/06/1999 - CA - Prise en compte de I_COMPTEREF
// 30/06/1999 - CA - Remplacement CO_TYPE par CC_TYPE
// 07/07/1999 - CA - Gestion de la barre d'avancement en insertion dans la base
// 08/07/1999 - CA - Si CB alors TVA récupérée = TVA récupérable
// 08/07/1999 - CA - I_ETAT='OUV'
// 09/07/1999 - CA - Chargement du fichier de sauvegarde en mémoire
// 01/08/2000 - CA - Prise en compte des fiches en historique ==> I_ETAT = "FER" si St[3]='H'.
// 01/08/2000 - CA - Possibilité d'intégrer l'ancien code dans le libellé.
// 06/06/2002 - CA - HMTitre.Mess[20] changé pour abrégé < 17 caractères
// 03/09/2002 - CA - Messages détaillés en import + suppression des conseils et avertissements
// FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur
// FQ 17259 - BTY - 01/06 Nouveau top dépréciation dans IMMO
// TGA - 26/01/2006 Ajout champs sur table IMMO
// TGA - 04/04/2006 Ajout champs sur table IMMO
// BTY - 04/06 FQ 17516 Nouveau top changement de regroupement dans IMMO
// TGA - 07/04/06 Initialisation de champs IMMO, IMMOLOG
// MBO - 02/05/2006 - FQ 17569 - alimentation des zones i_datedebeco et i_datedebfis suivant methode
// BTY - 05/06 FQ 18119 Positionner les indicateurs de modif compta dans PARAMSOC
// TGA - 07/09/2006 I_PRF => I_PFR
// MVG - 10/11/2006 FQ 19093
// MVG - 10/11/2006 FQ 19092
// MBO - 14/12/2006 FQ 19343 - amélioration pour import expert Winner Immo
// MBO - 06/03/2007 spécificités pour dossier winner, Iseries, quadra, Cegid (boolean ImportWIQC)
// MBO - 19/03/2007 FQ 19846 - ajout d'un rechargement tablette lieu géo
// MVG - 21/05/2007 FQ 19903 pour que la croix ferme la fenêtre
// BTY - 05/2007 FQ 20256 Enregistrer le type de dérogatoire dans i_typederoglia
// MBO - 10/07/07 ajout d'un rechargement tablette des regroupements
// MBO - 12/07/07 import du code etablissement si dossier gérant plusieurs etablissements
// MBO - 13/07/07 import de calcul s/VNF - gestion fiscale - remplact composant - FQ 21162
// MBO - 25/07/07 : correction faute orthographe sur message lié à TIMP_LGCPTEINCOHERENTE
// MBO - 13/09/07 FQ 21435 - import des motifs de cession (le code était pris sur une position alors qu'il en fait 3)
// MBO - 18/09/07 FQ 21454 - calculer l'antérieur réintégration si antérieur fiscal < antérieur éco
// MBO - 15/10/07 FQ 21657 - réinit du boolean 'remplacement'
// MBO - 31.10.07 FQ 21770 - Rechargement des tablettes TIMOTIFCESSION et AMREGROUPEMENT

interface

uses
  SysUtils, Classes, Controls, Forms, Dialogs,
  hmsgbox, Hctrls, StdCtrls, ExtCtrls,  HEnt1,
  ImEnt, HStatus{, Filtre SG6 10/11/04 Gestion des Filtres FQ 14826}, ComCtrls, Menus, HTB97,
  HPanel, UIUtil,ImContra,Outils,UTob, ImCRImp, Math, Mask,   //UtobDebug,
  (*DB, Windows,MajTable, Messages, HSysMenu , Buttons, Graphics,DBCtrls,Hqry, PGIEnv, *)
  Paramsoc, FileCtrl
  , HSysMenu
  , About
  {$IFNDEF SERIE1}
  , UObjFiltres {SG6 10/11/04 Gestion des filtres FQ 14826}
  {$ENDIF}
  ;

function CompareDateAchat (Item1,Item2:Pointer) : integer;
procedure MajParamSocImmo;

const
    TIMP_FILENULL            = -101;
    TIMP_FILENOTEXIST        = -102;
    TIMP_ERRORINSERTDB       = -103;
    TIMP_DATEINCOHERENTE     = -104;
    TIMP_CREATIONEXOFAILED   = -105;
    TIMP_LGCPTEINCOHERENTE   = -106;
    TIMP_TABLESANNEXES       = -107;
type
  TImportImmoInfo = procedure (Sender : TObject ; ErrCode : integer ; ErrMsg : string)  of Object ;
  TCodeEq = class
    CodeTheo : string;
    CodeStr : string;
    CodeNum : Int64;
    DateAchat : TDateTime;
  end;
  TFImpCegid = class(TForm)
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    PTools: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BStop: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Pages: TPageControl;
    TabSheet1: TTabSheet;
    MsgBox: THMsgBox;
    Sauve: TSaveDialog;
    Msg: THMsgBox;
    HMTitre: THMsgBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    HLabel2: THLabel;
    Bevel1: TBevel;
    GBNumerotation: TGroupBox;
    cbNumCpte: TRadioButton;
    cbNumInc: TRadioButton;
    cbIncAchat: TRadioButton;
    GBLibelle: TGroupBox;
    cbLibIdent: TRadioButton;
    cbCodeLib: TRadioButton;
    cbLibCode: TRadioButton;
    FileName: THCritMaskEdit;
    Timer: TTimer;
    Bevel2: TBevel;
    TITREAVANCEMENT: THLabel;
    DETAILAVANCEMENT: THLabel;
    cbCreaCpteAuto: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BStopClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FileNameElipsisClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    {$IFNDEF SERIE1}
    ObjFiltre : TObjFiltre; //SG6 10/11/04 Gestion des Filtres  FQ 14826
    {$ENDIF}
    OkFiscal,Arreter,OkLecture,OkAdd : boolean ;
    NbErreur, NbImport,NbModifies,NbLu : integer ;
    TErreur : TList ;
    CurNat : String ;
    FiltreImp:String ;
    fCurContrat : TImContrat;
    fTabCode : HTStrings;
    fLCode : TList;
    fTOB : TOB;
    fTOBLog : TOB;
    fTOBDiv : TOB;
    FTobGenExiste : TOB;
    FTobGenACreer : TOB;
    fFileStringList : HTStrings;
    fbGereStat : boolean;
    fbStat : boolean;
    fbFamille : boolean;
    fInfoImp : TInfoImpImmo;
    fOffsetCodeImmo : Int64;
    fCodeEnCours : string;
    fNoLigne : Integer;
    fModeAuto : boolean;
    fModeSilencieux : boolean;
    fOnInformation: TImportImmoInfo;
    fLgCpteGen,fbLngCpte : integer;
    fRappFile : string;
    ImportWIQC : boolean;
    Remplacement : boolean;
    
    procedure SetOnInformation(const Value: TImportImmoInfo);
    procedure ShowError(Code: integer; St : string = '') ;
    function  GetMsgErr(Code: integer): String;
    function IsExoClose(DateDebut, DateFin : string) : boolean ;
    function ImporteImmo : boolean;
    function ImporteImmo2(St : string) : boolean;
    function  TestBreak : boolean ;
    procedure ImporteImmo2Immo(St : string) ;
    procedure ImporteImmo2Cb(St : string) ;
    procedure ImporteImmo2Echeance(St : string) ;
    procedure ImporteImmo2Amort(St : string) ;
    procedure ImporteImmo2Compte(St : string) ;
    procedure ImporteImmo2Motif(St : string) ;
    procedure ImporteImmo2Lieu(St : string) ;
    procedure ImporteImmo2Famille(St : string) ;
    procedure ImporteImmo2Stat(St : string) ;
    procedure ImporteImmo2Regroup(St : string);  // nouveau mbo fq 19343
    function FormateDate(St : String) : TDateTime;
    function TrouveImmoDansTob(Code : string) : TOB;
    function CalculCodeImmo (St : string;Compte : string='') : string;
    procedure DetruitListeCode;
    procedure ClasseTobSuivantAchat;
    procedure ControleEtEnregistreCompte(stCompte: string);
    procedure InsertDB;
    function IsDossierImmoVide: boolean;
    procedure OnInformationImport(Sender: TObject; ErrCode: integer;
      ErrMsg: string);
    procedure UpdateAvancement(St1, St2: string);
    function ControleExercice(St: string): boolean;
    procedure CreationFichierRapport(stFileName: string);
    procedure UpdateFichierRapport(st: string);
    procedure CloseFichierRapport(bRestaurationOK: boolean);
    function GetCodeLieu(stCode: string): string;
    function GetCodeEtablissement(stCode: string): string;
    procedure SetPanelImmoInfo(Enable: boolean);
    procedure  EnregDepotGarantieT(T1 : TOB;St: string) ;
    procedure TraiteGeneral ( stGeneral, stLibelle : string );
    function TrouveAncienCode ( Code : string ) : string;
    function TrouveCompteImmo ( Code : string ) : string;
    procedure InitDossierEuro;
    procedure TrouveTypeDerogatoireDesImmos; // FQ 20256
  public
    property OnInformation : TImportImmoInfo read fOnInformation write SetOnInformation;
  end;

procedure ImportCegid ;
function GetNumeroOrdreSuivant(OBLog : TOB; Champ : string):integer;
procedure SavImo2PGI (FileName : string; bSilencieux : boolean);

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ImPlan,ImOutGen
  {$IFDEF SERIE1}
  , S1Util  //YCP 25/08/05
  {$ELSE}
  , Ent1
  {$ENDIF SERIE1}
  {$IFDEF eAGLClient}
  {$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  {$ENDIF eAGLClient}
  ;

{$R *.DFM}

procedure SavImo2PGI (FileName : string; bSilencieux : boolean);
var FImpCegid : TFImpCegid ;
    PP : THPanel ;
    {$IFNDEF SERIE1}
    Composants : TControlFiltre; //SG6   Gestion des Filtes 10/11/04   FQ 14826
    {$ENDIF}
BEGIN
  if Blocage(['nrCloture','nrBatchImmo'],False,'nrBatchImmo') then exit ;
  FImpCegid:=TFImpCegid.Create(Application) ;
  {$IFNDEF SERIE1}
  //SG6 10/11/04 Gestion de filtres FQ 14826
  Composants.PopupF   := FImpCegid.POPF;
  Composants.Filtres  := FImpCegid.FFILTRES;
  Composants.Filtre   := FImpCegid.BFILTRE;
  Composants.PageCtrl := FImpCegid.Pages;
  FImpCegid.ObjFiltre := TObjFiltre.Create(Composants, '');
  {$ENDIF}


  FImpCegid.fbStat := false;
  FImpCegid.fbFamille := false;
  FImpCegid.fModeSilencieux := bSilencieux;
  if bSilencieux then FImpCegid.OnInformation := FImpCegid.OnInformationImport;
  FImpCegid.fModeAuto := (FileName <> '');
  FImpCegid.FileName.Text := FileName;
  // en mode silencieux, on créé les comptes associés automatiquement
  FImpCegid.cbCreaCpteAuto.Checked := bSilencieux;
  {$IFDEF SERIE1}
  {$ELSE}
   If VH^.RecupSISCOPGI then
     BEGIN
     FImpCegid.cbCreaCpteAuto.Checked:=FALSE ;
     FImpCegid.CBLibIdent.Checked:=FALSE ;
     FImpCegid.CBCodeLib.Checked:=FALSE ;
     FImpCegid.CBLibCode.Checked:=FALSE ;
     FImpCegid.CBCodeLib.Checked:=TRUE ;
     END ;
  {$ENDIF}
  PP:=FindInsidePanel ;
  ChargeVHImmo;
  if ((FImpCegid.fModeAuto) or (PP=Nil)) then  // Pas d'inside si appel depuis transfert
  BEGIN
    try
      FImpCegid.ShowModal ;
    finally
      FImpCegid.Free ;
      Bloqueur('nrBatchImmo',False) ;
    end;
    Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
    InitInside(FImpCegid,PP) ;
    FImpCegid.Show ;
  END ;
end;

procedure ImportCegid ;
var FImpCegid : TFImpCegid ;
    PP : THPanel ;
BEGIN
  if Blocage(['nrCloture','nrBatchImmo'],False,'nrBatchImmo') then exit;
  FImpCegid:=TFImpCegid.Create(Application) ;
  FImpCegid.fbStat := false;
  FImpCegid.fbFamille := false;
  PP:=FindInsidePanel ;
  if PP=Nil then
  BEGIN
    try
      FImpCegid.ShowModal ;
    finally
      FImpCegid.Free ;
      Bloqueur('nrBatchImmo',False) ;
    end;
    Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
    InitInside(FImpCegid,PP) ;
    FImpCegid.Show ;
  END ;
END ;

function TFImpCegid.IsExoClose(DateDebut, DateFin : string) : boolean ;
var Deb, Fin : TDateTime ;
begin
Deb:=FormateDate(DateDebut) ;
Fin:=FormateDate(DateFin) ;
Result:=(Deb<VHImmo^.Encours.Deb) and (Fin<VHImmo^.Encours.Deb) ;
end ;

procedure TFImpCegid.SetPanelImmoInfo(Enable : boolean);
begin
 GBNumerotation.Enabled := Enable;
 GBLibelle.Enabled := Enable;
 cbCreaCpteAuto.Enabled := Enable;
 FileName.Enabled := Enable;
 BFiltre.Enabled := Enable;
 FFiltres.Enabled := Enable;
 BValider.Enabled := Enable;
 BFerme.Enabled := Enable;
 BStop.Enabled := Enable;
 BAide.Enabled := Enable;
end;

procedure TFImpCegid.FormShow(Sender: TObject);
begin
  // TobDebug(nil);
  if fModeAuto then SetPanelImmoInfo (False);
  Timer.Enabled := True;
  fCurContrat := nil;
  FiltreImp:='IMPIMO' ;
  CurNat:='IMM';
  Caption:=HMtitre.Mess[0] ;
  Sauve.Title:=Caption ;
  TErreur:=TList.Create;

  {$IFNDEF SERIE1}
  //SG6 10/11/04 Gestion des Filtres  FQ 14826
  ObjFiltre.FFI_TABLE:=FiltreImp;
  ObjFiltre.Charger;
  {$ENDIF}

  if (not IsDossierImmoVide) then
  begin
    cbNumInc.Checked := True;
    GBNumerotation.Enabled := False;
  end else GBNumerotation.Enabled := True;
  UpdateCaption(Self) ;
end;

procedure TFImpCegid.BStopClick(Sender: TObject);
begin Arreter:=True ; end;

procedure TFImpCegid.FormClose(Sender: TObject; var Action: TCloseAction);
begin
{$IFNDEF SERIE1}
FreeAndNil(ObjFiltre); //SG6 10/11/04 Gestion des filtres FQ 14826
{$ENDIF}
TErreur.Free ;
if IsInside(Self) then
   BEGIN
   Bloqueur('nrBatchImmo',False) ;
   Action:=caFree ;
   END ;
end;

function TFImpCegid.TestBreak : Boolean ;
BEGIN
Application.ProcessMessages ;
if Arreter then if Msgbox.Execute(18,Caption,'')<>mryes then Arreter:=False ;
Result:=Arreter ;
END ;

function TFImpCegid.FormateDate(St : String) : TDateTime;
var dt : TDateTime;
begin
  dt := Str8ToDate(St,TRUE);
  if dt = 0 then dt := iDate1900;
  result := dt;
end;

procedure TFImpCegid.BValiderClick(Sender: TObject);
Var OKInsert : boolean;
    wNewCode: string ;
begin
  UpdateAvancement ('','');
  TITREAVANCEMENT.Visible := True;
  DETAILAVANCEMENT.Visible := True;
  if FileName.Text <> '' then CreationFichierRapport(FileName.text)
  else CreationFichierRapport('IMOPGI_ERROR.LOG');
  if (FileName.text='') then
  begin
    ShowError (TIMP_FILENULL);
    exit ;
  end;
  if not FileExists(FileName.Text) then
  begin
    ShowError (TIMP_FILENOTEXIST);
    exit ;
  end;
  ActivePanels(Self,False,False) ;
  // Récupération des éléments significatifs du fichier
  UpdateAvancement (HMTitre.Mess[2],'');
  NbImport:=0 ; NbErreur:=0 ; NbModifies:=0 ; NbLu:=0 ;
  fFileStringList := HTStringList.Create;
  fFileStringList.LoadFromFile (FileName.Text);
  NbLu :=   fFileStringList.Count;
  fFileStringList.Free;

  //Initialisation
  UpdateAvancement (HMTitre.Mess[4],'');
  Arreter:=False ;
  fLgCpteGen := VHImmo^.Cpta[ImGeneTofb].Lg ; //GetParamSoc('SO_LGCPTEGEN');
  fTabCode := HTStringList.Create;
  fLCode := TList.Create;
  fTOB := TOB.Create('',nil,-1);
  fTOBLog := TOB.Create ('',nil,-1);
  fTOBDiv := TOB.Create ('', nil, -1);
  FTobGenExiste := TOB.Create ('', nil , -1);
  FTobGenACreer := TOB.Create ('', nil , -1);
  fInfoImp := TInfoImpImmo.Create;
  OkInsert := false;
  SourisSablier;
  fOffsetCodeImmo := 0;
  if (not IsDossierImmoVide) then
  begin
    wNewCode:=NouveauCodeImmo ;
    if wNewCode<>'' then fOffsetCodeImmo:=(StrToInt64(wNewCode)-1) ;
  end ;
  // Importation du fichier
  UpdateAvancement (HMTitre.Mess[5],'');
  if ImporteImmo then
  begin
//    if (cbCreaCpteAuto.Checked) then CreationCompteAuto;
    if (cbIncAchat.Checked) then ClasseTobSuivantAchat;
    if (Transactions(InsertDB,1)<>OeOk) then
    begin
      ShowError (TIMP_ERRORINSERTDB);
    end else
    begin
      OKInsert := True;
      cbNumInc.Checked := TRUE;
      GBNumerotation.Enabled := FALSE;
      // 05/06 FQ 18119 Positionner les indicateurs de modif
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi(True);
    end;
  end;

  // Mise à jour du fichier rapport
  UpdateFichierRapport ('Fiche récupérées : '+ IntToStr(fInfoImp.nFiche));
  UpdateFichierRapport ('  - en pleine propriété : '+ IntToStr(fInfoImp.nPRO));
  UpdateFichierRapport ('  - en crédit-bail : '+ IntToStr(fInfoImp.nCB));

  // Libération
  fTOB.Free;
  fTOBLog.Free;
  fTOBDiv.Free;
  FTobGenExiste.Free;
  FTobGenACreer.Free;
  fTabCode.Free;
  DetruitListeCode;
  SourisNormale;
  ActivePanels(Self,True,False) ;
  UpdateAvancement (HMTitre.Mess[3],'');

  // Compte-rendu
  if ((OKInsert) and (not fModeSilencieux)) then CompteRenduImportationImmo ( fInfoImp);
  CloseFichierRapport (OkInsert);
  fInfoImp.Free;
//  if ChoixFmt.Detruire then DeleteFile(FileName.text) ;
  Screen.cursor:=SynCrDefault ;
  if (fModeSilencieux) and (fModeAuto) then ModalResult := mrYes;
END;

{----------------------------- Lancement de l'importation ----------------------------------}
function TFImpCegid.ImporteImmo : boolean;
var St : String ;
    OBEche,OBImmo : TOB;
    OkRest : boolean;
    F : TextFile;
BEGIN
  OkRest := true;
  try
    InitMove(NbLu,'') ;MoveCur(False) ; TErreur.Clear ;
    OkLecture:=True ;
    SourisSablier ;
    AssignFile(F,FileName.Text);
    Reset (F);
    fNoLigne := 0;
    while not Eof (F) do
    begin
      Inc (fNoLigne,1);
      if TestBreak then Break ;
      MoveCur(False) ;
      if OkLecture  then Readln(F,St);
      if not ImporteImmo2(St) then
      begin OkRest := false; break; end;
    end;
    CloseFile(F);
    TrouveTypeDerogatoireDesImmos; // FQ 20256
  finally
    if (OKRest) and (fCurContrat <> nil) then
    begin
      // ajout mbo 29.12.06
      if ImportWIQC then
         fCurContrat.MetTrancheDansEcheance
      else
         fCurCOntrat.ConvertTrancheIntoEcheance;

      OBImmo := TrouveImmoDansTOB (fCurContrat.sCode);
      OBEche := TOB.Create('',OBImmo,-1);
      //fCurContrat.CalculEcheances;
      fCurContrat.MajTableEcheancesTOB(OBEche);
    end;
    if fCurContrat <> nil Then fCurContrat.Free;
    FiniMove ;
    SourisNormale ;
  end ;
  GBNumerotation.Enabled := (IsDossierImmoVide);
  result := OkRest;
END ;

function CompareDateAchat (Item1,Item2:Pointer) : integer;
var CodeEq1,CodeEq2 : TCodeEq;
begin
  CodeEq1 := Item1;CodeEq2 := Item2;
  if CodeEq1.DateAchat > CodeEq2.DateAchat then Result := 1
  else if CodeEq1.DateAchat < CodeEq2.DateAchat then Result := -1
  else Result := 0;
end;

procedure TFImpCegid.ClasseTobSuivantAchat;
var i,j,k : integer;
    OBImmo : TOB;
begin
  fLCode.Sort(CompareDateAchat);
  InitMove(fTOB.Detail.Count+fTOBLog.Detail.Count,'');
  for i:=0 to fTOB.Detail.Count - 1 do
  begin
    for j:=0 to fLCode.Count - 1 do
    begin
      if TCodeEq(fLCode.Items[j]).CodeNum = StrToInt64( fTOB.detail[i].GetValue('I_IMMO')) then
      begin
        OBImmo := fTOB.Detail[i];
        if OBImmo.Detail.Count > 0 then
        begin
          for k:=0 to OBImmo.Detail[0].Detail.Count - 1 do
            OBImmo.Detail[0].Detail[k].PutValue('IA_IMMO',Format('%.10d',[j+1]));
        end;
        fTOB.Detail[i].PutValue('I_IMMO',Format('%.10d',[j+1]));
        MoveCur(False) ;
        break;
      end;
    end;
  end;
  for i:=0 to fTOBLog.Detail.Count - 1 do
  begin
    for j:=0 to fLCode.Count - 1 do
    begin
      if TCodeEq(fLCode.Items[j]).CodeNum = StrToInt64( fTOBLog.detail[i].GetValue('IL_IMMO')) then
      begin
        fTOBLog.Detail[i].PutValue('IL_IMMO',Format('%.10d',[j+1]));
        MoveCur(False) ;
        break;
      end;
    end;
  end;
  FiniMove;
end;

function TFImpCegid.ControleExercice (St : string) : boolean;
var dtDebut, dtFin : TDateTime;
    ret : longint;
    sLibelleExo : string;
    Q : TQuery;
begin
  dtDebut := FormateDate(Trim( Copy(St,185,8)));
  dtFin := FormateDate(Trim( Copy(St,193,8)));
  Q := OpenSQL ('SELECT * FROM EXERCICE',TRUE);
  if Q.Eof then  // si aucun enregistrement
//  if ExecuteSQL ('SELECT * FROM EXERCICE') = 0 then
  begin  // Pas d'exercice créé, on en créé un avec les infos minimales
    sLibelleExo := TraduireMemoire('Exo du ')+DateToStr(dtDebut)+TraduireMemoire(' au ')+DateToStr(dtFin);
    ret := ExecuteSql('INSERT INTO EXERCICE (EX_EXERCICE,EX_LIBELLE,EX_ABREGE,EX_DATEDEBUT,EX_DATEFIN,'+
           'EX_ETATCPTA,EX_ETATBUDGET,EX_ETATADV,EX_ETATAPPRO,EX_ETATPROD,EX_SOCIETE,EX_VALIDEE,'+
           'EX_DATECUM,EX_DATECUMRUB,EX_BUDJAL,EX_NATEXO ) '+
           'VALUES ("001","'+sLibelleExo+'","'+Copy(sLibelleExo,1,14)+'","'+UsDateTime(dtDebut)+'",'+
                    '"'+UsDateTime(dtFin)+'","OUV","OUV","NON","NON","NON","'+V_PGI.CodeSociete+'",'+
                    '"------------------------","'+UsDateTime(0)+'","'+UsDateTime(0)+'","","")') ;
    if ret <> 1 then
    begin
      ShowError (TIMP_CREATIONEXOFAILED);
      result := false;
    end else
    begin
      SetParamSoc ('SO_LGCPTEGEN',StrToInt(Copy(St,274,5)));
      fLgCpteGen := GetParamSocSecur ('SO_LGCPTEGEN','');
      SetParamSoc ('SO_BOURREGEN',0);
      SetParamSoc ('SO_LGCPTEAUX',StrToInt(Copy(St,274,5)));
      SetParamSoc ('SO_BOURREAUX',0);
      SetParamSoc('SO_DATEDEBUTEURO',EncodeDate(1999,01,04)) ;
      if (UpperCase(St[284])='E') then
      begin
        SetParamSoc('SO_TENUEEURO', True);
        InitDossierEuro;
      end;
      {$IFDEF SERIE1}
      {$ELSE}
      RechargeParamSoc; // Pour recharger en mémoire le contexte euro
      {$ENDIF}
      result := True;
    end;
  end else
  begin  // Un exercice créé, on vérifie que ses dates correspondent à celles du fichier
    result := (dtDebut = VHImmo^.Encours.Deb) and (dtFin = VHImmo^.Encours.Fin);
    if not result then
    begin
      ShowError (TIMP_DATEINCOHERENTE);
      result := false;
    end;
  end;
  Ferme(Q);
end;

function TFImpCegid.ImporteImmo2(St : string) : boolean ;
var w: string ;
    //CodeImmo : string ;
    //OBImmo  : TOB;

begin
  result:=False ;
  OkAdd:=FALSE ;
  if (Copy(St,1,3)='PPA') then
  if not ControleExercice(St) then
  begin
    result := false;
    exit;
  end else
  begin
    V_PGI.TobSOC.Free;
    V_PGI.TobSoc := nil;
    MajParamSocImmo;  // Mise à jour des comptes immo dans le paramsoc
    ImChargeMagHalley;
    ChargeVHImmo;
    fbGereStat :=  UpperCase(St[212])='O';

    // Ajout mbo - fq 19343 - identification de l'import
    // si W en position 154 ---> on importe un dossier expert WInner Immo
    // si I                      on importe un dossier Iseries (as400)
    // si Q                      on importe un dossier Quadratus
    // si C                      on importe un dossier Cegid

    if (copy (St, 154, 1) = 'W') or
       (copy (St, 154, 1) = 'I') or
       (copy (St, 154, 1) = 'Q') or
       (copy (St, 154, 1) = 'C') then
       ImportWIQC := true
    else
       importWIQC := false;

    w:=copy(St,274,5) ;
    try fbLngCpte:=StrToInt(trim(w)) ; except fbLngCpte:=0 ; end ;
    if ((fbLngCpte>VHImmo^.Cpta[ImGeneTofb].Lg) and  cbNumCpte.Checked) or ((fbLngCpte>4) and cbNumCpte.Checked) then
    begin
      if fModeSilencieux then
      begin
        ShowError (TIMP_LGCPTEINCOHERENTE);
        result := false;
        exit;
      end;
      if PGIAsk('Attention : pour utiliser cette méthode de numérotation,#10#13 les comptes ne doivent pas posséder plus de 4 caractères significatifs.#10#13Voulez-vous continuer ?',Caption)<>mrYes then
        exit ;
    end ;
  end;
  if Copy(St,1,2)='FF' then ImporteImmo2Immo(St) ;
  //if Copy(St,1,2)='FC'  then ImporteImmo2Cession(St) ;
  if Copy(St,1,2)='FL' then ImporteImmo2Cb(St) ;
  //if Copy(St,1,2)='FM' then ImporteImmo2Mutation(St) ;
  if Copy(St,1,2)='FP' then ImporteImmo2Echeance(St) ;
  if Copy(St,1,2)='FA' then ImporteImmo2Amort(St) ;
  //if Copy(St,1,2)='FX' then ImporteImmo2AmortExcept(St) ;
  //if Copy(St,1,2)='FE' then ImporteImmo2Eclat(St) ;
  //if Copy(St,1,2)='FO' then ImporteImmo2Operation(St) ; *)
  {$IFDEF SERIE1}
  if (Copy(St,1,3)='TTC') then ImporteImmo2Compte(St) ;
  {$ELSE}
   If Not VH^.RecupSISCOPGI Then if (Copy(St,1,3)='TTC') then ImporteImmo2Compte(St) ;
  {$ENDIF}
  if (Copy(St,1,3)='TCE') then ImporteImmo2Motif(St) ;
  if (Copy(St,1,3)='TGE') then ImporteImmo2Lieu(St) ;
  if (Copy(St,1,3)='TST') then ImporteImmo2Stat(St) ;
  if (Copy(St,1,3)='TFA') then ImporteImmo2Famille(St) ;
  // ajout mbo - 19343 - pour import des lots winner immo
  if (Copy(St,1,3)='TRG') then ImporteImmo2Regroup(St) ;
  result:=True ;
END ;

procedure TFImpCegid.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure AddMessageLog (Msg : string);
var F : TextFile;
begin
  AssignFile (F,   'c:\immo.log');
  if FileExists ('c:\immo.log') then  Append (F)
  else Rewrite (F);
  Writeln (F,DateToStr(date)+ ' - '+TimeToStr(time) + ' : '+Msg);
  CloseFile (F);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 06/10/2004
Modifié le ... :   /  /
Description .. : - FQ 14306 - CA - 06/10/2004 - Récupération de la base
Suite ........ : d'amortissement
Mots clefs ... :
*****************************************************************}
procedure TFImpCegid.ImporteImmo2Immo(St : string) ;
  function _OperationEnCours ( St : string ) : boolean ;
  begin
    Result := ((St[98] = '1')        // cession
              or (St[99] = '1')     // éclatement
              or (St[100] = '1')    // mutation
              or (St[101] = '1')    // aide fiscale
              or (St[102] = '1')    // non déductible
              or (St[103] = '1')    // exceptionnel
              or (St[104] = '1')    // événement
              or (St[105] = '1')    // passage en réel
              or (St[106] = '1')    // avoir
              or (St[106] = '2')    // Dotation minimum
              or (St[106] = '3'))   // Avoir + Dot. minimum
  end;
var CodeImmo : string ;
    OBImmo, OBLog  : TOB;
    PremMois,PremAnnee,NbMois{,AA,MM,JJ} : Word;   wSt: string ;

begin

  CodeImmo := CalculCodeImmo (St);
  UpdateAvancement(HMTitre.Mess[13],CodeImmo + ' - '+Trim(Copy(St,23,25)));
  OBImmo := TrouveImmoDansTOB (CodeImmo);
  Inc(fInfoImp.nFiche);
  if OBImmo = nil then OBImmo := TOB.Create ('IMMO',fTOB,-1);
  if St[4]='C' then
  begin
    Inc(fInfoImp.nCB);
    OBImmo.PutValue('I_NATUREIMMO','CB');
    OBImmo.PutValue('I_COMPTELIE',AjusteCompteImmo(Copy(St,5,10)));
    ControleEtEnregistreCompte (OBImmo.GetValue('I_COMPTELIE'));
    OBImmo.PutValue('I_COMPTEREF',AjusteCompteImmo(Copy(St,5,10)));
  end
  else
  begin
    Inc(fInfoImp.nPRO);
    OBImmo.PutValue('I_NATUREIMMO','PRO');
    OBImmo.PutValue('I_COMPTEIMMO',AjusteCompteImmo(Copy(St,5,10)));
    OBImmo.PutValue('I_COMPTEREF',AjusteCompteImmo(Copy(St,5,10)));
    if (OBImmo.GetValue('I_COMPTEIMMO')<>'') then
    begin
      ControleEtEnregistreCompte (OBImmo.GetValue('I_COMPTEIMMO'));
      OBImmo.PutValue('I_COMPTEAMORT',CAssAmortissement (OBImmo.GetValue('I_COMPTEIMMO')));
      OBImmo.PutValue('I_COMPTEDOTATION',CAssDotation (OBImmo.GetValue('I_COMPTEIMMO')));
      OBImmo.PutValue('I_COMPTEDEROG',CAssDerog(OBImmo.GetValue('I_COMPTEIMMO')));
      OBImmo.PutValue('I_REPRISEDEROG',CAssRepriseDerog(OBImmo.GetValue('I_COMPTEIMMO')));
      OBImmo.PutValue('I_PROVISDEROG',CAssProvisDerog(OBImmo.GetValue('I_COMPTEIMMO')));
      OBImmo.PutValue('I_DOTATIONEXC',CAssDotationExc (OBImmo.GetValue('I_COMPTEDOTATION')));
      OBImmo.PutValue('I_VACEDEE',CAssVaCedee(OBImmo.GetValue('I_COMPTEIMMO')));
      OBImmo.PutValue('I_AMORTCEDE',OBImmo.GetValue('I_COMPTEAMORT'));
      OBImmo.PutValue('I_VAOACEDEE',OBImmo.GetValue('I_COMPTEIMMO'));
      OBImmo.PutValue('I_REPEXPLOIT',CAssRepExploit (OBImmo.GetValue('I_COMPTEDOTATION')));
      OBImmo.PutValue('I_REPEXCEP',CAssRepExc (OBImmo.GetValue('I_DOTATIONEXC')));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_COMPTEAMORT'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_COMPTEDOTATION'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_COMPTEDEROG'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_REPRISEDEROG'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_PROVISDEROG'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_DOTATIONEXC'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_VACEDEE'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_REPEXPLOIT'));
      ControleEtEnregistreCompte (OBImmo.GetValue('I_REPEXCEP'));
    end;
  end;
  OBImmo.PutValue('I_IMMO',CodeImmo);
  // ### Debut Modif CA le 01/08/2000
  // OBImmo.PutValue('I_ETAT','OUV');
  if St[3] = 'H' then OBImmo.PutValue('I_ETAT','FER')
  else OBImmo.PutValue('I_ETAT','OUV');
  // ### Fin Modif CA le 01/08/2000
  OBImmo.PutValue('I_NATUREBIEN','GEB');
  // ### Début modif CA le 01/08/2000
  // OBImmo.PutValue('I_LIBELLE',Trim(Copy(St,23,25)) );
  OBImmo.PutValue('I_TABLE1',Trim(Copy(St,21,2))); // ### Ajout CA le 06/09/2000
  if cbLibIdent.Checked then OBImmo.PutValue('I_LIBELLE',Trim(Copy(St,23,25)) )
  else if cbCodeLib.Checked then OBImmo.PutValue('I_LIBELLE',Trim(Copy(St,15,6))+ ' - ' + Trim(Copy(St,23,25)) )
  else if cbLibCode.Checked then OBImmo.PutValue('I_LIBELLE',Trim(Copy(St,23,25)) + ' - ' + Trim(Copy(St,15,6)) );
  // ### Fin Modif CA le 01/08/2000
  OBImmo.PutValue('I_ABREGE',Copy(OBImmo.GetValue('I_LIBELLE'),1,17));

  { On ajoute '***' au début du libellé si l'immobilisation a subi une opération sur l'exercice }
  if _OperationEnCours(St) then OBImmo.PutValue('I_LIBELLE',Copy('***'+OBImmo.GetValue('I_LIBELLE'),1,35));

  // MVG 10/11/2006 FQ 19093, Imo II met 'S' et 'F'
  if ((St[48]='F') or (St[48]='S')) then OBImmo.PutValue('I_QUALIFIMMO','S')
  else OBImmo.PutValue('I_QUALIFIMMO','R');
  OBImmo.PutValue('I_DATEPIECEA',FormateDate(Trim(Copy(St,49,8)))) ;
  // Enregistrement de l'acquisition dans le LOG
  OBLog := TOB.Create ('IMMOLOG',fTOBLog,-1);

   // TGA 07.04.2006
  OBLog.PutValue('IL_TAUX',0);

  OBLog.PutValue('IL_TYPEOP','ACQ' );
  OBLog.PutValue('IL_IMMO',CodeImmo);
  OBLog.Putvalue('IL_DATEOP',OBImmo.GetValue('I_DATEPIECEA'));
  OBLog.Putvalue('IL_ORDRE',1);
  OBLog.Putvalue('IL_ORDRESERIE',1);
  OBLog.Putvalue('IL_PLANACTIFAV',0);
  OBLog.Putvalue('IL_PLANACTIFAP',1);
  OBLog.Putvalue('IL_LIBELLE',RechDom('TIOPEAMOR', 'ACQ', FALSE)+' '+DateToStr(OBImmo.GetValue('I_DATEPIECEA')));
  OBLog.Putvalue('IL_TYPEMODIF',AffecteCommentaireOperation('ACQ'));
  // Fin enregistrement acquisition dans LOG
  if (cbNumInc.Checked) or (cbIncAchat.Checked) then
    TCodeEq(fLCode.Items[StrToInt64(CodeImmo)-fOffsetCodeImmo-1]).DateAchat := FormateDate(Trim(Copy(St,49,8)));

  wSt:=Trim(Copy(St,57,8)) ;
  if wSt<>'' then OBImmo.PutValue('I_DATEAMORT',FormateDate(wSt))
             else OBImmo.PutValue('I_DATEAMORT',OBImmo.GetValue('I_DATEPIECEA')) ;

  // ### AJOUT CA le 06/07/2000   : pour les gorets qui génèrent mal les fichiers de sauvegarde, on recalcule ...
  NOMBREMOIS (OBImmo.GetValue('I_DATEAMORT'),VHImmo^.Encours.Deb,PremMois,PremAnnee,NbMois);
  OBImmo.PutValue('I_DUREEREPRISE',IntToStr(MaxIntValue([0,NbMois-1])));
  // ### FIN AJOUT CA le 06/07/2000

  OBImmo.PutValue('I_REFINTERNEA',Trim(Copy(St,65,10))) ;
  OBImmo.PutValue('I_TIERSA',Trim(Copy(St,75,15)));
  OBImmo.PutValue('I_LIEUGEO',GetCodeLieu(Trim(Copy(St,91,3)))) ;

  // ajout mbo - fq 19343 pour expert winner immo
  // n° de lot pour regroupement
  wst := Trim(Copy(St,107,3)) ;
  if (wst <> '') and (ImportWIQC) then
     OBImmo.PutValue('I_GROUPEIMMO',wSt) ;

  // ajout mbo - gestion du code etablissement
  wst := '';

  {$IFDEF SERIE1}
  {$ELSE}
     wst := Trim(Copy(St, 110,3));
     if (wst = '') or not(VHImmo^.EtablisCpta) or not(ImportWIQC) then
        OBImmo.PutValue('I_ETABLISSEMENT',VHImmo^.EtablisDefaut)
     else
        OBImmo.PutValue('I_ETABLISSEMENT',GetCodeEtablissement(wst));
  {$ENDIF}

  if fbGereStat then OBImmo.PutValue('I_TABLE0',Trim(Copy(St,94,4))) ;
  OBImmo.PutValue('I_TVARECUPERABLE',Valeur(Trim(Copy(St,140,13)))/100) ;

  if OBImmo.GetValue('I_NATUREIMMO')='CB' then
    OBImmo.PutValue('I_TVARECUPEREE',OBImmo.GetValue('I_TVARECUPERABLE'))
  else OBImmo.PutValue('I_TVARECUPEREE',Valeur(Trim(Copy(St,153,13)))/100 );
  if st[98]='1' then
  begin
    { Cas d'une immobilisation cédée }
    OBImmo.PutValue('I_MONTANTHT',OBImmo.GetValue('I_BASEECO'));
    OBImmo.PutValue('I_BASETAXEPRO',OBImmo.GetValue('I_MONTANTHT'));
    OBImmo.PutValue('I_VALEURACHAT', OBImmo.GetValue('I_MONTANTHT'));
    OBImmo.PutValue('I_BASEAMORDEBEXO',OBImmo.GetValue('I_MONTANTHT'));
    OBImmo.PutValue('I_BASEAMORFINEXO',OBImmo.GetValue('I_MONTANTHT')) ;
    OBImmo.PutValue('I_MONTANTBASEAMORT',OBImmo.GetValue('I_MONTANTHT')) ;
  end else
  begin
    { Autres cas }
    OBImmo.PutValue('I_BASETAXEPRO',Valeur(Trim(Copy(St,166,13)))/100 );
    OBImmo.PutValue('I_MONTANTHT',Valeur(Trim(Copy(St,179,13)))/100 );
    // CA - 23/04/2002 - Mise à jour de la valeur d'achat
    OBImmo.PutValue('I_VALEURACHAT', OBImmo.GetValue('I_MONTANTHT'));
    OBImmo.PutValue('I_BASEAMORDEBEXO',OBImmo.GetValue('I_MONTANTHT'));
    OBImmo.PutValue('I_BASEAMORFINEXO',OBImmo.GetValue('I_MONTANTHT')) ;
    OBImmo.PutValue('I_MONTANTBASEAMORT',OBImmo.GetValue('I_MONTANTHT')) ;
  end;
  OBImmo.PutValue('I_QUANTITE',Valeur(Trim(Copy(St,231,13)))/100 );
  { Si la quantité est à 0 est qu'on n'est pas en historique, on force la quantité à 1 }
  if ((OBImmo.GetValue('I_QUANTITE')=0) and (St[3]<>'H')) then
    OBImmo.PutValue('I_QUANTITE',1);
  { Si la base est nulle on la recalcule à partir du montant HT
                    car on n'a pas toujours l'information dans le fichier }
  if OBImmo.GetValue('I_BASEECO') = 0 then
    OBImmo.PutValue('I_BASEECO', OBImmo.GetValue('I_MONTANTHT')+
                OBImmo.GetValue('I_TVARECUPERABLE')-OBImmo.GetValue('I_TVARECUPEREE'));
  // fq 19500
  if OBImmo.GetValue('I_BASEFISC') = 0 then
     OBImmo.PutValue('I_BASEFISC',OBImmo.GetValue('I_BASEECO')) ;

  // fq 17569 mbo 05/2006
  // fq 19343 mbo - pour expert winner les dates début amortissement sont transmises (enreg FA)
  if not ImportWIQC then
  begin
     if (OBImmo.GetValue('I_METHODEECO')='DEG') then
     begin
          OBImmo.PutValue('I_DATEDEBECO',OBImmo.GetValue('I_DATEPIECEA'));
          OBImmo.PutValue('I_DATEDEBFIS',OBImmo.GetValue('I_DATEPIECEA'));
     end else
     begin
          OBImmo.PutValue('I_DATEDEBECO',OBImmo.GetValue('I_DATEAMORT')) ;
          OBImmo.PutValue('I_DATEDEBFIS',OBImmo.GetValue('I_DATEAMORT')) ;
     end;
  end;
  // fq 17569 mbo 05/2006
  // fq 19343 pour expert winner les dates début amortissement sont transmises (enreg FA)
  if (OkFiscal) and (not ImportWIQC) then
  begin
    if (OBImmo.GetValue('I_METHODEFISC')='DEG') then
       OBImmo.PutValue('I_DATEDEBFIS',OBImmo.GetValue('I_DATEPIECEA'))
    else
       OBImmo.PutValue('I_DATEDEBFIS',OBImmo.GetValue('I_DATEAMORT')) ;
  end;
  // pour import winner si pas de fiscal ---> mettre date deb fiscal à date d'achat
  if (not OkFiscal) and (ImportWIQC) then
     OBImmo.PutValue('I_DATEDEBFIS',OBImmo.GetValue('I_DATEPIECEA'));


  // FQ 19343 - mbo - 14.12.2006 -
  if (ImportWIQC) then
  begin
     // calcul sur durée restante
     if (st[257] = 'O') then
        OBImmo.PutValue('I_JOURNALA','***');

     // ajout mbo - 13.07.07 - gestion fiscale - FQ 21162
     if st[258] = 'O' then
        OBImmo.PutValue('I_FUTURVNFISC', '***');
  end;

  // ajout pour initialisation  mbo fq 17569
  OBImmo.PutValue('I_SUSDEF','A');
  OBImmo.PutValue('I_REGLECESSION','NOR');
  OBImmo.PutValue('I_SUBVENTION','NON');
  OBImmo.PutValue('I_SBVEC','C');
  // ajout mbo - on importe le fiscal en provenance de winner immo uniquement
  // si la base fiscale est égale à la base éco ou à la base théorique
  // et que la date debut éco est inférieure ou égale à la date debut fiscal
  if (ImportWIQC) and (okFiscal) then
  begin
     //ajout mbo 13.07.07 - on ne fait pas de test si c'est un remplacement de composant - FQ 21162
     if not remplacement then
     begin
        if ((OBImmo.GetValue('I_BASEECO') <> OBImmo.GetValue('I_BASEFISC')) AND
           (OBImmo.GetValue('I_BASEFISC')<> (OBImmo.GetValue('I_MONTANTHT')+
                   OBImmo.GetValue('I_TVARECUPERABLE')-OBImmo.GetValue('I_TVARECUPEREE')))) or
            (OBImmo.GetValue('I_DATEDEBECO') < OBImmo.GetValue('I_DATEDEBFIS')) then
           // on ne reprend pas le fiscal
        begin
          OkFiscal := false;
          OBImmo.PutValue('I_BASEFISC', 0);
          OBImmo.PutValue('I_DATEDEBFIS',OBImmo.GetValue('I_DATEDEBECO'));
          OBImmo.PutValue('I_METHODEFISC', '');
          OBImmo.PutValue('I_REPRISEFISCAL', 0);
          OBImmo.PutValue('I_TAUXFISC',0);
          OBImmo.PutValue('I_DUREEFISC',0);
          OBImmo.PutValue('I_REPRISEDR',0);
          OBImmo.PutValue('I_REPRISEFEC',0);
          OBImmo.PutValue('I_REPRISEFECCEDEE',0);
          OBImmo.PutValue('I_REPRISEFDRCEDEE',0);
          OBImmo.PutValue('I_NONDED','-');
          OBImmo.PutValue('I_FUTURVNFISC', '');
        end;
     end;
  end;

end ;

//---------------------------------------------------------------------------------
// import des enregs de type FA (plan d'amortissement éco et fiscal)

procedure TFImpCegid.ImporteImmo2Amort(St : string) ;
  function _ControleDateValide ( StDate : string ) : boolean;
  var DateFichier : TDateTime;
  begin
    DateFichier := Str8ToDate ( stDate, True );
    if (DateFichier <= 0) then DateFichier := Str8ToDate ( stDate, False ); { Dans le cas des fichiers issus de Quadra, la première date est inversée }
    Result := ( DateFichier = VHImmo.Encours.Deb );
  end;

var CodeImmo : string ; bEco : boolean ;
    OBImmo : TOB;
    BaseAmort : double;
    Plafond : double;
    QP : double;
    wst : string;
    derog, reint :double;

begin

   // ajout mbo fq 17569
   OkFiscal := false;

  { Si la date de début n'est pas identique à celle de début d'exercice, on ne récupère pas les informations }
  if not _ControleDateValide ( Copy (St,22,8) ) then exit;

  { Récupération du code de l'immobilisation }
  CodeImmo:=CalculCodeImmo (St);

  bEco:=(upperCase(St[21])='E') ;

  UpdateAvancement(HMTitre.Mess[14],'Immo n° '+CodeImmo);

  OBImmo := TrouveImmoDansTOB (CodeImmo);

if OBImmo = nil then
  begin
  OBImmo := TOB.Create ('IMMO',fTOB,-1);
  OBImmo.PutValue('I_IMMO',CodeImmo);
  end;
if IsExoClose(Copy(St,22,8), Copy(St,32,8)) then
begin
  if bEco then
     OBImmo.PutValue('I_REPRISEECO',OBImmo.GetValue('I_REPRISEECO')+Valeur(Copy(St,81,13))/100 )
  else
    OBImmo.PutValue('I_REPRISEFISCAL',OBImmo.GetValue('I_REPRISEFISCAL')+Valeur(Copy(St,81,13))/100);

end ;

if bEco then
begin
  case St[30] of
    'L' : OBImmo.PutValue('I_METHODEECO','LIN');
    'D' : OBImmo.PutValue('I_METHODEECO','DEG');
    'N' : OBImmo.PutValue('I_METHODEECO','NAM');
  end ;
  if st[30] <> 'N' then
  begin
    OBImmo.PutValue('I_TAUXECO',Valeur(Copy(St,55,13))/100) ;
    OBImmo.PutValue('I_DUREEECO',StrToInt(Copy(St,198,5))) ;
    // Cas des durées multipliées par 100 suite reprise SST
    if OBImmo.GetValue('I_DUREEECO')>=1200 then OBImmo.PutValue('I_DUREEECO',Integer(OBImmo.GetValue('I_DUREEECO')/100));

    // attention a la dépréciation
    if (ImportWIQC) and (Copy(St,40,2) = 'DP') then
        OBImmo.PutValue('I_REPRISEECO',
                        (Valeur(Copy(St,120,13))/100)-(Valeur(Copy(St,81,13))/100)-(Valeur(Copy(St,244,13))/100))
    else
       OBImmo.PutValue('I_REPRISEECO',(Valeur(Copy(St,120,13))/100)-(Valeur(Copy(St,81,13))/100));
    { Si le taux et la durée sont nuls, on force la durée à 12 mois }
    if ((OBImmo.GetValue('I_TAUXECO')=0) and (OBImmo.GetValue('I_DUREEECO')=0)) then
      OBImmo.PutValue('I_DUREEECO',12) ;
  end;

  // ajout mbo pour winner immo - fq 19343
  // date début amortissement éco
  wst := copy(St, 348, 8);
  if ImportWIQC then
  begin
    OBImmo.PutValue('I_DATEDEBECO',FormateDate(wSt));
    // recup des antérieurs dépréciation
    if Copy(St,40,2) = 'DP' then
       OBImmo.PutValue('I_REPRISEDEP',Valeur(Copy(St,244,13))/100 );
  end;

  // modif mbo - fq 19343 - la base éco peut être différente de la base fiscale
  BaseAmort := ( Valeur(Trim(Copy(St,42,13)))/100 );
  if BaseAmort <> 0 then
  begin
    OBImmo.PutValue('I_BASEECO',BaseAmort);
  end;

  OBImmo.PutValue('I_REPRISEDR',0);
  OBImmo.PutValue('I_REPRISEFEC',0);
  OBImmo.PutValue('I_REPRISEFECCEDEE',0);
  OBImmo.PutValue('I_REPRISEFDRCEDEE',0);

end else
begin
  case St[30] of
    'L' : OBImmo.PutValue('I_METHODEFISC','LIN') ;
    'D' : OBImmo.PutValue('I_METHODEFISC','DEG') ;
    'N' : OBImmo.PutValue('I_METHODEFISC','NAM') ;
  end ;
  if st[30]<>'N' then
  begin
    OkFiscal := true;   // ajout mbo fq 17569
    OBImmo.PutValue('I_TAUXFISC',Valeur(Copy(St,55,13))/100);
    OBImmo.PutValue('I_DUREEFISC',StrToInt(Copy(St,198,5)));
    // Cas des durées multipliées par 100 suite reprise SST
    if OBImmo.GetValue('I_DUREEFISC')>=1200 then OBImmo.PutValue('I_DUREEFISC',Integer(OBImmo.GetValue('I_DUREEFISC')/100));

    // attention a la dépréciation
    if (ImportWIQC) and (Copy(St,40,2) = 'DP') then
        OBImmo.PutValue('I_REPRISEFISCAL',
                        (Valeur(Copy(St,120,13))/100)-(Valeur(Copy(St,81,13))/100)-(Valeur(Copy(St,244,13))/100))
    else
        OBImmo.PutValue('I_REPRISEFISCAL',(Valeur(Copy(St,120,13))/100)-(Valeur(Copy(St,81,13))/100));

    OBImmo.PutValue('I_REPRISEDR', 0);
    OBImmo.PutValue('I_REPRISEFEC', 0);

    if OBImmo.GetValue('I_REPRISEFISCAL') > OBImmo.GetValue('I_REPRISEECO') then
    begin
       //ajout pour chantier fiscal
       derog := (OBImmo.GetValue('I_REPRISEFISCAL') - OBImmo.GetValue('I_REPRISEECO'));
       derog := arrondi(derog, V_PGI.OkDecV);
       OBImmo.PutValue('I_REPRISEDR', derog);
    end else
    begin
       // ajout FQ 21454 -
       if OBImmo.GetValue('I_REPRISEFISCAL') < OBImmo.GetValue('I_REPRISEECO') then
       begin
          reint := (OBImmo.GetValue('I_REPRISEECO') - OBImmo.GetValue('I_REPRISEFISCAL'));
          reint := arrondi(reint, V_PGI.OkDecV);
          OBImmo.PutValue('I_REPRISEFEC', reint);
       end;
    end;

  end;

  // ajout mbo pour winner immo - fq 19343
  // Date début amortissement fiscal
  wst := copy(st, 348, 8);
  if ImportWIQC then
  begin
     OBImmo.PutValue('I_DATEDEBFIS',FormateDate(wSt)) ;
  end;
  { FQ 14306 - Récupération de la base d'amortissement }
  // modif mbo - fq 19500 - la base éco peut être différente de la base fiscale
  BaseAmort := ( Valeur(Trim(Copy(St,42,13)))/100 );
  if BaseAmort <> 0 then
  begin
    //OBImmo.PutValue('I_BASEECO',BaseAmort);
    OBImmo.PutValue('I_BASEFISC',BaseAmort);
  end;

  //ajout mbo 13.07.07 gestion fiscale - FQ 21162
  if St[394] = 'O' then
     OBImmo.PutValue('I_NONDED', 'X');

  //ajout mbo 13.07.07 remplacement de composant - FQ 21162
  if St[395] = 'O' then
  begin
     OBImmo.PutValue('I_STRING1','X');
     remplacement := true;
  end else
     remplacement := false;   // ajout pour fq 21657

end ;
  { FQ 14306 - Récupération de la base d'amortissement }
  // modif mbo - fq 19500 - la base éco peut être différente de la base fiscale
  //BaseAmort := ( Valeur(Trim(Copy(St,42,13)))/100 );
  //if BaseAmort <> 0 then
  //begin
    //OBImmo.PutValue('I_BASEECO',BaseAmort);
    //OBImmo.PutValue('I_BASEFISC',BaseAmort);
  //end;

  // MVG - 10/11/2006 FQ 19092
  if bEco then
  begin
     Plafond :=( Valeur(Trim(Copy(St,185,13)))/100 );
     if Plafond<BaseAmort then OBImmo.PutValue('I_REINTEGRATION',Plafond) else OBImmo.PutValue('I_REINTEGRATION',0.00);
     QP := Valeur(Trim(Copy(St,390,2)));
     OBImmo.PutValue('I_QUOTEPART',QP);
  end;


end ;

procedure TFImpCegid.ImporteImmo2Cb(St : string) ;
var CodeImmo : string ;
    sTmp : string;
    OBEche,OBImmo : TOB;
begin
  CodeImmo := CalculCodeImmo (St);
  UpdateAvancement(HMTitre.Mess[15],'Immo n° '+CodeImmo);
  OBImmo := TrouveImmoDansTOB (CodeImmo);
  if OBImmo = nil then
  begin
    OBImmo := TOB.Create ('IMMO',fTOB,-1);
    OBImmo.PutValue('I_IMMO',CodeImmo);
  end;
  if (fCurContrat = nil) or (fCurContrat.sCode <> CodeImmo) then
  begin
    if fCurContrat <> nil then
    begin
      // mbo 29.12.06
      if not ImportWIQC then
         fCurContrat.CalculEcheances
      else
         fCurContrat.MetTrancheDansEcheance;

      OBEche := TOB.Create ('',OBImmo,-1);
      fCurContrat.MajTableEcheancesTOB (OBEche);
      fCurContrat.Free;

    end;

    fCurContrat := TImContrat.Create;
  end;
  OBImmo.PutValue('I_COMPTEIMMO',AjusteCompteImmo(Copy(St,21,10))) ;
  OBImmo.PutValue('I_NUMCONTRATCB',Trim(Copy(St,31,15))) ;
  OBImmo.PutValue('I_ORGANISMECB',Trim(Copy(St,47,25))) ;
//  ControleEtEnregistreCompte (OBImmo.GetValue('I_ORGANISMECB'),True); à voir
  OBImmo.PutValue('I_TYPELOYERCB','LCO');
  sTmp := Trim(Copy(St,72,1));
  if sTmp = 'M' then OBImmo.PutValue('I_PERIODICITE','MEN')
  else   if sTmp = 'T' then OBImmo.PutValue('I_PERIODICITE','TRI')
  else   if sTmp = 'S' then OBImmo.PutValue('I_PERIODICITE','SEM')
  else   if sTmp = 'A' then OBImmo.PutValue('I_PERIODICITE','ANN');
  sTmp := Trim(Copy(St,73,1));
  if sTmp= 'A' then OBImmo.PutValue('I_VERSEMENTCB','AVA')
  else OBImmo.PutValue('I_VERSEMENTCB','ECH');
  if (Copy(St,74,1)='L') then OBImmo.PutValue('I_TYPELOYERCB','LCO')
  else OBImmo.PutValue('I_TYPELOYERCB','LVA');
  OBImmo.PutValue('I_DATEDEBUTECH',FormateDate(Trim(Copy(St,75,8)))) ;
  OBImmo.PutValue('I_DATEFINECH',FormateDate(Trim(Copy(St,83,8)))) ;
  EnregDepotGarantieT(OBImmo,St) ;
{
  La fiche dépôt de garantie n'existant pas physiquement dans le fichier sav,
  on ne récupère pas le dépôt de garantie. - CA le 01/03/2000
  if (Copy(St,91,10)<>'          ') then
  begin
    if (cbNumInc.Checked) or (cbIncAchat.Checked) then
    begin
      OBImmo.PutValue('I_IMMOLIEGAR',Format('%.10d',[fTabCode.IndexOf(AjusteCompteImmo(Copy(St,91,10))+Trim(Copy(St,101,6)))+1]));
    end
    else
    OBImmo.PutValue('I_IMMOLIEGAR',AjusteCompteImmo(Copy(St,91,10))+Trim(Copy(St,101,6))) ;
  end;}
  OBImmo.PutValue('I_MONTANTPREMECHE',Valeur(Trim(Copy(St,108,13)))/100) ;
  OBImmo.PutValue('I_FRAISECHE',Valeur(Trim(Copy(St,121,13)))/100) ;
  OBImmo.PutValue('I_MONTANTSUIVECHE',Valeur(Trim(Copy(St,134,13)))/100);
  OBImmo.PutValue('I_RESIDUEL',Valeur(Trim(Copy(St,160,13)))/100);

  OBImmo.PutValue('I_REPRISEDR',0);
  OBImmo.PutValue('I_REPRISEFEC',0);
  OBImmo.PutValue('I_REPRISEFECCEDEE',0);
  OBImmo.PutValue('I_REPRISEFDRCEDEE',0);
  
  fCurContrat.ChargeTOB (OBImmo);
  if fCurContrat.sTypeLoyer = 'LCO' then
  begin
    // mbo 29.12.06 - si import winner on traite les enreg de type FPEC
    if not ImportWIQC then
    begin
       fCurContrat.CalculEcheances;
       OBEche := TOB.Create ('',OBImmo, -1);
       fCurContrat.MajTableEcheancesTOB (OBEche);
       fCurContrat.Free;
       fCurContrat := nil;
    end;
    //fCurContrat.Free;
    //fCurContrat := nil;
  end;
end;

procedure TFImpCegid.EnregDepotGarantieT(T1 : TOB;St: string);
var T2: TOB ; mDepot,wCode,wCompte,CodeImmo: string ;
begin
  MDepot:=trim(copy(sT,173,13)) ;
  wCompte:=copy(sT,91,10) ;
  wCode:=copy(sT,101,6) ;
  if (T1=nil) or (Valeur(MDepot)=0) then exit ;
  CodeImmo:=CalculCodeImmo(wCompte,wCode) ;
  T2:=Tob.Create('IMMO',fTOB,-1) ;
  T2.PutValue('I_IMMO',CodeImmo) ;
  T2.PutValue('I_ABREGE',CodeImmo) ;
  T2.PutValue('I_NATUREIMMO','FI') ;
  T2.PutValue('I_LIBELLE','Dépôt de garantie '+T1.GetValue('I_LIBELLE')) ;
  T2.PutValue('I_MONTANTHT',Valeur(MDepot)/100) ;
  // CA - 23/04/2002 - Mise à jour de la valeur d'achat
  T2.PutValue('I_VALEURACHAT', Valeur(MDepot)/100);
  T2.PutValue('I_QUANTITE',1) ;
  T2.PutValue('I_TVARECUPERABLE',0) ;
  T2.PutValue('I_TVARECUPEREE',0) ;
  T2.PutValue('I_BASETAXEPRO',0) ;
  T2.PutValue('I_METHODEECO','') ;
  T2.PutValue('I_QUALIFIMMO',T1.GetValue('I_QUALIFIMMO')) ;
  T2.PutValue('I_COMPTEIMMO',AjusteCompteImmo(wCompte)) ;
  T2.PutValue('I_COMPTEREF',AjusteCompteImmo(wCompte)) ;
  T2.PutValue('I_DATEPIECEA',T1.GetValue('I_DATEPIECEA')) ;
  T2.PutValue('I_CODEPOSTAL',T1.GetValue('I_CODEPOSTAL')) ;
  T2.PutValue('I_VILLE',T1.GetValue('I_VILLE')) ;
  T2.PutValue('I_PAYS',T1.GetValue('I_PAYS')) ;
  T2.PutValue('I_ETABLISSEMENT',T1.GetValue('I_ETABLISSEMENT')) ;
  T2.PutValue('I_LIEUGEO',T1.GetValue('I_LIEUGEO')) ;
  T2.PutValue('I_IMMOLIEGAR',T1.GetValue('I_IMMO')) ;
  T2.PutValue('I_ETAT','OUV') ;
  T2.PutValue('I_OPEMUTATION','-') ;
  T2.PutValue('I_OPEECLATEMENT','-') ;
  T2.PutValue('I_OPECESSION','-') ;
  T2.PutValue('I_OPECHANGEPLAN','-') ;
  T2.PutValue('I_OPELIEUGEO','-') ;
  T2.PutValue('I_OPEETABLISSEMENT','-') ;
  T2.PutValue('I_OPELEVEEOPTION','-') ;
  T2.PutValue('I_OPEMODIFBASES','-') ;
  // BTY 01/06 FQ 17259 Nouveau top dépréciation
  T2.PutValue('I_OPEDEPREC','-') ;
  // BTY 04/06 FQ 17516
  T2.PutValue('I_OPEREG','-') ;
  T2.PutValue('I_OPERATION','-') ;
  T2.PutValue('I_IMMOLIE',T1.GetValue('I_IMMO')) ;
  // TGA 26/01/2006
  T2.PutValue('I_REPRISEDEP',0);
  T2.PutValue('I_ECCLEECR','');
  T2.PutValue('I_DOCGUID','');
  T2.PutValue('I_DATEFINCB',iDate1900);
  // TGA 04/04/2006
  T2.PutValue('I_REPRISEDEPCEDEE',0);
  T2.PutValue('I_SUSDEF','A');
  T2.PutValue('I_REGLECESSION','NOR');
  // en attente T2.PutValue('I_NONDED','-');
  T2.PutValue('I_REPRISEINT',0);
  T2.PutValue('I_REPRISEINTCEDEE',0);
  T2.PutValue('I_DPI','-');
  T2.PutValue('I_DPIEC','-');
  T2.PutValue('I_CORRECTIONVR',0);
  T2.PutValue('I_CORVRCEDDE',0);
  T2.PutValue('I_SUBVENTION','NON');
  T2.PutValue('I_SBVPRI',0);
  T2.PutValue('I_SBVMTC',0);
  T2.PutValue('I_SBVPRIC',0);
  T2.PutValue('I_SBVEC','C');
  T2.PutValue('I_SBVDATE',iDate1900);
  T2.PutValue('I_CPTSBVR','');
  T2.PutValue('I_CPTSBVB','');
  // TGA 07/09/2006 T2.PutValue('I_PRF','-');
  T2.PutValue('I_PFR','-');
  T2.PutValue('I_COEFDEG',0);
  T2.PutValue('I_AMTFOR',0);
  T2.PutValue('I_AMTFORC',0);
  T2.PutValue('I_ACHFOR',iDate1900);
  T2.PutValue('I_PRIXACFORC',0);
  T2.PutValue('I_VNCFOR',0);
  T2.PutValue('I_DURRESTFOR',0);
  T2.PutValue('I_DATEDEBECO',T2.GetValue('I_DATEAMORT'));
  T2.PutValue('I_DATEDEBFIS',T2.GetValue('I_DATEAMORT'));
  T2.PutValue('I_REPRISEDR',0);
  T2.PutValue('I_REPRISEFEC',0);
  T2.PutValue('I_REPRISEFECCEDEE',0);
  T2.PutValue('I_REPRISEFDRCEDEE',0);

  //Mise à jour des infos de l'immo de crédit bail
  T1.PutValue('I_IMMOLIEGAR',CodeImmo) ;
  T1.PutValue('I_DATEDEPOTGAR',T1.GetValue('I_DATEPIECEA')) ;
  T1.PutValue('I_DEPOTGARANTIE',Valeur(MDepot)/100) ;


end ;

procedure TFImpCegid.ImporteImmo2Echeance(St : string) ;
var CodeImmo : string ;
    dtDeb, dtFin : TDateTime;
    mtEch, mtFrais : double;
    nEch : integer;
begin
  CodeImmo := CalculCodeImmo (' '+St); // CA le 01/03/2000 : pour 'coller' avec les autres cas
  UpdateAvancement(HMTitre.Mess[16],'Immo n° '+CodeImmo);
  if fCurContrat = nil then exit;
  if fCurContrat.sCode = CodeImmo then
  begin
    dtDeb := FormateDate(Trim(Copy(St,24,8)));
    dtFin := FormateDate(Trim(Copy(St,32,8)));
    mtEch := Valeur(Trim(Copy(St,40,13)))/100 ;
    mtFrais := Valeur(Trim(Copy(St,53,13)))/100 ;
    nEch := StrToInt(Trim(Copy(St,66,5)));
    fCurContrat.AjouteTranche ( dtDeb,dtFin,nEch,mtEch,mtFrais);
  end;


end;

procedure TFImpCegid.ImporteImmo2Compte(St : string) ;
var stNoCompte : string;
begin
  stNoCompte := Copy (St,4,fLgCpteGen);
  UpdateAvancement (HMTitre.Mess[6], HMTitre.Mess[7]+stNoCompte);
  TraiteGeneral (stNoCompte, Copy (St,14,25));
end ;

procedure TFImpCegid.ImporteImmo2Motif(St : string) ;
var OBMotif : TOB;
begin
UpdateAvancement(HMTitre.Mess[16],'');
OBMotif := TOB.Create('CHOIXCOD',fTOBDiv,-1);
OBMotif.PutValue('CC_TYPE','MDC');

{ mbo - 13.09.07 - FQ 21435 - code est sur 3 positions et non sur 1
OBMotif.PutValue('CC_CODE',Trim(Copy (St,4,1)));
OBMotif.PutValue('CC_LIBELLE',Trim(Copy(St,5,15)));
}
OBMotif.PutValue('CC_CODE',Trim(Copy (St,4,3)));
OBMotif.PutValue('CC_LIBELLE',Trim(Copy(St,7,15)));
OBMotif.PutValue('CC_ABREGE',Trim(Copy(St,7,15)));
end ;

procedure TFImpCegid.ImporteImmo2Lieu(St : string) ;
var OBLieu : TOB;
begin
UpdateAvancement(HMTitre.Mess[17],'');
OBLieu := TOB.Create('CHOIXCOD',fTOBDiv,-1);
OBLieu.PutValue('CC_TYPE','GEO');
OBLieu.PutValue('CC_CODE',Trim(Copy (St,4,3)));
OBLieu.PutValue('CC_LIBELLE',Trim(Copy(St,7,15)));
OBLieu.PutValue('CC_ABREGE',Trim(Copy(St,7,15)));
end ;

// fq 19343 - import des lots winner immo dans la table regroupement
procedure TFImpCegid.ImporteImmo2Regroup(St : string) ;
var OBReg : TOB;
    code : string;
begin
UpdateAvancement(HMTitre.Mess[22],'');
code := copy(St, 4, 3);
OBreg := TOB.Create('CHOIXCOD',fTOBDiv,-1);
   if Obreg.FieldExists(code) = false then
   begin
      OBReg.PutValue('CC_TYPE','IGI');
      OBReg.PutValue('CC_CODE',Trim(Copy (St,4,3)));
      OBReg.PutValue('CC_LIBELLE',Trim(Copy(St,7,35)));
      OBReg.PutValue('CC_ABREGE',Trim(Copy(St,7,15)));
      // Ajout CA - 07/12/2007 pour besoin KPMG - 
      if (Length(St) >= 42) then OBReg.PutValue('CC_LIBRE',Trim(Copy(St,42,1)));
   end;
end ;


function TFImpCegid.GetCodeLieu (stCode : string) : string;
var OB : TOB;
begin
  if stCode = '' then stCode := '001';
  OB := fTOBDiv.FindFirst (['CC_TYPE','CC_CODE'],['GEO',stCode],False);
  if OB <> nil then result := OB.GetValue('CC_CODE')
  else
  begin
    OB := TOB.Create('CHOIXCOD',fTOBDiv,-1);
    if OB <> nil then
    begin
      OB.PutValue('CC_TYPE','GEO');
      OB.PutValue('CC_CODE',stCode);
      OB.PutValue('CC_LIBELLE',TraduireMemoire('A définir'));
      OB.PutValue('CC_ABREGE',TraduireMemoire('A définir'));
    end;
    result := stCode;
  end;
end;

function TFImpCegid.GetCodeEtablissement (stCode : string) : string;
var et: integer;
    TEtablissement : Tob;
    Cle : string;
begin
  result := VHImmo^.EtablisDefaut;

  TEtablissement := TOB.Create ('', nil, -1);
  TEtablissement.LoadDetailDB('ETABLISS',Cle,'',nil,False);

  for et:=0 to TEtablissement.Detail.Count - 1 do
  begin
      if stCode = TEtablissement.Detail[et].GetValue('ET_ETABLISSEMENT') then
      begin
         result := stCode;
         break;
      end;
  end;

  TEtablissement.Free;

end;


procedure TFImpCegid.ImporteImmo2Famille(St : string) ;
var  OBDetail : TOB;
begin
UpdateAvancement(HMTitre.Mess[18],'');
if not fbFamille then
begin
OBDetail := TOB.Create('COMMUN',fTOBDiv,-1);
{$IFDEF SERIE1}
OBDetail.PutValue('CO_TYPE','LCL');
OBDetail.PutValue('CO_CODE','I02');
{$ELSE}
OBDetail.PutValue('CO_TYPE','NAT');
OBDetail.PutValue('CO_CODE','I01');
{$ENDIF}
OBDetail.PutValue('CO_LIBELLE',Msg.Mess[36]);
OBDetail.PutValue('CO_ABREGE','X');
OBDetail.PutValue('CO_LIBRE','4');
fbFamille := True;
end;
{$IFDEF SERIE1}
OBDetail := TOB.Create('COMMUN',fTOBDiv,-1);
OBDetail.PutValue('CO_TYPE','LI2');
OBDetail.PutValue('CO_CODE',Copy(St,4,2));
OBDetail.PutValue('CO_LIBELLE',Trim(Copy(St,6,20)));
{$ELSE}
OBDetail := TOB.Create('NATCPTE',fTOBDiv,-1);
OBDetail.PutValue('NT_TYPECPTE','I01');
OBDetail.PutValue('NT_NATURE',Copy(St,4,2));
OBDetail.PutValue('NT_LIBELLE',Copy(St,6,20));
OBDetail.PutValue('NT_DATE1',date);
OBDetail.PutValue('NT_DATE2',date);
OBDetail.PutValue('NT_DATE3',date);
OBDetail.PutValue('NT_DATE4',date);
OBDetail.PutValue('NT_MONTANT1',0);
OBDetail.PutValue('NT_MONTANT1',1);
OBDetail.PutValue('NT_MONTANT1',2);
OBDetail.PutValue('NT_MONTANT1',3);
{$ENDIF}
end ;

procedure TFImpCegid.ImporteImmo2Stat(St : string) ;
var OBStat : TOB;
begin
if not fbStat then
begin
OBStat := TOB.Create('COMMUN',fTOBDiv,-1);
{$IFDEF SERIE1}
OBStat.PutValue('CO_TYPE','LCL');
OBStat.PutValue('CO_CODE','I01');
{$ELSE}
OBStat.PutValue('CO_TYPE','NAT');
OBStat.PutValue('CO_CODE','I00');
{$ENDIF}
OBStat.PutValue('CO_LIBELLE',Msg.Mess[35]);
OBStat.PutValue('CO_ABREGE','X');
OBStat.PutValue('CO_LIBRE','4');
fbStat := True;
end;
{$IFDEF SERIE1}
OBStat := TOB.Create('COMMUN',fTOBDiv,-1);
OBStat.PutValue('CO_TYPE','LI1');
OBStat.PutValue('CO_CODE',Copy(St,4,2));
OBStat.PutValue('CO_LIBELLE',Trim(Copy(St,8,20)));
{$ELSE}
OBStat := TOB.Create('NATCPTE',fTOBDiv,-1);
OBStat.PutValue('NT_TYPECPTE','I00');
OBStat.PutValue('NT_NATURE',Trim(Copy(St,4,4)));
OBStat.PutValue('NT_LIBELLE',Trim(Copy(St,8,20)));
OBStat.PutValue('NT_DATE1',date);
OBStat.PutValue('NT_DATE2',date);
OBStat.PutValue('NT_DATE3',date);
OBStat.PutValue('NT_DATE4',date);
{$ENDIF}
end ;

function TFImpCegid.TrouveImmoDansTob(Code : string) : TOB;
var i : integer;
begin
  result := nil;
  for i:=0 to fTOB.Detail.Count - 1 do
  begin
    if fTOB.detail[i].GetValue('I_IMMO') = Code then
    begin
      result := fTOB.detail[i];
      break;
    end;
  end;
end;

function GetNumeroOrdreSuivant(OBLog : TOB; Champ : string):integer;
var Ordre,i : integer;
begin
  Ordre := 0;
  for i:=0 to OBLog.Detail.Count - 1 do
  begin
    if OBLog.detail[i].GetValue(Champ) > Ordre then
      Ordre := OBLog.detail[i].GetValue(Champ);
  end;
  result := Ordre+1;
end;

procedure TFImpCegid.DetruitListeCode;
var i:integer; ARecord:TCodeEq;
begin
  if fLCode <> nil then
  begin
    for i := 0 to (fLCode.Count - 1) do
    begin
      ARecord := fLCode.Items[i];
      ARecord.Free;
    end;
    fLCode.Free;
    fLCode := nil;
  end;
end;

procedure TFImpCegid.ControleEtEnregistreCompte (stCompte : string);
begin
  if stCompte = '' then exit;
  TraiteGeneral(stCompte,TraduireMemoire('CREE PAR CEGID'));
end;

procedure TFImpCegid.InsertDB;
var Plan : TplanAmort;
    OBImmo, OBPlan, T : TOB;
    i : integer;
    CpteFinInf, CpteFinSup : string;
    MaxMove : integer;
begin
  hent1.EnableControls(Self,false) ;
  CpteFinInf := ImBourreLaDoncSurLesComptes(GetParamSocSecur('SO_CPTEFININF',''),'0');
  CpteFinSup := ImBourreLaDoncSurLesComptes(GetParamSocSecur('SO_CPTEFINSUP',''),'Z');
  MaxMove := 3*fTob.Detail.Count + FTobGenACreer.Detail.Count + fTobLog.Detail.Count + fTobDiv.Detail.Count;
  InitMove ( MaxMove, Caption );
  // Retraitement des immobilisations
  for i:=0 to fTob.Detail.Count - 1 do
  begin
    MoveCur ( False );
    OBImmo := fTob.Detail[i];
    // Mise à jour des immobilisations financières
    if (OBImmo.GetValue('I_COMPTEIMMO')>= CpteFinInf) and
      (OBImmo.GetValue('I_COMPTEIMMO')<= CpteFinSup) then
      OBImmo.PutValue('I_NATUREIMMO','FI');
    // Si pas d'amortissement fiscal, la base fiscale est égale à 0
    if (OBImmo.GetValue('I_METHODEFISC')='')  then OBImmo.PutValue('I_BASEFISC',0);
    // Mise à jour de l'analytique - FQ 10067 - CA - 30/11/2004
    OBImmo.PutValue('I_VENTILABLE','-');
    OBImmo.PutValue('I_VENTILABLE1','-');
    OBImmo.PutValue('I_VENTILABLE2','-');
    OBImmo.PutValue('I_VENTILABLE3','-');
    OBImmo.PutValue('I_VENTILABLE4','-');
    OBImmo.PutValue('I_VENTILABLE5','-');                    
  end;

  // Calcul des plans d'amortissement
  OBImmo := nil;
  try
    for i:=0 to fTob.Detail.Count - 1 do
    begin
      MoveCur ( False );
      OBImmo := fTob.Detail[i];
      // si immo fin., pas de calcul de plan
      if OBImmo.GetValue('I_NATUREIMMO')='FI' then continue;
      UpdateAvancement (HMTitre.Mess[21],'Immo n°'+OBImmo.GetValue ('I_IMMO')  );
      Plan:=TPlanAmort.Create(true) ;// := CreePlan (True);

      // ajout suite import antérieur dépréciation de winner immo
      if (ImportWIQC) and (OBImmo.GetValue('I_REPRISEDEP')<> 0) then
         Plan.DureeRest := true;

      Plan.CalculTOB(OBImmo, iDate1900);
      OBPlan := TOB.Create ('',OBImmo,-1);
      Plan.SauveTOB (OBPlan );
      OBImmo.PutValue('I_PLANACTIF',Plan.NumSeq);
      if Plan <> nil then Plan.free ; //Detruit;
    end;
  except
    on E: Exception do
    begin
      ShowError (0, E.Message);
      if OBImmo <> nil then
        ShowError (0,
          OBImmo.GetValue('I_COMPTEREF')+' - '+TrouveAncienCode(OBImmo.GetValue('I_IMMO'))+
          TraduireMemoire(' : problème lors du calcul du plan d''amortissement'));
      V_PGI.IOError := oeUnknown;
      FiniMove;
      exit;
    end;
  end;

  // Mise à jour des comptes généraux
  UpdateAvancement (HMTitre.Mess[8], HMTitre.Mess[9]);
  T := nil;
  try
    for i:=0 to FTobGenACreer.Detail.Count - 1 do
    begin
      MoveCur ( False );
      T := FTobGenACreer.Detail[i];
      T.InsertDB(nil);
    end;
    T := nil;
  except
    on E: Exception do
    begin
      ShowError (0, E.Message);
      if T <> nil then ShowError (0, T.GetValue('G_GENERAL')+
          TraduireMemoire(' : problème lors de la création du compte'));
      V_PGI.IOError := oeUnknown;
      FiniMove;
      exit;
    end;
  end;

  // Mise à jour des fiches et plans
  UpdateAvancement (HMTitre.Mess[8], HMTitre.Mess[10]);
  try
    InitMove( fTOB.Detail.Count , Caption );
    for i:=0 to fTOB.Detail.Count - 1 do
    begin
      MoveCur ( False );
      T := fTOB.Detail[i];
      T.InsertDB(nil);
    end;
    T := nil;
  except
    on E: Exception do
    begin
      ShowError (0, E.Message);
      if T <> nil then
        ShowError (0,
          T.GetValue('I_COMPTEREF')+' - '+TrouveAncienCode(T.GetValue('I_IMMO'))+
          TraduireMemoire(' : problème lors de la mise à jour de la fiche'));
      V_PGI.IOError := oeUnknown;
      FiniMove;
      exit;
    end;
  end;

  // Mise à jour de IMMOLOG
  UpdateAvancement (HMTitre.Mess[8], HMTitre.Mess[11]);
  T := nil;
  try
    for i := 0 to fTOBLog.Detail.Count - 1 do
    begin
      MoveCur (False);
      T := fTOBLog.Detail[i];
      T.InsertDB(nil);
    end;
  except
    on E: Exception do
    begin
      ShowError (0, E.Message);
      if T <> nil then
        ShowError (0,
          TrouveCompteImmo(T.GetValue('IL_IMMO'))+' - '+TrouveAncienCode(T.GetValue('IL_IMMO'))+
          TraduireMemoire(' : problème lors de la mise à jour de l''historique'));
      V_PGI.IOError := oeUnknown;
      FiniMove;
      exit;
    end;
  end;

  // Mise à jour des tables diverses
  UpdateAvancement (HMTitre.Mess[8], HMTitre.Mess[12]);
  try
    for i := 0 to fTOBDiv.Detail.Count - 1 do
    begin
      MoveCur ( False ) ;
      T := fTOBDiv.Detail[i];
      T.InsertOrUpdateDB(True);
    end;
  except
    on E: Exception do
    begin
      ShowError (0, E.Message);
      ShowError (TIMP_TABLESANNEXES);
      V_PGI.IOError := oeUnknown;
      FiniMove;
      exit;
    end;
  end;

  FiniMove;
  hent1.EnableControls(Self,true) ;

  AvertirTable('TILIEUGEO');         // AJOUT MBO POUR FQ 19846
  AvertirTable('AMREGROUPEMENT');    // mbo 31.10.07 FQ 21770
  AvertirTable('TIMOTIFCESSION');    // mbo 31.10.07 FQ 21770

end;

function TFImpCegid.IsDossierImmoVide : boolean;
var Q : TQuery;
begin
  Q := OpenSQL ('SELECT I_IMMO FROM IMMO', True);
  result :=  Q.Eof;
  Ferme (Q);
end;

procedure TFImpCegid.FileNameElipsisClick(Sender: TObject);
begin
ImDirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFImpCegid.TimerTimer(Sender: TObject);
begin
  // On stoppe le timer
  Timer.Enabled := false;
  // Si mode auto. , on lance la restauration automatiquement
  if fModeAuto then BValiderClick (nil);
end;

procedure TFImpCegid.SetOnInformation(const Value: TImportImmoInfo);
begin
  FOnInformation := Value;
end;

procedure TFImpCegid.ShowError (Code: integer; St : string = '') ;
begin
  if St = '' then St := GetMsgErr (Code)
  else Code := 0;
  if Assigned (OnInformation) then
    OnInformation (Self, Code, St)
  else
    PgiBox (St, TitreHalley) ;
end ;

function TFImpCegid.GetMsgErr (Code : integer) : string ;
begin
  case Code of
      TIMP_FILENULL : result := 'Fichier non renseigné';
      TIMP_FILENOTEXIST : result := 'Fichier introuvable';
      TIMP_ERRORINSERTDB : result := 'Problème lors de la mise à jour de la base';
      TIMP_DATEINCOHERENTE : result := 'Dates du fichier incohérentes avec les exercices du dossier.';
      TIMP_CREATIONEXOFAILED : result := 'Création de l''exercice impossible';
      TIMP_LGCPTEINCOHERENTE : result := 'Longueur des comptes incohérente avec le paramétrage du dossier';
      TIMP_TABLESANNEXES : result := 'Erreur lors de la mise à jour des tables annexes.';
  end ;
end ;

procedure TFImpCegid.OnInformationImport(Sender : TObject ; ErrCode : integer ; ErrMsg : string);
begin
  UpdateFichierRapport (ErrMsg);
end;

procedure TFImpCegid.UpdateAvancement (St1, St2 : string);
begin
  TITREAVANCEMENT.Caption := St1;
  DETAILAVANCEMENT.Caption := St2;
end;

function TFImpCegid.TrouveAncienCode ( Code : string ) : string;
var i : integer;
begin
  Result := '';
  for i:=0 to fLCode.Count-1 do
  begin
    if (TCodeEq(fLCode.Items[i]).CodeNum=StrToInt64( Code )) then
    begin
      Result := Copy(TCodeEq(fLCode.Items[i]).CodeTheo,VHImmo^.Cpta[fbGene].Lg+1,6);
      break;
    end;
  end;
end;

function TFImpCegid.CalculCodeImmo (St : string;Compte : string='') : string;
var CodeTheorique : string;
    NumCode : Int64;
    stCode, stCompte : string;
    ARecord : TCodeEq;
    i : integer;
begin
  if Compte = '' then
  begin
    stCode := Trim(Copy(St,15,6));
    stCompte := Trim(Copy(St,5,10));
  end else
  begin
    stCode := St;
    stCompte := Compte;
  end;
  CodeTheorique := AjusteCompteImmo(stCompte)+stCode;
  fCodeEnCours := CodeTheorique;
  if ((cbNumInc.Checked) or (cbIncAchat.Checked)) then
  begin
    NumCode := -1;
    for i:=0 to fLCode.Count-1 do
    begin
      if (TCodeEq(fLCode.Items[i]).CodeTheo=CodeTheorique) then
      begin
        NumCode := i+1+fOffsetCodeImmo;
        break;
      end;
    end;
    if NumCode = - 1 then
    begin
      NumCode := fLCode.Count + 1 + fOffsetCodeImmo;
      ARecord := TCodeEq.Create;
      ARecord.CodeTheo := CodeTheorique;
      ARecord.CodeNum := NumCode;
      fLCode.Add (ARecord);
    end;
    Result := Format('%.10d',[NumCode]);
  end else
  begin
    Result := Copy (stCompte,1,4) + stCode;  // pour limiter à 10 caractères;
  end;
end;

procedure TFImpCegid.CreationFichierRapport (stFileName : string);
var F : TextFile;
    stDirRapport : string;
begin
  stDirRapport := ExtractFileDir(stFileName)+'\Rapports';
  if not DirectoryExists(stDirRapport) then CreateDir(stDirRapport);
  fRappFile := ExtractFileDir(stFileName)+'\Rapports\'// ajout me +ExtractFileName(stFileName);
  +'RAPPORT_GLOBAL_'+HGetUserName+'.TXT';
  AssignFile (F, fRappFile);
  if FileExists (fRappFile) then
  {$i-} Append(F)  {$i+}
  else
        Rewrite (F);
  writeln (F,'---------------------------------------------------------------------------------------');
  writeln (F,'Restauration du fichier '+FileName.Text);
  Writeln (F,'Début : '+DateToStr(date)+ ' à '+TimeToStr(time));
  writeln (F,'---------------------------------------------------------------------------------------');
  CloseFile (F);
end;

procedure TFImpCegid.UpdateFichierRapport (st : string);
var F : TextFile;
begin
  AssignFile (F, fRappFile);
  Append (F);
  Writeln (F,st);
  CloseFile (F);
end;

procedure TFImpCegid.CloseFichierRapport (bRestaurationOK : boolean);
var F : TextFile;
    stFileOk : string;
    i : integer;
    Q : TQuery;
    TImp : TOB;
    Fich, Rep : string;
begin
  AssignFile (F, fRappFile);
  Append (F);
  writeln (F,'----------------------------------------------');
  if bRestaurationOK then Writeln (F,'RESTAURATION DU FICHIER OK')
  else writeln (F,'ERREUR EN RESTAURATION DU FICHIER');
// update de trfs5 ajout me  Pour majlot
  if (ctxPCL in V_PGI.PGIContexte)  then
  begin
        Fich := ExtractFileName(FileName.Text);
        Rep  := ExtractFileDir(FileName.Text);
        Q := OpenSQl ('SELECT * FROM TRFFICHIER Where TRF_FICHIER="' +Fich+ '" AND '+
        'TRF_REPERTOIRE="'+ Rep +'\"', TRUE);
        if not Q.EOF then
        begin
             TImp :=TOB.Create('',Nil,-1) ;
             TImp.LoadDetailDB('TRFFICHIER', '', '', Q, TRUE, FALSE);
             if bRestaurationOK then
             TImp.detail[0].PutValue ('TRF_ETATREST', 'ROK')
             else
             TImp.detail[0].PutValue ('TRF_ETATREST', 'NRE');
             TImp.InsertOrUpdateDB(True);
             TImp.free;
        end;
        Ferme(Q);
  end;

  writeln (F,'----------------------------------------------');
  CloseFile (F);
  if bRestaurationOK then
  begin
    for i:= Length(fRappFile) downto 1 do
    begin
       if fRappFile[i]='.' then
       begin
         stFileOK := Copy (fRappFile,1,i)+'OK';
         break;
       end;
    end;
    if stFileOK <> '' then
    begin
      AssignFile (F, stFileOK);
      Rewrite (F);
      writeln (F,'---------------------------------------------------------------------------------------');
      writeln (F,'Restauration du fichier '+FileName.Text);
      Writeln (F,'Terminée : '+DateToStr(date)+ ' à '+TimeToStr(time));
      writeln (F,'---------------------------------------------------------------------------------------');
      CloseFile (F);
    end;
  end;
end;

procedure MajParamSocImmo;
type
  TParamSocCompte = record
      N : string; // SOC_NOM
      D : string; // SOC_DATA
      C : string; // Caractère de bourrage
    end;
  TParamSocImmo = array[1..28] of TParamSocCompte;
const
  PARAMSOC_IMMO : TParamSocImmo = (( N : 'SO_CPTEAMORTINF'; D : '28'; C : '0'),
    ( N : 'SO_CPTEAMORTSUP'; D : '283'; C : '9'),
    ( N : 'SO_CPTECBINF'; D : '612'; C : '0'),( N : 'SO_CPTECBSUP'; D : '612'; C : '9'),
    ( N : 'SO_CPTEDEPOTINF'; D : '275'; C : '0'),( N : 'SO_CPTEDEPOTSUP'; D : '275'; C : '9'),
    ( N : 'SO_CPTEDEROGINF'; D : '145'; C : '0'),( N : 'SO_CPTEDEROGSUP'; D : '145'; C : '9'),
    ( N : 'SO_CPTEDOTEXCINF'; D : '6871'; C : '0'),( N : 'SO_CPTEDOTEXCSUP'; D : '6871'; C : '9'),
    ( N : 'SO_CPTEDOTINF'; D : '6811'; C : '0'),( N : 'SO_CPTEDOTSUP'; D : '6811'; C : '9'),
    ( N : 'SO_CPTEEXPLOITINF'; D : '7811'; C : '0'),( N : 'SO_CPTEEXPLOITSUP'; D : '7811'; C : '9'),
    ( N : 'SO_CPTEFININF'; D : '26'; C : '0'),( N : 'SO_CPTEFINSUP'; D : '27'; C : '9'),
    ( N : 'SO_CPTEIMMOINF'; D : '2'; C : '0'),( N : 'SO_CPTEIMMOSUP'; D : '23'; C : '9'),
    ( N : 'SO_CPTELOCINF'; D : '613'; C : '0'),( N : 'SO_CPTELOCSUP'; D : '613'; C : '9'),
    ( N : 'SO_CPTEPROVDERINF'; D : '68725'; C : '0'),( N : 'SO_CPTEPROVDERSUP'; D : '68725'; C : '9'),
    ( N : 'SO_CPTEREPDERINF'; D : '78725'; C : '0'),( N : 'SO_CPTEREPDERSUP'; D : '78725'; C : '9'),
    ( N : 'SO_CPTEREPEXCINF'; D : '7871'; C : '0'),( N : 'SO_CPTEREPEXCSUP'; D : '7871'; C : '9'),
    ( N : 'SO_CPTEVACEDEEINF'; D : '675'; C : '0'),( N : 'SO_CPTEVACEDEESUP'; D : '675'; C : '9'));
var i : integer;
begin
  for i:=1 to High(PARAMSOC_IMMO) do
  begin
    SetParamSoc(PARAMSOC_IMMO[i].N,
      ImBourreLaDoncSurLesComptes(PARAMSOC_IMMO[i].D,PARAMSOC_IMMO[i].C));
  end;
end;

procedure TFImpCegid.FormCreate(Sender: TObject);
begin

{$IFDEF SERIE1}
HelpContext:=400600 ;
{$ELSE}
//HelpContext:=0 ;
{$ENDIF}
end;

procedure TFImpCegid.TraiteGeneral ( stGeneral, stLibelle : string );
var T : TOB;
begin
  // Le compte est-il présent dans la liste des comptes existants ?
  T := FTobGenExiste.FindFirst(['G_GENERAL'],[stGeneral],False);
  if T = nil then // Si non
  begin
    // Le compte est-il présent dans la liste des comptes à créer ?
    T := FTobGenACreer.FindFirst(['G_GENERAL'],[stGeneral],False);
    if T = nil then  // si non
    begin
      // Le compte est -il présent dans la base ?
      if not FTobGenExiste.LoadDetailDB('GENERAUX','"'+stGeneral+'"','',nil,True) then // pas trouvé
      begin
        // Création du compte
        T := TOB.Create ('GENERAUX',fTOBGenACreer,-1);
        T.PutValue ('G_GENERAL',stGeneral);
        T.PutValue ('G_LIBELLE',stLibelle);
{$IFDEF SERIE1}
{$ELSE}
        T.PutValue ('G_ABREGE',Copy (stLibelle,1,17));
        T.PutValue ('G_SENS','M');
{$ENDIF}
        if stGeneral[1] = '2' then T.PutValue ('G_NATUREGENE','IMO')
        else if stGeneral[1] = '6' then T.PutValue ('G_NATUREGENE','CHA')
        else if stGeneral[1] = '7' then T.PutValue ('G_NATUREGENE','PRO')
        else T.PutValue ('G_NATUREGENE','DIV');
        T.PutValue ('G_VENTILABLE1','-'); T.PutValue ('G_VENTILABLE2','-');
        T.PutValue ('G_VENTILABLE3','-'); T.PutValue ('G_VENTILABLE4','-');
        T.PutValue ('G_VENTILABLE5','-');
{$IFDEF SERIE1}
{$ELSE}
        if VH^.RecupSISCOPGI then T.PutValue ('G_CREERPAR','REC')
        else T.PutValue ('G_CREERPAR','IMP');
{$ENDIF}
        T.PutValue ('G_CONFIDENTIEL','0');
        // pour le rapport
        fInfoImp.TCpte.Add(stGeneral);
      end;
    end;
  end
end;

function TFImpCegid.TrouveCompteImmo(Code: string): string;
var T : TOB;
begin
  T := fTOB.FindFirst(['I_IMMO'],[Code],False);
  if T<>nil then Result := T.GetValue('I_COMPTEIMMO')
  else Result := '';
end;

// FQ 20256
procedure TFImpCegid.TrouveTypeDerogatoireDesImmos;
var i: integer;
begin
  for i:=0 to fTOB.Detail.Count - 1 do
     fTOB.Detail[i].PutValue('I_TYPEDEROGLIA', TypeDerogatoire(fTOB.Detail[i], nil) );
end;


procedure TFImpCegid.InitDossierEuro;
var TD, TDetail : TOB;
begin
  SetParamSoc('SO_TENUEEURO', True);
  SetParamSoc('SO_TAUXEURO', 1);
  TD := TOB.Create('', nil, -1);
  TD.LoadDetailDB('DEVISE', '', '', nil, True);
  TDetail := TD.FindFirst(['D_DEVISE'], ['EUR'], False);
  if (GetParamSocSecur('SO_TENUEEURO',False) = True) then
  begin
    if (TDetail = nil) then
    begin
      TDetail := TOB.Create('DEVISE', TD, -1);
      TDetail.PutValue('D_DEVISE', 'EUR');
      TDetail.PutValue('D_LIBELLE', 'Euro');
      TDetail.PutValue('D_SYMBOLE', '€');
      TDetail.PutValue('D_FERME', '-');
      TDetail.PutValue('D_DECIMALE', 2);
      TDetail.PutValue('D_QUOTITE', 1);
      TDetail.PutValue('D_SOCIETE', V_PGI.CodeSociete);
      TDetail.PutValue('D_MONNAIEIN', '-');
      TDetail.PutValue('D_FONGIBLE', '-');
      TDetail.PutValue('D_PARITEEURO', 1);
    end
    else
    begin
      TDetail.PutValue('D_FERME', '-');
      TDetail.PutValue('D_DECIMALE', 2);
      TDetail.PutValue('D_QUOTITE', 1);
      TDetail.PutValue('D_MONNAIEIN', '-');
      TDetail.PutValue('D_FONGIBLE', '-');
      TDetail.PutValue('D_PARITEEURO', 1);
    end;
    TD.InsertOrUpdateDB(True);
    SetParamSoc('SO_DEVISEPRINC', 'EUR');
  end;
  TD.Free;
  // Suppression des monnaie 'In'
  ExecuteSQL ('DELETE FROM DEVISE WHERE D_MONNAIEIN="X" AND D_FONGIBLE="-"');
end;


procedure TFImpCegid.BFermeClick(Sender: TObject);
begin
// FQ 19903 - MVG 21/05/07
  Close ;
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
end;

end.
