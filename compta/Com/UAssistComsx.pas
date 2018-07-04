{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 23/07/2003
Modifié le ... :   /  /
Description .. : Export du fichier format PGI S5 format Standard ou etendu
Mots clefs ... :
****************************************************************}

unit UAssistComsx;

interface

uses
  Windows, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Mailol,
  assist, ComCtrls, ExtCtrls, hmsgbox, StdCtrls, HTB97, Hctrls, printers, Mask,
  {$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
  {$ENDIF}
  UTOB, Paramsoc, UtilTrans, Grids, Ent1, LookUp, HEnt1, HXLSPAS,
  Recordcom, CloPerio,
{$IFDEF EAGLCLIENT}
  uLanceProcess,
  uHTTP,
  MainEAGL,
  uWA,
{$ELSE}
  UExportCom,
  Fe_main,
{$ENDIF}
{$IFDEF SCANGED}
  galSystem,
{$ENDIF}
  RecupUtil,
  UTOZ, inifiles, GalOutil,
  Shellapi, HSysMenu, HPanel, Vierge,
  cbpPath, // TcbpPath.GetCegidUserTempPath
  uYFILESTD, UObjFiltres, filtre, Menus, ULibExercice;

type

  TFAssistCom = class(TFAssist)
    Choix: TTabSheet;
    Label7: TLabel;
    Bevel2: TBevel;
    HLabel2: THLabel;
    NATURETRANSFERT: THValComboBox;
    HISTOSYNCHRO: THValComboBox;
    TTYPEFORMAT: THLabel;
    TYPEFORMAT: THValComboBox;
    DATEECR: THCritMaskEdit;
    TDATEECR: THLabel;
    TypeEcr: TTabSheet;
    GRPBAL: TGroupBox;
    AUXILIARE: TCheckBox;
    BANA: TCheckBox;
    GRTYPECR: TGroupBox;
    Mail: TTabSheet;
    Label2: TLabel;
    Label10: TLabel;
    Bevel5: TBevel;
    Label11: TLabel;
    Label12: TLabel;
    FEMail: THCritMaskEdit;
    FHigh: TCheckBox;
    FFile: TRadioButton;
    FMail: TRadioButton;
    Fmonexpert: TRadioButton;
    PPARAMGENE: TTabSheet;
    TEXERCICE: THLabel;
    EXERCICE: THValComboBox;
    TNATUREJRL: THLabel;
    NATUREJRL: THValComboBox;
    TJOURNAUX: THLabel;
    JOURNAUX: THMultiValComboBox;
    Label1: TLabel;
    TDATEECR1: THLabel;
    TDATEECR2: THLabel;
    Bevel8: TBevel;
    DATEECR1: THCritMaskEdit;
    DATEECR2: THCritMaskEdit;
    TFETAB: THLabel;
    THISTOSYNCHRO: THLabel;
    PPARAM: TGroupBox;
    PGEN: TCheckBox;
    PTIERS: TCheckBox;
    PSectionana: TCheckBox;
    PJRL: TCheckBox;
    PANA: TCheckBox;
    PARAMSISCOII: TTabSheet;
    HLabel1: THLabel;
    CP_STAT: THValComboBox;
    GroupBox2: TGroupBox;
    SANCOLL: TCheckBox;
    COLLECTIFS: THValComboBox;
    COLLECT: THCritMaskEdit;
    CARAUX: THValComboBox;
    Label4: TLabel;
    AUX1: THCritMaskEdit;
    Label5: TLabel;
    AUX2: THCritMaskEdit;
    AJOUTAUX: TToolbarButton97;
    BDelete1: TToolbarButton97;
    THGRID: THGrid;
    Resume: TTabSheet;
    Label13: TLabel;
    Bevel6: TBevel;
    Bevel3: TBevel;
    FLib1: TLabel;
    FVal1: TLabel;
    FLib2: TLabel;
    FVal2: TLabel;
    FLib3: TLabel;
    FVal3: TLabel;
    FLib4: TLabel;
    FVal4: TLabel;
    PEXPORT: TTabSheet;
    LISTEEXPORT: TListBox;
    FICHENAME: THCritMaskEdit;
    Label3: TLabel;
    ZIPPE: TCheckBox;
    Label8: TLabel;
    Bevel4: TBevel;
    Label9: TLabel;
    Bevel1: TBevel;
    FTimer: TTimer;
    lEtap: THLabel;
    Label14: TLabel;
    Bevel7: TBevel;
    PrinterSetupDialog1: TPrinterSetupDialog;
    Label15: TLabel;
    bClickVisu: TToolbarButton97;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    FETAB: THMultiValComboBox;
    BNetExpert: TCheckBox;
    PTiersautre: TCheckBox;
    AvecLettrage: TCheckBox;
    BEXPORT: TCheckBox;
    GroupBox3: TGroupBox;
    TDATEMODIF1: TLabel;
    DATEMODIF1: THCritMaskEdit;
    TDATEMODIF2: TLabel;
    DATEMODIF2: THCritMaskEdit;
    ANNULVALIDATION: TCheckBox;
    GComptes: TGroupBox;
    TCPTEDEBUT: THLabel;
    TCPTEFIN: THLabel;
    SuppCarAuxi: TCheckBox;
    FTypeEcr: THMultiValComboBox;
    TFTypeEcr: THLabel;
    TLibre: TCheckBox;
    PARAMDOS: TCheckBox;
    PPARAMSOC: TCheckBox;
    PEXERCICE: TCheckBox;
    PECRBUD: TCheckBox;
    TExclure: THLabel;
    Exclure: THMultiValComboBox;
    TRANSFERTVERS: THRadioGroup;
    BGed: TCheckBox;
    Dock971: TDock97;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    GESTIONETAB: TCheckBox;
    BIni: TToolbarButton97;
    CPTEDEBUT: THCritMaskEdit;
    CPTEFIN: THCritMaskEdit;
    Quadra: TCheckBox;
    TDATEBAL: THLabel;
    DATARRETEBAL: THCritMaskEdit;
    TNature: THMultiValComboBox;
    HNature: THLabel;
    PVentil: TCheckBox;
    FCorpsMail: THMemo;
    GroupBox4: TGroupBox;
    Z_C1: THValComboBox;
    ZO1: THValComboBox;
    ZV1: TEdit;
    CBLib: TCheckBox;
    procedure TRANSFERTVERSClick(Sender: TObject);
    procedure NATURETRANSFERTChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure bFinClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure NATUREJRLClick(Sender: TObject);
    procedure PARAMDOSClick(Sender: TObject);
    procedure THGRIDColEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure COLLECTIFSChange(Sender: TObject);
    procedure COLLECTExit(Sender: TObject);
    procedure CARAUXChange(Sender: TObject);
    procedure BDelete1Click(Sender: TObject);
    procedure AJOUTAUXClick(Sender: TObject);
    procedure COLLECTElipsisClick(Sender: TObject);
    procedure FTimerTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure EXERCICEChange(Sender: TObject);
    procedure DATEECR2Exit(Sender: TObject);
    procedure SANCOLLClick(Sender: TObject);
    procedure AUX1Click(Sender: TObject);
    procedure AUX1Exit(Sender: TObject);
    procedure AUX2Click(Sender: TObject);
    procedure AUX2Exit(Sender: TObject);
    procedure FFileClick(Sender: TObject);
    procedure FICHENAMEExit(Sender: TObject);
    procedure bClickVisuClick(Sender: TObject);
    procedure BNetExpertClick(Sender: TObject);
    procedure DATEECRChange(Sender: TObject);
    procedure DATEECREnter(Sender: TObject);
    procedure DATEECRExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure PEXERCICEClick(Sender: TObject);
    procedure PPARAMSOCClick(Sender: TObject);
    procedure PECRBUDClick(Sender: TObject);
    procedure PGENClick(Sender: TObject);
    procedure SuppCarAuxiClick(Sender: TObject);
    procedure PTiersautreClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure GESTIONETABClick(Sender: TObject);
    procedure BIniClick(Sender: TObject);
    procedure CPTEDEBUTElipsisClick(Sender: TObject);
    procedure DATARRETEBALExit(Sender: TObject);
    procedure DATARRETEBALChange(Sender: TObject);
    procedure DATARRETEBALEnter(Sender: TObject);
    procedure PVentilClick(Sender: TObject);
    procedure CBLibClick(Sender: TObject);
  private
    { Déclarations privées }
    Trans                               : TFTransfertcom;
    // BVE 28.08.07 : Suivi des validations
    Archivage                           : Boolean;
    CarFou,CarCli                       : string;
    TobCol                              : TOB;
    FGrille                             : THGrid;
    OkExoEncours                        : Boolean;
    ListeCpteGen                        : TList;
    StArgsav                            : string;
    TOBTiers                            : TOB;
    EtablissUnique,OkExport,OkEnter     : Boolean;
    Email,ExclureEnreg                  : string;
    FichierImp,TYPEECRX,ECRINTEGRE      : string;
    PSuiv, NoSEqNet                     : integer;
    ErreurNetexpert                     : Boolean;
    InfoNet,DateFinMois                 : string;
    ListeFichierJoint, WhereAvancee     : string;
    ObjetFiltre                         : TObjFiltre;
    LienGamme                           : string;
    procedure ChargeNatureTransfert(TValue : string);
    procedure RendVisibleChamp;
    procedure TraiteListeJournal (Naturejal : string);
    procedure TraiteListeJournalBudgetaire;
    function  Creationparamcol : Boolean;
    function  Existedejacoll (Tg : THGRID; compte : string; var Rr : integer) : Boolean;
    function  Existedejaaux (Col,T1,T2 : string) : Boolean;
    function  RenseigneArg(var Fichier,Commande : String) : Boolean;
    procedure DoExoToDateOnChange(Exo: THValComboBox; Date1, Date2: THEdit);
    procedure ChargementCpte(ListeCpteGen : TList;Where : string);
    procedure RenWhereEcr( var Where: string; Ext : string);
    Function  ControleSisco : Boolean;
    function  InitTobParam : TOB ;
    Procedure ExportComSx;
    Procedure ExportComSage;
    Function  AfficheListeComExport(Chaine: string; Listecom : TListBox) : Boolean;
    procedure RenvoieCorpsMail (var  Hst : HTStringList);
    procedure ChargeAvances(OnLib : Boolean);
    function  DecodeWhereSup : string ;
    function  Rendw(Champ :string) : string; // fiche 10431
{$IFNDEF EAGLCLIENT}
    procedure FileClickzip (FFile : string=''; Directory : Boolean=FALSE);
    function  NetExpertEnvoi : Boolean;
{$ENDIF}
    procedure RemplirParam(var Fichier,Commande : string);
    Procedure ChangeCaption;
{$IFDEF EAGLCLIENT}
    procedure DownloadEnFile(Data, FichierDeSortie : string);
    procedure ComsxOnCallBack(Sender: TObject) ;
{$ENDIF}
  public
    { Déclarations publiques }
  end;

var
  FAssistCom: TFAssistCom;
  stArg     : string;

const iDate1900=2;

Procedure ExportDonnees(stArgument : string; Mode : Boolean=FALSE; Gamme : string=''; PourArchivage : Boolean=FALSE) ;
procedure COMDExoToDates ( Exo : String3 ; Var D1,D2 : TDateTime ) ;
{$IFDEF COMPTA}
Procedure EnvoiExportParDate (Date1, Date2, Code, Nature : string);
{$ENDIF}

implementation

{$R *.DFM}


uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  RecordComsis,
  PGIVersSisco,
  UtilPGI,
  ImprimeMaquette,
  {$IFDEF COMPTA}
  uObjEtatsChaines,
  {$ENDIF}
  uNEActions ;

Procedure ExportDonnees(stArgument : string; Mode : Boolean=FALSE; Gamme : string=''; PourArchivage : Boolean=FALSE) ;
var
FicIni             : TIniFile;
Fichier,Commande   : string;
ListeFile,tmp      : string;
Hst                : HTStringList;
{$IFDEF EAGLCLIENT}
//lTobResult              : TOB;
LT                      : TOB;
FileArchive,Filename    : string;
//CodeRetour              : integer;
{$ENDIF}
BEGIN
     stArg := stArgument;
     if Not ExJaiLeDroitConcept(TConcept(ccexport),True) then
        exit;

     FAssistCom :=TFAssistCom.Create(Application) ;
     With FAssistCom do
     begin
       try
           BeginTrans ;
           ErreurNetexpert := FALSE;
           LienGamme := Gamme;
           // BVE 28.08.07 : Suivi des validations
           Archivage := PourArchivage;
           if PourArchivage then BNetExpert.Checked := FALSE;
           if not Mode then ShowModal
           else
           begin
                tmp    := stArg;
                if (pos('.INI', UPPERCASE(tmp)) <> 0) then
                begin
                          tmp    := stArg;
                          Fichier := ReadTokenPipe(tmp, ';');
                          if not FileExists(Fichier) then
                          begin
                              PGIInfo ('Le fichier '+ Fichier+ ' n''existe pas','');
                              exit;
                          end;
                          FicIni       := TIniFile.Create(Fichier);
                          ListeFile    := FicIni.ReadString ('COMMANDE', 'LISTECOMMANDE', '');
                          FicIni.Free;
                end;
                Commande := 'COMMANDE';
                if ListeFile = '' Then
                begin
                     RemplirParam (Fichier, Commande);
                     FTimerTimer(Application);
                end
                else
                begin
                     Commande := ListeFile;
                     While Commande <> '' do
                     begin
                          Commande    := ReadTokenPipe(ListeFile, ';');
                          if Commande <> '' then
                          begin
                               LISTEEXPORT.Items.clear;
                               if not assigned(FTimer) then FTimer := TTimer.Create(nil);
                               RemplirParam (Fichier, Commande);
                               FTimerTimer(Application);
                          end;
                     end;
                end;
           end;
{$IFDEF EAGLCLIENT}
                          if ZIPPE.checked then
                          begin
                            FileName := Trans.FichierSortie;
                            Filearchive := ReadTokenPipe (Filename, '.');
                            Filearchive := Filearchive + '.ZIP';
                            FileName := './ECOMSX/'+ AppServer.SessionID ;
                            AppServer.RequestFileTo (Filearchive, FileName+'/'+ExtractFileName(Filearchive));
                            LT := InitTobParam;
                            with cWA.create do
                            begin
                                 Request('COMSX.SUPPDIR','', LT,'','');
                                 free;
                            end;
                            LT.free;
(*  SI utilisation en base
                            LT := InitTobParam;
                            with cWA.create do
                            begin
                                OnCallBack := ComsxOnCallBack ;
                                FileName := Trans.FichierSortie;
                                Filearchive := ReadTokenPipe (Filename, '.');
                                Filearchive := Filearchive + '.ZIP';
                                LT.Detail.Items[LT.detail.count-1].putValue ('FichierSortie', Filearchive);
                                lTobResult := Request('COMSX.ZIPPE','', LT,'','');
                                Filename := '';
                                CodeRetour :=  AGL_YFILESTD_EXTRACT (Filename, '0011', ExtractFileName(Filearchive), 'COM', 'EXP', 'TMP');
                                if CodeRetour <>-1 then ErreurNetexpert := FALSE
                                else
                                begin
                                  if FileExists(FileName) then
                                  begin
                                         Movefile (PChar(FileName), PChar(Filearchive));
                                         // suppression des fichiers dans la base
                                         ExecuteSQL ('DELETE NFILES WHERE NFI_FILEID IN (SELECT YFS_FILEID FROM YFILESTD WHERE YFS_CODEPRODUIT="0011" AND YFS_CRIT1="COM" AND YFS_CRIT2="EXP" AND YFS_CRIT3="TMP")');
                                         ExecuteSQL ('DELETE NFILEPARTS WHERE NFS_FILEID IN (SELECT YFS_FILEID FROM YFILESTD WHERE YFS_CODEPRODUIT="0011" AND YFS_CRIT1="COM" AND YFS_CRIT2="EXP" AND YFS_CRIT3="TMP")');
                                         ExecuteSQL ('DELETE YFILESTD WHERE YFS_CODEPRODUIT="0011" AND YFS_CRIT1="COM" AND YFS_CRIT2="EXP" AND YFS_CRIT3="TMP"');
                                   end;
                                end;
                            end ;
                            LT.Free;
*)
                          end;

{$ENDIF}
           if not ErreurNetexpert then CommitTrans
           else
           begin
                  RollBack ;
                  FCorpsMail.clear;
//GP                  FCorpsMail.Lines.add ('Veuillez trouver ci-joint le rapport d''erreur correspondant au dossier '+V_PGI.NoDossier);
                  FCorpsMail.Lines.add (Traduirememoire('Veuillez trouver ci-joint le rapport d''erreur correspondant au dossier ')+V_PGI.NoDossier);
                  FCorpsMail.Lines.add (InfoNet );
                  RenvoieCorpsMail (Hst);
                  AglSendMail('Rapport Comsx' ,'synchro-S1-S5@cegid.fr','', Hst,'',true, 1,'','') ;
                  Hst.free;
           end;
          EXCEPT
            On E: Exception do
              begin
              ShowMessage ( 'TozError : ' + E.Message ) ;
              AfficheListeCom('Export annulé ', LISTEEXPORT);
              RollBack ;
              end ;
          end;
          Free;
     end;
END ;


procedure TFAssistCom.TraiteListeJournal (Naturejal : string);
var
  Q : TQuery;
  sLejournal,sRequete,sWhere,sFrom, LibAbreg : String;
  ListJRL : string;
begin
  JOURNAUX.Items.clear;
  JOURNAUX.Values.clear;
  sRequete := 'select J_JOURNAL,J_LIBELLE,J_NATUREJAL ';
  sFrom := 'from JOURNAL ';
  if Naturejal <> '' then
  sWhere := 'where J_NATUREJAL="'+ Naturejal +'" ORDER BY J_JOURNAL';
  Q := OpenSQL(sRequete+sFrom+sWhere,true);
  while not Q.EOF do
    begin
    LibAbreg := Q.FindField('J_LIBELLE').AsString;
    sLejournal := Q.FindField('J_JOURNAL').AsString;
    if ListJRL = '' then ListJRL := sLejournal
    else ListJRL := ListJRL +';'+sLejournal;
    JOURNAUX.Items.Add(LibAbreg);
    JOURNAUX.Values.add(sLejournal);
    Q.Next;
    end;
  Ferme(Q);
  if Naturejal = '' then
//GP     JOURNAUX.text := '<<Tous>>'
     JOURNAUX.text := Traduirememoire('<<Tous>>')
  else
     JOURNAUX.text := ListJRL;

end;

procedure TFAssistCom.TraiteListeJournalBudgetaire;
var
  Q : TQuery;
  sLejournal,sRequete,sWhere,sFrom, LibAbreg : String;
begin
  JOURNAUX.Items.clear;
  JOURNAUX.Values.clear;
  sRequete := 'select BJ_BUDJAL,BJ_LIBELLE ';
  sFrom := 'from BUDJAL ';
  Q := OpenSQL(sRequete+sFrom+sWhere,true);
  while not Q.EOF do
    begin
    LibAbreg := Q.FindField('BJ_LIBELLE').AsString;
    sLejournal := Q.FindField('BJ_BUDJAL').AsString;
    JOURNAUX.Items.Add(LibAbreg);
    JOURNAUX.Values.add(sLejournal);
    Q.Next;
    end;
  Ferme(Q);
//GP  JOURNAUX.text := '<<Tous>>';
  JOURNAUX.text := Traduirememoire('<<Tous>>');
end;


procedure TFAssistCom.PARAMDOSClick(Sender: TObject);
begin
  inherited;
  // pour GI ligne de commande
  if starg <>'' then exit;

  PGEN.Enabled        := PARAMDOS.checked;
  PTIERS.Enabled      := PARAMDOS.checked;
  PTiersautre.Enabled:=  PARAMDOS.checked;
  PSectionana.Enabled := PARAMDOS.checked;
  PJRL.Enabled        := PARAMDOS.checked;
  FTypeEcr.Enabled    := not PARAMDOS.Checked;
  TFTypeEcr.Enabled   := not PARAMDOS.Checked;
  Exclure.Enabled     := not PARAMDOS.Checked;
  TExclure.Enabled    := not PARAMDOS.Checked;

  if not PARAMDOS.checked then
  begin
          EXERCICE.Text := '';
          DATEECR1.Text := stDate1900;
          DATEECR2.Text := stDate2099;
          JOURNAUX.Text := '';
          DATEMODIF1.Enabled  := FALSE;
          DATEMODIF2.Enabled  := FALSE;
          TDATEMODIF1.Enabled := FALSE;
          TDATEMODIF2.Enabled := FALSE;
          GroupBox3.Enabled   := FALSE;
          ANNULVALIDATION.Enabled := TRUE;
  end
  else
  begin
          DATEMODIF1.Enabled  := TRUE;
          DATEMODIF2.Enabled  := TRUE;
          TDATEMODIF1.Enabled := TRUE;
          TDATEMODIF2.Enabled := TRUE;
          GroupBox3.Enabled   := TRUE;
          ANNULVALIDATION.Checked := FALSE;
          ANNULVALIDATION.Enabled := FALSE;
  end;
  ChangeCaption;

end;

procedure TFAssistCom.THGRIDColEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var
  CollectifEncours              : string;
  AuxEncours1,AuxEncours2       : string;
  Vales                         : HTStrings;
begin
  inherited;
     COLLECTIFS.Onchange := nil;
     CARAUX.Onchange := nil;
     CollectifEncours := FGrille.Cells[0, FGrille.Row];
     AuxEncours1 := AGauche(FGrille.Cells[1, FGrille.Row], 10,'0');
     AuxEncours2 := AGauche(FGrille.Cells[2, FGrille.Row], 10,'0');
     if (CollectifEncours <> '') then
     begin
          COLLECT.Text := AGauche(CollectifEncours, VH^.Cpta[fbGene].Lg,'0');
          AUX1.Text    := Copy(AuxEncours1,1,VH^.Cpta[fbGene].Lg);
          AUX2.Text    := Copy(AuxEncours2,1,VH^.Cpta[fbGene].Lg);
          if (Copy(CollectifEncours,1,2) = '41') then
             COLLECTIFS.Value := 'C'
          else
             COLLECTIFS.Value := 'F';
          CARAUX.Items.clear;
          CARAUX.Values.clear;
          Vales := HTStringList.Create;
          if copy (CollectifEncours,1,2) = '40' then
          begin
               Vales.Add('0') ; Vales.Add('F') ;
               CARAUX.Items.Add('0');
               CARAUX.Items.Add('F');
          end;
          if copy (CollectifEncours,1,2) = '41' then
          begin
               Vales.Add('9') ; Vales.Add('C') ;
               CARAUX.Items.Add('9');
               CARAUX.Items.Add('C');
          end;
          CARAUX.Values := Vales;
          Vales.Free;

          CARAUX.Value := Copy(AuxEncours1,1,1);
     end;
end;

procedure TFAssistCom.COLLECTIFSChange(Sender: TObject);
var Vales : HTStrings;
begin
  inherited;
     if COLLECTIFS.Value = 'F' then
     begin
          AUX1.Text := '0';
          AUX2.Text := '0';
          COLLECT.Text := '40';
          CARAUX.Items.clear;
          CARAUX.Values.clear;
          Vales := HTStringList.Create;
          Vales.Add('0') ;
          Vales.Add('F') ;
          CARAUX.Items.Add('0');
          CARAUX.Items.Add('F');
          CARAUX.Values := Vales;
          Vales.Free;
     end;
     if COLLECTIFS.Value = 'C' then
     begin
          AUX1.Text := '9';
          AUX2.Text := '9';
          COLLECT.Text := '41';
          CARAUX.Items.clear;
          CARAUX.Values.clear;
          Vales := HTStringList.Create;
          Vales.Add('9') ;
          Vales.Add('C') ;
          CARAUX.Items.Add('9');
          CARAUX.Items.Add('C');
          CARAUX.Values := Vales;
          Vales.Free;
     end;

end;


procedure TFAssistCom.COLLECTExit(Sender: TObject);
var Vales : HTStrings;
begin
  inherited;
     COLLECT.Text := AGauche(COLLECT.Text, VH^.Cpta[fbGene].Lg,'0');

     if (COLLECTIFS.Value = 'F')
     and (Copy(COLLECT.Text,1,2) <> '40') then
     begin
          PgiInfo ('Le compte collectif doit commencer par 40','Envoyer');
          AUX1.Text := '0';
          AUX2.Text := '0';
          COLLECT.Text := '40';
          CARAUX.Items.clear;
          CARAUX.Values.clear;
          Vales := HTStringList.Create;
          Vales.Add('0') ;
          Vales.Add('F') ;
          CARAUX.Items.Add('0');
          CARAUX.Items.Add('F');
          CARAUX.Values := Vales;
          Vales.Free; exit;
     end;
     if (COLLECTIFS.Value = 'C')
     and (Copy(COLLECT.Text,1,2) <> '41') then
     begin
          PgiInfo ('Le compte collectif doit commencer par 41','Envoyer');
          AUX1.Text := '9';
          AUX2.Text := '9';
          COLLECT.Text := '41';
          CARAUX.Items.clear;
          CARAUX.Values.clear;
          Vales := HTStringList.Create;
          Vales.Add('9') ;
          Vales.Add('C') ;
          CARAUX.Items.Add('9');
          CARAUX.Items.Add('C');
          CARAUX.Values := Vales;
          Vales.Free; exit;
     end;
      if (Copy (COLLECT.Text,1,2)) = '40' then
      begin
           if not ExisteSQL ('SELECT G_GENERAL from GENERAUX Where G_GENERAL="'+
           COLLECT.Text+'" and G_NATUREGENE="COF"') then
           begin
               PgiInfo ('Le compte collectif n''existe pas dans le plan comptable', 'Envoyer');
               COLLECT.SetFocus ;
           end;
      end;
      if (Copy (COLLECT.Text,1,2)) = '41' then
      begin
           if not ExisteSQL ('SELECT G_GENERAL from GENERAUX Where G_GENERAL="'+
           COLLECT.Text+'" and G_NATUREGENE="COC"') then
           begin
               PgiInfo ('Le compte collectif n''existe pas dans le plan comptable', 'Envoyer');
               COLLECT.SetFocus ;
           end;
      end;
end;

procedure TFAssistCom.CARAUXChange(Sender: TObject);
begin
  inherited;
         Aux1.Text := CARAUX.value;
         Aux2.Text := CARAUX.value;
end;

procedure TFAssistCom.BDelete1Click(Sender: TObject);
var
Compte  : string;
iRef    : integer;
OkSup   : Boolean;
begin
  inherited;
  iRef := FGrille.Row;
  if FGrille.Cells[0, iRef] = '' then exit;
  Compte := AGauche(FGrille.Cells[0, iRef], 10,'0');
  if stArg <> '' then OkSup := TRUE
  else
  OkSup  := (PGIAsk ('Confirmez-vous la suppression du paramétrage sélectionné', 'Paramétrage Sisco II ')=mrYes);

  if not OkSup then  exit;

  FGrille.GotoLeBookmark(iRef);
  TOBCol.Detail[iRef-1].Free;
  if TobCol <> nil then
  begin
      FGrille.VidePile(FALSE);
      TobCol.PutGridDetail(THGRID,False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
  end;

end;

procedure TFAssistCom.AJOUTAUXClick(Sender: TObject);
var
ATob                  : Tob;
Collectif, Auxx1, Auxx2 : string;
iRef                  : integer;
begin
  inherited;
   if  TobCol = nil then TobCol := TOB.Create('', nil, -1);

   Collectif := AGauche(COLLECT.TEXT,10,'0');
   Auxx1 := AGauche(AUX1.TEXT,10,'0');
   Auxx2 := AGauche(AUX2.TEXT,10,'0');
   if (not Quadra.checked) then
   begin
         if (Copy (Collectif,1,2) = '40') and (Copy (Auxx1,1,1) <> '0')
         and (Copy (Auxx1,1,1) <> 'F') then
         begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;
         if (Copy (Collectif,1,2) = '41') and (Copy (Auxx1,1,1) <> '9')
         and (Copy (Auxx1,1,1) <> 'C') then
         begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;

         if (Copy (Collectif,1,2) = '40') and (Copy (Auxx2,1,1) <> '0')
         and (Copy (Auxx2,1,1) <> 'F') then
         begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;
         if (Copy (Collectif,1,2) = '41') and (Copy (Auxx2,1,1) <> '9')
         and (Copy (Auxx2,1,1) <> 'C') then
         begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;
   end;
   iRef := FGrille.Row;
   if Existedejacoll (THGRID, Collectif, iRef) then
   begin
    if (stArg <> '' ) or((stArg = '') and (PGIAsk('Voulez-vous enregistrer les modifications ?','Paramétrage Sisco II ')=mrYes)) then
    begin
         ATob := TOBCol.Detail[iRef-1];
         ATob.PutValue('CR_LIBELLE', AGauche(Auxx1, 10,'0'));
         ATob.PutValue('CR_ABREGE', AGauche(Auxx2, 10,'0'));
         TobCol.PutGridDetail(THGRID,False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
         exit;
    end
    else exit;
   end;
   if Existedejaaux (Collectif, Auxx1, Auxx2) then exit;
   Auxx1 := AUX1.TEXT;
   Auxx2 := AUX2.TEXT;
   ATob := TOB.Create('CORRESP',TobCol,-1);
   ATob.PutValue('CR_CORRESP', AGauche(Collectif, 10,'0'));
   ATob.PutValue('CR_LIBELLE', AGauche(Auxx1, 10,'0'));
   ATob.PutValue('CR_ABREGE', AGauche(Auxx2, 10,'0'));
   TobCol.PutGridDetail(THGRID,False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
end;

function TFAssistCom.Existedejacoll (Tg : THGRID; compte : string; var Rr : integer) : Boolean;
var
ii, indice   : integer;
cpte         : string;
begin
        indice := 0;
        for ii := 1 to Tg.RowCount - 1 do
        begin
          if Trim(Copy (Tg.Cells[0, ii],3,7)) <> '' then
           begin
             cpte := Copy (Tg.Cells[0, ii],0,10);
             Rr := ii;
             if compte = cpte then inc (indice);
             if indice >= 1 then break;
           end;
        end;
        if indice >= 1 then
        begin
             Result := TRUE; exit;
        end;
        Result := FALSE;
end;

function TFAssistCom.Existedejaaux (Col,T1,T2 : string) : Boolean;
var
ii                         : integer;
Err                        : Boolean;
Tranche1,Tranche2,cpte     : string;
begin
     Err := FALSE;
     for ii := 1 to FGrille.RowCount - 1 do
     begin
          cpte := Copy (FGrille.Cells[0, ii],0,10);
          if (copy (cpte,1,2)) = (copy (Col,1,2)) then
          begin
                Tranche1 := AGauche(FGrille.Cells[1, ii], 10,'0');
                Tranche2 := AGauche(FGrille.Cells[2, ii], 10,'0');
                if (Tranche1 < T1) then
                begin
                                 if (Tranche2 >= T2) then
                                 begin
                                      Err := TRUE; break;
                                 end;
                end
                else
                begin
                     if (Tranche1 <= T2) then
                     begin
                                  Err := TRUE; break;
                     end;

                end;
          end;
     end;
     if Err then
          PGiBox (' Les tranches de comptes sont déjà définies pour un autre collectif','Paramétrage Sisco II');
     Result := Err;
end;


procedure TFAssistCom.COLLECTElipsisClick(Sender: TObject);
begin
  inherited;
if (Copy (COLLECT.Text,1,2)) = '40' then
    LookupList(TControl(Sender),'','','G_GENERAL','G_LIBELLE','','',True,-1,'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX WHERE G_GENERAL LIKE "40%" and G_NATUREGENE="COF"');
if (Copy (COLLECT.Text,1,2)) = '41' then
    LookupList(TControl(Sender),'','','G_GENERAL','G_LIBELLE','','',True,-1,'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX WHERE G_GENERAL LIKE "41%" and G_NATUREGENE="COC"');
end;

function TFAssistCom.Creationparamcol : Boolean;
var
i         : integer;
Pcol      : ^ar_colreg;
Tcor      : TOB;
Cpte      : string;
begin
     Result := TRUE;
     ExecuteSQL('DELETE from CORRESP Where CR_TYPE="SIS"');
     for i := 1 to FGrille.RowCount - 1 do
     begin
          if Trim(Copy (FGrille.Cells[0, i],3,7)) <> '' then
          begin
            Tcor := TOB.Create('CORRESP', Nil, -1);
            system.New(Pcol);
            Cpte := FGrille.Cells[0, i];
            Pcol^.cpte := AGauche(Cpte,10,'0');
            Cpte := FGrille.Cells[1, i];
            Pcol^.inf := AGauche(Cpte,10,'0');
            Cpte := FGrille.Cells[2, i];
            Pcol^.sup := AGauche(Cpte,10,'Z');
            if Copy (Pcol^.cpte,1,2) = '40' then
               Pcol^.typ := 'F';
            if Copy (Pcol^.cpte,1,2) = '41' then
               Pcol^.typ := 'C';
            Tcor.PutValue('CR_TYPE', 'SIS');
            Tcor.PutValue('CR_CORRESP', Pcol^.cpte);
            Tcor.PutValue('CR_LIBELLE', Pcol^.inf);
            Tcor.PutValue('CR_ABREGE', Pcol^.sup);
            if ExisteSQl ('SELECT * from CORRESP Where CR_TYPE="SIS" and CR_CORRESP="'+ Pcol^.cpte +'"') then
            begin
                 Result := FALSE;
                 Pgiinfo ('Paramétrage incorrect.#10#13Collectif déjà existant : '+Pcol^.cpte,'Export Sisco II');
                 if Pcol <> nil then Dispose (Pcol);    exit;
            end;
            Tcor.InsertDB(nil, TRUE);
            if Pcol <> nil then Dispose (Pcol);
            Tcor.free;
          end;
     end;

end;

procedure TFAssistCom.FTimerTimer(Sender: TObject);
begin
  inherited;
  if not OkExport then exit;
  FTimer.Enabled := FALSE;
  bFinClick(Sender);
  Close ;
end;

procedure TFAssistCom.FormShow(Sender: TObject);
var
Q1                : TQuery;
ATob              : TOB;
Serie             : string;
Commande          : string;
Fichier           : string;
begin
  inherited;
  Label7.Caption := TraduireMemoire (Label7.Caption);
  TYPEFORMAT.ItemIndex := 1;
  NATURETRANSFERT.ItemIndex := 0;
  PSuiv := 1;  NoSeqNet := 0;
  lEtap.Caption := Msg.Mess[0] + ' ' + IntToStr(PSuiv);
  TraiteListeJournal ('');
  GRPBAL.Enabled  := FALSE;
  GRTYPECR.Enabled := FALSE;
  FTypeEcr.Enabled := FALSE;
  TFTypeEcr.Enabled := FALSE;
  Exclure.Enabled   := FALSE;
  TExclure.Enabled  := FALSE;

  PARAMSISCOII.Enabled := FALSE;
  OkEnter := FALSE;
  if LienGamme <> '' then
  TRANSFERTVERS.Itemindex := TRANSFERTVERS.Values.IndexOf(LienGamme)
  else
  TRANSFERTVERS.Itemindex := TRANSFERTVERS.Values.IndexOf(VH^.CPLienGamme);
  if TRANSFERTVERS.Itemindex = -1 then
  begin
      if VH^.CPLienGamme = 'S1' then  TRANSFERTVERS.Itemindex := 0
      else
      if VH^.CPLienGamme = 'S5' then  TRANSFERTVERS.Itemindex := 1
      else
      if VH^.CPLienGamme = 'S3' then  TRANSFERTVERS.Itemindex := 2
      else
      if VH^.CPLienGamme = 'SI' then  TRANSFERTVERS.Itemindex := 3
      else
      TRANSFERTVERS.Itemindex := 1;
  end;
  DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), now);
  if stArg = '' then  // par ligne de commande
  begin
       if TRANSFERTVERS.Itemindex = -1 then TRANSFERTVERS.Itemindex := 1;
       ChargeNatureTransfert(TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex]);
  end;
//  PPARAM.Enabled     :=  PARAMDOS.checked;
  PGEN.Enabled       :=  PARAMDOS.checked;
  PTIERS.Enabled     :=  PARAMDOS.checked;
  PTiersautre.Enabled:=  PARAMDOS.checked;
  PSectionana.Enabled:= PARAMDOS.checked;
  PJRL.Enabled       := PARAMDOS.checked;
//  TLibre.Enabled     := PARAMDOS.checked;

  PPARAMGENE.Enabled := not PARAMDOS.checked;

  CarFou := '0'; CarCli := '9';
  TobCol := TOB.Create('', nil, -1);
  Q1 := OpenSql ('SELECT * from CORRESP Where CR_TYPE="SIS"', TRUE);
  if Q1.EOF then
  begin
            COLLECTIFS.Value := 'F';
            CARAUX.value := '0';
  end
  else
  begin
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '40' then
                COLLECTIFS.Value := 'F';
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '41' then
                COLLECTIFS.Value := 'C';
           COLLECT.Text :=  AGauche(Q1.FindField ('CR_CORRESP').asstring, 10,'0');
           CARAUX.value := Copy (Q1.FindField ('CR_LIBELLE').asstring,1,1);
           AUX1.Text := AGauche(Q1.FindField ('CR_LIBELLE').asstring, 10,'0');
           AUX2.Text := AGauche(Q1.FindField ('CR_ABREGE').asstring, 10,'0');
  end;
  while not Q1.EOF do
  begin
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '40' then
                CarFou := copy (Q1.FindField ('CR_LIBELLE').asstring,0,1);
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '41' then
                CarCli := copy (Q1.FindField ('CR_LIBELLE').asstring,0,1);
           ATob := TOB.Create('CORRESP',TobCol,-1);
           ATob.PutValue('CR_CORRESP', AGauche(Q1.FindField ('CR_CORRESP').asstring, 10,'0'));
           ATob.PutValue('CR_LIBELLE', AGauche(Q1.FindField ('CR_LIBELLE').asstring, 10,'0'));
           ATob.PutValue('CR_ABREGE', AGauche(Q1.FindField ('CR_ABREGE').asstring, 10,'0'));
           Q1.next;
  end;
  ferme (Q1);
  FGrille := THGRID;
  TobCol.PutGridDetail(FGrille,False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
  COLLECT.EditMask := '';
  COLLECT.Text := '40';
  COLLECT.maxlength := VH^.Cpta[fbGene].Lg;
  AUX1.EditMask := '';
  AUX1.Text := '0';
  AUX1.maxlength := VH^.Cpta[fbGene].Lg;
  AUX2.EditMask := '';
  AUX2.Text := '0';
  AUX2.maxlength := VH^.Cpta[fbGene].Lg;
  TOBTiers := Nil ;
  FTimer.Enabled := FALSE;
  ExclureEnreg := '';
  OkExport := (stArg <> '');
  stArgsav := stArg;
  EtablissUnique := FALSE;
  if stArg <> '' then  // par ligne de commande
  begin
       if not RenseigneArg(Fichier, Commande) then exit;
  end
  else
  begin
       if V_PGI.LaSerie = S7 then serie := 'S5';
       if V_PGI.LaSerie = S5 then serie := 'S5';
       if V_PGI.LaSerie = S3 then serie := 'S3';
       if V_PGI.LaSerie = S1 then serie := 'S1';
       FCorpsMail.clear;
       FCorpsMail.Lines.add (TraduireMemoire('Veuillez trouver ci-joint le fichier correspondant au dossier '));
       FCorpsMail.Lines.add (V_PGI.NoDossier  + TraduireMemoire(' géré par le logiciel ')+ serie);
       FCorpsMail.Lines.add ('');
       FCorpsMail.Lines.add (TraduireMemoire('Merci de l''intégrer dès réception'));
       FEmail.Enabled := not(FFile.Checked);
       Label2.Enabled := not(FFile.Checked);
       FHigh.Enabled  := not(FFile.Checked);
       FCorpsMail.Enabled := not(FFile.Checked);
       Label12.Enabled := not(FFile.Checked);
       Label11.Enabled := not(FFile.Checked);
  end;
  VH^.CPLienGamme := GetParamSocSecur('SO_CPLIENGAMME', '', TRUE) ;
  DATARRETEBAL.Text := FormatDateTime(Traduitdateformat('dd/mm/yyyy'), VH^.Encours.Fin);
  ObjetFiltre.Charger;
  WhereAvancee := '';
end;

procedure TFAssistCom.EXERCICEChange(Sender: TObject);
begin
  inherited;
  if stArg <> '' then exit;

     DoExoToDateOnChange(EXERCICE, THEDIT(DATEECR1), THEDIT(DATEECR2)) ;
//GP     if EXERCICE.Text = '<<Tous>>' then
     if EXERCICE.Text = Traduirememoire('<<Tous>>') then
     begin
          DATEECR1.Text := stDate1900;
          DATEECR2.text := stDate2099;
     end;
end;

procedure TFAssistCom.DoExoToDateOnChange(Exo: THValComboBox; Date1, Date2: THEdit);
BEGIN
if (Exo=nil) or (Date1=nil) or (Date2=nil) then exit;
  ExoToDates(Exo.Value, Date1, Date2);
END;


procedure TFAssistCom.DATEECR2Exit(Sender: TObject);
Var D1,D2 : TDateTime ;
    Q     : TQuery;
    Okok  : boolean ;
    EXO   : String3;
begin
  inherited;
   EXO := EXERCICE.value;
//GP   if (EXERCICE.Text <> '<<Tous>>') and (EXERCICE.Text <> '') then
   if (EXERCICE.Text <> Traduirememoire('<<Tous>>')) and (EXERCICE.Text <> '') then
    begin
    Okok:=True ; D1:=Date ; D2:=Date ;
    If EXO=VH^.Precedent.Code Then BEGIN D1:=VH^.Precedent.Deb ; D2:=VH^.Precedent.Fin ; END Else
    If EXO=VH^.EnCours.Code Then BEGIN D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ; END Else
    If EXO=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END Else
    BEGIN
    Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE) ;
    if Not Q.EOF then
      BEGIN
      D1:=Q.FindField('EX_DATEDEBUT').asDateTime ; D2:=Q.FindField('EX_DATEFIN').asDateTime ;
      END else Okok:=False ;
      Ferme(Q) ;
    END;
    if Okok then
    begin
         if (D1 <= strTodate(DATEECR1.text)) and
         (D2  >= strTodate(DATEECR2.text)) then
         begin
              if (strTodate(DATEECR1.text)) <=
                 (strTodate(DATEECR2.text)) then exit
              else
                  DoExoToDateOnChange(EXERCICE, THEDIT(DATEECR1),THEDIT(DATEECR2)) ;
         end
         else
           DoExoToDateOnChange(EXERCICE, THEDIT(DATEECR1),THEDIT(DATEECR2)) ;
    end;
    end;

end;

procedure TFAssistCom.SANCOLLClick(Sender: TObject);
begin
  inherited;
  if ((TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex] = 'SISCO')) then
  begin
       if SANCOLL.checked then
       begin
        THGRID.VidePile(FALSE);
        AJOUTAUX.Enabled := FALSE;
        Bdelete1.Enabled := FALSE;
        COLLECTIFS.Enabled :=FALSE;
        COLLECT.Enabled :=FALSE;
        CARAUX.Enabled :=FALSE;
        AUX1.Enabled :=FALSE;
        AUX2.Enabled :=FALSE;
       end
       else
       begin
        if TobCol <> nil then
           TobCol.PutGridDetail(THGRID,False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
        AJOUTAUX.Enabled :=TRUE;
        Bdelete1.Enabled :=TRUE;
        COLLECTIFS.Enabled :=TRUE;
        COLLECT.Enabled :=TRUE;
        CARAUX.Enabled :=TRUE;
        AUX1.Enabled :=TRUE;
        AUX2.Enabled :=TRUE;
       end;

  end;

end;

procedure TFAssistCom.AUX1Click(Sender: TObject);
begin
  inherited;
         Aux1.Text := CARAUX.value;
end;

procedure TFAssistCom.AUX1Exit(Sender: TObject);
var
caraux1,car         : string;
begin
  inherited;
         caraux1 := CARAUX.value;
         car := Copy(AUX1.Text,1,1);
         if (caraux1 <> car) and (car <> '') and (not Quadra.checked) then
         begin
              AUX1.Text :=  caraux1+Copy(AUX1.Text,2,9);
              AUX2.Text :=  caraux1+Copy(AUX2.Text,2,9);
         end;
         AUX1.Text :=  AGauche(AUX1.Text, VH^.Cpta[fbGene].Lg,'0');

end;

procedure TFAssistCom.AUX2Click(Sender: TObject);
begin
  inherited;
         Aux2.Text := CARAUX.value;
end;

procedure TFAssistCom.AUX2Exit(Sender: TObject);
var
caraux1,car         : string;
begin
  inherited;
         caraux1 := CARAUX.value;
         car := Copy(AUX2.Text,1,1);
         if (caraux1 <> car) and (car <> '') and (not Quadra.checked) then
            AUX2.Text := caraux1+Copy(AUX2.Text,2,9);
         AUX2.Text :=  AGauche(AUX2.Text, VH^.Cpta[fbGene].Lg,'0');

end;

procedure TFAssistCom.FFileClick(Sender: TObject);
begin
  inherited;
       FEmail.Enabled := not(FFile.Checked);
       Label2.Enabled := not(FFile.Checked);
       FHigh.Enabled  := not(FFile.Checked);
       FCorpsMail.Enabled := not(FFile.Checked);
       Label12.Enabled := not(FFile.Checked);
       Label11.Enabled := not(FFile.Checked);
       if FFile.Checked then
             FEmail.Text := '';
       if FCorpsMail.Enabled then
            FCorpsMail.Font.Color := clBtnText
       else
            FCorpsMail.Font.Color := clGrayText;

end;

procedure TFAssistCom.FICHENAMEExit(Sender: TObject);
var
FileName  : string;
begin
  inherited;
                 if (Trans.FichierSortie = '') and (FicheName.text <> '') then
                   Trans.FichierSortie := FicheName.text;
                 FileName := Trans.FichierSortie;
                 FileName := ReadTokenPipe (Filename, '.');
                 if FileName = '' then
                 begin
                      PGIInfo ('Nom du fichier d''échange est obligatoire', 'Envoyer');
                 end;
end;

procedure TFAssistCom.bClickVisuClick(Sender: TObject);
var
Rep       : string;
begin
  inherited;
                     if FichierImp  = '' then
                     begin
                             FichierImp := 'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt';
                             EcrireDansfichierListeCom (FichierImp, LISTEEXPORT);
                     end;
                     Rep := ExtractFileDir (Trans.FichierSortie);
                     ShellExecute(0, PCHAR('open'), 'WordPad.exe' , PCHAR('"'+Rep + '\'+ FichierImp+'"' ),PCHAR(Rep + '\'+ FichierImp ),SW_RESTORE);
end;


procedure TFAssistCom.BNetExpertClick(Sender: TObject);
begin
  inherited;
     if  BNetExpert.Checked then
     begin
          FicheName.Text    := GetEnvVar('TEMP')+'\'+'PG'+ V_PGI.NoDossier+'_'+ Format ('%.03d', [NoSeqNet])+'.TRA';
          FicheName.Enabled := FALSE;
     end
     else
     begin
          FicheName.Text    := '';
          FicheName.Enabled := TRUE;
     end;
end;



procedure TFAssistCom.DATEECRChange(Sender: TObject);
var
Dfin,dd  : TDateTime;
jj       : integer;
st       : string;
Exerc    : TExoDate;
      procedure AffectDate;
      begin
                   dd := StrtoDate(DATEECR.Text);
                   Dfin := FinDeMois(dd);
                   DATEECR.Text := FormatDateTime(Traduitdateformat('dd/mm/yyyy'), Dfin) ;
                   DateFinMois := DATEECR.Text;
      end;
begin
  inherited;
  if not OkEnter then exit;
  if (stArg <> '') or (TRANSFERTVERS.Itemindex = -1) then exit;
  if (stArg = '') and (TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex] <> 'S1') and (NATURETRANSFERT.Value = 'SYN') then
  begin
       if (DATEECR.Text = '  /  /    ')  or (DATEECR.Text = stDate1900)
       or (pos(' ',DATEECR.Text) <> 0) then exit;
       if TransIsValidDate(DATEECR.Text) then
       begin
            AffectDate;
            // Fiche 10426
            if not CQuelExercice(StrToDate(DATEECR.Text), Exerc) then
             begin
                  PGIInfo('Date incorrecte');
                  if  BNetExpert.Checked then  DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), iDate1900)
                  else
                  begin
                         if (TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex] <> 'S1')  and (not(ctxPCL in V_PGI.PGIContexte)) then
                         begin
                            Dfin := PlusMois(V_PGI.DateEntree, -1);
                            Dfin := FinDeMois(Dfin);
                            DATEECR.Text := FormatDateTime(Traduitdateformat('dd/mm/yyyy'), DFin);
                         end
                         else
                            DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), now);
                  end;
                  exit;
             end;
       end
       else
       begin
        jj := Strtoint(Copy(DateEcr.Text, 0, 2));
        if jj > 31 then
        begin
             PGIInfo('Date incorrecte'); exit;
        end;
        // pour changement de mois
        if (DateFinMois <> '') and (Copy(DateEcr.Text, 4, 10) <> Copy(DateFinMois, 4, 10)) then
        begin
           if jj > 30 then
           begin
                st := '01/'+ Copy(DateEcr.Text, 4, 10);
                DateEcr.Text := St;
           end;
           AffectDate;
        end;
       end;
  end;
end;

procedure TFAssistCom.DATEECREnter(Sender: TObject);
begin
  inherited;
           OkEnter := TRUE;  DateFinMois := '';
end;

procedure TFAssistCom.DATEECRExit(Sender: TObject);
begin
  inherited;
           OkEnter := FALSE;
end;

procedure TFAssistCom.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
         if TobCol <> nil then
         begin
              TobCol.free;  TobCol := nil;
         end;
         if Assigned(ObjetFiltre) then FreeAndNil(ObjetFiltre);
end;

procedure TFAssistCom.PEXERCICEClick(Sender: TObject);
begin
  inherited;
    if (stArg = '') and not (PEXERCICE.Checked) then
    begin
         ExclureEnreg := ExclureEnreg + ';'+'EXO';
    end;
end;

procedure TFAssistCom.PPARAMSOCClick(Sender: TObject);
begin
  inherited;
    if (stArg = '') and not (PPARAMSOC.Checked) then
    begin
         ExclureEnreg := ExclureEnreg + ';'+'PARAM';
    end;
end;

procedure TFAssistCom.PECRBUDClick(Sender: TObject);
begin
  inherited;
    if PECRBUD.Checked then// écriture budgetaire
    begin
         if (NATURETRANSFERT.Value = 'BAL') then
            PARAMDOS.checked := TRUE
         else
            TraiteListeJournalBudgetaire;
    end
    else
        TraiteListeJournal (NATUREJRL.value);
end;


procedure TFAssistCom.ChargementCpte(ListeCpteGen : TList;Where : string);
var
Q1           :TQuery;
TCompteGene  : PListeCpteGen;
//i, j         : integer;
WW,WE,ORDERBY: string;
begin
// i:= gettickcount;
ORDERBY := ' ORDER BY G_GENERAL';
  if (NATURETRANSFERT.Value = 'JRL')
  or (NATURETRANSFERT.Value = 'SYN') then
  begin
   RenWhereEcr(WE,'E');
   WE :=' where g_general in (select distinct(e_general) from ecriture where '+ WE + ') ';
   WW :='SELECT G_GENERAL,G_LIBELLE,G_NATUREGENE,G_LETTRABLE,G_POINTABLE,G_VENTILABLE1,'+
   'G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_TABLE0,'+
   'G_TABLE1,G_TABLE2,G_TABLE3,G_TABLE4,G_TABLE5,G_TABLE6,G_TABLE7,G_TABLE8,G_TABLE9,'+
   'G_ABREGE,G_SENS,G_CORRESP1,G_CORRESP2,G_TVA,G_TVAENCAISSEMENT,G_TPF,'+
   'G_CUTOFF,G_CUTOFFPERIODE,G_CUTOFFECHUE,G_VISAREVISION,G_CYCLEREVISION,G_CUTOFFCOMPTE, G_CONFIDENTIEL'+
   ' FROM GENERAUX ' + WE + ORDERBY
  end
  else
   WW :='SELECT G_GENERAL,G_LIBELLE,G_NATUREGENE,G_LETTRABLE,G_POINTABLE,G_VENTILABLE1,'+
   'G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_TABLE0,'+
   'G_TABLE1,G_TABLE2,G_TABLE3,G_TABLE4,G_TABLE5,G_TABLE6,G_TABLE7,G_TABLE8,G_TABLE9,'
   +'G_ABREGE,G_SENS,G_CORRESP1,G_CORRESP2,G_TVA,G_TVAENCAISSEMENT,G_TPF, '+
   'G_CUTOFF,G_CUTOFFPERIODE,G_CUTOFFECHUE,G_VISAREVISION,G_CYCLEREVISION,G_CUTOFFCOMPTE, G_CONFIDENTIEL'+
   ' FROM GENERAUX '+ Where + ORDERBY;

 Q1:=OpenSQL(WW,TRUE) ;
 While Not Q1.Eof Do BEGIN
      System.New (TCompteGene);
      TCompteGene^.code        := Q1.FindField ('G_GENERAL').asstring;
      TCompteGene^.libelle     := Q1.FindField ('G_LIBELLE').asstring;
      TCompteGene^.nature      := Q1.FindField ('G_NATUREGENE').asstring;
      TCompteGene^.Lettrable   := Q1.FindField ('G_LETTRABLE').asstring;
      TCompteGene^.Pointage    := Q1.FindField ('G_POINTABLE').asstring;
      TCompteGene^.Ventilaxe1  := Q1.FindField ('G_VENTILABLE1').asstring;
      TCompteGene^.Ventilaxe2  := Q1.FindField ('G_VENTILABLE2').asstring;
      TCompteGene^.Ventilaxe3  := Q1.FindField ('G_VENTILABLE3').asstring;
      TCompteGene^.Ventilaxe4  := Q1.FindField ('G_VENTILABLE4').asstring;
      TCompteGene^.Ventilaxe5  := Q1.FindField ('G_VENTILABLE5').asstring;
      TCompteGene^.Table1      := Q1.FindField ('G_TABLE0').asstring;
      TCompteGene^.Table2      := Q1.FindField ('G_TABLE1').asstring;
      TCompteGene^.Table3      := Q1.FindField ('G_TABLE2').asstring;
      TCompteGene^.Table4      := Q1.FindField ('G_TABLE3').asstring;
      TCompteGene^.Table5      := Q1.FindField ('G_TABLE4').asstring;
      TCompteGene^.Table6      := Q1.FindField ('G_TABLE5').asstring;
      TCompteGene^.Table7      := Q1.FindField ('G_TABLE6').asstring;
      TCompteGene^.Table8      := Q1.FindField ('G_TABLE7').asstring;
      TCompteGene^.Table9      := Q1.FindField ('G_TABLE8').asstring;
      TCompteGene^.Table10     := Q1.FindField ('G_TABLE9').asstring;
      TCompteGene^.abrege      := Q1.FindField ('G_ABREGE').asstring;
      TCompteGene^.sens        := Q1.FindField ('G_SENS').asstring;
      // ajout me 13-05-2002
      TCompteGene^.Corresp1    := Q1.FindField ('G_CORRESP1').asstring;
      TCompteGene^.Corresp2    := Q1.FindField ('G_CORRESP2').asstring;
      // ajout me 12-01-2004
      TCompteGene^.Tva         := Q1.FindField ('G_TVA').asstring;
      TCompteGene^.Encaissement:= Q1.FindField ('G_TVAENCAISSEMENT').asstring;
      TCompteGene^.TPF         := Q1.FindField ('G_TPF').asstring;
// ajout me 24-05-2005 pour les cutoff
      TCompteGene^.ctoff         := Q1.FindField ('G_CUTOFF').asstring;
      TCompteGene^.ctoffper      := Q1.FindField ('G_CUTOFFPERIODE').asstring;
      TCompteGene^.cteche        := Q1.FindField ('G_CUTOFFECHUE').asstring;
      TCompteGene^.visarev       := Q1.FindField ('G_VISAREVISION').asstring;
      TCompteGene^.cyclerev      := Q1.FindField ('G_CYCLEREVISION').asstring;
      TCompteGene^.ctcpte        := Q1.FindField ('G_CUTOFFCOMPTE').asstring;
      TCompteGene^.condifentiel  := Q1.FindField ('G_CONFIDENTIEL').asstring;
      ListeCpteGen.Add(TCompteGene) ;
    Q1.NExt ;
END ;
Ferme(Q1) ;
// j:= gettickcount;
// showmessage(inttostr(j-i));
end;


procedure  TFAssistCom.RenWhereEcr( var Where : string; Ext : string);
var
Jourx                   : string;
EQUALIFPIECE            : string;
begin
          if Trans.Exo <> '' then
             Where := RendCommandeExo (Ext,Trans.Exo);
          if Trans.Jr <> ''  then
          begin
               Jourx := Trans.Jr;
               if Where <> '' then  Where := Where + ' AND ';
               if Ext = 'BE' then
                  Where := Where + '('+ RendCommandeJournalBud(Ext,Jourx)+') '
               else
                   Where := Where + '('+ RendCommandeJournal(Ext,Jourx)+') ';
          end;
         if (Trans.Typearchive = 'JRL') then
         begin
                if (pos ('N', FTypeEcr.Text) <> 0) then
                   EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="N"  ';
                if (pos ('S', FTypeEcr.Text) <> 0) then // écriture SIMULATION
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="S"  '
                   else
                      EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="S" ';
                end;

                if (pos ('I', FTypeEcr.Text) <> 0) then   // écriture IFRS
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="I"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="I" ';
                end;

                if (pos ('P', FTypeEcr.Text) <> 0) then   // écriture prévision
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="P"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="P" ';
                end;

                if (pos ('R', FTypeEcr.Text) <> 0) then   // écriture révision
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="R"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="R" ';
                end;

                //if BSITUATION.checked  then
                if (pos ('U', FTypeEcr.Text) <> 0) then  // écriture Situation
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="U"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="U" ';
                end;

                if (stArg <> '') and (TYPEECRX = 'X') then
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="X"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="X" ';
                end;
                if EQUALIFPIECE <> '' then
                begin
                  EQUALIFPIECE := EQUALIFPIECE + ') ';
                  if Where <> '' then  Where := Where + ' AND ';
                     Where := Where + EQUALIFPIECE;
                end;

                if not(BExport.State=cbGrayed) then
                  if BExport.State=cbChecked then Where := Where + ' AND '+Ext+'_EXPORTE="X"' else Where := Where +' AND '+Ext+'_EXPORTE<>"X"' ;

                if WhereAvancee <> '' then  Where := Where + WhereAvancee;

          end
          else
          begin
               if (Trans.Typearchive = 'DOS') then
               begin
                  if (Trans.Serie = 'S1') then
                     EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="N"  '
                     + 'or '+Ext+'_QUALIFPIECE="P" '
                     + 'or '+Ext+'_QUALIFPIECE="R" '
                     + 'or '+Ext+'_QUALIFPIECE="C" '
                     + ' ) '
                   else
                     EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="N"  '
                     + 'or '+Ext+'_QUALIFPIECE="S" '
                     + 'or '+Ext+'_QUALIFPIECE="P" '
                     + 'or '+Ext+'_QUALIFPIECE="R" '
                     + 'or '+Ext+'_QUALIFPIECE="C" '
                     + 'or '+Ext+'_QUALIFPIECE="U" ) ';
                   if Where <> '' then  Where := Where + ' AND ';
                     Where := Where + EQUALIFPIECE;
               end;
          end;
          if Trans.Etabi <> '' then
          begin
               if Where <> '' then  Where := Where + ' AND ';
                Where := Where + RendCommandeComboMulti(Ext,Trans.Etabi,'ETABLISSEMENT');
//               Where := Where + '('+Ext+'_ETABLISSEMENT="'+ Trans.Etabi+'") ';
          end;

          if (Trans.Dateecr1 <> iDate1900) and (Trans.Dateecr2 <> iDate1900) then
          begin
               if Where <> '' then  Where := Where + ' AND';
               Where := Where + ' ('+Ext+'_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'"' +
                              ' AND '+Ext+'_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")'
          end
          else
          begin
              if (Trans.Dateecr1 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' ('+Ext+'_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'")';
              end else
              if (Trans.Dateecr2 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' ('+Ext+'_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")';
              end;
          end ;
      if (Trans.Typearchive = 'SYN') and (Ext = 'E') then
      begin
             if (HISTOSYNCHRO.Value = '0') then
                Where :=  Where + ' AND E_QUALIFPIECE="N" AND E_IO="0"'
             else
             if (HISTOSYNCHRO.Value = '1') then
                Where :=  Where + ' AND E_QUALIFPIECE="N" AND E_IO="1"'
             else
                Where :=  Where + ' AND E_QUALIFPIECE="N" AND E_IO="X"'
      end;

end;

Function TFAssistCom.ControleSisco : Boolean;
var
Q1               : TQuery;
OkTiers          : Boolean;
OkExport         : Boolean;
ii               : integer;
cpte,Rep         : string;
F                : TextFile;
procedure EcritureErreur;
begin
                   Rep := GetParamSocSecur('SO_CPRDREPERTOIRE', '');
                   if Rep = '' then
                      Rep := ExtractFileDir (Trans.FichierSortie);
                   if (FichierImp = '') then
                      FichierImp := 'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt';
                   EcrireDansfichierListeCom (Rep+'\'+FichierImp, LISTEEXPORT);
end;
begin
            for ii := 1 to THGRID.RowCount - 1 do
            begin
              if Trim(Copy (THGRID.Cells[0, ii],3,7)) <> '' then
               begin
                 cpte := Copy (THGRID.Cells[0, ii],0,10);
                 if copy (cpte,1,2) = '40' then
                    CarFou := Copy (THGRID.Cells[1, ii],0,1);
                 if copy (cpte,1,2) = '41' then
                    CarCli := Copy (THGRID.Cells[1, ii],0,1);
               end;
            end;

             Result := TRUE;
             OkTiers := TRUE; OkExport := TRUE;
             if VH^.Cpta[fbGene].Lg <> VH^.Cpta[fbAux].Lg then
             begin
              AfficheListeCom('Les comptes généraux et auxiliaires sont de longueurs différentes, export impossible', LISTEEXPORT);
              AfficheListeCom('Longueur des comptes généraux : '+ IntToStr(VH^.Cpta[fbGene].Lg), LISTEEXPORT);
              AfficheListeCom('Longueur des comptes Tiers    : '+ IntToStr(VH^.Cpta[fbAux].Lg), LISTEEXPORT);
              OkExport := FALSE;
             end;
             if VH^.Cpta[fbGene].Lg > 10 then
             begin
              AfficheListeCom('Longueur > 10, export sera impossible', LISTEEXPORT);
              OkExport := FALSE;
             end;
             if CompteEstalpha (ListeCpteGen) then
             begin
              AfficheListeCom('Des comptes généraux alpha numériques sont présents, export impossible', LISTEEXPORT);
              OkExport := FALSE;
             end;
             Q1 := Opensql ('select count(t_auxiliaire) from tiers where t_natureauxi="CLI" and '+
                '(t_auxiliaire not like "' +CarCli+ '%")', TRUE);
             if Q1.Fields[0].asinteger > 0 then
             begin
              AfficheListeCom(IntToStr(Q1.Fields[0].asinteger)+' comptes tiers clients ne commençant pas par : '+ CarCli, LISTEEXPORT);
              OkTiers := FALSE;
             end;
             ferme (Q1);

             Q1 := Opensql ('select count(t_auxiliaire) from tiers where t_natureauxi="FOU" and '+
                '(t_auxiliaire not like "' +CarFou+ '%")', TRUE);
             if Q1.Fields[0].asinteger > 0 then
             begin
              AfficheListeCom(IntToStr(Q1.Fields[0].asinteger)+' comptes tiers fournisseurs ne commençant pas par :'+ CarFou, LISTEEXPORT);
              OkTiers := FALSE;
             end;
             ferme (Q1);

             if (not OkTiers) and OkExport then
             begin
              AfficheListeCom('Le caractère '+CarCli + ' ou '+ CarFou +' sera ajouté avant tous les comptes Tiers', LISTEEXPORT);
              AfficheListeCom('La longueur du fichier passe de '+ IntToStr(VH^.Cpta[fbGene].Lg) + ' à '+ IntToStr(VH^.Cpta[fbGene].Lg+1), LISTEEXPORT);
              if (stArg = '') then
                 PGIInfo ('Le caractère '+CarCli + ' ou '+ CarFou +' sera ajouté avant tous les comptes Tiers'+#10#13
                 + ' La longueur du fichier passe de '+ IntToStr(VH^.Cpta[fbGene].Lg) + ' à '+ IntToStr(VH^.Cpta[fbGene].Lg+1));
             end;
             if (not OkTiers)  then
             begin
                     bClickVisu.visible := TRUE;
                     EcritureErreur;
{$I-}
                     AssignFile(F, Rep+'\'+FichierImp);
                     Append(F) ;
                     writeln(F, 'Liste des comptes Tiers clients : ');
                     Q1 := Opensql ('select t_auxiliaire from tiers where t_natureauxi="CLI" and '+
                        '(t_auxiliaire not like "' +CarCli+ '%")', TRUE);
                     while not Q1.EOF do
                     begin
                      writeln(F, '         '+ Q1.FindField ('T_AUXILIAIRE').asstring);
                      Q1.next;
                     end;
                     ferme (Q1);
                     writeln(F, 'Liste des comptes Tiers fournisseurs : ');
                     Q1 := Opensql ('select t_auxiliaire from tiers where t_natureauxi="FOU" and '+
                        '(t_auxiliaire not like "' +CarFou+ '%")', TRUE);
                     While not Q1.EOF do
                     begin
                      writeln(F, '         '+ Q1.FindField ('T_AUXILIAIRE').asstring);
                      Q1.next;
                     end;
                     ferme (Q1);
                     CloseFile(F);
{$I+}
             end;
             if (not OkExport) then
             begin
               Result := FALSE;
               if (OkTiers) then EcritureErreur;
             end;
end;

{$IFDEF EAGLCLIENT}
procedure TFAssistCom.DownloadEnFile(Data, FichierDeSortie : string);
var
TDataStream             : TMemoryStream;
begin
          TDataStream := TMemoryStream.Create;
          TDataStream.Write(PChar(Data)^, Length(Data));
          TDataStream.Seek(0, 0);
          TDataStream.SaveToFile(FichierDeSortie);
          TDataStream.free;
end;
{$ENDIF}

Procedure TFAssistCom.ExportComSx;
var
Rep,Stt                 : string;
FileName                : string;
LT                      : TOB;
Exerc                   : TExoDate;
{$IFDEF EAGLCLIENT}
lTobResult              : TOB;
Filearchive             : string;
{$ENDIF}
OkExport                : Boolean;
{$IFNDEF EAGLCLIENT}
PEExport                : TExport;
{$ENDIF}
begin
    if WindowState  = wsMinimized then
    begin
       Stt := 'Execution Export Fichier : ' + Trans.FichierSortie;
       InitProgressbar(Stt);
    end;
    if Trans.FichierSortie='' then
      begin
        PgiBox('Veuillez choisir un fichier d''export #10'+Trans.FichierSortie,'Envoyer') ;
        Exit ;
      end ;
      FileName := Trans.FichierSortie;
      FileName := ReadTokenPipe (Filename, '.');
      if FileName = '' then
      begin
        PgiBox('Veuillez choisir un fichier d''export #10'+Trans.FichierSortie,'Envoyer') ;
        Exit ;
      end;

    // Fiche 10225 if (stArg = '') and (Not _BlocageMonoPoste(True)) then Exit ;
    Rep := GetParamSocSecur('SO_CPRDREPERTOIRE', '');
    if Rep = '' then
    Rep := ExtractFileDir (Trans.FichierSortie);

    // initialisation des informations en vu d'export
    LT := InitTobParam;
    PEXPORT.VISIBLE    := true;
    P.ACTIVEPAGE       := PEXPORT;
    bClickVisu.visible := TRUE;

    {$IFDEF EAGLCLIENT}
      with cWA.create do
      begin
           OnCallBack := ComsxOnCallBack ;
          lTobResult  := Request('COMSX.EXPORT','', LT,'','');
         free ;
      end ;
      if lTobResult = nil then
      begin
        PgiBox('Erreur, Server Comsx introuvable','Envoyer') ;
        OkExport := FALSE;
      end
      else
      OkExport := (lTobResult.GetValue ('ERROR') <> '');
    {$ELSE}
      // BVE 28.08.07 : Rajout de l'option d'archivage.
      OkExport := LanceExport (LT, stArg, AfficheListeComExport, PEExport, Archivage);
    {$ENDIF}
     if not OkExport Then
        MessageAlerte('Problème d''export')
      else
      begin
    {$IFNDEF EAGLCLIENT}
           // ajout  Envoi Net Expert
           if (BNetExpert.Checked) and ((Trans.Typearchive = 'SYN') or (Trans.Typearchive = 'DOS') or (Trans.Typearchive = 'DOSS')) then
           begin
              NetExpertEnvoi;
           end;
    {$ELSE}
        if BGed.checked then
        begin
          FileName := Trans.FichierSortie;
          Filearchive := ReadTokenPipe (Filename, '.');
          Filearchive := Filearchive + '.ZIP';
          FileName := './ECOMSX/'+ AppServer.SessionID ;
          AppServer.RequestFileTo (Filearchive, FileName+'/'+ExtractFileName(Filearchive));
          with cWA.create do
          begin
               Request('COMSX.SUPPDIR','', LT,'','');
               free;
          end;
        end;

        if lTobResult.FieldExists('DATA') then
        begin
               if not Zippe.Checked then
                  DownloadEnFile(lTobResult.getValue ('DATA'), Trans.FichierSortie);
        end;
    {$ENDIF}

           // ajout de clôture périodique
           if (Trans.Serie  <> 'S1') and (Trans.Typearchive = 'SYN') and (StrToDate(DATEECR.text) <> iDate1900)
           and CQuelExercice(StrToDate(DATEECR.Text), Exerc)  then // fiche 10426
              CloPerSynchro (DATEECR.Text);
      end;
      LT.free;

      if FichierImp = '' then
          FichierImp := Rep+'\'+'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt'
      else
          FichierImp := Rep+'\'+ FichierImp;
      EcrireDansfichierListeCom (FichierImp, LISTEEXPORT);
      if stArg = '' then
      begin
        if PGIAsk('Export terminé : voulez-vous imprimer le rapport ?','Envoyer')=mrYes then
        begin
          if PrinterSetupDialog1.Execute then
             ControlTextToPrinter(FichierImp,poPortrait);
        end;
      end;

    FTimer.free;
    ModalResult := 1;
    if WindowState  = wsMinimized then
       FiniProgressbar;
      // Fiche 10225 if (stArg = '') then _DeblocageMonoPoste(True) ;
end;


Procedure TFAssistCom.ExportComSage;
var
St                       : string;
Rep                      : string;
FileName                 : string;
LT                       : TOB;
{$IFDEF EAGLCLIENT}
lTobResult               : TOB;
Filearchive              : string;
{$ELSE}
ExportCom                 : TExport;
{$ENDIF}
begin
    if Trans.FichierSortie='' then
      begin
        PgiBox('Veuillez choisir un fichier d''export #10'+Trans.FichierSortie,'Envoyer') ;
        Exit ;
      end ;
      FileName := Trans.FichierSortie;
      FileName := ReadTokenPipe (Filename, '.');
      if FileName = '' then
      begin
        PgiBox('Veuillez choisir un fichier d''export #10'+Trans.FichierSortie,'Envoyer') ;
        Exit ;
      end;
      LT := InitTobParam;
      PEXPORT.VISIBLE    := true;
      P.ACTIVEPAGE       := PEXPORT;
      bClickVisu.visible := TRUE;
      if WindowState  = wsMinimized then
      begin
         St := 'Execution Export Fichier : ' + Trans.FichierSortie;
         InitProgressbar(St);
      end;

{$IFDEF EAGLCLIENT}
//      lTobResult := LanceProcessServer('CGIComsx', 'EXPORT', 'AUCUN', LT, True ) ;
//      if lTobResult.GetValue('RESULT') = '' then
      with cWA.create do
      begin
           OnCallBack := ComsxOnCallBack ;
          lTobResult  := Request('COMSX.EXPORT','', LT,'','');
         free ;
      end ;
      if lTobResult = nil then
      begin
        PgiBox('Erreur, Server Comsx introuvable','Envoyer') ;
        OkExport := FALSE;
      end
      else
      OkExport := (lTobResult.GetValue ('ERROR') <> '');

{$ELSE}
     // BVE 28.08.07 : Rajout de l'option d'archivage.
     OkExport := LanceExport (LT, stArg, AfficheListeComExport, ExportCom, Archivage);
{$ENDIF}
     if not OkExport Then
        MessageAlerte('Problème d''export');
{$IFDEF EAGLCLIENT}
        if BGed.checked then
        begin
          FileName := Trans.FichierSortie;
          Filearchive := ReadTokenPipe (Filename, '.');
          Filearchive := Filearchive + '.ZIP';
          FileName := './ECOMSX/'+ AppServer.SessionID ;
          AppServer.RequestFileTo (Filearchive, FileName+'/'+ExtractFileName(Filearchive));
          with cWA.create do
          begin
               Request('COMSX.SUPPDIR','', LT,'','');
               free;
          end;
        end;
        if lTobResult.FieldExists('DATA') then
        begin
               if not Zippe.Checked then
                  DownloadEnFile(lTobResult.getValue ('DATA'), Trans.FichierSortie);
        end;
{$ENDIF}
      LT.free;

    AfficheListeCom('Export terminé ', LISTEEXPORT);

    Rep := GetParamSocSecur('SO_CPRDREPERTOIRE', '');
    if Rep = '' then
    Rep := ExtractFileDir (Trans.FichierSortie);
    if FichierImp = '' then
      FichierImp := Rep+'\'+'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt'
    else
      FichierImp := Rep+'\'+ FichierImp;
    EcrireDansfichierListeCom (FichierImp, LISTEEXPORT);
    if stArg = '' then
    begin
      if PGIAsk('Export terminé : voulez-vous imprimer le rapport ?','Envoyer')=mrYes then
      begin
        if PrinterSetupDialog1.Execute then
           ControlTextToPrinter(FichierImp,poPortrait);
      end;
    end;
    FTimer.free;
    if WindowState  = wsMinimized then
       FiniProgressbar;
    ModalResult := 1;
end;



procedure TFAssistCom.ChargeNatureTransfert(TValue : string);
var
ModeS : boolean;
begin
  if (TValue = 'S1') then
  begin
    NATURETRANSFERT.Items.clear;
    NATURETRANSFERT.Values.clear;
    ModeS := GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE, FALSE);
    if(ModeS = TRUE) or (GetparamsocSecur ('SO_FLAGSYNCHRO', '') = '') then
    begin
         NATURETRANSFERT.Items.Add(TraduireMemoire('Synchronisation'));
         NATURETRANSFERT.Values.add('SYN');
    end;
    if ((NoSeqNet = 1) and (BNetExpert.Checked)) or (not BNetExpert.Checked) then
    begin
        NATURETRANSFERT.Items.Add(TraduireMemoire('Dossier Complet'));
        NATURETRANSFERT.Values.add('DOS');
        NATURETRANSFERT.Value := 'DOS';
        if (not BNetExpert.Checked) and (ModeS = FALSE) then
        begin
             NATURETRANSFERT.Items.Add(TraduireMemoire('Journal'));
             NATURETRANSFERT.Values.add('JRL');
        end;
        if (ModeS = TRUE) and ((GetParamSocSecur ('SO_CPDATESYNCHRO1', '') > iDate1900) or
               (GetParamSocSecur ('SO_CPDATESYNCHRO2', '') > iDate1900) or
               (GetparamsocSecur ('SO_FLAGSYNCHRO', '') = 'SYN')) then
                                 NATURETRANSFERT.ItemIndex := 0
        else
                                 NATURETRANSFERT.ItemIndex := 1;
        if (GetparamsocSecur ('SO_FLAGSYNCHRO', '') = 'JRL') then RendVisibleChamp;

    end
    else
    if not BNetExpert.Checked then
    begin
         NATURETRANSFERT.Items.Add(TraduireMemoire('Journal'));
         NATURETRANSFERT.Values.add('JRL');
         NATURETRANSFERT.ItemIndex := 0;
    end
    else
         NATURETRANSFERT.ItemIndex := 0;
    if (BNetExpert.Checked) and (NoSeqNet > 1) then
       NATURETRANSFERT.Value := 'SYN';
  end
  else
  begin
      if (VH^.CPLienGamme = 'S1') and (TValue <> 'SISCO') then
      begin
              NATURETRANSFERT.Items.clear;
              NATURETRANSFERT.Values.clear;
              NATURETRANSFERT.Items.Add(TraduireMemoire('Dossier Complet'));
              NATURETRANSFERT.Values.add('DOS');
              NATURETRANSFERT.Items.Add(TraduireMemoire('Balance'));
              NATURETRANSFERT.Values.add('BAL');
              if V_PGI.SAV then
              begin
                   NATURETRANSFERT.Items.Add(TraduireMemoire('Journal'));
                   NATURETRANSFERT.Values.add('JRL');
              end;
              NATURETRANSFERT.ItemIndex := 0;
      end
      else
      if (TValue = 'SISCO') then
      begin
        NATURETRANSFERT.Items.clear;
        NATURETRANSFERT.Values.clear;
        NATURETRANSFERT.Items.Add(TraduireMemoire('Balance'));
        NATURETRANSFERT.Values.add('BAL');
        NATURETRANSFERT.Items.Add(TraduireMemoire('Journal'));
        NATURETRANSFERT.Values.add('JRL');
        NATURETRANSFERT.Items.Add(TraduireMemoire('Exercice'));
        NATURETRANSFERT.Values.add('EXE');

        NATURETRANSFERT.ItemIndex := 1;
        NATURETRANSFERT.Value := 'JRL';
        FICHENAME.DataType := 'SAVEFILE(SI*.TRT)';
      end
      else
      begin
        NATURETRANSFERT.Items.clear;
        NATURETRANSFERT.Values.clear;
        if ((NoSeqNet = 1) and (BNetExpert.Checked)) or (not BNetExpert.Checked) then
        begin
              NATURETRANSFERT.Items.Add(TraduireMemoire('Dossier Complet'));
              NATURETRANSFERT.Values.add('DOS');
              NATURETRANSFERT.Items.Add(TraduireMemoire('Dossier Complet en vue d''une première synchronisation'));
              NATURETRANSFERT.Values.add('DOSS');
              NATURETRANSFERT.Items.Add(TraduireMemoire('Balance'));
              NATURETRANSFERT.Values.add('BAL');
        end;
        if not BNetExpert.Checked then
        begin
             NATURETRANSFERT.Items.Add(TraduireMemoire('Journal'));
             NATURETRANSFERT.Values.add('JRL');
        end;
        NATURETRANSFERT.Items.Add(TraduireMemoire('Synchronisation'));
        NATURETRANSFERT.Values.add('SYN');
        if (Trans.Serie <> 'SAGE') and (Trans.Serie <> 'IMMO') and (Trans.Serie <> 'COMPTA') then
        begin
             if not BNetExpert.Checked then
             begin
                    NATURETRANSFERT.ItemIndex := 0;
                    NATURETRANSFERT.Value := 'DOS';

             end
             else
             begin
                  if (NoSeqNet = 1) then
                    NATURETRANSFERT.ItemIndex := 1
                  else
                    NATURETRANSFERT.ItemIndex := 0;
             end;
        end;
      end;
  end;
end;

procedure TFAssistCom.TRANSFERTVERSClick(Sender: TObject);
var
TValue   : string;
begin
  inherited;
  FICHENAME.DataType := 'SAVEFILE(*.TRA)';
  TValue := TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex];
  if (TValue <>  VH^.CPLienGamme) then
         BNetExpert.Checked := FALSE
  else
      // CA - 28/10/2005 - Pour corriger problème affichage message NATURE en activation ASP
      if ((TValue = 'S1') and (stArg='')) then
         BNetExpert.Checked := IsDossierNetExpert (V_PGI.NoDossier, NoSEqNet);
  if (LienGamme <> '') then
  begin
   if (LienGamme = 'S1') and (TValue <> 'S1') then
   begin
     TValue := 'S1'; TRANSFERTVERS.Value := 'S1';
   end
   else
   if (LienGamme = 'S5') and (TValue = 'S1') then
   begin
     TValue := 'S5';
     TRANSFERTVERS.Value := 'S5';
   end;
  end;
  ChargeNatureTransfert(TValue);
  RendVisibleChamp;
end;

procedure TFAssistCom.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
    VK_F10    : bFinClick(Sender);
  end;

end;

function TFAssistCom.InitTobParam : TOB ;
var
LT, L1   : TOB;
WhereNature : string;
begin
    LT := TOB.Create('$PARAM', nil, -1) ;
// GP sur demande J POMAT le 5/5/2008 : (mise en commentaire)
(*
    LT.AddChampSupValeur('USERLOGIN' , V_PGI.UserLogin ) ;
    LT.AddChampSupValeur('INIFILE'   , HalSocIni ) ;
    LT.AddChampSupValeur('PASSWORD'  , V_PGI.Password ) ;
    LT.AddChampSupValeur('DOMAINNAME', '' ) ;
    LT.AddChampSupValeur('DATEENTREE', V_PGI.DateEntree ) ;
*)
// GP Réactivation pour les DLL du positionnement de 'DOSSIER', sinon la dll qui s'en sert plante grave (#0 pas cool....)
    LT.AddChampSupValeur('DOSSIER'   , V_PGI.NoDossier(*V_PGI.CurrentAlias*) ) ;

    //    LT.AddChampSupValeur('DOSSIER'   , V_PGI.NoDossier(*V_PGI.CurrentAlias*) ) ;
    LT.AddChampSupValeur('APPLICATION', NameAppli) ;

    LT.AddChampSupValeur('BaseCommune', EstBaseCommune);

    L1 := TOB.Create('Trans', LT, -1) ;
    L1.AddChampSupValeur('Balance'      , Trans.balance) ;
    L1.AddChampSupValeur('Complement'   , Trans.complement ) ;
    L1.AddChampSupValeur('DateEcr1'     , Trans.Dateecr1 ) ;
    L1.AddChampSupValeur('DateEcr2'     , Trans.Dateecr2 ) ;
    L1.AddChampSupValeur('Etabli'       , Trans.Etabi ) ;
    L1.AddChampSupValeur('Exo'          , Trans.Exo ) ;
    Trans.Exporte          := (BExport.Checked);
    L1.AddChampSupValeur('Exporte'      , Trans.Exporte ) ;
    L1.AddChampSupValeur('FichierSortie', Trans.FichierSortie ) ;
    L1.AddChampSupValeur('Jr'           , Trans.Jr) ;
    L1.AddChampSupValeur('naturejal'    , Trans.naturejal) ;
    L1.AddChampSupValeur('pana'         , Trans.pana) ;
    L1.AddChampSupValeur('Pgene'        , Trans.Pgene) ;
    L1.AddChampSupValeur('pjournaux'    , Trans.pjournaux) ;
    L1.AddChampSupValeur('psection'     , Trans.psection) ;
    L1.AddChampSupValeur('ptiers'       , Trans.ptiers) ;
    L1.AddChampSupValeur('ptiersautre'  , Trans.ptiersautre) ;
    L1.AddChampSupValeur('Serie'        , Trans.Serie) ;
    L1.AddChampSupValeur('Synch'        , Trans.Synch) ;
    if (NATURETRANSFERT.Value = 'DOSS') then Trans.Typearchive := 'DOSS';
    L1.AddChampSupValeur('Typearchive'  , Trans.Typearchive) ;
    L1.AddChampSupValeur('TypeFormat'   , Trans.TypeFormat) ;
    L1.AddChampSupValeur('PDos'         , (Paramdos.Checked)) ;
    L1.AddChampSupValeur('Aveclettrage' , (Aveclettrage.Checked));
    // fiche 10566  et fiche 10567
    if ((Trans.Typearchive = 'JRL') or (Trans.Typearchive = 'BAL')) then
    begin
        if (Pos('<<', Exclure.TEXT) >0 ) then ExclureEnreg := ExclureEnreg+'BQC;BQE;RIB;SAT;SSA;ETB;MDP;MDR;DEV;REG;SOU;CORR;TL;REL;CGN;CAE;JAL' // fiche 22010
        else ExclureEnreg := ExclureEnreg+Exclure.TEXT;
    end
    else
    if (Pos('<<', Exclure.TEXT) >0 ) then
          ExclureEnreg := ''
    else
          ExclureEnreg := ExclureEnreg+Exclure.TEXT;

    L1.AddChampSupValeur('ExclureEnreg' , ExclureEnreg);
    L1.AddChampSupValeur('FichierImp'   , FichierImp);
    L1.AddChampSupValeur('PECRBUD'      , (PECRBUD.Checked));
    L1.AddChampSupValeur('BANA'         , (BANA.Checked));
    L1.AddChampSupValeur('EtablissUnique', (EtablissUnique));
    L1.AddChampSupValeur('AUXILIARE'    , (AUXILIARE.Checked));
    L1.AddChampSupValeur('HISTOSYNCHROValue', HISTOSYNCHRO.Value);
    L1.AddChampSupValeur('PSECTIONana'  , (PSECTIONana.Checked));
    L1.AddChampSupValeur('PJRL'         , (PJRL.Checked));
//GP    if (FTypeEcr.TEXT = '<<Tous>>') then FTypeEcr.Text := 'N;S;U;I;P;R';
    if (FTypeEcr.TEXT = Traduirememoire('<<Tous>>')) then FTypeEcr.Text := 'N;S;U;I;P;R';
    L1.AddChampSupValeur('BNORMAL'      , (pos ('N', FTypeEcr.Text) <> 0));
    L1.AddChampSupValeur('BSIMULE'      , (pos ('S', FTypeEcr.Text) <> 0));
    L1.AddChampSupValeur('BSITUATION'   , (pos ('U', FTypeEcr.Text) <> 0));
    L1.AddChampSupValeur('BIFRS'        , (pos ('I', FTypeEcr.Text) <> 0));
    L1.AddChampSupValeur('BPREVISION'   , (pos ('P', FTypeEcr.Text) <> 0));
    L1.AddChampSupValeur('BREVISION'    , (pos ('R', FTypeEcr.Text) <> 0));

    L1.AddChampSupValeur('TYPEECRX'     , TYPEECRX);
    L1.AddChampSupValeur('STATE'        , BExport.State);
    L1.AddChampSupValeur('NATUREJRL'    , NATUREJRL.value);
    L1.AddChampSupValeur('ECRINTEGRE'    , ECRINTEGRE);
    L1.AddChampSupValeur('ModePCL'      , V_PGI.ModePCL);
    L1.AddChampSupValeur('DATEMODIF1'   , DATEMODIF1.Text);
    L1.AddChampSupValeur('DATEMODIF2'   , DATEMODIF2.Text);
    L1.AddChampSupValeur('ANNULVALIDATION', (ANNULVALIDATION.Checked));
    L1.AddChampSupValeur('CPTEDEBUT'    , CPTEDEBUT.Text);
    L1.AddChampSupValeur('CPTEFIN'      , CPTEFIN.Text);
    L1.AddChampSupValeur('Tlibre'       , (TLibre.Checked)) ;
    L1.AddChampSupValeur('Zippe'        , (Zippe.Checked)) ;
    L1.AddChampSupValeur('ListeFichierJoint', ListeFichierJoint);
    L1.AddChampSupValeur('BGed'         , (BGed.Checked)) ;
    L1.AddChampSupValeur('PVentil'       , (PVentil.Checked)) ;
    if  (ZV1.Text <> '') and ((Trans.Typearchive = 'JRL') or (Trans.Typearchive = 'BAL')) then
    begin
          WhereAvancee := DecodeWhereSup;
          L1.AddChampSupValeur('WHEREAVANCE'       , WhereAvancee) ;
    end
    else
    begin
      if (stArg <> '') and (WhereAvancee <> '') then
          L1.AddChampSupValeur('WHEREAVANCE'       , WhereAvancee)
      else
      begin
          WhereAvancee := '';
          L1.AddChampSupValeur('WHEREAVANCE'       , WhereAvancee);
      end;
    end;
    // fiche 10431
    WhereNature := '';
    if (PARAMDOS.checked) then
    begin
        if PGen.Checked then
              WhereNature := Rendw('G_NATUREGENE')
        else
        if PTIERS.Checked then
              WhereNature := Rendw('T_NATUREAUXI')
        else
        if PJRL.Checked then
              WhereNature := Rendw('J_NATUREJAL');
    end;
    L1.AddChampSupValeur('WHERENATURE'       , WhereNature);
    

    Result := LT;
end;

procedure TFAssistCom.bFinClick(Sender: TObject);
var
Bufnature    : string;
retour       : Boolean;
FileName     : string;
Filearchive  : string;
Datearr      : TDateTime;
Q1           : TQuery;
Dates        : string;
Importance,i : integer ;
Rep          : string;
D1, D2       : TDATETIME;
{$IFNDEF EAGLCLIENT}
Listefichiers : HTStringList;
sTmp          : string;
{$ENDIF}
compte0       : string;
OkSisco       : Boolean;
ModeS         : Boolean;
Pointage      : string;
Hst           : HTStringList;
begin
  inherited;
        OkExoEncours := FALSE;
        Trans.Jr     := '';
        Dates        := '';
        Pointage := GetParamSocSecur ('SO_CPPOINTAGESX', 'EXP', TRUE);

        if (Trans.Serie = '') and (stArg <> '') then
        begin
                 PGIBox ('Veuillez renseigner la serie .', 'Envoyer');
                 if TobCol <> nil then
                 begin TobCol.free;  TobCol := nil; end;
                 exit;
        end;
        if (Trans.Serie <> 'SAGE') and (Trans.Serie <> 'IMMO')  and (Trans.Serie <> 'COMPTA') then
        begin
             Trans.Serie := TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex];
        end;
//GP        if (EXERCICE.TEXT <> '<<Tous>>')
        if (EXERCICE.TEXT <> Traduirememoire('<<Tous>>'))
        and (EXERCICE.TEXT <> '') then
            Trans.Exo              := EXERCICE.value;
        if (NATURETRANSFERT.text = '') and (Trans.Serie <> 'SAGE') and (Trans.Serie <> 'IMMO')
        and (Trans.Serie <> 'COMPTA') then
        begin
                 if (Trans.Typearchive <> '') then NATURETRANSFERT.text :=  Trans.Typearchive
                 else
                 begin
                   PGIBox ('Renseignez la nature du transfert.', 'Envoyer');
                   if TobCol <> nil then
                   begin TobCol.free;  TobCol := nil; end;
                   exit;
                 end;
        end;
        if (Trans.Serie = 'SISCO') and (Trans.Exo = '') and (not Quadra.Checked) then
        begin
                 PGIBox ('Veuillez ne sélectionner qu''un seul exercice.', 'Envoyer');
                 if TobCol <> nil then
                 begin TobCol.free;  TobCol := nil; end;
                 exit;
        end;
        if (Trans.Serie = 'SISCO') then// sisco
        begin
             if not Creationparamcol then exit;
             if not ExisteSQL('SELECT * from CORRESP Where CR_TYPE="SIS"') then
             begin
                  if not SANCOLL.checked  then
                  begin
                       if GetParamSocSecur ('SO_DEFCOLFOU', '') = '' then
                       begin
                        PGiInfo ('Veuillez renseigner le collectif fournisseur dans les paramètres société.', 'Transfert Sisco II');
                        exit;
                       end;
                       if GetParamSocSecur ('SO_DEFCOLCLI', '') = '' then
                       begin
                        PGiInfo ('Veuillez renseigner le collectif client dans les paramètres sociétés.','Transfert Sisco II');
                        exit;
                       end;
                  end;
             end;
        end;
        if (Trans.Serie = 'SISCO') and (NATURETRANSFERT.Value = 'DOS') then
        begin
                 PGIBox ('Transfert du dossier vers Sisco II impossible.', 'Envoyer');
                 NATURETRANSFERT.Value := 'JRL';
                  if TobCol <> nil then
                  begin TobCol.free;  TobCol := nil; end;
                  exit;
        end;
        if (PARAMDOS.checked) and
            (NATURETRANSFERT.Value <> 'BAL') then
        begin
              PGIBox ('Vous devez être en mode balance pour effectuer l''export des paramètres dossier.', 'Envoyer');
              NATURETRANSFERT.Value := 'BAL';
              exit;
        end;

//GP        if (Journaux.Text <> '<<Tous>>') and (Journaux.Text <> '') then
        if (Journaux.Text <> Traduirememoire('<<Tous>>')) and (Journaux.Text <> '') then
           Trans.Jr            := Journaux.Text
        else
//GP        if (NATUREJRL.Text <> '<<Tous>>') and (NATUREJRL.Text <> '') then
        if (NATUREJRL.Text <> Traduirememoire('<<Tous>>')) and (NATUREJRL.Text <> '') then
        begin
             for i:= 1 to Journaux.Items.count-1 do
             begin
                  if i= 1 then
                     Trans.Jr := JOURNAUX.Values[i]
                  else
                     Trans.Jr := Trans.Jr + ';' + JOURNAUX.Values[i];
             end;
        end;

        Trans.FichierSortie    := FICHENAME.Text;
        Trans.Dateecr1         := StrToDate(DATEECR1.Text);
        Trans.Dateecr2         := StrToDate(DATEECR2.Text);
//GP        if FETAB.Text = '<<Tous>>' then  Trans.Etabi := ''
        if FETAB.Text = Traduirememoire('<<Tous>>') then  Trans.Etabi := ''
        else Trans.Etabi            := FETAB.Text;
        Trans.Typearchive      := NATURETRANSFERT.Value;

        Trans.balance          := (NATURETRANSFERT.Value = 'BAL');
        if PGEN.Checked then   Trans.Pgene  := 'X' else Trans.Pgene  := '-';
        if PTIERS.Checked then Trans.ptiers := 'X' else Trans.ptiers := '-';
        if PTiersautre.Checked then Trans.ptiersautre := 'X' else Trans.ptiersautre := '-';
        if PANA.Checked then   Trans.pana   :='X'  else Trans.pana   := '-';
        if PSECTIONana.Checked then Trans.psection := 'X' else Trans.psection := '-';
        if PJRL.Checked then Trans.pjournaux    := 'X' else Trans.pjournaux := '-';

//        if (Trans.Typearchive = 'DOSS') then Trans.Typearchive := 'DOS';
        ListeCpteGen := nil;
        if ExisteSQl ('SELECT D_PARITEEURO FROM DEVISE Where D_PARITEEURO <= 0') then
        begin
                PGIInfo ('Problème parité fixe EURO d''une devise,  elle ne doit pas être à zéro et / ou négative.'+#10#13+
                'Veuillez modifier les paramètres comptables' , 'Envoyer');
                exit;
        end;
        if (stArg = '') then
        begin
           if
            ((GetParamSocSecur('SO_CPLIENGAMME', '', TRUE)<>Trans.Serie) and
            ((Trans.Typearchive = 'SYN') or (Trans.Typearchive = 'DOSS')))
            or
            (((Pointage <> 'CLI') and (Pointage <> 'EXP')) and
            ((Trans.Typearchive = 'SYN') or (Trans.Typearchive = 'DOSS')))
            then
            begin
                 //if not LienS3 then
                 if GetParamSocSecur('SO_CPLIENGAMME', '', TRUE)<>Trans.Serie then
                 begin
                    PGIBox ('Vous n''êtes pas en liaison avec '+Trans.Serie+#10#13
                    +' Pour effectuer un export de données vers '+Trans.Serie+','+#10#13
                    +' veuillez renseigner l''option "Paramétrage des échanges" dans bureau PGI.'+#10#13
                    +' ou Accès : Click droit, Menu Société', 'Envoyer');
                      exit;
                 end;
                 if (Pointage <> 'CLI') and (Pointage <> 'EXP')then
                 begin
                    if Existesql('SELECT EE_GENERAL FROM EEXBQ') then
                    begin
                      PGIBox (' Vous avez indiqué "Pas de gestion du pointage" dans les paramètres sociétés alors qu''il existe des références de pointage dans le '+#10#13
                      +' dossier. Vous devez mettre en cohérence les éléments (supprimer les références qui n''ont pas lieu d''être ou indiquer qui effectue le '+#10#13
                      +' pointage) afin de pouvoir générer un fichier de synchronisation.'+#10#13, 'Envoi impossible');
                      exit;
                    end
                    else
                    if (Trans.Serie = 'S1') then
                    begin
                      PGIBox (' Vous avez indiqué "Pas de gestion du pointage" dans les paramètres sociétés. '+#10#13
                      +' Vous devez indiquer pointage par Expert comptable ou par Client '+#10#13
                      +' afin de pouvoir générer un fichier de synchronisation.'+#10#13, 'Envoi impossible');
                      exit;
                    end
                    else
                      PGIBox ('Vous avez choisi de ne pas faire de pointage.'+#10#13
                      +' Si vous voulez en faire, veuillez renseigner l''option  "Gestion du pointage."'+#10#13
                      +' "Paramétrage des échanges" dans Bureau PGI Expert,'+#10#13
                      +' ou accès par Click droit, Menu Société', 'Envoyer');
                 end;
            end;
        end;

        ModeS := GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE, FALSE);
        if (stArg = '') and (Trans.Serie = 'S1') then
        begin
                if (Trans.Typearchive = 'JRL') and (ModeS = TRUE) then
                begin
                       retour := (PGIAsk ('Attention, les données en mode "Journal" que vous allez'+ #10#13
                        +' envoyer ne contiennent aucune information de synchronisation (risque de doublon).'+ #10#13
                        +' Confirmez-vous l''export ?','Export')=mrYes);
                        if not retour then exit;
                end;
                if (Trans.Typearchive = 'SYN') and (ModeS = FALSE) then
                begin
                  if GetparamsocSecur ('SO_CPSYNCHROSX', TRUE) = TRUE then
                     SetParamsoc ('SO_FLAGSYNCHRO', 'SYN')
                  else
                  begin
                     PGIInfo ('Attention le dossier n''est pas paramétré en mode synchronisation.'+ #10#13
                     +'Paramètres société, Echanges, cochez l''option Synchronisation'); exit;
                  end;
                end;
                    // fiche 18836
                 if ModeS and (Trans.Typearchive <> 'JRL') and (Trans.Typearchive <> 'BAL') and (Pointage <> 'CLI') and (Pointage <> 'EXP')then
                 begin
                      PGIBox (' Vous avez indiqué "Pas de gestion du pointage" dans les paramètres sociétés. '+#10#13
                      +' Vous devez indiquer pointage par Expert comptable ou par Client '+#10#13
                      +' afin de pouvoir générer un fichier de synchronisation.'+#10#13, 'Envoi impossible');
                      exit;
                 end;
                 if ExisteSQL ('select  ##TOP 1##y_axe from analytiq where (y_axe="A2" or y_axe="A3" or y_axe="A4" or y_axe="A5")') then
                 begin
                             PGIInfo ('Vous effectuez des échanges synchronisés avec Business Line. Export impossible.'
                             + #10#13
                             +' Un axe (analytique ou TVA) autre que l''axe A1 est utilisé');  exit;
                 end;

        end;

        if (stArg = '') and ((Trans.Serie = 'S1') or (Trans.Typearchive = 'SYN')) or
        ((ctxPCL in V_PGI.PGIContexte) and (NATURETRANSFERT.Value = 'DOSS')) then
        begin
           if (Trans.Serie = 'S1') then
           begin
             if (GetParamSocSecur('SO_CPLIENGAMME', '', TRUE) = '')  then
                SetParamsoc ('SO_CPLIENGAMME', 'S1')
             else
             if ( GetParamSocSecur('SO_CPLIENGAMME', '', TRUE) <> 'S1') then
             begin
                    PGIBox ('Vous n''êtes pas en liaison avec Business Line'+#10#13
                    +' Pour effectuer un export de données vers Business Line,'+#10#13
                    +' veuillez renseigner l''option "Paramétrage des échanges" dans bureau PGI.'+#10#13
                    +' ou Accès : Click droit, Menu Société', 'Envoyer');
                      exit;
             end;
             // fiche 18836
             if ModeS and (Trans.Typearchive <> 'JRL') and (Trans.Typearchive <> 'BAL') and (Pointage <> 'CLI') and (Pointage <> 'EXP')then
             begin
                  if Existesql('SELECT EE_GENERAL FROM EEXBQ') then
                  begin
                    PGIBox (' Vous avez indiqué "Pas de gestion du pointage" dans les paramètres sociétés alors qu''il existe des références de pointage dans le '+#10#13
                    +' dossier. Vous devez mettre en cohérence les éléments (supprimer les références qui n''ont pas lieu d''être ou indiquer qui effectue le '+#10#13
                    +' pointage) afin de pouvoir générer un fichier de synchronisation.'+#10#13, 'Envoi impossible');
                    exit;
                  end
                  else
                    PGIBox ('Vous avez choisi de ne pas faire de pointage.'+#10#13
                    +' Si vous voulez en faire, veuillez renseigner l''option  "Gestion du pointage."'+#10#13
                    +' "Paramétrage des échanges" dans Bureau PGI Expert,'+#10#13
                    +' ou accès par Click droit, Menu Société', 'Envoyer');
             end;
           end;
           if (Trans.Typearchive = 'SYN') and
              (GetParamSocSecur ('SO_CPDECLOTUREDETAIL', '')= TRUE) then
           begin
                PGIBox ('L''export de type synchronisation sera impossible car une suppression du journal d''à nouveaux a été effectué dans la comptabilité.'+#10#13+
                ' Veuillez sélectionner l''export type "dossier"', 'Envoyer');
                exit;
           end;

           if (Trans.Typearchive = 'SYN') and
              (GetParamSocSecur ('SO_CPSYNCHROSX', '')= TRUE) then
           begin

               if (GetParamSocSecur ('SO_CPDATESYNCHRO1', '') <=  GetParamSocSecur ('SO_CPDATESYNCHRO2', '')) then
               begin
                  if (GetParamSocSecur('SO_CPDATESYNCHRO2', '') <> 0) and (GetParamSocSecur ('SO_CPDATESYNCHRO2', '') <> iDate1900) then Dates := 'le '+ DateToStr(GetParamSocSecur ('SO_CPDATESYNCHRO2', ''));
               end else   if (GetParamSocSecur('SO_CPDATESYNCHRO1', '') <> 0) and (GetParamSocSecur ('SO_CPDATESYNCHRO1', '') <> iDate1900) then Dates := 'le '+ DateToStr(GetParamSocSecur ('SO_CPDATESYNCHRO1', ''));

               if stArg <> '' then retour := TRUE
               else
               retour := (PGIAsk ('Un envoi de type synchronisation a été effectué '+ Dates +'.'+ #10#13
                +' A ce jour aucune réception n''a été enregistrée.'+ #10#13+' Confirmez-vous l''export ?','Synchronisation ')=mrYes);
                if not retour then exit;
           end;
        end
        else
           Trans.TypeFormat       := TYPEFORMAT.Value;

        if (Trans.Serie = 'S1') then  // fiche 15527
        begin
              ListeCpteGen:=TList.Create;
              ChargementCpte(ListeCpteGen, '');
              if (CompteEstalpha (ListeCpteGen)) then
              begin
                  PgiBox ('Export impossible.'+#10#13+
                  ' Vous avez des comptes généraux alpha numériques.', 'Envoyer');
                  LibererGen (ListeCpteGen); ListeCpteGen.free;
                  exit;
              end;
              // fiche 10326
              compte0 :=  BourreOuTronque('00000000000000000', fbGene);
              If ExisteSQL ('select G_GENERAL from generaux where G_GENERAL="'+Compte0+'"') then
              begin
                  PgiBox ('Export impossible.'+#10#13+
                  ' Vous avez un compte général '+Compte0, 'Envoyer');
                  LibererGen (ListeCpteGen); ListeCpteGen.free;
                  exit;
              end;
              LibererGen (ListeCpteGen); ListeCpteGen.free;
                // fiche 10574
              Trans.TypeFormat := 'ETE';
              Q1 := OpenSQL ('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" ORDER BY E_DATECOMPTABLE DESC',True);
              if not Q1.EOF then
              begin
                      Trans.Exo := Q1.FindField('E_EXERCICE').AsString;
                      OkExoEncours := TRUE;
              end;
              ferme (Q1);
              if OkExoEncours then
              begin
                          COMDExoToDates(Trans.Exo, D1, D2) ;
                          Q1 := OpenSql ('SELECT * from EXERCICE WHERE EX_DATEDEBUT > "'+ USDATETIME(D1)+'"',FALSE);
                          While not Q1.EOF do
                          begin
                                    Trans.Exo := Trans.Exo+';'+Q1.FindField ('EX_EXERCICE').asstring;
                                    Q1.next;
                          end;
                          ferme (Q1);
              end;
        end;

//GP        if NATUREJRL.Text = '<<Tous>>' then  Trans.naturejal := ''
        if NATUREJRL.Text = Traduirememoire('<<Tous>>') then  Trans.naturejal := ''
        else
        Trans.naturejal        := NATUREJRL.value;

        if(Trans.Typearchive = 'DOS') then
        Bufnature := 'Transfert : Dossier'
        else
            if(Trans.Typearchive = 'BAL') then
               Bufnature := 'Transfert : Balance '
               else if (Trans.Typearchive = 'JRL') then
                   Bufnature := 'Transfert : Journal'
               else if(Trans.Typearchive = 'SYN') then
                   Bufnature := 'Transfert : Synchronisation';
        if Trans.Serie = 'SISCO' then
        begin
                 FileName := UpperCase(Trans.FichierSortie);
                 if pos('.TRT', FileName) = 0 then   // Fiche
                 begin
                      //FileName := ReadTokenPipe (Filename, '.');
                      FileName := FileName + '.TRT';
                 end;
                 FICHENAME.TExt := FileName;
                 Trans.FichierSortie :=  FileName;
                 ListeCpteGen:=TList.Create;
                 if ( GetParamSocSecur('SO_CPLIENGAMME', '', TRUE) = '')  then
                    SetParamsoc ('SO_CPLIENGAMME', 'SI')
                 else
                 if (stArg = '') and ( GetParamSocSecur('SO_CPLIENGAMME', '', TRUE) <> 'SI') then
                 begin
                    PGIBox ('Vous n''êtes pas en liaison avec Sisco'+#10#13
                    +' Pour effectuer un export de données vers Sisco,'+#10#13
                    +' veuillez renseigner l''option "liaison avec une comptabilité".'+#10#13
                    +' Accès : "Révision à distance" des paramètres sociétés de la comptabilité,'+#10#13
                    +' l''option Liaison avec une comptabilité', 'Envoyer');
                      exit;
                 end;

                 ChargementCpte(ListeCpteGen, '');
                 While GetPage <> PEXPORT do
                      bSuivantClick(Sender);

                 if not ControleSisco then
                 begin
                      if TobCol <> nil then
                      begin TobCol.free;  TobCol := nil; end;
                      bPrecedent.Enabled := FALSE;
                      bFin.Enabled       := FALSE;
                      LibererGen (ListeCpteGen); ListeCpteGen.free;
                      exit;
                 end;
                 if VH^.Cpta[fbGene].Lg <> VH^.Cpta[fbAux].Lg then
                 begin
                      PGIBox ('Le transfert est impossible.#10#13Les comptes généraux et auxiliaires sont de longueurs différentes','Export Sisco II');
                      if TobCol <> nil then
                      begin TobCol.free;  TobCol := nil; end;
                      exit;
                 end;

                 if CompteEstalpha (ListeCpteGen) then
                 begin
                    PgiBox ('Export impossible.'+#10#13+' Vous avez des comptes généraux alpha numériques.', 'Envoyer');
                    if TobCol <> nil then
                    begin TobCol.free;  TobCol := nil; end;
                    LibererGen (ListeCpteGen); ListeCpteGen.free;
                    exit;
                 end;
                 // ajout me 28-01-2005
                 LibererGen (ListeCpteGen);
                 ListeCpteGen.free;
        end;

        if ZIPPE.checked then
        begin
             FileName := FICHENAME.Text;
             Filearchive := ReadTokenPipe (Filename, '.');
             Filearchive := Filearchive + '.zip';
             FileName :=  Filearchive;
        end
        else
             FileName := FICHENAME.Text;

        if stArg <> '' then Retour := TRUE
        else
        retour := (PGIAsk ('Confirmez-vous le traitement ?', 'Envoyer')=mrYes);

        if Retour then
        begin
             Datearr := StrToDate (DATEECR.Text);
             DateFichierClo :=  FormatDateTime(Traduitdateformat('ddmmyyyy'),Datearr);
             if (Trans.Serie <> 'SAGE') and (Trans.Serie <> 'IMMO') and (Trans.Serie <> 'COMPTA')then
                While GetPage <> PEXPORT do
                      bSuivantClick(Sender);
             if (Trans.Serie = 'SISCO') then// sisco
             begin
                 if Trans.Dateecr1 = iDate1900 then
                 begin
                      Trans.Dateecr1 := VH^.Encours.Deb;
                      Trans.Dateecr2 := VH^.Encours.Fin;
                 end;
                 Trans.psection         := CP_STAT.value;

                 if TobCol <> nil then
                 begin TobCol.free;  TobCol := nil; end;
                 AfficheListeCom('Debut Export ', LISTEEXPORT);
                 if TYPEECRX='' then TYPEECRX := 'N';
                 if Quadra.Checked then
                     OkSisco := TransferversQuadra (Trans.FichierSortie,LISTEEXPORT, Trans, SANCOLL.checked, CarCli, CarFou, TYPEECRX, Quadra.Checked)
                 else
                     OkSisco := TransfertPgiVersSisco(Trans.FichierSortie,LISTEEXPORT, Trans, SANCOLL.checked, CarCli, CarFou, TYPEECRX, Quadra.Checked, DATARRETEBAL.text) ;
                 if not OkSisco then
                          AfficheListeCom('Export annulé ', LISTEEXPORT)
                 else
                 begin
                           AfficheListeCom('Export terminé ', LISTEEXPORT);
                           if FichierImp = '' then
                           FichierImp := 'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt';
                           EcrireDansfichierListeCom (FichierImp, LISTEEXPORT);
                           if stArg = '' then
                           begin
                                if PGIAsk('Export terminé : voulez-vous imprimer le rapport ?','Envoyer')=mrYes then
                                   if PrinterSetupDialog1.Execute then
                                      ControlTextToPrinter(FichierImp,poPortrait);
                           end
                           else
                           begin
                                 Rep := GetParamSocSecur('SO_CPRDREPERTOIRE', '');
                                 if Rep = '' then
                                 Rep := ExtractFileDir (Trans.FichierSortie);
                                 if FichierImp = '' then
                                    FichierImp := Rep+'\'+'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt'
                                 else
                                    FichierImp := Rep+'\'+ FichierImp;
                           end;
                           FTimer.free;
                           ModalResult := 1;
                  end;
             end
             else
             begin
               if Trans.Serie = 'SAGE' then
               begin
                  Trans.TypeFormat := 'SAGE';
                  Trans.Typearchive := 'JRL';
                  ExportComSage;
               end
               else
                  ExportComSx;
             end;
{$IFNDEF EAGLCLIENT}
            if ZIPPE.checked then
            begin
               Rep := FICHENAME.Text;
               sTmp := ReadTokenPipe (Rep, '.')+ '.ZIP';
               if FileExists(sTmp) then
                  DeleteFile (sTmp);

               FileClickzip;
               if ListeFichierJoint <> '' then
               begin
                    sTmp     := ListeFichierJoint;
                    Listefichiers := HTStringList.Create;
                    while ( length(sTmp) <> 0 ) do Listefichiers.add(ReadTokenPipe(sTmp,';'));
                    for i:=0 to Listefichiers.Count-1 do
                        if (Listefichiers.Strings[i]<> '') then
                           FileClickzip(Listefichiers.Strings[i]);
                    Listefichiers.free;
               end;
{$IFDEF SCANGED}
               if BGed.Checked then
               begin
                   Rep := ExtractFileDir(Trans.FichierSortie)+'\GED';
                   FileClickzip(Rep, TRUE);
                   RemoveInDir2 (Rep, TRUE, TRUE, TRUE);
               end;
{$ENDIF}
            end;
{$IFDEF SCANGED}
            if BGed.Checked then
                 RemoveInDir2(GetEnvVar('TEMP'), FALSE, FALSE, FALSE,'\Agl*.pdf');
{$ENDIF}
{$ENDIF}
            if FMail.Checked then
            begin
                  if stArg <> '' then
                  begin
                       caption :=  'Fichier d''export Comsx';
                       FEMail.Text := Email;
                       FCorpsMail.Lines[0] := TraduireMemoire('Ci-joint le fichier d''export de Comsx et le rapport');
                       FCorpsMail.Lines[1] := '';
                       FCorpsMail.Lines[4] := '';

                  end;
                   if FHigh.Checked then Importance:=2 else Importance:=1 ;

                   RenvoieCorpsMail (Hst);
                   AglSendMail(caption ,FEMail.Text,'', Hst,FICHENAME.text+';'+FichierImp,true,Importance,'','') ;
                   Hst.free;
            end ;

        end;

end;


procedure TFAssistCom.NATURETRANSFERTChange(Sender: TObject);
var
OkBal,OkType     : Boolean;
Trans            : string;
Dfin             : TDateTime;
begin
  inherited;
  OkBal :=  (NATURETRANSFERT.Value = 'BAL');
  OkType :=  (NATURETRANSFERT.Value = 'BAL')
  or (NATURETRANSFERT.Value = 'JRL');
  GRPBAL.Enabled            := OkBal;
  AUXILIARE.Enabled         := OkBal;
  BAna.Enabled              := OkBal;
  GRTYPECR.Enabled          := OkType;
  Trans := TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex];

  if (Trans <> 'SISCO') then
  begin
    FTypeEcr.Enabled := OkType;
    TFTypeEcr.Enabled := OkType;
    Exclure.Enabled     := OkType;
    TExclure.Enabled    := OkType;
  end;
  RendVisibleChamp;

  if NATURETRANSFERT.Value = 'SYN' then
  begin
   HISTOSYNCHRO.Enabled      := TRUE;
   THISTOSYNCHRO.Enabled     := TRUE;

   TYPEFORMAT.Value          := 'ETE';
   TYPEFORMAT.Enabled        := FALSE;
   TTYPEFORMAT.Enabled       := FALSE;
   HISTOSYNCHRO.Items.clear;
   HISTOSYNCHRO.Values.clear;
   HISTOSYNCHRO.Items.Add(TraduireMemoire('<<Encours>>'));
   HISTOSYNCHRO.Values.add('X');
   if (GetParamSocSecur ('SO_CPDATESYNCHRO1', idate1900) <=  GetParamSocSecur ('SO_CPDATESYNCHRO2', idate1900)) then
   begin
         if (GetParamSocSecur('SO_CPDATESYNCHRO1', idate1900) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO1', idate1900) <> iDate1900) then
         begin
             HISTOSYNCHRO.Items.Add(FormatDateTime(Traduitdateformat('dddd dd mmmm yyyy "à" hh:nn'), GetParamSocSecur('SO_CPDATESYNCHRO1', FALSE)));
             HISTOSYNCHRO.Values.add('0');
         end;
         if (GetParamSocSecur('SO_CPDATESYNCHRO2', idate1900) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO2', idate1900) <> iDate1900) then
         begin
             HISTOSYNCHRO.Items.Add(FormatDateTime(Traduitdateformat('dddd dd mmmm yyyy "à" hh:nn'), GetParamSocSecur('SO_CPDATESYNCHRO2', idate1900)));
             HISTOSYNCHRO.Values.add('1');
         end;
   end
   else
   begin
         if (GetParamSocSecur('SO_CPDATESYNCHRO2', idate1900) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO2', idate1900) <> iDate1900) then
         begin
              HISTOSYNCHRO.Items.Add(FormatDateTime(Traduitdateformat('dddd dd mmmm yyyy "à" hh:nn'), GetParamSocSecur('SO_CPDATESYNCHRO2', FALSE)));
              HISTOSYNCHRO.Values.add('1');
         end;
         if (GetParamSocSecur('SO_CPDATESYNCHRO1', idate1900) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO1', idate1900) <> iDate1900) then
         begin
              HISTOSYNCHRO.Items.Add(FormatDateTime(Traduitdateformat('dddd dd mmmm yyyy "à" hh:nn'), GetParamSocSecur('SO_CPDATESYNCHRO1', idate1900)));
              HISTOSYNCHRO.Values.add('0');
         end;
    end;
         HISTOSYNCHRO.ItemIndex := 0;
  end
  else
  begin
         THISTOSYNCHRO.Enabled    := FALSE;
         HISTOSYNCHRO.Enabled     := FALSE;
  end;
  if (NATURETRANSFERT.Value = 'DOS') or (NATURETRANSFERT.Value = 'DOSS') then DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), iDate1900)
  else
  begin
       // Fiche 10054
       if  BNetExpert.Checked then  DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), iDate1900)
       else
       begin
             if (Trans <> 'S1')  and (not(ctxPCL in V_PGI.PGIContexte)) then
             begin
                Dfin := PlusMois(V_PGI.DateEntree, -1);
                Dfin := FinDeMois(Dfin);
                DATEECR.Text := FormatDateTime(Traduitdateformat('dd/mm/yyyy'), DFin);
             end
             else
                DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), now);
       end;
  end;
  if (Trans = 'SISCO') or (Trans = 'S1') then
  begin
     AvecLettrage.checked := FALSE;
     AvecLettrage.visible := FALSE;
  end
  else
  begin
     AvecLettrage.Enabled           := ((NATURETRANSFERT.Value = 'DOSS') or (NATURETRANSFERT.Value = 'JRL'));
     AvecLettrage.Checked           := FALSE;
     PEcrBud.Checked                := FALSE;
     PEcrBud.Enabled                := (NATURETRANSFERT.Value = 'JRL') or (NATURETRANSFERT.Value = 'BAL');
  end;
  DATEMODIF1.Text := stDate1900;
  DATEMODIF2.Text := stDate2099;
  // Pour international
  DATEECR1.Text := stDate1900;
  DATEECR2.Text := stDate2099;

  if (NATURETRANSFERT.Value = 'JRL') then
  begin
          GroupBox3.Caption := 'Ecritures Créées';
          DATEMODIF1.Enabled  := TRUE;
          DATEMODIF2.Enabled  := TRUE;
          TDATEMODIF1.Enabled := TRUE;
          TDATEMODIF2.Enabled := TRUE;
          GroupBox3.Enabled   := TRUE;
          if starg = '' then // fiche 10501
          begin
              PGEN.Checked        := TRUE;
              PTIERS.Checked      := TRUE;
              PTiersautre.Checked := FALSE;
              PSectionana.Checked := TRUE;
              PJRL.Checked        := TRUE;
              PPARAMSOC.Checked   := TRUE;
              PEXERCICE.Checked   := TRUE;
              TLibre.Checked      := FALSE;
              PVentil.Checked     := FALSE;
              FTypeEcr.Text       := 'N;';
              AvecLettrage.Checked := FALSE;
              ANNULVALIDATION.Checked := FALSE;
              PECRBUD.Checked     := FALSE;
              Exclure.Text        := '';
          end;
  end
  else
  begin
          GroupBox3.Caption := 'Modifiés';
          DATEMODIF1.Enabled  := FALSE;
          DATEMODIF2.Enabled  := FALSE;
          TDATEMODIF1.Enabled := FALSE;
          TDATEMODIF2.Enabled := FALSE;
          GroupBox3.Enabled   := FALSE;
  end;
end;

procedure TFAssistCom.RendVisibleChamp;
var
Typef     : Boolean;
Q1        : TQuery;
TValue    : string;
Repertoire: string;
formatSq  : string;
NoMax     : integer;
TNetRecep : TOB;
TA        : TOB;
Procedure GrisCombo(Grise : Boolean);
begin
       GrisechampCombo (EXERCICE, Grise);
       TEXERCICE.Enabled  := Grise;
       GrisechampCombo (NATUREJRL, Grise);
       TNATUREJRL.Enabled := Grise;
       JOURNAUX.Enabled := Grise;
       TJOURNAUX.Enabled  := Grise;
       FETAB.Enabled := Grise;
       TFETAB.Enabled     :=  Grise;
       GrisechampCritEdit (DATEECR1, Grise);
       TDATEECR1.Enabled  := Grise;
       GrisechampCritEdit (DATEECR2, Grise);
end;

begin
  TValue := TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex];
  typef := ((TValue <> 'SISCO') and (TValue <> 'S1') and
     (NATURETRANSFERT.value = 'BAL'));

  TYPEFORMAT.ItemIndex := 1;

  if (TValue = 'SISCO') then
   TYPEFORMAT.ItemIndex := 0;

  TYPEFORMAT.Enabled := typef;
  TTYPEFORMAT.Enabled := typef;

  PARAMDOS.Enabled :=  typef;
  if (NATURETRANSFERT.value <> 'BAL') then
     PARAMDOS.Checked :=  FALSE;

  if (TValue = 'SISCO') or (TValue = 'S1') then
  begin
     PARAMDOS.checked := typef;
     if (TValue = 'S1') then
         GRTYPECR.Enabled    := FALSE;
  end;

  BExport.Enabled := (NATURETRANSFERT.Value = 'JRL') or (NATURETRANSFERT.Value = 'BAL');

  if ((NATURETRANSFERT.Value = 'SYN')
  or (NATURETRANSFERT.Value = 'DOS')) then
  begin
       GrisCombo(FALSE);
       TYPEFORMAT.ItemIndex := 1;
       TYPEFORMAT.Value := 'ETE';
  end
  else
  begin
       if NATURETRANSFERT.Value <> 'EXE' then
       begin
                 EXERCICE.Items.clear;
                 EXERCICE.Values.clear;
                 EXERCICE.Items.Add(TraduireMemoire('<<Tous>>'));
                 EXERCICE.Values.add(' ');

                 Q1 := OpenSql ('SELECT EX_LIBELLE,EX_EXERCICE from EXERCICE ',FALSE);
                 While not Q1.EOF do
                 begin
                         EXERCICE.Items.Add(Q1.FindField('EX_LIBELLE').asstring);
                         EXERCICE.Values.add(Q1.FindField('EX_EXERCICE').asstring);
                         Q1.next;
                 end;
                 ferme (Q1);
       end;
       if (NATURETRANSFERT.Value = 'BAL')
       // ajout me 09-12-2002
       or (NATURETRANSFERT.Value = 'EXE')then
       begin
           FETAB.Enabled := TRUE;
           TFETAB.Enabled := TRUE;
           GrisechampCombo (EXERCICE, TRUE);
           TEXERCICE.Enabled  := TRUE;
           GrisechampCombo (NATUREJRL, FALSE);
           TNATUREJRL.Enabled := FALSE;
           JOURNAUX.Enabled := FALSE;
           TJOURNAUX.Enabled  := FALSE;
           TYPEFORMAT.ItemIndex := 0;
           TYPEFORMAT.Value := 'STD';
           if (NATURETRANSFERT.Value = 'EXE') then
           begin
                 EXERCICE.Items.clear;
                 EXERCICE.Values.clear;

                 // pour sisco acquisition on envoie l'exercice clôturé
                 Q1 := OpenSql ('SELECT EX_LIBELLE,EX_EXERCICE from EXERCICE WHERE EX_ETATCPTA="CDE" and EX_DATEDEBUT="'+USDateTime(VH^.Precedent.Deb) + '"',FALSE);
                 if not Q1.EOF  then
                 begin
                         EXERCICE.Items.Add(Q1.FindField('EX_LIBELLE').asstring);
                         EXERCICE.Values.add(Q1.FindField('EX_EXERCICE').asstring);
                 end;
                 ferme (Q1);
                 Q1 := OpenSql ('SELECT EX_LIBELLE,EX_EXERCICE from EXERCICE WHERE EX_ETATCPTA="OUV"  ORDER BY EX_EXERCICE',FALSE);
                 if not Q1.EOF  then
                 begin
                         EXERCICE.Items.Add(Q1.FindField('EX_LIBELLE').asstring);
                         EXERCICE.Values.add(Q1.FindField('EX_EXERCICE').asstring);
                 end;
                 ferme (Q1);
                 GrisechampCritEdit (DATEECR1, FALSE);
                 TDATEECR1.Enabled :=  FALSE;
                 GrisechampCritEdit (DATEECR2, FALSE);
                 TDATEECR2.Enabled := FALSE;
           end
           else
           begin
                 GrisechampCritEdit (DATEECR1, TRUE);
                 TDATEECR1.Enabled := TRUE;
                 GrisechampCritEdit (DATEECR2, TRUE);
                 TDATEECR2.Enabled := TRUE;
           end;
       end
       else
           GrisCombo(TRUE);
  end;
  if (TValue = 'SISCO') then
  begin
    TYPEFORMAT.Enabled   := FALSE;
    TTYPEFORMAT.Enabled  := FALSE;
    GRPBAL.Enabled       :=FALSE;
    GRTYPECR.Enabled     := FALSE;
    PARAMSISCOII.Enabled := TRUE;
    BDelete1.Enabled     := TRUE;
    if (NATURETRANSFERT.Value <> 'BAL') then
    begin
        DATARRETEBAL.visible := FALSE;
        TDATEBAL.visible     := FALSE;
    end
    else
    begin
        DATARRETEBAL.visible := TRUE;
        TDATEBAL.visible     := TRUE;
    end;
  end
  else
  begin
    PARAMSISCOII.Enabled := FALSE;
    BDelete1.Enabled     :=  FALSE;
  end;
  TDATEECR.Enabled :=  (NATURETRANSFERT.value = 'SYN');
  DATEECR.Enabled  :=  (NATURETRANSFERT.value = 'SYN');
  Label15.Visible  :=  (NATURETRANSFERT.value = 'SYN') and (not(ctxPCL in V_PGI.PGIContexte));
  if (TValue = VH^.CPLienGamme) then
  begin
     BNetExpert.Checked := IsDossierNetExpert (V_PGI.NoDossier, NoSEqNet);
     if BNetExpert.Checked and (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE) = FALSE) then
        Setparamsoc ('SO_CPMODESYNCHRO', TRUE);
  end;
  if  BNetExpert.Checked then
  begin
        formatSq := '%.03d';
        if NoSeqNet >= 999 then  formatSq := '%.04d';
        if NoSeqNet >= 9999 then  formatSq := '%.05d';
        if NoSeqNet = 0 then inc (NoSeqNet);
        TNetRecep := TOB.Create('', nil, -1);
        TA := TOB.Create ('',TNetRecep,-1);
        TA.AddChampSupValeur('DOMAINE', 'COMPTA');
        TA.AddChampSupValeur('CLIID',  V_PGI.NoDossier);
        TA.AddChampSupValeur('CLINOM', V_PGI.NomSociete);
        TA.AddChampSupValeur('FICMSQ', 'TRA');
        NoMax := MaxNumSeq(TNetRecep);
        TNetRecep.free;
        if NoMax >= NoSeqNet then NoSeqNet := NoMax+1;
        FicheName.Text := GetEnvVar('TEMP')+'\'+'PG'+ V_PGI.NoDossier+'_'+ Format (formatSq, [NoSeqNet])+'.TRA';
        FicheName.Enabled := FALSE;
        BNetExpert.Enabled := FALSE;
  end
  else
  begin
        Repertoire := GetParamSocSecur('SO_CPRDREPERTOIRE', '');
        if (Repertoire <> '') and (FicheName.Text = '') then
           FicheName.Text := Repertoire +'\'+'PG'+ V_PGI.NoDossier+'.TRA';
        BNetExpert.visible := FALSE;
  end;
  if (NATURETRANSFERT.Value = 'SYN') then
  begin
       if (TValue <> 'S1')  then
          TDATEECR.Caption := 'Date de clôture périodique'
       else
          TDATEECR.Caption := 'Date d''arrêté comptable';
       if (ctxPCL in V_PGI.PGIContexte) then
       begin
            Label15.Visible  := TRUE;
            Label15.Caption := 'Une clôture périodique sera générée chez votre client';
       end;
  end;
end;
procedure TFAssistCom.bSuivantClick(Sender: TObject);
var
St,Origine,Nature : string;
      Procedure     AfficheEtape;
      begin
        inc(Psuiv);
        lEtap.Caption := Msg.Mess[0] + ' ' + IntToStr(PSuiv);
      end;
begin
  inherited;
     if (Getpage = PPARAMGENE) and ((NATURETRANSFERT.Value = 'DOS')
       or (NATURETRANSFERT.Value = 'SYN')
       or (NATURETRANSFERT.Value = 'DOSS')) then
      begin
           bSuivantClick(Sender);
           exit;
      end;
      if (Getpage = TypeEcr) and (GRPBAL.Enabled = FALSE) and (GRTYPECR.Enabled = FALSE) then
      begin
           bSuivantClick(Sender);
           exit;
      end
      else
      if (Getpage = PPARAMGENE) and (PARAMDOS.checked) then
      begin
           bSuivantClick(Sender);
           exit;
      end
      else
      if (Getpage = PARAMSISCOII) and (TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex] <> 'SISCO') then
      begin
           bSuivantClick(Sender);
           exit;
      end;

      if (Getpage = Resume) or (Getpage = PEXPORT) then
      begin
         BFiltre.Enabled := TRUE;
         BINI.Enabled := TRUE;
      end
      else
      begin
         BFiltre.Enabled := FALSE;
         BINI.Enabled := FALSE;
      end;

      if Getpage = Resume then
      begin
           Origine := TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex];

           if Origine = 'SISCO' then
           St := 'Sisco II'
           else
           St := Origine;
           Nature := NATURETRANSFERT.Text;
          if Nature = 'JRL' then Nature := ' Nature : Journal' else
          if Nature = 'BAL' then Nature := ' Nature : Balance' else
          if (Nature = 'DOS') or (Nature = 'DOSS') then Nature := ' Nature : Dossier' else
          if Nature = 'SYN' then Nature := ' Nature : Synchronisation '+ Origine else
          if Nature = 'EXE' then Nature := ' Nature : Exercice';
          FVal1.caption := St + ' ' + Nature;
          FVal2.Caption := 'Du ' + DATEECR1.Text + ' au ' +  DATEECR2.Text;
          FVal3.Caption := FICHENAME.Text;
          FVal4.Caption := FEMAIL.Text;
      end;
      if Getpage = PEXPORT then
      begin
                 LISTEEXPORT.clear;
                 AfficheListeCom('*** Bouton <Fin> pour lancer le traitement *** ', LISTEEXPORT);
      end;
     AfficheEtape;
end;

procedure TFAssistCom.bPrecedentClick(Sender: TObject);
      Procedure     AfficheEtape;
      begin
        dec(PSuiv);
        if PSuiv <= 0 then PSuiv := 1;
        lEtap.Caption := Msg.Mess[0] + ' ' + IntToStr(PSuiv);
      end;
begin
  inherited;
     if (Getpage = PPARAMGENE) and ((NATURETRANSFERT.Value = 'DOS')
       or (NATURETRANSFERT.Value = 'SYN')
       or (NATURETRANSFERT.Value = 'DOSS')) then
      begin
           bPrecedentClick(Sender); exit;
      end;
     if (Getpage = TypeEcr) and (GRPBAL.Enabled = FALSE) and (GRTYPECR.Enabled = FALSE)then
      begin
           bPrecedentClick(Sender); exit;
      end
      else
      if (Getpage = PPARAMGENE) and (PARAMDOS.checked) then
      begin
           bPrecedentClick(Sender); exit;
      end
      else
      if (Getpage = PARAMSISCOII) and (TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex] <> 'SISCO') then
      begin
           bPrecedentClick(Sender); exit;
      end;
     AfficheEtape;
     if (Getpage = Resume) or (Getpage = PEXPORT) then
     begin
         BFiltre.Enabled := TRUE;
         BINI.Enabled := TRUE;
     end
     else
     begin
         BFiltre.Enabled := FALSE;
         BINI.Enabled := FALSE;
     end;

end;

procedure TFAssistCom.NATUREJRLClick(Sender: TObject);
begin
  inherited;
        TraiteListeJournal (NATUREJRL.value);
end;

function TFAssistCom.RenseigneArg(var Fichier,Commande : String) : Boolean;
var
tmp,tmp2,QuelExo   : string;
FicIni             : TIniFile;
NomFichT,St        : string;
               Procedure IniEnabled (Champ : string; CTR : TWinControl);
               begin
                  St := FicIni.ReadString (Commande, Champ, '');
                  TCheckBox(CTR).Checked := (St = 'TRUE') or (St = 'X');
               end;
begin
            Result := TRUE;
            ListeFichierJoint := '';
            if pos('Minimized', stArg) <> 0 then
                 WindowState  := wsMinimized;
            FTimer.Enabled := TRUE;
            tmp            := stArg;
            if (pos('.INI', Uppercase(tmp)) <> 0) then
            begin
                      if Fichier = '' then
                      begin
                            tmp    := ReadTokenPipe(stArg, ';');
                            if not FileExists(tmp) then
                            begin
                                PGIInfo ('Le fichier '+ tmp+ ' n''existe pas','');
                                Result := FALSE;
                                exit;
                            end;
                            Commande := 'COMMANDE';
                      end
                      else  tmp := Fichier;
                      FicIni                   := TIniFile.Create(tmp);
                      NomFichT                 := FicIni.ReadString (Commande, 'FICHIERTIERS', '');
                      Email                    := FicIni.ReadString (Commande, 'MAIL', '');
                      tmp                      := FicIni.ReadString (Commande, 'GESTIONETAB', '');
                      EtablissUnique           := (tmp='FALSE');
                      FICHENAME.Text           := FicIni.ReadString (Commande, 'NOMFICHIER', '');
                      tmp                      := FicIni.ReadString (Commande, 'TRANSFERTVERS', '');
                      Trans.Serie              := tmp;
                      if ((Trans.Serie <> 'SAGE') and (Trans.Serie <> 'IMMO')) and (Trans.Serie <> '')
                      and (Trans.Serie <> 'COMPTA') then
                      begin
                          TRANSFERTVERS.Itemindex  := TRANSFERTVERS.Values.IndexOf(tmp);
                          NATURETRANSFERT.Value    := FicIni.ReadString (Commande, 'NATURETRANSFERT', '');
                          Trans.Typearchive  := FicIni.ReadString (Commande, 'NATURETRANSFERT', '');
                      end;
                      FichierImp              := FicIni.ReadString (Commande, 'RAPPORT', '');

                      if (NATURETRANSFERT.Value = 'JRL')
                      or (NATURETRANSFERT.Value = 'BAL')
                      or (NATURETRANSFERT.Value = 'SYN')
                      or (Trans.Serie = 'SAGE') then
                      begin
                         //TYPEFORMAT.Value := 'STD';
                         TYPEFORMAT.Value       := FicIni.ReadString (Commande, 'FORMAT', '');
                         Trans.Exo              := FicIni.ReadString (Commande, 'EXERCICE', ''); // fiche 10615
                         DATEECR1.Text          := FicIni.ReadString (Commande, 'DATEECR1', '');
                         DATEECR2.Text          := FicIni.ReadString (Commande, 'DATEECR2', '');
                         tmp                    := FicIni.ReadString (Commande, 'TYPE', '');
                         if tmp <> '' then
                         begin
                            FTypeEcr.Text := tmp;
                            if (pos ('X', tmp) <> 0) then // pour la mode récupération typeecr X
                            begin
                             TYPEECRX := 'X';
                             ECRINTEGRE := FicIni.ReadString (Commande, 'TYPEECR', '');
                            end;
                         end;
                         tmp2 := FicIni.ReadString (Commande, 'JOURNAL', '');
                         if tmp2 <> '' then
                         begin
                            tmp                  := ReadTokenPipe(tmp2, '[');
                            Journaux.Text        := ReadTokenPipe(tmp2, ']');
                         end;
                         DATEECR.text            := FicIni.ReadString (Commande, 'DATEARRET', '');
                         ExclureEnreg            := FicIni.ReadString (Commande, 'EXCLURE', '');
                         IniEnabled ('LETTRAGE', AvecLettrage);
                         tmp2 := FicIni.ReadString (Commande, 'ETABLISSEMENT', '');
                         if tmp2 <> '' then
                         begin
                            tmp                  := ReadTokenPipe(tmp2, '[');
                            FETAB.Text           := ReadTokenPipe(tmp2, ']');
                         end;
                          HISTOSYNCHRO.Value      := FicIni.ReadString (Commande, 'HISTOSYNCHRO', '');
                      end;
                      IniEnabled ('ZIPPE', ZIPPE);

                      // ajout me pour GED
                      IniEnabled ('GED', BGed);

                      // pour GI ligne de commande
                      St := FicIni.ReadString (Commande, 'EXPORTE', '');
                      if St <>'' then //pour les trois positions <> cbGrayed
                      BExport.Checked := (St = 'TRUE') or (St = 'X');

                      tmp                     := FicIni.ReadString (Commande, 'NATUREJRL', '');
                      if tmp <> '' then
                          TraiteListeJournal (tmp);

                      IniEnabled ('ANNULVALIDATION', ANNULVALIDATION);
                      IniEnabled ('AVECLETTRAGE', AvecLettrage);

                      // en format balance
                      if (NATURETRANSFERT.Value = 'BAL') then
                      begin
                            IniEnabled ('PARAMDOS', PARAMDOS);
                            if PARAMDOS.Checked  then
                            begin
                               tmp                 := FicIni.ReadString (Commande, 'EXPARAMETRE', '');
                               if tmp <> '' then
                               begin
                                   PGEN.Checked        :=  (pos ('GEN', tmp) <> 0);
                                   PTIERS.Checked      :=  (pos ('TIE', tmp) <> 0);
                                   PTiersautre.Checked :=  (pos ('TIEA', tmp) <> 0);
                                   PSectionana.Checked :=  (pos ('SEC', tmp) <> 0);
                                   PJRL.Checked        :=  (pos ('JRL', tmp) <> 0);
                                   TLibre.Checked      :=  (pos ('TL', tmp) <> 0);
                                   PVentil.Checked     :=  (pos ('VEN', tmp) <> 0);
                                   CPTEDEBUT.text      := FicIni.ReadString (Commande, 'CPTEDEBUT', '');
                                   CPTEFIN.text        := FicIni.ReadString (Commande, 'CPTEFIN', '');
                                   DATEMODIF1.text     := FicIni.ReadString (Commande, 'DATEMODIF1', stDate1900);
                                   DATEMODIF2.text     := FicIni.ReadString (Commande, 'DATEMODIF2', stDate2099);
                               end;
                            end;
                      end
                      else
                      begin
                                tmp                 := FicIni.ReadString (Commande, 'EXPARAMETRE', '');
                                if tmp <> '' then
                                begin
                                   TLibre.Checked   :=  (pos ('TL', tmp) <> 0);
                                   PVentil.Checked  :=  (pos ('VEN', tmp) <> 0);
                                end;
                      end;

                      if ZIPPE.Checked then
                           ListeFichierJoint  := FicIni.ReadString (Commande, 'LISTEFICHIER', '');

                     tmp := FicIni.ReadString (Commande, 'WHEREAVANCE', '');
                     if tmp <> '' then WhereAvancee := tmp;

                      FicIni.free;
            end
            else
            begin
                      tmp := ReadTokenPipe(stArg, ';');
                      tmp := ReadTokenPipe(stArg, ';');
                      tmp := ReadTokenPipe(stArg, ';');
                      Email := ReadTokenPipe(stArg, ';');

                      NomFichT := ReadTokenPipe(stArg, ';');
                      tmp := ReadTokenPipe(stArg, ';');
                      // pour mettre les ecritures sur le même établissement
                      EtablissUnique := (tmp = '-');
                      tmp := ReadTokenPipe(stArg, ';');
                      //TRANSFERTVERS.Itemindex := TRANSFERTVERS.items.IndexOf(tmp);
                      Trans.Serie              := tmp;
                      // ajout me pour la purge si dossier netexpert fiche 17046
                      if Trans.serie <> '' then
                      TRANSFERTVERS.Itemindex  := TRANSFERTVERS.Values.IndexOf(tmp);

                      NATURETRANSFERT.Value := ReadTokenPipe(stArg, ';');   // nom du fichier

                      TYPEFORMAT.Value := 'ETE';
                      FICHENAME.Text   :=  ReadTokenPipe(stArg, ';');

                      if (NATURETRANSFERT.Value = 'JRL')
                      or (NATURETRANSFERT.Value = 'BAL')
                      or (NATURETRANSFERT.Value = 'SYN') then
                      begin
                         TYPEFORMAT.Value := 'STD';
                         EXERCICE.value:= ReadTokenPipe(stArg, ';');
                         DATEECR1.Text :=  ReadTokenPipe(stArg, ';');
                         DATEECR2.Text :=  ReadTokenPipe(stArg, ';');
                         if pos ('TYPE=', stArg) <> 0 then
                         begin
                              tmp := ReadTokenPipe(stArg, ']');
                              FTypeEcr.Text := tmp;
                              tmp := ReadTokenPipe(stArg, ';');
                         end
                         else
                              tmp := ReadTokenPipe(stArg, ';');

                         if pos ('JOURNAL=', stArg) <> 0 then
                         begin
                              tmp := ReadTokenPipe(stArg, '[');
                              Journaux.Text := ReadTokenPipe(stArg, ']');
                              tmp := ReadTokenPipe(stArg, ';');
                         end
                         else
                              tmp := ReadTokenPipe(stArg, ';');
                         if stArg <> '' then
                         begin
                              if pos (';', stArg) <> 0 then
                                 FichierImp := ReadTokenPipe(stArg, ';')
                              else
                                 FichierImp := stArg;
                         end;
                         if pos ('DATEARRET=', stArg) <> 0 then
                         begin
                              tmp := Copy(stArg, (pos('DATEARRET=', stArg))+10, Length(stArg));
                              if pos (';', stArg) <> 0 then
                                 DATEECR.text := ReadTokenPipe(tmp, ';')
                              else
                                 DATEECR.text := tmp;
                         end;
                      end;
            end;
            if (DATEECR.Text = '  /  /    ') then  DATEECR.Text := stDate1900;
            if (DATEECR.Text = stDate1900) and (NATURETRANSFERT.Value = 'SYN') then
               DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), now);

            if DATEECR1.Text = '  /  /    ' then
              DATEECR1.Text := stDate1900 ;
            if DATEECR2.Text = '  /  /    ' then
              DATEECR2.TExt := stDate2099;
(*   fiche 10615          if (DATEECR2.TExt <> stDate2099)  then
            begin
                QuelExo := QUELEXODTBUD (StrToDate(DATEECR2.Text));
                if QuelExo <> EXERCICE.value then
                EXERCICE.value := QuelExo;
            end
            else
                EXERCICE.value := '';
*)
            if (NomFichT <> '') and (FileExists (NomFichT)) then
            begin
                   if TOBTiers <> nil then TOBTiers.free;
                   TOBTiers := Nil ;
                   TOBTiers :=Tob.create('',Nil,-1) ;
                   if not TobLoadFromFile(NomFichT,nil,TOBTiers) then
                   begin
                        TOBTiers := Nil ;
                        TOBTiers.free;
                   end;
            end;
            if Email <> '' then
            begin
                 ZIPPE.Checked := TRUE;
                 if (pos('@', Email) = 0) then
                 begin
                  Email := '';  FMail.Checked := FALSE;
                 end
                 else
                 FMail.Checked := TRUE;
            end;
            stArg := stArgsav;
end;

Function TFAssistCom.AfficheListeComExport(Chaine: string; Listecom : TListBox) : Boolean;
var
Ch : string;
begin
  Result := TRUE;
  if ListeCom = nil then
  ListeCom := LISTEEXPORT;
  Ch := TraduireMemoire(Chaine);
  Listecom.Items.add (Ch);
  Listecom.ItemIndex := Listecom.Items.Count-1;
  if (WindowState  = wsMinimized) then Result := AfficheProgressbar(Ch);
end;

{$IFDEF EAGLCLIENT}
procedure TFAssistCom.ComsxOnCallBack(Sender: TObject) ;
var
  lemsg, param, msg,table: string ;
begin
if (Sender is Tob) then
  begin
        param:=Tob(Sender).getValue('Param') ;
        msg:=READTOKENPipe(param,'|') ;
        table:= READTOKENPipe(param,'|');
        table:=READTOKENPipe(param,'|') ;

        if (Table<>'')  then
        begin
          LeMsg:='Traitement en cours ...' ;
        end else
        begin
          LeMsg:=msg ;
        end ;
        if (LeMsg<>'') then AfficheListeComExport(LeMsg, LISTEEXPORT);
  end ;

end ;
{$ENDIF EAGLCLIENT}

{$IFNDEF EAGLCLIENT}
procedure TFAssistCom.FileClickzip(FFile : string=''; Directory : Boolean=FALSE);
var
  FileName       : String ;
  Commentaire    : String ;
  TheToz         : TOZ;
  Password       : string;
  Filearchive    : string;
  FFi            : string;
begin
  Password := '';
    // Récupération du nom du fichier a insérer
    //
  FileName := FICHENAME.Text;
  Filearchive := ReadTokenPipe (Filename, '.');
  Filearchive := Filearchive + '.ZIP';


  TheToz := TOZ.Create ;
  try
    if TheToz.OpenZipFile ( Filearchive, moCreate ) then
    begin
        if TheToz.OpenSession ( osAdd ) then
        begin
         if not Directory then
         begin
            if FFile = '' then
               FFi := FICHENAME.Text
            else
               FFi :=  FFile;

            if TheToz.ProcessFile ( FFi, Commentaire ) then
              begin
                   TheToz.CloseSession ;
                   if FFile = '' then DeleteFile(FicheName.Text);
              end
              else
              begin
              HShowMessage ( '0;Erreur;Soit le fichier : ' + ExtractFileName ( FileName ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '' ) ;
              TheToz.CancelSession ;
              end ;
         end
         else
         begin
{$IFDEF SCANGED}
              FFi := FFile;
              if Not ZippeDirectory(TheToz, FFI, '*.*', True, False) then
                 HShowMessage ( 'Erreur;compression répertoire dat, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '')
              else TheToz.CloseSession ;
{$ENDIF}
         end;
        end
        else
        HShowMessage ( '0;Erreur;Soit le fichier : ' + ExtractFileName ( FileName ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '' ) ;
    end
    else
    begin
      HShowMessage ( '0;Erreur;Erreur création du fichier archive : ' + 'archive.zip' + ' impossible;E;O;O;O', '', '' ) ;
      Exit ;
    end ;
 EXCEPT
    On E: Exception do
      begin
      ShowMessage ( 'TozError : ' + E.Message ) ;
      end ;
 END ;
 TheToz.Free;
 FICHENAME.Text := Filearchive;

end;
{$ENDIF EAGLCLIENT}

procedure TFAssistCom.RemplirParam (var Fichier,Commande : String);
var
Q1                : TQuery;
ATob              : TOB;
Serie             : string;
begin
  inherited;

  Label7.Caption := TraduireMemoire (Label7.Caption);
  TYPEFORMAT.ItemIndex := 1;
  NATURETRANSFERT.ItemIndex := 0;
  PSuiv := 1;
  lEtap.Caption := Msg.Mess[0] + ' ' + IntToStr(PSuiv);
  TraiteListeJournal ('');
  GRPBAL.Enabled  := FALSE;
  GRTYPECR.Enabled := FALSE;
  FTypeEcr.Enabled := FALSE;
  TFTypeEcr.Enabled := FALSE;

  PARAMSISCOII.Enabled := FALSE;

  TRANSFERTVERS.Itemindex := TRANSFERTVERS.Values.IndexOf(VH^.CPLienGamme);
  DATEECR.Text :=  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), now);
  if stArg = '' then  // par ligne de commande
  begin
       if TRANSFERTVERS.Itemindex = -1 then TRANSFERTVERS.Itemindex := 1;
       ChargeNatureTransfert(TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex]);
  end;
//  PPARAM.Enabled     :=  PARAMDOS.checked;
  PGEN.Enabled       :=  PARAMDOS.checked;
  PTIERS.Enabled     :=  PARAMDOS.checked;
  PTiersautre.Enabled:=  PARAMDOS.checked;
  PSectionana.Enabled:= PARAMDOS.checked;
  PJRL.Enabled       := PARAMDOS.checked;
//  TLibre.Enabled     := PARAMDOS.checked;

  PPARAMGENE.Enabled := not PARAMDOS.checked;

  CarFou := '0'; CarCli := '9';
  TobCol := TOB.Create('', nil, -1);
  Q1 := OpenSql ('SELECT * from CORRESP Where CR_TYPE="SIS"', TRUE);
  if Q1.EOF then
  begin
            COLLECTIFS.Value := 'F';
            CARAUX.value := '0';
  end
  else
  begin
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '40' then
                COLLECTIFS.Value := 'F';
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '41' then
                COLLECTIFS.Value := 'C';
           COLLECT.Text :=  AGauche(Q1.FindField ('CR_CORRESP').asstring, 10,'0');
           CARAUX.value := Copy (Q1.FindField ('CR_LIBELLE').asstring,1,1);
           AUX1.Text := AGauche(Q1.FindField ('CR_LIBELLE').asstring, 10,'0');
           AUX2.Text := AGauche(Q1.FindField ('CR_ABREGE').asstring, 10,'0');
  end;
  while not Q1.EOF do
  begin
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '40' then
                CarFou := copy (Q1.FindField ('CR_LIBELLE').asstring,0,1);
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '41' then
                CarCli := copy (Q1.FindField ('CR_LIBELLE').asstring,0,1);
           ATob := TOB.Create('CORRESP',TobCol,-1);
           ATob.PutValue('CR_CORRESP', AGauche(Q1.FindField ('CR_CORRESP').asstring, 10,'0'));
           ATob.PutValue('CR_LIBELLE', AGauche(Q1.FindField ('CR_LIBELLE').asstring, 10,'0'));
           ATob.PutValue('CR_ABREGE', AGauche(Q1.FindField ('CR_ABREGE').asstring, 10,'0'));
           Q1.next;
  end;
  ferme (Q1);
  FGrille := THGRID;
  TobCol.PutGridDetail(FGrille,False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
  COLLECT.EditMask := '';
  COLLECT.Text := '40';
  COLLECT.maxlength := VH^.Cpta[fbGene].Lg;
  AUX1.EditMask := '';
  AUX1.Text := '0';
  AUX1.maxlength := VH^.Cpta[fbGene].Lg;
  AUX2.EditMask := '';
  AUX2.Text := '0';
  AUX2.maxlength := VH^.Cpta[fbGene].Lg;
  TOBTiers := Nil ;
  FTimer.Enabled := FALSE;
  ExclureEnreg := '';
  OkExport := (stArg <> '');
  stArgsav := stArg;
  EtablissUnique := FALSE;
  if stArg <> '' then  // par ligne de commande
  begin
       if not RenseigneArg(Fichier, Commande) then exit;
  end
  else
  begin
       if V_PGI.LaSerie = S7 then serie := 'S5';
       if V_PGI.LaSerie = S5 then serie := 'S5';
       if V_PGI.LaSerie = S3 then serie := 'S3';
       if V_PGI.LaSerie = S1 then serie := 'S1';

       FCorpsMail.clear;
       FCorpsMail.Lines.add (TraduireMemoire('Veuillez trouver ci-joint le fichier correspondant au dossier '));
       FCorpsMail.Lines.add (V_PGI.NoDossier  + TraduireMemoire(' géré par le logiciel ')+ serie);
       FCorpsMail.Lines.add ('');
       FCorpsMail.Lines.add (TraduireMemoire('Merci de l''intégrer dès réception'));
       FEmail.Enabled := not(FFile.Checked);
       Label2.Enabled := not(FFile.Checked);
       FHigh.Enabled  := not(FFile.Checked);
       FCorpsMail.Enabled := not(FFile.Checked);
       Label12.Enabled := not(FFile.Checked);
       Label11.Enabled := not(FFile.Checked);
  end;
end;

{$IFNDEF EAGLCLIENT}
function TFAssistCom.NetExpertEnvoi : Boolean;
var
TNetEnvoi, TA, OutTOB       : TOB;
Filearchive, Ext,Info       : string;
i                           : integer;
formatSq                    : string;
begin
    // Result := FALSE;
//$|0||
//=|0||
//$|1||DOMAINE|CLIID |CLINOM|FICNOM |FICMSQ|FICTYPE|DATEARRET|FICSEQ|CREDAT|FICCTRL|DESCRIPTIF|
//=|1||COMPTA |10001|RUBI|PG10001_002|TRA|SYN|2004-04-21|002|2004-04-30 02:10:12|-|descriptif du fichier|

    TNetEnvoi := TOB.Create('', nil, -1);
    TA := TOB.Create ('',TNetEnvoi,-1);
    TA.AddChampSupValeur('DOMAINE', 'COMPTA');
    TA.AddChampSupValeur('CLIID', V_PGI.NoDossier);
    TA.AddChampSupValeur('CLINOM',V_PGI.NomSociete);
    Ext := ExtractFileName(FICHENAME.Text);
    Filearchive := ReadTokenPipe (Ext, '.');
    TA.AddChampSupValeur('FICNOM', Filearchive);
    Ext := ExtractFileExt(FICHENAME.Text);
    TA.AddChampSupValeur('FICMSQ', Copy(Ext, 2, length(Ext)));
    TA.AddChampSupValeur('FICTYPE', Trans.Typearchive);
//GP    TA.AddChampSupValeur('DATEARRET', FormatDateTime('yyyy-mm-dd', StrToDate (DATEECR.Text)));
    TA.AddChampSupValeur('DATEARRET', FormatDateTime(Traduitdateformat('yyyy-mm-dd'), StrToDate (DATEECR.Text)));
    formatSq := '%.03d';
    if NoSeqNet >= 999 then  formatSq := '%.04d';
    if NoSeqNet >= 9999 then  formatSq := '%.05d';
    TA.AddChampSupValeur('FICSEQ', Format (formatSq, [NoSeqNet]));
//GP    TA.AddChampSupValeur('CREDAT', FormatDateTime('yyyy-mm-dd hh:nn',NowH));
    TA.AddChampSupValeur('CREDAT', FormatDateTime(Traduitdateformat('yyyy-mm-dd hh:nn'),NowH));
    TA.AddChampSupValeur('FICCTRL', '-');
    if Trans.Typearchive = 'DOS' then
       Info := 'Dossier complet envoyé'
    else
       Info := 'Dossier : '+ V_PGI.NoDossier + ' synchronisé jusqu''au : ' +  FormatDateTime(Traduitdateformat('dd/mm/yyyy'), StrToDate (DATEECR.Text));
    TA.AddChampSupValeur('DESCRIPTIF', Info);
    // OutTOB := TOB.Create('', nil, -1);
    OutTOB := NEEnvoi (TNetEnvoi);
    ErreurNetexpert := FALSE;
    if OutTOB.GetValue ('ERROR') = '0' then
    begin
         for i:= 0 to OutTOB.detail.count-1 do
         begin
              if OutTOB.detail[i].GetValue('ERROR') <> '0' then
              begin
                   ErreurNetexpert := TRUE;
                   InfoNet := OutTOB.detail[i].GetValue('ERRORLIB');
                   PGIInfo (InfoNet);
                   DeleteFile(FicheName.Text);
                   break;
              end;
         end;
         if not ErreurNetexpert then
         begin
                  DeleteFile(FicheName.Text);
                  inc(NoSeqNet);
                  ExecuteSQL('UPDATE DOSSIER SET DOS_NECPSEQ=' +IntToStr(NoSeqNet)+
                  ' Where DOS_NODOSSIER="'+V_PGI.NoDossier+'"');
         end;
    end
    else
    begin
                  ErreurNetexpert := TRUE;
                  DeleteFile(FicheName.Text);
                  InfoNet := OutTOB.GetValue ('ERRORLIB');
                  PGIBox (InfoNet);
    end;
    Result := (not ErreurNetexpert);
    TNetEnvoi.Free;
end;
{$ENDIF}

Procedure TFAssistCom.ChangeCaption;
begin

CPTEDEBUT.Enabled   := TRUE;
CPTEFIN.Enabled     := TRUE;
TCPTEDEBUT.Enabled  := TRUE;
TCPTEFIN.Enabled    := TRUE;
GComptes.Enabled    := TRUE;
TNature.Enabled     := TRUE;
HNature.Enabled     := TRUE;
TNature.value       := '';

if PGen.Checked and (not PTIERS.Checked) and (not PSectionana.Checked )
and (not PJRL.Checked) then
begin
    GComptes.Caption := 'Comptes';
    TNature.DataType := 'TTNATGENE';
end
else
if (not PGen.Checked) and (PTIERS.Checked) and (not PSectionana.Checked )
and (not PJRL.Checked) then
begin
    GComptes.Caption := 'Tiers';
    TNature.DataType := 'TTNATTIERS';
end
else
if (not PGen.Checked) and (not PTIERS.Checked) and (PSectionana.Checked )
and (not PJRL.Checked) then
    GComptes.Caption := 'Sections'
else
if (not PGen.Checked) and (not PTIERS.Checked) and (not PSectionana.Checked )
and (PJRL.Checked) then
begin
    GComptes.Caption := 'Journaux';
    TNature.DataType := 'TTNATJAL';
end
else
begin
          CPTEDEBUT.Enabled   := FALSE;
          CPTEFIN.Enabled     := FALSE;
          TCPTEDEBUT.Enabled  := FALSE;
          TCPTEFIN.Enabled    := FALSE;
          GComptes.Enabled    := FALSE;
          TNature.Enabled     := FALSE;
          HNature.Enabled     := FALSE;

end;

end;

procedure TFAssistCom.PGENClick(Sender: TObject);
begin
  inherited;
  ChangeCaption;
end;

procedure TFAssistCom.SuppCarAuxiClick(Sender: TObject);
begin
  inherited;
  Trans.SuppCarAux := SuppCarAuxi.Checked;
end;


procedure TFAssistCom.PTiersautreClick(Sender: TObject);
begin
  inherited;
  if PTiersautre.Checked and not PTIERS.Checked then
     PTIERS.Checked := TRUE;
end;


procedure TFAssistCom.FormCreate(Sender: TObject);
var
TCF               : TControlFiltre;
begin
  inherited;

       TCF.Filtre   := BFiltre;
       TCF.Filtres  := FFiltres;
       TCF.PageCtrl := P;
       TCF.PopupF   := POPF;
       ObjetFiltre := TObjFiltre.create(TCF, 'COMEXPORT');
       BFiltre.Enabled := FALSE;
       BINI.Enabled := FALSE;
       CbLib.Checked := FALSE;
       Exclure.Items.Add('Bons à payer');
       Exclure.Values.add('BAP');

    // fiche 10438
    if (ctxPCL in V_PGI.PGIContexte) then GESTIONETAB.visible := FALSE;

end;

procedure COMDExoToDates ( Exo : String3 ; Var D1,D2 : TDateTime ) ;
Var Q     : TQuery;
begin
if EXO='' then Exit ;
D1:=Date ; D2:=Date ;
if EXO=VH^.Precedent.Code then begin D1:=VH^.Precedent.Deb ; D2:=VH^.Precedent.Fin ; end else
if EXO=VH^.EnCours.Code then begin D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ; end else
if EXO=VH^.Suivant.Code then begin D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; end else
   begin
   Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE) ;
   if Not Q.EOF then
      begin
      D1:=Q.FindField('EX_DATEDEBUT').asDateTime ; D2:=Q.FindField('EX_DATEFIN').asDateTime ;
      end;
   Ferme(Q) ;
   end;
end;

procedure TFAssistCom.GESTIONETABClick(Sender: TObject);
begin
  inherited;
   EtablissUnique := (not GESTIONETAB.checked);
   FETAB.Enabled := (GESTIONETAB.checked); // Fiche 10427
   if FETAB.Enabled = FALSE then FETAB.text := '';
end;

procedure TFAssistCom.BIniClick(Sender: TObject);
var
FicIni                   : TIniFile;
Rep,ParamExp             : string;
Fbat                     : TextFile;
Commande,CurrentFile     : string;
Currentbat,FF            : string;
Dossier, St              : string;
ii                       : integer;
begin
  inherited;
      Rep := ExtractFileDir (FICHENAME.Text);
      SetCurrentDirectory(PChar(Rep));
       with Topendialog.create(Self) do
       begin
         FileName := 'COMEXPORT.INI';
         Filter := 'Fichiers texte (*.INI)|*.INI|Tous les fichiers (*.*)|*.*';
         FilterIndex := 1;
         if Execute then
           CurrentFile := FileName
         else
            CurrentFile := '';
         Free;
       end;
      if  CurrentFile = '' then exit
      else
      Rep := ExtractFileDir (CurrentFile);
      Currentbat := CurrentFile;
      FF := ReadTokenPipe (Currentbat, '.');;
      Currentbat := FF +'.bat';


      if FileExists (CurrentFile) then DeleteFile( PChar(CurrentFile) );
      if FileExists (Currentbat) then DeleteFile( PChar(Currentbat) );

      FicIni        := TIniFile.Create(CurrentFile);
      FicIni.WriteString ('COMMANDE', 'NOMFICHIER', FICHENAME.Text);
      FicIni.WriteString ('COMMANDE', 'NATURETRANSFERT', NATURETRANSFERT.Value );
      FicIni.WriteString ('COMMANDE', 'TRANSFERTVERS', TRANSFERTVERS.Value );
      if FEMail.text <> '' then
      FicIni.WriteString ('COMMANDE', 'MAIL', FEMail.text );
      if not GESTIONETAB.checked then
           FicIni.WriteString ('COMMANDE', 'GESTIONETAB',  'FALSE' );

      FicIni.WriteString ('COMMANDE', 'FORMAT', TYPEFORMAT.Value);
      FicIni.WriteString ('COMMANDE', 'EXERCICE', EXERCICE.value);
      FicIni.WriteString ('COMMANDE', 'DATEECR1',  DATEECR1.Text);
      FicIni.WriteString ('COMMANDE', 'DATEECR2',  DATEECR2.Text);
//GP      if (JOURNAUX.Text <> '') and  (JOURNAUX.Text <> '<<Tous>>') then
      if (JOURNAUX.Text <> '') and  (JOURNAUX.Text <> Traduirememoire('<<Tous>>')) then
      FicIni.WriteString ('COMMANDE', 'JOURNAL', '['+JOURNAUX.Text+']');
      if DATEECR.Text <> stDate1900 then
         FicIni.WriteString ('COMMANDE', 'DATEARRET', DATEECR.text);

      if not PEXERCICE.Checked then
         FicIni.WriteString ('COMMANDE', 'EXCLURE', Exclure.TEXT +';EXO' )
      else
         FicIni.WriteString ('COMMANDE', 'EXCLURE', Exclure.TEXT );

      if AvecLettrage.Checked then
         FicIni.WriteString ('COMMANDE', 'LETTRAGE', 'TRUE' );

      if FETAB.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'ETABLISSEMENT', '['+FETAB.Text+']'  );
      if HISTOSYNCHRO.value <> '' then
         FicIni.WriteString ('COMMANDE', 'HISTOSYNCHRO', HISTOSYNCHRO.value );
      if ZIPPE.Checked then
         FicIni.WriteString ('COMMANDE', 'ZIPPE', 'TRUE' );
      if BGED.Checked then
         FicIni.WriteString ('COMMANDE', 'GED', 'TRUE' );
      if BExport.Checked then
         FicIni.WriteString ('COMMANDE', 'EXPORTE', 'TRUE' );
//GP      if (NATUREJRL.text <> '') and  (NATUREJRL.Text <> '<<Tous>>') then
      if (NATUREJRL.text <> '') and  (NATUREJRL.Text <> Traduirememoire('<<Tous>>')) then
         FicIni.WriteString ('COMMANDE', 'NATUREJRL', NATUREJRL.Value );   // Fiche 10615

      if ANNULVALIDATION.Checked then
         FicIni.WriteString ('COMMANDE', 'ANNULVALIDATION', 'TRUE');
      if BEXPORT.Checked then
         FicIni.WriteString ('COMMANDE', 'EXPORTE', 'TRUE' );
      if AvecLettrage.Checked then
         FicIni.WriteString ('COMMANDE', 'AVECLETTRAGE', 'TRUE');
      if PARAMDOS.Checked then
         FicIni.WriteString ('COMMANDE', 'PARAMDOS', 'TRUE' );
      if  (ZV1.Text <> '') and ((NATURETRANSFERT.Value = 'JRL') or (NATURETRANSFERT.Value = 'BAL')) then
      begin
            WhereAvancee := DecodeWhereSup;
           FicIni.WriteString ('COMMANDE', 'WHEREAVANCE', WhereAvancee );
      end;
      if (NATURETRANSFERT.Value = 'BAL') then
      begin
           if PARAMDOS.Checked then
             FicIni.WriteString ('COMMANDE', 'PARAMDOS', 'TRUE' );
           if PGEN.Checked  then ParamExp := 'GEN';
           if PTIERS.Checked then ParamExp := ParamExp + ',TIE';
           if PTiersautre.Checked then ParamExp := ParamExp + ',TIEA';
           if PSectionana.Checked then  ParamExp := ParamExp + ',SEC';
           if PJRL.Checked then  ParamExp := ParamExp + ',JRL';
           if TLibre.Checked then ParamExp := ParamExp + ',TL';
           if PVentil.Checked then ParamExp := ParamExp + ',VEN';

           ParamExp := '['+ParamExp +']';
           FicIni.WriteString ('COMMANDE', 'EXPARAMETRE', ParamExp );
           if CPTEDEBUT.text <> '' then
              FicIni.WriteString ('COMMANDE', 'CPTEDEBUT', CPTEDEBUT.text );
           if CPTEFIN.text <> '' then
              FicIni.WriteString ('COMMANDE', 'CPTEFIN', CPTEFIN.text );
           if DATEMODIF1.text <> stDate1900 then
              FicIni.WriteString ('COMMANDE', 'DATEMODIF1', DATEMODIF1.text );
           if DATEMODIF2.text <> stDate1900 then
              FicIni.WriteString ('COMMANDE', 'CPTEFIN', DATEMODIF2.text );
      end
      else
      begin
           if PVentil.Checked and TLibre.Checked then
              FicIni.WriteString ('COMMANDE', 'EXPARAMETRE', '[TL, VEN]')
           else
           if TLibre.Checked then
              FicIni.WriteString ('COMMANDE', 'EXPARAMETRE', '[TL]')
           else
           if PVentil.Checked then
              FicIni.WriteString ('COMMANDE', 'EXPARAMETRE', '[VEN]');
      end;
      FicIni.free;                          

      // fichier .bat
    AssignFile(Fbat, Currentbat);
    Rewrite(Fbat) ;
    for ii :=1 to ParamCount do
    begin
        St:=ParamStr(ii) ;
        Dossier :=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
        if Dossier='/DOSSIER'  then begin Dossier := St; break; end;
    end;
   // Fiche 10603
    if ((V_PGI.ModePCL='1')) then
      Commande := '"' + Application.ExeName +'" /USER='+V_PGI.UserLogin+ ' /PASSWORD='+V_PGI.PassWord+' /DOSSIER='+Dossier+ ' "/INI='+CurrentFile+';EXPORT;Minimized"'
    else
        Commande := '"'+ Application.ExeName +'" /USER='+V_PGI.UserLogin+ ' /PASSWORD='+V_PGI.PassWord+' /DOSSIER='+V_PGI.Currentalias+ ' "/INI='+CurrentFile+';EXPORT;Minimized"';
    Writeln(Fbat, Commande) ;
    CloseFile(Fbat);
    PGIInfo ('La génération du fichier de commande est terminée');

end;

// fiche 10431
function TFAssistCom.Rendw(Champ :string) : string;
 var
    x, y, St : string;
begin
          St := ''; y :=TNature.value;
          x := ReadTokenSt (y);
          if x <> '' then  St := Champ+'="' +x+ '"';
          while x <> '' do
          begin
                 x := ReadTokenSt (y);
                 if x <> '' then
                 St := St+ ' OR '+ Champ +'="' +x+ '"';
          end;
          Result := St;
end;

procedure TFAssistCom.CPTEDEBUTElipsisClick(Sender: TObject);
var
WhereNature : string;
begin
  inherited;
if PGen.Checked then
begin
     if TNature.value <> '' then
          WhereNature := Rendw('G_NATUREGENE')
     else
          WhereNature := '';
     if PECRBUD.Checked   then
          LookupList(TControl(Sender),TraduireMemoire('Comptes'),'BUDGENE','BG_BUDGENE','BG_LIBELLE','','BG_BUDGENE', True,1)
     else
          LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',WhereNature,'G_GENERAL', True,1)  ;
end
else
if PTIERS.Checked then
begin
     if TNature.value <> '' then
          WhereNature := Rendw('T_NATUREAUXI')
     else
          WhereNature := '';

     if PECRBUD.Checked   then
     begin
          CPTEDEBUT.Enabled   := FALSE;
          CPTEFIN.Enabled     := FALSE;
          TCPTEDEBUT.Enabled  := FALSE;
          TCPTEFIN.Enabled    := FALSE;
          GComptes.Enabled    := FALSE;
          TNature.Enabled     := FALSE;
          HNature.Enabled     := FALSE;
     end
     else
          LookUpList(TControl(Sender), TraduireMemoire('Auxiliaire'), 'TIERS',
          'T_AUXILIAIRE', 'T_LIBELLE', WhereNature, 'T_AUXILIAIRE', True, 2);
end
else
if PSectionana.Checked then
begin
     if PECRBUD.Checked   then
          LookupList(TControl(Sender),TraduireMemoire('Sections budgétaires'),'BUDSECT','BS_BUDSECT','BS_LIBELLE','','BS_BUDSECT', True,1)
     else
          LookupList(TControl(Sender),TraduireMemoire('Sections'),'SECTION','S_SECTION','S_LIBELLE','','S_SECTION', True,1)  ;
end
else
if PJRL.Checked then
begin
     if TNature.value <> '' then
          WhereNature := Rendw('J_NATUREJAL')
     else
          WhereNature := '';
     if PECRBUD.Checked   then
          LookupList(TControl(Sender),TraduireMemoire('Journaux budgétaires'),'BUDJAL','BJ_BUDJAL','BJ_LIBELLE','','BJ_BUDJAL', True,1)
     else
          LookupList(TControl(Sender),TraduireMemoire('Journaux'),'JOURNAL','J_JOURNAL','J_LIBELLE',WhereNature,'J_JOURNAL', True,1)  ;
end;

end;

procedure TFAssistCom.DATARRETEBALExit(Sender: TObject);
begin
  inherited;
             OkEnter := FALSE;
end;

procedure TFAssistCom.DATARRETEBALChange(Sender: TObject);
var
D1 : TDateTime;
begin
  inherited;
  if not OkEnter then exit;
  if (stArg <> '') or (TRANSFERTVERS.Itemindex = -1) then exit;
  if (stArg = '') and (TRANSFERTVERS.Values[TRANSFERTVERS.Itemindex] = 'SISCO')then
  begin
       if (NATURETRANSFERT.Value <> 'BAL') then exit;
       if (DATEECR.Text = '  /  /    ')  or (DATARRETEBAL.Text = stDate1900)
       or (pos(' ',DATARRETEBAL.Text) <> 0) then exit;
       if TransIsValidDate(DATARRETEBAL.Text) then
       begin
             D1 := StrToDate(DATARRETEBAL.Text);
             if (D1 < strTodate(DATEECR1.text)) or
             (D1  > strTodate(DATEECR2.text)) then
             DATARRETEBAL.text := FormatDateTime(Traduitdateformat('dd/mm/yyyy'), VH^.Encours.Fin);
       end;
  end;
end;

procedure TFAssistCom.DATARRETEBALEnter(Sender: TObject);
begin
  inherited;
             OkEnter := TRUE;
end;

procedure TFAssistCom.PVentilClick(Sender: TObject);
begin
  inherited;
if PVentil.checked then
   PSectionana.Checked := TRUE
else
   PSectionana.Checked := FALSE;

end;

procedure TFAssistCom.RenvoieCorpsMail (var  Hst : HTStringList);
var
i   : integer;
begin
         if Hst = nil then Hst := HTStringList.create;
         for i := 0 to FCorpsMail.Lines.Count-1 do
               Hst.Add(FCorpsMail.Lines[i]);
end;

Procedure AddZC(ZC : THValCombobox ; TDE : TDEChamp ; OnLib : Boolean) ;
BEGIN
If OnLib Then ZC.Items.Add(TDE.Libelle) Else ZC.Items.Add(TDE.Nom) ;
ZC.Values.Add(TDE.Nom) ;
END ;

procedure TFAssistCom.ChargeAvances(OnLib : Boolean) ;
Var i,j,k : Integer ;
    SST : Array[1..6] Of String ;
    TC : tControl ;
    ZC : THValComboBox ;
    St : String ;
BEGIN
St:='E' ;
i:=PrefixeToNum(St) ;
For k:=1 To 6 Do
BEGIN
  SST[k]:='' ;
  TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
  If TC<>NIL Then
    BEGIN
    ZC:=THValComboBox(TC) ; If ZC.ItemIndex>=0 Then SST[k]:=ZC.Value ; ZC.Items.Clear ; ZC.Values.Clear ;
    END ;
END ;
{$IFDEF EAGLCLIENT}
  if High(V_Pgi.DEChamps[i]) <= 0 then
    ChargeDeChamps(i, TableToPrefixe('ECRITURE'));
{$ENDIF EAGLCLIENT}

for j:=1 to High(V_PGI.DeChamps[i]) do
begin
   if (V_PGI.DEChamps[i,j].Nom <> 'E_QUALIFORIGINE') and  (V_PGI.DEChamps[i,j].Nom <> 'E_SOCIETE')
   and (V_PGI.DEChamps[i,j].Nom <> 'E_UTILISATEUR') then continue;
   if (V_PGI.DEChamps[i,j].Nom<>'') then
   BEGIN
        For k:=1 To 6 Do
          BEGIN
          TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
          If TC<>NIL Then
            BEGIN
            ZC:=THValComboBox(TC) ;
            AddZC(ZC,V_PGI.DEChamps[i,j],OnLib) ;
            END ;
          END ;
     END ;
end;
For k:=1 To 6 Do
  BEGIN
  TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
  If TC<>NIL Then
    BEGIN
    ZC:=THValComboBox(TC) ;
    If SST[k]<>'' Then ZC.Value:=SST[k] ;
    END ;
  END ;
END ;



procedure TFAssistCom.CBLibClick(Sender: TObject);
begin
  inherited;
ChargeAvances(CBLib.Checked) ;
end;


function TFAssistCom.DecodeWhereSup : string ;
Var St,St1 : String ;
    typ    : String ;
BEGIN
    St  := Trim(ZV1.Text) ; If St='' Then Exit ;
    St1 := ' AND ' + Z_C1.value;
    typ := ChampToType(Z_C1.value);
    if (Typ <>'INTEGER') and (Typ <> 'SMALLINT')and (Typ<>'DOUBLE') and (Typ<>'RATE') then
    Typ := '"'
    else Typ := '';
    Result := ST1 +ZO1.Value+typ+ St+typ;
END ;

{$IFDEF COMPTA}
procedure LancementEtat (CodeE, DateCDu, DateCAu, Code, FNomPdf : string); // Lancement des états
var TEtatC, LT     : TOB;
    lCEtatsChaines : TCEtatsChaines;
begin
  // Traitemet de la Compta : Création de l'objet des états chainés
  lCEtatsChaines := nil; TEtatC := nil;
  try
    TEtatC := TOB.Create('', nil, -1);
    LT := TOB.Create('ETAT', TEtatC, -1);
    LT.AddChampSupValeur('ID' , CodeE) ;
    LT.AddChampSupValeur('LIBELLEETAT' , '' ) ;
    LT.AddChampSupValeur('NBEXEMPLAIRE'   , 1 ) ;
    LT.AddChampSupValeur('FILTREUTILISE'  , '' ) ;
    lCEtatsChaines := TCEtatsChaines.Create ;
    // Utilisation du Format PDF pour l'impression
    begin
      lCEtatsChaines.CritEdtChaine.AuFormatPDF := True;
      lCEtatsChaines.CritEdtChaine.NomPDF      := FNomPdf;
      lCEtatsChaines.CritEdtChaine.MultiPdf    := False;
    end;

    // Utilisation des critères standards Compatibilité
    lCEtatsChaines.CritEdtChaine.UtiliseCritStd := True;
    lCEtatsChaines.CritEdtChaine.Exercice.Code  := CExerciceVersRelatif (Code);
    lCEtatsChaines.CritEdtChaine.Exercice.Deb   := StrToDate(DateCDu);
    lCEtatsChaines.CritEdtChaine.Exercice.Fin   := StrToDate(DateCAu);
    lCEtatsChaines.Execute( TEtatC ) ;

  finally
    lCEtatsChaines.Free;
    TEtatC.free;
  end;

end;

Procedure EnvoiExportParDate (Date1, Date2, Code, Nature : string);
var
Fichier, TempDir, Fileini : string;
IniFile                   : TIniFile;
Q                         : TQuery;
FileArchive               : string;
TJAL                      : TOB;
TheToz                    : TOZ;
OkArchive                 : Boolean;
FSaisieARC                : TFVierge;
Lb                        : THLabel;
Ed1                       : THCritMaskEdit;
Filereceive               : string;
begin
    OkArchive := FALSE;
    if  GetParamSocSecur('SO_PATHARCHIVE', '', TRUE) = '' then
    begin
        FSaisieARC := TFVierge.Create(nil);
        with FSaisieARC do
        begin
            Caption := 'Saisie répertoire d''archive';
            BorderStyle := bsDialog;
            Position := poScreenCenter;
            ClientWidth := 382;
            ClientHeight := 100;
            Lb :=  THLabel.create(nil);
            Lb.Height := 13;
            Lb.Left := 12;
            Lb.Top := 6;
            Lb.Width := length('Répertoire d''archive');

            Lb.caption := 'Répertoire d''archive';
            Lb.Parent := FSaisieARC;

            Ed1 :=  THCritMaskEdit.create(nil);
            Ed1.Name := 'REP';
            Ed1.tag := 100;
            Ed1.Height := 21;
            Ed1.Left := 12;
            Ed1.Top := 25;
            Ed1.text := '';
            Ed1.Width := 357;
            Ed1.Parent := FSaisieARC;
            Ed1.datatype := 'DIRECTORY';
            Ed1.ElipsisButton := TRUE;
            ShowModal;
            if (Ed1.Text <> '') and  DirectoryExists(Ed1.Text) then
              SetParamSoc('SO_PATHARCHIVE', Ed1.Text);
            Free;
        end;
    end;
    TempDir := TcbpPath.GetCegidUserTempPath +'\'+V_PGI.NoDossier;
    if not DirectoryExists(TempDir) then CreateDir(TempDir);
    Fileini := TempDir + '\COMSX.ini';
    Fichier := TempDir + '\PG'+V_PGI.NoDossier;
    IniFile := TIniFile.Create (Fileini);
    IniFile.WriteString ( 'COMMANDE' , 'NOMFICHIER', Fichier+'.TRA');
    IniFile.WriteString ( 'COMMANDE' , 'TRANSFERTVERS', 'S5');
    IniFile.WriteString ( 'COMMANDE' , 'NATURETRANSFERT', 'JRL');
    IniFile.WriteString ( 'COMMANDE' , 'FORMAT', 'ETE');
    IniFile.WriteString ( 'COMMANDE' , 'DATEECR1', Date1);
    IniFile.WriteString ( 'COMMANDE' , 'DATEECR2', Date2);
    IniFile.WriteString ( 'COMMANDE' , 'RAPPORT', Fichier+'.txt');
    IniFile.Free;
    ExportDonnees(Fileini+';EXPORT;Minimized',TRUE, 'S5', TRUE) ;
    TJAL := TOB.Create('', nil, -1);
    // sauvegarde du journal d'import
    Q := OpenSql ('SELECT * FROM CPJALIMPORT', TRUE);
    TJAL.LoadDetailDB('CPJALIMPORT', '', '', Q, TRUE, FALSE);
    TJAL.SaveToXmlFile(TempDir + '\CPJALIMPORT' + '.XML', TRUE);
    Ferme (Q);
    TJAL.free;

    // sauvegarde du Journal d'événement
    TJAL := TOB.Create('', nil, -1);
    Q := OpenSql ('SELECT * FROM JNALEVENT', TRUE);
    TJAL.LoadDetailDB('JNALEVENT', '', '', Q, TRUE, FALSE);
    TJAL.SaveToXmlFile(TempDir + '\JNALEVENT' + '.XML', TRUE);
    Ferme (Q);
    TJAL.Free;

    // sauvegarde du Journal de validation des écritures
    Q := OpenSql ('SELECT * FROM CPJALVALIDATION', TRUE);
    TJAL := TOB.Create('', nil, -1);
    TJAL.LoadDetailDB('CPJALVALIDATION', '', '', Q, TRUE, FALSE);
    TJAL.SaveToXmlFile(TempDir + '\CPJALVALIDATION' + '.XML', TRUE);
    Ferme (Q);
    TJAL.Free;

    // Lancement des états
    LancementEtat ('CP02', Date1, Date2, Code, TempDir + '\GrandLivre.pdf');
    LancementEtat ('CP06', Date1, Date2, Code, TempDir + '\JournalCentralisateur.pdf');
    LancementEtat ('CP05', Date1, Date2, Code, TempDir + '\Journaldesecritures.pdf');

   if (ctxPCL in V_PGI.PGIContexte) then
    FileArchive := TempDir + '\ARCHIVE' + V_PGI.NoDossier  + '_'
    + Nature + '_' + FormatDateTime(Traduitdateformat('ddmmyyyy'), StrToDate(Date1)) + '_' + FormatDateTime(Traduitdateformat('ddmmyyyy'), StrToDate(Date2))
    + '.ZIP'
  else
    FileArchive := TempDir + '\ARCHIVE' + GetParamsocSecur('SO_LIBELLE', '') + '_'
    + Nature + '_' + FormatDateTime(Traduitdateformat('ddmmyyyy'), StrToDate(Date1)) + '_' + FormatDateTime(Traduitdateformat('ddmmyyyy'), StrToDate(Date2))
    + '.ZIP';

    IniFile := TIniFile.Create (TempDir + '\ListeArchive'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt');
    IniFile.WriteString ( 'COMMANDE' , 'NOMFICHIER', ExtractFileName(FileArchive));
    IniFile.WriteString ( 'COMMANDE' , 'DATE', FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH));
    IniFile.WriteString ( 'COMMANDE' , 'UTILISATEUR', V_PGI.UserLogin);
    IniFile.WriteString ( 'COMMANDE' , 'Date de début', Date1);
    IniFile.WriteString ( 'COMMANDE' , 'Date de fin', Date2);
    IniFile.WriteString ( 'COMMANDE' , 'Fichiers', 'CPJALIMPORT.XML,JNALEVENT.XML,CPJALVALIDATION.XML,GrandLivre.pdf,JournalCentralisateur.pdf,Journaldesecritures.pdf,'+Fichier+'.TRA');
    IniFile.Free;

    TheToz := TOZ.Create ;
    try
      if TheToz.OpenZipFile (FileArchive, moCreate ) then
      begin
          if TheToz.OpenSession ( osAdd ) then
          begin
          {$IFDEF SCANGED}
               if Not ZippeDirectory(TheToz, TempDir, '*.*', True, False, '.INI') then
                 HShowMessage ( 'Erreur;compression répertoire dat, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '')
               else OkArchive := TRUE;
          {$ENDIF}
          end;
      end;
   except
      On E: Exception do
        begin
        ShowMessage ( 'TozError : ' + E.Message ) ;
        end ;
   end ;
   TheToz.CloseSession ;
   TheToz.Free;
   if OkArchive Then
   begin
      if GetParamSocSecur('SO_PATHARCHIVE', '', TRUE) <> ''  then
      begin
          Filereceive := GetParamSocSecur('SO_PATHARCHIVE', TempDir, TRUE) + '/' + ExtractFileName(FileArchive);
          if FileExists(Filereceive) then DeleteFile (Filereceive);
          CopyDirectory2 (TempDir, GetParamSocSecur('SO_PATHARCHIVE', TempDir, TRUE), '*.ZIP', FALSE, TRUE, TRUE);
          RemoveInDir2 (TempDir, FALSE, TRUE, TRUE);
      end;
   end;
end;
{$ENDIF COMPTA}


end.


