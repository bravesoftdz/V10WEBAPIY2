unit  UASSISTIMPORT;

interface

uses
  Windows, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, hmsgbox, StdCtrls, ComCtrls, ExtCtrls, HTB97, Hctrls,
  Mask, UTOB, Paramsoc,
  ControlParam, MesgErrCom,
  RecupUtil, CloPerio,
  Menus, Spin, HSysMenu, HPanel,ShellAPI,
{$IFDEF COMSX}
  ImpFic,
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  Fe_main,
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  UImportCom,
{$ELSE}
  MainEAgl,
  uLanceProcess,
  uHTTP, UWa,
{$ENDIF}

{$ifdef COMSX}
  UControlFicU,
{$endif}
  FichComm, HEnt1, Ent1, CORRESP_TOM, UtilTrans,
  UToz, ImprimeMaquette, printers, TImpFic,
  uLibStdCpta, Mailol, UlibExercice, uNEActions,
  Galoutil, Lookup, CLgCpte, Hcompte, CPTIERS_TOM,
  CPSECTION_TOM, CPJOURNAL_TOM, utilPGI, ImportConf,
  UObjFiltres, UListByUser, UTxml
  , ULibBonapayer
  , inifiles, CPGENERAUX_TOM;

const
  MaxEnreg = 100;

type
  TFAssistImport = class(TFAssist)
    Mail: TTabSheet;
    Label2: TLabel;
    Bevel5: TBevel;
    Label11: TLabel;
    Label12: TLabel;
    Label6: TLabel;
    Label3: TLabel;
    FEMail: THCritMaskEdit;
    FHigh: TCheckBox;
    FFile: TRadioButton;
    FMail: TRadioButton;
    FICHENAME: THCritMaskEdit;
    Resume: TTabSheet;
    Label13: TLabel;
    Bevel6: TBevel;
    Bevel3: TBevel;
    FLib1: TLabel;
    FVal1: TLabel;
    FLib3: TLabel;
    FVal3: TLabel;
    FLib4: TLabel;
    FVal4: TLabel;
    HLABEL2: THLabel;
    CP_STAT: THValComboBox;
    bClickVisu: TToolbarButton97;
    HLabel1: THLabel;
    TListeFichier: THLabel;
    FLISTEFICHIER: THMultiValComboBox;
    BNetExpert: TCheckBox;
    AVANCE: TTabSheet;
    HLabel3: THLabel;
    Bevel1: TBevel;
    POPZ: TPopupMenu;
    Generaux1: TMenuItem;
    Auxiliaires1: TMenuItem;
    Axeanalytique11: TMenuItem;
    Axeanalytique21: TMenuItem;
    Axeanalytique31: TMenuItem;
    Axeanalytique41: TMenuItem;
    Axeanalytique51: TMenuItem;
    ModedePaiement1: TMenuItem;
    Journaux1: TMenuItem;
    ParamtrageSISCOII1: TMenuItem;
    POPZ1: TPopupMenu;
    MenuItem1: TMenuItem;
    Zoomaux: TMenuItem;
    Zoomana: TMenuItem;
    Zoomsurlejournal1: TMenuItem;
    OPTIONAVANCE: TCheckBox;
    PIMPORT: TTabSheet;
    Label14: TLabel;
    Bevel7: TBevel;
    GroupBox2: TGroupBox;
    BInterdireCreatCompte: TCheckBox;
    INTERDIRECREAT: THMultiValComboBox;
    BGestiondoublon: TCheckBox;
    GroupBox3: TGroupBox;
    BRupturePiece: TCheckBox;
    BCalculPiece: TCheckBox;
    Zoomsurtablissement1: TMenuItem;
    GroupBox4: TGroupBox;
    BControlParam: TCheckBox;
    BControlEtab: TCheckBox;
    SuppComptable: TCheckBox;
    AVANCE2: TTabSheet;
    HLabel4: THLabel;
    Bevel2: TBevel;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    GComptes: TGroupBox;
    TCPTEDEBUT: THLabel;
    TCPTEFIN: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    GroupBox5: TGroupBox;
    HLabel8: THLabel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    HLabel11: THLabel;
    TRAxe5: THLabel;
    GroupBox6: TGroupBox;
    HLabel13: THLabel;
    HLabel14: THLabel;
    GroupBox7: TGroupBox;
    Label8: THLabel;
    Label9: THLabel;
    MainMenu1: TMainMenu;
    LISTEIMPORT: TListBox;
    MRDEFAUT: THValComboBox;
    REGDEFAUT: THValComboBox;
    GroupBox8: TGroupBox;
    BRejet: TCheckBox;
    BValider: TCheckBox;
    BCtrP: TCheckBox;
    BIntegre: TCheckBox;
    Etablissement1: TMenuItem;
    BEnvoiMail: TCheckBox;
    BINI: TToolbarButton97;
    BCTRLSOCIETE: TCheckBox;
    CPTEGENE: THCritMaskEdit;
    CPTEFOUR: THCritMaskEdit;
    CPTEDIVERS: THCritMaskEdit;
    CPTECLIENT: THCritMaskEdit;
    CPTESALAIRE: THCritMaskEdit;
    RAXE1: THCritMaskEdit;
    RAxe2: THCritMaskEdit;
    RAxe3: THCritMaskEdit;
    RAxe4: THCritMaskEdit;
    RAxe5: THCritMaskEdit;
    SUBCOLLCLI: THCritMaskEdit;
    SUBCOLLFOU: THCritMaskEdit;
    Label1: TLabel;
    FCORRESP: THMultiValComboBox;
    BCorresp: TToolbarButton97;
    HTYPEIMPORT: THLabel;
    TQUALPIECE: THValComboBox;
    Label7: TLabel;
    DecoupLon: TSpinEdit;
    Label4: TLabel;
    GroupBox9: TGroupBox;
    BMAJCPTE: TCheckBox;
    INTERDIRECREATTIERS: TCheckBox;
    POPZ2: TPopupMenu;
    BVOIR: TMenuItem;
    CTRPIECE: TMenuItem;
    GroupBox10: TGroupBox;
    BEcartChange: TCheckBox;
    LBBAP: THLabel;
    CBBAP: THValComboBox;
    BTPayeur: TCheckBox;
    TDATEECR1: THLabel;
    DATEECR1: THCritMaskEdit;
    LDATEECR1: THLabel;
    GroupBox11: TGroupBox;
    BRecalculSolde: TCheckBox;
    BCalcSoldeLigne: TCheckBox;
    BBlanc: TCheckBox;
    MODIFTIERS: THMultiValComboBox;
    FCorpsMail: THMemo;
    BMajDateDernMvt: TCheckBox;
    BDESEQUILIBRE: TCheckBox;
    CalculdateEche: TCheckBox;
    Label15: TLabel;
    Label5: TLabel;
    procedure bFinClick(Sender: TObject);
    procedure bClickVisuClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure BTWINZIPClick(Sender: TObject);
    procedure FMailClick(Sender: TObject);
    procedure FEMailExit(Sender: TObject);
    procedure FICHENAMEChange(Sender: TObject);
    procedure BNetExpertClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BCalculPieceClick(Sender: TObject);
    procedure BInterdireCreatCompteClick(Sender: TObject);
    procedure Generaux1Click(Sender: TObject);
    procedure Auxiliaires1Click(Sender: TObject);
    procedure Axeanalytique11Click(Sender: TObject);
    procedure Axeanalytique21Click(Sender: TObject);
    procedure Axeanalytique31Click(Sender: TObject);
    procedure Axeanalytique41Click(Sender: TObject);
    procedure Axeanalytique51Click(Sender: TObject);
    procedure ModedePaiement1Click(Sender: TObject);
    procedure Journaux1Click(Sender: TObject);
    procedure ParamtrageSISCOII1Click(Sender: TObject);
    procedure TQUALPIECEChange(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure ZoomauxClick(Sender: TObject);
    procedure ZoomanaClick(Sender: TObject);
    procedure LISTEIMPORTContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure LISTEIMPORTDblClick(Sender: TObject);
    procedure Zoomsurlejournal1Click(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure Zoomsurtablissement1Click(Sender: TObject);
    procedure OPTIONAVANCEClick(Sender: TObject);
    procedure Etablissement1Click(Sender: TObject);
    procedure BINIClick(Sender: TObject);
    procedure CTRPIECEClick(Sender: TObject);
    procedure DATEECR1Exit(Sender: TObject);
    procedure BBlancClick(Sender: TObject);
    procedure BMAJCPTEClick(Sender: TObject);
  private
    { Déclarations privées }
  OrigineFichier,NatureFichier      : string;
  OrigineEnvoie                     : string;
  DateArrete                        : TDateTime;
  OKexercice                        : boolean;
  Paramgene                         : boolean;
  Changementgene, Changementaux     : Boolean;
  OkImportS,Sansaux                 : Boolean;
  RepEncours                        : string;
  FichImp                           : string;
  TypeEcrR,TypeEcrI,Calculpiece     : string;
  AutoriseEcheck, OkCommit          : Boolean;
  RepriseComplement                 : string;
  InitGI,IsNetExpert, ScanBanck     : Boolean;
  FTimer                            : TTimer;
  OutTOBNet                         : TOB;
  QualifOrigine                     : string;
  stArg,FichierIni                  : string;
  TobFich                           : TOB;
  ObjetFiltre                       : TObjFiltre;
  AccesCommande                     : string;
  OptionEchange,ControlExercClos    : Boolean;
  {$IFDEF CERTIFNF}
  CpjCode                           : string;
  {$ENDIF}
  procedure OnMonTimer(Sender: TObject);
  procedure EnvoieImport(FichZipp : string='');
  function  EcritureTraceImport (OkFinImport : Boolean) : string;
  Function  EstUneLigneCom(St : String) : string;
  function  EnleveBlanc(Chaine: string; x, y: integer ; ValNumeric : String = ''): string;
  Function  Controlparamsociete  (SLect : string) : Boolean;
  Function  ControlParam5 (Ligne : string; Existeexerc : Boolean) : Boolean;
  function  COkExercice (SLect : string) : Boolean;
  Function  NbExoouv(SS : string; var Nb : integer): Boolean;
  Function  ControlEnreg (SS,Enreg : string; deb : integer; fin : integer): Boolean;
  Function  ControleAutreParam (Deb,SLect : string) : Boolean;
  Function  NombreExercice(WW : string; Max : Boolean=FALSE): integer;
  Function  ControleParamSoc(indice: integer) : Boolean;
  Function  CoherenceCompte (indice : integer; Lec : string) : Boolean;
  function  ImportComSx(FichierSortie: string; InfoImp : PtTInfoImport=nil): Boolean;
  Function  AfficheListeComImport(Chaine: string; Listecom : TListBox; pourcentage : double=0) : Boolean;
  function  InitTobParamImp (fichier : string) : TOB ;
  procedure RendreDispo(Ena : Boolean);
  procedure InitComImport ( Stt : string);
  procedure AffichageOptionImport;
  procedure RenvoieCorpsMail (var  Hst : HTStringList);
  {$IFDEF CERTIFNF}
  Procedure MajJournalImport (Filename : string);
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  procedure ComsxOnCallBack(Sender: TObject) ;
  {$ENDIF}

  public
    { Déclarations publiques }
    constructor create (stArgument, FileIni, Commande : string; OptE : Boolean); reintroduce; overload ; virtual ;

  end;

var
  FAssistImport: TFAssistImport;


//const  ElimChars =[',', ')','+', '-', '*', '/', '\', '''','=', '<', '>', '^', '%','.', ',', ';', ':', '|', '!', '?', '#','_', '$','&','µ','Â','Ö','É','Û', 'Ê', '§', ' '];


Function ImportDonnees(stArgument : string; Mode : Boolean=FALSE; OutN : TOB=nil; FileIni : string=''; Commande : string=''; OptE : Boolean=FALSE) : Boolean ;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  UProcGen;

Function ImportDonnees(stArgument : string; Mode : Boolean=FALSE; OutN : TOB=nil; FileIni : string=''; Commande : string=''; OptE : Boolean=FALSE) : Boolean;
var
lStSQl                     : string;
BEGIN
     Result := TRUE;
     if Not ExJaiLeDroitConcept(TConcept(ccimport),True) then
     begin
        Result := FALSE; exit;
     end;
     FAssistImport := TFAssistImport.Create(stArgument, FileIni, Commande, OptE);
     With FAssistImport do
     begin
         try
             // à porter sur V7XX
             OkCommit := FALSE; 
             if ExisteSQL ('SELECT 1 FROM COURRIER WHERE MG_TYPE=5555') then
             begin
               PGIINFO('Attention, un traitement Import est déjà encours');
               Result := FALSE;
             end
             else
             begin
                   lStSql := 'insert into courrier(mg_utilisateur,mg_combo,mg_type,mg_dejalu,mg_date,mg_averti,mg_expediteur) ' +
                   'values("' + W_W + '","IMPORT COMSX",5555,"-","' +
                   USTime(iDate1900) + '","-","' + V_PGI.User + '")';
                   try
                    ExecuteSQL(lStSQL);
                   except; end;

                   BeginTrans ;
                       stArg := stArgument;
                       OKexercice := TRUE;
                       FLISTEFICHIER.Value := stArg;
                       if (OutN <> nil ) then
                       begin
                            BNetExpert.Checked := TRUE;
                            OutTOBNet := OutN;
                        end;
                       if not Mode then ShowModal
                       else
                       begin
                            OnMonTimer(Application);
                       end;
                       if not Okexercice then Result := FALSE;
               end;
         EXCEPT
              On E: Exception do
                begin
                AfficheListeCom('Attention, aucune donnée n''a été importée',LISTEIMPORT);
                RollBackDiscret ;
                Result := FALSE;
                end ;
          end;
          if FTimer <> nil  then FTimer.free;
          if (OutN = nil) and (OutTOBNet <> nil) then
             OutTOBNet.free;
          Free;
     end;
END ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Module e traitement d'export
Créé le ...... : 28/04/2005
Modifié le ... :   /  /
Description .. : Procédure d'affichage dans la liste qui est appelée par le
Suite ........ : process serveur
Mots clefs ... :
*****************************************************************}

{$IFDEF EAGLCLIENT}
procedure TFAssistImport.ComsxOnCallBack(Sender: TObject) ;
var
  lemsg, param, msg,table: string ;
begin
if (Sender is Tob) then
  begin
  param:=Tob(Sender).getValue('Param') ;
  msg:=READTOKENPipe(param,'|') ;
  table:=READTOKENPipe(param,'|') ;
  table:=READTOKENPipe(param,'|') ;
  if (Table<>'')  then
    begin
    LeMsg:='Traitement en cours ...' ;
    end else
    begin
    LeMsg:=msg ;
    end ;
    AfficheListeComImport(LeMsg, LISTEIMPORT);
  end ;

end ;
{$ENDIF EAGLCLIENT}


procedure TFAssistImport.OnMonTimer(Sender: TObject);
begin
  if not OkImportS then exit;
  FTimer.Enabled := FALSE;
  bSuivant.Enabled := FALSE;
  bPrecedent.Enabled := FALSE;
  bFin.Enabled := FALSE;
  bFinClick(Sender);
  Close ;
end;

procedure TFAssistImport.bFinClick(Sender: TObject);
var
FileName                   : string;
TF,InTob                   : TOB;
Ind                        : integer;
Fichier                    : string;
PrinterSetupDialog1        : TPrinterSetupDialog;
FichierImp                 : string;
OkNet                      : Boolean;
st                         : array[0..255] of char;
afficheerreur              : Boolean;
 begin
  inherited;
    if bFin.Caption = 'Fermer' then begin ModalResult := 1; exit; end;

    While GetPage <> PIMPORT do
            bSuivantClick(Sender);
    FileName := FICHENAME.Text;

    // Si fichier ZIP, on l'extrait
    if UpperCase(ExtractFileExt(FileName)) = '.ZIP' then
      BTWINZIPClick(Sender);

    Label15.Caption := 'Traitement en cours ...'; // Fiche 10581
    if BNetExpert.Checked then
    begin
          OkNet := TRUE;
          InTob := TOB.create('Input_TOBM',nil,-1);
          InTob.Dupliquer(OutTOBNet,TRUE,TRUE,TRUE);
           For ind:=0 to InTob.Detail.count-1 do
            begin
            if not InTOB.detail[ind].FieldExists('ETAT') then
               begin
               InTOB.detail[ind].AddChampSupValeur('ETAT','');
               InTOB.detail[ind].AddChampSupValeur('ETATLIB','');
               end;
            end;

          if InTob.GetValue ('ERROR') = '0' then
          begin
               Ind := 0;   afficheerreur := FALSE;
               TF := InTob.Detail[0];
               while TF <> nil  do
               begin
                   if Ind = InTob.detail.Count then break;
                   Fichier := InTob.detail[Ind].Getvalue ('FICNOM')+'.'+ InTob.detail[Ind].Getvalue ('FICMSQ');
//GP                   if (FLISTEFICHIER.TEXT <> '<<Tous>>') and (pos (Fichier, FLISTEFICHIER.Text) = 0) then
                   if (FLISTEFICHIER.TEXT <> Traduirememoire('<<Tous>>')) and (pos (Fichier, FLISTEFICHIER.Text) = 0) then
                   begin
                        InTob.detail[Ind].free; dec(Ind);
                   end;
                 if Ind < InTob.Detail.Count-1 then  TF := InTob.detail[Ind+1]
                 else TF := nil;
                 inc (Ind);
               end;

               TF := NeREceptionFichier(InTob);
               for Ind :=0 to TF.detail.count-1 do
               begin
//                   InTob.detail[Ind].AddChampSupValeur('ETAT', '');
//                   InTob.detail[Ind].AddChampSupValeur('ETATLIB', '');
                   if TF.detail[Ind].GetValue('ERROR') <> '0' then
                   begin
                         PGIInfo (TF.detail[Ind].GetValue('ERRORLIB'));
                         OkNet := FALSE;
                         break;
                   end;
                   FICHENAME.Text := GetEnvVar('TEMP')+ '\' + TF.detail[Ind].Getvalue ('FICNOM')+'.'+ TF.detail[Ind].Getvalue ('FICMSQ');
                   if (FileExists(FICHENAME.Text)) then EnvoieImport
                   else
                   begin
                        OkNet := FALSE;
                        PGIBox('Fichier : ' + FICHENAME.Text + ' inexistant');
                   end;
                   if (not OKexercice) or (not OkNet) then
                   begin
                        InTob.detail[Ind].PutValue ('ETAT', '-');
                        InTob.detail[Ind].PutValue ('ETATLIB', 'Erreur de l''intégration ');
                        FichierImp := EcritureTraceImport(FALSE);
                        EcrireDansfichierListeCom (FichierImp, LISTEIMPORT);
                        if PGIAsk('Attention, import du fichier '+TF.detail[Ind].Getvalue ('FICNOM')+' est impossible. Voulez-vous Visualiser le rapport ?','Import')=mrYes then
                           WinExec(StrPCopy(st,'notepad '+FichierImp),SW_SHOWNORMAL);
                        AglSendMail('Rapport Comsx', 'synchro-S1-S5@cegid.fr', '', HTStringList(FCorpsMail.Lines), FichierImp, true, 1);
                        if (Not OkNet) then OKexercice := FALSE;
                        TF := NeReceptionCR(InTob);
                        if TF.GetValue ('ERROR') <> '0' then  // fiche 21543
                        begin
                            PGIInfo ('Il y a une erreur de réception');
                            OKexercice := FALSE;  afficheerreur := TRUE;
                        end;
                        break;
                   end
                   else
                       InTob.detail[Ind].Putvalue('ETAT', 'X');
                   if (FileExists(FICHENAME.Text)) and OKexercice then DeleteFile(FICHENAME.Text);
               end;
               TF := NeReceptionCR(InTob);                    // fiche 21543
               if (TF.GetValue ('ERROR') <> '0') and (not afficheerreur) then
               begin
                    PGIInfo ('Il y a une erreur de connection');
                    OKexercice := FALSE;
               end;
          end
          else
          begin
                PGIInfo (InTob.GetValue('ERRORLIB'));
                OkNet := FALSE;
          end;
          if Intob <> nil then InTob.free;
          FichierImp := EcritureTraceImport(OkNet);
          if (stArg = '') then
          begin
                if OKexercice then CommitTrans
                else RollBackDiscret ;
                OkCommit := TRUE;
                AfficheListeComImport   ('', LISTEIMPORT, 100);
                if PGIAsk('Import terminé : voulez-vous imprimer le rapport ?','Import')=mrYes then
                begin
                    PrinterSetupDialog1 := TPrinterSetupDialog.create(Application);
                    if PrinterSetupDialog1.Execute then
                       ControlTextToPrinter(FichierImp,poPortrait);
                    PrinterSetupDialog1.Free;
                end;
          end;
    end
    else
    begin
        if (FileName <> '') then
        begin
             if FileExists(FileName) then
{$IFDEF EAGLCLIENT}
                EnvoieImport(FileName)
{$ELSE}
                EnvoieImport
{$ENDIF}
             else
              PGIBox('Attention, le fichier '+ FileName+ ' d''import inexistant', 'Import');
        end
        else
              PGIBox('Attention, le fichier d''import n''est pas renseigné', 'Import');
    end;
end;

procedure TFAssistImport.EnvoieImport(FichZipp : string='');
var
SLect            : string;
FichierIE        : TextFile;
ExisteCollectif  : Boolean;
Gen,Aux1,Aux2    : string;
OkEntete         : Boolean;
Deb,Stt          : String;
OkImport         : array [0..4] of Boolean;
ii,Nbexr,id      : integer;
Nbexototal,Nbexot: integer;
FileName         : string;
ExisteDos,Cok    : Boolean;
//InfoImp          : PtTInfoImport ;
OkEnteteCom      : Boolean;
RetourConf       : string;
Exerc            : TExoDate;
EtatExercice     : string;
NbExerciceouv    : integer;
DImmo1, DFichierIMMO1 : TDateTime;
DImmo2, DFichierIMMO2 : TDateTime;
DFichierIMMO1OOUV     : TDateTime;
CodeEXO,ExocloIM : string;
OKIMMO           : Boolean;
Q1               : TQuery;

procedure SortieErreur;
begin
    CloseFile(FichierIE);
    if WindowState  = wsMinimized then FiniProgressbar;
end;
begin
    OkEntete     := FALSE;
    Nbexr        := 0;
    Nbexototal   := 0;
    Sansaux      := FALSE;
    DateArrete   := iDate1900;
    Changementgene := FALSE; Changementaux := FALSE;
    for ii := 0 to 4 do OkImport[ii] := FALSE;
    RetourConf       := '-';
    OkEnteteCom      := TRUE;
    NbExerciceouv    := 0;
    InitGI           := FALSE;
    ScanBanck        := TRUE;

     // Fiche 10423
     OkIMMO := FALSE;  DImmo1 := iDate1900; DImmo2 := iDate1900;   DFichierIMMO1 := iDate1900;
     DFichierIMMO2 := iDate1900;  DFichierIMMO1OOUV := iDate1900;

     if stArg <> '' then
     begin
                 if FileExists(RepEncours+'ListeCom'+ V_PGI.NoDossier+'.OK') then
                    DeleteFile(RepEncours+'ListeCom'+ V_PGI.NoDossier+'.OK');
                 if FileExists(RepEncours+'ListeCom'+ V_PGI.NoDossier+'.ERR') then
                    DeleteFile(RepEncours+'ListeCom'+ V_PGI.NoDossier+'.ERR');
     end;

    FileName := FICHENAME.Text;

    if WindowState  = wsMinimized then
    begin
       Stt := 'Import Fichier : ' + ExtractFileName(FileName);
       InitProgressbar(Stt);
    end;
    AssignFile(FichierIE, FileName);
    Reset(FichierIE);
    if not EOF(FichierIE) then
    begin
      Readln(FichierIE, SLect);

      Deb := EstUneLigneCom(SLect);
      // vérification de l'entête
      if (Copy (SLect,0, 4) = Deb+'S') and (Copy (SLect,0, 3) = '***') then
      begin
        if (Copy (SLect,4, 2) <> '') or  (Copy (SLect,6, 3) <> '') or (Copy (SLect,9, 3) <> '')
        or (Copy (SLect,12, 3) <> '') then OkEntete := TRUE;

        if (Copy (SLect, 54, 29) = '(c)xperts - ScanBANK/ScanFACT') then
        begin
            ScanBanck := VH^.OkModSCAN;
            if not ScanBanck then
                AfficheListeCom('Cegid Expert SCAN n''est pas sérialisé : les documents joints ne seront pas intégrés', LISTEIMPORT);
        end;
      end
      else
      begin
          if (Copy (SLect,0, 11) = '***DEBUT***') then OkEntete := TRUE
          else OkEntete := FALSE;
      end;
      if (Copy (SLect,0, 5) = Deb+'QU')  then OkEntete := TRUE;
      if (Copy (SLect,0, 5) = Deb+'WT') or (Copy (SLect,0, 5) = Deb+'WS') then OkEntete := TRUE;
      if (Copy (SLect,0, 1) = '!') or  (EnleveBlanc (SLect,0, 35) = EnleveBlanc(GetParamSocsecur('SO_LIBELLE', ''),0, 35)) then
      begin
               OkEntete := TRUE;
               OkEnteteCom := FALSE;
               NatureFichier  := 'JRL';
               OrigineFichier := 'S5';
               QualifOrigine  := 'ND0' ;
               Paramgene := FALSE;

               ExisteDos :=  ExisteSQl ( 'SELECT EX_EXERCICE from EXERCICE');
               If not ExisteDos then
               begin
                  PGIBox ('Le dossier comptable n''existe pas. Veuillez le créer.', 'Import');
                  SortieErreur;
                  exit;
               end;
      end;

      if (Pos ('Maman', SLect) <> 0) then
      begin
               NatureFichier  := 'DOS';
               OrigineFichier := 'IMMO';
               OkEntete := TRUE;
               QualifOrigine  := 'ND0' ;
               Paramgene := FALSE;
               RetourConf := 'X';
      end;

{$IFDEF COMSX}
      if Uppercase(ExtractFileExt(FileName)) = '.NSV' then
      begin
               NatureFichier  := 'DOS';
               OrigineFichier := 'PAIE';
               OkEntete := TRUE;
               QualifOrigine  := 'ND0' ;
               Paramgene := FALSE;
               RetourConf := 'X';
      end;
{$ENDIF}

      if not OkEntete then
      begin
               Readln(FichierIE, SLect);    // ajout me pour Fichier CCIMPEX
               if (Copy (SLect, 0, 3) = '***') or ((Copy (SLect, 0, 3) <> '***') and
               (ExisteSQL ('SELECT J_JOURNAL FROM JOURNAL Where J_JOURNAL="'+ EnleveBlanc (SLect,0, 3)+'"'))) then
               begin
                       OkEntete := TRUE;
                       OkEnteteCom := FALSE;
                       NatureFichier  := 'JRL';
                       OrigineFichier := 'S5';
                       QualifOrigine  := 'ND0' ;
                       Paramgene := FALSE;

                       ExisteDos :=  ExisteSQl ( 'SELECT EX_EXERCICE from EXERCICE');
                       If not ExisteDos then
                       begin
                          PGIBox ('Le dossier comptable n''existe pas. Veuillez le créer.', 'Import');
                          SortieErreur;
                          exit;
                       end;

              end
              else
              begin
                    if stArg <> '' then
                    begin
                      AfficheListeCom(ENTETE_INEXSISTANT, LISTEIMPORT);
                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                    end
                    else
                      PGIBox ('Entête du fichier inexistante ou fichier incorrecte.', 'Import');
                      SortieErreur;
                      exit;
              end;
      end;
      if (Copy (SLect,0, 5) = Deb+'S1') or (Copy (SLect,0, 5) = Deb+'S3') or  (Copy (SLect,0, 5) = Deb+'S2') or (Copy (SLect,0, 5) = Deb+'S5') or
         (Copy (SLect,0, 5) = Deb+'S7') or (Copy (SLect,0, 5) = Deb+'SI')
         // pour quadra
         or (Copy (SLect,0, 5) = Deb+'QU')
         // pour CCMX
         or (Copy (SLect,0, 5) = Deb+'WT')
         or (Copy (SLect,0, 5) = Deb+'WS')
         // pour compta finance
         or (Copy (SLect,0, 5) = Deb+'CF') then
      begin
           if Copy (SLect,4, 2) <> '' then
              OrigineFichier := Copy (SLect,4, 2);
           if Copy (SLect,9, 3) <> '' then
              NatureFichier := Copy (SLect,9, 3);
           if (NatureFichier = 'JRL') then
           begin
           // RED = cas reprise type JRL mais nature dossier forcée
           // REJ= cas reprise type JRL avec reprise lettrage
                RepriseComplement := EnleveBlanc (SLect, 124, 4);
                if RepriseComplement = 'RED' then
                 NatureFichier := 'DOS';
                if (EnleveBlanc (SLect, 125, 3) = 'IGI') then  // initialisation GI
                begin
                    RepriseComplement := 'IGI';
                    InitGI      := TRUE;
                end;
           end;
           if (NatureFichier = 'SYN') and (EnleveBlanc (SLect, 125, 3) = 'IGI') then
              RepriseComplement := 'IGI';
           if (NatureFichier = 'SYN') and (FileExists(RepEncours+'ComLettrage.log')) then DeleteFile(RepEncours+'ComLettrage.log');


           if (NatureFichier = 'DOS') and (EnleveBlanc (SLect, 125, 3) <> '')
           and (Copy (SLect,0, 5) <> Deb+'S1') then
                RepriseComplement := EnleveBlanc (SLect, 124, 4);


           QualifOrigine := Copy(NatureFichier, 0, 2)+ Copy(OrigineFichier, 2, 1);

           ExisteDos :=  ExisteSQl ( 'SELECT EX_EXERCICE from EXERCICE');
           if (NatureFichier = 'JRL') or (NatureFichier = 'SYN')then
           begin
               If not ExisteDos then
               begin
                    if stArg <> '' then
                    begin
                      AfficheListeCom(DOSSIER_INEXSISTANT,LISTEIMPORT);
                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                    end
                    else
                      PGIBox ('Le dossier comptable n''existe pas. Veuillez le créer.', 'Import');
                    SortieErreur;
                    exit;
               end;
               if (stArg = '') and (Copy (SLect,0, 5) = Deb+'S1') and (NatureFichier = 'SYN') and
               (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE) = FALSE) then
               begin
                    if GetparamsocSecur ('SO_CPSYNCHROSX', TRUE) = TRUE then
                        SetParamsoc ('SO_FLAGSYNCHRO', 'SYN')
                    else
                    begin
                       PGIInfo ('Attention, le dossier n''est pas paramétré en mode synchronisation.' + #10#13
                       +'Paramètres société, Echanges, cochez l''option Synchronisation');
                       SortieErreur; exit;
                    end;
               end;
               if (stArg = '') and (Copy (SLect,0, 5) = Deb+'S1') and
               (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE) = TRUE) and (NatureFichier = 'JRL') then
               begin
                       PGIInfo ('Le dossier est en mode synchronisation, la réception de fichiers de type journal ou balance en provenance de'+#10#13
                       + ' CEGID Business Line est impossible." ');
                       SortieErreur; exit;
               end;
               if (stArg = '') and (Copy (SLect,0, 5) = Deb+'S1')  then
               begin
                  if ExisteSQL ('select  ##TOP 1##y_axe from analytiq where (y_axe="A2" or y_axe="A3" or y_axe="A4" or y_axe="A5")') then
                  begin
                     PGIInfo ('Vous effectuez des échanges synchronisés avec Business Line. Export impossible.'
                     + #10#13
                     +' Un axe (analytique ou TVA) autre que l''axe A1 est utilisé');  exit;
                  end;
               end;
           end;
           if (NatureFichier = 'BAL') and  // fiche 10457
           (stArg = '') and (Copy (SLect,0, 5) = Deb+'S1') and
           (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE) = TRUE) then
           begin
                       PGIInfo ('Le dossier est en mode synchronisation, la réception de fichiers de type journal ou balance en provenance de'+#10#13
                       + ' CEGID Business Line est impossible." ');
                       SortieErreur; exit;
           end;
           if (NatureFichier <> 'JRL') and (NatureFichier <> 'DOS')
           and (NatureFichier <> 'BAL') and (NatureFichier <> 'SYN')
           and (NatureFichier <> 'JAL') then
           begin
                    if stArg <> '' then
                    begin
                      AfficheListeCom(ENTETE_TYPEINCORRECT + NatureFichier+ENTETE_TYPEINCSUITE,LISTEIMPORT);
                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                    end
                    else
                      PGIBox (ENTETE_TYPEINCORRECT + NatureFichier+#10#13+ENTETE_TYPEINCSUITE, 'Import');
                    SortieErreur;
                    exit;
           end;
           if  (NatureFichier = 'JAL') then NatureFichier := 'JRL';

           OrigineEnvoie := Copy (SLect,6, 3);

           // pour synchro de la base commune comporte comme origine client
           if (RepriseComplement = 'IGI') and (EnleveBlanc (SLect, 124, 1) = 'C') then OrigineEnvoie := 'CLI';

           if (EnleveBlanc(SLect, 26, 8,'') <> '') and (EnleveBlanc(SLect, 26, 8,'') <> stDate1900) then
              DateArrete    := Format_Date_HAL(EnleveBlanc(SLect, 26, 8,''));

           if (OrigineEnvoie <> 'EXP') and (OrigineEnvoie <> 'CLI') then
           begin
                    if stArg <> '' then
                    begin
                      AfficheListeCom(ENTETE_ORIGINE + OrigineEnvoie + ENTETE_ORIGINESUITE,LISTEIMPORT);
                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                    end
                    else
                      PGIBox (ENTETE_ORIGINE + OrigineEnvoie + #10#13 + ENTETE_ORIGINESUITE, 'Import');
                    SortieErreur;
                    exit;
           end;

           if (Copy (SLect,12, 3) <> 'STD') and (Copy (SLect,12, 3) <> 'ETE') then
           begin
                    if stArg <> '' then
                    begin
                      AfficheListeCom(ENTETE_FORMAT + Copy (SLect,12, 3) + ENTETE_FORMATSUITE,LISTEIMPORT);
                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                    end
                    else
                     PGIBox (ENTETE_FORMAT + Copy (SLect,12, 3) + #10#13 + ENTETE_FORMATSUITE, 'Import');
                    SortieErreur;
                    exit;
           end;
           CodeEXO := '001';
           if (NatureFichier = 'DOS')  and  AuMoinsUneImmo and (GetParamSocsecur ('So_exocloimmo', '') <>'') then
           begin
                   Q1 := OpenSQl ('SELECT EX_DATEDEBUT, EX_DATEFIN, EX_EXERCICE from EXERCICE where ex_exercice="'+GetParamSocsecur ('So_exocloimmo', '', TRUE)+'"', TRUE);
                   if Not Q1.EOF then
                   begin
                    DImmo1 := Q1.FindField ('EX_DATEDEBUT').asDateTime;
                    DImmo2 := Q1.FindField ('EX_DATEFIN').asDateTime;
                    CodeEXO  := Q1.FindField ('EX_EXERCICE'). asstring;
                    OkIMMO := TRUE;
                   end;
                   Ferme (Q1);
           end;
           // Controle préalable des paramètres et des comptes
           // avant remonter du fichier
           while not EOF(FichierIE) do
           begin
                  Readln(FichierIE, SLect);

                  Deb := EstUneLigneCom(SLect);
                  if Copy (SLect,0, 6) = Deb+'PS1' then
                       OkImport[0] := TRUE;

                  if Copy (SLect,0, 6) = Deb+'PS2' then
                  begin
                       OkImport[1] := TRUE;
                       if not Controlparamsociete  (SLect) then begin SortieErreur; exit; end;
                  end;
                  if Copy (SLect,0, 6) = Deb+'PS3' then
                  begin
                       OkImport[2] := TRUE;
                       if not Controlparamsociete  (SLect) then begin SortieErreur; exit; end;
                  end;
                  if Copy (SLect,0, 6) = Deb+'PS5' then
                  begin
                       OkImport[3] := TRUE;
                       if not ControlParam5 (SLect, ExisteDos)then begin SortieErreur; exit; end;
                  end;
                  if Copy (SLect,0, 6) = Deb+'EXO' then
                  begin
                       inc(Nbexototal);
                       Cok := TRUE;
                       EtatExercice := EnleveBlanc(SLect, 26, 3,'');
                       if (EtatExercice = 'OUV') then inc (NbExerciceouv);
                       if ((NatureFichier = 'BAL') or (NatureFichier = 'JRL')  or (NatureFichier = 'SYN')
                       or ((EtatExercice = 'OUV') and (NatureFichier = 'DOS')  and (NbExerciceouv = 1) and OkIMMO)) // Fiche 10423
                       and ExisteDos then  Cok := COkExercice (SLect);

                       if OkIMMO then // Fiche 10423
                       begin
                            if (EtatExercice = 'CDE')  then
                            begin
                              DFichierIMMO1 := Format_Date_HAL(EnleveBlanc(SLect, 10, 8,''));
                              DFichierIMMO2 := Format_Date_HAL(EnleveBlanc(SLect, 18, 8,''));
                              ExocloIM  := Format ('%.03d',[Nbexototal]);
                            end
                            else
                            if (EtatExercice = 'OUV') and (NbExerciceouv =1) then
                               DFichierIMMO1OOUV := Format_Date_HAL(EnleveBlanc(SLect, 10, 8,''));
                       end;


                       if (not Cok) then
                       begin
                                SortieErreur; exit;
                       end;
                       if (EtatExercice <> 'OUV') and (EtatExercice <> 'CDE') and (EtatExercice <> 'CPR') and
                          (EtatExercice <> 'CLO') and (EtatExercice <> 'NON') then
                       begin
                                if stArg <> '' then
                                begin
                                  AfficheListeCom(ERR_EXERCICEETAT,LISTEIMPORT);
                                  EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                  PgiBox (ERR_EXERCICEETAT, 'Import');
                                SortieErreur; exit;
                       end;
                       if not NbExoouv(SLect, Nbexr) then begin SortieErreur; exit; end;
                       if not ControlEnreg (SLect,'EXERCICE', 0, 3) then begin SortieErreur; exit; end;
                  end;
                  if not ControleAutreParam (Deb, SLect)then begin SortieErreur; exit; end;
                  // si format Balance
                  if (NatureFichier = 'BAL') and not EstUneLigneCpt(SLect) then
                  begin
                        if CQuelExercice (Format_Date_HAL(EnleveBlanc(SLect, 4, 8)), Exerc) then
                        begin
                          if (stArg = '') then
                          begin
                             if (Exerc.EtatCpta <> 'OUV') AND (Exerc.EtatCpta <> 'CPR') and (Exerc.EtatCpta <> 'NON') then
                             begin
                                  AfficheListeCom('L''exercice est clos. L''importation des écritures est impossible.' ,LISTEIMPORT);
                                  PgiBox ('L''exercice est clos. L''importation des écritures est impossible.');
                                  SortieErreur; exit;
                             end;
                          end;
                        end;
                  end;
           end;

           Nbexot := NombreExercice('');
           if (Nbexot <> Nbexototal) and (Nbexot <> 0) and (Nbexototal <> 0) then
           begin
             if (stArg = '') and (NatureFichier = 'DOS') then
                If PGIAsk('Le nombre d''exercice est différent entre le dossier et le fichier.#10#13Voulez-vous importer les écritures ?','Import')<>mrYes Then exit;
           end;

           if OkIMMO then  // Fiche 10423
           begin
                if (DFichierIMMO2 <> iDate1900) and (DImmo1 <> DFichierIMMO1) and (DImmo2 <> DFichierIMMO2) then
                begin
                    PGIInfo ('Import impossible. La réception du fichier va générer une incohérence des immobilisations.');
                    SortieErreur;
                    exit;
                end;
           end;

           for ii := 0 to 3 do
           begin
                if not OkImport[ii] then
                begin
                     if not ControleParamSoc(ii) then
                     begin
                                if ii = 3 then id := ii+2 // ps5
                                else id := ii+1;

                                if stArg <> '' then
                                begin
                                  AfficheListeCom(ERR_PARAMETRE+inttostr(id)+ ERR_PARAMETRESUITE,LISTEIMPORT);
                                  EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                  PGIBox (ERR_PARAMETRE+inttostr(id)+ ERR_PARAMETRESUITE, 'Import');
                                SortieErreur;
                                exit;
                     end;
                end;
           end;
           if (NatureFichier = 'DOS') and (stArg <> '') and (Nbexototal = 0) then
           begin
                                  AfficheListeCom(ERR_EXOMANQUANT,LISTEIMPORT);
                                  EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                  SortieErreur; exit;
           end;


           if (stArg <> '') or (NatureFichier = 'SYN') then RetourConf := 'X'
           else
           begin
               if (not SuppComptable.Checked) then RetourConf := 'X'
               else RetourConf := '-';
           end;

           if RetourConf = '-' then
           RetourConf := AglLanceFiche('CP','IMPORTCOMCONF','', '',
           'Origine du fichier : '+ OrigineFichier + '; Nature : ' + NatureFichier
           +'; Fichier : '+ FICHENAME.Text)  ;
      end;
      if (OrigineFichier <> 'PAIE')  and (Copy (SLect,0, 11) = '***DEBUT***') then
      begin
           OrigineFichier := 'SI';
           Readln(FichierIE, SLect);
           if Copy(SLect,20,1) = 'E' then NatureFichier := 'EXE'
           else
           if Copy(SLect,20,1) = 'B' then NatureFichier := 'BAL'
           else
           if Copy(SLect,20,1) = 'J' then NatureFichier := 'JRL';

           QualifOrigine := Copy(NatureFichier, 0, 2)+ Copy(OrigineFichier, 1, 1);
           if(NatureFichier = 'JRL') then
           begin
               If not ExisteSQl ( 'SELECT EX_EXERCICE from EXERCICE') then
               begin
                               if stArg <> '' then
                                begin
                                  AfficheListeCom(ERR_EXODOSINEXIST,LISTEIMPORT);
                                  EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                  PGIBox (ERR_EXODOSINEXIST, 'Import');
                                SortieErreur;
                                exit;
               end;
            end;

           ExisteCollectif := FALSE;
           While (Not EOF(FichierIE)) do
           BEGIN
                ReadLn(FichierIE,SLect) ;
                If Copy(SLect,1,2)='08' Then
                begin
                     Gen:=BourreEtLess(Trim(Copy(SLect,3,10)),fbGene) ;
                     Aux1:=BourreEtLess(Trim(Copy(SLect,13,10)),fbAux) ;
                     Aux2:=BourreEtLess(Trim(Copy(SLect,23,10)),fbAux) ;
                     if ExisteSQl('SELECT * FROM CORRESP WHERE CR_TYPE="SIS" AND CR_CORRESP="'+GEN+'"') then
                     begin
                          if Not ExisteSQl('SELECT * FROM CORRESP WHERE CR_TYPE="SIS" AND CR_CORRESP="'+GEN+'" And CR_LIBELLE="'+Aux1+'" AND CR_ABREGE="'+Aux2+'" ') then
                          begin
                               if stArg <> '' then
                                begin
                                  AfficheListeCom(ERR_TRANCHEAUXIL,LISTEIMPORT);
                                  EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                PGIBox (ERR_TRANCHEAUXIL, 'Import Sisco II');
                                exit;
                          end;
                          if (VH^.CPLienGamme = '')  then
                          SetParamsoc ('SO_CPLIENGAMME', 'SI')
                     end;
                     ExisteCollectif := TRUE;
                end;
           END;

           if (not ExisteCollectif) then
           begin
                if stArg <> '' then Sansaux := TRUE
                else
                begin
                    If PGIAsk('Vous n''avez pas de comptabilité auxiliaire, voulez-vous continuer ?','Import')<>mrYes Then
                         exit
                    else
                     Sansaux := TRUE;
                end;
           end;


           if (NatureFichier = 'EXE') or ((stArg <> '') and (OrigineFichier ='SI')) then RetourConf := 'X'
           else
           if RetourConf = '-' then
           RetourConf := AglLanceFiche('CP','IMPORTCOMCONF','', '',
           'Origine du fichier : '+ OrigineFichier + '; Nature : ' + NatureFichier
           +'; Fichier : '+ FICHENAME.Text)  ;
      end
      else
      begin
          if OptionEchange  and (not OkEntete) then // Fiche 16078
          begin
                Readln(FichierIE, SLect);
                if EOF(FichierIE) then
                begin
                          PGIInfo('Fichier incorrect, aucune donnée comptable n''est importée');
                          SortieErreur;
                          exit;
                end;
          end;
      end;
      CloseFile(FichierIE);
      if (FichZipp <> '') and (UpperCase(ExtractFileExt(FichZipp)) = '.ZIP') then
               DeleteFile(FICHENAME.Text);

      if RetourConf = 'X' then
      begin
      // dans le cas de la balance et dossier on regarde si on a supprimer le dossier
        if (NatureFichier = 'BAL')  or (NatureFichier = 'DOS') or (NatureFichier = 'EXE')then
        begin
            if not ExisteSQl ('SELECT EX_EXERCICE from EXERCICE') then
               Paramgene := TRUE
            else
               if Changementgene or Changementaux then
                  Paramgene := TRUE
               else
                  Paramgene := FALSE;
            // pour ne pas écraser les paramètres généraux , test à voir plus tard
            if (OrigineFichier = 'SI') and ((NatureFichier = 'BAL')or (NatureFichier='EXE')) then
                    Paramgene := TRUE
        end else  Paramgene := FALSE;
{$IFDEF EAGLCLIENT}
          if (FichZipp <> '') then FICHENAME.Text := FichZipp;
{$ENDIF}
          ImportComSx(FICHENAME.Text);
      end // cas ecriture
      else
      begin
      // s'il y a une entête ccimpex c'est ok
           if OkEnteteCom  then exit;
           NatureFichier  := 'JRL';
           OrigineFichier := 'S5';
           QualifOrigine  := 'ND0' ;
           Paramgene := FALSE;
{$IFDEF EAGLCLIENT}
          if (FichZipp <> '') then FICHENAME.Text := FichZipp;
{$ENDIF}
           ImportComSx(FICHENAME.Text);
      end;
      if OkExercice then
      begin
             // Fiche 10423 mise à jour du code exo immo
             if OkIMMO then
             begin
                   if (ExocloIM <> CodeEXO) and  (DImmo1 = DFichierIMMO1) and (DImmo2 = DFichierIMMO2) then
                   begin
                         Q1 := OpenSQl ('SELECT EX_DATEDEBUT, EX_DATEFIN, EX_EXERCICE from EXERCICE where EX_DATEDEBUT="'
                        + UsDateTime(DImmo1) +'" and EX_DATEFIN="'+ UsDateTime(DImmo2) + '"', TRUE);
                        if Not Q1.EOF then
                          CodeEXO  := Q1.FindField ('EX_EXERCICE'). asstring
                        else CodeExo := '';
                         Ferme (Q1);
                         SetParamSoc ('So_exocloimmo', CodeExo);
                   end
                   else
                   begin
                        if (DFichierIMMO2 = iDate1900) and (DImmo2 = DFichierIMMO1OOUV-1) then
                           SetParamSoc ('So_exocloimmo', '');
                   end;
             end;
      end;
    end;
  if WindowState  = wsMinimized then FiniProgressbar;
end;


Function TFAssistImport.EstUneLigneCom(St : String) : string;
Var Cod : String ;
BEGIN
Cod:=Copy(St,1,3) ;
Result:=Cod ;
END ;

function TFAssistImport.EcritureTraceImport (OkFinImport : Boolean): string;
Var
FichierImp, Rep : string;
posx            : integer;
FichTmp         : string;
begin
 Rep := GetParamSocSecur ('SO_CPRDREPERTOIRE', '', TRUE);
 if not DirectoryExists(Rep) then CreateDir(Rep);
 if stArg <> '' then
 begin
        if FichImp <> '' then
        begin
             FichierImp := RepEncours+FichImp;
             if not OkFinImport then
             begin
                posx := Pos('.',FichierImp);
                if Posx = 0 then Posx := length(FichierImp)+1;
                FichierImp := Copy(FichierImp,0,Posx-1) + '.ERR';
             end;
        end
        else
        begin
                 if not (ctxPCL in V_PGI.PGIContexte) then
                 begin
                  FichTmp := ExtractFileName(FICHENAME.Text);
                  FichTmp := 'ListeCom'+ ReadTokenPipe (FichTmp, '.');
                 end
                 else FichTmp := 'ListeCom'+ V_PGI.NoDossier;
                 if OkFinImport then
                   FichierImp := RepEncours+FichTmp+'.OK'
                 else
                   FichierImp := RepEncours+FichTmp+'.ERR';
        end;
        if BNetExpert.Checked then
        begin
                if Rep <> '' then Rep := Rep + '\'
                else Rep := GetEnvVar('TEMP') + '\';
                if OkFinImport then
                   FichierImp := Rep +'ListeCom'+ V_PGI.NoDossier+'.OK'
                else
                   FichierImp := Rep +'ListeCom'+ V_PGI.NoDossier+'.ERR';
        end;

        if Not OkFinImport then
             OKexercice    := FALSE;
 end
 else
 begin
    if Rep <> '' then
       FichierImp := Rep+'\'+'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt'
    else
    begin
       if BNetExpert.Checked then
          FichierImp :=  GetEnvVar('TEMP') + '\'+ 'ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt'
       else
       begin
            FichTmp := ExtractFileName(FICHENAME.Text);
            FichTmp := ReadTokenPipe (FichTmp, '.');
            if FichTmp <> '' then
               FichierImp := ExtractFileDir(FICHENAME.Text)+ '\ListeCom'+FichTmp+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt'
            else
               FichierImp := ExtractFileDir(FICHENAME.Text)+'\ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt';
       end;
    end;
 end;
 Result := FichierImp;
{$IFDEF CERTIFNF}
 MajJournalImport (FICHENAME.Text);
{$ENDIF}

end;



procedure TFAssistImport.bClickVisuClick(Sender: TObject);
{$ifdef COMSX}
var
nomfichier : string;
{$ENDIF}
begin
  inherited;
     if FICHENAME.Text = '' then exit;
{$ifdef COMSX}
          nomfichier:=ExtractFileName(FICHENAME.Text);
          ReadTokenPipe(nomfichier,'.');
          if UpperCase(nomfichier)<>'TRA' then
              PGIInfo('Le fichier selectionné n''est pas un fichier au format TRA')
          else
              ControlFichierTra(FICHENAME.Text);
{$endif}
end;

procedure TFAssistImport.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
    VK_F10    : bFinClick(Sender);
  end;
end;

procedure TFAssistImport.FormCreate(Sender: TObject);
var
serie, St     : string;
stArgsav      : string ;
TCF           : TControlFiltre;
begin
  inherited;
       Paramgene := TRUE;

       if V_PGI.LaSerie = S7 then serie := 'S5';
       if V_PGI.LaSerie = S5 then serie := 'S5';
       if V_PGI.LaSerie = S3 then serie := 'S3';

       FCorpsMail.clear;
       FCorpsMail.Lines.add ('Veuillez trouver ci-joint le fichier de rapport correspondant au dossier ');
       FCorpsMail.Lines.add (V_PGI.NoDossier  + ' géré par le logiciel '+ serie);

       FEmail.Enabled := not(FFile.Checked);
       Label2.Enabled := not(FFile.Checked);
       FHigh.Enabled  := not(FFile.Checked);
       FCorpsMail.Enabled := not(FFile.Checked);
       Label12.Enabled := not(FFile.Checked);
       Label11.Enabled := not(FFile.Checked);
       FTimer         := TTimer.Create(nil);
       FTimer.Enabled := FALSE;
       FTimer.Interval:= 2000;
       FTimer.Ontimer := OnMonTimer;
       Calculpiece    := '';
       AutoriseEcheck := TRUE;
       stArgsav := stArg;
       ControlExercClos := FALSE;
       if pos('Minimized', stArg) <> 0 then WindowState  := wsMinimized;
       AutoriseEcheck  := (pos('SANSECHEC', stArg) = 0);
       OkImportS := (stArg <> '');
       if stArg <> '' then  // par ligne de commande
       begin
                FTimer.Enabled := TRUE;
                RepEncours := ReadTokenPipe(stArg, ';');  // répertoire
                St := ReadTokenPipe(stArg, ';');   // nom du fichier
                if St = 'Minimized' then St := '';
                if FichierIni <> '' then
                      InitComImport (St)
                else
                begin
                      if pos('DECOUPELON=', St) <> 0 then
                      begin
                         StArg := ReadTokenPipe(St, '=');
                         if Isnumeric(St) then
                         begin
                            DecoupLon.value  := StrToint(St); St := '';
                         end;
                      end;
                      FICHENAME.Text := RepEncours+St;
                      FEmail.Text := ReadTokenPipe(stArg, ';');
                      if (FEmail.Text <> '') and (pos('@', FEmail.Text) = 0) then FEmail.Text := '';
                      TypeEcrR        := ReadTokenPipe(stArg, ';');
                      TypeEcrI        := ReadTokenPipe(stArg, ';');
                      if (TypeEcrR <> 'S') and (TypeEcrR <> 'N') then TypeEcrR := '';
                      if (TypeEcrI <> 'S') and (TypeEcrI <> 'N') then TypeEcrI := '';
                      Calculpiece     := ReadTokenPipe(stArg, ';'); // dans le cas journal s'il faut calculer ou non les pieces 'O' ou 'N'
                      FichImp         := ReadTokenPipe(stArg, ';');
                      St              := ReadTokenPipe(stArg, ';');
                      St              := ReadTokenPipe(stArg, ';'); // sans echec
                      BGestiondoublon.Checked  := (St = 'TRUE');
                      St              := ReadTokenPipe(stArg, ';');
                      BIntegre.Checked     := (St = 'TRUE');
                      St              := ReadTokenPipe(stArg, ';');
                      BControlParam.Checked  := (St = 'FALSE');
                end;
                stArg := stArgsav;
       end;
       TCF.Filtre := BFiltre;
       TCF.Filtres := FFiltres;
       TCF.PageCtrl := P;
       ObjetFiltre := TObjFiltre.create(TCF, 'COMIMPORT');
       BFiltre.Enabled := FALSE;
       BINI.Enabled := FALSE;
       // fiche 10631
       if (stArg = '') and (GetParamSocSecur('SO_CPRDREPERTOIRE', '') <> '') then
          Fichename.Text := GetParamSocSecur('SO_CPRDREPERTOIRE', '') + '\PG'+ V_PGI.NoDossier +'.TRA'

end;

procedure TFAssistImport.RendreDispo(Ena : Boolean);
begin
                   BMAJCPTE.Enabled               := Ena;
                   BControlParam.Enabled          := Ena;
                   BCalculPiece.Enabled           := Ena;
                   BRupturePiece.Enabled          := Ena;
                   BInterdireCreatCompte.Enabled  := Ena;
                   BGestiondoublon.Enabled        := Ena;
                   Label1.Enabled                 := Ena;
                   FCORRESP.Enabled               := Ena;
                   BCorresp.Enabled               := Ena;
                   HTYPEIMPORT.Enabled            := Ena;
                   TQUALPIECE.Enabled             := Ena;
                   Label7.Enabled                 := Ena;
                   DecoupLon.Enabled              := Ena;
                   Label4.Enabled                 := Ena;
                   BBlanc.Enabled                 := Ena;
end;

procedure TFAssistImport.bSuivantClick(Sender: TObject);
begin
  inherited;
      if (GetPage = AVANCE) and (not OPTIONAVANCE.Checked) then
         bSuivantClick(Sender);
      if (GetPage = AVANCE2) and (not OPTIONAVANCE.Checked) then
         bSuivantClick(Sender);
      if Getpage = Resume then
      begin
          FVal3.Caption := FICHENAME.Text;
          FVal4.Caption := FEMAIL.Text;
      end;
      if Getpage = Avance then
      begin
           if FICHENAME.Text = '' then
           begin
              RendreDispo(FALSE);
           end
           else  if BMAJCPTE.Enabled = FALSE then
              RendreDispo(TRUE);
           // fiche 10481
           if (NatureFichier = 'JRL') then
              BRupturePiece.Checked := TRUE
           else
              BRupturePiece.Checked := FALSE
      end;
      if Getpage = Avance2 then
      begin
             DATEECR1.visible  := (NatureFichier = 'JRL');
             LDATEECR1.visible := (NatureFichier = 'JRL');
      end;
      if Getpage = PIMPORT then
      begin
                 LISTEIMPORT.clear;
                 AfficheListeCom('*** Bouton <Fin> pour lancer le traitement *** ', LISTEIMPORT);
      end;

      if (Getpage = Resume) or (Getpage = PIMPORT) then
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


procedure TFAssistImport.FMailClick(Sender: TObject);
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

procedure TFAssistImport.FEMailExit(Sender: TObject);
begin
  inherited;
            if FEmail.Text <> '' then
            begin
                 if (pos('@', FEmail.Text) = 0) then
                 begin
                  FEmail.Text := '';  FMail.Checked := FALSE;
                 end
                 else
                 FMail.Checked := TRUE;
            end;
end;



procedure TFAssistImport.FICHENAMEChange(Sender: TObject);
var
FichierIE        : TextFile;
SLect,Deb,Format : string;
begin
  inherited;
  if FICHENAME.TEXT = '' then
  begin
       OPTIONAVANCE.enabled := FALSE;
       OPTIONAVANCE.checked := FALSE;
  end;
  if not FileExists(FICHENAME.TEXT) then exit;

    AssignFile(FichierIE, FICHENAME.TEXT);
    Reset(FichierIE);
    if not EOF(FichierIE) then
    begin
      Readln(FichierIE, SLect);
      Deb := EstUneLigneCom(SLect);
      CP_STAT.visible := (Copy (SLect,0, 11) = '***DEBUT***');
      HLABEL2.visible:= (Copy (SLect,0, 11) = '***DEBUT***');
      if (Copy (SLect,9, 3) = 'DOS') or (Copy (SLect,9, 3) = 'SYN') then
      begin
          OPTIONAVANCE.Enabled := FALSE;
          OPTIONAVANCE.Checked := FALSE;
          SuppComptable.Enabled := FALSE;
          SuppComptable.Checked := TRUE;
      end
      else
      begin
           OPTIONAVANCE.Enabled := TRUE;
           if (VH^.CPLienGamme = 'S1') then
              SuppComptable.Checked := FALSE;
           SuppComptable.Enabled := TRUE;
           NatureFichier := Copy (SLect,9, 3);
           if (Copy (SLect,0, 1) = '!') then NatureFichier := 'JRL';
      end;
      if CP_STAT.visible then
      begin
            Readln(FichierIE, SLect);
            Format := Copy (SLect, 20,1);
            SuppComptable.Enabled := FALSE;
      end;
      if (Copy (SLect,9, 3) ='BAL') and (SLect[124]='X') then
         INTERDIRECREATTIERS.Enabled := TRUE
      else
      begin
         INTERDIRECREATTIERS.Enabled := FALSE;
         INTERDIRECREATTIERS.Checked := FALSE;
      end;
    end;
    CloseFile(FichierIE);

    if IsNetExpert and (not BNetExpert.Checked) then  // si dossier net expert
    begin
         if not V_PGI.SAV and (Copy (SLect,9, 3)= 'DOS') or ((Format<>'') and (Format <> 'J') and  (Format <> 'B')) then
         begin
              PGIInfo ('Uniquement, le format journal ou balance est accepté');
              FICHENAME.Text := '';
         end;
    end;
end;


procedure TFAssistImport.BNetExpertClick(Sender: TObject);
begin
  inherited;
    BNetExpert.Enabled := IsNetExpert;
    if IsNetExpert then
    begin
         if (BNetExpert.Checked) then  // si dossier net expert
            FICHENAME.Text := '';
       FICHENAME.Enabled := (not (BNetExpert.Checked));
       Label3.Enabled := (not (BNetExpert.Checked));
       TListeFichier.Enabled := (BNetExpert.Checked);
       FLISTEFICHIER.Enabled := (BNetExpert.Checked);
    end;
end;

procedure TFAssistImport.FormShow(Sender: TObject);
var NoSEqNet,i       : integer;
TNetRecep,TA         : TOB;
Fichier              : string;
EnvoiRecep           : string;
begin
  inherited;
       IsNetExpert           := IsDossierNetExpert (V_PGI.NoDossier, NoSEqNet);
       EnvoiRecep            := 'R';
       BNetExpert.Checked    := IsNetExpert;
       BNetExpert.visible    := (BNetExpert.Checked);
       BRecalculSolde.Enabled := (V_PGI.SAV  = TRUE);
       CBBAP.visible := (not (V_PGI.ModePCL='1')) and (ExisteTypeVisa (''));
       LBBAP.visible := (not (V_PGI.ModePCL='1')) and (ExisteTypeVisa (''));
       if IsNetExpert then
       begin
           FICHENAME.Enabled := (not (BNetExpert.Checked));
           Label3.Enabled := (not (BNetExpert.Checked));
           TListeFichier.Enabled := (BNetExpert.Checked);
           FLISTEFICHIER.Enabled := (BNetExpert.Checked);
       end;

       if (BNetExpert.Checked) then
       begin
            TNetRecep := TOB.Create('', nil, -1);
            TA := TOB.Create ('',TNetRecep,-1);
            if Uppercase(ExtractFileExt(FICHENAME.Text)) = '.NSV' then
            begin
                 TA.AddChampSupValeur('DOMAINE', 'PAIE');
                 TA.AddChampSupValeur('CLIID',  V_PGI.NoDossier);
                 TA.AddChampSupValeur('CLINOM', V_PGI.NomSociete);
                 TA.AddChampSupValeur('FICMSQ', 'NSV');
            end
            else
            begin
                 TA.AddChampSupValeur('DOMAINE', 'COMPTA');
                 TA.AddChampSupValeur('CLIID',  V_PGI.NoDossier);
                 TA.AddChampSupValeur('CLINOM', V_PGI.NomSociete);
                 TA.AddChampSupValeur('FICMSQ', 'TRA');
            end;
            OutTOBNet := NEListeFichier (TNetRecep, UneAct_ATTENTE, EnvoiRecep);
            if OutTOBNet.GetValue ('ERROR') = '0' then
            begin
                 FLISTEFICHIER.Items.clear;
                 if FLISTEFICHIER.Values <> nil then FLISTEFICHIER.Values.clear;
                 for i:= 0 to OutTOBNet.detail.count-1 do
                 begin
                     Fichier := OutTOBNet.detail[i].Getvalue ('DESCRIPTIF') + ' ' + OutTOBNet.detail[i].Getvalue ('FICNOM')+'.'+ OutTOBNet.detail[i].Getvalue ('FICMSQ');
                     FLISTEFICHIER.Items.Add(Fichier);
                     FLISTEFICHIER.Values.add(Fichier);
                 end;
            end
            else
                    PGIBox (OutTOBNet.GetValue ('ERRORLIB'));
            if TNetRecep <> nil then TNetRecep.free;
       end;
       if stArg = '' then  // sans ligne de commande
        Label15.Caption := ''; // fiche   10581
  // pour international
  DATEECR1.Text := stDate1900;
  ObjetFiltre.Charger;
end;


procedure TFAssistImport.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
    bSuivant.Enabled := FALSE;
    bPrecedent.Enabled := FALSE;
    bFin.Enabled := FALSE;
    if (stArg <> '') or ( not OkCommit) then
    begin
      if OKexercice then CommitTrans
      else RollBackDiscret ;
    end;
     //ajout optimisation du dossier à faire uniquement dans le cas de SQL server
If IsMSSQL then  //10/08/2006 YMO Ajout test SQL2005
    begin
{$IFNDEF EAGLCLIENT}
          if (NatureFichier = 'DOS') and (OrigineFichier <> 'IMMO') then  // pour vider le .LOG
          begin
                    try
                          ExecuteSql ('checkpoint');
                          ExecuteSql ('backup log '+v_pgi.DBName+' with no_log');
                          ExecuteSql ('dbcc shrinkdatabase('+v_pgi.DBName+')');
                          ExecuteSql ('dbcc shrinkfile(2)');
                    except;
                    end;
          end;
{$ENDIF}

    end;
  if Assigned(ObjetFiltre) then FreeAndNil(ObjetFiltre);
  // à porter sur V7XX
  ExecuteSQL('DELETE FROM COURRIER WHERE MG_UTILISATEUR="'+W_W+'" AND MG_TYPE=5555 AND MG_EXPEDITEUR="'+V_PGI.User+'"') ;

{$IFDEF COMPTA}
{$IFDEF EAGLCLIENT}
    AvertirCacheServer( 'PARAMSOC' ) ; // fiche  10546
    AvertirCacheServer( 'EXERCICE' ) ;
    AvertirCacheServer('SOUCHE');
{$ENDIF}
    AvertirMultiTable('TTEXERCICE');  // fiche 10572
    AvertirMultiTable('TTJOURNAL');
    CHARGEMAGEXO (FALSE);

{$ENDIF}

end;



procedure TFAssistImport.BTWINZIPClick(Sender: TObject);
var
  FileName       : String ;
  TheToz         : TOZ;
  Password       : string;
  Filearchive    : string;
  FileDir        : string;
  FileTraite     : string;
  Namefichier    : string;
{$IFDEF EAGLCLIENT}
  ii             : integer;
{$ENDIF}
begin
  inherited;
  Password := '';
    // Récupération du nom du fichier a insérer
    //
  FileName := FICHENAME.Text;
  Namefichier := ExtractFileName(FileName);
  Filearchive := ExtractFileDir(FileName) +  '\' + ReadTokenPipe (Namefichier, '.');
  // fiche 10450 Filearchive := ReadTokenPipe (Filename, '.');

  FileTraite  := Filearchive + '.TRA';
  Filearchive := Filearchive + '.zip';
  if UpperCase(ExtractFileExt(FileName)) <> '.ZIP' then exit;

  TheToz := TOZ.Create ;
  try
    if TheToz.OpenZipFile ( Filearchive, moOpen ) then
    begin
        TobFich := TheToz.ConvertListInTob;
{$IFDEF EAGLCLIENT}
        if (TobFich.Detail.Count > 0) then
        begin
         FileDir := ExtractFileDir ( FICHENAME.Text ) ;
         for ii :=0 to TobFich.detail.count-1 do
         begin
             FileTraite := TobFich.detail[ii].getvalue('ZF_NAME');
             if (pos('.TRA', FileTraite) <> 0) or
             (pos('.tra', FileTraite) <> 0)  or
             (pos('.CPTA', FileTraite) <> 0) or
             (pos('.cpta', FileTraite) <> 0) then
             begin
                  if TheToz.OpenSession(osExt) then
                  begin
                       sysutils.DeleteFile(FileDir + '\' + ExtractFileName(FileTraite));
                       if ((TheToz.SetDirOut(FileDir)) and
                           (TheToz.ProcessFile(FileTraite)) and
                           (TheToz.CloseSession())) then
                          begin
                               FICHENAME.Text := FileDir + '\' + ExtractFileName(FileTraite);
                               break;
                          end;
                  end;
             end;
         end;
         TobFich.free;
         TheToz.Free;  TheToz := nil;
        end;
{$ELSE}
        if TheToz.OpenSession ( {osExtAll} osExt ) then
        begin
        FileDir := ExtractFileDir ( FICHENAME.Text ) ;
        if TheToz.SetDirOut ( FileDir ) then
          begin
            TheToz.CloseSession ;
            FICHENAME.Text := FileTraite;
          end
          else
          begin
          HShowMessage ( '0;Erreur;Le répertoire ' + FileDir + ' n''existe pas.;E;O;O;O', '', '' ) ;
          TheToz.CancelSession ;
          end ;
        end
        else
        HShowMessage ( '0;Erreur;Soit le fichier : ' + ExtractFileName ( FileName ) + ' n''existe plus, soit la session n''est pas ouverte en ajoute.;E;O;O;O', '', '' ) ;
{$ENDIF}
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
      TheToz.Free;
      end ;
 END ;
 if TheToz <> nil then TheToz.free;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 30/04/2002
Modifié le ... :   /  /
Description .. : Function qui fait une copy de la chaine à la longueur
Suite ........ : donnée + enlève les blancs et si la valeur et numérique et à
Suite ........ : blanc met la valeur par défaut passée au niveau des
Suite ........ : paramètres de la fonction
Mots clefs ... :
*****************************************************************}
function TFAssistImport.EnleveBlanc(Chaine: string; x, y: integer ; ValNumeric : String = ''): string;
var
  st: string;
  IsNumeric : Boolean ;
begin
  IsNumeric:=ValNumeric<>'' ;
  st := Copy(Chaine, x, y);
  If IsNumeric Then St:=StStrFPoint(St) ;
  Result := Trim(St);
  If (Result='') And (ValNumeric<>'') Then Result:=ValNumeric ;
end;

Function TFAssistImport.Controlparamsociete  (SLect : string) : Boolean;
var
Q         : TQuery;
Cpte      : string;
i         : integer;
Deb       : string;
Oksoc     : Boolean;
lg1,lg2   : integer;
W1,W2     : string;
indd,indf : integer;
lgsect1   : integer;
lgsect2   : integer;
lgsect3   : integer;
lgsect4   : integer;
lgsect5   : integer;
//RetourConf: string;

begin
    OkSoc := TRUE; indd := 0; indf := SizeRecControl;
    Deb := EstUneLigneCom(SLect);
    if Copy (SLect,0, 6) = Deb+'PS2' then
    begin
            if (EnleveBlanc(SLect, 7, 2,'0') = '') or (EnleveBlanc(SLect, 10, 2,'0') = '') then
            begin
               if stArg <> '' then
                begin
                  AfficheListeCom(ERR_LGCPTEPS2,LISTEIMPORT);
                  EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                end
                else
                PGIBox (ERR_LGCPTEPS2, 'Import');
                OkSoc := FALSE; Result := OkSoc; exit;
            end;
            Lg1 :=  StrToInt(EnleveBlanc(SLect, 7, 2,'10'));
            Lg2 :=  StrToInt(EnleveBlanc(SLect, 10, 2,'10'));
            Changementgene := FALSE; Changementaux := FALSE;
                  // pour recharger s'il y a eu un changement de lg
            VH^.Cpta[fbGene].Lg:=GetParamSocSecur('SO_LGCPTEGEN', 0, TRUE) ;
            VH^.Cpta[fbAux].Lg:=GetParamSocSecur('SO_LGCPTEAUX', 0, TRUE) ;

            // longueur de section
            if GetParamSocSecur ('SO_ZGEREANAL', TRUE) = TRUE then
            begin
                 lgsect1 := StrToInt(EnleveBlanc(SLect, 13, 2,'0'));
                 lgsect2 := StrToInt(EnleveBlanc(SLect, 16, 2,'0'));
                 lgsect3 := StrToInt(EnleveBlanc(SLect, 19, 2,'0'));
                 lgsect4 := StrToInt(EnleveBlanc(SLect, 22, 2,'0'));
                 lgsect5 := StrToInt(EnleveBlanc(SLect, 25, 2,'0'));
               if (NatureFichier = 'JRL') then
               begin
                  if(VH^.Cpta[fbAxe1].Lg <> lgsect1) and (VH^.Cpta[fbAxe1].Lg <> 0) or
                  (VH^.Cpta[fbAxe2].Lg <> lgsect2) and (VH^.Cpta[fbAxe2].Lg <> 0) or
                  (VH^.Cpta[fbAxe3].Lg <> lgsect3) and (VH^.Cpta[fbAxe3].Lg <> 0) or
                  (VH^.Cpta[fbAxe4].Lg <> lgsect4) and (VH^.Cpta[fbAxe4].Lg <> 0) or
                  (VH^.Cpta[fbAxe5].Lg <> lgsect5) and (VH^.Cpta[fbAxe5].Lg <> 0) then
                    begin
                       if stArg <> '' then
                        begin
                          AfficheListeCom(ERR_SECTIONANA,LISTEIMPORT);
                          EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                        end
                        else
                         PGIBox ('Enregistrement PS2 :#10#13la longueur des sections analytiques est incohérente avec le paramétrage du dossier.', 'Import');
                         OkSoc := FALSE; Result := OkSoc; exit;
                     end;
               end;
               if (NatureFichier = 'SYN')then
               begin
                  if(VH^.Cpta[fbAxe1].Lg <> lgsect1) and (VH^.Cpta[fbAxe1].Lg <> 0) then
                  begin
                      if stArg <> '' then
                        begin
                          AfficheListeCom(ERR_SECTIONANA, LISTEIMPORT);
                          EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                        end
                        else
                          PGIBox ('Enregistrement PS2 :#10#13la longueur des sections analytiques est incohérente avec le paramétrage du dossier.', 'Import');
                         OkSoc := FALSE; Result := OkSoc; exit;
                  end;
               end;
            end;
            if (VH^.Cpta[fbGene].Lg <> lg1) and (VH^.Cpta[fbGene].Lg <> 0) or
             (VH^.Cpta[fbAux].Lg <> lg2) and (VH^.Cpta[fbAux].Lg <> 0) then
            begin
               if (NatureFichier = 'JRL') or (NatureFichier = 'SYN')then
               begin
                      if stArg <> '' then
                        begin
                          AfficheListeCom(ERR_LGCPTE,LISTEIMPORT);
                          EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                        end
                        else
                        PGIBox ('Enregistrement PS2 :#10#13la longueur des comptes est incohérente avec le paramétrage du dossier.', 'Import');
                   OkSoc := FALSE; Result := OkSoc; exit;
               end;
               if ExisteSql ('SELECT EX_EXERCICE FROM EXERCICE') then
               begin
                    if stArg <> '' then
                    begin
                      AfficheListeCom('La longueur des comptes est différente',LISTEIMPORT);
                      OkSoc := FALSE; Result := OkSoc;
                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                      exit;
                    end
                    else
                    begin
                       If PGIAsk('La longueur des comptes est différente; Voulez-vous effectuez une RAZ des écritures dossier ?','Import')=mrYes Then
                       begin
                           (*RetourConf := AglLanceFiche('CP','IMPORTCOMCONF','', '',
                           'Origine du fichier : '+ OrigineFichier + '; Nature : ' + NatureFichier
                           +'; Fichier : '+ FICHENAME.Text+';RAZ')  ;
                           if RetourConf <> 'X' then
                           begin
                              OkSoc := FALSE; Result := OkSoc; exit;
                           end
                           else
                           begin
                                SuppComptable.Checked := FALSE;
                           *)
                                if  VH^.Cpta[fbGene].Lg > lg1 then
                                begin
                                  lg1 := VH^.Cpta[fbGene].Lg;
                                  AfficheListeCom('La longueur des comptes généraux est supérieure à la longueur du fichier.',LISTEIMPORT);
                                  AfficheListeCom('Pour la modifier, utilisez l''option changement de longueur des comptes.',LISTEIMPORT);
                                end
                                else
                                begin
                                     if VH^.Cpta[fbGene].Lg < lg1 then
                                     begin
                                          ChangeLgCpteGen ( lg1, VH^.Cpta[fbgene].Cb, nil);
                                          VH^.Cpta[fbGene].Lg := lg1;
                                          Changementgene := TRUE;
                                     end;
                                end;

                                 if VH^.Cpta[fbAux].Lg > lg2 then
                                 begin
                                      lg2 :=  VH^.Cpta[fbAux].Lg;
                                      AfficheListeCom('La longueur des comptes auxiliaires est supérieure à la longueur du fichier.',LISTEIMPORT);
                                      AfficheListeCom('Pour la modifier, utilisez l''option changement de longueur des comptes.',LISTEIMPORT);
                                 end
                                 else
                                 begin
                                      if VH^.Cpta[fbAux].Lg < lg2 then
                                      begin
                                           ChangeLgCpteAux (lg2, VH^.Cpta[fbAux].Cb, nil);
                                           VH^.Cpta[fbAux].Lg := lg2;
                                           Changementaux := TRUE;
                                      end;
                                 end;
                           //end;
                       end
                       else
                       begin
                               OkSoc := FALSE; Result := OkSoc; exit;
                       end;
                    end;
               end;

            end;

            if (EnleveBlanc(SLect, 13, 2,'0') = '') or (EnleveBlanc(SLect, 16, 2,'0') = '')
            or (EnleveBlanc(SLect, 19, 2,'0') = '') or (EnleveBlanc(SLect, 22, 2,'0') = '')
            or (EnleveBlanc(SLect, 25, 2,'0') = '') then
            begin
                      if stArg <> '' then
                        begin
                          AfficheListeCom(ERR_LGSECTION, LISTEIMPORT);
                          EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                        end
                        else
                        PGIBox (ERR_LGSECTION, 'Import');
                        OkSoc := FALSE; Result := OkSoc; exit;
            end;
            if VH^.Cpta[fbGene].Lg = 0 then
            begin
                 VH^.Cpta[fbGene].Lg := Lg1;
                 SetParamsoc('SO_LGCPTEGEN', lg1);
            end;
            if VH^.Cpta[fbAux].Lg = 0 then
            begin
                 VH^.Cpta[fbAux].Lg := Lg2;
                 SetParamsoc('SO_LGCPTEAUX', lg2);
            end;
            if (VH^.Cpta[fbGene].Cb = '') then
               SetParamsoc('SO_BOURREGEN', EnleveBlanc(SLect, 9, 1,'0'));
            if (VH^.Cpta[fbAux].Cb = '') then
               SetParamsoc('SO_BOURREAUX', EnleveBlanc(SLect, 12, 1,'0'));

    end;
    if (Copy (SLect,0, 6) = Deb+'PS2') or (Copy (SLect,0, 6) = Deb+'PS3') then
    begin
        if (Copy (SLect,0, 6) = Deb+'PS2') then begin indd := 0; indf := 4;  end;
        if (Copy (SLect,0, 6) = Deb+'PS3') then begin indd := 5; indf := SizeRecControlPS3; end;

        for i:=indd to indf do
        begin
             if EnleveBlanc(SLect, (RecControl[i].deb), (RecControl[i].fin),'') = '' then
             begin
                 if GetParamSocSecur (RecControl[i].NomChamp, '') = '' then
                 begin
                    Q := OpenSQL ('SELECT PRR_NUMPLAN,PRR_SOCNOM,PRR_SOCDATA from PARSOCREF where ' +
                      ' PRR_NUMPLAN =7 and PRR_SOCNOM="'+RecControl[i].NomChamp+'"', True);
                    if not Q.EOF then
                    begin
                           if RecControl[i].champ = 'CPT' then
                           begin
                               Cpte := BourreOuTronque(Q.Findfield ('PRR_SOCDATA').asstring, RecControl[i].LeFb);
                               SetParamsoc (RecControl[i].NomChamp, Cpte);
                               RecControl[i].Valeur := Cpte;
                               if RecControl[i].LeFb = fbGene then
                               begin
                                    W1 := ' AND GER_GENERAL="'+ Q.Findfield ('PRR_SOCDATA').asstring+'"';
                                    W2 := ' Where G_GENERAL="'+Cpte+'"';
                                    Ferme(Q);
                                    LoadStandardCompta (7, 'GENERAUX','GENERAUXREF', W2, W1);
                               end;
                               if RecControl[i].LeFb = fbAux then
                               begin
                                    W1 := ' AND TRR_AUXILIAIRE="'+ Q.Findfield ('PRR_SOCDATA').asstring+'"';
                                    W2 := ' Where T_AUXILIAIRE="'+Cpte+'"';
                                    Ferme(Q);
                                    LoadStandardCompta (7, 'TIERS','TIERSREF', W2, W1);
                               end;
                           end;
                           if RecControl[i].champ = 'JAL' then
                           begin
                                    SetParamsoc (RecControl[i].NomChamp, Q.Findfield ('PRR_SOCDATA').asstring);
                                    W1 := ' AND JR_JOURNAL="'+ Q.Findfield ('PRR_SOCDATA').asstring+'"';
                                    W2 := ' Where J_JOURNAL="'+Q.Findfield ('PRR_SOCDATA').asstring+'"';
                                    Ferme(Q);
                                    LoadStandardCompta (7, 'JOURNAL','JALREF', W2, W1);
                           end;
                    end
                    else
                    ferme (Q);
                 end
                 else
                 begin
                  if RecControl[i].champ = 'CPT' then
                           RecControl[i].Valeur := GetParamSocSecur (RecControl[i].NomChamp, '');
                  if RecControl[i].champ = 'JAL' then
                           RecControl[i].Valeur := GetParamSocSecur (RecControl[i].NomChamp, '');
                 end;
             end
             else
             begin
                  if RecControl[i].champ = 'CPT' then
                  begin
                           Cpte := BourreOuTronque(EnleveBlanc(SLect, (RecControl[i].deb), (RecControl[i].fin),''), RecControl[i].LeFb);
                           RecControl[i].Valeur := Cpte;
                           CoherenceCompte (i, '');
                  end;
                  if RecControl[i].champ = 'JAL' then
                           RecControl[i].Valeur := EnleveBlanc(SLect, (RecControl[i].deb), (RecControl[i].fin));
             end;
        end;
    end;
    // ajout V590
    if GetParamSocSecur ('SO_GCMODEREGLEDEFAUT', '') = '' then
    begin
             Q := OpenSQL ('SELECT PRR_NUMPLAN,PRR_SOCNOM,PRR_SOCDATA from PARSOCREF where ' +
             ' PRR_NUMPLAN =7 and PRR_SOCNOM="SO_GCMODEREGLEDEFAUT"', True);
             if not Q.EOF then SetParamsoc ('SO_GCMODEREGLEDEFAUT', Q.Findfield ('PRR_SOCDATA').asstring);
             ferme (Q);
    end;
    if GetParamSocSecur ('SO_REGIMEDEFAUT', '') = '' then
    begin
             Q := OpenSQL ('SELECT PRR_NUMPLAN,PRR_SOCNOM,PRR_SOCDATA from PARSOCREF where ' +
             ' PRR_NUMPLAN =7 and PRR_SOCNOM="SO_REGIMEDEFAUT"', True);
             if not Q.EOF then SetParamsoc ('SO_REGIMEDEFAUT', Q.Findfield ('PRR_SOCDATA').asstring);
             ferme (Q);
    end;
    if GetParamSocSecur ('SO_CODETVADEFAUT', '') = '' then
    begin
             Q := OpenSQL ('SELECT PRR_NUMPLAN,PRR_SOCNOM,PRR_SOCDATA from PARSOCREF where ' +
             ' PRR_NUMPLAN =7 and PRR_SOCNOM="SO_CODETVADEFAUT"', True);
             if not Q.EOF then SetParamsoc ('SO_CODETVADEFAUT', Q.Findfield ('PRR_SOCDATA').asstring);
             ferme (Q);
    end;

Result := OkSoc;
end;

Function TFAssistImport.ControlParam5 (Ligne : string; Existeexerc : Boolean) : Boolean;
var
Okp   : Boolean;
i     : integer;
Cpte  : string;
Q1    : TQuery;
begin
     okp := TRUE;
     if (EnleveBlanc(Ligne, 7, 3,'') = '') or (EnleveBlanc(Ligne, 10, 1,'') = '')
     or (EnleveBlanc(Ligne, 11, 1,'') = '') then
     begin
        if stArg <> '' then
          begin
            AfficheListeCom(ERR_PARAMPS5, LISTEIMPORT);
            EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
          end
          else
          PgiBox ('Enregistrement PS5 :#10#13vous devez renseigner tous les paramètres obligatoires dans le fichier.', 'Import');
          okp := FALSE; Result := okp; exit;
     end;
     if Existeexerc then
     begin
         if (GetParamSocSecur ('SO_DEVISEPRINC', '') <> '') and ((GetParamSocSecur ('SO_DEVISEPRINC', '')) <> (EnleveBlanc(Ligne, 7, 3,''))) then
         begin
            if stArg <> '' then
              begin
                AfficheListeCom(ERR_MONNAIEPS5, LISTEIMPORT);
                EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
              end
              else
              PgiBox ('Enregistrement PS5 :#10#13la monnaie de tenue est incohérente avec le paramétrage du dossier.', 'Import');
              okp := FALSE; Result := okp; exit;
         end;
     end;
     if (EnleveBlanc(Ligne, 7, 3,'') = 'EUR') and (EnleveBlanc(Ligne, 11, 1,'') <> 'X') then
     begin
          if stArg <> '' then
          begin
                AfficheListeCom(ERR_MONNAIETENUE,LISTEIMPORT);
                EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
          end
          else
          PgiBox (ERR_MONNAIETENUE, 'Import');
          okp := FALSE; Result := okp; exit;
     end;
     if (EnleveBlanc(Ligne, 7, 3,'') <> 'EUR') and (EnleveBlanc(Ligne, 11, 1,'') = 'X') then
     begin
         if stArg <> '' then
          begin
                AfficheListeCom(ERR_PARAMTENUE, LISTEIMPORT);
                EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
          end
          else
           PgiBox (ERR_PARAMTENUE, 'Import');
          okp := FALSE; Result := okp; exit;
     end;

    for i:= SizeRecControlPS3+1 to  SizeRecControl do
    begin
              if RecControl[i].champ = 'CPT' then
              begin
                       Cpte := BourreOuTronque(EnleveBlanc(Ligne, (RecControl[i].deb), (RecControl[i].fin),''), RecControl[i].LeFb);
                       RecControl[i].Valeur := Cpte;
                       CoherenceCompte (i, '');
              end;
              if RecControl[i].champ = 'JAL' then
                       RecControl[i].Valeur := EnleveBlanc(Ligne, (RecControl[i].deb), (RecControl[i].fin));
              if RecControl[i].champ = 'DIV' then
              begin
                       RecControl[i].Valeur := EnleveBlanc(Ligne, (RecControl[i].deb), (RecControl[i].fin));
                       if GetParamsocSecur (RecControl[i].NomChamp, '') = '' then
                       begin
                            Q1 := OpenSQL ('SELECT PRR_NUMPLAN,PRR_SOCNOM,PRR_SOCDATA from PARSOCREF where ' +
                              ' PRR_NUMPLAN =7 and PRR_SOCNOM="'+RecControl[i].NomChamp+'"', True);
                            if not Q1.EOF then
                               SetParamsoc (RecControl[i].NomChamp, Q1.Findfield ('PRR_SOCDATA').asstring);
                            ferme (Q1);
                       end;
              end;

    end;
    Result := okp;
end;

function TFAssistImport.COkExercice (SLect : string) : Boolean;
var
D1, D2   : TDateTime;
Q1       : TQuery;
PremJour, PremMois, PremAnnee, DernAnnee, DernMois, DernJour : Word ;
PremJourF, PremMoisF, PremAnneeF, DernAnneeF, DernMoisF, DernJourF : Word ;
  procedure afficheerreur ( option : integer);
  begin
                      if stArg <> '' then
                      begin
                        case option of
                          1 : AfficheListeCom(ERR_DATEEXEXER, LISTEIMPORT);
                          0 : AfficheListeCom(ERR_EXERCDECALE,LISTEIMPORT);
                        end;
                        EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                      end
                      else
                      begin
                          case option of
                            1 : PGIBox (ERR_DATEEXEXER, 'Import');
                            0 : PGIBox (ERR_EXERCDECALE, 'Import');
                          end;

                      end;
  end;
begin
        Result := FALSE;
        D1 := Format_Date_HAL(EnleveBlanc(SLect, 10, 8,''));
        D2 := Format_Date_HAL(EnleveBlanc(SLect, 18, 8,''));

        if (NatureFichier = 'DOS') then  // pour dossier avec les immos  // Fiche 10423
        begin
                  if ExisteSQl('SELECT EX_DATEDEBUT, EX_DATEFIN from EXERCICE where EX_DATEDEBUT="'
                  + UsDateTime(D1) +'" and EX_DATEFIN="'+ UsDateTime(D2) + '"') then
                  begin
                       Result := TRUE; exit;
                  end
                  else
                  begin
                       PGIBox (ERR_DATEEXEXER + 'La réception du fichier va générer une incohérence sur les immobilisations.');
                       Result := FALSE; exit;
                  end;
        end
        else
        begin
                if ExisteSQl('SELECT EX_DATEDEBUT, EX_DATEFIN from EXERCICE where EX_DATEDEBUT="'
                + UsDateTime(D1) +'" and EX_DATEFIN="'+ UsDateTime(D2) + '"') then
                begin
                     Result := TRUE; exit;
                end;

          (*  ne marche pas en cas synchro à voir exemple 2001,2002,2003,2004 dans fichier et la base de
          reception uniquement elle a 2004 pb voir avec CA  *)
                if VerifCoherenceExo (D1, D2) then
                begin
                     Result := TRUE; exit;
                end
                else
                begin   // Fiche 10186 modification date fin en synchro
                     if ExisteSQl('SELECT EX_DATEDEBUT, EX_DATEFIN from EXERCICE where EX_DATEDEBUT="'
                        + UsDateTime(D1) +'" and EX_DATEFIN<>"'+ UsDateTime(D2) + '"') then
                        begin
                              afficheerreur (1);
                              OKexercice := FALSE;
                              Result := FALSE; exit;
                        end;
                end;

                DecodeDate(D1,PremAnnee,PremMois,PremJour) ;
                DecodeDate(D2,DernAnnee,DernMois,DernJour) ;

                Q1 := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN from EXERCICE ', True);
                While not Q1.Eof do
                begin

                     if (D1 > Q1.FindField('EX_DATEDEBUT').AsDateTime) and
                        (D1 <= Q1.FindField('EX_DATEFIN').AsDateTime) and
                        (Q1.FindField('EX_DATEFIN').AsDateTime <= D2) then
                     begin
                        afficheerreur (0);
                        OKexercice := FALSE; Ferme (Q1);
                        exit;
                     end;
                    DecodeDate(Q1.FindField('EX_DATEDEBUT').AsDateTime,PremAnneeF,PremMoisF,PremJourF) ;
                    DecodeDate(Q1.FindField('EX_DATEFIN').AsDateTime,DernAnneeF,DernMoisF,DernJourF) ;
                    if PremAnneeF = PremAnnee then
                    begin
                         if PremMoisF <> PremMois  then
                         begin
                              afficheerreur(1);
                              OKexercice := FALSE; Ferme (Q1);
                              exit;
                         end;
                    end;
                    if DernAnneeF = DernAnnee then
                    begin
                         if DernMoisF <> DernMois then
                         begin
                              afficheerreur(1);
                              OKexercice := FALSE; Ferme (Q1);
                              exit;
                         end;
                    end;

                    Q1.next;
                end;
                Ferme (Q1);
                Q1 := OpenSQL('SELECT MAX(EX_DATEFIN) Datefin from EXERCICE '+
                ' Where EX_DATEFIN < "'+ USDATETime(D1)+'"', True);
                if not Q1.Eof then
                begin
                     if (Q1.FindField('Datefin').asDatetime <> 0) and
                      (Q1.FindField('Datefin').asDatetime+1 <> D1) then
                     begin
                            afficheerreur(0);
                            OKexercice := FALSE;   Ferme (Q1);
                            exit;
                     end;
                end;
                Ferme (Q1);
        end;
        Result := TRUE;
end;

Function TFAssistImport.NbExoouv(SS : string; var Nb : integer): Boolean;
var
Q1       : TQuery;
D1,D2    : TDatetime;
Nbexo    : integer;
begin
        NbExo := 0;
        if (NatureFichier = 'BAL') or (NatureFichier = 'JRL') then
           Nbexo := NombreExercice ('WHERE EX_ETATCPTA="OUV"');
        if Nbexo = 0 then begin Result := TRUE; exit; end;
        D1 := Format_Date_HAL(EnleveBlanc(SS, 10, 8,''));
        D2 := Format_Date_HAL(EnleveBlanc(SS, 18, 8,''));
        Q1 := OpenSQl ('SELECT EX_ETATCPTA,EX_EXERCICE,EX_DATEDEBUT,EX_DATEFIN FROM EXERCICE '+
        'WHERE EX_DATEDEBUT="'+UsDateTime(D1)+'" AND EX_DATEFIN="'+UsDateTime(D2)+'"',TRUE);
        if not Q1.EOF then
        begin
           if (Q1.FindField ('EX_ETATCPTA').asstring <> 'OUV') and
            (EnleveBlanc(SS, 26, 3,'') = 'OUV') then inc (Nb) ;
        end
        else if (EnleveBlanc(SS, 26, 3,'') = 'OUV') then inc (Nb);
         ferme (Q1);
        if (Nbexo+Nb > 2) then
        begin
                    if stArg <> '' then
                    begin
                       AfficheListeCom(ERR_EXOPLUSOUV,LISTEIMPORT);
                       EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT)
                    end
                    else
                       PgiBox (ERR_EXOPLUSOUV, 'Import');
                OKexercice := FALSE;
                Result := FALSE; exit;
        end;

        Result := TRUE;
end;

Function TFAssistImport.ControlEnreg (SS,Enreg : string; deb : integer; fin : integer): Boolean;
var
Oke      : Boolean;
ie       : integer;
Dd,nat   : string;
begin
     Oke := TRUE;
     Dd := EstUneLigneCom(SS);
     for ie := deb to fin do
     begin
             if EnleveBlanc(SS, (EnregControl[ie].deb), (EnregControl[ie].fin),'') = '' then
             begin
                                // pour ne pas bloquer le problème de création compte avec champ venti; à blanc
                                if (Enreg = 'Comptes généraux : ' + GetParamSocSecur ('SO_GENATTEND', '')) then
                                begin
                                 if (EnregControl[ie].mess=ERR_GENERAUX_VENTIL1) or
                                 (EnregControl[ie].mess=ERR_GENERAUX_VENTIL2) or
                                 (EnregControl[ie].mess=ERR_GENERAUX_VENTIL3) or
                                 (EnregControl[ie].mess=ERR_GENERAUX_VENTIL4) or
                                 (EnregControl[ie].mess=ERR_GENERAUX_VENTIL5)then continue;
                                end;
                                if stArg <> '' then
                                begin
                                      AfficheListeCom(Enreg + ' ' + EnregControl[ie].mess,LISTEIMPORT);
                                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                begin
                                     if (Copy (SS,0, 6) = '***CAE') then
                                     begin
                                         nat := Copy (SS, 59, 3);
                                         if (nat <> 'PRO') and (nat <> 'AUC') and
                                         (nat <> 'AUD') and (nat <> 'CON') and
                                         (EnregControl[ie].mess=ERR_GENERAUX_CDPOINT) then
                                          PgiBox ('Code Collectif manquant (pos 63)', Enreg)
                                          else
                                          continue;
                                     end
                                     else
                                        PgiBox (EnregControl[ie].mess, Enreg);
                                end;
                                oke := FALSE; break;
             end
             else
             begin
                 if (Copy (SS,0, 6) = Dd+'CGE')
                 or (Copy (SS,0, 6) = Dd+'CGN')then
                 begin
                   if (ie >= 22) and (ie <= SizeEnregControl) then
                   begin
                        if (EnleveBlanc(SS, (EnregControl[ie].deb), (EnregControl[ie].fin),'') <> 'X') and
                        (EnleveBlanc(SS, (EnregControl[ie].deb), (EnregControl[ie].fin),'') <> '-') then
                        begin
                               if stArg <> '' then
                                begin
                                      AfficheListeCom(Enreg + ' ' + EnregControl[ie].mess,LISTEIMPORT);
                                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                PgiBox (EnregControl[ie].mess, Enreg);
                                oke := FALSE; break;
                        end;
                   end;
                    if (ie = 21)  then // lettrage
                    begin   // fiche  10347
                      if (EnleveBlanc(SS, 59, 3,'') <> 'DIV') and (EnleveBlanc(SS, 59, 3,'') <> 'TIC') and (EnleveBlanc(SS, 59, 3,'') <> 'TID')
                      and (EnleveBlanc(SS, 62, 1,'') = 'X') then
                      begin
                           // ajout me 14-05-2002 pour les cpte gene lettrable
                          if (EnleveBlanc(SS, 7, 1,'') <> '4') and
                          (EnleveBlanc(SS, 7, 3,'') <> '511') and
                          (EnleveBlanc(SS, 7, 2,'') <> '58') then
                          begin
                              if (OrigineFichier = 'S1') then
                                 AfficheListeCom(Enreg+ ' ne doit pas être lettrable : '+
                                 '(Modifié par ComSx)',LISTEIMPORT)
                              else
                              begin
                               if stArg <> '' then
                                begin
                                      AfficheListeCom('Ce compte ne doit pas être lettrable',LISTEIMPORT);
                                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                PgiBox ('Ce compte ne doit pas être lettrable', Enreg);
                                oke := FALSE; break;
                              end;
                          end;
                      end;
                     end;
                 end;
             end;
     end;
Result := Oke;
end;

Function TFAssistImport.ControleAutreParam (Deb,SLect : string) : Boolean;
var
ii       : integer;
nat      : string;
begin
Result := FALSE;
        if (Copy (SLect,0, 5) = Deb+'TL') then
          if not ControlEnreg (SLect,'Tables libre', 4, 6) then exit;
        if Copy (SLect,0, 6) = Deb+'SSA' then
          if not ControlEnreg (SLect,'Sous-Section', 7, 9) then exit;
        if Copy (SLect,0, 6) = Deb+'SAT' then
          if not ControlEnreg (SLect,'Section', 7, 9) then exit;
        if Copy (SLect,0, 6) = Deb+'ETB' then
          if not ControlEnreg (SLect,'Etablissement', 10, 10) then exit;
        if Copy (SLect,0, 6) = Deb+'MDP' then
          if not ControlEnreg (SLect,'Mode de paiement', 10, 10) then exit;
        if Copy (SLect,0, 6) = Deb+'MDR' then
          if not ControlEnreg (SLect,'Mode de règlement', 10, 10) then exit;
        if Copy (SLect,0, 6) = Deb+'REG' then
          if not ControlEnreg (SLect,'Régime TVA', 10, 10) then exit;
        if Copy (SLect,0, 6) = Deb+'TVA' then
          if not ControlEnreg (SLect,'Information sur TVA', 10, 10) then exit;
        if Copy (SLect,0, 6) = Deb+'SOU' then
          if not ControlEnreg (SLect,'Souche', 10, 11) then exit;
        if Copy (SLect,0, 6) = Deb+'JAL' then
        begin
          if not ControlEnreg (SLect,'Journal', 10, 11) then exit;
          nat := EnleveBlanc(SLect, 45, 3);
          if (nat <> 'ACH') and (nat <> 'ANA') and
             (nat <> 'ANO') and (nat <> 'BQE') and
             (nat <> 'CAI') and (nat <> 'CLO') and
             (nat <> 'ECC') and (nat <> 'EXT') and
             (nat <> 'OD') and (nat <> 'ODA') and
             (nat <> 'REG') and (nat <> 'VTE') then
             begin
                               if stArg <> '' then
                                begin
                                      AfficheListeCom('Nature Journal incorrecte :'+EnleveBlanc(SLect, 45, 3),LISTEIMPORT);
                                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                PGIBox ('Nature Journal incorrecte :'+EnleveBlanc(SLect, 45, 3), 'Journal :' +  EnleveBlanc(SLect, 7, 3));
                                exit;
             end;
        end;
        if Copy (SLect,0, 6) = Deb+'DEV' then
          if not ControlEnreg (SLect,'Devise', 12, 17) then exit;
        if (Copy (SLect,0, 6) = Deb+'CGE')
        or (Copy (SLect,0, 6) = Deb+'CGN') then
        begin
          if not ControlEnreg (SLect,'Comptes généraux : ' + EnleveBlanc(SLect, 7, 17), 18, 26) then exit;
          for ii:=0 to SizeRecControl do
               if not CoherenceCompte (ii, SLect) then exit;
          nat := EnleveBlanc(SLect, 59, 3);
          if (nat <> 'BQE') and (nat <> 'CAI') and
             (nat <> 'COC') and (nat <> 'CHA') and
             (nat <> 'COD') and (nat <> 'COF') and
             (nat <> 'DIV') and (nat <> 'COS') and
             (nat <> 'EXT') and (nat <> 'IMO') and
             (nat <> 'TIC') and (nat <> 'PRO')
             and (nat <> 'TID') then
             begin
                               if stArg <> '' then
                                begin
                                      AfficheListeCom('Nature Compte incorrecte :'+ nat+ ' Compte :' +  EnleveBlanc(SLect, 7, 17),LISTEIMPORT);
                                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                PGIBox ('Nature Compte incorrecte :'+nat, 'Compte :' +  EnleveBlanc(SLect, 7, 17));
                                exit;
             end;

        end;
        if Copy (SLect,0, 6) = Deb+'CAE' then
        begin
          if not ControlEnreg (SLect,'Comptes auxiliaires : '+ EnleveBlanc(SLect, 7, 17), 18, 21) then exit;
          for ii:=1 to 3 do
               if not CoherenceCompte (ii, SLect) then exit;
          nat := EnleveBlanc(SLect, 59, 3);
          if (nat <> 'AUC') and (nat <> 'AUD') and
             (nat <> 'DIV') and (nat <> 'CLI') and
             (nat <> 'FOU') and (nat <> 'COF') and
             (nat <> 'SAL') and // les tiers comptables
             (nat <> 'PRO') and (nat <> 'AUC') and
             (nat <> 'AUD') and (nat <> 'CON') and
             (nat <> 'NCP') then // tiers autres
             begin
                               if stArg <> '' then
                                begin
                                      AfficheListeCom('Nature Compte incorrecte :'+nat+' Compte :' +  EnleveBlanc(SLect, 7, 17),LISTEIMPORT);
                                      EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
                                end
                                else
                                PGIBox ('Nature Compte incorrecte :'+nat, 'Compte :' +  EnleveBlanc(SLect, 7, 17));
                                exit;
             end;
        end;
        Result := TRUE;
end;

Function TFAssistImport.NombreExercice(WW : string; Max : Boolean=FALSE): integer;
var
Q1       : TQuery;
Nb       : integer;
begin
           if Max then
              Q1 := OpenSQl ('SELECT Max(EX_EXERCICE) FROM EXERCICE '+ WW,TRUE)
           else
              Q1 := OpenSQl ('SELECT Count(EX_EXERCICE) FROM EXERCICE '+ WW,TRUE);
           Nb := Q1.Fields[0].asinteger;
           ferme (Q1);
           Result := Nb;
end;

Function TFAssistImport.ControleParamSoc(indice: integer) : Boolean;
var
Qa       : Tquery;
procedure EcritureErreur (Mesg : string);
begin
     AfficheListeCom(Mesg,LISTEIMPORT);
     Result := FALSE;  OKexercice := FALSE;
end;
begin
  Result := TRUE;
  if not BControlParam.Checked then exit;
case  indice of
  1 : //ps2
  begin
        if (GetParamSocSecur ('SO_LGCPTEGEN', 0, TRUE) = 0  ) then
           EcritureErreur('Longueur des comptes généraux à 0');
        if (GetParamSocSecur ('SO_LGCPTEAUX', 0, TRUE) = 0  ) then
           EcritureErreur('Longueur des comptes auxiliaires à 0');
        if (GetParamSocSecur ('SO_BOURREGEN', 0, TRUE) = '' ) then
           EcritureErreur('Caractère de bourrage des comptes généraux incorrect');
        if (GetParamSocSecur ('SO_BOURREAUX', 0, TRUE) = '' ) then
           EcritureErreur('Caractère de bourrage des comptes auxiliaires incorrect');
        if GetParamSocSecur ('SO_ZGEREANAL', TRUE) = TRUE then
        begin
              Qa := Opensql('SELECT X_LONGSECTION,X_BOURREANA,X_SECTIONATTENTE from AXE',
                TRUE);
              while not Qa.EOF do
              begin
                if (Qa.findField('X_LONGSECTION').asinteger = 0) or
                   (Qa.findField('X_BOURREANA').asstring = '') or
                   (Qa.findField('X_SECTIONATTENTE').asstring = '') then
                   begin
                        AfficheListeCom(VERIF_AXE,LISTEIMPORT);
                        break;
                   end;
                   Qa.next;
              end;
              ferme(Qa);
        end;
        if (GetParamSocSecur ('SO_GENATTEND', '', TRUE) = '') then
           EcritureErreur('Compte d''attente général non renseigné');
        if (GetParamSocSecur ('SO_CLIATTEND', '', TRUE) = '') then
           EcritureErreur('Compte d''attente client non renseigné');
        if (GetParamSocSecur ('SO_FOUATTEND', '', TRUE) = '') then
           EcritureErreur('Compte d''attente fournisseur non renseigné');
        if (GetParamSocSecur ('SO_SALATTEND', '', TRUE) = '') then
           EcritureErreur('Compte d''attente salarié non renseigné');
        if (GetParamSocSecur ('SO_DIVATTEND', '', TRUE) = '') then
           EcritureErreur('Compte d''attente divers non renseigné');
  end;
  2 : //ps3
  begin
        if (GetParamSocSecur ('SO_OUVREBIL', '', TRUE) = '') then
           EcritureErreur('Compte ouverture bilan non renseigné');
        if (GetParamSocSecur ('SO_FERMEBIL', '', TRUE) = '') then
           EcritureErreur('Compte fermeture bilan non renseigné');
        if (GetParamSocSecur ('SO_RESULTAT', '', TRUE) = '')   then
           EcritureErreur('Compte résultat intermédiaire non renseigné');
        if (GetParamSocSecur ('SO_FERMEBEN', '', TRUE) = '') then
           EcritureErreur('Compte fermeture bénéfice non renseigné');
        if (GetParamSocSecur ('SO_FERMEPERTE', '', TRUE) = '') then
           EcritureErreur('Compte fermeture perte non renseigné');
        if (GetParamSocSecur ('SO_JALOUVRE', '', TRUE) = '') then
           EcritureErreur('Journal d''ouverture non renseigné');
        if (GetParamSocSecur ('SO_JALREPBALAN', '', TRUE) = '') then
           EcritureErreur('Journal de reprise balance non renseigné');
        if (GetParamSocSecur ('SO_DEFCOLCLI', '', TRUE) = '') then
           EcritureErreur('Compte collectif client non renseigné');
        if (GetParamSocSecur ('SO_DEFCOLFOU', '', TRUE) = '') then
           EcritureErreur('Compte collectif fournisseur non renseigné');
        if (GetParamSocSecur ('SO_DEFCOLSAL', '', TRUE) = '') then
           EcritureErreur('Compte collectif salarié non renseigné');
        if (GetParamSocSecur ('SO_DEFCOLDDIV', '', TRUE) = '') then
           EcritureErreur('Compte collectif divers débiteur non renseigné');
        if (GetParamSocSecur ('SO_DEFCOLCDIV', '', TRUE) = '') then
           EcritureErreur('Compte collectif divers créditeur non renseigné');
        if (GetParamSocSecur ('SO_DEFCOLDIV', '', TRUE) = '') then
           EcritureErreur('Compte collectif divers non renseigné');
  end;
  3 : //ps5
  begin
        if (GetParamSocSecur ('SO_DEVISEPRINC', '', TRUE) = '') then
           EcritureErreur('Devise principale non renseignée');
        if (GetParamSocSecur ('SO_DECVALEUR', 0, TRUE) = 0) then
           EcritureErreur('Nombres de décimales non renseigné');
        if (GetParamSocSecur ('SO_REGIMEDEFAUT', '', TRUE)= '') then
           EcritureErreur('Devise principale non renseignée');
        if (GetParamSocSecur ('SO_GCMODEREGLEDEFAUT', '', TRUE) = '') then
           EcritureErreur('Régime de facturation non renseigné');
        if (GetParamSocSecur ('SO_CODETVADEFAUT', '', TRUE) = '') then
           EcritureErreur('Code TVA non renseigné');
        if (GetParamSocSecur ('SO_CODETVAGENEDEFAULT', '', TRUE) = '') then
           EcritureErreur('Exigibilité de TVA non renseigné');
        if (GetParamSocSecur ('SO_ETABLISDEFAUT', '', TRUE) = '') then
           EcritureErreur('Etablissement non renseigné');
  end;
end;

  if Result then ChargeMagHalley ;
end;

Function TFAssistImport.CoherenceCompte (indice : integer; Lec : string) : Boolean;
var
Q1       : TQuery;
Req,Ext  : string;
nat, St  : string;
begin
    if RecControl[indice].LeFb = fbGene then
    begin
    Req := 'SELECT * from GENERAUX Where G_GENERAL="'; Ext := 'G_'; nat := 'NATUREGENE';
    end
    else
    begin
    Req := 'SELECT * from TIERS Where T_AUXILIAIRE="'; Ext := 'T_'; nat := 'NATUREAUXI';
    end;
    if RecControl[indice].champ <> 'CPT' then begin Result := TRUE; exit; end;
    if Lec <> '' then
    begin
         if RecControl[indice].valeur <> BourreOuTronque(EnleveBlanc(Lec, 7, 17), RecControl[indice].LeFb) then
         begin Result := TRUE; exit; end;
    end;
    if Lec = '' then
    begin
            Q1 := OpenSQl (Req+ RecControl[indice].valeur+'"', TRUE);
            if not Q1.EOF then
            begin
                 if (RecControl[indice].Nature <> Q1.findField (Ext+nat).asstring) then
                 begin
                      if (RecControl[indice].NomChamp = 'SO_DEFCOLDDIV') and (Q1.findField (Ext+nat).asstring = 'COC') then
                      begin
                          ferme (Q1); Result := TRUE;    exit;
                      end;
                      if (RecControl[indice].NomChamp = 'SO_DEFCOLCDIV') and (Q1.findField (Ext+nat).asstring = 'COF') then
                      begin
                          ferme (Q1); Result := TRUE;    exit;
                      end;
                      if (indice = 0 ) then
                      begin
                          St := RecControl[indice].valeur;
                          St[Length(St)] := Chr(Ord (St[Length(St)])+1);
                          RecControl[indice].valeur := St;
                          AfficheListeCom('Incohérence avec les paramètres du dossier'+
                                               ' Nature du compte : ' + RecControl[indice].valeur ,LISTEIMPORT);
                      end
                      else
                      begin
                           AfficheListeCom('Incohérence avec les paramètres du dossier'+
                           ' Nature du compte : ' + RecControl[indice].valeur ,LISTEIMPORT);
                           RecControl[indice].valeur := '';
                      end;
                 end;
            end;
            ferme (Q1);
     end
     else
     begin
                  if RecControl[indice].Nature <> EnleveBlanc(Lec, 59, 3) then
                  begin
                              if (RecControl[indice].NomChamp = 'SO_DEFCOLDDIV') and (EnleveBlanc(Lec, 59, 3) = 'COC') then
                              begin
                                  Result := TRUE;    exit;
                              end;
                              if (RecControl[indice].NomChamp = 'SO_DEFCOLCDIV') and (EnleveBlanc(Lec, 59, 3) = 'COF') then
                              begin
                                  Result := TRUE;    exit;
                              end;
                              AfficheListeCom(VERIF_PARAMSOC+ RecControl[indice].NomChamp ,LISTEIMPORT);
                              RecControl[indice].valeur := '';
                  end;
     end;
     Result := TRUE;

end;

procedure TFAssistImport.BCalculPieceClick(Sender: TObject);
begin
  inherited;
 Calculpiece := '';
 if (BCalculpiece.Checked) then Calculpiece := 'N'
end;



function TFAssistImport.ImportComSx(FichierSortie: string; InfoImp : PtTInfoImport=nil): Boolean;
var
  Tl, NameFichier       : string;
  FichierImp            : string;
  OkFinImport           : Boolean;
  Mesg                  : string;
  FileName, Filearchive : string;
  PrinterSetupDialog1   : TPrinterSetupDialog;
  RequestTOB            : TOB;
  Hst                   : HTStringList;
{$IFDEF EAGLCLIENT}
  SourceFile            : TFileStream;
  SourceFileString      : string;
  lTobResult            : TOB;
{$ELSE}
  ImpCom                : TImportCom;
  ii                    : integer;
{$ENDIF}

begin
  Tl               := 'TL';
  OkFinImport      := FALSE;

  OKexercice := TRUE;
//  ic := gettickcount;

  if FichierSortie = '' then
  begin
     if stArg <> '' then
      begin
        AfficheListeCom(FICHIER_INEXISTANT + FichierSortie,LISTEIMPORT);
        EcrireDansfichierListeCom (EcritureTraceImport(FALSE), LISTEIMPORT);
      end
      else
        PgiBox(FICHIER_INEXISTANT+' #10' + FichierSortie, 'Import');
      Result := FALSE; Exit;
  end;

//  if (stArg = '') and (Not _BlocageMonoPoste(True)) then begin Result := FALSE; Exit ; end;

  RequestTOB := InitTobParamImp(FichierSortie);
  bFin.Caption := 'Fermer';
  bSuivant.Enabled := FALSE;
  bPrecedent.Enabled := FALSE;

{$IFDEF EAGLCLIENT}  // Passage des données du fichier par la TOB pour le serveur
 if UpperCase(ExtractFileExt(FichierSortie)) <> '.ZIP' then
 begin
        SourceFile := TFileStream.Create(FichierSortie ,fmOpenRead or fmShareDenyNone);
        SetLength(SourceFileString, SourceFile.Size);
        SourceFile.Read(pChar(SourceFileString)^, SourceFile.Size);
        RequestTOB.AddChampSupValeur('SOURCE'   , SourceFileString) ;
        RequestTOB.AddChampSupValeur ('ERROR', '');
        SourceFile.Free;
 end;
{$ENDIF}

  SourisSablier;
  {$IFDEF EAGLCLIENT}
      with cWA.create do
      begin
           OnCallBack:=ComsxOnCallBack ;
          lTobResult := Request('COMSX.IMPORT','',RequestTOB,'','');
         free ;
      end ;
      if lTobResult = nil then
      begin
        PgiBox('Erreur, Server Comsx introuvable','Envoyer') ;
        OKexercice := FALSE;
      end
      else
      OKexercice :=  (lTobResult.GetValue ('ERROR') = '');
  {$ELSE}
      OKexercice := LanceImport (RequestTOB, stArg, AfficheListeComImport, ImpCom);
  {$ENDIF}
      if (not OKexercice) then
      begin
         Result := FALSE;
         if (stArg = '') then
         begin
              Mesg :=  'Attention, aucune donnée n''a été importée.';
              OKexercice := FALSE;
              bFin.Caption := 'Relancer';
              if pos(ERR_CREATE_CPTE, RequestTOB.GetValue ('ERROR')) <> 0 then
                   Mesg := Mesg + #10#13+ 'Vous pouvez créer les établissements, sections, journaux, comptes généraux et auxiliaires manquants en double cliquant sur la ligne';
              PGIInfo(Mesg);
         end;
         if (FEmail.Text <> '') and (pos('@', FEmail.Text) <> 0) then
         begin
             FichierImp := EcritureTraceImport(OkFinImport);
             EcrireDansfichierListeCom (FichierImp, LISTEIMPORT);
             RenvoieCorpsMail (Hst);
             AglSendMail ('Rapport Comsx', FEmail.Text, '', Hst, FichierImp, true, 1);
         end
         else
         begin
               FichierImp := EcritureTraceImport(OKexercice);
               EcrireDansfichierListeCom (FichierImp, LISTEIMPORT);
               if (stArg = '') and (not BNetExpert.Checked) then  // fiche 10593
               begin
                   if OKexercice then CommitTrans
                   else RollBackDiscret ;
                   OkCommit := TRUE;
                   AfficheListeComImport   ('', LISTEIMPORT, 100);
                   if PGIAsk('Import terminé : voulez-vous imprimer le rapport ?','Import')=mrYes then
                   begin
                        PrinterSetupDialog1 := TPrinterSetupDialog.create(Application);
                        if PrinterSetupDialog1.Execute then
                           ControlTextToPrinter(FichierImp,poPortrait);
                        PrinterSetupDialog1.Free;
                   end;
               end;
         end;
{$IFDEF EAGLCLIENT}
         if BRejet.Checked then
         begin
              FileName := FichierSortie;
              //Filearchive := ReadTokenPipe (Filename, '.');
              //Filearchive := Filearchive + '.REJET';
              Namefichier := ExtractFileName(FileName);
              Filearchive := ExtractFileDir(FileName) +  '\' + ReadTokenPipe (Namefichier, '.')+ '.REJET';
              FileName := './ECOMSX/'+ AppServer.SessionID ;
              AppServer.RequestFileTo (Filearchive, FileName+'/'+ExtractFileName(Filearchive));
              with cWA.create do
              begin
                   Request('COMSX.SUPPDIR','', RequestTOB,'','');
                   free;
              end;
         end;
{$ENDIF}
         RequestTOB.Free;  SourisNormale;
         exit;
  end
  else
  begin
{$IFDEF EAGLCLIENT}
         if BRejet.Checked then
         begin
              FileName := FichierSortie;
              //Filearchive := ReadTokenPipe (Filename, '.');
              //Filearchive := Filearchive + '.REJET';
              Namefichier := ExtractFileName(FileName);
              Filearchive := ExtractFileDir(FileName) +  '\' + ReadTokenPipe (Namefichier, '.')+ '.REJET';
              FileName := './ECOMSX/'+ AppServer.SessionID ;
              AppServer.RequestFileTo (Filearchive, FileName+'/'+ExtractFileName(Filearchive));
              with cWA.create do
              begin
                   Request('COMSX.SUPPDIR','', RequestTOB,'','');
                   free;
              end;
         end;
{$ENDIF}
         OkFinImport := TRUE;
  end;
  RequestTOB.free;
  SourisNormale;
  if (ctxPCL in V_PGI.PGIContexte) and  OkFinImport then // Fiche 18613
  begin
// ajout me car setflagappli plante en Eaglclient
          SetFlagAppliDossier ('CCS5.EXE', V_PGI.NoDossier, TRUE);
          SetFlagAppliDossier ('COMSX.EXE', V_PGI.NoDossier, TRUE);
          //initialisation des tables libres
          InitParamTablesLibres ('');
  end;

  if (OrigineEnvoie = 'EXP') and  (NatureFichier = 'SYN') and (DateArrete <> iDate1900) then
       CloPerSynchro (DateToStr(DateArrete));

  if stArg <> '' then
  begin
         if (not BNetExpert.Checked) then begin FTimer.free; FTimer := nil; end;
         ModalResult := 1;
  end;

  AfficheListeCom('Fin de Traitement ',LISTEIMPORT);
//  if (stArg = '') then _DeblocageMonoPoste(True) ;
  bFin.Caption := 'Fermer';
  bSuivant.Enabled := FALSE;
  bPrecedent.Enabled := FALSE;
  if not BNetExpert.Checked then
  begin
       FileName := FICHENAME.Text;
       Namefichier := ExtractFileName(FileName);
       Filearchive := ExtractFileDir(FileName) +  '\' + ReadTokenPipe (Namefichier, '.');
      // fiche 10450 Filearchive := ReadTokenPipe (Filename, '.');

       if  OrigineFichier = 'IMMO' then
       begin
           FileName  := Filearchive + '.ARM';
           if (FileExists(FileName)) then DeleteFile(FileName);
           RenameFile(Filearchive +'.BOB', FileName);
       end
       else
       begin
{$IFNDEF EAGLCLIENT} // correction par la ligne de commande plantage dans transport fichier
           if (TobFich <> nil) and (TobFich.Detail.Count > 0) then
           begin
               for ii :=0 to TobFich.detail.count-1 do
                 sysutils.DeleteFile(ExtractFileDir ( FICHENAME.Text ) + '\' + ExtractFileName(TobFich.detail[ii].getvalue('ZF_NAME')));
               TobFich.free;
           end
           else
{$ENDIF}
           begin
                if  OrigineFichier <> 'PAIE' then
                begin
                     FileName  := Filearchive + '.ARC';
                     if (FileExists(FileName)) and (FileName <> FICHENAME.Text)then DeleteFile(FileName);
                     if UpperCase(ExtractFileExt(FICHENAME.Text)) <> '.ZIP' then
                        RenameFile(FICHENAME.Text, FileName);
                end;
           end;
       end;
  end;

//  jc:= gettickcount;
 // showmessage(inttostr(jc-ic));

  FichierImp := EcritureTraceImport(OkFinImport);
  EcrireDansfichierListeCom (FichierImp, LISTEIMPORT);

  if (stArg = '') and (not BNetExpert.Checked) then
  begin
     if OKexercice then CommitTrans
     else RollBackDiscret ;
     OkCommit := TRUE;
     AfficheListeComImport   ('', LISTEIMPORT, 100);
     if PGIAsk('Import terminé : voulez-vous imprimer le rapport ?','Import')=mrYes then
     begin
          PrinterSetupDialog1 := TPrinterSetupDialog.create(Application);
          if PrinterSetupDialog1.Execute then
             ControlTextToPrinter(FichierImp,poPortrait);
          PrinterSetupDialog1.Free;
     end;                  
  end;

  if (not BEnvoiMail.checked) and (TRIM(FEmail.Text) <> '') and (pos('@', FEmail.Text) <> 0) then
  begin
         RenvoieCorpsMail (Hst);
         AglSendMail('Rapport Comsx', FEmail.Text, '', Hst, FichierImp, true, 1);
  end;

  if Hst <> nil then Hst.Free;

  Result := TRUE;
end;

Function TFAssistImport.AfficheListeComImport(Chaine: string; Listecom : TListBox; pourcentage : double=0) : Boolean;
var
Ch : string;
begin
  Result := TRUE;
  if ListeCom = nil then
       ListeCom := LISTEIMPORT;
  if Chaine <> '' then
  begin
        Ch := TraduireMemoire(Chaine);
        Listecom.Items.add (Ch);
        Listecom.ItemIndex := Listecom.Items.Count-1;
        IF Pos(' est manquant', Ch)  <> 0 then
           Listecom.Selected[Listecom.ItemIndex]:=True ;
  end;
  if (WindowState  = wsMinimized) then Result := AfficheProgressbar(Ch)
  else
  begin
        if pourcentage > 0 then
        begin
          if pourcentage = 100 then
            Label15.Caption := 'Traitement terminé  ... '
          else
          if pourcentage > 99 then
            Label15.Caption := 'Traitement en cours ... : 100 %'
          else
            Label15.Caption := 'Traitement en cours ... : ' + FloatTostr(Arrondi(pourcentage, 2)) + ' %';
        end;
  end;
end;

function TFAssistImport.InitTobParamImp (fichier : string) : TOB ;
var
LT, L1   : TOB;
{$IFDEF EAGLCLIENT}
FileName,Erreur      : string;
{$ENDIF}
{$IFDEF CERTIFNF}
id                   : integer;
Q1                   : TQuery;
{$ENDIF}
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
    LT.AddChampSupValeur('DOSSIER'   , V_PGI.NoDossier ) ;
    LT.AddChampSupValeur('BaseCommune', EstBaseCommune);
    LT.AddChampSupValeur('ERROR', '');
    LT.AddChampSupValeur('APPLICATION', NameAppli);

    L1 := TOB.Create('Trans', LT, -1) ;
    L1.AddChampSupValeur('FichierSortie'   , Fichier) ;
    L1.AddChampSupValeur('AutoriseEcheck'   , AutoriseEcheck) ;
//GP     if (not AutoriseEcheck) and ((INTERDIRECREAT.Text = '<<Tous>>') or (INTERDIRECREAT.Text = '')) then
    if (not AutoriseEcheck) and ((INTERDIRECREAT.Text = Traduirememoire('<<Tous>>')) or (INTERDIRECREAT.Text = '')) then
       L1.AddChampSupValeur ('INTERDIRECREAT', 'IGE;IAU;IJA;IET;ISE')
    else
    begin
       if (BControlEtab.Checked) and (pos('IET', INTERDIRECREAT.Text) = 0) then
          INTERDIRECREAT.Text := INTERDIRECREAT.Text+ 'IET';
       L1.AddChampSupValeur ('INTERDIRECREAT', INTERDIRECREAT.Text);
    end;
    L1.AddChampSupValeur('RepriseComplement',RepriseComplement);
    L1.AddChampSupValeur ('Paramgene', Paramgene);
    L1.AddChampSupValeur ('CPSTAT', CP_STAT.value);
    L1.AddChampSupValeur ('RupturePiece', (BRupturePiece.Checked));
    L1.AddChampSupValeur ('CalculPiece', Calculpiece);
    L1.AddChampSupValeur ('BMAJCPTE', (BMAJCPTE.Checked));
    L1.AddChampSupValeur ('TypeEcrR', TypeEcrR);
    L1.AddChampSupValeur ('TypeEcrI', TypeEcrI);
    {$IFDEF CERTIFNF}
        id := 0;
        if ExisteSQL('SELECT CPJ_CODE Code from CPJALIMPORT ' +
                    ' Where CPJ_DATE = "'+ USDATETime(DATE)+'"') then
        begin
                  Q1 := OpenSQL('SELECT MAX(CPJ_CODE) from CPJALIMPORT '+
                          ' Where CPJ_DATE = "'+ USDATETime(DATE)+'"', True);
                  id := StrToint((Copy ( Q1.Fields[0].asstring, 2, 2)));
                  inc(id);
                  Ferme (Q1);
        end;
        if OptionEchange then
                CpjCode := 'X' + FormatFloat('00',id)
        else
                CpjCode := 'I' + FormatFloat('00',id);

       L1.AddChampSupValeur ('QualifOrigine', CpjCode);
   {$ELSE}
       L1.AddChampSupValeur ('QualifOrigine', QualifOrigine);
   {$ENDIF CERTIFNF}
    L1.AddChampSupValeur ('Sansaux', Sansaux);
    L1.AddChampSupValeur ('ModePCL', V_PGI.ModePCL);
    L1.AddChampSupValeur ('FCorresp', FCorresp.text);
    L1.AddChampSupValeur ('OrigineEnvoie', OrigineEnvoie);
    L1.AddChampSupValeur ('OrigineFichier', OrigineFichier);
    L1.AddChampSupValeur ('DecoupLon', DecoupLon.value);
    L1.AddChampSupValeur ('BIntegre', (BIntegre.Checked));
    L1.AddChampSupValeur ('BRejet', (BRejet.Checked));
    L1.AddChampSupValeur ('BDESEQUILIBRE', (BDESEQUILIBRE.Checked));
    L1.AddChampSupValeur ('Gestiondoublon', (BGestiondoublon.Checked));
    L1.AddChampSupValeur ('BControlEtab', (BControlEtab.Checked));
    if (NatureFichier <> 'BAL') then
    begin
         INTERDIRECREATTIERS.Enabled := FALSE;
         INTERDIRECREATTIERS.Checked := FALSE;
    end;

    L1.AddChampSupValeur ('CREATIONTIERS', (INTERDIRECREATTIERS.Checked));

    // Mode avancée compte de remplacement
    L1.AddChampSupValeur ('CPTEGENE', (CPTEGENE.Text));
    L1.AddChampSupValeur ('CPTECLIENT', (CPTECLIENT.Text));
    L1.AddChampSupValeur ('CPTEFOUR', (CPTEFOUR.Text));
    L1.AddChampSupValeur ('CPTESALAIRE', (CPTESALAIRE.Text));
    L1.AddChampSupValeur ('CPTEDIVERS', (CPTEDIVERS.Text));
    L1.AddChampSupValeur ('RAXE1', (RAXE1.Text));
    L1.AddChampSupValeur ('RAXE2', (RAXE2.Text));
    L1.AddChampSupValeur ('RAXE3', (RAXE3.Text));
    L1.AddChampSupValeur ('RAXE4', (RAXE4.Text));
    L1.AddChampSupValeur ('RAXE5', (RAXE5.Text));
    L1.AddChampSupValeur ('SUBCOLLCLI', (SUBCOLLCLI.Text));
    L1.AddChampSupValeur ('SUBCOLLFOU', (SUBCOLLFOU.Text));
    L1.AddChampSupValeur ('MRDEFAUT', (MRDEFAUT.value));
    L1.AddChampSupValeur ('REGDEFAUT', (REGDEFAUT.value));
    L1.AddChampSupValeur ('BValider', (BValider.Checked));
    L1.AddChampSupValeur ('BCtrP', (BCtrP.Checked));
    L1.AddChampSupValeur ('BCtrS', (BCTRLSOCIETE.Checked));
    L1.AddChampSupValeur ('EcartChange', (BEcartChange.Checked));
    L1.AddChampSupValeur ('OptionEchange', (OptionEchange));
    L1.AddChampSupValeur ('BTPayeur', (BTPayeur.Checked));
    if DATEECR1.Text = '  /  /    ' then
       L1.AddChampSupValeur ('DATEECR', stDate1900)
    else
       L1.AddChampSupValeur ('DATEECR', (DATEECR1.Text));

    L1.AddChampSupValeur ('BRecalculS', (BRecalculSolde.Checked));

    L1.AddChampSupValeur ('BSoldeLigne', (BCalcSoldeLigne.Checked));

    L1.AddChampSupValeur ('TYPEVISA', (CBBAP.value));

    L1.AddChampSupValeur ('BBLANC', (BBlanc.Checked));
    L1.AddChampSupValeur ('MODIFTIERS', MODIFTIERS.Text);
    L1.AddChampSupValeur ('MajDateDernMvt', BMajDateDernMvt.Checked);
    L1.AddChampSupValeur ('CalculdateEche', CalculdateEche.Checked);
    L1.AddChampSupValeur ('ControlExercClos', ControlExercClos);
    L1.AddChampSupValeur ('ScanBanck', ScanBanck);


{$IFDEF EAGLCLIENT}
    if UpperCase(ExtractFileExt(Fichier)) = '.ZIP' then
    begin
              // déplacement du zip sur le serveur
              with cWA.create do
              begin
                   OnCallBack:=ComsxOnCallBack ;
                   Request('COMSX.CREATEDIR','', LT,'','');
                   FileName := './'+NameAppli+'/'+ AppServer.SessionID ;
                   AppServer.upload (FileName+'/'+ExtractFileName(Fichier), Fichier, Erreur);
                   free;
              end;
    end;
{$ENDIF}
    Result := LT;
    // affichage des options d'import
    AffichageOptionImport;

end;


procedure TFAssistImport.BInterdireCreatCompteClick(Sender: TObject);
begin
  inherited;
AutoriseEcheck  := not(BInterdireCreatCompte.checked);
if BInterdireCreatCompte.checked then
   INTERDIRECREAT.Enabled := TRUE
else
   INTERDIRECREAT.Enabled := FALSE;

end;

procedure TFAssistImport.Generaux1Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IGE') ;
end;

procedure TFAssistImport.Auxiliaires1Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IAU') ;
end;

procedure TFAssistImport.Axeanalytique11Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IA1') ;
end;

procedure TFAssistImport.Axeanalytique21Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IA2') ;
end;

procedure TFAssistImport.Axeanalytique31Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IA3') ;
end;

procedure TFAssistImport.Axeanalytique41Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IA4') ;
end;

procedure TFAssistImport.Axeanalytique51Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IA5') ;
end;

procedure TFAssistImport.ModedePaiement1Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IPM') ;
end;

procedure TFAssistImport.Journaux1Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IJA') ;
end;

procedure TFAssistImport.ParamtrageSISCOII1Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('SIS') ;
end;

procedure TFAssistImport.TQUALPIECEChange(Sender: TObject);
begin
  inherited;
TypeEcrI := TQUALPIECE.value;
end;

procedure TFAssistImport.MenuItem1Click(Sender: TObject);
begin
  inherited;
          //         Cache.ZoomTable:=tzGeneral ;
          //         GChercheCompte(Cache, Nil);
          LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL', True,1)  ;
end;

procedure TFAssistImport.ZoomauxClick(Sender: TObject);
begin
  inherited;
         //          Cache.ZoomTable:=tzTiers ;
         //          GChercheCompte(Cache, Nil);
          LookUpList(TControl(Sender), TraduireMemoire('Auxiliaire'), 'TIERS',
          'T_AUXILIAIRE', 'T_LIBELLE', '', 'T_AUXILIAIRE', True, 2);

end;

procedure TFAssistImport.ZoomanaClick(Sender: TObject);
begin
  inherited;

                   LookupList( TControl(Sender),          // Contrôle lookup à renseigné
                            'Recherche d''une section analytique',         // Titre
                            'SECTION',         // Table
                            'S_SECTION',       // Colonne de retour
                            'S_LIBELLE',       // Select
                            '',                // Where
                            'S_SECTION',       // Orber by
                            true,              // Affichage des colonnes de titres
                            CRenseigneTagLookup(fBAxe1),           // N° de tag de la fonction de creation / modification
                            '',                // Ordre SQL Complet
                            tlDefault,        // Par défaut
                            0        // Agrandissement de la largeur pour la recherche sur libelle
                        ) ;

end;


procedure TFAssistImport.LISTEIMPORTContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
begin
  inherited;

  PopZ1.Items[0].Visible := (pos ('Le compte général ', LISTEIMPORT.Items[LISTEIMPORT.ItemIndex]) <> 0);
  PopZ1.Items[1].Visible := (pos ('Le compte auxiliaire ', LISTEIMPORT.Items[LISTEIMPORT.ItemIndex]) <> 0);
  PopZ1.Items[2].Visible := (pos ('La section ', LISTEIMPORT.Items[LISTEIMPORT.ItemIndex]) <> 0);
  PopZ1.Items[3].Visible := (pos ('Le journal ', LISTEIMPORT.Items[LISTEIMPORT.ItemIndex]) <> 0);
  PopZ1.Items[4].Visible := (pos ('Etablissement ', LISTEIMPORT.Items[LISTEIMPORT.ItemIndex]) <> 0);

end;


procedure TFAssistImport.LISTEIMPORTDblClick(Sender: TObject);
var
ValeurManquant,Chaine,Axe      : string;
PosM,PosA                      : integer;
     function RechCtrl (RechChaine : string; ps : integer) : string;
     var
     ii : integer;
     begin
          ii := Pos (RechChaine, Chaine);
          if (ii <> 0) then
          begin
               ValeurManquant := Copy (Chaine, ii+ps, PosM-ps);
               Result := ValeurManquant;
          end;
     end;
begin
  inherited;
  Chaine := LISTEIMPORT.Items[LISTEIMPORT.ItemIndex];
  PosM := Pos(' est manquant', Chaine);
  if (PosM <> 0) then
  begin
      if RechCtrl ('Le compte général ', 18) <> '' then
          FicheGene(nil, '', ValeurManquant, tacreat, 0)
      else
      if RechCtrl ('Le compte auxiliaire ', 21) <> '' then
                FicheTiersZoom(nil, '', ValeurManquant, taCreat, 1, '')
      else
      if RechCtrl ('La section ', 11) <> '' then
      begin
                PosA := pos ('Axe :', Chaine);
                if PosA <> 0 then Axe := Copy (Chaine, PosA+5, 2);
                   FicheSectionMZS(Axe, ValeurManquant, taCreatOne, 0, '');
      end
      else
      if RechCtrl ('Etablissement ', 14) <> '' then
                  FicheEtablissement_AGL(taCreat)
      else
      if RechCtrl ('Le journal ', 11) <> '' then
                  FicheJournal(Nil,'', Copy (Chaine, 12, 3) , taCreat,0) ;

  end; // PosM <> 0
end;

procedure TFAssistImport.Zoomsurlejournal1Click(Sender: TObject);
begin
  inherited;
                   LookupList( TControl(Sender),          // Contrôle lookup à renseigné
                            'Recherche d''un journal',         // Titre
                            'JOURNAL',         // Table
                            'J_JOURNAL',       // Colonne de retour
                            'J_LIBELLE',       // Select
                            '',                // Where
                            'J_JOURNAL',       // Orber by
                            true,              // Affichage des colonnes de titres
                            CRenseigneTagLookup(fbJal),           // N° de tag de la fonction de creation / modification
                            '',                // Ordre SQL Complet
                            tlDefault,        // Par défaut
                            0        // Agrandissement de la largeur pour la recherche sur libelle
                        ) ;

end;

constructor TFAssistImport.create(stArgument, FileIni, Commande: string; OptE : Boolean);
begin
stArg := stArgument;
FichierIni := FileIni;
AccesCommande := Commande;
OptionEchange := OptE;
inherited Create(Application);

end;

procedure TFAssistImport.bPrecedentClick(Sender: TObject);
begin
  inherited;
      if (GetPage = AVANCE) and (not OPTIONAVANCE.Checked) or
      (GetPage = AVANCE2) and (not OPTIONAVANCE.Checked) then bPrecedentClick(Sender);
      if (Getpage = Resume) or (Getpage = PIMPORT) then
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

procedure TFAssistImport.Zoomsurtablissement1Click(Sender: TObject);
begin
  inherited;
  FicheEtablissement_AGL(taCreat);
end;

procedure TFAssistImport.OPTIONAVANCEClick(Sender: TObject);
begin
  inherited;
       if stArg = '' then  // sans ligne de commande
       begin
            if OPTIONAVANCE.Checked then
            begin
                 TQUALPIECE.value := 'N';
                 TypeEcrI := TQUALPIECE.values[TQUALPIECE.ItemIndex];
                 TypeEcrR := TypeEcrI;
                 SuppComptable.Checked := FALSE;
            end
            else
            begin
                 TypeEcrI := '';
                 if (VH^.CPLienGamme = 'S1') then
                    SuppComptable.Checked := FALSE
                 else
                    SuppComptable.Checked := TRUE;
            end;
       end;

end;

procedure TFAssistImport.Etablissement1Click(Sender: TObject);
begin
  inherited;
   CCLanceFiche_Correspondance ('IET') ;
end;

procedure TFAssistImport.InitComImport( Stt : string);
var
FicIni                   : TIniFile;
St,NomFichier            : string;
               Procedure IniEnabled (Champ : string; CTR : TWinControl);
               begin
                  St := FicIni.ReadString (AccesCommande, Champ, '');
                  TCheckBox(CTR).Checked := (St = 'TRUE') or (St = 'X');
               end;
begin
                FicIni           := TIniFile.Create(FichierIni);
                RepEncours       := FicIni.ReadString (AccesCommande, 'REPERTOIRE', '');
                if (RepEncours <> '') and (RepEncours[length (RepEncours)] <> '\') then RepEncours := RepEncours + '\';
                NomFichier := FicIni.ReadString (AccesCommande, 'NOMFICHIER', '');
                if (pos ('*', NomFichier) <> 0) then
                   FICHENAME.Text :=  RepEncours+Stt
                else
                   FICHENAME.Text   := RepEncours+NomFichier;
                FichImp          := FicIni.ReadString (AccesCommande, 'RAPPORT', '');
                TypeEcrR         :=  FicIni.ReadString (AccesCommande, 'TYPEECR', '');    // type d'écriture reçu  exemple simule
                TypeEcrI         := FicIni.ReadString (AccesCommande, 'ECRINTERGE', ''); // type d'écriture à intégrer exemple Normal
                if (TypeEcrR <> 'S') and (TypeEcrR <> 'N') then TypeEcrR := '';
                if (TypeEcrI <> 'S') and (TypeEcrI <> 'N') then TypeEcrI := '';

                IniEnabled ('DOUBLON', BGestiondoublon);
                IniEnabled ('CPARAM', BControlParam);

                St               := FicIni.ReadString (AccesCommande, 'DECOUPEFOLIO', '');
                if (St <> '') and Isnumeric(St) then
                   DecoupLon.value  := StrToint(St);

                IniEnabled ('INTEGRATION', BIntegre);
                IniEnabled ('RAPPORTREJET', BRejet);
                IniEnabled ('DESEQUILIBRE', BDESEQUILIBRE);


                FEMail.text      := FicIni.ReadString (AccesCommande, 'MAIL', '');
                // Mode avancée

                IniEnabled ('MAJCPTE', BMAJCPTE);
                IniEnabled ('INTERDIRECREATIERS', INTERDIRECREATTIERS);
                IniEnabled ('CETAB', BControlEtab);
                IniEnabled ('RUPTUREPIECE', BRupturePiece);
                IniEnabled ('CALCULPIECE', BCalculPiece);
                IniEnabled ('SANSECHEC', BInterdireCreatCompte);
                AutoriseEcheck  := not(BInterdireCreatCompte.checked);

                INTERDIRECREAT.Text := FicIni.ReadString (AccesCommande, 'INTERDIRECREAT', '');
                if (INTERDIRECREAT.Text <> '') and  (INTERDIRECREAT.Text <> '[]') then AutoriseEcheck  := FALSE;

                IniEnabled ('VALIDERECR', BValider);
                IniEnabled ('CALCONTREPARTIE', BCtrP);
                FCorresp.Text    := FicIni.ReadString (AccesCommande, 'CORRESP', '');
                CPTEGENE.Text    := FicIni.ReadString (AccesCommande, 'CPTEGENE', '');
                CPTEFOUR.Text    := FicIni.ReadString (AccesCommande, 'CPTEFOUR', '');
                CPTEDIVERS.Text  := FicIni.ReadString (AccesCommande, 'CPTEDIVERS', '');
                CPTECLIENT.Text  := FicIni.ReadString (AccesCommande, 'CPTECLIENT', '');
                CPTESALAIRE.Text := FicIni.ReadString (AccesCommande, 'CPTESALAIRE', '');
                RAXE1.Text       := FicIni.ReadString (AccesCommande, 'RAXE1', '');
                RAXE2.Text       := FicIni.ReadString (AccesCommande, 'RAXE2', '');
                RAXE3.Text       := FicIni.ReadString (AccesCommande, 'RAXE3', '');
                RAXE4.Text       := FicIni.ReadString (AccesCommande, 'RAXE4', '');
                RAXE5.Text       := FicIni.ReadString (AccesCommande, 'RAXE5', '');
                SUBCOLLCLI.Text  := FicIni.ReadString (AccesCommande, 'SUBCOLLCLI', '');
                SUBCOLLFOU.Text  := FicIni.ReadString (AccesCommande, 'SUBCOLLFOU', '');
                MRDEFAUT.value   := FicIni.ReadString (AccesCommande, 'MRDEFAUT', '');
                REGDEFAUT.value  := FicIni.ReadString (AccesCommande, 'REGDEFAUT', '');
                CBBAP.value      := FicIni.ReadString (AccesCommande, 'TYPEVISA', '');
                DATEECR1.Text    := FicIni.ReadString (AccesCommande, 'DATEECRITURE', '');
                IniEnabled ('BLANC', BBlanc);
                MODIFTIERS.Text  :=  FicIni.ReadString (AccesCommande, 'MODIFTIERS', '');

                if DATEECR1.Text = '' then   DATEECR1.Text := stDate1900;

                IniEnabled ('CTRS', BCTRLSOCIETE);
                IniEnabled('VALIDER', BValider);
                IniEnabled ('CTRP', BCtrP);
                IniEnabled ('ECARTCHANGE', BEcartChange);
                IniEnabled ('ANOMAIL', BEnvoiMail);
                IniEnabled ('TIERSPAYEUR',  BTPayeur);
                IniEnabled ('SOLDELIGNE',  BCalcSoldeLigne);
                // ajout me pour ne pas faire le recalcul des soldes
                IniEnabled ('RECALCULSOLDE', BRecalculSolde);
                IniEnabled ('MAJDATEDERNMVT',  BMajDateDernMvt);
                IniEnabled ('CALCULDATEECHE',  CalculdateEche);
                // fiche 10558
                  St := FicIni.ReadString (AccesCommande, 'CTREXERCLOS', '');
                ControlExercClos := (St = 'TRUE') or (St = 'X');
                FicIni.free;
end;


procedure TFAssistImport.AffichageOptionImport;
begin
    if (NatureFichier <> 'JRL') and (NatureFichier <> 'BAL') then exit;

    AfficheListeComImport('Options Avancées société : '+ V_PGI.NomSociete, LISTEIMPORT);
    AfficheListeComImport   ('   Nom du fichier : '+ FICHENAME.Text, LISTEIMPORT);

    if (BInterdireCreatCompte.checked) then
       AfficheListeComImport('   Interdire la création en l''absence du paramétrage : '+ INTERDIRECREAT.Text, LISTEIMPORT);
    if (BControlEtab.Checked) then
       AfficheListeComImport('   L''intégrité des établissements ', LISTEIMPORT);
    if (BRupturePiece.Checked) then
       AfficheListeComImport('   Conserver les ruptures des pièces ', LISTEIMPORT);
    if (Calculpiece = 'N')  then
      AfficheListeComImport('    Conserver les numéros de pièce ', LISTEIMPORT);
    if (BMAJCPTE.Checked) then
      AfficheListeComImport('    Mise à jour des comptes ', LISTEIMPORT);
    if (TypeEcrI <> '') then
      AfficheListeComImport('    Importer les écritures en type ' + TypeEcrI, LISTEIMPORT);
    if (FCorresp.text <> '') then
      AfficheListeComImport('    Tables de correspondance ' + FCorresp.text, LISTEIMPORT);
    if (DecoupLon.value <> 0) then
      AfficheListeComImport('    Découpage journal libre ou bordereau tout les ' + IntToStr(DecoupLon.value), LISTEIMPORT);
    if (BIntegre.Checked) then
      AfficheListeComImport('    Intégration partielle (uniquement les écritures cohérentes)', LISTEIMPORT);
    if (BRejet.Checked) then
      AfficheListeComImport('    Générer un fichier de rejet', LISTEIMPORT);
    if (BDESEQUILIBRE.Checked) then
      AfficheListeComImport('    Rejet en cas de déséquilibre des montants', LISTEIMPORT);

    if (BGestiondoublon.Checked) then
      AfficheListeComImport('    Interdire les doublons', LISTEIMPORT);
    if (INTERDIRECREATTIERS.Checked) then
      AfficheListeComImport('    Interdire la création des tiers', LISTEIMPORT);

    // Mode avancée compte de remplacement
    if (CPTEGENE.Text) <> '' then
      AfficheListeComImport('    Compte général de remplaçement :' + CPTEGENE.Text, LISTEIMPORT);
    if (CPTECLIENT.Text) <> '' then
      AfficheListeComImport('    Compte client de remplaçement :' + CPTECLIENT.Text, LISTEIMPORT);
    if (CPTEFOUR.Text) <> '' then
      AfficheListeComImport('    Compte fournisseur de remplaçement :' + CPTEFOUR.Text, LISTEIMPORT);
    if (CPTESALAIRE.Text) <> '' then
      AfficheListeComImport('    Compte salarié de remplaçement :' + CPTESALAIRE.Text, LISTEIMPORT);
    if (CPTEDIVERS.Text) <> '' then
      AfficheListeComImport('    Compte salarié de remplaçement :' + CPTEDIVERS.Text, LISTEIMPORT);

    if (RAXE1.Text) <> '' then
      AfficheListeComImport('    Section de remplaçement sur axe1 :' + RAXE1.Text, LISTEIMPORT);
    if (RAXE2.Text) <> '' then
      AfficheListeComImport('    Section de remplaçement sur axe2 :' + RAXE2.Text, LISTEIMPORT);
    if (RAXE3.Text) <> '' then
      AfficheListeComImport('    Section de remplaçement sur axe3 :' + RAXE3.Text, LISTEIMPORT);
    if (RAXE4.Text) <> '' then
      AfficheListeComImport('    Section de remplaçement sur axe4 :' + RAXE4.Text, LISTEIMPORT);
    if (RAXE5.Text) <> '' then
      AfficheListeComImport('    Section de remplaçement sur axe5 :' + RAXE5.Text, LISTEIMPORT);

    if (SUBCOLLCLI.Text) <> '' then
      AfficheListeComImport('    Collectif client de substitution :' + SUBCOLLCLI.Text, LISTEIMPORT);
    if (SUBCOLLFOU.Text) <> '' then
      AfficheListeComImport('    Collectif fournisseur de substitution :' + SUBCOLLFOU.Text, LISTEIMPORT);

    if (MRDEFAUT.value) <> '' then
      AfficheListeComImport('    Mode de règlement de remplaçement :' + MRDEFAUT.value, LISTEIMPORT);
    if (REGDEFAUT.value) <> '' then
      AfficheListeComImport('    Régime de TVA de remplaçement :' + REGDEFAUT.value, LISTEIMPORT);
    if (BValider.Checked) then
      AfficheListeComImport('    Valider les écritures à l''intégration ', LISTEIMPORT);
    if (BCtrP.Checked) then
      AfficheListeComImport('    Recalcule des contreparties' , LISTEIMPORT);
    if (BEcartChange.Checked) then
      AfficheListeComImport('    Génération d''écart de change' , LISTEIMPORT);
    if BTPayeur.Checked then
      AfficheListeComImport ('    Gestion tiers payeurs' , LISTEIMPORT);
    AfficheListeComImport('   ' , LISTEIMPORT);
end;


procedure TFAssistImport.BINIClick(Sender: TObject);
var
FicIni                   : TIniFile;
Rep                      : string;
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
         FileName := 'COMSX.INI';
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
      FicIni.WriteString ('COMMANDE', 'REPERTOIRE', Rep);
      FicIni.WriteString ('COMMANDE', 'NOMFICHIER', ExtractFileName (FICHENAME.Text));
      if FEMail.text <> '' then
      FicIni.WriteString ('COMMANDE', 'MAIL', FEMail.text );
      if TypeEcrI <> '' then
      begin
           FicIni.WriteString ('COMMANDE', 'TYPEECR', TypeEcrI );
           FicIni.WriteString ('COMMANDE', 'ECRINTERGE', TypeEcrI );
      end;
      if (BInterdireCreatCompte.checked) then
      begin
           FicIni.WriteString ('COMMANDE', 'SANSECHEC', 'TRUE' );
      end;
      if BGestiondoublon.Checked then
         FicIni.WriteString ('COMMANDE', 'DOUBLON', 'TRUE');
      if BIntegre.Checked then
         FicIni.WriteString ('COMMANDE', 'INTEGRATION', 'TRUE');
      if BControlParam.Checked then
         FicIni.WriteString ('COMMANDE', 'CPARAM', 'TRUE' );
      if DecoupLon.Value > 0 then
         FicIni.WriteString ('COMMANDE', 'DECOUPEFOLIO', IntTostr(DecoupLon.Value));
      if FichImp <> '' then
         FicIni.WriteString ('COMMANDE', 'RAPPORT', FichImp );
      if BMAJCPTE.Checked then
         FicIni.WriteString ('COMMANDE', 'MAJCPTE', 'TRUE' );
      if (INTERDIRECREATTIERS.Checked) then
         FicIni.WriteString ('COMMANDE', 'INTERDIRECREATIERS', 'TRUE' );
      if BControlEtab.Checked then
         FicIni.WriteString ('COMMANDE', 'CETAB', 'TRUE' );
      if BRupturePiece.Checked then
         FicIni.WriteString ('COMMANDE', 'RUPTUREPIECE', 'TRUE' );
      if BCalculPiece.Checked then
         FicIni.WriteString ('COMMANDE', 'CALCULPIECE', 'TRUE' );
      if BRejet.Checked then
         FicIni.WriteString ('COMMANDE', 'RAPPORTREJET', 'TRUE' );
      if BDESEQUILIBRE.Checked then
         FicIni.WriteString ('COMMANDE', 'DESEQUILIBRE', 'TRUE' );
      if BValider.Checked then
         FicIni.WriteString ('COMMANDE', 'VALIDERECR', 'TRUE' );
      if BCtrP.Checked then
         FicIni.WriteString ('COMMANDE', 'CALCONTREPARTIE', 'TRUE' );
      if FCorresp.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'CORRESP', FCorresp.Text );
      if CPTEGENE.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'CPTEGENE', CPTEGENE.Text );
      if CPTEFOUR.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'CPTEFOUR', CPTEFOUR.Text );
      if CPTEDIVERS.text <> '' then
         FicIni.WriteString ('COMMANDE', 'CPTEDIVERS', CPTEDIVERS.text );
      if CPTECLIENT.text <> '' then
         FicIni.WriteString ('COMMANDE', 'CPTECLIENT', CPTECLIENT.text );
      if CPTESALAIRE.text <>'' then
         FicIni.WriteString ('COMMANDE', 'CPTESALARIE', 'TRUE' );
      if RAXE1.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'RAXE1', RAXE1.Text );
      if RAXE2.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'RAXE2', RAXE2.Text );
      if RAXE3.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'RAXE3', RAXE3.Text );
      if RAXE4.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'RAXE4', RAXE4.Text );
      if RAXE5.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'RAXE5', RAXE5.Text );
      if SUBCOLLCLI.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'SUBCOLLCLI', SUBCOLLCLI.Text );
      if SUBCOLLFOU.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'SUBCOLLFOU', SUBCOLLFOU.Text );
      if MRDEFAUT.value <> '' then
         FicIni.WriteString ('COMMANDE', 'MRDEFAUT', MRDEFAUT.value);
      if REGDEFAUT.value <> '' then
         FicIni.WriteString ('COMMANDE', 'REGDEFAUT', 'TRUE' );
      if BEnvoiMail.checked  then
         FicIni.WriteString ('COMMANDE', 'ANOMAIL', 'TRUE');
      if BEcartChange.Checked then
         FicIni.WriteString ('COMMANDE', 'ECARTCHANGE', 'TRUE');
      if (BCTRLSOCIETE.Checked) then
         FicIni.WriteString ('COMMANDE', 'CTRS', 'TRUE');
      if BValider.Checked then
         FicIni.WriteString ('COMMANDE', 'VALIDER', 'TRUE');
      if BCtrP.Checked then
         FicIni.WriteString ('COMMANDE', 'CTRP', 'TRUE');
      if (BTPayeur.Checked) then
         FicIni.WriteString ('COMMANDE', 'TIERSPAYEUR', 'TRUE');
      if (CBBAP.value) <> '' then
         FicIni.WriteString ('COMMANDE', 'TYPEVISA', CBBAP.value);
      if (DATEECR1.Text <> '') and (DATEECR1.Text <> stDate1900) then
         FicIni.WriteString ('COMMANDE', 'DATEECRITURE', DATEECR1.Text);

      if BCalcSoldeLigne.Checked then
         FicIni.WriteString ('COMMANDE', 'SOLDELIGNE',  'TRUE');
      if BBlanc.Checked then
         FicIni.WriteString ('COMMANDE', 'BLANC',  'TRUE');
      if MODIFTIERS.Text <> '' then
         FicIni.WriteString ('COMMANDE', 'MODIFTIERS', MODIFTIERS.Text );

      if not BRecalculSolde.Checked  then  // ajout me pour ne pas faire le recalcul des soldes
         FicIni.WriteString ('COMMANDE', 'RECALCULSOLDE', 'FALSE');

      if BMajDateDernMvt.Checked then
         FicIni.WriteString ('COMMANDE', 'MAJDATEDERNMVT', 'TRUE');

      if CalculdateEche.Checked then
         FicIni.WriteString ('COMMANDE', 'CALCULDATEECHE', 'TRUE');

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
      Commande := '"' + Application.ExeName + '" /USER='+V_PGI.UserLogin+ ' /PASSWORD='+V_PGI.PassWord+' /DOSSIER='+Dossier+ ' "/INI='+CurrentFile+';IMPORT;Minimized"'
    else
      Commande := '"' + Application.ExeName + '" /USER='+V_PGI.UserLogin+ ' /PASSWORD='+V_PGI.PassWord+' /DOSSIER='+V_PGI.Currentalias+ ' "/INI='+CurrentFile+';IMPORT;Minimized"';
    Writeln(Fbat, Commande) ;
    CloseFile(Fbat);
    PGIInfo ('La génération du fichier de commande est terminée');
end;

procedure TFAssistImport.CTRPIECEClick(Sender: TObject);
{$IFDEF COMSX}
var
formatfic  : integer;
{$ENDIF}
begin
  inherited;
{$IFDEF COMSX}
     if FICHENAME.Text = '' then exit;
     formatfic := 4;  // format etendu
     VisuLignesErreurs(FICHENAME.Text,'',formatfic,TRUE) ;
{$ENDIF}
end;

// fiche
procedure TFAssistImport.DATEECR1Exit(Sender: TObject);
var
Exerc         : TExoDate;
begin
  inherited;
    if (NatureFichier <> 'JRL') then
       DATEECR1.Text := stDate1900
    else
    begin
          if not CQuelExercice (StrToDate(DATEECR1.Text), Exerc) then
          begin
               PGIBox(ERR_DATEECRITURE);
               DATEECR1.Text := stDate1900;
          end;
    end;
end;

procedure TFAssistImport.BBlancClick(Sender: TObject);
begin
  inherited;
if BBlanc.checked then
   MODIFTIERS.Enabled := TRUE
else
   MODIFTIERS.Enabled := FALSE;
end;

procedure TFAssistImport.BMAJCPTEClick(Sender: TObject);
begin
  inherited;
if BMAJCPTE.checked then
   BBlanc.Enabled := TRUE
else
   BBlanc.Enabled := FALSE;
end;

procedure TFAssistImport.RenvoieCorpsMail (var  Hst : HTStringList);
var
i   : integer;
begin
         if Hst = nil then Hst := HTStringList.create;
         for i := 0 to FCorpsMail.Lines.Count-1 do
               Hst.Add(FCorpsMail.Lines[i]);
end;

{$IFDEF CERTIFNF}
Procedure TFAssistImport.MajJournalImport (Filename : string);
var
i     : integer;
stBlob : string;
TB     : TOB;
StrStream : TStringStream;
Libelle, NFichier : string;
begin
        Case NatureFichier[1]  of
             'J' : NFichier := 'Journal';
             'D' : NFichier := 'Dossier';
             'B' : NFichier := 'Balance';
         end;
         Libelle := 'Import ' + NFichier ;
         if not OKexercice then  Libelle := Libelle + 'Statut : ERREUR ' ;

         TB := TOB.Create ('CPJALIMPORT', nil,-1);
         TB.Putvalue ( 'CPJ_CODE' , CpjCode);
         TB.Putvalue ( 'CPJ_LIBELLE' ,  Libelle );
         TB.Putvalue ( 'CPJ_DATE' ,  DATE);
         TB.Putvalue ( 'CPJ_ORIGINE' ,  QualifOrigine);
         TB.Putvalue ( 'CPJ_NOMFICHIER' ,  ExtractFileName(FileName));
         TB.Putvalue ( 'CPJ_CREERPAR' ,  OrigineEnvoie);

         StrStream := TStringStream.Create(stBlob);
         for i := 0 to  LISTEIMPORT.count-1 do
          StrStream.WriteString(LISTEIMPORT.Items[i]+#10);
         stBlob := StrStream.DataString;
         StrStream.Free;

         TB.Putvalue ( 'CPJ_RAPPORT' ,  stBlob);

         TB.InsertOrUpdateDB(TRUE);
         TB.free;
end;
{$ENDIF}


end.

