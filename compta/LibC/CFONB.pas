{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 14/04/2003
Modifié le ... : 21/07/2006
Description .. : Passage en eAGL
Suite ........ : 
Suite ........ : Version    Date        Auteur  Fiche Qualité
Suite ........ : 6.50       28/07/2005  MD      FQ16013
Suite ........ : 7.10       21/07/2006  SB      Dev 4184
Mots clefs ... : 
*****************************************************************}
unit CFONB;

interface

uses
  Forms,
  Classes,
  Controls,
  Dialogs,
  StdCtrls,
  Buttons,
  ComCtrls,
  ExtCtrls,
  SysUtils,      // FormatDateTime, Date, FormatFloat, IntToStr, DateToStr, FormatDateTime, StrToDate, ExtractFileName, ExtractFilePath, Trim
  Hctrls,
  HMsgBox,
  SaisUtil,      // TSuiviMP, TOBM,
  Ent1,          // VH, DirDefault, ChargeSocieteHalley
  HEnt1,         // Format_String, V_PGI
  HSysMenu,
  UProcGen, //RPos //SG6 24/01/05
  UTOB,
{$IFDEF EAGLCLIENT}
  UtileAGL,      // LanceEtat
  BANQUECP_TOM,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  EdtREtat,      // LanceEtat
  {$IFNDEF TRESO}
  {$ENDIF}
  BANQUECP_TOM,
{$ENDIF}

  HStatus,       // InitMove, MoveCur, FiniMove
  SaisComm,      // OBMToIdent
  uLibEcriture,  // function WhereEcritureTOB
  uLibWindows,   // AglUppercase

  {$IFNDEF TRESO}
  CFONBFIC,      // VisuExportCFONB
  {$ENDIF}

  ed_tools,
  HTB97,
  ParamSoc,      // GetParamSoc
  UtilSoc,       // ChargePageSoc, SauvePageSoc, InterfaceSoc
  UtilPGI       // DecodeRIB, EncodeRib, getSchemaName
  {$IFNDEF TRESO}
  ,EtbUser, Mask        // z_GetDestinataires, z_Teletransmission
  {$ENDIF}
  ;

{$IFNDEF TRESO}
Procedure ExportCFONB ( Enc : boolean ; CodeBanque,FormatCFONB,EnvoiTrans : String ; TECHE : TList ; smp : TSuiviMP = smpAucun; ModeleBordereau : String = '') ;
Function  WhereMulti ( TL : TList ) : String ;
Function  GetNumCFONB ( Cat : String ) : String ;
Procedure ExportCFONBMS ( vTobScenario : TOB ; vTLEche : TList ; vBoMultiSoc : Boolean = True ) ;
function  ParamCFONB ( vTobScenario : TOB ) : Boolean ;
Procedure ExportCFONBBatch ( vTobScenario : TOB ; vTLEche : TList ; vBoMultiSoc : Boolean = True ) ;
{$ENDIF}


Type Tot_Cli = Class
     Auxiliaire,RIB,LastMP,LastLib,Devise : String ;
     CodeAccept : String; // FQ 12336
     Debit,Credit : Double ;
     LastEche,LastDate : TDateTime ;
     TIDTIC : Boolean ;
     szRef : String;
     Libelle : string ; // pour éviter les multiples reqêtes sur les tables tiers / généraux
     exporte : boolean ;
     END ;

  {JP 23/09/03 : Infos d'un compte bancaire}
  TCompte = record
    General      : string;
    CodeBanque   : string;
    NumCompte    : string;
    CodeGuichet  : string;
    CleRib       : string;
    Domiciliation: string;
    RaisonSoc    : string;
    Divers       : string;
  end;

  {JP 23/09/03 : Infos pour un transfert}
  TTransfert = record
    CodeBIC  : string;
    CodeIBAN : string;
    Devise   : string;
    Nom      : string;
    Adresse1 : string;
    Ville    : string;
    Agence   : string;
    Pays     : string;
    Pays2    : string;
    CodePays : string;
    CodeIBAN2: string;
    Devise2  : string;
    IdClient : String;
    TypeRemise : String;
    Remise   : string;
    NatEco   : string;
    DateExect: TDateTime;
    Banque   : String;
    TypeDevise : string;
  end;

  {JP 23/09/03 : Infos générales de traitement}
  TInfosCFONB = record
    Code        : string;
    NumEmetteur : string;
    Motif       : string;
    Imputation  : string;
    Montant     : Double;
    DateCreat   : TDateTime;
  end;

  {JP 23/09/03 : Infos concernant la société}
  TInfosSoc = record
    Raison : string;
    Adr1   : string;
    Ville  : string;
    Pays   : string;
    Siret  : string;
  end;

type
  {$IFNDEF TRESO}
  TFEXPCFONB = class(TForm)
    TypeExport: TRadioGroup;
    Outils: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    H_GENERAL: THLabel;
    BQ_GENERAL: THValComboBox;
    H_NOMFICHIER: THLabel;
    NomFichier: TEdit;
    RechFile: TToolbarButton97;
    Sauve: TSaveDialog;
    HM: THMsgBox;
    H_Document: THLabel;
    Document: THValComboBox;
    HMTrad: THSystemMenu;
    BRIB: THBitBtn;
    BCompteRendu: THBitBtn;
    BParam: THBitBtn;
    Pages: TPageControl;
    TS1: TTabSheet;
    TS2: TTabSheet;
    CRibTiers: TCheckBox;
    CCumulTiers: TCheckBox;
    H_DATEREMISE: THLabel;
    DATEREMISE: THCritMaskEdit;
    RefTireLib: TCheckBox;
    FApercu: TCheckBox;
    REnvoi: TRadioGroup;
    cTeleTrans: TCheckBox;
    TCDestinataire: THLabel;
    CDestinataire: THValComboBox;
    TS3: TTabSheet;
    TypeDebit: TRadioGroup;
    ImputFrais: TRadioGroup;
    cbo_NatEco: THValComboBox;
    lbl_NatEco: THLabel;
    lbl_Remise: THLabel;
    txt_RefRemise: TEdit;
    procedure RechFileClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure TypeExportClick(Sender: TObject);
    procedure BQ_GENERALChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BCompteRenduClick(Sender: TObject);
    procedure BRIBClick(Sender: TObject);
    procedure BParamClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure DocumentChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    GeneCharge,ExisteMauvais,ExportEnEuro : boolean ;
    OkDebit : boolean ;
    FF       : TextFile ;
    NumE,NbDec : integer ;
    TotalG   : Double ;
    NumEmetteur,LastExport,Destinataire : String ;
    gszDateEcheance : String;
    gszDateEcheanceBis : TDateTime; {JP 24/09/03}
    szModeleBordereau : String;
    // GCO - 19/07/2006 - FQ 18603
    FStDevise : string; // Remplit avec la devise si monodevise, sinon ''
    Function  EmetOK ( NumEmet,Dom,Etab,Guichet,NumCompte : String ) : boolean ;
    procedure GetNomFichier ;
    function PutLeRIB ( O : TOB ) : Boolean ; // FQ 12336
    Function  GetLeRIB ( AuxiGen,RIB : String ; Var Etab,Guichet,NumCompte,Cle,Dom : String ) : boolean ;
    Function  ZeroG ( St : String ; L : integer ) : String ;
    Function  Date5 ( DD : TDateTime ) : String ;
    Procedure SommeCli ( O : TOB ) ;
    Function  ControleOk ( App : boolean ) : boolean ;
    Function  LanceCFONB ( App : Boolean ) : boolean ;
    procedure EditeCFONB ;
    procedure FlagCFONB ;
    Function  TransETEBAC ( NomComplet,Destinataire : String ) : boolean ;
    Procedure Get_Compte ( Q : TQuery ; Var Dom,Etab,Guichet,NumCompte : String ) ;
    Function  EmetteurCFONB : boolean ;
    Function  EmetteurLCR : boolean ;
    Function  EmetteurVIR : boolean ;
    Function  EmetteurPRE : boolean ;
    Function  EmetteurTRANS : boolean ;
    Function  DetailCFONB ( O : TOB ) : boolean ;
    Function  DetailLCR ( O : TOB ) : boolean ;
    Function  DetailVIR ( O : TOB ) : boolean ;
    Function  DetailPRE ( O : TOB ) : boolean ;
    Function  DetailTRANS ( O : TOB ) : boolean ;
    Function  DetailTRANS2( O : TOB ) : boolean ;
    Function  SommeDetailPRE ( X : Tot_Cli ) : boolean ;
    Function  SommeDetailLCR ( X : Tot_Cli ) : boolean ;
    Function  SommeDetailVIR ( X : Tot_Cli ) : boolean ;
    Function  SommeDetailTRANS ( X : Tot_Cli ) : boolean ;
    Function  SommeDetailTRANS2( X : Tot_Cli ) : boolean ;
    Function  SommeDetailCFONB ( X : Tot_Cli ) : boolean ;
    Procedure TotalCFONB ;
    Procedure TotalLCR ;
    Procedure TotalVIR ;
    Procedure TotalPRE ;
    Procedure TotalTRANS ;
    Procedure GereDateRem ;
    Procedure RemplirLesEches ;
    //SG6 18/01/05
    procedure TraitementFichier(var App:boolean);
    {JP 05/07/05 : FQ 15493 : Nouvelle numérotation des fichiers}
    function ExtractFileNameOnly(FileName : string; RacineNom : Boolean): string;
    {JP 05/07/05 : FQ 15493 : Génère un nom de fichier}
    procedure GenereNomFichier (Nom_Fichier : string);
    {JP 05/07/05 : FQ 16161 : Récupération des infos rib du Tiers}
    function GetInfosRibCompte(AuxiGen, RIB : string; var Ville, Pays, Bic, NomBQ : string) : string;
    // Gestion du mode par lot Dev 4184
    function  VerifParam : Boolean ; // Tests de paramètres
    procedure RecupParam ;           // Recup des paramètres
    procedure SetParam ;             // Mise en place des paramètres
    function ExecuteBatch( OkAppend : Boolean ) : Boolean ;  // Génération du fichier en mode batch
    procedure DeflagCFONBBatch ;
    function findTotCli( vAux: string ; vDev : string ): Tot_Cli;
    Function RecupGenOuAux  ( O : TOB ; AvecLib : boolean ; Var Cpt,Lib : String ) : Boolean ; // Opti récup libellé tiers
    function GetTobMP( vCodeMP : string ) : Tob ; // opti sur mode de paiement

    // Fin Dev 4184
  public
    CodeBanque,FormatCFONB,EnvoiTrans : String ;
    Enc                    : boolean ;
    smp                    : TSuiviMP ;
    TECHE,TCli,TECHES,TECHE1 : TList ;
    FBoMultiSoc              : Boolean ;
    FTobScenario             : Tob ;        // Tob de transit pour utilisation batch via scaneio de génération
    FBoRecupParam            : Boolean ;    // Mode saisie des paramètres (sans lancement de l'export)
    FBoModeBatch             : Boolean ;    // Génération batch du ficher, sans affichage de l'interface
    FTobMP                   : Tob ;        // Tob de mode de paiement pour opti base
  end;
  {$ENDIF}

  TGenerationCFONB = class
  public
    class function CompleteZero(St : string; L : Integer) : string;
    class function Date5       (DD : TDateTime) : string ;
    {Virements}
    class function EmetteurVIR(Cpt : TCompte; Inf : TInfosCFONB) : string;
    class function DetailVIR  (Cpt : TCompte; Inf : TInfosCFONB) : string;
    class function TotalVIR   (Inf : TInfosCFONB) : string;
    {Transferts}
    class function EmetteurTrans(Cpt : TTransfert; Inf : TInfosCFONB; Soc : TInfosSoc) : string;
    class function DetailTrans  (Cpt : TTransfert; Inf : TInfosCFONB) : string;
    class function DetailTrans2 (Cpt : TTransfert; Inf : TInfosCFONB) : string;
    class function TotalTrans   (Cpt : TTransfert; Inf : TInfosCFONB; Soc : TInfosSoc) : string;
  end;

implementation

{$IFNDEF TRESO}
{$R *.DFM}
uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  ULibEncaDeca,
  cbpPath,
  Commun {SetPlusBanqueCp};
{$ENDIF}

{$IFNDEF TRESO}
Procedure ExportCFONBBatch ( vTobScenario : TOB ; vTLEche : TList ; vBoMultiSoc : Boolean = True ) ;
Var lStFormatCFONB : string ;
    X              : TFEXPCFONB ;
begin

  X := TFEXPCFONB.Create(Application) ;

  try

    if vTobScenario.GetValue('CPG_TYPEENCADECA') = 'LCR'
      then lStFormatCFONB := 'LEN'
      else lStFormatCFONB  := vTobScenario.GetValue('CPG_TYPEENCADECA') ;

    X.CodeBanque        := vTobScenario.GetString('CPG_GENERAL') ;
    X.TECHE             := vTLEche ;
    X.smp               := smpAucun ;
    X.FormatCFONB       := lStFormatCFONB ;
    X.Enc               := vTobScenario.GetString('CPG_FLUXENCADECA')='ENC' ;
    X.EnvoiTrans        := '' ;
    X.szModeleBordereau := '';
    X.FBoMultiSoc       := vBoMultiSoc ;
    X.FTobScenario      := vTobScenario ;
    X.FBoRecupParam     := False ;
    X.FBoModeBatch      := True ;

    X.SetParam ;
    X.ExecuteBatch( vTobScenario.GetValue('CFONB_PREMIERLOT')<>'X' ) ;

  finally

    X.Free ;

  end ;

end ;

Procedure ExportCFONBMS ( vTobScenario : TOB ; vTLEche : TList ; vBoMultiSoc : Boolean = True ) ;
Var lStModeleBOR   : string ;
    lStFormatCFONB : string ;
    X              : TFEXPCFONB ;
begin

  X := TFEXPCFONB.Create(Application) ;

  try

    if (vTobScenario.GetValue('CPG_BORDEREAUEXP') = 'X')
      then lStModeleBOR := vTobScenario.GetValue('CPG_BORDEREAUMOD')
      else lStModeleBOR := '' ;

    if vTobScenario.GetValue('CPG_TYPEENCADECA') = 'LCR'
      then lStFormatCFONB := 'LEN'
      else lStFormatCFONB  := vTobScenario.GetValue('CPG_TYPEENCADECA') ;

    X.CodeBanque        := vTobScenario.GetString('CPG_GENERAL') ;
    X.TECHE             := vTLEche ;
    X.smp               := smpAucun ;
    X.FormatCFONB       := lStFormatCFONB ;
    X.Enc               := vTobScenario.GetString('CPG_FLUXENCADECA')='ENC' ;
    X.EnvoiTrans        := '' ;
    X.szModeleBordereau := lStModeleBOR;
    X.FBoMultiSoc       := vBoMultiSoc ;
    X.FTobScenario      := vTobScenario ;
    X.FBoRecupParam     := False ;
    X.FBoModeBatch      := False ;
    X.ShowModal ;

  finally

    X.Free ;

  end ;

end ;

function ParamCFONB ( vTobScenario : TOB ) : Boolean ;
Var lStModeleBOR   : string ;
    lStFormatCFONB : string ;
    X              : TFEXPCFONB ;
begin

  X := TFEXPCFONB.Create(Application) ;

  try

    if (vTobScenario.GetValue('CPG_BORDEREAUEXP') = 'X')
      then lStModeleBOR := vTobScenario.GetValue('CPG_BORDEREAUMOD')
      else lStModeleBOR := '' ;

    if vTobScenario.GetValue('CPG_TYPEENCADECA') = 'LCR'
      then lStFormatCFONB := 'LEN'
      else lStFormatCFONB  := vTobScenario.GetValue('CPG_TYPEENCADECA') ;

    if vTobScenario.GetNumChamp('CFONB_PARAMOK') > 0
      then vTobScenario.PutValue('CFONB_PARAMOK', '-')
      else vTobScenario.AddChampSupValeur( 'CFONB_PARAMOK', '-' ) ;


    X.CodeBanque        := vTobScenario.GetString('CPG_GENERAL') ;
    X.TECHE             := nil ;
    X.smp               := smpAucun ;
    X.FormatCFONB       := lStFormatCFONB ;
    X.Enc               := vTobScenario.GetString('CPG_FLUXENCADECA')='ENC' ;
    X.EnvoiTrans        := '' ;
    X.szModeleBordereau := lStModeleBOR;
    X.FBoMultiSoc       := False ;
    X.FTobScenario      := vTobScenario ;
    X.FBoRecupParam     := True ;
    X.FBoModeBatch      := False ;
    X.ShowModal ;
    
    result := ( vTobScenario.getValue('CFONB_PARAMOK') = 'X' ) ;

  finally

    X.Free ;

  end ;

end ;


Procedure ExportCFONB ( Enc : boolean ; CodeBanque,FormatCFONB,EnvoiTrans : String ; TECHE : TList ; smp : TSuiviMP = smpAucun; ModeleBordereau : String = '') ;
Var X : TFEXPCFONB ;
BEGIN
X:=TFEXPCFONB.Create(Application) ;
 Try
  X.CodeBanque        := CodeBanque ;
  X.TECHE             := TECHE ;
  X.smp               := smp ;
  X.FormatCFONB       := FormatCFONB ;
  X.Enc               := Enc ;
  X.EnvoiTrans        := EnvoiTrans ;
  X.szModeleBordereau := ModeleBordereau;
  X.FBoMultiSoc       := False ;
  X.FTobScenario      := nil ;
  X.FBoRecupParam     := False ;
  X.FBoModeBatch      := False ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
END ;

Function  GetNumCFONB ( Cat : String ) : String ;
Var St : String ;
    Num  : integer ;
BEGIN
St:=Cat+Format_String(V_PGI.User,3)+FormatDateTime('YYMMDD',Date) ;
Num:=GetParamSocSecur('SO_CPNUMCFONB',0)+1 ;
if Num>=100000 then Num:=1 ;
SetParamSoc('SO_CPNUMCFONB',Num) ;
Result:=St+FormatFloat('00000',Num) ;
END ;

function WhereMulti(TL : TList) : String;
Var
    SWhere : String;
    i : integer;
    O : TOB;
begin
    SWhere := '';

    for i:=0 to TL.Count-1 do
    begin
        O := TOB(TL[i]);
// BPY le 16/09/2004 : Fiche n° 14489 => edition de tt les bordereau a la place de juste celui que l'on export !
//        SWhere := SWhere + '(E_JOURNAL="'+O.GetMvt('E_JOURNAL')+'" AND E_NUMEROPIECE='+IntToStr(O.GetMvt('E_NUMEROPIECE'))+' AND E_QUALIFPIECE="'+O.GetMvt('E_QUALIFPIECE')
//                         + '" AND E_EXERCICE="'+O.GetMvt('E_EXERCICE')+'" AND E_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE'))+')' ;
        SWhere := SWhere + '(E_JOURNAL="' + O.GetValue('E_JOURNAL') + '" AND '
                         + 'E_EXERCICE="' + O.GetValue('E_EXERCICE') + '" AND '
                         + 'E_DATECOMPTABLE="' + UsDateTime(O.GetValue('E_DATECOMPTABLE')) + '" AND '
                         + 'E_NUMEROPIECE=' + IntToStr(O.GetValue('E_NUMEROPIECE')) + ' AND '
                         + 'E_NUMLIGNE=' + IntToStr(O.GetValue('E_NUMLIGNE')) + ' AND '
                         + 'E_NUMECHE=' + IntToStr(O.GetValue('E_NUMECHE')) + ' AND '
                         + 'E_QUALIFPIECE="' + O.GetValue('E_QUALIFPIECE') + '")';
// Fin BPY
        if (i < TL.Count-1) then SWhere := SWhere + ' OR ';
    end;
    if (TL.Count > 1) then SWhere := '(' + SWhere + ')';
    result := SWhere;
end;

Procedure TFEXPCFONB.GetNomFichier ;
Var Q : TQuery ;
    StFichier,StModele : String ;
BEGIN
StFichier:='' ; StModele:='' ; Destinataire:='' ;
if FBoModeBatch then Exit ;
Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+CodeBanque
          +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ;   // 19/10/2006 YMO Multisociétés
if Not Q.EOF then
   BEGIN
   Case TypeExport.ItemIndex of
    0,1,2,3,4 : BEGIN
                FormatCFONB:='LCR' ;
                StFichier:=Q.FindField('BQ_REPLCR').AsString ;
                StModele:=Q.FindField('BQ_LETTRELCR').AsString ;
                CCumulTiers.Checked:=False ;
                END ;
    5,6,7 : BEGIN
          FormatCFONB:='VIR' ;
          StFichier:=Q.FindField('BQ_REPVIR').AsString ;
          StModele:=Q.FindField('BQ_LETTREVIR').AsString ;
          CCumulTiers.Checked:=False ;
          END ;
      9 : BEGIN
          FormatCFONB:='TRI' ;
          StFichier:=Q.FindField('BQ_REPVIR').AsString ;
          StModele:=Q.FindField('BQ_LETTREVIR').AsString ;
          CCumulTiers.Checked:=False ;
          END ;
      8 : BEGIN
          FormatCFONB:='PRE' ;
          StFichier:=Q.FindField('BQ_REPPRELEV').AsString ;
          StModele:=Q.FindField('BQ_LETTREPRELV').AsString ;
          CCumulTiers.Checked:=True ;
          END ;
      END ;
   Destinataire:=Q.FindField('BQ_DESTINATAIRE').AsString ;
   END ;
Ferme(Q) ;
if StFichier='' then
   BEGIN
   Case TypeExport.ItemIndex of
    0,1,2,3,4 : StFichier := TcbpPath.GetCegidUserLocalAppData + '\LCRBOR.TXT' ;
    5,6,7,9   : StFichier := TcbpPath.GetCegidUserLocalAppData + '\VIREMENT.TXT' ;
      8       : StFichier := TcbpPath.GetCegidUserLocalAppData + '\PRELEV.TXT' ;
      END ;
   END ;
NomFichier.Text:=StFichier ;
if StModele<>'' then Document.Value:=StModele ;
cDestinataire.Value:=Destinataire ;
END ;

Procedure TFEXPCFONB.GereDateRem ;
BEGIN
Case TypeExport.ItemIndex of
   0,1,2,3,4 : BEGIN DateRemise.Enabled:=False ; DateRemise.Text:=DateToStr(V_PGI.DateEntree) ; H_DateRemise.Caption:=HM.Mess[17] ; END ;
   5,6,7,9 : BEGIN DateRemise.Enabled:=True  ; H_DateRemise.Caption:=HM.Mess[19] ; END ;
     8 : BEGIN DateRemise.Enabled:=True  ; H_DateRemise.Caption:=HM.Mess[18] ; END ;
   END ;
END ;

Function TFEXPCFONB.Date5 ( DD : TDateTime ) : String ;
Var St : String ;
BEGIN
St:=FormatDateTime('ddmmyy',DD) ;
Result:=Copy(St,1,4)+Copy(St,6,1) ;
END ;

procedure TFEXPCFONB.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,NomFichier.Text) ;
if Sauve.Execute then NomFichier.Text:=Sauve.FileName ;
end;

procedure TFEXPCFONB.FormShow(Sender: TObject);
begin
{JP 29/01/07 : gestion du partage de BanqueCP}
SetPlusBanqueCp(BQ_GENERAL);

GeneCharge:=True ; BCompteRendu.Enabled:=False ;
BQ_GENERAL.Value:=CodeBanque ; BRIB.Enabled:=(CodeBanque<>'') ;
if FormatCFONB='VRI' then FormatCFONB := 'TRI';
if FormatCFONB='VIR' then TypeExport.ItemIndex:=5 else
if FormatCFONB='PRE' then TypeExport.ItemIndex:=8 else
if FormatCFONB='LEN' then TypeExport.ItemIndex:=0 else
if FormatCFONB='LES' then TypeExport.ItemIndex:=1 Else
if FormatCFONB='LEV' then TypeExport.ItemIndex:=2 Else
if FormatCFONB='LEX' then TypeExport.ItemIndex:=3 Else
if FormatCFONB='LEZ' then TypeExport.ItemIndex:=4 else
if FormatCFONB='VI2' then TypeExport.ItemIndex:=6 else
if FormatCFONB='VI3' then TypeExport.ItemIndex:=7 else
if FormatCFONB='TRI' then TypeExport.ItemIndex:=9;
If VH^.OldTeletrans Then z_GetDestinataires(cDestinataire.Values,cDestinataire.Items) ;
GetNomFichier ; DateRemise.Text:=DateToStr(V_PGI.DateEntree) ;
GeneCharge:=False ; LastExport:='' ;
GereDateRem ;
If Not VH^.OldTeletrans Then EnvoiTrans:='NON' ;
if EnvoiTrans='NON' then BEGIN cTeleTrans.Enabled:=False ; cTeleTrans.Checked:=False ; END else
 if EnvoiTrans='AUT' then BEGIN cTeleTrans.Enabled:=False ; cTeleTrans.Checked:=True ; END else
  if ((EnvoiTrans='AUC') or (EnvoiTrans='')) then BEGIN cTeleTrans.Enabled:=True ; cTeleTrans.Checked:=False ; END ;
If Not VH^.OldTeletrans Then
  BEGIN
  cTeleTrans.Enabled:=FALSE ; cTeleTrans.Checked:=False ; CDestinataire.Enabled:=FALSE ; TCDestinataire.Enabled:=FALSE ;
  END ;
lbl_NatEco.Visible :=False; cbo_NatEco.Visible := False; TS3.TabVisible := False;
txt_RefRemise.visible :=False; lbl_Remise.Visible:= False;
If SMP<>smpAucun Then
  BEGIN
  TypeExport.Enabled:=FALSE ;
  Case SMP Of
    smpDecBorDec   : TypeExport.ItemIndex:= 0 ;
    smpDecVirBqe   : TypeExport.ItemIndex:= 5 ;
    smpDecVirInBqe,
    smpDecVirInEdt : begin
                       TypeExport.ItemIndex:= 9 ;
                       REnvoi.ItemIndex :=1;
                       TS3.TabVisible := True;
                       lbl_NatEco.Visible := True;
                       cbo_NatEco.Visible := True;
                       txt_RefRemise.Visible :=True;
                       lbl_Remise.Visible:=True;
                       RefTireLib.Visible := False;
                       BRIB.Hint := 'Modifier l''IBAN';
                     end;
                     
    {JP 06/06/06 : FQ 18284 : si on demande un transfert international depuis décaissement divers,
                   il faut rendre visible les zones idoines}
    smpDecDiv : if (FormatCFONB = 'TRI') or (FormatCFONB = 'VRI') then begin
                  REnvoi.ItemIndex := 1;
                  TS3.TabVisible := True;
                  lbl_NatEco   .Visible := True;
                  cbo_NatEco   .Visible := True;
                  txt_RefRemise.Visible := True;
                  lbl_Remise   .Visible := True;
                  RefTireLib   .Visible := False;
                  BRIB.Hint := 'Modifier l''IBAN';
                end;
    smpEncPreBqe   : TypeExport.ItemIndex:= 8 ;
    smpEncTraEsc   : BEGIN TypeExport.ItemIndex:= 1 ; TypeExport.Enabled:=TRUE ; END ;
    smpEncTraEnc   : BEGIN TypeExport.ItemIndex:= 0 ; TypeExport.Enabled:=TRUE ; END ;
    END ;
  END ;
if szModeleBordereau <> '' then Document.Value := szModeleBordereau; // 12339
{JP 15/05/07 : FQ 19421 : Pour gérer le RefTireLib}
if TypeExport.ItemIndex < 0 then TypeExport.ItemIndex := 0;
TypeExportClick(TypeExport);
end;

Procedure TFEXPCFONB.SommeCli ( O : TOB ) ;
Var Aux,RIB,MP,Lib : String ;
    D,C     : double ;
    i       : integer ;
    X       : Tot_Cli ;
    Find    : boolean ;
    Eche,DateC : TDateTime ;
    TIDTIC : Boolean ;
    Lib2,StRIB,szDev : String ;
    lTobMP      : Tob ;
begin
  if Not CCumulTiers.Checked then Exit ;
  TIDTIC:=RecupGenOuAux(O,True,Aux,Lib2) ; // SBO 26/10/2006 : passage ée arg à True pour récup le libellé
  { b md FQ16013 28/07/2005 }
  if Formatcfonb = 'TRI' then begin
    D:=O.GetValue('E_DEBITDEV') ; C:=O.GetValue('E_CREDITDEV') ;
  end
  else begin
    D:=O.GetValue('E_DEBIT') ; C:=O.GetValue('E_CREDIT') ;
  end;
 { e md }

  MP      := O.GetValue('E_MODEPAIE') ;
  Eche    := O.GetValue('E_DATEECHEANCE') ;
  DateC   := O.GetValue('E_DATECOMPTABLE') ;
  StRIB   := FullMajuscule(O.GetValue('E_RIB')); if StRIB<>'' then RIB:=StRIB ;
  szDev   := O.GetValue('E_DEVISE');
  if DateC>StrToDate(DATEREMISE.Text) then
    DateC:=StrToDate(DATEREMISE.Text) ;
  if ReftireLib.Checked
    then Lib := FullMajuscule(O.GetValue('E_LIBELLE'))
    else Lib := FullMajuscule(O.GetValue('E_REFEXTERNE'));
  Find:=False ;
  for i:=0 to TCli.Count-1 do begin
    X:=Tot_Cli(TCli[i]) ;
    if (X.Auxiliaire=Aux) And (X.TIDTIC=TIDTIC) and (X.Devise=szDev) then begin
       X.Debit:=X.Debit+D ; X.Credit:=X.Credit+C ;
       If Eche>X.LastEche Then X.LastEche:=Eche ;
       Find:=True ;

       // FQ 12336 : Si le code accpetation est différent du précédent : Récupère celui du mode de paiement
       if not (X.CodeAccept = O.GetValue('E_CODEACCEPT')) then
         begin
         lTobMP := GetTobMP( MP ) ;
         if lTObMP <> nil then
           X.CodeAccept := lTobMP.GetString( 'MP_CODEACCEPT' ) ;
         end;
       Break ;
    end;
  end;
  if Not Find then begin
    X:=Tot_Cli.Create ;
    X.Auxiliaire:=Aux ; X.Rib:=RIB ; X.Debit:=D ; X.Credit:=C ; X.Devise:=szDev;
    X.LastMP:=MP ; X.LastEche:=Eche ; X.LastDate:=DateC ; X.lastLib:=Lib ;
    X.CodeAccept := O.GetValue('E_CODEACCEPT'); // FQ 12336
    X.TIDTIC:=TIDTIC ;
    X.szRef := FullMajuscule(O.GetValue('E_REFINTERNE'));
    X.Libelle := Lib2 ;
    X.Exporte := False ;
    TCli.Add(X) ;
  end;
END ;

Function TFEXPCFONB.ControleOk ( App : boolean ) : boolean ;
Var i : integer ;
    O : TOB ;
    Deja,DiffMP,DiffBQE : boolean ;
    OldMP,OldBQE : String3 ;
    OkD,OkC : boolean ;
    QQ      : TQuery ;
    CptB    : String ;
BEGIN
ExisteMauvais:=False ; VideListe(TCLI) ;
Result:=False ; Deja:=False ; OkD:=False ; OkC:=False ; OldMP:='' ; OldBQE:='' ;
DiffMP:=False ; DiffBQE:=False ;
if TypeExport.ItemIndex<0 then BEGIN HM.Execute(10,'','') ; Exit ; END ;
if BQ_GENERAL.Value='' then BEGIN HM.Execute(0,'','') ; Exit ; END ;
if NomFichier.Text='' then BEGIN HM.Execute(1,'','') ; Exit ; END ;
if ((Not App) and (Document.Value='')) then BEGIN if HM.Execute(3,'','')<>mrYes then Exit ; END ;
if ((Not App) and (LastExport<>'')) then BEGIN if HM.Execute(6,'','')<>mrYes then Exit ; END ;

if TECHE.Count > 0 then
  FStDevise := TOB(TECHE[0]).GetString('E_DEVISE');

for i:=0 to TECHE.Count-1 do
    BEGIN
    O:=TOB(TECHE[i]);

    // GCO - 19/07/2007 - FQ 18603
    if (O.GetString('E_DEVISE') <> FStDevise) and (FStDevise <> '') then
      FStDevise := '';
    // FIN GCO

    if O.GetValue('E_CFONBOK')='X' then Deja:=True ;
    if O.GetValue('E_DEBIT')<>0 then OkD:=True ;
    if O.GetValue('E_CREDIT')<>0 then OkC:=True ;
    if i=0 then
       BEGIN
       OldMP:=O.GetValue('E_MODEPAIE') ;
       OldBQE:=O.GetValue('E_JOURNAL') ;
       END else
       BEGIN
       if O.GetValue('E_MODEPAIE')<>OldMP then DiffMP:=True ;
       if O.GetValue('E_JOURNAL')<>OldBQE then DiffBQE:=True ;
       END ;
    if not PutLeRIB(O) then exit; {Affectation du RIB} // VL 10/10/2003 FQ 12336
    SommeCli(O) ; {Cumul par tiers pour contrôle}
    END ;
if ((Not DiffBQE) and (Not ENC)) then
   BEGIN
   QQ:=OpenSQL('Select J_CONTREPARTIE from JOURNAL Where J_JOURNAL="'+OldBQE+'"',True) ;
   if Not QQ.EOF then
      BEGIN
      CptB:=QQ.Fields[0].AsString ; if ((CptB<>'') and (CptB<>BQ_GENERAL.Value)) then DiffBQE:=True ;
      END ;
   Ferme(QQ) ;
   END ;
if ((Deja) and (LastExport='') and (Not App)) then BEGIN if HM.Execute(5,'','')<>mrYes then Exit ; END ;
if DiffMP then BEGIN if HM.Execute(12,'','')<>mrYes then Exit ; END ;
if ((DiffBQE) and (Not ENC)) then BEGIN if HM.Execute(20,'','')<>mrYes then Exit ; END ;
if ((OkD) and (ENC)) then BEGIN if HM.Execute(13,'','')<>mrYes then Exit ; END ;
if ((OkC) and (Not ENC)) then BEGIN if HM.Execute(13,'','')<>mrYes then Exit ; END ;
if Not App then if HM.Execute(2,'','')<>mrYes then Exit ;
Result:=True ;
END ;

Procedure TFEXPCFONB.RemplirLesEches ;
Var i : integer ;
    O,OS,O2 : TOB ;
BEGIN
VideListe(TECHE1) ; VideListe(TECHES) ;
if REnvoi.ItemIndex<>1 then Exit ;
for i:=0 to TEche.Count-1 do
    BEGIN
    O  := TOB(TECHE[i]) ;
    OS := TOB.Create ( 'ECRITURE', nil, -1 ) ;
    OS.Dupliquer ( O , true , true );
    TECHES.Add(OS) ;
    O2 := TOB.Create ( 'ECRITURE', nil, -1 ) ;
    O2.Dupliquer ( O , true , true );
    TECHE1.Add(O2);
    END ;
END ;

Function TFEXPCFONB.TransETEBAC ( NomComplet,Destinataire : String ) : boolean ;
Var NomFichier,Chemin,CarteAppel : String ;
    ie : integer ;
BEGIN
Result:=True ;
if Not VH^.OkModEtebac then
   BEGIN
   HShowMessage('0;Télétransmission ETEBAC; Le module Liaison Bancaire n''est pas sérialisé ;W;O;O;O;','','') ;
   Exit ;
   END ;
NomFichier:=ExtractFileName(NomComplet) ;
Chemin:=ExtractFilePath(NomComplet) ;
if Chemin<>'' then if Chemin[Length(Chemin)]<>'\' then Chemin:=Chemin+'\' ;
Case TypeExport.ItemIndex of
   0,1,2,3,4 : CarteAppel:='E000' ;
   5,6,7 : CarteAppel:='E001' ;
     8 : CarteAppel:='E002' ;
   else Exit ;
   END ;
ie:=z_Teletransmission(PChar(NomFichier),PChar(Chemin),PChar(CarteAppel),PChar(Destinataire),1) ;
if ie<>0 then BEGIN if HM.Execute(21,'','')<>mrYes then Result:=False ; END ;
END ;

procedure TFEXPCFONB.BValiderClick(Sender: TObject);
Var io : TIoErr ;
    i : integer ;
    OkTrans,OkAppend : boolean ;
    DateEcheance,OldDateEcheance : TDateTime;
begin
//SG6 19/01/05 FQ 14321
OkAppend := False;

// Mode récup de paramètres uniquement :
//  on sort après validation des paramètres sans effecuter l'export, parmètres stocker dans la tob FTobScenario
if FBoRecupParam then
  begin
  if VerifParam then
    begin
    RecupParam ;
    FTobScenario.PutValue( 'CFONB_PARAMOK',  'X' ) ;
    Close ;
    end ;
  Exit ;
  end ;

// Virement à échéance E-2 ou E-3
DateEcheance:=0;
if (TypeExport.ItemIndex=6) or (TypeExport.ItemIndex=7) then begin
  for i:=0 to TECHE.Count-1 do begin
    if (DateEcheance <> 0) then begin
      OldDateEcheance := StrToDate(TOB(TECHE[i]).GetValue('E_DATEECHEANCE'));
      if (DateEcheance <> OldDateEcheance) then begin
        HM.Execute(22,caption,''); Exit;
      end;
    end else
      OldDateEcheance := StrToDate(TOB(TECHE[0]).GetValue('E_DATEECHEANCE'));
    DateEcheance := OldDateEcheance;
  end;
end;


RemplirLesEches ; ExportEnEuro:=FALSE ; OkTrans:=True ;

 Case REnvoi.ItemIndex of
    0 : BEGIN OkDebit:=VH^.TenueEuro ; ExportEnEuro:=TRUE ; END ;
    1 : BEGIN
        if TECHE.Count>0 then OkDebit:=True;
        ExportEnEuro:=FALSE ;
        If (OkDebit And VH^.TenueEuro) Or ((Not OkDebit) And (Not VH^.TenueEuro)) Then ExportEnEuro:=TRUE ;
        END ;
    END ;
 if OkDebit then NbDec:=V_PGI.OkDecV else NbDec:=V_PGI.OkDecE ;
 if Not ControleOK(False) then Exit ;
 //SG6 18/01/05 FQ 14321
 TraitementFichier(OkAppend);
 if Not LanceCFONB(OkAppend) then Exit ;

if VH^.OldTeletrans And cTeletrans.Checked then OkTrans:=TransETEBAC(NomFichier.Text,cDestinataire.Value) ;
if OkTrans then
   BEGIN
   io:=Transactions(FlagCFONB,2) ;
   Case io of
      oeOk      : HM.Execute(14,'','') ;
      oeUnknown : MessageAlerte(HM.Mess[11]) ;
      oeSaisie  : MessageAlerte(HM.Mess[4]) ;
      END ;
   EditeCFONB ; // 13712
   BCompteRendu.Enabled:=True ; LastExport:=NomFichier.Text ;
   END ;
end;

procedure TFEXPCFONB.TypeExportClick(Sender: TObject);
begin
  if GeneCharge then Exit ;
  if FBoModeBatch then Exit ;
  GetNomFichier ;
  GereDateRem ;
  {JP 15/05/07 : FQ 19421 : Pour gérer le RefTireLib}
  RefTireLib.Visible := (TypeExport.ItemIndex <= 4);
  if not RefTireLib.Visible then RefTireLib.Checked := False; 

  H_DATEREMISE.Visible :=(TypeExport.ItemIndex<=5) or (TypeExport.ItemIndex>=8);
  DATEREMISE.Visible   :=(TypeExport.ItemIndex<=5) or (TypeExport.ItemIndex>=8);
  if (TypeExport.ItemIndex=9) then CRibTiers.Caption := 'Affecter l''&IBAN principal du compte sur les mouvements sans IBAN'
                              else CRibTiers.Caption := 'Affecter le &RIB principal du compte sur les mouvements sans RIB';
  TS3.TabVisible        :=(TypeExport.ItemIndex=9);
  lbl_NatEco.Visible    :=(TypeExport.ItemIndex=9);
  cbo_NatEco.Visible    :=(TypeExport.ItemIndex=9);
  txt_RefRemise.Visible :=(TypeExport.ItemIndex=9);
  lbl_Remise.Visible    :=(TypeExport.ItemIndex=9);
  {JP 05/07/05 : FQ 15493 : Génère un nom de fichier unique à partir de GetNomFichier}
  GenereNomFichier(NomFichier.Text);
end;

{=============================== Methodes d'export ===================================}
Function TFEXPCFONB.EmetOK ( NumEmet,Dom,Etab,Guichet,NumCompte : String ) : boolean ;
BEGIN
Result:=False ;
if Trim(NumEmet)='' then BEGIN HM.Execute(7,'','') ; Exit ; END ;
if ((Trim(Dom)='') or (Trim(Etab)='') or (Trim(Guichet)='') or (Trim(NumCompte)='')) then BEGIN HM.Execute(8,'','') ; Exit ; END ;
Result:=True ;
END ;

Function TFEXPCFONB.ZeroG ( St : String ; L : integer ) : String ;
BEGIN
St:=Copy(St,1,L) ;
While Length(St)<L do St:='0'+St ;
Result:=St ;
END ;

function TFEXPCFONB.PutLeRIB ( O : TOB ) : Boolean ;
Var Q : TQuery ;
    AuxiGen,RIB,Etab,Guichet,NumCompte,Cle,Dom,Iban,Pays : String ;
    Lib : String ;
BEGIN
Result := True;
RecupGenOuAux(O,FALSE,AuxiGen,Lib) ;
RIB:=FullMajuscule(O.GetValue('E_RIB')); Cle:='' ;
// Pour l'IBAN
if Copy(RIB,1,1) = '*' then Exit;
DecodeRIB(Etab,Guichet,NumCompte,Cle,Dom,RIB) ;
if ((Trim(Etab)='') or (Trim(Guichet)='') or (Trim(NumCompte)='')) then
   if CRibTiers.Checked then
   BEGIN
   Q:=OpenSQL('SELECT * from RIB Where R_AUXILIAIRE="'+AuxiGen+'" AND R_PRINCIPAL="X"',True) ;
   if Not Q.EOF then
      BEGIN
      Etab      := FullMajuscule(Q.FindField('R_ETABBQ').AsString);
      Guichet   := Q.FindField('R_GUICHET').AsString ;
      NumCompte := Q.FindField('R_NUMEROCOMPTE').AsString ;
      Cle       := Q.FindField('R_CLERIB').AsString ;
      Dom       := FullMajuscule(Q.FindField('R_DOMICILIATION').AsString);
      Iban      := Q.FindField('R_CODEIBAN').AsString ;
      Pays      := FullMajuscule(Q.FindField('R_PAYS').AsString);
      {Maj OBM Rib}
      if (CodeIsoDuPays(Pays)<>'FR') then
        RIB:='*'+Iban
      else
      RIB:=EncodeRIB(Etab,Guichet,NumCompte,Cle,Dom) ;
      O.PutValue('E_RIB',RIB) ;
      END
      else begin
        HM.Execute(35,'',O.GetValue('E_AUXILIAIRE')+' n''est pas renseigné.');
        Result := False;
      end;
   Ferme(Q) ;
   END
   else begin
     HM.Execute(30,'',O.GetValue('E_AUXILIAIRE')+' n''est pas renseigné.');
     Result := False;
   end;
END ;

Function TFEXPCFONB.GetLeRIB ( AuxiGen,RIB : String ; Var Etab,Guichet,NumCompte,Cle,Dom : String ) : boolean ;
BEGIN
Cle:='' ;
DecodeRIB(Etab,Guichet,NumCompte,Cle,Dom,RIB) ;
Result:=(Trim(RIB)<>'') ;
END ;

{============================= Emetteurs ====================================}
Procedure TFEXPCFONB.Get_Compte ( Q : TQuery ; Var Dom,Etab,Guichet,NumCompte : String ) ;
BEGIN
Dom:=FullMajuscule(Q.FindField('BQ_DOMICILIATION').AsString);
Etab:=FullMajuscule(Q.FindField('BQ_ETABBQ').AsString);
Guichet:=Q.FindField('BQ_GUICHET').AsString ;
NumCompte:=Q.FindField('BQ_NUMEROCOMPTE').AsString ;
END ;

Function TFEXPCFONB.EmetteurLCR ;
Var St,Dom,Etab,Guichet,NumCompte : String ;
    Q  : TQuery ;
    D21 : Char ;
    D80 : Char;{JP 04/08/05 : FQ 15958}
BEGIN
Result:=False ;
{RIB banque cédant}

if FBoModeBatch then
  begin
  if Assigned( FTobScenario ) then
    begin
    Dom         := FTobScenario.GetString( 'BQ_DOMICILIATION') ;
    Etab        := FTobScenario.GetString( 'BQ_ETABBQ') ;
    Guichet     := FTobScenario.GetString( 'BQ_GUICHET') ;
    Numcompte   := FTobScenario.GetString( 'BQ_NUMEROCOMPTE') ;
    NumEmetteur := FTobScenario.GetString( 'CFONB_NUMEMETTEUR') ;
    end ;
  end
else
  begin
  Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+BQ_GENERAL.Value
            +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
  if Not Q.EOF then
     BEGIN
     Get_Compte(Q,Dom,Etab,Guichet,NumCompte) ;
     NumEmetteur:=Q.FindField('BQ_NUMEMETLCR').AsString ;
     END else
     BEGIN
     Dom:='' ; Etab:='' ; Guichet:='' ; NumCompte:='' ; NumEmetteur:='' ;
     END ;
  Ferme(Q) ;
  if Not EmetOK(NumEmetteur,Dom,Etab,Guichet,NumCompte) then Exit ;
  end ;

{Constituer l'enregistrement}
Inc(NumE) ;

  {JP 04/08/05 : FQ 15958 : Mauvaise gestion des positions 79 et 80
               Avant 80 était vide !
               la solution ne semble pas claire mais après analyse avec IB
               ItemIndex  | 79 | 80 |Old 79
               --------------------------
                  0       |  3 |  0 |  3
                  1       |  1 |  0 |  1
                  2       |  2 |  0 |  2
                  3       |  4 |  0 |  4
                  4       |  1 |  1 |  5}
  D80 := '0';
  Case TypeExport.ItemIndex Of
    0 : D21 := '3';
    1 : D21 := '1';
    2 : D21 := '2';
    3 : D21 := '4';
    4 : begin
          D21 := '1';
          D80 := '1';
        end;
    else D21 := '1';
  end;

{A}St:='03' ;
{B}St:=St+'60'+ZeroG(Inttostr(NumE),8)+Format_String(NumEmetteur,6) ;
{C}St:=St+Format_String('',6)+FormatDateTime('ddmmyy',V_PGI.DateEntree)+Format_String(V_PGI.NomSociete,24) ;

{JP 04/08/05 : FQ 15958 : Nouvelle gestion des positions 79 et 80 et suppression des exports en Francs}
{D}St := St + Format_String(Dom, 24) + D21 + D80 + 'E' + ZeroG(Etab, 5) + ZeroG(Guichet, 5) + ZeroG(NumCompte, 11);

{E}St:=St+Format_String('',16) ;
{F}St:=St+Format_String('',6)+Format_String('',10)+Format_String('',15) ;
{G}St:=St+Format_String('',11) ;
WriteLn(FF,St) ;
Result:=True ;
END ;

Function TFEXPCFONB.EmetteurVIR ;
Var St,Dom,Etab,Guichet,NumCompte,Libelle : String ;
    Q  : TQuery ;
    Cpt : TCompte;
    Inf : TInfosCFONB;
BEGIN
Result:=False ;
{RIB banque donneur d'ordre}

if FBoModeBatch then
  begin
  if Assigned( FTobScenario ) then
    begin
    Dom         := FTobScenario.GetString( 'BQ_DOMICILIATION') ;
    Etab        := FTobScenario.GetString( 'BQ_ETABBQ') ;
    Guichet     := FTobScenario.GetString( 'BQ_GUICHET') ;
    Numcompte   := FTobScenario.GetString( 'BQ_NUMEROCOMPTE') ;
    NumEmetteur := FTobScenario.GetString( 'CFONB_NUMEMETTEUR') ;
    Libelle     := FTobScenario.GetString( 'BQ_LIBELLE') ;
    end ;
  end
else
  begin
  Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+BQ_GENERAL.Value
            +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
  if Not Q.EOF then
     BEGIN
     Get_Compte(Q,Dom,Etab,Guichet,NumCompte) ;
     NumEmetteur:=Q.FindField('BQ_NUMEMETVIR').AsString ;
     Libelle:= FullMajuscule(Q.FindField('BQ_LIBELLE').AsString);
     END else
     BEGIN
     Dom:='' ; Etab:='' ; Guichet:='' ; NumCompte:='' ;
     NumEmetteur:='' ; Libelle:='' ;
     END ;
  Ferme(Q) ;
  if Not EmetOK(NumEmetteur,Dom,Etab,Guichet,NumCompte) then Exit ;
  end ;

{Constituer l'enregistrement}
Inc(NumE) ;

  {JP 24/09/03 : Nouvelle fonction}
  case TypeExport.ItemIndex of
    6 : begin
          Inf.Code := '28';
          Cpt.Divers := '';
          Inf.DateCreat := gszDateEcheanceBis;
        end;
    7 : begin
          Inf.Code := '27';
          Cpt.Divers := '7';
          Inf.DateCreat := gszDateEcheanceBis;
        end;
    else begin
      Inf.Code := '02';
      Cpt.Divers := '';
      Inf.DateCreat := StrToDate(DateRemise.Text);
    end;
  end;

  Inf.NumEmetteur   := NumEmetteur;
  Cpt.Domiciliation := FullMajuscule(V_PGI.NomSociete);
  Cpt.CodeGuichet   := Guichet;
  Cpt.NumCompte     := NumCompte;
  Cpt.CodeBanque    := Etab;
  St := TGenerationCFONB.EmetteurVIR(Cpt, Inf);

  WriteLn(FF,St) ;
  Result:=True ;

END ;

Function TFEXPCFONB.EmetteurPRE ;
Var St,Dom,Etab,Guichet,NumCompte,Libelle : String ;
    Q  : TQuery ;
BEGIN
Result:=False ;
{RIB banque donneur d'ordre}

if FBoModeBatch then
  begin
  if Assigned( FTobScenario ) then
    begin
    Dom         := FTobScenario.GetString( 'BQ_DOMICILIATION') ;
    Etab        := FTobScenario.GetString( 'BQ_ETABBQ') ;
    Guichet     := FTobScenario.GetString( 'BQ_GUICHET') ;
    Numcompte   := FTobScenario.GetString( 'BQ_NUMEROCOMPTE') ;
    NumEmetteur := FTobScenario.GetString( 'CFONB_NUMEMETTEUR') ;
    Libelle     := FTobScenario.GetString( 'BQ_LIBELLE') ;
    end ;
  end
else
  begin
  Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+BQ_GENERAL.Value
            +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
  if Not Q.EOF then
     BEGIN
     Get_Compte(Q,Dom,Etab,Guichet,NumCompte) ;
     Libelle:= FullMajuscule(Q.FindField('BQ_LIBELLE').AsString);
     END else
     BEGIN
     Dom:='' ; Etab:='' ; Guichet:='' ; NumCompte:='' ; Libelle:='' ;
     END ;
  Ferme(Q) ;
  {Lecture Société Emetteur national}
  NumEmetteur:=GetParamSocSecur('SO_LIBRETEXTE0', '');
  {Contrôles}
  if Not EmetOK(NumEmetteur,Dom,Etab,Guichet,NumCompte) then Exit ;
  end ;

{Constituer l'enregistrement}
Inc(NumE) ;
{A}St:='03' ;
{B}St:=St+'08'+Format_String('',8)+Format_String(NumEmetteur,6) ;
{C}St:=St+Format_String('',7)+Date5(StrToDate(DateRemise.Text))+Format_String(V_PGI.NomSociete,24) ;
{D}If ExportEnEuro Then St:=St+Format_String('',7)+Format_String('',19)+'E'+Format_String('',5)+ZeroG(Guichet,5)+ZeroG(NumCompte,11)
                   Else St:=St+Format_String('',7)+Format_String('',19)+'F'+Format_String('',5)+ZeroG(Guichet,5)+ZeroG(NumCompte,11) ;
{E}St:=St+Format_String('',16) ;
{F}St:=St+Format_String('',31) ;
{G}St:=St+ZeroG(Etab,5)+Format_String('',6) ;
WriteLn(FF,St) ;
Result:=True ;
END ;

Function TFEXPCFONB.EmetteurCFONB ;
BEGIN
Result:=False ;
Case TypeExport.ItemIndex of
 0,1,2,3,4 : Result:=EmetteurLCR ;
 5,6,7 : Result:=EmetteurVIR ;
   8 : Result:=EmetteurPRE ;
   9 : Result:=EmetteurTRANS;
   END ;
END ;

{=============================================================================}
{===================================== LCR ===================================}
{=============================================================================}
Function TFEXPCFONB.DetailLCR ( O : TOB ) : boolean ;
Var Lib,St,AuxiGen,LibelleAuxi,Dom,Etab,Cle,Guichet,NumCompte,MP,CodeAcc : String ;
    Montant,D,C : Double ;
    DateC : TDateTime ;
    CharAcc : Char ;
    CodeCat : String ;
    CodeAccMvt : String ;
    lTobMP     : Tob ;
BEGIN
Result:=False ;
{Analyse OBM}
RecupGenOuAux(O,TRUE,AuxiGen,LibelleAuxi) ;
D:=O.GetValue('E_DEBIT');
if ((D<>0) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
C:=O.GetValue('E_CREDIT');
if ((C<>0) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
if RefTireLib.Checked then
  Lib := O.GetValue('E_LIBELLE')
else
  Lib := O.GetValue('E_REFEXTERNE') ;

Lib := FullMajuscule(Lib);

Montant:=D+C ; if Montant=0 then Exit ;
{Analyse MP}
MP:=O.GetValue('E_MODEPAIE') ; CodeAcc:='' ; CharAcc:='1' ;
lTobMP :=  GetTobMP( MP ) ;
if lTobMP <> nil then
 begin
 CodeAcc := lTobMP.GetString('MP_CODEACCEPT') ;
 CodeCat := lTobMP.GetString('MP_CATEGORIE') ;
 end ;
If CodeCat='LCR' Then
  BEGIN
  If CodeAcc='BOR' Then CharAcc:='2' else
    BEGIN
    CodeAccMvt:=O.GetValue('E_CODEACCEPT') ;
    // FQ 12336
    if CodeAccMvt='ACC' then CharAcc:='1' else
    if CodeAccMvt='NAC' then CharAcc:='0' else
    if CodeAccMvt='TRA' then CharAcc:='3' Else
    if CodeAccMvt='NON' then CharAcc:='1' ;
    END ;
  END ;
{Recup du RIB}
if Not GetLeRIB(AuxiGen,O.GetValue('E_RIB'),Etab,Guichet,NumCompte,Cle,Dom) then Exit ;
{Constituer l'enregistrement}
Inc(NumE) ;
{A}St:='06' ;
{B}St:=St+'60';
   St:=St+ZeroG(Inttostr(NumE),8);
   St:=St+Format_String('',6);
{C}St:=St+Format_String('',2);
   St:=St+Format_String(Lib,10);
   St:=St+Format_String(LibelleAuxi,24);
{D}St:=St+Format_String(Dom,24);
   St:=St+CharAcc;
   St:=St+Format_String('',2)+ZeroG(Etab,5)+ZeroG(Guichet,5)+ZeroG(NumCompte,11) ;
{E}St:=St+ZeroG(InttoStr(Round(100.0*Montant)),12);
   St:=St+Format_String('',4) ;
{F}St:=St+FormatDateTime('ddmmyy',O.GetValue('E_DATEECHEANCE')) ;
   DateC:=O.GetValue('E_DATECOMPTABLE') ; if DateC>StrToDate(DATEREMISE.Text) then DateC:=StrToDate(DATEREMISE.Text) ;
   St:=St+FormatDateTime('ddmmyy',DateC) ;
   St:=St+Format_String('',4);
   St:=St+Format_String('',1+3+3+9);
{G}St:=St+Format_String('',10) ;
TotalG:=Arrondi(TotalG+Montant,NbDec) ;
WriteLn(FF,St) ;
Result:=True ;
END ;

Function TFEXPCFONB.SommeDetailLCR ( X : Tot_Cli ) : boolean ;
Var St,AuxiGen,Lib,LibelleAuxi,Dom,Etab,Cle,Guichet,NumCompte,RefPiece : String ;
    Montant,D,C : Double ;
    CodeAcc : String ;
    CharAcc     : Char ;
BEGIN
Result:=False ;
{Analyse OBM}
AuxiGen:=X.AUXILIAIRE ; RefPiece:='' ;
D:=X.Debit ; C:=X.Credit ;
if ((D>C) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
if ((C>D) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
if ENC then Montant:=C-D else Montant:=D-C ;
if Montant=0 then Exit ;
{Recup du RIB}
if Not GetLeRIB(AuxiGen,X.RIB,Etab,Guichet,NumCompte,Cle,Dom) then Exit ;
{Infos Tiers}
LibelleAuxi := X.Libelle ;
{Code Acc}
// VL FQ 12336
CodeAcc:=X.CodeAccept; CharAcc:='1' ; Lib:=X.lastLib ;
if CodeAcc='BOR' then CharAcc:='2' else
if CodeAcc='ACC' then CharAcc:='1' else
if CodeAcc='NAC' then CharAcc:='0' else // FQ 12336
if CodeAcc='TRA' then CharAcc:='3' else
if CodeAcc='NON' then CharAcc:='1' ;
{Constituer l'enregistrement}
Inc(NumE) ;
{A}St:='06' ;
{B}St:=St+'60';
   St:=St+ZeroG(Inttostr(NumE),8);
   St:=St+Format_String('',6);
{C}St:=St+Format_String('',2);
   St:=St+Format_String(Lib,10);
   St:=St+Format_String(LibelleAuxi,24);
{D}St:=St+Format_String(Dom,24);
   St:=St+CharAcc;
   St:=St+Format_String('',2)+ZeroG(Etab,5)+ZeroG(Guichet,5)+ZeroG(NumCompte,11) ;
{E}St:=St+ZeroG(InttoStr(Round(100.0*Montant)),12);
   St:=St+Format_String('',4) ;
{F}St:=St+FormatDateTime('ddmmyy',X.LastEche);
   St:=St+FormatDateTime('ddmmyy',X.LastDate);
   St:=St+Format_String('',4);
   St:=St+Format_String('',1+3+3+9) ;
{G}St:=St+Format_String('',10) ;
TotalG:=Arrondi(TotalG+Montant,NbDec) ;
WriteLn(FF,St) ;
Result:=True ;
END ;

{=============================================================================}
{================================== VIREMENTS ================================}
{=============================================================================}
Function TFEXPCFONB.DetailVIR ( O : TOB ) : boolean ;
Var St,AuxiGen,LibelleAuxi,Dom,Cle,Etab,Guichet,NumCompte,RefPiece : String ;
    Montant,D,C : Double ;
  Cpt : TCompte; {JP24/09/03}
  Inf : TInfosCFONB;
BEGIN
Result:=False ;
{Analyse OBM}
RecupGenOuAux(O,TRUE,AuxiGen,LibelleAuxi) ;
RefPiece:= FullMajuscule(O.GetValue('E_LIBELLE'));
D:=O.GetValue('E_DEBIT');
if ((D<>0) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
C:=O.GetValue('E_CREDIT');
if ((C<>0) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
Montant:=D+C ; if Montant=0 then Exit ;
{Recup du RIB}
if Not GetLeRIB(AuxiGen,O.GetValue('E_RIB'),Etab,Guichet,NumCompte,Cle,Dom) then Exit ;
{Constituer l'enregistrement}
Inc(NumE) ;

  {JP 24/09/03 : Nouvelle fonction}
  case TypeExport.ItemIndex of
    6 : Inf.Code := '28';
    7 : Inf.Code := '27';
  else
    Inf.Code := '02';
  end;

  Inf.NumEmetteur   := NumEmetteur;

// FD 4234 : ref interne de la pièce ( SBO 24/07/2006 )
  Cpt.Divers        := Copy( O.GetString('E_REFINTERNE'), 1, 12) ;
// Fin FD 4234 : ref interne de la pièce
  Cpt.RaisonSoc     := LibelleAuxi;
  Cpt.Domiciliation := Dom;
  Cpt.CodeGuichet   := Guichet;
  Cpt.NumCompte     := NumCompte;
  Inf.Montant       := Montant;
  Inf.Motif         := Copy(RefPiece,1,29);
  Cpt.CodeBanque    := Etab;

  st := TGenerationCFONB.DetailVIR(Cpt, Inf);

  TotalG:=Arrondi(TotalG+Montant,NbDec) ;
  WriteLn(FF,St) ;

  Result:=True ;
END ;

Function TFEXPCFONB.SommeDetailVIR ( X : Tot_Cli ) : boolean ;
Var St,AuxiGen,LibelleAuxi,Dom,Etab,Cle,Guichet,NumCompte,RefPiece : String ;
    Montant,D,C : Double ;
  Cpt : TCompte; {JP24/09/03}
  Inf : TInfosCFONB;
BEGIN
Result:=False ;
{Analyse OBM}
AuxiGen:=X.AUXILIAIRE ; RefPiece:='' ;
D:=X.Debit ; C:=X.Credit ;
if ((D>C) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
if ((C>D) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
if ENC then Montant:=C-D else Montant:=D-C ;
if Montant=0 then Exit ;
{Recup du RIB}
if Not GetLeRIB(AuxiGen,X.RIB,Etab,Guichet,NumCompte,Cle,Dom) then Exit ;
{Infos Tiers}
LibelleAuxi := X.Libelle ;
{Constituer l'enregistrement}
Inc(NumE) ;

  {JP 24/09/03 : Nouvelle fonction}
  case TypeExport.ItemIndex of
    6 : Inf.Code := '28';
    7 : Inf.Code := '27';
  else
    Inf.Code := '02';
  end;

  Inf.NumEmetteur   := NumEmetteur;

// FD 4234 : ref interne de la pièce ( SBO 24/07/2006 )
  Cpt.Divers        := Copy( X.Auxiliaire, 1, 12) ;
// Fin FD 4234 : ref interne de la pièce
  Cpt.RaisonSoc     := LibelleAuxi;
  Cpt.Domiciliation := Dom;
  Cpt.CodeGuichet   := Guichet;
  Cpt.NumCompte     := NumCompte;
  Inf.Montant       := Montant;
  Inf.Motif         := Copy(RefPiece,1,29);
  Cpt.CodeBanque    := Etab;

  st := TGenerationCFONB.DetailVIR(Cpt, Inf);

  TotalG:=Arrondi(TotalG+Montant,NbDec) ;
  WriteLn(FF,St) ;

  Result:=True ;
END ;

{=============================================================================}
{================================= PRELEVEMENTS ==============================}
{=============================================================================}
Function TFEXPCFONB.DetailPRE ( O : TOB ) : boolean ;
Var St,AuxiGen,LibelleAuxi,Dom,Etab,Cle,Guichet,NumCompte,RefPiece : String ;
    D,C,Montant : Double ;
    lStRef  : string ;
BEGIN
Result:=False ;
{Analyse OBM}
RecupGenOuAux(O,TRUE,AuxiGen,LibelleAuxi) ;
RefPiece:=FullMajuscule(O.GetValue('E_LIBELLE'));
// Oubli ???
D:=O.GetValue('E_DEBIT');
if ((D<>0) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
C:=O.GetValue('E_CREDIT');
if ((C<>0) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
Montant:=D+C ;
if Montant=0 then Exit ;
// Fin Oubli ???
//Montant:=O.GetValue('E_DEBIT')+O.GetValue('E_CREDIT');
{Recup du RIB}
if Not GetLeRIB(AuxiGen,O.GetValue('E_RIB'),Etab,Guichet,NumCompte,Cle,Dom) then Exit ;
{Constituer l'enregistrement}
Inc(NumE) ;
{A}St:='06' ;                                          // Zone 1 à 2     : Type d'enregistrement :  06 = Destinataire
{B}St:=St+'08';                                        // Zone 3 à 4     : Code opération
   St:=St+Format_String('',8);                         // Zone 5 à 12    : Réservée
   St:=St+Format_String(NumEmetteur,6) ;               // Zone 13 à 18   : N° émetteur
// FD 4234 : ref interne de la pièce ( SBO 24/07/2006 )
   lStRef := Copy( O.GetString('E_REFINTERNE'), 1, 12) ;
{C}St:=St+Format_String( lStRef, 12 );                 // Zone 19 à 30   : Référence interne du donneur d'ordre
// Fin FD 4234 : ref interne de la pièce
   St:=St+Format_String(LibelleAuxi,24) ;              // Zone 31 à 54   : Nom/Raison social du bénéficiaire
{D}St:=St+Format_String(Dom,24);                       // Zone 55 à 78   : Domiciliation /
// Si salaires,pensions,prest. assimilées en faveur non-résidents : 55 à 74 : Domiciliation + 75 à 78 : Nature et identification géographique
// Si déclaration à la balance des paiements des virements en faveur non-résidents : 55 à 63 : N° SIREN + 64 à 78 : Domiciliation
   St:=St+Format_String('',8);                         // Zone 79 à 86   : Déclaration à la balnce des paiements (suite)
   St:=St+ZeroG(Guichet,5);                            // Zone 87 à 91   : Code guichet banque bénéficiaire
   St:=St+ZeroG(NumCompte,11) ;                        // Zone 92 à 102  : N° compte
{E}St:=St+ZeroG(InttoStr(Round(100.0*Montant)),16) ;   // Zone 103 à 118 : Montant virement
{F}St:=St+Format_String(Copy(RefPiece,1,31),31);       // Zone 119 à 149 : Libellé virement
{G}St:=St+ZeroG(Etab,5);                               // Zone 150 à 154 : Code établissement bénéficiaire
   St:=St+Format_String('',6) ;                        // Zone 155 à 160 : Réservée
TotalG:=Arrondi(TotalG+Montant,NbDec) ;
WriteLn(FF,St) ;
Result:=True ;
END ;

Function TFEXPCFONB.SommeDetailPRE ( X : Tot_Cli ) : boolean ;
Var St,AuxiGen,LibelleAuxi,Dom,Etab,Cle,Guichet,NumCompte,RefPiece : String ;
    Montant,D,C : Double ;
    lStRef : string ;
BEGIN
Result:=False ;
{Analyse OBM}
AuxiGen:=X.AUXILIAIRE ; RefPiece:='' ;
D:=X.Debit ; C:=X.Credit ;
if ((D>C) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
if ((C>D) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
if ENC then Montant:=C-D else Montant:=D-C ;
if Montant=0 then Exit ;
{Recup du RIB}
if Not GetLeRIB(AuxiGen,X.RIB,Etab,Guichet,NumCompte,Cle,Dom) then Exit ;
{Infos Tiers}
LibelleAuxi := X.Libelle ;

{Constituer l'enregistrement}
Inc(NumE) ;
{A}St:='06' ;
{B}St:=St+'08'+Format_String('',8)+Format_String(NumEmetteur,6) ;

// FD 4234 : code auxiliaire ( SBO 24/07/2006 )
   lStRef := Copy( X.Auxiliaire, 1, 12) ;
{C}St:=St+Format_String( lStRef, 12 );                        // Zone 19 à 30   : Référence interne du donneur d'ordre
// Fin FD 4234 : ref interne de la pièce

//{C}St:=St+Format_String('',12); // SBO 26/10/2006 : bug ajout en trop de ces caractères

   St:=St+Format_String(LibelleAuxi,24) ;
{D}St:=St+Format_String(Dom,24);
   St:=St+Format_String('',8)+ZeroG(Guichet,5)+ZeroG(NumCompte,11) ;
{E}St:=St+ZeroG(InttoStr(Round(100.0*Montant)),16) ;
{F}St:=St+Format_String(Copy(RefPiece,1,31),31);
{G}St:=St+ZeroG(Etab,5);
   St:=St+Format_String('',6) ;

   TotalG:=Arrondi(TotalG+Montant,NbDec) ;
   WriteLn(FF,St) ;

   Result:=True ;
END ;

Function TFEXPCFONB.DetailCFONB (O : TOB ) : boolean ;
BEGIN
Result:=False ;
Case TypeExport.ItemIndex of
 0,1,2,3,4 : Result:=DetailLCR(O) ;
 5,6,7 : Result:=DetailVIR(O) ;
   8 : Result:=DetailPRE(O) ;
   9 : begin
         Result := DetailTRANS(O);
         Result := Result and DetailTRANS2(O);
       end;
   END ;
if FBoModeBatch then
  if not result then
    begin
    O.PutValue('E_CFONBOK', '-');
    O.PutValue('E_NUMCFONB', '');
    end ;
END ;

Function TFEXPCFONB.SommeDetailCFONB ( X : Tot_Cli ) : boolean ;
BEGIN
Result:=False ;
Case TypeExport.ItemIndex of
 0,1,2,3,4 : Result:=SommeDetailLCR(X) ;
 5,6,7 : Result:=SommeDetailVIR(X) ;
   8 : Result:=SommeDetailPRE(X) ;
   9 : begin
         Result:=SommeDetailTRANS(X) ;
         Result:= Result and SommeDetailTRANS2(X);
       end;
   END ;
X.exporte := result ;
END ;

{======================================= Total ====================================}
Procedure TFEXPCFONB.TotalLCR ;
Var St : String ;
BEGIN
Inc(NumE) ;
{A}St:='08' ;
{B}St:=St+'60';
   St:=St+ZeroG(Inttostr(NumE),8);
   St:=St+Format_String('',6);
{C}St:=St+Format_String('',12+24);
{D}St:=St+Format_String('',24+8+5+11);
{E}St:=St+ZeroG(InttoStr(Round(100.0*TotalG)),12);
{F}St:=St+Format_String('',4+31);
{G}St:=St+Format_String('',5+6);
WriteLn(FF,St) ;
END ;

Procedure TFEXPCFONB.TotalVIR ;
Var
  St : String ;
  Inf : TInfosCFONB;
BEGIN
Inc(NumE) ;
  {JP 24/09/03 : Nouvelle fonction}
  case TypeExport.ItemIndex of
    6 : Inf.Code := '28';
    7 : Inf.Code := '27';
  else
    Inf.Code := '02';
  end;

  Inf.NumEmetteur := NumEmetteur;
  Inf.Montant     := TotalG;
  st := TGenerationCFONB.TotalVIR(Inf);

  WriteLn(FF,St) ;
END ;

Procedure TFEXPCFONB.TotalPRE ;
Var St : String ;
BEGIN
Inc(NumE) ;
{A}St:='08' ;
{B}St:=St+'08';
   St:=St+Format_String('',8);
   St:=St+Format_String(NumEmetteur,6) ;
{C}St:=St+Format_String('',12+24) ;
{D}St:=St+Format_String('',24+8+5+11) ;
{E}St:=St+ZeroG(InttoStr(Round(100.0*TotalG)),16) ;
{F}St:=St+Format_String('',31) ;
{G}St:=St+Format_String('',5+6) ;
WriteLn(FF,St) ;
END ;

Procedure TFEXPCFONB.TotalCFONB ;
BEGIN
Case TypeExport.ItemIndex of
 0,1,2,3,4 : TotalLCR ;
 5,6,7 : TotalVIR ;
   8 : TotalPRE ;
   9 : TotalTRANS ;
   END ;
END ;

Function TFEXPCFONB.LanceCFONB ( App : boolean ) : boolean ;
Var {io,}i : integer ;
    O    : TOB ;
    X    : Tot_Cli ;
    Okok,UneLigne : boolean ;
    StFichier : String ;
BEGIN
Result:=False ; Okok:=True ; UneLigne:=False ;
{Gestion fichier}
StFichier:=NomFichier.Text ; if StFichier='' then Exit ;
AssignFile(FF,StFichier) ;
{$I-}
if App then Append(FF) else ReWrite(FF) ;
{$I+}
if IoResult<>0 then BEGIN {$I-} CloseFile(FF) ; {$I+} {io:=ioResult ;} Exit ; END ;
{Traitement export}
NumE:=0 ;
// Mode batch  : génération du total sur le dernier lot
if FBoModeBatch and Assigned( FTobScenario )then
  begin
  if FTobScenario.GetValue('CFONB_PREMIERLOT') <> 'X'
    then TotalG := FTobScenario.GetValue('CFONB_TOTALG')
    else TotalG := 0 ;
  end
else TotalG:=0 ;
gszDateEcheance := Date5(StrToDate(TOB(TECHE[0]).GetValue('E_DATEECHEANCE')));
{JP 24/09/03}
gszDateEcheanceBis := StrToDate(TOB(TECHE[0]).GetValue('E_DATEECHEANCE'));

// En mode batch, ecriture de l'émetteur CFONB uniquement sur le 1er passage
if FBoModeBatch and assigned( FTobScenario ) then
  begin
  if ( FTobScenario.GetValue('CFONB_PREMIERLOT') = 'X' )
    then EmetteurCFONB
    else NumEmetteur := FTobScenario.GetString( 'CFONB_NUMEMETTEUR') ;
  end
else if Not EmetteurCFONB then
     begin
     CloseFile(FF) ;
     Exit ;
     end ;

if Not CCumulTiers.Checked then
   BEGIN
   InitMove(TECHE.Count,'') ;
   for i:=0 to TECHE.Count-1 do
       BEGIN
       MoveCur(False) ;
       O:=TOB(TECHE[i]) ;
       if Not DetailCFONB(O) then Okok:=False else UneLigne:=True ;
       END ;
END else
   BEGIN
   InitMove(TCLI.Count,'') ;
   for i:=0 to TCLI.Count-1 do
       BEGIN
       MoveCur(False) ; X:=Tot_Cli(TCLI[i]) ;
       if Not SommeDetailCFONB(X) then Okok:=False else UneLigne:=True ;
       END ;
END ;
FiniMove ;
// Mode batch  : génération du total sur le dernier lot
if FBoModeBatch then
  begin
  if Assigned( FTobScenario ) then
    if (FTobScenario.GetValue('CFONB_DERNIERLOT') = 'X')
      then TotalCFONB
      else FTobScenario.PutValue('CFONB_TOTALG', TotalG) ;
  end
// Mode classique
else TotalCFONB ;

CloseFile(FF) ;
if Not UneLigne then
   BEGIN
   AssignFile(FF,StFichier) ;
   {$i-} Erase(FF) ; {$i+} {io:=ioResult ;}
   Okok:=False ;
   if FBoModeBatch then // Mode Batch : REport des erreurs dans la tob de transition
     FTobScenario.PutValue('CFONB_ERREUR', 16)
     else HM.Execute(16,'','') ;
   END else
   BEGIN
   if ((Not Okok) or (ExisteMauvais)) then
     BEGIN
     if FBoModeBatch then // Mode Batch : REport des erreurs dans la tob de transition
       begin
       FTobScenario.PutValue('CFONB_EXISTEMAUVAIS', 'X') ;
       Okok := True ;
       end
     else if HM.Execute(9,'','')=mrYes
       then Okok:=True ;
     END ;
   END ;
Result:=Okok ;
END ;

{=========================== Edition document ===================================}
procedure TFEXPCFONB.FlagCFONB ;
Var i,Nb       : integer ;
    O          : TOB ;
    SWhere     : String ;
    RIB        : String ;
    SQL        : String ;
    NumCFONB   : String ;
    NowFutur   : TDateTime ;
    lStDossier : string ;
begin
  if TECHE.Count<=0 then Exit ;

  NowFutur := NowH ;
  NumCFONB := GetNumCFONB(FormatCFONB) ;

  for i:=0 to TECHE.Count-1 do
    begin
    O      := TOB(TECHE[i]) ;
    SWhere := whereEcritureTob( tsGene, O, True ) ;
    RIB    := FullMajuscule(O.GetValue('E_RIB'));

    // Gestion multisoc
    if EstMultiSoc and ( O.GetString('E_SOCIETE') <> GetParamSocSecur('SO_SOCIETE', '') ) then
      begin
      lStDossier := GetSchemaName( O.GetString('E_SOCIETE') ) ;
      if lStDossier = '' then Continue ;
      end
    else lStDossier := V_PGI.SchemaName ;

    // MAJ BASE
    if Trim(RIB)<>'' then
      begin
      {JP 16/11/07 : FQ 21847 : Gestion de E_UTILISATEUR}
      SQL := 'UPDATE ' + GetTableDossier( lStDossier, 'ECRITURE' )
             + ' SET E_CFONBOK="X", E_NUMCFONB="' + NumCFONB + '", E_RIB="' + RIB + '", E_DATEMODIF="' 
             + UsTime(NowFutur) + '", E_UTILISATEUR = "' + V_PGI.User + '" Where ' + SWhere ;
      Nb  := ExecuteSQL(SQL) ;
      if Nb<>1 then V_PGI.IoError:=oeSaisie ;
      end ;

    if V_PGI.IoError<>oeOk then Break ;
    end ;

end ;

procedure TFEXPCFONB.EditeCFONB ;
Var SWhere : String ;
BEGIN
if Document.Value='' then Exit ;
if TECHE.Count<=0 then Exit ;
{$IFDEF CCMP}
if FBoMultiSoc then
  begin
  ExecuteEmissionBOR  ( TEche, Document.Value, True, FApercu.Checked, 'REF_REMISE=' + txt_RefRemise.Text + '`') ;
  end
else
{$ENDIF CCMP}
  begin
  SWhere := WhereMulti(TECHE) ;
  {JP 07/06/07 : FQ 18831 : ajout de la référence de remise dans les critères
                 FQ 19334 : ajout de la Date de remise dans les critères}
  LanceEtat('E','BOR',Document.Value,FApercu.Checked,False,False,Nil,SWhere,'',False,
            0, 'REF_REMISE=' + txt_RefRemise.Text + '`DATE_REMISE=' + DATEREMISE.Text + '`') ;
  end ;
END ;

procedure TFEXPCFONB.BQ_GENERALChange(Sender: TObject);
begin
if FBoModeBatch then Exit ;
BRIB.Enabled:=(BQ_GENERAL.Value<>'') ;
if GeneCharge then Exit ;
CodeBanque:=BQ_GENERAL.Value ;
GetNomFichier ;
{JP 05/07/05 : FQ 15493 : Génère un nom de fichier unique à partir de GetNomFichier}
GenereNomFichier(NomFichier.Text);
end;

procedure TFEXPCFONB.BFermeClick(Sender: TObject);
begin
Close ;
end;

procedure TFEXPCFONB.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFEXPCFONB.BCompteRenduClick(Sender: TObject);
begin
VisuExportCFONB(LastExport) ;
end;

procedure TFEXPCFONB.BRIBClick(Sender: TObject);
Var St : String ;
begin
St:=BQ_GENERAL.Value ;
if St<>'' then FicheBanqueCP(St,taModif,0) ;
end;

procedure TFEXPCFONB.BParamClick(Sender: TObject);
begin
  ParamSociete(False,'','SCO_SUIVITIERS','',ChargeSocieteHalley,ChargePageSoc,SauvePageSoc,InterfaceSoc,1105000);
end;

procedure TFEXPCFONB.FormCreate(Sender: TObject);
begin
Pages.ActivePageIndex := 0;
TCli:=TList.Create ;
TECHES:=TList.Create ; TECHE1:=TList.Create ;
NbDec:=V_PGI.OkDecV ; OkDebit:=True ;
FBoModeBatch := False ;
//SG6 18.04.05 FQ 15372
// CA - 16/10/2006 : Transferts internationaux disponibles uniquement dans CCMP
{$IFNDEF CCMP}
TypeExport.Items.Delete(9);
{$ENDIF}
FTobMP := Tob.Create('$MODEPAIE', nil, -1 ) ;
end;


procedure TFEXPCFONB.DocumentChange(Sender: TObject);
begin
if FBoModeBatch then Exit ;
FApercu.Enabled:=Document.Value<>'' ;
end;


function TFEXPCFONB.EmetteurTRANS: boolean;
var
  St,szCodeBIC,szCodeIBAN,szCodeDEV,szRef : String ;
  dtDateExec : TDateTime ;
  Q  : TQuery ;
  Cpt : TTransfert;
  Inf : TInfosCFONB;
  Soc : TInfosSoc;
begin
  Result:=False ;
  // Date d'exécution souhaitée
  dtDateExec:=StrToDate(DATEREMISE.Text) ;

if FBoModeBatch then
  begin
  if Assigned( FTobScenario ) then
    begin
    szCodeBIC   := FTobScenario.GetString('BQ_CODEBIC') ;
    szCodeIBAN  := FTobScenario.GetString('BQ_CODEIBAN') ;
    szCodeDEV   := FTobScenario.GetString('BQ_DEVISE') ;
    end ;
  end
else
  begin
  // Référence de la remise
  if (txt_RefRemise.Text='') then begin HM.Execute(32,caption,''); if txt_RefRemise.CanFocus then txt_RefRemise.SetFocus; Exit; end;
  szRef:=txt_RefRemise.Text;
  // Code BIC de la banque émettrice
  Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+BQ_GENERAL.Value
            +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True); // 19/10/2006 YMO Multisociétés
  if Not Q.EOF then begin
    szCodeBIC:=Q.FindField('BQ_CODEBIC').AsString;
    szCodeIBAN:=Q.FindField('BQ_CODEIBAN').AsString;
    szCodeDEV:=Q.FindField('BQ_DEVISE').AsString;
  end;
  Ferme(Q);
  if szCodeBIC = '' then begin HM.Execute(23,caption,''); Exit; end;
  if szCodeIBAN = '' then begin HM.Execute(24,caption,''); Exit; end;
  if szCodeDEV = '' then begin HM.Execute(25,caption,''); Exit; end;
  end ;

  // Constitue l'enregistrement
  // Toutes les dates ont la forme AAAAMMJJ
  NumE := 0;
  Inc(NumE);

  {JP 24/09/03 : Nouvelle fonction}
  Inf.DateCreat := Date;
  Soc.Raison    := FullMajuscule(GetParamSocSecur('SO_LIBELLE', ''));
  Soc.Adr1      := FullMajuscule(GetParamSocSecur('SO_ADRESSE1', ''));
  Soc.Ville     := FullMajuscule(GetParamSocSecur('SO_CODEPOSTAL', '')) + ' ' +
                   FullMajuscule(GetParamSocSecur('SO_VILLE', '')); // 13466

  Soc.Pays      := FullMajuscule(GetParamSocSecur('SO_PAYS', ''));
  Soc.Siret     := StringReplace(GetParamSocSecur('SO_SIRET', ''), ' ', '', [rfReplaceAll]);
  Cpt.Remise    := szRef;
  Cpt.CodeBIC   := szCodeBIC;
  Cpt.CodeIBAN  := szCodeIBAN;
  Cpt.Devise    := szCodeDEV;
  Cpt.IdClient  := GetParamSocSecur('SO_CPIDCLIENT', 0); // 13466
  Cpt.TypeRemise := IntToStr(TypeDebit.ItemIndex +1); // 13466
  Cpt.DateExect := dtDateExec;
  // GCO - 19/07/2006 - FQ 18603
  Cpt.TypeDevise := FStDevise; // '' si multidevise, la devise si monodevise

  st := TGenerationCFONB.EmetteurTrans(Cpt, Inf, Soc);

  WriteLn(FF,St);
  Result:=True;
end;

// Détail de l'opération
function TFEXPCFONB.DetailTRANS(O: TOB): boolean;
var
  St,szRef,szCodeIBAN,szNom,szAdresse1,szVille,szPays,szCodePays,szDevise,szCodeIBAN2,szDevise2,szNatEco : String ;
  Q,Q2  : TQuery ;
  dtDateExec : TDateTime ;
  Montant,D,C : Double ;
  Cpt : TTransfert; {JP 24/09/03}
  Inf : TInfosCFONB;
  TIDTIC : Boolean ;
  LibAG, CptAG : String ;
begin
  Result:=False ;

  szRef := FullMajuscule(O.GetValue('E_REFINTERNE'));

  // Code IBAN
  szCodeIBAN := FullMajuscule(O.GetValue('E_RIB'));

  {YMO 26/01/2005 Prise en compte des cas TIC et TID FQ17308
   JP 23/05/06 : FQ 18014 : Il faut mettre True pour récupérer le libellé}
  TIDTIC := RecupGenOuAux(O, {FALSE} True, CptAG, LibAG) ;
  if TIDTIC
    Then Q:=OpenSQL('SELECT G_ADRESSE1,G_VILLE,G_PAYS FROM GENERAUX WHERE G_GENERAL="'+CptAG+'"',True)
    else Q:=OpenSQL('SELECT T_ADRESSE1,T_VILLE,T_PAYS FROM TIERS WHERE T_AUXILIAIRE="'+CptAG+'"',True);
  if Not Q.EOF then begin
    szNom:=LibAG;
    if Copy(szCodeIBAN,1,1)<>'*' then begin
      if HM.Execute(26,'',IntToStr(O.GetValue('E_NUMEROPIECE'))+'-'+szNom+' n''est pas renseigné. Souhaitez-vous utiliser le RIB principal du compte ?')<>mrYes then begin Ferme(Q); Exit ; end;
      Q2:=OpenSQL('SELECT R_CODEIBAN,R_NATECO from RIB Where R_AUXILIAIRE="'+CptAG+'" AND R_PRINCIPAL="X"',True) ;
      if Not Q2.EOF then begin
        szCodeIBAN:=Q2.FindField('R_CODEIBAN').AsString; if szCodeIBAN='' then
        begin
          HM.Execute(28,'',''); Ferme(Q); Ferme(Q2); Exit;
        end;
        if (cbo_NatEco.Value = '') then begin
          szNatEco:=Q2.FindField('R_NATECO').AsString;
          if (szNatEco = '') then begin HM.Execute(30,'',CptAG+' n''a pas de nature économique renseignée.'); Ferme(Q); Ferme(Q2); Exit; end;
        end;
        end
      else
      begin
       HM.Execute(27,'',''); Ferme(Q); Ferme(Q2); Exit ;
      end;
      Ferme(Q2);
      end
    else
      Delete(szCodeIBAN,1,1);
    szAdresse1:= FullMajuscule(Q.Fields[0].AsString);
    szVille:= FullMajuscule(Q.Fields[1].AsString);
    szPays:= FullMajuscule(Q.Fields[2].AsString);
  end;
  Ferme(Q);

  if szCodeIBAN = '' then begin HM.Execute(28,'',''); Exit; end;

  if szPays = '' then begin HM.Execute(31,'',szNom+' n''a pas de pays renseigné.'); Exit; end;

  // Nature économique
  if (cbo_NatEco.Value <> '') then szNatEco := cbo_NatEco.Value;
  if (cbo_NatEco.Value = '') and (szNatEco='') then begin
    Q2:=OpenSQL('SELECT R_NATECO from RIB Where R_AUXILIAIRE="'+O.GetValue('E_AUXILIAIRE')+'" AND R_CODEIBAN="'+szCodeIBAN+'" AND R_NATECO<>""',True) ;
      if Not Q2.EOF then szNatEco:=Q2.FindField('R_NATECO').AsString
                    else begin HM.Execute(30,'',O.GetValue('E_AUXILIAIRE')+' n''a pas de nature économique renseignée.'); Ferme(Q2); Exit ; end;
    Ferme(Q2);
  end;

  // Code IBAN et devise de l'émetteur
  Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+BQ_GENERAL.Value
            +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
  if Not Q.EOF then begin
    szCodeIBAN2 := Q.FindField('BQ_CODEIBAN').AsString;
    szDevise2   := Q.FindField('BQ_DEVISE').AsString;
    szCodePays  := Q.FindField('BQ_PAYS').AsString; // FQ 12430
    if szCodePays = '' then szCodePays := 'FRA';  // France par défaut // FQ 12430

  end;
  Ferme(Q) ;

  {JP 25/09/03 : Il faut reprendre le libellé de l'auxiliaire et non celui de la banque
                qui est renseigné dans l'enregistrement 05}
  Q:=OpenSQL('SELECT PY_CODEISO2, PY_LIBELLE FROM PAYS WHERE PY_PAYS="'+szPays+'"',True) ;
  if not Q.EOF then begin
    szPays       := FullMajuscule(Q.FindField('PY_LIBELLE').AsString);
    Cpt.CodePays := Q.FindField('PY_CODEISO2').AsString;;
  end;
  Ferme(Q);
  if szCodePays = '' then begin HM.Execute(29,'',szPays+' n''est pas renseigné.'); Exit; end;

  //SG6 21.02.05 FQ 15374
  if (REnvoi.ItemIndex=1) or (TypeExport.ItemIndex = 9) then begin
    // Devises
    szDevise := O.GetValue('E_DEVISE');

    // Montant de la remise
    D:=O.GetValue('E_DEBITDEV');
    if ((D<>0) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
    C:=O.GetValue('E_CREDITDEV');
    if ((C<>0) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
    Montant:=D+C ; if Montant=0 then Exit ;
    end
  else begin
    // Devises
    szDevise := 'EUR';

    // Montant de la remise
    D:=O.GetValue('E_DEBIT');
    if ((D<>0) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
    C:=O.GetValue('E_CREDIT');
    if ((C<>0) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
    Montant:=D+C ; if Montant=0 then Exit ;
  end;

  // Date d'éxécution
  dtDateExec:=StrToDate(DATEREMISE.Text) ;

  // Constitue l'enregistrement
  // Toutes les dates ont la forme AAAAMMJJ
  Inc(NumE);

  {JP 24/09/03 : Nouvelle fonction}
  Inf.NumEmetteur := Inttostr(NumE);
  Cpt.CodeIBAN    := szCodeIBAN;
  Cpt.Nom         := szNom;
  Cpt.Adresse1    := szAdresse1;
  Cpt.Ville       := szVille;
  Cpt.Pays        := szPays;
  Cpt.Remise      := szRef;
  Inf.Montant     := Montant;
  Cpt.NatEco      := szNatEco;
  Inf.Imputation  := IntToStr(ImputFrais.ItemIndex + 13);
  Cpt.Devise      := szDevise;
  Cpt.CodeIBAN2   := szCodeIBAN2;
  Cpt.Devise2     := szDevise2;
  cpt.DateExect   := dtDateExec;
  St := TGenerationCFONB.DetailTrans(Cpt, Inf);
  TotalG:=Arrondi(TotalG+Montant,NbDec) ;
  WriteLn(FF,St) ;
  Result:=True ;
end;

// Banque du bénéficiaire
function TFEXPCFONB.DetailTRANS2(O: TOB): boolean;
var
  St, szNomBQ, szVille, szCodeBIC, szPays, szPays2 : String ;
  Q  : TQuery ;
  Cpt : TTransfert;
  Inf : TInfosCFONB;
  LibAG, CptAG : String ;
begin
  Result:=False ;

  //YMO 31/01/2005 Prise en compte des cas TIC et TID FQ17308
  RecupGenOuAux(O,FALSE,CptAG,LibAG) ;

  {JP 05/07/05 : FQ 16161 : Amélioration des de la reprise des compte}
  szPays2 := GetInfosRibCompte(CptAG, O.GetValue('E_RIB'), szVille, szPays, szCodeBIC, szNomBQ);

  Q := OpenSQL('SELECT PY_LIBELLE FROM PAYS WHERE PY_PAYS="' + szPays + '"',True);
  if not Q.EOF then szPays2 := FullMajuscule(Q.FindField('PY_LIBELLE').AsString)
               else szPays2 := '';
  Ferme(Q);

  if (szCodeBIC = '') then
  begin
    HM.Execute(33,'',CptAG+' n''est pas renseigné.');
    Exit ;
  end;

  if (szPays = '') then
  begin
    HM.Execute(34,'',CptAG+' n''est pas renseigné.');
    Exit ;
  end;

  // Constitue l'enregistrement
  Inc(NumE);

  {JP 24/09/03 : Nouvelle fonction}
  Inf.NumEmetteur := Inttostr(NumE);
  Cpt.Banque      := szNomBQ;
  Cpt.Ville       := szVille;
  Cpt.CodeBic     := szCodeBic;
  Cpt.Pays        := szPays;
  Cpt.Pays2       := szPays2;
  st := TGenerationCFONB.DetailTrans2(Cpt, Inf);

  WriteLn(FF,St) ;
  Result:=True ;
end;

procedure TFEXPCFONB.TotalTRANS;
var
  St,szCodeIBAN,szCodeDEV,szRef : String ;
  dtDateCreat : TDateTime ;
  Q : TQuery;
  Cpt : TTransfert;
  Inf : TInfosCFONB;
  Soc : TInfosSoc;
begin
  // Date de création
  dtDateCreat:=StrToDate(DATEREMISE.Text) ;

  // Référence de la remise
  if (txt_RefRemise.Text='') then begin HM.Execute(32,caption,''); if txt_RefRemise.CanFocus then txt_RefRemise.SetFocus; Exit; end;
  szRef:=txt_RefRemise.Text;

  // Code IBAN de la banque émettrice
  Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+BQ_GENERAL.Value
            +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True); // 19/10/2006 YMO Multisociétés
  if Not Q.EOF then begin
    szCodeIBAN:=Q.FindField('BQ_CODEIBAN').AsString;
    szCodeDEV:=Q.FindField('BQ_DEVISE').AsString;
  end;
  Ferme(Q);
  Inc(NumE) ;

  {JP 24/09/03 : Nouvelle fonction}
  Inf.NumEmetteur := Inttostr(NumE);
  Inf.DateCreat   := dtDateCreat;
  Soc.Siret       := StringReplace(GetParamSocSecur('SO_SIRET', ''), ' ', '', [rfReplaceAll]);
  Cpt.Remise      := szRef;
  Cpt.CodeIBAN    := szCodeIBAN;
  Cpt.Devise      := szCodeDEV;
  Cpt.IdClient    := GetParamSocSecur('SO_CPIDCLIENT', 0); // 13466
  Inf.Montant     := TotalG;

  st := TGenerationCFONB.TotalTrans(Cpt, Inf, Soc);
  WriteLn(FF,St) ;
end;

function TFEXPCFONB.SommeDetailTRANS(X: Tot_Cli): boolean;
var
  St,szRef,szCodeIBAN,szNom,szAdresse1,szVille,szPays,szCodePays,szDevise,szCodeIBAN2,szDevise2,szNatEco : String ;
  Q,Q2  : TQuery ;
  dtDateExec : TDateTime ;
  Montant,D,C : Double ;
  Cpt : TTransfert; {JP 24/09/03}
  Inf : TInfosCFONB;
begin
  Result:=False ;

  szRef := X.szRef;

  // Code IBAN
  szCodeIBAN:=X.RIB;
  Q:=OpenSQL('SELECT T_LIBELLE,T_ADRESSE1,T_VILLE,T_PAYS FROM TIERS WHERE T_AUXILIAIRE="'+X.Auxiliaire+'"',True) ;

  //YMO 31/01/2005 Prise en compte des cas TIC et TID FQ17308

  if X.TIDTIC Then
      Q:=OpenSQL('SELECT G_LIBELLE,G_ADRESSE1,G_VILLE,G_PAYS FROM GENERAUX WHERE G_GENERAL="'+X.Auxiliaire+'"',True)
  else
      Q:=OpenSQL('SELECT T_LIBELLE,T_ADRESSE1,T_VILLE,T_PAYS FROM TIERS WHERE T_AUXILIAIRE="'+X.Auxiliaire+'"',True) ;
  if Not Q.EOF then begin
    szNom:=Q.Fields[0].AsString ;
    if Copy(szCodeIBAN,1,1)<>'*' then begin
      if HM.Execute(26,'',szNom+' n''est pas renseigné. Souhaitez-vous utiliser le RIB principal du compte ?')<>mrYes then begin Ferme(Q); Exit ; end;
      Q2:=OpenSQL('SELECT R_CODEIBAN,R_NATECO from RIB Where R_AUXILIAIRE="'+X.Auxiliaire+'" AND R_PRINCIPAL="X"',True) ;
      if Not Q2.EOF then begin
        szCodeIBAN:=Q2.FindField('R_CODEIBAN').AsString; if szCodeIBAN='' then begin HM.Execute(28,'',''); Ferme(Q); Ferme(Q2); Exit; end;
        if (cbo_NatEco.Value = '') then begin
          szNatEco:=Q2.FindField('R_NATECO').AsString;
          if (szNatEco = '') then begin HM.Execute(30,'',X.AUXILIAIRE+' n''a pas de nature économique renseignée.'); Ferme(Q); Ferme(Q2); Exit; end;
        end;
        end
      else begin HM.Execute(27,'',''); Ferme(Q); Ferme(Q2); Exit ; end;
      Ferme(Q2);
      end
    else
      Delete(szCodeIBAN,1,1);
    szAdresse1:=Q.Fields[1].AsString ;
    szVille:=Q.Fields[2].AsString ;
    szPays:=Q.Fields[3].AsString ;
  end;
  Ferme(Q);
  if szCodeIBAN = '' then begin HM.Execute(28,'',''); Exit; end;
  if szPays = '' then begin HM.Execute(31,'',szNom+' n''a pas de pays renseigné.'); Exit; end;

  // Nature économique
  if (cbo_NatEco.Value <> '') then szNatEco := cbo_NatEco.Value;
  if (cbo_NatEco.Value = '') and (szNatEco='') then begin
    Q2:=OpenSQL('SELECT R_NATECO FROM RIB WHERE R_AUXILIAIRE="'+X.AUXILIAIRE+'" AND R_CODEIBAN="'+szCodeIBAN+'" AND R_NATECO<>""',True) ;
    if Not Q2.EOF then
      szNatEco:=Q2.FindField('R_NATECO').AsString
    else
      begin HM.Execute(30,'',X.AUXILIAIRE+' n''a pas de nature économique renseignée.'); Ferme(Q2); Exit ; end;
    Ferme(Q2);
  end;

  // Code IBAN et devise de l'émetteur
  Q:=OpenSQL('SELECT * from BANQUECP Where BQ_GENERAL="'+BQ_GENERAL.Value
            +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
  if Not Q.EOF then begin
    szCodeIBAN2:=Q.FindField('BQ_CODEIBAN').AsString;
    szDevise2:=Q.FindField('BQ_DEVISE').AsString;
    szCodePays:=Q.FindField('BQ_PAYS').AsString; // FQ 12430
    if szCodePays = '' then szCodePays := 'FRA';  // France par défaut // FQ 12430
  end;
  Ferme(Q) ;

  // Devises
  szDevise := X.Devise;

  // Code pays
  Q:=OpenSQL('SELECT PY_LIBELLE,PY_CODEISO2 from PAYS WHERE PY_PAYS="'+szPays+'"',True) ;
  if Not Q.EOF then begin
    szPays:= FullMajuscule(Q.FindField('PY_LIBELLE').AsString);
    szCodePays:=Q.FindField('PY_CODEISO2').AsString;
  end;
  Ferme(Q);
  if szCodePays = '' then begin HM.Execute(29,'',szPays+' n''est pas renseigné.'); Exit; end;

  // Montant de la remise
  D:=X.Debit; C:=X.Credit;
  if ((D>C) and (ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
  if ((C>D) and (Not ENC)) then BEGIN ExisteMauvais:=True ; Exit ; END ;
  if ENC then Montant:=C-D else Montant:=D-C ;
  if Montant=0 then Exit ;

  // Date d'éxécution
  dtDateExec:=StrToDate(DATEREMISE.Text) ;

  // Constitue l'enregistrement
  // Toutes les dates ont la forme AAAAMMJJ
  Inc(NumE);
  {JP 24/09/03 : Nouvelle fonction}
  Inf.NumEmetteur := Inttostr(NumE);
  Cpt.CodeIBAN    := szCodeIBAN;
  Cpt.Nom         := szNom;
  Cpt.Adresse1    := szAdresse1;
  Cpt.Ville       := szVille;
  Cpt.Pays        := szPays;
  Cpt.CodePays    := szCodePays;
  Cpt.Remise      := szRef;
  Inf.Montant     := Montant;
  Cpt.NatEco      := szNatEco; {JP 18/07/05 : FQ 16215}
  Inf.Imputation  := IntToStr(ImputFrais.ItemIndex + 13);
  Cpt.Devise      := szDevise;
  Cpt.CodeIBAN2   := szCodeIBAN2;
  Cpt.Devise2     := szDevise2;
  cpt.DateExect   := dtDateExec;
  St := TGenerationCFONB.DetailTrans(Cpt, Inf);

  TotalG:=Arrondi(TotalG+Montant,NbDec) ;
  WriteLn(FF,St) ;
  Result:=True ;
end;

function TFEXPCFONB.SommeDetailTRANS2(X: Tot_Cli): boolean;
var
  St, szVille, szCodeBIC, szPays : String ;
  Cpt : TTransfert;
  Inf : TInfosCFONB;
begin
  Result:=False ;
  {JP 05/07/05 : FQ 16161 : Amélioration des de la reprise des compte}
  GetInfosRibCompte(X.Auxiliaire, X.RIB, szVille, szPays, szCodeBIC, st);

  if (szCodeBIC = '') then begin
    HM.Execute(33,'', X.Auxiliaire +' n''est pas renseigné.');
    Exit ;
  end;

  if (szPays = '') then begin
    HM.Execute(34,'', X.Auxiliaire +' n''est pas renseigné.');
    Exit ;
  end;

  // Constitue l'enregistrement
  Inc(NumE);

  {JP 24/09/03 : Nouvelle fonction}
  Inf.NumEmetteur := Inttostr(NumE);
  Cpt.Ville       := szVille;
  Cpt.CodeBic     := szCodeBic;
  Cpt.Pays        := szPays;
  st := TGenerationCFONB.DetailTrans2(Cpt, Inf);
  WriteLn(FF,St) ;
  Result:=True ;

end;

{JP 05/07/05 : Récupèration d'infos diverses du Rib du compte courant
{---------------------------------------------------------------------------------------}
function TFEXPCFONB.GetInfosRibCompte(AuxiGen, RIB : string; var Ville, Pays, Bic, NomBQ : string) : string;
{---------------------------------------------------------------------------------------}
var
    lStEtab,
    lStGuichet,
    lStNumero,
    lStCle,
    lStDom : String ;
    Q      : TQuery;
begin

  DecodeRIB( lStEtab, lStGuichet, lStNumero, lStCle, lStDom, RIB, VH^.PaysLocalisation ) ;

  {On récupère tous les rib du compte en cours}
  Q := OpenSQL('SELECT R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB, R_CODEIBAN, R_DOMICILIATION, ' +
               'R_VILLE, R_PAYS, R_CODEBIC, R_PRINCIPAL FROM RIB WHERE R_AUXILIAIRE = "' + AuxiGen + '"', True);
  while not Q.EOF do begin
    {Si le Rib ou l'Iban correspondent au Rib de l'ecriture en cours, ou bien que l'on est sur le rib principal ...}
    if ( ( Q.FindField('R_ETABBQ').AsString = lStEtab ) and ( Q.FindField('R_GUICHET').AsString = lStGuichet ) and
         ( Q.FindField('R_NUMEROCOMPTE').AsString = lStNumero ) and ( Q.FindField('R_CLERIB').AsString = lStCle ) )
         or ( '*' + Q.FindField('R_CODEIBAN').AsString = RIB )
         or ( Q.FindField('R_PRINCIPAL').AsString = 'X' ) then
      begin
      {... O récupère la ville, le pays et le code bic}
      Ville := Q.FindField('R_VILLE'        ).AsString;
      BIC   := Q.FindField('R_CODEBIC'      ).AsString;
      Pays  := Q.FindField('R_PAYS'         ).AsString;
      NomBQ := Q.FindField('R_DOMICILIATION').AsString;

      {Si on a trouver le bon rib, on sort, sinon on a récupéré les infos du rib principal (par défaut),
       mais on continue la boucle avec l'espoir de trouver le bon rib}
      if not (Q.FindField('R_PRINCIPAL').AsString = 'X') then Break;
    end;
    Q.Next;
  end;
  Ferme(Q);

  {JP 25/09/03 : on récupère le code ISO 2}
  Result := Pays;
  Pays   := CodeIsoDuPays(Pays);
end;

{JP 05/07/05 : FQ 15493 : Génère un nom de fichier : faute de comprendre toutes les
               subtilités de l'unité, je préfère ne pas prendre de risque, mais je pense
               que ce morceau de code est appélè un peu trop de fois}
{---------------------------------------------------------------------------------------}
procedure TFEXPCFONB.GenereNomFichier(Nom_Fichier : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  Dir_Fic,
  Nom_Fic,
  Extension_Fic : string;
begin
  Dir_Fic := ExtractFilePath(Nom_Fichier);
  Nom_Fic := ExtractFileName(Nom_Fichier);
  {JP 06/06/07 : suite à FQ 18083 : Si le répertoire ou le nom de fichier est vide ou le répertoire incorrect ...}
  if (Trim(Dir_Fic) = '') or (Trim(Nom_Fic) = '') or
     not DirectoryExists(Dir_Fic) then begin
    {... On reprend les valeurs par défauf}
    case TypeExport.ItemIndex of
      0, 1, 2, 3, 4 : Nom_Fichier := TcbpPath.GetCegidUserLocalAppData + '\LCRBOR.TXT' ;
      5, 6, 7, 9    : Nom_Fichier := TcbpPath.GetCegidUserLocalAppData + '\VIREMENT.TXT' ;
      8             : Nom_Fichier := TcbpPath.GetCegidUserLocalAppData + '\PRELEV.TXT' ;
    end;
    Dir_Fic := ExtractFilePath(Nom_Fichier);
    Nom_Fic := ExtractFileName(Nom_Fichier);
    {Si pas de répertoire ou de nom de fichier, on sort}
    if (Trim(Dir_Fic) = '') or (Trim(Nom_Fic) = '') then Exit;
  end;

  {Suppression du premier caractère qui est le "."}
  Extension_Fic := Copy(ExtractFileExt(Nom_Fic), 2, 6);
  {Récupération ud nom de fichier sans extension et sans l'éventuelle date}
  Nom_Fic := ExtractFileNameOnly(Nom_Fic, True);
  n := Length(Dir_Fic);
  {Génération du nom de fichier à partir de la date}
  if Dir_Fic[n] = '\' then Nomfichier.Text := Dir_Fic + UProcGen.GetNomFichier(Dir_Fic, Nom_Fic, Extension_Fic)
                      else Nomfichier.Text := Dir_Fic + '\' + UProcGen.GetNomFichier(Dir_Fic, Nom_Fic, Extension_Fic);
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 19/01/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet de gérer le nom du fichier de l'export
Suite ........ : CFONB FQ 14321
Mots clefs ... : TRAITEMENT FICHIER EXPORT FQ 14321
*****************************************************************}
procedure TFEXPCFONB.TraitementFichier(var App:boolean);
var
  question : TmodalResult;
begin
  App:=False;
  if FileExists(NomFichier.Text) then
  begin
    {JP 07/06/07 : FQ 15490 : Modification du message}
    question := PGIAskCancel(TraduireMemoire('Le fichier ') + ExtractFileName(NomFichier.Text) +
                             TraduireMemoire(' existe déjà. Vous pouvez : ') + #13#13 +
                             Traduirememoire(' - Permettre la modification du nom et la création d''un nouveau fichier (Oui)') + #13 +
                             Traduirememoire(' - Écraser le contenu du fichier pour générer l''export (Non)')  + #13 +
                             Traduirememoire(' - Faire l''export à la suite du contenu existant dans le fichier (Annuler)'), 'Avertissement');
    if question = mrCancel then
    begin
      App:= True;
      Exit;
    end
    else if question = mrNo then Exit
    {JP 04/07/05 : FQ 15493 : Nouvelle gestion du nommage automatique des fichiers}
    else if question = mrYes then GenereNomFichier(NomFichier.Text);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 19/01/2005
Modifié le ... : 05/07/2005
Description .. : Fonction qui permet d'extraire le nom du fichier sans
Suite ........ : extention dans le chemin du fichier
Suite ........ : JP 05/07/05, FQ 15493 : ajout Racine nom, si l'on veut
Suite ........ : enlever les "_" du nom et ce qui suit
Mots clefs ... : NOM FICHIER SANS EXTENSION
*****************************************************************}
Function TFEXPCFONB.ExtractFileNameOnly(FileName:String; RacineNom : Boolean): String;
var ExtPart : TFileName;
    lgExt   : Integer;
begin
 ExtPart := ExtractFileExt(FileName);
 LgExt:=Length(ExtPart);
 Delete(FileName,Length(FileName)-lgExt+1,lgExt);
 Result:=FileName;
 if RacineNom and (Pos('_', Result) > 0) then {JP 05/07/05 : FQ 15493}
   Result := Copy(Result, 1, Pos('_', Result) - 1);
end;

procedure TFEXPCFONB.RecupParam;
begin
  if not Assigned( FTobScenario ) then Exit ;

  // Ajout champ cfonb
  FTobScenario.AddChampSupValeur( 'CFONB_TYPEEXPORT',    TypeExport.ItemIndex ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_BQ_GENERAL',    BQ_GENERAL.Value ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_NOMFICHIER',    NomFichier.Text ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_DOCUMENT',      Document.Value ) ;

  if CRibTiers.checked
    then FTobScenario.AddChampSupValeur( 'CFONB_CRIBTIERS',     'X' )
    else FTobScenario.AddChampSupValeur( 'CFONB_CRIBTIERS',     '-' ) ;

  if CCumulTiers.Checked
    then FTobScenario.AddChampSupValeur( 'CFONB_CCUMULTIERS',   'X' )
    else FTobScenario.AddChampSupValeur( 'CFONB_CCUMULTIERS',   '-' ) ;

  FTobScenario.AddChampSupValeur( 'CFONB_DATEREMISE',           DATEREMISE.Text   ) ;

  if RefTireLib.Checked
    then FTobScenario.AddChampSupValeur( 'CFONB_REFTIRELIB',    'X' )
    else FTobScenario.AddChampSupValeur( 'CFONB_REFTIRELIB',    '-' ) ;

  FTobScenario.AddChampSupValeur( 'CFONB_CBO_NATECO',           Cbo_NatEco.value  ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_TXT_REFREMISE',        Txt_RefRemise.Text    ) ;

  FTobScenario.AddChampSupValeur( 'CFONB_RENVOI',               Renvoi.itemIndex ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_IMPUTFRAIS',           ImputFrais.ItemIndex ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_TYPEDEBIT',            TypeDebit.ItemIndex ) ;

  // champ de gestion pour le traitement
  FTobScenario.AddChampSupValeur( 'CFONB_PREMIERLOT',          'X' ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_DERNIERLOT',          '-' ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_ERREUR',               0 ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_EXISTEMAUVAIS',        '-' ) ;
  FTobScenario.AddChampSupValeur( 'CFONB_TOTALG',               0 ) ;

  FTobScenario.AddChampSupValeur( 'CFONB_NUMCFONB',             GetNumCFONB( FormatCFONB ) ) ;

end;

function TFEXPCFONB.VerifParam: Boolean;
Var Dom        : String ;
    Etab       : String ;
    Guichet    : String ;
    NumCompte  : String ;
    szCodeBIC  : String ;
    szCodeIBAN : String ;
    szCodeDEV  : String ;
    Libelle    : String ;
    lQBqe      : TQuery ;
begin

  result := False ;

  // Format d'export
  if TypeExport.ItemIndex<0 then
    begin
    HM.Execute(10,'','') ;
    Exit ;
    end ;

  // Banque
  if BQ_GENERAL.Value='' then
    begin
    HM.Execute(0,'','') ;
    Exit ;
    end ;

  // Nom de ficher
  if NomFichier.Text='' then
    begin
    HM.Execute(1,'','') ;
    Exit ;
    end ;

  // Données de l'emetteur
  lQBqe := OpenSQL('SELECT BQ_LIBELLE, BQ_DOMICILIATION, BQ_ETABBQ, BQ_GUICHET, '
                        + 'BQ_NUMEROCOMPTE, BQ_NUMEMETLCR, BQ_NUMEMETVIR, '
                        + 'BQ_CODEBIC, BQ_CODEIBAN, BQ_DEVISE '
                        + 'FROM BANQUECP WHERE BQ_GENERAL="'+BQ_GENERAL.Value
                        +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"', True ) ; // 19/10/2006 YMO Multisociétés
  if Not lQBqe.Eof then
     begin
     Libelle     := lQBqe.FindField('BQ_LIBELLE').AsString ;
     Dom         := FullMajuscule(lQBqe.FindField('BQ_DOMICILIATION').AsString) ;
     Etab        := FullMajuscule(lQBqe.FindField('BQ_ETABBQ').AsString) ;
     Guichet     := lQBqe.FindField('BQ_GUICHET').AsString ;
     NumCompte   := lQBqe.FindField('BQ_NUMEROCOMPTE').AsString ;
     Case TypeExport.ItemIndex of
          0,1,2,3,4 : NumEmetteur := lQBqe.FindField('BQ_NUMEMETLCR').AsString ;
          5,6,7     : NumEmetteur := lQBqe.FindField('BQ_NUMEMETVIR').AsString ;
          8         : NumEmetteur := GetParamSocSecur('SO_LIBRETEXTE0', '') ;  // Lecture Société Emetteur national
          9         : begin
                      NumEmetteur := '' ;
                      szCodeBIC   := lQBqe.FindField('BQ_CODEBIC').AsString ;
                      szCodeIBAN  := lQBqe.FindField('BQ_CODEIBAN').AsString ;
                      szCodeDEV   := lQBqe.FindField('BQ_DEVISE').AsString ;
                      end ;
      end ;

     end
   else
     begin
     Libelle      := '' ;
     Dom          := '' ;
     Etab         := '' ;
     Guichet      := '' ;
     NumCompte    := '' ;
     NumEmetteur  := '' ;
     szCodeBIC    := '' ;
     szCodeIBAN   := '' ;
     szCodeDEV    := '' ;
     end ;

  Ferme( lQBqe ) ;

  if TypeExport.ItemIndex = 9 then
  // transferts internationnaux
    begin
    // Référence de la remise
    if (txt_RefRemise.Text='') then
      begin
      HM.Execute(32,caption,'');
      if txt_RefRemise.CanFocus then
        txt_RefRemise.SetFocus;
      Exit;
      end;
    if szCodeBIC = ''  then  begin HM.Execute(23,caption,''); Exit; end;
    if szCodeIBAN = '' then  begin HM.Execute(24,caption,''); Exit; end;
    if szCodeDEV = ''  then  begin HM.Execute(25,caption,''); Exit; end;
    end
  else
  // sinon test émetteur
    if Not EmetOK(NumEmetteur,Dom,Etab,Guichet,NumCompte) then Exit ;

  FTobScenario.AddChampSupValeur( 'CFONB_NUMEMETTEUR',   NumEmetteur ) ;
  // Recup info RIB
  FTobScenario.AddChampSupValeur( 'BQ_LIBELLE',          Libelle ) ;
  FTobScenario.AddChampSupValeur( 'BQ_DOMICILIATION',    Dom ) ;
  FTobScenario.AddChampSupValeur( 'BQ_ETABBQ',           Etab ) ;
  FTobScenario.AddChampSupValeur( 'BQ_GUICHET',          Guichet ) ;
  FTobScenario.AddChampSupValeur( 'BQ_NUMEROCOMPTE',     NumCompte ) ;
  FTobScenario.AddChampSupValeur( 'BQ_CODEBIC',          szCodeBIC ) ;
  FTobScenario.AddChampSupValeur( 'BQ_CODEIBAN',         szCodeIBAN ) ;
  FTobScenario.AddChampSupValeur( 'BQ_DEVISE',           szCodeDEV ) ;

  // tout est ok
  result := True ;

end;

procedure TFEXPCFONB.SetParam;
begin
  if not Assigned( FTobScenario ) then Exit ;

  TypeExport.ItemIndex := FTobScenario.GetInteger('CFONB_TYPEEXPORT') ;
  BQ_GENERAL.Value     := FTobScenario.GetString('CFONB_BQ_GENERAL') ;
  NomFichier.Text      := FTobScenario.GetString('CFONB_NOMFICHIER') ;

  CRibTiers.checked    := FTobScenario.GetString('CFONB_CRIBTIERS')='X'  ;
  CCumulTiers.Checked  := FTobScenario.GetString('CFONB_CCUMULTIERS')='X' ;
  DATEREMISE.Text      := FTobScenario.GetString('CFONB_DATEREMISE') ;

  RefTireLib.Checked := FTobScenario.GetString('CFONB_REFTIRELIB') = 'X';

  Cbo_NatEco.value     := FTobScenario.GetString( 'CFONB_CBO_NATECO');
  Txt_RefRemise.Text   := FTobScenario.GetString( 'CFONB_TXT_REFREMISE');

  Renvoi.itemIndex     := FTobScenario.GetInteger( 'CFONB_RENVOI' ) ;
  ImputFrais.itemIndex := FTobScenario.GetInteger( 'CFONB_IMPUTFRAIS' ) ;
  TypeDebit.itemIndex  := FTobScenario.GetInteger( 'CFONB_TYPEDEBIT' ) ;

  // paramètres non utilisé...
  Document.Value       := '' ;
  FApercu.Checked      := False ;
  cTeletrans.Checked   := False ;

end;

function TFEXPCFONB.ExecuteBatch( OkAppend : Boolean ) : Boolean ;
Var i : integer ;
    O : Tob ;
begin

  RemplirLesEches ;
  ExportEnEuro := FALSE ;

  Case REnvoi.ItemIndex of
    0 : begin
        OkDebit      := VH^.TenueEuro ;
        ExportEnEuro := TRUE ;
        end ;
    1 : begin
        if TECHE.Count>0 then
          OkDebit:=True;
        ExportEnEuro := FALSE ;
        If (OkDebit And VH^.TenueEuro) Or ((Not OkDebit) And (Not VH^.TenueEuro)) Then
          ExportEnEuro := TRUE ;
        end ;
    end ;

  if OkDebit
    then NbDec:=V_PGI.OkDecV
    else NbDec:=V_PGI.OkDecE ;

  ExisteMauvais:=False ;
  VideListe(TCLI) ;

  for i:=0 to TECHE.Count-1 do
      BEGIN
      O:=TOB(TECHE[i]) ;
// Affecation du rib non prévu pour le moment
//      if not PutLeRIB(O) then exit; // Affectation du RIB
      SommeCli(O) ; {Cumul par tiers pour contrôle}
      END ;

  result := LanceCFONB( OkAppend ) ;
  if result then
    DeflagCFONBBatch ;

end;

procedure TFEXPCFONB.DeflagCFONBBatch;
Var i          : integer ;
    O          : TOB ;
    SQL        : String ;
    lStDossier : string ;
    lTotC      : Tot_Cli ;
begin
  if TECHE.Count<=0 then Exit ;
  if not FBoModeBatch then Exit ;

  for i:=0 to TECHE.Count-1 do
    begin
    O      := TOB(TECHE[i]) ;

    // Faut-il mettre à jour l'écriture  ?
    if CCumulTiers.Checked then // cas des cumuls
      begin
      if O.GetString('E_AUXILIAIRE')=''
        then lTotC := FindTotCli( O.GetString('E_GENERAL'), O.GetString('E_DEVISE') )
        else lTotC := FindTotCli( O.GetString('E_AUXILIAIRE' ), O.GetString('E_DEVISE') ) ;
      if lTotC = nil
        then Continue
        else if lTotC.Exporte then Continue ;
      end
    else // cas par lignes
      if O.GetString('E_CFONBOK') = 'X' then Continue ;

    // Gestion multisoc
    if EstMultiSoc and ( O.GetString('E_SOCIETE') <> GetParamSocSecur('SO_SOCIETE', '') ) then
      begin
      if O.GetNumChamp('SYSDOSSIER') > 0
        then lStDossier := O.GetString('SYSDOSSIER')
        else Continue ;
      end
    else lStDossier := V_PGI.SchemaName ;

    // MAJ BASE
    SQL := 'UPDATE ' + GetTableDossier( lStDossier, 'ECRITURE' )
           + ' SET E_CFONBOK="-", E_NUMCFONB="" WHERE ' + whereEcritureTob( tsGene, O, True ) ;
    ExecuteSQL(SQL) ;
    
    end ;

end;

function TFEXPCFONB.findTotCli( vAux: string ; vDev : string ): Tot_Cli;
var i : integer ;
    X : Tot_Cli ;
begin
  result := nil ;
  for i:=0 to TCli.Count-1 do
    begin
    X := Tot_Cli( TCli[i] ) ;
    if ( X.Auxiliaire = vAux ) and ( X.Devise = vDev ) then
      begin
      result := X ;
      Exit ;
      end ;
    end ;
end;


function TFEXPCFONB.RecupGenOuAux(O: TOB; AvecLib: boolean; var Cpt, Lib: String): Boolean;
Var Q : TQuery ;
    TIDTIC : Boolean ;
BEGIN
Cpt:=O.GetValue('E_AUXILIAIRE') ; TIDTIC:=FALSE ;
If (Cpt='') And (O.GetValue('E_NUMECHE')>0) Then BEGIN TIDTIC:=TRUE ; Cpt:=O.GetValue('E_GENERAL') ; END ;
Lib:='' ;
If AvecLib Then
  BEGIN
  if FBoModeBatch then
    begin
    If TIDTIC Then Lib := FullMajuscule( O.GetString('G_LIBELLE') )
              Else Lib := FullMajuscule( O.GetString('T_LIBELLE') ) ;
    end
  else
    begin
    If TIDTIC Then Q:=OpenSQL('Select G_LIBELLE from GENERAUX Where G_GENERAL="'+Cpt+'"',True)
              Else Q:=OpenSQL('Select T_LIBELLE from TIERS Where T_AUXILIAIRE="'+Cpt+'"',True) ;
    If Not Q.Eof Then
      Lib:= FullMajuscule(Q.Fields[0].AsString);
    Ferme(Q) ;
    end ;
  END ;
Result:=TIDTIC ;
END ;

procedure TFEXPCFONB.FormDestroy(Sender: TObject);
begin

  VideListe(TCli) ;
  TCli.Free ;

  VideListe(TECHES) ;
  TECHES.Free ;

  VideListe(TECHE1) ;
  TECHE1.Free ;

  if Assigned(FTobMP) then
    FreeAndNil( FTobMP ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 01/01/1900
Modifié le ... : 21/03/2007
Description .. : 
Suite ........ : 
Suite ........ : 
Suite ........ : FQ19759 : pb de récupération du code acceptation
Suite ........ : ( SBO 21/03/2007)
Mots clefs ... : 
*****************************************************************}
function TFEXPCFONB.GetTobMP(vCodeMP: string): Tob;
var lTob : Tob ;
    Q    : TQuery ;
begin

  result := Nil ;
  if vCodeMP='' then Exit ;

  result := FTobMP.FindFirst( ['MP_MODEPAIE'], [ vCodeMP ], True ) ;
  if result = nil  then
    begin
    Q := OpenSQL('SELECT MP_MODEPAIE,MP_CODEACCEPT,MP_CATEGORIE from MODEPAIE Where MP_MODEPAIE="'+vCodeMP+'"',True) ;
    if Not Q.EOF then
      begin
      lTob := Tob.Create('MODEPAIE', FTobMP, -1 ) ;
      if lTob.SelectDB('', Q) then
        result := lTob ;
      end ;
    Ferme(Q) ;
    end ;

end;

{$ENDIF}

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.EmetteurVIR(Cpt : TCompte; Inf : TInfosCFONB) : string;
{---------------------------------------------------------------------------------------}
var
  Date5 : string;
begin
  Date5 := FormatDateTime('ddmmyy',Inf.DateCreat);
  Date5 := Copy(Date5, 1, 4) + Copy(Date5, 6, 1);

  {Zone 1 à 2 : Type d'enregistrement :  03 = Emetteur
   Zone 3 à 4 : Code opération        :  02 = pour virement ordinaire
   Zone 5 à 12 : Réservée
   Zone 13 à 18 : N° émetteur
   Zone 19 : Code décalage
   Zone 20 à 25 : Réservée
   Zone 26 à 30 : Date au format JJMMA
   Zone 31 à 54 : Nom/Raison social du bénéficiaire
   Zone 55 à 61 : Référence de la remise
   Zone 62 à 80 : Réservée
   Zone 81 : E pour Euro ; F pour Franc
   Zone 82 à 86 : Réservée
   Zone 87 à 91 : Code guichet donneur d'ordre
   Zone 92 à 102 : N° compte donneur d'ordre
   Zone 103 à 118 : Identifiant du donneur d'ordre
   Zone 119 à 149 : Réservée
   Zone 150 à 154 : Code établissement du donneur d'ordre
   Zone 155 à 160 : Réservée}

  Result := '03' + Inf.Code + Format_String('', 8) + Format_String(Inf.NumEmetteur, 6) +
            Format_String(Cpt.Divers, 1) + Format_String('', 6) + Date5 +
            Format_String(Cpt.Domiciliation, 24) + Format_String('', 26) + 'E' + Format_String('',5) +
            CompleteZero(Cpt.CodeGuichet, 5) + CompleteZero(Cpt.NumCompte, 11) + Format_String('',47) +
            CompleteZero(Cpt.CodeBanque, 5) + Format_String('', 6);
end;

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.DetailVIR(Cpt : TCompte; Inf : TInfosCFONB) : string;
{---------------------------------------------------------------------------------------}
begin
  {Zone 1 à 2 : Type d'enregistrement : 06 = Detail
  Zone 3 à 4 : Code opération         : 02 = Virement ordinaire  /76 : virements de compte à compte
  Zone 5 à 12 : Réservée
  Zone 13 à 18 : N° émetteur
  Zone 19 à 30 : Référence interne du donneur d'ordre
  Zone 31 à 54 : Nom/Raison social du bénéficiaire
  Zone 55 à 78 : Domiciliation / 20 + Résident et code pays 4 ???
  Zone 79 à 86 : Déclaration à la balnce des paiements (suite)
  Zone 87 à 91 : Code guichet banque bénéficiaire
  Zone 92 à 102 : N° compte
  Zone 103 à 118 : Montant virement
  Zone 119 à 147 : Libellé virement
  Zone 148 à 149 : Code motif rejet
  Zone 150 à 154 : Code établissement bénéficiaire
  Zone 155 à 160 : Réservée}
  Result := '06' + Inf.Code + Format_String('', 8) + Format_String(Inf.NumEmetteur, 6) + Format_String(Cpt.Divers, 12) +
           Format_String(Cpt.RaisonSoc, 24) + Format_String(Cpt.Domiciliation, 24) + Format_String('',8) +
           CompleteZero(Cpt.CodeGuichet, 5) + CompleteZero(Cpt.NumCompte, 11) + CompleteZero(InttoStr(Round(100.0 * Inf.Montant)),16) +
           Format_String(Inf.Motif, 29) + Format_String('', 2) + CompleteZero(Cpt.CodeBanque, 5) + Format_String('', 6);
end;

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.TotalVIR(Inf : TInfosCFONB) : string;
{---------------------------------------------------------------------------------------}
begin
  {Ecriture de l'enregistrement total (fin des virements de même compte source)}
  {Zone 1 à 2    : Type d'enregistrement : 08 = Total
  Zone 3 à 4     : Code opération         : 02 = Virement ordinaire
  Zone 5 à 12    : Réservée
  Zone 13 à 18   : N° émetteur
  Zone 19 à 102  : Réservée
  Zone 103 à 118 : Montant Total
  Zone 119 à 160 : Réservée}
  Result := '08' + Inf.Code + Format_String('', 8) + Format_String(Inf.NumEmetteur, 6) + Format_String('', 84) +
            CompleteZero(InttoStr(Round(100.0 * Inf.Montant)), 16) + Format_String('', 42);
end;

{   Transferts   }

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.EmetteurTrans(Cpt : TTransfert; Inf : TInfosCFONB; Soc : TInfosSoc) : string;
{---------------------------------------------------------------------------------------}
var
  TypCpt : string;
begin
  {Zone 1   à 2   : 2  : Code enregistrement
  Zone 3   à 4   : 2  : Code opération
  Zone 5   à 10  : 6  : N° séquentiel
  Zone 11  à 18  : 8  : Date de création
  Zone 19  à 53  : 35 : Raison sociale de l'émetteur
  Zone 54  à 88  : 35 : Adresse N° et rue
  Zone 89  à 123 : 35 : Adresse Ville
  Zone 124 à 158 : 35 : Adresse Pays
  Zone 159 à 172 : 14 : N° SIRET
  Zone 173 à 188 : 16 : Référence remise
  Zone 189 à 199 : 11 : Code BIC de la banque émettrice
  Zone 200       : 1  : Identifiant du compte à débiter 0:Autre ; 1:IBAN ; 2:BBAN
  Zone 201 à 234 : 34 : Compte à débiter
  Zone 235 à 237 : 3  : Code devise du compte à débiter
  Zone 238 à 253 : 16 : Identifiant du client
  Zone 254       : 1  : Identifiant du compte de frais 0:Autre ; 1:IBAN ; 2:BBAN
  Zone 255 à 288 : 34 : Compte de frais
  Zone 289 à 291 : 3  : Code devise du compte de frais
  Zone 292 à 307 : 16 : Zone réservée
  Zone 308       : 1  : Indice du type de débit de la remise 1:Débit global de la remise ; 2:Débit unitaire par opération ; 3:Débit global par devise de transfert
  Zone 309       : 1  : Indice du type de la remise 1:Mono-date d'éxécution et mono-devise ; 2:Mono-date d'éxécution et multi-devises ; 3:Multi-date d'éxécution et mono-devise ; 4:Multi-date d'éxécution et Multi-devises
  Zone 310 à 317 : 8  : Date d'exécution souhaitée
  Zone 318 à 320 : 3  : Code devise des transferts : Obligatoire pour les remises de type 'mono-devise' (voir 309) sinon,ne pas renseigner}

  {JP 10/07/07 : FQ 17553 : gestion du BBAN}
       if IsIBAN(Cpt.CodeIBAN) then TypCpt := '1'
  else if IsBBAN(Cpt.CodeIBAN) then TypCpt := '2'
                               else TypCpt := '0';

  Result := '03PI000001' + FormatDateTime('yyyymmdd', Inf.DateCreat) +
            Format_String(Soc.Raison, 35) +
            Format_String(Soc.Adr1, 35) +
            Format_String(Soc.Ville, 35) +
            Format_String(Soc.Pays,35) +
            Format_String(Soc.Siret, 14) +
            Format_String(Cpt.Remise, 16) +
            Format_String(Cpt.CodeBIC,11) + '1' +
            Format_String(Cpt.CodeIBAN, 34) +
            Format_String(Cpt.Devise, 3) +
            Format_String(Cpt.IdClient, 16) + //'1' + // 13466
            TypCpt + {JP 10/07/07 : FQ 17553}
            Format_String(Cpt.CodeIBAN, 34) +
            Format_String(Cpt.Devise, 3) +
            Format_String('', 16) +
            {JP 21/07/05 : FQ 15388 : a priori, après relecture du cahier des charges, ces zones sont
                           facultatives dans la mesure où elles sont renseignées dans le détail, ce
                           qui est le cas ici.
            Cpt.TypeRemise + '2' +
            FormatDateTime('yyyymmdd', Cpt.DateExect) +
            Format_String('', 3);}
            {JP 09/08/05 : FQ 15388 suite : j'y suis allé un peu fort avec le type de remise ci-dessus}
            Cpt.TypeRemise +
            IIF(Cpt.TypeDevise <> '', '1', '2') +
            {JP 28/09/07 : FQ 21529 : Attention, quand on gèrera ci dessus le '3' et '4', il ne
                           faudra plus de mettre de date !}
            FormatDateTime('yyyymmdd', Cpt.DateExect) +
            IIF(Cpt.TypeDevise <> '', Cpt.TypeDevise, '   ');
end;

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.DetailTrans(Cpt : TTransfert; Inf : TInfosCFONB) : string;
{---------------------------------------------------------------------------------------}
var
  TypCpt : string;
begin
  {Zone 1 à 2     : 2  : Code enregistrement
   Zone 3 à 4     : 2  : Code opération
   Zone 5 à 10    : 6  : N° séquentiel
   Zone 11        : 1  : Identifiant du compte du bénéficiaire  0:Autre ; 1:IBAN ; 2:BBAN
   Zone 12 à 45   : 34 : Compte du bénéficiaire
   Zone 46 à 80   : 35 : Nom du bénéficiaire
   Zone 81 à 115  : 35 : Adresse N° et rue
   Zone 116 à 150 : 35 : Adresse Ville
   Zone 151 à 185 : 35 : Adresse Pays
   Zone 186 à 202 : 17 : Identification nationale du bénéficiaire
   Zone 203 à 204 : 2  : Code pays du bénéficiaire
   Zone 205 à 220 : 16 : Référence de l'opération
   Zone 221       : 1  : Qualifiant du montant de l'ordre T:Montant dans la devise du transfert D:Montant exprimé dans la devise du compte à débiter
   Zone 222 à 225 : 4  : Zone réservée
   Zone 226 à 239 : 14 : Montant de l'ordre
   Zone 240       : 1  : Nombre de décimales
   Zone 241       : 1  : Zone réservée
   Zone 242 à 244 : 3  : Motif économique
   Zone 245 à 246 : 2  : Codepays pour la déclaration BDF
   Zone 247       : 1  : Mode de règlement 0:Autre ; 1:Par chèque de la banque émettrice ; 2:Par chèque de la banque réceptrice
   Zone 248 à 249 : 2  : Code imputation des frais 13:Bénéficiaire 14:Emetteur et bénéficiaire 15:Emetteur
   Zone 250       : 1  : Identifiant du compte de frais 0:Autre ; 1:IBAN ; 2:BBAN
   Zone 251 à 284 : 34 : Compte de frais
   Zone 285 à 287 : 3  : Code devise du compte de frais
   Zone 251 à 284 : 34 : Compte de frais
   Zone 285 à 287 : 3  : Code devise du compte de frais
   Zone 288 à 309 : 22 : Zone réservée
   Zone 310 à 317 : 8  : Date d'exécution souhaitée
   Zone 318 à 320 : 3  : Code devise du transfert}

  {JP 10/07/07 : FQ 17553 : gestion du BBAN}
       if IsIBAN(Cpt.CodeIBAN) then TypCpt := '1'
  else if IsBBAN(Cpt.CodeIBAN) then TypCpt := '2'
                               else TypCpt := '0';

  Result := '04PI' + CompleteZero(Inf.NumEmetteur, 6) + '1' + Format_String(Cpt.CodeIBAN,34)+ Format_String(Cpt.Nom,35) +
            Format_String(Cpt.Adresse1,35) + Format_String(Cpt.Ville,35) + Format_String(Cpt.Pays,35) +
            Format_String('', 17) + Format_String(Cpt.CodePays,2) + Format_String(Cpt.Remise, 16)+ 'T' + Format_String('', 4) +
            CompleteZero((InttoStr(Round(100.0 * Inf.Montant))), 14) + '2 ' + Format_String(Cpt.NatEco, 3) +
            Format_String(Cpt.CodePays, 2) + '0' + Inf.Imputation +
            TypCpt; {JP 10/07/07 : FQ 17553}


  {Imputation des frais}
  if (Inf.Imputation = '13') then
    Result := Result + Format_String(Cpt.CodeIBAN, 34) + Format_String(Cpt.Devise, 3)
  else
    Result := Result + Format_String(Cpt.CodeIBAN2, 34) + Format_String(Cpt.Devise2, 3);

  Result := Result + Format_String('', 22) +

  {JP 28/09/07 : FQ 21529 : Comme on ne gère pas les 3 et 4 (multi dates), on ne met pas de date}
  Format_String('', 8) + //FormatDateTime('yyyymmdd', Cpt.DateExect) +
  // GCO - 19/07/20006 - FQ 18603
  IIF(Cpt.TypeDevise <> '', Cpt.TypeDevise, '   ');
  //Format_String(Cpt.Devise, 3);
end;

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.DetailTrans2(Cpt : TTransfert; Inf : TInfosCFONB) : string;
{---------------------------------------------------------------------------------------}
begin
  {Zone 1 à 2     : 2   : Code enregistrement
   Zone 3 à 4     : 2   : Code opération
   Zone 5 à 10    : 6   : N° séquentiel
   Zone 11 à 45   : 35  : Nom de la banque du bénéficiaire
   Zone 46 à 80   : 35  : N° et nom de rue de l'agence
   Zone 81 à 115  : 35  : Ville de l'agence
   Zone 116 à 150 : 35  : Pays de l'agence
   Zone 151 à 161 : 11  : Code BIC
   Zone 162 à 163 : 2   : Code pays à la norme ISO
   Zone 164 à 320 : 157 : Réservée}
  Result := '05PI' + CompleteZero(Inf.NumEmetteur, 6) +
            Format_String(Cpt.Banque, 35) +
            Format_String(Cpt.Agence, 35) +
            Format_String(Cpt.Ville, 35) +
            Format_String(Cpt.Pays2, 35) +
            Format_String(Cpt.CodeBic, 11) +
            Format_String(Cpt.Pays, 2) +
            Format_String('', 157);
end;

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.TotalTrans(Cpt : TTransfert; Inf : TInfosCFONB; Soc : TInfosSoc) : string;
{---------------------------------------------------------------------------------------}
var
  TypCpt : string;
begin
  {Zone 1 à 2     : 2   : Code enregistrement :  08 = Total
  Zone 3 à 4     : 2   : Code opération
  Zone 5 à 10    : 6   : N° séquentiel
  Zone 11 à 18   : 8   : Date de création
  Zone 19 à 158  : 140 : Zone réservée
  Zone 159 à 172 : 14  : N° SIRET de l'émetteur
  Zone 173 à 188 : 16  : Référence remise
  Zone 189 à 199 : 11  : Zone réservée
  Zone 200       : 1   : Identifiant du compte à débiter 0:Autre ; 1:IBAN ; 2:BBAN
  Zone 201 à 234 : 34  : Compte à débiter
  Zone 235 à 237 : 3   : Code devise du compte à débiter
  Zone 238 à 253 : 16  : Identifiant du client
  Zone 254 à 271 : 18  : Total de contrôle
  Zone 272 à 320 : 49  : Zone réservée}

  {JP 10/07/07 : FQ 17553 : gestion du BBAN}
       if IsIBAN(Cpt.CodeIBAN) then TypCpt := '1'
  else if IsBBAN(Cpt.CodeIBAN) then TypCpt := '2'
                               else TypCpt := '0';

  Result := '08PI' + CompleteZero(Inf.NumEmetteur, 6) +
            FormatDateTime('yyyymmdd', Inf.DateCreat) +
            Format_String('', 140) +
            Format_String(Soc.Siret, 14) +
            Format_String(Cpt.Remise, 16) +
            Format_String('', 11) + 
            TypCpt + {JP 10/07/07 : FQ 17553}
            Format_String(Cpt.CodeIBAN, 34) +
            Format_String(Cpt.Devise, 3) +
            Format_String(Cpt.IdClient,16) + // 13466
            CompleteZero(InttoStr(Round(100.0 * Inf.Montant)), 18) +
            Format_String('', 49);
end;

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.CompleteZero(St : string; L : Integer) : string;
{---------------------------------------------------------------------------------------}
begin
  St := Copy(St, 1, L);
  while Length(St) < L do St := '0' + St;
  Result := St;
end;

{---------------------------------------------------------------------------------------}
class function TGenerationCFONB.Date5(DD : TDateTime) : string ;
{---------------------------------------------------------------------------------------}
var
  St : string ;
begin
  St := FormatDateTime('ddmmyy', DD) ;
  Result := Copy(St, 1, 4) + Copy(St, 6, 1);
end;


end.


