unit UConnectWSConst;

interface

type
  T_WSEntryType      = (wsetNone, wsetDocument, wsetPayment, wsetPayer, wsetExtourne, wsetSubContractPayment, wsetStock);
  T_WSPSocType       = (wspsNone, wspsServer, wspsPort, wspsFolder, wspsLastSynchro);
  T_WSDataService    = (  wsdsNone               // Aucun                Doc génération fichier .TRA
                        , wsdsThird              // Tiers                page 34
                        , wsdsAnalyticalSection  // Section analytique   page 25
                        , wsdsAccount            // Comptes comptable
                        , wsdsJournal            // Journaux comptable
                        , wsdsBankId             // RIB                  page 51
                        , wsdsChoixCod           // Table CHOIXCOD       page 64 (GDM;NVR;GOR;JUR;LGU;GCT;RTV;GZC;SCC;TRC)
                        , wsdsCommon             // Table COMMUN
                        , wsdsChoixExt           // Table CHOIXEXT 
                        , wsdsRecovery           // Relance
                        , wsdsCountry            // Pays                 page 
                        , wsdsCurrency           // Devise               page 22
                        , wsdsChangeRate         // Taux de change       page 77
                        , wsdsCorrespondence     // Correspondance
                        , wsdsPaymenChoice       // Mode de règlement    page 20
                        , wsdsFiscalYear         // Exercice comptable
                        , wsdsSocietyParameters  // Paramètre société
                        , wsdsEstablishment      // Etablissement
                        , wsdsPaymentMode        // Mode de paiment
                        , wsdsZipCode            // Code postaux
                        , wsdsContact            // Contact
                        , wsdsFieldsList         // Liste des champs
                        , wsdsAccForPayment      // Liste des écritures pour récup des règlements
                        , wsdsDocPayment         // Table ACOMPTES
                        );
  T_WSInfoFromDSType = (wsidNone, wsidTableName, wsidFieldsKey, wsidExcludeFields, wsidFieldsList, wsidRequest);
  T_WSAction         = (wsacNone, wsacUpdate, wsacInsert);
  T_WSType           = (wstypNone, wstypUpload, wstypImport, wstypGetState, wstypGetReport);
  T_SvcSyncBTPY2Log  = (ssbylNone, ssbylLog, ssbylWindows);

  T_WSBTPValues      = Record
                         ConnectionName : string;
                         UserAdmin      : string;
                         Server         : string;
                         DataBase       : string;
                         LastSynchro    : string;
                       end;
  T_WSY2Values       = Record
                         ConnectionName : string;
                         Server         : string;
                         DataBase       : string;
                       end;
  T_WSConnectionValues = Record
                           UserAdmin   : string;
                           BTPServer   : string;
                           BTPDataBase : string;
                           BTPLastSync : string;
                           Y2Server    : string;
                           Y2DataBase  : string;
                           Y2LastSync  : string;
                         end;
  T_WSLogValues       = Record
                          LogLevel          : integer;
                          LogMoMaxSize      : double;
                          LogMaxQty         : integer;
                          LogDebug          : integer;
                          LogDebugMoMaxSize : double;
                          DebugEvents       : integer;
                          OneLogPerDay      : boolean;
                       end;

  T_WSResponseImportEntries = record
                                Error        : string;
                                HasError     : Boolean;
                                ProcessId    : string;
                                ReportFileId : string;
                                Terminated   : Boolean;
                              end;

  T_WSDocumentInf = Record
                   dType   : string;
                   dStub   : string;
                   dNumber : integer;
                   dIndex  : integer;
                 end;

  TGetFromDSType = class (TObject)
  public
    class function dstPrefix(DSType : T_WSDataService) : string;
    class function dstTableName(DSType : T_WSDataService) : string;
    class function dstViewName(DSType : T_WSDataService) : string;
    class function dstWSName(DSType : T_WSDataService) : string;
    class function ExtractType(TableName : string) : string; overload;
    class function ExtractType(DSType : T_WSDataService) : string; overload;
    class procedure WriteLog(TypeDebug: T_SvcSyncBTPY2Log; Text: string; LogValues : T_WSLogValues; LineLevel: integer; WithoutDateTime: Boolean = true);
    class function dstFiedsList(DSType : T_WSDataService) : string;
  end;

const
  WSCDS_UpdDateFieldName          = 'DateModif';
  WSCDS_CreDateFieldName          = 'DateCreation';
  WSCDS_SocLastSync               = 'SO_BTWSLASTSYNC';
  WSCDS_SocCegidDos               = 'SO_BTWSCEGIDDOS';
  WSCDS_SocNumPort                = 'SO_BTWSCEGIDPORT';
  WSCDS_SocServer                 = 'SO_BTWSSERVEUR';
  WSCDS_EmptyValue                = '#None#';
  WSCDS_GetDataOk                 = 'OK';
  WSCDS_IndiceField               = '#INDICE';                                                                               
  WSCDS_NextUrlValue              = '"@odata.nextLink":"';
  WSCDS_SectionGlobalSettings     = 'GLOBALSETTINGS';
  WSCDS_SectionConnection         = 'CONNECTION';
  WSCDS_SectionUpdateFrequency    = 'UPDATEFREQUENCY';
  WSCDS_EndUrlEntries             = 'entries';
  WSCDS_EndUrlUploadBytes         = 'importEntries/upload/bytes';
  WSCDS_EndUrlImportEntries       = 'importEntries';
  WSCDS_EndUrlImportEntriesEnd    = 'importEntries/end';
  WSCDS_EndUrlImportEntriesReport = 'importEntries/Report';
  WSCDS_XmlNone                   = 'None';
  WSCDS_XmlTrue                   = 'true';
  WSCDS_XmlFalse                  = 'false';
  WSCDS_ErrorMsg                  = '##### ERREUR';
  WSCDS_DebugMsg                  = 'DEBUG : ';
  WSCDS_ServiceName               = 'SvcSynBTPY2';

implementation

uses
  CommonTools
  , SvcMgr
  , Windows
  , SysUtils
  ;
  
{ TGetFromDSType }

class function TGetFromDSType.dstPrefix(DSType: T_WSDataService): string;
begin
  case DSType of
    wsdsThird             : Result := 'T';
    wsdsAnalyticalSection : Result := 'S';
    wsdsAccount           : Result := 'G';
    wsdsJournal           : Result := 'J';
    wsdsBankId            : Result := 'R';
    wsdsChoixCod          : Result := 'CC';
    wsdsCommon            : Result := 'CO';
    wsdsChoixExt          : Result := 'YX';
    wsdsRecovery          : Result := 'RR';
    wsdsCountry           : Result := 'PY';
    wsdsCurrency          : Result := 'D';
    wsdsCorrespondence    : Result := 'CR';
    wsdsPaymenChoice      : Result := 'MR';
    wsdsChangeRate        : Result := 'H';
    wsdsFiscalYear        : Result := 'EX';
    wsdsSocietyParameters : Result := 'SOC';
    wsdsEstablishment     : Result := 'ET';
    wsdsPaymentMode       : Result := 'MP';
    wsdsZipCode           : Result := 'O';
    wsdsContact           : Result := 'C';
    wsdsAccForPayment     : Result := 'E';
    wsdsDocPayment        : Result := 'GAC';
  else
    Result := '';
  end;
end;

class function TGetFromDSType.dstTableName(DSType: T_WSDataService): string;
begin
  case DSType of
    wsdsThird             : Result := Tools.GetTableNameFromTtn(ttnTiers);
    wsdsAnalyticalSection : Result := Tools.GetTableNameFromTtn(ttnSection);
    wsdsAccount           : Result := Tools.GetTableNameFromTtn(ttnGeneraux);
    wsdsJournal           : Result := Tools.GetTableNameFromTtn(ttnJournal);
    wsdsBankId            : Result := Tools.GetTableNameFromTtn(ttnRib);
    wsdsChoixCod          : Result := Tools.GetTableNameFromTtn(ttnChoixCod);
    wsdsCommon            : Result := Tools.GetTableNameFromTtn(ttnCommun);
    wsdsChoixExt          : Result := Tools.GetTableNameFromTtn(ttnChoixExt);
    wsdsRecovery          : Result := Tools.GetTableNameFromTtn(ttnRelance);
    wsdsCountry           : Result := Tools.GetTableNameFromTtn(ttnPays);
    wsdsCurrency          : Result := Tools.GetTableNameFromTtn(ttnDevise);
    wsdsCorrespondence    : Result := Tools.GetTableNameFromTtn(ttnCorresp);
    wsdsPaymenChoice      : Result := Tools.GetTableNameFromTtn(ttnModeRegl);
    wsdsChangeRate        : Result := Tools.GetTableNameFromTtn(ttnChancell);
    wsdsFiscalYear        : Result := Tools.GetTableNameFromTtn(ttnExercice);
    wsdsSocietyParameters : Result := Tools.GetTableNameFromTtn(ttnParamSoc);
    wsdsEstablishment     : Result := Tools.GetTableNameFromTtn(ttnEtabliss);
    wsdsPaymentMode       : Result := Tools.GetTableNameFromTtn(ttnModePaie);
    wsdsZipCode           : Result := Tools.GetTableNameFromTtn(ttnCodePostaux);
    wsdsContact           : Result := Tools.GetTableNameFromTtn(ttnContact);
    wsdsAccForPayment     : Result := Tools.GetTableNameFromTtn(ttnEcriture);
    wsdsDocPayment        : Result := Tools.GetTableNameFromTtn(ttnAcomptes);
  else
    Result := '';
  end;
end;

class function TGetFromDSType.dstViewName(DSType: T_WSDataService): string;
begin
  case DSType of
    wsdsAccount           : Result := 'ZCDSBACCOUNT';
    wsdsAnalyticalSection : Result := 'ZCDSBANALYTICALSECTION';
    wsdsBankId            : Result := 'ZCDSBBANKID';
    wsdsChangeRate        : Result := 'ZCDSBCHANGERATE';
    wsdsChoixCod          : Result := 'ZCDSBCHXOIXCOD';
    wsdsChoixExt          : Result := 'ZCDSBCHOIXEXT';
    wsdsCommon            : Result := 'ZCDSBCOMMUN';
    wsdsContact           : Result := 'ZCDSBCONTACT';
    wsdsCorrespondence    : Result := 'ZCDSBCORRESP';
    wsdsCountry           : Result := 'ZCDSBCOUNTRY';
    wsdsCurrency          : Result := 'ZCDSBCURRENCY';
    wsdsAccForPayment     : Result := 'ZCDSBECRPOURRGT';
    wsdsEstablishment     : Result := 'ZCDSBESTABLISHMENT';
    wsdsFieldsList        : Result := 'ZCDSBFIELDSLIST';
    wsdsFiscalYear        : Result := 'ZCDSBFISCALYEAR';
    wsdsJournal           : Result := 'ZCDSBJOURNAL';
    wsdsPaymenChoice      : Result := 'ZCDSBPAYMENT';
    wsdsPaymentMode       : Result := 'ZCDSBPAYMENTMODE';
    wsdsRecovery          : Result := 'ZCDSBRECOVERY';
    wsdsSocietyParameters : Result := 'ZCDSBSOCIETYPARAM';
    wsdsThird             : Result := 'ZCDSBTHIRDSLIST';
    wsdsZipCode           : Result := 'ZCDSBZIPCODE';
  else
    Result := '';
  end;
end;

class function TGetFromDSType.dstWSName(DSType: T_WSDataService): string;
begin
  case DSType of
    wsdsAccForPayment     : Result := 'lseACCFORPAYMENT';
    wsdsAccount           : Result := 'lseACCOUNTLIST';
    wsdsAnalyticalSection : Result := 'lseANALYTICALSECTION';
    wsdsBankId            : Result := 'lseBANKILIST';
    wsdsChangeRate        : Result := 'lseCHANGERATE';
    wsdsContact           : Result := 'lseCONTACT';
    wsdsCorrespondence    : Result := 'lseCORRESPONDENCELIST';
    wsdsCountry           : Result := 'lseCOUNTRYLIST';
    wsdsCurrency          : Result := 'lseCURRENCYLIST';
    wsdsEstablishment     : Result := 'lseESTABLISHMENT';
    wsdsFieldsList        : Result := 'lseFIELDSLIST';
    wsdsFiscalYear        : Result := 'lseFISCALYEAR';
    wsdsJournal           : Result := 'lseJOURNALLIST';
    wsdsChoixCod          : Result := 'lsePARAMCC';
    wsdsCommon            : Result := 'lsePARAMCO';
    wsdsChoixExt          : Result := 'lsePARAMYX';
    wsdsPaymenChoice      : Result := 'lsePAYMENTLIST';
    wsdsPaymentMode       : Result := 'lsePAYMENTMODE';
    wsdsRecovery          : Result := 'lseRECOVERYLIST';
    wsdsSocietyParameters : Result := 'lseSOCIETYPARAM';
    wsdsThird             : Result := 'lseTHIRDSLIST';
    wsdsZipCode           : Result := 'lseZIPCODE';
  else
    Result := '';
  end;
end;

class function TGetFromDSType.ExtractType(TableName : string) : string;
begin
  case Tools.GetTtnFromTableName(TableName) of
   ttnChoixCod : Result := ExtractType(wsdsChoixCod);
   ttnChoixExt : Result := ExtractType(wsdsChoixExt);
   ttnCommun   : Result := ExtractType(wsdsCommon);
   ttnContact  : Result := ExtractType(wsdsContact);
   ttnRelance  : Result := ExtractType(wsdsRecovery);
   ttnTiers    : Result := ExtractType(wsdsThird);
  else ;
    Result := '';
  end;
end;

class function TGetFromDSType.ExtractType(DSType : T_WSDataService) : string;
begin
  case DSType of
    wsdsChoixCod : Result := 'GDM;NVR;GOR;JUR;LGU;GCT;RTV;GZC;SCC;TRC;FON;SRV;YTC;CIV;LIP';
    wsdsChoixExt : Result := 'LB1;LB2;LB3;LB4;LB5;LB6;LB7;LB8;LB9;LBA';
    wsdsCommon   : Result := 'SEP;MVI;NAE;TVS;YFJ;SEN;CMP;TLC;APA;ARR;NTT;PSX';
    wsdsContact  : Result := 'CLI';
    wsdsRecovery : Result := 'RRG;RTR';
    wsdsThird    : Result := 'CLI;PRO';
  else
    Result := '';
  end;
end;

class procedure TGetFromDSType.WriteLog(TypeDebug: T_SvcSyncBTPY2Log; Text: string; LogValues : T_WSLogValues; LineLevel: integer; WithoutDateTime: Boolean = true);
var
  LogText     : string;
  LogFilePath : string;
  WindowsLog  : TEventLogger;
  LogFile     : TextFile;

  procedure WriteWindowsLog;
  begin
    LogFilePath := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), WSCDS_ServiceName, 'exe']);
    WindowsLog := TEventLogger.Create(ExtractFileName(LogFilePath));
    try
      WindowsLog.LogMessage(Text, EVENTLOG_INFORMATION_TYPE);
    finally
      WindowsLog.Free;
    end;
  end;

begin
  case TypeDebug of
    ssbylLog:
      begin
        if LogValues.LogLevel > 0 then
        begin
          LogFilePath := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), WSCDS_ServiceName, 'log']);
          if Logvalues.OneLogPerDay then
            LogFilePath := Format('%s_%s.log', [Copy(LogFilePath, 1, pos('.log', LogFilePath) -1), Tools.CastDateTimeForQry(Now)]);
          AssignFile(LogFile, LogFilePath);
          try
            if FileExists(LogFilePath) then
              Append(LogFile)
            else
              Rewrite(LogFile);
            if Text <> '' then
            begin
              if WithoutDateTime then
                LogText := Format('%s : %s%s', [DateTimeToStr(Now), StringOfChar(' ', LineLevel), Text])
              else
                LogText := Format('%s : %s%s', [Copy(Text, 1, pos('=', Text) - 1), StringOfChar(' ', LineLevel), Copy(Text, Pos('=', Text) + 1, length(Text))]);
            end else
              LogText := '';
            Writeln(LogFile, LogText);
          finally
            CloseFile(LogFile);
          end;
        end else
        begin
          { Si pas de log métier, on écrit dans le log windows uniquement pour le débug } 
          if (LogValues.DebugEvents > 0) and (pos(WSCDS_DebugMsg, Text) > 0 ) then
            WriteWindowsLog;
        end;
      end;
    ssbylWindows: WriteWindowsLog;
  end;
end;

class function TGetFromDSType.dstFiedsList(DSType: T_WSDataService): string;
var
  Prefix : string;
begin
  { ***********************
  LES CHAMPS DOIVENT IMPERATIVEMENT ETRE EN ORDRE ALPHABETIQUE
   *********************** }
  case DSType of
    wsdsThird :
      Result := 'T_ABREGE=Abrege'
              + ';T_ADRESSE1=Adresse1'
              + ';T_ADRESSE2=Adresse2'
              + ';T_ADRESSE3=Adresse3'
              + ';T_ANNEENAISSANCE=AnneeNaissance'
              + ';T_APE=CodeNAF'
              + ';T_APPORTEUR=Apporteur'
              + ';T_AUXILIAIRE=Auxiliaire'
              + ';T_AVOIRRBT=RbtSurAvoir'
              + ';T_CLETELEPHONE=TelephoneFormate'
              + ';T_CODEIMPORT=CodeImport'
              + ';T_CODEPOSTAL=CodePostal'
              + ';T_COEFCOMMA=CommissionApporteur'
              + ';T_COLLECTIF=Collectif'
              + ';T_COMMENTAIRE=Commentaire'
              + ';T_COMPTATIERS=FamilleComptable'
              + ';T_CONFIDENTIEL=Confidentiel'
              + ';T_CONSO=CodeConsolidation'
              + ';T_CORRESP1=Corresp1'
              + ';T_CORRESP2=Corresp2'
              + ';T_COUTHORAIRE=CoutHoraire'
              + ';T_CREDITACCORDE=CreditAccord'
              + ';T_CREDITDEMANDE=CreditDemande'
              + ';T_CREDITDERNMVT=CreditDernierMvt'
              + ';T_CREDITPLAFOND=CreditPlafond'
              + ';T_CREERPAR=CreerPar'
              + ';T_DATECREATION=DateCreation'
              + ';T_DATECREDITDEB=DateDebutAssuranceCredit'
              + ';T_DATECREDITFIN=DateFinAssuranceCredit'
              + ';T_DATEDERNMVT=DateDernierMvt'
              + ';T_DATEDERNPIECE=DateDernierePiece'
              + ';T_DATEDERNRELEVE=DateDernierReleve'
              + ';T_DATEFERMETURE=DateFermeture'
              + ';T_DATEINTEGR=DateIntegration'
              + ';T_DATEMODIF=DateModif'
              + ';T_DATEOUVERTURE=DateOuverture'
              + ';T_DATEPLAFONDDEB=DateDebutPlafondAutorise'
              + ';T_DATEPLAFONDFIN=DateFinPlafondAutorise'
              + ';T_DATEPROCLI=DateClientDepuis'
              + ';T_DEBITDERNMVT=DebitDernierMvt'
              + ';T_DEBRAYEPAYEUR=DebrayageAutomatismeTP'
              + ';T_DELAIMOYEN=DelaiMoyenLivraison'
              + ';T_DERNLETTRAGE=DernierLettrage'
              + ';T_DEVISE=Devise'
              + ';T_DIVTERRIT=DivisionTerritoriale'
              + ';T_DOMAINE=DomaineActivite'
              + ';T_DOSSIERCREDIT=NumDossierAssuranceCredit'
              + ';T_EAN=CodeEAN'
              + ';T_EMAIL=AdresseMessagerie'
              + ';T_EMAILING=eMailing'
              + ';T_ENSEIGNE=Enseigne'
              + ';T_ESCOMPTE=Escompte'
              + ';T_ETATRISQUE=EtatRisque'
              + ';T_EURODEFAUT=Euro'
              + ';T_EXPORTE=TiersExportePar'
              + ';T_FACTURE=Facture'
              + ';T_FACTUREHT=FactureHT'
              + ';T_FAX=Telephone2'
              + ';T_FERME=Ferme'
              + ';T_FORMEJURIDIQUE=FormeJuridique'
              + ';T_FRANCO=FrancoPort'
              + ';T_FREQRELEVE=FrequenceReleve'
              + ';T_INVISIBLE=Invisible'
              + ';T_ISPAYEUR=EstPayeur'
              + ';T_JOURNAISSANCE=JourNaissance'
              + ';T_JOURPAIEMENT1=JourPaiement1'
              + ';T_JOURPAIEMENT2=JourPaiement2'
              + ';T_JOURRELEVE=JourReleve'
              + ';T_JURIDIQUE=AbreviationPostale'
              + ';T_LANGUE=Langue'
              + ';T_LETTRABLE=Lettrable'
              + ';T_LETTREPAIEMENT=ModeleLettrePaiement'
              + ';T_LIBELLE=Libelle'
              + ';T_LIGNEDERNMVT=NumLigneDernierMvt'
              + ';T_LOCALTAX=LocalisationTaxe'
              + ';T_MODEREGLE=ModeReglement'
              + ';T_MOISCLOTURE=MoisClotureEntreprise'
              + ';T_MOISNAISSANCE=MoisNaissance'
              + ';T_MOTIFVIREMENT=MotifVirement'
              + ';T_MULTIDEVISE=MultiDevise'
              + ';T_NATIONALITE=Nationalite'
              + ';T_NATUREAUXI=Nature'
              + ';T_NATUREECONOMIQUE=NatureEconomique'
              + ';T_NIF=CodeNIF'
              + ';T_NIVEAUIMPORTANCE=ImportanceClient'
              + ';T_NIVEAURISQUE=NiveauRisque'
              + ';T_NUMDERNMVT=NumPieceDernierMvt'
              + ';T_NUMDERNPIECE=NumDernierePiece'
              + ';T_ORIGINETIERS=OrigineTiers'
              + ';T_PARTICULIER=Particulier'
              + ';T_PASSWINTERNET=MotPasseInternet'
              + ';T_PAYEUR=Payeur'
              + ';T_PAYEURECLATEMENT=PayeurEclatement'
              + ';T_PAYS=Pays'
              + ';T_PHONETIQUE=LibellePhonetique'
              + ';T_PRENOM=Libelle2'
              + ';T_PRESCRIPTEUR=Prescripteur'
              + ';T_PROFIL=ProfilGestion'
              + ';T_PUBLIPOSTAGE=Publipostage'
              + ';T_QUALIFESCOMPTE=ModeApplicationEscompte'
              + ';T_REGIMETVA=RegimeTaxe'
              + ';T_REGION=Region'
              + ';T_RELANCEREGLEMENT=ModeRelanceReglement'
              + ';T_RELANCETRAITE=ModeRelanceTraite'
              + ';T_RELEVEFACTURE=ReleveSurFacture'
              + ';T_REMISE=Remise'
              + ';T_REPRESENTANT=Commercial'
              + ';T_RESIDENTETRANGER=ResidentEtranger'
              + ';T_RVA=SiteWeb'
              + ';T_SAUTPAGE=SautPage'
              + ';T_SCORECLIENT=ScoreClient'
              + ';T_SCORERELANCE=ScoreRelance'
              + ';T_SECTEUR=SecteurActivite'
              + ';T_SEXE=Sexe'
              + ';T_SIRET=Siret'
              + ';T_SOCIETE=CodeSociete'
              + ';T_SOCIETEGROUPE=Groupe'
              + ';T_SOLDEPROGRESSIF=SoldeProgressif'
              + ';T_SOUMISTPF=SoumisTPF'
              + ';T_TABLE0=TableLibre1'
              + ';T_TABLE1=TableLibre2'
              + ';T_TABLE2=TableLibre3'
              + ';T_TABLE3=TableLibre4'
              + ';T_TABLE4=TableLibre5'
              + ';T_TABLE5=TableLibre6'
              + ';T_TABLE6=TableLibre7'
              + ';T_TABLE7=TableLibre8'
              + ';T_TABLE8=TableLibre9'
              + ';T_TABLE9=TableLibre10'
              + ';T_TARIFTIERS=FamilleTarif'
              + ';T_TELEPHONE=Telephone'
              + ';T_TELEPHONE2=Telephone3'
              + ';T_TELEX=TelephonePortable'
              + ';T_TIERS=Code'
              + ';T_TOTALCREDIT=TotalCredit'
              + ';T_TOTALDEBIT=TotalDebit'
              + ';T_TOTAUXMENSUELS=TotauxMensuels'
              + ';T_TOTCREANO=TotalCreditANouveaux'
              + ';T_TOTCREANON1=ANouveauProvisoireCreditNP1'
              + ';T_TOTCREE=TotalCreditN'
              + ';T_TOTCREP=TotalCreditNM1'
              + ';T_TOTCRES=TotalCreditNP1'
              + ';T_TOTDEBANO=TotalDebbitANouveaux'
              + ';T_TOTDEBANON1=ANouveauProvisoireDebitNP1'
              + ';T_TOTDEBE=TotalDebitN'
              + ';T_TOTDEBP=TotalDebitNM1'
              + ';T_TOTDEBS=TotalDebitNP1'
              + ';T_TOTDERNPIECE=TotalHTDernierePiece'
              + ';T_TRANSPORTEUR=Transporteur'
              + ';T_TVAENCAISSEMENT=TVAEncaissement'
              + ';T_UTILISATEUR=Utilisateur'
              + ';T_VILLE=Ville'
              + ';T_ZONECOM=ZoneCommerciale'
              ;
    wsdsAnalyticalSection :
      Result := 'CSP_ABREGE=Abrege'
              + ';CSP_AFFAIREENCOUR=AffaireCours'
              + ';CSP_AXE=Axe'
              + ';CSP_BUDSECT=SectionBudgetaireDepassement'
              + ';CSP_CHANTIER=Chantier'
              + ';CSP_CLEREPARTITIO=CleRepartAutresSections'
              + ';CSP_CODEIMPORT=CodeSectionAutreAppli'
              + ';CSP_CONFIDENTIEL=Confidentielle'
              + ';CSP_CORRESP1=SectionCorresp1'
              + ';CSP_CORRESP2=SectionCorresp2'
              + ';CSP_CREDITDERNMVT=CreditDernierMvt'
              + ';CSP_CREERPAR=CreeePar'
              + ';CSP_DATECREATION=DateCreation'
              + ';CSP_DATEDERNMVT=DateDernierMvt'
              + ';CSP_DATEFERMETURE=DateDerniereFermeture'
              + ';CSP_DATEMODIF=DateModif'
              + ';CSP_DATEOUVERTURE=DateDerniereOuverture'
              + ';CSP_DEBCHANTIER=DateDebutChantier'
              + ';CSP_DEBITDERNMVT=DebitDernierMvt'
              + ';CSP_DOMAINE=DomaineActivite'
              + ';CSP_EXPORTE=ExporteePar'
              + ';CSP_FERME=Fermee'
              + ';CSP_FINCHANTIER=DateFinChantier'
              + ';CSP_INDIRECTE=Indirecte'
              + ';CSP_INVISIBLE=Invisible'
              + ';CSP_LIBELLE=Libelle'
              + ';CSP_LIGNEDERNMVT=NumLigneDernierMvt'
              + ';CSP_MAITREOEUVRE=MaitreOeuvre'
              + ';CSP_MODELE=Modele'
              + ';CSP_NUMDERNMVT=NumPieceDernierMvt'
              + ';CSP_REPARTAVECCPT=RepartAvecCCPT'
              + ';CSP_SAUTPAGE=SautPage'
              + ';CSP_SECTION=Code'
              + ';CSP_SECTIONTRIE=SectionRi'
              + ';CSP_SECTIONTRIE0=SectionRi0'
              + ';CSP_SECTIONTRIE1=SectionRi1'
              + ';CSP_SECTIONTRIE2=SectionRi2'
              + ';CSP_SECTIONTRIE3=SectionRi3'
              + ';CSP_SECTIONTRIE4=SectionRi4'
              + ';CSP_SECTIONTRIE5=SectionRi5'
              + ';CSP_SECTIONTRIE6=SectionRi6'
              + ';CSP_SECTIONTRIE7=SectionRi7'
              + ';CSP_SECTIONTRIE8=SectionRi8'
              + ';CSP_SECTIONTRIE9=SectionRi9'
              + ';CSP_SENS=Sens'
              + ';CSP_SOCIETE=Societe'
              + ';CSP_SOLDEPROGRESS=SoldeProgressif'
              + ';CSP_SOUSPLAN=SousPlan'
              + ';CSP_TABLE0=TableLibre1'
              + ';CSP_TABLE1=TableLibre2'
              + ';CSP_TABLE2=TableLibre3'
              + ';CSP_TABLE3=TableLibre4'
              + ';CSP_TABLE4=TableLibre5'
              + ';CSP_TABLE5=TableLibre6'
              + ';CSP_TABLE6=TableLibre7'
              + ';CSP_TABLE7=TableLibre8'
              + ';CSP_TABLE8=TableLibre9'
              + ';CSP_TABLE9=TableLibre10'
              + ';CSP_TOTALCREDIT=TotalCredit'
              + ';CSP_TOTALDEBIT=TotalDebit'
              + ';CSP_TOTAUXMENSUEL=TotauxMensuels'
              + ';CSP_TOTCREANO=TotalCreditAnouveaux'
              + ';CSP_TOTCREANON1=ANouveauProvisoireCreditNP1'
              + ';CSP_TOTCREE=TotalCreditN'
              + ';CSP_TOTCREP=TotalCreditNM1'
              + ';CSP_TOTCRES=TotalCreditNP1'
              + ';CSP_TOTDEBANO=TotalDebitANouveaux'
              + ';CSP_TOTDEBANON1=ANouveauProvisoireDebitNP1'
              + ';CSP_TOTDEBE=TotalDebitN'
              + ';CSP_TOTDEBP=TotalDebitNM1'
              + ';CSP_TOTDEBS=TotalDebitNP1'
              + ';CSP_TRANCHEGENEA=TrancheCompteA'
              + ';CSP_TRANCHEGENEDE=TrancheCompteDe'
              + ';CSP_UO=UniteOeuvre'
              + ';CSP_UOLIBELLE=LibelleUniteOeuvre'
              + ';CSP_UTILISATEUR=Utilisateur'
              ;
    wsdsAccount :
      Result := 'G_ABREGE=Abrege'
              + ';G_ADRESSE1=Adresse1'
              + ';G_ADRESSE2=Adresse2'
              + ';G_ADRESSE3=Adresse3'
              + ';G_APE=CodeNAF'
              + ';G_BUDGENE=CompteBudgetaireDepassement'
              + ';G_CENTRALISABLE=Centralisable'
              + ';G_CODEIMPORT=CompteApplicationAmont'
              + ';G_CODEPOSTAL=CodePostal'
              + ';G_COLLECTIF=Collectif'
              + ';G_COMPENS=GestionCompensation'
              + ';G_CONFIDENTIEL=Confidentiel'
              + ';G_CONSO=CodeConsolidation'
              + ';G_CORRESP1=CompteCorresp1'
              + ';G_CORRESP2=CompteCorresp2'
              + ';G_CREDITDERNMVT=CreditDernierMvt'
              + ';G_CREDNONPOINTE=TotalCreditNonPointeNM1'
              + ';G_CREERPAR=CreerPar'
              + ';G_CUTOFF=EstCompteChargePeriodique'
              + ';G_CUTOFFCOMPTE=CompteChargePeriodique'
              + ';G_CUTOFFECHUE=Echue'
              + ';G_CUTOFFPERIODE=Periode'
              + ';G_CYCLEREVISION=CycleRevision'
              + ';G_DATECREATION=DateCreation'
              + ';G_DATEDERNMVT=DateDernierMvt'
              + ';G_DATEFERMETURE=DateFermeture'
              + ';G_DATEMODIF=DateModif'
              + ';G_DATEOUVERTURE=DateOuverture'
              + ';G_DEBITDERNMVT=DebitDernierMvt'
              + ';G_DEBNONPOINTE=TotalDebitNonPointeNM1'
              + ';G_DERNLETTRAGE=DernierCodeLettrage'
              + ';G_DEVISECAISSE=CaisseEnDevise'
              + ';G_DIVTERRIT=DivisionTerritoriale'
              + ';G_DOMAINE=DomaineActivite'
              + ';G_EFFET=CompteEffetPortefeuille'
              + ';G_ETABLISSEMENT=Etablissement'
              + ';G_EXPORTE=ExportePar'
              + ';G_FAX=Fax'
              + ';G_FERME=Ferme'
              + ';G_GENERAL=Code'
              + ';G_GUIDASSOCIER=CodePersonneAssociee'
              + ';G_IAS14=GestionIAS14'
              + ';G_INVISIBLE=Invisible'
              + ';G_JOURPAIEMENT1=JourReglement1'
              + ';G_JOURPAIEMENT2=JourReglement2'
              + ';G_JURIDIQUE=FormeJuridique'
              + ';G_LANGUE=Langue'
              + ';G_LETTRABLE=Lettrable'
              + ';G_LETTREPAIEMENT=ModeleLettrePaiement'
              + ';G_LIBELLE=Libelle'
              + ';G_LIGNEDERNMVT=NumLigneDernierMvt'
              + ';G_MODELE=Modele'
              + ';G_MODEREGLE=ConditionsReglement'
              + ';G_MOTIFVIREMENT=MotifVirement'
              + ';G_NATUREECONOMIQUE=NatureEconomique'
              + ';G_NATUREGENE=Nature'
              + ';G_NIF=CodeNIF'
              + ';G_NONTAXABLE=CompteNonTaxable'
              + ';G_NUMDERNMVT=NumPieceDernierMvt'
              + ';G_PAYS=Pays'
              + ';G_PLAFOND=PlafondRisque'
              + ';G_POINTABLE=Pointable'
              + ';G_PURGEABLE=Purgeable'
              + ';G_QUALIFQTE1=QualifiantQte1'
              + ';G_QUALIFQTE2=QualifiantQte2'
              + ';G_REGIMETVA=RegimeTVA'
              + ';G_RELANCEREGLEMENT=ModeleRelanceRgt'
              + ';G_RELANCETRAITE=ModeleRelanceTraites'
              + ';G_RESIDENTETRANGER=CodeResidentEtranger'
              + ';G_RESTRICTIONA1=ModeleRrestrictionA1'
              + ';G_RESTRICTIONA2=ModeleRrestrictionA2'
              + ';G_RESTRICTIONA3=ModeleRrestrictionA3'
              + ';G_RESTRICTIONA4=ModeleRrestrictionA4'
              + ';G_RESTRICTIONA5=ModeleRrestrictionA5'
              + ';G_RISQUE=Risque'
              + ';G_RISQUETIERS=CompteParticipantRisque'
              + ';G_SAUTPAGE=SautPage'
              + ';G_SENS=Sens'
              + ';G_SIRET=CodeSIRET'
              + ';G_SOCIETE=Societe'
              + ';G_SOLDEPROGRESSIF=SoldeProgressif'
              + ';G_SOUMISTPF=SoumisTPF'
              + ';G_SUIVITRESO=SuiviTreso'
              + ';G_TABLE0=TableLibre1'
              + ';G_TABLE1=TableLibre2'
              + ';G_TABLE2=TableLibre3'
              + ';G_TABLE3=TableLibre4'
              + ';G_TABLE4=TableLibre5'
              + ';G_TABLE5=TableLibre6'
              + ';G_TABLE6=TableLibre7'
              + ';G_TABLE7=TableLibre8'
              + ';G_TABLE8=TableLibre9'
              + ';G_TABLE9=TableLibre10'
              + ';G_TELEPHONE=Telephone1'
              + ';G_TELEX=Telephone2'
              + ';G_TOTALCREDIT=TotalCredit'
              + ';G_TOTALDEBIT=TotalDebit'
              + ';G_TOTAUXMENSUELS=TotauxMois'
              + ';G_TOTCREANO=TotalCreditANouveaux'
              + ';G_TOTCREANON1=ANouveauProvisoireNM1'
              + ';G_TOTCREE=TotalCreditN'
              + ';G_TOTCREN2=TotalCreditN2'
              + ';G_TOTCREP=TotalCreditNM1'
              + ';G_TOTCREPTD=TotalCreditPointeDevise'
              + ';G_TOTCREPTP=TotalCreditPointePivot'
              + ';G_TOTCRES=TotalCreditNP1'
              + ';G_TOTDEBANO=TotalDebitANouveaux'
              + ';G_TOTDEBANON1=ANouveauProvisoireNP1'
              + ';G_TOTDEBE=TotalDebitN'
              + ';G_TOTDEBN2=TotalDebitN2'
              + ';G_TOTDEBP=TotalDebitNM1'
              + ';G_TOTDEBPTD=TotalDebitPointeDevise'
              + ';G_TOTDEBPTP=TotalDebitPointePivot'
              + ';G_TOTDEBS=TotalDebitNP1'
              + ';G_TPF=CodeTPF'
              + ';G_TVA=CodeTVA'
              + ';G_TVAENCAISSEMENT=ExigibiliteTVA'
              + ';G_TVASURENCAISS=SoumisTVASurEncaissement'
              + ';G_TYPECPTTVA=TypeTVA'
              + ';G_UTILISATEUR=Utilisateur'
              + ';G_VENTILABLE=Ventilable'
              + ';G_VENTILABLE1=VentilableAxe1'
              + ';G_VENTILABLE2=VentilableAxe2'
              + ';G_VENTILABLE3=VentilableAxe3'
              + ';G_VENTILABLE4=VentilableAxe4'
              + ';G_VENTILABLE5=VentilableAxe5'
              + ';G_VENTILTYPE=VentilationType'
              + ';G_VILLE=Ville'
              + ';G_VISAREVISION=VisaSurCompte'
              ;
    wsdsJournal :
      Result := 'J_ABREGE=Abrege'
              + ';J_ACCELERATEUR=AccelerateurSaisie'
              + ';J_AXE=Axe'
              + ';J_CENTRALISABLE=Centralisable'
              + ';J_CHOIXDATE=ChoixDateEtebac'
              + ';J_CHRONOID=CodeCompteur'
              + ';J_COMPTEAUTOMAT=ListeCptesAutomatiques'
              + ';J_COMPTEINTERDIT=ListeCpteInterdits'
              + ';J_COMPTEURNORMAL=CompteurNormal'
              + ';J_COMPTEURSIMUL=CompteurSimulation'
              + ';J_CONTREPARTIE=Contrepartie'
              + ';J_CONTREPARTIEAUX=ContrepartieAuxiliaire'
              + ';J_CPTEREGULCREDIT1=CpteRegulCredit1'
              + ';J_CPTEREGULCREDIT2=CpteRegulCredit2'
              + ';J_CPTEREGULCREDIT3=CpteRegulCredit3'
              + ';J_CPTEREGULDEBIT1=CpteRegulDebit1'
              + ';J_CPTEREGULDEBIT2=CpteRegulDebit2'
              + ';J_CPTEREGULDEBIT3=CpteRegulDebit3'
              + ';J_CREDITDERNMVT=CreditDernierMvt'
              + ';J_CREERPAR=CreerPar'
              + ';J_DATECREATION=DateCreation'
              + ';J_DATEDERNMVT=DateDernierMvt'
              + ';J_DATEFERMETURE=DateFermeture'
              + ';J_DATEMODIF=DateModif'
              + ';J_DATEOUVERTURE=DateOuverture'
              + ';J_DEBITDERNMVT=DebitDernierMvt'
              + ';J_EFFET=JournalSuiviEffet'
              + ';J_EQAUTO=SoldeAutoF10'
              + ';J_EXPORTE=ExportePar'
              + ';J_FERME=Ferme'
              + ';J_IMPORTCONFORME=ImportConforme'
              + ';J_INCNUM=IncrementerNumFolio'
              + ';J_INCREF=IncrementerRefEcriture'
              + ';J_INVISIBLE=Invisible'
              + ';J_JOURNAL=Code'
              + ';J_LIBELLE=Libelle'
              + ';J_LIBELLEAUTO=LibelleAutomatique'
              + ';J_MODESAISIE=ModeSaisie'
              + ';J_MONTANTNEGATIF=MontantsNegatifsAutorises'
              + ';J_MULTIDEVISE=MultiDevise'
              + ';J_NATCOMPL=AfficherNature'
              + ';J_NATDEFAUT=NatureParDefaut'
              + ';J_NATUREJAL=Nature'
              + ';J_NUMDERNMVT=NumPieceDernierMvt'
              + ';J_NUMLIGJOUR=NumLigneJour'
              + ';J_OUVRIRLETT=LettrageSaisie'
              + ';J_SOCIETE=Societe'
              + ';J_TOTALCREDIT=TotalCredit'
              + ';J_TOTALDEBIT=TotalDebit'
              + ';J_TOTCREE=TotalCreditN'
              + ';J_TOTCREP=TotalCreditNM1'
              + ';J_TOTCRES=TotalCreditNP1'
              + ';J_TOTDEBE=TotalDebitN'
              + ';J_TOTDEBP=TotalDebitNM1'
              + ';J_TOTDEBS=TotalDebitNP1'
              + ';J_TRESOCHAINAGE=GenerationEcritures'
              + ';J_TRESODATE=ModificationDateAutorisee'
              + ';J_TRESOECRITURE=RechercheTableEcriture'
              + ';J_TRESOIMPORT=TraitementImportationEtebac'
              + ';J_TRESOLIBELLE=ModificationLibelleAutorisee'
              + ';J_TRESOLIBRE=SoldeSurCompteDeBanque'
              + ';J_TRESOMONTANT=ModificationMontantAutorisee'
              + ';J_TRESOVALID=GenerationEcritureComptable'
              + ';J_TVACTRL=ControleTVAPiece'
              + ';J_TYPECONTREPARTIE=ScenarioSaisieTreso'
              + ';J_UTILISATEUR=Utilisateur'
              + ';J_VALIDEEN=ValideN'
              + ';J_VALIDEEN1=ValideNP1'
              ;
    wsdsBankId :
      Result :=  'R_ACOMPTE=PourAcompte'
              + ';R_AUXILIAIRE=Auxiliaire'
              + ';R_CLERIB=CleRIB'
              + ';R_CODEBIC=CodeBic'
              + ';R_CODEIBAN=Iban'
              + ';R_CREATEUR=Createur'
              + ';R_DATECREATION=DateCreation'
              + ';R_DATEMODIF=DateModif'
              + ';R_DEVISE=Devise'
              + ';R_DOMICILIATION=Domiciliation'
              + ';R_ETABBQ=CodeBanque'
              + ';R_FRAISPROF=PourFrais'
              + ';R_GUICHET=CodeGuichet'
              + ';R_NATECO=NatureEconomique'
              + ';R_NUMEROCOMPTE=NumeroCompte'
              + ';R_NUMERORIB=Identifiant'
              + ';R_PAYS=Pays'
              + ';R_PRINCIPAL=Principal'
              + ';R_SALAIRE=PourSalarie'
              + ';R_SOCIETE=Societe'
              + ';R_TYPEIDBQ=TypeIdentifiantBancaire'
              + ';R_TYPEPAYS=TypeIdentification'
              + ';R_UTILISATEUR=Utilisateur'
              + ';R_VILLE=Ville'
              ;                             
    wsdsChoixCod, wsdsCommon, wsdsChoixExt :
      begin
        Prefix := dstPrefix(DSType);
        Result := Prefix + '_ABREGE=Abrege'
                + ';' + Prefix + '_CODE=Code'
                + ';' + Prefix + '_LIBELLE=Libelle'
                + ';' + Prefix + '_LIBRE=Libre'
                + ';' + Prefix + '_TYPE=Type'
                ;
      end;
    wsdsRecovery :
      Result := 'RR_DELAI1=DelaiNiveau1'
              + ';RR_DELAI2=DelaiNiveau2'
              + ';RR_DELAI3=DelaiNiveau3'
              + ';RR_DELAI4=DelaiNiveau4'
              + ';RR_DELAI5=DelaiNiveau5'
              + ';RR_DELAI6=DelaiNiveau6'
              + ';RR_DELAI7=DelaiNiveau7'
              + ';RR_ENJOURS=RelanceNiveauJours'
              + ';RR_FAMILLERELANCE=FamilleRelance'
              + ';RR_GROUPELETTRE=RelancerPlusHautNiveau'
              + ';RR_INVISIBLE=Invisible'
              + ';RR_LIBELLE=Libelle'
              + ';RR_MODELE1=ModeleNiveau1'
              + ';RR_MODELE2=ModeleNiveau2'
              + ';RR_MODELE3=ModeleNiveau3'
              + ';RR_MODELE4=ModeleNiveau4'
              + ';RR_MODELE5=ModeleNiveau5'
              + ';RR_MODELE6=ModeleNiveau6'
              + ';RR_MODELE7=ModeleNiveau7'
              + ';RR_NONECHU=InclureNonEchus'
              + ';RR_SCOORING=AppliquerScoring'
              + ';RR_TYPERELANCE=TypeRelance'
              ;
    wsdsCountry :
      Result := 'PY_ABREGE=Abrege'
              + ';PY_CODEBANCAIRE=CodeBancaire'
              + ';PY_CODEDI=CodeEdi'
              + ';PY_CODEINSEE=CodeInsee'
              + ';PY_CODEISO2=CodeIso2'
              + ';PY_DEVISE=Devise'
              + ';PY_LANGUE=Langue'
              + ';PY_LIBELLE=Libelle'
              + ';PY_LIEUDISPO=Incoterm'
              + ';PY_LIMITROPHE=Limitrophe'
              + ';PY_MASQUENIF=MasqueSaisieNIF'
              + ';PY_MEMBRECEE=MembreCEE'
              + ';PY_NATIONALITE=Nationalite'
              + ';PY_PAYS=Code'
              + ';PY_REGION=Region'
              ;
    wsdsCurrency :
      Result := 'D_ARRONDIPRIXACHAT=ArrondiPrixAchat'
              + ';D_ARRONDIPRIXVENTE=ArrondiPrixVente'
              + ';D_CODEISO=CodeIso'
              + ';D_CPTLETTRCREDIT=CpteRegulationCredit'
              + ';D_CPTLETTRDEBIT=CpteRegulationDebit'
              + ';D_CPTPROVCREDIT=CpteGainChangeCredit'
              + ';D_CPTPROVDEBIT=CpteGainChangeDebit'
              + ';D_DECIMALE=NbreDecimales'
              + ';D_DEVISE=Code'
              + ';D_FERME=Ferme'
              + ';D_FONGIBLE=SubdivisionEuro'
              + ';D_LIBELLE=Libelle'
              + ';D_MAXCREDIT=PlafondCredit'
              + ';D_MAXDEBIT=PlafondDebit'
              + ';D_MONNAIEIN=EstMonnaieIn'
              + ';D_PARITEEURO=PariteEuro'
              + ';D_PARITEEUROFIXING=PartieEuroFixing'
              + ';D_QUOTITE=Quotite'
              + ';D_SOCIETE=Societe'
              + ';D_SYMBOLE=Symbole'
              ;
    wsdsChangeRate :
      Result := 'H_COMMENTAIRE=Commentaire'
              + ';H_COTATION=Cotation'
              + ';H_DATECOURS=DateTaux'
              + ';H_DEVISE=CodeDevise'
              + ';H_SOCIETE=Societe'
              + ';H_TAUXCLOTURE=TauxCloture'
              + ';H_TAUXLIBRE1=TauxLibre1'
              + ';H_TAUXLIBRE2=TauxLibre2'
              + ';H_TAUXMOYEN=TauxMoyen'
              + ';H_TAUXREEL=TauxReel'
              ;
    wsdsCorrespondence :
      Result := 'CR_ABREGE=Abrege'
              + ';CR_CORRESP=Corresponance'
              + ';CR_INVISIBLE=Invisible'
              + ';CR_LIBELLE=Libelle'
              + ';CR_LIBRETEXTE1=LibreTexte1'
              + ';CR_LIBRETEXTE2=LibreTexte2'
              + ';CR_LIBRETEXTE3=LibreTexte3'
              + ';CR_LIBRETEXTE4=LibreTexte4'
              + ';CR_LIBRETEXTE5=LibreTexte5'
              + ';CR_SOCIETE=Societe'
              + ';CR_TYPE=Type'
              ;
    wsdsPaymenChoice :
      Result := 'MR_ABREGE=Abrege'
              + ';MR_APARTIRDE=APartirDe'
              + ';MR_ARRONDIJOUR=JourArrondi'
              + ';MR_ECARTJOURS=EcartJours'
              + ';MR_EINTEGREAUTO=IntegrationEuto'
              + ';MR_ESC1=EscompteSurEch1'
              + ';MR_ESC10=EscompteSurEch10'
              + ';MR_ESC11=EscompteSurEch11'
              + ';MR_ESC12=eEscompteSurEch12'
              + ';MR_ESC2=EscompteSurEch2'
              + ';MR_ESC3=EscompteSurEch3'
              + ';MR_ESC4=EscompteSurEch4'
              + ';MR_ESC5=EscompteSurEch'
              + ';MR_ESC6=EscompteSurEch6'
              + ';MR_ESC7=EscompteSurEch7'
              + ';MR_ESC8=EscompteSurEch8'
              + ';MR_ESC9=EscompteSurEch9'
              + ';MR_INVISIBLE=Invisible'
              + ';MR_LIBELLE=Libelle'
              + ';MR_MODEGUIDE=ModeGuide'
              + ';MR_MODEREGLE=Code'
              + ';MR_MONTANTMIN=MinimumRequis'
              + ';MR_MP1=ModePaiementEch1'
              + ';MR_MP10=ModePaiementEch10'
              + ';MR_MP11=ModePaiementEch11'
              + ';MR_MP12=ModePaiementEch12'
              + ';MR_MP2=ModePaiementEch2'
              + ';MR_MP3=ModePaiementEch3'
              + ';MR_MP4=ModePaiementEch4'
              + ';MR_MP5=ModePaiementEch5'
              + ';MR_MP6=ModePaiementEch6'
              + ';MR_MP7=ModePaiementEch7'
              + ';MR_MP8=ModePaiementEch8'
              + ';MR_MP9=ModePaiementEch9'
              + ';MR_NOMBREECHEANCE=NbreEcheances'
              + ';MR_PLUSJOUR=JourPlus'
              + ';MR_REMPLACEMIN=ConditionsRemplacement'
              + ';MR_REPARTECHE=Repartition'
              + ';MR_SEPAREPAR=SepareesPar'
              + ';MR_SOCIETE=Societe'
              + ';MR_TAUX1=TauxSurEch1'
              + ';MR_TAUX10=TauxSurEch10'
              + ';MR_TAUX11=TauxSurEch11'
              + ';MR_TAUX12=TauxSurEch12'
              + ';MR_TAUX2=TauxSurEch2'
              + ';MR_TAUX3=TauxSurEch3'
              + ';MR_TAUX4=TauxSurEch4'
              + ';MR_TAUX5=TauxSurEch5'
              + ';MR_TAUX6=TauxSurEch6'
              + ';MR_TAUX7=TauxSurEch7'
              + ';MR_TAUX8=TauxSurEch8'
              + ';MR_TAUX9=TauxSurEch9'
              ;
    wsdsFiscalYear :
      Result := 'EX_ABREGE=Abrege'
              + ';EX_BUDJAL=JournalBudgetDepassement'
              + ';EX_DATECLOTDE=DateClotureDefinitive'
              + ';EX_DATECUM=DateButoirCumul'
              + ';EX_DATECUMBUD=DateButoirCumulBudget'
              + ';EX_DATECUMBUDGET=DateButoirCumulBudgetee'
              + ';EX_DATECUMRUB=DateBubtoirCumul'
              + ';EX_DATEDEBUT=DateDebut'
              + ';EX_DATEFIN=DateFin'
              + ';EX_ENTITY=Entite'
              + ';EX_ETATADV=EtatAdminVente'
              + ';EX_ETATAPPRO=EtatApprovisionnement'
              + ';EX_ETATBUDGET=EtatBudgetaire'
              + ';EX_ETATCPTA=EtatComptable'
              + ';EX_ETATPROD=EtatProduction'
              + ';EX_EXERCICE=Code'
              + ';EX_LIBELLE=Libelle'
              + ';EX_NATEXO=NatureExercice'
              + ';EX_NONSOUMISBOI=NonSoumisBOI'
              + ';EX_PASEQUILIBRE=PeriodesNonValidees'
              + ';EX_SOCIETE=Societe'
              + ';EX_VALIDEE=Valide'
              ;
    wsdsSocietyParameters  :
      Result := 'SOC_DATA=Donnee'
              + ';SOC_NOM=Nom'
              ;
    wsdsEstablishment :
      Result := 'ET_ABREGE=Abrege'
              + ';ET_ACTIVITE=Activite'
              + ';ET_ADRESSE1=Adresse1'
              + ';ET_ADRESSE2=Adresse2'
              + ';ET_ADRESSE3=Adresse3'
              + ';ET_ALBAT=Batiment'
              + ';ET_ALESC=Escalier'
              + ';ET_ALETA=Etage'
              + ';ET_ALNOAPP=NumeroAppartement'
              + ';ET_ALRESID=Residence'
              + ';ET_APE=Naf'
              + ';ET_AXE1=Groupe'
              + ';ET_AXE2=SousGroupe'
              + ';ET_BOOLLIBRE1=DecisionLibre1'
              + ';ET_BOOLLIBRE2=DecisionLibre2'
              + ';ET_BOOLLIBRE3=DecisionLibre3'
              + ';ET_CHARLIBRE1=TexteLibre1'
              + ';ET_CHARLIBRE2=TexteLibre2'
              + ';ET_CHARLIBRE3=TexteLibre3'
              + ';ET_CODEEDI=CodeEDI'
              + ';ET_CODEPOSTAL=CodePostal'
              + ';ET_DATECREATION=DateCreation'
              + ';ET_DATELIBRE1=DateLibre1'
              + ';ET_DATELIBRE2=DateLibre2'
              + ';ET_DATELIBRE3=DateLibre3'
              + ';ET_DATEMODIF=DateModif'
              + ';ET_DEPOT=DepotPrincipal'
              + ';ET_DEPOTLIE=ListeDepotsLies'
              + ';ET_DEVISE=Devise'
              + ';ET_DEVISEACH=DeviseTarifAchat'
              + ';ET_DIVTERRIT=DivsionTerritoriale'
              + ';ET_EAN=CodeEAN'
              + ';ET_EMAIL=ServeurWeb'
              + ';ET_ETABLIE=EtablissementLie'
              + ';ET_ETABLISSEMENT=Code'
              + ';ET_FAX=Fax'
              + ';ET_FICTIF=Fictif'
              + ';ET_INVISIBLE=Invisible'
              + ';ET_JURIDIQUE=FormeJuridique'
              + ';ET_LANGUE=Langue'
              + ';ET_LIBELLE=Libelle'
              + ';ET_LIBREET1=TableLibre1'
              + ';ET_LIBREET2=TableLibre2'
              + ';ET_LIBREET3=TableLibre3'
              + ';ET_LIBREET4=TableLibre4'
              + ';ET_LIBREET5=TableLibre5'
              + ';ET_LIBREET6=TableLibre6'
              + ';ET_LIBREET7=TableLibre7'
              + ';ET_LIBREET8=TableLibre8'
              + ';ET_LIBREET9=TableLibre9'
              + ';ET_LIBREETA=TableLibreA'
              + ';ET_LOCALTAX=LocalisationTaxe'
              + ';ET_NODOSSIER=NomDossier'
              + ';ET_PAYS=Pays'
              + ';ET_PROGRAMME=Programme'
              + ';ET_RESPONSABLE=Responsable'
              + ';ET_SIRET=Siret'
              + ';ET_SOCIETE=Societe'
              + ';ET_SURSITE=GereSurSite'
              + ';ET_SURSITEDISTANT=GereSurSiteDistant'
              + ';ET_TELEPHONE=Telephone1'
              + ';ET_TELEX=Telephone2'
              + ';ET_TYPETARIF=TarifUtilise'
              + ';ET_TYPETARIFACH=TarifAchat'
              + ';ET_UTILISATEUR=Utilisateur'
              + ';ET_VALLIBRE1=ValeurLibre1'
              + ';ET_VALLIBRE2=ValeurLibre2'
              + ';ET_VALLIBRE3=ValeurLibre3'
              + ';ET_VILLE=Ville'
              + ';ET_VOIENO=NumeroVoie'
              + ';ET_VOIENOCOMPL=ComplementNumeroVoie'
              + ';ET_VOIENOM=NomVoie'
              + ';ET_VOIETYPE=TypeVoie'
              ;
    wsdsPaymentMode :
      Result := 'MP_ABREGE=Abrege'
              + ';MP_AFFICHNUMCBUS=AffichageCarteFormatUS'
              + ';MP_ARRONDIFO=Arrondi'
              + ';MP_AVECINFOCOMPL=InformationsComplementaires'
              + ';MP_AVECNUMAUTOR=NumeroAutorisation'
              + ';MP_CATEGORIE=Categorie'
              + ';MP_CLIOBLIGFO=ClientObligatoire'
              + ';MP_CODEACCEPT=CodeAcceptation'
              + ';MP_CODEAFB=CodeFabrication'
              + ';MP_CONDITION=Condition'
              + ';MP_COPIECBDANSCTRL=CopieNumCarteDansControle'
              + ';MP_COREXPCREDIT=CodeCorrespondanceCredit'
              + ';MP_COREXPDEBIT=CodeCorrespondanceDebit'
              + ';MP_CPTECAISSE=CompteCaisse'
              + ';MP_CPTEREGLE=CompteReglement'
              + ';MP_CPTEREMBQ=CompteRemiseBanque'
              + ';MP_DELAIRETIMPAYE=DelaiRetourImpaye'
              + ';MP_DEVISEFO=Devise'
              + ';MP_EDITABLE=ModeleEdition'
              + ';MP_EDITCHEQUEFO=ImpressionChequeCaisse'
              + ';MP_ENCAISSEMENT=Sens'
              + ';MP_ENVOITPEFO=EnvoiMontantCBTpe'
              + ';MP_FORMATCFONB=CodeCFONB'
              + ';MP_GENERAL=CpteGeneSuiviEffets'
              + ';MP_GEREQTE=GestionQuantite'
              + ';MP_JALCAISSE=JournalCaisse'
              + ';MP_JALREGLE=JournalReglement'
              + ';MP_JALREMBQ=JournalRemiseBanque'
              + ';MP_LETTRECHEQUE=ChequeEditable'
              + ';MP_LETTRETRAITE=TraiteEditable'
              + ';MP_LIBELLE=Libelle'
              + ';MP_MODEPAIE=Code'
              + ';MP_MONTANTMAX=MontantMaximum'
              + ';MP_MONTANTMIN=MontantMinimum'
              + ';MP_POINTABLE=Pointable'
              + ';MP_REMPLACEMAX=ModePaiementRemplacement'
              + ';MP_TYPEMODEPAIE=TypeModePaiement'
              + ';MP_UTILFO=UtilisableCaisse'
              ;
    wsdsZipCode :
      Result := 'O_CODEINSEE=CodeInsee'
              + ';O_CODEPOSTAL=Code'
              + ';O_PAYS=Pays'
              + ';O_VILLE=Ville'
              ;
    wsdsContact :
      Result := 'C_ANNEENAIS=AnneeNaissance'
              + ';C_AUXILIAIRE=CodeAuxiliaire'
              + ';C_BOOLLIBRE1=BooleenLibre1'
              + ';C_BOOLLIBRE2=BooleenLibre12'
              + ';C_BOOLLIBRE3=BooleenLibre13'
              + ';C_CIVILITE=Civilite'
              + ';C_CLEFAX=CleTelephoneBureau'
              + ';C_CLETELEPHONE=CleTelephoneDomicile'
              + ';C_CLETELEX=CleTelephonePortable'
              + ';C_CREATEUR=Createur'
              + ';C_DATECREATION=DateCreation'
              + ';C_DATEFERMETURE=DateFermeture'
              + ';C_DATELIBRE1=DateLibre1'
              + ';C_DATELIBRE2=DateLibre2'
              + ';C_DATELIBRE3=DateLibre3'
              + ';C_DATEMODIF=DateModif'
              + ';C_EMAILING=EMailing'
              + ';C_FAX=TelephoneBureau'
              + ';C_FERME=Ferme'
              + ';C_FONCTION=Fonction'
              + ';C_FONCTIONCODEE=CodeFonction'
              + ';C_GUIDPER=CodePersonne'
              + ';C_GUIDPERANL=CodePersonneLiee'
              + ';C_JOURNAIS=JourNaissance'
              + ';C_LIBRECONTACT1=TableLibre1'
              + ';C_LIBRECONTACT2=TableLibre2'
              + ';C_LIBRECONTACT3=TableLibre3'
              + ';C_LIBRECONTACT4=TableLibre4'
              + ';C_LIBRECONTACT5=TableLibre5'
              + ';C_LIBRECONTACT6=TableLibre6'
              + ';C_LIBRECONTACT7=TableLibre7'
              + ';C_LIBRECONTACT8=TableLibre8'
              + ';C_LIBRECONTACT9=TableLibre9'
              + ';C_LIBRECONTACTA=TableLibre10'
              + ';C_LIENTIERS=TiersAssocie'
              + ';C_LIPARENT=LienParente'
              + ';C_MOISNAIS=MoisNaissance'
              + ';C_NATUREAUXI=Nature'
              + ';C_NOM=Nom'
              + ';C_NUMEROADRESSE=NumeroAdresseTiers'
              + ';C_NUMEROCONTACT=NumeroContact'
              + ';C_NUMIMPORT=NumeroContactApplication'
              + ';C_PRENOM=Prenom'
              + ';C_PRINCIPAL=ContactPrincipal'
              + ';C_PUBLIPOSTAGE=Publipostage'
              + ';C_RVA=Email'
              + ';C_SERVICE=Service'
              + ';C_SERVICECODE=CodeService'
              + ';C_SEXE=Sexe'
              + ';C_SOCIETE=Societe'
              + ';C_TELEPHONE=Telephone1'
              + ';C_TELEX=TelephonePortable'
              + ';C_TEXTELIBRE1=TexteLibre1'
              + ';C_TEXTELIBRE2=TexteLibre2'
              + ';C_TEXTELIBRE3=TexteLibre3'
              + ';C_TIERS=CodeTiers'
              + ';C_TYPECONTACT=TypeContact'
              + ';C_UTILISATEUR=Utilisateur'
              + ';C_VALLIBRE1=ValeurLibre1'
              + ';C_VALLIBRE2=ValeurLibre2'
              + ';C_VALLIBRE3=ValeurLibre3'
              ;
    wsdsFieldsList :
      Result := 'DH_NOMCHAMP=NomChamp';
    wsdsAccForPayment :
      Result := 'E_AUXILIAIRE=Auxiliaire'
              + ';E_COUVERTURE=MontantCouverture'
              + ';E_COUVERTUREDEV=MontantCouvertureDevise'
              + ';E_CREDIT=Credit'
              + ';E_CREDITDEV=CreditDevise'
              + ';E_DATECOMPTABLE=DateComptable'
              + ';E_DATECREATION=DateCreation'
              + ';E_DATEECHEANCE=DateEcheance'
              + ';E_DATEMODIF=DateModif'
              + ';E_DEBIT=Debit'
              + ';E_DEBITDEV=DebitDevise'
              + ';E_ENTITY=Entite'
              + ';E_ETATLETTRAGE=EtatLettrage'
              + ';E_EXERCICE=Exercice'
              + ';E_JOURNAL=Journal'
              + ';E_LETTRAGE=Lettrage'
              + ';E_LIBELLE=Libelle'
              + ';E_MODEPAIE=ModePaiement'
              + ';E_NATUREPIECE=Nature'
              + ';E_NUMEROPIECE=NumeroPiece'
              + ';E_NUMLIGNE=NumeroLigne'
              + ';E_NUMTRAITECHQ=NumeroTraiteCheque'
              + ';E_QUALIFPIECE=TypeEcriture'
              ;
  else
    Result := '';
  end;
end;

end.
