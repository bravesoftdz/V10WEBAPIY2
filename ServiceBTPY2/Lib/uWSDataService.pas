unit uWSDataService;

interface

uses
  Classes
  , UConnectWSConst
  ;

type
  TReadWSDataService = class (TObject)
  private
    class function GetODataOperator(Operator : string) : string;
  public
    class function GetPrefixFromDSType(DSType : T_WSDataService) : string;
    class function GetTableNameFromDSType(DSType : T_WSDataService) : string;
    class function GetViewNameFromDSType(DSType : T_WSDataService) : string;
    class function GetWSNameFromDSType(DSType : T_WSDataService) : string;
    class function GetFiedsListFromDsType(DSType : T_WSDataService) : string;
    class function GetData(DSType : T_WSDataService; ServerName, FolderName : string; var TslResult, TslViewFields : TStringList; TslFilter : TStringList=nil; KnownUrl : string='') : string;
  end;

implementation

uses
  CommonTools
  , WinHttp_TLB
  , uLkJSON
  , SysUtils
  , Variants
  {$IF not defined(APPSRV)}
  , UConnectWSCEGID
  {$IFEND !APPSRV}
  ;

class function TReadWSDataService.GetODataOperator(Operator : string) : string;
begin
  case Tools.CaseFromString(Operator, ['=', '<>', '>', '>=', '<', '<=', 'AND', 'OR', 'NOT']) of
    {=}   0 : Result := 'eq';
    {<>}  1 : Result := 'ne';
    {>}   2 : Result := 'gt';
    {>=}  3 : Result := 'ge';
    {<}   4 : Result := 'lt';
    {>=}  5 : Result := 'le';
    {AND} 6 : Result := 'and';
    {OR}  7 : Result := 'or';
    {NOT} 8 : Result := 'not';
  else
    Result := '';
  end;
end;

class function TReadWSDataService.GetPrefixFromDSType(DSType : T_WSDataService) : string;
begin
  case DSType of
    wsdsThird              : Result := 'T';
    wsdsAnalyticalSection  : Result := 'S';
    wsdsAccount            : Result := 'G';
    wsdsJournal            : Result := 'J';
    wsdsBankIdentification : Result := 'R';
    wsdsChoixCod           : Result := 'CC';
    wsdsCommon             : Result := 'CO';
    wsdsRecovery           : Result := 'RR';
    wsdsCountry            : Result := 'PY';
    wsdsCurrency           : Result := 'D';
    wsdsCorrespondence     : Result := 'CR';
    wsdsPaymenChoice       : Result := 'MR';
    wsdsChangeRate         : Result := 'H';
    wsdsFiscalYear         : Result := 'EX';
    wsdsSocietyParameters  : Result := 'SOC';
    wsdsEstablishment      : Result := 'ET';
    wsdsPaymentMode        : Result := 'MP';
    wsdsZipCode            : Result := 'O';
    wsdsContact            : Result := 'C';
  else
    Result := '';
  end;
end;

class function TReadWSDataService.GetTableNameFromDSType(DSType : T_WSDataService) : string;
begin
  case DSType of
    wsdsThird              : Result := Tools.GetTableNameFromTtn(ttnTiers);
    wsdsAnalyticalSection  : Result := Tools.GetTableNameFromTtn(ttnSection);
    wsdsAccount            : Result := Tools.GetTableNameFromTtn(ttnGeneraux);
    wsdsJournal            : Result := Tools.GetTableNameFromTtn(ttnJournal);
    wsdsBankIdentification : Result := Tools.GetTableNameFromTtn(ttnRib);
    wsdsChoixCod           : Result := Tools.GetTableNameFromTtn(ttnChoixCod);
    wsdsCommon             : Result := Tools.GetTableNameFromTtn(ttnCommun);
    wsdsRecovery           : Result := Tools.GetTableNameFromTtn(ttnRelance);
    wsdsCountry            : Result := Tools.GetTableNameFromTtn(ttnPays);
    wsdsCurrency           : Result := Tools.GetTableNameFromTtn(ttnDevise);
    wsdsCorrespondence     : Result := Tools.GetTableNameFromTtn(ttnCorresp);
    wsdsPaymenChoice       : Result := Tools.GetTableNameFromTtn(ttnModeRegl);
    wsdsChangeRate         : Result := Tools.GetTableNameFromTtn(ttnChancell);
    wsdsFiscalYear         : Result := Tools.GetTableNameFromTtn(ttnExercice);
    wsdsSocietyParameters  : Result := Tools.GetTableNameFromTtn(ttnParamSoc);
    wsdsEstablishment      : Result := Tools.GetTableNameFromTtn(ttnEtabliss);
    wsdsPaymentMode        : Result := Tools.GetTableNameFromTtn(ttnModePaie);
    wsdsZipCode            : Result := Tools.GetTableNameFromTtn(ttnCodePostaux);
    wsdsContact            : Result := Tools.GetTableNameFromTtn(ttnContact);
  else
    Result := '';
  end;
end;

class function TReadWSDataService.GetViewNameFromDSType(DSType : T_WSDataService) : string;
begin
  case DSType of
    wsdsThird              : Result := 'ZLSETHIRDSLIST';
    wsdsAnalyticalSection  : Result := 'ZLSEANALYTICALSECTION';
    wsdsAccount            : Result := 'ZLSEACCOUNT';
    wsdsJournal            : Result := 'ZLSEJOURNAL';
    wsdsBankIdentification : Result := 'ZLSEBANKID';
    wsdsChoixCod           : Result := 'ZLSECHXOIXCOD';
    wsdsCommon             : Result := 'ZLSECOMMUN';
    wsdsRecovery           : Result := 'ZLSERECOVERY';
    wsdsCountry            : Result := 'ZLSECOUNTRY';
    wsdsCurrency           : Result := 'ZLSECURRENCY';
    wsdsCorrespondence     : Result := 'ZLSECORRESP';
    wsdsPaymenChoice       : Result := 'ZLSEPAYMENT';
    wsdsChangeRate         : Result := 'ZLSECHANGERATE';
    wsdsFiscalYear         : Result := 'ZLSEFISCALYEAR';
    wsdsSocietyParameters  : Result := 'ZLSESOCIETYPARAM';
    wsdsEstablishment      : Result := 'ZLSEESTABLISHMENT';
    wsdsPaymentMode        : Result := 'ZLSEPAYMENTMODE';
    wsdsZipCode            : Result := 'ZLSEZIPCODE';
    wsdsContact            : Result := 'ZLSECONTACT';
  else
    Result := '';
  end;
end;

class function TReadWSDataService.GetWSNameFromDSType(DSType: T_WSDataService): string;
begin
  case DSType of
    wsdsThird              : Result := 'lseTHIRDSLIST';
    wsdsAnalyticalSection  : Result := 'lseANALYTICALSECTION';
    wsdsAccount            : Result := 'lseACCOUNTLIST';
    wsdsJournal            : Result := 'lseJOURNALLIST';
    wsdsBankIdentification : Result := 'lseBANKILIST';
    wsdsChoixCod           : Result := 'lsePARAMCC';
    wsdsCommon             : Result := 'lsePARAMCO';
    wsdsRecovery           : Result := 'lseRECOVERYLIST';
    wsdsCountry            : Result := 'lseCOUNTRYLIST';
    wsdsCurrency           : Result := 'lseCURRENCYLIST';
    wsdsCorrespondence     : Result := 'lseCORRESPONDENCELIST';
    wsdsPaymenChoice       : Result := 'lsePAYMENTLIST';
    wsdsChangeRate         : Result := 'lseCHANGERATE';
    wsdsFiscalYear         : Result := 'lseFISCALYEAR';
    wsdsSocietyParameters  : Result := 'lseSOCIETYPARAM';
    wsdsEstablishment      : Result := 'lseESTABLISHMENT';
    wsdsPaymentMode        : Result := 'lsePAYMENTMODE';
    wsdsZipCode            : Result := 'lseZIPCODE';
    wsdsContact            : Result := 'lseCONTACT';
    wsdsFieldsList         : Result := 'lseFIELDSLIST';
  else
    Result := '';
  end;
end;

class function TReadWSDataService.GetFiedsListFromDsType(DSType : T_WSDataService) : string;
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
      Result := ';G_ABREGE=Abrege'
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
    wsdsBankIdentification :
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
    wsdsChoixCod :
      Result := 'CC_ABREGE=Abrege'
              + ';CC_CODE=Code'
              + ';CC_LIBELLE=Libelle'
              + ';CC_LIBRE=Libre'
              + ';CC_TYPE=Type'
              ;
    wsdsCommon :
      Result := 'CO_ABREGE=Abrege'
              + ';CO_CODE=Code'
              + ';CO_LIBELLE=Libelle'
              + ';CO_LIBRE=Libre'
              + ';CO_TYPE=Type'
              ;
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
  else
    Result := '';
  end;
end;

{ DSType : WebApi
  TslResult : TstringList renvoyée
  TslFilter : TStringList des filtres sous la forme :
  . 1ère valeur : opérateur logique (; s'il n'y en a pas, sinon (ex) AND;)
  . 2ème valeur : nom du champ
  . 3ème valeur : opérateur (à traduire pour oData)
  . 4ème valeur :
    . valeur
      ou
    . ( début de groupe
    . ) fin de groupe

  EXEMPLE :
  Avoir la liste des clients et fournisseurs créés ou modifiés après le 01/04/2018 13:53:48
    TSlFilter.Add(';;;(');                                           : début groupe
    TSlFilter.Add(';T_NATUREAUXI;=;CLI');                            : condition
    TSlFilter.Add('OR;T_NATUREAUXI;=;FOU');                          : condition
    TSlFilter.Add(';;;)');                                           : fin groupe
    TSlFilter.Add('AND;T_DATEMODIF;>=;'202018-04-01T13:53:48.000Z'); : condition
  Le filtre ci dessus correspond :
    pour oData : $filter=(Nature eq 'CLI' or Nature eq 'FOU') and DateModif ge 02018-04-01T13:53:48.000Z
    pour SQL   : WHERE (T_NATUREAUXI='CLI OR T_NATUREAUXI='FOU') AND T_DATEMODIF >= '202018-04-01T13:53:48.000Z'

  ATTENTION, si passage de date en paramètre, celle-ci doit être au format "string"
}
class function TReadWSDataService.GetData(DSType : T_WSDataService; ServerName, FolderName : string; var TslResult, TslViewFields : TStringList; TslFilter : TStringList=nil; KnownUrl : string='') : string;
var
  http     : IWinHttpRequest;
  Url      : string;
  Response : string;
  Values   : string;
  Fields   : string;
  Alias    : string;
  Field    : string;
  NewUrl   : string;
  JSon     : TlkJSONBase;
  Items    : TlkJSONBase;
  Item     : TlkJSONBase;
  Cpt      : Integer;
  Cpt1     : Integer;

  function GetFieldName(Value : string) : string;
  begin
    Result := Copy(Value, 1, pos('=', Value) -1);
  end;

  function GetAliasName(Value : string) : string;
  begin
    Result := Copy(Value, pos('=', Value) + 1, Length(Value));
  end;

  function GetFieldsFromDSType : string;
  var
    iPos : integer;
  begin
    iPos := TslViewFields.IndexOfName(TReadWSDataService.GetWSNameFromDSType(DSType));
    if iPos > -1 then
      Result := TslViewFields.ValueFromIndex[iPos];
  end;

  function GetFilter : string;
  var
    CptF            : integer;
    Filter          : string;
    FilterField     : string;
    FilterOperator  : string;
    FilterValue     : string;
    FilterLogicalOp : string;
    Values          : string;
    Field           : string;
    Alias           : string;
    Value           : Variant;
    StartEndGroup   : Boolean;
  begin
    if (Assigned(TslFilter)) and (TslFilter[0] <> WSCDS_EmptyValue) then
    begin
      for CptF := 0 to pred(TslFilter.Count) do
      begin
        Filter          := TslFilter.Strings[CptF];
        FilterLogicalOp := GetODataOperator(Tools.ReadTokenSt_(Filter, ';'));
        FilterField     := Tools.ReadTokenSt_(Filter, ';');
        FilterOperator  := GetODataOperator(Tools.ReadTokenSt_(Filter, ';'));
        FilterValue     := Tools.ReadTokenSt_(Filter, ';');
        Values          := GetFieldsFromDSType;
        while Values <> '' do
        begin
          Value := Tools.ReadTokenSt_(Values, ';');
          Field := GetFieldName(Value);
          Alias := GetAliasName(Value);
          if Field = FilterField then
            Break;
        end;
        StartEndGroup := ((FilterValue = '(') or (FilterValue = ')'));
        if not StartEndGroup then
        begin
          if Tools.GetFieldType(FilterField{$IF defined(APPSRV)}, ServerName, FolderName{$IFEND !APPSRV}) = ttfDate then
            FilterValue := FormatDateTime('yyyy-mm-dd', Int(StrToDateTime(FilterValue))) + 'T' +  FormatDateTime('hh:nn:ss.zzz', StrToDateTime(FilterValue)) + 'Z'
          else
            FilterValue := '''' + FilterValue + '''';
          Result := Result
                  + Tools.iif(FilterLogicalOp <> '', ' ' + FilterLogicalOp, '')
                  + ' ' + Alias
                  + ' ' + FilterOperator
                  + ' ' + FilterValue
                    ;
        end else
          Result := Result + FilterValue;
      end;                                              
      Result := '?$filter=' + Result;
    end else
      Result := '';
  end;

begin
  if (Assigned(TslResult)) and (DSType <> wsdsNone) then
  begin
    if KnownUrl = '' then
      Url := 'http://'
           + ServerName
           + '/CegidDataService/odata'
           + '/' + FolderName
           + '/' + GetWSNameFromDSType(DSType)
           + GetFilter
    else
      Url := KnownUrl;
    http := CoWinHttpRequest.Create;
    try
      http.SetAutoLogonPolicy(0);
      http.Open('GET', Url, False);
      TRY
        http.Send(EmptyParam);
      EXCEPT
        on E: Exception do
        begin
          {$IF not defined(APPSRV)}
          ShowMessage(E.Message);
          {$IFEND !APPSRV}
          exit;
        end;
      END;
      if http.Status = 200 then
      begin
        Response := '[' + http.ResponseText + ']';
        JSon := TlkJSON.ParseText(Response);
        for Cpt := 0 to pred(JSon.Count) do
        begin
          Items := JSon.Child[Cpt].Field['value'];
          for Cpt1 := 0 to Pred(Items.Count) do
          begin
            Item := Items.Child[Cpt1];
            TslResult.Add(WSCDS_IndiceField + IntToStr(Cpt1) + '#=' + IntToStr(Cpt1));
            Values := GetFieldsFromDSType;
            while Values <> '' do
            begin
              Fields := Tools.ReadTokenSt_(Values, ';');
              Field  := GetFieldName(Fields);
              Alias  := GetAliasName(Fields);
              if Tools.GetFieldType(Field{$IF defined(APPSRV)}, ServerName, FolderName{$IFEND APPSRV}) = ttfBoolean then
                TslResult.Add(Field + '=' + Tools.iif(Item.Field[Alias].Value, 'X', '-'))
              else
                TslResult.Add(Field + '=' + VarToStr(Item.Field[Alias].Value));
            end;
          end;
        end;
        if pos(WSCDS_NextUrlValue, Response) > 0 then
        begin
          NewUrl := Copy(Response, Pos(WSCDS_NextUrlValue, Response) + Length(WSCDS_NextUrlValue), Length(Response)-2);
          NewUrl := Copy(NewUrl, 1, Pos('"', NewUrl) -1);
          GetData(DSType, ServerName, FolderName, TslResult, TslViewFields, TslFilter, NewUrl);
        end;
        Result := WSCDS_GetDataOk;
      end else
        Result := Format('Erreur %s - %s', [IntToStr(http.Status), http.StatusText]);
    finally
      http := nil;
    end;
  end else
    Result := 'TslResult ou DSType non assigné.';
end;

end.
