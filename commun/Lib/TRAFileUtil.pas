unit TRAFileUtil;

interface

uses
  Classes
  ;
  
type
  T_TraProduct = (trapNone, trapS1, trapS3, trapS5, trapSI, trapQU, trapWS, trapWT);
  T_TraFileOrigin = (trafoNone, trafoCLI, trafoEXP);
  T_TraFileContent = (trafcNone, trafcJRL, trafcDOS, trafcBAL, trafcSYN, trafcEXO);
  T_TraFileFormat = (traffNone, traffSTD, traffETE);
  T_TraRecordCode = (  trarcNone
                     , trarcGeneralSetting1   // Paramètres généraux 1 - PS1
                     , trarcGeneralSetting2   // Paramètres généraux 2 - PS2
                     , trarcGeneralSetting3   // Paramètres généraux 3 - PS3
                     , trarcGeneralSetting4   // Paramètres généraux 4 - PS4
                     , trarcGeneralSetting5   // Paramètres généraux 5 - PS5
                     , trarcFiscalYear        // Exercice - EXO
                     , trarcEstablishment     // Etablissement - ETB
                     , trarcPaymentMode       // Mode de paiement - MDP
                     , trarcPaymentCondition  // Condition de règlement - MDR
                     , trarcCurrency          // Devise - DEV
                     , trarcTaxSystem         // Régime de TVA - REG
                     , trarcAnalyticalSection // Section analytique - SAT
                     , trarcCorrespondence    // Compte de correspondance - CRR * CRI à voir
                     , trarcAccount           // Comptes - CGN
                     , trarcThird             // Auxiliaire - CAE
                     , trarcJournal           // Journal - JAL
                     , trarcBankId            // RIB - RIB
                     , trarcContact           // Contact - CON
                     , trarcRecovery          // Relance - RRL
                     , trarcChangeRate        // Chancellerie - CHA
                     , trarcChoixCod          // Choix code - CTL
                    );

  FileTraGeneration = class(TObject)
    private
      function GetHeaderStartLine(Product : T_TraProduct; Origin : T_TraFileOrigin; Content : T_TraFileContent; Format : T_TraFileFormat) : string;
      function GetRecordStartLine(traRCode : T_TraRecordCode) : string;
      function GetThirdLine(Data : string) : string;
      function GetBankIdLine(Data : string) : string;
      function GetAnalyticalSection(Data : string) : string;
      function GetCurrency(Data : string) : string;
      function GetChangeRate(Data : string) : string;
      function GetPaymentMode(Data : string) : string;
      function GetPaymentCondition(Data : string) : string;
      function GetChoixCod(Data : string) : string;
      function GetContact(Data : string) : string;
      function GetTraProductCode(tProduct : T_TraProduct) : string;
      function GetFileOrigin(tFileOrig : T_TraFileOrigin) : string;
      function GetFileContent(tFileContent : T_TraFileContent) : string;
      function GetFileFormat(tFileFormat : T_TraFileFormat) : string;
      function GetRecordCode(tRecordCode : T_TraRecordCode) : string;
    public
      NewAna : boolean;

      function GetTraRecordCodeFromTableName(TableName : string) : T_TraRecordCode;
      function GetTraLine(tRecordCode: T_TraRecordCode; Data : string): string;
      function GetTraFirstLine(Product : T_TraProduct; Origin : T_TraFileOrigin; Content : T_TraFileContent; Format : T_TraFileFormat; UserCode : string) : string;
  end;

const
  TRAFixedArea = '***';

implementation

uses
  CommonTools
  , UConnectWSConst
  , SysUtils
  ;
function FileTraGeneration.GetHeaderStartLine(Product : T_TraProduct; Origin : T_TraFileOrigin; Content : T_TraFileContent; Format : T_TraFileFormat) : string;
begin
  Result := TRAFixedArea
          + GetTraProductCode(Product)
          + GetFileOrigin(Origin)
          + GetFileContent(Content)
          + GetFileFormat(Format)
          ; 
end;

function FileTraGeneration.GetRecordStartLine(traRCode : T_TraRecordCode) : string;
begin
  Result := TRAFixedArea + GetRecordCode(traRCode);
end;

function FileTraGeneration.GetThirdLine(Data : string) : string;
var
  MCurrency : string;
  Currency  : string;
begin
  if Data <> '' then
  begin
    MCurrency := Tools.GetStValueFromTSl(Data , 'T_MULTIDEVISE');
    Currency  := Tools.iif(MCurrency = 'X', '', Tools.GetStValueFromTSl(Data , 'T_DEVISE'));
    Result  := GetRecordStartLine(trarcThird)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_AUXILIAIRE')      , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_LIBELLE')         , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_NATUREAUXI')      , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_LETTRABLE')       , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_COLLECTIF')       , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_EAN')             , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE0')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE1')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE2')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE3')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE4')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE5')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE6')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE7')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE8')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TABLE9')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_ADRESSE1')        , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_ADRESSE2')        , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_ADRESSE3')        , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_CODEPOSTAL')      , traaLeft, 9)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_VILLE')           , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_DOMICILIATION')   , traaLeft, 24)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_ETABBQ')          , traaLeft, 5)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_GUICHET')         , traaLeft, 5)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_NUMEROCOMPTE')    , traaLeft, 11)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_CLERIB')          , traaLeft, 2)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_PAYS')            , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_ABREGE')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_LANGUE')          , traaLeft, 3)
             + Tools.FormatValue(MCurrency                                          , traaLeft, 1)
             + Tools.FormatValue(Currency                                           , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TELEPHONE')       , traaLeft, 25)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_FAX')             , traaLeft, 25)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_REGIMETVA')       , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_MODEREGLE')       , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_COMMENTAIRE')     , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_NIF')             , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_SIRET')           , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_APE')             , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_PRENOM')          , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_SERVICE')         , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_FONCTION')        , traaLeft, 35)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TELEPHONE')       , traaLeft, 25)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_FAX')             , traaLeft, 25)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TELEX')           , traaLeft, 25)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_RVA')             , traaLeft, 25)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_CIVILITE')        , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_PRINCIPAL')       , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_FORMEJURIDIQUE')  , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_PRINCIPAL')       , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_TVAENCAISSEMENT') , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_PAYEUR')          , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_ISPAYEUR')        , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_AVOIRRBT')        , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_RELANCEREGLEMENT'), traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_RELANCETRAITE')   , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_CONFIDENTIEL')    , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_CORRESP1')        , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_CORRESP2')        , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_ESCOMPTE')        , traaRigth, 20, 2)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_REMISE')          , traaRigth, 20, 2)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_FACTURE')         , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_JURIDIQUE')       , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_CREDITDEMANDE')   , traaRigth, 20, 2)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_CREDITACCORDE')   , traaRigth, 20, 2)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_CREDITPLAFOND')   , traaRigth, 20, 2)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_FERME')           , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_FACTUREHT')       , traaLeft, 1)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_SOCIETEGROUPE')   , traaLeft, 17)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_RELANCETRAITE')   , traaLeft, 3)
             + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'T_RELANCEREGLEMENT'), traaLeft, 3)
             ;
  end else
    Result := '';
end;

function FileTraGeneration.GetBankIdLine(Data : string) : string;
begin
  if Data <> '' then
    Result := GetRecordStartLine(trarcBankId)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_AUXILIAIRE')   , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_NUMERORIB')    , traaRigth, 6)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_PRINCIPAL')    , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_ETABBQ')       , traaLeft, 5)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_GUICHET')      , traaLeft, 5)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_NUMEROCOMPTE') , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_CLERIB')       , traaLeft, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_DOMICILIATION'), traaLeft, 24)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_VILLE')        , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_PAYS')         , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_DEVISE')       , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_CODEBIC')      , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_SOCIETE')      , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_SALAIRE')      , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_ACOMPTE')      , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_FRAISPROF')    , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_CODEIBAN')     , traaLeft, 70)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_NATECO')       , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_TYPEPAYS')     , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_ETABBQ')       , traaLeft, 8)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'R_NUMEROCOMPTE') , traaLeft, 20)
  else
    Result := '';
end;

function FileTraGeneration.GetAnalyticalSection(Data : string) : string;
begin
  if Data <> '' then
    Result := GetRecordStartLine(trarcAnalyticalSection)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_SECTION') , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_LIBELLE') , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_AXE')     , traaLeft, 3)
            + Tools.FormatValue('S1'                                      , traaLeft, 3)      
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE0')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE1')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE2')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE3')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE4')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE5')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE6')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE7')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE8')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_TABLE9')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_ABREGE')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_SENS')    , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'S_FERME')   , traaLeft, 1)
  else
    Result := '';
end;

function FileTraGeneration.GetCurrency(Data : string) : string;
begin
  if Data <> '' then
    Result := GetRecordStartLine(trarcCurrency)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_DEVISE')        , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_LIBELLE')       , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_SYMBOLE')       , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_FERME')         , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_DECIMALE')      , traaLeft, 1) 
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_QUOTITE')       , traaRigth, 6, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_MONNAIEIN')     , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_PARITEEURO')    , traaRigth, 20, 9)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_CODEISO')       , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_FONGIBLE')      , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_CPTLETTRDEBIT') , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_CPTLETTRCREDIT'), traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_CPTPROVDEBIT')  , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_CPTPROVCREDIT') , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_MAXDEBIT')      , traaRigth, 20, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'D_MAXCREDIT')     , traaRigth, 20, 2)
  else
    Result := '';
end;

function FileTraGeneration.GetChangeRate(Data : string) : string;
var
  FreeDate1 : string;
begin
  if Data <> '' then
  begin
    FreeDate1 := Tools.GetStValueFromTSl(Data, 'H_DATECOURS');
    if StrToDateTime(FreeDate1) < 2 then FreeDate1 := DateTimeToStr(2);
    Result := GetRecordStartLine(trarcChangeRate)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'H_DEVISE')     , traaLeft, 3)
            + Tools.FormatValue(FreeDate1                                     , traaLeft, 8, 0, tvtDateTime)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'H_TAUXREEL')   , traaRigth, 9)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'H_TAUXCLOTURE'), traaRigth, 9)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'H_COMMENTAIRE'), traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'H_SOCIETE')    , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'H_COTATION')   , traaRigth, 9)
  end else
    Result := '';
end;

function FileTraGeneration.GetPaymentMode(Data : string) : string;
begin
  if Data <> '' then
    Result := GetRecordStartLine(trarcPaymentMode)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_MODEPAIE')    , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_LIBELLE')     , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_CATEGORIE')   , traaRigth, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_CODEACCEPT')  , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_LETTRECHEQUE'), traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_LETTRETRAITE'), traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_CONDITION')   , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_MONTANTMAX')  , traaRigth, 20, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_REMPLACEMAX') , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MP_GENERAL')     , traaLeft, 17)
  else
    Result:= '';
end;

function FileTraGeneration.GetPaymentCondition(Data : string) : string;
var
  Cpt : integer;
begin
  if Data <> '' then
  begin
    Result := GetRecordStartLine(trarcPaymentCondition)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_MODEREGLE')     , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_LIBELLE')       , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_APARTIRDE')     , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_PLUSJOUR')      , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_ARRONDIJOUR')   , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_NOMBREECHEANCE'), traaLeft, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_SEPAREPAR')     , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_MONTANTMIN')    , traaRigth, 20, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_REMPLACEMIN')   , traaLeft, 3)
            ;
    for Cpt := 1 to 12 do
      Result := Result + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_MP'   + IntToStr(Cpt)), traaLeft, 3)
                       + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'MR_TAUX' + IntToStr(Cpt)), traaRigth, 6, 2);
  end else
    Result:= '';
end;

function FileTraGeneration.GetChoixCod(Data : string) : string;
begin
  if Data <> '' then
    Result := GetRecordStartLine(trarcChoixCod)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'CC_TYPE')   , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'CC_CODE')   , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'CC_LIBELLE'), traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'CC_ABREGE') , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'CC_LIBRE')  , traaLeft, 70)
  else
    Result := '';    
end;

function FileTraGeneration.GetContact(Data : string) : string;
var
  FreeDate1 : string;
  FreeDate2 : string;
  FreeDate3 : string;
begin
  if Data <> '' then
  begin
    FreeDate1 := Tools.GetStValueFromTSl(Data, 'C_DATELIBRE1');
    FreeDate2 := Tools.GetStValueFromTSl(Data, 'C_DATELIBRE2');
    FreeDate3 := Tools.GetStValueFromTSl(Data, 'C_DATELIBRE3');
    if StrToDateTime(FreeDate1) < 2 then FreeDate1 := DateTimeToStr(2);
    if StrToDateTime(FreeDate2) < 2 then FreeDate2 := DateTimeToStr(2);
    if StrToDateTime(FreeDate3) < 2 then FreeDate3 := DateTimeToStr(2);
    Result := GetRecordStartLine(trarcContact)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TYPECONTACT')  , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_AUXILIAIRE')   , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_NUMEROCONTACT'), traaRigth, 6)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_NATUREAUXI')   , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_PRINCIPAL')    , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_NOM')          , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_PRENOM')       , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_SERVICE')      , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_FONCTION')     , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TELEPHONE')    , traaLeft, 25)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_FAX')          , traaLeft, 25)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TELEX')        , traaLeft, 25)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_RVA')          , traaLeft, 250)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_SOCIETE')      , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_CIVILITE')     , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_FONCTIONCODEE'), traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_LIPARENT')     , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_JOURNAIS')     , traaRigth, 6)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_MOISNAIS')     , traaRigth, 6)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_ANNEENAIS')    , traaRigth, 6)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_SEXE')         , traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_PUBLIPOSTAGE') , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TEXTELIBRE1')  , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TEXTELIBRE2')  , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TEXTELIBRE3')  , traaLeft, 35)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_BOOLLIBRE1')   , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_BOOLLIBRE2')   , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_BOOLLIBRE3')   , traaLeft, 1)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_LIBRECONTACT1'), traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_LIBRECONTACT2'), traaLeft, 3)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_LIBRECONTACT3'), traaLeft, 3)
            + Tools.FormatValue(FreeDate1                                       , traaLeft, 8, 0, tvtDate)
            + Tools.FormatValue(FreeDate2                                       , traaLeft, 8, 0, tvtDate)
            + Tools.FormatValue(FreeDate3                                       , traaLeft, 8, 0, tvtDate)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_VALLIBRE1')    , traaRigth, 15, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_VALLIBRE2')    , traaRigth, 15, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_VALLIBRE3')    , traaRigth, 15, 2)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_TIERS')        , traaLeft, 17)
            + Tools.FormatValue(Tools.GetStValueFromTSl(Data, 'C_CLETELEPHONE') , traaLeft, 25)
  end else
    Result := '';
end;
  
function FileTraGeneration.GetTraProductCode(tProduct : T_TraProduct) : string;
begin
  case tProduct of
    trapS1 : Result := 'S1';
    trapS3 : Result := 'S3';
    trapS5 : Result := 'S5';
    trapSI : Result := 'SI';
    trapQU : Result := 'QU';
    trapWS : Result := 'WS';
    trapWT : Result := 'WT';
  else
    Result := '';
  end;
end;

function FileTraGeneration.GetFileOrigin(tFileOrig : T_TraFileOrigin) : string;
begin
  case tFileOrig of
    trafoCLI : Result := 'CLI';
    trafoEXP : Result := 'EXP';
  else
    Result := '';
  end;
end;

function FileTraGeneration.GetFileContent(tFileContent: T_TraFileContent): string;
begin
  case tFileContent of
    trafcJRL : Result := 'JRL';
    trafcDOS : Result := 'DOS';
    trafcBAL : Result := 'BAL';
    trafcSYN : Result := 'SYN';
    trafcEXO : Result := 'EXO';
  else
    Result := '';
  end;
end;

function FileTraGeneration.GetFileFormat(tFileFormat: T_TraFileFormat): string;
begin
  case tFileFormat of
    traffSTD : Result := 'STD';
    traffETE : Result := 'ETE';
  else
    Result := '';
  end;
end;

function FileTraGeneration.GetRecordCode(tRecordCode: T_TraRecordCode): string;
begin
  case tRecordCode of
    trarcGeneralSetting1   : Result := 'PS1';
    trarcGeneralSetting2   : Result := 'PS2';
    trarcGeneralSetting3   : Result := 'PS3';
    trarcGeneralSetting4   : Result := 'PS4';
    trarcGeneralSetting5   : Result := 'PS4';
    trarcFiscalYear        : Result := 'EXO';
    trarcEstablishment     : Result := 'ETB';
    trarcPaymentMode       : Result := 'MDP';
    trarcPaymentCondition  : Result := 'MDR';
    trarcCurrency          : Result := 'DEV';
    trarcTaxSystem         : Result := 'REG';
    trarcAnalyticalSection : Result := Tools.iif(NewAna, 'NSA', 'SAT');
    trarcCorrespondence    : Result := 'CRR'; //CRI à voir
    trarcAccount           : Result := 'CGN';
    trarcThird             : Result := 'CAE';
    trarcJournal           : Result := 'JAL';
    trarcBankId            : Result := 'RIB';
    trarcContact           : Result := 'CON';
    trarcRecovery          : Result := 'RRL';
    trarcChangeRate        : Result := 'CHA';
    trarcChoixCod          : Result := 'CTL';
  else
    Result := '';
  end;
end;

function FileTraGeneration.GetTraRecordCodeFromTableName(TableName: string): T_TraRecordCode;
begin
  case Tools.CaseFromString(TableName, [  Tools.GetTableNameFromTtn(ttnTiers)
                                        , Tools.GetTableNameFromTtn(ttnRib)
                                        , Tools.GetTableNameFromTtn(ttnSection)
                                        , Tools.GetTableNameFromTtn(ttnDevise)
                                        , Tools.GetTableNameFromTtn(ttnChancell)
                                        , Tools.GetTableNameFromTtn(ttnModeRegl)
                                        , Tools.GetTableNameFromTtn(ttnChoixCod)
                                        , Tools.GetTableNameFromTtn(ttnContact)
                                       ]) of
    {TIERS}     0 : Result := trarcThird;
    {RIB}       1 : Result := trarcBankId;
    {CSECTION}  2 : Result := trarcAnalyticalSection;
    {DEVISE}    3 : Result := trarcCurrency;
    {CHANCELL}  4 : Result := trarcChangeRate;
    {MODEREGL}  5 : Result := trarcPaymentCondition;
    {CHOIXCOD}  6 : Result := trarcChoixCod;
    {CONTACT}   7 : Result := trarcContact;
  else
    Result := trarcNone;
  end;
end;

function FileTraGeneration.GetTraLine(tRecordCode: T_TraRecordCode; Data : string): string;
begin
  Result := '';
  if (tRecordCode <> trarcNone) and (Data <> '') then
  begin
    case tRecordCode of
      {TIERS}    trarcThird             : Result := GetThirdLine(Data);
      {RIB}      trarcBankId            : Result := GetBankIdLine(Data);
      {CSECTION} trarcAnalyticalSection : Result := GetAnalyticalSection(Data);
      {DEVISE}   trarcCurrency          : Result := GetCurrency(Data);
      {CHANCELL} trarcChangeRate        : Result := GetChangeRate(Data);
      {MODEPAIE} trarcPaymentMode       : Result := GetPaymentMode(Data);
      {MODEREGL} trarcPaymentCondition  : Result := GetPaymentCondition(Data);
      {CHOIXCOD} trarcChoixCod          : Result := GetChoixCod(Data);
      {CONTACT}  trarcContact           : Result := GetContact(Data);
    end;
  end;
end;

function FileTraGeneration.GetTraFirstLine(Product : T_TraProduct; Origin : T_TraFileOrigin; Content : T_TraFileContent; Format : T_TraFileFormat; UserCode : string) : string;
var
  DefaultDate : string;
begin
  DefaultDate := '01011900';
  Result :=  GetHeaderStartLine(Product, Origin, Content, Format)               // Début ligne
           + Tools.FormatValue(''                , traaLeft, 3)                 // Code exo
           + Tools.FormatValue(DefaultDate       , traaLeft, 8)                 // Date bascule euro
           + Tools.FormatValue(DefaultDate       , traaLeft, 8)                 // Date arrêtée périodique
           + Tools.FormatValue('010'             , traaLeft, 3)                 // Version
           + Tools.FormatValue(''                , traaLeft, 5)                 // N° dossier Cab
           + Tools.FormatValue(DateTimeToStr(Now), traaLeft, 3, 0, tvtDateTime) // Date et heure
           + Tools.FormatValue(UserCode          , traaLeft, 35)                // Utilisateur
           + Tools.FormatValue(''                , traaLeft, 35)                // Raison Sociale
           + Tools.FormatValue(''                , traaLeft, 4)                 // Indicateur de reprise
           + Tools.FormatValue(''                , traaLeft, 6)                 // N° dossier
           + Tools.FormatValue(''                , traaLeft, 3)                 // Info échange fréquence
           + Tools.FormatValue(DefaultDate       , traaLeft, 8)                 // Date début 1er exercice
           + Tools.FormatValue('001'             , traaLeft, 3)                 // Sous version du fichier
           + Tools.FormatValue('-'               , traaLeft, 3)                 // Nouvelle gestion analytique
           + Tools.FormatValue('-'               , traaLeft, 3)                 // Nouvelle gestion des transferts
           + Tools.FormatValue('-'               , traaLeft, 3)                 // Dossier agricole
           + Tools.FormatValue('0'               , traaRigth, 6)                // Numéro séquence
           ;
end;

//BTPValues

end.
