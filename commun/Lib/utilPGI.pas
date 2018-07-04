unit utilPGI;

interface

uses hent1,sysutils,UTOB,ParamSoc,Stdctrls,classes,
{$IFNDEF EAGLSERVER}
     graphics,extctrls,Forms,windows,
{$ELSE  EAGLSERVER}
     eSession,
{$ENDIF EAGLSERVER}
{$IFDEF EAGLCLIENT}
     eFichList,
{$ELSE EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     {$IFNDEF EAGLSERVER}
     UBob,
     FichList,
     {$ENDIF EAGLSERVER}
{$ENDIF EAGLCLIENT}
     {$IFNDEF EAGLSERVER}
     AglMail,MailOl,M3FP,
     {$ENDIF EAGLSERVER}
{$IFDEF GRC}
     EntRT,
{$ENDIF GRC}
{$IFDEF GIGI}
     LicUtil,
{$ENDIF GIGI}
   variants, Hdb,
{$IFNDEF NOVH}
EntGC, // pour R_CleDoc  supprimé car entgc contient des VH^.
{$ENDIF NOVH}

  HCtrls,Jpeg, comctrls,shellapi, HMsgBox,
  ULibIdentBancaire ,
  HRichOle
{$IFNDEF EAGLSERVER}
   ,CBPPath
{$ENDIF}
  ,MajTable
  ,Ent1
  {$IFDEF MODENT1}
  , CPTypeCons
  , CPVersion
  , CPProcMetier
  {$ENDIF MODENT1}
  ,UentCommun
   ; // js1 120107 pour connectedb


{$IFNDEF NOVH}
Function EncodeRefPresqueCPGescom ( TOBPiece : TOB ) : String ; overload;
Function EncodeRefPresqueCPGescom ( Cledoc : r_cledoc ) : String ; overload;

{Intégration GC <--> CPTA}
Function EncodeRefCPGescom ( LaTOB : TOB; FromAcc : boolean=false) : String ;
function  DecodeRefCPGescom ( RefG : String ) : R_CleDoc ;
procedure DecodeRefPiece(St: string; var CleDoc: R_CleDoc);
procedure AGLDecodeRefPiece(Parms : Array of Variant; nb: integer);
function EncodeRefGCCpta(TobEcr : TOB) : string;

{ Manipulation TClesDocs }
procedure AddCleDocInTClesDocs(const CleDoc: R_CleDoc; var ClesDocs: TClesDocs); overload;
procedure AddCleDocInTClesDocs(const RefPiece: String; var ClesDocs: TClesDocs); overload;
function ExistCleDocInTClesDocs(const CleDoc: R_CleDoc; ClesDocs: TClesDocs; const TestNumOrdre: Boolean): Boolean; overload;
function ExistCleDocInTClesDocs(const RefPiece: String; ClesDocs: TClesDocs; const TestNumOrdre: Boolean): Boolean; overload;
{$ENDIF !NOVH}

function CanCloseExoFromGC(Exercice : string) : boolean ;
{$IFNDEF NOVH}
function CanValidPieceFromGC(RefGescom : string) : boolean;
function MajRefCptaOnValidPce(TobE : TOB) : boolean;
{$ENDIF !NOVH}


Function  Evaluedate ( St : String ) : TDateTime ;
{Divers}
Function  ExisteCarInter ( St : String ) : boolean ;
Function ErreurDansIban(Rib : String) : boolean;
Function VerifRib(Banque,Guichet,Compte : String ; ISOPays : String = CodeISOFR) : String ; //XMG 14/07/03
Function  VerifSiret (Siret : string) : boolean ;
{$IFDEF CHR}  { CHR_AGF_20080417}
function  PGIEnvoiMail(SUJET: hstring; AQUI, CopieA: string; Corps: hTStrings; FICHIERS: string; EnvoiAuto: boolean = TRUE; Importance: Integer = 1; CATEGORIE: string = ''; COMPANIE: string = ''; SansAlerte: boolean = FALSE): boolean;
{$ELSE}
procedure PGIEnvoiMail(SUJET: hstring; AQUI, CopieA: string; Corps: hTStrings; FICHIERS: string; EnvoiAuto: boolean = TRUE; Importance: Integer = 1; CATEGORIE: string = ''; COMPANIE: string = ''; SansAlerte: boolean = FALSE);
{$ENDIF CHR}
Function isIntraComValide(num : String; var siren : String) : Boolean ;
function VerifTVA(stIsoPays, num : string) : boolean ;
{RIB}
function EncodeRIB ( Etab,Guichet,Numero,Cle,Dom : String ; ISOPays : String = CodeISOFR ; TypePays : String ='' ) : String ; //XMG 14/07/03
Procedure DecodeRIB ( Var Etab,Guichet,Numero,Cle,Dom : String ; RIB : String ; ISOPays : String = CodeISOFR ) ; //XMG 14/07/03
function  EncodeRIBIban ( Etab,Guichet,Numero,Cle,Dom,Iban : String ; ISOPays : String = CodeISOFR ) : String ; //XMG 14/07/03
procedure DecodeRIBIban ( Var Etab,Guichet,Numero,Cle,Dom,Iban : String ; RIB : String ) ; // XVI 24/02/2005

//Franco francaise
function  EncodeRIB_OLD ( Etab,Guichet,Numero,Cle,Dom : String ; ISOPays : String = CodeISOFR ) : String ; //XMG 14/07/03
Procedure DecodeRIB_OLD ( Var Etab,Guichet,Numero,Cle,Dom : String ; RIB : String ; ISOPays : String = CodeISOFR ) ; //XMG 14/07/03
function  EncodeRIBIban_OLD ( Etab,Guichet,Numero,Cle,Dom,Iban : String ; ISOPays : String = CodeISOFR ) : String ; //XMG 14/07/03
procedure DecodeRIBIban_OLD ( Var Etab,Guichet,Numero,Cle,Dom,Iban : String ; RIB : String ) ; // XVI 24/02/2005

function  EncodeIban(pszIban : String) : String;
function  DecodeIban(pszIban : String) : String;
function  CodeIsoDuPays(Pays : string) : String ;
function  CodePaysDeIso(codeISO : string) : String  ;
Function  FormatZonesRIB(szPays,szQueZone : String ) : String ; //XMG 14/07/03
Function  IBANtoRIB(vstIBAN : String ; var vstRIB, vstPays , vstCleIBAN : String ) : Boolean ;
function  ExisteRibSurCpt( vStCpt, vStRib : String ) : Boolean ;

{Divers}
function  CheckToString ( B : Boolean ) : String ;
function  StringToCheck ( St : String ) : boolean ;
procedure AddChampTableLibreToSQL (lequel: string; var MesChamps : string);
{$IFNDEF EAGLSERVER}
procedure LoadBitMapFromChamp(DS : TDataSet; Champ : string; Image : TImage; IsJpeg : boolean = False; Zoom : integer =0) ;
procedure SetJPEGOptions(Image1: TImage; Zoom : integer = 0);
procedure BlobToFile (VarName:string;Value:Variant) ;
{$ENDIF}
Procedure ControleUsers ;
function  Resolution ( Decim : byte ) : Double ;
Procedure AjouteTOBMontant ( TOBSource,TOBDest : TOB ; NomChamp : String ) ;
function  NullToVide(Valeur : variant) : string;
function StrToAlias (st : string) : string;
function MultiComboInSQL(st: string): string;
function JaiLeDroitNatureGCModif(Nature : string) : boolean;
// CCMX-CEGID - FQ 13344 - DEBUT
//function JaiLeDroitNatureGCCreat(Nature : string) : boolean;
function JaiLeDroitNatureGCCreat(Nature : string; TypePiece : String = 'Origine'; NatureOrig : String = '') : boolean;
// CCMX-CEGID - FQ 13344 - FIN
{Fonctions de conversion}
FUNCTION  PIVOTTOEURO ( X : Double ) : Double ;
FUNCTION  PIVOTTOEURONA ( X : Double ) : Double ;
FUNCTION  EUROTOPIVOT ( X : Double ) : Double ;
FUNCTION  EUROTOPIVOTNA ( X : Double ) : Double ;
FUNCTION  PIVOTTODEVISE(X,Taux,Quotite : Double ; Decim : byte ) : Double ;
FUNCTION  PIVOTTODEVISENA(X,Taux,Quotite : Double ) : Double ;
FUNCTION  DEVISETOPIVOT(X,Taux,Quotite : Double) : Double ;
FUNCTION DEVISETOPIVOTEx(X,Taux,Quotite : Double; NbDec : integer) : Double ;
FUNCTION  DEVISETOPIVOTNA(X,Taux,Quotite : Double) : Double ;
FUNCTION  DEVISETOEURO ( X,Taux,Quotite : Double ) : Double ;
FUNCTION  DEVISETOEURONA ( X,Taux,Quotite : Double ) : Double ;
FUNCTION  EUROTODEVISE ( X,Taux,Quotite : Double ; Decim : integer ) : Double ;
FUNCTION  EUROTODEVISENA ( X,Taux,Quotite : Double ) : Double ;
FUNCTION  DEVISETOFRANC(X,Taux,Quotite : Double) : Double ;
FUNCTION  DEVISETOFRANCNA(X,Taux,Quotite : Double) : Double ;
FUNCTION  FRANCTOEURO ( X : Double ) : Double ;
FUNCTION  FRANCTOEURONA ( X : Double ) : Double ;
FUNCTION  EUROTOFRANC ( X : Double ) : Double ;
FUNCTION  EUROTOFRANCNA ( X : Double ) : Double ;
FUNCTION  DEVISETODEVISE(X,Taux,Quotite,NewTaux,NewQuotite : Double ; Decim : byte ) : Double ;
Procedure AttribCotation ( TOBP : TOB ) ;
function  VarToDouble (V : Variant) : double ;
Function  FabricWhereToken ( sToken,Champ : String ) : String ;
{Créations à la volée}
Function CreerSectionVolee ( CodeSect,Libelle : String ; NumA : integer ) : boolean ; overload;
Function CreerSectionVolee ( CodeSect,Libelle,Axe : String ) : boolean ; overload;
{Limitation paramsoc}
Procedure BrancheParamSocAffiche (Var stVirerBranche, stAfficherBranche : string);
{Blocage réseau}
function _Bloqueur( Niveau : String ; Bloc : boolean  ; QueCloture : Boolean = FALSE) : boolean ;
Function _EstBloque( Niveau : String ; SaufMoi : boolean ) : boolean ;
Function _EstMultiBloque( Niveau : Array of String ; SaufMoi : boolean ) : boolean ;
Function _ExistBlocage : boolean ;
Function  _BlocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) : boolean ;
procedure _DeblocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) ;
function  _BlocCarFiche : boolean ;
function _Blocage ( Niveau : Array of String ; SaufMoi : boolean ; Quoi : String  ; QueCloture : Boolean = FALSE) : boolean ;
{Fin Blocage réseau}

function PGI_IMPORT_BOB(CodeProduit:string): Integer;
Function CompteImpression ( Const typeEtat, natureetat, Codeetat:string): boolean;//mcd 02/02/07

function CleTelForFind (strTelephone:string):string;
function CleTelephone (strTelephone:string; bFixedLen:boolean=TRUE):string;

// $$$JP 18/07/05 - pour l'instant, uniquement pour BUREAU PGI pour le CTI
{$IFDEF BUREAU} //{$IFNDEF ERADIO}
function MajCleTelephone (const strChampCle, strChampTel, strTable:string):boolean;
{$ENDIF}
function  fbToCumulType    ( vFB : TFichierBase ) : String ;
function  fbToTable        ( vFB : TFichierBase ) : String ;

// Gestion MULTISOC
Function  EstMultiSoc : Boolean ;
Function  GetTableDossier  ( vDossier, vNomTable : String ) : String ;
Function  EstTablePartagee ( vNomTable : String ) : Boolean ;
{JP 24/10/07 : gère les paramsoc partagés outre le dossier}
function  GetParamSocDossierMS(Nom : string; Defaut : Variant; Dossier : string) : Variant;
function GetBase (UneBase, LaTable : string) : String;
Function  GetBasesMS (CodeRegroupement : string = ''; BaseSql : boolean = True) : String ;
Function  PresenceMS ( vFichier : String ; vChamp : String ; vValeur : String ) : Boolean ;
Function  PresenceComplexeMS ( vFichier : String ; vChamps : Array of hString ; vComps : Array of hString ; vValeurs : Array of hString ; vTypes : Array of hString ) : Boolean ;
function  GetDossier    ( vSchemaName : String ; CodeRegroupement : String = '' ) : string ;
function  GetSchemaName ( vDossier    : String ; CodeRegroupement : String = '' ) : string ;
Function  OpenSelect ( vSQL : String ; vDossier : String = '' ; vRO : Boolean = True ; vStack : String = '' ) : TQuery ;
{Fonction d'insertion en Multi sociétés des filles de vTob}
function  InsertOrUpdateAllNivelsMS(vTob : TOB; vDossier : string; InsertOk : Boolean) : Boolean;
Function  InsertTobMS( vTob : TOB ; vDossier : String = '' ) : Boolean ;
Function  DeleteTobMS( vTob : TOB ; vDossier : String = '' ) : Boolean ;
Function  UpdateTobMS( vTob : TOB ; vDossier : String = '' ) : Boolean ;
Function  RecupInfosSocietes( vStListeParamSoc : String ; vStCodeRegroupement : string = '' ) : TOB ;
Function  SetParamSocDossier  ( vNomParam : String ; vValeur : Variant  ; vDossier : String = '' ) : Boolean ;

{ DBR : Multi soc }
function EstBaseMultiSoc : boolean;
function TablePartagee (LaTable : string) : boolean;
function GereNoDossier : boolean;
function RechDomZoneLibre (sValue : string; bAbrege : boolean; sPlus : string = ''; bLibre : boolean = false; sPrefixe : string='') : string;
function IsMESSeria: boolean;
function IsMESAutonome: boolean;

{JS1 : Driver SGBD}
function isMssql: boolean; //js1 04052006
function isOracle: boolean;
function isDB2: boolean;

// CA - 11/10/2005 - Gestion base allégée PCL
function EstBasePclAllegee : Boolean;
function EstBasePclOptimisee : Boolean;
{$IFDEF CHR}
function MODEOpenSQL( SQL: string; RO: Boolean; Nb: Integer = -1; Stack: string = ''; UniDirection: boolean = FALSE; stBases: string = '' ) : TQuery;
function RetParamCmdLine( sNomVal: string ): string;
{$ENDIF}
// ME  09/02/2006 Fiche 10327
Function ModePaiementParDefaut (var ModRegle: string) : string;
{JP 31/07/06 : Chargement du NoDossier en mode PGE}
procedure ChargeNoDossier;

{$IFDEF MAJPCL}
procedure InitPGIpourDossierPCL;
{$ENDIF}

{ Gestion des chaines PARAM1=VALEUR1;PARAM2=VALEUR2 }
function GetArgumentValue(Argument: string; Const MyArg : String; const WithUpperCase: Boolean = True; const Separator: String = ';'): String;
function AGLGetArgumentValue(Parms : array of variant; nb : integer): Variant;
function GetArgumentString(Argument: string; Const MyArg : String; WithUpperCase: Boolean = True; const Separator: String = ';'): String;
function GetArgumentInteger(Argument: string; Const MyArg : String; const Separator: String = ';'): Integer;
function GetArgumentDouble(Argument: string; Const MyArg : String; const Separator: String = ';'): Double;
function GetArgumentDateTime(Argument: string; Const MyArg : String; const Separator: String = ';'): tDateTime;
function GetArgumentBoolean(Argument: string; Const MyArg : String; const Separator: String = ';'): Boolean;
function GetArgumentTob(Argument: string; Const MyArg : String; const Separator: String = ';'): Tob;
function SetArgumentTob(const aTobName: String; const aTob: Tob; const aSepar: String = '='): String;

function FindInTobWithLike(LaTobMere : TOB; AuMoins1 : boolean; Champs : Array of string; Valeurs : Array of variant) : TOB;

{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}
Function  VersionBase(libdoss : String): Variant; //js1 120107 on externalise la fonction : vient de majsocparot_tof
{$ENDIF EAGLCLIENT}
{$ENDIF !EAGLSERVER}

// Gestion Identification Bancaire
function IsBBAN(BBAN : string) : boolean;
function IsIBAN(IBAN : string) : boolean;

function GetValueMontant(Montant : double) : string;

{ Pour les appels à Isoflex }
function GetParamIsoflex : string;

function GetSelectAll(PrefixesTables : string; SansBlob : boolean=True; ExclureChamps : string='') : string;
function GetParamSocPaysNumerique : string;
function CodePays3CVersNumerique(CodePays: String): String;
function GetSousSectionUnique(Ax : string) : string;

    function FabriqueConditionIn(Zone: string): string;
CONST
  MS_CODEREGROUPEMENT : String = '##MULTISOC' ;  // Code du regroupement multi-société


type
  TAbstractionRepertoire = (rpUserData, rpUserTemp, rpCommonData,rpBob, rpCegidData);

// SBO 19/11/2005 : Centralisation gestion du multi-axe analytique
{b FP}
type
  TSQLAnaCroise = class
  private
    FAxes: array[1..MaxAxe] of Boolean; // Axes ventilables ?
    FPremierAxe: Integer;

    procedure LoadInfo;
    function  GetPremierAxe: String;
    function  AxeToSousPlan(NatureCpt: String): Integer;  {A partir de l'axe, retourne le numéro du sous plan}
  public
    constructor Create;
    destructor  Destroy; override;

    function GetConditionAxe(NatureCpt: String): String;
    function GetChampSection(NatureCpt: String): String;

    class function  ConditionAxe(NatureCpt: String): String;
    class function  ChampSection(NatureCpt: String): String;

    class procedure TraduireRequete( NatureCpt: String ; var vStReq : String ) ;

  end;
{e FP}

function AbstractionRepertoire(Genre: TAbstractionRepertoire; Complement: String; Produit: Boolean): String;
function RepertoireExiste (LeRepertoire :string; bCreation : boolean = false) : boolean;
{$IFNDEF EAGLSERVER}
function FormatNumVersion(const Mask: String): String;
{$ENDIF !EAGLSERVER}
function GetMyTempPath : string;
function GetMyComputerName : string;

implementation
                
uses
{$IFDEF AFFAIRE}
  {$IFDEF GPAO}
  wAffaires,
  {$ENDIF GPAO}
  ConfidentAffaire ,
{$ENDIF AFFAIRE}
{$IFNDEF EAGLSERVER}
 wCommuns,
{$ENDIF EAGLSERVER}
  CbpMCD,
  CbpEnumerator,
 UProcGen  {StrRight}
 , CommonTools
 ;


{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 24/10/2006
Modifié le ... :   /  /    
Description .. : permet de vérifier 
Suite ........ : - la validité de la clé pour la belgique
Suite ........ : - le n° TVA et le siret pour la France
Mots clefs ... : TVA
*****************************************************************}
function VerifTVA(stIsoPays, num : string) : boolean ;
var stNum,stCle : string ;
    lgNum,NumPar97 : integer ;
begin
  stNum := stISOPays+num ;
  lgNum := length(num);
  result := False ;
    (*if (stISOPays='FR') and VerifSiret(num) and isIntraComValide(stNum,stNum) then
      result := true  //Cas français, on teste tout !
    else *)
      if isIntraComValide(stNum,stNum) then //cas autre
        if stISOPays = 'BE' then  //algo particulier pour la belgique
        begin                     //on vérifie la clé du n° TVA
          stCle:=copy(num, lgNum-1, length(num)); //les 2 dernier digit sont la clé
          stNum := copy(num,1,length(num)-2);
          numPar97 := trunc(strToInt(stNum) / 97)*97 ;
          if (97 - (StrToInt(stNum)-numPar97)) = strtoint(stCle) then
            result := true
          else
            result := false;
        end //if 'BE'
      else
        result := false ; // isIntraComValide
end;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 23/10/2006
Modifié le ... : 23/10/2006
Description .. : isIntraComValide : Teste que la structure du numéro de
Suite ........ : TVA intra communautaire soit valide.
Suite ........ : // num : numéro de TVA intra Com.
Suite ........ : // [siren] : cas pour la France afin d'afiner la vérification
Suite ........ : // Return : True/False
Mots clefs ... : TVA,INTRACOM
*****************************************************************}
Function isIntraComValide(num : String; var siren : String) : Boolean;
var
  //Découpage des données
  pays : String ;
  fin : String ;
  finLength : Integer ;
begin
  Result := false;
  //Nettoyer la chaine
  num := UpperCase(StringReplace(num, ' ', '',[rfReplaceAll])) ;
  If Length(num) < 3 Then  //Il n'existe pas de code de moins de 3 char
  begin
    isIntraComValide := False ;
    Exit;
  end ;
  pays := copy(num, 1, 2) ;  //pays pays
  fin := copy(num, 3, length(num)); //Num TVA
  finLength := Length(num) - 2; //longueur num - 2 code ISO Pays
  //Vérifier à partir de la clé Pays
  if pays = 'IE' then if (finLength = 8) then result := true ; //Irlande
  if pays = 'DK' then if (finLength = 8)and  IsNumeric(fin) then result := true ; //Danemark
  if pays = 'FI' then if (finLength = 8)and  IsNumeric(fin) then result := true; //Finlande
  if pays = 'LU' then if (finLength = 8)and  IsNumeric(fin) then result := true; //Luxembourg
  if pays = 'MT' then if (finLength = 8)and  IsNumeric(fin) then result := true; //Malte
  if pays = 'SI' then if (finLength = 8)and  IsNumeric(fin) then result := true; //Slovénie
  if pays = 'HU' then if (finLength = 8)and  IsNumeric(fin) then result := true; //Hongrie
  if pays = 'CZ' then if (finLength >= 8) And (finLength <= 10) then result := true; //République tchèque
  if pays = 'ES' then if (finLength = 9) then result := true; //Espagne
  if pays = 'CY' then if (finLength = 9) then result := true; //Chypre : CY + 8 caractères numériques + 1 caractères alphabétiques
  if pays = 'DE' then if (finLength = 9)and  IsNumeric(fin) then result := true; //Allemagne
  if pays = 'EL' then if (finLength = 9)and  IsNumeric(fin) then result := true; //Grèce
  if pays = 'GR' then if (finLength = 9)and  IsNumeric(fin) then result := true; //Grèce
  if pays = 'PT' then if (finLength = 9)and  IsNumeric(fin) then result := true; //Portugal
  if pays = 'EE' then if (finLength = 9)and  IsNumeric(fin) then result := true; //Estonie

  if pays = 'SK' then if ((finLength = 9) Or (finLength = 10)) and  IsNumeric(fin) then result := true; //Slovaquie
  if pays = 'LT' then if ((finLength = 9) Or (finLength = 12)) and  IsNumeric(fin) then result := true; //Lituanie
  if pays = 'GB' then if (((finLength = 9) Or ((finLength = 4) And (copy(fin, 1, 1) = '9'))))
                              Or ((finLength = 5) And (IsNumeric(copy(fin, 3, length(pays)))))
                      then result := true; //Grande-Bretagne

  if pays = 'AT' then if (finLength = 9) And ((copy(fin, 1, 1) = 'U') And IsNumeric(copy(fin, 2,length(pays))))
                        then result := true; //Autriche

  if pays = 'PL' then if (finLength = 10)and  IsNumeric(fin) then result := true; //Pologne

  if pays = 'BE' then if ((finLength = 9) and IsNumeric(fin)) then result := true; //Belgique
  if pays = 'BE' then if ((finLength = 10) and (copy(fin, 1, 1) = '0') and IsNumeric(fin)) then result := true; //Belgique

  if pays = 'LV' then if (finLength = 11) and IsNumeric(fin) then result := true; //Lettonie
  if pays = 'IT' then if (finLength = 11) and IsNumeric(fin) then result := true; //Italie
  if pays = 'FR' then if (finLength = 11) and IsNumeric(fin) then
                        if (length(siren)<>0) and (copy(fin, 3,length(pays)) = siren)
                        then
                          result := true; //France
  if pays = 'NL' then if (finLength = 12)and  IsNumeric(fin) then result := true; //Pays-Bas
  if pays = 'SE' then if (finLength = 12)and  IsNumeric(fin) then result := true; //Suède

End ;
Function Evaluedate ( St : String ) : TDateTime ;
Var dd,mm,yy : Word ;
BEGIN
Result:=0 ; if St='' then Exit ;
dd:=StrToInt(Copy(St,1,2)) ; mm:=StrToInt(Copy(St,3,2)) ; yy:=StrToInt(Copy(St,5,4)) ;
Result:=Encodedate(yy,mm,dd) ;
END ;

Function CompteImpression ( Const typeEtat, natureetat, Codeetat:string): boolean;
begin  //mcd 02/02/2007 13372 AGL. on ne compte que les impressions qui sont paramétré dans la base
  // PCH 07/02/2007 ajout critère yed_predefini="STD"
  Result:=false;
  if ExisteSql ('select yed_natureetat from yparamedition'
     + ' where yed_natureetat="' + NatureEtat + '" and yed_Codeetat="'+ CodeEtat
     + '" and yed_typeetat="' + TypeEtat + '" and yed_predefini="STD"') then result:=true;
end;

{$IFNDEF NOVH}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/03/2000
Modifié le ... : 27/07/2007
Description .. : Encodage d'un identifiant de pièce Gescom dans l'écriture Comptable
                 depuis une Tob PIECE ou LIGNE
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Function EncodeRefCPGescom (LaTOB : TOB; FromAcc : boolean=false) : String ;
var Prefixe, LaDate : string;
begin
  Result:='' ;
  if not assigned(LaTob) then Exit ;
  { Si c'est une Tob acomptes, il faut prendre la date GP }
  if not FromAcc then
  begin
    if LaTob.NomTable = 'LIGNE' then
      Prefixe := 'GL_'
      else
      Prefixe := 'GP_';
    LaDate := FormatDateTime('ddmmyyyy', LaTob.GetValue(Prefixe + 'DATEPIECE'));
  end else
  begin
    Prefixe := 'GAC_';
    LaDate := FormatDateTime('ddmmyyyy', LaTob.GetValue('GP_DATEPIECE'));
  end;
  Result := LaTob.GetString(Prefixe + 'NATUREPIECEG')
          + ';' + LaTob.GetString(Prefixe + 'SOUCHE')
//          + ';' + FormatDateTime('ddmmyyyy',LaTob.GetValue(Prefixe + 'DATEPIECE'))
          + ';' + LaDate
          + ';' + IntToStr(LaTob.GetInteger(Prefixe + 'NUMERO'))
          + ';' + IntToStr(LaTob.GetInteger(Prefixe + 'INDICEG'))+';';
end;

Function EncodeRefPresqueCPGescom ( TOBPiece : TOB ) : String ;    overload;
BEGIN
Result:='' ; if TOBPiece=Nil then Exit ;
Result:=TOBPiece.GetValue('GP_NATUREPIECEG')+';'+TOBPiece.GetValue('GP_SOUCHE')+';'+'%'+';'
       +IntToStr(TOBPiece.GetValue('GP_NUMERO'))+';'+IntToStr(TOBPiece.GetValue('GP_INDICEG'))+';' ;
END ;

Function EncodeRefPresqueCPGescom ( Cledoc : r_cledoc ) : String ; overload;
BEGIN
Result:='' ; if Cledoc.NaturePiece ='' then Exit ;
Result:=Cledoc.NaturePiece +';'+Cledoc.Souche +';'+'%'+';'
       +IntToStr(Cledoc.NumeroPiece )+';'+IntToStr(Cledoc.Indice )+';' ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/03/2000
Modifié le ... : 27/03/2000
Description .. : Décodage d'un identifiant de pièce Gescom en Compta vers un identifiant Gescom
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Function DecodeRefCPGescom ( RefG : String ) : R_CleDoc ;
Var CD : R_CleDoc ;
    StC  : String ;
BEGIN
FillChar(CD,Sizeof(CD),#0) ;
StC:=ReadTokenSt(RefG) ; CD.NaturePiece:=Trim(StC) ;
StC:=ReadTokenSt(RefG) ; CD.Souche:=Trim(StC) ;
StC:=ReadTokenSt(RefG) ; CD.DatePiece:=EvalueDate(StC) ;
StC:=ReadTokenSt(RefG) ; CD.NumeroPiece:=StrToInt(StC) ;
StC:=ReadTokenSt(RefG) ; CD.Indice:=StrToInt(StC) ;
Result:=CD ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 14/09/2005
Description .. : Charge un R_CleDoc à partir d'une chaine du type :

  (Mix de l'ancien DecodeRefPiece et de StringToCleDoc)

  [DATEPIECE1;]NATURE;SOUCHE;NUMERO;INDICEG;[NUMLIGNE]; (Ancien GL_PIECEPRECEDENTE)
  ou
  NATURE;[DATEPIECE2];SOUCHE;NUMERO;INDICEG;[NUMLIGNE];

  Format de DATEPIECE1 : jjmmaaaa (ecrite par l'ancien EncodeRefPiece)
  Format de DATEPIECE2 : Date au format 'normale'

*****************************************************************}
procedure DecodeRefPiece(St: string; var CleDoc: R_CleDoc);
var
  StC, StL: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  if St <> '' then
  begin
    StC := St;
    { Date pièce en première position }
//GP_20080318_TS_GC15906
    if Stc[1] in ['0'..'9'] then
    begin
    { Decode jjmmaaaa }
      CleDoc.DatePiece := EvalueDate(ReadTokenSt(StC))
    end;

    { Nature de la pièce en première ou deuxième position }
    CleDoc.NaturePiece := Trim(ReadTokenSt(StC));

    { Date pièce en deuxième position }
    // mcd 07/12/2005 if Stc[1] in ['0','1','2','3','4','5','6','7','8','9'] then
    // DBR pour expliquer :
    // ne pas considérer que c'est une date si souche numérique.
    // cette structure n'est possible que sur un AppelPiece (a changer la clé d'AppelPiece)
    // c'est une date si / en 3ème possition ou
    // le premier, le troisième et le quatrième caractère après la nature sont numériques.
//GP_20080318_TS_GC15906 >>>
    if (Stc<>'') and
       (Stc[3] = DateSeparator) or
       ((Stc[1] in ['0'..'9']) and
        (Stc[3] in ['0'..'9']) and
        (Stc[4] in ['0'..'9'])) then
      CleDoc.DatePiece := StrToDate(ReadTokenST(StC));
//GP_20080318_TS_GC15906 <<<

    { Souche }
    CleDoc.Souche := Trim(ReadTokenSt(StC));

    { Numéro de pièce }
    CleDoc.NumeroPiece := StrToInt(ReadTokenSt(StC));

    { Indice }
    CleDoc.Indice := StrToInt(ReadTokenSt(StC));

    { Numéro de ligne ou d'ordre }
    StL := ReadTokenSt(StC);
    if StL <> '' then
    begin
      CleDoc.NumLigne := StrToInt(StL);
      CleDoc.NumOrdre := StrToInt(StL);
    end;
  end;
end;

procedure AGLDecodeRefPiece(Parms : Array of Variant; nb: integer);
var
  s: String;
  CleDoc: R_CleDoc;
begin
  s := Parms[0]; { La chaine à décoder }
  DecodeRefPiece(s, CleDoc);
  Tob(LongInt(Parms[1])).AddChampSupValeur('NATUREPIECEG', CleDoc.NaturePiece);
  Tob(LongInt(Parms[1])).AddChampSupValeur('SOUCHE', Cledoc.Souche);
  Tob(LongInt(Parms[1])).AddChampSupValeur('NUMERO', Cledoc.NumeroPiece);
  Tob(LongInt(Parms[1])).AddChampSupValeur('INDICEG', Cledoc.Indice);
  Tob(LongInt(Parms[1])).AddChampSupValeur('NUMORDRE', Cledoc.NumOrdre); { C'est la même chose cf DecodeRefPiece }
  Tob(LongInt(Parms[1])).AddChampSupValeur('NUMLIGNE', Cledoc.NumLigne); { C'est la même chose cf DecodeRefPiece }
  Tob(LongInt(Parms[1])).AddChampSupValeur('NOPERSP', Cledoc.NoPersp);
end;

function EncodeRefGCCpta(TobEcr : TOB) : string;
begin
  Result := '' ;
  if not assigned(TobEcr) then exit;
  Result := TobEcr.GetString('E_JOURNAL')+';'+FormatDateTime('ddmmyyyy',TobEcr.GetDateTime('E_DATECOMPTABLE'))+';'
           +IntToStr(TobEcr.GetInteger('E_NUMEROPIECE'))+';'+TobEcr.GetString('E_QUALIFPIECE')+';'
           +TobEcr.GetString('E_NATUREPIECE')+';' ;
end;

{ Fonctions de manipulation d'un TClesDocs }
{ Ajout }
procedure AddCleDocInTClesDocs(const CleDoc: R_CleDoc; var ClesDocs: TClesDocs);
begin
  SetLength(ClesDocs, Length(ClesDocs) + 1);
  ClesDocs[Pred(Length(ClesDocs))] := CleDoc
end;

procedure AddCleDocInTClesDocs(const RefPiece: String; var ClesDocs: TClesDocs);
var
  CleDoc: R_CleDoc;
begin
  DecodeRefPiece(RefPiece, CleDoc);
  AddCleDocInTClesDocs(CleDoc, ClesDocs)
end;

{ Test d'existence }
function ExistCleDocInTClesDocs(const CleDoc: R_CleDoc; ClesDocs: TClesDocs; const TestNumOrdre: Boolean): Boolean;
var
  i: Integer;
begin
  Result := False;
  i := Low(ClesDocs);
  while not Result and (i <= High(ClesDocs)) do
  begin
    Result := (CleDoc.NaturePiece = ClesDocs[i].NaturePiece) and
              (CleDoc.NumeroPiece = ClesDocs[i].NumeroPiece) and
              (CleDoc.Souche      = ClesDocs[i].Souche     ) and
              (CleDoc.Indice      = ClesDocs[i].Indice     ) and
              (not TestNumOrdre or (CleDoc.NumOrdre = ClesDocs[i].NumOrdre));
    Inc(i)
  end
end;

function ExistCleDocInTClesDocs(const RefPiece: String; ClesDocs: TClesDocs; const TestNumOrdre: Boolean): Boolean;
var
  CleDoc: R_CleDoc;
begin
  DecodeRefPiece(RefPiece, CleDoc);
  Result := ExistCleDocInTClesDocs(CleDoc, ClesDocs, TestNumOrdre)
end;
{$ENDIF !NOVH}

function CanCloseExoFromGC(Exercice : string) : boolean ;
var Nb : integer;
    Qry : TQuery;
    DteDeb, DteFin : TDateTime;
begin
  Result := False;
  if Exercice = '' then exit;
  Qry := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE = "' + Exercice + '"', True);
  if not Qry.Eof then
  begin
    DteDeb := Qry.Fields[0].AsDateTime;
    DteFin := Qry.Fields[1].AsDateTime;
    Ferme(Qry);
  end else
  begin
    Ferme(Qry);
    exit;
  end;
  // Recherche dans COMPTADIFFEREE
  Qry := OpenSQL('SELECT COUNT(GCD_DATEPIECE)'
               + ' FROM COMPTADIFFEREE'
               + ' WHERE GCD_DATEPIECE >= "' + UsDateTime(DteDeb) + '"'
               + ' AND GCD_DATEPIECE <= "' + UsDateTime(DteFin) + '"'  , True);
  if not Qry.EOF then
    Nb := Qry.Fields[0].AsInteger
    else
    Nb := 0;
  Ferme(Qry);
  if Nb > 0 then
    exit;

  { CA - 29/11/2006 - A la demande de BTP, on ne fait pas le contrôle suivant pour le marché BTP }
  { CA - 04/01/2007 - FQ 19399  Idem en PCL }
  if ((GetParamSocSecur('SO_DISTRIBUTION','')='014') or (ctxPCL in V_PGI.PGIContexte)) then
  begin
    Result := True;
    exit;
  end;

  // Recherche les pièces
  Qry := OpenSQL('SELECT COUNT(GP_DATEPIECE)'
               + ' FROM PIECE, PARPIECE'
               + ' WHERE GP_VIVANTE = "X"'
               + ' AND GP_DATEPIECE >= "' + UsDateTime(DteDeb) + '"'
               + ' AND GP_DATEPIECE <= "' + UsDateTime(DteFin) + '"'
               + ' AND GPP_NATUREPIECEG = GP_NATUREPIECEG AND GPP_TYPEECRCPTA <> "RIE" AND GPP_TYPEECRCPTA <> ""', true);
  if not Qry.EOF then
    Nb := Qry.Fields[0].AsInteger
    else
    Nb := 0;
  Ferme(Qry);
  if Nb > 0 then
    exit;
  Result := True;
end;

{$IFNDEF NOVH}
{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 20/11/2005
Modifié le ... :
Description .. : Est-il possible de valider une pièce venant de la GC
                 Doit être en simulation
                 Doit être la dernière pièce de la chaine (ex FAC, FF, AVC, ...)
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
function CanValidPieceFromGC(RefGescom : string) : boolean;
var CleDoc : R_CleDoc;
    Sql : String;
begin
  Result := true;
  if RefGescom = '' then exit;
  // BDU - 25/04/07. Pas de traitement si lien Frais en compta
  if (Copy(RefGescom, 1, 3) = 'AA;') or (Copy(RefGescom, 1, 3) = 'AD;') or
    (Copy(RefGescom, 1, 3) = 'AR;') then Exit;
  CleDoc := DecodeRefCPGescom(RefGescom);
  if (CleDoc.NaturePiece = '') or (CleDoc.Souche = '') then exit;
  Sql := 'SELECT GPP_NATUREPIECEG, GPP_SOUCHE FROM PARPIECE P1'
       + ' WHERE P1.GPP_NATURESUIVANTE IN("", P1.GPP_NATUREPIECEG||";")'
       + ' AND P1.GPP_SOUCHE = "' + CleDoc.Souche + '"'
       + ' AND P1.GPP_NATUREPIECEG = "' + CleDoc.NaturePiece + '"'
       + ' AND NOT EXISTS (SELECT 1 FROM PARPIECE P2'
       + '                  WHERE P2.GPP_NATUREORIGINE LIKE ("%"||P1.GPP_NATUREPIECEG||";%"))';
  Result := (ExisteSQL(Sql));
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 12/05/2006
Modifié le ... :
Description .. : Mise à jour de GP_REFCOMPTABLE en validation de pièce
Suite ........ : (passage de Cpta simulation vers Cpta normale)
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
function MajRefCptaOnValidPce(TobE : TOB) : boolean;
var RefGC, RefCP : string;
    CleDoc : R_CleDoc;
begin
  Result := true;
  if not assigned(TobE) then exit;
  { Doit être une écriture de type Simulation
    JTR le 23/01/2008 Non, car le qualifpiece est égal à N }
//  if TobE.GetString('E_QUALIFPIECE') <> 'S' then exit;
  { Doit être une facture ou un avoir }
  if pos(TobE.GetString('E_NATUREPIECE') + ';', 'FC;AC;FF;AF;') = 0 then exit;
  { Doit avoir une référence Gescom }
  RefGC := TobE.GetString('E_REFGESCOM');
  if RefGC = '' then exit;
  CleDoc := DecodeRefCPGescom(RefGC);
  RefCP := EncodeRefGCCpta(TobE);
  { Doit mettre à jour une seule pièce }
  Result := (ExecuteSql('UPDATE PIECE SET GP_REFCOMPTABLE = "'+ RefCP +'" WHERE '
                      + ' GP_NATUREPIECEG="' + CleDoc.NaturePiece + '"'
                      + ' AND GP_SOUCHE="' + CleDoc.Souche + '"'
                      + ' AND GP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
                      + ' AND GP_INDICEG=' + IntToStr(CleDoc.Indice)) = 1);
end;
{$ENDIF !NOVH}
{===================== Début Blocage Reseau ==========================}
function _Bloqueur( Niveau : String ; Bloc : boolean  ; QueCloture : Boolean = FALSE) : boolean ;
BEGIN
//{$IFDEF GCGC}
//if GetParamSoc('SO_IFDEFCEGID') then
// CA - 09/09/2005 - Pas de SO_IFDEFCEGID dans les bases PCL
if GetParamSocSecur('SO_IFDEFCEGID',False) then
   if QueCloture and (Not Bloc) then Niveau:='nrCloture' ;
//{$ENDIF}
Result:=Bloqueur(Niveau,Bloc) ;
END ;

Function _EstBloque( Niveau : String ; SaufMoi : boolean ) : boolean ;
BEGIN
Result:=EstBloque( Niveau,SaufMoi) ;
END ;

Function _EstMultiBloque( Niveau : Array of String ; SaufMoi : boolean ) : boolean ;
BEGIN
Result:=EstMultiBloque( Niveau,SaufMoi) ;
END ;

Function _ExistBlocage : boolean ;
BEGIN
Result:=ExistBlocage ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 06/04/2005
Modifié le ... :   /  /
Description .. : - LG - 06/04/2005 - suppression du VH^
Mots clefs ... :
*****************************************************************}
Function  _BlocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) : boolean ;
BEGIN
If TypeBlocage='' Then TypeBlocage:='NRTOUTSEUL' ;
//{$IFDEF GCGC}
//if GetParamSoc('SO_IFDEFCEGID') then
// CA - 09/09/2005 - Pas de SO_IFDEFCEGID dans les bases PCL
if GetParamSocSecur('SO_IFDEFCEGID',False) then
   if Shunte then BEGIN Result:=TRUE ; Exit ; END ;
//{$ENDIF}
Result:=BlocageMonoPoste (Totale) ;
END ;

procedure _DeblocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) ;
BEGIN
If TypeBlocage='' Then TypeBlocage:='NRTOUTSEUL' ;
//{$IFDEF GCGC}
//if GetParamSoc('SO_IFDEFCEGID') then
// CA - 09/09/2005 - Pas de SO_IFDEFCEGID dans les bases PCL
if GetParamSocSecur('SO_IFDEFCEGID',False) then
   if Shunte then Exit ;
//{$ENDIF}
DeblocageMonoPoste (Totale) ;
END ;

function  _BlocCarFiche : boolean ;
BEGIN
//{$IFDEF GCGC}
//if GetParamSoc('SO_IFDEFCEGID') then
// CA - 09/09/2005 - Pas de SO_IFDEFCEGID dans les bases PCL
if GetParamSocSecur('SO_IFDEFCEGID',False) then
  Result:=_EstMultiBloque(['nrCloture'],True)
else
//{$ENDIF}
  Result:=_EstMultiBloque(['nrCloture','nrBatch','nrSaisieCreat','nrSaisieModif','nrLettrage'],True) ;
END ;

Function NiveauAPrendre(Niveau : String) : Boolean ;
BEGIN
Result:=FALSE ;
if (Niveau<>'nrEnca') and (Niveau<>'nrDeca') and (Niveau<>'nrLettrage') and (Niveau<>'nrBatch') and (Niveau<>'nrSaisieModif') And
   (Niveau<>'nrPointage') and (Niveau<>'nrSaisieCreat') and (Niveau<>'nrRelance') Then Result:=TRUE ;
END ;

function _Blocage ( Niveau : Array of String ; SaufMoi : boolean ; Quoi : String ; QueCloture : Boolean = FALSE) : boolean ;
VAr N :  Array of String ;
    i,j,k : Integer ;
    NrAEnlever : Boolean ;
BEGIN
//{$IFDEF GCGC}
//if GetParamSoc('SO_IFDEFCEGID') then
// CA - 09/09/2005 - Pas de SO_IFDEFCEGID dans les bases PCL
if GetParamSocSecur('SO_IFDEFCEGID',False) then
begin
  if QueCloture then
    BEGIN
    SetLength(N,1) ;
    N[0]:='nrCloture' ;
    Result:=Blocage( N,SaufMoi,Quoi) ; Exit ;
    END Else
    BEGIN
    i:=0 ; NrAEnlever:=FALSE ;
    for j:=0 to High(Niveau) do if NiveauAPrendre(Niveau[j]) then Inc(i) else nrAEnlever:=TRUE ;
    If NrAEnlever Then
      BEGIN
      SetLength(N,i) ; k:=0 ;
      for j:=0 to High(Niveau) do if NiveauAPrendre(Niveau[j]) then
        BEGIN
        N[k]:=Niveau[j] ;
        Inc(k) ;
        END ;
      Result:=Blocage( N,SaufMoi,Quoi) ; Exit ;
      END ;
    END ;
end;
//{$ENDIF}
Result:=Blocage( Niveau,SaufMoi,Quoi) ;
END ;
(************ FIN BLOCAGE RESEAU **********************)

Function  ExisteCarInter ( St : String ) : boolean ;
BEGIN
Result:=True ;
if Pos('"',St)>0 then Exit ; if Pos('*',St)>0 then Exit ;
if Pos('%',St)>0 then Exit ; if Pos('_',St)>0 then Exit ;
if Pos('?',St)>0 then Exit ; if Pos('''',St)>0 then Exit ;
Result:=False ;
END ;

function  Resolution ( Decim : byte ) : Double ;
Var i : integer ;
    X : double ;
BEGIN
X:=1.0 ; for i:=1 to Decim do X:=X/10.0 ; Resolution:=X ;
END ;

Procedure AjouteTOBMontant ( TOBSource,TOBDest : TOB ; NomChamp : String ) ;
Var im : integer ;
    X : Double ;
BEGIN
if ((TOBSource=Nil) or (TOBDest=Nil)) then Exit ;
im:=TOBSource.GetNumChamp(NomChamp) ; if im<0 then Exit ;
X:=TOBSource.GetValeur(im)+TOBDest.GetValeur(im) ; TOBDest.PutValeur(im,X) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 31/10/2002
Modifié le ... : 25/11/2002
Description .. : Formatage d'une chaine quelconque en une chaine
Suite ........ : valide pour les alias des requêtes SQL ;
Suite ........ : carctères valides : lettres ; '_' et chiffres si pas en 1ère
Suite ........ : posistion
Mots clefs ... : ALIAS;
*****************************************************************}
function StrToAlias (st : string) : string;
var i,j : integer;
    maj, stAlias : string;
begin
  j := 1 ;
  stAlias := '' ;
  maj := trim (uppercase (st));
  // si le 1er caractère est un chiffre on le précède par un '_'
  if maj [1] in ['0'..'9','_'] then
  begin
    stAlias := stAlias + 'X' + maj [1] ;
    j := 2 ;
  end ;
  for i := j to length (st) do
  begin
    if maj[i] in ['A'..'Z','_','0'..'9'] then stAlias := stAlias + st[i]
    // remplacement du caractère incorrect par un _ si le carctère précédent n'en pas '_'
    else if maj [i-1] <> '_' then stAlias := stAlias + '_';
  end;
  result := stAlias ;
end;

function MultiComboInSQL(st: string): string;
var Liste, Code: string;
begin
  Liste := '';
  while st <> '' do
  begin
    Code := ReadTokenSt(st);
    if Liste <> '' then Liste := Liste + ',';
    Liste := Liste + '"' + Code + '"';
  end;
  Result := Liste;
end;

function NullToVide(Valeur : variant) : string;
begin
  if VarIsNull(Valeur) then Result:='' else Result:=Valeur;
  if Result = #0 then Result :=''; //mcd 14/03/2007 en delphi7, null dans champsup des bob pas triaté pareil (equalité cbp 13602
end;

function CheckToString ( B : Boolean ) : String ;
begin
  Result := booltostr_(B);
end;

function StringToCheck ( St : String ) : boolean ;
begin
  Result := strtobool_(St)
end;

{***********A.G.L.***********************************************
Auteur  ...... : EPZ-PL
Créé le ...... : 07/06/2005
Modifié le ... :   /  /
Description .. : Procédure qui ajoute une champ de table à la liste MesChamps si ce champ est activé
Suite ........ : dans les paramètres (<> '.-') (champs tables libres)
Suite ........ : Sert à remplir des requêtes avec seulement les champs des tables libres activés
Mots clefs ... : SQL; GIGA
*****************************************************************}
procedure AddChampTableLibreToSQL (lequel: string; var MesChamps : string);
var titre : string;
begin
  titre := ChampToLibelle(lequel);
  if (Trim(titre) <> '') and (Length(titre) = 2) and (titre = '.-') then exit;
  MesChamps := MesChamps + ',' + lequel;
end;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 29/01/2003
Modifié le ... :   /  /
Description .. : Fonction qui retourne le code ISO d'un code pays passé en
Suite ........ : paramètre
Mots clefs ... : PAYS;ISO
*****************************************************************}
function CodeIsoDuPays(Pays : string) : String  ;
var Q : TQuery ;
  SQL : String ;
  codeISO : string ;
begin
  result := '';
  SQL := 'SELECT PY_CODEISO2 FROM PAYS WHERE PY_PAYS="' + Pays + '"' ;
  Q := OpenSQL(SQL,True);
  if not Q.EOF then CodeISO := UpperCase(Q.Fields[0].AsString); // FQ 13197
  ferme(Q);
  if codeISO>'' then
    result := codeISO
  else
  result:=CodeISOFR; //Par défaut on retourne le code ISO de la France //XMG 14/07/03
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 29/01/2003
Modifié le ... :   /  /
Description .. : retourne le code pays par rapport à un code ISO
Mots clefs ... : PAYS;ISO
*****************************************************************}
function CodePaysDeIso(codeISO : string) : String  ;
var Q : TQuery ;
  SQL : String ;
  Pays : string ;
begin
  if trim(CodeISO)='' then
     CodeISO:=CodeISOFR ; //XMG 14/07/03
  result := '';
  SQL := 'SELECT PY_PAYS FROM PAYS WHERE PY_CODEISO2="' + codeISO + '"' ;
  Q := OpenSQL(SQL,True);
  if not Q.EOF then Pays := Q.Fields[0].AsString;
  ferme(Q);
  if Pays>'' then
    result := Pays
  else
    result:='FRA'; //Par défaut on retourne le code pays de la France
end;

function EncodeRIB ( Etab,Guichet,Numero,Cle,Dom : String ; ISOPays : String = CodeISOFR ; TypePays : String ='' ) : String ; //XMG 14/07/03
var TRIB : tob ;
  stRetour, lequel : string ;
  i : integer ;
  st : string ;
BEGIN
    TRIB := Tob.Create('YIDENTBANCAIRE',nil,-1);
    Try
    stRetour := ParametreRIB( nil, ISOPays, TypePays, '',  TRIB, false) ;
    // FQ 19469
    if stRetour = 'VIDE' then
    begin
    	// Il n'y a pas de paramétrage pour ce pays.
            {if ISOPays = codeISOFR then
      	result := etab + '/' + guichet + '/' + numero + '/' + cle + '/' + Dom
      else if ISOPays = codeISOES then
 				result := etab + '/' + guichet + '/' + cle + '/' + numero + '/' + Dom
      }
      Result := encodeRIB_OLD(Etab,Guichet, Numero, cle, dom, ISOPays);
    end // END FQ 19469
    else
    begin
	    st := '' ;
  	  for i:=1 to 4 do
    	begin
        if i>1 then st := st + '/' ;
      	lequel := chercheLequel(i, TRIB);
	      if lequel='ETABBQ' then
  	      st:= st + Format_String(etab, cherchelong(lequel, ISOPays, Typepays)) else
    	  if lequel='GUICHET' then
      	  st := st +Format_String(Guichet, cherchelong(lequel, ISOPays, Typepays))  else
	      if lequel='NUMEROCOMPTE' then
  	      st := st + Format_String(Numero, cherchelong(lequel, ISOPays, Typepays))  else
    	  if lequel='CLERIB' then
      	  st := st + Format_String(cle, cherchelong(lequel, ISOPays, Typepays));
	    end ;
  	  st := st + '/' + Dom;
     // 	freeAndNil(TRIB);  gm le 21/12/07 déplacé
	    //SDA le 27/12/2007 result := copy(st ,2,length(st)) ;
      result := st; // SDA le 27/12/2007
      end;
    Finally
    FreeAndNil(TRIB); // gm le 21/12/07 déplacé
    end;
END ;

function EncodeRIB_OLD ( Etab,Guichet,Numero,Cle,Dom : String ; ISOPays : String = CodeISOFR ) : String ; //XMG 14/07/03
Var St : String ;
BEGIN
   if IsoPays=CodeISOES then //RIB au Format Espagnol
      St:=Format_String(Etab,4)+'/'+Format_String(Guichet,4)+'/'+Format_String(Cle,2)+'/'+Format_String(Numero,10)+'/'+Format_String(Dom,24)
   else //La reste (y compris la France)
      St:=Format_String(Etab,5)+'/'+Format_String(Guichet,5)+'/'+Format_String(Numero,11)
         +'/'+Format_String(Cle,2)+'/'+Format_String(Dom,24) ; // XVI 24/02/2005
   if ((Trim(Etab)='') or (Trim(Guichet)='') or (Trim(Numero)='') or (Trim(Cle)='')) then St:=Format_String('',Length(St)) ;
   Result:=St ;
END ;

Procedure DecodeRIB ( Var Etab,Guichet,Numero,Cle,Dom : String ; RIB : String ; ISOPays : String = CodeISOFR ) ; //XMG 14/07/03
  var TRIB : tob ;
  stRetour, lequel : string ;
  lg : integer ;
  i, debS : integer ;
BEGIN
    //initialisation variable de retour
    Etab := '';
    Guichet := '';
    Numero := '';
    Cle := '';
    Dom := '' ;
    TRIB := Tob.Create('YIDENTBANCAIRE',nil,-1);
    stRetour := ParametreRIB(nil, ISOPays, '', '',  TRIB, false) ;
    //on fait si le retour contient quelque chose
    if (length(trim(stRetour))<>0) and (trim(stRetour)<>'VIDE') then
    begin
      debS := 1 ;
      for i:=1 to 4 do
      begin
        lequel := chercheLequel(i, TRIB);
        lg := strtoint(TRIB.GetValue('YIB_LG' + lequel)) ;
        if lequel='ETABBQ' then
          Etab:=Copy(RIB,debS, lg) else
        if lequel='GUICHET' then
          Guichet:=Copy(RIB,debS, lg) else
        if lequel='NUMEROCOMPTE' then
          Numero:=Copy(RIB,debS, lg) else
        if lequel='CLERIB' then
          cle:=Copy(RIB,debS, lg) ;
        debS := debS + lg ;
      end ;
      Dom:=Copy(RIB,debS,length(RIB));
    end
    else
       // On fait comme avant :
       DecodeRIB_OLD(Etab,Guichet,Numero,Cle,Dom,RIB,ISOPays);
    freeAndNil(TRIB);
END ;

Procedure DecodeRIB_OLD ( Var Etab,Guichet,Numero,Cle,Dom : String ; RIB : String ; ISOPays : String = CodeISOFR ) ; //XMG 14/07/03
BEGIN
   if ISOPays=CodeISOES then
   Begin //RIB au format Espagnol
      Etab:=Copy(RIB,1,4) ; Guichet:=Copy(RIB,6,4) ; Cle:=Copy(RIB,11,2) ;
      Numero:=Copy(RIB,14,10) ;
      Dom:=Copy(RIB,26,24)   ;
   End else
   Begin //La reste (Y compris la France)
      Etab:=Copy(RIB,1,5) ; Guichet:=Copy(RIB,7,5) ; Numero:=Copy(RIB,13,11) ;
      Cle:=Copy(RIB,25,2) ; Dom:=Copy(RIB,28,24)   ;
   End ; // XVI 24/02/2005
END ;

function EncodeRIBIban ( Etab,Guichet,Numero,Cle,Dom,Iban : String ; ISOPays : String = CodeISOFR ) : String ; //XMG 14/07/03
Var St : String ;
begin
  St:=EncodeRIB(Etab,Guichet,Numero,Cle,Dom,ISOPays)+'/'
     +EncodeIban(Iban); // XVI 24/02/2005
  if ((Trim(Etab)='') or (Trim(Guichet)='') or (Trim(Numero)='') or (Trim(Cle)='')) then St:=Format_String('',Length(St)) ;
  Result:=St ;
end;

function  EncodeRIBIban_OLD ( Etab,Guichet,Numero,Cle,Dom,Iban : String ; ISOPays : String = CodeISOFR ) : String ; //XMG 14/07/03
Var St : String ;
begin
  St:=EncodeRIB_OLD(Etab,Guichet,Numero,Cle,Dom,ISOPays)+'/'
     +EncodeIban(Iban); // XVI 24/02/2005
  if ((Trim(Etab)='') or (Trim(Guichet)='') or (Trim(Numero)='') or (Trim(Cle)='')) then St:=Format_String('',Length(St)) ;
  Result:=St ;
end;

procedure DecodeRIBIban ( Var Etab,Guichet,Numero,Cle,Dom,Iban : String ; RIB : String ) ;
var ii, IbanPos : integer ;
    ISOPays     : String ;
begin
  IBANpos:=0 ;
  repeat
    ii:=pos('/',copy(RIB,IBANpos+1,length(RIB))) ;
    if ii>0 then IBANPos:=IBANpos+ii ;
  until ii<=0 ;
  iban:='' ;
  if IBANPos>0 then
  begin
     Iban:=DecodeIban(Copy(RIB,IBANPos+1,length(RIB)));
     ISOPays:=copy(Iban,1,2) ;
     DecodeRib(Etab,Guichet,Numero,Cle,Dom,Rib,ISOPays) ;
  End ; // XVI 24/02/2005
end;

procedure DecodeRIBIban_OLD ( Var Etab,Guichet,Numero,Cle,Dom,Iban : String ; RIB : String ) ;
var ii, IbanPos : integer ;
    ISOPays     : String ;
begin
  IBANpos:=0 ;
  repeat
    ii:=pos('/',copy(RIB,IBANpos+1,length(RIB))) ;
    if ii>0 then IBANPos:=IBANpos+ii ;
  until ii<=0 ;
  iban:='' ;
  if IBANPos>0 then
  begin
     Iban:=DecodeIban(Copy(RIB,IBANPos+1,length(RIB)));
     ISOPays:=copy(Iban,1,2) ;
     DecodeRIB_OLD(Etab,Guichet,Numero,Cle,Dom,Rib,ISOPays) ;
  End ; // XVI 24/02/2005
end;

function EncodeIban(pszIban : String) : String;
var
  St : String ;
  i,L : Integer;
begin
  if not IsNumeric(Copy(pszIban,3,length(pszIban))) then begin Result := pszIban; Exit; End;
  L := length(pszIban);
  St := Copy(pszIban,1,2);
  for i := 3 to L do begin
    // Code chaque paire de chiffre
    if (i mod 2 <> 0) then
      St := St + Chr(StrToInt(Copy(pszIban,i,2))+33);
  end;
  // Si nombre impaire : Récupère le dernier chiffre
  if (L mod 2 <> 0) then
    St := '0' + St + Copy(pszIban,L,1)
  else
    St := '1' + St + Copy(pszIban,L,1);
  Result := St;
end;

function DecodeIban(pszIban : String) : String;
var
  St,St2 : String ;
  i,L : Integer;
  bPair : boolean;
begin
  if (Copy(pszIban,1,1)<>'0') and (Copy(pszIban,1,1)<>'1') then begin Result := pszIban; Exit; End;
  L := length(pszIban);
  St := Copy(pszIban,2,2);
  bPair := (Copy(pszIban,1,1) = '1');
  for i := 4 to L do begin
    if bPair and (i=L) then begin Result := St; exit; end;
    St2 := IntToStr( Ord(pszIban[i])-33);
    If (length(St2) = 1) then St2 := '0'+St2;
    St := St + St2
  end;
  Delete(St,length(St)-3,4);
  St := St + Copy(pszIban,L,1);
  Result := St;
end;


function IsBBAN(BBAN : string) : boolean;
begin    
  Result := not(ErreurDansIban(BBAN));
end;

function IsIBAN(IBAN : string) : boolean;
begin
  Result := ErreurDansIban(IBAN);
end;
Function ErreurDansIban(Rib : String) : boolean;
var
  St2, St4, cleIBAN, cleIBAN2, ret, strInter : String ;
	ii : Byte ;
	cleL, i : integer ;
begin
	Result:=True;
  CleIban2 := Copy(Rib,3,2);
	St2 := Copy(RIB,5,length(Rib)-4) + Copy(RIB,1,2) + '00' ;
	if Length(St2)<10 then exit ;
	St2:=UpperCase(St2) ;
	//Transforme les lettres en chiffres selon le NEE 5.3
	i:=1 ;
	while i<Length(St2) do
	begin
		if St2[i] in ['A'..'Z'] then
		BEGIN
			ii:=Ord(St2[i])-65 ;
			st4:= copy(st2,1,i-1) + inttostr(10+ii) + copy(st2,i+1, length(st2));
			st2:=st4 ;
		END ;
		inc(i);
	end ;

	ret := '' ;
	cleL := 0 ;
	st4:='';
	//On découpe par tranche de 9
	//On calcul la clé via mod 97 puis on fait clé + reste du rib
	for i:=1 to (length(st2) div 9)+1 do
	begin
		st4 := copy(st2,1,9) ;
		delete(st2,1,9);
		strInter := inttostr(cleL)+st4 ;
		cleL := strtoint64(strinter) mod 97 ;
	end ;
	//une fois fini, on calcul 98-clé
	cleIBAN := inttostr(98-(cleL  mod 97));
	if length(cleIBAN)=1 then  cleIBAN := '0' + cleIBAN ;
	Result := Not (CleIBAN = CleIban2);
end;

//XMG 14/07/03 début
{***********A.G.L.***********************************************
Auteur  ...... : XMG
Créé le ...... : 14/07/2003
Modifié le ... :
Description .. : Renvoie le format pout Editmask des zones concernants au
Suite ........ : RIB, par rapport au Pays du RIB....
Mots clefs ... : RIB;ZONES;FORMAT;
*****************************************************************}
Function FormatZonesRIB(szPays,szQueZone : String ) : String ;
//szQueZone= 'BQ'.- Banque
//           'GU'.- Guichet
//           'CP'.- Numéro Compte
//           'DC'.- Dogot de controle
Begin
Result:='' ;
if szPays=CodeISOFR then
   begin
   if szQuezone='BQ' then Result:=stringofchar('9',5)  else
   if szQuezone='GU' then Result:=stringofchar('9',5)  else
   if szQuezone='CP' then Result:=stringofchar('a',11) else
   if szQuezone='DC' then Result:=stringofchar('9',2)  else
      ;
   end else
if szPays=CodeISOES then
   begin
   if szQuezone='BQ' then Result:=stringofchar('9',4)  else
   if szQuezone='GU' then Result:=stringofchar('9',4)  else
   if szQuezone='CP' then Result:=stringofchar('9',10) else
   if szQuezone='DC' then Result:=stringofchar('9',2)  else
      ;
   end ;
if Result<>'' then Result:=Result+';0;_' else
if (V_PGI.VersionDev) then PGIBox('Pays et/ou Zone inconnues (Pays='+szPays+'/Zone='+szQueZone+').','Erreur Developpeur') ;
End ;
{***********A.G.L.***********************************************
Auteur  ...... : X.Maluenda
Créé le ...... : 05/05/2004
Modifié le ... :   /  /
Description .. : Formate la partie RIB d'un IBAN (par rapport au pays)
Mots clefs ... : RIB;IBAN;FORMAT;PAYS
*****************************************************************}
Function RIBIBANtoRIB ( vstRibIban, vStCodePaysISO : String ) : String ;
Begin
 Result:='' ;
 if vstCodePaysISO=CodeISOES then
//Format Espagnol
//    123456789*123456789*123
//    |   |   | |
//    123/123/1/123456789/
    Result:=Copy(vStRibIBAN,1,4)+'/'
           +Copy(vStRibIBAN,5,4)+'/'
           +Copy(vStRibIBAN,9,2)+'/'
           +Copy(vStRibIBAN,11,10)
 else
//    123456789*123456789*123
//    |    |    |          |
//    1234/1234/1234567890/1/
    Result:=Copy(vStRibIBAN,1,5)+'/'
           +Copy(vStRibIBAN,6,5)+'/'
           +Copy(vStRibIBAN,11,11)+'/'
           +Copy(vStRibIBAN,22,2) ;
End ;

{***********A.G.L.***********************************************
Auteur  ...... : X.Maluenda
Créé le ...... : 05/05/2004
Modifié le ... :
Description .. : Depuis un IBAN renvoie:
Suite ........ : - Le RIB (Formaté par rapport au Pays)
Suite ........ : - Le Pays (Code ISO)
Suite ........ : - La clé de l'IBAN
Mots clefs ... : RIB;IBAN;
*****************************************************************}
Function  IBANtoRIB(vstIBAN : String ; var vstRIB, vstPays , vstCleIBAN : String ) : Boolean ;
Begin
  vStRIB:='' ;
  vStPays:='' ;
  vStCleIBAN:='' ;
  if copy(vstIBAN,1,1)='*' then
     Delete(vstIBAN,1,1) ;
  Result:=not ErreurDansIBAN(vStIBAN) ;
  if result then
     begin
     vstPays:=Copy(vStIBAN,1,2) ;
     vStCleIBAN:=Copy(vStIBAN,3,2) ;
     vStRIB:=RibIBANtoRIB(Copy(vStIBAN,5,length(vStIBAN)),vStPays) ;
     End ;
End ;

{***********A.G.L.***********************************************
Auteur  ...... :

Créé le ...... : 17/01/2000

Modifié le ... :   /  /

Description .. : Calcul clé RIB

Mots clefs ... : CLE;RIB

*****************************************************************}
//XMG 14/07/03 début
Function VerifRibFRA(Banque,Guichet,Compte : String) : String ;
Var St2,St3,Rib : String ;
    ii : Byte ;
    i : Integer ;

BEGIN
Result:='' ; St2:=Trim(Banque)+Trim(Guichet)+Trim(Compte)+'00' ; if Length(St2)<10 then exit ;
St2:=UpperCase(St2) ;
For i:=1 to Length(St2) do if St2[i] in ['A'..'Z'] then
    BEGIN
    if St2[i]>='S' then St2[i]:=Succ(St2[i]) ;
    ii:=(((Ord(St2[i])-64)-1) mod 9)+1 ; St2[i]:=Chr(48+ii) ;
    END ;
Repeat
  St3:=Copy(St2,1,2) ;
  if StrToInt(St3)>=97 then BEGIN Delete(St2,1,2) ; St2:=IntToStr(StrToInt(St3) mod 97)+St2 ; END
                       else BEGIN St3:=Copy(St2,1,3) ; Delete(St2,1,3) ; St2:=IntToStr(StrToInt(St3) mod 97)+St2 ; END ;
Until Length(St2)<=2 ;
Rib:=IntToStr(97-StrToInt(St2)) ; if Length(Rib)<2 then Rib:='0'+Rib ;
Result:=Rib ;
END ;
//////////////////////////////////////////////////////////////////////////////////
   Function CalculeDCESP ( szPartie : String) : String ;
   const ValXPos : Array[1..10] of byte = (1,2,4,8,5,10,9,7,3,6) ;
   Var ii,RR : integer ;
   Begin
   Result:='' ;
   RR:=0 ;
   For ii:=Length(szPartie) downto 1 do
      RR:=RR+Valeuri(Copy(szpartie,ii,1))*ValXPos[ii] ;
   RR:=RR mod 11 ;
   RR:=11 - RR ;
   if RR>10 then RR:=0 else
   if RR>9  then RR:=1 ;
   Result:=inttostr(RR) ;
   End ;
   //////////////////////////////////////////////////////////////////////////////////
Function VerifRibESP(Banque,Guichet,Compte : String) : String ;
BEGIN
Result:='' ;
Banque:=stringofchar('0',4)+Banque ;
Guichet:=stringofchar('0',4)+Guichet ;
Result:=CalculeDCESP('00'+Copy(Banque,length(Banque)-3,4)+Copy(Guichet,length(Guichet)-3,4)) ;
Compte:=stringofchar('0',10)+Compte;
Result:=Result+CalculeDCESP(Copy(Compte,length(Compte)-9,10)) ;
END ;
//////////////////////////////////////////////////////////////////////////////////
Function VerifRib(Banque,Guichet,Compte : String ; ISOPays : String = CodeISOFR) : String ;
Begin
if ISOPays=CodeISOFR then Result:=VerifRibFRA(Banque,Guichet,Compte) else
if ISOPays=CodeISOES then Result:=VerifRibESP(Banque,Guichet,Compte) else
   ;
End ;
//XMG 14/07/03 fin


{***********A.G.L.***********************************************
Auteur  ...... :

Créé le ...... : 17/01/2000

Modifié le ... :   /  /

Description .. : Calcul clé RIB

Mots clefs ... : CLE;RIB

*****************************************************************}

{Function VerifRib(Banque,Guichet,Compte : String ; ISOPays : String = CodeISOFR) : String ; //XMG 14/07/03
Var St2,St3,Rib : String ;
    ii : Byte ;
    i : Integer ;

BEGIN
Result:='' ; St2:=Trim(Banque)+Trim(Guichet)+Trim(Compte)+'00' ; if Length(St2)<10 then exit ;
St2:=UpperCase(St2) ;
For i:=1 to Length(St2) do if St2[i] in ['A'..'Z'] then
    BEGIN
    if St2[i]>='S' then St2[i]:=Succ(St2[i]) ;
    ii:=(((Ord(St2[i])-64)-1) mod 9)+1 ; St2[i]:=Chr(48+ii) ;
    END ;
Repeat
  St3:=Copy(St2,1,2) ;
  if StrToInt(St3)>=97 then BEGIN Delete(St2,1,2) ; St2:=IntToStr(StrToInt(St3) mod 97)+St2 ; END
                       else BEGIN St3:=Copy(St2,1,3) ; Delete(St2,1,3) ; St2:=IntToStr(StrToInt(St3) mod 97)+St2 ; END ;
Until Length(St2)<=2 ;
Rib:=IntToStr(97-StrToInt(St2)) ; if Length(Rib)<2 then Rib:='0'+Rib ;
Result:=Rib ;
END ;
}

// Verif code Siret
{Description	 Cette fonction permet de contrôler la validité du numéro
				 SIRET (ou SIREN) selon la table qui suit :

				 N° SIRET (14 chiffres)		N° SIREN (9 chiffres)

				 On multiplie le chiffre du n° SIRET ou SIREN, en commençant
				 par la fin par 1, 2, 1, 2, ...
				 Le total des produits de ces multiplications doit toujours
				 être un nombre de dizaines se terminant par 0.
				 Dans le cas contraire, le N° SIRET ou SIREN est faux.
}


{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 29/11/2000
Modifié le ... :   /  /
Description .. : Verification du numéro de Siet ou Siren
Mots clefs ... : SIREN;SIRET
*****************************************************************}
Function VerifSiret (Siret : string) : boolean ;
var i,x,y,z : integer;
    stY : string ;
begin
z:=0; x:=0;
result:=false;
If Siret='' then exit;
If (length(Siret)<>9) and (length(Siret)<>14) then Exit ;
for i:=1 to length(Siret) do begin if not (Siret[i] in ['0'..'9']) then exit; end;

for i:=length(Siret) downto 1 do
    begin
    if (x<>1) then x:=1 else x:=2;
    y:=strToInt(Siret[I])* x ;
    if y>9 then begin stY:=inttostr(y); y:=strToInt(stY[1])+strToInt(stY[2]); end;
    z:=z+y;
    end;
result:= ((z mod 10)=0);
end;

{***********A.G.L.***********************************************
Auteur  ...... :

Créé le ...... : 17/01/2000

Modifié le ... :   /  /

Description .. : converti un string ACTION=xxxxx en TActionFiche

Mots clefs ... : TACTIONFICHE;CONVERSION

*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure SetJPEGOptions(Image1: TImage; Zoom : integer = 0);
var
  Temp: Boolean;
begin
  Temp := Image1.Picture.Graphic is TJPEGImage;
  if Temp then
    with TJPEGImage(Image1.Picture.Graphic) do
    begin
      PixelFormat := TJPEGPixelFormat(0);
      Scale := TJPEGScale(Zoom);
      Grayscale := False;
      Performance := TJPEGPerformance(0);
      ProgressiveDisplay := false;
    end;

end ;
{$ENDIF}

{
***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis

Créé le ...... : 17/01/2000

Modifié le ... :   /  /

Description .. : Charge une image Bitmap à partir d'un champ DB dans un controle TImage

Mots clefs ... : IMAGE;BITMAP;TIMAGE

*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure LoadBitMapFromChamp(DS : TDataSet; Champ : string; Image : TImage; IsJpeg : boolean = False; Zoom : integer =0) ;
var TJ : TJpegImage ;
{$IFDEF EAGLCLIENT}
    s: TStringStream ;
{$ELSE}
    FField : TField ;
    s      : TmemoryStream ;
{$ENDIF}
begin

if Champ='' then exit;
{$IFDEF EAGLCLIENT}
s:=TStringStream.Create(DS.FindField(Champ).AsString) ;
      try
{$ELSE}
FField:=DS.FindField(Champ) ;
if FField=nil then exit ;
s:=TmemoryStream.Create ;
Image.Picture:=Nil ;
if (FField is TBlobField) And (Not TBlobField(FField).IsNull) then
   begin
      try
      TBlobField(FField).SaveToStream(s) ;
{$ENDIF}
      s.Seek(0,0) ;
      if IsJpeg then
         begin
             TJ:=TJpegImage.create ;
             TJ.LoadFromStream(s) ;
//             Image.Picture.Bitmap.assign(TJ) ;
             Image.Picture.assign(TJ) ;
             TJ.free ;
         end else
            Image.Picture.Bitmap.LoadFromStream(s) ;
      SetJPEGOptions (Image, zoom) ;
      finally end ;
{$IFNDEF EAGLCLIENT}
   end ;
{$ENDIF}
{$IFNDEF EAGLCLIENT}
if (FField is TStringField) and (Trim(FField.AsString)<> '') AND ( FileExists(FField.AsString)) then
   begin
      try
      Image.Picture.LoadFromFile(FField.AsString) ;
      SetJPEGOptions (Image );
      finally end ;
   end;
{$ENDIF}
s.Free ;
end;
{$ENDIF}

{$IFNDEF EAGLSERVER}
procedure BlobToFile (VarName:string;Value:Variant) ;
var ThePanel:TPanel;
	RichDest: TRichEdit;
    NomFichier:string;
begin
      ThePanel := TPanel.Create( nil );
      ThePanel.Visible := False;
      ThePanel.ParentWindow := GetDesktopWindow;

    NomFichier:=VarName+'.RTF';
    SysUtils.DeleteFile(NomFichier) ;
    RichDest := TRichEdit.Create( ThePanel );
    RichDest.Parent:=ThePanel;
    StringTorich(RichDest,Value);
    RichDest.Lines.SaveToFile( NomFichier);
    ThePanel.Free;
end;
{$ENDIF}

Procedure ControleUsers ;
Var Q : TQuery ;
    NbUss : Integer ;
BEGIN
NbUss:=0 ;
Q:=OpenSQL('SELECT COUNT(*) FROM UTILISAT WHERE US_SUPERVISEUR="X"',True) ;
if Not Q.EOF then NbUss:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
if NbUss>0 then Exit ;
Q:=OpenSQL('SELECT * FROM UTILISAT',False) ;
if Not Q.EOF then
   BEGIN
   Q.Edit ;
   Q.FindField('US_SUPERVISEUR').AsString:='X' ;
   Q.Post ;
   END ;
Ferme(Q) ;
END ;

{=========================== Fonctions de conversion ==========================}
FUNCTION PIVOTTOEURO ( X : Double ) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;
XX:=X/V_PGI.TauxEuro ;
if GetParamSoc('SO_TENUEEURO') then Result:=Arrondi(XX,V_PGI.OkDecV) else Result:=Arrondi(XX,V_PGI.OkDecE) ;
end ;

FUNCTION PIVOTTOEURONA ( X : Double ) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;
XX:=X/V_PGI.TauxEuro ; Result:=Arrondi(XX,6) ;
end ;

FUNCTION EUROTOPIVOT ( X : Double ) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;
XX:=X*V_PGI.TauxEuro ;
if GetParamSoc('SO_TENUEEURO') then Result:=Arrondi(XX,V_PGI.OkDecE) else Result:=Arrondi(XX,V_PGI.OkDecV) ;
end ;

FUNCTION EUROTOPIVOTNA ( X : Double ) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;
XX:=X*V_PGI.TauxEuro ; Result:=Arrondi(XX,6) ;
end ;

FUNCTION PIVOTTODEVISE(X,Taux,Quotite : Double ; Decim : byte ) : Double ;
begin
Result:=0 ; if X=0 then Exit ;
If Taux=0 Then Taux:=1 ; if Quotite=0 then Quotite:=1 ;
Result:=Arrondi(X*Quotite/Taux,Decim) ;
end ;

FUNCTION PIVOTTODEVISENA(X,Taux,Quotite : Double) : Double ;
begin
Result:=0 ; if X=0 then Exit ;
If Taux=0 Then Taux:=1 ; if Quotite=0 then Quotite:=1 ;
Result:=Arrondi(X*Quotite/Taux,6) ;
end ;

FUNCTION DEVISETOPIVOT(X,Taux,Quotite : Double) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;  if Taux = 0 then Taux := 1;
if Quotite=0 then Quotite:=1 ; XX:=X*Taux/Quotite ;
if GetParamSoc('SO_TENUEEURO') then Result:=Arrondi(XX,V_PGI.OkDecE) else Result:=Arrondi(XX,V_PGI.OkDecV) ;
end ;

FUNCTION DEVISETOPIVOTEx(X,Taux,Quotite : Double; NbDec : integer) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ; if Taux = 0 then Taux := 1;
if Quotite=0 then Quotite:=1 ; XX:=X*Taux/Quotite ;
Result:=Arrondi(XX,NbDec);
end ;

FUNCTION DEVISETOPIVOTNA(X,Taux,Quotite : Double) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;
if Quotite=0 then Quotite:=1 ; XX:=X*Taux/Quotite ;
if GetParamSoc('SO_TENUEEURO') then Result:=Arrondi(XX,6) else Result:=Arrondi(XX,6) ;
end ;

FUNCTION DEVISETOEURO ( X,Taux,Quotite : Double ) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;
if Quotite=0 then Quotite:=1 ; XX:=X*Taux/Quotite/V_PGI.TauxEuro ;
if GetParamSoc('SO_TENUEEURO') then Result:=Arrondi(XX,V_PGI.OkDecV) else Result:=Arrondi(XX,V_PGI.OkDecE) ;
end ;

FUNCTION DEVISETOEURONA ( X,Taux,Quotite : Double ) : Double ;
Var XX : Double ;
begin
Result:=0 ; if X=0 then Exit ;
if Quotite=0 then Quotite:=1 ; XX:=X*Taux/Quotite/V_PGI.TauxEuro ;
Result:=Arrondi(XX,6) ;
end ;

FUNCTION EUROTODEVISE ( X,Taux,Quotite : Double ; Decim : integer ) : Double ;
begin
Result:=0 ; if X=0 then Exit ;
if Quotite=0 then Quotite:=1 ;
Result:=Arrondi(X*V_PGI.TauxEuro*Quotite/Taux,Decim) ;
end ;

FUNCTION EUROTODEVISENA ( X,Taux,Quotite : Double ) : Double ;
begin
Result:=0 ; if X=0 then Exit ;
if Quotite=0 then Quotite:=1 ;
Result:=Arrondi(X*V_PGI.TauxEuro*Quotite/Taux,6) ;
end ;

FUNCTION DEVISETOFRANC(X,Taux,Quotite : Double) : Double ;
BEGIN Result:=DeviseToPivot(X,Taux,Quotite) ; END ;

FUNCTION DEVISETOFRANCNA(X,Taux,Quotite : Double) : Double ;
BEGIN Result:=DeviseToPivotNA(X,Taux,Quotite) ; END ;

FUNCTION FRANCTOEURO ( X : Double ) : Double ;
BEGIN Result:=PivotToEuro(X) ; END ;

FUNCTION FRANCTOEURONA ( X : Double ) : Double ;
BEGIN Result:=PivotToEuroNA(X) ; END ;

FUNCTION EUROTOFRANC ( X : Double ) : Double ;
BEGIN Result:=EuroToPivot(X) ; END ;

FUNCTION EUROTOFRANCNA ( X : Double ) : Double ;
BEGIN Result:=EuroToPivotNA(X) ; END ;

function GetSousSectionUnique(Ax : string) : string;
var QQ: TQuery;
begin
  Result := '';
  QQ := OpenSQL('SELECT CAS_SOUSPLAN FROM CSOUSPLAN WHERE CAS_AXE="'+Ax+'"',True,1,'',true);
  if not QQ.Eof then
  begin
    Result := QQ.Fields[0].AsString;
  end;
  ferme (QQ);
end;

Function CreerSectionVolee ( CodeSect,Libelle : String ; NumA : integer ) : boolean ; overload;
Var TOBS : TOB ;
    Ax,SousSection : String ;
BEGIN
  Result:=False ;
  if NumA<=0 then Exit ;
  if CodeSect='' then Exit ;
  Ax:='A'+IntToStr(NumA) ;
  if VH^.LiaisonY2ViaShare then
  begin
    if ExisteSQL('SELECT CSP_SECTION FROM CSECTION WHERE CSP_SECTION="'+CodeSect+'" AND CSP_AXE="'+Ax+'"') then Exit ;
    SousSection := GetSousSectionUnique(Ax);
    CodeSect:=BourreLaDonc(CodeSect,TFichierBase(Ord(fbAxe1)+NumA-1)) ;
    TOBS:=TOB.Create('CSECTION',Nil,-1) ;
    TOBS.PutValue('CSP_SECTION',CodeSect) ;
    TOBS.PutValue('CSP_SOUSPLAN',SousSection) ;
    TOBS.PutValue('CSP_LIBELLE',Copy(Libelle,1,35)) ;
    TOBS.PutValue('CSP_ABREGE',Copy(Libelle,1,17)) ;
    TOBS.PutValue('CSP_SENS','M') ;
    TOBS.PutValue('CSP_AXE',Ax) ;
    TOBS.PutValue('CSP_CREERPAR','AFF') ;
    TOBS.PutValue('CSP_SOLDEPROGRESS','X') ;
    TOBS.PutValue('CSP_CODEIMPORT',CodeSect) ;
    // CEGID V9
    TOBS.PutValue('CSP_INVISIBLE','-') ;
    // ---
    Result:=TOBS.InsertDB(Nil) ;
    TOBS.Free ;
  end else
  begin
    if ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="'+CodeSect+'" AND S_AXE="'+Ax+'"') then Exit ;
    CodeSect:=BourreLaDonc(CodeSect,TFichierBase(Ord(fbAxe1)+NumA-1)) ;
    TOBS:=TOB.Create('SECTION',Nil,-1) ;
    TOBS.PutValue('S_SECTION',CodeSect) ;
    TOBS.PutValue('S_LIBELLE',Copy(Libelle,1,35)) ;
    TOBS.PutValue('S_ABREGE',Copy(Libelle,1,17)) ;
    TOBS.PutValue('S_SENS','M') ;
    TOBS.PutValue('S_AXE',Ax) ;
    TOBS.PutValue('S_CREERPAR','AFF') ;
    TOBS.PutValue('S_SOLDEPROGRESSIF','X') ;
    TOBS.PutValue('S_CODEIMPORT',CodeSect) ;
    // CEGID V9
    TOBS.PutValue('S_INVISIBLE','-') ;
    // ---
    Result:=TOBS.InsertDB(Nil) ;
    TOBS.Free ;
  end;
END ;

Function CreerSectionVolee ( CodeSect,Libelle,Axe : String ) : boolean ; overload;
Var TOBS : TOB ;
    SousSection : string;
BEGIN
Result:=False ;
if CodeSect='' then Exit ;
  if VH^.LiaisonY2ViaShare then
  begin
    if ExisteSQL('SELECT CSP_SECTION FROM CSECTION WHERE CSP_SECTION="'+CodeSect+'" AND CSP_AXE="'+Axe+'"') then Exit ;
    SousSection := GetSousSectionUnique(Axe);
    TOBS:=TOB.Create('CSECTION',Nil,-1) ;
    TOBS.PutValue('CSP_SECTION',CodeSect) ;
    TOBS.PutValue('CSP_SOUSPLAN',SousSection) ;
    TOBS.PutValue('CSP_LIBELLE',Copy(Libelle,1,35)) ;
    TOBS.PutValue('CSP_ABREGE',Copy(Libelle,1,17)) ;
    TOBS.PutValue('CSP_SENS','M') ;
    TOBS.PutValue('CSP_AXE',Axe) ;
    TOBS.PutValue('CSP_CREERPAR','AFF') ;
    TOBS.PutValue('CSP_SOLDEPROGRESS','X') ;
    TOBS.PutValue('CSP_CODEIMPORT',CodeSect) ;
    // CEGID V9
    TOBS.PutValue('CSP_INVISIBLE','-') ;
    // ---
    Result:=TOBS.InsertDB(Nil) ;
    TOBS.Free ;
  end else
  begin
    if ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="'+CodeSect+'" AND S_AXE="'+Axe+'"') then Exit ;
    TOBS:=TOB.Create('SECTION',Nil,-1) ;
    TOBS.PutValue('S_SECTION',CodeSect) ;
    TOBS.PutValue('S_LIBELLE',Copy(Libelle,1,35)) ;
    TOBS.PutValue('S_ABREGE',Copy(Libelle,1,17)) ;
    TOBS.PutValue('S_SENS','M') ;
    TOBS.PutValue('S_AXE',Axe) ;
    TOBS.PutValue('S_CREERPAR','AFF') ;
    TOBS.PutValue('S_SOLDEPROGRESSIF','X') ;
    TOBS.PutValue('S_CODEIMPORT',CodeSect) ;
    // CEGID V9
    TOBS.PutValue('S_INVISIBLE','-') ;
    // ---
    Result:=TOBS.InsertDB(Nil) ;
    TOBS.Free ;
  end;
END ;

Procedure AttribCotation ( TOBP : TOB ) ;
Var Cote,Taux : Double ;
    Dev       : String3 ;
    Pref      : String ;
    i,iCote   : integer ;
BEGIN
iCote:=0;
if TOBP.NomTable='PIECE' then Pref:='GP' else
   if TOBP.NomTable='PIEDBASE' then Pref:='GPB' else
      if TOBP.NomTable='PIEDECHE' then Pref:='GPE' else Exit ;
if TOBP.FieldExists(Pref+'_TAUXDEV') then Taux:=TOBP.GetValue(Pref+'_TAUXDEV') else Exit ;
if TOBP.FieldExists(Pref+'_DEVISE') then Dev:=TOBP.GetValue(Pref+'_DEVISE') else Exit ;
if ((Dev=V_PGI.DevisePivot) or (Dev=V_PGI.DeviseFongible)) then Cote:=1.0 else
   if V_PGI.TauxEuro<>0 then Cote:=Taux/V_PGI.TauxEuro else Cote:=1 ;
if TOBP.FieldExists(Pref+'_COTATION') then
   BEGIN
   TOBP.PutValue(Pref+'_COTATION',Cote) ;
   if TOBP.NomTable='PIECE' then
      BEGIN
      for i:=0 to TOBP.Detail.Count-1 do
          BEGIN
          if i=0 then iCote:=TOBP.Detail[i].GetNumChamp('GL_COTATION') ;
          TOBP.Detail[i].PutValeur(iCote,Cote) ;
          END ;
      END ;
   END ;
END ;

Procedure BrancheParamSocAffiche (Var stVirerBranche, stAfficherBranche : string);
Var stAffiche :String;

  function GetBranchesPieces : string;
  begin
    result := ';SCO_PARAMPIECES;SCO_PARAMPIECECOMPL';
  end;

BEGIN
stVirerBranche:='' ; stAfficherBranche:='' ;
// Branches affichées
stAfficherBranche := 'SCO_COORDONNEES;SCO_COMPTABLES'; // Base comptable nécessaire
stAffiche:=StAfficherBranche;   //mcd08/03/02 pour stocker avant GEscom
if ctxGescom in V_PGI.PGIContexte then
   BEGIN
   // Paramétres gestion commerciale
   //stAfficherBranche:=stAfficherBranche+';SCO_DATESDIVERS;SCO_GESTIONCOMMERCIALE' ;
   { mng 13-10-05 pas de stock en BSuite stAfficherBranche:=stAfficherBranche+';SCO_DATESDIVERS;SCO_GCSTOCK' ;}
   stAfficherBranche:=stAfficherBranche+';SCO_DATESDIVERS' ;
   {$IFDEF GCGC}
   if VH_GC.GCAchatStockSeria then
     stAfficherBranche:=stAfficherBranche+';SCO_GCSTOCK;SCO_GCSTOCK1;SCO_GCSTOCK2';
   {$ENDIF GCGC}

   stAfficherBranche:=stAfficherBranche+';SCO_PARAMSDEFAUT;SCO_PARAMSARTICLES' ;
{$IFDEF CRM}
   stAfficherBranche:=stAfficherBranche+';SCO_PREFGC;SCO_CLIENTS' ;
{$ELSE}
   stAfficherBranche:=stAfficherBranche+';SCO_PREFGC;SCO_GCPONTCOMPTABLE;SCO_CLIENTS' ;
{$ENDIF CRM}
   stAfficherBranche:=stAfficherBranche+';SCO_GCCOMPTECOMPTABLE;SCO_GCMODELESEDITION' ;
{$IFDEF CRM}
   stAfficherBranche:=stAfficherBranche+';SCO_FORMZONECLIENT' ;
{$ELSE}
   stAfficherBranche:=stAfficherBranche+';SCO_FORMZONECLIENT;SCO_GCTOX;SCO_PARAMDEPOT_ETAB' ;
{$ENDIF CRM}
{$IFNDEF CRM}
   stAfficherBranche:=stAfficherBranche+';SCO_PREFLOGISTIQUE';
{$ENDIF !CRM}
   {$IFDEF QUALITE}
     if VH_GC.QualiteSeria then
     	stAfficherBranche:=stAfficherBranche+';SCO_QUALITE' ;
   {$ENDIF QUALITE}
   {$IFNDEF GPAO}
    {$IFDEF GCGC}
   if (VH_GC.OASeria) and (GetParamSocSecur('SO_ASSEMBLAGE', false)) then
    {$ENDIF GCGC}
     stAfficherBranche:=stAfficherBranche+';SCO_ASSEMBLAGE';
   {$ENDIF !GPAO}
   {$IFDEF ACCESCBN}
     stAfficherBranche:=stAfficherBranche+';SCO_CBN';
   {$ENDIF ACCESCBN}
   {$IFDEF ACCESSCM}
   if VH_GC.SCMSeria then
   begin
     stAfficherBranche:=stAfficherBranche+';SCO_SYNAPTIQUE';
     { Paramètres Import/Export GPAO }
     if stVirerBranche = '' then
       stVirerBranche := stVirerBranche + 'SCO_QIMPORTEXPORTPRODUFLEX'
     else
       stVirerBranche := stVirerBranche + ';SCO_QIMPORTEXPORTPRODUFLEX';
     { Paramètres Import }
     stVirerBranche := stVirerBranche + ';SCO_QIMPORTEXPORTGENERAL';
     { Paramètres Import/Export Dialogue HOST}
     stVirerBranche := stVirerBranche + ';SCO_QIMPEXPDIALOGUEHOST';
     { Paramètres Info à afficher dans le Plan Général des phases}
     stVirerBranche := stVirerBranche + ';SCO_QINFOPLANGP';
     { Paramètres Profil groupage}
     stVirerBranche := stVirerBranche + ';SCO_QPROFILGROUPAGE';
     { Paramètres Infos à afficher dans objectifs/prévisions}
     stVirerBranche := stVirerBranche + ';SCO_PREVISION';
     { Paramètres Calcul prévision}
     stVirerBranche := stVirerBranche + ';SCO_QBPPREVISION';
   end;
   {$ENDIF ACCESSCM}
   { Objectifs }
   {$IFDEF GESCOM OR PAIEGRH}
   if VH_GC.OBJSeria then
     stAfficherBranche:=stAfficherBranche+';SCO_OBJECTIFS';
   {$ENDIF GESCOM OR PAIEGRH}
   stAfficherBranche:=stAfficherBranche + GetBranchesPieces;
   stAfficherBranche:=stAfficherBranche+';SCO_TARIFS';
   {mng 06/03/07 plus de paramsoc alertes
   stAfficherBranche:=stAfficherBranche+';SCO_GCALERTES';}
   if GetParamSocSecur('SO_GCCPTAIMMODIV', False) then
     stAfficherBranche:=stAfficherBranche+';SCO_LIENSERVANTISSIMO';
   END ;
if ctxMode in V_PGI.PGIContexte then
   BEGIN
   // Paramétres Mode
   stAfficherBranche:=stAfficherBranche+';SCO_DATESDIVERS;SCO_EURO;SCO_DIVERS' ;
   stAfficherBranche:=stAfficherBranche+';SCO_PARAMSDEFAUT;SCO_PARAMSARTICLES;SCO_GCPRESTATION' ;
   stAfficherBranche:=stAfficherBranche+';SCO_GCDIMENSION' ;
   stAfficherBranche:=stAfficherBranche+';SCO_PREFGC;SCO_GCPONTCOMPTABLE;SCO_CLIENTS' ;
   stAfficherBranche:=stAfficherBranche+';SCO_FORMZONECLIENT;SCO_ECOMMERCE;SCO_GCTOX' ;
   stAfficherBranche:=stAfficherBranche+';SCO_GCFO;SCO_GCCONNEXIONS;SCO_GCORLI' ;
   stAfficherBranche:=stAfficherBranche+';SCO_PARAMDEPOT_ETAB;SCO_GCMODELESEDITION' ;
   // pas le cti stAfficherBranche:=stAfficherBranche+';SCO_PROSPECTS' ;
   stAfficherBranche:=stAfficherBranche+';SCO_GRCPREF;SCO_RTDOCUMENTS;SCO_RTPREFERENCES;SCO_RTPROJETS' ;
   stAfficherBranche:=stAfficherBranche+';SCO_RTPROCHQMPSLIBRES;SCO_RTACTIONS;SCO_RTOPERATIONS;SCO_SUSPECT;SCO_CHAINAGE' ;
   stAfficherBranche:=stAfficherBranche+';SCO_CONTACTS';
//   stAfficherBranche:=stAfficherBranche+';SCO_RTDOCUMENTS' ;           A conditionner suivant si GRC O/N
   exit;
   END;
{$IFDEF CRM}
//  BBI plus de branche Ecommerce dans les paramsoc
//   stAfficherBranche:=stAfficherBranche+';SCO_ECOMMERCE;' ;
   stAfficherBranche:=stAfficherBranche ;
{$ELSE}
//  BBI plus de branche Ecommerce dans les paramsoc
//   stAfficherBranche:=stAfficherBranche+';SCO_ECOMMERCE;SCO_GCPRESTATION;' ;
   stAfficherBranche:=stAfficherBranche+';SCO_GCPRESTATION;' ;
{$ENDIF CRM}
    // pour GA dans GC
{$IFDEF CRM}
   stAfficherBranche:=stAfficherBranche+';SCO_RESSOURCE';
   stAfficherBranche:=stAfficherBranche+';SCO_PREVISION';
{$ELSE}
   stAfficherBranche:=stAfficherBranche+';SCO_AFFPREFERENCES;SCO_AFFACTURE;SCO_RESSOURCE;SCO_AFFACTIVITE;SCO_COMPORTACTIVITE';
   stAfficherBranche:=stAfficherBranche+';SCO_AFEDITION;SCO_AFPREFERENCE;SCO_AFDATES';
   // BDU - 02/01/08 - FQ : 14805
   // stAfficherBranche:=stAfficherBranche+';SCO_AFEACTIVITE;SCO_AFBUDGET;SCO_AFCUTOFF';
   stAfficherBranche:=stAfficherBranche+';SCO_AFEACTIVITE;SCO_AFCUTOFF';
   //C.B 09/07/2007
   //supression lien SCO_LIENPLANACT
   stAfficherBranche:=stAfficherBranche+';SCO_AFFAIREREGROUPE;SCO_FRAISCOMPTA';
   //stAfficherBranche:=stAfficherBranche+';SCO_AFFAIREREGROUPE;SCO_FRAISCOMPTA;SCO_LIENPLANACT';
{$ENDIF CRM}
 
{$IFDEF GCGC}
   if VH_GC.GAPlanningSeria or VH_GC.GAPlanChargeSeria then
     stAfficherBranche:=stAfficherBranche+';SCO_AFPLANNING;SCO_LIENPLANACT;SCO_PLANNINGGENE;SCO_AFCOMMUNPLAN;SCO_AFPLANCHARGE'; //mcd 07/03/2005 ajout commun +gene

   //C.B 15/05/2007
   if VH_GC.GAPlanningSeria then
     stAfficherBranche := stAfficherBranche+';SCO_YPLANNING';
     
  // C.B 24/03/2006
(*
 {$IFDEF BUSINESSPLACE}
 // pas de révision en SUITE
  if ctxGCAff in V_PGI.PGIContexte then
     stAfficherBranche:=stAfficherBranche+';SCO_REVISIONPRIX';
 {$ENDIF}
*)
if (ctxAffaire in V_PGI.PGIContexte) {or (VH_GC.GASeria)} then       // PCS pour GA + GC
   BEGIN
        // mcd 08/03/02 tout revu pour passer en mode autorisation seulement
   stAfficherBranche:=stAffiche+';SCO_DATESDIVERS'; // fin paramètre compta
{$ifndef OGC}
   {mng 06/03/07 plus de paramsoc alertes
   stAfficherBranche:=stAffiche+';SCO_GCALERTES'; }//Alerte de la GC  mcd 30/03/2006
{$endif}
        // paramètres gescom ATTENTION, si GC + GA, voir si il n'y pas des
        // des branches à remettre ???? (dans le cas GA, on efface toute la branche
        // gescom pour la refaire avec les seules options voulues
   stAfficherBranche:=stAfficherBranche+';SCO_PARAMSDEFAUT;SCO_PARAMSARTICLES;';
   stAfficherBranche:=stAfficherBranche+';SCO_GCPRESTATION;SCO_PREFGC;SCO_GCCOMPTECOMPTABLE';
   stAfficherBranche:=stAfficherBranche+';SCO_GCPONTCOMPTABLE;SCO_CLIENTS;SCO_FORMZONECLIENT';
{$ifndef OGC}
   stAfficherBranche:=stAfficherBranche+';SCO_TARIFS;SCO_GCMODELESEDITION';
   stAfficherBranche:=stAfficherBranche+';SCO_FRAISCOMPTA'; // PL le 18/02/04 : saisie des frais en compta
{$else}
   stAfficherBranche:=stAfficherBranche+';SCO_GCMODELESEDITION';
{$endif}
   stAfficherBranche:=stAfficherBranche+';SCO_AFFACTURE'; // mcd 06/03/2006
        // Paramétres affaires
   stAfficherBranche:=stAfficherBranche+';SCO_AFFPREFERENCES;SCO_RESSOURCE;SCO_AFFACTIVITE;SCO_COMPORTACTIVITE';
   //C.B 09/07/2007
   //supression lien SCO_LIENPLANACT 
   stAfficherBranche:=stAfficherBranche+';SCO_AFEDITION;SCO_AFPREFERENCE;SCO_AFDATES';
   // stAfficherBranche:=stAfficherBranche+';SCO_AFEDITION;SCO_AFPREFERENCE;SCO_AFDATES;SCO_LIENPLANACT';
   // mcd 29/01/03 pas de budget en GA stAfficherBranche:=stAfficherBranche+';SCO_AFEACTIVITE;SCO_AFBUDGET;SCO_AFCUTOFF';
{$ifndef OGC}
   stAfficherBranche:=stAfficherBranche+';SCO_AFEACTIVITE;SCO_AFCUTOFF';
{$endif}
(*
   if not(CtxScot in V_PGI.PGICOntexte) then
   {$IFDEF BUSINESSPLACE}
        // BDU - 02/01/08 - FQ : 14805
        // stAfficherBranche:=stAfficherBranche+';SCO_AFBUDGET;SCO_REVISIONPRIX';
        stAfficherBranche:=stAfficherBranche+';SCO_REVISIONPRIX';
   {$ELSE}  // pas de révision en SUITE
        // BDU - 02/01/08 - FQ : 14805
        // stAfficherBranche:=stAfficherBranche+';SCO_AFBUDGET';
   {$ENDIF}
*)
   if CtxScot in V_PGI.PGICOntexte then
{$ifndef OGC}   //mcd 13/07/07 14360 suppression de parampiececompl
    stAfficherBranche:=stAfficherBranche+';SCO_APPREC;SCO_FACTUREECLATEE;SCO_AFFAIREREGROUPE;SCO_PARAMPIECES'
{$else}
    stAfficherBranche:=stAfficherBranche+';SCO_AFFAIREREGROUPE;SCO_PARAMPIECES'  //mcd 16/04/07 ajout parampiece
{$endif}
   else
   begin
    stAfficherBranche:=stAfficherBranche+';SCO_AFFAIREREGROUPE' + GetBranchesPieces; //AB-200510-Param Pièces
   end;
      //mcd 07/03/2005 nouvelle brnaches.. tout revu
   if VH_GC.GAPlanningSeria or VH_GC.GAPlanChargeSeria then stAfficherBranche:=stAfficherBranche+';SCO_AFCOMMUNPLAN';
   if VH_GC.GAPlanChargeSeria then  stAfficherBranche:=stAfficherBranche+';SCO_AFPLANCHARGE';
   if VH_GC.GAPlanningSeria   then  stAfficherBranche:=stAfficherBranche+';SCO_PLANNINGGENE;SCO_AFPLANNING;SCO_LIENPLANACT;SCO_YPLANNING';
   // 3OB 26/09/2007 : Nouvelle branche de segmentation
   if VH_GC.GASegment   then  stAfficherBranche:=stAfficherBranche+';SCO_OGC';
   // STR 14/01/2008 : nouvelle branche pour Facture Info
   if VH_GC.GAFactInfo   then  stAfficherBranche:=stAfficherBranche+';SCO_FACTUREINFO';
   { Objectifs }
   {$IFDEF GESCOM OR PAIEGRH}
   if VH_GC.OBJSeria          then  stAfficherBranche:=stAfficherBranche+';SCO_OBJECTIFS';
   {$ENDIF GESCOM OR PAIEGRH}
   END;
// MODIF BTP
if CtxBtp in V_PGI.PGIContexte then
   BEGIN
   stAfficherBranche:=stAfficherBranche+';SCO_BTBTP;';
   stVirerBranche := stVirerBranche + ';SCO_APPEL;SCO_AFEACTIVITE;';
{$IFDEF NOMADE}
   stAfficherBranche:=stAfficherBranche+';SCO_GCTOX;';
{$ENDIF}
{$IFDEF NOMADESERVER}
   stAfficherBranche:=stAfficherBranche+';SCO_GCTOX;';
{$ENDIF}
   stAfficherBranche:=stAfficherBranche+';SCO_PARAMDEPOT_ETAB' ;
   stAfficherBranche:=stAfficherBranche+';SCO_GCSTOCK' ;
   END;

if ctxGRC in V_PGI.PGIContexte then
   BEGIN
   // Paramétres gestion de la relation clients
   //stAfficherBranche:=stAfficherBranche+';SCO_PROSPECTS' ;
   stAfficherBranche:=stAfficherBranche+';SCO_GRCPREF;SCO_RTDOCUMENTS;SCO_RTPREFERENCES' ;
{$IFNDEF GRCLIGHT}
   stAfficherBranche:=stAfficherBranche+';SCO_RTPROJETS;SCO_RTOPERATIONS;SCO_SUSPECT' ;
   stAfficherBranche:=stAfficherBranche+';SCO_CIBLAGE';
{$ENDIF GRCLIGHT}
   stAfficherBranche:=stAfficherBranche+';SCO_RTPROCHQMPSLIBRES;SCO_RTACTIONS;SCO_CHAINAGE' ;
 {$ifdef GIGI}
    stAfficherBranche:=stAfficherBranche+';SCO_CONTACTS;SCO_RTMODELEFICHEACTION';   //mcd 11/07/05 pas de ged en GI
 {$ELSE}
   stAfficherBranche:=stAfficherBranche+';SCO_CONTACTS;SCO_RTMODELEFICHEACTION;SCO_GEDGRC';
 {$ENDIF GIGI}
{$IFDEF CTI}
   if VH_RT.CTISeria = true then stAfficherBranche:=stAfficherBranche+';SCO_CTI' ;
{$ENDIF}
{$IF Defined(CRM) or Defined(GRCLIGHT)}
   if VH_RT.FlyDocSeria = true then stAfficherBranche:=stAfficherBranche+';SCO_FLYDOC' ;
{$IFEND}
   END ;
if VH_GC.GRFSeria = true then stAfficherBranche:=stAfficherBranche+';SCO_GRF' ;

if EstbasemultiSoc then stAfficherBranche := stAfficherBranche + ';SCO_MULTISOC';
//GP_20071206_TP_GC15634 >>>

{ Planning de livraison }
stAfficherBranche:=stAfficherBranche+';SCO_WPLANLIVR';
{$ENDIF GCGC}

{$IFDEF GPAO}
if ctxGPAO in V_PGI.PGIContexte then
begin
   // Paramétres GP
   stAfficherBranche:=stAfficherBranche+';SCO_GPAO';
  {$IFDEF AFFAIRE}
    stAfficherBranche:=stAfficherBranche+';SCO_AFFAIRESINDUS' ;
  {$ENDIF AFFAIRE}
  { Branche MES }
  stAfficherBranche:=stAfficherBranche+';SCO_MES' ;
end;
{$ENDIF GPAO}
{$IFDEF SAV}
   {$IFDEF BUSINESSPLACE}
   if VH_GC.SAVSeria then
     stAfficherBranche:=stAfficherBranche+';SCO_SAV';
   {$ENDIF}
{$ENDIF}
{$ifdef GIGI}
{$ifndef EAGLSERVER}
    // pour la GI, si que GRC, on ne voit pas tout dans les paramsoc
if  (V_PGI.PassWord <> CryptageSt(DayPass(Date)))
    and not (VH_GC.GASeria)
    and (VH_GC.GRCSeria) then
     begin
     stAfficherBranche:='SCO_COORDONNEES;SCO_CLIENTS;SCO_FORMZONECLIENT;SCO_GCMODELESEDITION'; // GC
      //+option GRC
     stAfficherBranche:=stAfficherBranche+';SCO_GRCPREF;SCO_RTDOCUMENTS;SCO_RTPREFERENCES;SCO_RTPROJETS' ;
     stAfficherBranche:=stAfficherBranche+';SCO_RTPROCHQMPSLIBRES;SCO_RTACTIONS;SCO_RTOPERATIONS;SCO_SUSPECT;SCO_CHAINAGE' ;
     stAfficherBranche:=stAfficherBranche+';SCO_CONTACTS;SCO_RTMODELEFICHEACTION';
     stAfficherBranche:=stAfficherBranche+';SCO_CIBLAGE';
     end;
{$endif EAGLSERVER}
{$endif GIGI}

{$IFDEF GPAOLIGHT}
  stAfficherBranche:=stAfficherBranche+';SCO_PDR';
{$ENDIF GPAOLIGHT}
END;


function VarToDouble (V : Variant) : double ;
Var St : String ;
BEGIN
Case VarType(v) of
  varSmallint,varInteger,varSingle,varDouble,varCurrency : result:=VarAsType(V,VarDouble) ;
  varEmpty,VarNull : Result:=0 ;
  varString : BEGIN St:=VarAstype(v,VarString) ; if IsNumeric(St) then result:=Valeur(St) else result:=0 ; END ;
  else Result:=0 ;
  END ;
END ;

Function FabricWhereToken ( sToken,Champ : String ) : String ;
Var sWhere,sSub : String ;
BEGIN
Result:='' ; sWhere:='' ;
if ((sToken='') or (Champ='')) then Exit ;
Repeat
 sSub:=ReadTokenSt(sToken) ;
 if sSub<>'' then sWhere:=sWhere+Champ+'="'+sSub+'" OR ' ;
Until (sToken='');
if sWhere<>'' then BEGIN Delete(sWhere,Length(sWhere)-3,4) ; sWhere:='('+sWhere+')' ; END ;
Result:=sWhere ;
END ;


FUNCTION DEVISETODEVISE(X,Taux,Quotite,NewTaux,NewQuotite : Double ; Decim : byte ) : Double ;
var XL : Double;
begin
Result:=0 ; if X=0 then Exit ;
XL := DEVISETOEURONA (X,Taux,Quotite);
Result := EUROTODEVISE (XL,NewTaux,NewQuotite,Decim);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Joël TRIFILIEFF
Créé le ...... : 19/03/2003
Modifié le ... :
Description .. : Test autorisation de modif d'une pièce lors d'un accès
Suite ........ : autre que les menus
Mots clefs ... : CONFIDENTIALITE;MODIFICATION
*****************************************************************}
function JaiLeDroitNatureGCModif(Nature : string) : boolean;

  function AffecteTag(Devis, Prof, Cde, CdeEch, PrepLiv, Liv, LivEch, FactProv, Ticket, AchCde, AchLiv : integer) : boolean;
  begin
    //VENTE
    if Nature = 'DE' then Result := JaiLeDroitTag(Devis) //Devis
    else if Nature = 'PRO' then Result := JaiLeDroitTag(Prof) //Facture proforma
{$IFDEF BTP}
    else if Nature = 'CBT' then Result := JaiLeDroitTag(Cde) //Commande
    else if Nature = 'LBT' then Result := JaiLeDroitTag(Liv) //Commande
{$ENDIF}
    else if Nature = 'CC' then Result := JaiLeDroitTag(Cde) //Commande
    else if Nature = 'CCE' then Result := JaiLeDroitTag(CdeEch) //Commande échantillon
    else if Nature = 'PRE' then Result := JaiLeDroitTag(PrepLiv) //Préparation livraison
    else if Nature = 'BLC' then Result := JaiLeDroitTag(Liv) //Livraison
    else if Nature = 'LCE' then Result := JaiLeDroitTag(LivEch) //Livraison échantillon
    else if Nature = 'FPR' then Result := JaiLeDroitTag(FactProv) //Facture provisoire
    else if Nature = 'FFO' then Result := JaiLeDroitTag(Ticket) //Ticket
    //ACHAT
    else if Nature = 'CF' then Result := JaiLeDroitTag(AchCde) //Commande
    else if Nature = 'BLF' then Result := JaiLeDroitTag(AchLiv) //Livraison
    //TRANSFERTS DE STOCK
    {$IFNDEF MODE}
      else if (Nature = 'TEM') and (GetParamSoc('SO_GCTRV')) then Result := False //Transfert Emis
      else if (Nature = 'TRV') and (GetParamSoc('SO_GCTRV')) then Result := False //Transfert à valider
      else if (Nature = 'TRE') and (GetParamSoc('SO_GCTRV')) then Result := False //Transfert Reçu
    {$ENDIF MODE}
    else Result := true;
  end;

begin
    Result := True;
{$IFDEF GESCOM}
    //SBU - FQ 10175 - DEBUT - en CHR le menu ACHAT est le 163
    {$IFDEF CHR}
    if ctxGescom in V_PGI.PGIContexte then Result := AffecteTag(30221,30222,30223,30226,30229,30224,30227,30225,36101,163172,163173);
    {$ELSE CHR}
    if ctxGescom in V_PGI.PGIContexte then Result := AffecteTag(30221,30222,30223,30226,30229,30224,30227,30225,36101,31211,31212);
    {$ENDIF}
    //SBU - FQ 10175 - FIN
{$ENDIF}
    if (ctxMode in V_PGI.PGIContexte) or (ctxAffaire in V_PGI.PGIContexte) then Result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Joël TRIFILIEFF
Créé le ...... : 19/03/2003
Modifié le ... :
Description .. : Test autorisation de création d'une pièce lors d'un accès
Suite ........ : autre que les menus
Auteur  ...... : CCMX-CEGID - Dominique MARIS - FQ 13344
Modifié le ... : 09/02/07
Description .. : Ajout d'un argument pour savoir si on étudie les droits d'accès
Suite ........ : sur la pièce d'origine (droit uniquement en création) ou sur les
Suite ........ : sur les pièces destinations (droit en création ET en génération)
Auteur  ...... : CCMX-CEGID - Dominique MARIS - FQ 13344
Modifié le ... : 16/02/07
Description .. : Ajout d'un autre argument pour savoir quelle est la pièce d'origine
Suite ........ : de façon à différencier les droits en génération lorsque la pièce
Suite ........ : d'origine peut être différente alors que génération dans même nature
Suite ........ : de pièce. Ex génération de facture ou de retour en avoir sur stock
Mots clefs ... : CONFIDENTIALITE;CREATION
*****************************************************************}
//function JaiLeDroitNatureGCCreat(Nature : string) : boolean;
function JaiLeDroitNatureGCCreat(Nature : string; Typepiece :string = 'Origine'; NatureOrig : string = '') : boolean;

  function AffecteTag(Devis, Prof, Cde, CdeEch, PrepLiv, Liv, LivEch, FactProv, Facture, AvCli, AvStk, AvProv, Ticket, AchCde, AchLiv, AchFact,
                      AvAch, AvAchStock, AvAchStockPrix,RetourFournisseur
    //CCMX-CEGID - FQ 13344 - DEBUT
                      // : integer) : boolean;
    //CCMX-CEGID - FQ 13344 - fin
                      , PrepAchLiv, CdeSSTR
                      : integer) : boolean;
  //SBU - FQ 10175 - DEBUT - tag des menus sous forme de constantes (fait uniquement pour les tag commun a la Gescom et au CHR
  Const
    {$IFDEF CHR}
    AchGeneCde = 16331;                     // Générer cde
    AchGeneLiv = 16332;                     // Générer réception
    AchGeneManuLivCde = 163352;             // Générer manuellement une réception à partir d'une commande
    AchGeneAutoLivCde = 163362;             // Générer automatiquement une réception à partir d'une commande
    AchGeneAvFac = 163343;                  // Générer Facture frn en avoir financier
    AchGeneAvStockRetour = 163342;          // Générer Retour en avoir fournisseur valorisé
    AchGeneManuAvStockRetour = 163356;      // Générer manuellement un avoir fournisseur valorisé à partir d'un retour
    AchGeneAutoAvStockRetour = 163366;      // Générer automatiquement un avoir fournisseur valorisé à partir d'un retour
    AchGeneAvStockPrixFact = 163344;        // Générer Facture frn en avoir fournisseur sur stock
    AchGeneAvStockPrixRetour = 163341;      // Générer Retour en avoir fournisseur sur stock
    AchGeneManuAvStockPrixRetour = 163355;  // Générer manuellement un avoir sur stock à partir d'un retour
    AchGeneAutoAvStockPrixRetour = 163365;  // Générer automatiquement un avoir sur stock à partir d'un retour
    AchGeneFact = 16333;                    // Générer Facture
    AchGeneManuFactCde = 163353;            // Générer manuellement une facture à partir d'une cde
    AchGeneAutoFactCde = 163363;            // Générer automatiquement une facture à partir d'une cde
    AchGeneManuFactLiv = 163354;            // Générer manuellement une facture à partir d'une Livraison
    AchGeneAutoFacLiv = 163364;             // Générer automatiquement une facture à partir d'une Livraison
    {$ELSE}
    AchGeneCde = 31514;                     // Générer cde
    AchGeneLiv = 31511;                     // Générer réception
    AchGeneManuLivCde = 31453;              // Générer manuellement une réception à partir d'une commande
    AchGeneAutoLivCde = 31521;              // Générer automatiquement une réception à partir d'une commande
    AchGeneAvFac = 31543;                   // Générer Facture frn en avoir financier
    AchGeneAvStockRetour = 31542;           // Générer Retour en avoir fournisseur valorisé
    AchGeneManuAvStockRetour = 31459;       // Générer manuellement un avoir fournisseur valorisé à partir d'un retour
    AchGeneAutoAvStockRetour = 31528;       // Générer automatiquement un avoir fournisseur valorisé à partir d'un retour
    AchGeneAvStockPrixFact = 31544;         // Générer Facture frn en avoir fournisseur sur stock
    AchGeneAvStockPrixRetour = 31541;       // Générer Retour en avoir fournisseur sur stock
    AchGeneManuAvStockPrixRetour = 31458;   // Générer manuellement un avoir sur stock à partir d'un retour
    AchGeneAutoAvStockPrixRetour = 31527;   // Générer automatiquement un avoir sur stock à partir d'un retour
    AchGeneFact = 31512;                    // Générer Facture
    AchGeneManuFactCde = 31454;             // Générer manuellement une facture à partir d'une cde
    AchGeneAutoFactCde = 31525;              // Générer automatiquement une facture à partir d'une cde
    AchGeneManuFactLiv = 31457;             // Générer manuellement une facture à partir d'une Livraison
    AchGeneAutoFacLiv = 31522;              // Générer automatiquement une facture à partir d'une Livraison
    {$ENDIF}
  //SBU - FG 10175 - FIN
  begin
    //VENTE
    //CCMX-CEGID - FQ 13344 - DEBUT
    if Typepiece <> 'Destination' then
    begin
    //CCMX-CEGID - FQ 13344 - FIN
      if Nature = 'DE'  then Result := JaiLeDroitTag(Devis) //Devis
      else if Nature = 'PRO' then Result := JaiLeDroitTag(Prof) //Facture proforma
      else if Nature = 'CC' then Result := JaiLeDroitTag(Cde) //Commande
      else if Nature = 'CCE' then Result := JaiLeDroitTag(CdeEch) //Commande échantillon
      else if Nature = 'PRE' then Result := JaiLeDroitTag(PrepLiv) //Préparation livraison
      else if Nature = 'BLC' then Result := JaiLeDroitTag(Liv) //Livraison
      else if Nature = 'LCE' then Result := JaiLeDroitTag(LivEch) //Livraison échantillon
      else if Nature = 'FPR' then Result := JaiLeDroitTag(FactProv) //Facture provisoire
      else if Nature = 'FAC' then Result := JaiLeDroitTag(Facture) //Facture
      else if Nature = 'AVC' then Result := JaiLeDroitTag(AvCli) //Avoir client
      else if Nature = 'AVS' then Result := JaiLeDroitTag(AvStk) //Avoir sur stock
      else if Nature = 'APR' then Result := JaiLeDroitTag(AvProv) //Avoir provisoire
      else if Nature = 'FFO' then Result := JaiLeDroitTag(Ticket) //Ticket
    //ACHAT
      else if Nature = 'CF' then Result := JaiLeDroitTag(AchCde) //Commande
      else if Nature = 'BLF' then Result := JaiLeDroitTag(AchLiv) //Livraison
      else if Nature = 'AF' then Result := JaiLeDroitTag(AvAch) //avoir fournisseur financier
      else if Nature = 'AFP' then Result := JaiLeDroitTag(AvAchStock) //avoir fournisseur valorisé
      else if Nature = 'AFS' then Result := JaiLeDroitTag(AvAchStockPrix) //avoir fournisseur sur stock
      else if Nature = 'BFA' then Result := JaiLeDroitTag(RetourFournisseur) // Retour fournisseur
      else if Nature = 'FF' then Result := JaiLeDroitTag(AchFact) //Facture
      else Result := true;
    //CCMX-CEGID - FQ 13344 - DEBUT
    end
    else
    begin
      if Nature = 'DE'  then Result := JaiLeDroitTag(Devis) //Devis
      else if Nature = 'PRO' then Result := JaiLeDroitTag(Prof) //Facture proforma
      else if Nature = 'CC' then
        begin
          Result := JaiLeDroitTag(Cde); //Commande
          if Result then Result := JaiLeDroitTag(30403); // Générer cde
          if Result then
            if NatureOrig = 'DE' then
            Begin
              Result := JaiLeDroitTag(30451); // Générer manuellement une cde à partir d'un devis
              if result then Result := JaiLeDroitTag(30418); // Générer automatiquement une cde à partir d'un devis
            end
        end
      else if Nature = 'CCE' then Result := JaiLeDroitTag(CdeEch) //Commande échantillon
      else if Nature = 'PRE' then
        begin
          Result := JaiLeDroitTag(PrepLiv); //Préparation livraison
          if Result then Result := JaiLeDroitTag(30409); // Générer Préparation livraison
          if Result then
            if NatureOrig = 'CC' then
            Begin
              Result := JaiLeDroitTag(30452); // Générer manuellement une Prépa livr à partir d'une cde
              if result then Result := JaiLeDroitTag(30413); // Générer automatiquement une Prépa livr à partir d'une cde
            end
        end
      else if Nature = 'BLC' then
        begin
          Result := JaiLeDroitTag(Liv); //Livraison
          if Result then Result := JaiLeDroitTag(30421); // Générer livraison
          if Result then
            if NatureOrig = 'CC' then
            Begin
              Result := JaiLeDroitTag(30453); // Générer manuellement une livraison à partir d'une cde
              if result then Result := JaiLeDroitTag(30415); // Générer automatiquement une livraison à partir d'une cde
            end
            else if NatureOrig = 'PRE' then
            Begin
              Result := JaiLeDroitTag(30455); // Générer manuellement une livraison à partir d'une Prépa Livr
              if result then Result := JaiLeDroitTag(30411); // Générer automatiquement une livraison à partir d'une Prépa Livr
            end
        end
      else if Nature = 'LCE' then
        begin
          Result := JaiLeDroitTag(LivEch); //Livraison échantillon
          if Result then Result := JaiLeDroitTag(30422); // Générer livraison échantillon
          if Result then
            if NatureOrig = 'CCE' then
            Begin
              Result := JaiLeDroitTag(30459); // Générer manuellement une livr échantillon à partir d'une cde d'échantillon
              if result then Result := JaiLeDroitTag(30425); // Générer automatiquement une livr échantillon à partir d'une cde d'échantillon
            end
        end
      else if Nature = 'FPR' then
        begin
          Result := JaiLeDroitTag(FactProv); //Facture provisoire
          if Result then Result := JaiLeDroitTag(30414); // Validation des documents provisoires
        end
      else if Nature = 'FAC' then
        begin
          Result := JaiLeDroitTag(Facture); //Facture
          if Result then Result := JaiLeDroitTag(30405); // Générer Facture
          if Result then
            if NatureOrig = 'CC' then
            Begin
              Result := JaiLeDroitTag(30454); // Générer manuellement une facture à partir d'une cde
              if result then Result := JaiLeDroitTag(30416); // Générer automatiquement une facture à partir d'une cde
            end
            else if NatureOrig = 'PRE' then
            Begin
              Result := JaiLeDroitTag(30456); // Générer manuellement une facture à partir d'une Prépa Livr
              if result then Result := JaiLeDroitTag(30417); // Générer automatiquement une facture à partir d'une Prépa Livr
            end
            else if NatureOrig = 'BLC' then
            Begin
              Result := JaiLeDroitTag(30457); // Générer manuellement une facture à partir d'une Livraison
              if result then Result := JaiLeDroitTag(30412); // Générer automatiquement une facture à partir d'une Livraison
            end
        end
      else if Nature = 'AVC' then
        begin
          Result := JaiLeDroitTag(AvCli); //Avoir client
          if Result then
            if NatureOrig = 'FAC' then
              Result := JaiLeDroitTag(30419); // Générer Facture en avoir client (avoir financier)
        end
      else if Nature = 'AVS' then
        begin
          Result := JaiLeDroitTag(AvStk); //Avoir sur stock
          if Result then
            if NatureOrig = 'FAC' then
              Result := JaiLeDroitTag(30423)    // Générer Facture en avoir sur stock
            else if NatureOrig = 'BRC' then
            begin
              Result := JaiLeDroitTag(30461);   // Générer Retour en avoir sur stock
              if Result then
              begin
                Result := JaiLeDroitTag(30458); // Générer manuellement un avoir sur stock à partir d'un retour
                if result then Result := JaiLeDroitTag(30424); // Générer automatiquement un avoir sur stock à partir d'un retour
              end
            end
        end
      else if Nature = 'APR' then Result := JaiLeDroitTag(AvProv) //Avoir provisoire
      else if Nature = 'FFO' then Result := JaiLeDroitTag(Ticket) //Ticket
      //ACHAT
      else if Nature = 'CF' then
        begin
          Result := JaiLeDroitTag(AchCde); //Commande
          //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
          //if Result then Result := JaiLeDroitTag(31514); // Générer cde
          if Result then Result := JaiLeDroitTag(AchGeneCde); // Générer cde
          if Result then
            if NatureOrig = 'DEF' then
            Begin
              Result := JaiLeDroitTag(31451); // Générer manuellement une cde à partir d'une proposition
              if result then Result := JaiLeDroitTag(31524); // Générer automatiquement une cde à partir d'une proposition
            end
        end
        else if Nature = 'BSA' then
        begin
          Result := JaiLeDroitTag(211623); // Générer reception sous traitance
          if Result then
            if NatureOrig = 'CSA' then
            Begin
              Result := JaiLeDroitTag(211621); // Générer manuellement une Rec SST à partir d'une cde SST
              if result then Result := JaiLeDroitTag(211622); // Générer automatiquement une rec SST à partir d'une Cde SST
            end
        end
        else if Nature = 'PRF' then
        begin
          Result := JaiLeDroitTag(PrepAchLiv); //Préparation réception
          if Result then Result := JaiLeDroitTag(31513); // Générer Préparation reception
          if Result then
            if NatureOrig = 'CF' then
            Begin
              Result := JaiLeDroitTag(31452); // Générer manuellement une Prépa reception à partir d'une cde
              if result then Result := JaiLeDroitTag(31520); // Générer automatiquement une Prépa réception à partir d'une cde
            end
        end
        else if Nature = 'BLF' then
        begin
          Result := JaiLeDroitTag(AchLiv); // Réception fournisseur
          //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
          //if Result then Result := JaiLeDroitTag(31511); // Générer réception
          if Result then Result := JaiLeDroitTag(AchGeneLiv); // Générer réception
          if Result then
            if NatureOrig = 'CF' then
            Begin
              //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
              //Result := JaiLeDroitTag(31453); // Générer manuellement une réception à partir d'une commande
              //if result then Result := JaiLeDroitTag(31521); // Générer automatiquement une réception à partir d'une commande
              Result := JaiLeDroitTag(AchGeneManuLivCde); // Générer manuellement une réception à partir d'une commande
              if result then Result := JaiLeDroitTag(AchGeneAutoLivCde); // Générer automatiquement une réception à partir d'une commande
            end
            else if NatureOrig = 'PRF' then
            Begin
              Result := JaiLeDroitTag(31455); // Générer manuellement une réception à partir d'une préparation
              if result then Result := JaiLeDroitTag(31523); // Générer automatiquement une réception à partir d'une préparation
            end
        end
      else if Nature = 'AF' then
        begin
          Result := JaiLeDroitTag(AvAch); //avoir fournisseur financier
          if Result then
            if NatureOrig = 'FF' then
              //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
              //Result := JaiLeDroitTag(31543);    // Générer Facture frn en avoir financier
              Result := JaiLeDroitTag(AchGeneAvFac);    // Générer Facture frn en avoir financier
        end
      else if Nature = 'AFP' then
        begin
          Result := JaiLeDroitTag(AvAchStock); //avoir fournisseur valorisé
          if Result then
            if NatureOrig = 'BFA' then
            begin
              //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
              //Result := JaiLeDroitTag(31542);   // Générer Retour en avoir fournisseur valorisé
              Result := JaiLeDroitTag(AchGeneAvStockRetour);   // Générer Retour en avoir fournisseur valorisé
              if Result then
              begin
                //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
                //Result := JaiLeDroitTag(31459); // Générer manuellement un avoir fournisseur valorisé à partir d'un retour
                //if result then Result := JaiLeDroitTag(31528); // Générer automatiquement un avoir fournisseur valorisé à partir d'un retour
                Result := JaiLeDroitTag(AchGeneManuAvStockRetour); // Générer manuellement un avoir fournisseur valorisé à partir d'un retour
                if result then Result := JaiLeDroitTag(AchGeneAutoAvStockRetour); // Générer automatiquement un avoir fournisseur valorisé à partir d'un retour
              end
            end
        end
      else if Nature = 'AFS' then
        begin
          Result := JaiLeDroitTag(AvAchStockPrix); //avoir fournisseur sur stock
          if Result then
            if NatureOrig = 'FF' then
              //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
              //Result := JaiLeDroitTag(31544)    // Générer Facture frn en avoir fournisseur sur stock
              Result := JaiLeDroitTag(AchGeneAvStockPrixFact)    // Générer Facture frn en avoir fournisseur sur stock
            else if NatureOrig = 'BFA' then
            begin
              //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
              //Result := JaiLeDroitTag(31541);   // Générer Retour en avoir fournisseur sur stock
              Result := JaiLeDroitTag(AchGeneAvStockPrixRetour);   // Générer Retour en avoir fournisseur sur stock
              if Result then
              begin
                //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
                //Result := JaiLeDroitTag(31458); // Générer manuellement un avoir sur stock à partir d'un retour
                //if result then Result := JaiLeDroitTag(31527); // Générer automatiquement un avoir sur stock à partir d'un retour
                Result := JaiLeDroitTag(AchGeneManuAvStockPrixRetour); // Générer manuellement un avoir sur stock à partir d'un retour
                if result then Result := JaiLeDroitTag(AchGeneAutoAvStockPrixRetour); // Générer automatiquement un avoir sur stock à partir d'un retour
              end
            end
        end
      else if Nature = 'BFA' then Result := JaiLeDroitTag(RetourFournisseur) // Retour fournisseur
      else if Nature = 'FF' then
        begin
          Result := JaiLeDroitTag(AchFact); //Facture
          //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
          //if Result then Result := JaiLeDroitTag(31512); // Générer Facture
          if Result then Result := JaiLeDroitTag(AchGeneFact); // Générer Facture
          if Result then
            if NatureOrig = 'CF' then
            Begin
              //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
              //Result := JaiLeDroitTag(31454); // Générer manuellement une facture à partir d'une cde
              //if result then Result := JaiLeDroitTag(31525); // Générer automatiquement une facture à partir d'une cde
              Result := JaiLeDroitTag(AchGeneManuFactCde); // Générer manuellement une facture à partir d'une cde
              if result then Result := JaiLeDroitTag(AchGeneAutoFactCde); // Générer automatiquement une facture à partir d'une cde
            end
            else if NatureOrig = 'PRF' then
            Begin
              Result := JaiLeDroitTag(31456); // Générer manuellement une facture à partir d'une Prépa Livr
              if result then Result := JaiLeDroitTag(31526); // Générer automatiquement une facture à partir d'une Prépa Livr
            end
            else if NatureOrig = 'BLF' then
            Begin
              //SBU - FQ 10175 - utilisation de constantes pour certain tag de menu
              //Result := JaiLeDroitTag(31457); // Générer manuellement une facture à partir d'une Livraison
              //if result then Result := JaiLeDroitTag(31522); // Générer automatiquement une facture à partir d'une Livraison
              Result := JaiLeDroitTag(AchGeneManuFactLiv); // Générer manuellement une facture à partir d'une Livraison
              if result then Result := JaiLeDroitTag(AchGeneAutoFacLiv); // Générer automatiquement une facture à partir d'une Livraison
            end
        end
      else Result := true;
    end;
    //CCMX-CEGID - FQ 13344 - FIN
  end;

begin
    Result := True;
{$IFDEF GESCOM}
    if ctxGescom in V_PGI.PGIContexte then
      Result := AffecteTag (30201,30202,30241,30242,30209,30251,30252,30210,30205,30261,30262,30263,36101,
      //SBU - FQ 10175 - DEBUT - en CHR le menu ACHAT est le 163
      {$IFDEF CHR}
                            163121,16313,16314,163161,163162,163163,16315
      {$ELSE CHR}
                            31201,31202,31203,31231,31232,31233,31234
      {$ENDIF}
      //SBU - FQ 10175 - FIN
    //CCMX-CEGID - FQ 13344 - DEBUT - Ajout des préparations de réceptions
    //                                      des cdes de sous traitance
                            //);
                            ,31204,211611);
    //CCMX-CEGID - FQ 13344 - FIN
{$ENDIF}
{$IFDEF AFFAIRE}
    if ctxAffaire in V_PGI.PGIContexte then  //GA_200802_AB-GA14525
      Result := AFJaiLeDroitNatureGCCreat(Nature, Typepiece, NatureOrig);
{$ENDIF}
    if (ctxMode in V_PGI.PGIContexte) {or (ctxAffaire in V_PGI.PGIContexte)} then
      Result := True;
end;

function CleTelForFind (strTelephone:string):string;
var
   car   :string;
   i     :integer;
begin
  Result := '';
  strTelephone := Trim (strTelephone);
  if copy(StrTelephone,1,1)='+' then
  begin
  	strTelephone := Copy(StrTelephone,4,255);
  end;

  for i := 1 to length (strTelephone) do
  begin
    car := copy (strTelephone, i, 1);
    if (IsNumeric (car)) and (car <>'.') and (car<>',') and (car<>'-') and (car<>'+') and (car<>' ') and (car <>'(') and (car <> ')') then
    	Result := Result + car;
  end;

  if copy(Result,1,1)='0' then
  begin
  	Result := Copy(Result,2,255);
  end;
end;

{**************************************************************************************
Calcul de la clé téléphone: les 9 derniers caractères numériques d'un n° de téléphone
***************************************************************************************}
function CleTelephone (strTelephone:string; bFixedLen:boolean=TRUE):string;
var
   car   :string;
   i     :integer;
begin
     Result := '';
     strTelephone := Trim (strTelephone);
     for i := 1 to length (strTelephone) do
     begin
          car := copy (strTelephone, i, 1);
          if (IsNumeric (car)) and (car <>'.') and (car<>',') and (car<>'-') and (car<>'+') and (car<>' ') and (car <>'(') and (car <> ')') then
             Result := Result + car;
     end;

     // Si clé véritable (donc toujours 9 car.), on construit une chaine de 9 caractères. Sinon, on laisse telle quelle
     if bFixedLen = TRUE then
     begin
          i := Length (Result);
          if i > 9 then
              Result := Copy (Result, i-8, 9)
          else
              if i < 9 then
                 Result := StringOfChar ('0', 9-i) + Result;
     end;
end;

// $$$ JP 18/07/05 - que pour le bureau pour l'instant
//{$IFNDEF ERADIO}
{$IFDEF BUREAU}
{**************************************************************************************
Màj de tous les enregistrements contenant un champ clé téléphone
***************************************************************************************}
function MajCleTelephone (const strChampCle, strChampTel, strTable:string):boolean;
var
   TOBTel        :TOB;
   i             :integer;
   strValeurTel  :string;
begin
     //$$$ JP 06/12/05 - warning delphi -> Result := FALSE;
     TOBTel := TOB.Create ('les tels', nil, -1);
     try
        TOBTel.LoadDetailFromSQL ('SELECT DISTINCT ' + strChampTel + ' FROM ' + strTable);
        for i := 0 to TOBTel.Detail.Count-1 do
        begin
             strValeurTel := TOBTel.Detail [i].GetString (strChampTel);
             ExecuteSQL ('UPDATE ' + strTable + ' SET ' + strChampCle + '="' + CleTelephone (strValeurTel) + '" WHERE ' + strChampTel + '="' + strValeurTel + '"');
        end;
        Result := TRUE;
     finally
            TOBTel.Free;
     end;
end;
{$ENDIF}
//{$ENDIF !ERADIO}

{***********A.G.L.Privé.*****************************************
Version PGI de PCL_IMPORT_BOB sans contrainte sur V_PGI_ENV
*****************************************************************}
function PGI_IMPORT_BOB(CodeProduit:string): Integer;
{$IFNDEF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
var sFileBOB                :string;
    Chemin                  :string;
    SearchRec               :TSearchRec;
    NumVersion              :integer;
    ret                     :integer;
    RetInt : integer;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF !EAGLCLIENT}
BEGIN
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version 999
  // - extension .BOB
  // - exemple CCS50582F0001.BOB
  //
  // CODE RETOUR DE LA FONCTION
  Result := 0;
{$IFNDEF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFDEF BTP}
  Chemin := TCBPPath.GetCegidDistriBob + '\BAT2008\';
  ret := FindFirst(Chemin+'*.BOB', faAnyFile, SearchRec);
  NumVersion := 0;
  while ret = 0 do
  begin
       //RECUPERE NOM DU BOB
       sFileBOB := SearchRec.Name;
       //RECUPERE NUM VERSION
       try
          NumVersion := ValeurI(Copy(sFileBOB,5,4));
       except
          Result := -6;
          exit;
       end;

       if NumVersion > V_PGI.NumVersionBase then
       begin
            RetInt := AGLIntegreBob(Chemin + sFileBOB,FALSE,TRUE);
            case RetInt of
//            case TestAGLIntegreBob(Chemin + sFileBOB) of
            0  :// OK
                begin
                     if V_PGI.SAV then Pgiinfo('Intégration de : '+sFileBOB, TitreHalley);//Resultif not LIA_JOURNAL_EVENEMENT(sTempo) then Result := -1;
                     if copy(sFileBob,9,1) = 'M' then Result := 1; //SI BOB AVEC MENU, ON REND 1 POUR SORTIR DE L'APPLICATION
                end;
//Non en line
            1  : if V_PGI.SAV then Pgiinfo('Intégration déjà effectuée :'+sFileBOB, TitreHalley);// Intégration déjà effectuée
           -1  :// Erreur d'écriture dans la table YMYBOBS
                begin
                     if V_PGI.SAV then PGIInfo('Erreur d''écriture dans la table YMYBOBS :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
           -2  :// Erreur d'intégration dans la fonction AglImportBob
                begin
                     if V_PGI.SAV then PGIInfo('Erreur d''intégration dans la fonction AglImportBob :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
           -3  ://Erreur de lecture du fichier BOB.
                begin
                     if V_PGI.SAV then PGIInfo('Erreur de lecture du fichier BOB :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
           -4  :// Erreur inconnue.
                begin
                     if V_PGI.SAV then PGIInfo('Erreur inconnue :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
            end;

       end;
       ret := FindNext(SearchRec);
   end;
   sysutils.FindClose(SearchRec);
{$ELSE}
  Chemin := TCBPPath.GetCegidDistriBob + '\CGSX\';
  ret := FindFirst(Chemin+ 'CGSX' +'*.BOB', faAnyFile, SearchRec);
  NumVersion := 0;
  while ret = 0 do
  begin
       //RECUPERE NOM DU BOB
       sFileBOB := SearchRec.Name;
       //RECUPERE NUM VERSION
       try
          NumVersion := ValeurI(Copy(sFileBOB,5,4));
       except
          Result := -6;
          exit;
       end;

       if NumVersion > V_PGI.NumVersionBase then
       begin
            case AGLIntegreBob(Chemin + sFileBOB,FALSE,TRUE) of
//            case TestAGLIntegreBob(Chemin + sFileBOB) of
            0  :// OK
                begin
                     if V_PGI.SAV then Pgiinfo('Intégration de : '+sFileBOB, TitreHalley);//Resultif not LIA_JOURNAL_EVENEMENT(sTempo) then Result := -1;
                     if copy(sFileBob,9,1) = 'M' then Result := 1; //SI BOB AVEC MENU, ON REND 1 POUR SORTIR DE L'APPLICATION
                end;
            1  : if V_PGI.SAV then Pgiinfo('Intégration déjà effectuée :'+sFileBOB, TitreHalley);// Intégration déjà effectuée
           -1  :// Erreur d'écriture dans la table YMYBOBS
                begin
                     if V_PGI.SAV then PGIInfo('Erreur d''écriture dans la table YMYBOBS :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
           -2  :// Erreur d'intégration dans la fonction AglImportBob
                begin
                     if V_PGI.SAV then PGIInfo('Erreur d''intégration dans la fonction AglImportBob :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
           -3  ://Erreur de lecture du fichier BOB.
                begin
                     if V_PGI.SAV then PGIInfo('Erreur de lecture du fichier BOB :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
           -4  :// Erreur inconnue.
                begin
                     if V_PGI.SAV then PGIInfo('Erreur inconnue :'+Chemin + sFileBOB,'PCL_IMPORT_BOB');
                end;
            end;

       end;
       ret := FindNext(SearchRec);
   end;
   sysutils.FindClose(SearchRec);
{$ENDIF !BTP}
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
{$ENDIF !EAGLCLIENT}
END;

{JP 31/07/06 : Chargement du NoDossier en mode PGE}
{---------------------------------------------------------------------------------------}
procedure ChargeNoDossier;
{---------------------------------------------------------------------------------------}

    {-------------------------------------------------------------------}
    function _GetListeChamp : string;
    {-------------------------------------------------------------------}
    begin
      Result := ' BQ_ADRESSE1,BQ_ADRESSE2,BQ_ADRESSE3,BQ_AGENCE,BQ_BANQUE,BQ_CALENDRIER,BQ_CLERIB,BQ_CODE,BQ_CODEBIC,BQ_CODECIB,BQ_CODEIBAN,BQ_CODEPOSTAL,BQ_COMMENTAIRE,BQ_COMPTEFRAIS,BQ_CONTACT,BQ_CONVENTIONLCR,BQ_DATEDERNSOLDE,BQ_DELAIBAPLCR,';
      Result := Result + 'BQ_DELAILCR,BQ_DELAIPRELVACC,BQ_DELAIPRELVORD,BQ_DELAITRANSINT,BQ_DELAIVIRBRULANT,BQ_DELAIVIRCHAUD,BQ_DELAIVIRORD,BQ_DERNSOLDEDEV,BQ_DERNSOLDEFRS,BQ_DESTINATAIRE,BQ_DEVISE,BQ_DIVTERRIT,BQ_DOMICILIATION,BQ_ECHEREPLCR,BQ_ECHEREPPRELEV,BQ_ENCOURSLCR,';
      Result := Result + 'BQ_ETABBQ,BQ_FAX,BQ_GENERAL,BQ_GUICHET,BQ_GUIDECOMPATBLE,BQ_INDREMTRANS,BQ_JOURFERMETUE,BQ_LANGUE,BQ_LETTRECHQ,BQ_LETTRELCR,BQ_LETTREPRELV,BQ_LETTREVIR,BQ_LIBELLE,BQ_MULTIDEVISE,BQ_NATURECPTE,BQ_NODOSSIER,BQ_NUMEMETLCR,BQ_NUMEMETPRE,BQ_NUMEMETVIR,';
      Result := Result + 'BQ_NUMEROCOMPTE,BQ_PAYS,BQ_PLAFONDLCR,BQ_RAPPAUTOREL,BQ_RAPPROAUTOLCR,BQ_RELEVEETRANGER,BQ_REPBONAPAYER,BQ_REPIMPAYELCR,BQ_REPIMPAYEPRELV,BQ_REPLCR,BQ_REPLCRFOURN,BQ_REPPRELEV,BQ_REPRELEVE,BQ_REPVIR,BQ_SOCIETE,BQ_TELEPHONE,BQ_TELEX,';
      Result := Result + 'BQ_TYPEREMTRANS,BQ_VILLE ';
      if isMssql then Result := Result + ',BQ_BLOCNOTE ';
    end;

var
  LeDos : string;
  Q     : TQuery;
begin
  {Mise à jour de V_PGI.NoDossier en mode PGE}
  if not (ctxPCL in V_PGI.PGIContexte) then begin
    LeDos := GetParamSocSecur ('SO_NODOSSIER', '');
    if LeDos <> '' then V_PGI.NoDossier := LeDos
                   else V_PGI.NoDossier := '000000';
  end;
  {Théoriquement plus nécessaire à partir du PGIMAJVER 828
   JP 26/10/07 : malheureusement la demande 682 a été mal interprétée dans MajVer828. voir
                 mail de ce jour}
  (*
  {Mise à jour de BQ_NODOSSIER, si on n'est pas en mode Tréso, sinon, ce sera fait dans la Tréso :
   JP 13/10/06 : reste le cas PGE sans Tréso à Gérer et ce comment ??
                 (not GetParamSocSecur('SO_MODETRESO',False)) or}
  if (ctxPCL in V_PGI.PGIContexte) and
     (not GetParamSocSecur('SO_TRMOULINEBQCODE', False)) then begin
    ExecuteSQL('UPDATE BANQUECP SET BQ_NODOSSIER = "' + V_PGI.NoDossier + '" WHERE BQ_NODOSSIER = "000000" ' +
               'OR BQ_NODOSSIER = "" OR BQ_NODOSSIER IS NULL');
    {Mise à jour du ParamSoc pour signifier que le traitement a été fait}
    SetParamSoc('SO_TRMOULINEBQCODE', True);

  end
  *)
  {JP 12/01/07 : Mise à jour du champ RB_RUBRIQUE, car il est initialisé avec V_PGI.NoDossier : si cela ne
                 pose pas de problèmes en PCL, car V_PGI.NoDossier a toujours été initialisé, mais en PGE
                 c'est plus problèmatique surtout que RB_NODOSSIER appartient à la clef primaire}
  {else
   JP 26/10/07 : le PGIMajVer ne gère pas ce cas (demande mal prise en compte)}
  if not EstComptaTreso and not (ctxPCL in V_PGI.PGIContexte) and
     (not GetParamSocSecur('SO_TRNODOSSIERCOMPTA', False)) then begin

    if EstTablePartagee('BANQUECP') then begin
      if isMssql then LeDos := V_PGI.SchemaName + '.DBO.BANQUECP'
                 else LeDos := V_PGI.SchemaName + '.BANQUECP';
    end
    else
      LeDos := 'BANQUECP';
    {JP 09/01/07 : FQ 22146 : Gérer dans le Mul de suppression des Rubriques
    ExecuteSQL('UPDATE RUBRIQUE SET RB_NODOSSIER = "' + V_PGI.NoDossier + '" WHERE RB_NODOSSIER = "000000" ' +
               'OR RB_NODOSSIER = "" OR RB_NODOSSIER IS NULL');}
    ExecuteSQL('UPDATE ' + LeDos + ' SET BQ_NODOSSIER = "' + V_PGI.NoDossier + '" WHERE BQ_NODOSSIER = "000000" ' +
               'OR BQ_NODOSSIER = "" OR BQ_NODOSSIER IS NULL');
    ExecuteSQL('UPDATE ' + LeDos + ' SET BQ_NATURECPTE = "BQE" WHERE ' +
               'OR BQ_NATURECPTE = "" OR BQ_NATURECPTE IS NULL');
    if EstTablePartagee('BANQUECP') then begin
      BeginTrans;
      try
        {Fusion de BanqueCP}
        Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE = "BANQUECP"', True);
        {Théoriquement, il ne devrait pas y avoir d'else ...}
        if not Q.EOF then begin
          {Création d'une clef unqiue}
          ExecuteSQL('UPDATE ' + LeDos + ' SET BQ_CODE = SUBSTRING(BQ_GENERAL, 1, 12)||"' + StrRight(V_PGI.NoDossier, 5) + '" ' +
                     'WHERE BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');
          {On n'est pas dans la base de partage du référentiel}
          if Q.FindField('DS_NOMBASE').AsString <> V_PGI.SchemaName then begin
            {Reprise des données vers la table de partage}
            ExecuteSQL('INSERT INTO BANQUECP (' + _GetListeChamp + ') SELECT ' + _GetListeChamp + ' FROM ' + LeDos);
            {Suppression des données dans la base d'origine}
            ExecuteSQL('DELETE FROM ' + LeDos);
          end;
        end;
        CommitTrans;
      except
        on E : Exception do begin
          PGIError(TraduireMemoire('Erreur lors de la mise à jour les comptes bancaires avec le message :') +
                   #13#13 + E.Message, TraduireMemoire('Mise à jour du N° Dossier'));
          RollBack;
          Exit;
        end;
      end;
    end;
    {Mise à jour du ParamSoc pour signifier que le traitement a été fait}
    SetParamSoc('SO_TRNODOSSIERCOMPTA', True);
  end;
end;

{$IFDEF CHR} { CHR_AGF_20080417}
function PGIEnvoiMail(SUJET: hstring; AQUI, CopieA: string; Corps: hTStrings; FICHIERS: string; EnvoiAuto: boolean = TRUE; Importance: Integer = 1; CATEGORIE: string = ''; COMPANIE: string = ''; SansAlerte: boolean = FALSE): boolean;
var
{$IFNDEF EAGLSERVER}
  ResultMailForm : TResultMailForm;
  Liste :hTStringList;
{$ELSE}
  sServer,
  sFrom,
  sCorps: string;
{$ENDIF !EAGLSERVER}
begin
  {$IFNDEF EAGLSERVER}
  Result := True;
  try
  if not EnvoiAuto and ( not v_pgi.ficheoutlook or (V_PGI.MailMethod = mmNotes) ) then
  begin
      Liste:=hTStringList.Create ;
      try
        if Corps = Nil then
        begin
          Liste.Add('');
        end
        else
        begin
          Liste.Text := Corps.text;
        end;
        ResultMailForm := AglMailForm(Sujet, Aqui, CopieA, Liste, FICHIERS);
        EnvoiAUto:=True;
        if ResultMailForm = rmfOkButNotSend then SendMail( SUJET, AQUI, CopieA, Liste, FICHIERS, EnvoiAuto, Importance, CATEGORIE,COMPANIE, SansAlerte );
      finally
        Liste.free;
      end;
  end
  else SendMail( SUJET, AQUI, CopieA, Corps, FICHIERS, EnvoiAuto, Importance, CATEGORIE,COMPANIE, SansAlerte );
  except
    on E: Exception do
    begin
      if( V_PGI.SAV ) then
        PGIError( E.Message );
      Result := False;
    end;
  end;
  {$ELSE}
  sServer := GetParamSocSecur( 'SO_MBOSMTPSERVER', V_PGI.SMTPServer );
  sFrom := GetParamSocSecur( 'SO_MBOSMTPFROM', V_PGI.SMTPFrom );
  if( Assigned(Corps) ) then
    sCorps := Corps.Text;
  Result := SendMailSmtp( sServer, sFrom, AQui, Sujet, sCorps, Fichiers );
  {$ENDIF !EAGLSERVER}
End;

{$ELSE}
procedure PGIEnvoiMail(SUJET: hstring; AQUI, CopieA: string; Corps: hTStrings; FICHIERS: string; EnvoiAuto: boolean = TRUE; Importance: Integer = 1; CATEGORIE: string = ''; COMPANIE: string = ''; SansAlerte: boolean = FALSE);
{$IFNDEF EAGLSERVER}
var ResultMailForm : TResultMailForm;
    Liste :hTStringList;
{$ENDIF !EAGLSERVER}
begin
  {$IFNDEF EAGLSERVER}
  if not EnvoiAuto and ( not v_pgi.ficheoutlook or (V_PGI.MailMethod = mmNotes) ) then
  begin
      Liste:=hTStringList.Create ;
      try
        if Corps = Nil then
        begin
          Liste.Add('');
        end
        else
        begin
          Liste.Text := Corps.text;
        end;
        ResultMailForm := AglMailForm(Sujet, Aqui, CopieA, Liste, FICHIERS);
        EnvoiAUto:=True;
        if ResultMailForm = rmfOkButNotSend then SendMail( SUJET, AQUI, CopieA, Liste, FICHIERS, EnvoiAuto, Importance, CATEGORIE,COMPANIE, SansAlerte );
      finally
        Liste.free;
      end;
  end
  else SendMail( SUJET, AQUI, CopieA, Corps, FICHIERS, EnvoiAuto, Importance, CATEGORIE,COMPANIE, SansAlerte );
  {$ENDIF !EAGLSERVER}
End;
{$ENDIF CHR}

procedure AGLPGIEnvoiMail( parms: array of variant; nb: integer ) ;
Var T : hTStringList ;
begin
  T:=hTStringList.Create ;
  T.Text:=string(parms[3]) ;
  PGIEnvoiMail(parms[0],parms[1],parms[2],T,parms[4],parms[5]) ;
  T.Free ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 26/05/2005
Modifié le ... :   /  /    
Description .. : Utilisé dans les sources pour isoler les implémentations
Suite ........ : spécifiques au multisociété
Mots clefs ... : MULTISOC
*****************************************************************}
Function EstMultiSoc : Boolean ;
begin
{$IFDEF EAGLSERVER}
 result := false ;
{$ELSE EAGLSERVER}
  Result := not (ctxPCL in V_PGI.PGIContexte) and ( high(V_PGI.DEShares) > 0 ) ;
{$ENDIF !EAGLSERVER} 
end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 26/05/2005
Modifié le ... : 26/05/2005
Description .. : Retourne vrai si la table vNomTable fait partie d'un
Suite ........ : partage MULTISOC
Suite ........ : (Part du principe que DESHARE est renseigné dans toutes
Suite ........ : les bases du regroupement multisocété )
Mots clefs ... : MULTISOC
*****************************************************************}
Function EstTablePartagee( vNomTable : String ) : Boolean ;
begin
  if EstMultiSoc
    then Result := TableToBaseNum( vNomTable ) > 0
    else Result := False ;
end ;

{JP 24/10/07 : gère les paramsoc partagés outre le dossier}
{---------------------------------------------------------------------------------------}
function  GetParamSocDossierMS(Nom : string; Defaut : Variant; Dossier : string) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if TableToBaseNum(Nom, 'PAR') > 0 then Result := GetParamsocSecur(Nom, Defaut)
                                    else Result := GetParamsocDossierSecur(Nom, Defaut, Dossier);
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 26/05/2005
Modifié le ... : 26/05/2005
Description .. : Retourne la valeur utilisée dans la table cumuls pour 
Suite ........ : identifiée le type d'entité auquel il se rapporte :
Suite ........ : champ CU_TYPE
Mots clefs ... : MULTISOC
*****************************************************************}
function fbToCumulType( vFB : TFichierBase ) : String ;
begin
  case vFB of
   fbGene : result := 'MSG' ;
   fbAux  : result := 'MST' ;
   fbSect : result := 'MSS' ;
   fbJal  : result := 'MSJ' ;
   else result := '' ;
   end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 30/05/2005
Modifié le ... :   /  /
Description .. : Retourne le nom de la table correspondant au TfichierBase
Mots clefs ... : 
*****************************************************************}
function  fbToTable        ( vFB : TFichierBase ) : String ;
begin
  case vFB of
   fbGene : result := 'GENERAUX' ;
   fbAux  : result := 'TIERS' ;
   fbSect : begin
              if VH^.LiaisonY2ViaShare then result := 'CSECTION'
                                       else result := 'SECTION' ;
            end;
   fbAxe1..fbAxe5 : begin
                      if VH^.LiaisonY2ViaShare then result := 'CSECTION'
                                               else result := 'SECTION' ;
                    end;
   fbJal  : result := 'JOURNAL' ;
   else result := '' ;
   end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 30/05/2005
Modifié le ... : 19/07/2005
Description .. : Retourne la nom de la table préfixée par le nom du dossier
Suite ........ : si la table n'est pas partagée.
Mots clefs ... : MULTISOC
*****************************************************************}
Function  GetTableDossier( vDossier, vNomTable : String ) : String ;
begin
  result := vNomTable ;

  if (Trim(vDossier)='') or (vDossier=V_PGI.SchemaName) then Exit ;

  // Si la table est partagée, on laisse la main à l'AGL...
  if EstTablePartagee( vNomTable ) then Exit ;

{$IFDEF EAGLSERVER}
  if (V_PGI.Driver = dbMSSQL) or (V_PGI.Driver=dbMSSQL2005)
{$ELSE}
  if isMssql // de wCommuns non compatible eAglServer pour le moment
{$ENDIF EAGLSERVER}
    then result := vDossier + '.DBO.' + vNomTable
    else result := vDossier + '.' + vNomTable ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 19/12/2006
Modifié le ... :   /  /    
Description .. : en multi société, permet d'ajouter le préfixe sur la base en 
Suite ........ : fonction du driver sql
Mots clefs ... : MULTISOC;BASE;
*****************************************************************}
function GetBase (UneBase, LaTable : string) : String;
begin
  if isMssql then result := UneBase + '.DBO.' + LaTable
  else result := UneBase + '.' + LaTable;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 19/07/2005
Modifié le ... : 19/07/2005
Description .. : Retourne la liste des bases du regroupement utilisé pour le
Suite ........ : paramétrage du mutlisociété, sous la forme d'une chaîne
Suite ........ : séparée par des ';'
Suite ........ :
Suite ........ : Code du regroupement MultiSoc stocké dans la constante
Suite ........ : MS_CODEREGROUPEMENT
Mots clefs ... :
*****************************************************************}
Function GetBasesMS (CodeRegroupement : string = ''; BaseSql : boolean = True) : String ;
var lQMD     : TQuery ;
    lTSData  : TStringList ;
    lStVal   : String ;
    lStBase  : String ;
begin

  result := '' ;

  if CodeRegroupement = '' then CodeRegroupement := MS_CODEREGROUPEMENT;

  // récupération paramétrage du regroupement
  lStVal := '' ;
  lQMD   := OpenSQL('SELECT * FROM YMULTIDOSSIER WHERE YMD_CODE = "' + CodeRegroupement + '"', True ) ;
  if not lQMD.Eof then
    lStVal := lQMD.FindField('YMD_DETAILS').AsString ;
  Ferme( lQMD ) ;
  if lStVal = '' then Exit ;

  // Récupération 1ère ligne
  lTSData      := TStringList.Create ;
  lTSData.Text := lStVal ;
  lStVal       := lTSData.Strings[0] ;

  // On ne garde que le nom des bases
  while lStVal<>'' do
    begin
    lStBase := ReadTokenSt( lStVal ) ;
    if BaseSql then
    begin
      ReadTokenPipe( lStBase , '|' );
      result := result + lStBase + ';';
    end else
      result  := result + ReadTokenPipe( lStBase , '|' ) + ';' ;
    end ;

  FreeAndNil( lTSData ) ;

end ;

Function  PresenceMS ( vFichier : String ; vChamp : String ; vValeur : String ) : Boolean ;
var lStBases   : String ;
    lStDossier : String ;
begin
  result := False ;
  if not EstMultiSoc then Exit ;

  // Si la table est partagée, appel de la fonction standard, on laisse faire l'AGL...
  if EstTablePartagee( vFichier ) then
    begin
    result := Presence( vFichier, vChamp, vValeur ) ;
    end
  // Sinon on parcours les bases du regroupement
  else
    begin
    lStBases   := GetBasesMS ;
    lStDossier := ReadTokenSt( lStBases ) ;
    While lStDossier <> '' do
      begin
      result := ExisteSQL( 'SELECT ' + vChamp + ' FROM ' + GetTableDossier( lStDossier,  vFichier ) +
                           ' WHERE ' + vChamp + ' = "' + vValeur + '"' ) ;
      {JP 17/05/06 : Presence semble enlever le DBO !!
      result := Presence( GetTableDossier( lStDossier, vFichier ), vChamp, vValeur ) ;}
      if result then Exit ;
      lStDossier := ReadTokenSt( lStBases ) ;
      end ;
    end ;

end ;

Function  PresenceComplexeMS ( vFichier : String ; vChamps : Array of hString ; vComps : Array of hString ; vValeurs : Array of hString ; vTypes : Array of hString ) : Boolean ;
var lStBases : String ;
    lStDossier : String ;
begin
  result := False ;
  if not EstMultiSoc then Exit ;

  // Si la table est partagée, appel de la fonction standard, on laisse faire l'AGL...
  if EstTablePartagee( vFichier ) then
    begin
    result := PresenceComplexe( vFichier, vChamps, vComps, vValeurs, vTypes ) ;
    end
  else
  // Sinon on boucle sur les bases du regroupement
    begin
    lStBases   := GetBasesMS ;
    lStDossier := ReadTokenSt( lStBases ) ;
    While lStDossier <> '' do
      begin
//      result := ExisteSQL( 'SELECT ' + vChamp + ' FROM ' + GetTableDossier( vFichier ) + ' WHERE ' + vChamp + '="' + vValeur + '"' ) ;
      result := PresenceComplexe( GetTableDossier( lStDossier, vFichier ) , vChamps, vComps, vValeurs, vTypes ) ;
      if result then Exit ;
      lStDossier := ReadTokenSt( lStBases ) ;
      end ;
    end ;
end ;

function EstBaseMultiSoc : boolean;
begin
  Result := ExisteSql ('SELECT 1 FROM YMULTIDOSSIER WHERE YMD_CODE="' + MS_CODEREGROUPEMENT + '"');
end;

function TablePartagee (LaTable : string) : boolean;
begin
  Result := EstTablePartagee (LaTable);
end;

function RechDomZoneLibre (sValue : string; bAbrege : boolean; sPlus : string = ''; bLibre : boolean = false; sPrefixe : string='') : string;
begin
  if (sPrefixe = '') then // PL le 16/05/07 : on estime que seul le code est signifiant
  begin
  if pos (Copy (sValue, 1, 2), 'AM;AT;AD;AC;AB;AS') > 0 then
    Result := rechdom ('GCZONELIBREART', sValue, bAbrege, sPlus, bLibre)
  else if pos (Copy (sValue, 1, 2), 'WM;WT;WD;WC;WB') > 0 then
    Result := rechdom ('GCZONELIBRESAV', sValue, bAbrege, sPlus, bLibre)
  else if pos (Copy (sValue, 1, 2), 'CM;CT;CD;CC;CB;CR;FD;FM;FT') > 0 then   //mcd 16/08/05 oubli CR
    Result := rechdom ('GCZONELIBRETIE', sValue, bAbrege, sPlus, bLibre)
  else if pos (Copy (sValue, 1, 2), 'BT;BC;BD;BB;BM') > 0 then
    Result := rechdom ('GCZONELIBRECON', sValue, bAbrege, sPlus, bLibre)
  else if pos (Copy (sValue, 1, 2), 'DB;DC;DD;DM;DT') > 0 then
    Result := rechdom ('GCZONELIBREDEP', sValue, bAbrege, sPlus, bLibre)
  else if pos (Copy (sValue, 1, 2), 'TL') > 0 then
    Result := rechdom ('GCZONELIBREDEP', sValue, bAbrege, sPlus, bLibre)
  else
    Result := RechDom ('GCZONELIBRE', sValue, bAbrege, sPlus, bLibre);

  end
  else
    // PL le 16/05/07 : sinon, il faut le préfixe pour préciser la tablette associée
    begin
    if sPrefixe = 'RPE' then
      Result := rechdom ('RTLIBPERSPECTIVE', sValue, bAbrege, sPlus, bLibre)
    else  
    if sPrefixe = 'RPJ' then
      Result := rechdom ('RTLIBPROJET', sValue, bAbrege, sPlus, bLibre)
    else  
    if sPrefixe = 'ROP' then
      Result := rechdom ('RTLIBOPERATION', sValue, bAbrege, sPlus, bLibre)
    else
    ;
    end;
end;

function IsMESSeria: boolean;
begin
  Result := ( ctxGPAO in V_PGI.PGIContexte );
end;

function IsMESAutonome: boolean;
begin
  Result := False;
end;

function GereNoDossier : boolean;
var
  GuidDossier : string;
begin
  Result := true;
  V_PGI.NoDossier := '00000000';
  if EstMultiSoc then
  begin
    GuidDossier := GetParamSocSecur ('SO_GUIDDOSSIER', '');
    V_PGI.NoDossier := GetParamSocSecur ('SO_NODOSSIER', '');
    Result := ExisteSQL ('SELECT 1 FROM DOSSIER WHERE DOS_NODOSSIER="' + V_PGI.NoDossier + '" AND ' +
                            ' DOS_GUIDDOSSIER="' + GuidDossier + '"');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/10/2005
Modifié le ... :   /  /    
Description .. : Indique si la base courante est une base PCL allégée
Mots clefs ... : 
*****************************************************************}
function EstBasePclAllegee : Boolean;
begin
  Result := False;
  try
    Result := Not ExisteSQL('SELECT 1 FROM DECHAMPS WHERE DH_PREFIXE="AFF"');
  except
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/03/2006
Modifié le ... :   /  /
Description .. : Indique si la base courante est une base PCL optimisée
Mots clefs ... :
*****************************************************************}
function EstBasePclOptimisee : Boolean;
begin
  Result := False;
  try
    Result := Not ExisteSQL('SELECT 1 FROM DETABLES WHERE DT_PREFIXE="AFF"');
  except
  end;
end;

{$IFDEF MAJPCL}
procedure InitPGIpourDossierPCL;
begin
    V_PGI.PCLDomainesExclus :='"A","0","B","E","G","H","O","Q","R","U","W","X"';
    // CA - 01/09/2005
    V_PGI.PclDomainesVuesExclus :='"A","0","B","E","G","H","O","Q","R","U","W","X"';
    // Fin CA
    //Christophe AYEL Demande n° 1246 et 1304
    V_PGI.PCLModulesInclus := '2,4,6,17,18,24,25,26,27,37,38,39,41,42,43,44,45,46,47,48,49,52,53,54,55,56,61,64,66,69,81,96,155,156,157,158,159,171,173,175,176,177,200,270,272,273,276,303,326,360,362';
    V_PGI.PCLParamSocInclus := '(SOC_TREE LIKE "001;000%") OR (SOC_TREE LIKE "001;001%") OR (SOC_TREE LIKE "001;002%") OR (SOC_TREE LIKE "001;003%") OR (SOC_TREE LIKE "001;005%") OR (SOC_TREE LIKE "001;018%") OR (SOC_TREE LIKE "001;026%") OR (SOC_TREE LIKE "001;027%")';
end;
{$ENDIF}

{ GC/GRC : JTR / - eQualité 13254}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 25/03/2005
Modifié le ... : 25/03/2005
Description .. : OpenSQL qui retourne un  TQuery dans tous les cas
Mots clefs ... :
*****************************************************************}
{$IFDEF CHR}
function MODEOpenSQL( SQL: string; RO: Boolean; Nb: Integer = -1; Stack: string = ''; UniDirection: boolean = FALSE; stBases: string = '' ) : TQuery;
begin
  Result := OpenSQL( SQL, RO, Nb, Stack, UniDirection, stBases );
  if not Assigned( Result ) then
  begin
{$IFDEF EAGLCLIENT}
    Result := TQuery.Create( '', nil, -1 );
{$ELSE}
    Result := TQuery.Create( Appli );
{$ENDIF}
  end;
end;
{$ENDIF CHR}
{ TSQLAnaCroise }

function TSQLAnaCroise.AxeToSousPlan(NatureCpt: String): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to MaxAxe do begin
//    if FAxes[i] then
      Inc(Result);
    if NatureCpt = 'A'+IntToStr(i) then
      break;
  end;
end;

constructor TSQLAnaCroise.Create;
begin
  LoadInfo;
end;

destructor TSQLAnaCroise.Destroy;
begin
  inherited;
end;

function TSQLAnaCroise.GetConditionAxe(NatureCpt: String): String;
begin
  if (not GetParamSocSecur('SO_CROISAXE', false) ) or (NatureCpt = GetPremierAxe) then
    Result := 'Y_AXE = "'+NatureCpt+'"'
  else
    Result := 'Y_AXE = "'+GetPremierAxe+'" AND '+GetChampSection(NatureCpt)+' <> ""';
end;

function TSQLAnaCroise.GetChampSection(NatureCpt: String): String;
begin
  if (not GetParamSocSecur('SO_CROISAXE', false) ) or (NatureCpt = GetPremierAxe) then
    Result := 'Y_SECTION'
  else
    Result := 'Y_SOUSPLAN'+IntToStr(AxeToSousPlan(NatureCpt));
end;

function TSQLAnaCroise.GetPremierAxe: String;
begin
  Result := 'A' + IntToStr(FPremierAxe);
end;

procedure TSQLAnaCroise.LoadInfo;
var
  i: Integer;
begin
  FPremierAxe := 0;

  for i := 1 to MaxAxe do begin
    FAxes[i] := GetParamSoc('SO_VENTILA' + IntToStr(i));
    if FAxes[i] then begin
      if (FPremierAxe = 0) then FPremierAxe := i;
    end;
  end;
end;

class function TSQLAnaCroise.ChampSection(NatureCpt: String): String;
var
  Ana: TSQLAnaCroise;
begin
  Ana := TSQLAnaCroise.Create;
  try
    Result := Ana.GetChampSection(NatureCpt);
  finally
    Ana.Free;
  end;
end;

class function TSQLAnaCroise.ConditionAxe(NatureCpt: String): String;
var
  Ana: TSQLAnaCroise;
begin
  Ana := TSQLAnaCroise.Create;
  try
    Result := Ana.GetConditionAxe(NatureCpt);
  finally
    Ana.Free;
  end;
end;

class procedure TSQLAnaCroise.TraduireRequete( NatureCpt : String ; var vStReq : String ) ;
var Ana    : TSQLAnaCroise;
    lStChp : String ;
begin
  Ana := TSQLAnaCroise.Create;
  try
    lStChp := Ana.GetChampSection( NatureCpt );
    if lStChp <> 'Y_SECTION' then
      vstReq := FindEtReplace( vstReq, 'Y_SECTION', lStChp, True ) ;
  finally
    Ana.Free;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 23/11/2005
Modifié le ... :   /  /
Description .. : teste l'existence du rib ou de l'iban dans la liste des rib du
Suite ........ : tiers ou du tic / tid
Mots clefs ... :
*****************************************************************}
function  ExisteRibSurCpt( vStCpt, vStRib : String ) : Boolean ;
var lStEtab,
    lStGuichet,
    lStNumero,
    lStCle,
    lStDom : String ;
    lStLoc : String ;
begin

  lStLoc := GetParamSocSecur('SO_PAYSLOCALISATION', CodeISOFR) ;

  DecodeRIB( lStEtab, lStGuichet, lStNumero, lStCle, lStDom, vStRIB, lStLoc ) ;

  result := ExisteSQL( 'SELECT R_AUXILIAIRE FROM RIB WHERE R_AUXILIAIRE = "'     + vStCpt
                                         + '" AND ( R_CODEIBAN="'                + vStRib
                                         + '" OR ( R_ETABBQ = "'                 + lStEtab
                                                    + '" AND R_GUICHET = "'      + lStGuichet
                                                    + '" AND R_NUMEROCOMPTE = "' + lStNumero
                                                    + '" AND R_CLERIB = "'       + lStCle
                                                    + '" ) ) ' ) ;
end ;


// ================================================================================
// == Fonction OpenSelect
// ==   Retourne un TQuery sur le dossier cible ( local par défaut)
// ================================================================================
Function  OpenSelect( vSQL : String ; vDossier : String = '' ; vRO : Boolean = True ; vStack : String = '' ) : TQuery ;
begin

 if //EstMultiSoc and                                                  // Multisociété actif ?
    ( Trim( vDossier ) <> '' )  and ( vDossier <> V_PGI.SchemaName ) // Dossier cible non local
   // Open sur le / les dossiers
   then result := OpenSQL( vSQL, vRO , -1, vStack, False, vDossier )
   // Open classique
   else result := OpenSQL( vSQL, vRO,  -1, vStack ) ;

end ;

// ================================================================================
// == Fonction GetDossier
// ==   Retourne le nom Décla correspondant au nom physique de la base
// ================================================================================
function  GetDossier    ( vSchemaName : String ; CodeRegroupement : String = '' ) : string ;
var lQMD     : TQuery ;
    lTSData  : TStringList ;
    lStVal   : String ;
    lStBase  : String ;
    lStCode  : String ;
begin

  result := '' ;

  if CodeRegroupement = '' then
    CodeRegroupement := MS_CODEREGROUPEMENT ;

  // récupération paramétrage du regroupement
  lStVal := '' ;
  lQMD := OpenSQL( 'SELECT * FROM YMULTIDOSSIER WHERE YMD_CODE = "' + CodeRegroupement + '"', True, -1, 'GETDOSSIER') ;
  if not lQMD.Eof then
    lStVal := lQMD.FindField('YMD_DETAILS').AsString ;
  Ferme( lQMD ) ;
  if lStVal = '' then Exit ;

  // Récupération 1ère ligne
  lTSData      := TStringList.Create ;
  lTSData.Text := lStVal ;
  lStVal       := lTSData.Strings[0] ;

  // Parcours des bases du regroupement
  while lStVal<>'' do
    begin
    lStBase    := ReadTokenSt( lStVal ) ;
    lStCode    := ReadTokenPipe( lStBase , '|' ) ;
    if lStBase = vSchemaName then
      begin
      result := lStCode ;
      break ;
      end ;
    end ;

  FreeAndNil( lTSData ) ;

end;

// ================================================================================
// == Fonction GetSchemaName
// ==   Retourne le nom de la base en se basant sur le nom Décla
// ================================================================================
function  GetSchemaName ( vDossier    : String ; CodeRegroupement : String = '' ) : string ;
var lQMD     : TQuery ;
    lTSData  : TStringList ;
    lStVal   : String ;
    lStBase  : String ;
    lStCode  : String ;
begin

  result := '' ;

  if CodeRegroupement = '' then
    CodeRegroupement := MS_CODEREGROUPEMENT ;

  // récupération paramétrage du regroupement
  lStVal := '' ;
  lQMD := OpenSQL( 'SELECT * FROM YMULTIDOSSIER WHERE YMD_CODE = "' + CodeRegroupement + '"', True, -1, 'GETDOSSIER') ;
  if not lQMD.Eof then
    lStVal := lQMD.FindField('YMD_DETAILS').AsString ;
  Ferme( lQMD ) ;
  if lStVal = '' then Exit ;

  // Récupération 1ère ligne
  lTSData      := TStringList.Create ;
  lTSData.Text := lStVal ;
  lStVal       := lTSData.Strings[0] ;

  // Parcours des bases du regroupement
  while lStVal<>'' do
    begin
    lStBase    := ReadTokenSt( lStVal ) ;
    lStCode    := ReadTokenPipe( lStBase , '|' ) ;
    if lStCode = vDossier then
      begin
      result := lStBase ;
      break ;
      end ;
    end ;

  FreeAndNil( lTSData ) ;

end;

{Fonction d'insertion en Multi sociétés des filles de vTob}
{---------------------------------------------------------------------------------------}
function InsertOrUpdateAllNivelsMS(vTob : TOB; vDossier : string; InsertOk : Boolean) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  Result := True;
  BeginTrans;
  try
  
    for n := 0 to vTob.Detail.Count - 1 do begin
      F := vTob.Detail[n];
      {Si la tob a des filles, on rappelle la fonction de manière récursive}
      if (F.Detail.Count > 0) and EstMultiSoc and (vDossier <> V_PGI.SchemaName) and (vDossier <> '') then
        InsertOrUpdateAllNivelsMS(F, vDossier, InsertOk);
      {Si la tob est réelle, on met à jour la table}
      if (F.NumTable > 0) then begin
        if InsertOk then Result := InsertTobMs(F, vDossier)
                    else Result := UpdateTobMs(F, vDossier);
        if not Result then begin
          PgiError('Impossible de mettre à jour la table : ' + F.NomTable);
          RollBack;
          Exit;
        end;
      end;
    end;
    CommitTrans;
  except
    on E : Exception do begin
      RollBack;
      Result := False;
    end;
  end;
end;

// ================================================================================
// == Fonction InsertTobMS
// ==   Enregistre un Tob dans le Dossier spécifié, en local par défaut
// ==   Retourne True en cas de Réussite, False sinon.
// ================================================================================

Function  InsertTobMS( vTob : TOB ; vDossier : String = '' ) : Boolean ;
var lStSQL   : String ;
    lStTable : String ;
begin

  result := False ;

  lStTable := vTob.NomTable ;
  if TableToNum(lStTable) <= 0 then Exit ;

  Try

    // Cas classique
    if not EstMultiSoc or ( vDossier = '' ) or ( vDossier = V_PGI.SchemaName ) then
      begin
      vTob.InsertDB( nil ) ;
      result := True ;
      end
    else
    // Cas spécifique au dossier !
      begin

      // Enregistrement de la Tob
      lStSQL := vTob.MakeInsertSQL ;
      lStSQL := FindEtReplace( lStSql, lStTable, GetTableDossier( vDossier, lStTable ), False ) ;
      Result := ExecuteSQL( lStSQL ) = 1 ;

      end ;

  finally

  end ;

end ;

// ================================================================================
// == Fonction UpdateTobMS
// ==   MAJ d'une Tob dans le Dossier spécifié, en local par défaut
// ==   Retourne True en cas de Réussite, False sinon.
// ================================================================================

Function  UpdateTobMS( vTob : TOB ; vDossier : String = '' ) : Boolean ;
var lStSQL   : String ;
    lStTable : String ;
begin

  result := False ;

  lStTable := vTob.NomTable ;
  if TableToNum(lStTable) <= 0 then Exit ;

  Try

    // Cas classique
    if not EstMultiSoc or ( vDossier = '' ) or ( vDossier = V_PGI.SchemaName ) then
      begin
      vTob.UpdateDB ;
      result := True ;
      end
    else
    // Cas spécifique au dossier !
      begin

      // Enregistrement de la Tob
      lStSQL := vTob.MakeUpdateSQL ;  
      lStSQL := FindEtReplace( lStSql, lStTable, GetTableDossier( vDossier, lStTable ), False ) ;
      lStSQL := lStSQL + ' WHERE ' + vTob.Cle1 ;
      result := ExecuteSQL( lStSQL ) = 1 ;

      end ;

  finally

  end ;

end ;

// ================================================================================
// == Fonction UpdateTobMS
// ==   Suppression d'une Tob dans le Dossier spécifié, en local par défaut
// ==   Retourne True en cas de Réussite, False sinon.
// ================================================================================

Function  DeleteTobMS( vTob : TOB ; vDossier : String = '' ) : Boolean ;
var lStSQL   : String ;
    lStTable : String ;
begin

  result := False ;

  lStTable := vTob.NomTable ;
  if TableToNum(lStTable) <= 0 then Exit ;

  Try

    // Cas classique
    if not EstMultiSoc or ( vDossier = '' ) or ( vDossier = V_PGI.SchemaName ) then
      begin
      vTob.DeleteDB ;
      result := True ;
      end
    else
    // Cas spécifique au dossier !
      begin

      // Effacement de la Tob
      lStSQL := vTob.MakeUpdateSQL ;
      lStSQL := Copy( lStSql, pos('WHERE', lStSQL), length( lStSQL ) ) ;
      lStSQL := 'DELETE FROM ' + GetTableDossier( vDossier, lStTable ) + ' ' + lStSQL ;
      result := ExecuteSQL( lStSQL ) = 1 ;

      end ;

  finally

  end ;

end ;


// ================================================================================
// == Fonction RecupInfosSocietes
// ==   Retourne une TOB contenant autant de filles que de société du regroupement
// ==   Chaque fille possède comme champs supplémentaires, la liste des
// ==    champs demandés en paramètres
// ================================================================================
Function  RecupInfosSocietes( vStListeParamSoc : string ; vStCodeRegroupement : string = '' ) : TOB ;
var lStBases   : string ;
    lStDossier : string ;
    lStListePS : string ;
    lTobSoc    : TOB ;
    lStParam   : string ;
    lQParam    : TQuery ;
    lStData    : string ;
    lStDesign  : string ;
    lStType    : Char ;
begin

  result := TOB.Create('$Result', nil, -1 ) ;

  lStBases := GetBasesMS ( vStCodeRegroupement ) ;
  if lStBases = '' then Exit ;

  while (lStBases <> '') do
    begin
    lStDossier := ReadTokenSt( lStBases ) ;
    lStListePS := vStListeParamSoc ;
    lTobSoc    := TOB.Create('$INFOSSOC', result, -1 ) ;
    lTobSoc.AddChampSupValeur('SCHEMANAME', lStDossier ) ;

    while (lStListePS<>'') do
      begin
      lStParam := ReadTokenSt( lStListePS ) ;
      lQParam  := OpenSelect( 'SELECT SOC_DATA, SOC_DESIGN FROM PARAMSOC WHERE SOC_NOM LIKE "' + lStParam + '%"'
                             , lStDossier ) ;
      // Recup valeur
      if not lQParam.Eof then
        begin
        lStDesign := lQParam.FindField('SOC_DESIGN').AsString ;
        if length( lStDesign ) > 1 then
          begin
          lStType   := lStDesign[1] ;
          lStData   := lQParam.FindField('SOC_DATA').AsString ;
          lTobSoc.AddChampSupValeur( lStParam, ParamSocDataToVariant( lStParam, lStData, lStType ) ) ;
          end ;
        end ;

      Ferme( lQParam ) ;

      end ;
    end ;

end ;

Function  SetParamSocDossier  ( vNomParam : String ; vValeur : Variant  ; vDossier : String = '' ) : Boolean ;
begin
  if (vDossier = '') or (vDossier = V_PGI.SchemaName) then
    result := SetParamSoc( vNomParam, vValeur )
  else
    result := ExecuteSQL ( 'UPDATE ' + GetTableDossier( vDossier, 'PARAMSOC')
                           + ' SET SOC_DATA = "' + VariantToParamSocData( vValeur )
                           + '" WHERE SOC_NOM = "' + vNomParam + '"' ) > 0 ;
end ;

// ME  09/02/2006 Fiche 10327
// Centralisation de la fontion pour importCom,cpgeneraux_tom etc...
Function ModePaiementParDefaut (var ModRegle: string) : string;
var
ModPaie           : string;
lQMR              : TQuery;
begin
        ModRegle := GetParamSocSecur('SO_GCMODEREGLEDEFAUT', '');
        ModPaie  := 'DIV' ;
        if ModRegle <> '' then
        begin
          lQMR := OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+ModRegle+'" ',TRUE) ;
          if Not lQMR.Eof Then
            ModPaie := lQMR.FindField('MR_MP1').AsString ;
          Ferme(lQMR) ;
        end ;
        Result := ModPaie;
end;

function isMssql: boolean; //js1 04052006
begin
  Result := V_PGI.Driver in [dbMssql, dbMssql2005]
end;

function isOracle: boolean;
begin
  Result := V_PGI.Driver in [dbORACLE7, dbORACLE8, dbORACLE9, dbORACLE10]
end;

function isDB2: boolean;
begin
  Result := V_PGI.Driver in [dbDB2]
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 13/09/2001
Modifié le ... : 09/01/2002 (JLS/TS)
Modifié le ... : 04/02/2003 (TS)
Description .. : Dans une chaine de type S = 'ACTION=CREATION;PARAMS=ADONF;MAZEL=TOF;'
Description .. : GetArgumentValue('MAZEL', S) retourne 'TOF'
Mots clefs ... : ARGUMENTS
*****************************************************************}
Function GetArgumentValue(Argument: string; Const MyArg : String; Const WithUpperCase: Boolean = True; const Separator: String = ';'): String;
begin
  Result := Tools.GetArgumentValue(Argument, MyArg, WithUpperCase, Separator);
end;

function AGLGetArgumentValue(Parms : array of variant; nb : integer): Variant;
var
  Parms2: Boolean;
  Parms3: String;
begin
  if VarIsEmpty(Parms[2]) then
    Parms2 := True
  else
    Parms2 := VarAsType(Parms[2], varBoolean);

  if VarIsEmpty(Parms[3]) then
    Parms3 := ';'
  else
    Parms3 := VarToStr(Parms[3]);

  Result := GetArgumentValue(Parms[0], Parms[1], Parms2, Parms3);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 14/12/2001
Modifié le ... :   /  /
Description .. : Renvoie directement une chaine.
Suite ........ : test si l'argument existe
Mots clefs ... :
*****************************************************************}
Function GetArgumentString(Argument: string; Const MyArg : String; WithUpperCase: Boolean = True; const Separator: String = ';'):String;
begin
  Result := Tools.GetArgumentString(Argument, MyArg, WithUpperCase, Separator);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 14/12/2001
Modifié le ... :   /  /
Description .. : Renvoie directement un entier.
Suite ........ : test si l'argument existe
Mots clefs ... :
*****************************************************************}
Function GetArgumentInteger(Argument: string; Const MyArg : String; const Separator: String = ';'):Integer;
begin
	if Pos(MyArg, Argument) > 0 then
		Result := Valeuri(GetArgumentValue(Argument, MyArg, True, Separator))
  else
   	Result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 14/12/2001
Modifié le ... :   /  /
Description .. : Renvoie directement un Double
Suite ........ : Test si l'argument n'existe pas
Mots clefs ... :
*****************************************************************}
Function GetArgumentDouble(Argument: string; Const MyArg : String; const Separator: String = ';'):Double;
begin
	if Pos(MyArg, Argument) > 0 then
		Result := Valeur(GetArgumentValue(Argument, MyArg, True, Separator))
  else
   	Result := 0.0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 14/12/2001
Modifié le ... :   /  /
Description .. : Renvoie directement un Double
Suite ........ : Test si l'argument n'existe pas
Mots clefs ... :
*****************************************************************}
function GetArgumentDateTime(Argument: string; Const MyArg : String; const Separator: String = ';'):tDateTime;
begin
	if Pos(MyArg, Argument) > 0 then
		Result := StrToDateTime(GetArgumentValue(Argument, MyArg, True, Separator))
  else
   	Result := iDate1900;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/02/2003
Modifié le ... :   /  /
Description .. : Renvoie directement un Booléen
Suite ........ : Test si l'argument n'existe pas
Mots clefs ... :
*****************************************************************}
function GetArgumentBoolean(Argument: string; Const MyArg : String; const Separator: String = ';'): Boolean;
begin
	if Pos(MyArg, Argument) > 0 then
		Result := StrToBool_(GetArgumentValue(Argument, MyArg, True, Separator))
  else
   	Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/06/2004
Modifié le ... :   /  /
Description .. : Renvoie directement unetob
Mots clefs ... :depuis une chaîne d'arguments
*****************************************************************}
function GetArgumentTob(Argument: string; Const MyArg : String; const Separator: String = ';'): Tob;
begin
	if Pos(MyArg, Argument) > 0 then
		Result := Tob(GetArgumentInteger(Argument, MyArg, Separator))
  else
   	Result := nil;
end;

function SetArgumentTob(const aTobName: String; const aTob: Tob; const aSepar: String = '='): String;
begin
  if (aTobName <> '') and Assigned(aTob) then
    Result := aTobName + aSepar + IntToStr(LongInt(aTob))
  else
    Result := '';
end;

{$IFDEF CHR}
{ GC/GRC : JTR / - eQualité 13254}
// Retourne la valeur d'un paramètre de la ligne de commande
function RetParamCmdLine( sNomVal: string ): string;
var iPos: Integer;
    sDelimiter: string;
begin
  sNomVal := '/' + AnsiUpperCase(sNomVal) + '=';
  iPos := Pos( sNomVal, AnsiUpperCase(CmdLine) );
  if( iPos > 0 ) then
  begin
    Result := TrimLeft( Copy( CmdLine, iPos + Length(sNomVal), Length(CmdLine) ) );
    if( Copy(Result, 1, 1) = '"' ) then
      sDelimiter := '"'
    else
      sDelimiter := ' ';
    iPos := Pos( sDelimiter, Result );
    if( iPos > 0 ) then
      Result := Trim( Copy(Result, 1, iPos) );
  end else
    Result := '';
end;
{ GC/GRC : JTR Fin / - eQualité 13254}
{$ENDIF CHR}
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction qui retourne le montant avec 2 decimales sans 
Suite ........ : virgule :
Suite ........ : 125 ==> 12500
Suite ........ : 125,3 ==> 12530
Suite ........ : 125,33 ==> 12533
Mots clefs ... :
*****************************************************************}
function GetValueMontant(Montant : double) : string;
var
	i,j,k	: integer;
	temp 	: string;
   virgule : boolean;
begin
	k := 0;
	j := 1;
	virgule := false;
   temp := FloatToStr(Montant);
   setlength(result,Length(temp)+2);
   for i := 1 to Length(temp) do
   begin
       if virgule then
       begin
       	Inc(k);
         if k > 2 then exit;
         result[j] := Temp[i];
         Inc(j);
       end
       else
       begin
       	if (Temp[i] = ',') or (Temp[i] = '.') then
         	virgule := true
         else
         begin
	       	result[j] := Temp[i];
   	      Inc(j);
         end;
       end;
   end;
  	if not(virgule) or (k < 2) then
   begin
   	for i := (k+1) to 2 do
      begin
      	result[j] := '0';
         Inc(j);
      end;
   end;
   setlength(result,j-1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 19/07/2006
Modifié le ... :
Description .. : Renvoi la 1ère tob fille qui contient tout ou partie de la valeur passée
                 AuMoins1 = Au moins une des valeurs est présente dans une des filles
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
function FindInTobWithLike(LaTobMere : TOB; AuMoins1 : boolean; Champs : Array of string; Valeurs : Array of variant) : TOB;
var Cpt, Cpt1 : integer;
    TobTmp : TOB;
    Nbre, CompteTous : integer;
begin
  Result := nil;
  if (not assigned(LaTobMere)) then exit;
  Nbre := length(Champs);
  if Nbre = 0 then exit;
  for Cpt := 0 to LaTobMere.detail.count -1 do
  begin
    CompteTous := 0;
    TobTmp := LaTobMere.detail[Cpt];
    { On traite les champs de chacune des filles }
    for Cpt1 := 0 to Nbre -1 do
    begin
      if (TobTmp.FieldExists(Champs[Cpt1])) and (pos(Valeurs[Cpt1], TobTmp.GetValue(Champs[Cpt1])) > 0) then
      begin
        if AuMoins1 then
        begin
          Result := TobTmp;
          break;
        end else
          CompteTous := CompteTous + 1;
      end;
    end;
    { Si Aumoin1 et ok, on quitte }
    if (Aumoins1) and (assigned(Result)) then
      break
    { Sinon, si toutes les valeurs sont identiques, on quitte }
    else if (not Aumoins1) and (Comptetous = Nbre) then
    begin
      Result := TobTmp;
      break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. : Retourne une chaine contenant des paramètres a passer à
Suite ........ : Isoflex (User, Nom société)
Mots clefs ... : ISOFLEX PARAMETRE USER NOM SOCIETE
*****************************************************************}
function GetParamIsoflex : string;
begin
  Result := V_PGI.User + ',' + V_PGI.NomSociete
end;

function RepertoireExiste (LeRepertoire :string; bCreation : boolean = false) : boolean;
begin
  Result := DirectoryExists(LeRepertoire);
  if not Result and bCreation then
  begin
    Result := ForceDirectories (LeRepertoire);
  end;
end;

// BDU - 03/05/07 - Suppression chemin en dur
function AbstractionRepertoire(Genre: TAbstractionRepertoire; Complement: String; Produit: Boolean): String;
begin
  // Répertoire de base
{$IFNDEF EAGLSERVER}
  case Genre of
    rpUserData : Result := TCbpPath.GetCegidUserLocalAppData + '\';
    rpUserTemp : Result := TCbpPath.GetCegidUserTempPath + '\';
    rpCommonData : Result := TCbpPath.GetCommonAppdata + '\';
    rpCegidData : Result := TCbpPath.GetCegidData + '\';
    rpBob:result :=  TCbpPath.GetCegidDistriBob   + '\';
  end;
  // Ajout du nom du produit
  if Produit then
  begin
  {$IFDEF GPAO}
    Result := Result + Copy(NomHalley, 1, 5) + '\';
  {$ELSE GPAO}
    if (ctxAffaire in V_PGI.PGIContexte) or (CtxScot in V_PGI.PGICOntexte) then
      Result := Result + Copy(NomHalley, 2, 4) + '\'
    else
      Result := Result + Copy(NomHalley, 1, 4) + '\';
  {$ENDIF GPAO}
  end;
  // Ajout du complément
  if Complement <> '' then
    Result := Result + Complement + '\';

  RepertoireExiste (Result, true);
{$ENDIF !EAGLSERVER}
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 29/06/2007
Modifié le ... :   /  /
Description .. : Retourne l'équivalent d'un SELECT * avec le nom des champs
Suite ........ : ExlcureChamps = Nom des des champs à ne pas prendre en comptes
Mots clefs ... : 
*****************************************************************}
function GetSelectAll(PrefixesTables : string; SansBlob : boolean=True; ExclureChamps : string='') : string;
var Prefixe, NomChamp : string;
    NumTable, NumChamp : integer;
    IsBlob : boolean;
    Mcd : IMCDServiceCOM;
  	Table     : ITableCOM ;
    FieldList : IEnumerator ;
    Field     : IFieldCOM ;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  //
  Result := '';
  if PrefixesTables = '' then exit;
  while PrefixesTables <> '' do
  begin
    Prefixe := ReadTokenSt(PrefixesTables);
    Table := Mcd.GetTable(Mcd.PrefixeToTable(Prefixe));
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      Field := FieldList.Current as IFieldCOM ;
      NomChamp := Field.Name;
      IsBlob := (pos(Field.Tipe, 'BLOB;DATA') > 0);
      if (NomChamp <> '') and (pos(NomChamp, ExclureChamps) = 0) and (((SansBlob) and  (not IsBlob)) or (not SansBlob)) then
        Result := Result + ', ' + NomChamp;
    end;
  end;
  if Result <> '' then
    Result := copy(Result, 3, length(Result));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Joël SICH
Créé le ...... : 12/01/2007
Modifié le ... :   /  /
Description .. : Renvoie la version de base
Mots clefs ... :
*****************************************************************}
//function TOF_MAJSOCPARLOT.VersionBase(libdoss : String): Variant;
{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}
function VersionBase(libdoss : String): Variant; //js1 120107 on externalise la fonction
var Q: TQuery;
    OldDriver : TDBDriver ;
    OldOdbc,Ok        : Boolean ;
    DBSource: TDatabase;
begin
  Result := 0;
  OldDriver:=V_PGI.Driver ;
  OldOdbc:=V_PGI.ODBC ;
  Application.ProcessMessages ;
  DBSource := TDatabase.Create(nil);
  ok := ConnecteDB(libdoss,DBSource,libdoss);
  if ok then
  begin
    Q := TQuery.Create(nil) ;
    try
      Q.DataBaseName := DBSource.DataBaseName ;
      Q.SessionName  := DBSource.SessionName ;
      Q.SQL.Clear ;
      Q.SQL.Add('SELECT SO_VERSIONBASE FROM SOCIETE') ;
      ChangeSQL(Q) ;
      Try
        Q.Open ;
        if Not Q.Eof then Result := Q.Fields[0].Value;
        if VarToStr(Result)='' then Result := 0;
      finally
        Q.Close;
      end;
    except
      ok := False;
    end;
    Q.Free;
  end;
  if not ok then
    Result := -1;

  DBSource.Connected := False;
  DBSource.Free;
  V_PGI.Driver:=OldDriver ;
  V_PGI.ODBC:=OldOdbc ;
end;
{$ENDIF EAGLCLIENT}

function GetApplicationVersion: String;
var
  VerInfoSize, VerValueSize, Dummy: DWord;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
begin
  VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
  {Deux solutions : }
  if VerInfoSize <> 0 then
  {- Les info de version sont inclues }
  begin
    { On alloue de la mémoire pour un pointeur sur les info de version : }
    GetMem(VerInfo, VerInfoSize);
    try
      {On récupère ces informations : }
      GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo);
      VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
      {On traite les informations ainsi récupérées : }
      with VerValue^ do
      begin
        Result := IntTostr(dwFileVersionMS shr 16);
        Result := Result + '.' + IntTostr(dwFileVersionMS and $FFFF);
        Result := Result + '.' + IntTostr(dwFileVersionLS shr 16);
        Result := Result + '.' + IntTostr(dwFileVersionLS and $FFFF);
      end;
    finally
      {On libère la place précédemment allouée : }
      FreeMem(VerInfo, VerInfoSize);
    end
  end
  else
    {- Les infos de version ne sont pas inclues }
    {On déclenche une exception dans le programme : }
    raise EAccessViolation.Create('Les informations de version de sont pas inclues');
end;

function FormatNumVersion(const Mask: String): String;

  function ApplyMask(Digits, Mask: String): String;
  var
    SubMask, Digit: String;
    i: Integer;
  begin
    while Mask <> '' do
    begin
      SubMask := ReadTokenPipe(Mask  , '.');
      Digit   := ReadTokenPipe(Digits, '.');
      { SubMask = X : on récupère numéro }
      if SubMask = wTrue then
        Result := Result + iif(Result <> '', '.', '') + Digit
      { SubMask = [X.X] : on éclate le numéro de 2 digits en 2 numéros }
      else if Pos('[', SubMask) = 1 then
      begin
        SubMask := StringReplace(Copy(SubMask, 2, Pos(']', SubMask) - 2), '|', '.', [rfReplaceAll]);
        Digit := wPadLeft(Digit, wCountToken(SubMask, '.'), '0');
        for i := 1 to Length(SubMask) do
          if SubMask[i] = '.' then
            Insert('.', Digit, i);
        Result := Result + iif(Result <> '', '.', '') + ApplyMask(Digit, SubMask)
      end
      { SubMask <> - : on récupère le caractère du mask }
      else if SubMask <> wFalse then
        Result := Result + iif(Result <> '', '.', '') + SubMask
      { SubMask = - : on ignore le numéro trouvé }
    end
  end;

begin
  Result := ApplyMask(GetApplicationVersion(), Mask);
end;
{$ENDIF !EAGLSERVER}

{ GC_DBR_ERACANADA_DEBUT }
function GetParamSocPaysNumerique : string;
begin
  Result := CodePays3CVersNumerique (GetParamSocSecur ('SO_PAYS', ''));
end;
{ GC_DBR_ERACANADA_FIN }

{ GC_BDU_ERACANADA_DEBUT - Posé par DBR}
function CodePays3CVersNumerique(CodePays: String): String;
var
  Q: TQuery;
const
  // Le contenu des deux variables ci-dessous est statique donc conservé d'un appel à l'autre
  CodeNumerique: String = '';
  Ok: Boolean = False;
  CodePaysTraite : string = '';
begin
  // Si la recherche a déjà été effectuée, on ne la refait pas
  if Ok and (CodeNumerique <> '') and (CodePaysTraite = CodePays) then
    Result := CodeNumerique
  else
  begin
    // Récupération du code ISO numérique à partir du code ISO 3 caractères
    Q := OpenSQL(Format('SELECT YPY_CODEPAYS FROM YYPAYSISO WHERE YPY_CODEISOA3 = "%s"', [CodePays]), True);
    try
      if not Q.EOF then
      begin
        Result := Q.FindField('YPY_CODEPAYS').AsString;
        // Conserve la valeur pour un prochain appel
        CodeNumerique := Result;
        Ok := True;
        CodePaysTraite := CodePays;
      end;
    finally
      Ferme(Q);
    end;
  end;
end;
{ GC_BDU_ERACANADA_FIN - Posé par DBR }

Function FabriqueConditionIn(Zone : string) : string;
Var TmpZone : String;
Begin

  Result := '';

  if (zone = '') then Exit;

  Tmpzone :=  Trim(ReadTokenSt(Zone));

  While (Tmpzone <>'') do
  Begin
    if Result = '' then
      Result := '("' + Tmpzone + '"'
    else
      Result := Result + ',"' + Tmpzone + '"';
    Tmpzone:=Trim(ReadTokenSt(Zone));
  End;

  Result := result + ')';

end;

function GetMyTempPath : string;
  var
  Buffer: array[0..1023] of Char;
begin
  GetTempPath(Sizeof(Buffer)-1,Buffer);
  SetString(Result, Buffer, StrLen(Buffer));
end;

function GetMyComputerName : string;
var
  ComputerName : array[0..MAX_COMPUTERNAME_LENGTH] of Char;
  nSize        : DWord;
begin
  nSize := SizeOf(ComputerName);
  GetComputerName(@ComputerName, nSize);
  Result := ComputerName;
end;

Initialization
  {$IFNDEF EAGLSERVER}
    RegisterAglProc( 'PGIEnvoiMail', FALSE , 6, AGLPGIEnvoiMail);
    {$IFNDEF NOVH}
      RegisterAglProc( 'AglDecodeRefPiece', False, 2, AglDecodeRefPiece);
    {$ENDIF !NOVH}
    RegisterAGLFunc('AGLGetArgumentValue', False, 2, AGLGetArgumentValue);
  {$ENDIF !EAGLSERVER}
end.



