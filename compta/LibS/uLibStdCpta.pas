{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 05/08/2004
Modifié le ... :   /  /
Description .. : - FQ 13995 - CA - 05/08/2004 - mise à jour de
Suite ........ : SO_ETABLISDEFAUT en création de dossier via l'assistant
Mots clefs ... :
*****************************************************************}
unit uLibStdCpta;

interface

uses sysutils,

{$IFDEF EAGLSERVER}
  uLanceProcess, // LanceProcessServer
{$ELSE}
{$ENDIF}

{$IFDEF VER150}
  Variants,
{$ENDIF}

{$IFDEF EAGLCLIENT}
  LicUtil,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}

  Ed_Tools,      // InitMoveProgressForm
  hctrls,
  Controls,      // MrYes

{$IFNDEF EAGLSERVER}
  hstatus,
  HMsgBox,       // PGIINFO
  Htb97,         // ToolBarButton97
  uIUtil,        // FindInsidePanel
  Forms,         // TForm
{$ENDIF}
  hent1,
  Ent1,
  Classes,       // TStrings
  utob,
  ULibExercice ,
{$IFNDEF EAGLSERVER}
  GalOutil,
{$ENDIF}
  uTXML,         // XMLDecodeSt
  uListByUser,   // TListByUser
  paramsoc,uEntCommun;

//{$IFDEF EAGLSERVER}
//{$ELSE}
type
  TTraitementFiltre = class(TObject)
  public
    FStTableSrc : string;  // Nom de la Table Source
    FStTableDes : string;  // Nom de la Table Destination
    FNumStd     : integer; // Numéro de Standard

    constructor Create;
    destructor Destroy; override;
    procedure MAJFiltreVersionXML;
    procedure Execute;

  private
    FListeByUserSrc : TListByUser;
    FListeByUserDes : TListByUser;

    function  RecopieFiltresRef : Boolean; // Traitement n°4
    procedure FusionFiltres;
    procedure SuppressionFiltreTemp;       // Traitement n°6

    // Code du MUL pour la nouvelle Gestion des FILTRES
    procedure ParseParamsFiltreDes(Params: HTStrings);

  published

  end;

//{$ENDIF}

type
  Lstcpte = record
              compte: string;
            end;

type
  TInfoMajAgricole = procedure ( Erreur : integer; Msg : string ) of object;

const
  YFS_EXTENSION = 'XML';
  //YFS_EXTENSION = 'TXT';

procedure LoadStandardCompta (NumStd : integer; TableDos,TableStd : string; WhereDos:string=''; WhereStd : string=''; LgCpteGen : integer=-1; LgCpteAux : integer = -1);
function  EviteDoublon(var ValChamp: variant; tab, champ: string; LgCpte : Integer): Boolean;
procedure LoadStandardMaj(NumStd: integer; TableDos, TableStd, cle, condition: string ; RazFiltres : Boolean = false ; LgCpteGen : integer=-1; LgCpteAux : integer = -1);

procedure SupprimeDonneStd(TableStd: string; numstd: integer);
procedure SupprimeDonneRefStd(TableStd: string);
procedure SuppressionPlanNonreference;
function  SupprimeDossierType(NumPlan: integer) : boolean;
procedure EnregistreParamSocStandard(NumStd: integer);
procedure EnregistreInfoStandard(Numero: integer; Libelle, Abrege: string);
procedure SaveAsStandardCompta(NumStd: integer; TableDos, TableStd: string; AvecInitMove : boolean = FALSE);

// à dégager
// procedure SaveInfoListe(NumStd: integer; LIST: string);

procedure ChargeStandardCompta ( NumPlan : integer; TableDos, TableStd : string ; Cle: array of string );

procedure InitDossierStandardEuro;
procedure BasculeStandardEuro;
function  StandardEnEuro : boolean;
procedure InitTableExerciceStandard;
procedure CreationSectionAttente ( NomBase : string='');
{$IFNDEF EAGLSERVER}
procedure DeverseStandardDansDossier ( NumStandard : integer; NomBase : string; Monnaie : String) ;
{$ENDIF}
procedure InitEtablissement (NomBase : string = '');
procedure InitEtablissementCompl (NomBase : String; Etablissement : String; Libelle : String);
procedure LoadStandardFiltreDB ( NomBase : string; NumStd : integer ; bMaj : boolean );
procedure LoadStandardListeDB ( NomBase : string; NumStd : integer );
function  PresenceAncienFiltreDansSTD( vNumStd : integer  ) : Boolean;
function  MessageAncienFiltreDansSTD( vNumStd : integer; vNumMessage : integer ) : string;

function ChargeToutLeStandard( vNumStd : integer; vBoCtxStd : Boolean; vBoAvecProgressForm : Boolean) : Boolean;
function EnregistreToutLeStandard( vNumStd : integer ; vStLibelle, vStAbrege : string; vBoAvecProgressForm : Boolean) : Boolean;

procedure ChargeAxe( vNumStd : integer );
procedure InitParamTablesLibres ( stNomBase : string );
function  SaveDossierAsStandard (pNumStd : Integer; pstParamStd : string; pstLibelleStd : string) : integer ;

{$IFDEF EAGLCLIENT}
Function InitTobParamProcessServer : TOB ;
{$ENDIF}

function  ActivationTvaSurDossier( vBoParAligneStd : Boolean = False; vNumStd : integer = 0) : boolean ;
procedure InitialisationTVA      ( vBoParAligneStd : Boolean = False; vNumStd : integer = 0);
procedure CreationSectionTva( vStAxe : string = '');
procedure DesactivationAnalytique;
{$IFNDEF EAGLSERVER}
function  DesactivationTvasurDossier : boolean ;
{$ENDIF EAGLSERVER}

// GCO - 25/04/2006 - PB ADER
procedure CreateCodTXX;
{$IFNDEF EAGLSERVER}
function MiseAJourAxeAgricole (stNoDossier : string; PO : TInfoMajAgricole) : integer;
{$ENDIF EAGLSERVER}

// Fonctions de gestion des standards dans YFILESTD
function YFSChargeTableDepuisStandard ( pstTable : string; NumStd : integer; pstBaseName : string) : boolean;
function YFSEnregistreTableVersStandard ( pstTable : string; NumStd : integer) : boolean;
function YFSSupprimeTableStandard ( pstTable : string; NumStd : integer) : boolean;

implementation

uses
{$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
{$ENDIF MODENT1}

{$IFDEF CCSTD}
{$ELSE}
  {$IFNDEF EAGLSERVER}
    SoldeCpt, // MajTotTousComptes
  {$ENDIF}
{$ENDIF}

  ImSaiCoef
  {$IFNDEF EAGLSERVER}
  ,uLibAnalytique // MAJGVentil
  {$ENDIF}
  , cbpPath
  ,uYFileStd;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/03/2005
Modifié le ... :   /  /
Description .. : Procédure d'initialisation des tables libres en création de
Suite ........ : dossier.
Suite ........ : Cette procédure ne doit être appelée uniquement dans le
Suite ........ : cas où aucune table libre n'a été paramétrée dans le
Suite ........ : standard.
Mots clefs ... :
*****************************************************************}
procedure InitParamTablesLibres ( stNomBase : string );
var TTl : TOB;
    i : integer;
    stCode, stLibelle, stAbrege,stLibre : string;
    stDB : string;
begin
  if stNomBase <> '' then stDb := stNomBase+'.dbo.'
  else stDB := '';
  TTl := TOB.Create ('', nil, -1);
  try
    { Chargement des enregistrements de COMMUN }
    TTl.LoadDetailFromSQL ('SELECT * FROM ##DP##.COMMUN WHERE CO_TYPE="NAT"');
    for i :=0 to TTl.Detail.Count - 1 do
    begin
      { Mise à jour de l'enregistrement CHOIXCOD }
      stCode := TTl.Detail[i].GetValue('CO_CODE');
      stLibelle := TraduireMemoire('Table n°')+IntToStr(StrToInt(Copy(stCode,2,2))+1);
      stAbrege := '-';
      stLibre := TTl.Detail[i].GetValue('CO_LIBRE');
      { Insertion dans la base }
      // CA - 10/05/2006 - Ajout stDB pour créer correctement enregistrements dans CHOIXCOD si création depuis bureau
      if not ExisteSQL ('select * from '+stDB+'choixcod where cc_type="NAT" and cc_code="'+stCode+'"') then
         ExecuteSQL ('INSERT INTO '+stDB+'CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) '+
         'VALUES ("NAT","'+stCode+'","'+stLibelle+'","'+stAbrege+'","'+stLibre+'")');
    end;
  finally
    TTl.Free;
  end;
end;

function EviteDoublon(var ValChamp: variant; tab, champ: string; LgCpte : Integer): Boolean;
var
  Val: string;
  OK: Boolean;
begin
  OK := TRUE;
  Val := VarToStr(ValChamp);
  if Length(Val) > LgCpte then
  begin
    Val := copy(ValChamp, 1, LgCpte);
    if not Presence(tab, champ, Val) then
      ValChamp := Val
    else
      OK := FALSE;
  end;
  Result := OK;
end;

Function StdBourreLaDonc ( St : String ; LgCompte : integer ) : string ;
var ll,i : Integer ;
    Bourre  : Char ;
begin
  Bourre:='0';
  Result:=St ; ll:=Length(Result) ;
  If ll<LgCompte then for i:=ll+1 to LgCompte do Result:=Result+Bourre;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Initialisation de la monnaie du dossier. Pour l'instant
Suite ........ : systématiquement en Euros.
Mots clefs ... :
*****************************************************************}
procedure InitMonnaieDossier (NomBase : string; Monnaie : String);
var TobTableDansBaseDossier : Tob;
    RSql,QDev               : TQuery;
    ChCodeSociete           : String;
    i                       : integer;
begin
  //--- Récupère so_societe directement dans la requete car pb dans V_PGI.CodeSociete
  RSql:=OpenSql ('Select SO_SOCIETE from '+NomBase+'.dbo.SOCIETE',True,-1,'',true);
  ChCodeSociete:=RSql.FindField ('SO_SOCIETE').AsString;
  Ferme (RSql);

  //--- Mise à jour du taux par rapport à l'euro
  if (Monnaie='EUR') then
   begin
    ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_TENUEEURO"');
    ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="1" WHERE SOC_NOM="SO_TAUXEURO"');
   end
  else
   begin
    ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_TENUEEURO"');
    ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="6.55957" WHERE SOC_NOM="SO_TAUXEURO"');
   end;

  //--- Mise à jour de l'enregistrement devise Euro
  // RSql := OpenSQL ('SELECT * FROM '+NomBase+'.dbo.DEVISE WHERE D_DEVISE="EUR"',True);
  // Suppression des devises éventuellement présentes dans le dossier
  ExecuteSQL ('DELETE FROM '+NomBase+'.dbo.DEVISE');

  TobTableDansBaseDossier:=Tob.Create ('DEVISE',Nil,-1);
  try
    TobTableDansBaseDossier.LoadDetailDB('DEVISE','','',nil,TRUE);
    for i := 0 to TobTableDansBaseDossier.Detail.Count - 1 do
    begin
      TobTableDansBaseDossier.Detail[i].PutValue ('D_SOCIETE',ChCodeSociete);
      ExecuteSQL (StringReplace (TobTableDansBaseDossier.Detail[i].MakeInsertSql,'INTO DEVISE','INTO '+NomBase+'.dbo.Devise',[RfIgnoreCase]));
    end;
  finally
    TobTableDansBaseDossier.Free;
  end;

  //--- Suppression de la devise FRF
  ExecuteSQL ('DELETE FROM '+NomBase+'.dbo.DEVISE WHERE D_DEVISE="FRF"');
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+Monnaie+'" WHERE SOC_NOM="SO_DEVISEPRINC"');

  QDev:=OpenSQL('SELECT D_DECIMALE FROM '+NomBase+'.dbo.DEVISE WHERE D_DEVISE="'+Monnaie+'"',True,-1,'',true) ;
  if Not QDev.EOF then ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA='+IntToStr (QDev.Fields[0].AsInteger)+' WHERE SOC_NOM="SO_DECVALEUR"');
  Ferme(QDev) ;

  // $$$ JPO 26/04/06 - si l'euro est monnaie in (db0 foireuse), on met à jour pour qu'elle ne le soit plus
  ExecuteSQL ('UPDATE ' + NomBase + '.dbo.DEVISE SET D_MONNAIEIN="-" WHERE D_DEVISE="EUR" AND D_MONNAIEIN="X"');

  //--- Suppression des monnaie 'In'
  ExecuteSQL ('DELETE FROM '+NomBase+'.dbo.DEVISE WHERE D_MONNAIEIN="X" AND D_FONGIBLE="-"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Déversement des paramsoc d'un standard dans un dossier
Mots clefs ... :
*****************************************************************}
procedure DeverseParamSocRef ( NumPlan: integer; NomBase : string; var LgCpteGen : integer; var LgCpteAux : integer );
var ChSql,StVal             : string;
    RSql                    : TQuery;
begin
 LgCpteGen := 6;
 LgCpteAux := 6;

 //--- Longueur des comptes généraux
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' AND PRR_SOCNOM="SO_LGCPTEGEN"';
 RSql:=OpenSQL(ChSql,True,-1,'',true);
 if (RSql.Eof) then
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="6" WHERE SOC_NOM="SO_LGCPTEGEN"')
 else
  begin
   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+RSql.FindField('PRR_SOCDATA').AsString+'" WHERE SOC_NOM="SO_LGCPTEGEN"');
   LgCpteGen:=RSql.FindField('PRR_SOCDATA').AsInteger;
  end;
 ferme (RSql);

 //--- Longueur des comptes auxiliaires
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' AND PRR_SOCNOM="SO_LGCPTEAUX"';
 RSql:=OpenSql (ChSql,True,-1,'',true);
 if (RSql.Eof) then
   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="6" WHERE SOC_NOM="SO_LGCPTEAUX"')
 else
  begin
   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+RSql.FindField('PRR_SOCDATA').AsString+'" WHERE SOC_NOM="SO_LGCPTEAUX"');
   LgCpteAux:=RSql.FindField('PRR_SOCDATA').AsInteger;
  end;
 ferme (RSql);

 //--- Caractere de bourrage des comptes généraux
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' AND PRR_SOCNOM="SO_BOURREGEN"';
 RSql:=OpenSql (ChSql,True,-1,'',true);
 if (RSql.Eof) then
   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="0" WHERE SOC_NOM="SO_BOURREGEN"')
 else
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+RSql.FindField('PRR_SOCDATA').AsString+'" WHERE SOC_NOM="SO_BOURREGEN"');
 ferme (RSql);

 //--- Caractere de bourrage des comptes auxiliaires
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' AND PRR_SOCNOM="SO_BOURREAUX"';
 RSql:=OpenSql (ChSql,True,-1,'',true);
 if (RSql.Eof) then
   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="0" WHERE SOC_NOM="SO_BOURREAUX"')
 else
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+RSql.FindField('PRR_SOCDATA').AsString+'" WHERE SOC_NOM="SO_BOURREAUX"');
 ferme (RSql);

 //--- Paramètres sociétés
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' ORDER BY PRR_SOCNOM';
 RSql:=OpenSql (ChSql,True);
 {$IFNDEF EAGLSERVER}
 InitMove(RSql.RecordCount, 'Chargement des paramètres société');
 {$ENDIF}
 while not (RSql.Eof) do
  begin
   StVal:=RSql.FindField('PRR_SOCDATA').AsString;
   if (RSql.FindField('PRR_COMPTE').AsString = 'G') then
    StVal:=StdBourreLaDonc(stVal, LgCpteGen)
   else
    if (RSql.FindField('PRR_COMPTE').AsString = 'T') then
     StVal:=StdBourreLaDonc(stVal, LgCpteGen);

   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+StVal+'" WHERE SOC_NOM="'+RSql.FindField('PRR_SOCNOM').AsString+'"');
   RSql.Next;
   {$IFNDEF EAGLSERVER}
   MoveCur(False);
   {$ENDIF}
  end;

 Ferme(RSql);
 {$IFNDEF EAGLSERVER}
 FiniMove;
 {$ENDIF}
end;

{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Deversement d"un standard comptable dans un dossier
Suite ........ : quelconque.
Mots clefs ... :
*****************************************************************}
procedure DeverseStandardDansDossier (NumStandard : integer; NomBase : string; Monnaie : String) ;
var LgCpteGen, LgCpteAux    : integer;
begin
  InitEtablissement (NomBase);
  DeverseParamSocRef(NumStandard, NomBase, LgCpteGen, LgCpteAux );
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_TENUEEURO"');
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_I_CPTAPGI"');
  InitMonnaieDossier (NomBase,Monnaie);

  LoadStandardCompta(NumStandard, NomBase+'.dbo.GENERAUX', 'GENERAUXREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.JOURNAL', 'JALREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.GUIDE', 'GUIDEREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.ECRGUI', 'ECRGUIREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.ANAGUI', 'ANAGUIREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.TIERS', 'TIERSREF','','',LgCpteGen,LgCpteAux);

  if ExisteSQL('SELECT * FROM MODEPAIEREF WHERE MPR_NUMPLAN='+IntToStr(NumStandard)) then
   begin
    ExecuteSQL('DELETE FROM '+NomBase+'.dbo.MODEPAIE');
    LoadStandardCompta(NumStandard, NomBase+'.dbo.MODEPAIE', 'MODEPAIEREF','','',LgCpteGen,LgCpteAux);
   end;

  if ExisteSQL('SELECT * FROM MODEREGLREF WHERE MRR_NUMPLAN='+IntToStr(NumStandard)) then
   begin
    ExecuteSQL('DELETE FROM '+NomBase+'.dbo.MODEREGL');
    LoadStandardCompta(NumStandard, NomBase+'.dbo.MODEREGL', 'MODEREGLREF','','',LgCpteGen,LgCpteAux);
   end;

  LoadStandardCompta(NumStandard, NomBase+'.dbo.CORRESP', 'CORRESPREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.RUPTURE', 'RUPTUREREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.NATCPTE', 'NATCPTEREF','','',LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.FILTRES', 'FILTRESREF','','',LgCpteGen,LgCpteAux);
  LoadStandardMaj(NumStandard, NomBase+'.dbo.LISTE', 'LISTEREF', 'LISTE', ' ',False,LgCpteGen,LgCpteAux);
  LoadStandardMaj(NumStandard, NomBase+'.dbo.AXE', 'AXEREF', 'AXE', ' ',False,LgCpteGen,LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.SECTION', 'SECTIONREF','','',LgCpteGen,LgCpteAux);
  CreationSectionAttente (NomBase);

  LoadStandardMaj(NumStandard, NomBase+'.dbo.CHOIXCOD', 'CHOIXCODREF', 'TYPE;CODE', ' ',False,LgCpteGen,LgCpteAux);
  if not ExisteSQL ('SELECT * FROM '+NomBase+'.dbo.CHOIXCOD WHERE CC_TYPE="NAT"') then InitParamTablesLibres ( NomBase );
  LoadStandardCompta(NumStandard, NomBase+'.dbo.VENTIL', 'VENTILREF','','',LgCpteGen,LgCpteAux);

  // GCO - 04/05/2005
  LoadStandardCompta(NumStandard, NomBase+'.dbo.TXCPTTVA', 'TXCPTTVAREF', '', '', LgCpteGen, LgCpteAux);
  LoadStandardCompta(NumStandard, NomBase+'.dbo.REFAUTO',  'REFAUTOREF',  '', '', LgCpteGen, LgCpteAux);
  // FIN GCO

  // CA - 25/06/2007
  //LoadStandardCompta(NumStandard, NomBase+'.dbo.CMASQUECRITERES', 'CMASQUECRITREF', '', '', LgCpteGen, LgCpteAux);
  //LoadStandardCompta(NumStandard, NomBase+'.dbo.CMASQUESAISIE', 'CMASQUESAISIEREF', '', '', LgCpteGen, LgCpteAux);
  { Déversement des masques de saisie }
  YFSChargeTableDepuisStandard ( 'CMASQUECRITERES',NumStandard,Nombase);
  YFSChargeTableDepuisStandard ( 'CMASQUESAISIE',NumStandard,Nombase);
  { Déversement des lois de répartition }
  YFSChargeTableDepuisStandard ( 'CLOIVENTIL',NumStandard,Nombase);
  // + TARD
  // FIN CA

  InstalleLesCoefficientsDegressifs (NomBase);

  { AJOUT CA 18/10/2007 - Synchronisation PARAMSOC et SOCIETE pour chargement
  correct sur le serveur Web Access - FQ 21680 }
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.SOCIETE SET SO_DECVALEUR='+IntToStr(GalOutil.GetParamSocDossier('SO_DECVALEUR',NomBase)));
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.SOCIETE SET SO_DECQTE='+IntToStr(GalOutil.GetParamSocDossier('SO_DECQTE',NomBase)));
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.SOCIETE SET SO_DECPRIX='+IntToStr(GalOutil.GetParamSocDossier('SO_DECPRIX',NomBase)));
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. : Chargement d'un standard filtre
Suite ........ : Code spécifique pour gérer la copie du contenu des blobs
Suite ........ : d'une base à une autre
Mots clefs ... :
*****************************************************************}
procedure LoadStandardFiltreDB ( NomBase : string; NumStd : integer; bMaj : boolean );
var stSQL : string;
begin
  if NomBase <> '' then
    stSQL := 'INSERT INTO '+NomBase+'.dbo.FILTRES '
  else
    stSQL := 'INSERT INTO FILTRES ';

  StSql := StSql + '(FI_TABLE, FI_LIBELLE, FI_CRITERES, FI_DATECREATION, ' +
                   'FI_DATEMODIF, FI_CREATEUR, FI_UTILISATEUR)';

  stSQL := stSQL + ' SELECT FIR_TABLE FI_TABLE, FIR_LIBELLE FI_LIBELLE, ' +
                   'FIR_CRITERES FI_CRITERES, ' +
                   '"' + USDateTime(Now) + '" FI_DATECREATION, ' +
                   '"' + USDateTime(Now) + '" FI_DATEMODIF, ' +
                   '"' + V_Pgi.User + '" FI_CREATEUR, ' +
                   '"' + V_Pgi.User + '" FI_UTILISATEUR FROM FILTRESREF ' +
                   'WHERE FIR_NUMPLAN = ' + IntToStr(NumStd);
  if bMaj then
  begin
    if NomBase <> '' then
      stSQL := stSQL + ' AND FIR_TABLE NOT IN (SELECT FI_TABLE FROM ' + NomBase + '.dbo.FILTRES)'
    else
      stSQL := stSQL + ' AND FIR_TABLE NOT IN (SELECT FI_TABLE FROM FILTRES)';
  end;
  ExecuteSQL ( stSQL );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. : Chargement d'un standard liste
Suite ........ : Code spécifique pour gérer la copie du contenu des blobs
Suite ........ : d'une base à une autre
Mots clefs ... :
*****************************************************************}
procedure LoadStandardListeDB ( NomBase : string; NumStd : integer );
var stSQL : string;
begin
  if NomBase <> '' then stSQL := 'INSERT INTO '+NomBase+'.dbo.LISTE '
  else stSQL := 'INSERT INTO LISTE ';
  stSQL := stSQL + '(LI_LISTE,LI_UTILISATEUR,LI_LIBELLE,LI_SOCIETE,LI_NUMOK,LI_TRIOK,LI_LANGUE,LI_DATA,LI_DATECREATION,LI_DATEMODIF,LI_CREATEUR)';
  stSQL := stSQL + ' SELECT LIR_LISTE LI_LISTE,LIR_UTILISATEUR LI_UTILISATEUR, LIR_LIBELLE LI_LIBELLE, LIR_SOCIETE LI_SOCIETE,LIR_NUMOK LI_NUMOK,LIR_TRIOK LI_TRIOK,LIR_LANGUE LI_LANGUE,LIR_DATA LI_DATA, ';
  stSQL := stSQL + '"' + USDateTime(Now) + '" LI_DATECREATION, ' +
                   '"' + USDateTime(Now) + '" LI_DATEMODIF, ' +
                   '"' + V_Pgi.User + '" FI_CREATEUR ';
  stSQL := stSQL + ' FROM LISTEREF ';
  stSQL := stSQL + ' WHERE LIR_NUMPLAN='+IntToStr(NumStd);
  ExecuteSQL ( stSQL );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  PresenceAncienFiltreDansSTD( vNumStd : integer  ) : Boolean;
begin
  Result := ExisteSQL('SELECT FIR_TABLE FROM ##DP##.FILTRESREF WHERE ' +
                      'FIR_NUMPLAN = ' + IntToStr( vNumStd ) + ' AND ' +
                      'FIR_LIBELLE NOT LIKE "#%#" ORDER BY FIR_NUMPLAN, FIR_TABLE');
end;

////////////////////////////////////////////////////////////////////////////////
function MessageAncienFiltreDansSTD( vNumStd : integer; vNumMessage : integer ) : string;
begin
  Result := 'Le dossier type ' + GetColonneSQL('STDCPTA', 'STC_LIBELLE', 'STC_NUMPLAN = ' + IntToStr(vNumStd)) +
            ' contient des filtres dont le format a été modifié.' + #13 + #10;

  case vNumMessage of
    // Message du BUREAU PGI
    0 : Result := Result + ' Afin de pouvoir créer le dossier comptable,';
    // Message Création Dossier COMPTA
    1 : Result := Result + ' Afin de pouvoir créer le dossier,';
    // Message Alignement STANDARD
    2 : Result := Result + ' Afin de pouvoir effectuer l''alignement,';
  else

  end;

  Result := Result + ' enregistrez de nouveau votre dossier type' + #13 + #10 +
                     ' dans la Gestion des standards de la comptabilité (module Dossier type) pour le mettre' + #13 + #10 +
                     ' en conformité.';
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/03/2002
Modifié le ... : 15/12/2004
Description .. : Chargement d'une table de standard comptable
Suite ........ : - Si des éléments sont présents dans la table de destination,
Suite ........ : la table n'est pas modifiée.
Suite ........ : CA - 29/04/2002 - traitement de tous les champs car
Suite ........ : NUMPLAN n'est pas forcément le premier champ de la liste
Suite ........ : GCO - 15/12/2004 -> Mise en ProcessSERVER
Mots clefs ... : STANDARD;COMPTA;CHARGEMENT
*****************************************************************}
procedure LoadStandardCompta (NumStd : integer; TableDos,TableStd : string; WhereDos:string=''; WhereStd : string=''; LgCpteGen : integer=-1; LgCpteAux : integer = -1);

  function _DeviseDossier ( MonnaieStd : string ) : string;
  begin
    //Si la devise est renseignée dans le standard , on met à jour avec la
    // devise du dossier sinon, on laisse à blanc - cas du multidevises -
    if MonnaieStd = '' then Result := ''
    else Result := V_PGI.DevisePivot
  end;

var PrefDos, PrefStd, Suffixe , ChampDos  : String;
    NomBase,Table                         : String;
    TobDossier,TobTiersComplementaire     : Tob;
    RSqlStd                               : TQuery;
    ValChamp                              : Variant;
    Position, i                           : Integer;
    OkMaj                                 : Boolean;
begin
  // Pour éviter de passer en paramétre Nombase + Table
  Position := pos ('.dbo.',TableDos);
  if (Position > 0) then
  begin
    NomBase := copy (TableDos,1,pos ('.',TableDos)-1);
    Table   := copy (TableDos,Position+5,Length (TableDos)-(Position+4));
  end
  else
  begin
    NomBase :='';
    Table   := TableDos;
  end;

  // Cas spécifique des filtres - CA - 28/09/2004
  if Table = 'FILTRES' then
  begin
    LoadStandardFiltreDB ( NomBase, NumStd , False );
    Exit;
  end
  else
    if Table = 'LISTE' then
    begin
      LoadStandardListeDB ( NomBase, NumStd );
      Exit;
    end;
{$IFDEF NOVH}
  if LgCpteGen=-1 then LgCpteGen := GetInfoCpta(fbGene).Lg;
  if LgCpteAux=-1 then LgCpteAux := GetInfoCpta(fbAux).Lg;
{$ELSE}
  if LgCpteGen=-1 then LgCpteGen := VH^.Cpta[fbGene].Lg;
  if LgCpteAux=-1 then LgCpteAux := VH^.Cpta[fbAux].Lg;
{$ENDIF}
  PrefDos := TableToPrefixe(Table);
  PrefStd := TableToPrefixe(TableStd);

  // Chargement des enregistrements du standard pour la table concernée
  RSqlStd := OpenSQL ('SELECT * FROM ' + TableStd + ' WHERE '  +PrefStd + '_NUMPLAN=' + IntToStr(NumStd) + WhereStd, True);
  if not RSqlStd.Eof then  // Si un tel standard existe
  begin
      {$IFNDEF EAGLSERVER}
      {$IFNDEF EAGLCLIENT}
      InitMove(QCount(RSqlStd),'Chargement Table : '+TableDos);
      {$ENDIF}
      {$ENDIF}


    // Chargement des enregistrements déjà présents dans la table dossier
    TobDossier := Tob.Create (Table,Nil,-1);
    TobDossier.LoadDetailFromSQL('SELECT * FROM ' + TableDos + WhereDos);

    if (TobDossier.Detail.Count=0) then
    begin   { La table du dossier est vide }
      while not RSqlStd.Eof do
      begin
        OkMaj:=True;
        //--- Traitement des tiers complementaires
        if (Pos('TIERSREF',TableStd)>0) then
        begin
          TobTiersComplementaire:=Tob.Create ('TIERSCOMPL',Nil,-1);
          TobTiersComplementaire.LoadDetailFromSQL ('SELECT * FROM TIERSCOMPL WHERE YTC_AUXILIAIRE="'+StdBourreLaDonc(RSqlStd.FindField('TRR_AUXILIAIRE').AsString, LgCpteAux)+'"');
          if (TobTiersComplementaire.Detail.Count=0) then
          begin
            TobTiersComplementaire.PutValue ('YTC_AUXILIAIRE',StdBourreLaDonc(RSqlStd.FindField('TRR_AUXILIAIRE').AsString, LgCpteAux));
            TobTiersComplementaire.PutValue ('YTC_TIERS',RSqlStd.FindField('TRR_TIERS').AsString);
            TobTiersComplementaire.PutValue ('YTC_SCHEMAGEN',RSqlStd.FindField('TRR_SCHEMAGEN').AsString);
            TobTiersComplementaire.PutValue ('YTC_ACCELERATEUR',RSqlStd.FindField('TRR_ACCELERATEUR').AsString);
            if NomBase = '' then
              ExecuteSQl (StringReplace (TobTiersComplementaire.MakeInsertSql,'INTO TIERSCOMPL','INTO TIERSCOMPL',[RfIgnoreCase]))
            else ExecuteSQl (StringReplace (TobTiersComplementaire.MakeInsertSql,'INTO TIERSCOMPL','INTO '+NomBase+'.dbo.TIERSCOMPL',[RfIgnoreCase]));
          end;
          TobTiersComplementaire.Free;
        end;
        for i:=0 to RSqlStd.FieldCount - 1 do
        begin
          Suffixe := ExtractSuffixe(RSqlStd.Fields[i].FieldName);
          //--- Gestion des exceptions à la règle ici
          if ((PrefStd='PR') and (Suffixe='COMPTE')) then ChampDos := 'G_GENERAL'
          else if ((PrefStd='PR') and (Suffixe='REPORTDETAIL')) then continue
          else if Suffixe='PREDEFINI' then continue
          else if Suffixe='NUMPLAN' then continue
          else if (UpperCase(Suffixe)='ROWGUID') then continue  // Cas particulier KPMG
          //--- GCO - 01/06/2004 - Gestion des champs Accelerateur
          else if (TableStd = 'TIERSREF') and ((Suffixe = 'SCHEMAGEN') or (Suffixe = 'ACCELERATEUR')) then Continue
          // GCO - 10/11/2004 - FQ 14946
          //else if (TableStd = 'JALREF') and (Suffixe = 'ACCELERATEUR') then Continue
          //--- Fin de la gestion des exceptions
          else ChampDos := PrefDos+'_'+Suffixe;

          //--- Calcul de la valeur du champ
          if ChampDos = 'G_GENERAL' then
          begin
            ValChamp := StdBourreLaDonc(RSqlStd.FindField('GER_GENERAL').AsString,LgCpteGen);
            OkMaj := EviteDoublon (ValChamp,'GENERAUX','G_GENERAL',LgCpteGen);
          end
          else
          if ChampDos = 'T_AUXILIAIRE' then
          begin
            ValChamp := StdBourreLaDonc(RSqlStd.FindField('TRR_AUXILIAIRE').AsString,LgCpteAux);
            OKMAJ := EviteDoublon (ValChamp,'TIERS','T_AUXILIAIRE',LgCpteAux);
          end
          else
          if ChampDos = 'T_TIERS' then
            ValChamp := StdBourreLaDonc(RSqlStd.FindField('TRR_TIERS').AsString,LgCpteAux)
          else
          if ChampDos = 'T_COLLECTIF' then
            ValChamp := StdBourreLaDonc(RSqlStd.FindField('TRR_COLLECTIF').AsString,LgCpteGen)
          else
          if ((NomBase='') and ((ChampDos = 'T_DEVISE') or (ChampDos = 'GU_DEVISE'))) then
            ValChamp := _DeviseDossier (RSqlStd.FindField(PrefStd+'_'+Suffixe).AsString)
          else
            ValChamp := RSqlStd.FindField(PrefStd+'_'+Suffixe).AsVariant;

          try
            TobDossier.PutValue (ChampDos,ValChamp);
          except
            {$IFNDEF EAGLSERVER}
            Pgiinfo(TableStd, Suffixe);
            {$ELSE}
            DDWriteln(TableStd + ' ' + Suffixe ) ;
            {$ENDIF}
          end;
        end;

        if OkMaj then
        begin
          if NomBase = '' then
            ExecuteSQl (StringReplace (TobDossier.MakeInsertSql,'INTO '+Table,'INTO '+Table,[RfIgnoreCase]))
          else ExecuteSQl (StringReplace (TobDossier.MakeInsertSql,'INTO '+Table,'INTO '+NomBase+'.dbo.'+Table,[RfIgnoreCase]));
        end;
        TobDossier.InitValeurs;

        {$IFNDEF EAGLSERVER}
        {$IFNDEF EAGLCLIENT}
        MoveCur(False);
        {$ENDIF}
        {$ENDIF}
        RSqlStd.Next;
      end;
    end;
    TobDossier.Free;

    {$IFNDEF EAGLSERVER}
    {$IFNDEF EAGLCLIENT}
    FiniMove;
    {$ENDIF}
    {$ENDIF}
  end;
  Ferme (RSqlStd);
end;
////////////////////////////////////////////////////////////////////////////////

procedure LoadStandardMaj(NumStd: integer; TableDos, TableStd, cle, condition: string ; RazFiltres : Boolean = false; LgCpteGen : integer=-1; LgCpteAux : integer = -1);
var PrefDos,PrefStd,Suffixe,ChampDos,ChampStd : string;
    XWhere,NomBase,Table,Cle1,Cle2,Cle3       : string;
    RSqlStd                                   : TQuery;
    TobDossier                                : Tob;
    ValChamp                                  : Variant;
    Position,i                                : integer;
begin
  //--- Pour eviter de passer en paramétre Nombase + Table
  Position:=pos ('.dbo.',TableDos);
  if (Position>0) then
   begin
    NomBase:=copy (TableDos,1,pos ('.',TableDos)-1);
    Table:=copy (TableDos,Position+5,Length (TableDos)-(Position+4));
   end
  else
   begin
    NomBase:='';
    Table:=TableDos;
   end;

  { Mise à jour des tables particulières FILTRES et LISTE }
  if Table = 'FILTRES' then
  begin
    if RazFiltres then
      ExecuteSQL ( 'DELETE FROM '+TableDos+' WHERE FI_TABLE IN (SELECT FIR_TABLE FROM FILTRESREF WHERE FIR_NUMPLAN='+IntToStr(NumStd)+')');
    LoadStandardFiltreDB ( NomBase, NumStd, not RazFiltres );
    exit;
  end else if Table = 'LISTE' then
  begin
    ExecuteSQL ( 'DELETE FROM '+TableDos+' WHERE LI_LISTE IN (SELECT LIR_LISTE FROM LISTEREF WHERE LIR_NUMPLAN='+IntToStr(NumStd)+')');
    LoadStandardListeDB ( NomBase, NumStd );
    exit;
  end;

{$IFDEF NOVH}
  if LgCpteGen=-1 then LgCpteGen := GetInfoCpta(fbGene).Lg;
  if LgCpteAux=-1 then LgCpteAux := GetInfoCpta(fbAux).Lg;
{$ELSE}
  if LgCpteGen=-1 then LgCpteGen := VH^.Cpta[fbGene].Lg;
  if LgCpteAux=-1 then LgCpteAux := VH^.Cpta[fbAux].Lg;
{$ENDIF}

  PrefDos := TableToPrefixe(Table);
  PrefStd := TableToPrefixe(TableStd);
  Cle1 := ReadTokenSt(cle);
  Cle2 := ReadTokenSt(cle);
  Cle3 := ReadTokenSt(cle);


  RSqlStd := OpenSQL('SELECT * FROM ' + TableStd + ' WHERE ' +PrefStd + '_NUMPLAN=' + IntToStr(NumStd) + condition, True);
  while not RSqlStd.Eof do // Si un tel standard existe
   begin
    {$IFNDEF EAGLSERVER}
    {$IFDEF EAGLCLIENT}
    InitMove(RSQlStd.RecordCount, 'Chargement Table : ' + TableDos);
    {$ELSE}
    InitMove(QCount(RSqlStd), 'Chargement Table : ' + TableDos);
    {$ENDIF}
    {$ENDIF}
    ChampStd := PrefStd + '_' + Cle1;
    ChampDos := PrefDos + '_' + cle1;
    XWhere := ChampDos + '="' + RSqlStd.FindField(ChampStd).AsVariant + '"';
    if Cle2 <> '' then
     begin
      ChampStd := PrefStd + '_' + Cle2;
      ChampDos := PrefDos + '_' + Cle2;
      XWhere := XWhere+' and '+ChampDos+'="'+RSqlStd.FindField(ChampStd).AsVariant+'"';
     end;
    if Cle3 <> '' then
     begin
      ChampStd := PrefStd + '_' + Cle3;
      ChampDos := PrefDos + '_' + Cle3;
      XWhere := XWhere+' and '+ChampDos+'="'+RSqlStd.FindField(ChampStd).AsVariant+'"';
    end;

    TobDossier:=Tob.Create (Table,Nil,-1);
    TobDossier.LoadDetailFromSQL('SELECT * FROM ' + TableDos + ' Where ' + XWhere);
    if not (TobDossier.Detail.Count=0) then
     begin
       for i := 0 to RSqlStd.FieldCount - 1 do
       begin
        Suffixe := ExtractSuffixe(RSqlStd.Fields[i].FieldName);
        //--- Gestion des exceptions à la règle ici
        if ((PrefStd = 'PR') and (Suffixe = 'COMPTE')) then ChampDos := 'G_GENERAL'
        else if ((PrefStd = 'PR') and (Suffixe = 'REPORTDETAIL')) then continue
        else if Suffixe = 'PREDEFINI' then continue
        else if Suffixe = 'NUMPLAN' then continue
        else if (UpperCase(Suffixe)='ROWGUID') then continue  // Cas particulier KPMG
        else ChampDos := PrefDos + '_' + Suffixe;
        //--- Calcul de la valeur du champ
        if ChampDos = 'G_GENERAL' then ValChamp := StdBourreLaDonc(RSqlStd.FindField('GER_GENERAL').AsString,LgCpteGen)
        else if ChampDos = 'T_AUXILIAIRE' then ValChamp := StdBourreLaDonc(RSqlStd.FindField('TRR_AUXILIAIRE').AsString,LgCpteAux)
        else if ChampDos = 'T_TIERS' then ValChamp := StdBourreLaDonc(RSqlStd.FindField('TRR_TIERS').AsString, LgCpteAux)
        else if ChampDos = 'T_COLLECTIF' then ValChamp := StdBourreLaDonc(RSqlStd.FindField('TRR_COLLECTIF').AsString, LgCpteGen)
        else ValChamp := RSqlStd.FindField(PrefStd + '_' + Suffixe).AsVariant;

        TobDossier.Putvalue (ChampDos,ValChamp)
       end;
      ExecuteSQl (StringReplace (TobDossier.MakeUpdateSql+' where '+XWhere,'Update '+Table,'Update '+NomBase+'.dbo.'+Table,[RfIgnoreCase]))
     end
    else
     begin
       for i := 0 to RSqlStd.FieldCount - 1 do
       begin
        Suffixe := ExtractSuffixe(RSqlStd.Fields[i].FieldName);
        if Suffixe = 'PREDEFINI' then continue
        else if Suffixe = 'NUMPLAN' then continue
        else if (UpperCase(Suffixe)='ROWGUID') then continue;  // Cas particulier KPMG
        ChampDos := PrefDos + '_' + Suffixe;
        ValChamp := RSqlStd.FindField(PrefStd + '_' + Suffixe).AsVariant;
        TobDossier.PutValue (ChampDos,ValChamp)
       end;
      ExecuteSQl (StringReplace (TobDossier.MakeInsertSQL,'INTO '+Table,'INTO '+NomBase+'.dbo.'+Table,[RfIgnoreCase]));;
     end;
    {$IFNDEF EAGLSERVER}
    MoveCur(False);
    {$ENDIF}
    TobDossier.Free;
    RSqlStd.Next;
   end;
  {$IFNDEF EAGLSERVER}
  FiniMove;
  {$ENDIF}
  Ferme(RSqlStd);
end;

procedure SupprimeDonneStd(TableStd: string; numstd: integer);
begin
  ExecuteSQL('DELETE FROM ' + TableStd + ' WHERE ' +
             TableToPrefixe(TableStd) + '_NUMPLAN=' + IntToStr(numstd));
end;

procedure SupprimeDonneRefStd(TableStd: string);
var PrefStd: string;
begin
  PrefStd := TableToPrefixe(TableStd);

  ExecuteSQL('DELETE FROM ' + TableStd + ' WHERE ' +
             PrefStd + '_NUMPLAN < 21 AND ' + PrefStd + '_PREDEFINI = "STD"');

  ExecuteSQL('DELETE FROM ' + TableStd + ' WHERE ' + PrefStd + '_NUMPLAN >= 100');
end;

procedure SuppressionPlanNonreference;
var
  Q1: TQuery;
  Where: string;
begin
  // pour supprimer tous les standards predefinis STD et < 21
  // et tous les standards > 100
  SupprimeDonneRefStd('PARSOCREF');
  SupprimeDonneRefStd('GENERAUXREF');
  SupprimeDonneRefStd('JALREF');
  SupprimeDonneRefStd('GUIDEREF');
  SupprimeDonneRefStd('ECRGUIREF');
  SupprimeDonneRefStd('ANAGUIREF');
  SupprimeDonneRefStd('TIERSREF');
  SupprimeDonneRefStd('CORRESPREF');
  SupprimeDonneRefStd('RUPTUREREF');
  SupprimeDonneRefStd('AXEREF');
  SupprimeDonneRefStd('SECTIONREF');
  SupprimeDonneRefStd('CHOIXCODREF');
  SupprimeDonneRefStd('VENTILREF');
  SupprimeDonneRefStd('NATCPTEREF');
  SupprimeDonneRefStd('FILTRESREF');
  SupprimeDonneRefStd('LISTEREF');
  // CA - 02/04/2002
  SupprimeDonneRefStd('MODEPAIEREF');
  SupprimeDonneRefStd('MODEREGLREF');
  // Fin CA - 02/04/2002

  // GCO - 04/05/2005
  SupprimeDonneRefStd('TXCPTTVAREF');
  SupprimeDonneRefStd('REFAUTOREF');
  // FIN GCO

  // CA - 25/06/2007
  // SupprimeDonneRefStd('CMASQUECRITREF');
  // SupprimeDonneRefStd('CMASQUESAISIEREF');
  // FIN CA

  Where := ' WHERE STC_NUMPLAN < 21 AND STC_PREDEFINI="STD"';
  Q1 := OpenSQL('SELECT * FROM STDCPTA ' + Where, True,-1,'',true);
  if not Q1.Eof then
    ExecuteSQL('Delete from STDCPTA' + Where);
  Ferme(Q1);

  Where := ' WHERE STC_NUMPLAN >= 100';
  Q1 := OpenSQL('SELECT * FROM STDCPTA ' + Where, True,-1,'',true);
  if not Q1.Eof then
    ExecuteSQL('Delete from STDCPTA' + Where);
  Ferme(Q1);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/03/2002
Modifié le ... :   /  /
Description .. : Suppression d'un dossier type
Suite ........ : Paramètres : No du dossier type
Mots clefs ... : DOSSIER TYPE;SUPPRESSION
*****************************************************************}

function SupprimeDossierType(NumPlan: integer) : boolean;
begin
  SupprimeDonneStd('PARSOCREF', NumPlan);
  SupprimeDonneStd('GENERAUXREF', NumPlan);
  SupprimeDonneStd('JALREF', NumPlan);
  SupprimeDonneStd('GUIDEREF', NumPlan);
  SupprimeDonneStd('ECRGUIREF', NumPlan);
  SupprimeDonneStd('ANAGUIREF', NumPlan);
  SupprimeDonneStd('TIERSREF', NumPlan);
  SupprimeDonneStd('CORRESPREF', NumPlan);
  SupprimeDonneStd('RUPTUREREF', NumPlan);
  SupprimeDonneStd('AXEREF', NumPlan);
  SupprimeDonneStd('SECTIONREF', NumPlan);
  SupprimeDonneStd('CHOIXCODREF', NumPlan);
  SupprimeDonneStd('VENTILREF', NumPlan);
  SupprimeDonneStd('NATCPTEREF', NumPlan);
  SupprimeDonneStd('FILTRESREF', NumPlan);
  SupprimeDonneStd('LISTEREF', NumPlan);
// CA - 02/04/2002
  SupprimeDonneStd('MODEPAIEREF', NumPlan);
  SupprimeDonneStd('MODEREGLREF', NumPlan);
// Fin CA - 02/04/2002

  SupprimeDonneStd('TXCPTTVAREF', NumPlan);
  SupprimeDonneStd('REFAUTOREF', NumPlan);

  YFSSupprimeTableStandard ('CMASQUECRITERES', NumPlan);
  YFSSupprimeTableStandard ('CMASQUESAISIE', NumPlan);
  YFSSupprimeTableStandard ('CLOIVENTIL', NumPlan);

  ExecuteSQL('DELETE FROM STDCPTA WHERE STC_NUMPLAN=' + IntToStr(NumPlan));
  Result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/03/2002
Modifié le ... :   /  /
Description .. : Enregistre les paramsoc dans un standard
Mots clefs ... : PARAMSOC;STANDARD;DOSSIER TYPE
*****************************************************************}

procedure EnregistreParamSocStandard(NumStd: integer);
var
  QDos, QStd: TQuery;
  StTypeParam, StChamp: string;
begin
  QStd := OpenSQL('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' + IntToStr(NumStd), False);
  if QStd.Eof then
  begin
    // ATTENTION : Exclusion de la branche coordonnées.
    QDos := OpenSQL('SELECT * FROM PARAMSOC WHERE SOC_DATA<>"" ' +
      'AND (((SOC_TREE LIKE "001;001;%") AND (SOC_TREE NOT LIKE "001;001;001;%")) ' +
      'OR SOC_TREE LIKE "001;002;%")', True,-1,'',true);
    while not QDos.Eof do
    begin
      StChamp := QDos.FindField('SOC_NOM').AsString;
      // Exceptions à ne pas sauvegarder :
// ajout me      if (stChamp = 'SO_ETABLISDEFAUT') or (StChamp='SO_LGCPTEGEN') or
//                  (StChamp='SO_LGCPTEAUX') or (StChamp='SO_BOURREGEN') or
//                  (StChamp='SO_BOURREAUX')then
      if ((stChamp = 'SO_ETABLISDEFAUT') or
                (StChamp = 'SO_BOURREAUX') or (StChamp = 'SO_DATEDERNENTREE')  or (StChamp = 'SO_DATEREVISION') or
                (StChamp = 'SO_IMMOMIGEURO') or (StChamp = 'SO_DATEBASCULE')  or (StChamp = 'SO_DATECLOTUREIMMO') or
                (StChamp = 'SO_DATECLOTUREPER') or (StChamp = 'SO_DATECLOTUREPRO')  or (StChamp = 'SO_EXOCLOIMMO') or
                (StChamp = 'SO_EXOV8') or (StChamp = 'SO_ZLASTDATE') or (StChamp = 'SO_CPLASTSAISIE')) then
      begin
        QDos.Next;
        continue;
      end;
      // Fin exceptions
      QStd.Insert;
      InitNew(QStd);
      QStd.FindField('PRR_NUMPLAN').AsInteger := NumStd;
      StTypeParam := UpperCase(Copy(QDos.FindField('SOC_DESIGN').AsString, 1,
        5));
      QStd.FindField('PRR_SOCNOM').AsString :=
        QDos.FindField('SOC_NOM').AsString;
      if (StTypeParam = 'L;TZG') then
        QStd.FindField('PRR_SOCDATA').AsString :=
          BourreLess(QDos.FindField('SOC_DATA').AsString, fbGene)
      else if (StTypeParam = 'L;TZT') then
        QStd.FindField('PRR_SOCDATA').AsString :=
          BourreLess(QDos.FindField('SOC_DATA').AsString, fbAux)
      else
        QStd.FindField('PRR_SOCDATA').AsString :=
          QDos.FindField('SOC_DATA').AsString;
      if (StTypeParam = 'L;TZG') then
        QStd.FindField('PRR_COMPTE').AsString := 'G'
      else if (StTypeParam = 'L;TZT') then
        QStd.FindField('PRR_COMPTE').AsString := 'T'
      else
        QStd.FindField('PRR_COMPTE').AsString := '-';

      // GCO - 09/09/2004 - FQ 13220
      if NumStd > 20 then
        QStd.FindField('PRR_PREDEFINI').AsString := 'STD'
      else
        QStd.FindField('PRR_PREDEFINI').AsString := 'CEG';

      QStd.Post;
      QDos.Next;
    end;
    Ferme(QDos);
  end;
  Ferme(QStd);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/03/2002
Modifié le ... :   /  /
Description .. : Enregistrement des informations générales du standards
Suite ........ : dans la table STDCPTA
Mots clefs ... : STANDARD;COMPTA;
*****************************************************************}

procedure EnregistreInfoStandard(Numero: integer; Libelle, Abrege: string);
var
  T, Ts: TOB;
begin
  if (Numero > 99) then
    V_PGI.IOError := oeUnknown;
  Libelle := Copy(Libelle, 1, 35);
  Abrege := Copy(Abrege, 1, 17);
  T := TOB.Create('', nil, -1);
  try
    T.LoadDetailDB('STDCPTA', IntToStr(Numero), '', nil, False);
    if T.Detail.Count = 0 then
    begin
      Ts := TOB.Create('STDCPTA', T, -1);
      Ts.PutValue('STC_NUMPLAN', Numero);
      Ts.PutValue('STC_LIBELLE', Libelle);
      Ts.PutValue('STC_ABREGE', Abrege);

      if Numero > 20 then
        Ts.PutValue('STC_PREDEFINI', 'STD')
      else
        Ts.PutValue('STC_PREDEFINI', 'CEG');

      // GCO - 10/05/2005
      Ts.SetDateTime('STC_DATEMODIF', Now);
      // FIN GCO
      T.InsertDB(nil);
    end;
  finally
    T.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/03/2002
Modifié le ... :   /  /
Description .. : Copie d'une table dossier vers une table standard
Mots clefs ... :
*****************************************************************}

procedure SaveAsStandardCompta(NumStd: integer; TableDos, TableStd: string; AvecInitMove : boolean = FALSE);
var
  QDos, QStd: TQuery;
  PrefDos, PrefStd, Suffixe, ChampDos: string;
  ValChamp: variant;
  i: integer;
begin
  PrefDos := TableToPrefixe(TableDos);
  PrefStd := TableToPrefixe(TableStd);
  QStd := OpenSQL('SELECT * FROM ' + TableStd + ' WHERE ' + PrefStd + '_NUMPLAN='
    + IntToStr(NumStd), False);
  if QStd.Eof then
  begin
    // GCO - 02/06/2004
    if TableDos <> 'TIERS' then
      QDos := OpenSQL('SELECT * FROM ' + TableDos, True)
    else
      QDos := OpenSQL('SELECT TIERS.*, YTC_SCHEMAGEN, YTC_ACCELERATEUR FROM TIERS LEFT JOIN TIERSCOMPL ON ' +
              'T_AUXILIAIRE = YTC_AUXILIAIRE ORDER BY T_AUXILIAIRE', True);
    // FIN GCO
    {$IFNDEF EAGLSERVER}
    if AvecInitMove then
      {$IFDEF EAGLCLIENT}
      InitMove(QDos.RecordCount, 'Enregistrement Table ' + TableDos);
      {$ELSE}
      InitMove(QCount(QDos), 'Enregistrement Table ' + TableDos);
      {$ENDIF}
    {$ENDIF}
    while not QDos.Eof do
    begin
      if TableDos = 'CHOIXCOD' then
      begin
        if (QDos.FindField(PrefDos + '_TYPE').AsString <> 'VTY') and
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'RUG') and
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'RUT') and
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'NAT') and
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'RTV') and
           // GCO - 20/04/2006 - Ajout des quanlifiants de quantités
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'YTN') and
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'TX1') and
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'TX2') and
           (QDos.FindField(PrefDos + '_TYPE').AsString <> 'QME') then
        begin
          QDos.Next;
          continue;
        end;
      end;
      QStd.Insert;
      InitNew(QStd);
      QStd.FindField(PrefStd + '_NUMPLAN').AsInteger := NumStd;
      if NumStd < 21 then
        QStd.FindField(PrefStd + '_PREDEFINI').AsString := 'CEG'
      else
        QStd.FindField(PrefStd + '_PREDEFINI').AsString := 'STD';
      for i := 0 to QStd.FieldCount - 1 do
      begin
        Suffixe := ExtractSuffixe(QStd.Fields[i].FieldName);
        // Gestion des exceptions à la règle ici
        if ((PrefStd = 'PR') and (Suffixe = 'COMPTE')) then
          ChampDos := 'G_GENERAL'
        else if ((PrefStd = 'PR') and (Suffixe = 'REPORTDETAIL')) then
          continue
        else if (Suffixe = 'PREDEFINI') then
          continue
        else if (Suffixe = 'NUMPLAN') then
          continue
        else if (UpperCase(Suffixe)='ROWGUID') then continue  // Cas particulier KPMG
        // FIN de la gestion des exceptions
        else
          ChampDos := PrefDos + '_' + Suffixe;
        if ChampDos = 'G_GENERAL' then
          ValChamp := BourreLess(QDos.FindField('G_GENERAL').AsString, fbGene)
        else if ChampDos = 'T_AUXILIAIRE' then
          ValChamp := BourreLess(QDos.FindField('T_AUXILIAIRE').AsString, fbAux)
        else if ChampDos = 'T_COLLECTIF' then
          ValChamp := BourreLess(QDos.FindField('T_COLLECTIF').AsString, fbGene)
        else if ChampDos = 'T_TIERS' then
          ValChamp := BourreLess(QDos.FindField('T_TIERS').AsString, fbAux)
        // GCO - 02/06/2004
        else if (TableDos = 'TIERS') and ((Suffixe = 'SCHEMAGEN') or (Suffixe = 'ACCELERATEUR')) then
           ValChamp := QDos.FindField('YTC_' + Suffixe).AsString
        // FIN GCO
        else ValChamp := QDos.FindField(ChampDos).AsVariant;

        try
          QStd.FindField(PrefStd + '_' + Suffixe).AsVariant := ValChamp;
        except
          {$IFNDEF EAGLSERVER}
          pgiinfo( VarToStr(QDos.FindField(ChampDos).AsVariant), PrefStd + '_' + Suffixe);
          {$ELSE}
           DDWriteln(VarToStr(QDos.FindField(ChampDos).AsVariant) + ' ' + PrefStd + '_' + Suffixe ) ;
          {$ENDIF}
        end;
      end;
      QStd.Post;
      {$IFNDEF EAGLSERVER}
      if AvecInitMove then MoveCur(False);
      {$ENDIF}
      QDos.Next;
    end;
    Ferme(QDos);
    {$IFNDEF EAGLSERVER}
    if AvecInitMove then FiniMove;
    {$ENDIF}
  end;
  Ferme(QStd);
end;

(*
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/03/2002
Modifié le ... :   /  /
Description .. : Enregistrement d'une liste en standard
Mots clefs ... :
*****************************************************************}
procedure SaveInfoListe(NumStd: integer; LIST: string);
var
  Q1, QPlanRef: TQuery;
begin
  Q1 := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="' + LIST + '" '
    + 'AND LI_UTILISATEUR="---"', FALSE);

  if not Q1.Eof then
  begin
    QPlanRef := OpenSQL('SELECT * FROM LISTEREF WHERE LIR_NUMPLAN=' +
      IntToStr(NumStd) + ' AND LIR_LISTE="' + LIST + '" '
      + 'AND LIR_UTILISATEUR="---"', FALSE);

    if QPlanRef.Eof then
    begin
      QPlanRef.Insert;
      InitNew(QPlanRef);
      QPlanRef.FindField('LIR_NUMPLAN').AsInteger := NumStd;
      QPlanRef.FindField('LIR_LISTE').AsString := LIST;
      //            QPlanRef.FindField('LIR_UTILISATEUR').AsString := Q1.findField ('LI_UTILISATEUR').asstring;
      QPlanRef.FindField('LIR_UTILISATEUR').AsString := '---';
      QPlanRef.FindField('LIR_LIBELLE').AsString :=
        Q1.findField('LI_LIBELLE').asstring;
      QPlanRef.FindField('LIR_SOCIETE').AsString :=
        Q1.findField('LI_SOCIETE').asstring;
      QPlanRef.FindField('LIR_NUMOK').AsVariant :=
        Q1.findField('LI_NUMOK').AsVariant;
      QPlanRef.FindField('LIR_TRIOK').AsVariant :=
        Q1.findField('LI_TRIOK').AsVariant;
      QPlanRef.FindField('LIR_LANGUE').AsString :=
        Q1.findField('LI_LANGUE').asstring;
      QPlanRef.FindField('LIR_DATA').AsVariant :=
        Q1.findField('LI_DATA').AsVariant;
      if NumStd < 21 then
        QPlanRef.FindField('LIR_PREDEFINI').Asstring := 'CEG'
      else
        QPlanRef.FindField('LIR_PREDEFINI').Asstring := 'STD';
      QPlanRef.Post;
    end
    else
    begin
      QPlanRef.Edit;
      QPlanRef.FindField('LIR_DATA').Asstring :=
        Q1.findField('LI_DATA').Asstring;
      QPlanRef.Post;
    end;
    Ferme(QPlanRef);
  end;
  Ferme(Q1);
end;*)

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : Analyse des champs spécifiques pour le retraitement le cas
Suite ........ : échéant.
Suite ........ : Exemple : seule la racine des champs de type comptes est 
Suite ........ : enregistrée dans la base. Il faut donc effectuer un bourrage 
Suite ........ : des comptes au préalable.
Mots clefs ... : 
*****************************************************************}
function AnalyseChampStandard ( Champ : string; Valeur : variant ) : variant;
begin
  Result := '';
  if Champ = 'G_GENERAL' then
    Result := BourreLaDonc(Valeur, fbGene)
  else if Champ = 'T_AUXILIAIRE' then
    Result := BourreLaDonc(Valeur, fbAux)
  else if Champ = 'T_TIERS' then
    Result := BourreLaDonc(Valeur, fbAux)
  else if Champ = 'T_COLLECTIF' then
    Result := BourreLaDonc(Valeur,  fbGene)
  else Result := Valeur;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : Nouvelle fonction de chargement d'un standard dans un
Suite ........ : dossier comptable.
Suite ........ : Exemple de Paramètres ( 7 , 'GENERAUX','GENERAUXREF','GENERAL')
Suite ........ : Cette fonction permet aussi la mise à jour d'une table dossier
Suite ........ : D'après C AYEL, appelé par le BUREAU
Mots clefs ... : CHARGEMENT;STANDARD;
*****************************************************************}
procedure ChargeStandardCompta ( NumPlan : integer; TableDos, TableStd : string ; Cle: array of string );
var TDos, TDosInsert, T : TOB;
    QStd : TQuery;
    PrefDos, PrefStd, Suffixe : string;
    i : integer;
    CleDos, CleStd : array of string;
    ValCleStd : array of variant;
begin
  // Initialisation
  SetLength(CleDos , Length (Cle));
  SetLength(CleStd , Length (Cle));
  SetLength(ValCleStd, Length(Cle));
  PrefDos := TableToPrefixe(TableDos);
  PrefStd := TableToPrefixe(TableStd);
  for i:=Low(Cle) to High(Cle) do
  begin
    CleDos[i] := PrefDos+'_'+Cle[i];
    CleStd[i] := PrefStd+'_'+Cle[i];
  end;
  TDos := TOB.Create ('', nil, -1);
  TDosInsert := TOB.Create ('', nil, -1);
  try
    TDos.LoadDetailDB (TableDos,'','',nil,False);
    if TDos <> nil then
    begin
      QStd := OpenSQL('SELECT * FROM ' + TableStd + ' WHERE ' + PrefStd + '_NUMPLAN='
        + IntToStr(NumPlan), True,-1,'',true);
      try
        while not QStd.Eof do
        begin
          // Cet enregistrement est-il déjà présent dans la table dossier
          for i:=Low(Cle) to High(Cle) do
            ValCleStd[i] := AnalyseChampStandard(CleDos[i],QStd.FindField(CleStd[i]).AsString);
          if (TDos.FindFirst(CleDos,ValCleStd,False) = nil ) then
          begin
            // On ajoute l'enregistrement dans la table
            T := TOB.Create (TableDos,TDosInsert,-1);
            for i := 0 to QStd.FieldCount - 1 do
            begin
              Suffixe := ExtractSuffixe(QStd.Fields[i].FieldName);
              // Gestion des exceptions à la règle ici
              if (Suffixe = 'PREDEFINI') or (Suffixe='NUMPLAN') then continue
              else if (UpperCase(Suffixe)='ROWGUID') then continue;  // Cas particulier KPMG
              // Mise à jour du champ en cours de traitement
              T.PutValue ( PrefDos + '_' + Suffixe,AnalyseChampStandard(PrefDos + '_' + Suffixe,QStd.FindField (PrefStd + '_' + Suffixe).AsVariant));
            end;
          end;
          QStd.Next;
        end;
      finally
        Ferme (QStd);
      end;
    end;
    TDosInsert.InsertDB (nil);
  finally
    TDosInsert.Free;
    TDos.Free;
  end;
end;

procedure InitDossierStandardEuro;
var TDevise, T : TOB;
begin
  SetParamSoc('SO_TENUEEURO', True);
  SetParamSoc('SO_TAUXEURO', 1);
  TDevise := TOB.Create('', nil, -1);
  TDevise.LoadDetailDB('DEVISE', '', '', nil, True);
  T := TDevise.FindFirst(['D_DEVISE'], ['EUR'], False);
  if (T = nil) then
  begin
    T := TOB.Create('DEVISE', TDevise, -1);
    T.PutValue('D_DEVISE', 'EUR');
    T.PutValue('D_LIBELLE', 'Euro');
    T.PutValue('D_SYMBOLE', '€');
    T.PutValue('D_FERME', '-');
    T.PutValue('D_DECIMALE', 2);
    T.PutValue('D_QUOTITE', 1);
    T.PutValue('D_SOCIETE', V_PGI.Codesociete);
    T.PutValue('D_MONNAIEIN', '-');
    T.PutValue('D_FONGIBLE', '-');
    T.PutValue('D_PARITEEURO', 1);
  end
  else
  begin
    T.PutValue('D_FERME', '-');
    T.PutValue('D_DECIMALE', 2);
    T.PutValue('D_QUOTITE', 1);
    T.PutValue('D_MONNAIEIN', '-');
    T.PutValue('D_FONGIBLE', '-');
    T.PutValue('D_PARITEEURO', 1);
  end;
  TDevise.InsertOrUpdateDB(True);
  // Suppression des devises MONNAIEIN
  ExecuteSQL ('DELETE FROM DEVISE WHERE D_MONNAIEIN="X" AND D_FONGIBLE="-"');
  SetParamSoc('SO_DEVISEPRINC', 'EUR');
  TDevise.Free;
  AvertirTable('TTDEVISE');
  AvertirTable('TTDEVISETOUTES');
  AvertirTable('TTDEVISEETAT');
  AvertirTable('TTDEVISEOUT');
end;

function StandardEnEuro : boolean;
begin
  Result := not ExisteSQL ('SELECT * FROM PARSOCREF WHERE PRR_SOCDATA="FRF" AND PRR_SOCNOM="SO_DEVISEPRINC"');
end;

procedure BasculeStandardEuro;
var T : TOB;
    i : integer;
begin
  if  not StandardEnEuro then
  begin
    try
      BeginTrans;
      T := TOB.Create ('', nil, - 1);
      try
        T.LoadDetailDB ('DEVISE','','',nil,False);
        for i := 0 to T.Detail.Count - 1 do
        begin
          if T.Detail[i].GetValue('D_MONNAIEIN')='X' then
          begin
            // Mise à jour des guides
            ExecuteSQL ('UPDATE GUIDEREF SET GDR_DEVISE="EUR" WHERE GDR_DEVISE="'+
              T.Detail[i].GetValue('D_DEVISE')+'"');
            // Mise à jour des Tiers
            ExecuteSQL ('UPDATE TIERSREF SET TRR_DEVISE="EUR" WHERE TRR_DEVISE="'+
              T.Detail[i].GetValue('D_DEVISE')+'"');
            // Mise à jour des paramsoc
            ExecuteSQL ('UPDATE PARSOCREF SET PRR_SOCDATA="X" WHERE PRR_SOCNOM="SO_TENUEEURO" AND '+
                'PRR_SOCDATA="'+T.Detail[i].GetValue('D_DEVISE')+'"');
            ExecuteSQL ('UPDATE PARSOCREF SET PRR_SOCDATA="EUR" WHERE PRR_SOCNOM="SO_DEVISEPRINC" AND '+
                'PRR_SOCDATA="'+T.Detail[i].GetValue('D_DEVISE')+'"');
          end;
        end;
      finally
        T.Free;
      end;
      CommitTrans;
    except
      Rollback;
    end;
  end;
end;

procedure InitTableExerciceStandard;

  procedure InsereExerciceStandard ( Code, Libelle, Abrege, EtatCpta : string; DateDebut, DateFin : TDateTime );
  begin
    ExecuteSQL ('INSERT INTO EXERCICE (EX_DATEDEBUT,EX_DATEFIN,EX_ETATCPTA,EX_EXERCICE,EX_LIBELLE,EX_ABREGE) '+
            ' VALUES ("'+USDateTime(DateDebut)+'","'+USDateTime(DateFin)+'","'+EtatCpta+'","'+
                        Code+'","'+Libelle+'","'+Abrege+'")');
  end;

var y,m,d , iAnnee               : Word;
    i,j, iExo                    : integer;
    DateDebut, DateFin           : TDateTime;
    CodeExo,Libelle,EtatCpta     : string;
begin
  ExecuteSQL ('DELETE FROM EXERCICE');
  iExo := 1;
  DecodeDate ( V_PGI.DateEntree, y, m, d);
  for i:=0 to 1 do
  begin
    if i=0 then EtatCpta:='CLO' else EtatCpta:='NON';
    for j:=9 downto 1 do
    begin
      CodeExo := Format('%.03d',[iExo]);
      if (i=0) then iAnnee := y-j
      else iAnnee := 10-j+y;
      DateDebut := EncodeDate( iAnnee,1,1);
      DateFin := EncodeDate( iAnnee,12,31);
      Libelle := TraduireMemoire ('Exercice '+IntToStr(iAnnee));
      InsereExerciceStandard ( CodeExo, Libelle, Libelle, EtatCpta, DateDebut, DateFin);
      Inc (iExo, 1);
    end;
    if i = 0 then
    begin
      CodeExo := Format('%.03d',[iExo]);
      DateDebut := EncodeDate( y,1,1);
      DateFin := EncodeDate( y,12,31);
      Libelle := TraduireMemoire ('Exercice '+IntToStr(y));
      InsereExerciceStandard ( CodeExo, Libelle, Libelle, 'OUV', DateDebut, DateFin);
      Inc(iExo,1);
    end;
  end;
end;

//{$IFNDEF EAGLSERVER}
////////////////////////////////////////////////////////////////////////////////
{ TTraitementFiltre }
////////////////////////////////////////////////////////////////////////////////
constructor TTraitementFiltre.Create;
begin

end;

////////////////////////////////////////////////////////////////////////////////

destructor TTraitementFiltre.Destroy;
begin
  FreeAndNil(FListeByUserDes);
  FreeAndNil(FListeByUserSrc);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/12/2003
Modifié le ... : 09/12/2003
Description .. : Lancement du traitement de l'objet
Suite ........ : Principe de fonctionnement de l'objet
Suite ........ : 1°) Conversion de la table FILTRES ( changement de format des anciens filtres )
Suite ........ : 2°) Recopie des anciens filtres de FILTRESREF dans FILTRES
Suite ........ : 4°) Recopie des nouveaux filtres de FILTRESREF avec modification
Suite ........ :     du FI_LBELLE pour éviter les violations Clé si existe déjà ( Filtres temporaires )
Suite ........ : 5°) Fusion des filtres de FILTRESREF et FILTRES
Suite ........ : 6°) Suppression des filtres temporaires pour le traitement
Mots clefs ... :
*****************************************************************}
procedure TTraitementFiltre.Execute;
begin
  // 1°) Conversion éventuelle de Filtres Anciens Format en XML
  MAJFiltreVersionXML;
  // 2°) Insertion des Filtres de FILTRESREF dans FILTRES avec un nom particulier
  RecopieFiltresREF;
  // 3°)
  FusionFiltres;
  // 4°)
  SuppressionFiltreTemp;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/12/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TTraitementFiltre.RecopieFiltresRef : Boolean;
var lQuery : TQuery;
    lSt : string;
begin
  Result := False;
  try
    try
      ExecuteSQL('DELETE FILTRES WHERE FI_TABLE LIKE "@%"');
      // Ajout GCO - 11/01/2005 - FQ 15236 - On ne recopie que les filtres #---#
      lQuery := OpenSQL('SELECT * FROM ##DP##.FILTRESREF WHERE FIR_NUMPLAN = ' +
                IntToStr( FNumStd) + ' AND FIR_LIBELLE = "#---#"', True,-1,'',true);
      while not lQuery.Eof do
      begin
        lSt := 'INSERT INTO FILTRES (FI_TABLE, FI_LIBELLE, FI_CRITERES) ' +
               '(SELECT "@" + FIR_TABLE, FIR_LIBELLE, FIR_CRITERES FROM ##DP##.FILTRESREF ' +
               'WHERE FIR_NUMPLAN = ' + IntToStr(FNumStd) + ' AND ' +
               'FIR_TABLE = "' + lQuery.FindField('FIR_TABLE').AsString + '" AND ' +
               'FIR_LIBELLE = "#---#")';
        ExecuteSQL(lSt);
        lQuery.Next;
      end;
      Result := True;

    except
      on E: Exception do
    {$IFDEF EAGLSERVER}
    {$ELSE}
        PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : RecopieNouveauxFiltres');
    {$ENDIF}
    end;
  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/05/2004
Modifié le ... :   /  /
Description .. : Fusion des Filtres de  FListeByUserSrc et FListeByUserDes
Mots clefs ... :
*****************************************************************}
procedure TTraitementFiltre.FusionFiltres;
var lQuery          : TQuery;
    lSt             : string;
    lStNomFiltreSrc : string;
    lStNomFiltreDes : string;
begin
  lQuery := nil;
  try
    try
      lSt := 'SELECT * FROM FILTRES WHERE FI_TABLE LIKE "@%" ORDER BY FI_TABLE';

      lQuery := OpenSql( lSt, True,-1,'',true );
      while not lQuery.Eof do
      begin
        lStNomFiltreSrc := lQuery.FindField('FI_TABLE').AsString;
        lStNomFiltreDes := Copy(lStNomFiltreSrc, 2, Length(lStNomFiltreSrc));

        FListeByUserDes := TListByUser.Create;
        FListeByUSerDes.LoadDB( lStNomFiltreDes );

        FListeByUserSrc := TListByUser.Create;
        FListeByUserSrc.LoadDB( lStNomFiltreSrc );

        FListeByUserDes.Fusionne( FListeByUserSrc );
        FListeByUserDes.Save;
        FreeAndNil( FListeByUserDes );
        FreeAndNil( FListeByUserSrc );
        lQuery.Next;
      end;
    except
      on E: Exception do
    {$IFDEF EAGLSERVER}

    {$ELSE}
      PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : FusionFiltres');
    {$ENDIF}
    end;

  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/05/2004
Modifié le ... :   /  /
Description .. : Suppression des Filtres @+NomFiltre qui proviennent de la
Suite ........ : recopie de FILTRESREF
Mots clefs ... :
*****************************************************************}
procedure TTraitementFiltre.SuppressionFiltreTemp;
begin
  try
    ExecuteSQL('DELETE FROM FILTRES WHERE FI_TABLE LIKE "@%"');
  except
    on E: Exception do
  {$IFDEF EAGLSERVER}
  {$ELSE}
    PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : SuppressionFiltreTemp');
  {$ENDIF}
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/12/2003
Modifié le ... :   /  /
Description .. : ParseParams du Filtre Dossier
Mots clefs ... :
*****************************************************************}
procedure TTraitementFiltre.ParseParamsFiltreDes(Params: HTStrings);
var T: TOB;
begin
  FListeByUserDes.AddVersion;
  T := FListeByUserDes.Add;
  //en position 0 de Params se trouve le nom du filtre
  T.PutValue('NAME', XMLDecodeSt(Params[0]));
  T.PutValue('USER', '---');
  Params.Delete(0);
  FListeByUserDes.AffecteTOBFiltreMemoire(T, Params);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/05/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : Conversion des Filtres ancien FORMAT en Filtres XML
*****************************************************************}
procedure TTraitementFiltre.MAJFiltreVersionXML;
var lQuery : TQuery;
    lSt    : string;
begin
  lSt := 'SELECT DISTINCT FI_TABLE FROM FILTRES WHERE FI_LIBELLE NOT LIKE "#%#"';
  lQuery := nil;
  try
    try
      lQuery := OpenSQL( lSt, True,-1,'',true);
      while not lQuery.Eof do
      begin
        FListeByUserDes := TListByUser.Create;
        FListeByUserDes.OnParams := ParseParamsFiltreDes;
        FListeByUserDes.LoadDB(lQuery.FindField('FI_TABLE').AsString);
        FreeAndNil(FListeByUserDes);
        lQuery.Next;
      end;
    except
      on E: Exception do
    {$IFDEF EAGLSERVER}
    {$ELSE}
        PgiError('Erreur de requête SQL : ' + E.Message, 'MAJFiltreVersionXML');
    {$ENDIF}    
    end;
  finally
    Ferme(lQuery);
  end;
end;
//{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/07/2004
Modifié le ... : 07/10/2004
Description .. : Création des sections d'attente dans un dossier vide
Suite ........ : - 07/10/2004 - CA - Mise à jour de TOUS les champs de
Suite ........ : l'enregistrement section  en création
Mots clefs ... :
*****************************************************************}
procedure CreationSectionAttente ( NomBase : string = '' );
var
  stSection: string;
  i: integer;
  stTableAxe, stTableSection : string;
  TSection : TOB;
begin
  if NomBase='' then stTableAxe := 'AXE' else stTableAxe := NomBase+'.dbo.AXE';
  if NomBase='' then stTableSection := 'SECTION' else stTableSection := NomBase+'.dbo.SECTION';
  for i := 1 to 5 do
{$IFDEF NOVH}
    if GetInfoCpta(AxeToFb('A' + IntToStr(i))).Attente <> '' then
{$ELSE}
    if VH^.Cpta[AxeToFb('A' + IntToStr(i))].Attente <> '' then
{$ENDIF}
    begin
{$IFDEF NOVH}
      stSection := BourreLaDonc(GetInfoCpta(AxeToFb('A' + IntToStr(i))).Attente,
        AxeToFb('A' + IntToStr(i)));
{$ELSE}
      stSection := BourreLaDonc( VH^.Cpta[AxeToFb('A' + IntToStr(i))].Attente,
        AxeToFb('A' + IntToStr(i)));
{$ENDIF}
      ExecuteSQL('Update '+stTableAxe+' set X_SECTIONATTENTE="' + stSection +
        '" Where X_AXE="A' + IntToStr(i) + '"');

{$IFDEF NOVH}
      if not ExisteSQL('SELECT * FROM '+stTableSection+' Where S_SECTION="' +
        GetInfoCpta(AxeToFb('A' + IntToStr(i))).Attente + '"') then
{$ELSE}
      if not ExisteSQL('SELECT * FROM '+stTableSection+' Where S_SECTION="' +
        VH^.Cpta[AxeToFb('A' + IntToStr(i))].Attente + '"') then
{$ENDIF}
      begin   { Création de la section d'attente }
        TSection := TOB.Create ('SECTION',nil,-1);
        try
          TSection.PutValue('S_SECTION',stSection);
          TSection.PutValue('S_LIBELLE',TraduireMemoire('Section d''attente'));
          TSection.PutValue('S_ABREGE',Copy(TraduireMemoire('Section d''attente'),1,17));
          TSection.PutValue('S_SENS','M');
          TSection.PutValue('S_SOLDEPROGRESSIF','X');
          TSection.PutValue('S_AXE','A'+IntToStr(i));
          TSection.PutValue('S_CONFIDENTIEL','0');
          // CEGID V9
          TSection.PutValue('S_INVISIBLE','-');
          // ------
          ExecuteSQL (StringReplace (TSection.MakeInsertSql,'INTO SECTION','INTO '+stTableSection,[RfIgnoreCase]));
        finally
          TSection.Free;
        end;
      end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : Initalisation de l'enregistrement établissement en création de
Suite ........ : dossier
Mots clefs ... :
*****************************************************************}
procedure InitEtablissement (NomBase : string = '');
var RSqlEtablissement,RSqlSociete : TQuery;
    TobEtablissement, TobSociete  : Tob;
    ChSql                         : String;
begin
  //--- Si on n'accède pas à une base dossier, traitement normal
  if (NomBase='') then
  begin
    //--- Si il existe déjà un établissement dans la base, on ne crée pas d'établissement
    if ExisteSQL ('SELECT * FROM ETABLISS') then
      SetParamSoc('SO_ETABLISDEFAUT', GetColonneSQL('ETABLISS','ET_ETABLISSEMENT',''))
    else
    begin
      ChSql:='SELECT * FROM SOCIETE';
      RSqlSociete := OpenSQL (ChSql,True,-1,'',true);

      if (not RSqlSociete.Eof) then
      begin
        TOBEtablissement := TOB.Create ('ETABLISS',nil,-1);
        try
          TOBEtablissement.PutValue('ET_ETABLISSEMENT',RSqlSociete.FindField('SO_SOCIETE').AsString);
          TOBEtablissement.PutValue('ET_SOCIETE',RSqlSociete.FindField('SO_SOCIETE').AsString);
          TOBEtablissement.PutValue('ET_LIBELLE',RSqlSociete.FindField('SO_LIBELLE').AsString);
          TOBEtablissement.PutValue('ET_ABREGE',Copy (RSqlSociete.FindField('SO_LIBELLE').AsString,1,17));
          TOBEtablissement.PutValue('ET_ADRESSE1',RSqlSociete.FindField('SO_ADRESSE1').AsString);
          TOBEtablissement.PutValue('ET_ADRESSE2',RSqlSociete.FindField('SO_ADRESSE2').AsString);
          TOBEtablissement.PutValue('ET_ADRESSE3',RSqlSociete.FindField('SO_ADRESSE3').AsString);
          TOBEtablissement.PutValue('ET_CODEPOSTAL',RSqlSociete.FindField('SO_CODEPOSTAL').AsString);
          TOBEtablissement.PutValue('ET_VILLE',RSqlSociete.FindField('SO_VILLE').AsString);
          TOBEtablissement.PutValue('ET_PAYS',RSqlSociete.FindField('SO_PAYS').AsString);
          TOBEtablissement.PutValue('ET_TELEPHONE',RSqlSociete.FindField('SO_TELEPHONE').AsString);
          TOBEtablissement.PutValue('ET_FAX',RSqlSociete.FindField('SO_FAX').AsString);
          TOBEtablissement.PutValue('ET_TELEX',RSqlSociete.FindField('SO_TELEX').AsString);
          TOBEtablissement.PutValue('ET_SIRET',RSqlSociete.FindField('SO_SIRET').AsString);
          TOBEtablissement.PutValue('ET_APE',RSqlSociete.FindField('SO_APE').AsString);
          TOBEtablissement.PutValue('ET_JURIDIQUE',RSqlSociete.FindField('SO_NATUREJURIDIQUE').AsString);
          TOBEtablissement.InsertDB(nil);
        finally
          TOBEtablissement.Free;
        end;
        SetParamSoc('SO_ETABLISDEFAUT', RSqlSociete.FindField('SO_SOCIETE').AsString);
        ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="'+RSqlSociete.FindField('SO_SOCIETE').AsString+'" WHERE SOC_NOM="SO_ETABLISDEFAUT"');
      end;
      Ferme (RSqlSociete);
    end;
  end
 else
  begin
    //--- Si il existe déjà un établissement dans la base, on ne crée pas d'établissement
    ChSql:='SELECT ET_ETABLISSEMENT FROM '+Nombase+'.DBO.ETABLISS ORDER BY ET_ETABLISSEMENT';
    RSqlEtablissement:= OpenSQL (ChSql, True,-1,'',true);
    if not RSqlEtablissement.Eof then
      //--- Maj table PARAMSOC
      ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+RSqlEtablissement.FindField('ET_ETABLISSEMENT').AsString+'" WHERE SOC_NOM="SO_ETABLISDEFAUT"')
    else
    begin
     TobEtablissement:=Tob.Create ('ETABLISS',Nil,-1);
     TobEtablissement.LoadDB;
     TobSociete:=Tob.Create ('SOCIETE',Nil,-1);
     TobSociete.LoadDetailFromSQL('SELECT * FROM '+NomBase+'.dbo.SOCIETE');

     if (TobSociete.Detail.count>0) then
      begin
       TobEtablissement.PutValue ('ET_ETABLISSEMENT',TobSociete.Detail[0].GetValue('SO_SOCIETE'));
       TobEtablissement.PutValue ('ET_SOCIETE',TobSociete.Detail[0].GetValue('SO_SOCIETE'));
       TobEtablissement.PutValue ('ET_LIBELLE',TobSociete.Detail[0].GetValue('SO_LIBELLE'));
       TobEtablissement.PutValue ('ET_ABREGE',Copy (TobSociete.Detail[0].GetValue('SO_LIBELLE'),1,17));
       TobEtablissement.PutValue ('ET_ADRESSE1',TobSociete.Detail[0].GetValue('SO_ADRESSE1'));
       TobEtablissement.PutValue ('ET_ADRESSE2',TobSociete.Detail[0].GetValue('SO_ADRESSE2'));
       TobEtablissement.PutValue ('ET_ADRESSE3',TobSociete.Detail[0].GetValue('SO_ADRESSE3'));
       TobEtablissement.PutValue ('ET_CODEPOSTAL',TobSociete.Detail[0].GetValue('SO_CODEPOSTAL'));
       TobEtablissement.PutValue ('ET_VILLE',TobSociete.Detail[0].GetValue('SO_VILLE'));
       TobEtablissement.PutValue ('ET_PAYS',TobSociete.Detail[0].GetValue('SO_PAYS'));
       TobEtablissement.PutValue ('ET_TELEPHONE',TobSociete.Detail[0].GetValue('SO_TELEPHONE'));
       TobEtablissement.PutValue ('ET_FAX',TobSociete.Detail[0].GetValue('SO_FAX'));
       TobEtablissement.PutValue ('ET_TELEX',TobSociete.Detail[0].GetValue('SO_TELEX'));
       TobEtablissement.PutValue ('ET_SIRET',TobSociete.Detail[0].GetValue('SO_SIRET'));
       TobEtablissement.PutValue ('ET_APE',TobSociete.Detail[0].GetValue('SO_APE'));
       TobEtablissement.PutValue ('ET_JURIDIQUE',TobSociete.Detail[0].GetValue('SO_NATUREJURIDIQUE'));
       ExecuteSQl (StringReplace (TobEtablissement.MakeInsertSql,'INTO ETABLISS','INTO '+NomBase+'.dbo.ETABLISS',[RfIgnoreCase]));

       //--- Maj table PARAMSOC
       ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+TobSociete.Detail[0].GetValue('SO_SOCIETE')+'" WHERE SOC_NOM="SO_ETABLISDEFAUT"');
      end;

     Ferme(RSqlEtablissement);
     TobSociete.Free;
     TobEtablissement.Free;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CATALA David
Créé le ...... : 31/08/2004
Modifié le ... :
Description .. : Initalisation de l'enregistrement établissement
Suite ........ : complémentaire  en création de
Suite ........ : dossier
Mots clefs ... :
*****************************************************************}
procedure InitEtablissementCompl (NomBase : String; Etablissement : String; Libelle : String);
var TobEtablissementCompl  : TOB;
    DateFermeture          : TDateTime;
    ANow,MNow,DNow         : Word;
    RSql                   : TQuery;
    RsqlEtablissementCompl : TQuery;
    ChSql                  : String;
begin
 //--- Si il existe déjà un établissement Complémentaire dans la base, on ne crée pas d'établissement complémentaire
 ChSql:='SELECT * FROM '+Nombase+'.DBO.ETABCOMPL WHERE ETB_ETABLISSEMENT="'+Etablissement+'"';
 RSqlEtablissementCompl:= OpenSQL (ChSql, True,-1,'',true);

 if RSqlEtablissementCompl.Eof then
  begin
   //--- Calcul DateFermetue
   DecodeDate (DATE, ANow, MNow, DNow);
   DateFermeture:=EncodeDate(ANow, 05, 31);
   if (DateFermeture<=Now) then
    DateFermeture := EncodeDate(ANow + 1, 05, 31);

   TobEtablissementCompl:=Tob.Create ('ETABCOMPL',Nil,-1);
   TobEtablissementCompl.LoadDB;

   TobEtablissementCompl.PutValue ('ETB_CONGESPAYES','-');
   TobEtablissementCompl.PutValue ('ETB_DATECLOTURECPN',DateFermeture);
   TobEtablissementCompl.PutValue ('ETB_CONGESPAYES', '-');
   TobEtablissementCompl.PutValue ('ETB_CODESECTION', '1');
   TobEtablissementCompl.PutValue ('ETB_BCJOURPAIEMENT', 0);
   TobEtablissementCompl.PutValue ('ETB_JEDTDU', 0);
   TobEtablissementCompl.PutValue ('ETB_JEDTAU', 0);
   TobEtablissementCompl.PutValue ('ETB_PRUDHCOLL', '1');
   TobEtablissementCompl.PutValue ('ETB_PRUDHSECT', '4');
   TobEtablissementCompl.PutValue ('ETB_PRUDHVOTE', '1');
   TobEtablissementCompl.PutValue ('ETB_ISLICSPEC', '-');
   TobEtablissementCompl.PutValue ('ETB_ISOCCAS', '-');
   TobEtablissementCompl.PutValue ('ETB_ISLABELP', '-');
   TobEtablissementCompl.PutValue ('ETB_ETABLISSEMENT', Etablissement);
   TobEtablissementCompl.PutValue ('ETB_LIBELLE', Libelle);
   TobEtablissementCompl.PutValue ('ETB_NBJOUTRAV', 0);
   TobEtablissementCompl.PutValue ('ETB_NBREACQUISCP', 0);
   TobEtablissementCompl.PutValue ('ETB_NBACQUISCP', '');
   TobEtablissementCompl.PutValue ('ETB_NBRECPSUPP', 0);
   TobEtablissementCompl.PutValue ('ETB_TYPDATANC', '1');
   TobEtablissementCompl.PutValue ('ETB_DATEACQCPANC', '');
   TobEtablissementCompl.PutValue ('ETB_VALORINDEMCP', '');
   TobEtablissementCompl.PutValue ('ETB_RELIQUAT', '');
   TobEtablissementCompl.PutValue ('ETB_PROFILCGE', '');
   TobEtablissementCompl.PutValue ('ETB_BASANCCP', '');
   TobEtablissementCompl.PutValue ('ETB_VALANCCP', '');
   TobEtablissementCompl.PutValue ('ETB_PERIODECP', 0);
   TobEtablissementCompl.PutValue ('ETB_1ERREPOSH', '');
   TobEtablissementCompl.PutValue ('ETB_2EMEREPOSH', '');
   TobEtablissementCompl.PutValue ('ETB_MVALOMS', '');
   TobEtablissementCompl.PutValue ('ETB_VALODXMN', 0);
   TobEtablissementCompl.PutValue ('ETB_TYPDATANC', '');
   TobEtablissementCompl.PutValue ('ETB_MVALOMS', '');
   TobEtablissementCompl.PutValue ('ETB_PCTFRAISPROF', 0);
   TobEtablissementCompl.PutValue ('ETB_HORAIREETABL', 0);
   TobEtablissementCompl.PutValue ('ETB_DATEVALTRANS', Idate1900);
   TobEtablissementCompl.PutValue ('ETB_SEUILTEMPSPAR', 0);
   TobEtablissementCompl.PutValue ('ETB_PRORATATVA', 0);
   TobEtablissementCompl.PutValue ('ETB_JOURPAIEMENT', 0);
   TobEtablissementCompl.PutValue ('ETB_EDITBULCP', '');
   TobEtablissementCompl.PutValue ('ETB_REGIMEALSACE', '-');
   TobEtablissementCompl.PutValue ('ETB_MEDTRAV', -1);
   TobEtablissementCompl.PutValue ('ETB_CODEDDTEFP', -1);
   TobEtablissementCompl.PutValue ('ETB_SUBROGATION','-');

   RSql:=OpenSQL('SELECT SOC_DATA FROM '+NomBase+'.dbo.PARAMSOC WHERE SOC_NOM="SO_PGJOURHEURE"', True,-1,'',true);
   if Not RSql.Eof then TobEtablissementCompl.PutValue ('ETB_JOURHEURE',RSql.FindField('SOC_DATA').AsString);
    Ferme(RSql);

   ExecuteSQl (StringReplace (TobEtablissementCompl.MakeInsertSql,'INTO ETABCOMPL','INTO '+NomBase+'.dbo.ETABCOMPL',[RfIgnoreCase]));
   TobEtablissementCompl.Free;
  end;

 Ferme (RSqlEtablissementCompl);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLCLIENT}
function InitTobParamProcessServer : TOB ;
begin
  Result := TOB.Create('$PARAM', nil, -1) ;
  // Paramètres pour l'appel du process server :
(*
 CA - 17/07/2007 - Suppression car renseigné par CBP (dixit XP)
Avec ce code, on avait des soucis de connexion en authentification NT

  Result.AddChampSupValeur('USERLOGIN' , V_PGI.UserLogin ) ;
  Result.AddChampSupValeur('INIFILE'   , HalSocIni ) ;
  Result.AddChampSupValeur('PASSWORD'  , V_PGI.Password ) ;
  Result.AddChampSupValeur('DOMAINNAME', V_PGI.DomainName ) ;
  Result.AddChampSupValeur('DATEENTREE', V_PGI.DateEntree ) ;
  Result.AddChampSupValeur('DOSSIER'   , V_PGI.CurrentAlias ) ;
*)
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/12/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure SupprimeTableStd(TableDos: string);
begin
  ExecuteSQL('Delete from ' + TableDos);
end;

////////////////////////////////////////////////////////////////////////////////
procedure SupprimeTableChoixCod(Code: string);
begin
  ExecuteSQL('Delete from CHOIXCOD Where CC_TYPE="' + Code + '"');
end;

////////////////////////////////////////////////////////////////////////////////
(*
procedure SupprimeTableChoixListe(LIST: string);
begin
  ExecuteSQL('Delete from LISTE Where LI_LISTE="' + LIST + '"');
end;*)

procedure ChargeParamSocRef(NumPlan: integer);
var
  QParamSocRef: TQuery;
  lgene, lgaux: integer;
  TPSoc, TPSocI : TOB;
  i : integer;
  StVal: WideString;
  BourreGen, BourreAux : string;
begin
  // GCO - 24/04/2006 - FQ 17783,17785
  // CA - 25/06/2007 - Ajout 'WHERE' pour limitée la réinitialisation à la comptabilité
  ExecuteSql('UPDATE PARAMSOC SET SOC_DATA = "" WHERE SOC_TREE LIKE "001;001%"');

  lgene := 8; lgaux := 8; BourreGen := '0'; BourreAux := '';
  TPSoc := TOB.Create ('',nil,-1);
  try
    TPSoc.LoadDetailFromSQL('SELECT PRR_SOCNOM,PRR_SOCDATA FROM PARSOCREF WHERE PRR_SOCNOM IN ("SO_LGCPTEGEN","SO_LGCPTEAUX","SO_BOURREGEN","SO_BOURREAUX") AND PRR_NUMPLAN='+IntToStr(NumPlan));
    for i:= 0 to TPSoc.Detail.Count - 1 do
    begin
      TPSocI := TPSoc.Detail[i];
      if TPSocI.GetString ('PRR_SOCNOM')='SO_LGCPTEGEN' then
        lgene := StrToInt(TPSocI.GetString ('PRR_SOCDATA'))
      else if TPSocI.GetString ('PRR_SOCNOM')='SO_LGCPTEAUX' then
        lgaux := StrToInt(TPSocI.GetString ('PRR_SOCDATA'))
      else if TPSocI.GetString ('PRR_SOCNOM')='SO_BOURREGEN' then
        BourreGen := Copy(TPSocI.GetString ('PRR_SOCDATA'),1,1)
      else if TPSocI.GetString ('PRR_SOCNOM')='SO_BOURREAUX' then
        BourreAux := Copy(TPSocI.GetString ('PRR_SOCDATA'),1,1);
    end;
  finally
    TPSoc.Free;
  end;
  SetParamsoc ('SO_LGCPTEGEN',lgene);
  SetParamsoc ('SO_LGCPTEAUX',lgaux);
  SetParamsoc ('SO_BOURREGEN',BourreGen);
  if BourreAux = '' then BourreAux := BourreGen;
  SetParamsoc ('SO_BOURREAUX',BourreAux);

(*
  // Recherche des infos de bourrage des comptes
  lgene := StrToInt (GetColonneSQL ('PARSOCREF','PRR_SOCDATA','PRR_SOCNOM="SO_LGCPTEGEN" AND PRR_NUMPLAN='+IntToStr(NumPlan)));
  lgaux := StrToInt (GetColonneSQL ('PARSOCREF','PRR_SOCDATA','PRR_SOCNOM="SO_LGCPTEAUX" AND PRR_NUMPLAN='+IntToStr(NumPlan)));
//  BourreGen := GetColonneSQL ('PARSOCREF','PRR_SOCDATA','PRR_SOCNOM="SO_BOURREGEN" AND PRR_NUMPLAN='+IntToStr(NumPlan))[1];
//  StVal := GetColonneSQL ('PARSOCREF','PRR_SOCDATA','PRR_SOCNOM="SO_BOURREAUX" AND PRR_NUMPLAN='+IntToStr(NumPlan));

  if StVal <> '' then BourreAux := StVal[1] else BourreAux := BourreGen;
  SetParamsoc ('SO_LGCPTEGEN',lgene);
  SetParamsoc ('SO_LGCPTEAUX',lgaux);
  SetParamsoc ('SO_BOURREGEN',BourreGen);
  SetParamsoc ('SO_BOURREAUX',BourreAux);
*)
  // CA - 25/06/2007 - Par défaut, on force la conformité BOI
  SetParamSoc ('SO_CPCONFORMEBOI',True);
  // Mise à jour des VH^
{$IFNDEF NOVH}
  RechargeParamSoc;
{$ENDIF}
  // Chargement des paramètres société
  QParamSocRef := OpenSQL('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
    IntToStr(NumPlan) + ' ORDER BY PRR_SOCNOM', True);
  {$IFNDEF EAGLSERVER}
  InitMove(RecordsCount(QParamSocRef), 'Chargement des paramètres société');
  {$ENDIF}
  while not QParamSocRef.EOF do
  begin
    StVal := QParamSocRef.FindField('PRR_SOCDATA').AsString;
    // On ne bourre plus les comptes car dans le PARSOCREF, les comptes
    // ont la taille indiquée dans le paramétrage.
    if (QParamSocRef.FindField('PRR_COMPTE').AsString = 'G') then
      StVal := BourreLaDonc(stVal, fbGene)
    else if (QParamSocRef.FindField('PRR_COMPTE').AsString = 'T') then
      StVal := BourreLaDonc(stVal, fbAux);
    if QParamSocRef.FindField('PRR_SOCNOM').AsString = 'SO_LGCPTEGEN' then
    begin
      lgene := StrToInt(QParamSocRef.FindField('PRR_SOCDATA').AsString);
      SetParamSoc('SO_LGCPTEGEN', lgene);
    end;
    if QParamSocRef.FindField('PRR_SOCNOM').AsString = 'SO_LGCPTEAUX' then
    begin
      lgaux := StrToInt(QParamSocRef.FindField('PRR_SOCDATA').AsString);
      SetParamSoc('SO_LGCPTEAUX', lgaux);
    end;
    SetParamSoc(QParamSocRef.FindField('PRR_SOCNOM').AsString, StVal);
    QParamSocRef.Next;
    {$IFNDEF EAGLSERVER}
    MoveCur(False);
    {$ENDIF}
  end;
  Ferme(QParamSocRef);
  {$IFNDEF EAGLSERVER}
  FiniMove;
  {$ENDIF}

  ChargeTobSoc;

end;

procedure LoadStandardRef(NumStd: integer; TableDos, TableStd: string);
var
  QDos, QStd: TQuery;
  PrefDos, PrefStd, Suffixe, ChampDos: string;
  ValChamp: variant;
  i: integer;
begin
  PrefDos := TableToPrefixe(TableDos);
  PrefStd := TableToPrefixe(TableStd);
  QStd := OpenSQL('SELECT * FROM ' + TableStd + ' WHERE ' + PrefStd + '_NUMPLAN='
    + IntToStr(NumStd), True);
  if not QStd.Eof then // Si un tel standard existe
  begin
    {$IFNDEF EAGLSERVER}
    InitMove(RecordsCount(QStd), 'Chargement Table : ' + TableDos);
    {$ENDIF}
    QDos := OpenSQL('SELECT * FROM ' + TableDos, False);
    while not QStd.Eof do
    begin
      QDos.Insert;
      InitNew(QDos);
      for i := 0 to QStd.FieldCount - 1 do
      begin
        Suffixe := ExtractSuffixe(QStd.Fields[i].FieldName);
        // Gestion des exceptions à la règle ici
        if ((PrefStd = 'PR') and (Suffixe = 'COMPTE')) then
          ChampDos := 'GER_GENERAL'
        else if ((PrefStd = 'PR') and (Suffixe = 'REPORTDETAIL')) then
          continue
            // Fin de la gestion des exceptions
        else
          ChampDos := PrefDos + '_' + Suffixe;
        if ChampDos = 'GER_BLOCNOTE' then
          continue;
        // calcul de la valeur du champ
        if ChampDos = 'PR_COMPTE' then
          ValChamp := BourreLaDonc(QStd.FindField('GER_GENERAL').AsString,
            fbGene)
            // Fin modification
        else
          ValChamp := QStd.FindField(PrefStd + '_' + Suffixe).AsVariant;
        QDos.FindField(ChampDos).AsVariant := ValChamp;
      end;
      QDos.Post;
      {$IFNDEF EAGLSERVER}
      MoveCur(False);
      {$ENDIF}
      QStd.Next;
    end;
    Ferme(QDos);
    {$IFNDEF EAGLSERVER}
    FiniMove;
    {$ENDIF}
  end;
  Ferme(QStd);
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/12/2004
Modifié le ... : 16/12/2004
Description .. :
Mots clefs ... :
*****************************************************************}
function ChargeToutLeStandard( vNumStd : integer; vBoCtxStd : Boolean; vBoAvecProgressForm : Boolean) : Boolean;
var lInNbMoveCur : integer;
begin
  BeginTrans;
  try
    if vBoCtxStd then
      lInNbMoveCur := 5
    else
      lInNbMoveCur := 4;

    if vBoAvecProgressForm then InitMoveProgressForm( nil, 'Chargement du dossier type', 'Veuillez patienter...', lInNbMoveCur , False, True);

    if vBoCtxStd then
    begin
      // Suppression du contenu des tables à mettre à jour
      if vBoAvecProgressForm then MoveCurProgressForm('Initialisation de l''environnement...');
      SupprimeTableStd('GENERAUX');
      SupprimeTableStd('TIERS');
      SupprimeTableStd('TIERSCOMPL'); // GCO - 02/06/2004
      SupprimeTableStd('JOURNAL');
      SupprimeTableStd('GUIDE');
      SupprimeTableStd('ECRGUI');
      SupprimeTableStd('ANAGUI');
      SupprimeTableStd('CORRESP');
      SupprimeTableStd('RUPTURE');
      SupprimeTableStd('AXE');
      SupprimeTableStd('SECTION');
      SupprimeTableChoixCod('VTY');
      SupprimeTableChoixCod('RUG');
      SupprimeTableChoixCod('RUT');
      SupprimeTableChoixCod('RTV');
      SupprimeTableChoixCod('QME'); // Qualififant quantité
      SupprimeTableChoixCod('TX1'); // TVA par régime fiscal
      SupprimeTableChoixCod('TX2'); // TPF Par régime fiscal
      SupprimeTableStd('VENTIL');
      SupprimeTableStd('NATCPTE');
      SupprimeTableStd('FILTRES');
      // GCO - 04/05/2005
      SupprimeTableStd('TXCPTTVA');
      SupprimeTableStd('REFAUTO');
      // FIN GCO
      ExecuteSQL ('DELETE FROM CMASQUESAISIE WHERE CMS_TYPE<>"CEG"');
      ExecuteSQL ('DELETE FROM CMASQUECRITERES WHERE CMC_TYPE<>"CEG"');
      SupprimeTableStd('CLOIVENTIL');

      // A dégager , après délibération avec CJ
      //SupprimeTableChoixListe('MULMMVTS');
      //SupprimeTableChoixListe('MULMANAL');
      //SupprimeTableChoixListe('MULVMVTS');

      // Chargement du paramsoc
      ChargeParamSocRef( vNumStd );

      SetParamSoc('SO_NUMPLANREF', vNumStd);
      SetParamSoc('SO_BOURREGEN', '0');
      SetParamSoc('SO_BOURREAUX', '0');

      if (vNumStd > 0) and (vNumStd <> GetParamSocSecur('SO_NUMPLANREF',0)) then {Lek 100206}
        SetParamSoc('SO_NUMPLANREF', vNumStd);
{$IFNDEF NOVH}
      ChargeMagHalley;
{$ENDIF}
    end;

    // Chargement des tables depuis les standards
    if vBoAvecProgressForm then MoveCurProgressForm('Chargement des comptes généraux...');
    LoadStandardCompta(vNumStd, 'GENERAUX', 'GENERAUXREF');

    if vBoAvecProgressForm then MoveCurProgressForm('Chargement des journaux...');
    LoadStandardCompta(vNumStd, 'JOURNAL', 'JALREF');
    LoadStandardCompta(vNumStd, 'GUIDE', 'GUIDEREF');
    LoadStandardCompta(vNumStd, 'ECRGUI', 'ECRGUIREF');
    LoadStandardCompta(vNumStd, 'ANAGUI', 'ANAGUIREF');

    if vBoAvecProgressForm then MoveCurProgressForm('Chargement des comptes auxiliaires...');
    LoadStandardCompta(vNumStd, 'TIERS', 'TIERSREF');
    LoadStandardCompta(vNumStd, 'CORRESP', 'CORRESPREF');
    LoadStandardCompta(vNumStd, 'RUPTURE', 'RUPTUREREF');
    // GCO - 09/05/2005
    LoadStandardCompta(vNumStd, 'TXCPTTVA', 'TXCPTTVAREF');
    // GCO - 25/04/2006 - Correction du Probleme ADER,
    // Regenere dans choixcod à partir de TXCPTTVA
    CreateCodTXX;

    LoadStandardCompta(vNumStd, 'REFAUTO', 'REFAUTOREF');
    // FIN GCO

    LoadStandardMaj(vNumStd, 'AXE', 'AXEREF', 'AXE', ' ');
    LoadStandardCompta(vNumStd, 'SECTION', 'SECTIONREF');

    { Déversement des masques de saisie }
    YFSChargeTableDepuisStandard ( 'CMASQUECRITERES',vNumStd,'');
    YFSChargeTableDepuisStandard ( 'CMASQUESAISIE',vNumStd,'');
    { Déversement des lois de répartition }
    YFSChargeTableDepuisStandard ( 'CLOIVENTIL',vNumStd,'');

    LoadStandardMaj(vNumStd, 'CHOIXCOD', 'CHOIXCODREF', 'TYPE;CODE', ' ');
    if not ExisteSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="NAT"') then InitParamTablesLibres ('');
    // GCO - 11/10/2005 - FQ 16510
    if not ExisteSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE = "RTV"') then
    begin
      ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE, CC_CODE, CC_LIBELLE, CC_ABREGE, CC_LIBRE) ' +
               '(SELECT CCR_TYPE, CCR_CODE, CCR_LIBELLE, CCR_ABREGE, CCR_LIBRE FROM ' +
               'CHOIXCODREF WHERE CCR_NUMPLAN = ' + IntToStr(vNumStd) + ' AND CCR_TYPE = "RTV")');
    end;

    LoadStandardCompta(vNumStd, 'VENTIL', 'VENTILREF');
    LoadStandardCompta(vNumStd, 'NATCPTE', 'NATCPTEREF');

    if vBoCtxStd then
    begin
      LoadStandardCompta(vNumStd, 'MODEPAIE', 'MODEPAIEREF'); // CA - 02/04/2002
      LoadStandardCompta(vNumStd, 'MODEREGL', 'MODEREGLREF');
    end
    else
    begin // Assistant de création de dossier
      // CA - 02/04/2002
      // Pour les tables modepaie et moderegl, des éléments sont déjà présents dans une base vide
      // De plus les standards cegid n'ont pas d'informations sur ces tables.
      // On ne déverse ces infos que si les tables de standards contiennent de telles informations
      if ExisteSQL('SELECT * FROM MODEPAIEREF WHERE MPR_NUMPLAN = ' + IntToStr(vNumStd)) then
      begin
        ExecuteSQL('DELETE FROM MODEPAIE');
        LoadStandardCompta(vNumStd, 'MODEPAIE', 'MODEPAIEREF');
      end;
      if ExisteSQL('SELECT * FROM MODEREGLREF WHERE MRR_NUMPLAN = ' + IntToStr(vNumStd)) then
      begin
        ExecuteSQL('DELETE FROM MODEREGL');
        LoadStandardCompta(vNumStd, 'MODEREGL', 'MODEREGLREF');
      end;
      // Fin CA - 02/04/2002
    end;

    if vBoAvecProgressForm then MoveCurProgressForm('Chargement des filtres...');
    LoadStandardCompta(vNumStd, 'FILTRES', 'FILTRESREF');

    // GCO - 15/12/2004 - à dégager
    //LoadStandardMaj(NumPlanCompte, 'LISTE', 'LISTEREF', 'LISTE', ' ');

    if vBoCtxStd then
      ChargeAxe( vNumStd );

    CommitTrans;
    
    Result := True;
  except
    on E: Exception do
    begin
      RollBack;
      Result := False;
      {$IFNDEF EAGLSERVER}
      PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : ChargeToutLeStandard');
      {$ENDIF}
    end;
  end;

  if vBoAvecProgressForm then FiniMoveProgressForm;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function EnregistreToutLeStandard( vNumStd : integer ; vStLibelle, vStAbrege : string; vBoAvecProgressForm : Boolean) : Boolean;
begin
  BeginTrans;
  try
    // Suppression du dossier type
    if vBoAvecProgressForm then InitMoveProgressForm( nil, 'Enregistrement du standard', 'Veuillez patienter...', 8, False, True);
    SupprimeDossierType(vNumStd);

    // Enregistrement de l'entête
    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement de l''entête ...');
    EnregistreInfoStandard(vNumStd, vStLibelle, vStAbrege);

    // Enregistrement des éléments du dossier type
    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des paramètres ...');
    EnregistreParamSocStandard(vNumStd);

    // Enregistrement du plan comptable
    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des comptes généraux ...');
    SaveAsStandardCompta(vNumStd, 'GENERAUX', 'GENERAUXREF');

    // Enregistrement des tiers
    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des comptes auxiliaires ...');
    SaveAsStandardCompta(vNumStd, 'TIERS', 'TIERSREF');

    // Enregistrement des journaux
    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des journaux ...');
    SaveAsStandardCompta(vNumStd, 'JOURNAL', 'JALREF');

    // Enregistrement des guides
    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des guides ...');
    SaveAsStandardCompta(vNumStd, 'GUIDE', 'GUIDEREF');
    SaveAsStandardCompta(vNumStd, 'ECRGUI', 'ECRGUIREF');
    SaveAsStandardCompta(vNumStd, 'ANAGUI', 'ANAGUIREF');

    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des filtres ...');
    SaveAsStandardCompta(vNumStd, 'FILTRES', 'FILTRESREF');

    // GCO - 15/12/2004 - à dégager
    //if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des listes ...');
    //SaveInfoListe(vNumStd, 'MULMMVTS');
    //SaveInfoListe(vNumStd, 'MULMANAL');
    //SaveInfoListe(vNumStd, 'MULVMVTS');

    if vBoAvecProgressForm then MoveCurProgressForm('Enregistrement des tables annexes ...');
    SaveAsStandardCompta(vNumStd, 'CORRESP', 'CORRESPREF');
    SaveAsStandardCompta(vNumStd, 'RUPTURE', 'RUPTUREREF');
    SaveAsStandardCompta(vNumStd, 'NATCPTE', 'NATCPTEREF');

    // GCO - 09/05/2005
    SaveAsStandardCompta(vNumStd, 'TXCPTTVA', 'TXCPTTVAREF');
    SaveAsStandardCompta(vNumStd, 'REFAUTO', 'REFAUTOREF');
    // FIN GCO

    SaveAsStandardCompta(vNumStd, 'AXE', 'AXEREF');
    SaveAsStandardCompta(vNumStd, 'SECTION', 'SECTIONREF');
    SaveAsStandardCompta(vNumStd, 'CHOIXCOD', 'CHOIXCODREF');
    SaveAsStandardCompta(vNumStd, 'VENTIL', 'VENTILREF');
    // CA - 02/04/2002
    SaveAsStandardCompta(vNumStd, 'MODEPAIE', 'MODEPAIEREF');
    SaveAsStandardCompta(vNumStd, 'MODEREGL', 'MODEREGLREF');
    // Fin CA - 02/04/2002

    //SaveAsStandardCompta(vNumStd, 'CMASQUECRITERES', 'CMASQUECRITREF');
    //SaveAsStandardCompta(vNumStd, 'CMASQUESAISIE', 'CMASQUESAISIEREF');

    // Enregistrements des tables en base
    YFSEnregistreTableVersStandard ( 'CMASQUECRITERES', vNumStd);
    YFSEnregistreTableVersStandard ( 'CMASQUESAISIE', vNumStd);
    YFSEnregistreTableVersStandard ( 'CLOIVENTIL', vNumStd);

    SetParamSoc('SO_NUMPLANREF', vNumStd);
    CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      RollBack;
      Result := False;
      {$IFNDEF EAGLSERVER}
      PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : EnregistreToutLeStandard');
      {$ENDIF}
    end;
  end;

  if vBoAvecProgressForm then FiniMoveProgressForm;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ChargeAxe( vNumStd : integer );
var Q1, Q2 : TQuery;
    TobAxe : TOB;
begin
  Q1 := OpenSQL('SELECT * FROM AXEREF WHERE XRE_NUMPLAN="' + IntToStr( vNumStd ) + '"', True,-1,'',true);
  if Q1.EOF then
  begin
    Ferme(Q1);
    Tobaxe := TOB.Create('AXE', nil, -1);
    Q2 := OpenSql('SELECT * FROM ##DP##.AXE', True,-1,'',true);
    Q2.First;
    while not Q2.EOF do
    begin
      Tobaxe.PutValue('X_AXE', Q2.FindField('X_AXE').Asstring);
      Tobaxe.PutValue('X_LIBELLE', Q2.FindField('X_LIBELLE').Asstring);
      Tobaxe.PutValue('X_COMPTABLE', Q2.FindField('X_COMPTABLE').Asstring);
      Tobaxe.PutValue('X_CHANTIER', Q2.FindField('X_CHANTIER').Asstring);
      Tobaxe.PutValue('X_MODEREOUVERTURE', Q2.FindField('X_MODEREOUVERTURE').Asstring);
      Tobaxe.PutValue('X_SECTIONATTENTE', Q2.FindField('X_SECTIONATTENTE').Asstring);
      Tobaxe.PutValue('X_REGLESAISIE', Q2.FindField('X_REGLESAISIE').Asstring);
      Tobaxe.PutValue('X_ABREGE', Q2.FindField('X_ABREGE').Asstring);
      Tobaxe.PutValue('X_LONGSECTION', Q2.FindField('X_LONGSECTION').Asinteger);
      Tobaxe.PutValue('X_BOURREANA', Q2.FindField('X_BOURREANA').Asstring);
      Tobaxe.PutValue('X_SOCIETE', Q2.FindField('X_SOCIETE').Asstring);
      Tobaxe.PutValue('X_STRUCTURE', Q2.FindField('X_STRUCTURE').Asstring);
      Tobaxe.PutValue('X_GENEATTENTE', Q2.FindField('X_GENEATTENTE').Asstring);
      Tobaxe.PutValue('X_CPTESTRUCT', Q2.FindField('X_CPTESTRUCT').Asstring);
      Tobaxe.PutValue('X_FERME', Q2.FindField('X_FERME').Asstring);
      Tobaxe.PutValue('X_SAISIETRANCHE', Q2.FindField('X_SAISIETRANCHE').Asstring);
      Tobaxe.InsertOrUpdateDB(TRUE);
      Q2.Next;
    end;

    Ferme( Q2 );
    if Tobaxe <> nil then
      Tobaxe.Free;
  end
  else
    Ferme(Q1);
end;

{$IFNDEF EAGLSERVER}
function _TestTvaEncaissement : boolean ;
begin
 result := false ;
 // on bloque si des comptes 41 des ecritures multidevise
 if ExisteSQL('select E_GENERAL from ecriture where ( E_ECHE="X" ) and ( E_NUMECHE > 1 ) and ( E_GENERAL like "41%" ) ' ) then
  begin
   PgiInfo('Gestion de la TVA impossible. Des comptes 41 ont des écritures multi-échéance !');
   exit ;
  end ;
 result := true ;
end ;
{$ENDIF} 

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/04/2006
Modifié le ... : 31/08/2006
Description .. : - LG - 19/06/2006 - FB 18553 - controle des compte multi
Suite ........ : eche pour l'activation de la tva
Suite ........ : -LG - 31/08/2006 - FB 18492 - Quelque soit le choix fait 
Suite ........ : dans le DP, le champs so_ouitvaenc reste à - au lieu de 
Suite ........ : passer à X. Du coup se sont les comptes 7 qui sont rendus 
Suite ........ : ventilables.
Mots clefs ... : 
*****************************************************************}
function ActivationTvaSurDossier( vBoParAligneStd : Boolean = False; vNumStd : integer = 0) : boolean ;
{$IFNDEF EAGLSERVER}
var
 lBoTvaSurEncaissement : Boolean;
 lQ                    : TQuery ;
 lBoTVAEnc             : boolean ;
{$ENDIF}
begin
 result := false ;
{$IFDEF EAGLSERVER}
  DDWriteLn('ActivationTvaSurDossier, Fonction non disponible en EAGLSERVER');
{$ELSE}
  if VH^.AnaCroisaxe then
  begin
    if vBoParAligneStd then
      // Pas de message bloquant
    else
      PgiInfo('Gestion de la TVA impossible. Analytiques "croise-axes" activés');
  end
  else
  begin
   if vBoParAligneStd then // On va cherche le paramsoc du standard dans PARSOCREF
    begin
      lBoTvaSurEncaissement := GetColonneSQL('PARSOCREF', 'PRR_SOCDATA',
        ' PRR_NUMPLAN= '+ IntToStr(vNumStd) + ' AND PRR_SOCNOM = "SO_OUITVAENC"') = 'X';

      if not _TestTvaEncaissement then exit ;
      SetParamSoc('SO_OUITVAENC', lBoTvasurEncaissement);
    end
    else
    begin
      lQ         := OpenSQL('select DFI_EXIGIBILITE,DFI_OPTIONEXIG FROM DPFISCAL LEFT JOIN ' +
                            'DOSSIER ON DOS_GUIDPER = DFI_GUIDPER WHERE '+
                            'DOS_NODOSSIER = "' + V_PGI.NoDossier + '"'
                            , true,-1,'',true) ;
      if not lQ.Eof then
        lBoTVAEnc  := ( lQ.FindField('DFI_EXIGIBILITE').asString = 'TE' ) and ( lQ.FindField('DFI_OPTIONEXIG').asString <> 'X' )
      else lBoTVAEnc  := False;
      Ferme(lQ) ;
    //  if PgiAsk('Voulez vous gérer la TVA sur les encaissements ?', 'Assistant Comptabilité') = MrYes then
      if lBoTVAEnc then
      begin
        if not _TestTvaEncaissement then exit ;
        SetParamSoc('SO_OUITVAENC', True);
      end
      else
       SetParamSoc('SO_OUITVAENC', False) ;
    end ;
    SetParamSoc('SO_CPPCLSAISIETVA', True);
    // Ajout GCO - 06/12/2005
    if GetParamSocSecur('SO_CPPCLSANSANA', False) then
    begin
      SetParamSoc('SO_ZSAISIEANAL', True);
      SetParamSoc('SO_ZGEREANAL', True);
      SetParamSoc('SO_CPPCLSANSANA', False);
    end;
    InitialisationTVA( vBoParAligneStd, vNumStd);
          // ajout me E_IO
    ExecuteSQL('UPDATE  ECRITURE SET E_IO = "X"');

    result := true ;
    AvertirCacheServer('') ; // LG 01/062007 - pour forcer la prise en compte eneAGL
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLSERVER}
procedure InitialisationTVA( vBoParAligneStd : Boolean = False; vNumStd : integer = 0);
begin
  DDWriteLn('InitialisationTVA, Fonction non disponible en EAGLSERVER');
end;
{$ELSE}
procedure InitialisationTVA( vBoParAligneStd : Boolean = False; vNumStd : integer = 0);
  procedure _ChargeInfoAxe (piAxe : integer);
  var
    lQ         : TQuery;
    fb         : TFichierBase ;
    Stb        : string;
  begin
    lQ:=OpenSQL('SELECT * FROM AXE WHERE X_AXE="A'+IntToStr(piAxe)+'"',TRUE, -1, 'AXE',true) ;
    try
      if not lQ.EOF then
      begin
        fb:=AxeToFb(LQ.FindField('X_AXE').AsString) ;
        VH^.Cpta[fb].Lg:=LQ.FindField('X_LONGSECTION').AsInteger ;
        Stb:=lQ.FindField('X_BOURREANA').AsString ;
        if Stb<>'' then
        begin
          VH^.Cpta[fb].Cb:=Stb[1] ;
        end else
        begin
          VH^.Cpta[fb].Cb:='0' ;
        end;
        VH^.SaisieTranche[piAxe]:=(lQ.FindField('X_SAISIETRANCHE').AsString='X') ;
        VH^.Cpta[fb].Chantier:=lQ.FindField('X_CHANTIER').AsString='X' ;
        VH^.Cpta[fb].Structure:=lQ.FindField('X_STRUCTURE').AsString='X' ;
        VH^.Cpta[fb].Attente:=lQ.FindField('X_SECTIONATTENTE').AsString ;
        VH^.Cpta[fb].AxGenAttente:=lQ.FindField('X_GENEATTENTE').AsString ;
      end;
    finally
      Ferme(lQ) ;
    end;
  end;
var i         : integer;
    lQuery    : TQuery;
begin
  // Recherche d'un Axe analytique disponible
  for i := 0 to 5 do
  begin
    if not ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE ' +
                     'G_VENTILABLE' + IntToStr(i+1) + ' = "X"') then
    begin
      // Création de l'AXE TVA
      ExecuteSQL('UPDATE AXE SET X_LIBELLE = "TVA", X_ABREGE = "TVA", X_SECTIONATTENTE = "ZZZZ", ' +
                 'X_STRUCTURE = "-", X_GENEATTENTE = "", X_LONGSECTION = 4 WHERE ' +
                 'X_AXE = "A' + IntToStr(i+1) + '"');

      SetParamSoc('SO_CPPCLAXETVA', 'A' + IntToStr(i+1));

      Break;
    end;
  end;

  //
  CreationSectionTva('A'+ IntToStr(i+1));

  if vBoParAligneStd then exit ;

  // Affectation
  try
   if GetParamSocSecur('SO_OUITVAENC', false) then
      begin
        ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE = "X", ' +
                   'G_VENTILABLE' + IntToStr(i+1) + ' = "X" WHERE G_GENERAL LIKE "41%"');

        lQuery := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL LIKE "41%"', True);
      end
      else
      begin
        ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE = "X", ' +
                   'G_VENTILABLE' + IntToStr(i+1) + ' = "X" WHERE G_GENERAL LIKE "70%"');

        lQuery := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL LIKE "70%"', True);
      end;

      InitMoveProgressForm( nil, 'Activation de la TVA', 'Veuillez patienter...', lQuery.RecordCount , False, True);

      while not lQuery.Eof do
      begin
        _ChargeInfoAxe (i+1);

        {$IFDEF COMPTA}
        BeginTrans ;
        try
         MoveCurProgressForm('Traitement du compte : ' + lQuery.FindField('G_GENERAL').AsString );
         MajGVentil(False, lQuery.FindField('G_GENERAL').AsString);
         CommitTrans ;
        except
         on E : Exception do
          begin
           PGIError('Erreur sur le compte ' + lQuery.FindField('G_GENERAL').AsString + #10#13 + E.Message ) ;
           Rollback ;
          end ;
        end ;
        {$ENDIF}

        {$IFDEF CCSTD}
        MajGVentil(False, lQuery.FindField('G_GENERAL').AsString);
        {$ENDIF}

        lQuery.Next;
      end;

    FiniMoveProgressForm

  finally
{$IFDEF COMPTA}
  // fiche 19491
    MAJHistoparam ('GENERAUX',  '');
    MAJHistoparam ('SECTION',  '');
{$ENDIF}
    Ferme(lQuery);
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/04/2006
Modifié le ... :   /  /
Description .. : Création des SECTIONS de TVA
Mots clefs ... :
*****************************************************************}
procedure CreationSectionTva( vStAxe : string = '');
var lStAxe : string;
    lQuery : TQuery;
    lTob : Tob;
begin
  if vStAxe = '' then
  begin
    lStAxe := GetParamSocSecur('SO_CPPCLAXETVA', '', True);
    if lStAxe = '' then Exit;
  end
  else
    lStAxe := vStAxe;

  try
    ExecuteSQL('DELETE SECTION WHERE S_AXE = "' + lStAxe + '"');

    lQuery := OpenSQL('SELECT * FROM CHOIXDPSTD WHERE YDS_TYPE = "TVA" ORDER BY YDS_CODE', True,-1,'',true);
    while not lQuery.Eof do
    begin
      lTob := Tob.Create('SECTION', nil, -1);
      lTob.SetString('S_SECTION', lQuery.FindField('YDS_CODE').AsString);
      lTob.SetString('S_LIBELLE', lQuery.FindField('YDS_LIBELLE').AsString);
      lTob.SetString('S_ABREGE',  lQuery.FindField('YDS_ABREGE').AsString);
      lTob.SetString('S_SENS',    'M');
      lTob.SetDateTime('S_DATECREATION', Now);
      lTob.SetDateTime('S_DATEMODIF', Now);
      lTob.SetString('S_SOLDEPROGRESSIF', 'X');
      lTob.SetString('S_AXE', lStAxe);
      lTob.SetString('S_UTILISATEUR', V_PGI.User);
      lTob.SetString('S_SOCIETE', V_Pgi.CodeSociete);
      lTob.SetString('S_CREERPAR', 'TVA');
      lTob.InsertDb( nil, False);
      FreeAndNil(lTob);
      lQuery.Next;
    end;
  finally
    Ferme(lQuery);
  end;
end;

{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 12/09/2006
Modifié le ... :   /  /    
Description .. : LG - 12/09/2006 - chg du message sur l'activation de la tva
Mots clefs ... : 
*****************************************************************}
function DesactivationTvasurDossier : boolean;
var
 lStAxe        : string ;
 lBoOuiTvaEnc  : boolean ;
 lStAxeTva     : string ;
 lStSql        : string ;
begin

 result := false ;

 if not GetParamSocSecur('SO_CPPCLSAISIETVA', false) then exit ;

 if PgiAsk('Voulez vous désactiver la TVA ?', 'Assistant Comptabilité') = MrYes then
  if PgiAsk('Cette opération est irréversible, voulez vous continuer ?', 'Assistant Comptabilité') = MrYes then
   begin
    InitMoveProgressForm( nil, 'Désactivation de la TVA', 'Veuillez patienter...', 6 , False, True);
    lStAxe       := GetParamSocSecur('SO_CPPCLAXETVA', '') ;
    if lStAxe =  '' then exit ;
    lBoOuiTvaEnc := GetParamSocSecur('SO_OUITVAENC', '') ;
    lStAxeTva    := Copy(lStAxe,2,1) ;

    if lBoOuiTvaEnc then
     lStSql := ' G_GENERAL LIKE "41%"'
      else
       lStSql := ' G_GENERAL LIKE "70%"' ;

    BeginTrans ;

    try

    MoveCurProgressForm('Traitement des généraux...');
    ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE' + lStAxeTva + ' = "-" WHERE ' + lStSql ) ;
    MoveCurProgressForm('Traitement des sections...');
    ExecuteSQL('DELETE VENTIL WHERE V_SECTION IN ( SELECT S_SECTION FROM SECTION WHERE S_AXE="A' + lStAxeTva + '") ' ) ;
    MoveCurProgressForm('Traitement des axes...');
    ExecuteSQL('UPDATE AXE SET X_LIBELLE = "Axe n°' + lStAxeTva + '" , X_SECTIONATTENTE = "" ' +
               'WHERE X_AXE = "A' + lStAxeTva + '" ') ;
    MoveCurProgressForm('Suppression de l''analytique...');
    ExecuteSQL('DELETE ANALYTIQ WHERE Y_AXE = "A' + lStAxeTva + '" ' ) ;
    MoveCurProgressForm('Suppression des sections de TVA...') ;
    ExecuteSQL('DELETE SECTION WHERE S_AXE = "A' + lStAxeTva + '"') ;
    MoveCurProgressForm('Traitement des écritures...');
    ExecuteSQL('UPDATE ECRITURE SET  E_ANA="-" WHERE E_ANA="X" AND NOT EXISTS ( SELECT Y_NUMLIGNE FROM ' +
               'ANALYTIQ WHERE E_JOURNAL=Y_JOURNAL AND E_EXERCICE=Y_EXERCICE ' +
               'AND E_NUMEROPIECE=Y_NUMEROPIECE AND E_QUALIFPIECE=Y_QUALIFPIECE AND E_NUMLIGNE=Y_NUMLIGNE ' +
               'AND E_DATECOMPTABLE=Y_DATECOMPTABLE ) ' ) ;
    // ajout me E_IO
    ExecuteSQL('UPDATE  ECRITURE SET E_IO = "X"');
    ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE="-" WHERE G_VENTILABLE1="-" AND G_VENTILABLE2="-" AND G_VENTILABLE3="-" AND G_VENTILABLE4="-" AND G_VENTILABLE5="-" ') ;
    SetParamSoc('SO_CPPCLSAISIETVA', false) ;
    FiniMoveProgressForm ;
    CommitTrans ;
    AvertirCacheServer('') ;   // LG 01/062007 - pour forcer la prise en compte en eAGL
    result := true ;
{$IFDEF COMPTA}
  // fiche 19491
    MAJHistoparam ('GENERAUX',  '');
    MAJHistoparam ('SECTION',  '');
{$ENDIF}
    except
     on E : Exception do
      begin
       PGIError('Erreur à la désactivation de la tva' + #10#13 + E.Message ) ;
       Rollback ;
      end ;
    end ;


   end ;

end ;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
procedure DesactivationAnalytique;
var
  Q1: Tquery;
  Lst: array of Lstcpte;
  Count: integer;
begin
  // ajout me
  SetParamSoc('SO_ZSAISIEANAL', FALSE);
  SetParamSoc('SO_ZGEREANAL', FALSE);
  SetParamSoc('SO_CPPCLSANSANA', True);

  Q1 := Opensql('SELECT * From GENERAUX', TRUE,-1,'',true);
  if not Q1.EOF then
    ExecuteSQL('UPDATE  GENERAUX SET G_VENTILABLE="-",G_VENTILABLE1="-",' +
      'G_VENTILABLE2="-",G_VENTILABLE3="-",G_VENTILABLE4="-",G_VENTILABLE5="-" ');

  Ferme(Q1);
  Q1 := Opensql('SELECT * From IMMO', TRUE,-1,'',true);
  if not Q1.EOF then
    ExecuteSQL('UPDATE  IMMO SET I_VENTILABLE="-",I_VENTILABLE1="-",' +
      'I_VENTILABLE2="-",I_VENTILABLE3="-",I_VENTILABLE4="-",I_VENTILABLE5="-"');
  Ferme(Q1);
  SetLength(Lst, 5);
  Count := 0;
  Q1 := Opensql('SELECT * From AXE', TRUE,-1,'',true);
  while not Q1.EOF do
  begin
    Lst[Count].compte := Q1.FindField('X_SECTIONATTENTE').asstring;
    inc(Count);
    Q1.next;
  end;
  Ferme(Q1);
  // suppression des sections
  Q1 := Opensql('SELECT * From SECTION', TRUE,-1,'',true);
  if not Q1.EOF then
    ExecuteSQL('DELETE FROM SECTION WHERE S_SECTION<>"' +
      Lst[0].compte + '"' + ' and S_SECTION<>"' +
      Lst[1].compte + '"' + ' and S_SECTION<>"' +
      Lst[2].compte + '"' + ' and S_SECTION<>"' +
      Lst[3].compte + '"' + ' and S_SECTION<>"' +
      Lst[4].compte + '"');
  Ferme(Q1);

  // création section d'attente si aucune section
  CreationSectionAttente;

  Q1 := Opensql('SELECT * From ANALYTIQ', TRUE,-1,'',true);
  if not Q1.EOF then
    ExecuteSQL('DELETE FROM ANALYTIQ');
  Ferme(Q1);

  Q1 := Opensql('SELECT * From ECRITURE', TRUE,-1,'',true);
  if not Q1.EOF then
    ExecuteSQL('UPDATE  ECRITURE SET E_IO = "X",E_ANA="-"');
  Ferme(Q1);

  // ajout me 30/04/01
  Q1 := Opensql('SELECT * From ANAGUI', TRUE,-1,'',true);
  if not Q1.EOF then
    ExecuteSQL('DELETE FROM ANAGUI');
  Ferme(Q1);

  Q1 := Opensql('SELECT * From VENTIL', TRUE,-1,'',true);
  if not Q1.EOF then
    ExecuteSQL('DELETE FROM VENTIL');
  Ferme(Q1);

{$IFDEF COMPTA}
  // fiche 19491
    MAJHistoparam ('GENERAUX',  '');
    MAJHistoparam ('SECTION',  '');
{$ENDIF}
  
{$IFDEF CCSTD}

{$ELSE}
  {$IFNDEF EAGLSERVER}
    //Lancement de la recalcul
    MajTotTousComptes(False, '');
  {$ENDIF}
{$ENDIF}  

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure CreateCodTXX;
var lQuery : TQuery;
begin
  lQuery := OpenSQl('SELECT DISTINCT TV_TVAOUTPF, TV_CODETAUX FROM TXCPTTVA WHERE ' +
                    'TV_CODETAUX NOT IN (SELECT CC_CODE FROM CHOIXCOD WHERE ' +
                    '(CC_TYPE = "TX1") OR (CC_TYPE = "TX2"))', True,-1,'',true);
  while not lQuery.Eof do
  begin
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE, CC_CODE, CC_LIBELLE, CC_ABREGE, CC_LIBRE) ' +
               'VALUES ("'+ lQuery.FindField('TV_TVAOUTPF').AsString + '","' +
               lQuery.FindField('TV_CODETAUX').AsString + '","' +
               lQuery.FindField('TV_CODETAUX').AsString + '", "", "")');
    lQuery.Next;
  end;

  Ferme(lQuery);
end;

{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 13/03/2007
Modifié le ... : 14/03/2007
Description .. : Fonction de mise à jour des sections agricoles en fonction
Suite ........ : des choix effectués dans le dossier permanent.
Suite ........ : Paramètres :
Suite ........ : - stNoDossier : Numero de dossier
Suite ........ : - PO : Procédure of Object pour la gestion des messages
Suite ........ : d'erreur
Suite ........ : Code retour
Suite ........ : - renvoie 0 si OK sinon renvoie le code erreur (<0)
Mots clefs ... :
*****************************************************************}
function MiseAJourAxeAgricole (stNoDossier : string; PO : TInfoMajAgricole) : integer;
  procedure _MajUO (pstNodossier, pstAxe : string);
  var
    stSQL : string;
  begin
    stSQL := 'UPDATE DB'+stNoDossier+'.dbo.SECTION SET S_UO=(SELECT SUM(DPC_SUPERFICIE) FROM DPPARCELLES '+
              ' WHERE DPC_NODOSSIER="'+pstNoDossier+'" AND DPC_CULTURE=S_SECTION) '+
              ' WHERE S_AXE="'+pstAxe+'"';
    ExecuteSQL ( stSQL );
  end;

  procedure _CreationSectionDansDossier (pstSQL : string; pstNoDossier : string; pstAxe : string );
  var
    Q : TQuery;
    i : integer;
    TSection : TOB;
  begin
    // Chargement des sections
    TSection := TOB.Create ('',nil,-1);
    Q := OpenSQL (pstSQL, True,-1,'',True);
    try
      if not Q.Eof then
      begin
        while not Q.Eof do
        begin
          with TOB.Create ('SECTION',TSection,-1) do
          begin
            PutValue('S_SECTION',Q.FindField('YDS_CODE').AsString);
            PutValue('S_LIBELLE',Q.FindField('YDS_LIBELLE').AsString);
            PutValue('S_ABREGE',Copy(Q.FindField('YDS_LIBELLE').AsString,1,17));
            PutValue('S_AXE',pstAxe);
            PutValue('S_SENS','M');
            PutValue('S_CREERPAR','BUR');
            // --- CEGID V9
            PutValue('S_INVISIBLE','-');
          end;
        Q.Next;
        end;
      end;
    finally
      Ferme (Q);
    end;
    // Enregistrement des sections dans la base dossier
    if TSection.Detail.Count <> 0 then
    begin
      for i := 0 to TSection.Detail.Count - 1 do
        ExecuteSQL (StringReplace (TSection.Detail[i].MakeInsertSql,'INTO SECTION','INTO DB'+pstNoDossier+'.dbo.SECTION',[RfIgnoreCase]));
    end;
  end;

var
  stSQL : string;
  stAxe : string;
  TSection : TOB;
  i : integer;
begin
  result := 0;

  // Si Quelqu'un sur le dossier, on refuse le traitement
  if (stNoDossier <> NoDossierBaseCommune) then
  begin
    if ExisteSQL('SELECT * FROM DB'+stNoDossier+'.dbo.COURRIER WHERE MG_TYPE=-666') then
    begin
      if assigned (PO) then PO(-1,TraduireMemoire('Un utilisateur est présent sur le dossier')+
        ' '+stNoDossier+'. '+TraduireMemoire('Veuillez-vous déconnecter.'));
      result := -1;
      exit;
    end;
  end;

  {
     Si l'axe agricole n'est pas renseigné, on affecte l'axe destiné à recevoir
    les sections agricoles.
  }
  BeginTrans;
  try
    if (GalOutil.GetParamSocDossier('SO_CPAXEAGRICOLE','DB'+StNoDossier)='') then
    begin

      { 1 - Calcul et mise à jour de l'axe dans les paramètres société }

      for i := 0 to 5 do
      begin
        if not ExisteSQL('SELECT G_GENERAL FROM DB'+stNoDossier+'.dbo.GENERAUX WHERE ' +
                       'G_VENTILABLE' + IntToStr(i+1) + ' = "X"') then
        begin
          stAxe := 'A' + IntToStr(i+1);
          ExecuteSQL('UPDATE DB'+stNoDossier+'.dbo.PARAMSOC SET SOC_DATA="'+stAxe+'" WHERE SOC_NOM="SO_CPAXEAGRICOLE"');
          Break;
        end;
      end;

      { 2 - Activation de l'analytique sur le dossier }
      if GalOutil.GetParamsocDossier('SO_CPPCLSANSANA','DB'+StNoDossier) then
      begin
        ExecuteSQL('UPDATE DB'+stNoDossier+'.dbo.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_ZSAISIEANAL"');
        ExecuteSQL('UPDATE DB'+stNoDossier+'.dbo.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_ZGEREANAL"');
        ExecuteSQL('UPDATE DB'+stNoDossier+'.dbo.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_CPPCLSANSANA"');
      end;

      if (stAxe <> '') then
      begin
        { 3 - Mise à jour de l'axe agricole dans la table AXE }
        ExecuteSQL('UPDATE DB'+stNoDossier+'.dbo.AXE SET X_LIBELLE = "Agricole", X_ABREGE = "Agricole", X_SECTIONATTENTE = "000000", ' +
                     'X_STRUCTURE = "-", X_GENEATTENTE = "", X_LONGSECTION = 6 WHERE ' +
                     'X_AXE = "'+stAxe + '"');
        { Pour être tranquille, on supprime les sections éventuellement créées auparavant }
        ExecuteSQL ('DELETE FROM DB'+stNoDossier+'.dbo.SECTION WHERE S_AXE="'+stAxe+'"');

        { 4 - Création de la section d'attente }
        TSection := TOB.Create ('SECTION',nil,-1);
        try
          TSection.PutValue('S_SECTION','000000');
          TSection.PutValue('S_LIBELLE',TraduireMemoire('Section d''attente'));
          TSection.PutValue('S_ABREGE',TraduireMemoire('Attente'));
          TSection.PutValue('S_AXE',stAxe);
          TSection.PutValue('S_SENS','M');
          TSection.PutValue('S_CREERPAR','BUR');
          // --- CEGID V9
          TSection.PutValue('S_INVISIBLE','-');
          ExecuteSQL (StringReplace (TSection.MakeInsertSql,'INTO SECTION','INTO DB'+stNoDossier+'.dbo.SECTION',[RfIgnoreCase]));
        finally
          TSection.Free;
        end;

        { 5 - Création des sections associées en fonction des éléments du DP}
        { CA/CAT - 23/04/2007 - A la demande de CJ on ne prend pas en compte les sections 600 }
        stSQL := 'SELECT * FROM CHOIXDPSTD WHERE YDS_TYPE="DAG" AND YDS_PREDEFINI="DOS" AND YDS_NODOSSIER="'+stNoDossier+'" AND (YDS_CODE NOT LIKE "600%")';
        _CreationSectionDansDossier (stSQL, stNoDossier, stAxe);

        { 6 - Mise à jour des unités d'oeuvre }
        _MajUO (stNoDossier, stAxe);
      end else
      begin
        if assigned (PO) then PO(-2,TraduireMemoire('Affectation de l''axe agricole impossible.'));
        result := -2;
        exit;
      end;
    end else
    {
      Alignement des sections agricoles soit :
     - Création des nouvelles sections
     - Fermeture des sections plus utilisées.
     }
    begin
      stAxe := GalOutil.GetParamSocDossier('SO_CPAXEAGRICOLE','DB'+stNoDossier);
      if stAxe <> '' then
      begin
        // Ajout des nouvelles sections
        // CA/CAT - 23/04/2007 - A la demande de CJ on ne prend pas en compte les sections 600 }
        stSQL := 'SELECT * FROM CHOIXDPSTD WHERE YDS_TYPE="DAG" AND YDS_NODOSSIER="'+stNoDossier+
              '" AND (YDS_CODE NOT LIKE "600%") AND YDS_CODE NOT IN (SELECT S_SECTION FROM DB'+stNoDossier+
              '.dbo.SECTION WHERE S_AXE="'+stAxe+'")';
        _CreationSectionDansDossier (stSQL, stNoDossier,stAxe);
        // Fermeture des sections qui ne sont plus utilisés
        stSQL := 'UPDATE DB'+stNoDossier+'.dbo.SECTION SET S_FERME="X" WHERE S_AXE="'+stAxe+'" AND S_SECTION<>"000000" ';
        stSQL := stSQL + ' AND S_SECTION NOT IN (SELECT YDS_CODE FROM CHOIXDPSTD WHERE YDS_TYPE="DAG" AND YDS_NODOSSIER="'+stNoDossier+'")';
        ExecuteSQL (stSQL);
        // Réouverture des section réactivées
        stSQL := 'UPDATE DB'+stNoDossier+'.dbo.SECTION SET S_FERME="-" WHERE S_AXE="'+stAxe+'" AND S_SECTION<>"000000" ';
        stSQL := stSQL + ' AND S_SECTION IN (SELECT YDS_CODE FROM CHOIXDPSTD WHERE YDS_TYPE="DAG" AND YDS_NODOSSIER="'+stNoDossier+'")';
        ExecuteSQL (stSQL);
        { Mise à jour des unités d'oeuvre }
        _MajUO (stNoDossier, stAxe);
      end;
    end;
    AvertirCacheServer( 'PARAMSOC' ) ;
    CommitTrans;
  except
    on E : Exception do
    begin
      if assigned(PO) then PO (-3,TraduireMemoire('Erreur SQL : ') + E.Message);
      RollBack;
      result := -3;
    end;
  end;
end;
{$ENDIF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/05/2007
Modifié le ... : 24/05/2007
Description .. : Enregistre le dossier comme standard
Suite ........ : Paramètres :
Suite ........ : - pNumStd : integer : numéro de standard de destination
Suite ........ : - pstParamStd : string : Flags d'enregistrement (X ou - )
Suite ........ :   [1] : PARAMSTD : enregistrement des PARAMSOCs
Suite ........ :   [2] : COMPTESTD : enregistrement des généraux
Suite ........ :   [3] : AUXILIAIRESTD : enregistrement des auxiliaires
Suite ........ :   [4] : JOURNALSTD : enregistrement des journaux
Suite ........ :   [5] : GUIDESTD : enregistrement des guides
Suite ........ : - pstLibelleStd : string : libellé du standard
Mots clefs ... :
*****************************************************************}
function SaveDossierAsStandard (pNumStd : Integer; pstParamStd : string; pstLibelleStd : string) : integer ;

  procedure _SaveInfoStandard (pNumStd : integer; pstLibelleStd : string);
  var  TRef : TOB;
  begin
    if not ExisteSQL ('SELECT 1 FROM STDCPTA WHERE STC_NUMPLAN='+IntToStr(pNumStd)) then
    begin
      TRef := TOB.Create('STDCPTA',nil,-1);
      try
        TRef.PutValue('STC_NUMPLAN', pNumStd);
        TRef.PutValue('STC_LIBELLE', pstLibelleStd);
        TRef.PutValue('STC_ABREGE', Copy(pstLibelleStd,1,17));
        TRef.PutValue('STC_PREDEFINI','STD');
        TRef.InsertDB(nil);
      finally
        TRef.Free;
      end;
    end;
  end;

begin
  Result := -1;
  if (Length(pstParamStd) <> 5 ) then exit;
  // Enregistrement des paramètres dossier
  try
    BeginTrans;
    if (pstParamStd[1] = 'X') then EnregistreParamSocStandard(pNumStd);
    // Enregistrement du plan comptable
    if (pstParamStd[2] = 'X') then
    begin
      SaveAsStandardCompta(pNumStd, 'GENERAUX','GENERAUXREF');
      { Remise à zéro des cumuls }
      ExecuteSQl ('Update GENERAUXREF set GER_TOTDEBP=0,GER_TOTCREP=0,GER_TOTDEBE=0,GER_TOTCREE=0,'+
         'GER_TOTDEBS=0,GER_TOTCRES=0 Where GER_NUMPLAN='+IntToStr(pNumStd));
    end;
    // Enregistrement des tiers
    if (pstParamStd[3] = 'X') then SaveAsStandardCompta(pNumStd, 'TIERS','TIERSREF');
    // Enregistrement des journaux
    if (pstParamStd[4] = 'X')  then SaveAsStandardCompta(pNumStd, 'JOURNAL','JALREF');
    // Enregistrement des guides
    if (pstParamStd[5] = 'X') then
    begin
      SaveAsStandardCompta(pNumStd, 'GUIDE','GUIDEREF');
      SaveAsStandardCompta(pNumStd, 'ECRGUI','ECRGUIREF');
      SaveAsStandardCompta(pNumStd, 'ANAGUI','ANAGUIREF');
    end;
    _SaveInfoStandard (pNumStd,pstLibelleStd);
    CommitTrans;
    Result := 0;
  except
    Rollback;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 26/06/2007
Modifié le ... : 26/06/2007
Description .. : Chargement d'une table depus un standard
Mots clefs ... :
*****************************************************************}
function YFSChargeTableDepuisStandard ( pstTable : string; NumStd : integer; pstBaseName : string) : boolean;
var
  lstFileName             : string;
  lstPredefini            : string;
  lRet                    : integer;
  lTStd                   : TOB;
  lbEntete                : boolean;
  lstEncoding             : string;
  i                       : integer;
begin
  if NumStd > 20 then lstPredefini := 'STD' else lstPredefini :='CEG';
  { Extraction des enregistrements vers le disque }
  lstFileName := AGL_YFILESTD_GET_PATH ('COMPTA', pstTable+Format('%.03d',[NumStd])+'.'+YFS_EXTENSION,
    'STANDARD','DOSSIERTYPE', Format('%.03d',[NumStd]), pstTable, '', V_PGI.LanguePrinc, lStPredefini);
  DeleteFile (lstFileName);
  // Récupération du fichier en base
  lRet := AGL_YFILESTD_EXTRACT( lstFileName, 'COMPTA',pstTable+Format('%.03d',[NumStd])+'.'+YFS_EXTENSION,
    'STANDARD','DOSSIERTYPE', Format('%.03d',[NumStd]), pstTable,'',False,V_PGI.LanguePrinc,lstPredefini);
  if (lRet = -1) then
  begin
    lTStd := TOB.Create ('',nil,-1);
    try
      lTStd.LoadFromXMLFile(lstFileName,lbEntete,lstEncoding);
      //lTStd.LoadDetailFile(lstFileName);
      if (pstBaseName = '' ) then lTStd.InsertOrUpdateDB()  // si on est connecté à la base dossier
      else
      begin // Si on n'est pas connecté à la base dossier
        // Suppression des enregistrements en base
        ExecuteSQL ('DELETE FROM '+pstTable);
        for i := 0 to lTStd.Detail.Count - 1 do
          ExecuteSQl (StringReplace (lTStd.Detail[i].MakeInsertSql,
              'INTO '+pstTable,'INTO '+pstBaseName+'.dbo.'+pstTable,[RfIgnoreCase]));
      end;
    finally
      lTStd.Free;
    end;
  end;
  Result := (lRet = - 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 26/06/2007
Modifié le ... : 26/06/2007
Description .. : Enregistrement d'une table comme un standard dans la
Suite ........ : base de données (table YFILESTD)
Mots clefs ... :
*****************************************************************}
function YFSEnregistreTableVersStandard ( pstTable : string; NumStd : integer) : boolean;
var
  lTStd         : TOB;
  lstFileName   : string;
  lRet          : integer;
  lbExit        : boolean;
begin
  { Chargement des enregistrements de la table 'pstTable' en TOB }
  lTStd := TOB.Create ('', nil, -1);
  try
    lTStd.LoadDetailDBFromSQL (pstTable,'SELECT * FROM '+pstTable);
    lbExit := (lTStd.Detail.Count = 0);
    if not lbExit then
    begin
      { Enregistrement de la TOB dans un fichier }
      lstFileName := pstTable+Format('%.03d',[NumStd])+'.'+YFS_EXTENSION;
      lTStd.SaveToXmlFile(TCBPPath.GetTemplates+'\'+lstFileName,False);
      //lTStd.SaveToFile(lstFileName,False,False,True);
    end;
  finally
    lTStd.Free;
  end;

  if not lbExit then
  begin
    { Insertion du fichier dans la base de données - Table YFILESTD }
    lRet := AGL_YFILESTD_IMPORT(TCBPPath.GetTemplates+'\'+lstFileName,'COMPTA',ExtractFileName(lstFileName), YFS_EXTENSION,
                  'STANDARD',
                  'DOSSIERTYPE',
                  Format('%.03d',[NumStd]),
                  pstTable,'','-','-','-','-','-',
                  V_PGI.LanguePrinc,
                  'STD',
                  TraduireMemoire('Standard')+' '+pstTable+' '+TraduireMemoire('n°')+IntToStr(NumStd),'000000');
  end else lRet := -1;
  Result := (lRet = -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 26/06/2007
Modifié le ... :   /  /
Description .. : Suppression des éléments d'un standard pour une table
Suite ........ : donnée
Mots clefs ... :
*****************************************************************}
function YFSSupprimeTableStandard ( pstTable : string; NumStd : integer) : boolean;
var
  lstPredefini : string;
begin
  if NumStd > 20 then lstPredefini := 'STD'
  else lstPredefini := 'CEG';
  Result :=
    (Agl_YFileStd_Delete ('COMPTA', pstTable+Format('%.03d',[NumStd])+'.'+YFS_EXTENSION,V_PGI.LanguePrinc,
      lstPredefini, 'STANDARD', '000000') = -1);
end;

end.



