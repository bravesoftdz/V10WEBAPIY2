unit uWAINIBase;
{$IFDEF BASEEXT}azertyuiop{$ENDIF BASEEXT}

interface


uses hctrls,
  sysutils,
  hent1,
  classes,
  uWAINI,
  uWAInfoIni,
  uWAEnvironnement,
  utob
  ,
  Variants
{$IFNDEF EAGLCLIENT}
  , DB
{$IFDEF ODBCDAC}
  ,
  odbcconnection,
  odbctable,
  odbcquery,
  odbcdac
{$ELSE}
,
  uDbxDataSet
{$ENDIF}
  ,
  seriaPGI
{$ENDIF EAGLCLIENT}
  ;

type


  {  cWAINIBase Class }
  cWAINIBase = class(cWAINI)
  private
{$IFNDEF EAGLCLIENT}
    xxDB: TDataBase;
    DBDriver: TDBDriver;
    function getDB: TDataBase;
    function getDB2(NomBase,NomIni:string): TDataBase;
    property leDB: TDataBase read getDB;
    function getNouveauDossier(): string;
    function SaveLaTob(laTob: tob): boolean;
    function getSocieteFromHeberge(laTob: TOB; AvecRecupDescriptionBase: boolean): boolean;
    function chargeSQL2TOB(nomTob, sql: string): TOB;
    function myExecuteSQL(sSQL: string; NoChangeSQL: boolean = FALSE): integer;
    function myExisteSQL(sSQL: string): boolean;
    procedure setEntiteFromNomEntite;
    procedure setGroupeFromNomGroupe;
    function CtrlHeberge(laTob: TOB): boolean;
    function CtrlClient(laTob: TOB): boolean;
    function CtrlDossier(laTob: TOB): boolean;
    function GetWhereCle(NomTable: string; laTob: tob; var ChampsAIgnorer: string): string;
{$ENDIF EAGLCLIENT}
    function getISCreate: boolean;
    procedure setIsCreate(valeur: boolean);
  public
    constructor create(); override;
    destructor destroy; override;

    {$IFNDEF EAGLCLIENT}
    function Chargedossier(VireRef, VireDP, VireModel: boolean): HTStringList; override;
    function Ajoutedossier(WAInfoIni: cWAInfoIni): boolean; override;
    function Supprimedossier(WAInfoIni: cWAInfoIni): boolean; override;
    function Modifiedossier(WAInfoIni: cWAInfoIni): boolean; override;
    function Existedossier(WAInfoIni: cWAInfoIni): boolean; override;
    procedure ChargeDBParams(WAInfoIni: cWAInfoIni); override;
    function GetSocRef(): string; override;
    function SetDossierVerrou(BaseNom: string; Etat: TDossierVerrou): boolean; override ;
    function GetDossierVerrou(BaseNom: string): TDossierVerrou; override ;
    function getListOfHeberge(): HTStringList; override ;
    function getListOfClient(): HTStringList; override ;
    function getListOfDossier(): HTStringList; override ;
    function UserPassValidate(user, pass: string): boolean;override ;
    function IsAdmin(login: string): boolean;override ;
    function GetHeberge: TOB;override ;
    function GetClient: TOB;override ;
    function GetDossier: TOB;override ;
    function setHeberge(laTob: TOB): boolean;override ;
    function setClient(laTob: TOB): boolean;override ;
    function setDossier(laTob: TOB; BaseFrom: string; AvecRecupHeberge: boolean): boolean;override ;
    function DelHeberge: boolean;override ;
    function DelClient: boolean;override ;
    function DelDossier(AvecSupprPhysique, ForceSupprSiEchecPhysique: boolean): boolean;override ;
    function GetNextID(leQuel: integer): string;override ;
    function getProduit: string; override;
    function GetMaxUserNumber: integer; override;
    function GetEnvironnement(WAEnvironnement : cWAEnvironnement): boolean; override ;
    function IsFicBase(): boolean; override;
    function DeleteTableToBase(Base,Tables:String) : boolean;
    procedure DeleteSpecJournal(xxDBSource:TDataBase);
    procedure DeleteSpecSouche(xxDBSource:TDataBase);
    procedure DeleteSpecUTILISAT(xxDBSource:TDataBase);
    procedure DeleteSpecFiles(xxDBSource:TDataBase);
    procedure InitModeEchange(xxDBSource:TDataBase);  // MKO 25/06/09
    procedure InitTVA(xxDBSource:TDataBase);  // MKO 29/06/09

    {$ENDIF EAGLCLIENT}

    function getNomBase((*wGamme,*) wEntite, wDossier: string): string;
    property isCreate: boolean read getISCreate write setIsCreate; // Flag qui permet de specifier à setHeberge qu'il est en création ;

  public
    class function TokenCar(): string;
    procedure getElementFromNom(var nom, stEntite, stGroupe: string); overload; //, stGamme: string); overload;
    procedure getElementFromNom(var nom: string); overload;
    procedure getElementFromNom(WAInfoIni: cWAInfoIni); overload;
{$IFNDEF EAGLCLIENT}
  private
    procedure initHeberge(T1: TOB);
  private
    function chercheUserSuivant(var User: string): boolean;
  public
    function GetListofEnvironnement() : HTStringList; override ;
{$ENDIF EAGLCLEINT}
  end;

const LISTETABLE = 'BUDGET;PREVISION;VIREMENTINTERNE;CALENDRIER;CATEGORIEBNC;COMPTE;'+
                   'ECRITUREBNC;ECRITUREBNCD;ANALYTIQ;CCNFCONTROLE;CNOTESTRAVAIL;'+
                   'CNTVDETAIL;CONTABON;CONTACT;DOCREGLE;ECRITURE;'+
                   'ECRVENTILTVA;EEXBQ;EEXBQIMPUT;EEXBQLIG;JOURNAL;REGLEIMPUTBQ;'+
                   'SOUCHE;TIERS;VENTILREGLEIMP;ADRLIVR;ARTICLE;COMMERC;COMPOART;'+
                   'DISPO;GLOSSAIR;LIBELLEETRANGER;LIENS;LIGNES;MOUVEMENTSSTOCK;PIECES;'+
                   'PIEDECHE;TARIF;TAUXTVAARTICLE;IMMO;IMMOAMOR;IMMOARD;IMMOCPTE;'+
                   'IMMODPI;IMMOECHE;IMMOEXERCICE;IMMOLOG;IMMOMVTD;IMMOREGFR;IMMOUO;'+
                   'ABSENCESALARIE;DEPORTSAL;DROITACCES;ETABCOMPL;HISTOSAISRUB;'+
                   'MASQUESAISRUB;MEMOSAL;MOTIFABSENCE;MOTIFENTREESAL;MOTIFSORTIEPAY;'+
                   'PROFILRUB;RECAPSALARIES;REMUNERATION;SALARIES;ACTIONS;CIBLE;UTILISAT;'+
                   'OPERATION;RELEVESFACT;SESSIONLOG;FERMECAISSE;REGLEDETAIL;TEMPETIQ;'+
                   'TEMPFONDS;AGENDA;ETABLISS;LOG;RIB;YFILEPARTS;YFILES;YIMAGES';
                 //  DETAILGROUPE;   enlever pour assoc

implementation

uses paramsoc,
  uWAIniFile,
  majtable,
  licUtil,
  licJFD,
  Forms,
  uWABD,
  uWABDSQLSERVER  ,
  cbpdates
{$IFDEF EAGLSERVER}
  , eSession
{$ENDIF EAGLSERVER}
  , cbpDatabases
  ;

const
  MAX_LEN_HEBERGE = 17;
  MAX_LEN_CLIENT = 17;
  MAX_LEN_DOSSIER = 17;

  /////////////////////////////////////////////////////////////////////////////////////////////
function nString(n: integer; car: string): string;
//<=> String VB
var
  StOut: string;
  i: Word;
begin
  if n <= 0 then
  begin
    result := '';
    exit;
  end;
  StOut := '';
  for i := 1 to n do
    StOut := StOut + car;
  Result := StOut;
end;

//////////////////////////////////////////////////////////////////////////////
function BourreGauche(St: string; n: integer; car: string): string;
var
  w: string;
  i: integer;
begin
  i := n - length(st);
  if i < 0 then
    result := st;
  w := nString(i, car);
  result := w + st;
end;

{  cWAINIBase }

const
  fTokenCar = ';';

class function cWAINIBase.TokenCar(): string;
begin
  result := fTokenCar;
end;

procedure cWAINIBase.getElementFromNom(var nom, stEntite, stGroupe: string) ; //, stGamme: string);
var
  st: string;
begin
  st := nom;
  if (pos(fTokenCar, nom) > 0) then
  begin
    nom := READTOKENPipe(st, fTokenCar);
    stEntite := READTOKENPipe(st, fTokenCar);
    stGroupe := READTOKENPipe(st, fTokenCar);
    //stGamme := st;
  end
  else
  if (pos('@', nom) > 0) then
  begin
    nom := READTOKENPipe(st, '@');
    stEntite := READTOKENPipe(st, '@');
    stGroupe := READTOKENPipe(st, '@');
    //stGamme := st;
  end
  else
  begin //YCP 05/04/05
    nom := st;
    stEntite := IDHeberge;
    stGroupe := IDClient;
    //stGamme := Gamme;
  end;
end;

procedure cWAINIBase.getElementFromNom(var nom: string);
var
  stEntite, stGroupe(*, stGamme*): string;
begin
  getElementFromNom(nom, stEntite, stGroupe) ; //, stGamme);
  IDHeberge := stEntite;
  IDClient := stGroupe;
  //Gamme := stGamme;
end;

procedure cWAINIBase.getElementFromNom(WAInfoIni: cWAInfoIni);
var
  st: string;
begin
  st := WAInfoIni.NomDeBase;
  getElementFromNom(st);
  WAInfoIni.NomDeBase := st;
end;

{$IFNDEF EAGLCLIENT}
function cWAINIBase.IsAdmin(login: string): boolean;
begin
  if ModeEagl then
  begin
    result:=inherited IsAdmin(login) ;
  end
  else
  begin
  result := false;
{$IFNDEF EAGLCLIENT}
  if not result then
    result := myExisteSQL('SELECT * FROM ADMHEBERGE WHERE AHE_IDHEBERGE="' + login + '" AND AHE_CEGID="X"');
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.myExecuteSQL(sSQL: string; NoChangeSQL: boolean = FALSE): integer;
var
  OldDriver: TDBDriver;
begin
  traceexecution(sSql);
  OldDriver := V_PGI.Driver;
{$IFNDEF EAGLSERVER}
  V_PGI.Driver := DBDriver;
{$ELSE}
  LookupCurrentSession.LeDBDriver := DBDriver;
{$ENDIF eAglServer}
  try
  result := ExecuteSQLDB(sSQL, leDB, NoChangeSQL);
  finally
{$IFNDEF EAGLSERVER}
  V_PGI.Driver := OldDriver;
{$ELSE}
  LookupCurrentSession.LeDBDriver := OldDriver;
{$ENDIF eAglServer}
    end ;
end;

function cWAINIBase.myExisteSQL(sSQL: string): boolean;
var
  q: TQuery;
begin
  traceexecution(sSql);
  q := DBOpenSQL(leDB, sSql, DBDriver);
  result := not q.eof;
  ferme(q);
end;

procedure cWAINIBase.setEntiteFromNomEntite;
var
  q: Tquery;
  sql: string;
begin
  q := nil;
  try
    if (IDHeberge = '') and (NomHeberge <> '') then
    begin
      //recherche sur l'id de l'entite
      sql := 'SELECT AHE_IDHEBERGE FROM ADMHEBERGE WHERE AHE_IDHEBERGE="' + NomHeberge + '" AND (AHE_FERME<>"X" OR AHE_FERME IS NULL)';
      traceExecution(Sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      if not q.eof then
        IDHeberge := q.Fields[0].AsString
      else
        IDHeberge := NomHeberge; //YCP 11/04/05
      ferme(Q);
      //recherche sur le nom de l'entite
      (*if (Entite='') then
        begin
        sql:='SELECT AHE_IDHEBERGE FROM ADMHEBERGE WHERE AHE_NOM="'+Nomentite+'"' ;
        traceExecution(Sql) ;
        q:=DBOpenSQL(leDB,Sql,DBDriver) ;
        if not q.eof then Entite:=q.Fields[0].AsString ;
        ferme(Q) ;
        end ;*)
    end;
  except
    on E: Exception do
    begin
      if Assigned(q) then
        ferme(Q);
      ErrorMessage := 'setEntiteFromNomEntite: ' + E.Message;
    end;
  end;
end;

procedure cWAINIBase.setGroupeFromNomGroupe;
var
  q: Tquery;
  sql: string;
begin
  q := nil;
  try
    if (IDClient = '') and (NomClient <> '') then
    begin
      //recherche sur l'id du groupe
      sql := 'SELECT ACL_IDCLIENT FROM ADMCLIENT WHERE ACL_IDCLIENT="' + NomClient + '" AND ACL_IDHEBERGE="' + IDHeberge + '" AND (ACL_FERME<>"X" OR ACL_FERME IS NULL)';
      traceExecution(Sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      if not q.eof then
        IDClient := q.Fields[0].AsString
      else
        IDClient := NomClient; //YCP 11/04/05
      ferme(Q);
      //recherche sur le nom du groupe
      (*if (groupe='') then
        begin
        sql:='SELECT ACL_IDCLIENT FROM ADMCLIENT WHERE ACL_NOM="'+NomGroupe+'" AND ACL_IDHEBERGE="'+Entite+'"' ;
        traceExecution(Sql) ;
        q:=DBOpenSQL(leDB,Sql,DBDriver) ;
        if not q.eof then Groupe:=q.Fields[0].AsString ;
        ferme(Q) ;
        end ;*)
    end;
  except
    on E: Exception do
    begin
      if Assigned(q) then
        ferme(Q);
      ErrorMessage := 'setGroupeFromNomGroupe: ' + E.Message;
    end;
  end;
end;
{$ENDIF EAGLCLIENT}

//var cpte: integer = 0; //YCP 15/05/07 

constructor cWAINIBase.create();
begin
  inherited;
  ReaffecteNomTable(className, true);
  mode := 'BAS';
  forceBase := true;
{$IFNDEF EAGLCLIENT}
  xxDB := nil;
{$ENDIF EAGLCLIENT}
end;

destructor cWAINIBase.destroy;
begin
{$IFNDEF EAGLCLIENT}
  TCBPDatabases.FreeDataBase(xxDB) ;
{$ENDIF EAGLCLIENT}
  inherited;
end;

{$IFNDEF EAGLCLIENT}
function cWAINIBase.Chargedossier(VireRef, VireDP, VireModel: boolean): HTStringList;
var
  q: TQuery;
  sRef, where, sql: string;
begin
  if ModeEagl then
  begin
    result := inherited Chargedossier(Vireref, VireDP, VireModel);
  end
  else
  begin
    result := HTStringList.Create;
    afficheTrace(self);
    q := nil;
    try
      setEntiteFromNomEntite;
      setGroupeFromNomGroupe;

      sRef := GetSocRef();
      sql := 'SELECT ASO_NOM,ASO_IDCLIENT FROM ADMSOCIETE WHERE (ASO_FERME<>"X" OR ASO_FERME IS NULL) AND (ASO_VERROU="AUC" OR ASO_VERROU="")';
      where := ' AND ASO_IDHEBERGE="' + IDHeberge + '"';
      if (IDClient <> '') then
        where := where + ' AND ASO_IDCLIENT="' + IDClient + '"'; //YCP 08/03/05
      if (VireModel) then
        where := where + ' AND ASO_MODELE="-"'; //YCP 08/03/05
      //where := where + ' AND ASO_GAMME="' + gamme + '"'; //if (gamme<>'')  then
      if (server <> '') then
        where := where + ' AND ASO_SERVER="' + server + '"';
      if (Driver <> '') then
        where := where + ' AND ASO_DRIVER="' + Driver + '"';

      traceExecution(Sql + where);
      q := DBOpenSQL(leDB, Sql + where, DBDriver);
      while not Q.eof do
      begin
        if (sRef = q.fields[0].AsString) then
        begin
          q.next;
          continue;
        end;
        if VireDP and (Copy(q.fields[0].AsString, 1, 3) = '###') then
          continue;
        result.Add(q.fields[0].AsString + fTokenCar + IDHeberge + fTokenCar + q.fields[1].AsString) ; // + fTokenCar + Gamme);
        q.next;
      end;
      ferme(Q);
      if not VireRef then
      begin
        result.Add(sRef);
      end;
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'Chargedossier: ' + E.Message;
      end;
    end;
  end;
end;

function cWAINIBase.Ajoutedossier(WAInfoIni: cWAInfoIni): boolean;
{$IFNDEF EAGLCLIENT}
var
  q: TQuery;
  where, sql: string;
  EncryptUserPW: boolean;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result := inherited AjouteDossier(WAInfoIni);
  end
  else
  begin
    result := true;
{$IFNDEF EAGLCLIENT}
    try
      EncryptUserPW := true;

      (*if pos(fTokenCar ,WAInfoIni.NomDeBase)>0 then
        begin
        st:=WAInfoIni.NomDeBase ;
        WAInfoIni.NomDeBase:=READTOKENPipe(st,fTokenCar) ;
        Entite:=READTOKENPipe(st,fTokenCar) ;
        Groupe:=READTOKENPipe(st,fTokenCar) ;
        Gamme:=st ;
        end ;*)
      getElementFromNom(WAInfoIni);

      sql := 'SELECT * FROM ADMSOCIETE WHERE ASO_NOM ="' + WAInfoIni.NomDeBase + '" ';
      where := ' AND ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"' ; // AND ASO_GAMME="' + gamme + '"';
      traceexecution(sql + where);
      q := DBOpenSQL(leDB, Sql + where, DBDriver, false, false);
      if Q.eof then
      begin
        q.Insert;
        q.findfield('ASO_IDDOSSIER').AsString := getNouveauDossier();
        q.findfield('ASO_IDHEBERGE').AsString := IDHeberge;
        q.findfield('ASO_IDCLIENT').AsString := IDCLient;
        //q.findfield('ASO_GAMME').AsString := Gamme;
      end
      else
      begin
        q.Edit;
      end;
      if EncryptUserPW then
      begin
        WAInfoIni.User := CryptageSt(WAInfoIni.User);
        WAInfoIni.Pass := CryptageSt(WAInfoIni.Pass);
      end;
      q.findfield('ASO_NOM').AsString := WAInfoIni.NomDeBase;
      q.findfield('ASO_DRIVER').AsString := WAInfoIni.Driver;
      q.findfield('ASO_SERVEUR').AsString := WAInfoIni.server;
      q.findfield('ASO_PATH').AsString := WAInfoIni.Path;
      q.findfield('ASO_PILOTE').AsString := WAInfoIni.ODBC;
      q.findfield('ASO_NOMBASE').AsString := WAInfoIni.Base;
      q.findfield('ASO_LOGIN').AsString := WAInfoIni.User;
      q.findfield('ASO_PASSWORD').AsString := WAInfoIni.Pass;
      q.findfield('ASO_SHARE').AsString := WAInfoIni.Share;
      q.findfield('ASO_OPTIONS').AsString := WAInfoIni.Options ; //08/08/2007
      q.post;
      //StDir:=IniFile.ReadString(StNom,'Dir','');
      ferme(Q);
    except
      on E: Exception do
      begin
        result := false;
        ErrorMessage := 'AjouteDossier: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.Supprimedossier(WAInfoIni: cWAInfoIni): boolean;
{$IFNDEF EAGLCLIENT}
var
  where, sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result := inherited Supprimedossier(WAInfoIni);
  end
  else
  begin
    result := true;
{$IFNDEF EAGLCLIENT}
    try
      (*if pos(fTokenCar ,WAInfoIni.NomDeBase)>0 then
        begin
        st:=WAInfoIni.NomDeBase ;
        WAInfoIni.NomDeBase:=READTOKENPipe(st,fTokenCar) ;
        Entite:=READTOKENPipe(st,fTokenCar) ;
        Groupe:=READTOKENPipe(st,fTokenCar) ;
        Gamme:=st ;
        end ; *)
      getElementFromNom(WAInfoIni);

      sql := 'DELETE FROM ADMSOCIETE WHERE ASO_NOM ="' + WAInfoIni.NomDeBase + '" ';
      where := ' AND ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"' ; // AND ASO_GAMME="' + gamme + '"';
      MyExecuteSQL(sql + where);
    except
      on E: Exception do
      begin
        result := false;
        ErrorMessage := 'Supprimedossier: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.Modifiedossier(WAInfoIni: cWAInfoIni): boolean;
begin
  if ModeEagl then
  begin
    result := inherited Modifiedossier(WAInfoIni);
  end
  else
  begin
    result := Ajoutedossier(WAInfoIni);
  end;
end;

function cWAINIBase.Existedossier(WAInfoIni: cWAInfoIni): boolean;
{$IFNDEF EAGLCLIENT}
var
  q: TQuery;
  where, sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result := inherited Existedossier(WAInfoIni);
  end
  else
  begin
{$IFDEF EAGLCLIENT}
    result := false;
{$ELSE}
    q := nil;
    try
      (*if pos(fTokenCar ,WAInfoIni.NomDeBase)>0 then
        begin
        st:=WAInfoIni.NomDeBase ;
        WAInfoIni.NomDeBase:=READTOKENPipe(st,fTokenCar) ;
        Entite:=READTOKENPipe(st,fTokenCar) ;
        Groupe:=READTOKENPipe(st,fTokenCar) ;
        Gamme:=st ;
        end ;*)
      getElementFromNom(WAInfoIni);

      sql := 'SELECT ASO_NOM FROM ADMSOCIETE WHERE ASO_NOM ="' + WAInfoIni.NomDeBase + '" ';
      where := ' AND ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"' ; // AND ASO_GAMME="' + gamme + '"';
      traceexecution(sql + where);
      q := DBOpenSQL(leDB, Sql + where, DBDriver);
      result := not q.eof;
      ferme(Q);
    except
      on E: Exception do
      begin
        result := false;
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'Existedossier: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

procedure cWAINIBase.ChargeDBParams(WAInfoIni: cWAInfoIni);
{$IFNDEF EAGLCLIENT}
var
  EncryptUserPW: boolean;
  q: TQuery;
  where, sql: string;
  OldDriver: TDBDriver;
{$ENDIF EAGLCLIENT}
begin
  traceexecution('cWAINIBase.ChargeDBParams');
  afficheTrace(self);
  if ModeEagl then
  begin
    inherited ChargeDBParams(WAInfoIni);
  end
  else
  begin
{$IFNDEF EAGLCLIENT}
    EncryptUserPW := true;
    q := nil;
    try
      (*if pos(fTokenCar ,WAInfoIni.NomDeBase)>0 then
        begin
        st:=WAInfoIni.NomDeBase ;
        WAInfoIni.NomDeBase:=READTOKENPipe(st,fTokenCar) ;
        Entite:=READTOKENPipe(st,fTokenCar) ;
        Groupe:=READTOKENPipe(st,fTokenCar) ;
        Gamme:=st ;
        end ;*)
      getElementFromNom(WAInfoIni);
      //08/08/2007
      sql := 'SELECT * FROM ADMSOCIETE WHERE ASO_NOM ="' + WAInfoIni.NomDeBase + '" and ';
      where := 'ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"' ; // AND ASO_GAMME="' + gamme + '"';
      traceexecution(sql + where);
      q := DBOpenSQL(leDB, Sql + where, DBDriver);
      if not Q.eof then
      begin
        WAInfoIni.Driver := q.findfield('ASO_DRIVER').AsString;
        WAInfoIni.server := q.findfield('ASO_SERVEUR').AsString;
        WAInfoIni.Path := q.findfield('ASO_PATH').AsString;
        WAInfoIni.ODBC := q.findfield('ASO_PILOTE').AsString;
        WAInfoIni.Base := q.findfield('ASO_NOMBASE').AsString;
        WAInfoIni.User := q.findfield('ASO_LOGIN').AsString;
        WAInfoIni.Pass := q.findfield('ASO_PASSWORD').AsString;
        //Groupe := q.findfield('ACL_GROUPECWAS').AsString ;
        WAInfoIni.Share := q.findfield('ASO_SHARE').AsString;
        //WAInfoIni.Gamme := q.findfield('ASO_GAMME').AsString;
        //StDir:=IniFile.ReadString(StNom,'Dir','');
        WAInfoIni.Options := q.findfield('ASO_OPTIONS').AsString; ; //08/08/2007
        if EncryptUserPW then
        begin
          WAInfoIni.User := DeCryptageSt(WAInfoIni.User);
          WAInfoIni.Pass := DeCryptageSt(WAInfoIni.Pass);
        end;
        WAInfoIni.Pass := DecryptJFD(WAInfoIni.Pass);
      end;
      ferme(Q);
      //YCP pour usdatetime qui utilise le V_PGI.DBDriver ! et comme on est pas sur la DBSOC
      // on est mort !!!!
      OldDriver := V_PGI.Driver;
      {$IFNDEF EAGLSERVER} V_PGI.Driver := DBDriver; {$ELSE} LookupCurrentSession.LeDBDriver := DBDriver; {$ENDIF eAglServer}
      try
        sql:='update ADMSOCIETE set ASO_DATEACCESS="'+TCbpDates.UsDateTime(NowH)+'" where ' ;
      finally
        {$IFNDEF EAGLSERVER}V_PGI.Driver := OldDriver;{$ELSE}LookupCurrentSession.LeDBDriver := OldDriver;{$ENDIF eAglServer}
        end ;
      traceexecution(sql + where);
      MyExecuteSQL(sql+where) ;
      //08/08/2007
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'ChargeDBParams: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.GetSocRef(): string;
var
  q: TQuery;
  sql: string;
  nomsocref: string ;
  WAEnvironnement: cWAEnvironnement ;
begin
  if ModeEagl then
  begin
    result := inherited GetSocRef();
  end
  else
  begin
    result := '';
    q := nil;
    try
     WAEnvironnement:=cWAEnvironnement.create() ;
     // $$$ JP 18/01/08
     try
      if getEnvironnement(WAEnvironnement) then
        begin
        nomsocref:=WAenvironnement.Socref ;
        sql := 'SELECT ASO_NOM,ASO_IDHEBERGE,ASO_IDCLIENT '
           + ' FROM ADMSOCIETE WHERE ASO_NOM="'+nomsocref+'"';
        traceexecution(sql);
        q := DBOpenSQL(leDB, Sql, DBDriver);
        if not q.eof then
          result := q.fields[0].AsString + fTokenCar + q.fields[1].AsString + fTokenCar + q.fields[2].AsString ; //+ fTokenCar + Gamme;
        ferme(Q);
        end else
        begin //YCP cas ou il n'existe pas d'environnement par defaut < AGL580
        sql := 'SELECT ASO_NOM,ASO_IDHEBERGE,ASO_IDCLIENT '
             + ' FROM ADMSOCIETE WHERE ASO_REFERENCE="X" ';
        traceexecution(sql);
        q := DBOpenSQL(leDB, Sql, DBDriver);
        if not q.eof then
          result := q.fields[0].AsString + fTokenCar + q.fields[1].AsString + fTokenCar + q.fields[2].AsString ; //+ fTokenCar + Gamme;
        ferme(Q);
        end ;
     // $$$ JP 18/01/08
     finally
       WAEnvironnement.free;
     end;
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'GetSocRef: ' + E.Message;
      end;
    end;
  end;
end;


function cWAINIBase.SetDossierVerrou(BaseNom: string; Etat: TDossierVerrou): boolean;
var
{$IFNDEF EAGLCLIENT}
  where, sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited setDossierVerrou(BaseNom,Etat) ;
  end
  else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    try
      (*if pos(fTokenCar ,BaseNom)>0 then
        begin
        st:=BaseNom ;
        BaseNom:=READTOKENPipe(st,fTokenCar) ;
        Entite:=READTOKENPipe(st,fTokenCar) ;
        Groupe:=READTOKENPipe(st,fTokenCar) ;
        Gamme:=st ;
        end ;*)
      getElementFromNom(BaseNom);

      sql := 'UPDATE ADMSOCIETE SET ASO_VERROU="' + DossierVerrouToString(Etat) + '" WHERE ASO_NOM ="' + BaseNom + '" ';
      where := ' AND ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"' ;// AND ASO_GAMME="' + gamme + '"';
      MyExecuteSQL(sql + where);
    except
      on E: Exception do
      begin
        ErrorMessage := 'SetDossierVerrou: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.GetDossierVerrou(BaseNom: string): TDossierVerrou;
var
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  where, sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited getDossierVerrou(BaseNom) ;
  end
  else
  begin
    result := dvinconnu;
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      (*if pos(fTokenCar ,BaseNom)>0 then
        begin
        st:=BaseNom ;
        BaseNom:=READTOKENPipe(st,fTokenCar) ;
        Entite:=READTOKENPipe(st,fTokenCar) ;
        Groupe:=READTOKENPipe(st,fTokenCar) ;
        Gamme:=st ;
        end ;*)
      getElementFromNom(BaseNom);

      sql := 'SELECT ASO_VERROU FROM ADMSOCIETE WHERE ASO_NOM ="' + BaseNom + '" ';
      where := ' AND ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"' ; // AND ASO_GAMME="' + gamme + '"';
      traceexecution(sql + where);
      q := DBOpenSQL(leDB, Sql + where, DBDriver);
      // $$$ JP 18/01/08
      try
        if not q.eof then
          result := StringToDossierVerrou(q.fields[0].AsString);
      // $$$ JP 18/01/08
      finally
        ferme (Q);
      end;
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'GetDossierVerrou: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{$IFNDEF EAGLCLIENT}
function cWAINIBase.getDB: TDataBase;
var
  WAINIFile: cWAINIFile;
  WAInfoIni: cWAInfoIni;
begin
  if (xxDB = nil) then
  begin
   WAINIFile := cWAINIFile.create;
   // $$$ JP 18/01/08
   try
    WAINIFile.NomIni := BASECHAPEAUINI;
    WAInfoIni := cWAInfoIni.create;

    // $$$ JP 18/01/08
    try
      WAInfoIni.NomDeBase := 'PGIADMIN';
      WAINIFile.ChargeDBParams(WAInfoIni);
      WAInfoIni.afficheTrace(WAInfoIni);
      //inc(cpte);
      xxDB := TCBPDatabases.CreateNamedDataBase('ADMIN') ;

      WAINIFile.connecteDB(WAInfoIni, xxDB);
      DBDriver := WAInfoIni.DBDriver;
    // $$$ JP 18/01/08
    finally
      WAInfoIni.Free;
    end;
   // $$$ JP 18/01/08
   finally
    WAINIFile.free;
   end;
  end;
  if xxDB = nil then
    TraceExecution('uWAIniBase: Erreur de création de TDataBase');
  result := xxDB;
end;

function cWAINIBase.getDB2(NomBase,NomIni:string): TDataBase;
var
  WAIniBase: cWAIniBase;
  WAInfoIni: cWAInfoIni;
  xxDBNew  : TDataBase;
  stEntite,
  stGroupe: string ;
begin
   //WAINI:= cWAIni.HeritedInstance();
   WAIniBase := cWAIniBase.Create;
   try
    WAIniBase.NomIni := NomIni;
    WAInfoIni := cWAInfoIni.create;
    try
      getElementFromNom(NomBase, stEntite, stGroupe) ; //, stGamme);
      if (NomBase <>'') then
         begin
         WAInfoIni.NomDeBase := NomBase;//'PGIADMIN';
         WAIniBase.IdHeberge:= stEntite;
         WAIniBase.IdClient:= stGroupe;
         WAIniBase.ChargeDBParams(WAInfoIni);
         WAInfoIni.afficheTrace(WAInfoIni);
         xxDBNew := TCBPDatabases.CreateNamedDataBase('COPYTABLE') ;
         WAIniBase.connecteDB(WAInfoIni, xxDBNew);
         DBDriver := WAInfoIni.DBDriver;
         end
         else
         TraceExecution('uWAIniBase: Erreur de création de TDataBase, nom de la base vide');
    finally
      WAInfoIni.Free;
    end;
   finally
    WAIniBase.free;
   end;
  if xxDBNew = nil then
    TraceExecution('uWAIniBase: Erreur de création de TDataBase');
  result := xxDBNew;
end;


function cWAINIBase.getNouveauDossier(): string;
var
  q: TQuery;
  sql: string;
  ii: integer;
begin
  q := nil;
  try
    sql := 'SELECT max(ASO_IDDOSSIER) FROM ADMSOCIETE WHERE ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"' ; //+ ' AND ASO_GAMME="' + gamme + '" ';
    traceexecution(sql);
    q := DBOpenSQL(leDB, Sql, DBDriver);
    if not q.eof then
      ii := q.fields[0].asInteger
    else
      ii := 0;
    inc(ii);
    result := format('%03.3d', [ii]);
    ferme(Q);
  except
    on E: Exception do
    begin
      result := '001';
      if Assigned(q) then
        ferme(Q);
      ErrorMessage := 'GetNouveauDossier: ' + E.Message;
    end;
  end;
end;
{$ENDIF EAGLCLIENT}

function cWAINIBase.getListOfHeberge(): HTStringList;
var
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  sql: string;
  admin: boolean;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited getListOfHeberge ;
  end
  else
  begin
    result := HTStringList.Create();
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      admin := IsAdmin(V_PGI.userlogin);
      sql := 'SELECT * FROM ADMHEBERGE';
      if not admin then
        sql := sql + ' WHERE AHE_IDHEBERGE="' + V_PGI.userlogin + '"';
      traceexecution(sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      while not Q.Eof do
      begin
        result.Add(q.findfield('AHE_IDHEBERGE').AsString + '=' + q.findfield('AHE_NOM').AsString);
        q.next;
      end;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'getListOfHeberge: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.GetListofEnvironnement(): HTStringList;
var
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEAGL then
  begin
  result:=inherited GetListofEnvironnement ;
  end
  else
  begin
  result := HTStringList.Create;
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      sql := 'SELECT AEN_IDENV,AEN_LIBELLE FROM ADMENVIRONNEMENT';
      traceexecution(sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      while not Q.Eof do
      begin
        result.Add(q.findfield('AEN_IDENV').AsString + '=' + q.findfield('AEN_LIBELLE').AsString);
        q.next;
      end;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'GetListofEnvironnement: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.getListOfClient(): HTStringList;
var
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEAGL then
  begin
  result:=inherited getListOfClient ;
  end
  else
  begin
  result := HTStringList.Create;
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      sql := 'SELECT * FROM ADMCLIENT WHERE ACL_IDHEBERGE="' + IDHeberge + '"';
      traceexecution(sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      while not Q.Eof do
      begin
        result.Add(q.findfield('ACL_IDCLIENT').AsString + '=' + q.findfield('ACL_NOM').AsString);
        q.next;
      end;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'getListOfClient: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.getListOfDossier(): HTStringList;
var
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  where, sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEAGL then
  begin
  result:=inherited getListOfDossier ;
  end
  else
  begin
  result := HTStringList.Create;
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      where := '';
      sql := 'SELECT * FROM ADMSOCIETE WHERE ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '"';
      //if (gamme <> '') then where := where + ' AND ASO_GAMME="' + gamme + '"';

      traceexecution(sql + where);
      q := DBOpenSQL(leDB, Sql + where, DBDriver);
      while not Q.Eof do
      begin
        result.Add(q.findfield('ASO_IDDOSSIER').AsString + '=' + q.findfield('ASO_NOM').AsString);
        q.next;
      end;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'getListOfDossier: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.UserPassValidate(user, pass: string): boolean;
var
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited UserPassValidate(user, pass) ;
  end
  else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      sql := 'SELECT US_UTILISATEUR FROM UTILISAT WHERE US_ABREGE="' + User + '" AND US_PASSWORD="' + CryptageSt(pass) + '"';
      traceexecution(sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      result := not Q.Eof;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'UserPassValidate: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.IsFicBase(): boolean;
var
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited IsFicBase() ;
  end
  else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      sql := 'SELECT AHE_FICBASE FROM ADMHEBERGE WHERE AHE_IDHEBERGE="' + IDheberge + '" and AHE_FICBASE="X"';
      traceexecution(sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      result := not Q.Eof;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'IsFicBase: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{$IFNDEF EAGLCLIENT}
function cWAINIBase.chargeSQL2TOB(nomTob, sql: string): TOB;
var
  i: integer;
  q: TQuery;
begin
  result := TOB.create(nomTOB, nil, -1);
  traceexecution(sql);
  q := DBOpenSQL(leDB, Sql, DBDriver);
  // $$$ JP 18/01/08
  try
   try
    for i := 0 to q.fields.count - 1 do
    begin
      result.AddChampSup(q.fields[i].Fieldname, false);
      if not q.eof then
        result.PutValue(q.fields[i].Fieldname, q.Fields[i].AsVariant);
    end;
    // $$$ JP 18/01/08 ferme(Q);
   except
    on E: Exception do
    begin
      if Assigned(q) then
        ferme(Q);
      ErrorMessage := 'ChargeSQL2TOB: ' + E.Message;
    end;
   end;
  // $$$ JP 18/01/08
  finally
   ferme (Q);
  end;
end;
{$ENDIF EAGLCLIENT}

function cWAINIBase.getHeberge(): TOB;
var
{$IFNDEF EAGLCLIENT}
  T1: Tob;
  st: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited getHeberge() ;
  end
  else
  begin
  result := TOB.Create('', nil, -1);
{$IFNDEF EAGLCLIENT}
    T1 := chargeSQL2TOB('XXHEBERGE'
      , 'SELECT * FROM ADMHEBERGE WHERE AHE_IDHEBERGE="' + IDHeberge + '"');
    //YCP 16/07/05 initialisation des infos de serveur depuis les paramsocs
    st := trim(T1.getValue('AHE_IDHEBERGE'));
    if (st = '') then
      initHeberge(T1);
    //YCP 16/07/05 fin

    T1.ChangeParent(result, -1);
{$ENDIF EAGLCLIENT}
  end;
end;




{$IFNDEF EAGLCLIENT}
procedure cWAINIBase.initHeberge(T1: TOB);
var
  val: string ;
begin
  if T1 = nil then
    exit;
  if not T1.FieldExists('AHE_DRIVER') then
    T1.AddChampSup('AHE_DRIVER', false);
  val:=getParamsoc('SO_DRIVER') ;
  val:=getParamsocSecur('SO_DRIVER', '') ;
  T1.putValue('AHE_DRIVER', val);
  if not T1.FieldExists('AHE_SERVEUR') then
    T1.AddChampSup('AHE_SERVEUR', false);
  T1.putValue('AHE_SERVEUR', getParamsocSecur('SO_SERVEUR', ''));
  if not T1.FieldExists('AHE_PILOTE') then
    T1.AddChampSup('AHE_PILOTE', false);
  T1.putValue('AHE_PILOTE', getParamsocSecur('SO_PILOTE', ''));
  if not T1.FieldExists('AHE_LOGIN') then
    T1.AddChampSup('AHE_LOGIN', false);
  T1.putValue('AHE_LOGIN', getParamsocSecur('SO_LOGIN', ''));
  if not T1.FieldExists('AHE_PASSWORD') then
    T1.AddChampSup('AHE_PASSWORD', false);
  T1.putValue('AHE_PASSWORD', getParamsocSecur('SO_PASSWORD', ''));
  if not T1.FieldExists('AHE_CHEMIN') then
    T1.AddChampSup('AHE_CHEMIN', false);
  T1.putValue('AHE_CHEMIN', getParamsocSecur('SO_CHEMIN', ''));
  //08/08/2007
  if not T1.FieldExists('AHE_OPTIONS') then
    T1.AddChampSup('AHE_OPTIONS', false);
  T1.putValue('AHE_OPTIONS', getParamsocSecur('SO_OPTIONS', ''));
  //08/08/2007
end;
{$ENDIF EAGLCLEINT}

function cWAINIBase.GetClient(): TOB;
var
{$IFNDEF EAGLCLIENT}
  T1: Tob;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited getClient ;
  end
  else
  begin
  result := TOB.Create('', nil, -1);
{$IFNDEF EAGLCLIENT}
    T1 := chargeSQL2TOB('XXCLIENT'
      , 'SELECT * FROM ADMCLIENT WHERE ACL_IDHEBERGE="' + IDHeberge + '" '
      + 'AND ACL_IDCLIENT="' + IDClient + '"');
    T1.ChangeParent(result, -1);
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.GetDossier(): TOB;
var
{$IFNDEF EAGLCLIENT}
  T1: Tob;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited getDossier ;
  end
  else
  begin
  result := TOB.Create('', nil, -1);
{$IFNDEF EAGLCLIENT}
    T1 := chargeSQL2TOB('XXSOCIETE'
      , 'SELECT * FROM ADMSOCIETE '
      + 'WHERE ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '" '
      + 'AND ASO_IDDOSSIER="' + IDDossier + '" ');
    T1.ChangeParent(result, -1);
{$ENDIF EAGLCLIENT}
  end;
end;

procedure incrementeCombo(var val: string; ln: integer);
var
  i: integer;
  st: string;
begin
  st := val;
  i := ln;
  while (i >= 0) do
  begin
    if (st[i] = '9') then
    begin
      st[i] := 'A';
      i := -1;
    end
    else if (st[i] = 'Z') then
    begin
      st[i] := '0';
      dec(i);
    end
    else
    begin
      st[i] := char(ord(st[i]) + 1);
      i := -1;
    end;
  end;
  if st > val then
    val := st; //pour repasser à 0
end;

function cWAINIBase.chercheUserSuivant(var User: string): boolean;
var
  tUser: tob;
  q: TQuery;
begin
  tUser := TOB.Create('UTILISAT_s', nil, -1);
  q := DBOpenSQL(leDB, 'SELECT * FROM UTILISAT ORDER BY US_UTILISATEUR', DBDriver);
  tUser.LoadDetailDB('UTILISAT', '', '', Q, false);
  ferme(q);

  try
    user := '000';
    while user <> 'ZZZ' do
    begin
      if tUser.FindFirst(['US_UTILISATEUR'], [user], true) = nil then
        break;
      incrementeCombo(user, 3);
    end;
    result := user <> 'ZZZ';
  finally
    tUser.free;
    //raise;
  end;
end;

function cWAINIBase.setHeberge(laTob: TOB): boolean;
var
{$IFNDEF EAGLCLIENT}
  //num: integer;
  tUser: tob;
  utilisat: string;
  //q: tquery;
  //sql: string;
  maxUser: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited setHeberge(laTob) ;
  end
  else
  begin
    result := false;
    if laTob = nil then
    begin
      ErrorMessage := 'Le paramètre (tob) est inexistant';
      exit;
    end;
{$IFNDEF EAGLCLIENT}
    try
      if CtrlHeberge(laTob) then
      begin
        utilisat := laTob.getValue('AHE_UTILISAT');
        tUser := chargeSQL2TOB('##UTILISAT'
          , 'SELECT * FROM UTILISAT WHERE US_UTILISATEUR="' + utilisat + '"');
        if trim(utilisat) = '' then
        begin
          //num := 0;
          (*sql := 'SELECT max(US_UTILISATEUR) FROM UTILISAT';
          traceexecution(sql);
          q := DBOpenSQL(leDB, Sql, DBDriver);
          ///YCP passage en alpha des comptes
          maxUser:=q.Fields[0].AsString ;
          ferme(Q) ;
          if (maxUser<>'') then
            begin      *)
          if not chercheUserSuivant(MaxUser) then
          begin
            Result := false;
            ErrorMessage := 'Le nombre d''utilisateurs maximum est atteint (ZZZ)';
            tUser.free;
            exit;
          end;
          (*end else
          begin
          MaxUser:='001' ;
          end ;*)
          utilisat := bourreGauche(MaxUser, 3, '0');

          (*if IsNumeric(q.Fields[0].AsString) then num := q.Fields[0].AsInteger + 1;
          if (num > 999) then
          begin
            Result := false;
            ErrorMessage := 'Le nombre d''utilisateurs maximum est atteint (999)';
            exit;
            tUser.free;
          end;
          utilisat := bourreGauche(IntToStr(num), 3, '0');*)
        end;
        laTob.putValue('AHE_UTILISAT', utilisat);
        tUser.PutValue('US_UTILISATEUR', utilisat);
        tUser.PutValue('US_ABREGE', laTob.getValue('AHE_IDHEBERGE'));
        tUser.PutValue('US_LIBELLE', laTob.getValue('AHE_NOM'));
        tUser.PutValue('US_PASSWORD', laTob.getValue('AHE_PASSUTILISAT'));
        SaveLaTob(tUser);
        tUser.free;
         //24/05/2007 début
        if (trim(laTob.GetValue('AHE_IDENV'))='') and (isCreate) then
           laTob.PutValue('AHE_IDENV',GetParamSocSecur('SO_IDENVCREATE','')) ;
        if (isCreate) then
           begin
           if laTob.FieldExists('AHE_FICBASE') then
            laTob.PutValue('AHE_FICBASE','X')
            else
            laTob.AddChampSupValeur('AHE_FICBASE','X') ; //mko 13/08/09 Créer par défaut en fichier en base
           end;
         //24/05/2007 fin
        SaveLaTob(laTob);
        result := true;
      end;
    except
      on E: Exception do
      begin
        ErrorMessage := 'SetHeberge: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.setClient(laTob: TOB): boolean;
begin
  if ModeEagl then
  begin
    result:=inherited setClient(laTob) ;
  end
  else
  begin
  result := false;
  if laTob = nil then
  begin
    ErrorMessage := 'Le paramètre (tob) est inexistant';
    exit;
  end;
{$IFNDEF EAGLCLIENT}
    try
      if CtrlClient(laTob) then
      begin
        SaveLaTob(laTob);
        result := true;
      end;
    except
      on E: Exception do
      begin
        ErrorMessage := 'SetClient: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.setDossier(laTob: TOB; BaseFrom: string; AvecRecupHeberge: boolean): boolean;
var
{$IFNDEF EAGLCLIENT}
  ioWABD: cWABD;
  sql, nom, chemin,
  Cheminmdf,
  Cheminldf,
  NomBaseDest : string;
  q: TQuery;
  okok   : boolean;
  where, stEntite, stGroupe(*, stGamme*): string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited SetDossier(laTob,BaseFrom,AvecRecupHeberge) ;
  end
  else
  begin
  result := false;
  if laTob = nil then
  begin
    ErrorMessage := 'Le paramètre (tob) est inexistant';
    exit;
  end;
{$IFNDEF EAGLCLIENT}
    try
      if CtrlDossier(laTob) then
      begin
        okok := getSocieteFromHeberge(laTob, AvecRecupHeberge);
        if not okok then
        begin
          ErrorMessage := 'Erreur de récupération d''informations de l''hébergé';
        end
        else
        begin
          ioWABD := cWABD.create(string(laTOB.getValue('ASO_DRIVER')));
          with ioWABD do
          begin
            Base := laTOB.getValue('ASO_NOMBASE');
            if (ioWABD is cWABDSQLSERVER) then
            begin
              chemin := IncludeTrailingPathDelimiter(laTOB.getValue('CHEMIN'))
                + 'D' + bourreGauche(laTOB.getValue('ASO_IDHEBERGE'), 3, '0')
                + bourreGauche(laTOB.getValue('ASO_IDCLIENT'), 3, '0');
              Cheminmdf := IncludeTrailingPathDelimiter(chemin) + laTOB.getValue('ASO_NOMBASE') + '.mdf';
              Cheminldf := IncludeTrailingPathDelimiter(chemin) + laTOB.getValue('ASO_NOMBASE') + '.ldf';
              cWABDSQLSERVER(ioWABD).Masterserver := laTOB.getValue('ASO_SERVEUR');
              cWABDSQLSERVER(ioWABD).MasterUser := DeCryptageSt(laTOB.getValue('ASO_LOGIN'));
              cWABDSQLSERVER(ioWABD).MasterPass := DeCryptageSt(laTOB.getValue('ASO_PASSWORD'));
              cWABDSQLSERVER(ioWABD).mdfFile := Cheminmdf;
              cWABDSQLSERVER(ioWABD).ldfFile := Cheminldf;
            end;
            if not ExisteBase() then
              begin
              //ON RECUP Les identifiants de LA BASE FROM=> Base modele
              getElementFromNom(BaseFrom, StEntite, StGroupe);
              //RECUP LE NOM PHYSIQUE DE LA BASE MODELE POUR CREER LA BASE
              sql := 'select ASO_NOMBASE from ADMSOCIETE where ASO_NOM="' + BaseFrom + '"';
              where := ' AND ASO_IDHEBERGE="' + stEntite + '" AND ASO_IDCLIENT="' + stGroupe + '"' ; // AND ASO_GAMME="' + stGamme + '"';
              TraceExecution(sql + where);
              Q := DBOpenSQL(leDB, Sql + where, DBDriver);
              try
              if not Q.eof then
                nom := q.Fields[0].AsString
                else
                nom := '';
              finally
               ferme(Q);
              end;
              if not CreerBaseFrom(nom) then  ////Creation de la base a partir de la base modele
                begin
                result := false;
                ErrorMessage := 'La création physique de la base a échouée';
                end
                else
                begin
                SaveLaTob(laTob);
                result:= True;
                //RECUP DES INFOS DE LA BASE DE DESTINATION POUR CONNECTION
                NomBaseDest:=laTOB.getValue('ASO_NOM')+FTokenCar+ laTOB.getValue('ASO_IDHEBERGE')+ FTokenCar+ laTOB.getValue('ASO_IDCLIENT');
                //SUPPRESSION DES ELEMENTS A NE PAS RECUP DE LA BASE MODELE
                if not DeleteTableToBase(NomBaseDest,ListeTable) then
                   begin
                   result := False;
                   ErrorMessage := 'La copie des données de la base modèle vers la nouvelle base a échouée';
                   end;
                end;
            end
            else
            begin
              SaveLaTob(laTob);
              result := true;
            end;
            free;
          end;
        end;
      end;
    except
      on E: Exception do
      begin
        ErrorMessage := 'SetDossier: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : MKO
Créé le ...... : 17/06/2009
Modifié le ... :   /  /    
Description .. : Copie les tables de la base modèle listées dans TABLES
Suite ........ : séparées avec des ; , dans la base que l'on vient de créer:
Suite ........ : de BaseFrom vers Base
Mots clefs ... : 
*****************************************************************}
function cWAINIBase.DeleteTableToBase(Base,Tables:String) : boolean;
var xxDBSource  : TDataBase;
    TableListe,
    TableCopy   : String;
    QQ         : TQUERY;
begin
    result:=True;
    //CONNECTION SUR LA BASES
    xxDBSource:=getDB2(Base,'ES1.ini');
    try
      //RECUP DES TABLES A COPIER
      TableListe:= Tables;
      if (Assigned(xxDBSource) and (xxDBSource.Connected)) then
         begin
         InitModeEchange(xxDBSource); //MKO 25/06/09 REPASSE LE MODE ECHANGE A SYN DANS TOUS LES CAS
         InitTVA(xxDBSource); //MKO 30/06/09 ON DCOCHE SYSTEMATIQUEMENT LA GESTION DE LA TVA
         While (TableListe<>'') do
            begin
            TableCopy:= ReadTokenSt(TableListe);
            if (TableCopy<>'') then
               begin
               QQ:=DBOpenSQL(xxDBSource,'Select DT_NOMTABLE FROM DETABLES where DT_NOMTABLE="'+TableCopy+'"' ,TDBDriver(StToDriver(xxDBSource.DriverName,false)));
                try
                if Not QQ.Eof then
                   begin
                   if TableCopy='JOURNAL' then
                      DeleteSpecJournal(xxDBSource)
                      else
                   if TableCopy='SOUCHE' then
                      DeleteSpecSouche(xxDBSource)
                      else
                   if TableCopy='UTILISAT' then
                      DeleteSpecUTILISAT(xxDBSource)
                      else
                   if (TableCopy='YFILES') OR (TableCopy='YFILEPARTS') then
                      DeleteSpecFiles(xxDBSource)
                      else
                      ExecuteSQLDB('DELETE FROM '+TableCopy,xxDBSource); 
                   end;
                finally
                Ferme(QQ) ;
                end;
               end
               else
               begin
               result:=False;
               break;
               end;
            end;
         end;
   finally
     begin
     cWAINI.Deconnectedb(xxDBSource) ;
     FreeAndNil(xxDBSource) ;
     end;
   end;
end;

procedure cWAINIBase.InitTVA(xxDBSource:TDataBase);  // MKO 25/06/09
begin
   ExecuteSQLDB('UPDATE PARAMSOC set SOC_DATA="-" where SOC_NOM="SO_TVACOMPLETE" ' ,xxDBSource);
end;

procedure cWAINIBase.InitModeEchange(xxDBSource:TDataBase);  // MKO 25/06/09
begin
   ExecuteSQLDB('UPDATE PARAMSOC set SOC_DATA="SYN" where SOC_NOM="SO_MODEECHANGE" ' ,xxDBSource);
end;

procedure cWAINIBase.DeleteSpecJournal(xxDBSource:TDataBase);
begin
   ExecuteSQLDB('UPDATE JOURNAL set J_REFJAL="" ' ,xxDBSource);
end;

procedure cWAINIBase.DeleteSpecSouche(xxDBSource:TDataBase);
begin
   ExecuteSQLDB('UPDATE SOUCHE set SH_NUMERO=0' ,xxDBSource);
end;

procedure cWAINIBase.DeleteSpecUTILISAT(xxDBSource:TDataBase);
begin
   ExecuteSQLDB('DELETE UTILISAT WHERE US_LIBELLE<>"CEGID"' ,xxDBSource);
end;

procedure cWAINIBase.DeleteSpecFiles(xxDBSource:TDataBase);
begin
   ExecuteSQLDB('DELETE FROM YFILES where YFI_FILEGUID not in (select KIM_MODEOPERATOIRE FROM KITIMPORT)' ,xxDBSource); 
   ExecuteSQLDB('DELETE FROM YFILEPARTS where YFP_FILEGUID not in (select KIM_MODEOPERATOIRE FROM KITIMPORT) ' ,xxDBSource);
end;



function cWAINIBase.delHeberge: boolean;
begin
  if ModeEagl then
  begin
    result:=inherited DelHeberge() ;
  end
  else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    try
      result := not myExisteSQL('SELECT ACL_IDHEBERGE FROM ADMCLIENT WHERE ACL_IDHEBERGE="' + IDHeberge + '"');
      if result then
      begin
        MyExecuteSQL('DELETE FROM UTILISAT WHERE US_UTILISATEUR=(SELECT AHE_UTILISAT FROM ADMHEBERGE WHERE AHE_IDHEBERGE="' + IDHeberge + '")');
        MyExecuteSQL('DELETE FROM ADMHEBERGE WHERE AHE_IDHEBERGE="' + IDHeberge + '"');
      end
      else
        ErrorMessage := 'Ce compte contient des clients, il ne peux être effacé';
    except
      on E: Exception do
      begin
        ErrorMessage := 'DelHeberge: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.delClient: boolean;
begin
  if ModeEagl then
  begin
    result:=inherited delClient ;
  end
  else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    try
      result := not myExisteSQL('SELECT ASO_IDHEBERGE FROM ADMSOCIETE WHERE ASO_IDHEBERGE="' + IDHeberge + '" '
        + 'AND ASO_IDCLIENT="' + IDClient + '"');
      if result then
        MyExecuteSQL('DELETE FROM ADMCLIENT WHERE ACL_IDHEBERGE="' + IDHeberge + '" AND ACL_IDCLIENT="' + IDClient + '"')
      else
        ErrorMessage := 'Ce client contient des dossiers, il ne peux être effacé';
    except
      on E: Exception do
      begin
        ErrorMessage := 'DelClient: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.delDossier(AvecSupprPhysique, ForceSupprSiEchecPhysique: boolean): boolean;
{$IFNDEF EAGLCLIENT}
var
  ioWABD: cWABD;
  q: TQuery;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited delDossier(AvecSupprPhysique, ForceSupprSiEchecPhysique) ;
  end
  else
  begin
    result := true;
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      sql := 'SELECT * FROM ADMSOCIETE '
        + 'WHERE ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '" '
        + 'AND ASO_IDDOSSIER="' + IDDossier + '" ';
      traceexecution(sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      if q.eof then
      begin
      result:=false ;
      ErrorMessage := 'Le dossier n''existe pas' ;
      end else
      begin
        if result and AvecSupprPhysique then
        begin
          ioWABD := cWABD.create(q.findField('ASO_PILOTE').asString);
          with ioWABD do
          begin
            parle := false;
            if (ioWABD is cWABDSQLSERVER) then
            begin
              cWABDSQLSERVER(ioWABD).Masterserver := q.findField('ASO_SERVEUR').asString;
              cWABDSQLSERVER(ioWABD).MasterUser := DeCryptageSt(q.findField('ASO_LOGIN').asString);
              cWABDSQLSERVER(ioWABD).MasterPass := DeCryptageSt(q.findField('ASO_PASSWORD').asString);
              cWABDSQLSERVER(ioWABD).Base := q.findField('ASO_NOMBASE').asString;
            end;
            result := SupprimerBase();
            if not result then
              ErrorMessage := 'Une erreur c''est produite lors de la suppression physique de la base';
            free;
          end;
        end;
        if result or ForceSupprSiEchecPhysique then
        begin
          sql := 'DELETE FROM ADMSOCIETE '
            + 'WHERE ASO_IDHEBERGE="' + IDHeberge + '" AND ASO_IDCLIENT="' + IDClient + '" '
            + 'AND ASO_IDDOSSIER="' + IDDossier + '" ';
          traceexecution(sql);
          MyExecuteSQL(sql);
        end;
      end;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'DelDossier: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function GetNextIDString(var MaxString: string): string;
var
  i, pos: integer;
  car: char;
  zero: string;
begin
  zero := '';
  for i := 0 to 5 - length(MaxString) do
    zero := zero + '0';
  MaxString := zero + MaxString;
  car := '0';
  pos := length(MaxString);
  while (car = '0') and (pos > 0) do
  begin
    car := char(ord(MaxString[pos]) + 1);
    MaxString[pos] := car;
    if (car > 'Z') then
    begin
      car := '0';
      MaxString[pos] := car;
      dec(pos);
    end;
  end;
  result := MaxString;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Yann PELUD
Créé le ...... : 23/07/2004
Modifié le ... :   /  /
Description .. : Récupère le prochain ID de :
Suite ........ : 0: Hebergé
Suite ........ : 1: Client
Suite ........ : 2: Dossier
Mots clefs ... :
*****************************************************************}
function cWAINIBase.GetNextID(leQuel: integer): string;
var
  st: string;
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result:=inherited getNextId(lequel) ;
  end
  else
  begin
  result := '';
{$IFNDEF EAGLCLIENT}
    q := nil;
    try
      case leQuel of
        0: sql := 'SELECT MAX(AHE_IDHEBERGE) FROM ADMHEBERGE';
        1: sql := 'SELECT MAX(ACL_IDCLIENT)  FROM ADMCLIENT  WHERE ACL_IDHEBERGE="' + IDHeberge + '"';
        2: sql := 'SELECT MAX(ASO_IDDOSSIER) FROM ADMSOCIETE WHERE ASO_IDHEBERGE="' + IDHeberge + '"'; //un chrono de dossier unique par expert
      end;
      traceexecution(sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      st := q.fields[0].AsString;
      result := GetNextIDString(st);
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'GetNextID: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.CtrlHeberge(laTob: TOB): boolean;
var
  tmp, idHeberge: string;
  WABD: cWABD;
begin
  idHeberge := LaTob.GetValue('AHE_IDHEBERGE');

  if (iscreate) then
    traceexecution('CREATION !!!');
  if (length(idHeberge) > MAX_LEN_HEBERGE) then
    ErrorMessage := 'La longueur de l''identifiant hébergé doit être au maximum de ' + intToStr(MAX_LEN_HEBERGE) + ' caractères'
  else if (iscreate) and myExisteSQL('SELECT AHE_IDHEBERGE FROM ADMHEBERGE ' + GetWhereCle('ADMHEBERGE', LaTob, tmp)) then
    ErrorMessage := 'L''identifiant existe déjà'
  else if (trim(LaTob.GetString('AHE_NOM')) = '') then
    ErrorMessage := 'Le nom de l''hébergé ne peut être vide'
  else
  begin
    WABD := cWABD.create(string(laTob.GetValue('ASO_PILOTE')));
    if (WABD is cWABDSQLSERVER) then
    begin
      cWABDSQLSERVER(WABD).parle := false;
      TraceExecution('AHE_SERVEUR ' + laTob.GetValue('AHE_SERVEUR'));
      cWABDSQLSERVER(WABD).Masterserver := laTob.GetValue('AHE_SERVEUR');
      TraceExecution('AHE_LOGIN ' + laTob.GetValue('AHE_LOGIN'));
      cWABDSQLSERVER(WABD).MasterUser := DeCryptageSt(laTob.GetValue('AHE_LOGIN'));
      TraceExecution('AHE_PASSWORD ' + laTob.GetValue('AHE_PASSWORD'));
      cWABDSQLSERVER(WABD).MasterPass := DeCryptageSt(laTob.GetValue('AHE_PASSWORD'));
      TraceExecution('master ' + laTob.GetValue('master'));
      cWABDSQLSERVER(WABD).Base := 'master';
      if not cWABDSQLSERVER(WABD).TestConnexion then
        ErrorMessage := 'Le paramétrage du serveur est incorrect';
    end
    else
    begin
      ErrorMessage := 'Le paramétrage du serveur est incorrect (Pilote non pris en charge)';
    end;
    WABD.free;
  end;

  result := ErrorMessage = '';

  //non bloquant
  if (trim(LaTob.GetString('AHE_PASSUTILISAT')) = '') then
    ActionMessage := 'Attention, un mot de passe vide n''est pas conseillé';
end;

function cWAINIBase.CtrlClient(laTob: TOB): boolean;
var
  tmp, idHeberge, idClient,cdkey: string;
begin
  if (iscreate) then
    traceexecution('CREATION !!!');

  idHeberge := LaTob.GetValue('ACL_IDHEBERGE');
  idClient := LaTob.GetValue('ACL_IDCLIENT');
  cdkey:=trim(LaTob.GetValue('ACL_CDKEY')) ; //YCP 15/05/07 
  if (length(idHeberge) > MAX_LEN_HEBERGE) then
    ErrorMessage := 'La longueur de l''identifiant hébergé doit être au maximum de ' + intToStr(MAX_LEN_HEBERGE) + ' caractères'
  else if (length(idClient) > MAX_LEN_CLIENT) then
    ErrorMessage := 'La longueur de l''identifiant client doit être au maximum de ' + intToStr(MAX_LEN_CLIENT) + ' caractères'
  else if (iscreate) and myExisteSQL('SELECT ACL_IDHEBERGE FROM ADMCLIENT ' + GetWhereCle('ADMCLIENT', LaTob, tmp)) then
    ErrorMessage := 'Le client existe déjà'
  else if (cdkey<>'-    -    -') and myExisteSQL('SELECT ACL_CDKEY FROM ADMCLIENT WHERE ACL_CDKEY="' + cdkey + '" ' //YCP 15/05/07 sauf si cdkey est vide
    + ' AND (ACL_IDHEBERGE<>"' + laTob.getValue('ACL_IDHEBERGE') + '" OR ACL_IDCLIENT<>"' + laTob.getValue('ACL_IDCLIENT') + '")') then
    ErrorMessage := 'Le numéro de CDKey est déjà utilisé'
  else if (trim(LaTob.GetString('ACL_NOM')) = '') then
    ErrorMessage := 'Le nom du client ne peut être vide';

  result := ErrorMessage = '';
end;

function cWAINIBase.CtrlDossier(laTob: TOB): boolean;
var
  st, nomsoc, (*idGamme,*) idHeberge, idClient, idDossier: string;
begin
  if (iscreate) then
    traceexecution('CREATION !!!');

  idHeberge := LaTob.GetValue('ASO_IDHEBERGE');
  idClient := LaTob.GetValue('ASO_IDCLIENT');
  idDossier := LaTob.GetValue('ASO_IDDOSSIER');
  //idGamme := LaTob.GetValue('ASO_GAMME');
  NomSoc := LaTob.GetValue('ASO_NOM');
  if (NomSoc = '') then
  begin
    NomSoc := getNomBase((*idGamme,*) idHeberge, iddossier); // ErrorMessage:='Vous devez renseigner le nom de la société' else
    LaTob.PutValue('ASO_NOMBASE', NomSoc);
  end;
  if (length(idHeberge) > MAX_LEN_HEBERGE) then
    ErrorMessage := 'La longueur de l''identifiant hébergé doit être au maximum de ' + intToStr(MAX_LEN_HEBERGE) + ' caractères'
  else if (length(idClient) > MAX_LEN_CLIENT) then
    ErrorMessage := 'La longueur de l''identifiant client doit être au maximum de ' + intToStr(MAX_LEN_CLIENT) + ' caractères'
  else if (length(idDossier) > MAX_LEN_DOSSIER) then
    ErrorMessage := 'La longueur de l''identifiant dossier doit être au maximum de ' + intToStr(MAX_LEN_DOSSIER) + ' caractères'
  else if (trim(LaTob.GetString('ASO_NOM')) = '') then
    ErrorMessage := 'Le nom de la société ne peut être vide'
  (*else if (idGamme = '') then
    ErrorMessage := 'Vous devez choisir une gamme'*)
  else if (iscreate)
       and myExisteSQL('SELECT ASO_IDHEBERGE FROM ADMSOCIETE '
                     + ' WHERE ASO_IDHEBERGE="' + laTob.getValue('ASO_IDHEBERGE') + '" '
                     + ' AND ASO_IDDOSSIER="' + laTob.getValue('ASO_IDDOSSIER')+ '"') then
                     //+ '" AND ASO_GAMME="' + laTob.getValue('ASO_GAMME') + '"') then
    ErrorMessage := 'Le dossier existe déjà'
  else if myExisteSQL('SELECT ASO_NOM FROM ADMSOCIETE '
                     +' WHERE ASO_NOM="' + NomSoc + '"'
                     + ' AND ASO_IDHEBERGE="' + idHeberge + '" AND ASO_IDCLIENT="' + idClient + '"'
                     + ' AND ASO_IDDOSSIER<>"' + idDossier + '"') then //AND ASO_GAMME="' + idGamme + '"
    ErrorMessage := 'Le nom de la société existe déjà';

  if (LaTob.GetValue('ASO_VERROU') = '') then
    LaTob.PutValue('ASO_VERROU', 'AUC');

  //YCP on s'assure que les valeurs sont bonnes
  st := LaTob.getValue('ASO_DATEACCESS') ;
  if not isDateTimeText(st) then
    LaTob.PutValue('ASO_DATEACCESS',date()) ;

  st := LaTob.getValue('ASO_DATEBACKUP') ;
  if not isDateTimeText(st) then
    LaTob.PutValue('ASO_DATEBACKUP',iDate1900) ;

  result := ErrorMessage = '';
end;

function cWAINIBase.getSocieteFromHeberge(laTob: TOB; AvecRecupDescriptionBase: boolean): boolean;
var
  tHeberge: Tob;
  i: integer;
  s, sql, chp1, chp2: string;
  q: TQuery;
  chSupp: TCS;
begin
  result := false;
  tHeberge := TOB.Create('HEBERGE', nil, -1);
  s := laTOB.getValue('ASO_IDHEBERGE');
  sql := 'SELECT * FROM ADMHEBERGE WHERE AHE_IDHEBERGE="' + s + '"';
  traceexecution(sql);
  q := DBOpenSQL(leDB, Sql, DBDriver);
  if q.eof then
  begin
    ErrorMessage := 'L''identifiant de l''hébergé n''a pas été trouvé :' + laTOB.getValue('ASO_IDHEBERGE');
  end
  else
  begin
    tHeberge.LoadDetailDB('HEBERGE', '', '', Q, false);
    if (tHeberge.Detail.count <= 0) then
    begin
      ErrorMessage := 'L''identifiant de l''hébergé n''a pas été trouvé :' + laTOB.getValue('ASO_IDHEBERGE');
    end
    else
    begin
      for i := 0 to tHeberge.detail[0].ChampsSup.Count - 1 do
      begin
        chSupp := TCS(tHeberge.detail[0].ChampsSup[i]);
        chp2 := chSupp.Nom;
        if (AvecRecupDescriptionBase) then
        begin
          chp1 := 'ASO_' + copy(chp2, 5, length(chp2));
          if (laTob.FieldExists(chp1)) and (chp1 <> 'ASO_NOM') then
            laTOB.putValue(chp1, chSupp.Valeur);
        end;
        if (chp2 = 'AHE_CHEMIN') then
          laTOB.AddChampSupValeur('CHEMIN', chSupp.Valeur);
      end;
      result := true;
    end;
  end;
  tHeberge.free;
end;

function cWAINIBase.GetWhereCle(NomTable: string; laTob: tob; var ChampsAIgnorer: string): string;
begin
  result := '';
  ChampsAIgnorer := '';
  if nomTable = 'ADMHEBERGE' then
  begin
    result := 'WHERE AHE_IDHEBERGE="' + laTob.getValue('AHE_IDHEBERGE') + '"'
  end
  else if nomTable = 'ADMCLIENT' then
  begin
    result := 'WHERE ACL_IDHEBERGE="' + laTob.getValue('ACL_IDHEBERGE') + '" '
      + 'AND ACL_IDCLIENT="' + laTob.getValue('ACL_IDCLIENT') + '"'
  end
  else if nomTable = 'ADMSOCIETE' then
  begin
    result := 'WHERE ASO_IDHEBERGE="' + laTob.getValue('ASO_IDHEBERGE') + '" '
      + 'AND ASO_IDCLIENT="' + laTob.getValue('ASO_IDCLIENT') + '" '
      + 'AND ASO_IDDOSSIER="' + laTob.getValue('ASO_IDDOSSIER') + '"'
  end
  else if nomTable = 'UTILISAT' then
  begin
    result := 'WHERE US_UTILISATEUR="' + laTob.getValue('US_UTILISATEUR') + '" ';
    ChampsAIgnorer := 'US_CRC;US_DATECONNEXION;';
  end
  else
    ErrorMessage := 'Table inconnue en sauvegarde:' + nomtable;
end;

function cWAINIBase.SaveLaTob(laTob: tob): boolean;
var
  ChampsAIgnorer, where, sql, nomTable, st: string;
  q: TQuery;
  i: integer;
  insert: boolean ;
begin
  result := false;
  if (copy(laTob.NomTable, 1, 2) = 'XX') then
    nomtable := 'ADM' + copy(laTob.NomTable, 3, length(laTob.NomTable));
  if (copy(laTob.NomTable, 1, 2) = '##') then
    nomtable := copy(laTob.NomTable, 3, length(laTob.NomTable));

  Where := GetWhereCle(NomTable, LaTob, ChampsAIgnorer);
  if (ErrorMessage <> '') then
    exit;

  sql := 'SELECT * FROM ' + nomTable + ' ' + where;
  TraceExecution(sql);
  Q := DBOpenSQL(leDB, Sql, DBDriver, false, false);
  if not q.eof then
    begin q.edit; insert:=false; end
  else
    begin q.insert; insert:=true; end ;
  for i := 0 to q.FieldCount - 1 do
  begin
    if (Pos(';' + q.fields[i].FieldName + ';', ';' + ChampsAIgnorer) > 0) then
      continue;
    st:=laTob.GetString(q.fields[i].FieldName) ;
    if laTob.FieldExists(q.fields[i].FieldName) then
      begin
      if (pos('_DATECREATION',q.fields[i].FieldName)>0) and insert then
         begin
         q.fields[i].AsVariant:=now ;
         end else
      if (pos('_DATEMODIF',q.fields[i].FieldName)>0) then
         begin
         q.fields[i].AsVariant:=now ;
         end else
      if (st<>#0) and not VarIsNull(laTob.getvalue(q.fields[i].FieldName)) then
        begin
        q.fields[i].AsVariant := laTob.getvalue(q.fields[i].FieldName) ;
        end else
        begin
        if (q.fields[i].DataType=ftDateTime) then q.fields[i].AsVariant:=iDate1900
        else if (q.fields[i].DataType=ftInteger) then q.fields[i].AsVariant:=0
          else q.fields[i].AsVariant:='' ;
        end ;
      end ;
  end;
  q.post;
  ferme(Q);
  result := true;
end;

function cWAINIBase.getProduit: string;
{$IFNDEF EAGLCLIENT}
var
  q: TQuery;
  seria, sql: string;
  ss: TPGISeriaServer;
  Product, Users, Order: LongWord;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result := inherited getProduit;
  end
  else
  begin
  result := '';
{$IFNDEF EAGLCLIENT}
    q := nil;
    ss := TPGISeriaServer.Create;
    try
      setEntiteFromNomEntite;
      setGroupeFromNomGroupe;

      //ajout du produit de ADMCLIENT
      sql := 'SELECT ACL_CDKEY FROM ADMCLIENT WHERE ACL_IDHEBERGE="' + IDHeberge + '"'
        + ' AND ACL_IDCLIENT="' + IDClient + '"';
      traceExecution(Sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      if not q.eof then
      begin
        seria := q.Fields[0].AsString;
        if SS.DecodeKey(seria, Product, Users, Order) > 0 then
        begin
          result := IntToStr(Product);
        end;
      end;
      ferme(Q);

      //ajout des produits de ADMCDKEY
      sql := 'SELECT ACK_CDKEY FROM ADMCDKEY WHERE ACK_IDHEBERGE="' + IDHeberge + '"'
        + ' AND ACK_IDCLIENT="' + IDClient + '"';
      traceExecution(Sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      while not q.eof do
      begin
        seria := q.Fields[0].AsString;
        if SS.DecodeKey(seria, Product, Users, Order) > 0 then
        begin
          result := result + si(result <> '', ';', '') + IntToStr(Product);
        end;
        Q.next();
      end;
      ferme(Q);
      ss.free;

    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        if assigned(ss) then
          ss.free;
        ErrorMessage := 'getProduit: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINIBase.GetEnvironnement(WAEnvironnement : cWAEnvironnement): boolean;
var
  q: TQuery;
  sql: string;
  tADMENV: tob;
begin
  if ModeEagl then
  begin
    result := inherited GetEnvironnement(WAEnvironnement) ;
  end
  else
  begin
    result:=false ;
    q := nil;
    try
      setEntiteFromNomEntite;
      setGroupeFromNomGroupe;
      //24/05/07 Gestion Environnement sur le Client début
      sql := 'SELECT * FROM ADMENVIRONNEMENT WHERE AEN_IDENV='
        +' (SELECT ACL_IDENV FROM ADMCLIENT WHERE ACL_IDHEBERGE="'+ IDHeberge + '" AND ACL_IDCLIENT="'+IDClient+'")' ;
      traceExecution(Sql);
      q := DBOpenSQL(leDB, sql, DBDriver);
      if q.eof then
        begin
        Ferme(q) ;
        sql := 'SELECT * FROM ADMENVIRONNEMENT WHERE AEN_IDENV='
          +' (SELECT AHE_IDENV FROM ADMHEBERGE WHERE AHE_IDHEBERGE="'+ IDHeberge + '")' ;
        traceExecution(Sql);
        q := DBOpenSQL(leDB, sql, DBDriver);
        if q.eof then
          begin
          Ferme(q) ;
          //sql := 'SELECT * FROM ADMENVIRONNEMENT WHERE AEN_DEFAUT="X"' ; //24/05/2007 Gestion des environnements par défaut dans les ParamSoc
          //YCP pb de getparamsosql := 'SELECT * FROM ADMENVIRONNEMENT WHERE AEN_IDENV="'+GetParamSocSecur('SO_IDENVEXPLOIT','')+'"' ;
          sql:= 'SELECT * FROM ADMENVIRONNEMENT WHERE AEN_IDENV=(select soc_data from paramsoc where soc_nom="SO_IDENVEXPLOIT")' ;
            traceExecution(Sql);
          q := DBOpenSQL(leDB, sql, DBDriver);
          end ;
        end ;
      //24/05/07 Gestion Environnement sur le Client fin
      tADMENV := TOB.Create('ADMENVIRONNEMENT_s', nil, -1);
      tADMENV.LoadDetailDB('ADMENVIRONNEMENT', '', '', Q, false);
      result:=WAEnvironnement.setvaluesFromTob(tADMENV) ;
      tADMENV.free ;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        ErrorMessage := 'GetEnvironnement: ' + E.Message;
      end;
    end;
  end;
end ;


function cWAINIBase.GetMaxUserNumber: integer;
{$IFNDEF EAGLCLIENT}
var
  q: TQuery;
  seria, sql: string;
  ss: TPGISeriaServer;
  Product, Users, Order: LongWord;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result := inherited GetMaxUserNumber;
  end
  else
  begin
  result := -1;
{$IFNDEF EAGLCLIENT}
    q := nil;
    ss := nil;
    try
      setEntiteFromNomEntite;
      setGroupeFromNomGroupe;

      sql := 'SELECT ACL_CDKEY FROM ADMCLIENT WHERE ACL_IDHEBERGE="' + IDHeberge + '"'
        + ' AND ACL_IDCLIENT="' + IDClient + '"';
      traceExecution(Sql);
      q := DBOpenSQL(leDB, Sql, DBDriver);
      if not q.eof then
      begin
        seria := q.Fields[0].AsString;
        ss := TPGISeriaServer.Create;
        if SS.DecodeKey(seria, Product, Users, Order) > 0 then
        begin
          result := Users;
        end;
        ss.free;
      end;
      ferme(Q);
    except
      on E: Exception do
      begin
        if Assigned(q) then
          ferme(Q);
        if assigned(ss) then
          ss.free;
        ErrorMessage := 'GetMaxUserNumber: ' + E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{$ENDIF EAGLCLIENT}

function cWAINIBase.getNomBase((*wGamme,*) wEntite, wDossier: string): string;
var
  WAEnvironnement: cWAEnvironnement ;
begin
  IDHeberge:=wEntite ;
  WAEnvironnement:=cWAEnvironnement.create() ;
  if getEnvironnement(WAEnvironnement) then
    begin
    result:=WAenvironnement.PrefixeBase ;
    end ;
  WAEnvironnement.free ;

  (*WAIniBase:=cWAINIBase.create ;
  WAIniBase.IDHeberge:=wEntite ;
  WAEnvironnement:=cWAEnvironnement.create() ;
  if WAIniBase.getEnvironnement(WAEnvironnement) then
    begin
    result:=WAenvironnement.PrefixeBase ;
    end ;
  WAEnvironnement.free ;
  WAIniBase.free ;*)

  (*if wGamme = '001' then
    result := 'S1'
  else if wGamme = '002' then
    result := 'DB'
  else if wGamme = '003' then
    result := 'S5';   *)
  result := result + '#' + wEntite + '#' + wDossier;
end;

function cWAINIBase.getISCreate: boolean;
begin
  if (not self.FieldExists('ISCREATE')) then
    self.AddChampSupValeur('ISCREATE', '');
  result := (self.getString('ISCREATE') = 'X');
end;

procedure cWAINIBase.setIsCreate(valeur: boolean);
begin
  if (not self.FieldExists('ISCREATE')) then
    self.AddChampSup('ISCREATE', false);
  self.PutValue('ISCREATE', si(valeur, 'X', '-'));
end;


end.

