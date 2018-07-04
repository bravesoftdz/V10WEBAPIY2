unit uTOBActive;

interface

uses
  ComObj, ActiveX, AspTlb, ETOBACTIVE_TLB, StdVcl, SysUtils, eSession, eISAPI, MajTable ;

type
  TTOBActive = class(TASPObject, ITOBActive)
  private

  protected
    procedure OnEndPage; safecall;
    procedure OnStartPage(const AScriptingContext: IUnknown); safecall;
    function  EGETTARIF ( FichierIni,NomSoc,CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte,RemiseTiers: OleVariant;
                          var LePrix,LaRemise: OleVariant): OleVariant; safecall;
    function  EGETTAUXTVA(FichierIni, NomSoc, Regime, Code: OleVariant): OleVariant; safecall;
    procedure EGetDBParams(FichierIni, NomSoc: OleVariant; var StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions: OleVariant); safecall;
    function  EGETTARIF_WP ( CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte,RemiseTiers: OleVariant;
                             var LePrix,LaRemise: OleVariant;
                             StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions: OleVariant): OleVariant; safecall;
    procedure EGetInfosDossier ( FichierIni,NomSoc: OleVariant; var DevisePivot,TenueEuro,DepotDefaut: OleVariant ); safecall;
  end;

implementation

uses ComServ, HEnt1, DBTables, HCtrls, UtilTarif ;

Var FF : TextFile ;

procedure TTOBActive.OnEndPage;
begin
  inherited OnEndPage;
end;

procedure TTOBActive.OnStartPage(const AScriptingContext: IUnknown);
begin
  inherited OnStartPage(AScriptingContext);
end;

function TTOBActive.EGETTARIF ( FichierIni,NomSoc,CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte,RemiseTiers: OleVariant;
                                var LePrix,LaRemise: OleVariant ): OleVariant;
Var pSession: tISession;
begin
pSession:=OpenComSession ;
HalSocIni:=FichierINI ; ConnecteDB(NomSoc,pSession.LaDB,NomSoc) ;
ChargeTablePrefixe(FALSE,TRUE) ;
Result:=ChercheWEBtarif(CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte,RemiseTiers,LePrix,LaRemise) ;
CloseComSession(pSession) ;
end;

function TTOBActive.EGETTAUXTVA (FichierIni, NomSoc, Regime, Code : OleVariant ): OleVariant ;
Var QQ : TQuery ;
    pSession: tISession;
begin
pSession:=OpenComSession ;
HalSocIni:=FichierINI ; ConnecteDB(NomSoc,pSession.LaDB,NomSoc) ;
ChargeTablePrefixe(FALSE,TRUE) ;
Result:=0 ;
QQ:=OpenSQL('SELECT TV_TAUXVTE FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_REGIME="'+Regime+'" AND TV_CODETAUX="'+Code+'"',True) ;
if Not QQ.EOF then Result:=QQ.Fields[0].AsVariant ;
Ferme(QQ) ;
CloseComSession(pSession) ;
end;

Function ConnecteDBWithParams ( StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions : String ; DB : TDataBase ) : boolean ;
Var DBDriver : TDBDriver ;
    OldFork : Boolean ;
    i       : integer ;
BEGIN
Result:=FALSE ; OldFork:=V_PGI.StandardSurDP ; V_PGI.StandardSurDP:=False ;
 Try
  DBDriver:=AssignDBParamsStrings('',StDriver,StServer,StPath,StBase,StUser,StPassWord,StODBC,StOptions,DB) ;
  DB.Exclusive:=False ;
  if V_PGI.PrivateBDEDir<>'' then DB.Session.PrivateDir:=V_PGI.PrivateBDEDir ;
  if V_PGI.NetBDEDir<>'' then DB.Session.NetFileDir:=V_PGI.NetBDEDir ;
  DB.Connected:=TRUE ;
  V_PGI.Driver:=DBDriver ; MajOracleSession(DB) ;
  Result:=TRUE ;
 Except
  On E : Exception do BEGIN V_PGI.StandardSurDP:=OldFork ; Result:=FALSE ; END ;
 End ;
V_PGI.StandardSurDP:=OldFork ;
END ;

procedure TTOBActive.EGetDBParams ( FichierIni,NomSoc : OleVariant; Var StDriver, StServer, StPath, StODBC, StBase, StUser, StPassWord, StOptions: OleVariant) ;
Var pSession: tISession;
    sNomSoc,sStDriver,sStServer,sStPath,sStODBC,sStBase,sStUser,sStPassWord,sStOptions : String ;
begin
pSession:=OpenComSession ;
HalSocIni:=FichierINI ; sNomSoc:=NomSoc ;
ChargeDBParams(sNomSoc,sStDriver,sStServer,sStPath,sStbase,sStUser,sStPassWord,sStODBC,sStOptions) ;
StDriver:=sStDriver ; StServer:=sStServer ; StPath:=sStPath         ; StODBC:=sStODBC ;
StBase:=sStBase     ; StUser:=sStUser     ; StPassWord:=sStPassWord ; StOptions:=sStOptions ;
CloseComSession(pSession) ;
end;

procedure TTOBActive.EGetInfosDossier ( FichierIni,NomSoc: OleVariant; var DevisePivot,TenueEuro,DepotDefaut: OleVariant);
Var pSession: tISession;
    Q         : TQuery ;
    StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions : OLEVariant ;
begin
EGetDBParams(FichierIni,NomSoc,StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions) ;
pSession:=OpenComSession ;
ConnecteDBWithParams(StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions,pSession.LaDB) ;
ChargeTablePrefixe(FALSE,TRUE) ;
Q:=OpenSQL('SELECT SOC_NOM, SOC_DATA FROM PARAMSOC WHERE (SOC_NOM="SO_TENUEEURO" OR SOC_NOM="SO_DEVISEPRINC" OR SOC_NOM="SO_GCDEPOTDEFAUT")',True) ;
While Not Q.EOF do
   BEGIN
   if Q.Fields[0].AsString='SO_TENUEEURO' then TenueEuro:=Q.Fields[1].AsString else
    if Q.Fields[0].AsString='SO_GCDEPOTDEFAUT' then DepotDefaut:=Q.Fields[1].AsString
       else DevisePivot:=Q.Fields[1].AsString ;
   Q.Next ;
   END ;
Ferme(Q) ;
CloseComSession(pSession) ;
end;

function TTOBActive.EGETTARIF_WP ( CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte,RemiseTiers : OleVariant;
                                   var LePrix,LaRemise : OleVariant;
                                   StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions: OleVariant): OleVariant;
Var pSession: tISession;
begin
pSession:=OpenComSession ;
ConnecteDBWithParams(StDriver,StServer,StPath,StODBC,StBase,StUser,StPassWord,StOptions,pSession.LaDB) ;
ChargeTablePrefixe(FALSE,TRUE) ;
Result:=ChercheWEBtarif(CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte,RemiseTiers,LePrix,LaRemise) ;
CloseComSession(pSession) ;
end;

initialization
TAutoObjectFactory.Create(ComServer, TTOBActive, Class_TOBActive,ciMultiInstance, tmApartment);
end.

