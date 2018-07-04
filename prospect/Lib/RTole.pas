unit RTOLE;

interface

uses ComObj, ActiveX, CRCS5_TLB, HEnt1, HCtrls, DB, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} StdVcl,OleDicoPGI, Forms, Utob ;

type
  TGRCWINDOW = class(TAutoObject,IGRCWINDOW)
  protected
    function GRCRechDOMOLE(const TT, Code: WideString): OleVariant; safecall;
    function GRCCumulPersp(const etat: WideString; datcredeb: TDateTime; datcrefin: TDateTime;
                            datfinviedeb: TDateTime; datfinviefin: TDateTime): OleVariant; safecall;
    procedure GRCInitDico; safecall;
    procedure GRCNewDico;  safecall;
    procedure GRCShowDico; safecall;
    procedure GRCFreeDico; safecall;
  end;

procedure GetTobDico(theTobDico: TOB);

implementation

uses ComServ, Dialogs;

Function DecodeToken ( StOrig,Champ : String ) : String ;
Var ii : integer ;
    St,StLoc,StRes : String ;
BEGIN
Result:='' ;
St:=StOrig ; StRes:='' ; ii:=0 ;
Repeat
 StLoc:=ReadTokenSt(St) ;
 if StLoc<>'' then BEGIN StRes:=StRes+Champ+'="'+StLoc+'" OR ' ; Inc(ii) ; END ;
Until ((St='') or (StLoc='')) ;
if StRes<>'' then Delete(StRes,Length(StRes)-3,4) ;
if ii>1 then StRes:='('+Stres+')' ;
Result:=StRes ;
END ;

Function CompleteSQLOLE ( StVals,Champ : String ; Var PasWhere : Boolean ) : String ;
Var StLoc : String ;
BEGIN
Result:='' ;
if StVals<>'' then
   BEGIN
   StLoc:=DecodeToken(StVals,Champ) ;
   if PasWhere then
   BEGIN
   Result:=' WHERE ' ;
   PasWhere:=False ;
   END
   else
       Result:=' AND ' ;
   Result:=Result+StLoc ;
   END ;
END ;

Procedure VarSQLWhere ( Var SQL : String ; Var PasWhere : boolean ) ;
BEGIN
if PasWhere then
   BEGIN
   SQL:=SQL+' WHERE ' ;
   PasWhere:=False ;
   END
   else SQL:=SQL+' AND ' ;
END ;

function TGRCWINDOW.GRCRechDOMOLE(const TT, Code: WideString): OleVariant;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result:=RechDom(TT,Code,False) ;
end;

function TGRCWINDOW.GRCCumulPersp(const etat: WideString; datcredeb: TDateTime; datcrefin: TDateTime;
                            datfinviedeb: TDateTime; datfinviefin: TDateTime): OleVariant;
Var Q   : TQuery ;
    SQL : String ;
    PasWhere : Boolean ;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;

SQL:='Select COUNT (RPE_ETATPER) FROM PERSPECTIVES' ;
PasWhere:=True ;
SQL:=SQL+CompleteSQLOLE(etat,'RPE_ETATPER',PasWhere) ;

VarSQLWhere(SQL,PasWhere) ;
SQL:=SQL+'RPE_DATECREATION>="'+USDateTime(datcredeb)+'"' ;

VarSQLWhere(SQL,PasWhere) ;
SQL:=SQL+'RPE_DATECREATION<="'+USDateTime(datcrefin)+'"' ;

VarSQLWhere(SQL,PasWhere) ;
SQL:=SQL+'RPE_DATEFINVIE>="'+USDateTime(datfinviedeb)+'"' ;

VarSQLWhere(SQL,PasWhere) ;
SQL:=SQL+'RPE_DATEFINVIE<="'+USDateTime(datfinviefin)+'"' ;

Q:=OpenSQL(SQL,True) ; if Not Q.EOF then Result:=Q.Fields[0].AsVariant ; Ferme(Q) ;
end;

procedure TGRCWINDOW.GRCFreeDico;
begin
FreeDictionnairePGI;
end;

procedure TGRCWINDOW.GRCInitDico;
begin
//InitDictionnairePGI(GetTobDico,GetTobFunc,ValidateScript);
InitDictionnairePGI(GetTobDico,Nil,Nil);
end;

procedure TGRCWINDOW.GRCNewDico;
begin
NewDictionnairePGI;
end;

procedure TGRCWINDOW.GRCShowDico;
begin
Application.Minimize;
AfficherDictionnairePGI;
end;

Procedure InitTobDico( Nom, Prefixe : string; TheTobDico : Tob);
var T1 : Tob;
    Q : TQuery;
begin
Q:=OpenSql('SELECT * FROM DECHAMPS WHERE (DH_PREFIXE="'+Prefixe+'") AND (DH_CONTROLE like "%D%")',True);
T1 := TOB.Create(Nom,theTobDico,-1);
T1.LoadDetailDB('DECHAMPS','','DH_NUMCHAMP',Q,false);
Ferme(Q);
end;

procedure GetTobDico(theTobDico: TOB);
begin
theTobDico.ClearDetail;
InitTobDico('Tiers','T',theTobDico);
InitTobDico('Adresse','ADR',theTobDico);
InitTobDico('Contact','C',theTobDico);
InitTobDico('Champs libres','YTC',theTobDico);
InitTobDico('Informations complémentaires','RPR',theTobDico);
end;

initialization
  TAutoObjectFactory.Create(ComServer, TGRCWINDOW, Class_GRCWINDOW, ciMultiInstance, tmApartment);
end.
