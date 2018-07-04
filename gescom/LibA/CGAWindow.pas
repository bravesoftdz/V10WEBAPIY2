unit CGAWindow;

interface

uses ComObj, ActiveX, CGAS5_TLB, HEnt1, HCtrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
 DB, DBtables,
{$ENDIF}
 StdVcl,OleDicoPGI, Forms, Utob, UtilOLE ;

type
  TCGAWindow = class(TAutoObject,ICGAWindow)
  protected
    function GARechDOMOLE(const TT, Code: WideString): OleVariant; safecall;
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

function TCGAWindow.GARechDOMOLE(const TT, Code: WideString): OleVariant;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result:=RechDom(TT,Code,False) ;
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
  TAutoObjectFactory.Create(ComServer, TCGAWindow, Class_CGAWindow, ciMultiInstance, tmApartment);
end.
