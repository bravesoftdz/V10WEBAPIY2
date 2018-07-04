unit GCOLE;

interface

uses ComObj, ActiveX, HEnt1, HCtrls, CGS5_TLB, DB, DBtables, StdVcl ;

type
  TGCWINDOW = class(TAutoObject,IGCWINDOW)
  protected
    function  GCCumulPiece(const NomChamp,Natures,Etabs,Devises,Repres,Depots: WideString; DateDeb,DateFin: TDateTime; const Arg: WideString): OleVariant; safecall;
    function  GCCumulLigne(const NomChamp,Natures,Etabs,Devises,Repres,Depots: WideString; DateDeb,DateFin: TDateTime; const Arg: WideString): OleVariant; safecall;
    function  GCRechDOMOLE(const TT, Code: WideString): OleVariant; safecall;
    function  GCCumulArticle(const NomChamp, Natures, Repres, FM1, FM2, FM3,Types, Statuts: WideString; DateDeb, DateFin: TDateTime; const Arg: WideString): OleVariant; safecall;
    procedure OnRelease (var shutdown : boolean) ;
    procedure GCZoom(const Table: WideString; const Qui: WideString); safecall;
    function  GCGetIDCourant: OleVariant; safecall;
    function  GCTransformePiece(OldNat: OleVariant; NewNat: OleVariant; Numero: OleVariant;
                                Souche: OleVariant; Indice: OleVariant): OleVariant; safecall;
  public
    procedure initialize ; override ;
  end;

implementation

uses ComServ, Dialogs, Forms, UTom, FichList, UtomLiensOle, FactGrp;

procedure TGCWINDOW.OnRelease(var shutdown: boolean);
begin
ShutDown:=FALSE ;
end;

procedure TGCWINDOW.initialize;
begin
  inherited initialize ;
  ComServer.onlastRelease:=OnRelease ;

end;                

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
   if PasWhere then BEGIN Result:=' WHERE ' ; PasWhere:=False ; END else Result:=' AND ' ;
   Result:=Result+StLoc ;
   END ;
END ;

Procedure VarSQLWhere ( Var SQL : String ; Var PasWhere : boolean ) ;
BEGIN
if PasWhere then BEGIN SQL:=SQL+' WHERE ' ; PasWhere:=False ; END else SQL:=SQL+' AND ' ;
END ;

function TGCWINDOW.GCRechDOMOLE(const TT, Code: WideString): OleVariant;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result:=RechDom(TT,Code,False) ;
end;

function TGCWINDOW.GCCumulPiece(const NomChamp, Natures, Etabs, Devises,Repres, Depots: WideString; DateDeb, DateFin: TDateTime;const Arg: WideString): OleVariant;
Var Q   : TQuery ;
    SQL : String ;
    PasWhere : Boolean ;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
SQL:='Select SUM('+NomChamp+') FROM PIECE' ;
PasWhere:=True ;
SQL:=SQL+CompleteSQLOLE(Natures,'GP_NATUREPIECEG',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Etabs,'GP_ETABLISSEMENT',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Devises,'GP_DEVISE',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Repres,'GP_REPRESENTANT',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Depots,'GP_DEPOT',PasWhere) ;
if Arg<>''   then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+Arg ; END ;
if DateDeb>2 then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+'GP_DATEPIECE>="'+USDateTime(DateDeb)+'"' ; END ;
if DateFin>2 then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+'GP_DATEPIECE<="'+USDateTime(DateFin)+'"' ; END ;
Q:=OpenSQL(SQL,True) ; if Not Q.EOF then Result:=Q.Fields[0].AsVariant ; Ferme(Q) ;
end;

function TGCWINDOW.GCCumulLigne(const NomChamp,Natures,Etabs,Devises,Repres,Depots: WideString; DateDeb,DateFin: TDateTime;const Arg: WideString): OleVariant;
Var Q : TQuery ;
    SQL : String ;
    PasWhere : Boolean ;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
SQL:='Select SUM('+NomChamp+') FROM LIGNE' ;
PasWhere:=True ;
SQL:=SQL+CompleteSQLOLE(Natures,'GL_NATUREPIECEG',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Etabs,'GL_ETABLISSEMENT',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Devises,'GL_DEVISE',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Repres,'GL_REPRESENTANT',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Depots,'GL_DEPOT',PasWhere) ;
if Arg<>''   then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+Arg ; END ;
if DateDeb>2 then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+'GL_DATEPIECE>="'+USDateTime(DateDeb)+'"' ; END ;
if DateFin>2 then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+'GL_DATEPIECE<="'+USDateTime(DateFin)+'"' ; END ;
if PasWhere then SQL:=SQL+' WHERE GL_TYPELIGNE="ART"' else SQL:=SQL+' AND GL_TYPELIGNE="ART"' ;
Q:=OpenSQL(SQL,True) ; if Not Q.EOF then Result:=Q.Fields[0].AsVariant ; Ferme(Q) ;
end;

function TGCWINDOW.GCCumulArticle(const NomChamp, Natures, Repres, FM1, FM2, FM3, Types, Statuts: WideString; DateDeb, DateFin: TDateTime; const Arg: WideString): OleVariant;
Var Q : TQuery ;
    SQL : String ;
    PasWhere : Boolean ;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
SQL:='Select SUM('+NomChamp+') FROM LIGNE LEFT JOIN ARTICLE ON GL_ARTICLE=GA_ARTICLE' ;
PasWhere:=True ;
SQL:=SQL+CompleteSQLOLE(Natures,'GL_NATUREPIECEG',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Repres,'GL_REPRESENTANT',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(FM1,'GA_FAMILLENIV1',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(FM2,'GA_FAMILLENIV2',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(FM3,'GA_FAMILLENIV3',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Types,'GA_TYPEARTICLE',PasWhere) ;
SQL:=SQL+CompleteSQLOLE(Statuts,'GA_STATUTART',PasWhere) ;
if Arg<>''   then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+Arg ; END ;
if DateDeb>2 then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+'GL_DATEPIECE>="'+USDateTime(DateDeb)+'"' ; END ;
if DateFin>2 then BEGIN VarSQLWhere(SQL,PasWhere) ; SQL:=SQL+'GL_DATEPIECE<="'+USDateTime(DateFin)+'"' ; END ;
if PasWhere then SQL:=SQL+' WHERE GL_TYPELIGNE="ART"' else SQL:=SQL+' AND GL_TYPELIGNE="ART"' ;
Q:=OpenSQL(SQL,True) ; if Not Q.EOF then Result:=Q.Fields[0].AsVariant ; Ferme(Q) ;
end;

procedure TGCWINDOW.GCZoom(const Table: WideString; const Qui: WideString);
var SQL,Champ : string ;
begin
   if Table='COMMERCIAL' then  begin Champ:='GCL_COMMERCIAL'; end ;
   SQL:= 'SELECT '+Champ+' from '+Table+' WHERE '+Champ+'="'+Qui+'"' ;
   If ExisteSQL(SQL) then
      begin
      InitZoomOLE ;
      if Table='COMMERCIAL' then V_PGI.DispatchTT (9,taConsult ,Qui, '','');
      FinZoomOLE ;
      end ;
end;

function TGCWINDOW.GCGetIDCourant: OleVariant;
var F :TForm ;
    OM :TOM ;
Begin
result:='';
F:=TForm(Screen.ActiveForm) ;
if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit ;
if (OM is TOM_LiensOle) then result:=TOM_LiensOle(OM).GetIDCourant else exit ;
end;

function TGCWINDOW.GCTransformePiece(OldNat: OleVariant; NewNat: OleVariant; Numero: OleVariant;
                                     Souche: OleVariant; Indice: OleVariant): OleVariant; 
BEGIN
Result:=TransfoBatchPiece(OldNat,NewNat,Souche,Numero,Indice) ; 
END ;

initialization
  TAutoObjectFactory.Create(ComServer, TGCWINDOW, Class_GCWINDOW, ciMultiInstance, tmApartment);
end.
