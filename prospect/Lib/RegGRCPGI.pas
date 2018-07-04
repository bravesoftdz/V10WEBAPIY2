unit RegGRCPGI;

interface

uses OLEAuto,OleDicoPGI,UTob;

Type
// Automation Server
  TOleAutoGRCPGI = class(TAutoObject)
  automated
// Dico PGI
    function ShowDico : variant;
    function Free : variant;
    function New : variant;
    function Init : variant;
// GRC
    function GRCRechDOMOLE(const TT, Code: WideString): variant;
    function GRCCumulPersp(const etat: WideString; datcredeb: TDateTime; datcrefin: TDateTime;
                            datfinviedeb: TDateTime; datfinviefin: TDateTime): variant;
  end;

procedure RegisterGRCPGI ;

// fonction utilisateurs

procedure GetTobFunc ( theTobFunc : TOB {a ne pas vider : a utiliser comme parent});
procedure GetTobDico ( theTobDico : TOB);

implementation

uses
  {$IFDEF VER150}
    Variants,
  {$ENDIF VER150}
  HEnt1,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}hctrls,UtilOle
  ;

procedure RegisterGRCPGI ;
const
  AutoClassInfo: TAutoClassInfo = (
    AutoClass: TOleAutoGRCPGI;
    ProgID: 'CegidPgi.GRCPgi';
// ATTENTION : CE NUMERO DE CLASSE ID EST A CHANGER
// IL DOIT LIE A UNE APPLICATION ET A UNE SEULE
// POUR CALCULER UN GUID UNIQUE, FAIRE DANS DELPHI :
// CTRL+SHIFT+G
    ClassID: '{4B2646C8-952A-43F4-AACA-0D72F1450696}';
    Description: 'Cegid GRC PGI';
    Instancing: acMultiInstance);
begin
  Automation.RegisterClass(AutoClassInfo);
end;


{ TOleAutoGRCPGI }

function TOleAutoGRCPGI.ShowDico : variant;
begin
AfficherDictionnairePGI;
result := unassigned;
end;

function TOleAutoGRCPGI.Free : variant;
begin
FreeDictionnairePGI;
result := unassigned;
end;

function TOleAutoGRCPGI.New : variant;
begin
NewDictionnairePGI;
result := unassigned;
end;

function TOleAutoGRCPGI.Init : variant;
begin
InitDictionnairePGI(GetTobDico,GetTobFunc,nil,'');
result := unassigned;
end;

{ Exemples d'initialisation des tob champs et fonctions }


procedure GetTobDico ( theTobDico : TOB);
begin
theTobDico.ClearDetail;
InitTobDico('Tiers','T',theTobDico);
InitTobDico('Adresse','ADR',theTobDico);
InitTobDico('Contact','C',theTobDico);
InitTobDico('Champs libres','YTC',theTobDico);
InitTobDico('Informations complémentaires','RPR',theTobDico);
end;

procedure GetTobFunc ( theTobFunc : TOB {a ne pas vider : a utiliser comme parent});
var T,TT : TOB;
begin
// pour cacher l'onglet des fonctions, sortir APRES AVOIR VIDER LE DETAIL
// theTobFunc.ClearDetail;
// exit;

  TT := TOB.Create('Spécifiques',theTobFunc,-1);
// créer une tob fille pour chaque fonction
  T := TOB.Create('ExisteLien',TT,-1);
//??  T.AddchampSupValeur('TYPE','');
  T.AddchampSupValeur('NOM','EXISTELIEN');  // le nom de ta fonction
  T.AddchampSupValeur('PROTO','@EXISTELIEN ( Quoi , Lequel )'); // le prototype : ce qui est ajouté au script
  T.AddchampSupValeur('HELP','Permet de savoir si un lien existe'); // aide hint sur click
  ////
  T := TOB.Create('Fonction1',TT,-1);
//??  T.AddchampSupValeur('TYPE','');
  T.AddchampSupValeur('NOM','FONCTION1');
  T.AddchampSupValeur('PROTO','@FONCTION1 ( Lequel )');
  T.AddchampSupValeur('HELP','N''importe quoi ce test...');
//
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

function TOleAutoGRCPGI.GRCRechDOMOLE(const TT, Code: WideString): Variant;
begin
Result:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result:=RechDom(TT,Code,False) ;
end;

function TOleAutoGRCPGI.GRCCumulPersp(const etat: WideString; datcredeb: TDateTime; datcrefin: TDateTime;
                            datfinviedeb: TDateTime; datfinviefin: TDateTime): Variant;
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

Q:=OpenSQL(SQL,True,-1,'',true) ; if Not Q.EOF then Result:=Q.Fields[0].AsVariant ; Ferme(Q) ;
end;


end.
