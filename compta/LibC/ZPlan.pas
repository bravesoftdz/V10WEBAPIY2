unit ZPlan;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// GENERAUX : G_GENERAL

interface

uses SysUtils,
  {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
     UTob, HCtrls,
     Ent1,
     TZ ;

type
  TZPlan = class
  private
  Plan : TZF ;
  function    GetCount : Integer ;
  public
  OuvrePerte, OuvreBenef, FermePerte, FermeBenef : string ;
  constructor Create ;
  destructor  Destroy ; override ;
  function    Load(bAux, bRes, bColl : Boolean) : Boolean ;
  function    FindCompte(var NumCompte : string) : Integer ;
  function    GetValue(Nom : string; Niveau : integer) : Variant ;
  property    Count : Integer read GetCount ;
  end ;

implementation


uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus, ParamSoc; 

constructor TZPlan.Create ;
begin
Plan:=TZF.Create('VCOMPTE', nil, -1) ;
end ;

destructor TZPlan.Destroy ;
begin
if Plan<>nil then Plan.Free ;
end ;

function TZPlan.Load(bAux, bRes, bColl : Boolean) : Boolean ;
var Q : TQuery ; Cols, Table, Where, OrderBy : string ;
begin
Cols:='' ; Table:='' ; Where:='' ; OrderBy:='' ;
if bColl then
  begin
  Table:='GENERAUX' ; OrderBy:='G_GENERAL' ;
  Cols:='G_GENERAL, G_COLLECTIF, G_NATUREGENE' ;
  Where:='WHERE G_COLLECTIF="X" AND G_NATUREGENE<>"EXT"' ;
  end else
  if bAux then
    begin
    Table:='TIERS' ; OrderBy:='T_COLLECTIF, T_AUXILIAIRE' ;
    Cols:='T_AUXILIAIRE , T_LIBELLE, T_COLLECTIF, T_NATUREAUXI, T_LETTRABLE, T_MODEREGLE' ;
    end else
    begin
    Table:='GENERAUX' ; OrderBy:='G_GENERAL' ;
    Cols:='G_GENERAL, G_LIBELLE, G_COLLECTIF, G_NATUREGENE, G_LETTRABLE, G_POINTABLE, G_MODEREGLE' ;
    Where:='WHERE G_NATUREGENE<>"EXT"' ;
    end ;
if (not bAux) and (bRes) then
   begin
   FermePerte:=GetParamSoc('SO_FERMEPERTE') ;
   OuvrePerte:=GetParamSoc('SO_OUVREPERTE') ;
   FermeBenef:=GetParamSoc('SO_FERMEBEN') ;
   OuvreBenef:=GetParamSoc('SO_OUVREBEN') ;
   end ;
{ FQ 21777 BVE 06.11.07 }
Where := Where + ' AND G_FERME <> "X" ';
{ END FQ 21777 }
Q:=OpenSQL('SELECT '+Cols+' FROM '+Table+' '+Where+'ORDER BY '+OrderBy, TRUE);
Result:=not Q.EOF ; if not Result then begin Ferme(Q) ; Exit ; end ;
if bAux then Result:=Plan.LoadDetailDB('VTIERS', '', '', Q, TRUE)
        else Result:=Plan.LoadDetailDB('VGENERAUX', '', '', Q, TRUE) ;
Ferme(Q) ;
end ;

(* ANCIEN CODE
function TZPlan.Load(bAux, bRes : Boolean) : Boolean ;
var Q, QSoc : TQuery ; Col, ColAux, Join, Where : string ;
begin
//'SELECT G_GENERAL, G_LIBELLE, T_AUXILIAIRE, T_LIBELLE FROM GENERAUX LEFT JOIN TIERS ON G_GENERAL=T_COLLECTIF ORDER BY G_GENERAL'
ColAux:='' ; Join:='' ; Where:='' ;
Col:='G_GENERAL, G_LIBELLE, G_COLLECTIF, G_NATUREGENE' ;
if bAux then
   begin
//   ColAux:=', T_AUXILIAIRE AS G_NIF, T_LIBELLE AS G_ADRESSE1' ;
   ColAux:=', T_AUXILIAIRE , T_LIBELLE' ;
   Join:='LEFT JOIN TIERS ON G_GENERAL=T_COLLECTIF ' ;
   end ;
if not bRes then
   begin
   QSoc:=OpenSQL('SELECT SO_FERMEPERTE, SO_OUVREPERTE, SO_FERMEBEN, SO_OUVREBEN FROM SOCIETE', TRUE) ;
   FermePerte:=QSoc.Fields[0].AsString ;
   OuvrePerte:=QSoc.Fields[1].AsString ;
   FermeBenef:=QSoc.Fields[2].AsString ;
   OuvreBenef:=QSoc.Fields[3].AsString ;
   Ferme(QSoc) ;
   Where:='WHERE G_GENERAL<>"'+FermePerte+'"'
          +' AND G_GENERAL<>"'+OuvrePerte+'"'
          +' AND G_GENERAL<>"'+FermeBenef+'"'
          +' AND G_GENERAL<>"'+OuvreBenef+' "' ;
   end ;
Q:=OpenSQL('SELECT '+Col+ColAux+' FROM GENERAUX '+Join+Where+'ORDER BY G_GENERAL', TRUE);
Result:=not Q.EOF ; if not Result then begin Ferme(Q) ; Exit ; end ;
Result:=Plan.LoadDetailDB('VGENERAUX', '', '', Q, FALSE) ;
Ferme(Q) ;
end ;
*)

function TZPlan.GetCount : Integer ;
begin
Result:=Plan.Detail.Count ;
end ;

function TZPlan.FindCompte(var NumCompte : string) : Integer ;
var i : integer ; Col : string ;
begin
Result:=-1 ; Col:='' ;
NumCompte:=Trim(NumCompte) ; if NumCompte='' then Exit ;
NumCompte:=BourreLaDonc(NumCompte, fbGene) ;
if Plan.Detail.Count>0 then
  if Plan.Detail[0].NomTable='VTIERS'
    then Col:='T_AUXILIAIRE'
    else Col:='G_GENERAL' ;
for i:=0 to Plan.Detail.Count-1 do
  if TZF(Plan.Detail[i]).GetValue(Col)=NumCompte then
     begin Result:=i ; Exit ; end ;
end ;

function TZPlan.GetValue(Nom : string; Niveau : integer) : Variant ;
begin
Result:=TZF(Plan.Detail[Niveau]).GetValue(Nom) ;
end ;

end.
