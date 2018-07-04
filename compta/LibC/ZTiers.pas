unit ZTiers;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// TIERS    : T_AUXILIAIRE

interface

uses SysUtils,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     {$ENDIF}
     HCtrls,
     UTob;

type
  TZTiers = class
  private
  Tiers : TOB ;
  function    GetCount : Integer ;
  public
  constructor Create ;
  destructor  Destroy ; override ;
  function    GetCompte(var NumCompte : string; bFindByLib: Boolean=FALSE) : Integer ;
  function    GetValue(Nom : string; Niveau : integer) : Variant ;
  property    Count : Integer read GetCount ;
  end ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  Ent1;

constructor TZTiers.Create ;
begin
Tiers:=TOB.Create('VTIERS', nil, -1) ;
end ;

destructor TZTiers.Destroy ;
begin
if Tiers<>nil then Tiers.Free ;
end ;

function TZTiers.GetCount : Integer ;
begin
Result:=Tiers.Detail.Count ;
end ;

function TZTiers.GetCompte(var NumCompte : string; bFindByLib: Boolean(*=FALSE*)) : Integer ;
var Q: TQuery; i : integer ; sNumCompte : string ;
begin
Result:=-1 ;
NumCompte:=Trim(NumCompte) ; if NumCompte='' then Exit ;
if not bFindByLib then NumCompte:=BourreLaDonc(NumCompte, fbAux) ;
for i:=0 to Tiers.Detail.Count-1 do
  if Tiers.Detail[i].GetValue('T_AUXILIAIRE')=UpperCase(NumCompte) then begin Result:=i ; Exit ; end ;
if bFindByLib then
  begin
  for i:=0 to Tiers.Detail.Count-1 do
    if Tiers.Detail[i].GetValue('T_LIBELLE')=NumCompte then
      begin Result:=i ; NumCompte:=Tiers.Detail[i].GetValue('T_AUXILIAIRE') ; Exit ; end ;
  end ;
if Tiers.LoadDetailDB('TIERS', '"'+UpperCase(NumCompte)+'"', '', nil, TRUE) then
  begin Result:=Tiers.Detail.Count-1 ; Exit ; end ;
if bFindByLib then
  begin
  Q:=OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_LIBELLE LIKE "'+UpperCase(NumCompte)+'%"', TRUE) ;
  if not Q.EOF then
    begin
    sNumCompte:=Q.Fields[0].AsString ;
    Q.Next ;
    if Q.EOF then
      begin
      // on relance la recherche pour savoir s'il existait pas deja
      for i:=0 to Tiers.Detail.Count-1 do
       if Tiers.Detail[i].GetValue('T_AUXILIAIRE')=UpperCase(sNumCompte) then
        begin NumCompte:=sNumCompte ; Result:=i ; Exit ; end ;
      if Tiers.LoadDetailDB('TIERS', '"'+UpperCase(sNumCompte)+'"', '', nil, TRUE) then
        begin
        Result:=Tiers.Detail.Count-1 ;
        NumCompte:=Tiers.Detail[Tiers.Detail.Count-1].GetValue('T_AUXILIAIRE') ;
        end ;
      end ;
    end ;
  Ferme(Q) ;
  end ;
end ;

function TZTiers.GetValue(Nom : string; Niveau : integer) : Variant ;
begin
Result:=Tiers.Detail[Niveau].GetValue(Nom) ;
end ;

end.
