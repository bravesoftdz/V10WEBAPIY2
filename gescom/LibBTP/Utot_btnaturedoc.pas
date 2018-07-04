{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/06/2002
Modifié le ... :   /  /
Description .. : Source TOT de la TABLETTE : BTNATUREDOC (BTNATUREDOC)
Mots clefs ... : TOT;BTNATUREDOC
*****************************************************************}
Unit Utot_btnaturedoc ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     UTOT,
     Dicobtp,
     UtilPaieAffaire
      ;

Type
  TOT_BTNATUREDOC = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ; 
  end ;

Implementation

procedure TOT_BTNATUREDOC.OnNewRecord ;
begin
  Inherited ;
SetField ('CO_LIBRE','SECOND');
end ;

procedure TOT_BTNATUREDOC.OnDeleteRecord ;
var code,lib,stit,Libre: string;
    CodeUse : boolean;
BEGIN
inherited;
  Code := GetField('CO_CODE');
  Libre := GetField('CO_Libre');
  Lib :=  RechDom('BTNATUREDOC',Code,False);
  stit := format('Nature de document : %s - %s',[Code,Lib]);
if (Libre = 'PRINC') or (Libre = 'PRINC1') then
    Begin
    PGIBoxAF('Suppression interdite, cette nature de document est obligatoire',stit);
    Lasterror:=1;
    exit;
    End;
CodeUse := SupTablesLiees ('BDETETUDE', 'BDE_NATURE', Code, '' , false);
if CodeUse then
   Begin
   PGIBoxAF('Suppression interdite, cette nature de document est utilisé',stit);
   Lasterror:=1; Exit;
   End;

end ;

procedure TOT_BTNATUREDOC.OnUpdateRecord ;
var Libre : string;
    poset : integer;
begin
  Inherited ;
  Libre := GetField('CO_LIBELLE');
  repeat
  poset := pos ('&',libre);
  if poset = 1 then libre := copy (libre,2,255)
  else if poset > 1 then libre := copy(libre,1,poset-1)+copy (libre,poset+1,255);
  until poset <= 0;
  Setfield ('CO_LIBELLE',libre);
end ;

procedure TOT_BTNATUREDOC.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOT_BTNATUREDOC.OnClose ;
begin
  Inherited ;
end ;

procedure TOT_BTNATUREDOC.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOT_BTNATUREDOC ] ) ;
end.

