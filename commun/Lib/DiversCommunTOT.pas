Unit DiversCommunTOT ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     utob,
{$ELSE}
     db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     forms, sysutils,ComCtrls,HCtrls, HEnt1, HMsgBox, UTOT ;

Type
  TOT_YYDOMAINE = Class ( TOT )
    procedure OnDeleteRecord           ; override ;
    Public
    Function DetruitDomaine ( sDom : String ) : boolean ;
  end ;

Implementation

Function TOT_YYDOMAINE.DetruitDomaine ( sDom : String ) : Boolean ;
Var Trouver : boolean ;
BEGIN
Result:=False ; //Trouver:=False ;
Trouver:=ExisteSQL('SELECT GA_DOMAINE FROM ARTICLE WHERE GA_DOMAINE="'+sDom+'"') ;
if Trouver then BEGIN LastError:=1 ; Exit ; END ;
Trouver:=ExisteSQL('SELECT CPU_DOMAINE FROM CPPROFILUSERC WHERE CPU_DOMAINE="'+sDom+'"') ;
if Trouver then BEGIN LastError:=1 ; Exit ; END ;
Trouver:=ExisteSQL('SELECT GP_DOMAINE FROM PIECE WHERE GP_DOMAINE="'+sDom+'"') ;
if Trouver then BEGIN LastError:=1 ; Exit ; END ;
Result:=True ;
END ;

procedure TOT_YYDOMAINE.OnDeleteRecord ;
Var sDom : String ;
    OkDet : Boolean ;
begin
sDom:=GetField('CC_CODE') ; if sDom='' then Exit ;
OkDet:=DetruitDomaine(sDom) ;
if Not OkDet then BEGIN LastError:=1 ; LastErrorMsg:='Ce domaine est utilisé dans les paramétrages ou les pièces' ; END ;
  Inherited ;
end ;

Initialization
  registerclasses([TOT_YYDOMAINE]) ;
end.

