Unit UTotRegimeTva ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     utob,
{$ELSE}
     db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     forms, sysutils,ComCtrls,HCtrls, HEnt1, HMsgBox, UTOT ;

Type
  TOT_TTREGIMETVA = Class ( TOT )
    procedure OnDeleteRecord           ; override ;
    Public
    Function DetruitReg ( CodeReg : String ) : boolean ;
  end ;

Implementation

Function TOT_TTREGIMETVA.DetruitReg ( CodeReg : String ) : Boolean ;
Var Trouver : boolean ;
BEGIN
Result:=False ;
Trouver:=ExisteSQL('SELECT T_REGIMETVA FROM TIERS WHERE T_REGIMETVA="'+CodeReg+'"') ;
if Trouver then BEGIN LastError:=1 ; Exit ; END ;
Trouver:=ExisteSQL('SELECT G_REGIMETVA FROM GENERAUX WHERE G_REGIMETVA="'+CodeReg+'"') ;
if Trouver then BEGIN LastError:=1 ; Exit ; END ;
Trouver:=ExisteSQL('SELECT TV_TAUXACH,TV_TAUXVTE,TV_CPTEACH,TV_CPTEVTE FROM TXCPTTVA WHERE TV_REGIME="'+CodeReg+'" AND (TV_CPTEACH<>"" OR TV_CPTEVTE<>"")') ;
if Trouver then BEGIN LastError:=1 ; Exit ; END ;
Result:=True ;
END ;

procedure TOT_TTREGIMETVA.OnDeleteRecord ;
Var CodeReg : String ;
    OkDet : Boolean ;
begin
CodeReg:=GetField('CC_CODE') ; if CodeReg='' then Exit ;
OkDet:=DetruitReg(CodeReg) ;
if Not OkDet then BEGIN LastError:=1 ; LastErrorMsg:='Ce régime est utilisé par le paramétrage des tiers ou taxes' ; END ;
  Inherited ;
if OkDet then ExecuteSql('Delete From TXCPTTVA Where TV_REGIME="'+CodeReg+'"') ;
end ;

Initialization
  registerclasses([TOT_TTREGIMETVA]) ;
end.

