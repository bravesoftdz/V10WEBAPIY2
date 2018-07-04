{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMETHODVENTILS1 ()
Mots clefs ... : TOF;BTMETHODVENTILS1
*****************************************************************}
Unit BTMETHODVENTILS1_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     Paramsoc,
     entgc,
     UtilLine,
     BTPARAMSOCGEN_TOF,
     UTOF ;

Type
  TOF_BTMETHODVENTILS1 = Class (TOF_BTPARAMSOCGEN)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	function VerifInfos : boolean;
  end ;

Implementation

procedure TOF_BTMETHODVENTILS1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMETHODVENTILS1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMETHODVENTILS1.OnUpdate ;
begin
  Inherited ;
  if not verifInfos then
  begin
  	TForm(Ecran).ModalResult:=0;
    exit;
  end;
	StockeInfos(Ecran, LaTob);
  // on stocke directement pour que la suite soit ok
  SetParamSoc('SO_GCVENTCPTAART',GetControltext('SO_GCVENTCPTAART'));
  SetParamSoc('SO_GCVENTCPTATIERS',GetControltext('SO_GCVENTCPTATIERS'));
  SetParamSoc('SO_GCVENTCPTAAFF',GetControltext('SO_GCVENTCPTAAFF'));
{ Devenu inutile en V8
  // + stockage dans les VH_GC...grrr
  VH_GC.GCVentCptaArt:=GetParamSocSecur('SO_GCVENTCPTAART', False) ;
  VH_GC.GCVentCptaTiers:=GetParamSocSecur('SO_GCVENTCPTATIERS', False) ;
  VH_GC.GCVentCptaAff:=GetParamSocSecur('SO_GCVENTCPTAAFF', False) ;
}
end ;

procedure TOF_BTMETHODVENTILS1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMETHODVENTILS1.OnArgument (S : String ) ;
begin
  Inherited ;
  chargeEcran(ecran, LaTob);
end ;

procedure TOF_BTMETHODVENTILS1.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMETHODVENTILS1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMETHODVENTILS1.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTMETHODVENTILS1.VerifInfos: boolean;
begin
	result := true;
end;

Initialization
  registerclasses ( [ TOF_BTMETHODVENTILS1 ] ) ; 
end.

