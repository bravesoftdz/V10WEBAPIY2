{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARSOCCPTAGS1 ()
Mots clefs ... : TOF;BTPARSOCCPTAGS1
*****************************************************************}
Unit BTPARSOCCPTAGS1_TOF ;

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
     BTPARAMSOCGEN_TOF,
     utilLine,
     UTOF ; 

Type
  TOF_BTPARSOCCPTAGS1 = Class (TOF_BTPARAMSOCGEN)
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

procedure TOF_BTPARSOCCPTAGS1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARSOCCPTAGS1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARSOCCPTAGS1.OnUpdate ;
begin
  Inherited ;
  if not verifInfos then
  begin
  	TForm(Ecran).ModalResult:=0;
    exit;
  end;
	StockeInfos(Ecran, Latob);
end ;

procedure TOF_BTPARSOCCPTAGS1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARSOCCPTAGS1.OnArgument (S : String ) ;
begin
  Inherited ;
  chargeEcran(Ecran, LaTob);
end ;

procedure TOF_BTPARSOCCPTAGS1.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPARSOCCPTAGS1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARSOCCPTAGS1.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTPARSOCCPTAGS1.VerifInfos: boolean;
begin
  result := true;
end;

Initialization
  registerclasses ( [ TOF_BTPARSOCCPTAGS1 ] ) ; 
end.

