{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECALCPIECEUNI ()
Mots clefs ... : TOF;BTRECALCPIECEUNI
*****************************************************************}
Unit BTRECALCPIECEUNI ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
		 MaineAGL,
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     AglInit,
     HMsgBox, 
     UTOF,
     UTOB ; 

Type
  TOF_BTRECALCPIECEUNI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function RecalcPieceUni (var WithPuv : boolean) : boolean;

Implementation

function RecalcPieceUni (var WithPuv : boolean) : boolean;
var TOBResult : TOB;
begin
	result := true;
	TOBresult := TOB.create ('LES PARAMS',nil,-1);
  TOBresult.AddChampSupValeur ('CALCULPV','-');
	TheTOB := TOBresult;
  AglLanceFiche ('BTP','BTRECALCPIECEUNI','','','');
  WithPuv := (TOBresult.GetValue('CALCULPV')='X');
  if TheTOB = nil then result := false;
  TheTOB := Nil;
end;

procedure TOF_BTRECALCPIECEUNI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECEUNI.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECEUNI.OnUpdate ;
begin
  Inherited ;
  if TCheckBox(GetControl('CBCALCPV')).Checked then LaTOB.putValue('CALCULPV','X');
  TheTOB := LaTob;
end ;

procedure TOF_BTRECALCPIECEUNI.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECEUNI.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECEUNI.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECEUNI.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECEUNI.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTRECALCPIECEUNI ] ) ; 
end.

