{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 10/06/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTESTFORMULE ()
Mots clefs ... : TOF;BTESTFORMULE
*****************************************************************}
Unit BTESTFORMULE_TOF ;

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
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UCalculsMetre,
     HTB97,
     UTOF ;

Type
  TOF_BTESTFORMULE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	BCALC : TToolbarButton97;
    procedure BCalcClick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTESTFORMULE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTESTFORMULE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTESTFORMULE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTESTFORMULE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTESTFORMULE.OnArgument (S : String ) ;
begin
  Inherited ;
  BCALC := TToolbarButton97 (getControl('BCALC'));
  BCALC.OnClick := BCalcClick;
end ;

procedure TOF_BTESTFORMULE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTESTFORMULE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTESTFORMULE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTESTFORMULE.BCalcClick(Sender: TObject);
var TheChaine : string;
begin
	SetControlProperty('RESULT','Value',0);
  TheChaine := getControltext('FORMULE');
  if TheChaine = '' then
  BEGIN
    Exit;
  END;
  SetControlProperty('RESULT','Value',CalculeFormuleLigne(TheChaine));
end;

Initialization
  registerclasses ( [ TOF_BTESTFORMULE ] ) ; 
end.
