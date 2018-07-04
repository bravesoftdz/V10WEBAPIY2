{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/05/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMUL ()
Mots clefs ... : TOF;BTMUL
*****************************************************************}
Unit BTMUL_TOF;

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
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox, 
     UTOF ; 

Type
  TOF_BTMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_BTMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMUL.OnArgument (S : String ) ;
var BParamListe : TToolbarButton97;
begin

		BparamListe := TToolbarButton97 (GetCOntrol('BparamListe'));

//uniquement en line
{*
    if V_PGI.Sav then
    begin
    	if BparamListe <> nil then BParamListe.Visible := True
    end
    else
    begin
    	if BparamListe <> nil then BParamListe.Visible := False;
    end;
*}

  Inherited ;

end ;

procedure TOF_BTMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTMUL ] ) ; 
end.
