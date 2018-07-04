{***********UNITE*************************************************
Auteur  ...... : Santucci Lionel
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DATECLOTURE ()
Mots clefs ... : TOF;DATECLOTURE
*****************************************************************}
Unit UTOFBTClotureDev ;

Interface

Uses StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     fe_main,
{$ELSE}
     mainEAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     UTOB,
     HMsgBox,
     AglInit,
     UTOF ;

Type
  TOF_DATECLOTURE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function SaisieDateCloture (var DateCloture : TDateTime) : boolean;

Implementation

function SaisieDateCloture (var DateCloture : TDateTime) : boolean;
var TOBParam : TOB;
begin
result := false;
TOBPARAM := TOB.Create ('DATE CLOTURE',nil,-1);
TRY
TOBPAram.addChampsupValeur ('DATECLOTURE',DateCloture,false);
TheTob := TOBParam;
AGLLanceFiche ('BTP','BTCLOTUREDEV','','','');
if TheTob <> nil then
  BEGIN
  DateCloture := TheTOB.GetValue('DATECLOTURE');
  result := true;
  END;
FINALLY
TheTob := nil;
TOBPARAM.free;
end;
end;

procedure TOF_DATECLOTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DATECLOTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DATECLOTURE.OnUpdate ;
begin
  Inherited ;
  TheTob := LaTOB;
end ;

procedure TOF_DATECLOTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DATECLOTURE.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_DATECLOTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_DATECLOTURE.OnCancel;
begin
  inherited;

end;

Initialization
  registerclasses ( [ TOF_DATECLOTURE ] ) ; 
end.

