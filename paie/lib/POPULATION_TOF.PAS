{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 31/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : POPULATION ()  Saisie des populations
                
Mots clefs ... : TOF;POPULATION
*****************************************************************}
Unit POPULATION_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     HDB,
     FE_Main,
     Fiche,
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul, 
     uTob,
     eFiche,
     MaineAgl,
{$ENDIF}
     HCtrls,
     sysutils,
     ComCtrls,
     Entpaie,
     UTOF ;

Type
  TOF_POPULATION = Class (TOF)
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

procedure TOF_POPULATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_POPULATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_POPULATION.OnUpdate ;
begin
  Inherited ;
end;


procedure TOF_POPULATION.OnLoad ;
begin
  Inherited ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 28/06/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_POPULATION.OnArgument (S : String ) ;
var
Stplus : string;

begin
  Inherited ;
 // Afficher type population PRE (Presence) et FOR (formation) uniquement si les modules sont sérialisés

  stplus := '';
  if not VH_PAIE.PGSeriaFormation then stplus := ' AND CO_CODE <> "FOR"';
  if not VH_PAIE.PGSeriaPresence then  stplus := stplus + ' AND CO_CODE <> "PRE"';
  if stplus <> '' then SetControlProperty('PPO_TYPEPOP', 'Plus', stplus);

end;


procedure TOF_POPULATION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_POPULATION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_POPULATION.OnCancel () ;
begin
  Inherited ;
end ;



Initialization
  registerclasses ( [ TOF_POPULATION ] ) ;
end.
