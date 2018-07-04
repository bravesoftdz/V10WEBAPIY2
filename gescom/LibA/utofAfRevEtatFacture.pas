{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVETATFACTURE ()
Mots clefs ... : TOF;AFREVETATFACTURE
*****************************************************************}
Unit utofAfRevEtatFacture ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$Else}
     MainEagl,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,UTofAfBaseCodeAffaire,UtilRevision ;

Type
  TOF_AFREVETATFACTURE = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  end ;

procedure LanceEtatFactureARegulariser ;

Implementation

procedure TOF_AFREVETATFACTURE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1')); Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3')); Aff4:=THEdit(GetControl('AFF_AVENANT'));
end;


procedure TOF_AFREVETATFACTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVETATFACTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVETATFACTURE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFREVETATFACTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFREVETATFACTURE.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_AFREVETATFACTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFREVETATFACTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVETATFACTURE.OnCancel () ;
begin
  Inherited ;
end ;


procedure LanceEtatFactureARegulariser ;
begin
  FlagueLesLignesRegularisables(True) ;
  AGLLanceFiche('AFF','AFREVETATFACTURE','','','');
end ;

Initialization
  registerclasses ( [ TOF_AFREVETATFACTURE ] ) ;
end.
