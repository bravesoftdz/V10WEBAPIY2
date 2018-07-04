{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/08/2017
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CONDITIONNEMENT (CONDITIONNEMENT)
Mots clefs ... : TOM;CONDITIONNEMENT
*****************************************************************}
Unit CONDITIONNEMENT_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob ;

Type
  TOM_CONDITIONNEMENT = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
    end ;

Implementation

procedure TOM_CONDITIONNEMENT.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_CONDITIONNEMENT.OnAfterRollBack;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_CONDITIONNEMENT ] ) ; 
end.
