{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/12/2013
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BTSOUSFAMILLETARIF (BTSOUSFAMILLETARIF)
Mots clefs ... : TOM;BTSOUSFAMILLETARIF
*****************************************************************}
Unit TOMBTSOUSFAMILLETARIF ;

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
  TOM_BTSOUSFAMILLETARIF = Class (TOM)
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

procedure TOM_BTSOUSFAMILLETARIF.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_BTSOUSFAMILLETARIF.OnAfterRollBack;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOM_BTSOUSFAMILLETARIF ] ) ; 
end.

