{***********UNITE*************************************************
Auteur  ...... : Guillon stéphane - SG6
Créé le ...... : 16/11/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : EEXBQ (EEXBQ)
Mots clefs ... : TOM;EEXBQ
*****************************************************************}
Unit EEXBQ_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
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
  TOM_EEXBQ = Class (TOM)
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
    end ;

Implementation

procedure TOM_EEXBQ.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_EEXBQ.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_EEXBQ ] ) ; 
end.

