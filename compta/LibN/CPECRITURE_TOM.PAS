{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/11/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ECRITURE (ECRITURE)
Mots clefs ... : TOM;ECRITURE
*****************************************************************}
Unit CPECRITURE_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
{$ELSE}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uLibEcriture,
     UTOM,
     UTob ;

Type
  TOM_ECRITURE = Class (TOM)
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
  private

    end ;

Implementation

procedure TOM_ECRITURE.OnNewRecord ;
begin
  Inherited ;

  if assigned( FTOB ) then
    CPutDefautEcr( FTOB ) ;

end ;

procedure TOM_ECRITURE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnLoadRecord ;
var lCode : String ;
begin
  Inherited ;

  if assigned( FTOB ) then
    lCode := FTOB.GetValue('E_GENERAL') ;

end ;

procedure TOM_ECRITURE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnArgument ( S: String ) ;
var lNew : Boolean ;
begin
  Inherited ;

  if assigned( FTOB ) then
    lNew := FTOB.FieldExists('E_GENERAL') ;

end ;

procedure TOM_ECRITURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_ECRITURE ] ) ; 
end.
