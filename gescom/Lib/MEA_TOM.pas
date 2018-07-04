{***********UNITE*************************************************
Auteur  ...... :  LS
Créé le ...... : 03/05/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : MEA (MEA)
Mots clefs ... : TOM;MEA
*****************************************************************}
Unit MEA_TOM ;

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
  TOM_MEA = Class (TOM)
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

procedure TOM_MEA.OnNewRecord ;
begin
  Inherited ;
//uniquement en line
{*
  SetField('GME_QUALIFMESURE','PIE');
  SetField('GME_QUOTITE',1);
*}
end ;

procedure TOM_MEA.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_MEA.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_MEA.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_MEA.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_MEA.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_MEA.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_MEA.OnArgument ( S: String ) ;
begin
  Inherited ;
  //
//uniquement en line
{*
  TGroupBox(GetControl('TCORRESPONDANCE')).visible := false;
*}
end ;

procedure TOM_MEA.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_MEA.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_MEA ] ) ; 
end.

