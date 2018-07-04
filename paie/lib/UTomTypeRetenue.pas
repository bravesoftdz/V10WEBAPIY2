{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TYPERETENUE (TYPERETENUE)
Mots clefs ... : TOM;TYPERETENUE
*****************************************************************}
Unit UTomTypeRetenue ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,
{$ELSE}
   eFiche,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,PGOutils ;

Type
  TOM_TYPERETENUE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    private
    CEG,LectureSeule,STD,DOS : Boolean;
    end ;

Implementation

procedure TOM_TYPERETENUE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_TYPERETENUE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_TYPERETENUE.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TYPERETENUE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TYPERETENUE.OnLoadRecord ;
begin
  Inherited ;
  AccesPredefini('TOUS',CEG,STD,DOS);
  LectureSeule := False;
  if (Getfield('PTR_PREDEFINI') = 'CEG')  then
  begin
  LectureSeule := (CEG=False);
  PaieLectureSeule(TFFiche(Ecran),(CEG=False));
  SetControlEnabled('BDelete',CEG);
  end;
end ;

procedure TOM_TYPERETENUE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_TYPERETENUE.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_TYPERETENUE.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_TYPERETENUE ] ) ; 
end.

