{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : SCORINGDEF (SCORINGDEF)
Mots clefs ... : TOM;SCORINGDEF
*****************************************************************}
Unit UTomScoringDef ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,
{$else}
     eFiche,
     eFichList,
     UTob,
{$ENDIF}
     UTOM;

Type
  TOM_SCORINGDEF = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation

procedure TOM_SCORINGDEF.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_SCORINGDEF.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_SCORINGDEF ] ) ; 
end.

