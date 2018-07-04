{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AFPUBLICATION (AFPUBLICATION)
Mots clefs ... : TOM;AFPUBLICATION
*****************************************************************}
Unit uTomAfPublication;

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
     UTOM,
     UTob ;

Type
  TOM_AFPUBLICATION = Class (TOM)
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
  const
	TexteMsg: array[1..1] of string 	= (
          {1}        'Suppression impossible certaines valeurs d''indices dépendent de cette publication'
          );

procedure AglLanceFicheAFPUBLICATION(cle,Action : string ) ;

Implementation

procedure TOM_AFPUBLICATION.OnNewRecord ;
begin
  Inherited ;
  SetControlText('AFP_PUBMAJ', 'SAI');
end ;                                  

procedure TOM_AFPUBLICATION.OnDeleteRecord ;
var st : string ;

begin
  Inherited ;
  st:='SELECT * FROM afvalindice WHERE afv_pubcode = "'+getfield('AFP_PUBCODE')+'"' ;
  If ExisteSQL(st) then
  begin
     LastError    := 1;
     LastErrorMsg := TexteMsg[LastError];
     exit;
  end;
end;


procedure TOM_AFPUBLICATION.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_AFPUBLICATION.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_AFPUBLICATION.OnLoadRecord ;
begin
  Inherited ;
  SetFocusControl('AFP_PUBCODE') ;
end ;

procedure TOM_AFPUBLICATION.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_AFPUBLICATION.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_AFPUBLICATION.OnClose ;
begin
  Inherited ;
  avertirTable('AFTPUBLICATION') ;
  avertirTable('AFPUBLICATION_LIB') ;
  avertirTable('AFPUBLICATION') ;
end ;

procedure TOM_AFPUBLICATION.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure AglLanceFicheAFPUBLICATION(cle,Action : string ) ;
begin
  AglLanceFiche('AFF','AFPUBLICATION','',cle,Action);
end ;

Initialization
  registerclasses ( [ TOM_AFPUBLICATION ] ) ; 
end.
