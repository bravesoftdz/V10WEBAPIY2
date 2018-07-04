{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/09/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YREPERTPERSO (YREPERTPERSO)
Mots clefs ... : TOM;YREPERTPERSO
*****************************************************************}
Unit UTomYRepertPerso ;

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
  TOM_YREPERTPERSO = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  public
  private
  end ;

////////// IMPLEMENTATION //////////
Implementation

uses
    utilPgi
{$IFDEF DP}
    ,entDP
{$ENDIF}
    ;

procedure TOM_YREPERTPERSO.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_YREPERTPERSO.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YREPERTPERSO.OnUpdateRecord ;
begin
  Inherited ;
  // if DS.State in [dsInsert] then => ne pas utiliser,
  // car tom utilisée sans fiche dans AssistImportCarnet via CreateTom
  // et dans ce cas le param Ecran est à Nil, et DS aussi
  if GetField('YRP_GUIDREP')='' then
    begin
    SetField('YRP_GUIDREP', AglGetGuid());
    SetField('YRP_USER', V_PGI.User);
    end;

  // $$$ JP 13/06/05
  SetField ('YRP_CLETELEPHONE',  CleTelephone (GetField ('YRP_TEL1')));
  SetField ('YRP_CLETELEPHONE2', CleTelephone (GetField ('YRP_TEL2')));
end ;

procedure TOM_YREPERTPERSO.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YREPERTPERSO.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_YREPERTPERSO.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YREPERTPERSO.OnArgument ( S: String ) ;
begin
  inherited ;

{$IFDEF DP}
  SetControlVisible ('BTEL1', VH_DP.ctiAlerte <> nil);
  SetControlVisible ('BTEL2', VH_DP.ctiAlerte <> nil);
{$ELSE}
  SetControlVisible ('BTEL1', FALSE);
  SetControlVisible ('BTEL2', FALSE);
{$ENDIF}
end ;

procedure TOM_YREPERTPERSO.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_YREPERTPERSO.OnCancelRecord ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOM_YREPERTPERSO ] ) ; 
end.

