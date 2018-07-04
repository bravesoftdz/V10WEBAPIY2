{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/02/2013
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DEPOTS (DEPOTS)
Mots clefs ... : TOM;DEPOTS
*****************************************************************}
Unit DEPOTS_TOM ;

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

  TOM_DEPOTS = Class (TOM)
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
  private
    procedure SetLastError(Num: integer; ou: string);
    end ;

const
  TexteMessage: array[1..1] of string = (
    {1} 'Merci de renseigner une section analytique existante pour ce dépot'
    );

Implementation
uses   ParamSoc;

procedure TOM_DEPOTS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnUpdateRecord ;
begin
  Inherited ;
  if GetParamSocSecur('SO_OPTANALSTOCK',false) then
  begin
    if not ExisteSQL ('SELECT S_SECTION FROM SECTION WHERE S_SECTION="'+THedit(GetControl('GDE_SECTION')).Text+'" '+
                      'AND S_AXE="'+GetparamSocSecur('SO_BTAXEANALSTOCK','')+'"') then
    begin
       SetLastError(1,'GDE_SECTION');
       exit;
    end;
  end;
end ;

procedure TOM_DEPOTS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnArgument ( S: String ) ;
begin
  Inherited ;
  if (GetParamSocSecur('SO_OPTANALSTOCK',false)) and (GetparamSocSecur('SO_BTAXEANALSTOCK','')<> '')then
  begin
    SetControlVisible('GB_ANALYTIQUE',true);
  end;

end ;

procedure TOM_DEPOTS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterRollBack;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
end;

Initialization
  registerclasses ( [ TOM_DEPOTS ] ) ;
end.
