{ Unité : Source TOT de la TABLETTE : CPRAPPROSURLIB
--------------------------------------------------------------------------------------
    Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 8.01.001.008   18/09/03  JP   Création de l'unité
--------------------------------------------------------------------------------------}
unit CPRAPPROSURLIB_TOT;

interface

uses
  Controls, Classes, UProcGen,
  {$IFDEF EAGLCLIENT}
    UTob, {dsInsert, dsEdit}
  {$ELSE}
    db,
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, HMsgBox, UTOT;

type
  TOT_CPRAPPROSURLIB = class ( TOT )
    procedure OnNewRecord           ; override;
    procedure OnUpdateRecord        ; override;
    procedure OnAfterUpdateRecord   ; override;
    procedure OnArgument(S : string); override;
  end;

implementation

{---------------------------------------------------------------------------------------}
procedure TOT_CPRAPPROSURLIB.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  (*
  if ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "' + RAPPROSURLIB +
               '" AND CC_CODE = "' + DS.FindField('CC_CODE').AsString + '"') then
  if (Trim(DS.FindField('CO_LIBRE').AsString) = '') and (Trim(GetField('CO_LIBRE')) = '') then begin

    LastError := 1;
  end;
  *)
end;

{---------------------------------------------------------------------------------------}
procedure TOT_CPRAPPROSURLIB.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  AvertirTable('CPRAPPROSURLIB');
  MajInfoTablette('CPRAPPROSURLIB');
end;

{---------------------------------------------------------------------------------------}
procedure TOT_CPRAPPROSURLIB.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  OkComplement := False;

end;

{---------------------------------------------------------------------------------------}
procedure TOT_CPRAPPROSURLIB.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
//
end;


initialization
  RegisterClasses([TOT_CPRAPPROSURLIB]);

end.
