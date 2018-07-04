{ Unité : Source TOT de la TABLETTE : TTJOURFERIER
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            18/09/03    JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit TTJOURFERIER_TOT ;

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
  TOT_TTJOURFERIER = class ( TOT )
    procedure OnNewRecord           ; override;
    procedure OnUpdateRecord        ; override;
    procedure OnAfterUpdateRecord   ; override;
    procedure OnArgument(S : string); override;
    procedure OnComplement(Prefixe, Tipe, Code : string) ; override;
  end;

implementation

uses TRJOURFERIE_TOF;

{---------------------------------------------------------------------------------------}
procedure TOT_TTJOURFERIER.OnUpdateRecord () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  if (Trim(DS.FindField('CO_LIBRE').AsString) = '') and (Trim(GetField('CO_LIBRE')) = '') then begin
    HShowMessage('0;Saisie de jours fériés; Veuillez saisir une date !;W;O;O;O;', '', '');
    OnComplement('', '', '');
    LastError := 1;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOT_TTJOURFERIER.OnAfterUpdateRecord () ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  AvertirTable('TTJOURFERIER');
  MajInfoTablette('TTJOURFERIER');
end ;

{---------------------------------------------------------------------------------------}
procedure TOT_TTJOURFERIER.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50011;
  OkComplement := True;
  bComplement.Hint := TraduireMemoire('Saisie de la date');
  bComplement.GlobalIndexImage := 'Z0008_S16G1';
end ;

{---------------------------------------------------------------------------------------}
procedure TOT_TTJOURFERIER.OnComplement(Prefixe, Tipe, Code: string);
{---------------------------------------------------------------------------------------}
var
  ret : string;
begin
  inherited;
  {Saisie de la date}
  Ret := TR_DetailJourFerier(GetField('CO_LIBRE'));
  if Ret <> CONST_ANNULER then begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;

    SetField('CO_LIBRE', Ret)
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOT_TTJOURFERIER.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
//
end;


Initialization
  registerclasses ( [ TOT_TTJOURFERIER ] ) ;
end.
