{ Unité : Source TOF de la FICHE : TRJOURFERIE
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            18/09/03    JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit TRJOURFERIE_TOF ;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
    FE_Main,
  {$ELSE}
    MaineAGL,
  {$ENDIF}
  Vierge, SysUtils, HCtrls, UTOF;

type
  TOF_TRJOURFERIE = class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : string ) ; override ;
  private
    procedure RadioClick(Sender : TObject);
  end ;

function TR_DetailJourFerier(Arg : string) : string;

const
  CONST_ANNULER = 'ANNULER';

Implementation

{---------------------------------------------------------------------------------------}
function TR_DetailJourFerier(Arg : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AGLLanceFiche('TR', 'TRJOURFERIE', '', '', Arg);
  if Result = '' then Result := CONST_ANNULER;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRJOURFERIE.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  j, m, a : Word;
  dt      : string;
begin
  inherited;
  Ecran.HelpContext := 500011;
  if TRadioButton(GetControl('RBFIXE')).Checked then begin
    dt := GetControlText('DATEFIXE');
    DecodeDate(StrToDate(dt), a, m, j);
    dt := DateToStr(EncodeDate(1970, m, j));
    TFVierge(Ecran).Retour := dt;
  end
  else if THSpinEdit(GetControl('ECARTBASE')).Value = 0 then
    TFVierge(Ecran).Retour := 'P'
  else
    TFVierge(Ecran).Retour := 'P+' + IntToStr(THSpinEdit(GetControl('ECARTBASE')).Value);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRJOURFERIE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  dt : string;
  n  : Integer;
begin
  inherited;
  TRadioButton(GetControl('RBMOBILE')).OnClick := RadioClick;
  TRadioButton(GetControl('RBFIXE'  )).OnClick := RadioClick;

  dt := ReadTokenSt(S);
  {Si on a déjà renseigner la date}
  if Trim(dt) <> '' then begin
    {Jour mobile par rapport à Pâques}
    if dt[1] = 'P' then begin
      if Trim(dt) = 'P' then n:= 0 {Il s'agit de Pâques}
                        else n := StrToInt(Copy(dt, 3, 9));

      THSpinEdit(GetControl('ECARTBASE')).Value := n;
      TRadioButton(GetControl('RBMOBILE')).Checked := True;
      RadioClick(GetControl('RBMOBILE'));
    end
    else
      SetControlText('DATEFIXE', dt);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRJOURFERIE.RadioClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
begin
  Ok := ((UpperCase(TComponent(Sender).Name) = 'RBMOBILE') and TRadioButton(Sender).Checked) or
        ((UpperCase(TComponent(Sender).Name) = 'RBFIXE') and TRadioButton(Sender).Checked);
  SetControlEnabled('ECARTBASE',    Ok);
  SetControlEnabled('LBECARTBASE',  Ok);
  SetControlEnabled('DATEFIXE', not Ok);
end;

Initialization
  registerclasses ( [ TOF_TRJOURFERIE ] ) ;
end.
