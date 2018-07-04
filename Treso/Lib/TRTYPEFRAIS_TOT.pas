{ Unité : Source TOT de la TABLETTE : TRTYPEFRAIS
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
6.20.001.001     29/11/04   JP    Création de l'unité
6.30.001.004     15/03/05   JP    Création de la fonction d'appel et application de la FQ 10227
--------------------------------------------------------------------------------------}
unit TRTYPEFRAIS_TOT;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Classes, UTOT, HEnt1, HPanel;

type
  TOT_TRTYPEFRAIS = class ( TOT )
    procedure OnNewRecord           ; override;
    procedure OnDeleteRecord        ; override;
    procedure OnAfterUpdateRecord   ; override;
    procedure OnArgument(S : string); override;
  end;

procedure LanceTot_TypeFrais(Action : TActionFiche; Inside : THPanel);

implementation

uses
  {$IFDEF EAGL}
  eTablette,
  {$ELSE}
  Tablette,
  {$ENDIF EAGL}
  Constantes, HCtrls;

{---------------------------------------------------------------------------------------}
procedure LanceTot_TypeFrais(Action : TActionFiche; Inside : THPanel);
{---------------------------------------------------------------------------------------}
begin
  ParamTable('TRTYPEFRAIS', Action, 150, Inside, 3, 'Services bancaires');
end;

{---------------------------------------------------------------------------------------}
procedure TOT_TRTYPEFRAIS.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  if (GetField('CC_CODE') = CODEFRAISOPCVM) or (GetField('CC_CODE') = CODEFCOMOPCVM) then begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer ce service.');
    LastError    := 2;
  end
  {15/03/05 : FQ 10227 : on s'assure que le service n'est pas utilisé}
  else if ExisteSQL('SELECT TFR_TYPEFRAIS FROM FRAIS WHERE TFR_TYPEFRAIS = "' + VarToStr(GetField('CC_CODE')) + '"')then begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer ce service.'#13 +
                                    'Il est utilisé dans la définition d''un frais.');
    LastError    := 2;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOT_TRTYPEFRAIS.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  AvertirTable('TRTYPEFRAIS');
  MajInfoTablette('TRTYPEFRAIS');
end;

{---------------------------------------------------------------------------------------}
procedure TOT_TRTYPEFRAIS.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOT_TRTYPEFRAIS.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
//
end;


initialization
  RegisterClasses([TOT_TRTYPEFRAIS]);

end.
