{ Unité : Source TOT de la TABLETTE : TRPORTEFEUILLE
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
6.30.001.004     15/03/05   JP    Création de l'unité en Correction de la FQ 10227

--------------------------------------------------------------------------------------}
unit TRPORTEFEUILLE_TOT;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Classes, UTOT, HEnt1, HPanel;

type
  TOT_TRPORTEFEUILLE = class ( TOT )
    procedure OnDeleteRecord        ; override;
    procedure OnAfterUpdateRecord   ; override;
    procedure OnArgument(S : string); override;
  end;

procedure LanceTot_Portefeuille(Action : TActionFiche; Inside : THPanel);


implementation

uses
  {$IFDEF EAGL}
  eTablette,
  {$ELSE}
  Tablette,
  {$ENDIF EAGL}
  HCtrls;

{---------------------------------------------------------------------------------------}
procedure LanceTot_Portefeuille(Action : TActionFiche; Inside : THPanel);
{---------------------------------------------------------------------------------------}
begin
  ParamTable('TRPORTEFEUILLE', Action, 150, Inside, 3, 'Portefeuilles de valeurs mobilières');
end;

{---------------------------------------------------------------------------------------}
procedure TOT_TRPORTEFEUILLE.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {15/03/05 : FQ 10227 : on s'assure que le portefeuille n'est pas utilisé}
  if ExisteSQL('SELECT TOF_PORTEFEUILLE FROM TROPCVMREF WHERE TOF_PORTEFEUILLE = "' + VarToStr(GetField('CC_CODE')) + '"')then begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer ce portefeuille.'#13 +
                                    'Il est utilisé dans la définition d''un OPCVM.');
    LastError    := 2;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOT_TRPORTEFEUILLE.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  AvertirTable('TRPORTEFEUILLE');
  MajInfoTablette('TRPORTEFEUILLE');
end;

{---------------------------------------------------------------------------------------}
procedure TOT_TRPORTEFEUILLE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;


initialization
  RegisterClasses([TOT_TRPORTEFEUILLE]);

end.
