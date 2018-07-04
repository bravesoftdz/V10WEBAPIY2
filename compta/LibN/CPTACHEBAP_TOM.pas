{ Unité : Source TOM de la TABLE : CPTACHEBAP
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001   20/02/06    JP     Création de l'unité
--------------------------------------------------------------------------------------}
unit CPTACHEBAP_TOM;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFichList,
  {$ELSE}
  FE_Main, FichList, db, dbtables,
  {$ENDIF}
  SysUtils, UTOM, UTob;

type
  TOM_CPTACHEBAP = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnDeleteRecord           ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnLoadRecord             ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnArgument   (S : string); override;
    procedure OnClose                  ; override;
    procedure OnCancelRecord           ; override;
  private

  end;

procedure CpLanceFiche_TacheBAP(Range, Lequel, Argument : string);


implementation

uses
  HCtrls, ULibBonAPayer;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_TacheBAP(Range, Lequel, Argument : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPTACHEBAP', Range, Lequel, Argument);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetPlusCombo(THValComboBox(GetControl('CTA_DESTINATAIRE')) , 'US');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnCancelRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTACHEBAP.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

initialization
  RegisterClasses([TOM_CPTACHEBAP]);

end.
