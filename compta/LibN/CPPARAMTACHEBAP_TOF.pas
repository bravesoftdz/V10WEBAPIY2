{ Unité : Source TOF de la fiche CPPARAMTACHEBAP
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001   08/02/06    JP     Création de l'unité
--------------------------------------------------------------------------------------}
unit CPPARAMTACHEBAP_TOF;

interface

uses
  Controls, Classes, Vierge, 
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, DBTables,
  {$ENDIF}
  SysUtils, UTob, UTOF;

type
  TOF_CPPARAMTACHEBAP = class(TOF)

    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
  private

  public

  end;

procedure CpLanceFiche_ParamTache;

implementation

uses
  HCtrls, HEnt1;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_ParamTache;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPPARAMTACHEBAP', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPARAMTACHEBAP.OnArgument(S: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509125;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPARAMTACHEBAP.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPARAMTACHEBAP.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  
end;


initialization
  RegisterClasses([TOF_CPPARAMTACHEBAP]);

end.
