{ Unité : Source TOM de la TABLE : REGLEACCRO
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91            10/10/03   JP   Création de l'unité
 8.01.001.006    06/03/07   JP   Ajout du champ prédéfini
--------------------------------------------------------------------------------------}
unit REGLEACCRO_TOM;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  FE_Main,
  {$else}
  MaineAGL,
  {$ENDIF}
   UTOM;

type
  TOM_REGLEACCRO = class (TOM)
    procedure OnArgument(S : string); override;
    procedure OnNewRecord           ; override; {06/03/07}
  private

  end;

procedure TRLanceFiche_ParamRegleAccro(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  HEnt1;

{Je pense qu'il va falloir envisager des règles d'incompatibilité entre les cases à cocher
{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_ParamRegleAccro(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_REGLEACCRO.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000141;

  SetControlVisible('TRG_BNUMCOMPTE', False);
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_REGLEACCRO.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited; {06/03/07}
  if CtxPcl in V_PGI.PGIContexte then SetField('TRG_PREDEFINI', 'CEG')
                                 else SetField('TRG_PREDEFINI', 'STD');
end;

initialization
  RegisterClasses([TOM_REGLEACCRO]);
   
end.
