{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 29/09/2005
Modifié le ... :   /  /
Description .. :
Suite ........ :
Mots clefs ... :
*****************************************************************}
unit UTofPG_MulPGExcept;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  Fe_Main,
{$ELSE}
  MaineAgl,
{$ENDIF}
  HCtrls, HEnt1, UTOF, Vierge;
type
  TOF_PG_MulPGExcept = class(TOF)
  private
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation
uses P5Def;

procedure TOF_PG_MulPGExcept.OnArgument(Arguments: string);
begin
  inherited;
  if (V_PGI.ModePcl <> '1') OR (PGRendModeFonc () = 'MONO') then
  begin
    SetControlText('PEN_PREDEFINI', 'DOS');
    SetControlEnabled('PEN_PREDEFINI', FALSE);
  end;
end;

initialization
  registerclasses([TOF_PG_MulPGExcept]);
end.

