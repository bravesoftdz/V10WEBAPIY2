unit UTOFPG_MulPGCumExcept;

interface
uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
     HTB97,
{$IFNDEF EAGLCLIENT}
     Fe_Main,
{$ELSE}
     MaineAgl,
{$ENDIF}
     HCtrls,
     HEnt1,
     UTOF,
     Vierge;
type
  TOF_PG_MulPGCumExcept = class(TOF)
  private
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation
uses P5Def;

procedure TOF_PG_MulPGCumExcept.OnArgument(Arguments: string);
begin
inherited;
end;


initialization
  registerclasses([TOF_PG_MulPGCumExcept]);
end.

