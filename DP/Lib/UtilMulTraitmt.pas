unit UtilMulTraitmt;

interface

uses
    HEnt1,
{$IFDEF EAGLCLIENT}
    eMul,
{$ELSE}
    Mul,
{$ENDIF}
    HCtrls;

Procedure FinTraitmtMul (Mul : TFMul);

/////////////// IMPLEMENTATION //////////////
implementation

Procedure FinTraitmtMul (Mul : TFMul);
// cf FinTraiteMul dans UtilMulTrt de gescom\liba...
begin
  if Mul.FListe.AllSelected then Mul.FListe.AllSelected:=False
                            else Mul.FListe.ClearSelected;
  Mul.bSelectAll.Down := False ;
end;

end.
 