unit uTOFZoomELignes;

interface
uses UTOF;

type
    TOF_ZoomELignes = class(TOF)

    published
      procedure OnArgument(stArgument : String); override;

    end;

implementation
uses Classes, HCtrls;

procedure TOF_ZoomELignes.OnArgument(stArgument : String);
begin
inherited;
SetControlText('EL_NATUREPIECEG', ReadTokenSt(stArgument));
SetControlText('EL_SOUCHE', ReadTokenSt(stArgument));
SetControlText('EL_NUMERO', ReadTokenSt(stArgument));
SetControlText('EL_INDICEG', ReadTokenSt(stArgument));
end;


initialization
RegisterClasses([TOF_ZoomELignes]);

end.
