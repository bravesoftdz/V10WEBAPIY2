unit utof_assfourchettes;

interface

uses classes, UTOF, HMsgBox, HTB97, ParamSoc, StdCtrls;

type
  TOF_ASSFOURCHETTES = Class (TOF)
     procedure OnArgument (stArgument : String ) ; override ;
     procedure OnLoad ; override ;
     procedure OnValiderClick(Sender: TObject);
  private
     procedure InitZone (St : string);
     procedure UpdateChamp (St : string);
  end;

implementation

procedure TOF_ASSFOURCHETTES.InitZone(St: string);
begin
  if GetControl(St) is TCheckBox then
  begin
    SetControlChecked(St,GetParamSoc(St)) ;
  end else SetControlText(St,getParamSoc(St));
end;

procedure TOF_ASSFOURCHETTES.UpdateChamp(St: string);
begin
  SetParamSoc(St,GetControlText(St));
end;

procedure TOF_ASSFOURCHETTES.OnArgument(stArgument: String);
begin
   TToolbarButton97(GetControl('BVALIDER')).OnClick := OnValiderClick;
end;

procedure TOF_ASSFOURCHETTES.OnLoad ;
begin
  InitZone('SO_BILDEB1');
  InitZone('SO_BILFIN1');
  InitZone('SO_BILDEB2');
  InitZone('SO_BILFIN2');
  InitZone('SO_BILDEB3');
  InitZone('SO_BILFIN3');
  InitZone('SO_BILDEB4');
  InitZone('SO_BILFIN4');
  InitZone('SO_BILDEB5');
  InitZone('SO_BILFIN5');

  InitZone('SO_PRODEB1');
  InitZone('SO_PROFIN1');
  InitZone('SO_PRODEB2');
  InitZone('SO_PROFIN2');
  InitZone('SO_PRODEB3');
  InitZone('SO_PROFIN3');
  InitZone('SO_PRODEB4');
  InitZone('SO_PROFIN4');
  InitZone('SO_PRODEB5');
  InitZone('SO_PROFIN5');

  InitZone('SO_CHADEB1');
  InitZone('SO_CHAFIN1');
  InitZone('SO_CHADEB2');
  InitZone('SO_CHAFIN2');
  InitZone('SO_CHADEB3');
  InitZone('SO_CHAFIN3');
  InitZone('SO_CHADEB4');
  InitZone('SO_CHAFIN4');
  InitZone('SO_CHADEB5');
  InitZone('SO_CHAFIN5');

  InitZone('SO_EXTDEB1');
  InitZone('SO_EXTFIN1');
  InitZone('SO_EXTDEB2');
  InitZone('SO_EXTFIN2');
end;

procedure TOF_ASSFOURCHETTES.OnValiderClick(Sender: TObject);
begin
  UpdateChamp('SO_BILDEB1');
  UpdateChamp('SO_BILFIN1');
  UpdateChamp('SO_BILDEB2');
  UpdateChamp('SO_BILFIN2');
  UpdateChamp('SO_BILDEB3');
  UpdateChamp('SO_BILFIN3');
  UpdateChamp('SO_BILDEB4');
  UpdateChamp('SO_BILFIN4');
  UpdateChamp('SO_BILDEB5');
  UpdateChamp('SO_BILFIN5');

  UpdateChamp('SO_PRODEB1');
  UpdateChamp('SO_PROFIN1');
  UpdateChamp('SO_PRODEB2');
  UpdateChamp('SO_PROFIN2');
  UpdateChamp('SO_PRODEB3');
  UpdateChamp('SO_PROFIN3');
  UpdateChamp('SO_PRODEB4');
  UpdateChamp('SO_PROFIN4');
  UpdateChamp('SO_PRODEB5');
  UpdateChamp('SO_PROFIN5');

  UpdateChamp('SO_CHADEB1');
  UpdateChamp('SO_CHAFIN1');
  UpdateChamp('SO_CHADEB2');
  UpdateChamp('SO_CHAFIN2');
  UpdateChamp('SO_CHADEB3');
  UpdateChamp('SO_CHAFIN3');
  UpdateChamp('SO_CHADEB4');
  UpdateChamp('SO_CHAFIN4');
  UpdateChamp('SO_CHADEB5');
  UpdateChamp('SO_CHAFIN5');

  UpdateChamp('SO_EXTDEB1');
  UpdateChamp('SO_EXTFIN1');
  UpdateChamp('SO_EXTDEB2');
  UpdateChamp('SO_EXTFIN2');
end;

Initialization
RegisterClasses([TOF_ASSFOURCHETTES]) ;
end.
