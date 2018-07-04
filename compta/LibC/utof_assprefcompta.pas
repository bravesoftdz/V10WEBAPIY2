unit utof_assprefcompta;

interface

uses classes, UTOF, HMsgBox, HTB97, ParamSoc, StdCtrls;

type
  TOF_ASSPREFCOMPTA = Class (TOF)
     procedure OnArgument (stArgument : String ) ; override ;
     procedure OnLoad ; override ;
     procedure OnValiderClick(Sender: TObject);
  private
     procedure InitZone (St : string);
     procedure UpdateChamp (St : string);
  end;

implementation

procedure TOF_ASSPREFCOMPTA.InitZone(St: string);
begin
  if GetControl(St) is TCheckBox then
  begin
    SetControlChecked(St,GetParamSoc(St)) ;
  end else SetControlText(St,getParamSoc(St));
end;

procedure TOF_ASSPREFCOMPTA.UpdateChamp(St: string);
begin
  SetParamSoc(St,GetControlText(St));
end;

procedure TOF_ASSPREFCOMPTA.OnArgument(stArgument: String);
begin
   TToolbarButton97(GetControl('BVALIDER')).OnClick := OnValiderClick;
end;

procedure TOF_ASSPREFCOMPTA.OnLoad ;
begin
  InitZone('SO_MONTANTNEGATIF');
  InitZone('SO_ALERTEPARTIEL');
  InitZone('SO_CPDATEOBLI');
  InitZone('SO_NBJECRAVANT');
  InitZone('SO_NBJECRAPRES');
  InitZone('SO_JALLOOKUP');
  InitZone('SO_ETABLOOKUP');
  InitZone('SO_ALERTEREGUL');
  InitZone('SO_BOUCLERSAISIECREAT');
  InitZone('SO_AUTOEXTRASOLDE');
  InitZone('SO_ATTRIBRIBAUTO');
  InitZone('SO_CPCONTROLEDOUBLON');
  InitZone('SO_CPCHAMPDOUBLON');
  InitZone('SO_CPSTATUSBARRE');
  InitZone('SO_CPREFLETTRAGE');
end;

procedure TOF_ASSPREFCOMPTA.OnValiderClick(Sender: TObject);
begin
  UpdateChamp('SO_MONTANTNEGATIF');
  UpdateChamp('SO_ALERTEPARTIEL');
  UpdateChamp('SO_CPDATEOBLI');
  UpdateChamp('SO_NBJECRAVANT');
  UpdateChamp('SO_NBJECRAPRES');
  UpdateChamp('SO_JALLOOKUP');
  UpdateChamp('SO_ETABLOOKUP');
  UpdateChamp('SO_ALERTEREGUL');
  UpdateChamp('SO_BOUCLERSAISIECREAT');
  UpdateChamp('SO_AUTOEXTRASOLDE');
  UpdateChamp('SO_ATTRIBRIBAUTO');
  UpdateChamp('SO_CPCONTROLEDOUBLON');
  UpdateChamp('SO_CPCHAMPDOUBLON');
  UpdateChamp('SO_CPSTATUSBARRE');
  UpdateChamp('SO_CPREFLETTRAGE');
end;

Initialization
RegisterClasses([TOF_ASSPREFCOMPTA]) ;

end.



