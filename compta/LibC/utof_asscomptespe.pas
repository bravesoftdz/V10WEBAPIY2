unit utof_asscomptespe;

interface

uses classes, UTOF, HMsgBox, HTB97, ParamSoc, StdCtrls;

type
  TOF_ASSCOMPTESPE = Class (TOF)
     procedure OnArgument (stArgument : String ) ; override ;
     procedure OnLoad ; override ;
     procedure OnValiderClick(Sender: TObject);
  private
     procedure InitZone (St : string);
     procedure UpdateChamp (St : string);
  end;

implementation

procedure TOF_ASSCOMPTESPE.InitZone(St: string);
begin
  if GetControl(St) is TCheckBox then
  begin
    SetControlChecked(St,GetParamSoc(St)) ;
  end else SetControlText(St,getParamSoc(St));
end;

procedure TOF_ASSCOMPTESPE.UpdateChamp(St: string);
begin
  SetParamSoc(St,GetControlText(St));
end;

procedure TOF_ASSCOMPTESPE.OnArgument(stArgument: String);
begin
   TToolbarButton97(GetControl('BVALIDER')).OnClick := OnValiderClick;
end;

procedure TOF_ASSCOMPTESPE.OnLoad ;
begin
  SetControlEnabled('SO_LGCPTEGEN',False);
  SetControlEnabled('SO_BOURREGEN',False);
  InitZone('SO_LGCPTEGEN');
  InitZone('SO_BOURREGEN');
  InitZone('SO_GENATTEND');
  InitZone('SO_JALOUVRE');
  InitZone('SO_JALFERME');
  InitZone('SO_OUVREBIL');
  InitZone('SO_FERMEBIL');
  InitZone('SO_OUVREPERTE');
  InitZone('SO_FERMEPERTE');
  InitZone('SO_RESULTAT');
  InitZone('SO_OUVREBEN');
  InitZone('SO_FERMEBEN');
end;

procedure TOF_ASSCOMPTESPE.OnValiderClick(Sender: TObject);
begin
  UpdateChamp('SO_LGCPTEGEN');
  UpdateChamp('SO_BOURREGEN');
  UpdateChamp('SO_GENATTEND');
  UpdateChamp('SO_JALOUVRE');
  UpdateChamp('SO_JALFERME');
  UpdateChamp('SO_OUVREBIL');
  UpdateChamp('SO_FERMEBIL');
  UpdateChamp('SO_OUVREPERTE');
  UpdateChamp('SO_FERMEPERTE');
  UpdateChamp('SO_RESULTAT');
  UpdateChamp('SO_OUVREBEN');
  UpdateChamp('SO_FERMEBEN');
end;

Initialization
RegisterClasses([TOF_ASSCOMPTESPE]) ;
end.
