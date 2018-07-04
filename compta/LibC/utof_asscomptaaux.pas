unit utof_asscomptaaux;

interface

uses classes, UTOF, HMsgBox, HTB97, ParamSoc, StdCtrls;

type
  TOF_ASSCOMPTAAUX = Class (TOF)
    procedure OnArgument (stArgument : String ) ; override ;
     procedure OnLoad ; override ;
     procedure OnValiderClick(Sender: TObject);
  private
     procedure InitZone (St : string);
     procedure UpdateChamp (St : string);
  end;


implementation

procedure TOF_ASSCOMPTAAUX.InitZone(St: string);
begin
  if GetControl(St) is TCheckBox then
  begin
    SetControlChecked(St,GetParamSoc(St)) ;
  end else SetControlText(St,getParamSoc(St));
end;

procedure TOF_ASSCOMPTAAUX.UpdateChamp(St: string);
begin
  SetParamSoc(St,GetControlText(St));
end;

procedure TOF_ASSCOMPTAAUX.OnArgument(stArgument: String);
begin
   TToolbarButton97(GetControl('BVALIDER')).OnClick := OnValiderClick;
end;

procedure TOF_ASSCOMPTAAUX.OnLoad ;
begin
  SetControlEnabled('SO_LGCPTEAUX',False);
  SetControlEnabled('SO_BOURREAUX',False);
  InitZone('SO_LGCPTEAUX');
  InitZone('SO_BOURREAUX');
  InitZone('SO_CPCODEAUXIONLY');
  InitZone('SO_DEFCOLCLI');
  InitZone('SO_CLIATTEND');
  InitZone('SO_LETREFFC');
  InitZone('SO_LETDATEFC');
  InitZone('SO_LETEGALC');
  InitZone('SO_LETREFRC');
  InitZone('SO_LETDATERC');
  InitZone('SO_LETTOLERC');
  InitZone('SO_DEFCOLFOU');
  InitZone('SO_FOUATTEND');
  InitZone('SO_LETREFFF');
  InitZone('SO_LETDATEFF');
  InitZone('SO_LETEGALF');
  InitZone('SO_LETREFRF');
  InitZone('SO_LETDATERF');
  InitZone('SO_LETTOLERF');
  InitZone('SO_DEFCOLSAL');
  InitZone('SO_SALATTEND');
  InitZone('SO_DEFCOLDDIV');
  InitZone('SO_DEFCOLCDIV');
  InitZone('SO_DEFCOLDIV');
  InitZone('SO_DIVATTEND');
end;

procedure TOF_ASSCOMPTAAUX.OnValiderClick(Sender: TObject);
begin
  UpdateChamp('SO_LGCPTEAUX');
  UpdateChamp('SO_BOURREAUX');
  UpdateChamp('SO_CPCODEAUXIONLY');
  UpdateChamp('SO_DEFCOLCLI');
  UpdateChamp('SO_CLIATTEND');
  UpdateChamp('SO_LETREFFC');
  UpdateChamp('SO_LETDATEFC');
  UpdateChamp('SO_LETEGALC');
  UpdateChamp('SO_LETREFRC');
  UpdateChamp('SO_LETDATERC');
  UpdateChamp('SO_LETTOLERC');
  UpdateChamp('SO_DEFCOLFOU');
  UpdateChamp('SO_FOUATTEND');
  UpdateChamp('SO_LETREFFF');
  UpdateChamp('SO_LETDATEFF');
  UpdateChamp('SO_LETEGALF');
  UpdateChamp('SO_LETREFRF');
  UpdateChamp('SO_LETDATERF');
  UpdateChamp('SO_LETTOLERF');
  UpdateChamp('SO_DEFCOLSAL');
  UpdateChamp('SO_SALATTEND');
  UpdateChamp('SO_DEFCOLDDIV');
  UpdateChamp('SO_DEFCOLCDIV');
  UpdateChamp('SO_DEFCOLDIV');
  UpdateChamp('SO_DIVATTEND');
  UpdateChamp('SO_LGCPTEAUX');
  UpdateChamp('SO_BOURREAUX');
  UpdateChamp('SO_CPCODEAUXIONLY');
  UpdateChamp('SO_DEFCOLCLI');
  UpdateChamp('SO_CLIATTEND');
  UpdateChamp('SO_LETREFFC');
  UpdateChamp('SO_LETDATEFC');
  UpdateChamp('SO_LETEGALC');
  UpdateChamp('SO_LETREFRC');
  UpdateChamp('SO_LETDATERC');
  UpdateChamp('SO_LETTOLERC');
  UpdateChamp('SO_DEFCOLFOU');
  UpdateChamp('SO_FOUATTEND');
  UpdateChamp('SO_LETREFFF');
  UpdateChamp('SO_LETDATEFF');
  UpdateChamp('SO_LETEGALF');
  UpdateChamp('SO_LETREFRF');
  UpdateChamp('SO_LETDATERF');
  UpdateChamp('SO_LETTOLERF');
  UpdateChamp('SO_DEFCOLSAL');
  UpdateChamp('SO_SALATTEND');
  UpdateChamp('SO_DEFCOLDDIV');
  UpdateChamp('SO_DEFCOLCDIV');
  UpdateChamp('SO_DEFCOLDIV');
  UpdateChamp('SO_DIVATTEND');
end;


Initialization
RegisterClasses([TOF_ASSCOMPTAAUX]);
end.
