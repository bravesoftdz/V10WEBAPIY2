unit utof_assanalytique;

interface

uses classes, UTOF, HMsgBox, HTB97, ParamSoc,Section, Axe, Structur, menus,
     mulsecti,HEnt1, venttype,Fe_Main,
{$IFNDEF CCS3}
     corresp,
{$ENDIF}
     rupture, stdctrls;

type
  TOF_ASSANALYTIQUE = Class (TOF)
    procedure OnArgument (stArgument : String ) ; override ;
     procedure OnLoad ; override ;
     procedure OnUpdate; override ;
     procedure OnAxeAnaClick(Sender : TObject);
     procedure OnStructSectionClick(Sender : TObject);
     procedure OnSectionClick(Sender : TObject);
     procedure OnVentilationClick(Sender : TObject);
     procedure OnPlanCorresp1Click(Sender : TObject);
     procedure OnPlanCorresp2Click(Sender : TObject);
     procedure OnPlanRupture1Click(Sender : TObject);
     procedure OnPlanRupture2Click(Sender : TObject);
// ajout me
     procedure OnModifAxeClick(Sender: TObject);
  private
     procedure InitZone (St : string);
     procedure UpdateChamp (St : string);
  end;


implementation

procedure TOF_ASSANALYTIQUE.InitZone(St: string);
var v : variant;
begin
  if GetControl(St) is TCheckBox then
  begin
    SetControlChecked(St,GetParamSoc(St)) ;
  end else SetControlText(St,getParamSoc(St));
end;

procedure TOF_ASSANALYTIQUE.UpdateChamp(St: string);
begin
  SetParamSoc(St,GetControlText(St));
end;

procedure TOF_ASSANALYTIQUE.OnArgument(stArgument: String);
var popcorresp, poprupture : TPopupMenu;
    menucorresp1,menucorresp2,menurupture1,menurupture2 : TMEnuItem;
begin
   TToolbarButton97(GetControl('BPERSOAXE')).OnClick := OnAxeAnaClick ;
   TToolbarButton97(GetControl('BSTRUCTSECTION')).OnClick := OnStructSectionClick;
   TToolbarButton97(GetControl('BSECTION')).OnClick := OnSectionClick;
   TToolbarButton97(GetControl('BVENTILATION')).OnClick := OnVentilationClick;
   popcorresp := TPopUpMenu(GetControl('POPCORRESP'));
   poprupture := TPopUpMenu(GetControl('POPRUPTURE'));
   menucorresp1 := popcorresp.Items[0];
   menucorresp1.OnClick := OnPlanCorresp1Click;
   menucorresp2 := popcorresp.Items[1];
   menucorresp2.OnClick := OnPlanCorresp2Click;
   menurupture1 := poprupture.Items[0];
   menurupture1.OnClick := OnPlanRupture1Click;
   menurupture2 := poprupture.Items[1];
   menurupture2.OnClick := OnPlanRupture2Click;
   // ajout me
   TToolbarButton97(GetControl('BMODIFAXE')).OnClick    := OnModifAxeClick;
end;


procedure TOF_ASSANALYTIQUE.OnLoad ;
begin
  InitZone('SO_ZSAISIEANAL');
end;

procedure TOF_ASSANALYTIQUE.OnAxeAnaClick(Sender: TObject);
begin
FicheAxeAna ('');
end;

procedure TOF_ASSANALYTIQUE.OnStructSectionClick(Sender: TObject);
begin
ParamPlanAnal('') ;
end;

procedure TOF_ASSANALYTIQUE.OnSectionClick(Sender: TObject);
begin
MulticritereSection(taModif);
end;

procedure TOF_ASSANALYTIQUE.OnVentilationClick(Sender: TObject);
begin
ParamVentilType;
end;

procedure TOF_ASSANALYTIQUE.OnPlanCorresp1Click(Sender: TObject);
begin
{$IFNDEF CCS3}
ParamCorresp('A1') ;
{$ENDIF}
end;

procedure TOF_ASSANALYTIQUE.OnPlanCorresp2Click(Sender: TObject);
begin
{$IFNDEF CCS3}
ParamCorresp('A2')
{$ENDIF}
end;

procedure TOF_ASSANALYTIQUE.OnPlanRupture1Click(Sender: TObject);
begin
PlanRupture('RU1') ;
end;

procedure TOF_ASSANALYTIQUE.OnPlanRupture2Click(Sender: TObject);
begin
PlanRupture('RU2') ;
end;

// ajout me
procedure TOF_ASSANALYTIQUE.OnModifAxeClick(Sender: TObject);
begin
           V_PGI.ExtendedFieldSelection := 'A';
           AGLLanceFiche('CP','MULGENERAUX','','','');
end;

procedure TOF_ASSANALYTIQUE.OnUpdate;
begin
  UpdateChamp('SO_ZSAISIEANAL');
end;

Initialization
RegisterClasses([TOF_ASSANALYTIQUE]);
end.
