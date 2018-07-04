unit utof_assanalytique;

interface

uses classes, UTOF, HMsgBox, HTB97, ParamSoc,CPSection_TOM, { BVE 03.05.07 Structur} CPSTRUCTURE_TOF, menus,
     UTOFMULPARAMGEN, HEnt1, UTOTVENTILTYPE,Fe_Main,
     CPAXE_TOM, //Axe, 
{$IFNDEF CCS3}
     CORRESP_TOM,
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
     procedure OnPlanCorresp3Click(Sender : TObject);
     procedure OnPlanCorresp4Click(Sender : TObject);
     procedure OnPlanCorresp5Click(Sender : TObject);
     procedure OnPlanRupture1Click(Sender : TObject);
     procedure OnPlanRupture2Click(Sender : TObject);
     procedure OnPlanRupture3Click(Sender : TObject);
     procedure OnPlanRupture4Click(Sender : TObject);
     procedure OnPlanRupture5Click(Sender : TObject);
// ajout me
     procedure OnModifAxeClick(Sender: TObject);
  private
     procedure InitZone (St : string);
     procedure UpdateChamp (St : string);
  end;


implementation

procedure TOF_ASSANALYTIQUE.InitZone(St: string);
begin
  if GetControl(St) is TCheckBox then
  begin
    SetControlChecked(St,GetParamSocSecur(St,'')) ;{Lek 100206}
  end else SetControlText(St,getParamSocSecur(St,'')); {Lek 100206}
end;

procedure TOF_ASSANALYTIQUE.UpdateChamp(St: string);
begin
  SetParamSoc(St,GetControlText(St));
end;

procedure TOF_ASSANALYTIQUE.OnArgument(stArgument: String);
var popcorresp, poprupture : TPopupMenu;
    menucorresp1,menucorresp2,menucorresp3,menucorresp4,menucorresp5 : TMEnuItem;
    menurupture1,menurupture2,menurupture3,menurupture4,menurupture5 : TMEnuItem;
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
   menucorresp3 := popcorresp.Items[2];						// Ajout
   menucorresp3.OnClick := OnPlanCorresp3Click;		// Ajout
   menucorresp4 := popcorresp.Items[3];						// Ajout
   menucorresp4.OnClick := OnPlanCorresp4Click;		// Ajout
   menucorresp5 := popcorresp.Items[4];						// Ajout
   menucorresp5.OnClick := OnPlanCorresp5Click;		// Ajout
   menurupture1 := poprupture.Items[0];
   menurupture1.OnClick := OnPlanRupture1Click;
   menurupture2 := poprupture.Items[1];
   menurupture2.OnClick := OnPlanRupture2Click;
   menurupture3 := poprupture.Items[2];						// Ajout
   menurupture3.OnClick := OnPlanRupture3Click;		// Ajout
   menurupture4 := poprupture.Items[3];						// Ajout
   menurupture4.OnClick := OnPlanRupture4Click;		// Ajout
   menurupture5 := poprupture.Items[4];						// Ajout
   menurupture5.OnClick := OnPlanRupture5Click;		// Ajout

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
CPLanceFiche_MULSection('C;7178000');
end;

procedure TOF_ASSANALYTIQUE.OnVentilationClick(Sender: TObject);
begin
ParamVentilType;
end;

procedure TOF_ASSANALYTIQUE.OnPlanCorresp1Click(Sender: TObject);
begin
{$IFNDEF CCS3}
CCLanceFiche_Correspondance('A1') ;		// Modification
{$ENDIF}
end;

procedure TOF_ASSANALYTIQUE.OnPlanCorresp2Click(Sender: TObject);
begin
{$IFNDEF CCS3}
CCLanceFiche_Correspondance('A2')		// Modification
{$ENDIF}
end;

procedure TOF_ASSANALYTIQUE.OnPlanCorresp3Click(Sender: TObject);	// Ajout
begin
{$IFNDEF CCS3}
CCLanceFiche_Correspondance('A3')
{$ENDIF}
end;

procedure TOF_ASSANALYTIQUE.OnPlanCorresp4Click(Sender: TObject);	// Ajout
begin
{$IFNDEF CCS3}
CCLanceFiche_Correspondance('A4')
{$ENDIF}
end;

procedure TOF_ASSANALYTIQUE.OnPlanCorresp5Click(Sender: TObject);	// Ajout
begin
{$IFNDEF CCS3}
CCLanceFiche_Correspondance('A5')
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

procedure TOF_ASSANALYTIQUE.OnPlanRupture3Click(Sender: TObject);	// Ajout
begin
PlanRupture('RU3') ;
end;

procedure TOF_ASSANALYTIQUE.OnPlanRupture4Click(Sender: TObject);	// Ajout
begin
PlanRupture('RU4') ;
end;

procedure TOF_ASSANALYTIQUE.OnPlanRupture5Click(Sender: TObject);	// Ajout
begin
PlanRupture('RU5') ;
end;

// ajout me
procedure TOF_ASSANALYTIQUE.OnModifAxeClick(Sender: TObject);
begin
           V_PGI.ExtendedFieldSelection := '1';
           AGLLanceFiche('CP','MULGENERAUX','','','');
end;

procedure TOF_ASSANALYTIQUE.OnUpdate;
begin
  UpdateChamp('SO_ZSAISIEANAL');
end;

Initialization
RegisterClasses([TOF_ASSANALYTIQUE]);
end.
