{***********UNITE*************************************************
Auteur  ...... : GF
Créé le ...... : 02/03/2001
Modifié le ... : 28/03/2001
Description .. : Source TOF de la Fiche
Suite ........ : (cf UTOFAFPLANNINGEXEMPLE)
Mots clefs ... : TOF;BTPLANNINGEXEMPLE
*****************************************************************}
unit UTOFBTPLANNINGEXEMPLE;

interface

uses
  Classes,
  HEnt1,
  UTOF,
  UTob,
  Hplanning,
  controls,
  PlanUtil;
  //BTPlanningConfig,
  //HrUtils;

type
  TOF_BTPLANNINGEXEMPLE = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnClose; override;

  public
    TobEtats	: Tob;
    TobItems	: Tob;
    TOBRes		: Tob;
    TobModelePlanning	: TOB;
    Planning	: THPlanningBTP;
    procedure ParamPlanningEx;

  end;

implementation

var     TobPlannings			: array[0..1] of RecordPlanning;

procedure TOF_BTPLANNINGEXEMPLE.OnArgument(stArgument: string);
begin
  inherited;

  TobModelePlanning := LaTob;

//  Planning := THPlanning(Ecran.FindComponent('Planning'));
	Planning := THplanningBTP.create(ecran);
  Planning.Align := alclient;
  Planning.Parent := ecran;

  ParamPlanningEx;

end;

procedure TOF_BTPLANNINGEXEMPLE.OnNew;
begin
  inherited;
end;

procedure TOF_BTPLANNINGEXEMPLE.OnDelete;
begin
  inherited;
end;

procedure TOF_BTPLANNINGEXEMPLE.OnUpdate;
begin
  inherited;
end;

procedure TOF_BTPLANNINGEXEMPLE.OnLoad;
begin
  inherited;
end;

procedure TOF_BTPLANNINGEXEMPLE.OnClose;
begin
  inherited;
(*
  with TobPlannings[0] do
  		 begin
    	 if TobItems <> nil then TobItems.free;
    	 if TobEtats <> nil then TobEtats.free;
    	 if TobRes <> nil then TobRes.free;
    	 if TobCols <> nil then TobCols.free;
    	 if TobRows <> nil then TobRows.free;
    	 if TobEvents <> nil then TobEvents.free;
       end;
  *)
	if Planning <> nil then planning.free;
end;

procedure TOF_BTPLANNINGEXEMPLE.ParamPlanningEx;
begin

  with TobPlannings[0] do
		   begin
    	 TobItems := nil;
    	 TobEtats := nil;
    	 TobRes := nil;
    	 TobCols := nil;
    	 TobRows := nil;
    	 TobEvents := nil;
  		 end;

  
  // Chargement des paramètres du planning
  try
    ChargeParamPlanning(TobPlannings[0], Planning, TobModelePlanning, V_PGI.DateEntree, False);
  finally
  end;
end;

initialization
  registerclasses([TOF_BTPLANNINGEXEMPLE]);

end.
