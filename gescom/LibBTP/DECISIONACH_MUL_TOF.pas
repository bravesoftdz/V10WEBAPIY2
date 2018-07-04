{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 27/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DECISIONACH_MUL ()
Mots clefs ... : TOF;DECISIONACH_MUL
*****************************************************************}
Unit DECISIONACH_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
		 Hdb,
		 DBGrids,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     M3FP,
     UTOF,
     HTB97,
     DECISIONACHAT_TOF,
 		 ValideDecisionAch	 ;

Type
	TModegestionDecision = (TmoValidation,TmoSelection);
  TOF_DECISIONACH_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private

{$IFDEF EAGLCLIENT}
    	GS : THGrid;
{$ELSE}
    	GS : THDBGrid;
{$ENDIF}
			Modegestion : TModegestionDecision;
			procedure GSDblClick (Sender : Tobject);

  end ;

Implementation

procedure TOF_DECISIONACH_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  GS := THDBGrid (GetControl('Fliste'));
  GS.OnDblClick := GSDblClick;
  if S='VALIDATION' then
  begin
  	Modegestion := TmoValidation;
  end else
  begin
  	Modegestion := TmoSelection;
  end;
  if Modegestion = TmoValidation then
  begin
    TCheckBox(GeTControl('BAE_VALIDE')).Visible := false;
    TCheckBox(GeTControl('BAE_VALIDE')).Checked := false;
  end;
end ;

procedure TOF_DECISIONACH_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_MUL.GSDblClick(Sender: Tobject);
var Numero : integer;
		Valide : string;
begin
	if ModeGestion = TmoSelection then
  begin
    Numero:=GS.datasource.dataset.FindField('BAE_NUMERO').AsInteger;
    if (GS.datasource.dataset.FindField('BAE_VALIDE').AsString='X') then Valide := 'CONSULTATION'
    																													      else Valide := 'MODIFICATION';
    Ouvredecisionnel('NUMDECISIONNEL='+IntToStr(Numero)+';ACTION='+Valide);
    TToolBarButton97(GetControl('BCherche')).click;
  end else
  begin
    Numero:=GS.datasource.dataset.FindField('BAE_NUMERO').AsInteger;
    ValideDecisionnel(Numero);
    TToolBarButton97(GetControl('BCherche')).click;
  end;
end;

Initialization
  registerclasses ( [ TOF_DECISIONACH_MUL ] ) ;
end.

