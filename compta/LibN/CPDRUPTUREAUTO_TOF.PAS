{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPDRUPTUREAUTO ()
Mots clefs ... : TOF;CPDRUPTUREAUTO
*****************************************************************}
Unit CPDRUPTUREAUTO_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     MainEagl,  // AGLLanceFiche
{$ELSE}
     Db,
     Dbtables,
     FE_Main,  // AGLLanceFiche
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_CPDRUPTUREAUTO = Class (TOF)

    SEAuto    : THSpinEdit;
    
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnChangeSEAuto   ( Sender : TObject );

  private
    gszNature : string;
    gszPlan   : string;
  end ;

procedure DetailPlanRuptureAuto( vNature, vPlan, vEcranCaption : string );

Implementation

uses Rupanal,
     DRupanal,
     DRupGene,
     UtilPgi ;

////////////////////////////////////////////////////////////////////////////////
procedure DetailPlanRuptureAuto( vNature, vPlan, vEcranCaption : string );
begin
  AGLLanceFiche('CP','CPDRUPTUREAUTO', '', '', vNature + ';' + vPlan + ';' + vEcranCaption);
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPDRUPTUREAUTO.OnArgument (S : String ) ;
var lStArgument : string;
begin
  Inherited ;
  lStArgument   := S;
  gszNature     := ReadTokenSt(lStArgument);
  gszPlan       := ReadTokenSt(lStArgument);
  Ecran.Caption := ReadTokenSt(lStArgument);
  UpdateCaption( Ecran );

  // Récupération des composants
  SEAuto    := THSpinEdit(GetControl('SEAUTO', True));
  // Branchement des événements
  SEAuto.OnChange := OnChangeSEAuto;

  SEAuto.Value := 0 ;

  OnChangeSEAuto(nil) ;
  case gszNature[3] of
    'G' : Ecran.HelpContext := 1370200 ;
    'T' : Ecran.HelpContext := 1380200 ;
    else Ecran.HelpContext := 0 ;
  end ;
end ;

procedure TOF_CPDRUPTUREAUTO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPDRUPTUREAUTO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPDRUPTUREAUTO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPDRUPTUREAUTO.OnDelete ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPDRUPTUREAUTO.OnUpdate ;
var RupClass : TabBool ;
    LibClass : TabLib ;
    i : Byte ;
begin
  Inherited ;
  if SEAuto.Value <= 0 then Exit ;
  for i:=1 to 6 do
  begin
    if TCheckBox(Ecran.FindComponent('Classe'+InttoStr(i))).Checked then
      RupClass[i] := True
    else
      RupClass[i] := False ;
    LibClass[i] := TEdit(Ecran.FindComponent('EClasse'+InttoStr(i))).Text ;
  end;
  DetailPlanRuptureGene(RupClass, LibClass, gszNature, gszPlan) ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPDRUPTUREAUTO.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPDRUPTUREAUTO.OnChangeSEAuto(Sender: TObject);
var i : Integer ;
    C : TComponent ;
begin
  for i := 1 to 6 do
  begin
    C := Ecran.FindComponent('Classe' + InttoStr(i)) ;
    TControl(C).Enabled := (i <= SEAuto.Value) ;
    if not TControl(C).Enabled then TCheckBox(C).Checked:=False ;
    C:= Ecran.FindComponent('EClasse' + InttoStr(i)) ;
    TControl(C).Enabled := (i <= SEAuto.Value) ;
    if not TControl(C).Enabled then TEdit(C).Text := '' ;
    C := Ecran.FindComponent('TEClasse' + InttoStr(i)) ;
    TControl(C).Enabled := (i <= SEAuto.Value) ;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPDRUPTUREAUTO.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPDRUPTUREAUTO ] ) ; 
end.
