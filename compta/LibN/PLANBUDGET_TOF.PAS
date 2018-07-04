{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 24/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANBUDGET ()
Mots clefs ... : TOF;PLANBUDGET
*****************************************************************}
Unit PLANBUDGET_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db, dbtables, mul, FE_Main, QRS1,
{$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF;

procedure PlanBudget(UnCpte : String ; DuBouton : Boolean);

Type
  TOF_PLANBUDGET = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    FTri, TYPEEDITION : THRadioGroup;
    procedure TriOnClick(Sender : TObject);
    procedure TypeOnClick(Sender : TObject);
  end ;

Implementation

procedure PlanBudget(UnCpte : String ; DuBouton : Boolean);
begin
  AGLLanceFiche('CP','CPPLANBUDGET','','',UnCpte);
end;

procedure TOF_PLANBUDGET.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDGET.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDGET.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDGET.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDGET.OnArgument (S : String ) ;
var
  i : Integer;
begin
  Inherited ;

  SetControlText('BG_BUDGENE', S);
  SetControlText('BG_BUDGENE_', S);

  TComboBox(GetControl('BG_SENS', True)).ItemIndex := 0;

  FTri := THRadioGroup(GetControl('FTRI', True));
  FTri.OnClick := TriOnClick;

  TYPEEDITION := THRadioGroup(GetControl('TYPEEDITION', True));
  TYPEEDITION.OnClick := TypeOnClick;
  TypeOnClick(nil);

  // Libellé des tables libres (Evite des requêtes à chaque fiche)
  for i := 0 to 9 do
    SetControlText('T'+IntToStr(i), RechDom('TTTABLESLIBRESBUDG', 'B0'+IntToStr(i), False));

  Ecran.HelpContext := 15125100;
end ;

procedure TOF_PLANBUDGET.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDGET.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDGET.OnCancel () ;
begin
  Inherited ;
end ;

// Choix du tri de l'édition : Par compte général ou Libellé
procedure TOF_PLANBUDGET.TriOnClick(Sender: TObject);
begin
  if (FTri.ItemIndex = 0) then SetControlText('XX_ORDERBY' , 'BG_BUDGENE')
                          else SetControlText('XX_ORDERBY' , 'BG_LIBELLE')
end;

procedure TOF_PLANBUDGET.TypeOnClick(Sender: TObject);
var
  b : Boolean;
begin
  b := (TYPEEDITION.ItemIndex = 0);
  if b then TFQRS1(Ecran).CodeEtat := 'BGE'
       else TFQRS1(Ecran).CodeEtat := 'BG2';
end;

Initialization
  registerclasses ( [ TOF_PLANBUDGET ] ) ; 
end.
