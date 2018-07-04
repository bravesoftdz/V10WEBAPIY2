{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANSECT ()
Mots clefs ... : TOF;PLANSECT
*****************************************************************}
Unit PLANSECT_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul, FE_Main, QRS1,
{$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF;

procedure CPLanceFiche_PLANSECT(const Axe, Compte : String);

Type
  TOF_PLANSECT = Class (TOF)
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

procedure CPLanceFiche_PLANSECT(const Axe, Compte : String);
begin
  AGLLanceFiche('CP','CPPLANSECT','','',Axe+';'+Compte);
end;

procedure TOF_PLANSECT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PLANSECT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PLANSECT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PLANSECT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PLANSECT.OnArgument (S : String ) ;
var
  i : Integer;
begin
  Inherited ;
  //28/10/2005 Rajout de l'aide
  Ecran.HelpContext:=7199000;

  SetControlText('S_AXE', ReadTokenST(S));
  SetControlText('S_SECTION', S);
  SetControlText('S_SECTION_', S);
  
  TComboBox(GetControl('S_SENS', True)).ItemIndex := 0;

  FTri := THRadioGroup(GetControl('FTRI', True));
  FTri.OnClick := TriOnClick;

  TYPEEDITION := THRadioGroup(GetControl('TYPEEDITION', True));
  TYPEEDITION.OnClick := TypeOnClick;
  TypeOnClick(nil);

  // Libellé des tables libres (Evite des requêtes à chaque fiche)
  for i := 0 to 9 do
    SetControlText('T'+IntToStr(i), RechDom('TTTABLESLIBRESSEC', 'S0'+IntToStr(i), False));
end ;

procedure TOF_PLANSECT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PLANSECT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PLANSECT.OnCancel () ;
begin
  Inherited ;
end ;

// Choix du tri de l'édition : Par Section ou Libellé
procedure TOF_PLANSECT.TriOnClick(Sender: TObject);
begin
  if (FTri.ItemIndex = 0) then SetControlText('XX_ORDERBY' , 'S_SECTION')
                          else SetControlText('XX_ORDERBY' , 'S_LIBELLE')
end;

// Choix du type d'édition : Liste ou Fiche
procedure TOF_PLANSECT.TypeOnClick(Sender: TObject);
var
  b : Boolean;
begin
  b := (TYPEEDITION.ItemIndex = 0);
  if b then TFQRS1(Ecran).CodeEtat := 'AN2'
       else TFQRS1(Ecran).CodeEtat := 'ANA';

  SetControlEnabled('CLEREPART', not b);

  {JP 02/11/05 : On n'autorise l'export qu'en mode liste}
  if not B then SetControlChecked('FLISTE', False);
  SetControlEnabled('FLISTE', B);
end;

Initialization
  registerclasses ( [ TOF_PLANSECT ] ) ; 
end.
