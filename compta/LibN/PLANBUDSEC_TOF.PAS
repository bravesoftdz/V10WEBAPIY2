{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 24/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANBUDSEC ()
Mots clefs ... : TOF;PLANBUDSEC
*****************************************************************}
Unit PLANBUDSEC_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db, dbtables, mul, FE_Main, QRS1,
{$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF;

procedure PlanBudSec(Axe : String ; UnCpte : String ; DuBouton : Boolean);

Type
  TOF_PLANBUDSEC = Class (TOF)
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

procedure PlanBudSec(Axe : String ; UnCpte : String ; DuBouton : Boolean) ;
begin
  AGLLanceFiche('CP','CPPLANBUDSECT','','',Axe+';'+UnCpte);
end;

procedure TOF_PLANBUDSEC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDSEC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDSEC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDSEC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDSEC.OnArgument (S : String ) ;
var
  i : Integer;
begin
  Inherited ;

  SetControlText('BS_AXE', ReadTokenSt(S));
  SetControlText('BS_BUDSECT', S);
  SetControlText('BS_BUDSECT_', S);

  TComboBox(GetControl('BS_SENS', True)).ItemIndex := 0;

  FTri := THRadioGroup(GetControl('FTRI', True));
  FTri.OnClick := TriOnClick;

  TYPEEDITION := THRadioGroup(GetControl('TYPEEDITION', True));
  TYPEEDITION.OnClick := TypeOnClick;
  TypeOnClick(nil);

  // Libellé des tables libres (Evite des requêtes à chaque fiche)
  for i := 0 to 9 do
    SetControlText('T'+IntToStr(i), RechDom('TTTABLESLIBRESBUDS', 'D0'+IntToStr(i), False));

  Ecran.HelpContext := 15145100;
end ;

procedure TOF_PLANBUDSEC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDSEC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDSEC.OnCancel () ;
begin
  Inherited ;
end ;

// Choix du tri de l'édition : Par compte général ou Libellé
procedure TOF_PLANBUDSEC.TriOnClick(Sender: TObject);
begin
  if (FTri.ItemIndex = 0) then SetControlText('XX_ORDERBY' , 'BS_BUDSECT')
                          else SetControlText('XX_ORDERBY' , 'BS_LIBELLE')
end;

procedure TOF_PLANBUDSEC.TypeOnClick(Sender: TObject);
var
  b : Boolean;
begin
  b := (TYPEEDITION.ItemIndex = 0);
  if b then TFQRS1(Ecran).CodeEtat := 'BSE'
       else TFQRS1(Ecran).CodeEtat := 'BS2';
end;

Initialization
  registerclasses ( [ TOF_PLANBUDSEC ] ) ; 
end.
