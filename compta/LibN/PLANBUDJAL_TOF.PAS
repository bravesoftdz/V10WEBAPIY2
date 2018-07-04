{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... : 11/08/2005
Description .. : Source TOF de la FICHE : PLANBUDJAL ()
Mots clefs ... : TOF;PLANBUDJAL
*****************************************************************}
Unit PLANBUDJAL_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db, dbtables, mul, FE_Main, QRS1,
{$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF;

procedure PlanBudJal(Axe : String ;UnCpte : String ; DuBouton : Boolean);

Type
  TOF_PLANBUDJAL = class (TOF)
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
    procedure GereCodeEtat;
    procedure FormAfterShow;
  end ;

implementation

uses
  LicUtil;

procedure PlanBudJal(Axe : String ;UnCpte : String ; DuBouton : Boolean);
begin
  AGLLanceFiche('CP','CPPLANBUDJAL','','',Axe+';'+UnCpte);
end;

procedure TOF_PLANBUDJAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDJAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDJAL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDJAL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDJAL.OnArgument (S : String ) ;
begin
  Inherited ;


  SetControlText('BJ_AXE', ReadTokenSt(S));
  SetControlText('BJ_BUDJAL', S);
  SetControlText('BJ_BUDJAL_', S);

  TComboBox(GetControl('BJ_SENS', True)).ItemIndex := 0;

  FTri := THRadioGroup(GetControl('FTRI', True));
  FTri.OnClick := TriOnClick;

  TYPEEDITION := THRadioGroup(GetControl('TYPEEDITION', True));
  TYPEEDITION.OnClick := TypeOnClick;
  TypeOnClick(nil);

  Ecran.HelpContext := 15165100;

  {Gestion du paramétrage des états}
  TFQRS1(Ecran).ChoixEtat  := False;
  TFQRS1(Ecran).ParamEtat  := V_PGI.Superviseur or (V_PGI.Password = CryptageSt(DayPass(Date)));
  {Pour afficher le bon état}
  TFQRS1(Ecran).OnAfterFormShow := FormAfterShow;
end ;

procedure TOF_PLANBUDJAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDJAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PLANBUDJAL.OnCancel () ;
begin
  Inherited ;
end ;

// Choix du tri de l'édition : Par compte général ou Libellé
procedure TOF_PLANBUDJAL.TriOnClick(Sender: TObject);
begin
  if (FTri.ItemIndex = 0) then SetControlText('XX_ORDERBY' , 'BJ_BUDJAL')
                          else SetControlText('XX_ORDERBY' , 'BJ_LIBELLE')
end;

{---------------------------------------------------------------------------------------}
procedure TOF_PLANBUDJAL.TypeOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  GereCodeEtat;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_PLANBUDJAL.FormAfterShow;
{---------------------------------------------------------------------------------------}
begin
  GereCodeEtat;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_PLANBUDJAL.GereCodeEtat;
{---------------------------------------------------------------------------------------}
begin
  if TypeEdition.ItemIndex = 0 then TFQRS1(Ecran).CodeEtat := 'BJA'
                               else TFQRS1(Ecran).CodeEtat := 'BJ2';
  SetControlText('FETAT', TFQRS1(Ecran).CodeEtat);
end;


Initialization
  registerclasses ( [ TOF_PLANBUDJAL ] ) ;
end.
