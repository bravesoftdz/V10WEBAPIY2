{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANJOUR ()
Mots clefs ... : TOF;PLANJOUR
*****************************************************************}
Unit PLANJOUR_TOF ;

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

procedure CPLanceFiche_PLANJOUR(const Compte : String);

Type
  TOF_PLANJOUR = Class (TOF)
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

procedure CPLanceFiche_PLANJOUR(const Compte : String);
begin
  AGLLanceFiche('CP','CPPLANJOUR','','',Compte);
end;

procedure TOF_PLANJOUR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PLANJOUR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PLANJOUR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PLANJOUR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PLANJOUR.OnArgument (S : String ) ;
begin
  Inherited ;
  //28/10/2005 Rajout de l'aide
  Ecran.HelpContext:=7229000;

  SetControlText('J_JOURNAL', S);
  SetControlText('J_JOURNAL_', S);
  
  TComboBox(GetControl('J_NATUREJAL', True)).ItemIndex := 0;

  FTri := THRadioGroup(GetControl('FTRI', True));
  FTri.OnClick := TriOnClick;

  TYPEEDITION := THRadioGroup(GetControl('TYPEEDITION', True));
  TYPEEDITION.OnClick := TypeOnClick;
  TypeOnClick(nil);
end ;

procedure TOF_PLANJOUR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PLANJOUR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PLANJOUR.OnCancel () ;
begin
  Inherited ;
end ;

// Choix du tri de l'édition : Par journal ou Libellé
procedure TOF_PLANJOUR.TriOnClick(Sender: TObject);
begin
  if (FTri.ItemIndex = 0) then SetControlText('XX_ORDERBY' , 'J_JOURNAL')
                          else SetControlText('XX_ORDERBY' , 'J_LIBELLE')
end;

// Choix du type d'édition : Liste ou Fiche
procedure TOF_PLANJOUR.TypeOnClick(Sender: TObject);
begin
  if (TYPEEDITION.ItemIndex = 0) then TFQRS1(Ecran).CodeEtat := 'JO2'
                                 else TFQRS1(Ecran).CodeEtat := 'JOU';
  {JP 02/11/05 : On n'autorise l'export qu'en mode liste}
  if (TYPEEDITION.ItemIndex > 0) then SetControlChecked('FLISTE', False);
  SetControlEnabled('FLISTE', (TYPEEDITION.ItemIndex = 0));
end;

Initialization
  registerclasses ( [ TOF_PLANJOUR ] ) ; 
end.
