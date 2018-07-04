{***********UNITE*************************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 23/10/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TIERSFRAIS ()
Mots clefs ... : TOF;TIERSFRAIS
*****************************************************************}
Unit TIERSFRAIS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     mul,
     FE_main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTOB,
     SaisieList,
     uTableFiltre,
     htb97,
     AGLInit,
     menus
     ;

Type
  TOF_TIERSFRAIS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  Private
    Action : string ;
    TF : TTableFiltre;
    Procedure MnVisuFrais_OnClick(Sender: TObject);

    procedure DoSetNavigate(Sender:TObject);

  end ;

Implementation

Uses
  wCommuns;

procedure TOF_TIERSFRAIS.OnArgument (S : String ) ;
begin
  Inherited ;

  Action := GetArgumentValue(S, 'ACTION');

  TF := TFSaisieList( ecran ).LeFiltre;
  TF.OnSetNavigate:= DoSetNavigate;

  if Assigned(GetControl('MNVISUFRAIS')) then
    TMenuItem(GetControl('MNVISUFRAIS')).OnClick := MnVisuFrais_OnClick;

end ;

procedure TOF_TIERSFRAIS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_TIERSFRAIS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_TIERSFRAIS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_TIERSFRAIS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_TIERSFRAIS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_TIERSFRAIS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_TIERSFRAIS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_TIERSFRAIS.MnVisuFrais_OnClick(Sender: TObject);
begin
  AGLLanceFiche('GC','GCPORT','',GetField('GTF_CODEPORT'),'ACTION=' + iif(Action='CREATION', 'MODIFICATION', Action) + ';MONOFICHE') ;
end;

procedure TOF_TIERSFRAIS.DoSetNavigate(Sender: TObject);
Var
  ListeNonVide: boolean;
begin
  ListeNonVide := TF.Recno<>0;
  SetControlProperty('BPOPMENU', 'ENABLED', (TF.CanReload and ListeNonVide) );
end;

Initialization
  registerclasses ( [ TOF_TIERSFRAIS ] ) ;
end.
