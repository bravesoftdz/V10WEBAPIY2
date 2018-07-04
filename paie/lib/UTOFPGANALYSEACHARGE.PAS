{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 30/05/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPGANALYSEACHARGE ()
Mots clefs ... : TOF;UTOFPGANALYSEACHARGE Tobviewer Personnes à charge
*****************************************************************}
Unit UTOFPGANALYSEACHARGE;

Interface

Uses StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     EdtREtat,
{$else}
     eMul,
     UtilEAGL,
{$ENDIF}
     forms,
     uTob,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     EntPaie,
     PGOutils2,
     stat ;

Type
  TOF_UTOFPGANALYSEACHARGE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure ExitEdit(Sender: TObject);
  end ;

Implementation

procedure TOF_UTOFPGANALYSEACHARGE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.OnArgument (S : String ) ;
var
  Defaut : ThEdit;
begin
  Inherited ;

  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut <> nil then
    Defaut.OnExit := ExitEdit;

  Defaut := ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut <> nil then
    Defaut.OnExit := ExitEdit;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFPGANALYSEACHARGE.ExitEdit(Sender: TObject);
var
  edit : thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit,10);
end;


Initialization
  registerclasses ( [ TOF_UTOFPGANALYSEACHARGE ] ) ; 
end.
