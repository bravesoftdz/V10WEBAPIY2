{ Unité : Source TOF de la FICHE : TRCIBREFERENCE ()
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            08/09/03    JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit TRCIBREFERENCE_TOF ;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db,dbtables, mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, SysUtils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, uTob, UTableFiltre, Constantes,
  SaisieList;

type
  TOF_TRCIBREFERENCE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TF : TTableFiltre;
  end ;

procedure TRLanceFiche_CibReference(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_CibReference(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnNew ;
{---------------------------------------------------------------------------------------}
begin
  Inherited;
  TF.StartUpdate;
  TF.PutValue('TCI_BANQUE', CodeCibRef);
  TF.EndUpdate;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnDelete ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnUpdate ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;

end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  TF := TFSaisieList(Ecran).LeFiltre;
  Ecran.HelpContext := 150;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnDisplay () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;

end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCIBREFERENCE.OnCancel () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

initialization
  RegisterClasses ( [ TOF_TRCIBREFERENCE ] ) ;
end.
