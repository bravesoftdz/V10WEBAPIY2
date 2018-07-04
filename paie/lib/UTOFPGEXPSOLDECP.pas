{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 13/02/2006
Modifié le ... :   /  /
Description .. : Gestion des recapitulatifs congés payés et RTT des salariés
Suite ........ : Lancer a partir du menu :
Suite ........ :   Administration >Traitement > Export Solde CP
Suite ........ :   Num : 49814
Mots clefs ... :   TOF;PGEXPSOLDECP
*****************************************************************}
unit UTOFPGEXPSOLDECP;

interface

Uses SysUtils, HCtrls, HEnt1, Classes, Vierge,
   HMsgBox,  UTOF
  {$IFDEF EAGLCLIENT}
    ,eFiche
  {$ELSE}
  {$ENDIF}
  ;

Type
  TOF_PGEXPSOLDECP = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end;

implementation
 uses PgOutilseAgl;

procedure TOF_PGEXPSOLDECP.OnUpdate ;
begin
  Inherited ;
  ChargeTobSalRecap('A', StrToDate(GetControlText('DATE_INTEGRATION')));
  PGIInfo('Le traitement s''est correctement terminé','Export Solde CP');
  TFVierge(Ecran).BFerme.Click;
End ;

procedure TOF_PGEXPSOLDECP.OnArgument (S : String ) ;
begin
  Inherited ;
  SetControltext('DATE_INTEGRATION',DateToStr(Date));
End ;

procedure TOF_PGEXPSOLDECP.OnClose ;
Var
  Date_int : THEdit;
begin
  Inherited ;
  Date_int := THEdit(GetControl('DATE_INTEGRATION'));
  if Date_int <> nil then
     if (not(IsValidDate(Date_int.text))) then
        SetControltext('DATE_INTEGRATION',DateToStr(Date));
end ;

Initialization
  registerclasses ([TOF_PGEXPSOLDECP]) ;
End.




