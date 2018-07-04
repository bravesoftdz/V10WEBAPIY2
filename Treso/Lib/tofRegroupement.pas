{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/04/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : REGROUPEMENT ()
Mots clefs ... : TOF;REGROUPEMENT
*****************************************************************}
unit tofRegroupement ;

interface

uses
  StdCtrls, Controls, Classes, SaisieList, 
  {$IFDEF EAGLCLIENT}
   MaineAGL,
  {$ELSE}
  db, dbtables, FE_Main, 
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, uTableFiltre, hTB97, UTOF;


procedure TRLanceFiche_Regroupement(Dom : string ; fiche : string ; Range : string ; lequel :string ; arguments : string )  ;

type
  TOF_REGROUPEMENT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  private
    TF: TTableFiltre;
  end;

implementation

uses
  ExtCtrls, Commun, Constantes;

procedure TRLanceFiche_Regroupement(Dom : string ; fiche : string ; Range : string ; lequel :string ; arguments : string )  ;
begin
  AglLanceFiche(Dom, fiche, range, lequel, arguments ); // Regroupement
end ;

procedure TOF_REGROUPEMENT.OnArgument (S : String ) ;
begin
  Inherited ;
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
    TF := TFSaisieList(Ecran).LeFiltre;

end ;

Initialization
  registerclasses ( [ TOF_REGROUPEMENT ] ) ; 
end.

