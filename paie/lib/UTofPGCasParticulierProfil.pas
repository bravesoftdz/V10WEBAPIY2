{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 29/05/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCASPARTPROFIL ()
Mots clefs ... : TOF;PGCASPARTPROFIL
*****************************************************************}
Unit UTofPGCasParticulierProfil ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     Vierge,
     HMsgBox, 
     UTOF ; 

Type
  TOF_PGCASPARTPROFIL = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PGCASPARTPROFIL.OnUpdate ;
begin
  Inherited ;
  TFVierge(Ecran).Retour := GetControltext('PROFIL');
end ;

procedure TOF_PGCASPARTPROFIL.OnArgument (S : String ) ;
var Theme,Profil : String;
begin
  Inherited ;
  Theme := ReadTokenPipe(S,';');
  Profil := ReadTokenPipe(S,';');
  SetControlProperty('PROFIL','Plus',Theme);
  SetControlText('PROFIL',Profil);
  TFVierge(Ecran).Retour := 'NONMODIFIE';
  TFVierge(Ecran).Caption := 'Saisie profil thème : '+rechDom('PGTHEMEPROFIL',theme,False);
  UpdateCaption(TFVierge(Ecran));
end ;

Initialization
  registerclasses ( [ TOF_PGCASPARTPROFIL ] ) ; 
end.

