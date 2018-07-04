{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 18/10/2007
Modifié le ... :   /  /
Description .. : Source TOF ancêtre des FICHES communes à la GIGA et la GRC
Mots clefs ... : TOF;ANCETRE_GIGA_GRC
*****************************************************************}
Unit utofAFANCETRE_GIGA_GRC ;

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
     UtofAfTraducChampLibre,
     HMsgBox,
     UtilGC,
     UTOF ;

Type
                //PL 18/10/07 : pour gérer tout ce qui est commun GIGA/GRC et qui ne doit être géré qu'en IFDEF AFFAIRE
  TOF_ANCETRE_GIGA_GRC = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    public
    giNumMemoLine : integer; // Numéro de ligne du Mémo pour mettre la clé Ressource YTC
    Memo : TMemo;
  end ;

Implementation

procedure TOF_ANCETRE_GIGA_GRC.OnArgument (S : String ) ;
begin
  Inherited ;

// On suppose qu'on a les champs YTC_RESSOURCE et T_MOISCLOTURE que si on a au-moins le champ Ressource1
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then
    // On cache l'onglet Tables libres Ressources si on n'est pas en contexte affaire
    SetControlVisible ('PRESSOURCE', false)
  else
    begin
    // on affiche l'onglet Tables libres Ressources
    SetControlVisible ('PRESSOURCE', true);

    // Gestion des libellés libres des Ressources libres
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');

    // Si on a un Memo et qu'on ne veut pas bloquer (giNumMemoLine<> -1),
    // on peut mettre à jour le Memo.lines
    if Assigned(Memo) and (giNumMemoLine <> -1) then
      Memo.lines[giNumMemoLine] := ',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3,YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3';

    // On cache les champs MoisCloture si on n'est pas en contexte Scot
    if (GetControl('T_MOISCLOTURE') <> nil)  then
      if not (ctxscot in V_PGI.PGICOntexte) then
         begin
         SetControlVisible ('T_MOISCLOTURE', false);
         SetControlVisible ('T_MOISCLOTURE_', false);
         SetControlVisible ('TT_MOISCLOTURE', false);
         SetControlVisible ('TT_MOISCLOTURE_', false);
         end;
    end;
  end;

end ;

procedure TOF_ANCETRE_GIGA_GRC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ANCETRE_GIGA_GRC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ANCETRE_GIGA_GRC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_ANCETRE_GIGA_GRC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_ANCETRE_GIGA_GRC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_ANCETRE_GIGA_GRC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ANCETRE_GIGA_GRC.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_ANCETRE_GIGA_GRC ] ) ; 
end.
