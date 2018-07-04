{***********UNITE*************************************************
Auteur  ...... : MCD
Créé le ...... : 17/10/2002
Modifié le ... : 18/11/2002
Description .. : Source TOF de la FICHE : UTOFAFMODIF_UNITE ()
Mots clefs ... : TOF;UTOFAFMODIF_UNITE
*****************************************************************}
Unit UTOFAFMODIF_UNITE ;

Interface

Uses

{$IFDEF EAGLCLIENT}
       Maineagl,
{$ELSE}
   Fe_Main ,
{$ENDIF}
      Controls,AffaireUtil,Classes,ParamSOc, UTOF, HEnt1, Ent1, HmsgBox ,UtilRessource;

Type
  TOF_AFMODIF_UNITE = Class (TOF)
    procedure OnClose                   ; override ;
    procedure OnUpdate                   ; override ;
    procedure OnArgument (S : String ) ; override;
    procedure AppelChangeUnite;
  Private
    Unite : string;
    Erreur:Boolean;
  end ;
  
Procedure AFLanceFiche_Modif_Unite;


Implementation


procedure TOF_AFMODIF_UNITE.OnClose ;

begin
  Inherited ;
If Erreur then begin
   LastError:=(-1);
   LastErrorMsg:='';
   Erreur:=false;
   exit;     // il y eu une erreur, on reste sur la même fiche
   end;
end ;

procedure TOF_AFMODIF_UNITE.OnArgument (S : String ) ;
begin
  Inherited ;
Erreur:=False;
Unite := GetParamSoc('So_AfMesureActivite');
SetControltext ('UNITE',Unite);
end ;


procedure TOF_AFMODIF_UNITE.OnUpdate;
var
text, titre : string;
RepVal : TIOErr;
begin
  Inherited ;

If (GetControlText('UNITE1') = '' ) or
   (GetControlText('UNITE1') = Unite ) then
  begin
    PGIInfo ('Vous n''avez pas changé de valeur ou valeur vide', titre);
    Lasterror:=1;
    Erreur:=True;
    exit;
  end
else
  begin
    text := 'Confirmez-vous le changement de l''unité de référence';
    If (PGIAsk (text,titre) <> mrYes) then exit;

    SetControlText ('SUIVI', 'Mise à jour table en cours');

      // dans tous les cas on détruit les enrgt existant dans la table
    RepVal := Transactions (AppelChangeUnite, 2);
    Case RepVal of
        oeOk :
            BEGIN
            PGIInfo ('La modification de l''unité de référence s''est déroulée avec succés.', titre);
            END;
        oeUnknown, oeSaisie :
            BEGIN
            MessageAlerte('ATTENTION : Changement d''unité de référence non effectué !');
            Exit;
            END ;
    END ;
  end;

end ;

procedure TOF_AFMODIF_UNITE.AppelChangeUnite;
begin
  // stocke nouvelle valeur du paramsoc
  SetParamsoc ('SO_AFMESUREACTIVITE', GetControlText('UNITE1'));

  // processus de changement d'unité
  ChangeUnite;

  // on recharge les paramsoc dans la variable globale
if (V_PGI.IoError = oeOk) then
  ChargeSocieteHalley;

end;

Procedure AFLanceFiche_Modif_Unite;
begin
if ToutSeulAff  then  exit; // quelqu'un d'autre travaille sur la base
try
AGLLanceFiche ('AFF','AFMODIF_UNITE','','','');
finally
Bloqueur ('AffToutSeul',False);
end;
end;

Initialization
  registerclasses ( [ TOF_AFMODIF_UNITE ] ) ;
end.
