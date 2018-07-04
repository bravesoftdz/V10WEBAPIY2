{***********UNITE*************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 26/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DPSAISIE_TIERS ()
Mots clefs ... : TOF;DPSAISIE_TIERS
*****************************************************************}
Unit DPTOFSAISIE_TIERS ;

Interface

Uses  Classes,
{$IFNDEF EAGLCLIENT}
        Fe_Main,
{$else}
     MaineAGL,
{$ENDIF}
     sysutils,
     HCtrls,
     UTOF
     ;

Type
  TOF_DPSAISIE_TIERS = Class (TOF)
    procedure OnArgument (StArgument : String ) ; override ;
  end ;

Function DPLanceFiche_SaisieTiers(Argument:String):string;

Implementation

function DPLanceFiche_SaisieTiers(Argument:String):string;
begin
result:=AGLLanceFiche ('DP','DPSAISIE_TIERS','','',Argument);
end;

procedure TOF_DPSAISIE_TIERS.OnArgument (Stargument : String ) ;
Var critere, valeur, champ : string ;
    X : integer;
begin
  Inherited ;
           // on récupère les paramètres éventuels
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'TIERS' then  SetControlText('TIERS',valeur)
        else if Champ = 'AUXI'  then begin
             SetControlText('AUXILIAIRE',valeur);
             SetControlEnabled ('TIERS',false);
             end
        else if Champ = 'NATUREAUXI'  then SetControlText('NATUREAUXI',valeur)
        else if Champ = 'NOM'  then SetControlText('NOM1',valeur)
        else if Champ = 'NOMPER'  then SetControlText('NOMPER',valeur)
        else if Champ = 'ERR'  then SetControlCaption('TERR',valeur)
        else if Champ = 'GUIDPER' then SetControlText('GUIDPER',valeur);
        end;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    end;
end ;

Initialization
  registerclasses ( [ TOF_DPSAISIE_TIERS ] ) ;
end.
