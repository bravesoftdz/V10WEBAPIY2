{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFEACTIVITEMODIF ()
Mots clefs ... : TOF;AFEACTIVITEMODIF
*****************************************************************}
Unit UTOFAFEACTIVITEMODIF ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls,
{$IFDEF EAGLCLIENT}
   eMul,MaineAGL,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, mul, FE_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF, utob, UTofAfBaseCodeAffaire, AffaireUtil ;

Type
  TOF_AFEACTIVITEMODIF = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument(S : String )  ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    procedure bSelectAff1Click(Sender: TObject);     override ;
    procedure bEffaceAff1Click(Sender: TObject);     override ;
    procedure InitLibelleAffaire ;
    public
  end ;

Function AFLanceFiche_EActiviteModif(Argument:string):variant;
Implementation

procedure TOF_AFEACTIVITEMODIF.OnArgument (S : String ) ;
var
    Aff0, Aff1, Aff2, Aff3, Aff4 : string;
    Critere, Champ, valeur  : String;
    i,x : integer;
begin
     Aff0:=''; Aff1:=''; Aff2:=''; Aff3:=''; Aff4:='';

     Critere:=(Trim(ReadTokenSt(S)));
     while (Critere <>'') do
     begin
          if Critere<>'' then
          begin
               X:=pos(':',Critere);
               if x<>0 then
               begin
                    Champ:=copy(Critere,1,X-1);
                    Valeur:=Copy(Critere,X+1,length(Critere)-X);
               end;
               if Champ = 'DATE' then SetControlText('EAC_DATEACTIVITE', Valeur);
               if Champ = 'RESS' then SetControlText('EAC_RESSOURCE', Valeur);
               if Champ = 'AFF' then SetControlText('EAC_AFFAIRE', Valeur);
               if Champ = 'ART' then SetControlText('EAC_CODEARTICLE', Valeur);
               if Champ = 'TIE' then SetControlText('EAC_TIERS', Valeur);
               if Champ = 'AFF0' then Aff0:= Valeur;
               if Champ = 'AFF1' then Aff1:= Valeur;
               if Champ = 'AFF2' then Aff2:= Valeur;
               if Champ = 'AFF3' then Aff3:= Valeur;
               if Champ = 'AFF4' then Aff4:= Valeur;
          end;
          Critere:=(Trim(ReadTokenSt(S)));
    end;

    inherited ;

    if GetControlText('EAC_AFFAIRE')='' then
    begin
         SetControlText('EAC_AFFAIRE', Aff0+Aff1+Aff2+Aff3+Aff4);
         SetControlText('EAC_AFFAIRE0', Aff0);
         SetControlText('EAC_AFFAIRE1', Aff1);
         SetControlText('EAC_AFFAIRE2', Aff2);
         SetControlText('EAC_AFFAIRE3', Aff3);
         SetControlText('EAC_AVENANT', Aff4);
    end;

    for i:=1 to 10 do
        nextcontrol(TForm(Ecran), true);

    InitLibelleAffaire;
    THEdit(GetControl('EAC_DATEACTIVITE')).SetFocus;
end;

procedure TOF_AFEACTIVITEMODIF.bSelectAff1Click(Sender: TObject);
begin
     inherited;
     InitLibelleAffaire;
end;

procedure TOF_AFEACTIVITEMODIF.bEffaceAff1Click(Sender: TObject);
begin
     inherited;
     InitLibelleAffaire;
end;

procedure TOF_AFEACTIVITEMODIF.InitLibelleAffaire ;
begin
     THEdit(GetControl('EAC_TIERS')).Enabled := TRUE;

     // Met à jour le libelle du tiers
     THEdit(GetControl('EAC_TIERS')).SetFocus;
     nextprevcontrol(TForm(Ecran), true);
     SetControlVisible('EAC_AFFAIRE',true);
     THEdit(GetControl('EAC_AFFAIRE')).SetFocus;
     nextprevcontrol(TForm(Ecran), true);
     SetControlVisible('EAC_AFFAIRE',False);
     THEdit(GetControl('EAC_AFFAIRE1')).SetFocus;

     THEdit(GetControl('EAC_TIERS')).Enabled := FALSE;
end ;

procedure TOF_AFEACTIVITEMODIF.OnClose ;
begin
     inherited ;
end ;

procedure TOF_AFEACTIVITEMODIF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('EAC_AFFAIRE'));  Aff0:=THEdit(GetControl('EAC_AFFAIRE0'));
Aff1:=THEdit(GetControl('EAC_AFFAIRE1')); Aff2:=THEdit(GetControl('EAC_AFFAIRE2'));
Aff3:=THEdit(GetControl('EAC_AFFAIRE3')); Aff4:=THEdit(GetControl('EAC_AVENANT'));
Tiers:=THEdit(GetControl('EAC_TIERS'));
end;


Function AFLanceFiche_EActiviteModif(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFEACTIVITEMODIF','','',Argument);
end;

Initialization
  registerclasses ( [ TOF_AFEACTIVITEMODIF ] ) ;
end.
