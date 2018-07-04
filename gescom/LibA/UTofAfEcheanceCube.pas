unit UTofAfEcheanceCube;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,paramsoc,
{$IFDEF EAGLCLIENT}
  Maineagl,
{$ELSE}
  dbTables, db, FE_Main,
{$ENDIF}
  HCtrls,HTB97,Cube, UTOF,UTofAfBaseCodeAffaire,AfUtilArticle,DicoAf, UtofAftableaubord;

Type

TOF_AfEcheanceCube = Class (TOF_AFBASECODEAFFAIRE)
 Public
  procedure OnArgument (stArgument : string); override;
  procedure NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);override ;
 Private
  BBureau : boolean  ;
  procedure AfterShow;
  END;

Procedure AFLanceFiche_EcheanceCube (Argument : string);

implementation

//******************************************************************************
//**************** Cube d'activité  ********************************************
//******************************************************************************

Procedure AFLanceFiche_EcheanceCube(Argument:string);
begin
  AGLLanceFiche ('AFF','AFECHEANCE_CUBE','','',Argument);
end;

procedure TOF_AfEcheanceCube.AfterShow;
begin
If Bbureau then  // si appel depuis le bureau
   begin
   TFCube(Ecran).BChercheClick(Self);   //force la recherche avec les paramètre passé lors de l'appel
   TFCube(Ecran).BOKClick(Self);  // force l'affichage du résultat, sans appuyer sur le bouton OK
   TfCube(Ecran).ForceSelect('','CEGID_BUREAU');
   end;
end;

procedure TOF_AfEcheanceCube.OnArgument(stArgument : String );
Var Critere ,champ, valeur: string;
Begin
Inherited;
BBureau :=false;
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        if pos(':',Critere)<>0 then
          begin
           Champ:=ReadTokenPipe(Critere,':');
           Valeur:=critere;
          end
        else
          begin
           if pos('=',Critere)<>0 then
             begin
              Champ:=ReadTokenPipe(Critere,'=');
              Valeur:= Critere;
             end else champ := critere;
          end ;
        if (Champ = 'BUREAU')   then
          begin
          TFCube(Ecran).OnAfterFormShow := AfterShow;
          BBureau :=true;
          SetControlVisible ('FCLIENT',false);
          SetControlEnabled('AFF_TIERS',false); SetControlEnabled('AFF_TIERS_',false);
          SetControlVisible ('AFF_AFFAIRE1',false);  SetControlVisible ('AFF_AFFAIRE2',false);
          SetControlVisible ('AFF_AFFAIRE3',false);  SetControlVisible ('AFF_AVENANT',false);
          SetControlVisible ('TAFA_AFFAIRE',false); SetControlVisible ('TAFA_AFFAIRE_',false);
          SetControlVisible ('BSELECTAFF1',false); SetControlVisible ('BEFFACEAFF1',false);
          SetControlVisible ('BSELECTAFF2',false); SetControlVisible ('BEFFACEAFF2',false);
          SetControlVisible ('AFF_AFFAIRE1_',false);  SetControlVisible ('AFF_AFFAIRE2_',false);
          SetControlVisible ('AFF_AFFAIRE3_',false); SetControlVisible ('AFF_AVENANT_',false);  SetControlVisible ('TAFF_AFFAIRE_',false);
          SetControlVisible ('TAFF_LIBREAFF1',false); SetControlVisible ('AFF_LIBREAFF1',false);
          SetControlVisible ('TAFF_LIBREAFF2',false); SetControlVisible ('AFF_LIBREAFF2',false);
          SetControlVisible ('TAFF_LIBREAFF3',false); SetControlVisible ('AFF_LIBREAFF3',false);
          SetControlVisible ('TAFF_LIBREAFF4',false); SetControlVisible ('AFF_LIBREAFF4',false);
          SetControlVisible ('TAFF_LIBREAFF5',false); SetControlVisible ('AFF_LIBREAFF5',false);
          SetControlVisible ('TAFF_LIBREAFF6',false);  SetControlVisible ('AFF_LIBREAFF6',false);
          SetControlVisible ('TAFF_LIBREAFF7',false); SetControlVisible ('AFF_LIBREAFF7',false);
          SetControlVisible ('TAFF_LIBREAFF8',false); SetControlVisible ('AFF_LIBREAFF8',false);
          SetControlVisible ('TAFF_LIBREAFF9',false); SetControlVisible ('AFF_LIBREAFF9',false);
          SetControlVisible ('TAFF_LIBREAFFA',false); SetControlVisible ('AFF_LIBREAFFA',false);
          SetControlVisible ('TAFF_RESSOURCE1',false); SetControlVisible ('AFF_RESSOURCE1',false);
          SetControlVisible ('TAFF_RESSOURCE2',false); SetControlVisible ('AFF_RESSOURCE2',false);
          SetControlVisible ('TAFF_RESSOURCE3',false); SetControlVisible ('AFF_RESSOURCE3',false);
          SetControlVisible ('TAFF_DEPARTEMENT',false); SetControlVisible ('AFF_DEPARTEMENT',false);
          SetControlVisible ('TAFF_RESPONSABLE',false); SetControlVisible ('AFF_RESPONSABLE',false);
          SetControlVisible ('TAFF_ETATAFFAIRE',false);  SetControlVisible ('AFF_ETATAFFAIRE',false);
          SetControlVisible ('TAFF_GROUPECONF',false); SetControlVisible ('AFF_GROUPECONF',false);
          SetControlVisible ('TAFF_DEPARTEMENT',false);  SetControlVisible ('AFF_DEPARTEMENT',false);
          SetControlVisible ('TAFF_ETABLISSEMENT',false);  SetControlVisible ('AFF_ETABLISSEMENT',false);
          SetControlVisible ('BVEL4',false);   SetControlVisible ('BVEL5',false);
          SetControlVisible ('TCOMPLAFF',false);  SetControlVisible ('TLIBRESAFF',false);
          TfCube(Ecran).ParamDIm:=MpOndemand;
          end
        else if (Champ = 'DATEFIN') then SetCOntroLText('AFA_DATEECHE_',Valeur)
        else if (Champ = 'DATEDEB')then SetCOntroLText('AFA_DATEECHE',Valeur)
        else if (Champ = 'CLT')  then
         begin
         SetCOntroLText('AFF_TIERS',Valeur);
         SetCOntroLText('AFF_TIERS_',Valeur);
         end;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
End;


procedure TOF_AfEcheanceCube.NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('AFF_TIERS'));
  Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
  Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
  Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
  Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
  Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

Initialization
registerclasses([TOF_AfEcheanceCube]);
end.
