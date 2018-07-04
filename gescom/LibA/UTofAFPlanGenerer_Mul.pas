unit UTofAFPlanGenerer_Mul;

interface

uses  StdCtrls,Controls,Classes,M3FP,HTB97,AFPlanning, AFPlanningCst, hqry,
{$IFDEF EAGLCLIENT}
      eMul,utob,Maineagl,
{$ELSE}
      db,dbTables,DBGrids,mul, FE_main,
{$ENDIF}
      forms,sysutils,HDB,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ConfidentAffaire,
      AffaireUtil,AffaireRegroupeUtil,DicoAF,SaisUtil,EntGC,
      utofAfBaseCodeAffaire,utilpgi,AglInit,UtilGc,TraducAffaire,
      UtilTaches, AFPlanningGene, utilPlanning, utofAfPlanningGenerer;

Type
     TOF_AFPLANGENERER_MUL = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument(stArgument : String ); override;
        procedure OnUpdate; override;
        procedure OnLoad; override;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        procedure TraiteRetour;
        procedure TraiteInvite;
        procedure BOuvrirOnClick(Sender : TObject);

     private
       fUserInvite : boolean;
{$IFDEF EAGLCLIENT}
       fListe      : THGrid;
{$ELSE}
       fListe      : THDBGrid;

{$ENDIF}
       fQ          : THQuery;

     End;

  //function AFLanceFicheAFPlanning(Lequel, Argument : String) : String;
  function AFLanceFicheAFPlanGenerer_Mul : String;

implementation           

Function AFLanceFicheAFPlanGenerer_Mul : String;
begin
  result := AGLLanceFiche('AFF','AFPLANGENERER_MUL','', '','');
end;
                                 
procedure TOF_AFPLANGENERER_MUL.OnArgument(stArgument : String);
Begin

  Inherited;

  TFMUL(ecran).DBListe := 'AFMULPLANNINGGENE';

  // on ne peut pas modifier les affaires
  SetControlVisible ('BInsert',False);
  SetControlVisible ('BInsert1',False);
  SetControlVisible('bAssistantCreation',False);

  // et les sous affaires
  if not(GereSousAffaire) then
    begin
      SetcontrolVisible('TAFF_AFFAIREREF',False);
      SetcontrolVisible('AFFAIREREF1',False); SetcontrolVisible('AFFAIREREF2',False);
      SetcontrolVisible('AFFAIREREF3',False); SetcontrolVisible('AFFAIREREF4',False);
      SetcontrolVisible('AFF_ISAFFAIREREF',False);
      SetcontrolVisible('BSELECTAFF2',False);
    end;

  {$IFDEF EAGLCLIENT}
    TraduitAFLibGridSt(TFMul(Ecran).FListe);
  {$ELSE}
    TraduitAFLibGridDB(TFMul(Ecran).FListe);
  {$ENDIF}

   TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick ;
   fListe := TFMul(Ecran).FListe;
   fQ := TFMul(Ecran).Q;
   {$IFDEF CCS3}
  if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  if (getcontrol('PZONE') <> Nil)  then SetControlVisible ('PZONE', False);
  {$ENDIF}
End;

procedure TOF_AFPLANGENERER_MUL.OnUpdate;
Begin

  // Gestion repositionnement auto sur l'affaire en cours si sortie rapide / bug par prg ( ou Eagl)
  inherited;

  if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('AFF_AFFAIRE2','');
  if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('AFF_AFFAIRE3','');

End;

procedure TOF_AFPLANGENERER_MUL.TraiteRetour;
Var stRetour  : string;
    Champ     : string;

begin

  stRetour := GetControlText ('RETOUR');
  if stRetour <> '' then
    begin
      // On récupére dans l'ordre le client + Affaire 1,2,3
      Champ:=(Trim(ReadTokenSt(stRetour))); if Champ <> '' then SetControltext('AFF_TIERS',Champ);
      Champ:=(Trim(ReadTokenSt(stRetour))); if Champ <> '' then SetControltext('AFF_AFFAIRE1',Champ);
      Champ:=(Trim(ReadTokenSt(stRetour))); if (Champ <> '') And (VH_GC.CleAffaire.Co2Visible) then SetControltext('AFF_AFFAIRE2',Champ);
      Champ:=(Trim(ReadTokenSt(stRetour))); if (Champ <> '') And (VH_GC.CleAffaire.Co3Visible) then SetControltext('AFF_AFFAIRE3',Champ);
    end;
  SetControlText('RETOUR', '');

end;

procedure TOF_AFPLANGENERER_MUL.OnLoad;
Var Affaire,Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant : string;

begin

  inherited;

  if fUserInvite and ( not V_PGI.Superviseur) and ( not V_PGI.Controleur) then
    TraiteInvite;

  if GereSousAffaire then
    begin
      if (GetControlText('AFFAIREREF1')='') then SetControlText('AFF_AFFAIREREF','')
      else
        begin
          Affaire0 := 'A';
          Affaire1 := GetControlText('AFFAIREREF1'); Affaire2 := GetControlText('AFFAIREREF2');
          Affaire3 := GetControlText('AFFAIREREF3'); Avenant := GetControlText('AFFAIREREF4');
          Affaire:=CodeAffaireRegroupe(Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant, taModif, false,false,false);
          if not ExisteAffaire(Affaire,'') then SetControlText('AFF_AFFAIREREF',Trim(Copy(Affaire,1,15)));
        end;
    end;

end;

procedure TOF_AFPLANGENERER_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1')); Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3')); Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('AFF_TIERS'));

  // affaire de référence pour recherche
  Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
  Aff1_:=THEdit(GetControl('AFFAIREREF1'));
  Aff2_:=THEdit(GetControl('AFFAIREREF2'));Aff3_:=THEdit(GetControl('AFFAIREREF3'));
  Aff4_:=THEdit(GetControl('AFFAIREREF4'));
end;

procedure TOF_AFPLANGENERER_MUL.TraiteInvite;
var                                 
lstzone,zone : string;

begin

  lstzone := 'AFF_RESPONSABLE;AFF_RESSOURCE1;AFF_RESSOURCE2;AFF_RESSOURCE3';
  while lstzone <> '' do
   begin
     zone := ReadTokenSt(Lstzone);
     if ThEdit(GetControl(zone))<> nil then
       begin
         SetControlText(zone,'');
         SetControlEnabled(zone,True);
       end;
   end;
  TToolBarButton97(GetControl('BInsert')).enabled := False;
  TToolBarButton97(GetControl('BInsert1')).enabled := False;
  TToolBarButton97(GetControl('BVOIRAFFAIRE')).enabled := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/10/2002
Modifié le ... :
Description .. : génération du planning pour un ensemble d'affaire
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANGENERER_MUL.BOuvrirOnClick(Sender : TObject);
var
  i           : Integer;
  vGenerateur : TAFGenerateur;
  vBoStop     : Boolean;
  vInRemplace : Integer;
  pTache      : RecordTache;
  vStRetour   : String;
  vBoValorise : Boolean;

begin
  pTache.StNumeroTache  := '';
  pTache.StLibTache     := '';
  pTache.StFctTache     := '';
  pTache.StArticle      := '';
  pTache.StUnite        := '';
  pTache.StLastDateGene := '';
  pTache.BoCompteur     := False;

  vBoStop := false;

  vStRetour   := AFLanceFicheAFPlanningGenerer('', 'SELECTION');
  vInRemplace := strToInt(ReadTokenSt(vStRetour));
  vBoValorise := ReadTokenSt(vStRetour) = 'TRUE';

  try
    if (fListe.nbSelected <> 0) then
      begin
        for i := 0 to fListe.nbSelected -1 do
          begin
            vGenerateur := TAFGenerateur.create;
            try                 
              fListe.GotoLeBookMark(i);
{              if TCheckBox(GetControl('CBREMPLACER')).checked then
                if TCheckBox(GetControl('CBSUPPRIMER')).checked then
                  vInRemplace := cInRemplaceTout
                else
                  vInRemplace := cInRemplace
              else if TCheckBox(GetControl('CBSUPPRIMERPLA')).checked then
                if TCheckBox(GetControl('CBSUPPRIMER')).checked then
                  vInRemplace := cInSupprimerTout
                else
                  vInRemplace := cInSupprimer
              else
                vInRemplace := cInAjout;
}
              if not EtatInterdit(vInRemplace, fQ.Findfield('AFF_AFFAIRE').AsString, pTache) then
              begin
                if not vGenerateur.GenererPlanning(fQ.Findfield('AFF_AFFAIRE').AsSTring, pTache, vInRemplace, vBoValorise) then
                begin
                  vBoStop := true;
                  break;
                end;
              end;
            finally
              vGenerateur.Free;
            end;
          end;
        if not vBoStop then PGIInfoAF ('La génération du planning est terminée', '');
      end;

  except
    on E:Exception do
      begin
        MessageAlerte('Erreur lors de la génération du planning' +#10#13#10#13 + E.message ) ;
      end; // on
  end; // try
end;

Initialization
  registerclasses([TOF_AFPLANGENERER_MUL]);
end.
