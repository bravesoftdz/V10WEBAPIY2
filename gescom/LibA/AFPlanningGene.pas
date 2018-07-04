{***********UNITE*************************************************
Auteur  ...... : C.B
Créé le ...... : 04/10/2002
Modifié le ... :   /  /
Description .. : Generation du planning récurrent
Mots clefs ... :
*****************************************************************}
unit AFPlanningGene;


interface

  uses classes, hctrls, utob, UAFO_Ressource, UAFO_REGLES, HEnt1, ed_tools,
       DicoAF, Sysutils,  
{$IFDEF EAGLCLIENT}

{$ELSE}                   
   dbtables,
{$ENDIF}

  afPlanningCst, utilplanning, utilTaches, Paramsoc;


  Type
    TAFGenerateur = class
      public
        function GenererPlanning(pStAffaire : String; pTache : RecordTache; pInModeGene : Integer; pBoValorisation : Boolean) : Boolean;

      private
        fStAffaire      : String;
        fTache          : RecordTache;
        fInModeGene     : Integer;
        fBoValorisation : Boolean;
        fListeRes       : TAFO_Ressources;

        procedure ExecuterGeneration;
  end;

const
	TexteMsgTache: array[1..2] of string 	= (
          {1}        '',
          {2}        'La génération du planning ne s''est pas effectuée correctement.');

implementation

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/10/2002
Modifié le ... :
Description .. : generation pour une affaire ou pour une tache
Mots clefs ... :
*****************************************************************}
procedure TAFGenerateur.ExecuterGeneration;
var
  i,j             : Integer;
  vTob            : Tob;
  vQr             : TQuery;
  vInNumLigne     : Integer;
  vTobMere        : Tob;
  vAFReglesTache  : TAFRegles;
  vSt             : String;
  vArticle        : RecordArticle;

begin

  vTobMere := TOB.Create('Planning mere', nil, -1);

  try
    try
      if fTache.StNumeroTache <> '' then
        begin

          // chargement de la liste des ressources
          fListeRes := TAFO_Ressources.Create;
          vAFReglesTache := TAFRegles.create;
          InitMoveProgressForm (nil,'Traitement en cours...','', 50, false, true) ;
 
          try
            vAFReglesTache.LoadDBReglesTaches(fStAffaire, fTache.StNumeroTache);
            fListeRes.RessourceD1TacheAffaire(fStAffaire, fTache.StNumeroTache);
            vInNumLigne := GetNumLignePlanning(fStAffaire);
            if (fInModeGene = cInRemplace) or (fInModeGene = cInSupprimer) then
              ExecuteSql('DELETE FROM AFPLANNING WHERE APL_AFFAIRE = "' + fStAffaire + '" AND APL_NUMEROTACHE = ' + fTache.StNumeroTache + ' AND APL_LIGNEGENEREE = "X" AND APL_ACTIVITEGENERE <> "X"' )
            else if (fInModeGene = cInRemplaceTout) or (fInModeGene = cInSupprimerTout) then
              ExecuteSql('DELETE FROM AFPLANNING WHERE APL_AFFAIRE = "' + fStAffaire + '" AND APL_NUMEROTACHE = ' + fTache.StNumeroTache + ' AND APL_ACTIVITEGENERE <> "X"');

            // suppression du planning sans regenerer
            if (fInModeGene <> cInSupprimer) and (fInModeGene <> cInSupprimerTout) then
            begin
              // les données article sont les mêmes pour toutes les ressources
              GetArticle(fTache.StArticle, vArticle);
              for i := 0 to fListeRes.Count - 1 do
                begin
                  fListeRes.GetRessource(i).GenererPlanning(vInNumLigne, vTobMere, vAFReglesTache, fStAffaire, fTache, vArticle, fBoValorisation, fInModeGene = cInAjout);
                end;
            end;
            vTobMere.InsertDB(nil, false);
          finally
            fListeRes.Free;
            vAFReglesTache.Free;
            FiniMoveProgressForm;
          end;
        end
      else
        begin
          // recherche des numero de taches de cette affaire
          vSt :='SELECT ATA_NUMEROTACHE, ATA_LIBELLETACHE1, ATA_FONCTION, ';
          vSt := vSt + 'ATA_ARTICLE, ATA_UNITETEMPS, ATA_LASTDATEGENE, ATA_TIERS, ';
          vSt := vSt + 'ATA_QTEINITIALE FROM TACHE WHERE ATA_AFFAIRE = "' + fStAffaire + '"';
          vTob := tob.create('TACHE', nil,-1);
					InitMoveProgressForm (nil,'Traitement en cours...','', 50, false, true) ;

          try
            vQr := openSql(vSt, true);
            if not vQr.eof then
              begin
                vTob.LoadDetailDB('TACHE','','',vQR,False);
                vInNumLigne := GetNumLignePlanning(fStAffaire);
                if (fInModeGene = cInRemplace) or (fInModeGene = cInSupprimer) then
                  ExecuteSql('DELETE FROM AFPLANNING WHERE APL_AFFAIRE = "' + fStAffaire + '" AND APL_LIGNEGENEREE = "X"  AND APL_ACTIVITEGENERE <> "X"')
                else if (fInModeGene = cInRemplaceTout) or (fInModeGene = cInSupprimerTout) then
                  ExecuteSql('DELETE FROM AFPLANNING WHERE APL_AFFAIRE = "' + fStAffaire + '" AND APL_ACTIVITEGENERE <> "X"');

                // suppression du planning sans regenerer
                if (fInModeGene <> cInSupprimer) and  (fInModeGene <> cInSupprimerTout) then

                  // pour chaque tache
                  for i := 0 to vTob.detail.count - 1 do
                    begin
                      // chargement de la liste des ressources
                      // et chargement des regles de la tache
                      fListeRes := TAFO_Ressources.Create;
                      vAFReglesTache := TAFRegles.create;

                      try
                        fTache.StNumeroTache := vTob.detail[i].GetValue('ATA_NUMEROTACHE');
                        fTache.StLibTache := vTob.detail[i].GetValue('ATA_LIBELLETACHE1');
                        fTache.StFctTache := vTob.detail[i].GetValue('ATA_FONCTION');
                        fTache.StArticle := vTob.detail[i].GetValue('ATA_ARTICLE');
                        fTache.StTiers := vTob.detail[i].GetValue('ATA_TIERS');
                        fTache.StUnite := vTob.detail[i].GetValue('ATA_UNITETEMPS');
                        fTache.StLastDateGene := vTob.detail[i].GetValue('ATA_LASTDATEGENE');
                        fTache.StActiviteRepris := vTob.detail[i].GetValue('ATA_ACTIVITEREPRIS');
                        fTache.BoCompteur := ((vTob.detail[i].GetValue('ATA_QTEINITIALE') <> 0) and getParamsoc('SO_AFREALPLAN'));

                        fListeRes.RessourceD1TacheAffaire(fStAffaire, vTob.detail[i].GetValue('ATA_NUMEROTACHE'));
                        vAFReglesTache.LoadDBReglesTaches(fStAffaire, vTob.detail[i].GetValue('ATA_NUMEROTACHE'));
                        GetArticle(fTache.StArticle, vArticle);

                        for j := 0 to fListeRes.Count - 1 do
                          begin
                            fListeRes.GetRessource(j).GenererPlanning(vInNumLigne,
                                                                      vTobMere, vAFReglesTache,
                                                                      fStAffaire, fTache, vArticle,
                                                                      fBoValorisation,
                                                                      fInModeGene = cInAjout);
                          end
                      finally
                        fListeRes.Free;
                        vAFReglesTache.Free;
                      end;
                    end;
                vTobMere.InsertDB(nil, false);
              end;

          finally
            vtob.Free;
            ferme(vQr);
            FiniMoveProgressForm;
          end;
        end;
    except
      on e:exception do PGIBoxAF(TexteMsgTache[2]+#10#13+e.Message,'');
    end;
  finally
    vTobMere.Free;
  end;
end;

function TAFGenerateur.GenererPlanning(pStAffaire : String; pTache : RecordTache; pInModeGene : Integer; pBoValorisation : Boolean) : Boolean;
var
  io  : TIoErr;

begin

  fStAffaire      := pStAffaire;
  fTache          := pTache;
  fInModeGene     := pInModeGene;
  fBoValorisation := pBoValorisation;

  io:=Transactions(ExecuterGeneration, 1);
  if io<>oeOk then
    result := false
  else
    result := true;
end;

end.
