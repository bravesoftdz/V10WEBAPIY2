{***********UNITE*************************************************
Auteur  ...... : MCDesseignet
Créé le ...... : 13/05/2005
Modifié le ... :   /  /
Description .. : Generation des dates pour les obligation DP (à partir AfModelTAche
Mots clefs ... :
*****************************************************************}
unit DPObligGene;


interface

  uses
  utob,UAFO_REGLES,classes, hctrls, ed_tools,
{$IFDEF EAGLCLIENT}

{$ELSE}                   
   dbtables,
{$ENDIF}
  afPlanningCst
  ;

Procedure DpObligGeneration (Datedeb, datefin : Tdatetime; var  pTobMere :tob);

implementation

Procedure DpObligGeneration (Datedeb, datefin : Tdatetime;var  pTobMere :tob);
var
  i, jj           : Integer;
  vQr             : TQuery;
  vTob,vTobMere   : Tob;
  vSt             : String;
  vAFRegles       : TAFRegles;
  vListeJours     : TStringList;
  vRdDuree        : Double; // durée des interventions
  vInNbJoursDecal : Integer; // nombre de jours de decalage maximum
begin
  vRdDuree  := 1;
  vInNbJoursDecal := 0;
  vAFRegles := TAFRegles.create;
  pTobMere := TOB.Create('obligation mere', nil, -1);
      //on charge les taches modeles avec les champs à mettre dans la TOB
  vSt :='SELECT AFM_MODELETACHE, AFM_LIBELLETACHE1,AFM_LIBELLETACHE2,AFM_CHARLIBRE3,';
  vSt := vSt + ' AFM_DESCRIPTIF,AFM_FAMILLETACHE FROM AFMODELETACHE WHERE AFM_MODELETACHE like "Z%"';
  vTob := tob.create('liste obli', nil,-1);
  try
   vQr := openSql(vSt, true);
   if not vQr.eof then
   begin
    vTob.LoadDetailDB('AFMODELETACHE','','',vQR,False);
      // pour chaque tache  modele
    for i := 0 to vTob.detail.count - 1 do
    begin
      // chargement des regles de la tache
      vAFRegles := TAFRegles.create;
     try
     vAFRegles.LoadDBReglesModele(Vtob.detail[i].getstring('AFM_MODELETACHE'),CinRemplace,DateDeb,DateFin);
        //calcul de la liste des jours
     vListeJours := vAFRegles.GenereListeJours(vRdDuree, vInNbJoursDecal, cInNoCompteur);
     if vListeJours <> nil then
        begin
        for jj := 0 to vListeJours.count - 1 do
          begin //alimentation TOB pour chaque jour calculé
           MoveCurProgressForm('');
           vtobmere:= tob.create ('',ptobmere,-1);
           vtobmere.AddChampSupValeur('AFM_MODELETACHE', Vtob.detail[i].Getstring ('AFM_MODELETACHE'));
           vtobmere.AddChampSupValeur('DATE',vlistejours[jj]);
          end;
        end;
     vListeJours.Free;
     finally
     vAFRegles.Free;
     end;
    end; //fin boucle sur tob des modèles
  end; // fin requête renseignée
  finally
    vtob.Free;
    ferme(vQr);
    FiniMoveProgressForm;
    end;
end;
end.
