{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 18/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : REVAL_RESS_MUL ()
Mots clefs ... : TOF;REVAL_RESS_MUL
*****************************************************************}
Unit UTOFREVAL_RESS_MUL ;

Interface
Uses
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     FE_Main, Mul, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,
{$ENDIF}
    StdCtrls, Controls, Classes, forms, sysutils,  ComCtrls, HTB97, HStatus,
    UTOB, UtilMulTrt,paramsoc, entgc,
    HCtrls, HEnt1, HMsgBox, UTOF, uafo_ressource,
    ActiviteUtil,
    FactUtil, UtilArticle,TraducAffaire,Dicobtp,
    UtofAfChoixDatesReval,UTOFAFTRADUCCHAMPLIBRE, UtilRessource,UtilTaches;

Type
  TOF_REVAL_RESS_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                  ; override ;
    procedure BOuvrirOnClick(Sender: TObject);
    private
    fAppel : string;    
    public
    du:integer;
    function RevaloriseRessources(TobResSelec:TOB):integer;

  end ;
Procedure AFLanceFiche_RevalActiv;
Procedure AFLanceFicheAFPlanningRevalRess;
Implementation


procedure TOF_REVAL_RESS_MUL.Onupdate;
begin
Inherited;
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end;


procedure TOF_REVAL_RESS_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolBarButton97(GetControl('BOuvrir')).OnClick:= BOuvrirOnClick;
  fAppel := S;
  {$IFDEF CCS3}
  if (getcontrol('PCOMP')  <> Nil) then  SetControlVisible('PCOMP', False);
  {$ENDIF}
end ;

procedure TOF_REVAL_RESS_MUL.OnClose ;
begin
  Inherited ;
end ;


procedure TOF_REVAL_RESS_MUL.BOuvrirOnClick (Sender : TObject);
  var
  TobRes : TOB;
  iNbRessource, RepRevalo : integer;
  begin
    RepRevalo := 0;
    TobRes := TOB.Create ('', Nil, -1);
    try
      SourisSablier;
      //  TraiteEnregMulTable (TFMul(ecran),'SELECT * FROM RESSOURCE','ARS_RESSOURCE','RESSOURCE', 'ARS_RESSOURCE','RESSOURCE',TobRes, True);
      // PL le 31/01/03 : on réduit aux champs strictement nécessaires
      TraiteEnregMulTable (TFMul(ecran), 'SELECT ARS_RESSOURCE FROM RESSOURCE', 'ARS_RESSOURCE', 'RESSOURCE',
                            'ARS_RESSOURCE', 'RESSOURCE', TobRes, True);

      iNbRessource := TobRes.Detail.count;

      // Revalorisation de toutes les ressources
      if (iNbRessource <> 0) then
        RepRevalo := RevaloriseRessources (TobRes)
      else
        PGIInfoAf ('Vous devez sélectionner au-moins une ressource', Ecran.Caption);

      if (RepRevalo = -1) then
        PgiInfoAf ('La revalorisation des ressources a été annulée.', Ecran.Caption)
      else
      if (RepRevalo = 0) then
        PgiInfoAf ('La revalorisation des ressources selectionnées a été effectuée avec succès.', Ecran.Caption)
      else
        PgiInfoAf ('Un problème est survenu lors de la revalorisation.', Ecran.Caption);

      // Rafraichit la liste
      TFMul(Ecran).ChercheClick;

    finally
      TobRes.Free;
      SourisNormale;
    end;
  end;

function TOF_REVAL_RESS_MUL.RevaloriseRessources (TobResSelec : TOB) : integer;
  var
  x, i, j, iNbRessources, iNbInitMove : integer;
  RepDates, critere, Champ, Valeur : string;
  DateDebut, DateFin, DateDebPlan, DateFinPlan : TDateTime;
  bValoPR, bValoPV, bValoPLanning : boolean;
  TOBAffaires, TOBArticles, TOBActivite,vTobPlanning : TOB;
  AFOAssistants : TAFO_Ressources;
  QQ : TQuery;
  sReq, sRessCourante, sReqDeBase : string;
  begin
    Result := -1;
    if fAppel = '' then fAppel := 'RES';
    // Recuperation des dates intervalles et de la décision de revalorisation sur PR et/ou PV
    // Appel depuis Revalorisation du planning  fAppel=PLANNING
    RepDates := AFLanceFiche_ChoixDatesRevalActiv (fAppel);
    if (RepDates = '') then exit;

    Result := 0;
    DateDebut := 0;
    DateFin := 0;
    bValoPR := false;
    bValoPV := false;
    DateDebPlan := 0;
    DateFinPlan := 0;
    bValoPLanning  := false;
    Critere := (Trim (ReadTokenSt (RepDates)));
    while (Critere <> '') do
      begin
        x := pos (':', Critere);
        if (x<>0) then
          begin
            Champ := copy (Critere, 1, x-1);
            Valeur := Copy (Critere, x+1, length(Critere) - x);
          end;

        if (Champ = 'DATEDEBUT') and (Valeur <> '') and IsValidDate(Valeur) then
          DateDebut := strtodate (Valeur);

        if (Champ = 'DATEFIN') and (Valeur <> '') and IsValidDate(Valeur) then
          DateFin := strtodate (Valeur);

        if (Champ = 'PR') and (Valeur = 'OUI') then
          bValoPR := true;

        if (Champ = 'PV') and (Valeur = 'OUI') then
          bValoPV := true;

        if (Champ = 'PL') and (Valeur = 'OUI') then
          bValoPLanning := true;

        if (Champ = 'DATEDEBPL') and (Valeur <> '') and IsValidDate(Valeur) then
          DateDebPlan := strtodate (Valeur);

        if (Champ = 'DATEFINPL') and (Valeur <> '') and IsValidDate(Valeur) then
          DateFinPlan := strtodate (Valeur);

        Critere := (Trim (ReadTokenSt (RepDates)));
      end;

    try
// PL le 05/03/03 : on revoie l'ensemble de la revalorisation en découpant en PR et PV
// pour gérer en PV les NumPieceAch = "", ce qui n'était pas le cas jusqu'à maintenant
      // Lancement de la revalorisation sur les ressources
      iNbRessources := TobResSelec.Detail.Count;
      TOBArticles := TOB.Create ('Les Articles', Nil, -1);
      TOBAffaires := TOB.Create ('Les Affaires', Nil, -1);
      AFOAssistants := TAFO_Ressources.Create;
      try
        iNbInitMove := 0;
        if bValoPV then Inc(iNbInitMove);
        if bValoPR then Inc(iNbInitMove);
        if bValoPlanning then Inc(iNbInitMove);
        initMove (iNbRessources * iNbInitMove, 'Revalorisation');
        for i := 0 to iNbRessources - 1 do
          begin
            sRessCourante := TobResSelec.Detail[i].GetValue ('ARS_RESSOURCE');
            // PL le 05/05/03 : on rajoute ACT_NUMLIGNEUNIQUE partie de la nouvelle clé de la ligne ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_NUMLIGNEUNIQUE
            sReqDeBase := 'SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_FOLIO,ACT_TYPEARTICLE'
                    + ',ACT_NUMLIGNE,ACT_NUMLIGNEUNIQUE,ACT_ARTICLE,ACT_QTE,ACT_ACTIVITEREPRIS,ACT_NUMPIECEACH'
                    + ',ACT_ETATVISAFAC,ACT_TYPEHEURE,ACT_CODEARTICLE,ACT_UNITE,ACT_DEVISE'
                    + ',ACT_PUPRCHARGE,ACT_PUPR,ACT_PUPRCHINDIRECT,ACT_TOTPRCHARGE,ACT_TOTPR,ACT_TOTPRCHINDI'
                    + ',ACT_PUVENTE,ACT_PUVENTEDEV,ACT_TOTVENTE,ACT_TOTVENTEDEV'
                    + ' FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" AND ACT_RESSOURCE="'+sRessCourante;

            // PL le 31/01/03 : on réduit la sélection au nombre de champs nécessaires à la revalorisation
            // Il faut au minimum la clé plus les autres champs
            // SELECT * : le nombre d'enreg est relativement restreint (une ressource, intervalle de dates...)
            //sReq := 'SELECT * FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" AND ACT_RESSOURCE="' + sRessCourante;
            if bValoPV then // PL le 05/03/03 : Pas de condition sur le ACT_NumPieceAch
            begin
              sReq := sReqDeBase;
              ////////////////////////////

              if (DateDebut <> 0) then
                sReq := sReq + '" AND ACT_DATEACTIVITE>="' + UsDateTime (DateDebut);
              if (DateFin<>0) then
                sReq := sReq + '" AND ACT_DATEACTIVITE<="' + UsDateTime (DateFin);
              // PL le 16/09/02 : on bride à nouveau l'accès aux prestation uniquement car on a maintenant la possibilité de revaloriser par article
              sReq := sReq + '" AND ACT_TYPEARTICLE="PRE" AND ACT_QTE<>0 AND ACT_ACTIVITEREPRIS<>"FAC"';
              //    sReq := sReq + '" AND ACT_QTE<>0 AND ACT_ACTIVITEREPRIS<>"FAC"';
              if (GetParamSoc ('SO_AFAPPPOINT') = true) then
                sReq := sReq + ' AND ACT_ETATVISAFAC<>"VIS"';

              TOBActivite := TOB.Create ('les activites', Nil, -1);

              QQ := nil;
              try
                QQ := OpenSQL (sReq, True);
                if Not QQ.EOF then
                begin
                  TobActivite.LoadDetailDB ('ACTIVITE', '', '', QQ, False);
                  // PL le 30/01/03 : on remet les flags à false pour que l'agl puisse faire la différence
                  // des champs vraiment modifiés lors du update
                  TobActivite.SetAllModifie (false);

                  ValoriseLesActivites (TobActivite, TOBArticles, AFOAssistants, TOBAffaires, false, true);
                end;
              finally
                Ferme (QQ);
                TOBActivite.Free;
              end;
            MoveCur(False);
            end;

            //
            // On traite ensuite pour le PR
            //
            if bValoPR then // PL le 05/03/03 : Pas de condition sur le ACT_ACTIVITEREPRIS ni sur ACT_ETATVISAFAC
              begin
                // PL le 31/01/03 : on réduit la sélection au nombre de champs nécessaires à la revalorisation
                // Il faut au minimum la clé plus les autres champs
                // SELECT * : le nombre d'enreg est relativement restreint (une ressource, intervalle de dates...)
                //sReq := 'SELECT * FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" AND ACT_RESSOURCE="'+sRessCourante;
                sReq := sReqDeBase;
                ////////////////////////////////

                if (DateDebut <> 0) then
                  sReq := sReq + '" AND ACT_DATEACTIVITE>="' + UsDateTime(DateDebut);
                if (DateFin <> 0) then
                  sReq := sReq + '" AND ACT_DATEACTIVITE<="' + UsDateTime(DateFin);

                // PL le 16/09/02 : on bride à nouveau l'accès aux prestation uniquement car on a maintenant la possibilité de revaloriser par article
                sReq := sReq + '" AND ACT_TYPEARTICLE="PRE" AND ACT_QTE<>0 AND ACT_NUMPIECEACH=""';

//                if (GetParamSoc ('SO_AFAPPPOINT') = true) then
//                  sReq := sReq + ' AND ACT_ETATVISAFAC="VIS"';
//                  sReq := sReq + ' AND (ACT_ACTIVITEREPRIS="FAC" OR ACT_ETATVISAFAC="VIS")'
//                else
//                  sReq := sReq + ' AND ACT_ACTIVITEREPRIS="FAC"'; PL le 05/03/03

                TOBActivite := TOB.Create ('les activites', Nil, -1);
                QQ := nil;
                try
                  QQ := OpenSQL (sReq, True);

                  if Not QQ.EOF then
                    begin
                      TobActivite.LoadDetailDB ('ACTIVITE', '', '', QQ, False);
                      // PL le 30/01/03 : on remet les flags à false pour que l'agl puisse faire la différence
                      // des champs vraiment modifiés lors du update
                      TobActivite.SetAllModifie (false);

                      ValoriseLesActivites (TobActivite, TOBArticles, AFOAssistants, TOBAffaires, true, false);
                    end;

                finally
                  Ferme (QQ);
                  TOBActivite.Free;
                end;
              MoveCur(False);
              end;

            if bValoPlanning then
            begin
              sReq := ' SELECT * FROM AFPLANNING ';
              sReq := sReq + ' WHERE APL_RESSOURCE = "'+sRessCourante+'"';
              if (DateDebPlan <> 0) then
              sReq := sReq + ' AND APL_DATEDEBPLA >= "' + UsDateTime(DateDebPlan)+'"';
              if (DateFinPlan <> 0) and (DateFinPlan > DateDebPlan) then
              sReq := sReq + ' AND APL_DATEFINPLA <= "' + UsDateTime(DateFinPlan)+'"';
              vTobPlanning := Tob.create('Les plannings', nil, -1);
              try
                QQ := OpenSQL (sReq, True);
                if Not QQ.EOF then
                begin
                  vTobPlanning.LoadDetailDB('AFPLANNING','','', QQ,True);
                  For j:=0 to vTobPlanning.detail.count-1 do
//Modif FV            Valorisation(vTobPlanning.detail[j], 'APL', AFOAssistants, TOBAffaires, TOBArticles);
                  vTobPlanning.UpdateDB(true);
                end;
              finally
                vTobPlanning.Free;
                Ferme (QQ);
              end;
            MoveCur(False);
            end;

          end;

finally
TOBArticles.Free;
AFOAssistants.Free;
TOBAffaires.Free;
FiniMove();
end;
// Fin modif PL le 05/03/03
////////////////////////////////////////////

except
Result:=1;
end;
end;

Procedure AFLanceFiche_RevalActiv;
begin
AGLLanceFiche('AFF','REVAL_RESS_MUL','','','RES');
end;

Procedure AFLanceFicheAFPlanningRevalRess;
begin
AGLLanceFiche('AFF','REVAL_RESS_MUL','','','PLANNING');
end;

Initialization
  registerclasses ( [ TOF_REVAL_RESS_MUL ] ) ;
end.
