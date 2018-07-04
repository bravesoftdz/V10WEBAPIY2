{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 16/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFARTICLEREVALMUL ()
Mots clefs ... : TOF;AFARTICLEREVALMUL
*****************************************************************}
Unit UtofAFARTICLEREVALMUL ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     FE_Main, Mul, db, dbTables,HDB,
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls, HTB97, HStatus,
     UTOB, UtilMulTrt,paramsoc, entgc,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uafo_ressource,AfUtilArticle,
     ActiviteUtil,
     //FactUtil, UtilArticle,
     TraducAffaire,DicoAF,
     UtofAfChoixDatesReval,
     UTOFAFTRADUCCHAMPLIBRE,
     UtilTaches
     //UtilRessource
     ;

Type
  TOF_AFARTICLEREVALMUL = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure BOuvrirOnClick(Sender: TObject);
    private
    fAppel : string;    
    public
    function RevaloriseArticles(TobArtSelec:TOB):integer;
  end ;

Procedure AFLanceFiche_RevalActivArticle;
Procedure AFLanceFicheAFPlanningRevalArticle;

Implementation

procedure TOF_AFARTICLEREVALMUL.BOuvrirOnClick(Sender: TObject);
var
TobArt:TOB;
iNbArticle,RepRevalo:integer;
begin
RepRevalo:=0;

TobArt:=TOB.Create('',Nil,-1);
try
SourisSablier;
//TraiteEnregMulTable (TFMul(ecran),'SELECT * FROM ARTICLE','ARS_RESSOURCE','ARTICLE', 'ARS_RESSOURCE','RESSOURCE',TobArt, True);
TraiteEnregMulListe (TFMul(Ecran), 'GA_ARTICLE','ARTICLE', TobArt, True);

iNbArticle:=TobArt.Detail.count;

// Revalorisation de tous les articles
if (iNbArticle<>0) then
    RepRevalo := RevaloriseArticles(TobArt)
else
    PGIInfoAf('Vous devez sélectionner au-moins un article', Ecran.Caption);

if (RepRevalo=-1) then
   PgiInfoAf('La revalorisation des articles a été annulée.', Ecran.Caption)
else
if (RepRevalo=0) then
   PgiInfoAf('La revalorisation des articles selectionnés a été effectuée avec succès.', Ecran.Caption)
else
   PgiInfoAf('Un problème est survenu lors de la revalorisation.', Ecran.Caption);

// Rafraichit la liste
TFMul(Ecran).ChercheClick;

finally
TobArt.Free;
SourisNormale;
end;
end;


function TOF_AFARTICLEREVALMUL.RevaloriseArticles (TobArtSelec:TOB) : integer;
  var
  x, i, j, iNbArticles, iNbInitMove : integer;
  RepDates, critere, Champ, Valeur : string;
  DateDebut, DateFin, DateDebPlan, DateFinPlan : TDateTime;
  bValoPR, bValoPV, bValoPLanning : boolean;
  TOBArticles, TOBActivite, vTobPlanning, TOBAffaires : TOB;
  AFOAssistants : TAFO_Ressources;
  QQ : TQuery;
  sReq, sArtCourant, sReqDeBase : string;
  begin
  Result := -1;
  if fAppel = '' then fAppel := 'ART';
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
  While (Critere <> '') do
    begin
      x := pos(':', Critere);
      if (x <> 0) then
        begin
          Champ := copy (Critere, 1, x-1);
          Valeur := Copy (Critere, x+1, length (Critere) - x);
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

    // Lancement de la revalorisation sur les articles
    iNbArticles := TobArtSelec.Detail.Count;
    TOBArticles := TOB.Create ('Les Articles', Nil, -1) ;
    AFOAssistants := TAFO_Ressources.Create;
    TOBAffaires := TOB.Create('Les Affaires', Nil, -1);
    try
      iNbInitMove := 0;
      if bValoPV then Inc(iNbInitMove);
      if bValoPR then Inc(iNbInitMove);
      if bValoPlanning then Inc(iNbInitMove);
      initMove (iNbArticles * iNbInitMove, 'Revalorisation');
      for i := 0 to iNbArticles - 1 do
        begin
        sArtCourant := TobArtSelec.Detail[i].GetValue ('GA_ARTICLE');
        // PL le 05/05/03 : on rajoute ACT_NUMLIGNEUNIQUE partie de la nouvelle clé de la ligne ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_NUMLIGNEUNIQUE
        sReqDeBase := 'SELECT  ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_FOLIO,ACT_TYPEARTICLE'
              + ',ACT_NUMLIGNE,ACT_NUMLIGNEUNIQUE,ACT_ARTICLE,ACT_QTE,ACT_ACTIVITEREPRIS,ACT_NUMPIECEACH'
              + ',ACT_ETATVISAFAC,ACT_TYPEHEURE,ACT_CODEARTICLE,ACT_UNITE,ACT_DEVISE'
              + ',ACT_PUPRCHARGE,ACT_PUPR,ACT_PUPRCHINDIRECT,ACT_TOTPRCHARGE,ACT_TOTPR,ACT_TOTPRCHINDI'
              + ',ACT_PUVENTE,ACT_PUVENTEDEV,ACT_TOTVENTE,ACT_TOTVENTEDEV'
              + ' FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" AND ACT_ARTICLE="' + sArtCourant;

        // PL le 31/01/03 : on réduit la sélection au nombre de champs nécessaires à la revalorisation
        //    sReq := 'SELECT * FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" AND ACT_ARTICLE="'+sArtCourant;
        if bValoPV then
        begin
          sReq := sReqDeBase;
          //////////////////////////////

          if (DateDebut <> 0) then
            sReq := sReq + '" AND ACT_DATEACTIVITE>="' + UsDateTime (DateDebut);
          if (DateFin<>0) then
            sReq := sReq + '" AND ACT_DATEACTIVITE<="' + UsDateTime (DateFin);

          sReq := sReq + '" AND ACT_QTE<>0 AND ACT_ACTIVITEREPRIS<>"FAC"';
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

        if bValoPR then  // PL le 05/03/03 : Pas de condition sur le ACT_ACTIVITEREPRIS ni sur ACT_ETATVISAFAC
          begin
            // PL le 31/01/03 : on réduit la sélection au nombre de champs nécessaires à la revalorisation
            // Il faut au minimum la clé plus les autres champs
            // sReq := 'SELECT * FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" AND ACT_ARTICLE="' + sArtCourant;
            sReq := sReqDeBase;
            //////////////////////////////

            if (DateDebut <> 0) then
              sReq := sReq + '" AND ACT_DATEACTIVITE>="' + UsDateTime (DateDebut);
            if (DateFin <> 0) then
              sReq := sReq + '" AND ACT_DATEACTIVITE<="' + UsDateTime (DateFin);

            sReq := sReq + '" AND ACT_QTE<>0 AND ACT_NUMPIECEACH=""';
//            if (GetParamSoc ('SO_AFAPPPOINT') = true) then
//              sReq := sReq + ' AND (ACT_ACTIVITEREPRIS="FAC" OR ACT_ETATVISAFAC="VIS")'
//            else
//              sReq := sReq + ' AND ACT_ACTIVITEREPRIS="FAC"';

            QQ := nil;
            TOBActivite := TOB.Create ('les activites', Nil, -1);
            try
              QQ := OpenSQL (sReq,True);

              if Not QQ.EOF then
                begin
                  TobActivite.LoadDetailDB ('ACTIVITE', '', '', QQ, False);
                  // PL le 30/01/03 : on remet les flags à false pour que l'agl puisse faire la différence
                  // des champs vraiment modifiés lors du update
                  TobActivite.SetAllModifie (false);

                  ValoriseLesActivites (TobActivite, TOBArticles, AFOAssistants, TOBAffaires, true, false);
                end;


            Finally
              Ferme (QQ);
              TOBActivite.Free;
            end;
         MoveCur(False);
         end;

        if bValoPlanning then
        begin
          sReq := ' SELECT * FROM AFPLANNING ';
          sReq := sReq + ' WHERE APL_ARTICLE = "' + sArtCourant+'"';
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
                Valorisation(vTobPlanning.detail[j], 'APL', AFOAssistants, TOBAffaires, TOBArticles);
              vTobPlanning.UpdateDB(true);
            end;
          finally
            vTobPlanning.Free;
            Ferme (QQ);
          end;
        end;
        MoveCur (False);
      end;

    finally
      TOBArticles.Free;
      AFOAssistants.Free;
      TOBAffaires.Free;
      FiniMove ();
    end;
// Fin modif PL le 05/03/03

  except
    Result:=1;
  end;

end;

procedure TOF_AFARTICLEREVALMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFARTICLEREVALMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFARTICLEREVALMUL.OnUpdate ;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end ;

procedure TOF_AFARTICLEREVALMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFARTICLEREVALMUL.OnArgument (S : String ) ;
var    ComboTypeArticle : THValComboBox;
begin
  Inherited ;
TToolBarButton97(GetControl('BOuvrir')).OnClick:= BOuvrirOnClick;

// permet de renseigner la condition plus sur la tablette des
// typearticle en fct de natures gérées dans l'appli (Btp, Affaire,
ComboTypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle;
fAppel := S;
{$IFDEF CCS3}
if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
{$ENDIF}
end ;

procedure TOF_AFARTICLEREVALMUL.OnClose ;
begin
  Inherited ;
end ;

Procedure AFLanceFiche_RevalActivArticle;
begin
AGLLanceFiche('AFF','AFARTICLEREVALMUL','','','ART');
end;

// Appel depuis Planning Traitement Revalorisation par Article
Procedure AFLanceFicheAFPlanningRevalArticle;
begin
AGLLanceFiche('AFF','AFARTICLEREVALMUL','','','PLANNING');
end;

Initialization
  registerclasses ( [ TOF_AFARTICLEREVALMUL ] ) ; 
end.
