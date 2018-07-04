{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 19/06/2006
Modifié le ... :   /  /
Description .. : PGCALCULSIMULATION
                 Calcul de la simulation
Mots clefs ... :
*****************************************************************
PT1   : 12/04/2007 NA  Prise en compte des données de la table dynamique  pour la tables Droit
PT2   : 12/04/2007 NA  Calcul de l'indemnité du salarié
PT3   : 12/04/2007 NA  Mémorisation des tables IDR
PT4   : 12/04/2007 NA  Ajout méthode de calcul ABO
PT5   : 17/04/2007 NA  Chargement de la table CODTAB et SALAIREMOYEN et CODTABVALEUR pour la survie
                       dans une TOB pour améliorer la performance}

Unit PGCALCULSIMULATION ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
     Controls, 
     Classes,
     {$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     Utob,
     P5Util,
     Pgpopuloutils,
     math,
     PGTablesDyna, // pt1
     HEnt1;

   Type Tparamsimul = record
       Tdatesimul : string;
       Tcodesimul : string;
       TLibellesimul: string;
       Ttypedepart : string;
       Tmethcalcul : string;
       Ttauxactual : double;
       Tdatecalcul : string;
       Tcodecalcul : string;
end;
  procedure calculsimulation(tob_pop : TOB; paramsimul : TPARAMSIMUL);
  function coeffsurvie(sexe: string; ageret,agesal : integer; tablesur: string) : DOUBLE;
  //  function nbmoisacquis(ageancien: integer ; tabledroit: string) : DOUBLE;   // pt1
  function nbmoisacquis(salarie,convention,etab : string ; datesimul,daterefage,daterefanc : Tdatetime;
   ageancien:integer; tabledroit:string):double; //pt1
  Function probapresence(actuel, restant : integer; tableturn : string): double;
  Function revalorisation(actuel, restant : integer; tablereval : string): double;


Implementation
var
Tablesdynamiques : TTablesdynamiques;
tob_codtab : TOB; // pt5
tob_valsurvie : TOB; // pt5

{*****************************************************************}
{Créé le ...... : 19/06/2006
Modifié le ... :   /  /
Description .. : Calcul de la simulation 
Suite ........ : En paramètre :
Suite ........ :  - La Tob contenant les populations à traiter 
Suite ........ :  - Les paramètres de calcul de la simulation
Mots clefs ... :
*****************************************************************}
procedure calculsimulation(Tob_pop : Tob; paramsimul : TParamsimul);
Var
codesimul, dtsimul , libellesimul, wdatecalcul, typedepart, methcalc, codecalc, population, codtabtur,
codtabrev, codtabsur, codtabdrd, codtabdrr, codepop, WDatenaiss, WDateancien, Wdateretr, wdateentree,
codtabdroit, erreur, selectpop , stpop, salarie, sexe, typetab, typecontrat, pred, convention, etab: string;
agedepd, agedepr, Anciennete,AncAnnee,AncMois, ageannee, agemois ,Ancfinale, ancienfinalaa,
ancienfinalmm, i, ageretr, nbrest, anneesrestAGE, anneesrestANC, ageretrSV : integer;
txactu,  salaireactuel, salairefincarr, txchargepat, survie, moisacquis, actualisation, probapres ,
revaloris , salfincarr, indemniteactual, indemnite , indemnsal: double;
Tob_popfille, Tob_SAL , Tob_salfille, Tob_resultsimul, T1, Tob_SIMUL, T2,Tsalexist,TCD, TOB_SALMOYEN,TSM : tob;     // pt5
QQ, Q1 , Qsqlctr, Qcodtab, Qsalmoy, Qvalsur: TQuery;      // pt5
datesimul , datenaiss, dateancien , datedepretr, datecalcul, dateentree : Tdatetime;
PremMois, PremAnnee,Age , agerest : WORD;
z1 , z2: extended;
Traitsal : boolean;

begin

// Récupération des données de la simulation
  dtsimul := paramsimul.tdatesimul ;

  codesimul :=paramsimul.tcodesimul;
  libellesimul := paramsimul.Tlibellesimul;
  wdatecalcul := paramsimul.Tdatecalcul ;
  typedepart := paramsimul.ttypedepart ;
  methcalc := paramsimul.tmethcalcul ;
  codecalc := paramsimul.Tcodecalcul;
  txactu := paramsimul.Ttauxactual ;
  datesimul := strtodate(dtsimul);
  datecalcul := strtodate(wdatecalcul);

  // Suppression de la simulation si elle existe
  Executesql('Delete from SIMULATION where PSL_CODESIMUL = "'+codesimul+'"');
  Executesql('Delete from RESULTSIMUL where PSR_CODESIMUL = "'+codesimul+'"');

  //  construction de l'objet TABLES dynamiques
  Tablesdynamiques:= TTablesDynamiques.Create;

  // deb pt5
  //chargement de la table CODTAB dans une tob
  Qcodtab := Opensql('select PTI_TYPTAB, PTI_CODTAB,PTI_QUATAB,PTI_CODTABLDYN from CODTAB ', true);
  Tob_CODTAB := Tob.Create('Tables IDR', nil, -1);
  Tob_CODTAB.LoadDetailDB('Tables IDR', '', '', Qcodtab, False);
  ferme(Qcodtab);

  // Chargement de la table SALAIREMOYEN
  Qsalmoy := opensql('select PSY_SALAIREMOY, PSY_SALARIE from SALAIREMOYEN WHERE PSY_DATECALC'+
  '= "'+usdatetime(datecalcul)+'" and PSY_CODECALSAL = "'+codecalc+'"' , true);
  Tob_SALMOYEN := Tob.Create('Salaire moyen', nil, -1);
  Tob_SALMOYEN.LoadDetailDB('Salaire moyen', '', '', Qsalmoy, False);
  ferme(Qsalmoy);

  // chargement des valeurs des tables de survie
  Qvalsur := opensql('select PTS_CODTAB,PTS_VALEUR,PTS_EFFHOMME,PTS_EFFFEMME from CODTABVALEUR WHERE PTS_TYPTAB'+
  '= "SUR" ' , true);
  Tob_VALSURVIE := Tob.Create('Valeurs tables de survie', nil, -1);
  Tob_VALSURVIE.loadDetailDB('Valeurs tables de survie', '', '', Qvalsur, False);
  ferme(Qvalsur);
  // fin pt5

  Tob_resultsimul := TOB.CREATE('RESULTSIMUL', nil, -1);

  selectpop := '';
  // Traitement de chaque population
  if TOB_pop <> nil then
  begin
    Tob_popfille := TOB_pop.FindFirst([''], [''], TRUE);

    while Tob_popfille <> nil do
    begin
    population := Tob_popfille.GetValue('POPULATION');
    pred := Tob_popfille.GetValue('PREDEFINI');
    selectpop := selectpop +  population +';';

    // Recherche des règles de simulation pour cette population
   QQ := opensql('SELECT * FROM REGLESPOP WHERE PPR_POPULATION = "'+population+'" AND PPR_PREDEFINI = "'+pred+'" ORDER BY PPR_PREDEFINI', true);
   if not QQ.EOF then
   begin
   codepop := QQ.findfield('PPR_CODEPOP').asstring;
   codtabsur := QQ.findfield('PPR_CODTABSUR').asstring;
   codtabrev := QQ.findfield('PPR_CODTABREV').asstring;
   codtabtur := QQ.findfield('PPR_CODTABTUR').asstring;;
   codtabdrd := QQ.findfield('PPR_CODTABDRD').asstring;
   codtabdrr := QQ.findfield('PPR_CODTABDRR').asstring;
   agedepd := QQ.findfield('PPR_AGEDEPD').asinteger;
   agedepr := QQ.findfield('PPR_AGEDEPR').asinteger;
   txchargepat := QQ.findfield('PPR_TXCHARGEPAT').asfloat;
   end;

   if typedepart = 'DV'   then
   begin
   // si Départ volontaire
   ageretr := agedepd;
   codtabdroit := codtabdrd;
   end
   else
   begin
   // si Mise en retraite
   ageretr := agedepr;
   codtabdroit := codtabdrr;
   end;
   ageretrSV := ageretr;
   ferme(QQ);

   // Pour chaque salarié appartenant à cette population
   TOB_SAL := salariepop(Codepop, population , datesimul);

    if TOB_SAL <> nil then
    begin
    Tob_salfille := TOB_SAL.FindFirst([''], [''], TRUE);

     while Tob_salfille <> nil do
     begin
      salarie := Tob_salfille.getvalue('SALARIE');
      convention := Tob_salfille.getvalue('CONVENTION'); // pt1
      etab := Tob_salfille.getvalue('ETABL'); // pt1
      
      // contrôle que ce salarié n'a pas été déjà calculé pour cette simulation
       traitsal := true;
      if (T1 <> nil) then
      begin
      Tsalexist := Tob_RESULTSIMUL.FindFirst(['PSR_SALARIE'], [salarie], TRUE);
      if (Tsalexist <> nil) then traitsal := false;
      end;

        If (traitsal = true) then
        begin

        // Traiter les salariés uniquement si non CDD :
        Qsqlctr := Opensql('SELECT PCI_TYPECONTRAT  FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+ Salarie +'" AND' +
        ' PCI_DEBUTCONTRAT <="' + UsDateTime(datesimul) + '" ORDER BY PCI_DEBUTCONTRAT DESC ', true);

        if not Qsqlctr.EOF then   // si au moins un contrat existant
        typecontrat := Qsqlctr.findfield('PCI_TYPECONTRAT').asstring else
        typecontrat := 'CDI';
        ferme(qsqlctr);

        if typecontrat = 'CDI' then
        begin


        erreur := '000';
        sexe := Tob_salfille.getvalue('SEXE');

        Wdateentree:= Tob_salfille.GetValue('DATEENTREE');
        dateentree := strtodate(wdateentree);
        Wdatenaiss := Tob_salfille.GetValue('DATENAISS');
        datenaiss := strtodate(wdatenaiss);

        if Tob_salfille.GetValue('DATEANCIENNETE') = null then
        Wdateancien := Tob_salfille.GetValue('DATEENTREE')  else
        Wdateancien := Tob_salfille.GetValue('DATEANCIENNETE');
        dateancien := strtodate(wdateancien);

        // calcul age salarié, ancienneté actuelle

        If DateNaiss > IDate1900 then AglNombreDeMoisComplet(DateNaiss, Datesimul, PremMois, PremAnnee, Age)
        else Age := 0;
        if dateancien > IDate1900 then Anciennete := AncienneteMois(Dateancien, Datesimul)else anciennete :=0;
        AncAnnee := Anciennete div 12;
        AncMois := Anciennete - (AncAnnee *12) ;
        AgeAnnee := Age div 12;
        Agemois := Age - (AgeAnnee * 12) ;



        // calcul date départ retraite  et ancienneté finale
        datedepretr := plusdate(datenaiss, ageretr, 'A');
        wdateretr := datetostr(datedepretr);
        ancfinale := AncienneteMois(Dateancien, Datedepretr);
        AncienfinalAA := Ancfinale div 12;
        AncienfinalMM := Ancfinale - (AncienfinalAA *12) ;

        anneesrestANC := (ancfinale - anciennete) div 12;   // années d'ancienneté restantes
        anneesrestAGE := ((ageretr * 12) - age) div 12;     // années restantes jusqu'à l'age de la retraite

        // Recherche salaire moyen dans la table SALAIREMOYEN pour le code calcul et la date de calcul saisis
        // en paramètre
         salaireactuel := 0;
        // Q1 := opensql('select PSY_SALAIREMOY from SALAIREMOYEN where PSY_salarie = "'+salarie+'" '+
        //' and PSY_DATECALC = "'+usdatetime(datecalcul)+'" and PSY_CODECALSAL = "'+codecalc+'"' , true);
        //if not Q1.EOF then  salaireactuel := Q1.findfield('PSY_SALAIREMOY').asfloat;    pt5
        TSM := Tob_SALMOYEN.findfirst(['PSY_SALARIE'], [salarie], true);  // pt5
        if TSM <> NIL then  salaireactuel := TSM.getvalue('PSY_SALAIREMOY');  // pt5

        if salaireactuel = 0 then erreur := '001';
        ferme(Q1);

        ageretr := ageretrSV;
        // si age du salarié > age retraite : age retraite = age actuel
        if ageannee > ageretr then
        begin
        ageretr := ageannee;
        ancienfinalaa := ageannee;
        ancienfinalmm := agemois;
        datedepretr := datesimul;
        end;


        // calcul actualisation
        AglNombreDeMoisComplet(Datesimul, Datedepretr, PremMois, PremAnnee, agerest);
        nbrest := agerest div 12;
        z1 := 1 + (txactu / 100);
        z2 :=   power(z1, nbrest);
        actualisation := (1 / Z2) * 100 ;
        actualisation := arrondi(actualisation, 4);

        // calcul coefficient de survie
        survie := coeffsurvie(sexe,ageretr,AgeAnnee, Codtabsur);
        if survie = 0 then erreur := '002';

        // Calcul nombre de mois acquis

        if ((methcalc = '001') or (methcalc = '004'))  // pt4
        then       //  nb mois correspondant à l'ancienneté finale si méthode des unités de crédit projetés
        //moisacquis := nbmoisacquis(ancienfinalaa, codtabdroit)    pt1
        moisacquis := nbmoisacquis(salarie,convention,etab,datesimul,datedepretr,datedepretr,ancienfinalaa,codtabdroit) // pt1
        else       // nb mois correspondant à l'ancienneté actuelle sinon
        // moisacquis := nbmoisacquis(AncAnnee, codtabdroit);     pt1
        moisacquis := nbmoisacquis(salarie,convention,etab,datesimul,datedepretr,datesimul,Ancannee,codtabdroit); // pt1
        if moisacquis = 0 then erreur := '003';

        // calcul probabilité de présence
        // pt5 Q1 := opensql('Select PTI_QUATAB from CODTAB where PTI_TYPTAB = "TUR" and PTI_CODTAB = "'+codtabtur+'"', true);
        probapres := 0;

        TCD := Tob_codtab.findfirst(['PTI_TYPTAB', 'PTI_CODTAB'], ['TUR', codtabtur], true);  // pt5

       // pt5 if not Q1.EOF then
        if TCD <> nil then  //pt5
        begin
        // pt5 typetab := Q1.findfield('PTI_QUATAB').asstring;
        typetab := TCD.getvalue('PTI_QUATAB'); // pt5
        if typetab = 'AGE' then // si qualifiant table = AGE
        probapres := probapresence(ageannee, anneesrestAGE, codtabtur)
        else                   // si qualifiant table = ANC
        probapres := probapresence(AncAnnee, anneesrestANC, codtabtur);
        end;
       //  ferme(Q1);  pt5

        if probapres = 0 then erreur := '005';

        // calcul revalorisation
       // pt5 Q1 := opensql('Select PTI_QUATAB from CODTAB where PTI_TYPTAB = "REV" and PTI_CODTAB = "'+codtabrev+'"', true);
        revaloris:= 0;
        TCD := Tob_codtab.findfirst(['PTI_TYPTAB', 'PTI_CODTAB'], ['REV', codtabrev], true);  // pt5

       // if not Q1.EOF then  pt5
        if TCD <> Nil then // pt5
        begin
        // pt5 typetab := Q1.findfield('PTI_QUATAB').asstring;
        typetab := TCD.getvalue('PTI_QUATAB'); // pt5
        if typetab = 'AGE' then // si qualifiant table = AGE
        revaloris := revalorisation(ageannee, anneesrestAGE, codtabrev)
        else                   // si qualifiant table = ANC
        revaloris := revalorisation(AncAnnee, anneesrestANC, codtabrev);
        end;
        // ferme(Q1);  pt5
        if revaloris = 0 then erreur := '004';

        salfincarr := salaireactuel * revaloris;
        salfincarr := arrondi(salfincarr, 2);

        // calcul de l'indemnité de départ
        indemnite :=  0;
        indemnsal := 0;   // PT2
        if ((methcalc = '002') or (methcalc = '004'))  then   // Méthode rétrospective pure ou PT4 Méth rétro salaire actuel
        begin
        indemnite := salaireactuel * moisacquis * (survie/ 100) * (probapres / 100) * (actualisation/100);
        indemnite := arrondi(indemnite, 2);
        end
        else       // si méthode unités de crédit ou salaire fin carrière : calcul indemnité actualisée
        indemniteactual := salfincarr * moisacquis * (survie/ 100) * (probapres / 100) * (actualisation/100);

        indemnsal := salfincarr * moisacquis;   // PT2

        // si méthode salaire fin carrière : indemnité = indemnité actualisée
        // si méthode des unités de crédit projetés ou Rétro avec salaire actuel: indemnité = indemnité actualisée * (nb années prés / nb années totales)
        if methcalc = '003' then
        indemnite := arrondi(indemniteactual, 2);
        if (methcalc = '001')  and (ancfinale <> 0) then
        begin
        indemnite := indemniteactual * (Anciennete / ancfinale);
        indemnite := arrondi(indemnite, 2);
        end;
        // pt4
        if (methcalc = '004') and (ancfinale <> 0) then
        begin
        indemnite := indemnite * (Anciennete / ancfinale);
        indemnite := arrondi(indemnite, 2);
        end;
        // fin pt4

       // si départ volontaire : Indemnite = Indemnite * (1 + taux charge)
        if typedepart = 'DV' then
        begin
        if (txchargepat = 0) and (erreur = '000') then erreur := '006';
        
        indemnite := indemnite * ( 1 + (txchargepat/100));
        indemnite := arrondi(indemnite, 2);
        end;

        // Alimente la TOB RESULTSIMUL avec les données calculées
        T1 := TOB.CREATE('RESULTSIMUL', Tob_RESULTSIMUL, -1);

        T1.PutValue('PSR_SALARIE', salarie);
        T1.putvalue('PSR_CODESIMUL', codesimul);
        T1.putvalue('PSR_CODEPOP', codepop);
        T1.putvalue('PSR_POPULATION', population);
        T1.putvalue('PSR_SALAIREACTU', salaireactuel);
        T1.putvalue('PSR_SALAIREFINCAR', salfincarr);
        T1.putvalue('PSR_AGEACTUELAA', ageannee);
        T1.putvalue('PSR_AGEACTUELMM', agemois);
        T1.putvalue('PSR_ANCIENACTUAA', ancannee);
        T1.putvalue('PSR_ANCIENACTUMM', ancmois);
        T1.putvalue('PSR_ANCIENFINALAA', ancienfinalaa);
        T1.putvalue('PSR_ANCIENFINALMM', ancienfinalmm);
        T1.putvalue('PSR_AGEDEPRETR', ageretr);
        T1.putvalue('PSR_DATEDEPRETR', datedepretr);
        T1.putvalue('PSR_TXCHARGE',  txchargepat);
        T1.putvalue('PSR_ACTUALISATION', actualisation);
        T1.putvalue('PSR_PROBASUIVIE', survie);
        T1.putvalue('PSR_MOISACQUIS', moisacquis);
        T1.putvalue('PSR_PROBAPRESENCE', probapres);
        T1.putvalue('PSR_REVALORISATION', revaloris);
        T1.putvalue('PSR_INDEMNBRUT', indemnsal);   // PT2
        T1.putvalue('PSR_INDEMNLEGALE', indemnite);
        T1.putvalue('PSR_NOM', Tob_salfille.getvalue('NOM'));
        T1.putvalue('PSR_PRENOM', Tob_salfille.getvalue('PRENOM'));
        T1.putvalue('PSR_SEXE', Tob_salfille.getvalue('SEXE'));
        T1.putvalue('PSR_NUMEROSS', Tob_salfille.getvalue('NUMEROSS'));
        T1.putvalue('PSR_ETABLISSEMENT', Tob_salfille.getvalue('ETABL'));
        T1.putvalue('PSR_DATECALANC', dateancien);
        T1.putvalue('PSR_DATEENTREE', dateentree);
        T1.putvalue('PSR_DATENAISSANCE', dateNAISS);
        T1.putvalue('PSR_TRAVAILN1', Tob_salfille.getvalue('TRAVAILN1'));
        T1.putvalue('PSR_TRAVAILN2', Tob_salfille.getvalue('TRAVAILN2'));
        T1.putvalue('PSR_TRAVAILN3', Tob_salfille.getvalue('TRAVAILN3'));
        T1.putvalue('PSR_TRAVAILN4', Tob_salfille.getvalue('TRAVAILN4'));
        T1.putvalue('PSR_LIBREPCMB1', Tob_salfille.getvalue('LIBREPCMB1'));
        T1.putvalue('PSR_LIBREPCMB2', Tob_salfille.getvalue('LIBREPCMB2'));
        T1.putvalue('PSR_LIBREPCMB3', Tob_salfille.getvalue('LIBREPCMB3'));
        T1.putvalue('PSR_LIBREPCMB4', Tob_salfille.getvalue('LIBREPCMB4'));
        T1.putvalue('PSR_CODESTAT', Tob_salfille.getvalue('CODESTAT'));
        T1.putvalue('PSR_DADSPROF', Tob_salfille.getvalue('DADSPROF'));
        T1.putvalue('PSR_LIBELLEEMPLOI', Tob_salfille.getvalue('LIBELLEEMPLOI'));
        T1.putvalue('PSR_CONVENTION', Tob_salfille.getvalue('CONVENTION'));
        T1.putvalue('PSR_CONDEMPLOI', Tob_salfille.getvalue('CONDEMPLOI'));
        T1.putvalue('PSR_DADSCAT', Tob_salfille.getvalue('DADSCAT'));
        T1.putvalue('PSR_TYPECONTRAT', Typecontrat);
        // PT3
        T1.putvalue('PSR_CODTABSUR', codtabsur);
        T1.putvalue('PSR_CODTABTUR', codtabtur);
        T1.putvalue('PSR_CODTABREVAL', codtabrev);
        T1.putvalue('PSR_CODTABDROIT', codtabdroit);
        // fin PT3
        T1.putvalue('PSR_STATUTSIMUL', ERREUR);

        end;
        end;
        
       Tob_salfille := TOB_sal.FindNext([''], [''], TRUE);
      end;
      end;
      freeandnil(TOB_SAL);
      Tob_popfille := TOB_pop.FindNext([''], [''], TRUE);
   end;

   // Mise à jour de la table SIMULATION
   stpop := copy(selectpop, 1, length(selectpop) - 1);
   Tob_simul := TOB.CREATE('SIMULATION', nil, -1);
   T2 := TOB.CREATE('SIMULATION', Tob_SIMUL, -1);
   T2.putvalue('PSL_CODESIMUL', codesimul);
   T2.putvalue('PSL_LIBELLE', libellesimul );
   T2.putvalue('PSL_DATECALC', datecalcul);
   T2.putvalue('PSL_DATESIMUL', datesimul );
   T2.putvalue('PSL_TYPEDEPART', typedepart);
   T2.putvalue('PSL_METHCALCUL', methcalc);
   T2.putvalue('PSL_TAUXACTUAL', Txactu);
   T2.putvalue('PSL_CODECALSAL', Codecalc);
   T2.putvalue('PSL_SELECTPOP', Stpop);

   T2.InsertDB(nil, false);

   freeandnil(Tob_simul);

   // Mise à jour de la table RESULTSIMUL
   For i := 0 to Tob_resultsimul.detail.Count - 1 do
   begin
    T1 := Tob_resultsimul.Detail[i];
    T1.InsertDB(nil, false);
   end;

  freeandnil(Tob_resultsimul);
  freeandnil(Tob_pop);

  end;
// pt1 destruction objet tables dynamiques
If assigned(Tablesdynamiques) then Tablesdynamiques.free;
if assigned(Tob_codtab) then freeandnil(Tob_codtab);   // pt5
if assigned(Tob_salmoyen) then freeandnil(Tob_salmoyen);   // pt5
if assigned(Tob_VALSURVIE) then freeandnil(Tob_VALSURVIE);   // pt5
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/06/2006
Modifié le ... :   /  /
Description .. : Calcul du coefficient de survie
Suite ........ : Paramètres: Sexe, age départ retraite, age salarié, code de 
Suite ........ : la table
Mots clefs ... : 
*****************************************************************}
function coeffsurvie(sexe: string; ageret,agesal : integer; tablesur: string) : DOUBLE;
var
//Q : Tquery;    pt5
effactu , effretr : integer;
resultat : double;
TSUR : TOB; // pt5

begin
effactu := 0;
effretr := 0;
result := 0;


//Q:= opensql('Select PTS_EFFHOMME, PTS_EFFFEMME from CODTABVALEUR where PTS_TYPTAB = "SUR"'+
//' and PTS_CODTAB = "'+tablesur+'" and PTS_VALEUR = "'+inttostr(ageret)+'"', true);  pt5
TSUR := Tob_VALSURVIE.findfirst(['PTS_CODTAB', 'PTS_VALEUR'], [tablesur,ageret], true);  // pt5

//if not Q.EOF then  pt5
if TSUR <> NIL then // pt5
begin
  if sexe = 'F' then
  // pt5 effretr := Q.findfield('PTS_EFFFEMME').asinteger else
  // pt5 effretr := Q.findfield('PTS_EFFHOMME').asinteger;
  effretr := TSUR.getvalue('PTS_EFFFEMME') else    // pt5
  effretr := TSUR.getvalue('PTS_EFFHOMME');        // pt5
end;

//ferme(Q); pt5

// Q:= opensql('Select PTS_EFFHOMME, PTS_EFFFEMME from CODTABVALEUR where PTS_TYPTAB = "SUR"' +
//' and PTS_CODTAB = "'+tablesur+'" and PTS_VALEUR = "'+inttostr(agesal)+'"', true);   pt5
TSUR := Tob_VALSURVIE.findfirst(['PTS_CODTAB', 'PTS_VALEUR'], [tablesur,agesal], true);  // pt5
// pt5 if not Q.EOF then
if  TSUR <> NIL then // pt5
begin
  if sexe = 'F' then
 // effactu := Q.findfield('PTS_EFFFEMME').asinteger else
 // effactu := Q.findfield('PTS_EFFHOMME').asinteger;
  effactu := TSUR.getvalue('PTS_EFFFEMME') else    // pt5
  effactu := TSUR.getvalue('PTS_EFFHOMME');        // pt5
end;

//ferme(Q);  pt5

if effactu <> 0 then resultat := (effretr / effactu) * 100 else resultat := 0;
result := arrondi(resultat, 4);

end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/06/2006
Modifié le ... :   /  /    
Description .. : Calcul nombre de mois acquis :
Suite ........ : En Entrée:     - Salarié, convention collective, établissement
                                - Date de simulation, date référence age, date référence ancienneté
                                - l'ancienneté finale
                                - le code de la table des droits
                 En sorite      - Le nombre de mois acquis
Mots clefs ... : 
*****************************************************************}
function nbmoisacquis(salarie,convention,etab: string; datesimul, daterefage,daterefanc : TDatetime; ageancien : integer; tabledroit: string) : DOUBLE;    // pt1
var
Q : Tquery;    // pt1
codtabldyna : string;  // pt1
TCD : Tob;  // pt5
begin
  result := 0;
  // début pr1
   // pt5 Q1 := opensql('select  PTI_CODTABLDYN from CODTAB where PTI_CODTAB = "'+tabledroit+'"', true);
   TCD := Tob_codtab.findfirst(['PTI_CODTAB'], [tabledroit], true);  // pt5

  //if not Q1.EOF then pt5
   if TCD <> NIL then // pt5
   begin
     // pt5 codtabldyna := Q1.findfield('PTI_CODTABLDYN').asstring;
     codtabldyna := TCD.getvalue('PTI_CODTABLDYN'); // pt5
     if codtabldyna <> ''  // si table dynamique renseigné
     then
     begin
      daterefage := plusdate(daterefage, 1, 'J');
      daterefanc := plusdate(daterefanc, 1, 'J'); // ajout 1 jour pour prendre en compte la date comme journée complète
      result := Tablesdynamiques.GetValueFromDynaTableDSA(salarie,codtabldyna,datesimul,
      etab,convention,daterefage,daterefanc);
     end
     else
     begin // fin pt1

      Q:= opensql('Select PTS_NBMOIS  from CODTABVALEUR where PTS_TYPTAB = "DRT" and '+
      ' PTS_CODTAB = "'+tabledroit+'" and PTS_VALEUR = "'+inttostr(ageancien)+'"', true);

      if not Q.EOF then result := Q.findfield('PTS_NBMOIS').asfloat;
      ferme(Q);
     end;
  end;
// ferme(Q1)  pt5
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 29/06/2006
Modifié le ... :   /  /
Description .. : Calcul probabilité de présence :
Suite ........ : Paramètres :  ancienneté ou age actuel 
Suite ........ :               Nb années restantes avant le départ en retraite
Suite ........ :               code table TURN OVER
Mots clefs ... :
*****************************************************************}
function probapresence(actuel, restant : integer; tableturn: string) : DOUBLE;
var
reste, Val, total, nbrest : integer;
z1 , z2: extended;
Q1, Q2 : TQuery;
tx, resultat : double;
fini : boolean;
stactu : string;

begin
fini := false;
stactu := inttostr(actuel);
result := 0;

// Lecture de la borne correspondant à l'age ou ancienneté actuel
Q1:= opensql('SELECT PTS_TAUX, PTS_VALEUR from codtabvaleur where pts_typtab = "TUR" and' +
' pts_codtab = "'+tableturn+'" and pts_valeur <= "'+stactu+'"  order by pts_valeur desc', true);

if not Q1.EOF then
begin

tx := Q1.findfield('PTS_TAUX').asfloat;
val := Q1.findfield('PTS_VALEUR').asinteger;
reste := actuel;
resultat := 1;
result := 0;
total := 0;
nbrest := 0;

// Lecture de la borne supérieure
Q2:= opensql('SELECT PTS_TAUX, PTS_VALEUR from codtabvaleur where pts_typtab = "TUR" and' +
' pts_codtab = "'+tableturn+'" and pts_valeur > "'+stactu+'"', true);

  if not Q2.EOF then
  begin
  Q2.First;
    while not Q2.EOF do
    begin
    val := Q2.findfield('PTS_VALEUR').asinteger;
    total := total + nbrest;
    
    if (total + nbrest) > restant then
    begin
    nbrest := restant - total;
    fini := true;
    end
    else
    nbrest := val - reste;

    z1 := 1 - (tx/100);
    z2 := power(z1, nbrest);
    resultat := resultat * z2;

    if fini then break;

    reste := val;
    tx := Q2.findfield('PTS_TAUX').asfloat;
    Q2.next;
    end;

  end;
    if (fini = false) then
     // s'il n'y a plus de borne >, le nombre d'années restant correspond au dernier taux
    begin
    nbrest := restant - total;
    z1 := 1 - (tx/100);
    z2 := power(z1, nbrest);
    resultat := resultat * z2;
    end;
  resultat := resultat * 100;
  result := arrondi(resultat, 4);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 29/06/2006
Modifié le ... :   /  /
Description .. : Calcul taux de revalorisation :
Suite ........ : Paramètres :  ancienneté ou age actuel 
Suite ........ :               Nb années restantes avant le départ en retraite
Suite ........ :               code table TURN OVER
Mots clefs ... :
*****************************************************************}
function revalorisation(actuel, restant : integer; tablereval: string) : DOUBLE;
var
reste, Val, total, nbrest : integer;
z1 , z2: extended;
Q1, Q2 : TQuery;
tx, resultat : double;
fini : boolean;
stactu : string;

begin
fini := false;
stactu := inttostr(actuel);
result := 0;

// Lecture de la borne correspondant à l'age ou ancienneté actuel
Q1:= opensql('SELECT PTS_TAUX, PTS_VALEUR from codtabvaleur where pts_typtab = "REV" and' +
' pts_codtab = "'+tablereval+'" and pts_valeur <= "'+stactu+'"  order by pts_valeur desc', true);

if not Q1.EOF then
begin

tx := Q1.findfield('PTS_TAUX').asfloat;
val := Q1.findfield('PTS_VALEUR').asinteger;
reste := actuel;
resultat := 1;
result := 0;
total := 0;
nbrest := 0;

// Lecture de la borne supérieure
Q2:= opensql('SELECT PTS_TAUX, PTS_VALEUR from codtabvaleur where pts_typtab = "REV" and' +
' pts_codtab = "'+tablereval+'" and pts_valeur > "'+stactu+'"', true);

  if not Q2.EOF then
  begin
  Q2.First;
    while not Q2.EOF do
    begin
    val := Q2.findfield('PTS_VALEUR').asinteger;
    total := total + nbrest;
    
    if (total + nbrest) > restant then
    begin
    nbrest := restant - total;
    fini := true;
    end
    else
    nbrest := val - reste;

    z1 := 1 + (tx/100);
    z2 := power(z1, nbrest);
    resultat := resultat * z2;

    if fini then break;

    reste := val;
    tx := Q2.findfield('PTS_TAUX').asfloat;
    Q2.next;
    end;

  end;
    if (fini = false) then
     // s'il n'y a plus de borne >, le nombre d'années restant correspond au dernier taux
    begin
    nbrest := restant - total;
    z1 := 1 + (tx/100);
    z2 := power(z1, nbrest);
    resultat := resultat * z2;
    end;

  result := arrondi(resultat, 5);
end;
end;

end.


