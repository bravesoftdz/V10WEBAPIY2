{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 16/05/2006
Modifié le ... :   /  /
Description .. : PGCALCULSALAIRE ()  Calcul du salaire moyen
Mots clefs ... :
*****************************************************************
PT1    :  17/04/2007  NA  Calcul du salaire moyen des 3 derniers mois proratisé ou non, le salaire
                          à retenir est le plus avantageux entre le salaire moyen et celui des 3 derniers mois}

Unit PGCALCULSALAIRE ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     UTobDebug,
     ENTPAIE;


  procedure calculsalmoyen(stwhere, Codecal : string; Zdatecal, Zdatedebut, Zdatefin : Tdatetime;salcompare : boolean; var tob_erreur : Tob);
  procedure calculsaldernmois(dtfin : Tdatetime);
  function calculcumulsal(salarie, listecum : string; dtdeb, dtfin : TDatetime):  double;
  function calculbaseremsal(salarie, listerem : string; dtdeb, dtfin : Tdatetime) : double;
  function calculmtremsal(salarie, listerem : string; dtdeb, dtfin : Tdatetime) : double;
  function calculnbtrentsal(salarie : string; dtdeb, dtfin : Tdatetime) : integer;
  procedure calculcumul(listecum, sens : string; calculperiode, proratis : boolean; dtdeb, dtfin : TDatetime);
  procedure calculbaserem(listerem, sens : string; calculperiode, proratis : boolean; dtdeb, dtfin : Tdatetime);
  procedure calculmtrem(listerem,sens : string; calculperiode, proratis : boolean; dtdeb, dtfin : Tdatetime);
  procedure calculnbtrent(calculperiode : boolean; dtdeb, dtfin : Tdatetime);
  procedure calculsal(salcompare : boolean);
Implementation
var // pt1
listecumplusPR, listecummoinPR, listebasremplusPR,listebasremmoinPR,listemtremplusPR,
listemtremmoinPR , listecumplusNPR, listecummoinNPR, listebasremplusNPR,listebasremmoinNPR,listemtremplusNPR,
listemtremmoinNPR ,listecumplus, listecummoin, listebasremplus, listebasremmoin, listemtremplus,
listemtremmoin,  lewhere: string;
tob_calculsalaire, Tob_calculsalairedernmois :  Tob;
datedebut, datefin: Tdatetime;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/05/2006
Modifié le ... :   /  /
Description .. : Calcul du salaire moyen avec les paramètres suivants :
                 - les salariés stockés dans la tob Tob_sal
                 - le code calcul
                 - la date de calcul
                 - la date de début de période
                 - la date de fin de période
                 - salaire moyen de la période comparé au salaire des 3 derniers mois
Mots clefs ... :
*****************************************************************}
procedure calculsalmoyen(stwhere, Codecal : string; Zdatecal, Zdatedebut, Zdatefin : Tdatetime; salcompare : boolean; var tob_erreur : Tob);
var
st , sens, remcum, Basemont, naturecal, sal, typecontrat, travailn1, travailn2, travailn3, travailn4,
codestat, convention, condemploi, libelleemploi, etabl, dadsprof, dadscat, librepcmb1, librepcmb2,
librepcmb3, librepcmb4 ,proratise, senscalculsal, salarie: string;
Qsql, qsqlsal, qsqlhis : Tquery;
Tob_salmoyen, T1, T,  tob_erreurfille, TCD, Tob_HISTOSAL, THS: Tob;
histo, trouvehisto, calculperiode, proratis : boolean;
dtdeb,dtfin : Tdatetime;
salairemoy, salairemoyper,salairemoyderm : double;
i : integer;
begin
lewhere := stwhere;
datedebut := zdatedebut;
datefin := zdatefin;
// chargement de la méthode de calcul
listecumplus := '';
listecummoin := '';
listebasremplus := '';
listebasremmoin := '';
listemtremplus := '';
listemtremmoin := '';
// deb pt1
listecumplusPR := '';
listecummoinPR := '';
listebasremplusPR := '';
listebasremmoinPR := '';
listemtremplusPR := '';
listemtremmoinPR := '';
listecumplusNPR := '';
listecummoinNPR := '';
listebasremplusNPR := '';
listebasremmoinNPR := '';
listemtremplusNPR := '';
listemtremmoinNPR := '';
// fin pt1

// Recherche des codes cumul   et des rémunérations
st := 'Select PSM_RUBRCUM, PSM_SENSCALC, PSM_NATURECAL, PSM_BASEMONT,PSM_PRORATISE from METHCALCULSALMOY where PSM_CODECALSAL = "'+codecal+'"';

Qsql := Opensql(st, true);

while not Qsql.eof do
begin

  remcum := Qsql.findfield('PSM_RUBRCUM').asstring;
  sens := Qsql.findfield('PSM_SENSCALC').asstring;
  naturecal := Qsql.findfield('PSM_NATURECAL').asstring;
  basemont := Qsql.findfield('PSM_BASEMONT').asstring;
  proratise := QSqL.findfield('PSM_PRORATISE').asstring;    // pt1

  // si CUMUL
  if naturecal = 'CUM' then
  begin

   if sens = '09' then
   begin
      if (proratise = 'X') then listecumplusPR :=  listecumplusPR + '"'+remcum+'" ,' // à ajouter si proratisé
      else
      listecumplusNPR :=  listecumplusNPR + '"'+remcum+'" ,';    // à ajouter si non proratisé pt1

      listecumplus := listecumplus + '"'+remcum+'" ,';   // à ajouter
   end
   else
   begin
      if (proratise = 'X') then listecummoinPR :=  listecummoinPR + '"'+remcum+'" ,'  // à soustraire si proratisé   pt1
      else
      listecummoinNPR :=  listecummoinNPR + '"'+remcum+'" ,'; // à soustraire si non proratisé  pt1

      listecummoin := listecummoin + '"'+remcum+'" ,';     // à soustraire
   end;
  end;

  // Si remunération
  if naturecal = 'REM' then
  begin
    if sens = '09' then
    begin    // à ajouter
      if basemont = 'MON' then // Montant rem
      begin
        if (proratise = 'X') then listemtremplusPR:= listemtremplusPR + '"'+remcum+'" ,' // Si montant rem et proratisé pt1
        else
        listemtremplusNPR:= listemtremplusNPR + '"'+remcum+'" ,'; // Si montant rem et non  proratisé pt1

        listemtremplus := listemtremplus + '"'+remcum+'" ,'; // Si montant rem
      end
      else
      begin                   // Base rem
        if (proratise = 'X') then listebasremplusPR:= listebasremplusPR + '"'+remcum+'" ,' // Si base rem et proratisé pt1
        else
        listebasremplusNPR:= listebasremplusNPR + '"'+remcum+'" ,'; // Si base rem et non  proratisé pt1

        listebasremplus := listebasremplus + '"'+remcum+'" ,'; // si base rem
      end;
    end;

    if sens = '10' then
    begin   // à soustraire
      if basemont = 'MON' then  // Montant rem
      begin
        if (proratise = 'X') then listemtremmoinPR:= listemtremmoinPR + '"'+remcum+'" ,' // Si montant rem et proratisé pt1
        else
        listemtremmoinNPR:= listemtremmoinNPR + '"'+remcum+'" ,'; // Si montant rem et non  proratisé pt1

        listemtremmoin := listemtremmoin + '"'+remcum+'" ,'; // si montant rem
      end
      else
      begin     // Base rem
        if (proratise = 'X') then listebasremmoinPR:= listebasremmoinPR + '"'+remcum+'" ,' // Si base rem et proratisé pt1
        else
        listebasremmoinNPR:= listebasremmoinNPR + '"'+remcum+'" ,'; // Si base rem et non  proratisé pt1

        listebasremmoin := listebasremmoin + '"'+remcum+'" ,'; // si base rem
      end;
   end;
 end;

 Qsql.next;
end;

ferme(Qsql);
listecumplus := '(' + copy(listecumplus, 1, length(listecumplus) - 2) + ')';
listecummoin := '(' + copy(listecummoin, 1, length(listecummoin) - 2) + ')';
listemtremplus := '(' + copy(listemtremplus, 1, length(listemtremplus) - 2) + ')';
listemtremmoin := '(' + copy(listemtremmoin, 1, length(listemtremmoin) - 2) + ')';
listebasremplus := '(' + copy(listebasremplus, 1, length(listebasremplus) - 2) + ')';
listebasremmoin := '(' + copy(listebasremmoin, 1, length(listebasremmoin) - 2) + ')';
// deb pt1
listecumplusPR := '(' + copy(listecumplusPR, 1, length(listecumplusPR) - 2) + ')';
listecummoinPR := '(' + copy(listecummoinPR, 1, length(listecummoinPR) - 2) + ')';
listemtremplusPR := '(' + copy(listemtremplusPR, 1, length(listemtremplusPR) - 2) + ')';
listemtremmoinPR := '(' + copy(listemtremmoinPR, 1, length(listemtremmoinPR) - 2) + ')';
listebasremplusPR := '(' + copy(listebasremplusPR, 1, length(listebasremplusPR) - 2) + ')';
listebasremmoinPR := '(' + copy(listebasremmoinPR, 1, length(listebasremmoinPR) - 2) + ')';
listecumplusNPR := '(' + copy(listecumplusNPR, 1, length(listecumplusNPR) - 2) + ')';
listecummoinNPR := '(' + copy(listecummoinNPR, 1, length(listecummoinNPR) - 2) + ')';
listemtremplusNPR := '(' + copy(listemtremplusNPR, 1, length(listemtremplusNPR) - 2) + ')';
listemtremmoinNPR := '(' + copy(listemtremmoinNPR, 1, length(listemtremmoinNPR) - 2) + ')';
listebasremplusNPR := '(' + copy(listebasremplusNPR, 1, length(listebasremplusNPR) - 2) + ')';
listebasremmoinNPR := '(' + copy(listebasremmoinNPR, 1, length(listebasremmoinNPR) - 2) + ')';
// fin pt1

 Tob_salmoyen := TOB.CREATE('Salairemoyen', nil, -1);
 Tob_erreur := TOB.CREATE('Salariés en erreur', nil, -1);

 Tob_calculsalaire := TOB.CREATE('Calculsalairemoyen', nil, -1);


    // calcul du salaire sur la période pour tous les salariés
    dtdeb := datedebut;
    dtfin := datefin;
    calculperiode := true;
    proratis := false;
    senscalculsal := '+';
    calculcumul(listecumplus,senscalculsal,calculperiode, proratis, dtdeb, dtfin);     // calcul montant cumul plus
    calculbaserem(listebasremplus,senscalculsal, calculperiode,proratis,dtdeb, dtfin); // calcul base rem plus
    calculmtrem(listemtremplus, senscalculsal, calculperiode,proratis,dtdeb, dtfin); // calcul montant rem plus
    senscalculsal := '-';
    calculcumul(listecummoin,senscalculsal,calculperiode,proratis, dtdeb, dtfin);     // calcul montant cumul moins
    calculbaserem(listebasremmoin,senscalculsal, calculperiode,proratis,dtdeb, dtfin); // calcul base rem moins
    calculmtrem(listemtremmoin,senscalculsal, calculperiode,proratis,dtdeb, dtfin); // calcul montant rem moins
    calculnbtrent(calculperiode,dtdeb,dtfin);
    calculsal(salcompare);



    // Mise à jour de la table SALAIREMOYEN

       // Recherche des informations du salarié
     Qsqlsal := Opensql('Select PSA_SALARIE, PSA_ETABLISSEMENT, PSA_DATEENTREE, PSA_LIBELLE, PSA_PRENOM,' +
     ' PSA_DATENAISSANCE, PSA_SEXE, PSA_TRAVAILN1, PSA_TRAVAILN2, PSA_TRAVAILN3, PSA_TRAVAILN4, PSA_CODESTAT,' +
     ' PSA_CONVENTION,PSA_CONDEMPLOI, PSA_LIBELLEEMPLOI, PSA_DADSPROF, PSA_DADSCAT, PSA_LIBREPCMB1, PSA_LIBREPCMB2'+
     ', PSA_LIBREPCMB3, PSA_LIBREPCMB4, PCI_TYPECONTRAT from SALARIES LEFT JOIN CONTRATTRAVAIL ON '+
     ' ((PCI_SALARIE = PSA_SALARIE) And (PCI_DEBUTCONTRAT <= "' + UsDateTime(zdatecal) + '"))' +
     ' '+lewhere+'  and ((PCI_DEBUTCONTRAT is NULL) or (PCI_DEBUTCONTRAT = (SELECT MAX(PCI_DEBUTCONTRAT)'+
     ' FROM CONTRATTRAVAIL WHERE PCI_SALARIE = PSA_SALARIE AND PCI_DEBUTCONTRAT <= "'+UsDateTime(zdatecal)+'")))' +
     '', true);


     histo := false;
     if vh_paie.pghistorisation then     // si historisation
     begin
     qsqlhis := Opensql('SELECT PSA_SALARIE, PHS_SALARIE, PHS_ETABLISSEMENT, PHS_TRAVAILN1, PHS_TRAVAILN2, '+
     'PHS_TRAVAILN3, PHS_TRAVAILN4, PHS_CODESTAT, PHS_CONDEMPLOI, PHS_LIBELLEEMPLOI, PHS_DADSPROF, PHS_DADSCAT,'+
     ' PHS_CBLIBRE1, PHS_CBLIBRE2,PHS_CBLIBRE3, PHS_CBLIBRE4 from SALARIES LEFT JOIN HISTOSALARIE ON '+
     ' (PSA_SALARIE = PHS_SALARIE) '+lewhere+' AND PHS_DATEAPPLIC in (SELECT MAX(PHS_DATEAPPLIC) FROM'+
     ' HISTOSALARIE WHERE PHS_SALARIE = PSA_SALARIE AND PHS_DATEAPPLIC <= "'+ UsDateTime(zdatecal)+'")', true);
     if not qsqlhis.EOF then
     begin
     histo := true;
     Tob_Histosal:= Tob.create('Histosal', nil, -1);
     Tob_Histosal.loaddetaildb('Histosal', '', '', Qsqlhis, False, FALSE);
     end;
     ferme(qsqlhis);
     end;

    
     // mise à jour table SALAIREMOYEN
      while not qsqlsal.Eof do
      begin
        salarie :=  qsqlsal.findfield('PSA_SALARIE').asstring;
        convention := qsqlsal.findfield('PSA_CONVENTION').asstring;
        typecontrat := qsqlsal.findfield('PCI_TYPECONTRAT').asstring;
        if typecontrat = '' then typecontrat := 'CDI';

        trouvehisto := false;
        if histo = true then
        begin

        // recherche s'il existe un historique pour le salarié traité
        THS := Tob_Histosal.findfirst(['PHS_SALARIE'], [salarie], true);
        if THS <>  nil then trouvehisto := true;
        end;


        if not trouvehisto then
        begin     // données prises dans la fiche salarié si non historisation ou non trouvé dans historique

          etabl := Qsqlsal.findfield('PSA_ETABLISSEMENT').asstring;
          travailn1 := qsqlsal.findfield('PSA_TRAVAILN1').asstring;
          travailn2 := qsqlsal.findfield('PSA_TRAVAILN2').asstring;
          travailn3 := qsqlsal.findfield('PSA_TRAVAILN3').asstring;
          travailn4 := qsqlsal.findfield('PSA_TRAVAILN4').asstring;
          codestat := qsqlsal.findfield('PSA_codestat').asstring;
          condemploi := qsqlsal.findfield('PSA_CONDEMPLOI').asstring;
          libelleemploi := qsqlsal.findfield('PSA_LIBELLEEMPLOI').asstring;
          dadsprof := qsqlsal.findfield('PSA_DADSPROF').asstring;
          dadscat := qsqlsal.findfield('PSA_DADSCAT').asstring;
          librepcmb1 := qsqlsal.findfield('PSA_LIBREPCMB1').asstring;
          librepcmb2 := qsqlsal.findfield('PSA_LIBREPCMB2').asstring;
          librepcmb3 := qsqlsal.findfield('PSA_LIBREPCMB3').asstring;
         librepcmb4 := qsqlsal.findfield('PSA_LIBREPCMB4').asstring;
       end
       else
       begin
          // données prises dans HISTOSALARIE
        etabl := THS.getvalue('PHS_ETABLISSEMENT');
        travailn1 := THS.getvalue('PHS_TRAVAILN1');
        travailn2 := THS.getvalue('PHS_TRAVAILN2');
        travailn3 := THS.getvalue('PHS_TRAVAILN3');
        travailn4 := THS.getvalue('PHS_TRAVAILN4');
        codestat := THS.getvalue('PHS_codestat');
        condemploi := THS.getvalue('PHS_CONDEMPLOI');
        libelleemploi := THS.getvalue('PHS_LIBELLEEMPLOI');
        dadsprof := THS.getvalue('PHS_DADSPROF');
        dadscat := THS.getvalue('PHS_DADSCAT');
        librepcmb1 := THS.getvalue('PHS_CBLIBRE1');
        librepcmb2 := THS.getvalue('PHS_CBLIBRE2');
        librepcmb3 := THS.getvalue('PHS_CBLIBRE3');
        librepcmb4 := THS.getvalue('PHS_CBLIBRE4');
       end;
     
     // Suppression du calcul précédent

     ExecuteSQL('DELETE FROM SALAIREMOYEN WHERE PSY_SALARIE = "'+salarie+'" and '+
     'PSY_CODECALSAL = "'+codecal+'" AND PSY_DATECALC ="'+USDATETIME(zdatecal)+'" ');

     // recherche du calcul pour chaque salarié
      TCD := Tob_calculsalaire.findfirst(['SALARIE'], [salarie], true);
      if TCD <> nil then
      begin
      salairemoy := TCD.getvalue('SALAIREMOY');
      salairemoyper := TCD.getvalue('SALAIREMOYPER');
      salairemoyderm := TCD.getvalue('SALAIREMOYDERM');
      end
      else
      begin
      salairemoy := 0;
      salairemoyper := 0;
      salairemoyderm := 0;
      end;

     // traitement des erreurs
     if salairemoy = 0 then
     begin
     tob_erreurfille := Tob.Create('Salariés en erreur',Tob_erreur,-1);
     Tob_erreurfille.AddChampSupValeur('PSA_SALARIE',sal);
     Tob_erreurfille.AddChampSupValeur('PSA_LIBELLE',Qsqlsal.findfield('PSA_LIBELLE').asstring);
     Tob_erreurfille.AddChampSupValeur('PSA_PRENOM',Qsqlsal.findfield('PSA_PRENOM').asstring);
     end;

     T1 := TOB.CREATE('SALAIREMOYEN', Tob_Salmoyen, -1);

     T1.PutValue('PSY_SALARIE', salarie);
     T1.putvalue('PSY_DATECALC', zdatecal);
     T1.putvalue('PSY_CODECALSAl', codecal);
     T1.putvalue('PSY_DATEDEBUT', dtdeb);
     T1.putvalue('PSY_DATEFIN', dtfin);
     T1.putvalue('PSY_SALAIREMOYPER', salairemoyper);  // pt1
     T1.putvalue('PSY_SALAIREMOY', salairemoy); // pt1
     T1.putvalue('PSY_SALAIREMOYDERM', salairemoyderm); // pt1
     T1.putvalue('PSY_ETABLISSEMENT',etabl);
     T1.putvalue('PSY_DATEENTREE',Qsqlsal.findfield('PSA_DATEENTREE').asdatetime);
     T1.putvalue('PSY_NOM',Qsqlsal.findfield('PSA_LIBELLE').asstring);
     T1.putvalue('PSY_PRENOM',Qsqlsal.findfield('PSA_PRENOM').asstring);
     T1.putvalue('PSY_DATENAISSANCE',Qsqlsal.findfield('PSA_DATENAISSANCE').asdatetime);
     T1.putvalue('PSY_SEXE',Qsqlsal.findfield('PSA_SEXE').asstring);
     T1.putvalue('PSY_TRAVAILN1', travailn1);
     T1.putvalue('PSY_TRAVAILN2', travailn2);
     T1.putvalue('PSY_TRAVAILN3', travailn3);
     T1.putvalue('PSY_LIBREPCMB1', librepcmb1);
     T1.putvalue('PSY_LIBREPCMB2', librepcmb2);
     T1.putvalue('PSY_LIBREPCMB3', librepcmb3);
     T1.putvalue('PSY_LIBREPCMB4', librepcmb4);
     T1.putvalue('PSY_CODESTAT', Codestat);
     T1.putvalue('PSY_CONVENTION', Convention);
     T1.putvalue('PSY_CONDEMPLOI', Condemploi);
     T1.putvalue('PSY_LIBELLEEMPLOI', libelleemploi);
     T1.putvalue('PSY_DADSPROF', dadsprof);
     T1.putvalue('PSY_DADSCAT', dadscat);
     T1.putvalue('PSY_TYPECONTRAT', typecontrat);

     Qsqlsal.next;
     end;

    ferme(Qsqlsal);
   For i := 0 to Tob_salmoyen.detail.Count - 1 do
   begin
    T := Tob_salmoyen.Detail[i];
    T.InsertDB(nil, false);
   end;


if Assigned(Tob_Salmoyen) then freeandnil(TOB_salmoyen);
if Assigned(Tob_calculsalaire) then freeandnil(TOB_Calculsalaire);
if Assigned(Tob_Histosal) then freeandnil(Tob_Histosal);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/04/2007
Modifié le ... :   /  /    
Description .. : Calcul du salaire moyen des 3 derniers mois
Mots clefs ... : 
*****************************************************************}
// deb pt1
procedure calculsaldernmois(dtfin : Tdatetime);
var
dtdeb : Tdatetime;
calculperiode, proratis : boolean;
senscalculsal : string;

begin

  Tob_calculsalairedernmois := TOB.CREATE('Calculsalairedernmois', nil, -1);


    // calcul du salaire sur la période pour tous les salariés
    dtfin := datefin;
    dtdeb := plusdate(dtfin, -3, 'M');
    calculperiode := false;

    // calcul du paramétrage "A proratiser"
    proratis := true;
    senscalculsal := '+';
    calculcumul(listecumplusPR,senscalculsal,calculperiode,proratis, dtdeb, dtfin); // calcul cumul plus proratisé
    calculbaserem(listebasremplusPR,senscalculsal,calculperiode,proratis, dtdeb, dtfin); // calcul base rem plus proratisé
    calculmtrem(listemtremplusPR, senscalculsal, calculperiode, proratis,dtdeb, dtfin); // calcul montant cumul plus proratisé

    senscalculsal := '-';
    calculcumul(listecummoinPR,senscalculsal, calculperiode,proratis, dtdeb, dtfin);  // calcul montant cumul moins proratisé
    calculbaserem(listebasremmoinPR, senscalculsal, calculperiode,proratis,dtdeb, dtfin); // calcul base rem moins proratisé
    calculmtrem(listemtremmoinPR,senscalculsal,calculperiode,proratis,dtdeb,dtfin); // calcul montant rem moins proratisé

    // calcul du paramétraga "Non proratisé"
    senscalculsal := '+';
    proratis := false;
    calculcumul(listecumplusNPR,senscalculsal, calculperiode, proratis, dtdeb, dtfin); // calcul montant cumul plus non proratisé
    calculbaserem(listebasremplusNPR,senscalculsal, calculperiode,proratis, dtdeb, dtfin); // calcul base rem plus non proratisé
    calculmtrem(listemtremplusNPR, senscalculsal, calculperiode,proratis, dtdeb, dtfin); // calcul montant rem plus non proratisé
    
    senscalculsal := '-';
    calculcumul(listecummoinNPR,senscalculsal, calculperiode, proratis, dtdeb, dtfin); // calcul montant cumul moins non proratisé
    calculbaserem(listebasremmoinNPR,senscalculsal,calculperiode,proratis,dtdeb, dtfin); // calcul base rem moins non proratisé
    calculmtrem(listemtremmoinNPR,senscalculsal,calculperiode,proratis, dtdeb, dtfin);  // calcul montant rem moins non proratisé
    calculnbtrent(calculperiode,dtdeb,dtfin);

    
end;
// fin pt1

{***********A.G.L.********************************************************************
Auteur  ...... : NA
Créé le ...... : 19/05/2006
Modifié le ... :   /  /
Description .. : Calcul du total des cumuls pour un salarié (selon la liste passée en
Suite ........ : paramètre)
Mots clefs ... :
**************************************************************************************}
function calculcumulsal(salarie, listecum : string;  dtdeb, dtfin : Tdatetime): double;
var
st1 : string;
sqlcum : Tquery;

begin
result := 0;
if listecum <> '()' then
begin
st1 := 'SELECT sum(PHC_MONTANT) as CUM from HISTOCUMSAL where PHC_SALARIE = "'+salarie+'"'+
' and PHC_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and PHC_DATEFIN <= "'+USDATETIME(dtfin)+'" '+
' and  PHC_CUMULPAIE IN '+listecum+' group by PHC_SALARIE';
sqlcum := opensql(st1, true);
if not sqlcum.eof then
result := sqlcum.findfield('CUM').asfloat;
ferme(sqlcum);
end;
end;


{***********A.G.L.*************************************************************************
Auteur  ...... : NA
Créé le ...... : 19/05/2006
Modifié le ... :   /  /    
Description .. : Calcul du total des bases des  rubriques de rémunération  pour un salarié
Suite ........ : (selon la liste passée en paramètre)
Mots clefs ... : 
*******************************************************************************************}
function calculbaseremsal(salarie, listerem : string; dtdeb, dtfin : Tdatetime): double;
var
st1 : string;
sqlrem : Tquery;

begin
result := 0;
if listerem <> '()' then
begin
st1 := 'SELECT sum(PHB_BASEREM) as BASEREM from HISTOBULLETIN where PHB_SALARIE = "'+salarie+'"'+
' and PHB_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and PHB_DATEFIN <= "'+USDATETIME(dtfin)+'" '+
' and  PHB_RUBRIQUE IN '+listerem+' group by PHB_SALARIE';
sqlrem := opensql(st1, true);
if not sqlrem.eof then
result := sqlrem.findfield('BASEREM').asfloat;
ferme(sqlrem);
end;
end;


{***********A.G.L.**********************************************************
Auteur  ...... : NA
Créé le ...... : 19/05/2006
Modifié le ... :   /  /    
Description .. : Calcul du total des montants des rubriques de rémunération 
Suite ........ : (selon la liste passée en paramètre)
Mots clefs ... :
****************************************************************************}
function calculmtremsal(salarie, listerem : string; dtdeb, dtfin : Tdatetime): double;
var
st1 : string;
sqlrem : Tquery;

begin
result := 0;
if listerem <> '()' then
begin
st1 := 'SELECT sum(PHB_MTREM) as MTREM from HISTOBULLETIN where PHB_SALARIE = "'+salarie+'" '+
' and PHB_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and PHB_DATEFIN <= "'+USDATETIME(dtfin)+'" '+
' and  PHB_RUBRIQUE IN '+listerem+' group by PHB_SALARIE';
sqlrem := opensql(st1, true);
if not SqlRem.eof then
result := sqlrem.findfield('MTREM').asfloat;
Ferme(sqlrem);
end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/05/2006
Modifié le ... :   /  /    
Description .. : Calcul du nombre de trentième (pour un salarié)
Mots clefs ... : 
*****************************************************************}
function calculnbtrentsal(salarie : string; dtdeb, dtfin : Tdatetime): integer;
var
st1 : string;
sqltrent : Tquery;

begin
result := 0;
st1 := 'SELECT sum(PPU_NUMERATTRENT) as NBTRENT from PAIEENCOURS where PPU_SALARIE = "'+salarie+'"'+
' and PPU_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and PPU_DATEFIN <= "'+USDATETIME(dtfin)+'" '+
' group by PPU_SALARIE';
sqltrent := opensql(st1, true);
if not sqltrent.eof then
result := sqltrent.findfield('NBTRENT').asinteger;
ferme(sqltrent);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 31/05/2007
Modifié le ... :   /  /    
Description .. : Recherche des cumuls pour les salariés traités
Mots clefs ... : 
*****************************************************************}
procedure calculcumul(listecum,sens : string; calculperiode,proratis : boolean; dtdeb, dtfin : Tdatetime);
var
st1 , salarie: string;
sqlcum : Tquery;
result : double;
TCD : TOB;

begin

if listecum <> '()' then
begin
  st1 := 'SELECT PHC_SALARIE, sum(PHC_MONTANT) as CUM from HISTOCUMSAL LEFT JOIN SALARIES ON (PHC_SALARIE = PSA_SALARIE)'+
  '  '+lewhere+' and PHC_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and PHC_DATEFIN <= "'+USDATETIME(dtfin)+'" '+
  ' and  PHC_CUMULPAIE IN '+listecum+' group by PHC_SALARIE';
  sqlcum := opensql(st1, true);
  while not sqlcum.eof do
  begin
    result := sqlcum.findfield('CUM').asfloat;
    if proratis then result := result / 12;
    salarie := sqlcum.findfield('PHC_SALARIE').asstring;
    if calculperiode then TCD := Tob_calculsalaire.findfirst(['SALARIE'], [salarie], true)
    else TCD := Tob_calculsalairedernmois.findfirst(['SALARIE'], [salarie], true);
    if TCD <> nil then
    begin

      if sens = '+' then
      begin
        if TCD.fieldExists('CUMPLUS') then result := result + TCD.getvalue('CUMPLUS') else
        TCD.Addchampsup('CUMPLUS', false);
        TCD.putvalue('CUMPLUS',result );
      end
      else
      begin
        if TCD.fieldExists('CUMMOINS') then result := result + TCD.getvalue('CUMMOINS') else
        TCD.Addchampsup('CUMMOINS',false);
        TCD.putvalue('CUMMOINS',result );
      end;

    end
    else
    begin
      if calculperiode then TCD := Tob.create('Données caculées', TOB_calculsalaire, -1)
      else TCD := Tob.create('Données caculées', TOB_calculsalairedernmois, -1);

      TCD.Addchampsup('SALARIE', true);
      TCD.putvalue('SALARIE', salarie);
      if sens = '+' then
      begin
        TCD.Addchampsup('CUMPLUS',false);
        TCD.putvalue('CUMPLUS',result );
      end
      else
      begin
        TCD.Addchampsup('CUMMOINS',false);
        TCD.putvalue('CUMMOINS',result );
      end;
    end;
   sqlcum.next;
  end;
  ferme(sqlcum);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 31/05/2007
Modifié le ... :   /  /    
Description .. : Recherche des bases de cotisation pour tous les salariés
Mots clefs ... : 
*****************************************************************}
procedure calculbaserem(listerem,sens : string; calculperiode,proratis : boolean; dtdeb, dtfin : Tdatetime);
var
st1 , salarie: string;
sqlrem : Tquery;
result : double;
TCD : TOB;
begin

if listerem <> '()' then
begin
  st1 := 'SELECT PHB_SALARIE, sum(PHB_BASEREM) as BASEREM from HISTOBULLETIN LEFT JOIN SALARIES '+
  ' ON (PHB_SALARIE = PSA_SALARIE)   '+lewhere+' and PHB_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and' +
  ' PHB_DATEFIN <= "'+USDATETIME(dtfin)+'" and  PHB_RUBRIQUE IN '+listerem+' group by PHB_SALARIE';
  sqlrem := opensql(st1, true);
  while not sqlrem.eof do
  begin
    result := sqlrem.findfield('BASEREM').asfloat;
    if proratis then result := result /12;
    salarie := sqlrem.findfield('PHB_SALARIE').asstring;
    if calculperiode then TCD := Tob_calculsalaire.findfirst(['SALARIE'], [salarie], true)
    else TCD := Tob_calculsalairedernmois.findfirst(['SALARIE'], [salarie], true);
    if TCD <> nil then
    begin

      if sens = '+' then
      begin
        if TCD.fieldExists('BASEREMPLUS') then result := result + TCD.getvalue('BASEREMPLUS') else
        TCD.addchampsup('BASEREMPLUS', false);
        TCD.putvalue('BASEREMPLUS',result);
      end
      else
      begin
         if TCD.fieldExists('BASEREMMOINS') then result := result + TCD.getvalue('BASEREMMOINS') else
         TCD.addchampsup('BASEREMMOINS', false);
         TCD.putvalue('BASEREMMOINS',result);
      end;
    end
    else
    begin
      if calculperiode then TCD := Tob.create('Données caculées', TOB_calculsalaire, -1) else
      TCD := Tob.create('Données caculées', TOB_calculsalairedernmois, -1);

      TCD.addchampsup('SALARIE', true);
      TCD.putvalue('SALARIE', salarie);
      if sens = '+' then
      begin
        TCD.addchampsup('BASEREMPLUS', false);
        TCD.putvalue('BASEREMPLUS',result);
      end
      else
      begin
        TCD.addchampsup('BASEREMMOINS', false);
        TCD.putvalue('BASEREMMOINS',result);
      end;
    end;
   sqlrem.next;
  end;
  ferme(sqlrem);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 31/05/2007
Modifié le ... :   /  /    
Description .. : Recherche des montants de rémunération pour tous les 
Suite ........ : salariés
Mots clefs ... : 
*****************************************************************}
procedure calculmtrem(listerem,sens : string; calculperiode, proratis : boolean; dtdeb, dtfin : Tdatetime);
var
st1 , salarie : string;
sqlrem : Tquery;
result : double;
TCD : TOB;

begin

if listerem <> '()' then
begin
  st1 :=  'SELECT PHB_SALARIE, sum(PHB_MTREM) as MTREM from HISTOBULLETIN LEFT JOIN SALARIES '+
  ' ON (PHB_SALARIE = PSA_SALARIE)   '+lewhere+' and PHB_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and' +
  ' PHB_DATEFIN <= "'+USDATETIME(dtfin)+'" and  PHB_RUBRIQUE IN '+listerem+' group by PHB_SALARIE';
  sqlrem := opensql(st1, true);
  while not SqlRem.eof do
  begin
  result := sqlrem.findfield('MTREM').asfloat;
  if proratis then result := result /12;
  salarie := sqlrem.findfield('PHB_SALARIE').asstring;
    if  calculperiode then TCD := Tob_calculsalaire.findfirst(['SALARIE'], [salarie], true)
    else TCD := Tob_calculsalairedernmois.findfirst(['SALARIE'], [salarie], true);

    if TCD <> nil then
    begin

      if sens = '+' then
      begin
        if TCD.fieldExists('REMPLUS') then result := result + TCD.getvalue('REMPLUS') else
        TCD.addchampsup('REMPLUS', false);
        TCD.putvalue('REMPLUS',result);
      end
      else
      begin
        if TCD.fieldExists('REMMOINS') then result := result + TCD.getvalue('REMMOINS') else
        TCD.addchampsup('REMMOINS', false);
        TCD.putvalue('REMMOINS',result);
      end;
    end
    else
    begin
      if calculperiode then TCD := Tob.create('Données caculées', TOB_calculsalaire, -1)
      else TCD := Tob.create('Données caculées', TOB_calculsalairedernmois, -1);

      TCD.Addchampsup('SALARIE',true);
      TCD.putvalue('SALARIE', salarie);
      if sens = '+' then
      begin
        TCD.addchampsup('REMPLUS', false);
        TCD.putvalue('REMPLUS',result);
      end
      else
      begin
        TCD.addchampsup('REMMOINS', false);
        TCD.putvalue('REMMOINS',result);
      end;
    end;
   sqlrem.next;
  end;
  ferme(sqlrem);
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 31/05/2007
Modifié le ... :   /  /    
Description .. : Recherche nombre de trentiéme
Mots clefs ... : 
*****************************************************************}
procedure calculnbtrent(calculperiode : boolean; dtdeb, dtfin : Tdatetime);
var
st1 , salarie : string;
sqltrent : Tquery;
TCD : TOB;
result: integer;
begin

  st1 := 'SELECT PPU_SALARIE ,sum(PPU_NUMERATTRENT) as NBTRENT from PAIEENCOURS LEFT JOIN SALARIES ON '+
  '(PPU_SALARIE = PSA_SALARIE)  '+lewhere+' and PPU_DATEDEBUT >= "'+USDATETIME(dtdeb)+'" and '+
  'PPU_DATEFIN <= "'+USDATETIME(dtfin)+'" group by PPU_SALARIE';
  sqltrent := opensql(st1, true);
  while not sqltrent.eof do
  begin
    result := sqltrent.findfield('NBTRENT').asinteger;
    salarie := sqltrent.findfield('PPU_SALARIE').asstring;
    if  calculperiode then TCD := Tob_calculsalaire.findfirst(['SALARIE'], [salarie], true)
    else TCD := Tob_calculsalairedernmois.findfirst(['SALARIE'], [salarie], true);

    if TCD <> nil then
    begin
    TCD.Addchampsup('NBTRENT',false);
    TCD.Putvalue('NBTRENT', result);
    end;
    sqltrent.next;
  end;
  ferme(sqltrent);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 31/05/2007
Modifié le ... :   /  /    
Description .. : Calcul du salaire à partir de la tob contenant les résultats 
Suite ........ : des rubriques et cumuls pour tous les salariés
Mots clefs ... : 
*****************************************************************}
procedure calculsal(salcompare : boolean);
var
Tfille, Tfille_saldernmois: TOB;
mttot, cumplus,cummoins,baseremplus,baseremmoins, remplus,remmoins,nbtrent, salairemoyper, salairemoy, salairemoyderm: double;
dtdeb,dtfin : Tdatetime;
sal: string;
an : integer;

begin
// calcul du salaire des 3 derniers mois si demande la comparaison
    if salcompare then calculsaldernmois(datefin);

if tob_calculsalaire <> nil then
begin
  Tfille := tob_calculsalaire.findfirst([''],[''], true);
  while tfille <> nil do
  begin
   cumplus := 0;
   cummoins := 0;
   remplus := 0;
   remmoins :=0;
   baseremplus :=0;
   baseremmoins :=0;
   nbtrent := 0;

   sal := Tfille.getvalue('SALARIE');
   if tfille.fieldExists('CUMPLUS')   then  cumplus := Tfille.getvalue('CUMPLUS');
   if tfille.fieldExists('CUMMOINS')  then  cummoins := Tfille.getvalue('CUMMOINS');
   if tfille.fieldExists('REMPLUS')   then  remplus := Tfille.getvalue('REMPLUS');
   if tfille.fieldExists('REMMOINS')  then  remmoins := Tfille.getvalue('REMMOINS');
   if tfille.fieldExists('BASEREMPLUS')  then Baseremplus := Tfille.getvalue('BASEREMPLUS');
   if tfille.fieldExists('BASEREMMOINS') then Baseremmoins := Tfille.getvalue('BASEREMMOINS');
   if tfille.fieldExists('NBTRENT')   then Nbtrent := Tfille.getvalue('NBTRENT');
   mttot := cumplus - cummoins + baseremplus - baseremmoins + remplus - remmoins;
   // si le montant = 0 recherche du salaire pour les périodes A-1 , A-2, A-3
   if mttot = 0 then
   begin
     for an := 1 to 3 do
     begin
       dtdeb := plusdate(datedebut, -an, 'A');
       dtfin := plusdate(datefin, -an, 'A');
       cumplus :=  calculcumulsal(sal,listecumplus, dtdeb, dtfin);     // calcul montant cumul plus
       cummoins :=  calculcumulsal(sal, listecummoin, dtdeb, dtfin);     // calcul montant cumul moins
       baseremplus := calculbaseremsal(sal,listebasremplus, dtdeb, dtfin); // calcul base rem plus
       baseremmoins := calculbaseremsal(sal,listebasremmoin, dtdeb, dtfin); // calcul base rem moins
       remplus := calculmtremsal(sal,listemtremplus, dtdeb, dtfin); // calcul montant rem plus
       remmoins := calculmtremsal(sal,listemtremmoin, dtdeb, dtfin); // calcul montant rem moins
       mttot := cumplus - cummoins + baseremplus - baseremmoins + remplus - remmoins;
       if mttot <> 0 then
       begin
         nbtrent := calculnbtrentsal(sal, dtdeb, dtfin);
         break;
       end;
    end;
   end;
  if nbtrent <> 0 then salairemoyper := (mttot * 30) / nbtrent else salairemoyper := 0;
  salairemoyper := ARRONDI(salairemoyper, 2);
  if salairemoyper < 0 then salairemoyper := 0;
  // calcul du salaire des 3 derniers mois  si demande la comparaison
    if salcompare then
    begin
      // recherche des données calculées dans la tob "Calcul derniers mois TOB_CALCULSALAIREDERNMOIS)
      Tfille_saldernmois := tob_calculsalairedernmois.findfirst(['SALARIE'],[sal], true);
      if Tfille_saldernmois <> nil then
      begin
       // calcul du salaire moyen des 3 derniers mois
        if tfille_saldernmois.fieldExists('CUMPLUS')   then  cumplus := Tfille_saldernmois.getvalue('CUMPLUS');
        if tfille_saldernmois.fieldExists('CUMMOINS')  then  cummoins := Tfille_saldernmois.getvalue('CUMMOINS');
        if tfille_saldernmois.fieldExists('REMPLUS')   then  remplus := Tfille_saldernmois.getvalue('REMPLUS');
        if tfille_saldernmois.fieldExists('REMMOINS')  then  remmoins := Tfille_saldernmois.getvalue('REMMOINS');
        if tfille_saldernmois.fieldExists('BASEREMPLUS')  then Baseremplus := Tfille_saldernmois.getvalue('BASEREMPLUS');
        if tfille_saldernmois.fieldExists('BASEREMMOINS') then Baseremmoins := Tfille_saldernmois.getvalue('BASEREMMOINS');
        if tfille_saldernmois.fieldExists('NBTRENT')   then Nbtrent := Tfille_saldernmois.getvalue('NBTRENT');
        mttot := cumplus - cummoins + baseremplus - baseremmoins + remplus - remmoins;
        if mttot < 0 then mttot := 0;
        if nbtrent <> 0 then salairemoyderm := (mttot * 30) / nbtrent else salairemoyderm := 0;
        salairemoyderm:= ARRONDI(salairemoyderm, 2);
        // pt1 : le salaire retenu est le + avantageux entre le salaire de la période et le salaire des 3 derniers mois
        if salairemoyper > salairemoyderm then salairemoy := salairemoyper else salairemoy := salairemoyderm;
      end
      else
      begin
        salairemoyderm := 0;
        salairemoy := salairemoyper;
      end;
    end
    else
    begin // pas de comparaison, dans ce cas le salaire de la période et le salaire des 3 derniers mois =0
        salairemoyderm := 0;
        salairemoy := salairemoyper;
        salairemoyper := 0;
    end;
      Tfille.addchampsup('SALAIREMOY', true);
      Tfille.addchampsup('SALAIREMOYDERM', true);
      Tfille.addchampsup('SALAIREMOYPER', true);
      Tfille.putvalue('SALAIREMOY',salairemoy);
      Tfille.putvalue('SALAIREMOYDERM',salairemoyderm);
      Tfille.putvalue('SALAIREMOYPER',salairemoyper);

  Tfille := tob_calculsalaire.findnext([''],[''], true);
 end;
end;
if Assigned(Tob_calculsalairedernmois) then freeandnil(tob_calculsalairedernmois);
end;

end.
