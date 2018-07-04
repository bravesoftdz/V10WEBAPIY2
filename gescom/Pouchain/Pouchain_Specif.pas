unit Pouchain_Specif;

//---
// ATTENTION, ne pas changer l'ordre des champs sup dans la TOB
// à partir du champ Pv_Negocie
// car les lignes de totalisations sont renseignées à l'aide des numéros de champs
//---

interface

uses
  Classes,
  Hctrls,
  db,
  controls,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  Forms,
  sysutils,
  Windows,
  HEnt1,
  HMsgBox,
  UTOB,
  Entgc,
  uEntCommun,
  Ent1,
  AGLInit,
  AGLInitBTP,
  DateUtils,
  ED_TOOLS,
  Variants,
  utilxlsBTP,
  Ulog;

Type T_Valeurs = array [0..29] of double;

var TOBBTB             : TOB;
    DateDeb, DateFin   : TDateTime;
    Ordre_RA_UO        : Boolean;
    CodeRA, CodeUO, CodeAffaire, CodeClient, RepertServeur, RepertSortie, RepertLog, NomFicLog, LibAffPrinc, TypTableau : String;
    TxMarge : Double;
    T_Affaire, T_UO, T_RA, T_General : T_Valeurs;

Procedure LanceSpecifPouchain (NumeroMenu: integer; RepertoireServeur, RepertoireSortie, Annee : string);
Procedure LanceTdbcom07(Typ : string; RepertoireServeur, RepertoireSortie, Annee : string; Direct : Boolean=False);

implementation

Uses UTOF_VideInside;

procedure InitTableau (var resul : T_valeurs);
var indice : Integer;
begin
   for Indice := 0 to length(resul)-1 do
   begin
     Resul[Indice] := valeur('0');
   end;
end;

Procedure ChargeLeNegocie (TOBL : tob);
var Req         : string;
    Q          : Tquery;
begin
  // PV Négocié
  TOBL.AddChampSupValeur ('Pv_Negocie', 0);
  Req := 'select sum (GP_TOTALHTDEV)'+
         ' from PIECE WHERE GP_AFFAIRE="'+TOBL.GetValue('AFF_AFFAIRE')+'"'+
         ' and (select aff_etataffaire from affaire where aff_affaire = gp_affairedevis) <> "ENC"'+
         ' AND GP_LIBREPIECE1 = "X" AND gp_naturepieceg = "DBT"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Pv_Negocie',Q.Fields[0].AsFloat);
  end;
  ferme(Q);
end;

Procedure ChargeLePrevisionnel (TOBL : tob);
var Req         : string;
    Q          : Tquery;
begin
  // MOD prévisionnelle
  TOBL.AddChampSupValeur ('Prevu_MOD_H', 0);
//  Req := 'SELECT SUM(QTEMOAT+QTEMOCH+QTEMOENC+QTEMOETU+QTEMOREA+QTEMODAO+QTEMOAPI) FROM Z_TDB_PREVUOUV WHERE BOP_AFFAIRE1+BOP_AFFAIRE2+BOP_AFFAIRE3="'+
//         TOBL.GetValue('AFF_AFFAIRE1')+TOBL.GetValue('AFF_AFFAIRE2')+TOBL.GetValue('AFF_AFFAIRE3')+'"';
  Req := 'SELECT SUM(QTEMOAT+QTEMOCH+QTEMOENC+QTEMOETU+QTEMOREA+QTEMODAO+QTEMOAPI) FROM Z_TDB_PREVUOUV WHERE BOP_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_MOD_H',Q.Fields[0].AsFloat);
  end;
  ferme(Q);
  Req := 'SELECT SUM(QTEMOAT+QTEMOCH+QTEMOENC+QTEMOETU+QTEMOREA+QTEMODAO+QTEMOAPI) FROM Z_TDB_PREVUDEV WHERE GL_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_MOD_H',TOBL.GetValue('Prevu_MOD_H')+Q.Fields[0].AsFloat);
  end;
  ferme(Q);

  // Achats directs prévisionnels en PA
  TOBL.AddChampSupValeur ('Prevu_Achats_PA', 0);
  Req := 'SELECT SUM(MTCHPA+MTATPA+MTLOCPA+MTSTIPA+MTSTEPA+MTFDPA) FROM Z_TDB_PREVUOUV WHERE BOP_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_Achats_PA',Q.Fields[0].AsFloat);
  end;
  ferme(Q);
  Req := 'SELECT SUM(MTCHPA+MTATPA+MTLOCPA+MTSTIPA+MTSTEPA+MTFDPA) FROM Z_TDB_PREVUDEV WHERE GL_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_Achats_PA',TOBL.GetValue('Prevu_Achats_PA')+Q.Fields[0].AsFloat);
  end;
  ferme(Q);

  // Déboursé total prévisionnels en PA
  TOBL.AddChampSupValeur ('Prevu_Debourse_PA', 0);
  Req := 'SELECT SUM(MTCHPA+MTATPA+MTLOCPA+MTSTIPA+MTSTEPA+MTFDPA+MTMOATPA+MTMOCHPA+MTMOENCPA+MTMOETUPA+MTMOREAPA+MTMODAOPA+MTMOAPIPA+MTMOERAPA+MTMOERCPA) FROM Z_TDB_PREVUOUV WHERE BOP_AFFAIRE="'+
//  Req := 'SELECT SUM(MTCHPA+MTATPA+MTLOCPA+MTSTIPA+MTSTEPA+MTFDPA+MTMOATPA+MTMOCHPA+MTMOENCPA+MTMOETUPA+MTMOREAPA+MTMODAOPA+MTMOAPIPA) FROM Z_TDB_PREVUOUV WHERE BOP_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_Debourse_PA',Q.Fields[0].AsFloat);
  end;
  ferme(Q);
  Req := 'SELECT SUM(MTCHPA+MTATPA+MTLOCPA+MTSTIPA+MTSTEPA+MTFDPA+MTMOATPA+MTMOCHPA+MTMOENCPA+MTMOETUPA+MTMOREAPA+MTMODAOPA+MTMOAPIPA+MTMOERAPA+MTMOERCPA) FROM Z_TDB_PREVUDEV WHERE GL_AFFAIRE="'+
//  Req := 'SELECT SUM(MTCHPA+MTATPA+MTLOCPA+MTSTIPA+MTSTEPA+MTFDPA+MTMOATPA+MTMOCHPA+MTMOENCPA+MTMOETUPA+MTMOREAPA+MTMODAOPA+MTMOAPIPA) FROM Z_TDB_PREVUDEV WHERE GL_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_Debourse_PA',TOBL.GetValue('Prevu_Debourse_PA')+Q.Fields[0].AsFloat);
  end;
  ferme(Q);

  // Déboursé total prévisionnels en PR
  TOBL.AddChampSupValeur ('Prevu_Debourse_PR', 0);
  Req := 'SELECT SUM(MTCHPR+MTATPR+MTLOCPR+MTSTIPR+MTSTEPR+MTFDPR+MTMOATPR+MTMOCHPR+MTMOENCPR+MTMOETUPR+MTMOREAPR+MTMODAOPR+MTMOAPIPR+MTMOERAPR+MTMOERCPR) FROM Z_TDB_PREVUOUV WHERE BOP_AFFAIRE="'+
//  Req := 'SELECT SUM(MTCHPR+MTATPR+MTLOCPR+MTSTIPR+MTSTEPR+MTFDPR+MTMOATPR+MTMOCHPR+MTMOENCPR+MTMOETUPR+MTMOREAPR+MTMODAOPR+MTMOAPIPR) FROM Z_TDB_PREVUOUV WHERE BOP_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_Debourse_PR',Q.Fields[0].AsFloat);
  end;
  ferme(Q);
  Req := 'SELECT SUM(MTCHPR+MTATPR+MTLOCPR+MTSTIPR+MTSTEPR+MTFDPR+MTMOATPR+MTMOCHPR+MTMOENCPR+MTMOETUPR+MTMOREAPR+MTMODAOPR+MTMOAPIPR+MTMOERAPR+MTMOERCPR) FROM Z_TDB_PREVUDEV WHERE GL_AFFAIRE="'+
//  Req := 'SELECT SUM(MTCHPR+MTATPR+MTLOCPR+MTSTIPR+MTSTEPR+MTFDPR+MTMOATPR+MTMOCHPR+MTMOENCPR+MTMOETUPR+MTMOREAPR+MTMODAOPR+MTMOAPIPR) FROM Z_TDB_PREVUDEV WHERE GL_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Prevu_Debourse_PR',TOBL.GetValue('Prevu_Debourse_PR')+Q.Fields[0].AsFloat);
  end;
  ferme(Q);

end;

Procedure ChargeLeRealise (TOBL : tob);
var Req         : string;
    Q          : Tquery;
begin
  // Réalisé MO
  TOBL.AddChampSupValeur ('Realise_MOD_STA', 0);
  TOBL.AddChampSupValeur ('Realise_MOD_BET', 0);
  TOBL.AddChampSupValeur ('Realise_MOD_API', 0);
  TOBL.AddChampSupValeur ('Realise_MOD_TOTAL', 0);
  TOBL.AddChampSupValeur ('Realise_MOD_PA', 0);
  TOBL.AddChampSupValeur ('Realise_MOD_PR', 0);
  Req := 'select sum (REAMOATQTE+REAMOCHQTE+REAMOENCQTE),sum (REAMOETUQTE+REAMOREAQTE+REAMODAOQTE),sum (REAMOAPIQTE)'+
         ' ,sum (REAMOATPA+REAMOCHPA+REAMOENCPA+REAMOETUPA+REAMOREAPA+REAMODAOPA+REAMOAPIPA+REAMOERAPA+REAMOERCPA)'+
         ' ,sum (REAMOATPR+REAMOCHPR+REAMOENCPR+REAMOETUPR+REAMOREAPR+REAMODAOPR+REAMOAPIPR+REAMOERAPR+REAMOERCPR)'+
         ' ,sum (REAMOERAQTE+REAMOERCQTE)'+
//         ' ,sum (REAMOATPA+REAMOCHPA+REAMOENCPA+REAMOETUPA+REAMOREAPA+REAMODAOPA+REAMOAPIPA)'+
//         ' ,sum (REAMOATPR+REAMOCHPR+REAMOENCPR+REAMOETUPR+REAMOREAPR+REAMODAOPR+REAMOAPIPR)'+
         ' from Z_TDB_REALISE WHERE BCO_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+
         '" AND BCO_DATEMOUV <= "'+USDATETIME(DateFin)+'"';

  // Pour les tdbcom07 c et d, on ne prend que le réalisé de la période sélectionnée
  if (TypTableau = 'c') or (TypTableau = 'd') then
     Req := req + ' AND BCO_DATEMOUV >= "'+USDATETIME(DateDeb)+'"';

  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Realise_MOD_STA',Q.Fields[0].AsFloat);
    TOBL.Putvalue('Realise_MOD_BET',Q.Fields[1].AsFloat);
    TOBL.Putvalue('Realise_MOD_API',Q.Fields[2].AsFloat);
    TOBL.Putvalue('Realise_MOD_TOTAL',Q.Fields[0].AsFloat+Q.Fields[1].AsFloat+Q.Fields[2].AsFloat+Q.Fields[5].AsFloat);
    TOBL.Putvalue('Realise_MOD_PA',Q.Fields[3].AsFloat);
    TOBL.Putvalue('Realise_MOD_PR',Q.Fields[4].AsFloat);
  end;
  ferme(Q);
  // Déboursé Achats directs
  TOBL.AddChampSupValeur ('Realise_MAT_PA', 0);
  TOBL.AddChampSupValeur ('Realise_LOC_PA', 0);
  TOBL.AddChampSupValeur ('Realise_ST_PA', 0);
  TOBL.AddChampSupValeur ('Realise_FD_PA', 0);
  TOBL.AddChampSupValeur ('Realise_Achats_PA', 0);
  TOBL.AddChampSupValeur ('Realise_debourse_PA', 0);
  TOBL.AddChampSupValeur ('Realise_debourse_PR', 0);
  Req := 'select sum (DEBMATCHPA+DEBMATATPA), sum (DEBLOCMATPA), sum (DEBSTIPA+DEBSTEPA), sum (DEBFDPA)'+
         ' , sum (DEBMATCHPA+DEBMATATPA+DEBLOCMATPA+DEBSTIPA+DEBSTEPA+DEBFDPA)'+
         ' , sum (DEBMATCHPR+DEBMATATPR+DEBLOCMATPR+DEBSTIPR+DEBSTEPR+DEBFDPR)'+
         ' from Z_TDB_DEBOURSE WHERE BCO_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+
         '" AND BCO_DATEMOUV <= "'+USDATETIME(DateFin)+'"';

  // Pour les tdbcom07 c et d, on ne prend que le réalisé de la période sélectionnée
  if (TypTableau = 'c') or (TypTableau = 'd') then
     Req := req + ' AND BCO_DATEMOUV >= "'+USDATETIME(DateDeb)+'"';

  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Realise_MAT_PA',Q.Fields[0].AsFloat);
    TOBL.Putvalue('Realise_LOC_PA',Q.Fields[1].AsFloat);
    TOBL.Putvalue('Realise_ST_PA',Q.Fields[2].AsFloat);
    TOBL.Putvalue('Realise_FD_PA',Q.Fields[3].AsFloat);
    TOBL.Putvalue('Realise_Achats_PA',Q.Fields[0].AsFloat+Q.Fields[1].AsFloat+Q.Fields[2].AsFloat+Q.Fields[3].AsFloat);
    TOBL.Putvalue('Realise_debourse_PA',Q.Fields[4].AsFloat + TOBL.Getvalue('Realise_MOD_PA'));
    TOBL.Putvalue('Realise_debourse_PR',Q.Fields[5].AsFloat + TOBL.Getvalue('Realise_MOD_PR'));
  end;
  ferme(Q);

end;

Procedure ChargeLeFacture (TOBL : tob);
var Req        : string;
    Q          : Tquery;
begin
  // Facturé
  TOBL.AddChampSupValeur ('Facture', 0);
  Req := 'select sum (GP_TOTALHTDEV)'+
         ' from PIECE WHERE GP_AFFAIRE="'+TOBL.GetValue('AFF_AFFAIRE')+'"';

  if (TypTableau = 'd') then
    Req := Req + ' AND (gp_naturepieceg = "FAC" OR gp_naturepieceg = "AVC") AND GP_DATEPIECE <= "'+USDATETIME(DateFin)+'"'
  else
    Req := Req + ' AND (gp_naturepieceg = "FBT" OR gp_naturepieceg = "ABT") AND GP_DATEPIECE <= "'+USDATETIME(DateFin)+'"';

  // Pour les tdbcom07 c et d, on ne prend que le facturé de la période sélectionnée
  if (TypTableau = 'c') or (TypTableau = 'd') then
     Req := req + ' AND GP_DATEPIECE >= "'+USDATETIME(DateDeb)+'"';

  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Facture',Q.Fields[0].AsFloat);
  end;
  ferme(Q);
end;

procedure DateTransfoOK (TOBCDE : TOB; var Montant : Double);
var Q : Tquery;
    DateTransfo : TDateTime ;
begin
  Montant := 0;
  Q := OpenSQL('SELECT BCO_DATEMOUV, BCO_MONTANTACH FROM CONSOMMATIONS WHERE BCO_QUANTITE<>0 AND BCO_LIENTRANSFORME='+FloatToStr(TOBCDE.GetValue('BCO_NUMMOUV')), True);
  while Not Q.EOF do
  begin
  	DateTransfo := StrToDate(Q.FindField('BCO_DATEMOUV').AsString);

    if (TypTableau = 'c') or (TypTableau = 'd') then
    begin
      if not ((DateTransfo >= DateDeb) and (DateTransfo <= DateFin)) then
        Montant := Montant + Q.FindField('BCO_MONTANTACH').AsFloat;
    end else
    begin
      if not (DateTransfo <= DateFin) then
        Montant := Montant + Q.FindField('BCO_MONTANTACH').AsFloat;
    end;
    Q.Next;
  end;
	Ferme (Q);
end;

Procedure ChargeEngage (TOBL : tob);
var Req         : string;
    Q           : Tquery;
//    TOBCDE      : TOB;
//    Indice      : Integer;
    Montant     : Double;
begin
  // Engagé
  TOBL.AddChampSupValeur ('Engage', 0);

   // Calcul de l'engagé fournisseurs à partir des lignes d'achats
   // Cumul des commandes fournisseurs non livrées
   Req := 'SELECT SUM(GL_QTERESTE*GL_PUHTNETDEV) AS ACHAT FROM LIGNE ' +
          'WHERE GL_NATUREPIECEG IN ("CF","CFR","BFA") AND GL_TYPELIGNE="ART" ' +
          'AND GL_AFFAIRE="' + TOBL.GetValue('AFF_AFFAIRE') + '"';
   Req := Req + ' AND GL_DATEPIECE <= "' + UsDateTime(DateFin) + '" ';
   // Pour les tdbcom07 c et d, on ne prend que l'engagé de la période sélectionnée
   if (TypTableau = 'c') or (TypTableau = 'd') then
      Req := req + ' AND GL_DATEPIECE >= "'+USDATETIME(DateDeb)+'"';

   Q := OpenSQL(Req, true,-1,'',true);
   if not Q.EOF then
   begin
      Montant := Q.findfield('ACHAT').AsFloat;
      TOBL.PutValue('Engage', TOBL.GetValue('Engage') + Montant);
   end;
   ferme(Q);
   // Cumul des réceptions fournisseurs et factures issues de commandes et hors période (postérieures à date de fin) : donc à ajouter
   Req := 'SELECT SUM(GL_QTEFACT*GL_PUHTNETDEV) AS ACHAT FROM LIGNE ' +
          'WHERE GL_NATUREPIECEG IN ("BLF","LFR","FF") AND GL_TYPELIGNE="ART" ' +
          'AND GL_AFFAIRE="' + TOBL.GetValue('AFF_AFFAIRE') + '"';
   Req := Req + ' AND GL_PIECEPRECEDENTE LIKE "%;CF;%"';
   Req := Req + ' AND GL_DATEPIECE > "' + UsDateTime(DateFin) + '" ';
   Q := OpenSQL(Req, true,-1,'',true);
   if not Q.EOF then
   begin
      Montant := Q.findfield('ACHAT').AsFloat;
      TOBL.PutValue('Engage', TOBL.GetValue('Engage') + Montant);
   end;
   ferme(Q);

{
  Req := 'SELECT SUM(LIVMATCHAPA+LIVMATATPA+LIVLOCMATPA+LIVSTIPA+LIVSTEPA+LIVFDPA) FROM Z_TDB_CONSO WHERE BCO_AFFAIRE="'+
         TOBL.GetValue('AFF_AFFAIRE')+
         '" AND BCO_DATEMOUV <= "'+USDateTime(DateFin)+'"';

  // Pour les tdbcom07 c et d, on ne prend que l'engagé de la période sélectionnée
  if (TypTableau = 'c') or (TypTableau = 'd') then
     Req := req + ' AND BCO_DATEMOUV >= "'+USDATETIME(DateDeb)+'"';

  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    TOBL.Putvalue('Engage',Q.Fields[0].AsFloat);
  end;
  ferme(Q);

  // Gestion des commandes déjà livrées en fonction des dates de traitement
  Req := 'SELECT BCO_NUMMOUV FROM CONSOMMATIONS ' +
         'WHERE BCO_NATUREPIECEG IN ("CF","CFR","BFA") AND BCO_AFFAIRE="' + TOBL.GetValue('AFF_AFFAIRE') + '"';
  Req := Req + ' AND BCO_DATEMOUV <= "'+USDateTime(DateFin)+'"';
  Q := OpenSQL(Req, true,-1,'',true);
  TOBCDE := TOB.create('Commandes', nil, -1);
  TOBCDE.LoadDetailDB('', '', '', Q, true);
  ferme(Q);
  //Boucle sur tob pour traitement
  Montant := 0.0;
  for Indice := 0 TO TOBCDE.detail.count -1 do
  Begin
     DateTransfoOK(TOBCDE.Detail[Indice], Montant);
     if Montant <> 0 then
     begin
       TOBL.PutValue('Engage', TOBL.GetValue('Engage') + Montant);
     end;
  end;
  FreeAndNil(TOBCDE);
}

end;

Procedure ChargeAvancement (TOBL : tob);
var Req :string;
    Q : Tquery;
    DateDebutPrev, Dt, Dt2 : TDateTime;
    i : Integer;
begin
  // Avancement
  TOBL.AddChampSupValeur ('AvancementTotal', 0);
  TOBL.AddChampSupValeur ('AvancementPeriode', 0);
  TOBL.AddChampSupValeur ('CAMeriteTotal', 0);
  TOBL.AddChampSupValeur ('CAMeritePeriode', 0);

  // récupération date de début
  Req := 'SELECT ZPR_DATEDEBUT FROM ZPREVAFF WHERE ZPR_AFFAIRE ="'+TOBL.GetValue('AFF_AFFAIRE')+'"';
  Q := OpenSQL(Req, True);
  if Not Q.EOF then
  begin
    DateDebutPrev := Q.Fields[0].AsDateTime;
  end else
  begin
    DateDebutPrev := Date();
  end;
  ferme(Q);

  // On prend l'avancement si la date de début de prévisionnel est inférieure ou égale à la date de fin
  if DateDebutPrev <= DateFin then
  begin
    //détermination de l'avancement total
    // Calcul du numéro du champ correspondant au mois de la date de fin
    for i := 0 to 47 do
    begin
      Dt := PLUSMOIS(DateDebutPrev,i);
      if (YearOf(Dt) = YearOf(DateFin)) and (MonthOf(Dt) = MonthOf(DateFin)) then Break;
    end;
    // Préparation de la requête
    Req := 'SELECT ';
    if i = 0 then Req := Req + 'ZAV_POURCENTAGE1 '
    else if i = 1 then Req := Req + 'ZAV_POURCENTAGE2 '
    else if i = 2 then Req := Req + 'ZAV_POURCENTAGE3 '
    else if i = 3 then Req := Req + 'ZAV_POURCENTAGE4 '
    else if i = 4 then Req := Req + 'ZAV_POURCENTAGE5 '
    else if i = 5 then Req := Req + 'ZAV_POURCENTAGE6 '
    else if i = 6 then Req := Req + 'ZAV_POURCENTAGE7 '
    else if i = 7 then Req := Req + 'ZAV_POURCENTAGE8 '
    else if i = 8 then Req := Req + 'ZAV_POURCENTAGE9 '
    else if i = 9 then Req := Req + 'ZAV_POURCENTAGE10 '
    else if i = 10 then Req := Req + 'ZAV_POURCENTAGE11 '
    else if i = 11 then Req := Req + 'ZAV_POURCENTAGE12 '
    else if i = 12 then Req := Req + 'ZAV_POURCENTAGE13 '
    else if i = 13 then Req := Req + 'ZAV_POURCENTAGE14 '
    else if i = 14 then Req := Req + 'ZAV_POURCENTAGE15 '
    else if i = 15 then Req := Req + 'ZAV_POURCENTAGE16 '
    else if i = 16 then Req := Req + 'ZAV_POURCENTAGE17 '
    else if i = 17 then Req := Req + 'ZAV_POURCENTAGE18 '
    else if i = 18 then Req := Req + 'ZAV_POURCENTAGE19 '
    else if i = 19 then Req := Req + 'ZAV_POURCENTAGE20 '
    else if i = 20 then Req := Req + 'ZAV_POURCENTAGE21 '
    else if i = 21 then Req := Req + 'ZAV_POURCENTAGE22 '
    else if i = 22 then Req := Req + 'ZAV_POURCENTAGE23 '
    else if i = 23 then Req := Req + 'ZAV_POURCENTAGE24 '
    else if i = 24 then Req := Req + 'ZAV_POURCENTAGE25 '
    else if i = 25 then Req := Req + 'ZAV_POURCENTAGE26 '
    else if i = 26 then Req := Req + 'ZAV_POURCENTAGE27 '
    else if i = 27 then Req := Req + 'ZAV_POURCENTAGE28 '
    else if i = 28 then Req := Req + 'ZAV_POURCENTAGE29 '
    else if i = 29 then Req := Req + 'ZAV_POURCENTAGE30 '
    else if i = 30 then Req := Req + 'ZAV_POURCENTAGE31 '
    else if i = 31 then Req := Req + 'ZAV_POURCENTAGE32 '
    else if i = 32 then Req := Req + 'ZAV_POURCENTAGE33 '
    else if i = 33 then Req := Req + 'ZAV_POURCENTAGE34 '
    else if i = 34 then Req := Req + 'ZAV_POURCENTAGE35 '
    else if i = 35 then Req := Req + 'ZAV_POURCENTAGE36 '
    else if i = 36 then Req := Req + 'ZAV_POURCENTAGE37 '
    else if i = 37 then Req := Req + 'ZAV_POURCENTAGE38 '
    else if i = 38 then Req := Req + 'ZAV_POURCENTAGE39 '
    else if i = 39 then Req := Req + 'ZAV_POURCENTAGE40 '
    else if i = 40 then Req := Req + 'ZAV_POURCENTAGE41 '
    else if i = 41 then Req := Req + 'ZAV_POURCENTAGE42 '
    else if i = 42 then Req := Req + 'ZAV_POURCENTAGE43 '
    else if i = 43 then Req := Req + 'ZAV_POURCENTAGE44 '
    else if i = 44 then Req := Req + 'ZAV_POURCENTAGE45 '
    else if i = 45 then Req := Req + 'ZAV_POURCENTAGE46 '
    else if i = 46 then Req := Req + 'ZAV_POURCENTAGE47 '
    else Req := Req + 'ZAV_POURCENTAGE48 ';
    Req := Req + 'FROM ZAVANCEMENTAFF WHERE ZAV_AFFAIRE ="'+TOBL.GetValue('AFF_AFFAIRE')+'"';
    // Lecture avancement
    Q := OpenSQL(Req, True);
    if Not Q.EOF then
    begin
      TOBL.Putvalue('AvancementTotal',Q.Fields[0].AsFloat);
      TOBL.Putvalue('CAMeriteTotal',TOBL.Getvalue('pv_negocie')*(Q.Fields[0].AsFloat/100));
    end;
    ferme(Q);

    //détermination de l'avancement de la période
    // Calcul de la date correspondant au mois précédent la date de début
    Dt2 := PLUSMOIS(DateDeb,-1);
    // Calcul du numéro du champ correspondant
    for i := 0 to 47 do
    begin
      Dt := PLUSMOIS(DateDebutPrev,i);
      if (YearOf(Dt) = YearOf(Dt2)) and (MonthOf(Dt) = MonthOf(Dt2)) then Break;
    end;
    // Préparation de la requête
    Req := 'SELECT ';
    if i = 0 then Req := Req + 'ZAV_POURCENTAGE1 '
    else if i = 1 then Req := Req + 'ZAV_POURCENTAGE2 '
    else if i = 2 then Req := Req + 'ZAV_POURCENTAGE3 '
    else if i = 3 then Req := Req + 'ZAV_POURCENTAGE4 '
    else if i = 4 then Req := Req + 'ZAV_POURCENTAGE5 '
    else if i = 5 then Req := Req + 'ZAV_POURCENTAGE6 '
    else if i = 6 then Req := Req + 'ZAV_POURCENTAGE7 '
    else if i = 7 then Req := Req + 'ZAV_POURCENTAGE8 '
    else if i = 8 then Req := Req + 'ZAV_POURCENTAGE9 '
    else if i = 9 then Req := Req + 'ZAV_POURCENTAGE10 '
    else if i = 10 then Req := Req + 'ZAV_POURCENTAGE11 '
    else if i = 11 then Req := Req + 'ZAV_POURCENTAGE12 '
    else if i = 12 then Req := Req + 'ZAV_POURCENTAGE13 '
    else if i = 13 then Req := Req + 'ZAV_POURCENTAGE14 '
    else if i = 14 then Req := Req + 'ZAV_POURCENTAGE15 '
    else if i = 15 then Req := Req + 'ZAV_POURCENTAGE16 '
    else if i = 16 then Req := Req + 'ZAV_POURCENTAGE17 '
    else if i = 17 then Req := Req + 'ZAV_POURCENTAGE18 '
    else if i = 18 then Req := Req + 'ZAV_POURCENTAGE19 '
    else if i = 19 then Req := Req + 'ZAV_POURCENTAGE20 '
    else if i = 20 then Req := Req + 'ZAV_POURCENTAGE21 '
    else if i = 21 then Req := Req + 'ZAV_POURCENTAGE22 '
    else if i = 22 then Req := Req + 'ZAV_POURCENTAGE23 '
    else if i = 23 then Req := Req + 'ZAV_POURCENTAGE24 '
    else if i = 24 then Req := Req + 'ZAV_POURCENTAGE25 '
    else if i = 25 then Req := Req + 'ZAV_POURCENTAGE26 '
    else if i = 26 then Req := Req + 'ZAV_POURCENTAGE27 '
    else if i = 27 then Req := Req + 'ZAV_POURCENTAGE28 '
    else if i = 28 then Req := Req + 'ZAV_POURCENTAGE29 '
    else if i = 29 then Req := Req + 'ZAV_POURCENTAGE30 '
    else if i = 30 then Req := Req + 'ZAV_POURCENTAGE31 '
    else if i = 31 then Req := Req + 'ZAV_POURCENTAGE32 '
    else if i = 32 then Req := Req + 'ZAV_POURCENTAGE33 '
    else if i = 33 then Req := Req + 'ZAV_POURCENTAGE34 '
    else if i = 34 then Req := Req + 'ZAV_POURCENTAGE35 '
    else if i = 35 then Req := Req + 'ZAV_POURCENTAGE36 '
    else if i = 36 then Req := Req + 'ZAV_POURCENTAGE37 '
    else if i = 37 then Req := Req + 'ZAV_POURCENTAGE38 '
    else if i = 38 then Req := Req + 'ZAV_POURCENTAGE39 '
    else if i = 39 then Req := Req + 'ZAV_POURCENTAGE40 '
    else if i = 40 then Req := Req + 'ZAV_POURCENTAGE41 '
    else if i = 41 then Req := Req + 'ZAV_POURCENTAGE42 '
    else if i = 42 then Req := Req + 'ZAV_POURCENTAGE43 '
    else if i = 43 then Req := Req + 'ZAV_POURCENTAGE44 '
    else if i = 44 then Req := Req + 'ZAV_POURCENTAGE45 '
    else if i = 45 then Req := Req + 'ZAV_POURCENTAGE46 '
    else if i = 46 then Req := Req + 'ZAV_POURCENTAGE47 '
    else if i = 47 then Req := Req + 'ZAV_POURCENTAGE48 ';
    if (i = 48) then
    begin
      if Dt2 <= DateDebutPrev then  // cas où la date de début saisie est inférieure ou égale à la date de début du prévisionnel
      begin
        TOBL.Putvalue('AvancementPeriode',TOBL.GetValue('AvancementTotal'));
        TOBL.Putvalue('CAMeritePeriode',TOBL.Getvalue('pv_negocie')*(TOBL.GetValue('AvancementPeriode')/100));
      end else
      begin
        TOBL.Putvalue('AvancementPeriode',0.0);
        TOBL.Putvalue('CAMeritePeriode',0.0);
      end;
    end else
    begin
      Req := Req + 'FROM ZAVANCEMENTAFF WHERE ZAV_AFFAIRE ="'+TOBL.GetValue('AFF_AFFAIRE')+'"';
      // Lecture avancement
      Q := OpenSQL(Req, True);
      if Not Q.EOF then
      begin
        TOBL.Putvalue('AvancementPeriode',TOBL.GetValue('AvancementTotal')-Q.Fields[0].AsFloat);
        TOBL.Putvalue('CAMeritePeriode',TOBL.Getvalue('pv_negocie')*(TOBL.GetValue('AvancementPeriode')/100));
      end;
      ferme(Q);
    end;
  end;

  // Autres données calculées
  if (TypTableau = 'c') then
  begin
     TOBL.AddChampSupValeur ('Marge_nette', TOBL.getvalue('CAMeritePeriode')-TOBL.getvalue('Realise_debourse_PR'));
     TOBL.AddChampSupValeur ('Marge_brute', TOBL.getvalue('CAMeritePeriode')-TOBL.getvalue('Realise_debourse_PA'));
  end else if (TypTableau = 'd') then
  begin
     TOBL.AddChampSupValeur ('Marge_nette', TOBL.getvalue('Facture')-TOBL.getvalue('Realise_debourse_PR'));
     TOBL.AddChampSupValeur ('Marge_brute', TOBL.getvalue('Facture')-TOBL.getvalue('Realise_debourse_PA'));
  end else
  begin
     TOBL.AddChampSupValeur ('Marge_nette', TOBL.getvalue('CAmeriteTotal')-TOBL.getvalue('Realise_debourse_PR'));
     TOBL.AddChampSupValeur ('Marge_brute', TOBL.getvalue('CAmeriteTotal')-TOBL.getvalue('Realise_debourse_PA'));
  end;
  TOBL.AddChampSupValeur ('Encours_Facturation', TOBL.getvalue('Facture')-TOBL.getvalue('CAmeriteTotal'));

  if (TypTableau = 'b') then
  begin
    if TOBL.getvalue('Realise_debourse_PR') > TOBL.getvalue('pv_negocie') then
      TOBL.AddChampSupValeur ('Encours_Production', TOBL.getvalue('pv_negocie')-TOBL.getvalue('Facture'))
    else
      TOBL.AddChampSupValeur ('Encours_Production', TOBL.getvalue('Realise_debourse_PR')-TOBL.getvalue('Facture'));
  end else
  begin
    TOBL.AddChampSupValeur ('Encours_Production', TOBL.getvalue('CAmeriteTotal')-TOBL.getvalue('Facture'));
  end;

  TOBL.AddChampSupValeur ('MCD_Prevue', 0);
  TOBL.AddChampSupValeur ('MCD_Realise', 0);
  if (TypTableau = 'c') then
  begin
     if TOBL.getvalue('CAMeritePeriode') <> 0 then
        TOBL.PutValue ('MCD_Realise', TOBL.getvalue('Marge_brute')/TOBL.getvalue('CAMeritePeriode'));
  end else if (TypTableau = 'd') then
  begin
     if TOBL.getvalue('Facture') <> 0 then
        TOBL.PutValue ('MCD_Realise', TOBL.getvalue('Marge_brute')/TOBL.getvalue('Facture'));
  end else
  begin
     if TOBL.getvalue('CAmeriteTotal') <> 0 then
        TOBL.PutValue ('MCD_Realise', TOBL.getvalue('Marge_brute')/TOBL.getvalue('CAmeriteTotal'));
  end;
  if TOBL.getvalue('pv_negocie') <> 0 then
     TOBL.PutValue ('MCD_Prevue', (TOBL.getvalue('pv_negocie')-TOBL.getvalue('Prevu_Debourse_PA'))/TOBL.getvalue('pv_negocie'));
end;

Procedure ChargeLesAffaires;
var Req         : string;
    QQ          : Tquery;
begin

  if TypTableau = 'd' then
  begin
    Req := 'SELECT AFF_AFFAIRE, AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_RESPONSABLE AS AFF_LIBREAFFA, AFF_LIBELLE, AFF_TIERS, AFF_LIBREAFFA AS AFF_RESPONSABLE, AFF_CHANTIER, T_LIBELLE,' +
         ' AFF_ETATAFFAIRE as ETATAFFAIRE, "" AS STATUTPOUCHAIN' +
         ' FROM AFFAIRE left join TIERS on (T_Tiers = AFF_TIERS and T_NATUREAUXI="CLI") where'+
         ' not ((aff_datedebut > "'+USDATETIME(DateFin)+'") or (aff_datelibre2 <> "" and aff_datelibre2 < "'+USDATETIME(DateDeb)+'")) and aff_affaire0="W"';
  end else
  begin
    Req := 'SELECT AFF_AFFAIRE, AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_RESPONSABLE, AFF_LIBELLE, AFF_TIERS, AFF_CHANTIER, T_LIBELLE,' +
         ' (CASE WHEN (AFF_DATELIBRE2 > "'+USDATETIME(DateFin)+'" OR AFF_DATELIBRE2 = "") then ((CASE WHEN (ZPR_DATEDEBUT > "'+USDATETIME(DateFin)+'" OR ZPR_DATEDEBUT IS NULL) then ("ENC") else ("ACP") END)) else ("TER") END) as ETATAFFAIRE,' +
         ' (CASE WHEN (AFF_DATELIBRE2 > "'+USDATETIME(DateFin)+'") then ("") else (AFF_LIBREAFF1) END) as STATUTPOUCHAIN' +
         ' FROM AFFAIRE left join TIERS on (T_Tiers = AFF_TIERS and T_NATUREAUXI="CLI") left join ZPREVAFF on ZPR_AFFAIRE = AFF_AFFAIRE where'+
         ' not ((aff_datedebut > "'+USDATETIME(DateFin)+'") or (aff_datelibre2 <> "" and aff_datelibre2 < "'+USDATETIME(DateDeb)+'")) and aff_affaire0="A"';
  end;

  //Pour tdbcom07 et tdbcom07b, ne prendre que les affaires acceptées :
  if (TypTableau = '') or (TypTableau = 'b') then
    Req := Req + ' and (CASE WHEN (AFF_DATELIBRE2 > "'+USDATETIME(DateFin)+'" OR AFF_DATELIBRE2 = "") then ((CASE WHEN (ZPR_DATEDEBUT > "'+USDATETIME(DateFin)+'" OR ZPR_DATEDEBUT IS NULL) then ("ENC") else ("ACP") END)) else ("TER") END) = "ACP"';

  //Pour tdbcom07c , ne pas prendre l'affaire XXXFG :
  if (TypTableau = 'c') then
    Req := Req + ' and aff_affaire2 <> "XXXFG"';

  //Pour tdbcom07b , ne pas prendre les affaires commençant par SV :
  if (TypTableau = 'b') then
    Req := Req + ' and aff_affaire2 not like "SV%"';

  if CodeUO <> '' then Req := Req + ' and aff_affaire1 in ('+CodeUO+')';
  if CodeRA <> '' then Req := Req + ' and aff_Responsable in ('+CodeRA+')';
  if CodeAffaire <> '' then Req := Req + ' and aff_affaire2 ="'+CodeAffaire+'"';
  if CodeClient <> '' then Req := Req + ' and aff_tiers ="'+CodeClient+'"';
  if Ordre_RA_UO then Req := Req + ' ORDER BY AFF_RESPONSABLE, ETATAFFAIRE, AFF_AFFAIRE1,'
  else                Req := Req + ' ORDER BY AFF_AFFAIRE1, ETATAFFAIRE, AFF_RESPONSABLE,';
  Req := Req + ' AFF_AFFAIRE2, AFF_AFFAIRE3';


  QQ := OpenSQL(Req, true);
  TOBBTB.LoadDetailDB('', '', '', QQ, true);
  ferme(QQ);

end;

Procedure MajTotGeneral (TOBBTB : TOB; ind : Integer);
var       TOBL : TOB;
          Indice, iPv_Negocie : Integer;
begin
  TOBL := TOB.Create ('',TOBBTB,ind);
  TOBL.Dupliquer(TOBBTB.Detail[ind-1],False,True);

  iPv_Negocie := TOBL.GetNumChamp('Pv_Negocie');
  TOBL.PutValue('AFF_RESPONSABLE','Total Général');
  for Indice := 0 to length(T_General)-1 do
  begin
    TOBL.PutValeur(iPv_Negocie+Indice,T_General[Indice]);
  end;
  TOBL.PutValue('AFF_AFFAIRE1','');
  TOBL.PutValue('AFF_AFFAIRE2','');
  TOBL.PutValue('AFF_AFFAIRE3','');
  TOBL.PutValue('ETATAFFAIRE','');
  TOBL.PutValue('STATUTPOUCHAIN','');
  TOBL.PutValue('AFF_LIBELLE','');
  TOBL.PutValue('AFF_TIERS','');
  if (TypTableau = 'd') then TOBL.PutValue('AFF_LIBREAFFA','');
end;

Procedure MajTotRAUO (TOBBTB : TOB; ind : Integer; TypeTot, LibTot : String);
var       TOBL : TOB;
          Indice, iPv_Negocie : Integer;
begin
  TOBL := TOB.Create ('',TOBBTB,ind);
  TOBL.Dupliquer(TOBBTB.Detail[ind-1],False,True);

  iPv_Negocie := TOBL.GetNumChamp('Pv_Negocie');
  if TypeTot = 'RA' then
  begin
    TOBL.PutValue('AFF_RESPONSABLE','Total '+LibTot);
    TOBL.PutValue('AFF_AFFAIRE1','');
    for Indice := 0 to length(T_RA)-1 do
    begin
      TOBL.PutValeur(iPv_Negocie+Indice,T_RA[Indice]);
    end;
  end else
  begin
    TOBL.PutValue('AFF_AFFAIRE1','Total '+LibTot);
    TOBL.PutValue('AFF_RESPONSABLE','');
    for Indice := 0 to length(T_UO)-1 do
    begin
      TOBL.PutValeur(iPv_Negocie+Indice,T_UO[Indice]);
    end;
  end;

  TOBL.PutValue('AFF_AFFAIRE2','');
  TOBL.PutValue('AFF_AFFAIRE3','');
  TOBL.PutValue('ETATAFFAIRE','');
  TOBL.PutValue('STATUTPOUCHAIN','');
  TOBL.PutValue('AFF_LIBELLE','');
  TOBL.PutValue('AFF_TIERS','');
  if (TypTableau = 'd') then TOBL.PutValue('AFF_LIBREAFFA','');
end;

Procedure MajTotAff (TOBBTB : TOB; ind : Integer);
var       TOBL : TOB;
          Indice, iPv_Negocie : Integer;
begin
  TOBL := TOB.Create ('',TOBBTB,ind);
  TOBL.Dupliquer(TOBBTB.Detail[ind-1],False,True);

  TOBL.PutValue('AFF_AFFAIRE2','Total '+TOBL.Getvalue('AFF_AFFAIRE2'));
  TOBL.PutValue('AFF_AFFAIRE3','');
  // Affectation libellé affaire principale si elle existe
  if LibAffPrinc <> '' then
  begin
       TOBL.PutValue('AFF_LIBELLE',LibAffPrinc);
       LibAffPrinc := '';
  end;

 iPv_Negocie := TOBL.GetNumChamp('Pv_Negocie');
 for Indice := 0 to length(T_Affaire)-1 do
  begin
    TOBL.PutValeur(iPv_Negocie+Indice,T_Affaire[Indice]);
  end;

  for Indice := 0 to length(T_Affaire)-1 do
  begin
    T_UO[Indice] := T_UO[Indice] + T_Affaire[Indice];
    T_RA[Indice] := T_RA[Indice] + T_Affaire[Indice];
    T_General[Indice] := T_General[Indice] + T_Affaire[Indice];
  end;

  if T_UO[0] <> 0 then
  begin
       T_UO[20] := (T_UO[22]/T_UO[0])*100;
       T_UO[21] := (T_UO[23]/T_UO[0])*100;
       T_UO[28] := (T_UO[0] - T_UO[3])/T_UO[0];
  end;
  if T_RA[0] <> 0 then
  begin
       T_RA[20] := (T_RA[22]/T_RA[0])*100;
       T_RA[21] := (T_RA[23]/T_RA[0])*100;
       T_RA[28] := (T_RA[0] - T_RA[3])/T_RA[0];
  end;
  if T_General[0] <> 0 then
  begin
       T_General[20] := (T_General[22]/T_General[0])*100;
       T_General[21] := (T_General[23]/T_General[0])*100;
       T_General[28] := (T_General[0] - T_General[3])/T_General[0];
  end;

  if (TypTableau = 'c') then
  begin
     if T_UO[23] <> 0 then T_UO[29] := T_UO[25]/T_UO[23];
     if T_RA[23] <> 0 then T_RA[29] := T_RA[25]/T_RA[23];
     if T_General[23] <> 0 then T_General[29] := T_General[25]/T_General[23];
  end else if (TypTableau = 'd') then
  begin
     if T_UO[18] <> 0 then T_UO[29] := T_UO[25]/T_UO[18];
     if T_RA[18] <> 0 then T_RA[29] := T_RA[25]/T_RA[18];
     if T_General[18] <> 0 then T_General[29] := T_General[25]/T_General[18];
  end else
  begin
      if T_UO[22] <> 0 then T_UO[29] := T_UO[25]/T_UO[22];
      if T_RA[22] <> 0 then T_RA[29] := T_RA[25]/T_RA[22];
      if T_General[22] <> 0 then T_General[29] := T_General[25]/T_General[22];
  end;

  // modif BRL 5/11/2013 sur demande Mr Stienne dans com7b
  if (TypTableau = 'b') then
  begin
    if T_UO[17] > T_UO[0] then        // si Realise_debourse_PR > Pv_Negocie
      T_UO[27] := T_UO[0]-T_UO[18]    // Encours_Production = Pv_Negocie - Facture
    else                              // sinon
      T_UO[27] := T_UO[17]-T_UO[18];  // Encours_Production = Realise_debourse_PR - Facture
  end else
  begin
    T_UO[27] := T_UO[22]-T_UO[18];    // Encours_Production = CAMeriteTotal - Facture
  end;
  if (TypTableau = 'b') then
  begin
    if T_RA[17] > T_RA[0] then        // si Realise_debourse_PR > Pv_Negocie
      T_RA[27] := T_RA[0]-T_RA[18]    // Encours_Production = Pv_Negocie - Facture
    else                              // sinon
      T_RA[27] := T_RA[17]-T_RA[18];  // Encours_Production = Realise_debourse_PR - Facture
  end else
  begin
    T_RA[27] := T_RA[22]-T_RA[18];    // Encours_Production = CAMeriteTotal - Facture
  end;
  if (TypTableau = 'b') then
  begin
    if T_General[17] > T_General[0] then             // si Realise_debourse_PR > Pv_Negocie
      T_General[27] := T_General[0]-T_General[18]    // Encours_Production = Pv_Negocie - Facture
    else                                             // sinon
      T_General[27] := T_General[17]-T_General[18];  // Encours_Production = Realise_debourse_PR - Facture
  end else
  begin
    T_General[27] := T_General[22]-T_General[18];    // Encours_Production = CAMeriteTotal - Facture
  end;

end;

Procedure CalculeLesTotaux (TOBBTB : TOB);
Var i : Integer;
    RACourant, UOCourant, AffaireCourante : string;
    RALu, UOLu, AffaireLu : string;
begin
  RACourant := '';
  UOCourant := '';
  AffaireCourante := '';
  LibAffPrinc := '';

  InitTableau (T_GENERAL);
  InitTableau (T_RA);
  InitTableau (T_UO);
  InitTableau (T_Affaire);

  i := 0;
  while i < TOBBTB.Detail.Count do
  begin
    If RACourant = '' then RACourant := TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE');
    If UOCourant = '' then UOCourant := TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1');
    If AffaireCourante = '' then AffaireCourante := TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2');

    RALu := TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE');
    UOLu := TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1');
    AffaireLu := TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2');

    // traitement des ruptures
    If (RALu <> RACourant) Or (UOLu <> UOCourant) Or (AffaireLu <> AffaireCourante) then
    begin
      // Rupture sur affaire
      If (AffaireLu <> AffaireCourante) or (UOLu <> UOCourant) or (RALu <> RACourant) then
      begin
        MajTotAff(TOBBTB,i);
        Inc(i);
        AffaireCourante := TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2');
        InitTableau (T_Affaire);
      end;
      // Test sur l'ordre de présentation choisi
      if Ordre_RA_UO then
      begin
        // Rupture sur UO
        If (UOLu <> UOCourant) or (RALu <> RACourant) then
        begin
          MajTotRAUO(TOBBTB,i,'UO',UOCourant);
          Inc(i);
          UOCourant := TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1');
          InitTableau (T_UO);
        end;
        // Rupture sur RA
        If RALu <> RACourant then
        begin
          MajTotRAUO(TOBBTB,i,'RA',RACourant);
          Inc(i);
          RACourant := TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE');
          InitTableau (T_RA);
        end;
      end else
      begin
        // Rupture sur RA
        If (RALu <> RACourant) or (UOLu <> UOCourant) then
        begin
          MajTotRAUO(TOBBTB,i,'RA',RACourant);
          Inc(i);
          RACourant := TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE');
          InitTableau (T_RA);
        end;
        // Rupture sur UO
        If UOLu <> UOCourant then
        begin
          MajTotRAUO(TOBBTB,i,'UO',UOCourant);
          Inc(i);
          UOCourant := TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1');
          InitTableau (T_UO);
        end;
      end;

      // Stockage libellé affaire principale pour la ligne de totalisation
      if (TOBBTB.Detail[i].GetValue('AFF_AFFAIRE3') = '000') and (LibAffPrinc = '') then LibAffPrinc := TOBBTB.Detail[i].GetValue('AFF_LIBELLE');

    end else
    begin
      // Cumuls des lignes
      T_Affaire[0] := T_Affaire[0] + TOBBTB.Detail[i].GetValue('Pv_Negocie');
      T_Affaire[1] := T_Affaire[1] + TOBBTB.Detail[i].GetValue('Prevu_MOD_H');
      T_Affaire[2] := T_Affaire[2] + TOBBTB.Detail[i].GetValue('Prevu_Achats_PA');
      T_Affaire[3] := T_Affaire[3] + TOBBTB.Detail[i].GetValue('Prevu_Debourse_PA');
      T_Affaire[4] := T_Affaire[4] + TOBBTB.Detail[i].GetValue('Prevu_Debourse_PR');
      T_Affaire[5] := T_Affaire[5] + TOBBTB.Detail[i].GetValue('Realise_MOD_STA');
      T_Affaire[6] := T_Affaire[6] + TOBBTB.Detail[i].GetValue('Realise_MOD_BET');
      T_Affaire[7] := T_Affaire[7] + TOBBTB.Detail[i].GetValue('Realise_MOD_API');
      T_Affaire[8] := T_Affaire[8] + TOBBTB.Detail[i].GetValue('Realise_MOD_TOTAL');
      T_Affaire[9] := T_Affaire[9] + TOBBTB.Detail[i].GetValue('Realise_MOD_PA');
      T_Affaire[10] := T_Affaire[10] + TOBBTB.Detail[i].GetValue('Realise_MOD_PR');
      T_Affaire[11] := T_Affaire[11] + TOBBTB.Detail[i].GetValue('Realise_MAT_PA');
      T_Affaire[12] := T_Affaire[12] + TOBBTB.Detail[i].GetValue('Realise_LOC_PA');
      T_Affaire[13] := T_Affaire[13] + TOBBTB.Detail[i].GetValue('Realise_ST_PA');
      T_Affaire[14] := T_Affaire[14] + TOBBTB.Detail[i].GetValue('Realise_FD_PA');
      T_Affaire[15] := T_Affaire[15] + TOBBTB.Detail[i].GetValue('Realise_Achats_PA');
      T_Affaire[16] := T_Affaire[16] + TOBBTB.Detail[i].GetValue('Realise_debourse_PA');
      T_Affaire[17] := T_Affaire[17] + TOBBTB.Detail[i].GetValue('Realise_debourse_PR');
      T_Affaire[18] := T_Affaire[18] + TOBBTB.Detail[i].GetValue('Facture');
      T_Affaire[19] := T_Affaire[19] + TOBBTB.Detail[i].GetValue('Engage');
      T_Affaire[22] := T_Affaire[22] + TOBBTB.Detail[i].GetValue('CAMeriteTotal');
      T_Affaire[23] := T_Affaire[23] + TOBBTB.Detail[i].GetValue('CAMeritePeriode');
      T_Affaire[24] := T_Affaire[24] + TOBBTB.Detail[i].GetValue('Marge_nette');
      T_Affaire[25] := T_Affaire[25] + TOBBTB.Detail[i].GetValue('Marge_brute');
      T_Affaire[26] := T_Affaire[26] + TOBBTB.Detail[i].GetValue('Encours_Facturation');

      // modif BRL 5/11/2013 sur demande Mr Stienne dans com7b
      //      T_Affaire[27] := T_Affaire[27] + TOBBTB.Detail[i].GetValue('Encours_Production');
      if (TypTableau = 'b') then
      begin
        if T_Affaire[17] > T_Affaire[0] then             // si Realise_debourse_PR > Pv_Negocie
          T_Affaire[27] := T_Affaire[0]-T_Affaire[18]    // Encours_Production = Pv_Negocie - Facture
        else                                             // sinon
          T_Affaire[27] := T_Affaire[17]-T_Affaire[18];   // Encours_Production = Realise_debourse_PR - Facture
      end else
      begin
        T_Affaire[27] := T_Affaire[22]-T_Affaire[18];     // Encours_Production = CAMeriteTotal - Facture
      end;


      if T_Affaire[0] <> 0 then
      begin
           T_Affaire[20] := (T_Affaire[22]/T_Affaire[0])*100; // Avancement Total
           T_Affaire[21] := (T_Affaire[23]/T_Affaire[0])*100; // Avancement Période
           T_Affaire[28] := (T_Affaire[0] - T_Affaire[3])/T_Affaire[0];
      end;
      if (TypTableau = 'c') then
      begin
         if T_Affaire[23] <> 0 then T_Affaire[29] := T_Affaire[25]/T_Affaire[23];
      end else if (TypTableau = 'd') then
      begin
         if T_Affaire[18] <> 0 then T_Affaire[29] := T_Affaire[25]/T_Affaire[18];
      end else
      begin
          if T_Affaire[22] <> 0 then T_Affaire[29] := T_Affaire[25]/T_Affaire[22];
      end;
      Inc(i);
    end;
  end;

  // Dernière totalisations
  if i > 0 then
  begin
    MajTotAff(TOBBTB, i);
    if Ordre_RA_UO then
    begin
      MajTotRAUO(TOBBTB,i+1,'UO',UOCourant);
      MajTotRAUO(TOBBTB,i+2,'RA',RACourant);
    end else
    begin
      MajTotRAUO(TOBBTB,i+1,'RA',RACourant);
      MajTotRAUO(TOBBTB,i+2,'UO',UOCourant);
    end;
    MajTotGeneral(TOBBTB,i+3);
  end;

end;

Procedure LanceExportXls(TOBBTB : TOB; Direct : Boolean=False);
Var i, LigneXls, ColXls : Integer;
    iRow,Icol,NbCols : integer;
    fWinExcel   : OleVariant;
    fWorkBook   : Variant;
    fWorkSheet  : Variant;
    fnewInst    : boolean;
    fDocExcel : string;
    db1, db2 : Double;
begin
  EcritLog (RepertLog,'Ouverture Excel',NomFicLog);
  if not OpenExcel(true,fWinExcel,fNewInst) then
  begin
    if Not Direct then
      PgiError ('Excel n''est pas installé sur ce poste','')
    else
      EcritLog (RepertLog,'Excel n''est pas installé sur ce poste',NomFicLog);
    exit;
  end;

  EcritLog (RepertLog,'Ouverture Maquette',NomFicLog);
  fDocExcel := RepertServeur+'\TdbCom07'+Typtableau+'_Maquette.xlsx';
  if not fileExists (fDocExcel) then
  begin
    ExcelClose (fWinExcel);
    if Not Direct then
      PGIBox('La maquette '+fDocExcel+' est introuvable.','Traitement impossible')
    else
      EcritLog (RepertLog,'La maquette '+fDocExcel+' est introuvable.',NomFicLog);
    Exit;
  end;

  EcritLog (RepertLog,'Ouverture Workbook',NomFicLog);
  fWorkBook := OpenWorkBook (fDocExcel ,fWinExcel);
  fWorkSheet := fWorkBook.activeSheet;

  EcritLog (RepertLog,'Ecriture entête',NomFicLog);
 	ExcelFindRange (fWorkBook,'TdbCom07'+Typtableau,'DEBUTMOIS',iRow,iCol,NbCols);
	ExcelValue(fWorkSheet,iRow-1,iCol-1,DateDeb);
 	ExcelFindRange (fWorkBook,'TdbCom07'+Typtableau,'FINMOIS',iRow,iCol,NbCols);
	ExcelValue(fWorkSheet,iRow-1,iCol-1,DateFin);
 	ExcelFindRange (fWorkBook,'TdbCom07'+Typtableau,'DATEGENE',iRow,iCol,NbCols);
	ExcelValue(fWorkSheet,iRow-1,iCol-1,'Généré le : '+FormatDateTime('dd/mm/yyyy hh:mm',NowH));

 	ExcelFindRange (fWorkBook,'TdbCom07'+Typtableau,'RA',iRow,iCol,NbCols);
  LigneXls := iRow;
  ColXls := 0;

  EcritLog (RepertLog,'Début écriture lignes',NomFicLog);
  for i := 0 to TOBBTB.Detail.Count - 1 do
  begin
    if Not Direct then MoveCurProgressForm('');
    if Typtableau = 'a' then
    begin
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('ETATAFFAIRE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('STATUTPOUCHAIN')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_CHANTIER')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE3')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_LIBELLE')); Inc(ColXls);
      if (Copy(TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE'),1,5)<> 'Total') and (Copy(TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1'),1,5)<> 'Total') then
      begin
        ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('T_LIBELLE')); Inc(ColXls);
      end else
        Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Pv_Negocie')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_MOD_H')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Achats_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Debourse_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_STA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_API')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_BET')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_TOTAL')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MAT_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_LOC_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_ST_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_FD_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_Achats_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Facture')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AvancementTotal')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('CAMeriteTotal')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AvancementPeriode')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('CAMeritePeriode')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Engage')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Marge_nette')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Marge_brute')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Encours_Facturation')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('MCD_Prevue')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('MCD_Realise'));
    end else if Typtableau = '' then
    begin
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('ETATAFFAIRE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('STATUTPOUCHAIN')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_CHANTIER')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE3')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_LIBELLE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_TIERS')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Pv_Negocie')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_MOD_H')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Achats_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_TOTAL')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_Achats_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Facture')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AvancementTotal')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('CAMeriteTotal')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AvancementPeriode')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('CAMeritePeriode')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Engage')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Marge_nette')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Marge_brute')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Encours_Facturation')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('MCD_Prevue')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('MCD_Realise'));
    end else if Typtableau = 'b' then
    begin
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_CHANTIER')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE3')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_LIBELLE')); Inc(ColXls);
      if (Copy(TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE'),1,5)<> 'Total') and (Copy(TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1'),1,5)<> 'Total') then
      begin
        ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('T_LIBELLE')); Inc(ColXls);
      end else
        Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Pv_Negocie')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Facture')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Encours_Production'));
    end else if (TypTableau = 'c') or (TypTableau = 'd') then
    begin
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE')); Inc(ColXls);
      if (TypTableau = 'd') then
      begin
        ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_LIBREAFFA')); Inc(ColXls);
      end;
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('ETATAFFAIRE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('STATUTPOUCHAIN')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_CHANTIER')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE3')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_LIBELLE')); Inc(ColXls);
      if (TypTableau = 'd') then
      begin
        if (Copy(TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE'),1,5)<> 'Total') and (Copy(TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1'),1,5)<> 'Total') then
        begin
          ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('T_LIBELLE')); Inc(ColXls);
        end else
          Inc(ColXls);
      end;
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_STA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_API')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_BET')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_TOTAL')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MAT_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_LOC_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_ST_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_FD_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_Achats_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Facture')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AvancementPeriode')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('CAMeritePeriode')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Marge_nette')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Marge_brute')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('MCD_Realise'));
    end else if Typtableau = 'f' then
    begin
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_RESPONSABLE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('ETATAFFAIRE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('STATUTPOUCHAIN')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_CHANTIER')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE1')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE2')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_AFFAIRE3')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('AFF_LIBELLE')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Pv_Negocie')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_MOD_H')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Achats_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Debourse_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Prevu_Debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_STA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_API')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_BET')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MOD_TOTAL')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_MAT_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_LOC_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_ST_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_FD_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_Achats_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PR')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Realise_debourse_PA')); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TOBBTB.Detail[i].GetValue('Facture')); Inc(ColXls);

      ExcelValue(fWorkSheet,LigneXls+i,ColXls,TxMarge); Inc(ColXls);
      db1:=0.0;
      if TxMarge <> 0.0 then db1 := TOBBTB.Detail[i].GetValue('Realise_debourse_PA')/(1-TxMarge/100);
      Db2 := db1 - TOBBTB.Detail[i].GetValue('Facture');

      ExcelValue(fWorkSheet,LigneXls+i,ColXls,db1); Inc(ColXls);
      ExcelValue(fWorkSheet,LigneXls+i,ColXls,db2);

    end;
    ColXls := 0;
  end;
  EcritLog (RepertLog,'Sauvegarde Fichier Excel',NomFicLog);
  ExcelSave (fWorkBook ,RepertSortie+'\TdbCom07'+Typtableau+'-lse.xlsx');
  EcritLog (RepertLog,'Fermeture Excel',NomFicLog);
  ExcelClose (fWinExcel);

  if Not Direct then
  begin
    if Pgiask ('Document '+'TdbCom07'+Typtableau+'-lse.xlsx'+' généré avec succès.#13#10 #13#10 Désirez-vous le visualiser ?')= mryes then
    begin
      if not VarIsEmpty(fWinExcel) then
      begin
        fWinExcel.quit;
        fWinExcel := unassigned;
      end;
  //    FileExecAndWait (GetExcelPath + 'EXCEL.exe "'+RepertServeur+'\TdbCom07'+Typtableau+'-lse.xlsx' + '"');
      FileExec (GetExcelPath + 'EXCEL.exe "'+RepertSortie+'\TdbCom07'+Typtableau+'-lse.xlsx' + '"', False, False);
    end;
  end;
  EcritLog (RepertLog,'Fin export Excel',NomFicLog);

end;

Procedure LanceTdbcom07(Typ : string; RepertoireServeur, RepertoireSortie, Annee : string; Direct : Boolean=False);
Var i : Integer;
    TOBR : TOB;
    SurBureau : Boolean;
begin
  RepertLog := 'C:\ProgramData\Cegid';
  NomFicLog := 'Tdbcom07.log';

  EcritLog(RepertLog,'Lancement Tdbcom07 : '+FormatDateTime('dd/mm/yyyy hh:mm:ss',NowH),NomFicLog);

  // Saisie des paramètres de lancement :
  TOBR := TOB.Create ('TOB des parametres',nil,-1);
  TOBR.AddChampSupValeur ('DATEDEB',Debutannee(V_PGI.DateEntree));
  TOBR.AddChampSupValeur ('DATEFIN',Finannee(V_PGI.DateEntree));
  TOBR.AddChampSupValeur ('ORDRERAUO','X');
  TOBR.AddChampSupValeur ('CODERA','');
  TOBR.AddChampSupValeur ('CODEUO','');
  TOBR.AddChampSupValeur ('AFFAIRE','');
  TOBR.AddChampSupValeur ('CLIENT','');
  TOBR.AddChampSupValeur ('REPERT','');
  TOBR.AddChampSupValeur ('TYPTAB',Typ);
  TOBR.AddChampSupValeur ('TXMARGE',0.0);
  TOBR.AddChampSupValeur ('CODERETOUR','-');

  if Typ = 'x' then
    TypTableau := 'a'
  else if Typ = 'z' then
    TypTableau := ''
  else
    TypTableau := Typ;

  SurBureau := False;
  if (Direct) and (RepertoireSortie = '') then
  begin
    Direct := False;
    RepertoireSortie:=GetFromRegistry(HKEY_CURRENT_USER,'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders','Desktop',RepertoireSortie,TRUE) ;
    RepertSortie := RepertoireSortie;
    SurBureau := True;
  end;

  if Direct then
  begin
     RepertServeur:=RepertoireServeur;
     RepertSortie:=RepertoireSortie;
     if TypTableau = 'b' then
       DateFin := Now // type b  : à la date du jour
     else
       DateFin := FinDeMois(PlusDate(Now,-1,'M')); // Fin du mois précédent la date du jour
     DateDeb := Debutannee(DateFin);
     Ordre_RA_UO := True;
     CodeRA := '';
     CodeUO := '';
     CodeAffaire := '';
     CodeClient := '';
     TxMarge := 0.0;
     if Typ = 'z' then CodeAffaire := '22134';
     if Typ = 'x' then DateFin := Finannee(DateDeb);
     if Annee <> '' then
     begin
       DateDeb := StrToDate('01/01/'+Annee);
       DateFin := StrToDate('31/12/'+Annee);
     end;
  end else
  begin
    TheTOB := TOBR;
    if SurBureau then
       AGLLANCEFICHE ('Z','ZTDBCOM07','','','ACTION=MODIFICATION')
    else
       SaisirHorsInside('ZTDBCOM07','Z');
    TheTOB := nil;

    if TOBR.GetValue('CODERETOUR') = '-' Then Exit;
    if Typ = 'f' then TxMArge := TOBR.GetValue('TXMARGE');

       RepertServeur := TOBR.GetValue('REPERT');
    if RepertoireSortie = '' then RepertSortie := RepertServeur;

    DateDeb := TOBR.GetValue('DATEDEB');
    DateFin := TOBR.GetValue('DATEFIN');
    if TOBR.GetValue('ORDRERAUO') = 'X' then
       Ordre_RA_UO := True
    else
       Ordre_RA_UO := False;
    CodeRA := TOBR.GetValue('CODERA');
    CodeUO := TOBR.GetValue('CODEUO');
    CodeAffaire := TOBR.GetValue('AFFAIRE');
    CodeClient := TOBR.GetValue('CLIENT');
    TxMarge := TOBR.Getvalue('TXMARGE');

    //traitement des codes issus de multivalcombobox (valeurs séparées par des point-virgules)
    if CodeRA <> '' then
    begin
      // On remplace le ; par une , pour la requête SQL à venir
      CodeRA := StringReplace(CodeRA, ';', '","', [rfReplaceAll]);
      // On enlève le guillemet et la virgule en fin de zone pour la requête SQL à venir
      CodeRA := '"'+Copy(CodeRA,0,(Length(CodeRA)-2));
    end;
    if CodeUO <> '' then
    begin
      // On remplace le ; par une , pour la requête SQL à venir
      CodeUO := StringReplace(CodeUO, ';', '","', [rfReplaceAll]);
      // On enlève le guillemet et la virgule en fin de zone pour la requête SQL à venir
      CodeUO := '"'+Copy(CodeUO,0,(Length(CodeUO)-2));
    end;
  end;
//---
// ATTENTION, ne pas changer l'ordre des champs sup dans la TOB
// à partir du champ Pv_Negocie
// car les lignes de totalisations sont renseignées à l'aide des numéros de champs
//---
  TOBBTB := TOB.create('Ma Tob', nil, -1);

  EcritLog(RepertLog,'Chargement affaires',NomFicLog);
  ChargeLesAffaires;

  if TOBBTB.Detail.Count > 0 then
  begin
    if Not Direct then
       InitMoveProgressForm(Application.MainForm,Application.Title, 'Préparation du tableau TdbCom07'+TypTableau+' en cours, veuillez patienter...', TOBBTB.Detail.Count, False, true) ;

    for i := 0 to TOBBTB.Detail.Count - 1 do
    begin
         MoveCurProgressForm('');

  EcritLog(RepertLog,'Chargement négocié',NomFicLog);
         ChargeLeNegocie (TOBBTB.Detail[i]);
  EcritLog(RepertLog,'Chargement prévisionnel',NomFicLog);
         ChargeLePrevisionnel (TOBBTB.Detail[i]);
  EcritLog(RepertLog,'Chargement réalisé',NomFicLog);
         ChargeLeRealise (TOBBTB.Detail[i]);
  EcritLog(RepertLog,'Chargement facturé',NomFicLog);
         ChargeLeFacture (TOBBTB.Detail[i]);
  EcritLog(RepertLog,'Chargement engagé',NomFicLog);
         ChargeEngage (TOBBTB.Detail[i]);
  EcritLog(RepertLog,'Chargement avancement',NomFicLog);
         ChargeAvancement (TOBBTB.Detail[i]);
    end;
  EcritLog(RepertLog,'Calculs totaux',NomFicLog);
    CalculeLesTotaux(TOBBTB);

    if Not Direct then
       FiniMoveProgressForm ;

    if Not Direct then
    begin
       InitMoveProgressForm(Application.MainForm,Application.Title, 'TdbCom07'+TypTableau+' : Transfert EXCEL en cours, veuillez patienter...', TOBBTB.Detail.Count, False, true) ;
       MoveCurProgressForm('');
    end;
  EcritLog(RepertLog,'Export Excel',NomFicLog);
    LanceExportXls(TOBBTB, Direct);
    if Not Direct then
        FiniMoveProgressForm ;

  end else
  begin
    if Not Direct then
       PGIBox('La sélection ne renvoie aucun enregistrement','Traitement impossible');
  end;

  EcritLog(RepertLog,'Fin du traitement',NomFicLog);
  FreeAndNil(TOBR);
  FreeAndNil(TOBBTB);
end;

// Lancement pgm depuis fichier de commande
Procedure LanceSpecifPouchain (NumeroMenu: integer; RepertoireServeur, RepertoireSortie, Annee : string);
begin
  if NumeroMenu = 1 then LanceTdbcom07('', RepertoireServeur, RepertoireSortie, Annee, True)
  else if NumeroMenu = 2 then LanceTdbcom07('a', RepertoireServeur, RepertoireSortie, Annee, True)
  else if NumeroMenu = 3 then LanceTdbcom07('b', RepertoireServeur, RepertoireSortie, Annee, True)
  else if NumeroMenu = 4 then LanceTdbcom07('c', RepertoireServeur, RepertoireSortie, Annee, True)
  else if NumeroMenu = 5 then LanceTdbcom07('d', RepertoireServeur, RepertoireSortie, Annee, True)
  else if NumeroMenu = 6 then LanceTdbcom07('f', RepertoireServeur, RepertoireSortie, Annee, True)
  else if NumeroMenu = 98 then LanceTdbcom07('x', RepertoireServeur, RepertoireSortie, Annee, True)
  else if NumeroMenu = 99 then LanceTdbcom07('z', RepertoireServeur, RepertoireSortie, Annee, True);
end;

end.
