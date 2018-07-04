{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/03/2004
Modifié le ... :   /  /
Description .. : Unit de gestion du multi critère salariés calcul dotation provision CP
Mots clefs ... : PAIE
*****************************************************************
PT1  : 31/08/2004 : V_500 PH FQ 11556 Ergonomie
PT2  : 02/09/2004 : V_500 PH FQ 11575 Ergonomie
PT3  : 02/09/2004 : V_500 PH FQ 11575 tests sur existence de la provision pour afficher le salarié
PT4  : 22/12/2004 : V_600 PH FQ 11865 Erreur SQL sur la liste des provisions à comptabiliser
PT5  : 05/01/2005 : V_600 PH Suite FQ 11865 On ne voit que la liste des salariés ayant une provision calculée
PT6  : 08/06/2005 : V_600 PH Suite FQ 12371 Gestion des dotations dans le cas d'un seul salarié
PT7  : 22/06/2005 : V_600 PH FQ 11648 Initialisation de la date d'arrété à la date de fin de mois
PT8  : 29/06/2005 : V_600 PH FQ 12394 Traitement du motif type RTT compatibilité ascendante
PT9  : 10/05/2006 : V_650 PH Prise en compte du taux de charges mis dans les paramsoc
PT10   25/06/2007 : V_700 PH FQ 11556 Ergonomie Titre
PT11 : 09/07/2007 : V_72  FC FQ 13911 Ergonomie
}

unit UTOFPG_SalDotCp_mul;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLSERVER}
  UTOF, HPanel, ParamDat, Dialogs,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, FE_Main,
{$ELSE}
  eMul, MaineAgl,
{$ENDIF}
{$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox,
  UTOB, HTB97,
  ParamSoc,
  Windows;

{$IFNDEF EAGLSERVER}
type
  TOF_PG_SalDotCP_mul = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
  private
    StEtab, TypeAction: string;
    procedure ParamCalcul(Sender: TObject);
    procedure LanceCalculProv(Sender: TObject);
    procedure ExitEtab(Sender: TObject);
  end;
{$ENDIF}

function CalculProv(DATEARRET: TDateTime; Etab1, Etab2: string): Boolean;
function Process_CalculProv(DD: TDateTime; Etab1, Etab2: string): TOB;
implementation
uses
  Ent1,
  PGoutils,
  ED_Tools,
  PgCongesPayes,
  P5Def,
  EntPaie,
{$IFNDEF EAGLSERVER}
  uFicheJob,
{$ENDIF}
  HQry;

function SqlEtab(Prefixe, Etab1, Etab2: string): string;
begin
  if Etab1 <> '' then
  begin
    if Etab2 <> '' then result := ' AND ' + Prefixe + '_ETABLISSEMENT>="' + Etab1 + '" AND ' + Prefixe + '_ETABLISSEMENT <="' + Etab2 + '"'
    else Result := ' AND ' + Prefixe + '_ETABLISSEMENT="' + Etab1 + '"';
  end
  else Result := '';
end;

function Process_CalculProv(DD: TDateTime; Etab1, Etab2: string): TOB;
{$IFDEF EAGLSERVER}
var TobResult: TOB;
  LaDate: TDateTime;
  OkOk: Boolean;
{$ENDIF}
begin
  result := nil;
{$IFDEF EAGLSERVER}
  TOBResult := TOB.Create('PROCESSCALCULPROV', nil, -1);
  MessageCgi('Param = ' + DateToStr(DD) + '|' + Etab1 + '|' + Etab2);
  MessageCgi('apres laDate');
  okok := CalculProv(DD, Etab1, Etab2);
  if OkOk then TOBResult.AddChampSupValeur('OKTRAITE', TRUE)
  else TOBResult.AddChampSupValeur('OKTRAITE', FALSE);
  result := TOBResult;
{$ENDIF}
end;

function CalculProv(DATEARRET: TDateTime; Etab1, Etab2: string): Boolean;
var
  salarie, st, Sql, TypGest: string;
  StAcquis, StPris, ZDateArret: string;
  TobProvisionCp, TSalProvCP, T1, TSal, TS: TOB;
  TCum, TC, Tob_CalendrierSalarie, TOB_Etab, TEtab: TOB;
  DateDN1, DateFN1, DateDN, DD, DF: TDateTime;
  i, j, TypJ: Integer;
  OkTrouve: Boolean;
  Mt, Cpat, CBrut, BasProv, NbRTT, TauxJ, Coeff, Jtrav, Htrav: Double;
  Q: TQuery;
  StEtab: string;
  VV: Variant;
  Acq, Pris: Double;
begin
  MessageCgi('DEB CalculProv');
  DD := DebutDeMois(DATEARRET);
  DF := FINDeMois(DATEARRET);
{$IFNDEF EAGLSERVER}
  InitMoveProgressForm(nil, 'Calcul des provisions en cours ...', 'Veuillez patienter SVP ...', 4, FALSE, TRUE);
{$ENDIF}
  StEtab := SqlEtab('PSA', Etab1, etab2);
  // Calcul de la provision RTT
  StAcquis := '';
  StPris := '';
  St := VH_Paie.PgRttAcquis;
  while St <> '' do
    StAcquis := StAcquis + ' PHC_CUMULPAIE="' + ReadTokenSt(St) + '" OR';
  if StAcquis <> '' then StAcquis := ' AND (' + Copy(StAcquis, 1, Length(StAcquis) - 2) + ')';
  St := VH_Paie.PgRttPris;
  while St <> '' do
    StPris := StPris + ' PHC_CUMULPAIE="' + ReadTokenSt(St) + '" OR';
  if StPris <> '' then StPris := ' AND (' + Copy(StPris, 1, Length(StPris) - 2) + ')';
  ZDateArret := UsDateTime(DATEARRET);
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement des cumuls RTT en cours ...');
{$ENDIF}
  MessageCgi('Avt Tob Etab');
  TOB_Etab := Tob.Create('etablissements_de_paie', nil, -1);
  q := opensql('SELECT * FROM ETABCOMPL ORDER BY ETB_ETABLISSEMENT', True);
  if not q.eof then TOB_Etab.LoadDetailDb('ETABCOMPL', '', '', q, False);
  Ferme(q);

  Sql := 'SELECT PMA_JOURHEURE FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_TYPERTT = "X" OR PMA_TYPEABS="RTT"'; // PT8
  Q := OpenSql(sql, TRUE);
  if not Q.EOF then
    TypGest := Q.FindField('PMA_JOURHEURE').AsString;
  Ferme(Q);
  Sql := 'SELECT DISTINCT PSA_SALARIE,' +
    '(SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_DATEFIN<="' + ZDateArret + '" ' +
    'AND PHC_SALARIE=PSA_SALARIE ' + StAcquis + ') ACQUIS,' +
    '(SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_DATEFIN<="' + ZDateArret + '" ' +
    'AND PHC_SALARIE=PSA_SALARIE ' + StPris + ') PRIS FROM SALARIES' +
    ' WHERE (PSA_DATESORTIE <= "' + usdatetime(10) + '" OR PSA_DATESORTIE >= "' + usdatetime(DATEARRET) + '"' +
    ' OR PSA_DATESORTIE IS NULL) ' + StEtab +
    ' ORDER BY PSA_SALARIE ';
  MessageCgi('SQL Cum=' + Sql);
  Q := OpenSql(sql, TRUE);
  TCum := TOB.Create('Les cumuls', nil, -1);
  TCum.LoadDetailDB('SALARIES', '', '', Q, False);
  Ferme(Q);

  st := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,' +
    'PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_TAUXHORAIRE,ETB_NBJOUTRAV,ETB_STANDCALEND,PSA_CALENDRIER,PSA_STANDCALEND FROM SALARIES';
  st := st + ' LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT ';
  st := St + ' WHERE (PSA_DATESORTIE <= "' + usdatetime(10) + '" OR PSA_DATESORTIE >= "' + usdatetime(DATEARRET) + '"' +
    ' OR PSA_DATESORTIE IS NULL) ' + StEtab + ' ORDER BY PSA_SALARIE';
  MessageCgi('SQL Sal=' + Sql);
  Q := OpenSql(st, TRUE);
  TSal := TOB.Create('Les Salaries', nil, -1);
  TSal.LoadDetailDB('SALARIES', '', '', Q, False);
  Ferme(Q);
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Calcul des provisions CP en cours ...');
{$ENDIF}
  MessageCgi('Avt CalculeprovCp');
  TobProvisionCp := CalculeProvisionCp(StEtab, DateArret, DateDN1, DateFN1, DateDN);
  MessageCgi('Apres CalculeprovCp');
  TSalProvCP := TOB.Create('Mes provisions', nil, -1);
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Calcul des provisions CP terminée, calcul des charges patronales en cours...');
{$ENDIF}
  for i := 0 to TobProvisionCp.detail.count - 1 do
  begin
    Salarie := TobProvisionCp.detail[i].GetValue('SALARIE');
    BasProv := Valeur(TobProvisionCp.detail[i].GetValue('CPPROVISION'));
    T1 := TOB.Create('PROVCP', TSalProvCP, -1);
    T1.PutValue('PDC_SALARIE', Salarie);
    T1.PutValue('PDC_DATEARRET', DATEARRET);
    T1.PutValue('PDC_CPPROVISION', Arrondi(BasProv, 2));

    St := 'SELECT PPU_CBRUT, PPU_CCOUTPATRON, ETB_NBJOUTRAV,PPU_JOURSOUVRES,PPU_JOURSOUVRABLE, PPU_ETABLISSEMENT,PPU_TRAVAILN1,PPU_TRAVAILN2,PPU_TRAVAILN3,PPU_TRAVAILN4,' +
      'PPU_CODESTAT,PPU_LIBREPCMB1,PPU_LIBREPCMB2,PPU_LIBREPCMB3,PPU_LIBREPCMB4,PPU_DATEDEBUT,PPU_DATEFIN ' +
      'FROM PAIEENCOURS LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PPU_ETABLISSEMENT WHERE PPU_SALARIE="' + Salarie +
      '" AND PPU_DATEFIN >="' + UsDateTime(DateDN1) + '" AND PPU_DATEFIN <="' + UsDateTime(DATEARRET) + '" ORDER BY PPU_DATEFIN DESC';
    Q := OpenSQL(St, true);

    if not Q.EOF then
    begin
      Q.First;
      T1.PutValue('PDC_ETABLISSEMENT', Q.FindField('PPU_ETABLISSEMENT').AsString);
      CPat := Q.FindField('PPU_CCOUTPATRON').AsFloat;
      Cbrut := Q.FindField('PPU_CBRUT').AsFloat;
      TypJ := Q.FindField('ETB_NBJOUTRAV').AsInteger;
      if TypJ = 5 then JTrav := Q.FindField('PPU_JOURSOUVRES').AsFloat
      else if TypJ = 6 then JTrav := Q.FindField('PPU_JOURSOUVRABLE').AsFloat
      else JTrav := 21;
      if CBrut <> 0 then
      begin
        if (VH_PAIE.PGtauxCharges <> 0) then Coeff := VH_PAIE.PGtauxCharges // PT9
        else Coeff := CPat / CBrut;
        Mt := Arrondi((Coeff * BasProv), 2);
      end
      else
      begin
        if (VH_PAIE.PGtauxCharges <> 0) then Coeff := VH_PAIE.PGtauxCharges // PT9
        else Coeff := 0;
        Mt := 0;
      end;
      T1.PutValue('PDC_CPCHARGESPAT', Mt);
      T1.PutValue('PDC_TRAVAILN1', Q.FindField('PPU_TRAVAILN1').AsString);
      T1.PutValue('PDC_TRAVAILN2', Q.FindField('PPU_TRAVAILN2').AsString);
      T1.PutValue('PDC_TRAVAILN3', Q.FindField('PPU_TRAVAILN3').AsString);
      T1.PutValue('PDC_TRAVAILN4', Q.FindField('PPU_TRAVAILN4').AsString);
      T1.PutValue('PDC_CODESTAT', Q.FindField('PPU_CODESTAT').AsString);
      T1.PutValue('PDC_LIBREPCMB1', Q.FindField('PPU_LIBREPCMB1').AsString);
      T1.PutValue('PDC_LIBREPCMB2', Q.FindField('PPU_LIBREPCMB2').AsString);
      T1.PutValue('PDC_LIBREPCMB3', Q.FindField('PPU_LIBREPCMB3').AsString);
      T1.PutValue('PDC_LIBREPCMB4', Q.FindField('PPU_LIBREPCMB4').AsString);
      T1.PutValue('PDC_DATEDEBUT', Q.FindField('PPU_DATEDEBUT').AsDateTime);
      T1.PutValue('PDC_DATEFIN', Q.FindField('PPU_DATEFIN').AsDateTime);
      if JTrav <> 0 then TauxJ := Arrondi(CBrut / JTrav, 2)
      else TauxJ := 0;
    end
    else
    begin
      if (VH_PAIE.PGtauxCharges <> 0) then Coeff := VH_PAIE.PGtauxCharges // PT9
      else Coeff := 0.45;
      TS := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);
      if TS <> nil then
      begin
        T1.PutValue('PDC_ETABLISSEMENT', TS.GetValue('PSA_ETABLISSEMENT'));
        Mt := Arrondi((Coeff * BasProv), 2); // Pourcentage forfaitaire
        T1.PutValue('PDC_CPCHARGESPAT', Mt);
        T1.PutValue('PDC_TRAVAILN1', TS.GetValue('PSA_TRAVAILN1'));
        T1.PutValue('PDC_TRAVAILN2', TS.GetValue('PSA_TRAVAILN2'));
        T1.PutValue('PDC_TRAVAILN3', TS.GetValue('PSA_TRAVAILN3'));
        T1.PutValue('PDC_TRAVAILN4', TS.GetValue('PSA_TRAVAILN4'));
        T1.PutValue('PDC_CODESTAT', TS.GetValue('PSA_CODESTAT'));
        T1.PutValue('PDC_LIBREPCMB1', TS.GetValue('PSA_LIBREPCMB1'));
        T1.PutValue('PDC_LIBREPCMB2', TS.GetValue('PSA_LIBREPCMB2'));
        T1.PutValue('PDC_LIBREPCMB3', TS.GetValue('PSA_LIBREPCMB3'));
        T1.PutValue('PDC_LIBREPCMB4', TS.GetValue('PSA_LIBREPCMB4'));
        TauxJ := TS.GetValue('PSA_TAUXHORAIRE') * 7;
        T1.PutValue('PDC_DATEDEBUT', IDate1900);
        T1.PutValue('PDC_DATEFIN', IDate1900);
      end;
    end;
    FERME(Q);
    T1.PutValue('PDC_COEFFCHARGES', Coeff);
    T1.PutValue('PDC_TAUXJOUR', TauxJ);
    T1.PutValue('PDC_RTTCHARGESPAT', 0);
    T1.PutValue('PDC_NBRTT', 0);
    T1.PutValue('PDC_CPBASERELN1', Valeur(TobProvisionCp.detail[i].GetValue('CPBASERELN1')));
    T1.PutValue('PDC_CPMOISRELN1', Valeur(TobProvisionCp.detail[i].GetValue('CPMOISRELN1')));
    T1.PutValue('PDC_CPRESTRELN1', Valeur(TobProvisionCp.detail[i].GetValue('CPRESTRELN1')));
    T1.PutValue('PDC_CPACQUISRELN1', Valeur(TobProvisionCp.detail[i].GetValue('CPACQUISRELN1')));
    T1.PutValue('PDC_CPPRISRELN1', Valeur(TobProvisionCp.detail[i].GetValue('CPPRISRELN1')));
    T1.PutValue('PDC_CPPROVRELN1', Valeur(TobProvisionCp.detail[i].GetValue('CPPROVRELN1')));
    T1.PutValue('PDC_CPBASEN1', Valeur(TobProvisionCp.detail[i].GetValue('CPBASEN1')));
    T1.PutValue('PDC_CPMOISN1', Valeur(TobProvisionCp.detail[i].GetValue('CPMOISN1')));
    T1.PutValue('PDC_CPACQUISN1', Valeur(TobProvisionCp.detail[i].GetValue('CPACQUISN1')));
    T1.PutValue('PDC_CPPRISN1', Valeur(TobProvisionCp.detail[i].GetValue('CPPRISN1')));
    T1.PutValue('PDC_CPRESTN1', Valeur(TobProvisionCp.detail[i].GetValue('CPRESTN1')));
    T1.PutValue('PDC_CPPROVN1', Valeur(TobProvisionCp.detail[i].GetValue('CPPROVN1')));
    T1.PutValue('PDC_CPBASEN', Valeur(TobProvisionCp.detail[i].GetValue('CPBASEN')));
    T1.PutValue('PDC_CPMOISN', Valeur(TobProvisionCp.detail[i].GetValue('CPMOISN')));
    T1.PutValue('PDC_CPACQUISN', Valeur(TobProvisionCp.detail[i].GetValue('CPACQUISN')));
    T1.PutValue('PDC_CPPRISN', Valeur(TobProvisionCp.detail[i].GetValue('CPPRISN')));
    T1.PutValue('PDC_CPRESTN', Valeur(TobProvisionCp.detail[i].GetValue('CPRESTN')));
    T1.PutValue('PDC_CPPROVN', Valeur(TobProvisionCp.detail[i].GetValue('CPPROVN')));
  end;
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Calcul des provisions RTT en cours...');
{$ENDIF}
  for i := 0 to TCum.detail.count - 1 do
  begin
    Salarie := TCum.detail[i].GetValue('PSA_SALARIE');
    VV := TCum.detail[i].GetValue('ACQUIS');
    if not IsNumeric(VV) then Acq := 0
    else Acq := TCum.detail[i].GetDouble('ACQUIS');
    VV := TCum.detail[i].GetValue('PRIS');
    if not IsNumeric(VV) then Pris := 0
    else Pris := TCum.detail[i].GetDouble('PRIS');
    NbRTT := Acq - Pris;
// PT6    if NbRTT = 0 then continue;
    for j := 0 to TSalProvCP.detail.count - 1 do
    begin
      OkTrouve := FALSE;
      if Salarie = TSalProvCP.detail[j].GetValue('PDC_SALARIE') then
      begin
        OkTrouve := TRUE;
        Coeff := TSalProvCP.detail[j].GetValue('PDC_COEFFCHARGES');
        TauxJ := TSalProvCP.detail[j].GetValue('PDC_TAUXJOUR');
        break;
      end;
    end;
    if OkTrouve then
    begin
      if TypGest = 'JOU' then
      begin
        Mt := Arrondi((TauxJ * NbRTT), 2);
        TSalProvCP.detail[j].PutValue('PDC_PROVRTT', Mt);
        TSalProvCP.detail[j].PutValue('PDC_NBRTT', NbRTT);
        Mt := Arrondi((Mt * Coeff), 2);
        TSalProvCP.detail[j].PutValue('PDC_RTTCHARGESPAT', Mt);
      end
      else
      begin // Gestion en heures des RTT, on cherche donc à trouver le nombre d'heures travaillées par jour
        TS := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);
        if TS <> nil then
        begin
          if Assigned(Tob_CalendrierSalarie) then FreeAndNil(Tob_CalendrierSalarie);
          Tob_CalendrierSalarie := ChargeCalendrierSalarie(TS.getvalue('PSA_ETABLISSEMENT'),
            TS.getvalue('PSA_SALARIE'), TS.getvalue('PSA_CALENDRIER'),
            TS.getvalue('PSA_STANDCALEND'), TS.getvalue('ETB_STANDCALEND'));
          TEtab := TOB_Etab.FindFirst(['ETB_ETABLISSEMENT'], [TS.getvalue('PSA_ETABLISSEMENT')], FALSE);
          if TEtab <> nil then
          begin
            TypJ := TS.getvalue('ETB_NBJOUTRAV');
            HTrav := 0;
            if TypJ = 5 then CalculVarOuvresOuvrablesMois(TEtab, nil, Tob_CalendrierSalarie, DD, DF, TRUE, JTrav, HTrav)
            else if TypJ = 6 then CalculVarOuvresOuvrablesMois(TEtab, nil, Tob_CalendrierSalarie, DD, DF, FALSE, JTrav, HTrav);
            if Htrav > 0 then Mt := Arrondi(((TauxJ / HTrav) * NbRTT), 2)
            else Mt := Arrondi(((TauxJ / 7.0) * NbRTT), 2);
          end;
        end;
        TSalProvCP.detail[j].PutValue('PDC_PROVRTT', Mt);
        Mt := Arrondi((Mt * Coeff), 2);
        TSalProvCP.detail[j].PutValue('PDC_RTTCHARGESPAT', Mt);
        TSalProvCP.detail[j].PutValue('PDC_NBRTT', NbRTT);
      end
    end
    else
    begin // TOB non trouvée
      TS := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);
      if TS <> nil then
      begin
        T1 := TOB.Create('PROVCP', TSalProvCP, -1);
        T1.PutValue('PDC_SALARIE', Salarie);
        T1.PutValue('PDC_DATEARRET', DATEARRET);
        if (VH_PAIE.PGtauxCharges <> 0) then Coeff := VH_PAIE.PGtauxCharges // PT9
        else Coeff := 0.45;
        T1.PutValue('PDC_ETABLISSEMENT', TS.GetValue('PSA_ETABLISSEMENT'));
        T1.PutValue('PDC_CPCHARGESPAT', 0);
        T1.PutValue('PDC_CPPROVISION', 0);
        T1.PutValue('PDC_TRAVAILN1', TS.GetValue('PSA_TRAVAILN1'));
        T1.PutValue('PDC_TRAVAILN2', TS.GetValue('PSA_TRAVAILN2'));
        T1.PutValue('PDC_TRAVAILN3', TS.GetValue('PSA_TRAVAILN3'));
        T1.PutValue('PDC_TRAVAILN4', TS.GetValue('PSA_TRAVAILN4'));
        T1.PutValue('PDC_CODESTAT', TS.GetValue('PSA_CODESTAT'));
        T1.PutValue('PDC_LIBREPCMB1', TS.GetValue('PSA_LIBREPCMB1'));
        T1.PutValue('PDC_LIBREPCMB2', TS.GetValue('PSA_LIBREPCMB2'));
        T1.PutValue('PDC_LIBREPCMB3', TS.GetValue('PSA_LIBREPCMB3'));
        T1.PutValue('PDC_LIBREPCMB4', TS.GetValue('PSA_LIBREPCMB4'));
        TypJ := TS.GetValue('ETB_NBJOUTRAV'); // Recherche du nombre d'heures travaillées par jour


        TauxJ := (TS.GetValue('PSA_TAUXHORAIRE') * 7.2);
        T1.PutValue('PDC_DATEDEBUT', IDate1900);
        T1.PutValue('PDC_DATEFIN', IDate1900);
        if NbRTT <> 0 then Mt := Arrondi((TauxJ * NbRTT), 2)
        else Mt := 0;
        T1.PutValue('PDC_NBRTT', NbRTT);
        T1.PutValue('PDC_PROVRTT', Mt);
        Mt := Arrondi((Mt * Coeff), 2);
        T1.PutValue('PDC_RTTCHARGESPAT', Mt);
      end;
    end;
  end;
  result := FALSE;
  T1 := TSalProvCP.FindFirst([''], [''], FALSE);
  while T1 <> nil do
  begin // Boucle pour annuler les enregistrements non renseignés (significatifs)
    if (T1.GetValue('PDC_CPPROVISION') = 0) and (T1.GetValue('PDC_CPCHARGESPAT') = 0) and (T1.GetValue('PDC_PROVRTT') = 0) and
      (T1.GetValue('PDC_RTTCHARGESPAT') = 0) and (T1.GetValue('PDC_NBRTT') = 0) and
      (T1.GetValue('PDC_CPBASERELN1') = 0) and (T1.GetValue('PDC_CPMOISRELN1') = 0) and
      (T1.GetValue('PDC_CPRESTRELN1') = 0) and
      (T1.GetValue('PDC_CPBASEN1') = 0) and (T1.GetValue('PDC_CPMOISN1') = 0) and (T1.GetValue('PDC_CPACQUISN1') = 0) and
      (T1.GetValue('PDC_CPPRISN1') = 0) and (T1.GetValue('PDC_CPRESTN1') = 0) and (T1.GetValue('PDC_CPBASEN') = 0) and
      (T1.GetValue('PDC_CPMOISN') = 0) and (T1.GetValue('PDC_CPACQUISN') = 0) and (T1.GetValue('PDC_CPPRISN') = 0) and
      (T1.GetValue('PDC_CPRESTN') = 0) and (T1.GetValue('PDC_CPPROVN') = 0) and (T1.GetValue('PDC_CPPROVN1') = 0) and
      (T1.GetValue('PDC_CPPROVRELN1') = 0) and (T1.GetValue('PDC_CPPRISRELN1') = 0) and (T1.GetValue('PDC_CPACQUISRELN1') = 0)
      then T1.Free;
    T1 := TSalProvCP.FindNext([''], [''], FALSE);
  end;
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Fin de calcul des provisions RTT, Insertion des données en cours...');
{$ENDIF}
  if TSalProvCP.detail.count - 1 >= 0 then // PT6
  begin
    st := 'DELETE FROM PROVCP WHERE PDC_DATEARRET="' + UsDateTime(DATEARRET) + '"';
    if StEtab <> '' then
    begin
      st := st + StEtab;
      st := FindEtReplace(st, 'PSA', 'PDC', TRUE);
    end;
    ExecuteSQL(St);
    TSalProvCP.InsertDB(nil, FALSE);
    result := TRUE;
  end;
{$IFNDEF EAGLSERVER}
  FiniMoveProgressForm;
{$ENDIF}
  FreeAndNil(TSalProvCP);
  FreeAndNil(TobProvisionCp);
  FreeAndNil(TSal);
  FreeAndNil(TCum);
  if Assigned(Tob_CalendrierSalarie) then FreeAndNil(Tob_CalendrierSalarie);
  FreeAndNil(TOB_Etab);
{$IFNDEF EAGLSERVER}
  PGIBox('Calcul des provisions terminé !', 'Calcul des provisions CP et RTT salariés');
{$ENDIF}
end;

{$IFNDEF EAGLSERVER}

procedure TOF_PG_SalDotCP_mul.ParamCalcul(Sender: TObject);
var TobParam, T1: TOB;
  MaDate: THEdit;
  YY, MM, DD: WORD;
begin
  TobParam := TOB.create('Ma Tob de Param', nil, -1);
  T1 := TOB.Create('XXX', TobParam, -1);
  MaDate := THEdit(GetControl('_DATEFIN'));
  if MaDate <> nil then
  begin
    DecodeDate(StrToDate(GetControlTExt('_DATEFIN')), YY, MM, DD);
    T1.AddChampSupValeur('DD', IntToStr(DD));
    T1.AddChampSupValeur('MM', IntToStr(MM));
    T1.AddChampSupValeur('YY', IntToStr(YY));
    T1.AddChampSupValeur('ETAB1', GetControltext('PSA_ETABLISSEMENT'));
    T1.AddChampSupValeur('ETAB2', GetControltext('PSA_ETABLISSEMENT_'));
//  LancePRocess ....
    AGLFicheJob(0, taCreat, 'cgiPaieS5', 'PROVCP', TobParam);
  end;
  FreeAndNil(TOBParam);
end;

procedure TOF_PG_SalDotCP_mul.ExitEtab(Sender: TObject);
begin
  if GetControlText('PSA_ETABLISSEMENT_') = '' then SetControlText('PSA_ETABLISSEMENT_', GetControlText('PSA_ETABLISSEMENT'));
end;

procedure TOF_PG_SalDotCP_mul.LanceCalculProv(Sender: TObject);
var
  st, CbxMulti: string;
  rep: Integer;
  OkOk: Boolean;
  Q: TQuery;
  DATEARRET: TDateTime;
  StEtab: string;
  DD: THEdit; // PT7
begin
  CbxMulti := GetControlText('VCBXMULTI');
  // DEB PT7
  DD := THEdit(GetControl('DATEARRET'));
  if DD <> nil then DATEARRET := StrToDate(GetControlText('DATEARRET'))
  else DATEARRET := StrToDate(GetControlText('_DATEFIN'));
  // FIN PT7
  rep := MrYes;
  if TypeAction <> 'CALC' then
  begin
    if (CbxMulti = 'MON') and (GetControlText('PSA_ETABLISSEMENT') = '') then
    begin
      PGIBox('Vous devez saisir un établissement !', Ecran.Caption);
      rep := mrNo;
    end;
    if (CbxMulti = 'MUL') then
    begin
      if VH^.EtablisDefaut = '' then
      begin
        PGIBox('Vous devez obligatoirement renseigner un établissement principal dans la comptabilité', Ecran.Caption);
        rep := mrNo;
      end;
      // PT2 ergonomie dans contenu du message
      if rep = mrYes then rep := PGIAsk('Vous allez générer vos ODs pour tous vos établissements, établissement par établissement !', Ecran.Caption);
    end;
    if (CbxMulti = 'PRI') then rep := PGIAsk('Vous allez générer vos ODs pour tous vos établissements sur l''établissement principal !', Ecran.Caption);
  end;
  if rep = mrYes then
  begin
    StEtab := SqlEtab('PDC', GetControlText('PSA_ETABLISSEMENT'), GetControlText('PSA_ETABLISSEMENT_'));
    st := 'SELECT COUNT (*) NBRE FROM PROVCP WHERE PDC_DATEARRET="' + UsDateTime(DATEARRET) + '"';
    if StEtab <> '' then st := st + StEtab;
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    begin
      if Q.FindField('NBRE').AsInteger > 0 then // PT1  Faute orthographe
        rep := PGIAsk('Des provisions ont déjà été calculées à cette date, #13#10 Voulez vous recalculer les provisions à la date du ' + DateToStr(DATEARRET) + ' ?',
          Ecran.Caption)
      else rep := MrYes;
    end
    else rep := MrNo;
    Ferme(Q);
    if Rep = MrYes then OkOk := CalculProv(DATEARRET, GetControlText('PSA_ETABLISSEMENT'), GetControlText('PSA_ETABLISSEMENT_'))
    else OkOk := TRUE;
    if OkOk and (TypeAction <> 'CALC') then
    begin
      st := GetControlText('_DATEFIN'); // PT7
      St := St + ';';
      St := St + GetControlText('PSA_ETABLISSEMENT');
      St := St + ';' + GetControlText('PSA_ETABLISSEMENT_');
      St := St + ';' + CbxMulti;
      AglLanceFiche('PAY', 'CALCDOTCP', '', '', st);
    end
    else
      if TypeAction <> 'CALC' then PgiBox('Vous n''avez aucune provision CP à la date d''arrêté', Ecran.caption);
  end;
end;

procedure TOF_PG_SalDotCP_mul.OnArgument(stArgument: string);
var
  BtnValidMul, BTNPROCESS: TToolbarButton97;
  CbxMul: THValComboBox;
  CbxEtab: THValComboBox;
begin
  inherited;
  CbxMul := THValComboBox(GetControl('VCBXMULTI'));
  if CbxMul <> nil then CbxMul.Value := VH_Paie.PgVentilMulEtab;
  BtnValidMul := TToolbarButton97(GetControl('BTNLANCE'));
  if BtnValidMul <> nil then BtnValidMul.OnClick := LanceCalculProv;
  TypeAction := stArgument;
  if TypeAction = 'CALC' then
  begin
    Ecran.Caption := 'Calcul des provisions CP et RTT salariés';
    SetControlVisible('TMULTI', FALSE);
    SetControlVisible('VCBXMULTI', FALSE);
    BTNPROCESS := TToolbarButton97(GetControl('BPROG'));
    if BTNPROCESS <> nil then
    begin
      SetControlVisible('BPROG', FALSE);
      BTNPROCESS.OnClick := ParamCalcul;
    end;
  end
  else
  begin
    Ecran.Caption := 'Dotations pour provisions CP et RTT salariés';
    SetcontrolEnabled('CHBXAVEC', FALSE); //PT5
  end;
  UpdateCaption (Ecran); //  PT10
  CbxEtab := THValComboBox(GetControl('PSA_ETABLISSEMENT'));
  if CbxEtab <> nil then CbxEtab.OnExit := ExitEtab;
  if (VH_Paie.PGtauxCharges > 0) then
  begin
    SetControlText ('LBLTXCHARGES', 'Taux de charges '+ FloatToStr(VH_Paie.PGtauxCharges));
    SetControlVisible ('LBLTXCHARGES', TRUE);
  end
  else SetControlVisible ('LBLTXCHARGES', FALSE);
end;

procedure TOF_PG_SalDotCP_mul.OnLoad;
var
  stTemp, LePlus: string;
begin
  inherited;
  // PT3 Rajout de la clause exists
// DEB PT5
//  if TypeAction = 'CALC' then
//  begin
  if GetControlText('CHBXAVEC') = 'X' then
//    SetControlProperty('CHBXAVEC', 'Caption', 'Calcul des dotations effectuées'); //PT11
    LePlus := ''
  else
//    SetControlProperty('CHBXAVEC', 'Caption', 'Calcul des dotations non effectuées');  //PT11
    LePlus := ' NOT ';
//  end;  // FIN PT5
  // PT4  Rajout d'un espace devant EXIST
  StTemp := '(PSA_DATESORTIE <= "' + usdatetime(10) + '" OR PSA_DATESORTIE >= "' + usdatetime(StrToDate(GetControlText('_DATEFIN'))) + '"' + // PT7
    ' OR PSA_DATESORTIE IS NULL) AND' + LePlus + ' EXISTS (SELECT PDC_SALARIE FROM PROVCP WHERE PDC_SALARIE=PSA_SALARIE AND PDC_DATEARRET="' + usdatetime(StrToDate(GetControlText('_DATEFIN'))) + '")'; // PT7
  SetControlText('XX_WHERE', StTemp);
end;
{$ENDIF}

initialization
{$IFNDEF EAGLSERVER}
  registerclasses([TOF_PG_SalDotCP_mul]);
{$ENDIF}
end.

