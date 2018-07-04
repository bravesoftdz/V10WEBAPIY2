{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 18/03/2005
Modifié le ... :   /  /
Description .. : Fonctions @ publiques AGL
Mots clefs ... : PAIE;EDITION
*****************************************************************}
{
PT1   : 18/03/2004 SB V_60 FQ 11398 Ajout paramètre function fiche individuelle
PT2   : 20/09/2005 SB V_65 FQ 12372 Modif Order by tob des cumuls
PT3   : 18/01/2006 SB V_65 modification de la clause de Confidentialité
PT4   : 22/02/2007 FC V_72 Intégration des zones dynamiques sur le bulletin
                           Gestion de ces zones par une fonction générique
                           @PAIEGETVALZONE
PT5   : 09/07/2007 FC V_72 FQ 14535 Historisation population
PT6   : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Caractéristique
                           activité" - FQ N°14568
PT7   : 23/07/2007 FC V_72 FQ 14595 Libellé de l'élément dynamique
PT8   : 28/08/2007 FC V_80 Externalisation des fonctions @ d'édition du bulletin dans une DLL
PT9   : 26/12/2007 FC V_81 FQ 14120 Rupture par salarié non coché :
                           message "erreur:champ PPU_ETABLISSEMENT n'existe pas
}
unit CalcOLEGenericPaie;

interface

uses SysUtils,
     {$IFDEF EAGLCLIENT}
     {$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     Hctrls, Utob, Hent1, paramsoc, HDebug,EntPaie,StrUtils;

function PGCalcOLEGenericPaie(Sf, Sp: string): variant;
//PT8 function FunctPgPaieEditBulletin(Sf, Sp: string): variant; //PT4
//PT8 procedure FnPaieCreateBull(DateDeb, DateFin: TDateTime; Salarie, Etab, Page, Population: string); //PT4 //PT5
//PT8 function CalcZoneBasTauCoeMnt(ChampPHB,Etab,Salarie,ValChamp,NumZone:String;DateDeb,DateFin:TDateTime):double; //PT4

implementation

uses PGCongespayes,
{$IFNDEF CPS1}
     PGPOPULOUTILS,
{$ENDIF}
     DB;  //PT4

var
  { Fiche individuelle,Cumul,Réductions }
  GblTabLibEtab, GblTabMois: array[1..12] of string;
  GblTabMont, GblTabBaseRub, GblTabPatRub: array[1..12] of Double;
  GblTFicheInd, GblTCumPaie, GblTReduction: Tob;
  GblCumulMonnaie: Boolean;
  { Bulletin }
  //DEB PT4
(*PT8  GblTob_ParamPop:Tob;
  Tob_ExerciceSocial:Tob;
  DateDebutExerSoc:TDateTime;
  Tob_DateClotureCp,Tob_DateCp:Tob;
  MemDateDeb,MemDateFin:TDateTime;
  MemEtab:string;
  DateJamais,DateCp:TDateTime;
  CpAnnee,CpMois,CpJour:Word;
  Tab_ZonesCalculees: array[1..24] of Double;*)
  //FIN PT4


function PGColleZeroDevant(Nombre, Long: integer): string;
begin
  result := Format('%-' + IntToStr(Long) + '.' + IntToStr(Long) + 'd', [Nombre])
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 21/11/2003
Modifié le ... :   /  /
Description .. : Chargement des tobs
Mots clefs ... : PAIE,EDITION
*****************************************************************}
function FnCreateFid(DateDeb, DateFin: TDateTime; Salarie, Rupt, StWhere, ChampRupt, StRubNonImpr,StRubPat : string): string;
var
  QRech: TQuery;
  i, NbMM, MMAnt : Integer;
  ExpRupt, RuptSal, StRubImpr: string;
  YY, MM, JJ: Word;
  TobEtab: Tob;
begin
  ExpRupt := '';
  if Pos('PSA_SALARIE', salarie) = 1 then RuptSal := '-' else RuptSal := 'X';
  if Pos('PHB_', ChampRupt) > 0 then ExpRupt := 'AND ' + ChampRupt + '="' + Rupt + '" ';
  if StRubNonImpr = '-' then StRubImpr := 'AND PHB_IMPRIMABLE="X"' else StRubImpr := '';
  if StRubPat = 'X' then StRubPat := '' else StRubPat := 'OR PHB_MTPATRONAL<>0'; { PT1 }
  Debug('Paie Pgi : Chargement des données du salarié=' + salarie);
  { init Tob Montant rubrique à récupérer }
  if RuptSal = 'X' then
    QRech := OpenSql('SELECT PHB_NATURERUB,PHB_RUBRIQUE,PHB_DATEDEBUT,PHB_DATEFIN,' +
      'SUM(PHB_MTREM) AS MTREM,SUM(PHB_MTSALARIAL) AS MTSALARIAL,' +
      'SUM(PHB_BASECOT) AS BASECOT,SUM(PHB_MTPATRONAL) AS MTPATRONAL,' +
      'PHB_SALARIE,PHB_ORDREETAT,PHB_SENSBUL FROM HISTOBULLETIN ' +
      'LEFT JOIN SALARIES ON PHB_SALARIE=PSA_SALARIE ' +
      'WHERE PHB_SALARIE="' + Salarie + '" ' + ExpRupt + ' ' + StWhere + ' ' +
      'AND PHB_DATEFIN>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
      'AND (((PHB_MTREM<>0 OR PHB_MTSALARIAL<>0 '+StRubPat+') ' + StRubImpr + ') AND PHB_NATURERUB<>"BAS") ' + { PT1 }
      'GROUP BY PHB_DATEDEBUT,PHB_DATEFIN,PHB_SALARIE,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE,PHB_SENSBUL ' +
      'ORDER BY PHB_DATEDEBUT,PHB_DATEFIN,PHB_SALARIE,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE', TRUE)
  else
    QRech := OpenSql('SELECT PHB_NATURERUB,PHB_RUBRIQUE,PHB_DATEDEBUT,PHB_DATEFIN, ' +
      'PHB_ORDREETAT,PHB_SENSBUL,SUM(PHB_MTREM) AS MTREM,SUM(PHB_MTSALARIAL) AS ' +
      'MTSALARIAL,SUM(PHB_BASECOT) AS BASECOT,SUM(PHB_MTPATRONAL) AS MTPATRONAL ' +
      'FROM HISTOBULLETIN ' +
      'WHERE PHB_DATEFIN>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
      'AND (((PHB_MTREM<>0 OR PHB_MTSALARIAL<>0 '+StRubPat+') ' + StRubImpr + ') AND PHB_NATURERUB<>"BAS") ' + { PT1 }
      ExpRupt + ' ' + StWhere + ' ' +
      'GROUP BY PHB_DATEDEBUT,PHB_DATEFIN,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE,PHB_SENSBUL ' +
      'ORDER BY PHB_DATEDEBUT,PHB_DATEFIN,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE', TRUE);
  if not QRech.eof then
  begin
    GblTFicheInd := Tob.create('La_Tob_chargee', nil, -1);
    GblTFicheInd.LoadDetailDB('HISTO_BULLETIN', '', '', QRech, FALSE);
  end;
  Ferme(QRech);
  { Changement du préfixe des critères }
  if ExpRupt <> '' then
  begin
    Pos('HB', ExpRupt);
    ExpRupt[Pos('HB', ExpRupt)] := 'P';
    ExpRupt[Pos('B_', ExpRupt)] := 'U';
  end;
  for i := 1 to 14 do { 14 champs critères possible }
    if StWhere <> '' then
    begin
      if (Pos('HB_', StWhere) = 0) and (Pos('SA_', StWhere) = 0) then Break;
      if (Pos('HB_', StWhere) > 0) then
      begin
        StWhere[Pos('HB_', StWhere)] := 'P';
        StWhere[Pos('B_', StWhere)] := 'U';
      end;
      if (Pos('SA_', StWhere) > 0) then
      begin
        StWhere[Pos('SA_', StWhere)] := 'P';
        StWhere[Pos('A_', StWhere)] := 'U';
      end;
    end;

  { init Tob Montant Cumul à récuperer }
  if RuptSal = 'X' then
    QRech := OpenSql('SELECT PPU_DATEDEBUT,PPU_DATEFIN,SUM(PPU_CHEURESTRAV) AS CHEURESTRAV, ' +
      'SUM(PPU_CBRUT) AS CBRUT, SUM(PPU_CBRUTFISCAL) AS CBRUTFISCAL,SUM(PPU_CPLAFONDSS) AS CPLAFONDSS, ' +
      'SUM(PPU_CCOUTPATRON) AS CCOUTPATRON,SUM(PPU_CCOUTSALARIE) AS CCOUTSALARIE, ' +
      'SUM(PPU_CNETIMPOSAB) AS CNETIMPOSAB,SUM(PPU_CNETAPAYER) AS CNETAPAYER ' +
      'FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(DateDeb) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
      'AND PPU_SALARIE="' + Salarie + '" ' + ExpRupt + ' ' + StWhere + ' ' +
      'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN ' +
      'ORDER BY PPU_DATEDEBUT,PPU_DATEFIN', TRUE)
  else
    QRech := OpenSql('SELECT PPU_DATEDEBUT,PPU_DATEFIN,SUM(PPU_CHEURESTRAV) AS CHEURESTRAV, ' +
      'SUM(PPU_CBRUT) AS CBRUT, SUM(PPU_CBRUTFISCAL) AS CBRUTFISCAL,SUM(PPU_CPLAFONDSS) AS CPLAFONDSS, ' +
      'SUM(PPU_CCOUTPATRON) AS CCOUTPATRON,SUM(PPU_CCOUTSALARIE) AS CCOUTSALARIE, ' +
      'SUM(PPU_CNETIMPOSAB) AS CNETIMPOSAB,SUM(PPU_CNETAPAYER) AS CNETAPAYER ' +
      'FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(DateDeb) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
      ' ' + ExpRupt + ' ' + StWhere + ' ' +
      'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN ' +
      'ORDER BY PPU_DATEDEBUT,PPU_DATEFIN', TRUE);
  if not QRech.eof then
  begin
    GblTCumPaie := Tob.create('La Tob chargee', nil, -1);
    GblTCumPaie.LoadDetailDB('PAIE_ENCOURS', '', '', QRech, FALSE);
  end;
  Ferme(QRech);
  { Init Tob Etablissement et des mois edités }
  if RuptSal = 'X' then
    QRech := OpenSql('SELECT PPU_DATEDEBUT,PPU_DATEFIN,PPU_ETABLISSEMENT FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(DateDeb) + '" ' +
      'AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
      'AND PPU_SALARIE="' + Salarie + '" ORDER BY PPU_DATEDEBUT,PPU_DATEFIN', True)
  else
    QRech := OpenSql('SELECT DISTINCT PPU_DATEDEBUT,PPU_DATEFIN,PPU_ETABLISSEMENT FROM PAIEENCOURS ' + //PT9
      'WHERE PPU_DATEFIN>="' + UsDateTime(DateDeb) + '" ' +
      'AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '" ' + ExpRupt + ' ' + StWhere +
      'ORDER BY PPU_DATEDEBUT,PPU_DATEFIN', True);
  if not QRech.eof then
  begin
    TobEtab := Tob.create('Les établissements', nil, -1);
    TobEtab.LoadDetailDB('PAIE_ENCOURS', '', '', QRech, FALSE);
  end;
  Ferme(QRech);
  NbMM := 0;
  MMAnt := 0;
  if assigned(TobEtab) then
    for i := 0 to TobEtab.detail.count - 1 do
    begin
      DecodeDate(TobEtab.detail[i].GetValue('PPU_DATEFIN'), YY, MM, JJ);
      If MMAnt<>MM then
        Begin
        if (MMAnt<>0) and (MM<>MMAnt+1) then
          Begin
          if MM<MMAnt then
            NbMM:=NbMM+(12-MMAnt)+MM
          else
            if MM>MMAnt then
              NbMM:=NbMM+(MM-MMAnt);
          End
        else Inc(NbMM);
        if (NbMM<1) or (NbMM>12) then NbMM:=12;
        GblTabMois[NbMM] := IntToStr(MM);
        GblTabLibEtab[MM] := TobEtab.detail[i].GetValue('PPU_ETABLISSEMENT');
        End
      else
        if RuptSal = 'X' then
        begin
          if GblTabLibEtab[MM] = '' then
            GblTabLibEtab[MM] := TobEtab.detail[i].GetValue('PPU_ETABLISSEMENT')
          else
            GblTabLibEtab[MM] := GblTabLibEtab[MM] + ' | ' + TobEtab.detail[i].GetValue('PPU_ETABLISSEMENT');
        end;
     MMAnt:=MM;
    end;
  FreeAndNil(TobEtab);
  result := '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 20/11/2003
Modifié le ... :   /  /
Description .. : Function @ de la fiche individuelle ou récapitulative
Suite ........ : Englobe le traitement des functions
Mots clefs ... : PAIE;EDITION
*****************************************************************}
{----------------------------------------------------------------------------
Initialisation de tobs pour l'affichage des differents cumuls (HISTOCUMSAL):
La fiche ind reprend l'historique de paie sur une ou pls session de paie
Le tout est edité sous forme d'un tableau de 12 mois..
Dans le cas de pls bulletins pour une session de paie alors somme des bulletins..
----------------------------------------------------------------------------}
function FunctPgEditFicInd(Sf, Sp: string): variant;
var
  Salarie, Rupt, nature, Mois, StVal, StChamp, StWhere, StRubNonImpr, StRubPat : string;
  DateDeb, Datefin: TDateTime;
  i, NbMM: integer;
  somme: double;
  YY, MM, JJ, MMPrec: WORD;
  TRech: Tob;
begin
  try
    Result := '';
    { @PAIECREATEFID([PSA_SALARIE];[CRI_XX_RUPTURE1];[CRI_XX_VARIABLEDEB];[CRI_XX_VARIABLEFIN];[§CRI_XX_RUPTURE1];[CRI_CSAL]) }
    { Fn chargement des tobs }
    if sf = 'PAIECREATEFID' then
    begin
      Salarie := Trim(ReadTokenSt(sp));
      if (isnumeric(Salarie) and (GetParamSoc('SO_PGTYPENUMSAL') = 'NUM')) then
        Salarie := PGColleZeroDevant(ValeurI(Salarie), 10);
      Rupt := ReadTokenSt(sp);
      DateDeb := StrToDate(ReadTokenSt(sp));
      Datefin := StrToDate(ReadTokenSt(sp));
      StChamp := ReadTokenSt(sp);
      StWhere := ReadTokenSt(sp);
      StRubNonImpr := ReadTokenSt(sp);
      StRubPat := ReadTokenSt(sp);  { PT1 }
      for i := 1 to 12 do
      begin
        GblTabMois[i] := '';
        GblTabLibEtab[i] := '';
      end;
      if (Salarie <> '') and (DateDeb > idate1900) and (DateFin > idate1900) then
        Result := FnCreateFid(DateDeb, DateFin, Salarie, Rupt, StWhere, StChamp, StRubNonImpr,StRubPat); { PT1 }
    end
    else
      { @PAIEINITMONTRUB([XX_NATURERUB];[XX_RUBRIQUE]);  Fiche Individuelle et Recap }
      { Sur chaque ligne de rubrique on recherche les sommes des différents mois }
      { On initialise dans les tableaux les montants pouvant être utilisés : }
      { bases, mts salarials et patronals, mt rémunération..}
      if sf = 'PAIEINITMONTRUB' then
    begin
      if not Assigned(GblTFicheInd) then exit;
      nature := Trim(ReadTokenSt(sp));
      StVal := Trim(ReadTokenSt(sp));
      Debug('Paie Pgi : Initialisation rubrique=' + StVal);
      for i := 1 to 12 do
      begin
        GblTabMont[i] := 0;
        GblTabBaseRub[i] := 0;
        GblTabPatRub[i] := 0;
      end;
      { StChamp = Montant à récupérer }
      if nature = 'AAA' then StChamp := 'MTREM'
      else if nature = 'COT' then StChamp := 'MTSALARIAL';
      NbMM := 0;
      MMPrec := 0;
      { Balayage de la tob des historiques de rubrique }
      TRech := GblTFicheInd.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [Nature, StVal], FALSE);
      while TRech <> nil do
      begin { Chargement des tableaux }
        DecodeDate(TRech.GetValue('PHB_DATEFIN'), YY, MM, JJ);
        if MM <> MMPrec then Inc(NbMM);
        MMPrec := MM;
        if NbMM < 13 then { On se limite à 12 mois }
        begin
          if (TRech.GetValue('PHB_SENSBUL') <> 'M') then
          begin
            GblTabMont[MM] := GblTabMont[MM] + TRech.GetValue(StChamp);
            GblTabBaseRub[MM] := GblTabBaseRub[MM] + TRech.GetValue('BASECOT');
            GblTabPatRub[MM] := GblTabPatRub[MM] + TRech.GetValue('MTPATRONAL');
          end
          else
            if (TRech.GetValue('PHB_SENSBUL') = 'M') then
          begin
            GblTabMont[MM] := GblTabMont[MM] - TRech.GetValue(StChamp);
            GblTabBaseRub[MM] := GblTabBaseRub[MM] - TRech.GetValue('BASECOT');
            GblTabPatRub[MM] := GblTabPatRub[MM] - TRech.GetValue('MTPATRONAL');
          end;
        end;
        TRech.free; { Suppression des filles au fur à mesure pour gagner du temps de traitement }
        TRech := GblTFicheInd.FindNext(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [Nature, StVal], FALSE);
      end;
    end
    else

      {  @PAIEGETMONTRUB([MOIS]); Montant rémunération ou cotisation }
      {  @PAIEGETBASERUB([MOIS]); Base de cotisation }
      {  @PAIEGETPATRUB([MOIS]);  Montant patronal }
      {  retourne le montant de la rubrique du mois désiré à afficher }
      {  Le tableau est initialise dans INITMONTRUB }
      if sf = 'PAIEGETMONTRUB' then
    begin
      mois := Trim(ReadTokenSt(sp));
      if IsNumeric(mois) then
        if GblTabMont[ValeurI(mois)] <> 0 then
        begin
          result := GblTabMont[ValeurI(mois)];
        end;
    end
    else
      if sf = 'PAIEGETBASERUB' then
    begin
      mois := Trim(ReadTokenSt(sp));
      if IsNumeric(mois) then
        if GblTabBaseRub[ValeurI(mois)] <> 0 then
        begin
          result := GblTabBaseRub[ValeurI(mois)];
        end;
    end
    else

      if sf = 'PAIEGETPATRUB' then
    begin
      mois := Trim(ReadTokenSt(sp));
      if IsNumeric(mois) then
        if GblTabPatRub[ValeurI(mois)] <> 0 then
        begin
          result := GblTabPatRub[ValeurI(mois)];
        end;
    end
    else

      { @PAIEGETCUMPAIE([Cumul];[Mois]) ; }
      { retourne le montant du cumul à afficher }
      if sf = 'PAIEGETCUMPAIE' then
    begin
      if not assigned(GblTCumPaie) then exit;
      StChamp := '';
      Somme := 0;
      StVal := Trim(ReadTokenSt(sp));
      StVal := PGColleZeroDevant(ValeurI(StVal), 2);
      mois := Trim(ReadTokenSt(sp));
      if not IsNumeric(mois) then exit;
      if StVal = '01' then StChamp := 'CBRUT' else
        if StVal = '02' then StChamp := 'CBRUTFISCAL' else
        if StVal = '07' then StChamp := 'CCOUTPATRON' else
        if StVal = '08' then StChamp := 'CCOUTSALARIE' else
        if StVal = '09' then StChamp := 'CNETIMPOSAB' else
        if StVal = '10' then StChamp := 'CNETAPAYER' else
        if StVal = '30' then StChamp := 'CPLAFONDSS' else
        if StVal = '20' then StChamp := 'CHEURESTRAV';
      if StVal = '' then exit;
      if (IsNumeric(Mois)) and (GblTCumPaie.detail.count > 0) then
        for i := 0 to GblTCumPaie.detail.count - 1 do
        begin
          TRech := GblTCumPaie.detail[i];
          DecodeDate(TRech.GetValue('PPU_DATEFIN'), YY, MM, JJ);
          if MM = ValeurI(Mois) then
            Somme := Somme + TRech.GetValue(StChamp);
        end;
      if somme = 0 then result := ''
      else
      begin
        Result := somme;
      end;
    end
    else

      { [@PAIEGETETAB([M1])] Récupération de l'établissement mensuel }
      if sf = 'PAIEGETETAB' then
    begin
      mois := Trim(ReadTokenSt(sp));
      if not IsNumeric(mois) then exit;
      result := GblTabLibEtab[ValeurI(Mois)];
    end
    else

      if Sf = 'PAIEGETMOIS' then
    begin
      mois := Trim(ReadTokenSt(sp));
      if not IsNumeric(mois) then exit;
      Result := GblTabMois[ValeurI(mois)];
    end;
  except
    Result := '';
  end;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 08/12/2003
Modifié le ... : 08/12/2003
Description .. : Function @ de l'édition des cumuls
Mots clefs ... : PAIE;EDITION
*****************************************************************}
function FunctPgEditCumul(Sf, Sp: string): variant;
var
  Salarie, StVal, RuptSal, StWhere, StChamp: string;
  DateDeb, Datefin: TDateTime;
  QRech: TQuery;
  TRech: Tob;
  YY, MM, JJ: Word;
  i, NbMM, MMAnt: integer;
begin
  result := '';
  StChamp := '';
  { @PAIECREATECUMSAL([PHC_SALARIE];[CRI_XX_RUPTURE1];[CRI_XX_VARIABLEDEB];[CRI_XX_VARIABLEFIN]) }
  if Sf = 'PAIECREATECUMSAL' then
  begin
    StWhere := '';
    for i := 1 to 12 do GblTabMois[i] := '';
    Salarie := Trim(ReadTokenSt(sp));
    if (isnumeric(Salarie) and (GetParamSoc('SO_PGTYPENUMSAL') = 'NUM')) then
      Salarie := PGColleZeroDevant(ValeurI(Salarie), 10);
    StVal := ReadTokenSt(sp);
    DateDeb := StrToDate(ReadTokenSt(sp));
    Datefin := StrToDate(ReadTokenSt(sp));
    StChamp := ReadTokenSt(sp);
    StWhere := ReadTokenSt(sp);
    if Pos('PHC_SALARIE', salarie) = 1 then RuptSal := '-' else RuptSal := 'X';
    if RuptSal = 'X' then StWhere := StWhere + ' AND PHC_SALARIE="' + Salarie + '"'
    else if (RuptSal = '-') and (V_PGI.Confidentiel = '0') then StWhere := StWhere + ' AND PHC_CONFIDENTIEL<="'+IntToStr(V_PGI.NiveauAccesConf)+'"'; { PT3 } 
    Debug('Paie Pgi : Chargement des données');
    if (StChamp <> '') and (StVal <> '') then StWhere := StWhere + ' AND ' + StChamp + '="' + StVal + '"';
    { Chargement des données }
    Qrech := OpenSql('SELECT DISTINCT PHC_CUMULPAIE,PHC_DATEDEBUT,PHC_DATEFIN, ' +
      'SUM(PHC_MONTANT) AS MONTANT ' +
      'FROM HISTOCUMSAL WHERE PHC_DATEDEBUT>="' + UsDateTime(DateDeb) + '"  ' +
      'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
      '' + StWhere + ' ' +
      'GROUP BY PHC_DATEFIN,PHC_DATEDEBUT,PHC_CUMULPAIE ' +       {PT2}
      'ORDER BY PHC_DATEFIN,PHC_DATEDEBUT,PHC_CUMULPAIE', True);  {PT2}
    if not QRech.eof then
    begin
      GblTCumPaie := Tob.create('La Tob chargee', nil, -1);
      GblTCumPaie.LoadDetailDB('HISTO_CUMSAL', '', '', QRech, FALSE);
    end;
    Ferme(QRech);
    { Paramétrage des périodes }
    NbMM := 0;
    MMAnt := 0;
    if assigned(GblTCumPaie) then
    for i := 0 to GblTCumPaie.detail.count - 1 do
    begin
    TRech := GblTCumPaie.detail[i];
    DecodeDate(TRech.GetValue('PHC_DATEFIN'), YY, MM, JJ);
    TRech.AddChampSupValeur('PERIODE',MM);
    If MMAnt<>MM then
       Begin
       if (MMAnt<>0) and (MM<>MMAnt+1) then
          Begin
          if MM<MMAnt then
            NbMM:=NbMM+(12-MMAnt)+MM
          else
            if MM>MMAnt then
              NbMM:=NbMM+(MM-MMAnt);
          End
        else Inc(NbMM);
        if (NbMM<1) or (NbMM>12) then NbMM:=12;
        GblTabMois[NbMM] := IntToStr(MM);
        End;
     MMAnt:=MM;
    end;

  end

  else
    { @PAIEINITMONTCUM([PHC_CUMULPAIE]) Initialisation du tableau de données }
    if Sf = 'PAIEINITMONTCUM' then
  begin
    StVal := Trim(ReadTokenSt(sp));
    if StVal = '' then exit;
    if RechDom('PGTYPECUMUL', StVal, False) = 'X' then GblCumulMonnaie := True else GblCumulMonnaie := False;
    for i := 1 to 12 do GblTabMont[i] := 0;
    { On balaye la tob sur les cumuls }
    if assigned(GblTCumPaie) then
    begin
      TRech := GblTCumPaie.FindFirst(['PHC_CUMULPAIE'], [StVal], FALSE);
      while assigned(TRech) do
      begin
        NbMM := TRech.GetValue('PERIODE');
        GblTabMont[NbMM] := GblTabMont[NbMM] + TRech.GetValue('MONTANT');
        TRech.free;
        TRech := GblTCumPaie.FindNext(['PHC_CUMULPAIE'], [StVal], FALSE);
      end;
    end;
  end

  else
    { @PAIEGETMONTCUM([MOIS]) Récupération de la valeur }
    if sf = 'PAIEGETMONTCUM' then
  begin
    StVal := Trim(ReadTokenSt(sp));
    if (StVal <> '') then
      if GblTabMont[ValeurI(StVal)] <> 0 then
        result := GblTabMont[ValeurI(StVal)];
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 08/12/2003
Modifié le ... :   /  /
Description .. : Function @ de l'édition des charges sociales
Mots clefs ... : PAIE;EDITION
*****************************************************************}
function FunctPgEditChSocial(Sf, Sp: string): variant;
var
  DateDeb, Datefin, StVal, Rupt, StWhere, Critere, Org: string;
  DateRupt: TDateTime;
  QRech: TQuery;
begin
  result := '';
  { Renvoie le nombre d'hommes / femmes affecté à l'organisme selon rupture }
  if Sf = 'PAIEGETNBSEXE' then
  begin
    Rupt := ReadTokenSt(sp);
    StVal := ReadTokenSt(sp);
    DateDeb := ReadTokenSt(sp);
    Datefin := ReadTokenSt(sp);
    Org := ReadTokenSt(sp);
    Critere := ReadTokenSt(sp);
    if Sp <> '' then DateRupt := StrToDate(ReadTokenSt(sp)) else DateRupt := Idate1900;
    result := '';
    StWhere := '';
    if (Rupt <> '') and (StVal <> '') then StWhere := 'AND ' + Rupt + '="' + StVal + '" ';
    if DateRupt <> Idate1900 then StWhere := StWhere + 'AND PHB_DATEFIN="' + UsDateTime(DateRupt) + '" '
    else
    begin
      if (isValidDate(DateDeb)) then StWhere := StWhere + 'AND PHB_DATEFIN>="' + USDateTime(StrToDate(DateDeb)) + '" ';
      if (isValidDate(Datefin)) then StWhere := StWhere + 'AND PHB_DATEFIN<="' + USDateTime(StrToDate(Datefin)) + '" ';
    end;
    if Org <> '' then StWhere := StWhere + 'AND PHB_ORGANISME="' + Org + '" ';
    QRech := OpenSql('SELECT COUNT(DISTINCT PSA_SALARIE) SAL,PSA_SEXE ' +
      'FROM HISTOBULLETIN LEFT JOIN SALARIES ON PHB_SALARIE=PSA_SALARIE ' +
      'WHERE PHB_NATURERUB="COT" AND (PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0) ' + StWhere + Critere +
      'GROUP BY PSA_SEXE ORDER BY PSA_SEXE DESC', True);
    while not Qrech.eof do
    begin
      if Qrech.FindField('PSA_SEXE').AsString = 'M' then
        result := Result + 'Homme : ' + InttoStr(QRech.FindField('SAL').asinteger) + '     '
      else
        if Qrech.FindField('PSA_SEXE').AsString = 'F' then
        result := result + 'Femme : ' + InttoStr(QRech.FindField('SAL').asinteger) + '     ';
      Qrech.Next;
    end;
    Ferme(QRech);
    Result := trim(result);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 08/12/2003
Modifié le ... :   /  /
Description .. : Function @ de l'édition des différentes réductions
Mots clefs ... : PAIE;EDITION
*****************************************************************}
function InitFilleGblTReduction(StSal : string ): Tob;
begin
  Result := nil;
  if not Assigned(GblTReduction) then exit;
  Result := Tob.Create('Les Filles réductions', GblTReduction, -1);
  Result.AddChampSupValeur('SALARIE'       , StSal);
  Result.AddChampSupValeur('HEURES'        , 0);
  Result.AddChampSupValeur('REMUNERATION'  , 0);
  Result.AddChampSupValeur('REPAS'         , 0);
  Result.AddChampSupValeur('ALLEGEMENT'    , 0);
  Result.AddChampSupValeur('MAJORATION'    , 0);
  Result.AddChampSupValeur('MINORATION'    , 0);
end;

function FunctPgEditReduction(Sf, Sp: string): variant;
var
  Tob_Red, T1: Tob;
  Datedeb, DateFin : TDateTime;
  StSal, StTemp, TypeEtat, Etab, Annee, Mois : string;
  i, j: Integer;
  QRech: TQuery;
begin
  { Chargement des données nécessaires aux éditions }
  if Sf = 'PAIECHARGETOBREDUCTION' then
  begin
    TypeEtat := ReadTokenSt(sp);
    Datedeb := StrToDate(ReadTokenSt(sp));
    Datefin := StrToDate(ReadTokenSt(sp));
    Etab := ReadTokenSt(sp);
    Annee := ReadTokenSt(sp);
    Mois  := ReadTokenSt(sp);
    if (GetParamSoc('SO_PGREDHEURE') <> '') and (GetParamSoc('SO_PGREDREM') <> '') then
      StTemp := ' AND (PHC_CUMULPAIE="' + GetParamSoc('SO_PGREDHEURE') + '" OR PHC_CUMULPAIE="' + GetParamSoc('SO_PGREDREM') + '") '
    else
      if (GetParamSoc('SO_PGREDHEURE') <> '') then
      StTemp := ' AND PHC_CUMULPAIE="' + GetParamSoc('SO_PGREDHEURE') + '" '
    else
      if (GetParamSoc('SO_PGREDREM') <> '') then
      StTemp := ' AND PHC_CUMULPAIE="' + GetParamSoc('SO_PGREDREM') + '" '
    else
      StTemp := '';
    { Si rupture périodique, chargement session }
    if (IsNumeric(Annee)) AND (ISNumeric(Mois)) then
       Begin
       DateDeb := EncodeDate(ValeurI(Annee),ValeurI(Mois),1);
       DateFin := FindeMois (DateDeb);
       End;
    if (TypeEtat = 'PBS') or (TypeEtat = 'PLF') then StTemp := StTemp + ' AND PCT_REDUCBASSAL="X" '
    else if (TypeEtat = 'PRP') then StTemp := StTemp + ' AND PCT_REDUCREPAS="X" '
    else if (TypeEtat = 'PLA') then StTemp := StTemp + ' AND (PCT_ALLEGEMENTA2="X" OR PCT_MAJORATA2="X" OR PCT_MINORATA2="X") ';

    QRech := OpenSql('SELECT SUM(PHC_MONTANT) as MONTANT,PHC_CUMULPAIE,PHC_SALARIE '+
            'FROM HISTOCUMSAL ' +
            'WHERE PHC_ETABLISSEMENT="' + Etab + '" AND PHC_REPRISE="-" '+
            'AND PHC_DATEFIN>="' + UsDateTime(Datedeb) + '" AND PHC_DATEFIN<="' + UsDateTime(Datefin) + '" ' +
            'AND PHC_SALARIE IN (SELECT PHB_SALARIE FROM HISTOBULLETIN LEFT JOIN COTISATION '+
            'ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE '+
            'WHERE PHB_DATEDEBUT=PHC_DATEDEBUT AND PHB_DATEFIN=PHC_DATEFIN '+ StTemp +
            'AND PHB_DATEFIN>="' + UsDateTime(Datedeb) + '" AND PHB_DATEFIN<="' + UsDateTime(Datefin) + '" ' +
            'AND PHB_SALARIE=PHC_SALARIE AND PHB_ETABLISSEMENT=PHC_ETABLISSEMENT  AND PHB_MTPATRONAL<>0 ) '+
            'GROUP BY PHC_SALARIE,PHC_CUMULPAIE', TRUE);
    if not QRech.eof then
    begin
      Tob_Red := Tob.Create('Les Reductions', nil, -1);
      Tob_Red.LoadDetailDB('HISTO_CUMSAL', '', '', QRech, False);
    end;
    Ferme(QRech);
    { Création de la tob principale des réductions }
    GblTReduction := Tob.Create('Les Réductions', nil, -1);
    if Assigned(Tob_Red) then
    begin
      for i := 0 to Tob_Red.detail.Count - 1 do
      begin
        StSal := Tob_Red.detail[i].GetValue('PHC_SALARIE');
        T1 := GblTReduction.FindFirst(['SALARIE'], [StSal], False);
        if not Assigned(T1) then T1 := InitFilleGblTReduction(StSal);
        if Assigned(T1) then
        begin
          if Tob_Red.detail[i].GetValue('PHC_CUMULPAIE') = GetParamSoc('SO_PGREDHEURE') then
            T1.PutValue('HEURES', T1.GetValue('HEURES') + Tob_Red.detail[i].GetValue('MONTANT'));
          if Tob_Red.detail[i].GetValue('PHC_CUMULPAIE') = GetParamSoc('SO_PGREDREM') then
            T1.PutValue('REMUNERATION', T1.GetValue('REMUNERATION') + Tob_Red.detail[i].GetValue('MONTANT'));
        end;
      end;
      FreeAndNil(Tob_Red);
    end;
    { Concernant l'édition des réductions repas }
    if TypeEtat = 'PRP' then
    begin
      if (GetParamSoc('SO_PGTYPREDREPAS') <> '') and (GetParamSoc('SO_PGREDREPAS') <> '') then
      begin
        if GetParamSoc('SO_PGTYPREDREPAS') = '05' then StTemp := 'PHB_BASEREM' //Base
        else if GetParamSoc('SO_PGTYPREDREPAS') = '06' then StTemp := 'PHB_TAUXREM' //Taux
        else if GetParamSoc('SO_PGTYPREDREPAS') = '07' then StTemp := 'PHB_COEFFREM' //Coeff
        else if GetParamSoc('SO_PGTYPREDREPAS') = '08' then StTemp := 'PHB_MTREM'; //Montant
        QRech := OpenSql('SELECT SUM(' + StTemp + ') MONTANT,PHB_SALARIE FROM HISTOBULLETIN ' +
          'WHERE PHB_MTREM<>0  AND PHB_ETABLISSEMENT="' + Etab + '" ' +
          'AND PHB_RUBRIQUE="' + GetParamSoc('SO_PGREDREPAS') + '" ' +
          'AND PHB_DATEFIN>="' + UsDateTime(DateDeb) + '" ' +
          'AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" GROUP BY PHB_SALARIE', TRUE);
        if not QRech.eof then
        begin
          Tob_Red := Tob.Create('Divers ', nil, -1);
          Tob_Red.LoadDetailDB('HISTO_CUMSAL', '', '', QRech, False);
        end;
        Ferme(QRech);
        if Assigned(Tob_Red) then
        begin
          for i := 0 to Tob_Red.detail.Count - 1 do
          begin
            StSal := Tob_Red.detail[i].GetValue('PHB_SALARIE');
            T1 := GblTReduction.FindFirst(['SALARIE'], [StSal], False);
            if not Assigned(T1) then T1 := InitFilleGblTReduction(StSal);
            if Assigned(T1) then T1.PutValue('REPAS', T1.GetValue('REPAS') + Tob_Red.detail[i].GetValue('MONTANT'));
          end;
          FreeAndNil(Tob_Red);
        end;
      end;
    end
    else
      { Concernant l'édition des réductions Loi Aubry 2 }
      if TypeEtat = 'PLA' then
      for j := 1 to 3 do
      begin
        if j = 1 then StTemp := 'PCT_ALLEGEMENTA2'
        else if j = 2 then StTemp := 'PCT_MAJORATA2'
        else if j = 3 then StTemp := 'PCT_MINORATA2';
        QRech := OpenSql('SELECT SUM(PHB_MTPATRONAL) MONTANT,PHB_SALARIE FROM HISTOBULLETIN ' +
          'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB ' +
          'AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
          'WHERE PHB_DATEFIN>="' + UsDateTime(DateDeb) + '" ' +
          'AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
          'AND PHB_ETABLISSEMENT="' + Etab + '" ' +
          'AND ' + StTemp + '="X" AND PHB_MTPATRONAL<>0 ' +
          'GROUP BY PHB_SALARIE', True);
        if not QRech.eof then
        begin
          Tob_Red := Tob.Create('Divers ', nil, -1);
          Tob_Red.LoadDetailDB('HISTO_BULLETIN', '', '', QRech, False);
        end;
        Ferme(QRech);
        if Assigned(Tob_Red) then
        begin
          for i := 0 to Tob_Red.detail.Count - 1 do
          begin
            StSal := Tob_Red.detail[i].GetValue('PHB_SALARIE');
            T1 := GblTReduction.FindFirst(['SALARIE'], [StSal], False);
            if not Assigned(T1) then T1 := InitFilleGblTReduction(StSal);
            if Assigned(T1) then
            begin
              if StTemp = 'PCT_ALLEGEMENTA2' then
                T1.PutValue('ALLEGEMENT', T1.GetValue('ALLEGEMENT') + Tob_Red.detail[i].GetValue('MONTANT'))
              else
                if StTemp = 'PCT_MAJORATA2' then
                T1.PutValue('MAJORATION', T1.GetValue('MAJORATION') + Tob_Red.detail[i].GetValue('MONTANT'))
              else
                if StTemp = 'PCT_MINORATA2' then
                T1.PutValue('MINORATION', T1.GetValue('MINORATION') +Tob_Red.detail[i].GetValue('MONTANT'));
            end;
          end;
          FreeAndNil(Tob_Red);
        end;
      end;
  end
  else
    { Récupération des données }
    if SF = 'PAIEGETVALREDUCTION' then
  begin
    StTemp := ReadTokenSt(sp);
    StSal := ReadTokenSt(sp);
    if (StTemp <> '') and (StSal <> '') and (Assigned(GblTReduction)) then
    begin
      Tob_Red := GblTReduction.findFirst(['SALARIE'], [StSal], False);
      if Assigned(Tob_Red) then
         Begin
         Result := Tob_Red.GetValue(StTemp);
         Tob_Red.PutValue(StTemp,0);
         If (Tob_Red.GetValue('HEURES')=0) AND (Tob_Red.GetValue('REMUNERATION')=0) AND
            (Tob_Red.GetValue('REPAS')=0)  AND (Tob_Red.GetValue('ALLEGEMENT')=0)   AND
            (Tob_Red.GetValue('MAJORATION')=0) AND (Tob_Red.GetValue('MINORATION')=0) Then Tob_Red.free;
         End;
    end;
  end;

end;

//DEB PT4
{***********A.G.L.Privé.*****************************************
Auteur  ...... : FC
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Function @ de l'édition des bulletins
Mots clefs ... : PAIE;EDITION
*****************************************************************}
(*  PT8
procedure FnPaieCreateBull(DateDeb, DateFin: TDateTime; Salarie, Etab, Page, Population: string); //PT5
var
  W_PgParametre,W_TypeParametre : String;
  QTemp,QParamAssoc,Q       :TQuery;
  i,j           :integer;
  DTcloturePi, DTDebutPi, Tdate: TDateTime;
  EnAnnee, EnMois, EnJour, RMois: Word;
  CodeAssocie : String;
begin
  if (Datefin < 1) or (DateDeb < 1) or (Salarie = '') or (Etab = '') then exit;

  if (Page = '1') or (Page = '') or (DateDeb <> MemDateDeb) or (DateFin <> MemDateFin) or (Etab <> MemEtab) then
  begin
    MemDateDeb := DateDeb; // rupture sur les dates donc rechargement
    MemDateFin := DateFin;
    MemEtab := Etab;

    //init date debut exercice social pour l'alimentation des cumuls
    QTemp := OpenSql('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL ORDER BY PEX_DATEFIN',True);
    if not QTemp.eof then //PORTAGECWAS
      Begin
      Tob_ExerciceSocial := Tob.Create('Les exercices social',nil,-1);
      Tob_ExerciceSocial.LoadDetailDB('','','',QTemp,False);
      End
    else
      Tob_ExerciceSocial := nil;
    Ferme(QTemp);

    //init DateCP pour l'alimentation des cumuls
    Tob_DateClotureCp := Tob.create('Date de cloture', nil, -1);
    if Tob_DateClotureCp <> nil then
    begin
      QTemp := OpenSql('SELECT ETB_ETABLISSEMENT,ETB_DATECLOTURECPN FROM ETABCOMPL ', True);
      while not QTemp.eof do
      begin
        Tob_DateCp := Tob.create('Date de cloture', Tob_DateClotureCp, -1);
        if Tob_DateCp = nil then break;
        Tob_DateCp.AddChampSup('ETABLISSEMENT', False);
        Tob_DateCp.AddChampSup('DATECLOTURE', False);
        TDate := QTemp.FindField('ETB_DATECLOTURECPN').AsDateTime;
        if (TDate > 0) then
        begin
          RMois := 1;
          RendPeriode(DTcloturePi, DTDebutPi, TDate, DateFin);
          DecodeDate(DateFin, EnAnnee, EnMois, EnJour);
          DecodeDate(DTcloturePi, CpAnnee, Cpmois, Cpjour);
          if EnMois <= CpMois then
          begin
            EnAnnee := EnAnnee - 1;
            RMois := Cpmois + 1;
          end;
          if EnMois > Cpmois then RMois := Cpmois + 1;
          if CpMois = 12 then RMois := 1;
          DateCP := EncodeDate(EnAnnee, RMois, 1);
        end
        else DateCP := idate1900;
        Tob_DateCp.PutValue('ETABLISSEMENT', QTemp.FindField('ETB_ETABLISSEMENT').AsString);
        Tob_DateCp.PutValue('DATECLOTURE', DateCP);
        QTemp.next;
      end;
      Ferme(QTemp);
    end;
  end;

  if (Page = '1') or (Page = '') then
  begin
    // Réinitialiser le tableau contenant les valeurs déjà calculées
    for j := 1 to 24 do
      Tab_ZonesCalculees[j] := 0;
  end;

{$IFNDEF CPS1}
  if Assigned(GblTob_ParamPop) then FreeAndNil(GblTob_ParamPop);

  //DEB PT5
  CodeAssocie := '';
  if Population = '' then
  begin
    Q := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
      + ' WHERE PNA_SALARIE = "' + Salarie + '"'
      + ' AND PNA_TYPEPOP = "PAI"', True);   // Anciennement ZLB
    if not Q.Eof then
      CodeAssocie := Q.FindField('PNA_POPULATION').AsString;
    Ferme(Q);
  end
  else
    CodeAssocie := Population;
  //FIN PT5

  QParamAssoc := OpenSql('SELECT PGO_PGPARAMETRE, PGO_TYPEPARAMETRE ' +
    ' FROM PGPARAMETRESASSOC ' +
    ' WHERE PGO_CODEASSOCIE = "' + CodeAssocie + '"' +
    ' AND PGO_DATEVALIDITE <= "' + UsDateTime(DateFin) + '"' +
    ' AND PGO_PGPARAMETRE LIKE "PAI%"' +  // Anciennement ZLB
    ' ORDER BY PGO_DATEVALIDITE DESC', True);
  // Chargement des paramètres propre à la population
  if not QParamAssoc.Eof then
  begin
    W_PgParametre   := QParamAssoc.FindField('PGO_PGPARAMETRE').AsString;
    W_TypeParametre := QParamAssoc.FindField('PGO_TYPEPARAMETRE').AsString;
    QTemp := OpenSql('SELECT PGP_PGNOMCHAMP, PGP_PGVALCHAMP FROM PGPARAMETRES ' +
      'WHERE ##PGP_PREDEFINI## PGP_PGPARAMETRE = "'+ W_PgParametre + '"' +
      ' AND PGP_TYPEPARAMETRE = "' + W_TypeParametre + '"', True);
    if not QTemp.eof then
    begin
      GblTob_ParamPop := Tob.create('PARAMPOP', nil, -1);
      GblTob_ParamPop.LoadDetailDB('PARAMPOP', '', '', QTemp, FALSE);
    end
    else GblTob_ParamPop := nil;
    Ferme(QTemp);
  end;
  Ferme(QParamAssoc);
{$ENDIF}

  QTemp := OpenSQl('SELECT MIN(PHC_DATEDEBUT) AS DATEDEB FROM HISTOCUMSAL ' +
    'WHERE PHC_SALARIE="' + Salarie + '"', True);
  if not QTemp.EOF then //PORTAGECWAS
    DateJamais := QTemp.FindField('DATEDEB').AsDateTime
  else
    DateJamais := idate1900;
  Ferme(QTemp);

  if Tob_DateClotureCp <> nil then
  begin
    Tob_DateCP := Tob_DateClotureCp.FindFirst(['ETABLISSEMENT'], [etab], False);
    if Tob_DateCP <> nil then
      DateCp := Tob_DateCP.GetValue('DATECLOTURE')
    else
      DateCp := Idate1900;
  end
  else
    DateCp := Idate1900;

  DateDebutExerSoc := idate1900;
  For i := 0 to Tob_ExerciceSocial.detail.count-1 do
  Begin
    if (Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEDEBUT') <= DateDeb)
    and (Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEFIN') >= DateFin)  then
    Begin
      DateDebutExerSoc :=  Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEDEBUT');
      Break;
      End;
  End;
end; *)

(*PT8
function FunctPgPaieEditBulletin(Sf, Sp: string): variant;
var
  NumZone,Salarie,Etab,Page,Libelle, Popul :String;
  iDateDeb,iDateFin                 :integer;
  DateDeb,DateFin                   :TDateTime;
  NomChamp,Requete,Raz              :String;
  Tob_Temp,Tob_Temp2                :Tob;
  QCumulPaie,QCumulMois,QEltDynamique,QRechCumul,QSalarie  :TQuery;
  CumulDD                           :TDateTime;
  Resultat,Resultat1,Resultat2      :double;
  Qt : TQuery;
  CodeTabLib, St, Lib : String;
  Q : TQuery;
begin
  //Bulletin de paie , pour une session de paie
  //@PAIECREATEBULL([XX_DATEDEBUT];[XX_DATEFIN];[XX_SALARIE];[XX_ETABLISSEMENT];[PAGE])
  if Sf = 'PAIECREATEBULL' then
  begin
    iDateDeb := StrToInt(ReadTokenSt(sp));
    iDateFin := StrToInt(ReadTokenSt(sp));
    Salarie := Trim(ReadTokenSt(sp));
    if (isnumeric(Salarie) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      Salarie := PGColleZeroDevant(StrToInt(Salarie), 10);
    Etab := ReadTokenSt(sp);
    Page := ReadTokenSt(sp);
    Popul := ReadTokenSt(sp);
    if (Etab <> '') and (Salarie <> '') and (iDateDeb > 0) and (iDateFin > 0) then
      FnPaieCreateBull(iDateDeb, iDateFin, Salarie, Etab, Page, Popul);
    result := '';
    exit;
  end;

  // Fonction de récupération des éléments dynamiques présents sur le bulletin
  if Sf = 'PAIEGETVALZONE' then
  begin
    // Récupération des paramètres de la fonction @
    NumZone := Trim(ReadTokenSt(sp));
    Salarie := Trim(ReadTokenSt(sp));
    if (isnumeric(Salarie) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      Salarie := PGColleZeroDevant(StrToInt(Salarie), 10);
    Etab := ReadTokenSt(sp);
    if (isnumeric(Etab)) then
      Etab := PGColleZeroDevant(StrToInt(Etab), 3);
    // La récupération des dates peut poser problème si sur l'état, un format <> 'C' est renseigné pour la fonction
    // car dans ce cas, les dates ne reviennent pas sous le format "dd/mm/yyyy" mais sous la forme d'un numérique
//    iDateDeb := StrToInt(ReadTokenSt(sp));
//    iDateFin := StrToInt(ReadTokenSt(sp));
    DateDeb := StrToDate(ReadTokenSt(sp));
    DateFin := StrToDate(ReadTokenSt(sp));
    Page := ReadTokenSt(sp);

    // Inutile de continuer si on a aucun paramètre
    if GblTob_ParamPop = nil then exit;

    Result := '';
    Resultat  := 0; Resultat1 := 0; Resultat2 := 0;
    if (NumZone <> '') and (Salarie <> '') and (Assigned(GblTob_ParamPop)) then
    begin
      NomChamp := 'TYPZONE' + NumZone;
      Tob_Temp := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
      if assigned(Tob_Temp) then
      begin
        // -------- ETIQUETTE --------
        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ETQ') then
        begin
          NomChamp := 'VALZONE' + NumZone;
          Tob_Temp2 := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
            Result := Tob_Temp2.GetValue('PGP_PGVALCHAMP');
        end;

        // -------- DONNEE FICHE SALARIE --------
        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'SAL') then
        begin
          NomChamp := 'ZSAL' + NumZone;
          Tob_Temp2 := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
          begin
            //Si on gère l'historisation des données salariés
            result := '';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'COEFFICIENT') then
              Libelle := 'Coefficient';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'QUALIFICATION') then
              Libelle := 'Qualification';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CODEEMPLOI') then
              Libelle := 'Emploi PCS';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBELLEEMPLOI') then
              Libelle := 'Emploi';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONDEMPLOI') then
{PT6
              Libelle := 'Condition d''emploi';
}
              Libelle := 'Caractéristique activité';
//FIN PT6              
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATEANCIENNETE') then
              Libelle := 'Date d''ancienneté';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATEENTREE') then
              Libelle := 'Date d''entrée';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'HORAIREMOIS') then        //******
              Libelle := 'Horaire';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'SALAIRETHEO') then
              Libelle := 'Salaire';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TAUXHORAIRE') then
              Libelle := 'Taux horaire';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'SALAIREMOIS1') then
              Libelle := 'Salaire mensuel';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'SALAIRANN1') then
              Libelle := 'Salaire annuel';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'INDICE') then
              Libelle := 'Indice';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'NIVEAU') then
              Libelle := 'Niveau';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSPROF') then
              Libelle := 'Statut professionnel DADSU';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSCAT') then
              Libelle := 'Catégorie DUCS';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CATBILAN') then
              Libelle := 'Catégorie bilan social';
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONVENTION') then
              Libelle := 'Convention collective';
            if (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 1, Length(Tob_Temp2.GetValue('PGP_PGVALCHAMP')) - 1)) = 'TRAVAILN' then
            begin
              if GetParamSocSecur('SO_PGLIBORGSTAT' + (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 9, 1)),'') <> '' then
                Libelle := GetParamSocSecur('SO_PGLIBORGSTAT' + (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 9, 1)),'')
              else
                Libelle := 'Libellé organisation ' + (MidStr(Tob_Temp2.GetValue('PGP_PGVALCHAMP'), 9, 1));
            end;
            if VH_Paie.PgHistorisation then
            begin
              Requete := 'SELECT PHD_NEWVALEUR,PHD_TABLETTE FROM PGHISTODETAIL ' +
                ' WHERE PHD_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                ' AND PHD_SALARIE = "' + Salarie + '"' +
                ' AND PHD_DATEAPPLIC <= "' + USDATETIME(DateFin) + '"' +
                ' ORDER BY PHD_DATEAPPLIC DESC';
              QEltDynamique := OpenSql(Requete, TRUE);
              if not QEltDynamique.eof then //PORTAGECWAS
              begin
                if QEltDynamique.FindField('PHD_TABLETTE').AsString <> '' then
                begin
                  // Cas particulier pour la qualification car nom tablette tronqué. Il faut la récupérer dans PPP
                  if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'QUALIFICATION') then
                  begin
                    Qt := OpenSql('SELECT PPP_LIENASSOC FROM PARAMSALARIE ' +
                      ' WHERE PPP_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '"' +
                      ' AND PPP_PREDEFINI = "CEG"', True);
                    if not Qt.Eof then
                      result := Libelle + ' : ' + RechDom(Qt.FindField('PPP_LIENASSOC').AsString,QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False)
                    else
                      result := Libelle + ' : ' + QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
                    Ferme(Qt);
                  end
                  else
                  begin
                    if RechDom(QEltDynamique.FindField('PHD_TABLETTE').AsString,QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False) <> '' then
                      result := Libelle + ' : ' + RechDom(QEltDynamique.FindField('PHD_TABLETTE').AsString,QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False)
                    else
                      result := Libelle + ' : ' + QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
                  end;
                end
                else
                  result := Libelle + ' : ' + QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
              end
              else
                result := '';
              Ferme(QEltDynamique);
            end;
            if (result = '') then
            begin
              Requete := 'SELECT PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + ' FROM SALARIES' +
                ' WHERE PSA_SALARIE = "' + Salarie + '"';
              QSalarie := OpenSql(Requete, TRUE);
              if not QSalarie.Eof then //PORTAGECWAS
                if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TAUXPARTIEL') then
                  Result := Libelle + ' : ' + FloatToStr(QSalarie.Fields[0].AsFloat) + ' %'
                else
                begin
                  if      (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CODEEMPLOI') then
                    Result := Libelle + ' : ' + RechDom('PGCODEPCSESE',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBELLEEMPLOI') then
                    Result := Libelle + ' : ' + RechDom('PGLIBEMPLOI',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'QUALIFICATION') then
                    Result := Libelle + ' : ' + RechDom('PGLIBQUALIFICATION',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'COEFFICIENT') then
                    Result := Libelle + ' : ' + RechDom('PGLIBCOEFFICIENT',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'INDICE') then
                    Result := Libelle + ' : ' + RechDom('PGLIBINDICE',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'NIVEAU') then
                    Result := Libelle + ' : ' + RechDom('PGLIBNIVEAU',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSCAT') then
                    Result := Libelle + ' : ' + RechDom('PGSCATEGORIEL',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DADSPROF') then
                    Result := Libelle + ' : ' + RechDom('PGSPROFESSIONNEL',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONDEMPLOI') then
                    Result := Libelle + ' : ' + RechDom('PGCONDEMPLOI',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CATBILAN') then
                    Result := Libelle + ' : ' + RechDom('PGCATBILAN',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'CONVENTION') then
                    Result := Libelle + ' : ' + RechDom('PGCATBILAN',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN1') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN1',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN2') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN2',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN3') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN3',QSalarie.Fields[0].AsString,False)
                  else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'TRAVAILN4') then
                    Result := Libelle + ' : ' + RechDom('PGTRAVAILN4',QSalarie.Fields[0].AsString,False)
                  else
                    Result := Libelle + ' : ' + QSalarie.Fields[0].AsString;
                end;
              Ferme(QSalarie);
            end;
          end;
        end;

        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') <> 'ETQ') and (Tob_Temp.GetValue('PGP_PGVALCHAMP') <> 'RES') then
        begin
          NomChamp := 'VALZONE' + NumZone;
          Tob_Temp2 := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
          begin
            // --------  CUMUL MOIS --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'CMM') then
            begin
              Requete := 'SELECT SUM(PHC_MONTANT) AS MONTANT ' +
                'FROM HISTOCUMSAL WHERE PHC_SALARIE="' + Salarie + '" AND ' +
                'PHC_CUMULPAIE="' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                'AND PHC_DATEDEBUT>="' + UsDateTime(DateDeb) + '" ' +
                'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ';
              QCumulPaie := OpenSql(Requete, TRUE);
              if not QCumulPaie.eof then //PORTAGECWAS
                Resultat := QCumulPaie.FindField('MONTANT').AsFloat;
              Ferme(QCumulPaie);

              Tab_ZonesCalculees[strtoint(NumZone)] := Resultat;
              Result := RechDom('PGCUMULPAIE',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + FormatFloat('#,##0.00',Resultat);
            end;

            // -------- ELEMENT DYNAMIQUE LIBRE (ex carte orange) --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ELD') then
            begin
              Requete := 'SELECT PHD_NEWVALEUR,PHD_CODTABL FROM PGHISTODETAIL ' +
                ' WHERE PHD_PGINFOSMODIF = "' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                ' AND PHD_SALARIE = "' + Salarie + '"' +
                ' AND PHD_DATEAPPLIC <= "' + USDATETIME(DateFin) + '"' +
                ' ORDER BY PHD_DATEAPPLIC DESC';
              QEltDynamique := OpenSql(Requete, TRUE);
              if not QEltDynamique.eof then //PORTAGECWAS
              begin
                //DEB PT7
                CodeTabLib := QEltDynamique.FindField('PHD_CODTABL').AsString;
                St := ' SELECT PTD_LIBELLECODE FROM TABLEDIMDET WHERE PTD_DTVALID IN (SELECT MAX(PTE_DTVALID) FROM TABLEDIMENT WHERE PTD_CODTABL=PTE_CODTABL '+
                'AND PTE_DTVALID<="'+UsDateTime(DateFin)+'") AND PTD_CODTABL="'+CodeTabLib+'" AND PTD_VALCRIT1="'+QEltDynamique.FindField('PHD_NEWVALEUR').AsString+'"';
                Q := OpenSql(St, TRUE);
                Lib := '';
                if not Q.eof then //PORTAGECWAS
                  Lib:= Q.FindField('PTD_LIBELLECODE').AsString;
                Ferme(Q);
                if Lib = '' then
                  Lib := QEltDynamique.FindField('PHD_NEWVALEUR').AsString;
                //FIN PT7
                result := RechDom('PGZONEHISTOSAL',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + Lib;
              end
              else
                result := '';
              Ferme(QEltDynamique);
            end;

            // -------- ELEMENT DYNAMIQUE MONTANT -------- (valeur particulière d'un élément national pour un salarié)
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ELM') then
            begin
              Requete := 'SELECT PED_MONTANTEURO,PED_MONTANT FROM ELTNATIONDOS ' +
                ' WHERE PED_TYPENIVEAU = "SAL" ' +
                ' AND PED_VALEURNIVEAU = "' + Salarie + '"' +
                ' AND PED_DATEVALIDITE <= "' + UsDateTime(DateDeb) + '"' +
                ' AND PED_CODEELT = "' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '"' +
                ' ORDER BY PED_DATEVALIDITE DESC';
              QEltDynamique := OpenSql(Requete, TRUE);
              if not QEltDynamique.eof then //PORTAGECWAS
                Resultat := QEltDynamique.FindField('PED_MONTANTEURO').AsFloat;
              if Resultat = 0 then
                Resultat := QEltDynamique.FindField('PED_MONTANT').AsFloat;
              Ferme(QEltDynamique);

              Tab_ZonesCalculees[strtoint(NumZone)] := Resultat;
              Result := RechDom('PGELEMENTNAT',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + FormatFloat('#,##0.00',Resultat);
            end;

            // -------- ZONES LIBRES -------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'ZLS') then
            begin
              NomChamp := 'ZLS' + NumZone;
              Tob_Temp2 := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
              if assigned(Tob_Temp2) then
              begin
                result := '';
                if VH_Paie.PgHistorisation then
                begin
                  if ExisteSQL('SELECT PPP_PGINFOSMODIF FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '"' +
                    ' AND PPP_HISTORIQUE="X" AND ##PPP_PREDEFINI##') then
                    begin
                    Requete := 'SELECT PHD_NEWVALEUR,PHD_TABLETTE FROM PGHISTODETAIL ' +
                      ' WHERE PHD_PGINFOSMODIF = "PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                      ' AND PHD_SALARIE = "' + Salarie + '"' +
                      ' AND PHD_DATEAPPLIC <= "' + USDATETIME(DateFin) + '"' +
                      ' ORDER BY PHD_DATEAPPLIC DESC';
                    QEltDynamique := OpenSql(Requete, TRUE);
                    if not QEltDynamique.eof then //PORTAGECWAS
                    begin
                      // Dates libres
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE1') then
                        result := GetParamsoc('SO_PGLIBDATE1') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE2') then
                        result := GetParamsoc('SO_PGLIBDATE2') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE3') then
                        result := GetParamsoc('SO_PGLIBDATE3') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE4') then
                        result := GetParamsoc('SO_PGLIBDATE4') + ' : ' + DateTimeToStr(QEltDynamique.FindField('PHD_NEWVALEUR').AsDateTime);

                      // Combos libres
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB1') then
                        result := GetParamsoc('SO_PGLIBCOMBO1') + ' : ' + RechDom('PGLIBREPCMB1',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB2') then
                        result := GetParamsoc('SO_PGLIBCOMBO2') + ' : ' + RechDom('PGLIBREPCMB2',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB3') then
                        result := GetParamsoc('SO_PGLIBCOMBO3') + ' : ' + RechDom('PGLIBREPCMB3',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB4') then
                        result := GetParamsoc('SO_PGLIBCOMBO4') + ' : ' + RechDom('PGLIBREPCMB4',QEltDynamique.FindField('PHD_NEWVALEUR').AsString,False);

                      // Cases à cocher libres
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE1') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Oui';
                      end;
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE2') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Oui';
                      end;
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE3') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Oui';
                      end;
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE4') then
                      begin
                        if (QEltDynamique.FindField('PHD_NEWVALEUR').AsString = '-') then
                          Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Non'
                        else
                          Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Oui';
                      end;
                    end
                    else
                      result := '';
                    Ferme(QEltDynamique);
                  end;
                end;
                if (result = '') then
                begin
                  Requete := 'SELECT PSA_' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + ' FROM SALARIES' +
                    ' WHERE PSA_SALARIE = "' + Salarie + '"';
                  QSalarie := OpenSql(Requete, TRUE);
                  if not QSalarie.Eof then //PORTAGECWAS
                    if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE1') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE1') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE2') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE2') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE3') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE3') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'DATELIBRE4') then
                    begin
                      if (QSalarie.Fields[0].AsDateTime <> iDate1900) then
                        Result := GetParamsoc('SO_PGLIBDATE4') + ' : ' + DateTimeToStr(QSalarie.Fields[0].AsDateTime);
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE1') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE1') + ' : Oui';
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE2') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE2') + ' : Oui';
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE3') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE3') + ' : Oui';
                    end
                    else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'BOOLLIBRE4') then
                    begin
                      if (QSalarie.Fields[0].AsString = '-') then
                        Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Non'
                      else
                        Result := GetParamsoc('SO_PGLIBCOCHE4') + ' : Oui';
                    end
                    else
                    begin
                      if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB1') then
                        Result := GetParamsoc('SO_PGLIBCOMBO1') + ' : ' + RechDom('PGLIBREPCMB1',QSalarie.Fields[0].AsString,False)
                      else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB2') then
                        Result := GetParamsoc('SO_PGLIBCOMBO2') + ' : ' + RechDom('PGLIBREPCMB2',QSalarie.Fields[0].AsString,False)
                      else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB3') then
                        Result := GetParamsoc('SO_PGLIBCOMBO3') + ' : ' + RechDom('PGLIBREPCMB3',QSalarie.Fields[0].AsString,False)
                      else if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = 'LIBREPCMB4') then
                        Result := GetParamsoc('SO_PGLIBCOMBO4') + ' : ' + RechDom('PGLIBREPCMB4',QSalarie.Fields[0].AsString,False)
                      else
                        Result := QSalarie.Fields[0].AsString;
                    end;
                  Ferme(QSalarie);
                end;
              end;
            end;

            // -------- CUMUL PAIE --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'CMP') then
            begin
              // Déterminer la période de début et fin de prise en compte pour le cumul
              QRechCumul := OpenSql('SELECT PCL_RAZCUMUL FROM HISTOCUMSAL  ' +
              'LEFT JOIN CUMULPAIE ON PCL_CUMULPAIE=PHC_CUMULPAIE ' +
              'WHERE ##PCL_PREDEFINI## PHC_SALARIE="' + salarie + '" ' +
              'AND (PHC_CUMULPAIE="' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" )', True);
              if not QRechCumul.Eof then
              begin
                Raz := QRechCumul.FindField('PCL_RAZCUMUL').AsString;
                if Raz = '00' then
                  CumulDD := DateDebutExerSoc;
                if (raz = '01') or (raz = '02') or (raz = '03') or (raz = '04') or (raz = '05') or (raz = '06') or (raz = '07') or (raz = '08') or (raz = '09') or (raz = '10') or (raz = '11') or (raz = '12') then
                begin
                  DecodeDate(Datedeb, CpAnnee, CpMois, CpJour);
                  if CpMois < StrtoInt(raz) then Cpannee := Cpannee - 1;
                    CumulDD := EncodeDate(CpAnnee, StrToInt(raz), 1);
                end;
                if raz = '99' then cumulDD := DateJamais;
                if raz = '20' then cumulDD := DateCP;
                if (Datefin > 0) and (CumulDD > 0) and (Salarie <> '') and (Tob_Temp2.GetValue('PGP_PGVALCHAMP') <> '') then
                begin
                  // Récupérer le cumul sur la période déterminée
                  Requete := 'SELECT SUM(PHC_MONTANT) AS MONTANT ' +
                    'FROM HISTOCUMSAL WHERE PHC_SALARIE="' + Salarie + '" AND ' +
                    'PHC_CUMULPAIE="' + Tob_Temp2.GetValue('PGP_PGVALCHAMP') + '" ' +
                    'AND PHC_DATEDEBUT>="' + UsDateTime(CumulDD) + '" ' +
                    'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ';
                  QCumulMois := OpenSql(Requete, TRUE);
                  if not QCumulMois.eof then //PORTAGECWAS
                    Resultat := QCumulMois.FindField('MONTANT').AsFloat;
                  Ferme(QCumulMois);
                end;
              end;
              Ferme(QRechCumul);

              Tab_ZonesCalculees[strtoint(NumZone)] := Resultat;
              Result := RechDom('PGCUMULPAIE',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' : ' + FormatFloat('#,##0.00',Resultat);
            end;

            // -------- BASE, TAUX, COEFFICIENT, MONTANT  DE REMUNERATION --------
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'BAS') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_BASEREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (base) : ' + FormatFloat('#,##0.00',Result);
            end;
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'TAU') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_TAUXREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (taux) : ' + FormatFloat('#,##0.0000',Result);
            end;
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'COE') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_COEFFREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (coefficient) : ' + FormatFloat('#,##0.00',Result);
            end;
            if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'MNT') then
            begin
              Result := CalcZoneBasTauCoeMnt('PHB_MTREM',Etab,Salarie,Tob_Temp2.GetValue('PGP_PGVALCHAMP'),NumZone,DateDeb,DateFin);
              Result := RechDom('PGREMUNERATION',Tob_Temp2.GetValue('PGP_PGVALCHAMP'),False) + ' (montant) : ' + FormatFloat('#,##0.00',Result);
            end;
          end;
        end;

        // -------- RESULTAT --------
        if (Tob_Temp.GetValue('PGP_PGVALCHAMP') = 'RES') then
        begin
          // Récupération de la valeur de la zone à gauche de l'opérateur
          NomChamp := 'CBVALZONE' + NumZone;
          Tob_Temp2 := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
            Resultat1 := Tab_ZonesCalculees[StrToInt(Tob_Temp2.GetValue('PGP_PGVALCHAMP'))];

          // Récupération de la valeur de la zone à droite de l'opérateur
          NomChamp := 'CBVALZONE' + NumZone + 'B';
          Tob_Temp2 := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
            Resultat2 := Tab_ZonesCalculees[StrToInt(Tob_Temp2.GetValue('PGP_PGVALCHAMP'))];

          // Opération
          NomChamp := 'OPERATEUR' + NumZone;
          Tob_Temp2 := GblTob_ParamPop.FindFirst(['PGP_PGNOMCHAMP'], [NomChamp], False);
          if assigned(Tob_Temp2) then
          begin
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '+') then
              Resultat := Resultat1 + Resultat2;
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '-') then
              Resultat := Resultat1 - Resultat2;
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '*') then
              Resultat := Resultat1 * Resultat2;
            if (Tob_Temp2.GetValue('PGP_PGVALCHAMP') = '/') then
              if (Resultat2 <> 0) then
                Resultat := Resultat1 / Resultat2;
          end;
          Result := FormatFloat('#,##0.00',Resultat);
        end;
      end;
    end;
    exit;
  end;
end;

// Fonction de récupération d'une base, d'un taux, d'un coefficient ou d'un montant de rémunération
function CalcZoneBasTauCoeMnt(ChampPHB,Etab,Salarie,ValChamp,NumZone:String;DateDeb,DateFin:TDateTime):double;
var
  Requete : String;
  Q:TQuery;
begin
  Result := 0;
  Requete := 'SELECT ' + ChampPHB +
    ' FROM HISTOBULLETIN ' +
    ' WHERE PHB_NATURERUB="AAA" ' +
    ' AND PHB_ETABLISSEMENT = "' + Etab + '"' +
    ' AND PHB_SALARIE = "' + Salarie + '" ' +
    ' AND PHB_DATEDEBUT >= "' + UsDateTime(DateDeb) + '" ' +
    ' AND PHB_DATEFIN <= "' + UsDateTime(DateFin) + '" ' +
    ' AND PHB_RUBRIQUE = "' + ValChamp + '"';

  Q := OpenSql(Requete, TRUE);
  if not Q.eof then //PORTAGECWAS
  begin
    Result := Q.FindField(ChampPHB).AsFloat;
  end;
  Ferme(Q);

  Tab_ZonesCalculees[strtoint(NumZone)] := Result;
end;
//FIN PT4 *)

{***********A.G.L.***********************************************
Auteur  ...... : S. BELMAHJOUB
Créé le ...... : 24/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function PGCalcOLEGenericPaie(Sf, Sp: string): variant;
var
  StType: string;
begin
  Result := '';

  { Suppression des tobs }
  if sf = 'PAIEFREEFN' then
  begin
    Debug('Paie PGI : Déallocation memoire des données');
    StType := Trim(ReadTokenSt(sp));
    if StType = 'PFI' then
    begin
      FreeAndNil(GblTFicheInd);
      FreeAndNil(GblTCumPaie);
    end
    else
      if StType = 'PCU' then
      FreeAndNil(GblTCumPaie)
    else
      if StType = 'PRD' then
      FreeAndNil(GblTReduction);
(*PT8    else        //DEB PT4
      if StType = 'PBP' then
        if Sp = 'X' then
        begin
          FreeAndNil(GblTob_ParamPop);
          FreeAndNil(Tob_ExerciceSocial);
          FreeAndNil(Tob_DateClotureCp);
        end;      //FIN PT4 *)
  end
  else
    { Fiche individuelle }
    if (sf = 'PAIECREATEFID')
    or (sf = 'PAIEINITMONTRUB') or (sf = 'PAIEINITBASERUB')
    or (sf = 'PAIEGETMONTRUB') or (sf = 'PAIEGETBASERUB') or (sf = 'PAIEGETPATRUB')
    or (sf = 'PAIEGETCUMPAIE') or (Sf = 'PAIEGETETAB') or (Sf = 'PAIEGETMOIS') then
  begin
    result := FunctPGEditFicInd(Sf, Sp);
  end
  else
    { Cumul historique de paie }
    if (Sf = 'PAIECREATECUMSAL') or (Sf = 'PAIEINITMONTCUM') or (sf = 'PAIEGETMONTCUM') then
  begin
    result := FunctPgEditCumul(Sf, Sp);
  end
  else
    { Charges sociales }
    if (Sf = 'PAIEGETNBSEXE') then
  begin
    result := FunctPgEditChSocial(Sf, Sp);
  end
    { Réductions }
  else
    if (sf = 'PAIECHARGETOBREDUCTION') or (sf = 'PAIEGETVALREDUCTION') then
  begin
    result := FunctPgEditReduction(Sf, Sp);
  end;
  (*PT8
  //DEB PT4
    { Edition du bulletin }
  else
    if (Sf = 'PAIEGETVALZONE') or (Sf = 'PAIECREATEBULL') then
  begin
    result := FunctPgPaieEditBulletin(Sf, Sp);
  end;
  //FIN PT4*)

end; //FIN CalcOLEGenericPaie


////////////////////////////////////////////////////////////////////////////////

initialization
  GblTFicheInd := nil;
  GblTCumPaie := nil;

finalization
  if assigned(GblTFicheInd) then FreeAndNil(GblTFicheInd);
  if assigned(GblTCumPaie) then FreeAndNil(GblTCumPaie);

end.

