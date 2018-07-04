{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... :   /  /
Description .. : Tableau de bord des effectifs (Vignette vierge)
               : Vignette : PG_VIG_EFFECTIFS
               : Tablette : PGPERIODEVIGNETTE
               : Table    : PAIEENCOURS, SALARIES
Mots clefs ... : 
*****************************************************************}
unit PGVIGEFFECTIFS;

interface
uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  PGVIGUTIL,
  uToolsOWC,
  uToolsPlugin,
  HCtrls;
type
  PG_VIG_EFFECTIFS = class (TGAVignettePaie)
 private
    sN1                        : string;
    DataTob,DataTobN1          : TOB;
    SalariesTob,SalariesTobN1  : TOB;
    EnDateDu, EnDateAu         : TDatetime;
    EnDateDuN1, EnDateAuN1     : TDatetime;
    PeriodeRech                : string;

    procedure AfficheChamps(AParam : string);

 protected
    procedure RecupDonnees; override;
    function GetInterface (NomGrille: string = ''): Boolean; override;
    procedure GetClauseWhere; override;
    procedure DrawGrid (Grille: string); override;
    function SetInterface : Boolean ; override;
 public
 end;
implementation
 uses
  SysUtils,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_EFFECTIFS.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');

  if (ParamFich = '') then DateJour := DateRef;

  PeriodeRech := GetControlValue('PERIODE')  ;
  if (PeriodeRech = '') then
  begin
    PeriodeRech := '003';                        // valeur par défaut = 003 (mois)
    SetControlValue ('PERIODE',PeriodeRech) ;
    DatesPeriode(Datejour, EnDateDu, EnDateAu, PeriodeRech,sN1);

    SetControlValue ('N1',sN1);
  end
  else
  begin
    sN1 := GetControlValue('N1');
  end;

  if ParamFich <> '' then
  // on fait glisser les périodes
  begin
    GlisserPeriode (ParamFich,DateJour, PeriodeRech);
  end;

  DatesPeriode(Datejour, EnDateDu, EnDateAu, PeriodeRech, SN1);
  SetControlValue('N1',sN1);

  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_EFFECTIFS.GetClauseWhere;
begin
  inherited;
     // période : An, Quadrimestre, Trimestre, Mois
     ClauseWhere := 'where ppu_datedebut >= "'+USDATETIME (EnDateDu)+'" and '+
     'ppu_datefin <= "'+USDATETIME (EnDateAu)+'"';

     if (PeriodeRech = '001') or (PeriodeRech='002')then
     // période : Semaine, mois
       ClauseWhere := 'where ppu_datedebut <= "'+USDATETIME (EnDateDu)+'" and '+
                      'ppu_datefin >= "'+USDATETIME (EnDateAu)+'"';

end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_EFFECTIFS.RecupDonnees;
var
  StSQL             : String;
  ClauseWhereN1     : string;
  ClauseWhereS      : string;
  TobL              : TOB;   //@@
begin
  inherited;
  try
     //ANNEE N

     // table PAIEENCOURS
     StSQL := 'select distinct (ppu_salarie),'+
              'psa_priseffectif, psa_unitepriseff, psa_dateentree, '+
              'psa_datesortie '+
              'from paieencours left join salaries on psa_salarie = ppu_salarie ';
     DataTob := OpenSelectInCache (StSQL+ClauseWhere);
//DataTob.SaveToFile ('C:\temp\DataTob.TXT',FALSE,TRUE, TRUE);

     // table SALARIES
     StSQL := 'select psa_salarie, psa_PRISEFFECTIF, psa_unitepriseff '+
              'from salaries ';
     ClauseWhereS := 'where psa_dateentree <= "'+USDATETIME (EnDateDu)+'" and '+
                     '(psa_datesortie = "'+USDATETIME (IDate1900)+'" or '+
                     'psa_datesortie >= "'+USDATETIME (EnDateAu)+'") and '+
                     'psa_priseffectif = "X"';
     SalariesTob := OpenSelectInCache (StSQL+ClauseWhereS);

     // ANNEE N-1
     EnDateDuN1 := PlusDate (EnDateDu, -1, 'A');
     EnDateAuN1 := PlusDate (EnDateAu, -1, 'A');


     // table PAIEENCOURS
     // période : An, Quadrimestre, Trimestre, Mois
     ClauseWhereN1 := 'where ppu_datedebut >= "'+USDATETIME (EnDateDuN1)+'" and '+
     'ppu_datefin <= "'+USDATETIME (EnDateAuN1)+'"';
     if (PeriodeRech = '001') or (PeriodeRech='002')then
     // période : Semaine, mois
       ClauseWhereN1 := 'where ppu_datedebut <= "'+USDATETIME (EnDateDuN1)+'" and '+
                      'ppu_datefin >= "'+USDATETIME (EnDateAuN1)+'"';
     StSQL := 'select distinct (ppu_salarie),'+
              'psa_priseffectif, psa_unitepriseff, psa_dateentree, '+
              'psa_datesortie '+
              'from paieencours left join salaries on psa_salarie = ppu_salarie ';
     DataTobN1 := OpenSelectInCache (StSQL+ClauseWhereN1);
//DataTobN1.SaveToFile ('C:\temp\DataTobN1.TXT',FALSE,TRUE, TRUE);

     // table SALARIES
     StSQL := 'select psa_salarie, psa_PRISEFFECTIF, psa_unitepriseff '+
              'from salaries ';
     ClauseWhereS := 'where psa_dateentree <= "'+USDATETIME (EnDateDuN1)+'" and '+
                     '(psa_datesortie = "'+USDATETIME (IDate1900)+'" or '+
                     'psa_datesortie >= "'+USDATETIME (EnDateAuN1)+'") and '+
                     'psa_priseffectif = "X"';
     SalariesTobN1 := OpenSelectInCache (StSQL+ClauseWhereS);

     if ((SalariesTobN1.Detail.Count <> 0) or
         (DataTobN1.Detail.Count <> 0) or
         (SalariesTob.Detail.Count <> 0) or
         (DataTob.Detail.Count <> 0)) then
     begin
       TobDonnees := Tob.Create ('LES EFFECTIFS', nil, -1);
       TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
       TobL.AddChampSupValeur ('EFFECTIF', 'effectifs');
     end;
     
  finally
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_EFFECTIFS.DrawGrid (Grille: string);
begin
  inherited;
end;

function PG_VIG_EFFECTIFS.SetInterface: Boolean;
begin
  inherited SetInterface;
  AfficheChamps('');
  result:=true;
end;


procedure PG_VIG_EFFECTIFS.AfficheChamps(AParam : string);
var
  NPEffPaye, NPEntree, NPSortie, NPPerman, NPEffTot            : double;
  NTPEffPaye, NTPEntree, NTPSortie, NTPPerman, NTPEffTot       : integer;
  N1PEffPaye, N1PEntree, N1PSortie, N1PPerman, N1PEffTot       : double;
  N1TPEffPaye, N1TPEntree, N1TPSortie, N1TPPerman, N1TPEffTot  : integer;
  i_ind                                                        : integer;
  EntreLe, SortiLe                                             : TDATEtime;
  Entre, Sorti, Permanent                                      : boolean;
  sANN1                                                        : string;

begin
  NPEffPaye := 0.0;
  NPEntree := 0.0;
  NPSortie := 0.0;
  NPPerman := 0.0;
  NPEffTot := 0.0;

  NTPEffPaye := 0;
  NTPEntree := 0;
  NTPSortie := 0;
  NTPPerman := 0;
  NTPEffTot := 0;

  N1TPEffPaye := 0;
  N1TPEntree := 0;
  N1TPSortie := 0;
  N1TPPerman := 0;
  N1TPEffTot := 0;

  N1PEffPaye := 0.0;
  N1PEntree := 0.0;
  N1PSortie := 0.0;
  N1PPerman := 0.0;
  N1PEffTot := 0.0;


  // ANNEE N
//  Entre := false;
//  Sorti := false;
//  Permanent := false;

  if (DataTob.detail.count <> 0) then
  begin
    for i_ind := 0 to DataTob.detail.count-1 do
    begin
      NTPEffPaye := NTPEffPaye + 1;
      EntreLe :=  DataTob.detail [i_ind].GetValue('PSA_DATEENTREE');
      SortiLe :=  DataTob.detail [i_ind].GetValue('PSA_DATESORTIE');
      if (EntreLe >= EnDateDu) and (EntreLe <= EnDateAu) then
      // salarié entré dans la période
      begin
        Entre := true;
        NTPEntree := NTPEntree+1;
      end
      else
      // salarié entré hors période période
        Entre := false;
      if (SortiLe >= EnDateDu) and (SortiLe <= EnDateAu) then
      // salarié sorti dans la période
      begin
        Sorti := true ;
        NTPSortie := NTPSortie+1;
      end
      else
      // salarié sorti hors période
        Sorti := false;

      if (EntreLe <= EnDateDu) and ((SortiLe >= EnDateAu) or (SortiLe = IDate1900)) then
      // salarié fait partie de l'effectif permanent pour la période
      begin
        permanent := true;
        NTPPerman := NTPPerman + 1;
      end
      else
        permanent := false;

      if (DataTob.detail [i_ind].GetValue('PSA_PRISEFFECTIF') = 'X') then
      // salarié pris dans l'effectif : calcul des effectifs payés
      begin
        NPEffPaye := NPEffPaye+DataTob.detail [i_ind].GetValue('PSA_UNITEPRISEFF');
        if entre then
        begin
          NPEntree := NPEntree + (1 * DataTob.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));
        end;
        if sorti then
        begin
          NPSortie := NPSortie + (1 * DataTob.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));
        end;
        if Permanent then
          NPPerman := NPPerman + (1 * DataTob.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));

      end;
    end;
  end;

  if (SalariesTob.detail.count <> 0) then
  begin
    NTPEffTot := SalariesTob.detail.count;
    for i_ind := 0 to SalariesTob.detail.count-1 do
    begin
      NPEffTot := NPEffTot + (1 * SalariesTob.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));
    end;
  end;

  // ANNEE N-1
//  Entre := false;
//  Sorti := false;
//  Permanent := false;
  if (DataTobN1.detail.count <> 0) then
  begin
    for i_ind := 0 to DataTobN1.detail.count-1 do
    begin
      N1TPEffPaye := N1TPEffPaye + 1;
      EntreLe :=  DataTobN1.detail [i_ind].GetValue('PSA_DATEENTREE');
      SortiLe :=  DataTobN1.detail [i_ind].GetValue('PSA_DATESORTIE');
      if (EntreLe >= EnDateDuN1) and (EntreLe <= EnDateAuN1) then
      // salarié entré dans la période
      begin
        Entre := true;
        N1TPEntree := N1TPEntree+1;
      end
      else
      // salarié entré hors période période
        Entre := false;
      if (SortiLe >= EnDateDuN1) and (SortiLe <= EnDateAuN1) then
      // salarié sorti dans la période
      begin
        Sorti := true ;
        N1TPSortie := N1TPSortie+1;
      end
      else
      // salarié sorti hors période
        Sorti := false;

      if (EntreLe <= EnDateDu) and ((SortiLe >= EnDateAu) or (SortiLe = IDate1900)) then
      // salarié fait partie de l'effectif permanent pour la période
      begin
        permanent := true;
        N1TPPerman := N1TPPerman + 1;
      end
      else
        permanent := false;

      if (DataTobN1.detail [i_ind].GetValue('PSA_PRISEFFECTIF') = 'X') then
      // salarié pris dans l'effectif : calcul des effectifs payés
      begin
        N1PEffPaye := N1PEffPaye + DataTobN1.detail [i_ind].GetValue('PSA_UNITEPRISEFF');
        if entre then
        begin
          N1PEntree := N1PEntree + (1 * DataTobN1.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));
        end;
        if sorti then
        begin
          N1PSortie := N1PSortie + (1 * DataTobN1.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));
        end;
        if Permanent then
          N1PPerman := N1PPerman + (1 * DataTobN1.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));

      end;
    end;
  end;
  if (SalariesTobN1.detail.count <> 0) then
  begin
    N1TPEffTot := SalariesTobN1.detail.count;
    for i_ind := 0 to SalariesTobN1.detail.count-1 do
    begin
      N1PEffTot := N1PEffTot + (1 * SalariesTobN1.detail [i_ind].GetValue('PSA_UNITEPRISEFF'));
    end;
  end;

  SetControlValue('NPEFFPAYE',NPEffPaye);
  SetControlValue('NPENTREE',NPEntree);
  SetControlValue('NPSORTIE',NPSortie);
  SetControlValue('NPPERMAN',NPPerman);
  SetControlValue('NPEFFTOT',NPEffTot);

  SetControlValue('NTPEFFPAYE',NTPEffPaye);
  SetControlValue('NTPENTREE',NTPEntree);
  SetControlValue('NTPSORTIE',NTPSortie);
  SetControlValue('NTPPERMAN',NTPPerman);
  SetControlValue('NTPEFFTOT',NTPEffTot);

  SetControlValue('ANNEEN',sN1);

  SetControlValue('N1PEFFPAYE',N1PEffPaye);
  SetControlValue('N1PENTREE',N1PEntree);
  SetControlValue('N1PSORTIE',N1PSortie);
  SetControlValue('N1PPERMAN',N1PPerman);
  SetControlValue('N1PEFFTOT',N1PEffTot);

  SetControlValue('N1TPEFFPAYE',N1TPEffPaye);
  SetControlValue('N1TPENTREE',N1TPEntree);
  SetControlValue('N1TPSORTIE',N1TPSortie);
  SetControlValue('N1TPPERMAN',N1TPPerman);
  SetControlValue('N1TPEFFTOT',N1TPEffTot);

  DatesPeriode(PlusDate (Datejour, -1, 'A'), EnDateDuN1, EnDateAuN1, PeriodeRech, sANN1);

  SetControlValue('ANNEEN1',SANN1);


  FreeAndNil (DataTob);
  FreeAndNil (DataTobN1);
  FreeAndNil (SalariesTob);
  FreeAndNil (SalariesTobN1);

end;

end.
