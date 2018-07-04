{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 07/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{
PT1   : 07/11/2005 SB V_65 FQ 12625 Procedure d'épurement des salariés chargés
                           en cas de filtre absence
PT2   : 28/02/2006 SB V_65 FQ 12926 Optimisation + Refonte tri du planning par
                           nom resp., puis nom sal.
PT3   : 27/06/2006 SB V_65 Refonte de l'export des compteurs du planning
PT4   :	06/07/2006 SB V_65 Nouveaux paramètres d'execution
PT5   : 07/09/2006 SB V_65 FQ 13394 Correction violation d'acces
PT6   : 14/05/2007 FC V_72 FQ 13758 Inclure le responsable dans le récap
                           planning
}
unit PgPlanningOutils;

interface
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  Classes,
  SysUtils,
  Messages,
  Graphics,
  HCtrls,
  HPlanning,
  HEnt1,
  hmsgbox,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Utob;

const
  SBWM_APPSTARTUP = WM_USER + 1;

type
  TNiveauRupture = record // ENREGISTREMENT DE TYPE NIVEAU DE RUPTURE
    NiveauRupt: 0..4; // Libellé niveau de rupture
    ChampsRupt: array[1..4] of string; // Nom de champs niveau de rupture
    CondRupt: string; // Clause conditionnel niveau de rupture
  end;

  TChampLibOrg = record
    ChampsOrg: array[1..4] of string;
    Champslibre: array[1..4] of string;
  end;

type
  TPlanningIntervalAffichage = (piaHeure, piaDemiJournee, piaSemaine, piaMois, piaTrimestre, piaSemestre, piaOutlook);
  TPlanningIntervalList = set of TPlanningIntervalAffichage ;

function ListeChampRupt(Sep: string; Rupt: TNiveauRupture): string;
procedure MiseEnFormeColonneFixeRupt ( HPlan : THPlanning ; Rupt : TNiveauRupture; NbColDefaut : integer = 2 ) ;
procedure LoadConfigPlanning(HPlan: THPlanning);
procedure MiseEnFormeTobRupt(TypeTob: string; Tob_DeportSal, TobRes: Tob; Rupt: TNiveauRupture);
function AddRessTob(T_DeportSal, TobRes: Tob; Rupt: TNiveauRupture): Boolean;
function CreateEnr(Champ: string; T_DeportSal, TobRes: Tob; Rupt: TNiveauRupture): Boolean;
procedure EpureTobRessource(TRes, TItems: Tob); {PT1}
function RendCritereSalarie(LcMultiNiveau: Boolean; LcTypUtilisat, LcStUtilisat, LcStNivHierar: string): string; { PT2 }
function RendListeService(StWhere: string): string; { PT2 }
procedure InitialiseChampLibOrg(var ChampLibOrg: TChampLibOrg); { PT3 }
Function PgGenereTobExport(TobSal, TobAbs, TobRecap: Tob; ChampCol,Titre,Mode: string; ReturnTob : Boolean = False) : Tob; { PT3 }
Function  RendIndexCol ( LaGrille : THGrid ; St : String) : Integer; { PT3 }

implementation

uses
  UTOFPGExportPlanning, Entpaie;

function ListeChampRupt(Sep: string; Rupt: TNiveauRupture): string;
var
  Prefixe: string;
begin
  if Sep = ';' then Prefixe := 'LIB_' else Prefixe := '';
  if Rupt.ChampsRupt[1] <> '' then Result := Prefixe + Rupt.ChampsRupt[1] + Sep;
  if Rupt.ChampsRupt[2] <> '' then Result := Result + Prefixe + Rupt.ChampsRupt[2] + Sep;
  if Rupt.ChampsRupt[3] <> '' then Result := Result + Prefixe + Rupt.ChampsRupt[3] + Sep;
  if Rupt.ChampsRupt[4] <> '' then Result := Result + Prefixe + Rupt.ChampsRupt[4] + Sep;
end;


procedure MiseEnFormeColonneFixeRupt(HPlan: THPlanning; Rupt: TNiveauRupture; NbColDefaut : integer = 2); { PT5 }
var
  FixeColDefaut, FixeAlignDefaut: string; { PT3 }
begin
   FixeColDefaut := '75;200'; { PT5 }
   FixeAlignDefaut := 'C;L'; { PT5 }
   
 if NbColDefaut = 3 then
    Begin
    FixeColDefaut := '75;200;75'; { PT3 }
    FixeAlignDefaut := 'C;L;L'; { PT3 }
    end;
  if Rupt.NiveauRupt = 0 then
  begin
    HPlan.TokenSizeColFixed := FixeColDefaut; 
    HPlan.TokenAlignColFixed := FixeAlignDefaut; 
  end
  else
    if Rupt.NiveauRupt = 1 then
    begin
      HPlan.TokenSizeColFixed := '150;' + FixeColDefaut;
      HPlan.TokenAlignColFixed := 'L;' + FixeAlignDefaut;
    end
    else
      if Rupt.NiveauRupt = 2 then
      begin
        HPlan.TokenSizeColFixed := '150;150;' + FixeColDefaut;
        HPlan.TokenAlignColFixed := 'L;L;' + FixeAlignDefaut;
      end
      else
        if Rupt.NiveauRupt = 3 then
        begin
          HPlan.TokenSizeColFixed := '150;150;150;' + FixeColDefaut;
          HPlan.TokenAlignColFixed := 'L;L;L;' + FixeAlignDefaut;
        end
        else
          if Rupt.NiveauRupt = 4 then
          begin
            HPlan.TokenSizeColFixed := '150;150;150;150;' + FixeColDefaut;
            HPlan.TokenAlignColFixed := 'L;L;L;L;' + FixeAlignDefaut;
          end;
end;

procedure LoadConfigPlanning(HPlan: THPlanning);
var
  S, ST: string;
  L: TStrings;
begin
  L := TStringList.Create;
  S := GetSynRegKey('ConfigPlanningConges', S, True);
  if S <> '' then
  begin
    L.Text := S;
    { Les couleurs Fond+Fériés}
    S := L.Strings[0];
    HPlan.ColorBackground := StringToColor(ReadTokenSt(S));
    HPlan.ColorJoursFeries := StringToColor(ReadTokenSt(S));
    HPlan.ColorSelection := StringToColor(ReadTokenSt(S));

    { La forme graphique }
    S := L.Strings[1];
    HPlan.FormeGraphique := AglPlanningStringToFormeGraphique(ReadTokenSt(S));

    { Hauteur,Largeur, cumuldate et mode de cumul }
    S := L.Strings[2];
    HPlan.RowSizeData := ValeurI(ReadTokenSt(S));
    HPlan.ColSizeData := ValeurI(ReadTokenSt(S));
    HPlan.ActiveLigneGroupeDate := ReadTokenSt(S) = '1';
    HPlan.CumulInterval := AglPlanningStringToCumulInterval(ReadTokenSt(S));
    St := ReadTokenSt(S);
    if St = '' then
      if (HPlan.ColSizeData = 15) and (HPlan.CumulInterval = pciMois) then St := 'dd' else St := 'dd/mm';
    HPlan.DateFormat := St;
  end;
  L.Free;
end;


procedure MiseEnFormeTobRupt(TypeTob: string; Tob_DeportSal, TobRes: Tob; Rupt: TNiveauRupture);
var
  I, J: Integer;
  St, Champ, Pref, Suf: string;
begin
  St := '';
  { DEB PT2 }
  if TypeTob = 'RES' then
  begin
    for I := 0 to Tob_DeportSal.Detail.Count - 1 do
    begin
      St := '';
      for j := 1 to Rupt.NiveauRupt do
      begin
        Champ := Rupt.champsRupt[j];
        if Pos('ETAB', Champ) > 0 then Pref := 'TT' else Pref := 'PG';
        Suf := Copy(Champ, 5, Length(Champ));
        Tob_DeportSal.Detail[i].AddChampSupValeur('TRI_' + Suf, RechDom(Pref + Suf, Tob_DeportSal.Detail[i].GetValue(Champ), False));
        St := St + 'TRI_' + Suf + ';';
      end;
    end;
    Tob_DeportSal.Detail.Sort(St + 'NOM_SAL;PSA_SALARIE');
    { FIN PT2 }
    for I := 0 to Tob_DeportSal.Detail.Count - 1 do
      AddRessTob(Tob_DeportSal.Detail[i], TobRes, Rupt);
  end;

end;


function AddRessTob(T_DeportSal, TobRes: Tob; Rupt: TNiveauRupture): Boolean;
var
  T1: Tob;
  ChRupt1, ValRupt1, ChRupt2, ValRupt2, ChRupt3, ValRupt3, ChRupt4, ValRupt4, RessSal: string;
begin
  result := False;
 // if Rupt.ChampsRupt[1]='PSE_RESPONSABS' then   { PT2 }
   // T_DeportSal.AddChampSupValeur('LIBELLE_RESP',RechDom('PGSALARIE',T_DeportSal.GetValue('PSE_RESPONSABS'),False)); { PT2 }
  while Result = False do
  begin
    RessSal := T_DeportSal.GetValue('PSA_SALARIE');
    if Rupt.NiveauRupt > 0 then //Niv1
    begin //B1
      ChRupt1 := Rupt.ChampsRupt[1];
      ValRupt1 := T_DeportSal.GetValue(ChRupt1);
      T1 := TobRes.FindFirst([ChRupt1], [ValRupt1], False); //ChampRupt1
      if T1 <> nil then
      begin
        if Rupt.NiveauRupt > 1 then //Niv2
        begin //B2
          ChRupt2 := Rupt.ChampsRupt[2];
          ValRupt2 := T_DeportSal.GetValue(ChRupt2);
          T1 := TobRes.FindFirst([ChRupt1, ChRupt2, 'COD_' + ChRupt2], [ValRupt1, ValRupt2, ValRupt2], False); //ChampRupt2
          if T1 <> nil then
          begin
            if Rupt.NiveauRupt > 2 then //Niv3
            begin //B3
              ChRupt3 := Rupt.ChampsRupt[3];
              ValRupt3 := T_DeportSal.GetValue(ChRupt3);
              T1 := TobRes.FindFirst([ChRupt1, ChRupt2, ChRupt3, 'COD_' + ChRupt3], [ValRupt1, ValRupt2, ValRupt3, ValRupt3], False); //ChampRupt3
              if T1 <> nil then
              begin
                if Rupt.NiveauRupt > 3 then //Niv4
                begin //B4
                  ChRupt4 := Rupt.ChampsRupt[4];
                  ValRupt4 := T_DeportSal.GetValue(ChRupt4);
                  T1 := TobRes.FindFirst([ChRupt1, ChRupt2, ChRupt3, ChRupt4, 'COD_' + ChRupt4], [ValRupt1, ValRupt2, ValRupt3, ValRupt4, ValRupt4], False); //ChampRupt4
                  if T1 <> nil then
                    result := CreateEnr('PSA_SALARIE', T_DeportSal, TobRes, Rupt) //Niv 4
                  else
                    result := CreateEnr(ChRupt4, T_DeportSal, TobRes, Rupt); //ChamptRupt4
                end //E4
                else
                  result := CreateEnr('PSA_SALARIE', T_DeportSal, TobRes, Rupt); //Niv 3
              end
              else
                result := CreateEnr(ChRupt3, T_DeportSal, TobRes, Rupt); //ChamptRupt3
            end //E3
            else
              result := CreateEnr('PSA_SALARIE', T_DeportSal, TobRes, Rupt); //Niv 2
          end
          else
            result := CreateEnr(ChRupt2, T_DeportSal, TobRes, Rupt); //ChamptRupt2
        end //E2
        else
          result := CreateEnr('PSA_SALARIE', T_DeportSal, TobRes, Rupt); //Niv1
      end
      else
        result := CreateEnr(ChRupt1, T_DeportSal, TobRes, Rupt); //ChamptRupt1
    end //E1
    else
      result := CreateEnr('PSA_SALARIE', T_DeportSal, TobRes, Rupt); //Niv 0
  end; //End While

end;

function CreateEnr(Champ: string; T_DeportSal, TobRes: Tob; Rupt: TNiveauRupture): Boolean;
var
  T1: Tob;
  Pref, Tablette: string;
begin
  result := False;
  T1 := Tob.Create('Les ressources', TobRes, -1);
  with T1 do
  begin
    AddChampSupValeur('PSA_SALARIE', '');
    AddChampSupValeur('NOM_SAL', '');
    AddChampSupValeur('NOM_SERVICE', ''); { PT3 }
    if Rupt.NiveauRupt >= 1 then
    begin
      AddChampSupValeur(Rupt.ChampsRupt[1], T_DeportSal.GetValue(Rupt.ChampsRupt[1]));
      AddChampSupValeur('COD_' + Rupt.ChampsRupt[1], '');
      AddChampSupValeur('LIB_' + Rupt.ChampsRupt[1], '');
    end;
    if Rupt.NiveauRupt >= 2 then
    begin
      AddChampSupValeur(Rupt.ChampsRupt[2], T_DeportSal.GetValue(Rupt.ChampsRupt[2]));
      AddChampSupValeur('COD_' + Rupt.ChampsRupt[2], '');
      AddChampSupValeur('LIB_' + Rupt.ChampsRupt[2], '');
    end;
    if Rupt.NiveauRupt >= 3 then
    begin
      AddChampSupValeur(Rupt.ChampsRupt[3], T_DeportSal.GetValue(Rupt.ChampsRupt[3]));
      AddChampSupValeur('COD_' + Rupt.ChampsRupt[3], '');
      AddChampSupValeur('LIB_' + Rupt.ChampsRupt[3], '');
    end;
    if Rupt.NiveauRupt >= 4 then
    begin
      AddChampSupValeur(Rupt.ChampsRupt[4], T_DeportSal.GetValue(Rupt.ChampsRupt[4]));
      AddChampSupValeur('COD_' + Rupt.ChampsRupt[4], '');
      AddChampSupValeur('LIB_' + Rupt.ChampsRupt[4], '');
    end;
    if Champ = 'PSA_SALARIE' then
    begin
      {if (Service <> '') and (Service <> 'Error') then
        Service := ' - ' + Service;  }
      PutValue('NOM_SAL', T_DeportSal.Getvalue('NOM_SAL'));
      PutValue('PSA_SALARIE', T_DeportSal.GetValue('PSA_SALARIE'));
      if T_DeportSal.GetValue('PSE_CODESERVICE') <> null then { PT3 }
        PutValue('NOM_SERVICE', RechDom('PGSERVICE', T_DeportSal.GetValue('PSE_CODESERVICE'), False)); { PT3 }

      Result := True;
    end
    else
    begin
      if Pos('ETAB', Champ) > 0 then Pref := 'TT' else Pref := 'PG';
      Tablette := Pref + Copy(Champ, 5, Length(Champ));
      PutValue('COD_' + Champ, T_DeportSal.GetValue(Champ));
      PutValue('LIB_' + Champ, RechDom(Tablette, T_DeportSal.GetValue(Champ), False));
    end;
  end;

  //IF T_DeportSal.GetValue('PSA_SALARIE')='0000000003' then
  //   PGVisuUnObjet(TobRessources,Champ,'');
end;

{ DEB PT1 }

procedure EpureTobRessource(TRes, TItems: Tob);
var TFille, T1: Tob;
  Sal: string;
begin
  TFille := TRes.FindFirst([''], [''], False);
  while assigned(TFille) do
  begin
    Sal := TFille.GetValue('PSA_SALARIE');
    if Sal <> '' then
    begin
      T1 := TItems.FindFirst(['PCN_SALARIE'], [Sal], False);
      if not Assigned(T1) then TFille.free;
    end;
    TFille := TRes.FindNext([''], [''], False);
  end;
end;
{ FIN PT1 }
{ DEB PT2 }

function RendCritereSalarie(LcMultiNiveau: Boolean; LcTypUtilisat, LcStUtilisat, LcStNivHierar: string): string;
var
  StListServ: string;
begin
  result := '';
  if LcStUtilisat <> '' then
  begin
    if LcTypUtilisat = 'ASS' then
    begin
      if LcMultiNiveau then
      begin
        StListServ := PGPlanningOutils.RendListeService('PSE_ASSISTABS="' + LcStUtilisat + '"');
        result := '(SELECT PSE_SALARIE FROM DEPORTSAL,SERVICEORDRE ' +
          'WHERE (PSE_CODESERVICE=PSO_CODESERVICE ' + LcStNivHierar +
          'AND (PSO_SERVICESUP IN (' + StListServ + ') ' +
          'OR PSO_CODESERVICE IN (' + StListServ + '))) ' +
          ' OR (PSE_SALARIE = "' + LcStUtilisat + '")'      //PT6
      end
      else
        result := '(SELECT PSE_SALARIE FROM DEPORTSAL ' +
          'WHERE PSE_ASSISTABS="' + LcStUtilisat + '" ' +
          ' OR PSE_SALARIE = "' + LcStUtilisat + '"';        //PT6
    end
    else
      if LcMultiNiveau then
      begin
        StListServ := RendListeService('PSE_RESPONSABS="' + LcStUtilisat + '"');
        result := '(SELECT PSE_SALARIE FROM DEPORTSAL,SERVICEORDRE ' +
          'WHERE (PSE_CODESERVICE=PSO_CODESERVICE ' + LcStNivHierar +
          'AND (PSO_SERVICESUP IN (' + StListServ + ') ' +
          'OR PSO_CODESERVICE IN (' + StListServ + '))) ' +
          ' OR (PSE_SALARIE = "' + LcStUtilisat + '")'      //PT6
      end
      else
        result := '(SELECT PSE_SALARIE FROM DEPORTSAL ' +
          'WHERE PSE_RESPONSABS="' + LcStUtilisat + '"' +
          ' OR PSE_SALARIE = "' + LcStUtilisat + '"';       //PT6
  end;
end;
{ FIN PT2 }

{ DEB PT2 }

function RendListeService(StWhere: string): string;
var Q: TQuery;
begin
  Result := '""';
  Q := OpenSql('SELECT DISTINCT PSE_CODESERVICE FROM DEPORTSAL WHERE ' + StWhere, True);
  if not Q.eof then Result := '';
  while not Q.Eof do
  begin
    Result := Result + '"' + Q.FindField('PSE_CODESERVICE').AsString + '",';
    Q.Next;
  end;
  Ferme(Q);
  if Length(Result) > 2 then Result := Copy(Result, 1, Length(Result) - 1);
end;
{ FIN PT2 }

{ DEB PT3 }
procedure InitialiseChampLibOrg(var ChampLibOrg: TChampLibOrg); 
begin
  ChampLibOrg.ChampsOrg[1] := VH_Paie.PGLibelleOrgStat1;
  ChampLibOrg.ChampsOrg[2] := VH_Paie.PGLibelleOrgStat2;
  ChampLibOrg.ChampsOrg[3] := VH_Paie.PGLibelleOrgStat3;
  ChampLibOrg.ChampsOrg[4] := VH_Paie.PGLibelleOrgStat4;
  ChampLibOrg.Champslibre[1] := VH_Paie.PgLibCombo1;
  ChampLibOrg.Champslibre[2] := VH_Paie.PgLibCombo2;
  ChampLibOrg.Champslibre[3] := VH_Paie.PgLibCombo3;
  ChampLibOrg.Champslibre[4] := VH_Paie.PgLibCombo4;
end;
{ FIN PT3 }

{ DEB PT3 }
Function PgGenereTobExport(TobSal, TobAbs, TobRecap: Tob; ChampCol,Titre,Mode: string; ReturnTob : Boolean = False) : Tob;
var
  Tob_Export, T, T1, T2: Tob;
  I, s: integer;
  Val,NbjoursPri, NbJoursRtt, NbJoursAutre: Double;
  Col, ColName,st, listechamp: string;
  TabChampRupt, TabValRupt: array[1..4] of string;
  ChampLibOrg: TChampLibOrg;
begin
  InitialiseChampLibOrg(ChampLibOrg);
  Tob_Export := Tob.Create('EXPORT_ABSENCE', nil, -1);
  col := '';
  ListeChamp := ChampCol;
  while ListeChamp <> '' do
  begin
    St := ReadTokenSt(ListeChamp);
    if St = 'PSA_SALARIE' then col := col + 'Matricule;'
    else if St = 'NOM_SAL' then col := col + 'Libellé;'
    else if St = 'NOM_SERVICE' then col := col + 'Service;'
    else if St = 'LIB_PSA_ETABLISSEMENT' then col := col + 'Etablissement;'
    else if St = 'LIB_PSE_RESPONSABS' then col := col + 'Responsable;'
    else if St = 'LIB_PSA_CODESTAT' then col := col + 'Code statistique;'
    else if Pos('PSA_TRAVAILN', St) > 0 then col := col + ChampLibOrg.ChampsOrg[StrToInt(Copy(St, Length(St), 1))] + ';'
    else if Pos('PSA_LIBREPCMB', St) > 0 then col := col + ChampLibOrg.Champslibre[StrToInt(Copy(St, Length(St), 1))] + ';';
  end;

  for i := 0 to TobSal.Detail.count - 1 do
  begin
    T1 := TobSal.detail[i];
    s := 0;
    ListeChamp := ChampCol;
    while ListeChamp <> '' do
    begin
      St := ReadTokenSt(ListeChamp);
      if (Pos('NOM_SAL', St) > 0) or (Pos('PSA_SALARIE', St) > 0) or (Pos('NOM_SERVICE', St) > 0) then Continue;
      Inc(s);   
      if (Pos('PSA_TRAVAILN', St) > 0) and (T1.GetValue(St) <> '') then
      begin
        TabChampRupt[s] := St; //ChampLibOrg.ChampsOrg[StrToInt(Copy(St, Length(St), 1))];
        TabValRupt[s] := T1.GetValue(St);
      end
      else
        if (Pos('PSA_LIBREPCMB', St) > 0) and (T1.GetValue(St) <> '') then
        begin
          TabChampRupt[s] := St; //ChampLibOrg.Champslibre[StrToInt(Copy(St, Length(St), 1))];
          TabValRupt[s] := T1.GetValue(St);
        end
        else
          if (T1.GetValue(St) <> '') then
          begin
            TabChampRupt[s] := St;
            TabValRupt[s] := T1.GetValue(St);
          end;
    end;
    if T1.GetValue('PSA_SALARIE') <> '' then
    begin
      T2 := Tob.Create('FILLE_EXPORT_ABSENCE', Tob_Export, -1);
      T2.Dupliquer(T1, True, True, True);
      T2.AddChampSupValeur('Matricule', T1.GetValue('PSA_SALARIE'));
      T2.AddChampSupValeur('Libellé', T1.GetValue('NOM_SAL'));
      T2.AddChampSupValeur('Service', T1.GetValue('NOM_SERVICE'));
      T2.AddChampSup('CPACQUISN1', False);
      T2.AddChampSup('CPPRISN1', False);
      T2.AddChampSup('CPRESTN1', False); { PT9-1 }
      T2.AddChampSup('CPACQUISN', False);
      T2.AddChampSup('CPPRISN', False);
      T2.AddChampSup('CPRESTN', False); { PT9-1 }
      T2.AddChampSup('CPENCOURS', False);
      T2.AddChampSup('CPSOLDE', False);
      T2.AddChampSup('RTTACQUIS', False);
      T2.AddChampSup('RTTPRIS', False);
      T2.AddChampSup('RTTREST', False); { PT9-1 }
      T2.AddChampSup('RTTENCOURS', False);
      T2.AddChampSup('RTTSOLDE', False);
      T2.AddChampSup('ABSENCE', False);
      for s := 1 to 4 do
        if TabChampRupt[s] <> '' then
        begin
          T2.AddChampSupValeur(TabChampRupt[s], TabValRupt[s]);
        end;
      //PT2-2 Ajout des mouvements en cours d'intégration dans calcul des compteurs
      //      Calcul des RTT en fonction du type
      //Congés payés mouvement non intégré et en cours d'intégration
      NbjoursPri := TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_TYPECONGE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'PRI', '-', 'ATT'], False);
      NbjoursPri := NbjoursPri + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_TYPECONGE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'PRI', 'ENC', 'ATT'], False);
      NbjoursPri := NbjoursPri + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_TYPECONGE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'PRI', '-', 'VAL'], False);
      NbjoursPri := NbjoursPri + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_TYPECONGE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'PRI', 'ENC', 'VAL'], False);
      if NbJoursPri > 0 then T2.PutValue('CPENCOURS', NbJoursPri);
      //Type RTT
      NbJoursRtt := TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PMA_TYPEABS', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'RTT', '-', 'ATT'], False); { PT5 }
      NbJoursRtt := NbJoursRtt + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PMA_TYPEABS', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'RTT', 'ENC', 'ATT'], False); { PT5 }
      NbJoursRtt := NbJoursRtt + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PMA_TYPEABS', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'RTT', '-', 'VAL'], False); { PT5 }
      NbJoursRtt := NbJoursRtt + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PMA_TYPEABS', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'RTT', 'ENC', 'VAL'], False); { PT5 }
      if NbJoursRtt > 0 then T2.PutValue('RTTENCOURS', NbJoursRtt);
      //autre absence
      NbJoursAutre := TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), '-', 'ATT'], False);
      NbJoursAutre := NbJoursAutre + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'ENC', 'ATT'], False);
      NbJoursAutre := NbJoursAutre + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), '-', 'VAL'], False);
      NbJoursAutre := NbJoursAutre + TobAbs.Somme('PCN_JOURS', ['PCN_SALARIE', 'PCN_EXPORTOK', 'PCN_VALIDRESP'], [T2.GetValue('PSA_SALARIE'), 'ENC', 'VAL'], False);
      NbJoursAutre := NbJoursAutre - NbjoursPri - NbJoursRtt;
      if NbJoursAutre > 0 then T2.PutValue('ABSENCE', NbJoursAutre);

      if TobRecap <> nil then
      begin
        T := TobRecap.FindFirst(['PRS_SALARIE'], [T2.GetValue('PSA_SALARIE')], False);
        if T <> nil then
        begin
          if T.GetValue('PRS_ACQUISN1') <> 0 then
            T2.PutValue('CPACQUISN1', T.GetValue('PRS_ACQUISN1'));
          if T.GetValue('PRS_PRISN1') <> 0 then
            T2.PutValue('CPPRISN1', T.GetValue('PRS_PRISN1'));
          if T.GetValue('PRS_RESTN1') <> 0 then { PT9-1 }
            T2.PutValue('CPRESTN1', T.GetValue('PRS_RESTN1'));
          if T.GetValue('PRS_ACQUISN') <> 0 then
            T2.PutValue('CPACQUISN', T.GetValue('PRS_ACQUISN'));
          if T.GetValue('PRS_PRISN') <> 0 then
            T2.PutValue('CPPRISN', T.GetValue('PRS_PRISN'));
          if T.GetValue('PRS_RESTN') <> 0 then { PT9-1 }
            T2.PutValue('CPRESTN', T.GetValue('PRS_RESTN'));
          if T.GetValue('PRS_CUMRTTACQUIS') <> 0 then
            T2.PutValue('RTTACQUIS', T.GetValue('PRS_CUMRTTACQUIS'));
          if T.GetValue('PRS_CUMRTTPRIS') <> 0 then
            T2.PutValue('RTTPRIS', T.GetValue('PRS_CUMRTTPRIS'));
          if T.GetValue('PRS_CUMRTTREST') <> 0 then { PT9-1 }
            T2.PutValue('RTTREST', T.GetValue('PRS_CUMRTTREST'));
        end;
      end;
      Val := 0;
      if isnumeric(T2.GetValue('CPRESTN1')) then Val := T2.GetValue('CPRESTN1');
      if isnumeric(T2.GetValue('CPRESTN')) then Val := Val + T2.GetValue('CPRESTN');
      if isnumeric(T2.GetValue('CPENCOURS')) then Val := Val - T2.GetValue('CPENCOURS');
      if Val <> 0 then T2.PutValue('CPSOLDE',Val);
      Val := 0;
      if isnumeric(T2.GetValue('RTTREST')) then Val := T2.GetValue('RTTREST');
      if isnumeric(T2.GetValue('RTTENCOURS')) then Val := Val - T2.GetValue('RTTENCOURS');
      if Val <> 0 then T2.PutValue('RTTSOLDE',Val);
    end;
  end;

  ColName := ChampCol + ';CPACQUISN1;CPPRISN1;CPRESTN1;CPACQUISN;CPPRISN;CPRESTN;CPENCOURS;CPSOLDE;RTTACQUIS;RTTPRIS;RTTREST;RTTENCOURS;RTTSOLDE;ABSENCE';
  Col := Col + 'CP Acq. N-1;CP Pris N-1;CP Rest N-1;CP Acq. N;CP Pris N;CP Rest N;CP en cours;Solde CP;RTT Acq;RTT Pris;RTT Rest;RTT en cours;Solde RTT;Absence'; { PT9-1 }

  St := Copy(ChampCol, 1, (Pos('PSA_SALARIE', ChampCol) - 1));
  St := St + 'NOM_SAL;PSA_SALARIE;';
  Tob_Export.Detail.Sort(St);
	{ DEB PT4 }
  if ReturnTob then
    Begin
    Result := Tob_Export;
    End
 else { FIN PT4 }
    Begin
    PGExportPlanning(Tob_Export, Titre, ColName, Col);
    FreeAndNil(Tob_Export);
    End;
end;
{ FIN PT3 }

{ DEB PT3 }
Function  RendIndexCol ( LaGrille : THGrid ; St : String) : Integer;
Var i : integer;
Begin
Result := -1;
For i := 0 to LaGrille.ColCount-1 do
  if LaGrille.ColNames [i] = UpperCase(St) then
    Begin
    result := i;
    Break;
    End;
if Result = -1 then PGibox('Nom de colonne incorrect : ' +St, '');
End;
{ FIN PT3 }

end.

