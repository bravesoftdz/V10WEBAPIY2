{***********UNITE*************************************************
Auteur  ...... : Paie
Créé le ...... : 15/02/2006
Modifié le ... : 15/02/2006
Description .. : Vignette Rémunérations  (histogramme 3D)
               : Vignette : PGVIGNETTECAHRTBR
               : Tablette : PGBRUTSVIGNETTE
               : Table    : PAIENCOURS
Mots clefs ... :
*****************************************************************}

unit AFVignetteChartBrutPaie;

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  uToolsOWC,
  OWC11_TLB,
  PGVignettePaie,
  HCtrls;

type
  TGChartBrutPaie = class (TGAVignettePaie)
  private
    DateDu, DateAu, DateDuN1, DateAuN1 : TDateTime;
    Annee, AnneeN1                      : string;
    MasseSalariale  : string;
  procedure TobToGraph(var Obj : TOWCChart);
  procedure ColorSpecific(oChart: chChart);

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

function TGChartBrutPaie.GetInterface (NomGrille: string): Boolean;
var
  Year, Month, Day                  : Word;
  SDate, SN1                        : string;
begin
  Result := inherited GetInterface ('MasseSalariale');



  if (ParamFich = '') then
  begin
    if (DateJour <= Idate1900) then
      DateJour := DateRef ;
    DecodeDate(DateJour,Year, month, Day);
  end;

  MasseSalariale := GetControlValue('TYPSALAIRE')  ;
  if (MasseSalariale = '') then
  begin
    MasseSalariale := '001' ;
    SetControlValue('TYPSALAIRE',MasseSalariale);
  end;

  // on fait glisser les périodes
  if ParamFich <> '' then
  begin
    if ParamFich = 'plus'  then
    begin
      DateJour := PlusDate (DateJour, +1, 'A');
    end
    else
    if ParamFich = 'moins' then
    begin
      DateJour := PlusDate (DateJour, -1, 'A');
    end;
    DecodeDate(DateJour,Year, month, Day);
  end;

  // affichage période
  // année N
  SDate := '01/01/'+InttoStr(Year);
  DateDu := StrToDate (SDate);
  SDate:= '31/12/'+InttoStr(Year);
  DateAu := StrToDate (SDate);
  Annee := InttoStr (Year);
  // année N-1
  DateDuN1 := PlusDate (DateDu, -1, 'A');
  DateAuN1 := PlusDate (DateDu, -1, 'A');
  DecodeDate(DateDuN1,Year, month, Day);
  AnneeN1 := InttoStr (Year);
  sN1 := Annee;

  SetControlValue ('N1',sN1);
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure TGChartBrutPaie.GetClauseWhere;
begin
  inherited;
   ClauseWhere := ' WHERE Year(PPU_DATEFIN) =  "'+annee+'" group by Year(PPU_DATEFIN),Month(PPU_DATEFIN)';
end;

{-----Chargement des données -----------------------------------------------------------}

procedure TGChartBrutPaie.RecupDonnees;
var
//@@  Q                                    : TQuery;
  TobL, MaTobI, MaTob, T               : TOB;
  i_ind, nomois, noAn                  : integer;
  St                                   : String;
  St1                                  : String;
  ClauseWhere1                         : string;

begin
  inherited;

  st := 'select Month(PPU_DATEFIN) MOIS, Year(PPU_DATEFIN) AN ,sum(PPU_CBRUT) BRUT, '+
        'sum(PPU_CCOUTPATRON) CHARGES, sum(PPU_CNETAPAYER) NET '+
        'from PAIEENCOURS ';

  st1 := 'UNION select Month(PPU_DATEFIN) MOIS, Year(PPU_DATEFIN) AN ,sum(PPU_CBRUT) BRUT, '+
        'sum(PPU_CCOUTPATRON) CHARGES, sum(PPU_CNETAPAYER) NET '+
        'from PAIEENCOURS ';

  ClauseWhere1 := ' WHERE Year(PPU_DATEFIN) =  "'+
                  anneeN1+'" group by Year(PPU_DATEFIN),Month(PPU_DATEFIN)'+
                  'ORDER BY Year(PPU_DATEFIN) DESC' ;

  try
    try
      MaTobI := OpenSelectInCache (st + ClauseWhere+ st1 + ClauseWhere1);    //@@

      MaTob:= TOB.Create('LaTOB',NIL,-1);
      for i_ind := 1 to 24 do
      begin
        if i_ind > 12 then
        begin
          Nomois := i_ind - 12;
          NoAn := strtoint(AnneeN1);
        end
        else
        begin
          Nomois := i_ind;
          NoAn := strtoint(Annee);
        end;

        T := TOB.create('TOBPAIEENCOURS',MaTob, -1);
        T.AddChampSupValeur('MOIS',nomois,false);
        T.AddChampSupValeur('AN',noAn,false);
        T.AddChampSupValeur('BRUT',0.0,false);
        T.AddChampSupValeur('CHARGES',0.0,false);
        T.AddChampSupValeur('NET',0.0,false);
      end;


      for i_ind := 0 to MaTobI.detail.count - 1 do
      begin
        if (MaTobI.detail [i_ind].GetVALUE ('AN') = Annee) then
        begin
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')-1].PutValue('MOIS', MaTobI.detail [i_ind].GetValue('MOIS'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')-1].PutValue('AN', MaTobI.detail [i_ind].GetValue('AN'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')-1].PutValue('BRUT', MaTobI.detail [i_ind].GetValue('BRUT'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')-1].PutValue('CHARGES', MaTobI.detail [i_ind].GetValue('CHARGES'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')-1].PutValue('NET', MaTobI.detail [i_ind].GetValue('NET'));
        end
        else
        begin
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')+11].PutValue('MOIS', MaTobI.detail [i_ind].GetValue('MOIS'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')+11].PutValue('AN', MaTobI.detail [i_ind].GetValue('AN'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')+11].PutValue('BRUT', MaTobI.detail [i_ind].GetValue('BRUT'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')+11].PutValue('CHARGES', MaTobI.detail [i_ind].GetValue('CHARGES'));
          MaTob.detail [MaTobI.detail [i_ind].GetInteger ('MOIS')+11].PutValue('NET', MaTobI.detail [i_ind].GetValue('NET'));
        end;
      end;
//MaTob.SaveToFile ('C:\temp\MATOB.TXT',FALSE,TRUE, TRUE);
      for i_ind := 0 to MaTob.detail.count - 1 do
      begin
       TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
       case (MaTob.detail [i_ind].GetInteger ('MOIS')) of
         1 :St := 'Janv.';
         2 :St := 'Fév.';
         3 :St := 'Mars';
         4 :St := 'Avril';
         5 :St := 'Mai';
         6 :St := 'Juin';
         7 :St := 'Juil.';
         8 :St := 'Août';
         9 :St := 'Sept.' ;
         10 :St := 'Oct.';
         11 :St := 'Nov.';
         12 :St := 'Déc.';
       end;

       TobL.AddChampSupValeur ('LEMOIS', St, CSTString);
       St := FloatToStr (MaTob.detail [i_ind].GetDouble ('AN'));
       TobL.AddChampSupValeur ('ANNEE', St, CSTString);

       if (MasseSalariale = '001') then
       begin
         // Brut (cumul 01)
         TobL.AddChampSupValeur ('SALAIRE', MaTob.detail [i_ind].GetDouble ('BRUT'));
       end
       else
       if (MasseSalariale = '002') then
       begin
         // Brut + Charges patronales (cumul 01 + cumul 07)
         TobL.AddChampSupValeur ('SALAIRE', MaTob.detail [i_ind].GetDouble ('BRUT')+MaTob.detail [i_ind].GetDouble ('CHARGES'));
       end
       else
       if (MasseSalariale = '003') then
       begin
         // net à payer
         TobL.AddChampSupValeur ('SALAIRE', MaTob.detail [i_ind].GetDouble ('NET'));
       end;
      end;

    except
      on E: Exception do
        MessageErreur := 'Erreur lors du traitement des données avec le message :'#13#13 + E.Message;
    end;
  finally
    FreeAndNIL(MaTob);
    FreeAndNIL(MaTobI);
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure TGChartBrutPaie.DrawGrid (Grille: string);
begin
  inherited;
end;

function TGChartBrutPaie.SetInterface: Boolean;
var  Obj : TOWCChart;
begin
  result := TRUE;

  if TOBDonnees.detail.Count <= 1 then exit;
  Obj := TOWCChart.Create(TobResponse,AspectGraph);
  try
    try

      TobToGraph(Obj);
//TObDonnees.SaveToFile ('C:\temp\TOBDONNEE.TXT',FALSE,TRUE, TRUE);
      Obj.AddSpecific := ColorSpecific;

      Obj.TypeGraph := tyHistogramme;
      Obj.owLegend.Position := poPositionBottom;
      Obj.owAxeBot.Bold := true;
      Obj.owAxeLeft.InfoAxe.NumFormat := '#,#0 €';
      Obj.Build;

    except
      on E : Exception do
        MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
    end;
  finally
    if Assigned(Obj) then FreeAndNil(Obj);
  end;
     result := TRUE;
end;

procedure TGChartBrutPaie.TobToGraph(var Obj: TOWCChart);
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
  begin
    Obj.SetValue(TobDonnees.Detail[n].GetString('LEMOIS'),
                 TobDonnees.Detail[n].GetString('ANNEE'),
                 TobDonnees.Detail[n].GetDouble('SALAIRE'));
  end
end;
{---------------------------------------------------------------------------------------}
procedure TGChartBrutPaie.ColorSpecific(oChart: chChart);
{---------------------------------------------------------------------------------------}
begin
  if TobDonnees.Detail.Count > 0 then begin
    if TobDonnees.Detail[0].GetString('ANNEE') = Annee then
    begin
      oChart.SeriesCollection[0].Interior.Set_Color('LightSteelBlue');
      oChart.SeriesCollection[1].Interior.Set_Color('LightGoldenrodYellow');
    end
    else
    begin
      oChart.SeriesCollection[0].Interior.Set_Color('LightGoldenrodYellow');
      oChart.SeriesCollection[1].Interior.Set_Color('LightSteelBlue');
    end;
  end;

end;

end.
