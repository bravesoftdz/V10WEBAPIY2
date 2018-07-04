{***********UNITE*************************************************
Auteur  ...... : Paie
Créé le ...... : 15/02/2006
Modifié le ... : 15/02/2006
Description .. : Vignette des absences par motif (Camembert)
               : Vignette : PGVIGNETCHARTABS
               : Tablette : PGPERIODEVIGNETTE
               : Table    : ABSENCESALARIE
Mots clefs ... :
*****************************************************************
PT1  | 29/01/2008 | FLO | Ajout d'une sécurité sur les couleurs du camembert et plus de specif de la taille pour éviter l'étirement de l'image
}

unit AFVignetteChartAbsPaie;

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  dialogs,
  graphics,
//  mxcommon,
  uToolsOWC,
  OWC11_TLB,
  PGVignettePaie,
  PGVIGUTIL,
  HCtrls;

type
  TGChartAbsPAIE = class (TGAVignettePaie)
  private
    EnDateDu, EnDateAu: TDateTime;

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
  HEnt1,
  SysUtils;

{-----Lit les critères ------------------------------------------------------------------}

function TGChartAbsPAIE.GetInterface (NomGrille: string): Boolean;
var
  sN1        : string;
begin
  Result := inherited GetInterface ('');
  EnDateDu := iDate1900;
  EnDateAu := iDate2099;

  // Qd on lance le portail ParamFich = '', par défaut DateJour = date système
  // calcul des dates de début et fin de période
  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);


end;

{-----Critère de la requète ------------------------------------------------------------}

procedure TGChartAbsPAIE.GetClauseWhere;
begin
  inherited;
   ClauseWhere := ' WHERE (PCN_TYPEMVT="ABS" OR (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" and pcn_mvtduplique<>"X"))'+
                  ' AND (((PCN_DATEDEBUTABS >= "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEDEBUTABS <= "'+
                  USDATETIME (EnDateAu)+'"))'+
                  ' OR ((PCN_DATEFINABS >= "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEFINABS <= "'+
                  USDATETIME (EnDateAu)+'"))'+
                  ' OR ((PCN_DATEDEBUTABS < "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEFINABS > "'+
                  USDATETIME (EnDateAu)+'")))';
end;

{-----Chargement des données -----------------------------------------------------------}

procedure TGChartAbsPAIE.RecupDonnees;
var
  TobL, MaTob: TOB;
  i_ind: integer;
  St : String;
  Nbval : integer;
  NbAutres : double;
begin
  inherited;

  NbAutres := 0;
  Nbval := StrToInt(GetCritereValue('NBVAL'))  ;
  if (Nbval = 0) then
    Nbval := 6;

  st := 'select PCN_TYPECONGE, SUM (PCN_JOURS) LASOMME, PMA_PGCOLORACTIF from Absencesalarie '+
        'LEFT JOIN MOTIFABSENCE ON PMA_MOTIFABSENCE=PCN_TYPECONGE' ;
  try
    try
      MaTob := OpenSelectInCache (st + ClauseWhere+' group BY PCN_TYPECONGE,PMA_PGCOLORACTIF ORDER BY LASOMME DESC');

      if MaTob.Detail.Count < Nbval then
          Nbval := MaTob.Detail.Count  ;

      for i_ind := 0 to Nbval - 1 do
      begin
       TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
       St := RechDom ('PGMOTIFABSENCE', MaTob.detail [i_ind].GetString ('PCN_TYPECONGE'), FALSE);
       TobL.AddChampSupValeur ('TYPECONGE', st);
       TobL.AddChampSupValeur ('LENOMBRE', MaTob.detail [i_ind] .GetDouble ('LASOMME'));
       TobL.AddChampSupValeur ('COULEUR', MaTob.detail [i_ind] .GetString ('PMA_PGCOLORACTIF'));
      end;

      if MaTob.Detail.Count > Nbval then
      begin
       TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
       TobL.AddChampSupValeur ('TYPECONGE', 'Autres');
       TobL.AddChampSupValeur ('COULEUR', colNone);

       for i_ind := Nbval to MaTob.Detail.Count - 1 do
       begin
        NbAutres := NbAutres +  MaTob.detail [i_ind] .GetDouble ('LASOMME');
       end;
       TobL.AddChampSupValeur ('LENOMBRE', NbAutres);
      end;

    except
      on E: Exception do
        MessageErreur := 'Erreur lors du traitement des données avec le message :'#13#13 + E.Message;
    end;
  finally
    FreeAndNil (MaTob);
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure TGChartAbsPAIE.DrawGrid (Grille: string);
begin
  inherited;
end;

function TGChartAbsPAIE.SetInterface: Boolean;
var  Obj : TOWCChart;
begin
  result := TRUE;

  if TOBDonnees.detail.Count < 1 then exit;

  Obj := TOWCChart.Create(TobResponse, AspectGraph);
  try
    try

      TobToGraph(Obj);

      Obj.AddSpecific := ColorSpecific;

      Obj.TypeGraph := tyCamembert;
      //Obj.SetWidth (250); //PT1
      //Obj.SetHEight(250);
      Obj.owLegend.Position := poPositionLeft;
      Obj.owLabel.Visible := True;
      Obj.owLabel.Color := colLightYellow;

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

procedure TGChartAbsPAIE.TobToGraph(var Obj: TOWCChart);
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
    Obj.SetValue(TobDonnees.Detail[n].GetString('TYPECONGE'), 'Nbre jours', TobDonnees.Detail[n].GetDouble('LENOMBRE'));
end;

{---------------------------------------------------------------------------------------}
procedure TGChartAbsPAIE.ColorSpecific(oChart: chChart);
{---------------------------------------------------------------------------------------}
var
   i  : integer;
   nbpoints : integer;
begin
  if TobDonnees.Detail.Count > 0 then
  begin
    nbpoints := oChart.SeriesCollection[0].Points.count ;
    for i := 0 to  nbpoints-1 do
    begin
      if TobDonnees.Detail[i].GetString('COULEUR') <> '' Then //PT1
      	oChart.SeriesCollection[0].Points[i].Interior.Set_Color(StringToColor(TobDonnees.Detail[i].GetString('COULEUR')));
    end;
  end;
end;

end.
