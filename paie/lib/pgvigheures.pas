{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... :   /  /    
Description .. : Heures travaillées et heures payée (Histogramme 3D)
               : Vignette : PG_VIG_HEURES
               : Tablette : PGPERIODEVIGNETTE
               : Table    : HISTOCUMSAL
Mots clefs ... : 
*****************************************************************}
unit PGVIGHEURES;

interface
uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  OWC11_TLB,
  PGVignettePaie,
  PGVIGUTIL,
  uToolsOWC,
  uToolsPlugin,
  HCtrls;

type
  PG_VIG_HEURES = class (TGAVignettePaie)
 private
    sN1                        : string;
    DataTob                    : TOB;
    EnDateDu, EnDateAu         : TDatetime;

    procedure TobToGraph(var Obj: TOWCChart);

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

function PG_VIG_HEURES.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, SN1);

  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_HEURES.GetClauseWhere;
begin
  inherited;
     // période : An, Quadrimestre, Trimestre, Mois
     ClauseWhere := 'and '+
                    'PHC_DATEDEBUT >= "'+USDATETIME (EnDateDu)+'" and '+
                    'PHC_DATEFIN <= "'+USDATETIME (EnDateAu)+'" ';
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_HEURES.RecupDonnees;
var
  StSQL             : string;
  StSQL1            : string;
  TobL              : TOB;

begin
  inherited;

  //ANNEE N

  // table HISTOCUMSAL
  StSQL := 'select sum(PHC_MONTANT) HEURES '+
           'from histocumsal where PHC_CUMULPAIE="20" ';
  StSQL1 := 'union all select sum(PHC_MONTANT) HEURES '+
           'from histocumsal where PHC_CUMULPAIE="21" ';
  ddWriteLN ('PAIE PG_VIG_HEURES.RecupDonnees : '+StSQL+ClauseWhere+StSQL1+ClauseWhere);
 try
   try
     DataTob := OpenSelectInCache (StSQL+ClauseWhere+StSQL1+ClauseWhere);

     TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
     TobL.AddChampSupValeur ('HEURES', '', CSTString);
     TobL.AddChampSupValeur ('TRAV', 'heures travaillées', CSTString);
     TobL.AddChampSupValeur ('PAYE', 'heures payées', CSTString);

     if DataTob.detail.Count <= 1 then
     begin
      TobL.AddChampSupValeur ('HTRAV', 0.0);
      TobL.AddChampSupValeur ('HPAYE', 0.0);
      exit;
     end
     else
     begin

       TobL.AddChampSupValeur ('HTRAV', DataTob.detail [0].GetDouble ('HEURES'));
       TobL.AddChampSupValeur ('HPAYE', DataTob.detail [1].GetDouble ('HEURES'));
//TobDonnees.SaveToFile ('C:\temp\TobDonnees.TXT',FALSE,TRUE, TRUE);
     end;
     ddWriteLN ('PAIE : Fin de RecupDonnees');

   except
     on E: Exception do
        MessageErreur := 'Erreur lors du traitement des données avec le message :'#13#13 + E.Message;
   end;

  finally
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_HEURES.DrawGrid (Grille: string);
begin
  inherited;
end;

function PG_VIG_HEURES.SetInterface: Boolean;
var
  Obj : TOWCChart;

begin
  inherited SetInterface;

  Obj := TOWCChart.Create(TobResponse,AspectGraph);
  try
    try

     TobToGraph(Obj);

//TObDonnees.SaveToFile ('C:\temp\TOBDONNEE.TXT',FALSE,TRUE, TRUE);

      Obj.TypeGraph := tyHistogramme;
      Obj.owLegend.Position := poPositionBottom;
      Obj.owAxeBot.Bold := true;
      Obj.owAxeLeft.InfoAxe.NumFormat := '#,#0';
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


procedure PG_VIG_HEURES.TobToGraph(var Obj: TOWCChart);
begin
      Obj.SetValue(TobDonnees.Detail[0].GetString('HEURES'),
                   TobDonnees.Detail[0].GetString('TRAV'),
                   TobDonnees.Detail[0].GetDouble('HTRAV'));
      Obj.SetValue(TobDonnees.Detail[0].GetString('HEURES'),
                   TobDonnees.Detail[0].GetString('PAYE'),
                   TobDonnees.Detail[0].GetDouble('HPAYE'));
end;

end.
