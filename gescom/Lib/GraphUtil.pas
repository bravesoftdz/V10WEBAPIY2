unit GraphUtil;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HCtrls,UTob, TeEngine, GRS1, Series;

procedure LanceGraph (FGraph : TFGRS1; TobGraph : TOB;
                      stTable, stColonnes, stWhere : string;
                      stTitresCol, stColGraph1, stColGraph2 : string;
                      tstTitre1, tstTitre2 : Tstrings;
                      TCSCTypeChart : TChartSeriesClass;
                      stChampLigneTitre : string; AxeIsDate : Boolean);
function CompteLesChamps (stCol : string) : integer;
procedure SeriesGraph (stColGraph, stColGrid : string; TobGraph : TOB;
                       var iColGraph : array of integer);
procedure TitresColonnes (FGraph : TFGRS1; stTitresCol : string);

implementation

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Lance l'affichage du Graphe
{***********A.G.L.***********************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 20/06/2000
Modifié le ... :   /  /    
Description .. : Dessin d'un GRAPHE
Mots clefs ... : GRAPHE
*****************************************************************}
procedure LanceGraph (FGraph : TFGRS1; TobGraph : TOB;
                      stTable,  stColonnes, stWhere : string;
                      stTitresCol, stColGraph1, stColGraph2 : string;
                      tstTitre1, tstTitre2 : Tstrings;
                      TCSCTypeChart : TChartSeriesClass;
                      stChampLigneTitre : string; AxeIsDate : Boolean);
(*  FGraph : Form (TFGRS1) contenant les méthodes du graphe
    TobGraph : Données du Graph. Si nil, utilisation des params suivants
    stTable : Table sur laquelle se fait la séléction des données
    stColonnes : Champs de la table séparés d'un point virgule et champs
                 constituant la Grid
    stWhere : Condition de lecture des données
    stTitresCol : Titres des colonnes de la grid (séparés d'un point virgule)
    stColGraph1 : Noms des colonnes de la Grid générant le Graph1 (Gauche de la Forme)
    stColGraph2 : Noms des colonnes de la Grid générant le Graph2 (Droite de la Forme)
    tstTitre1 : Titre du Graphe1
    tstTitre2 : Titre du Graphe2
    TCSCTypeChart : Type de graphe par défaut
    stChampLigneTitre : Champ désignant la colonne des titres des séries
    AxeIsDate : inidique si les axes sont des dates *)
Var stSelect, stChamp, stCol : string;
    TSql : TQuery;
    TOBG : TOB;
    iColLigneTitre, iInd : Integer;
    iColGraph1, iColGraph2 : array of integer;
begin
FGraph.FListe.VidePile(False);
if TobGraph = nil then
    begin
    if (stTable = '') or (stColonnes = '') then exit;
    stSelect := '';
    stCol := stColonnes;
    stChamp := ReadTokenSt (stCol);
    While stChamp <> '' do
        begin
        if stSelect = '' then stSelect := 'SELECT ' else stSelect := stSelect + ',';
        stSelect := stSelect + stChamp;
        stChamp := ReadTokenSt (stCol);
        end;
    if stSelect = '' then exit;
    stSelect := stSelect + ' FROM ' + stTable;
    if stWhere <> '' Then stSelect := stSelect + ' WHERE ' + stWhere;
    TSql := OpenSql (stSelect, True,-1, '', True);
    if TSql.EOF then begin Ferme (TSql); exit; end;
    TobG := TOB.Create ('', nil, -1);
    TobG.LoadDetailDB ('', '', '', TSql, False);
    Ferme (TSql);
    end else
    begin
    TobG := TobGraph;
    end;
TobG.PutGridDetail (FGraph.FListe, True, False, stColonnes);

SetLength(iColGraph1, CompteLesChamps(stColGraph1));
SetLength(iColGraph2, CompteLesChamps(stColGraph2));

SeriesGraph (stColGraph1, stColonnes, TobG, iColGraph1);
SeriesGraph (stColGraph2, stColonnes, TobG, iColGraph2);

if stChampLigneTitre <> '' then
    begin
    if stColonnes = '' then
        begin
        iColLigneTitre := TobG.GetNumChamp(stChampLigneTitre) - 1;
        end else
        begin
        iColLigneTitre := 0;
        stCol := stColonnes;
        stChamp := ReadTokenSt (stCol);
        while (stChamp <> '') and (stChampLigneTitre <> stChamp) do
            begin
            Inc(iColLigneTitre);
            stChamp := ReadTokenSt(stCol);
            end;
        if stChamp = '' then iColLigneTitre := 0;
        end;
    end else iColLigneTitre := 0;

TitresColonnes (FGraph, stTitresCol);

FGraph.UpdateGraph (iColLigneTitre, 0, 1, TobG.Detail.Count, AxeIsDate,
                    iColGraph1, iColGraph2, TCSCTypeChart);

FGraph.FChart1.Title.Text.Clear;
FGraph.FChart2.Title.Text.Clear;

for iInd := 0 to tstTitre1.Count - 1 do
    begin
    FGraph.FChart1.Title.Text.Add (tstTitre1[iInd]);
    end;
if tstTitre2 <> nil then
    begin
    for iInd := 0 to tstTitre2.Count - 1 do
        begin
        FGraph.FChart2.Title.Text.Add (tstTitre2[iInd]);
        end;
    end;

if TobGraph = nil then TobG.Free;
end;

function CompteLesChamps (stCol : string) : integer;
var stChamp : string;
    iChamp : integer;
begin
iChamp := 0;
repeat
    stChamp := ReadTokenSt (stCol);
    if stChamp <> '' then Inc(iChamp);
until stChamp = '';
Result := iChamp;
end;

procedure SeriesGraph (stColGraph, stColGrid : string; TobGraph : TOB;
                       var iColGraph : array of integer);
var iChamp, iInd : integer;
    stChamp, stChampRech, stColRech : string;
begin
iChamp := 0;
stChamp := ReadTokenSt(stColGraph);
while stchamp <> '' do
    begin
    if stColGrid = '' then
        begin
        iColGraph[iChamp] := TobGraph.GetNumChamp(stChamp) - 1;
        end else
        begin
        iInd := 0;
        stColRech := stColGrid;
        stChampRech := ReadTokenSt(stColRech);
        while stChampRech <> '' do
            begin
            if stChampRech = stChamp then
                begin
                iColGraph[iChamp] := iInd;
                stChampRech := '';
                end else
                begin
                Inc(iInd);
                stChampRech := ReadTokenSt(stColRech);
                end;
            end;
        end;
    Inc(iChamp);
    stChamp := ReadTokenSt(stColGraph);
    end;
end;

procedure TitresColonnes (FGraph : TFGRS1; stTitresCol : string);
var iInd : integer;
    stTitre : string;
begin
iInd := 0;
stTitre := ReadTokenSt (stTitresCol);
While stTitre <> '' do
    begin
    FGraph.FListe.Cells [iInd, 0] := stTitre;
    Inc (iInd);
    stTitre := ReadTokenSt (stTitresCol);
    end;
end;

end.
