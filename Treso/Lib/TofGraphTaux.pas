{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GRAPHTAUX ()
Mots clefs ... : TOF;GRAPHTAUX
*****************************************************************}
unit TofGraphTaux ;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  SysUtils, HCtrls, GRS1, Series, GraphUtil,
  Graphics, TeEngine, Chart, UTOF, Constantes, Commun;

type

  TOF_GRAPHTAUX = Class (TOF)
    Taux	:	THValComboBox ;
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    DateDe	: THEdit ;
    DateA	: THEdit ;
    procedure InitEchelle;
  end ;

procedure TRLanceFiche_GraphTaux(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

implementation

uses
  HEnt1;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_GraphTaux(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_GRAPHTAUX.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  F : TFGRS1;
  sttitre, stColonnes,
  stChampTitre, stCriteres,
  stTitresCol1, stTitresCol2,
  stColGraph1, stColGraph2,
  stWhere                  : string;
  tstTitres                : Tstrings;
  ft                       : TFont;
begin
  inherited ;
  stCriteres := '';
  stColGraph1 := '';
  stColGraph2 := '';
  stTitresCol1 := '';
  stTitresCol2 := '';
  sttitre := '';

  if (Taux <> nil) then begin
    {Definition des fonts des titres}
    ft := TFont.Create;
    ft.Style := [fsBold];
    ft.Color := clBlue;
    ft.Size  := 12;

    tstTitres := TStringList.Create;
    F := TFGRS1(Ecran);

    tstTitres.Add ('Cours des taux ' + Taux.Text );
    stColonnes := 'TTA_DATE;TTA_COTATION';
    stTitre := 'Date;Cotation';            // légende des colonnes
    stColGraph1 := 'TTA_DATE;TTA_COTATION';
    stColGraph2 := '';
    stChampTitre := stTitre;
    stWhere := 'TTA_CODE="'+ Taux.Value + '"' ;
    stWhere := stWhere + ' AND TTA_DATE>="' + usdatetime(strToDate(dateDe.Text)) + '"' ;
    stWhere := stWhere + ' AND TTA_DATE<="' + usdatetime(strToDate(dateA.Text)) + '"' ;
    stWhere := stWhere + ' ORDER BY TTA_DATE' ;

    LanceGraph( F, nil, 'COTATIONTAUX', stColonnes, stWhere, stTitre, stColGraph1, stColGraph2,
                tstTitres, nil, TLineSeries , stChampTitre, False);

    if F.Fchart1.SeriesCount > 0 then begin
      F.FChart1.SeriesList.series[0].Marks.Visible := False;
      F.FChart1.SeriesList.series[0].DataSource := nil;
      F.FChart1.LeftAxis.Title.caption := Taux.Text;
      F.FChart1.BottomAxis.Title.caption := 'Date';
      F.FChart1.SeriesList.series[1].Title := F.FChart1.LeftAxis.Title.caption;
      {Pour n'afficher qu'une rubrique dans la légende : la série 0 est la date, on s'en moque dans la légende}
      F.FChart1.Legend.FirstValue := 1;
      F.FChart1.LeftAxis .LabelsFont.Color := clGreen;
      {Font des titres des axes du graph}
      F.FChart1.LeftAxis  .Title.Font.Assign(ft);
      F.FChart1.BottomAxis.Title.Font.Assign(ft);
      F.FChart1.LeftAxis  .Title.Font.Color := clGreen;

      ft.Size := 9;
      ft.Style := [];
      {Font des échelles des axes}
      F.FChart1.LeftAxis  .LabelsFont.Assign(ft);
      F.FChart1.BottomAxis.LabelsFont.Assign(ft);
      F.FChart1.LeftAxis  .LabelsFont.Color := clGreen;
      {Initialisation des échelles}
      InitEchelle;
    end ;
    FreeAndNil(ft);
    tstTitres.Free;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_GRAPHTAUX.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  SetControlVisible('BUNDO'  , False); {FQ 10128}
  SetControlVisible('BSAUVE' , False); {FQ 10128}
  SetControlVisible('BDELETE', False); {FQ 10128}

  Taux   := THValComboBox(Getcontrol('TAUX'));
  DateDe := THEdit(GetControl('CRITDATE')); {FQ 10129}
  DateA  := THEdit(GetControl('CRITDATE_'));{FQ 10129}
  if Taux <> nil then Taux.OnChange := TFGRS1(Ecran).BValiderClick;
  TFGRS1(Ecran).FListe.FColFormats[1] := '#,####0.0000';
  TFGRS1(Ecran).HelpContext := 150;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_GRAPHTAUX.InitEchelle;
{---------------------------------------------------------------------------------------}
var
  Tmp,
  MaxC,
  MinC : Double;
  Dbr  : TDblResult;
  Deb,
  Fin  : TDateTime;
  C    : TChart;
  Tx,
  Whe  : string;
begin
  C := TFGRS1(Ecran).FChart1;
  Tx  := Taux.Value;
  Fin := StrToDate(DateA.Text);
  Deb := StrToDate(DateDe.Text);

  {Récupération des valeurs extrèmes des fluctuation des Cotations}
  Whe := 'TTA_CODE = "' + Tx + '" AND TTA_DATE >= "' + USDateTime(Deb)
                             + '" AND TTA_DATE <= "' + USDateTime(Fin) + '"';

  Dbr := GetValeursMinMax(Tx, 'TTA_COTATION', 'COTATIONTAUX', Whe, Deb, Fin);
  MaxC := Valeur(Dbr.RC);
  MinC := Valeur(Dbr.RE);

  C.LeftAxis .Automatic := False;
  {Définition de l'échelle de l'axe de gauche/Série 1/Taux Réel/Jaune}
  Tmp := (MaxC - MinC) / 22;
  C.LeftAxis.SetMinMax(MinC - Tmp , MaxC + Tmp);
  C.LeftAxis.Increment := Tmp;
  C.LeftAxis.AxisValuesFormat := '#0.0000"';
end;


Initialization
  registerclasses ( [ TOF_GRAPHTAUX ] ) ;
end.

