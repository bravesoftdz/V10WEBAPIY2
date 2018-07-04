{***********UNITE*************************************************
Auteur  ...... : RRO
Créé le ...... : 23/01/2002
Modifié le ... : 13/02/2002
Description .. : Source TOF de la FICHE : GRAPHCHANCEL (TR)
Suite ........ :
Suite ........ : Prend la fiche TRGRAPHCHANCEL
Mots clefs ... : TOF;GRAPHCHANCELL
*****************************************************************}
Unit TofGraphChancel ;

Interface

Uses
  StdCtrls, Controls, Classes, Graphics, TeEngine, Chart, Commun, Constantes, 
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  SysUtils, HCtrls, GRS1, Series, GraphUtil, UTOF, Grids, Math;

Type
  TOF_GRAPHCHANCEL = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    devise	: THValComboBox ;
    DateDe	: THEdit ;
    DateA	: THEdit ;
    FListe      : THGrid;

    procedure InitEchelle(NbDec : Byte);
    procedure DeviseOnChange(Sender : TObject);
  end ;

procedure TRLanceFiche_GraphChancel(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

implementation

uses
  HEnt1, ExtCtrls{TImage};

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_GraphChancel(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_GRAPHCHANCEL.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  F : TFGRS1;
  sttitre,
  stColonnes, stChampTitre,
  stCriteres, stWhere,
  stTitresCol1, stTitresCol2,
  stColGraph1, stColGraph2 : string;
  tstTitres                : Tstrings;
  ft                       : TFont;
  Dev                      : string;
  InfoDevise               : TDblResult;
  Quotite                  : string;
  Format                   : string;
  NbDecimales              : Byte;
  n                        : Integer;
begin
  Inherited ;
  stCriteres := '';
  stColGraph1 := '';
  stColGraph2 := '';
  stTitresCol1 := '';
  stTitresCol2 := '';
  sttitre := '';

  if (devise <> nil) and (DateDe <> nil) and (dateA <> nil) then begin
    tstTitres := TStringList.Create;
    {Definition des fonts des titres}
    ft := TFont.Create;
    ft.Style := [fsBold];
    ft.Color := clBlue;
    ft.Size  := 12;

    F := TFGRS1(Ecran);
    Dev := devise.Value;
    tstTitres.Add ('Cours de la devise ' + Devise.Text );
    stColonnes := 'H_DATECOURS;H_TAUXREEL;H_COTATION'; //;H_DEVISE;H_DATECOURS';
    stTitre := 'Date;Taux;Cotation';            // légende des colonnes
    stColGraph1 := 'H_DATECOURS;H_TAUXREEL;H_COTATION';
    stColGraph2 := '';
    stChampTitre := stTitre;
    stWhere := 'H_DEVISE = "' + Dev + '"' ;
    stWhere := stWhere + ' AND H_DATECOURS>="' + usdatetime(strToDate(dateDe.Text)) + '"' ;
    stWhere := stWhere + ' AND H_DATECOURS<="' + usdatetime(strToDate(dateA.Text)) + '"' ;
    stWhere := stWhere + ' ORDER BY H_DATECOURS';

    LanceGraph(F, nil, 'CHANCELL', stColonnes, stWhere , stTitre , stColGraph1,
                 stColGraph2, tstTitres, nil, TLineSeries , stChampTitre, False);

    if F.Fchart1.SeriesCount > 0 then begin
      {Récupération des infos de la devise pour formater le graph}
      InfoDevise := GetInfoDevise(Dev);
      Quotite    := InfoDevise.RC;
      NbDecimales := StrToInt(InfoDevise.RE);
      if NbDecimales > 0 then Format := Copy('x.xxxxxxxxxxxx', 0, NbDecimales + 1)
                         else Format := 'x';

      {????}
      F.FChart1.SeriesList.series[2].Marks.Visible := False;
      F.FChart1.SeriesList.series[0].Marks.Visible := False;
      F.FChart1.SeriesList.series[0].DataSource := nil;
      {Intitulés des axes du graph}
      n := Max(StrToInt(Quotite), 1);
      F.FChart1.LeftAxis.Title.Caption := IntToStr(n) + ' ' + Dev + ' = ' + Format + ' EUR';
      if StrToInt(Quotite) > 1 then
        F.FChart1.LeftAxis.Title.Caption := F.FChart1.LeftAxis  .Title.Caption + ' (/ ' + Quotite +')';

      F.FChart1.RightAxis.Title.Caption := '1 EUR = ' + format + ' ' + Dev;
      if StrToInt(Quotite) > 1 then
        F.FChart1.RightAxis.Title.Caption := F.FChart1.RightAxis .Title.Caption + ' (* ' + Quotite +')';

      F.FChart1.BottomAxis.Title.Caption := 'Date';
      {Font des titres des axes du graph}
      F.FChart1.LeftAxis  .Title.Font.Assign(ft);
      F.FChart1.RightAxis .Title.Font.Assign(ft);
      F.FChart1.BottomAxis.Title.Font.Assign(ft);

      ft.Size := 9;
      ft.Style := [];
      {Font des échelles des axes}
      F.FChart1.LeftAxis  .LabelsFont.Assign(ft);
      F.FChart1.RightAxis .LabelsFont.Assign(ft);
      F.FChart1.BottomAxis.LabelsFont.Assign(ft);
      {Pour la légende}
      F.FChart1.SeriesList.series[2].Title := F.FChart1.LeftAxis .Title.Caption;
      F.FChart1.SeriesList.series[1].Title := F.FChart1.RightAxis.Title.caption;
      F.FChart1.SeriesList.series[0].Title := '';
      {Pour n'afficher que deux rubriques dans la légende : la série 0 est la date, on s'en moque dans la légende}
      F.FChart1.Legend.FirstValue := 1;
      {Définition des échelles des axes verticaux, l'horizontal étant en mode automatique}
      InitEchelle(NbDecimales);
      {Attribution des axes aux séries}
      F.FChart1.SeriesList.series[1].VertAxis := aLeftAxis;
      F.FChart1.SeriesList.series[2].VertAxis := aRightAxis;
      F.FChart1.SeriesList.series[1].Active := True;
      {On met tout à la même couleur pour rendre le graph plus lisible
       Je n'ai pas encore trouvé comment modifié la couleur de la légende !!!!!}
      F.FChart1.SeriesList.series[1].ColorRange(F.FChart1.SeriesList.series[1].XValues, F.FChart1.SeriesList.series[1].XValues.First, F.FChart1.SeriesList.series[1].XValues.Last, clYellow);
      F.FChart1.SeriesList.series[2].ColorRange(F.FChart1.SeriesList.series[2].XValues, F.FChart1.SeriesList.series[2].XValues.First, F.FChart1.SeriesList.series[2].XValues.Last, clGreen);
      F.FChart1.RightAxis.LabelsFont.Color := clGreen;
      F.FChart1.LeftAxis .LabelsFont.Color := clYellow;
      F.FChart1.RightAxis.Title.Font.Color := clGreen;
      F.FChart1.LeftAxis .Title.Font.Color := clYellow;
    end ;
    tstTitres.Free;
    FreeAndNil(ft);
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_GRAPHCHANCEL.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  Ecran.HelpContext := 500008;
  SetControlVisible('BUNDO'  , False); {FQ 10128}
  SetControlVisible('BSAUVE' , False); {FQ 10128}
  SetControlVisible('BDELETE', False); {FQ 10128}

  devise := THValComboBox(Getcontrol('DEVISE'));
  DateDe := THEdit(GetControl('CRITDATE')); {FQ 10129}
  DateA  := THEdit(GetControl('CRITDATE_'));{FQ 10129}
  Devise.OnChange := TFGRS1(Ecran).BValiderClick;
  FListe := TFGRS1(Ecran).FListe;
  {On force le titre des colonne en ouverture de fiche car on se retrouve avec des titres surprenants !}
  FListe.Cells[0, 0] := 'Date';
  FListe.Cells[1, 0] := 'Taux';
  FListe.Cells[2, 0] := 'Cotation';
  FListe.FColFormats[1] := '#,####0.0000';
  FListe.FColFormats[2] := '#,####0.0000';
  TFGRS1(Ecran).FChart2.Visible := False;
  Devise.OnChange := DeviseOnChange;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_GRAPHCHANCEL.DeviseOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), Devise.Value);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_GRAPHCHANCEL.InitEchelle(NbDec : Byte);
{---------------------------------------------------------------------------------------}
var
  Tmp,
  MaxC,
  MinC,
  MaxT,
  MinT : Double;
  Dbr  : TDblResult;
  Deb,
  Fin  : TDateTime;
  Dev  : string;
  Whe  : string;
  C    : TChart;
begin
  C := TFGRS1(Ecran).FChart1;
  Dev := devise.Value;
  Fin := StrToDate(DateA.Text);
  Deb := StrToDate(DateDe.Text);

  Whe := 'H_DEVISE = "' + Dev + '" AND H_DATECOURS >= "' + USDateTime(Deb)
                              + '" AND H_DATECOURS <= "' + USDateTime(Fin) + '"';

  {Récupération des valeurs extrèmes des fluctuation des taux}
  Dbr := GetValeursMinMax(Dev, 'H_TAUXREEL', 'CHANCELL', Whe, Deb, Fin);
  MaxT := Valeur(Dbr.RC);
  MinT := Valeur(Dbr.RE);
  {Récupération des valeurs extrèmes des fluctuation des Cotations}
  Dbr := GetValeursMinMax(Dev, 'H_COTATION', 'CHANCELL', Whe, Deb, Fin);
  MaxC := Valeur(Dbr.RC);
  MinC := Valeur(Dbr.RE);

  C.LeftAxis .Automatic := False;
  C.RightAxis.Automatic := False;
  {Définition de l'échelle de l'axe de gauche/Série 1/Taux Réel/Jaune}
  Tmp := (MaxC - MinC) / 22;
  C.RightAxis.SetMinMax(MinC - Tmp , MaxC + Tmp);
  C.RightAxis.Increment := Tmp;
  C.RightAxis.AxisValuesFormat := '#0.0000"';
  {Définition de l'échelle de l'axe de droite/Série 2/Cotation/Vert}
  Tmp := (MaxT - MinT) / 22;
  C.LeftAxis.SetMinMax(MinT - Tmp, MaxT + Tmp);
  C.LeftAxis.Increment := Tmp;
  C.LeftAxis.AxisValuesFormat := '#0.0000"';
end;

initialization
  RegisterClasses ( [ TOF_GRAPHCHANCEL ] ) ;

end.

