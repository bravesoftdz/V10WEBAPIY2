{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/05/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFGBPHPR ()
Mots clefs ... : TOF;QUFGBPHPR
*****************************************************************}
Unit QUFGBPHPR_TOF ;

Interface

Uses UTOF,UGraph;

Type
  TOF_QUFGBPHPR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
     QT: TQTob;
  end ;

Implementation

Uses Classes,GraphUtil,Chart,HTB97,GRS1,Series;

procedure TOF_QUFGBPHPR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPHPR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPHPR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPHPR.OnLoad ;
var sColonnesListe, stChampLigneTitre: hString;
    sColonnesGraph: hString;
    sTitres : hstring;
    {$IFDEF MODE}  //EVI_TEMP_V800
    tsTitre: HTStrings;
    {$ELSE}
    tsTitre: TStrings;
    {$ENDIF}
    TheForm: TFGRS1;
begin
  Inherited ;
  //création de la tob
  if QT <> nil
    then QT.Free;

  QT :=TQTob.create(['LADATE'],['VAL1']);

  TheForm := TFGRS1(Ecran);
  sColonnesListe := 'LADATE;VAL1;VAL2;VAL3';
  sTitres := 'Date;Valeur';
  sColonnesGraph := 'VAL1;VAL2;VAL3';
  {$IFDEF MODE}  //EVI_TEMP_V800
  tsTitre := HTStringList.Create;
  {$ELSE}
  tsTitre := TStringList.Create;
  {$ENDIF}
  tsTitre.Add('% eclatement');
  stChampLigneTitre := 'LADATE';
  LanceGraph(TheForm, QT.LaTob, '', sColonnesListe, '',
             sTitres, sColonnesGraph, '', tsTitre, nil,
             TLineSeries, stChampLigneTitre, false);

  TToolBarButton97(TFGRS1(Ecran).FindComponent('bAffGraph')).down := true;
  TToolBarButton97(TFGRS1(Ecran).FindComponent('bliste')).down := false;
  GetControl('FLISTE').Visible := False;
  tChart(GetControl('FCHART1')).Visible := True;
end ;

procedure TOF_QUFGBPHPR.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPHPR.OnClose ;
begin
  Inherited ; 
  QT.free;
end ;

procedure TOF_QUFGBPHPR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPHPR.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_QUFGBPHPR ] ) ; 
end.
