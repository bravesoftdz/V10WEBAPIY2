{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/05/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFGBPLOI ()
Mots clefs ... : TOF;QUFGBPLOI
*****************************************************************}
Unit QUFGBPLOI_TOF ;

Interface

Uses Controls,
  Classes,
  GraphUtil,
  UTOF,
  Chart,
  HTB97,
  GRS1,
  Series,
  UGraph,
  HEnt1
  ;

Type
  TOF_QUFGBPLOI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

var QTL:TQTob;

Implementation


procedure TOF_QUFGBPLOI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPLOI.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPLOI.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPLOI.OnLoad ;
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
  TheForm := TFGRS1(Ecran);
  sColonnesListe := 'LADATE;VAL1';
  sTitres := 'Date;%';
  sColonnesGraph := 'VAL1';
  {$IFDEF MODE}  //EVI_TEMP_V800
  tsTitre := HTStringList.Create;
  {$ELSE}
  tsTitre := TStringList.Create;
  {$ENDIF}
  tsTitre.Add(TraduireMemoire('% d''éclatement'));
  stChampLigneTitre := 'LADATE';

  LanceGraph(TheForm, QTL.LaTob, '', sColonnesListe, '',
             sTitres, sColonnesGraph, '', tsTitre, nil,
             TLineSeries, stChampLigneTitre, false); 
  
  TToolBarButton97(TFGRS1(Ecran).FindComponent('bAffGraph')).down := true;
  TToolBarButton97(TFGRS1(Ecran).FindComponent('bliste')).down := false;
  GetControl('FLISTE').Visible := False;
  tChart(GetControl('FCHART1')).Visible := True;
end ;

procedure TOF_QUFGBPLOI.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPLOI.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPLOI.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPLOI.OnCancel () ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_QUFGBPLOI ] ) ; 
end.
