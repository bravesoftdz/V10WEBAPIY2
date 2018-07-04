{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 16/09/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFGBPGRAPHHPR ()
Modifié le ... : 29/01/2007
Description .. : Gestion d'une colonne supplémentaire : CA Presque réalisé
Mots clefs ... : TOF;QUFGBPGRAPHHPR
*****************************************************************}
Unit QUFGBPGRAPHHPR_TOF ;

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
  HCtrls
  ;
Type
  TOF_QUFGBPGRAPHHPR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;
                  
var QTHPR:TQTob;
    titreGraph:hString;

Implementation


procedure TOF_QUFGBPGRAPHHPR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPGRAPHHPR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPGRAPHHPR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPGRAPHHPR.OnLoad ;
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
  TheForm.Caption := titreGraph;
  //CEGID-CCMX le 20/11/2006 DEBUT
  if (QTHPR.LaTob.Detail[0].FieldExists('CAPRESQUEREAL')) then
  begin
    sColonnesListe := 'LADATE;CAHISTO;CAPREVU;CAPRESQUEREAL;CAREALISE';
    sTitres := 'Date;Historique;Prévu;PresqueRéal.;Réalisé';
    sColonnesGraph := 'CAHISTO;CAPREVU;CAPRESQUEREAL;CAREALISE';
  end
  else  //CEGID-CCMX le 20/11/2006 FIN
  begin
    sColonnesListe := 'LADATE;CAHISTO;CAPREVU;CAREALISE';
   sTitres := 'Date;Historique;Prévu;Réalisé';
    sColonnesGraph := 'CAHISTO;CAPREVU;CAREALISE';
  end;
  {$IFDEF MODE}  //EVI_TEMP_V800
  tsTitre := HTStringList.Create;
  {$ELSE}
  tsTitre := TStringList.Create;
  {$ENDIF}
  tsTitre.Add(titreGraph);
  stChampLigneTitre := 'LADATE';

  LanceGraph(TheForm, QTHPR.LaTob, '', sColonnesListe, '',
             sTitres, sColonnesGraph, '', tsTitre, nil,
             TLineSeries, stChampLigneTitre, false);

  TToolBarButton97(TFGRS1(Ecran).FindComponent('bAffGraph')).down := true;
  TToolBarButton97(TFGRS1(Ecran).FindComponent('bliste')).down := false;
  GetControl('FLISTE').Visible := False;
  tChart(GetControl('FCHART1')).Visible := True;
end ;

procedure TOF_QUFGBPGRAPHHPR.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPGRAPHHPR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPGRAPHHPR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFGBPGRAPHHPR.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_QUFGBPGRAPHHPR ] ) ; 
end.
