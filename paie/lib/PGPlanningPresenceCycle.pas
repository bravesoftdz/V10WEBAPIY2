{***********UNITE*************************************************
Auteur  ...... : GGU
Créé le ...... : 03/07/2007
Modifié le ... :   /  /
Description .. : Planning des présences et exceptions des
Suite ........ : cycles et modèles de cycles
Mots clefs ... : PLANNING;PRESENCE
*****************************************************************
PT1 GGU 03/07/2007 : Affichage d'une ligne avec les items
                     par Modeles en plus de la ligne des items par journées
}
unit PGPlanningPresenceCycle;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
{$IFDEF EAGLCLIENT}
  UtileAGL,eMul,MaineAgl,HStatus,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
  HCtrls, StdCtrls, HTB97, ExtCtrls, hplanning, UTOB, Dialogs,
  HEnt1, ComCtrls,
  PGPlanningTemplate, PgPresence,PgPlanningOutils;

type
  TPlanningPresenceCycle = class(TFormPlanning)
  private
    { Composants spécifiques pour le récap/détail }
    GBRecap: THGroupBox;
    LblTJType: THLabel;
    LblJType: THLabel;
    LblTPlage1de: THLabel;
    LblTPlage1a: THLabel;
    LblPlage1a: THLabel;
    LblPlage1de: THLabel;
    LblTPlage2de: THLabel;
    LblTPlage2a: THLabel;
    LblPlage2de: THLabel;
    LblPlage2a: THLabel;
    LblPlage1Lendemain: THLabel;
    LblPlage2Lendemain: THLabel;
    { Propriétés spécifiques }
    StWhTobEtats,StWhTobRessource, TypePlanning : string;
    GestionPresence : TGestionPresence;
    Procedure AskRefreshDetailsProcedure(Planning : THPlanning; Tob : TOB);
    Procedure SpecifPlanningProperties(Planning : THPlanning; Sender : TObject);
    procedure LoadEtats(Deb, Fin: TDateTime);
    procedure LoadItems(Deb, Fin: TDateTime);
    procedure LoadRessources(Deb, Fin: TDateTime);
    procedure GereOptionPopup(Item: TOB; ZCode: Integer;
      var ReDraw: Boolean; IdRessource : String; CurrentDate : TDateTime; ZoneSelectionnee : TZoneSelected);
  public
    { Déclarations publiques }
  end;

  procedure PGPlanningPresenceCycleOpen(DateDeb,DateFin : TDateTime ;WhereTobEtats, WhereTobRessource : string; ListDesIntervales : TPlanningIntervalList; TypeDuPlanning : String = 'CYC');   // NiveauRupt : TNiveauRupture;  ; DateReference : TDateTime = 0  ,StWhTobItems

implementation

uses DateUtils, TntStdCtrls, hmsgBox;

procedure PGPlanningPresenceCycleOpen(DateDeb,DateFin : TDateTime ;WhereTobEtats, WhereTobRessource : string; ListDesIntervales : TPlanningIntervalList; TypeDuPlanning : String = 'CYC');   // NiveauRupt : TNiveauRupture;  ; DateReference : TDateTime = 0  ,StWhTobItems
var
  Planning : TPlanningPresenceCycle;
Begin
  Planning := TPlanningPresenceCycle.Create(Application);
  { Ajout des composants pour le détail }
  with Planning do
  begin
    CreateDetailsGroupBox(GBRecap, 1, 1, 906, 73, 'Détail de la journée');
    CreateDetailsLabel(LblTJType,    GBRecap, 12 ,21,61,13,'Journée type');
    CreateDetailsLabel(LblJType,     GBRecap, 99 ,21,3 ,13,'', False);
    CreateDetailsLabel(LblTPlage1de, GBRecap, 414,21,51,13,'Plage 1 de');
    CreateDetailsLabel(LblTPlage1a,  GBRecap, 582,21,6 ,13,'à');
    CreateDetailsLabel(LblPlage1a,   GBRecap, 597,21,3 ,13,'', False);
    CreateDetailsLabel(LblPlage1de,  GBRecap, 477,21,3 ,13,'', False);
    CreateDetailsLabel(LblTPlage2de, GBRecap, 414,45,51,13,'Plage 2 de');
    CreateDetailsLabel(LblTPlage2a,  GBRecap, 582,45,6 ,13,'à');
    CreateDetailsLabel(LblPlage2de,  GBRecap, 477,45,3 ,13,'', False);
    CreateDetailsLabel(LblPlage2a,   GBRecap, 597,45,3 ,13,'', False);
    CreateDetailsLabel(LblPlage1Lendemain, GBRecap, 702,21,59,13,'le lendemain', True, False);
    CreateDetailsLabel(LblPlage2Lendemain, GBRecap, 702,45,59,13,'le lendemain', True, False);
  end;
  try
    { Affectation des évènements }
    with Planning do
    begin
      OnAskRefreshDetails  := AskRefreshDetailsProcedure;
      OnCreatePlanning     := SpecifPlanningProperties;
//      OnChargeTobPlanning  := ;
      OnPopup              := GereOptionPopup;
      LoadItemsProcedure   := LoadItems;
      LoadEtatsProcedure   := LoadEtats;
      LoadRessourcesProcedure:= LoadRessources;
      { Affectation des propriétés hérités }
      AllPlanIntervalDebut := DateDeb;
      AllPlanIntervalFin := DateFin;
      ListIntervales := ListDesIntervales;
      { Affectation des propriétés spécifiques }
      StWhTobEtats:=WhereTobEtats;
      StWhTobRessource:=WhereTobRessource;
      TypePlanning := TypeDuPlanning;
      GestionPresence := TGestionPresence.Create(False, False, True, True, False);
      if TypeDuPlanning = 'CYC' then
        Planning.Caption := 'Planning des rythmes de travail'
      else if TypeDuPlanning = 'MOD' then
        Planning.Caption := 'Planning des modèles de cycle';
      { Affichage du planning }
      ShowPlanning;
    end;
  finally
    if Assigned(Planning.GestionPresence) then FreeAndNil(Planning.GestionPresence);
    if Planning<>nil then Planning.Free;
  end;
end;

{ TPlanningPresenceCycle }

procedure TPlanningPresenceCycle.AskRefreshDetailsProcedure(Planning: THPlanning;
  Tob: TOB);
begin
  if Assigned(Tob) then
  begin
    LblJType.Caption    := Tob.GetString('ETAT')+' '+Tob.GetString('ETAT_LIBELLE');
    LblPlage1de.Caption := TimeToStr(Tob.GetDateTime('PJO_HORDEBPLAGE1'));
    LblPlage1a.Caption  := TimeToStr(Tob.GetDateTime('PJO_HORFINPLAGE1'));
    LblPlage1Lendemain.Visible := Tob.GetBoolean('PJO_JOURJ1PLAGE1');
    LblPlage2de.Caption := TimeToStr(Tob.GetDateTime('PJO_HORDEBPLAGE2'));
    LblPlage2a.Caption  := TimeToStr(Tob.GetDateTime('PJO_HORFINPLAGE2'));
    LblPlage2Lendemain.Visible := Tob.GetBoolean('PJO_JOURJ1PLAGE2');
  end else
  begin
    LblJType.Caption    := '';
    LblPlage1de.Caption := '';
    LblPlage1a.Caption  := '';
    LblPlage1Lendemain.Visible := False;
    LblPlage2de.Caption := '';
    LblPlage2a.Caption  := '';
    LblPlage2Lendemain.Visible := False;
  end;
end;

    { Chargement des Ressources }
procedure TPlanningPresenceCycle.LoadRessources(Deb, Fin : TDateTime);
Var
  StSQLLoadRessources : String;
  indexRessource : Integer;  //PT1
begin
    if TypePlanning = 'MOD' then //Affichage des modèles de cycle  (Type défini dans la tablette PGTYPEAFFECTSALARIE)
    begin
      if StWhTobRessource <> '' then StWhTobRessource := ' AND PMY_MODELECYCLE in ('+StWhTobRessource+')';
      StSQLLoadRessources := 'SELECT '
                           + ' PMY_MODELECYCLE as CYCLE, PMY_LIBELLE as LIBELLE, PMY_DATEVALIDITE, PMY_NBJOUR, PMY_COLORMODCYCLE'
                           + ' FROM MODELECYCLEENT '
                           + ' WHERE PMY_DATEVALIDITE <= "'+USDATETIME(AllPlanIntervalFin)+'"' + StWhTobRessource
                           + ' ORDER BY PMY_LIBELLE, PMY_MODELECYCLE';
    end else begin
      if StWhTobRessource <> '' then StWhTobRessource := ' AND PYC_CYCLE in ('+StWhTobRessource+')';
      StSQLLoadRessources := 'SELECT '
                           + ' PYC_CYCLE as CYCLE, PYC_LIBELLE as LIBELLE, PYC_DATEVALIDITE, PYC_TYPECYCLE,'
                           + ' PYC_ETATCYCLE, PYC_NBCYCLE, PYC_COLORCYCLE FROM CYCLEENTETE '
                           + ' WHERE PYC_DATEVALIDITE <= "'+USDATETIME(AllPlanIntervalFin)+'"' + StWhTobRessource
                           + ' ORDER BY PYC_LIBELLE, PYC_CYCLE';
    end;
    TobRessources.LoadDetailDBFromSQL('CYCLEENTETE',StSQLLoadRessources);
    { PT1 Duplication des ressources pour afficher une ligne avec les items par Modeles en plus des items par journées  }
    For indexRessource := TobRessources.FillesCount(0)-1 downto 0 do
    begin

    end;
end;

procedure TPlanningPresenceCycle.LoadEtats(Deb, Fin : TDateTime);
Var
  StSQLLoadEtats : String;
begin
  { Chargement des Etats }
  StSQLLoadEtats := 'SELECT PJO_JOURNEETYPE as ETAT, PJO_LIBELLE as ETAT_LIBELLE, PJO_HORDEBPLAGE1, '
                  + ' PJO_HORFINPLAGE1, PJO_DUREEPLAGE1, PJO_JOURJ1PLAGE1, PJO_TYPEJOUR1, '
                  + ' PJO_HORDEBPLAGE2, PJO_HORFINPLAGE2, PJO_DUREEPLAGE2, PJO_JOURJ1PLAGE2, '
                  + ' PJO_TYPEJOUR2, PJO_TYPEJOUR, PJO_EQUIVTPSPLEIN, PJO_JOURTRAVFERIE, '
                  + ' PJO_POIDSJOUR, PJO_TEMPSLIBRE1, PJO_TEMPSLIBRE2, PJO_TEMPSLIBRE3, '
                  + ' PJO_TEMPSLIBRE4, PJO_TEMPSLIBRE5, PJO_TEMPSLIBRE6, PJO_TEMPSLIBRE7, '
                  + ' PJO_DUREEPAUSE, PJO_PAIEMENTPAUSE, PJO_PAUSEEFFECTIF, PJO_DUREENOTRAV, '
                  + ' PJO_DUREENOTRAVEFF, PJO_ABREGE, PJO_COLORJOURTYPE'
                  + ' FROM JOURNEETYPE ';
  if StWhTobEtats <> '' then StSQLLoadEtats := StSQLLoadEtats +' WHERE PJO_JOURNEETYPE in ('+StWhTobEtats+')';
  TobEtats.LoadDetailDBFromSQL('JOURNEETYPE',StSQLLoadEtats);
  TobEtats.detail[0].AddChampSupValeur('FONTCOLOR',ColorToString(FontColor),true);
  TobEtats.detail[0].AddChampSupValeur('FONTSTYLE',FontStyle,true);
  TobEtats.detail[0].AddChampSupValeur('FONTSIZE',FontSize,true);
  TobEtats.detail[0].AddChampSupValeur('FONTNAME',FontName,true);
end;

procedure TPlanningPresenceCycle.LoadItems(Deb, Fin : TDateTime);
var
  TempTobRessources : Tob;
  indexRes : Integer;
  Cycle, JourneeType : String;
  DateCourante : TDateTime;
begin
  { Chargement des Items }
  for indexRes := 0 to TobRessources.FillesCount(0) -1 do
  begin
    TempTobRessources := TobRessources.Detail[indexRes];
    Cycle := TempTobRessources.GetString('CYCLE');
    DateCourante := Deb;
    While DateCourante <= Fin do
    begin
      { A chaque journée type, on ajoute un item depuis la date de début jusqu'à la date de fin }
      JourneeType := GestionPresence.GetJourneeType(DateCourante, Cycle, TypePlanning);
      if JourneeType <> '' then
        CreateNewPresenceItem(JourneeType, DateCourante, Cycle, 'CYCLE_OU_MODELECYCLE', TobEtats, TobItems, False);
      DateCourante := DateCourante +1 ;
    end;
  end;
end;

procedure TPlanningPresenceCycle.SpecifPlanningProperties(Planning: THPlanning;
  Sender: TObject);
begin
  with Planning do
  begin
    { Ressources }
    ResChampID := 'CYCLE';
    { Items }
    ChampLineID := 'CYCLE_OU_MODELECYCLE';
    ChampdateDebut := 'DATEDEBUT';
    ChampDateFin := 'DATEFIN';
    ChampEtat := 'ETAT';
    ChampLibelle := 'ETAT_LIBELLE';
    ChampHint := 'ETAT_LIBELLE';
    { Etats }
    EtatChampCode := 'ETAT';
    EtatChampLibelle := 'ETAT_LIBELLE';
    EtatChampBackGroundColor := 'PJO_COLORJOURTYPE';
    { PopupMenu }
    AddOptionPopup(99,'Gérer une exception');
    EnableOptionPopup(99,True);
    { Alignement des colonnes fixes }
    TokenSizeColFixed   := '50;200';
    TokenAlignColFixed  := 'C;L';
    { Liste des champs d'entete }
    TokenFieldColEntete := '';
    TokenFieldColFixed  := 'CYCLE;LIBELLE';   
  end;

end;

procedure TPlanningPresenceCycle.GereOptionPopup(Item: TOB; ZCode: Integer;
  var ReDraw: Boolean; IdRessource : String; CurrentDate : TDateTime; ZoneSelectionnee : TZoneSelected);
//Var
//  CurrentDate : TDateTime;
//  TempTob : Tob;
//  Cycle : String;
// d1, d2 : String;
// da1, da2 : TDateTime;
// i1, i2 : Integer;
// ZoneSel : TZoneSelected;
begin
  if (ZCode = 99) then  { Création d'une exception }
  begin
//    Cycle := '';
//    { On cherche la date et le cycle courants }
//    if not Assigned (Item) then
//    begin
//      CurrentDate := Plannings[PCPlannings.ActivePageIndex].GetDateOfCol(Plannings[PCPlannings.ActivePageIndex].Col);
//      TempTob := nil;
//      { En mode outlook }
//      if Plannings[PCPlannings.ActivePageIndex].ModeOutlook then
//      begin
//        if TobRessources.FillesCount(0) > 1 then
//          {Si on a plusieurs ressources, on ne sais pas laquelle choisir }
//          PGIInfo('Vous devez selectionner le cycle sur lequel l''exception doit être appliquée.')
//        else
//          {Si on a 1 seule ressource }
//          TempTob := TobRessources.Detail[0];
//      end else
//        TempTob := Plannings[PCPlannings.ActivePageIndex].GetRessourceOfRow(Plannings[PCPlannings.ActivePageIndex].Row);
//      if Assigned(TempTob) then
//        Cycle := TempTob.GetString('CYCLE');
//    end else begin
//      CurrentDate := Item.GetDateTime('JOUR');
//      Cycle := Item.GetString('CYCLE_OU_MODELECYCLE');
//    end;
    if (IdRessource <> '') and (CurrentDate <> iDate1900) then
    begin
      if GestionPresence.DateOfException(CurrentDate, IdRessource, TypePlanning) then
        AGLLanceFiche('PAY', 'EXCEPTCYCLE','',TypePlanning+';'+IdRessource+';'+DateToStr(CurrentDate),'ACTION=MODIFICATION')
      else
        AGLLanceFiche('PAY', 'EXCEPTCYCLE','','','ACTION=CREATION;PYA_TYPECYCLE='+TypePlanning+';PYA_CYCLE='+IdRessource+';PYA_DATEEXCEPTION='+DateToStr(CurrentDate));
      Plannings[PCPlannings.ActivePageIndex].Activate := False;
      GestionPresence.ReloadException;
      DelItems (CurrentDate, CurrentDate, 'JOUR');
      LoadItems(CurrentDate, CurrentDate);
      Plannings[PCPlannings.ActivePageIndex].Activate := True;
    end;
  end;
end;

end.
 