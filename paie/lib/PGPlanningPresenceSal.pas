{***********UNITE*************************************************
Auteur  ...... : GGU
Créé le ...... : 03/07/2007
Modifié le ... :   /  /
Description .. : Planning des absences/présences et exceptions des
Suite ........ : salariés
Mots clefs ... : PLANNING;PRESENCE
*****************************************************************}
{
PT1 25/10/2007 GGU V_80 Ajout des uinformations concernant la journée lors d'un clique sur le planning
}
unit PGPlanningPresenceSal;
          
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
  TPlanningPresenceSal = class(TFormPlanning)
  private
    { Composants spécifiques pour le récap/détail }
    GBRecap: THGroupBox;
    TobJourneesTypes : Tob;//PT1
    MiHighAnnule : THMenuItem;
    MiHighAbs    : THMenuItem;
    MiHighPre    : THMenuItem;
    MiHighExcept : THMenuItem;
    MiHighCycle  : THMenuItem;
    MiHighModele : THMenuItem;
    DetailBox : THListBox;
    { Propriétés spécifiques }
    StWhTobEtats,StWhTobRessource : string;
    stWhereSalarie : String;
    GestionPresence : TGestionPresence;
    Realise, previsionnel : Boolean;
    Rupture: TNiveauRupture;
    Procedure AskRefreshDetailsProcedure(Planning : THPlanning; laTob : TOB);
    Procedure SpecifPlanningPropertys(Planning : THPlanning; Sender : TObject);
    procedure LoadEtats(Deb, Fin: TDateTime);
    procedure LoadItems(Deb, Fin: TDateTime);
    procedure LoadRessources(Deb, Fin: TDateTime);
    procedure OnClickHighlight(Sender : TObject);
    procedure BeforeShowForm(Sender : TObject);
    procedure GereOptionPopup(Item: TOB; ZCode: Integer;
      var ReDraw: Boolean; IdRessource : String; CurentDate : TDateTime; ZoneSelectionnee : TZoneSelected);
  public
    { Déclarations publiques }
  end;

  procedure PGPlanningPresenceSalOpen(DateDeb,DateFin : TDateTime ;WhereTobEtats, WhereTobRessource : string; NiveauRupt: TNiveauRupture; ListDesIntervales : TPlanningIntervalList; ShowRealise, ShowPrevisionnel : Boolean; WhereSalarie : String = '');   // NiveauRupt : TNiveauRupture;  ; DateReference : TDateTime = 0  ,StWhTobItems

implementation

uses DateUtils, TntStdCtrls, hmsgBox, StrUtils, ed_tools, Menus;

procedure PGPlanningPresenceSalOpen(DateDeb,DateFin : TDateTime ;WhereTobEtats, WhereTobRessource : string; NiveauRupt: TNiveauRupture; ListDesIntervales : TPlanningIntervalList; ShowRealise, ShowPrevisionnel : Boolean; WhereSalarie : String = '');   // NiveauRupt : TNiveauRupture;  ; DateReference : TDateTime = 0  ,StWhTobItems
var
  Planning : TPlanningPresenceSal;
  TobDesJourneesTypes : Tob;
Begin
  Planning := TPlanningPresenceSal.Create(Application);
  //Debut PT1
  Planning.AskRefreshDetailsForAllClick := True;
  TobDesJourneesTypes := Tob.Create('Tob des journées types', nil, -1);
  TobDesJourneesTypes.LoadDetailFromSQL('SELECT PJO_JOURNEETYPE, PJO_LIBELLE, PJO_HORDEBPLAGE1, '
                                      + ' PJO_HORFINPLAGE1, PJO_DUREEPLAGE1, PJO_JOURJ1PLAGE1, PJO_TYPEJOUR1, '
                                      + ' PJO_HORDEBPLAGE2, PJO_HORFINPLAGE2, PJO_DUREEPLAGE2, PJO_JOURJ1PLAGE2, '
                                      + ' PJO_TYPEJOUR2, PJO_TYPEJOUR, PJO_EQUIVTPSPLEIN, PJO_JOURTRAVFERIE, '
                                      + ' PJO_POIDSJOUR, PJO_TEMPSLIBRE1, PJO_TEMPSLIBRE2, PJO_TEMPSLIBRE3, '
                                      + ' PJO_TEMPSLIBRE4, PJO_TEMPSLIBRE5, PJO_TEMPSLIBRE6, PJO_TEMPSLIBRE7, '
                                      + ' PJO_DUREEPAUSE, PJO_PAIEMENTPAUSE, PJO_PAUSEEFFECTIF, PJO_DUREENOTRAV, '
                                      + ' PJO_DUREENOTRAVEFF, PJO_ABREGE, PJO_COLORJOURTYPE'
                                      + ' FROM JOURNEETYPE ');
  Planning.TobJourneesTypes := TobDesJourneesTypes;
  //Fin PT1
  { Ajout des composants pour le détail }
  with Planning do
  begin
    CreateDetailsGroupBox(GBRecap, 1, 1, 906, 73, 'Détail de la journée');
    DetailBox := THListBox.Create(Planning);
    DetailBox.Parent := GBRecap;
    DetailBox.Align := alClient;
    MiHighAnnule    := THMenuItem.Create(Planning);
    MiHighAnnule.Caption := 'Enlever la mise en évidence';
    MiHighAnnule.OnClick := OnClickHighlight;
    MiHighAbs    := THMenuItem.Create(Planning);
    MiHighAbs.Caption := 'Mise en évidence des absences';
    MiHighAbs.OnClick := OnClickHighlight;
    MiHighPre    := THMenuItem.Create(Planning);
    MiHighPre.Caption := 'Mise en évidence des présences';
    MiHighPre.OnClick := OnClickHighlight;
    MiHighExcept := THMenuItem.Create(Planning);
    MiHighExcept.Caption := 'Mise en évidence des exceptions';
    MiHighExcept.OnClick := OnClickHighlight;
    MiHighCycle  := THMenuItem.Create(Planning);
    MiHighCycle.Caption := 'Mise en évidence des cycles de présence';
    MiHighCycle.OnClick := OnClickHighlight;
    MiHighModele  := THMenuItem.Create(Planning);
    MiHighModele.Caption := 'Mise en évidence des modèles de cycle de présence';
    MiHighModele.OnClick := OnClickHighlight;
  end;
  try
    { Affectation des évènements }
    with Planning do
    begin
      OnAskRefreshDetails  := AskRefreshDetailsProcedure;
      OnCreatePlanning     := SpecifPlanningPropertys;
      OnBeforeShow         := BeforeShowForm;
      OnPopup              := GereOptionPopup;
      LoadItemsProcedure   := LoadItems;
      LoadEtatsProcedure   := LoadEtats;
      LoadRessourcesProcedure:= LoadRessources;
      { Affectation des propriétés hérités }
      AllPlanIntervalDebut := DateDeb;
      AllPlanIntervalFin := DateFin;
      ListIntervales := ListDesIntervales;
      { Affectation des propriétés spécifiques }
      StWhTobEtats     := WhereTobEtats;
      StWhTobRessource := WhereTobRessource;
      stWhereSalarie   := WhereSalarie;
      Rupture := NiveauRupt;
      { Pour optimiser, On ne charge pas immédiatement les absences, on les chargera lors du chargement des items }
      GestionPresence := TGestionPresence.Create(True, True, True, True, False, DateDeb,DateFin);
      Realise := ShowRealise;
      previsionnel := Showprevisionnel;
      Planning.Caption := 'Planning des rythmes de travail des salariés';
      { Affichage du planning }
      ShowPlanning;
    end;
  finally
    if Assigned(Planning.GestionPresence) then FreeAndNil(Planning.GestionPresence);
    if Assigned(Planning.TobJourneesTypes) then FreeAndNil(Planning.TobJourneesTypes);
    if Planning<>nil then Planning.Free;
  end;
end;

{ TPlanningPresenceSal }

procedure TPlanningPresenceSal.AskRefreshDetailsProcedure(Planning: THPlanning;
  laTob: TOB);
var
  indexChamp : Integer;
  stRNom : String;
  //Début PT1
  DateClique : TDateTime;
  TobSalClique, TobJourneeType : TOB;
  CycleOuModele, TypeDeCycle, JourneeType : String;
  isExceptionSalarie, isExceptionCycle : Boolean;
  //Fin PT1
  function ReplaceNomChamp(NomChamp, Valeur : String) : String;
  var
   st1, st2 : String;
  begin
    result := NomChamp+' : '+Valeur;
    if NomChamp = 'ETAT' then
    begin
      st1 := Valeur;
      While st1 <> '' do
        st2 := readTokenPipe(st1, '_');
      result := 'Code : '+st2;
    end;
    if NomChamp = 'ETAT_LIBELLE' then result := 'Libellé : '+Valeur;
    if NomChamp = 'ETAT_ABREGE' then result := '';
    if NomChamp = 'PJO_HORDEBPLAGE1' then
      if (Valeur <> '01/01/1900') and (Valeur <> '') then
        result := 'Heure de début de la plage 1 : '+Valeur else result := '';
    if NomChamp = 'PJO_HORFINPLAGE1' then
      if (Valeur <> '01/01/1900') and (Valeur <> '') then
        result := 'Heure de fin de la plage 1 : '+Valeur else result := '';
    if (NomChamp = 'PJO_JOURJ1PLAGE1') and ((Valeur = '-') or (Valeur = '')) then result := '';
    if (NomChamp = 'PJO_JOURJ1PLAGE1') and (Valeur = 'X') then result := 'La plage 1 se termine le lendemain';
    if NomChamp = 'PJO_HORDEBPLAGE2' then
      if (Valeur <> '01/01/1900') and (Valeur <> '') then
        result := 'Heure de début de la plage 2 : '+Valeur else result := '';
    if NomChamp = 'PJO_HORFINPLAGE2' then
      if (Valeur <> '01/01/1900') and (Valeur <> '') then
        result := 'Heure de fin de la plage 2 : '+Valeur else result := '';
    if (NomChamp = 'PJO_JOURJ1PLAGE2') and ((Valeur = '-') or (Valeur = '')) then result := '';
    if (NomChamp = 'PJO_JOURJ1PLAGE2') and (Valeur = 'X') then result := 'La plage 2 se termine le lendemain';
    if NomChamp = 'SALARIE' then result := 'Salarié : '+LeftStr(Valeur, Length(Valeur)-1);
    if NomChamp = 'DATEDEBUT' then result := 'Date de début : '+Valeur;
    if NomChamp = 'DATEFIN' then result := 'Date de fin : '+Valeur;
    if (NomChamp = 'TYPE_ITEM') and (Valeur = 'ABS') then result := 'Type : Absence';
    if (NomChamp = 'TYPE_ITEM') and (Valeur = 'PRE') then result := 'Type : Présence';
    if (NomChamp = 'TYPE_ITEM') and (Valeur = 'EXCEPTION') then result := 'Type : Exception';
    if (NomChamp = 'TYPE_ITEM') and (Valeur = 'CYC') then result := 'Type : Rythme de travail';
    if (NomChamp = 'TYPE_ITEM') and (Valeur = 'MOD') then result := 'Type : Modèle de cyle';
    if (NomChamp = 'CLEF') then result := '';
    if (NomChamp = 'JOUR') then result := '';
  end;
begin
  if Assigned(laTob) then
  begin
    DetailBox.Items.Clear;
    for indexChamp := 0 to latob.NombreChampReel -1 do
    begin
      stRNom := ReplaceNomChamp(latob.GetNomChamp(indexChamp), latob.GetString(indexChamp));
      if stRNom <> '' then
        DetailBox.Items.Add(stRNom);
    end;
    for indexChamp := 0 to latob.NombreChampSup -1 do
    begin
      stRNom := ReplaceNomChamp(latob.GetNomChamp(1000+indexChamp), latob.GetString(1000+indexChamp));
      if stRNom <> '' then
        DetailBox.Items.Add(stRNom);
    end;
  end else
    DetailBox.Items.Clear;
  //DEBUT PT1
  { Recherche de la date sélectionnée }
  DateClique := Planning.GetDateOfCol(Planning.Col);
  { Recherche du salarié sélectionné }
  TobSalClique := Planning.GetRessourceOfRow(Planning.Row);
  { Recherche des informations concernant cette journée du salarié }
  if (DateClique > iDate1900) and (Assigned(TobSalClique)) then
  begin
    JourneeType := GestionPresence.GetJourneeTypeSalarie(DateClique, TobSalClique.getString('PSA_SALARIE'), CycleOuModele, TypeDeCycle, isExceptionSalarie, isExceptionCycle, True);
    TobJourneeType := TobJourneesTypes.FindFirst(['PJO_JOURNEETYPE'], [JourneeType], False);
    if Assigned(TobJourneeType) then
    begin
      if isExceptionCycle   then
        DetailBox.Items.Add('Une exception est saisie au niveau du cycle');
      if isExceptionSalarie then
        DetailBox.Items.Add('Une exception est saisie au niveau du salarié');
      stRNom := 'Journée type du '+AGLDateToStr(DateClique)+': '+JourneeType+' '+TobJourneeType.GetString('PJO_LIBELLE');
      DetailBox.Items.Add(stRNom);
      stRNom := 'Plage 1 de '+TimeToStr(TobJourneeType.GetDateTime('PJO_HORDEBPLAGE1'))+' à '+TimeToStr(TobJourneeType.GetDateTime('PJO_HORFINPLAGE1'))+' (Durée :'+TobJourneeType.GetString('PJO_DUREEPLAGE1')+' heures)';
      if TobJourneeType.GetBoolean('PJO_JOURJ1PLAGE1') then stRNom := stRNom + ' le lendemain';
      DetailBox.Items.Add(stRNom);
      stRNom := 'Plage 2 de '+TimeToStr(TobJourneeType.GetDateTime('PJO_HORDEBPLAGE2'))+' à '+TimeToStr(TobJourneeType.GetDateTime('PJO_HORFINPLAGE2'))+' (Durée :'+TobJourneeType.GetString('PJO_DUREEPLAGE2')+' heures)';
      if TobJourneeType.GetBoolean('PJO_JOURJ1PLAGE2') then stRNom := stRNom + ' le lendemain';
      DetailBox.Items.Add(stRNom);
    end;
  end;
  //FIN PT1
end;

procedure TPlanningPresenceSal.LoadRessources(Deb, Fin : TDateTime);
Var
  StSQLLoadRessources, stChampsRupture : String;
  TempTobSal, TempTobRessources, tempTobRes, findTob : Tob;
  indexSal, indexPlanning : Integer;
begin
  { Chargement des Ressources }
  TobRessources.ClearDetail;
  TempTobSal := Tob.Create('tob des salariés',nil,-1);
  TempTobRessources := Tob.Create('tob des ressources avec gestion des ruptures',nil,-1);

  stChampsRupture := ListeChampRupt(',',Rupture);
  StSQLLoadRessources := 'SELECT '+stChampsRupture+' PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM, PSA_LIBELLE||" "||PSA_PRENOM AS RESSOURCE_LIBELLE FROM SALARIES left outer join deportsal on psa_salarie = pse_salarie '+stWhereSalarie;
  TempTobSal.LoadDetailFromSQL(StSQLLoadRessources, True);
  { Si on affiche plusieurs salariés, on désactive le planning Outlook }
  if TempTobSal.FillesCount(0)-1 > 1 then
  begin
    for indexPlanning := 0 to PCPlannings.PageCount -1 do
    begin
      if PCPlannings.Pages[indexPlanning].Caption = 'Outlook' then
        PCPlannings.Pages[indexPlanning].TabVisible := False;
    end;
  end;
  TempTobSal.Detail.Sort('RESSOURCE');
  { Gestion des ruptures }
  MiseEnFormeTobRupt('RES', TempTobSal,TempTobRessources,Rupture);
  { On ajoute la gestion du type réalisé ou prévisionnel }
  for indexSal := 0 to TempTobRessources.FillesCount(0)-1 do
  begin
    if TempTobRessources.Detail[indexSal].GetString('PSA_SALARIE') <> '' then
    begin
      { Cas des ressources correspondant à un salarié }
      if previsionnel then
      begin
        tempTobRes := Tob.Create('tob d''une ressource salarié (Prévisionnel)',TobRessources,-1);
        tempTobRes.Dupliquer(TempTobRessources.Detail[indexSal], True, True);
        findTob := TempTobSal.FindFirst(['PSA_SALARIE'], [TempTobRessources.Detail[indexSal].GetString('PSA_SALARIE')], False);
        tempTobRes.AddChampSupValeur('RESSOURCE', findTob.GetString('PSA_SALARIE')+'P');
        tempTobRes.AddChampSupValeur('RESSOURCE_LIBELLE', findTob.GetString('RESSOURCE_LIBELLE'));
        tempTobRes.AddChampSupValeur('TYPE', 'Rythme de travail');
        tempTobRes.AddChampSupValeur('PSA_ETABLISSEMENT', findTob.GetString('PSA_ETABLISSEMENT'));
        tempTobRes.AddChampSupValeur('PSA_SALARIE_SHOW', findTob.GetString('PSA_SALARIE'));
      end;
      if Realise then
      begin
        tempTobRes := Tob.Create('tob d''une ressource salarié (réalisé)',TobRessources,-1);
        tempTobRes.Dupliquer(TempTobRessources.Detail[indexSal], True, True);
        findTob := TempTobSal.FindFirst(['PSA_SALARIE'], [TempTobRessources.Detail[indexSal].GetString('PSA_SALARIE')], False);
        tempTobRes.AddChampSupValeur('RESSOURCE', findTob.GetString('PSA_SALARIE')+'R');
        if previsionnel then
        begin
          tempTobRes.AddChampSupValeur('RESSOURCE_LIBELLE', '');
          tempTobRes.AddChampSupValeur('PSA_SALARIE_SHOW', '');
        end else begin
          tempTobRes.AddChampSupValeur('RESSOURCE_LIBELLE', findTob.GetString('RESSOURCE_LIBELLE'));
          tempTobRes.AddChampSupValeur('PSA_SALARIE_SHOW', findTob.GetString('PSA_SALARIE'));
        end;
        tempTobRes.AddChampSupValeur('TYPE', 'Evènement');
        tempTobRes.AddChampSupValeur('PSA_ETABLISSEMENT', findTob.GetString('PSA_ETABLISSEMENT'));
      end;
    end else begin
      { Cas des ressources de gestion des ruptures (ne correspondant à aucun salarié) }
      tempTobRes := Tob.Create('tob d''une ressource de rupture',TobRessources,-1);
      tempTobRes.Dupliquer(TempTobRessources.Detail[indexSal], True, True);
    end;
  end;
  FreeAndNil(TempTobRessources);
  FreeAndNil(TempTobSal);
end;

procedure TPlanningPresenceSal.LoadEtats(Deb, Fin : TDateTime);
Var
  StSQLLoadEtats : String;
  indexTob : Integer;
  newEtat : Tob;
begin
  { Chargement des Etats }
  { Journées types }
  StSQLLoadEtats := 'SELECT "PJO_"||PJO_JOURNEETYPE AS ETAT, PJO_LIBELLE AS ETAT_LIBELLE, PJO_COLORJOURTYPE AS ETAT_COLOR, '
                  + ' PJO_HORDEBPLAGE1, PJO_HORFINPLAGE1, PJO_JOURJ1PLAGE1, "JOT" AS TYPE_ITEM, '
                  + ' PJO_HORDEBPLAGE2, PJO_HORFINPLAGE2, PJO_JOURJ1PLAGE2, PJO_ABREGE AS ETAT_ABREGE '
//                  + ' PJO_DUREEPLAGE1, PJO_TYPEJOUR1, PJO_DUREEPLAGE2, PJO_TYPEJOUR2, '
//                  + ' PJO_TYPEJOUR, PJO_EQUIVTPSPLEIN, PJO_JOURTRAVFERIE, '
//                  + ' PJO_POIDSJOUR, PJO_TEMPSLIBRE1, PJO_TEMPSLIBRE2, PJO_TEMPSLIBRE3, '
//                  + ' PJO_TEMPSLIBRE4, PJO_TEMPSLIBRE5, PJO_TEMPSLIBRE6, PJO_TEMPSLIBRE7, '
//                  + ' PJO_DUREEPAUSE, PJO_PAIEMENTPAUSE, PJO_PAUSEEFFECTIF, PJO_DUREENOTRAV, '
//                  + ' PJO_DUREENOTRAVEFF, PJO_ABREGE'
                  + ' FROM JOURNEETYPE';
//  if StWhTobEtats <> '' then StSQLLoadEtats := StSQLLoadEtats +' WHERE PJO_JOURNEETYPE in ('+StWhTobEtats+')';
  TobEtats.LoadDetailFromSQL(StSQLLoadEtats, True);
  { Absences et présences}
  StSQLLoadEtats := 'SELECT "PMA_"||PMA_MOTIFABSENCE AS ETAT, PMA_LIBELLE AS ETAT_LIBELLE, PMA_PGCOLORACTIF AS ETAT_COLOR, '
                  + ' 0, 0, "-", PMA_TYPEMOTIF AS TYPE_ITEM, '
                  + ' 0, 0, "-", PMA_ABREGE AS ETAT_ABREGE  '
//                  + 'PMA_TYPERTT, PMA_ABREGE, PMA_RUBRIQUE, PMA_RUBRIQUEJ, PMA_ALIMENT, PMA_ALIMENTJ, PMA_GERECOMM, '
//                  + 'PMA_JOURHEURE, PMA_GESTIONMAXI, PMA_TYPEPERMAXI, PMA_PERMAXI, PMA_JRSMAXI, '
//                  + 'PMA_CALENDSAL, PMA_CALENDCIVIL, PMA_OUVRES, PMA_OUVRABLE, PMA_SSJOURFERIE, '
//                  + 'PMA_CONTROLMOTIF, PMA_PROFILABS, PMA_PROFILABSJ, PMA_TYPEATTEST, '
//                  + 'PMA_PREDEFINI, PMA_NODOSSIER, PMA_EDITION, PMA_PRISTOTAL, PMA_MOTIFEAGL, '
//                  + 'PMA_OKSAISIERESP, PMA_OKSAISIESAL, PMA_GESTIONIJSS, PMA_TYPEABS, '
//                  + 'PMA_PGCOLORPAY, PMA_PGCOLORVAL, PMA_PGCOLORATT, '
//                  + 'PMA_PGCOLORREF, PMA_PGCOLORANN, PMA_EDITPLANPAIE, PMA_EDITPLANABS, '
//                  + 'PMA_ANNULABLE, PMA_ALIMNETH, PMA_ALIMNETJ, PMA_TYPEMOTIF, PMA_PGCOLORPRE, '
//                  + 'PMA_CTRLABSEXISTE, PMA_CTRLPREEXISTE, PMA_CTRLPLAGEH,  PMA_CARENCEIJSS '
                  + 'FROM MOTIFABSENCE';  // WHERE PMA_CALENDSAL = "X"
//  if StWhTobEtats <> '' then StSQLLoadEtats := StSQLLoadEtats +' WHERE PJO_JOURNEETYPE in ('+StWhTobEtats+')';
  TobEtats.LoadDetailFromSQL(StSQLLoadEtats, True);
  { Cycles }
  StSQLLoadEtats := 'SELECT "CYC_"||PYC_CYCLE||"_"||LTRIM(STR(1000000*DAY(PYC_DATEVALIDITE)+10000*MONTH(PYC_DATEVALIDITE)+YEAR(PYC_DATEVALIDITE))) AS ETAT, PYC_LIBELLE AS ETAT_LIBELLE, PYC_COLORCYCLE AS ETAT_COLOR, '
                  + ' 0, 0, "-", "CYC" AS TYPE_ITEM, '
                  + ' 0, 0, "-", PYC_LIBELLE AS ETAT_ABREGE  '
                  + 'FROM CYCLEENTETE';
  TobEtats.LoadDetailFromSQL(StSQLLoadEtats, True);
  { Modèles de cycles }
  StSQLLoadEtats := 'SELECT "MOD_"||PMY_MODELECYCLE||"_"||LTRIM(STR(1000000*DAY(PMY_DATEVALIDITE)+10000*MONTH(PMY_DATEVALIDITE)+YEAR(PMY_DATEVALIDITE))) AS ETAT, PMY_LIBELLE AS ETAT_LIBELLE, PMY_COLORMODCYCLE AS ETAT_COLOR, '
                  + ' 0, 0, "-", "MOD" AS TYPE_ITEM, '
                  + ' 0, 0, "-", PMY_LIBELLE AS ETAT_ABREGE  '
                  + 'FROM MODELECYCLEENT';
  TobEtats.LoadDetailFromSQL(StSQLLoadEtats, True);
  { Affectation d'une couleur par défaut }
  for indexTob := 0 to TobEtats.FillesCount(0) -1 do
  begin
    if Trim(TobEtats.Detail[indexTob].GetString('ETAT_COLOR')) = '' then
      TobEtats.Detail[indexTob].PutValue('ETAT_COLOR', 'clWhite');
  end;
  { Duplication des états pour les exceptions }
  for indexTob := TobEtats.FillesCount(0)-1 downto 0 do
  begin
    newEtat := Tob.Create('Etat Exception', TobEtats, -1);
    newEtat.Dupliquer(TobEtats.Detail[indexTob], True, True);
    newEtat.PutValue('TYPE_ITEM', 'EXCEPTION');
    newEtat.PutValue('ETAT', 'EXCEPTION_'+newEtat.GetString('ETAT'));
  end;
  TobEtats.detail[0].AddChampSupValeur('FONTCOLOR',ColorToString(FontColor),true);
  TobEtats.detail[0].AddChampSupValeur('FONTSTYLE',FontStyle,true);
  TobEtats.detail[0].AddChampSupValeur('FONTSIZE',FontSize,true);
  TobEtats.detail[0].AddChampSupValeur('FONTNAME',FontName,true);
end;

procedure TPlanningPresenceSal.LoadItems(Deb, Fin : TDateTime);
var
  TempTobRessources, newItem, tempTob, tempTob2, temptob3, tempTobAllCycles : Tob;
  indexRes, Ordre, indextob : Integer;
  SALARIE, AbsencePresenceSal, TypeMvt : String;
  DateCourante : TDateTime;
  ressource : String;
  TypeRessource, stclef : String;
  stType, stcycle, stdate, stLib : String;
  tobProfilsPresenceSalarie, tobTypesEtProfilsPresenceSalarie, tobRythmesSalarie, tobCyclesSalarie, tobModelesCyclesSalarie : Tob;
  indexCycle : Integer;
  dateDebutException, dateFinException, datevalid : TDateTime;
  intdatevalid : Integer;
begin
  { Chargement des Items }
  InitMoveProgressForm(nil, TraduireMemoire('Chargement des évènements'), TraduireMemoire('Veuillez patienter...'), 3*TobRessources.FillesCount(0), False, True);
  MoveCurProgressForm('Chargement des absences...');
  GestionPresence.ReloadAbsencePresenceSalarie(Deb, Fin, True);
  for indexRes := 0 to TobRessources.FillesCount(0) -1 do
  begin
    TempTobRessources := TobRessources.Detail[indexRes];
    SALARIE := TempTobRessources.GetString('RESSOURCE');
    TypeRessource := RightStr(SALARIE,1);
    SALARIE := LeftStr(SALARIE,Length(SALARIE)-1);

    MoveCurProgressForm('Chargement des rythmes de travail...');
    if TypeRessource = 'P' then // Prévisionnel
    begin
      tobProfilsPresenceSalarie := GestionPresence.GetProfilsPresenceSalarie(SALARIE);
      tobTypesEtProfilsPresenceSalarie :=  GestionPresence.GetTypesEtProfilsPresenceSalarie(SALARIE, tobProfilsPresenceSalarie);
      tobRythmesSalarie := GestionPresence.GetRythmesSalarie(tobTypesEtProfilsPresenceSalarie);
      tobCyclesSalarie := GestionPresence.GetCyclesSalarie(tobRythmesSalarie);
      tobModelesCyclesSalarie := GestionPresence.GetDecompositionModelesCyclesSalarie(tobCyclesSalarie, Deb, Fin);
      for indexCycle := tobModelesCyclesSalarie.FillesCount(0)-1 downto 0 do
      begin
        tempTob := tobModelesCyclesSalarie.Detail[indexCycle];
        if   ((tempTob.GetDateTime('DATEDEBUT') >= deb) and (tempTob.GetDateTime('DATEDEBUT') <= fin))
          or ((tempTob.GetDateTime('DATEDEBUT') <= deb) and (tempTob.GetDateTime('DATEFIN') >= fin))
          or ((tempTob.GetDateTime('DATEFIN') >= deb) and (tempTob.GetDateTime('DATEFIN') <= fin)) then
        begin
          datevalid := tempTob.GetDateTime('DATEVALIDITE');
          intdatevalid := 1000000*DayOf(datevalid)+10000*MonthOf(datevalid)+YearOf(datevalid);
          newItem := Tob.Create('Item Modele cycle de travail',TobItems, -1);
          newItem.AddChampSupValeur('SALARIE',SALARIE+TypeRessource);
          newItem.AddChampSupValeur('DATEDEBUT',tempTob.GetDateTime('DATEDEBUT'));
          newItem.AddChampSupValeur('JOUR', DateOf(tempTob.GetDateTime('DATEDEBUT')));
          newItem.AddChampSupValeur('DATEFIN',tempTob.GetDateTime('DATEFIN'));
          newItem.AddChampSupValeur('ETAT',tempTob.GetString('TYPE')+'_'+tempTob.GetString('CYCLE')+'_'+ TrimLeft(intToStr(intdatevalid)));
          newItem.AddChampSupValeur('ETAT_LIBELLE',tempTob.GetString('LIBELLE'));
          newItem.AddChampSupValeur('ETAT_ABREGE', tempTob.GetString('LIBELLE'));
          newItem.AddChampSupValeur('PJO_HORDEBPLAGE1','');
          newItem.AddChampSupValeur('PJO_HORFINPLAGE1','');
          newItem.AddChampSupValeur('PJO_HORDEBPLAGE2','');
          newItem.AddChampSupValeur('PJO_HORFINPLAGE2','');
          newItem.AddChampSupValeur('PJO_JOURJ1PLAGE1','-');
          newItem.AddChampSupValeur('PJO_JOURJ1PLAGE2','-');
          newItem.AddChampSupValeur('TYPE_ITEM',tempTob.GetString('TYPE'));
          newItem.AddChampSupValeur('CLEF',tempTob.GetString('TYPE')+';'+tempTob.GetString('CYCLE')+';'+tempTob.GetString('DATEDEBUT'));
        end;
      end;
      FreeAndNil(tobModelesCyclesSalarie);
      FreeAndNil(tobCyclesSalarie);
      FreeAndNil(tobRythmesSalarie);
      FreeAndNil(tobTypesEtProfilsPresenceSalarie);
      FreeAndNil(tobProfilsPresenceSalarie);
    end;

    MoveCurProgressForm('Chargement des Exceptions...');
    if TypeRessource = 'P' then // Prévisionnel
    begin
      { On parcours les cycles du salarié }
      tempTob := TobItems.FindFirst(['SALARIE'], [SALARIE+TypeRessource], False);
      while Assigned(tempTob) do
      begin
        stclef := tempTob.GetString('CLEF');
        stType := ReadtokenSt(stclef);
        stcycle := ReadtokenSt(stclef);
        stdate := ReadtokenSt(stclef);
        stType := tempTob.GetString('TYPE_ITEM');
        if (stType = 'CYC') or (stType = 'MOD') then
        begin
          tempTob2 := GestionPresence.TobExceptions.FindFirst(['PYA_TYPECYCLE', 'PYA_CYCLE'], [stType, stcycle], False);
          while Assigned(tempTob2) and (tempTob2.GetString('PYA_JOURNEETYPE') <> '') do
          begin
            dateDebutException := tempTob2.GetDateTime('PYA_DATEEXCEPTION');
            if (dateDebutException >= Deb) and (dateDebutException <= Fin) then
            begin
              temptob3 := TobEtats.findFirst(['ETAT'], ['PJO_'+tempTob2.GetString('PYA_JOURNEETYPE')], False);
              if Assigned(temptob3) then
                stLib := temptob3.GetString('ETAT_LIBELLE')
              else
                stLib := tempTob2.GetString('PYA_JOURNEETYPE');
              newItem := Tob.Create('Item Exception cycle',TobItems, -1);
              newItem.AddChampSupValeur('SALARIE',SALARIE+TypeRessource);
              newItem.AddChampSupValeur('DATEDEBUT',dateDebutException);
              newItem.AddChampSupValeur('JOUR', DateOf(dateDebutException));
              newItem.AddChampSupValeur('DATEFIN',dateDebutException);
              newItem.AddChampSupValeur('ETAT','EXCEPTION_PJO_'+tempTob2.GetString('PYA_JOURNEETYPE'));
              newItem.AddChampSupValeur('ETAT_LIBELLE',stLib);
              newItem.AddChampSupValeur('ETAT_ABREGE', stLib);
              newItem.AddChampSupValeur('PJO_HORDEBPLAGE1','');
              newItem.AddChampSupValeur('PJO_HORFINPLAGE1','');
              newItem.AddChampSupValeur('PJO_HORDEBPLAGE2','');
              newItem.AddChampSupValeur('PJO_HORFINPLAGE2','');
              newItem.AddChampSupValeur('PJO_JOURJ1PLAGE1','-');
              newItem.AddChampSupValeur('PJO_JOURJ1PLAGE2','-');
              newItem.AddChampSupValeur('TYPE_ITEM','EXCEPTION');
              newItem.AddChampSupValeur('CLEF',tempTob2.GetString('PYA_TYPECYCLE')+';'+tempTob2.GetString('PYA_CYCLE')+';'+tempTob2.GetString('PYA_DATEEXCEPTION'));
            end;
            tempTob2 := GestionPresence.TobExceptions.FindNext(['PYA_TYPECYCLE', 'PYA_CYCLE'], [stType, stcycle], False);
          end;
        end;
        tempTob := TobItems.FindNext(['SALARIE'], [SALARIE+TypeRessource], False);
      end;
      { On parcours les exceptions du salarié }
      tempTob := GestionPresence.TobExceptionSalaries.FindFirst(['PYE_SALARIE'], [SALARIE], False);
      while Assigned(tempTob) do
      begin
        dateDebutException := tempTob.GetDateTime('PYE_DATEDEBUT');
        dateFinException := tempTob.GetDateTime('PYE_DATEFIN');
        stType := tempTob.GetString('PYE_TYPEAFFECT');
        if stType = 'JOU' then
          stType := 'PJO'
        else if stType = 'CYC' then
        tempTobAllCycles := GetAllValidesFromTobValidite(GestionPresence.TobCycle, ['PYC_CYCLE'],
           [tempTob.GetString('PYE_CYCLEAFFECT')], 'PYC_DATEVALIDITE', 'DATEDEBUT', dateDebutException,
           'DATEFIN', dateFinException)
        else if stType = 'MOD' then
        tempTobAllCycles := GetAllValidesFromTobValidite(GestionPresence.TobModeleCycle, ['PMY_MODELECYCLE'],
           [tempTob.GetString('PYE_CYCLEAFFECT')], 'PMY_DATEVALIDITE', 'DATEDEBUT', dateDebutException,
           'DATEFIN', dateFinException);
        //Mettre la date de validite du cycle dans le champs etat
        if Assigned(tempTobAllCycles) then
        begin
          for indextob := 0 to tempTobAllCycles.FillesCount(0)-1 do
          begin
            datevalid := idate1900;
            if stType = 'CYC' then
            begin
              datevalid := tempTobAllCycles.Detail[indextob].GetDateTime('PYC_DATEVALIDITE');
              stLib := tempTobAllCycles.Detail[indextob].GetString('PYC_LIBELLE');
            end else if stType = 'MOD' then
            begin
              datevalid := tempTobAllCycles.Detail[indextob].GetDateTime('PMY_DATEVALIDITE');
              stLib := tempTobAllCycles.Detail[indextob].GetString('PMY_LIBELLE');
            end;
            intdatevalid := 1000000*DayOf(datevalid)+10000*MonthOf(datevalid)+YearOf(datevalid);
            dateDebutException := tempTobAllCycles.Detail[indextob].GetDateTime('DATEDEBUT');
            dateFinException := tempTobAllCycles.Detail[indextob].GetDateTime('DATEFIN');
            if   ((dateDebutException >= Deb) and (dateDebutException <= Fin))
              or ((dateFinException >= Deb) and (dateFinException <= Fin))
              or ((dateDebutException <= Deb) and (dateFinException >= Fin)) then
            begin
              newItem := Tob.Create('Item Exception salarie',TobItems, -1);
              newItem.AddChampSupValeur('SALARIE',SALARIE+TypeRessource);
              newItem.AddChampSupValeur('DATEDEBUT',dateDebutException);
              newItem.AddChampSupValeur('JOUR', DateOf(dateDebutException));
              newItem.AddChampSupValeur('DATEFIN',dateFinException);
              newItem.AddChampSupValeur('ETAT','EXCEPTION_'+stType+'_'+tempTob.GetString('PYE_CYCLEAFFECT')+'_'+ TrimLeft(intToStr(intdatevalid)));
              newItem.AddChampSupValeur('ETAT_LIBELLE',stLib);
              newItem.AddChampSupValeur('ETAT_ABREGE', stLib);
              newItem.AddChampSupValeur('PJO_HORDEBPLAGE1','');
              newItem.AddChampSupValeur('PJO_HORFINPLAGE1','');
              newItem.AddChampSupValeur('PJO_HORDEBPLAGE2','');
              newItem.AddChampSupValeur('PJO_HORFINPLAGE2','');
              newItem.AddChampSupValeur('PJO_JOURJ1PLAGE1','-');
              newItem.AddChampSupValeur('PJO_JOURJ1PLAGE2','-');
              newItem.AddChampSupValeur('TYPE_ITEM', 'EXCEPTION');
              newItem.AddChampSupValeur('CLEF',tempTob.GetString('PYE_SALARIE')+';'+tempTob.GetString('PYE_DATEDEBUT'));
            end;
          end;
        end else begin
          if   ((dateDebutException >= Deb) and (dateDebutException <= Fin))
            or ((dateFinException >= Deb) and (dateFinException <= Fin))
            or ((dateDebutException <= Deb) and (dateFinException >= Fin)) then
          begin
            temptob3 := TobEtats.findFirst(['ETAT'], ['PJO_'+tempTob.GetString('PYE_CYCLEAFFECT')], False);
            if Assigned(temptob3) then
              stLib := temptob3.GetString('ETAT_LIBELLE')
            else
              stLib := tempTob.GetString('PYE_CYCLEAFFECT');
            newItem := Tob.Create('Item Exception salarie',TobItems, -1);
            newItem.AddChampSupValeur('SALARIE',SALARIE+TypeRessource);
            newItem.AddChampSupValeur('DATEDEBUT',dateDebutException);
            newItem.AddChampSupValeur('JOUR', DateOf(dateDebutException));
            newItem.AddChampSupValeur('DATEFIN',dateFinException);
            newItem.AddChampSupValeur('ETAT','EXCEPTION_'+stType+'_'+tempTob.GetString('PYE_CYCLEAFFECT'));
            newItem.AddChampSupValeur('ETAT_LIBELLE',stLib);
            newItem.AddChampSupValeur('ETAT_ABREGE', stLib);
            newItem.AddChampSupValeur('PJO_HORDEBPLAGE1','');
            newItem.AddChampSupValeur('PJO_HORFINPLAGE1','');
            newItem.AddChampSupValeur('PJO_HORDEBPLAGE2','');
            newItem.AddChampSupValeur('PJO_HORFINPLAGE2','');
            newItem.AddChampSupValeur('PJO_JOURJ1PLAGE1','-');
            newItem.AddChampSupValeur('PJO_JOURJ1PLAGE2','-');
            newItem.AddChampSupValeur('TYPE_ITEM', 'EXCEPTION');
            newItem.AddChampSupValeur('CLEF',tempTob.GetString('PYE_SALARIE')+';'+tempTob.GetString('PYE_DATEDEBUT'));
          end;
        end;
        if Assigned(tempTobAllCycles) then FreeAndNil(tempTobAllCycles);
        tempTob := GestionPresence.TobExceptionSalaries.FindNext(['PYE_SALARIE'], [SALARIE], False);
      end;
    end;
    MoveCurProgressForm('Chargement des journées...');
//    if TypeRessource = 'P' then // Prévisionnel
//    begin
//      DateCourante := Deb;
//      While DateCourante <= Fin do
//      begin
//        { Pour chaque jour on recherche la journée type du rythme de travail avec ses exceptions }
//        JourneeType := GestionPresence.GetJourneeTypeSalarie(DateCourante, SALARIE, Cycle, TypeCycle, isExSal, isExCycle);
//        if JourneeType <> '' then
//          CreateNewPresenceItem('PJO_'+JourneeType, DateCourante, SALARIE+TypeRessource, 'SALARIE', TobEtats, TobItems, False);
//        DateCourante := DateCourante +1 ;
//      end;
//    end else
    if TypeRessource = 'R' then // Réalisé
    begin
      DateCourante := Deb;
      While DateCourante <= Fin do
      begin
        { Pour chaque jour on recherche les absences/présences }
        AbsencePresenceSal := GestionPresence.GetAbsencePresenceSalarie(DateCourante, SALARIE, TypeMvt, ressource, Ordre);
        if AbsencePresenceSal <> '' then
          CreateNewPresenceItem('PMA_'+AbsencePresenceSal, DateCourante, SALARIE+TypeRessource, 'SALARIE', TobEtats, TobItems, False, TypeMvt+';'+SALARIE+';'+IntToStr(Ordre));
//        AbsenceSal := GestionPresence.GetAbsenceSalarie(DateCourante, SALARIE, ressource, Ordre);
//        if AbsenceSal <> '' then
//          CreateNewPresenceItemAbsence(AbsenceSal, DateCourante, SALARIE+TypeRessource, 'SALARIE', TobEtats, TobItems, False);
//        { Pour chaque jour on recherche les présences }
//        PresenceSal := GestionPresence.GetAbsenceSalarie(DateCourante, SALARIE, ressource, Ordre);
//        if PresenceSal <> '' then
//          CreateNewPresenceItemPresence(PresenceSal, DateCourante, SALARIE+TypeRessource, 'SALARIE', TobEtats, TobItems, False);
        DateCourante := DateCourante +1 ;
      end;
    end;
  end;
  FiniMoveProgressForm;
end;

procedure TPlanningPresenceSal.SpecifPlanningPropertys(Planning: THPlanning;
  Sender: TObject);
begin
  with Planning do
  begin
    { Ressources }
    ResChampID := 'RESSOURCE';
    { Items }
    ChampLineID := 'SALARIE';
    ChampdateDebut := 'DATEDEBUT';
    ChampDateFin := 'DATEFIN';
    ChampEtat := 'ETAT';
    ChampLibelle := 'ETAT_LIBELLE';
    ChampHint := 'ETAT_LIBELLE';
    { Etats }
    EtatChampCode := 'ETAT';
    EtatChampLibelle := 'ETAT_LIBELLE';
    EtatChampBackGroundColor := 'ETAT_COLOR';
    { PopupMenu }
    AddOptionPopup(97,'Gérer une absence pour le salarié');
    EnableOptionPopup(97,True);
    AddOptionPopup(98,'Gérer une présence pour le salarié');
    EnableOptionPopup(98,True);
//    AddOptionPopup(99,'Gérer une exception pour le rythme de travail');
//    EnableOptionPopup(99,True);
    AddOptionPopup(100,'Gérer une exception pour le salarié');
    EnableOptionPopup(100,True);
    { Alignement des colonnes fixes }
//    TokenSizeColFixed   := '70;150';
//    TokenAlignColFixed  := 'C;L';
    { Liste des champs d'entete }
    TokenFieldColEntete := '';
    TokenFieldColFixed  := ListeChampRupt(';',Rupture) + 'PSA_SALARIE_SHOW;RESSOURCE_LIBELLE';
    MiseEnFormeColonneFixeRupt(Planning,Rupture);
    if Realise and previsionnel then
    begin
      TokenSizeColFixed   := TokenSizeColFixed+';85';
      TokenAlignColFixed  := TokenAlignColFixed+';L';
      TokenFieldColFixed  := TokenFieldColFixed+';TYPE';
    end;
  end;
end;

procedure TPlanningPresenceSal.GereOptionPopup(Item: TOB; ZCode: Integer;
  var ReDraw: Boolean; IdRessource : String; CurentDate : TDateTime; ZoneSelectionnee : TZoneSelected);
var
  JourneeType, CycleOuModele, TypeCycle : String;
  isExSal, isExCycle : Boolean;
  SALARIE, TypeRessource : String;
  tobSal, TempTob : Tob;
  debut, Fin : TDateTime;
begin
  if (IdRessource <> '') and (CurentDate <> iDate1900) then
  begin
    SALARIE := IdRessource;
    TypeRessource := RightStr(SALARIE,1);
    SALARIE := LeftStr(SALARIE,Length(SALARIE)-1);
    debut := AllPlanIntervalDebut;
    Fin   := AllPlanIntervalFin;
    tobSal := TobRessources.FindFirst(['RESSOURCE'], [IdRessource], False);
    JourneeType := GestionPresence.GetJourneeTypeSalarie(CurentDate, SALARIE, CycleOuModele, TypeCycle, isExSal, isExCycle);
    Plannings[PCPlannings.ActivePageIndex].Activate := False;
    if (ZCode = 97) then  { Gestion d'une absence }
    begin
      TempTob := GestionPresence.GetTobAbsencePresenceSalarie(CurentDate, SALARIE, 'ABS' );
      if not Assigned(TempTob) then TempTob := GestionPresence.GetTobAbsencePresenceSalarie(CurentDate, SALARIE, 'CPA' );
      if Assigned(TempTob) then
      begin
        debut := TempTob.GetDateTime('PCN_DATEDEBUTABS');
        Fin   := TempTob.GetDateTime('PCN_DATEFINABS');
        AglLanceFiche('PAY', 'EABSENCE', 'PCN_SALARIE=' + SALARIE, TempTob.GetString('PCN_TYPEMVT')+';'+SALARIE+';'+TempTob.GetString('PCN_ORDRE'),
                      SALARIE+';E;'+tobSal.GetString('PSA_ETABLISSEMENT')+';ACTION=MODIFICATION;;ADM;'+ DateToStr(debut)+';'+DateToStr(Fin));
        FreeAndNil(TempTob)
      end else begin
        debut := ZoneSelectionnee.DebutSelection;
        Fin   := ZoneSelectionnee.FinSelection;
        AglLanceFiche('PAY', 'EABSENCE', 'PCN_SALARIE=' + SALARIE, SALARIE,
                      SALARIE+ '!' + tobSal.GetString('RESSOURCE_LIBELLE') +';E;;ACTION=CREATION;;ADM;'+DateToStr(debut)+';'+DateToStr(Fin));
      end;
    end else if (ZCode = 98) then  { Gestion d'une présence }
    begin
      TempTob := GestionPresence.GetTobAbsencePresenceSalarie(CurentDate, SALARIE, 'PRE' );
      if Assigned(TempTob) then
      begin
        debut := TempTob.GetDateTime('PCN_DATEDEBUTABS');
        Fin   := TempTob.GetDateTime('PCN_DATEFINABS');
        AglLanceFiche('PAY', 'MVTPRESENCE', 'PCN_SALARIE=' + SALARIE, TempTob.GetString('PCN_TYPEMVT')+';'+SALARIE+';'+TempTob.GetString('PCN_ORDRE'),
                      SALARIE+';P;'+tobSal.GetString('PSA_ETABLISSEMENT')+';ACTION=MODIFICATION;;ADM;'+   DateToStr(debut)+';'+DateToStr(Fin));
        FreeAndNil(TempTob)
      end else begin
        debut := ZoneSelectionnee.DebutSelection;
        Fin   := ZoneSelectionnee.FinSelection;
        AglLanceFiche('PAY', 'MVTPRESENCE', 'PCN_SALARIE=' + SALARIE, SALARIE,
                      SALARIE+ '!' + tobSal.GetString('RESSOURCE_LIBELLE') +';P;;ACTION=CREATION;;ADM;'+DateToStr(debut)+';'+DateToStr(Fin));
      end;
//        AGLLanceFiche('PAY', 'MVTPRESENCE','','',SALARIE+';P;'+tobSal.GetString('PSA_ETABLISSEMENT')+';ACTION=MODIFICATION;;' + 'Planning.TypUtilisat'+';'+DateToStr(ZoneSelectionnee.DebutSelection)+';'+DateToStr(ZoneSelectionnee.FinSelection));
//        AGLLanceFiche('PAY', 'MVTPRESENCE','',DateToStr(CurentDate),'ACTION=MODIFICATION');
//        PYP_SALARIE,PYP_DATEDEBUTPRES,PYP_DATEFINPRES,PYP_COMPTEURPRES,PYP_TYPECALPRES
      //else
     //   AGLLanceFiche('PAY', 'MVTPRESENCE','','','ACTION=CREATION;PYA_TYPECYCLE='+TypeCycle+';PYA_CYCLE='+CycleOuModele+';PYA_DATEEXCEPTION='+DateToStr(CurentDate));
//      GestionPresence.ReloadMvtPresence;
    end else if (ZCode = 99) then  { Gestion d'une exception au rythme de travail }
    begin
      debut := CurentDate;
      Fin   := CurentDate;
      if isExCycle then
        AGLLanceFiche('PAY', 'EXCEPTCYCLE','',TypeCycle+';'+CycleOuModele+';'+DateToStr(debut),'ACTION=MODIFICATION')
      else
        AGLLanceFiche('PAY', 'EXCEPTCYCLE','','','ACTION=CREATION;PYA_TYPECYCLE='+TypeCycle+';PYA_CYCLE='+CycleOuModele+';PYA_DATEEXCEPTION='+DateToStr(debut));
      GestionPresence.ReloadException;
    end else if (ZCode = 100) then  { Gestion d'une exception salarié }
    begin
      TempTob := GestionPresence.GetTobExceptionSalarie(CurentDate, SALARIE);
      if Assigned(TempTob) then
      begin
        debut := TempTob.GetDateTime('PYE_DATEDEBUT');
        Fin   := TempTob.GetDateTime('PYE_DATEFIN');
        AGLLanceFiche('PAY', 'EXCEPTPRESSAL','',SALARIE+';'+DateToStr(debut),'ACTION=MODIFICATION');
        FreeAndNil(TempTob)
      end else begin
        debut := ZoneSelectionnee.DebutSelection;
        Fin   := ZoneSelectionnee.FinSelection;
        AGLLanceFiche('PAY', 'EXCEPTPRESSAL','','','ACTION=CREATION;'+SALARIE+';;;'+DateToStr(debut)+';'+DateToStr(Fin));
      end;
      GestionPresence.ReloadExceptionSal;
    end;
    DelItems (debut, Fin, 'DATEDEBUT', 'DATEFIN');
    LoadItems(debut, Fin);
    Plannings[PCPlannings.ActivePageIndex].Activate := True;
  end;
end;

procedure TPlanningPresenceSal.OnClickHighlight(Sender: TObject);
begin
  if Sender = MiHighAnnule then
    Highlight([], [], false)
  else if Sender = MiHighAbs then
    Highlight(['TYPE_ITEM'], ['ABS'], True)
  else if Sender = MiHighPre then
    Highlight(['TYPE_ITEM'], ['PRE'], True)
  else if Sender = MiHighExcept then
    Highlight(['TYPE_ITEM'], ['EXCEPTION'], True)    //'CLEF','EXCEPTION'
  else if Sender = MiHighCycle then
    Highlight(['TYPE_ITEM'], ['CYC'], True)
  else if Sender = MiHighModele then
    Highlight(['TYPE_ITEM'], ['MOD'], True);
end;

procedure TPlanningPresenceSal.BeforeShowForm(Sender: TObject);
begin
  PopupHighlight.Items.Insert(0, MiHighAnnule);
  PopupHighlight.Items.Insert(1, MiHighAbs);
  PopupHighlight.Items.Insert(2, MiHighPre);
  PopupHighlight.Items.Insert(3, MiHighExcept);
  PopupHighlight.Items.Insert(4, MiHighCycle);
  PopupHighlight.Items.Insert(5, MiHighModele);
  BtnHighlight.Visible := True;
  PRECAP.Height := 150;
end;

end.

