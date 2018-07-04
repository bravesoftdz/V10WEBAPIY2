{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/07/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PRESCALCCOMPTEURS ()
Mots clefs ... : TOF;PRESCALCCOMPTEURS
*****************************************************************}
{
PT1 07/02/2008 GGU V_81 Utilisation de 2 paramsoc pour gérer les dates de clôture
}
Unit UTofPG_PRESCALCCOMPTEURS;

Interface

uses Classes, Sysutils, Controls,
  {$IFNDEF EAGLSERVER}
 AglInit,
  {$ENDIF}
StdCtrls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF EAGLSERVER}
    mul,
    FE_Main,
  {$ENDIF}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} UTOB,
{$ELSE}
  emul, MaineAgl, Utob,
{$ENDIF}
{$IFNDEF EAGLSERVER}
  Utof,
{$ENDIF}
  HCtrls,HTB97,HEnt1,PGPresence,ComCtrls;

Type
  TCompteursCalculator = class
  private
    DateSaisieDebutJour, DateFinJour, DateSaisieDebutHebdo, DateFinHebdo, DateSaisieDebutMens,
    DateFinMens, DateSaisieDebutAnn, DateFinAnn, DateSaisieDebutFinCycle, DateFinFinCycle : TDateTime;
    DateDebutJour, DateDebutHebdo, DateDebutMens, DateDebutAnn, DateDebutFinCycle : TDateTime;
//    CalculJour, CalculHebdo, CalculFinCycle, CalculMens, CalculAnn : Boolean;
    DateDebTousSalaries : TDateTime;
//    ExerPerEnCours, DebPer, FinPer : String; { Exercice }
    RecalculEnCours : Integer;
    TobCompteursARecalculer : Tob;
    SelectionUtilisateur : Boolean; // Des critères ont été positionnés par l'utilisateur
    Recalcul       : Boolean; // Un recalcul global a été lancé (avec une coche)
    CurrentSal : String;
//    RecalculCurrentSal : Boolean; // Le salarié en cours a besoin d'être recalculé : on a coché le recalcul ou on a un flag "ARE" pour tous ou juste pour lui.
    DateEntreeSalarie, DateSortieSalarie : TDateTime;
    procedure CalculCompteursProfil(TobCompteursProfils, TobCompteurs,
      Tob_Rub, TobSalarie: Tob; ProfilPres: String; DateDebut,
      DateFin: TDateTime);
    procedure CalculCompteurs(TobCompteursACalculer, Tob_Rub,
      TobSalarie: Tob);
    procedure TraiteARecalculer(TobValeursCompteurs : Tob); //; Salarie: String);
  public
//    Function LancementDuCalcul(WhereSalaries : String; CalculDesCompteursMensuels : Boolean; DateDebutCalcul, DateFinCalcul : TDateTime) : Boolean;
    diag : TObject;
    Function LancementDuCalcul(WhereSalaries : String;
      CalculDesCompteursJournaliers, CalculDesCompteursHebdomadaires, CalculDesCompteursFinDeCycle,
      CalculDesCompteursMensuels, CalculDesCompteursAnnuels : Boolean;
      RecalcEnCours : Integer;
      ForceRecalcul : Boolean;
      DateDeDebutPourTousLesSalaries :TDateTime;
      DateSaisieDebutJournaliers, DateSaisieDebutHebdomadaires, DateSaisieDebutFinDeCycle, DateSaisieDebutMensuels, DateSaisieDebutAnnuels : TDateTime;
      DateFinJournaliers, DateFinHebdomadaires, DateFinFinDeCycle, DateFinMensuels, DateFinAnnuels : TDateTime;
      DateDebutCalcul, DateFinCalcul : TDateTime) : Boolean; OverLoad;
    Function LancementDuCalcul(TobParametres : Tob) : Boolean; OverLoad;
    Constructor Create(ObjDiag : TObject = nil);
    Destructor Destroy(); Override;
  end;

{$IFNDEF EAGLSERVER}
Type
  TOF_PRESCALCCOMPTEURS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    DebPourTousSalaries : TDateTime;
    Recalculer       : Boolean; // Un recalcul a été lancé
    CriteresSaisis : Boolean; // Des critères ont été positionnés par l'utilisateur
    RecalcEnCours : Integer;
    CalculJour, CalculHebdo, CalculFinCycle, CalculMens, CalculAnn : Boolean;
    diag : TListBox;
    LBDetails : TListBox;
    Calculator : TCompteursCalculator;
    procedure LancementDuCalcul  (Sender : TObject);
    Procedure Calcul(AsJob : Boolean = False);
    Procedure ClickSortie        (Sender : TObject);
    Procedure ClickCalcMensuel   (Sender : TObject);
    Procedure ClickRecalcul      (Sender : TObject);
    Procedure ClickCalcAnnuel    (Sender : TObject);
    Procedure ClickChangePeriode (Sender : TObject);
    Procedure GrilleDblClick     (Sender : TObject);
    Procedure ClickAddJob        (Sender : TObject);
  end ;
{$ENDIF}

  Function ConvertPrefixe(StWhere, DePref, APref: string): string;

Implementation

Uses
  EntPaie, P5Util, P5Def, uPaieEtabCompl,
  ed_Tools, uPaieVariables, DateUtils, Math, PgOutils2, HQry, HMsgBox, HDB, ParamSoc
{$IFNDEF EAGLSERVER}
  ,UFicheJob
{$ENDIF}
  ,pgcalendrier
  ,PGTobOutils
  ;

{$IFNDEF EAGLSERVER}
procedure TOF_PRESCALCCOMPTEURS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PRESCALCCOMPTEURS.OnDelete ;
begin
  Inherited ;
end ;
procedure TOF_PRESCALCCOMPTEURS.OnUpdate ;
begin
  Inherited;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Chargement de la liste des salariés
Mots clefs ... :
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.OnLoad ;
var
  DateArret : TDateTime;
  St        : String;
begin
     // Si aucun traitement n'a été effectué, on laisse la date de début modifiable
     //If GetParamSocSecur(PresenceRenvoieParamSoc(JOUR), '') = '' Then
           SetControlEnabled ('EDDATEDEBUT', True);
     //Else
     //     SetControlEnabled ('EDDATEDEBUT', False);

     // Si on a sélectionné des salariés dans la liste, nul besoin de prendre en
     // compte le reste des critères.
     If GetControlText('SELECTEDSAL') = '' Then
     Begin
          if  TCheckBox(GetControl('CKSORTIE')) <> Nil Then
          Begin
               if (GetControlText('CKSORTIE')='X') {and (IsValidDate(GetControlText('DATEARRET')))} Then
               Begin
                    DateArret := StrtoDate(GetControlText('DATEARRET'));
                    St := ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'"'+
                          ' OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'"'+
                          ' OR PSA_DATESORTIE IS NULL)'+
                          ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'" ';
               End;
               SetControlText('XX_WHERE', St);
          End;
     End
     Else
          SetControlText('XX_WHERE', GetControlText('SELECTEDSAL'));
end ;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 27/07/2007
Modifié le ... :   /  /
Description .. : Chargement de l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.OnArgument (S : String ) ;
var Num : Integer;
    Debut, Fin : TDateTime;
     {$IFNDEF EAGLCLIENT}
          Liste : THDBGrid;
     {$ELSE}
          Liste : THGrid;
     {$ENDIF}
begin
  Inherited ;
  // En mode SAV, on active l'onglet "Détails du calcul"
  if V_PGI.SAV then
  begin
      LBDetails := TListBox.Create(Ecran);
      LBDetails.Parent := (GetControl('PDETAIL') as TTabSheet);
      LBDetails.Align := alClient;
      diag := LBDetails;
  end else
      diag := nil;

  Calculator := TCompteursCalculator.Create(Diag);
//  tobExercices := TOB.Create('liste des exercices de paie', nil, -1);
//  tobExercices.LoadDetailFromSQL('SELECT PEX_EXERCICE, PEX_LIBELLE, PEX_DATEDEBUT, PEX_DATEFIN FROM EXERSOCIAL');
//  tobEtabCompl := TOB.Create('liste des dates de clotures de CP par etablissement', nil, -1);
//  tobEtabCompl.LoadDetailFromSQL('Select ETB_ETABLISSEMENT, ETB_DATECLOTURECPN from ETABCOMPL');
     if GetControl('BOuvrir') <> nil then
          TToolBarButton97( GetControl('BOuvrir') ).OnClick := LancementDuCalcul;
     if GetControl('BAddJob') <> nil then
          TToolBarButton97( GetControl('BAddJob') ).OnClick := ClickAddJob;

     SetControlProperty('PDETAIL', 'TabVisible', Assigned(diag));
     SetControlProperty('PDETAIL', 'Visible', Assigned(diag));

     // Recherche des libellés des zones TRAVAIL et CODESTAT
     For Num := 1 To 4 Do
          VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
     VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));

     // Coches
     (GetControl('CKSORTIE')        As TCheckBox).OnClick := ClickSortie;
     (GetControl('CBCALCCOMPTMENS') As TCheckBox).OnClick := ClickCalcMensuel;
     (GetControl('CBRECALCUL')      As TCheckBox).OnClick := ClickRecalcul;
     (GetControl('CBCALCCOMPTANNU') As TCheckBox).OnClick := ClickCalcAnnuel;

     // Double-clic dans la grille
     {$IFNDEF EAGLCLIENT}
          Liste := THDBGrid(GetControl('FListe'));
     {$ELSE}
          Liste := THGrid(GetControl('FListe'));
     {$ENDIF}
     If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick;

     // Boutons Précedent et suivant
     (GetControl('BTNPREC') As TToolBarButton97).OnClick := ClickChangePeriode;
     (GetControl('BTNSUIV') As TToolBarButton97).OnClick := ClickChangePeriode;

     PresenceDonneSemaineCalculActuelle (Debut, Fin);
     SetControlText('EDDATEDEBUT', DateToStr(Debut));
     SetControlText('EDDATEFIN',   DateToStr(Fin));

     {$IFDEF EAGLCLIENT}
     (GetControl('PAGES') as THPageControl2).ActivePageIndex := 1; //1e onglet
     {$ENDIF}
end ;

procedure TOF_PRESCALCCOMPTEURS.OnClose ;
begin
  FreeAndNil(Calculator);
  Inherited ;
end ;

procedure TOF_PRESCALCCOMPTEURS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PRESCALCCOMPTEURS.OnCancel () ;
begin
  Inherited ;
end ;

//function TOF_PRESCALCCOMPTEURS.NeedRaz(periodeRaz: String; debut,
//  fin: TDateTime; TobExercices, TobEtabcompl : Tob; Etablissement : String) : Boolean;
//var
//  iPeriodeRaz, index : Integer;
//  TempDate : TDateTime;
//  TempTob : Tob;
//begin
//  result := False;
//  if (periodeRaz = '') or (not IsNumeric(periodeRaz)) then exit;
//  iPeriodeRaz := StrToInt(periodeRaz);
//  case iPeriodeRaz of
//    0 : begin //Fin d'exercice
//      { On parcours les exercices pour voir si l'une des dates de début fait partie de la période }
//      for index := 0 to tobExercices.FillesCount(0)-1 do
//      begin
//        TempDate := tobExercices.Detail[index].GetDateTime('PEX_DATEDEBUT');
//        if (debut <= TempDate) And (TempDate <= fin) then result := True;
//      end;
//    end;
//    1..12 : begin  //Fin de mois
//      TempDate := EncodeDate(YearOf(debut), iPeriodeRaz, 1);
//      TempDate := IncMonth(TempDate); //On test le premier jour du mois suivant
//      While TempDate <= fin do
//      begin
//        if (debut <= TempDate) And (TempDate <= fin) then
//        begin
//          result := True;
//          exit;
//        end;
//        TempDate := IncYear(TempDate);
//      end;
//    end;
//    20 : begin  //Fin de période CP
//      TempTob := TobEtabcompl.FindFirst(['ETB_ETABLISSEMENT'], [Etablissement], False);
//      if Assigned(TempTob) then
//      begin
//        TempDate := TempTob.GetDateTime('ETB_DATECLOTURECPN');
//        if (debut <= TempDate) And (TempDate <= fin) then result := True;
//      end;
//    end;
//    99 : begin  //Jamais
//      exit;
//    end;
//  end;
//end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU / FLO
Créé le ...... : 01/08/2007
Modifié le ... :   /  /
Description .. : Lancement du calcul des compteurs sur clic du
Suite ........ : bouton "BOuvrir"
Mots clefs ... :
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.LancementDuCalcul(Sender : TObject);
begin
  Calcul();
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Clic sur la coche "Calcul annuel".
Suite ........ : Permet la saisie des dates de l'année
Mots clefs ... :
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.ClickCalcAnnuel(Sender: TObject);
var Actif : Boolean;
    Q  : TQuery;
begin
     Actif := (GetControlText('CBCALCCOMPTANNU') = 'X');

     SetControlEnabled('EDDATEDEBUTANNEE', Actif);
     SetControlEnabled('EDDATEFINANNEE',   Actif);

     If Actif Then
     Begin
          // Mise à jour des dates avec celles de l'année en cours par défaut
          Q := OpenSQL('SELECT PEX_DATEDEBUT, PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_EXERCICE DESC', True);
          If Not Q.EOF Then
          Begin
               SetControlText('EDDATEDEBUTANNEE', DateToStr(Q.FindField('PEX_DATEDEBUT').AsDateTime));
               SetControlText('EDDATEFINANNEE',   DateToStr(Q.FindField('PEX_DATEFIN').AsDateTime));
          End;
          Ferme(Q);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Clic sur la coche "Calcul Mensuel".
Suite ........ : Permet la saisie de la date de début et met à jour la date de
Suite ........ : fin.
Mots clefs ... :
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.ClickCalcMensuel(Sender: TObject);
var Debut, Fin : TDateTime;
    Actif : Boolean;
begin
     Actif := (GetControlText('CBCALCCOMPTMENS') = 'X');

     If Not Actif Then SetControlText ('CBRECALCUL',  '');
     SetControlEnabled('CBRECALCUL',  Actif);

     If Actif Then
     Begin
          // Mise à jour des dates de début et de fin
          PresenceDonneMoisCalculActuel(Debut, Fin);

          SetControlText('EDDATEDEBUT', DateToStr(Debut));
          SetControlText('EDDATEFIN',   DateToStr(Fin));

          // Cas particulier : Si pas d'historique mensuel, la coche recalcul ne doit donc pas être accessible
          // et la date de début est modifiable
          If (GetParamSocDatesCloture(MOIS) = '') Then //PT1 GetParamSocSecur(PresenceRenvoieParamSoc(MOIS), '') = '') Then
          Begin
               SetControlEnabled('CBRECALCUL',  False);
               SetControlEnabled('EDDATEDEBUT', True);
          End;
     End
     Else
     Begin
          // On remet les dates hebdomadaires
          PresenceDonneSemaineCalculActuelle(Debut, Fin);

          SetControlText('EDDATEDEBUT', DateToStr(Debut));
          SetControlText('EDDATEFIN',   DateToStr(Fin));

          If (GetParamSocDatesCloture(JOUR) = '') Then //PT1 GetParamSocSecur(PresenceRenvoieParamSoc(JOUR), '') = '') Then
               SetControlEnabled('EDDATEDEBUT', True)
          Else
               SetControlEnabled('EDDATEDEBUT', False);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Clic sur la coche "Recalculer la période".
Suite ........ : Réaffecte les dates du dernier mois pris en compte
Mots clefs ... :
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.ClickRecalcul(Sender: TObject);
var Debut, Fin : TDateTime;
    Actif : Boolean;
begin
     Actif := (GetControlText('CBRECALCUL')='X');
     SetControlEnabled('EDDATEDEBUT', true);
     SetControlEnabled('EDDATEFIN',   Not Actif);
     SetControlEnabled('BTNPREC',     Actif);
     SetControlEnabled('BTNSUIV',     Actif);

     If Actif Then
          RecalcEnCours := -1
     Else
          RecalcEnCours := 0;

     // On recherche la période désirée
     PresenceDonnePeriodeCalcul(MOIS,RecalcEnCours,Debut, Fin);

     SetControlText('EDDATEDEBUT', DateToStr(Debut));
     SetControlText('EDDATEFIN',   DateToStr(Fin));
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Clic sur les boutons de périodes précédentes ou suivantes
Mots clefs ... :
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.ClickChangePeriode(Sender: TObject);
var Debut, Fin : TDateTime;
begin

     // On recherche l'ancienne période calculée
     If TToolBarButton97(Sender).Name = 'BTNPREC' Then
     Begin
          RecalcEnCours := RecalcEnCours - 1;
          PresenceDonnePeriodeCalcul(MOIS, RecalcEnCours, Debut, Fin);
          If Debut = iDate1900 Then
          Begin
               Debut := StrToDate(GetControlText('EDDATEDEBUT'));
               Fin   := StrToDate(GetControlText('EDDATEFIN'));
               RecalcEnCours := RecalcEnCours + 1;
          End;
     End
     Else
     // On recherche la période calculée ensuite
     Begin
          If RecalcEnCours < -1 Then
          Begin
               RecalcEnCours := RecalcEnCours + 1;
               PresenceDonnePeriodeCalcul(MOIS, RecalcEnCours, Debut, Fin);
          End
          Else
          Begin
               Debut := StrToDate(GetControlText('EDDATEDEBUT'));
               Fin   := StrToDate(GetControlText('EDDATEFIN'));
          End;
     End;

     SetControlText('EDDATEDEBUT', DateToStr(Debut));
     SetControlText('EDDATEFIN',   DateToStr(Fin));
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Clic sur la coche "Exclure les salariés sortis".
Mots clefs ... :
*****************************************************************}
procedure TOF_PRESCALCCOMPTEURS.ClickSortie(Sender: TObject);
Var Actif : Boolean;
    Debut,Fin : TDateTime;
begin
      Actif := (GetControlText('CKSORTIE')='X');
      SetControlVisible ('TDATEARRET', Actif);
      SetControlVisible ('DATEARRET',  Actif);
      SetControlEnabled ('DATEARRET',  Actif);

      If Actif Then
      Begin
          // Date de sortie des salariés : Début de période mensuelle
          PresenceDonneMoisCalculActuel (Debut, Fin);
          SetControlText('DATEARRET', DateToStr(Debut));
      End;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : FLO
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Double-clic dans la grille
Mots clefs ... :
*****************************************************************}
Procedure TOF_PRESCALCCOMPTEURS.GrilleDblClick(Sender: TObject);
var
  st : string;
  Q_Mul: THQuery;
begin
     If PgiAsk(TraduireMemoire('Voulez-vous calculer uniquement les compteurs du salarié sélectionné?'), TraduireMemoire('Calcul des compteurs de présence')) = mrNo Then
          Exit;

     Q_Mul := THQuery(Ecran.FindComponent('Q'));
     if Q_Mul = nil then exit;

     { Composition du clause WHERE pour limiter le mul à ces salariés }
     {$IFDEF EAGLCLIENT}
          TFmul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
     {$ENDIF}
     Q_Mul := THQuery(Ecran.FindComponent('Q')) ;

     St := ' AND PSA_SALARIE="' + Q_Mul.FindField('PSA_SALARIE').AsString + '"';

     SetControlText('SELECTEDSAL', St);
     TFMul(Ecran).BCherche.Click;

     { Récupération de la Query pour traitement }
     {$IFDEF EAGLCLIENT}
          if TFMul(Ecran).Fetchlestous then TheMulQ := TOB(Ecran.FindComponent('Q'));
     {$ELSE}
          TheMulQ := THQuery(Ecran.FindComponent('Q'));
     {$ENDIF}

     { Lancement du traitement }
     LancementDuCalcul (Nil);

     SetControlText('SELECTEDSAL', '');
     TheMulQ := nil;
     TFMul(Ecran).BCherche.Click;
end;
{$ENDIF EAGLSERVER}

//  TOF_PRESCALCCOMPTEURS.
procedure TCompteursCalculator.CalculCompteursProfil(TobCompteursProfils, TobCompteurs, Tob_Rub, TobSalarie : Tob; ProfilPres : String; DateDebut, DateFin : TDateTime);
var
  ListeCompteursProfils, TobListCompteursPeriode, TobProfil : Tob;
  indexCompteurList : Integer;
  { L'ordre de calcul est important : par exemple, les compteurs journaliers servent au calcul des compteurs hebdo... }
  Procedure TriCompteurs(var ListCompteurs : Tob);
  var
    indexCompteur : Integer;
    TobAn, TobMe, TobHe, TobJo, TobFi : Tob;
  begin
    TobAn := tob.Create('Liste des compteurs de périodicité annuelle', nil, -1);
    TobMe := tob.Create('Liste des compteurs de périodicité mensuelle', nil, -1);
    TobHe := tob.Create('Liste des compteurs de périodicité hebdomadaire', nil, -1);
    TobJo := tob.Create('Liste des compteurs de périodicité journalière', nil, -1);
    TobFi := tob.Create('Liste des compteurs de périodicité fin de cycle', nil, -1);
    for indexCompteur := ListCompteurs.FillesCount(0) -1 downto 0 do
    begin
      if ListCompteurs.Detail[indexCompteur].GetString('PYR_PERIODICITEPRE') = periodiciteAnnuelle then
        ListCompteurs.Detail[indexCompteur].ChangeParent(TobAn,-1)
      else if ListCompteurs.Detail[indexCompteur].GetString('PYR_PERIODICITEPRE') = periodiciteMensuelle then
        ListCompteurs.Detail[indexCompteur].ChangeParent(TobMe,-1)
      else if ListCompteurs.Detail[indexCompteur].GetString('PYR_PERIODICITEPRE') = periodiciteHebdomadaire then
        ListCompteurs.Detail[indexCompteur].ChangeParent(TobHe,-1)
      else if ListCompteurs.Detail[indexCompteur].GetString('PYR_PERIODICITEPRE') = periodiciteJournaliere then
        ListCompteurs.Detail[indexCompteur].ChangeParent(TobJo,-1)
      else if ListCompteurs.Detail[indexCompteur].GetString('PYR_PERIODICITEPRE') = periodiciteFinDeCycle then
        ListCompteurs.Detail[indexCompteur].ChangeParent(TobFi,-1);
    end;
    { Insertion dans l'ordre de calcul : journalière, hebdomadaire, fin de cycle, mensuelle, annuelle }
    for indexCompteur := TobJo.FillesCount(0) -1 downto 0 do
      TobJo.Detail[indexCompteur].ChangeParent(ListCompteurs, -1);
    for indexCompteur := TobHe.FillesCount(0) -1 downto 0 do
      TobHe.Detail[indexCompteur].ChangeParent(ListCompteurs, -1);
    for indexCompteur := TobFi.FillesCount(0) -1 downto 0 do
      TobFi.Detail[indexCompteur].ChangeParent(ListCompteurs, -1);
    for indexCompteur := TobMe.FillesCount(0) -1 downto 0 do
      TobMe.Detail[indexCompteur].ChangeParent(ListCompteurs, -1);
    for indexCompteur := TobAn.FillesCount(0) -1 downto 0 do
      TobAn.Detail[indexCompteur].ChangeParent(ListCompteurs, -1);
    FreeAndNil(TobAn);
    FreeAndNil(TobMe);
    FreeAndNil(TobHe);
    FreeAndNil(TobJo);
    FreeAndNil(TobFi);
  end;
begin
{  LogMessage(diag,'TobCompteursProfils.count '+IntToStr(TobCompteursProfils.Detail.Count));
  LogMessage(diag,'TobCompteurs.count '+IntToStr(TobCompteurs.Detail.Count));
  LogMessage(diag,'Tob_Rub.count '+IntToStr(Tob_Rub.Detail.Count));
  LogMessage(diag,'TobSalarie.count '+IntToStr(TobSalarie.Detail.Count));
  LogMessage(diag,'ProfilPres : String; DateDebut, DateFin '+ProfilPres + ' '+DateToStr(DateDebut)+' '+DateToStr(DateFin));
}
  ListeCompteursProfils := GetAllValidesFromTobValiditeMultiPass(TobCompteursProfils, ['PPV_PROFILPRES'], ['PPV_COMPTEURPRES'], [ProfilPres], 'PPV_DATEVALIDITE', 'DATEDEBUT', DateDebut, 'DATEFIN', DateFin);
  { Boucle sur les compteurs de la période }
  for indexCompteurList := 0 to ListeCompteursProfils.FillesCount(0) -1 do
  begin
    TobProfil := ListeCompteursProfils.Detail[indexCompteurList];
    TobListCompteursPeriode := GetAllValidesFromTobValidite(TobCompteurs, ['PYR_COMPTEURPRES'], [TobProfil.GetString('PPV_COMPTEURPRES')],
       'PYR_DATEVALIDITE', 'DATEDEBUT', TobProfil.GetDateTime('DATEDEBUT'), 'DATEFIN', TobProfil.GetDateTime('DATEFIN'));
    TriCompteurs(TobListCompteursPeriode);
    CalculCompteurs(TobListCompteursPeriode, Tob_Rub, TobSalarie);
    FreeAndNil(TobListCompteursPeriode);
  end;
  FreeAndNil(ListeCompteursProfils);
end;


//   TOF_PRESCALCCOMPTEURS.
{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 03/08/2007
Modifié le ... :   /  /
Description .. : Fonction qui recherche les lignes à recalculer pour en
Suite ........ : déduire les nouvelles dates de début
Mots clefs ... :
*****************************************************************}
procedure TCompteursCalculator.TraiteARecalculer(TobValeursCompteurs : Tob); //; Salarie : String);
var
  TempTob, TobValComptSal, TobValComptSalMem : Tob;
  RecalcDate, CompteurDate : TDateTime;
  perio : String;
  indexCompteur, indexValeurCompteur : Integer;
  DateSupprJour, DateSupprHebdo, DateSupprFinCycle, DateSupprMens, DateSupprAnn : TDateTime;
  ListeParam, Param, stSQL : String;

  procedure CheckRecalcul(DateRecalcul : TDateTime);
  begin
    if DateDebutJour     > DateRecalcul then DateDebutJour     := DateRecalcul;
    if DateDebutHebdo    > DateRecalcul then DateDebutHebdo    := DateRecalcul;
    if DateDebutFinCycle > DateRecalcul then DateDebutFinCycle := DateRecalcul;
    if DateDebutMens     > DateRecalcul then DateDebutMens     := DateRecalcul;
    if DateDebutAnn      > DateRecalcul then DateDebutAnn      := DateRecalcul;

    if DateSupprJour     > DateRecalcul then DateSupprJour     := DateRecalcul;
    if DateSupprHebdo    > DateRecalcul then DateSupprHebdo    := DateRecalcul;
    if DateSupprFinCycle > DateRecalcul then DateSupprFinCycle := DateRecalcul;
    if DateSupprMens     > DateRecalcul then DateSupprMens     := DateRecalcul;
    if DateSupprAnn      > DateRecalcul then DateSupprAnn      := DateRecalcul;
  end;
begin
  if Recalcul then
  begin
    DateSupprJour     := DateDebutJour;
    DateSupprHebdo    := DateDebutHebdo;
    DateSupprFinCycle := DateDebutFinCycle;
    DateSupprMens     := DateDebutMens;
    DateSupprAnn      := DateDebutAnn;
  end else begin
    DateSupprJour     := DateDebTousSalaries;
    DateSupprHebdo    := DateDebTousSalaries;
    DateSupprFinCycle := DateDebTousSalaries;
    DateSupprMens     := DateDebTousSalaries;
    DateSupprAnn      := DateDebTousSalaries;
  end;

  CheckRecalcul(DateDebTousSalaries);
  { On recherche toutes les lignes qui sont indiquées "à recalculer" pour le salarie }
  TempTob := TobCompteursARecalculer.FindFirst(['PYP_SALARIE'], [CurrentSal], False);
  RecalcDate := idate2099;
  while Assigned(TempTob) do
  begin
    if (TempTob.GetDateTime('PYP_DATEDEBUTPRES') < RecalcDate) and (TempTob.GetDateTime('PYP_DATEDEBUTPRES')  < DateDebutJour) then
      RecalcDate := TempTob.GetDateTime('PYP_DATEDEBUTPRES');
    if SelectionUtilisateur then
    begin
      { Suppression du compteur traité }
      TempTob.DeleteDB();
      TempTob.ChangeParent(nil, -1);
      FreeAndNil(TempTob);
    end;
    TempTob := TobCompteursARecalculer.FindNext(['PYP_SALARIE'], [CurrentSal], False);
  end;
  { La date de début de ces lignes deviens la nouvelle date de début }
  CheckRecalcul(RecalcDate);
  { Si on a un flag de recalcul et qu'on doit recalculer les compteurs depuis une date déjà clôturée
    on calcul les mensuels, même si ce n'est pas coché }
  if (DateDebutMens <> iDate2099) then
  begin
    ListeParam := GetParamSocDatesCloture(MOIS); //PT1 GetParamSocSecur(PresenceRenvoieParamSoc(MOIS), '');
    Param := ReadTokenSt(ListeParam);
    While (Param <> '') Do
    Begin
      if StrToDate(Param) >= RecalcDate then
      begin
        DateDebutMens := DateDebutJour;
        DateSupprMens := DateDebutJour;
        DateFinMens   := DateFinJour;
        Break;
      end;
      Param := ReadTokenSt(ListeParam);
    End;
  end;
  if (RecalcDate <> idate2099) then
  begin
    LogMessage(Diag,'  Les compteurs du salarié doivent être recalculés depuis le '+DateToStr(DateDebutJour));
    { Suppression de l'existant pour le salarié }
    ObjCalcuVarPre.FreeMemSalarie(CurrentSal);
    { On recharge les infos du salarié sur la période étendue }
    { Chargement des Compteurs calculés }
    TobValComptSal := Tob.Create('valeurs des Compteurs calculés pour le salarié courant', nil, -1);
    stSQL := ObjCalcuVarPre.BuildLoadingQuery(DateDebutJour, DateFinJour);
    stSQL := stSQL + ' and (PYP_SALARIE = "'+CurrentSal+'")';
    TobValComptSal.LoadDetailDBFromSQL('PRESENCESALARIE', stSQL);
    { On éclate la tob des compteurs du salarié dans la tob global, en vérifiant qu'on insère pas de doublons}
    TobValComptSalMem := ObjCalcuVarPre.GetTobValeursCompteursSalarie(CurrentSal);
    EclateTob(TobValComptSal, TobValComptSalMem, 'PYP_COMPTEURPRES', False);  //, True
    FreeAndNil(TobValComptSal);
    { Chargement des absences-présences }
    ObjCalcuVarPre.ReloadAbsencePresenceSalarie(DateDebutJour, DateFinJour, True, CurrentSal);
    { Chargement des jours feriés sur la période étendue }
    if Assigned(ObjCalcuVarPre.Tob_Ferie) then FreeAndNil(ObjCalcuVarPre.Tob_Ferie);
    ObjCalcuVarPre.Tob_Ferie := ChargeTobFerie(DateDebutJour, DateFinJour);
  end;
  { On supprime les lignes postérieures aux dates de calcul }
  { Suppression en base }
  ExecuteSQL('Delete from presencesalarie where PYP_SALARIE = "'+CurrentSal+'" '
            +'and ('
            +'    (PYP_PERIODICITEPRE = "'+periodiciteJournaliere +'" and PYP_DATEFINPRES >= "'+USDATETIME(DateSupprJour    ) +'") '
            +' or (PYP_PERIODICITEPRE = "'+periodiciteHebdomadaire+'" and PYP_DATEFINPRES >= "'+USDATETIME(DateSupprHebdo   ) +'") '
            +' or (PYP_PERIODICITEPRE = "'+periodiciteFinDeCycle  +'" and PYP_DATEFINPRES >= "'+USDATETIME(DateSupprFinCycle) +'") '
            +' or (PYP_PERIODICITEPRE = "'+periodiciteMensuelle   +'" and PYP_DATEFINPRES >= "'+USDATETIME(DateSupprMens    ) +'") '
            +' or (PYP_PERIODICITEPRE = "'+periodiciteAnnuelle    +'" and PYP_DATEFINPRES >= "'+USDATETIME(DateSupprAnn     ) +'") '
            +'    )'
            +'and PYP_ETATPRES <> "SAI" and PYP_PGINDICATPRES <> "INP" AND PYP_PGINDICATPRES <> "AIN"');
  { Suppression dans la tob }
  { Parcours des compteurs du salarie }
  for indexCompteur := 0 to TobValeursCompteurs.detail.Count -1 do
  begin
    { Parcours des valeurs de compteur }
    for indexValeurCompteur := TobValeursCompteurs.detail[indexCompteur].detail.Count -1 downto 0 do
    begin
      TempTob := TobValeursCompteurs.detail[indexCompteur].detail[indexValeurCompteur];
      { Contrôle pour ne pas supprimer de lignes saisies ou intégrées en paie }
      if ( (TempTob.GetString('PYP_ETATPRES') <> 'SAI') and (TempTob.GetString('PYP_PGINDICATPRES') <> 'INP') and (TempTob.GetString('PYP_PGINDICATPRES') <> 'AIN') ) then
      begin
        CompteurDate := TempTob.GetDateTime('PYP_DATEFINPRES');
        perio := TempTob.GetString('PYP_PERIODICITEPRE');
        if      ((perio = periodiciteJournaliere)  and (CompteurDate >= DateSupprJour    ))
             or ((perio = periodiciteHebdomadaire) and (CompteurDate >= DateSupprHebdo   ))
             or ((perio = periodiciteFinDeCycle)   and (CompteurDate >= DateSupprFinCycle))
             or ((perio = periodiciteMensuelle)    and (CompteurDate >= DateSupprMens    ))
             or ((perio = periodiciteAnnuelle)     and (CompteurDate >= DateSupprAnn     )) then
        begin
          TempTob.ChangeParent(nil, -1);
          FreeAndNil(TempTob);
        end;
      end;
    end;
  end;
end;

 //  TOF_PRESCALCCOMPTEURS.
procedure TCompteursCalculator.CalculCompteurs(TobCompteursACalculer, Tob_Rub, TobSalarie : Tob);
var
  TobCompteurACalculer, TobValeursCompteursSalarie : Tob;
  indexCompteur : Integer;
  periodicite, periodeRaz, Compteur : String; //, Salarie
  debut, fin : TDateTime;
  NeedSave, IsRegul, HadRegul : Boolean;
  Value, ValueToSave : Extended;
  Indicateur : String;
  Procedure SaveValue(TobSalarie, TobCompteur : Tob; Value : Extended; Regul : Boolean = False);
  var
    NewTob, TobParent : Tob;
    CodeCompteurPres : String;
  begin
    { Modification et insertion dans la tob, setmodifie pour ceux qui ont été insérés ou modifiés }

    CodeCompteurPres := TobCompteur.GetString('PYR_COMPTEURPRES');
    TobParent := TobValeursCompteursSalarie.FindFirst(['PYP_COMPTEURPRES'], [CodeCompteurPres], False);
    if not Assigned(TobParent) then
    begin
      TobParent := Tob.Create(TobValeursCompteursSalarie.NomTable, TobValeursCompteursSalarie, -1);
      TobParent.AddChampSupValeur('PYP_COMPTEURPRES', CodeCompteurPres);
    end;

    NewTob := tob.Create('PRESENCESALARIE', TobParent, -1);
    NewTob.AddChampSupValeur('PYP_SALARIE',        CurrentSal); // TobSalarie.GetString('PSA_SALARIE'));
    NewTob.AddChampSupValeur('PYP_DATEDEBUTPRES',  TobCompteur.GetDateTime('DATEDEBUT'));
    NewTob.AddChampSupValeur('PYP_DATEFINPRES',    TobCompteur.GetDateTime('DATEFIN'));
    NewTob.AddChampSupValeur('PYP_COMPTEURPRES',   TobCompteur.GetString('PYR_COMPTEURPRES'));
    NewTob.AddChampSupValeur('PYP_PERIODICITEPRE', TobCompteur.GetString('PYR_PERIODICITEPRE'));
    NewTob.AddChampSupValeur('PYP_THEMEPRE',       TobCompteur.GetString('PYR_THEMEPRE'));
    NewTob.AddChampSupValeur('PYP_QUANTITEPRES',   Value);
    NewTob.AddChampSupValeur('PYP_ETATPRES',       'CAL');
    NewTob.AddChampSupValeur('PYP_PGINDICATPRES',  'CAL');
    if not Regul then
      NewTob.AddChampSupValeur('PYP_TYPECALPRES',  '001')
    else
      NewTob.AddChampSupValeur('PYP_TYPECALPRES',  '002');
    NewTob.AddChampSupValeur('PYP_ETABLISSEMENT',  TobSalarie.GetString('PSA_ETABLISSEMENT'));
    NewTob.AddChampSupValeur('PYP_TRAVAILN1',      TobSalarie.GetString('PSA_TRAVAILN1'));
    NewTob.AddChampSupValeur('PYP_TRAVAILN2',      TobSalarie.GetString('PSA_TRAVAILN2'));
    NewTob.AddChampSupValeur('PYP_TRAVAILN3',      TobSalarie.GetString('PSA_TRAVAILN3'));
    NewTob.AddChampSupValeur('PYP_TRAVAILN4',      TobSalarie.GetString('PSA_TRAVAILN4'));
    NewTob.AddChampSupValeur('PYP_CODESTAT',       TobSalarie.GetString('PSA_CODESTAT'));
    NewTob.AddChampSupValeur('PYP_CONVENTION',     TobSalarie.GetString('PSA_CONVENTION'));
    NewTob.AddChampSupValeur('PYP_DADSPROF',       TobSalarie.GetString('PSA_DADSPROF'));
    NewTob.AddChampSupValeur('PYP_DADSCAT',        TobSalarie.GetString('PSA_DADSCAT'));
    NewTob.AddChampSupValeur('PYP_LIBREPCMB1',     TobSalarie.GetString('PSA_LIBREPCMB1'));
    NewTob.AddChampSupValeur('PYP_LIBREPCMB2',     TobSalarie.GetString('PSA_LIBREPCMB2'));
    NewTob.AddChampSupValeur('PYP_LIBREPCMB3',     TobSalarie.GetString('PSA_LIBREPCMB3'));
    NewTob.AddChampSupValeur('PYP_LIBREPCMB4',     TobSalarie.GetString('PSA_LIBREPCMB4'));
    NewTob.AddChampSupValeur('PYP_DATEDEBUTBUL',   iDate1900); //DebPer);
    NewTob.AddChampSupValeur('PYP_DATEFINBUL',     iDate1900); //FinPer);
    NewTob.SetAllModifie(True);
  end;
begin
{  LogMessage(diag,'TobCompteursACalculer.count '+IntToStr(TobCompteursACalculer.Detail.Count));
  LogMessage(diag,'Tob_Rub.count '+IntToStr(Tob_Rub.Detail.Count));
  LogMessage(diag,'TobSalarie.count '+IntToStr(TobSalarie.Detail.Count));
}
  //  Salarie := TobSalarie.GetString('PSA_SALARIE');
  ObjCalcuVarPre.Traiteperiodicite(TobCompteursACalculer, CurrentSal, DateEntreeSalarie, DateSortieSalarie,
    DateDebutJour, DateDebutHebdo, DateDebutMens, DateDebutAnn, DateDebutFinCycle,
    DateFinJour, DateFinHebdo, DateFinMens, DateFinAnn, DateFinFinCycle);
  TobValeursCompteursSalarie := ObjCalcuVarPre.GetTobValeursCompteursSalarie(CurrentSal); // Salarie);
{  LogMessage(diag,'TobValeursCompteursSalarie.count '+IntToStr(TobValeursCompteursSalarie.Detail.Count));}
  { Boucle sur les compteurs }
  for indexCompteur := 0 to TobCompteursACalculer.FillesCount(0)-1 do
  begin
{  LogMessage(diag,'TobCompteurs X '+IntToStr(indexCompteur)); }
    TobCompteurACalculer := TobCompteursACalculer.Detail[indexCompteur];
{  LogMessage(diag,'TobCompteurs X - Chargt ok'); }
    periodicite := TobCompteurACalculer.GetString('PYR_PERIODICITEPRE');
{  LogMessage(diag,'periodicite '+periodicite);
  LogMessage(diag,'CalculJour '+BoolToStr(ObjCalcuVarPre.CalculJour));
  LogMessage(diag,'CalculHebdo '+BoolToStr(ObjCalcuVarPre.CalculHebdo));
  LogMessage(diag,'CalculFinCycle '+BoolToStr(ObjCalcuVarPre.CalculFinCycle));
  LogMessage(diag,'CalculMens '+BoolToStr(ObjCalcuVarPre.CalculMens));
  LogMessage(diag,'CalculAnn '+BoolToStr(ObjCalcuVarPre.CalculAnn));  }
    if (not ObjCalcuVarPre.CalculJour)     and (periodicite = periodiciteJournaliere) then continue;
    if (not ObjCalcuVarPre.CalculHebdo)    and (periodicite = periodiciteHebdomadaire) then continue;
    if (not ObjCalcuVarPre.CalculFinCycle) and (periodicite = periodiciteFinDeCycle) then continue;
    if (not ObjCalcuVarPre.CalculMens)     and (periodicite = periodiciteMensuelle) then continue;
    if (not ObjCalcuVarPre.CalculAnn)      and (periodicite = periodiciteAnnuelle) then continue;
    periodeRaz  := TobCompteurACalculer.GetString('PYR_PERIODERAZ');
    debut       := TobCompteurACalculer.GetDateTime('DATEDEBUT');
    fin         := TobCompteurACalculer.GetDateTime('DATEFIN');
    Compteur    := TobCompteurACalculer.GetString('PYR_COMPTEURPRES');
//    LogMessage(Diag,'  - Compteur '+Compteur+' du '+DateToStr(debut)+' au '+DateToStr(fin));
    IsRegul := False;
//    Value := ObjCalcuVarPre.EvalVarPreCompPres(CurrentSal, Compteur, periodicite, debut, fin, TOB_Rub, NeedSave, HadRegul, Indicateur, Diag);
    Value := ObjCalcuVarPre.EvalVarPreCompPres(CurrentSal, Compteur, periodicite, debut, fin, TOB_Rub, NeedSave, HadRegul, Indicateur, nil );
    LogMessage(Diag,'  - Compteur '+Compteur+' du '+DateToStr(debut)+' au '+DateToStr(fin)+', Resultat : '+FloatToStr(Value));
    ValueToSave := Value;
    if (Indicateur = 'INP') or (Indicateur = 'AIN') then { Si le compteur a déjà été passé en paie }
    begin
      if Recalcul then  { Si on a demandé le recalcul }
      begin
        { On concatène la régule avec le compteur passé en paie }
        if HadRegul then
          ObjCalcuVarPre.ConcatRegul(CurrentSal, Compteur, periodicite, debut, fin);
        { Recalcul du compteur et calcul de la valeur à sauvegarder }
//        Value := ObjCalcuVarPre.EvalVarPreCompPres(CurrentSal, Compteur, periodicite, debut, fin, TOB_Rub, NeedSave, HadRegul, Indicateur, Diag, True);
        Value := ObjCalcuVarPre.EvalVarPreCompPres(CurrentSal, Compteur, periodicite, debut, fin, TOB_Rub, NeedSave, HadRegul, Indicateur, nil, True);
        ValueToSave :=  Value - ValueToSave;
        IsRegul := True;
      end;
    end;
    { Mise à jour du compteur avec la valeur calculée }
    if NeedSave then SaveValue(TobSalarie, TobCompteurACalculer, ValueToSave, IsRegul);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 08/08/2007
Modifié le ... :   /  /
Description .. : Conversion des préfixes de colonnes
Mots clefs ... :
*****************************************************************}
function ConvertPrefixe(StWhere, DePref, APref: string): string;
var
  pospref: integer;
begin
  if StWhere <> '' then
    while Pos(DePref, StWhere) > 0 do
    begin
      pospref := Pos(DePref, StWhere);
      StWhere[(pospref + 1)] := APref[2];
      StWhere[(pospref + 2)] := APref[3];
    end;
  result := Trim(StWhere);
end;

destructor TCompteursCalculator.Destroy;
begin
  FreeAndNil(diag);
  FreeAndNil(TobCompteursARecalculer);
  inherited;
end;

//Function TCompteursCalculator.LancementDuCalcul(WhereSalaries : String; CalculDesCompteursMensuels : Boolean; DateDebutCalcul, DateFinCalcul : TDateTime): Boolean;
Function TCompteursCalculator.LancementDuCalcul(WhereSalaries : String;
    CalculDesCompteursJournaliers, CalculDesCompteursHebdomadaires, CalculDesCompteursFinDeCycle,
    CalculDesCompteursMensuels, CalculDesCompteursAnnuels : Boolean;
    RecalcEnCours : Integer;
    ForceRecalcul : Boolean;
    DateDeDebutPourTousLesSalaries :TDateTime;
    DateSaisieDebutJournaliers, DateSaisieDebutHebdomadaires, DateSaisieDebutFinDeCycle, DateSaisieDebutMensuels, DateSaisieDebutAnnuels : TDateTime;
    DateFinJournaliers, DateFinHebdomadaires, DateFinFinDeCycle, DateFinMensuels, DateFinAnnuels : TDateTime;
    DateDebutCalcul, DateFinCalcul : TDateTime) : Boolean;
var
  TobSalaries, TobSalarie, TobProfilsSalaries, TobSalariePopul, TobProfils, TobProfil, TempTob : Tob;
  TobCompteursProfils, TobParametresAssocPopul, TobParametres, TobParametrePopul, TobParametresPopul, Tob_Rub,  tobParametre : Tob;
  indexSalarie, indexProfil, indexCompteur, indexCompteurFille, indexCompteurPetiteFille : integer;
  indexParametre : Integer;
  TypeProfilPres, ProfilPres, parametrepop, SalariePop : String;
  DateFinEstFinDeMois : Boolean;
  StSQL : String;
  StListResult : TStringList;


  TobValComptSal : Tob;

//{$IFDEF TESTGGU}
//    F: TextFile;
//{$ENDIF TESTGGU}
begin
{  LogMessage(diag,'CalculJour '+BoolToStr(CalculDesCompteursJournaliers));
  LogMessage(diag,'CalculHebdo '+BoolToStr(CalculDesCompteursHebdomadaires));
  LogMessage(diag,'CalculFinCycle '+BoolToStr(CalculDesCompteursFinDeCycle));
  LogMessage(diag,'CalculMens '+BoolToStr(CalculDesCompteursMensuels));
  LogMessage(diag,'CalculAnn '+BoolToStr(CalculDesCompteursAnnuels));  }
{$IFNDEF EAGLSERVER}
  InitMoveProgressForm(Nil, 'Calcul des compteurs de présence', 'Initialisation', 8, False, True);
{$ENDIF}
//{$IFDEF TESTGGU}
//        AssignFile(F, 'C:\LogGGU.txt');  { fichier suivant dans la propriété Files }
//        Append(F);
//        Writeln(F, 'Initialisation');
//        Flush(f);  { Pour que le texte soit réellement écrit dans le fichier }
//        CloseFile(F);
//{$ENDIF TESTGGU}
  SelectionUtilisateur := (WhereSalaries <> '');
  DateSaisieDebutJour := DateSaisieDebutJournaliers;
  DateSaisieDebutHebdo := DateSaisieDebutHebdomadaires;
  DateSaisieDebutFinCycle := DateSaisieDebutFinDeCycle;
  DateSaisieDebutMens := DateSaisieDebutMensuels;
  DateSaisieDebutAnn := DateSaisieDebutAnnuels;
  DateFinJour := DateFinJournaliers;
  DateFinHebdo := DateFinHebdomadaires;
  DateFinFinCycle := DateFinFinDeCycle;
  DateFinMens := DateFinMensuels;
  DateFinAnn := DateFinAnnuels;
  RecalculEnCours := RecalcEnCours;
//  RendPeriodeEnCours(ExerPerEnCours, DebPer, FinPer);
  Recalcul := ForceRecalcul;
  DateDebTousSalaries :=  DateDeDebutPourTousLesSalaries;

{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement des salariés');
{$ENDIF}
//{$IFDEF TESTGGU}
//        AssignFile(F, 'C:\LogGGU.txt');  { fichier suivant dans la propriété Files }
//        Append(F);
//        Writeln(F, 'Chargement des salariés');
//        Flush(f);  { Pour que le texte soit réellement écrit dans le fichier }
//        CloseFile(F);
//{$ENDIF TESTGGU}
  LogMessage(Diag,'Chargement des salariés. '+TimeToStr(Time));
  TobSalaries := Tob.Create('Liste des salariés', nil, -1);
  StSQL := 'SELECT PSA_SALARIE FROM SALARIES ' + WhereSalaries;
  TobSalaries.LoadDetailFromSQL(stSQL);
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement des données de la paie...');
{$ENDIF}
//{$IFDEF TESTGGU}
//        AssignFile(F, 'C:\LogGGU.txt');  { fichier suivant dans la propriété Files }
//        Append(F);
//        Writeln(F, 'Chargement des données de la paie...');
//        Flush(f);  { Pour que le texte soit réellement écrit dans le fichier }
//        CloseFile(F);
//{$ENDIF TESTGGU}
  LogMessage(Diag,'Chargement des données de la paie. '+TimeToStr(Time));
  InitLesTOBPaie;
  ChargeLesTOBPaie;
  DateFinEstFinDeMois := CalculDesCompteursMensuels;// and (not recalcul);
//     if GetControlText('SELECTEDSAL') <> '' then
//       ChargePresence(GetControlText('SELECTEDSAL'), Deb, Fin, DateFinEstFinDeMois )
//     else if GetControlText('PSA_SALARIE') <> '' then
//       ChargePresence('WHERE PSA_SALARIE = "'+GetControlText('PSA_SALARIE')+'"', Deb, Fin, DateFinEstFinDeMois)
//     else

  { DONE 3 -oGGU -cPrésence : Tenir compte des flag de recalcul pour la date de chargement }
  if (DateDebTousSalaries <> iDate2099) and (DateDebTousSalaries < DateDebutCalcul) then
    ChargePresence(WhereSalaries, DateDebTousSalaries, DateFinCalcul, DateFinEstFinDeMois)
  else
    ChargePresence(WhereSalaries, DateDebutCalcul, DateFinCalcul, DateFinEstFinDeMois);
  ObjCalcuVarPre.CalculJour     := CalculDesCompteursJournaliers;
  ObjCalcuVarPre.CalculHebdo    := CalculDesCompteursHebdomadaires;
  ObjCalcuVarPre.CalculFinCycle := CalculDesCompteursFinDeCycle;
  ObjCalcuVarPre.CalculMens     := CalculDesCompteursMensuels;
  ObjCalcuVarPre.CalculAnn      := CalculDesCompteursAnnuels;
//       ChargePresence(RecupWhereCritere(TPageControl(GetControl('Pages'))), Deb, Fin, DateFinEstFinDeMois);
  { Si calcul des mensuels on force le calcul annuel quand le salarié sort dans le mois
  //@@GGU  -> Quelles dates prendre pour le calcul des annuels ?
  }
  initTOB_EtabCompl();
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement des profils salariés...');
{$ENDIF}
  LogMessage(Diag,'Chargement des profils salariés. '+TimeToStr(Time));
  TobProfilsSalaries := Tob.Create('profils salariés', nil, -1);
  stSQL := 'SELECT PPZ_SALARIE, PPZ_DATEVALIDITE, PPZ_TYPPROFILPRES, PPZ_PROFILPRES FROM PROFILPRESSALARIE';
  ObjCalcuVarPre.AddFiltreSalarie('PPZ', stSQL);
  TobProfilsSalaries.LoadDetailFromSQL(stSQL);
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement des associations Compteurs - profils...');
{$ENDIF}
  LogMessage(Diag,'Chargement des associations compteurs-profils. '+TimeToStr(Time));
  TobCompteursProfils := Tob.Create('Compteurs - Profils', nil, -1);
  TobCompteursProfils.LoadDetailFromSQL('SELECT PPV_PROFILPRES, PPV_DATEVALIDITE, PPV_COMPTEURPRES, PPV_LIBELLE FROM PROFILCOMPTEURPRES');
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement de la table d''association Salarié - Population...');
{$ENDIF}
  LogMessage(Diag,'Chargement des associations salarié-population. '+TimeToStr(Time));
  TobSalariePopul := Tob.create('table d''association Salarié - Population', nil, -1);
  stSQL := 'SELECT PNA_SALARIE, PNA_POPULATION FROM SALARIEPOPUL WHERE PNA_TYPEPOP = "PRE"';
  ObjCalcuVarPre.AddFiltreSalarie('PNA', stSQL);
  TobSalariePopul.LoadDetailDBFromSQL('SALARIEPOPUL',stSQL);
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement de la table d''association Population - paramètres...');
{$ENDIF}
  LogMessage(Diag,'Chargement des associations population-paramètres. '+TimeToStr(Time));
  TobParametresAssocPopul := Tob.create('table d''association Population - parametres', nil, -1);
  TobParametresAssocPopul.LoadDetailDBFromSQL('PGPARAMETRESASSOC','SELECT PGO_CODEASSOCIE, PGO_PGPARAMETRE, PGO_DATEVALIDITE FROM PGPARAMETRESASSOC WHERE ##PGO_PREDEFINI## AND PGO_TYPEPARAMETRE = "POP" AND PGO_TYPEPOP = "PRE"');
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Chargement de la table des paramètres population...');
{$ENDIF}
  LogMessage(Diag,'Chargement des paramètres population. '+TimeToStr(Time));
  TobParametresPopul := Tob.create('table des parametres population', nil, -1);
  TobParametresPopul.LoadDetailDBFromSQL('PGPARAMETRES','SELECT PGP_PGPARAMETRE, PGP_LIBELLE, PGP_PGNOMCHAMP, PGP_PGVALCHAMP FROM PGPARAMETRES WHERE ##PGP_PREDEFINI## AND PGP_TYPEPARAMETRE = "POP" AND PGP_MODULECHAMP = "PRE"');
{$IFNDEF EAGLSERVER}
  FiniMoveProgressForm;
{$ENDIF}
  BEGINTRANS;
  try
    { Calcul des compteurs }
{$IFNDEF EAGLSERVER}
    InitMoveProgressForm(nil, 'Calcul des compteurs de présence', 'Calcul en cours...', TobSalaries.FillesCount(0) + 2, False, True);
{$ENDIF}
    try
      for indexSalarie := 0 to TobSalaries.FillesCount(0) -1 do
      begin
        CurrentSal        := TobSalaries.Detail[indexSalarie].GetString('PSA_SALARIE');
        DateDebutJour     := DateSaisieDebutJour;
        DateDebutHebdo    := DateSaisieDebutHebdo;
        DateDebutFinCycle := DateSaisieDebutFinCycle;
        DateDebutMens     := DateSaisieDebutMens;
        DateDebutAnn      := DateSaisieDebutAnn;
{$IFNDEF EAGLSERVER}
        MoveCurProgressForm(CurrentSal);
{$ENDIF}
        LogMessage(Diag,'Salarié : '+CurrentSal+' - '+TimeToStr(Time));
        RecupTobSalarie(CurrentSal, DateDebutCalcul, DateFinCalcul);
        TobSalarie := TOB_Salarie;
        DateEntreeSalarie := TobSalarie.GetDateTime('PSA_DATEENTREE');
        DateSortieSalarie := TobSalarie.GetDateTime('PSA_DATESORTIE');
        { Chargement des rubriques du salarié }
        try
          Tob_Rub := ChargeRubriqueSalarie(TobSalarie, DateDebutCalcul, DateFinCalcul, taCalcul, 0, '');
        except
        end;
        if Assigned(Tob_Rub) then
        begin
          { Prise en compte des flags "A recalculer" }
          TraiteARecalculer(ObjCalcuVarPre.GetTobValeursCompteursSalarie(CurrentSal)); //, CurrentSal);
          { Recherche des profils du salarié }
          TobProfils := GetAllValidesFromTobValidite(TobProfilsSalaries, ['PPZ_SALARIE'], [CurrentSal],
                                                       'PPZ_DATEVALIDITE', 'DATEDEBUT', DateDebutJour, 'DATEFIN', DateFinCalcul);
          { Boucle sur les profils }
          for indexProfil := 0 to TobProfils.FillesCount(0) -1 do
          begin
            TobProfil := TobProfils.Detail[indexProfil];
            TypeProfilPres := TobProfil.GetString('PPZ_TYPPROFILPRES');
            if TypeProfilPres = 'PER' then
            begin
              ProfilPres := TobProfil.GetString('PPZ_PROFILPRES');
              LogMessage(Diag,' - Profil '+ProfilPres+' de type '+TypeProfilPres);
              CalculCompteursProfil(TobCompteursProfils, ObjCalcuVarPre.TobCompteurs, Tob_Rub, TobSalarie, ProfilPres, TobProfil.GetDateTime('DATEDEBUT'), TobProfil.GetDateTime('DATEFIN'));
            end else if TypeProfilPres = 'POP' then
            begin
              { Recherche de la population du salarié }
              TempTob := TobSalariePopul.FindFirst(['PNA_SALARIE'], [CurrentSal], False);
              if Assigned(TempTob) then
              begin
                SalariePop := TempTob.GetString('PNA_POPULATION');
                { Recherche du parametre de la population }
                TobParametres := GetAllValidesFromTobValidite(TobParametresAssocPopul, ['PGO_CODEASSOCIE'], [SalariePop], 'PGO_DATEVALIDITE', 'DATEDEBUT', TobProfil.GetDateTime('DATEDEBUT'), 'DATEFIN', TobProfil.GetDateTime('DATEFIN') );
                for indexParametre := 0 to TobParametres.FillesCount(0)-1 do
                begin
                  tobParametre := TobParametres.Detail[indexParametre];
                  parametrepop := tobParametre.GetString('PGO_PGPARAMETRE');
                  { Recherche du profil de présence du parametre }
                  TobParametrePopul := TobParametresPopul.FindFirst(['PGP_PGPARAMETRE', 'PGP_PGNOMCHAMP'], [parametrepop, 'PROFILPRES'], False);
                  if Assigned(TobParametrePopul) then
                  begin
                    ProfilPres := TobParametrePopul.GetString('PGP_PGVALCHAMP');
                    LogMessage(Diag,' - Profil '+ProfilPres+' de type '+TypeProfilPres);
                    CalculCompteursProfil(TobCompteursProfils, ObjCalcuVarPre.TobCompteurs, Tob_Rub, TobSalarie, ProfilPres, TobProfil.GetDateTime('DATEDEBUT'), TobProfil.GetDateTime('DATEFIN'));
                  end;
                end;
                FreeAndNil(TobParametres);
              end;
            end;
          end;
          FreeAndNil(Tob_Rub);
        end;
        FreeAndNil(TobProfils);
        if Assigned(TOB_DUSALARIE) then FreeAndNil(TOB_DUSALARIE);
        { On vide les infos du salarié (Sauf les compteurs calculés à mettre en base) }
        ObjCalcuVarPre.FreeMemSalarie(CurrentSal);
        { Mise à jour de la base pour le salarié courant }
        TobValComptSal := ObjCalcuVarPre.GetTobValeursCompteursSalarie(CurrentSal);
        TobValComptSal.InsertDB(nil);
        { On vide les compteurs calculés du salarié }
        TobValComptSal.ChangeParent(nil, -1);
        FreeAndNil(TobValComptSal);
      end;
    finally
      LogMessage(Diag,'Libération du salarié. '+TimeToStr(Time));
      FreeAndNil(TobSalariePopul);
      FreeAndNil(TobParametresAssocPopul);
      FreeAndNil(TobParametresPopul);
      FreeAndNil(TobCompteursProfils);
      FreeAndNil(TobProfilsSalaries);
      Nettoyage_Etablissement();
      VideLesTOBPaie(False);
      FreeAndNil(TobSalaries);
    end;
//{$IFNDEF EAGLSERVER}
//    MoveCurProgressForm('Mise à jour de la base');
//{$ENDIF}
//    LogMessage(Diag,'Mise à jour de la base. '+TimeToStr(Time));
    { Suppression des compteurs-flag "A recalculer" }
    if not SelectionUtilisateur then
    begin
      ExecuteSQL('Delete from presencesalarie where PYP_PGINDICATPRES = "ARE"');
    end;
    { Suppression de la tob des compteurs non mis à jour (sinon, insertDB essaye de les insérer à nouveau -> violation de clef) }
{    for indexCompteur := ObjCalcuVarPre.TobValeursCompteurs.FillesCount(0)-1 downto 0 do
    begin
      for indexCompteurFille := ObjCalcuVarPre.TobValeursCompteurs.Detail[indexCompteur].detail.count -1 downto 0 do
      begin
        for indexCompteurPetiteFille := ObjCalcuVarPre.TobValeursCompteurs.Detail[indexCompteur].Detail[indexCompteurFille].detail.count -1 downto 0 do
        begin
          TempTob := ObjCalcuVarPre.TobValeursCompteurs.Detail[indexCompteur].Detail[indexCompteurFille].Detail[indexCompteurPetiteFille];
          if (not TempTob.IsOneModifie) or (TempTob.GetDouble('PYP_QUANTITEPRES') = 0) then
          begin
            TempTob.ChangeParent(nil,-1);
            FreeAndNil(TempTob);
          end;
        end;
//        TempTob := ObjCalcuVarPre.TobValeursCompteurs.Detail[indexCompteur].Detail[indexCompteurFille];
//        if (not TempTob.IsOneModifie) or (TempTob.GetDouble('PYP_QUANTITEPRES') = 0) then
//        begin
//          TempTob.ChangeParent(nil,-1);
//          FreeAndNil(TempTob);
//        end;
      end;
    end;
    ObjCalcuVarPre.TobValeursCompteurs.InsertDB(nil); }
    { Sauvegarde de la date de fin de traitement  }
    // Attention : Pas de sauvegarde de date si des critères ont été positionnés et qu'un recalcul n'a pas été lancé globalement
    If Not SelectionUtilisateur Then
    Begin
         PresenceSauveJourCalcule(DateFinCalcul);
//         If CalculMens Then PresenceSauveMoisCalcule(Fin, Recalcul);
         If Recalcul Then RecalculEnCours := 0;
    End;
    COMMITTRANS;
    result := True;
  except
    on E : Exception do begin
      result := False;
      ROLLBACK;
      LogMessage(Diag,'Erreur lors du traitement - Les modifications ont été annulées');
      LogMessage(Diag,'Message d''erreur : '+E.Message);
    end;
  end;
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Libération des objets en mémoire');
{$ENDIF}
  LogMessage(Diag,'Libération de la mémoire. '+TimeToStr(Time));
//  FreeAndNil(TobSalariePopul);
//  FreeAndNil(TobParametresAssocPopul);
//  FreeAndNil(TobParametresPopul);
//  FreeAndNil(TobCompteursProfils);
//  FreeAndNil(TobProfilsSalaries);
//  Nettoyage_Etablissement();
//  VideLesTOBPaie(False);
  VideLaPresence;
//  FreeAndNil(TobSalaries);
{$IFNDEF EAGLSERVER}
  FiniMoveProgressForm;
{$ENDIF}
  LogMessage(Diag,'Traitement terminé. '+TimeToStr(Time));
  { DONE 3 -oGGU -cPrésence : Ajouter dans les tablettes de famille d'évènement PGFAMEVT
et d'évènement PGEVT les codes et les libellés de la présence }
  StListResult := LogToStrings(Diag);
  if result then
    CreeJnalEvt('006', '220', 'OK', nil, nil, StListResult)
  else
    CreeJnalEvt('006', '220', 'ERR', nil, nil, StListResult);
  FreeAndNil(StListResult);
end;

constructor TCompteursCalculator.Create(ObjDiag : TObject = nil);
begin
  diag := ObjDiag;
  if not Assigned(Diag) then
    diag := Tob.Create('Log du calcul des compteurs', nil, -1);
  RecalculEnCours := 0;
  { Chargement des lignes de compteurs "ARE" (A recalculer) }
  TobCompteursARecalculer := Tob.Create('tob PRESENCESALARIE', nil, -1);
  TobCompteursARecalculer.LoadDetailDBFromSQL('PRESENCESALARIE', 'Select PYP_SALARIE, PYP_DATEDEBUTPRES, '
        +'PYP_DATEFINPRES, PYP_COMPTEURPRES, PYP_PERIODICITEPRE, PYP_THEMEPRE, '
        +'PYP_ETATPRES, PYP_PGINDICATPRES, PYP_TYPECALPRES '
        +' from presencesalarie where PYP_PGINDICATPRES = "ARE"'); 
end;

function TCompteursCalculator.LancementDuCalcul(
  TobParametres: Tob): Boolean;
  Function S2B(sStr : String) : Boolean;
  begin
    if sStr = 'X' then result := True else result := False;
  end;
begin
{  LogMessage(diag,'CalculJour '+BoolToStr(TobParametres.GetBoolean('CalculDesCompteursJournaliers')));
  LogMessage(diag,'CalculHebdo '+BoolToStr(TobParametres.GetBoolean('CalculDesCompteursHebdomadaires')));
  LogMessage(diag,'CalculFinCycle '+BoolToStr(TobParametres.GetBoolean('CalculDesCompteursFinDeCycle')));
  LogMessage(diag,'CalculMens '+BoolToStr(TobParametres.GetBoolean('CalculDesCompteursMensuels')));
  LogMessage(diag,'CalculAnn '+BoolToStr(TobParametres.GetBoolean('CalculDesCompteursAnnuels')));
  LogMessage(diag,'CalculJour '+TobParametres.GetString('CalculDesCompteursJournaliers'));
  LogMessage(diag,'CalculHebdo '+TobParametres.GetString('CalculDesCompteursHebdomadaires'));
  LogMessage(diag,'CalculFinCycle '+TobParametres.GetString('CalculDesCompteursFinDeCycle'));
  LogMessage(diag,'CalculMens '+TobParametres.GetString('CalculDesCompteursMensuels'));
  LogMessage(diag,'CalculAnn '+TobParametres.GetString('CalculDesCompteursAnnuels'));
}
  if Assigned(TobParametres) then
    result := LancementDuCalcul(
      TobParametres.GetString('WhereSalaries'),
      S2B(TobParametres.GetString('CalculDesCompteursJournaliers')),
      S2B(TobParametres.GetString('CalculDesCompteursHebdomadaires')),
      S2B(TobParametres.GetString('CalculDesCompteursFinDeCycle')),
      S2B(TobParametres.GetString('CalculDesCompteursMensuels')),
      S2B(TobParametres.GetString('CalculDesCompteursAnnuels')),
      TobParametres.GetInteger('RecalcEnCours'),
      S2B(TobParametres.GetString('ForceRecalcul')),
      TobParametres.GetDateTime('DateDeDebutPourTousLesSalaries'),
      TobParametres.GetDateTime('DateSaisieDebutJournaliers'),
      TobParametres.GetDateTime('DateSaisieDebutHebdomadaires'),
      TobParametres.GetDateTime('DateSaisieDebutFinDeCycle'),
      TobParametres.GetDateTime('DateSaisieDebutMensuels'),
      TobParametres.GetDateTime('DateSaisieDebutAnnuels'),
      TobParametres.GetDateTime('DateFinJournaliers'),
      TobParametres.GetDateTime('DateFinHebdomadaires'),
      TobParametres.GetDateTime('DateFinFinDeCycle'),
      TobParametres.GetDateTime('DateFinMensuels'),
      TobParametres.GetDateTime('DateFinAnnuels'),
      TobParametres.GetDateTime('DateDebutCalcul'),
      TobParametres.GetDateTime('DateFinCalcul'))
  else
    result := False;
end;

{$IFNDEF EAGLSERVER}
procedure TOF_PRESCALCCOMPTEURS.ClickAddJob(Sender: TObject);
begin
  Calcul(True);
end;
{$ENDIF}

{$IFNDEF EAGLSERVER}
procedure TOF_PRESCALCCOMPTEURS.Calcul(AsJob: Boolean = False);
var
  Deb, Fin : TDateTime;
  Q : TQuery;
  Salarie, st : string;
  i: integer;
  Q_Mul: THQuery;
  ListeParam, Param : String;
  stSQL : String;
  SaisieDebutJour, SaisieDebutHebdo, SaisieDebutFinCycle, SaisieDebutMens, SaisieDebutAnn : TDateTime;
  SaisieFinJour, SaisieFinHebdo, SaisieFinFinCycle, SaisieFinMens, SaisieFinAnn : TDateTime;
  TobParam, T1: TOB;
  procedure RecupSaisUtilisateur;
    procedure RecupDatesCalcul(Recup : Boolean; var DateDeb, DateFin : TDateTime; NomControlDeb, NomControlFin : String );
    begin
      if Recup then
      begin
        DateDeb     := AGLStrToDate(GetControlText(NomControlDeb));
        DateFin     := AGLStrToDate(GetControlText(NomControlFin));
      end else begin
        DateDeb     := iDate2099;
        DateFin     := iDate1900;
      end;
    end;
  begin
    { Récupération de la saisie de l'utilisateur }
    RecupDatesCalcul(CalculJour,     SaisieDebutJour,     SaisieFinJour,     'EDDATEDEBUT',      'EDDATEFIN');
    RecupDatesCalcul(CalculHebdo,    SaisieDebutHebdo,    SaisieFinHebdo,    'EDDATEDEBUT',      'EDDATEFIN');
    RecupDatesCalcul(CalculFinCycle, SaisieDebutFinCycle, SaisieFinFinCycle, 'EDDATEDEBUT',      'EDDATEFIN');
    RecupDatesCalcul(CalculMens,     SaisieDebutMens,     SaisieFinMens,     'EDDATEDEBUT',      'EDDATEFIN');
    RecupDatesCalcul(CalculAnn,      SaisieDebutAnn,      SaisieFinAnn,      'EDDATEDEBUTANNEE', 'EDDATEFINANNEE');
  end;
  Function B2S(bBool : Boolean) : String;
  begin
    if bBool then result := 'X' else result := '-';
  end;
begin
  Inherited ;
     { Gestion de la sélection de salarié }
     Q_Mul := THQuery(Ecran.FindComponent('Q'));
     if Q_Mul = nil then exit;

     If (TFMul(Ecran).FListe.nbSelected > 0) Then
     Begin
          If PgiAsk(TraduireMemoire('Voulez-vous calculer les compteurs pour les salariés sélectionnés?'), TraduireMemoire('Calcul des compteurs de présence')) = mrYes Then
          Begin
               St := '';

               { Composition du clause WHERE pour limiter le mul à ces salariés }
               for i := 0 to TFMul(Ecran).FListe.NbSelected - 1 do
               begin
                    {$IFDEF EAGLCLIENT}
                         TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
                    {$ENDIF}
                    TFMul(Ecran).FListe.GotoLeBookmark(i);

                    Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').AsString;
                    St := St + ' PSA_SALARIE="' + Salarie + '" OR';
               end;

               TFMul(Ecran).FListe.ClearSelected;
               
               if St <> '' then St := ' AND (' + Copy(St, 1, Length(st) - 2) + ')';
               SetControlText('SELECTEDSAL', St);
               TFMul(Ecran).BCherche.Click;

               { Récupération de la Query pour traitement }
               {$IFDEF EAGLCLIENT}
                    if TFMul(Ecran).Fetchlestous then TheMulQ := TOB(Ecran.FindComponent('Q'));
               {$ELSE}
                    TheMulQ := THQuery(Ecran.FindComponent('Q'));
               {$ENDIF}
          End
          Else
               Exit;
     End;

     If GetControlText('CBRECALCUL') = 'X' Then Recalculer := True Else Recalculer := False;

     { Recherche d'un éventuel compteur "A recalculer" pour l'ensemble des salariés }

     DebPourTousSalaries := iDate2099;
     Q := OpenSql('select min(PYP_DATEDEBUTPRES) as DEBUT, count(*) as X from presencesalarie where PYP_PGINDICATPRES = "ARE" and PYP_SALARIE = "TOUS"', True);
     if not Q.Eof then
     begin
          if (Q.FindField('X').AsInteger >= 1) and (Q.FindField('DEBUT').AsDateTime <= StrToDate(GetControlText('EDDATEDEBUT'))) then
               DebPourTousSalaries := Q.FindField('DEBUT').AsDateTime;
     end;
     Ferme(Q);

     // Récupération des critères saisis
     StSQL := RecupWhereCritere(TPageControl(GetControl('Pages')));
     If ((StSQL <> '') Or (GetControlText('SELECTEDSAL') <> '')) Then CriteresSaisis := True Else CriteresSaisis := False;

     // Teste si un recalcul général est nécessaire
     if (DebPourTousSalaries <> iDate2099) And ((DebPourTousSalaries < StrToDate(GetControlText('EDDATEDEBUT'))) Or (CriteresSaisis)) then
     begin
          if PGIAsk('Les compteurs de tous les salariés doivent être recalculés depuis le '+AGLDateToStr(DebPourTousSalaries)+'.#10#13Voulez vous effectuer ce traitement ?', Ecran.Caption) = mrNo then
          begin
               FiniMoveProgressForm;
               exit;
          end else begin
               Recalculer := True;
               // En cas de recalcul général, on supprime le filtre saisi par l'utilisateur
               If (GetControlText('SELECTEDSAL') <> '') Then
               Begin
                    SetControlText('SELECTEDSAL', '');
                    St := '';
               End;
               If StSQL <> '' Then
               Begin
                    SetControlText('PSA_SALARIE', '');
                    SetControlText('PSA_LIBELLE', '');
                    SetControlText('PSA_ETABLISSEMENT', '');
                    SetControlText('CKSORTIE', '-');
                    SetControlText('PSA_TRAVAILN1', '');
                    SetControlText('PSA_TRAVAILN2', '');
                    SetControlText('PSA_TRAVAILN3', '');
                    SetControlText('PSA_TRAVAILN4', '');
                    SetControlText('PSA_CODESTAT', '');
                    StSQL := '';
               End;
               // On modifie la date de début de traitement affichée
               SetControlText('EDDATEDEBUT', DateToStr(DebPourTousSalaries));
               TheMulQ := nil;
               TFMul(Ecran).BCherche.Click;
          End;
     end;

     Deb := StrToDate(GetControlText('EDDATEDEBUT'));
     Fin := StrToDate(GetControlText('EDDATEFIN'));

     if ((DebPourTousSalaries <> iDate2099) and (DebPourTousSalaries < Fin)) or Recalculer then
     begin
       if PGIAsk('Les compteurs postérieurs au '+DateToStr(Fin)+' seront supprimés pour les salariés sélectionnés et ne seront pas recalculés.#10#13#10#13 Voulez-vous continuer ?', Ecran.Caption) = mrNo then
         exit;
     end;
     CalculJour     := True;
     CalculHebdo    := True;
     CalculFinCycle := True;
     CalculMens     := (GetControlText('CBCALCCOMPTMENS') = 'X');
     CalculAnn      := (GetControlText('CBCALCCOMPTANNU') = 'X');
     { Si on a un flag de recalcul et qu'on doit recalculer les compteurs depuis une date déjà clôturée
       on calcul les mensuels, même si ce n'est pas coché }
     if (not CalculMens) and (DebPourTousSalaries <> iDate2099) then
     begin
       ListeParam := GetParamSocDatesCloture(MOIS); //PT1 GetParamSocSecur(PresenceRenvoieParamSoc(MOIS), '');
       Param := ReadTokenSt(ListeParam);
       While (Param <> '') Do
       Begin
         if StrToDate(Param) >= DebPourTousSalaries then
         begin
           CalculMens := True;
           Break;
         end;
         Param := ReadTokenSt(ListeParam);
       End;
     end;

     RecupSaisUtilisateur;

     { Contrôle des périodes }
     If (Fin < Deb) Or (CalculAnn And (Calculator.DateFinAnn < Calculator.DateSaisieDebutAnn)) Then
     Begin
          PGIBox(TraduireMemoire('La date de début doit être inférieure à la date de fin.'));
          Exit;
     End;

     { Chargement des données }

     // En mode SAV on affiche l'onglet de détails
     If Assigned(Diag) Then
     Begin
          (GetControl('PAGES') as THPageControl2).ActivePageIndex := 4; //4e onglet après SQL
          LBDetails.Items.Clear;
     End;  

     if GetControlText('SELECTEDSAL') <> '' then
       stSQL := 'WHERE '+GetControlText('SELECTEDSAL')
     else if GetControlText('PSA_SALARIE') <> '' then
       stSQL := 'WHERE PSA_SALARIE = "'+GetControlText('PSA_SALARIE')+'"';
//     else
//       stSQL := stSQL;

     TobParam := TOB.create('Ma Tob de Param', nil, -1);
     T1 := TOB.Create('XXX', TobParam, -1);
     T1.AddChampSupValeur('WhereSalaries',                  stSQL);
     {08/11/2007 GGU
      On transforme les boolean en strings, sinon ils ne sont pas récupérés comme ils faut en process serveur
      detail du probleme :
      On met un boolean à True dans la tob qui est stockée 'True' dans la tob
      il est récupéré comme une chaine de caractère 'True' par le process Serveur
      Donc quand on fait un GetBoolean, on récupère False....
      Alors qu'en execution normal, l'affectation du boolean puis la récupération
      par getBoolean ne pose pas de problème.
      Le problème se situe à priori lors du stockage de la tob dans le job }
     T1.AddChampSupValeur('CalculDesCompteursJournaliers',  B2S(CalculJour));
     T1.AddChampSupValeur('CalculDesCompteursHebdomadaires',B2S(CalculHebdo));
     T1.AddChampSupValeur('CalculDesCompteursFinDeCycle',   B2S(CalculFinCycle));
     T1.AddChampSupValeur('CalculDesCompteursMensuels',     B2S(CalculMens));
     T1.AddChampSupValeur('CalculDesCompteursAnnuels',      B2S(CalculAnn));
     T1.AddChampSupValeur('RecalcEnCours',                  RecalcEnCours);
     T1.AddChampSupValeur('ForceRecalcul',                  B2S(Recalculer));
     T1.AddChampSupValeur('DateDeDebutPourTousLesSalaries', DebPourTousSalaries);
     T1.AddChampSupValeur('DateSaisieDebutJournaliers',     SaisieDebutJour);
     T1.AddChampSupValeur('DateSaisieDebutHebdomadaires',   SaisieDebutHebdo);
     T1.AddChampSupValeur('DateSaisieDebutFinDeCycle',      SaisieDebutFinCycle);
     T1.AddChampSupValeur('DateSaisieDebutMensuels',        SaisieDebutMens);
     T1.AddChampSupValeur('DateSaisieDebutAnnuels',         SaisieDebutAnn);
     T1.AddChampSupValeur('DateFinJournaliers',             SaisieFinJour);
     T1.AddChampSupValeur('DateFinHebdomadaires',           SaisieFinHebdo);
     T1.AddChampSupValeur('DateFinFinDeCycle',              SaisieFinFinCycle);
     T1.AddChampSupValeur('DateFinMensuels',                SaisieFinMens);
     T1.AddChampSupValeur('DateFinAnnuels',                 SaisieFinAnn);
     T1.AddChampSupValeur('DateDebutCalcul',                Deb);
     T1.AddChampSupValeur('DateFinCalcul',                  Fin);
     if AsJob then
     begin
       AGLFicheJob(0, taCreat, 'cgiPaieS5', 'CALCULCOMPTEURSPRESENCE', TobParam);
     end else begin
       if not Calculator.LancementDuCalcul(TobParam.detail[0]) then
//         stSQL,
//         CalculJour, CalculHebdo, CalculFinCycle, CalculMens, CalculAnn,
//         RecalcEnCours, Recalculer, DebPourTousSalaries,
//         SaisieDebutJour, SaisieDebutHebdo, SaisieDebutFinCycle, SaisieDebutMens, SaisieDebutAnn,
//         SaisieFinJour, SaisieFinHebdo, SaisieFinFinCycle, SaisieFinMens, SaisieFinAnn,
//         Deb, Fin) then
       begin
         PGIError('Erreur lors du traitement#10#13Les modifications ont été annulées', Ecran.Caption);
       end;
     end;

     FreeAndNil(TOBParam);

     { Mise à jour de l'affichage avec le jour suivant }
     If Fin <> Date Then
     Begin
          If CalculMens Then
               PresenceDonneMoisCalculActuel(Deb, Fin)
          Else
               PresenceDonneSemaineCalculActuelle(Deb, Fin);

          SetControlText('EDDATEDEBUT', DateToStr(Deb));
          SetControlText('EDDATEFIN',   DateToStr(Fin));
     End;

     // Remise à 0 de la sélection des salariés
     If St <> '' Then
     Begin
          SetControlText('SELECTEDSAL', '');
          TheMulQ := nil;
     End;
     TFMul(Ecran).BCherche.Click;

end;
{$ENDIF}

Initialization
{$IFNDEF EAGLSERVER}
  registerclasses ( [ TOF_PRESCALCCOMPTEURS ] ) ;
{$ENDIF EAGLSERVER}
end.
