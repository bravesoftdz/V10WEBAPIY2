{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 03/06/2003
Modifié le ... : 03/06/2003
Description .. : Fiche de visualisation des travaux en cours sous forme de planning
Mots clefs ... : TOF;TRAVAUX;PLANNING;UTILISATEUR
*****************************************************************}
unit utoftravail; 

Interface

Uses
    StdCtrls, Controls, Classes, forms, sysutils,ComCtrls,Messages,Windows,
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   dbTables, db,FE_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, HPanel, UTOF , vierge, HTB97,
     UTOB,Affaireutil,AglInitGC,M3FP,Saisutil,Grids,
     DicoAF,graphics,EntGC, Ent1,CalcOleGenericAff,
     utilressource,AfUtilArticle,Utilarticle, UTofAfBaseCodeAffaire,
     HPlanning;

Type
  TOF_TRAVAIL = Class (TOF)
  public
        procedure OnNew;                           override ;
        procedure OnArgument (strArgument:String); override ;
        procedure OnClose;                         override ;

//        procedure OnClickJour      (Sender:TObject);
        procedure OnClickSemaine   (Sender:TObject);
        procedure OnClickMois      (Sender:TObject);
        procedure OnClickPrev      (Sender:TObject);
        procedure OnClickNext      (Sender:TObject);
//        procedure OnChangePage     (Sender:TObject);

//        procedure OnClickActInterne (Sender:TObject);
  //      procedure OnClickActExterne (Sender:TObject);
    //    procedure OnClickActAbsence (Sender:TObject);

  protected
           Pages                                    :TPageControl;  // Onglet
           TravauxPlanning                          :THPlanning;  // Planning pour travaux
           PanelTravaux                             :THPanel;

           TOBRessTravaux                           :TOB;
           TOBItemTravaux                           :TOB;       // Les 3 TOBS utiles pour planning travaux
           TOBEtatTravaux                           :TOB;

//           TOBGroupes                               :TOB;         // Liste des groupes de travail "consultables"
//           TOBDroits                                :TOB;         // Liste des droits du user connecté sur les travaux de chaque user

//           strGroupe                               :string;       // Groupe d'utilisateur à traiter (exclusif de strUser)

           DateDebut, DateFin                      :TDateTime;    // Fourchette de date d'analyse
           TypeDate                                :integer;      // 1=heure/jour, 2=jour/semaine, 3=jour/mois

           procedure BuildEtatPlanning;
           function  PlanningValide      :boolean;
           procedure AfficheTravaux      (DatePosition:TDateTime);
           function  AfficheTravailFiche (Item:TOB):boolean; //; strCodeEvt:string):boolean;

           function  UpdateTravail       (Item:TOB):boolean; //FromItem, ToItem:TOB):boolean;
           function  DeleteTravail       (Item:TOB):boolean;

           procedure OnClickValider      (Sender: TObject);
           procedure OnClickImprimer     (Sender: TObject);
           procedure OnClickExporter     (Sender: TObject);
           procedure OnClickMaximiser    (Sender: TObject);
           procedure OnClickRechDossier  (Sender: TObject);

           // Evénements sur les éléments du THPlanning
           procedure OnDblClickItemTravail  (Sender: TObject);
           procedure OnEvtTravail           (Sender:TObject; FromItem:TOB; ToItem:TOB; Actions:THPlanningAction);
           procedure OnModifyItemTravail    (Sender:TObject; Item:TOB; var Cancel:boolean);
           procedure OnDeleteItemTravail    (Sender:TObject; Item:TOB; var Cancel:boolean);
           procedure OnMoveItemTravail      (Sender:TObject; Item:TOB; var Cancel:boolean);
           procedure OnLinkItemTravail      (Sender:TObject; Source,Destination:TOB; Option:THPlanningOptionLink; var Cancel: boolean);

           procedure MaximisePlanning;
  end;

procedure DPLance_Travaux ();

function DebutDeSemaine (DateDep:TDateTime):TDateTime;


Implementation

uses dialogs, afactivite, menus, dpOutilsAgenda;

function DebutDeSemaine (DateDep:TDateTime):TDateTime;
var
   iDay    :integer;
begin
     iDay   := DayOfWeek (DateDep);
     if iDay = 1 then
         Result := DateDep - 6
     else
         Result := DateDep - (iDay-2);
end;

procedure TOF_TRAVAIL.BuildEtatPlanning;
begin
     TOBEtatTravaux := TOB.Create ('les etats travaux', nil, -1);
     try
        // Création de l'état "vide"
{        with TOB.Create('etat vide', TOBEtatTravaux, -1) do
        begin
             AddChampSupValeur ('E_CODE', '000');
             AddChampSupValeur ('E_LIBELLE', '');
             AddChampSupValeur ('E_COULEURFOND', ColorToString (RGB(255,255,255)));
             AddChampSupValeur ('E_COULEURFONTE', 'clBlack');
             AddChampSupValeur ('E_NOMFONTE', 'Arial');
             AddChampSupValeur ('E_TAILLEFONTE', 8);
             AddChampSupValeur ('E_STYLEFONTE', '');
             AddChampSupValeur ('E_ICONE', -1);
        end;}

        // Création d'un état "fait"
        with TOB.Create('etat fait', TOBEtatTravaux, -1) do
        begin
             AddChampSupValeur ('E_CODE', 'FAIT');
             AddChampSupValeur ('E_LIBELLE', '');
             AddChampSupValeur ('E_COULEURFOND', ColorToString (RGB(0,220,0)));
             AddChampSupValeur ('E_COULEURFONTE', 'clBlack');
             AddChampSupValeur ('E_NOMFONTE', 'Arial');
             AddChampSupValeur ('E_TAILLEFONTE', 9);
             AddChampSupValeur ('E_STYLEFONTE', 'bold');
             AddChampSupValeur ('E_ICONE', -1);
        end;

        // Création d'un état "activité externe"
        with TOB.Create('etat en cours', TOBEtatTravaux, -1) do
        begin
             AddChampSupValeur ('E_CODE', 'NONFAIT');
             AddChampSupValeur ('E_LIBELLE', '');
             AddChampSupValeur ('E_COULEURFOND', ColorToString (RGB(0,0,220)));
             AddChampSupValeur ('E_COULEURFONTE', 'clBlack');
             AddChampSupValeur ('E_NOMFONTE', 'Arial');
             AddChampSupValeur ('E_TAILLEFONTE', 9);
             AddChampSupValeur ('E_STYLEFONTE', 'bold');
             AddChampSupValeur ('E_ICONE', -1);
        end;

        // Création d'un état "fait"
        with TOB.Create('etat en retard', TOBEtatTravaux, -1) do
        begin
             AddChampSupValeur ('E_CODE', 'RETARD');
             AddChampSupValeur ('E_LIBELLE', '');
             AddChampSupValeur ('E_COULEURFOND', ColorToString (RGB(220,0,0)));
             AddChampSupValeur ('E_COULEURFONTE', 'clBlack');
             AddChampSupValeur ('E_NOMFONTE', 'Arial');
             AddChampSupValeur ('E_TAILLEFONTE', 9);
             AddChampSupValeur ('E_STYLEFONTE', 'bold');
             AddChampSupValeur ('E_ICONE', -1);
        end;
     except
           PgiInfo ('Initialisation états travaux impossible');
           TOBEtatTravaux.ClearDetail;
     end;
end;

procedure TOF_TRAVAIL.OnArgument (strArgument:string);
var
   i             :integer;
   Tab           :TTabSheet;
   varCodeGroupe :variant;
   varLibGroupe  :variant;
begin
     inherited;

     // Le panneau où s'affichera le planning
     PanelTravaux := THPanel (GetControl ('PTRAVAUX'));

     // Boutons de paramétrage de l'affichage planning
//     TToolbarButton97 (GetControl ('BJOUR')).OnClick := OnClickJour;
     TToolbarButton97 (GetControl ('BSEMAINE')).OnClick := OnClickSemaine;
     TToolbarButton97 (GetControl ('BMOIS')).OnClick := OnClickMois;
     TToolbarButton97 (GetControl ('BPREV')).OnClick := OnClickPrev;
     TToolbarButton97 (GetControl ('BNEXT')).OnClick := OnClickNext;
//     TPageControl (GetControl ('PAGES')).OnChange    := OnChangePage;

     // Bouton d'impression, export excel et max visu
     TToolbarButton97 (GetControl ('BIMPRIMER')).OnClick := OnClickImprimer;
     TToolbarButton97 (GetControl ('BEXPORTEXCEL')).OnClick := OnClickExporter;
     TToolbarButton97 (GetControl ('BMAX')).OnClick := OnClickMaximiser;

     // Recherche dossier
     THEdit (GetControl ('EDOSSIER')).OnElipsisClick  := OnClickRechDossier;

     // Uniquement les codes événements de type "TRA"vail
     SetControlProperty ('JEV_CODEEVT', 'Plus', 'JTE_FAMEVT="TRA"');

     // Bouton valider: lancement planning travaux
     TToolbarButton97 (GetControl ('BVALIDER')).OnClick := OnClickValider;

     // Traitement associés aux menus de création
//     TMenuItem (GetControl ('ACT_INTERNE')).OnClick := OnClickActInterne;
  //   TMenuItem (GetControl ('ACT_EXTERNE')).OnClick := OnClickActExterne;
    // TMenuItem (GetControl ('ACT_ABSENCE')).OnClick := OnClickActAbsence;

     // Par défaut sur l'utilisateur en cours, et non pas sur le groupe
//     strUser := ''; //V_PGI.User;
  //   strGroupe  := '';
//     TTabSheet (GetControl ('PUSER')).Caption := V_PGI.User + ' - ' + V_PGI.UserName;

     // Date début et type d'affichage date par défaut
     TypeDate  := 2;
     DateDebut := Trunc (DebutDeSemaine (Date));
     DateFin   := DateDebut + 6.9999;

     // Le planning travaux n'existe pas encore
     TOBRessTravaux  := nil;
     TOBEtatTravaux  := nil;
     TOBItemTravaux  := nil;
     TravauxPlanning := nil;

     // Création des état du THPlanning: vide, absence, activité interne et activité externe
     BuildEtatPlanning;

     // Affichage des travaux sur les dossiers visibles par l'utilisateur connecté
     AfficheTravaux (DateDebut);
end;

procedure TOF_TRAVAIL.OnClose ;
begin
     inherited ;

     TOBRessTravaux.Free;
     TOBItemTravaux.Free;
     TOBEtatTravaux.Free;
     TravauxPlanning.Free;
end;

procedure TOF_TRAVAIL.AfficheTravaux (DatePosition:TDateTime);
var
   TOBTravaux          :TOB;
   TOBFille            :TOB;
   TOBLigneRessource   :TOB;
//   TOBCodeEvt          :TOB;
   strRequeteEv        :string;
   strClePlanning      :string;
   strCurClePlanning   :string;
   i                   :integer;
   DateTravauxDeb      :TDateTime;
   DateTravauxFin      :TDateTime;
   strUser             :string;       // Utilisateur à traiter
   strTypeEvt          :string;
   strNoDossier        :string;
   cbsFait             :TCheckBoxState;
   cbsRetard           :TCheckBoxState;
begin
     // Paramètres de filtre
     strUser      := GetControlText ('JEV_USER1');
     strTypeEvt   := THMultiValComboBox (GetControl ('JEV_CODEEVT')).Value;
     strNoDossier := GetControlText ('JEV_NODOSSIER');
     cbsFait      := GetCheckBoxState ('CBFAIT');
     cbsRetard    := GetCheckBoxState ('CBRETARD');

     // Requête sur les événements pour le (ou les) utilisateurs demandés
     strRequeteEv := 'SELECT JEV_NODOSSIER,DOS_LIBELLE AS LIBDOSS,JEV_CODEEVT,JEV_NOEVT,JEV_DATEFIN,JEV_EVTLIBELLE';
     strRequeteEv := strRequeteEv + ' FROM JUEVENEMENT JOIN DOSSIER ON JEV_NODOSSIER=DOS_NODOSSIER';
     strRequeteEv := strRequeteEv + ' WHERE JEV_FAMEVT="TRA" AND JEV_DOMAINEACT<>"JUR"';
     strRequeteEv := strRequeteEv + ' AND JEV_DATEFIN>="' + USDATETIME (DateDebut) + '"';
     strRequeteEv := strRequeteEv + ' AND JEV_DATEFIN<="' + USDATETIME (DateFin) + '"';
     if strUser <> '' then
        strRequeteEv := strRequeteEv + ' AND JEV_USER1="' + strUser + '"';
     if strTypeEvt <> '' then
     begin
          strTypeEvt   := StringReplace (strTypeEvt, ';', '","', [rfReplaceAll]);
          strTypeEvt   := Copy (strTypeEvt, 1, Length (strTypeEvt)-2);
          strRequeteEv := strRequeteEv + ' AND JEV_CODEEVT IN ("' + strTypeEvt + ')';
     end;
     if cbsFait = cbChecked then
         strRequeteEv := strRequeteEv + ' AND JEV_FAIT="X"'
     else
         if cbsFait = cbUnchecked then
            strRequeteEv := strRequeteEv + ' AND JEV_FAIT="-"';
     if cbsRetard = cbChecked then
         strRequeteEv := strRequeteEv + ' AND JEV_DATEFIN<"' + USDATETIME (V_PGI.DateEntree) + '"'
     else
         if cbsRetard = cbUnchecked then
            strRequeteEv := strRequeteEv + ' AND JEV_DATEFIN>="' + USDATETIME (V_PGI.DateEntree) + '"';
     strRequeteEv := strRequeteEv + ' ORDER BY JEV_NODOSSIER,JEV_DATE,JEV_CODEEVT';

     // Construction de la TOB des lignes activité, soit en une seule requête (si pas complétude des ressources) ou en deux
     TOBTravaux  := TOB.Create ('les travaux', nil, -1);
     TOBTravaux.LoadDetailFromSQL (strRequeteEv);

     // Création et affichage de l'objet THPlanning
     TravauxPlanning.Free;
     TravauxPlanning := THPlanning.Create (Ecran);
     if TravauxPlanning = nil then
         PgiInfo ('Initialisation de la fenêtre planning impossible')
     else
     begin
          TravauxPlanning.Parent                          := PanelTravaux;
          TravauxPlanning.Activate                        := FALSE;
          TravauxPlanning.Visible                         := FALSE;
          TravauxPlanning.Align                           := alClient;
          TravauxPlanning.Title                           := 'Travaux';
          TravauxPlanning.Legende                         := FALSE;
          TravauxPlanning.SurBooking                      := TRUE;
          TravauxPlanning.DisplayOptionCreation           := FALSE;
          TravauxPlanning.DisplayOptionCopie              := FALSE;
          TravauxPlanning.DisplayOptionLiaison            := FALSE;
          TravauxPlanning.DisplayOptionLier               := FALSE;
          TravauxPlanning.DisplayOptionSuppressionLiaison := FALSE;

          // Call-back des événements sur les éléments du THplanning
          TravauxPlanning.OnDblClick           := OnDblClickItemTravail;
          TravauxPlanning.OnAvertirApplication := OnEvtTravail;
          TravauxPlanning.OnModifyItem         := OnModifyItemTravail;
          TravauxPlanning.OnDeleteItem         := OnDeleteItemTravail;
          TravauxPlanning.OnMoveItem           := OnMoveItemTravail;
          TravauxPlanning.OnLink               := OnLinkItemTravail;

          // Configuration des états (les etats possible pour les items)
          TravauxPlanning.EtatChampBackGroundColor := 'E_COULEURFOND';
          TravauxPlanning.EtatChampCode            := 'E_CODE';
          TravauxPlanning.EtatChampFontColor       := 'E_COULEURFONTE';
          TravauxPlanning.EtatChampFontName        := 'E_NOMFONTE';
          TravauxPlanning.EtatChampFontSize        := 'E_TAILLEFONTE';
          TravauxPlanning.EtatChampFontStyle       := 'E_STYLEFONTE';
          TravauxPlanning.EtatChampIcone           := 'E_ICONE';
          TravauxPlanning.EtatChampLibelle         := 'E_LIBELLE';

          // Configuration des items (les éléments du corps du planning)
          TravauxPlanning.ChampLineID     := 'I_CODE';
          TravauxPlanning.ChampLibelle    := 'I_LIBELLE';
          TravauxPlanning.ChampdateDebut  := 'I_DATEDEBUT';
          TravauxPlanning.ChampDateFin    := 'I_DATEFIN';
          TravauxPlanning.ChampEtat       := 'I_ETAT';
//          TravauxPlanning.ChampColor      := 'I_COLOR';
          TravauxPlanning.ChampHint       := 'I_HINT';
          TravauxPlanning.ChampIcone      := 'I_ICONE';

          // Configuration des ressources (les entêtes de ligne)
          TravauxPlanning.ResChampColor       := '';
          TravauxPlanning.ResChampFixedColor  := '';
          TravauxPlanning.ResChampID          := 'R_CODE';
          TravauxPlanning.ResChampReadOnly    := '';

          // Couleurs WE/jours fériés et case en sélection
          TravauxPlanning.GestionJoursFeriesActive := FALSE;
          TravauxPlanning.ActiveSaturday           := FALSE;
          TravauxPlanning.ActiveSunday             := FALSE;
          TravauxPlanning.ColorJoursFeries         := clLtGray;
          TravauxPlanning.ColorOfSaturday          := clLtGray;
          TravauxPlanning.ColorOfSunday            := clLtGray;
          TravauxPlanning.ColorSelection           := clNone;

          // Taille des lignes/colonnes
          TravauxPlanning.RowSizeData := 16;
          TravauxPlanning.ColSizeData := 35;
          TravauxPlanning.FrameOn     := TRUE;

          // Taille, alignement des ressources planning (=entête de ligne du planning)
          TravauxPlanning.TokenFieldColFixed  := 'R_LIBDOSS';
          TravauxPlanning.TokenSizeColFixed   := '200';
          TravauxPlanning.TokenAlignColFixed  := 'L';

          // Date périmètres du planning
          TravauxPlanning.IntervalDebut      := DateDebut;
          TravauxPlanning.IntervalFin        := DateFin;
          TravauxPlanning.DateOfStart        := DatePosition;

          // Forme des items (flèches)
          TravauxPlanning.FormeGraphique     := pgRectangle;
          TravauxPlanning.Autorisation       := [patMove, patModify];
          TravauxPlanning.MouseAlready       := TRUE;

          // Entête de colonne
          TravauxPlanning.ActiveLigneGroupeDate := TRUE;
          TravauxPlanning.ActiveLigneDate       := TRUE;
{          if TypeDate = 1 then
          begin
               TravauxPlanning.Interval              := piHeure;
               TravauxPlanning.CumulInterval         := pciJour;
               TravauxPlanning.DateFormat            := 'hh';
               TravauxPlanning.JourneeDebut          := StrToDateTime('00:00');
               TravauxPlanning.JourneeFin            := StrToDateTime('23:59');
          end
          else}
              if TypeDate = 2 then
              begin
                   TravauxPlanning.Interval              := piJour;
                   TravauxPlanning.CumulInterval         := pciSemaine;
                   TravauxPlanning.DateFormat            := 'dd/mm';
              end
              else
              begin
                   TravauxPlanning.Interval              := piJour;
                   TravauxPlanning.CumulInterval         := pciMois;
                   TravauxPlanning.DateFormat            := 'dd';
              end;

          // Génération de la TOB ressources-planning (pas ressources au sens métier)
          strCurClePlanning := '';
//          TOBLigneRessource := nil;

          // Création TOB pour le planning
          TOBRessTravaux.Free;
          TOBRessTravaux := TOB.Create ('ress travaux', nil, -1);
          TOBItemTravaux.Free;
          TOBItemTravaux := TOB.Create ('items travaux', nil, -1);
          for i := 0 to TOBTravaux.Detail.Count-1 do
          begin
               // Agrégat de ligne d'activité à traiter
               TOBFille := TOBTravaux.Detail [i];

               // Valeur "clé" de l'enregistrement: n° dossier
               strClePlanning := '';
               strClePlanning := TOBFille.GetValue ('JEV_NODOSSIER');
               if strClePlanning <> strCurClePlanning then
               begin
                    // Nouvelle clé courante et cumul à 0 car nouvelle ligne
                    strCurClePlanning := strClePlanning;

                    // Création d'une nouvelle ressource planning avec comme code unique le n° dossier
                    TOBLigneRessource := TOB.Create ('une ressource', TOBRessTravaux, -1);
                    with TOBLigneRessource do
                    begin
                         AddChampSupValeur ('R_CODE', strCurClePlanning);
                         AddChampSupValeur ('R_LIBDOSS', strCurClePlanning + ' - ' + Trim (string (TOBFille.GetValue('LIBDOSS'))));
                    end;
               end;

               // Création de l'item associé (si date existante)
               //DateTravauxDeb := TOBFille.GetValue ('JEV_DATE');
               DateTravauxFin := TOBFille.GetValue ('JEV_DATEFIN');
               with TOB.Create ('item travaux', TOBItemTravaux, -1) do
               begin
                    AddChampSupValeur ('I_CODE', strCurClePlanning);
                    if DateTravauxFin > 0 then
                    begin
                         AddChampSupValeur ('I_LIBELLE', TOBFille.GetValue ('JEV_USER1'));
                         AddChampSupValeur ('I_DATEDEBUT', DateTravauxFin);
                         AddChampSupValeur ('I_DATEFIN', DateTravauxFin);
                         if TOBFille.GetValue ('JEV_FAIT') = 'X' then
                             AddChampSupValeur ('I_ETAT', 'FAIT')
                         else
                             if DateTravauxFin < Trunc (V_PGI.DateEntree) then
                                 AddChampSupValeur ('I_ETAT', 'RETARD')
                             else
                                 AddChampSupValeur ('I_ETAT', 'NONFAIT');
                         AddChampSupValeur ('I_NOEVT', IntToStr (TOBFille.GetValue ('JEV_NOEVT')));
                         AddChampSupValeur ('I_CODEEVT', TOBFille.GetValue ('JEV_CODEEVT'));
                    end;
//                    AddChampSupValeur ('I_COLOR', 'clAqua');
                    AddChampSupValeur ('I_HINT', TOBFille.GetValue ('JEV_EVTLIBELLE'));
                    AddChampSupValeur ('I_ICONE', -1);
               end;
          end;

          // Si aucune données, on affiche pas le planning
          if TOBItemTravaux.Detail.Count > 0 then
          begin
               // Affichage planning
               TravauxPlanning.TobRes   := TOBRessTravaux;
               TravauxPlanning.TobEtats := TOBEtatTravaux;
               TravauxPlanning.TobItems := TOBItemTravaux;
               TravauxPlanning.Activate := TRUE;
               TravauxPlanning.Visible  := TRUE;
          end;
     end;
     TOBTravaux.Free;
end;

procedure TOF_TRAVAIL.MaximisePlanning;
var
   ZoomSuivi     :TForm;
begin
     if PlanningValide = TRUE then
     begin
          // Affichage dans une fenêtre maximisée
          ZoomSuivi             := TForm.Create (Ecran);
          ZoomSuivi.Caption     := 'Travaux';
          ZoomSuivi.BorderIcons := [biSystemMenu, biMinimize, biMaximize];
          ZoomSuivi.BorderStyle := bsSizeable;
          ZoomSuivi.WindowState := wsMaximized;
          ZoomSuivi.KeyPreview  := TRUE;
          TravauxPlanning.Parent  := ZoomSuivi;
          PanelTravaux.Caption := 'Affichage des travaux dans une fenêtre agrandie';
          ZoomSuivi.ShowModal;

          // Fin de la fenêtre maximiser: le planning "revient" dans la fenêtre de lancement
          TravauxPlanning.Parent := PanelTravaux;
          ZoomSuivi.Free;
     end;
end;

procedure TOF_TRAVAIL.OnClickValider (Sender:TObject);
begin
     AfficheTravaux (DateDebut);
end;

procedure TOF_TRAVAIL.OnClickImprimer (Sender:TObject);
begin
     if PlanningValide = TRUE then
     begin
          TravauxPlanning.TypeEtat    := 'E';
          TravauxPlanning.NatureEtat  := 'APL';
          TravauxPlanning.CodeEtat    := 'APL';
          TravauxPlanning.Print ('Liste des travaux du ' + FormatDateTime ('dd/mm/yyyy', DateDebut) + ' au ' + FormatDateTime ('dd/mm/yyyy', DateFin));
     end;
end;

procedure TOF_TRAVAIL.OnClickExporter (Sender: TObject);
var
   C  :TSaveDialog;
begin
     if PlanningValide = TRUE then
     begin
          C := TSaveDialog.Create (Ecran);
          try
             C.Options := [ofOverwritePrompt];
             C.Filter  := 'Microsoft Excel (*.xls)|*.XLS';
             if C.Execute then
             begin
                  SourisSablier;
                  try
                     TravauxPlanning.Visible := FALSE;
                     PanelTravaux.Caption := 'Export vers classeur Microsoft Excel "' + C.FileName + '" en cours...';
                     PanelTravaux.Refresh;
                     TravauxPlanning.ExportToExcel (TRUE, C.FileName);
                  finally
                         TravauxPlanning.Visible := TRUE;
                         PanelTravaux.Caption := '';
                         PanelTravaux.Refresh;
                         SourisNormale;
                  end;
             end;
          finally
                 C.Free;
          end;
     end;
end;

procedure TOF_TRAVAIL.OnClickMaximiser (Sender: TObject);
begin
     MaximisePlanning;
end;

procedure TOF_TRAVAIL.OnClickSemaine(Sender: TObject);
begin
     TypeDate := 2;

     DateDebut := Trunc (DebutDeSemaine (TravauxPlanning.GetDateOfCol (TravauxPlanning.Col)));
     DateFin   := DateDebut + 6.9999;
     AfficheTravaux (DateDebut);
end;

procedure TOF_TRAVAIL.OnClickMois(Sender: TObject);
begin
     TypeDate := 3;

     DateDebut := Trunc (DebutDeMois (TravauxPlanning.GetDateOfCol (TravauxPlanning.Col)));
     DateFin   := FinDeMois (DateDebut) + 0.9999;
     AfficheTravaux (DateDebut); //DoPlannings (DateDebut);
end;

procedure TOF_TRAVAIL.OnClickPrev(Sender: TObject);
begin
     case TypeDate of
          2:
               DateDebut := DateDebut - 7;

          3:
               DateDebut := DebutDeMois (PlusMois (DateDebut, -1));
     end;

     AfficheTravaux (DateDebut);
end;

procedure TOF_TRAVAIL.OnClickNext(Sender: TObject);
var
   DateStart     :TDateTime;
begin
     case TypeDate of
          2:
          begin
               DateStart := Trunc (DateFin+1);
               DateFin   := DateStart + 6.9999;
          end;

          3:
          begin
               DateStart := Trunc (DebutDeMois (PlusMois (DateFin,1)));
               DateFin   := FinDeMois (DateStart) + 0.9999;
          end;
     else
         DateStart := DateDebut;
     end;

     AfficheTravaux (DateDebut); //DoPlannings (DateStart);
end;

function TOF_TRAVAIL.AfficheTravailFiche (Item:TOB):boolean; //; strCodeEvt:string):boolean;
var
   NewDateFin       :TDateTime;
   NewUser          :string;
   strArg, strNoEvt :string;
   TOBMaj           :TOB; //, TOBUnDroit        :TOB;
begin
     // Par défaut, aucune modification n'a été effectuée
     Result := FALSE;

     // Selon création ou modification/consultation, les arguments diffèrent
     strArg   := 'ACTION=';
     strNoEvt := '';
     if Item = nil then
     begin
          // Construction argument pour valeur par défaut
          strArg := strArg + 'CREATION';

          if PlanningValide = TRUE then
          begin
               if TravauxPlanning.Col <> -1 then
                   NewDateFin := TravauxPlanning.GetDateOfCol (TravauxPlanning.Col)
               else
                   NewDateFin := Now;
               if TravauxPlanning.Row > 1 then
                   NewUser := TravauxPlanning.TobRes.Detail [TravauxPlanning.Row-2].GetValue ('R_CODE')
               else
                   NewUser := V_PGI.User;
          end
          else
          begin
               NewDateFin := Now;
               NewUser := V_PGI.User;
          end;

          // Lancement en création de la fiche travail
          strArg := strArg + ';JEV_DATEFIN=' + FormatDateTime ('dd/mm/yyyy', NewDateFin) + ';JEV_USER1=' + NewUser;
          strNoEvt := DPLanceFiche_Agenda ('', strArg);
     end
     else
         if Item.GetValue ('I_NOEVT') <> '' then
         begin
              strArg := strArg + 'MODIFICATION';

              // Lancement en création de la fiche travail
              strNoEvt := DPLanceFiche_Agenda (IntToStr (Item.GetValue ('I_NOEVT')), strArg);

              // Màj de l'item dans le THPlanning: relecture enregistrement
              if strNoEvt <> '' then
              begin
                   TOBMaj := TOB.Create ('un evt travail', nil, -1);
                   try
                      TOBMaj.LoadDetailFromSQL ('SELECT JEV_DATE,JEV_USER1,JEV_DATEFIN,JEV_EVTLIBELLE FROM JUEVENEMENT WHERE JEV_NOEVT=' + strNoEvt);
                      if TOBMaj.Detail.Count > 0 then
                      begin
                           NewDateFin := TOBMaj.Detail [0].GetValue ('JEV_DATEFIN');
                           Item.PutValue ('I_CODE', TOBMaj.Detail [0].GetValue ('JEV_NODOSSIER'));
                           Item.PutValue ('I_DATEDEBUT', NewDateFin);
                           Item.PutValue ('I_DATEFIN', NewDateFin);
                           Item.PutValue ('I_HINT', TOBMaj.Detail [0].GetValue ('JEV_EVTLIBELLE'));
                      end;
                   finally
                          TOBMaj.Free;
                   end;
              end;
         end;

     // Si un code événement est renvoyé, cela indique qu'une modification a eu lieu
     Result := strNoEvt <> '';
end;

procedure TOF_TRAVAIL.OnNew;
begin
     if AfficheTravailFiche (nil) = TRUE then
        AfficheTravaux (DateDebut);
end;


function TOF_TRAVAIL.UpdateTravail (Item:TOB):boolean; //FromItem, ToItem:TOB):boolean;
var
   dtNewDebut, dtNewFin    :TDateTime;
   strUpdate               :string;
   OldSep                  :char;
begin
     Result := FALSE;

     // Nouvelles dates de l'activité
     dtNewDebut := Item.GetValue ('I_DATEDEBUT');
     dtNewFin   := Item.GetValue ('I_DATEFIN');

     // En mode semaine ou mois, il faut compléter les nouvelles dates par l'heure déjà définie (non calculé par THPlanning)
     dtNewDebut := Int (dtNewDebut) + Item.GetValue ('I_HEUREDEBUT');
     dtNewFin   := Int (dtNewFin) + Item.GetValue ('I_HEUREFIN');

     // Mise à jour enregistrement avec les nouvelles dates
     try
        OldSep := DecimalSeparator;
        DecimalSeparator := '.';
        strUpdate := 'UPDATE JUEVENEMENT SET JEV_DATE=' + FormatFloat ('0.###############', dtNewDebut) + ', ';
        strUpdate := strUpdate + 'JEV_DATEFIN=' + FormatFloat ('0.###############', dtNewFin) + ', ';
        strUpdate := strUpdate + 'JEV_ALERTEDATE=' + FormatFloat ('0.###############', dtNewDebut) + '-(JEV_DATE-JEV_ALERTEDATE) ';
        strUpdate := strUpdate + 'WHERE JEV_NOEVT=' + Item.GetValue ('I_NOEVT');
        ExecuteSQL (strUpdate);
        DecimalSeparator := OldSep;
        Result := TRUE;
     except
           PgiInfo ('Impossible de mettre à jour le travail');
           DecimalSeparator := OldSep;
     end;
end;

function TOF_TRAVAIL.DeleteTravail (Item: TOB):boolean;
var
   strMessage :string;
begin
     Result := FALSE;

     // Construction du message explicite
     strMessage := 'Confirmez-vous la suppression du travail ';
     strMessage := strMessage + ' de ' + Item.GetValue ('I_CODE') + ' ?';

     // Si demande de suppression confirmée, on supprime...
     if PgiAsk (strMessage) = mrYes then
        try
           ExecuteSQL ('DELETE JUEVENEMENT WHERE JEV_NOEVT=' + IntToStr (Item.GetValue ('I_NOEVT')));
           Result := TRUE;
        except
              PgiInfo ('Impossible de supprimer cet élément');
        end;
end;

{procedure TOF_TRAVAIL.EnterCell (iRow:integer; Item:TOB);
var
   strDroit  :string;
begin
     // Droit sur l'utilisateur associé à la ligne du THPlanning
     strDroit := AgendaPlanning.TobRes.Detail [iRow].GetValue ('R_DROIT');

     // Menus création
     TMenuItem (GetControl ('ACT_INTERNE')).Enabled := FALSE;
     TMenuItem (GetControl ('ACT_EXTERNE')).Enabled := FALSE;
     TMenuItem (GetControl ('ACT_ABSENCE')).Enabled := FALSE;
     if (strDroit = '2') or (strDroit = '4') then
         TMenuItem (GetControl ('ACT_ABSENCE')).Enabled := TRUE;
     if (strDroit = '3') or (strDroit = '4') then
     begin
          TMenuItem (GetControl ('ACT_INTERNE')).Enabled := TRUE;
          TMenuItem (GetControl ('ACT_EXTERNE')).Enabled := TRUE;
     end;

     // Menu contextuel sur l'éventuel item en sélection
     if Item <> nil then
     begin
          if AgendaAutorise (Item) = FALSE then
          begin
               AgendaPlanning.DisplayOptionDeplacement  := FALSE;
               AgendaPlanning.DisplayOptionEtirer       := FALSE;
               AgendaPlanning.DisplayOptionReduire      := FALSE;
               AgendaPlanning.DisplayOptionModification := FALSE;
               AgendaPlanning.DisplayOptionSuppression  := FALSE;
          end
          else
          begin
               AgendaPlanning.DisplayOptionDeplacement  := TRUE;
               AgendaPlanning.DisplayOptionEtirer       := TRUE;
               AgendaPlanning.DisplayOptionReduire      := TRUE;
               AgendaPlanning.DisplayOptionModification := TRUE;
               AgendaPlanning.DisplayOptionSuppression  := TRUE;
          end;
     end;
end;}

procedure TOF_TRAVAIL.OnEvtTravail (Sender: TObject; FromItem, ToItem: TOB; Actions: THPlanningAction);
begin
{     case Actions of
          // En déplacement cellule
          paCellEnter:
                      EnterCell (AgendaPlanning.Row-2);

          // Sur clic droit ou gauche, on s'assure de ne pas proposer les traitements non autorisés
          paClickLeft,
          paClickRight:
                       EnterCell (AgendaPlanning.Row-2, FromItem);
     end;}
end;

procedure TOF_TRAVAIL.OnDblClickItemTravail (Sender: TObject);
var
   NewItem   :TOB;
begin
     NewItem := TOB.Create ('item travail', nil, -1);
     NewItem.Dupliquer (TravauxPlanning.GetCurItem, TRUE, TRUE);
     if AfficheTravailFiche (NewItem) = TRUE then
     begin
          TravauxPlanning.DeleteItem (TravauxPlanning.GetCurItem);
          TravauxPlanning.AddItem (NewItem);
     end;
end;

procedure TOF_TRAVAIL.OnDeleteItemTravail (Sender:TObject; Item:TOB; var Cancel:boolean);
begin
     if DeleteTravail (Item) = TRUE then
        Cancel := TRUE;
end;

procedure TOF_TRAVAIL.OnModifyItemTravail (Sender:TObject; Item:TOB; var Cancel:boolean);
begin
     if AfficheTravailFiche (Item) = FALSE then
        Cancel := TRUE;
end;

procedure TOF_TRAVAIL.OnMoveItemTravail (Sender: TObject; Item: TOB; var Cancel: boolean);
begin
     if UpdateTravail (Item) = TRUE then //, Item) = TRUE then
        exit;

     Cancel := TRUE;
end;

procedure DPLance_Travaux ();
begin
     AGLLanceFiche ('YY','YYTRAVAUX','','', '');
end;

procedure TOF_TRAVAIL.OnLinkItemTravail (Sender:TObject; Source,Destination:TOB; Option:THPlanningOptionLink; var Cancel:boolean);
begin
     case Option of
          polExtend:
          begin
               if UpdateTravail (Destination) = FALSE then
                  Cancel := TRUE;
          end;

          polReduce:
          begin
               if UpdateTravail (Destination) = FALSE then
                  Cancel := TRUE;
          end;
     end;
end;

procedure TOF_TRAVAIL.OnClickRechDossier(Sender: TObject);
var
   strCodeDos   :string;
begin
     strCodeDos := AGLLanceFiche ('YY', 'YYDOSSIER_SEL', '', '', '');
     if strCodeDos <> '' then
        SetControlText ('JEV_NODOSSIER', ReadTokenSt (strCodeDos));
end;

function TOF_TRAVAIL.PlanningValide: boolean;
begin
     Result := (TravauxPlanning <> nil) AND (TravauxPlanning.TobItems <> nil) AND (TravauxPlanning.TobItems.Detail.Count > 0);
end;

initialization
              RegisterClasses ([TOF_TRAVAIL]);
end.

