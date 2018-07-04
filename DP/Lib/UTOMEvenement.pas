{***********UNITE*************************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 17/12/2002
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Description .. : TOM fusionnée entre JURI et DP
Mots clefs ... : TOM;Evenement
*****************************************************************}
unit UTOMEvenement;

interface

uses
      HDB, // => effectue un cast TDBCheckBox = TcheckBox en eagl
{$IFNDEF EAGLCLIENT}
      Db,
{$ENDIF}
      UTob, UTom, ComCtrls, StdCtrls, HCtrls, Controls, classes, DPTofAnnuSel,
      dpOutilsAgenda;

const
     QuartH = 1/96;



//////////// DECLARATION //////////////
type
{$IFDEF EAGLCLIENT}
    TField = THVariant;
{$ENDIF}

   TOM_Evenement = Class (TOM)
      public
         sGuidPerDos_f : string;
         sNoDossier_f : string;

         procedure OnLoadRecord  ; override ;
         procedure OnNewRecord ; override ;
         procedure OnChangeField (F : TField) ; override ;
         procedure OnUpdateRecord ; override;

{$IFDEF BUREAU}
         procedure OnAfterUpdateRecord; override;
{$ENDIF}

         procedure OnArgument(stArgument: String); override;
         procedure OnClose;                        override;
         procedure OnDeleteRecord;                 override; // $$$ JP 24/09/03: pour savoir quel enreg. on supprime
         procedure OnAfterDeleteRecord;            override; // $$$ JP 21/08/07

         //procedure OnExit_Dossier (Sender : TObject); virtual;
         //procedure OnExit_Operation (Sender : TObject); virtual;
         //procedure OnExit_Personne (Sender : TObject); virtual;
         procedure OnElipsisClick_Personne (Sender : TObject); virtual;
         procedure OnElipsisClick_Operation (Sender : TObject); virtual;
         procedure OnElipsisClick_DossierJur (Sender:TObject); virtual;
         procedure OnElipsisClick_DossierBur (Sender:TObject); virtual;

         procedure OnChangeCodeEvt   (Sender:TObject);
         procedure OnElipsisClick_DOCNOM(Sender : TObject);

         procedure FiltreCombo(sControl_p, sRestriction_p : string);

      private
         sAction_f : string;
         sCodeEvt_f : string;
         sFamEvt_f : string;
         sGuidEvt_f : string;
         // $$$ JP 21/08/07 utilisation de TomLoading et/ou TomUpdating   m_bLoadOrSave  :boolean; // $$$ JP 30/10/06

{$IFDEF BUREAU}
         m_TypePeriode  :TPageControl;
         m_chJours      :array [1..7] of TCheckBox;

         m_rbPerJourTousLes   :TRadioButton;
         m_rbPerJourOuvrables :TRadioButton;

         m_rbPerMois          :TRadioButton;
         m_rbPerMoisN         :TRadioButton;
         m_cbPerMoisNInstance :THValComboBox;
         m_cbPerMoisNJour     :THValComboBox;

         m_rbPerAnnee          :TRadioButton;
         m_rbPerAnneeN         :TRadioButton;
         m_cbPerAnneeMois      :THValComboBox;
         m_cbPerAnneeNInstance :THValComboBox;
         m_cbPerAnneeNJour     :THValComboBox;
         m_cbPerAnneeNMois     :THValComboBox;

         m_rbNoEnd      :TRadioButton;
         m_rbEndOcc     :TRadioButton;
         m_rbEndDate    :TRadioButton;
         m_eEndOcc      :THNumEdit;
         m_eStartDate   :THEdit;
         m_eEndDate     :THEdit;

         m_bCalcPer     :boolean;

         m_TOBRecur           :TOB;
         // $$$ JP 21/08/07 inutile, que pour DELETE m_TOBExcept          :TOB;
         m_iTypePer           :ActFrequence; // $$$ JP: type de périodicité actuelle lors du LoadRecord
         m_bMustDelExceptions :boolean; // $$$ JP: toute modif' sur dates et propriétés de périodicité entraine la suppression des exceptions lié à l'enreg actuel
         m_dtDuree            :TDateTime; // $$$ JP 04/10/06: pour pouvoir recalculer date fin si date début changée (et l'inverse)

{$ENDIF}

         sCodeUser_f    :string;  // $$$JP 11/08/03 - code utilisateur par défaut
         sExterne_f     :string;  // $$$JP 22/08/03 - externe ou non par défaut
         sAbsence_f     :string;  // $$$JP 22/08/03 - absence ou non par défaut
         sDateEvt_f     :string;  // $$$JP 11/08/03 - date événement de famille ACT(ivité)
         sDateEvtFin_f  :string;  // $$$JP 26/07/04 - date fin d'événément de famille ACT(ivité)
         sUsers_f       :string;  // $$$JP 20/08/03 - code utilisateurs utilisables pour JEV_USER1
         strEvtMaj      :string;  // $$$JP 20/08/03 - n° événement mis à jour en mode agenda
         sLibEvt_f      :string;  // $$$JP 06/05/04 - libellé par défaut de l'événement agenda
         sLieu_f        :string;  // $$$JP 06/05/04 - lieu par défaut de l'événement agenda
         sGuidPer_f     :string;  // $$$JP 06/05/04 - code personne par défaut de l'événement agenda
         sDossier_f     :string;  // $$$JP 06/05/04 - dossier par défaut de l'événement agenda
         sAffaire_f     :string;  // $$$JP 06/05/04 - affaire par défaut de l'événement agenda

         m_bAutoriseModif  :boolean; // $$$JP 22/09/03 - autorisation de modification
         m_bDeletable      :boolean; // $$$ JP 16/08/06
         m_bException      :boolean; // $$$ JP 28/05/07 - TRUE si on affiche/modifie l'exception, FALSE sinon.
         m_strLastHour     :string;

//         procedure OnExit_DateDeb (Sender : TObject);
//         procedure OnExit_DateFin (Sender : TObject);
//         procedure OnExit_HeureDeb (Sender : TObject);
//         procedure OnExit_HeureFin (Sender : TObject);
//         procedure OnExit_DateAlerte (Sender : TObject);
//         procedure OnExit_HeureAlerte (Sender : TObject);

         procedure OnClick_Alerte(Sender : TObject);
         procedure OnClick_BOUTLOOK(Sender: TObject);
         // BM : remplace le script qui ne gère pas tout
			procedure OnClick_BAffiche(Sender: TObject);

         procedure RecupererLibelleNatureEvt;
         // BM : gestion DateTime -> Date & Time
         procedure OnChange_DateTime(Sender : TObject);
         procedure CalculerDuree;
         procedure DateDecompose(dtDateHeure_p : TDateTime; var sDate_p, sHeure_p : string);
         function  DateRecompose( sChampDate_p, sChampHeure_p : String) : string;
         procedure LoadFieldTime(sFieldDate_p, sChampDate_p, sChampHeure_p : string);
         procedure LoadFieldDuree(dtDuree_p, dtDeb_p, dtFin_p : TDateTime);
//         procedure ChangeFieldTime( sChampDate_p, sChampHeure_p : string );

         procedure AutoriserAlerte;
         procedure CalculerAlerte;
         procedure CalculerInterval( var dtDate_p : TDateTime; sTypeInterval_p : string; nInterval_p : integer);
         procedure ControlerAlerte;
         procedure MajDateAlerte;  // $$$JP 12/08/03: màj date alerte

         procedure SetUser  (strUser:string);
         procedure SetPerso (bPerso:boolean);

{$IFDEF BUREAU}
         procedure OnClickRechAnnuaire (Sender:TObject);
         procedure OnClickRechDossier  (Sender:TObject);
         procedure OnClickRechMission  (Sender:TObject);
         procedure OnClickFait         (Sender:TObject);
         procedure OnChangeUser        (Sender:TObject);
         procedure OnClickPers         (Sender:TObject);
         procedure OnEnterHeure        (Sender:TObject);
         procedure OnExitHeure         (Sender:TObject);
         procedure OnClickMatin        (Sender:TObject);
         procedure OnClickAprem        (Sender:TObject);
         procedure OnClickJournee      (Sender:TObject);

         // $$$ JP 25/09/06 - gestion onglet périodicité
         procedure UpdatePlagePeriodicite;
         procedure UpdateHelpPeriodicite;
         procedure SelectTypeJour      (bOuvrable:boolean);
         procedure SelectTypeMois      (bMoisN:boolean);
         procedure SelectTypeAnnee     (bAnneeN:boolean);

         procedure OnClickJourTousLes      (Sender:TObject);
         procedure OnClickJourOuvrables    (Sender:TObject);
         procedure OnChangeTypePeriode     (Sender:TObject);
         procedure OnClickMois             (Sender:TObject);
         procedure OnClickMoisN            (Sender:TObject);
         procedure OnClickAnnee            (Sender:TObject);
         procedure OnClickAnneeN           (Sender:TObject);

         procedure OnClickNoEnd            (Sender:TObject);
         procedure OnClickEndOcc           (Sender:TObject);
         procedure OnClickEndDate          (Sender:TObject);
         procedure OnChangeEndOcc          (Sender:TObject);
         procedure OnChangeEndDate         (Sender:TObject);
         //procedure OnChangeStartDate       (Sender:TObject);

         procedure OnClickUnJourSemaine    (Sender:TObject);
         procedure OnChangeJourInterval    (Sender:TObject);
         procedure OnChangeSemaineInterval (Sender:TObject);
         procedure OnChangeMoisJour        (Sender:TObject);
         procedure OnChangeMoisInterval    (Sender:TObject);
         procedure OnChangeMoisNInstance   (Sender:TObject);
         procedure OnChangeMoisNJour       (Sender:TObject);
         procedure OnChangeMoisNInterval   (Sender:TObject);

         procedure OnChangeAnneeJour       (Sender:TObject);
         procedure OnChangeAnneeMois       (Sender:TObject);
         procedure OnChangeAnneeNInstance  (Sender:TObject);
         procedure OnChangeAnneeNJour      (Sender:TObject);
         procedure OnChangeAnneeNMois      (Sender:TObject);

         procedure OnFormKeyDown       (Sender:TObject; var Key:Word; Shift:TShiftState);
{$ENDIF}
end;

//////////// IMPLEMENTATION //////////////

implementation

uses
{$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} FE_Main, Fiche, dbctrls,
{$ELSE}
   MaineAGL, eFiche,
{$ENDIF}
   sysutils, extctrls, HEnt1, HMsgBox, HRichEdt,

   {$IFDEF VER150}
   Variants, DateUtils,
   {$ENDIF}

   HTB97, DpJurOutils, ParamSoc, rtfCounter, forms, m3fp, hsysmenu, uDossierSelect, galOutil
   , buttons, UJurOutils 
// unités UJurOutils et UFMaquette dans DP\LIB


{$IFDEF BUREAU}
        // $$$JP 06/04/2004: compil de l'alerte agenda seulement dans mode BUREAU
        , AlerteAgenda, utofaffaire_mul, windows
        , yplanning, yressource // $$$ JP 24/04/07: pour planning unifié
{$IFDEF VER150}
        , TntStdCtrls;
{$ELSE}
  ;
{$ENDIF}

{$ELSE}
       ;
{$ENDIF}


const
     // Codes jours (jour ouvrable, lundi...)
     cjJour         = 0;
     cjJourOuvrable = 1;
     cjJourWeekEnd  = 2;
     cjLundi        = 3;
     cjMardi        = 4;
     cjMercredi     = 5;
     cjJeudi        = 6;
     cjVendredi     = 7;
     cjSamedi       = 8;
     cjDimanche     = 9;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnArgument
Description .. :
Paramètres ... : Liste des arguments
*****************************************************************}
procedure TOM_Evenement.OnArgument (stArgument : String ) ;
var
   Critere   :string;
   i         :integer;
begin
     inherited;
     if stArgument = '' then
        exit;

     // $$$JP 11/08/03: Traitement spécifique fiche agenda
     sCodeUser_f    := '';
     sCodeEvt_f     := '';
     sDateEvt_f     := '';
     sDateEvtFin_f  := '';
     sUsers_f       := '';
     sLibEvt_f      := '';
     sLieu_f        := '';
     sGuidPer_f     := '';
     sDossier_f     := '';
     sAffaire_f     := '';

     m_bAutoriseModif := FALSE;
     m_bDeletable     := TRUE;  // $$$ JP 16/08/06;
     m_bException     := FALSE; // $$$ JP 28/05/07
     // $$$JP 06/05/2004: seulement si mode BUREAU
{$IFDEF BUREAU}
     m_bCalcPer       := FALSE; // $$$ JP 02/10/06

     if (Ecran.Name = 'YYAGENDA_FIC') then //or (Ecran.Name = 'YYTRAVAIL_FIC') then
     begin
          strEvtMaj := '';

          // Interdiction modification dans l'alerte
          for i := 0 to Screen.Formcount-1 do
              if Screen.Forms [i].Name = 'FAlerteAgenda' then
                 TFAlerteAgenda (Screen.Forms [i]).DisableModif;

          // Récupère les critères de lancement de la fiche agenda/fiche travail
          Critere := (Trim (ReadTokenSt (stArgument)));
          while (Critere <> '') or (stArgument <> '') do // $$$ JP 21/09/07: non, peut y avoir des tokens vides au milieu, genre ";...;..." ou "...;;..." // (Critere <>'') do
          begin
             // $$$ JP 21/09/07: on regarde de suite si Critere vide
             if Critere <> '' then
             begin
               // Si action=consultation, alors pas d'enregistrement autorisée
               if Critere = 'MODIFAUTORISEE' then
                  m_bAutoriseModif := TRUE

               // $$$ JP 16/08/06: suppression autorisée ou non
               else if Critere = 'NOTDELETABLE' then
                    m_bDeletable := FALSE

               // $$$ JP 28/05/07: mot clé pour indiquer explicitement si on souhaite ouvrir l'exception
               //else if (Copy (Critere, 1, 13) = 'OPENEXCEPTION' then
               //     m_bException := TRUE

               // Code utilisateur par défaut
               else if (Copy (Critere,1,10) = 'JEV_USER1=') then
                  sCodeUser_f := Copy (Critere, 11, Length (Critere)-10)

               // Date agenda par defaut
               else if (Copy (Critere,1,9) = 'JEV_DATE=') then
                  sDateEvt_f := Copy (Critere, 10, Length (Critere)-9)

               // Date échéance par defaut
               else if (Copy (Critere,1,12) = 'JEV_DATEFIN=') then
                  sDateEvtFin_f := Copy (Critere, 13, Length (Critere)-12)

               // Absence par defaut
               else if (Copy (Critere,1,12) = 'JEV_ABSENCE=') then
                  sAbsence_f := Copy (Critere, 13, Length (Critere)-12)

               // Externe par defaut
               else if (Copy (Critere,1,12) = 'JEV_EXTERNE=') then
                  sExterne_f := Copy (Critere, 13, Length (Critere)-12)

               // Code utilisateurs disponibles en saisie sur cet événement
               else if (Copy (Critere,1,6) = 'USERS=') then
                  sUsers_f := Copy (Critere, 7, Length (Critere)-6)

               // Valeurs à renseigner par défaut
               else if (Copy (Critere,1,12) = 'JEV_CODEEVT=') then
                    sCodeEvt_f := Copy (Critere, 13, Length (Critere)-12)
               else if (Copy (Critere,1,15) = 'JEV_EVTLIBELLE=') then
                    sLibEvt_f := Copy (Critere, 16, Length (Critere)-15)
               else if (Copy (Critere,1,9) = 'JEV_LIEU=') then
                    sLieu_f := Copy (Critere, 10, Length (Critere)-9)
               else if (Copy (Critere,1,12) = 'JEV_GUIDPER=') then
                    sGuidPerDos_f := Copy (Critere, 13, Length (Critere)-12)
               else if (Copy (Critere,1,14) = 'JEV_NODOSSIER=') then
                    sDossier_f := Copy (Critere, 15, Length (Critere)-14)
               else if (Copy (Critere,1,12) = 'JEV_AFFAIRE=') then
                    sAffaire_f := Copy (Critere, 13, Length (Critere)-12);
             end;

             // Paramètre suivant
             Critere := (Trim (ReadTokenSt (stArgument)));
          end;

          // filtre sur les codes utilisateur possible
          if sUsers_f <> '' then
             THDBValComboBox (GetControl ('JEV_USER1')).Plus := 'AND US_UTILISATEUR IN (' + sUsers_f + ')';

          // Sélection des décalage d'alerte courts (minutes, heures et jours)
          THValComboBox (GetControl ('CBDECALALERTE')).Plus       := 'AND (SUBSTRING (CO_CODE,1,1)="+" OR SUBSTRING(CO_CODE,1,1)="0")';
          THEdit (GetControl ('EANNUAIRE')).OnElipsisClick        := OnClickRechAnnuaire;
          THEdit (GetControl ('EMISSION')).OnElipsisClick         := OnClickRechMission;
          TToolbarButton97 (GetControl ('BFAIT')).OnClick         := OnClickFait;

          // Traitement associés aux menus de recherche
          THEdit (GetControl ('EDOSSIER')).OnElipsisClick  := OnClickRechDossier;

          // sur changement code evt, il faut récupérer le lieu par défaut
          THDBValComboBox (GetControl ('JEV_CODEEVT')).OnChange := OnChangeCodeEvt;

          // Changement user ou "activité perso" doit provoquer immédiatement une réaction
          THDBValComboBox (GetControl ('JEV_USER1')).OnChange := OnChangeUser;
          TDBCheckBox     (GetControl ('JEV_PERS')).OnClick   := OnClickPers;

          // $$$ JP 25/08/05 - il faut màj le JEV_DATE et JEV_DATEFIN sur chaque modif' d'heure, ou click matin/jour...
          THEdit  (GetControl ('EHEUREDEB')).OnEnter := OnEnterHeure;
          THEdit  (GetControl ('EHEUREFIN')).OnEnter := OnEnterHeure;
          THEdit  (GetControl ('EHEUREDEB')).OnExit  := OnExitHeure;
          THEdit  (GetControl ('EHEUREFIN')).OnExit  := OnExitHeure;
          TBitBtn (GetControl ('BMATIN')).OnClick    := OnClickMatin;
          TBitBtn (GetControl ('BAPREM')).OnClick    := OnClickAprem;
          TBitBtn (GetControl ('BJOURNEE')).OnClick  := OnClickJournee;

          // $$$ JP 25/09/06 - gestion onglet périodicité
          m_TypePeriode := TPageControl (GetControl ('PTYPEPERIODE'));
          if m_TypePeriode <> nil then
          begin
               // Boutons du type de fréquence (hebdo, mensuelle...)
               m_TypePeriode.OnChange := OnChangeTypePeriode;

               // Plage de fréquence
               m_rbNoEnd    := TRadioButton (GetControl ('RBPERNOEND'));
               m_rbEndOcc   := TRadioButton (GetControl ('RBPERENDOCC'));
               m_rbEndDate  := TRadioButton (GetControl ('RBPERENDDATE'));
               m_eEndOcc    := THNumEdit    (GetControl ('EPERENDOCC'));
               m_eEndDate   := THEdit       (GetControl ('EPERDATEFIN'));
               m_eStartDate := THEdit       (GetControl ('EPERDATEDEB'));
               m_rbNoEnd.OnClick     := OnClickNoEnd;
               m_rbEndOcc.OnClick    := OnClickEndOcc;
               m_rbEndDate.OnClick   := OnClickEndDate;
               m_eEndOcc.OnChange    := OnChangeEndOcc;
               m_eEndDate.OnChange   := OnChangeEndDate;

               // Quotidienne
               THNumEdit    (GetControl ('EPERJOURINTERVAL')).OnChange    := OnChangeJourInterval;
               m_rbPerJourTousLes   := TRadioButton (GetControl ('RBPERJOURTOUSLES'));
               m_rbPerJourOuvrables := TRadioButton (GetControl ('RBPERJOUROUVRABLES'));
               m_rbPerJourTousLes.OnCLick   := OnClickJourTousLes;
               m_rbPerJourOuvrables.OnClick := OnClickJourOuvrables;

               // Hebdomadaire
               THNumEdit    (GetControl ('EPERSEMAINEINTERVAL')).OnChange := OnChangeSemaineInterval;
               m_chJours [1] := TCheckBox (GetControl ('CBLUNDI'));
               m_chJours [2] := TCheckBox (GetControl ('CBMARDI'));
               m_chJours [3] := TCheckBox (GetControl ('CBMERCREDI'));
               m_chJours [4] := TCheckBox (GetControl ('CBJEUDI'));
               m_chJours [5] := TCheckBox (GetControl ('CBVENDREDI'));
               m_chJours [6] := TCheckBox (GetControl ('CBSAMEDI'));
               m_chJours [7] := TCheckBox (GetControl ('CBDIMANCHE'));
               for i := 1 to 7 do
                   m_chJours [i].Onclick := OnClickUnJourSemaine;

               // Choix des périodicités mensuelles
               m_rbPerMois          := TRadioButton (GetControl ('RBPERMOIS'));
               m_rbPerMoisN         := TRadioButton (GetControl ('RBPERMOISN'));
               m_cbPerMoisNInstance := THValComboBox (GetControl ('CBPERMOISNINSTANCE'));
               m_cbPerMoisNJour     := THValComboBox (GetControl ('CBPERMOISNJOUR'));
               m_rbPerMois.OnClick           := OnClickMois;
               m_rbPerMoisN.OnClick          := OnClickMoisN;
               m_cbPerMoisNInstance.OnChange := OnChangeMoisNInstance;
               m_cbPerMoisNJour.OnChange     := OnChangeMoisNJour;
               THNumEdit (GetControl ('EPERMOISJOUR')).OnChange       := OnChangeMoisJour;
               THNumEdit (GetControl ('EPERMOISINTERVAL')).OnChange   := OnChangeMoisInterval;
               THNumEdit (GetControl ('EPERMOISNINTERVAL')).OnChange  := OnChangeMoisNInterval;

               // Choix des périodicités annuelles
               m_rbPerAnnee          := TRadioButton (GetControl ('RBPERANNEE'));
               m_rbPerAnneeN         := TRadioButton (GetControl ('RBPERANNEEN'));
               m_cbPerAnneeMois      := THValComboBox (GetControl ('CBPERANNEEMOIS'));
               m_cbPerAnneeNInstance := THValComboBox (GetControl ('CBPERANNEENINSTANCE'));
               m_cbPerAnneeNJour     := THValComboBox (GetControl ('CBPERANNEENJOUR'));
               m_cbPerAnneeNMois     := THValComboBox (GetControl ('CBPERANNEENMOIS'));
               m_rbPerAnnee.OnClick  := OnClickAnnee;
               m_rbPerAnneeN.OnClick := OnClickAnneeN;
               m_cbPerAnneeMois.OnChange      := OnChangeAnneeMois;
               m_cbPerAnneeNInstance.OnChange := OnChangeAnneeNInstance;
               m_cbPerAnneeNJour.OnChange     := OnChangeAnneeNJour;
               m_cbPerAnneeNMois.OnChange     := OnChangeAnneeNMois;
               THNumEdit (GetControl ('EPERANNEEJOUR')).OnChange       := OnChangeAnneeJour;
          end;

          // $$$ JP 24/08/05 - On doit gérer quelques touches raccourcis (F7, ...)
          Ecran.OnKeyDown := OnFormKeyDown;

          // $$$ JP 03/10/06 - les enregistrements attachés
          m_TOBRecur  := TOB.Create ('JUEVTRECURRENCE', nil, -1);
          // $$$ JP 21/08/07 m_TOBExcept := TOB.Create ('JURECEXCEPTION', nil, -1);
          // On termine tout de suite dans le cas de fiche agenda
          exit;
     end;
{$ENDIF}

   sAction_f := ReadTokenSt(stArgument);
   sNoDossier_f := ReadTokenSt(stArgument);
   sGuidPerDos_f := ReadTokenSt(stArgument);
   sFamEvt_f := ReadTokenSt(stArgument);
   sGuidEvt_f := ReadTokenSt(stArgument);

   FiltreCombo('JEV_CODEEVT', 'JTE_FAMEVT = "' + sFamEvt_f + '"');
   if (GetControl('TJEV_CODEDOS' ) <> nil) then
   begin
        // $$$ JP 11/08/2004 - toujours avoir la fonction activée (mais onexit vide, donc inutile)
        THEdit(GetControl('TJEV_CODEDOS')).OnElipsisClick := OnElipsisClick_DossierJur;
        if GetControl('DOS_LIBELLE' ) <> nil then
        THEdit(GetControl('DOS_LIBELLE')).OnElipsisClick  := OnElipsisClick_DossierBur;
{      if GetControlVisible('TJEV_CODEDOS' ) then
      begin
         THEdit(GetControl('TJEV_CODEDOS')).OnElipsisClick := OnElipsisClick_Dossier;
         THEdit(GetControl('TJEV_CODEDOS')).OnExit         := OnExit_Dossier;
      end
      else if GetControlVisible('DOS_LIBELLE' ) then
      begin
         THEdit(GetControl('DOS_LIBELLE')).OnElipsisClick := OnElipsisClick_Dossier;
         THEdit(GetControl('DOS_LIBELLE')).OnExit         := OnExit_Dossier;
      end}
   end;

   if (GetControl('TJEV_CODEOP') <> nil ) then
   begin
      THEdit(GetControl('TJEV_CODEOP')).OnElipsisClick := OnElipsisClick_Operation;
      //THEdit(GetControl('TJEV_CODEOP')).OnExit := OnExit_Operation;
   end;

   if (Copy(Ecran.Name, 1, 7) = 'YYEVEN_') then
   begin

      if (GetControl('DATE') <> nil ) then
         THEdit(GetControl('DATE')).OnExit := OnChange_DateTime;

      if (GetControl('DATEDEB') <> nil ) then
         THEdit(GetControl('DATEDEB')).OnExit := OnChange_DateTime;

      if (GetControl('DATEFIN') <> nil ) then
         THEdit(GetControl('DATEFIN')).OnExit := OnChange_DateTime;

      if (GetControl('HEUREDEB') <> nil ) then
         THEdit(GetControl('HEUREDEB')).OnExit := OnChange_DateTime;

      if (GetControl('HEUREFIN') <> nil ) then
         THEdit(GetControl('HEUREFIN')).OnExit := OnChange_DateTime;

      if (GetControl('JEV_ALERTE') <> nil) then
      begin
         TCheckBox(GetControl('JEV_ALERTE')).OnClick := OnClick_Alerte;
         if (GetControl('ALERTEDATE') <> nil) then
            THEdit(GetControl('ALERTEDATE')).OnExit := OnChange_DateTime;
         if (GetControl('ALERTEHEURE') <> nil) then
            THEdit(GetControl('ALERTEHEURE')).OnExit := OnChange_DateTime;
      end;
   end;

   if (GetControl('TJEV_GUIDPER') <> nil) then
   begin
      THEdit(GetControl('TJEV_GUIDPER')).OnElipsisClick := OnElipsisClick_Personne;
      //THEdit(GetControl('TJEV_GUIDPER')).OnExit := OnExit_Personne;
   end;


   if (GetControl('BOUTLOOK') <> nil) then
      TToolBarButton97(GetControl('BOUTLOOK')).OnClick := OnClick_BOUTLOOK;

   if (GetControl('BAFFICHE') <> nil) then
      TBitBtn(GetControl('BAFFICHE')).OnClick := OnClick_BAffiche;

   if (GetControl('DOCNOM') <> nil) then
	   THEdit(GetControl('DOCNOM')).OnElipsisClick := OnElipsisClick_DOCNOM;

end;

{$IFDEF BUREAU}
procedure TOM_Evenement.OnFormKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
var
   TD :TDateTime;
begin
     case Key of
          VK_F7:
          begin
               if DS.State <> dsBrowse then
               begin
                    TD := Trunc (GetField ('JEV_DATE')) + StrToTime (Trim (GetControlText ('EHEUREFIN')));
                    SetField ('JEV_DATEFIN', TD);
               end;
          end;
     else
         TFFiche (Ecran).FormKeyDown (Sender, Key, Shift);
     end;
end;
{$ENDIF}


{*****************************************************************
Auteur ....... : BM
Date ......... : 29/09/02
Procédure .... : OnExit_DateHeureDebut
Description .. : Sortie des champs JEV_ALERTEDATE, JEV_DATE, JEV_ALERTEHEURE, JEV_HEUREDEB
                 pour vérification de la date et de l'heure d'alerte.
                 ReCalculer l'alerte par défaut en cas d'erreur
Paramètres ... : L'objet
*****************************************************************}
procedure TOM_Evenement.ControlerAlerte;
var
   dDate_l, dAlerteDate_l : TDate;
begin
   if (GetControl( 'JEV_ALERTE') <> nil ) and ( GetCheckBoxState( 'JEV_ALERTE') = cbChecked ) then
   begin
      dAlerteDate_l := GetField( 'JEV_ALERTEDATE' );
      dDate_l := GetField( 'JEV_DATE' );

      if (dAlerteDate_l <> iDate1900) and ( dAlerteDate_l > dDate_l ) then
      begin
         PgiInfo('La date et l''heure d''alerte doivent être préalables à la date et l''heure de début', titreHalley);
         CalculerAlerte;
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 12/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_Evenement.FiltreCombo(sControl_p, sRestriction_p : string);
var
   sPlus_l : string;
begin
   sPlus_l := THValComboBox(GetControl(sControl_p)).Plus;
   if sPlus_l <> '' then
      sPlus_l := sPlus_l + ' AND ';
   sPlus_l := sPlus_l + '(' + sRestriction_p + ')';
   SetControlProperty(sControl_p, 'Plus', sPlus_l);
end;

procedure TOM_Evenement.SetUser (strUser:string);
begin
     if V_PGI.User = strUser then
         SetControlEnabled ('JEV_PERS', ModifAutorisee or m_bAutoriseModif) // $$$ JP 11/04/07 TRUE)
     else
     begin
          if GetField ('JEV_PERS') = 'X' then
             SetField ('JEV_PERS', '-');
          SetControlEnabled ('JEV_PERS', FALSE);
     end;
end;

procedure TOM_Evenement.SetPerso (bPerso:boolean);
begin
     if bPerso = TRUE then
     begin
          SetField ('JEV_GUIDPER', '');
          SetField ('JEV_NODOSSIER', '');
          SetField ('JEV_AFFAIRE', '');
     end;

     SetControlVisible ('LGUIDPER',  not bPerso);
     SetControlVisible ('EANNUAIRE', not bPerso);
     SetControlVisible ('LCODEDOS',  not bPerso);
     SetControlVisible ('EDOSSIER',  not bPerso);
     SetControlVisible ('LCODEAFF',  not bPerso);
     SetControlVisible ('EMISSION',  not bPerso);
     SetControlVisible ('BEFFACEANNUAIRE', not bPerso);
     SetControlVisible ('BEFFACEDOSSIER',  not bPerso);
     SetControlVisible ('BEFFACEMISSION',  not bPerso);
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnChangeField
Description .. :
Paramètres ... : Le champ modifié
*****************************************************************}
procedure TOM_Evenement.OnChangeField(F: TField);
var
   sLib_l         :string;
   sGuidPer_l     :string;
   TOBLibelle     :TOB;
   dtDate         :TDateTime;
   dtHeure        :TDateTime;
   strHeure       :string;
   i              :integer;
begin
     inherited;

     // $$$JP 12/08/03: si fiche agenda, on fait rien
     // $$$JP 06/05/2004: seulement si mode BUREAU
{$IFDEF BUREAU}
     if (Ecran.Name = 'YYAGENDA_FIC') then //or (Ecran.Name = 'YYTRAVAIL_FIC') then
     begin
          if F.FieldName = 'JEV_GUIDPER' then
          begin
               // Recherche libellé de la personne sélectionnée, et dossier associé (si existant)
               TOBLibelle := TOB.Create ('le libelle', nil, -1);
               TOBLibelle.LoadDetailFromSQL ('SELECT ANN_NOM1 FROM ANNUAIRE WHERE ANN_GUIDPER = "' + GetField ('JEV_GUIDPER') + '"');
               if TOBLibelle.Detail.Count > 0 then
                   SetControlText ('EANNUAIRE', TOBLibelle.Detail [0].GetValue ('ANN_NOM1'))
               else
                   SetControlText ('EANNUAIRE', '<aucune>');
               TOBLibelle.Free;
          end
          else if F.FieldName = 'JEV_AFFAIRE' then
          begin
               // Recherche libellé de la mission sélectionnée
               TOBLibelle := TOB.Create ('le libelle', nil, -1);
               TOBLibelle.LoadDetailFromSQL('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="' + GetField ('JEV_AFFAIRE') + '"');
               if TOBLibelle.Detail.Count > 0 then
                   SetControlText ('EMISSION', TOBLibelle.Detail [0].GetValue ('AFF_LIBELLE'))
               else
                   SetControlText ('EMISSION', '<aucune>');
               TOBLibelle.Free;
          end
          else if F.FieldName = 'JEV_DATE' then
          begin
               // $$$ JP 25/08/05 - pour garder l'heure déjà saisie (si le changement provient d'une saisie)
               dtDate   := GetField ('JEV_DATE');
               dtHeure  := Frac (dtDate);
               strHeure := GetControlText ('EHEUREDEB');
               if (dtHeure = 0.0) and (strHeure <> '00:00') and (strHeure <> '  :  ') and (strHeure <> '') then
               begin
                    dtDate := Trunc (dtDate) + StrToTime (strHeure);
                    SetField ('JEV_DATE', dtDate); // Trunc (dtDate) + StrToTime (strHeure);
               end
               else
               begin
                    SetControlText ('EHEUREDEB', FormatDateTime ('hh:nn', dtHeure));
                    // $$$ JP 31/08/2007: toujours màj datefin même si updating
                    if TomLoading = FALSE then //and (TomUpdating = FALSE) then // $$$ JP 21/08/07 m_bLoadOrSave = FALSE then !!: avoir CBP >=7.0.20.35
                    begin
                         dtDate := dtDate+m_dtDuree;
                         SetField ('JEV_DATEFIN', dtDate); //dtDate+m_dtDuree)
                    end
                    else
                    begin
                         SetControlEnabled ('BMATIN', GetControlEnabled ('JEV_CODEEVT'));
                         SetControlEnabled ('BAPREM', GetControlEnabled ('JEV_CODEEVT'));
                    end;
               end;
               m_dtDuree := GetField ('JEV_DATEFIN') - GetField ('JEV_DATE'); // $$$ JP 04/10/06: durée définie, pour recalculer date fin

               // $$$ JP 21/08/07: uniquement si en saisie
               if (TomLoading = FALSE) and (TomUpdating = FALSE) then // $$$ JP 21/08/07 m_bLoadOrSave = FALSE then
                  m_bMustDelExceptions := not m_bException; // $$$ JP 27/10/06: nécessaire, car change l'éventuelle récurrence

               // $$$ JP 26/09/06 - on doit recalculer quelques propriétés des périodicités
               // $$$ JP 21/08/07: tomloading
               if (TomLoading = FALSE) and (TomUpdating = FALSE){m_bLoadOrSave = FALSE)} and (DS.State <> dsBrowse) and (m_bCalcPer = FALSE) then
               begin
                    m_bCalcPer := TRUE;
                    try
                       for i := 1 to 7 do
                           m_chJours [i].Checked := FALSE;
                       SetControlText ('EPERMOISJOUR', '0');
                       SetControlText ('EPERANNEEJOUR', '0');
                       m_cbPerMoisNInstance.ItemIndex  := -1;
                       m_cbPerAnneeNInstance.ItemIndex := -1;
                    finally
                           m_bCalcPer := FALSE;
                    end;
                    OnChangeTypePeriode (nil);
               end;
          end
          else if F.FieldName = 'JEV_DATEFIN' then
          begin
               // $$$ JP 25/08/05 - pour garder l'heure déjà saisie (si le changement provient d'une saisie)
               dtDate   := GetField ('JEV_DATEFIN');
               dtHeure  := Frac (dtDate);
               strHeure := GetControlText ('EHEUREFIN');
               if (dtHeure = 0.0) and (strHeure <> '00:00') and (strHeure <> '  :  ') and (strHeure <> '') then
               begin
                    dtDate := Trunc (dtDate) + StrToTime (strHeure);
                    SetField ('JEV_DATEFIN', dtDate); //Trunc (dtDate) + StrToTime (strHeure))
               end
               else
               begin
                    SetControlText ('EHEUREFIN', FormatDateTime ('hh:nn', dtHeure));
                    // $$$ JP 21/08/07 tomloading
                    // $$$ JP 31/08/07: la date fin doit toujous être recalculée si elle est < à la date début (même si updating)
                    //if (TomLoading = FALSE) and (TomUpdating = FALSE) {m_bLoadOrSave = FALSE)} and ((dtDate-QuartH) < GetField ('JEV_DATE')) then
                    if (TomLoading = FALSE) and  ((dtDate-QuartH) < GetField ('JEV_DATE')) then
                    begin
                         dtDate := GetField ('JEV_DATE') + QuartH;
                         SetField ('JEV_DATEFIN', dtDate); //GetField ('JEV_DATE') + QuartH) // $$$ JP 09/10/06: on reforce pas la date début, mais plutôt la date fin
                    end
                    else if TomUpdating = FALSE then // $$$ JP 31/08/07: inutile d'invalider des boutons si on enregistre: on va sortir de la fiche)
                    begin
                         SetControlEnabled ('BMATIN', GetControlEnabled ('JEV_CODEEVT'));
                         SetControlEnabled ('BAPREM', GetControlEnabled ('JEV_CODEEVT'));
                    end;
               end;

               m_dtDuree := GetField ('JEV_DATEFIN') - GetField ('JEV_DATE'); // $$$ JP 04/10/06: durée définie, pour recalculer date fin

               // $$$ JP 21/08/07: uniquement si en saisie
               if (TomLoading = FALSE) and (TomUpdating = FALSE) then // $$$ JP 21/08/07 m_bLoadOrSave = FALSE then
                  m_bMustDelExceptions := not m_bException; // $$$ JP 27/10/06: nécessaire, car change l'éventuelle récurrence
          end
          else if F.FieldName = 'JEV_NODOSSIER' then
          begin
               if GetField ('JEV_NODOSSIER') <> '' then
               begin
                    // Recherche libellé du dossier sélectionné
                    TOBLibelle := TOB.Create ('le libelle', nil, -1);
                    TOBLibelle.LoadDetailFromSQL('SELECT DOS_LIBELLE FROM DOSSIER WHERE DOS_NODOSSIER="' + GetField ('JEV_NODOSSIER') + '"');
                    if TOBLibelle.Detail.Count > 0 then
                    begin
                         SetControlText ('EDOSSIER', TOBLibelle.Detail [0].GetValue ('DOS_LIBELLE'));
                         SetControlEnabled ('EMISSION', TRUE); // $$$ JP 16/08/06: si dossier existe, mission doit pouvoir être séléctionnable GetControlEnabled ('JEV_GUIDPER'));
                         SetControlEnabled ('BEFFACEMISSION', TRUE); // $$$ JP 16/08/06 GetControlEnabled ('JEV_GUIDPER'));
                    end
                    else
                    begin
                         SetControlText ('EDOSSIER', '<' + GetField ('JEV_NODOSSIER') + '>');
                         SetControlEnabled ('EMISSION', FALSE);
                         SetControlEnabled ('BEFFACEMISSION', FALSE);
                    end;
                    TOBLibelle.Free;
               end
               else
               begin
                    SetControlText ('EDOSSIER', '<aucun>');
                    SetControlEnabled ('EMISSION', FALSE);
                    SetControlEnabled ('BEFFACEMISSION', FALSE);
               end;
          end
          else if F.FieldName = 'JEV_USER1' then
          begin
               // $$$ JP 05/08/04 - pas de case "activité personnelle" si "on" est pas le user concerné par l'activité
               SetUser (GetField ('JEV_USER1'));
          end
          else if F.FieldName = 'JEV_PERS' then
          begin
               // $$$ JP 05/08/04 - si activité perso, on ne doit pas saisir fiche annuaire/code dossier/ code mission
               SetPerso (GetField ('JEV_PERS') = 'X');
          end;

          // On ne doit pas traiter après, car fice spécifique agenda
          exit;
     end;
{$ENDIF}

   if F.FieldName = 'JEV_GUIDEVT' then
      SetControlEnabled( 'BOUTLOOK', GetField('JEV_GUIDEVT') <> '');

   if F.FieldName = 'JEV_CODEEVT' then
   begin
        sCodeEvt_f := GetField('JEV_CODEEVT');

        // redondance dans table juevenement
        RecupererLibelleNatureEvt;
        UpdateCaption(Ecran);
   end;

   if (F.FieldName = 'JEV_DATE') or (F.FieldName = 'JEV_DATEFIN') then
   begin
//      if (DS.State <> dsBrowse) then
//         CalculerDuree;
   end;

   if (F.FieldName = 'JEV_ALERTE' ) then
   begin
      AutoriserAlerte;
   end;
   if (F.FieldName = 'JEV_ALERTEDATE') then
   begin
//      if (DS.State <> dsBrowse) then
//         ControlerAlerte;
   end;

   if ( F.FieldName = 'JEV_CODEDOS' ) and
      ( GetControl('TJEV_CODEDOS') <> Nil) then
   begin
      sLib_l := GetNomDos(GetField('JEV_CODEDOS'));
      SetControlText('TJEV_CODEDOS', sLib_l);
      if sLib_l = '' then
         SetControlText('JEV_CODEDOS', '');
   end;

   if ( F.FieldName = 'JEV_CODEOP' ) and
      ( GetControl('TJEV_CODEOP') <> Nil ) then
   begin
      sLib_l := GetLibOp(GetField('JEV_CODEOP'), GetField('JEV_CODEDOS'));
      SetControlText('TJEV_CODEOP', sLib_l);
      if sLib_l = '' then
         SetControlText('JEV_CODEOP', '');
   end;

   if ( F.FieldName = 'JEV_GUIDPER' ) and
      ( GetControl('TJEV_GUIDPER') <> Nil ) then
   begin
      if GetField('JEV_GUIDPER') = null then
         sGuidPer_l := ''
      else
         sGuidPer_l := GetField('JEV_GUIDPER');
      sLib_l := GetNomCompPer(sGuidPer_l);
      SetControlText('TJEV_GUIDPER', sLib_l);
      if sLib_l = '' then
         SetControlText('JEV_GUIDPER', '');
   end;

end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnLoadRecord
Description .. :
Paramètres ... :
*****************************************************************}
procedure TOM_Evenement.OnLoadRecord;
var
   dtActivite     :TDateTime;
   dtAlerte       :TDateTime;
   dtDelai        :double;
   strCodeDelai   :string;
   strNoDossier   :string;
   bAlerteAbs     :boolean;
   bFait          :boolean;
   bBoutonEnabled :boolean;
{$IFDEF BUREAU}
   strTypePer     :hstring;
   strDays        :hstring;
   strGuidEvt     :string;
   i              :integer;
{$ENDIF}
begin
     inherited;

     // $$$ JP 30/10/06
     // $$$ JP 21/08/07
     //m_bLoadOrSave := TRUE;
     //try
     if (Ecran.Name = 'YYEVEN_DOC') then
     begin
         SetControlText('DOCNOM', GetField('JEV_DOCNOM'));
         SetControlEnabled('BAFFICHE', GetControlText('DOCNOM') <> '');
     end;

     // BM : décomposition du champ datetime en date et heure
     if (Copy(Ecran.Name, 1, 7) = 'YYEVEN_') then
     begin
         LoadFieldTime('JEV_DATE', 'DATEDEB', 'HEUREDEB');
         LoadFieldTime('JEV_DATEFIN', 'DATEFIN', 'HEUREFIN');
         LoadFieldTime('JEV_ALERTEDATE', 'ALERTEDATE', 'ALERTEHEURE');
         LoadFieldDuree(GetField('JEV_DUREE'), GetField('JEV_DATE'), GetField('JEV_DATEFIN'));
     end;
     
     // $$$JP 11/08/03: gestion spécifique pour fiche agenda
     // $$$JP 06/05/2004: seulement si mode BUREAU
{$IFDEF BUREAU}
     if (Ecran.Name = 'YYAGENDA_FIC') then
     begin
          // $$$ JP 21/10/03: si dossier non existant, interdit de modifier la fiche
          strNoDossier := GetField ('JEV_NODOSSIER');

          // $$$ JP 21/08/07: guid de l'événement
          strGuidEvt := GetField ('JEV_GUIDEVT');

          // $$$ JP 28/05/07: dans la fiche, seuls les exceptions "complètes" sont acceptables (pas les exceptions non représentées dans JUEVENEMENT)
          if GetField ('JEV_OCCURENCEEVT') = 'EXC' then
          begin
               m_bException := TRUE;

               // Pas de modification de périodicité sur une exception
               SetControlVisible ('PRECURRENCE', FALSE);
               SetControlVisible ('LEXCEPTION', TRUE);
          end;

          // $$$JP 28/10/2003: manquait parenthèse englobante sur "existesql", on exécutait donc parfois "FicheReadOnly" quand il ne le fallait pas
          if (strNoDossier <> '') and (ExisteSQL ('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_NODOSSIER="' + strNoDossier + '"') = FALSE) then
          begin
               FicheReadOnly (Ecran);
               m_bAutoriseModif := FALSE;
          end;

          // Si activité réalisée, on ne peut plus faire de modification
          bFait := GetField ('JEV_FAIT') = 'X';
          if bFait = TRUE then
          begin
               ModifAutorisee := FALSE;
               m_bAutoriseModif  := FALSE; // $$$ JP 11/04/07: de toute façon, trop tard pour éclaircir ces 2 booléens: modifautorise
               SetControlEnabled ('EHEUREDEB', FALSE);
               SetControlEnabled ('EHEUREFIN', FALSE);
               SetControlEnabled ('BVALIDER', FALSE);
               SetControlEnabled ('BDELETE', FALSE);
               SetControlEnabled ('PGENERAL', FALSE);
               SetControlEnabled ('PNOTE', FALSE);
               SetControlProperty ('BFAIT', 'hint', 'Considérer l''activité comme NON réalisée');
          end
          else
              SetControlProperty ('BFAIT', 'hint', 'Considérer l''activité comme réalisée');
          SetControlVisible ('LFAIT', bFait);

          if GetField ('JEV_ABSENCE') = 'X' then
          begin
               Ecran.Caption := 'Absence : ';

               // Absence, donc personne, dossier et mission invisibles
               // $$$ JP 11/05/04 - désormais visible, mais non accessible
               SetControlEnabled ('EANNUAIRE', FALSE);
               SetControlEnabled ('BEFFACEANNUAIRE', FALSE);

               // Si alerte absences non autorisé, champs correspondant sont invisibles
               bAlerteAbs  := GetParamSocSecur ('SO_AGEALERTABS', True);
               SetControlVisible ('JEV_ALERTE', bAlerteAbs);
               SetControlvisible ('CBDECALALERTE', bAlerteAbs);
               SetControlVisible ('JEV_ALERTEDATE', bAlerteAbs);
               SetControlVisible ('LALERTEDECAL', bAlerteAbs);
               SetControlVisible ('LALERTEDECALAVANT', bAlerteAbs);
          end
          else
          begin
               if GetField ('JEV_EXTERNE') = 'X' then
                   Ecran.Caption := 'Activité externe : '
               else
                   Ecran.Caption := 'Activité interne : ';

               SetControlVisible ('JEV_ALERTE', TRUE);
               SetControlVisible ('CBDECALALERTE', TRUE);
               SetControlVisible ('JEV_ALERTEDATE', TRUE);
               SetControlVisible ('LALERTEDECALAVANT', TRUE);
          end;

          bBoutonEnabled := GetControlEnabled ('JEV_CODEEVT');
          SetControlEnabled ('BMATIN',          bBoutonEnabled);
          SetControlEnabled ('BAPREM',          bBoutonEnabled);
          SetControlEnabled ('BJOURNEE',        bBoutonEnabled);
          SetControlEnabled ('BEFFACEANNUAIRE', bBoutonEnabled);
          SetControlEnabled ('BEFFACEDOSSIER',  bBoutonEnabled);
          SetControlEnabled ('BEFFACEMISSION',  bBoutonEnabled);
          if strGuidEvt <> '' then // $$$ JP 21/08/07 GetField ('JEV_GUIDEVT') <> '' then
              SetControlEnabled ('BFAIT', TRUE) // $$$ JP 11/04/07: s'il y a un identifiant, on doit toujours pouvoir modifier cet état m_bAutoriseModif)
          else
              SetControlEnabled ('BFAIT', FALSE);

          // $$$ JP 16/08/06: si suppression autorisée
          SetControlEnabled ('BDELETE', m_bDeletable and m_bAutoriseModif);

          // Sélection des motifs en fonction du code événement
          if sAbsence_f = 'X' then
              THValComboBox (GetControl ('JEV_CODEEVT')).Plus := 'JTE_FAMEVT="ACT" AND JTE_ABSENCE="X"'
          else
              if sExterne_f = 'X' then
                  THValComboBox (GetControl ('JEV_CODEEVT')).Plus := 'JTE_FAMEVT="ACT" AND JTE_ABSENCE="-" AND JTE_EXTERNE="X"'
              else
                  THValComboBox (GetControl ('JEV_CODEEVT')).Plus := 'JTE_FAMEVT="ACT" AND JTE_ABSENCE="-" AND JTE_EXTERNE="-"';

          // Calcul délai d'alerte
          strCodeDelai := '';
          dtActivite   := GetField ('JEV_DATE');
          dtAlerte     := GetField ('JEV_ALERTEDATE');
          if (GetField ('JEV_ALERTE') = 'X') and (dtAlerte > iDate1900) then
          begin
               dtDelai := dtActivite - dtAlerte;
               if dtDelai <= StrToTime ('00:59') then
                   strCodeDelai := FormatDateTime ('"+"nn', dtDelai)
               else
                   strCodeDelai := Format ('%.3d', [Round (24*dtDelai)]);

               // $$$ JP 11/04/07: activer les controle seulement si autorisé
               SetControlEnabled ('CBDECALALERTE', bBoutonEnabled);
               SetControlEnabled ('LALERTEDECAL', bBoutonEnabled);
               SetControlEnabled ('LALERTEDECALAVANT', bBoutonEnabled);
               SetControlEnabled ('JEV_ALERTEDATE', bBoutonEnabled);
          end
          else
          begin
               if GetControlText ('CBDECALALERTE') = '' then
                  strCodeDelai := '+15';
               SetControlEnabled ('CBDECALALERTE', FALSE);
               SetControlEnabled ('LALERTEDECAL', FALSE);
               SetControlEnabled ('LALERTEDECALAVANT', FALSE);
               SetControlEnabled ('JEV_ALERTEDATE', FALSE);
          end;

          if RechDom ('RTRAPPELS', strCodeDelai, FALSE) <> '' then
              SetControlProperty ('CBDECALALERTE', 'Value', strCodeDelai)
          else
              SetControlProperty ('CBDECALALERTE', 'Value', '');

          // Mémorisation des valeurs de champ par défaut
          if sCodeUser_f = '' then
             sCodeUser_f := GetField ('JEV_USER1');
          if sAbsence_f = '' then
             sAbsence_f := GetField ('JEV_ABSENCE');
          if sExterne_f = '' then
             sExterne_f := GetField ('JEV_EXTERNE');
          if sDateEvt_f = '' then
             sDateEvt_f := GetField ('JEV_DATE');
          if sDateEvtFin_f = '' then
             sDateEvtFin_f := GetField ('JEV_DATEFIN');

          // $$$ JP 03/10/06: activer la périodicité si saisissable
          m_TypePeriode.Enabled := bBoutonEnabled;

          // $$$ JP 04/10/06: durée définie, pour recalculer date fin
          m_dtDuree := GetField ('JEV_DATEFIN') - dtActivite;
          if m_dtDuree < 0 then
          begin
               m_dtDuree  := QuartH;
               dtActivite := dtActivite + QuartH;
               SetField ('JEV_DATEFIN', dtActivite); //dtActivite + QuartH);
          end;

          // $$$ JP 03/10/06: lecture des tables attachées: JUEVTRECURRENCE et JURECEXCEPTION + alimentation des zones
          // $$$ JP 28/05/07: pas de recurrence (et donc pas non plus d'exception) si on est sur une exception
          if m_bException = FALSE then
          begin
            m_TOBRecur.LoadDetailDB ('JUEVTRECURRENCE', '"' + strGuidEvt {// $$$ JP 21/08/07 GetField ('JEV_GUIDEVT')} + '";1', '', nil, FALSE);
            if m_TOBRecur.Detail.Count > 0 then
            begin
               m_bCalcPer := TRUE;
               strTypePer := m_TOBRecur.Detail [0].GetValue ('JEC_RECURRENCETYPE');
               SetControlVisible ('GBPLAGEPERIODICITE', TRUE);

               if strTypePer = 'QUO' then
               begin
                    m_iTypePer := ACF_DAILY;
                    SetControlText ('EPERJOURINTERVAL', IntToStr (m_TOBRecur.Detail [0].GetValue ('JEC_INTERVALLE')));
                    m_TypePeriode.ActivePageIndex := 1;
               end
               else if strTypePer = 'HEB' then
               begin
                    m_iTypePer := ACF_WEEKLY;

                    // Jours définis éligibles à la répétition
                    strDays := Trim (m_TOBRecur.Detail [0].GetValue ('JEC_WEEKDAYS'));
                    if strDays = '' then
                       strDays := '-------';

                    // Intervalle de semaine (si 0, et jours ouvrables, doit apparaitre dans onglet quotidien même si c'est une répétition hebdomadaire, idem outlook)
                    i := m_TOBRecur.Detail [0].GetValue ('JEC_INTERVALLE');
                    if (i = 0) and (strDays = 'XXXXX--') then
                    begin
                         m_rbPerJourOuvrables.Checked := TRUE;
                         m_TypePeriode.ActivePageIndex := 1;
                    end
                    else
                    begin
                         SetControlText ('EPERSEMAINEINTERVAL', IntToStr (i));
                         for i := 1 to 7 do
                             if strDays [i] <> '-' then
                                m_chJours [i].Checked := TRUE;
                         m_TypePeriode.ActivePageIndex := 2;
                    end;
               end
               else if strTypePer = 'MEN' then
               begin
                    m_iTypePer := ACF_MONTHLY;

                    m_rbPerMois.Checked := TRUE;
                    SetControlText ('EPERMOISINTERVAL', IntToStr (m_TOBRecur.Detail [0].GetValue ('JEC_INTERVALLE')));
                    setControlText ('EPERMOISJOUR', IntToStr (m_TOBRecur.Detail [0].GetValue ('JEC_DAYOFMONTH')));

                    m_TypePeriode.ActivePageIndex := 3;
               end
               else if strTypePer = 'MEX' then
               begin
                    m_iTypePer := ACF_MONTHLY;

                    m_rbPerMoisN.Checked := TRUE;
                    SetControlText ('EPERMOISNINTERVAL', IntToStr (m_TOBRecur.Detail [0].GetValue ('JEC_INTERVALLE')));
                    strDays := Trim (m_TOBRecur.Detail [0].GetValue ('JEC_WEEKDAYS'));
                    if (strDays = '') or (strDays = '-------') then
                        m_cbPerMoisNJour.ItemIndex := -1
                    else
                    begin
                         if strDays = 'XXXXXXX' then
                              m_cbPerMoisNJour.ItemIndex := cjJour
                         else if strDays = 'XXXXX--' then
                              m_cbPerMoisNJour.ItemIndex := cjJourOuvrable
                         else if strDays = '-----XX' then
                              m_cbPerMoisNJour.ItemIndex := cjJourWeekEnd
                         else
                              for i := 1 to 7 do
                                  if strDays [i] <> '-' then
                                  begin
                                       m_cbPerMoisNJour.ItemIndex := i + 2;
                                       break;
                                  end;
                    end;
                    m_cbPerMoisNInstance.ItemIndex := m_TOBRecur.Detail [0].GetValue ('JEC_WEEKOFMONTH') - 1;

                    m_TypePeriode.ActivePageIndex := 3;
               end
               else if strTypePer = 'ANN' then
               begin
                    m_iTypePer := ACF_YEARLY;

                    m_rbPerAnnee.Checked := TRUE;
                    SetControlText ('EPERANNEEJOUR', IntToStr (m_TOBRecur.Detail [0].GetValue ('JEC_DAYOFMONTH')));
                    m_cbPerAnneeMois.ItemIndex := m_TOBRecur.Detail [0].GetValue ('JEC_MONTHOFYEAR') - 1;

                    m_TypePeriode.ActivePageIndex := 4;
               end
               else
               begin
                    m_iTypePer := ACF_YEARLY;

                    m_rbPerAnneeN.Checked := TRUE;
                    strDays := Trim (m_TOBRecur.Detail [0].GetValue ('JEC_WEEKDAYS'));
                    if (strDays = '') or (strDays = '-------') then
                        m_cbPerAnneeNJour.ItemIndex := -1
                    else
                    begin
                         if strDays = 'XXXXXXX' then
                              m_cbPerAnneeNJour.ItemIndex := cjJour
                         else if strDays = 'XXXXX--' then
                              m_cbPerAnneeNJour.ItemIndex := cjJourOuvrable
                         else if strDays = '-----XX' then
                              m_cbPerAnneeNJour.ItemIndex := cjJourWeekEnd
                         else
                              for i := 1 to 7 do
                                  if strDays [i] <> '-' then
                                  begin
                                       m_cbPerAnneeNJour.ItemIndex := i + 2;
                                       break;
                                  end;
                    end;
                    m_cbPerAnneeNInstance.ItemIndex := m_TOBRecur.Detail [0].GetValue ('JEC_WEEKOFMONTH') - 1;
                    m_cbPerAnneeNMois.ItemIndex := m_TOBRecur.Detail [0].GetValue ('JEC_MONTHOFYEAR') - 1;

                    m_TypePeriode.ActivePageIndex := 4;
               end;

               // Plage de périodicité
               SetControlText ('EPERDATEDEB', DateToStr (m_TOBRecur.Detail [0].GetValue ('JEC_START')));
               if m_TOBRecur.Detail [0].GetValue ('JEC_NBOCCURRENCE') <> 0 then //-1 then
               begin
                    m_rbEndOcc.Checked := TRUE;
                    SetControlText ('EPERENDOCC',  IntToStr  (m_TOBRecur.Detail [0].GetValue ('JEC_NBOCCURRENCE')));
               end
               else if m_TOBRecur.Detail [0].GetValue ('JEC_END') <> iDate1900 then
               begin
                    m_rbEndDate.Checked := TRUE;
                    SetControlText ('EPERDATEFIN', DateToStr (m_TOBRecur.Detail [0].GetValue ('JEC_END')));
               end
               else
               begin
                    m_rbNoEnd.Checked := TRUE;
               end;
               m_bCalcPer := FALSE;

               // Chargement des exceptions à la récurrence
               // $$$ JP 21/08/07 m_TOBExcept.LoadDetailDBFromSQL ('JURECEXCEPTION', 'SELECT * FROM JURECEXCEPTION WHERE JEX_GUIDEVT="' + GetField ('JEV_GUIDEVT') + '" AND JEX_RECURNUM=1');
            end
            else
              m_iTypePer := ACF_NONE;
          end;

          // Force les calculs de propriétés périodique par défaut et calcul plage de périodicité
          OnChangeTypePeriode (nil);
          m_bMustDelExceptions := FALSE;

          // Titre: user + code événément + libellé
          Ecran.Caption := Ecran.Caption + GetField ('JEV_USER1') + ' - ' + GetField ('JEV_CODEEVT') + ' ' + GetField ('JEV_EVTLIBABREGE');
     end;
{$ENDIF}
     // $$$ JP 30/10/06
     // $$$ JP 21/08/07
     //finally
     //       m_bLoadOrSave := FALSE;
     //end;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnNewRecord
Description .. :
Paramètres ... :
*****************************************************************}
procedure TOM_Evenement.OnNewRecord;
Var
   tsJevNoteEvt_l :HTStrings;
begin
     inherited;

     // $$$JP 11/08/03: gestion spécifique pour fiche agenda
     // $$$JP 06/05/2004: seulement si mode BUREAU
{$IFDEF BUREAU}
     if (Ecran.Name = 'YYAGENDA_FIC') then
     begin
          if sExterne_f = '' then
             sExterne_f := '-';
          if sAbsence_f = '' then
             sAbsence_f := '-';
          if (sDossier_f = '') and (sAbsence_f <> 'X') and (VH_DOSS <> nil) and (VH_DOSS.NoDossier <> '') then
             sDossier_f := VH_DOSS.NoDossier;
          if sDateEvt_f = '' then
          begin
               sDateEvt_f    := FormatDateTime ('dd/mm/yyyy 08:00', Date);
               sDateEvtFin_f := '';
          end;
          if sDateEvtFin_f = '' then
             sDateEvtFin_f := FormatDateTime ('dd/mm/yyyy hh:nn', StrToDateTime (sDateEvt_f) + StrToTime ('00:30'));

          SetField ('JEV_FAMEVT', 'ACT');
          SetField ('JEV_DATE',    StrToDateTime (sDateEvt_f));
          SetField ('JEV_DATEFIN', StrToDateTime (sDateEvtFin_f));
          if sAbsence_f = 'X' then
          begin
               SetField ('JEV_ALERTE', '-');
               SetField ('JEV_ALERTEDATE', iDate1900); // $$$ JP 11/04/07 0);

               // $$$JP 28/10/2003: si création d'une absence, pas de dossier associé
               SetField ('JEV_NODOSSIER', '');
          end
          else
          begin
               SetField ('JEV_ALERTE', 'X');
               SetField ('JEV_ALERTEDATE', StrToDateTime (sDateEvt_f) - QuartH);
          end;

          // Autres valeurs par défaut
          if sCodeEvt_f <> '' then
             SetField ('JEV_CODEEVT', sCodeEvt_f);
          if sLibEvt_f <> '' then
             SetField ('JEV_EVTLIBELLE', sLibEvt_f);
          if sLieu_f <> '' then
             SetField ('JEV_LIEU', sLieu_f);
          if sGuidPerDos_f <> '' then
             SetField ('JEV_GUIDPER', sGuidPerDos_f);
          if (sDossier_f <> '') or (sAbsence_f = 'X') then
             SetField ('JEV_NODOSSIER', sDossier_f);
          if sAffaire_f <> '' then
             SetField ('JEV_AFFAIRE', sAffaire_f);
          SetField ('JEV_USER1', sCodeUser_f);
          SetField ('JEV_EXTERNE', sExterne_f);
          SetField ('JEV_ABSENCE', sAbsence_f);

          // $$$ JP 03/10/06: les liens vers les tables attachées
          m_TOBRecur.ClearDetail;
          // $$$ JP 21/08/07 m_TOBExcept.ClearDetail;

          // $$$ JP 11/04/07: en création, modification autorisée forcément
          m_bAutoriseModif := TRUE;

          exit; // SORTIE
     end;
{$ENDIF}

   // Allocation de mémoire pour le THRichEdit
   tsJevNoteEvt_l := HTStringList.Create;
   StringsToRich(THRichEdit(GetControl('JEV_NOTEEVT')), tsJevNoteEvt_l) ;
   tsJevNoteEvt_l.Free;

   if sNoDossier_f <> '' then
      SetField('JEV_CODEDOS', sNoDossier_f);
   if sGuidPerDos_f <> '' then
      SetField('JEV_GUIDPER', sGuidPerDos_f);

   SetField('JEV_FAMEVT', sFamEvt_f);

   SetField('JEV_DATE', Now); // Date);
   SetField('JEV_DATEFIN', Date);
   SetField('JEV_DUREE', iDate1900);

   if sFamEvt_f <> 'DOC' then
      SetControlText('HEUREDEB', TimeToStr(Time) );
   SetControlText('HEUREFIN', TimeToStr(Time + EncodeTime ( 0, 30, 0, 0)) );

   SetField('JEV_ALERTEDATE', iDate1900 );
   AutoriserAlerte;

   SetControlText('TJEV_GUIDPER', '');
   SetField('JEV_USER1', GetValChamp('JURIDIQUE', 'JUR_RESP', 'JUR_CODEDOS = "' + sNoDossier_f + '"'));
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnUpdateRecord
Description .. :
Paramètres ... :
*****************************************************************}
procedure TOM_Evenement.OnUpdateRecord;
var
{$IFDEF BUREAU}
   DateRdv     :TDateTime;
   DateFinRdv  :TDateTime;
   DateRec     :TDateTime; // $$$ JP 04/10/06: date début récurrence, pour controle si = date début activité
   dDateMinDeb :TDateTime;
   strMessage  :hstring;
   strValue    :string;
{$ENDIF}
   sGuid_l { $$$ JP 31/08/07 ,sFile_l} :string;
begin
     // $$$JP 11/08/03: gestion spécifique pour fiche agenda
     // $$$JP 06/05/2004: seulement si mode BUREAU
{$IFDEF BUREAU}
     if Ecran.Name = 'YYAGENDA_FIC' then
     begin
        // $$$ JP 21/08/07
        //m_bLoadOrSave := TRUE;
        //try
          // $$$ JP EN EAGLCLIENT, le valider conserve que les date, pas les heures
          DateRdv     := Trunc (GetField ('JEV_DATE')) + StrToTime (GetControlText ('EHEUREDEB')); //;    //Trunc (GetField ('JEV_DATE'));
          SetField    ('JEV_DATE', DateRdv);
          DateFinRdv  := Trunc (GetField ('JEV_DATEFIN')) + StrToTime (GetControlText ('EHEUREFIN')); //; //Trunc (GetField ('JEV_DATEFIN'));
          SetField    ('JEV_DATEFIN', DateFinRdv);

          dDateMinDeb := DebutDeMois (PlusMois (V_PGI.DateEntree, -2)); // $$$ JP 10/08/07: un peu plus permissif, 3 mois max au lieu de 2  // -1));

          // Contrôle qu'un code événement est bien renseigné
          if GetField ('JEV_CODEEVT') = '' then
          begin
               // $$$ JP 10/08/07: LastError réinitialisé lors du SetFocus!!! Faire d'abord le SetFocus
               SetFocusControl ('JEV_CODEEVT');
               LastError    := 1;
               LastErrorMsg := 'Veuillez spécifier le type d''activité';
               exit;
          end;

          // Contrôle qu'un objet est bien renseigné
          if GetField ('JEV_EVTLIBELLE') = '' then
          begin
               // $$$ JP 10/08/07: LastError réinitialisé lors du SetFocus!!! Faire d'abord le SetFocus
               SetFocusControl ('JEV_EVTLIBELLE');
               LastError    := 1;
               LastErrorMsg := 'Veuillez spécifier l''objet de l''activité';
               exit;
          end;

          // Pas de rdv personnel si on est pas l'utilisateur concerné par le rdv
          if (GetField ('JEV_PERS') = 'X') and (GetField ('JEV_USER1') <> V_PGI.User) then
          begin
               // $$$ JP 10/08/07: LastError réinitialisé lors du SetFocus!!! Faire d'abord le SetFocus
               SetFocusControl ('JEV_PERS');
               LastError    := 1;
               LastErrorMsg := 'Vous ne pouvez pas créer d''activité personnelle pour le collaborateur: ' + GetField ('JEV_USER1');
               exit;
          end;

          // Contrôle des dates
          // $$$ JP 06/11/06: seulement si la date à changer, ou bien en création
          if ((DS.State = dsInsert) or (Trunc (StrToDateTime (sDateEvt_f)) <> Trunc (DateRdv))) and (Trunc (DateRdv) < dDateMinDeb) then
          begin
               // $$$ JP 10/08/07: LastError réinitialisé lors du SetFocus!!! Faire d'abord le SetFocus
               SetFocusControl ('JEV_DATE');
               LastError    := 1;
               LastErrorMsg := 'Veuillez spécifier une date de début postérieure au ' + FormatDateTime ('dd/mm/yyyy', dDateMinDeb);
               exit;
          end
          else if Trunc (DateRdv) > Trunc (DateFinRdv) then
          begin
               // $$$ JP 10/08/07: LastError réinitialisé lors du SetFocus!!! Faire d'abord le SetFocus
               SetFocusControl ('JEV_DATEFIN');
               LastError    := 1;
               LastErrorMsg := 'L''activité doit se terminer APRES qu''elle a commencé';
               exit;
          end
          else if (Trunc (DateRdv) = Trunc (DateFinRdv)) and (Frac (DateRdv) > Frac (DateFinRdv)) then
          begin
               // $$$ JP 10/08/07: LastError réinitialisé lors du SetFocus!!! Faire d'abord le SetFocus
               SetFocusControl ('EHEUREDEB');
               LastError    := 1;
               LastErrorMsg := 'L''heure de fin doit être supérieure à l''heure de début';
               exit;
          end;

          // Contrôle qu'une activité n'est pas déjà prévue en chevauchement des nouvelles dates
          // $$$ JP 10/10/06: plus possible pour les activités répétables, et pas grave en fait...
          {if ExisteSQL ('SELECT JEV_GUIDEVT FROM JUEVENEMENT WHERE JEV_USER1="' + GetField ('JEV_USER1') + '" AND JEV_FAMEVT="ACT" AND JEV_GUIDEVT <> "' + GetField ('JEV_GUIDEVT') + '" AND JEV_DATE<"' + Ustime (DateFinRdv) + '" AND JEV_DATEFIN>"' + UsTime (DateRdv) + '"') = TRUE then
             if PgiAsk ('Cette activité va chevaucher une ou plusieurs autres activités' + #10 + ' Confirmez-vous ce traitement?') <> mrYes then
             begin
                  LastError    := 1;
                  LastErrorMsg := 'Chevauchement d''activité refusé';
                  SetFocusControl ('JEV_DATE');
                  exit;
             end;}

          // Demande confirmation si exceptions existent et que les modifications faite sur l'activité implique la suppressions de ses exceptions (modif' date, périodicité...)
          if (m_bMustDelExceptions = TRUE) and (Getfield ('JEV_OCCURENCEEVT') = 'REC') and
             (ExisteSQL ('SELECT 1 FROM JURECEXCEPTION WHERE JEX_GUIDEVT="' + GetField ('JEV_GUIDEVT') + '"') = TRUE) then // $$$ JP 21/08/07 and (m_TOBExcept.Detail.Count > 0) then
          begin
               strMessage := 'Ces modifications entraîneront la suppression des exceptions associées à cette activité répétable'#10' ';
               {if m_TOBExcept.Detail.Count > 1 then
                   strMessage := strMessage + 'les ' + IntToStr (m_TOBExcept.Detail.Count) + ' exceptions liées à cette activité seront perdues.'
               else
                   strMessage := strMessage + 'l''exception liée à cette activité sera perdue.';}
               strMessage := strMessage + #10' Confirmez-vous les modifications apportées à cette activité?';
               if PgiAsk (strMessage) <> mrYes then
               begin
                    LastError    := 1;
                    LastErrorMsg := 'Modification annulée';
                    exit;
               end;
          end;

          // $$$ JP 06/06/07: sur une activité d'exception, pas de répétition
          if m_bException = FALSE then
          begin
            if ActFrequence (m_TypePeriode.ActivePageIndex) = ACF_NONE then
              SetField ('JEV_OCCURENCEEVT', 'SIN')
            else
            begin
               // $$$ JP 26/10/06: vérification des éléments de périodicité
               case ActFrequence (m_TypePeriode.ActivePageIndex) of
                    ACF_DAILY:
                    begin
                         strValue := GetControlText ('EPERJOURINTERVAL');
                         if (strValue = '') or (strValue = '0') then
                         begin
                              LastError    := 1;
                              LastErrorMsg := 'Veuillez saisir une fréquence valide (répétition quotidienne)';
                              exit;
                         end;
                    end;

                    ACF_WEEKLY:
                    begin
                         strValue := GetControlText ('EPERSEMAINEINTERVAL');
                         if (strValue = '') or (strValue = '0') then
                         begin
                              LastError    := 1;
                              LastErrorMsg := 'Veuillez saisir une fréquence valide (répétition hebdomadaire)';
                              exit;
                         end;
                    end;

                    ACF_MONTHLY:
                    begin
                         if m_rbPerMois.Checked = TRUE then
                         begin
                              strValue := GetControlText ('EPERMOISINTERVAL');
                              if (strValue = '') or (strValue = '0') then
                              begin
                                   LastError    := 1;
                                   LastErrorMsg := 'Veuillez saisir une fréquence valide (répétition mensuelle)';
                                   exit;
                              end;
                              strValue := GetControlText ('EPERMOISJOUR');
                              if (strValue = '') or (strValue = '0') then
                              begin
                                   LastError    := 1;
                                   LastErrorMsg := 'Veuillez saisir un numéro de jour valide (répétition mensuelle)';
                                   exit;
                              end;
                         end
                         else
                         begin
                              strValue := GetControlText ('EPERMOISNINTERVAL');
                              if (strValue = '') or (strValue = '0') then
                              begin
                                   LastError    := 1;
                                   LastErrorMsg := 'Veuillez saisir une fréquence valide (répétition mensuelle)';
                                   exit;
                              end;
                         end;
                    end;

                    ACF_YEARLY:
                    begin
                         if m_rbPerAnnee.Checked = TRUE then
                         begin
                              strValue := GetControlText ('EPERANNEEJOUR');
                              if (strValue = '') or (strValue = '0') then
                              begin
                                   LastError    := 1;
                                   LastErrorMsg := 'Veuillez saisir un numéro de jour valide (répétition annuelle)';
                                   exit;
                              end;
                         end;
                    end;
               end;

               // $$$ JP 04/10/06: il faut vérifier que la date définie dans JEV est identique à celle de la 1ère occurence dans JEC
               DateRec := StrToDate (m_eStartDate.Text);
               if Trunc (DateRdv) <> Trunc (DateRec) then
               begin
                    if PgiAsk ('La première occurence de la série doit débuter le ' + m_eStartDate.Text +','#10' Confirmer-vous cette date?') <> mrYes then
                    begin
                         LastError := 1;
                         LastErrorMsg := 'Veuillez saisir une date début et une périodicité qui vous conviennent';
                         exit;
                    end
                    else
                    begin
                         m_bCalcPer := TRUE;
                         DateRec    := DateRec + Frac (DateRdv);
                         SetField ('JEV_DATE', DateRec); //DateRec + Frac (DateRdv));
                         m_bCalcPer := FALSE;
                    end;
               end;

               // Activité récurrente
               SetField ('JEV_OCCURENCEEVT', 'REC');
            end;
          end;
        // $$$ JP 21/08/07
        //finally
        //    m_bLoadOrSave := FALSE;
        //end;
     end;
{$ENDIF}

     // Contrôle qu'un code utilisateur est bien renseigné
     if GetField ('JEV_USER1') = '' then
     begin
          // $$$ JP 10/08/07: LastError réinitialisé lors du SetFocus!!! Faire d'abord le SetFocus
          SetFocusControl ('JEV_USER1');
          LastError    := 1;
          LastErrorMsg := 'Veuillez spécifier le collaborateur concerné';
          exit;
     end;

     // Calcul de la nouvelle clé événement si mode création
     if DS.State = dsInsert then
     begin
        SetField('JEV_GUIDEVT', AglGetGuid);
        SetField('JEV_NOEVT', -2);
     end;
     strEvtMaj := GetField ('JEV_GUIDEVT');

     inherited;

end;

// $$$ JP 03/10/06: il faut mettre à jour ou alimenter les tables liées: JUEVTRECURRENCE & JURECEXCEPTION
{$IFDEF BUREAU}
procedure TOM_Evenement.OnAfterUpdateRecord;
var
   strDays    :hstring;
   i          :integer;
   strGuidEvt :string;
   tobYPL     :TOB;
begin
     inherited;

     // $$$ JP 26/04/07: si pas agenda, on fait rien (normalement jamais le cas, mais bon...)
     if Ecran.Name <> 'YYAGENDA_FIC' then
        exit;

     // $$$ JP 24/04/07
     strGuidEvt := GetField ('JEV_GUIDEVT');

     // $$$ JP 24/04/07: màj ou création dans le planning unifié YPLANNING
     tobYPL := TOB.Create ('YPLANNING', nil, -1);
     try
        LoadTobYPL ('YPL_PREFIXE="JEV" AND YPL_GUIDORI="' + strGuidEvt + '"', tobYPL);
        if tobYPL.Detail.Count < 1 then
        begin
             with TOB.Create ('#YPLANNING', tobYPL, -1) do
             begin
                  AddChampSupValeur ('YPL_PREFIXE',   'JEV');
                  AddChampSupValeur ('YPL_GUIDORI',   GetField ('JEV_GUIDEVT'));
                  AddChampSupValeur ('YPL_GUIDYRS',   GetYRS_GUID ('', '', GetField ('JEV_USER1')));
                  AddChampSupValeur ('YPL_DATEDEBUT', GetField ('JEV_DATE'));
                  AddChampSupValeur ('YPL_DATEFIN',   GetField ('JEV_DATEFIN'));
                  AddChampSupValeur ('YPL_TYPEYPL',   'ACT'); // $$$ JP la famille d'événement dans JEV, car JUEVENEMENT ne contient pas que l'agenda
                  AddChampSupValeur ('YPL_STATUTYPL', '');

                  AddChampSupValeur ('YPL_LIBELLE',   GetField ('JEV_EVTLIBELLE'));
                  AddChampSupValeur ('YPL_ABREGE',    GetField ('JEV_EVTLIBABREGE'));
                  AddChampSupValeur ('YPL_PRIVE',     GetField ('JEV_PERS'));
             end;
             createYPL (tobYPL.Detail [0])
        end
        else
        begin
             with tobYPL.Detail [0] do
             begin
                  PutValue ('YPL_PREFIXE',   'JEV');
                  PutValue ('YPL_GUIDORI',   GetField ('JEV_GUIDEVT'));
                  PutValue ('YPL_GUIDYRS',   GetYRS_GUID ('', '', GetField ('JEV_USER1')));
                  PutValue ('YPL_DATEDEBUT', GetField ('JEV_DATE'));
                  PutValue ('YPL_DATEFIN',   GetField ('JEV_DATEFIN'));
                  PutValue ('YPL_TYPEYPL',   'ACT'); // $$$ JP la famille d'événement dans JEV, car JUEVENEMENT ne contient pas que l'agenda
                  PutValue ('YPL_STATUTYPL', '');
                  PutValue ('YPL_LIBELLE',   GetField ('JEV_EVTLIBELLE'));
                  PutValue ('YPL_ABREGE',    GetField ('JEV_EVTLIBABREGE'));
                  PutValue ('YPL_PRIVE',     GetField ('JEV_PERS'));
             end;
             updateYPL (tobYPL.Detail [0]);
        end;
     finally
            tobYPL.Free;
     end;

     // Si nécessite suppression des exceptions
     // $$$ JP 21/08/07
     if m_bMustDelExceptions = TRUE then // and (m_TOBExcept.Detail.Count > 0) then
        AgendaDeleteExceptions (strguidEvt);
     {begin
          for i := 0 to m_TOBExcept.Detail.Count - 1 do
              m_TOBExcept.Detail [i].DeleteDB;
          m_TOBExcept.ClearDetail;
     end;}

     // S'il y a une périodicité, il faut initialiser les champs
     if ActFrequence (m_TypePeriode.ActivePageIndex) <> ACF_NONE then
     begin
          // Initialisation enreg' de récurrence si pas encore existant
          if m_TOBRecur.Detail.Count < 1 then
          begin
               with TOB.Create ('JUEVTRECURRENCE', m_TOBRecur, -1) do
               begin
                    InitValeurs;
                    PutValue ('JEC_GUIDEVT', GetField ('JEV_GUIDEVT'));
                    PutValue ('JEC_RECURNUM', 1);
               end;
          end;

          m_TOBRecur.Detail [0].PutValue ('JEC_START', StrToDateTime (m_eStartDate.Text));
          if m_rbEndOcc.Checked = TRUE then
               m_TOBRecur.Detail [0].PutValue ('JEC_NBOCCURRENCE', StrToInt (m_eEndOcc.Text))
          else
          begin
               m_TOBRecur.Detail [0].PutValue ('JEC_NBOCCURRENCE', 0);
               if m_rbEndDate.Checked = TRUE then
                   m_TOBRecur.Detail [0].PutValue ('JEC_END', StrToDateTime (m_eEndDate.Text))
               else
                   m_TOBRecur.Detail [0].PutValue ('JEC_END', iDate1900);
          end;
     end;

     // Si changement de périodicité
     case ActFrequence (m_TypePeriode.ActivePageIndex) of
          ACF_NONE:
          begin
               // S'il y a des enreg' de périodicité, on les vire
               ExecuteSQL ('DELETE JUEVTRECURRENCE WHERE JEC_GUIDEVT="' + strGuidEvt + '"');
               m_TobRecur.ClearDetail;
               {if m_TOBRecur.Detail.Count > 0 then
               begin
                    for i := 0 to m_TOBRecur.Detail.Count - 1 do
                        m_TOBRecur.Detail [i].DeleteDB;
                    m_TOBRecur.ClearDetail;
                    {if m_TOBExcept.Detail.Count > 0 then
                    begin
                         for i := 0 to m_TOBExcept.Detail.Count - 1 do
                             m_TOBExcept.Detail [i].DeleteDB;
                         m_TOBExcept.ClearDetail;
                    end;
               end;}
          end;

          ACF_DAILY:
          begin
               // Alimentation des champs de la récurrence
               with m_TOBRecur.Detail [0] do
               begin
                    if m_rbPerJourTousLes.Checked = TRUE then
                    begin
                         PutValue ('JEC_RECURRENCETYPE', 'QUO');
                         PutValue ('JEC_INTERVALLE',     StrToInt (GetControlText ('EPERJOURINTERVAL')));
                         PutValue ('JEC_WEEKDAYS',       '');
                    end
                    else
                    begin
                         // $$$ JP 26/10/06: répétition hebdomadaire bien qu'afficher dans l'onglet des répétitions quotidiennes
                         PutValue ('JEC_RECURRENCETYPE', 'HEB');
                         PutValue ('JEC_INTERVALLE',     0);
                         PutValue ('JEC_WEEKDAYS',       'XXXXX--');
                    end;
                    PutValue ('JEC_DAYOFMONTH',  0);
                    PutValue ('JEC_WEEKOFMONTH', 0);
                    PutValue ('JEC_MONTHOFYEAR', 0);
               end;

               // Mise à jour dans la base
               m_TOBRecur.InsertOrUpdateDB;
          end;

          ACF_WEEKLY:
          begin
               // Alimentation des champs de la récurrence
               with m_TOBRecur.Detail [0] do
               begin
                    PutValue ('JEC_RECURRENCETYPE', 'HEB');
                    PutValue ('JEC_INTERVALLE',  StrToInt (GetControlText ('EPERSEMAINEINTERVAL')));
                    strDays := '-------';
                    for i := 1 to 7 do
                        if m_chJours [i].Checked = TRUE then
                           strDays [i] := 'X';
                    PutValue ('JEC_WEEKDAYS',    strDays);
                    PutValue ('JEC_DAYOFMONTH',  0);
                    PutValue ('JEC_WEEKOFMONTH', 0);
                    PutValue ('JEC_MONTHOFYEAR', 0);
               end;

               // Mise à jour dans la base
               m_TOBRecur.InsertOrUpdateDB;
          end;

          ACF_MONTHLY:
          begin
               if m_rbPerMois.Checked = TRUE then
               begin
                    with m_TOBRecur.Detail [0] do
                    begin
                         PutValue ('JEC_RECURRENCETYPE', 'MEN');
                         PutValue ('JEC_INTERVALLE',  StrToInt (GetControlText ('EPERMOISINTERVAL')));
                         PutValue ('JEC_WEEKDAYS',    '');
                         PutValue ('JEC_DAYOFMONTH',  StrToInt (GetControlText ('EPERMOISJOUR')));
                         PutValue ('JEC_WEEKOFMONTH', 0);
                         PutValue ('JEC_MONTHOFYEAR', 0);
                    end;
               end
               else
               begin
                    with m_TOBRecur.Detail [0] do
                    begin
                         PutValue ('JEC_RECURRENCETYPE', 'MEX');
                         PutValue ('JEC_INTERVALLE',  StrToInt (GetControlText ('EPERMOISNINTERVAL')));
                         case m_cbPerMoisNJour.ItemIndex of
                              cjJour:
                                strDays := 'XXXXXXX';
                              cjJourOuvrable:
                                strDays := 'XXXXX--';
                              cjJourWeekEnd:
                                strDays := '-----XX';
                         else
                              strDays := '-------';
                              strDays [m_cbPerMoisNJour.ItemIndex - 2] := 'X';
                         end;
                         PutValue ('JEC_WEEKDAYS',    strDays);
                         PutValue ('JEC_DAYOFMONTH',  0);
                         PutValue ('JEC_WEEKOFMONTH', m_cbPerMoisNInstance.ItemIndex + 1);
                         PutValue ('JEC_MONTHOFYEAR', 0);
                    end;
               end;

               // Mise à jour dans la base
               m_TOBRecur.InsertOrUpdateDB;
          end;

          ACF_YEARLY:
          begin
               if m_rbPerAnnee.Checked = TRUE then
               begin
                    with m_TOBRecur.Detail [0] do
                    begin
                         PutValue ('JEC_RECURRENCETYPE', 'ANN');
                         PutValue ('JEC_INTERVALLE',  0);
                         PutValue ('JEC_WEEKDAYS',    '');
                         PutValue ('JEC_DAYOFMONTH',  StrToInt (GetControlText ('EPERANNEEJOUR')));
                         PutValue ('JEC_WEEKOFMONTH', 0);
                         PutValue ('JEC_MONTHOFYEAR', m_cbPerAnneeMois.ItemIndex + 1);
                    end;
               end
               else
               begin
                    with m_TOBRecur.Detail [0] do
                    begin
                         PutValue ('JEC_RECURRENCETYPE', 'ANX');
                         PutValue ('JEC_INTERVALLE',  0);
                         case m_cbPerAnneeNJour.ItemIndex of
                              cjJour:
                                strDays := 'XXXXXXX';
                              cjJourOuvrable:
                                strDays := 'XXXXX--';
                              cjJourWeekEnd:
                                strDays := '-----XX';
                         else
                              strDays := '-------';
                              strDays [m_cbPerAnneeNJour.ItemIndex - 2] := 'X';
                         end;
                         PutValue ('JEC_WEEKDAYS',    strDays);
                         PutValue ('JEC_WEEKDAYS',    strDays);
                         PutValue ('JEC_DAYOFMONTH',  0);
                         PutValue ('JEC_WEEKOFMONTH', m_cbPerAnneeNInstance.ItemIndex + 1);
                         PutValue ('JEC_MONTHOFYEAR', m_cbPerAnneeNMois.ItemIndex + 1);
                    end;
               end;

               // Mise à jour dans la base
               m_TOBRecur.InsertOrUpdateDB;
          end;
     end;
end;
{$ENDIF}

{*****************************************************************
Auteur ....... : BM
Date ......... : 23/09/02
Procédure .... : OnClick_Alerte
Description .. : Clic sur check box des alertes : invalide la zone et
                 re-initialise la date et l'heure d'alerte si décochée.
                 Active la zone pour permettre la saisie si cochée.
Paramètres ... : L'objet
*****************************************************************}
procedure TOM_Evenement.OnClick_Alerte(Sender : TObject);
begin
   if Not (DS.State in [dsInsert, dsEdit]) then
      exit;

   if (GetCheckBoxState( 'JEV_ALERTE') = cbChecked) then
   begin
      AutoriserAlerte;
      SetField('JEV_ALERTEUSER', GetField('JEV_USER1'));
      CalculerAlerte;
   end
   else
   begin
      SetField('JEV_ALERTEUSER', '');
      SetField('JEV_ALERTEDATE', iDate1900);
      SetControlText('ALERTEDATE', '01/01/1900');
      SetControlText('ALERTEHEURE', '00:00:00');
      AutoriserAlerte;
   end;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnClick_BOUTLOOK
Description .. : Envoi vers Outlook
Paramètres ... : L'objet
*****************************************************************}
procedure TOM_Evenement.OnClick_BOUTLOOK(Sender: TObject);
var
   sMsg_l : string;
begin
   if AjoutOutlook(sFamEvt_f, GetField('JEV_GUIDEVT')) > 0 then
   begin
      if sFamEvt_f = 'REU' then
         sMsg_l := 'dans le calendrier.'
      else if sFamEvt_f = 'TAC' then
         sMsg_l := 'dans le gestionnaire des tâches.';

      PGIInfo( 'Elément ajouté ' + sMsg_l, 'Liaison Outlook' );
   end;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnElipsisClick_Personne
Description .. :
Paramètres ... : L'objet
*****************************************************************}
procedure TOM_Evenement.OnElipsisClick_Personne(Sender: TObject);
var
   sValCode_l : string;
begin
   //sValCode_l := AGLLanceFiche('YY', 'ANNUAIRE_SEL', '', '', '');
   sValCode_l := LancerAnnuSel ('', '', '');

   if sValCode_l <> '' then
   begin
      ModeEdition(DS);
      SetField('JEV_GUIDPER', sValCode_l);
   end;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnElipsisClick_DossierJur
Description .. :
Paramètres ... : L'objet
*****************************************************************}
procedure TOM_Evenement.OnElipsisClick_DossierJur (Sender:TObject);
var
   sValCode_l  :string;
begin
     sValCode_l := AGLLanceFiche('JUR', 'JURIDIQUE_SEL', '', '', '');
     if sValCode_l <> '' then
     begin
          ModeEdition(DS);
          SetField ('JEV_CODEDOS', sValCode_l);

          // pas de SetField sinon ça boucle
          SetControlText('JEV_CODEOP', '');
          SetControlText('TJEV_CODEOP', '');
   end;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX - JP
Date ......... : 11/08/04
Procédure .... : OnElipsisClick_DossierBur
Description .. :
Paramètres ... :
*****************************************************************}
procedure TOM_Evenement.OnElipsisClick_DossierBur (Sender:TObject);
var
   strCodeDos   :string;
begin
     strCodeDos := AGLLanceFiche ('YY', 'YYDOSSIER_SEL', '', '', '');
     if strCodeDos <> '' then
     begin
          ModeEdition (DS);
          SetField ('JEV_NODOSSIER', ReadTokenSt (strCodeDos));
          ReadTokenSt (strCodeDos);
          SetControlText ('DOS_LIBELLE', ReadTokenSt (strCodeDos));

          // pas de SetField sinon ça boucle
          SetControlText('JEV_CODEOP', '');
          SetControlText('TJEV_CODEOP', '');
     end;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : OnElipsisClick_Operation
Description .. :
Paramètres ... : L'objet
*****************************************************************}
procedure TOM_Evenement.OnElipsisClick_Operation(Sender: TObject);
var
   sCleSel_l, sValCode_l : string;
begin
   if GetField('JEV_CODEDOS') <> '' then
      sCleSel_l := 'JOP_CODEDOS=' + GetField('JEV_CODEDOS');

   sValCode_l := AGLLanceFiche('JUR', 'DOSOPER_SEL', sCleSel_l, '', '');

   if sValCode_l <> '' then
   begin
      ModeEdition(DS);
      SetField('JEV_CODEOP', sValCode_l);
   end;

end;


{*****************************************************************
Auteur ....... : BM
Date ......... : 23/09/02
Procédure .... : AutoriserAlerte
Description .. : Invalide si la zone bVisible_p à FALSE
                 Active la zone si bVisible_p à TRUE
Paramètres ... : Etat à prendre en compte
*****************************************************************}

procedure TOM_Evenement.AutoriserAlerte;
var
   bVisible_l : boolean;
begin
   if GetControl('JEV_ALERTE') <> nil then
   begin
      bVisible_l := GetCheckBoxState( 'JEV_ALERTE') = cbChecked;
      SetControlEnabled( 'JEV_ALERTEUSER', bVisible_l);
      SetControlEnabled( 'ALERTEDATE', bVisible_l);
      SetControlEnabled( 'ALERTEHEURE', bVisible_l);
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 23/09/02
Procédure .... : CalculerAlerte
Description .. : Calculer l'alerte et initialise les champs  JEV_ALERTEDATE, JEV_ALERTEHEURE
Paramètres ... :
*****************************************************************}

procedure TOM_Evenement.CalculerAlerte;
var
   dtDateAlerte_l : TDateTime;
   sHeureDeb_l : string;
begin
   dtDateAlerte_l := GetField('JEV_DATE');
   sHeureDeb_l := FormatDateTime( 'hh:nn:ss', GetField('JEV_DATE') );

   CalculerInterval( dtDateAlerte_l,  GetParamSocSecur('SO_JUDELAITYPE', 'JOU'),
            StrToInt( GetParamSocSecur('SO_JUDELAIDEF', -1) ) );

   SetField( 'JEV_ALERTEDATE', dtDateAlerte_l );
   LoadFieldTime('JEV_ALERTEDATE', 'ALERTEDATE', 'ALERTEHEURE' );
end;
{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Fonction ..... : CalculerDuree
Description .. :
Paramètres ... : Le temps entre 2 dates
*****************************************************************}
procedure TOM_Evenement.CalculerDuree;
var
   dtDateHeureDeb_l, dtDateHeureFin_l, dtDuree_l : TDateTime;
   wYear_l, wMonth_l, wDay_l : word;
   sDuree_l : string;
begin
   dtDateHeureDeb_l := GetField('JEV_DATE');
   dtDateHeureFin_l := GetField('JEV_DATEFIN');
   dtDuree_l := dtDateHeureFin_l - dtDateHeureDeb_l;
   if dtDuree_l <> GetField('JEV_DUREE') then
   begin
{      dtDuree_l := PlusDate(dtDuree_l, 1, 'J');
      DecodeDate(dtDuree_l, wYear_l, wMonth_l, wDay_l);
      if wYear_l
      sDuree_l := FloatToStr(wYear_l)FloatToStr(wYear_l)FloatToStr(wYear_l)}
      LoadFieldDuree(dtDuree_l, dtDateHeureDeb_l, dtDateHeureFin_l);
      SetField('JEV_DUREE', dtDuree_l);
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 29/09/02
Procédure .... : CalculerInterval
Description .. : Calculer la date et de l'heure d'alerte par défaut
                 en fonction de la date et de l'heure de début et du
                 délai par défaut.
Paramètres ... : La date de début
                 L'heure de début
                 Le type de délai : ANN année, JOU jour...
                 Le délai
*****************************************************************}

procedure TOM_Evenement.CalculerInterval( var dtDate_p : TDateTime;
                                        sTypeInterval_p : string; nInterval_p : integer);
var
   nReliqJour_l : integer;
begin
   if nInterval_p = 0 then
      exit;

   if (sTypeInterval_p = 'SEC') or
      (sTypeInterval_p = 'MIN') or
      (sTypeInterval_p = 'HEU') then
   begin
      dtDate_p := PlusHeure( dtDate_p, nInterval_p, Copy(sTypeInterval_p,0,1), nReliqJour_l );
   end
   else if (sTypeInterval_p = 'JOU') or
      (sTypeInterval_p = 'MOI') or
      (sTypeInterval_p = 'ANN') then
      dtDate_p := PlusDate( dtDate_p, nInterval_p, Copy(sTypeInterval_p,0,1) );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 21/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_Evenement.LoadFieldTime(sFieldDate_p, sChampDate_p, sChampHeure_p : string);
var
   sDate_l, sHeure_l : string;
begin
   if (GetControl(sChampDate_p) <> nil) and (GetControl(sChampHeure_p) <> nil) then
   begin
      DateDecompose(GetField(sFieldDate_p), sDate_l, sHeure_l);
      SetControlText(sChampDate_p, sDate_l );
      SetControlText(sChampHeure_p, sHeure_l );
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/08/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_Evenement.LoadFieldDuree(dtDuree_p, dtDeb_p, dtFin_p : TDateTime);
var
   sDuree_l, sPluriel_l : string;
   iJour_l : integer;
begin
   iJour_l := DaysBetween(dtDeb_p, dtFin_p);
   if GetControl('DUREE') = nil then exit;

   if iJour_l >= 1 then
   begin
      if iJour_l > 1 then
         sPluriel_l := 's';
      sDuree_l := Format('%d jour%s ', [iJour_l, sPluriel_l]);
      if dtDuree_p - iJour_l > 0 then
         sDuree_l := sDuree_l + 'et ';
   end;

   if dtDuree_p - iJour_l > 0 then
      sDuree_l := sDuree_l + FormatDateTime('h ''h'' nn', dtDuree_p);
   SetControlText('DUREE', sDuree_l);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/08/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Evenement.OnChange_DateTime(Sender : TObject);
var
   sOldDate_l, sNewDate_l : string;
   sChampName_l, sChampDate_l, sChampHeure_l : string;
begin
//   if TomLoading then exit;
   sChampName_l := THEdit(Sender).Name;

   if (sChampName_l = 'DATEDEB') or (sChampName_l = 'HEUREDEB') then
   begin
      sChampDate_l := 'DATEDEB';
      sChampHeure_l := 'HEUREDEB';
      sChampName_l := 'JEV_DATE';
   end
   else if (sChampName_l = 'DATEFIN') or (sChampName_l = 'HEUREFIN') then
   begin
      sChampDate_l := 'DATEFIN';
      sChampHeure_l := 'HEUREFIN';
      sChampName_l := 'JEV_DATEFIN';
   end
   else if (sChampName_l = 'ALERTEDATE') or (sChampName_l = 'ALERTEHEURE') then
   begin
      sChampDate_l := 'ALERTEDATE';
      sChampHeure_l := 'ALERTEHEURE';
      sChampName_l := 'JEV_ALERTEDATE';
   end
   else
      exit;

   if (GetControl(sChampDate_l) <> nil) and (GetControl(sChampHeure_l) <> nil) then
   begin
      sOldDate_l := FormatDateTime('dd/mm/yyyy hh:nn', GetField(sChampName_l));
      sNewDate_l := DateRecompose(sChampDate_l, sChampHeure_l);

      if sNewDate_l <> sOldDate_l then
      begin
         ModeEdition(DS);
         SetField(sChampName_l, StrToDateTime(sNewDate_l));
      end;
   end;

   if (sChampDate_l = 'DATEDEB') or (sChampHeure_l = 'HEUREDEB') then
      CalculerDuree
   else if (sChampDate_l = 'DATEFIN') or (sChampHeure_l = 'HEUREFIN') then
      CalculerDuree
   else if (sChampDate_l = 'ALERTEDATE') or (sChampHeure_l = 'ALERTEHEURE') then
      ControlerAlerte;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 21/08/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Evenement.DateDecompose(dtDateHeure_p : TDateTime; var sDate_p, sHeure_p : string);
begin
   sDate_p := FormatDateTime('ddd dd/mm/yyyy', dtDateHeure_p);
   if sDate_p = '  :  :  ' then
      sDate_p := '01/01/1900 ';
   sHeure_p := FormatDateTime('hh:nn', dtDateHeure_p);
   if sHeure_p = '  :  ' then
      sHeure_p := '00:00';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_Evenement.DateRecompose(sChampDate_p, sChampHeure_p : String) : string;
var
   sDate_l, sHeure_l : string;
begin
   sDate_l := GetControlText(sChampDate_p);
   sHeure_l := GetControlText(sChampHeure_p);
   result := sDate_l + ' ' + sHeure_l;
end;


{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 17/12/02
Procédure .... : RecupererLibelleNatureEvt
Description .. :
Paramètres ... :
*****************************************************************}
procedure TOM_Evenement.RecupererLibelleNatureEvt;
Var
   sRequete_l : string;
   QRYTypeEvt_l : TQuery ;
begin
   sRequete_l := 'SELECT JTE_EVTLIBABREGE, JTE_DOMAINEACT ' +
                 'FROM JUTYPEEVT ' +
                 'WHERE JTE_CODEEVT = "' + sCodeEvt_f + '"';
   QRYTypeEvt_l := OpenSQL( sRequete_l, TRUE );
   if Not QRYTypeEvt_l.EOF then
      SetField('JEV_EVTLIBABREGE', QRYTypeEvt_l.FindField('JTE_EVTLIBABREGE').AsString);
   Ferme(QRYTypeEvt_l) ;
end;

// $$$JP 12/08/03 - pour màj date alerte dès que màj date activité ou heure début ou coche "alerte"
procedure TOM_Evenement.MajDateAlerte;
var
   strDecal   :string;
   strHeure   :string;
   iDecalDate :TDateTime;
begin
     // $$$ JP 11/04/07: ne rien modifier si pas autorisé (en cours de chargement, ...)
     if (ModifAutorisee = FALSE) or (m_bAutoriseModif = FALSE) or (DS.State = dsBrowse) then // $$ JP 16/05/07: pas trouvé mieux pour éviter les problèmes de checkbox.click agl
        exit;

     if GetCheckBoxState( 'JEV_ALERTE') = cbChecked then //if GetControlText ('JEV_ALERTE') = 'X' then
     begin
          strDecal := GetControlText ('CBDECALALERTE');
          if strDecal <> '' then
          begin
               if strDecal [1] = '+' then
                   // En nb de minutes
                   iDecalDate := StrToTime ('00:' + Copy (strDecal, 2, 2))
               else
                   // En nb d'heure
                   iDecalDate := StrToInt (strDecal) / 24;
          end
          else
              iDecalDate := 0;

          // On enlève déjà le décalage, et ensuite on ajoute l'heure définie (sauf si indéfinie
          iDecalDate := Trunc (GetField ('JEV_DATE')) - iDecalDate;

          // $$$ JP 07/09/2004: gestion heure vide
          strHeure :=  Trim (GetControlText ('EHEUREDEB'));
          if strHeure <> ':' then
             iDecalDate := iDecalDate + StrToTime (strHeure);
              //SetField ('JEV_ALERTEDATE', Trunc (GetField ('JEV_DATE')) + StrToTime (strHeure) - iDecalDate)
          //else
          SetField ('JEV_ALERTEDATE', iDecalDate); //Trunc (GetField ('JEV_DATE')) - iDecalDate);

          SetControlEnabled ('CBDECALALERTE',     TRUE);
          SetControlEnabled ('LALERTEDECAL',      TRUE);
          SetControlEnabled ('LALERTEDECALAVANT', TRUE);
          SetControlEnabled ('JEV_ALERTEDATE',    TRUE);
     end
     else
     begin
          iDecalDate := iDate1900;
          SetField ('JEV_ALERTEDATE', iDecalDate); //iDate1900); //0);

          SetControlEnabled ('CBDECALALERTE',     FALSE);
          SetControlEnabled ('LALERTEDECAL',      FALSE);
          SetControlEnabled ('LALERTEDECALAVANT', FALSE);
          SetControlEnabled ('JEV_ALERTEDATE',    FALSE);
     end;
end;

{$IFDEF BUREAU}
procedure TOM_Evenement.OnClickRechAnnuaire(Sender: TObject);
var
   strGuidPer   :string;
   TOBDossier   :TOB;
begin
     //strGuidPer := AGLLanceFiche ('YY', 'ANNUAIRE_SEL', '', '', '');
     strGuidPer := LancerAnnuSel ('', '', '');

     if strGuidPer <> '' then
     begin
          ModeEdition (DS);
          SetField ('JEV_GUIDPER', strGuidPer);

          // Si dossier associé, on met à jour champ jev_nodossier
          TOBDossier := TOB.Create ('le dossier', nil, -1);
          try
             TOBDossier.LoadDetailFromSQL ('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_GUIDPER= "' + strGuidPer + '"');
             if TOBDossier.Detail.Count > 0 then
             begin
                  SetField ('JEV_NODOSSIER', TOBDossier.Detail [0].GetValue ('DOS_NODOSSIER'));
                  SetField ('JEV_AFFAIRE', '');
             end;
          finally
                 TOBDossier.Free;
          end;
     end;
end;

procedure TOM_Evenement.OnClickRechDossier(Sender: TObject);
var
   strCodeDos   :string;
begin
     strCodeDos := AGLLanceFiche ('YY', 'YYDOSSIER_SEL', '', '', '');
     if strCodeDos <> '' then
     begin
          ModeEdition (DS);
          strCodeDos := ReadTokenSt (strCodeDos);
          SetField ('JEV_NODOSSIER', strCodeDos);
     end;
end;

procedure TOM_Evenement.OnClickRechMission(Sender: TObject);
var
   strCodeMission   :string;
   TOBTiers         :TOB;
   strTiers         :string;
begin
     // Lien vers tiers associé au code dossier sélectionné (pas code personne, car peut être non lié)
     TOBTiers := TOB.Create ('le tiers', nil, -1);
     TOBTiers.LoadDetailFromSQL ('SELECT T_TIERS FROM TIERS JOIN ANNUAIRE ON T_TIERS=ANN_TIERS JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="' + GetField ('JEV_NODOSSIER') + '"');
     if TOBTiers.Detail.Count > 0 then
         strTiers := TOBTiers.Detail [0].GetValue ('T_TIERS')
     else
         strTiers := '';
     FreeAndNil (TOBTiers);

     // $$$JP 27/11/03: pour utiliser la tof affaire
     if strTiers <> '' then
     begin
          strCodeMission := AFLanceFiche_AffaireRech ('', 'AFF_TIERS=' + strTiers);
          if strCodeMission <> '' then
          begin
               ModeEdition (DS);
               strCodeMission := ReadTokenSt (strCodeMission);
               SetField ('JEV_AFFAIRE', strCodeMission);
          end;
     end
     else
         PgiInfo ('Aucun tiers identifié pour ce dossier');
end;
{$ENDIF}

procedure TOM_Evenement.OnClose;
var
   i    :integer;
begin
     inherited;

     // $$$JP 06/05/2004: seulement si mode BUREAU
{$IFDEF BUREAU}
     if Ecran.Name = 'YYAGENDA_FIC' then
     begin
          // Valeur de retour = code de l'événement créé, mis à jour ou supprimé
          TFFiche (Ecran).Retour := strEvtMaj;

          // Interdiction modification dans l'alerte
          for i := 0 to Screen.Formcount-1 do
              if Screen.Forms [i].Name = 'FAlerteAgenda' then
                 TFAlerteAgenda (Screen.Forms [i]).EnableModif;

          // $$$ JP 03/10/06
          // $$$ JP 21/08/07 FreeAndNil (m_TOBExcept);
          FreeAndNil (m_TOBRecur);
     end;
{$ENDIF}
     if Copy(Ecran.Name, 1, 7) = 'YYEVEN_' then
     begin
          // Valeur de retour = code de l'événement créé, mis à jour ou supprimé
          TFFiche (Ecran).Retour := strEvtMaj;
     end;
end;

{$IFDEF BUREAU}
procedure TOM_Evenement.OnClickFait (Sender: TObject);
var
   strMessage   :string;
   iGenereGi    :integer;
begin
     if GetField ('JEV_FAIT') = 'X' then
     begin
          strMessage := 'Confirmez-vous que cette activité est NON réalisée?';
          if GetField ('JEV_GENEREGI') = 'X' then
             strMessage := strMessage + #10 + ' (celle-ci a déjà été transférée en GI)';
          if PgiAsk (strMessage) = mrYes then
          begin
               ModeEdition (DS);
               SetField ('JEV_FAIT', '-');
               TToolbarButton97 (GetControl ('BVALIDER')).Click;
          end;
     end
     else
     begin
          strMessage := 'Confirmez-vous que cette activité est réalisée?';
          if PgiAsk (strMessage) = mrYes then
          begin
               iGenereGi := AgendaSendToGi ('"' + GetField ('JEV_GUIDEVT') + '"', FALSE);
               if iGenereGi <> -1 then
               begin
                    ModeEdition (DS);
                    SetField ('JEV_FAIT', 'X');
                    if iGenereGi > 0 then
                       SetField ('JEV_GENEREGI', 'X')
                    else
                        if GetParamSocSecur ('SO_AGEGENEREGI', True) = TRUE then
                            PgiInfo ('Activité réalisée (non intégrée en GI)')
                        else
                            PgiInfo ('Activité réalisée');
                    TToolbarButton97 (GetControl ('BVALIDER')).Click;
               end;
          end;
     end;
end;

procedure TOM_Evenement.OnChangeUser (Sender: TObject);
begin
     SetUser (Trim (GetControlText ('JEV_USER1')));
end;

procedure TOM_Evenement.OnClickPers (Sender: TObject);
begin
     SetPerso (TCheckBox (Sender).Checked);
end;

procedure TOM_Evenement.OnEnterHeure (Sender:TObject);
begin
     m_strLastHour := THEdit (Sender).Text;
end;

procedure TOM_Evenement.OnExitHeure (Sender:TObject);
var
   dtHour  :TDateTime;
begin
     if THEdit (Sender).Text <> m_strLastHour then
     begin
          ModeEdition (DS);
          try
             dtHour := StrToTime (THEdit (Sender).Text);
             if THEdit (Sender).Name = 'EHEUREDEB' then
             begin
                  dtHour := Trunc (GetField ('JEV_DATE')) + dtHour;
                  SetField ('JEV_DATE', dtHour); //Trunc (GetField ('JEV_DATE')) + dtHour);
                  MajDateAlerte;
             end
             else
             begin
                  dtHour := Trunc (GetField ('JEV_DATEFIN')) + dtHour;
                  SetField ('JEV_DATEFIN', dtHour); //Trunc (GetField ('JEV_DATEFIN')) + dtHour);
             end;
          except
                PgiInfo ('Veuillez saisir une heure valide');
                THEdit (Sender).SetFocus;
          end;
     end;
end;

procedure TOM_Evenement.OnClickMatin (Sender:TObject);
var
   TD  :TDateTime;
begin
     ModeEdition (DS);

     TD := Trunc (GetField ('JEV_DATE')) + StrToTime ('08:00');
     SetField ('JEV_DATE',    TD); //Trunc (GetField ('JEV_DATE')) + StrToTime ('08:00'));

     TD := Trunc (GetField ('JEV_DATE')) + StrToTime ('13:00');
     SetField ('JEV_DATEFIN', TD); // $$$ JP 07/11/06: FQ 11146

     MajDateAlerte;
end;

procedure TOM_Evenement.OnClickAprem (Sender:TObject);
var
   TD  :TDateTime;
begin
     ModeEdition (DS);

     TD := Trunc (GetField ('JEV_DATE')) + StrToTime ('13:00');
     SetField ('JEV_DATE',    TD); //Trunc (GetField ('JEV_DATE')) + StrToTime ('13:00'));

     TD := Trunc (GetField ('JEV_DATE')) + StrToTime ('20:00');
     SetField ('JEV_DATEFIN', TD); //Trunc (GetField ('JEV_DATE')) + StrToTime ('20:00')); // $$$ JP 07/11/06: FQ 11146

     MajDateAlerte;
end;

procedure TOM_Evenement.OnClickJournee (Sender:TObject);
var
   TD  :TDateTime;
begin
     ModeEdition (DS);

     TD := Trunc (GetField ('JEV_DATE')) + StrToTime ('08:00');
     SetField ('JEV_DATE',    TD); //Trunc (GetField ('JEV_DATE')) + StrToTime ('08:00'));

     TD := Trunc (GetField ('JEV_DATE')) + StrToTime ('20:00');
     SetField ('JEV_DATEFIN', TD); //Trunc (GetField ('JEV_DATE')) + StrToTime ('20:00')); // $$$ JP 07/11/06: FQ 11146

     MajDateAlerte;
end;

// Reconstruction des propriétés de la plage de périodicité: comme outlook, dans la mesure du possible
procedure TOM_Evenement.UpdatePlagePeriodicite;
var
   iTypePer         :ActFrequence;
   dtStart, dtEnd   :TDateTime;
   i, iOccurence    :integer;
   iInterval, iJour :integer;
   strDaysOfWeek    :string;
begin
     if m_bCalcPer = TRUE then
        exit;

     m_bCalcPer := TRUE;

     // $$$ JP 08/11/06: libellé date du 1er onglet, pour qu'on voit bien que l'activité se répète ou non
     if m_iTypePer = ACF_NONE then
         SetControlCaption ('GBDATE', 'Date')
     else
         SetControlCaption ('GBDATE', 'Date de début de répétition');

     try
        iTypePer   := ActFrequence (m_TypePeriode.ActivePageIndex);
        dtStart    := Trunc (GetField ('JEV_DATE'));
        if m_rbEndDate.Checked = TRUE then
        begin
             // Date fin spécifiée, on calcul à partir de celle-ci et non un nombre d'occurence
             iOccurence := 0;
             try
                dtEnd := StrToDate (m_eEndDate.Text);
             except
                   exit;
             end;
        end
        else
        begin
             // Nb d'occurence spécifiée ou sans fin: on calcul la date de fin théorique
             dtEnd := iDate1900;
             try
                iOccurence := StrToInt (m_eEndOcc.Text);
             except
                   exit;
             end;
        end;

        case iTypePer of
          ACF_DAILY:
          begin
               // Cas particulier: "tous les jours ouvrables" = répétition hebdomadaire et non pas quotidienne
               if m_rbPerJourTousLes.Checked = TRUE then
               begin
                    try
                       iInterval := StrToInt (GetControlText ('EPERJOURINTERVAL'));
                    except
                          exit;
                    end;

                    AgendaGetOccurence (ACF_DAILY, dtStart, dtEnd, iOccurence, iInterval);
               end
               else
                    AgendaGetOccurence (ACF_WEEKLY, dtStart, dtEnd, iOccurence, 1, 0, 'XXXXX--');
          end;

          ACF_WEEKLY:
          begin
               try
                  iInterval := StrToInt (GetControlText ('EPERSEMAINEINTERVAL'));
               except
                     exit;
               end;
               strDaysOfWeek := '-------';
               for i := 1 to 7 do
                   if m_chJours [i].Checked = TRUE then
                      strDaysOfWeek [i] := 'X';
               if strDaysOfWeek = '-------' then
               begin
                    i := DayOfWeek (dtStart);
                    if i = 1 then
                        i := 7
                    else
                        i := i - 1;
                    m_chJours [i].Checked := TRUE;
                    strDaysOfWeek [i] := 'X';
               end;

               AgendaGetOccurence (iTypePer, dtStart, dtEnd, iOccurence, iInterval, 0, strDaysOfWeek);
          end;

          ACF_MONTHLY:
          begin
               // Si zone vide, on prend la valeur par défaut selon la date début
               if m_rbPerMois.Checked = TRUE then
               begin
                    try
                       iInterval := StrToInt (GetControlText ('EPERMOISINTERVAL'))
                    except
                          exit;
                    end;
                    try
                       iJour := StrToInt (GetControlText ('EPERMOISJOUR'));
                    except
                          exit;
                    end;
                    AgendaGetOccurence (iTypePer, dtStart, dtEnd, iOccurence, iInterval, 0, '', iJour);
               end
               else
               begin
                    try
                       iInterval := StrToInt (GetControlText ('EPERMOISNINTERVAL'));
                    except
                          exit;
                    end;
                    case m_cbPerMoisNJour.ItemIndex of
                         cjJour:
                            strDaysOfWeek := 'XXXXXXX';

                         cjJourOuvrable:
                            strDaysOfWeek := 'XXXXX--';

                         cjJourWeekEnd:
                            strDaysOfWeek := '-----XX';
                    else
                        // Jour précis de la semaine
                        strDaysOfWeek := '-------';
                        strDaysOfWeek [m_cbPerMoisNJour.ItemIndex-2] := 'X';
                    end;

                    AgendaGetOccurence (iTypePer, dtStart, dtEnd, iOccurence, iInterval, m_cbPerMoisNInstance.ItemIndex + 1, strDaysOfWeek);
               end;
          end;

          ACF_YEARLY:
          begin
               if m_rbPerAnnee.Checked = TRUE then
               begin
                    try
                       iJour := StrToInt (GetControlText ('EPERANNEEJOUR'));
                    except
                          exit;
                    end;
                    AgendaGetOccurence (iTypePer, dtStart, dtEnd, iOccurence, 0, 0, '', iJour, m_cbPerAnneeMois.ItemIndex + 1);
               end
               else
               begin
                    case m_cbPerAnneeNJour.ItemIndex of
                         cjJour:
                            strDaysOfWeek := 'XXXXXXX';

                         cjJourOuvrable:
                            strDaysOfWeek := 'XXXXX--';

                         cjJourWeekEnd:
                            strDaysOfWeek := '-----XX';
                    else
                        // Jour précis de la semaine
                        strDaysOfWeek := '-------';
                        strDaysOfWeek [m_cbPerAnneeNJour.ItemIndex-2] := 'X';
                    end;

                    AgendaGetOccurence (iTypePer, dtStart, dtEnd, iOccurence, 0, m_cbPerAnneeNInstance.ItemIndex + 1, strDaysOfWeek, 0, m_cbPerAnneeNMois.ItemIndex + 1);
               end;
          end;
        end;

        SetControlText ('EPERENDOCC',  IntToStr (iOccurence));
        SetControlText ('EPERDATEDEB', DateToStr (dtStart));
        SetControlText ('EPERDATEFIN', DateToStr (dtEnd));

        // $$$ JP 21/08/07: uniquement si en saisie
        if (TomLoading = FALSE) and (TomUpdating = FALSE) then // $$$ JP 21/08/07 m_bLoadOrSave = FALSE then
           m_bMustDelExceptions := not m_bException; // Pour indiquer qu'une modification a eu lieu sur la périodicité, donc les exceptions peuvent ne plus correspondre (comme ms outlook)

        // Si en consultation, il faut se mettre en édition (et si pas en cours de chargement)
        if (DS.State = dsBrowse) and (TomLoading = FALSE) and (TomUpdating = FALSE) then // $$$ JP 21/08/07 m_bLoadOrSave = FALSE) then
           ModeEdition (DS);

        // Pas d'alerte possible sur les activités répétables, ni la possibilité de cocher "fait"
        // $$$ JP 11/04/07: SetField à faire seulement si pas en chargement/sauvegarde
        if (iTypePer <> ACF_NONE) and (TomLoading = FALSE) and (TomUpdating = FALSE) then // $$$ JP 21/08/07 m_bLoadOrSave = FALSE) then
           SetField ('JEV_ALERTE', '-');

        // $$$ JP 11/04/07: ne pas activer les controles suivants lorsqu'on a pas le droit de modifier
        if (ModifAutorisee = TRUE) and (m_bAutoriseModif = TRUE) then
        begin
             SetControlEnabled ('JEV_ALERTE', iTypePer = ACF_NONE);
             SetControlVisible ('BFAIT',      iTypePer = ACF_NONE);
        end;

        // Texte d'aide
        UpdateHelpPeriodicite;
     finally
            m_bCalcPer := FALSE;
     end;
end;

procedure TOM_Evenement.UpdateHelpPeriodicite;
var
  iFrequence     :ActFrequence;
  strInterval    :hstring;
  iNbDays        :integer;
  strStart       :hstring;
  strEnd         :hstring;
  strOccurence   :hstring;
  strHelp        :hstring;
  i              :integer;
begin
     // Paramètres de la périodicité
     iFrequence   := ActFrequence (m_TypePeriode.ActivePageIndex);
     strStart     := m_eStartDate.Text; //GetControlText ('EPERDATEDEB');
     strOccurence := m_eEndOcc.Text; //GetControlText ('EPERENDOCC');
     strEnd       := GetControlText ('EPERDATEFIN');

     // Selon le type de fréquence, on construit un texte d'explication
     strHelp := '';
     case iFrequence of
          ACF_NONE:
          begin
               strHelp := 'le ' + strStart;
          end;

          ACF_DAILY:
          begin
               if m_rbPerJourTousLes.Checked = TRUE then
               begin
                    strInterval := GetControlText ('EPERJOURINTERVAL');
                    if strInterval = '1' then
                        strHelp := 'tous les jours'
                    else
                        strHelp := '1 jour sur ' + strInterval;
               end
               else
                   strHelp := 'tous les jours ouvrables (lundi à vendredi)';
          end;

          ACF_WEEKLY:
          begin
               strInterval := GetControlText ('EPERSEMAINEINTERVAL');
               iNbDays := 0;
               for i := 1 to 7 do
               begin
                    if m_chJours [i].Checked = TRUE then
                       strHelp := strHelp + ' ' + m_chJours [i].Caption + ',';
                    Inc (iNbDays);
               end;

               if iNbDays > 0 then
               begin
                    Delete (strHelp, Length (strHelp), 1);
                    if strInterval = '1' then
                        if iNbDays > 1 then
                            strHelp := 'toutes les semaines, les' + strHelp
                        else
                            strHelp := 'toutes les semaines, le' + strHelp
                    else
                        if iNbDays > 1 then
                            strHelp := 'toutes les ' + strInterval + ' semaines, les' + strHelp
                        else
                            strHelp := 'toutes les ' + strInterval + ' semaines, le' + strHelp;
               end
               else
                   strHelp := 'jamais';
          end;

          ACF_MONTHLY:
          begin
               if TRadioButton (GetControl ('RBPERMOIS')).Checked = TRUE then
               begin
                    strHelp := GetControlText ('EPERMOISJOUR');
                    if strHelp = '1' then
                       strHelp := strHelp + 'er';

                    strInterval := GetControlText ('EPERMOISINTERVAL');
                    if strInterval = '1' then
                        strHelp := 'tous les mois, le ' + strHelp
                    else
                        strHelp := 'tous les ' + strInterval + ' mois, le ' + strHelp;
                    strHelp := strHelp + ' du mois';
               end
               else
               begin
                    strInterval := GetControlText ('EPERMOISNINTERVAL');
                    if strInterval = '1' then
                        strHelp := 'tous les mois, '
                    else
                        strHelp := 'tous les ' + strInterval + ' mois, ';
                    strHelp := strHelp + 'le ' + THValComboBox (GetControl ('CBPERMOISNINSTANCE')).Text + ' ' + m_cbPerMoisNJour.Text + ' du mois';
               end;
          end;

          ACF_YEARLY:
          begin
               strHelp := 'toutes les années, ';
               if TRadioButton (GetControl ('RBPERANNEE')).Checked = TRUE then
               begin
                    strInterval := GetControlText ('EPERANNEEJOUR');
                    if strInterval = '1' then
                        strHelp := strHelp + 'le 1er ' + THValComboBox (GetControl ('CBPERANNEEMOIS')).Text
                    else
                        strHelp := strHelp + 'le ' + strInterval + ' ' + THValComboBox (GetControl ('CBPERANNEEMOIS')).Text;
               end
               else
               begin
                    strInterval := THValComboBox (GetControl ('CBPERANNEENMOIS')).Text;
                    strHelp := strHelp + 'le ' + THValComboBox (GetControl ('CBPERANNEENINSTANCE')).Text + ' ' + THValComboBox (GetControl ('CBPERANNEENJOUR')).Text;
                    if (strInterval <> '') and (char (strInterval [1]) in ['a','e','i','o','u']) then
                        strHelp := strHelp + ' d'''
                    else
                        strHelp := strHelp + ' de ';
                    strHelp := strHelp + strInterval;
               end;
          end;
     end;

     // Complément: plage de la fréquence
     if strHelp = 'jamais' then
        strHelp := 'Cette activité n''a jamais lieu: veuillez cocher au moins un jour de semaine.'
     else
     begin
          strHelp := 'Cette activité a lieu ' + strHelp + '.'#10;
          if iFrequence <> ACF_NONE then
          begin
               strHelp := strHelp + 'La série commence le ' + strStart;
               if m_rbNoEnd.Checked = TRUE then
                   strHelp := strHelp + ', sans limite dans le temps.'
               else if m_rbEndOcc.Checked = TRUE then
                   strHelp := strHelp + ' et est limitée à ' + strOccurence + ' répétitions.'
               else if m_rbEndDate.Checked = TRUE then
                   strHelp := strHelp + ' et se termine le ' + strEnd + '.';
         end
         else
             strHelp := strHelp + '(aucune répétition).';
     end;

     SetControlText ('LFREQUENCE', strHelp);
end;

procedure TOM_Evenement.OnChangeTypePeriode (Sender:TObject);
var
   dtDate    :TDateTime;
begin
     case ActFrequence (m_TypePeriode.ActivePageIndex) of
          ACF_NONE:
          begin
               SetControlVisible ('GBPLAGEPERIODICITE', FALSE);
          end;

          ACF_DAILY:
          begin
               // Calcul date fin, nb occurence, ... en fonction de la fréquence et de ce qui est déjà saisi
               SetControlVisible ('GBPLAGEPERIODICITE', TRUE);
          end;

          ACF_WEEKLY:
          begin
               SetControlVisible ('GBPLAGEPERIODICITE', TRUE);
          end;

          ACF_MONTHLY:
          begin
               m_bCalcPer := TRUE;

               // Valeurs par défaut si zone "vides", en fonction de la date début
               try
                  dtDate := GetField ('JEV_DATE');
                  if GetControlText ('EPERMOISJOUR') = '0' then
                     SetControlText ('EPERMOISJOUR', IntToStr (DayOf (dtDate)));
                  if m_cbPerMoisNInstance.ItemIndex = -1 then
                  begin
                       m_cbPerMoisNInstance.ItemIndex := AgendaGetInstanceOfDate (dtDate) - 1;
                       if DayOfWeek (dtDate) = 1 then
                           m_cbPerMoisNJour.ItemIndex  := 9
                       else
                           m_cbPerMoisNJour.ItemIndex  := DayOfWeek (dtDate) + 1;
                  end;
                  SetControlVisible ('GBPLAGEPERIODICITE', TRUE);
               finally
                      m_bCalcPer := FALSE;
               end;
          end;

          ACF_YEARLY:
          begin
               m_bCalcPer := TRUE;

               // Valeurs par défaut si zone "vides", en fonction de la date début
               try
                  // Valeurs par défaut si zone "vides", en fonction de la date début
                  dtDate := GetField ('JEV_DATE');
                  if GetControlText ('EPERANNEEJOUR') = '0' then
                  begin
                       SetControlText ('EPERANNEEJOUR', IntToStr (DayOf (dtDate)));
                       m_cbPerAnneeMois.ItemIndex := MonthOf (dtDate) - 1;
                  end;
                  if m_cbPerAnneeNInstance.ItemIndex = -1 then
                  begin
                       m_cbPerAnneeNInstance.ItemIndex := AgendaGetInstanceOfDate (dtDate) - 1;
                       if DayOfWeek (dtDate) = 1 then
                           m_cbPerAnneeNJour.ItemIndex  := 9
                       else
                           m_cbPerAnneeNJour.ItemIndex  := DayOfWeek (dtDate) + 1;
                       m_cbPerAnneeNMois.ItemIndex := MonthOf (dtDate) - 1;
                  end;
                  SetControlVisible ('GBPLAGEPERIODICITE', TRUE);
               finally
                      m_bCalcPer := FALSE;
               end;
          end;
     end;

     UpdatePlagePeriodicite;
end;

// $$$ JP 26/10/06 - sélection du type périodicité quotidien
procedure TOM_Evenement.SelectTypeJour (bOuvrable:boolean);
begin
     SetControlEnabled ('EPERJOURINTERVAL', not bOuvrable);

     UpdatePlagePeriodicite;
end;

// $$$ JP 25/09/06 - sélection du type périodicité mensuelle
procedure TOM_Evenement.SelectTypeMois (bMoisN:boolean);
begin
     SetControlEnabled ('EPERMOISJOUR',       not bMoisN);
     SetControlEnabled ('LPERMOISTOUSLES',    not bMoisN);
     SetControlEnabled ('EPERMOISINTERVAL',   not bMoisN);
     SetControlEnabled ('LPERMOISMOIS',       not bMoisN);

     SetControlEnabled ('CBPERMOISNINSTANCE', bMoisN);
     SetControlEnabled ('CBPERMOISNJOUR',     bMoisN);
     SetControlEnabled ('LPERMOISNTOUSLES',   bMoisN);
     SetControlEnabled ('EPERMOISNINTERVAL',  bMoisN);
     SetControlEnabled ('LPERMOISNMOIS',      bMoisN);

     UpdatePlagePeriodicite;
end;

// $$$ JP 25/09/06 - sélection du type périodicité annuelle
procedure TOM_Evenement.SelectTypeAnnee (bAnneeN:boolean);
begin
     SetControlEnabled ('EPERANNEEJOUR',       not bAnneeN);
     SetControlEnabled ('CBPERANNEEMOIS',      not bAnneeN);

     SetControlEnabled ('CBPERANNEENINSTANCE', bAnneeN);
     SetControlEnabled ('CBPERANNEENJOUR',     bAnneeN);
     SetControlEnabled ('LPERANNEENDE',        bAnneeN);
     SetControlEnabled ('CBPERANNEENMOIS',     bAnneeN);

     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnClickJourTousLes (Sender:TObject);
begin
     SelectTypeJour (FALSE);
end;

procedure TOM_Evenement.OnClickJourOuvrables (Sender:TObject);
begin
     SelectTypeJour (TRUE);
end;

procedure TOM_Evenement.OnClickMois (Sender:TObject);
begin
     SelectTypeMois (FALSE);
end;

procedure TOM_Evenement.OnClickMoisN (Sender:TObject);
begin
     SelectTypeMois (TRUE);
end;

procedure TOM_Evenement.OnClickAnnee (Sender:TObject);
begin
     SelectTypeAnnee (FALSE);
end;

procedure TOM_Evenement.OnClickAnneeN (Sender:TObject);
begin
     SelectTypeAnnee (TRUE);
end;

procedure TOM_Evenement.OnClickNoEnd (Sender:TObject);
begin
     SetControlEnabled ('EPERENDOCC',     FALSE);
     SetControlEnabled ('LPEROCCURENCE',  FALSE);
     SetControlEnabled ('EPERDATEFIN',    FALSE);
     SetControlText    ('EPERENDOCC',     '10');

     // $$$ JP 22/01/07
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnClickEndOcc (Sender:TObject);
begin
     SetControlEnabled ('EPERENDOCC',     TRUE);
     SetControlEnabled ('LPEROCCURENCE',  TRUE);
     SetControlEnabled ('EPERDATEFIN',    FALSE);
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnClickEndDate (Sender:TObject);
begin
     SetControlEnabled ('EPERENDOCC',     FALSE);
     SetControlEnabled ('LPEROCCURENCE',  FALSE);
     SetControlEnabled ('EPERDATEFIN',    TRUE);
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeEndOcc (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeEndDate (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeJourInterval (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeSemaineInterval (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnClickUnJourSemaine (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeMoisJour (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeMoisInterval (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeMoisNInstance (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeMoisNJour (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeMoisNInterval (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeAnneeJour (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeAnneeMois (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeAnneeNInstance (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeAnneeNJour (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;

procedure TOM_Evenement.OnChangeAnneeNMois (Sender:TObject);
begin
     UpdatePlagePeriodicite;
end;
{$ENDIF}


procedure TOM_Evenement.OnChangeCodeEvt(Sender: TObject);
var
   TOBTypeEvt    :TOB;
   strLieu       :string;
   strMissionAdm :string;
begin
     // $$$ JP 11/04/07: ne pas modifier le JEV_LIEU si déjà ok. Pas d'autre choix hélas, on passe par là à des moments non voulu!
{$IFDEF EAGLCLIENT}
     if GetField ('JEV_LIEU') <> '' then
        exit;
{$ENDIF}

     TOBTypeEvt := TOB.Create ('lieu evt', nil, -1);
     try
        TOBTypeEvt.LoadDetailFromSQL ('SELECT JTE_LIEU,JTE_AFFAIRE FROM JUTYPEEVT WHERE JTE_CODEEVT="' + GetControlText ('JEV_CODEEVT') + '" AND JTE_FAMEVT="ACT"');
        if TOBTypeEvt.Detail.Count > 0 then
        begin
             strLieu := TOBTypeEvt.Detail [0].GetValue ('JTE_LIEU');
             SetField ('JEV_LIEU', strLieu);
             if GetField ('JEV_ABSENCE') = 'X' then
             begin
                  strMissionAdm := TOBTypeEvt.Detail [0].GetValue ('JTE_AFFAIRE');
                  if strMissionAdm <> '' then
                  begin
                       SetField ('JEV_NODOSSIER', NoDossierBaseCommune);
                       SetField ('JEV_AFFAIRE', strMissionAdm);
                  end
                  else
                  begin
                       SetField ('JEV_NODOSSIER', '');
                       SetField ('JEV_AFFAIRE', '');
                  end;
             end;
        end;
     finally
            TOBTypeEvt.Free;
     end;
end;


procedure TOM_Evenement.OnDeleteRecord;
{$IFDEF BUREAU}
var
   strOcc  :string;
{$ENDIF}
begin
     strEvtMaj := GetField ('JEV_GUIDEVT');

{$IFDEF BUREAU}
     if Ecran.Name = 'YYAGENDA_FIC' then
     begin
          // $$$ JP 21/09/07: inutile tant que pas gestion des exception
          exit;

          strOcc := GetField ('JEV_OCCURENCEEVT');
          if strOcc <> 'SIN' then
          begin
               if strOcc = 'REC' then
               begin
                    if PgiAsk ('Confirmez-vous la suppression de la série?') <> mrYes then
                    begin
                         LastError    := 1;
                         LastErrorMsg := 'Suppression annulée';
                         exit;
                    end;
               end
               else
               begin
                    if PgiAsk ('Confirmez-vous la suppression de cette occurence ?') <> mrYes then
                    begin
                         LastError    := 1;
                         LastErrorMsg := 'Suppression annulée';
                         exit;
                    end;
               end;
          end;
     end;
{$ENDIF}

     inherited;
end;

// $$$ JP 21/08/07
procedure TOM_Evenement.OnAfterDeleteRecord;
begin
{$IFDEF BUREAU}
     if (Ecran.Name = 'YYAGENDA_FIC') and (strEvtMaj <> '') then
     begin
          // $$$ JP 26/04/07: on met à jour le planning unifié
          deleteYPL ('JEV', strEvtMaj);

          // $$$ JP 21/08/07: ménage dans les répétition et/ou exception
          if m_bException = FALSE then
          begin
               BeginTrans;

               try
                  ExecuteSQL ('DELETE JUEVTRECURRENCE WHERE JEC_GUIDEVT="' + strEvtMaj + '"');
                  AgendaDeleteExceptions (strEvtMaj);

                  CommitTrans;
               except
                  RollBack;
               end;
          end
          else
          begin
               // Pour une exception: il faut màj l'enreg d'exception (JEX) pour indiquer une exception de suppression
               ExecuteSQL ('UPDATE JURECEXCEPTION SET JEX_EXCEPTTYPE=0,JEX_GUIDEVTEXCEPT="" WHERE JEX_GUIDEVTEXCEPT="' + strEvtMaj + '"');
          end;

          // Indiquer la suppression par une * en début
          strEvtMaj := '*' + strEvtMaj;
     end;
{$ENDIF}

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/07/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Evenement.OnClick_BAffiche(Sender: TObject);
begin
   LanceWord(GetControlText('DOCNOM'), '');
//   OnDocChange;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_Evenement.OnElipsisClick_DOCNOM(Sender : TObject);
var
	sNewDocPath_l, sMessage_l : string;
begin
	sNewDocPath_l := ChoixMaquette('*.doc', 'Documents Word (*.doc)|*.doc|' +
                                           'Documents RTF (*.rtf)|*.rtf|' +
                                           'Documents RPT (*.rpt)|*.rpt|');

   if sNewDocPath_l = '' then exit;
   if GetControlText('DOCNOM') = sNewDocPath_l then exit;

   ModeEdition(DS);
   sMessage_l := 'Le document a été modifié.'#10#13 +
              'Voulez-vous enregistrer les modifications?';
   TFFiche(Ecran).HM.Mess[0] := '0;?caption?;' + sMessage_l + ';Q;YNC;Y;C;';

   SetField('JEV_DOCNOM', sNewDocPath_l);
	SetControlText('DOCNOM', sNewDocPath_l);

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure AGLMajDateAlerte (parms:array of variant; nb:integer);
var
   F   :TForm;
   OM  :TOM ;
begin
     F:=TForm(Longint(Parms[0])) ;
     if (F is TFFiche) then
         OM := TFFiche(F).OM
     else
         exit;
     if (OM is TOM_Evenement) then
         TOM_Evenement(OM).MajDateAlerte
     else
         exit;
end;


Initialization
   registerclasses([TOM_Evenement]) ;
   RegisterAGLProc ('MajDateAlerte', TRUE, 0, AGLMajDateAlerte);

end.

